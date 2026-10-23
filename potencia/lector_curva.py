"""
Lectura robusta de curvas de carga (cuartohorarias u horarias).

Formatos admitidos (autodetectados):
  - CSV / TXT con cualquier separador (; , tabulador |) y codificación habitual.
  - Excel (.xlsx, .xlsm, .xls), eligiendo la hoja con datos de curva.
  - Con o sin cabecera (p. ej. ficheros P5D/F5D de REE sin cabecera).
  - Fecha y hora en una sola columna, o fecha + hora (+ cuarto) en columnas.
  - Formato "ancho": una fila por día y 24/25 (o 96/100) columnas de valores.
  - Valores en kW (potencia media) o en kWh / Wh (energía del intervalo).

Convenio horario:
  El estándar del Sistema de Medidas (REE, "Ficheros para el intercambio de
  información de medida") es que la etiqueta de tiempo de cada registro
  corresponde al FINAL del periodo de integración: el primer cuarto del día
  31/01 es "31/01 00:15" y el último "01/02 00:00" (en curvas horarias,
  hora 1 = de 0 h a 1 h y hora 24 = de 23 h a 24 h). Si no hay pistas claras
  en el fichero, se asume este convenio.

Cambios de hora (petición de GE&PE):
  - Día de 23 h (marzo): la hora que falta se rellena con el promedio de la
    hora anterior y la hora posterior.
  - Día de 25 h (octubre): la hora repetida se sustituye por el promedio de
    sus dos registros.
  Así todos los días quedan con 96 cuartos de hora.
"""

import csv
import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

from .calendario import dias_cambio_hora, hora_cambio_horario, zona_horaria

UNIDADES = ["Automática", "kW", "kWh", "Wh", "MW", "MWh", "W"]
# factor para pasar el valor a kW: potencia -> constante; energía del intervalo -> depende del paso (min)
_POTENCIA = {"kW": 1.0, "MW": 1000.0, "W": 0.001}
_ENERGIA = {"kWh": 1.0, "MWh": 1000.0, "Wh": 0.001}
# símbolo de unidad aislado (no pegado a otras letras): "(kW)", "_kWh", "[MWh]", "kW·h", "kW-h"
_RE_UNIDAD = re.compile(r"(?<![a-z])(mw\s*[-.]?\s*h|kw\s*[-.]?\s*h|wh|mw|kw|w)(?![a-z0-9])")
_TOKENS = {"mwh": "MWh", "kwh": "kWh", "wh": "Wh", "mw": "MW", "kw": "kW", "w": "W"}
CONVENIOS = ["Automático", "Fin de periodo", "Inicio de periodo"]

_RE_FECHA_DMY = re.compile(r"^\s*\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}")
_RE_FECHA_YMD = re.compile(r"^\s*\d{4}[/.-]\d{1,2}[/.-]\d{1,2}")
_RE_CUPS = re.compile(r"^ES\d{16}[A-Z]{2}(\d[A-Z])?$", re.I)
_RE_HHMM = re.compile(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$")
_RE_24H = re.compile(r"\s24:00(?::00)?\s*$")
_RE_TZ = re.compile(r"(?:Z|[+-]\d{2}:?\d{2})\s*$")


@dataclass
class CurvaCargada:
    """Curva normalizada a cuartos de hora en hora oficial local."""
    datos: pd.DataFrame            # columnas: inicio, kw, estimado
    archivo: str
    hoja: str | None
    resolucion_min: int            # resolución original (15 o 60)
    unidad: str                    # unidad original de los valores
    convenio: str                  # "Fin de periodo" / "Inicio de periodo"
    cups: str | None
    columnas: dict = field(default_factory=dict)
    avisos: list = field(default_factory=list)
    n_registros: int = 0
    origen_unidad: str = ""
    huecos_rellenados: int = 0
    dias_cambio_hora: list = field(default_factory=list)
    ficheros: list = field(default_factory=list)       # resumen de cada fichero leído
    sin_datos: list = field(default_factory=list)      # periodos largos sin datos (no se rellenan)

    @property
    def fecha_inicio(self):
        return self.datos["inicio"].iloc[0]

    @property
    def fecha_fin(self):
        return self.datos["inicio"].iloc[-1] + pd.Timedelta(minutes=15)

    @property
    def es_horaria(self):
        return self.resolucion_min > 15

    @property
    def descripcion_unidad(self):
        """Unidad, de dónde se ha sacado y cómo se pasa a potencia en kW."""
        if self.unidad not in _POTENCIA and self.unidad not in _ENERGIA:      # varios ficheros con unidades distintas
            return f"{self.unidad} – {self.origen_unidad}"
        if self.unidad in _POTENCIA:
            como = "valores de potencia: se usan directamente" + (
                "" if self.unidad == "kW" else f" (pasados de {self.unidad} a kW)")
        else:
            como = (f"valores de energía de cada intervalo de {self.resolucion_min} min: se convierten a potencia "
                    f"media en kW (× {60 / self.resolucion_min:g}" +
                    ("" if self.unidad == "kWh" else f", pasando de {self.unidad} a kWh") + ")")
        return f"{self.unidad} – {self.origen_unidad}; {como}"


# --------------------------------------------------------------------------
# utilidades de conversión
# --------------------------------------------------------------------------

def _norm(texto):
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().lower()


def _vacio(v):
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    return isinstance(v, str) and v.strip() == ""


def _parece_fecha(v):
    if isinstance(v, (datetime, date, pd.Timestamp)) and not pd.isna(v):
        return True
    if isinstance(v, str):
        return bool(_RE_FECHA_DMY.match(v) or _RE_FECHA_YMD.match(v))
    return False


def a_numero(serie):
    """Convierte textos con formato español o inglés a float."""
    def conv(v):
        if _vacio(v):
            return np.nan
        if isinstance(v, (int, float, np.number)) and not isinstance(v, bool):
            return float(v)
        s = str(v).strip().replace("\xa0", "").replace(" ", "")
        if s in ("-", "--"):
            return np.nan
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return np.nan
    return pd.Series([conv(v) for v in serie], index=serie.index, dtype=float)


def _a_fecha_hora(serie, zona):
    """
    Convierte una serie a datetime (hora local, sin zona).
    Devuelve (serie_datetime, hubo_24h) donde hubo_24h indica etiquetas "24:00".
    """
    tz = zona_horaria(zona)
    res = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    utc = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns, UTC]")
    hubo_24 = False

    es_dt = serie.map(lambda v: isinstance(v, (datetime, pd.Timestamp)) and not pd.isna(v))
    if es_dt.any():
        vals = pd.to_datetime(serie[es_dt])
        if getattr(vals.dt, "tz", None) is not None:
            utc[es_dt] = vals.dt.tz_convert("UTC")
            vals = vals.dt.tz_convert(tz).dt.tz_localize(None)
        res[es_dt] = vals
    es_d = serie.map(lambda v: isinstance(v, date) and not isinstance(v, datetime))
    if es_d.any():
        res[es_d] = pd.to_datetime(serie[es_d])

    # números de serie de Excel (fecha sin formato)
    es_num = serie.map(lambda v: isinstance(v, (int, float, np.number)) and not isinstance(v, bool)
                       and not pd.isna(v) and 20000 < float(v) < 80000)
    if es_num.any():
        res[es_num] = pd.to_datetime("1899-12-30") + pd.to_timedelta(serie[es_num].astype(float), unit="D")

    es_txt = serie.map(lambda v: isinstance(v, str) and v.strip() != "")
    if es_txt.any():
        txt = serie[es_txt].str.strip()
        mas_un_dia = txt.str.contains(_RE_24H)
        if mas_un_dia.any():
            hubo_24 = True
            txt = txt.where(~mas_un_dia, txt.str.replace(_RE_24H, " 00:00", regex=True))
        con_tz = txt.str.contains(_RE_TZ) & txt.str.contains(r"\d{1,2}:\d{2}")
        if con_tz.any():
            v = pd.to_datetime(txt[con_tz], utc=True, errors="coerce", format="mixed")
            utc[txt[con_tz].index] = v
            res[txt[con_tz].index] = v.dt.tz_convert(tz).dt.tz_localize(None)
        resto = txt[~con_tz]
        if len(resto):
            ymd = resto.str.match(_RE_FECHA_YMD)
            for mascara, dayfirst in ((ymd, False), (~ymd, True)):
                sub = resto[mascara]
                if len(sub):
                    v = pd.to_datetime(sub, dayfirst=dayfirst, errors="coerce", format="mixed")
                    res[sub.index] = v
        if mas_un_dia.any():
            idx = mas_un_dia[mas_un_dia].index
            res[idx] = res[idx] + pd.Timedelta(days=1)
    if utc.notna().sum() < 0.99 * res.notna().sum():
        utc = None
    return res, hubo_24, utc


def _a_minutos(serie):
    """
    Convierte una columna de hora a minutos desde las 0 h.
    Devuelve (minutos, tipo) con tipo "entera" (1, 2, ..., 24) o "reloj" (HH:MM).
    """
    enteros = []
    minutos = []
    tipo_reloj = False
    for v in serie:
        if _vacio(v):
            minutos.append(np.nan); enteros.append(False); continue
        if isinstance(v, time):
            minutos.append(v.hour * 60 + v.minute); tipo_reloj = True; enteros.append(False); continue
        if isinstance(v, timedelta):
            minutos.append(round(v.total_seconds() / 60)); tipo_reloj = True; enteros.append(False); continue
        if isinstance(v, (datetime, pd.Timestamp)):
            base = datetime(1899, 12, 31)
            dias = (v.date() - base.date()).days if v.year < 1901 else 0
            minutos.append(dias * 1440 + v.hour * 60 + v.minute); tipo_reloj = True
            enteros.append(False); continue
        if isinstance(v, (int, float, np.number)) and not isinstance(v, bool):
            f = float(v)
            if 0 < f < 1 and not f.is_integer():   # fracción de día de Excel
                minutos.append(round(f * 1440)); tipo_reloj = True; enteros.append(False)
            else:
                minutos.append(f * 60); enteros.append(True)
            continue
        s = str(v).strip()
        m = _RE_HHMM.match(s)
        if m:
            minutos.append(int(m.group(1)) * 60 + int(m.group(2))); tipo_reloj = True
            enteros.append(False); continue
        m = re.match(r"^\s*(\d{1,2})\s*(h|H)?\s*$", s)
        if m:
            minutos.append(int(m.group(1)) * 60); enteros.append(True); continue
        m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*$", s)   # "0-1", "23-24"
        if m:
            minutos.append(int(m.group(2)) * 60); enteros.append(True); continue
        minutos.append(np.nan); enteros.append(False)
    tipo = "reloj" if tipo_reloj else "entera"
    return pd.Series(minutos, index=serie.index, dtype=float), tipo


# --------------------------------------------------------------------------
# lectura de la tabla cruda
# --------------------------------------------------------------------------

def _leer_texto(ruta):
    crudo = Path(ruta).read_bytes()
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = crudo.decode(cod)
            break
        except UnicodeDecodeError:
            continue
    lineas = [l for l in texto.splitlines() if l.strip()][:200]
    mejor, mejor_puntos = ";", -1
    for sep in (";", "\t", "|", ","):
        cuentas = [l.count(sep) for l in lineas]
        if not cuentas or max(cuentas) == 0:
            continue
        moda = max(set(cuentas), key=cuentas.count)
        puntos = cuentas.count(moda) * (1 if moda > 0 else 0)
        if sep == ",":           # la coma suele ser separador decimal
            puntos *= 0.8
        if puntos > mejor_puntos:
            mejor, mejor_puntos = sep, puntos
    filas = list(csv.reader(io.StringIO(texto), delimiter=mejor))
    ancho = max((len(f) for f in filas), default=0)
    filas = [f + [""] * (ancho - len(f)) for f in filas if any(c.strip() for c in f)]
    return pd.DataFrame(filas, dtype=object)


def _leer_tablas(ruta):
    ext = Path(ruta).suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        hojas = pd.read_excel(ruta, sheet_name=None, header=None,
                              engine="xlrd" if ext == ".xls" else None)
        return {nombre: df.dropna(how="all").dropna(axis=1, how="all") for nombre, df in hojas.items()}
    return {None: _leer_texto(ruta)}


def _puntuar_hoja(nombre, df):
    if df.empty:
        return 0
    muestra = df.head(3000)
    filas_fecha = muestra.apply(lambda f: any(_parece_fecha(v) for v in f), axis=1).sum()
    puntos = filas_fecha * (len(df) / max(len(muestra), 1))
    n = _norm(nombre or "")
    if any(k in n for k in ("curva", "potencia", "consumo", "carga", "datos")):
        puntos *= 1.5
    if re.fullmatch(r"curva[ _]?potencia", n):
        puntos *= 3          # hoja de la curva en los Excel de estudio de GE&PE
    if "campos" in n:
        puntos *= 0.2        # hojas auxiliares «Curva 2 campos» / «Curva 4 campos» de esos Excel
    return puntos


def _cups_en_texto(texto):
    """CUPS que aparecen en un texto (nombre de fichero, notas encima de la tabla), sin repetir."""
    vistos = []
    for m in re.finditer(r"ES\d{16}[A-Z]{2}(?:\d[A-Z])?", str(texto).upper()):
        if m.group(0) not in vistos:
            vistos.append(m.group(0))
    return vistos


def _a_hora_reloj(reg, paso, ordinal, zona):
    """
    Registros con posición cronológica dentro del día (horas 1-25, columnas por hora…) -> hora de reloj
    en los días de cambio de hora. Si ya están en hora de reloj, no cambia nada.
    """
    if not ordinal:
        return reg
    reg = reg.copy()
    h_cambio = hora_cambio_horario(zona) * 60
    anios = sorted({d.year for d in pd.DatetimeIndex(reg["dia"])})
    marzo = {pd.Timestamp(dias_cambio_hora(a)[0]) for a in anios}
    octubre = {pd.Timestamp(dias_cambio_hora(a)[1]) for a in anios}
    max_min = reg.groupby("dia")["minuto"].transform("max")
    es_mar = reg["dia"].isin(marzo) & (max_min <= 1440 - 60 - paso)
    es_oct = reg["dia"].isin(octubre) & (max_min >= 1440)
    reg.loc[es_mar & (reg["minuto"] >= h_cambio), "minuto"] += 60
    reg.loc[es_oct & (reg["minuto"] >= h_cambio + 60), "minuto"] -= 60
    return reg[reg["minuto"] < 1440]


def _a_paso(tabla, paso, paso_final):
    """Registros horarios -> cuartos de hora (cada cuarto con la potencia media de su hora)."""
    if paso == paso_final:
        return tabla
    rep = paso // paso_final
    t = tabla.loc[tabla.index.repeat(rep)].copy()
    t["minuto"] = t["minuto"].to_numpy() + np.tile(np.arange(rep) * paso_final, len(tabla))
    return t


def _componer(lista, zona):
    """
    Une los registros de uno o varios ficheros (p. ej. uno por mes) en una única curva normalizada.
    lista: resultados de LectorCurva.registros (todos del mismo suministro).
    """
    avisos = []
    varios = len(lista) > 1
    for r in lista:
        avisos += [f"{Path(r['archivo']).name}: {a}" if varios else a for a in r["avisos"]]
    paso = min(r["paso"] for r in lista)
    if len({r["paso"] for r in lista}) > 1:
        avisos.append("Hay ficheros horarios y cuartohorarios: en los horarios cada cuarto de hora toma la potencia "
                      "media de su hora.")
    tablas = []
    for k, r in enumerate(lista):
        t = _a_paso(r["tabla"][["dia", "minuto", "kw"]], r["paso"], paso)
        tablas.append(t.assign(fichero=k))
    reg = pd.concat(tablas, ignore_index=True)

    if varios:
        # cuartos de hora que vienen en más de un fichero (ficheros que se solapan)
        reg["inicio"] = pd.to_datetime(reg["dia"]) + pd.to_timedelta(reg["minuto"], unit="m")
        en_varios = reg.groupby("inicio")["fichero"].nunique()
        n_solape = int((en_varios > 1).sum())
        if n_solape:
            avisos.append(f"{n_solape} cuartos de hora aparecen en más de un fichero (los ficheros se solapan); "
                          "se ha usado el promedio.")
            # para que no cuenten además como «duplicados»: un valor por fichero y cuarto
            reg = reg.groupby(["inicio", "fichero"], as_index=False).agg(
                dia=("dia", "first"), minuto=("minuto", "first"), kw=("kw", "mean"))
        reg = reg.drop(columns=["inicio"], errors="ignore")

    serie, dias_ch, huecos, duplicados, sin_datos = LectorCurva._normalizar(reg.drop(columns="fichero"), paso,
                                                                            False, zona)
    if duplicados and not varios:
        avisos.append(f"{duplicados} registros duplicados (distintos del cambio de hora) se han promediado.")
    if huecos:
        pct = 100.0 * huecos / len(serie)
        avisos.append(f"Se han rellenado {huecos} intervalos sin dato ({pct:.2f} %) con el promedio "
                      f"de la hora anterior y posterior.")
        if pct > 5:
            avisos.append("¡Atención! Más del 5 % de la curva es estimada: el resultado puede no ser fiable.")
    for a, b in sin_datos:
        avisos.append(f"¡Atención! No hay datos del {a:%d/%m/%Y %H:%M} al {b:%d/%m/%Y %H:%M}"
                      + (" (¿falta algún fichero?)" if varios else "") + ": ese periodo no se rellena ni se "
                      "factura, así que el estudio no corresponde a un año completo.")

    if paso > 15:
        serie = LectorCurva._a_cuartos(serie, paso)
        avisos.append("Curva HORARIA: se ha estimado cada cuarto de hora con la potencia media de su hora "
                      "(criterio del art. 9 de la Circular 3/2020 cuando no hay registro cuartohorario). "
                      "Los picos de 15 minutos quedan suavizados y los excesos reales pueden ser mayores.")
    serie = LectorCurva._recortar_12_meses(serie, avisos)

    unidades = list(dict.fromkeys(r["unidad"] for r in lista))
    convenios = list(dict.fromkeys(r["convenio"] for r in lista))
    if len(unidades) == 1:
        unidad, origen = unidades[0], lista[0]["origen_unidad"] + (" (en todos los ficheros)" if varios else "")
    else:
        unidad = "varias (" + ", ".join(unidades) + ")"
        origen = "; ".join(f"{Path(r['archivo']).name}: {r['unidad']}, {r['origen_unidad']}" for r in lista)
    ficheros = []
    if varios:
        for r in lista:
            dias = pd.DatetimeIndex(r["tabla"]["dia"])
            ficheros.append(f"{Path(r['archivo']).name}: {r.get('cups') or 'sin CUPS'} · "
                            f"{dias.min():%d/%m/%Y} – {dias.max():%d/%m/%Y} · {r['unidad']}")
    cups = next((r["cups"] for r in lista if r.get("cups")), None)
    return CurvaCargada(
        datos=serie.reset_index(drop=True),
        archivo="; ".join(Path(r["archivo"]).name if varios else r["archivo"] for r in lista),
        hoja=lista[0]["hoja"] if not varios else None,
        resolucion_min=paso, unidad=unidad, origen_unidad=origen,
        convenio=convenios[0] if len(convenios) == 1 else "; ".join(convenios), cups=cups,
        columnas=lista[0]["columnas"], avisos=avisos, n_registros=sum(r["n_registros"] for r in lista),
        huecos_rellenados=int(huecos), dias_cambio_hora=dias_ch, ficheros=ficheros, sin_datos=sin_datos)


# --------------------------------------------------------------------------
# lector
# --------------------------------------------------------------------------

class LectorCurva:
    """Lee el fichero una sola vez; `procesar` puede repetirse con otros ajustes."""

    def __init__(self, ruta):
        self.ruta = str(ruta)
        self.tablas = _leer_tablas(ruta)
        if not self.tablas:
            raise ValueError("El fichero está vacío.")
        puntos = {n: _puntuar_hoja(n, df) for n, df in self.tablas.items()}
        self.hojas = sorted(self.tablas, key=lambda n: -puntos[n])
        self.hoja_defecto = self.hojas[0]
        if puntos[self.hoja_defecto] == 0:
            raise ValueError("No se han encontrado fechas en el fichero. ¿Es una curva de carga?")
        self.cups_disponibles = []

    # ------------------------------------------------------------------
    def procesar(self, zona="Península", unidad="Automática", convenio="Automático",
                 hoja=None, cups=None):
        """Curva del fichero (del CUPS indicado si el fichero tiene varios), normalizada a cuartos de hora."""
        return _componer([self.registros(zona, unidad, convenio, hoja, cups)], zona)

    def cups_en_fichero(self, hoja=None):
        """CUPS del fichero: los de su columna de CUPS o, si no tiene, el del nombre o el texto encima de la tabla."""
        hoja = self.hoja_defecto if hoja is None else hoja
        df = self.tablas[hoja].reset_index(drop=True)
        cab, datos, preambulo = self._separar_cabecera(df)
        nombres = [(_norm(c) if not _vacio(c) else "") for c in cab] if cab is not None else [""] * df.shape[1]
        datos = datos.copy()
        datos.columns = range(datos.shape[1])
        col = self._clasificar_columnas(datos, nombres)
        if col.get("cups") is not None:
            serie = datos[col["cups"]].astype(str).str.strip().str.upper()
            return [c for c in serie.value_counts().index if _RE_CUPS.match(c)], True
        return _cups_en_texto(Path(self.ruta).stem + " " + preambulo), False

    def registros(self, zona="Península", unidad="Automática", convenio="Automático", hoja=None, cups=None):
        """
        Registros del fichero en hora de reloj (sin normalizar todavía): dict con la tabla (dia, minuto, kw),
        el paso, la unidad, el convenio, los avisos y el CUPS. Varios ficheros se unen con _componer.
        """
        hoja = self.hoja_defecto if hoja is None else hoja
        df = self.tablas[hoja].reset_index(drop=True)
        avisos = []

        cab, datos, preambulo = self._separar_cabecera(df)
        nombres = [(_norm(c) if not _vacio(c) else "") for c in cab] if cab is not None else [""] * df.shape[1]
        datos = datos.copy()
        datos.columns = range(datos.shape[1])

        col = self._clasificar_columnas(datos, nombres)

        # filtro de CUPS
        self.cups_disponibles = []
        cups_sel = None
        if col.get("cups") is not None:
            serie_cups = datos[col["cups"]].astype(str).str.strip().str.upper()
            self.cups_disponibles = [c for c in serie_cups.value_counts().index if _RE_CUPS.match(c)]
            if self.cups_disponibles:
                cups_sel = cups if cups in self.cups_disponibles else self.cups_disponibles[0]
                if len(self.cups_disponibles) > 1:
                    avisos.append(f"El fichero contiene {len(self.cups_disponibles)} CUPS; "
                                  f"se usa {cups_sel}.")
                datos = datos[serie_cups == cups_sel]
        else:
            # sin columna de CUPS: el del nombre del fichero o el del texto encima de la tabla
            en_texto = _cups_en_texto(Path(self.ruta).stem + " " + preambulo)
            if len(en_texto) == 1:
                cups_sel = en_texto[0]
                self.cups_disponibles = en_texto

        registros, paso, conv_txt, ordinal = self._construir_registros(datos, col, zona, convenio, avisos)
        if len(registros) == 0:
            raise ValueError("No se han podido interpretar registros de fecha/hora y valor.")

        # unidad
        unidad_det, origen = unidad, "elegida a mano"
        if unidad == "Automática":
            unidad_det, origen = self._unidad_automatica(col, datos, preambulo)
            if unidad_det is None:
                unidad_det, origen = "kWh", "supuesta (no indicada en el fichero)"
                avisos.append("No se ha podido identificar la unidad (ni en el título de las columnas, ni en una "
                              "columna de unidades, ni encima de la tabla); se asume kWh (energía del intervalo, "
                              "estándar de las distribuidoras). Revísalo.")
        if unidad_det in _POTENCIA:
            factor = _POTENCIA[unidad_det]
        else:
            factor = _ENERGIA[unidad_det] * 60.0 / paso
        registros["kw"] = registros["valor"] * factor

        negativos = (registros["kw"] < 0).sum()
        if negativos:
            avisos.append(f"{negativos} registros con valor negativo se han puesto a 0.")
            registros.loc[registros["kw"] < 0, "kw"] = 0.0

        return {"tabla": _a_hora_reloj(registros, paso, ordinal, zona), "paso": paso, "convenio": conv_txt,
                "unidad": unidad_det, "origen_unidad": origen, "avisos": avisos, "cups": cups_sel,
                "archivo": self.ruta, "hoja": hoja, "n_registros": len(registros),
                "columnas": {k: v for k, v in col.items() if k.startswith("nombre_") or k == "formato"}}

    # ------------------------------------------------------------------
    @staticmethod
    def _separar_cabecera(df):
        n = len(df)

        def fila_de_datos(i):
            fila = df.iloc[i]
            tiene_fecha = any(_parece_fecha(v) for v in fila)
            numeros = a_numero(fila).notna().sum()
            return tiene_fecha and numeros >= 1

        inicio = None
        for i in range(min(n, 200)):
            if all(fila_de_datos(j) for j in range(i, min(i + 5, n))):
                inicio = i
                break
        if inicio is None:
            raise ValueError("No se ha encontrado el bloque de datos (filas con fecha y valor).")
        def textos(fila):
            return [v for v in fila if isinstance(v, str) and v.strip() and not _parece_fecha(v)]

        # fila de títulos: la más cercana a los datos con varios textos (o solo texto); las filas que haya
        # entre ella y los datos (p. ej. una fila con las unidades) se unen a los títulos por columna
        cab, j_cab = None, inicio
        for j in range(inicio - 1, max(-1, inicio - 4), -1):
            fila = df.iloc[j]
            tx = textos(fila)
            if len(tx) >= 2 or (tx and sum(not _vacio(v) for v in fila) == len(tx) and j == inicio - 1
                                and j - 1 >= 0 and len(textos(df.iloc[j - 1])) < 2):
                j_cab = j
                break

        def es_fila_unidades(fila):
            tx = [_norm(v) for v in textos(fila)]
            return bool(tx) and all(_RE_UNIDAD.fullmatch(v.strip("()[] ")) or re.fullmatch(r"[dmayhs/:.\- ]+", v)
                                    for v in tx)

        # si la fila elegida solo tiene unidades o formatos (kW, dd/mm/aaaa, hh:mm), los títulos están encima
        if 0 < j_cab < inicio and es_fila_unidades(df.iloc[j_cab]) and len(textos(df.iloc[j_cab - 1])) >= 2:
            j_cab -= 1
        if j_cab < inicio:
            bloque = df.iloc[j_cab:inicio]
            cab = [" ".join(str(v).strip() for v in bloque[c] if isinstance(v, str) and v.strip()) or None
                   for c in df.columns]
        # texto de las filas anteriores (títulos, notas del tipo «Unidades: kW»)
        preambulo = " ".join(" ".join(textos(df.iloc[k])) for k in range(max(0, j_cab - 30), j_cab))
        return cab, df.iloc[inicio:], preambulo

    # ------------------------------------------------------------------
    @staticmethod
    def _clasificar_columnas(datos, nombres):
        muestra = datos.head(3000)
        info = {}
        for c in datos.columns:
            s = muestra[c]
            no_vacios = s[~s.map(_vacio)]
            if len(no_vacios) == 0:
                continue
            frac_fecha = no_vacios.map(_parece_fecha).mean()
            nums = a_numero(no_vacios)
            frac_num = nums.notna().mean()
            distintos = set(nums.dropna().round(6).unique()[:50])
            info[c] = dict(frac_fecha=frac_fecha, frac_num=frac_num, distintos=distintos,
                           nombre=nombres[c] if c < len(nombres) else "",
                           cups=no_vacios.astype(str).str.strip().str.match(_RE_CUPS).mean())

        col = {"formato": "largo"}
        # CUPS
        for c, i in info.items():
            if i["cups"] > 0.8 or "cups" in i["nombre"]:
                col["cups"] = c
                break

        # columnas de fecha
        fechas = [c for c, i in info.items() if i["frac_fecha"] > 0.9]
        fecha_hora = fecha = None
        for c in fechas:
            vals = muestra[c][~muestra[c].map(_vacio)]
            tiene_hora = vals.map(lambda v: (isinstance(v, (datetime, pd.Timestamp)) and
                                             (v.hour or v.minute)) or
                                  (isinstance(v, str) and bool(re.search(r"\d{1,2}:\d{2}", v)))).mean()
            if tiene_hora > 0.3 and fecha_hora is None:
                fecha_hora = c
            elif fecha is None:
                fecha = c
        if fecha_hora is not None:
            col["fecha_hora"] = fecha_hora
            col["nombre_fecha"] = info[fecha_hora]["nombre"] or f"columna {fecha_hora + 1}"
        elif fecha is not None:
            col["fecha"] = fecha
            col["nombre_fecha"] = info[fecha]["nombre"] or f"columna {fecha + 1}"
        else:
            raise ValueError("No se ha identificado ninguna columna de fecha.")
        usadas = {col.get("cups"), fecha_hora, fecha}
        if fecha_hora is not None:
            for c, i in info.items():
                if c in usadas:
                    continue
                if any(k in i["nombre"] for k in ("verano", "invierno", "estacion", "bandera", "flag", "dst"))                         or (not i["nombre"] and i["distintos"] == {0.0, 1.0} and c == fecha_hora + 1):
                    col["bandera"] = c
                    usadas.add(c)
                    break

        def es_bandera(i):
            return i["distintos"] <= {0.0, 1.0} or any(k in i["nombre"] for k in
                                                      ("verano", "invierno", "estacion", "bandera", "flag", "dst"))

        if "fecha" in col:
            # columna de hora
            cand_hora = [c for c, i in info.items() if c not in usadas and
                         (i["nombre"] in ("hora", "hour", "h", "horas", "hora fin", "hora final", "periodo")
                          or i["nombre"].startswith("hora"))]
            if not cand_hora:
                for c, i in info.items():
                    if c in usadas or es_bandera(i):
                        continue
                    vals = muestra[c][~muestra[c].map(_vacio)].astype(str)
                    if vals.str.match(_RE_HHMM).mean() > 0.9 or \
                            muestra[c].map(lambda v: isinstance(v, (time, timedelta))).mean() > 0.9:
                        cand_hora = [c]
                        break
                    d = i["distintos"]
                    if d and all(float(x).is_integer() for x in d) and min(d) >= 0 and max(d) <= 25 and len(d) >= 20:
                        cand_hora = [c]
                        break
            if cand_hora:
                col["hora"] = cand_hora[0]
                col["nombre_hora"] = info[cand_hora[0]]["nombre"] or f"columna {cand_hora[0] + 1}"
                usadas.add(cand_hora[0])
                # columna de cuarto
                for c, i in info.items():
                    if c in usadas:
                        continue
                    d = i["distintos"]
                    if any(k in i["nombre"] for k in ("cuarto", "qh", "minuto", "cuart")) or \
                            (d and (d <= {1.0, 2.0, 3.0, 4.0} or d <= {0.0, 15.0, 30.0, 45.0}) and len(d) >= 3):
                        col["cuarto"] = c
                        col["nombre_cuarto"] = i["nombre"] or f"columna {c + 1}"
                        usadas.add(c)
                        break
            else:
                numericas = [c for c, i in info.items() if c not in usadas and i["frac_num"] > 0.5]
                if len(numericas) >= 23:
                    col["formato"] = "ancho"
                    col["valores"] = numericas
                    col["nombre_valor"] = " ".join(info[c]["nombre"] for c in numericas[:2])
                    return col
                raise ValueError("Hay columna de fecha pero no se ha encontrado la columna de hora.")

        # columna de valor
        excluir = ("reactiva", "saliente", "generac", "export", "exced", "metodo", "calidad",
                   "flag", "bandera", "verano", "invierno", "estacion", "cuarto", "firmeza", "tipo")
        positivos = [("activa entrante", 6), ("entrante", 5), ("consumo", 4), ("potencia", 4),
                     ("energia", 3), ("kwh", 3), ("kw", 3), ("activa", 2), ("valor", 1),
                     ("medida", 1), ("lectura", 1)]
        candidatas = []
        for c, i in info.items():
            if c in usadas or i["frac_num"] < 0.9 or es_bandera(i):
                continue
            n = i["nombre"]
            if any(k in n for k in excluir) or n in ("as", "r1", "r2", "r3", "r4", "hora", "h"):
                continue
            puntos = max([p for k, p in positivos if k in n] + [0])
            if re.match(r"^ae\b", n):
                puntos = 6
            if n in ("r", "q"):
                continue
            candidatas.append((-puntos, c))
        if not candidatas:
            raise ValueError("No se ha encontrado la columna de valores de consumo/potencia.")
        candidatas.sort()
        c = candidatas[0][1]
        col["valor"] = c
        col["nombre_valor"] = info[c]["nombre"] or f"columna {c + 1}"
        usadas.add(c)
        # columna con la unidad de cada registro («Unidad», «Magnitud»… o valores kW / kWh)
        for c, i in info.items():
            if c in usadas:
                continue
            vals = muestra[c][~muestra[c].map(_vacio)].astype(str).map(_norm)
            son_unidades = vals.map(lambda v: v in _TOKENS or v in ("kw h", "kw-h", "mw h")).mean() > 0.9
            if son_unidades or any(k in i["nombre"] for k in ("unidad", "unit", "magnitud", "uom")):
                col["unidad"] = c
                col["nombre_unidad"] = i["nombre"] or f"columna {c + 1}"
                break
        return col

    # ------------------------------------------------------------------
    @staticmethod
    def _unidades_en(texto):
        """Unidades explícitas (kW, kWh, Wh, MW, MWh, W) que aparecen en un texto, sin repetir."""
        res = []
        for m in _RE_UNIDAD.finditer(_norm(texto)):
            u = _TOKENS[re.sub(r"[\s.-]", "", m.group(1))]
            if u not in res:
                res.append(u)
        return res

    @classmethod
    def _detectar_unidad(cls, nombre):
        """Unidad a partir del título de la columna: símbolo explícito o, si no hay, palabras clave."""
        explicitas = cls._unidades_en(nombre)
        if explicitas:
            return explicitas[0]
        n = _norm(nombre)
        if "potencia" in n:
            return "kW"
        if any(k in n for k in ("consumo", "energia", "entrante")) or re.match(r"^ae\b", n):
            return "kWh"
        return None

    def _unidad_automatica(self, col, datos, preambulo):
        """(unidad, de dónde se ha sacado) o (None, "") si el fichero no la indica."""
        titulo = col.get("nombre_valor", "")
        explicitas = self._unidades_en(titulo)
        if explicitas:
            return explicitas[0], f"indicada en el título de la columna «{titulo}»"
        if "unidad" in col:
            vals = [u[0] for u in datos[col["unidad"]].astype(str).map(self._unidades_en) if u]
            if vals:
                u = max(set(vals), key=vals.count)
                return u, f"indicada en la columna de unidades «{col['nombre_unidad']}»"
        nota = self._unidades_en(preambulo)
        if len(nota) == 1:
            return nota[0], "indicada en el texto encima de la tabla"
        u = self._detectar_unidad(titulo)
        if u:
            return u, f"deducida de la palabra del título «{titulo}»"
        return None, ""

    # ------------------------------------------------------------------
    def _construir_registros(self, datos, col, zona, convenio, avisos):
        """
        Devuelve (registros, paso_min, convenio_texto, ordinal).
        registros: DataFrame con dia (Timestamp), minuto (minuto de inicio del
        intervalo dentro del día) y valor. Si `ordinal` es True, `minuto` es la
        posición cronológica real dentro del día (no la hora de reloj).
        """
        if col["formato"] == "ancho":
            fechas = _a_fecha_hora(datos[col["fecha"]], zona)[0]
            n_col = len(col["valores"])
            paso = 60 if n_col <= 25 else (30 if n_col <= 50 else 15)
            partes = []
            for k, c in enumerate(col["valores"]):
                partes.append(pd.DataFrame({"dia": fechas.dt.normalize(), "minuto": k * paso,
                                            "valor": a_numero(datos[c]).to_numpy()}))
            reg = pd.concat(partes, ignore_index=True).dropna()
            reg = reg.sort_values(["dia", "minuto"], kind="stable")
            return reg, paso, "Fin de periodo (columnas por hora)", True

        valores = a_numero(datos[col["valor"]])

        if "fecha_hora" in col:
            ts, hubo_24, utc = _a_fecha_hora(datos[col["fecha_hora"]], zona)
            ok = ts.notna() & valores.notna()
            ts, valores = ts[ok], valores[ok]
            utc = utc[ok] if utc is not None else None
            bandera = a_numero(datos[col["bandera"]])[ok] if "bandera" in col else None
            return self._desde_etiquetas_reloj(ts, valores, convenio, hubo_24, avisos, zona, utc, bandera)

        fechas = _a_fecha_hora(datos[col["fecha"]], zona)[0].dt.normalize()
        minutos, tipo = _a_minutos(datos[col["hora"]])
        ok = fechas.notna() & minutos.notna() & valores.notna()
        fechas, minutos, valores = fechas[ok], minutos[ok], valores[ok]
        cuarto = None
        if "cuarto" in col:
            q = a_numero(datos[col["cuarto"]])[ok]
            cuarto = (q / 15) if set(q.dropna().unique()) <= {0, 15, 30, 45} else (q - 1)

        if tipo == "reloj" and cuarto is not None and (minutos % 60 == 0).all():
            tipo = "entera"
        if tipo == "reloj":
            ts = fechas + pd.to_timedelta(minutos, unit="m")
            hubo_24 = bool((minutos >= 1440).any())
            if (minutos > 1440).any():   # "25:00" -> posiciones ordinales
                dif = minutos.groupby(fechas).diff()
                paso_r = int(dif[dif > 0].mode().iloc[0]) if (dif > 0).any() else 60
                return self._desde_ordinal(fechas, minutos - paso_r, valores, paso_r, "Fin de periodo")
            return self._desde_etiquetas_reloj(ts, valores, convenio, hubo_24, avisos, zona)

        # horas enteras (0..23 = inicio, 1..24/25 = fin de periodo)
        horas = (minutos / 60).round().astype(int)
        base1 = horas.min() >= 1
        if convenio == "Inicio de periodo":
            base1 = False
        elif convenio == "Fin de periodo":
            base1 = True
        h0 = horas - (1 if base1 else 0)
        if cuarto is None:
            grupo = pd.DataFrame({"f": fechas, "h": horas})
            por_hora = grupo.groupby(["f", "h"]).size()
            if por_hora.median() >= 4:            # 4 registros por hora sin columna de cuarto
                cuarto = grupo.groupby(["f", "h"]).cumcount().astype(float)
                paso = 15
            else:
                cuarto = pd.Series(0.0, index=fechas.index)
                paso = 60
        else:
            paso = 15
        minuto = h0 * 60 + cuarto.round().astype(int) * 15
        conv_txt = "Fin de periodo (horas 1-24)" if base1 else "Inicio de periodo (horas 0-23)"
        return self._desde_ordinal(fechas, minuto, valores, paso, conv_txt)

    # ------------------------------------------------------------------
    @staticmethod
    def _desde_ordinal(fechas, minuto, valores, paso, conv_txt):
        reg = pd.DataFrame({"dia": fechas.to_numpy(), "minuto": np.asarray(minuto, dtype=int),
                            "valor": valores.to_numpy()})
        return reg, paso, conv_txt, True

    @staticmethod
    def _desde_etiquetas_reloj(ts, valores, convenio, hubo_24, avisos, zona, utc=None, bandera=None):
        ts_ord = ts.sort_values()
        dif = ts_ord.diff().dropna().dt.total_seconds() / 60
        dif = dif[dif > 0]
        paso = int(dif.mode().iloc[0]) if len(dif) else 15
        if paso not in (15, 30, 60):
            paso = 15 if paso < 30 else 60
            avisos.append(f"Resolución temporal irregular; se interpreta como {paso} minutos.")

        if convenio == "Fin de periodo":
            fin = True
        elif convenio == "Inicio de periodo":
            fin = False
        elif hubo_24:
            fin = True
        else:
            primero, ultimo = ts_ord.iloc[0], ts_ord.iloc[-1]
            m_pri = primero.hour * 60 + primero.minute
            m_ult = ultimo.hour * 60 + ultimo.minute
            if m_pri == paso:
                fin = True
            elif m_pri == 0 and m_ult == 1440 - paso:
                fin = False
            else:
                fin = True
                avisos.append("No se ha podido deducir si la hora marca el inicio o el fin del intervalo; "
                              "se asume FIN de periodo (estándar REE).")
        delta = pd.Timedelta(minutes=paso)
        inicio = None
        tz = zona_horaria(zona)
        if utc is not None:
            inicio = (utc - delta if fin else utc).dt.tz_convert(tz).dt.tz_localize(None)
        elif fin:
            # en los días de cambio de hora, restar el intervalo en hora física (no de reloj)
            nominal = 1440 // paso
            cuenta = (ts - delta).dt.normalize().value_counts()
            dias_ch = {pd.Timestamp(d) for a in {x.year for x in cuenta.index} for d in dias_cambio_hora(a)}
            reales = [d for d in cuenta.index if d in dias_ch and cuenta[d] != nominal]
            if reales:
                try:
                    amb = bandera.fillna(0).astype(bool).to_numpy() if bandera is not None else "infer"
                    loc = ts.dt.tz_localize(tz, ambiguous=amb, nonexistent="NaT")
                    if loc.notna().all():
                        inicio = (loc - delta).dt.tz_convert(tz).dt.tz_localize(None)
                except Exception:
                    inicio = None
        if inicio is None:
            inicio = ts - delta if fin else ts
        reg = pd.DataFrame({"dia": inicio.dt.normalize().to_numpy(),
                            "minuto": (inicio.dt.hour * 60 + inicio.dt.minute).to_numpy(),
                            "valor": valores.to_numpy()})
        return reg, paso, ("Fin de periodo" if fin else "Inicio de periodo"), False

    # ------------------------------------------------------------------
    @staticmethod
    def _normalizar(reg, paso, ordinal, zona):
        """Pasa a rejilla regular de `paso` minutos en hora de reloj, con 24 h por día."""
        reg = reg.copy()
        h_cambio = hora_cambio_horario(zona) * 60
        dias_ch = []
        anios = sorted({d.year for d in pd.DatetimeIndex(reg["dia"])})
        marzo = {pd.Timestamp(dias_cambio_hora(a)[0]) for a in anios}
        octubre = {pd.Timestamp(dias_cambio_hora(a)[1]) for a in anios}

        if ordinal:
            reg = _a_hora_reloj(reg, paso, ordinal, zona)

        reg["inicio"] = pd.to_datetime(reg["dia"]) + pd.to_timedelta(reg["minuto"], unit="m")

        # duplicados: en octubre la hora repetida se promedia
        cuenta = reg.groupby("inicio").size()
        repetidos = cuenta[cuenta > 1]
        en_oct = repetidos.index.normalize().isin(list(octubre)) & \
            (repetidos.index.hour * 60 >= h_cambio) & (repetidos.index.hour * 60 < h_cambio + 60)
        duplicados = int((repetidos[~en_oct] - 1).sum())
        for d in sorted({i.date() for i in repetidos.index[en_oct]}):
            dias_ch.append(f"{d:%d/%m/%Y} (25 h: hora repetida promediada)")
        serie = reg.groupby("inicio")["kw"].mean()

        # rejilla completa
        rejilla = pd.date_range(serie.index.min().normalize(),
                                serie.index.max().normalize() + pd.Timedelta(days=1) - pd.Timedelta(minutes=paso),
                                freq=f"{paso}min")
        serie = serie.reindex(rejilla)
        falta = serie.isna()
        marzo_falta = falta & serie.index.normalize().isin(list(marzo)) & \
            (serie.index.hour * 60 >= h_cambio) & (serie.index.hour * 60 < h_cambio + 60)
        for d in sorted({i.date() for i in serie.index[marzo_falta]}):
            dias_ch.append(f"{d:%d/%m/%Y} (23 h: hora inexistente rellenada)")

        # días incompletos al principio/fin (menos de la mitad de datos): se descartan
        por_dia = (~falta).groupby(serie.index.normalize()).mean()
        buenos = por_dia[por_dia >= 0.5].index
        if len(buenos):
            serie = serie[(serie.index >= buenos.min()) & (serie.index < buenos.max() + pd.Timedelta(days=1))]
            falta = serie.isna()
            marzo_falta = marzo_falta.reindex(serie.index, fill_value=False)

        # relleno: promedio de la hora anterior y posterior
        n = 60 // paso
        for _ in range(6):
            hueco = serie.isna()
            if not hueco.any():
                break
            estim = (serie.shift(n) + serie.shift(-n)) / 2
            serie = serie.where(~hueco, estim)
        if serie.isna().any():
            serie = serie.interpolate(limit_direction="both")

        # los huecos de más de 2 días (p. ej. un mes sin fichero) NO se inventan: se dejan sin datos
        tramo = (falta != falta.shift()).cumsum()
        largo = falta & falta.groupby(tramo).transform("size").ge(2 * 1440 // paso)
        sin_datos = []
        for _, bloque in serie.index[largo].to_series().groupby(tramo[largo]):
            sin_datos.append((bloque.iloc[0], bloque.iloc[-1] + pd.Timedelta(minutes=paso)))
        huecos = int((falta & ~marzo_falta & ~largo).sum())
        salida = pd.DataFrame({"inicio": serie.index, "kw": serie.to_numpy(), "estimado": falta.to_numpy()})
        salida = salida[~largo.to_numpy()].reset_index(drop=True)
        return salida, dias_ch, huecos, duplicados, sin_datos

    # ------------------------------------------------------------------
    @staticmethod
    def _a_cuartos(serie, paso):
        rep = paso // 15
        inicio = np.repeat(serie["inicio"].to_numpy(), rep) + \
            np.tile(np.arange(rep) * np.timedelta64(15, "m"), len(serie))
        return pd.DataFrame({"inicio": inicio, "kw": np.repeat(serie["kw"].to_numpy(), rep),
                             "estimado": True})

    # ------------------------------------------------------------------
    @staticmethod
    def _recortar_12_meses(serie, avisos):
        ini = serie["inicio"]
        primero = ini.iloc[0]
        ultimo_fin = ini.iloc[-1] + pd.Timedelta(minutes=15)
        meses = pd.period_range(primero, ini.iloc[-1], freq="M")
        completos = [m for m in meses
                     if m.start_time >= primero and (m + 1).start_time <= ultimo_fin]
        if len(completos) >= 12:
            sel = completos[-12:]
            desde, hasta = sel[0].start_time, sel[-1].end_time
            if len(meses) > 12 or len(completos) != len(meses):
                avisos.append(f"La curva abarca {len(meses)} meses; se usan los 12 últimos meses completos "
                              f"({sel[0].strftime('%m/%Y')} – {sel[-1].strftime('%m/%Y')}).")
            return serie[(ini >= desde) & (ini <= hasta)]
        dias = (ultimo_fin - primero).days
        if dias > 366:
            desde = ultimo_fin - pd.Timedelta(days=365)
            avisos.append("Se usan los últimos 365 días de la curva.")
            return serie[ini >= desde]
        if dias < 360:
            avisos.append(f"La curva solo cubre {dias} días: los importes NO corresponden a un año completo.")
        return serie


class LectorCurvas:
    """
    Varios ficheros de curva leídos como uno solo: por ejemplo, un Excel por mes y por CUPS.
    Cada fichero se interpreta por separado (formato, unidad, convenio horario) y después se
    unen los registros de cada CUPS. Misma interfaz que LectorCurva.
    """

    def __init__(self, rutas):
        self.rutas = [str(r) for r in rutas]
        self.lectores, errores = [], []
        for r in self.rutas:
            try:
                self.lectores.append(LectorCurva(r))
            except Exception as e:
                errores.append(f"{Path(r).name}: {e}")
        if errores:
            raise ValueError("No se han podido leer estos ficheros:\n" + "\n".join(errores))
        self.ruta = f"{len(self.rutas)} ficheros"
        self.hojas = [None]          # cada fichero usa su hoja con la curva
        self.hoja_defecto = None
        # CUPS de cada fichero (columna CUPS, nombre del fichero o texto encima de la tabla)
        self._cups = []
        for lector in self.lectores:
            try:
                self._cups.append(lector.cups_en_fichero())
            except Exception:
                self._cups.append(([], False))
        self.cups_disponibles = list(dict.fromkeys(c for lista, _ in self._cups for c in lista))

    def procesar(self, zona="Península", unidad="Automática", convenio="Automático", hoja=None, cups=None):
        objetivo = cups if cups in self.cups_disponibles else (self.cups_disponibles[0] if self.cups_disponibles
                                                                else None)
        lista, sin_cups, errores = [], [], []
        for lector, (cups_fichero, _) in zip(self.lectores, self._cups):
            nombre = Path(lector.ruta).name
            if objetivo is not None and cups_fichero and objetivo not in cups_fichero:
                continue                                   # fichero de otro CUPS
            if objetivo is not None and not cups_fichero:
                if len(self.cups_disponibles) > 1:
                    sin_cups.append(nombre)                # no se sabe de qué CUPS es
                    continue
            try:
                lista.append(lector.registros(zona, unidad, convenio, None, objetivo))
            except Exception as e:
                errores.append(f"{nombre}: {e}")
        if not lista:
            raise ValueError("Ningún fichero tiene datos válidos" + (f" del CUPS {objetivo}" if objetivo else "")
                             + (":\n" + "\n".join(errores) if errores else "."))
        curva = _componer(lista, zona)
        if objetivo:
            curva.cups = objetivo
        if sin_cups:
            curva.avisos.append("Estos ficheros no indican su CUPS (ni en una columna, ni en el nombre, ni encima "
                                "de la tabla) y hay varios CUPS, así que no se han usado: " + ", ".join(sin_cups))
        curva.avisos += [f"No se ha podido leer {e}" for e in errores]
        if len(self.cups_disponibles) > 1:
            curva.avisos.append(f"Los ficheros contienen {len(self.cups_disponibles)} CUPS; se usa {objetivo}.")
        self.cups_disponibles = list(dict.fromkeys(c for lista_c, _ in self._cups for c in lista_c))
        return curva


def abrir_curvas(rutas):
    """LectorCurva para un fichero; LectorCurvas para varios."""
    rutas = list(rutas)
    return LectorCurva(rutas[0]) if len(rutas) == 1 else LectorCurvas(rutas)
