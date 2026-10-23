# -*- coding: utf-8 -*-
"""
ESTUDIO DE POTENCIA: logica de la pagina (sin Streamlit)
========================================================
Lo que hacia la interfaz tkinter del programa de escritorio EstudioPotencia (gui.py y
gui_gemweb.py) con el calculo de potencia/: leer las curvas (ficheros subidos o Gemweb),
la zona de cada suministro a partir de su codigo postal en el inventario de Gemweb, calcular
la situacion actual y las propuestas, las tablas de resultados y anexos, los graficos y el
PDF, y empaquetar un estudio para guardarlo en la base de datos.
"""

import gzip
import pickle
import re
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from matplotlib.figure import Figure

from potencia import informe_pdf
from potencia.calculo import (N_P, Estudio, ResumenMultipunto, Tarifa, cargar_precios, evaluar,
                              fmt, fmt_pot, opciones_inversion_defecto, texto_conceptos,
                              validar_potencias)
from potencia.calendario import ZONAS
from potencia.lector_curva import CONVENIOS, UNIDADES, LectorCurva, abrir_curvas

TARIFAS = ["6.1TD", "3.0TD"]
MAX_PROPUESTAS = 6
EXTENSIONES = ["csv", "txt", "xlsx", "xlsm", "xls", "dat"]
TODOS = "Todos (suma)"
PENINSULA = ZONAS[0]
CONCEPTOS_INVERSION = [("acometida", "¿Modifica acometida?"), ("acceso", "Derechos de acceso"),
                       ("extension", "Derechos de extensión"),
                       ("enganche", "Derechos de enganche")]
AVISO_HORARIA = ("En Gemweb este suministro tiene lectura HORARIA: cada cuarto de hora es un "
                 "reparto de la energía de su hora, así que los picos de 15 minutos no son reales "
                 "y los excesos pueden ser mayores.")


def precios():
    return cargar_precios()


# ------------------------------------------------------------------- suministro
@dataclass
class Suministro:
    """Un punto de suministro del estudio: datos, curva, zona, potencias y propuestas.
    propuestas: [{"pc": [6] o None, "opciones": {...} o None}] desde la Propuesta 1 (la
    optima: su pc se ignora); opciones None = las de por defecto (opciones_inversion_defecto)."""
    cups: str = ""
    curva: object = None                       # potencia.lector_curva.CurvaCargada
    zona: str = PENINSULA
    zona_origen: str = ""
    zona_manual: bool = False
    tarifa: str = TARIFAS[0]
    denominacion: str = ""
    direccion: str = ""
    actual: list = field(default_factory=lambda: [None] * N_P)
    propuestas: list = field(default_factory=lambda: [{"pc": None, "opciones": None},
                                                      {"pc": None, "opciones": None}])
    avisos_extra: list = field(default_factory=list)      # de Gemweb: se conservan al reinterpretar
    zona_curva: str = None                                # zona con la que se leyo la curva
    estudio: object = None
    optimo: object = None
    escenarios: list = None
    _clave: tuple = None

    @property
    def nombre(self):
        return self.denominacion.strip() or self.cups or "Suministro"

    def invalidar(self):
        self.estudio = self.optimo = self.escenarios = self._clave = None


# ------------------------------------------------------------------- curvas
def guardar_subidos(ficheros, carpeta) -> list:
    """[(nombre, bytes)] -> rutas en `carpeta` con su nombre original: el lector saca de el el
    CUPS y el formato."""
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    rutas = []
    for nombre, contenido in ficheros:
        ruta = carpeta / Path(nombre).name
        ruta.write_bytes(contenido)
        rutas.append(ruta)
    return rutas


def abrir(rutas):
    """LectorCurva (un fichero) o LectorCurvas (varios)."""
    return abrir_curvas([Path(r) for r in rutas])


def hoja_elegida(lector, hoja):
    return next((h for h in lector.hojas if str(h) == str(hoja)), lector.hoja_defecto)


def interpretar(lector, zona, ajustes=None, cups=None, multipunto=False):
    """Curva de un CUPS con la zona y los ajustes de lectura (hoja, unidad, convenio)."""
    a = ajustes or {}
    curva = lector.procesar(zona, a.get("unidad") or UNIDADES[0],
                            a.get("convenio") or CONVENIOS[0], hoja_elegida(lector, a.get("hoja")),
                            cups or a.get("cups"))
    if multipunto:          # en multipunto se usan todos los CUPS: el aviso «se usa X» no aplica
        curva.avisos = [x for x in curva.avisos if not re.search(r"contienen? \d+ CUPS; se usa", x)]
    return curva


def leer_suministro(s, lector, ajustes=None, multipunto=False):
    """(Re)lee la curva del suministro con su zona, conservando los avisos de Gemweb."""
    s.curva = interpretar(lector, s.zona, ajustes, s.cups if multipunto else None, multipunto)
    s.curva.avisos += [a for a in s.avisos_extra if a not in s.curva.avisos]
    s.zona_curva = s.zona
    s.invalidar()
    return s.curva


def resumen_interpretacion(curva, lector=None, multipunto=False) -> list:
    """Lineas con lo que se ha entendido del fichero (el dialogo de carga del programa)."""
    d = curva.datos
    lineas = [
        "Columnas usadas: " + ", ".join(f"{k.replace('nombre_', '')} = «{v}»"
                                         for k, v in curva.columnas.items()
                                         if k.startswith("nombre_")),
        "Resolución original: " + ("CUARTOHORARIA (15 min)" if curva.resolucion_min == 15
                                   else f"HORARIA ({curva.resolucion_min} min)")
        + f" · Hora del registro: {curva.convenio}",
        f"Unidad: {curva.descripcion_unidad}",
        (f"Periodo: {curva.fecha_inicio:%d/%m/%Y %H:%M} – {curva.fecha_fin:%d/%m/%Y %H:%M} · "
         f"{len(d):,} cuartos de hora ({curva.n_registros:,} registros leídos)").replace(",", "."),
        f"Potencia máxima: {fmt(d['kw'].max(), 1)} kW · Energía total: "
        f"{fmt(d['kw'].sum() * 0.25)} kWh",
        f"CUPS: {curva.cups or 'no incluido en el fichero'}",
    ]
    if multipunto and lector is not None:
        n = len(lector.cups_disponibles)
        lineas.insert(0, f"ESTUDIO MULTIPUNTO: {n} CUPS detectados"
                      + (" (los datos de abajo son del primero)." if n else
                         ". El fichero no tiene columna de CUPS: no se puede hacer un estudio "
                         "multipunto."))
    if curva.ficheros:
        lineas.append(f"Ficheros leídos ({len(curva.ficheros)}), unidos en una sola curva: "
                      + ", ".join(curva.ficheros))
    if curva.dias_cambio_hora:
        lineas.append("Cambios de hora: " + "; ".join(curva.dias_cambio_hora))
    return lineas


def primeros_registros(curva, n=12) -> pd.DataFrame:
    d = curva.datos.head(n)
    return pd.DataFrame({"Inicio del cuarto": d["inicio"].dt.strftime("%d/%m/%Y %H:%M"),
                         "Fin del cuarto": (d["inicio"] + pd.Timedelta(minutes=15))
                         .dt.strftime("%d/%m/%Y %H:%M"),
                         "Potencia (kW)": d["kw"].round(2),
                         "Estimado": np.where(d["estimado"], "Sí", "")})


def registro_carga(suministros, multipunto=False, errores=()) -> str:
    """Texto de la pestaña «Registro de carga»."""
    lineas = []
    for s in suministros:
        c = s.curva
        if c is None:
            continue
        if multipunto:
            lineas.append(f"══ CUPS {s.cups} ══")
        lineas += [f"Fichero: {c.archivo}", f"Hoja: {c.hoja}" if c.hoja else None,
                   f"Zona: {s.zona}" + (f" ({s.zona_origen})" if s.zona_origen else ""),
                   f"Resolución original: {c.resolucion_min} min · Hora del registro: "
                   f"{c.convenio}",
                   f"Unidad: {c.descripcion_unidad}",
                   f"Registros leídos: {c.n_registros:,}".replace(",", "."),
                   f"Periodo del estudio: {c.fecha_inicio:%d/%m/%Y %H:%M} – "
                   f"{c.fecha_fin:%d/%m/%Y %H:%M}",
                   f"Intervalos rellenados (huecos): {c.huecos_rellenados}",
                   "Cambios de hora: " + ("; ".join(c.dias_cambio_hora) or "ninguno")]
        if c.ficheros:
            lineas += [f"Ficheros ({len(c.ficheros)}):"] + [f"  · {f}" for f in c.ficheros]
        if c.sin_datos:
            lineas.append("Periodos SIN DATOS (no se rellenan): " + "; ".join(
                f"{a:%d/%m/%Y %H:%M} – {b:%d/%m/%Y %H:%M}" for a, b in c.sin_datos))
        lineas.append("Avisos:")
        lineas += [f"  • {a}" for a in c.avisos] or ["  (ninguno)"]
        lineas.append("")
    if errores:
        lineas += ["CUPS no interpretados:"] + [f"  • {e}" for e in errores]
    return "\n".join(l for l in lineas if l is not None)


# ------------------------------------------------------------------- Gemweb
def zona_por_codigo_postal(cp):
    """Sistema electrico a partir del codigo postal (None si no se puede deducir)."""
    cp = re.sub(r"\D", "", str(cp or ""))
    if not 4 <= len(cp) <= 5:
        return None
    return {"07": "Illes Balears", "35": "Canarias", "38": "Canarias", "51": "Ceuta",
            "52": "Melilla"}.get(cp.zfill(5)[:2], PENINSULA)


def zona_de(sumi, centro=None) -> tuple:
    """(zona, de donde sale) del suministro de Gemweb: el codigo postal de su direccion (si
    falta, el del centro de consumo). Sin el: Peninsula, a revisar."""
    for registro, origen in ((sumi, "Gemweb"), (centro, "centro de consumo en Gemweb")):
        cp = (registro or {}).get("codi_postal")
        zona = zona_por_codigo_postal(cp)
        if zona:
            return zona, f"CP {str(cp).strip()}, {origen}"
    if not sumi:
        return PENINSULA, "no está en Gemweb: revisa la zona"
    return PENINSULA, "sin código postal en Gemweb: revisa la zona"


def numero(texto):
    try:
        return float(str(texto).replace(",", "."))
    except (TypeError, ValueError):
        return None


def lista_cups(texto):
    """CUPS unicos de un texto (separados por saltos de linea, espacios, comas o punto y coma)."""
    return list(dict.fromkeys(c.strip().upper() for c in re.split(r"[\s,;]+", texto or "")
                              if c.strip()))


def periodo_defecto(hoy=None):
    """Ultimos 12 meses completos: (primer dia, ultimo dia)."""
    hoy = hoy or date.today()
    fin = hoy.replace(day=1) - timedelta(days=1)
    ini = (fin.replace(day=1) - timedelta(days=1)).replace(day=1)
    for _ in range(10):
        ini = (ini - timedelta(days=1)).replace(day=1)
    return ini, fin


def direccion_gemweb(sumi) -> str:
    partes = [sumi.get("direccio"), " ".join(x for x in (sumi.get("codi_postal"),
                                                          sumi.get("poblacio")) if x)]
    return ", ".join(x.strip() for x in partes if x and x.strip())


def rellenar_desde_gemweb(s, sumi):
    """Tarifa (si es 6.1TD o 3.0TD), potencias contratadas (si estan las 6), denominacion y
    direccion del inventario de Gemweb."""
    tarifa = (sumi.get("tarifa_acces") or "").strip()
    if tarifa in TARIFAS:
        s.tarifa = tarifa
    potencias = [numero(sumi.get(f"pot_contract_p{k}")) for k in range(1, 7)]
    if all(p and p > 0 for p in potencias):
        s.actual = potencias
    s.denominacion = sumi.get("nom") or s.denominacion
    s.direccion = direccion_gemweb(sumi) or s.direccion


def escribir_csv(curvas, ruta):
    """Curvas descargadas ({cups: DataFrame[fecha, kwh]}) en un CSV que el lector interpreta
    directamente: CUPS; fecha-hora de fin del cuarto; energia en kWh."""
    partes = [pd.DataFrame({"CUPS": cups, "Fecha hora (fin del cuarto)": df["fecha"],
                            "Consumo (kWh)": df["kwh"]}) for cups, df in curvas.items() if len(df)]
    pd.concat(partes, ignore_index=True).to_csv(ruta, sep=";", index=False, decimal=",",
                                                encoding="utf-8-sig")


def descargar_gemweb(cliente, lista, desde, hasta, carpeta, al_avanzar=None) -> tuple:
    """Curva cuartohoraria de cada CUPS de Gemweb (un dia de margen a cada lado: el lector se
    queda con los meses completos). cliente: gemweb.ClienteGemweb de la plataforma.
    -> (ruta del CSV o None, {cups: suministro de Gemweb}, [avisos])."""
    curvas, info, avisos = {}, {}, []
    total = len(lista)
    for i, cups in enumerate(lista):
        sumi = cliente.buscar_suministro(cups)
        if sumi is None:
            avisos.append(f"{cups}: no está en el inventario de Gemweb.")
            continue
        tarifa = (sumi.get("tarifa_acces") or "").strip()
        if tarifa not in TARIFAS:
            avisos.append(f"{cups}: tarifa {tarifa or 'desconocida'} (el estudio es para 6.1TD y "
                          "3.0TD); se descarga igualmente.")

        def avance(n, n_tramos, a, b, i=i, cups=cups):
            if al_avanzar:
                al_avanzar((i + (n - 1) / n_tramos) / total,
                           f"{cups}: tramo {n}/{n_tramos} ({a:%d/%m/%Y} – {b:%d/%m/%Y})")
        valores, fallidos = cliente.descargar(sumi["id"], desde - timedelta(days=1),
                                              hasta + timedelta(days=1), avance)
        if not valores:
            avisos.append(f"{cups}: Gemweb no tiene telelecturas cuartohorarias en el periodo."
                          + (f" ({fallidos[0]})" if fallidos else ""))
            continue
        avisos += [f"{cups}: tramo sin descargar {f}" for f in fallidos]
        ultimo = datetime.strptime(valores[-1][0][:16], "%Y-%m-%d %H:%M") - timedelta(minutes=15)
        if ultimo.date() < hasta:
            avisos.append(f"{cups}: Gemweb solo tiene datos hasta el {ultimo:%d/%m/%Y %H:%M}.")
        sumi["_lectura_horaria"] = sumi.get("tipus_lectura") != "quart_hora"
        curvas[cups] = pd.DataFrame(valores, columns=["fecha", "kwh"])
        info[cups] = sumi
    if not curvas:
        return None, info, avisos
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"gemweb_{datetime.now():%Y%m%d_%H%M%S}.csv"
    escribir_csv(curvas, ruta)
    return ruta, info, avisos


def avisos_gemweb(cups, avisos, sumi) -> list:
    """Avisos de la descarga de un CUPS (se guardan en el suministro y van a su curva)."""
    res = [a for a in avisos if a.startswith(cups)]
    if (sumi or {}).get("_lectura_horaria"):
        res.append(AVISO_HORARIA)
    return res


# ------------------------------------------------------------------- calculo
def preparar(s, precios_):
    """Estudio y optimo del suministro (se reutilizan si no cambian curva, zona y tarifa)."""
    clave = (id(s.curva), s.zona, s.tarifa)
    if s.estudio is None or s._clave != clave:
        s.estudio = Estudio(s.curva.datos, s.zona, Tarifa.desde_precios(precios_, s.tarifa),
                            precios_.get("festivos"))
        s.optimo = s.estudio.optimo()
        s._clave = clave
    return s.estudio


def propuestas_de(s) -> list:
    """[(nombre, pc, opciones)] para evaluar; las propuestas vacias toman el optimo. Lanza
    ValueError si alguna no es valida."""
    res = []
    for k, p in enumerate(s.propuestas, start=1):
        pc = list(s.optimo) if k == 1 or not p.get("pc") or any(
            v is None for v in p["pc"]) else [float(v) for v in p["pc"]]
        err = validar_potencias(pc)
        if err:
            raise ValueError(f"Propuesta {k}: {err}")
        opciones = p.get("opciones") or opciones_inversion_defecto(pc, s.actual)
        res.append((f"Propuesta {k}", pc, opciones))
    return res


def calcular(s, precios_):
    """Escenarios del suministro. Devuelve None o un mensaje de error."""
    if s.curva is None:
        return "Falta la curva de carga."
    if any(p is None or p <= 0 for p in s.actual):
        return "Faltan las 6 potencias contratadas actuales."
    est = preparar(s, precios_)
    try:
        propuestas = propuestas_de(s)
    except ValueError as e:
        return str(e)
    s.escenarios = evaluar(est, propuestas, s.actual)
    return None


def avisos_calculo(s) -> list:
    """Lo que conviene revisar despues de calcular."""
    res = []
    if validar_potencias(s.actual):
        res.append("Las potencias actuales no cumplen la condición P1 ≤ … ≤ P6.")
    pmax = s.estudio.max_demanda()
    if pmax > 3 * max(s.actual) or pmax < 0.15 * max(s.actual):
        res.append(f"La potencia máxima de la curva ({fmt(pmax, 1)} kW) es muy distinta de la "
                   f"contratada ({fmt(max(s.actual))} kW). Revisa la unidad de la curva "
                   "(kW / kWh).")
    if s.tarifa == "3.0TD" and max(s.optimo) <= 50:
        res.append("La potencia óptima no supera 50 kW: con 3.0TD ≤ 50 kW el punto de medida "
                   "sería tipo 4 y los excesos se facturarían por maxímetro, no por cuartos de "
                   "hora.")
    return res


def resumen(suministros):
    return ResumenMultipunto([(s.estudio, s.escenarios) for s in suministros])


# ------------------------------------------------------------------- tablas
def _corto(nombre):
    return nombre if nombre in ("Actual", "Óptima") else f"Prop. {nombre.split()[-1]}"


def tabla_resultados(escenarios) -> pd.DataFrame:
    """Actual y propuestas: potencias, terminos, ahorro, inversion y PRS."""
    filas = []
    for k, e in enumerate(escenarios):
        fila = {"Tipo": e.nombre + (" (óptima)" if k == 1 else "")}
        fila.update({f"P{i + 1} (kW)": fmt_pot(e.pc[i]) for i in range(N_P)})
        fila.update({"T. Fijo (€)": fmt(e.coste.total_fijo),
                     "T. Excesos (€)": fmt(e.coste.total_exceso),
                     "Total (€)": fmt(e.coste.total),
                     "Ahorro (€) *": "-" if k == 0 else fmt(e.ahorro),
                     "%": "-" if k == 0 else f"{fmt(100 * e.ahorro_pct, 1)} %",
                     "Inversión (€)": fmt(e.inversion["total"], 2)
                     if e.inversion.get("total") else "-",
                     "PRS (años)": fmt(e.prs, 2) if e.prs is not None else "-"})
        filas.append(fila)
    return pd.DataFrame(filas)


def tabla_mensual(etiquetas, escenarios) -> pd.DataFrame:
    """Coste de la potencia mes a mes de cada escenario y ahorro respecto al actual."""
    actual, prop = escenarios[0], escenarios[1:]
    cols = ["Mes"] + [_corto(e.nombre) for e in escenarios] + \
        [("Ahorro" if e.nombre == "Óptima" else f"Ahorro P{e.nombre.split()[-1]}") for e in prop]
    filas = [[m] + [fmt(e.coste.total_mes[i]) for e in escenarios]
             + [fmt(actual.coste.total_mes[i] - e.coste.total_mes[i]) for e in prop]
             for i, m in enumerate(etiquetas)]
    filas.append(["Total"] + [fmt(e.coste.total) for e in escenarios]
                 + [fmt(e.ahorro) for e in prop])
    return pd.DataFrame(filas, columns=cols)


def tabla_energia(est) -> pd.DataFrame:
    e = est.energia_kwh()
    cols = ["Mes"] + [f"P{i}" for i in range(1, 7)] + ["Total"]
    filas = [[m] + [fmt(v) for v in e[i]] + [fmt(e[i].sum())]
             for i, m in enumerate(est.etiquetas_meses)]
    tot = e.sum(0)
    filas.append(["Total"] + [fmt(v) for v in tot] + [fmt(tot.sum())])
    filas.append(["%"] + [f"{fmt(100 * v / tot.sum(), 1)} %" if tot.sum() else "-"
                          for v in tot] + [""])
    return pd.DataFrame(filas, columns=cols)


def tabla_maximos(est) -> pd.DataFrame:
    m = est.maximos()
    cols = ["Mes"] + [f"P{i}" for i in range(1, 7)] + ["Máximo"]
    filas = [[mes] + [fmt(v) for v in m[i]] + [fmt(m[i].max())]
             for i, mes in enumerate(est.etiquetas_meses)]
    filas.append(["Máximo"] + [fmt(v) for v in m.max(0)] + [fmt(m.max())])
    return pd.DataFrame(filas, columns=cols)


def tabla_maximos_cups(suministros) -> pd.DataFrame:
    """Todos (suma): potencia maxima anual de cada CUPS por periodo."""
    filas = []
    for s in suministros:
        m = s.estudio.maximos().max(0)
        filas.append([s.cups] + [fmt(v) for v in m] + [fmt(m.max())])
    return pd.DataFrame(filas, columns=["CUPS"] + [f"P{i}" for i in range(1, 7)] + ["Máximo"])


def tabla_detalle(est, esc) -> pd.DataFrame:
    """Fijo y excesos por periodo y mes de un escenario."""
    c = esc.coste
    cols = (["Mes", "Días"] + [f"Fijo P{i}" for i in range(1, 7)] + ["Total fijo"]
            + [f"Exceso P{i}" for i in range(1, 7)] + ["Total excesos", "TOTAL"])
    filas = [[m, int(est.dias_mes[i])] + [fmt(v, 2) for v in c.fijo[i]]
             + [fmt(c.fijo[i].sum(), 2)] + [fmt(v, 2) for v in c.exceso[i]]
             + [fmt(c.exceso[i].sum(), 2), fmt(c.total_mes[i], 2)]
             for i, m in enumerate(est.etiquetas_meses)]
    filas.append(["Total", int(est.dias_totales)] + [fmt(v, 2) for v in c.fijo.sum(0)]
                 + [fmt(c.total_fijo, 2)] + [fmt(v, 2) for v in c.exceso.sum(0)]
                 + [fmt(c.total_exceso, 2), fmt(c.total, 2)])
    return pd.DataFrame(filas, columns=cols)


def filas_cups(suministros) -> list:
    """Una fila por CUPS para la tabla del multipunto y el PDF (actual frente a la optima)."""
    filas = []
    for s in suministros:
        act, opt = s.escenarios[0], s.escenarios[1]
        filas.append({"cups": s.cups, "denominacion": s.denominacion.strip(),
                      "tarifa": s.tarifa, "zona": s.zona, "actual": act.coste.total,
                      "optima": opt.coste.total, "ahorro": opt.ahorro, "pct": opt.ahorro_pct,
                      "inversion": opt.inversion.get("total", 0), "prs": opt.prs})
    return filas


def tabla_cups(filas, res) -> pd.DataFrame:
    datos = []
    for k, f in enumerate(filas, 1):
        datos.append([k, f["cups"], f["denominacion"] or "-", f["tarifa"], f["zona"],
                      fmt(f["actual"]), fmt(f["optima"]), fmt(f["ahorro"]),
                      f"{fmt(100 * f['pct'], 1)} %",
                      fmt(f["inversion"], 2) if f["inversion"] else "-",
                      fmt(f["prs"], 2) if f["prs"] is not None else "-"])
    act, opt = res.escenarios[:2]
    datos.append(["", "TOTAL", "", "", "", fmt(act.coste.total), fmt(opt.coste.total),
                  fmt(opt.ahorro), f"{fmt(100 * opt.ahorro_pct, 1)} %",
                  fmt(opt.inversion["total"], 2) if opt.inversion["total"] else "-",
                  fmt(opt.prs, 2) if opt.prs is not None else "-"])
    return pd.DataFrame(datos, columns=["Nº", "CUPS", "Denominación", "Tarifa", "Zona",
                                        "Coste actual", "Coste óptimo", "Ahorro", "%",
                                        "Inversión", "PRS (años)"])


def texto_notas(estudio, escenarios) -> str:
    inv = [f"{e.nombre}: {texto_conceptos(e.inversion['conceptos'])}" for e in escenarios[1:]
           if e.inversion.get("total")]
    d = estudio.tarifa.derechos
    return ("* Ahorro respecto al coste actual; % sobre el coste de potencia actual. Importes "
            "sin impuesto eléctrico.\n\n"
            f"Inversión ({d.get('nombre', '')}): acceso {fmt(d['acceso'], 6)} €/kW y extensión "
            f"{fmt(d['extension'], 6)} €/kW sobre el aumento de la potencia máxima respecto a la "
            f"actual; enganche {fmt(d['enganche'], 2)} €. Si la ampliación es > 250 kW o en suelo "
            "no urbanizable, la extensión no se incluye (la ejecuta el titular)."
            + ("\n\nConceptos → " + " · ".join(inv) if inv else ""))


# ------------------------------------------------------------------- graficos y PDF
def figura(dibujar, *args, figsize=(9, 4), **kwargs) -> Figure:
    """Figure de matplotlib con uno de los graficos de potencia.graficos (los del PDF)."""
    fig = Figure(figsize=figsize, dpi=110, layout="constrained")
    dibujar(fig, *args, **kwargs)
    return fig


def nombre_pdf(datos, multipunto=False) -> str:
    if multipunto:
        nombre = f"Estudio potencia multipunto {datos.get('titular') or ''}".strip()
    else:
        nombre = (f"Estudio potencia {datos.get('titular') or ''} "
                  f"({datos.get('instalacion') or ''})").strip()
    return "".join(ch for ch in nombre if ch not in '\\/:*?"<>|') + ".pdf"


def _pdf(generar, *args, **kwargs) -> bytes:
    with tempfile.TemporaryDirectory() as carpeta:
        ruta = Path(carpeta) / "estudio.pdf"
        generar(ruta, *args, **kwargs)
        return ruta.read_bytes()


def pdf_individual(datos, s, precios_, anexos=True, seleccion=(0, None)) -> bytes:
    """datos: titular, cups, instalacion, direccion, fecha (tarifa y zona salen de s)."""
    d = dict(datos, tarifa=s.tarifa, zona=s.zona)
    return _pdf(informe_pdf.generar_pdf, d, s.estudio, s.escenarios, s.curva,
                precios_.get("normativa", []), incluir_anexos=anexos, seleccion_curva=seleccion)


def pdf_multipunto(datos, suministros, res, precios_, anexos=True, seleccion=(0, None)) -> bytes:
    sums = [{"datos": {"titular": datos.get("titular", ""), "fecha": datos.get("fecha", ""),
                       "cups": s.cups, "tarifa": s.tarifa, "instalacion": s.denominacion.strip(),
                       "direccion": s.direccion.strip(), "zona": s.zona},
             "estudio": s.estudio, "escenarios": s.escenarios, "curva": s.curva}
            for s in suministros]
    return _pdf(informe_pdf.generar_pdf_multipunto, datos, sums, res, filas_cups(suministros),
                precios_.get("normativa", []), incluir_anexos=anexos, seleccion_curva=seleccion)


# ------------------------------------------------------------------- guardar
def empaquetar(suministros) -> bytes:
    """Suministros con sus curvas ya leidas (sin el calculo, que se rehace al cargarlos),
    comprimidos para guardarlos en la base de datos de la plataforma (solo en local)."""
    ligeros = []
    for s in suministros:
        copia = Suministro(**{k: v for k, v in s.__dict__.items()
                              if k not in ("estudio", "optimo", "escenarios", "_clave")})
        ligeros.append(copia)
    return gzip.compress(pickle.dumps(ligeros, protocol=pickle.HIGHEST_PROTOCOL))


def desempaquetar(datos: bytes) -> list:
    """Lo contrario de empaquetar (solo para lo que guarda la propia plataforma)."""
    return pickle.loads(gzip.decompress(datos))


def resumen_guardar(tipo, datos, suministros) -> dict:
    """Columnas de la tabla de estudios guardados."""
    if tipo == "multipunto":
        res = resumen(suministros)
        opt = res.escenarios[1]
    else:
        opt = suministros[0].escenarios[1]
    return {"tipo": tipo, "titular": datos.get("titular", ""),
            "cups": [s.cups for s in suministros if s.cups],
            "tarifa": ", ".join(dict.fromkeys(s.tarifa for s in suministros)),
            "ahorro": round(float(opt.ahorro), 2),
            "inversion": round(float(opt.inversion.get("total", 0) or 0), 2)}
