# -*- coding: utf-8 -*-
"""
APUNTAMIENTOS («PROFILE») DEL PRECIO DEL MERCADO DIARIO
=======================================================
Apuntamiento de un periodo en un mes = precio medio de las horas del periodo ÷ precio medio de
todas las horas del mes (como la fila «Profile» de los Excel mensuales de OMIE). Se calcula para
cada sistema eléctrico (Península, Baleares, Canarias, Ceuta, Melilla) y calendario:

    6P    peajes de 6 periodos (3.0TD, 6.xTD): potencia/calendario.asignar_periodos
    3P    2.0TD: P1 10-14 h y 18-22 h; P2 8-10 h, 14-18 h y 22-24 h; P3 0-8 h, fines de semana
          y festivos nacionales (Ceuta y Melilla: una hora más tarde)

Precio hora a hora del mercado diario español:
    ESIOS   columna «Mercado diario €/MWh» de las hojas «PFMHORAS_COM mmm-aa» de
            AAAA_Historico_componentes_ESIOS.xlsx (ssaa_datos_esios.pfmhoras)
    OMIE    respaldo para los meses que no están completos en ESIOS: matriz día × hora de los
            Excel mensuales de config.CARPETA_OMIE_HORARIO (Omie/AAAA/«9 septiembre'26.xls»; el
            año sale del nombre del fichero, porque hay ficheros en la carpeta de otro año)

Las horas van numeradas como en ESIOS (1..23/24/25 por orden dentro del día) y se pasan a la hora
de inicio peninsular; en Canarias se aplican las franjas sobre esa hora, igual que los Excel de
OMIE. Solo se usan meses completos.

Se guardan en almacen (apuntamientos_calculados) y el análisis usa, para cada mes del año, el
de los 12 últimos meses calculados, con su mes, año y fuente. Con la curva horaria de un año
completo de un CUPS se calcula su apuntamiento propio (precio ponderado por su consumo).
Sin Streamlit.
"""

import calendar
import datetime as dt
import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

import config
from potencia import calendario

SISTEMAS = ["Península", "Baleares", "Canarias", "Ceuta", "Melilla"]
CALENDARIOS = ["6P", "3P"]
ZONA_CALENDARIO = {"Baleares": "Illes Balears"}       # nombre en potencia/calendario
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
ESIOS, OMIE = "ESIOS", "OMIE"
DIAS_CURVA_COMPLETA = 360


class SinPrecios(Exception):
    """No hay precios horarios completos del mes."""


def calendario_de(tarifa: str) -> str:
    return "3P" if tarifa == "2.0TD" else "6P"


def n_periodos(cal: str) -> int:
    return 3 if cal == "3P" else 6


# ------------------------------------------------------------------- horas y periodos
def hora_inicio(fecha: dt.date, hora: int) -> int:
    """Hora de inicio (0..23) de la hora ESIOS `hora` (1..25) del día."""
    marzo, octubre = calendario.dias_cambio_hora(fecha.year)
    if fecha == marzo and hora >= 3:
        return hora                       # no existe la hora de las 2:00
    if fecha == octubre and hora >= 4:
        return hora - 2                   # la de las 2:00 se repite (horas 3 y 4)
    return hora - 1


def periodos(claves: list, sistema: str, cal: str) -> np.ndarray:
    """Periodo de cada (fecha, hora ESIOS) en el calendario 6P o 3P del sistema."""
    fechas = [f for f, _ in claves]
    horas = np.array([hora_inicio(f, h) for f, h in claves])
    if cal == "6P":
        inicio = pd.to_datetime(fechas) + pd.to_timedelta(horas, unit="h")
        return calendario.asignar_periodos(inicio, ZONA_CALENDARIO.get(sistema, sistema))
    desfase = 1 if sistema in ("Ceuta", "Melilla") else 0
    h = horas - desfase
    fest = calendario.festivos(sorted({f.year for f in fechas}))
    fin_semana = np.array([f.weekday() >= 5 or f in fest for f in fechas])
    p = np.full(len(claves), 3)
    p[np.isin(h, [8, 9, 14, 15, 16, 17, 22, 23])] = 2
    p[np.isin(h, [10, 11, 12, 13, 18, 19, 20, 21])] = 1
    if desfase:                            # 23-24 h (P2) en Ceuta y Melilla
        p[horas == 23] = 2
        p[horas == 8] = 2
    p[horas < 8] = 3
    p[fin_semana] = 3
    return p


def profile(precios: dict, sistema: str, cal: str, pesos: dict = None) -> list:
    """Apuntamiento de cada periodo (lista de 6; None si el periodo no tiene horas).
    Sin pesos: precio medio de las horas del periodo / precio medio del mes. Con pesos (consumo
    de un CUPS por hora): precio medio ponderado por el consumo del periodo / precio medio del
    mes."""
    claves = sorted(precios)
    p = periodos(claves, sistema, cal)
    valores = np.array([precios[k] for k in claves], float)
    base = valores.mean()
    salida = []
    for per in range(1, 7):
        sel = p == per
        if per > n_periodos(cal) or not sel.any() or base == 0:
            salida.append(None)
            continue
        if pesos is None:
            salida.append(float(valores[sel].mean() / base))
        else:
            w = np.array([pesos.get(k, 0.0) for k, s in zip(claves, sel) if s], float)
            salida.append(float((valores[sel] * w).sum() / w.sum() / base) if w.sum() > 0
                          else None)
    return salida


# ------------------------------------------------------------------- precios horarios
def _horas_mes(anio: int, mes: int) -> int:
    dias = calendar.monthrange(anio, mes)[1]
    marzo, octubre = calendario.dias_cambio_hora(anio)
    return dias * 24 - (marzo.month == mes) + (octubre.month == mes)


def completo(precios: dict, anio: int, mes: int) -> bool:
    return len({k for k in precios if k[0].year == anio and k[0].month == mes}) == \
        _horas_mes(anio, mes)


def precios_esios(anio: int, mes: int) -> dict:
    """{(fecha, hora): €/MWh} del mercado diario del mes (ESIOS PFMHORAS)."""
    import ssaa_datos_esios as esios
    ini = dt.date(anio, mes, 1)
    fin = dt.date(anio, mes, calendar.monthrange(anio, mes)[1])
    try:
        datos, _ = esios.pfmhoras(ini, fin, columnas=("Mercado diario",))
    except Exception as e:                  # sin fichero del año u hoja sin esa columna
        raise SinPrecios(f"ESIOS {MESES[mes - 1]} {anio}: {e}")
    return datos


def _sin_acentos(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def mes_anio_fichero(nombre: str, carpeta_anio: int = None):
    """(mes, año) de un Excel mensual de OMIE («9 septiembre'26.xls», «1 Enero'26.XLS»)."""
    n = _sin_acentos(Path(nombre).stem)
    mes = next((i + 1 for i, m in enumerate(MESES) if m in n or m[:3] + "'" in n), None)
    a = re.search(r"['_\-\s](\d{2,4})$", n)
    anio = (int(a.group(1)) + (2000 if len(a.group(1)) == 2 else 0)) if a else carpeta_anio
    return mes, anio


def fichero_omie(anio: int, mes: int, carpeta: Path = None):
    """Excel mensual de OMIE del mes (se busca en todas las carpetas de año) o None."""
    carpeta = Path(carpeta or config.CARPETA_OMIE_HORARIO)
    try:
        subcarpetas = [c for c in carpeta.iterdir() if c.is_dir() and c.name.isdigit()]
    except OSError:
        return None
    # primero la carpeta del año
    for c in sorted(subcarpetas, key=lambda c: c.name != str(anio)):
        for f in sorted(c.iterdir()):
            if f.suffix.lower() in (".xls", ".xlsx") and \
                    mes_anio_fichero(f.name, int(c.name)) == (mes, anio):
                return f
    return None


def _filas_excel(ruta: Path) -> list:
    """Filas de la primera hoja que tenga la matriz día × hora."""
    if ruta.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(ruta, read_only=True, data_only=True)
        try:
            hojas = [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
        finally:
            wb.close()
    else:
        import xlrd
        libro = xlrd.open_workbook(ruta, on_demand=True)
        hojas = []
        for n in libro.sheet_names():
            sh = libro.sheet_by_name(n)
            hojas.append([sh.row_values(r) for r in range(sh.nrows)])
    for filas in hojas:
        if any(str(f[0]).strip().lower() in ("día", "dia") for f in filas if f):
            return filas
    raise SinPrecios(f"{ruta.name}: no tiene la tabla de precios por día y hora.")


def precios_omie_fichero(ruta: Path, anio: int, mes: int) -> dict:
    """{(fecha, hora): €/MWh} de la matriz día × hora de un Excel mensual de OMIE. Las columnas
    son las horas 1..24 (25 el día de 25 h); los días, «L 01», «M 02»…"""
    filas = _filas_excel(Path(ruta))
    cab = next(i for i, f in enumerate(filas) if f and str(f[0]).strip().lower() in ("día", "dia"))
    columnas = {}
    for c, v in enumerate(filas[cab][1:], start=1):
        try:
            columnas[c] = int(float(str(v).strip()))
        except ValueError:
            # la hora 25 (día de 25 h de octubre) viene sin título tras la 24
            if str(v or "").strip() == "" and columnas and max(columnas.values()) == 24:
                columnas[c] = 25
                continue
            break
    salida = {}
    for f in filas[cab + 1:]:
        m = re.search(r"(\d{1,2})\s*$", str(f[0]).strip()) if f else None
        if not m:
            if f and str(f[0]).strip().lower().startswith("total"):
                break
            continue
        try:
            fecha = dt.date(anio, mes, int(m.group(1)))
        except ValueError:
            continue
        for c, hora in columnas.items():
            if c < len(f) and isinstance(f[c], (int, float)) and f[c] != "":
                salida[(fecha, hora)] = float(f[c])
    return salida


def precios_mes(anio: int, mes: int, carpeta_omie: Path = None):
    """(precios horarios completos del mes, fuente): ESIOS y, si no está completo, OMIE."""
    motivos = []
    try:
        p = precios_esios(anio, mes)
        if completo(p, anio, mes):
            return p, ESIOS
        motivos.append(f"ESIOS incompleto ({len(p)} h)")
    except SinPrecios as e:
        motivos.append(str(e))
    ruta = fichero_omie(anio, mes, carpeta_omie)
    if ruta:
        p = precios_omie_fichero(ruta, anio, mes)
        if completo(p, anio, mes):
            return p, OMIE
        motivos.append(f"{ruta.name} incompleto ({len(p)} h)")
    else:
        motivos.append("sin Excel mensual de OMIE")
    raise SinPrecios(f"{MESES[mes - 1]} {anio}: " + "; ".join(motivos))


# ------------------------------------------------------------------- cálculo mensual
def calcular_mes(anio: int, mes: int, carpeta_omie: Path = None) -> dict:
    """Apuntamientos del mes para todos los sistemas y calendarios:
    {"anio", "mes", "fuente", "factores": {"Península|6P": [6 valores], ...}}."""
    precios, fuente = precios_mes(anio, mes, carpeta_omie)
    return {"anio": anio, "mes": mes, "fuente": fuente,
            "factores": {f"{s}|{c}": profile(precios, s, c)
                         for s in SISTEMAS for c in CALENDARIOS}}


def meses_a_calcular(hoy: dt.date, n: int = 12) -> list:
    """Los n meses completos anteriores al mes de hoy, del más reciente al más antiguo."""
    a, m = hoy.year, hoy.month
    salida = []
    for _ in range(n):
        a, m = (a - 1, 12) if m == 1 else (a, m - 1)
        salida.append((a, m))
    return salida


def actualizar(hoy: dt.date = None, guardados: dict = None, carpeta_omie: Path = None,
               guardar=None, meses_atras: int = 15) -> list:
    """Calcula los meses que faltan (o que solo tenían el respaldo de OMIE y ya están en ESIOS)
    de los últimos `meses_atras`. guardados: {(año, mes): {"fuente": ...}} de lo ya calculado.
    guardar(calculo) lo guarda (almacen). Devuelve los avisos de los meses sin precios."""
    hoy = hoy or dt.date.today()
    if guardados is None:
        import almacen
        guardados = {(r["anio"], r["mes"]): r for r in almacen.listar_apuntamientos()}
    if guardar is None:
        import almacen
        guardar = almacen.guardar_apuntamientos
    avisos = []
    for a, m in meses_a_calcular(hoy, meses_atras):
        previo = guardados.get((a, m))
        if previo and previo.get("fuente") == ESIOS:
            continue
        try:
            c = calcular_mes(a, m, carpeta_omie)
        except SinPrecios as e:
            avisos.append(str(e))
            continue
        if previo and previo.get("fuente") == c["fuente"]:
            continue
        guardar(c)
    return avisos


def vigentes(calculados: list, hoy: dt.date = None) -> dict:
    """Apuntamiento de cada mes del año (1..12): el más reciente de los 12 últimos meses
    completos ya calculados (si falta alguno, el del mismo mes del año anterior).
    {"Península|6P": {mes: {"factores": [...], "anio": 2025, "fuente": "ESIOS"}}}."""
    hoy = hoy or dt.date.today()
    por_mes = {}
    for c in sorted(calculados, key=lambda c: (c["anio"], c["mes"])):
        if (c["anio"], c["mes"]) < (hoy.year, hoy.month):
            por_mes[c["mes"]] = c                    # se queda el más reciente
    salida = {}
    for clave in (f"{s}|{c}" for s in SISTEMAS for c in CALENDARIOS):
        salida[clave] = {m: {"factores": c["factores"][clave], "anio": c["anio"],
                             "fuente": c["fuente"]}
                         for m, c in por_mes.items() if clave in c["factores"]}
    return salida


def etiqueta(anio: int, mes: int, fuente: str) -> str:
    """«sep-26 (ESIOS)»."""
    return f"{MESES[mes - 1][:3]}-{anio % 100:02d} ({fuente})"


# ------------------------------------------------------------------- apuntamiento del cliente
def cobertura(curva: dict) -> dict:
    """¿La curva {(fecha, hora): kWh} tiene un año completo? Días con dato, meses presentes,
    primer y último día y los 12 meses consecutivos más recientes que se pueden usar."""
    dias = sorted({k[0] for k, v in curva.items() if v is not None})
    if not dias:
        return {"valida": False, "dias": 0, "motivo": "La curva está vacía."}
    fin = dias[-1]
    # 12 meses completos hasta el último mes con datos casi completo
    ultimo = (fin.year, fin.month)
    if fin.day < calendar.monthrange(*ultimo)[1] - 1:
        ultimo = (fin.year - (fin.month == 1), 12 if fin.month == 1 else fin.month - 1)
    meses = []
    a, m = ultimo
    for _ in range(12):
        meses.append((a, m))
        a, m = (a - 1, 12) if m == 1 else (a, m - 1)
    meses.reverse()
    ventana = {d for d in dias if (d.year, d.month) in meses}
    presentes = {(d.year, d.month) for d in ventana}
    valida = len(ventana) >= DIAS_CURVA_COMPLETA and presentes == set(meses)
    motivo = "" if valida else (
        f"Faltan meses: {', '.join(f'{MESES[mm - 1][:3]}-{aa % 100:02d}' for aa, mm in meses if (aa, mm) not in presentes)}."
        if presentes != set(meses) else
        f"Solo {len(ventana)} días con datos en los 12 meses (hacen falta {DIAS_CURVA_COMPLETA}).")
    return {"valida": valida, "dias": len(ventana), "desde": dt.date(*meses[0], 1),
            "hasta": dt.date(*meses[-1], calendar.monthrange(*meses[-1])[1]),
            "meses": meses, "motivo": motivo}


def apuntamiento_cliente(curva: dict, sistema: str, tarifa: str, precios_de=None) -> dict:
    """Apuntamiento propio de un CUPS con su curva de un año completo: para cada mes del año,
    precio ponderado por su consumo de cada periodo / precio medio del mes (con los precios del
    mismo mes de la curva). {"valida", "motivo", "cobertura", "factores": {mes: [6]},
    "fuentes": {mes: "sep-26 (ESIOS)"}}. precios_de(año, mes) -> (precios, fuente)."""
    precios_de = precios_de or precios_mes
    cob = cobertura(curva)
    if not cob["valida"]:
        return {"valida": False, "motivo": cob["motivo"], "cobertura": cob}
    cal = calendario_de(tarifa)
    factores, fuentes = {}, {}
    for a, m in cob["meses"]:
        try:
            precios, fuente = precios_de(a, m)
        except SinPrecios as e:
            return {"valida": False, "motivo": f"Sin precios horarios del mercado: {e}",
                    "cobertura": cob}
        pesos = {k: v for k, v in curva.items() if k[0].year == a and k[0].month == m}
        factores[m] = profile(precios, sistema, cal, pesos)
        fuentes[m] = etiqueta(a, m, fuente)
    return {"valida": True, "motivo": "", "cobertura": cob, "factores": factores,
            "fuentes": fuentes}


def pendientes(calculados: list, hoy: dt.date = None) -> bool:
    """¿Falta el mes pasado (o solo está con el respaldo de OMIE)? Para recalcular al abrir el
    análisis sin esperar a la tarea programada."""
    hoy = hoy or dt.date.today()
    a, m = meses_a_calcular(hoy, 1)[0]
    return not any(c["anio"] == a and c["mes"] == m and c["fuente"] == ESIOS for c in calculados)


def main(argv=None):
    """Tarea programada del servidor (una vez al día; el cálculo solo se hace si hay un mes
    nuevo):  python -m analisis_ofertas.apuntamientos"""
    import argparse
    import logging
    parser = argparse.ArgumentParser(description=main.__doc__)
    parser.parse_args(argv)
    config.CARPETA_DATOS.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(config.CARPETA_DATOS / "apuntamientos.log",
                                      encoding="utf-8"), logging.StreamHandler()])
    try:
        avisos = actualizar()
    except Exception:                                     # noqa: BLE001
        logging.exception("Error al calcular los apuntamientos")
        return 1
    for a in avisos:
        logging.info("Sin precios: %s", a)
    logging.info("Apuntamientos al día.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
