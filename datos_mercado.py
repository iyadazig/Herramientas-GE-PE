# -*- coding: utf-8 -*-
"""
DATOS DE MERCADO PARA EL ANALISIS DE OFERTAS (solo lectura)
===========================================================
    OMIP   config.FICHERO_OMIP (omip_history.xlsx, hoja «Espana», lo actualiza OMIP_Scraper
           cada día): cotizaciones de cierre de los futuros de carga base de electricidad
           (instrumento FTB) mensuales (M), trimestrales (Q) y anuales (YR).
    OMIE   config.FICHERO_OMIE_MEDIO (PRECIO MEDIO OMIE.xls, hoja «desde 2019», mantenido a
           mano): precio medio mensual del mercado diario por año.

Precio de mercado previsto de cada mes del contrato:
    previsto = % OMIP × OMIP del mes + (1 − % OMIP) × media de OMIE de ese mes en los N
               últimos años con dato
con % OMIP y N de los datos regulados del año (datos_regulados, «mercado»). OMIP de cada mes:
el contrato mensual si cotiza; si no, el trimestral; si no, el anual (M > Q > YR). Sin
cotización, solo la media de OMIE (y un aviso).

Los ficheros están en la red y otros programas los escriben: se leen de una copia temporal
(con reintentos si están bloqueados) y se guardan en caché mientras no cambien (fecha de
modificación y tamaño). Cada análisis guarda lo que usó (congelar) para poder rehacerlo igual
aunque los ficheros cambien. Sin Streamlit.
"""

import datetime as dt
import re
import shutil
import tempfile
import time
import unicodedata
from pathlib import Path

import pandas as pd

import config

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MESES_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DIAS_AVISO_OMIP = 7          # OMIP con más días que esto desde la última sesión: aviso
FUENTE_OMIP = "OMIP – www.omip.pt"
FUENTE_OMIE = "OMIE – www.omie.es"

_cache = {}


class DatosMercado(Exception):
    """Fichero de mercado que no se puede leer o sin el formato esperado."""


# ------------------------------------------------------------------- lectura robusta
def _firma(ruta: Path):
    st = ruta.stat()
    return (str(ruta), st.st_mtime_ns, st.st_size)


def copia_local(ruta: Path, intentos: int = 3, espera: float = 1.0) -> Path:
    """Copia el fichero a una carpeta temporal (otro programa puede estar escribiéndolo). Si
    está bloqueado, lo reintenta. La copia se borra al leerla."""
    ruta = Path(ruta)
    ultimo = None
    for n in range(intentos):
        destino = Path(tempfile.mkdtemp(prefix="mercado_")) / ruta.name
        try:
            shutil.copy2(ruta, destino)
            return destino
        except FileNotFoundError:
            shutil.rmtree(destino.parent, ignore_errors=True)
            raise DatosMercado(f"No se encuentra el fichero {ruta}.")
        except OSError as e:
            shutil.rmtree(destino.parent, ignore_errors=True)
            ultimo = e
            if n < intentos - 1:
                time.sleep(espera)
    raise DatosMercado(f"No se puede leer {ruta} (puede que esté abierto o bloqueado): {ultimo}")


def _leer(ruta: Path, lector):
    """lector(copia_local) en caché por fichero, fecha de modificación y tamaño."""
    ruta = Path(ruta)
    try:
        firma = _firma(ruta)
    except OSError:
        raise DatosMercado(f"No se encuentra el fichero {ruta} (¿carpeta de red accesible?).")
    if firma not in _cache:
        copia = copia_local(ruta)
        try:
            _cache[firma] = lector(copia)
        finally:
            shutil.rmtree(copia.parent, ignore_errors=True)
    return _cache[firma]


def _sin_acentos(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


# ------------------------------------------------------------------- OMIP
_CONTRATO = re.compile(r"^FTB\s+(?:M\s+(?P<mes>[A-Za-z]{3})-(?P<am>\d{2})|"
                       r"Q(?P<q>[1-4])-(?P<aq>\d{2})|YR-(?P<ay>\d{2}))$")


def _leer_omip(ruta: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(ruta, sheet_name="Espana")
    except ValueError as e:
        raise DatosMercado(f"El fichero de OMIP no tiene la hoja «Espana»: {e}")
    faltan = {"fecha_sesion", "producto", "instrumento", "contrato",
              "settlement_price_eur_mwh"} - set(df.columns)
    if faltan:
        raise DatosMercado("Al fichero de OMIP le faltan las columnas " + ", ".join(sorted(faltan)))
    df = df[(df["producto"].astype(str).str.strip() == "Electricidad")
            & (df["instrumento"].astype(str).str.strip() == "FTB")].copy()
    # fecha_sesion mezcla fechas y textos («2026-05-05» o «05/05/2026»)
    texto = df["fecha_sesion"].astype(str).str.strip().str[:10]
    fechas = pd.to_datetime(texto, format="%Y-%m-%d", errors="coerce")
    fechas = fechas.fillna(pd.to_datetime(texto, format="%d/%m/%Y", errors="coerce"))
    df["sesion"] = fechas.dt.date
    filas = []
    for sesion, contrato, precio in zip(df["sesion"], df["contrato"],
                                        df["settlement_price_eur_mwh"]):
        m = _CONTRATO.match(str(contrato).strip())
        if not m or pd.isna(sesion) or pd.isna(precio):
            continue
        if m["mes"]:
            if m["mes"].title() not in MESES_EN:
                continue
            filas.append((sesion, "M", 2000 + int(m["am"]), MESES_EN.index(m["mes"].title()) + 1,
                          float(precio), contrato.strip()))
        elif m["q"]:
            filas.append((sesion, "Q", 2000 + int(m["aq"]), int(m["q"]), float(precio),
                          contrato.strip()))
        else:
            filas.append((sesion, "YR", 2000 + int(m["ay"]), 0, float(precio), contrato.strip()))
    if not filas:
        raise DatosMercado("El fichero de OMIP no tiene cotizaciones de electricidad (FTB M/Q/YR).")
    return pd.DataFrame(filas, columns=["sesion", "tipo", "anio", "periodo", "precio",
                                        "contrato"])


def omip(ruta: Path = None) -> pd.DataFrame:
    """Cotizaciones de electricidad: sesion (date), tipo (M/Q/YR), anio, periodo (mes,
    trimestre o 0), precio (€/MWh) y contrato."""
    return _leer(ruta or config.FICHERO_OMIP, _leer_omip)


def sesiones(cotizaciones: pd.DataFrame) -> list:
    return sorted(set(cotizaciones["sesion"]))


def sesion_omip(cotizaciones: pd.DataFrame, fecha: dt.date = None) -> dict:
    """Cotizaciones de una sesión: la última (o la última hasta `fecha`). {"fecha": date,
    "M": {(año, mes): precio}, "Q": {(año, trimestre): precio}, "YR": {año: precio}}."""
    disponibles = [s for s in sesiones(cotizaciones) if fecha is None or s <= fecha]
    if not disponibles:
        raise DatosMercado(f"No hay sesiones de OMIP hasta el {fecha:%d/%m/%Y}.")
    s = disponibles[-1]
    c = cotizaciones[cotizaciones["sesion"] == s]
    return {"fecha": s,
            "M": {(int(a), int(p)): float(v) for a, p, v in
                  c[c.tipo == "M"][["anio", "periodo", "precio"]].itertuples(index=False)},
            "Q": {(int(a), int(p)): float(v) for a, p, v in
                  c[c.tipo == "Q"][["anio", "periodo", "precio"]].itertuples(index=False)},
            "YR": {int(a): float(v) for a, v in
                   c[c.tipo == "YR"][["anio", "precio"]].itertuples(index=False)}}


def omip_mes(sesion: dict, anio: int, mes: int):
    """(precio, contrato usado) del mes con prioridad M > Q > YR, o (None, None)."""
    if (anio, mes) in sesion["M"]:
        return sesion["M"][(anio, mes)], f"M {MESES[mes - 1][:3]}-{anio % 100:02d}"
    q = (mes - 1) // 3 + 1
    if (anio, q) in sesion["Q"]:
        return sesion["Q"][(anio, q)], f"Q{q}-{anio % 100:02d}"
    if anio in sesion["YR"]:
        return sesion["YR"][anio], f"YR-{anio % 100:02d}"
    return None, None


# ------------------------------------------------------------------- OMIE (medias mensuales)
def _leer_omie_medio(ruta: Path) -> dict:
    try:
        hoja = pd.read_excel(ruta, sheet_name="desde 2019", header=None)
    except ValueError as e:
        raise DatosMercado(f"El fichero de precios medios de OMIE no tiene la hoja «desde 2019»: {e}")
    cabecera = None
    for f in range(hoja.shape[0]):
        for c in range(hoja.shape[1]):
            if _sin_acentos(hoja.iat[f, c]).startswith("mes/a"):
                cabecera = (f, c)
                break
        if cabecera:
            break
    if not cabecera:
        raise DatosMercado("No se encuentra la cabecera «mes/año» en el fichero de OMIE.")
    f0, c0 = cabecera
    anios = {}
    for c in range(c0 + 1, hoja.shape[1]):
        v = hoja.iat[f0, c]
        try:
            a = int(float(v))
        except (TypeError, ValueError):
            break
        if not 2000 <= a <= 2100:
            break
        anios[c] = a
    if not anios:
        raise DatosMercado("No se encuentran las columnas de los años en el fichero de OMIE.")
    precios = {}
    for f in range(f0 + 1, min(f0 + 20, hoja.shape[0])):
        nombre = _sin_acentos(hoja.iat[f, c0])
        if nombre not in MESES:
            continue
        mes = MESES.index(nombre) + 1
        for c, a in anios.items():
            v = pd.to_numeric(hoja.iat[f, c], errors="coerce")
            if pd.notna(v):
                precios[(a, mes)] = float(v)
    if not precios:
        raise DatosMercado("El fichero de OMIE no tiene precios mensuales.")
    return precios


def omie_medio(ruta: Path = None) -> dict:
    """{(año, mes): precio medio del mercado diario (€/MWh)}."""
    return _leer(ruta or config.FICHERO_OMIE_MEDIO, _leer_omie_medio)


def media_omie(precios: dict, mes: int, n_anios: int):
    """(media, años usados) del mes en los n últimos años con dato."""
    anios = sorted(a for (a, m) in precios if m == mes)[-n_anios:]
    if not anios:
        return None, []
    return sum(precios[(a, mes)] for a in anios) / len(anios), anios


# ------------------------------------------------------------------- precio previsto
def meses_horizonte(inicio: dt.date, n_meses: int) -> list:
    """[(año, mes)] de n meses desde el mes de inicio."""
    a, m = inicio.year, inicio.month
    salida = []
    for _ in range(n_meses):
        salida.append((a, m))
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return salida


def precio_previsto(meses: list, sesion: dict, precios_omie: dict, pct_omip: float,
                    n_anios: int) -> list:
    """Una fila por mes: anio, mes, omip, contrato, omie_media, anios_omie, previsto, aviso."""
    filas = []
    for a, m in meses:
        p_omip, contrato = omip_mes(sesion, a, m) if sesion else (None, None)
        media, usados = media_omie(precios_omie, m, n_anios)
        aviso = ""
        if p_omip is None and media is None:
            previsto, aviso = None, "Sin OMIP ni OMIE"
        elif p_omip is None:
            previsto, aviso = media, "Sin cotización de OMIP: solo la media de OMIE"
        elif media is None:
            previsto, aviso = p_omip, "Sin media de OMIE: solo OMIP"
        else:
            previsto = pct_omip * p_omip + (1 - pct_omip) * media
        filas.append({"anio": a, "mes": m, "omip": p_omip, "contrato": contrato,
                      "omie_media": media, "anios_omie": usados, "previsto": previsto,
                      "aviso": aviso})
    return filas


# ------------------------------------------------------------------- estado y copia congelada
def _modificado(ruta: Path):
    try:
        return dt.datetime.fromtimestamp(Path(ruta).stat().st_mtime)
    except OSError:
        return None


def estado(ruta_omip: Path = None, ruta_omie: Path = None, hoy: dt.date = None) -> dict:
    """Estado de los ficheros de mercado para enseñarlo en la app: fechas y avisos."""
    hoy = hoy or dt.date.today()
    ruta_omip = Path(ruta_omip or config.FICHERO_OMIP)
    ruta_omie = Path(ruta_omie or config.FICHERO_OMIE_MEDIO)
    salida = {"omip": {"fichero": str(ruta_omip), "modificado": _modificado(ruta_omip),
                       "ultima_sesion": None},
              "omie": {"fichero": str(ruta_omie), "modificado": _modificado(ruta_omie),
                       "ultimo_mes": None},
              "avisos": []}
    try:
        ultima = sesiones(omip(ruta_omip))[-1]
        salida["omip"]["ultima_sesion"] = ultima
        if (hoy - ultima).days > DIAS_AVISO_OMIP:
            salida["avisos"].append(f"La última sesión de OMIP es del {ultima:%d/%m/%Y} "
                                    f"(hace {(hoy - ultima).days} días): revisa que OMIP_Scraper "
                                    "se esté ejecutando.")
    except DatosMercado as e:
        salida["avisos"].append(str(e))
    try:
        ultimo = max(omie_medio(ruta_omie))
        salida["omie"]["ultimo_mes"] = ultimo
        # el mes pasado puede no estar aún a principios de mes; el anterior, sí
        if (hoy.year * 12 + hoy.month) - (ultimo[0] * 12 + ultimo[1]) > 2:
            salida["avisos"].append(f"El fichero de OMIE llega a {MESES[ultimo[1] - 1]} de "
                                    f"{ultimo[0]}: falta actualizarlo con los últimos meses.")
    except DatosMercado as e:
        salida["avisos"].append(str(e))
    return salida


def congelar(inicio: dt.date, n_meses: int, pct_omip: float, n_anios: int,
             fecha_sesion: dt.date = None, ruta_omip: Path = None, ruta_omie: Path = None) -> dict:
    """Lo que usa un análisis, para guardarlo con él: sesión de OMIP, ficheros y su fecha, y el
    precio previsto de cada mes del horizonte (JSON)."""
    sesion = sesion_omip(omip(ruta_omip), fecha_sesion)
    precios = omie_medio(ruta_omie)
    filas = precio_previsto(meses_horizonte(inicio, n_meses), sesion, precios, pct_omip, n_anios)
    return {"sesion_omip": sesion["fecha"].isoformat(),
            "fichero_omip": str(ruta_omip or config.FICHERO_OMIP),
            "fichero_omie": str(ruta_omie or config.FICHERO_OMIE_MEDIO),
            "omie_modificado": (_modificado(ruta_omie or config.FICHERO_OMIE_MEDIO) or
                                dt.datetime.min).isoformat(timespec="seconds"),
            "congelado": dt.datetime.now().isoformat(timespec="seconds"),
            "pct_omip": pct_omip, "anios_media_omie": n_anios,
            "fuentes": [FUENTE_OMIP, FUENTE_OMIE],
            "meses": filas}
