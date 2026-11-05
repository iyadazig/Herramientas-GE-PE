# -*- coding: utf-8 -*-
"""
PARAMETROS UTILIZADOS
=====================
Todos los valores con que se valoran las ofertas, con su fuente y su fecha:
    regulados      peajes y cargos de potencia y energía, pagos por capacidad, pérdidas y SRAD
                   por tarifa y año
    generales      IEE y sus mínimos, impuesto municipal, FNEE, OS y OM por año
    mercado        OMIP (sesión y contrato), media de OMIE (años), % OMIP y precio previsto por
                   mes; SSAA por mes
    apuntamientos  por sistema, calendario y mes, con el mes, año y fuente de cada valor

Sin estudio: los vigentes (datos regulados, ficheros de mercado y apuntamientos calculados). Con
estudio: los congelados con él. tablas() da DataFrames para la app y excel() el libro para
descargar. Sin Streamlit.
"""

import datetime as dt
import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import datos_mercado as dm
import datos_regulados as dr
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo
from analisis_ofertas import modelo as mo

GRANATE = "970000"
PER = mo.PERIODOS
MESES_CORTOS = [m[:3] for m in ap.MESES]


def _origen_reg(d: dict) -> str:
    if d.get("copia_de") not in (None, d.get("anio")):
        return f"Copia de {d['copia_de']} (sin publicar)"
    return {"guardado": "Administración", "semilla": "Programa"}.get(d.get("origen"), "Estudio")


def tabla_regulados(regs: dict) -> pd.DataFrame:
    """Una fila por año, tarifa y concepto con P1..P6. regs: {año: datos}."""
    filas = []
    for anio, d in sorted(regs.items()):
        for t in mo.TARIFAS:
            if t not in d.get("tarifas", {}):
                continue
            term = dr.terminos(d, t)
            v = d["tarifas"][t]
            for nombre, valores, unidad in (
                    ("Peajes de potencia", v.get("peajes_potencia"), "€/kW·año"),
                    ("Cargos de potencia", v.get("cargos_potencia"), "€/kW·año"),
                    ("Peajes + cargos de potencia", term["potencia"], "€/kW·año"),
                    ("Peajes de energía", v.get("peajes_energia"), "€/kWh"),
                    ("Cargos de energía", v.get("cargos_energia"), "€/kWh"),
                    ("Peajes + cargos de energía", term["energia"], "€/kWh"),
                    ("Pagos por capacidad", term["pagos_capacidad"], "€/kWh"),
                    ("Pérdidas", term["perdidas"], "tanto por uno"),
                    ("SRAD", term["srad"], "€/kWh")):
                if valores is None:
                    continue
                # periodos de la tarifa (la 2.0TD: 2 de potencia y 3 de energía)
                n = (2 if "potencia" in nombre else 3) if t == "2.0TD" else 6
                vals = mo.seis(list(valores)[:n])
                filas.append({"Año": anio, "Tarifa": t, "Concepto": nombre, "Unidad": unidad,
                              **dict(zip(PER, vals)),
                              "Origen": _origen_reg(dict(d, anio=anio))})
        for s, por_t in (d.get("perdidas_sistema") or {}).items():
            for t, valores in por_t.items():
                if valores:
                    filas.append({"Año": anio, "Tarifa": t, "Concepto": f"Pérdidas en {s}",
                                  "Unidad": "tanto por uno",
                                  **dict(zip(PER, mo.seis(valores))),
                                  "Origen": _origen_reg(dict(d, anio=anio))})
    return pd.DataFrame(filas)


def tabla_generales(regs: dict) -> pd.DataFrame:
    filas = []
    for anio, d in sorted(regs.items()):
        for clave, nombre, unidad in dr.GENERALES:
            filas.append({"Año": anio, "Parámetro": nombre, "Valor": d["generales"].get(clave),
                          "Unidad": unidad, "Origen": _origen_reg(dict(d, anio=anio))})
    return pd.DataFrame(filas)


def tabla_fuentes(regs: dict) -> pd.DataFrame:
    filas = []
    for anio, d in sorted(regs.items()):
        for f in d.get("fuentes") or []:
            filas.append({"Año": anio, "Fuente normativa": f,
                          "Revisado": d.get("revisado", "")})
    return pd.DataFrame(filas).drop_duplicates("Fuente normativa") if filas else pd.DataFrame()


def tabla_mercado(e: mo.Estudio, n_meses: int = None) -> pd.DataFrame:
    """Una fila por mes del horizonte del estudio."""
    ctx = calculo.Contexto(e)
    filas = []
    for k in range(n_meses or e.horizonte()):
        f = mo.sumar_meses(e.inicio, k)
        clave = f"{f.year}-{f.month:02d}"
        fila = e.mercado.get(clave) or {}
        previsto, omip = ctx.mercado(f.year, f.month)
        manual = clave in e.omip_manual
        filas.append({
            "Mes": f"{MESES_CORTOS[f.month - 1]}-{f.year % 100:02d}",
            "OMIP (€/MWh)": omip or None,
            "Contrato OMIP": "manual" if manual else (fila.get("contrato") or ""),
            "Media OMIE (€/MWh)": fila.get("omie_media"),
            "Años de la media OMIE": ", ".join(str(a) for a in fila.get("anios_omie") or []),
            "% OMIP": e.pct_omip, "Precio previsto (€/MWh)": previsto,
            "SSAA (€/MWh)": ctx.ssaa(f.year, f.month),
            "SSAA": "manual" if clave in e.ssaa_manual else "por mes natural"})
    return pd.DataFrame(filas)


def tabla_apuntamientos(ap_omie: dict, origen: dict = None, ap_ssaa: dict = None) -> pd.DataFrame:
    filas = []
    for clave in sorted(ap_omie, key=lambda c: (mo.SISTEMAS.index(c.split("|")[0])
                                                if c.split("|")[0] in mo.SISTEMAS else 9, c)):
        sistema, cal = clave.split("|")
        for m in range(1, 13):
            valores = mo.seis((ap_omie[clave] or {}).get(m))
            if not any(v is not None for v in valores):
                continue
            fila = {"Sistema": sistema, "Calendario": "2.0TD" if cal == "3P" else "6 periodos",
                    "Mes": ap.MESES[m - 1], **dict(zip(PER, valores)),
                    "Dato de": ((origen or {}).get(clave) or {}).get(m, "")}
            if ap_ssaa and ap_ssaa.get(clave) and ap_ssaa[clave].get(m) and \
                    ap_ssaa[clave][m] != ap_omie[clave].get(m):
                fila["SSAA distinto"] = "sí"
            filas.append(fila)
    return pd.DataFrame(filas)


def apuntamientos_vigentes(calculados: list, hoy: dt.date = None):
    """(ap_omie, origen) de los apuntamientos calculados (12 últimos meses)."""
    v = ap.vigentes(calculados, hoy)
    ap_omie = {clave: {m: d["factores"] for m, d in meses.items()} for clave, meses in v.items()}
    origen = {clave: {m: ap.etiqueta(d["anio"], m, d["fuente"]) for m, d in meses.items()}
              for clave, meses in v.items()}
    return ap_omie, origen


def tablas_estudio(e: mo.Estudio) -> dict:
    """Los parámetros congelados de un estudio guardado."""
    anios = sorted({mo.sumar_meses(e.inicio, k).year for k in range(e.horizonte())})
    ctx = calculo.Contexto(e)
    regs = {a: dict(ctx.regulados(a), anio=a) for a in anios}
    return {"Regulados por tarifa": tabla_regulados(regs),
            "Generales": tabla_generales(regs),
            "Mercado y SSAA": tabla_mercado(e),
            "Apuntamientos": tabla_apuntamientos(e.ap_omie, e.ap_origen, e.ap_ssaa),
            "Fuentes": tabla_fuentes(regs)}


def tablas_vigentes(anios: list = None, calculados: list = None, hoy: dt.date = None,
                    n_meses: int = 36) -> dict:
    """Los parámetros vigentes hoy (para la página «Parámetros utilizados»)."""
    hoy = hoy or dt.date.today()
    anios = anios or [hoy.year, hoy.year + 1, hoy.year + 2]
    regs = {a: dr.datos_anio(a) for a in anios}
    salida = {"Regulados por tarifa": tabla_regulados(regs), "Generales": tabla_generales(regs)}
    avisos = []
    try:
        inicio = mo.sumar_meses(dt.date(hoy.year, hoy.month, 1), 1)
        m = regs[anios[0]]["mercado"]
        cong = dm.congelar(inicio, n_meses, m["pct_omip"], m["anios_media_omie"])
        e = mo.Estudio(inicio=inicio, pct_omip=m["pct_omip"], sesion_omip=cong["sesion_omip"],
                       mercado={f"{f['anio']}-{f['mes']:02d}": f for f in cong["meses"]},
                       ssaa_mes=[None] * 12, regulados={a: regs[a] for a in anios})
        mercado = tabla_mercado(e, n_meses).drop(columns=["SSAA (€/MWh)", "SSAA"])
        salida["Mercado"] = mercado
        avisos.append(f"OMIP: sesión del {dt.date.fromisoformat(cong['sesion_omip']):%d/%m/%Y} "
                      f"({dm.FUENTE_OMIP}). OMIE: {dm.FUENTE_OMIE}.")
    except dm.DatosMercado as e:
        avisos.append(f"Datos de mercado no disponibles: {e}")
    if calculados is None:
        try:
            import almacen
            calculados = almacen.listar_apuntamientos()
        except Exception:                                   # noqa: BLE001
            calculados = []
    ap_omie, origen = apuntamientos_vigentes(calculados, hoy)
    salida["Apuntamientos"] = tabla_apuntamientos(ap_omie, origen)
    salida["Fuentes"] = tabla_fuentes(regs)
    salida["_avisos"] = avisos
    return salida


def hoja(wb, nombre: str, df: pd.DataFrame, titulo: str) -> None:
    """Hoja con título y la tabla (cabecera granate)."""
    ws = wb.create_sheet(nombre[:31])
    ws["A1"] = f"{titulo} — {nombre}"
    ws["A1"].font = Font(bold=True, size=12, color=GRANATE)
    ws["A2"] = f"Generado el {dt.datetime.now():%d/%m/%Y %H:%M} con Herramientas GE&PE"
    ws["A2"].font = Font(italic=True, size=9, color="7F7F7F")
    if df is None or df.empty:
        ws["A4"] = "Sin datos."
        return
    for c, col in enumerate(df.columns, start=1):
        celda = ws.cell(4, c, col)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=GRANATE)
        celda.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = max(10, min(45, len(str(col)) + 4))
    for r, fila in enumerate(df.itertuples(index=False), start=5):
        for c, v in enumerate(fila, start=1):
            if isinstance(v, float) and pd.isna(v):
                v = None
            celda = ws.cell(r, c, v)
            if isinstance(v, float):
                celda.number_format = "0.000000" if abs(v) < 1 else "#,##0.00"
    ws.freeze_panes = "A5"


def excel(tablas: dict, titulo: str = "Parámetros utilizados") -> bytes:
    """Libro con una hoja por tabla (las claves que empiezan por _ no se exportan)."""
    wb = Workbook()
    wb.remove(wb.active)
    for nombre, df in tablas.items():
        if not nombre.startswith("_"):
            hoja(wb, nombre, df, titulo)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()
