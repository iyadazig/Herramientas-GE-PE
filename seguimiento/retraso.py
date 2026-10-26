# -*- coding: utf-8 -*-
"""
Retraso en el envio de informes de seguimiento
==============================================
La macro GenerarDashboardRetraso del Excel «Retraso_Seguimiento_…xlsm», en Python:

- Solo los suministros con Tipo contratación «Sgto. y Control» y una cadencia reconocible.
- Cada año se parte en bloques naturales del tamaño de la cadencia: mensual 1 mes,
  trimestral 3 (ene-mar, abr-jun…), semestral 6 y anual 12.
- En cada mes del bloque: vacio, «Sin fra», «Cadencia» y «Temporal» no esperan informe;
  cualquier otro estado lo espera y solo «Inf enviado» cuenta como enviado.
- Un bloque sin ningun informe esperado no se evalua.
- Fecha de control: dia 1 del segundo mes despues del ultimo mes del bloque (M+2). Si aun no
  ha llegado, el bloque no se evalua.
- ENVIADO si todos los esperados estan enviados; si no, RETRASO.
"""

import datetime as dt

import pandas as pd

from seguimiento import columnas as C

ENVIADO, RETRASO = "ENVIADO", "RETRASO"
NO_ESPERAN = {"", "sin fra", "cadencia", "temporal"}

COLUMNAS_AUDITORIA = [
    "Técnico", "Auxiliar", "Cliente/Grupo", "Denominación suministro", "CUPS", "Cadencia",
    "Bloque", "Mes inicio", "Mes fin", "Fecha control (M+2)", "Tam. bloque", "Inf. enviados",
    "Inf. esperados", "Resultado", "Valores del bloque", "Observaciones", "_id"]


def tam_bloque(cadencia):
    """Meses por bloque segun la cadencia (0 si no se reconoce), como la macro: busca el texto
    dentro de la cadencia y gana el ultimo que encaja."""
    cad = (cadencia or "").strip().lower()
    tam = 0
    for texto, meses in (("mensual", 1), ("trimestral", 3), ("semestral", 6), ("anual", 12)):
        if texto in cad:
            tam = meses
    return tam


def fecha_control(anio, mes_fin):
    """Dia 1 del mes M+2 tras el ultimo mes del bloque."""
    total = anio * 12 + (mes_fin - 1) + 2
    return dt.date(total // 12, total % 12 + 1, 1)


def auditoria(suministros, estados, anio, hoy=None):
    """Una fila (dict) por bloque evaluado. suministros: dicts id y datos; estados:
    {(id, anio, mes): {estado, ...}} (almacen.listar_sgc_estados)."""
    hoy = hoy or dt.date.today()
    filas = []
    for s in suministros:
        d = s["datos"]
        if (d.get("Tipo contratación") or "").strip() != C.SGTO_Y_CONTROL:
            continue
        tam = tam_bloque(d.get("Cadencia informe"))
        if not tam:
            continue
        for bloque in range(1, -(-12 // tam) + 1):
            ini, fin = (bloque - 1) * tam + 1, min(bloque * tam, 12)
            enviados = esperados = 0
            valores = []
            for mes in range(ini, fin + 1):
                v = ((estados.get((s["id"], anio, mes)) or {}).get("estado") or "").strip().lower()
                valores.append("[%s]" % (v or "vacio"))
                if v not in NO_ESPERAN:
                    esperados += 1
                    enviados += v == "inf enviado"
            if not esperados:
                continue
            control = fecha_control(anio, fin)
            if hoy < control:
                continue
            filas.append({
                "Técnico": (d.get("Técnico") or "").strip(),
                "Auxiliar": (d.get("Auxiliar") or "").strip(),
                "Cliente/Grupo": d.get("Cliente/Grupo"),
                "Denominación suministro": d.get("Denominación suministro"),
                "CUPS": (d.get("CUPS") or "").strip(), "Cadencia": d.get("Cadencia informe"),
                "Bloque": bloque, "Mes inicio": dt.date(anio, ini, 1),
                "Mes fin": dt.date(anio, fin, 1), "Fecha control (M+2)": control,
                "Tam. bloque": tam, "Inf. enviados": enviados, "Inf. esperados": esperados,
                "Resultado": ENVIADO if enviados == esperados else RETRASO,
                "Valores del bloque": " ".join(valores),
                "Observaciones": d.get("Observaciones"), "_id": s["id"]})
    filas.sort(key=lambda f: (str(f["Cliente/Grupo"] or "").lower(), f["Resultado"] != RETRASO))
    return filas


def auditoria_df(filas):
    return pd.DataFrame(filas, columns=COLUMNAS_AUDITORIA)


def resumen(df, por):
    """Retrasos, enviados y % de retraso por «Técnico» o «Auxiliar» (los vacios no cuentan,
    como en la macro), de mas a menos retrasos."""
    datos = df[df[por].fillna("") != ""]
    tabla = datos.groupby(por)["Resultado"].agg(
        Retrasos=lambda r: int((r == RETRASO).sum()),
        Enviados=lambda r: int((r == ENVIADO).sum())).reset_index()
    total = tabla["Retrasos"] + tabla["Enviados"]
    tabla["% Retraso"] = (tabla["Retrasos"] / total.where(total > 0)).fillna(0)
    return tabla.sort_values(["Retrasos", por], ascending=[False, True]).reset_index(drop=True)


def indicadores(df):
    """(retrasos, enviados, % cumplimiento o None)."""
    retrasos = int((df["Resultado"] == RETRASO).sum())
    enviados = int((df["Resultado"] == ENVIADO).sum())
    total = retrasos + enviados
    return retrasos, enviados, (enviados / total if total else None)
