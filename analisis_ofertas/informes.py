# -*- coding: utf-8 -*-
"""
INFORMES DEL ANALISIS DE OFERTAS (Excel y PDF para el cliente)
==============================================================
    excel(e, res, ref, tecnico)   Excel de cálculo con fórmulas vivas (informe_excel): Resumen
                    con el formato de la hoja «Informe» de la plantilla, Info suministro,
                    Parámetros utilizados, contrato actual y una hoja por oferta
    pdf(e, res, ref, tecnico)     informe para el cliente con el formato de la plantilla
                    (informe_pdf): resumen, relación de suministros o ficha monopunto con
                    gráficas, y condiciones; A4 o A3 apaisado según las columnas
    relacion_suministros(e)   tabla de los CUPS (también para la app)

Importes sin decimales en €/año y €/kWh con 4 decimales; costes sin IVA. Sin Streamlit.
"""

import datetime as dt

import pandas as pd

from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import modelo as mo

# ------------------------------------------------------------------- formatos
def eur(v) -> str:
    """262194.3 → «262.194»."""
    return "—" if v is None else f"{v:,.0f}".replace(",", ".")


def ekwh(v) -> str:
    return "—" if v is None else f"{v:.4f}".replace(".", ",")


def pct(v) -> str:
    return "—" if v is None else f"{v * 100:+.2f} %".replace(".", ",").replace("-", "−")


def mes_corto(a: int, m: int) -> str:
    return f"{ap.MESES[m - 1][:3]}-{a % 100:02d}"


def titulo_oferta(f: dict) -> str:
    extra = ", ".join(x for x in (f["ssaa"], f["cierres"]) if x)
    ronda = f" · {f['ronda']}.ª ronda" if f["ronda"] and f["ronda"] > 1 else ""
    return f"{f['comercializadora'] or f['nombre']}\n{f['duracion']} meses ({extra}){ronda}"


def notas(e: mo.Estudio, res: dict) -> list:
    n = [f"Costes sin IVA. Incluyen potencia, energía, peajes y cargos, SSAA, FNEE, garantías de "
         f"origen, fees, impuesto municipal, impuesto eléctrico y alquileres; no incluyen excesos "
         f"de potencia ni energía reactiva.",
         "€/año = coste medio anual durante la duración de cada oferta, comparado con el coste "
         "estimado del contrato actual en esos mismos meses."]
    if any(o.indexada for o in e.ofertas + ([e.actual] if e.actual else [])):
        n.append(f"Precio de mercado de los indexados: {e.pct_omip * 100:.0f} % la cotización de "
                 f"OMIP{' del ' + dt.date.fromisoformat(e.sesion_omip).strftime('%d/%m/%Y') if e.sesion_omip else ''}"
                 f" y {100 - e.pct_omip * 100:.0f} % la media del precio de OMIE del mismo mes de "
                 f"los últimos años. Fuentes: {' y '.join(['OMIP – www.omip.pt', 'OMIE – www.omie.es'])}.")
        cliente = sorted({c for r in res["ofertas"] for c in r.apuntamiento_cliente})
        if cliente:
            n.append(f"Apuntamiento: en {len(cliente)} suministro(s) con curva horaria de un año "
                     "se usa su apuntamiento real; en el resto, el estimado de los últimos 12 "
                     "meses del mercado diario.")
        else:
            n.append("Apuntamiento estimado con el precio horario del mercado diario de los "
                     "últimos 12 meses publicados (ESIOS/OMIE).")
    return n


def relacion_suministros(e: mo.Estudio) -> pd.DataFrame:
    filas = []
    for s in e.suministros:
        if not s.activo:
            continue
        por_p = [sum(s.consumo_mes(m)[p] for m in range(1, 13)) for p in range(6)]
        filas.append({"CUPS": s.cups, "Alias": s.alias, "Dirección": s.direccion,
                      "Tarifa": s.tarifa, "Sistema": s.sistema,
                      **{f"Potencia P{p + 1} (kW)": s.potencias[p] for p in range(6)},
                      **{f"Consumo P{p + 1} (kWh)": round(por_p[p]) for p in range(6)},
                      "Consumo anual (kWh)": round(sum(por_p)),
                      "Reducción IEE": s.reduccion_iee})
    return pd.DataFrame(filas)


def coste_por_cups(res: dict) -> pd.DataFrame:
    """Coste medio anual (€/año) de cada CUPS en cada oferta."""
    filas = {}
    for r in res["ofertas"]:
        nombre = f"{r.oferta.comercializadora or r.oferta.nombre} {r.n_meses} m"
        for cups, d in r.por_cups.items():
            filas.setdefault(cups, {"CUPS": cups})[nombre] = d["total"] * 12 / r.n_meses
    return pd.DataFrame(list(filas.values()))


# ------------------------------------------------------------------- Excel
def excel(e: mo.Estudio, res: dict, ref: str = "", tecnico: dict = None) -> bytes:
    """Excel de cálculo con fórmulas vivas (ver informe_excel)."""
    from analisis_ofertas import informe_excel
    return informe_excel.excel(e, res, ref, tecnico)


# ------------------------------------------------------------------- PDF
def pdf(e: mo.Estudio, res: dict, ref: str = "", tecnico: dict = None) -> bytes:
    """Informe PDF con el formato de la plantilla (ver informe_pdf)."""
    from analisis_ofertas import informe_pdf
    return informe_pdf.pdf(e, res, ref, tecnico)
