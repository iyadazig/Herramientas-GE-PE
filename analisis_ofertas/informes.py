# -*- coding: utf-8 -*-
"""
INFORMES DEL ANALISIS DE OFERTAS (Excel y PDF para el cliente)
==============================================================
    excel(e, res)   hojas «Informe» (contrato actual, tablas de precio fijo e indexado con la
                    oferta mínima, cuadro de mercado y notas), «Relación de suministros»,
                    «Coste por CUPS» y las de «Parámetros utilizados»
    pdf(e, res)     el mismo informe en A4 apaisado
    relacion_suministros(e)   tabla de los CUPS (también para la app)

Importes sin decimales en €/año y €/kWh con 4 decimales; costes sin IVA. Sin Streamlit.
"""

import datetime as dt
import io
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import config
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo, parametros, resultados
from analisis_ofertas import modelo as mo

GRANATE, BLANCO, GRIS = "970000", "FFFFFF", "7F7F7F"
VERDE_CLARO, NARANJA, ROSA, BEIS = "E2EFDA", "F4B183", "F2DCDB", "FBF5EC"
LOGO = config.RECURSOS / "assets" / "logo_geype.png"
FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)


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
def _c(ws, ref, valor=None, negrita=False, color=None, fondo=None, formato=None, centro=False,
       tam=10, borde=True, ajustar=True):
    celda = ws[ref]
    celda.value = valor
    celda.font = Font(bold=negrita, color=color or "000000", size=tam)
    if fondo:
        celda.fill = PatternFill("solid", fgColor=fondo)
    if formato:
        celda.number_format = formato
    celda.alignment = Alignment(horizontal="center" if centro else None, vertical="center",
                                wrap_text=ajustar)
    if borde:
        celda.border = BORDE
    return celda


FILAS_TABLA = [("Duración (meses)", "duracion", "0"),
               ("Coste medio anual (€/año)", "media_anual", "#,##0"),
               ("Coste unitario (€/kWh)", "eur_kwh", "0.0000"),
               ("Contrato actual en el mismo periodo (€/año)", "actual_mismo_periodo", "#,##0"),
               ("Ahorro frente al contrato actual (€/año)", "ahorro", "#,##0"),
               ("Variación frente al contrato actual", "variacion", "+0.00%;-0.00%"),
               ("Ranking", "ranking", "0")]


def _tabla_excel(ws, fila, titulo, filas, mejor):
    _c(ws, f"B{fila}", titulo, negrita=True, color=BLANCO, fondo=GRANATE)
    for i, f in enumerate(filas):
        col = get_column_letter(3 + i)
        _c(ws, f"{col}{fila}", titulo_oferta(f), negrita=True, centro=True,
           fondo=NARANJA if f["duracion"] > 12 else ROSA)
        ws.column_dimensions[col].width = 20
    ws.row_dimensions[fila].height = 45
    for j, (nombre, clave, formato) in enumerate(FILAS_TABLA, start=1):
        _c(ws, f"B{fila + j}", nombre, negrita=clave == "media_anual")
        for i, f in enumerate(filas):
            col = get_column_letter(3 + i)
            v = f.get(clave)
            es_mejor = mejor and f["nombre"] == mejor["nombre"]
            color = None
            if clave in ("ahorro", "variacion") and v is not None:
                color = "00703C" if (v > 0 if clave == "ahorro" else v < 0) else "C00000"
            _c(ws, f"{col}{fila + j}", v, negrita=es_mejor and clave == "media_anual",
               formato=formato, centro=True, color=color,
               fondo=VERDE_CLARO if es_mejor else None)
    return fila + len(FILAS_TABLA) + 2


def excel(e: mo.Estudio, res: dict, ref: str = "") -> bytes:
    filas = resultados.comparativa(res)
    mejores = resultados.mejores(filas)
    wb = Workbook()
    ws = wb.active
    ws.title = "Informe"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 44
    if LOGO.exists():
        img = XLImage(str(LOGO))
        img.height, img.width = 60, int(60 * 700 / 352)
        ws.add_image(img, "B1")
    ws.merge_cells("B5:J5")
    _c(ws, "B5", "RESUMEN DEL ESTUDIO DE OFERTAS DE RENOVACIÓN", negrita=True, color=BLANCO,
       fondo=GRANATE, tam=12, centro=True)
    ws.row_dimensions[5].height = 26
    _c(ws, "B7", (e.cliente or "").upper(), negrita=True, tam=13, borde=False)
    _c(ws, "B8", " · ".join(x for x in (
        f"Ref. {ref}" if ref else "", f"Inicio de suministro: {e.inicio:%d/%m/%Y}" if e.inicio else "",
        f"{len([s for s in e.suministros if s.activo])} suministros",
        f"Fecha del estudio: {(e.fecha or dt.date.today()):%d/%m/%Y}") if x),
       color=GRIS, borde=False)
    fila = 10
    actual = res.get("actual")
    if actual:
        primer = actual.por_anio()[0]
        _c(ws, f"B{fila}", "Contrato actual", negrita=True, color=BLANCO, fondo=GRANATE)
        _c(ws, f"C{fila}", f"{actual.oferta.comercializadora} ({actual.oferta.modalidad})",
           negrita=True)
        ws.merge_cells(f"C{fila}:F{fila}")
        _c(ws, f"B{fila + 1}", "Coste estimado del primer año (€/año)")
        _c(ws, f"C{fila + 1}", primer["total"], formato="#,##0", negrita=True, centro=True)
        _c(ws, f"B{fila + 2}", "Coste unitario (€/kWh)")
        _c(ws, f"C{fila + 2}", primer["total"] / primer["consumo"] if primer["consumo"] else None,
           formato="0.0000", centro=True)
        fila += 4
    for t, titulo in (("Fijo", "PRECIO FIJO"), ("Indexado", "PRECIO INDEXADO")):
        de_tipo = [f for f in filas if f["tipo"] == t]
        if de_tipo:
            fila = _tabla_excel(ws, fila, titulo, de_tipo, mejores.get(t))
    for t, f in mejores.items():
        _c(ws, f"B{fila}", f"Oferta mínima ({t.lower()})", negrita=True, fondo=VERDE_CLARO)
        _c(ws, f"C{fila}", f"{f['comercializadora']} {f['duracion']} meses: "
                           f"{eur(f['media_anual'])} €/año ({pct(f['variacion'])} frente al actual)",
           fondo=VERDE_CLARO)
        ws.merge_cells(f"C{fila}:J{fila}")
        fila += 1
    fila += 1
    cuadro = resultados.cuadro_mercado(e, n_meses=max((r.n_meses for r in res["ofertas"]),
                                                      default=12))
    if cuadro:
        _c(ws, f"B{fila}", "MERCADO PREVISTO (€/MWh)", negrita=True, color=BLANCO, fondo=GRANATE)
        for i, (nombre, clave) in enumerate((("Precio previsto", "previsto"), ("OMIP", "omip"),
                                             ("Media OMIE", "omie"), ("SSAA", "ssaa")), start=1):
            _c(ws, f"B{fila + i}", nombre)
        for j, q in enumerate(cuadro[:12]):
            col = get_column_letter(3 + j)
            _c(ws, f"{col}{fila}", q["trimestre"], negrita=True, centro=True, fondo=BEIS)
            for i, clave in enumerate(("previsto", "omip", "omie", "ssaa"), start=1):
                _c(ws, f"{col}{fila + i}", q[clave], formato="#,##0.00", centro=True)
        fila += 6
    _c(ws, f"B{fila}", "Notas", negrita=True, borde=False)
    for n in notas(e, res):
        fila += 1
        ws.merge_cells(f"B{fila}:J{fila}")
        _c(ws, f"B{fila}", n, color=GRIS, borde=False, tam=9)
        ws.row_dimensions[fila].height = 28

    parametros.hoja(wb, "Relación de suministros", relacion_suministros(e),
                    "Relación de suministros")
    parametros.hoja(wb, "Coste por CUPS", coste_por_cups(res), "Coste medio anual (€/año)")
    for nombre, df in parametros.tablas_estudio(e).items():
        parametros.hoja(wb, nombre, df, "Parámetros utilizados")
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


# ------------------------------------------------------------------- PDF
def pdf(e: mo.Estudio, res: dict, ref: str = "") -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    gr, verde = colors.HexColor("#" + GRANATE), colors.HexColor("#" + VERDE_CLARO)
    est = {"n": ParagraphStyle("n", fontName="Helvetica", fontSize=8, leading=10),
           "b": ParagraphStyle("b", fontName="Helvetica-Bold", fontSize=8, leading=10),
           "t": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=16,
                               textColor=gr),
           "c": ParagraphStyle("c", fontName="Helvetica-Bold", fontSize=7, leading=8.5,
                               alignment=1, textColor=colors.white),
           "p": ParagraphStyle("p", fontName="Helvetica", fontSize=7, leading=9,
                               textColor=colors.HexColor("#595959"))}
    filas = resultados.comparativa(res)
    mejores = resultados.mejores(filas)
    hist = []
    if LOGO.exists():
        hist.append(Image(str(LOGO), width=40 * mm, height=20 * mm, hAlign="LEFT"))
    hist.append(Paragraph("RESUMEN DEL ESTUDIO DE OFERTAS DE RENOVACIÓN", est["t"]))
    hist.append(Paragraph(escape((e.cliente or "").upper()) + (f" · Ref. {escape(ref)}" if ref
                                                               else ""), est["b"]))
    hist.append(Spacer(1, 3 * mm))
    actual = res.get("actual")
    if actual:
        p = actual.por_anio()[0]
        hist.append(Paragraph(
            f"<b>Contrato actual:</b> {escape(actual.oferta.comercializadora or '')} "
            f"({escape(actual.oferta.modalidad)}) · coste estimado del primer año "
            f"<b>{eur(p['total'])} €/año</b> ({ekwh(p['total'] / p['consumo'] if p['consumo'] else None)} €/kWh)",
            est["n"]))
        hist.append(Spacer(1, 3 * mm))
    for t, titulo in (("Fijo", "Precio fijo"), ("Indexado", "Precio indexado")):
        de_tipo = [f for f in filas if f["tipo"] == t]
        if not de_tipo:
            continue
        cab = [Paragraph(titulo, est["c"])] + [
            Paragraph(escape(titulo_oferta(f)).replace("\n", "<br/>"), est["c"]) for f in de_tipo]
        datos = [cab]
        formatos = {"duracion": lambda v: str(v), "media_anual": eur, "eur_kwh": ekwh,
                    "actual_mismo_periodo": eur, "ahorro": eur, "variacion": pct,
                    "ranking": lambda v: "—" if v is None else str(v)}
        for nombre, clave, _ in FILAS_TABLA:
            datos.append([Paragraph(nombre, est["n"])] +
                         [formatos[clave](f.get(clave)) for f in de_tipo])
        ancho_col = min(38, (255 - 70) / max(1, len(de_tipo)))
        tabla = Table(datos, colWidths=[70 * mm] + [ancho_col * mm] * len(de_tipo), hAlign="LEFT")
        estilo = [("BACKGROUND", (0, 0), (-1, 0), gr), ("GRID", (0, 0), (-1, -1), 0.4,
                                                          colors.HexColor("#BFBFBF")),
                  ("FONTSIZE", (1, 1), (-1, -1), 8), ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                  ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                  ("TOPPADDING", (0, 1), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 1), (-1, -1), 1.5)]
        for i, f in enumerate(de_tipo, start=1):
            if f["duracion"] > 12:
                estilo.append(("BACKGROUND", (i, 0), (i, 0), colors.HexColor("#C55A11")))
            if mejores.get(t) and f["nombre"] == mejores[t]["nombre"]:
                estilo.append(("BACKGROUND", (i, 1), (i, -1), verde))
        tabla.setStyle(TableStyle(estilo))
        hist += [tabla, Spacer(1, 4 * mm)]
    for t, f in mejores.items():
        hist.append(Paragraph(f"<b>Oferta mínima ({t.lower()}):</b> {escape(f['comercializadora'] or '')} "
                              f"{f['duracion']} meses · {eur(f['media_anual'])} €/año "
                              f"({pct(f['variacion'])} frente al contrato actual)", est["n"]))
    cuadro = resultados.cuadro_mercado(e, n_meses=max((r.n_meses for r in res["ofertas"]),
                                                      default=12))[:12]
    if cuadro:
        hist.append(Spacer(1, 4 * mm))
        datos = [["€/MWh"] + [q["trimestre"] for q in cuadro]]
        for nombre, clave in (("Precio previsto", "previsto"), ("OMIP", "omip"),
                              ("Media OMIE", "omie"), ("SSAA", "ssaa")):
            datos.append([nombre] + [("—" if q[clave] is None else f"{q[clave]:.2f}".replace(".", ","))
                                     for q in cuadro])
        tabla = Table(datos, hAlign="LEFT", colWidths=[30 * mm] + [17 * mm] * len(cuadro))
        tabla.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), gr),
                                   ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                   ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                                   ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
                                   ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                                   ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
        hist.append(tabla)
    hist.append(Spacer(1, 3 * mm))
    for n in notas(e, res):
        hist.append(Paragraph(escape(n), est["p"]))
    # relación de suministros
    rel = relacion_suministros(e)
    if not rel.empty:
        hist += [PageBreak(), Paragraph("Relación de suministros", est["t"]), Spacer(1, 2 * mm)]
        cols = ["CUPS", "Alias", "Tarifa", "Sistema"] + [f"Potencia P{p} (kW)" for p in range(1, 7)] \
            + ["Consumo anual (kWh)"]
        datos = [[Paragraph(c.replace(" (kW)", "<br/>(kW)"), est["c"]) for c in cols]]
        for _, f in rel.iterrows():
            datos.append([f["CUPS"], Paragraph(escape(str(f["Alias"] or "")), est["n"]),
                          f["Tarifa"], f["Sistema"]]
                         + [("" if f["Tarifa"] == "2.0TD" and i >= 2 else
                             f"{f[c]:g}".replace(".", ",")) for i, c in enumerate(cols[4:10])]
                         + [eur(f["Consumo anual (kWh)"])])
        datos.append(["TOTAL", "", "", ""] + [""] * 6 + [eur(rel["Consumo anual (kWh)"].sum())])
        tabla = Table(datos, repeatRows=1, hAlign="LEFT",
                      colWidths=[45 * mm, 50 * mm, 17 * mm, 20 * mm] + [16 * mm] * 6 + [25 * mm])
        tabla.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), gr),
                                   ("FONTSIZE", (0, 1), (-1, -1), 7),
                                   ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                                   ("ALIGN", (2, 1), (-1, -1), "CENTER"),
                                   ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]))
        hist.append(tabla)

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#7F7F7F"))
        canvas.drawString(10 * mm, 7 * mm, f"GE&PE · Estudio de ofertas · {dt.date.today():%d/%m/%Y}")
        canvas.drawRightString(287 * mm, 7 * mm, f"Página {doc.page}")
        canvas.restoreState()
    salida = io.BytesIO()
    doc = SimpleDocTemplate(salida, pagesize=landscape(A4), leftMargin=10 * mm,
                            rightMargin=10 * mm, topMargin=8 * mm, bottomMargin=12 * mm,
                            title="Estudio de ofertas", author="GE&PE")
    doc.build(hist, onFirstPage=pie, onLaterPages=pie)
    return salida.getvalue()
