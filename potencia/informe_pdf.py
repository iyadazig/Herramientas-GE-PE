"""Informe PDF del estudio de optimización de potencia (formato GE&PE)."""

import io
import textwrap
from dataclasses import replace
from xml.sax.saxutils import escape

import pandas as pd

from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (Image, KeepInFrame, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

from .calculo import MESES_NOMBRE, fmt, fmt_pot, ruta_recurso, texto_conceptos
from .graficos import (AZUL_OSCURO, COLORES_SUAVES, color, dibujar_costes, dibujar_curva_selector,
                       dibujar_maximos)

GRIS = colors.HexColor("#D9D9D9")
AZUL = colors.HexColor(AZUL_OSCURO)
FILA_CLARA = colors.HexColor("#DCE6F1")
ROJO = colors.HexColor("#C00000")

_E = {
    "normal": ParagraphStyle("n", fontName="Helvetica", fontSize=7.5, leading=9.5),
    "negrita": ParagraphStyle("b", fontName="Helvetica-Bold", fontSize=7.5, leading=9.5),
    "titulo": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=11, leading=13),
    "celda": ParagraphStyle("c", fontName="Helvetica-Bold", fontSize=7, leading=8.2, alignment=TA_CENTER,
                            textColor=colors.white),
    "celda_n": ParagraphStyle("cn", fontName="Helvetica-Bold", fontSize=7, leading=8.2, alignment=TA_CENTER),
    "pie": ParagraphStyle("p", fontName="Helvetica", fontSize=6.5, leading=8),
    "subrayado": ParagraphStyle("s", fontName="Helvetica", fontSize=7.5, leading=9.5, underlineWidth=0.5),
}

ANCHO_PAGINA = 277  # mm útiles en A4 apaisado con márgenes de 10 mm
SEPARACION = 4      # mm entre la tabla mensual y el gráfico


def _nombre(nombre):
    """Nombre de un escenario tal como aparece en el informe."""
    return "Propuesta 1 (óptima)" if nombre in ("Propuesta 1", "Óptima") else nombre


def _con_nombres(escenarios):
    return [replace(e, nombre=_nombre(e.nombre)) for e in escenarios]


def _marco(est, filas_total=1):
    """Contorno común a todas las tablas: rejilla blanca entre celdas y línea azul sobre los totales (sin marco)."""
    est.append(("INNERGRID", (0, 0), (-1, -1), 0.6, colors.white))
    if filas_total:
        est.append(("LINEABOVE", (0, -filas_total), (-1, -filas_total), 1, AZUL))
    return est


def _lista_y(elementos):
    """['3', '4', '5'] -> '3, 4 y 5'."""
    return elementos[0] if len(elementos) == 1 else ", ".join(elementos[:-1]) + " y " + elementos[-1]


def _ancho_mensual(escenarios):
    """Ancho (mm) de la tabla mensual según el número de columnas (costes + ahorros)."""
    columnas = 2 * len(escenarios) - 1
    # el gráfico conserva al menos 110 mm aunque haya muchas propuestas
    return min(ANCHO_PAGINA - SEPARACION - 110, max(120, 16 + 21 * columnas))


def _tabla_y_grafico(tabla, ancho_tabla, estudio, escenarios):
    """Tabla mensual a la izquierda y gráfico de costes a la derecha, ocupando todo el ancho."""
    ancho_graf = ANCHO_PAGINA - ancho_tabla - SEPARACION
    # leyenda en varias filas si el gráfico es estrecho
    ncol = min(len(escenarios), 4 if ancho_graf >= 140 else 3)
    return Table([[tabla, _figura_png(dibujar_costes, ancho_graf, 68, estudio, escenarios, ncol=ncol)]],
                 colWidths=[(ancho_tabla + SEPARACION) * mm, ancho_graf * mm], hAlign="LEFT",
                 style=[("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0)])


def _figura_png(dibujar, ancho_mm, alto_mm, *args, **kwargs):
    fig = Figure(figsize=(ancho_mm / 25.4, alto_mm / 25.4), dpi=200, layout="constrained")
    FigureCanvasAgg(fig)
    dibujar(fig, *args, **kwargs)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200)
    buf.seek(0)
    return Image(buf, width=ancho_mm * mm, height=alto_mm * mm)


def _logo(alto_mm):
    ruta = ruta_recurso("assets/logo_geype.png")
    if not ruta.exists():
        return Spacer(1, alto_mm * mm)
    from PIL import Image as PILImage
    with PILImage.open(ruta) as im:
        w, h = im.size
    return Image(str(ruta), width=alto_mm * mm * w / h, height=alto_mm * mm)


def _cabecera(datos, titulo):
    banda = Table([[Paragraph(titulo, _E["titulo"])]], colWidths=[100 * mm],
                  style=[("BACKGROUND", (0, 0), (-1, -1), GRIS), ("TOPPADDING", (0, 0), (-1, -1), 3),
                         ("BOTTOMPADDING", (0, 0), (-1, -1), 4)])
    t = Table([[banda, _logo(14),
                Table([[Paragraph("<font color='white'><b>Fecha</b></font>", _E["normal"]),
                        Paragraph(datos.get("fecha", ""), _E["normal"])]],
                      colWidths=[18 * mm, 22 * mm],
                      style=[("BACKGROUND", (0, 0), (0, 0), AZUL), ("ALIGN", (0, 0), (-1, -1), "CENTER")])]],
              colWidths=[105 * mm, 67 * mm, 105 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (1, 0), (1, 0), "CENTER"),
                           ("ALIGN", (2, 0), (2, 0), "RIGHT"), ("LEFTPADDING", (0, 0), (0, 0), 0)]))
    t.hAlign = "LEFT"
    return t


def _tabla_datos(filas):
    """Bloque «etiqueta  valor»: ambos como párrafos con el mismo estilo para que queden en la misma línea."""
    etiqueta = ParagraphStyle("etq", parent=_E["normal"], textColor=colors.HexColor("#404040"))
    t = Table([[Paragraph(k, etiqueta), Paragraph(v, _E["normal"])] for k, v in filas],
              colWidths=[20 * mm, 160 * mm], hAlign="LEFT")
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, -1), 0),
                           ("LEFTPADDING", (1, 0), (1, -1), 2),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6), ("TOPPADDING", (0, 0), (-1, -1), 0.6)]))
    return t


def _datos_cliente(datos):
    filas = [("Titular", datos.get("titular", "")), ("CUPS", datos.get("cups", "")),
             ("Tarifa", datos.get("tarifa", "")), ("Instalación", datos.get("instalacion", "")),
             ("Dirección", datos.get("direccion", "")), ("Zona", datos.get("zona", ""))]
    return _tabla_datos([(k, f"<b>{escape(v)}</b>" if k in ("Titular", "Tarifa", "Instalación") else escape(v))
                         for k, v in filas])


def _tabla_mensual(estudio, escenarios, ancho_mm):
    n = len(escenarios)
    prop = escenarios[1:]
    cab1 = [Paragraph("Mes", _E["celda"]), Paragraph("Coste potencia facturada (€)<br/>"
                                                     "<font color='#FFFF00'>(sin i.e.)</font>", _E["celda"])]
    cab1 += [""] * (n - 1)
    cab1 += [Paragraph("Ahorro anual (€)<br/><font color='#FFFF00'>(sin i.e.)</font>", _E["celda"])] + \
        [""] * (len(prop) - 1)
    w_mes = 16
    w = (ancho_mm - w_mes) / (n + len(prop))
    estrecha = w < 13          # muchas propuestas: letra algo menor para que «Propuesta» quepa entera
    celda = ParagraphStyle("cm", parent=_E["celda"], fontSize=5.8, leading=6.8) if estrecha else _E["celda"]
    celda_n = ParagraphStyle("cmn", parent=_E["celda_n"], fontSize=5.8, leading=6.8) if estrecha else _E["celda_n"]
    cab2 = [""] + [Paragraph(e.nombre, celda) for e in escenarios] + \
        [Paragraph(f"Actual -<br/>{e.nombre}", celda_n) for e in prop]
    filas = [cab1, cab2]
    for i, mes in enumerate(estudio.etiquetas_meses):
        fila = [mes] + [fmt(e.coste.total_mes[i]) for e in escenarios]
        fila += [fmt(escenarios[0].coste.total_mes[i] - e.coste.total_mes[i]) for e in prop]
        filas.append(fila)
    filas.append(["Total"] + [fmt(e.coste.total) for e in escenarios] + [fmt(e.ahorro) for e in prop])

    t = Table(filas, colWidths=[w_mes * mm] + [w * mm] * (n + len(prop)), repeatRows=2)
    est = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 6.5 if estrecha else 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (1, 0), (n, 0)), ("SPAN", (0, 0), (0, 1)),
        ("BACKGROUND", (0, 0), (-1, 0), AZUL), ("BACKGROUND", (0, 1), (0, 1), AZUL),
        ("TEXTCOLOR", (0, 0), (0, 1), colors.white), ("FONTNAME", (0, 0), (0, 1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 2), (-1, -2), FILA_CLARA),
        ("FONTNAME", (0, 2), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1 if estrecha else 2), ("RIGHTPADDING", (0, 0), (-1, -1), 1 if estrecha else 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2),
    ]
    _marco(est)
    if prop:
        # separación más marcada entre los costes y los ahorros
        est.append(("LINEBEFORE", (n + 1, 0), (n + 1, -1), 1.2, colors.white))
    if prop:
        est.append(("SPAN", (n + 1, 0), (-1, 0)))
    for i in range(n):
        est.append(("BACKGROUND", (1 + i, 1), (1 + i, 1), colors.HexColor(color(i))))
    for j, e in enumerate(prop):
        c = n + 1 + j
        est.append(("BACKGROUND", (c, 1), (c, 1), colors.HexColor(COLORES_SUAVES[(j + 1) % len(COLORES_SUAVES)])))
        est.append(("BACKGROUND", (c, 2), (c, -2), colors.HexColor(COLORES_SUAVES[(j + 1) % len(COLORES_SUAVES)])))
        est.append(("BACKGROUND", (c, -1), (c, -1), colors.HexColor(color(j + 1))))
        est.append(("TEXTCOLOR", (c, -1), (c, -1), colors.white))
        for i in range(len(estudio.meses)):
            if escenarios[0].coste.total_mes[i] - e.coste.total_mes[i] < 0:
                est.append(("TEXTCOLOR", (c, 2 + i), (c, 2 + i), ROJO))
    t.setStyle(TableStyle(est))
    return t


def _tabla_optimizacion(escenarios):
    cab0 = [Paragraph("OPTIMIZACIÓN DE POTENCIA <font color='#FFFF00'>(sin i.e.)</font>", _E["celda"])] + [""] * 13
    cab1 = [Paragraph("Tipo", _E["celda"]), Paragraph("Potencias contratadas (kW)", _E["celda"]), "", "", "", "", "",
            Paragraph("Coste de la potencia facturada (€)", _E["celda"]), "", "",
            Paragraph("Ahorro estimado (€) (*)", _E["celda"]), "",
            Paragraph("Inversión (€) (**)", _E["celda"]), Paragraph("PRS (años)", _E["celda"])]
    cab2 = [""] + [f"P{i}" for i in range(1, 7)] + \
        [Paragraph("T. Fijo", _E["celda"]), Paragraph("T. Excesos", _E["celda"]),
         Paragraph("Total", _E["celda"]), Paragraph("Total (€)", _E["celda"]), "%", "", ""]
    filas = [cab0, cab1, cab2]
    for i, e in enumerate(escenarios):
        fila = [Paragraph(e.nombre, _E["celda"])] + [fmt_pot(p) for p in e.pc] + \
            [fmt(e.coste.total_fijo), fmt(e.coste.total_exceso), fmt(e.coste.total)]
        if i == 0:
            fila += ["-", "-", "", ""]
        else:
            prs = e.prs
            fila += [fmt(e.ahorro), f"{fmt(100 * e.ahorro_pct, 1)}%",
                     fmt(e.inversion.get("total", 0), 2) if e.inversion.get("total") else "-",
                     fmt(prs, 2) if prs is not None else "-"]
        filas.append(fila)
    anchos = [34] + [17.5] * 6 + [20, 20, 20, 20, 14, 22, 22]      # suma = ANCHO_PAGINA
    t = Table(filas, colWidths=[a * mm for a in anchos])
    est = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-1, 0)), ("SPAN", (0, 1), (0, 2)), ("SPAN", (1, 1), (6, 1)), ("SPAN", (7, 1), (9, 1)),
        ("SPAN", (10, 1), (11, 1)), ("SPAN", (12, 1), (12, 2)), ("SPAN", (13, 1), (13, 2)),
        ("BACKGROUND", (0, 0), (-1, 2), AZUL), ("TEXTCOLOR", (0, 0), (-1, 2), colors.white),
        ("FONTNAME", (0, 0), (-1, 2), "Helvetica-Bold"),
        ("BACKGROUND", (7, 2), (7, 2), colors.HexColor("#948A54")),
        ("BACKGROUND", (8, 2), (8, 2), colors.HexColor("#E26B6B")),
        ("BACKGROUND", (1, 3), (-1, -1), FILA_CLARA),
        ("FONTNAME", (10, 4), (10, -1), "Helvetica-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.3), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.3),
    ]
    _marco(est, filas_total=0)
    for i in range(len(escenarios)):
        est.append(("BACKGROUND", (0, 3 + i), (0, 3 + i), colors.HexColor(color(i))))
    t.setStyle(TableStyle(est))
    t.hAlign = "LEFT"
    return t


def _tabla_anexo(titulo, estudio, matriz, dec, fila_extra, etiqueta_extra, max_col=False):
    cab = [Paragraph(titulo, _E["celda"])] + [""] * 7
    cab2 = ["Desde", "Hasta"] + [f"P{i}" for i in range(1, 7)] + (["Máximo"] if max_col else ["Total"])
    filas = [cab + [""], cab2]
    for i, m in enumerate(estudio.meses):
        fila = [m.start_time.strftime("%d/%m/%Y"), m.end_time.strftime("%d/%m/%Y")]
        fila += [fmt(v, dec) for v in matriz[i]]
        fila.append(fmt(matriz[i].max() if max_col else matriz[i].sum(), dec))
        filas.append(fila)
    filas.append(["", "Total" if not max_col else "Máximo"] +
                 [fmt(v, dec) for v in fila_extra[0]] + [fmt(fila_extra[1], dec)])
    filas.append(["", etiqueta_extra] + [f"{fmt(100 * v, 1)}%" for v in fila_extra[2]] + [""])
    t = Table(filas, colWidths=[20 * mm, 20 * mm] + [22 * mm] * 6 + [24 * mm])
    t.setStyle(TableStyle(_marco([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 1), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 1), colors.white), ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 2), (-1, -3), FILA_CLARA),
        ("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.1),
    ], filas_total=2)))
    t.hAlign = "LEFT"
    return t


def _tabla_coste_mensual(estudio, escenarios):
    """Coste mensual de cada escenario: término fijo, excesos y total."""
    n = len(escenarios)
    fuente = 7 if n <= 4 else 6.3
    cab0 = [Paragraph("COSTE DE LA POTENCIA POR MES (€) <font color='#FFFF00'>(sin i.e.)</font>", _E["celda"])] +         [""] * (3 * n)
    cab1 = [""]
    for e in escenarios:
        cab1 += [Paragraph(e.nombre, _E["celda"]), "", ""]
    cab2 = ["Mes"] + ["T. Fijo", "T. Excesos", "Total"] * n
    filas = [cab0, cab1, cab2]
    for i, mes in enumerate(estudio.etiquetas_meses):
        fila = [mes]
        for e in escenarios:
            f, x = e.coste.fijo[i].sum(), e.coste.exceso[i].sum()
            fila += [fmt(f), fmt(x), fmt(f + x)]
        filas.append(fila)
    fila = ["Total"]
    for e in escenarios:
        fila += [fmt(e.coste.total_fijo), fmt(e.coste.total_exceso), fmt(e.coste.total)]
    filas.append(fila)

    w = (ANCHO_PAGINA - 18) / (3 * n)
    t = Table(filas, colWidths=[18 * mm] + [w * mm] * (3 * n))
    est = _marco([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), fuente),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-1, 0)), ("SPAN", (0, 1), (0, 2)),
        ("BACKGROUND", (0, 0), (-1, 0), AZUL), ("BACKGROUND", (0, 1), (0, 2), AZUL),
        ("TEXTCOLOR", (0, 1), (0, 2), colors.white), ("FONTNAME", (0, 0), (-1, 2), "Helvetica-Bold"),
        ("FONTNAME", (0, 3), (0, -1), "Helvetica-Bold"), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
    ])
    for k in range(n):
        c0 = 1 + 3 * k
        est += [("SPAN", (c0, 1), (c0 + 2, 1)),
                ("BACKGROUND", (c0, 1), (c0 + 2, 1), colors.HexColor(color(k))),
                ("BACKGROUND", (c0, 2), (c0 + 2, 2), colors.HexColor(COLORES_SUAVES[k % len(COLORES_SUAVES)])),
                ("BACKGROUND", (c0, 3), (c0 + 1, -2), FILA_CLARA),
                ("BACKGROUND", (c0 + 2, 3), (c0 + 2, -1), colors.HexColor(COLORES_SUAVES[k % len(COLORES_SUAVES)])),
                ("FONTNAME", (c0 + 2, 3), (c0 + 2, -1), "Helvetica-Bold"),
                ("LINEBEFORE", (c0, 1), (c0, -1), 1.2, colors.white)]
    t.setStyle(TableStyle(est))
    return t


def _documento(ruta):
    return SimpleDocTemplate(str(ruta), pagesize=landscape(A4), leftMargin=10 * mm, rightMargin=10 * mm,
                             topMargin=8 * mm, bottomMargin=8 * mm,
                             title="Optimización de la potencia contratada", author="GE&PE Ingeniería")


def _mes_seleccion(estudio, mes):
    """El mes del selector puede venir como índice o como etiqueta ('feb.-26')."""
    if isinstance(mes, str):
        return estudio.etiquetas_meses.index(mes) if mes in estudio.etiquetas_meses else None
    return mes


def _historia_suministro(datos, estudio, escenarios, curva, normativa, incluir_anexos, seleccion_curva):
    """Páginas de un suministro: resumen (página 1) y, opcionalmente, sus anexos."""
    escenarios = _con_nombres(escenarios)
    historia = [_cabecera(datos, "Optimización de la potencia contratada"), Spacer(1, 1 * mm),
                _datos_cliente(datos), Spacer(1, 2 * mm)]

    ancho_tabla = _ancho_mensual(escenarios)
    historia.append(_tabla_y_grafico(_tabla_mensual(estudio, escenarios, ancho_tabla), ancho_tabla,
                                     estudio, escenarios))
    historia += [Spacer(1, 3 * mm), _tabla_optimizacion(escenarios), Spacer(1, 1 * mm)]

    notas = ["* Este porcentaje es sobre el coste de potencia actual"]
    por_concepto = {}      # misma nota de inversión para varias propuestas -> una sola línea
    for e in escenarios[1:]:
        if e.inversion.get("total"):
            por_concepto.setdefault(texto_conceptos(e.inversion["conceptos"]), []).append(e.nombre)
    for conceptos, nombres in por_concepto.items():
        notas.append(f"** Inversión en concepto de {conceptos} ({_lista_y(nombres)})")
    if any(max(e.pc) > max(escenarios[0].pc) for e in escenarios[1:]):
        notas.append("(1) Las propuestas con potencia máxima superior a la actual requieren modificación técnica "
                     "(ampliación de la potencia asociada a la acometida).")
    izquierda = [Paragraph(n, _E["pie"]) for n in notas]

    ini, fin = estudio.meses[0], estudio.meses[-1]
    tipo = ("horaria (cuartos de hora estimados con la potencia media horaria)" if curva.es_horaria
            else "cuartohoraria")
    derecha = [Paragraph("<u>Normativa de referencia:</u>", _E["normal"])]
    derecha += [Paragraph(n, _E["pie"]) for n in normativa]
    derecha += [Spacer(1, 1.5 * mm), Paragraph("<u>Notas:</u>", _E["normal"]), Paragraph(
        f"Estudio realizado utilizando la curva {tipo} de potencia correspondiente al periodo de consumo entre "
        f"{MESES_NOMBRE[ini.month - 1]} {ini.year} y {MESES_NOMBRE[fin.month - 1]} {fin.year}. "
        f"Curva de carga utilizada: desde {curva.fecha_inicio:%d/%m/%Y} hasta "
        f"{(curva.fecha_fin - pd.Timedelta(days=1)):%d/%m/%Y}.", _E["pie"])]
    historia.append(Table([[izquierda, derecha]], colWidths=[118 * mm, 159 * mm], hAlign="LEFT",
                          style=[("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (0, -1), 0),
                                 ("RIGHTPADDING", (-1, 0), (-1, -1), 0)]))

    # la página de resumen se reduce si hace falta para que quepa siempre en una hoja
    historia = [KeepInFrame(277 * mm, 192 * mm, historia, mode="shrink")]

    if incluir_anexos:
        energia = estudio.energia_kwh()
        maximos = estudio.maximos()
        historia += [PageBreak(), _cabecera(datos, "ANEXO: Energía y potencia registradas"), Spacer(1, 4 * mm)]
        tot_e = energia.sum(0)
        historia.append(_tabla_anexo("PERFIL DE CONSUMO (kWh)", estudio, energia, 0,
                                     (tot_e, tot_e.sum(), tot_e / tot_e.sum() if tot_e.sum() else tot_e * 0),
                                     "Porcentaje"))
        historia.append(Spacer(1, 5 * mm))
        max_p = maximos.max(0)
        historia.append(_tabla_anexo("POTENCIAS MÁXIMAS DEMANDADAS (kW)", estudio, maximos, 0,
                                     (max_p, max_p.max(), max_p / max_p.max() if max_p.max() else max_p * 0),
                                     "Porcentaje", max_col=True))
        historia += [PageBreak(), _cabecera(datos, "ANEXO: Coste de la potencia por mes"), Spacer(1, 4 * mm),
                     _tabla_coste_mensual(estudio, escenarios), Spacer(1, 2 * mm),
                     Paragraph("T. Fijo: término de potencia contratada (peajes + cargos). T. Excesos: facturación por "
                               "excesos de potencia cuartohorarios (art. 9 de la Circular 3/2020).", _E["pie"])]
        historia += [PageBreak(), _cabecera(datos, "ANEXO: Curva de carga"), Spacer(1, 2 * mm),
                     _figura_png(dibujar_curva_selector, 275, 105, estudio, escenarios, seleccion_curva[0],
                                 _mes_seleccion(estudio, seleccion_curva[1])),
                     Spacer(1, 1 * mm),
                     _figura_png(dibujar_maximos, 275, 62, estudio, escenarios[0].pc)]

    return historia


def generar_pdf(ruta, datos, estudio, escenarios, curva, normativa, incluir_anexos=True, seleccion_curva=(0, None)):
    """
    Informe de un suministro.
    datos: dict con titular, cups, tarifa, instalacion, direccion, zona, fecha.
    escenarios: [Actual, Propuesta 1, ...] (calculo.Escenario).
    curva: potencia.lector_curva.CurvaCargada (para las notas).
    seleccion_curva: (periodo, mes) del gráfico de la curva, igual que el selector de la interfaz
                     (periodo 0 = todos; mes None = año completo, índice o etiqueta 'feb.-26').
    """
    _documento(ruta).build(_historia_suministro(datos, estudio, escenarios, curva, normativa, incluir_anexos,
                                                seleccion_curva))


_CAMPOS_CUPS = {  # clave: (cabecera, cabecera corta, peso en el reparto del ancho)
    "coste": ("Coste (€)", "Coste", 1.0),
    "ahorro": ("Ahorro (€)", "Ahorro", 1.0),
    "pct": ("%", "%", 0.8),
    "inv": ("Inversión (€)", "Inv.", 1.2),
    "prs": ("PRS (años)", "PRS", 0.7),
}


def _celdas_propuesta(e, campos):
    """Valores de una propuesta para la tabla de CUPS (guiones si el CUPS no la tiene)."""
    if e is None:
        return ["-"] * len(campos)
    inv = e.inversion.get("total", 0)
    valores = {"coste": fmt(e.coste.total), "ahorro": fmt(e.ahorro), "pct": f"{fmt(100 * e.ahorro_pct, 1)}%",
               "inv": fmt(inv, 2) if inv else "-", "prs": fmt(e.prs, 2) if e.prs is not None else "-"}
    return [valores[c] for c in campos]


def _tabla_resumen_cups(filas, suministros, escenarios):
    """
    Una fila por suministro: coste actual y, por cada propuesta, coste, ahorro, %, inversión y PRS
    (con 5 o más propuestas, solo coste y ahorro para que siga siendo legible).
    Fila final con los totales del resumen conjunto.
    """
    n_prop = len(escenarios) - 1
    compacta = n_prop >= 3
    campos = ["coste", "ahorro"] if n_prop >= 5 else list(_CAMPOS_CUPS)
    nc = len(campos)
    fuente = 7 if not compacta else 6.2
    celda = ParagraphStyle("cr", parent=_E["celda"], fontSize=fuente, leading=fuente + 1.2)
    celda_n = ParagraphStyle("crn", parent=_E["celda_n"], fontSize=fuente, leading=fuente + 1.2)
    fijas = ["Nº", "CUPS", "Denominación", "Tarifa", "Zona", "Coste actual (€)"]
    titulo = (lambda e: f"{e.nombre} (€)") if compacta else (lambda e: e.nombre)
    cab0 = [Paragraph(c, celda) for c in fijas]
    cab1 = [""] * len(fijas)
    for e in escenarios[1:]:
        cab0 += [Paragraph(titulo(e), celda)] + [""] * (nc - 1)
        cab1 += [Paragraph(_CAMPOS_CUPS[c][1 if compacta else 0], celda_n) for c in campos]
    datos = [cab0, cab1]

    anchos_fijos = [7, 37, 30, 11, 16, 18] if not compacta else [6, 33, 22, 10, 14, 16]
    # la denominación se parte en líneas a mano (texto plano) para que quede a la altura del resto de la fila
    ancho_den = int(anchos_fijos[2] * mm / (fuente * 0.52))
    for k, (f, s) in enumerate(zip(filas, suministros), 1):
        escs = s["escenarios"]
        fila = [str(k), f["cups"], textwrap.fill(f["denominacion"] or "-", ancho_den), f["tarifa"], f["zona"],
                fmt(escs[0].coste.total)]
        for j in range(1, n_prop + 1):
            fila += _celdas_propuesta(escs[j] if j < len(escs) else None, campos)
        datos.append(fila)
    fila = ["", "TOTAL", "", "", "", fmt(escenarios[0].coste.total)]
    for e in escenarios[1:]:
        fila += _celdas_propuesta(e, campos)
    datos.append(fila)

    reparto = [_CAMPOS_CUPS[c][2] for c in campos]
    unidad = (ANCHO_PAGINA - sum(anchos_fijos)) / (sum(reparto) * n_prop)
    anchos = anchos_fijos + [r * unidad for r in reparto] * n_prop
    t = Table(datos, colWidths=[a * mm for a in anchos], repeatRows=2)
    ult = len(fijas) - 1
    est = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), fuente),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (ult, 2), (-1, -1), "RIGHT"), ("ALIGN", (2, 2), (2, -1), "LEFT"),
        ("BACKGROUND", (0, 0), (ult, 1), AZUL),
        ("BACKGROUND", (0, 2), (-1, -2), FILA_CLARA),
        ("BACKGROUND", (ult, 2), (ult, -2), colors.HexColor(COLORES_SUAVES[0])),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2 if not compacta else 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2 if not compacta else 1.5),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]
    est += [("SPAN", (c, 0), (c, 1)) for c in range(len(fijas))]
    _marco(est)
    for j in range(1, n_prop + 1):
        c0 = len(fijas) + nc * (j - 1)
        suave = colors.HexColor(COLORES_SUAVES[j % len(COLORES_SUAVES)])
        est += [("SPAN", (c0, 0), (c0 + nc - 1, 0)),
                ("BACKGROUND", (c0, 0), (c0 + nc - 1, 0), colors.HexColor(color(j))),
                ("BACKGROUND", (c0, 1), (c0 + nc - 1, 1), suave),
                ("BACKGROUND", (c0, 2), (c0, -2), suave),
                ("BACKGROUND", (c0 + 1, -1), (c0 + 1, -1), colors.HexColor(color(j))),
                ("TEXTCOLOR", (c0 + 1, -1), (c0 + 1, -1), colors.white),
                ("LINEBEFORE", (c0, 0), (c0, -1), 1.2, colors.white)]
        for k, s in enumerate(suministros):
            escs = s["escenarios"]
            if j < len(escs) and escs[j].ahorro < 0:
                est.append(("TEXTCOLOR", (c0 + 1, 2 + k), (c0 + 1, 2 + k), ROJO))
        if escenarios[j].ahorro < 0:
            est.append(("BACKGROUND", (c0 + 1, -1), (c0 + 1, -1), ROJO))
    t.setStyle(TableStyle(est))
    t.hAlign = "LEFT"
    return t


def generar_pdf_multipunto(ruta, datos, suministros, resumen, filas_cups, normativa, incluir_anexos=True,
                           seleccion_curva=(0, None)):
    """
    Informe multipunto: resumen conjunto y después las páginas de cada suministro.
    datos: dict con titular y fecha.
    suministros: [{"datos": dict, "estudio": Estudio, "escenarios": [...], "curva": CurvaCargada}, ...]
    resumen: calculo.ResumenMultipunto.  filas_cups: filas de la tabla resumen por CUPS.
    """
    info = _tabla_datos([("Titular", f"<b>{escape(datos.get('titular', ''))}</b>"),
                         ("Suministros", f"<b>{len(suministros)}</b> (estudio multipunto)")])
    escs = _con_nombres(resumen.escenarios)
    ancho_tabla = _ancho_mensual(escs)
    portada = [_cabecera(datos, "Optimización de la potencia – Resumen multipunto"), Spacer(1, 1 * mm), info,
               Spacer(1, 2 * mm),
               _tabla_y_grafico(_tabla_mensual(resumen, escs, ancho_tabla), ancho_tabla, resumen, escs),
               Spacer(1, 3 * mm), _tabla_resumen_cups(filas_cups, suministros, escs), Spacer(1, 1.5 * mm),
               Paragraph("El resumen conjunto compara la situación actual con cada propuesta: la Propuesta 1 (óptima) "
                         "suma las potencias óptimas de cada suministro y la Propuesta 2 y siguientes suman la "
                         "propuesta con ese número de cada suministro. El detalle de cada suministro figura en las "
                         "páginas siguientes.", _E["pie"])]
    faltan = {}      # suministros sin la propuesta -> propuestas afectadas
    for k, idx in sorted(resumen.sin_propuesta.items()):
        faltan.setdefault(tuple(idx), []).append(k)
    for idx, props in faltan.items():
        nums = _lista_y([str(i + 1) for i in idx])
        uno, una = len(idx) == 1, len(props) == 1
        texto = f"Propuesta {props[0]}" if una else "Propuestas " + _lista_y([str(k) for k in props])
        quien = f"el suministro n.º {nums} no tiene" if uno else f"los suministros n.º {nums} no tienen"
        portada.append(Paragraph(f"{texto}: {quien} {'esta propuesta' if una else 'estas propuestas'}; en el total "
                                 f"{'se suma' if uno else 'se suman'} con su situación actual.", _E["pie"]))
    if resumen.periodos_distintos:
        portada.append(Paragraph("Atención: las curvas de los suministros no cubren los mismos meses; el resumen "
                                 "mensual suma los meses disponibles de cada uno.", _E["pie"]))
    portada.append(Spacer(1, 1.5 * mm))
    portada.append(Paragraph("<u>Normativa de referencia:</u>", _E["normal"]))
    portada += [Paragraph(n, _E["pie"]) for n in normativa]
    # con pocos CUPS la portada se reduce si hace falta para que quepa en una hoja;
    # con muchos, la tabla de CUPS sigue en la hoja siguiente (repite la cabecera)
    historia = [KeepInFrame(ANCHO_PAGINA * mm, 192 * mm, portada, mode="shrink")] if len(suministros) <= 12 \
        else portada
    for s in suministros:
        historia.append(PageBreak())
        historia += _historia_suministro(s["datos"], s["estudio"], s["escenarios"], s["curva"], normativa,
                                         incluir_anexos, seleccion_curva)
    _documento(ruta).build(historia)
