# -*- coding: utf-8 -*-
"""
INFORME PDF DEL ANALISIS DE OFERTAS (mismo formato que la plantilla Excel)
=========================================================================
    pdf(e, res, ref, tecnico)
Páginas:
    1. Ficha del suministro con su perfil de consumo y gráficas (monopunto) o «Relación de puntos
       de suministro» con las gráficas de la cartera: mensual por periodo, anual por periodo, por
       tarifa y suministros de mayor consumo (multipunto). Con quién ha hecho el estudio y su
       email.
    2. «Resumen del estudio de ofertas de renovación»: igual que la hoja «Informe» de la plantilla
       (tablas de precio fijo e indexado con el contrato actual, la mejor oferta, la variación y
       el ranking por la variación; cuadro del precio de mercado por trimestre de cada año y notas).
    3. «Condiciones de las ofertas».
Cada página va sobre el papel de GE&PE con su marca de agua (assets/papel_<orientación>_<n>.pdf):
los datos quedan entre el logo y el «GE&PE» gris, sin pisarlos. En cada orientación se usa el
papel 1 (logo grande) si los datos caben y si no el 2 (logo pequeño). El tamaño se adapta a
las columnas: A4 y, si no caben con letra legible, A3 (el papel se escala). Sin Streamlit.
"""

import datetime as dt
import functools
import io
from xml.sax.saxutils import escape

from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, landscape, portrait
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle)

import config
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo, resultados
from analisis_ofertas import modelo as mo

LOGO = config.RECURSOS / "assets" / "logo_geype.png"
C = {k: colors.HexColor(v) for k, v in {
    "navy": "#002060", "teal": "#31869B", "lteal": "#DCE6F1", "orange": "#FCD5B4",
    "act": "#FFFFCC", "minh": "#215967", "min": "#F2F2F2", "verde": "#C6EFCE",
    "verde_t": "#006100", "rojo": "#C00000", "gris": "#7F7F7F", "linea": "#BFBFBF",
    "azul": "#1F4E79", "azul_c": "#DDEBF7", "amarillo": "#FFF2CC", "total": "#E2EFDA"}.items()}
COL_PERIODOS = ["#1F4E79", "#2E75B6", "#9DC3E6", "#F4B183", "#C55A11", "#A5A5A5"]
MARGEN = 10 * mm
COL_ETIQ, COL_MIN, COL_OF_MAX, COL_OF_MIN = 34 * mm, 28 * mm, 24 * mm, 20 * mm


# ------------------------------------------------------------------- formatos
def _eur(v, sufijo=""):
    return "-" if v is None else f"{v:,.0f}".replace(",", ".") + sufijo


def _num(v, d=2):
    return "-" if v is None else f"{v:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _pct(v):
    return "-" if v is None else f"{v * 100:.2f}%".replace(".", ",")


def _transp(color, alfa):
    """El mismo color con transparencia (0 = invisible, 1 = opaco)."""
    return colors.Color(color.red, color.green, color.blue, alpha=alfa)


def _st(nombre, tam=8, negrita=False, color=colors.black, alin=1, leading=None, cursiva=False):
    fuente = "Helvetica-BoldOblique" if negrita and cursiva else "Helvetica-Bold" if negrita \
        else "Helvetica-Oblique" if cursiva else "Helvetica"
    return ParagraphStyle(nombre, fontName=fuente, fontSize=tam, leading=leading or tam * 1.22,
                          textColor=color, alignment=alin)


def nombre_corto(comercializadora: str, nombre: str = "") -> str:
    """ENDESA, EDP, TOTAL… (primera palabra, sin comas; TOTALENERGIES → TOTAL)."""
    base = (comercializadora or nombre or "").strip().split(" ")[0].replace(",", "").upper()
    return "TOTAL" if base == "TOTALENERGIES" else base


def cabecera_oferta(f: dict) -> str:
    """Como la plantilla: COMERCIALIZADORA [index. PT/PP] [NN M] + (cómo van los SSAA)."""
    texto = nombre_corto(f["comercializadora"], f["nombre"])
    if f["tipo"] != "Fijo":
        marca = f" {(f.get('producto') or '')} {f['nombre']} ".upper()
        tipo = "PP" if " PP " in marca else "PT" if " PT " in marca else (
            "PP" if f["modalidad"] == mo.PASS_POOL else "PT")
        texto += f" index. {tipo}"
    if f["duracion"] > 12:
        texto += f" {f['duracion']} M"
    extra = f["ssaa"] + (", cierres" if f.get("cierres") else "")
    return f"{texto}<br/>({extra})"


def tecnico_de(usuario: str) -> dict:
    """{'nombre', 'email'} del usuario de la plataforma (Administración → usuarios)."""
    try:
        import almacen
        u = next((x for x in almacen.listar_usuarios() if x["usuario"] == usuario), None)
    except Exception:
        u = None
    return {"nombre": (u or {}).get("nombre") or usuario or "", "email": (u or {}).get("email") or ""}


# ------------------------------------------------------------------- piezas comunes
def _cabecera(e: mo.Estudio, ancho: float, titulo: str, ref: str, con_logo: bool = True) -> list:
    sub = " · ".join(x for x in (f"CIF {e.cif}" if e.cif else "", f"Ref. {ref}" if ref else "",
                                  f"Inicio de suministro {e.inicio:%d/%m/%Y}" if e.inicio else "")
                     if x)
    textos = [Paragraph(escape((e.cliente or "").upper()), _st("cli", 12, True, alin=2)),
              Spacer(1, 1 * mm), Paragraph(escape(sub), _st("sub", 8, color=C["gris"], alin=2))]
    if con_logo:
        logo = Image(str(LOGO), width=32 * mm, height=16 * mm) if LOGO.exists() else ""
        arriba = Table([[logo, textos]], colWidths=[40 * mm, ancho - 40 * mm])
    else:
        # sobre el papel de GE&PE (que ya lleva el logo): cliente y datos en una sola línea
        arriba = Table([[Paragraph(escape((e.cliente or "").upper()), _st("cli", 12, True, alin=0)),
                         Paragraph(escape(sub), _st("sub", 8, color=C["gris"], alin=2))]],
                       colWidths=[ancho * 0.6, ancho * 0.4])
    arriba.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    barra = Table([[Paragraph(titulo, _st("tit", 11, True, colors.white))]], colWidths=[ancho],
                  rowHeights=[8 * mm])
    barra.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), C["navy"]),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    return [arriba, Spacer(1, 2 * mm), barra, Spacer(1, 3 * mm)]


def _realizado(tecnico: dict, fecha: dt.date, ancho: float) -> Table:
    filas = [[Paragraph("Estudio realizado por", _st("r1", 8, True, colors.white, 0)),
              Paragraph(escape(tecnico.get("nombre") or ""), _st("r2", 8, alin=0)),
              Paragraph("Email", _st("r3", 8, True, colors.white, 0)),
              Paragraph(escape(tecnico.get("email") or ""), _st("r4", 8, alin=0)),
              Paragraph("Fecha", _st("r5", 8, True, colors.white, 0)),
              Paragraph(f"{fecha:%d/%m/%Y}", _st("r6", 8))]]
    pesos = [0.2, 0.28, 0.07, 0.27, 0.07, 0.11]
    t = Table(filas, colWidths=[ancho * p for p in pesos], hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), C["azul"]),
                           ("BACKGROUND", (2, 0), (2, 0), C["azul"]),
                           ("BACKGROUND", (4, 0), (4, 0), C["azul"]),
                           ("BACKGROUND", (1, 0), (1, 0), C["azul_c"]),
                           ("BACKGROUND", (3, 0), (3, 0), C["azul_c"]),
                           ("BACKGROUND", (5, 0), (5, 0), C["azul_c"]),
                           ("GRID", (0, 0), (-1, -1), 0.4, colors.white),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    return t


# ------------------------------------------------------------------- resumen de ofertas
def texto_mejora(m: dict) -> str:
    """«−3,20 % · −12.345 € vs 1ª ronda» (mejora del coste del primer año)."""
    if not m:
        return "-"
    return f"{_pct(m['variacion'])} · {_eur(m['dif'], ' €')} vs {m['ronda_anterior']}ª ronda"


def _tabla_ofertas(titulo: str, filas: list, act: dict, n_cols: int, ancho_of: float,
                   mejora: dict = None) -> Table:
    """Tabla como la de la hoja «Informe» (contrato actual · ofertas · mejor oferta)."""
    s_cab = _st("cab", 7.5, True, alin=1)
    s_val = _st("val", 8)
    s_lbl = _st("lbl", 8, True)
    s_gris = _st("gris", 6.5, color=C["gris"], cursiva=True)
    filas = filas + [None] * (n_cols - len(filas))
    validas = [f for f in filas if f]
    # bloque principal: primer año frente al contrato actual en esos meses (ranking y mejor
    # oferta); bloque gris: periodo total de cada oferta con su propio ranking
    minimo = min(validas, key=resultados.orden_anio1) if validas else None
    rank1 = {f["nombre"]: i for i, f in enumerate(sorted(validas, key=resultados.orden_anio1), 1)}
    rankp = {f["nombre"]: i for i, f in enumerate(sorted(validas, key=resultados.orden), 1)}
    rojo = lambda v: C["rojo"] if (v or 0) < 0 else colors.black

    def celda(f, texto, estilo):
        return Paragraph(texto(f), estilo) if f else Paragraph("-", estilo)

    def gris(texto):
        return ([celda(f, texto, s_gris) for f in filas]
                + [Paragraph(texto(minimo) if minimo else "-", s_gris)])

    actual_txt = f"Contrato actual<br/>{escape(act['nombre'])}"
    eur_mwh = lambda v: "-" if v is None else _num(v, 2) + " €/MWh"
    datos = [
        [Paragraph(actual_txt, _st("ac", 8, True, colors.white)),
         Paragraph(escape(titulo), _st("tt", 9, True, colors.white))] + [""] * (n_cols - 1)
        + [Paragraph("Mejor oferta", _st("om", 9, True, colors.white))],
        [""] + [Paragraph(cabecera_oferta(f), s_cab) if f else Paragraph("-", s_cab) for f in filas]
        + [""],
        # primer año
        [Paragraph(_eur(act["total"], " €/año"), s_val)]
        + [Paragraph(_eur(f["anio1"], " €/año"), s_val) if f else "-" for f in filas]
        + [Paragraph(_eur(minimo["anio1"], " €/año") if minimo else "-", _st("mv", 8, True))],
        [Paragraph(eur_mwh(act["eur_mwh"]), s_val)]
        + [Paragraph(eur_mwh(f["anio1_eur_mwh"]), s_val) if f else "-" for f in filas]
        + [Paragraph(eur_mwh(minimo["anio1_eur_mwh"]) if minimo else "-", _st("mk", 8, True))],
        [Paragraph("Variación sobre<br/>contrato actual", s_lbl)]
        + [Paragraph(_pct(f["variacion_anio1"]), _st("v", 8, color=rojo(f["variacion_anio1"])))
           if f else "-" for f in filas]
        + [Paragraph(_pct(minimo["variacion_anio1"]) if minimo else "-",
                     _st("mv2", 8, True, rojo(minimo and minimo["variacion_anio1"])))],
        [""] + [Paragraph(_eur(f["dif_anio1"], " €"), _st("d", 8, color=rojo(f["dif_anio1"])))
                if f and f["dif_anio1"] is not None else "-" for f in filas]
        + [Paragraph(_eur(minimo["dif_anio1"], " €") if minimo and minimo["dif_anio1"] is not None
                     else "-", _st("md", 8, True, rojo(minimo and minimo["dif_anio1"])))],
        [Paragraph("Ranking de oferta", s_lbl)]
        + [Paragraph(str(rank1[f["nombre"]]), _st("rk", 9, True)) if f else "-" for f in filas]
        + [Paragraph(cabecera_oferta(minimo).replace("<br/>", " ") if minimo else "-",
                     _st("mn", 7, True))],
        # periodo total de cada oferta (en gris)
        [Paragraph("Contrato actual en el mismo periodo que la oferta analizada (€/año)", s_gris)]
        + gris(lambda f: _eur(f["actual_mismo_periodo"])),
        [Paragraph("Oferta analizada en su periodo total (€/año)", s_gris)]
        + gris(lambda f: _eur(f["media_anual"])),
        [Paragraph("Diferencia en su periodo total", s_gris)]
        + gris(lambda f: _pct(f["variacion"])),
        [""] + gris(lambda f: _eur(f["media_anual"] - f["actual_mismo_periodo"], " €")
                    if f["actual_mismo_periodo"] else "-"),
        [Paragraph("Ranking en su periodo total", s_gris)]
        + gris(lambda f: str(rankp[f["nombre"]]))]
    altos = [6.5 * mm, 11.5 * mm, 5 * mm, 5 * mm, 4.6 * mm, 4.6 * mm, 8 * mm,
             8.5 * mm, 6 * mm, 4.4 * mm, 4.4 * mm, 4.8 * mm]
    if mejora is not None:
        # mejora del primer año frente a la oferta de la misma comercializadora en la ronda
        # anterior (en gris, como el periodo total)
        datos.append([Paragraph("Mejora frente a la ronda anterior (primer año)", s_gris)]
                     + gris(lambda f: texto_mejora(mejora.get(f["nombre"]))))
        altos.append(8 * mm)
    t = Table(datos, colWidths=[COL_ETIQ] + [ancho_of] * n_cols + [COL_MIN], rowHeights=altos)
    ult = n_cols + 1
    est = [("SPAN", (0, 0), (0, 1)), ("SPAN", (1, 0), (n_cols, 0)), ("SPAN", (ult, 0), (ult, 1)),
           ("SPAN", (0, 4), (0, 5)), ("SPAN", (0, 9), (0, 10)),
           ("BACKGROUND", (0, 0), (0, 1), C["teal"]), ("BACKGROUND", (1, 0), (n_cols, 0), C["teal"]),
           ("BACKGROUND", (ult, 0), (ult, 1), C["minh"]),
           ("BACKGROUND", (1, 1), (n_cols, 1), C["lteal"]),
           ("BACKGROUND", (0, 2), (0, 3), C["act"]), ("BACKGROUND", (0, 4), (0, 6), C["lteal"]),
           ("BACKGROUND", (ult, 2), (ult, 6), C["min"]),
           ("GRID", (0, 0), (-1, 6), 0.4, C["linea"]),
           ("BOX", (0, 0), (0, 6), 1.2, colors.black), ("BOX", (1, 0), (n_cols, 6), 1.2, colors.black),
           ("BOX", (ult, 0), (ult, 6), 1.2, colors.black),
           ("BOX", (0, 7), (-1, -1), 0.4, C["linea"]), ("INNERGRID", (0, 7), (-1, -1), 0.4, C["linea"]),
           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
           ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
           ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]
    for i, f in enumerate(filas, start=1):
        if f and f["duracion"] > 12:
            est.append(("BACKGROUND", (i, 1), (i, 1), C["orange"]))
        if f and minimo and f["nombre"] == minimo["nombre"]:
            est.append(("BACKGROUND", (i, 2), (i, 2), C["verde"]))
    t.setStyle(TableStyle(est))
    return t


def _cuadro_mercado(e: mo.Estudio, n_meses: int) -> Table:
    ctx = calculo.Contexto(e)
    anios = []
    for y in range(3):
        if n_meses > 12 * y:
            anios.append(y)
    cab = [Paragraph("Trimestre", _st("q0", 8, True))] + [
        Paragraph(str(mo.sumar_meses(e.inicio, 12 * y).year), _st("qa", 8, True)) for y in anios]
    datos = [[Paragraph("Precio de mercado estimado (€/MWh)", _st("qt", 8.5, True, colors.white, 0))]
             + [""] * len(anios), cab]
    medias = []
    for q in range(4):
        fila = [Paragraph(f"Q{(mo.sumar_meses(e.inicio, 3 * q).month - 1) // 3 + 1}",
                          _st("qq", 8, True))]
        for y in anios:
            ks = [12 * y + 3 * q + i for i in range(3) if 12 * y + 3 * q + i < n_meses]
            if ks:
                v = [ctx.mercado(f.year, f.month)[0] for f in (mo.sumar_meses(e.inicio, k) for k in ks)]
                fila.append(Paragraph(_num(sum(v) / len(v)), _st("qv", 8, alin=1)))
            else:
                fila.append("")
        datos.append(fila)
    fila = [Paragraph("Media anual", _st("qm", 8, True, colors.white, 1))]
    for y in anios:
        ks = range(12 * y, min(n_meses, 12 * y + 12))
        v = [ctx.mercado(f.year, f.month)[0] for f in (mo.sumar_meses(e.inicio, k) for k in ks)]
        medias.append(sum(v) / len(v))
        fila.append(Paragraph(_num(medias[-1]), _st("qmv", 8, True, colors.white, 1)))
    datos.append(fila)
    t = Table(datos, colWidths=[28 * mm] + [22 * mm] * len(anios))
    t.setStyle(TableStyle([("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 0), C["teal"]),
                           ("BACKGROUND", (0, 1), (-1, 1), C["lteal"]),
                           ("BACKGROUND", (0, 2), (0, 5), C["lteal"]),
                           ("BACKGROUND", (0, 6), (-1, 6), C["teal"]),
                           ("GRID", (0, 1), (-1, -1), 0.4, C["linea"]),
                           ("BOX", (0, 0), (-1, -1), 1.2, colors.black),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 1), (-1, -1), "CENTER"),
                           ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
    return t, [mo.sumar_meses(e.inicio, 12 * y).year for y in anios], medias


def _notas(e: mo.Estudio, anios: list, medias: list) -> list:
    sesion = dt.date.fromisoformat(e.sesion_omip).strftime("%d/%m/%Y") if e.sesion_omip else "-"
    lista = [f"{_num(m)} €/MWh en {a}" for a, m in zip(anios, medias)]
    texto_m = (", ".join(lista[:-1]) + " y " + lista[-1]) if len(lista) > 1 else (lista[0] if lista else "")
    return [
        f"El precio de mercado para el análisis de las ofertas indexadas se ha estimado ponderando al "
        f"{e.pct_omip * 100:.0f} % las cotizaciones de OMIP publicadas el {sesion} para el periodo de "
        f"contrato y al {100 - e.pct_omip * 100:.0f} % la media mensual del precio OMIE de los últimos "
        f"3 años. Resulta un precio de mercado estimado (media de los meses que cubren las ofertas) de "
        f"{texto_m}.",
        "El coste recoge todos los conceptos de la factura eléctrica (término de potencia, energía, "
        "peajes y cargos, servicios de ajuste, FNEE, garantías de origen, impuesto municipal e impuesto "
        "eléctrico) salvo IVA, excesos de potencia y energía reactiva.",
        "€/año y €/MWh: coste del primer año de cada oferta y del contrato actual. La variación, el "
        "ranking y la mejor oferta comparan ese primer año con el contrato actual en esos mismos "
        "meses.",
        "En gris, el periodo total de cada oferta: su coste medio anual durante toda su duración "
        "frente al contrato actual en esos mismos meses, la diferencia y el ranking por esa "
        "diferencia. Las ofertas con cabecera naranja duran más de 12 meses."]


def _pagina_informe(e, res, ref, n_cols, ancho_of, ancho_util, con_logo=True, sufijo="",
                    mejora=None) -> list:
    filas = resultados.comparativa(res)
    actual = res.get("actual")
    act = {"nombre": (nombre_corto(actual.oferta.comercializadora, "") +
                      ("" if actual.oferta.modalidad == mo.FIJO else " index.")) if actual else "",
           **resultados.actual_primer_anio(res)}
    h = _cabecera(e, ancho_util, "RESUMEN DEL ESTUDIO DE OFERTAS DE RENOVACIÓN" + sufijo, ref,
                  con_logo)
    for tipo, titulo in (("Fijo", "Estimación facturación ofertas a precio fijo (€/año)"),
                         ("Indexado", "Estimación facturación ofertas indexadas (€/año)")):
        de_tipo = [f for f in filas if f["tipo"] == tipo]
        if de_tipo:
            h += [_tabla_ofertas(titulo, de_tipo, act, n_cols, ancho_of, mejora),
                  Spacer(1, 3 * mm)]
    n_meses = max((r.n_meses for r in res["ofertas"]), default=12)
    cuadro, anios, medias = _cuadro_mercado(e, n_meses)
    sesion = dt.date.fromisoformat(e.sesion_omip).strftime("%d/%m/%Y") if e.sesion_omip else "-"
    izq = [cuadro, Spacer(1, 1 * mm),
           Paragraph(f"Cotización OMIP del {sesion}", _st("cot", 7, color=C["gris"], alin=0, cursiva=True))]
    der = [Paragraph("Notas:", _st("nt", 8.5, True, alin=0))]
    for n in _notas(e, anios, medias):
        der += [Spacer(1, 1 * mm), Paragraph(escape(n), _st("nn", 7.5, alin=4, cursiva=True))]
    ancho_cuadro = 28 * mm + 22 * mm * len(anios) + 8 * mm
    bloque = Table([[izq, der]], colWidths=[ancho_cuadro, ancho_util - ancho_cuadro])
    bloque.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    h.append(bloque)
    return h


# ------------------------------------------------------------------- ficha de suministros
def _consumo_periodos(s) -> list:
    return [sum(s.consumo_mes(m)[p] for m in range(1, 13)) for p in range(6)]


def _relacion(e: mo.Estudio, ancho_util: float) -> Table:
    activos = [s for s in e.suministros if s.activo]
    cab1 = ["Nº", "Denominación", "Dirección de suministro", "CUPS", "Tarifa",
            "Potencia contratada (kW)", "", "", "", "", "",
            "Perfil de consumo anual (kWh)", "", "", "", "", "", "Consumo anual (kWh)"]
    cab2 = [""] * 5 + [f"P{p}" for p in range(1, 7)] * 2 + [""]
    s_c = _st("rc", 7.5, True, colors.white)
    datos = [[Paragraph(x, s_c) if x else "" for x in cab1], [Paragraph(x, s_c) if x else "" for x in cab2]]
    tot = [0.0] * 6
    for i, s in enumerate(activos, 1):
        cp = _consumo_periodos(s)
        tot = [a + b for a, b in zip(tot, cp)]
        n = mo.n_periodos(s.tarifa)
        datos.append([str(i), Paragraph(escape(s.alias or ""), _st("ra", 7.5, True)),
                      Paragraph(escape(s.direccion or ""), _st("rd", 7, alin=0)), s.cups, s.tarifa]
                     + [(_num(s.potencias[p], 0) if p < n and s.potencias[p] else "") for p in range(6)]
                     + [(_eur(cp[p]) if p < n else "") for p in range(6)] + [_eur(sum(cp))])
    datos.append(["", Paragraph(f"TOTAL ({len(activos)} suministros)", _st("rt", 7.5, True, alin=0)),
                  "", "", ""] + [""] * 6 + [_eur(v) for v in tot] + [_eur(sum(tot))])
    fijos = [8, 30, 0, 38, 12] + [12] * 6 + [16] * 6 + [21]
    resto = ancho_util / mm - sum(fijos)
    fijos[2] = max(40, resto)
    t = Table(datos, colWidths=[w * mm for w in fijos], repeatRows=2)
    t.setStyle(TableStyle([("SPAN", (5, 0), (10, 0)), ("SPAN", (11, 0), (16, 0)),
                           *[("SPAN", (c, 0), (c, 1)) for c in (0, 1, 2, 3, 4, 17)],
                           ("BACKGROUND", (0, 0), (-1, 1), C["azul"]),
                           ("BACKGROUND", (1, 2), (1, -2), C["amarillo"]),
                           ("BACKGROUND", (0, -1), (-1, -1), C["total"]),
                           ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                           ("FONTSIZE", (0, 2), (-1, -1), 7.5),
                           ("ALIGN", (0, 2), (-1, -1), "CENTER"),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("GRID", (0, 0), (-1, -1), 0.4, C["linea"]),
                           ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    return t


def _ficha_monopunto(e: mo.Estudio, s, ancho_util: float, extra: float = 0) -> list:
    s_l = _st("fl", 8.5, True, colors.white)
    s_v = _st("fv", 8.5)
    datos = [[Paragraph("Commodity", s_l), Paragraph("ELECTRICIDAD", s_v), "", ""],
             [Paragraph("Titular del contrato", s_l), Paragraph(escape(e.cliente or ""), s_v),
              Paragraph("CIF", s_l), Paragraph(escape(e.cif or ""), s_v)],
             [Paragraph("Dirección de suministro", s_l), Paragraph(escape(s.direccion or ""), s_v), "", ""],
             [Paragraph("Número de CUPS", s_l), Paragraph(escape(s.cups), s_v),
              Paragraph("Denominación", s_l), Paragraph(escape(s.alias or ""), s_v)]]
    w = ancho_util / mm
    t = Table(datos, colWidths=[w * .22 * mm, w * .38 * mm, w * .16 * mm, w * .24 * mm])
    t.setStyle(TableStyle([("SPAN", (1, 0), (3, 0)), ("SPAN", (1, 2), (3, 2)),
                           *[("BACKGROUND", (0, r), (0, r), C["azul"]) for r in range(4)],
                           ("BACKGROUND", (2, 1), (2, 1), C["azul"]), ("BACKGROUND", (2, 3), (2, 3), C["azul"]),
                           *[("BACKGROUND", (1, r), (1, r), C["azul_c"]) for r in range(4)],
                           ("BACKGROUND", (3, 1), (3, 1), C["azul_c"]), ("BACKGROUND", (3, 3), (3, 3), C["azul_c"]),
                           ("GRID", (0, 0), (-1, -1), 0.6, colors.white),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    n = mo.n_periodos(s.tarifa)
    pot = [[""] + [Paragraph(f"P{p}", s_l) for p in range(1, n + 1)] + [Paragraph("Tarifa acceso", s_l)],
           [Paragraph("Potencias (kW)", s_l)] + [Paragraph(_num(s.potencias[p], 0), s_v) for p in range(n)]
           + [Paragraph(s.tarifa, s_v)]]
    tp = Table(pot, colWidths=[w * .22 * mm] + [w * .78 / (n + 1) * mm] * (n + 1))
    tp.setStyle(TableStyle([("BACKGROUND", (1, 0), (-1, 0), C["azul"]), ("BACKGROUND", (0, 1), (0, 1), C["azul"]),
                            ("BACKGROUND", (1, 1), (-1, 1), C["azul_c"]),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.white),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    # perfil de consumo (12 meses del año tipo, en el orden del contrato)
    orden = [mo.sumar_meses(e.inicio, k).month for k in range(12)] if e.inicio else list(range(1, 13))
    s_h = _st("ph", 8, True, colors.white)
    filas = [[Paragraph("PERFIL DE CONSUMO (kWh)", _st("pt", 9, True, colors.white))] + [""] * (n + 1),
             [Paragraph("Mes", s_h)] + [Paragraph(f"P{p}", s_h) for p in range(1, n + 1)] + [Paragraph("Total", s_h)]]
    for m in orden:
        c = s.consumo_mes(m)[:n]
        filas.append([Paragraph(ap.MESES[m - 1], _st("pm", 8, True))] + [_eur(v) for v in c] + [_eur(sum(c))])
    tot = _consumo_periodos(s)[:n]
    filas.append([Paragraph("Total", _st("ptt", 8, True, colors.white))] + [_eur(v) for v in tot] + [_eur(sum(tot))])
    filas.append([Paragraph("Porcentaje", _st("ppc", 8, True, colors.white))]
                 + [_pct(v / sum(tot)) if sum(tot) else "-" for v in tot] + [""])
    tf = Table(filas, colWidths=[w * .16 * mm] + [w * .84 / (n + 1) * mm] * (n + 1))
    tf.setStyle(TableStyle([("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 1), C["azul"]),
                            ("BACKGROUND", (0, 2), (0, -3), C["azul_c"]),
                            ("BACKGROUND", (0, -2), (-1, -1), C["azul"]),
                            ("TEXTCOLOR", (1, -2), (-1, -1), colors.white),
                            ("FONTNAME", (1, -2), (-1, -1), "Helvetica-Bold"),
                            ("FONTNAME", (-1, 2), (-1, -1), "Helvetica-Bold"),
                            ("FONTSIZE", (1, 2), (-1, -1), 8), ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                            ("ALIGN", (0, 0), (-1, 1), "CENTER"),
                            ("GRID", (0, 1), (-1, -1), 0.4, C["linea"]),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 0.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6)]))
    return [t, Spacer(1, 2 * mm), tp, Spacer(1, 3 * mm), tf, Spacer(1, 3 * mm),
            _graficas(s, orden, n, ancho_util, 52 * mm + extra)]


def _graficas(s, orden, n, ancho_util, alto=72 * mm) -> Drawing:
    """Barras apiladas por mes (P1…Pn) y barras del consumo anual por periodo con su %."""
    w = ancho_util
    d = Drawing(w, alto)
    datos = [[s.consumo_mes(m)[p] / 1000 for m in orden] for p in range(n)]
    g = VerticalBarChart()
    g.x, g.y, g.width, g.height = 14 * mm, 14 * mm, w * 0.58 - 14 * mm, alto - 26 * mm
    g.data = datos
    g.categoryAxis.categoryNames = [ap.MESES[m - 1][:3] for m in orden]
    g.categoryAxis.labels.fontSize = 7
    g.valueAxis.labels.fontSize = 7
    g.categoryAxis.labels.fontName = g.valueAxis.labels.fontName = "Helvetica"
    g.valueAxis.valueMin = 0
    g.valueAxis.valueMax = max(sum(datos[p][i] for p in range(n)) for i in range(12)) * 1.08 or 1
    g.valueAxis.labelTextFormat = lambda v: f"{v:,.0f}".replace(",", ".")
    g.categoryAxis.style = "stacked"
    g.barWidth, g.groupSpacing = 6, 6
    for p in range(n):
        g.bars[p].fillColor = colors.HexColor(COL_PERIODOS[p])
        g.bars[p].strokeColor = None
    d.add(g)
    d.add(String(g.x, alto - 8 * mm, "Consumo mensual por periodo (MWh)", fontName="Helvetica-Bold",
                 fontSize=9, fillColor=C["azul"]))
    ley = Legend()
    ley.x, ley.y = g.x + 2 * mm, 5 * mm
    ley.alignment, ley.columnMaximum, ley.fontSize, ley.fontName = "right", 1, 7, "Helvetica"
    ley.dx, ley.dy, ley.deltax = 7, 7, 34
    ley.colorNamePairs = [(colors.HexColor(COL_PERIODOS[p]), f"P{p + 1}") for p in range(n)]
    d.add(ley)
    tot = [sum(s.consumo_mes(m)[p] for m in range(1, 13)) / 1000 for p in range(n)]
    g2 = VerticalBarChart()
    g2.x, g2.y, g2.width, g2.height = w * 0.66, 14 * mm, w * 0.32, alto - 26 * mm
    g2.data = [tot]
    g2.categoryAxis.categoryNames = [f"P{p + 1}" for p in range(n)]
    g2.categoryAxis.labels.fontSize = 7
    g2.valueAxis.labels.fontSize = 7
    g2.categoryAxis.labels.fontName = g2.valueAxis.labels.fontName = "Helvetica"
    g2.valueAxis.valueMin = 0
    g2.valueAxis.labelTextFormat = lambda v: f"{v:,.0f}".replace(",", ".")
    g2.bars[0].fillColor = C["azul"]
    g2.bars[0].strokeColor = None
    for p in range(n):
        g2.bars[(0, p)].fillColor = colors.HexColor(COL_PERIODOS[p])
    suma = sum(tot) or 1
    g2.barLabelFormat = lambda v: _pct(v / suma)
    g2.barLabels.fontSize, g2.barLabels.nudge, g2.barLabels.fontName = 7, 6, "Helvetica"
    g2.valueAxis.valueMax = max(tot) * 1.15 if max(tot) else 1
    d.add(g2)
    d.add(String(g2.x, alto - 8 * mm, "Consumo anual por periodo (MWh y %)", fontName="Helvetica-Bold",
                 fontSize=9, fillColor=C["azul"]))
    return d


# ------------------------------------------------------------------- condiciones
def _condiciones(e: mo.Estudio, res: dict, ancho_util: float) -> Table:
    s_h = _st("ch", 7.5, True, colors.white)
    s_v = _st("cv", 7, alin=0)
    cab = ["Oferta", "Contacto", "Validez y duración", "Pago y garantías", "Volumen y cierres",
           "Rescisión anticipada", "GdO y otros"]
    datos = [[Paragraph(x, s_h) for x in cab]]
    filas = resultados.comparativa(res)
    todas = ([(res["actual"].oferta, None)] if res.get("actual") else []) + \
        [(r.oferta, f) for r, f in zip(res["ofertas"], filas)]
    for o, f in todas:
        c = o.condiciones or {}
        nombre = "Contrato actual<br/>" + escape(o.comercializadora or "") if f is None \
            else cabecera_oferta(f)
        juntar = lambda *xs: "<br/>".join(escape(str(x)) for x in xs if x)
        datos.append([Paragraph(nombre, _st("cn", 7, True, alin=0)),
                      Paragraph(juntar(o.comercializadora, o.gestor, o.email_gestor), s_v),
                      Paragraph(juntar(c.get("validez"), f"{o.duracion} meses" if f else ""), s_v),
                      Paragraph(juntar(c.get("pago"), c.get("garantias")), s_v),
                      Paragraph(juntar(c.get("banda"), c.get("cierres")), s_v),
                      Paragraph(juntar(c.get("rescision")), s_v),
                      Paragraph(juntar(c.get("gdo"), c.get("otros")), s_v)])
    pesos = [0.13, 0.14, 0.15, 0.14, 0.15, 0.13, 0.16]
    t = Table(datos, colWidths=[ancho_util * p for p in pesos], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), C["azul"]),
                           # fondos transparentes: se ve la marca de agua del papel debajo
                           ("BACKGROUND", (0, 1), (0, -1), _transp(C["lteal"], 0.75)),
                           ("ROWBACKGROUNDS", (1, 1), (-1, -1),
                            [None, _transp(colors.HexColor("#DCE6F1"), 0.35)]),
                           ("GRID", (0, 0), (-1, -1), 0.4, C["linea"]),
                           ("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]))
    return t


# ------------------------------------------------------------------- papel con marca de agua
HUECO = 3 * mm          # separación entre los datos y el logo o el «GE&PE» del papel


def papel(orientacion: str, n: int):
    """Ruta del papel con marca de agua («vertical»/«horizontal», 1 o 2) o None si no existe.
    1: logo grande (para páginas con menos datos); 2: logo pequeño (más sitio)."""
    ruta = config.RECURSOS / "assets" / f"papel_{orientacion}_{n}.pdf"
    return ruta if ruta.exists() else None


@functools.lru_cache(maxsize=None)
def zona_libre(ruta) -> dict:
    """Márgenes (pt, en el tamaño del papel) que deja libres el papel: debajo del logo, encima
    del pie con el nombre y la dirección de GE&PE y a la derecha del texto del margen. El
    «GE&PE» gris es una marca de agua clara: los datos pueden ir encima. Se leen del propio PDF
    (el pie, dibujando el papel y buscando la primera fila oscura del cuarto inferior)."""
    import fitz
    import numpy as np
    with fitz.open(str(ruta)) as d:
        p = d[0]
        ancho, alto = p.rect.width, p.rect.height
        arriba = max((p.get_image_bbox(i).y1 for i in p.get_images(full=True)), default=0)
        izq = max((b[2] for b in p.get_text("blocks") if b[4].strip() and b[0] < 40
                   and b[3] - b[1] > alto / 4), default=0)
        desde = alto * 0.75
        zona = fitz.Rect(izq + 2, desde, ancho, alto)
        pix = p.get_pixmap(matrix=fitz.Matrix(2, 2), clip=zona, colorspace=fitz.csGRAY)
        gris = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width)
        filas = np.where((gris < 160).any(axis=1))[0]
        pie = desde + filas[0] / 2 if len(filas) else alto
    return {"ancho": ancho, "alto": alto, "arriba": arriba + HUECO, "abajo": float(alto - pie) + HUECO,
            "izq": max(MARGEN, izq + HUECO), "der": MARGEN}


def _marco(tam, ruta):
    """(x, y, ancho, alto) del marco de datos de una página de tamaño tam con ese papel
    (el papel se escala al tamaño de la página: A4 → A3 × √2)."""
    if ruta is None:
        return MARGEN, 13 * mm, tam[0] - 2 * MARGEN, tam[1] - 13 * mm - 8 * mm
    z = zona_libre(ruta)
    f = tam[0] / z["ancho"]
    return (z["izq"] * f, z["abajo"] * f, tam[0] - (z["izq"] + z["der"]) * f,
            tam[1] - (z["arriba"] + z["abajo"]) * f)


def _igualar(contenido: bytes, tam) -> bytes:
    """Amplía las hojas más pequeñas (p. ej. la ficha A4 vertical) al tamaño de las demás (A3),
    en su misma orientación, para que todo el informe se imprima en el mismo papel."""
    import fitz
    grande = max(tam)
    doc = fitz.open(stream=contenido, filetype="pdf")
    if all(round(max(p.rect.width, p.rect.height)) == round(grande) for p in doc):
        doc.close()
        return contenido
    pequeno = min(tam)
    nuevo = fitz.open()
    for i, p in enumerate(doc):
        if round(max(p.rect.width, p.rect.height)) == round(grande):
            nuevo.insert_pdf(doc, from_page=i, to_page=i)
        else:
            vertical = p.rect.height > p.rect.width
            hoja = nuevo.new_page(width=pequeno if vertical else grande,
                                  height=grande if vertical else pequeno)
            hoja.show_pdf_page(hoja.rect, doc, i)
    salida = nuevo.tobytes(garbage=3, deflate=True)
    nuevo.close()
    doc.close()
    return salida


def _poner_papel(contenido: bytes, papeles: list) -> bytes:
    """Pone debajo de cada página su papel (ruta o None)."""
    import fitz
    doc = fitz.open(stream=contenido, filetype="pdf")
    abiertos = {}
    for pagina, ruta in zip(doc, papeles):
        if ruta is None:
            continue
        if ruta not in abiertos:
            abiertos[ruta] = fitz.open(str(ruta))
        pagina.show_pdf_page(pagina.rect, abiertos[ruta], 0, overlay=False)
    salida = doc.tobytes(garbage=3, deflate=True)
    for d in abiertos.values():
        d.close()
    doc.close()
    return salida


# ------------------------------------------------------------------- gráficas multipunto
class _Cartera:
    """Suma de los suministros: se dibuja con las mismas gráficas que un monopunto."""

    def __init__(self, suministros):
        self.s = suministros

    def consumo_mes(self, mes):
        return [sum(s.consumo_mes(mes)[p] for s in self.s) for p in range(6)]


def _eje(g, maximo):
    g.categoryAxis.labels.fontSize = g.valueAxis.labels.fontSize = 7
    g.categoryAxis.labels.fontName = g.valueAxis.labels.fontName = "Helvetica"
    g.valueAxis.valueMin = 0
    g.valueAxis.valueMax = maximo or 1
    g.valueAxis.labelTextFormat = lambda v: f"{v:,.0f}".replace(",", ".")


def _grafica_tarifas(activos, ancho, alto) -> Drawing:
    """Consumo anual por tarifa (MWh), apilado por periodo, con el % de cada tarifa."""
    tarifas = sorted({s.tarifa for s in activos}, key=lambda t: -sum(
        s.consumo_anual() for s in activos if s.tarifa == t))
    datos = [[sum(sum(s.consumo_mes(m)[p] for m in range(1, 13)) for s in activos
                  if s.tarifa == t) / 1000 for t in tarifas] for p in range(6)]
    tot = [sum(datos[p][i] for p in range(6)) for i in range(len(tarifas))]
    suma = sum(tot) or 1
    d = Drawing(ancho, alto)
    g = VerticalBarChart()
    g.x, g.y, g.width, g.height = 14 * mm, 16 * mm, ancho - 18 * mm, alto - 30 * mm
    g.data = datos
    g.categoryAxis.categoryNames = [f"{t} ({_pct(v / suma)})" for t, v in zip(tarifas, tot)]
    g.categoryAxis.style = "stacked"
    _eje(g, max(tot, default=0) * 1.1)
    g.barWidth = 18
    g.groupSpacing = 10
    for p in range(6):
        g.bars[p].fillColor = colors.HexColor(COL_PERIODOS[p])
        g.bars[p].strokeColor = None
    d.add(g)
    paso = g.width / max(1, len(tarifas))
    for i, v in enumerate(tot):
        d.add(String(g.x + paso * (i + 0.5), g.y + g.height * v / (g.valueAxis.valueMax or 1) + 2,
                     _num(v, 0) + " MWh", fontName="Helvetica-Bold", fontSize=7,
                     textAnchor="middle"))
    d.add(String(g.x, alto - 8 * mm, "Consumo anual por tarifa (MWh, apilado por periodo)",
                 fontName="Helvetica-Bold", fontSize=9, fillColor=C["azul"]))
    ley = Legend()
    ley.x, ley.y = g.x + 2 * mm, 5 * mm
    ley.alignment, ley.columnMaximum, ley.fontSize, ley.fontName = "right", 1, 7, "Helvetica"
    ley.dx, ley.dy, ley.deltax = 7, 7, 30
    ley.colorNamePairs = [(colors.HexColor(COL_PERIODOS[p]), f"P{p + 1}") for p in range(6)
                          if any(datos[p])]
    d.add(ley)
    return d


def _grafica_principales(activos, ancho, alto, n=10) -> Drawing:
    """Los n suministros de más consumo (MWh) y el resto agrupado en «Otros»."""
    orden = sorted(activos, key=lambda s: -s.consumo_anual())
    filas = [((s.alias or s.cups[-8:])[:28], s.consumo_anual() / 1000) for s in orden[:n]]
    resto = sum(s.consumo_anual() for s in orden[n:]) / 1000
    if resto > 0:
        filas.append((f"Otros ({len(orden) - n})", resto))
    suma = sum(v for _, v in filas) or 1
    filas = filas[::-1]                                   # el mayor arriba
    d = Drawing(ancho, alto)
    g = HorizontalBarChart()
    g.x, g.y, g.width, g.height = 48 * mm, 4 * mm, ancho - 48 * mm - 26 * mm, alto - 16 * mm
    g.data = [[v for _, v in filas]]
    g.categoryAxis.categoryNames = [x for x, _ in filas]
    _eje(g, max((v for _, v in filas), default=0) * 1.05)
    g.valueAxis.visible = False
    g.bars[0].fillColor = C["azul"]
    g.bars[0].strokeColor = None
    if resto > 0:
        g.bars[(0, 0)].fillColor = C["gris"]
    g.barLabelFormat = lambda v: f"{_num(v, 0)} MWh · {_pct(v / suma)}"
    g.barLabels.fontSize, g.barLabels.fontName = 7, "Helvetica"
    g.barLabels.boxAnchor, g.barLabels.dx = "w", 3
    d.add(g)
    d.add(String(0, alto - 8 * mm, f"Suministros de mayor consumo anual (MWh y % del total)",
                 fontName="Helvetica-Bold", fontSize=9, fillColor=C["azul"]))
    return d


def _graficas_cartera(e: mo.Estudio, activos, ancho_util, extra: float = 0) -> list:
    """Gráficas de la ficha multipunto: mensual de la cartera por periodo y anual por periodo;
    consumo por tarifa; suministros de mayor consumo."""
    orden = [mo.sumar_meses(e.inicio, k).month for k in range(12)] if e.inicio else list(range(1, 13))
    n = max(mo.n_periodos(s.tarifa) for s in activos)
    varias = len({s.tarifa for s in activos}) > 1        # con una sola tarifa no aporta
    alto = max(62 * mm, (min(11, len(activos)) * 6 + 18) * mm) + extra

    def fila(piezas, pesos):
        anchos = [ancho_util * p / sum(pesos) for p in pesos]
        t = Table([[hacer(a - 4 * mm) for hacer, a in zip(piezas, anchos)]], colWidths=anchos)
        t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return KeepTogether([t])

    mensual = lambda a: _graficas(_Cartera(activos), orden, n, a, 72 * mm + extra)
    tarifas = lambda a: _grafica_tarifas(activos, a, alto)
    principales = lambda a: _grafica_principales(activos, a, alto)
    if ancho_util >= 340 * mm:
        # hoja ancha (A3): mensual + anual por periodo, tarifas y principales en una fila
        return [fila([mensual, tarifas, principales] if varias else [mensual, principales],
                     [0.5, 0.18, 0.32] if varias else [0.6, 0.4])]
    segunda = fila([tarifas, principales], [0.42, 0.58]) if varias else         fila([principales], [1])
    return [KeepTogether([mensual(ancho_util)]), Spacer(1, 4 * mm), segunda]


# ------------------------------------------------------------------- documento
def _tam_informe(n_cols: int):
    """(página, ancho de cada columna de oferta): A4 apaisado si caben con letra legible."""
    for tam in (landscape(A4), landscape(A3)):
        util = _marco(tam, papel("horizontal", 2))[2] - COL_ETIQ - COL_MIN
        ancho = util / max(1, n_cols)
        if ancho >= COL_OF_MIN:
            return tam, min(COL_OF_MAX, ancho)
    tam = landscape(A3)
    return tam, (_marco(tam, papel("horizontal", 2))[2] - COL_ETIQ - COL_MIN) / max(1, n_cols)


def _repartir(flujos, ancho, alto, desde=4, max_hueco=12 * mm):
    """Si el bloque cabe en una hoja, reparte el sitio que sobra entre sus separaciones (las de
    después de la barra del título) para que ocupe toda la hoja y no quede todo arriba. Se deja
    un 10 % del sobrante abajo."""
    sobra = alto - _alto(flujos, ancho)
    huecos = [f for f in flujos[desde:] if isinstance(f, Spacer)]
    if sobra <= 0 or not huecos:
        return flujos
    extra = min(max_hueco, sobra * 0.9 / len(huecos))
    for f in huecos:
        f.height += extra
    return flujos


def _alto(flujos, ancho) -> float:
    """Alto que ocupan los flujos en una sola columna de ese ancho (sin cortar)."""
    total = 0.0
    for f in flujos:
        if isinstance(f, KeepTogether):
            total += _alto(f._content, ancho)
        else:
            total += f.wrap(ancho, 100_000)[1] + f.getSpaceBefore() + f.getSpaceAfter()
    return total


def pdf(e: mo.Estudio, res: dict, ref: str = "", tecnico: dict = None, ronda: int = None,
        mejora: dict = None) -> bytes:
    """Ficha del suministro (o relación de suministros con gráficas) → resumen de ofertas →
    condiciones. Cada página lleva debajo el papel con la marca de agua de GE&PE. ronda: informe
    de una ronda (va en el título del resumen y de las condiciones); mejora: fila «Mejora frente
    a la ronda anterior» (resultados.mejora_ronda_anterior)."""
    tecnico = tecnico or {}
    sufijo = f" · {ronda}ª RONDA" if ronda else ""
    filas = resultados.comparativa(res)
    n_cols = max(1, max(sum(f["tipo"] == "Fijo" for f in filas),
                        sum(f["tipo"] != "Fijo" for f in filas)))
    tam_inf, ancho_of = _tam_informe(n_cols)
    activos = [s for s in e.suministros if s.activo]
    mono = len(activos) == 1
    hoy = e.fecha or dt.date.today()
    if mono:
        tam_fic = portrait(A4)
    else:
        # la relación cabe en A4 apaisado con dirección de al menos 40 mm
        necesario = (8 + 30 + 40 + 38 + 12 + 72 + 96 + 21) * mm
        tam_fic = landscape(A4) if _marco(landscape(A4), papel("horizontal", 2))[2] >= necesario \
            else landscape(A3)
    tam_cond = landscape(A4)

    def orientacion(tam):
        return "horizontal" if tam[0] > tam[1] else "vertical"

    # ---- contenido de cada bloque
    def ficha(ancho, con_logo, extra=0):
        if mono:
            h = _cabecera(e, ancho, "FICHA DEL PUNTO DE SUMINISTRO", ref, con_logo)
            h += [_realizado(tecnico, hoy, ancho), Spacer(1, 3 * mm)]
            return h + _ficha_monopunto(e, activos[0], ancho, extra)
        h = _cabecera(e, ancho, "RELACIÓN DE PUNTOS DE SUMINISTRO", ref, con_logo)
        h += [_realizado(tecnico, hoy, min(ancho, 230 * mm)), Spacer(1, 4 * mm),
              _relacion(e, ancho), Spacer(1, 6 * mm)]
        return h + _graficas_cartera(e, activos, ancho, extra)

    def informe(ancho, con_logo, extra=0):
        return _pagina_informe(e, res, ref, n_cols, ancho_of, ancho, con_logo, sufijo, mejora)

    def condiciones(ancho, con_logo, extra=0):
        return _cabecera(e, ancho, "CONDICIONES DE LAS OFERTAS" + sufijo, ref, con_logo) + \
            [_condiciones(e, res, ancho)]

    # ---- tamaño y papel de cada bloque: el primero en el que los datos caben en una hoja
    # (A4 con el papel 1, A4 con el 2, A3 con el 1, A3 con el 2); si no caben en ninguna, A4
    # con el papel de más sitio y los datos siguen en las hojas siguientes
    bloques = []
    for nombre, tam, hacer in (("ficha", tam_fic, ficha), ("informe", tam_inf, informe),
                               ("condiciones", tam_cond, condiciones)):
        o = orientacion(tam)
        rutas = [r for r in (papel(o, 1), papel(o, 2)) if r] or [None]
        tamanos = [tam] + ([landscape(A3)] if tam == landscape(A4) else [])
        elegido = (tam, rutas[-1])
        for t in tamanos:
            for r in rutas:
                m = _marco(t, r)
                if _alto(hacer(m[2], r is None), m[2]) <= m[3]:
                    elegido = (t, r)
                    break
            else:
                continue
            break
        bloques.append((nombre, elegido[0], hacer, elegido[1]))
    # todas las hojas del mismo tamaño: si alguna necesita A3, las apaisadas se maquetan en A3
    # y la ficha vertical de un monopunto se amplía a A3 vertical al final (_igualar)
    if any(t == landscape(A3) for _, t, _, _ in bloques):
        for k, (nombre, t, hacer, r) in enumerate(bloques):
            if t == landscape(A4):
                rutas = [x for x in (papel("horizontal", 1), papel("horizontal", 2)) if x] or [None]
                elegido = next((x for x in rutas if _alto(hacer(_marco(landscape(A3), x)[2], x is None),
                                                         _marco(landscape(A3), x)[2])
                                <= _marco(landscape(A3), x)[3]), rutas[-1])
                bloques[k] = (nombre, landscape(A3), hacer, elegido)

    plantillas, h, papel_de = [], [], {}
    for i, (nombre, tam, hacer, ruta) in enumerate(bloques):
        x, y, ancho, alto = _marco(tam, ruta)
        papel_de[nombre] = ruta
        plantillas.append(PageTemplate(
            id=nombre, pagesize=tam, onPage=lambda c, d, r=ruta: _pie(c, d, r),
            frames=[Frame(x, y, ancho, alto, id=nombre, leftPadding=0, rightPadding=0,
                          topPadding=0, bottomPadding=0)]))
        if i:
            h += [NextPageTemplate(nombre), PageBreak()]
        flujos = hacer(ancho, ruta is None)
        sobra = alto - _alto(flujos, ancho)
        if nombre == "ficha" and sobra > 10 * mm:
            # en la ficha el sitio que sobra agranda primero las gráficas (hasta 1,8 veces)
            flujos = hacer(ancho, ruta is None, extra=min(sobra * 0.6, 55 * mm))
        h += _repartir(flujos, ancho, alto)

    paginas = []

    def _pie(canvas, doc, ruta):
        paginas.append(ruta)
        canvas.saveState()
        ancho = canvas._pagesize[0]
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(C["gris"])
        if ruta is None:
            canvas.setStrokeColor(C["linea"])
            canvas.line(MARGEN, 10 * mm, ancho - MARGEN, 10 * mm)
            texto = " · ".join(x for x in ("GE&PE Ingeniería · Estudio de ofertas de electricidad",
                                            e.cliente, f"{hoy:%d/%m/%Y}") if x)
            canvas.drawString(MARGEN, 6 * mm, texto)
        canvas.drawRightString(ancho - MARGEN, 6 * mm, f"Página {doc.page}")
        canvas.restoreState()

    salida = io.BytesIO()
    doc = BaseDocTemplate(salida, pagesize=bloques[0][1], pageTemplates=plantillas,
                          title="Estudio de ofertas de electricidad",
                          author=tecnico.get("nombre") or "GE&PE")
    doc.build(h)
    final = _poner_papel(salida.getvalue(), paginas)
    return _igualar(final, max((t for _, t, _, _ in bloques), key=lambda t: t[0] * t[1]))
