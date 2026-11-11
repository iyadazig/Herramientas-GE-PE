# -*- coding: utf-8 -*-
"""
EXCEL DE CALCULO DEL ANALISIS DE OFERTAS (fórmulas vivas, fáciles de seguir)
===========================================================================
    excel(e, res, ref, tecnico) -> bytes
Hojas (como la plantilla original de análisis de ofertas):
    Resumen                 mismo formato que la hoja «Informe» de la plantilla, enlazado a las
                            hojas de cada oferta
    Info suministro         datos de los CUPS (potencias, IEE, alquiler) y perfil de consumo
                            mensual P1-P6
    Parámetros utilizados   impuestos y costes regulados por año natural, peajes y cargos por
                            tarifa, pérdidas, mercado previsto por mes de contrato (OMIP y OMIE)
                            y apuntamientos
    Contrato actual y una hoja por oferta:
        condiciones de precio (celdas amarillas) y precios por año de contrato
        resultado de la oferta (total, media anual, €/MWh y comparación con el contrato actual)
        TOTAL de la cartera mes a mes y TOTAL por tarifa
        precios unitarios de cada mes y periodo (€/kWh) por grupo de tarifa y sistema
        cálculo de cada CUPS: consumo × precio unitario, potencia, K, IEE y servicios
Todas las cuentas son fórmulas de Excel: si se cambia un parámetro o un precio se recalcula.
Las fórmulas dependen de la modalidad de cada oferta (fijo, Pass-Pool, Pass-Through) y del modo
de los SSAA: se escriben las de la modalidad elegida. Mismo resultado que la aplicación.
Sin Streamlit.
"""

import datetime as dt
import io
import re

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

import config
import datos_regulados as dr
from analisis_ofertas import calculo, resultados
from analisis_ofertas import modelo as mo
from analisis_ofertas.informe_pdf import cabecera_oferta, nombre_corto

LOGO = config.RECURSOS / "assets" / "logo_geype.png"
PAR, INFO, RES = "Parámetros utilizados", "Info suministro", "Resumen"
QP = "'Parámetros utilizados'!"
QI = "'Info suministro'!"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]

F = {k: PatternFill("solid", fgColor=v) for k, v in {
    "navy": "002060", "teal": "31869B", "lteal": "DCE6F1", "orange": "FCD5B4", "act": "FFFFCC",
    "minh": "215967", "min": "F2F2F2", "azul": "1F4E79", "azul_c": "DDEBF7", "in": "FFF2CC",
    "tot": "E2EFDA", "gris": "F2F2F2", "verde": "C6EFCE"}.items()}
FINO = Side(style="thin", color="BFBFBF")
CAJA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
EUR, EUR2, KWH, P6, PCT = '#,##0 €;[Red]-#,##0 €;"-"', '#,##0.00;[Red]-#,##0.00;"-"', \
    '#,##0;[Red]-#,##0;"-"', '0.000000;[Red]-0.000000;"-"', '0.00%;[Red]-0.00%;"-"'


# ------------------------------------------------------------------- utilidades
def _put(ws, ref, v=None, fmt=None, b=False, color="000000", fill=None, size=10, al=None,
         wrap=False, borde=False, it=False):
    c = ws[ref]
    c.value = v
    c.font = Font(name="Calibri", bold=b, italic=it, color=color, size=size)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = F[fill]
    if al or wrap:
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=wrap)
    if borde:
        c.border = CAJA
    return c


def _hdr(ws, ref, v, fill="azul", color="FFFFFF", size=9):
    return _put(ws, ref, v, b=True, color=color, fill=fill, size=size, al="center", wrap=True,
                borde=True)


def _titulo(ws, fila, texto, c1=1, c2=12, fill="navy"):
    ws.merge_cells(start_row=fila, start_column=c1, end_row=fila, end_column=c2)
    _put(ws, f"{CL(c1)}{fila}", texto, b=True, color="FFFFFF", fill=fill, size=11)
    ws.row_dimensions[fila].height = 20


def _nombre_hoja(nombre, usados):
    n = re.sub(r"[\[\]\*\?/\\:]", "-", nombre)[:31]
    base, i = n, 2
    while n.lower() in usados:
        n = f"{base[:28]} {i}"
        i += 1
    usados.add(n.lower())
    return n


def _q(hoja):
    return "'" + hoja.replace("'", "''") + "'!"


# ------------------------------------------------------------------- Parámetros utilizados
class Parametros:
    """Escribe la hoja y guarda dónde está cada dato (para las fórmulas de las ofertas)."""

    def __init__(self, wb, e: mo.Estudio, ctx: calculo.Contexto, H: int):
        self.ws = ws = wb.create_sheet(PAR)
        self.e, self.ctx, self.H = e, ctx, H
        ws.sheet_view.showGridLines = False
        _titulo(ws, 1, "PARÁMETROS UTILIZADOS EN EL CÁLCULO", 1, 26)
        _put(ws, "A2", "Valores congelados al guardar el análisis. Las celdas amarillas se pueden "
                       "cambiar: todas las hojas se recalculan.", it=True, color="7F7F7F", size=9)
        meses = [mo.sumar_meses(e.inicio, k) for k in range(H)]
        self.anios = sorted({f.year for f in meses})
        tarifas = sorted({s.tarifa for s in e.suministros if s.activo})
        sistemas = sorted({(s.tarifa, s.sistema) for s in e.suministros if s.activo})
        r = 4
        # --- 1. generales por año natural
        _titulo(ws, r, "1. IMPUESTOS Y COSTES REGULADOS POR AÑO NATURAL", 1, 10, "teal")
        cab = ["Año", "IEE (tipo)", "Impuesto municipal (%)", "FNEE (€/MWh)", "Pago OM (€/MWh)",
               "Pago OS (€/MWh)", "Datos de"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        ws.row_dimensions[r + 1].height = 40
        self.gen = {}
        for j, a in enumerate(self.anios):
            f = r + 2 + j
            g = ctx.regulados(a)["generales"]
            _put(ws, f"A{f}", a, b=True, al="center", borde=True)
            for i, (k, fmt) in enumerate((("iee", "0.0000000%"), ("impuesto_municipal", PCT),
                                         ("fnee", "0.0000"), ("om", "0.0000"), ("os", "0.0000"))):
                _put(ws, f"{CL(2 + i)}{f}", g[k], fmt, fill="in", borde=True)
            copia = ctx.regulados(a).get("copia_de")
            _put(ws, f"G{f}", f"copia de {copia}" if copia not in (None, a) else str(a), size=9,
                 borde=True)
            self.gen[a] = f
        self.r_gen = (r + 2, r + 1 + len(self.anios))
        r = self.r_gen[1] + 2
        # --- 2. crecimiento y % OMIP
        _titulo(ws, r, "2. ESTIMACIÓN DEL MERCADO Y CRECIMIENTO DEL CONSUMO", 1, 10, "teal")
        _put(ws, f"A{r + 1}", "Ponderación OMIP en el precio previsto (resto = media OMIE)", b=True)
        _put(ws, f"F{r + 1}", e.pct_omip, PCT, fill="in", borde=True)
        self.pct = f"{QP}$F${r + 1}"
        _put(ws, f"A{r + 2}", "Sesión OMIP utilizada", b=True)
        _put(ws, f"F{r + 2}", dt.date.fromisoformat(e.sesion_omip) if e.sesion_omip else None,
             "dd/mm/yyyy", borde=True)
        n_anios = (H + 11) // 12
        crec = calculo._crecimiento(e, max(1, n_anios))
        _put(ws, f"A{r + 3}", "Factor de consumo por año de contrato (crecimiento)", b=True)
        self.crec = {}
        for y in range(n_anios):
            _hdr(ws, f"{CL(6 + y)}{r + 3}", f"Año {y + 1}", "lteal", "000000")
            _put(ws, f"{CL(6 + y)}{r + 4}", crec[y], "0.0000", fill="in", borde=True)
            self.crec[y + 1] = f"{QP}${CL(6 + y)}${r + 4}"
        r += 6
        # --- 3. mercado por mes de contrato
        _titulo(ws, r, "3. MERCADO PREVISTO POR MES DE CONTRATO (€/MWh)", 1, 14, "teal")
        cab = ["Mes contrato", "Fecha", "Año", "Mes", "OMIP aplicado", "Media OMIE (años)",
               "Precio previsto", "SSAA previstos", "Ph − diario (Canarias)",
               "Factor de carga pico (8-20 h L-V)", "Horas", "Horas pico"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        ws.row_dimensions[r + 1].height = 40
        self.m0 = r + 2
        for k, fch in enumerate(meses):
            f = self.m0 + k
            fila_m = e.mercado.get(f"{fch.year}-{fch.month:02d}") or {}
            omip = e.omip_manual.get(f"{fch.year}-{fch.month:02d}")
            omip = fila_m.get("omip") if omip is None else omip
            _put(ws, f"A{f}", k + 1, al="center", borde=True)
            _put(ws, f"B{f}", fch, "mmm-yy", borde=True)
            _put(ws, f"C{f}", fch.year, al="center", borde=True)
            _put(ws, f"D{f}", fch.month, al="center", borde=True)
            _put(ws, f"E{f}", omip if omip else None, EUR2, fill="in", borde=True)
            _put(ws, f"F{f}", fila_m.get("omie_media"), EUR2, fill="in", borde=True)
            _put(ws, f"G{f}", f'=IF(N(E{f})>0,IF(F{f}="",E{f},{self.pct}*E{f}+(1-{self.pct})*F{f}),N(F{f}))',
                 EUR2, b=True, borde=True)
            _put(ws, f"H{f}", ctx.ssaa(fch.year, fch.month), EUR2, fill="in", borde=True)
            _put(ws, f"I{f}", ctx.ph_dif(fch.year, fch.month), EUR2, fill="in", borde=True)
            _put(ws, f"J{f}", float(e.pico.get(fch.month) or 1.0), "0.0000", fill="in", borde=True)
            _put(ws, f"K{f}", calculo._horas(fch.year, fch.month), "#,##0", borde=True)
            _put(ws, f"L{f}", calculo._horas_pico(fch.year, fch.month), "#,##0", borde=True)
        ws["J" + str(self.m0)].comment = Comment(
            "Factor de carga pico = precio medio del mercado diario de 8 a 20 h de lunes a viernes "
            "÷ precio medio del mes (12 últimos meses). Solo se usa para liquidar cierres en carga "
            "pico.", "GE&PE")
        if not any(o.indice_ph for o in e.ofertas + ([e.actual] if e.actual else [])):
            ws.column_dimensions["I"].hidden = True          # Ph de Canarias: no se usa
        ws["G" + str(self.m0)].comment = Comment(
            "Precio previsto = % OMIP × OMIP + (1 − % OMIP) × media OMIE del mismo mes de los "
            "últimos años; sin OMIP, la media de OMIE.", "GE&PE")
        r = self.m0 + H + 1
        # --- 4. peajes, cargos, pagos por capacidad y SRAD
        _titulo(ws, r, "4. PEAJES + CARGOS, PAGOS POR CAPACIDAD Y SRAD POR AÑO NATURAL Y TARIFA", 1, 27, "teal")
        _hdr(ws, f"A{r + 1}", "Clave")
        _hdr(ws, f"B{r + 1}", "Año")
        _hdr(ws, f"C{r + 1}", "Tarifa")
        for i, nom in enumerate(("Potencia (€/kW·año)", "Energía (€/kWh)", "Pagos capacidad (€/kWh)",
                                 "SRAD (€/kWh)")):
            for p in range(6):
                _hdr(ws, f"{CL(4 + 6 * i + p)}{r + 1}", f"{nom} P{p + 1}")
        ws.row_dimensions[r + 1].height = 52
        self.t0 = r + 2
        f = self.t0
        for a in self.anios:
            for t in tarifas:
                tm = dr.terminos(ctx.regulados(a), t)
                _put(ws, f"A{f}", f'=B{f}&"|"&C{f}', borde=True, size=9)
                _put(ws, f"B{f}", a, al="center", borde=True)
                _put(ws, f"C{f}", t, al="center", borde=True)
                for i, k in enumerate(("potencia", "energia", "pagos_capacidad", "srad")):
                    vals = tm[k] or [0.0] * 6
                    for p in range(6):
                        _put(ws, f"{CL(4 + 6 * i + p)}{f}", vals[p], P6, fill="in", borde=True)
                f += 1
        self.t1 = f - 1
        r = f + 1
        # --- 5. pérdidas
        _titulo(ws, r, "5. COEFICIENTES DE PÉRDIDAS POR AÑO, TARIFA Y SISTEMA", 1, 10, "teal")
        for i, h in enumerate(["Clave", "Año", "Tarifa", "Sistema"] + [f"P{p}" for p in range(1, 7)]):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        self.l0 = r + 2
        f = self.l0
        for a in self.anios:
            for t, si in sistemas:
                perd = dr.perdidas(ctx.regulados(a), t, si)
                _put(ws, f"A{f}", f'=B{f}&"|"&C{f}&"|"&D{f}', borde=True, size=9)
                _put(ws, f"B{f}", a, al="center", borde=True)
                _put(ws, f"C{f}", t, al="center", borde=True)
                _put(ws, f"D{f}", si, al="center", borde=True)
                for p in range(6):
                    _put(ws, f"{CL(5 + p)}{f}", perd[p], "0.0%", fill="in", borde=True)
                f += 1
        self.l1 = f - 1
        r = f + 1
        # --- 6. apuntamientos estimados
        _titulo(ws, r, "6. APUNTAMIENTOS ESTIMADOS POR SISTEMA, CALENDARIO Y MES NATURAL "
                       "(precio del periodo / precio medio del mes; vacío = 1)", 1, 17, "teal")
        cab = ["Clave", "Sistema|calendario", "Mes"] + [f"OMIE P{p}" for p in range(1, 7)] + \
              [f"SSAA P{p}" for p in range(1, 7)] + ["Dato de"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        self.a0 = r + 2
        f = self.a0
        claves = sorted(set(e.ap_omie) | set(e.ap_ssaa))
        for clave in claves:
            om, ss = e.ap_omie.get(clave) or {}, e.ap_ssaa.get(clave) or e.ap_omie.get(clave) or {}
            for m in range(1, 13):
                _put(ws, f"A{f}", f'=B{f}&"|"&C{f}', borde=True, size=9)
                _put(ws, f"B{f}", clave, borde=True)
                _put(ws, f"C{f}", m, al="center", borde=True)
                vo, vs = calculo._factores(om.get(m) or om.get(str(m))), \
                    calculo._factores(ss.get(m) or ss.get(str(m)))
                for p in range(6):
                    _put(ws, f"{CL(4 + p)}{f}", vo[p], "0.0000", fill="in", borde=True)
                    _put(ws, f"{CL(10 + p)}{f}", vs[p], "0.0000", fill="in", borde=True)
                _put(ws, f"P{f}", ((e.ap_origen.get(clave) or {}).get(m) or ""), size=8, borde=True)
                f += 1
        self.a1 = f - 1
        ws.column_dimensions["A"].width = 16
        for c in range(2, 28):
            ws.column_dimensions[CL(c)].width = 11.5
        ws.freeze_panes = "B4"

    # referencias para las fórmulas
    def rng(self, col, desde, hasta):
        return f"{QP}${col}${desde}:${col}${hasta}"

    def termino(self, clave_ref, bloque, p):
        """bloque 0 potencia, 1 energía, 2 capacidad, 3 SRAD; clave_ref: celda con año|tarifa."""
        col = CL(4 + 6 * bloque + p)
        return (f"INDEX({self.rng(col, self.t0, self.t1)},MATCH({clave_ref},"
                f"{self.rng('A', self.t0, self.t1)},0))")

    def perdida(self, clave_ref, p):
        return (f"INDEX({self.rng(CL(5 + p), self.l0, self.l1)},MATCH({clave_ref},"
                f"{self.rng('A', self.l0, self.l1)},0))")

    def apunt(self, clave_ref, p, ssaa=False):
        col = CL((10 if ssaa else 4) + p)
        return (f"IFERROR(INDEX({self.rng(col, self.a0, self.a1)},MATCH({clave_ref},"
                f"{self.rng('A', self.a0, self.a1)},0)),1)")

    def general(self, anio_ref, col):
        """col: B IEE, C impuesto municipal, D FNEE, E OM, F OS."""
        return (f"INDEX({self.rng(col, *self.r_gen)},MATCH({anio_ref},"
                f"{self.rng('A', *self.r_gen)},0))")

    def mes(self, k, col):
        return f"{QP}${col}${self.m0 + k - 1}"


# ------------------------------------------------------------------- Info suministro
class Info:
    def __init__(self, wb, e: mo.Estudio, tecnico: dict, ref: str):
        self.ws = ws = wb.create_sheet(INFO)
        ws.sheet_view.showGridLines = False
        activos = [s for s in e.suministros if s.activo]
        _titulo(ws, 1, "INFORMACIÓN DEL SUMINISTRO", 1, 17)
        datos = [("Estudio realizado por", tecnico.get("nombre") or ""),
                 ("Email del técnico", tecnico.get("email") or ""),
                 ("Fecha", e.fecha or dt.date.today()), ("Commodity", "ELECTRICIDAD"),
                 ("Titular del contrato", e.cliente), ("CIF", e.cif),
                 ("Inicio de suministro", e.inicio), ("Referencia", ref)]
        for i, (k, v) in enumerate(datos):
            _put(ws, f"A{3 + i}", k, b=True, color="FFFFFF", fill="azul", borde=True)
            ws.merge_cells(f"B{3 + i}:E{3 + i}")
            _put(ws, f"B{3 + i}", v, "dd/mm/yyyy" if isinstance(v, dt.date) else None,
                 fill="azul_c", borde=True)
        r = 12
        _titulo(ws, r, "PUNTOS DE SUMINISTRO", 1, 17, "teal")
        cab = ["Nº", "CUPS", "Denominación", "Dirección", "Tarifa", "Sistema"] + \
              [f"Pot. P{p} (kW)" for p in range(1, 7)] + ["Reducción IEE", "Alquiler (€/mes)",
                                                          "Consumo anual (kWh)"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        ws.row_dimensions[r + 1].height = 32
        self.fila = {}
        self.s0 = r + 2
        for n, s in enumerate(activos):
            f = self.s0 + n
            self.fila[s.cups] = f
            for i, v in enumerate([n + 1, s.cups, s.alias, s.direccion, s.tarifa, s.sistema]):
                _put(ws, f"{CL(1 + i)}{f}", v, borde=True, size=9)
            for p in range(6):
                _put(ws, f"{CL(7 + p)}{f}", float(s.potencias[p] or 0), "#,##0.###", fill="in", borde=True)
            _put(ws, f"M{f}", float(s.reduccion_iee or 0), "0%", fill="in", borde=True)
            _put(ws, f"N{f}", float(s.alquiler or 0), EUR2, fill="in", borde=True)
        r = self.s0 + len(activos) + 1
        _titulo(ws, r, "PERFIL DE CONSUMO (kWh) — año tipo por mes natural y periodo", 1, 17, "teal")
        for i, h in enumerate(["Clave", "CUPS", "Denominación", "Mes nº", "Mes", "P1", "P2", "P3", "P4",
                               "P5", "P6", "Total"]):
            _hdr(ws, f"{CL(1 + i)}{r + 1}", h)
        self.c0 = r + 2
        f = self.c0
        for s in activos:
            ini = f
            for m in range(1, 13):
                _put(ws, f"A{f}", f'=B{f}&"|"&D{f}', size=8, borde=True)
                _put(ws, f"B{f}", s.cups, size=9, borde=True)
                _put(ws, f"C{f}", s.alias, size=9, borde=True)
                _put(ws, f"D{f}", m, al="center", borde=True)
                _put(ws, f"E{f}", MESES[m - 1], borde=True)
                for p, v in enumerate(s.consumo_mes(m)):
                    _put(ws, f"{CL(6 + p)}{f}", v, KWH, fill="in", borde=True)
                _put(ws, f"L{f}", f"=SUM(F{f}:K{f})", KWH, b=True, borde=True)
                f += 1
            _put(ws, f"E{f}", "Total", b=True, fill="tot", borde=True)
            for p in range(7):
                c = CL(6 + p)
                _put(ws, f"{c}{f}", f"=SUM({c}{ini}:{c}{f - 1})", KWH, b=True, fill="tot", borde=True)
            ws[f"O{self.fila[s.cups]}"].value = f"=L{f}"
            ws[f"O{self.fila[s.cups]}"].number_format = KWH
            ws[f"O{self.fila[s.cups]}"].border = CAJA
            f += 2
        self.c1 = f
        anchos = {"A": 12, "B": 24, "C": 18, "D": 30, "E": 11, "F": 11}
        for c, w in anchos.items():
            ws.column_dimensions[c].width = w
        for c in range(7, 16):
            ws.column_dimensions[CL(c)].width = 11

    def consumo(self, cups_ref, mes_ref, p):
        col = CL(6 + p)
        return (f"INDEX({QI}${col}${self.c0}:${col}${self.c1},MATCH({cups_ref}&\"|\"&{mes_ref},"
                f"{QI}$A${self.c0}:$A${self.c1},0))")


# ------------------------------------------------------------------- hoja de oferta
CONCEPTOS = ["Potencia", "Energía / mercado", "Peajes y cargos energía", "SSAA",
             "Otros (FNEE, GdO, fee)", "K / imp. municipal", "Impuesto eléctrico",
             "Servicios y alquiler"]
# precio fijo: (título de la columna, clave, bloque de precios unitarios; None = se calcula aparte)
COLS_FIJO = [("Potencia", "POT", "POTENCIA €/kW·mes"), ("Energía / mercado", "ENER", "ENERGÍA €/kWh"),
             ("Peajes y cargos energía", "PEAJ", "PEAJES €/kWh"), ("SSAA", "SSAA", "SSAA €/kWh"),
             ("Otros (FNEE, GdO, fee)", "OTROS", "OTROS €/kWh"), ("K / imp. municipal", "K", None),
             ("Impuesto eléctrico", "IEE", None), ("Servicios y alquiler", "SERV", None)]
# indexados: una columna por concepto de la fórmula (los términos de energía, ya elevados a
# pérdidas en Pass-Through y multiplicados por el Ci en Pass-Pool)
BLOQUES_IDX = [("TÉRMINO B €/kWh", "Energía fijo"), ("OMIE €/kWh", "Precio OMIE"),
               ("CAPACIDAD €/kWh", "Pagos por capacidad"), ("OM €/kWh", "Pago O.M."),
               ("OS €/kWh", "Pago O.S."), ("DESVÍOS €/kWh", "Desvíos / CC"), ("SRAD €/kWh", "SRAD")]
BLOQUES_OTROS = [("FNEE €/kWh", "FNEE"), ("CG €/kWh", "CG (fee)"), ("GDO €/kWh", "GdO")]
# claves de los conceptos sujetos al K / impuesto municipal (energía y otros; el SSAA si se elige)
K_BASE = {"ENER", "OTROS", "OMIE", "TÉRMINO B", "CAPACIDAD", "OM", "OS", "DESVÍOS", "SRAD", "FNEE",
          "CG", "GDO"}


def columnas_oferta(o: mo.Oferta, hay_b: bool) -> list:
    """Columnas de € por mes de la hoja de una oferta: (título, clave, bloque unitario)."""
    if not o.indexada:
        return COLS_FIJO
    usa = {"TÉRMINO B": hay_b, "CAPACIDAD": o.pagos_capacidad, "OM": o.om, "OS": o.os,
           "DESVÍOS": bool(o.desvios), "SRAD": o.srad, "FNEE": o.fnee != mo.FNEE_INCLUIDO,
           "CG": bool(o.fee), "GDO": bool(o.gdo)}
    clave = lambda bloque: bloque.split(" €")[0]
    cols = [("ATR potencia", "POT", "POTENCIA €/kW·mes")]
    if not o.atr_incluido:
        cols.append(("ATR energía", "PEAJ", "PEAJES €/kWh"))
    for bloque, titulo in BLOQUES_IDX:
        k = clave(bloque)
        if k == "OMIE":
            cols.append(("Precio OMIE" + (" y cierres" if o.cierres_tipo != mo.SIN_CIERRES else ""),
                         k, bloque))
            if o.ssaa_modo != mo.SSAA_INCLUIDOS:
                cols.append(("Costes de ajuste (SSAA)", "SSAA", "SSAA €/kWh"))
        elif usa[k]:
            cols.append((titulo, k, bloque))
    for bloque, titulo in BLOQUES_OTROS:
        if usa[clave(bloque)]:
            cols.append((titulo, clave(bloque), bloque))
    if o.k:
        cols.append(("Impuesto municipal", "K", None))
    return cols + [("IEE", "IEE", None), ("Servicios y alquiler", "SERV", None)]


def _explicacion(o: mo.Oferta) -> str:
    if o.modalidad == mo.FIJO:
        return ("Energía = precio fijo del periodo / unidad × (1 − descuento)"
                + ("" if o.atr_incluido else " + peajes y cargos de energía"))
    regul = "pagos por capacidad + OM + OS + desvíos + SRAD (los que añada la oferta)"
    if o.modalidad == mo.PASS_POOL:
        ap = "media mensual (sin apuntamiento)" if o.ci_media_mensual else "precio previsto × apuntamiento"
        return (f"Energía = Ci × ({ap} + {regul}) + término B. Cada concepto va en su columna, "
                f"ya multiplicado por el Ci")
    return (f"Energía = (precio previsto × apuntamiento + {regul}) × (1 + pérdidas) + término B. "
            f"Cada concepto va en su columna, ya elevado a pérdidas")


class HojaOferta:
    def __init__(self, wb, nombre_hoja, e, o, r: calculo.ResultadoOferta, par: Parametros,
                 info: Info, ctx, es_actual=False):
        self.ws = ws = wb.create_sheet(nombre_hoja)
        self.nombre = nombre_hoja
        ws.sheet_view.showGridLines = False
        self.e, self.o, self.r, self.par, self.info = e, o, r, par, info
        D = r.n_meses
        self.D = D
        _titulo(ws, 1, ("CONTRATO ACTUAL: " if es_actual else "OFERTA: ") + o.nombre, 1, 22)
        _put(ws, "A2", " · ".join(x for x in (o.comercializadora, o.producto, o.modalidad,
                                               f"{D} meses" + (" (proyectado al horizonte del "
                                                               "estudio)" if es_actual else ""))
                                   if x), b=True)
        _put(ws, "A3", _explicacion(o), it=True, color="1F4E79")
        _put(ws, "A4", "Los importes se calculan mes a mes con el consumo de cada CUPS × el precio "
                       "unitario de cada periodo (tablas de abajo).", it=True, color="7F7F7F", size=9)
        # --- condiciones (celdas amarillas)
        _titulo(ws, 6, "CONDICIONES DE PRECIO", 1, 3, "teal")
        cond = [("Modalidad", o.modalidad), ("Unidad del precio fijo", o.unidad),
                ("Divisor de unidad", mo.UNIDADES.get(o.unidad, 1)),
                ("¿Precio con peajes y cargos?", "Sí" if o.atr_incluido else "No"),
                ("Término de potencia", o.potencia), ("Descuento (fijo)", o.descuento),
                ("Base del término B", o.base_b),
                ("Añade pagos por capacidad", 1 if o.pagos_capacidad else 0),
                ("Añade pago OS", 1 if o.os else 0), ("Añade pago OM", 1 if o.om else 0),
                ("Añade SRAD", 1 if o.srad else 0), ("Desvíos (€/MWh, en barras)", o.desvios),
                ("Fee / coste de gestión (€/MWh)", o.fee), ("Base del fee", o.base_fee),
                ("K / impuesto municipal", o.k), ("K también sobre SSAA", 1 if o.k_ssaa else 0),
                ("GdO (€/MWh)", o.gdo), ("FNEE", o.fnee),
                ("Otros fijos (€/CUPS·mes)", o.otros_fijos), ("Modo SSAA", o.ssaa_modo),
                ("Banda mínima (€/MWh)", o.banda_min), ("Banda máxima (€/MWh)", o.banda_max),
                ("Techo (€/MWh)", o.techo), ("SSAA con perfil", 1 if o.ssaa_perfil else 0),
                ("Pérdidas en SSAA", o.ssaa_perdidas), ("% pérdidas fijo SSAA", o.ssaa_pct_perdidas),
                ("Factor SSAA 1", o.ssaa_f1), ("Factor SSAA 2", o.ssaa_f2),
                ("Ci sobre la media mensual (PP)", 1 if o.ci_media_mensual else 0),
                ("Cierres", o.cierres_tipo), ("Coste por click (€/MWh)", o.coste_click)]
        self.c = {}
        for i, (k, v) in enumerate(cond):
            f = 7 + i
            _put(ws, f"A{f}", k, borde=True, size=9)
            ws.merge_cells(f"A{f}:B{f}")
            _put(ws, f"C{f}", v, PCT if k in ("Descuento (fijo)", "K / impuesto municipal",
                                             "% pérdidas fijo SSAA") else None,
                 fill="in", borde=True, size=9, al="center")
            self.c[k] = f"$C${f}"
        fila = 7 + len(cond) + 1
        # --- resultado (arriba a la derecha)
        self.res_fila = {}
        _titulo(ws, 6, "RESULTADO", 5, 8, "minh")
        # --- precios por año de contrato y grupo
        self.lista, _ = calculo.incluidos(e, o)
        grupos = calculo.grupos_de(e, o)
        ny = (D + 11) // 12
        self.precios = {}
        _titulo(ws, fila, "PRECIOS DE LA OFERTA POR GRUPO Y AÑO DE CONTRATO", 1, 27, "teal")
        cab = ["Grupo", "Año contrato"] + [f"Energía P{p}" for p in range(1, 7)] + \
              [f"Término B P{p} (€/MWh)" for p in range(1, 7)] + [f"Ci P{p}" for p in range(1, 7)] + \
              [f"Potencia propia P{p} (€/kW·año)" for p in range(1, 7)]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{fila + 1}", h)
        ws.row_dimensions[fila + 1].height = 40
        f = fila + 2
        for g in sorted({gn for _, gn in self.lista}):
            tar = grupos[g][0]
            ini = f
            for y in range(1, ny + 1):
                pr = o.precios_grupo(g, y, tar)
                _put(ws, f"A{f}", g, borde=True, size=9)
                _put(ws, f"B{f}", y, al="center", borde=True)
                for i, (attr, defecto) in enumerate((("energia", 0.0), ("termino_b", 0.0),
                                                     ("ci", 1.0), ("potencia", 0.0))):
                    vals = [defecto if v is None else float(v) for v in mo.seis(getattr(pr, attr))]
                    for p in range(6):
                        _put(ws, f"{CL(3 + 6 * i + p)}{f}", vals[p], "0.000000", fill="in", borde=True)
                f += 1
            self.precios[g] = (ini, f - 1)
        fila = f + 1
        hay_b = any(v for g in self.precios for y in range(1, ny + 1)
                    for v in mo.seis(o.precios_grupo(g, y, grupos[g][0]).termino_b) if v)
        self.cols = columnas_oferta(o, hay_b)
        # --- tablas de precios unitarios por grupo (y por CUPS con apuntamiento propio)
        self.tablas = {}
        bloques = []
        for s, g in self.lista:
            tar, sis = grupos[g]
            propio = (s.apuntamiento == mo.CLIENTE and e.ap_cliente.get(s.cups) and o.indexada)
            clave = f"{g}|{s.cups}" if propio else g
            if clave not in self.tablas and clave not in [b[0] for b in bloques]:
                bloques.append((clave, g, tar, sis, s.cups if propio else None))
        self.fila_tot = fila
        n_tar = len({grupos[g][0] for _, g in self.lista})
        fila = fila + D + 4 + ((n_tar * (D + 4)) if len(self.lista) > 1 else 0)
        for clave, g, tar, sis, cups in bloques:
            fila = self._tabla_precios(fila, clave, g, tar, sis, cups, ctx)
        # --- bloques por CUPS
        self.cups = {}
        for s, g in self.lista:
            clave = f"{g}|{s.cups}" if f"{g}|{s.cups}" in self.tablas else g
            fila = self._bloque_cups(fila, s, clave)
        # --- totales (arriba)
        self._totales(grupos)
        ws.column_dimensions["A"].width = 13
        ws.column_dimensions["B"].width = 11
        for col in range(3, 130):
            ws.column_dimensions[CL(col)].width = 11.5
        ws.column_dimensions["C"].width = 21
        ws.freeze_panes = "C6"

    # celdas de las condiciones
    def cnd(self, k):
        return self.c[k]

    def _tabla_precios(self, fila, clave, g, tar, sis, cups, ctx):
        ws, o, par, e, D = self.ws, self.o, self.par, self.e, self.D
        cal = mo.calendario_de(tar)
        ph = calculo.usa_ph(o, sis)
        meses = [(f.year, f.month) for f in (mo.sumar_meses(o.inicio or e.inicio, k) for k in range(D))]
        ap_om, ap_ss = calculo.apuntamientos(e, o, sis, cal, meses, cups, ph=ph)
        clave_ap = mo.clave_ap(mo.SISTEMA_PH if ph and e.ap_omie.get(mo.clave_ap(mo.SISTEMA_PH, cal))
                               else sis, cal)
        est_om = [calculo._factores((e.ap_omie.get(clave_ap) or {}).get(m)) for _, m in meses]
        est_ss = [calculo._factores((e.ap_ssaa.get(mo.clave_ap(sis, cal)) or
                                     e.ap_omie.get(mo.clave_ap(sis, cal)) or {}).get(m)) for _, m in meses]
        formula_om = all(abs(a - b) < 1e-12 for fa, fb in zip(ap_om.tolist(), est_om) for a, b in zip(fa, fb))
        formula_ss = all(abs(a - b) < 1e-12 for fa, fb in zip(ap_ss.tolist(), est_ss) for a, b in zip(fa, fb))
        tit = f"PRECIOS UNITARIOS — grupo {g}" + (f" (CUPS {cups}, apuntamiento de su curva)" if cups else "")
        _titulo(ws, fila, tit, 1, 66 + (54 if o.indexada else 0), "azul")
        idx = o.indexada
        grupos_cab = [("Mes", 1), ("Fecha", 1), ("Año natural", 1), ("Año contrato", 1),
                      ("Mes natural", 1), ("Previsto €/MWh", 1), ("Mercado €/MWh", 1),
                      ("Cierres €/MWh", 1),
                      ("SSAA €/MWh", 1), ("FNEE €/MWh", 1), ("Apuntamiento OMIE", 6),
                      ("Apuntamiento SSAA", 6), ("Pérdidas", 6)] + \
                     ([(b, 6) for b, _ in BLOQUES_IDX] if idx else [("Regulados €/MWh", 6)]) + \
                     [("ENERGÍA €/kWh", 6), ("PEAJES €/kWh", 6), ("SSAA €/kWh", 6)] + \
                     ([(b, 6) for b, _ in BLOQUES_OTROS] if idx else []) + \
                     [("OTROS €/kWh", 6), ("POTENCIA €/kW·mes", 6), ("Clave tarifa", 1),
                      ("Clave pérdidas", 1), ("Clave apuntamiento", 1)]
        col, cols = 1, {}
        for nom, n in grupos_cab:
            cols[nom] = col
            if n == 1:
                _hdr(ws, f"{CL(col)}{fila + 1}", nom)
                ws.merge_cells(start_row=fila + 1, start_column=col, end_row=fila + 2, end_column=col)
            else:
                ws.merge_cells(start_row=fila + 1, start_column=col, end_row=fila + 1, end_column=col + n - 1)
                _hdr(ws, f"{CL(col)}{fila + 1}", nom)
                for p in range(n):
                    _hdr(ws, f"{CL(col + p)}{fila + 2}", f"P{p + 1}", "lteal", "000000")
            col += n
        ws.row_dimensions[fila + 1].height = 30
        L = {k: CL(v) for k, v in cols.items()}
        P = lambda nom, p: CL(cols[nom] + p)
        pr0, pr1 = self.precios[g]
        c = self.cnd
        f0 = fila + 3
        for k in range(1, D + 1):
            f = f0 + k - 1
            a, m = meses[k - 1]
            _put(ws, f"{L['Mes']}{f}", k, al="center", borde=True)
            _put(ws, f"{L['Fecha']}{f}", f"={par.mes(k, 'B')}", "mmm-yy", borde=True)
            _put(ws, f"{L['Año natural']}{f}", f"={par.mes(k, 'C')}", borde=True)
            _put(ws, f"{L['Año contrato']}{f}", f"=ROUNDUP({L['Mes']}{f}/12,0)", borde=True)
            _put(ws, f"{L['Mes natural']}{f}", f"={par.mes(k, 'D')}", borde=True)
            _put(ws, f"{L['Previsto €/MWh']}{f}", f"={par.mes(k, 'G')}", EUR2, borde=True)
            prev = f"{L['Previsto €/MWh']}{f}"
            if o.cierres_tipo == mo.SIN_CIERRES:
                merc = f"={prev}"
            else:
                d = self.r.cierres[k - 1]
                merc = f"={prev}*(1-{d['pct']})+{d['precio']}*{d['pct']}"
            if ph:                                   # Ph demanda de Canarias: + (Ph − diario)
                merc += f"+{par.mes(k, 'I')}"
            _put(ws, f"{L['Mercado €/MWh']}{f}", merc if o.indexada else None, EUR2, b=True, borde=True)
            cierre_v = 0.0
            if o.cierres_tipo != mo.SIN_CIERRES:
                d = self.r.cierres[k - 1]
                cierre_v = float(d["liquidacion"] + d["clicks"])
            _put(ws, f"{L['Cierres €/MWh']}{f}", cierre_v, EUR2, borde=True,
                 fill="in" if o.cierres_tipo != mo.SIN_CIERRES else None)
            _put(ws, f"{L['SSAA €/MWh']}{f}", f"={par.mes(k, 'H')}", EUR2, borde=True)
            _put(ws, f"{L['FNEE €/MWh']}{f}", "=" + par.general(f"{L['Año natural']}{f}", "D"), "0.0000", borde=True)
            _put(ws, f"{L['Clave tarifa']}{f}", f'={L["Año natural"]}{f}&"|{tar}"', size=8, borde=True)
            _put(ws, f"{L['Clave pérdidas']}{f}", f'={L["Año natural"]}{f}&"|{tar}|{sis}"', size=8, borde=True)
            _put(ws, f"{L['Clave apuntamiento']}{f}", f'="{clave_ap}|"&{L["Mes natural"]}{f}', size=8, borde=True)
            kt, kp, ka = (f"${L['Clave tarifa']}{f}", f"${L['Clave pérdidas']}{f}",
                          f"${L['Clave apuntamiento']}{f}")
            ka_ss = f'"{mo.clave_ap(sis, cal)}|"&${L["Mes natural"]}{f}'
            anio_c = f"${L['Año contrato']}{f}"
            fila_pr = f"MIN({anio_c},{pr1 - pr0 + 1})"
            for p in range(6):
                aom, ass, per = P("Apuntamiento OMIE", p), P("Apuntamiento SSAA", p), P("Pérdidas", p)
                _put(ws, f"{aom}{f}", ("=" + par.apunt(ka, p)) if formula_om else float(ap_om[k - 1][p]),
                     "0.0000", borde=True, fill=None if formula_om else "in")
                _put(ws, f"{ass}{f}", ("=" + par.apunt(ka_ss, p, True)) if formula_ss else float(ap_ss[k - 1][p]),
                     "0.0000", borde=True, fill=None if formula_ss else "in")
                _put(ws, f"{per}{f}", "=" + par.perdida(kp, p), "0.0%", borde=True)
                yr = f"{L['Año natural']}{f}"
                # regulados de mercado (€/MWh): cada uno por separado en los indexados
                regs = {"CAPACIDAD": f"{c('Añade pagos por capacidad')}*{par.termino(kt, 2, p)}*1000",
                        "SRAD": f"{c('Añade SRAD')}*{par.termino(kt, 3, p)}*1000",
                        "OS": f"{c('Añade pago OS')}*{par.general(yr, 'F')}",
                        "OM": f"{c('Añade pago OM')}*{par.general(yr, 'E')}",
                        "DESVÍOS": c('Desvíos (€/MWh, en barras)')}
                if not idx:
                    reg = P("Regulados €/MWh", p)
                    _put(ws, f"{reg}{f}", "=" + "+".join(regs[k_] for k_ in ("CAPACIDAD", "SRAD", "OS", "OM",
                                                                           "DESVÍOS")), EUR2, borde=True)
                b_cel = f"INDEX(${CL(9 + p)}${pr0}:${CL(9 + p)}${pr1},{fila_pr})"
                ci_cel = f"INDEX(${CL(15 + p)}${pr0}:${CL(15 + p)}${pr1},{fila_pr})"
                en_cel = f"INDEX(${CL(3 + p)}${pr0}:${CL(3 + p)}${pr1},{fila_pr})"
                b_aj = f"{b_cel}*IF({c('Base del término B')}=\"{mo.BARRAS}\",1+{per}{f},1)"
                merc_c = f"{L['Mercado €/MWh']}{f}"
                cie_c = f"{L['Cierres €/MWh']}{f}"
                if o.modalidad == mo.FIJO:
                    ener = f"={en_cel}/{c('Divisor de unidad')}*(1-{c('Descuento (fijo)')})"
                else:
                    # Pass-Pool: × Ci; Pass-Through: elevado a pérdidas
                    if o.modalidad == mo.PASS_POOL:
                        apm = f"IF({c('Ci sobre la media mensual (PP)')}=1,1,{aom}{f})"
                        fac, omie = ci_cel, f"({merc_c}*{apm}+{cie_c})"
                    else:
                        fac, omie = f"(1+{per}{f})", f"({merc_c}*{aom}{f}+{cie_c})"
                    partes = {"OMIE": f"{omie}*{fac}/1000", "TÉRMINO B": f"{b_aj}/1000"}
                    partes.update({k_: f"{v}*{fac}/1000" for k_, v in regs.items()})
                    for bloque, _ in BLOQUES_IDX:
                        _put(ws, f"{P(bloque, p)}{f}", "=" + partes[bloque.split(' €')[0]], P6, borde=True)
                    ener = "=" + "+".join(f"{P(b_, p)}{f}" for b_, _ in BLOQUES_IDX)
                _put(ws, f"{P('ENERGÍA €/kWh', p)}{f}", ener, P6, b=True, borde=True)
                _put(ws, f"{P('PEAJES €/kWh', p)}{f}",
                     f"=IF({c('¿Precio con peajes y cargos?')}=\"Sí\",0,{par.termino(kt, 1, p)})", P6, borde=True)
                s_ = f"{L['SSAA €/MWh']}{f}"
                sp = f"{s_}*IF({c('SSAA con perfil')}=1,{ass}{f},1)"
                modo = o.ssaa_modo
                if modo == mo.SSAA_INCLUIDOS:
                    base = "0"
                elif modo == mo.SSAA_PERIODO:
                    base = f"{s_}*{ass}{f}"
                elif modo == mo.SSAA_PLANO:
                    base = s_
                elif modo == mo.SSAA_BANDA:
                    base = (f"IF({sp}<{c('Banda mínima (€/MWh)')},{sp}-{c('Banda mínima (€/MWh)')},"
                            f"IF({sp}>{c('Banda máxima (€/MWh)')},{sp}-{c('Banda máxima (€/MWh)')},0))")
                else:
                    base = f"MAX(0,{sp}-{c('Techo (€/MWh)')})"
                lf = (f"IF({c('Pérdidas en SSAA')}=\"{mo.PERD_TARIFA}\",1+{per}{f},"
                      f"IF({c('Pérdidas en SSAA')}=\"{mo.PERD_FIJO}\",1+{c('% pérdidas fijo SSAA')},1))")
                _put(ws, f"{P('SSAA €/kWh', p)}{f}",
                     f"={base}*{lf}*{c('Factor SSAA 1')}*{c('Factor SSAA 2')}/1000", P6, borde=True)
                fn = (f"IF({c('FNEE')}=\"{mo.FNEE_INCLUIDO}\",0,{L['FNEE €/MWh']}{f}"
                      f"*IF({c('FNEE')}=\"{mo.BARRAS}\",1+{per}{f},1))")
                fee = f"{c('Fee / coste de gestión (€/MWh)')}*IF({c('Base del fee')}=\"{mo.BARRAS}\",1+{per}{f},1)"
                if idx:
                    for bloque, v in zip([b_ for b_, _ in BLOQUES_OTROS], (fn, fee, c('GdO (€/MWh)'))):
                        _put(ws, f"{P(bloque, p)}{f}", f"={v}/1000", P6, borde=True)
                    otros = "=" + "+".join(f"{P(b_, p)}{f}" for b_, _ in BLOQUES_OTROS)
                else:
                    otros = f"=({fn}+{c('GdO (€/MWh)')}+{fee})/1000"
                _put(ws, f"{P('OTROS €/kWh', p)}{f}", otros, P6, borde=True)
                pot_prop = f"INDEX(${CL(21 + p)}${pr0}:${CL(21 + p)}${pr1},{fila_pr})"
                _put(ws, f"{P('POTENCIA €/kW·mes', p)}{f}",
                     f"=IF({c('Término de potencia')}=\"{mo.POT_PROPIA}\",{pot_prop},{par.termino(kt, 0, p)})/12",
                     P6, borde=True)
        self.tablas[clave] = {"f0": f0, "cols": cols}
        return f0 + D + 2

    def _bloque_cups(self, fila, s, clave):
        ws, D, par, info, c = self.ws, self.D, self.par, self.info, self.cnd
        t = self.tablas[clave]
        P = lambda nom, p: CL(t["cols"][nom] + p)
        tf0 = t["f0"]
        fi = info.fila[s.cups]
        cols = self.cols
        ci = {k_: 10 + i for i, (_, k_, _b) in enumerate(cols)}          # columna de cada concepto
        CT, CB = CL(10 + len(cols)), CL(11 + len(cols))                   # TOTAL € y base del IEE
        X = lambda k_: CL(ci[k_])
        _titulo(ws, fila, f"CUPS {s.cups} · {s.alias or ''} · {s.tarifa} · {s.sistema}", 1,
                max(24, 12 + len(cols)), "azul")
        cab = ["Mes", "Fecha"] + [f"kWh P{p}" for p in range(1, 7)] + ["kWh total"] + \
              [t_ for t_, _, _b in cols] + ["TOTAL €", "Base IEE €"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{fila + 1}", h)
        ws.row_dimensions[fila + 1].height = 32
        f0 = fila + 2
        kw = f"{QI}$G${fi}:$L${fi}"
        cups_ref = f"{QI}$B${fi}"
        red = f"{QI}$M${fi}"
        alq = f"{QI}$N${fi}"
        for k in range(1, D + 1):
            f, ft = f0 + k - 1, tf0 + k - 1
            _put(ws, f"A{f}", k, al="center", borde=True)
            _put(ws, f"B{f}", f"={par.mes(k, 'B')}", "mmm-yy", borde=True)
            mes_ref = par.mes(k, "D")
            for p in range(6):
                _put(ws, f"{CL(3 + p)}{f}",
                     f"={info.consumo(cups_ref, mes_ref, p)}*{par.crec[(k - 1) // 12 + 1]}", KWH, borde=True)
            _put(ws, f"I{f}", f"=SUM(C{f}:H{f})", KWH, b=True, borde=True)
            kwh = f"$C{f}:$H{f}"
            fila_p = lambda nom: f"${P(nom, 0)}${ft}:${P(nom, 5)}${ft}"
            for _t, k_, bloque in cols:
                if bloque:
                    _put(ws, f"{X(k_)}{f}", f"=SUMPRODUCT({kw if k_ == 'POT' else kwh},{fila_p(bloque)})",
                         EUR2, borde=True)
            if "K" in ci:
                base_k = [f"{X(k_)}{f}" for _t, k_, _b in cols if k_ in K_BASE]
                ss = f"+IF({c('K también sobre SSAA')}=1,{X('SSAA')}{f},0)" if "SSAA" in ci else ""
                _put(ws, f"{X('K')}{f}", f"={c('K / impuesto municipal')}*({'+'.join(base_k) or '0'}{ss})",
                     EUR2, borde=True)
            # base del IEE: todos los conceptos salvo el propio IEE y los servicios
            _put(ws, f"{CB}{f}", f"=SUM(J{f}:{CL(ci['IEE'] - 1)}{f})", EUR2, borde=True, color="7F7F7F")
            _put(ws, f"{X('SERV')}{f}", f"={c('Otros fijos (€/CUPS·mes)')}+{alq}", EUR2, borde=True)
            _put(ws, f"{CT}{f}", f"=SUM(J{f}:{X('SERV')}{f})", EUR2, b=True, borde=True)
        # IEE por año de contrato
        fy = f0 + D + 1
        _hdr(ws, f"A{fy}", "Año contrato")
        for i, h in enumerate(["Año natural (1.er mes)", "Base IEE €", "MWh", "Tipo IEE",
                               "Reducción", "IEE del año €"]):
            _hdr(ws, f"{CL(2 + i)}{fy}", h)
        ny = (D + 11) // 12
        for y in range(1, ny + 1):
            f = fy + y
            a1, a2 = f0 + 12 * (y - 1), min(f0 + D - 1, f0 + 12 * y - 1)
            _put(ws, f"A{f}", y, al="center", borde=True)
            _put(ws, f"B{f}", f"={par.mes(12 * (y - 1) + 1, 'C')}", borde=True)
            _put(ws, f"C{f}", f"=SUM({CB}{a1}:{CB}{a2})", EUR2, borde=True)
            _put(ws, f"D{f}", f"=SUM(I{a1}:I{a2})/1000", "#,##0.000", borde=True)
            _put(ws, f"E{f}", "=" + par.general(f"B{f}", "B"), "0.0000000%", borde=True)
            _put(ws, f"F{f}", f"={red}", "0%", borde=True)
            _put(ws, f"G{f}", f"=C{f}*(1-F{f})*E{f}", EUR2, b=True, borde=True)
            for k in range(a1, a2 + 1):
                _put(ws, f"{X('IEE')}{k}", f"=IF($C${f}=0,G{f}/{a2 - a1 + 1},G{f}*{CB}{k}/$C${f})", EUR2,
                     borde=True)
        self.cups[s.cups] = (f0, s.tarifa)
        return fy + ny + 3

    def _totales(self, grupos):
        ws, D = self.ws, self.D
        fila = self.fila_tot
        nc = len(self.cols)
        CT, CM = CL(4 + nc), CL(5 + nc)                  # TOTAL € y €/MWh de la cartera
        _titulo(ws, fila, "TOTAL DE LA CARTERA (€ por mes)", 1, max(19, 6 + nc), "navy")
        cab = ["Mes", "Fecha", "kWh total"] + [t_ for t_, _, _b in self.cols] + ["TOTAL €", "€/MWh"]
        for i, h in enumerate(cab):
            _hdr(ws, f"{CL(1 + i)}{fila + 1}", h)
        ws.row_dimensions[fila + 1].height = 32
        f0 = fila + 2
        src = [8] + list(range(9, 10 + nc))            # columnas I (kWh), conceptos y TOTAL de los CUPS
        for k in range(1, D + 1):
            f = f0 + k - 1
            _put(ws, f"A{f}", k, al="center", borde=True)
            _put(ws, f"B{f}", f"={self.par.mes(k, 'B')}", "mmm-yy", borde=True)
            for j, cs in enumerate(src):
                ref = "+".join(f"{CL(cs + 1)}{b0 + k - 1}" for b0, _ in self.cups.values()) or "0"
                _put(ws, f"{CL(3 + j)}{f}", "=" + ref, KWH if j == 0 else EUR2,
                     b=j == len(src) - 1, borde=True)
            _put(ws, f"{CM}{f}", f"=IF(C{f}=0,\"\",{CT}{f}/C{f}*1000)", EUR2, borde=True)
        ft = f0 + D
        _put(ws, f"B{ft}", "TOTAL", b=True, fill="tot", borde=True)
        for j in range(len(src)):
            col = CL(3 + j)
            _put(ws, f"{col}{ft}", f"=SUM({col}{f0}:{col}{ft - 1})", KWH if j == 0 else EUR,
                 b=True, fill="tot", borde=True)
        _put(ws, f"{CM}{ft}", f"=IF(C{ft}=0,\"\",{CT}{ft}/C{ft}*1000)", EUR2, b=True, fill="tot", borde=True)
        self.tot = {"f0": f0, "ft": ft, "total": CT}
        # total por tarifa
        if len(self.lista) > 1:
            fila = ft + 2
            for tar in sorted({t for _, t in self.cups.values()}):
                bl = [b0 for b0, t in self.cups.values() if t == tar]
                _titulo(ws, fila, f"TOTAL TARIFA {tar} ({len(bl)} CUPS)", 1, 12, "teal")
                for i, h in enumerate(["Mes", "Fecha"] + [f"kWh P{p}" for p in range(1, 7)] +
                                      ["kWh total", "TOTAL €", "€/MWh"]):
                    _hdr(ws, f"{CL(1 + i)}{fila + 1}", h)
                g0 = fila + 2
                for k in range(1, D + 1):
                    f = g0 + k - 1
                    _put(ws, f"A{f}", k, al="center", borde=True)
                    _put(ws, f"B{f}", f"={self.par.mes(k, 'B')}", "mmm-yy", borde=True)
                    for j, cs in enumerate([3, 4, 5, 6, 7, 8, 9, 10 + nc]):
                        _put(ws, f"{CL(3 + j)}{f}", "=" + "+".join(f"{CL(cs)}{b0 + k - 1}" for b0 in bl),
                             KWH if j < 7 else EUR2, borde=True)
                    _put(ws, f"K{f}", f"=IF(I{f}=0,\"\",J{f}/I{f}*1000)", EUR2, borde=True)
                fila = g0 + D + 2
        # resultado
        c0 = 5
        m1 = min(12, D)                          # primer año (anualizado si dura menos)
        f1 = f0 + m1 - 1
        filas = [("Meses del contrato", D, "0"), ("Coste total del contrato (€)", f"={CT}{ft}", EUR),
                 ("Consumo del contrato (MWh)", f"=C{ft}/1000", "#,##0.0"),
                 ("Coste medio anual (€/año)", f"={CT}{ft}*12/{D}", EUR),
                 ("Coste unitario (€/MWh)", f"=IF(C{ft}=0,\"\",{CT}{ft}/C{ft}*1000)", EUR2),
                 ("Coste año 1 (€/año)", f"=SUM({CT}{f0}:{CT}{f1})*12/{m1}", EUR),
                 ("Consumo año 1 (MWh)", f"=SUM(C{f0}:C{f1})/1000", "#,##0.0"),
                 ("Coste unitario año 1 (€/MWh)",
                  f"=IF(SUM(C{f0}:C{f1})=0,\"\",SUM({CT}{f0}:{CT}{f1})/SUM(C{f0}:C{f1})*1000)", EUR2)]
        self._escribir_resultado(filas, 7)
        self.res_fila["_fila_libre"] = 7 + len(filas)

    def _escribir_resultado(self, filas, f_ini):
        ws = self.ws
        for i, (k, v, fmt) in enumerate(filas):
            f = f_ini + i
            ws.merge_cells(start_row=f, start_column=5, end_row=f, end_column=7)
            _put(ws, f"E{f}", k, b=True, borde=True, size=9)
            _put(ws, f"H{f}", v, fmt, b=True, fill="min", borde=True, al="right")
            self.res_fila[k] = f"{_q(self.nombre)}$H${f}"

    def comparar_con_actual(self, actual: "HojaOferta"):
        """Contrato actual en los mismos meses que la oferta (primer año y periodo total) y la
        diferencia y variación de cada uno."""
        f = self.res_fila["_fila_libre"]
        a = actual
        rng_f = f"{_q(a.nombre)}$B${a.tot['f0']}:$B${a.tot['ft'] - 1}"
        rng_t = f"{_q(a.nombre)}${a.tot['total']}${a.tot['f0']}:${a.tot['total']}${a.tot['ft'] - 1}"
        ini = f"$B${self.tot['f0']}"
        fin = f"$B${self.tot['ft'] - 1}"
        m1 = min(12, self.D)
        fin1 = f"$B${self.tot['f0'] + m1 - 1}"
        local = lambda k: "$H$" + self.res_fila[k].split("$H$")[1]       # en la misma hoja
        filas = [("Contrato actual en el primer año de la oferta (€/año)",
                  f"=SUMIFS({rng_t},{rng_f},\">=\"&{ini},{rng_f},\"<=\"&{fin1})*12/{m1}", EUR),
                 ("Diferencia año 1 frente al actual (€/año)", "", EUR),
                 ("Variación año 1 frente al actual", "", PCT),
                 ("Contrato actual en los mismos meses (€/año)",
                  f"=SUMIFS({rng_t},{rng_f},\">=\"&{ini},{rng_f},\"<=\"&{fin})*12/{self.D}", EUR),
                 ("Diferencia frente al actual (€/año)", "", EUR),
                 ("Variación frente al actual", "", PCT)]
        # las diferencias se escriben después: necesitan las celdas de las filas de arriba
        self._escribir_resultado(filas, f)
        a1, act1 = local("Coste año 1 (€/año)"), local("Contrato actual en el primer año de la oferta (€/año)")
        ap, actp = local("Coste medio anual (€/año)"), local("Contrato actual en los mismos meses (€/año)")
        for k, v in (("Diferencia año 1 frente al actual (€/año)", f"={a1}-{act1}"),
                     ("Variación año 1 frente al actual", f"=IF(N({act1})=0,\"\",{a1}/{act1}-1)"),
                     ("Diferencia frente al actual (€/año)", f"={ap}-{actp}"),
                     ("Variación frente al actual", f"=IF(N({actp})=0,\"\",{ap}/{actp}-1)")):
            self.ws[local(k).replace("$", "")] = v


# ------------------------------------------------------------------- Resumen (formato Informe)
def _resumen(wb, e, res, hojas, actual_h, ref):
    ws = wb.create_sheet(RES, 0)
    ws.sheet_view.showGridLines = False
    filas = resultados.comparativa(res)
    n = max(1, max(sum(f["tipo"] == "Fijo" for f in filas), sum(f["tipo"] != "Fijo" for f in filas)))
    MC = CL(3 + n)
    if LOGO.exists():
        img = XLImage(str(LOGO))
        img.height, img.width = 60, int(60 * 700 / 352)
        ws.add_image(img, "B1")
    ws.merge_cells(f"B5:{MC}5")
    _put(ws, "B5", "RESUMEN DEL ESTUDIO DE OFERTAS DE RENOVACIÓN", b=True, color="FFFFFF", fill="navy",
         size=12, al="center")
    ws.row_dimensions[5].height = 26
    ws.merge_cells(f"B7:{MC}7")
    _put(ws, "B7", (e.cliente or "").upper(), b=True, size=13, al="center")
    act_nombre = nombre_corto(actual_h.o.comercializadora) + ("" if actual_h.o.modalidad == mo.FIJO else " index.")
    act_y1 = actual_h.res_fila["Coste año 1 (€/año)"]
    act_mwh = f"={actual_h.res_fila['Coste unitario año 1 (€/MWh)']}"
    fila = 9
    for tipo, titulo in (("Fijo", "Estimación facturación ofertas a precio fijo (€/año)"),
                         ("Indexado", "Estimación facturación ofertas indexadas (€/año)")):
        de_tipo = [(f, hojas[f["nombre"]]) for f in filas if (f["tipo"] == "Fijo") == (tipo == "Fijo")]
        if not de_tipo:
            continue
        t = fila
        ws.merge_cells(start_row=t, start_column=2, end_row=t + 1, end_column=2)
        _put(ws, f"B{t}", f"Contrato actual\n{act_nombre}", b=True, color="FFFFFF", fill="teal", al="center", wrap=True)
        ws.merge_cells(start_row=t, start_column=3, end_row=t, end_column=2 + n)
        _put(ws, f"C{t}", titulo, b=True, color="FFFFFF", fill="teal", size=11, al="center")
        ws.merge_cells(start_row=t, start_column=3 + n, end_row=t + 1, end_column=3 + n)
        _put(ws, f"{MC}{t}", "Mejor oferta", b=True, color="FFFFFF", fill="minh", size=11, al="center")
        ws.row_dimensions[t].height = 24
        ws.row_dimensions[t + 1].height = 48
        # primer año: coste, €/MWh, variación frente al actual en esos meses y ranking
        rv, rk, rp, re_, rn = t + 2, t + 3, t + 4, t + 5, t + 6
        # periodo total (en gris): actual en los mismos meses, oferta, diferencia y ranking
        ga, go, gp, ge, gn = t + 7, t + 8, t + 9, t + 10, t + 11
        EURA, EURM = '#,##0 "€/año"', '#,##0.00 "€/MWh"'
        GRIS = dict(it=True, color="7F7F7F", size=8, al="center")
        _put(ws, f"B{rv}", f"={act_y1}", EURA, fill="act", al="center", borde=True)
        _put(ws, f"B{rk}", act_mwh, EURM, fill="act", al="center", borde=True)
        ws.merge_cells(f"B{rp}:B{re_}")
        _put(ws, f"B{rp}", "Variación sobre contrato actual", b=True, fill="lteal", al="center", wrap=True, borde=True)
        _put(ws, f"B{rn}", "Ranking de oferta", b=True, fill="lteal", al="center", borde=True)
        _put(ws, f"B{ga}", "Contrato actual en el mismo periodo que la oferta analizada (€/año)",
             it=True, color="7F7F7F", size=8, wrap=True)
        _put(ws, f"B{go}", "Oferta analizada en su periodo total (€/año)", it=True, color="7F7F7F",
             size=8, wrap=True)
        ws.merge_cells(f"B{gp}:B{ge}")
        _put(ws, f"B{gp}", "Diferencia en su periodo total", it=True, color="7F7F7F", size=8, wrap=True)
        _put(ws, f"B{gn}", "Ranking en su periodo total", it=True, color="7F7F7F", size=8, wrap=True)
        ws.row_dimensions[ga].height = 34
        ws.row_dimensions[go].height = 24
        rng = lambda r_: f"$C${r_}:${CL(2 + n)}${r_}"
        for i in range(n):
            col = CL(3 + i)
            if i < len(de_tipo):
                fo, h = de_tipo[i]
                R = h.res_fila
                _put(ws, f"{col}{t + 1}", cabecera_oferta(fo).replace("<br/>", "\n"), b=True,
                     fill="orange" if fo["duracion"] > 12 else "lteal", size=9, al="center", wrap=True, borde=True)
                _put(ws, f"{col}{rv}", f"={R['Coste año 1 (€/año)']}", EURA, al="center", borde=True)
                _put(ws, f"{col}{rk}", f"={R['Coste unitario año 1 (€/MWh)']}", EURM, al="center", borde=True)
                _put(ws, f"{col}{rp}", f"={R['Variación año 1 frente al actual']}", PCT, al="center", borde=True)
                _put(ws, f"{col}{re_}", f"={R['Diferencia año 1 frente al actual (€/año)']}", EUR, al="center", borde=True)
                _put(ws, f"{col}{rn}", f"=RANK({col}{rp},{rng(rp)},1)", "0", b=True, al="center", borde=True)
                _put(ws, f"{col}{ga}", f"={R['Contrato actual en los mismos meses (€/año)']}", KWH, **GRIS)
                _put(ws, f"{col}{go}", f"={R['Coste medio anual (€/año)']}", KWH, **GRIS)
                _put(ws, f"{col}{gp}", f"={R['Variación frente al actual']}", PCT, **GRIS)
                _put(ws, f"{col}{ge}", f"={R['Diferencia frente al actual (€/año)']}", EUR, **GRIS)
                _put(ws, f"{col}{gn}", f"=RANK({col}{gp},{rng(gp)},1)", "0", **GRIS)
            else:
                for r_ in (t + 1, rv, rk, rp, re_, rn):
                    _put(ws, f"{col}{r_}", "-", al="center", borde=True, fill="lteal" if r_ == t + 1 else None)
                for r_ in (ga, go, gp, ge, gn):
                    _put(ws, f"{col}{r_}", "-", **GRIS)
        pos = f"MATCH(MIN({rng(rp)}),{rng(rp)},0)"           # mejor variación del primer año
        for r_, fmt in ((rv, EURA), (rk, EURM), (rp, PCT), (re_, EUR)):
            _put(ws, f"{MC}{r_}", f"=INDEX({rng(r_)},{pos})", fmt, b=True, fill="min", al="center", borde=True)
        for r_, fmt in ((ga, KWH), (go, KWH), (gp, PCT), (ge, EUR), (gn, "0")):
            _put(ws, f"{MC}{r_}", f"=INDEX({rng(r_)},{pos})", fmt, **GRIS)
        _put(ws, f"{MC}{rn}", f"=SUBSTITUTE(INDEX({rng(t + 1)},{pos}),CHAR(10),\" \")", b=True, fill="min",
             al="center", wrap=True, borde=True, size=8)
        ws.row_dimensions[rn].height = 30
        fila = gn + 3
    # notas
    _put(ws, f"B{fila}", "Notas:", b=True)
    sesion = dt.date.fromisoformat(e.sesion_omip).strftime("%d/%m/%Y") if e.sesion_omip else "-"
    notas = [f"Precio de mercado de los indexados: {e.pct_omip * 100:.0f} % la cotización de OMIP del {sesion} "
             f"y {100 - e.pct_omip * 100:.0f} % la media del precio OMIE de los últimos 3 años (hoja «{PAR}»).",
             "El coste recoge todos los conceptos de la factura (potencia, energía, peajes y cargos, SSAA, "
             "FNEE, garantías de origen, impuesto municipal e impuesto eléctrico) salvo IVA, excesos de "
             "potencia y energía reactiva.",
             "€/año y €/MWh: coste del primer año de cada oferta y del contrato actual. La variación, el "
             "ranking y la mejor oferta comparan ese primer año con el contrato actual en esos mismos meses.",
             "En gris, el periodo total de cada oferta: su coste medio anual durante toda su duración frente "
             "al contrato actual en esos mismos meses, la diferencia y el ranking por esa diferencia. El "
             "detalle de cada cálculo está en la hoja de cada oferta."]
    for i, nt in enumerate(notas):
        ws.merge_cells(f"B{fila + 1 + i}:{MC}{fila + 1 + i}")
        _put(ws, f"B{fila + 1 + i}", nt, it=True, size=9, wrap=True)
        ws.row_dimensions[fila + 1 + i].height = 28
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    for i in range(n):
        ws.column_dimensions[CL(3 + i)].width = 15.5
    ws.column_dimensions[MC].width = 18
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ------------------------------------------------------------------- libro
def excel(e: mo.Estudio, res: dict, ref: str = "", tecnico: dict = None) -> bytes:
    tecnico = tecnico or {}
    ctx = calculo.Contexto(e)
    wb = Workbook()
    wb.remove(wb.active)
    actual = res.get("actual")
    H = max([r.n_meses for r in res["ofertas"]] + [actual.n_meses if actual else 0] + [12])
    info = Info(wb, e, tecnico, ref)
    par = Parametros(wb, e, ctx, H)
    usados = {RES.lower(), INFO.lower(), PAR.lower()}
    hojas = {}
    actual_h = None
    if actual:
        actual_h = HojaOferta(wb, _nombre_hoja("Contrato actual", usados), e, actual.oferta, actual,
                              par, info, ctx, es_actual=True)
    for r in res["ofertas"]:
        h = HojaOferta(wb, _nombre_hoja(r.oferta.nombre, usados), e, r.oferta, r, par, info, ctx)
        if actual_h:
            h.comparar_con_actual(actual_h)
        hojas[r.oferta.nombre] = h
    if actual_h:
        _resumen(wb, e, res, hojas, actual_h, ref)
    orden = [RES, INFO, PAR] + ([actual_h.nombre] if actual_h else []) + [h.nombre for h in hojas.values()]
    wb._sheets = [wb[n] for n in orden if n in wb.sheetnames]
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()
