# -*- coding: utf-8 -*-
"""
Iberdrola Clientes (grandes cuentas).

Maqueta observada (el texto ordenado mezcla columnas: se busca cada dato por su patron):
  Periodo de facturación dd/mm/aaaa - dd/mm/aaaa
  Número de factura 0000...            Fecha de emisión de factura dd de mes de aaaa
  Peaje de acceso a la red (ATR): 6.1TD
  Energía consumida
  P3 10.000 kWh x 0,100000 €/kWh     1.000,00 €
  Restricciones técnicas y POS
  30.000 kWh x 0,020000 €/kWh          600,00 €
  Total 30000 kWh hasta dd/mm/aaaa
La fecha inicial del periodo es la de la lectura anterior (a las 24 h): el consumo
facturado empieza el dia siguiente.
"""

import datetime as dt
import re

from .base import (RE_FECHA_LETRA, LineaSSAA, comprobar_lineas, fecha_es, lector_generico,
                   numero_es)

NUM = r"-?\d[\d.]*(?:,\d+)?"

RE_PERIODO = re.compile(r"Periodo\s+de\s+facturaci[oó]n\s+(\d{2}/\d{2}/\d{4})\s*-\s*"
                        r"(\d{2}/\d{2}/\d{4})", re.I)
RE_NUMERO = re.compile(r"N[uú]mero\s+de\s+factura\s+(\d{10,})", re.I)
RE_EMISION = re.compile(r"Fecha\s+de\s+emisi[oó]n\s+de\s+factura\s+" + RE_FECHA_LETRA, re.I)
RE_TARIFA = re.compile(r"\(ATR\)\s*:\s*([236]\.[0-4]TD(?:VE)?)", re.I)
RE_ENERGIA = re.compile(r"\bP([1-6])\s+(" + NUM + r")\s*kWh\s*x\s*" + NUM + r"\s*€/kWh", re.I)
RE_TOTAL = re.compile(r"\bTotal\s+(\d+)\s*kWh\s+hasta", re.I)
RE_SSAA = re.compile(r"Restricciones\s+t[eé]cnicas\s+y\s+POS\s+(" + NUM + r")\s*kWh\s*x\s*("
                     + NUM + r")\s*€/kWh\s+(" + NUM + r")\s*€", re.I)


def leer(texto):
    f = lector_generico(texto)
    f.avisos = []

    m = RE_NUMERO.search(texto)
    if m:
        f.numero = m.group(1)
    m = RE_EMISION.search(texto)
    if m:
        f.fecha_emision = fecha_es(m.group(1))
    m = RE_PERIODO.search(texto)
    if m:
        f.inicio = fecha_es(m.group(1)) + dt.timedelta(days=1)
        f.fin = fecha_es(m.group(2))
    m = RE_TARIFA.search(texto)
    if m:
        f.tarifa = m.group(1).upper()

    m = RE_TOTAL.search(texto)
    energia = {}
    for e in RE_ENERGIA.finditer(texto):
        energia.setdefault(e.group(1), numero_es(e.group(2)))
    if m:
        f.consumo_kwh = float(m.group(1))
    elif energia:
        f.consumo_kwh = round(sum(energia.values()), 3)

    f.lineas_ssaa = []
    for m in RE_SSAA.finditer(texto):
        l = LineaSSAA(concepto="Restricciones técnicas y POS", inicio=f.inicio, fin=f.fin,
                      texto=" ".join(m.group(0).split()))
        l.kwh, l.precio, l.importe = (numero_es(m.group(i)) for i in (1, 2, 3))
        f.lineas_ssaa.append(l)
    if not f.lineas_ssaa:
        f.avisos.append("No se encuentra la línea «Restricciones técnicas y POS».")
    return comprobar_lineas(f)
