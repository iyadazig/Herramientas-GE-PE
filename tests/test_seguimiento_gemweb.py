# -*- coding: utf-8 -*-
"""Estados del cuadro de seguimiento desde las facturas de Gemweb (API simulada)."""

import datetime as dt
import sys
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
from seguimiento import columnas as C
from seguimiento import estados_gemweb as G
from tests.test_seguimiento import CUPS1, CUPS2, CUPS3, BaseDatos


def _xml(filas):
    raiz = ET.Element("raiz")
    for f in filas:
        e = ET.SubElement(raiz, "fila")
        for k, v in f.items():
            ET.SubElement(e, k).text = v
    return raiz


class GemwebFalso:
    def __init__(self, facturas):
        self.facturas = facturas
        self.llamadas = []

    def get_inventory(self, category, **kw):
        self.llamadas.append((category, kw))
        if category == "subministraments":
            return _xml([{"id": "1", "cups": CUPS1 + "0F"}, {"id": "2", "cups": CUPS2},
                         {"id": "3", "cups": CUPS3}])
        return _xml(self.facturas)


def _factura(id_sum, ini, fin, verificada="0", numero="F", **marcas):
    return dict({"id_subministraments": id_sum, "periode_ini": ini, "periode_fin": fin,
                 "verificada": verificada, "num_factura": numero}, **marcas)


class EstadosGemweb(BaseDatos):
    def setUp(self):
        super().setUp()
        self.importar()
        # CUPS1 (mensual): ene Inf enviado, mar Recibida; CUPS2: ene-jun Pdte inf, jul-dic Sin fra
        self.ids = {s["cups"]: s["id"] for s in almacen.listar_sgc_suministros()}

    def estado(self, cups, mes):
        return (almacen.listar_sgc_estados(2026).get((self.ids[cups], 2026, mes)) or {}).get(
            "estado")

    def test_reglas(self):
        self.assertEqual(G.nuevo_estado(None, False), C.RECIBIDA)
        self.assertEqual(G.nuevo_estado(C.SIN_FRA, True), C.PDTE_INF)
        self.assertEqual(G.nuevo_estado(C.RECIBIDA, True), C.PDTE_INF)
        self.assertIsNone(G.nuevo_estado(C.RECIBIDA, False))
        for final in (C.PDTE_INF, C.INF_ENVIADO, C.INF_NO_ENVIADO, C.TEMPORAL, C.CADENCIA):
            self.assertIsNone(G.nuevo_estado(final, True))

    def test_actualizar(self):
        api = GemwebFalso([
            _factura("1", "01/07/2026", "31/07/2026", "0", "F-jul"),
            _factura("1", "01/03/2026", "31/03/2026", "1", "F-mar"),    # antes del inicio
            _factura("2", "01/07/2026", "31/08/2026", "1", "F-bim"),    # bimestral verificada
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-anul", anulada="1"),
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-comp", complementaria="1"),
            _factura("3", "01/08/2026", "31/08/2026", "0", "F-ases")])  # asesoramiento
        cambios = G.actualizar(hoy=dt.date(2026, 10, 26), cliente=api)
        self.assertEqual(api.llamadas[1][1]["search_values"], "2026-07-01")
        self.assertEqual(sorted((c["cups"], c["mes"], c["antes"], c["despues"]) for c in cambios),
                         [(CUPS1, 7, None, C.RECIBIDA), (CUPS2, 7, C.SIN_FRA, C.PDTE_INF),
                          (CUPS2, 8, C.SIN_FRA, C.PDTE_INF)])
        self.assertEqual(self.estado(CUPS1, 7), C.RECIBIDA)
        self.assertEqual(self.estado(CUPS1, 3), C.RECIBIDA)                # sin tocar
        self.assertEqual(self.estado(CUPS2, 9), C.SIN_FRA)                 # anulada y comp.
        self.assertIsNone(self.estado(CUPS3, 8))
        self.assertEqual(len(almacen.listar_cambios_gemweb()), 3)
        self.assertEqual(G.actualizar(hoy=dt.date(2026, 10, 26), cliente=api), [])

    def test_simular_no_guarda(self):
        api = GemwebFalso([_factura("1", "01/07/2026", "31/07/2026")])
        self.assertEqual(len(G.actualizar(simular=True, hoy=dt.date(2026, 10, 26), cliente=api)),
                         1)
        self.assertIsNone(self.estado(CUPS1, 7))
        self.assertEqual(almacen.listar_cambios_gemweb(), [])

    def test_inicio_en_enero_cubre_el_año_anterior(self):
        self.assertEqual(G.desde(dt.date(2027, 1, 15)), dt.date(2026, 10, 1))


if __name__ == "__main__":
    unittest.main()
