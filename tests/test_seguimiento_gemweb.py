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
            _factura("1", "01/03/2026", "31/03/2026", "1", "F-mar"),    # Recibida -> Pdte inf
            _factura("1", "01/01/2026", "31/01/2026", "1", "F-ene"),    # ya Inf enviado
            _factura("2", "01/07/2026", "31/08/2026", "1", "F-bim"),    # bimestral verificada
            _factura("2", "01/07/2026", "31/07/2026", "0", "F-jul2"),   # cuenta la verificada
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-anul", anulada="1"),
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-comp", complementaria="1"),
            _factura("3", "01/08/2026", "31/08/2026", "0", "F-ases"),   # asesoramiento
            _factura("9", "01/08/2026", "31/08/2026", "0", "F-fuera")])  # no esta en Gemweb
        r = G.actualizar(hoy=dt.date(2026, 10, 26), cliente=api)
        self.assertEqual(r["inicio"], dt.date(2026, 1, 1))
        self.assertEqual(api.llamadas[1][1]["search_values"], "2026-01-01")
        self.assertEqual(sorted((c["cups"], c["mes"], c["antes"], c["despues"])
                                for c in r["cambios"]),
                         [(CUPS1, 3, C.RECIBIDA, C.PDTE_INF), (CUPS1, 7, None, C.RECIBIDA),
                          (CUPS2, 7, C.SIN_FRA, C.PDTE_INF), (CUPS2, 8, C.SIN_FRA, C.PDTE_INF)])
        self.assertEqual(self.estado(CUPS1, 7), C.RECIBIDA)
        self.assertEqual(self.estado(CUPS1, 1), C.INF_ENVIADO)              # sin tocar
        self.assertEqual(self.estado(CUPS2, 9), C.SIN_FRA)                 # anulada y comp.
        self.assertIsNone(self.estado(CUPS3, 8))
        self.assertEqual(len(almacen.listar_cambios_gemweb()), 4)
        resultado = {(f["Factura"], f.get("Mes")): f["Resultado"] for f in r["listado"]}
        self.assertEqual(resultado[("F-ene", "ene-26")], G.SIN_CAMBIOS)
        self.assertEqual(resultado[("F-mar", "mar-26")], G.CAMBIA)
        self.assertEqual(resultado[("F-jul2", "jul-26")], G.OTRA_FACTURA)
        self.assertEqual(resultado[("F-anul", None)], "Descartada (anulada)")
        self.assertEqual(resultado[("F-ases", "ago-26")], G.NO_SYC)
        self.assertNotIn(("F-fuera", None), resultado)       # su suministro no esta en Gemweb
        self.assertEqual(len(r["facturas"]), 8)
        self.assertEqual(G.actualizar(hoy=dt.date(2026, 10, 26), cliente=api)["cambios"], [])

    def test_cups_que_no_esta_en_el_cuadro(self):
        api = GemwebFalso([_factura("4", "01/07/2026", "31/07/2026", "0", "F-otro")])
        api.get_inventory_original = api.get_inventory

        def inventario(category, **kw):
            if category == "subministraments":
                return _xml([{"id": "4", "cups": "ES0000000000000009XX"}])
            return api.get_inventory_original(category, **kw)
        api.get_inventory = inventario
        r = G.actualizar(hoy=dt.date(2026, 10, 26), cliente=api)
        self.assertEqual([f["Resultado"] for f in r["listado"]], [G.NO_EN_CUADRO])
        self.assertEqual(r["cambios"], [])

    def test_simular_no_guarda(self):
        api = GemwebFalso([_factura("1", "01/07/2026", "31/07/2026")])
        self.assertEqual(len(G.actualizar(simular=True, hoy=dt.date(2026, 10, 26),
                                          cliente=api)["cambios"]), 1)
        self.assertIsNone(self.estado(CUPS1, 7))
        self.assertEqual(almacen.listar_cambios_gemweb(), [])

    def test_inicio_en_enero_cubre_el_año_anterior(self):
        self.assertEqual(G.desde(dt.date(2027, 1, 15)), dt.date(2026, 10, 1))


if __name__ == "__main__":
    unittest.main()
