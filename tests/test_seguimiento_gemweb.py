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
        self.assertEqual(G.nuevo_estado(None), C.RECIBIDA)
        self.assertEqual(G.nuevo_estado(""), C.RECIBIDA)
        self.assertEqual(G.nuevo_estado(C.SIN_FRA), C.RECIBIDA)
        for otro in (C.RECIBIDA, C.PDTE_INF, C.INF_ENVIADO, C.INF_NO_ENVIADO, C.TEMPORAL,
                     C.CADENCIA, "-"):
            self.assertIsNone(G.nuevo_estado(otro))

    def _api(self):
        return GemwebFalso([
            _factura("1", "01/07/2026", "31/07/2026", "0", "F-jul"),
            _factura("1", "01/03/2026", "31/03/2026", "1", "F-mar"),    # Recibida: no se toca
            _factura("1", "01/01/2026", "31/01/2026", "1", "F-ene"),    # ya Inf enviado
            _factura("1", "01/12/2025", "31/12/2025", "1", "F-2025"),   # año anterior: fuera
            _factura("2", "01/07/2026", "31/08/2026", "1", "F-bim"),    # bimestral verificada
            _factura("2", "01/07/2026", "31/07/2026", "0", "F-jul2"),   # mismo mes que la anterior
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-anul", anulada="1"),
            _factura("2", "01/09/2026", "30/09/2026", "0", "F-comp", complementaria="1"),
            _factura("3", "01/08/2026", "31/08/2026", "0", "F-ases"),   # asesoramiento
            _factura("9", "01/08/2026", "31/08/2026", "0", "F-fuera")])  # no esta en Gemweb

    def test_buscar_propone_sin_cambiar_el_cuadro(self):
        api = self._api()
        r = G.buscar(hoy=dt.date(2026, 10, 26), cliente=api)
        self.assertEqual(r["inicio"], "2026-01-01")
        self.assertEqual(api.llamadas[1][1]["search_values"], "2026-01-01")
        self.assertEqual(sorted((c["cups"], c["mes"], c["antes"], c["despues"])
                                for c in r["cambios"]),
                         [(CUPS1, 7, None, C.RECIBIDA), (CUPS2, 7, C.SIN_FRA, C.RECIBIDA),
                          (CUPS2, 8, C.SIN_FRA, C.RECIBIDA)])
        self.assertIsNone(self.estado(CUPS1, 7))                           # nada aplicado
        resultado = {(f["Factura"], f.get("Mes")): f["Resultado"] for f in r["listado"]}
        self.assertEqual(resultado[("F-ene", "ene-26")], G.SIN_CAMBIOS)
        self.assertEqual(resultado[("F-mar", "mar-26")], G.SIN_CAMBIOS)
        self.assertEqual(resultado[("F-bim", "ago-26")], G.CAMBIA)
        self.assertEqual(resultado[("F-jul2", "jul-26")], G.OTRA_FACTURA)
        self.assertEqual(resultado[("F-anul", None)], "Descartada (anulada)")
        self.assertEqual(resultado[("F-ases", "ago-26")], G.NO_SYC)
        self.assertNotIn("F-2025", {f for f, _ in resultado})
        self.assertNotIn("F-fuera", {f for f, _ in resultado})

    def test_aplicar_todos_o_seleccionados(self):
        r = G.buscar(hoy=dt.date(2026, 10, 26), cliente=self._api())
        G.guardar_propuesta(r)
        agosto = [c for c in r["cambios"] if c["mes"] == 8]
        hechos, saltados = G.aplicar(agosto, "prueba")
        self.assertEqual((len(hechos), saltados), (1, []))
        self.assertEqual(self.estado(CUPS2, 8), C.RECIBIDA)
        self.assertIsNone(self.estado(CUPS1, 7))
        self.assertEqual(len(G.propuesta()["cambios"]), 2)                # quedan pendientes
        # alguien cambia una celda despues de la busqueda: ese cambio no se aplica
        almacen.guardar_sgc_estados([{"id_suministro": self.ids[CUPS1], "anio": 2026, "mes": 7,
                                      "estado": C.TEMPORAL}], "prueba")
        hechos, saltados = G.aplicar(G.propuesta()["cambios"], "prueba")
        self.assertEqual((len(hechos), len(saltados)), (1, 1))
        self.assertEqual(self.estado(CUPS1, 7), C.TEMPORAL)
        self.assertEqual(self.estado(CUPS2, 7), C.RECIBIDA)
        self.assertEqual(self.estado(CUPS1, 3), C.RECIBIDA)                # no se toca
        self.assertEqual(G.propuesta()["cambios"], [])
        self.assertEqual(len(almacen.listar_cambios_gemweb()), 2)
        self.assertEqual(G.buscar(hoy=dt.date(2026, 10, 26), cliente=self._api())["cambios"], [])

    def test_cups_del_cuadro_fuera_de_gemweb(self):
        fuera = G.cups_fuera_de_gemweb(cliente=GemwebFalso([]))
        self.assertEqual(fuera, [])
        api = GemwebFalso([])
        api.get_inventory = lambda category, **kw: _xml([{"id": "1", "cups": CUPS1}])
        self.assertEqual(sorted(f["CUPS"] for f in G.cups_fuera_de_gemweb(cliente=api)),
                         [CUPS2, CUPS3])

    def test_solo_este_año(self):
        self.assertEqual(G.desde(dt.date(2027, 1, 15)), dt.date(2027, 1, 1))


if __name__ == "__main__":
    unittest.main()
