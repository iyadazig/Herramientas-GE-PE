# -*- coding: utf-8 -*-
"""Páginas del análisis de ofertas y de los datos regulados: conversión de las tablas editables
(sin Streamlit) y carga de la página en una base de datos temporal (AppTest)."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
import datos_regulados as dr
from analisis_ofertas import paginas_regulados as pr


class TablasRegulados(unittest.TestCase):
    def test_ida_y_vuelta_de_cada_concepto(self):
        d = dr.datos_anio(2026, {})
        nuevo = dr.limpiar(d)
        for c, _, _ in dr.CONCEPTOS:
            pr.aplicar_concepto(nuevo, c, pr.tabla_concepto(d, c))
        self.assertEqual(nuevo["tarifas"], dr.limpiar(d)["tarifas"])

    def test_cambiar_un_valor_y_la_2_0td_solo_sus_periodos(self):
        d = dr.datos_anio(2026, {})
        t = pr.tabla_concepto(d, "peajes_potencia")
        t.loc[t["Tarifa"] == "6.1TD", "P1"] = 30.0
        t.loc[t["Tarifa"] == "2.0TD", "P5"] = 99.0           # la 2.0TD no tiene P5: se ignora
        pr.aplicar_concepto(d, "peajes_potencia", t)
        self.assertEqual(d["tarifas"]["6.1TD"]["peajes_potencia"][0], 30.0)
        self.assertEqual(len(d["tarifas"]["2.0TD"]["peajes_potencia"]), 2)
        self.assertEqual(dr.validar(d), [])

    def test_srad_vacio_queda_sin_srad(self):
        d = dr.datos_anio(2026, {})
        t = pr.tabla_concepto(d, "srad")
        self.assertTrue(t.loc[t["Tarifa"] == "2.0TD", "P1"].isna().all())
        pr.aplicar_concepto(d, "srad", t)
        self.assertIsNone(d["tarifas"]["2.0TD"]["srad"])
        self.assertEqual(len(d["tarifas"]["6.1TD"]["srad"]), 6)

    def test_perdidas_por_sistema(self):
        d = dr.datos_anio(2026, {})
        t = pr.tabla_perdidas_sistema(d)
        self.assertEqual(len(t), 4 * len(dr.TARIFAS))
        fila = (t["Sistema"] == "Canarias") & (t["Tarifa"] == "6.1TD")
        for p in pr.PERIODOS:
            t.loc[fila, p] = 0.08
        pr.aplicar_perdidas_sistema(d, t)
        self.assertEqual(d["perdidas_sistema"], {"Canarias": {"6.1TD": [0.08] * 6}})
        self.assertEqual(dr.perdidas(d, "6.1TD", "Canarias"), [0.08] * 6)
        self.assertEqual(dr.validar(d), [])


class PaginaRegulados(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def test_la_pagina_carga_y_guarda(self):
        from streamlit.testing.v1 import AppTest

        def app():
            from analisis_ofertas import paginas_regulados
            paginas_regulados.pagina_datos_regulados()
        with mock.patch("estilo._usuario", return_value="ana"):
            at = AppTest.from_function(app, default_timeout=60).run()
            self.assertFalse(at.exception)
            self.assertIn("semilla", at.info[0].value)
            at.number_input(key=[k for k in at.session_state.filtered_state
                                 if k.endswith("_g_fnee")][0]).set_value(2.5).run()
            at.button[0].click().run()
            self.assertFalse(at.exception)
        self.assertEqual(almacen.datos_regulados_guardados()[2026]["generales"]["fnee"], 2.5)


if __name__ == "__main__":
    unittest.main()
