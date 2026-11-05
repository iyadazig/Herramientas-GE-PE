# -*- coding: utf-8 -*-
"""Datos regulados por año: semilla, copia del último año, lo guardado prevalece, validación y
adaptador para el estudio de potencia."""

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
import datos_regulados as dr
from potencia import calendario

RAIZ = Path(__file__).resolve().parents[1]


class Semilla(unittest.TestCase):
    def test_la_semilla_es_valida(self):
        self.assertEqual(dr.validar(dr.datos_anio(2026, {})), [])
        self.assertEqual(dr.validar_comunes(dr.comunes()), [])
        self.assertEqual(dr.anios({}), [2026])

    def test_todas_las_tarifas_con_sus_periodos(self):
        d = dr.datos_anio(2026, {})
        self.assertEqual(list(d["tarifas"]), dr.TARIFAS)
        for t in dr.TARIFAS:
            n = dr.n_periodos(t)
            self.assertEqual(len(d["tarifas"][t]["peajes_energia"]), n, t)
            self.assertEqual(len(d["tarifas"][t]["perdidas"]), n, t)

    def test_temporadas_y_periodos_iguales_que_el_calendario_de_potencia(self):
        c = dr.comunes()
        letras = {"A": "Alta", "B": "Media alta", "B1": "Media", "C": "Baja"}
        for zona, por_letra in calendario._TEMPORADAS.items():
            sistema = "Baleares" if zona == "Illes Balears" else zona
            for letra, meses in por_letra.items():
                self.assertEqual(c["temporadas"][sistema][letras[letra]], meses, (zona, letra))
                self.assertEqual(tuple(c["periodos_temporada"][sistema][letras[letra]]),
                                 calendario._PERIODOS_TIPO_DIA[zona][letra], (zona, letra))

    def test_periodos_validos_por_sistema(self):
        self.assertEqual(dr.periodos_validos("Península", "6.1TD", 1), [1, 2, 6])
        self.assertEqual(dr.periodos_validos("Canarias", "6.1TD", 8), [1, 3, 6])
        self.assertEqual(dr.periodos_validos("Ceuta", "3.0TD", 4), [3, 5, 6])
        self.assertEqual(dr.periodos_validos("Ceuta", "2.0TD", 4), [1, 2, 3])

    def test_terminos_y_perdidas(self):
        d = dr.datos_anio(2026, {})
        t = dr.terminos(d, "2.0TD")
        self.assertEqual(t["potencia"], [round(23.324952 + 4.379461, 6),
                                         round(0.443770 + 0.281653, 6), 0, 0, 0, 0])
        self.assertEqual(t["energia"][:3], [round(0.033261 + 0.064292, 6),
                                            round(0.016409 + 0.012858, 6),
                                            round(0.000077 + 0.003215, 6)])
        self.assertIsNone(t["srad"])
        self.assertEqual(len(dr.terminos(d, "6.1TD")["srad"]), 6)
        self.assertEqual(dr.perdidas(d, "6.1TD", "Canarias"), d["tarifas"]["6.1TD"]["perdidas"])
        d["perdidas_sistema"] = {"Canarias": {"6.1TD": [0.1] * 6}}
        self.assertEqual(dr.perdidas(d, "6.1TD", "Canarias"), [0.1] * 6)
        self.assertEqual(dr.perdidas(d, "6.1TD", "Península"), d["tarifas"]["6.1TD"]["perdidas"])


class Anios(unittest.TestCase):
    def test_año_sin_publicar_copia_el_ultimo(self):
        d = dr.datos_anio(2029, {})
        self.assertEqual((d["anio"], d["origen"], d["copia_de"]), (2029, "copia", 2026))
        self.assertEqual(dr.datos_anio(2024, {})["copia_de"], 2026)    # anterior: el primero
        self.assertEqual(dr.datos_anio(2026, {})["origen"], "semilla")

    def test_lo_guardado_prevalece(self):
        g = {2026: dr.limpiar(dr.datos_anio(2026, {})), 2027: dr.limpiar(dr.datos_anio(2026, {}))}
        g[2026]["generales"]["fnee"] = 3.0
        g[2027]["generales"]["fnee"] = 4.0
        self.assertEqual(dr.datos_anio(2026, g)["generales"]["fnee"], 3.0)
        self.assertEqual(dr.datos_anio(2026, g)["origen"], "guardado")
        d = dr.datos_anio(2028, g)
        self.assertEqual((d["copia_de"], d["generales"]["fnee"]), (2027, 4.0))
        exportado = json.loads(dr.exportar_json(g))
        self.assertEqual(list(exportado["anios"]), ["2026", "2027"])
        self.assertNotIn("origen", exportado["anios"]["2027"])

    def test_la_copia_no_cambia_la_semilla(self):
        d = dr.datos_anio(2026, {})
        d["tarifas"]["6.1TD"]["perdidas"][0] = 9
        self.assertNotEqual(dr.datos_anio(2026, {})["tarifas"]["6.1TD"]["perdidas"][0], 9)


class Validar(unittest.TestCase):
    def setUp(self):
        self.d = dr.datos_anio(2026, {})

    def test_desvios_de_referencia(self):
        g = dr.datos_anio(2026, {})["generales"]
        self.assertAlmostEqual(g["desvios_coste_ref"] + g["desvios_saldo_ref"], 0.001)
        self.assertIn("C2", g["desvios_fuente"])
        self.d["generales"]["desvios_saldo_ref"] = "x"
        self.assertIn("tiene que ser un número", "\n".join(dr.validar(self.d)))

    def test_errores(self):
        self.d["tarifas"]["3.0TD"]["peajes_energia"] = [0.1] * 5
        self.d["tarifas"]["6.1TD"]["perdidas"][2] = -0.1
        self.d["generales"]["iee"] = 5.11
        self.d["mercado"]["pct_omip"] = 1.2
        self.d["mercado"]["anios_media_omie"] = 0
        del self.d["tarifas"]["6.4TD"]
        e = "\n".join(dr.validar(self.d))
        for texto in ("3.0TD peajes de energía: tiene que tener 6", "6.1TD pérdidas",
                      "tanto por uno", "ponderación de OMIP", "años de la media de OMIE",
                      "Falta la tarifa 6.4TD"):
            self.assertIn(texto, e)

    def test_temporadas_que_no_cubren_el_año(self):
        c = dr.comunes()
        c["temporadas"]["Ceuta"]["Alta"] = [1, 2, 8]
        self.assertIn("Temporadas de Ceuta", "\n".join(dr.validar_comunes(c)))


class Potencia(unittest.TestCase):
    def test_adaptador_igual_que_precios_json(self):
        precios = json.loads((RAIZ / "potencia" / "precios.json").read_text(encoding="utf-8"))
        precios.pop("_comentario")
        self.assertEqual(dr.precios_potencia(2026, {}), precios)

    def test_estudio_de_potencia_lee_los_datos_regulados(self):
        from estudio_potencia import servicio
        with mock.patch.object(dr, "_guardados", return_value={}):
            self.assertEqual(servicio.precios()["tarifas"]["6.1TD"]["tep"],
                             dr.datos_anio(2026, {})["tarifas"]["6.1TD"]["tep"])
        with mock.patch.object(dr, "precios_potencia", side_effect=OSError):
            self.assertIn("6.1TD", servicio.precios()["tarifas"])        # respaldo


class Guardado(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def test_guardar_leer_y_volver_a_la_semilla(self):
        self.assertEqual(dr._guardados(), {})                  # sin base de datos: no la crea
        self.assertFalse(almacen.ruta_bd().exists())
        d = dr.datos_anio(2027, {})
        d["generales"]["fnee"] = 2.9
        dr.guardar(2027, d, "ana")
        self.assertEqual(dr.anios(), [2026, 2027])
        self.assertEqual(dr.datos_anio(2027)["generales"]["fnee"], 2.9)
        self.assertEqual(dr.datos_anio(2027)["origen"], "guardado")
        self.assertNotIn("copia_de", almacen.datos_regulados_guardados()[2027])
        d["generales"]["fnee"] = 3.1
        dr.guardar(2027, d, "luis")
        self.assertEqual([h["usuario"] for h in almacen.historial_datos_regulados(2027)],
                         ["luis", "ana"])
        self.assertEqual(almacen.info_datos_regulados()[2027][1], "luis")
        almacen.borrar_datos_regulados(2027)
        self.assertEqual(dr.datos_anio(2027)["origen"], "copia")

    def test_no_guarda_datos_con_errores(self):
        d = dr.datos_anio(2026, {})
        d["mercado"]["pct_omip"] = 2
        with self.assertRaises(dr.DatosRegulados):
            dr.guardar(2026, d, "ana")


if __name__ == "__main__":
    unittest.main()
