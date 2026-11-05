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


class TablasAnalisis(unittest.TestCase):
    def setUp(self):
        from tests.test_analisis_calculo import estudio, fijo
        self.e = estudio(2)
        self.e.actual = fijo(0.15, nombre="Actual", es_actual=True)
        self.o = fijo(0.1, nombre="F", duracion=24)
        self.e.ofertas = [self.o]

    def test_suministros(self):
        from analisis_ofertas import paginas as ap_
        t = ap_.tabla_suministros(self.e)
        t.loc[1, "Sistema"] = "Canarias"
        t.loc[1, "Reducción IEE"] = 0.85
        t.loc[0, "Incluir"] = False
        t.loc[0, "Apuntamiento"] = "Del cliente (curva)"     # sin curva: se queda el estimado
        ap_.aplicar_suministros(self.e, t)
        a, b = self.e.suministros
        self.assertEqual((a.activo, a.apuntamiento), (False, "Estimado (a mano)"))
        self.assertEqual((b.sistema, b.reduccion_iee), ("Canarias", 0.85))

    def test_mercado_y_ssaa(self):
        from analisis_ofertas import paginas as ap_
        t = ap_.tabla_mercado(self.e)
        self.assertEqual(len(t), 24)
        t.loc[1, "OMIP (€/MWh)"] = 150.0
        ap_.aplicar_mercado(self.e, t)
        self.assertEqual(self.e.omip_manual, {"2027-02": 150.0})
        t = ap_.tabla_mercado(self.e)
        self.assertEqual(t.loc[1, "Contrato"], "manual")
        t.loc[1, "OMIP (€/MWh)"] = 100.0                     # vuelve al de la sesión
        ap_.aplicar_mercado(self.e, t)
        self.assertEqual(self.e.omip_manual, {})
        s = ap_.tabla_ssaa(self.e)
        s.loc[0, "SSAA (€/MWh)"] = 20.0
        ap_.aplicar_ssaa(self.e, s)
        self.assertEqual(self.e.ssaa_mes[:2], [20.0, 10.0])

    def test_apuntamientos_manual(self):
        from analisis_ofertas import paginas as ap_
        clave = "Península|6P"
        self.e.ap_omie[clave] = {m: [1.0] * 6 for m in range(1, 13)}
        self.e.ap_origen[clave] = {m: "sep-26 (ESIOS)" for m in range(1, 13)}
        t = ap_.tabla_ap(self.e, clave)
        t.loc[2, "P2"] = 1.2
        ap_.aplicar_ap(self.e, clave, t)
        self.assertEqual(self.e.ap_omie[clave][3][1], 1.2)
        self.assertEqual((self.e.ap_origen[clave][3], self.e.ap_origen[clave][4]),
                         ("manual", "sep-26 (ESIOS)"))
        self.assertEqual(list(ap_.tabla_ap(self.e, "Península|3P").columns),
                         ["Mes", "P1", "P2", "P3", "Dato de"])

    def test_precios_por_grupo_y_año(self):
        from analisis_ofertas import paginas as ap_
        self.e.suministros[1].sistema = "Canarias"
        t = ap_.tabla_precios(self.e, self.o, "energia")
        self.assertEqual(list(zip(t["Grupo"], t["Año"])),
                         [("6.1TD", 1), ("6.1TD", 2), ("6.1TD-Canarias", 1),
                          ("6.1TD-Canarias", 2)])
        t.loc[1, "P1"] = 0.2                                  # año 2 de la Península: solo P1
        ap_.aplicar_precios(self.o, "energia", t)
        lista = self.o.precios["6.1TD"]
        self.assertEqual(lista[1].energia[:2], [0.2, None])
        self.assertNotIn("6.1TD-Canarias", {g for g, v in self.o.precios.items() if v and
                                            any(x is not None for x in v[0].energia)})
        t = ap_.tabla_precios(self.e, self.o, "energia")
        t.loc[1, "P1"] = None                                 # año 2 vacío: copia el año 1
        ap_.aplicar_precios(self.o, "energia", t)
        self.assertEqual(len(self.o.precios["6.1TD"]), 1)

    def test_grupos_cups_y_cierres(self):
        from analisis_ofertas import paginas as ap_
        import pandas as pd
        ap_.aplicar_grupos_propios(self.o, pd.DataFrame(
            [{"Grupo": "Especial", "Tarifa": "6.1TD", "Sistema": "Península"},
             {"Grupo": None, "Tarifa": None, "Sistema": None}]))
        self.assertEqual([g.nombre for g in self.o.grupos_propios], ["Especial"])
        t = ap_.tabla_cups_oferta(self.e, self.o)
        t.loc[0, "Grupo propio"] = "Especial"
        t.loc[1, "Incluido"] = False
        ap_.aplicar_cups_oferta(self.o, t)
        self.assertEqual(self.o.asignacion, {self.e.suministros[0].cups: "Especial"})
        self.assertEqual(self.o.excluidos, [self.e.suministros[1].cups])
        self.o.modalidad, self.o.cierres_tipo = "Indexado Pass-Pool", "Por %"
        c = ap_.tabla_cierres(self.e, self.o)
        self.assertEqual(len(c), 24)
        c.loc[0, "% cerrado"] = 0.5
        ap_.aplicar_cierres(self.o, c)
        self.assertEqual([(x.mes, x.pct) for x in self.o.cierres], [(1, 0.5)])


class Formato(unittest.TestCase):
    def test_numeros_pegados_de_excel(self):
        from analisis_ofertas import paginas as ap_
        casos = {"16,56": 16.56, "1.234,5": 1234.5, "16.56": 16.56, "1.234.567": 1234567.0,
                 " 27,40 ": 27.4, "-0,5": -0.5, "": None, None: None, "abc": None, 3: 3.0,
                 float("nan"): None, "12 €": 12.0}
        for texto, esperado in casos.items():
            self.assertEqual(ap_.numero_es(texto), esperado, texto)
        self.assertEqual([ap_.texto_es(v, 2) for v in (16.56, 20.0, None, 1234.567)],
                         ["16,56", "20", "", "1234,57"])
        self.assertEqual(ap_.contrato_es("M Jan-27"), "M ene-27")

    def test_meses_en_espanol_en_el_mercado(self):
        from analisis_ofertas import paginas as ap_
        from tests.test_analisis_calculo import estudio
        e = estudio()
        e.mercado["2027-01"]["contrato"] = "M Jan-27"
        t = ap_.tabla_mercado(e)
        self.assertEqual((t.loc[0, "Mes"], t.loc[0, "Contrato"], t.loc[0, "Clave"]),
                         ("ene-27", "M ene-27", "2027-01"))

    def test_comercializadoras(self):
        from analisis_ofertas import comercializadoras as coms
        self.assertEqual(len(coms.COMERCIALIZADORAS), 57)
        self.assertEqual(coms.COMERCIALIZADORAS[0], "ACCIONA GREEN ENERGY DEVELOPMENTS SL")
        self.assertIn("ENDESA ENERGÍA, S.A.U.", coms.COMERCIALIZADORAS)
        self.assertEqual(coms.opciones()[0], "")
        self.assertEqual(coms.opciones("OTRA SA")[:2], ["", "OTRA SA"])     # guardada antes

    def test_filtro_de_solicitudes(self):
        from analisis_ofertas import paginas as ap_
        sols = [{"usuario": "ana", "cliente": "EMPRESA UNO", "ref": "R-1", "cups": "ES01",
                 "denominaciones": "Sede central | Almacén"},
                {"usuario": "luis", "cliente": "OTRA EMPRESA", "ref": "", "cups": "ES02",
                 "denominaciones": "Oficina"}]
        self.assertEqual(len(ap_.solicitudes_filtradas(sols)), 2)
        self.assertEqual([s["usuario"] for s in ap_.solicitudes_filtradas(sols, usuarios=["ana"])],
                         ["ana"])
        self.assertEqual(len(ap_.solicitudes_filtradas(sols, cliente="otra")), 1)
        self.assertEqual(len(ap_.solicitudes_filtradas(sols, denominacion="almacén")), 1)
        self.assertEqual(len(ap_.solicitudes_filtradas(sols, denominacion="es02")), 1)


class Gemweb(unittest.TestCase):
    def test_comprobar_curvas(self):
        import datetime as dt
        from analisis_ofertas import servicio
        from analisis_ofertas import apuntamientos as apm
        from tests.test_analisis_calculo import estudio
        from tests.test_apuntamientos import curva_año, horas_del_mes
        import curva_consumo

        class Falso:
            def curva(self, cups, desde, hasta):
                if cups.endswith("BB"):
                    raise RuntimeError("El CUPS no está en el inventario de Gemweb.")
                self.periodo = (desde, hasta)
                valores = {(f, h, q): 0.25 for (f, h) in curva_año(dt.date(2025, 10, 1))
                           for q in range(1, 5)}
                return curva_consumo.Curva("qh", valores), {}

        e = estudio(2)
        falso = Falso()
        curvas, filas = servicio.comprobar_curvas_gemweb(
            e, falso, hoy=dt.date(2026, 11, 5),
            precios_de=lambda a, m: ({k: 80.0 for k in horas_del_mes(a, m)}, apm.ESIOS))
        self.assertEqual(falso.periodo, (dt.date(2025, 11, 1), dt.date(2026, 10, 31)))
        a, b = filas
        self.assertEqual((a["En Gemweb"], a["Se puede usar"], b["En Gemweb"]), ("Sí", "Sí", "No"))
        self.assertIn("inventario", b["Motivo"])
        self.assertIn(e.suministros[0].cups, e.ap_cliente)
        self.assertEqual(list(curvas), [e.suministros[0].cups])
        self.assertAlmostEqual(next(iter(curvas[e.suministros[0].cups].values())), 1.0)


class PaginaAnalisis(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def test_denominaciones_de_las_solicitudes(self):
        almacen.guardar_solicitud("ana", {"cliente": "X", "cups": ["ES01", "ES02"]},
                                  {"suministros": [{"cups": "ES01", "alias": "Sede"},
                                                   {"cups": "ES02", "alias": "Almacén"}]}, "", b"")
        self.assertEqual(almacen.listar_solicitudes()[0]["denominaciones"], "Sede | Almacén")

    def test_recorrido_completo(self):
        from streamlit.testing.v1 import AppTest
        from tests.test_analisis_calculo import estudio, fijo

        def app():
            import streamlit as st
            from analisis_ofertas import paginas
            if "ao_estudio" not in st.session_state:
                st.session_state.ao_estudio = st.session_state.pop("_inicial")
            paginas.pagina_nuevo_analisis()

        e = estudio(2)
        e.actual = fijo(0.15, nombre="Contrato actual", es_actual=True,
                        comercializadora="ACTUAL SA")
        e.ofertas = [fijo(0.1, nombre="F", duracion=24, comercializadora="COM A",
                          gestor="Ana", email_gestor="ana@example.com")]
        estado = {"omip": {"ultima_sesion": None}, "omie": {"ultimo_mes": None}, "avisos": []}
        import datos_mercado as dm
        with mock.patch("estilo._usuario", return_value="ana"), \
                mock.patch.object(dm, "estado", return_value=estado):
            at = AppTest.from_function(app, default_timeout=120)
            at.session_state["_inicial"] = e
            at.run()
            self.assertFalse(at.exception, at.exception)
            boton = next(b for b in at.button if b.label.startswith("Guardar el análisis"))
            boton.click().run()
            self.assertFalse(at.exception, at.exception)
            boton = next(b for b in at.button if b.label.startswith("Preparar informes"))
            boton.click().run()
            self.assertFalse(at.exception, at.exception)
        guardados = almacen.listar_analisis()
        self.assertEqual(len(guardados), 1)
        self.assertEqual(guardados[0]["comercializadoras"], "COM A")
        self.assertEqual(almacen.contactos_comercializadoras()["COM A"]["email"],
                         "ana@example.com")


if __name__ == "__main__":
    unittest.main()
