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
            clave = [k for k in at.session_state.filtered_state if k.endswith("_gt_fnee")][0]
            self.assertEqual(at.text_input(key=clave).value, "2,658")    # con coma decimal
            at.text_input(key=clave).input("2.5").run()                  # escrito con punto
            self.assertEqual(at.text_input(key=clave).value, "2,5")      # se ve con coma
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
        t.loc[1, "Reducción IEE (%)"] = 85.0
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
        self.o.cierres_tipo = "Por MW (carga base)"
        c = ap_.tabla_cierres(self.e, self.o)
        c.loc[2, "MW pico"] = 0.2
        c.loc[2, "Precio pico (€/MWh)"] = 150.0
        ap_.aplicar_cierres(self.o, c)
        pico = [x for x in self.o.cierres if x.mes == 3][0]
        self.assertEqual((pico.mw_pico, pico.precio_pico), (0.2, 150.0))

    def test_ph_de_canarias_y_desvios(self):
        from analisis_ofertas import paginas as ap_
        from analisis_ofertas import modelo as mo
        self.e.suministros[1].sistema = "Canarias"
        self.assertIn("Canarias Ph|6P", ap_.claves_ap(self.e))
        self.e.ph_dif, self.e.ph_origen = {1: 20.0}, {1: "ene-26 (ESIOS)"}
        t = ap_.tabla_ph(self.e)
        t.loc[0, "Ph − diario (€/MWh)"] = 18.5
        t.loc[1, "Ph − diario (€/MWh)"] = 21.0
        ap_.aplicar_ph(self.e, t)
        self.assertEqual(self.e.ph_dif, {1: 18.5, 2: 21.0})
        self.assertEqual(self.e.ph_origen, {1: "manual", 2: "manual"})
        self.e.regulados[2027]["generales"].update(desvios_coste_ref=0.441,
                                                   desvios_saldo_ref=-0.44,
                                                   desvios_fuente="C2")
        self.assertIn("coste de desvíos 0,441", ap_.referencia_desvios(self.e))
        self.assertEqual(mo.SISTEMA_PH, "Canarias Ph")


class RegaladosCongelados(unittest.TestCase):
    def test_avisa_si_un_anio_se_congelo_antes_de_cambiar_y_se_actualiza(self):
        """Un análisis congela cada año la primera vez que lo necesita: 2027 con el FNEE de
        antes y 2028-2029 con el nuevo. Se avisa de la diferencia y «actualizar» la corrige."""
        import datetime as dt
        from analisis_ofertas import modelo as mo
        from analisis_ofertas import servicio
        e = mo.Estudio(inicio=dt.date(2027, 1, 1))
        e.ofertas = [mo.Oferta(nombre="36 M", duracion=36)]
        servicio.congelar_regulados(e, guardados={})
        self.assertEqual(sorted(e.regulados), [2027, 2028, 2029])
        self.assertEqual(servicio.diferencias_regulados(e, guardados={}), [])
        e.regulados[2027]["generales"]["fnee"] = 2.4374639           # congelado antes
        cambios = servicio.diferencias_regulados(e, guardados={})
        self.assertEqual(cambios, ["FNEE 2027: 2,4374639 €/MWh → 2,658 €/MWh"])
        e.regulados[2028]["tarifas"]["6.1TD"]["peajes_energia"][0] += 0.001
        self.assertTrue(any("tarifa en 2028" in x for x in
                            servicio.diferencias_regulados(e, guardados={})))
        servicio.congelar_regulados(e, guardados={}, todos=True)
        self.assertEqual(servicio.diferencias_regulados(e, guardados={}), [])
        self.assertEqual({a: r["generales"]["fnee"] for a, r in e.regulados.items()},
                         {2027: 2.658, 2028: 2.658, 2029: 2.658})


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

    def test_doce_valores_pegados_en_fila_o_columna(self):
        from analisis_ofertas import paginas as ap_
        doce = [10.5 + i for i in range(12)]
        textos = [f"{v:.1f}".replace(".", ",") for v in doce]
        self.assertEqual(ap_.valores_pegados("\t".join(textos) + "\r\n"), doce)     # fila
        self.assertEqual(ap_.valores_pegados("\r\n".join(textos) + "\r\n"), doce)   # columna
        self.assertEqual(ap_.valores_pegados(" ; ".join(textos)), doce)
        self.assertEqual(ap_.valores_pegados("1,5\t\t2,5"), [1.5, None, 2.5])     # hueco
        self.assertEqual(ap_.valores_pegados(""), [])

    def test_pegar_precios_de_excel(self):
        import pandas as pd
        from analisis_ofertas import paginas as ap_
        per = ["P1", "P2", "P3", "P4", "P5", "P6"]
        tabla = pd.DataFrame([{"Grupo": g, "Año": a, **{p: None for p in per}}
                              for g in ("6.1TD", "3.0TD") for a in (1, 2)])
        tabla.loc[0, "P6"] = 0.09                                 # ya había un P6
        # una fila de 5 celdas con coma decimal: P6 no se toca
        nueva, error = ap_.pegar_precios(tabla, "0,1512\t0,1401\t0,1203\t0,1105\t0,1002\r\n")
        self.assertIsNone(error)
        self.assertEqual(list(nueva.loc[0, per]), [0.1512, 0.1401, 0.1203, 0.1105, 0.1002, 0.09])
        self.assertTrue(pd.isna(tabla.loc[0, "P1"]))              # la original no cambia
        # varias filas desde la fila 2 (3.0TD año 1), con punto decimal y un hueco
        nueva, error = ap_.pegar_precios(tabla, "0.2\t0.19\t0.18\n0.21\t\t0.17", desde=2)
        self.assertIsNone(error)
        self.assertEqual(list(nueva.loc[2, ["P1", "P2", "P3"]]), [0.2, 0.19, 0.18])
        self.assertEqual(nueva.loc[3, "P1"], 0.21)
        self.assertIsNone(nueva.loc[3, "P2"])
        self.assertEqual(nueva.loc[3, "P3"], 0.17)
        # una columna de 6 celdas = P1…P6 de la fila elegida
        nueva, error = ap_.pegar_precios(tabla, "1\n2\n3\n4\n5\n6", desde=1)
        self.assertIsNone(error)
        self.assertEqual(list(nueva.loc[1, per]), [1.0, 2.0, 3.0, 4.0, 5.0, 6.0])
        # errores: más de 6 periodos, más filas que la tabla, texto que no es número
        for texto, desde, trozo in (("1\t2\t3\t4\t5\t6\t7", 0, "6 periodos"),
                                    ("1\n2\n3\n4\n5\n6\n7", 0, "6 periodos"),
                                    ("1\t2\n3\t4", 3, "solo quedan"),
                                    ("1\tabc", 0, "no es un número"),
                                    ("", 0, "nada pegado")):
            nueva, error = ap_.pegar_precios(tabla, texto, desde)
            self.assertIn(trozo, error or "", texto)
            self.assertIs(nueva, tabla)

    def test_punto_decimal_se_ve_con_coma(self):
        from analisis_ofertas import paginas as ap_
        casos = {"0.394": "0,394", "-16.5": "-16,5", ".5": ",5", "1,5228": "1,5228",
                 "1.234.567": "1.234.567", "16": "16", "abc": "abc", "": ""}
        for texto, esperado in casos.items():
            self.assertEqual(ap_.con_coma(texto), esperado, texto)
        self.assertIsNone(ap_.con_coma(None))
        self.assertEqual(ap_.numero_es(ap_.con_coma("0.394")), 0.394)

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
    def test_comprobar_curvas_y_elegir(self):
        import datetime as dt
        from analisis_ofertas import servicio
        from analisis_ofertas import apuntamientos as apm
        from analisis_ofertas import modelo as mo
        from analisis_ofertas import paginas as ap_
        from tests.test_analisis_calculo import estudio
        from tests.test_apuntamientos import curva_año, horas_del_mes
        import curva_consumo

        class Falso:
            def curva(self, cups, desde, hasta):
                if cups.endswith("BB"):
                    raise RuntimeError("El CUPS no está en el inventario de Gemweb.")
                self.periodo = (desde, hasta)
                # Gemweb tiene datos del 01/06/2025 al 24/08/2026 (agosto, a medias)
                valores = {(f, h, q): 0.25 for (f, h) in curva_año(dt.date(2025, 6, 1), 450)
                           for q in range(1, 5)}
                return curva_consumo.Curva("qh", valores), {}

        e = estudio(2)
        falso = Falso()
        curvas, filas = servicio.comprobar_curvas_gemweb(
            e, falso, hoy=dt.date(2026, 11, 5),
            precios_de=lambda a, m: ({k: 80.0 for k in horas_del_mes(a, m)}, apm.ESIOS))
        self.assertEqual(falso.periodo, (dt.date(2024, 11, 1), dt.date(2026, 11, 4)))
        a, b = filas
        self.assertEqual((a["En Gemweb"], a["Se puede usar"], b["En Gemweb"]), ("Sí", "Sí", "No"))
        self.assertEqual((a["Datos desde"], a["Datos hasta"]),
                         (dt.date(2025, 6, 1), dt.date(2026, 8, 24)))
        # los 12 últimos meses completos que tiene: ago-25 a jul-26
        self.assertEqual((a["Año usado desde"], a["Año usado hasta"]),
                         (dt.date(2025, 8, 1), dt.date(2026, 7, 31)))
        self.assertIn("inventario", b["Motivo"])
        cups_a, cups_b = e.suministros[0].cups, e.suministros[1].cups
        self.assertIn(cups_a, e.ap_cliente)
        self.assertEqual(min(k[0] for k in curvas[cups_a]), dt.date(2025, 8, 1))
        self.assertAlmostEqual(next(iter(curvas[cups_a].values())), 1.0)
        # hasta que el usuario lo decide, todos con el estimado
        self.assertEqual({s.apuntamiento for s in e.suministros}, {mo.ESTIMADO})
        t = ap_.tabla_eleccion(e, filas)
        self.assertEqual(list(t["Usar su curva"]), [False, False])
        no = servicio.elegir_apuntamiento(e, {cups_a: True, cups_b: True})
        self.assertEqual(no, [cups_b])
        self.assertEqual([s.apuntamiento for s in e.suministros], [mo.CLIENTE, mo.ESTIMADO])
        self.assertEqual(list(ap_.tabla_eleccion(e, filas)["Usar su curva"]), [True, False])
        servicio.elegir_apuntamiento(e, {cups_a: False})
        self.assertEqual(e.suministros[0].apuntamiento, mo.ESTIMADO)


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
