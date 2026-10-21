# -*- coding: utf-8 -*-
"""Cliente comun de Gemweb y parsers de las consultas (XML simulado, sin red)."""

import sys
import threading
import time
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import gemweb
from gemweb_extractor import cliente as c


def xml(texto):
    return ET.fromstring(texto)


class Parsers(unittest.TestCase):
    def test_inventario(self):
        df = c.inventory_a_dataframe(xml(
            "<r><subministrament><id>7</id><cups>ES0000000000000000XX</cups></subministrament>"
            "<subministrament><id>8</id><cups>ES0000000000000001XX</cups></subministrament></r>"))
        self.assertEqual(list(df["id"]), ["7", "8"])

    def test_consumo_por_periodo_mensual_y_diario(self):
        mensual = c.consumo_periodos_a_dataframe(xml(
            "<r><subministrament><id>7</id><units>kWh</units><year value='2026'>"
            "<month value='4'><consumption_p1>10</consumption_p1><consumption_p2>5</consumption_p2>"
            "</month></year></subministrament></r>"))
        self.assertEqual((mensual["fecha"][0], mensual["total"][0]), ("2026-04", 15.0))
        diario = c.consumo_periodos_a_dataframe(xml(
            "<r><subministrament><id>7</id><units>kWh</units><year value='2026'>"
            "<month value='4'><day value='2'><consumption_p6>3</consumption_p6></day></month>"
            "</year></subministrament></r>"))
        self.assertEqual((diario["fecha"][0], diario["total"][0]), ("2026-04-02", 3.0))

    def test_telelecturas(self):
        df = c.metering_a_dataframe(xml(
            "<r><subministrament><id>7</id><units>kWh</units><values>"
            "<value date='2026-04-01  00:15'>1.5</value><value date='2026-04-01  00:30'>2</value>"
            "</values></subministrament></r>"))
        self.assertEqual(list(df["valor"]), [1.5, 2.0])

    def test_coste_consumo(self):
        df = c.cost_consumption_a_dataframe(xml(
            "<r><subministrament><id>7</id><units>EUR</units><completeness>98</completeness>"
            "<year value='2026'><month value='3'>120.5</month></year></subministrament></r>"))
        self.assertEqual((df["fecha"][0], df["valor"][0], df["completitud_%"][0]),
                         ("2026-03", 120.5, "98"))

    def test_optimizacion(self):
        df = c.power_optimization_a_dataframe(xml(
            "<r><subministrament><id>7</id><completeness>100</completeness>"
            "<current><power><p1>100</p1></power><costs><total>900</total></costs></current>"
            "<optimum><power><p1>80</p1></power><costs><total>800</total></costs>"
            "<savings><total>100</total></savings></optimum></subministrament></r>"))
        fila = df.iloc[0]
        self.assertEqual((fila["actual_p1_kW"], fila["optimo_p1_kW"], fila["ahorro_total_€"]),
                         (100.0, 80.0, 100.0))

    def test_partir_rango(self):
        tramos = c.partir_rango("2026-01-01", "2026-03-15", 31)
        self.assertEqual(tramos[0], ("2026-01-01", "2026-02-01"))
        self.assertEqual(tramos[-1][1], "2026-03-15")
        for (_a, fin), (ini, _b) in zip(tramos, tramos[1:]):
            self.assertEqual(fin, ini)                  # sin huecos entre tramos


class AyudasDePaginas(unittest.TestCase):
    """Rodeos de los fallos de Gemweb comprobados contra la API real (oct-2026)."""

    def setUp(self):
        from gemweb_extractor import paginas
        self.p = paginas

    def test_co2_sin_agrupaciones_inferiores_a_la_hora(self):
        self.assertNotIn("quart-horari", self.p.periodos_para("co2"))
        self.assertIn("horari", self.p.periodos_para("co2"))
        self.assertIn("quart-horari", self.p.periodos_para("consum"))

    def test_un_mensaje_por_error_repetido_en_los_tramos(self):
        fallidos = ["2026-01-01 → 2026-02-01: get_metering: El dispositivo no obtiene lecturas.",
                    "2026-02-01 → 2026-03-01: get_metering: El dispositivo no obtiene lecturas."]
        self.assertEqual(self.p.errores_tramos(fallidos), ["El dispositivo no obtiene lecturas."])

    def test_mes_anterior(self):
        import datetime as dt
        self.assertEqual(self.p.mes_anterior(dt.date(2026, 3, 15)),
                         (dt.date(2026, 2, 1), dt.date(2026, 2, 28)))
        self.assertEqual(self.p.mes_anterior(dt.date(2026, 1, 3)),
                         (dt.date(2025, 12, 1), dt.date(2025, 12, 31)))

    def test_consumo_mensual_desde_telelecturas(self):
        import pandas as pd

        class Falso:
            def get_metering_por_tramos(self, **k):
                self.k = k
                return pd.DataFrame({"id_suministro": ["7", "7"], "fecha": ["2026-01", "2026-02"],
                                     "valor": [10.0, 20.0], "unidades": ["kWh", "kWh"]}), []

        api = Falso()
        df = self.p.consumo_mensual_telelecturas(api, 7, "2026-01-01", "2026-02-28")
        self.assertEqual((api.k["period"], api.k["field"]), ("mensual", "consum"))
        self.assertEqual(list(df["mes"]), [1, 2])
        self.assertEqual(list(df["valor"]), [10.0, 20.0])

    def test_xml_roto_de_gemweb_se_explica(self):
        api = c.GemwebClient("u", "s")
        api._token_valido = lambda: True
        import requests
        original = requests.post
        requests.post = lambda *a, **k: type("R", (), {
            "text": "<r><day value='0'>1</day></month></r>",
            "raise_for_status": lambda self: None})()
        try:
            with self.assertRaises(c.GemwebError) as e:
                api._post("get_consumption_by_period")
        finally:
            requests.post = original
        self.assertIn("fallo del servidor de Gemweb", str(e.exception))
        self.assertNotIn("temporal", str(e.exception))

    def test_campos_de_la_simulacion(self):
        campos, valores = self.p.campos_simulacion({"p1": 0.15, "p2": None, "p3": 0.1},
                                                   {"p1": 30.5})
        self.assertEqual(campos, "preu_ene_p1,preu_ene_p3,preu_pot_p1")
        self.assertEqual(valores, "0.15,0.1,30.5")             # decimales con punto

    def test_resultado_de_la_simulacion(self):
        df, gen = self.p.simulacion_a_dataframe(xml(
            "<resultat><simulation><id>7</id><data_ini_sim>01/01/2026</data_ini_sim>"
            "<data_fin_sim>31/03/2026</data_fin_sim><import_ene_p1>10</import_ene_p1>"
            "<import_pot_p1>5</import_pot_p1><import_react_p1>0</import_react_p1>"
            "<import_ene>10</import_ene><import_tot>15</import_tot></simulation></resultat>"))
        self.assertEqual(list(df["Periodo"]), ["P1"])            # los totales van aparte
        self.assertEqual((gen["energia"], gen["total"], gen["desde"]), (10.0, 15.0, "01/01/2026"))
        vacio, gen0 = self.p.simulacion_a_dataframe(xml(
            "<resultat><simulation><id>7</id><import_tot>0</import_tot></simulation></resultat>"))
        self.assertTrue(vacio.empty)
        self.assertEqual(gen0["total"], 0.0)

    def _factura(self, ini, fin, p1=0, p6=0, **marcas):
        extra = "".join("<%s>%s</%s>" % (k, v, k) for k, v in marcas.items())
        return ("<factura_elec><periode_ini>%s</periode_ini><periode_fin>%s</periode_fin>"
                "<tarifa>6.1TD</tarifa><comercialitzadora>COMERCIAL S.A.</comercialitzadora>"
                "<consum_p1>%s</consum_p1><consum_p6>%s</consum_p6><consum_tot>%s</consum_tot>"
                "<anulada>0</anulada><abonada>0</abonada>%s</factura_elec>"
                % (ini, fin, p1, p6, p1 + p6, extra))

    def test_facturas_por_periodo_descarta_anuladas_y_abonos(self):
        root = xml("<registre>%s%s%s</registre>" % (
            self._factura("31/12/2025", "31/01/2026", p1=10, p6=5),
            self._factura("31/12/2025", "31/01/2026", p1=10, anulada=1),
            self._factura("31/01/2026", "28/02/2026", p1=3, abonament=1)))
        filas, descartadas = self.p.facturas_por_periodo("ES0000000000000000XX", root)
        self.assertEqual((len(filas), descartadas), (1, 2))
        f = filas[0]
        self.assertEqual((f["Fin factura"].isoformat(), f["P1 kWh"], f["P6 kWh"], f["Total kWh"]),
                         ("2026-01-31", 10.0, 5.0, 15.0))
        self.assertEqual((f["Tarifa"], f["Comercializadora"]), ("6.1TD", "COMERCIAL S.A."))

    def test_elegir_ultimos_12_meses_y_periodo(self):
        import datetime as dt
        filas = [{"Fin factura": dt.date(2025, m, 28)} for m in range(1, 13)] + \
                [{"Fin factura": dt.date(2026, m, 28)} for m in range(1, 10)] + \
                [{"Fin factura": dt.date(2026, 9, 15)}]          # un mes facturado en dos
        ult = self.p.elegir_facturas(filas, "ultimos12")
        self.assertEqual(ult[0]["Fin factura"], dt.date(2025, 10, 28))
        self.assertEqual(ult[-1]["Fin factura"], dt.date(2026, 9, 28))
        self.assertEqual(len(ult), 13)                           # 12 meses, 13 facturas
        per = self.p.elegir_facturas(filas, "periodo", dt.date(2026, 1, 1), dt.date(2026, 3, 31))
        self.assertEqual([f["Fin factura"].month for f in per], [1, 2, 3])
        self.assertEqual(self.p.elegir_facturas([], "ultimos12"), [])

    def test_meses_naturales_de_los_ultimos_12(self):
        import datetime as dt
        facturas = [{"Fin factura": dt.date(2026, 2, 28)}]
        self.assertEqual(self.p.meses_naturales(facturas, "ultimos12"),
                         (dt.date(2025, 3, 1), dt.date(2026, 2, 28)))
        self.assertEqual(self.p.meses_naturales([], "ultimos12"), (None, None))
        a, b = dt.date(2026, 1, 1), dt.date(2026, 6, 30)
        self.assertEqual(self.p.meses_naturales(facturas, "periodo", a, b), (a, b))

    def test_variantes_cups_de_22_caracteres(self):
        self.assertEqual(self.p.variantes_cups("es0000000000000000xx0f"),
                         ["ES0000000000000000XX0F", "ES0000000000000000XX"])
        self.assertEqual(self.p.variantes_cups("ES0000000000000000XX"), ["ES0000000000000000XX"])

    def test_mes_natural_y_excel(self):
        import io
        import openpyxl
        filas = self.p.mes_natural_por_periodo("ES0000000000000000XX", xml(
            "<r><subministrament><id>7</id><units>kWh</units><values><year value='2026'>"
            "<month value='3'><consumption_p1>4</consumption_p1><consumption_p6>6"
            "</consumption_p6></month></year></values></subministrament></r>"))
        self.assertEqual((filas[0]["Mes"], filas[0]["P1 kWh"], filas[0]["Total kWh"]),
                         ("2026-03", 4.0, 10.0))
        import pandas as pd
        datos = self.p.excel_consumo_periodos(pd.DataFrame([{"CUPS": "x"}]), pd.DataFrame(filas),
                                              pd.DataFrame([{"CUPS": "x", "Estado": "OK"}]))
        wb = openpyxl.load_workbook(io.BytesIO(datos))
        self.assertEqual(wb.sheetnames, ["Facturas por periodo", "Mes natural por periodo",
                                         "Resumen"])

    def test_avisos_de_gemweb_en_paginas_existentes(self):
        # las claves son los titulos del menu (inicio.py): si se renombra uno, que se note
        import re
        menu = open(Path(__file__).resolve().parents[1] / "inicio.py", encoding="utf-8").read()
        titulos = set(re.findall(r'gw\.pagina_\w+, "([^"]+)"', menu))
        self.assertTrue(set(self.p.AVISOS_GEMWEB) <= titulos)


class ClienteComun(unittest.TestCase):
    def test_un_solo_token_con_varios_a_la_vez(self):
        class Lento(c.GemwebClient):
            pedidos = 0

            def _renovar_token(self):
                time.sleep(0.2)
                Lento.pedidos += 1
                self._token = "t"
                import datetime as dt
                self._token_expira = dt.datetime.now() + dt.timedelta(hours=1)

        api = Lento("u", "s")
        import requests
        original = requests.post
        requests.post = lambda *a, **k: type("R", (), {"text": "<r/>",
                                                      "raise_for_status": lambda self: None})()
        try:
            hilos = [threading.Thread(target=api._post, args=("get_inventory",))
                     for _ in range(8)]
            for h in hilos:
                h.start()
            for h in hilos:
                h.join()
        finally:
            requests.post = original
        self.assertEqual(Lento.pedidos, 1)

    def test_obtener_api_reutiliza_el_cliente(self):
        a = gemweb.obtener_api("usuario", "clave")
        self.assertIs(a, gemweb.obtener_api("usuario", "clave"))
        self.assertIsNot(a, gemweb.obtener_api("usuario", "otra"))

    def test_mismo_error_en_toda_la_plataforma(self):
        self.assertIs(gemweb.GemwebError, c.GemwebError)


if __name__ == "__main__":
    unittest.main()
