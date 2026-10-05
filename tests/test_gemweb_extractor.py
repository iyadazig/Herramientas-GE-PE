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
