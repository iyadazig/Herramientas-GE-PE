# -*- coding: utf-8 -*-
"""Parámetros utilizados: tablas con su fuente y fecha y exportación a Excel."""

import datetime as dt
import io
import sys
import unittest
from pathlib import Path
from unittest import mock

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import datos_mercado as dm
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import parametros as pm
from tests.test_analisis_calculo import estudio


def calculado(anio, mes, fuente=ap.ESIOS):
    return {"anio": anio, "mes": mes, "fuente": fuente,
            "factores": {f"{s}|{c}": [1.1, 0.9, None, None, None, 1.0]
                         for s in ap.SISTEMAS for c in ap.CALENDARIOS}}


class Estudio(unittest.TestCase):
    def test_tablas_congeladas_y_excel(self):
        e = estudio()
        e.ofertas = []
        e.ap_omie["Península|6P"] = {1: [1.2, 1.0, None, None, None, 0.8]}
        e.ap_origen["Península|6P"] = {1: "ene-26 (ESIOS)"}
        e.omip_manual["2027-02"] = 120.0
        e.mercado["2027-01"]["contrato"] = "M Jan-27"
        t = pm.tablas_estudio(e)
        self.assertEqual(set(t), {"Regulados por tarifa", "Generales", "Mercado y SSAA",
                                  "Apuntamientos", "Fuentes"})
        reg = t["Regulados por tarifa"]
        fila = reg[(reg["Tarifa"] == "2.0TD") & (reg["Concepto"] == "Peajes + cargos de potencia")]
        self.assertEqual(list(fila[["P1", "P2"]].iloc[0]), [12.0, 6.0])
        self.assertTrue(fila[["P3", "P4", "P5", "P6"]].isna().all(axis=None))
        merc = t["Mercado y SSAA"]
        self.assertEqual(list(merc["Mes"][:2]), ["ene-27", "feb-27"])
        self.assertEqual(list(merc["Contrato OMIP"][:2]), ["M Jan-27", "manual"])
        self.assertEqual(merc["OMIP (€/MWh)"][1], 120.0)
        aps = t["Apuntamientos"]
        self.assertEqual((aps["Dato de"][0], aps["P1"][0]), ("ene-26 (ESIOS)", 1.2))
        libro = openpyxl.load_workbook(io.BytesIO(pm.excel(t)))
        self.assertEqual(libro.sheetnames, ["Regulados por tarifa", "Generales",
                                            "Mercado y SSAA", "Apuntamientos", "Fuentes"])
        self.assertEqual(libro["Apuntamientos"]["A4"].value, "Sistema")


class Vigentes(unittest.TestCase):
    def test_sin_ficheros_de_mercado(self):
        calculados = [calculado(2025, m) for m in range(10, 13)] + \
                     [calculado(2026, m) for m in range(1, 10)]
        with mock.patch.object(dm, "congelar", side_effect=dm.DatosMercado("sin red")):
            t = pm.tablas_vigentes([2026, 2027], calculados, hoy=dt.date(2026, 11, 5))
        self.assertIn("sin red", t["_avisos"][0])
        self.assertNotIn("Mercado", t)
        aps = t["Apuntamientos"]
        self.assertEqual(len(aps), 5 * 2 * 12)
        octubre = aps[(aps["Sistema"] == "Canarias") & (aps["Mes"] == "octubre")]
        self.assertEqual(list(octubre["Dato de"]), ["oct-25 (ESIOS)"] * 2)
        self.assertEqual(sorted(set(t["Generales"]["Año"])), [2026, 2027])
        self.assertIn("Copia de 2026 (sin publicar)", set(t["Generales"]["Origen"]))
        pm.excel(t)                                         # «_avisos» no se exporta


if __name__ == "__main__":
    unittest.main()
