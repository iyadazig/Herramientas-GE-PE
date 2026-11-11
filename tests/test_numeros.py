# -*- coding: utf-8 -*-
"""Números en formato español en todo el programa: coma decimal y punto de miles al mostrar;
coma o punto decimal al escribir o pegar."""

import math
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import numeros as N


class Numeros(unittest.TestCase):
    def test_leer(self):
        casos = {"16,56": 16.56, "1.234,5": 1234.5, "16.56": 16.56, "0.394": 0.394,
                 "1.234.567": 1234567.0, " 27,40 ": 27.4, "-0,5": -0.5, "12 €": 12.0,
                 "5 %": 5.0, "": None, None: None, "abc": None, "-": None, 3: 3.0,
                 float("nan"): None, True: None, pd.NA: None}
        for v, esperado in casos.items():
            self.assertEqual(N.leer(v), esperado, repr(v))

    def test_texto(self):
        self.assertEqual(N.texto(1234.5678), "1.234,5678")
        self.assertEqual(N.texto(0.0511269632), "0,0511269632")
        self.assertEqual(N.texto(2.658), "2,658")
        self.assertEqual(N.texto(1500.0), "1.500")
        self.assertEqual(N.texto(-1234567.25), "-1.234.567,25")
        self.assertEqual(N.texto(1234.5678, 2), "1.234,57")
        self.assertEqual(N.texto(1234.5, miles=False), "1234,5")
        self.assertEqual(N.texto(-0.0), "0")
        self.assertEqual(N.texto(None), "")
        self.assertEqual(N.texto(float("nan")), "")
        for v in (0.0511269632, 1234.5678, -16.5, 2.658, 1e-7):
            self.assertTrue(math.isclose(N.leer(N.texto(v)), v), v)   # ida y vuelta

    def test_con_coma(self):
        for t, esperado in {"0.394": "0,394", "-16.5": "-16,5", ".5": ",5", "1,5": "1,5",
                            "1.234.567": "1.234.567", "16": "16", "abc": "abc"}.items():
            self.assertEqual(N.con_coma(t), esperado)
        self.assertIsNone(N.con_coma(None))

    def test_tabla(self):
        df = pd.DataFrame({"Año": [2026.0, 2027.0], "Concepto": ["FNEE", "IEE"],
                           "Valor": [2.658, 0.0511269632], "Kilos": [1500.5, 2.0],
                           "Entero": [3, 4], "Mezcla": [1.5, None]})
        t = N.tabla(df)
        self.assertEqual(list(t["Año"]), ["2026", "2027"])          # años sin punto de miles
        self.assertEqual(list(t["Valor"]), ["2,658", "0,0511269632"])
        self.assertEqual(list(t["Kilos"]), ["1.500,5", "2"])
        self.assertEqual(list(t["Entero"]), [3, 4])                  # enteros sin tocar
        self.assertEqual(list(t["Concepto"]), ["FNEE", "IEE"])
        self.assertEqual(list(t["Mezcla"]), ["1,5", ""])
        self.assertEqual(df["Valor"][0], 2.658)                      # el original no cambia


if __name__ == "__main__":
    unittest.main()
