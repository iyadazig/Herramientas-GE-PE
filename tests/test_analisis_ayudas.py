# -*- coding: utf-8 -*-
"""Ayudas de «Ofertas»: textos del ? de cada campo, avisos de coherencia y fórmula en una
línea."""

import re
import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from analisis_ofertas import ayudas
from analisis_ofertas import modelo as mo
from ofertas import solicitud as so


class Textos(unittest.TestCase):
    def test_todas_las_claves_usadas_existen(self):
        codigo = (RAIZ / "analisis_ofertas" / "paginas.py").read_text(encoding="utf-8")
        usadas = set(re.findall(r'AY\["(\w+)"\]', codigo))
        codigo = (RAIZ / "ofertas" / "paginas.py").read_text(encoding="utf-8")
        usadas |= set(re.findall(r'ayudas\.AYUDA\["(\w+)"\]', codigo))
        self.assertTrue(usadas)
        self.assertEqual(usadas - set(ayudas.AYUDA), set())

    def test_modos_de_ssaa_explicados(self):
        for modo in mo.MODOS_SSAA:
            self.assertIn(f"**{modo}**", ayudas.AYUDA["ssaa_modo"])

    def test_tipologias_de_la_solicitud(self):
        self.assertEqual([o for o in so.PRECIO_INDEXADO if o not in ayudas.TIPOLOGIAS], [])


class Avisos(unittest.TestCase):
    def pt(self, **kw):
        o = mo.Oferta(modalidad=mo.PASS_THROUGH, atr_incluido=False, pagos_capacidad=True,
                      os=True, om=True, fee=1.0, k=0.015, ssaa_modo=mo.SSAA_PERIODO)
        for k, v in kw.items():
            setattr(o, k, v)
        return o

    def test_oferta_normal_sin_avisos(self):
        self.assertEqual(ayudas.avisos_coherencia(self.pt()), [])

    def test_errores_tipicos(self):
        av = " ".join(ayudas.avisos_coherencia(self.pt(
            pagos_capacidad=False, os=False, om=False, fee=0, k=1.5,
            ssaa_modo=mo.SSAA_BANDA, banda_min=20, banda_max=10)))
        for texto in ("sin pagos por capacidad", "sin fee", "K del 150", "máxima debe ser mayor"):
            self.assertIn(texto, av)

    def test_fee_cero_con_termino_b_no_avisa(self):
        o = self.pt(fee=0, ssaa_modo=mo.SSAA_INCLUIDOS, pagos_capacidad=False, os=False, om=False)
        o.precios["6.1TD"] = [mo.PreciosAnio(termino_b=[15.0] * 6)]
        self.assertEqual(ayudas.avisos_coherencia(o), [])

    def test_k_y_factor_de_ssaa_a_la_vez(self):
        av = ayudas.avisos_coherencia(self.pt(k_ssaa=True, ssaa_f2=1.015))
        self.assertTrue(any("dos veces" in x for x in av))
        self.assertFalse(any("dos veces" in x for x in ayudas.avisos_coherencia(
            self.pt(k_ssaa=False, ssaa_f2=1.015))))
        self.assertFalse(any("dos veces" in x for x in ayudas.avisos_coherencia(
            self.pt(k_ssaa=True, ssaa_f1=1.02))))

    def test_fijo_con_componentes_de_mercado(self):
        o = mo.Oferta(modalidad=mo.FIJO, pagos_capacidad=True)
        self.assertTrue(ayudas.avisos_coherencia(o))

    def test_fijo_con_k(self):
        self.assertTrue(any("suele ir incluida" in x for x in ayudas.avisos_coherencia(
            mo.Oferta(modalidad=mo.FIJO, k=0.015))))
        self.assertEqual(ayudas.avisos_coherencia(mo.Oferta(modalidad=mo.FIJO)), [])


class Formula(unittest.TestCase):
    def test_pass_through(self):
        o = mo.Oferta(modalidad=mo.PASS_THROUGH, atr_incluido=False, pagos_capacidad=True,
                      os=True, fee=2.5, k=0.015, ssaa_modo=mo.SSAA_PERIODO,
                      ssaa_perdidas=mo.PERD_TARIFA)
        f = ayudas.resumen_formula(o)
        self.assertIn("(OMIE × apuntamiento + pagos por capacidad + OS) × (1 + pérdidas)", f)
        self.assertIn("peajes y cargos", f)
        self.assertIn("fee 2,5 €/MWh", f)
        self.assertIn("K 1,5 %", f)

    def test_pass_pool_media_mensual_y_fijo(self):
        pp = mo.Oferta(modalidad=mo.PASS_POOL, ci_media_mensual=True)
        self.assertIn("Ci × (OMIE medio del mes)", ayudas.resumen_formula(pp))
        fijo = mo.Oferta(modalidad=mo.FIJO, unidad="c€/kWh", ssaa_modo=mo.SSAA_TECHO, techo=11)
        f = ayudas.resumen_formula(fijo)
        self.assertIn("precio fijo (c€/kWh)", f)
        self.assertIn("techo 11 €/MWh", f)


class FormulaLatex(unittest.TestCase):
    def test_pass_through_con_k_sobre_ssaa(self):
        o = mo.Oferta(modalidad=mo.PASS_THROUGH, atr_incluido=False, pagos_capacidad=True,
                      fee=1.0, k=0.015, k_ssaa=True, ssaa_modo=mo.SSAA_PERIODO,
                      ssaa_perdidas=mo.PERD_TARIFA)
        f, ley = ayudas.formula_latex(o)
        self.assertTrue(f.startswith(r"\begin{aligned}") and f.endswith(r"\end{aligned}"))
        self.assertIn(r"\Big[", f)
        self.assertIn(r"\Big]\times(1+1{,}5\%)", f)
        # los SSAA van dentro del corchete de la K y los peajes fuera
        self.assertLess(f.index("SSAA"), f.index(r"\Big]"))
        self.assertGreater(f.index("Peajes y cargos"), f.index(r"\Big]"))
        self.assertTrue(any("Apunt." in x for x in ley))
        self.assertNotIn(chr(8), f)

    def test_ssaa_fuera_de_la_k(self):
        o = mo.Oferta(modalidad=mo.PASS_THROUGH, k=0.015, k_ssaa=False,
                      ssaa_modo=mo.SSAA_TECHO, techo=11)
        f, _ = ayudas.formula_latex(o)
        self.assertGreater(f.index(r"\max"), f.index(r"\Big]"))

    def test_perdidas_en_un_solo_parentesis_e_impuesto_electrico(self):
        o = mo.Oferta(modalidad=mo.PASS_THROUGH, atr_incluido=False, pagos_capacidad=True,
                      fee=1.0, base_fee=mo.BARRAS, fnee=mo.BARRAS, gdo=2.0, k=0.015,
                      k_ssaa=True, ssaa_modo=mo.SSAA_PERIODO, ssaa_perdidas=mo.PERD_TARIFA)
        f, ley = ayudas.formula_latex(o)
        self.assertEqual(f.count(r"\times (1+\text{Pérdidas})"), 1)    # todo en un paréntesis
        grupo = f[:f.index(r"\times (1+\text{Pérdidas})")]
        for termino in ("OMIE", "Capacidad", "Fee", "FNEE", "SSAA"):
            self.assertIn(termino, grupo)
        self.assertNotIn("GdO", grupo)                                   # la GdO va sin pérdidas
        self.assertTrue(f.endswith(r"\bigg\}\times(1+\text{I.E.})\end{aligned}"))
        self.assertTrue(any(x.startswith("I.E.") for x in ley))

    def test_fijo_con_atr(self):
        f, _ = ayudas.formula_latex(mo.Oferta(modalidad=mo.FIJO, atr_incluido=True))
        self.assertIn("Precio fijo", f)
        self.assertNotIn("Peajes", f)


if __name__ == "__main__":
    unittest.main()
