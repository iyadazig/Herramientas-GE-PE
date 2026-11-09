# -*- coding: utf-8 -*-
"""Informes del análisis de ofertas: PDF con el formato de la plantilla (informe, relación de
suministros o ficha monopunto, condiciones; A4 o A3 según las columnas) y Excel de cálculo con
fórmulas vivas. Datos ficticios. La igualdad de las fórmulas con el motor se comprobó con Excel
en local (no hay Excel en las pruebas)."""

import io
import re
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl

from analisis_ofertas import calculo, informe_excel, informe_pdf, informes
from analisis_ofertas import modelo as mo
from test_analisis_calculo import estudio, fijo, oferta


def _ofertas(e, n_fijas=2, n_index=2):
    e.actual = oferta(nombre="Contrato actual", es_actual=True, modalidad=mo.PASS_THROUGH,
                      atr_incluido=False, comercializadora="ELEIA ENERGÍA", duracion=12)
    e.actual.precios["6.1TD"] = [mo.PreciosAnio(termino_b=[0.0] * 6)]
    e.ofertas = []
    for i in range(n_fijas):
        e.ofertas.append(fijo(0.1 + i / 100, nombre=f"Fijo {i}", duracion=12 * (i % 3 + 1),
                              comercializadora=f"TOTALENERGIES {i}"))
    for i in range(n_index):
        o = oferta(nombre=f"Index PP {i}", modalidad=mo.PASS_POOL, ci_media_mensual=False,
                   atr_incluido=False, duracion=24, comercializadora="AXPO IBERIA",
                   condiciones={"validez": "Hasta el 28/11/2026", "pago": "15 días"})
        o.precios["6.1TD"] = [mo.PreciosAnio(termino_b=[20.0] * 6, ci=[1.08] * 6)]
        e.ofertas.append(o)
    return e


def _paginas(pdf: bytes) -> list:
    """[(ancho, alto, texto)] de cada página."""
    import fitz
    with fitz.open(stream=pdf, filetype="pdf") as d:
        return [(round(p.rect.width), round(p.rect.height), p.get_text().replace("\xa0", " "))
                for p in d]


def _invade_papel(pdf: bytes) -> list:
    """Palabras de los datos (no del papel ni el nº de página) que caen en la zona del logo, del
    texto del margen o del «GE&PE» gris."""
    import fitz
    malas = []
    with fitz.open(stream=pdf, filetype="pdf") as d:
        for n, p in enumerate(d):
            o = "horizontal" if p.rect.width > p.rect.height else "vertical"
            ruta = informe_pdf.papel(o, 1)
            z = informe_pdf.zona_libre(ruta)
            f = p.rect.width / z["ancho"]
            with fitz.open(str(ruta)) as papel:
                propias = {w[4] for w in papel[0].get_text("words")}
            for x0, y0, x1, y1, w, *_ in p.get_text("words"):
                if w in propias or w.startswith("Página") or w.isdigit() and y0 > p.rect.height - 30:
                    continue
                arriba = (z["arriba"] - informe_pdf.HUECO) * f
                abajo = p.rect.height - (z["abajo"] - informe_pdf.HUECO) * f
                if y0 < arriba - 1 or y1 > abajo + 1 or x0 < (z["izq"] - informe_pdf.HUECO) * f:
                    malas.append((n, w, round(y0), round(y1)))
    return malas


class Pdf(unittest.TestCase):
    def test_multipunto_ficha_resumen_y_condiciones(self):
        e = _ofertas(estudio(n_cups=2))
        res = calculo.calcular(e)
        b = informes.pdf(e, res, "REF", {"nombre": "Técnica GE&PE", "email": "tecnica@geype.com"})
        self.assertTrue(b.startswith(b"%PDF"))
        paginas = _paginas(b)
        self.assertGreaterEqual(len(paginas), 3)
        textos = [t for _, _, t in paginas]
        orden = [next(i for i, t in enumerate(textos) if clave in t)
                 for clave in ("RELACIÓN DE PUNTOS", "RESUMEN DEL ESTUDIO", "CONDICIONES DE LAS")]
        self.assertEqual(orden, sorted(orden))                 # ficha → resumen → condiciones
        self.assertIn("Consumo mensual por periodo", textos[0])   # gráficas de la cartera
        self.assertIn("Suministros de mayor consumo", textos[0])
        # cada hoja lleva el papel de GE&PE (marca de agua) y sin Fax
        self.assertTrue(all("Productividad" in t for t in textos))
        self.assertFalse(any("Fax" in t or "902 511 341" in t for t in textos))
        apaisadas = [t for a, al, t in paginas if a > al]
        self.assertTrue(apaisadas and all("954 636 737" in t for t in apaisadas))
        self.assertEqual(_invade_papel(b), [])

    def test_muchas_ofertas_pasan_a_a3(self):
        e = _ofertas(estudio(n_cups=2), n_fijas=3, n_index=14)
        b = informes.pdf(e, calculo.calcular(e))
        resumen = next(p for p in _paginas(b) if "RESUMEN DEL ESTUDIO" in p[2])
        self.assertEqual(resumen[:2], (1191, 842))             # A3 apaisado: letra legible
        # todas las hojas en el mismo papel (A3), también las condiciones
        self.assertEqual({max(a, al) for a, al, _ in _paginas(b)}, {1191})
        self.assertEqual(_invade_papel(b), [])

    def test_monopunto_ficha_vertical_primero(self):
        e = _ofertas(estudio(n_cups=1))
        b = informes.pdf(e, calculo.calcular(e))
        paginas = _paginas(b)
        self.assertEqual(paginas[0][:2], (595, 842))           # ficha en A4 vertical, la primera
        self.assertIn("FICHA DEL PUNTO DE SUMINISTRO", paginas[0][2])
        self.assertEqual(_invade_papel(b), [])

    def test_zona_libre_de_los_papeles(self):
        for o in ("vertical", "horizontal"):
            z = informe_pdf.zona_libre(informe_pdf.papel(o, 1))
            # los datos pueden ir sobre el «GE&PE» gris, pero no sobre el pie (nombre y
            # dirección de GE&PE)
            self.assertTrue(25 < z["abajo"] < 60, z["abajo"])
            self.assertGreater(z["arriba"], 50)                # el logo
            self.assertGreater(z["izq"], 25)                   # texto del Registro Mercantil

    def test_nombres_cortos_como_la_plantilla(self):
        f = {"comercializadora": "TOTALENERGIES CLIENTES, S.A.U.", "nombre": "Total PP 24 m",
             "producto": "", "tipo": "Indexado", "modalidad": mo.PASS_POOL, "duracion": 24,
             "ssaa": "SSAA incluidos", "cierres": ""}
        self.assertEqual(informe_pdf.cabecera_oferta(f), "TOTAL index. PP 24 M<br/>(SSAA incluidos)")
        f.update(comercializadora="ENDESA ENERGÍA", tipo="Fijo", duracion=12, nombre="Endesa PF")
        self.assertEqual(informe_pdf.cabecera_oferta(f), "ENDESA<br/>(SSAA incluidos)")


class Excel(unittest.TestCase):
    def setUp(self):
        self.e = _ofertas(estudio(n_cups=2))
        self.res = calculo.calcular(self.e)
        self.wb = openpyxl.load_workbook(io.BytesIO(informes.excel(self.e, self.res, "REF")))

    def test_hojas(self):
        self.assertEqual(self.wb.sheetnames[:4], ["Resumen", "Info suministro",
                                                  "Parámetros utilizados", "Contrato actual"])
        self.assertEqual(len(self.wb.sheetnames), 4 + len(self.e.ofertas))

    def test_formulas_vivas(self):
        ws = self.wb["Index PP 0"]
        self.assertTrue(str(ws["H8"].value).startswith("="))           # coste total del contrato
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("=")]
        self.assertTrue(any("SUMPRODUCT" in f for f in formulas))       # consumo × precio
        self.assertTrue(any("'Parámetros utilizados'!" in f for f in formulas))
        self.assertTrue(any("'Info suministro'!" in f for f in formulas))
        rs = self.wb["Resumen"]
        enlaces = [c.value for row in rs.iter_rows() for c in row
                   if isinstance(c.value, str) and "'Index PP 0'!" in c.value]
        self.assertTrue(enlaces)

    def test_total_por_tarifa_en_multipunto(self):
        ws = self.wb["Fijo 0"]
        textos = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
        self.assertTrue(any(t.startswith("TOTAL DE LA CARTERA") for t in textos))
        self.assertTrue(any(t.startswith("TOTAL TARIFA 6.1TD") for t in textos))

    def test_nombres_de_hoja_unicos_y_validos(self):
        usados = set()
        a = informe_excel._nombre_hoja("Oferta: 12/24 [x]", usados)
        b = informe_excel._nombre_hoja("Oferta: 12/24 [x]", usados)
        self.assertNotEqual(a, b)
        self.assertFalse(re.search(r"[\[\]\*\?/\\:]", a))


if __name__ == "__main__":
    unittest.main()
