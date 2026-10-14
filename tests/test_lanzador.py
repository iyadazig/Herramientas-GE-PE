# -*- coding: utf-8 -*-
"""Pagina inicial servida en espanol (sin traduccion automatica del navegador)."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import lanzador


class PaginaEnEspanol(unittest.TestCase):
    def test_html_marcado_en_espanol(self):
        html = b'<!DOCTYPE html>\n<html lang="en">\n  <head>\n    <meta charset="UTF-8" />'
        out = lanzador._html_en_espanol(html)
        self.assertIn(b'<html lang="es" translate="no">', out)
        self.assertIn(b'<meta name="google" content="notranslate" />', out)
        self.assertNotIn(b'lang="en"', out)

    def test_index_real_de_streamlit(self):
        # si una version nueva de Streamlit cambia la cabecera, que se note aqui
        import streamlit
        ruta = Path(streamlit.__file__).parent / "static" / "index.html"
        self.assertIn(b'<html lang="es" translate="no">',
                      lanzador._html_en_espanol(ruta.read_bytes()))


if __name__ == "__main__":
    unittest.main()
