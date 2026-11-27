"""«Nuevo gestor» en la ficha de una oferta: «Guardar gestor» lo da de alta en la base de
datos de gestores y lo deja elegido en la oferta (con datos ficticios en una carpeta temporal)."""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import almacen
import config
from streamlit.testing.v1 import AppTest

SCRIPT = '''
import streamlit as st
from analisis_ofertas import modelo as mo, paginas as pg
if "o" not in st.session_state:
    st.session_state.o = mo.Oferta(comercializadora="Comercializadora Ficticia")
o = st.session_state.o
c1, c2, c3 = st.columns(3)
pg._gestor(o, c2, c3, "t")
st.write(f"OFERTA:{o.gestor}|{o.email_gestor}")
'''


class NuevoGestorEnOferta(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen.guardar_gestor("Otro Gestor", "otro@ejemplo.com", "Otra Comercializadora")

    def tearDown(self):
        self.p.stop()
        self.tmp.cleanup()

    def test_guardar_gestor_lo_registra_y_lo_deja_elegido(self):
        at = AppTest.from_string(SCRIPT, default_timeout=60).run()
        sb = at.selectbox[0]
        self.assertEqual(len(sb.options), 2)                 # vacío + «Nuevo gestor…»
        sb.select_index(1).run()
        at.text_input[0].input("Gestor Ficticio").run()
        self.assertEqual(at.text_input[0].label, "Nombre del nuevo gestor")  # no se pierde
        at.text_input[1].input("gestor@ejemplo.com").run()
        at.button[0].click().run()
        self.assertFalse(at.exception)
        self.assertEqual(at.selectbox[0].value, "Gestor Ficticio")
        self.assertIn("OFERTA:Gestor Ficticio|gestor@ejemplo.com",
                      [m.value for m in at.markdown])
        g = [x for x in almacen.listar_gestores() if x["nombre"] == "Gestor Ficticio"]
        self.assertEqual((g[0]["email"], g[0]["comercializadora"]),
                         ("gestor@ejemplo.com", "Comercializadora Ficticia"))


if __name__ == "__main__":
    unittest.main()
