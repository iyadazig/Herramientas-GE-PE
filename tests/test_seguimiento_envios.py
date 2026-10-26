# -*- coding: utf-8 -*-
"""Envios de Seguimiento y Control: informes, avisos de incidencias, «Inf enviado» en el
cuadro y el correo por Microsoft Graph (con requests simulado). Datos ficticios."""

import datetime as dt
import sys
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
from seguimiento import columnas as C
from seguimiento import correo
from seguimiento import datos as D
from seguimiento import incidencias as INC
from seguimiento import informes as INF
from tests.test_seguimiento import CUPS1, CUPS2, BaseDatos

CONTACTOS = [
    {"Técnico": "Tecnico Uno", "Auxiliar": "Aux Uno", "Cliente/Grupo": "Grupo Uno ",
     "Denominación suministro": "Planta Uno", "CUPS": CUPS1, "Nombre cliente": "Cliente Uno",
     "Para": "uno@ejemplo.invalid; otro@ejemplo.invalid", "CC": "copia@ejemplo.invalid"},
    {"Técnico": "Tecnico Uno", "Auxiliar": "Aux Uno", "Cliente/Grupo": "Grupo Uno",
     "Denominación suministro": "Planta Dos", "CUPS": CUPS2, "Nombre cliente": "Cliente Uno",
     "Para": "uno@ejemplo.invalid", "CC": None},
    {"Técnico": "Tecnico Dos", "Auxiliar": "Aux Dos", "Cliente/Grupo": "Grupo Dos",
     "Denominación suministro": "Planta Tres", "CUPS": "ES0000000000000000ZZ",
     "Nombre cliente": "Cliente Dos", "Para": "dos@ejemplo.invalid", "CC": None},
]
TECNICOS = {"Tecnico Uno": "tecnico@ejemplo.invalid"}


class Informes(unittest.TestCase):
    def test_nombre_del_pdf(self):
        i = INF.parsear_nombre_pdf("20260101_20260131_Inf_SyC_Planta_Uno_Refacturacion.pdf")
        self.assertEqual((i["periodo_inicial"], i["periodo_final"], i["denominacion_suministro"],
                          i["tipo_factura"]), ("20260101", "20260131", "Planta_Uno",
                                               "Refacturacion"))
        self.assertIsNone(INF.parsear_nombre_pdf("informe.pdf")["denominacion_suministro"])

    def test_un_correo_por_grupo_con_sus_informes(self):
        pdfs = [("20260101_20260131_Inf_SyC_Planta_Uno.pdf", b"1"),
                ("20260201_20260228_Inf_SyC_Planta Dos.pdf", b"2"),
                ("20260101_20260131_Inf_SyC_Planta Tres.pdf", b"3"), ("otro.pdf", b"4")]
        envios, sin_casar, aviso = INF.preparar(pdfs, [dict(c) for c in CONTACTOS], "Aux Uno",
                                                TECNICOS)
        self.assertIsNone(aviso)
        self.assertEqual(sorted(sin_casar), ["20260101_20260131_Inf_SyC_Planta Tres.pdf",
                                             "otro.pdf"])
        self.assertEqual(len(envios), 1)
        e = envios[0]
        self.assertEqual(e["grupo"], "Grupo Uno")
        self.assertEqual(e["asunto"], "Informe Seguimiento: Grupo Uno")
        self.assertEqual(e["para"], "otro@ejemplo.invalid; uno@ejemplo.invalid")
        self.assertEqual(e["cc"], "copia@ejemplo.invalid; tecnico@ejemplo.invalid")
        self.assertEqual(len(e["adjuntos"]), 2)
        self.assertIn("los siguientes informes de seguimiento", e["html"])
        self.assertIn("01/01/26 - 31/01/26", e["html"])
        self.assertEqual(sorted(INF.meses_a_marcar(e)), [(CUPS1, 2026, 1), (CUPS2, 2026, 2)])

    def test_un_solo_suministro_va_en_el_asunto(self):
        envios, _, _ = INF.preparar([("20260101_20260131_Inf_SyC_Planta Dos.pdf", b"2")],
                                    [dict(c) for c in CONTACTOS], "Aux Uno", {})
        self.assertEqual(envios[0]["asunto"], "Informe Seguimiento: Planta Dos")
        self.assertIn("el siguiente informe de seguimiento", envios[0]["html"])

    def test_meses_cubiertos(self):
        self.assertEqual(D.meses_cubiertos(dt.date(2026, 1, 1), dt.date(2026, 6, 30)),
                         [(2026, m) for m in range(1, 7)])
        self.assertEqual(D.meses_cubiertos(dt.date(2026, 1, 5), dt.date(2026, 2, 4)),
                         [(2026, 1)])
        self.assertEqual(D.meses_cubiertos(dt.date(2026, 1, 20), dt.date(2026, 2, 3)),
                         [(2026, 1)])
        self.assertEqual(D.meses_cubiertos(dt.date(2025, 12, 1), dt.date(2026, 1, 31)),
                         [(2025, 12), (2026, 1)])


class MarcarEnviados(BaseDatos):
    def test_solo_estados_con_factura(self):
        self.importar()                       # CUPS1: ene Inf enviado, mar Recibida
        uno = next(s["id"] for s in almacen.listar_sgc_suministros() if s["cups"] == CUPS1)
        almacen.guardar_sgc_estados([{"id_suministro": uno, "anio": 2026, "mes": 4,
                                      "estado": C.SIN_FRA}], "prueba")
        marcados, sin_marcar = D.marcar_enviados(
            [(CUPS1, 2026, 1), (CUPS1, 2026, 3), (CUPS1, 2026, 4),
             ("ES0000000000000000NO", 2026, 1)], "prueba")
        self.assertEqual([m["mes"] for m in marcados], ["mar-26"])
        self.assertEqual([(m["mes"], m["estado"]) for m in sin_marcar],
                         [("abr-26", C.SIN_FRA), ("ene-26", "no está en el cuadro")])
        estados = almacen.listar_sgc_estados(2026)
        self.assertEqual(estados[(uno, 2026, 3)]["estado"], C.INF_ENVIADO)
        self.assertEqual(estados[(uno, 2026, 3)]["origen"], "envío")


def _facturas(filas):
    columnas = ["nom_contacte_1_1", "nom_contacte_1_2", "nombre", "nombre.1", "CUPS",
                "dirección", "Número de factura", "Período inicial", "Período final",
                "Importe total exceso potencia", "Importe reactiva", "Lectura"]
    return pd.DataFrame(filas, columns=columnas)


class Incidencias(unittest.TestCase):
    def setUp(self):
        base = ["Tecnico Uno", "Aux Uno", "Grupo Uno", "Planta Uno", CUPS1, "Calle 1"]
        self.facturas = _facturas([
            base + ["F1", dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 31), "12,5", 0, "R"],
            base + ["F2", dt.datetime(2026, 2, 1), dt.datetime(2026, 2, 28), 0, 3.2, "R"],
            base + ["F3", dt.datetime(2026, 3, 1), dt.datetime(2026, 3, 31), 0, 0, "E"],
            base + ["F4", dt.datetime(2026, 4, 1), dt.datetime(2026, 4, 30), 0, 0, None],
            ["Tecnico Uno", "Aux Uno", "Grupo X", "Planta X", "ES0000000000000000XY", "",
             "F5", dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 31), 5, 0, "R"]])

    def test_incidencias_y_correo(self):
        envios, sin_contacto, aviso = INC.preparar(self.facturas, CONTACTOS, "Aux Uno", TECNICOS)
        self.assertIsNone(aviso)
        self.assertEqual(list(sin_contacto["numero_factura"]), ["F5"])
        self.assertEqual(len(envios), 1)
        e = envios[0]
        self.assertEqual(e["asunto"], "Aviso incidencia en factura_Grupo Uno")
        self.assertEqual([f["Número factura"] for f in e["filas"]], ["F1", "F2", "F3"])
        self.assertEqual([f["Incidencias"] for f in e["filas"]],
                         ["Exceso de Potencia (12.50 €)", "Exceso de Reactiva (3.20 €)",
                          "Lectura estimada"])
        self.assertIn("revisar los equipos de compensación de energía reactiva", e["html"])
        self.assertEqual(e["cc"], "copia@ejemplo.invalid; tecnico@ejemplo.invalid")
        self.assertEqual(e["filas"][0]["Inicio factura"], "01/01/26")

    def test_otro_auxiliar_no_tiene_incidencias(self):
        envios, _, aviso = INC.preparar(self.facturas, CONTACTOS, "Aux Dos", {})
        self.assertEqual(envios, [])
        self.assertEqual(aviso, "No hay incidencias para este auxiliar.")

    def test_segundo_envio_solo_facturas_nuevas(self):
        nuevas = INC.incidencias_nuevas(
            self.facturas.iloc[:2].rename(columns={}), self.facturas)
        self.assertEqual(list(nuevas["Número de factura"]), ["F3", "F4", "F5"])
        self.assertIsNone(INC.incidencias_nuevas(self.facturas, self.facturas))
        with self.assertRaises(ValueError):
            INC.incidencias_nuevas(pd.DataFrame({"a": [1]}), self.facturas)


class Respuesta:
    def __init__(self, estado, datos=None):
        self.status_code, self._datos, self.text = estado, datos or {}, ""

    def json(self):
        return self._datos


class Correo(BaseDatos):
    def setUp(self):
        super().setUp()
        correo._tokens.clear()
        correo.guardar_configuracion("inquilino-ficticio", "aplicacion-ficticia")
        cifrar = mock.patch.object(correo, "_cifrar", lambda t: "texto:" + t)
        descifrar = mock.patch.object(correo, "_descifrar", lambda t: t.split(":", 1)[1])
        cifrar.start()
        descifrar.start()
        self.addCleanup(cifrar.stop)
        self.addCleanup(descifrar.stop)

    def test_sin_configuracion(self):
        correo.guardar_configuracion("", "")
        with self.assertRaises(correo.CorreoError):
            correo.iniciar_conexion()

    @mock.patch("seguimiento.correo.time.sleep")
    @mock.patch("seguimiento.correo.requests")
    def test_conectar_con_codigo(self, req, _):
        req.post.side_effect = [
            Respuesta(200, {"device_code": "d", "user_code": "ABC", "interval": 1,
                            "verification_uri": "https://microsoft.com/devicelogin"}),
            Respuesta(400, {"error": "authorization_pending"}),
            Respuesta(200, {"access_token": "acceso", "refresh_token": "renovar",
                            "expires_in": 3600})]
        req.get.return_value = Respuesta(200, {"mail": "persona@ejemplo.invalid"})
        codigo = correo.iniciar_conexion()
        self.assertEqual(codigo["user_code"], "ABC")
        self.assertEqual(correo.completar_conexion("u1", codigo), "persona@ejemplo.invalid")
        self.assertEqual(correo.conexion("u1"), "persona@ejemplo.invalid")
        self.assertEqual(almacen.outlook_sgc("u1")["token"], "texto:renovar")

    @mock.patch("seguimiento.correo.requests")
    def test_codigo_caducado(self, req):
        req.post.return_value = Respuesta(400, {"error": "expired_token"})
        with self.assertRaises(correo.CorreoError):
            correo.completar_conexion("u1", {"device_code": "d", "interval": 1})

    @mock.patch("seguimiento.correo.requests")
    def test_envio_directo_renovando_el_token(self, req):
        almacen.guardar_outlook_sgc("u1", "persona@ejemplo.invalid", "texto:renovar")
        req.post.side_effect = [Respuesta(200, {"access_token": "nuevo", "expires_in": 3600}),
                                Respuesta(202)]
        correo.enviar("u1", "a@ejemplo.invalid; b@ejemplo.invalid", "", "Asunto", "<p>x</p>",
                      [("informe.pdf", b"%PDF")])
        llamada = req.post.call_args_list[1]
        self.assertTrue(llamada.args[0].endswith("/me/sendMail"))
        mensaje = llamada.kwargs["json"]["message"]
        self.assertEqual(len(mensaje["toRecipients"]), 2)
        self.assertEqual(mensaje["attachments"][0]["contentType"], "application/pdf")
        self.assertEqual(llamada.kwargs["headers"]["Authorization"], "Bearer nuevo")

    @mock.patch("seguimiento.correo.requests")
    def test_envio_grande_con_sesion_de_subida(self, req):
        correo._tokens["u1"] = ("acceso", 1e12)
        grande = b"x" * (correo.LIMITE_DIRECTO + 10)
        req.post.side_effect = [Respuesta(201, {"id": "m1"}),
                                Respuesta(201, {"uploadUrl": "https://subida.invalid"}),
                                Respuesta(202)]
        req.put.return_value = Respuesta(201)
        correo.enviar("u1", "a@ejemplo.invalid", "", "Asunto", "<p>x</p>", [("g.pdf", grande)])
        self.assertEqual(req.put.call_count, 1 + len(grande) // correo.TROZO)
        self.assertTrue(req.post.call_args_list[-1].args[0].endswith("/me/messages/m1/send"))

    @mock.patch("seguimiento.correo.requests")
    def test_token_retirado_desconecta(self, req):
        almacen.guardar_outlook_sgc("u1", "persona@ejemplo.invalid", "texto:renovar")
        req.post.return_value = Respuesta(400, {"error": "invalid_grant"})
        with self.assertRaises(correo.CorreoError):
            correo.enviar("u1", "a@ejemplo.invalid", "", "Asunto", "<p>x</p>")
        self.assertIsNone(correo.conexion("u1"))

    def test_sin_destinatarios(self):
        with self.assertRaises(correo.CorreoError):
            correo.enviar("u1", "", "", "Asunto", "<p>x</p>")


if __name__ == "__main__":
    unittest.main()
