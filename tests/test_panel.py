# -*- coding: utf-8 -*-
"""Mi espacio: fecha clave y semáforo de las licitaciones, su estado según solicitudes y
análisis, carga por meses, informes y facturas pendientes según la periodicidad (regla de
retraso M+2), calendario, ficha de cliente, bloques del panel y preferencias. Datos ficticios."""

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
from mi_espacio import datos as P
from seguimiento import columnas as C

D = dt.date
U = P.umbrales()


def sum_(id_, **d):
    base = {"Técnico": "Ana", "Auxiliar": "Eva", "Cliente/Grupo": "Cliente A",
            "CUPS": f"ES00000000000000{id_:04d}XX", "Denominación suministro": f"Sum {id_}",
            "Tipo contratación": C.SGTO_Y_CONTROL, "Cadencia informe": "Mensual",
            "Tarifa de acceso": "6.1TD", "Consumo anual (kWh)": 100000}
    base.update(d)
    return {"id": id_, "datos": base}


EQUIPO = [{"nombre": "Ana", "tecnico": 1, "auxiliar": 0, "usuario": "ana"},
          {"nombre": "Eva", "tecnico": 0, "auxiliar": 1, "usuario": "eva"},
          {"nombre": "Luis", "tecnico": 1, "auxiliar": 0, "usuario": "luis"}]


class FechaClave(unittest.TestCase):
    def test_prioridad_y_tipo(self):
        d = {P.F_LANZAMIENTO: "2026-05-01", P.F_DENUNCIA: "2026-06-01", P.F_FIN: "2026-09-30"}
        self.assertEqual(P.fecha_clave(d), (D(2026, 5, 1), "Lanzamiento"))
        self.assertEqual(P.fecha_clave(d, "Denuncia")[0], D(2026, 6, 1))
        self.assertEqual(P.fecha_clave(d, "Fin de contrato")[0], D(2026, 9, 30))
        self.assertEqual(P.fecha_clave({P.F_DENUNCIA: "2026-06-01"})[1], "Denuncia")
        f, origen = P.fecha_clave({P.F_FIN: "2026-09-30"}, margen=3)
        self.assertEqual((f, origen), (D(2026, 6, 30), "Fin de contrato − 3 meses"))
        self.assertEqual(P.fecha_clave({}), (None, ""))

    def test_semaforo(self):
        self.assertEqual(P.semaforo(10, U), P.ROJO)
        self.assertEqual(P.semaforo(30, U), P.AMBAR)
        self.assertEqual(P.semaforo(90, U), P.AMBAR)
        self.assertEqual(P.semaforo(91, U), P.VERDE)
        self.assertEqual(P.semaforo(-5, U), P.GRIS)                 # pasó sin solicitud
        self.assertEqual(P.semaforo(-5, U, hay_trabajo=True), P.ROJO)
        propios = P.umbrales({"rojo": 15}, {"ambar": 60, "inf_rojo": ""})
        self.assertEqual((propios["rojo"], propios["ambar"], propios["inf_rojo"]), (15, 60, 15))
        self.assertEqual(P.semaforo(20, propios), P.AMBAR)

    def test_fechas_de_texto(self):
        self.assertEqual(P.fecha("15/11/2026 14:00"), D(2026, 11, 15))
        self.assertEqual(P.fecha("15/11/2026"), D(2026, 11, 15))
        self.assertEqual(P.fecha("2026-11-15T10:00:00"), D(2026, 11, 15))
        self.assertIsNone(P.fecha("NO PRÓRROGA"))


class Licitaciones(unittest.TestCase):
    def fuente(self):
        return P.Fuente(suministros=[
            sum_(1, **{P.F_DENUNCIA: "2026-06-01"}),
            sum_(2, **{P.F_DENUNCIA: "2026-06-01"}),                      # misma licitación
            sum_(3, **{"Cliente/Grupo": "Cliente B", "Técnico": "Luis",
                       P.F_DENUNCIA: "2026-07-15"}),
            sum_(4, **{"Cliente/Grupo": "Cliente C", P.F_FIN: "2026-03-31"}),
            sum_(5, **{"Cliente/Grupo": "Cliente D"})],                    # sin fechas
            equipo=EQUIPO)

    def test_agrupar_estado_y_carga(self):
        f = self.fuente()
        hoy = D(2026, 5, 15)
        f.solicitudes = [{"id": 9, "fecha": "2026-04-01T10:00:00", "estado": "Enviada",
                          "cups": "ES000000000000000001XX, ES000000000000000002XX"}]
        f.analisis = [{"id": 4, "fecha": "2025-01-10T10:00:00", "estado": "Adjudicada",
                       "cups": "ES000000000000000003XX"}]           # anterior a la ventana
        lic = P.agrupar_licitaciones(P.licitaciones(f, hoy, U), f, U)
        self.assertEqual(len(lic), 3)
        a = lic[lic["Cliente/Grupo"] == "Cliente A"].iloc[0]
        self.assertEqual((a["Suministros"], a["Estado"], a["_id_trabajo"]),
                         (2, "Solicitud enviada", 9))
        self.assertEqual((a["Días"], a["Semáforo"]), (17, P.ROJO))
        b = lic[lic["Cliente/Grupo"] == "Cliente B"].iloc[0]
        self.assertEqual((b["Estado"], b["Semáforo"]), ("Sin solicitud", P.AMBAR))
        c = lic[lic["Cliente/Grupo"] == "Cliente C"].iloc[0]
        self.assertEqual((c["Fecha clave"], c["Semáforo"]), (D(2025, 12, 31), P.GRIS))
        carga = P.carga_por_mes(lic, por_tecnico=True)
        self.assertEqual(int(carga["Licitaciones"].sum()), 3)
        self.assertEqual(set(carga["Técnico"]), {"Ana", "Luis"})

    def test_analisis_manda_sobre_la_solicitud(self):
        cups = {"ES000000000000000001XX"}
        sols = [{"id": 1, "fecha": "2026-03-01", "estado": "Ofertas recibidas",
                 "cups": "ES000000000000000001XX"}]
        ans = [{"id": 2, "fecha": "2026-04-01", "estado": "En curso",
                "cups": "ES000000000000000001XX"}]
        self.assertEqual(P.estado_licitacion(cups, D(2026, 6, 1), sols, ans)[:2],
                         ("Análisis en curso", "analisis"))
        ans[0]["estado"] = "Adjudicada"
        self.assertEqual(P.estado_licitacion(cups, D(2026, 6, 1), sols, ans)[0], "Adjudicada")
        self.assertEqual(P.estado_licitacion(cups, D(2026, 6, 1), sols, [])[0],
                         "Ofertas recibidas")

    def test_mis_suministros(self):
        n = P.mis_nombres("ana", EQUIPO)
        self.assertEqual(n, {"tecnico": {"Ana"}, "auxiliar": set()})
        f = self.fuente()
        filas = P.datos_bloque("licitaciones", f, D(2026, 5, 15), U, "ana", n)
        # Cliente C pasó hace más de 30 días: no sale en el panel (sí en Próximas licitaciones)
        self.assertEqual({x["Cliente/Grupo"] for x in filas}, {"Cliente A"})
        todas = P.datos_bloque("licitaciones", f, D(2026, 5, 15), U, "ana", n, todos=True)
        self.assertEqual(len(todas), 2)
        sin = P.datos_bloque("sin_fechas", f, D(2026, 5, 15), U, "ana", n)
        self.assertEqual({x["CUPS"] for x in sin},
                         {f"ES00000000000000{i:04d}XX" for i in (1, 2, 4, 5)})


class Informes(unittest.TestCase):
    """Regla de retraso.py: bloques naturales y control el día 1 de M+2."""

    def fuente(self, cadencia, estados):
        s = sum_(1, **{"Cadencia informe": cadencia})
        return P.Fuente(suministros=[s], estados={(1, 2026, m): {"estado": v}
                                                  for m, v in estados.items()})

    def test_mensual(self):
        f = self.fuente("Mensual", {1: "Pdte inf", 2: "Sin fra", 3: "Sin fra"})
        inf, fra = P.informes_pendientes(f, D(2026, 2, 20), U)
        self.assertEqual([(x["Bloque"], x["Fecha de control"], x["Días"], x["Semáforo"])
                          for x in inf], [("ene-26", D(2026, 3, 1), 9, P.ROJO)])
        inf, _ = P.informes_pendientes(f, D(2026, 3, 2), U)
        self.assertEqual(inf[0]["Semáforo"], P.RETRASO)
        # la factura de febrero falta desde marzo; la de marzo aún no ha acabado
        _, fra = P.informes_pendientes(f, D(2026, 3, 10), U)
        self.assertEqual([x["Meses sin factura"] for x in fra], ["feb-26"])
        self.assertEqual(fra[0]["Fecha de control"], D(2026, 4, 1))

    def test_trimestral_semestral_y_anual(self):
        hoy = D(2026, 4, 20)
        for cad, control, sem in (("Trimestral", D(2026, 5, 1), P.ROJO),
                                  ("Semestral", D(2026, 8, 1), P.VERDE),
                                  ("Anual", D(2027, 2, 1), P.VERDE)):
            f = self.fuente(cad, {1: "Recibida", 2: "Inf enviado"})
            inf, _ = P.informes_pendientes(f, hoy, U)
            self.assertEqual((inf[0]["Fecha de control"], inf[0]["Semáforo"],
                              inf[0]["Meses pendientes"]), (control, sem, "ene-26"), cad)
        f = self.fuente("Trimestral", {1: "Recibida"})
        self.assertEqual(P.informes_pendientes(f, D(2026, 4, 5), U)[0][0]["Semáforo"], P.AMBAR)

    def test_sin_informes_esperados_y_otros_tipos(self):
        f = self.fuente("Mensual", {1: "Inf enviado", 2: "Cadencia", 3: "Temporal"})
        self.assertEqual(P.informes_pendientes(f, D(2026, 6, 1), U)[0], [])
        f.suministros[0]["datos"]["Tipo contratación"] = "Otro"
        f.estados[(1, 2026, 4)] = {"estado": "Pdte inf"}
        self.assertEqual(P.informes_pendientes(f, D(2026, 6, 1), U), ([], []))

    def test_cambio_de_anio_y_anio_sin_estados(self):
        f = self.fuente("Mensual", {})
        f.estados = {(1, 2025, 12): {"estado": "Recibida"}, (1, 2025, 11): {"estado": "Inf enviado"}}
        inf, fra = P.informes_pendientes(f, D(2026, 1, 20), U)
        self.assertEqual([(x["Bloque"], x["Fecha de control"]) for x in inf],
                         [("dic-25", D(2026, 2, 1))])
        # 2025 tiene estados: los meses vacíos de 2025 cuentan como sin factura; 2026 no
        self.assertTrue(all(x["Bloque"].endswith("-25") for x in fra))
        self.assertIn("ene-25", fra[0]["Meses sin factura"])

    def test_solo_los_del_auxiliar(self):
        f = self.fuente("Mensual", {1: "Pdte inf"})
        self.assertEqual(P.informes_pendientes(f, D(2026, 2, 20), U,
                                               P.mis_nombres("luis", EQUIPO)), ([], []))
        self.assertEqual(len(P.informes_pendientes(f, D(2026, 2, 20), U,
                                                   P.mis_nombres("eva", EQUIPO))[0]), 1)


class Calendario(unittest.TestCase):
    def test_eventos(self):
        f = P.Fuente(
            suministros=[sum_(1, **{P.F_DENUNCIA: "2026-06-10", P.F_FIN: "2026-09-30",
                                    P.F_LANZAMIENTO: "2026-06-20"})],
            estados={(1, 2026, 4): {"estado": "Pdte inf"}},
            solicitudes=[{"id": 3, "recepcion_hasta": "15/06/2026 14:00", "estado": "Enviada",
                          "ref": "R1", "cliente": "Cliente A", "cups": "", "usuario": "ana"},
                         {"id": 4, "recepcion_hasta": "16/06/2026", "estado": "Cancelada",
                          "ref": "R2", "cliente": "X", "cups": "", "usuario": "ana"}],
            catalogo=[{"id": 7, "vigencia_hasta": "2026-06-30", "comercializadora": "ENDESA",
                       "nombre": "Fijo", "usuario": "ana"}])
        evs = P.eventos(f, D(2026, 6, 1), D(2026, 6, 30), U, D(2026, 5, 20))
        self.assertEqual([e["Tipo"] for e in evs],
                         ["Control de informe", "Denuncia", "Fin de recepción de ofertas",
                          "Lanzamiento previsto", "Caduca oferta de catálogo"])
        self.assertEqual(evs[0]["Fecha"], D(2026, 6, 1))           # abril → control 1 de junio


class Ficha(unittest.TestCase):
    def test_buscar_y_reunir(self):
        f = P.Fuente(
            suministros=[sum_(1, CIF="B1"), sum_(2, CIF="B1"), sum_(3, **{"Cliente/Grupo": "Otro"})],
            solicitudes=[{"id": 1, "cliente": "Cliente A", "cups": ""},
                         {"id": 2, "cliente": "Z", "cups": "ES000000000000000003XX"}],
            analisis=[{"id": 5, "cliente": "Nadie", "cups": "ES000000000000000002XX"}],
            estudios=[{"id": 8, "titular": "Cliente A", "cups": ""}],
            revisiones=[{"id": 9, "cups": "ES000000000000000001XX"}],
            fichas={"ES000000000000000001XX": ("2026-01-01", "ana")})
        enc = P.buscar(f, "CIF", "B1")
        self.assertEqual(len(enc["suministros"]), 2)
        todo = P.ficha(f, enc)
        self.assertEqual([r["id"] for r in todo["Solicitudes"]], [1])
        self.assertEqual([r["id"] for r in todo["Análisis"]], [5])
        self.assertEqual(len(todo["Estudios de potencia"]), 1)
        self.assertEqual(len(todo["Revisiones de SSAA"]), 1)
        self.assertEqual(len(todo["Fichas de contrato"]), 1)
        enc = P.buscar(f, "CUPS", "ES000000000000000003XX")
        self.assertEqual([r["id"] for r in P.ficha(f, enc)["Solicitudes"]], [2])
        self.assertTrue(any(t == "CIF" for _, (t, _v) in P.opciones_busqueda(f)))


class Panel(unittest.TestCase):
    def test_config_por_rol_y_personalizada(self):
        tec = P.config_panel("Técnico")
        visibles = [c["bloque"] for c in tec if c["visible"]]
        self.assertIn("licitaciones", visibles)
        self.assertNotIn("informes", visibles)
        aux = [c["bloque"] for c in P.config_panel("Auxiliar") if c["visible"]]
        self.assertIn("informes", aux)
        self.assertNotIn("licitaciones", aux)
        ambos = [c["bloque"] for c in P.config_panel("Técnico y auxiliar") if c["visible"]]
        self.assertTrue({"informes", "licitaciones"} <= set(ambos))
        propia = P.config_panel("Técnico", [{"bloque": "informes", "visible": True, "filas": 3},
                                            {"bloque": "licitaciones", "visible": False}])
        self.assertEqual(propia[0], {"bloque": "informes", "visible": True, "filas": 3})
        self.assertFalse(propia[1]["visible"])
        self.assertEqual(len(propia), len(P.BLOQUES))           # los demás, ocultos al final
        self.assertEqual(set(P.contador_por_defecto("Auxiliar")),
                         {"informes", "facturas", "gemweb", "sin_contacto"})

    def test_avisos_solo_lo_elegido(self):
        f = P.Fuente(suministros=[sum_(1, **{P.F_DENUNCIA: "2026-05-20"})],
                     estados={(1, 2026, 1): {"estado": "Pdte inf"}}, equipo=EQUIPO)
        hoy = D(2026, 5, 10)
        n = P.mis_nombres("ana", EQUIPO)
        self.assertEqual(P.avisos(f, hoy, U, "ana", n, ["licitaciones"]), {"licitaciones": 1})
        self.assertEqual(P.avisos(f, hoy, U, "ana", n, ["licitaciones", "informes"],
                                  todos=True), {"licitaciones": 1, "informes": 1})

    def test_gemweb_y_sin_contacto(self):
        f = P.Fuente(suministros=[sum_(1), sum_(2, **{"Cliente/Grupo": "Cliente B"})],
                     cambios_gemweb=[{"id_suministro": 1, "cups": "X", "anio": 2026, "mes": 2,
                                      "antes": "Sin fra", "despues": "Recibida"}],
                     contactos_envio=[{"datos": {"Cliente/Grupo": "Cliente A", "Para": "a@b.es"}}])
        n = P.mis_nombres("eva", EQUIPO)
        self.assertEqual(len(P.datos_bloque("gemweb", f, D(2026, 3, 1), U, "eva", n)), 1)
        sin = P.datos_bloque("sin_contacto", f, D(2026, 3, 1), U, "eva", n)
        self.assertEqual([x["Cliente/Grupo"] for x in sin], ["Cliente B"])


class Preferencias(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def test_rol_y_preferencias(self):
        almacen.crear_usuario("eva", "Eva", "clave-larga-1", cambiar_clave=False,
                              rol="Auxiliar")
        almacen.crear_usuario("ana", "Ana", "clave-larga-2", cambiar_clave=False)
        roles = {u["usuario"]: u["rol"] for u in almacen.listar_usuarios()}
        self.assertEqual(roles, {"eva": "Auxiliar", "ana": "Técnico"})
        self.assertEqual(almacen.comprobar("eva", "clave-larga-1")["rol"], "Auxiliar")
        almacen.actualizar_usuario("ana", rol="Técnico y auxiliar")
        self.assertEqual(almacen.roles_de("Técnico y auxiliar"), {"Técnico", "Auxiliar"})
        self.assertEqual(almacen.roles_de(None), {"Técnico"})
        self.assertIsNone(almacen.leer_preferencia("ana", "panel"))
        almacen.guardar_preferencia("ana", "favoritos", ["licitaciones", "seguimiento"])
        almacen.guardar_preferencia("ana", "panel", [{"bloque": "informes", "visible": True}])
        self.assertEqual(almacen.leer_preferencia("ana", "favoritos"),
                         ["licitaciones", "seguimiento"])
        self.assertEqual(almacen.leer_preferencia("eva", "favoritos", []), [])
        almacen.borrar_preferencia("ana", "panel")
        self.assertIsNone(almacen.leer_preferencia("ana", "panel"))


class Paginas(unittest.TestCase):
    """Las páginas de Mi espacio se abren sin errores con cada rol (base temporal)."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()
        hoy = dt.date.today()
        almacen.guardar_sgc_suministros([
            {"id": None, "datos": dict(sum_(i)["datos"], **{
                P.F_DENUNCIA: (hoy + dt.timedelta(days=20 * i)).isoformat()})}
            for i in range(1, 4)], "ana")
        almacen.guardar_sgc_equipo([{"nombre": "Ana", "tecnico": 1, "usuario": "ana"},
                                    {"nombre": "Eva", "auxiliar": 1, "usuario": "eva"}])

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def test_se_abren(self):
        from streamlit.testing.v1 import AppTest

        def app(pagina, usuario, rol):
            import streamlit as st
            from mi_espacio import paginas as mi
            st.session_state.setdefault("usuario", {"usuario": usuario, "nombre": usuario,
                                                    "admin": 0, "supervisor": 0, "rol": rol})
            getattr(mi, pagina)()

        for pagina in ("pagina_mi_panel", "pagina_licitaciones", "pagina_ficha",
                       "pagina_calendario", "seccion_ajustes"):
            for usuario, rol in (("ana", "Técnico"), ("eva", "Auxiliar")):
                at = AppTest.from_function(app, args=(pagina, usuario, rol),
                                           default_timeout=120).run()
                self.assertEqual([e.value for e in at.exception], [], (pagina, rol))
        at = AppTest.from_function(app, args=("pagina_mi_panel", "ana", "Técnico"),
                                   default_timeout=120).run()
        textos = " ".join(m.value for m in at.markdown)
        self.assertIn("Mis próximas licitaciones", textos)
        self.assertNotIn("Informes por vencer", textos)


if __name__ == "__main__":
    unittest.main()
