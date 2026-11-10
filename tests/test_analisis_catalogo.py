# -*- coding: utf-8 -*-
"""Catálogo de ofertas: qué se guarda (solo lo genérico), validación, carga en otro análisis
con sus avisos, versiones y permisos en la base de datos (carpeta temporal). Datos ficticios."""

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import almacen
import config
from analisis_ofertas import calculo, catalogo
from analisis_ofertas import modelo as mo
from test_analisis_calculo import CUPS_A, estudio, fijo, oferta

FECHAS = dict(vigencia_desde=dt.date(2026, 11, 1), vigencia_hasta=dt.date(2026, 11, 30),
              inicio_desde=dt.date(2027, 1, 1), inicio_hasta=dt.date(2027, 3, 1))


def oferta_cliente():
    """Fija con cosas propias del cliente (gestor, grupo propio, CUPS excluido, inicio)."""
    o = fijo(0.12, comercializadora="ENDESA ENERGÍA, S.A.U.", nombre="Endesa 12 M",
             gestor="Gestor de Prueba", email_gestor="gestor@ejemplo.com", k=0.015,
             duracion=12, inicio=dt.date(2027, 2, 1), ronda=2)
    o.textos = {"k": "1,5"}
    o.grupos_propios = [mo.GrupoPropio("6.1TD grandes", "6.1TD")]
    o.precios["6.1TD grandes"] = [mo.PreciosAnio(energia=[0.11] * 6)]
    o.asignacion = {CUPS_A: "6.1TD grandes"}
    o.excluidos = ["ES0000000000000000ZZ"]
    o.condiciones = {"validez": "Hasta el 30/11/2026"}
    return o


def entrada(o, **kw):
    ficha = catalogo.ficha_de(o, kw.pop("nombre", "Endesa Empresas nov-26"), **FECHAS)
    d = {"id": 7, "version": 1, "estado": catalogo.ACTIVA, **ficha,
         "datos": catalogo.a_catalogo(o)}
    d.update(kw)
    return d


class Guardar(unittest.TestCase):
    def test_solo_lo_generico(self):
        d = catalogo.a_catalogo(oferta_cliente())
        for campo in ("gestor", "email_gestor", "inicio", "ronda", "cierres", "grupos_propios",
                      "asignacion", "excluidos", "apuntamientos_propios", "nombre"):
            self.assertNotIn(campo, d)
        self.assertEqual(list(d["precios"]), ["6.1TD"])          # sin el grupo propio
        self.assertEqual(d["k"], 0.015)
        self.assertEqual(d["textos"], {"k": "1,5"})              # decimales como se escribieron
        self.assertEqual(d["condiciones"]["validez"], "Hasta el 30/11/2026")

    def test_validar(self):
        o = oferta_cliente()
        ficha = catalogo.ficha_de(o, "Endesa nov-26", **FECHAS)
        self.assertEqual(catalogo.validar(ficha, catalogo.a_catalogo(o)), [])
        mal = dict(ficha, nombre="", vigencia_hasta=None,
                   inicio_desde=dt.date(2027, 5, 1))
        errores = " ".join(catalogo.validar(mal, catalogo.a_catalogo(o)))
        for texto in ("nombre", "vigencia", "posterior"):
            self.assertIn(texto, errores)
        sin_precios = oferta(modalidad=mo.FIJO, comercializadora="X")
        self.assertTrue(any("precios" in x for x in catalogo.validar(
            catalogo.ficha_de(sin_precios, "X", **FECHAS), catalogo.a_catalogo(sin_precios))))
        # Pass-Through con solo fee: no hacen falta precios por tarifa
        pt = oferta(modalidad=mo.PASS_THROUGH, fee=2.0, comercializadora="X")
        self.assertEqual(catalogo.validar(catalogo.ficha_de(pt, "X", **FECHAS),
                                          catalogo.a_catalogo(pt)), [])


class Cargar(unittest.TestCase):
    def test_ida_y_vuelta_calcula_igual(self):
        e = estudio(n_cups=2)
        original = fijo(0.12, k=0.015, gdo=1.0, comercializadora="ENDESA")
        original.gestor, original.email_gestor = "Gestor de otro cliente", "otro@ejemplo.com"
        cargada = catalogo.desde_catalogo(entrada(original))
        ctx = calculo.Contexto(e)
        self.assertAlmostEqual(calculo.calcular_oferta(e, original, ctx).total,
                               calculo.calcular_oferta(e, cargada, ctx).total, places=6)
        self.assertEqual(cargada.origen_catalogo, {"id": 7, "version": 1,
                                                   "nombre": "Endesa Empresas nov-26"})
        self.assertEqual((cargada.gestor, cargada.email_gestor), ("", ""))   # siempre vacío
        self.assertEqual(cargada.grupos_propios, [])
        self.assertIsNone(cargada.inicio)

    def test_avisos(self):
        e = estudio()
        ent = entrada(oferta_cliente())
        self.assertEqual(catalogo.avisos_carga(ent, e, dt.date(2026, 11, 15)), [])
        avisos = " ".join(catalogo.avisos_carga(ent, e, dt.date(2026, 12, 15)))
        self.assertIn("caducó", avisos)
        e.inicio = dt.date(2027, 6, 1)
        self.assertTrue(any("ventana de inicio" in x
                            for x in catalogo.avisos_carga(ent, e, dt.date(2026, 11, 15))))
        baja = dict(ent, estado=catalogo.BAJA)
        self.assertTrue(any("baja" in x for x in catalogo.avisos_carga(baja, estudio(),
                                                                       dt.date(2026, 11, 15))))

    def test_cubre_tarifas_y_grupos_territoriales(self):
        e = estudio(n_cups=2)
        e.suministros[1].tarifa = "3.0TD"
        ent = entrada(oferta_cliente())
        self.assertFalse(catalogo.cubre(ent, e))
        self.assertEqual(catalogo.sin_precio(ent, e), ["3.0TD"])
        self.assertTrue(any("3.0TD" in x for x in catalogo.avisos_carga(ent, e,
                                                                        dt.date(2026, 11, 15))))
        # un CUPS de Canarias usa los precios de su tarifa si no hay grupo territorial
        e.suministros[1].tarifa, e.suministros[1].sistema = "6.1TD", "Canarias"
        self.assertTrue(catalogo.cubre(ent, e))
        pt = oferta(modalidad=mo.PASS_THROUGH, fee=2.0)
        self.assertTrue(catalogo.cubre(entrada(pt), e))

    def test_versiones_y_actualizar(self):
        e = estudio()
        o = catalogo.desde_catalogo(entrada(oferta_cliente()))
        o.nombre, o.gestor = "Endesa (mi nombre)", "Mi gestor"
        o.grupos_propios = [mo.GrupoPropio("Propio", "6.1TD")]
        o.precios["Propio"] = [mo.PreciosAnio(energia=[0.2] * 6)]
        e.ofertas = [o, fijo(0.1)]
        self.assertEqual(catalogo.desactualizadas(e, {7: 1}), [])
        self.assertEqual([v for _, v in catalogo.desactualizadas(e, {7: 2})], [2])
        nueva = oferta_cliente()
        nueva.precios["6.1TD"] = [mo.PreciosAnio(energia=[0.13] * 6)]
        catalogo.actualizar_desde(o, entrada(nueva, version=2))
        self.assertEqual(o.precios["6.1TD"][0].energia, [0.13] * 6)
        self.assertIn("Propio", o.precios)                        # lo del análisis se queda
        self.assertEqual((o.nombre, o.gestor), ("Endesa (mi nombre)", "Mi gestor"))
        self.assertEqual(o.origen_catalogo["version"], 2)

    def test_duplicar_conserva_el_origen(self):
        o = catalogo.desde_catalogo(entrada(oferta_cliente()))
        self.assertEqual(mo.duplicar(o).origen_catalogo["id"], 7)


class BaseDeDatos(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.p = mock.patch.object(config, "CARPETA_DATOS", Path(self.tmp.name))
        self.p.start()
        almacen._iniciado.clear()
        almacen.crear_usuario("ana", "Ana", "clave-larga-1", cambiar_clave=False)
        almacen.crear_usuario("luis", "Luis", "clave-larga-2", cambiar_clave=False)
        almacen.crear_usuario("jefa", "Jefa", "clave-larga-3", admin=True, cambiar_clave=False)

    def tearDown(self):
        self.p.stop()
        almacen._iniciado.clear()
        self.tmp.cleanup()

    def _guardar(self, **kw):
        o = oferta_cliente()
        ficha = catalogo.ficha_de(o, "Endesa nov-26", **FECHAS)
        ficha.update(kw)
        return almacen.guardar_catalogo("ana", ficha, catalogo.a_catalogo(o)), ficha

    def test_guardar_listar_y_cargar(self):
        id_, _ = self._guardar()
        self._guardar(nombre="Endesa oct-26", vigencia_desde=dt.date(2026, 10, 1),
                      vigencia_hasta=dt.date(2026, 10, 31))
        self.assertEqual(len(almacen.listar_catalogo()), 2)
        vig = almacen.listar_catalogo(vigentes_en=dt.date(2026, 11, 15))
        self.assertEqual([r["nombre"] for r in vig], ["Endesa nov-26"])
        self.assertEqual(vig[0]["tarifas"], ["6.1TD"])
        self.assertNotIn("datos", vig[0])
        r = almacen.catalogo(id_)
        self.assertEqual(r["datos"]["k"], 0.015)
        o = catalogo.desde_catalogo(r)
        self.assertEqual(o.origen_catalogo["id"], id_)
        self.assertEqual(len(almacen.listar_catalogo(comercializadora="endesa energía, s.a.u.")), 2)

    def test_permisos_y_versiones(self):
        id_, ficha = self._guardar()
        datos = almacen.catalogo(id_)["datos"]
        self.assertFalse(almacen.actualizar_catalogo(id_, "luis", ficha, datos))
        self.assertTrue(almacen.actualizar_catalogo(id_, "ana", ficha, datos))
        self.assertTrue(almacen.actualizar_catalogo(id_, "jefa", ficha, datos))
        self.assertEqual(almacen.versiones_catalogo([id_]), {id_: 3})
        # cambiar solo la cabecera no sube la versión
        self.assertTrue(almacen.actualizar_catalogo(id_, "ana", dict(ficha, nombre="Otro")))
        r = almacen.catalogo(id_)
        self.assertEqual((r["version"], r["nombre"], r["actualizado_por"]), (3, "Otro", "ana"))
        self.assertFalse(almacen.estado_catalogo(id_, "luis", catalogo.BAJA))
        self.assertTrue(almacen.estado_catalogo(id_, "ana", catalogo.BAJA))
        self.assertEqual(almacen.listar_catalogo(), [])
        self.assertEqual(len(almacen.listar_catalogo(incluir_bajas=True)), 1)
        self.assertFalse(almacen.borrar_catalogo(id_, "luis"))
        self.assertTrue(almacen.borrar_catalogo(id_, "jefa"))
        self.assertIsNone(almacen.catalogo(id_))

    def test_guardar_y_cargar_desde_las_paginas(self):
        from analisis_ofertas import paginas
        e = estudio()
        o12 = oferta_cliente()
        o24 = fijo(0.115, comercializadora=o12.comercializadora, nombre="Endesa 24 M",
                   duracion=24)
        mensajes, errores = paginas.guardar_en_catalogo(e, [o12, o24], "Endesa nov-26", FECHAS,
                                                        "luis", id_analisis=5)
        self.assertEqual(errores, [])
        filas = almacen.listar_catalogo()
        self.assertEqual(sorted(r["nombre"] for r in filas),
                         ["Endesa nov-26 · 12 M", "Endesa nov-26 · 24 M"])
        self.assertEqual({r["id_analisis_origen"] for r in filas}, {5})
        self.assertEqual(o12.origen_catalogo["version"], 1)
        # otro usuario lo carga en su análisis
        otro = estudio(n_cups=2)
        otro.suministros[1].tarifa = "3.0TD"
        avisos = paginas.cargar_del_catalogo(otro, [r["id"] for r in filas],
                                             hoy=dt.date(2026, 11, 15), usuario="ana")
        self.assertEqual(len(otro.ofertas), 2)
        self.assertTrue(all(x.gestor == "" for x in otro.ofertas))
        self.assertTrue(any("3.0TD" in x for x in avisos))
        df = paginas.tabla_catalogo(filas, otro)
        self.assertTrue(df["Cubre mis tarifas"].str.startswith("⚠").all())
        # versión nueva: el autor sí; otro usuario guarda una entrada nueva
        o12.precios["6.1TD"] = [mo.PreciosAnio(energia=[0.125] * 6)]
        paginas.guardar_en_catalogo(e, [o12], "Endesa nov-26 · 12 M", FECHAS, "luis",
                                    como_version=True)
        self.assertEqual(o12.origen_catalogo["version"], 2)
        self.assertEqual(len(almacen.listar_catalogo()), 2)
        cargada = otro.ofertas[0]
        paginas.guardar_en_catalogo(otro, [cargada], "Copia de Ana", FECHAS, "ana",
                                    como_version=True)
        self.assertEqual(len(almacen.listar_catalogo()), 3)
        # errores: sin fechas no se guarda
        _, errores = paginas.guardar_en_catalogo(e, [o24], "X", dict(FECHAS, vigencia_hasta=None),
                                                 "luis")
        self.assertTrue(errores)

    def test_supervision(self):
        from supervision import datos as sd
        self._guardar()
        filas = sd.trabajo(dt.date(2000, 1, 1), dt.date(2100, 1, 1))
        self.assertTrue(any(f["origen"] == "catalogo" and f["usuario"] == "ana" for f in filas))


if __name__ == "__main__":
    unittest.main()
