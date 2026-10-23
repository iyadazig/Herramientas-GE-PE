# -*- coding: utf-8 -*-
"""Modulo «Estudio de potencia» (estudio_potencia/servicio.py): zona por codigo postal, curvas
subidas, descarga de Gemweb simulada, calculo, tablas, PDF y estudios guardados. Datos
ficticios (tests/generador_curvas.py)."""


import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import generador_curvas as g
from estudio_potencia import servicio as sv

CUPS_A = "ES0000000000000001AA0F"
CUPS_B = "ES0000000000000002BB0F"
ACTUAL = [150, 150, 150, 150, 150, 200]


def suministro_calculado(tmp, zona="Península"):
    Path(tmp).mkdir(parents=True, exist_ok=True)
    g.escribir_p5d(g.curva_fisica(), Path(tmp) / "p5d.txt")
    rutas = sv.guardar_subidos([("p5d.txt", (Path(tmp) / "p5d.txt").read_bytes())],
                               Path(tmp) / "subidos")
    s = sv.Suministro(zona=zona)
    sv.leer_suministro(s, sv.abrir(rutas))
    s.actual = list(ACTUAL)
    assert sv.calcular(s, sv.precios()) is None
    return s


class Zona(unittest.TestCase):
    def test_por_codigo_postal(self):
        for cp, zona in (("07001", "Illes Balears"), ("7001", "Illes Balears"),
                         ("35001", "Canarias"), ("38001", "Canarias"), ("51001", "Ceuta"),
                         ("52001", "Melilla"), ("41012", "Península"), ("28037", "Península")):
            self.assertEqual(sv.zona_por_codigo_postal(cp), zona)
        self.assertIsNone(sv.zona_por_codigo_postal(None))

    def test_zona_del_suministro_de_gemweb(self):
        self.assertEqual(sv.zona_de({"codi_postal": "35001"}), ("Canarias", "CP 35001, Gemweb"))
        # sin codigo postal en el suministro: el del centro de consumo
        self.assertEqual(sv.zona_de({"codi_postal": ""}, {"codi_postal": "07001"}),
                         ("Illes Balears", "CP 07001, centro de consumo en Gemweb"))
        zona, origen = sv.zona_de({"codi_postal": None})
        self.assertEqual(zona, "Península")
        self.assertIn("revisa la zona", origen)
        self.assertIn("no está en Gemweb", sv.zona_de(None)[1])

    def test_sin_codigo_postal_lee_la_direccion_y_la_poblacion(self):
        # codigo postal escrito en la direccion del centro («-» en su campo)
        centro = {"codi_postal": "-", "direccio": "CM FINCA 1 CM 35110 SANTA LUCIA (LAS PALMAS)",
                  "poblacio": "Santa Lucía"}
        self.assertEqual(sv.zona_de({"codi_postal": None}, centro),
                         ("Canarias", "CP 35110 de la dirección, centro de consumo en Gemweb"))
        # sin codigo postal: la provincia o la isla por su nombre
        self.assertEqual(sv.zona_de({}, {"direccio": "C/ MAYOR 3 MANACOR (Illes Balears)",
                                         "poblacio": "Manacor"})[0], "Illes Balears")
        self.assertEqual(sv.zona_de({}, {"direccio": "AV. PRINCIPAL 1", "poblacio": "Mallorca"}),
                         ("Illes Balears", "«Mallorca», centro de consumo en Gemweb"))
        self.assertEqual(sv.zona_de({"poblacio": "Santa Cruz de La Palma"})[0], "Canarias")
        self.assertEqual(sv.zona_de({"poblacio": "Melilla"})[0], "Melilla")
        # una poblacion que no es de las islas, Ceuta ni Melilla: Peninsula
        self.assertEqual(sv.zona_de({}, {"direccio": "CALLE FALSA 1 MADRID  MADRID",
                                         "poblacio": "Madrid", "codi_postal": "-"}),
                         ("Península", "población «Madrid» (sin código postal), centro de "
                                       "consumo en Gemweb"))
        # un numero de portal de 5 cifras fuera de rango no es un codigo postal
        self.assertIsNone(sv._cp_en_texto("POLIGONO 12, NAVE 99999"))

    def test_utilidades(self):
        self.assertEqual(sv.periodo_defecto(date(2026, 9, 25)), (date(2025, 9, 1),
                                                                 date(2026, 8, 31)))
        self.assertEqual(sv.lista_cups(f"{CUPS_A.lower()};\n {CUPS_A}, {CUPS_B}"),
                         [CUPS_A, CUPS_B])
        s = sv.Suministro()
        sv.rellenar_desde_gemweb(s, {"tarifa_acces": "3.0TD", "nom": "Nave", "direccio": "C/ A 1",
                                     "codi_postal": "41001", "poblacio": "Sevilla",
                                     **{f"pot_contract_p{k}": str(10 * k) for k in range(1, 7)}})
        self.assertEqual((s.tarifa, s.denominacion, s.direccion, s.actual),
                         ("3.0TD", "Nave", "C/ A 1, 41001 Sevilla", [10, 20, 30, 40, 50, 60]))


class DatosDelEstudio(unittest.TestCase):
    def test_titular_comun(self):
        a = sv.Suministro(cups=CUPS_A, titular="EMPRESA A SA", grupo="GRUPO DE PRUEBA")
        b = sv.Suministro(cups=CUPS_B, titular="EMPRESA B SA", grupo="GRUPO DE PRUEBA")
        fuera = sv.Suministro(cups="ES0000000000000003CC0F")        # no esta en Gemweb
        self.assertEqual(sv.titular_comun([a]), "EMPRESA A SA")
        self.assertEqual(sv.titular_comun([a, b, fuera]), "GRUPO DE PRUEBA")
        b.grupo = "OTRO GRUPO"
        self.assertEqual(sv.titular_comun([a, b]), "")
        b.titular = "EMPRESA A SA"
        self.assertEqual(sv.titular_comun([a, b]), "EMPRESA A SA")

    def test_cups_del_fichero(self):
        self.assertEqual(sv.cups_del_fichero(CUPS_A[:20], [CUPS_B, CUPS_A]), CUPS_A)
        self.assertEqual(sv.cups_del_fichero(CUPS_A, [CUPS_A[:20]]), CUPS_A[:20])
        self.assertIsNone(sv.cups_del_fichero(CUPS_A, [CUPS_B]))

    def test_estudio_guardado_con_la_version_anterior(self):
        import pickle
        s = sv.Suministro(cups=CUPS_A)
        antiguo = {k: v for k, v in s.__dict__.items()
                   if k not in ("titular", "grupo", "en_gemweb", "fuente")}
        copia = sv.Suministro.__new__(sv.Suministro)
        copia.__setstate__(antiguo)
        self.assertEqual((copia.cups, copia.titular, copia.en_gemweb, copia.fuente),
                         (CUPS_A, "", False, ""))
        self.assertEqual(pickle.loads(pickle.dumps(s)).cups, CUPS_A)


class Curvas(unittest.TestCase):
    def test_fichero_subido_con_su_nombre(self):
        with tempfile.TemporaryDirectory() as tmp:
            g.escribir_p5d_multipunto({CUPS_A: g.curva_fisica(semilla=1),
                                       CUPS_B: g.curva_fisica(semilla=2)},
                                      Path(tmp) / "multi.txt")
            rutas = sv.guardar_subidos([("multi.txt", (Path(tmp) / "multi.txt").read_bytes())],
                                       Path(tmp) / "subidos")
            self.assertEqual(rutas[0].name, "multi.txt")
            lector = sv.abrir(rutas)
            sv.interpretar(lector, "Península")
            self.assertEqual(sorted(lector.cups_disponibles), [CUPS_A, CUPS_B])
            # multipunto: cada CUPS con su zona y sin el aviso «se usa X»
            s = sv.Suministro(cups=CUPS_B, zona="Canarias", avisos_extra=["aviso de Gemweb"])
            curva = sv.leer_suministro(s, lector, multipunto=True)
            self.assertEqual((curva.cups, s.zona_curva), (CUPS_B, "Canarias"))
            self.assertFalse(any("se usa" in a for a in curva.avisos))
            self.assertIn("aviso de Gemweb", curva.avisos)       # se conserva al releer
            self.assertTrue(sv.resumen_interpretacion(curva, lector, True)[0].startswith(
                "ESTUDIO MULTIPUNTO: 2 CUPS"))
            self.assertEqual(len(sv.primeros_registros(curva)), 12)


class GemwebSimulado:
    """El ClienteGemweb de la plataforma: buscar_suministro y descargar."""

    def __init__(self, valores):
        self.valores = valores

    def buscar_suministro(self, cups):
        if cups == CUPS_B:
            return None
        return {"id": 1, "tarifa_acces": "6.1TD", "tipus_lectura": "hora", "codi_postal": "38001"}

    def descargar(self, id_suministro, desde, hasta, al_avanzar=None):
        if al_avanzar:
            al_avanzar(1, 1, desde, hasta)
        return self.valores, []


class Gemweb(unittest.TestCase):
    def test_descarga_y_lectura(self):
        fis = g.curva_fisica()
        fin = g._etiquetas_fin(fis).dt.tz_localize(None)
        valores = list(zip(fin.dt.strftime("%Y-%m-%d %H:%M"), (fis["kw"] / 4).round(4)))
        avance = []
        with tempfile.TemporaryDirectory() as tmp:
            ruta, info, avisos = sv.descargar_gemweb(
                GemwebSimulado(valores), [CUPS_A, CUPS_B], date(2025, 6, 1), date(2026, 5, 31),
                tmp, al_avanzar=lambda f, t: avance.append(t))
            self.assertEqual(list(info), [CUPS_A])
            self.assertIn(f"{CUPS_B}: no está en el inventario de Gemweb.", avisos)
            self.assertTrue(info[CUPS_A]["_lectura_horaria"])
            self.assertTrue(avance)
            s = sv.Suministro(cups=CUPS_A, avisos_extra=sv.avisos_gemweb(CUPS_A, avisos,
                                                                          info[CUPS_A]))
            s.zona, s.zona_origen = sv.zona_de(info[CUPS_A])
            self.assertEqual(s.zona, "Canarias")
            from potencia.lector_curva import LectorCurva
            curva = sv.leer_suministro(s, LectorCurva(ruta),
                                       {"convenio": "Fin de periodo"}, multipunto=True)
            self.assertIn(sv.AVISO_HORARIA, curva.avisos)
            self.assertEqual(curva.unidad, "kWh")


class Calculo(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.s = suministro_calculado(cls._tmp.name)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_propuestas_y_tablas(self):
        s = self.s
        self.assertEqual([e.nombre for e in s.escenarios], ["Actual", "Propuesta 1",
                                                            "Propuesta 2"])
        # la Propuesta 2 vacia toma el optimo; una no valida da error
        self.assertEqual(list(s.escenarios[2].pc), list(s.optimo))
        s.propuestas.append({"pc": [300, 200, 200, 200, 200, 200], "opciones": None})
        self.assertIn("Propuesta 3", sv.calcular(s, sv.precios()))
        s.propuestas[-1]["pc"] = [300] * 6
        s.propuestas[-1]["opciones"] = {"acometida": False, "acceso": False, "extension": False,
                                        "enganche": False}
        self.assertIsNone(sv.calcular(s, sv.precios()))
        self.assertEqual(s.escenarios[3].inversion["total"], 0)
        tabla = sv.tabla_resultados(s.escenarios)
        self.assertEqual(list(tabla["Tipo"])[:2], ["Actual", "Propuesta 1 (óptima)"])
        self.assertEqual(len(sv.tabla_mensual(s.estudio.etiquetas_meses, s.escenarios)), 13)
        self.assertEqual(list(sv.tabla_energia(s.estudio)["Mes"])[-2:], ["Total", "%"])
        self.assertEqual(len(sv.tabla_detalle(s.estudio, s.escenarios[1]).columns), 17)
        self.assertEqual(sv.avisos_calculo(s), [])
        s.propuestas.pop()
        sv.calcular(s, sv.precios())

    def test_pdf_y_guardar(self):
        s = self.s
        pdf = sv.pdf_individual({"titular": "EMPRESA DE PRUEBA SA", "cups": s.cups,
                                 "instalacion": "Nave 1", "fecha": "23/10/2026"}, s, sv.precios(),
                                anexos=False)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertEqual(sv.nombre_pdf({"titular": "A/B", "instalacion": "N:1"}),
                         "Estudio potencia AB (N1).pdf")
        copia = sv.desempaquetar(sv.empaquetar([s]))[0]
        self.assertIsNone(copia.escenarios)
        self.assertEqual(len(copia.curva.datos), len(s.curva.datos))
        self.assertIsNone(sv.calcular(copia, sv.precios()))
        self.assertAlmostEqual(copia.escenarios[1].coste.total, s.escenarios[1].coste.total)


class Multipunto(unittest.TestCase):
    def test_resumen_y_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = suministro_calculado(Path(tmp) / "a")
            b = suministro_calculado(Path(tmp) / "b", zona="Canarias")
            a.cups, b.cups = CUPS_A, CUPS_B
            res = sv.resumen([a, b])
            filas = sv.filas_cups([a, b])
            self.assertEqual([f["zona"] for f in filas], ["Península", "Canarias"])
            tabla = sv.tabla_cups(filas, res)
            self.assertEqual(list(tabla["CUPS"]), [CUPS_A, CUPS_B, "TOTAL"])
            pdf = sv.pdf_multipunto({"titular": "GRUPO DE PRUEBA", "fecha": "23/10/2026"},
                                    [a, b], res, sv.precios(), anexos=False)
            self.assertTrue(pdf.startswith(b"%PDF"))
            r = sv.resumen_guardar("multipunto", {"titular": "GRUPO DE PRUEBA"}, [a, b])
            self.assertEqual((r["cups"], r["tarifa"]), ([CUPS_A, CUPS_B], "6.1TD"))
            self.assertAlmostEqual(r["ahorro"], round(res.escenarios[1].ahorro, 2))


class Guardados(unittest.TestCase):
    def test_guardar_listar_y_leer(self):
        from unittest import mock
        import almacen
        import config
        with tempfile.TemporaryDirectory() as tmp, \
                mock.patch.object(config, "CARPETA_DATOS", Path(tmp)), \
                mock.patch.object(config, "FICHERO_CONTRATOS", Path(tmp) / "no_existe.json"):
            almacen._iniciado.clear()
            try:
                eid = almacen.guardar_estudio_potencia(
                    "ana", {"tipo": "individual", "titular": "EMPRESA DE PRUEBA SA",
                            "cups": [CUPS_A], "tarifa": "6.1TD", "ahorro": 1000.5,
                            "inversion": 79.49},
                    {"tipo": "individual", "datos": {"titular": "EMPRESA DE PRUEBA SA"}},
                    b"curvas", b"%PDF")
                fila = almacen.listar_estudios_potencia()[0]
                self.assertEqual((fila["id"], fila["cups"], fila["suministros"], fila["ahorro"]),
                                 (eid, CUPS_A, 1, 1000.5))
                self.assertNotIn("pdf", fila)
                e = almacen.estudio_potencia(eid)
                self.assertEqual((e["datos"]["datos"]["titular"], e["curvas"], e["pdf"]),
                                 ("EMPRESA DE PRUEBA SA", b"curvas", b"%PDF"))
                self.assertIsNone(almacen.estudio_potencia(999))
            finally:
                almacen._iniciado.clear()


if __name__ == "__main__":
    unittest.main()
