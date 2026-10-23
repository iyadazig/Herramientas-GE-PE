# -*- coding: utf-8 -*-
"""Calculo del estudio de potencia (paquete potencia/): las pruebas del programa de escritorio
EstudioPotencia, pasadas a unittest, con curvas sinteticas (tests/generador_curvas.py)."""

import itertools
import math
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import numpy as np
import pandas as pd

import generador_curvas as g
from potencia.calculo import (Estudio, ResumenMultipunto, Tarifa, calcular_inversion,
                              cargar_precios, evaluar, opciones_inversion_defecto,
                              validar_potencias)
from potencia.calendario import asignar_periodos, festivos
from potencia.lector_curva import LectorCurva, LectorCurvas, abrir_curvas

CUPS_A = "ES0000000000000001AA0F"
CUPS_B = "ES0000000000000002BB0F"


def cerca(a, b, rel=1e-6):
    return math.isclose(a, b, rel_tol=rel, abs_tol=1e-9)


class ConCarpeta(unittest.TestCase):
    """Carpeta temporal por prueba (el tmp_path de pytest)."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()


# ------------------------------------------------------------------- calculo
class Calculo(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        esp = g.esperado(g.curva_fisica())
        curva = pd.DataFrame({"inicio": esp.index, "kw": esp.to_numpy()})
        cls.estudio = Estudio(curva, "Península", Tarifa.desde_precios(cargar_precios(), "6.1TD"))

    def test_periodos_peninsula(self):
        ts = pd.to_datetime(["2026-01-13 09:00", "2026-01-13 08:45", "2026-01-13 07:45",
                             "2026-03-10 10:00", "2026-06-10 22:15", "2026-04-14 19:00",
                             "2026-01-17 12:00", "2026-01-06 12:00", "2026-01-01 12:00",
                             "2026-05-01 12:00"])
        self.assertEqual(list(asignar_periodos(ts, "Península")), [1, 2, 6, 2, 4, 4, 6, 6, 6, 6])

    def test_periodos_canarias_y_ceuta(self):
        # Canarias: julio (alta) -> P1 punta, P3 llano; Ceuta: enero (alta) -> P1 punta, P4 llano
        ts = pd.to_datetime(["2026-07-14 11:00", "2026-07-14 09:00"])
        self.assertEqual(list(asignar_periodos(ts, "Canarias")), [1, 3])
        ts = pd.to_datetime(["2026-01-13 20:00", "2026-01-13 16:00", "2026-04-14 11:00"])
        self.assertEqual(list(asignar_periodos(ts, "Ceuta")), [1, 4, 3])

    def test_festivos_fijos(self):
        f = festivos([2026])
        self.assertIn(pd.Timestamp("2026-10-12").date(), f)
        self.assertNotIn(pd.Timestamp("2026-04-03").date(), f)   # Viernes Santo: sin fecha fija

    def test_coste_formula_manual(self):
        estudio = self.estudio
        pc = np.array([200, 210, 220, 230, 240, 250.0])
        c = estudio.coste(pc)
        t = estudio.tarifa
        fijo = (pc * t.tp).sum() * estudio.dias_totales / t.dias_anio
        self.assertTrue(cerca(c.total_fijo, fijo))
        exc = 0.0
        for i in range(len(estudio.meses)):
            for p in range(6):
                v = estudio.kw[(estudio.idx_mes == i) & (estudio.periodo == p + 1)]
                d = np.clip(v - pc[p], 0, None)
                exc += t.tep[p] * np.sqrt((d ** 2).sum())
        self.assertTrue(cerca(c.total_exceso, exc))

    def test_optimo_es_minimo_con_orden(self):
        o = self.estudio.optimo()
        self.assertTrue(all(o[i] <= o[i + 1] for i in range(5)))
        base = self.estudio.coste(o).total
        # ningun vecino entero (que respete el orden) mejora el optimo
        for delta in itertools.product((-1, 0, 1), repeat=6):
            pc = o + np.array(delta)
            if validar_potencias(pc) is None:
                self.assertGreaterEqual(self.estudio.coste(pc).total, base - 1e-6)

    def test_inversion(self):
        derechos = {"acceso": 16.992541, "extension": 15.718632, "enganche": 79.49197}
        actual = [250, 265, 265, 265, 265, 265]
        op = opciones_inversion_defecto([340] * 6, actual)
        inv = calcular_inversion([340] * 6, actual, op, derechos)
        self.assertTrue(cerca(inv["total"], 75 * (16.992541 + 15.718632) + 79.49197))
        op = opciones_inversion_defecto([265] * 6, actual)
        self.assertTrue(cerca(calcular_inversion([265] * 6, actual, op, derechos)["total"],
                              79.49197))

    def test_validar_potencias(self):
        self.assertIsNone(validar_potencias([1, 2, 3, 4, 5, 6]))
        self.assertIn("crecientes", validar_potencias([10, 5, 5, 5, 5, 5]))


# ------------------------------------------------------------------- lector
class Lector(ConCarpeta):
    @classmethod
    def setUpClass(cls):
        cls.fisica = g.curva_fisica()
        cls.esperado_qh = g.esperado(cls.fisica)
        cls.esperado_h = g.esperado(cls.fisica, horaria=True)

    def comparar(self, curva, esperado, tol):
        s = curva.datos.set_index("inicio")["kw"]
        self.assertEqual(len(s), len(esperado))
        self.assertTrue((s.index == esperado.index).all())
        self.assertLess(np.abs(s.to_numpy() - esperado.to_numpy()).max(), tol)

    def test_cuartohorarias(self):
        for escritor, nombre, unidad, conv in (
                (g.escribir_p5d, "p5d.txt", "kWh", "Fin de periodo"),
                (g.escribir_datadis_cuartos, "datadis.csv", "kWh", "Fin de periodo"),
                (g.escribir_iso_offset, "iso.csv", "kW", "Fin de periodo")):
            with self.subTest(nombre):
                ruta = self.tmp / nombre
                escritor(self.fisica, ruta)
                c = LectorCurva(ruta).procesar()
                self.assertEqual((c.resolucion_min, c.unidad, c.huecos_rellenados,
                                  len(c.dias_cambio_hora)), (15, unidad, 0, 2))
                self.assertTrue(c.convenio.startswith(conv))
                self.comparar(c, self.esperado_qh, tol=0.01)

    def test_excel_geype(self):
        ruta = self.tmp / "geype.xlsx"
        g.escribir_excel_geype(self.esperado_qh, ruta)
        lector = LectorCurva(ruta)
        self.assertEqual(lector.hoja_defecto, "Curva_potencia")
        c = lector.procesar()
        self.assertEqual(c.unidad, "kW")
        self.comparar(c, self.esperado_qh, tol=1e-6)

    def test_horaria_ordinal_wh(self):
        ruta = self.tmp / "horaria.csv"
        g.escribir_distribuidora_horaria(self.fisica, ruta)
        c = LectorCurva(ruta).procesar()
        self.assertEqual((c.resolucion_min, c.unidad, c.es_horaria), (60, "Wh", True))
        self.comparar(c, self.esperado_h, tol=0.002)

    def test_ancho_horario(self):
        ruta = self.tmp / "ancho.xlsx"
        g.escribir_ancho_horario(self.fisica, ruta)
        c = LectorCurva(ruta).procesar()
        self.assertEqual(c.resolucion_min, 60)
        self.comparar(c, self.esperado_h, tol=1e-6)

    def test_huecos_y_recorte(self):
        fis = g.curva_fisica("2025-03-01", "2026-06-15")
        fis = fis.drop(fis.index[1000:1008])               # dos horas sin dato
        ruta = self.tmp / "p5d.txt"
        g.escribir_p5d(fis, ruta)
        c = LectorCurva(ruta).procesar()
        self.assertEqual(c.huecos_rellenados, 8)
        self.assertEqual((c.fecha_inicio, c.fecha_fin),
                         (pd.Timestamp("2025-06-01"), pd.Timestamp("2026-06-01")))
        self.assertTrue(any("12 últimos meses" in a for a in c.avisos))

    def test_forzar_unidad(self):
        ruta = self.tmp / "p5d.txt"
        g.escribir_p5d(self.fisica, ruta)
        c = LectorCurva(ruta).procesar(unidad="kW")
        self.assertTrue(np.allclose(c.datos["kw"].to_numpy() * 4, self.esperado_qh.to_numpy(),
                                    atol=0.01))


# ------------------------------------------------------------------- unidades
class Unidades(ConCarpeta):
    @classmethod
    def setUpClass(cls):
        cls.fisica = g.curva_fisica()
        cls.esperado = g.esperado(cls.fisica)

    def escribir(self, ruta, cabecera, factor, antes=(), columna_unidad=None):
        """CSV con etiqueta de fin de periodo; `factor` convierte kW al valor escrito."""
        loc = g._etiquetas_fin(self.fisica)
        filas = list(antes) + cabecera
        for l, k in zip(loc, self.fisica["kw"]):
            fila = f"{l:%d/%m/%Y};{l:%H:%M};{str(round(k * factor, 6)).replace('.', ',')}"
            filas.append(fila + (f";{columna_unidad}" if columna_unidad else ""))
        ruta.write_text("\n".join(filas), encoding="utf-8")

    def test_unidad_detectada(self):
        for caso, cabecera, factor, antes, col_unidad, unidad, origen in (
                ("titulo_kw", ["Fecha;Hora;Potencia media (kW)"], 1, (), None, "kW", "título"),
                ("titulo_guion", ["Fecha;Hora;AE_kW"], 1, (), None, "kW", "título"),
                ("fila_unidades", ["Fecha;Hora;Valor", ";;kW"], 1, (), None, "kW", "título"),
                ("fila_formatos", ["Fecha;Hora;Valor", "dd/mm/aaaa;hh:mm;kWh"], 0.25, (), None,
                 "kWh", "«valor kwh»"),
                ("columna_unidad", ["Fecha;Hora;Medida;Unidad"], 1, (), "kW", "kW",
                 "columna de unidades"),
                ("nota_encima", ["Fecha;Hora;Medida"], 1,
                 ("Curva de carga del suministro", "Unidades: kW"), None, "kW",
                 "encima de la tabla"),
                ("mwh", ["Fecha;Hora;Consumo (MWh)"], 0.25 / 1000, (), None, "MWh", "título"),
                ("simbolo_manda", ["Fecha;Hora;Potencia (kWh)"], 0.25, (), None, "kWh", "título"),
                ("palabra", ["Fecha;Hora;Potencia"], 1, (), None, "kW", "palabra")):
            with self.subTest(caso):
                ruta = self.tmp / f"{caso}.csv"
                self.escribir(ruta, cabecera, factor, antes, col_unidad)
                c = LectorCurva(ruta).procesar()
                self.assertEqual(c.unidad, unidad)
                self.assertIn(origen, c.origen_unidad)
                self.assertFalse(any("No se ha podido identificar la unidad" in a
                                     for a in c.avisos))
                self.assertLess(np.abs(c.datos["kw"].to_numpy()
                                       - self.esperado.to_numpy()).max(), 0.01)

    def test_sin_unidad_avisa(self):
        ruta = self.tmp / "sin_unidad.csv"
        self.escribir(ruta, ["Fecha;Hora;Medida"], 0.25)
        c = LectorCurva(ruta).procesar()
        self.assertEqual(c.unidad, "kWh")
        self.assertTrue(any("No se ha podido identificar la unidad" in a for a in c.avisos))


# ------------------------------------------------------------------- multipunto
class Multipunto(ConCarpeta):
    @classmethod
    def setUpClass(cls):
        cls.curvas = {CUPS_A: g.curva_fisica(semilla=1), CUPS_B: g.curva_fisica(semilla=2)}

    def test_lector_varios_cups(self):
        ruta = self.tmp / "multi.txt"
        g.escribir_p5d_multipunto(self.curvas, ruta)
        lector = LectorCurva(ruta)
        lector.procesar()
        self.assertEqual(sorted(lector.cups_disponibles), sorted(self.curvas))
        for cups, fis in self.curvas.items():
            c = lector.procesar(cups=cups)
            self.assertEqual(c.cups, cups)
            self.assertLess(np.abs(c.datos["kw"].to_numpy()
                                   - g.esperado(fis).to_numpy()).max(), 0.01)

    def test_resumen_suma_actual_y_optima(self):
        pr = cargar_precios()
        lista = []
        for k, (cups, fis) in enumerate(self.curvas.items()):
            esp = g.esperado(fis)
            tarifa = Tarifa.desde_precios(pr, "6.1TD" if k == 0 else "3.0TD")
            est = Estudio(pd.DataFrame({"inicio": esp.index, "kw": esp.to_numpy()}),
                          "Península", tarifa)
            actual = [150] * 5 + [200]
            opt = est.optimo()
            escs = evaluar(est, [("Propuesta 1", opt, opciones_inversion_defecto(opt, actual))],
                           actual)
            lista.append((est, escs))
        r = ResumenMultipunto(lista)
        act, optima = r.escenarios
        self.assertTrue(cerca(act.coste.total, sum(e[0].coste.total for _, e in lista)))
        self.assertTrue(cerca(optima.coste.total, sum(e[1].coste.total for _, e in lista)))
        self.assertTrue(cerca(optima.inversion["total"],
                              sum(e[1].inversion["total"] for _, e in lista)))
        self.assertTrue(cerca(optima.ahorro, act.coste.total - optima.coste.total))
        self.assertTrue(len(r.etiquetas_meses) == 12 and not r.periodos_distintos)
        self.assertTrue(cerca(r.energia_kwh().sum(),
                              sum(est.energia_kwh().sum() for est, _ in lista)))


# ------------------------------------------------------------------- varios ficheros
def _por_meses(fis):
    """Trozos de la curva fisica por mes de la hora local de inicio de cada cuarto."""
    mes = fis["utc"].dt.tz_convert(g.TZ).dt.strftime("%Y-%m")
    return [(m, fis[mes == m].reset_index(drop=True)) for m in sorted(mes.unique())]


def _p5d(fis, ruta, cups):
    g.escribir_p5d_multipunto({cups: fis}, ruta)


def _excel_kw(fis, ruta):
    """Excel con fecha-hora de fin del cuarto y potencia en kW, sin columna de CUPS (va en el
    nombre)."""
    fin = g._etiquetas_fin(fis).dt.tz_localize(None)
    pd.DataFrame({"Fecha y hora": fin.dt.strftime("%d/%m/%Y %H:%M"),
                  "Potencia (kW)": fis["kw"]}).to_excel(ruta, index=False)


class VariosFicheros(ConCarpeta):
    def comparar(self, curva, esperado, tol=0.01):
        s = curva.datos.set_index("inicio")["kw"]
        esperado = esperado.reindex(s.index)
        self.assertTrue(len(s))
        self.assertLess(np.abs(s.to_numpy() - esperado.to_numpy()).max(), tol)

    def test_un_fichero_por_mes_formatos_mezclados(self):
        fis = g.curva_fisica()
        rutas = []
        for k, (mes, trozo) in enumerate(_por_meses(fis)):
            if k % 3 == 0:
                ruta = self.tmp / f"curva_{mes}.txt"
                _p5d(trozo, ruta, CUPS_A)                         # kWh, hora de fin, con CUPS
            elif k % 3 == 1:
                ruta = self.tmp / f"{CUPS_A}_{mes}.xlsx"
                _excel_kw(trozo, ruta)                            # kW, CUPS solo en el nombre
            else:
                ruta = self.tmp / f"datadis_{mes}.csv"
                g.escribir_datadis_cuartos(trozo, ruta)           # kWh, hora 24:00
                ruta.write_text(ruta.read_text(encoding="cp1252").replace(g.CUPS, CUPS_A),
                                encoding="cp1252")
            rutas.append(ruta)
        lector = abrir_curvas(rutas)
        self.assertIsInstance(lector, LectorCurvas)
        c = lector.procesar()
        self.assertEqual((c.cups, lector.cups_disponibles), (CUPS_A, [CUPS_A]))
        self.assertEqual((c.fecha_inicio, c.fecha_fin),
                         (pd.Timestamp("2025-06-01"), pd.Timestamp("2026-06-01")))
        self.assertEqual((len(c.ficheros), c.sin_datos, c.huecos_rellenados), (12, [], 0))
        self.assertIn("varias", c.unidad)                         # kWh y kW segun el fichero
        self.comparar(c, g.esperado(fis))

    def test_varios_cups_en_ficheros_mensuales(self):
        fisicas = {CUPS_A: g.curva_fisica(semilla=1), CUPS_B: g.curva_fisica(semilla=2)}
        rutas = []
        for cups, fis in fisicas.items():
            for mes, trozo in _por_meses(fis):
                ruta = self.tmp / f"{cups}_{mes}.xlsx"
                _excel_kw(trozo, ruta)
                rutas.append(ruta)
        lector = LectorCurvas(sorted(rutas))                      # mezclados por orden alfabetico
        self.assertEqual(sorted(lector.cups_disponibles), sorted(fisicas))
        for cups, fis in fisicas.items():
            c = lector.procesar(cups=cups)
            self.assertEqual((c.cups, len(c.datos)), (cups, len(g.esperado(fis))))
            self.comparar(c, g.esperado(fis))

    def test_falta_un_mes_no_se_inventa(self):
        fis = g.curva_fisica()
        rutas = []
        for mes, trozo in _por_meses(fis):
            if mes == "2025-11":
                continue                                          # falta el fichero de noviembre
            ruta = self.tmp / f"curva_{mes}.txt"
            _p5d(trozo, ruta, CUPS_A)
            rutas.append(ruta)
        c = LectorCurvas(rutas).procesar()
        self.assertEqual(c.sin_datos, [(pd.Timestamp("2025-11-01"), pd.Timestamp("2025-12-01"))])
        self.assertFalse((c.datos["inicio"].dt.strftime("%Y-%m") == "2025-11").any())
        self.assertTrue(any("No hay datos del 01/11/2025" in x for x in c.avisos))
        self.comparar(c, g.esperado(fis))

    def test_ficheros_solapados(self):
        fis = g.curva_fisica()
        meses = _por_meses(fis)
        rutas = []
        for k, (mes, trozo) in enumerate(meses):
            if k + 1 < len(meses):              # cada fichero lleva tambien el dia 1 del siguiente
                trozo = pd.concat([trozo, meses[k + 1][1].head(96)], ignore_index=True)
            ruta = self.tmp / f"curva_{mes}.txt"
            _p5d(trozo, ruta, CUPS_A)
            rutas.append(ruta)
        c = LectorCurvas(rutas).procesar()
        self.assertTrue(any("más de un fichero" in x for x in c.avisos))
        self.assertFalse(any("duplicados" in x for x in c.avisos))
        self.comparar(c, g.esperado(fis))

    def test_un_fichero_con_hueco_largo(self):
        fis = g.curva_fisica()
        loc = fis["utc"].dt.tz_convert(g.TZ)
        fis = fis[~((loc >= "2025-09-10") & (loc < "2025-09-15"))]          # 5 dias sin datos
        ruta = self.tmp / "curva.txt"
        _p5d(fis, ruta, CUPS_A)
        c = LectorCurva(ruta).procesar()
        self.assertEqual(c.sin_datos, [(pd.Timestamp("2025-09-10"), pd.Timestamp("2025-09-15"))])
        self.assertEqual(c.huecos_rellenados, 0)
        self.assertEqual(len(c.datos), len(g.esperado(g.curva_fisica())) - 5 * 96)


if __name__ == "__main__":
    unittest.main()
