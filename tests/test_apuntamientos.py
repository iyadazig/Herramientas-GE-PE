# -*- coding: utf-8 -*-
"""Apuntamientos del mercado diario: periodos 6P y 3P por sistema, Profile exacto con precios
sintéticos, Excel mensual de OMIE (días de 23 y 25 h), respaldo, selección de los 12 últimos
meses, cobertura de la curva y apuntamiento del cliente."""

import calendar
import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
from analisis_ofertas import apuntamientos as ap


def horas_del_mes(anio, mes):
    """[(fecha, hora ESIOS)] del mes, con los días de cambio de hora."""
    marzo, octubre = ap.calendario.dias_cambio_hora(anio)
    claves = []
    for d in range(1, calendar.monthrange(anio, mes)[1] + 1):
        f = dt.date(anio, mes, d)
        n = 23 if f == marzo else 25 if f == octubre else 24
        claves += [(f, h) for h in range(1, n + 1)]
    return claves


def precios_por_periodo(anio, mes, sistema, cal, precio_de_periodo):
    claves = horas_del_mes(anio, mes)
    per = ap.periodos(claves, sistema, cal)
    return {k: float(precio_de_periodo[p]) for k, p in zip(claves, per)}


class Horas(unittest.TestCase):
    def test_hora_de_inicio_en_los_cambios_de_hora(self):
        marzo, octubre = dt.date(2026, 3, 29), dt.date(2026, 10, 25)
        self.assertEqual([ap.hora_inicio(marzo, h) for h in (1, 2, 3, 23)], [0, 1, 3, 23])
        self.assertEqual([ap.hora_inicio(octubre, h) for h in (1, 3, 4, 5, 25)], [0, 2, 2, 3, 23])
        self.assertEqual(ap.hora_inicio(dt.date(2026, 9, 1), 24), 23)

    def test_periodos_2_0td(self):
        martes, sabado = dt.date(2026, 9, 1), dt.date(2026, 9, 5)
        pen = ap.periodos([(martes, h) for h in range(1, 25)], "Península", "3P")
        self.assertEqual(list(pen), [3] * 8 + [2, 2] + [1] * 4 + [2] * 4 + [1] * 4 + [2, 2])
        ceuta = ap.periodos([(martes, h) for h in range(1, 25)], "Ceuta", "3P")
        self.assertEqual(list(ceuta), [3] * 8 + [2, 2, 2] + [1] * 4 + [2] * 4 + [1] * 4 + [2])
        self.assertEqual(set(ap.periodos([(sabado, h) for h in range(1, 25)], "Península", "3P")),
                         {3})
        festivo = dt.date(2026, 10, 12)                    # lunes festivo nacional
        self.assertEqual(set(ap.periodos([(festivo, h) for h in range(1, 25)], "Melilla", "3P")),
                         {3})

    def test_periodos_6p_por_sistema(self):
        martes_julio = dt.date(2026, 7, 7)
        claves = [(martes_julio, 12)]                       # 11-12 h: bloque punta
        self.assertEqual(ap.periodos(claves, "Península", "6P")[0], 1)       # julio: alta
        self.assertEqual(ap.periodos(claves, "Canarias", "6P")[0], 1)        # julio: alta
        self.assertEqual(ap.periodos(claves, "Baleares", "6P")[0], 1)
        agosto = [(dt.date(2026, 8, 4), 16)]                # 15-16 h: llano
        self.assertEqual(ap.periodos(agosto, "Península", "6P")[0], 4)       # media: P3/P4
        self.assertEqual(ap.periodos(agosto, "Canarias", "6P")[0], 3)        # alta: P1/P3
        self.assertEqual(ap.periodos(agosto, "Ceuta", "6P")[0], 4)           # alta: P1/P4


class Profile(unittest.TestCase):
    def test_profile_exacto(self):
        # septiembre en la Península (temporada media: P3 y P4 en laborables, P6 el resto)
        precios = precios_por_periodo(2026, 9, "Península", "6P", {3: 300, 4: 200, 6: 100})
        n = {p: sum(1 for v in precios.values() if v == x) for p, x in ((3, 300), (4, 200),
                                                                        (6, 100))}
        media = (300 * n[3] + 200 * n[4] + 100 * n[6]) / len(precios)
        f = ap.profile(precios, "Península", "6P")
        self.assertEqual(f[:2] + [f[4]], [None, None, None])
        self.assertAlmostEqual(f[2], 300 / media)
        self.assertAlmostEqual(f[3], 200 / media)
        self.assertAlmostEqual(f[5], 100 / media)
        tres = ap.profile(precios, "Península", "3P")
        self.assertEqual(tres[3:], [None, None, None])

    def test_ponderado_con_consumo_uniforme_es_el_mismo(self):
        precios = precios_por_periodo(2026, 2, "Canarias", "6P", {1: 5, 2: 7, 3: 3, 4: 1,
                                                                  5: 1, 6: 2})
        pesos = {k: 3.0 for k in precios}
        for a, b in zip(ap.profile(precios, "Canarias", "6P"),
                        ap.profile(precios, "Canarias", "6P", pesos)):
            self.assertAlmostEqual(a or 0, b or 0)


def excel_omie(ruta, anio, mes, precio=lambda f, h: 50.0, hoja="3.0TD-6.xTD"):
    """Como los Excel mensuales de OMIE: título, «Día» y horas 1..24 (la 25 sin título),
    filas «L 01»..., «Total» y debajo la tabla «Profile» (que no se lee)."""
    wb = Workbook()
    ws = wb.active
    ws.title = hoja
    ws["A1"] = "OMIE - Mercado de electricidad"
    ws.append([])
    ws.append([])
    marzo, octubre = ap.calendario.dias_cambio_hora(anio)
    ws.append(["Día"] + [str(h) for h in range(1, 25)] + ["", "Media"])
    letras = "LMMJVSD"
    for d in range(1, calendar.monthrange(anio, mes)[1] + 1):
        f = dt.date(anio, mes, d)
        n = 23 if f == marzo else 25 if f == octubre else 24
        ws.append([f"{letras[f.weekday()]} {d:02d}"] + [precio(f, h) for h in range(1, n + 1)]
                  + [""] * (25 - n) + [1.0])
    ws.append(["Total "] + [1.0] * 25)
    ws.append(["Profile", "", "", 0.9])
    wb.save(ruta)
    return ruta


class Ficheros(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        (self.dir / "2025").mkdir()
        (self.dir / "2026").mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def test_nombres_de_fichero(self):
        self.assertEqual(ap.mes_anio_fichero("9 septiembre'26.XLS"), (9, 2026))
        self.assertEqual(ap.mes_anio_fichero("1 Enero'26.xls"), (1, 2026))
        self.assertEqual(ap.mes_anio_fichero("12 Diciembre'25.XLS"), (12, 2025))
        self.assertEqual(ap.mes_anio_fichero("4 abril.xls", 2024), (4, 2024))

    def test_dias_de_23_y_25_horas_y_fichero_en_otra_carpeta(self):
        excel_omie(self.dir / "2026" / "10 octubre'25.xlsx", 2025, 10)   # mal colocado
        excel_omie(self.dir / "2026" / "3 Marzo'26.xlsx", 2026, 3)
        ruta = ap.fichero_omie(2025, 10, self.dir)
        self.assertEqual(ruta.name, "10 octubre'25.xlsx")
        octubre = ap.precios_omie_fichero(ruta, 2025, 10)
        self.assertTrue(ap.completo(octubre, 2025, 10))
        self.assertEqual(len(octubre), 31 * 24 + 1)
        self.assertIn((dt.date(2025, 10, 26), 25), octubre)
        marzo = ap.precios_omie_fichero(ap.fichero_omie(2026, 3, self.dir), 2026, 3)
        self.assertEqual(len(marzo), 31 * 24 - 1)
        self.assertNotIn((dt.date(2026, 3, 29), 24), marzo)
        self.assertIsNone(ap.fichero_omie(2026, 4, self.dir))

    def test_respaldo_de_omie_si_esios_no_esta_completo(self):
        excel_omie(self.dir / "2025" / "11 noviembre'25.xlsx", 2025, 11)
        with mock.patch.object(ap, "precios_esios", side_effect=ap.SinPrecios("sin libro")):
            precios, fuente = ap.precios_mes(2025, 11, self.dir)
            self.assertEqual((fuente, len(precios)), (ap.OMIE, 30 * 24))
            with self.assertRaisesRegex(ap.SinPrecios, "sin Excel mensual"):
                ap.precios_mes(2025, 12, self.dir)
        completo = {k: 10.0 for k in horas_del_mes(2025, 11)}
        with mock.patch.object(ap, "precios_esios", return_value=completo):
            self.assertEqual(ap.precios_mes(2025, 11, self.dir)[1], ap.ESIOS)
        a_medias = dict(list(completo.items())[:100])
        with mock.patch.object(ap, "precios_esios", return_value=a_medias):
            self.assertEqual(ap.precios_mes(2025, 11, self.dir)[1], ap.OMIE)


class Seleccion(unittest.TestCase):
    def calculo(self, anio, mes, fuente=ap.ESIOS, valor=1.0):
        return {"anio": anio, "mes": mes, "fuente": fuente,
                "factores": {f"{s}|{c}": [valor] * 6 for s in ap.SISTEMAS
                             for c in ap.CALENDARIOS}}

    def test_actualizar_solo_lo_que_falta(self):
        hoy = dt.date(2026, 11, 5)
        self.assertEqual(ap.meses_a_calcular(hoy, 3), [(2026, 10), (2026, 9), (2026, 8)])
        guardados = {(2026, 10): {"fuente": ap.ESIOS}, (2026, 9): {"fuente": ap.OMIE}}
        hechos = []

        def calcular(a, m, carpeta=None):
            if (a, m) == (2026, 7):
                raise ap.SinPrecios("julio 2026: nada")
            return self.calculo(a, m, ap.ESIOS if (a, m) != (2026, 8) else ap.OMIE)
        with mock.patch.object(ap, "calcular_mes", side_effect=calcular):
            avisos = ap.actualizar(hoy, guardados, guardar=hechos.append, meses_atras=4)
        self.assertEqual([(c["anio"], c["mes"]) for c in hechos], [(2026, 9), (2026, 8)])
        self.assertEqual(avisos, ["julio 2026: nada"])

    def test_vigentes_los_12_ultimos(self):
        calculados = [self.calculo(2025, m, valor=0.5) for m in range(1, 13)] + \
                     [self.calculo(2026, m, valor=0.9) for m in range(1, 12)]
        v = ap.vigentes(calculados, hoy=dt.date(2026, 11, 5))
        pen = v["Península|6P"]
        self.assertEqual((pen[10]["anio"], pen[10]["factores"][0]), (2026, 0.9))
        self.assertEqual(pen[11]["anio"], 2025)              # nov-26 aún no ha terminado
        self.assertEqual(pen[12]["anio"], 2025)
        self.assertEqual(ap.etiqueta(2025, 12, ap.OMIE), "dic-25 (OMIE)")
        self.assertFalse(ap.pendientes(calculados, dt.date(2026, 11, 5)))    # octubre está
        self.assertTrue(ap.pendientes(calculados, dt.date(2027, 1, 5)))      # falta diciembre


def curva_año(desde=dt.date(2025, 10, 1), dias=365, kwh=lambda f, h: 1.0, saltar=()):
    curva = {}
    for n in range(dias):
        f = desde + dt.timedelta(days=n)
        if f in saltar or (f.year, f.month) in saltar:
            continue
        marzo, octubre = ap.calendario.dias_cambio_hora(f.year)
        for h in range(1, (23 if f == marzo else 25 if f == octubre else 24) + 1):
            curva[(f, h)] = kwh(f, h)
    return curva


class Cliente(unittest.TestCase):
    def test_cobertura(self):
        c = ap.cobertura(curva_año())
        self.assertTrue(c["valida"])
        self.assertEqual((c["desde"], c["hasta"]), (dt.date(2025, 10, 1), dt.date(2026, 9, 30)))
        sin_enero = ap.cobertura(curva_año(saltar={(2026, 1)}))
        self.assertFalse(sin_enero["valida"])
        self.assertIn("ene-26", sin_enero["motivo"])
        huecos = ap.cobertura(curva_año(saltar={dt.date(2026, 5, d) for d in range(1, 11)}))
        self.assertFalse(huecos["valida"])
        self.assertIn("355 días", huecos["motivo"])
        self.assertFalse(ap.cobertura(curva_año(dias=200))["valida"])

    def test_apuntamiento_del_cliente(self):
        def precios_de(a, m):
            return {k: (200.0 if 10 <= k[1] <= 14 else 100.0) for k in horas_del_mes(a, m)}, \
                ap.ESIOS
        # consume solo de 10 a 14 h (horas ESIOS 10..14): paga más que la media
        curva = curva_año(kwh=lambda f, h: 1.0 if 10 <= h <= 14 else 0.0)
        r = ap.apuntamiento_cliente(curva, "Península", "2.0TD", precios_de)
        self.assertTrue(r["valida"])
        self.assertEqual(sorted(r["factores"]), list(range(1, 13)))
        self.assertEqual(r["fuentes"][10], "oct-25 (ESIOS)")
        factores_julio = r["factores"][7]
        self.assertEqual(factores_julio[3:], [None, None, None])
        self.assertGreater(factores_julio[0], 1.0)
        # precios planos: todo 1
        plano = ap.apuntamiento_cliente(curva_año(), "Canarias", "6.1TD",
                                        lambda a, m: ({k: 80.0 for k in horas_del_mes(a, m)},
                                                      ap.OMIE))
        for f in plano["factores"].values():
            self.assertTrue(all(x is None or abs(x - 1) < 1e-12 for x in f))
        # sin precios de algún mes: no válida, con el motivo
        def sin_precios(a, m):
            raise ap.SinPrecios("marzo 2026: nada")
        r = ap.apuntamiento_cliente(curva_año(), "Península", "3.0TD", sin_precios)
        self.assertFalse(r["valida"])
        self.assertIn("Sin precios horarios", r["motivo"])


class Guardado(unittest.TestCase):
    def test_almacen(self):
        with tempfile.TemporaryDirectory() as tmp, \
                mock.patch.object(config, "CARPETA_DATOS", Path(tmp)):
            almacen._iniciado.clear()
            c = {"anio": 2026, "mes": 9, "fuente": ap.OMIE,
                 "factores": {"Península|6P": [None, None, 0.9, 0.8, None, 1.1]}}
            almacen.guardar_apuntamientos(c)
            almacen.guardar_apuntamientos(dict(c, fuente=ap.ESIOS))
            filas = almacen.listar_apuntamientos()
            self.assertEqual(len(filas), 1)
            self.assertEqual((filas[0]["fuente"], filas[0]["factores"]["Península|6P"][2]),
                             (ap.ESIOS, 0.9))
            almacen._iniciado.clear()


if __name__ == "__main__":
    unittest.main()
