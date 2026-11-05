# -*- coding: utf-8 -*-
"""Lectores de OMIP y OMIE con ficheros sintéticos: fechas mezcladas, prioridad M > Q > YR,
mes sin cotización, cabecera de OMIE por texto, fichero bloqueado, caché y avisos."""

import datetime as dt
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import datos_mercado as dm


def fichero_omip(ruta: Path, filas: list) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Espana"
    ws.append(["fecha_sesion", "fecha_descarga", "producto", "zona", "instrumento", "contrato",
               "maturity", "settlement_price_eur_mwh"])
    for f in filas:
        ws.append(f)
    wb.create_sheet("Portugal")
    wb.save(ruta)
    return ruta


def fila(sesion, contrato, precio, producto="Electricidad", instrumento="FTB"):
    return [sesion, "2026-11-05 08:00:00", producto, "ES", instrumento, contrato,
            contrato.split()[1][:2] if contrato.startswith("FTB") else "SPOT", precio]


def fichero_omie(ruta: Path, precios: dict) -> Path:
    """Como «PRECIO MEDIO OMIE.xls»: título, cabecera «mes/año» desplazada, años en columnas,
    medias y trimestres debajo."""
    wb = Workbook()
    ws = wb.active
    ws.title = "desde 2019"
    ws["B2"] = "EVOLUCIÓN HISTÓRICA DE LOS PRECIOS DE MERCADO ELÉCTRICO (OMIE)"
    anios = sorted({a for a, _ in precios})
    ws.cell(7, 3, "mes/año")
    for n, a in enumerate(anios):
        ws.cell(7, 4 + n, float(a))
    ws.cell(7, 5 + len(anios), None)
    ws.cell(6, 6 + len(anios), "Promedio 3 años")
    for m in range(1, 13):
        ws.cell(7 + m, 3, dm.MESES[m - 1])
        for n, a in enumerate(anios):
            if (a, m) in precios:
                ws.cell(7 + m, 4 + n, precios[(a, m)])
        ws.cell(7 + m, 6 + len(anios), 999)          # columnas de después: no se leen
    ws.cell(20, 3, "Media anual")
    ws.cell(23, 3, "Q1")
    ws.cell(23, 4, 12345)
    wb.create_sheet("desde 2015")
    wb.save(ruta)
    return ruta


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        dm._cache.clear()

    def tearDown(self):
        dm._cache.clear()
        self.tmp.cleanup()


class Omip(Base):
    def setUp(self):
        super().setUp()
        self.ruta = fichero_omip(self.dir / "omip_history.xlsx", [
            fila(dt.datetime(2026, 10, 30), "FTB M Jan-27", 100.0),
            fila("2026-11-03", "FTB M Jan-27", 110.0),
            fila("2026-11-04", "FTB M Jan-27", 120.0),         # fechas en texto y en fecha
            fila("2026-11-04", "FTB Q1-27", 90.0),
            fila("2026-11-04", "FTB Q2-27", 70.0),
            fila("2026-11-04", "FTB YR-27", 80.0),
            fila("2026-11-04", "FTB YR-28", 60.0),
            fila("2026-11-04", "FTB Wk46-26", 1.0),            # semanas, días, PPA: no
            fila("2026-11-04", "FTB PPA 27/36", 1.0),
            fila("2026-11-04", "FGE YR-28", 42.0, "Gas natural", "FGE"),
            fila("2026-11-04", "SPEL BASE 2026-11-05", 1.0, instrumento="SPELBASE")])

    def test_ultima_sesion_y_prioridad(self):
        c = dm.omip(self.ruta)
        self.assertEqual(dm.sesiones(c), [dt.date(2026, 10, 30), dt.date(2026, 11, 3),
                                          dt.date(2026, 11, 4)])
        s = dm.sesion_omip(c)
        self.assertEqual(s["fecha"], dt.date(2026, 11, 4))
        self.assertEqual(dm.omip_mes(s, 2027, 1), (120.0, "M ene-27"))      # mensual
        self.assertEqual(dm.omip_mes(s, 2027, 2), (90.0, "Q1-27"))          # trimestral
        self.assertEqual(dm.omip_mes(s, 2027, 8), (80.0, "YR-27"))          # anual
        self.assertEqual(dm.omip_mes(s, 2028, 1), (60.0, "YR-28"))
        self.assertEqual(dm.omip_mes(s, 2029, 1), (None, None))             # sin cotización
        self.assertEqual(dm.omip_mes(dm.sesion_omip(c, dt.date(2026, 11, 1)), 2027, 1),
                         (100.0, "M ene-27"))

    def test_precio_previsto(self):
        s = dm.sesion_omip(dm.omip(self.ruta))
        omie = {(2024, 1): 60.0, (2025, 1): 70.0, (2026, 1): 80.0, (2023, 1): 1000.0}
        f = dm.precio_previsto([(2027, 1), (2029, 1), (2027, 2)], s, omie, 0.95, 3)
        self.assertAlmostEqual(f[0]["previsto"], 0.95 * 120 + 0.05 * 70)
        self.assertEqual(f[0]["anios_omie"], [2024, 2025, 2026])
        self.assertEqual((f[1]["previsto"], f[1]["omip"]), (70.0, None))
        self.assertIn("Sin cotización de OMIP", f[1]["aviso"])
        self.assertEqual((f[2]["previsto"], f[2]["aviso"]), (90.0, "Sin media de OMIE: solo OMIP"))

    def test_cache_por_fecha_de_modificacion(self):
        dm.omip(self.ruta)
        with mock.patch.object(dm, "_leer_omip", side_effect=AssertionError("sin caché")):
            dm.omip(self.ruta)                         # no vuelve a leer
        fichero_omip(self.ruta, [fila("2026-11-05", "FTB YR-27", 81.0),
                                 fila("2026-11-05", "FTB YR-28", 61.0, )])
        import os
        os.utime(self.ruta, ns=(1, 2_000_000_000_000_000_000))   # fecha distinta seguro
        self.assertEqual(dm.sesion_omip(dm.omip(self.ruta))["YR"][2027], 81.0)

    def test_fichero_bloqueado_se_reintenta(self):
        original = shutil.copy2
        llamadas = []

        def bloqueado(*a, **k):
            llamadas.append(1)
            if len(llamadas) == 1:
                raise PermissionError("en uso")
            return original(*a, **k)
        with mock.patch.object(dm.shutil, "copy2", side_effect=bloqueado), \
                mock.patch.object(dm.time, "sleep"):
            self.assertEqual(len(dm.omip(self.ruta)), 7)
        self.assertEqual(len(llamadas), 2)
        with mock.patch.object(dm.shutil, "copy2", side_effect=PermissionError("en uso")), \
                mock.patch.object(dm.time, "sleep"):
            dm._cache.clear()
            with self.assertRaisesRegex(dm.DatosMercado, "bloqueado"):
                dm.omip(self.ruta)

    def test_errores_claros(self):
        with self.assertRaisesRegex(dm.DatosMercado, "No se encuentra"):
            dm.omip(self.dir / "no_existe.xlsx")
        wb = Workbook()
        wb.active.title = "Otra"
        wb.save(self.dir / "mal.xlsx")
        with self.assertRaisesRegex(dm.DatosMercado, "Espana"):
            dm.omip(self.dir / "mal.xlsx")


class Omie(Base):
    def test_lee_por_la_cabecera_y_media_de_los_ultimos_años(self):
        precios = {(a, m): a - 2000 + m / 100 for a in range(2019, 2027) for m in range(1, 13)
                   if (a, m) < (2026, 11)}
        p = dm.omie_medio(fichero_omie(self.dir / "omie.xlsx", precios))
        self.assertEqual(p, precios)
        media, anios = dm.media_omie(p, 1, 3)
        self.assertEqual(anios, [2024, 2025, 2026])
        self.assertAlmostEqual(media, (24.01 + 25.01 + 26.01) / 3)
        self.assertEqual(dm.media_omie(p, 11, 3)[1], [2023, 2024, 2025])   # 2026 aún no
        self.assertEqual(dm.media_omie(p, 11, 5)[1], [2021, 2022, 2023, 2024, 2025])

    def test_sin_cabecera(self):
        wb = Workbook()
        wb.active.title = "desde 2019"
        wb.active["A1"] = "otra cosa"
        wb.save(self.dir / "omie.xlsx")
        with self.assertRaisesRegex(dm.DatosMercado, "mes/año"):
            dm.omie_medio(self.dir / "omie.xlsx")


class Estado(Base):
    def test_avisos(self):
        omip = fichero_omip(self.dir / "omip.xlsx", [fila("2026-10-01", "FTB YR-27", 80.0)])
        omie = fichero_omie(self.dir / "omie.xlsx", {(2026, 7): 50.0})
        e = dm.estado(omip, omie, hoy=dt.date(2026, 11, 5))
        self.assertEqual(e["omip"]["ultima_sesion"], dt.date(2026, 10, 1))
        self.assertEqual(e["omie"]["ultimo_mes"], (2026, 7))
        self.assertEqual(len(e["avisos"]), 2)
        self.assertIn("hace 35 días", e["avisos"][0])
        self.assertIn("julio de 2026", e["avisos"][1])
        e = dm.estado(self.dir / "no.xlsx", omie, hoy=dt.date(2026, 8, 5))
        self.assertIn("No se encuentra", e["avisos"][0])
        self.assertEqual(len(e["avisos"]), 1)

    def test_congelar(self):
        omip = fichero_omip(self.dir / "omip.xlsx", [fila("2026-11-04", "FTB YR-27", 80.0)])
        omie = fichero_omie(self.dir / "omie.xlsx", {(2025, m): 50.0 for m in range(1, 13)})
        c = dm.congelar(dt.date(2027, 11, 1), 3, 0.95, 3, ruta_omip=omip, ruta_omie=omie)
        self.assertEqual(c["sesion_omip"], "2026-11-04")
        self.assertEqual([(f["anio"], f["mes"]) for f in c["meses"]],
                         [(2027, 11), (2027, 12), (2028, 1)])
        self.assertAlmostEqual(c["meses"][0]["previsto"], 0.95 * 80 + 0.05 * 50)
        self.assertEqual(c["meses"][2]["previsto"], 50.0)
        import json
        json.dumps(c)                                   # se puede guardar tal cual


if __name__ == "__main__":
    unittest.main()
