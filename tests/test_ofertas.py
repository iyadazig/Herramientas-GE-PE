# -*- coding: utf-8 -*-
"""Excel de solicitud de ofertas (datos ficticios)."""

import datetime as dt
import io
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

from ofertas import correo
from ofertas import solicitud as so


def factura(anio, mes, dia, p1=0.0, p6=0.0):
    return {"Fin factura": dt.date(anio, mes, dia), "P1 kWh": p1, "P2 kWh": 0, "P3 kWh": 0,
            "P4 kWh": 0, "P5 kWh": 0, "P6 kWh": p6}


class Datos(unittest.TestCase):
    def test_alias_cif_y_discriminacion(self):
        self.assertEqual(so.alias_de("Empresa de prueba SA_TC4"), "TC4")
        self.assertEqual(so.alias_de("Empresa de prueba_Nave_2 Norte"), "Nave_2 Norte")
        self.assertEqual(so.alias_de("Sin alias"), "")
        self.assertEqual(so.cif_limpio("a-00.000.000"), "A00000000")
        self.assertEqual((so.discriminacion("2.0TD"), so.discriminacion("6.1TD")), ("DH3", "DH6"))

    def test_perfil_por_mes_de_fin_de_factura(self):
        # diciembre facturado en dos facturas (alta a mitad de mes): se suman
        perfil, aviso = so.perfil_mensual([factura(2025, 12, 15, p1=10), factura(2025, 12, 31, p1=5),
                                           factura(2026, 1, 31, p6=7)])
        self.assertEqual(perfil[12][0], 15.0)
        self.assertEqual(perfil[1][5], 7.0)
        self.assertEqual(perfil[6], [0.0] * 6)
        self.assertEqual(aviso, "")

    def test_aviso_si_se_mezclan_anios(self):
        _, aviso = so.perfil_mensual([factura(2025, 3, 31, p1=1), factura(2026, 3, 31, p1=1)])
        self.assertIn("Mar", aviso)

    def test_nombres_de_hojas(self):
        self.assertEqual(so.nombres_de_hojas([{"alias": "TC1"}, {"alias": "",
                                                                 "cups": "ES0000000000000000XX"}]),
                         ["1. TC1", "2. 000000XX"])


class Excel(unittest.TestCase):
    def test_libro_completo(self):
        perfil, _ = so.perfil_mensual([factura(2026, m, 28, p1=100, p6=50) for m in range(1, 13)])
        s = {"grupo": "GRUPO DE PRUEBA", "alias": "TC1", "cups": "ES0000000000000000XX",
             "razon": "EMPRESA DE PRUEBA SA",
             "direccion": "Calle Falsa 1", "cif": "A00000000", "tarifa": "6.1TD",
             "fecha_inicio": "31/12/2026", "potencias": [100, 100, 100, 100, 100, 200],
             "perfil": perfil}
        cab = {"ref": "R-1", "fecha": "01/01/2026", "gestor": "Persona de prueba",
               "telefono": so.TELEFONO_GEYPE, "email": "prueba@example.com", "producto": "X"}
        wb = openpyxl.load_workbook(io.BytesIO(so.excel_solicitud(cab, [s])))
        self.assertEqual(wb.sheetnames, ["Resumen", "1. TC1"])
        h = wb["1. TC1"]
        self.assertEqual((h["C6"].value, h["D8"].value, h["H9"].value, h["D10"].value),
                         ("R-1", "Persona de prueba", "prueba@example.com", "X"))
        self.assertEqual((h["E12"].value, h["D13"].value, h["D14"].value, h["H14"].value),
                         ("EMPRESA DE PRUEBA SA", "A00000000", "ES0000000000000000XX", "TC1"))
        self.assertEqual([h[c + "21"].value for c in "DEFGHI"], [100, 100, 100, 100, 100, 200])
        self.assertEqual((h["B25"].value, h["C25"].value, h["H25"].value), ("Ene / Jan", 100, 50))
        self.assertEqual(h["I25"].value, "=SUM(C25:H25)")
        self.assertEqual(h["H23"].value, "DH6")
        self.assertEqual(len(h._charts), 1)
        r = wb["Resumen"]
        self.assertEqual([r.cell(4, n).value for n in range(1, 9)],
                         ["GRUPO DE PRUEBA", "TC1", "ES0000000000000000XX", "EMPRESA DE PRUEBA SA",
                          "Calle Falsa 1", "A00000000", "6.1TD", "31/12/2026"])
        self.assertEqual((r["A2"].value, r["I2"].value, r["O2"].value),
                         ("Cliente/Grupo", "POTENCIAS A CONTRATAR (kW)", "PREVISIÓN DE CONSUMO ANUAL"))
        self.assertEqual([r.cell(4, n).value for n in range(9, 15)], [100] * 5 + [200])
        self.assertEqual(r["O4"].value, "='1. TC1'!C37")
        self.assertEqual(r["U4"].value, "='1. TC1'!I37")
        # fila TOTAL bajo el ultimo suministro
        self.assertEqual((r["T5"].value, r["U5"].value), ("TOTAL", "=SUM(U4:U4)"))
        self.assertEqual([r.cell(5, n).value for n in range(1, 20)], [None] * 19)
        self.assertFalse([m for m in r.merged_cells.ranges if m.min_row == 5])

    def test_consideraciones_en_la_hoja(self):
        c = {"indexado": ["Indexada Omie Pass Through", "Indexada Omie Pass Pool"],
             "inicio": "01/01/2027", "duracion": "31/12/2027", "duracion2": "31/12/2028",
             "temporada": False, "exencion": False, "gdo": True, "excedentes": False}
        cab = {"recepcion_desde": "22/10/2026", "recepcion_hasta": "23/10/2026",
               "recepcion_hora": "12:00"}
        h = openpyxl.load_workbook(io.BytesIO(so.excel_solicitud(
            cab, [{"cups": "ES0000000000000000XX", "tarifa": "2.0TD",
                   "consideraciones": c}])))["1. 000000XX"]
        valores = {celda.value for fila in h.iter_rows(min_row=61) for celda in fila}
        self.assertEqual(h["B61"].value[:32], "Tipología de ofertas solicitadas")
        # el precio fijo sale de la tarifa
        self.assertEqual((h["B62"].value, h["D62"].value), ("Precio fijo / Fixed price",
                                                            "Binómica en DH2"))
        self.assertEqual((h["D63"].value, h["D64"].value, h["D65"].value),
                         ("Indexada Omie Pass Through", "Indexada Omie Pass Pool", None))
        self.assertEqual(h["E67"].value, "Anual / Annual")
        self.assertEqual(h["E70"].value, "01/01/2027")
        self.assertEqual((h["F73"].value, h["F74"].value), ("31/12/2027", "31/12/2028"))
        self.assertEqual((h["B76"].value[:21], h["E76"].value),
                         ("Garantías de origen (", "Sí / Yes"))
        self.assertEqual((h["F79"].value, h["F80"].value), ("22/10/2026", "23/10/2026"))
        self.assertEqual(h["G80"].value, "antes de las 12:00 h / before 12:00")
        self.assertEqual(h["F80"].fill.fgColor.rgb[-6:], so.NARANJA)
        self.assertFalse(any("exención" in str(v) or "excedentes" in str(v) for v in valores))

    def test_precio_fijo_segun_tarifa(self):
        self.assertEqual([so.precio_fijo(t) for t in ("2.0TD", "3.0TD", "6.1TD", "6.4TD")],
                         ["Binómica en DH2"] + ["Binómica en DH6"] * 3)


class TablaConsideraciones(unittest.TestCase):
    def test_comunes_y_cambio_por_suministro(self):
        from ofertas import paginas as op
        comun = {o: False for o in op.COLUMNAS_CASILLA}
        comun.update({"PVPC": True, op.INICIO: dt.date(2027, 1, 1), op.DURACION: None,
                      op.DURACION2: None, op.DURACION3: None, op.PCT_CONSUMO: None, op.PCT_TIPO: None})
        t = op.tabla_consideraciones([{"cups": "A"}, {"cups": "B"}], comun)
        t.loc[1, op.GDO] = True
        t.loc[1, op.EXENCION] = True
        t.loc[1, op.PCT_CONSUMO] = 50.0
        a, b = (op.consideraciones_de(f) for _, f in t.iterrows())
        self.assertEqual((a["indexado"], a["inicio"], a["gdo"]), (["PVPC"], "01/01/2027", False))
        self.assertEqual((a["exencion"], a["pct_consumo_exento"]), (False, None))
        self.assertEqual((b["gdo"], b["exencion"], b["pct_consumo_exento"], b["duracion"]),
                         (True, True, 50.0, ""))


class Correo(unittest.TestCase):
    def suministro(self, cups, tarifa="6.1TD", **c):
        base = {"indexado": ["PVPC"], "inicio": "01/01/2027", "duracion": "31/12/2027",
                "duracion2": "", "duracion3": "", "temporada": False, "exencion": False,
                "gdo": False, "excedentes": False}
        base.update(c)
        return {"grupo": "GRUPO DE PRUEBA", "razon": "EMPRESA DE PRUEBA SA", "cups": cups,
                "tarifa": tarifa, "perfil": {1: [1000, 0, 0, 0, 0, 500]},
                "consideraciones": base}

    cab = {"email": "prueba@example.com", "recepcion_desde": "22/10/2026",
           "recepcion_hasta": "23/10/2026", "recepcion_hora": "12:00"}

    def test_igual_en_todos(self):
        texto = correo.texto_correo(self.cab, [
            self.suministro("ES0000000000000000XX", gdo=True, duracion2="31/12/2028",
                            duracion3="31/12/2029"),
            self.suministro("ES0000000000000001XX", gdo=True, duracion2="31/12/2028",
                            duracion3="31/12/2029")])
        self.assertIn("del cliente GRUPO DE PRUEBA en España", texto)
        self.assertIn("• CUPS: ES0000000000000000XX (EMPRESA DE PRUEBA SA)\n", texto)
        self.assertIn("consumo anual previsto total de 3.000 kWh", texto)
        self.assertIn("• Precio indexado: PVPC.\n", texto)
        self.assertIn("• Tipo de contrato: anual.\n", texto)
        self.assertIn("inicio del contrato: 1 de enero de 2027.\n", texto)
        self.assertIn("Duración del contrato: tres alternativas, hasta el 31 de diciembre de "
                      "2027, hasta el 31 de diciembre de 2028 y hasta el 31 de diciembre de "
                      "2029.", texto)
        self.assertIn("deben incluir garantías de origen (GdO).", texto)
        self.assertIn("Fecha máxima: 23 de octubre de 2026, antes de las 12:00 horas.", texto)
        self.assertIn("\nprueba@example.com\n", texto)
        self.assertNotIn("inómica", texto)              # el precio fijo no se menciona
        self.assertNotIn("—", texto)

    def test_exencion_y_excedentes_junto_al_cups(self):
        texto = correo.texto_correo(self.cab, [
            self.suministro("ES0000000000000000XX", exencion=True, pct_consumo_exento=85.0,
                            excedentes=True),
            self.suministro("ES0000000000000001XX", exencion=True, pct_consumo_exento=85.0)])
        self.assertIn("ES0000000000000000XX (EMPRESA DE PRUEBA SA) — sujeto a exención del "
                      "impuesto eléctrico (85 % del consumo); con compensación de excedentes\n",
                      texto)
        self.assertIn("ES0000000000000001XX (EMPRESA DE PRUEBA SA) — sujeto a exención del "
                      "impuesto eléctrico (85 % del consumo)\n", texto)
        resumen = texto.split("Resumen de la solicitud:")[1].split("Adjuntamos")[0]
        self.assertNotIn("exención", resumen)
        self.assertNotIn("excedentes", resumen)

    def test_lo_distinto_en_el_resumen_con_sus_cups(self):
        pt = ["Indexada Omie Pass Through", "Indexada Omie Pass Pool"]
        texto = correo.texto_correo(self.cab, [
            self.suministro("ES0000000000000000XX", indexado=pt, gdo=True),
            self.suministro("ES0000000000000001XX", indexado=pt, gdo=True,
                            inicio="01/03/2027"),
            self.suministro("ES0000000000000002XX", indexado=["PVPC"], temporada=True)])
        self.assertIn("• Precio indexado: Indexada Omie Pass Through y Indexada Omie Pass Pool "
                      "(ES0000000000000002XX: PVPC).", texto)
        self.assertIn("• Tipo de contrato: anual (ES0000000000000002XX: de temporada).", texto)
        self.assertIn("inicio del contrato: 1 de enero de 2027 (ES0000000000000001XX: 1 de marzo "
                      "de 2027).", texto)
        self.assertIn("garantías de origen (GdO), salvo en ES0000000000000002XX.", texto)
        self.assertNotIn("—", texto)


if __name__ == "__main__":
    unittest.main()
