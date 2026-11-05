# -*- coding: utf-8 -*-
"""Motor del análisis de ofertas: un caso a mano por cada rama de fórmula (fijo con y sin
peajes, Pass-Pool, Pass-Through, cierres por % y por MW, los 5 modos de SSAA, FNEE y fee en
neto o barras, IEE con mínimo y reducción), duraciones libres, contrato actual en los mismos
meses, apuntamiento del cliente por CUPS, resultados y JSON. Datos ficticios."""

import copy
import datetime as dt
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import datos_regulados as dr
from analisis_ofertas import calculo, resultados
from analisis_ofertas import modelo as mo

CUPS_A, CUPS_B = "ES0000000000000000AA", "ES0000000000000001BB"


def regulados_simples():
    """Datos regulados de 2026 con números redondos para hacer las cuentas a mano."""
    d = dr.datos_anio(2026, {})
    d["generales"] = {"iee": 0.05, "iee_minimo_general": 1.0, "iee_minimo_reducido": 0.5,
                      "impuesto_municipal": 0.015, "fnee": 2.0, "om": 0.04, "os": 0.2}
    for t in dr.TARIFAS:
        n = dr.n_periodos(t)
        d["tarifas"][t].update({
            "peajes_potencia": [10.0, 5.0] + ([2.0] * 4 if n == 6 else []),
            "cargos_potencia": [2.0, 1.0] + ([1.0] * 4 if n == 6 else []),
            "peajes_energia": [0.02] * n, "cargos_energia": [0.01] * n,
            "pagos_capacidad": [0.001] * n, "perdidas": [0.1] * n,
            "srad": [0.0005] * n if n == 6 else None})
    d["perdidas_sistema"] = {}
    return d


def estudio(n_cups=1, inicio=dt.date(2027, 1, 1)):
    regs = {a: copy.deepcopy(regulados_simples()) for a in range(2027, 2033)}
    e = mo.Estudio(cliente="Cliente de prueba", inicio=inicio, regulados=regs, pct_omip=1.0)
    for i, cups in enumerate([CUPS_A, CUPS_B][:n_cups]):
        e.suministros.append(mo.Suministro(
            cups=cups, tarifa="6.1TD", potencias=[100.0] * 6,
            consumo={m: [1000.0, 0, 0, 0, 0, 1000.0] for m in range(1, 13)}))
    for k in range(72):
        f = mo.sumar_meses(inicio, k)
        e.mercado[f"{f.year}-{f.month:02d}"] = {"omip": 100.0, "omie_media": 50.0}
    e.ssaa_mes = [10.0] * 12
    return e


def oferta(**kw):
    o = mo.Oferta(nombre="Oferta", comercializadora="COMERCIALIZADORA A")
    for k, v in kw.items():
        setattr(o, k, v)
    return o


def fijo(precio=0.1, **kw):
    o = oferta(**kw)
    o.precios["6.1TD"] = [mo.PreciosAnio(energia=[precio] * 6)]
    return o


class Fijo(unittest.TestCase):
    def test_con_peajes_incluidos(self):
        e = estudio()
        r = calculo.calcular_oferta(e, fijo(0.1, k=0.015), calculo.Contexto(e))
        # 24.000 kWh/año × 0,1 €/kWh
        self.assertAlmostEqual(r.conceptos["energia"].sum(), 2400.0)
        self.assertAlmostEqual(r.conceptos["peajes"].sum(), 0.0)
        # potencia: 100 kW × (12+6+3+3+3+3) €/kW·año
        self.assertAlmostEqual(r.conceptos["potencia"].sum(), 100 * 30)
        self.assertAlmostEqual(r.conceptos["k"].sum(), 0.015 * 2400)
        base = 3000 + 2400 + 36
        self.assertAlmostEqual(r.conceptos["iee"].sum(), base * 0.05)
        self.assertEqual((r.n_meses, round(r.mwh, 6)), (12, 24.0))

    def test_sin_peajes_unidad_y_descuento(self):
        e = estudio()
        r = calculo.calcular_oferta(e, fijo(100.0, unidad="€/MWh", atr_incluido=False,
                                            descuento=0.1), calculo.Contexto(e))
        self.assertAlmostEqual(r.conceptos["energia"].sum(), 24000 * 0.1 * 0.9)
        self.assertAlmostEqual(r.conceptos["peajes"].sum(), 24000 * 0.03)

    def test_iee_minimo_y_reduccion(self):
        e = estudio()
        e.suministros[0].potencias = [0.0] * 6
        r = calculo.calcular_oferta(e, fijo(0.001), calculo.Contexto(e))
        # base 24 € × 5 % = 1,2 € < 24 MWh × 1 €/MWh
        self.assertAlmostEqual(r.conceptos["iee"].sum(), 24.0)
        e.suministros[0].reduccion_iee = 0.85
        r = calculo.calcular_oferta(e, fijo(0.001), calculo.Contexto(e))
        self.assertAlmostEqual(r.conceptos["iee"].sum(), 12.0)          # mínimo reducido
        e.suministros[0].reduccion_iee = 1.0
        self.assertEqual(calculo.calcular_oferta(e, fijo(0.001)).conceptos["iee"].sum(), 0.0)

    def test_potencia_propia_servicios_y_alquiler(self):
        e = estudio()
        e.suministros[0].alquiler = 20.0
        o = fijo(0.1, potencia=mo.POT_PROPIA, otros_fijos=5.0)
        o.precios["6.1TD"][0].potencia = [30.0] + [0.0] * 5
        r = calculo.calcular_oferta(e, o)
        self.assertAlmostEqual(r.conceptos["potencia"].sum(), 100 * 30)
        self.assertAlmostEqual(r.conceptos["servicios"].sum(), 12 * 25)


class Indexado(unittest.TestCase):
    def test_pass_through(self):
        e = estudio()
        o = oferta(modalidad=mo.PASS_THROUGH, atr_incluido=False, pagos_capacidad=True,
                   os=True, om=True, srad=True, desvios=1.0)
        o.precios["6.1TD"] = [mo.PreciosAnio(termino_b=[3.0] * 6)]
        r = calculo.calcular_oferta(e, o)
        reg = 1.0 + 0.2 + 0.04 + 0.5 + 1.0                 # pc + OS + OM + SRAD + desvíos
        unit = ((100.0 + reg) * 1.1 + 3.0) / 1000
        self.assertAlmostEqual(r.conceptos["energia"].sum(), 24000 * unit)
        self.assertAlmostEqual(r.conceptos["peajes"].sum(), 24000 * 0.03)

    def test_pass_pool_omip_omie_y_b_en_barras(self):
        e = estudio()
        e.pct_omip = 0.9                                   # 0,9 × 100 + 0,1 × 50 = 95
        o = oferta(modalidad=mo.PASS_POOL, base_b=mo.BARRAS)
        o.precios["6.1TD"] = [mo.PreciosAnio(termino_b=[2.0] * 6, ci=[1.05] * 6)]
        r = calculo.calcular_oferta(e, o)
        unit = (1.05 * 95.0 + 2.0 * 1.1) / 1000
        self.assertAlmostEqual(r.conceptos["energia"].sum(), 24000 * unit)

    def test_apuntamiento_estimado_y_propio_de_la_oferta(self):
        e = estudio()
        e.ap_omie[mo.clave_ap("Península", "6P")] = {m: [1.2, None, None, None, None, 0.8]
                                                     for m in range(1, 13)}
        o = oferta(modalidad=mo.PASS_POOL)
        r = calculo.calcular_oferta(e, o)
        self.assertAlmostEqual(r.conceptos["energia"].sum(), 12000 * 0.12 + 12000 * 0.08)
        o.apuntamientos_propios[mo.clave_ap("Península", "6P")] = {1: [2.0] + [None] * 5}
        r = calculo.calcular_oferta(e, o)
        self.assertAlmostEqual(r.conceptos["energia"][0], 1000 * 0.2 + 1000 * 0.08)

    def test_cierres_por_porcentaje(self):
        e = estudio()
        o = oferta(modalidad=mo.PASS_POOL, cierres_tipo=mo.CIERRE_PCT, coste_click=0.5,
                   cierres=[mo.Cierre(1, pct=0.5, precio=80.0)])
        r = calculo.calcular_oferta(e, o)
        enero = (100 * 0.5 + 80 * 0.5 + 0.5 * 0.5) / 1000 * 2000
        self.assertAlmostEqual(r.conceptos["energia"][0], enero)
        self.assertAlmostEqual(r.conceptos["energia"][1], 2000 * 0.1)

    def test_cierres_por_mw(self):
        e = estudio()
        o = oferta(modalidad=mo.PASS_POOL, cierres_tipo=mo.CIERRE_MW, coste_click=1.0,
                   cierres=[mo.Cierre(1, mw=0.001, precio=90.0)])
        r = calculo.calcular_oferta(e, o)
        barras = 2000 * 1.1 / 1000                          # MWh en barras de enero
        horas = 31 * 24
        liq = 0.001 * horas * (90 - 100) / barras
        click = 0.001 * horas * 1.0 / barras
        self.assertAlmostEqual(r.conceptos["energia"][0], 2000 * (100 + liq + click) / 1000)
        self.assertEqual(len(r.cierres), 12)
        self.assertAlmostEqual(r.cierres[0]["barras_mwh"], barras)


class Ssaa(unittest.TestCase):
    def ssaa(self, **kw):
        e = estudio()
        e.ap_ssaa[mo.clave_ap("Península", "6P")] = {m: [1.5, None, None, None, None, 0.5]
                                                     for m in range(1, 13)}
        return calculo.calcular_oferta(e, fijo(0.1, **kw)).conceptos["ssaa"].sum()

    def test_modos(self):
        self.assertEqual(self.ssaa(ssaa_modo=mo.SSAA_INCLUIDOS), 0.0)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_PLANO), 24000 * 0.01)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_PERIODO),
                               12000 * 0.015 + 12000 * 0.005)
        # banda 8-12 sin perfil: 10 está dentro → 0; con perfil: 15 (+3) y 5 (−3)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_BANDA, banda_min=8, banda_max=12), 0)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_BANDA, banda_min=8, banda_max=12,
                                         ssaa_perfil=True), 12000 * 0.003 - 12000 * 0.003)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_TECHO, techo=6), 24000 * 0.004)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_TECHO, techo=6, ssaa_perfil=True),
                               12000 * 0.009)

    def test_perdidas_y_factores(self):
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_PLANO, ssaa_perdidas=mo.PERD_TARIFA,
                                         ssaa_f1=1.02, ssaa_f2=1.015),
                               24000 * 0.01 * 1.1 * 1.02 * 1.015)
        self.assertAlmostEqual(self.ssaa(ssaa_modo=mo.SSAA_PLANO, ssaa_perdidas=mo.PERD_FIJO,
                                         ssaa_pct_perdidas=0.05), 24000 * 0.01 * 1.05)

    def test_k_sobre_ssaa(self):
        e = estudio()
        o = fijo(0.1, ssaa_modo=mo.SSAA_PLANO, k=0.02, k_ssaa=True)
        r = calculo.calcular_oferta(e, o)
        self.assertAlmostEqual(r.conceptos["k"].sum(), 0.02 * (2400 + 240))


class Otros(unittest.TestCase):
    def test_fnee_gdo_y_fee(self):
        e = estudio()
        r = calculo.calcular_oferta(e, fijo(0.1, fnee=mo.BARRAS, gdo=1.0, fee=2.0,
                                            base_fee=mo.BARRAS))
        self.assertAlmostEqual(r.conceptos["otros"].sum(), 24000 * (2.0 * 1.1 + 1.0 + 2.2) / 1000)
        r = calculo.calcular_oferta(e, fijo(0.1, fnee=mo.NETO))
        self.assertAlmostEqual(r.conceptos["otros"].sum(), 24000 * 2.0 / 1000)


class Duraciones(unittest.TestCase):
    def test_6_18_33_y_60_meses(self):
        e = estudio()
        e.crecimiento = [0.1]                              # año 2: +10 %; años siguientes, igual
        for meses in (6, 18, 33, 60):
            r = calculo.calcular_oferta(e, fijo(0.1, duracion=meses))
            self.assertEqual(r.n_meses, meses)
            anios = r.por_anio()
            self.assertEqual(len(anios), -(-meses // 12))
            self.assertEqual(anios[-1]["meses"], meses - 12 * (len(anios) - 1))
            self.assertAlmostEqual(r.media_anual, r.total * 12 / meses)
        r18 = calculo.calcular_oferta(e, fijo(0.1, duracion=18))
        self.assertAlmostEqual(r18.consumo.sum(), 24000 + 12000 * 1.1)
        # potencia: mes a mes, sin redondear a años
        self.assertAlmostEqual(r18.conceptos["potencia"].sum(), 3000 * 1.5)

    def test_precio_de_cada_año_de_contrato(self):
        e = estudio()
        o = fijo(0.1, duracion=30)
        o.precios["6.1TD"].append(mo.PreciosAnio(energia=[0.2] * 6))
        r = calculo.calcular_oferta(e, o)
        anios = r.por_anio()
        self.assertAlmostEqual(anios[0]["energia"], 2400)
        self.assertAlmostEqual(anios[1]["energia"], 4800)
        self.assertAlmostEqual(anios[2]["energia"], 2400)  # año 3 parcial: precio del año 2

    def test_actual_en_los_mismos_meses(self):
        e = estudio()
        e.actual = fijo(0.15, nombre="Contrato actual", es_actual=True, duracion=12)
        e.ofertas = [fijo(0.1, nombre="A", duracion=18),
                     fijo(0.12, nombre="B", duracion=6, inicio=dt.date(2027, 7, 1))]
        self.assertEqual(e.horizonte(), 18)
        res = calculo.calcular(e)
        self.assertEqual(res["actual"].n_meses, 18)        # se proyecta al horizonte
        filas = resultados.comparativa(res)
        a, b = filas
        actual_a = res["actual"].total * 12 / 18
        self.assertAlmostEqual(a["actual_mismo_periodo"], actual_a)
        meses_b = [(2027, m) for m in range(7, 13)]
        self.assertAlmostEqual(b["actual_mismo_periodo"],
                               res["actual"].total_en(meses_b) * 2)
        self.assertLess(a["variacion"], 0)
        self.assertEqual(sorted(f["ranking"] for f in filas), [1, 2])
        self.assertEqual(resultados.mejores(filas)["Fijo"]["nombre"], "A")


class Grupos(unittest.TestCase):
    def test_territoriales_propios_y_excluidos(self):
        e = estudio(2)
        e.suministros[1].sistema = "Canarias"
        o = fijo(0.1)
        r = calculo.calcular_oferta(e, o)
        self.assertEqual(r.por_cups[CUPS_B]["grupo"], "6.1TD-Canarias")
        self.assertAlmostEqual(r.por_cups[CUPS_B]["anios"][0]["energia"], 2400)  # copia la tarifa
        o.precios["6.1TD-Canarias"] = [mo.PreciosAnio(energia=[0.2] * 6)]
        o.grupos_propios.append(mo.GrupoPropio("Especial", "6.1TD", "Península"))
        o.precios["Especial"] = [mo.PreciosAnio(energia=[0.05] * 6)]
        o.asignacion[CUPS_A] = "Especial"
        r = calculo.calcular_oferta(e, o)
        self.assertAlmostEqual(r.por_cups[CUPS_B]["anios"][0]["energia"], 4800)
        self.assertAlmostEqual(r.por_cups[CUPS_A]["anios"][0]["energia"], 1200)
        o.asignacion[CUPS_A] = "No existe"
        o.excluidos = [CUPS_B]
        r = calculo.calcular_oferta(e, o)
        self.assertEqual(r.por_cups, {})
        self.assertIn("Grupo no válido", r.incidencias[CUPS_A])


class ApuntamientoCliente(unittest.TestCase):
    def test_por_cups_y_solo_en_indexadas(self):
        e = estudio(2)
        clave = mo.clave_ap("Península", "6P")
        e.ap_omie[clave] = {m: [1.0] * 6 for m in range(1, 13)}
        e.ap_cliente[CUPS_A] = {"omie": {m: [1.3] * 6 for m in range(1, 13)}}
        pp = oferta(modalidad=mo.PASS_POOL)
        base = calculo.calcular_oferta(e, pp)
        e.suministros[0].apuntamiento = mo.CLIENTE
        con = calculo.calcular_oferta(e, pp)
        self.assertEqual(con.apuntamiento_cliente, [CUPS_A])
        self.assertAlmostEqual(con.por_cups[CUPS_A]["anios"][0]["energia"],
                               base.por_cups[CUPS_A]["anios"][0]["energia"] * 1.3)
        self.assertAlmostEqual(con.por_cups[CUPS_B]["anios"][0]["energia"],
                               base.por_cups[CUPS_B]["anios"][0]["energia"])
        self.assertEqual(calculo.calcular_oferta(e, fijo(0.1)).apuntamiento_cliente, [])
        # apuntamiento del cliente igual a 1 en todo = estimado igual a 1
        e.ap_cliente[CUPS_A] = {"omie": {m: [1.0] * 6 for m in range(1, 13)}}
        self.assertAlmostEqual(calculo.calcular_oferta(e, pp).total, base.total)
        e.actual = oferta(nombre="Actual", modalidad=mo.PASS_POOL)
        e.ofertas = [pp]
        e.ap_cliente[CUPS_A] = {"omie": {m: [1.3] * 6 for m in range(1, 13)}}
        comp = resultados.apuntamientos(e)
        self.assertEqual(len(comp), 1)
        self.assertGreater(comp[0]["diferencia"], 0)


class Resultados(unittest.TestCase):
    def test_escenarios_cuadro_y_rondas(self):
        e = estudio()
        e.actual = fijo(0.15, nombre="Actual", es_actual=True)
        e.ofertas = [fijo(0.1, nombre="F"), oferta(nombre="PP", modalidad=mo.PASS_POOL)]
        esc = resultados.escenarios(e, (-0.1, 0.0, 0.1))
        f_pp = [x["filas"][1]["media_anual"] for x in esc]
        f_fijo = [x["filas"][0]["media_anual"] for x in esc]
        self.assertLess(f_pp[0], f_pp[1])
        self.assertLess(f_pp[1], f_pp[2])
        self.assertAlmostEqual(f_fijo[0], f_fijo[2])
        cuadro = resultados.cuadro_mercado(e)
        self.assertEqual([c["trimestre"] for c in cuadro], ["Q1-27", "Q2-27", "Q3-27", "Q4-27"])
        self.assertAlmostEqual(cuadro[0]["previsto"], 100.0)
        r2 = mo.duplicar(e.ofertas[0], nueva_ronda=True)
        r2.precios["6.1TD"][0].energia = [0.09] * 6
        e.ofertas.append(r2)
        filas = resultados.comparativa(calculo.calcular(e))
        rondas = resultados.rondas(filas)
        self.assertEqual([f["ronda"] for f in rondas[0]["rondas"]], [1, 1, 2])
        self.assertEqual(e.ofertas[0].precios["6.1TD"][0].energia[0], 0.1)   # copia aparte


class Json(unittest.TestCase):
    def test_ida_y_vuelta(self):
        import json
        e = estudio(2)
        e.actual = fijo(0.15, es_actual=True, inicio=dt.date(2027, 1, 1))
        o = oferta(modalidad=mo.PASS_THROUGH, cierres_tipo=mo.CIERRE_MW,
                   cierres=[mo.Cierre(2, mw=1.5)], gestor="Ana Gestora",
                   email_gestor="ana@example.com")
        o.grupos_propios.append(mo.GrupoPropio("X", "6.1TD"))
        o.apuntamientos_propios["Península|6P"] = {3: [1.1] * 6}
        e.ofertas = [o]
        e.ap_omie["Península|6P"] = {1: [1.0] * 6}
        e.ap_cliente[CUPS_A] = {"omie": {1: [1.2] * 6}, "fuentes": {1: "ene-26 (ESIOS)"}}
        texto = json.dumps(mo.estudio_a_dict(e))
        d = mo.estudio_de_dict(json.loads(texto))
        self.assertEqual(d.ofertas[0].cierres[0].mw, 1.5)
        self.assertEqual(d.ofertas[0].email_gestor, "ana@example.com")
        self.assertEqual(d.ap_omie["Península|6P"][1], [1.0] * 6)
        self.assertEqual(d.ap_cliente[CUPS_A]["omie"][1], [1.2] * 6)
        self.assertEqual(d.ofertas[0].apuntamientos_propios["Península|6P"][3], [1.1] * 6)
        self.assertEqual(d.suministros[0].consumo[5], [1000.0, 0, 0, 0, 0, 1000.0])
        self.assertEqual(d.actual.inicio, dt.date(2027, 1, 1))
        self.assertEqual(sorted(d.regulados), list(range(2027, 2033)))
        self.assertAlmostEqual(calculo.calcular_oferta(d, d.ofertas[0]).total,
                               calculo.calcular_oferta(e, o).total)


if __name__ == "__main__":
    unittest.main()
