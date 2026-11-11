# -*- coding: utf-8 -*-
"""
CALCULO DEL COSTE DE CADA OFERTA (mes a mes)
============================================
Misma metodología que la plantilla Excel multipunto (generar_plantilla.py), con la duración
libre: se calcula cada mes del contrato (k = 1..duración) y cada periodo.

Para cada CUPS incluido, con el grupo de precios de su tarifa y sistema (o un grupo propio):
    consumo[k, p]  = perfil del mes natural × crecimiento del año de contrato
    energía        Fijo:          precio / unidad × (1 − descuento)
                   Pass-Pool:     (Ci × (mercado + regulados) + B) / 1000; si «el Ci se aplica
                                  a la media mensual», el mercado va sin apuntamiento
                   Pass-Through:  ((mercado + regulados) × (1 + pérdidas) + B) / 1000
                   mercado = (previsto × (1 − %cerrado) + precio cierre × %cerrado) × apuntamiento
                             + liquidación de los MW cerrados (base y pico) + clicks
                   Ph demanda (indexados con «Canarias al Ph», solo CUPS de Canarias):
                   mercado = (previsto + diferencial Ph − diario del mes) × apuntamiento del Ph
                             (el Ph ya lleva los costes de ajuste; sin diferencial calculado se
                             usa el SSAA del estudio)
                   regulados = pagos por capacidad + OS + OM + SRAD + desvíos (si la oferta los
                               añade); B en barras si su base es «Barras central»
    peajes         peajes + cargos de energía del año natural, si el precio no los incluye
    SSAA           según el modo (incluidos, por periodo, plano, banda, techo), con perfil,
                   elevación a pérdidas y factores 1 y 2
    otros          FNEE (incluido, neto o barras) + GdO + fee (neto o barras)
    potencia       kW × peaje+cargo del año natural del mes / 12 (o la potencia propia)
    K / IM         K × (energía + otros [+ SSAA])
    IEE            base × (1 − reducción) × tipo del año de contrato (sin mínimo por MWh)
    servicios      otros costes fijos + alquiler, por mes

El apuntamiento de cada CUPS: el propio de la oferta si lo tiene; si no, el del cliente (si
tiene curva válida y se ha elegido «Del cliente»); si no, el estimado del estudio; vacío = 1.
Sin Streamlit.
"""

import calendar
import datetime as dt
from dataclasses import dataclass, field

import numpy as np

import datos_regulados as dr
from analisis_ofertas import modelo as mo

CONCEPTOS = ["potencia", "energia", "peajes", "ssaa", "otros", "k", "iee", "servicios"]
NOMBRES = {"potencia": "Término de potencia", "energia": "Energía / mercado",
           "peajes": "Peajes y cargos de energía", "ssaa": "SSAA",
           "otros": "Otros (FNEE, GdO, fee)", "k": "K / impuesto municipal",
           "iee": "Impuesto eléctrico", "servicios": "Servicios y alquiler"}


# ------------------------------------------------------------------- entradas comunes
class Contexto:
    """Lo que no depende de la oferta: datos regulados por año, mercado y SSAA por mes."""

    def __init__(self, estudio: mo.Estudio, guardados_regulados: dict = None):
        self.e = estudio
        self._reg = {}
        self._guardados = guardados_regulados

    def regulados(self, anio: int) -> dict:
        if anio not in self._reg:
            if anio in self.e.regulados:
                d = self.e.regulados[anio]
            else:
                previos = [a for a in self.e.regulados if a <= anio]
                if previos:
                    d = dict(self.e.regulados[max(previos)], copia_de=max(previos))
                elif self.e.regulados:
                    d = dict(self.e.regulados[min(self.e.regulados)],
                             copia_de=min(self.e.regulados))
                else:
                    d = dr.datos_anio(anio, self._guardados)
            self._reg[anio] = d
        return self._reg[anio]

    def anios_copiados(self, anios) -> dict:
        """{año: año copiado} de los años sin datos regulados propios."""
        return {a: self.regulados(a)["copia_de"] for a in sorted(set(anios))
                if self.regulados(a).get("copia_de") not in (None, a)}

    def mercado(self, anio: int, mes: int):
        """(previsto, OMIP aplicado) €/MWh del mes: % OMIP × OMIP + (1 − %) × media OMIE; sin
        OMIP, la media de OMIE (como la plantilla)."""
        clave = f"{anio}-{mes:02d}"
        fila = self.e.mercado.get(clave) or {}
        omip = self.e.omip_manual.get(clave)
        if omip is None:
            omip = fila.get("omip")
        omie = fila.get("omie_media")
        if omip and omip > 0:
            if omie is None:
                return float(omip), float(omip)
            return self.e.pct_omip * omip + (1 - self.e.pct_omip) * omie, float(omip)
        return float(omie or 0.0), 0.0

    def ph_dif(self, anio: int, mes: int) -> float:
        """Lo que el Ph demanda de Canarias lleva encima del mercado diario (€/MWh): el
        calculado con ESIOS (12 últimos meses) o, si no hay, el SSAA del estudio."""
        v = self.e.ph_dif.get(mes) if self.e.ph_dif else None
        return float(v) if v is not None else self.ssaa(anio, mes)

    def ssaa(self, anio: int, mes: int) -> float:
        manual = self.e.ssaa_manual.get(f"{anio}-{mes:02d}")
        if manual is not None:
            return float(manual)
        return float(self.e.ssaa_mes[mes - 1] or 0.0) if self.e.ssaa_mes else 0.0


def _factores(valores) -> list:
    """Apuntamiento de 6 periodos: vacío o 0 = 1 (el precio base)."""
    return [float(v) if isinstance(v, (int, float)) and v else 1.0 for v in mo.seis(valores)]


def apuntamientos(e: mo.Estudio, o: mo.Oferta, sistema: str, cal: str, meses: list,
                  cups: str = None, ph: bool = False):
    """(OMIE, SSAA) de 6 periodos para cada mes [(año, mes)] → dos arrays D × 6. ph: el de
    mercado es el del Ph demanda de Canarias (el estimado «Canarias Ph» o, si no está, el de
    Canarias; el del cliente, el calculado con el Ph)."""
    clave = mo.clave_ap(sistema, cal)
    cliente = None
    if cups:
        s = e.suministro(cups)
        if s and s.apuntamiento == mo.CLIENTE:
            cliente = e.ap_cliente.get(cups)
    est_ss = e.ap_ssaa.get(clave) or e.ap_omie.get(clave) or {}
    if ph:
        clave_ph = mo.clave_ap(mo.SISTEMA_PH, cal)
        clave = clave_ph if e.ap_omie.get(clave_ph) else clave
        if cliente:
            cliente = {"omie": cliente["ph"]} if cliente.get("ph") else None
    propio = (o.apuntamientos_propios or {}).get(clave) or {}
    est_om = e.ap_omie.get(clave) or {}
    om, ss = [], []
    for _, m in meses:
        base_om = (cliente or {}).get("omie", {}).get(m) if cliente else None
        base_om = _factores(base_om if base_om else est_om.get(m))
        prop = mo.seis(propio.get(m)) if not ph else [None] * 6
        om.append([float(p) if isinstance(p, (int, float)) else b for p, b in zip(prop, base_om)])
        base_ss = (cliente or {}).get("ssaa", {}).get(m) if cliente else None
        ss.append(_factores(base_ss if base_ss else est_ss.get(m)))
    return np.array(om, float), np.array(ss, float)


# ------------------------------------------------------------------- resultados
@dataclass
class ResultadoOferta:
    oferta: mo.Oferta
    meses: list                                   # [(año, mes)] de cada mes de contrato
    consumo: np.ndarray                           # kWh por mes
    conceptos: dict                               # {concepto: € por mes}
    por_cups: dict = field(default_factory=dict)  # {cups: {"anios": [{concepto: €}], ...}}
    incidencias: dict = field(default_factory=dict)   # {cups: texto}
    apuntamiento_cliente: list = field(default_factory=list)   # CUPS con el suyo
    cierres: list = field(default_factory=list)   # detalle mensual de los cierres

    @property
    def total_mes(self) -> np.ndarray:
        return sum(self.conceptos[c] for c in CONCEPTOS)

    @property
    def total(self) -> float:
        return float(self.total_mes.sum())

    @property
    def mwh(self) -> float:
        return float(self.consumo.sum()) / 1000

    @property
    def n_meses(self) -> int:
        return len(self.meses)

    @property
    def media_anual(self) -> float:
        return self.total * 12 / self.n_meses if self.n_meses else 0.0

    @property
    def eur_mwh(self) -> float:
        return self.total / self.mwh if self.mwh else 0.0

    def por_anio(self) -> list:
        """[{concepto: €, "consumo": kWh, "meses": n}] por año de contrato (el último puede
        ser parcial)."""
        salida = []
        for i in range(0, self.n_meses, 12):
            d = {c: float(self.conceptos[c][i:i + 12].sum()) for c in CONCEPTOS}
            d["consumo"] = float(self.consumo[i:i + 12].sum())
            d["meses"] = min(12, self.n_meses - i)
            d["total"] = sum(d[c] for c in CONCEPTOS)
            salida.append(d)
        return salida

    def total_en(self, meses: list) -> float:
        """Coste en esos meses [(año, mes)] (los que no son del contrato no cuentan)."""
        idx = {am: i for i, am in enumerate(self.meses)}
        tm = self.total_mes
        return float(sum(tm[idx[am]] for am in meses if am in idx))


# ------------------------------------------------------------------- cálculo
def _crecimiento(e: mo.Estudio, n_anios: int) -> list:
    """Factor de consumo de cada año de contrato (1, 1+g2, (1+g2)(1+g3)...)."""
    f, salida = 1.0, [1.0]
    for y in range(2, n_anios + 1):
        g = e.crecimiento[y - 2] if len(e.crecimiento) >= y - 1 else 0.0
        f *= 1 + float(g or 0)
        salida.append(f)
    return salida


def _horas(anio: int, mes: int) -> int:
    return calendar.monthrange(anio, mes)[1] * 24        # como la plantilla (sin cambios de hora)


def _horas_pico(anio: int, mes: int) -> int:
    """Horas de carga pico del mes: 12 (8-20 h) por día de lunes a viernes."""
    dias = calendar.monthrange(anio, mes)[1]
    return sum(12 for d in range(1, dias + 1) if dt.date(anio, mes, d).weekday() < 5)


def usa_ph(o: mo.Oferta, sistema: str) -> bool:
    """¿El CUPS de ese sistema se calcula con el Ph demanda de Canarias en esta oferta?"""
    return bool(o.indice_ph and o.indexada and sistema == "Canarias")


def grupos_de(e: mo.Estudio, o: mo.Oferta) -> dict:
    """{nombre: (tarifa, sistema)} de los grupos de la oferta: los automáticos de las tarifas y
    sistemas de los suministros, y los propios."""
    g = {mo.grupo_automatico(s.tarifa, s.sistema): (s.tarifa, s.sistema)
         for s in e.suministros if s.activo}
    g.update({gp.nombre: (gp.tarifa, gp.sistema) for gp in o.grupos_propios if gp.nombre})
    return g


def incluidos(e: mo.Estudio, o: mo.Oferta):
    """[(suministro, grupo)] que entran en la oferta e {cups: incidencia}."""
    grupos = grupos_de(e, o)
    excluidos = set(o.excluidos or [])
    salida, incidencias = [], {}
    for s in e.suministros:
        if not s.activo or s.cups in excluidos:
            continue
        g = o.asignacion.get(s.cups) or mo.grupo_automatico(s.tarifa, s.sistema)
        if g not in grupos:
            incidencias[s.cups] = f"Grupo no válido ({g})"
            continue
        tar, sis = grupos[g]
        if sis != s.sistema:
            incidencias[s.cups] = "Sistema del grupo distinto al del CUPS"
        elif tar != s.tarifa:
            incidencias[s.cups] = "Tarifa del grupo distinta a la del CUPS"
        salida.append((s, g))
    return salida, incidencias


def calcular_oferta(e: mo.Estudio, o: mo.Oferta, ctx: Contexto = None,
                    n_meses: int = None) -> ResultadoOferta:
    """Coste de la oferta en su duración (o en n_meses: el contrato actual se proyecta al
    horizonte del estudio repitiendo los precios de su último año)."""
    ctx = ctx or Contexto(e)
    D = int(n_meses or o.duracion)
    inicio = o.inicio or e.inicio
    meses = [(f.year, f.month) for f in (mo.sumar_meses(inicio, k) for k in range(D))]
    anio_k = np.array([k // 12 + 1 for k in range(D)])
    crec = _crecimiento(e, int(anio_k.max()) if D else 1)
    f_crec = np.array([crec[y - 1] for y in anio_k])
    lista, incidencias = incluidos(e, o)
    grupos = grupos_de(e, o)

    # datos regulados de cada mes (año natural)
    regs = [ctx.regulados(a) for a, _ in meses]

    def por_mes(fn):
        return np.array([fn(r) for r in regs], float)

    # consumo de cada CUPS (D × 6) con el crecimiento
    consumo = {s.cups: np.array([s.consumo_mes(m) for _, m in meses], float) * f_crec[:, None]
               for s, _ in lista}

    # ---- cierres: consumo de la cartera en barras (MWh) de cada mes
    previsto = np.array([ctx.mercado(a, m)[0] for a, m in meses])
    omip = np.array([ctx.mercado(a, m)[1] for a, m in meses])
    horas = np.array([_horas(a, m) for a, m in meses], float)
    barras = np.zeros(D)
    for s, _ in lista:
        perd = np.array([dr.perdidas(r, s.tarifa, s.sistema) for r in regs])
        barras += (consumo[s.cups] * (1 + perd)).sum(axis=1)
    barras_mwh = barras / 1000
    cierres = {c.mes: c for c in o.cierres if c.mes}
    pct = np.zeros(D)
    precio_cierre = np.where(omip > 0, omip, previsto)
    liquidacion = np.zeros(D)
    clicks = np.zeros(D)
    horas_pico = np.array([_horas_pico(a, m) for a, m in meses], float)
    previsto_pico = np.array([previsto[k] * float(e.pico.get(m) or 1.0)
                              for k, (_, m) in enumerate(meses)], float)
    precio_pico = previsto_pico.copy()
    for k in range(D):
        c = cierres.get(k + 1)
        if c and c.precio is not None:
            precio_cierre[k] = c.precio
        if c and c.precio_pico is not None:
            precio_pico[k] = c.precio_pico
        if o.cierres_tipo == mo.CIERRE_PCT and c:
            pct[k] = min(1.0, float(c.pct or 0))
            clicks[k] = pct[k] * o.coste_click
        elif o.cierres_tipo == mo.CIERRE_MW and c and barras_mwh[k] > 0:
            mw, mwp = float(c.mw or 0), float(c.mw_pico or 0)
            liquidacion[k] = (mw * horas[k] * (precio_cierre[k] - previsto[k])
                              + mwp * horas_pico[k] * (precio_pico[k] - previsto_pico[k])
                              ) / barras_mwh[k]
            clicks[k] = (mw * horas[k] + mwp * horas_pico[k]) * o.coste_click / barras_mwh[k]
    base_mercado = previsto * (1 - pct) + precio_cierre * pct
    detalle_cierres = [{"mes": k + 1, "anio_mes": meses[k], "previsto": previsto[k],
                        "omip": omip[k], "pct": pct[k], "precio": precio_cierre[k],
                        "mw_pico": float((cierres.get(k + 1) or mo.Cierre(0)).mw_pico or 0),
                        "previsto_pico": previsto_pico[k], "precio_pico": precio_pico[k],
                        "horas_pico": horas_pico[k],
                        "barras_mwh": barras_mwh[k], "liquidacion": liquidacion[k],
                        "clicks": clicks[k]}
                       for k in range(D)] if o.cierres_tipo != mo.SIN_CIERRES else []

    # ---- matrices de coste unitario (€/kWh) de un grupo, D × 6
    ssaa = np.array([ctx.ssaa(a, m) for a, m in meses])
    dif_ph = np.array([ctx.ph_dif(a, m) for a, m in meses])
    unidad = mo.UNIDADES.get(o.unidad, 1)

    def matrices(tar, sis, gname, ap_om, ap_ss, ph=False):
        term = [dr.terminos(r, tar) for r in regs]
        perd = np.array([dr.perdidas(r, tar, sis) for r in regs])
        gen = [r["generales"] for r in regs]
        precios = [o.precios_grupo(gname, int(y), tar) for y in anio_k]

        def tabla(attr, defecto):
            return np.array([[defecto if v is None else float(v) for v in mo.seis(getattr(p, attr))]
                             for p in precios], float)
        regul = np.zeros((D, 6))
        if o.pagos_capacidad:
            regul += np.array([t["pagos_capacidad"] for t in term]) * 1000
        if o.srad:
            regul += np.array([t["srad"] or [0.0] * 6 for t in term]) * 1000
        regul += (np.array([g["os"] for g in gen]) * o.os + np.array([g["om"] for g in gen]) * o.om
                  + o.desvios)[:, None]
        if o.modalidad == mo.PASS_POOL and o.ci_media_mensual:
            ap_om = np.ones_like(ap_om)            # el Ci ya lleva el perfil
        base = base_mercado + (dif_ph if ph else 0.0)
        mercado = base[:, None] * ap_om + (liquidacion + clicks)[:, None]
        b = tabla("termino_b", 0.0) * ((1 + perd) if o.base_b == mo.BARRAS else 1)
        if o.modalidad == mo.FIJO:
            E = tabla("energia", 0.0) / unidad * (1 - o.descuento)
        elif o.modalidad == mo.PASS_POOL:
            E = (tabla("ci", 1.0) * (mercado + regul) + b) / 1000
        else:
            E = ((mercado + regul) * (1 + perd) + b) / 1000
        A = np.zeros((D, 6)) if o.atr_incluido else np.array([t["energia"] for t in term])
        s = ssaa[:, None] * np.ones((1, 6))
        sp = s * ap_ss if o.ssaa_perfil else s
        if o.ssaa_modo == mo.SSAA_INCLUIDOS:
            base = np.zeros((D, 6))
        elif o.ssaa_modo == mo.SSAA_PERIODO:
            base = s * ap_ss
        elif o.ssaa_modo == mo.SSAA_PLANO:
            base = s
        elif o.ssaa_modo == mo.SSAA_BANDA:
            base = np.where(sp < o.banda_min, sp - o.banda_min,
                            np.where(sp > o.banda_max, sp - o.banda_max, 0.0))
        else:
            base = np.maximum(0.0, sp - o.techo)
        lf = (1 + perd) if o.ssaa_perdidas == mo.PERD_TARIFA else (
            1 + o.ssaa_pct_perdidas if o.ssaa_perdidas == mo.PERD_FIJO else 1.0)
        S = base * lf * o.ssaa_f1 * o.ssaa_f2 / 1000
        fnee = np.array([g["fnee"] for g in gen])[:, None]
        otros = (0.0 if o.fnee == mo.FNEE_INCLUIDO else
                 fnee * ((1 + perd) if o.fnee == mo.BARRAS else 1.0))
        otros = otros + o.gdo + o.fee * ((1 + perd) if o.base_fee == mo.BARRAS else 1.0)
        O = np.broadcast_to(otros / 1000, (D, 6))
        pot = tabla("potencia", 0.0)
        return E, A, S, O, pot

    cache = {}
    conceptos = {c: np.zeros(D) for c in CONCEPTOS}
    consumo_total = np.zeros(D)
    por_cups, con_cliente = {}, []
    for s, gname in lista:
        tar, sis = grupos[gname]
        cal = mo.calendario_de(tar)
        ph = usa_ph(o, sis)
        propio = (s.apuntamiento == mo.CLIENTE and e.ap_cliente.get(s.cups) and o.indexada)
        if propio:
            con_cliente.append(s.cups)
            mats = matrices(tar, sis, gname,
                            *apuntamientos(e, o, sis, cal, meses, s.cups, ph=ph), ph=ph)
        else:
            if gname not in cache:
                cache[gname] = matrices(tar, sis, gname,
                                        *apuntamientos(e, o, sis, cal, meses, ph=ph), ph=ph)
            mats = cache[gname]
        E, A, S, O, pot_propia = mats
        c = consumo[s.cups]
        kw = np.array([float(v or 0) for v in mo.seis(s.potencias, 0.0)])
        if o.potencia == mo.POT_PROPIA:
            potencia = (pot_propia * kw).sum(axis=1) / 12
        else:
            potencia = np.array([np.dot(dr.terminos(r, s.tarifa)["potencia"], kw)
                                 for r in regs]) / 12
        energia = (c * E).sum(axis=1)
        peajes = (c * A).sum(axis=1)
        s_ss = (c * S).sum(axis=1)
        otros = (c * O).sum(axis=1)
        # K / impuesto municipal sobre energía y otros; la de los SSAA, con su propio valor
        k_im = o.k * (energia + otros) + (o.k_ssaa_valor * s_ss if o.k_ssaa else 0)
        servicios = np.full(D, o.otros_fijos + float(s.alquiler or 0))
        kwh = c.sum(axis=1)
        # IEE: tipo del año de contrato sobre la base (sin servicios ni alquiler); ya no se
        # aplica el importe mínimo por MWh
        base = potencia + energia + peajes + s_ss + otros + k_im
        iee = np.zeros(D)
        red = float(s.reduccion_iee or 0)
        for i in range(0, D, 12):
            sl = slice(i, min(D, i + 12))
            if red < 1:
                iee[sl] = base[sl] * (1 - red) * regs[i]["generales"]["iee"]
        filas = {"potencia": potencia, "energia": energia, "peajes": peajes, "ssaa": s_ss,
                 "otros": otros, "k": k_im, "iee": iee, "servicios": servicios}
        for cpt, v in filas.items():
            conceptos[cpt] += v
        consumo_total += kwh
        anios = []
        for i in range(0, D, 12):
            sl = slice(i, min(D, i + 12))
            d = {cpt: float(v[sl].sum()) for cpt, v in filas.items()}
            d["consumo"] = float(kwh[sl].sum())
            d["total"] = sum(d[cpt] for cpt in CONCEPTOS)
            anios.append(d)
        por_cups[s.cups] = {"grupo": gname, "anios": anios,
                            "total": sum(a["total"] for a in anios),
                            "consumo": sum(a["consumo"] for a in anios)}
    return ResultadoOferta(o, meses, consumo_total, conceptos, por_cups, incidencias,
                           con_cliente, detalle_cierres)


def calcular(e: mo.Estudio, ctx: Contexto = None) -> dict:
    """{"actual": ResultadoOferta (proyectado al horizonte), "ofertas": [ResultadoOferta]}."""
    ctx = ctx or Contexto(e)
    actual = None
    if e.actual:
        n = e.horizonte()
        actual = calcular_oferta(e, e.actual, ctx, n_meses=max(n, int(e.actual.duracion or 0)))
    return {"actual": actual,
            "ofertas": [calcular_oferta(e, o, ctx) for o in e.ofertas]}
