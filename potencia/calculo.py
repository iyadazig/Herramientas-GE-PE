"""
Cálculo del coste de la potencia (término fijo + excesos) y optimización.

Metodología (art. 9 de la Circular 3/2020 de la CNMC, redacción dada por la
Circular 1/2025, vigente desde el 1/4/2025):

  Término fijo  = Σp Pc_p · Tp_p · días / 365          (Tp = peajes + cargos)
  Excesos (puntos de medida tipo 1, 2 y 3; cada mes de facturación):
      FEP = Σp tep_p · √( Σj (Pd_j − Pc_p)² )
      sumando solo los cuartos de hora j del periodo p con Pd_j > Pc_p.

La potencia óptima minimiza la suma anual con la restricción normativa
P1 ≤ P2 ≤ … ≤ P6. Cada periodo es una función convexa de su potencia, así que
el óptimo entero exacto se obtiene evaluando todas las potencias enteras y
aplicando el algoritmo "pool adjacent violators" para la restricción de orden.
"""

import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

from .calendario import asignar_periodos

N_P = 6
MESES_ABR = ["ene.", "feb.", "mar.", "abr.", "may.", "jun.",
             "jul.", "ago.", "sep.", "oct.", "nov.", "dic."]
MESES_NOMBRE = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def ruta_recurso(relativa):
    """Ruta a un recurso del proyecto (también dentro de un ejecutable PyInstaller)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relativa


def cargar_precios(ruta=None):
    """Precios regulados. En Herramientas GE&PE están en potencia/precios.json, copia del
    config/precios.json del programa de escritorio EstudioPotencia: se actualizan los dos cada año."""
    ruta = Path(ruta) if ruta else Path(__file__).resolve().parent / "precios.json"
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


@dataclass
class Tarifa:
    nombre: str
    tp: np.ndarray            # €/kW·año (peajes + cargos)
    tep: np.ndarray           # €/kW (término de exceso de potencia, tipos 1-3)
    derechos: dict            # precios de acceso, extensión y enganche
    dias_anio: int = 365

    @classmethod
    def desde_precios(cls, precios, nombre, nivel_tension=None):
        t = precios["tarifas"][nombre]
        tp = np.array(t["peajes_potencia"], float) + np.array(t["cargos_potencia"], float)
        nivel = nivel_tension or t["nivel_tension"]
        return cls(nombre, tp, np.array(t["tep"], float), dict(precios["derechos"][nivel]),
                   int(precios.get("dias_anio", 365)))


@dataclass
class Coste:
    fijo: np.ndarray          # (meses, 6) €
    exceso: np.ndarray        # (meses, 6) €

    @property
    def total_mes(self):
        return self.fijo.sum(1) + self.exceso.sum(1)

    @property
    def total_fijo(self):
        return float(self.fijo.sum())

    @property
    def total_exceso(self):
        return float(self.exceso.sum())

    @property
    def total(self):
        return self.total_fijo + self.total_exceso


class Estudio:
    """Curva cuartohoraria + calendario + tarifa, con los cálculos de coste."""

    def __init__(self, curva, zona, tarifa, festivos=None):
        festivos = festivos or {}
        self.curva = curva.reset_index(drop=True)
        self.zona = zona
        self.tarifa = tarifa
        inicio = pd.DatetimeIndex(self.curva["inicio"])
        self.kw = self.curva["kw"].to_numpy(float)
        self.periodo = asignar_periodos(inicio, zona, festivos.get("fijos"),
                                        festivos.get("adicionales", ()), festivos.get("excluidos", ()))
        mes = inicio.to_period("M")
        self.meses = sorted(set(mes))
        pos = {m: i for i, m in enumerate(self.meses)}
        self.idx_mes = np.array([pos[m] for m in mes])
        self.dias_mes = np.array([len(set(inicio[self.idx_mes == i].date)) for i in range(len(self.meses))])

        # valores ordenados y sumas acumuladas por (mes, periodo) para evaluar excesos rápido
        self._orden = {}
        for i in range(len(self.meses)):
            for p in range(N_P):
                v = np.sort(self.kw[(self.idx_mes == i) & (self.periodo == p + 1)])
                self._orden[i, p] = (v, np.concatenate([[0], np.cumsum(v)]),
                                     np.concatenate([[0], np.cumsum(v * v)]))

    # ------------------------------------------------------------------
    @property
    def etiquetas_meses(self):
        return [f"{MESES_ABR[m.month - 1]}-{m.year % 100:02d}" for m in self.meses]

    @property
    def dias_totales(self):
        return int(self.dias_mes.sum())

    def energia_kwh(self):
        e = np.zeros((len(self.meses), N_P))
        np.add.at(e, (self.idx_mes, self.periodo - 1), self.kw * 0.25)
        return e

    def maximos(self):
        m = np.zeros((len(self.meses), N_P))
        np.maximum.at(m, (self.idx_mes, self.periodo - 1), self.kw)
        return m

    def max_demanda(self):
        return float(self.kw.max()) if len(self.kw) else 0.0

    # ------------------------------------------------------------------
    def _suma_cuadrados(self, i, p, x):
        """Σ max(0, v - x)² para el mes i y periodo p; x puede ser un array."""
        v, s1, s2 = self._orden[i, p]
        x = np.asarray(x, float)
        k = np.searchsorted(v, x, side="right")
        n = len(v)
        cnt = n - k
        a = s1[n] - s1[k]
        b = s2[n] - s2[k]
        return np.maximum(b - 2 * x * a + cnt * x * x, 0.0)

    def coste(self, pc):
        pc = np.asarray(pc, float)
        fijo = np.outer(self.dias_mes / self.tarifa.dias_anio, pc * self.tarifa.tp)
        exceso = np.zeros_like(fijo)
        for i in range(len(self.meses)):
            for p in range(N_P):
                exceso[i, p] = self.tarifa.tep[p] * np.sqrt(self._suma_cuadrados(i, p, pc[p]))
        return Coste(fijo, exceso)

    def coste_periodo(self, p, x):
        """Coste anual (fijo + excesos) del periodo p para cada potencia de x (array)."""
        x = np.asarray(x, float)
        total = x * self.tarifa.tp[p] * self.dias_totales / self.tarifa.dias_anio
        for i in range(len(self.meses)):
            total = total + self.tarifa.tep[p] * np.sqrt(self._suma_cuadrados(i, p, x))
        return total

    def optimo(self, minimo=1):
        """Potencias enteras óptimas con P1 ≤ … ≤ P6."""
        hi = int(np.ceil(self.max_demanda())) + 1
        rejilla = np.arange(minimo, max(hi, minimo + 1) + 1)
        costes = [self.coste_periodo(p, rejilla) for p in range(N_P)]

        # pool adjacent violators para funciones convexas separables
        bloques = []   # [suma de costes, [periodos], índice del óptimo]
        for p in range(N_P):
            bloques.append([costes[p].copy(), [p], int(np.argmin(costes[p]))])
            while len(bloques) > 1 and bloques[-2][2] > bloques[-1][2]:
                c2, p2, _ = bloques.pop()
                c1, p1, _ = bloques.pop()
                c = c1 + c2
                bloques.append([c, p1 + p2, int(np.argmin(c))])
        pc = np.zeros(N_P)
        for c, periodos, k in bloques:
            pc[periodos] = rejilla[k]
        return pc.astype(int)


# --------------------------------------------------------------------------
# inversión (derechos de acometida y enganche)
# --------------------------------------------------------------------------

def opciones_inversion_defecto(pc, pc_actual):
    """Por defecto: si se amplía por encima de la máxima potencia actual, todos los derechos."""
    amplia = max(pc) > max(pc_actual) + 1e-9
    return {"acometida": amplia, "acceso": amplia, "extension": amplia, "enganche": True}


def calcular_inversion(pc, pc_actual, opciones, derechos):
    incremento = max(0.0, float(max(pc)) - float(max(pc_actual)))
    acometida = opciones.get("acometida", False)
    acceso = incremento * derechos["acceso"] if acometida and opciones.get("acceso") else 0.0
    extension = incremento * derechos["extension"] if acometida and opciones.get("extension") else 0.0
    enganche = derechos["enganche"] if opciones.get("enganche") else 0.0
    conceptos = [n for n, v in (("acceso", acceso), ("extensión", extension), ("enganche", enganche)) if v > 0]
    return {"incremento": incremento, "acceso": acceso, "extension": extension,
            "enganche": enganche, "total": acceso + extension + enganche, "conceptos": conceptos}


def texto_conceptos(conceptos):
    if not conceptos:
        return "sin inversión"
    if len(conceptos) == 1:
        return f"derechos de {conceptos[0]}"
    return "derechos de " + ", ".join(conceptos[:-1]) + " y " + conceptos[-1]


def validar_potencias(pc):
    """Devuelve un mensaje de error o None."""
    pc = list(pc)
    if any(p is None or not np.isfinite(p) or p <= 0 for p in pc):
        return "Todas las potencias deben ser números mayores que 0."
    for i in range(N_P - 1):
        if pc[i + 1] < pc[i]:
            return (f"P{i + 2} ({pc[i + 1]:g} kW) es menor que P{i + 1} ({pc[i]:g} kW). La Circular 3/2020 "
                    "exige potencias crecientes: P1 ≤ P2 ≤ P3 ≤ P4 ≤ P5 ≤ P6.")
    return None


# --------------------------------------------------------------------------
# escenarios
# --------------------------------------------------------------------------

@dataclass
class Escenario:
    nombre: str
    pc: np.ndarray
    coste: Coste
    inversion: dict = field(default_factory=dict)
    ahorro: float = 0.0
    ahorro_pct: float = 0.0

    @property
    def prs(self):
        if self.ahorro > 0 and self.inversion.get("total", 0) > 0:
            return self.inversion["total"] / self.ahorro
        return None


def evaluar(estudio, propuestas, pc_actual):
    """
    propuestas: lista de (nombre, pc, opciones_inversion).
    Devuelve [Escenario actual, escenarios de propuestas...].
    """
    actual = Escenario("Actual", np.asarray(pc_actual, float), estudio.coste(pc_actual))
    res = [actual]
    for nombre, pc, opciones in propuestas:
        c = estudio.coste(pc)
        inv = calcular_inversion(pc, pc_actual, opciones, estudio.tarifa.derechos)
        ahorro = actual.coste.total - c.total
        pct = ahorro / actual.coste.total if actual.coste.total else 0.0
        res.append(Escenario(nombre, np.asarray(pc, float), c, inv, ahorro, pct))
    return res


class ResumenMultipunto:
    """
    Suma de varios suministros (estudio multipunto), mes a mes: situación actual,
    propuesta óptima (Propuesta 1) y resto de propuestas de cada uno.
    La Propuesta k conjunta suma la Propuesta k de cada CUPS; los CUPS que no la tienen
    entran con su situación actual (se listan en `sin_propuesta`).
    Imita la interfaz de Estudio que usan las tablas y gráficos (meses, etiquetas_meses,
    dias_mes, dias_totales, energia_kwh) y expone `escenarios` = [Actual, Propuesta 1, ...].
    """

    def __init__(self, lista):
        """lista: [(estudio, escenarios), ...] de cada suministro ya calculado."""
        self.meses = sorted(set().union(*(set(est.meses) for est, _ in lista)))
        pos = {m: i for i, m in enumerate(self.meses)}
        n = len(self.meses)
        n_esc = max(len(escs) for _, escs in lista)
        self.dias_mes = np.zeros(n, dtype=int)
        self._energia = np.zeros((n, N_P))
        fijo = np.zeros((n_esc, n, N_P))
        exceso = np.zeros((n_esc, n, N_P))
        inversion = np.zeros(n_esc)
        self.sin_propuesta = {}          # nº de propuesta -> índices de los CUPS que no la tienen
        for j, (est, escs) in enumerate(lista):
            idx = [pos[m] for m in est.meses]
            self.dias_mes[idx] = np.maximum(self.dias_mes[idx], est.dias_mes)
            self._energia[idx] += est.energia_kwh()
            for k in range(n_esc):
                e = escs[k] if k < len(escs) else escs[0]
                if k >= len(escs):
                    self.sin_propuesta.setdefault(k, []).append(j)
                fijo[k, idx] += e.coste.fijo
                exceso[k, idx] += e.coste.exceso
                inversion[k] += e.inversion.get("total", 0.0)
        actual = Escenario("Actual", np.full(N_P, np.nan), Coste(fijo[0], exceso[0]))
        self.escenarios = [actual]
        for k in range(1, n_esc):
            c = Coste(fijo[k], exceso[k])
            ahorro = actual.coste.total - c.total
            pct = ahorro / actual.coste.total if actual.coste.total else 0.0
            self.escenarios.append(Escenario(f"Propuesta {k}", np.full(N_P, np.nan), c,
                                             {"total": float(inversion[k]), "conceptos": []}, ahorro, pct))
        self.periodos_distintos = len({tuple(est.meses) for est, _ in lista}) > 1

    @property
    def etiquetas_meses(self):
        return [f"{MESES_ABR[m.month - 1]}-{m.year % 100:02d}" for m in self.meses]

    @property
    def dias_totales(self):
        return int(self.dias_mes.sum())

    def energia_kwh(self):
        return self._energia.copy()


# --------------------------------------------------------------------------
# formato numérico español
# --------------------------------------------------------------------------

def fmt(x, dec=0):
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return "-"
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pot(x):
    return fmt(x, 0) if float(x).is_integer() else fmt(x, 2)
