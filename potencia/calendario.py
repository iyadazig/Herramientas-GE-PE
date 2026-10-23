"""
Calendario de periodos horarios de 6 periodos (peajes 3.0TD y 6.xTD).

Transcripción literal del art. 7 de la Circular 3/2020 de la CNMC:
temporadas por sistema eléctrico, tipos de día (A, B, B1, C, D) y horas de
cada periodo. Ojo: en Canarias y Ceuta la asignación de periodos por tipo de
día NO sigue el mismo patrón que en la Península.
"""

from datetime import date

import numpy as np
import pandas as pd

ZONAS = ["Península", "Illes Balears", "Canarias", "Ceuta", "Melilla"]

# Temporada de cada mes (1..12): A = alta, B = media alta, B1 = media, C = baja
_TEMPORADAS = {
    "Península":     {"A": [1, 2, 7, 12], "B": [3, 11], "B1": [6, 8, 9], "C": [4, 5, 10]},
    "Canarias":      {"A": [7, 8, 9, 10], "B": [11, 12], "B1": [1, 2, 3], "C": [4, 5, 6]},
    "Illes Balears": {"A": [6, 7, 8, 9], "B": [5, 10], "B1": [1, 2, 12], "C": [3, 4, 11]},
    "Ceuta":         {"A": [1, 2, 8, 9], "B": [7, 10], "B1": [3, 11, 12], "C": [4, 5, 6]},
    "Melilla":       {"A": [1, 7, 8, 9], "B": [2, 12], "B1": [6, 10, 11], "C": [3, 4, 5]},
}

# Horas (hora de inicio 0..23) de los dos bloques de cada día laborable:
# "punta" = bloque más caro del día, "llano" = bloque intermedio. El resto
# de 0 h a 8 h es siempre P6.
_HORAS = {
    "Península":     {"punta": [9, 10, 11, 12, 13, 18, 19, 20, 21],
                      "llano": [8, 14, 15, 16, 17, 22, 23]},
    "Illes Balears": {"punta": [10, 11, 12, 13, 14, 18, 19, 20, 21],
                      "llano": [8, 9, 15, 16, 17, 22, 23]},
    "Canarias":      {"punta": [10, 11, 12, 13, 14, 18, 19, 20, 21],
                      "llano": [8, 9, 15, 16, 17, 22, 23]},
    "Ceuta":         {"punta": [10, 11, 12, 13, 14, 19, 20, 21, 22],
                      "llano": [8, 9, 15, 16, 17, 18, 23]},
    "Melilla":       {"punta": [10, 11, 12, 13, 14, 19, 20, 21, 22],
                      "llano": [8, 9, 15, 16, 17, 18, 23]},
}

# Periodo aplicable a (bloque punta, bloque llano) según el tipo de día
_PERIODOS_TIPO_DIA = {
    "Península":     {"A": (1, 2), "B": (2, 3), "B1": (3, 4), "C": (4, 5)},
    "Illes Balears": {"A": (1, 2), "B": (2, 3), "B1": (3, 4), "C": (4, 5)},
    "Canarias":      {"A": (1, 3), "B": (2, 3), "B1": (2, 4), "C": (4, 5)},
    "Ceuta":         {"A": (1, 4), "B": (2, 3), "B1": (2, 4), "C": (3, 5)},
    "Melilla":       {"A": (1, 2), "B": (2, 3), "B1": (3, 4), "C": (4, 5)},
}

FESTIVOS_FIJOS_DEFECTO = ["01-01", "01-06", "05-01", "08-15", "10-12",
                          "11-01", "12-06", "12-08", "12-25"]


def zona_horaria(zona):
    """Zona horaria IANA del sistema eléctrico (para curvas con desfase UTC)."""
    return "Atlantic/Canary" if zona == "Canarias" else "Europe/Madrid"


def hora_cambio_horario(zona):
    """Hora local (inicio) de la hora que falta en marzo / se repite en octubre."""
    return 1 if zona == "Canarias" else 2


def ultimo_domingo(anio, mes):
    d = pd.Timestamp(anio, mes, 1) + pd.offsets.MonthEnd(0)
    return (d - pd.Timedelta(days=(d.weekday() + 1) % 7)).date()


def dias_cambio_hora(anio):
    """(día de 23 h en marzo, día de 25 h en octubre)."""
    return ultimo_domingo(anio, 3), ultimo_domingo(anio, 10)


def festivos(anios, fijos=None, adicionales=(), excluidos=()):
    """Conjunto de fechas festivas (tipo D) para los años indicados."""
    fijos = FESTIVOS_FIJOS_DEFECTO if fijos is None else fijos
    excluidos = {pd.Timestamp(x).date() for x in excluidos}
    res = set()
    for anio in anios:
        for md in fijos:
            mes, dia = (int(x) for x in md.split("-"))
            res.add(date(anio, mes, dia))
    res |= {pd.Timestamp(x).date() for x in adicionales}
    return res - excluidos


def asignar_periodos(inicio, zona, fijos=None, adicionales=(), excluidos=()):
    """
    Periodo (1..6) de cada intervalo a partir de su hora local de inicio.

    `inicio` es una serie/array de datetime sin zona horaria (hora oficial local).
    """
    if zona not in _TEMPORADAS:
        raise ValueError(f"Zona desconocida: {zona}")
    ts = pd.DatetimeIndex(inicio)
    mes = ts.month.to_numpy()
    hora = ts.hour.to_numpy()
    dias = ts.normalize()

    anios = sorted(set(ts.year))
    fest = festivos(anios, fijos, adicionales, excluidos)
    es_festivo = np.isin(dias.date, list(fest)) if fest else np.zeros(len(ts), bool)
    tipo_d = (ts.weekday.to_numpy() >= 5) | es_festivo

    # periodo del bloque punta y del bloque llano según el mes
    per_punta = np.zeros(13, int)
    per_llano = np.zeros(13, int)
    for tipo, meses in _TEMPORADAS[zona].items():
        p_punta, p_llano = _PERIODOS_TIPO_DIA[zona][tipo]
        per_punta[meses] = p_punta
        per_llano[meses] = p_llano

    es_punta = np.isin(hora, _HORAS[zona]["punta"])
    es_llano = np.isin(hora, _HORAS[zona]["llano"])

    periodo = np.full(len(ts), 6, dtype=int)
    laborable = ~tipo_d
    periodo[laborable & es_punta] = per_punta[mes[laborable & es_punta]]
    periodo[laborable & es_llano] = per_llano[mes[laborable & es_llano]]
    return periodo
