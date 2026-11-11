# -*- coding: utf-8 -*-
"""
NUMEROS EN FORMATO ESPAÑOL (sin Streamlit)
==========================================
Criterio de todo el programa: los decimales se ven con coma y los miles con punto
(1.234,5678), y lo que se escribe o se pega se entiende con coma o con punto decimal.

    leer(v)              número de un texto escrito o pegado: «16,56», «1.234,5», «16.56»,
                         «1.234.567», «12 €», «5 %» (None si está vacío o no es un número)
    texto(v, dec, miles) 1234.5678 → «1.234,5678» (sin ceros de más; dec = decimales máximos)
    con_coma(t)          lo escrito con punto decimal, con coma: «0.394» → «0,394»
    tabla(df)            copia de un DataFrame con las columnas de decimales como texto español
                         (para mostrarlo: las tablas de Streamlit los ponen con punto)

Un punto solo se toma como separador de miles si hay más de uno («1.234.567») o si el texto
también lleva coma («1.234,5»); un único punto es decimal («0.394», «1.234» = 1,234).
"""

import math
import re

DECIMALES = 10                  # máximos al mostrar (sin ceros de más)


def leer(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if isinstance(v, float) and math.isnan(v) else float(v)
    try:
        import pandas as pd
        if v is pd.NA:
            return None
    except ImportError:                                     # pragma: no cover
        pass
    t = str(v).strip().replace(" ", "").replace(" ", "").replace("€", "").replace("%", "")
    if t in ("", "None", "nan", "NaN", "—", "-"):
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


def texto(v, decimales: int = DECIMALES, miles: bool = True) -> str:
    """Número en formato español, sin ceros de más; '' si está vacío."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v)
    t = f"{n:,.{decimales}f}" if miles else f"{n:.{decimales}f}"
    if "." in t:
        t = t.rstrip("0").rstrip(".")
    if t in ("-0", ""):
        t = "0"
    return t.replace(",", "\0").replace(".", ",").replace("\0", ".")


def con_coma(t):
    """El número tal y como se escribió pero con coma decimal: «0.394» → «0,394»,
    «-16.5» → «-16,5». Lo que ya lleva coma, los miles con varios puntos («1.234.567») y lo que
    no es un número se dejan igual."""
    if not isinstance(t, str):
        return t
    s = t.strip()
    if re.fullmatch(r"-?\d*\.\d+", s):
        return s.replace(".", ",")
    return t


def _es_decimal(serie) -> bool:
    """La columna tiene números con decimales (float, o mezcla de textos y números)."""
    import pandas as pd
    if pd.api.types.is_bool_dtype(serie):
        return False
    if pd.api.types.is_float_dtype(serie):
        return True
    if serie.dtype == object:
        valores = [v for v in serie if v is not None and not (isinstance(v, float) and math.isnan(v))]
        return bool(valores) and any(isinstance(v, float) for v in valores) and all(
            isinstance(v, (int, float)) and not isinstance(v, bool) for v in valores)
    return False


def tabla(df, decimales: int = DECIMALES, sin_miles=("Año", "Mes", "Periodo", "Nº", "ID")):
    """Copia del DataFrame para mostrarlo: las columnas con decimales pasan a texto español
    (1.234,5678); las de años, meses, números de orden e ids no llevan punto de miles."""
    salida = df.copy()
    for c in salida.columns:
        if _es_decimal(salida[c]):
            miles = not any(str(c).startswith(x) for x in sin_miles)
            salida[c] = [texto(v, decimales, miles) for v in salida[c]]
    return salida
