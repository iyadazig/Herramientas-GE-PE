# -*- coding: utf-8 -*-
"""
DATOS REGULADOS DEL SISTEMA ELECTRICO (por año natural)
=======================================================
Peajes y cargos de todas las tarifas de acceso, pagos por capacidad, pérdidas por tarifa y
sistema, SRAD, impuestos (IEE e impuesto municipal), FNEE, pagos OS/OM, temporadas por sistema
eléctrico y los parámetros de estimación del mercado (% OMIP, años de la media de OMIE).

    regulados/regulados.json    semilla del repositorio (valores publicados en el BOE)
    almacen.datos_regulados     lo guardado desde Administración → Datos regulados: prevalece
                                sobre la semilla, así editar desde la app no cambia el repositorio

Si un año no está (años futuros sin publicar), se usa una copia del último año anterior que sí
está, marcada con "copia_de" para avisar en los controles del análisis.

Lo usan el análisis de ofertas y el estudio de potencia (precios_potencia(): misma forma que
potencia/precios.json). Sin Streamlit.
"""

import copy
import datetime as dt
import json
from pathlib import Path

SEMILLA = Path(__file__).resolve().parent / "regulados" / "regulados.json"

TARIFAS = ["2.0TD", "3.0TD", "3.0TDVE", "6.1TD", "6.1TDVE", "6.2TD", "6.3TD", "6.4TD"]
SISTEMAS = ["Península", "Baleares", "Canarias", "Ceuta", "Melilla"]
TEMPORADAS = ["Alta", "Media alta", "Media", "Baja"]
# conceptos de cada tarifa: (clave, nombre, unidad)
CONCEPTOS = [("peajes_potencia", "Peajes de potencia", "€/kW·año"),
             ("cargos_potencia", "Cargos de potencia", "€/kW·año"),
             ("peajes_energia", "Peajes de energía", "€/kWh"),
             ("cargos_energia", "Cargos de energía", "€/kWh"),
             ("pagos_capacidad", "Pagos por capacidad", "€/kWh"),
             ("perdidas", "Coeficientes de pérdidas", "tanto por uno"),
             ("srad", "SRAD (estimación)", "€/kWh")]
GENERALES = [("iee", "Impuesto eléctrico (IEE)", "tanto por uno"),
             ("iee_minimo_general", "IEE: mínimo general", "€/MWh"),
             ("iee_minimo_reducido", "IEE: mínimo usos con reducción", "€/MWh"),
             ("impuesto_municipal", "Impuesto municipal (K estándar)", "tanto por uno"),
             ("fnee", "FNEE", "€/MWh"),
             ("om", "Pago al Operador del Mercado (OM)", "€/MWh"),
             ("os", "Pago al Operador del Sistema (OS)", "€/MWh")]
MERCADO = [("pct_omip", "Ponderación de OMIP en el precio previsto", "tanto por uno"),
           ("anios_media_omie", "Años de la media mensual de OMIE", "años")]

_cache = {}


class DatosRegulados(Exception):
    """Datos regulados no válidos o que faltan."""


def _leer_semilla() -> dict:
    clave = SEMILLA.stat().st_mtime_ns
    if _cache.get("clave") != clave:
        _cache.update(clave=clave, semilla=json.loads(SEMILLA.read_text(encoding="utf-8")))
    return _cache["semilla"]


def semilla() -> dict:
    return copy.deepcopy(_leer_semilla())


def _guardados() -> dict:
    """{año: datos} guardados en la base de datos (vacío si no hay base de datos)."""
    try:
        import almacen
        if not almacen.ruta_bd().exists():      # no se crea la base de datos solo para leer
            return {}
        return almacen.datos_regulados_guardados()
    except Exception:          # base de datos no accesible: solo la semilla
        return {}


def anios(guardados: dict = None) -> list:
    """Años con datos propios (semilla o guardados), ordenados."""
    g = _guardados() if guardados is None else guardados
    return sorted({int(a) for a in _leer_semilla()["anios"]} | {int(a) for a in g})


def datos_anio(anio: int, guardados: dict = None) -> dict:
    """Datos regulados del año (copia: se puede modificar). Añade 'anio', 'origen'
    ('guardado', 'semilla' o 'copia') y, si es copia, 'copia_de' con el año copiado."""
    anio = int(anio)
    g = _guardados() if guardados is None else {int(a): d for a, d in guardados.items()}
    s = _leer_semilla()["anios"]
    disponibles = anios(g)
    if not disponibles:
        raise DatosRegulados("No hay datos regulados de ningún año.")
    previos = [a for a in disponibles if a <= anio]
    fuente = previos[-1] if previos else disponibles[0]
    datos = copy.deepcopy(g[fuente] if fuente in g else s[str(fuente)])
    datos["anio"] = anio
    datos["origen"] = ("copia" if fuente != anio else
                       "guardado" if fuente in g else "semilla")
    if fuente != anio:
        datos["copia_de"] = fuente
    else:
        datos.pop("copia_de", None)
    return datos


def comunes() -> dict:
    """Lo que no depende del año: sistemas, temporadas, periodos por temporada y festivos."""
    s = _leer_semilla()
    return {k: copy.deepcopy(s[k]) for k in ("sistemas", "temporadas", "periodos_temporada",
                                             "festivos")}


# ------------------------------------------------------------------- consultas
def n_periodos(tarifa: str) -> int:
    return 3 if tarifa == "2.0TD" else 6


def temporada(sistema: str, mes: int) -> str:
    for t, meses in _leer_semilla()["temporadas"][sistema].items():
        if not t.startswith("_") and mes in meses:
            return t
    raise DatosRegulados(f"El mes {mes} no tiene temporada en {sistema}.")


def periodos_validos(sistema: str, tarifa: str, mes: int) -> list:
    """Periodos en los que puede haber consumo ese mes (3P: los tres; 6P: los dos de la
    temporada del mes en el sistema y el P6)."""
    if n_periodos(tarifa) == 3:
        return [1, 2, 3]
    return list(_leer_semilla()["periodos_temporada"][sistema][temporada(sistema, mes)]) + [6]


def terminos(datos: dict, tarifa: str) -> dict:
    """Términos de una tarifa en un año: potencia (peaje + cargo, €/kW·año, 6 valores con 0
    en los periodos que no tiene), energía (peaje + cargo, €/kWh), pagos por capacidad,
    pérdidas y SRAD (6 valores; None si no hay)."""
    t = datos["tarifas"][tarifa]

    def seis(lista):
        lista = list(lista or [])
        return [float(v) if v is not None else 0.0 for v in lista] + [0.0] * (6 - len(lista))
    pot = [round(a + b, 6) for a, b in zip(seis(t["peajes_potencia"]), seis(t["cargos_potencia"]))]
    ene = [round(a + b, 6) for a, b in zip(seis(t["peajes_energia"]), seis(t["cargos_energia"]))]
    return {"potencia": pot, "energia": ene, "pagos_capacidad": seis(t["pagos_capacidad"]),
            "perdidas": seis(t["perdidas"]),
            "srad": seis(t["srad"]) if t.get("srad") else None}


def perdidas(datos: dict, tarifa: str, sistema: str) -> list:
    """Coeficientes de pérdidas (6 valores) de la tarifa en el sistema: los propios del sistema
    si se han puesto; si no, los de la tarifa (Circular 3/2020, art. 11, iguales en todo el
    territorio)."""
    propio = (datos.get("perdidas_sistema") or {}).get(sistema, {}).get(tarifa)
    base = propio if propio else datos["tarifas"][tarifa]["perdidas"]
    return [float(v) for v in base] + [0.0] * (6 - len(base))


def precios_potencia(anio: int = None, guardados: dict = None) -> dict:
    """Precios del estudio de potencia con la forma de potencia/precios.json (las tarifas que
    tienen término de excesos de potencia, tep)."""
    anio = anio or dt.date.today().year
    d = datos_anio(anio, guardados)
    return {"vigencia": str(d.get("copia_de", anio)), "dias_anio": d.get("dias_anio", 365),
            "normativa": list(d.get("normativa_potencia") or []),
            "tarifas": {t: {"_fuente": v.get("_fuente", ""), "nivel_tension": v["nivel_tension"],
                            "peajes_potencia": v["peajes_potencia"],
                            "cargos_potencia": v["cargos_potencia"], "tep": v["tep"]}
                        for t, v in d["tarifas"].items() if v.get("tep")},
            "derechos": d["derechos"], "festivos": comunes()["festivos"]}


# ------------------------------------------------------------------- validación y guardado
def _numeros(valor, n, nombre, errores, admite_vacio=False, maximo=None):
    if valor is None and admite_vacio:
        return
    if not isinstance(valor, list) or len(valor) != n:
        errores.append(f"{nombre}: tiene que tener {n} valores.")
        return
    for v in valor:
        if not isinstance(v, (int, float)) or isinstance(v, bool) or v != v or v < 0:
            errores.append(f"{nombre}: hay valores vacíos o negativos.")
            return
        if maximo is not None and v > maximo:
            errores.append(f"{nombre}: valor {v} fuera de rango (máximo {maximo}).")
            return


def validar(datos: dict) -> list:
    """Errores de los datos de un año (lista vacía si están bien)."""
    errores = []
    tarifas = datos.get("tarifas") or {}
    for t in TARIFAS:
        if t not in tarifas:
            errores.append(f"Falta la tarifa {t}.")
            continue
        v, n = tarifas[t], n_periodos(t)
        npot = 2 if t == "2.0TD" else 6
        _numeros(v.get("peajes_potencia"), npot, f"{t} peajes de potencia", errores)
        _numeros(v.get("cargos_potencia"), npot, f"{t} cargos de potencia", errores)
        _numeros(v.get("peajes_energia"), n, f"{t} peajes de energía", errores, maximo=1)
        _numeros(v.get("cargos_energia"), n, f"{t} cargos de energía", errores, maximo=1)
        _numeros(v.get("pagos_capacidad"), n, f"{t} pagos por capacidad", errores, maximo=1)
        _numeros(v.get("perdidas"), n, f"{t} pérdidas", errores, maximo=1)
        _numeros(v.get("srad"), n, f"{t} SRAD", errores, admite_vacio=True, maximo=1)
        if v.get("tep") is not None:
            _numeros(v.get("tep"), 6, f"{t} término de excesos de potencia", errores)
    for sistema, por_tarifa in (datos.get("perdidas_sistema") or {}).items():
        if sistema not in SISTEMAS:
            errores.append(f"Pérdidas por sistema: «{sistema}» no es un sistema eléctrico.")
            continue
        for t, valores in por_tarifa.items():
            if t in TARIFAS and valores:
                _numeros(valores, n_periodos(t), f"Pérdidas {t} en {sistema}", errores, maximo=1)
    g = datos.get("generales") or {}
    for clave, nombre, _ in GENERALES:
        v = g.get(clave)
        if not isinstance(v, (int, float)) or v < 0:
            errores.append(f"{nombre}: falta o es negativo.")
    for clave in ("iee", "impuesto_municipal"):
        if isinstance(g.get(clave), (int, float)) and g[clave] >= 1:
            errores.append("Los porcentajes van en tanto por uno (5,11 % = 0,0511).")
            break
    m = datos.get("mercado") or {}
    if not isinstance(m.get("pct_omip"), (int, float)) or not 0 <= m["pct_omip"] <= 1:
        errores.append("La ponderación de OMIP tiene que estar entre 0 y 1.")
    if not isinstance(m.get("anios_media_omie"), int) or not 1 <= m["anios_media_omie"] <= 10:
        errores.append("Los años de la media de OMIE tienen que ser un número entero de 1 a 10.")
    return errores


def validar_comunes(c: dict) -> list:
    """Errores de las temporadas y los periodos por temporada."""
    errores = []
    for s in SISTEMAS:
        meses = sorted(m for t, ms in c["temporadas"].get(s, {}).items() if not t.startswith("_")
                       for m in ms)
        if meses != list(range(1, 13)):
            errores.append(f"Temporadas de {s}: cada mes tiene que estar en una sola temporada.")
        for t in TEMPORADAS:
            p = c["periodos_temporada"].get(s, {}).get(t)
            if not p or len(p) != 2 or not all(v in range(1, 6) for v in p):
                errores.append(f"Periodos de la temporada {t} en {s}: dos periodos entre P1 y P5.")
    return errores


def limpiar(datos: dict) -> dict:
    """Quita las marcas que añade datos_anio antes de guardar."""
    d = copy.deepcopy(datos)
    for k in ("anio", "origen", "copia_de"):
        d.pop(k, None)
    return d


def guardar(anio: int, datos: dict, usuario: str) -> None:
    errores = validar(datos)
    if errores:
        raise DatosRegulados("\n".join(errores))
    import almacen
    almacen.guardar_datos_regulados(int(anio), limpiar(datos), usuario)


def exportar_json(guardados: dict = None) -> str:
    """La semilla con lo guardado en la base de datos, para actualizar el repositorio o el
    programa de escritorio."""
    s = semilla()
    g = _guardados() if guardados is None else guardados
    for a, d in g.items():
        s["anios"][str(a)] = limpiar(d)
    s["anios"] = dict(sorted(s["anios"].items()))
    return json.dumps(s, ensure_ascii=False, indent=1)
