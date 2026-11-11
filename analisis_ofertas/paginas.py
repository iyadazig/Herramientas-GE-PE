# -*- coding: utf-8 -*-
"""
PAGINAS DEL «ANALISIS DE OFERTAS» (sección Ofertas de Herramientas GE&PE)
========================================================================
    pagina_nuevo_analisis       1. origen (solicitud guardada o plantilla de suministros)
                                2. suministros · 3. mercado, SSAA, apuntamientos y curvas
                                4. contrato actual · 5. ofertas · 6. controles · 7. resultados
    pagina_analisis_guardados   lista, detalle, descargas y «Cargar»
    pagina_parametros           «Parámetros utilizados» (los vigentes) con exportación a Excel

La conversión entre las tablas editables y el modelo (tabla_* / aplicar_*) no usa Streamlit y
se prueba en tests/test_analisis_paginas.py. Claves de sesión: prefijo «ao_».
"""

import datetime as dt
import math
import re

import pandas as pd
import streamlit as st

import almacen
import datos_mercado as dm
import datos_regulados as dr
import estilo
import numeros
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo, correo, informes, parametros, resultados, servicio
from analisis_ofertas import comercializadoras as coms
from analisis_ofertas import ayudas, catalogo
from analisis_ofertas import modelo as mo

MODULO = "Análisis de ofertas"
AY = ayudas.AYUDA
REINTENTO_APUNTAMIENTOS = 3600          # s entre intentos de calcular un mes aún sin publicar
_ultimo_intento = [0.0]
PAGINA_FORMULARIO = None            # la fija inicio.py («Nuevo análisis»)
PER = mo.PERIODOS
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MESES = ap.MESES


# =================================================================== tablas <-> modelo
def _num(v, defecto=None):
    if v is None or (isinstance(v, float) and math.isnan(v)) or v is pd.NA:
        return defecto
    try:
        return float(v)
    except (TypeError, ValueError):
        return defecto


def numero_es(v):
    """Número escrito o pegado al estilo español: «16,56», «1.234,5», «16.56» (None si vacío o
    no es un número)."""
    if v is None or v is pd.NA:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if isinstance(v, float) and math.isnan(v) else float(v)
    t = str(v).strip().replace("\u00a0", "").replace(" ", "").replace("€", "").replace("%", "")
    if t in ("", "None", "nan", "—", "-"):
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


def con_coma(texto):
    """El número tal y como se escribió pero con coma decimal: «0.394» → «0,394»,
    «-16.5» → «-16,5». Lo que ya lleva coma, los miles con varios puntos («1.234.567») y lo
    que no es un número se dejan igual."""
    if not isinstance(texto, str):
        return texto
    t = texto.strip()
    if re.fullmatch(r"-?\d*\.\d+", t):
        return t.replace(".", ",")
    return texto


def valores_pegados(texto: str) -> list:
    """Números pegados de Excel, de una fila (separados por tabuladores) o de una columna (uno
    por línea); si no hay ninguno de los dos, separados por «;» o espacios. Una celda vacía
    en medio cuenta (None) para no correr los meses; los vacíos del final no."""
    t = (texto or "").strip("\r\n ")
    if not t:
        return []
    if "\t" in t or "\n" in t:
        partes = re.split(r"\t|\r?\n", t)
    else:
        partes = re.split(r"[;\s]+", t)
    while partes and not partes[-1].strip():
        partes.pop()
    return [numero_es(p) for p in partes]


def pegar_precios(tabla: pd.DataFrame, texto: str, desde: int = 0) -> tuple:
    """Pone en la tabla de precios (filas Grupo/Año, columnas P1-P6) lo pegado de Excel a
    partir de la fila `desde`. Devuelve (tabla nueva, error o None).
    - Una fila de Excel (celdas separadas por tabuladores) = P1, P2… de una fila; con menos de
      6 celdas, los periodos que faltan no se tocan.
    - Varias filas = varias filas seguidas de la tabla.
    - Una columna (una celda por línea, hasta 6) = P1, P2… de la fila elegida.
    Una celda vacía en medio deja ese periodo vacío."""
    t = (texto or "").strip("\r\n")
    if not t.strip():
        return tabla, "No hay nada pegado."
    lineas = [l for l in re.split(r"\r?\n", t)]
    while lineas and not lineas[-1].strip():
        lineas.pop()
    if any("\t" in l for l in lineas):
        filas = [l.split("\t") for l in lineas]
    elif len(lineas) > 1:
        if len(lineas) > len(PER):
            return tabla, (f"Se han pegado {len(lineas)} valores en columna y hay como mucho "
                           f"{len(PER)} periodos.")
        filas = [[l for l in lineas]]                       # columna = periodos de una fila
    else:
        filas = [re.split(r"[;\s]+", lineas[0].strip())]
    nueva = tabla.copy()
    if desde + len(filas) > len(nueva):
        return tabla, (f"Se han pegado {len(filas)} filas y desde la fila elegida solo quedan "
                       f"{len(nueva) - desde}.")
    for i, celdas in enumerate(filas):
        while celdas and not celdas[-1].strip():
            celdas = celdas[:-1]
        if len(celdas) > len(PER):
            return tabla, (f"La fila {i + 1} tiene {len(celdas)} valores y hay como mucho "
                           f"{len(PER)} periodos (P1-P6).")
        for p, celda in zip(PER, celdas):
            v = numero_es(celda)
            if v is None and celda.strip():
                return tabla, f"«{celda.strip()}» no es un número."
            nueva.iloc[desde + i, nueva.columns.get_loc(p)] = v
    return nueva, None


def texto_es(v, decimales: int = 6) -> str:
    """16.56 → «16,56» (sin ceros de más)."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return f"{float(v):.{decimales}f}".rstrip("0").rstrip(".").replace(".", ",")


MESES_EN = {"Jan": "ene", "Feb": "feb", "Mar": "mar", "Apr": "abr", "May": "may", "Jun": "jun",
            "Jul": "jul", "Aug": "ago", "Sep": "sep", "Oct": "oct", "Nov": "nov", "Dec": "dic"}


def contrato_es(texto) -> str:
    """«M Jan-27» (análisis guardados antes) → «M ene-27»."""
    texto = texto or ""
    for en, es in MESES_EN.items():
        texto = texto.replace(f" {en}-", f" {es}-")
    return texto


def tabla_suministros(e: mo.Estudio) -> pd.DataFrame:
    return pd.DataFrame([{
        "Incluir": s.activo, "CUPS": s.cups, "Alias": s.alias, "Tarifa": s.tarifa,
        "Sistema": s.sistema, "Reducción IEE": s.reduccion_iee, "Alquiler (€/mes)": s.alquiler,
        "Consumo anual (kWh)": round(s.consumo_anual()),
        **{f"Pot. {p} (kW)": s.potencias[i] for i, p in enumerate(PER)},
        "Apuntamiento": s.apuntamiento if e.ap_cliente.get(s.cups) else mo.ESTIMADO}
        for s in e.suministros])


def aplicar_suministros(e: mo.Estudio, df: pd.DataFrame) -> None:
    for _, f in df.iterrows():
        s = e.suministro(f["CUPS"])
        if not s:
            continue
        s.activo = bool(f["Incluir"])
        s.alias = str(f["Alias"] or "")
        s.tarifa = f["Tarifa"] or s.tarifa
        s.sistema = f["Sistema"] or s.sistema
        s.reduccion_iee = _num(f["Reducción IEE"], 0.0)
        s.alquiler = _num(f["Alquiler (€/mes)"], 0.0)
        s.potencias = [_num(f[f"Pot. {p} (kW)"], 0.0) for p in PER]
        s.apuntamiento = (f["Apuntamiento"] if e.ap_cliente.get(s.cups) else mo.ESTIMADO)


def tabla_mercado(e: mo.Estudio) -> pd.DataFrame:
    ctx = calculo.Contexto(e)
    filas = []
    for k in range(e.horizonte()):
        f = mo.sumar_meses(e.inicio, k)
        clave = f"{f.year}-{f.month:02d}"
        fila = e.mercado.get(clave) or {}
        previsto, omip = ctx.mercado(f.year, f.month)
        filas.append({"Clave": clave, "Mes": f"{MESES[f.month - 1][:3]}-{f.year % 100:02d}",
                      "OMIP (€/MWh)": omip or None,
                      "Contrato": "manual" if clave in e.omip_manual else contrato_es(
                          fila.get("contrato")),
                      "Media OMIE (€/MWh)": fila.get("omie_media"),
                      "Previsto (€/MWh)": previsto})
    return pd.DataFrame(filas)


def aplicar_mercado(e: mo.Estudio, df: pd.DataFrame) -> None:
    """Un OMIP distinto del de la sesión queda como ajuste a mano del mes."""
    for _, f in df.iterrows():
        clave = f["Clave"]
        original = (e.mercado.get(clave) or {}).get("omip")
        nuevo = _num(f["OMIP (€/MWh)"])
        if nuevo is not None and (original is None or abs(nuevo - original) > 1e-9):
            e.omip_manual[clave] = nuevo
        elif nuevo is None or (original is not None and abs(nuevo - original) <= 1e-9):
            e.omip_manual.pop(clave, None)
        omie = _num(f["Media OMIE (€/MWh)"])
        if omie is not None:
            e.mercado.setdefault(clave, {})["omie_media"] = omie


def tabla_ssaa(e: mo.Estudio) -> pd.DataFrame:
    return pd.DataFrame([{"Mes": MESES[m], "SSAA (€/MWh)": e.ssaa_mes[m]} for m in range(12)])


def aplicar_ssaa(e: mo.Estudio, df: pd.DataFrame) -> None:
    e.ssaa_mes = [_num(v) for v in df["SSAA (€/MWh)"]]


def tabla_ap(e: mo.Estudio, clave: str) -> pd.DataFrame:
    n = 3 if clave.endswith("3P") else 6
    datos = e.ap_omie.get(clave) or {}
    return pd.DataFrame([{"Mes": MESES[m - 1], **{p: mo.seis(datos.get(m))[i]
                                                 for i, p in enumerate(PER[:n])},
                          "Dato de": (e.ap_origen.get(clave) or {}).get(m, "")}
                         for m in range(1, 13)])


def aplicar_ap(e: mo.Estudio, clave: str, df: pd.DataFrame) -> None:
    """Los meses cambiados quedan como «manual»."""
    n = 3 if clave.endswith("3P") else 6
    for m, (_, f) in enumerate(df.iterrows(), start=1):
        nuevos = [_num(f[p]) for p in PER[:n]] + [None] * (6 - n)
        antes = mo.seis((e.ap_omie.get(clave) or {}).get(m))
        if any((a is None) != (b is None) or (a is not None and abs(a - b) > 1e-9)
               for a, b in zip(nuevos, antes)):
            e.ap_omie.setdefault(clave, {})[m] = nuevos
            e.ap_origen.setdefault(clave, {})[m] = "manual"


CONCEPTOS_PRECIO = {"energia": "Energía", "termino_b": "Término B (€/MWh)",
                    "ci": "Coeficiente Ci", "potencia": "Potencia propia (€/kW·año)"}


def conceptos_precio(o: mo.Oferta) -> list:
    salida = ["energia"] if o.modalidad == mo.FIJO else (
        ["termino_b", "ci"] if o.modalidad == mo.PASS_POOL else ["termino_b"])
    if o.potencia == mo.POT_PROPIA:
        salida.append("potencia")
    return salida


def grupos_tabla(e: mo.Estudio, o: mo.Oferta) -> list:
    """Grupos de precios que salen en la tabla: los automáticos y los propios."""
    return list(calculo.grupos_de(e, o))


def tabla_precios(e: mo.Estudio, o: mo.Oferta, concepto: str) -> pd.DataFrame:
    anios = max(1, math.ceil(int(o.duracion or 12) / 12))
    filas = []
    for g in grupos_tabla(e, o):
        lista = o.precios.get(g) or []
        for y in range(1, anios + 1):
            valores = mo.seis(getattr(lista[y - 1], concepto)) if y <= len(lista) else [None] * 6
            filas.append({"Grupo": g, "Año": y, **dict(zip(PER, valores))})
    return pd.DataFrame(filas)


def aplicar_precios(o: mo.Oferta, concepto: str, df: pd.DataFrame) -> None:
    """Años sin ningún valor = los del año anterior; los años vacíos del final se quitan."""
    for g, filas in df.groupby("Grupo", sort=False):
        lista = list(o.precios.get(g) or [])
        anios = sorted(filas["Año"].astype(int))
        for _, f in filas.sort_values("Año").iterrows():
            y = int(f["Año"])
            while len(lista) < y:
                lista.append(mo.PreciosAnio())
            valores = [_num(f[p]) for p in PER]
            if concepto == "ci":
                valores = [1.0 if v is None else v for v in valores]
            setattr(lista[y - 1], concepto, valores)
        # quitar los años del final iguales a «vacíos» (heredan del anterior)
        while len(lista) > 1 and _vacio(lista[-1]):
            lista.pop()
        for y in range(1, len(lista)):
            if _vacio(lista[y]):
                lista[y] = mo.PreciosAnio(**{k: list(v) for k, v in vars(lista[y - 1]).items()})
        if len(lista) > max(anios, default=1):
            lista = lista[:max(anios)]
        o.precios[g] = lista


def _vacio(p: mo.PreciosAnio) -> bool:
    return (all(v is None for v in p.energia) and all(v is None for v in p.termino_b)
            and all(v is None for v in p.potencia) and all(v in (None, 1.0) for v in p.ci))


def tabla_grupos_propios(o: mo.Oferta) -> pd.DataFrame:
    return pd.DataFrame([{"Grupo": g.nombre, "Tarifa": g.tarifa, "Sistema": g.sistema}
                         for g in o.grupos_propios], columns=["Grupo", "Tarifa", "Sistema"])


def aplicar_grupos_propios(o: mo.Oferta, df: pd.DataFrame) -> None:
    o.grupos_propios = [mo.GrupoPropio(str(f["Grupo"]).strip(), f["Tarifa"],
                                       f["Sistema"] or "Península")
                        for _, f in df.iterrows()
                        if isinstance(f["Grupo"], str) and f["Grupo"].strip() and f["Tarifa"]]
    nombres = {g.nombre for g in o.grupos_propios}
    o.asignacion = {c: g for c, g in o.asignacion.items() if g in nombres}


def tabla_cups_oferta(e: mo.Estudio, o: mo.Oferta) -> pd.DataFrame:
    return pd.DataFrame([{"CUPS": s.cups, "Alias": s.alias, "Tarifa": s.tarifa,
                          "Sistema": s.sistema, "Incluido": s.cups not in o.excluidos,
                          "Grupo propio": o.asignacion.get(s.cups) or ""}
                         for s in e.suministros if s.activo])


def aplicar_cups_oferta(o: mo.Oferta, df: pd.DataFrame) -> None:
    o.excluidos = [f["CUPS"] for _, f in df.iterrows() if not bool(f["Incluido"])]
    o.asignacion = {f["CUPS"]: f["Grupo propio"] for _, f in df.iterrows()
                    if isinstance(f["Grupo propio"], str) and f["Grupo propio"]}


def tabla_cierres(e: mo.Estudio, o: mo.Oferta) -> pd.DataFrame:
    ctx = calculo.Contexto(e)
    inicio = o.inicio or e.inicio
    por_mes = {c.mes: c for c in o.cierres}
    filas = []
    for k in range(int(o.duracion or 12)):
        f = mo.sumar_meses(inicio, k)
        c = por_mes.get(k + 1) or mo.Cierre(k + 1)
        previsto, omip = ctx.mercado(f.year, f.month)
        filas.append({"Mes": k + 1, "Fecha": f"{MESES[f.month - 1][:3]}-{f.year % 100:02d}",
                      "Previsto (€/MWh)": previsto, "OMIP (€/MWh)": omip or None,
                      "% cerrado": c.pct, "MW": c.mw, "Precio de cierre (€/MWh)": c.precio,
                      "MW pico": c.mw_pico, "Precio pico (€/MWh)": c.precio_pico})
    return pd.DataFrame(filas)


def aplicar_cierres(o: mo.Oferta, df: pd.DataFrame) -> None:
    def v(f, c):
        return _num(f[c]) if c in f else None
    columnas = ("% cerrado", "MW", "Precio de cierre (€/MWh)", "MW pico", "Precio pico (€/MWh)")
    o.cierres = [mo.Cierre(int(f["Mes"]), pct=v(f, "% cerrado"), mw=v(f, "MW"),
                           precio=v(f, "Precio de cierre (€/MWh)"), mw_pico=v(f, "MW pico"),
                           precio_pico=v(f, "Precio pico (€/MWh)"))
                 for _, f in df.iterrows() if any(v(f, c) is not None for c in columnas)]


MOSTRAR_PH = False          # Ph demanda de Canarias: de momento no se usa (no se muestra)


def tabla_ph(e: mo.Estudio) -> pd.DataFrame:
    """Lo que el Ph demanda de Canarias lleva encima del mercado diario, por mes natural."""
    return pd.DataFrame([{"Mes": MESES[m - 1], "Ph − diario (€/MWh)": e.ph_dif.get(m),
                          "Dato de": e.ph_origen.get(m, "")} for m in range(1, 13)])


def aplicar_ph(e: mo.Estudio, df: pd.DataFrame) -> None:
    for m, (_, f) in enumerate(df.iterrows(), start=1):
        nuevo = _num(f["Ph − diario (€/MWh)"])
        antes = e.ph_dif.get(m)
        if nuevo is None:
            e.ph_dif.pop(m, None)
            e.ph_origen.pop(m, None)
        elif antes is None or abs(nuevo - antes) > 1e-9:
            e.ph_dif[m] = nuevo
            e.ph_origen[m] = "manual"


def claves_ap(e: mo.Estudio) -> list:
    """Sistemas y calendarios de los suministros (y el Ph de Canarias si hay CUPS allí)."""
    usados = {mo.clave_ap(s.sistema, mo.calendario_de(s.tarifa))
              for s in e.suministros if s.activo and s.tarifa in mo.TARIFAS}
    usados |= {mo.clave_ap(mo.SISTEMA_PH, mo.calendario_de(s.tarifa))
               for s in e.suministros if s.activo and s.sistema == "Canarias"
               and s.tarifa in mo.TARIFAS}
    return sorted(usados)


def referencia_desvios(e: mo.Estudio) -> str:
    """Texto de ayuda con los desvíos de referencia de los datos regulados."""
    reg = (e.regulados.get(e.inicio.year) if e.inicio and e.regulados else None) or \
        dr.datos_anio(e.inicio.year if e.inicio else dt.date.today().year)
    g = reg.get("generales", {})
    if g.get("desvios_coste_ref") is None:
        return ""
    neto = g["desvios_coste_ref"] + g.get("desvios_saldo_ref", 0.0)
    return (f"Referencia ({g.get('desvios_fuente', 'ESIOS')}): coste de desvíos "
            f"{texto_es(g['desvios_coste_ref'], 3)} €/MWh, saldo "
            f"{texto_es(g.get('desvios_saldo_ref', 0.0), 3)} €/MWh, neto {texto_es(neto, 3)} "
            "€/MWh. Usa el que cobre la oferta (p. ej. «coste de desvío del sistema» = coste; "
            "«saldo en valor absoluto» = coste + |saldo|).")


def tabla_comparativa(filas: list) -> pd.DataFrame:
    return pd.DataFrame([{
        "Ranking": f["ranking"], "Oferta": f["nombre"], "Comercializadora": f["comercializadora"],
        "Ronda": f["ronda"], "Modalidad": f["modalidad"], "SSAA": f["ssaa"],
        "Cierres": f["cierres"], "Meses": f["duracion"], "€/año": f["media_anual"],
        "€/kWh": f["eur_kwh"], "Actual mismo periodo (€/año)": f["actual_mismo_periodo"],
        "Ahorro (€/año)": f["ahorro"],
        "Variación": None if f["variacion"] is None else f["variacion"] * 100,
        "CUPS con su apuntamiento": f["apuntamiento_cliente"] or None,
        "Incidencias": f["incidencias"] or None} for f in filas])


# =================================================================== utilidades de página
def _usuario() -> str:
    return estilo._usuario()


def _estudio() -> mo.Estudio:
    return st.session_state.get("ao_estudio")


def _v() -> int:
    """Versión de las tablas (cambia al crear, cargar o reordenar ofertas)."""
    return st.session_state.get("ao_v", 0)


def _subir_version():
    st.session_state.ao_v = _v() + 1


def _limpiar():
    for k in [k for k in st.session_state if str(k).startswith("ao_")]:
        del st.session_state[k]


def _empezar(e: mo.Estudio, avisos: list, ref: str = "", id_analisis: int = None,
             curvas: dict = None):
    v = _v() + 1
    _limpiar()
    st.session_state.ao_estudio = e
    st.session_state.ao_avisos = avisos
    st.session_state.ao_ref = ref
    st.session_state.ao_id = id_analisis
    st.session_state.ao_curvas = curvas or {}
    st.session_state.ao_v = v


def cargar_analisis(id_analisis: int) -> bool:
    an = almacen.analisis(id_analisis)
    if not an:
        return False
    e = servicio.estudio_de_datos(an["datos"])
    _empezar(e, [], an["datos"].get("ref") or an.get("ref") or "", an["id"],
             servicio.desempaquetar_curvas(an.get("curvas")))
    st.session_state.ao_estado = an["estado"]
    return True


def _apuntamientos_al_dia() -> list:
    """Apuntamientos calculados; si falta el último mes, intenta calcularlo (como mucho una vez
    por hora en el servidor: mientras no se publique, no se repite en cada página)."""
    import time
    calc = almacen.listar_apuntamientos()
    if ap.pendientes(calc) and time.time() - _ultimo_intento[0] > REINTENTO_APUNTAMIENTOS:
        _ultimo_intento[0] = time.time()
        with st.spinner("Calculando los apuntamientos de los últimos meses…"):
            ap.actualizar()
        calc = almacen.listar_apuntamientos()
    return calc


def editor(df: pd.DataFrame, clave: str, numeros=(), decimales: dict = None, **kw):
    """st.data_editor con dos arreglos:
    - las columnas numéricas van como texto en formato español, así al pegar desde Excel
      «16,56» se lee 16,56 (el editor numérico de Streamlit tomaba la coma como separador de
      miles y daba 1656);
    - la tabla de partida se guarda la primera vez y no cambia mientras no cambie la clave: si
      se le pasara cada vez la tabla rehecha con lo editado, Streamlit pierde la edición y hay
      que pulsar Enter dos veces."""
    bases = st.session_state.setdefault("ao_bases", {})
    if clave not in bases:
        base = df.copy()
        for c in numeros:
            base[c] = [texto_es(v, (decimales or {}).get(c, 6)) for v in base[c]]
        bases[clave] = base
    config = dict(kw.pop("column_config", None) or {})
    for c in numeros:
        config[c] = st.column_config.TextColumn(c)
    versiones = st.session_state.setdefault("ao_bases_v", {})
    n = versiones.get(clave, 0)
    editada = st.data_editor(bases[clave], key=f"{clave}~{n}" if n else clave,
                             column_config=config, **kw).copy()
    # lo escrito con punto decimal se ve con coma: la tabla se rehace con lo editado y la coma
    # (el editor no deja cambiar lo escrito) y se vuelve a dibujar
    if any(con_coma(v) != v for c in numeros for v in editada[c] if isinstance(v, str)):
        base = editada.copy()
        for c in numeros:
            base[c] = [con_coma(v) for v in base[c]]
        bases[clave] = base
        versiones[clave] = n + 1
        st.rerun()
    malos = []
    for c in numeros:
        valores = list(editada[c])
        editada[c] = [numero_es(v) for v in valores]
        malos += [str(v) for v, n in zip(valores, editada[c])
                  if n is None and str(v or "").strip() not in ("", "None", "nan")]
    if malos:
        st.warning("No se entienden como números (se dejan vacíos): " + ", ".join(malos[:6]))
    return editada


def _fecha_hora(iso: str) -> str:
    return dt.datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M") if iso else ""


# =================================================================== 1. origen
def solicitudes_filtradas(sols: list, cliente: str = "", denominacion: str = "",
                          usuarios: list = None) -> list:
    """Solicitudes que contienen el texto del cliente (cliente, Ref. o CIF) y de la
    denominación del suministro (alias o CUPS), de los usuarios elegidos (vacío = todos)."""
    cliente, denominacion = cliente.strip().lower(), denominacion.strip().lower()
    return [s for s in sols
            if (not usuarios or s["usuario"] in usuarios)
            and (not cliente or cliente in " ".join(str(s.get(k) or "") for k in (
                "cliente", "ref")).lower())
            and (not denominacion or denominacion in " ".join(str(s.get(k) or "") for k in (
                "denominaciones", "cups")).lower())]


def filtrar_solicitudes(sols: list) -> list:
    """Filtros de la lista de solicitudes; por defecto, las del usuario conectado."""
    nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
    yo = _usuario()
    c1, c2, c3 = st.columns(3)
    cliente = c1.text_input("Cliente", key="ao_f_cliente", placeholder="Cliente o Ref.")
    denominacion = c2.text_input("Denominación del suministro", key="ao_f_denom",
                                 placeholder="Alias o CUPS")
    todos = sorted({s["usuario"] for s in sols if s["usuario"]} | ({yo} if yo else set()))
    usuarios = c3.multiselect("Usuario", todos, default=[yo] if yo in todos else [],
                              key="ao_f_usuarios", placeholder="Todos",
                              format_func=lambda u: nombres.get(u, u))
    filtradas = solicitudes_filtradas(sols, cliente, denominacion, usuarios)
    st.caption(f"{len(filtradas)} de {len(sols)} solicitudes" + (
        " (por defecto, las tuyas: quita el filtro de usuario para ver todas)"
        if usuarios == [yo] else ""))
    if filtradas:
        st.dataframe(pd.DataFrame([{
            "Nº": s["id"], "Guardada": _fecha_hora(s["fecha"]),
            "Usuario": nombres.get(s["usuario"], s["usuario"]), "Ref.": s["ref"],
            "Cliente": s["cliente"], "Suministros": s["suministros"],
            "Denominaciones": s.get("denominaciones") or "", "Estado": s["estado"]}
            for s in filtradas]), hide_index=True, use_container_width=True, height=min(
                38 + 35 * len(filtradas), 300))
    return filtradas
def _origen():
    estilo.seccion(1, "Origen", "Solicitud de ofertas guardada o plantilla de suministros")
    modo = st.radio("Partir de", ["Una solicitud de ofertas guardada",
                                  "Suministros de la plantilla Excel"], horizontal=True,
                    key="ao_modo_origen", help=AY["origen"])
    if modo.startswith("Una solicitud"):
        sols = [s for s in almacen.listar_solicitudes() if s.get("producto") != "gas"]
        if not sols:
            st.info("No hay solicitudes de electricidad guardadas. Crea una en «Crear nueva "
                    "solicitud» (con los datos de Gemweb) o usa la plantilla de suministros.")
            return
        sols = filtrar_solicitudes(sols)
        if not sols:
            st.info("Ninguna solicitud con estos filtros.")
            return
        por_id = {s["id"]: s for s in sols}
        elegida = st.selectbox("Solicitud", list(por_id), key="ao_sol", format_func=lambda i: (
            f"Nº {i} · {por_id[i]['cliente'] or '—'}"
            + (f" · Ref. {por_id[i]['ref']}" if por_id[i]["ref"] else "")
            + f" · {por_id[i]['suministros']} suministros · {por_id[i]['estado']}"))
        if st.button("Crear el análisis", type="primary"):
            sol = almacen.solicitud(elegida)
            with st.spinner("Preparando datos regulados, mercado y apuntamientos…"):
                e = servicio.estudio_desde_solicitud(sol)
                avisos = servicio.preparar(e, _apuntamientos_al_dia())
            _empezar(e, avisos, sol.get("ref") or "")
            almacen.registrar_actividad(_usuario(), MODULO, "Nuevo análisis", f"Solicitud {elegida}")
            st.rerun()
    else:
        from ofertas import plantilla as pl
        st.caption("Descarga la plantilla, rellena los suministros y su consumo mes a mes por "
                   "periodo, y súbela.")
        estilo.boton_descarga(MODULO, st, "Descargar plantilla (Excel)", pl.plantilla_excel([]),
                              "plantilla_suministros.xlsx", XLSX, key="ao_pl_descargar")
        c1, c2 = st.columns(2)
        cliente = c1.text_input("Cliente", key="ao_pl_cliente")
        inicio = c2.date_input("Inicio de suministro", format="DD/MM/YYYY", key="ao_pl_inicio",
                               value=mo.sumar_meses(dt.date.today().replace(day=1), 1))
        subida = st.file_uploader("Plantilla rellena", type=["xlsx"], key="ao_pl_fichero")
        if subida and st.button("Crear el análisis", type="primary"):
            leido = pl.leer_plantilla(subida.getvalue())
            errores = leido[2] + pl.errores_de_manuales(leido[0], leido[1], [], False)
            if errores:
                st.error("Corrige la plantilla:\n" + "\n".join("- " + x for x in errores))
                return
            fichas, _ = pl.aplicar_plantilla([], leido)
            with st.spinner("Preparando datos regulados, mercado y apuntamientos…"):
                e = servicio.estudio_desde_fichas(fichas, cliente, inicio)
                avisos = servicio.preparar(e, _apuntamientos_al_dia())
            _empezar(e, avisos)
            st.rerun()


# =================================================================== 2. suministros
def _suministros(e: mo.Estudio):
    estilo.seccion(2, "Suministros", f"{len(e.suministros)} CUPS")
    c1, c2, c3 = st.columns([1, 1, 2])
    inicio = c1.date_input("Inicio de suministro", value=e.inicio, format="DD/MM/YYYY",
                           key=f"ao_inicio_{_v()}")
    if inicio and inicio.replace(day=1) != e.inicio:
        e.inicio = inicio.replace(day=1)
        servicio.congelar_regulados(e)
        st.session_state.ao_avisos = servicio.congelar_mercado(e)
        _subir_version()
        st.rerun()
    n_anios = max(1, math.ceil(e.horizonte() / 12))
    crec = c2.text_input("Crecimiento del consumo por año (%)",
                         value=", ".join(f"{g * 100:g}" for g in e.crecimiento),
                         key=f"ao_crec_{_v()}",
                         help=AY["crecimiento"] + f" Separados por comas (el horizonte tiene "
                              f"{n_anios} años). Vacío = mismo consumo todos los años.")
    try:
        e.crecimiento = [float(x.replace(",", ".")) / 100
                         for x in re.split(r"[;\s]+|,\s+", crec.strip()) if x]
    except ValueError:
        c2.error("Escribe porcentajes separados por comas.")
    e.cliente = c3.text_input("Cliente", value=e.cliente, key=f"ao_cliente_{_v()}")
    st.caption("Sistema eléctrico propuesto por la dirección (revísalo). Reducción del impuesto "
               "eléctrico: 0, 0,85 (usos del art. 98.1) o 1 (exento). Apuntamiento: «Del "
               "cliente» solo en los CUPS con curva de un año completo (sección 3).")
    editada = editor(
        tabla_suministros(e), f"ao_sum_{_v()}",
        ["Alquiler (€/mes)"] + [f"Pot. {p} (kW)" for p in PER],
        hide_index=True, use_container_width=True, disabled=["CUPS", "Consumo anual (kWh)"],
        column_config={
            "Tarifa": st.column_config.SelectboxColumn(options=mo.TARIFAS, required=True),
            "Sistema": st.column_config.SelectboxColumn(options=mo.SISTEMAS, required=True),
            "Reducción IEE": st.column_config.SelectboxColumn(options=mo.REDUCCIONES_IEE,
                                                              required=True),
            "Apuntamiento": st.column_config.SelectboxColumn(options=[mo.ESTIMADO, mo.CLIENTE],
                                                             required=True),
            "Consumo anual (kWh)": st.column_config.NumberColumn(format="localized"),
            "Incluir": st.column_config.CheckboxColumn()})
    aplicar_suministros(e, editada)


# =================================================================== 3. mercado
def _avisar_regulados(e: mo.Estudio):
    """Si los datos regulados congelados en el análisis ya no coinciden con los de
    Administración (p. ej. un año se congeló antes de cambiar el FNEE), avisa de qué cambia y
    deja actualizarlos."""
    try:
        cambios = servicio.diferencias_regulados(e)
    except Exception:                                       # noqa: BLE001
        return
    if not cambios:
        return
    c1, c2 = st.columns([3, 1], vertical_alignment="center")
    c1.warning("Los datos regulados de este análisis no coinciden con los de Administración → "
               "Datos regulados (se congelaron antes de cambiarlos):  \n"
               + "  \n".join("- " + x for x in cambios[:10])
               + ("  \n- …" if len(cambios) > 10 else ""), icon=":material/warning:")
    if c2.button("Actualizar datos regulados", type="primary", key=f"ao_reg_act_{_v()}",
                 help="Vuelve a tomar todos los años de Administración → Datos regulados. "
                      "Recuerda guardar el análisis después."):
        servicio.congelar_regulados(e, todos=True)
        almacen.registrar_actividad(_usuario(), MODULO, "Actualizar datos regulados",
                                    "; ".join(cambios)[:300])
        _subir_version()
        st.rerun()


def _mercado(e: mo.Estudio):
    estilo.seccion(3, "Mercado, SSAA y apuntamientos",
                   f"OMIP sesión {e.sesion_omip or '—'} · {e.pct_omip * 100:.0f} % OMIP")
    estado = dm.estado()
    for a in estado["avisos"]:
        st.warning(a)
    c1, c2 = st.columns([3, 1])
    c1.caption(f"OMIP: última sesión del fichero {estado['omip']['ultima_sesion'] or '—'} · OMIE: "
               f"hasta {MESES[estado['omie']['ultimo_mes'][1] - 1] + ' ' + str(estado['omie']['ultimo_mes'][0]) if estado['omie']['ultimo_mes'] else '—'}. "
               "El estudio usa los datos congelados al crearlo; los OMIP cambiados a mano se "
               "conservan al actualizar.")
    if c2.button("Actualizar mercado", help=AY["actualizar_mercado"]):
        st.session_state.ao_avisos = servicio.congelar_mercado(e)
        _subir_version()
        st.rerun()
    _avisar_regulados(e)
    t_merc, t_ssaa, t_ap, t_curvas = st.tabs(["Precio de mercado", "SSAA", "Apuntamientos",
                                              "Curvas del cliente"])
    with t_merc:
        st.caption("Precio previsto = % OMIP × OMIP + (1 − % OMIP) × media de OMIE del mes. "
                   "Cambia el OMIP de un mes para ajustarlo a mano.")
        c_ed, c_prev = st.columns([3, 1])
        with c_ed:
            editada = editor(tabla_mercado(e).drop(columns=["Previsto (€/MWh)"]),
                             f"ao_merc_{_v()}_{e.horizonte()}",
                             ["OMIP (€/MWh)", "Media OMIE (€/MWh)"],
                             {"OMIP (€/MWh)": 2, "Media OMIE (€/MWh)": 2},
                             hide_index=True, use_container_width=True,
                             disabled=["Mes", "Contrato"], column_config={"Clave": None})
        aplicar_mercado(e, editada)
        with c_prev:
            previsto = tabla_mercado(e)[["Mes", "Previsto (€/MWh)"]].copy()
            previsto["Previsto (€/MWh)"] = [texto_es(v, 2) for v in previsto["Previsto (€/MWh)"]]
            st.dataframe(previsto, hide_index=True, use_container_width=True)
    with t_ssaa:
        st.caption("SSAA medios de cada mes natural (€/MWh). Por defecto, los SSAA estimados de "
                   "Parámetros (Administración → Datos regulados); se pueden cambiar en cada "
                   "estudio.")
        if st.button("Poner los SSAA estimados de Parámetros", key=f"ao_ssaa_def_{_v()}",
                     help="Sustituye los 12 valores de este estudio por los SSAA estimados de "
                          "Parámetros."):
            if servicio.ssaa_por_defecto(e, forzar=True):
                _subir_version()
                st.rerun()
            st.warning("En Parámetros no hay SSAA estimados para el año de inicio.")
        c1, c2 = st.columns([4, 1], vertical_alignment="bottom")
        pegado = c1.text_area(
            "Pegar los 12 valores de una vez", key=f"ao_ssaa_pegar_{_v()}", height=68,
            placeholder="Copia en Excel las 12 celdas de enero a diciembre (en fila o en "
                        "columna) y pégalas aquí con Ctrl+V",
            help="Vale una fila o una columna de Excel, con coma decimal. Los 12 valores van "
                 "de enero a diciembre.")
        if c2.button("Poner los 12 valores", key=f"ao_ssaa_poner_{_v()}",
                     disabled=not (pegado or "").strip(), use_container_width=True):
            valores = valores_pegados(pegado)
            if len(valores) != 12:
                st.error(f"Se han pegado {len(valores)} valores y hacen falta 12 (de enero a "
                         "diciembre).")
            elif any(v is None for v in valores):
                st.error("Hay celdas vacías o que no son números: revisa lo copiado.")
            else:
                e.ssaa_mes = valores
                _subir_version()
                st.rerun()
        st.caption("También puedes escribir o pegar en la tabla mes a mes.")
        aplicar_ssaa(e, editor(tabla_ssaa(e), f"ao_ssaa_{_v()}", ["SSAA (€/MWh)"],
                               {"SSAA (€/MWh)": 4}, hide_index=True, disabled=["Mes"]))
        if MOSTRAR_PH and any(s.activo and s.sistema == "Canarias" for s in e.suministros):
            st.markdown("**Ph demanda de Canarias**")
            st.caption("Lo que el Ph demanda (ESIOS, indicador 573) lleva encima del mercado "
                       "diario: costes de ajuste, capacidad... Media del mes de los 12 últimos "
                       "meses. Las ofertas indexadas al Ph usan: (previsto + este valor) × "
                       "apuntamiento «Canarias Ph». Vacío = el SSAA del mes.")
            aplicar_ph(e, editor(tabla_ph(e), f"ao_ph_{_v()}", ["Ph − diario (€/MWh)"],
                                 {"Ph − diario (€/MWh)": 2}, hide_index=True,
                                 disabled=["Mes", "Dato de"]))
    with t_ap:
        _apuntamientos(e)
    with t_curvas:
        _curvas(e)


def _apuntamientos(e: mo.Estudio):
    usados = claves_ap(e)
    if not usados:
        return
    st.caption("Apuntamiento del precio del mercado diario (precio medio de las horas de cada "
               "periodo ÷ precio medio del mes) de los 12 últimos meses publicados, calculado "
               "por la aplicación con ESIOS (y los Excel de OMIE de respaldo). «Dato de» indica "
               "el mes y la fuente; lo cambiado a mano queda como «manual». Vacío = 1. "
               "«Canarias Ph»: forma del Ph demanda de Canarias (ESIOS, indicador 573), para "
               "las ofertas indexadas al Ph.")
    clave = st.selectbox("Sistema y calendario", usados, key="ao_ap_clave",
                         help=AY["apuntamiento"],
                         format_func=lambda c: c.replace("|3P", " · 2.0TD").replace(
                             "|6P", " · 6 periodos"))
    tabla = tabla_ap(e, clave)
    editada = editor(tabla, f"ao_ap_{clave}_{_v()}", [p for p in PER if p in tabla.columns],
                     hide_index=True, use_container_width=True, disabled=["Mes", "Dato de"])
    aplicar_ap(e, clave, editada)
    c1, c2 = st.columns(2)
    if c1.button("Actualizar con los últimos calculados",
                 help="Sustituye todos los apuntamientos (también los cambiados a mano)."):
        _ultimo_intento[0] = 0.0                     # el usuario lo pide: se intenta ya
        servicio.apuntamientos_vigentes(e, _apuntamientos_al_dia())
        _subir_version()
        st.rerun()
    otros = [a for a in almacen.listar_analisis() if a["id"] != st.session_state.get("ao_id")]
    if otros:
        por_id = {a["id"]: a for a in otros}
        de = c2.selectbox("Copiar de otro análisis", [None] + list(por_id), key="ao_ap_copiar",
                          format_func=lambda i: "—" if i is None else
                          f"Nº {i} · {por_id[i]['cliente'] or '—'}")
        if de and c2.button("Copiar sus apuntamientos"):
            otro = servicio.estudio_de_datos(almacen.analisis(de)["datos"])
            e.ap_omie, e.ap_ssaa, e.ap_origen = otro.ap_omie, otro.ap_ssaa, otro.ap_origen
            e.ph_dif, e.ph_origen, e.pico = otro.ph_dif, otro.ph_origen, otro.pico
            _subir_version()
            st.rerun()


def _curvas(e: mo.Estudio):
    from curva_consumo import ErrorCurva, leer_curva
    st.caption("Sube la curva horaria (o cuartohoraria) de un año completo de un CUPS para "
               "calcular su apuntamiento real (precio ponderado por su consumo) con los precios "
               "horarios del mismo año. Después elige «Del cliente (curva)» en la tabla de "
               "suministros o con el botón de abajo. Las curvas se guardan con el análisis.")
    curvas = st.session_state.setdefault("ao_curvas", {})
    _comprobar_gemweb(e, curvas)
    st.markdown("**O sube la curva de un CUPS en un fichero**")
    c1, c2 = st.columns([2, 3])
    cups = c1.selectbox("CUPS", [s.cups for s in e.suministros if s.activo], key="ao_curva_cups")
    fichero = c2.file_uploader("Curva (CSV o Excel de Gemweb, Datadis o distribuidora)",
                               type=["csv", "xlsx", "xls"], key=f"ao_curva_fich_{_v()}")
    if fichero and st.button("Leer la curva y calcular su apuntamiento"):
        try:
            curva = leer_curva(fichero.getvalue(), fichero.name)
        except ErrorCurva as err:
            st.error(f"No se ha podido leer la curva: {err}")
            return
        valores = curva.valores
        if curva.resolucion == "qh":
            horaria = {}
            for (f, h, _), v in valores.items():
                horaria[(f, h)] = horaria.get((f, h), 0.0) + v
            valores = horaria
        curvas[cups] = valores
        with st.spinner("Calculando con los precios horarios del mercado…"):
            r = servicio.apuntamientos_cliente(e, {cups: valores})
        if r[cups]["valida"]:
            st.success(f"{cups}: apuntamiento calculado ({r[cups]['cobertura']['desde']:%d/%m/%Y} – "
                       f"{r[cups]['cobertura']['hasta']:%d/%m/%Y}).")
        else:
            st.warning(f"{cups}: {r[cups]['motivo']} Se usará el estimado.")
        _subir_version()
    if e.ap_cliente or curvas:
        filas = []
        for c in sorted(set(curvas) | set(e.ap_cliente)):
            a = e.ap_cliente.get(c)
            filas.append({"CUPS": c, "Válida": bool(a), "Desde": a.get("desde") if a else None,
                          "Hasta": a.get("hasta") if a else None,
                          "Días": a.get("dias") if a else None,
                          "Usa": (e.suministro(c).apuntamiento if e.suministro(c) else "")})
        st.dataframe(pd.DataFrame(filas), hide_index=True, use_container_width=True)
        b1, b2 = st.columns(2)
        if b1.button("Usar el del cliente en todos los CUPS con curva válida"):
            for s in e.suministros:
                if e.ap_cliente.get(s.cups):
                    s.apuntamiento = mo.CLIENTE
            _subir_version()
            st.rerun()
        if b2.button("Usar el estimado en todos"):
            for s in e.suministros:
                s.apuntamiento = mo.ESTIMADO
            _subir_version()
            st.rerun()


def tabla_eleccion(e: mo.Estudio, filas: list) -> pd.DataFrame:
    """Resultado de la comprobación con la columna «Usar su curva» (lo que usa cada CUPS)."""
    return pd.DataFrame([dict(f, **{"Usar su curva": bool(
        e.suministro(f["CUPS"]) and e.suministro(f["CUPS"]).apuntamiento == mo.CLIENTE)})
        for f in filas])


def _comprobar_gemweb(e: mo.Estudio, curvas: dict):
    """Busca en Gemweb la curva de todos los suministros, enseña de qué fechas es el año que se
    usaría y deja que el usuario decida en cuáles se usa."""
    import gemweb
    st.markdown("**Comprobar la curva horaria en Gemweb**")
    st.caption("Busca la curva de todos los suministros en los últimos 24 meses y toma los 12 "
               "últimos meses completos que tenga cada uno. Si es un año completo calcula su "
               "apuntamiento; después decides en cuáles se usa. Puede tardar unos minutos.")
    if st.button("Comprobar en Gemweb", key="ao_gemweb"):
        cred, _origen = gemweb.cargar_credenciales()
        if not cred:
            st.error("No hay credenciales de Gemweb (Administración → Credenciales de Gemweb).")
            return
        barra = st.progress(0.0, text="Conectando con Gemweb…")
        try:
            cliente = gemweb.ClienteGemweb(*cred)
            nuevas, filas = servicio.comprobar_curvas_gemweb(
                e, cliente, al_avanzar=lambda f, t: barra.progress(min(1.0, f), text=t))
        except Exception as err:                            # noqa: BLE001
            st.error(f"No se ha podido consultar Gemweb: {err}")
            return
        curvas.update(nuevas)
        st.session_state.ao_gemweb_filas = filas
        almacen.registrar_actividad(_usuario(), MODULO, "Consulta Gemweb",
                                    f"Curvas de {len(filas)} suministros")
        _subir_version()
    filas = st.session_state.get("ao_gemweb_filas")
    if not filas:
        return
    validas = sum(1 for f in filas if f["Se puede usar"] == "Sí")
    st.markdown(f"{validas} de {len(filas)} suministros con un año completo de curva. Marca "
                "«Usar su curva» en los que quieras usar su apuntamiento (solo en los que se "
                "puede) y pulsa «Guardar la elección».")
    fechas = {c: st.column_config.DateColumn(format="DD/MM/YYYY") for c in (
        "Datos desde", "Datos hasta", "Año usado desde", "Año usado hasta")}
    elegida = editor(tabla_eleccion(e, filas), f"ao_gemweb_tabla_{_v()}", hide_index=True,
                     use_container_width=True,
                     disabled=[c for c in filas[0] if c != "Usar su curva"],
                     column_config=dict(fechas, **{
                         "Usar su curva": st.column_config.CheckboxColumn()}))
    if st.button("Guardar la elección", key="ao_gemweb_elegir"):
        no = servicio.elegir_apuntamiento(e, {f["CUPS"]: bool(f["Usar su curva"])
                                              for _, f in elegida.iterrows()})
        _subir_version()
        if no:
            st.warning("Sin un año completo de curva, se queda el estimado en: " + ", ".join(no))
        else:
            st.rerun()


# =================================================================== 4-5. ofertas

def _sel(contenedor, etiqueta, opciones, valor, clave, **kw):
    """Desplegable que cambia a la primera: el valor inicial va al estado de la sesión y no
    se pasa «index» (si «index» cambia entre ejecuciones, Streamlit rehace el control y hay
    que elegir dos veces)."""
    if clave not in st.session_state or st.session_state[clave] not in opciones:
        st.session_state[clave] = valor if valor in opciones else opciones[0]
    return contenedor.selectbox(etiqueta, opciones, key=clave, **kw)


def _chk(contenedor, etiqueta, valor, clave, **kw) -> bool:
    """Casilla con el mismo arreglo que _sel."""
    if clave not in st.session_state:
        st.session_state[clave] = bool(valor)
    return contenedor.checkbox(etiqueta, key=clave, **kw)


def _pegar_precios(o: mo.Oferta, concepto: str, tabla: pd.DataFrame, clave: str):
    """Cuadro para pegar los precios copiados de Excel (una fila con P1…P6, varias filas o
    una columna). Funciona aunque el navegador no deje pegar en la tabla (sin https)."""
    with st.expander("Pegar precios de Excel", icon=":material/content_paste:"):
        filas = [f"{g} · año {a}" for g, a in zip(tabla["Grupo"], tabla["Año"])]
        c1, c2 = st.columns([3, 2])
        texto = c1.text_area(
            "Celdas copiadas de Excel", key=clave + "_txt", height=68,
            placeholder="Copia en Excel la fila con los precios de P1 a P6 (o varias filas) y "
                        "pégala aquí con Ctrl+V",
            help="Una fila de Excel rellena P1, P2… de la fila elegida (con 5 celdas, P6 no se "
                 "toca). Varias filas rellenan las filas siguientes de la tabla. Una columna "
                 "de hasta 6 celdas se toma como P1, P2… Coma o punto decimal.")
        desde = c2.selectbox("Empezar en la fila", range(len(filas)), key=clave + "_fila",
                             format_func=lambda i: filas[i]) if len(filas) > 1 else 0
        if c2.button("Poner precios", key=clave + "_btn", type="primary",
                     disabled=not (texto or "").strip(), use_container_width=True):
            nueva, error = pegar_precios(tabla, texto, desde or 0)
            if error:
                st.error(error)
            else:
                aplicar_precios(o, concepto, nueva)
                _subir_version()
                st.rerun()


def _importe(contenedor, etiqueta, o: mo.Oferta, attr: str, clave: str, escala: float = 1.0,
             **kw):
    """Importe escrito como texto (coma decimal): se guarda tal y como se escribe, con sus
    decimales (0,394 sigue siendo 0,394; 1,5228 sigue siendo 1,5228). escala: 100 para %."""
    valor = float(getattr(o, attr) or 0) * escala
    texto = (o.textos or {}).get(attr)
    if texto is None or numero_es(texto) is None or abs(numero_es(texto) - valor) > 1e-9:
        texto = texto_es(valor)
    if clave not in st.session_state:
        st.session_state[clave] = texto

    def a_coma():                      # al escribir 0.394 se ve 0,394 (mismos decimales)
        st.session_state[clave] = con_coma(st.session_state.get(clave))

    escrito = contenedor.text_input(etiqueta, key=clave, on_change=a_coma, **kw).strip()
    numero = 0.0 if escrito == "" else numero_es(escrito)
    if numero is None:
        contenedor.caption(":red[No es un número: se mantiene el anterior]")
        return
    setattr(o, attr, numero / escala)
    if o.textos is None:
        o.textos = {}
    o.textos[attr] = escrito


def formulario_oferta(e: mo.Estudio, o: mo.Oferta, k: str):
    """Ficha de una oferta (o del contrato actual) en pestañas. k: prefijo de las claves."""
    k = f"ao_{k}_{_v()}"
    t_conf, t_precios, t_grupos, t_cierres, t_cond = st.tabs(
        ["Configuración", "Precios", "Grupos y CUPS", "Cierres OMIP", "Condiciones"])
    with t_conf:
        c1, c2, c3 = st.columns(3)
        opciones = coms.opciones(o.comercializadora)
        o.comercializadora = _sel(c1, "Comercializadora", opciones, o.comercializadora,
                                  k + "com", format_func=lambda c: c or "— Elige —",
                                  help="Escribe para buscar en la lista.")
        # el gestor y su email se escriben en cada oferta: no se copian de otras ofertas (un
        # mismo gestor no lleva a todos los clientes de la comercializadora)
        o.gestor = c2.text_input("Gestor de la comercializadora", o.gestor, key=k + "gestor")
        o.email_gestor = c3.text_input("Email del gestor", o.email_gestor, key=k + "email").strip()
        if o.email_gestor and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", o.email_gestor):
            c3.caption(":red[Email no válido]")
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        o.producto = c1.text_input("Producto / descripción", o.producto, key=k + "prod")
        o.modalidad = _sel(c2, "Modalidad", mo.MODALIDADES, o.modalidad, k + "mod",
                           help=AY["modalidad"])
        o.duracion = int(c3.number_input("Duración (meses)", 1, 240, int(o.duracion or 12),
                                         key=k + "dur", help=AY["duracion"]))
        if not o.es_actual:
            o.ronda = int(c4.number_input("Ronda", 1, 20, int(o.ronda or 1), key=k + "ronda",
                                          help=AY["ronda"]))
        c1, c2, c3, c4 = st.columns(4)
        propio = _chk(c1, "Inicio distinto al del estudio", bool(o.inicio), k + "inip",
                      help=AY["inicio"])
        if propio:
            ini = c1.date_input("Inicio", o.inicio or e.inicio, format="DD/MM/YYYY",
                                key=k + "ini")
            o.inicio = ini.replace(day=1) if ini else None
        else:
            o.inicio = None
        # la unidad y el descuento solo en fijo; la potencia siempre en la segunda columna
        if o.modalidad == mo.FIJO:
            o.unidad = _sel(c3, "Unidad de la energía", list(mo.UNIDADES), o.unidad, k + "uni",
                            help=AY["unidad"])
            _importe(c4, "Descuento sobre energía (%)", o, "descuento", k + "desc", 100,
                     help=AY["descuento"])
        o.potencia = _sel(c2, "Término de potencia", [mo.POT_REGULADA, mo.POT_PROPIA],
                          o.potencia, k + "pot", help=AY["potencia"])
        o.atr_incluido = _chk(c1, "El precio incluye peajes y cargos", o.atr_incluido,
                              k + "atr", help=AY["atr"])
        # Ph de Canarias: de momento no se usa
        if not MOSTRAR_PH:
            o.indice_ph = False
        if o.modalidad == mo.PASS_THROUGH:
            st.markdown("**Componentes de mercado que suma la oferta**")
            c1, c2, c3, c4, c5 = st.columns(5)
            o.pagos_capacidad = _chk(c1, "Pagos por capacidad", o.pagos_capacidad, k + "pc",
                                     help=AY["pc"])
            o.os = _chk(c2, "Pago OS", o.os, k + "os", help=AY["os"])
            o.om = _chk(c3, "Pago OM", o.om, k + "om", help=AY["om"])
            o.srad = _chk(c4, "SRAD", o.srad, k + "srad", help=AY["srad"])
            ref = referencia_desvios(e)
            _importe(c5, "Desvíos (€/MWh, en barras)", o, "desvios", k + "desv",
                     help=AY["desvios"] + ("  \n" + ref if ref else ""))
            c1, c2 = st.columns(2)
            o.base_b = _sel(c1, "Base del término B (si lo hay)", mo.BASES, o.base_b, k + "bb",
                            help=AY["base_b"])
        elif o.modalidad == mo.PASS_POOL:
            # Pass-Pool: Ci × OMIE + término B (que ya lleva capacidad, OS, OM, SRAD y desvíos)
            o.pagos_capacidad = o.os = o.om = o.srad = False
            o.desvios = 0.0
            st.markdown("**Precio Pass-Pool: Ci × OMIE + término B** (capacidad, OS, OM, SRAD y "
                        "desvíos van dentro del término B)")
            c1, c2 = st.columns(2)
            o.base_b = _sel(c1, "Base del término B", mo.BASES, o.base_b, k + "bb",
                            help=AY["base_b"])
            o.ci_media_mensual = _chk(c2, "El Ci se aplica a la media mensual de OMIE",
                                      o.ci_media_mensual, k + "cimm", help=AY["cimm"])
        if MOSTRAR_PH and o.modalidad != mo.FIJO:
            o.indice_ph = _chk(st, "Canarias indexado al Ph demanda", o.indice_ph, k + "ph",
                               help=AY["ph"])
        st.markdown("**Otros conceptos**")
        c1, c2, c3, c4 = st.columns(4)
        _importe(c1, "Fee / gestión (€/MWh)", o, "fee", k + "fee", help=AY["fee"])
        o.base_fee = _sel(c2, "Base del fee", mo.BASES, o.base_fee, k + "bfee",
                          help=AY["base_fee"])
        _importe(c3, "Garantías de origen (€/MWh)", o, "gdo", k + "gdo", help=AY["gdo"])
        o.fnee = _sel(c4, "FNEE", mo.MODOS_FNEE, o.fnee, k + "fnee", help=AY["fnee"])
        c1, c2, c3 = st.columns(3)
        _importe(c1, "K / impuesto municipal (%)", o, "k", k + "k", 100, help=AY["k"])
        o.k_ssaa = _chk(c2, "La K se aplica a los SSAA", o.k_ssaa, k + "kss",
                        help=AY["k_ssaa"])
        _importe(c3, "Otros costes fijos (€/CUPS·mes)", o, "otros_fijos", k + "otf",
                 help=AY["otros_fijos"])
        st.markdown("**Servicios de ajuste (SSAA)**")
        c1, c2, c3, c4 = st.columns(4)
        o.ssaa_modo = _sel(c1, "Modo", mo.MODOS_SSAA, o.ssaa_modo, k + "ssm",
                           help=AY["ssaa_modo"])
        if o.ssaa_modo == mo.SSAA_BANDA:
            _importe(c2, "Banda mínima (€/MWh)", o, "banda_min", k + "bmin", help=AY["banda_min"])
            _importe(c3, "Banda máxima (€/MWh)", o, "banda_max", k + "bmax", help=AY["banda_max"])
        if o.ssaa_modo == mo.SSAA_TECHO:
            _importe(c2, "Techo (€/MWh)", o, "techo", k + "techo", help=AY["techo"])
        if o.ssaa_modo in (mo.SSAA_BANDA, mo.SSAA_TECHO):
            o.ssaa_perfil = _chk(c4, "Aplicar perfil por periodo", o.ssaa_perfil, k + "sperf",
                                 help=AY["ssaa_perfil"])
        if o.ssaa_modo != mo.SSAA_INCLUIDOS:
            c1, c2, c3, c4 = st.columns(4)
            o.ssaa_perdidas = _sel(c1, "Elevación a pérdidas", mo.PERDIDAS_SSAA,
                                   o.ssaa_perdidas, k + "sloss", help=AY["ssaa_perdidas"])
            if o.ssaa_perdidas == mo.PERD_FIJO:
                _importe(c2, "% de pérdidas", o, "ssaa_pct_perdidas", k + "slp", 100)
            _importe(c3, "Factor 1", o, "ssaa_f1", k + "sf1", help=AY["ssaa_f"])
            _importe(c4, "Factor 2", o, "ssaa_f2", k + "sf2", help=AY["ssaa_f"])
        if o.modalidad != mo.FIJO:
            st.markdown("**Cierres en OMIP**")
            c1, c2 = st.columns(2)
            o.cierres_tipo = _sel(c1, "Tipo de cierres", mo.TIPOS_CIERRE, o.cierres_tipo,
                                  k + "cie", help=AY["cierres"])
            if o.cierres_tipo != mo.SIN_CIERRES:
                _importe(c2, "Coste por click (€/MWh cerrado)", o, "coste_click", k + "click",
                         help=AY["click"])
        else:
            o.cierres_tipo = mo.SIN_CIERRES
        # fórmula resultante y avisos de coherencia (para compararla con la del PDF)
        formula, leyenda = ayudas.formula_latex(o)
        with st.container(border=True):
            st.markdown("##### Fórmula de esta oferta")
            st.caption("Precio de la energía de cada periodo, tal y como lo calcula la "
                       "aplicación. Compárala con la fórmula del contrato de la comercializadora.")
            st.latex(formula)
            st.caption("  \n".join("· " + x for x in leyenda))
        for aviso in ayudas.avisos_coherencia(o):
            st.warning(aviso)
    with t_precios:
        unidad = o.unidad if o.modalidad == mo.FIJO else ""
        st.caption("Una fila por grupo (tarifa y sistema, o grupo propio) y año de contrato. "
                   "Los grupos de Baleares, Canarias, Ceuta y Melilla sin precios usan los de su "
                   "tarifa; un año sin precios, los del anterior.")
        for concepto in conceptos_precio(o):
            st.markdown(f"**{CONCEPTOS_PRECIO[concepto]}**" + (f" ({unidad})" if concepto ==
                                                                "energia" else ""))
            tabla = tabla_precios(e, o, concepto)
            _pegar_precios(o, concepto, tabla, f"{k}pg_{concepto}")
            editada = editor(tabla, f"{k}pr_{concepto}_{len(tabla)}_{o.modalidad}_{o.unidad}",
                             PER, hide_index=True, use_container_width=True,
                             disabled=["Grupo", "Año"])
            aplicar_precios(o, concepto, editada)
    with t_grupos:
        st.caption("Cada CUPS va solo al grupo de su tarifa. Crea un grupo propio solo si la "
                   "oferta da a algunos CUPS un precio distinto del resto de su tarifa.")
        with st.expander("¿Cómo funcionan los grupos propios? (paso a paso y ejemplo)"):
            st.markdown(ayudas.GUIA_GRUPOS)
        st.markdown("**Grupos propios** (nombre, tarifa y sistema)")
        grupos = editor(tabla_grupos_propios(o), k + "gp", hide_index=True, num_rows="dynamic",
                        column_config={
                                    "Tarifa": st.column_config.SelectboxColumn(options=mo.TARIFAS),
                                    "Sistema": st.column_config.SelectboxColumn(
                                        options=mo.SISTEMAS)})
        aplicar_grupos_propios(o, grupos)
        st.markdown("**CUPS de la oferta**: «Incluido» y grupo propio de cada uno")
        cups = editor(
            tabla_cups_oferta(e, o), f"{k}cups_{len(o.grupos_propios)}", hide_index=True,
            use_container_width=True, disabled=["CUPS", "Alias", "Tarifa", "Sistema"],
            column_config={"Grupo propio": st.column_config.SelectboxColumn(
                options=[""] + [g.nombre for g in o.grupos_propios])})
        aplicar_cups_oferta(o, cups)
    with t_cierres:
        with st.expander("¿Cómo funcionan los cierres en OMIP? (por % y por MW, con ejemplos)",
                         expanded=o.cierres_tipo != mo.SIN_CIERRES and not o.cierres):
            st.markdown(ayudas.GUIA_CIERRES)
        if o.cierres_tipo == mo.SIN_CIERRES:
            st.caption("Sin cierres (elige el tipo en Configuración, solo indexados).")
        else:
            st.caption("Por %: parte del consumo cerrada a ese precio. Por MW: MW de carga base "
                       "cerrados (se liquida la diferencia con el previsto en el consumo de la "
                       "cartera en barras). Precio vacío = OMIP del mes. MW pico: carga de 8 a "
                       "20 h de lunes a viernes; se liquida contra el previsto del pico "
                       "(previsto × apuntamiento del pico del mercado diario); precio pico "
                       "vacío = ese previsto.")
            st.caption("% cerrado en tanto por uno (0,5 = 50 %).")
            editada = editor(
                tabla_cierres(e, o), f"{k}cier_{o.duracion}_{o.cierres_tipo}",
                ["Previsto (€/MWh)", "OMIP (€/MWh)", "% cerrado", "MW",
                 "Precio de cierre (€/MWh)", "MW pico", "Precio pico (€/MWh)"],
                {"Previsto (€/MWh)": 2, "OMIP (€/MWh)": 2},
                hide_index=True, use_container_width=True,
                disabled=["Mes", "Fecha", "Previsto (€/MWh)", "OMIP (€/MWh)"],
                column_config=({} if o.cierres_tipo == mo.CIERRE_MW else
                               {"MW pico": None, "Precio pico (€/MWh)": None}))
            aplicar_cierres(o, editada)
    with t_cond:
        c1, c2 = st.columns(2)
        for n, (clave, nombre) in enumerate((("validez", "Validez de la oferta"),
                                             ("pago", "Forma y plazo de pago"),
                                             ("garantias", "Garantías / depósito"),
                                             ("banda", "Banda de volumen / penalizaciones"),
                                             ("rescision", "Rescisión / permanencia"),
                                             ("gdo", "Garantías de origen"))):
            o.condiciones[clave] = (c1 if n % 2 == 0 else c2).text_input(
                nombre, o.condiciones.get(clave, ""), key=k + "cond_" + clave)
        o.notas = st.text_area("Notas", o.notas, key=k + "notas")


def _contrato_actual(e: mo.Estudio):
    estilo.seccion(4, "Contrato actual", "Se proyecta a toda la duración de las ofertas")
    st.caption("Se calcula con la misma ficha que una oferta. Para comparar cada oferta se suma "
               "el coste del contrato actual en sus mismos meses; si dura menos, se repiten los "
               "precios de su último año.")
    formulario_oferta(e, e.actual, "actual")


def _ofertas(e: mo.Estudio):
    estilo.seccion(5, "Ofertas", f"{len(e.ofertas)} oferta(s)")
    c1, c2, c3, c4 = st.columns(4)
    with st.expander("Guía rápida: cómo meter una oferta"):
        st.markdown(ayudas.GUIA)
    _cargar_catalogo(e)
    for aviso in st.session_state.pop("ao_avisos_catalogo", []):
        st.warning(aviso)
    if c1.button("Añadir oferta", type="primary"):
        e.ofertas.append(mo.Oferta(nombre=f"Oferta {len(e.ofertas) + 1:02d}"))
        st.session_state.ao_oferta = len(e.ofertas) - 1
        _subir_version()
        st.rerun()
    if not e.ofertas:
        st.info("Añade las ofertas recibidas.")
        return
    i = st.session_state.get("ao_oferta", 0)
    i = min(i, len(e.ofertas) - 1)
    i = st.radio("Oferta", list(range(len(e.ofertas))), index=i, horizontal=True,
                 key=f"ao_oferta_sel_{_v()}", format_func=lambda n: (
                     f"{e.ofertas[n].comercializadora or e.ofertas[n].nombre} "
                     f"{e.ofertas[n].duracion} m" + (f" · R{e.ofertas[n].ronda}"
                                                    if e.ofertas[n].ronda > 1 else "")))
    st.session_state.ao_oferta = i
    o = e.ofertas[i]
    if c2.button("Duplicar"):
        e.ofertas.insert(i + 1, mo.duplicar(o))
        st.session_state.ao_oferta = i + 1
        _subir_version()
        st.rerun()
    if c3.button("Nueva ronda", help="Copia la oferta como siguiente ronda de esa "
                                     "comercializadora."):
        e.ofertas.insert(i + 1, mo.duplicar(o, nueva_ronda=True))
        st.session_state.ao_oferta = i + 1
        _subir_version()
        st.rerun()
    if c4.button("Quitar esta oferta"):
        e.ofertas.pop(i)
        st.session_state.ao_oferta = max(0, i - 1)
        _subir_version()
        st.rerun()
    _origen_catalogo(e, o, f"ao_of{i}_{_v()}")
    o.nombre = st.text_input("Nombre de la oferta", o.nombre, key=f"ao_of{i}_{_v()}_nombre")
    formulario_oferta(e, o, f"of{i}")
    _guardar_catalogo(e, o, i)


# =================================================================== catálogo de ofertas
NO_SE_GUARDA = ("el nombre de esta oferta en el análisis, el gestor y su email, la ronda, la "
                "fecha de inicio, los grupos propios y sus CUPS, los CUPS excluidos, los cierres "
                "de meses concretos y los apuntamientos propios")


def _es_admin() -> bool:
    return bool((st.session_state.get("usuario") or {}).get("admin"))


def _f(v) -> str:
    return dt.date.fromisoformat(str(v)[:10]).strftime("%d/%m/%Y") if v else ""


def tabla_catalogo(filas: list, e: mo.Estudio = None, nombres: dict = None) -> pd.DataFrame:
    """Tabla del buscador: una fila por entrada (con «Cubre» si se pasa el análisis)."""
    nombres = nombres or {}
    datos = []
    for r in filas:
        fila = {"Nº": r["id"], "Comercializadora": r["comercializadora"], "Nombre": r["nombre"],
                "Modalidad": r["modalidad"], "Meses": r["duracion"],
                "Vigencia": f"{_f(r['vigencia_desde'])} - {_f(r['vigencia_hasta'])}",
                "Inicio de suministro": f"{_f(r['inicio_desde'])} - {_f(r['inicio_hasta'])}",
                "Tarifas con precio": ", ".join(r.get("tarifas") or []) or "(sin precios por tarifa)"}
        if e is not None:
            faltan = catalogo.sin_precio(r, e)
            fila["Cubre mis tarifas"] = "✔" if not faltan else "⚠ falta " + ", ".join(faltan)
        fila.update({"Versión": r["version"], "Estado": r["estado"],
                     "Autor": nombres.get(r["usuario"], r["usuario"])})
        datos.append(fila)
    return pd.DataFrame(datos)


def guardar_en_catalogo(e: mo.Estudio, ofertas: list, nombre: str, fechas: dict, usuario: str,
                        id_analisis=None, como_version: bool = False) -> tuple:
    """Guarda las ofertas en el catálogo (una entrada por oferta; con varias, el nombre lleva
    la duración). como_version: las que vienen del catálogo y el usuario puede tocar se guardan
    como versión nueva de su entrada. Devuelve (mensajes, errores)."""
    mensajes, errores = [], []
    for o in ofertas:
        nom = nombre if len(ofertas) == 1 else f"{nombre} · {o.duracion} M"
        ficha = catalogo.ficha_de(o, nom, id_analisis=id_analisis, **fechas)
        datos = catalogo.a_catalogo(o)
        malos = catalogo.validar(ficha, datos)
        if malos:
            errores += [f"«{o.nombre}»: {x}" for x in malos]
            continue
        origen = (o.origen_catalogo or {}).get("id")
        if como_version and origen and almacen.catalogo(origen) and \
                almacen.actualizar_catalogo(origen, usuario, ficha, datos):
            v = almacen.versiones_catalogo([origen]).get(origen, 1)
            o.origen_catalogo = {"id": origen, "version": v, "nombre": nom}
            almacen.registrar_actividad(usuario, MODULO, "Versión de catálogo", f"{nom} v{v}")
            mensajes.append(f"«{nom}»: versión {v} del catálogo nº {origen}.")
        else:
            id_ = almacen.guardar_catalogo(usuario, ficha, datos)
            o.origen_catalogo = {"id": id_, "version": 1, "nombre": nom}
            almacen.registrar_actividad(usuario, MODULO, "Alta en catálogo", nom)
            mensajes.append(f"«{nom}»: guardada en el catálogo (nº {id_}).")
    return mensajes, errores


def cargar_del_catalogo(e: mo.Estudio, ids: list, hoy=None, usuario: str = None) -> list:
    """Añade al análisis una copia de cada entrada (con el gestor vacío). Devuelve los
    avisos."""
    avisos = []
    for id_ in ids:
        r = almacen.catalogo(id_)
        if not r:
            continue
        avisos += catalogo.avisos_carga(r, e, hoy)
        e.ofertas.append(catalogo.desde_catalogo(r))
        if usuario:
            almacen.registrar_actividad(usuario, MODULO, "Carga desde catálogo", r["nombre"])
    return avisos


def _seleccion(df: pd.DataFrame, clave: str, multiple: bool = True) -> list:
    """Filas elegidas en la tabla (índices)."""
    ev = st.dataframe(df, hide_index=True, use_container_width=True, key=clave,
                      on_select="rerun",
                      selection_mode="multi-row" if multiple else "single-row")
    return list(ev.selection.rows) if ev and ev.selection else []


def _cargar_catalogo(e: mo.Estudio):
    with st.expander("Cargar ofertas del catálogo"):
        st.caption("Ofertas de catálogo guardadas por el equipo. Se añade una copia al análisis: "
                   "puedes retocarla sin cambiar el catálogo. El gestor y su email quedan "
                   "vacíos: rellénalos en cada oferta.")
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        todas = almacen.listar_catalogo()
        coms_ = sorted({r["comercializadora"] for r in todas})
        com = c1.selectbox("Comercializadora", ["Todas"] + coms_, key="ao_cat_com")
        modal = c2.selectbox("Modalidad", ["Todas"] + mo.MODALIDADES, key="ao_cat_mod")
        fecha = e.inicio or dt.date.today()
        vigentes = c3.checkbox("Solo vigentes", True, key="ao_cat_vig",
                               help=f"Que se puedan contratar hoy ({dt.date.today():%d/%m/%Y}) y "
                                    f"cuya ventana de inicio incluya el inicio del análisis "
                                    f"({fecha:%d/%m/%Y}).")
        cubren = c4.checkbox("Solo las que cubren mis tarifas", False, key="ao_cat_cub")
        filas = [r for r in todas
                 if (com == "Todas" or r["comercializadora"] == com)
                 and (modal == "Todas" or r["modalidad"] == modal)
                 and (not vigentes or (catalogo.vigente_en(r, dt.date.today())
                                       and (not e.inicio or not r["inicio_desde"]
                                            or _fecha_iso(r["inicio_desde"]) <= e.inicio
                                            <= _fecha_iso(r["inicio_hasta"]))))
                 and (not cubren or catalogo.cubre(r, e))]
        if not filas:
            st.info("No hay ofertas de catálogo con esos filtros." if todas else
                    "El catálogo está vacío. Guarda una oferta con «Guardar en el catálogo».")
            return
        nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
        elegidas = _seleccion(tabla_catalogo(filas, e, nombres), f"ao_cat_tabla_{_v()}")
        st.caption("Marca una o varias filas (por ejemplo, las de 12, 24 y 36 meses).")
        if st.button("Cargar las elegidas", type="primary", disabled=not elegidas,
                     key="ao_cat_cargar"):
            avisos = cargar_del_catalogo(e, [filas[k]["id"] for k in elegidas],
                                         usuario=_usuario())
            st.session_state.ao_oferta = len(e.ofertas) - 1
            st.session_state.ao_avisos_catalogo = avisos
            _subir_version()
            st.rerun()


def _fecha_iso(v):
    return dt.date.fromisoformat(str(v)[:10])


def _origen_catalogo(e: mo.Estudio, o: mo.Oferta, k: str):
    """Etiqueta de la oferta que viene del catálogo y aviso si hay una versión más nueva."""
    origen = o.origen_catalogo or {}
    if not origen.get("id"):
        return
    try:
        v = almacen.versiones_catalogo([origen["id"]]).get(origen["id"])
    except Exception:                                       # noqa: BLE001
        v = None
    texto = f"Del catálogo: «{origen.get('nombre')}» (nº {origen['id']}, versión " \
            f"{origen.get('version') or 1})"
    if v is None:
        st.caption(texto + " · ya no está en el catálogo.")
    elif v > int(origen.get("version") or 1):
        c1, c2 = st.columns([3, 1])
        c1.warning(texto + f". En el catálogo ya está la versión {v}.")
        if c2.button("Actualizar desde catálogo", key=f"{k}_catup",
                     help="Cambia la fórmula, los precios de las tarifas y las condiciones por "
                          "los del catálogo. Conserva el nombre, el gestor, la ronda, el inicio, "
                          "los grupos propios, los CUPS y los cierres de este análisis."):
            catalogo.actualizar_desde(o, almacen.catalogo(origen["id"]))
            _subir_version()
            st.rerun()
    else:
        st.caption(texto + ".")


def _guardar_catalogo(e: mo.Estudio, o: mo.Oferta, i: int):
    with st.expander("Guardar en el catálogo de ofertas"):
        st.caption("Para ofertas de catálogo (iguales para cualquier cliente). Se guarda la "
                   f"fórmula, los precios por tarifa y las condiciones; no se guarda {NO_SE_GUARDA}.")
        k = f"ao_catg_{i}_{_v()}"
        misma = [n for n, x in enumerate(e.ofertas)
                 if x.comercializadora == o.comercializadora and n != i]
        otras = st.multiselect(
            "Guardar también (una entrada por oferta)", misma, key=k + "otras",
            format_func=lambda n: f"{e.ofertas[n].nombre} · {e.ofertas[n].duracion} M",
            help="Por ejemplo, las de 24 y 36 meses del mismo catálogo. El nombre de cada entrada "
                 "lleva su duración.")
        ofertas = [o] + [e.ofertas[n] for n in otras]
        hoy = dt.date.today()
        corto = (o.comercializadora or "").split(" ")[0].replace(",", "").upper()
        propuesto = (o.origen_catalogo or {}).get("nombre") or             f"{corto} {o.producto or o.modalidad} {hoy:%m/%Y}".strip()
        nombre = st.text_input("Nombre del catálogo o producto", propuesto,
                               key=k + "nombre", help=AY["cat_nombre"])
        c1, c2, c3, c4 = st.columns(4)
        fechas = {
            "vigencia_desde": c1.date_input("Vigencia desde", hoy, format="DD/MM/YYYY",
                                            key=k + "vd", help=AY["cat_vigencia"]),
            "vigencia_hasta": c2.date_input("Vigencia hasta", hoy + dt.timedelta(days=30),
                                            format="DD/MM/YYYY", key=k + "vh",
                                            help=AY["cat_vigencia"]),
            "inicio_desde": c3.date_input("Inicio de suministro desde", e.inicio or hoy,
                                          format="DD/MM/YYYY", key=k + "id",
                                          help=AY["cat_inicio"]),
            "inicio_hasta": c4.date_input("Inicio de suministro hasta",
                                          mo.sumar_meses((e.inicio or hoy).replace(day=1), 3),
                                          format="DD/MM/YYYY", key=k + "ih",
                                          help=AY["cat_inicio"])}
        como_version = False
        origenes = [x for x in ofertas if (x.origen_catalogo or {}).get("id")]
        if origenes:
            como_version = st.radio(
                "Las que ya vienen del catálogo", ["Guardar como versión nueva de su entrada",
                                                   "Guardar como entradas nuevas"],
                key=k + "ver", help="Solo su autor o un administrador puede crear una versión "
                                    "nueva; si no, se guarda como entrada nueva."
            ).startswith("Guardar como versión")
        if st.button("Guardar en el catálogo", type="primary", key=k + "btn"):
            mensajes, errores = guardar_en_catalogo(
                e, ofertas, nombre, fechas, _usuario(), st.session_state.get("ao_id"),
                como_version)
            for m in mensajes:
                st.success(m)
            for x in errores:
                st.error(x)


# =================================================================== página del catálogo
def pagina_catalogo():
    todas = almacen.listar_catalogo(incluir_bajas=True)
    if not todas:
        st.info("El catálogo está vacío. Las ofertas se guardan desde «Nuevo análisis» → "
                "Ofertas → «Guardar en el catálogo de ofertas».")
        return
    nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
    c1, c2, c3 = st.columns([2, 1, 1])
    coms_ = sorted({r["comercializadora"] for r in todas})
    com = c1.selectbox("Comercializadora", ["Todas"] + coms_, key="cat_com")
    modal = c2.selectbox("Modalidad", ["Todas"] + mo.MODALIDADES, key="cat_mod")
    caducadas = c3.checkbox("Ver ofertas caducadas", False, key="cat_caducadas",
                            help="Por defecto solo salen las vigentes hoy. Marca para ver "
                                 "también las caducadas y las dadas de baja.")
    hoy = dt.date.today()
    filas = [r for r in todas
             if (com == "Todas" or r["comercializadora"] == com)
             and (modal == "Todas" or r["modalidad"] == modal)
             and (caducadas or catalogo.vigente_en(r, hoy))]
    if not filas:
        st.info("No hay ofertas de catálogo con esos filtros.")
        return
    elegidas = _seleccion(tabla_catalogo(filas, nombres=nombres), "cat_tabla", multiple=False)
    if not elegidas:
        st.caption("Elige una fila para ver su detalle.")
        return
    r = almacen.catalogo(filas[elegidas[0]]["id"])
    o = catalogo.desde_catalogo(r)
    st.subheader(f"{r['nombre']} · versión {r['version']}")
    st.caption(f"{r['comercializadora']} · {r['modalidad']} · {r['duracion']} meses · alta "
               f"{_fecha_hora(r['fecha'])} por {nombres.get(r['usuario'], r['usuario'])}"
               + (f" · último cambio {_fecha_hora(r['actualizado'])} por "
                  f"{nombres.get(r['actualizado_por'], r['actualizado_por'])}"
                  if r.get("actualizado") else ""))
    if not catalogo.vigente_en(r, hoy):
        st.warning("No vigente hoy." if r["estado"] != catalogo.BAJA else "Dada de baja.")
    formula, leyenda = ayudas.formula_latex(o)
    with st.container(border=True):
        st.latex(formula)
        st.caption("  \n".join(leyenda))
    precios = [{"Grupo": g, "Año": n, "Concepto": c,
                **{p: v for p, v in zip(PER, getattr(pa, campo))}}
               for g, anios in o.precios.items() for n, pa in enumerate(anios, 1)
               for campo, c in (("energia", "Energía"), ("termino_b", "Término B"),
                                ("ci", "Ci"), ("potencia", "Potencia"))
               if any(v not in (None, 1.0) if campo == "ci" else v is not None
                      for v in getattr(pa, campo))]
    if precios:
        st.markdown("**Precios**")
        estilo.tabla_es(pd.DataFrame(precios), hide_index=True, use_container_width=True)
    condiciones = {k: v for k, v in (o.condiciones or {}).items() if v}
    if condiciones or o.notas:
        st.markdown("**Condiciones**")
        for k, v in condiciones.items():
            st.markdown(f"- {k.capitalize()}: {v}")
        if o.notas:
            st.markdown(f"- Notas: {o.notas}")
    usuario = _usuario()
    if not (usuario == r["usuario"] or _es_admin()):
        st.caption("Solo su autor o un administrador puede cambiarla.")
        return
    with st.expander("Cambiar nombre, fechas o estado"):
        k = f"cat_ed_{r['id']}_{r['version']}"
        nombre = st.text_input("Nombre", r["nombre"], key=k + "n")
        c1, c2, c3, c4 = st.columns(4)
        ficha = dict(r, nombre=nombre,
                     vigencia_desde=c1.date_input("Vigencia desde", _fecha_iso(r["vigencia_desde"]),
                                                  format="DD/MM/YYYY", key=k + "vd"),
                     vigencia_hasta=c2.date_input("Vigencia hasta", _fecha_iso(r["vigencia_hasta"]),
                                                  format="DD/MM/YYYY", key=k + "vh"),
                     inicio_desde=c3.date_input("Inicio desde", _fecha_iso(r["inicio_desde"]),
                                                format="DD/MM/YYYY", key=k + "id"),
                     inicio_hasta=c4.date_input("Inicio hasta", _fecha_iso(r["inicio_hasta"]),
                                                format="DD/MM/YYYY", key=k + "ih"))
        b1, b2, b3 = st.columns(3)
        if b1.button("Guardar cambios", type="primary", key=k + "g"):
            errores = catalogo.validar(ficha, r["datos"])
            if errores:
                for x in errores:
                    st.error(x)
            elif almacen.actualizar_catalogo(r["id"], usuario, ficha):
                almacen.registrar_actividad(usuario, MODULO, "Cambio en catálogo", nombre)
                st.rerun()
        baja = r["estado"] == catalogo.BAJA
        if b2.button("Reactivar" if baja else "Dar de baja", key=k + "b",
                     help="Una oferta de baja no sale en el buscador del análisis."):
            almacen.estado_catalogo(r["id"], usuario, catalogo.ACTIVA if baja else catalogo.BAJA)
            almacen.registrar_actividad(usuario, MODULO,
                                        "Alta en catálogo" if baja else "Baja de catálogo",
                                        r["nombre"])
            st.rerun()
        seguro = b3.checkbox("Confirmo que quiero borrarla", key=k + "s")
        if b3.button("Borrar", disabled=not seguro, key=k + "x",
                     help="Se borra del catálogo; las ofertas ya cargadas en análisis no cambian."):
            almacen.borrar_catalogo(r["id"], usuario)
            almacen.registrar_actividad(usuario, MODULO, "Borrado de catálogo", r["nombre"])
            st.rerun()
    st.caption("Para cambiar la fórmula o los precios, cárgala en un análisis, modifícala y "
               "guárdala con «Guardar como versión nueva de su entrada».")


# =================================================================== 6-7. controles y resultados
def _controles(e: mo.Estudio, res: dict):
    estilo.seccion(6, "Controles")
    lista = servicio.controles(e, res)
    errores = [t for n, t in lista if n == "error"]
    avisos = [t for n, t in lista if n == "aviso"] + list(st.session_state.get("ao_avisos") or [])
    if not lista and not avisos:
        st.success("Sin incidencias.")
    if errores:
        st.error("**Revisa antes de usar los resultados:**\n" + "\n".join(f"- {t}"
                                                                          for t in errores))
    if avisos:
        with st.expander(f"Avisos ({len(avisos)})", expanded=not errores):
            st.markdown("\n".join(f"- {t}" for t in avisos))


def _resultados(e: mo.Estudio, res: dict):
    estilo.seccion(7, "Resultados", "Coste medio anual en la duración de cada oferta")
    filas = resultados.comparativa(res)
    if not filas:
        st.info("Añade ofertas para ver la comparativa.")
        return filas
    actual = res.get("actual")
    if actual:
        p = actual.por_anio()[0]
        c1, c2, c3 = st.columns(3)
        c1.metric("Contrato actual (primer año)", f"{informes.eur(p['total'])} €/año")
        c2.metric("€/kWh", informes.ekwh(p["total"] / p["consumo"] if p["consumo"] else None))
        c3.metric("Consumo anual", f"{informes.eur(p['consumo'] / 1000)} MWh")
    st.caption("€/año: coste medio anual en la duración de cada oferta. " + AY["mismo_periodo"])
    estilo.tabla_es(tabla_comparativa(filas), hide_index=True, use_container_width=True,
                    formatos={"€/año": 0, "Actual mismo periodo (€/año)": 0,
                              "Ahorro (€/año)": 0, "€/kWh": 4,
                              "Variación": lambda v: numeros.porcentaje(v, 2, signo=True)})
    for t, f in resultados.mejores(filas).items():
        st.markdown(f"**Mejor oferta ({t.lower()}):** {f['comercializadora']} "
                    f"{f['duracion']} meses · {informes.eur(f['media_anual'])} €/año "
                    f"({informes.pct(f['variacion'])} frente al contrato actual)")
    t_merc, t_esc, t_ap, t_rondas, t_cups = st.tabs(
        ["Mercado previsto", "Escenarios OMIP", "Apuntamiento estimado / cliente", "Rondas",
         "Coste por CUPS"])
    with t_merc:
        estilo.tabla_es(pd.DataFrame(resultados.cuadro_mercado(e)).drop(columns=["anio", "q"]),
                        hide_index=True)
    with t_esc:
        pasos = st.slider("Variación de OMIP (±%)", 5, 50, 20, 5, key="ao_esc")
        if st.button("Calcular escenarios", help=AY["escenarios"]):
            esc = resultados.escenarios(e, (-pasos / 100, 0.0, pasos / 100))
            tabla = pd.DataFrame([{"Oferta": f["nombre"], "Comercializadora": f["comercializadora"],
                                   **{f"OMIP {x['variacion_omip'] * 100:+.0f} % (€/año)":
                                      x["filas"][i]["media_anual"] for x in esc},
                                   **{f"Variación con OMIP {x['variacion_omip'] * 100:+.0f} %":
                                      (x["filas"][i]["variacion"] or 0) * 100 for x in esc}}
                                  for i, f in enumerate(esc[1]["filas"])])
            estilo.tabla_es(tabla, hide_index=True, use_container_width=True, formatos={
                c: (0 if "€/año" in c else (lambda v: numeros.porcentaje(v, 2, signo=True)))
                for c in tabla.columns if c.startswith(("OMIP", "Variación"))})
    with t_ap:
        comp = resultados.apuntamientos(e)
        if not comp:
            st.caption("Ningún CUPS tiene curva de un año completo (sección 3, «Curvas del "
                       "cliente»).")
        else:
            estilo.tabla_es(pd.DataFrame(comp), hide_index=True, use_container_width=True)
    with t_rondas:
        rondas = resultados.rondas(filas)
        if not rondas:
            st.caption("Ninguna comercializadora tiene más de una ronda.")
        for r in rondas:
            st.markdown(f"**{r['comercializadora']}**")
            estilo.tabla_es(pd.DataFrame([{"Ronda": f["ronda"], "Meses": f["duracion"],
                                           "€/año": f["media_anual"],
                                           "Variación": (f["variacion"] or 0) * 100}
                                          for f in r["rondas"]]), hide_index=True,
                            formatos={"€/año": 0, "Variación": lambda v: numeros.porcentaje(
                                v, 2, signo=True)})
    with t_cups:
        cpc = informes.coste_por_cups(res)
        estilo.tabla_es(cpc, hide_index=True, use_container_width=True,
                        formatos={c: 0 for c in cpc.columns if c != "CUPS"})
    return filas


def _ronda_del_informe(res: dict) -> tuple:
    """(resultado con las ofertas del informe, ronda o None, mejora frente a la ronda anterior o
    None). Con ofertas de más de una ronda se elige la ronda (se cargan sus ofertas y se pueden
    añadir o quitar otras) y si se añade la fila de mejora frente a la ronda anterior."""
    rondas_ = resultados.rondas_con_ofertas(res)
    if len(rondas_) < 2:
        return res, None, None
    todas = [r.oferta.nombre for r in res["ofertas"]]
    ronda_de = {r.oferta.nombre: int(r.oferta.ronda or 1) for r in res["ofertas"]}
    c1, c2 = st.columns([1, 3])
    opciones = rondas_ + ["Todas"]
    ronda = c1.selectbox("Ronda del informe", opciones, index=len(rondas_) - 1,
                         key="ao_inf_ronda",
                         format_func=lambda r: "Todas las ofertas" if r == "Todas" else f"{r}ª ronda",
                         help="El PDF y el Excel salen solo con las ofertas de esa ronda (puedes "
                              "añadir o quitar otras al lado). El título y el nombre del fichero "
                              "llevan la ronda.")
    defecto = todas if ronda == "Todas" else resultados.ofertas_de_ronda(res, ronda)
    elegidas = c2.multiselect("Ofertas que entran en el informe", todas, default=defecto,
                              key=f"ao_inf_ofertas_{ronda}",
                              format_func=lambda n: f"{n} · {ronda_de[n]}ª ronda",
                              help="Se cargan las de la ronda elegida; añade o quita las que "
                                   "necesites.")
    if not elegidas:
        st.warning("Elige al menos una oferta para el informe.")
    mejora = None
    if ronda != "Todas" and ronda > 1:
        if st.checkbox("Añadir la fila «Mejora frente a la ronda anterior»",
                       key=f"ao_inf_mejora_{ronda}",
                       help="Una fila más en el resumen con lo que ha mejorado el coste del primer "
                            "año de cada oferta frente a la de la misma comercializadora en la "
                            "ronda anterior (misma modalidad y duración si la hay)."):
            mejora = resultados.mejora_ronda_anterior(res, elegidas)
    return resultados.solo_ofertas(res, elegidas), (None if ronda == "Todas" else ronda), mejora


def _guardar_y_descargas(e: mo.Estudio, res: dict, filas: list):
    st.markdown("#### Informes y correo")
    ref = st.text_input("Referencia del estudio", st.session_state.get("ao_ref", ""),
                        key="ao_ref_txt", help=AY["referencia"])
    st.session_state.ao_ref = ref
    res_inf, ronda_inf, mejora = _ronda_del_informe(res)
    if st.button("Preparar informes (Excel y PDF) y borrador del correo",
                 disabled=not res_inf["ofertas"]):
        with st.spinner("Generando…"):
            from analisis_ofertas import informe_pdf
            tecnico = informe_pdf.tecnico_de(_usuario())
            st.session_state.ao_excel = informes.excel(e, res_inf, ref, tecnico, ronda=ronda_inf,
                                                       mejora=mejora)
            st.session_state.ao_pdf = informes.pdf(e, res_inf, ref, tecnico, ronda=ronda_inf,
                                                   mejora=mejora)
            st.session_state.ao_correo = correo.borrador(e, res_inf, ref)
            st.session_state.ao_inf_sufijo = f"_ronda{ronda_inf}" if ronda_inf else ""
    nombre = re.sub(r"[^\w-]", "_", ref or e.cliente or "analisis")[:40] + \
        st.session_state.get("ao_inf_sufijo", "")
    if st.session_state.get("ao_excel"):
        c1, c2, c3 = st.columns(3)
        estilo.boton_descarga(MODULO, c1, "Informe Excel", st.session_state.ao_excel,
                              f"estudio_ofertas_{nombre}.xlsx", XLSX, type="primary")
        estilo.boton_descarga(MODULO, c2, "Informe PDF", st.session_state.ao_pdf,
                              f"estudio_ofertas_{nombre}.pdf", "application/pdf")
        b = st.session_state.ao_correo
        estilo.boton_descarga(MODULO, c3, "Correo (HTML)", b["html"].encode("utf-8"),
                              f"correo_{nombre}.html", "text/html")
        st.text_input("Asunto", b["asunto"], key="ao_asunto")
        st.text_area("Borrador del correo (edítalo antes de enviarlo)", b["texto"], height=360,
                     key="ao_texto_correo")
    st.markdown("#### Guardar")
    estado = st.selectbox("Estado", servicio.ESTADOS, key="ao_estado_sel", help=AY["estado"],
                          index=servicio.ESTADOS.index(st.session_state.get("ao_estado",
                                                                           "En curso")))
    adjudicada = None
    if estado == "Adjudicada" and filas:
        por_nombre = {f["nombre"]: f for f in filas}
        adjudicada = st.selectbox("Oferta adjudicada", list(por_nombre), key="ao_adjudicada",
                                  help=AY["adjudicada"],
                                  format_func=lambda n: f"{por_nombre[n]['comercializadora']} "
                                                        f"{por_nombre[n]['duracion']} m ({n})")
    c1, c2 = st.columns(2)
    id_actual = st.session_state.get("ao_id")
    if c1.button("Guardar el análisis" if not id_actual else f"Guardar los cambios (nº {id_actual})",
                 type="primary"):
        resumen = servicio.resumen(e, res, estado, ref)
        if adjudicada:
            resumen["mejor"] = f"Adjudicada: {adjudicada}"
        datos = servicio.datos(e, ref)
        curvas = servicio.empaquetar_curvas(st.session_state.get("ao_curvas"))
        excel = st.session_state.get("ao_excel")
        pdf = st.session_state.get("ao_pdf")
        if id_actual:
            almacen.actualizar_analisis(id_actual, _usuario(), resumen, datos, curvas, excel, pdf)
        else:
            id_actual = almacen.guardar_analisis(_usuario(), resumen, datos, curvas, excel, pdf)
            st.session_state.ao_id = id_actual
        st.session_state.ao_estado = estado
        if adjudicada and e.id_solicitud:
            com = next((o.comercializadora for o in e.ofertas if o.nombre == adjudicada), "")
            servicio.adjudicar_solicitud(e.id_solicitud, com, _usuario())
        st.session_state.pop("ao_contactos", None)
        st.success(f"Análisis guardado con el nº {id_actual}."
                   + (" La solicitud queda «Adjudicada»." if adjudicada and e.id_solicitud else ""))
    if c2.button("Empezar un análisis nuevo"):
        _limpiar()
        st.rerun()


# =================================================================== páginas
def pagina_nuevo_analisis():
    e = _estudio()
    if e is None:
        st.markdown("Valora las **ofertas de electricidad recibidas** frente al contrato actual "
                    "con la metodología de la plantilla multipunto de GE&PE: fijo, indexado "
                    "Pass-Pool y Pass-Through, cierres en OMIP, todos los modos de SSAA, "
                    "cualquier duración y rondas.")
        _origen()
        return
    if st.session_state.get("ao_id"):
        st.info(f"Análisis nº {st.session_state.ao_id} cargado. «Guardar los cambios» lo "
                "actualiza (queda quién y cuándo).")
    # meses del horizonte sin precio de mercado (una oferta más larga) y SSAA vacíos: se
    # rellenan con lo que hay en Parámetros
    completados = servicio.completar_mercado(e)
    if servicio.ssaa_por_defecto(e):
        completados.append("SSAA vacíos: se han puesto los SSAA estimados de Parámetros "
                           "(se pueden cambiar en la sección 3).")
    if completados:
        st.session_state.ao_avisos = (st.session_state.get("ao_avisos") or []) + completados
        _subir_version()
    _suministros(e)
    _mercado(e)
    _contrato_actual(e)
    _ofertas(e)
    try:
        res = calculo.calcular(e)
    except Exception as err:                                # noqa: BLE001
        st.error(f"No se ha podido calcular: {err}")
        return
    _controles(e, res)
    filas = _resultados(e, res)
    _guardar_y_descargas(e, res, filas)


def pagina_analisis_guardados():
    lista = almacen.listar_analisis()
    if not lista:
        st.info("Todavía no hay análisis guardados. Se guardan al final de «Nuevo análisis».")
        return
    nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
    c1, c2 = st.columns([2, 1])
    buscar = c1.text_input("Buscar por cliente, CUPS, Ref. o comercializadora").strip().lower()
    estados = c2.multiselect("Estado", servicio.ESTADOS)
    filtrados = [a for a in lista if (not estados or a["estado"] in estados)
                 and (not buscar or buscar in " ".join(str(a.get(k) or "") for k in (
                     "cliente", "cups", "ref", "comercializadoras")).lower())]
    estilo.tabla_es(pd.DataFrame([{
        "Nº": a["id"], "Guardado": _fecha_hora(a["fecha"]),
        "Usuario": nombres.get(a["usuario"], a["usuario"]), "Ref.": a["ref"],
        "Cliente": a["cliente"], "Suministros": a["suministros"], "Ofertas": a["ofertas"],
        "Ronda": a["ronda"], "Mejor": a["mejor"], "Ahorro (€/año)": a["ahorro"],
        "Estado": a["estado"], "Solicitud": a["id_solicitud"],
        "Actualizado": _fecha_hora(a["actualizado"])} for a in filtrados]),
        hide_index=True, use_container_width=True, formatos={"Ahorro (€/año)": 0})
    estilo.borrar_mios(MODULO, filtrados, lambda a: "Nº %d · %s · %s" % (
        a["id"], _fecha_hora(a["fecha"]), a["cliente"] or a["ref"] or "—"),
        almacen.borrar_analisis, "ao_borrar", que="análisis")
    if not filtrados:
        return
    por_id = {a["id"]: a for a in filtrados}
    elegido = st.selectbox("Análisis", list(por_id), format_func=lambda i: (
        f"Nº {i} · {por_id[i]['cliente'] or '—'}"
        + (f" · Ref. {por_id[i]['ref']}" if por_id[i]["ref"] else "")))
    an = almacen.analisis(elegido)
    e = servicio.estudio_de_datos(an["datos"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Ofertas", an["ofertas"] or 0)
    c2.metric("Estado", an["estado"])
    c3.metric("Mejor", an["mejor"] or "—")
    contactos = [{"Comercializadora": o.comercializadora, "Ronda": o.ronda, "Gestor": o.gestor,
                  "Email": o.email_gestor, "Meses": o.duracion} for o in e.ofertas]
    if contactos:
        with st.expander("Ofertas y contactos de las comercializadoras"):
            st.dataframe(pd.DataFrame(contactos), hide_index=True, use_container_width=True)
    with st.expander("Parámetros utilizados (congelados con el análisis)"):
        tablas = parametros.tablas_estudio(e)
        for nombre, df in tablas.items():
            st.markdown(f"**{nombre}**")
            estilo.tabla_es(df, hide_index=True, use_container_width=True)
        estilo.boton_descarga(MODULO, st, "Descargar parámetros (Excel)", parametros.excel(tablas),
                              f"parametros_analisis_{an['id']}.xlsx", XLSX)
    b1, b2, b3 = st.columns(3)
    estilo.boton_descarga(MODULO, b1, "Informe Excel", an["excel"] or b"",
                          f"estudio_ofertas_{an['id']}.xlsx", XLSX, disabled=not an["excel"])
    estilo.boton_descarga(MODULO, b2, "Informe PDF", an["pdf"] or b"",
                          f"estudio_ofertas_{an['id']}.pdf", "application/pdf",
                          disabled=not an["pdf"])
    if b3.button("Cargar en «Nuevo análisis»", type="primary"):
        cargar_analisis(an["id"])
        if PAGINA_FORMULARIO is not None:
            st.switch_page(PAGINA_FORMULARIO)
        st.success("Cargado: ve a «Nuevo análisis».")


def pagina_parametros():
    st.markdown("Todos los valores con los que se valoran las ofertas hoy, con su **fuente y "
                "fecha**: términos regulados (Administración → Datos regulados), mercado (OMIP "
                "de la última sesión y medias de OMIE) y apuntamientos calculados por la "
                "aplicación con los 12 últimos meses publicados. Cada análisis guardado conserva "
                "los suyos (en «Análisis guardados»).")
    calc = _apuntamientos_al_dia()
    tablas = parametros.tablas_vigentes(calculados=calc)
    for a in tablas.get("_avisos", []):
        st.caption(a)
    estilo.boton_descarga(MODULO, st, "Descargar todo (Excel)", parametros.excel(tablas),
                          f"parametros_utilizados_{dt.date.today():%Y%m%d}.xlsx", XLSX,
                          type="primary")
    nombres = [n for n in tablas if not n.startswith("_")]
    for pestaña, nombre in zip(st.tabs(nombres), nombres):
        with pestaña:
            df = tablas[nombre]
            if nombre == "Regulados por tarifa" and not df.empty:
                c1, c2 = st.columns(2)
                anio = c1.selectbox("Año", sorted(df["Año"].unique()), key="pu_anio")
                tarifa = c2.selectbox("Tarifa", mo.TARIFAS, key="pu_tarifa")
                df = df[(df["Año"] == anio) & (df["Tarifa"] == tarifa)]
            if nombre == "Carga pico":
                st.caption(parametros.EXPLICACION_PICO)
            estilo.tabla_es(df, hide_index=True, use_container_width=True)
