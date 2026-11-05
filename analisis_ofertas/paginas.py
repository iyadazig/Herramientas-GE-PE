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
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo, correo, informes, parametros, resultados, servicio
from analisis_ofertas import comercializadoras as coms
from analisis_ofertas import modelo as mo

MODULO = "Análisis de ofertas"
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
    editada = st.data_editor(bases[clave], key=clave, column_config=config, **kw).copy()
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
                    key="ao_modo_origen")
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
                         help=f"Año 2, año 3… separados por comas (el horizonte tiene {n_anios} "
                              "años). Vacío = mismo consumo todos los años.")
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
    if c2.button("Actualizar mercado", help="Vuelve a leer la última sesión de OMIP y las medias "
                                           "de OMIE para todos los meses del horizonte."):
        st.session_state.ao_avisos = servicio.congelar_mercado(e)
        _subir_version()
        st.rerun()
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
        st.caption("SSAA medios de cada mes natural (€/MWh), a mano en cada estudio.")
        st.caption("Se pueden pegar los 12 valores copiados de Excel (con coma decimal).")
        aplicar_ssaa(e, editor(tabla_ssaa(e), f"ao_ssaa_{_v()}", ["SSAA (€/MWh)"],
                               {"SSAA (€/MWh)": 4}, hide_index=True, disabled=["Mes"]))
        if any(s.activo and s.sistema == "Canarias" for s in e.suministros):
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
def _contactos() -> dict:
    if "ao_contactos" not in st.session_state:
        try:
            st.session_state.ao_contactos = almacen.contactos_comercializadoras()
        except Exception:                                   # noqa: BLE001
            st.session_state.ao_contactos = {}
    return st.session_state.ao_contactos


def _sel(contenedor, etiqueta, opciones, valor, clave, **kw):
    return contenedor.selectbox(etiqueta, opciones, key=clave, index=opciones.index(valor)
                                if valor in opciones else 0, **kw)


def formulario_oferta(e: mo.Estudio, o: mo.Oferta, k: str):
    """Ficha de una oferta (o del contrato actual) en pestañas. k: prefijo de las claves."""
    k = f"ao_{k}_{_v()}"
    t_conf, t_precios, t_grupos, t_cierres, t_cond = st.tabs(
        ["Configuración", "Precios", "Grupos y CUPS", "Cierres OMIP", "Condiciones"])
    with t_conf:
        c1, c2, c3 = st.columns(3)
        opciones = coms.opciones(o.comercializadora)
        o.comercializadora = c1.selectbox(
            "Comercializadora", opciones, key=k + "com",
            index=opciones.index(o.comercializadora) if o.comercializadora in opciones else 0,
            format_func=lambda c: c or "— Elige —", help="Escribe para buscar en la lista.")
        contacto = _contactos().get(o.comercializadora.upper()) if o.comercializadora else None
        if contacto and not (o.gestor or o.email_gestor):
            o.gestor, o.email_gestor = contacto["gestor"], contacto["email"]
        o.gestor = c2.text_input("Gestor de la comercializadora", o.gestor, key=k + "gestor")
        o.email_gestor = c3.text_input("Email del gestor", o.email_gestor, key=k + "email").strip()
        if o.email_gestor and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", o.email_gestor):
            c3.caption(":red[Email no válido]")
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        o.producto = c1.text_input("Producto / descripción", o.producto, key=k + "prod")
        o.modalidad = _sel(c2, "Modalidad", mo.MODALIDADES, o.modalidad, k + "mod")
        o.duracion = int(c3.number_input("Duración (meses)", 1, 240, int(o.duracion or 12),
                                         key=k + "dur"))
        if not o.es_actual:
            o.ronda = int(c4.number_input("Ronda", 1, 20, int(o.ronda or 1), key=k + "ronda"))
        c1, c2, c3, c4 = st.columns(4)
        propio = c1.checkbox("Inicio distinto al del estudio", bool(o.inicio), key=k + "inip")
        if propio:
            ini = c1.date_input("Inicio", o.inicio or e.inicio, format="DD/MM/YYYY",
                                key=k + "ini")
            o.inicio = ini.replace(day=1) if ini else None
        else:
            o.inicio = None
        if o.modalidad == mo.FIJO:
            o.unidad = _sel(c2, "Unidad de la energía", list(mo.UNIDADES), o.unidad, k + "uni")
            o.descuento = c3.number_input("Descuento sobre energía (%)", 0.0, 100.0,
                                          float(o.descuento * 100), key=k + "desc") / 100
        o.atr_incluido = c4.checkbox("El precio incluye peajes y cargos", o.atr_incluido,
                                     key=k + "atr")
        o.potencia = _sel(c2 if o.modalidad != mo.FIJO else c1, "Término de potencia",
                          [mo.POT_REGULADA, mo.POT_PROPIA], o.potencia, k + "pot")
        if o.modalidad != mo.FIJO:
            st.markdown("**Componentes de mercado (indexados)**")
            c1, c2, c3, c4, c5 = st.columns(5)
            o.pagos_capacidad = c1.checkbox("Pagos por capacidad", o.pagos_capacidad,
                                            key=k + "pc")
            o.os = c2.checkbox("Pago OS", o.os, key=k + "os")
            o.om = c3.checkbox("Pago OM", o.om, key=k + "om")
            o.srad = c4.checkbox("SRAD", o.srad, key=k + "srad")
            o.base_b = _sel(c5, "Base del término B", mo.BASES, o.base_b, k + "bb")
            c1, c2 = st.columns(2)
            o.desvios = c1.number_input("Desvíos (€/MWh, en barras)", value=float(o.desvios),
                                        key=k + "desv", help=referencia_desvios(e) or None)
            o.indice_ph = c2.checkbox(
                "Canarias indexado al Ph demanda", o.indice_ph, key=k + "ph",
                help="Precio horario de adquisición de la demanda en Canarias (Orden "
                     "TEC/1172/2018, ESIOS). Ya incluye los costes de ajuste: en los CUPS de "
                     "Canarias el mercado es (previsto + Ph − diario) × apuntamiento del Ph. "
                     "No añadas SSAA si la oferta no los cobra aparte.")
            if o.modalidad == mo.PASS_POOL:
                o.ci_media_mensual = c2.checkbox(
                    "El Ci se aplica a la media mensual de OMIE", o.ci_media_mensual,
                    key=k + "cimm",
                    help="Precio de mercado mensual × Ci, con el Ci que ya incluye el perfil "
                         "(y las pérdidas): no se aplica el apuntamiento. Desmárcalo si la "
                         "oferta aplica el Ci sobre el precio apuntado.")
        st.markdown("**Otros conceptos**")
        c1, c2, c3, c4 = st.columns(4)
        o.fee = c1.number_input("Fee / gestión (€/MWh)", value=float(o.fee), key=k + "fee")
        o.base_fee = _sel(c2, "Base del fee", mo.BASES, o.base_fee, k + "bfee")
        o.gdo = c3.number_input("Garantías de origen (€/MWh)", value=float(o.gdo), key=k + "gdo")
        o.fnee = _sel(c4, "FNEE", mo.MODOS_FNEE, o.fnee, k + "fnee")
        c1, c2, c3 = st.columns(3)
        o.k = c1.number_input("K / impuesto municipal (%)", value=float(o.k * 100),
                              key=k + "k") / 100
        o.k_ssaa = c2.checkbox("La K se aplica a los SSAA", o.k_ssaa, key=k + "kss")
        o.otros_fijos = c3.number_input("Otros costes fijos (€/CUPS·mes)",
                                        value=float(o.otros_fijos), key=k + "otf")
        st.markdown("**Servicios de ajuste (SSAA)**")
        c1, c2, c3, c4 = st.columns(4)
        o.ssaa_modo = _sel(c1, "Modo", mo.MODOS_SSAA, o.ssaa_modo, k + "ssm")
        if o.ssaa_modo == mo.SSAA_BANDA:
            o.banda_min = c2.number_input("Banda mínima (€/MWh)", value=float(o.banda_min),
                                          key=k + "bmin")
            o.banda_max = c3.number_input("Banda máxima (€/MWh)", value=float(o.banda_max),
                                          key=k + "bmax")
        if o.ssaa_modo == mo.SSAA_TECHO:
            o.techo = c2.number_input("Techo (€/MWh)", value=float(o.techo), key=k + "techo")
        if o.ssaa_modo in (mo.SSAA_BANDA, mo.SSAA_TECHO):
            o.ssaa_perfil = c4.checkbox("Aplicar perfil por periodo", o.ssaa_perfil,
                                        key=k + "sperf")
        if o.ssaa_modo != mo.SSAA_INCLUIDOS:
            c1, c2, c3, c4 = st.columns(4)
            o.ssaa_perdidas = _sel(c1, "Elevación a pérdidas", mo.PERDIDAS_SSAA,
                                   o.ssaa_perdidas, k + "sloss")
            if o.ssaa_perdidas == mo.PERD_FIJO:
                o.ssaa_pct_perdidas = c2.number_input("% de pérdidas", value=float(
                    o.ssaa_pct_perdidas * 100), key=k + "slp") / 100
            o.ssaa_f1 = c3.number_input("Factor 1", value=float(o.ssaa_f1), format="%.4f",
                                        key=k + "sf1")
            o.ssaa_f2 = c4.number_input("Factor 2", value=float(o.ssaa_f2), format="%.4f",
                                        key=k + "sf2")
        if o.modalidad != mo.FIJO:
            st.markdown("**Cierres en OMIP**")
            c1, c2 = st.columns(2)
            o.cierres_tipo = _sel(c1, "Tipo de cierres", mo.TIPOS_CIERRE, o.cierres_tipo,
                                  k + "cie")
            if o.cierres_tipo != mo.SIN_CIERRES:
                o.coste_click = c2.number_input("Coste por click (€/MWh cerrado)",
                                                value=float(o.coste_click), key=k + "click")
        else:
            o.cierres_tipo = mo.SIN_CIERRES
    with t_precios:
        unidad = o.unidad if o.modalidad == mo.FIJO else ""
        st.caption("Una fila por grupo (tarifa y sistema, o grupo propio) y año de contrato. "
                   "Los grupos de Baleares, Canarias, Ceuta y Melilla sin precios usan los de su "
                   "tarifa; un año sin precios, los del anterior.")
        for concepto in conceptos_precio(o):
            st.markdown(f"**{CONCEPTOS_PRECIO[concepto]}**" + (f" ({unidad})" if concepto ==
                                                                "energia" else ""))
            tabla = tabla_precios(e, o, concepto)
            editada = editor(tabla, f"{k}pr_{concepto}_{len(tabla)}_{o.modalidad}_{o.unidad}",
                             PER, hide_index=True, use_container_width=True,
                             disabled=["Grupo", "Año"])
            aplicar_precios(o, concepto, editada)
    with t_grupos:
        st.caption("Grupos propios: para dar a algunos CUPS un precio distinto del de su tarifa. "
                   "Después asígnalos en la tabla de CUPS.")
        grupos = editor(tabla_grupos_propios(o), k + "gp", hide_index=True, num_rows="dynamic",
                        column_config={
                                    "Tarifa": st.column_config.SelectboxColumn(options=mo.TARIFAS),
                                    "Sistema": st.column_config.SelectboxColumn(
                                        options=mo.SISTEMAS)})
        aplicar_grupos_propios(o, grupos)
        cups = editor(
            tabla_cups_oferta(e, o), f"{k}cups_{len(o.grupos_propios)}", hide_index=True,
            use_container_width=True, disabled=["CUPS", "Alias", "Tarifa", "Sistema"],
            column_config={"Grupo propio": st.column_config.SelectboxColumn(
                options=[""] + [g.nombre for g in o.grupos_propios])})
        aplicar_cups_oferta(o, cups)
    with t_cierres:
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
    o.nombre = st.text_input("Nombre de la oferta", o.nombre, key=f"ao_of{i}_{_v()}_nombre")
    formulario_oferta(e, o, f"of{i}")


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
    st.dataframe(tabla_comparativa(filas), hide_index=True, use_container_width=True,
                 column_config={c: st.column_config.NumberColumn(format="localized")
                                for c in ("€/año", "Actual mismo periodo (€/año)",
                                          "Ahorro (€/año)")}
                 | {"€/kWh": st.column_config.NumberColumn(format="%.4f"),
                    "Variación": st.column_config.NumberColumn(format="%+.2f %%")})
    for t, f in resultados.mejores(filas).items():
        st.markdown(f"**Oferta mínima ({t.lower()}):** {f['comercializadora']} "
                    f"{f['duracion']} meses · {informes.eur(f['media_anual'])} €/año "
                    f"({informes.pct(f['variacion'])} frente al contrato actual)")
    t_merc, t_esc, t_ap, t_rondas, t_cups = st.tabs(
        ["Mercado previsto", "Escenarios OMIP", "Apuntamiento estimado / cliente", "Rondas",
         "Coste por CUPS"])
    with t_merc:
        st.dataframe(pd.DataFrame(resultados.cuadro_mercado(e)).drop(columns=["anio", "q"]),
                     hide_index=True)
    with t_esc:
        pasos = st.slider("Variación de OMIP (±%)", 5, 50, 20, 5, key="ao_esc")
        if st.button("Calcular escenarios"):
            esc = resultados.escenarios(e, (-pasos / 100, 0.0, pasos / 100))
            tabla = pd.DataFrame([{"Oferta": f["nombre"], "Comercializadora": f["comercializadora"],
                                   **{f"OMIP {x['variacion_omip'] * 100:+.0f} % (€/año)":
                                      x["filas"][i]["media_anual"] for x in esc},
                                   **{f"Variación con OMIP {x['variacion_omip'] * 100:+.0f} %":
                                      (x["filas"][i]["variacion"] or 0) * 100 for x in esc}}
                                  for i, f in enumerate(esc[1]["filas"])])
            st.dataframe(tabla, hide_index=True, use_container_width=True)
    with t_ap:
        comp = resultados.apuntamientos(e)
        if not comp:
            st.caption("Ningún CUPS tiene curva de un año completo (sección 3, «Curvas del "
                       "cliente»).")
        else:
            st.dataframe(pd.DataFrame(comp), hide_index=True, use_container_width=True)
    with t_rondas:
        rondas = resultados.rondas(filas)
        if not rondas:
            st.caption("Ninguna comercializadora tiene más de una ronda.")
        for r in rondas:
            st.markdown(f"**{r['comercializadora']}**")
            st.dataframe(pd.DataFrame([{"Ronda": f["ronda"], "Meses": f["duracion"],
                                        "€/año": f["media_anual"],
                                        "Variación (%)": (f["variacion"] or 0) * 100}
                                       for f in r["rondas"]]), hide_index=True)
    with t_cups:
        st.dataframe(informes.coste_por_cups(res), hide_index=True, use_container_width=True)
    return filas


def _guardar_y_descargas(e: mo.Estudio, res: dict, filas: list):
    st.markdown("#### Informes y correo")
    ref = st.text_input("Referencia del estudio", st.session_state.get("ao_ref", ""),
                        key="ao_ref_txt")
    st.session_state.ao_ref = ref
    if st.button("Preparar informes (Excel y PDF) y borrador del correo"):
        with st.spinner("Generando…"):
            st.session_state.ao_excel = informes.excel(e, res, ref)
            st.session_state.ao_pdf = informes.pdf(e, res, ref)
            st.session_state.ao_correo = correo.borrador(e, res, ref)
    nombre = re.sub(r"[^\w-]", "_", ref or e.cliente or "analisis")[:40]
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
    estado = st.selectbox("Estado", servicio.ESTADOS, key="ao_estado_sel",
                          index=servicio.ESTADOS.index(st.session_state.get("ao_estado",
                                                                           "En curso")))
    adjudicada = None
    if estado == "Adjudicada" and filas:
        por_nombre = {f["nombre"]: f for f in filas}
        adjudicada = st.selectbox("Oferta adjudicada", list(por_nombre), key="ao_adjudicada",
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
    st.dataframe(pd.DataFrame([{
        "Nº": a["id"], "Guardado": _fecha_hora(a["fecha"]),
        "Usuario": nombres.get(a["usuario"], a["usuario"]), "Ref.": a["ref"],
        "Cliente": a["cliente"], "Suministros": a["suministros"], "Ofertas": a["ofertas"],
        "Ronda": a["ronda"], "Mejor": a["mejor"], "Ahorro (€/año)": a["ahorro"],
        "Estado": a["estado"], "Solicitud": a["id_solicitud"],
        "Actualizado": _fecha_hora(a["actualizado"])} for a in filtrados]),
        hide_index=True, use_container_width=True,
        column_config={"Ahorro (€/año)": st.column_config.NumberColumn(format="localized")})
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
            st.dataframe(df, hide_index=True, use_container_width=True)
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
            st.dataframe(df, hide_index=True, use_container_width=True)
