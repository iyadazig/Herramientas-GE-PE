# -*- coding: utf-8 -*-
"""
Administración → Datos regulados
================================
Edición por año de los peajes, cargos, pagos por capacidad, pérdidas, SRAD, impuestos, FNEE,
pagos OS/OM y parámetros de estimación del mercado (datos_regulados). Lo guardado prevalece
sobre regulados/regulados.json. Las tablas se convierten sin Streamlit (probadas en
tests/test_analisis_paginas.py).
"""

import json

import pandas as pd
import streamlit as st

import almacen
import datos_regulados as dr
import estilo
from analisis_ofertas.paginas import editor

MODULO = "Administración"
PERIODOS = ["P1", "P2", "P3", "P4", "P5", "P6"]


def _n_concepto(tarifa: str, clave: str) -> int:
    """Periodos de un concepto (la 2.0TD tiene 2 de potencia y 3 de energía)."""
    if tarifa == "2.0TD":
        return 2 if clave.endswith("potencia") else 3
    return 6


def tabla_concepto(datos: dict, clave: str) -> pd.DataFrame:
    """Una fila por tarifa con los valores del concepto en P1..P6 (vacío si no tiene)."""
    filas = []
    for t in dr.TARIFAS:
        valores = list(datos["tarifas"][t].get(clave) or [])
        fila = {"Tarifa": t}
        fila.update({p: (valores[n] if n < len(valores) else None)
                     for n, p in enumerate(PERIODOS)})
        filas.append(fila)
    return pd.DataFrame(filas)


def aplicar_concepto(datos: dict, clave: str, tabla: pd.DataFrame) -> None:
    """Pasa a datos lo editado en la tabla del concepto (solo los periodos de cada tarifa).
    SRAD: si una tarifa queda sin ningún valor, sin SRAD (None)."""
    for _, fila in tabla.iterrows():
        t = fila["Tarifa"]
        n = _n_concepto(t, clave)
        valores = [None if pd.isna(fila[p]) else float(fila[p]) for p in PERIODOS[:n]]
        if clave == "srad" and all(v is None for v in valores):
            datos["tarifas"][t][clave] = None
        else:
            datos["tarifas"][t][clave] = valores


def tabla_perdidas_sistema(datos: dict) -> pd.DataFrame:
    """Pérdidas propias de cada tarifa en los sistemas no peninsulares (vacío = las de la
    tarifa)."""
    propias = datos.get("perdidas_sistema") or {}
    filas = []
    for s in dr.SISTEMAS[1:]:
        for t in dr.TARIFAS:
            valores = (propias.get(s) or {}).get(t) or []
            fila = {"Sistema": s, "Tarifa": t}
            fila.update({p: (valores[n] if n < len(valores) else None)
                         for n, p in enumerate(PERIODOS)})
            filas.append(fila)
    return pd.DataFrame(filas)


def aplicar_perdidas_sistema(datos: dict, tabla: pd.DataFrame) -> None:
    propias = {}
    for _, fila in tabla.iterrows():
        n = dr.n_periodos(fila["Tarifa"])
        valores = [None if pd.isna(fila[p]) else float(fila[p]) for p in PERIODOS[:n]]
        if any(v is not None for v in valores):
            propias.setdefault(fila["Sistema"], {})[fila["Tarifa"]] = valores
    datos["perdidas_sistema"] = propias


def _origen(d: dict, info: dict) -> str:
    if d["origen"] == "copia":
        return (f"Año sin datos propios: es una **copia de {d['copia_de']}**. Al guardar se crean "
                f"los de {d['anio']}.")
    if d["origen"] == "guardado":
        fecha, usuario = info.get(d["anio"], ("", ""))
        return (f"Datos guardados en la plataforma (última modificación: {fecha[:16].replace('T', ' ')}"
                f" por {usuario}).")
    return "Datos de la semilla del programa (regulados/regulados.json), sin cambios."


def pagina_datos_regulados():
    st.markdown("Peajes y cargos de todas las tarifas, pagos por capacidad, pérdidas, SRAD, "
                "impuestos, FNEE, pagos al operador del sistema y del mercado, y los parámetros "
                "con que se estima el precio de mercado. Los usan el **análisis de ofertas** y el "
                "**estudio de potencia**. Los años sin datos propios copian el último año que sí "
                "tiene (el análisis lo avisa).")
    usuario = estilo._usuario()
    anios = dr.anios()
    opciones = anios + [anios[-1] + 1]
    anio = st.selectbox(
        "Año natural", opciones, index=len(anios) - 1, key="dr_anio",
        format_func=lambda a: f"{a} (nuevo, copia de {anios[-1]})" if a not in anios else str(a))
    d = dr.datos_anio(anio)
    info = almacen.info_datos_regulados()
    st.info(_origen(d, info))
    clave = f"dr_{anio}_{info.get(anio, ('', ''))[0]}"   # se rehace al guardar

    t_tar, t_sis, t_gen, t_temp, t_fuentes = st.tabs(
        ["Tarifas", "Pérdidas por sistema", "Generales y mercado", "Temporadas", "Fuentes"])
    editadas = {}
    with t_tar:
        st.caption("Potencia en €/kW·año; energía, pagos por capacidad y SRAD en €/kWh; pérdidas "
                   "en tanto por uno. La 2.0TD tiene 2 periodos de potencia y 3 de energía: el "
                   "resto se ignora.")
        for c, nombre, unidad in dr.CONCEPTOS:
            st.markdown(f"**{nombre}** ({unidad})")
            editadas[c] = editor(tabla_concepto(d, c), f"ao_{clave}_{c}", PERIODOS,
                                 hide_index=True, use_container_width=True,
                                 disabled=["Tarifa"])
    with t_sis:
        st.caption("La Circular 3/2020 (art. 11) fija los mismos coeficientes en todo el "
                   "territorio. Rellena una fila solo si quieres usar otras pérdidas en un sistema "
                   "(por ejemplo, las publicadas por REE); vacía = las de la tarifa.")
        sistemas = editor(tabla_perdidas_sistema(d), f"ao_{clave}_sis", PERIODOS,
                          hide_index=True, use_container_width=True,
                          disabled=["Sistema", "Tarifa"])
    with t_gen:
        gen = {}
        c1, c2 = st.columns(2)
        for n, (c, nombre, unidad) in enumerate(dr.GENERALES):
            gen[c] = (c1 if n % 2 == 0 else c2).number_input(
                f"{nombre} ({unidad})", min_value=0.0, value=float(d["generales"][c]),
                format="%.10f" if c == "iee" else "%.7f", key=f"{clave}_g_{c}")
        st.markdown("**Desvíos de referencia**")
        st.caption("Media de los 12 últimos meses con liquidación C2 (coste + saldo = neto). "
                   "Solo informativo: aparece como ayuda en el campo «Desvíos» de las ofertas.")
        c1, c2 = st.columns(2)
        for n, (c, nombre, unidad) in enumerate(dr.DESVIOS):
            gen[c] = (c1 if n % 2 == 0 else c2).number_input(
                f"{nombre} ({unidad})", value=float(d["generales"].get(c) or 0.0),
                format="%.3f", key=f"{clave}_g_{c}")
        gen["desvios_fuente"] = st.text_input(
            "Periodo y fuente de los desvíos", d["generales"].get("desvios_fuente", ""),
            key=f"{clave}_g_desvios_fuente")
        st.markdown("**Estimación del precio de mercado**")
        st.caption("Precio previsto de cada mes = % OMIP × cotización OMIP + (1 − % OMIP) × media "
                   "del precio OMIE de ese mes en los últimos años.")
        m1, m2 = st.columns(2)
        pct = m1.number_input("Ponderación de OMIP (tanto por uno)", min_value=0.0, max_value=1.0,
                              value=float(d["mercado"]["pct_omip"]), step=0.05,
                              key=f"{clave}_m_pct")
        anios_omie = m2.number_input("Años de la media de OMIE", min_value=1, max_value=10,
                                     value=int(d["mercado"]["anios_media_omie"]), step=1,
                                     key=f"{clave}_m_anios")
    with t_temp:
        com = dr.comunes()
        st.caption(com["temporadas"]["_fuente"] + " " + com["periodos_temporada"]["_fuente"])
        st.dataframe(pd.DataFrame([{"Mes": m, **{s: dr.temporada(s, m) for s in dr.SISTEMAS}}
                                   for m in range(1, 13)]), hide_index=True)
        st.dataframe(pd.DataFrame([{"Temporada": t, **{
            s: " / ".join(f"P{p}" for p in com["periodos_temporada"][s][t]) for s in dr.SISTEMAS}}
            for t in dr.TEMPORADAS]), hide_index=True)
        st.caption("Las temporadas y los periodos son los de la Circular 3/2020; no se editan "
                   "aquí.")
    with t_fuentes:
        fuentes = st.text_area("Fuentes normativas (una por línea)",
                               "\n".join(d.get("fuentes") or []), height=220,
                               key=f"{clave}_fuentes")
        revisado = st.text_input("Revisado el", d.get("revisado", ""), key=f"{clave}_rev")

    nuevo = dr.limpiar(d)
    for c, tabla in editadas.items():
        aplicar_concepto(nuevo, c, tabla)
    aplicar_perdidas_sistema(nuevo, sistemas)
    nuevo["generales"] = gen
    nuevo["mercado"] = {"pct_omip": float(pct), "anios_media_omie": int(anios_omie)}
    nuevo["fuentes"] = [f.strip() for f in fuentes.splitlines() if f.strip()]
    nuevo["revisado"] = revisado.strip()

    b1, b2, b3, b4 = st.columns(4)
    if b1.button(f"Guardar los datos de {anio}", type="primary"):
        try:
            dr.guardar(anio, nuevo, usuario)
        except dr.DatosRegulados as e:
            st.error("No se ha guardado:\n" + "\n".join("- " + x for x in str(e).splitlines()))
        else:
            almacen.registrar_actividad(usuario, MODULO, "Datos regulados", f"Año {anio}")
            st.success(f"Datos regulados de {anio} guardados.")
            st.rerun()
    if d["origen"] == "guardado" and b2.button("Volver a los de la semilla",
                                               help="Borra lo guardado de este año (queda en el "
                                                    "historial)."):
        almacen.borrar_datos_regulados(anio)
        almacen.registrar_actividad(usuario, MODULO, "Datos regulados",
                                    f"Año {anio}: vuelta a la semilla")
        st.rerun()
    estilo.boton_descarga(MODULO, b3, "Descargar todo (JSON)", data=dr.exportar_json(),
                          file_name="regulados.json", mime="application/json",
                          help="Para actualizar regulados/regulados.json del repositorio.")
    estilo.boton_descarga(MODULO, b4, "precios.json del estudio de potencia",
                          data=json.dumps(dr.precios_potencia(anio), ensure_ascii=False,
                                          indent=2),
                          file_name="precios.json", mime="application/json",
                          help="Para el programa de escritorio EstudioPotencia (config/).")
    historial = almacen.historial_datos_regulados(anio)
    if historial:
        with st.expander("Historial de cambios de este año"):
            st.dataframe(pd.DataFrame([{"Fecha": h["fecha"][:16].replace("T", " "),
                                        "Usuario": h["usuario"]} for h in historial]),
                         hide_index=True)
