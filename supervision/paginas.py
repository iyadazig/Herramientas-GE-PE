# -*- coding: utf-8 -*-
"""
PANEL DEL EQUIPO (Supervision)
==============================
Para supervisores y administradores: que ha hecho cada usuario y cuando. Arriba los filtros
(periodo, usuarios, modulos), los indicadores frente al periodo anterior y los graficos; debajo
el trabajo guardado (se abre cada registro para descargar su informe, Excel o PDF) y la
actividad (accesos, paginas, consultas a Gemweb y descargas). Solo consulta: no cambia nada.
"""

import datetime as dt

import pandas as pd
import streamlit as st

import almacen
import estilo
from supervision import datos as sd

ss = st.session_state
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MODULO = "Supervisión"


def _filtros(nombres):
    c1, c2, c3, c4 = st.columns([1.2, 1.4, 2, 2])
    nombre_periodo = c1.selectbox("Periodo", sd.PERIODOS, key="sup_periodo")
    rango = sd.periodo(nombre_periodo)
    if rango is None:
        hoy = dt.date.today()
        rango = c2.date_input("Desde / hasta", value=(hoy.replace(day=1), hoy),
                              format="DD/MM/YYYY", key="sup_rango")
        if len(rango) != 2:
            st.info("Elige la fecha final del periodo.")
            st.stop()
    else:
        c2.text_input("Desde / hasta", disabled=True,
                      value="%s – %s" % (rango[0].strftime("%d/%m/%Y"),
                                         rango[1].strftime("%d/%m/%Y")))
    usuarios = c3.multiselect("Usuarios", sorted(nombres, key=lambda u: nombres[u].lower()),
                              format_func=nombres.get, placeholder="Todos", key="sup_usuarios")
    modulos = c4.multiselect("Módulos", sd.MODULOS, placeholder="Todos", key="sup_modulos")
    return rango[0], rango[1], usuarios, modulos


def _indicadores(registros, previos, eventos, eventos_previos):
    cols = st.columns(len(sd.MODULOS) + 1)
    for col, modulo in zip(cols, sd.MODULOS):
        n = sum(f["modulo"] == modulo for f in registros)
        antes = sum(f["modulo"] == modulo for f in previos)
        col.metric(sd.INDICADORES[modulo], n, n - antes, help="Diferencia con el periodo anterior de la "
                                               "misma duración.")
    activos = {x["usuario"] for x in registros + eventos} - {None}
    antes = {x["usuario"] for x in previos + eventos_previos} - {None}
    cols[-1].metric("Usuarios con actividad", len(activos), len(activos) - len(antes))


def _detalle(fila, nombres):
    """Ficha del registro elegido con su descarga."""
    st.markdown("**%s** · %s · %s" % (fila["tipo"], sd.nombre(fila["usuario"], nombres),
                                       pd.to_datetime(fila["fecha"]).strftime("%d/%m/%Y %H:%M")))
    if fila["origen"] == "revision":
        informe, cups, factura = almacen.informe_revision(fila["id"])
        st.caption("%s · %s" % (fila["descripcion"] or "—", fila["resultado"] or "—"))
        if informe:
            estilo.boton_descarga(MODULO, st, "Descargar informe Excel", informe,
                                  "revision_SSAA_%s_%s.xlsx" % (cups or "sin_cups",
                                                                factura or fila["id"]), XLSX)
    elif fila["origen"] == "solicitud":
        sol = almacen.solicitud(fila["id"])
        if sol:
            st.caption("Ref. %s · %s · estado: %s%s" % (
                sol["ref"] or "—", sol["cliente"] or "—", sol["estado"],
                " · comercializadoras: " + sol["comercializadoras"]
                if sol["comercializadoras"] else ""))
            estilo.boton_descarga(MODULO, st, "Descargar Excel", sol["excel"] or b"",
                                  "solicitud_ofertas_%d.xlsx" % sol["id"], XLSX,
                                  disabled=not sol["excel"])
    elif fila["origen"] == "estudio":
        est = almacen.estudio_potencia(fila["id"])
        if est:
            st.caption("%s · %s" % (est["titular"] or "—", fila["resultado"]))
            estilo.boton_descarga(MODULO, st, "Descargar el PDF", est["pdf"] or b"",
                                  "estudio_potencia_%d.pdf" % est["id"], "application/pdf",
                                  disabled=not est["pdf"])
    elif fila["origen"] == "ficha":
        st.caption("Versiones de la ficha de contrato de %s:" % fila["cups"])
        st.dataframe(pd.DataFrame([{
            "Fecha": pd.to_datetime(h["fecha"]), "Usuario": sd.nombre(h["usuario"], nombres),
            "Acción": h["accion"]} for h in almacen.historial_contrato(fila["cups"])]),
            hide_index=True, use_container_width=True,
            column_config={"Fecha": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm")})


def _tabla_trabajo(registros, nombres):
    if not registros:
        st.info("No hay trabajo guardado con estos filtros.")
        return
    df = pd.DataFrame([{
        "Fecha": pd.to_datetime(f["fecha"]), "Usuario": sd.nombre(f["usuario"], nombres),
        "Módulo": f["modulo"], "Tipo": f["tipo"], "Descripción": f["descripcion"],
        "CUPS": f["cups"], "Resultado": f["resultado"]} for f in registros])
    st.caption("Pulsa en una fila para abrir el registro y descargar su informe, Excel o PDF.")
    sel = st.dataframe(df, hide_index=True, use_container_width=True, on_select="rerun",
                       selection_mode="single-row", key="sup_trabajo",
                       column_config={"Fecha": st.column_config.DatetimeColumn(
                           format="DD/MM/YYYY HH:mm")})
    filas = sel.selection.rows if sel else []
    if filas and filas[0] < len(registros):
        with st.container(border=True):
            _detalle(registros[filas[0]], nombres)


def _tabla_actividad(eventos, nombres):
    if not eventos:
        st.info("No hay actividad registrada con estos filtros. El registro de accesos, "
                "páginas, consultas y descargas empezó con esta versión de la plataforma.")
        return
    acciones = sorted({e["accion"] for e in eventos})
    elegidas = st.multiselect("Acción", acciones, placeholder="Todas", key="sup_acciones")
    vista = [e for e in eventos if not elegidas or e["accion"] in elegidas]
    st.dataframe(pd.DataFrame([{
        "Fecha": pd.to_datetime(e["fecha"]), "Usuario": sd.nombre(e["usuario"], nombres),
        "Módulo": e["modulo"], "Acción": e["accion"], "Detalle": e["detalle"] or ""}
        for e in vista]),
        hide_index=True, use_container_width=True,
        column_config={"Fecha": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm")})
    accesos = [e for e in eventos if e["accion"] == sd.ACCESO]
    if accesos:
        st.markdown("**Inicios de sesión por día**")
        df = pd.DataFrame({"Día": [e["fecha"][:10] for e in accesos],
                           "Usuario": [sd.nombre(e["usuario"], nombres)
                                       for e in accesos]})
        tabla = df.pivot_table(index="Día", columns="Usuario", aggfunc="size",
                               fill_value=0).sort_index(ascending=False)
        tabla.index = [d[8:10] + "/" + d[5:7] + "/" + d[:4] for d in tabla.index]
        st.dataframe(tabla, use_container_width=True)


def pagina_panel():
    if not almacen.puede_supervisar(ss.usuario):
        st.warning("Solo los supervisores y los administradores pueden ver el panel del equipo.")
        return
    usuarios = almacen.listar_usuarios()
    nombres = {u["usuario"]: u["nombre"] for u in usuarios}
    ultimos = {u["usuario"]: u["ultimo_acceso"] for u in usuarios}
    desde, hasta, elegidos, modulos = _filtros(nombres)

    registros = sd.trabajo(desde, hasta, elegidos, modulos)
    eventos = sd.actividad(desde, hasta, elegidos)
    previo = sd.periodo_anterior(desde, hasta)
    _indicadores(registros, sd.trabajo(*previo, elegidos, modulos), eventos,
                 sd.actividad(*previo, elegidos))

    resumen, trabajo, actividad = st.tabs(["Resumen por usuario", "Trabajo guardado",
                                           "Actividad"])
    with resumen:
        filas = sd.resumen_por_usuario(registros, eventos, nombres, ultimos)
        if not filas:
            st.info("Nadie ha trabajado en la plataforma en este periodo.")
        else:
            st.dataframe(pd.DataFrame(filas).assign(**{"Último acceso": lambda d: pd.to_datetime(
                d["Último acceso"])}), hide_index=True, use_container_width=True,
                column_config={"Último acceso": st.column_config.DatetimeColumn(
                    format="DD/MM/YYYY HH:mm")})
            g1, g2 = st.columns(2)
            tabla = sd.por_usuario_y_modulo(registros, nombres)
            if not tabla.empty:
                g1.markdown("**Trabajos por usuario y módulo**")
                g1.bar_chart(tabla, horizontal=True)
            meses = sd.por_mes(registros, nombres)
            if not meses.empty:
                g2.markdown("**Trabajos por mes**")
                g2.bar_chart(meses)
    with trabajo:
        _tabla_trabajo(registros, nombres)
    with actividad:
        _tabla_actividad(eventos, nombres)

    st.divider()
    estilo.boton_descarga(MODULO, st, "Exportar a Excel",
                          sd.excel_supervision(registros, eventos, nombres, desde, hasta,
                                               ultimos),
                          "Supervision_%s_%s.xlsx" % (desde.isoformat(), hasta.isoformat()),
                          XLSX, type="primary",
                          help="Resumen por usuario, trabajo y actividad con los filtros de "
                               "arriba.")
