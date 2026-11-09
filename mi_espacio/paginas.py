# -*- coding: utf-8 -*-
"""
MI ESPACIO: PÁGINAS
===================
    pagina_mi_panel()        página de entrada: contadores y bloques personalizables por usuario
    pagina_licitaciones()    próximas licitaciones con semáforo, estado, carga por meses y acciones
    pagina_ficha()           todo lo de un cliente o CUPS
    pagina_calendario()      vencimientos del mes
La lógica está en mi_espacio/datos.py. PAGINAS ({ruta: st.Page}) la rellena inicio.py para poder
abrir otras páginas.
"""

import calendar
import datetime as dt
import io
import time

import altair as alt
import pandas as pd
import streamlit as st

import almacen
import estilo
from mi_espacio import datos as P

MODULO = "Mi espacio"
PAGINAS = {}                    # ruta → st.Page (inicio.py)
TTL = 60                        # s que se reutilizan los datos cargados
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ss = st.session_state


# ------------------------------------------------------------------- utilidades
def _yo() -> dict:
    return ss.get("usuario") or {}


def _fuente(completa: bool = False) -> P.Fuente:
    """Datos del panel, reutilizados TTL segundos en la sesión."""
    clave = "mi_fuente_completa" if completa else "mi_fuente"
    guardada = ss.get(clave)
    if not guardada or time.time() - guardada[0] > TTL:
        guardada = (time.time(), P.cargar(dt.date.today(), completa))
        ss[clave] = guardada
    return guardada[1]


def refrescar():
    for k in ("mi_fuente", "mi_fuente_completa", "mi_avisos"):
        ss.pop(k, None)


def _pref(clave, defecto=None):
    try:
        return almacen.leer_preferencia(_yo().get("usuario"), clave, defecto)
    except Exception:                                       # noqa: BLE001
        return defecto


def _guardar_pref(clave, valor):
    almacen.guardar_preferencia(_yo().get("usuario"), clave, valor)
    ss.pop("mi_avisos", None)


def generales() -> dict:
    """Umbrales generales (Administración)."""
    import json
    try:
        return json.loads(almacen.ajuste_sgc("panel_umbrales") or "{}")
    except ValueError:
        return {}


def mis_umbrales() -> dict:
    return P.umbrales(generales(), _pref("umbrales", {}))


def _nombres() -> dict:
    return P.mis_nombres(_yo().get("usuario"), _fuente().equipo)


def _todos_por_defecto(nombres) -> bool:
    """Sin nombres en la tabla del equipo, un supervisor o administrador ve todo."""
    return not P.tiene_nombres(nombres) and almacen.puede_supervisar(_yo())


def _abrir(ruta: str):
    pagina = PAGINAS.get(ruta)
    if pagina is not None:
        st.switch_page(pagina)


def abrir_registro(origen: str, id_):
    """Carga el registro en su módulo y abre la página."""
    if origen == "analisis":
        from analisis_ofertas import paginas as ap
        if ap.cargar_analisis(int(id_)):
            _abrir("analisis-ofertas")
    elif origen == "solicitud":
        from ofertas import paginas as op
        sol = almacen.solicitud(int(id_))
        if sol:
            op.cargar_en_formulario(sol)
            _abrir("solicitud-ofertas")
    elif origen == "estudio":
        from estudio_potencia import paginas as ep
        est = almacen.estudio_potencia(int(id_))
        if est:
            ep.cargar_en_formulario(est)
            _abrir("estudio-potencia")
    elif origen == "catalogo":
        _abrir("catalogo-ofertas")
    elif origen in ("suministro", "gemweb"):
        _abrir("seguimiento")
    elif origen == "envio":
        _abrir("seguimiento-envios")


def acciones_cups(cups: list, id_solicitud=None, clave: str = "acc"):
    """Botones «Crear solicitud», «Nuevo análisis» y «Estudio de potencia» con esos CUPS."""
    c1, c2, c3 = st.columns(3)
    texto = "\n".join(cups)
    if c1.button("Crear solicitud", key=clave + "_sol", disabled=not cups, type="primary",
                 help="Abre «Crear nueva solicitud» con estos CUPS ya pegados."):
        from ofertas import paginas as op
        op._limpiar_formulario()
        ss["gw_cups_texto"] = texto
        almacen.registrar_actividad(_yo().get("usuario"), MODULO, "Crear solicitud desde Mi espacio",
                                    f"{len(cups)} CUPS")
        _abrir("solicitud-ofertas")
    if c2.button("Nuevo análisis", key=clave + "_an", disabled=not id_solicitud,
                 help="Crea el análisis desde la solicitud de esta licitación." if id_solicitud
                 else "Primero hace falta una solicitud de ofertas guardada con estos CUPS."):
        ss["ao_modo_origen"] = "Una solicitud de ofertas guardada"
        ss["ao_f_usuarios"] = []
        ss["ao_sol"] = int(id_solicitud)
        _abrir("analisis-ofertas")
    if c3.button("Estudio de potencia", key=clave + "_pot", disabled=not cups,
                 help="Abre «Nuevo estudio» de potencia con estos CUPS ya pegados."):
        ss["pot_cups_texto"] = texto
        _abrir("estudio-potencia")


def _excel(hojas: dict) -> bytes:
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return salida.getvalue()


def _visible(df: pd.DataFrame) -> pd.DataFrame:
    """Sin las columnas internas (las que empiezan por _) y con el semáforo con su icono."""
    df = df[[c for c in df.columns if not str(c).startswith("_")]].copy()
    if "Semáforo" in df.columns:
        df["Semáforo"] = [f"{P.ICONO.get(v, '')} {v}" for v in df["Semáforo"]]
    return df


# =================================================================== contador del menú
def contador() -> int:
    """Nº de avisos del usuario para el título de «Mi panel» (0 si algo falla)."""
    try:
        guardado = ss.get("mi_avisos")
        if guardado and time.time() - guardado[0] < TTL:
            return guardado[1]
        yo = _yo()
        rol = yo.get("rol")
        claves = _pref("contador") or P.contador_por_defecto(rol)
        nombres = _nombres()
        n = sum(P.avisos(_fuente(), dt.date.today(), mis_umbrales(), yo.get("usuario"),
                         nombres, claves, _todos_por_defecto(nombres)).values())
        ss["mi_avisos"] = (time.time(), n)
        return n
    except Exception:                                       # noqa: BLE001
        return 0


# =================================================================== Mi panel
def pagina_mi_panel():
    yo = _yo()
    rol = yo.get("rol")
    nombres = _nombres()
    f = _fuente()
    hoy = dt.date.today()
    u = mis_umbrales()
    c1, c2 = st.columns([3, 1])
    quien = " y ".join(sorted(nombres["tecnico"] | nombres["auxiliar"])) or "—"
    c1.markdown(f"**{yo.get('nombre', '')}** · {rol or 'Técnico'} · en el cuadro de "
                f"Seguimiento: {quien}")
    if not P.tiene_nombres(nombres):
        c1.caption("No estás en la tabla del equipo de Seguimiento y Control (Administración → "
                   "Seguimiento y Control → Equipo): no se puede saber qué suministros son tuyos.")
    todos = c2.toggle("Ver lo de todo el equipo", _todos_por_defecto(nombres), key="mi_todos")
    if c2.button("Actualizar datos", icon=":material/refresh:"):
        refrescar()
        st.rerun()

    config = P.config_panel(rol, _pref("panel"))
    claves_contador = _pref("contador") or P.contador_por_defecto(rol)
    datos = {c["bloque"]: P.datos_bloque(c["bloque"], f, hoy, u, yo.get("usuario"), nombres,
                                         todos)
             for c in config if c["visible"]}
    # contadores
    avisos = P.avisos(f, hoy, u, yo.get("usuario"), nombres, claves_contador, todos)
    if avisos:
        cols = st.columns(min(6, len(avisos)))
        for i, (clave, n) in enumerate(avisos.items()):
            cols[i % len(cols)].metric(P.POR_CLAVE[clave].aviso, n)
    # bloques, de dos en dos
    visibles = [c for c in config if c["visible"]]
    for i in range(0, len(visibles), 2):
        cols = st.columns(2)
        for col, c in zip(cols, visibles[i:i + 2]):
            with col.container(border=True):
                _bloque(c, datos.get(c["bloque"], []))
    _personalizar(rol, config, claves_contador)


def _bloque(c: dict, filas: list):
    b = P.POR_CLAVE[c["bloque"]]
    st.markdown(f"**{b.titulo}**" + (f" · {len(filas)}" if filas and b.clave not in
                                      ("ultimo", "favoritos") else ""))
    if b.clave == "favoritos":
        favs = _pref("favoritos", [])
        if not favs:
            st.caption("Marca tus páginas con ☆ en su cabecera y saldrán aquí.")
        for ruta in favs:
            if ruta in PAGINAS:
                st.page_link(PAGINAS[ruta], icon=":material/star:")
        return
    if b.clave == "ultimo":
        filas = ultimo_trabajo(c["filas"])
    if not filas:
        st.caption("Nada pendiente." if b.clave != "ultimo" else "Aún no hay trabajo guardado.")
        return
    df = pd.DataFrame(filas[:c["filas"]])
    sel = st.dataframe(_visible(df), hide_index=True, use_container_width=True,
                       on_select="rerun", selection_mode="single-row",
                       key=f"mi_bl_{b.clave}")
    c1, c2 = st.columns(2)
    elegida = sel.selection.rows if sel and sel.selection else []
    if elegida and "_origen" in df.columns:
        fila = df.iloc[elegida[0]]
        if c1.button("Abrir", key=f"mi_ab_{b.clave}", type="primary"):
            abrir_registro(fila["_origen"], fila.get("_id"))
    if b.pagina and c2.button("Ver todo", key=f"mi_vt_{b.clave}"):
        _abrir(b.pagina)


def ultimo_trabajo(n: int = 6) -> list:
    from supervision import datos as sd
    hoy = dt.date.today()
    filas = sd.trabajo(hoy - dt.timedelta(days=120), hoy, usuarios=[_yo().get("usuario")])
    filas = sorted(filas, key=lambda x: x["fecha"], reverse=True)
    vistos, salida = set(), []
    for x in filas:
        if (x["origen"], x["id"]) in vistos:
            continue
        vistos.add((x["origen"], x["id"]))
        salida.append({"Fecha": str(x["fecha"])[:16].replace("T", " "), "Qué": x["tipo"],
                       "Descripción": x["descripcion"], "_origen": x["origen"], "_id": x["id"]})
        if len(salida) >= n:
            break
    return salida


def _personalizar(rol, config, claves_contador):
    with st.expander("Personalizar mi panel", icon=":material/tune:"):
        st.caption("Elige qué bloques ves, en qué orden y cuántas filas. Puedes añadir bloques "
                   "del otro rol. Los cambios solo te afectan a ti.")
        tabla = pd.DataFrame([{
            "Mostrar": c["visible"], "Orden": i + 1, "Bloque": P.POR_CLAVE[c["bloque"]].titulo,
            "Rol": P.POR_CLAVE[c["bloque"]].rol or "Común", "Filas": c["filas"],
            "_clave": c["bloque"]} for i, c in enumerate(config)])
        editada = st.data_editor(
            tabla, hide_index=True, use_container_width=True, key="mi_pers_tabla",
            disabled=["Bloque", "Rol"], column_order=["Mostrar", "Orden", "Bloque", "Rol", "Filas"],
            column_config={"Orden": st.column_config.NumberColumn(min_value=1, step=1),
                           "Filas": st.column_config.NumberColumn(min_value=1, max_value=50,
                                                                  step=1)})
        st.markdown("**Qué suma el contador de «Mi panel» en el menú**")
        con_aviso = [b for b in P.BLOQUES if b.aviso]
        cols = st.columns(3)
        elegidos = [b.clave for i, b in enumerate(con_aviso)
                    if cols[i % 3].checkbox(b.aviso, b.clave in claves_contador,
                                            key=f"mi_cont_{b.clave}")]
        st.markdown("**Mis umbrales** (vacío = los generales)")
        propios = _pref("umbrales", {})
        gen = P.umbrales(generales())
        cols = st.columns(2)
        nuevos = {}
        for i, (k, texto) in enumerate(P.NOMBRES_UMBRAL.items()):
            v = cols[i % 2].number_input(f"{texto} · general {gen[k]}", min_value=0,
                                         value=propios.get(k), step=1, key=f"mi_umb_{k}",
                                         placeholder=str(gen[k]))
            if v is not None:
                nuevos[k] = int(v)
        b1, b2 = st.columns(2)
        if b1.button("Guardar mi panel", type="primary", key="mi_pers_guardar"):
            orden = editada.sort_values(["Orden"], kind="stable")
            _guardar_pref("panel", [{"bloque": r["_clave"], "visible": bool(r["Mostrar"]),
                                     "filas": int(r["Filas"] or 5)}
                                    for _, r in orden.iterrows()])
            _guardar_pref("contador", elegidos)
            _guardar_pref("umbrales", nuevos)
            almacen.registrar_actividad(_yo().get("usuario"), MODULO, "Personalizar panel")
            st.rerun()
        if b2.button("Volver a la configuración de mi rol", key="mi_pers_rol"):
            for clave in ("panel", "contador", "umbrales"):
                almacen.borrar_preferencia(_yo().get("usuario"), clave)
            for k in list(ss.keys()):
                if str(k).startswith(("mi_pers_", "mi_cont_", "mi_umb_")):
                    del ss[k]
            ss.pop("mi_avisos", None)
            st.rerun()


# =================================================================== Próximas licitaciones
def pagina_licitaciones():
    f = _fuente()
    hoy = dt.date.today()
    u = mis_umbrales()
    nombres = _nombres()
    previos = _pref("filtros_licitaciones", {})
    c1, c2, c3, c4 = st.columns([1.2, 1, 1, 1])
    tipo = c1.selectbox("Fecha clave", P.TIPOS_FECHA,
                        index=P.TIPOS_FECHA.index(previos.get("tipo", "Automática")),
                        key="lic_tipo",
                        help="Automática: fecha prevista de lanzamiento; si no hay, la de "
                             f"denuncia; si tampoco, el fin de contrato menos {u['margen']} "
                             "meses.")
    desde = c2.date_input("Desde", hoy - dt.timedelta(days=30), format="DD/MM/YYYY",
                          key="lic_desde")
    hasta = c3.date_input("Hasta", hoy + dt.timedelta(days=365), format="DD/MM/YYYY",
                          key="lic_hasta")
    todos = c4.toggle("Todo el equipo", previos.get("todos", _todos_por_defecto(nombres)),
                      key="lic_todos", help="Si no, solo los suministros donde eres técnico o "
                                            "auxiliar.")
    base = P.licitaciones(f, hoy, u, tipo)
    if base.empty:
        st.info("Ningún suministro tiene fechas de contrato en Seguimiento y Control.")
        return
    if not todos:
        base = base[[(t in nombres["tecnico"]) or (a in nombres["auxiliar"])
                     for t, a in zip(base[P.TECNICO], base[P.AUXILIAR])]]
    base = base[(base["Fecha clave"] >= desde) & (base["Fecha clave"] <= hasta)]
    c1, c2, c3, c4 = st.columns(4)
    tecnicos = c1.multiselect("Técnico", sorted({x for x in base[P.TECNICO] if x}), key="lic_tec")
    clientes = c2.multiselect("Cliente", sorted({x for x in base[P.CLIENTE] if x}), key="lic_cli")
    coms = c3.multiselect("Comercializadora", sorted({x for x in base["Comercializadora"] if x}),
                          key="lic_com")
    tarifas = c4.multiselect("Tarifa", sorted({x for x in base["Tarifa"] if x}), key="lic_tar")
    for col, valores in ((P.TECNICO, tecnicos), (P.CLIENTE, clientes),
                         ("Comercializadora", coms), ("Tarifa", tarifas)):
        if valores:
            base = base[base[col].isin(valores)]
    lic = P.agrupar_licitaciones(base, f, u)
    semaforos = st.multiselect("Semáforo", [P.ROJO, P.AMBAR, P.VERDE, P.GRIS], key="lic_sem",
                               format_func=lambda s: f"{P.ICONO[s]} {s}",
                               help=f"Rojo: menos de {u['rojo']} días; ámbar: hasta "
                                    f"{u['ambar']}; verde: más; gris: la fecha ya pasó y no hay "
                                    "solicitud ni análisis.")
    if semaforos and not lic.empty:
        lic = lic[lic["Semáforo"].isin(semaforos)]
    if (previos.get("tipo"), previos.get("todos")) != (tipo, todos):
        _guardar_pref("filtros_licitaciones", {"tipo": tipo, "todos": todos})
    if lic.empty:
        st.info("Ninguna licitación con estos filtros.")
        return
    m = st.columns(5)
    for col, s in zip(m, (P.ROJO, P.AMBAR, P.VERDE, P.GRIS)):
        col.metric(f"{P.ICONO[s]} {s}", int((lic["Semáforo"] == s).sum()))
    m[4].metric("Suministros", int(lic["Suministros"].sum()))

    t1, t2 = st.tabs(["Licitaciones", "Carga de trabajo por meses"])
    with t1:
        st.caption("Una licitación = los suministros del mismo cliente con la misma fecha clave. "
                   "Elige una para ver sus CUPS y crear la solicitud.")
        sel = st.dataframe(_visible(lic), hide_index=True, use_container_width=True,
                           on_select="rerun", selection_mode="single-row", key="lic_tabla",
                           column_config={"Fecha clave": st.column_config.DateColumn(
                               format="DD/MM/YYYY")})
        filas = sel.selection.rows if sel and sel.selection else []
        if filas:
            x = lic.iloc[filas[0]]
            st.markdown(f"**{x[P.CLIENTE]}** · {x['Fecha clave']:%d/%m/%Y} · {x['Estado']}")
            st.dataframe(_visible(base[base["_id"].isin(x["_ids"])]), hide_index=True,
                         use_container_width=True)
            id_sol = x["_id_trabajo"] if x["_origen"] == "solicitud" else None
            if x["_origen"] in ("solicitud", "analisis") and st.button(
                    f"Abrir {'la solicitud' if x['_origen'] == 'solicitud' else 'el análisis'}",
                    key="lic_abrir"):
                abrir_registro(x["_origen"], x["_id_trabajo"])
            acciones_cups(list(x["_cups"]), id_sol, "lic")
        estilo.boton_descarga(MODULO, st, "Exportar a Excel", _excel(
            {"Licitaciones": _visible(lic), "Suministros": _visible(base)}),
            f"proximas_licitaciones_{hoy:%Y%m%d}.xlsx", XLSX)
    with t2:
        _graficas_carga(lic)


def _graficas_carga(lic: pd.DataFrame):
    mes = P.carga_por_mes(lic)
    por_tec = P.carga_por_mes(lic, por_tecnico=True)
    eje = alt.X("yearmonth(Mes):T", title=None, axis=alt.Axis(format="%b %y"))
    st.markdown("**Licitaciones por mes**")
    st.altair_chart(alt.Chart(mes).mark_bar(color=estilo.GRANATE).encode(
        eje, alt.Y("Licitaciones:Q", title="Licitaciones"),
        tooltip=[alt.Tooltip("yearmonth(Mes):T", title="Mes", format="%B %Y"),
                 "Licitaciones"]).properties(height=260), use_container_width=True)
    st.markdown("**Licitaciones por mes y técnico**")
    st.altair_chart(alt.Chart(por_tec).mark_bar().encode(
        eje, alt.Y("Licitaciones:Q", title="Licitaciones", stack="zero"),
        alt.Color(f"{P.TECNICO}:N", title="Técnico"),
        tooltip=[alt.Tooltip("yearmonth(Mes):T", title="Mes", format="%B %Y"), P.TECNICO,
                 "Licitaciones"]).properties(height=300), use_container_width=True)
    tabla = por_tec.pivot_table(index=P.TECNICO, columns="Mes", values="Licitaciones",
                                aggfunc="sum", fill_value=0)
    tabla.columns = [f"{c:%m/%y}" for c in tabla.columns]
    st.dataframe(tabla, use_container_width=True)


# =================================================================== Ficha de cliente / CUPS
def pagina_ficha():
    f = _fuente(completa=True)
    opciones = P.opciones_busqueda(f)
    elegida = st.selectbox("Cliente, titular, CIF, CUPS o suministro", [o[0] for o in opciones],
                           index=None, placeholder="Escribe para buscar…", key="ficha_busca")
    if not elegida:
        st.caption("Busca un cliente o un CUPS para ver todo lo que hay de él en la aplicación.")
        return
    tipo, valor = dict(opciones)[elegida]
    enc = P.buscar(f, tipo, valor)
    todo = P.ficha(f, enc)
    titulo = ", ".join(sorted(enc["clientes"])) or valor
    st.subheader(titulo)
    sums = enc["suministros"]
    if sums:
        d0 = sums[0]["datos"]
        c = st.columns(4)
        c[0].metric("Suministros", len(sums))
        c[1].metric("Consumo anual", f"{sum(float(s['datos'].get('Consumo anual (kWh)') or 0) for s in sums) / 1000:,.0f} MWh".replace(",", "."))
        c[2].metric("Técnico", ", ".join(sorted({P._texto(s["datos"].get(P.TECNICO)) for s in sums} - {""})) or "—")
        c[3].metric("Auxiliar", ", ".join(sorted({P._texto(s["datos"].get(P.AUXILIAR)) for s in sums} - {""})) or "—")
        st.caption(" · ".join(x for x in (P._texto(d0.get(P.CIF)), P._texto(d0.get(P.TITULAR)),
                                          P._texto(d0.get("Cadencia informe"))) if x))
    pestanas = ["Suministros"] + list(todo)
    tabs = st.tabs([f"{p} ({len(sums) if p == 'Suministros' else len(todo[p])})"
                    for p in pestanas])
    with tabs[0]:
        if sums:
            df = pd.DataFrame([{
                "CUPS": P._texto(s["datos"].get(P.CUPS)),
                "Denominación": P._texto(s["datos"].get(P.DENOMINACION)),
                "Tarifa": P._texto(s["datos"].get("Tarifa de acceso")),
                "Comercializadora": P._texto(s["datos"].get("Comercializadora")),
                "Consumo (kWh)": s["datos"].get("Consumo anual (kWh)"),
                "Denuncia": P.fecha(s["datos"].get(P.F_DENUNCIA)),
                "Fin de contrato": P.fecha(s["datos"].get(P.F_FIN)),
                "Lanzamiento": P.fecha(s["datos"].get(P.F_LANZAMIENTO)),
                "Periodicidad": P._texto(s["datos"].get("Cadencia informe")),
                "Técnico": P._texto(s["datos"].get(P.TECNICO)),
                "Auxiliar": P._texto(s["datos"].get(P.AUXILIAR))} for s in sums])
            st.dataframe(df, hide_index=True, use_container_width=True)
        else:
            st.caption("No está en Seguimiento y Control.")
        cups = sorted({P._texto(s["datos"].get(P.CUPS)) for s in sums} - {""}) or \
            sorted(enc["cups"])
        sol = next((r for r in todo["Solicitudes"]), None)
        acciones_cups(cups, sol["id"] if sol else None, "ficha")
    columnas = {
        "Solicitudes": ("solicitud", ["id", "fecha", "usuario", "ref", "cliente", "suministros",
                                      "estado", "comercializadoras"]),
        "Análisis": ("analisis", ["id", "fecha", "usuario", "ref", "cliente", "ofertas",
                                  "estado", "mejor", "ahorro"]),
        "Estudios de potencia": ("estudio", ["id", "fecha", "usuario", "tipo", "titular",
                                             "suministros", "ahorro"]),
        "Revisiones de SSAA": ("revision", ["id", "fecha", "usuario", "cups", "comercializadora",
                                            "factura", "periodo", "diferencia", "veredicto"]),
        "Fichas de contrato": ("ficha", ["cups", "actualizado", "usuario"])}
    for tab, nombre in zip(tabs[1:], todo):
        with tab:
            filas = todo[nombre]
            if not filas:
                st.caption("Nada guardado.")
                continue
            origen, cols = columnas[nombre]
            df = pd.DataFrame(filas)[[c for c in cols if c in pd.DataFrame(filas).columns]]
            sel = st.dataframe(df, hide_index=True, use_container_width=True,
                               on_select="rerun", selection_mode="single-row",
                               key=f"ficha_{origen}")
            elegidas = sel.selection.rows if sel and sel.selection else []
            if elegidas and origen in ("solicitud", "analisis", "estudio"):
                if st.button("Cargar en su módulo", key=f"ficha_ab_{origen}", type="primary"):
                    abrir_registro(origen, df.iloc[elegidas[0]]["id"])
            elif elegidas and origen == "revision":
                informe, c, fra = almacen.informe_revision(int(df.iloc[elegidas[0]]["id"]))
                if informe:
                    estilo.boton_descarga(MODULO, st, "Descargar informe de la revisión", informe,
                                          f"revision_SSAA_{c}_{fra}.xlsx", XLSX)
            elif origen == "ficha":
                if st.button("Abrir Fichas de contrato", key="ficha_fc"):
                    _abrir("fichas")


# =================================================================== Calendario
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def pagina_calendario():
    f = _fuente()
    hoy = dt.date.today()
    nombres = _nombres()
    c1, c2, c3, c4 = st.columns([1, 1, 2, 1])
    anio = c1.number_input("Año", 2020, 2100, hoy.year, key="cal_anio")
    mes = c2.selectbox("Mes", range(1, 13), index=hoy.month - 1, key="cal_mes",
                       format_func=lambda m: MESES[m - 1])
    tipos = c3.multiselect("Tipos", list(P.TIPOS_EVENTO), default=list(P.TIPOS_EVENTO),
                           key="cal_tipos")
    todos = c4.toggle("Todo el equipo", _todos_por_defecto(nombres), key="cal_todos")
    ini = dt.date(int(anio), mes, 1)
    fin = dt.date(int(anio), mes, calendar.monthrange(int(anio), mes)[1])
    evs = [e for e in P.eventos(f, ini, fin, mis_umbrales(), hoy, None if todos else nombres)
           if e["Tipo"] in tipos]
    tecnicos = sorted({e[P.TECNICO] for e in evs if e[P.TECNICO]})
    elegidos = st.multiselect("Técnico o usuario", tecnicos, key="cal_tec")
    if elegidos:
        evs = [e for e in evs if e[P.TECNICO] in elegidos]
    st.markdown(_cuadricula(int(anio), mes, evs, hoy), unsafe_allow_html=True)
    st.markdown(" ".join(f'<span style="background:{c};color:white;border-radius:4px;'
                         f'padding:1px 6px;font-size:0.8rem;margin-right:4px">{t}</span>'
                         for t, c in P.TIPOS_EVENTO.items() if t in tipos),
                unsafe_allow_html=True)
    if evs:
        df = pd.DataFrame(evs)
        st.markdown(f"**{len(evs)} vencimientos en {MESES[mes - 1].lower()} de {anio}**")
        sel = st.dataframe(_visible(df), hide_index=True, use_container_width=True,
                           on_select="rerun", selection_mode="single-row", key="cal_tabla",
                           column_config={"Fecha": st.column_config.DateColumn(
                               format="DD/MM/YYYY")})
        elegidas = sel.selection.rows if sel and sel.selection else []
        if elegidas and st.button("Abrir", key="cal_abrir", type="primary"):
            e = evs[elegidas[0]]
            abrir_registro(e["_origen"], e["_id"])
        estilo.boton_descarga(MODULO, st, "Exportar a Excel", _excel({"Vencimientos": _visible(df)}),
                              f"vencimientos_{anio}_{mes:02d}.xlsx", XLSX)
    else:
        st.info("Ningún vencimiento este mes con estos filtros.")


def _cuadricula(anio: int, mes: int, evs: list, hoy: dt.date) -> str:
    por_dia = {}
    for e in evs:
        por_dia.setdefault(e["Fecha"].day, []).append(e)
    cab = "".join(f"<th>{d}</th>" for d in ("Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"))
    filas = []
    for semana in calendar.Calendar(0).monthdayscalendar(anio, mes):
        celdas = []
        for d in semana:
            if not d:
                celdas.append("<td class='vacio'></td>")
                continue
            lista = por_dia.get(d, [])
            pastillas = "".join(
                f"<div class='ev' style='background:{P.TIPOS_EVENTO[e['Tipo']]}' "
                f"title='{e['Tipo']}: {e['Título']}'>{e['Título'][:26]}</div>"
                for e in lista[:3])
            if len(lista) > 3:
                pastillas += f"<div class='mas'>+{len(lista) - 3} más</div>"
            es_hoy = " hoy" if dt.date(anio, mes, d) == hoy else ""
            celdas.append(f"<td class='dia{es_hoy}'><div class='num'>{d}</div>{pastillas}</td>")
        filas.append("<tr>" + "".join(celdas) + "</tr>")
    estilo_css = (
        "<style>.cal{width:100%;border-collapse:collapse;table-layout:fixed}"
        ".cal th{background:#7B1E3A;color:white;padding:4px;font-size:.85rem}"
        ".cal td{border:1px solid #E5D5D5;vertical-align:top;height:92px;padding:3px}"
        ".cal td.vacio{background:#FAF6F6}.cal td.hoy{background:#F9E6EA}"
        ".cal .num{font-weight:600;font-size:.8rem;color:#555}"
        ".cal .ev{color:white;border-radius:4px;font-size:.7rem;padding:1px 4px;margin-top:2px;"
        "white-space:nowrap;overflow:hidden;text-overflow:ellipsis}"
        ".cal .mas{font-size:.7rem;color:#7B1E3A}</style>")
    return (estilo_css + f"<table class='cal'><tr>{cab}</tr>" + "".join(filas) + "</table>")


# =================================================================== Administración
def seccion_ajustes():
    """Umbrales generales del panel (Administración)."""
    import json
    st.markdown("### Mi espacio: umbrales generales")
    st.caption("Valores por defecto para todos; cada usuario puede poner los suyos en "
               "«Personalizar mi panel».")
    gen = P.umbrales(generales())
    cols = st.columns(2)
    nuevos = {k: int(cols[i % 2].number_input(texto, min_value=0, value=gen[k], step=1,
                                              key=f"adm_umb_{k}"))
              for i, (k, texto) in enumerate(P.NOMBRES_UMBRAL.items())}
    if st.button("Guardar umbrales generales", type="primary", key="adm_umb_guardar"):
        almacen.guardar_ajuste_sgc("panel_umbrales", json.dumps(nuevos))
        refrescar()
        st.success("Guardados.")
