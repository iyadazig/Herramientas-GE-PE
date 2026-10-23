# -*- coding: utf-8 -*-
"""
PAGINAS DE «SOLICITUD DE OFERTAS» de Herramientas GE&PE
=======================================================
Excel de solicitud de ofertas (Resumen + una hoja por suministro) a partir de los datos de
Gemweb: facturas de los ultimos 12 meses, potencias de la ultima factura y datos de la
entidad y del centro, mas las consideraciones de la oferta (tipologias de precio, fechas,
exencion...), comunes a todos los suministros y modificables en cada uno. Usa las funciones
de consulta de gemweb_extractor.paginas; el formato del Excel esta en ofertas/solicitud.py y
el texto del correo en ofertas/correo.py.

Las solicitudes se guardan en la base de datos de la plataforma (almacen) con el Excel, el
texto y los datos del formulario: «Solicitudes guardadas» las lista para todos, con su
seguimiento, y permite cargarlas de nuevo en el formulario (al guardar se crea otra).
"""

import re
from datetime import date, datetime, time

import pandas as pd
import streamlit as st

import almacen
from gemweb_extractor import paginas as gw
from ofertas import correo
from ofertas import solicitud as so

# la registra inicio.py: «Solicitudes guardadas» vuelve a ella al cargar una solicitud
PAGINA_FORMULARIO = None

ESTADOS = ["Preparada", "Enviada", "Ofertas recibidas", "Adjudicada", "Cancelada"]

# columnas de la tabla de datos, en este orden (grupo = division de la entidad en Gemweb,
# razon social = nombre de la entidad, alias = final del nombre del suministro)
COLUMNAS_FICHA = {"grupo": "Cliente/Grupo", "razon": "Razón social", "cif": "CIF",
                  "alias": "Alias", "cups": "CUPS", "direccion": "Dirección de suministro",
                  "tarifa": "Tarifa"}

# tabla de consideraciones de la oferta (una fila por suministro)
INICIO, DURACION, DURACION2, DURACION3 = (
    "Fecha prevista de inicio", "Duración del contrato hasta", "Duración del contrato hasta 2",
    "Duración del contrato hasta 3")
TEMPORADA, EXENCION = "De temporada", "Sujeto a exención de impuesto eléctrico"
GDO, EXCEDENTES = "Incluye GdO", "Compensación de excedentes"
PCT_CONSUMO, PCT_TIPO = "% consumo sujeto a exención", "% tipo de exención"
COLUMNAS_FECHA = [INICIO, DURACION, DURACION2, DURACION3]
PLAZOS = ("duracion", "duracion2", "duracion3")
TIPOLOGIAS = [so.PRECIO_FIJO] + so.PRECIO_INDEXADO
COLUMNAS_CASILLA = TIPOLOGIAS + [TEMPORADA, EXENCION, GDO, EXCEDENTES]

# gas natural: tipologias, particularidades y aspectos tecnicos (clave -> columna)
TIPOLOGIAS_GAS = [so.PRECIO_FIJO] + so.PRECIO_INDEXADO_GAS
GDO_GAS, IEH, PCT_IEH = ("GdO de gas renovable", "Impuesto sobre Hidrocarburos",
                         "% consumo (impuesto)")
COLUMNAS_CASILLA_GAS = TIPOLOGIAS_GAS + [GDO_GAS]
COLUMNAS_TECNICOS = {"qd": "Qd a contratar (kWh/día)", "telemedida": "Telemedida",
                     "termino_fijo": "Término fijo", "presion": "Presión (bar)",
                     "cogeneracion": "Cogeneración", "planta_satelite": "Planta satélite GNL",
                     "dias": "Días de uso a la semana", "interrumpible": "Interrumpible"}
NUMEROS_TECNICOS = ("qd", "presion", "dias")

# claves de los controles del formulario que se guardan para volver a cargarlo
CLAVES_FORMULARIO = (["sol_ref", "sol_fecha", "sol_gestor", "sol_tel", "sol_email",
                      "sol_producto", "sol_rec_desde", "sol_rec_hasta", "sol_rec_hora",
                      "sol_temporada", "sol_gdo", "sol_excedentes", "sol_exencion",
                      "sol_pct_consumo", "sol_pct_tipo", "sol_inicio", "sol_duracion",
                      "sol_duracion2", "sol_duracion3", "sol_gdo_gas", "sol_ieh",
                      "sol_pct_ieh"]
                     + ["sol_op_" + o for o in TIPOLOGIAS]
                     + ["sol_opg_" + o for o in TIPOLOGIAS_GAS])


def pagina_solicitud_ofertas():
    cargada = st.session_state.get("sol_cargada")
    if cargada:
        st.info(f"Solicitud nº {cargada['id']} cargada (Ref. {cargada['ref'] or '—'}, guardada "
                f"por {cargada['usuario']} el {cargada['fecha']}). Cambia lo que necesites; al "
                "guardar se crea una solicitud nueva y la original no se modifica.")
        if st.button("Empezar una solicitud nueva"):
            _limpiar_formulario()
            st.rerun()
        seccion_solicitud_ofertas({"clave": ("guardada", cargada["id"]),
                                   "fichas": cargada["fichas"], "comun": cargada["comun"]})
        return

    cliente = gw.obtener_cliente()
    st.markdown("Prepara el Excel de **solicitud de ofertas** de varios suministros a la vez, "
                "de electricidad o de gas natural: datos del suministro, potencias (o, en gas, "
                "caudal diario Qd) de su última factura y perfil de consumo de los últimos 12 "
                "meses, sacados de Gemweb.")
    st.subheader("1. Suministros")
    cups_list = gw.pedir_cups()
    st.subheader("2. Facturas para el perfil de consumo")
    modo, ini, fin = gw.elegir_periodo()

    st.subheader("3. Datos de Gemweb")
    clave = (tuple(cups_list), modo, ini, fin)
    if st.button("Buscar datos en Gemweb", type="primary", disabled=not cups_list):
        st.session_state.solicitud_ofertas = dict(
            gw.calcular_consumo_periodos(cliente, cups_list, modo, ini, fin, telelecturas=False,
                                         fichas_solicitud=True), clave=clave)
    res = st.session_state.get("solicitud_ofertas")
    if not res or res["clave"] != clave:
        return
    problemas = res["resumen"][res["resumen"]["Estado"] != "OK"]
    if not problemas.empty:
        st.warning("Algunos CUPS no tienen todos los datos:")
        st.dataframe(problemas[["CUPS", "Facturas", "Estado"]], hide_index=True,
                     use_container_width=True)
    seccion_solicitud_ofertas(res)


# ------------------------------------------------------------------- utilidades
def _texto(valor) -> str:
    """Celda de una tabla editada como texto ('' si vacia; fechas dd/mm/aaaa)."""
    if valor is None or valor is pd.NaT or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


def _fecha(texto):
    return datetime.strptime(texto, "%d/%m/%Y").date()


def _valor(clave, defecto):
    """value= de un control: el de la solicitud cargada si la hay, si no el de por defecto.
    Siempre el mismo en cada recarga (si cambiara, Streamlit reiniciaria el control)."""
    return {"value": st.session_state.get("sol_valores", {}).get(clave, defecto)}


def _indice(clave, opciones, defecto):
    """index= de un desplegable, como _valor (si lo guardado ya no es una opcion, el de por
    defecto)."""
    valor = st.session_state.get("sol_valores", {}).get(clave, defecto)
    return {"index": opciones.index(valor if valor in opciones else defecto)}


def _si(valor) -> bool:
    """Casilla de una tabla editada (vacia = no)."""
    return _texto(valor) != "" and bool(valor)


def codificar(valor):
    """Valor de un control a JSON (fechas y horas marcadas para poder leerlas despues)."""
    if isinstance(valor, datetime):
        valor = valor.date()
    if isinstance(valor, date):
        return {"fecha": valor.isoformat()}
    if isinstance(valor, time):
        return {"hora": valor.strftime("%H:%M")}
    return valor


def decodificar(valor):
    if isinstance(valor, dict) and "fecha" in valor:
        return date.fromisoformat(valor["fecha"])
    if isinstance(valor, dict) and "hora" in valor:
        return datetime.strptime(valor["hora"], "%H:%M").time()
    return valor


def _limpiar_formulario():
    for k in [k for k in st.session_state if str(k).startswith("sol_")]:
        del st.session_state[k]


# ------------------------------------------------------------------- consideraciones
def consideraciones_de(fila) -> dict:
    """Consideraciones de la oferta de un suministro a partir de su fila de la tabla."""
    exento = bool(fila[EXENCION])
    return {"fijo": bool(fila[so.PRECIO_FIJO]),
            "indexado": [o for o in so.PRECIO_INDEXADO if bool(fila[o])],
            "inicio": _texto(fila[INICIO]), "duracion": _texto(fila[DURACION]),
            "duracion2": _texto(fila[DURACION2]), "duracion3": _texto(fila[DURACION3]),
            "temporada": bool(fila[TEMPORADA]),
            "gdo": bool(fila[GDO]), "excedentes": bool(fila[EXCEDENTES]),
            "exencion": exento,
            "pct_consumo_exento": (float(fila[PCT_CONSUMO]) if exento and
                                   _texto(fila[PCT_CONSUMO]) else None),
            "pct_tipo_exencion": (float(fila[PCT_TIPO]) if exento and
                                  _texto(fila[PCT_TIPO]) else None)}


def fila_de(c: dict) -> dict:
    """Lo contrario de consideraciones_de: fila de la tabla a partir de lo guardado."""
    fila = {so.PRECIO_FIJO: c.get("fijo", True)}
    fila.update({o: o in (c.get("indexado") or []) for o in so.PRECIO_INDEXADO})
    for col, k in ((INICIO, "inicio"), (DURACION, "duracion"), (DURACION2, "duracion2"),
                   (DURACION3, "duracion3")):
        fila[col] = _fecha(c[k]) if c.get(k) else None
    fila.update({TEMPORADA: bool(c.get("temporada")), GDO: bool(c.get("gdo")),
                 EXCEDENTES: bool(c.get("excedentes")), EXENCION: bool(c.get("exencion")),
                 PCT_CONSUMO: c.get("pct_consumo_exento"), PCT_TIPO: c.get("pct_tipo_exencion")})
    return fila


def consideraciones_de_gas(fila) -> dict:
    """Gas: consideraciones de la oferta de un suministro a partir de su fila de la tabla."""
    return {"fijo": _si(fila[so.PRECIO_FIJO]),
            "indexado": [o for o in so.PRECIO_INDEXADO_GAS if _si(fila[o])],
            "inicio": _texto(fila[INICIO]), "duracion": _texto(fila[DURACION]),
            "duracion2": _texto(fila[DURACION2]), "duracion3": _texto(fila[DURACION3]),
            "temporada": False, "gdo": _si(fila[GDO_GAS]),
            "ieh": _texto(fila[IEH]) or so.IEH_REDUCIDO,
            "pct_ieh": float(fila[PCT_IEH]) if _texto(fila[PCT_IEH]) else None}


def fila_de_gas(c: dict) -> dict:
    """Lo contrario de consideraciones_de_gas."""
    fila = {so.PRECIO_FIJO: c.get("fijo", True)}
    fila.update({o: o in (c.get("indexado") or []) for o in so.PRECIO_INDEXADO_GAS})
    for col, k in ((INICIO, "inicio"), (DURACION, "duracion"), (DURACION2, "duracion2"),
                   (DURACION3, "duracion3")):
        fila[col] = _fecha(c[k]) if c.get(k) else None
    fila.update({GDO_GAS: bool(c.get("gdo")), IEH: c.get("ieh") or so.IEH_REDUCIDO,
                 PCT_IEH: c.get("pct_ieh")})
    return fila


# ------------------------------------------------------------------- aspectos tecnicos (gas)
def tabla_tecnicos(fichas: list) -> pd.DataFrame:
    """Una fila por suministro de gas con sus aspectos tecnicos, para editarla."""
    filas = []
    for f in fichas:
        consumo = so.consumo_anual(f)
        t = f.get("tecnicos") or so.tecnicos_por_defecto(consumo)
        filas.append(dict({"CUPS": f["cups"], "Tarifa": so.tarifa_gas(f.get("tarifa")),
                           "Consumo anual kWh": round(consumo)},
                          **{col: t.get(k) for k, col in COLUMNAS_TECNICOS.items()}))
    tabla = pd.DataFrame(filas)
    for k in NUMEROS_TECNICOS:
        tabla[COLUMNAS_TECNICOS[k]] = pd.to_numeric(tabla[COLUMNAS_TECNICOS[k]])
    return tabla


def tecnicos_de(fila) -> dict:
    """Aspectos tecnicos de un suministro a partir de su fila de la tabla."""
    t = {}
    for k, col in COLUMNAS_TECNICOS.items():
        valor = fila[col]
        if k in NUMEROS_TECNICOS:
            t[k] = float(valor) if _texto(valor) else None
        elif k == "termino_fijo":
            t[k] = _texto(valor) or so.CUOTA_FIJA
        else:
            t[k] = _si(valor)
    if t["dias"] is not None:
        t["dias"] = int(t["dias"])
    return t


def editor_tecnicos(fichas: list, clave) -> list:
    """Tabla editable de aspectos tecnicos (gas) y lo obligatorio aplicado: [tecnicos]."""
    with st.expander("Aspectos técnicos de cada suministro (gas)", expanded=True):
        st.caption("Tarifa, consumo y Qd (de la última factura) salen de Gemweb. La telemedida, "
                   "la presión, la cogeneración y la planta satélite no están en Gemweb: "
                   "revísalos. Término fijo: «Cuota fija» = por cliente; «Variable» = por "
                   "capacidad (Qd), que exige telemedida. Con más de 5 GWh/año la telemedida es "
                   "obligatoria (Orden IET/2446/2013) y el término fijo, variable.")
        cols = COLUMNAS_TECNICOS
        editada = st.data_editor(
            tabla_tecnicos(fichas), hide_index=True, use_container_width=True,
            disabled=["CUPS", "Tarifa", "Consumo anual kWh"],
            column_config={
                "Consumo anual kWh": st.column_config.NumberColumn(format="localized"),
                cols["qd"]: st.column_config.NumberColumn(min_value=0, format="localized"),
                cols["presion"]: st.column_config.NumberColumn(min_value=0, format="%g"),
                cols["dias"]: st.column_config.NumberColumn(min_value=1, max_value=7, step=1),
                cols["termino_fijo"]: st.column_config.SelectboxColumn(
                    options=[so.CUOTA_FIJA, so.VARIABLE], required=True),
                **{cols[k]: st.column_config.CheckboxColumn() for k in (
                    "telemedida", "cogeneracion", "planta_satelite", "interrumpible")}},
            key="sol_tecnicos_%d" % hash(clave))
    tecnicos, corregidos = [], []
    for f, (_, fila) in zip(fichas, editada.iterrows()):
        t, cambios = so.reglas_tecnicas(tecnicos_de(fila), so.consumo_anual(f))
        tecnicos.append(t)
        if cambios:
            corregidos.append(f"{f['cups']}: {', '.join(cambios)}")
    if corregidos:
        st.info("Se ha aplicado lo obligatorio con más de 5 GWh/año: " + "; ".join(corregidos))
    return tecnicos


def consideraciones_comunes() -> dict:
    """Selecciones para todos los suministros, con la disposicion del formulario de ofertas."""
    st.markdown("#### Consideraciones de la oferta")
    st.caption("Lo que elijas aquí se aplica a todos los suministros. Debajo puedes cambiarlo en "
               "los que necesiten otra cosa.")
    comun = {}
    c_tipo, c_part = st.columns(2)
    with c_tipo:
        st.markdown("**Tipología de precios**")
        for o in TIPOLOGIAS:
            clave = "sol_op_" + o
            comun[o] = st.checkbox(o, key=clave, **_valor(clave, o == so.PRECIO_FIJO),
                                   help=("Binómica en DH2 para la 2.0TD y en DH6 para la 3.0TD "
                                         "y las 6.xTD." if o == so.PRECIO_FIJO else None))
    with c_part:
        st.markdown("**Particularidades**")
        comun[TEMPORADA] = st.checkbox(TEMPORADA, key="sol_temporada",
                                       **_valor("sol_temporada", False),
                                       help="Sin marcar, el contrato es anual.")
        comun[GDO] = st.checkbox("Incluye garantías de origen (GdO)", key="sol_gdo",
                                 **_valor("sol_gdo", False))
        comun[EXCEDENTES] = st.checkbox(EXCEDENTES, key="sol_excedentes",
                                        **_valor("sol_excedentes", False))
        comun[EXENCION] = st.checkbox(EXENCION, key="sol_exencion",
                                      **_valor("sol_exencion", False))
        p1, p2 = st.columns(2)
        comun[PCT_CONSUMO] = p1.number_input("Porcentaje de consumo sujeto a exención (%)",
                                             min_value=0.0, max_value=100.0,
                                             disabled=not comun[EXENCION], key="sol_pct_consumo",
                                             **_valor("sol_pct_consumo", None))
        comun[PCT_TIPO] = p2.number_input("Porcentaje tipo de exención (%)", min_value=0.0,
                                          max_value=100.0, disabled=not comun[EXENCION],
                                          key="sol_pct_tipo", **_valor("sol_pct_tipo", None))
    return _fechas_comunes(comun)


def consideraciones_comunes_gas() -> dict:
    """Gas: tipologias (cada indexada sin y con coberturas), GdO de gas renovable e Impuesto
    sobre Hidrocarburos, para todos los suministros."""
    st.markdown("#### Consideraciones de la oferta")
    st.caption("Lo que elijas aquí se aplica a todos los suministros. Debajo puedes cambiarlo en "
               "los que necesiten otra cosa.")
    comun = {}
    c_tipo, c_part = st.columns(2)
    with c_tipo:
        st.markdown("**Tipología de precios**")
        for o in TIPOLOGIAS_GAS:
            clave = "sol_opg_" + o
            comun[o] = st.checkbox(o, key=clave, **_valor(clave, o == so.PRECIO_FIJO),
                                   help=("La comercializadora permite fijar el precio de parte "
                                         "del consumo durante el contrato (coberturas)."
                                         if o.endswith(so.CON_COBERTURAS) else None))
    with c_part:
        st.markdown("**Particularidades**")
        comun[GDO_GAS] = st.checkbox("Incluye garantías de origen de gas renovable",
                                     key="sol_gdo_gas", **_valor("sol_gdo_gas", False))
        comun[IEH] = st.selectbox(
            IEH, so.TIPOS_IEH, key="sol_ieh", **_indice("sol_ieh", so.TIPOS_IEH, so.IEH_REDUCIDO),
            help="Reducido (epígrafe 1.10.2): usos profesionales, sin cogeneración ni generación "
                 "eléctrica. General (1.10.1): el resto. Exento: generación eléctrica o "
                 "cogeneración con autorización de Hacienda, reducción química o usos no "
                 "combustibles (Ley 38/1992, art. 51).")
        comun[PCT_IEH] = st.number_input(
            "Porcentaje del consumo al que se aplica (%)", min_value=0.0, max_value=100.0,
            key="sol_pct_ieh", **_valor("sol_pct_ieh", None),
            help="Solo si no se aplica a todo el consumo: usos mixtos (art. 50.4) o una parte "
                 "exenta. Vacío = todo el consumo.")
    return _fechas_comunes(comun)


def _fechas_comunes(comun: dict) -> dict:
    """Fecha prevista de inicio y hasta tres fechas fin (comun a electricidad y gas)."""
    comun[INICIO] = st.date_input("Fecha prevista de inicio del contrato", format="DD/MM/YYYY",
                                  key="sol_inicio", **_valor("sol_inicio", None))
    st.caption("Duración del contrato: rellena al menos una fecha fin; cada una es una "
               "alternativa de plazo para la oferta.")
    c1, c2, c3 = st.columns(3)
    for col, clave, n in ((c1, DURACION, ""), (c2, DURACION2, "2"), (c3, DURACION3, "3")):
        comun[clave] = col.date_input(f"Duración del contrato hasta {n}".strip(),
                                      format="DD/MM/YYYY", key="sol_duracion" + n,
                                      **_valor("sol_duracion" + n, None))
    return comun


def tabla_consideraciones(fichas: list, comun: dict, guardado: dict = None,
                          gas: bool = False) -> pd.DataFrame:
    """Una fila por suministro con las selecciones comunes, para editarla. Si es una solicitud
    guardada y lo comun no ha cambiado desde entonces, con lo que tenia cada suministro."""
    usar_guardado = guardado is not None and guardado == comun
    de_guardado = fila_de_gas if gas else fila_de
    tabla = pd.DataFrame([dict({"CUPS": f["cups"], "Alias": f.get("alias", "")},
                               **(de_guardado(f["consideraciones"])
                                  if usar_guardado and f.get("consideraciones") else comun))
                          for f in fichas])
    for col in COLUMNAS_FECHA:
        tabla[col] = pd.to_datetime(tabla[col])
    for col in (PCT_IEH,) if gas else (PCT_CONSUMO, PCT_TIPO):
        tabla[col] = pd.to_numeric(tabla[col])
    return tabla


# ------------------------------------------------------------------- formulario
def seccion_solicitud_ofertas(res: dict) -> None:
    """Cabecera, tablas revisables de los suministros y botones de descarga, texto del correo
    y guardado."""
    fichas = res.get("fichas") or []
    if not fichas:
        st.warning("No se ha encontrado ningún suministro en Gemweb.")
        return
    st.subheader("4. Solicitud de ofertas")
    st.caption("Excel con el formato de GE&PE: una hoja «Resumen» y otra por suministro con sus "
               "datos, las potencias de su última factura (en gas, los aspectos técnicos), el "
               "perfil de consumo de los últimos 12 meses (cada factura en el mes de su fecha "
               "de fin) y las consideraciones de la oferta.")
    # producto de cada suministro segun Gemweb (las solicitudes antiguas, de electricidad)
    tipos = [f.get("tipo") or so.ELECTRICIDAD for f in fichas]
    if len(set(tipos)) > 1:
        por_tipo = {t: [f["cups"] for f, ti in zip(fichas, tipos) if ti == t] for t in set(tipos)}
        st.error("Hay CUPS de gas y de electricidad. Haz una solicitud para cada producto. "
                 f"Gas: {', '.join(por_tipo[so.GAS])}. Electricidad: "
                 f"{', '.join(por_tipo[so.ELECTRICIDAD])}.")
        return
    usuario = st.session_state.get("usuario") or {}
    c1, c2, c3 = st.columns(3)
    with c1:
        ref = st.text_input("Ref.", key="sol_ref", **_valor("sol_ref", ""))
        fecha = st.date_input("Fecha", format="DD/MM/YYYY", key="sol_fecha",
                              **_valor("sol_fecha", date.today()))
    with c2:
        gestor = st.text_input("Gestor de cuenta", key="sol_gestor",
                               **_valor("sol_gestor", usuario.get("nombre", "")))
        telefono = st.text_input("Teléfono", key="sol_tel",
                                 **_valor("sol_tel", so.TELEFONO_GEYPE))
    with c3:
        email = st.text_input("Email", key="sol_email",
                              help="El del usuario conectado. Si sale vacío, un administrador "
                                   "puede añadirlo en Administración → Usuarios.",
                              **_valor("sol_email", usuario.get("email") or ""))
        productos = list(so.PRODUCTOS)
        producto = st.selectbox(
            "Producto", productos, key="sol_producto",
            format_func=lambda p: "Gas natural" if p == so.GAS else "Electricidad",
            help="Sale elegido según el tipo de suministro de Gemweb.",
            **_indice("sol_producto", productos, tipos[0]))
    if producto != tipos[0]:
        st.warning("Según Gemweb, estos suministros son de "
                   f"{'gas' if tipos[0] == so.GAS else 'electricidad'}. Comprueba el producto.")
    gas = producto == so.GAS
    st.markdown("**Periodo para la recepción de las ofertas**")
    c1, c2, c3 = st.columns(3)
    desde = c1.date_input("Desde", format="DD/MM/YYYY", key="sol_rec_desde",
                          **_valor("sol_rec_desde", date.today()))
    hasta = c2.date_input("Hasta (fecha límite)", format="DD/MM/YYYY", key="sol_rec_hasta",
                          **_valor("sol_rec_hasta", None))
    hora = c3.time_input("Hora límite", key="sol_rec_hora", step=900,
                         **_valor("sol_rec_hora", time(12, 0)))
    if desde and hasta and hasta < desde:
        st.warning("La fecha límite de recepción es anterior a la de inicio.")

    st.caption("Revisa los datos de cada suministro (salen de Gemweb).")
    tabla = pd.DataFrame([{v: f.get(k, "") for k, v in COLUMNAS_FICHA.items()} for f in fichas])
    editada = st.data_editor(tabla, hide_index=True, use_container_width=True,
                             disabled=["CUPS", "Tarifa"], key="sol_tabla_%d" % hash(res["clave"]))
    potencias, tecnicos = None, []
    if gas:
        tecnicos = editor_tecnicos(fichas, res["clave"])
    else:
        with st.expander("Potencias que se usarán (última factura de cada suministro)"):
            st.caption("Salen de la última factura; cámbialas si hay que pedir otras (kW).")
            potencias = st.data_editor(
                pd.DataFrame([dict({"CUPS": f["cups"], "Última factura": f["ultima_factura"]},
                                   **{p: f["potencias"][n] for n, p in enumerate(so.PERIODOS)})
                              for f in fichas]),
                hide_index=True, use_container_width=True, disabled=["CUPS", "Última factura"],
                column_config={p: st.column_config.NumberColumn(min_value=0, format="%g")
                               for p in so.PERIODOS},
                key="sol_potencias_%d" % hash(res["clave"]))
    for f in fichas:
        if f.get("aviso"):
            st.warning(f"{f['cups']}: {f['aviso']}. El perfil Ene–Dic mezcla años; usa «Últimos "
                       "12 meses con factura» o un periodo de 12 meses.")

    comun = consideraciones_comunes_gas() if gas else consideraciones_comunes()
    st.caption("Consideraciones de cada suministro: salen con lo elegido arriba. Si después "
               "cambias algo de arriba, esta tabla vuelve a empezar con ello.")
    columnas = {c: st.column_config.DateColumn(format="DD/MM/YYYY") for c in COLUMNAS_FECHA}
    columnas.update({c: st.column_config.NumberColumn(min_value=0, max_value=100, format="%.2f")
                     for c in ((PCT_IEH,) if gas else (PCT_CONSUMO, PCT_TIPO))})
    columnas.update({c: st.column_config.CheckboxColumn()
                     for c in (COLUMNAS_CASILLA_GAS if gas else COLUMNAS_CASILLA)})
    tabla_c = tabla_consideraciones(fichas, comun, res.get("comun"), gas)
    if gas:
        columnas[IEH] = st.column_config.SelectboxColumn(options=so.TIPOS_IEH, required=True)
        # con cogeneracion no cabe el tipo reducido: se propone la exencion
        tabla_c[IEH] = [so.ieh_valido(v, t["cogeneracion"]) for v, t in zip(tabla_c[IEH],
                                                                             tecnicos)]
    editada_c = st.data_editor(
        tabla_c, hide_index=True, use_container_width=True, disabled=["CUPS", "Alias"],
        column_config=columnas,
        key="sol_consid_%d_%d" % (hash(res["clave"]), hash(repr(
            (producto, sorted(comun.items()), [t["cogeneracion"] for t in tecnicos])))))

    inversa = {v: k for k, v in COLUMNAS_FICHA.items()}
    suministros, sin_reducido = [], []
    filas_p = potencias.iterrows() if potencias is not None else ((None, None) for _ in fichas)
    for n, (f, (_, fila), (_, fila_c), (_, fila_p)) in enumerate(zip(
            fichas, editada.iterrows(), editada_c.iterrows(), filas_p)):
        s = dict(f, **{inversa[col]: _texto(valor) for col, valor in fila.items()})
        if gas:
            s["tecnicos"] = tecnicos[n]
            s["consideraciones"] = c = consideraciones_de_gas(fila_c)
            ieh = so.ieh_valido(c["ieh"], s["tecnicos"]["cogeneracion"])
            if ieh != c["ieh"]:
                c["ieh"] = ieh
                sin_reducido.append(s["cups"])
        else:
            s["potencias"] = [None if _texto(fila_p[p]) == "" else float(fila_p[p])
                              for p in so.PERIODOS]
            s["consideraciones"] = consideraciones_de(fila_c)
        s["fecha_inicio"] = s["consideraciones"]["inicio"]
        suministros.append(s)
    if sin_reducido:
        st.warning("Con cogeneración no cabe el tipo reducido del Impuesto sobre Hidrocarburos "
                   "(epígrafe 1.10.2): se pide «Exento» en " + ", ".join(sin_reducido)
                   + ". Cámbialo a «Tipo general» si no tiene autorización de exención.")
    avisos = [f"{s['cups']}: {a}" for s in suministros if gas for a in so.avisos_gas(s)]
    if avisos:
        st.warning("Revisa:\n" + "\n".join(f"- {a}" for a in avisos))
    plazos = [s["cups"] for s in suministros if s["consideraciones"]["inicio"] and any(
        s["consideraciones"][d] and _fecha(s["consideraciones"][d]) <= _fecha(
            s["consideraciones"]["inicio"]) for d in PLAZOS)]
    if plazos:
        st.warning("La duración del contrato termina antes de la fecha de inicio en: "
                   + ", ".join(plazos))
    sin_plazo = [s["cups"] for s in suministros
                 if not any(s["consideraciones"][d] for d in PLAZOS)]
    if sin_plazo:
        st.info("Falta al menos una fecha fin de duración del contrato en: "
                + ", ".join(sin_plazo))

    cabecera = {"ref": ref, "fecha": fecha.strftime("%d/%m/%Y") if fecha else "",
                "gestor": gestor, "telefono": telefono, "email": email,
                "producto": so.PRODUCTOS[producto], "producto_tipo": producto,
                "recepcion_desde": desde.strftime("%d/%m/%Y") if desde else "",
                "recepcion_hasta": hasta.strftime("%d/%m/%Y") if hasta else "",
                "recepcion_hora": hora.strftime("%H:%M") if hora else ""}
    if not hasta:
        st.info("Elige la fecha límite de recepción de ofertas (arriba, en «Periodo para la "
                "recepción de las ofertas») para poder descargar, ver el texto y guardar.")
    listo = bool(hasta) and not sin_plazo
    excel = so.excel_solicitud(cabecera, suministros) if listo else b""
    texto = correo.texto_correo(cabecera, suministros) if listo else ""
    nombre = "solicitud_ofertas_%s.xlsx" % (re.sub(r"[^\w-]", "_", ref) if ref else
                                             date.today().strftime("%d%m%Y"))
    b1, b2, b3 = st.columns(3)
    b1.download_button("Descargar solicitud de ofertas", data=excel, file_name=nombre,
                       disabled=not listo,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    if b2.button("Texto para el correo a las comercializadoras", disabled=not listo):
        st.session_state.sol_ver_correo = True
    if b3.button("Guardar solicitud", disabled=not listo,
                 help="Guarda el Excel, el texto del correo y todos los datos en el ordenador "
                      "de la plataforma. Se consultan en «Solicitudes guardadas»."):
        id_sol = guardar(cabecera, comun, suministros, excel, texto,
                         usuario.get("usuario", ""))
        st.session_state.sol_guardada = id_sol
    if st.session_state.get("sol_guardada"):
        st.success(f"Solicitud guardada con el nº {st.session_state.sol_guardada}. Está en "
                   "«Solicitud de ofertas → Solicitudes guardadas».")
    if listo and st.session_state.get("sol_ver_correo"):
        st.markdown("#### Texto para el correo")
        st.caption("Cópialo con el botón de la esquina del recuadro y pégalo en el correo. Se "
                   "actualiza solo si cambias algo de arriba.")
        st.code(texto, language=None, wrap_lines=True)


def cliente_de(suministros: list) -> str:
    grupos = list(dict.fromkeys(s.get("grupo") for s in suministros if s.get("grupo")))
    razones = list(dict.fromkeys(s.get("razon") for s in suministros if s.get("razon")))
    return ", ".join(grupos or razones)


def guardar(cabecera, comun, suministros, excel, texto, usuario) -> int:
    resumen = {"ref": cabecera.get("ref"), "cliente": cliente_de(suministros),
               "cups": [s["cups"] for s in suministros],
               "consumo": round(sum(so.consumo_anual(s) for s in suministros)),
               "recepcion_desde": cabecera.get("recepcion_desde"),
               "recepcion_hasta": " ".join(x for x in (cabecera.get("recepcion_hasta"),
                                                       cabecera.get("recepcion_hora")) if x)}
    datos = {"formulario": {k: codificar(st.session_state[k]) for k in CLAVES_FORMULARIO
                            if k in st.session_state},
             "comun": {k: codificar(v) for k, v in comun.items()},
             "cabecera": cabecera, "suministros": suministros}
    return almacen.guardar_solicitud(usuario, resumen, datos, texto, excel)


# ------------------------------------------------------------------- guardadas
def cargar_en_formulario(sol: dict) -> None:
    """Deja la solicitud guardada lista en el formulario (sus controles y sus suministros)."""
    _limpiar_formulario()
    datos = sol["datos"]
    st.session_state.sol_valores = {k: decodificar(v)
                                    for k, v in datos.get("formulario", {}).items()}
    fichas = []
    for s in datos["suministros"]:
        s = dict(s)
        s["perfil"] = {int(m): v for m, v in (s.get("perfil") or {}).items()}
        if s.get("perfil_gas"):
            s["perfil_gas"] = {int(m): v for m, v in s["perfil_gas"].items()}
        fichas.append(s)
    st.session_state.sol_cargada = {
        "id": sol["id"], "ref": sol.get("ref"), "usuario": sol.get("usuario"),
        "fecha": _fecha_hora(sol.get("fecha")), "fichas": fichas,
        "comun": {k: decodificar(v) for k, v in datos.get("comun", {}).items()}}


def _fecha_hora(iso: str) -> str:
    return datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M") if iso else ""


def pagina_solicitudes_guardadas():
    solicitudes = almacen.listar_solicitudes()
    if not solicitudes:
        st.info("Todavía no hay solicitudes guardadas. Se guardan con el botón «Guardar "
                "solicitud», al final de «Crear nueva solicitud».")
        return
    nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
    c1, c2 = st.columns([2, 1])
    buscar = c1.text_input("Buscar por cliente, CUPS, Ref. o comercializadora").strip().lower()
    estados = c2.multiselect("Estado", ESTADOS)
    filtradas = [s for s in solicitudes
                 if (not estados or s["estado"] in estados)
                 and (not buscar or buscar in " ".join(str(s.get(k) or "") for k in (
                     "cliente", "cups", "ref", "comercializadoras")).lower())]
    tabla = pd.DataFrame([{
        "Nº": s["id"], "Guardada": _fecha_hora(s["fecha"]),
        "Usuario": nombres.get(s["usuario"], s["usuario"]), "Ref.": s["ref"],
        "Producto": "Gas natural" if s.get("producto") == so.GAS else "Electricidad",
        "Cliente/Grupo": s["cliente"], "Suministros": s["suministros"],
        "Consumo anual kWh": s["consumo"], "Recepción desde": s["recepcion_desde"],
        "Recepción hasta": s["recepcion_hasta"], "Estado": s["estado"],
        "Comercializadoras": s["comercializadoras"]} for s in filtradas])
    st.dataframe(tabla, hide_index=True, use_container_width=True,
                 column_config={"Consumo anual kWh": st.column_config.NumberColumn(
                     format="localized")})
    if not filtradas:
        return

    st.subheader("Detalle")
    por_id = {s["id"]: s for s in filtradas}
    elegida = st.selectbox(
        "Solicitud", list(por_id), format_func=lambda i: "Nº %d · %s · %s%s" % (
            i, _fecha_hora(por_id[i]["fecha"]), por_id[i]["cliente"] or "—",
            f" · Ref. {por_id[i]['ref']}" if por_id[i]["ref"] else ""))
    sol = almacen.solicitud(elegida)
    c1, c2, c3 = st.columns(3)
    c1.metric("Suministros", sol["suministros"])
    c2.metric("Consumo anual previsto", f"{sol['consumo'] or 0:,.0f} kWh".replace(",", "."))
    c3.metric("Estado", sol["estado"])
    st.caption(f"Guardada por {nombres.get(sol['usuario'], sol['usuario'])} el "
               f"{_fecha_hora(sol['fecha'])}"
               + (f" · seguimiento actualizado por "
                  f"{nombres.get(sol['actualizado_por'], sol['actualizado_por'])} el "
                  f"{_fecha_hora(sol['actualizado'])}" if sol.get("actualizado") else ""))
    with st.expander(f"CUPS ({sol['suministros']})"):
        st.code(sol["cups"].replace(", ", "\n"), language=None)

    b1, b2 = st.columns(2)
    b1.download_button("Descargar Excel", data=sol["excel"] or b"",
                       file_name="solicitud_ofertas_%d%s.xlsx" % (
                           sol["id"], "_" + re.sub(r"[^\w-]", "_", sol["ref"]) if sol["ref"]
                           else ""),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", disabled=not sol["excel"])
    if b2.button("Cargar en el formulario",
                 help="Abre «Crear nueva solicitud» con esta solicitud para corregirla o "
                      "reutilizarla. Al guardar se crea una nueva."):
        cargar_en_formulario(sol)
        if PAGINA_FORMULARIO is not None:
            st.switch_page(PAGINA_FORMULARIO)
        st.success("Cargada: ve a «Crear nueva solicitud».")

    with st.expander("Texto del correo"):
        st.code(sol["texto"] or "", language=None, wrap_lines=True)

    with st.form(f"seguimiento_{sol['id']}"):
        st.markdown("**Seguimiento**")
        estado = st.selectbox("Estado", ESTADOS, index=ESTADOS.index(sol["estado"])
                              if sol["estado"] in ESTADOS else 0)
        comercializadoras = st.text_area("Comercializadoras a las que se ha enviado",
                                         value=sol["comercializadoras"] or "",
                                         placeholder="Una por línea o separadas por comas")
        if st.form_submit_button("Guardar seguimiento"):
            almacen.actualizar_seguimiento(sol["id"], estado, comercializadoras,
                                           (st.session_state.get("usuario") or {}).get(
                                               "usuario", ""))
            st.rerun()
