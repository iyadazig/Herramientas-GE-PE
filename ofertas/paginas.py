# -*- coding: utf-8 -*-
"""
PAGINAS DE «SOLICITUD DE OFERTAS» de Herramientas GE&PE
=======================================================
Excel de solicitud de ofertas (Resumen + una hoja por suministro) a partir de los datos de
Gemweb: facturas de los ultimos 12 meses, potencias de la ultima factura y datos de la
entidad y del centro, mas las consideraciones de la oferta (tipologias de precio, fechas,
exencion), comunes a todos los suministros y modificables en cada uno. Usa las funciones de
consulta de gemweb_extractor.paginas; el formato del Excel esta en ofertas/solicitud.py.
"""

import re
from datetime import date, datetime

import pandas as pd
import streamlit as st

from gemweb_extractor import paginas as gw
from ofertas import solicitud as so

# columnas de la tabla de datos, en este orden (grupo = division de la entidad en Gemweb,
# razon social = nombre de la entidad, alias = final del nombre del suministro)
COLUMNAS_FICHA = {"grupo": "Cliente/Grupo", "razon": "Razón social", "cif": "CIF",
                  "alias": "Alias", "cups": "CUPS", "direccion": "Dirección de suministro",
                  "tarifa": "Tarifa"}

# tabla de consideraciones de la oferta (una fila por suministro)
INICIO, DURACION, DURACION2 = ("Fecha prevista de inicio", "Duración del contrato hasta",
                               "Duración del contrato hasta 2")
TEMPORADA, EXENCION = "De temporada", "Sujeto a exención de impuesto eléctrico"
PCT_CONSUMO, PCT_TIPO = "% consumo sujeto a exención", "% tipo de exención"
COLUMNAS_FECHA = [INICIO, DURACION, DURACION2]
COLUMNAS_CASILLA = so.PRECIO_FIJO + so.PRECIO_INDEXADO + [TEMPORADA, EXENCION]


def pagina_solicitud_ofertas():
    cliente = gw.obtener_cliente()
    st.markdown("Prepara el Excel de **solicitud de ofertas** de varios suministros a la vez: "
                "datos del suministro, potencias de su última factura y perfil de consumo por "
                "periodos (P1–P6) de los últimos 12 meses, sacados de Gemweb.")
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


def _texto(valor) -> str:
    """Celda de una tabla editada como texto ('' si vacia; fechas dd/mm/aaaa)."""
    if valor is None or valor is pd.NaT or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


def _fecha(texto):
    return datetime.strptime(texto, "%d/%m/%Y").date()


def consideraciones_de(fila) -> dict:
    """Consideraciones de la oferta de un suministro a partir de su fila de la tabla."""
    exento = bool(fila[EXENCION])
    return {"fijo": [o for o in so.PRECIO_FIJO if bool(fila[o])],
            "indexado": [o for o in so.PRECIO_INDEXADO if bool(fila[o])],
            "inicio": _texto(fila[INICIO]), "duracion": _texto(fila[DURACION]),
            "duracion2": _texto(fila[DURACION2]), "temporada": bool(fila[TEMPORADA]),
            "exencion": exento,
            "pct_consumo_exento": (float(fila[PCT_CONSUMO]) if exento and
                                   _texto(fila[PCT_CONSUMO]) else None),
            "pct_tipo_exencion": (float(fila[PCT_TIPO]) if exento and
                                  _texto(fila[PCT_TIPO]) else None)}


def consideraciones_comunes() -> dict:
    """Selecciones para todos los suministros, con la disposicion del formulario de ofertas."""
    st.markdown("#### Consideraciones de la oferta")
    st.caption("Lo que elijas aquí se aplica a todos los suministros. Debajo puedes cambiarlo en "
               "los que necesiten otra cosa.")
    comun = {}
    c_fijo, c_index = st.columns(2)
    with c_fijo:
        st.markdown("**Precio fijo**")
        for o in so.PRECIO_FIJO:
            comun[o] = st.checkbox(o, key="sol_op_" + o)
    with c_index:
        st.markdown("**Precio indexado**")
        for o in so.PRECIO_INDEXADO:
            comun[o] = st.checkbox(o, key="sol_op_" + o)
    c1, c2, c3 = st.columns(3)
    comun[INICIO] = c1.date_input("Fecha prevista de inicio del contrato", value=None,
                                  format="DD/MM/YYYY", key="sol_inicio")
    comun[DURACION] = c2.date_input("Duración del contrato hasta", value=None,
                                    format="DD/MM/YYYY", key="sol_duracion")
    comun[DURACION2] = c3.date_input("Duración del contrato hasta 2 (opcional)", value=None,
                                     format="DD/MM/YYYY", key="sol_duracion2",
                                     help="Para pedir además la oferta a un segundo plazo.")
    c1, c2, c3, c4 = st.columns(4)
    comun[TEMPORADA] = c1.checkbox(TEMPORADA, key="sol_temporada",
                                   help="Sin marcar, el contrato es anual.")
    comun[EXENCION] = c2.checkbox(EXENCION, key="sol_exencion")
    comun[PCT_CONSUMO] = c3.number_input("Porcentaje de consumo sujeto a exención (%)",
                                         min_value=0.0, max_value=100.0, value=None,
                                         disabled=not comun[EXENCION], key="sol_pct_consumo")
    comun[PCT_TIPO] = c4.number_input("Porcentaje tipo de exención (%)", min_value=0.0,
                                      max_value=100.0, value=None,
                                      disabled=not comun[EXENCION], key="sol_pct_tipo")
    return comun


def tabla_consideraciones(fichas: list, comun: dict) -> pd.DataFrame:
    """Una fila por suministro con las selecciones comunes, para editarla."""
    tabla = pd.DataFrame([dict({"CUPS": f["cups"], "Alias": f.get("alias", "")}, **comun)
                          for f in fichas])
    for col in COLUMNAS_FECHA:
        tabla[col] = pd.to_datetime(tabla[col])
    for col in (PCT_CONSUMO, PCT_TIPO):
        tabla[col] = pd.to_numeric(tabla[col])
    return tabla


def seccion_solicitud_ofertas(res: dict) -> None:
    """Cabecera, tablas revisables de los suministros y boton de descarga del Excel."""
    fichas = res.get("fichas") or []
    if not fichas:
        st.warning("No se ha encontrado ningún suministro en Gemweb.")
        return
    st.subheader("4. Solicitud de ofertas")
    st.caption("Excel con el formato de GE&PE: una hoja «Resumen» y otra por suministro con sus "
               "datos, las potencias de su última factura, el perfil de consumo de los últimos "
               "12 meses (cada factura en el mes de su fecha de fin) y las consideraciones de "
               "la oferta.")
    usuario = st.session_state.get("usuario") or {}
    c1, c2, c3 = st.columns(3)
    with c1:
        ref = st.text_input("Ref.", key="sol_ref")
        fecha = st.date_input("Fecha", value=date.today(), format="DD/MM/YYYY", key="sol_fecha",
                              help="También es el inicio del periodo de recepción de ofertas.")
        limite = st.date_input("Fecha límite de recepción de ofertas", value=None,
                               format="DD/MM/YYYY", key="sol_limite")
    with c2:
        gestor = st.text_input("Gestor de cuenta", value=usuario.get("nombre", ""),
                               key="sol_gestor")
        telefono = st.text_input("Teléfono", value=so.TELEFONO_GEYPE, key="sol_tel")
    with c3:
        email = st.text_input("Email", value=usuario.get("email") or "", key="sol_email",
                              help="El del usuario conectado. Si sale vacío, un administrador "
                                   "puede añadirlo en Administración → Usuarios.")
        producto = st.text_input("Producto", key="sol_producto",
                                 placeholder="p. ej. ELECTRICIDAD / ELECTRICITY")

    st.caption("Revisa los datos de cada suministro (salen de Gemweb).")
    tabla = pd.DataFrame([{v: f.get(k, "") for k, v in COLUMNAS_FICHA.items()} for f in fichas])
    editada = st.data_editor(tabla, hide_index=True, use_container_width=True,
                             disabled=["CUPS", "Tarifa"], key="sol_tabla_%d" % hash(res["clave"]))
    with st.expander("Potencias que se usarán (última factura de cada suministro)"):
        st.dataframe(pd.DataFrame([dict({"CUPS": f["cups"], "Última factura": f["ultima_factura"]},
                                        **{p: f["potencias"][n] for n, p in
                                           enumerate(so.PERIODOS)}) for f in fichas]),
                     hide_index=True, use_container_width=True)
    for f in fichas:
        if f.get("aviso"):
            st.warning(f"{f['cups']}: {f['aviso']}. El perfil Ene–Dic mezcla años; usa «Últimos "
                       "12 meses con factura» o un periodo de 12 meses.")

    comun = consideraciones_comunes()
    st.caption("Consideraciones de cada suministro: salen con lo elegido arriba. Si después "
               "cambias algo de arriba, esta tabla vuelve a empezar con ello.")
    columnas = {c: st.column_config.DateColumn(format="DD/MM/YYYY") for c in COLUMNAS_FECHA}
    columnas.update({c: st.column_config.NumberColumn(min_value=0, max_value=100, format="%.2f")
                     for c in (PCT_CONSUMO, PCT_TIPO)})
    columnas.update({c: st.column_config.CheckboxColumn() for c in COLUMNAS_CASILLA})
    editada_c = st.data_editor(
        tabla_consideraciones(fichas, comun), hide_index=True, use_container_width=True,
        disabled=["CUPS", "Alias"], column_config=columnas,
        key="sol_consid_%d_%d" % (hash(res["clave"]), hash(repr(sorted(comun.items())))))

    inversa = {v: k for k, v in COLUMNAS_FICHA.items()}
    suministros = []
    for f, (_, fila), (_, fila_c) in zip(fichas, editada.iterrows(), editada_c.iterrows()):
        s = dict(f, **{inversa[col]: _texto(valor) for col, valor in fila.items()})
        s["consideraciones"] = consideraciones_de(fila_c)
        s["fecha_inicio"] = s["consideraciones"]["inicio"]
        suministros.append(s)
    sin_tipo = [s["cups"] for s in suministros
                if not s["consideraciones"]["fijo"] and not s["consideraciones"]["indexado"]]
    if sin_tipo:
        st.info("Sin ninguna tipología de precio elegida: " + ", ".join(sin_tipo))
    plazos = [s["cups"] for s in suministros if s["consideraciones"]["inicio"] and any(
        s["consideraciones"][d] and _fecha(s["consideraciones"][d]) <= _fecha(
            s["consideraciones"]["inicio"]) for d in ("duracion", "duracion2"))]
    if plazos:
        st.warning("La duración del contrato termina antes de la fecha de inicio en: "
                   + ", ".join(plazos))

    cabecera = {"ref": ref, "fecha": fecha.strftime("%d/%m/%Y") if fecha else "",
                "gestor": gestor, "telefono": telefono, "email": email, "producto": producto,
                "recepcion_desde": fecha.strftime("%d/%m/%Y") if fecha else "",
                "recepcion_hasta": limite.strftime("%d/%m/%Y") if limite else ""}
    if not limite:
        st.info("Elige la fecha límite de recepción de ofertas (arriba, junto a la fecha) para "
                "poder descargar la solicitud.")
    nombre = "solicitud_ofertas_%s.xlsx" % (re.sub(r"[^\w-]", "_", ref) if ref else
                                             date.today().strftime("%d%m%Y"))
    st.download_button("Descargar solicitud de ofertas",
                       data=so.excel_solicitud(cabecera, suministros) if limite else b"",
                       file_name=nombre, disabled=not limite,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
