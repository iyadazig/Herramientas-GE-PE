# -*- coding: utf-8 -*-
"""
PAGINAS DE «SOLICITUD DE OFERTAS» de Herramientas GE&PE
=======================================================
Excel de solicitud de ofertas (Resumen + una hoja por suministro) a partir de los datos de
Gemweb: facturas de los ultimos 12 meses, potencias de la ultima factura y datos de la
entidad y del centro. Usa las funciones de consulta de gemweb_extractor.paginas; el formato
del Excel esta en ofertas/solicitud.py.
"""

import re
from datetime import date

import pandas as pd
import streamlit as st

from gemweb_extractor import paginas as gw
from ofertas import solicitud as so

# columnas de la tabla editable, en este orden (grupo = division de la entidad en Gemweb,
# razon social = nombre de la entidad, alias = final del nombre del suministro)
COLUMNAS_FICHA = {"grupo": "Cliente/Grupo", "razon": "Razón social", "cif": "CIF",
                  "alias": "Alias", "cups": "CUPS", "direccion": "Dirección de suministro",
                  "tarifa": "Tarifa", "fecha_fin": "Fecha fin contrato actual"}


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


def seccion_solicitud_ofertas(res: dict) -> None:
    """Cabecera, tabla revisable de los suministros y boton de descarga del Excel."""
    fichas = res.get("fichas") or []
    if not fichas:
        st.warning("No se ha encontrado ningún suministro en Gemweb.")
        return
    st.subheader("4. Solicitud de ofertas")
    st.caption("Excel con el formato de GE&PE: una hoja «Resumen» y otra por suministro con sus "
               "datos, las potencias de su última factura y el perfil de consumo de los últimos "
               "12 meses (cada factura en el mes de su fecha de fin).")
    usuario = st.session_state.get("usuario") or {}
    c1, c2, c3 = st.columns(3)
    with c1:
        ref = st.text_input("Ref.", key="sol_ref")
        fecha = st.date_input("Fecha", value=date.today(), format="DD/MM/YYYY", key="sol_fecha")
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

    st.caption("Revisa los datos de cada suministro (salen de Gemweb). Gemweb no tiene la fecha "
               "fin del contrato actual: elige una para todos y cambia en la tabla la de los "
               "suministros que tengan otra.")
    comun = st.date_input("Fecha fin del contrato actual (para todos los suministros)",
                          value=None, format="DD/MM/YYYY", key="sol_fecha_fin_comun")
    tabla = pd.DataFrame([{v: (comun if k == "fecha_fin" else f.get(k, ""))
                           for k, v in COLUMNAS_FICHA.items()} for f in fichas])
    tabla["Fecha fin contrato actual"] = pd.to_datetime(tabla["Fecha fin contrato actual"])
    # al cambiar la fecha comun, la tabla vuelve a empezar con ella en todos los suministros
    editada = st.data_editor(
        tabla, hide_index=True, use_container_width=True, disabled=["CUPS", "Tarifa"],
        column_config={"Fecha fin contrato actual": st.column_config.DateColumn(
            format="DD/MM/YYYY")},
        key="sol_tabla_%d_%s" % (hash(res["clave"]), comun))
    with st.expander("Potencias que se usarán (última factura de cada suministro)"):
        st.dataframe(pd.DataFrame([dict({"CUPS": f["cups"], "Última factura": f["ultima_factura"]},
                                        **{p: f["potencias"][n] for n, p in
                                           enumerate(so.PERIODOS)}) for f in fichas]),
                     hide_index=True, use_container_width=True)
    for f in fichas:
        if f.get("aviso"):
            st.warning(f"{f['cups']}: {f['aviso']}. El perfil Ene–Dic mezcla años; usa «Últimos "
                       "12 meses con factura» o un periodo de 12 meses.")
    inversa = {v: k for k, v in COLUMNAS_FICHA.items()}
    suministros = []
    for f, (_, fila) in zip(fichas, editada.iterrows()):
        s = dict(f)
        for col, valor in fila.items():
            if valor is None or valor is pd.NaT or (isinstance(valor, float) and pd.isna(valor)):
                s[inversa[col]] = ""
            elif hasattr(valor, "strftime"):
                s[inversa[col]] = valor.strftime("%d/%m/%Y")
            else:
                s[inversa[col]] = str(valor).strip()
        suministros.append(s)
    cabecera = {"ref": ref, "fecha": fecha.strftime("%d/%m/%Y") if fecha else "",
                "gestor": gestor, "telefono": telefono, "email": email, "producto": producto}
    nombre = "solicitud_ofertas_%s.xlsx" % (re.sub(r"[^\w-]", "_", ref) if ref else
                                             date.today().strftime("%d%m%Y"))
    st.download_button("Descargar solicitud de ofertas",
                       data=so.excel_solicitud(cabecera, suministros), file_name=nombre,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
