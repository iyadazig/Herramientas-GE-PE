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
from datetime import date, datetime, time

import pandas as pd
import streamlit as st

from gemweb_extractor import paginas as gw
from ofertas import correo
from ofertas import solicitud as so

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
COLUMNAS_CASILLA = so.PRECIO_INDEXADO + [TEMPORADA, EXENCION, GDO, EXCEDENTES]


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
    return {"indexado": [o for o in so.PRECIO_INDEXADO if bool(fila[o])],
            "inicio": _texto(fila[INICIO]), "duracion": _texto(fila[DURACION]),
            "duracion2": _texto(fila[DURACION2]), "duracion3": _texto(fila[DURACION3]),
            "temporada": bool(fila[TEMPORADA]),
            "gdo": bool(fila[GDO]), "excedentes": bool(fila[EXCEDENTES]),
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
    c_index, c_part = st.columns(2)
    with c_index:
        st.markdown("**Precio indexado**")
        st.caption("El precio fijo se pide siempre según la tarifa: binómica en DH2 para la "
                   "2.0TD y en DH6 para la 3.0TD y las 6.xTD.")
        for o in so.PRECIO_INDEXADO:
            comun[o] = st.checkbox(o, key="sol_op_" + o)
    with c_part:
        st.markdown("**Particularidades**")
        comun[TEMPORADA] = st.checkbox(TEMPORADA, key="sol_temporada",
                                       help="Sin marcar, el contrato es anual.")
        comun[GDO] = st.checkbox("Incluye garantías de origen (GdO)", key="sol_gdo")
        comun[EXCEDENTES] = st.checkbox(EXCEDENTES, key="sol_excedentes")
        comun[EXENCION] = st.checkbox(EXENCION, key="sol_exencion")
        p1, p2 = st.columns(2)
        comun[PCT_CONSUMO] = p1.number_input("Porcentaje de consumo sujeto a exención (%)",
                                             min_value=0.0, max_value=100.0, value=None,
                                             disabled=not comun[EXENCION], key="sol_pct_consumo")
        comun[PCT_TIPO] = p2.number_input("Porcentaje tipo de exención (%)", min_value=0.0,
                                          max_value=100.0, value=None,
                                          disabled=not comun[EXENCION], key="sol_pct_tipo")
    comun[INICIO] = st.date_input("Fecha prevista de inicio del contrato", value=None,
                                  format="DD/MM/YYYY", key="sol_inicio")
    st.caption("Duración del contrato: rellena al menos una fecha fin; cada una es una "
               "alternativa de plazo para la oferta.")
    c1, c2, c3 = st.columns(3)
    for col, clave, n in ((c1, DURACION, ""), (c2, DURACION2, "2"), (c3, DURACION3, "3")):
        comun[clave] = col.date_input(f"Duración del contrato hasta {n}".strip(), value=None,
                                      format="DD/MM/YYYY", key="sol_duracion" + n)
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
    st.markdown("**Periodo para la recepción de las ofertas**")
    c1, c2, c3 = st.columns(3)
    desde = c1.date_input("Desde", value=date.today(), format="DD/MM/YYYY", key="sol_rec_desde")
    hasta = c2.date_input("Hasta (fecha límite)", value=None, format="DD/MM/YYYY",
                          key="sol_rec_hasta")
    hora = c3.time_input("Hora límite", value=time(12, 0), key="sol_rec_hora", step=900)
    if desde and hasta and hasta < desde:
        st.warning("La fecha límite de recepción es anterior a la de inicio.")

    st.caption("Revisa los datos de cada suministro (salen de Gemweb).")
    tabla = pd.DataFrame([{v: f.get(k, "") for k, v in COLUMNAS_FICHA.items()} for f in fichas])
    editada = st.data_editor(tabla, hide_index=True, use_container_width=True,
                             disabled=["CUPS", "Tarifa"], key="sol_tabla_%d" % hash(res["clave"]))
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
    for f, (_, fila), (_, fila_c), (_, fila_p) in zip(
            fichas, editada.iterrows(), editada_c.iterrows(), potencias.iterrows()):
        s = dict(f, **{inversa[col]: _texto(valor) for col, valor in fila.items()})
        s["potencias"] = [None if _texto(fila_p[p]) == "" else float(fila_p[p])
                          for p in so.PERIODOS]
        s["consideraciones"] = consideraciones_de(fila_c)
        s["fecha_inicio"] = s["consideraciones"]["inicio"]
        suministros.append(s)
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
                "gestor": gestor, "telefono": telefono, "email": email, "producto": producto,
                "recepcion_desde": desde.strftime("%d/%m/%Y") if desde else "",
                "recepcion_hasta": hasta.strftime("%d/%m/%Y") if hasta else "",
                "recepcion_hora": hora.strftime("%H:%M") if hora else ""}
    if not hasta:
        st.info("Elige la fecha límite de recepción de ofertas (arriba, en «Periodo para la "
                "recepción de las ofertas») para poder descargar la solicitud y el texto.")
    listo = bool(hasta) and not sin_plazo
    nombre = "solicitud_ofertas_%s.xlsx" % (re.sub(r"[^\w-]", "_", ref) if ref else
                                             date.today().strftime("%d%m%Y"))
    b1, b2 = st.columns(2)
    b1.download_button("Descargar solicitud de ofertas",
                       data=so.excel_solicitud(cabecera, suministros) if listo else b"",
                       file_name=nombre, disabled=not listo,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    if b2.button("Texto para el correo a las comercializadoras", disabled=not listo):
        st.session_state.sol_ver_correo = True
    if listo and st.session_state.get("sol_ver_correo"):
        st.markdown("#### Texto para el correo")
        st.caption("Cópialo con el botón de la esquina del recuadro y pégalo en el correo. Se "
                   "actualiza solo si cambias algo de arriba.")
        st.code(correo.texto_correo(cabecera, suministros), language=None, wrap_lines=True)
