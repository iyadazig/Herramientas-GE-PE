# -*- coding: utf-8 -*-
"""
Paginas de envio: informes de seguimiento, avisos de incidencias e historial
============================================================================
Los correos salen de la cuenta de Outlook del usuario conectado (correo.py). Siempre hay un
modo prueba (activado por defecto) que no envia nada y da el Excel de lo que se enviaria,
como los programas de escritorio.
"""

import datetime as dt
import io
import time

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import almacen
import estilo
from seguimiento import correo
from seguimiento import datos as D
from seguimiento import incidencias as INC
from seguimiento import informes as INF

MODULO = "Seguimiento y Control"
ss = st.session_state
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _usuario():
    return ss.usuario["usuario"]


def _excel(filas):
    salida = io.BytesIO()
    pd.DataFrame(filas).to_excel(salida, index=False)
    return salida.getvalue()


# ------------------------------------------------------------------- outlook
def bloque_outlook():
    """Estado de la conexion con Outlook del usuario y botones para conectar/desconectar.
    Devuelve la cuenta conectada o None."""
    tenant, cliente = correo.configuracion()
    if not tenant or not cliente:
        st.warning("Los envíos reales aún no están disponibles: falta registrar la aplicación de "
                   "Outlook (Administración → «Seguimiento y Control» → Outlook). Mientras tanto "
                   "se puede usar el modo prueba.", icon=":material/mail_lock:")
        return None
    cuenta = correo.conexion(_usuario())
    if cuenta:
        c1, c2 = st.columns([4, 1])
        c1.success("Los correos saldrán de tu cuenta **%s** y quedarán en tus Elementos "
                   "enviados." % cuenta, icon=":material/mail:")
        if c2.button("Desconectar", key="sgc_outlook_fuera", use_container_width=True):
            correo.desconectar(_usuario())
            st.rerun()
        return cuenta
    st.info("Para enviar, conecta tu cuenta de Outlook (solo la primera vez).",
            icon=":material/mail:")
    codigo = ss.get("sgc_codigo_outlook")
    if not codigo or time.time() - codigo["inicio"] > int(codigo.get("expires_in", 900)):
        if st.button("Conectar mi Outlook", type="primary", key="sgc_outlook_conectar"):
            try:
                ss.sgc_codigo_outlook = correo.iniciar_conexion()
            except (correo.CorreoError, OSError) as e:
                st.error(str(e))
            st.rerun()
        return None
    st.markdown("1. Abre **[%s](%s)** (se abre en otra pestaña).\n"
                "2. Escribe este código y entra con tu cuenta de GEYPE:"
                % (codigo["verification_uri"], codigo["verification_uri"]))
    st.code(codigo["user_code"], language=None)
    st.markdown("3. Cuando Microsoft diga que ya puedes cerrar la ventana, pulsa:")
    if st.button("Ya he entrado", type="primary", key="sgc_outlook_completar"):
        with st.spinner("Esperando a Microsoft…"):
            try:
                cuenta = correo.completar_conexion(_usuario(), codigo, esperar=30)
            except (correo.CorreoError, OSError) as e:
                ss.pop("sgc_codigo_outlook", None)
                st.error(str(e))
                return None
        if cuenta:
            ss.pop("sgc_codigo_outlook", None)
            almacen.registrar_actividad(_usuario(), MODULO, "Conectar Outlook", cuenta)
            st.rerun()
        st.warning("Microsoft aún no confirma la entrada. Termina en su página y vuelve a "
                   "pulsar «Ya he entrado».")
    return None


def admin_outlook():
    st.caption("Datos de la aplicación registrada en Microsoft Entra para enviar desde la cuenta "
               "de cada usuario (guía paso a paso en servidor/GUIA_OUTLOOK.md). No son "
               "contraseñas: la aplicación no tiene secreto y cada usuario entra con la suya.")
    tenant, cliente = correo.configuracion()
    with st.form("sgc_outlook_config"):
        nuevo_tenant = st.text_input("ID de directorio (inquilino)", value=tenant or "")
        nuevo_cliente = st.text_input("ID de aplicación (cliente)", value=cliente or "")
        if st.form_submit_button("Guardar", type="primary"):
            correo.guardar_configuracion(nuevo_tenant, nuevo_cliente)
            st.success("Guardado.")
            st.rerun()


# ------------------------------------------------------------------- comun
def _auxiliar(clave):
    """Selector de auxiliar: por defecto el del usuario conectado (tabla Equipo)."""
    equipo = almacen.listar_sgc_equipo()
    auxiliares = [e["nombre"] for e in equipo if e["auxiliar"]]
    if not auxiliares:
        st.warning("No hay auxiliares en el equipo (Administración → «Seguimiento y Control»).")
        return None, {}
    propio = next((e["nombre"] for e in equipo
                   if e["auxiliar"] and e.get("usuario") == _usuario()), None)
    auxiliar = st.selectbox("Auxiliar", auxiliares,
                            index=auxiliares.index(propio) if propio else 0, key=clave,
                            help="Se envían los correos de los clientes de este auxiliar. El "
                                 "correo sale siempre de tu cuenta.")
    emails = {e["nombre"]: e["email"] for e in equipo if e.get("email")}
    return auxiliar, emails


def _contactos():
    return [f["datos"] for f in almacen.listar_sgc_filas("contactos_envio")]


def _vista_previa(envios, clave):
    if not envios:
        return
    nombres = ["%s · %s" % (e["grupo"], e["asunto"]) for e in envios]
    elegido = st.selectbox("Ver un correo", range(len(envios)), format_func=nombres.__getitem__,
                           key=clave)
    e = envios[elegido]
    st.markdown("**Para:** %s  \n**CC:** %s  \n**Asunto:** %s" % (e["para"] or "—",
                                                                    e["cc"] or "—", e["asunto"]))
    if e.get("adjuntos"):
        st.caption("Adjuntos: " + ", ".join(n for n, _ in e["adjuntos"]))
    components.html(e["html"], height=420, scrolling=True)


def _enviar(envios, tipo, auxiliar, cuenta, detalle):
    """Envia los correos uno a uno. Devuelve la lista de resultados (dicts)."""
    resultados = []
    barra = st.progress(0.0, text="Enviando…")
    for n, e in enumerate(envios, start=1):
        try:
            correo.enviar(_usuario(), e["para"], e["cc"], e["asunto"], e["html"],
                          e.get("adjuntos") or ())
            resultado = "Enviado"
        except (correo.CorreoError, OSError) as error:
            resultado = "ERROR: %s" % error
        almacen.guardar_envio_sgc(_usuario(), cuenta, tipo, auxiliar, e["grupo"], e["para"],
                                  e["cc"], e["asunto"], detalle(e), False, resultado)
        resultados.append({"Cliente/Grupo": e["grupo"], "Para": e["para"], "Asunto": e["asunto"],
                           "Resultado": resultado, "_envio": e})
        barra.progress(n / len(envios), text="Enviados %d de %d" % (n, len(envios)))
    almacen.registrar_actividad(_usuario(), MODULO, "Enviar %s" % tipo, "%d correos, %d con error"
                                % (len(resultados),
                                   sum(r["Resultado"] != "Enviado" for r in resultados)))
    return resultados


def _confirmar_y_enviar(envios, tipo, auxiliar, cuenta, detalle, clave):
    """Casilla de modo prueba, Excel de prueba y boton de envio real. Devuelve los resultados
    del envio real o None."""
    prueba = st.toggle("Modo prueba (no se envía nada)", value=True, key=clave + "_prueba")
    if prueba:
        st.caption("En modo prueba solo se genera el Excel con lo que se enviaría.")
        return None
    if not cuenta:
        st.warning("Conecta tu Outlook (arriba) para enviar.")
        return None
    st.warning("Vas a enviar **%d correos reales** desde %s. Revisa los destinatarios antes de "
               "continuar." % (len(envios), cuenta), icon=":material/warning:")
    revisado = st.checkbox("He revisado los destinatarios", key=clave + "_ok")
    if st.button("Enviar %d correos" % len(envios), type="primary", disabled=not revisado,
                 icon=":material/send:", key=clave + "_enviar"):
        return _enviar(envios, tipo, auxiliar, cuenta, detalle)
    return None


def _mostrar_resultados(resultados):
    df = pd.DataFrame([{k: v for k, v in r.items() if k != "_envio"} for r in resultados])
    errores = (df["Resultado"] != "Enviado").sum()
    (st.success if not errores else st.warning)(
        "Enviados %d de %d correos." % (len(df) - errores, len(df)))
    st.dataframe(df, hide_index=True, use_container_width=True)


# ------------------------------------------------------------------- informes
def pagina_informes():
    cuenta = bloque_outlook()
    auxiliar, emails = _auxiliar("sgc_inf_aux")
    if not auxiliar:
        return
    subidos = st.file_uploader(
        "Informes en PDF (puedes seleccionar todos los de la carpeta a la vez)", type=["pdf"],
        accept_multiple_files=True, key="sgc_inf_pdfs",
        help="Nombre: AAAAMMDD_AAAAMMDD_Inf_SyC_Denominación.pdf, opcionalmente con el tipo de "
             "factura al final (…_Refacturacion.pdf).")
    st.caption("Formatos válidos: 20260101_20260131_Inf_SyC_Desaladora.pdf · "
               "20260101_20260131_Inf_SyC_Desaladora_Refacturacion.pdf")
    if not subidos:
        return
    envios, sin_casar, aviso = INF.preparar([(f.name, f.getvalue()) for f in subidos],
                                            _contactos(), auxiliar, emails)
    if aviso:
        st.warning(aviso)
    if sin_casar:
        with st.expander(":material/warning: %d PDF sin destinatario" % len(sin_casar),
                         expanded=True):
            st.caption("Su nombre no tiene el formato o su denominación no coincide con ningún "
                       "suministro de «Contactos de envío» de %s. No se enviarán." % auxiliar)
            st.dataframe(pd.DataFrame({"PDF": sin_casar}), hide_index=True)
    if not envios:
        st.info("No hay informes que coincidan con los contactos de envío de %s." % auxiliar)
        return
    st.markdown("**%d correos** con %d informes" % (len(envios),
                                                    sum(len(e["informes"]) for e in envios)))
    filas = INF.filas_prueba(envios, auxiliar)
    st.dataframe(pd.DataFrame(filas).drop(columns=["Informes"]), hide_index=True,
                 use_container_width=True)
    _vista_previa(envios, "sgc_inf_ver")
    estilo.boton_descarga(MODULO, st, "Descargar Excel del modo prueba", _excel(filas),
                          "modo_prueba_informes_%s_%s.xlsx" % (
                              auxiliar, dt.datetime.now().strftime("%Y%m%d_%H%M")), XLSX)
    resultados = _confirmar_y_enviar(
        envios, "informe", auxiliar, cuenta,
        lambda e: ", ".join(n for n, _ in e["adjuntos"]), "sgc_inf")
    if resultados is None:
        return
    _mostrar_resultados(resultados)
    meses = []
    for r in resultados:
        if r["Resultado"] == "Enviado":
            meses += INF.meses_a_marcar(r["_envio"])
    marcados, sin_marcar = D.marcar_enviados(meses, _usuario())
    st.success("Marcados como «Inf enviado» en el cuadro: %d meses." % len(marcados))
    if sin_marcar:
        with st.expander("%d meses sin marcar" % len(sin_marcar)):
            st.caption("Solo se marcan los meses en «Recibida», «Pdte inf» o «Inf no enviado». "
                       "Revísalos en el cuadro.")
            st.dataframe(pd.DataFrame(sin_marcar), hide_index=True, use_container_width=True)


# ------------------------------------------------------------------- incidencias
def pagina_incidencias():
    cuenta = bloque_outlook()
    auxiliar, emails = _auxiliar("sgc_av_aux")
    if not auxiliar:
        return
    primero = st.radio("¿Primer envío del mes?", ["Sí", "No"], horizontal=True, key="sgc_av_1")
    if primero == "Sí":
        fichero = st.file_uploader("Excel de facturas", type=["xls", "xlsx"], key="sgc_av_fact")
        if not fichero:
            return
        try:
            facturas = INC.leer_excel(fichero)
        except Exception as e:                                       # noqa: BLE001
            st.error("No se ha podido leer el Excel: %s" % e)
            return
    else:
        c1, c2 = st.columns(2)
        anterior = c1.file_uploader("Excel de incidencias ANTERIOR", type=["xls", "xlsx"],
                                    key="sgc_av_ant")
        actual = c2.file_uploader("Excel de incidencias ACTUAL", type=["xls", "xlsx"],
                                  key="sgc_av_act")
        if not (anterior and actual):
            return
        try:
            facturas = INC.incidencias_nuevas(INC.leer_excel(anterior), INC.leer_excel(actual))
        except Exception as e:                                       # noqa: BLE001
            st.error(str(e))
            return
        if facturas is None:
            st.info("No hay incidencias nuevas.")
            return
        st.caption("%d facturas nuevas respecto al Excel anterior." % len(facturas))
    try:
        envios, sin_contacto, aviso = INC.preparar(facturas, _contactos(), auxiliar, emails)
    except ValueError as e:
        st.error(str(e))
        return
    if len(sin_contacto):
        with st.expander(":material/warning: %d clientes y %d facturas sin contacto: NO recibirán "
                         "aviso" % (sin_contacto["cliente_grupo"].nunique(), len(sin_contacto)),
                         expanded=True):
            st.caption("Su combinación de técnico, auxiliar, cliente/grupo, denominación y CUPS no "
                       "está en «Contactos de envío».")
            st.dataframe(sin_contacto, hide_index=True, use_container_width=True)
            estilo.boton_descarga(MODULO, st, "Descargar Excel", _excel(sin_contacto),
                                  "avisos_sin_contacto_%s.xlsx"
                                  % dt.datetime.now().strftime("%Y%m%d_%H%M"), XLSX)
    if aviso:
        st.info(aviso)
    if not envios:
        return
    filas = [f for e in envios for f in e["filas"]]
    st.markdown("**%d correos** con %d facturas con incidencias" % (len(envios), len(filas)))
    st.dataframe(pd.DataFrame([{"Cliente/Grupo": e["grupo"], "Para": e["para"], "CC": e["cc"],
                                "Asunto": e["asunto"], "Facturas": e["facturas"]}
                               for e in envios]), hide_index=True, use_container_width=True)
    _vista_previa(envios, "sgc_av_ver")
    estilo.boton_descarga(MODULO, st, "Descargar Excel del modo prueba", _excel(filas),
                          "modo_prueba_%s_%s.xlsx" % (
                              auxiliar, dt.datetime.now().strftime("%Y%m%d_%H%M")), XLSX)
    resultados = _confirmar_y_enviar(
        envios, "aviso", auxiliar, cuenta,
        lambda e: ", ".join(str(f["Número factura"]) for f in e["filas"]), "sgc_av")
    if resultados is not None:
        _mostrar_resultados(resultados)


# ------------------------------------------------------------------- historial
def pagina_historial():
    c1, c2, c3, c4 = st.columns(4)
    desde = c1.date_input("Desde", dt.date.today() - dt.timedelta(days=90), format="DD/MM/YYYY",
                          key="sgc_his_desde")
    hasta = c2.date_input("Hasta", dt.date.today(), format="DD/MM/YYYY", key="sgc_his_hasta")
    df = pd.DataFrame(almacen.listar_envios_sgc(desde.isoformat(), hasta.isoformat()))
    if df.empty:
        st.info("No hay envíos en esas fechas.")
        return
    tipo = c3.multiselect("Tipo", sorted(df["tipo"].unique()), key="sgc_his_tipo")
    texto = c4.text_input("Buscar (cliente, destinatario, asunto)",
                          key="sgc_his_txt").strip().lower()
    if tipo:
        df = df[df["tipo"].isin(tipo)]
    if texto:
        df = df[df[["cliente", "para", "cc", "asunto"]].fillna("").apply(
            lambda c: c.str.lower().str.contains(texto, regex=False)).any(axis=1)]
    df["fecha"] = pd.to_datetime(df["fecha"])
    st.metric("Correos", len(df))
    st.dataframe(df.drop(columns=["id", "prueba"]).rename(columns={
        "fecha": "Fecha", "usuario": "Usuario", "cuenta": "Cuenta", "tipo": "Tipo",
        "auxiliar": "Auxiliar", "cliente": "Cliente/Grupo", "para": "Para", "cc": "CC",
        "asunto": "Asunto", "detalle": "Adjuntos / facturas", "resultado": "Resultado"}),
        hide_index=True, use_container_width=True, column_config={
            "Fecha": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm")})
