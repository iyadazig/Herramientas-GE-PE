# -*- coding: utf-8 -*-
"""
Correo desde la cuenta de Outlook (Microsoft 365) de cada usuario
=================================================================
Microsoft Graph con permisos delegados y el flujo de «codigo de dispositivo»: no hay secreto
ni hay que instalar nada. Un administrador registra una vez la aplicacion en Microsoft Entra
(servidor/GUIA_OUTLOOK.md) y guarda su ID de inquilino y de aplicacion en Administracion.

Cada usuario pulsa «Conectar mi Outlook»: la app le da un codigo, entra en
microsoft.com/devicelogin con su cuenta y la app guarda su refresh token cifrado con DPAPI
(almacen.sgc_outlook). Los correos salen de su cuenta y quedan en sus Elementos enviados.
"""

import base64
import threading
import time

import requests

import almacen

LOGIN = "https://login.microsoftonline.com/%s/oauth2/v2.0/"
GRAPH = "https://graph.microsoft.com/v1.0"
PERMISOS = "offline_access User.Read Mail.Send"
LIMITE_DIRECTO = 3 * 1024 * 1024        # mas adjuntos que esto: borrador + sesion de subida
TROZO = 320 * 1024 * 10                 # 3,2 MB (multiplo de 320 KiB, como pide Graph)
TIEMPO = 60

_tokens = {}                            # usuario -> (access token, caduca)
_candado = threading.Lock()


class CorreoError(Exception):
    pass


# ------------------------------------------------------------------- configuracion
def configuracion():
    """(tenant, client_id) guardados en Administracion, o (None, None)."""
    return almacen.ajuste_sgc("outlook_tenant"), almacen.ajuste_sgc("outlook_cliente")


def guardar_configuracion(tenant, cliente):
    almacen.guardar_ajuste_sgc("outlook_tenant", (tenant or "").strip() or None)
    almacen.guardar_ajuste_sgc("outlook_cliente", (cliente or "").strip() or None)


def _config():
    tenant, cliente = configuracion()
    if not tenant or not cliente:
        raise CorreoError("Falta configurar la conexión con Outlook: un administrador tiene que "
                          "guardar el ID de inquilino y el ID de aplicación en Administración → "
                          "«Seguimiento y Control» → Outlook.")
    return tenant, cliente


def _cifrar(texto):
    try:
        from gemweb import _dpapi
        return "dpapi:" + base64.b64encode(_dpapi(texto.encode("utf-8"), True)).decode("ascii")
    except Exception:                                     # noqa: BLE001  fuera de Windows
        return "texto:" + texto


def _descifrar(guardado):
    tipo, _, valor = guardado.partition(":")
    if tipo == "dpapi":
        from gemweb import _dpapi
        return _dpapi(base64.b64decode(valor), False).decode("utf-8")
    return valor


# ------------------------------------------------------------------- conexion
def iniciar_conexion():
    """Pide un codigo de dispositivo. dict con user_code, verification_uri, device_code,
    interval y expires_in (lo que hay que enseñar al usuario y guardar para completar)."""
    tenant, cliente = _config()
    r = requests.post(LOGIN % tenant + "devicecode",
                      data={"client_id": cliente, "scope": PERMISOS}, timeout=TIEMPO)
    datos = r.json()
    if r.status_code != 200:
        raise CorreoError("Microsoft no ha dado el código: %s" % (
            datos.get("error_description") or datos.get("error") or r.status_code))
    datos["inicio"] = time.time()
    return datos


def completar_conexion(usuario, codigo, esperar=60):
    """Espera a que el usuario entre con el codigo (hasta `esperar` segundos). Devuelve la
    cuenta conectada o None si aun no ha entrado; CorreoError si caduca o lo rechaza."""
    tenant, cliente = _config()
    limite = time.time() + esperar
    intervalo = max(int(codigo.get("interval") or 5), 1)
    while True:
        r = requests.post(LOGIN % tenant + "token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": cliente, "device_code": codigo["device_code"]}, timeout=TIEMPO)
        datos = r.json()
        if r.status_code == 200:
            return _guardar_tokens(usuario, datos)
        error = datos.get("error")
        if error == "authorization_pending":
            pass
        elif error == "slow_down":
            intervalo += 5
        elif error == "expired_token":
            raise CorreoError("El código ha caducado. Vuelve a pulsar «Conectar mi Outlook».")
        elif error == "authorization_declined":
            raise CorreoError("Se ha cancelado el acceso en la página de Microsoft.")
        else:
            raise CorreoError("Microsoft ha rechazado la conexión: %s" % (
                datos.get("error_description") or error))
        if time.time() + intervalo > limite:
            return None
        time.sleep(intervalo)


def _guardar_tokens(usuario, datos):
    with _candado:
        _tokens[usuario] = (datos["access_token"], time.time() + int(datos.get("expires_in", 3600))
                            - 120)
    cuenta = _cuenta(datos["access_token"])
    almacen.guardar_outlook_sgc(usuario, cuenta, _cifrar(datos["refresh_token"]))
    return cuenta


def _cuenta(token):
    r = requests.get(GRAPH + "/me", headers={"Authorization": "Bearer " + token}, timeout=TIEMPO)
    if r.status_code != 200:
        raise CorreoError("No se ha podido leer la cuenta de Outlook (%s)." % r.status_code)
    d = r.json()
    return d.get("mail") or d.get("userPrincipalName")


def conexion(usuario):
    """Cuenta de Outlook conectada del usuario o None."""
    c = almacen.outlook_sgc(usuario)
    return c["cuenta"] if c else None


def desconectar(usuario):
    with _candado:
        _tokens.pop(usuario, None)
    almacen.guardar_outlook_sgc(usuario, None, None)


def token_de(usuario):
    """Access token valido del usuario; lo renueva con el refresh token si hace falta."""
    with _candado:
        token, caduca = _tokens.get(usuario, (None, 0))
    if token and caduca > time.time():
        return token
    guardado = almacen.outlook_sgc(usuario)
    if not guardado:
        raise CorreoError("Conecta tu Outlook antes de enviar.")
    tenant, cliente = _config()
    r = requests.post(LOGIN % tenant + "token", data={
        "grant_type": "refresh_token", "client_id": cliente, "scope": PERMISOS,
        "refresh_token": _descifrar(guardado["token"])}, timeout=TIEMPO)
    datos = r.json()
    if r.status_code != 200:
        if datos.get("error") == "invalid_grant":
            desconectar(usuario)
            raise CorreoError("La conexión con tu Outlook ha caducado o se ha retirado. Vuelve a "
                              "pulsar «Conectar mi Outlook».")
        raise CorreoError("No se ha podido renovar el acceso a Outlook: %s" % (
            datos.get("error_description") or datos.get("error")))
    with _candado:
        _tokens[usuario] = (datos["access_token"],
                            time.time() + int(datos.get("expires_in", 3600)) - 120)
    if datos.get("refresh_token"):
        almacen.guardar_outlook_sgc(usuario, guardado["cuenta"], _cifrar(datos["refresh_token"]))
    return datos["access_token"]


# ------------------------------------------------------------------- envio
def _destinatarios(texto):
    return [{"emailAddress": {"address": e.strip()}}
            for e in str(texto or "").replace(",", ";").split(";") if e.strip()]


def _error(r, accion):
    try:
        d = r.json().get("error", {})
        detalle = d.get("message") or d.get("code")
    except ValueError:
        detalle = r.text[:200]
    return CorreoError("%s: %s (%s)" % (accion, detalle, r.status_code))


def enviar(usuario, para, cc, asunto, html, adjuntos=()):
    """Envia un correo HTML desde la cuenta del usuario. adjuntos: [(nombre, bytes)]."""
    if not _destinatarios(para):
        raise CorreoError("El correo no tiene destinatarios en «Para».")
    cabeceras = {"Authorization": "Bearer " + token_de(usuario)}
    mensaje = {"subject": asunto, "body": {"contentType": "HTML", "content": html},
               "toRecipients": _destinatarios(para), "ccRecipients": _destinatarios(cc)}
    total = sum(len(d) for _, d in adjuntos)
    if total <= LIMITE_DIRECTO:
        mensaje["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment", "name": nombre,
            "contentType": "application/pdf" if nombre.lower().endswith(".pdf")
            else "application/octet-stream",
            "contentBytes": base64.b64encode(datos).decode("ascii")} for nombre, datos in adjuntos]
        r = requests.post(GRAPH + "/me/sendMail", headers=cabeceras, timeout=TIEMPO * 2,
                          json={"message": mensaje, "saveToSentItems": True})
        if r.status_code != 202:
            raise _error(r, "No se ha enviado")
        return
    # adjuntos grandes: borrador, una sesion de subida por adjunto y envio del borrador
    r = requests.post(GRAPH + "/me/messages", headers=cabeceras, json=mensaje, timeout=TIEMPO)
    if r.status_code != 201:
        raise _error(r, "No se ha podido crear el borrador")
    id_mensaje = r.json()["id"]
    for nombre, datos in adjuntos:
        r = requests.post(GRAPH + "/me/messages/%s/attachments/createUploadSession" % id_mensaje,
                          headers=cabeceras, timeout=TIEMPO, json={"AttachmentItem": {
                              "attachmentType": "file", "name": nombre, "size": len(datos)}})
        if r.status_code != 201:
            raise _error(r, "No se ha podido adjuntar %s" % nombre)
        url = r.json()["uploadUrl"]
        for inicio in range(0, len(datos), TROZO):
            trozo = datos[inicio:inicio + TROZO]
            r = requests.put(url, data=trozo, timeout=TIEMPO * 2, headers={
                "Content-Length": str(len(trozo)),
                "Content-Range": "bytes %d-%d/%d" % (inicio, inicio + len(trozo) - 1,
                                                     len(datos))})
            if r.status_code not in (200, 201):
                raise _error(r, "No se ha podido subir %s" % nombre)
    r = requests.post(GRAPH + "/me/messages/%s/send" % id_mensaje, headers=cabeceras,
                      timeout=TIEMPO)
    if r.status_code != 202:
        raise _error(r, "No se ha enviado")
