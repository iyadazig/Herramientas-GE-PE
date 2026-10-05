# -*- coding: utf-8 -*-
"""
API DE GEMWEB — credenciales y curvas para la revision de SSAA
================================================================
Usa el cliente comun de la plataforma (gemweb_extractor/cliente.py, el de las
paginas de Gemweb), con un solo objeto por credenciales. Para la revision de SSAA:
  buscar_suministro: CUPS -> id interno y datos del suministro (get_inventory)
  curva:             kWh de cada cuarto de hora (get_metering, quart-horari, consum)

Formato de la API: XML; fechas "AAAA-MM-DD  HH:MM" con la hora de FIN del
cuarto (la primera del dia es 00:15) y 96 cuartos por dia tambien en los dias
de cambio de hora. curva_desde_gemweb() lo pasa a la convencion de ESIOS.

CREDENCIALES (nunca en el repositorio). Se buscan por este orden:
  1. variables de entorno GEMWEB_CLIENT_ID y GEMWEB_CLIENT_SECRET
  2. %APPDATA%\\ValidadorSSAA\\gemweb.json  (guardadas desde la app; secreto
     cifrado con DPAPI, solo legible por el mismo usuario de Windows)
  3. %APPDATA%\\EstudioPotencia\\gemweb.json (las del programa de potencia)
  4. ..\\API_Gemweb\\.streamlit\\secrets.toml
"""

import base64
import datetime as dt
import hashlib
import json
import os
import threading
from pathlib import Path

import config
import curva_consumo
from gemweb_extractor.cliente import GemwebClient, GemwebError   # noqa: F401 (se reexporta)


# ---------------------------------------------------------------- credenciales
def _appdata():
    return Path(os.environ.get("APPDATA") or Path.home())


def _ruta_propia():
    return _appdata() / "ValidadorSSAA" / "gemweb.json"


def _dpapi(datos, cifrar):
    """Cifra / descifra con la proteccion de datos de Windows (mismo usuario)."""
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    entrada = BLOB(len(datos), ctypes.create_string_buffer(datos, len(datos)))
    salida = BLOB()
    fn = ctypes.windll.crypt32.CryptProtectData if cifrar else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(salida)):
        raise GemwebError("No se ha podido cifrar/descifrar la clave de Gemweb.")
    try:
        return ctypes.string_at(salida.pbData, salida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(salida.pbData)


def guardar_credenciales(client_id, client_secret):
    ruta = _ruta_propia()
    ruta.parent.mkdir(parents=True, exist_ok=True)
    try:
        secreto = {"dpapi": base64.b64encode(_dpapi(client_secret.encode("utf-8"), True)).decode("ascii")}
    except Exception:
        secreto = {"texto": client_secret}      # fuera de Windows
    ruta.write_text(json.dumps({"client_id": client_id, "client_secret": secreto}), encoding="utf-8")


def borrar_credenciales():
    if _ruta_propia().exists():
        _ruta_propia().unlink()


def _leer_json(ruta):
    d = json.loads(ruta.read_text(encoding="utf-8"))
    s = d["client_secret"]
    secreto = _dpapi(base64.b64decode(s["dpapi"]), False).decode("utf-8") if "dpapi" in s else s["texto"]
    return d["client_id"], secreto


def _leer_toml(ruta):
    import tomllib
    with open(ruta, "rb") as f:
        s = tomllib.load(f)
    return s["GEMWEB_CLIENT_ID"], s["GEMWEB_CLIENT_SECRET"]


def cargar_credenciales():
    """((client_id, client_secret), origen) o (None, None)."""
    if os.environ.get("GEMWEB_CLIENT_ID") and os.environ.get("GEMWEB_CLIENT_SECRET"):
        return (os.environ["GEMWEB_CLIENT_ID"], os.environ["GEMWEB_CLIENT_SECRET"]), \
            "variables de entorno"
    fuentes = [(_ruta_propia(), _leer_json, "guardadas en esta app"),
               (_appdata() / "EstudioPotencia" / "gemweb.json", _leer_json,
                "programa de estudio de potencia"),
               (config.CARPETA.parent / "API_Gemweb" / ".streamlit" / "secrets.toml",
                _leer_toml, "proyecto API_Gemweb")]
    for ruta, lector, origen in fuentes:
        if ruta.exists():
            try:
                cred = lector(ruta)
            except Exception:
                continue
            if cred[0] and cred[1]:
                return cred, origen
    return None, None


# ---------------------------------------------------------------------- cliente
_apis = {}
_candado_apis = threading.Lock()


def obtener_api(client_id, client_secret):
    """Cliente de la API (gemweb_extractor.cliente.GemwebClient) compartido por toda la
    plataforma: un solo token por credenciales, aunque haya varios compañeros a la vez."""
    clave = (client_id, hashlib.sha256(client_secret.encode("utf-8")).hexdigest())
    with _candado_apis:
        if clave not in _apis:
            _apis[clave] = GemwebClient(client_id, client_secret)
        return _apis[clave]


class ClienteGemweb:
    """Operaciones que usa la revision de SSAA (CUPS -> suministro, curva cuartohoraria)
    sobre el cliente comun de la API."""

    def __init__(self, client_id=None, client_secret=None, api=None):
        self.api = api or obtener_api(client_id, client_secret)

    @classmethod
    def desde_configuracion(cls):
        cred, _origen = cargar_credenciales()
        if cred is None:
            raise GemwebError("No hay credenciales de Gemweb configuradas.")
        return cls(*cred)

    def comprobar(self):
        """Pide un token para validar las credenciales."""
        import requests
        try:
            self.api._renovar_token()
        except requests.RequestException as e:
            raise GemwebError("No se ha podido conectar con Gemweb: %s" % e) from e
        except GemwebError as e:
            raise GemwebError("Credenciales de Gemweb no válidas: %s" % e) from e

    def buscar_suministro(self, cups):
        """Datos del suministro en el inventario de Gemweb (dict) o None."""
        cups = cups.strip().upper()
        candidatos = [cups] + ([cups[:20]] if len(cups) > 20 else [])
        for c in candidatos:
            try:
                raiz = self.api.get_inventory(category="subministraments",
                                              search_by="subministraments.cups",
                                              search_values=c, limit=5)
            except GemwebError as e:
                if "no se han encontrado" in str(e).lower():
                    continue
                raise
            filas = [{h.tag: h.text for h in nodo} for nodo in raiz]
            filas = [f for f in filas if f.get("id")]
            if filas:
                # si hay varios (altas y bajas), el que sigue de alta
                activos = [f for f in filas if not f.get("data_baixa")]
                return (activos or filas)[0]
        return None

    def descargar(self, id_suministro, desde, hasta, al_avanzar=None):
        """[(fecha 'AAAA-MM-DD HH:MM' fin del cuarto, kWh)] y tramos fallidos."""
        df, fallidos = self.api.get_metering_por_tramos(
            id=int(id_suministro), date_from=desde.isoformat(), date_to=hasta.isoformat(),
            data_source="comptador", period="quart-horari", field="consum",
            al_avanzar=al_avanzar)
        valores = {}
        for fila in df.itertuples(index=False) if len(df) else []:
            fecha = " ".join(str(fila.fecha or "").split())
            if fecha:
                factor = {"kwh": 1.0, "wh": 0.001, "mwh": 1000.0}.get(
                    str(fila.unidades or "kWh").strip().lower(), 1.0)
                valores[fecha] = float(fila.valor) * factor        # sin duplicar solapes
        return sorted(valores.items()), fallidos

    def curva(self, cups, desde, hasta, al_avanzar=None):
        """Curva (convencion ESIOS) del CUPS entre dos fechas incluidas, y el suministro."""
        sum_ = self.buscar_suministro(cups)
        if sum_ is None:
            raise GemwebError("El CUPS %s no está en el inventario de Gemweb." % cups)
        # la API etiqueta con el fin del cuarto: se pide hasta el dia siguiente
        valores, fallidos = self.descargar(sum_["id"], desde, hasta + dt.timedelta(days=1),
                                           al_avanzar)
        if not valores:
            motivo = fallidos[0].split(": ", 1)[-1] if fallidos else "respuesta vacía"
            raise GemwebError("Gemweb no tiene curva de %s entre el %s y el %s (%s). El "
                              "suministro existe en el inventario, pero sin telemedida "
                              "cargada; sube la curva como fichero."
                              % (cups, desde.strftime("%d/%m/%Y"), hasta.strftime("%d/%m/%Y"),
                                 motivo))
        c = curva_desde_gemweb(valores, desde, hasta)
        c.avisos = ["Tramo sin descargar: " + f for f in fallidos] + c.avisos
        return c, sum_


# ------------------------------------------------------------------ conversion
def curva_desde_gemweb(valores, desde=None, hasta=None):
    """[(fecha fin del cuarto, kWh)] -> Curva cuartohoraria con claves ESIOS."""
    regs = []
    for fecha, kwh in valores:
        fin = dt.datetime.strptime(fecha, "%Y-%m-%d %H:%M")
        ini = fin - dt.timedelta(minutes=15)
        if (desde and ini.date() < desde) or (hasta and ini.date() > hasta):
            continue
        regs.append((ini.date(), ini.hour * 60 + ini.minute, kwh))
    out, avisos = curva_consumo.curva_reloj_a_esios(regs, "qh")
    return curva_consumo.Curva(resolucion="qh", valores=out, avisos=avisos,
                               columnas={"origen": "API Gemweb (consum, quart-horari)"},
                               origen="Gemweb")
