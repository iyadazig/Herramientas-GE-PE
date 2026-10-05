"""
Cliente Python para la API de Gemweb (copiado de API_Gemweb/gemweb_client.py).

Unico cliente de Herramientas GE&PE: lo usan las paginas de Gemweb y la descarga de
curvas de la revision de SSAA (gemweb.py). Cambio respecto al original: un candado al
renovar el token, porque en el servidor varios compañeros comparten el mismo objeto.

Gestiona automáticamente el token de autenticación (validez 1h) y expone
un método por cada operación documentada en el manual.

Todas las llamadas devuelven un objeto xml.etree.ElementTree.Element
(la raíz del XML de respuesta). Para convertirlo a DataFrame de pandas,
ver helpers al final del fichero.
"""

import threading
import time
import requests
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
import pandas as pd


class GemwebError(Exception):
    """Error devuelto por la API en una etiqueta <error>."""
    pass


class GemwebClient:
    BASE_URL = "https://api.gemweb.es"
    TIMEOUT = 60  # segundos
    TIMEOUT_METERING = 180  # las telelecturas pueden tardar bastante más
    REINTENTOS = 3
    ESPERAS_REINTENTO = [5, 15]  # segundos antes del 2º y 3er intento

    # Días por tramo al partir rangos largos de telelecturas, según la agrupación.
    # Las agrupaciones mensual/anual/total no se parten (se falsearía el agregado).
    DIAS_TRAMO_METERING = {
        "diari": 92,
        "horari": 31,
        "quart-horari": 31,
        "cinc-min": 7,
        "un-min": 3,
    }

    def __init__(self, client_id: str, client_secret: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self._token = None
        self._token_expira = None
        self._candado = threading.Lock()

    # ---------- Gestión de token ----------

    def _token_valido(self) -> bool:
        if not self._token or not self._token_expira:
            return False
        # Renovamos con 5 minutos de margen para evitar pillarnos justo en el filo
        return datetime.now() < (self._token_expira - timedelta(minutes=5))

    def _renovar_token(self) -> None:
        payload = {
            "request": "get_token",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
        }
        r = requests.post(self.BASE_URL, data=payload, timeout=self.TIMEOUT)
        r.raise_for_status()
        root = ET.fromstring(r.text)

        # Comprobar si la API devolvió un error
        err = root.find("error")
        if err is not None:
            raise GemwebError(f"Error al obtener token: {err.text}")

        token = root.findtext("access_token")
        expires_in = int(root.findtext("expires_in", default="3600"))

        if not token:
            raise GemwebError("La respuesta no contiene access_token")

        self._token = token
        self._token_expira = datetime.now() + timedelta(seconds=expires_in)

    def info_token(self) -> dict:
        """Devuelve info del token actual (útil para mostrar en la UI)."""
        if not self._token:
            return {"valido": False, "minutos_restantes": 0}
        minutos = int((self._token_expira - datetime.now()).total_seconds() / 60)
        return {
            "valido": self._token_valido(),
            "minutos_restantes": max(0, minutos),
            "caduca": self._token_expira,
        }

    # ---------- Llamada genérica ----------

    def _post(self, request_name: str, timeout: int = None, **params) -> ET.Element:
        """
        Llamada genérica. Renueva el token si hace falta y comprueba errores.
        Si el servidor no responde a tiempo, reintenta hasta REINTENTOS veces.
        """
        timeout = timeout or self.TIMEOUT

        for intento in range(1, self.REINTENTOS + 1):
            if not self._token_valido():
                with self._candado:          # un solo compañero pide el token nuevo
                    if not self._token_valido():
                        self._renovar_token()

            # Limpiar valores None (parámetros opcionales no informados)
            payload = {"request": request_name, "access_token": self._token}
            for k, v in params.items():
                if v is not None and v != "":
                    payload[k] = v

            try:
                r = requests.post(self.BASE_URL, data=payload, timeout=timeout)
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                if intento == self.REINTENTOS:
                    raise GemwebError(
                        f"{request_name}: Gemweb no ha respondido tras {self.REINTENTOS} "
                        f"intentos ({type(e).__name__}, espera máxima {timeout}s por intento)"
                    ) from e
                time.sleep(self.ESPERAS_REINTENTO[intento - 1])

        r.raise_for_status()

        # Intentar parsear el XML. Si la API devuelve XML mal formado,
        # capturamos el error y mostramos la respuesta cruda al usuario
        # para que pueda diagnosticar el problema.
        try:
            root = ET.fromstring(r.text)
        except ET.ParseError as e:
            # Guardar respuesta cruda para inspección
            self.ultima_respuesta_cruda = r.text
            raise GemwebError(
                f"La API devolvió un XML mal formado ({e}). "
                f"Esto suele ser un problema temporal del servidor de Gemweb. "
                f"Respuesta recibida (primeros 1000 caracteres):\n\n{r.text[:1000]}"
            ) from e

        # Si la respuesta es solo un <error>, lanzar excepción
        err = root.find("error")
        if err is not None:
            raise GemwebError(f"{request_name}: {err.text}")

        return root

    # ---------- Operaciones de la API ----------

    def get_inventory(
        self,
        category: str,
        search_by: str = None,
        search_values: str = None,
        search_operator: str = None,
        order_by: str = None,
        limit: int = None,
    ) -> ET.Element:
        """
        Inventario: entidades, centros_consumo, suministros (subministraments),
        instalaciones_solares (instalacions_solars) o facturas (factures).
        """
        return self._post(
            "get_inventory",
            category=category,
            search_by=search_by,
            search_values=search_values,
            search_operator=search_operator,
            order_by=order_by,
            limit=limit,
        )

    def get_invoice_inconsistencies(
        self,
        category: str,
        id: int,
        inconsistencies: str = None,
        percentages: str = None,
        language: str = "es",
    ) -> ET.Element:
        """
        Incoherencias en facturas. category: subministraments, instalacions_solars o factures.
        """
        return self._post(
            "get_invoice_inconsistencies",
            category=category,
            id=id,
            inconsistencies=inconsistencies,
            percentages=percentages,
            language=language,
        )

    def get_power_optimization(
        self,
        id: int,
        date_from: str,
        date_to: str,
        language: str = "es",
    ) -> ET.Element:
        """Optimización de potencia contratada para un suministro de electricidad."""
        return self._post(
            "get_power_optimization",
            id=id,
            date_from=date_from,
            date_to=date_to,
            language=language,
        )

    def get_cost_consumption(
        self,
        id: int,
        type: str,
        date_from: str,
        date_to: str,
        source: str = "factures",
        billing_date: int = 0,
    ) -> ET.Element:
        """
        Análisis mensual de coste (type='cost') o consumo (type='consum').
        source: 'factures' o 'telelectures'.
        """
        return self._post(
            "get_cost_consumption",
            id=id,
            type=type,
            date_from=date_from,
            date_to=date_to,
            source=source,
            billing_date=billing_date,
        )

    def get_calculated_invoice(
        self,
        id: int,
        date_to: str,
        date_from: str = None,
        language: str = "es",
    ) -> ET.Element:
        """Genera una factura calculada a partir de telelecturas."""
        return self._post(
            "get_calculated_invoice",
            id=id,
            date_from=date_from,
            date_to=date_to,
            language=language,
        )

    def get_consumption_by_period(
        self,
        id: int,
        period: str,
        date_from: str,
        date_to: str,
        language: str = "es",
    ) -> ET.Element:
        """Consumos agrupados por periodos tarifarios. period: 'mensual' o 'diari'."""
        return self._post(
            "get_consumption_by_period",
            id=id,
            period=period,
            date_from=date_from,
            date_to=date_to,
            language=language,
        )

    def get_contract_simulation(
        self,
        id: int,
        contract_date_from: str,
        contract_date_to: str,
        contract_fields: str,
        contract_values: str,
        date_from: str,
        date_to: str,
        language: str = "es",
    ) -> ET.Element:
        """Simulación de facturación con precios contratados."""
        return self._post(
            "get_contract_simulation",
            id=id,
            contract_date_from=contract_date_from,
            contract_date_to=contract_date_to,
            contract_fields=contract_fields,
            contract_values=contract_values,
            date_from=date_from,
            date_to=date_to,
            language=language,
        )

    def get_metering(
        self,
        id: int,
        date_from: str,
        date_to: str,
        data_source: str = "comptador",
        period: str = "diari",
        field: str = "consum",
        language: str = "es",
    ) -> ET.Element:
        """
        Telelecturas. field: consum, potencia_activa, potencia_maxima, reactiva, co2.
        period: total, anual, mensual, diari, horari, quart-horari, cinc-min, un-min.
        """
        return self._post(
            "get_metering",
            id=id,
            date_from=date_from,
            date_to=date_to,
            data_source=data_source,
            period=period,
            field=field,
            language=language,
            timeout=self.TIMEOUT_METERING,
        )

    def get_metering_por_tramos(
        self,
        id: int,
        date_from: str,
        date_to: str,
        data_source: str = "comptador",
        period: str = "diari",
        field: str = "consum",
        language: str = "es",
        al_avanzar=None,
    ) -> tuple:
        """
        Como get_metering, pero parte el rango en tramos para no saturar el
        servidor en peticiones largas (p. ej. un año cuartohorario).

        al_avanzar(n, total, ini, fin): callback opcional para mostrar progreso.

        Devuelve (DataFrame, lista de tramos fallidos). Un tramo que falla no
        aborta el resto: se devuelven los datos parciales y se informa del fallo.
        """
        dias = self.DIAS_TRAMO_METERING.get(period)
        if dias:
            tramos = partir_rango(date_from, date_to, dias)
        else:
            tramos = [(date_from, date_to)]

        dfs, fallidos = [], []
        for n, (ini, fin) in enumerate(tramos, start=1):
            if al_avanzar:
                al_avanzar(n, len(tramos), ini, fin)
            try:
                xml = self.get_metering(
                    id=id, date_from=ini, date_to=fin, data_source=data_source,
                    period=period, field=field, language=language,
                )
                dfs.append(metering_a_dataframe(xml))
            except GemwebError as e:
                fallidos.append(f"{ini} → {fin}: {e}")

        dfs = [d for d in dfs if not d.empty]
        if not dfs:
            return pd.DataFrame(), fallidos

        # Los tramos se solapan un día: quitamos los valores repetidos
        df = pd.concat(dfs, ignore_index=True)
        df = df.drop_duplicates(subset=["id_suministro", "fecha"]).reset_index(drop=True)
        return df, fallidos


def partir_rango(date_from: str, date_to: str, dias: int) -> list:
    """
    Parte [date_from, date_to] en tramos de 'dias' días como máximo.
    Cada tramo empieza el día en que acaba el anterior, así no se pierde
    ningún dato sea cual sea el criterio de la API con la fecha fin.
    """
    ini = date.fromisoformat(date_from)
    fin = date.fromisoformat(date_to)
    tramos = []
    while True:
        fin_tramo = min(ini + timedelta(days=dias), fin)
        tramos.append((ini.isoformat(), fin_tramo.isoformat()))
        if fin_tramo >= fin:
            break
        ini = fin_tramo
    return tramos


# ---------- Helpers para parsear XML a DataFrame ----------

def xml_a_dict(elem: ET.Element) -> dict:
    """Convierte un elemento XML 'plano' (hijos directos con texto) a dict."""
    return {hijo.tag: hijo.text for hijo in elem}


def inventory_a_dataframe(root: ET.Element) -> pd.DataFrame:
    """
    Parsea la respuesta de get_inventory.
    Los elementos cuelgan directamente de la raíz, uno por registro.
    """
    filas = [xml_a_dict(nodo) for nodo in root]
    return pd.DataFrame(filas)


def consumo_periodos_a_dataframe(root: ET.Element) -> pd.DataFrame:
    """
    Parsea get_consumption_by_period. Devuelve filas tipo:
    fecha | p1 | p2 | p3 | ... | total
    """
    filas = []
    for sub in root.findall(".//subministrament"):
        sub_id = sub.findtext("id")
        units = sub.findtext("units")
        for year in sub.findall(".//year"):
            y = year.get("value")
            for month in year.findall("month"):
                m = month.get("value")
                # ¿Es diario o mensual?
                dias = month.findall("day")
                if dias:
                    for day in dias:
                        fila = {
                            "id_suministro": sub_id,
                            "unidades": units,
                            "fecha": f"{y}-{int(m):02d}-{int(day.get('value')):02d}",
                        }
                        for hijo in day:
                            if hijo.tag.startswith("consumption_"):
                                fila[hijo.tag] = float(hijo.text or 0)
                        filas.append(fila)
                else:
                    fila = {
                        "id_suministro": sub_id,
                        "unidades": units,
                        "fecha": f"{y}-{int(m):02d}",
                    }
                    for hijo in month:
                        if hijo.tag.startswith("consumption_"):
                            fila[hijo.tag] = float(hijo.text or 0)
                    filas.append(fila)

    df = pd.DataFrame(filas)
    if not df.empty:
        # Columna total con la suma de los periodos
        cols_p = [c for c in df.columns if c.startswith("consumption_")]
        if cols_p:
            df["total"] = df[cols_p].sum(axis=1)
    return df


def metering_a_dataframe(root: ET.Element) -> pd.DataFrame:
    """Parsea get_metering. Cada <value date='...'>X</value> es una fila."""
    filas = []
    for sub in root.findall(".//subministrament"):
        sub_id = sub.findtext("id")
        units = sub.findtext("units")
        for val in sub.findall(".//value"):
            filas.append({
                "id_suministro": sub_id,
                "fecha": val.get("date"),
                "valor": float(val.text or 0),
                "unidades": units,
            })
    return pd.DataFrame(filas)


def cost_consumption_a_dataframe(root: ET.Element) -> pd.DataFrame:
    """Parsea get_cost_consumption. Devuelve filas año-mes-valor."""
    filas = []
    for sub in root.findall(".//subministrament"):
        sub_id = sub.findtext("id")
        units = sub.findtext("units")
        completeness = sub.findtext("completeness")
        for year in sub.findall(".//year"):
            y = year.get("value")
            for month in year.findall("month"):
                filas.append({
                    "id_suministro": sub_id,
                    "anio": int(y),
                    "mes": int(month.get("value")),
                    "fecha": f"{y}-{int(month.get('value')):02d}",
                    "valor": float(month.text or 0),
                    "unidades": units,
                    "completitud_%": completeness,
                })
    return pd.DataFrame(filas)


def power_optimization_a_dataframe(root: ET.Element) -> pd.DataFrame:
    """Parsea get_power_optimization. Una fila por suministro con potencias y ahorros."""
    filas = []
    for sub in root.findall(".//subministrament"):
        fila = {"id_suministro": sub.findtext("id"),
                "completitud_%": sub.findtext("completeness")}

        # Potencias actuales
        for p in sub.findall("./current/power/*"):
            fila[f"actual_{p.tag}_kW"] = float(p.text or 0)
        # Potencias óptimas
        for p in sub.findall("./optimum/power/*"):
            fila[f"optimo_{p.tag}_kW"] = float(p.text or 0)
        # Costes
        for c in sub.findall("./current/costs/*"):
            fila[f"coste_actual_{c.tag}_€"] = float(c.text or 0)
        for c in sub.findall("./optimum/costs/*"):
            fila[f"coste_optimo_{c.tag}_€"] = float(c.text or 0)
        # Ahorros
        for s in sub.findall("./optimum/savings/*"):
            fila[f"ahorro_{s.tag}_€"] = float(s.text or 0)
        filas.append(fila)
    return pd.DataFrame(filas)
