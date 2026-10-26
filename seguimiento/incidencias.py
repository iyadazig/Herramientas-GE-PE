# -*- coding: utf-8 -*-
"""
Avisos de incidencias en facturas (logica, sin Streamlit)
=========================================================
Del programa «Version_avisos_discrepancias_260923_app_envio_emails_incidencias.py», con las
mismas reglas: se cruza el Excel de facturas con los contactos de envio por tecnico,
auxiliar, cliente/grupo, denominacion y CUPS; una factura es incidencia si tiene exceso de
potencia, exceso de reactiva o lectura estimada; un correo por cliente/grupo con una tabla
por suministro. En el segundo envio del mes solo cuentan las facturas que no estaban en el
Excel anterior. Los contactos salen de la tabla «Contactos de envio».
"""

import pandas as pd

from seguimiento.informes import FIRMA_HTML, extraer_emails

COLUMNAS_FACTURAS = {
    "nom_contacte_1_1": "tecnico", "nom_contacte_1_2": "auxiliar", "Fecha de baja": "fecha_baja",
    "nombre": "cliente_grupo", "nombre.1": "denominacion_suministro", "CUPS": "cups",
    "dirección": "direccion_suministro", "Número de factura": "numero_factura",
    "Período inicial": "periodo_inicial", "Período final": "periodo_final",
    "Fecha de emisión": "fecha_emision", "Importe total exceso potencia": "importe_potencia",
    "Importe reactiva": "importe_reactiva", "Lectura": "lectura"}
COLUMNAS_CONTACTOS = {
    "Técnico": "tecnico", "Auxiliar": "auxiliar", "Cliente/Grupo": "cliente_grupo",
    "Nombre cliente": "nombre_cliente", "Para": "email_to", "CC": "email_cc", "CUPS": "cups",
    "Denominación suministro": "denominacion_suministro"}
CLAVES_CRUCE = ["tecnico", "auxiliar", "cliente_grupo", "denominacion_suministro", "cups"]
NUMERO_FACTURA = "Número de factura"


def leer_excel(contenido):
    return pd.read_excel(contenido)


def incidencias_nuevas(df_anterior, df_actual):
    """Filas del Excel actual cuya factura no estaba en el anterior (None si no hay)."""
    if NUMERO_FACTURA not in df_anterior.columns or NUMERO_FACTURA not in df_actual.columns:
        raise ValueError("La columna «Número de factura» no existe en uno de los Excel.")
    anteriores = set(df_anterior[NUMERO_FACTURA].astype(str))
    df_actual = df_actual.copy()
    df_actual[NUMERO_FACTURA] = df_actual[NUMERO_FACTURA].astype(str)
    nuevas = df_actual[~df_actual[NUMERO_FACTURA].isin(anteriores)]
    return None if nuevas.empty else nuevas


def _fecha(v):
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%y")
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d/%m/%y")
    except (TypeError, ValueError):
        return str(v)


def preparar(facturas, contactos, auxiliar, emails_tecnicos):
    """facturas: DataFrame del Excel; contactos: datos de «Contactos de envio».
    Devuelve (envios, sin_contacto (DataFrame), aviso o None). Cada envio: grupo,
    nombre_cliente, tecnico, para, cc, asunto, html, filas (para el modo prueba)."""
    facturas = facturas.rename(columns=COLUMNAS_FACTURAS)
    faltan = [c for c in CLAVES_CRUCE + ["numero_factura", "importe_potencia",
                                         "importe_reactiva", "lectura"] if c not in facturas]
    if faltan:
        raise ValueError("Al Excel de facturas le faltan columnas: %s." % ", ".join(faltan))
    contactos = pd.DataFrame(contactos, columns=list(COLUMNAS_CONTACTOS)).rename(
        columns=COLUMNAS_CONTACTOS)
    # quita espacios al principio y al final (nunca los interiores) para cruzar
    for col in CLAVES_CRUCE:
        facturas[col] = facturas[col].fillna("").astype(str).str.strip()
        contactos[col] = contactos[col].fillna("").astype(str).str.strip()

    chequeo = pd.merge(facturas, contactos[CLAVES_CRUCE].drop_duplicates(), on=CLAVES_CRUCE,
                       how="left", indicator=True)
    sin_contacto = chequeo[chequeo["_merge"] == "left_only"][
        CLAVES_CRUCE + ["numero_factura"]].drop_duplicates()

    df = pd.merge(facturas, contactos, on=CLAVES_CRUCE, how="inner").fillna(0)
    for col in ("importe_potencia", "importe_reactiva"):
        df[col] = (df[col].astype(str).str.replace(",", ".", regex=False)
                   .replace(["", "nan", "None"], "0"))
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["lectura"] = (df["lectura"].astype(str).str.upper()
                     .replace(["", "NAN", "NONE", "0"], "R"))

    def factura_valida(f):
        return not (f["importe_potencia"] <= 0 and f["importe_reactiva"] <= 0
                    and f["lectura"].startswith("R"))

    if not df.empty:
        df = df[df.apply(factura_valida, axis=1)]
    df = df[df["auxiliar"] == auxiliar]
    if df.empty:
        return [], sin_contacto, "No hay incidencias para este auxiliar."
    df = df.drop_duplicates()

    envios = []
    for grupo, df_grupo in df.groupby("cliente_grupo"):
        nombre_cliente = df_grupo["nombre_cliente"].iloc[0]
        tecnico = df_grupo["tecnico"].iloc[0]
        filtro = (contactos["cliente_grupo"] == grupo) & (contactos["auxiliar"] == auxiliar)
        para = extraer_emails(contactos.loc[filtro, "email_to"])
        cc = extraer_emails(contactos.loc[filtro, "email_cc"])
        if emails_tecnicos.get(tecnico):
            cc.append(emails_tecnicos[tecnico])
        para, cc = "; ".join(sorted(set(para))), "; ".join(sorted(set(cc)))
        asunto = f"Aviso incidencia en factura_{grupo}"
        html = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size: 13px;">
            <p>Estimado/a <b>{nombre_cliente}</b>,</p>
            <p>En relación con las facturas emitidas, se han detectado las siguientes incidencias:</p>
            """
        filas = []
        for _, s in df_grupo[["cups", "denominacion_suministro"]].drop_duplicates().iterrows():
            df_s = df_grupo[df_grupo["cups"] == s["cups"]]
            direccion = df_s["direccion_suministro"].iloc[0] if "direccion_suministro" in df_s \
                else ""
            html += f"""
                <p><b>Suministro:</b> {s["denominacion_suministro"]}<br>
                <b>CUPS:</b> {s["cups"]}<br>
                <b>Dirección de suministro:</b> {direccion}</p>

                <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;">
                <tr><th>Nº Factura</th><th>Período</th><th>Incidencias</th></tr>
                """
            for _, f in df_s.iterrows():
                incidencias = []
                if f["importe_potencia"] > 0:
                    incidencias.append(f"Exceso de Potencia ({f['importe_potencia']:.2f} €)")
                if f["importe_reactiva"] > 0:
                    incidencias.append(f"Exceso de Reactiva ({f['importe_reactiva']:.2f} €)")
                if f["lectura"] == "E":
                    incidencias.append("Lectura estimada")
                if not incidencias:
                    continue
                inicio, fin = _fecha(f.get("periodo_inicial")), _fecha(f.get("periodo_final"))
                html += f"""
                    <tr>
                        <td>{f['numero_factura']}</td>
                        <td>{inicio} – {fin}</td>
                        <td>{"<br>".join(incidencias)}</td>
                    </tr>
                    """
                filas.append({"Técnico": tecnico, "Auxiliar": auxiliar, "Cliente": grupo,
                              "Suministro": s["denominacion_suministro"],
                              "Inicio factura": inicio, "Fin factura": fin, "Para": para,
                              "CC": cc, "Asunto": asunto, "Número factura": f["numero_factura"],
                              "Incidencias": " | ".join(incidencias)})
            html += "</table><br>"
            if (df_s["importe_reactiva"] > 0).any():
                html += """
                    <p><b>Se recomienda revisar los equipos de compensación de energía reactiva, con el fin de evitar mayores recargos en el futuro.</b></p><br>
                    """
        html += f"""
            <p>Para cualquier duda, por favor contacte con su técnico asignado.</p>
            <p>Saludos.</p>
            {FIRMA_HTML}
            </body></html>
            """
        envios.append({"grupo": grupo, "nombre_cliente": nombre_cliente, "tecnico": tecnico,
                       "para": para, "cc": cc, "asunto": asunto, "html": html, "filas": filas,
                       "facturas": len(filas)})
    return envios, sin_contacto, None
