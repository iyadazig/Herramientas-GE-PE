# -*- coding: utf-8 -*-
"""
Envio de informes de seguimiento (logica, sin Streamlit)
========================================================
Del programa «260715_app_envio_emails_informes.py»: los PDF se llaman
AAAAMMDD_AAAAMMDD_Inf_SyC_Denominacion[_Tipo].pdf, se cruzan con los contactos de envio del
auxiliar por la denominacion del suministro y se envia un correo por cliente/grupo con todos
sus informes, en copia el tecnico. La diferencia: los contactos salen de la tabla «Contactos
de envio» y los PDF que no casan con nadie se avisan en vez de ignorarse.
"""

import datetime as dt
import os

FIRMA_HTML = """
<hr style="margin-top:20px; margin-bottom:15px;">

<table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 12px;">
<tr>
    <td style="padding-right:15px; vertical-align:top;">
        <img src="https://geype.es/wp-content/uploads/2017/02/logo.png"
             alt="GEYPE Ingeniería"
             style="width:120px;">
    </td>

    <td style="vertical-align:top;">
        <p style="margin:0; font-weight:bold; color:#9b1c1f;">
            GEYPE INGENIERÍA
        </p>
        <p style="margin:2px 0;">
            Gestión Energética
        </p>
        <p style="margin:6px 0;">
            📞 +34 954 636 737
        </p>
        <p style="margin:6px 0;">
            🌐 <a href="https://www.geype.com" target="_blank">www.geype.com</a>
        </p>
        <p style="margin:6px 0; font-style:italic;">
            "Más de 40 años de experiencia nos avalan"
        </p>

        <div style="margin-top:10px; background-color:#7ec0f4; padding:8px;">
            <b>¿Está satisfecho/a con el servicio que le estamos prestando?</b><br>
            <a href="https://docs.google.com/forms/d/e/1FAIpQLScfOw2uyEdOSb0_JOwlIsLNcOXLXRVxdd2KnV6XPS686eRV3g/viewform?vc=0&c=0&w=1&flr=0"
               target="_blank"
               style="color:#ffffff; font-weight:bold; text-decoration:none;">
               👉 Haga click aquí para dejarnos una reseña
            </a>
        </div>
    </td>
</tr>
</table>
"""

TIPOS_FACTURA = ["refacturacion", "rectificativa", "abono", "regularizacion", "ajuste", "nota",
                 "complementaria", "anulada", "extraordinaria"]


def extraer_emails(valores):
    """Emails de una lista de celdas separados por «;» o «,», sin repetir."""
    emails = []
    for valor in valores:
        if valor is None or str(valor).strip() in ("", "nan", "None"):
            continue
        for e in str(valor).replace(",", ";").split(";"):
            e = e.strip()
            if e and e not in emails:
                emails.append(e)
    return emails


def formatear_fecha(fecha):
    if len(fecha) == 8 and fecha.isdigit():
        return f"{fecha[6:8]}/{fecha[4:6]}/{fecha[2:4]}"
    return fecha


def fecha_de(texto):
    try:
        return dt.datetime.strptime(texto, "%Y%m%d").date()
    except (TypeError, ValueError):
        return None


def parsear_nombre_pdf(nombre_pdf):
    resultado = {"archivo": nombre_pdf, "periodo_inicial": None, "periodo_final": None,
                 "denominacion_suministro": None, "tipo_factura": "",
                 "suministro_completo": None}
    partes = os.path.splitext(nombre_pdf)[0].split("_")
    if len(partes) >= 4 and partes[2] == "Inf" and partes[3] == "SyC":
        resultado["periodo_inicial"] = partes[0]
        resultado["periodo_final"] = partes[1]
        tipo_factura, indice_tipo = "", -1
        for i, p in enumerate(partes[4:], start=4):
            if p.lower() in TIPOS_FACTURA:
                tipo_factura, indice_tipo = p.capitalize(), i
                break
        denominacion = "_".join(partes[4:] if indice_tipo == -1 else partes[4:indice_tipo])
        resultado["tipo_factura"] = tipo_factura
        resultado["denominacion_suministro"] = denominacion
        resultado["suministro_completo"] = partes[4] if len(partes) > 4 else denominacion
    return resultado


def _misma_denominacion(pdf, excel):
    return pdf == excel or pdf.replace("_", " ") == excel.replace("_", " ")


def preparar(pdfs, contactos, auxiliar, emails_tecnicos):
    """pdfs: [(nombre, bytes)]; contactos: datos de la tabla «Contactos de envio»;
    emails_tecnicos: {nombre: email}. Devuelve (envios, pdfs sin casar, aviso o None).
    Cada envio: grupo, nombre_cliente, tecnico, para, cc, asunto, html, adjuntos, informes
    (dicts del nombre del PDF) y cups (los de los contactos que casan)."""
    contactos_aux = [c for c in contactos if (c.get("Auxiliar") or "") == auxiliar]
    for c in contactos_aux:
        c["Cliente/Grupo"] = str(c.get("Cliente/Grupo") or "").strip()
        c["Denominación suministro"] = str(c.get("Denominación suministro") or "").strip()
    if not contactos_aux:
        return [], [n for n, _ in pdfs], "No hay contactos de envío para este auxiliar."
    lista_pdfs = []
    for nombre, datos in sorted(pdfs):
        info = parsear_nombre_pdf(nombre)
        info["datos"] = datos
        lista_pdfs.append(info)
    por_denominacion = {}
    for info in lista_pdfs:
        if info.get("denominacion_suministro"):
            por_denominacion.setdefault(info["denominacion_suministro"], []).append(info)

    envios, casados, vistos = {}, set(), set()
    for c in contactos_aux:
        clave = (c["Cliente/Grupo"], c.get("Nombre cliente"), c["Denominación suministro"])
        if clave in vistos:
            continue
        vistos.add(clave)
        grupo = c["Cliente/Grupo"]
        env = envios.setdefault(grupo, {"grupo": grupo, "nombre_cliente": c.get("Nombre cliente"),
                                        "informes": [], "cups": set()})
        for denom, infos in por_denominacion.items():
            if _misma_denominacion(denom, c["Denominación suministro"]):
                for info in infos:
                    if not any(x is info for x in env["informes"]):
                        env["informes"].append(info)
                    casados.add(info["archivo"])
                    if c.get("CUPS"):
                        info.setdefault("cups", set()).add(str(c["CUPS"]).strip().upper()[:20])
    sin_casar = [i["archivo"] for i in lista_pdfs if i["archivo"] not in casados]

    salida = []
    for grupo, env in envios.items():
        if not env["informes"]:
            continue
        filas = [c for c in contactos_aux if c["Cliente/Grupo"] == grupo]
        tecnico = filas[0].get("Técnico")
        para = extraer_emails(c.get("Para") for c in filas)
        cc = extraer_emails(c.get("CC") for c in filas)
        if emails_tecnicos.get(tecnico):
            cc.append(emails_tecnicos[tecnico])
        suministros = sorted({i["denominacion_suministro"] for i in env["informes"]
                              if i.get("denominacion_suministro")})
        asunto = ("Informe Seguimiento: %s" % suministros[0] if len(suministros) == 1
                  else "Informe Seguimiento: %s" % grupo)
        salida.append({**env, "tecnico": tecnico, "asunto": asunto,
                       "para": "; ".join(sorted(set(para))), "cc": "; ".join(sorted(set(cc))),
                       "suministros": suministros, "html": cuerpo(env),
                       "adjuntos": [(i["archivo"], i["datos"]) for i in env["informes"]]})
    return salida, sin_casar, None


def cuerpo(env):
    informes = env["informes"]
    texto_adjunto = ("el siguiente informe de seguimiento:" if len(informes) == 1
                     else "los siguientes informes de seguimiento:")
    tabla = """
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
            <tr style="background-color:#1976D2; color:white;">
                <th style="white-space:nowrap;">Denominación Suministro</th>
                <th style="white-space:nowrap;">Periodo informe</th>
                <th style="white-space:nowrap;">Tipo Factura</th>
            </tr>
            """
    for pdf in informes:
        peri = ""
        if pdf.get("periodo_inicial") and pdf.get("periodo_final"):
            peri = "%s - %s" % (formatear_fecha(pdf["periodo_inicial"]),
                                formatear_fecha(pdf["periodo_final"]))
        tabla += f"""
                <tr>
                    <td style="white-space:nowrap;">{pdf.get("denominacion_suministro", "")}</td>
                    <td style="white-space:nowrap;">{peri}</td>
                    <td style="white-space:nowrap;">{pdf.get("tipo_factura", "")}</td>
                </tr>
                """
    tabla += "</table>"
    return ("""
            <html>
            <body style="font-family: Arial, sans-serif; font-size: 13px;">
            <p>Estimado/a """ + str(env["nombre_cliente"] or "") + """,</p>
            <p>Mediante el presente, adjunto """ + texto_adjunto + """</p>
            <br/>
            """ + tabla + """
            <br/>
            <p>Reciba un cordial saludo.</p>
            """ + FIRMA_HTML + """
            </body></html>
            """)


def filas_prueba(envios, auxiliar):
    """Lo que el programa guardaba en modo_prueba_informes_*.xlsx."""
    filas = []
    for e in envios:
        periodos = sorted({(i["periodo_inicial"], i["periodo_final"]) for i in e["informes"]
                           if i.get("periodo_inicial") and i.get("periodo_final")})
        tabla = ""
        for i in e["informes"]:
            peri = ""
            if i.get("periodo_inicial") and i.get("periodo_final"):
                peri = "%s - %s" % (formatear_fecha(i["periodo_inicial"]),
                                    formatear_fecha(i["periodo_final"]))
            tabla += "%s | %s | %s\n" % (i.get("denominacion_suministro", ""), peri,
                                         i.get("tipo_factura", ""))
        filas.append({"Técnico": e["tecnico"], "Auxiliar": auxiliar,
                      "Suministro(s)": ", ".join(e["suministros"]), "Para": e["para"],
                      "CC": e["cc"], "Asunto": e["asunto"],
                      "Periodo": "%s - %s" % (periodos[0][0], periodos[-1][1]) if periodos else "",
                      "Informes": tabla, "Total informes": len(e["informes"])})
    return filas


def meses_a_marcar(envio):
    """[(cups, anio, mes)] que pasan a «Inf enviado» tras enviar el correo."""
    from seguimiento.datos import meses_cubiertos
    salida = []
    for i in envio["informes"]:
        ini, fin = fecha_de(i.get("periodo_inicial")), fecha_de(i.get("periodo_final"))
        for cups in sorted(i.get("cups") or ()):
            salida += [(cups, a, m) for a, m in meses_cubiertos(ini, fin)]
    return salida
