# -*- coding: utf-8 -*-
"""
BORRADOR DEL CORREO AL CLIENTE CON LOS RESULTADOS
=================================================
Plantillas de texto y reglas (sin servicios de IA), con la estructura de los correos de los
estudios hechos con la plantilla Excel:
    1. Contrato actual (coste estimado)
    2. Resultados (tabla por modalidad y duración: 12 meses / más de 12 meses)
    3. Conclusiones (reglas: qué ofertas mejoran el actual, la más barata de cada tipo, las de
       12 meses frente a las largas, el indexado frente al actual)
    4. Recomendación (reglas) y notas (apuntamiento, mercado)
borrador() da el asunto, el texto plano (para editarlo en la app) y el HTML. Sin Streamlit.
"""

from xml.sax.saxutils import escape

from analisis_ofertas import informes, resultados
from analisis_ofertas import modelo as mo

AZUL = "#1F4E79"


def _categoria(f: dict) -> str:
    plazo = "12 meses" if f["duracion"] <= 12 else "más de 12 meses"
    return f"{'Precio fijo' if f['tipo'] == 'Fijo' else 'Indexado'} – {plazo}"


def _modalidad(f: dict) -> str:
    if f["tipo"] == "Fijo":
        return "Fijo · " + f["ssaa"]
    return f["modalidad"] + (f" · {f['cierres']}" if f["cierres"] else "")


def conclusiones(filas: list) -> list:
    """Frases con las reglas sobre los resultados."""
    salida = []
    con_dato = [f for f in filas if f["variacion"] is not None]
    mejoran = sorted((f for f in con_dato if f["variacion"] < 0), key=lambda f: f["variacion"])
    if mejoran:
        nombres = [f"{f['comercializadora']} a {f['duracion']} meses ({informes.pct(f['variacion'])})"
                   for f in mejoran]
        salida.append(("Mejoran el contrato actual: " if len(mejoran) > 1 else
                       "La única oferta que mejora el contrato actual es ") + ", ".join(nombres)
                      + ".")
    elif con_dato:
        salida.append("Ninguna oferta mejora el coste estimado del contrato actual en su mismo "
                      "periodo.")
    mejores = resultados.mejores(filas)
    for t, f in mejores.items():
        salida.append(f"La mejor oferta {('a precio fijo' if t == 'Fijo' else 'indexada')} frente "
                      f"al contrato actual es {f['comercializadora']} a {f['duracion']} meses: "
                      f"{informes.eur(f['media_anual'])} €/año ({informes.pct(f['variacion'])}).")
    cortas = [f for f in con_dato if f["tipo"] == "Fijo" and f["duracion"] <= 12]
    if cortas and all(f["variacion"] > 0 for f in cortas):
        lo, hi = min(f["variacion"] for f in cortas), max(f["variacion"] for f in cortas)
        salida.append(f"Las ofertas a precio fijo de 12 meses encarecen el suministro entre un "
                      f"{lo * 100:.0f} % y un {hi * 100:.0f} %.")
    return salida


def _mejor_variacion(filas: list, tipo: str):
    """La oferta del tipo con mejor variación frente al actual (la que más ahorra en %)."""
    de_tipo = [f for f in filas if f["tipo"] == tipo and f["variacion"] is not None]
    return min(de_tipo, key=lambda f: f["variacion"]) if de_tipo else None


def recomendacion(filas: list) -> list:
    salida = []
    fijo, idx = _mejor_variacion(filas, "Fijo"), _mejor_variacion(filas, "Indexado")
    if fijo and fijo["variacion"] is not None and fijo["variacion"] < 0:
        salida.append(f"Si priorizáis la estabilidad presupuestaria, la alternativa más interesante "
                      f"es el precio fijo a {fijo['duracion']} meses con {fijo['comercializadora']}.")
    elif fijo:
        salida.append("No cerrar un precio fijo mientras el mercado cotice en los niveles "
                      "actuales: ninguna oferta a precio fijo mejora el contrato actual.")
    if idx and idx["variacion"] is not None:
        if idx["variacion"] < 0:
            salida.append(f"En el indexado, la oferta de {idx['comercializadora']} mejora el "
                          f"contrato actual ({informes.pct(idx['variacion'])}).")
        else:
            salida.append("En el indexado, el contrato actual sigue siendo hoy la mejor opción.")
    candidatas = sorted((f for f in filas if f["variacion"] is not None),
                        key=lambda f: f["variacion"])[:3]
    if candidatas:
        salida.append("Solicitar una nueva ronda a " + ", ".join(
            f"{f['comercializadora']} ({f['duracion']} meses)" for f in candidatas)
            + " para que mejoren precios y detallen sus condiciones contractuales.")
    return salida


def borrador(e: mo.Estudio, res: dict, ref: str = "", adjuntos=("resumen de ofertas",
                                                                  "relación de suministros")):
    """{"asunto", "texto", "html"}."""
    filas = resultados.comparativa(res)
    actual = res.get("actual")
    asunto = f"{e.cliente} – Resultados del estudio de ofertas eléctricas" + (
        f" (Ref. {ref})" if ref else "")
    t = ["Buenos días:", "",
         f"Os enviamos los resultados del estudio de las ofertas de suministro eléctrico "
         f"recibidas para {e.cliente}" + (f", con inicio de suministro el "
                                         f"{e.inicio:%d/%m/%Y}" if e.inicio else "") + ".", ""]
    h = [f"<p>Buenos días:</p><p>Os enviamos los resultados del estudio de las ofertas de "
         f"suministro eléctrico recibidas para <b>{escape(e.cliente)}</b>"
         + (f", con inicio de suministro el {e.inicio:%d/%m/%Y}" if e.inicio else "") + ".</p>"]

    def seccion(n, titulo):
        t.extend([f"{n}. {titulo.upper()}", ""])
        h.append(f'<h3 style="font-size:12pt;color:{AZUL};margin:22px 0 6px 0;border-bottom:2px '
                 f'solid {AZUL};padding-bottom:3px;">{n}. {escape(titulo)}</h3>')

    n = 1
    if actual:
        p = actual.por_anio()[0]
        seccion(n, f"Contrato actual ({actual.oferta.comercializadora}, "
                   f"{actual.oferta.modalidad.lower()})")
        frase = (f"El coste estimado del contrato actual para el primer año es de "
                 f"{informes.eur(p['total'])} €/año"
                 + (f" ({p['total'] / p['consumo'] * 1000:.2f} €/MWh)".replace(".", ",")
                    if p["consumo"] else "")
                 + ". Es el valor de referencia frente al que se comparan las ofertas.")
        t.extend([frase, ""])
        h.append(f"<p>{escape(frase)}</p>")
        n += 1
    seccion(n, "Resultados")
    n += 1
    cab = ["Oferta", "Modalidad", "Coste medio anual (€/año)",
           "Contrato actual en el mismo periodo (€/año)", "Diferencia vs contrato actual"]
    celda = "border:1px solid #9bb7d4;padding:5px 10px;"
    h.append('<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;'
             'font-size:10pt;">')
    h.append("<tr>" + "".join(f'<th style="{celda}background:{AZUL};color:#fff;">{c}</th>'
                              for c in cab) + "</tr>")
    mejores = resultados.mejores(filas)
    for cat in sorted({_categoria(f) for f in filas}, key=lambda c: (not c.startswith("Precio"),
                                                                   "más" in c)):
        t.append(cat + ":")
        h.append(f'<tr><td colspan="5" style="{celda}background:#DDEBF7;font-weight:bold;">'
                 f"{escape(cat)}</td></tr>")
        for f in [f for f in filas if _categoria(f) == cat]:
            dif = (f"{'+' if f['ahorro'] < 0 else '−'}{informes.eur(abs(f['ahorro']))} € "
                   f"({informes.pct(f['variacion'])})") if f["ahorro"] is not None else "—"
            nombre = f"{f['comercializadora']} {f['duracion']} meses" + (
                f" ({f['ronda']}.ª ronda)" if f["ronda"] > 1 else "")
            t.append(f"  • {nombre} – {_modalidad(f)}: {informes.eur(f['media_anual'])} €/año "
                     f"frente a {informes.eur(f['actual_mismo_periodo'])} €/año del contrato "
                     f"actual ({dif}).")
            destacado = any(m["nombre"] == f["nombre"] for m in mejores.values())
            color = "#00703C" if (f["variacion"] or 0) < 0 else "#C00000"
            h.append("<tr>" + f'<td style="{celda}font-weight:bold;">{escape(nombre)}</td>'
                     f'<td style="{celda}text-align:center;">{escape(_modalidad(f))}</td>'
                     f'<td style="{celda}text-align:right;{"background:#E2EFDA;font-weight:bold;" if destacado else ""}">'
                     f"{informes.eur(f['media_anual'])}</td>"
                     f'<td style="{celda}text-align:right;background:#FFFBEA;">'
                     f"{informes.eur(f['actual_mismo_periodo'])}</td>"
                     f'<td style="{celda}text-align:right;white-space:nowrap;color:{color};">'
                     f"{escape(dif)}</td></tr>")
        t.append("")
    h.append("</table>")
    nota = ("Costes sin IVA. En las ofertas de más de 12 meses se muestra el coste medio anual "
            "durante su duración, comparado con el coste estimado del contrato actual en esos "
            "mismos meses.")
    t.extend([nota, ""])
    h.append(f'<p style="font-size:9pt;color:#7f7f7f;">{escape(nota)}</p>')
    for titulo, frases in (("Conclusiones", conclusiones(filas)),
                           ("Recomendación", recomendacion(filas))):
        if not frases:
            continue
        seccion(n, titulo)
        n += 1
        t.extend([f"  • {x}" for x in frases] + [""])
        h.append("<ul>" + "".join(f"<li>{escape(x)}</li>" for x in frases) + "</ul>")
    otras = [x for x in informes.notas(e, res)[2:]]
    if otras:
        t.extend(otras + [""])
        h.extend(f'<p style="font-size:9pt;color:#7f7f7f;">{escape(x)}</p>' for x in otras)
    if adjuntos:
        frase = "Adjuntamos el " + " y la ".join(adjuntos) + "."
        t.extend([frase, ""])
        h.append(f"<p>{escape(frase)}</p>")
    cierre = ["Quedamos a vuestra disposición para cualquier aclaración.", "", "Un saludo,"]
    t.extend(cierre)
    h.append("<p>Quedamos a vuestra disposición para cualquier aclaración.</p><p>Un saludo,</p>")
    html = ('<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>'
            + escape(asunto) + '</title></head><body style="font-family:Calibri,Arial,sans-serif;'
            'font-size:11pt;color:#1f1f1f;line-height:1.45;"><div style="max-width:860px;">'
            + f'<p style="color:#7f7f7f;font-size:9.5pt;"><b>Asunto:</b> {escape(asunto)}</p>'
            + "".join(h) + "</div></body></html>")
    return {"asunto": asunto, "texto": "\n".join(t), "html": html}
