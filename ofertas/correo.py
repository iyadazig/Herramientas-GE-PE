# -*- coding: utf-8 -*-
"""
TEXTO DEL CORREO de peticion de ofertas a las comercializadoras
================================================================
Se genera con los mismos datos que el Excel. En el listado, junto a cada CUPS, solo la
exencion del impuesto electrico y la compensacion de excedentes (si las tiene). En el
resumen, para el resto de condiciones, lo mas habitual entre los suministros y, si alguno es
distinto, la excepcion con sus CUPS. Del precio fijo solo se dice si se pide (la binomica
segun la tarifa va en el Excel).

En gas natural, junto al CUPS: el Impuesto sobre Hidrocarburos si no es el tipo reducido, la
cogeneracion y la planta satelite. En el resumen, ademas, el Qd total, el impuesto, el termino
fijo y la telemedida.
"""

from collections import Counter

from ofertas import solicitud as so

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga(texto: str) -> str:
    """'22/10/2026' -> '22 de octubre de 2026' ('' si no hay fecha)."""
    if not texto:
        return ""
    d, m, a = (int(x) for x in texto.split("/"))
    return f"{d} de {MESES[m - 1]} de {a}"


def _kwh(valor) -> str:
    return f"{valor:,.0f}".replace(",", ".") + " kWh"


def _pct(valor) -> str:
    return f"{valor:g}".replace(".", ",") + " %"


def _lista(nombres: list) -> str:
    """'A', 'A y B', 'A, B y C'."""
    nombres = list(nombres)
    return nombres[0] if len(nombres) == 1 else ", ".join(nombres[:-1]) + " y " + nombres[-1]


def _unicos(valores: list) -> list:
    return list(dict.fromkeys(v for v in valores if v))


def _exencion(c: dict) -> str:
    datos = []
    if c.get("pct_consumo_exento") is not None:
        datos.append(f"{_pct(c['pct_consumo_exento'])} del consumo")
    if c.get("pct_tipo_exencion") is not None:
        datos.append(f"tipo de exención del {_pct(c['pct_tipo_exencion'])}")
    return "sujeto a exención del impuesto eléctrico" + (f" ({'; '.join(datos)})"
                                                         if datos else "")


def _duracion(c: dict) -> str:
    plazos = [f"hasta el {fecha_larga(c[k])}" for k in ("duracion", "duracion2", "duracion3")
              if c.get(k)]
    if len(plazos) > 1:
        return f"{['', '', 'dos', 'tres'][len(plazos)]} alternativas, {_lista(plazos)}"
    return plazos[0] if plazos else ""


def _ieh(c: dict) -> str:
    """'tipo reducido 0,15 €/GJ (usos profesionales)', 'exento (60 % del consumo)'..."""
    tipo = c.get("ieh") or so.IEH_REDUCIDO
    texto = tipo[0].lower() + tipo[1:]
    if c.get("pct_ieh") is not None:
        texto += f" ({_pct(c['pct_ieh'])} del consumo)"
    return texto


def _nota_cups_gas(s: dict) -> str:
    """Gas: junto al CUPS, el impuesto si no es el reducido, cogeneracion y planta satelite."""
    c, t = s.get("consideraciones") or {}, s.get("tecnicos") or {}
    notas = []
    if (c.get("ieh") or so.IEH_REDUCIDO) != so.IEH_REDUCIDO:
        notas.append(f"Impuesto sobre Hidrocarburos: {_ieh(c)}")
    if t.get("cogeneracion"):
        notas.append("con cogeneración")
    if t.get("planta_satelite"):
        notas.append("suministrado desde planta satélite de GNL")
    return "; ".join(notas)


def _nota_cups(c: dict) -> str:
    """Lo que se indica junto al CUPS en el listado: exencion y excedentes."""
    notas = []
    if c.get("exencion"):
        notas.append(_exencion(c))
    if c.get("excedentes"):
        notas.append("con compensación de excedentes")
    return "; ".join(notas)


def _mas_habitual(suministros: list, valor) -> tuple:
    """(valor mas repetido, [(otro valor, [CUPS])]) — a igualdad, el del primer suministro."""
    valores = [valor(s) for s in suministros]
    cuenta = Counter(valores)
    principal = max(dict.fromkeys(valores), key=lambda v: cuenta[v])
    otros = {}
    for s, v in zip(suministros, valores):
        if v != principal:
            otros.setdefault(v, []).append(s.get("cups", ""))
    return principal, list(otros.items())


def _con_excepciones(texto: str, otros: list, describir) -> str:
    """'texto (ES1 y ES2: otro valor; ES3: ...)'."""
    if not otros:
        return texto
    return texto + " (" + "; ".join(f"{_lista(cups)}: {describir(v)}" for v, cups in otros) + ")"


def _resumen(suministros: list, gas: bool = False) -> list:
    cons = [s.get("consideraciones") or {} for s in suministros]
    datos = [dict(s, _c=c, _t=s.get("tecnicos") or {}) for s, c in zip(suministros, cons)]
    lineas = [f"Número de suministros: {len(suministros)}, con un consumo anual previsto total "
              f"de {_kwh(sum(so.consumo_anual(s) for s in suministros))}."]
    if gas:
        qd = sum(s["_t"].get("qd") or 0 for s in datos)
        if qd:
            lineas[0] = lineas[0][:-1] + (f" y un caudal diario (Qd) total de "
                                          f"{_kwh(qd)}/día.")

    def tipologia(v):
        return _lista(v) if v else "ninguna"
    principal, otros = _mas_habitual(datos, lambda s: tuple(
        ([so.PRECIO_FIJO] if s["_c"].get("fijo", True) else [])
        + list(s["_c"].get("indexado") or ())))
    if principal:
        lineas.append(_con_excepciones(f"Tipología de precios: {tipologia(principal)}", otros,
                                       tipologia) + ".")
    elif otros:
        lineas.append("Tipología de precios, solo en algunos suministros: " + "; ".join(
            f"{_lista(cups)}: {tipologia(v)}" for v, cups in otros) + ".")

    principal, otros = _mas_habitual(
        datos, lambda s: "de temporada" if s["_c"].get("temporada") else "anual")
    lineas.append(_con_excepciones(f"Tipo de contrato: {principal}", otros, str) + ".")
    if gas:
        terminos = {so.CUOTA_FIJA: "cuota fija (por cliente)",
                    so.VARIABLE: "variable (por caudal)"}
        for nombre, valor in (
                ("Impuesto sobre Hidrocarburos", lambda s: _ieh(s["_c"])),
                ("Término fijo", lambda s: terminos.get(s["_t"].get("termino_fijo"),
                                                        terminos[so.CUOTA_FIJA])),
                ("Telemedida", lambda s: "sí" if s["_t"].get("telemedida") else "no")):
            principal, otros = _mas_habitual(datos, valor)
            lineas.append(_con_excepciones(f"{nombre}: {principal}", otros, str) + ".")

    def inicio(v):
        return fecha_larga(v) or "por determinar"
    principal, otros = _mas_habitual(datos, lambda s: s["_c"].get("inicio") or "")
    if principal or otros:
        lineas.append(_con_excepciones(f"Fecha prevista de inicio del contrato: "
                                       f"{inicio(principal)}", otros, inicio) + ".")

    def duracion(v):
        return v or "por determinar"
    principal, otros = _mas_habitual(datos, lambda s: _duracion(s["_c"]))
    if principal or otros:
        lineas.append(_con_excepciones(f"Duración del contrato: {duracion(principal)}", otros,
                                       duracion) + ".")

    gdo = "garantías de origen de gas renovable" if gas else "garantías de origen (GdO)"
    con_gdo = [s.get("cups", "") for s, c in zip(suministros, cons) if c.get("gdo")]
    if len(con_gdo) == len(suministros):
        lineas.append(f"Las ofertas deben incluir {gdo}.")
    elif con_gdo and len(con_gdo) * 2 >= len(suministros):
        sin = [s.get("cups", "") for s, c in zip(suministros, cons) if not c.get("gdo")]
        lineas.append(f"Las ofertas deben incluir {gdo}, salvo en {_lista(sin)}.")
    elif con_gdo:
        lineas.append(f"Las ofertas de {_lista(con_gdo)} deben incluir {gdo}.")
    return lineas


def texto_correo(cabecera: dict, suministros: list) -> str:
    """Texto para pegar en el correo de peticion de ofertas a los gestores de las
    comercializadoras."""
    gas = cabecera.get("producto_tipo") == so.GAS
    energia = "gas natural" if gas else "electricidad"
    cliente = _lista(_unicos([s.get("grupo") for s in suministros])
                     or _unicos([s.get("razon") for s in suministros]) or ["(cliente)"])
    listado = []
    for s in suministros:
        linea = f"    • CUPS: {s.get('cups', '')}"
        if s.get("razon"):
            linea += f" ({s['razon']})"
        nota = _nota_cups_gas(s) if gas else _nota_cups(s.get("consideraciones") or {})
        listado.append(linea + (f" — {nota}" if nota else ""))

    desde, hasta = (fecha_larga(cabecera.get("recepcion_desde")),
                    fecha_larga(cabecera.get("recepcion_hasta")))
    hora = cabecera.get("recepcion_hora")
    plazo = []
    if desde:
        plazo.append(f"    • Inicio: {desde}.")
    if hasta:
        plazo.append(f"    • Fecha máxima: {hasta}" + (f", antes de las {hora} horas." if hora
                                                        else "."))
    partes = [
        "Estimado comercializador:",
        f"GE&PE va a ser el encargado de realizar el proceso de licitación de los suministros de "
        f"{energia} del cliente {cliente} en España.",
        f"Nos ponemos en contacto con usted para trasladarle el interés de {cliente} en recibir "
        f"ofertas de {energia} para los siguientes suministros ubicados en España:",
        "\n".join(listado),
        "Resumen de la solicitud:\n" + "\n".join("    • " + l
                                                 for l in _resumen(suministros, gas)),
        "Adjuntamos la documentación necesaria para la preparación de las ofertas. En ella se "
        "detallan las tipologías de las ofertas solicitadas, las condiciones requeridas y los "
        "datos de consumo esperados para el próximo año de contrato, entre otros datos. Rogamos "
        "la lean atentamente para cumplir los requerimientos particulares de cada suministro.",
    ]
    if plazo:
        partes.append("Plazo de recepción de las ofertas:\n" + "\n".join(plazo))
    partes += [
        "Las propuestas deberán remitirse a la siguiente dirección de correo electrónico:\n"
        + (cabecera.get("email") or "(email del gestor)"),
        "Para cualquier duda o cuestión sobre esta solicitud, no dude en contactar con "
        "nosotros.",
        "Sin otro particular, a la espera de sus noticias, reciba un cordial saludo.",
        "Atentamente,",
    ]
    return "\n\n".join(partes) + "\n"
