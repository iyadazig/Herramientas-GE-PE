# -*- coding: utf-8 -*-
"""
TEXTO DEL CORREO de peticion de ofertas a las comercializadoras
================================================================
Se genera con los mismos datos que el Excel. Lo que es igual en todos los suministros va en
el resumen general; lo que solo se aplica a algunos, al lado de su CUPS en el listado.
"""

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
    plazos = [f"hasta el {fecha_larga(d)}" for d in (c.get("duracion"), c.get("duracion2")) if d]
    if len(plazos) == 2:
        return f"dos alternativas, {plazos[0]} y {plazos[1]}"
    return plazos[0] if plazos else ""


# caracteristicas que pueden cambiar de un suministro a otro: (clave, valor del suministro)
def _caracteristicas(s: dict) -> dict:
    c = s.get("consideraciones") or {}
    return {"fijo": so.precio_fijo(s.get("tarifa")),
            "indexado": tuple(c.get("indexado") or ()),
            "tipo": "de temporada" if c.get("temporada") else "anual",
            "inicio": c.get("inicio") or "",
            "duracion": _duracion(c),
            "exencion": _exencion(c) if c.get("exencion") else "",
            "gdo": bool(c.get("gdo")), "excedentes": bool(c.get("excedentes"))}


def _nota_cups(car: dict, distintas: set) -> str:
    """Lo que solo se aplica a algunos suministros, para ponerlo junto al CUPS."""
    notas = []
    if "fijo" in distintas:
        notas.append(f"precio fijo {car['fijo'][0].lower() + car['fijo'][1:]}")
    if "indexado" in distintas:
        notas.append("indexado: " + (", ".join(car["indexado"]) if car["indexado"]
                                     else "no se solicita"))
    if "tipo" in distintas:
        notas.append(f"contrato {car['tipo']}")
    if "inicio" in distintas:
        notas.append("inicio previsto el " + fecha_larga(car["inicio"]) if car["inicio"]
                     else "inicio por determinar")
    if "duracion" in distintas and car["duracion"]:
        notas.append("duración " + car["duracion"])
    if "exencion" in distintas and car["exencion"]:
        notas.append(car["exencion"])
    if "gdo" in distintas and car["gdo"]:
        notas.append("con garantías de origen (GdO)")
    if "excedentes" in distintas and car["excedentes"]:
        notas.append("con compensación de excedentes")
    return "; ".join(notas)


def _resumen(car: dict, distintas: set, suministros: list) -> list:
    lineas = [f"Número de suministros: {len(suministros)}, con un consumo anual previsto total "
              f"de {_kwh(sum(so.consumo_anual(s) for s in suministros))}."]
    fijo = ("precio fijo según la tarifa de cada suministro (binómica en DH2 para la 2.0TD y "
            "en DH6 para la 3.0TD y las 6.xTD)" if "fijo" in distintas else
            f"precio fijo ({car['fijo'][0].lower() + car['fijo'][1:]})")
    if "indexado" in distintas:
        tipologias = f"{fijo} y, según el suministro, precio indexado (ver listado)"
    elif car["indexado"]:
        tipologias = f"{fijo} y precio indexado ({', '.join(car['indexado'])})"
    else:
        tipologias = fijo
    lineas.append(f"Tipologías de oferta solicitadas: {tipologias}.")
    if "tipo" not in distintas:
        lineas.append(f"Tipo de contrato: {car['tipo']}.")
    if "inicio" in distintas:
        lineas.append("Fecha prevista de inicio del contrato: según el suministro (ver listado).")
    elif car["inicio"]:
        lineas.append(f"Fecha prevista de inicio del contrato: {fecha_larga(car['inicio'])}.")
    if "duracion" in distintas:
        lineas.append("Duración del contrato: según el suministro (ver listado).")
    elif car["duracion"]:
        lineas.append(f"Duración del contrato: {car['duracion']}.")
    if "exencion" not in distintas and car["exencion"]:
        lineas.append(f"Todos los suministros están {car['exencion'].replace('sujeto', 'sujetos', 1)}.")
    if "gdo" not in distintas and car["gdo"]:
        lineas.append("Las ofertas deben incluir garantías de origen (GdO).")
    if "excedentes" not in distintas and car["excedentes"]:
        lineas.append("Los suministros cuentan con compensación de excedentes.")
    return lineas


def texto_correo(cabecera: dict, suministros: list) -> str:
    """Texto para pegar en el correo de peticion de ofertas a los gestores de las
    comercializadoras."""
    cliente = _lista(_unicos([s.get("grupo") for s in suministros])
                     or _unicos([s.get("razon") for s in suministros]) or ["(cliente)"])
    cars = [_caracteristicas(s) for s in suministros]
    distintas = {k for k in cars[0] if len({repr(c[k]) for c in cars}) > 1}

    listado = []
    for s, car in zip(suministros, cars):
        linea = f"    • CUPS: {s.get('cups', '')}"
        if s.get("razon"):
            linea += f" ({s['razon']})"
        nota = _nota_cups(car, distintas)
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
        f"electricidad del cliente {cliente} en España.",
        f"Nos ponemos en contacto con usted para trasladarle el interés de {cliente} en recibir "
        f"ofertas de electricidad para los siguientes suministros ubicados en España:",
        "\n".join(listado),
        "Resumen de la solicitud:\n" + "\n".join(
            "    • " + l for l in _resumen(cars[0], distintas, suministros)),
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
