# -*- coding: utf-8 -*-
"""
CATÁLOGO DE OFERTAS (sin Streamlit)
===================================
Ofertas «de catálogo» de una comercializadora: fórmula y precios iguales para cualquier
cliente, válidos en una ventana de fechas. Se guardan una vez (almacen.catalogo_ofertas) y se
cargan en cualquier análisis como copia independiente que recuerda su origen.

Qué se guarda: solo lo genérico de la oferta. No se guardan el gestor ni su email ni nada
propio del cliente (grupos propios, CUPS asignados o excluidos, cierres de meses concretos,
fecha de inicio, apuntamientos propios).

    a_catalogo(o)                    datos de la oferta que se guardan
    ficha_de(o, nombre, fechas...)   cabecera de la entrada (columnas de la tabla)
    validar(ficha, datos)            lo que falta o no cuadra (lista de textos)
    desde_catalogo(entrada, contactos)   oferta lista para el análisis
    avisos_carga(entrada, e, hoy)    caducada, inicio fuera de ventana, tarifas sin precio, baja
    cubre(entrada, e)                ¿tiene precio para todas las tarifas y sistemas del análisis?
    desactualizadas(e, versiones)    ofertas del análisis con una versión más nueva en el catálogo
    actualizar_desde(o, entrada)     cambia lo genérico de la oferta y conserva lo del análisis
"""

import datetime as dt

from analisis_ofertas import modelo as mo

ACTIVA, BAJA = "Activa", "De baja"

# propios del cliente o del análisis: no van al catálogo
CAMPOS_EXCLUIDOS = ("nombre", "gestor", "email_gestor", "ronda", "es_actual", "inicio",
                    "cierres", "grupos_propios", "asignacion", "excluidos",
                    "apuntamientos_propios", "origen_catalogo")

GRUPOS_AUTOMATICOS = {mo.grupo_automatico(t, s) for t in mo.TARIFAS for s in mo.SISTEMAS}


# ------------------------------------------------------------------- guardar
def a_catalogo(o: mo.Oferta) -> dict:
    """Datos genéricos de la oferta (diccionario JSON) para guardarla en el catálogo."""
    d = mo.oferta_a_dict(o)
    for campo in CAMPOS_EXCLUIDOS:
        d.pop(campo, None)
    d["precios"] = {g: v for g, v in (d.get("precios") or {}).items() if g in GRUPOS_AUTOMATICOS}
    return d


def _tiene_valor(p: dict, campo: str) -> bool:
    valores = p.get(campo) or []
    if campo == "ci":
        return any(v not in (None, 1.0) for v in valores)
    return any(v is not None for v in valores)


def grupos_con_precio(datos: dict) -> list:
    """Grupos (tarifa o tarifa-sistema) que tienen algún precio en la oferta."""
    campos = ("energia",) if datos.get("modalidad") == mo.FIJO else \
        ("energia", "termino_b", "ci", "potencia")
    return sorted(g for g, anios in (datos.get("precios") or {}).items()
                  if any(_tiene_valor(p, c) for p in anios for c in campos))


def ficha_de(o: mo.Oferta, nombre: str, vigencia_desde, vigencia_hasta, inicio_desde,
             inicio_hasta, id_analisis=None) -> dict:
    """Cabecera de la entrada del catálogo (lo que se ve en el buscador)."""
    datos = a_catalogo(o)
    return {"comercializadora": (o.comercializadora or "").strip(), "nombre": (nombre or "").strip(),
            "modalidad": o.modalidad, "duracion": int(o.duracion or 0),
            "vigencia_desde": vigencia_desde, "vigencia_hasta": vigencia_hasta,
            "inicio_desde": inicio_desde, "inicio_hasta": inicio_hasta,
            "tarifas": grupos_con_precio(datos), "id_analisis_origen": id_analisis}


def validar(ficha: dict, datos: dict) -> list:
    """Lo imprescindible para guardar: comercializadora, nombre, modalidad, duración, las dos
    ventanas de fechas (coherentes) y precios (en fijo, de energía; en Pass-Pool, Ci o B)."""
    errores = []
    for clave, texto in (("comercializadora", "la comercializadora"),
                         ("nombre", "el nombre del catálogo o producto"),
                         ("modalidad", "la modalidad")):
        if not ficha.get(clave):
            errores.append(f"Falta {texto}.")
    if not ficha.get("duracion") or int(ficha["duracion"]) < 1:
        errores.append("Falta la duración (meses).")
    for desde, hasta, texto in (("vigencia_desde", "vigencia_hasta", "la vigencia"),
                                ("inicio_desde", "inicio_hasta",
                                 "la ventana de inicio de suministro")):
        a, b = ficha.get(desde), ficha.get(hasta)
        if not a or not b:
            errores.append(f"Faltan las fechas de {texto} (desde y hasta).")
        elif _fecha(a) > _fecha(b):
            errores.append(f"En {texto}, «desde» es posterior a «hasta».")
    if datos.get("modalidad") in (mo.FIJO, mo.PASS_POOL) and not grupos_con_precio(datos):
        errores.append("Faltan los precios: mete al menos los de una tarifa en la pestaña "
                       "Precios.")
    return errores


# ------------------------------------------------------------------- cargar
def _fecha(v):
    if v is None or isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v)[:10])


def desde_catalogo(entrada: dict, contactos: dict = None) -> mo.Oferta:
    """Oferta para el análisis a partir de una entrada del catálogo (copia independiente que
    recuerda de dónde viene). El gestor y su email se proponen de los contactos de la
    comercializadora (los últimos usados en los análisis)."""
    o = mo.oferta_de_dict(entrada["datos"])
    o.comercializadora = entrada.get("comercializadora") or o.comercializadora
    o.nombre = entrada.get("nombre") or o.producto or o.comercializadora
    o.origen_catalogo = {"id": entrada.get("id"), "version": entrada.get("version") or 1,
                         "nombre": entrada.get("nombre") or ""}
    c = (contactos or {}).get((o.comercializadora or "").upper())
    if c:
        o.gestor, o.email_gestor = c.get("gestor") or "", c.get("email") or ""
    return o


def _necesarios(e: mo.Estudio) -> set:
    return {(s.tarifa, s.sistema) for s in e.suministros if s.activo}


def sin_precio(entrada: dict, e: mo.Estudio) -> list:
    """Grupos del análisis sin precio en la entrada (un grupo territorial sin precios propios
    usa los de su tarifa, como en el cálculo). En indexados sin ningún precio (p. ej. Pass-
    Through con solo fee) no falta nada."""
    datos = entrada.get("datos") or {}
    con = set(entrada.get("tarifas") or grupos_con_precio(datos))
    if datos.get("modalidad", entrada.get("modalidad")) == mo.PASS_THROUGH and not con:
        return []
    return sorted(mo.grupo_automatico(t, s) for t, s in _necesarios(e)
                  if mo.grupo_automatico(t, s) not in con and t not in con)


def cubre(entrada: dict, e: mo.Estudio) -> bool:
    return not sin_precio(entrada, e)


def avisos_carga(entrada: dict, e: mo.Estudio, hoy: dt.date = None) -> list:
    """Avisos al cargar una entrada en un análisis (no impiden cargarla)."""
    hoy = hoy or dt.date.today()
    nombre = entrada.get("nombre") or entrada.get("comercializadora") or "Catálogo"
    avisos = []
    if entrada.get("estado") == BAJA:
        avisos.append(f"«{nombre}» está dada de baja en el catálogo.")
    hasta = _fecha(entrada.get("vigencia_hasta"))
    desde = _fecha(entrada.get("vigencia_desde"))
    if hasta and hoy > hasta:
        avisos.append(f"«{nombre}» caducó el {hasta:%d/%m/%Y}: confirma con la "
                      "comercializadora que sigue valiendo.")
    elif desde and hoy < desde:
        avisos.append(f"«{nombre}» aún no está vigente (desde el {desde:%d/%m/%Y}).")
    i0, i1 = _fecha(entrada.get("inicio_desde")), _fecha(entrada.get("inicio_hasta"))
    if e.inicio and i0 and i1 and not i0 <= e.inicio <= i1:
        avisos.append(f"«{nombre}»: el suministro empieza el {e.inicio:%d/%m/%Y}, fuera de la "
                      f"ventana de inicio del catálogo ({i0:%d/%m/%Y} - {i1:%d/%m/%Y}).")
    faltan = sin_precio(entrada, e)
    if faltan:
        avisos.append(f"«{nombre}»: sin precio para {', '.join(faltan)}. Complétalo en Precios "
                      "o excluye esos CUPS (sale en Controles).")
    return avisos


def vigente_en(entrada: dict, fecha: dt.date) -> bool:
    """La entrada se puede contratar en esa fecha (y no está de baja)."""
    d, h = _fecha(entrada.get("vigencia_desde")), _fecha(entrada.get("vigencia_hasta"))
    return entrada.get("estado", ACTIVA) != BAJA and (not d or d <= fecha) and \
        (not h or fecha <= h)


# ------------------------------------------------------------------- versiones
def desactualizadas(e: mo.Estudio, versiones: dict) -> list:
    """[(oferta, versión del catálogo)] de las ofertas cargadas del catálogo que tienen una
    versión más nueva. versiones: {id del catálogo: versión actual}."""
    salida = []
    for o in e.ofertas:
        origen = o.origen_catalogo or {}
        v = versiones.get(origen.get("id"))
        if v and v > int(origen.get("version") or 1):
            salida.append((o, v))
    return salida


def actualizar_desde(o: mo.Oferta, entrada: dict) -> None:
    """Pone en la oferta lo genérico de la versión actual del catálogo y conserva lo propio
    del análisis (nombre, gestor, ronda, inicio, cierres, grupos propios, CUPS...). Los precios
    de los grupos propios se conservan; los de los grupos automáticos se sustituyen."""
    nueva = mo.oferta_de_dict(entrada["datos"])
    propios = {g: v for g, v in o.precios.items() if g not in GRUPOS_AUTOMATICOS}
    for campo in a_catalogo(nueva):
        setattr(o, campo, getattr(nueva, campo))
    o.precios.update(propios)
    o.origen_catalogo = {"id": entrada.get("id"), "version": entrada.get("version") or 1,
                         "nombre": entrada.get("nombre") or ""}
