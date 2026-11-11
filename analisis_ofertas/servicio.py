# -*- coding: utf-8 -*-
"""
SERVICIO DEL ANALISIS DE OFERTAS (sin Streamlit)
================================================
    estudio_desde_solicitud    Estudio a partir de una solicitud de ofertas guardada (sus
                               suministros, potencias y perfil de 12 meses)
    estudio_desde_fichas       ídem desde fichas de Gemweb o de la plantilla de suministros
    preparar                   congela los datos regulados, el mercado (OMIP y OMIE) y los
                               apuntamientos vigentes
    apuntamientos_cliente      apuntamiento propio de cada CUPS con su curva de un año
    controles                  avisos antes de dar por bueno el análisis
    resumen / datos            lo que se guarda en almacen (y de vuelta)
"""

import datetime as dt
import json
import re
import zlib

import datos_mercado as dm
import datos_regulados as dr
from analisis_ofertas import apuntamientos as ap
from analisis_ofertas import calculo
from analisis_ofertas import modelo as mo
from analisis_ofertas import resultados

ESTADOS = ["En curso", "Adjudicada", "Descartada"]
VERSION = 1


# ------------------------------------------------------------------- suministros
def normalizar_tarifa(tarifa) -> str:
    """«6.1 TD», «6.1TD », «3.0TDVE» → nombre de la lista de tarifas (o el texto limpio)."""
    t = re.sub(r"\s+", "", str(tarifa or "")).upper().replace("TDVE", "TDVE")
    for nombre in mo.TARIFAS:
        if t == nombre.upper():
            return nombre
    return t


def sistema_de(direccion: str) -> str:
    """Sistema eléctrico a partir de la dirección (código postal o nombre de la isla, Ceuta o
    Melilla); si no se puede saber, Península."""
    from estudio_potencia import servicio as sp
    zona = None
    cp = sp._cp_en_texto(direccion)
    if cp:
        zona = sp.zona_por_codigo_postal(cp)
    zona = zona or sp.zona_por_nombre(direccion or "")
    return {"Illes Balears": "Baleares"}.get(zona, zona) or "Península"


def suministros_de_fichas(fichas: list) -> list:
    """Fichas de la solicitud (electricidad) → [Suministro]."""
    salida = []
    for f in fichas:
        if (f.get("tipo") or "electricidad") == "gas":
            continue
        perfil = {int(m): [float(v or 0) for v in mo.seis(p, 0.0)]
                  for m, p in (f.get("perfil") or {}).items()}
        salida.append(mo.Suministro(
            cups=f["cups"], tarifa=normalizar_tarifa(f.get("tarifa")),
            sistema=f.get("sistema") or sistema_de(f.get("direccion")),
            alias=f.get("alias") or "", razon=f.get("razon") or "",
            direccion=f.get("direccion") or "",
            potencias=[float(v or 0) for v in mo.seis(f.get("potencias"), 0.0)],
            consumo=perfil))
    return salida


def _primero_de_mes(fecha: dt.date) -> dt.date:
    return dt.date(fecha.year, fecha.month, 1)


def estudio_desde_fichas(fichas: list, cliente: str = "", inicio: dt.date = None,
                         hoy: dt.date = None) -> mo.Estudio:
    hoy = hoy or dt.date.today()
    inicio = _primero_de_mes(inicio) if inicio else mo.sumar_meses(_primero_de_mes(hoy), 1)
    e = mo.Estudio(cliente=cliente, fecha=hoy, inicio=inicio,
                   suministros=suministros_de_fichas(fichas))
    e.cif = next((f.get("cif") for f in fichas if f.get("cif")), "")
    e.actual = mo.Oferta(nombre="Contrato actual", es_actual=True, duracion=12)
    return e


def estudio_desde_solicitud(sol: dict, hoy: dt.date = None) -> mo.Estudio:
    """sol: almacen.solicitud(id). Inicio = la fecha prevista de la solicitud (primer día del
    mes); duraciones pedidas, informativas."""
    from ofertas import solicitud as so
    fichas = sol["datos"]["suministros"]
    inicio, duraciones = None, []
    for f in fichas:
        c = f.get("consideraciones") or {}
        if c.get("inicio") and not inicio:
            inicio = dt.datetime.strptime(c["inicio"], "%d/%m/%Y").date()
        for fin in so.plazos(c):
            meses = so.meses_de_plazo(c.get("inicio"), fin)
            if meses and meses not in duraciones:
                duraciones.append(meses)
    e = estudio_desde_fichas(fichas, sol.get("cliente") or "", inicio, hoy)
    e.id_solicitud = sol.get("id")
    e.duraciones_solicitadas = sorted(duraciones)
    return e


# ------------------------------------------------------------------- datos congelados
def anios_horizonte(e: mo.Estudio) -> list:
    n = e.horizonte()
    return sorted({mo.sumar_meses(e.inicio, k).year for k in range(n)})


def _fmt_regulado(v, unidad):
    if not isinstance(v, (int, float)):
        return str(v)
    if unidad == "tanto por uno":
        return f"{v * 100:.4f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
    return f"{v:.7f}".rstrip("0").rstrip(".").replace(".", ",") + f" {unidad}"


def diferencias_regulados(e: mo.Estudio, guardados: dict = None) -> list:
    """Lo que ha cambiado en Administración → Datos regulados desde que se congelaron en el
    análisis, año a año del horizonte: textos «FNEE 2027: 2,4374639 €/MWh → 2,658 €/MWh» (y uno
    por año si cambian peajes, cargos, pagos por capacidad, pérdidas o SRAD de alguna tarifa)."""
    salida = []
    for a in anios_horizonte(e):
        congelado = e.regulados.get(a)
        if congelado is None:
            continue
        actual = dr.datos_anio(a, guardados)
        g0, g1 = congelado.get("generales") or {}, actual.get("generales") or {}
        for clave, nombre, unidad in dr.GENERALES:
            v0, v1 = g0.get(clave), g1.get(clave)
            iguales = v0 == v1 or (isinstance(v0, (int, float)) and isinstance(v1, (int, float))
                                   and abs(v0 - v1) < 1e-12)
            if not iguales:
                salida.append(f"{nombre} {a}: {_fmt_regulado(v0, unidad)} → "
                              f"{_fmt_regulado(v1, unidad)}")
        if (congelado.get("tarifas") or {}) != (actual.get("tarifas") or {}):
            salida.append(f"Peajes, cargos, pagos por capacidad, pérdidas o SRAD de alguna "
                          f"tarifa en {a}")
    return salida


def congelar_regulados(e: mo.Estudio, guardados: dict = None, todos: bool = False) -> None:
    """Datos regulados de los años del horizonte (los que falten; todos=True los vuelve a
    tomar de Administración)."""
    for a in anios_horizonte(e):
        if todos or a not in e.regulados:
            e.regulados[a] = dr.datos_anio(a, guardados)


def congelar_mercado(e: mo.Estudio, ruta_omip=None, ruta_omie=None) -> list:
    """OMIP de la última sesión y media de OMIE para los meses del horizonte (los ajustes a mano
    se conservan). % OMIP y años de la media: los de los datos regulados del año de inicio."""
    reg = e.regulados.get(e.inicio.year) or dr.datos_anio(e.inicio.year)
    m = reg.get("mercado") or {"pct_omip": 0.95, "anios_media_omie": 3}
    try:
        c = dm.congelar(e.inicio, e.horizonte(), m["pct_omip"], m["anios_media_omie"],
                        ruta_omip=ruta_omip, ruta_omie=ruta_omie)
    except dm.DatosMercado as err:
        return [f"No se han podido leer los datos de mercado: {err}"]
    e.pct_omip = m["pct_omip"]
    e.sesion_omip = c["sesion_omip"]
    e.mercado = {f"{f['anio']}-{f['mes']:02d}": {k: f[k] for k in (
        "omip", "contrato", "omie_media", "anios_omie")} for f in c["meses"]}
    return [f"{ap.MESES[f['mes'] - 1][:3]}-{f['anio'] % 100:02d}: {f['aviso']}"
            for f in c["meses"] if f["aviso"]]


def completar_mercado(e: mo.Estudio, ruta_omip=None, ruta_omie=None) -> list:
    """Meses del horizonte sin precio de mercado (p. ej. al añadir una oferta más larga que las
    de cuando se congeló): se rellenan con los datos de mercado de Parámetros, de la misma
    sesión de OMIP que el resto del estudio. Los meses que ya tiene no se tocan."""
    congelar_regulados(e)
    faltan = [k for k in range(e.horizonte())
              if not (e.mercado.get(_clave_mes(e, k)) or {}).get("omie_media")
              and not (e.mercado.get(_clave_mes(e, k)) or {}).get("omip")]
    if not faltan:
        return []
    reg = e.regulados.get(e.inicio.year) or dr.datos_anio(e.inicio.year)
    m = reg.get("mercado") or {"pct_omip": 0.95, "anios_media_omie": 3}
    sesion = dt.date.fromisoformat(e.sesion_omip) if e.sesion_omip else None
    try:
        c = dm.congelar(e.inicio, e.horizonte(), e.pct_omip or m["pct_omip"],
                        m["anios_media_omie"], fecha_sesion=sesion, ruta_omip=ruta_omip,
                        ruta_omie=ruta_omie)
    except dm.DatosMercado as err:
        return [f"No se han podido completar los datos de mercado: {err}"]
    if not e.sesion_omip:
        e.sesion_omip, e.pct_omip = c["sesion_omip"], m["pct_omip"]
    nuevos = 0
    for f in c["meses"]:
        clave = f"{f['anio']}-{f['mes']:02d}"
        actual = e.mercado.get(clave) or {}
        if actual.get("omie_media") or actual.get("omip"):
            continue
        e.mercado[clave] = {k: f[k] for k in ("omip", "contrato", "omie_media", "anios_omie")}
        nuevos += 1
    return [f"Se han añadido {nuevos} meses de precio de mercado desde Parámetros (sesión de OMIP "
            f"del {dt.date.fromisoformat(c['sesion_omip']):%d/%m/%Y})."] if nuevos else []


def _clave_mes(e: mo.Estudio, k: int) -> str:
    f = mo.sumar_meses(e.inicio, k)
    return f"{f.year}-{f.month:02d}"


def ssaa_por_defecto(e: mo.Estudio, forzar: bool = False) -> bool:
    """SSAA estimados de Parámetros (datos regulados del año de inicio) en el estudio, si no
    tiene ninguno (o siempre con forzar). Devuelve si los ha puesto."""
    if not forzar and any(v is not None for v in e.ssaa_mes):
        return False
    valores = dr.ssaa_estimados(e.inicio.year)
    if not valores:
        return False
    e.ssaa_mes = list(valores)
    return True


def apuntamientos_vigentes(e: mo.Estudio, calculados: list, hoy: dt.date = None,
                           solo_vacios: bool = False) -> None:
    """Apuntamientos estimados del estudio = los calculados de los 12 últimos meses (con su mes,
    año y fuente). solo_vacios: no toca los que ya tiene (los cambiados a mano se quedan)."""
    v = ap.vigentes(calculados, hoy)
    for m, d in v.pop(ap.CLAVE_DIF_PH, {}).items():
        if solo_vacios and e.ph_dif.get(m) is not None:
            continue
        e.ph_dif[m] = d["factores"][0]
        e.ph_origen[m] = ap.etiqueta(d["anio"], m, d["fuente"])
    for m, d in v.pop(ap.CLAVE_PICO, {}).items():
        if not (solo_vacios and e.pico.get(m)):
            e.pico[m] = d["factores"][0]
    for clave, meses in v.items():
        for m, d in meses.items():
            if solo_vacios and (e.ap_omie.get(clave) or {}).get(m):
                continue
            e.ap_omie.setdefault(clave, {})[m] = d["factores"]
            e.ap_origen.setdefault(clave, {})[m] = ap.etiqueta(d["anio"], m, d["fuente"])


def preparar(e: mo.Estudio, calculados: list = None, hoy: dt.date = None) -> list:
    """Congela regulados, mercado y apuntamientos de un estudio nuevo. Devuelve avisos."""
    congelar_regulados(e)
    avisos = congelar_mercado(e)
    ssaa_por_defecto(e)
    if calculados is None:
        import almacen
        calculados = almacen.listar_apuntamientos()
    apuntamientos_vigentes(e, calculados, hoy)
    if not calculados:
        avisos.append("Todavía no hay apuntamientos calculados: se usará 1 (el precio medio) "
                      "hasta que los rellenes o se calculen.")
    return avisos


# ------------------------------------------------------------------- curvas del cliente
def empaquetar_curvas(curvas: dict) -> bytes:
    """{cups: {(fecha, hora): kWh}} → bytes comprimidos (para almacen; nunca a git)."""
    datos = {c: {f"{k[0].isoformat()}|{k[1]}": v for k, v in curva.items()}
             for c, curva in (curvas or {}).items()}
    return zlib.compress(json.dumps(datos).encode("utf-8"))


def desempaquetar_curvas(blob: bytes) -> dict:
    if not blob:
        return {}
    datos = json.loads(zlib.decompress(blob).decode("utf-8"))
    salida = {}
    for c, curva in datos.items():
        salida[c] = {}
        for k, v in curva.items():
            f, h = k.split("|")
            salida[c][(dt.date.fromisoformat(f), int(h))] = v
    return salida


def precios_ph_mes(anio: int, mes: int):
    """(Ph demanda de Canarias del mes, fuente) para el apuntamiento del cliente."""
    ph = ap.precios_ph(anio, mes)
    if not ap.completo(ph, anio, mes):
        raise ap.SinPrecios(f"Ph demanda {ap.MESES[mes - 1]} {anio} incompleto")
    return ph, "ESIOS Ph"


def apuntamientos_cliente(e: mo.Estudio, curvas: dict, precios_de=None, ph_de=None) -> dict:
    """Calcula el apuntamiento propio de cada CUPS con curva y lo guarda en el estudio (solo
    los válidos). En Canarias calcula también el del Ph demanda (para las ofertas indexadas al
    Ph; si no se puede, esas usan el estimado). Devuelve {cups: resultado} con la cobertura y
    el motivo de los no válidos."""
    salida = {}
    for cups, curva in (curvas or {}).items():
        s = e.suministro(cups)
        if not s:
            continue
        r = ap.apuntamiento_cliente(curva, s.sistema, s.tarifa, precios_de)
        salida[cups] = r
        if r["valida"]:
            cob = r["cobertura"]
            e.ap_cliente[cups] = {"omie": r["factores"], "fuentes": r["fuentes"],
                                  "desde": cob["desde"].isoformat(),
                                  "hasta": cob["hasta"].isoformat(), "dias": cob["dias"]}
            if s.sistema == "Canarias":
                rph = ap.apuntamiento_cliente(curva, s.sistema, s.tarifa,
                                              ph_de or precios_ph_mes)
                if rph["valida"]:
                    e.ap_cliente[cups]["ph"] = rph["factores"]
        else:
            e.ap_cliente.pop(cups, None)
            if s.apuntamiento == mo.CLIENTE:
                s.apuntamiento = mo.ESTIMADO
    return salida


MESES_BUSQUEDA_CURVA = 24          # meses hacia atrás en los que se busca en Gemweb


def periodo_curva(hoy: dt.date = None, meses: int = MESES_BUSQUEDA_CURVA):
    """(desde, hasta) que se pide a Gemweb: los últimos `meses` meses hasta ayer. De ahí se
    toman los 12 últimos meses completos que tenga la curva."""
    hoy = hoy or dt.date.today()
    return mo.sumar_meses(dt.date(hoy.year, hoy.month, 1), -meses), hoy - dt.timedelta(days=1)


def a_horaria(curva) -> dict:
    """Curva (horaria o cuartohoraria, claves ESIOS) → {(fecha, hora): kWh}."""
    if curva.resolucion != "qh":
        return dict(curva.valores)
    salida = {}
    for (f, h, _), v in curva.valores.items():
        salida[(f, h)] = salida.get((f, h), 0.0) + v
    return salida


def comprobar_curvas_gemweb(e: mo.Estudio, cliente, hoy: dt.date = None, precios_de=None,
                            al_avanzar=None):
    """Busca en Gemweb la curva de todos los suministros del estudio (últimos 24 meses), toma
    los 12 últimos meses completos que tenga cada una, comprueba si es un año completo y
    calcula su apuntamiento. NO cambia lo que usa cada CUPS: lo decide el usuario después,
    viendo el desde y el hasta (elegir_apuntamiento). cliente: gemweb.ClienteGemweb.
    al_avanzar(fraccion, texto).
    -> ({cups: curva horaria de esos 12 meses}, [fila por CUPS: CUPS, Alias, En Gemweb,
    Datos desde, Datos hasta, Año usado desde, Año usado hasta, Días con dato, Se puede
    usar, Motivo])."""
    desde, hasta = periodo_curva(hoy)
    curvas, filas = {}, []
    activos = [s for s in e.suministros if s.activo]
    for i, s in enumerate(activos):
        if al_avanzar:
            al_avanzar(i / max(1, len(activos)), f"{s.cups}: buscando la curva en Gemweb…")
        fila = {"CUPS": s.cups, "Alias": s.alias, "En Gemweb": "No", "Datos desde": None,
                "Datos hasta": None, "Año usado desde": None, "Año usado hasta": None,
                "Días con dato": 0, "Se puede usar": "No", "Motivo": ""}
        try:
            curva, _ = cliente.curva(s.cups, desde, hasta)
        except Exception as err:                            # noqa: BLE001 (GemwebError, red)
            fila["Motivo"] = str(err)
            filas.append(fila)
            continue
        fila["En Gemweb"] = "Sí"
        horaria = a_horaria(curva)
        dias = sorted({k[0] for k in horaria})
        if dias:
            fila["Datos desde"], fila["Datos hasta"] = dias[0], dias[-1]
        cob = ap.cobertura(horaria)
        if cob.get("meses"):
            fila["Año usado desde"], fila["Año usado hasta"] = cob["desde"], cob["hasta"]
            # solo los 12 meses que se usan (se guardan con el análisis)
            horaria = {k: v for k, v in horaria.items() if (k[0].year, k[0].month)
                       in set(cob["meses"])}
        curvas[s.cups] = horaria
        r = apuntamientos_cliente(e, {s.cups: horaria}, precios_de)[s.cups]
        fila["Días con dato"] = r["cobertura"].get("dias", 0)
        fila["Se puede usar"] = "Sí" if r["valida"] else "No"
        fila["Motivo"] = r["motivo"] or "Año completo: puedes usar su apuntamiento."
        filas.append(fila)
    if al_avanzar:
        al_avanzar(1.0, "Comprobación terminada.")
    return curvas, filas


def elegir_apuntamiento(e: mo.Estudio, usar: dict) -> list:
    """Aplica lo que decide el usuario: {cups: True (su curva) / False (el estimado)}. Solo se
    puede usar la curva si su apuntamiento está calculado (año completo). Devuelve los CUPS
    que no se han podido poner con su curva."""
    sin_curva = []
    for cups, quiere in usar.items():
        s = e.suministro(cups)
        if not s:
            continue
        if quiere and not e.ap_cliente.get(cups):
            sin_curva.append(cups)
            quiere = False
        s.apuntamiento = mo.CLIENTE if quiere else mo.ESTIMADO
    return sin_curva


# ------------------------------------------------------------------- controles
def controles(e: mo.Estudio, res: dict = None) -> list:
    """[(nivel, texto)] con nivel 'error' (el resultado no es fiable) o 'aviso'."""
    salida = []
    for s in e.suministros:
        if not s.activo:
            continue
        if s.tarifa not in mo.TARIFAS:
            salida.append(("error", f"{s.cups}: tarifa «{s.tarifa}» desconocida."))
            continue
        if s.sistema not in mo.SISTEMAS:
            salida.append(("error", f"{s.cups}: sistema eléctrico «{s.sistema}» desconocido."))
            continue
        if not s.consumo_anual():
            salida.append(("error", f"{s.cups}: sin consumo."))
        if not any(s.potencias):
            salida.append(("aviso", f"{s.cups}: sin potencias contratadas."))
        if mo.n_periodos(s.tarifa) == 6:
            malos = [f"{ap.MESES[m - 1][:3]} P{p + 1}" for m in range(1, 13)
                     for p, v in enumerate(s.consumo_mes(m))
                     if v and p + 1 not in dr.periodos_validos(s.sistema, s.tarifa, m)]
            if malos:
                salida.append(("aviso", f"{s.cups}: consumo en periodos que no existen ese mes en "
                                        f"{s.sistema} (Circular 3/2020): " + ", ".join(malos[:8])
                               + ("…" if len(malos) > 8 else "")))
    ctx = calculo.Contexto(e)
    copiados = ctx.anios_copiados(anios_horizonte(e)) if e.inicio else {}
    for a, de in copiados.items():
        salida.append(("aviso", f"Datos regulados de {a} sin publicar: se usan los de {de}."))
    n = e.horizonte()
    sin_omip, sin_omie = [], []
    for k in range(n):
        f = mo.sumar_meses(e.inicio, k)
        fila = e.mercado.get(f"{f.year}-{f.month:02d}") or {}
        etiqueta = f"{ap.MESES[f.month - 1][:3]}-{f.year % 100:02d}"
        if not (fila.get("omip") or e.omip_manual.get(f"{f.year}-{f.month:02d}")):
            sin_omip.append(etiqueta)
        if fila.get("omie_media") is None:
            sin_omie.append(etiqueta)
    indexadas = [o for o in e.ofertas + [e.actual] if o and o.indexada]
    if indexadas and sin_omip:
        salida.append(("aviso", "Meses sin cotización de OMIP (se usa la media de OMIE): "
                       + ", ".join(sin_omip)))
    if indexadas and sin_omie:
        salida.append(("error", "Meses sin media de OMIE: " + ", ".join(sin_omie)))
    sin_dif_ph = len([m for m in range(1, 13) if e.ph_dif.get(m) is not None]) < 12
    con_ph = [o for o in e.ofertas + [e.actual] if o and o.indexada and o.indice_ph]
    if any(o and o.ssaa_modo != mo.SSAA_INCLUIDOS for o in e.ofertas + [e.actual]) or (
            con_ph and sin_dif_ph):
        if any(v is None for v in e.ssaa_mes) and not e.ssaa_manual:
            salida.append(("error", "Faltan los SSAA de algún mes (sección Mercado)."))
    usados = {mo.clave_ap(s.sistema, mo.calendario_de(s.tarifa)) for s in e.suministros
              if s.activo and s.tarifa in mo.TARIFAS}
    for clave in sorted(usados):
        vacios = [m for m in range(1, 13) if not (e.ap_omie.get(clave) or {}).get(m)]
        if indexadas and vacios:
            salida.append(("aviso", f"Apuntamientos de {clave.replace('|', ' ')} vacíos en "
                                    f"{len(vacios)} meses: se usa 1."))
    if con_ph and sin_dif_ph:
        salida.append(("aviso", "Falta el Ph demanda de Canarias calculado con ESIOS en algún "
                                "mes: en esos meses el Ph se estima como mercado + SSAA del "
                                "estudio."))
    for o in e.ofertas + ([e.actual] if e.actual else []):
        canarias = [s for s, _ in calculo.incluidos(e, o)[0] if s.sistema == "Canarias"]
        if o.indexada and o.indice_ph:
            if not canarias:
                salida.append(("aviso", f"«{o.nombre}»: indexada al Ph demanda de Canarias, "
                                        "pero no tiene CUPS de Canarias (no cambia nada)."))
            elif o.ssaa_modo != mo.SSAA_INCLUIDOS:
                salida.append(("aviso", f"«{o.nombre}»: posible doble cobro. El Ph demanda de "
                                        "Canarias ya incluye los costes de ajuste (SSAA) y la "
                                        f"oferta añade además SSAA ({o.ssaa_modo}). Revisa "
                                        "la oferta."))
        if not o.comercializadora:
            salida.append(("aviso", f"«{o.nombre}»: falta la comercializadora."))
        if o.email_gestor and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", o.email_gestor):
            salida.append(("aviso", f"«{o.nombre}»: el email del gestor no es válido."))
        grupos = {g for s, g in calculo.incluidos(e, o)[0]}
        if o.modalidad == mo.FIJO:
            tarifa_de = {g: t for g, (t, _) in calculo.grupos_de(e, o).items()}
            sin = [g for g in grupos if not any(
                v is not None for y in range(1, 2 + len(o.precios.get(g) or []))
                for v in o.precios_grupo(g, y, tarifa_de.get(g)).energia)]
            if sin:
                salida.append(("error", f"«{o.nombre}»: sin precio de energía para "
                                        + ", ".join(sorted(sin)) + "."))
        if (not o.es_actual and e.duraciones_solicitadas
                and o.duracion not in e.duraciones_solicitadas):
            salida.append(("aviso", f"«{o.nombre}»: dura {o.duracion} meses (se pidieron "
                                    + ", ".join(str(d) for d in e.duraciones_solicitadas)
                                    + "). Solo informativo."))
    if res:
        for r in res["ofertas"] + ([res["actual"]] if res.get("actual") else []):
            for cups, texto in r.incidencias.items():
                salida.append(("error", f"«{r.oferta.nombre}» {cups}: {texto}."))
    return salida


# ------------------------------------------------------------------- almacen
def resumen(e: mo.Estudio, res: dict, estado: str = "En curso", ref: str = "") -> dict:
    filas = resultados.comparativa(res)
    mejor = min((f for f in filas if f["media_anual"]), key=resultados.orden,
                default=None)
    return {"id_solicitud": e.id_solicitud, "ref": ref, "cliente": e.cliente,
            "cups": [s.cups for s in e.suministros if s.activo],
            "consumo": round(sum(s.consumo_anual() for s in e.suministros if s.activo)),
            "inicio": e.inicio.isoformat() if e.inicio else None, "ofertas": len(e.ofertas),
            "ronda": max((o.ronda for o in e.ofertas), default=1),
            "comercializadoras": [o.comercializadora for o in e.ofertas],
            "mejor": (f"{mejor['comercializadora']} {mejor['duracion']} m" if mejor else None),
            "ahorro": mejor["ahorro"] if mejor else None, "estado": estado}


def datos(e: mo.Estudio, ref: str = "") -> dict:
    return {"version": VERSION, "ref": ref, "estudio": mo.estudio_a_dict(e)}


def estudio_de_datos(d: dict) -> mo.Estudio:
    return mo.estudio_de_dict(d.get("estudio", d))


def adjudicar_solicitud(id_solicitud: int, comercializadora: str, usuario: str) -> None:
    """Marca la solicitud «Adjudicada» añadiendo la comercializadora a las que ya tenía."""
    import almacen
    sol = almacen.solicitud(id_solicitud)
    if not sol:
        return
    previas = [c.strip() for c in (sol.get("comercializadoras") or "").split(",") if c.strip()]
    if comercializadora and comercializadora.upper() not in {c.upper() for c in previas}:
        previas.append(comercializadora)
    almacen.actualizar_seguimiento(id_solicitud, "Adjudicada", ", ".join(previas), usuario)
