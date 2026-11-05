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


def apuntamientos_vigentes(e: mo.Estudio, calculados: list, hoy: dt.date = None,
                           solo_vacios: bool = False) -> None:
    """Apuntamientos estimados del estudio = los calculados de los 12 últimos meses (con su mes,
    año y fuente). solo_vacios: no toca los que ya tiene (los cambiados a mano se quedan)."""
    v = ap.vigentes(calculados, hoy)
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


def apuntamientos_cliente(e: mo.Estudio, curvas: dict, precios_de=None) -> dict:
    """Calcula el apuntamiento propio de cada CUPS con curva y lo guarda en el estudio (solo
    los válidos). Devuelve {cups: resultado} con la cobertura y el motivo de los no válidos."""
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
        else:
            e.ap_cliente.pop(cups, None)
            if s.apuntamiento == mo.CLIENTE:
                s.apuntamiento = mo.ESTIMADO
    return salida


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
    if any(o and o.ssaa_modo != mo.SSAA_INCLUIDOS for o in e.ofertas + [e.actual]):
        if any(v is None for v in e.ssaa_mes) and not e.ssaa_manual:
            salida.append(("error", "Faltan los SSAA de algún mes (sección Mercado)."))
    usados = {mo.clave_ap(s.sistema, mo.calendario_de(s.tarifa)) for s in e.suministros
              if s.activo and s.tarifa in mo.TARIFAS}
    for clave in sorted(usados):
        vacios = [m for m in range(1, 13) if not (e.ap_omie.get(clave) or {}).get(m)]
        if indexadas and vacios:
            salida.append(("aviso", f"Apuntamientos de {clave.replace('|', ' ')} vacíos en "
                                    f"{len(vacios)} meses: se usa 1."))
    for o in e.ofertas + ([e.actual] if e.actual else []):
        if not o.comercializadora:
            salida.append(("aviso", f"«{o.nombre}»: falta la comercializadora."))
        if o.email_gestor and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", o.email_gestor):
            salida.append(("aviso", f"«{o.nombre}»: el email del gestor no es válido."))
        grupos = {g for s, g in calculo.incluidos(e, o)[0]}
        if o.modalidad == mo.FIJO:
            sin = [g for g in grupos if not any(
                v is not None for p in (o.precios.get(g) or o.precios.get(g.split("-")[0]) or [])
                for v in p.energia)]
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
    mejor = min((f for f in filas if f["media_anual"]), key=lambda f: f["media_anual"],
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
