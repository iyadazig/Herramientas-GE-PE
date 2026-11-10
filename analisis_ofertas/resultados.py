# -*- coding: utf-8 -*-
"""
RESULTADOS DEL ANALISIS DE OFERTAS
==================================
Tablas para la app, el informe y el correo:
    comparativa       una fila por oferta: coste medio anual en su duración frente al contrato
                      actual en los mismos meses, €/kWh, ahorro, variación y ranking
    mejores           oferta de menor coste anual de cada tipo (fijo / indexado)
    cuadro_mercado    precio previsto, OMIP, OMIE y SSAA por trimestre del horizonte
    escenarios        coste de cada oferta con OMIP ± X %
    apuntamientos     indexadas con el apuntamiento estimado y con el del cliente
    rondas            evolución de cada comercializadora entre rondas
Sin Streamlit.
"""

import statistics

from analisis_ofertas import calculo
from analisis_ofertas import modelo as mo


def _cierres(o: mo.Oferta) -> str:
    return "" if o.cierres_tipo == mo.SIN_CIERRES else (
        "cierres %" if o.cierres_tipo == mo.CIERRE_PCT else "cierres MW")


def tipo(o: mo.Oferta) -> str:
    return "Fijo" if o.modalidad == mo.FIJO else "Indexado"


def comparativa(res: dict) -> list:
    """Una fila por oferta (en el orden del estudio)."""
    actual = res.get("actual")
    filas = []
    for r in res["ofertas"]:
        o = r.oferta
        act = actual.total_en(r.meses) * 12 / r.n_meses if actual and r.n_meses else None
        filas.append({
            "nombre": o.nombre, "comercializadora": o.comercializadora, "producto": o.producto,
            "gestor": o.gestor, "email_gestor": o.email_gestor, "ronda": o.ronda,
            "modalidad": o.modalidad, "tipo": tipo(o), "ssaa": o.resumen_ssaa(),
            "cierres": _cierres(o), "duracion": r.n_meses, "total": r.total, "mwh": r.mwh,
            "media_anual": r.media_anual, "eur_kwh": r.eur_mwh / 1000 if r.mwh else None,
            "actual_mismo_periodo": act,
            "ahorro": (act - r.media_anual) if act else None,
            "variacion": (r.media_anual / act - 1) if act else None,
            "incidencias": len(r.incidencias), "apuntamiento_cliente": len(r.apuntamiento_cliente),
            "inicio": r.meses[0] if r.meses else None})
    con_variacion = sorted((f for f in filas if f["variacion"] is not None), key=orden)
    for n, f in enumerate(con_variacion, start=1):
        f["ranking"] = n
    for f in filas:
        f.setdefault("ranking", None)
    return filas


def orden(f: dict):
    """Criterio del ranking y de la mejor oferta: la variación frente al contrato actual en los
    mismos meses. Con duraciones distintas el coste medio anual no es comparable (una oferta a 36
    meses promedia años más baratos que una a 24). Sin contrato actual, el coste medio anual."""
    v = f.get("variacion")
    return (v is None, v if v is not None else 0.0, f.get("media_anual") or 0.0)


def mejores(filas: list) -> dict:
    """{"Fijo": fila, "Indexado": fila} con más ahorro frente al contrato actual (en %)."""
    salida = {}
    for t in ("Fijo", "Indexado"):
        de_tipo = [f for f in filas if f["tipo"] == t and f["media_anual"]]
        if de_tipo:
            salida[t] = min(de_tipo, key=orden)
    return salida


def cuadro_mercado(e: mo.Estudio, ctx: calculo.Contexto = None, n_meses: int = None) -> list:
    """Precio previsto, OMIP, media de OMIE y SSAA (€/MWh) por trimestre del horizonte."""
    ctx = ctx or calculo.Contexto(e)
    n = n_meses or e.horizonte()
    trimestres = {}
    for k in range(n):
        f = mo.sumar_meses(e.inicio, k)
        clave = (f.year, (f.month - 1) // 3 + 1)
        previsto, omip = ctx.mercado(f.year, f.month)
        fila = e.mercado.get(f"{f.year}-{f.month:02d}") or {}
        t = trimestres.setdefault(clave, {"previsto": [], "omip": [], "omie": [], "ssaa": []})
        t["previsto"].append(previsto)
        if omip:
            t["omip"].append(omip)
        if fila.get("omie_media") is not None:
            t["omie"].append(fila["omie_media"])
        t["ssaa"].append(ctx.ssaa(f.year, f.month))
    return [{"trimestre": f"Q{q}-{a % 100:02d}", "anio": a, "q": q,
             **{k: (statistics.mean(v) if v else None) for k, v in t.items()}}
            for (a, q), t in sorted(trimestres.items())]


def _copia(e: mo.Estudio) -> mo.Estudio:
    return mo.estudio_de_dict(mo.estudio_a_dict(e))


def escenarios(e: mo.Estudio, variaciones=(-0.2, -0.1, 0.0, 0.1, 0.2)) -> list:
    """Coste medio anual y variación frente al actual de cada oferta con el OMIP de todos los
    meses multiplicado por (1 + x). [{"variacion_omip": x, "filas": comparativa}]."""
    salida = []
    for x in variaciones:
        c = _copia(e)
        for clave, fila in c.mercado.items():
            omip = c.omip_manual.get(clave, fila.get("omip"))
            if omip:
                c.omip_manual[clave] = omip * (1 + x)
        salida.append({"variacion_omip": x, "filas": comparativa(calculo.calcular(c))})
    return salida


def apuntamientos(e: mo.Estudio) -> list:
    """Indexadas: coste medio anual con el apuntamiento estimado en todos los CUPS y con el del
    cliente en los que tienen curva válida. Vacío si ningún CUPS tiene curva válida."""
    con_curva = [s.cups for s in e.suministros if e.ap_cliente.get(s.cups)]
    if not con_curva:
        return []
    est, cli = _copia(e), _copia(e)
    for s in est.suministros:
        s.apuntamiento = mo.ESTIMADO
    for s in cli.suministros:
        s.apuntamiento = mo.CLIENTE if s.cups in con_curva else mo.ESTIMADO
    r_est, r_cli = calculo.calcular(est), calculo.calcular(cli)
    salida = []
    for a, b in zip(r_est["ofertas"], r_cli["ofertas"]):
        if not a.oferta.indexada:
            continue
        salida.append({"nombre": a.oferta.nombre, "comercializadora": a.oferta.comercializadora,
                       "estimado": a.media_anual, "cliente": b.media_anual,
                       "diferencia": b.media_anual - a.media_anual,
                       "variacion": (b.media_anual / a.media_anual - 1) if a.media_anual else None,
                       "cups_con_curva": len(con_curva)})
    return salida


def rondas(filas: list) -> list:
    """Comercializadoras con más de una ronda: [{"comercializadora", "rondas": [filas]}]."""
    por_com = {}
    for f in filas:
        por_com.setdefault((f["comercializadora"] or "").strip().upper(), []).append(f)
    salida = []
    for com, lista in por_com.items():
        if len({f["ronda"] for f in lista}) > 1:
            salida.append({"comercializadora": lista[0]["comercializadora"],
                           "rondas": sorted(lista, key=lambda f: (f["ronda"], f["duracion"]))})
    return salida
