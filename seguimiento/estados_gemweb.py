# -*- coding: utf-8 -*-
"""
Estados del cuadro de seguimiento a partir de las facturas de Gemweb
====================================================================
Comprueba que las facturas emitidas en Gemweb esten señaladas en el cuadro y PROPONE los
cambios; no se aplica nada hasta que alguien los acepta en el cuadro (todos o los
seleccionados). La busqueda se hace con el boton «Buscar cambios en Gemweb» y a las 7:00 y a
las 12:00 (tarea programada del servidor), que deja la propuesta preparada:

    python -m seguimiento.estados_gemweb             (busca y guarda la propuesta)
    python -m seguimiento.estados_gemweb --aplicar   (busca y aplica todos los cambios)

1. Una consulta con las facturas cuyo periodo termina desde el 1 de enero del año en curso
   (las de periodos del año anterior quedan para mas adelante) y otra con los suministros,
   para pasar de su id al CUPS. Las anuladas, simuladas, abonos, abonadas y complementarias
   se listan pero no cuentan.
2. Cada factura cuenta en los meses del año que cubre en mas de la mitad de sus dias (una
   bimestral, en dos meses).
3. Regla, que nunca hace retroceder un estado:
     vacio o «Sin fra»  -> «Recibida» (factura sin verificar) o «Pdte inf» (verificada)
     «Recibida»         -> «Pdte inf» si esta verificada
   «Inf enviado», «Inf no enviado», «Temporal», «Cadencia» y cualquier otro no se tocan.
Solo los suministros activos «Sgto. y Control» del cuadro cuyo CUPS esta en Gemweb.

Aparte, cups_fuera_de_gemweb() da los suministros del cuadro cuyo CUPS no esta en Gemweb.
"""

import argparse
import datetime as dt
import json
import logging
import sys
from pathlib import Path

if __package__ in (None, ""):                         # python seguimiento\estados_gemweb.py
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
from seguimiento import columnas as C
from seguimiento import datos as D

EXCLUIR = {"anulada": "anulada", "simulada": "simulada", "abonament": "abono",
           "abonada": "abonada", "complementaria": "complementaria"}
USUARIO = "Gemweb automático"
PROPUESTA = "gemweb_propuesta"                 # clave de sgc_ajustes con la ultima busqueda
LOG = logging.getLogger("estados_gemweb")

# resultado de cada factura y mes en el listado
CAMBIA, SIN_CAMBIOS, OTRA_FACTURA = ("Cambio propuesto", "Ya señalada en el cuadro",
                                     "Ya señalada (cuenta otra factura del mes)")
NO_SYC, DESCARTADA = "No es «Sgto. y Control»", "Descartada"
COLUMNAS_LISTADO = ["CUPS", "Denominación suministro", "Factura", "Comercializadora", "Emisión",
                    "Periodo inicio", "Periodo fin", "Verificada", "Mes", "Estado en el cuadro",
                    "Estado propuesto", "Resultado"]


def _fecha(texto):
    try:
        return dt.datetime.strptime((texto or "").strip()[:10], "%d/%m/%Y").date()
    except ValueError:
        try:
            return dt.date.fromisoformat((texto or "").strip()[:10])
        except ValueError:
            return None


def desde(hoy):
    """1 de enero del año en curso."""
    return dt.date(hoy.year, 1, 1)


def _cliente():
    import gemweb
    credenciales, _ = gemweb.cargar_credenciales()
    if not credenciales:
        raise RuntimeError("No hay credenciales de Gemweb.")
    return gemweb.obtener_api(*credenciales)


def cups_gemweb(cliente):
    """{id del suministro en Gemweb: CUPS de 20 caracteres}."""
    cups = {}
    for s in cliente.get_inventory(category="subministraments", limit=100000):
        d = {h.tag: h.text for h in s}
        if d.get("cups"):
            cups[d.get("id")] = d["cups"].strip().upper()[:20]
    return cups


def facturas_gemweb(cliente, inicio, cups=None):
    """Facturas de Gemweb cuyo periodo termina desde `inicio`: dicts cups (20 caracteres),
    factura, comercializadora, emision, inicio, fin, verificada y descartada (motivo o None)."""
    cups = cups_gemweb(cliente) if cups is None else cups
    salida = []
    for f in cliente.get_inventory(category="factures", search_by="factures.periode_fin",
                                   search_values=inicio.isoformat(), search_operator=">=",
                                   limit=100000):
        d = {h.tag: h.text for h in f}
        ini, fin = _fecha(d.get("periode_ini")), _fecha(d.get("periode_fin"))
        if d.get("id_subministraments") not in cups or not ini or not fin:
            continue
        motivos = [t for m, t in EXCLUIR.items()
                   if str(d.get(m) or "0").strip() not in ("0", "")]
        salida.append({"cups": cups[d["id_subministraments"]], "factura": d.get("num_factura"),
                       "comercializadora": d.get("comercialitzadora"),
                       "emision": _fecha(d.get("data_emissio")), "inicio": ini, "fin": fin,
                       "verificada": str(d.get("verificada") or "0").strip() == "1",
                       "descartada": ", ".join(motivos) or None})
    return salida


def nuevo_estado(actual, verificada):
    """Estado que corresponde con una factura (None si no cambia)."""
    if actual in (None, "", C.SIN_FRA):
        return C.PDTE_INF if verificada else C.RECIBIDA
    if actual == C.RECIBIDA and verificada:
        return C.PDTE_INF
    return None


def _es_syc(s):
    return (s["datos"].get("Tipo contratación") or "").strip() == C.SGTO_Y_CONTROL


def revisar(facturas, por_cups, estados, inicio):
    """(cambios, listado). cambios: dicts id_suministro, cups, denominacion, anio, mes, antes,
    despues, factura y verificada. listado: una fila por factura, suministro y mes de las
    facturas de los CUPS que estan en el cuadro, con lo que ha encontrado."""
    mejor = {}                                    # (id, anio, mes) -> factura que cuenta
    for f in facturas:
        if f["descartada"]:
            continue
        for anio, mes in D.meses_cubiertos(f["inicio"], f["fin"]):
            if dt.date(anio, mes, 1) < inicio:
                continue
            for s in por_cups.get(f["cups"], []):
                clave = (s["id"], anio, mes)
                if _es_syc(s) and (clave not in mejor or (
                        f["verificada"] and not mejor[clave]["verificada"])):
                    mejor[clave] = f
    cambios, nuevos = [], {}
    for (id_, anio, mes), f in sorted(mejor.items(), key=lambda x: x[0]):
        antes = (estados.get((id_, anio, mes)) or {}).get("estado")
        despues = nuevo_estado(antes, f["verificada"])
        if despues:
            s = next(x for x in por_cups[f["cups"]] if x["id"] == id_)
            cambios.append({"id_suministro": id_, "cups": f["cups"],
                            "denominacion": s["datos"].get("Denominación suministro"),
                            "anio": anio, "mes": mes, "antes": antes, "despues": despues,
                            "factura": f["factura"], "verificada": f["verificada"]})
            nuevos[(id_, anio, mes)] = despues

    listado = []
    for f in sorted(facturas, key=lambda x: (x["cups"], x["inicio"], str(x["factura"]))):
        suministros = por_cups.get(f["cups"], [])
        if not suministros:
            continue                               # CUPS de Gemweb que no estan en el cuadro
        base = {"CUPS": f["cups"], "Factura": f["factura"],
                "Comercializadora": f["comercializadora"], "Emisión": f["emision"],
                "Periodo inicio": f["inicio"], "Periodo fin": f["fin"],
                "Verificada": "Sí" if f["verificada"] else "No"}
        if f["descartada"]:
            listado.append({**base, "Denominación suministro": suministros[0]["datos"].get(
                "Denominación suministro"), "Resultado": "%s (%s)" % (DESCARTADA,
                                                                       f["descartada"])})
            continue
        meses = [(a, m) for a, m in D.meses_cubiertos(f["inicio"], f["fin"])
                 if dt.date(a, m, 1) >= inicio]
        for s in suministros:
            for anio, mes in meses:
                clave = (s["id"], anio, mes)
                fila = {**base, "Denominación suministro": s["datos"].get(
                    "Denominación suministro"), "Mes": C.columna_mes(anio, mes),
                    "Estado en el cuadro": (estados.get(clave) or {}).get("estado")}
                if not _es_syc(s):
                    fila["Resultado"] = NO_SYC
                elif mejor.get(clave) is not f:
                    fila["Resultado"] = OTRA_FACTURA
                elif clave in nuevos:
                    fila["Resultado"], fila["Estado propuesto"] = CAMBIA, nuevos[clave]
                else:
                    fila["Resultado"] = SIN_CAMBIOS
                listado.append(fila)
    return cambios, listado


def buscar(hoy=None, cliente=None):
    """Busca en Gemweb sin cambiar el cuadro. dict con fecha, inicio, facturas (nº),
    descartadas (nº), cambios (propuestos) y listado."""
    hoy = hoy or dt.date.today()
    cliente = cliente or _cliente()
    inicio = desde(hoy)
    facturas = facturas_gemweb(cliente, inicio)
    estados = almacen.listar_sgc_estados(inicio.year)
    cambios, listado = revisar(facturas, D.suministros_por_cups(), estados, inicio)
    LOG.info("%d facturas de Gemweb desde %s; %d cambios propuestos", len(facturas), inicio,
             len(cambios))
    return {"fecha": dt.datetime.now().isoformat(timespec="seconds"),
            "inicio": inicio.isoformat(), "facturas": len(facturas),
            "descartadas": sum(1 for f in facturas if f["descartada"]),
            "cambios": cambios, "listado": listado}


def guardar_propuesta(resultado):
    almacen.guardar_ajuste_sgc(PROPUESTA, json.dumps(resultado, ensure_ascii=False,
                                                     default=str))


def propuesta():
    """La ultima busqueda guardada (con los cambios aun sin aplicar) o None."""
    texto = almacen.ajuste_sgc(PROPUESTA)
    return json.loads(texto) if texto else None


def aplicar(cambios, usuario):
    """Aplica los cambios aceptados. Solo los que siguen teniendo el estado de cuando se
    buscaron (si alguien ha cambiado la celda despues, no se pisa). Los aplicados se quitan
    de la propuesta guardada. Devuelve (aplicados, no aplicados)."""
    estados = {}
    for anio in {c["anio"] for c in cambios}:
        estados.update(almacen.listar_sgc_estados(anio))
    hechos, saltados = [], []
    for c in cambios:
        previo = estados.get((c["id_suministro"], c["anio"], c["mes"])) or {}
        (hechos if previo.get("estado") == c["antes"] else saltados).append(
            dict(c, obs=previo.get("obs")))
    if hechos:
        almacen.guardar_sgc_estados(
            [{"id_suministro": c["id_suministro"], "anio": c["anio"], "mes": c["mes"],
              "estado": c["despues"], "obs": c["obs"], "origen": "Gemweb"} for c in hechos],
            usuario, origen="Gemweb")
        almacen.guardar_cambios_gemweb(hechos)
        almacen.registrar_actividad(usuario, "Seguimiento y Control",
                                    "Aplicar estados de Gemweb", "%d cambios" % len(hechos))
    guardada = propuesta()
    if guardada:
        quitar = {(c["id_suministro"], c["anio"], c["mes"]) for c in hechos + saltados}
        guardada["cambios"] = [c for c in guardada["cambios"]
                               if (c["id_suministro"], c["anio"], c["mes"]) not in quitar]
        guardar_propuesta(guardada)
    return hechos, saltados


def cups_fuera_de_gemweb(cliente=None):
    """Suministros activos del cuadro cuyo CUPS no esta en Gemweb (o no tienen CUPS)."""
    en_gemweb = set(cups_gemweb(cliente or _cliente()).values())
    salida = []
    for s in almacen.listar_sgc_suministros(bajas=False):
        cups = (s["cups"] or "").strip().upper()[:20]
        if cups not in en_gemweb:
            d = s["datos"]
            salida.append({"Técnico": d.get("Técnico"), "Auxiliar": d.get("Auxiliar"),
                           "Tipo contratación": d.get("Tipo contratación"),
                           "Cliente/Grupo": d.get("Cliente/Grupo"),
                           "Denominación suministro": d.get("Denominación suministro"),
                           "CUPS": s["cups"] or "(sin CUPS)",
                           "Tipo suministro": d.get("Tipo suministro"),
                           "Comercializadora": d.get("Comercializadora"),
                           "Validación Gemweb": d.get("Validación Gemweb"),
                           "Estado en Gemweb": d.get("Estado en Gemweb")})
    return salida


def main(argv=None):
    p = argparse.ArgumentParser(description="Busca en Gemweb las facturas que faltan por "
                                            "señalar en el cuadro de seguimiento.")
    p.add_argument("--aplicar", action="store_true",
                   help="aplica todos los cambios encontrados (si no, solo los propone)")
    a = p.parse_args(argv)
    config.CARPETA_DATOS.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(config.CARPETA_DATOS / "estados_gemweb.log",
                                      encoding="utf-8"), logging.StreamHandler()])
    try:
        resultado = buscar()
        guardar_propuesta(resultado)
        if a.aplicar:
            hechos, saltados = aplicar(resultado["cambios"], USUARIO)
            LOG.info("Aplicados %d cambios (%d no, porque la celda ya habia cambiado)",
                     len(hechos), len(saltados))
    except Exception:                                     # noqa: BLE001
        LOG.exception("Error al buscar los estados en Gemweb")
        return 1
    for c in resultado["cambios"]:
        LOG.info("%s %s %s: %s -> %s (factura %s)", c["cups"], c["denominacion"],
                 C.columna_mes(c["anio"], c["mes"]), c["antes"] or "vacío", c["despues"],
                 c["factura"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
