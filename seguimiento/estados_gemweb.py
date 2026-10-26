# -*- coding: utf-8 -*-
"""
Estados del cuadro de seguimiento a partir de las facturas de Gemweb
====================================================================
Se ejecuta a las 7:00 y a las 12:00 (tarea programada del servidor) y con el boton
«Actualizar ahora desde Gemweb» del cuadro:

    python -m seguimiento.estados_gemweb             (actualiza)
    python -m seguimiento.estados_gemweb --simular   (solo dice que cambiaria)

1. Una consulta con todas las facturas cuyo periodo termina desde el 1 de enero del año en
   curso (en enero y febrero, desde el dia 1 de hace 3 meses, para no dejar fuera las de
   final del año anterior) y otra con los suministros, para pasar de su id al CUPS. Las
   anuladas, simuladas, abonos, abonadas y complementarias se listan pero no cuentan.
2. Cada factura cuenta en los meses que cubre en mas de la mitad de sus dias (una bimestral,
   en dos meses).
3. Regla, que nunca hace retroceder un estado:
     vacio o «Sin fra»  -> «Recibida» (factura sin verificar) o «Pdte inf» (verificada)
     «Recibida»         -> «Pdte inf» si esta verificada
   «Inf enviado», «Inf no enviado», «Temporal», «Cadencia» y cualquier otro no se tocan.
Solo afecta a los suministros activos «Sgto. y Control» del cuadro cuyo CUPS esta en Gemweb
(los de asesoramiento no llevan el seguimiento mensual de facturas).

Ademas del cambio, devuelve un listado de todo lo encontrado (una fila por factura y mes) con
lo que ha pasado en cada caso, para verlo y descargarlo desde el cuadro.
"""

import argparse
import datetime as dt
import logging
import sys
from pathlib import Path

if __package__ in (None, ""):                         # python seguimiento\estados_gemweb.py
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import almacen
import config
from seguimiento import columnas as C
from seguimiento import datos as D

MESES_ATRAS = 3
EXCLUIR = {"anulada": "anulada", "simulada": "simulada", "abonament": "abono",
           "abonada": "abonada", "complementaria": "complementaria"}
USUARIO = "Gemweb automático"
LOG = logging.getLogger("estados_gemweb")

# resultado de cada factura y mes en el listado
CAMBIA, SIN_CAMBIOS, OTRA_FACTURA = "Cambia", "Sin cambios", "Sin cambios (cuenta otra factura)"
NO_EN_CUADRO, NO_SYC, DESCARTADA = ("CUPS no está en el cuadro", "No es «Sgto. y Control»",
                                   "Descartada")
COLUMNAS_LISTADO = ["CUPS", "Denominación suministro", "Factura", "Comercializadora", "Emisión",
                    "Periodo inicio", "Periodo fin", "Verificada", "Mes", "Estado anterior",
                    "Estado nuevo", "Resultado"]


def _fecha(texto):
    try:
        return dt.datetime.strptime((texto or "").strip()[:10], "%d/%m/%Y").date()
    except ValueError:
        try:
            return dt.date.fromisoformat((texto or "").strip()[:10])
        except ValueError:
            return None


def desde(hoy):
    """1 de enero del año en curso, o el dia 1 de hace MESES_ATRAS meses si es anterior."""
    total = hoy.year * 12 + hoy.month - 1 - MESES_ATRAS
    return min(dt.date(hoy.year, 1, 1), dt.date(total // 12, total % 12 + 1, 1))


def facturas_gemweb(cliente, inicio):
    """Facturas de Gemweb cuyo periodo termina desde `inicio`: dicts cups (20 caracteres),
    factura, comercializadora, emision, inicio, fin, verificada y descartada (motivo o None)."""
    cups = {}
    for s in cliente.get_inventory(category="subministraments", limit=100000):
        d = {h.tag: h.text for h in s}
        if d.get("cups"):
            cups[d.get("id")] = d["cups"].strip().upper()[:20]
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
    despues, factura y obs (la que ya tenia la celda). listado: una fila por factura, suministro
    y mes (una por factura si no llega a ningun suministro) con lo que ha pasado."""
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
        previo = estados.get((id_, anio, mes)) or {}
        despues = nuevo_estado(previo.get("estado"), f["verificada"])
        if despues:
            s = next(x for x in por_cups[f["cups"]] if x["id"] == id_)
            cambios.append({"id_suministro": id_, "cups": f["cups"],
                            "denominacion": s["datos"].get("Denominación suministro"),
                            "anio": anio, "mes": mes, "antes": previo.get("estado"),
                            "despues": despues, "factura": f["factura"],
                            "obs": previo.get("obs")})
            nuevos[(id_, anio, mes)] = despues

    listado = []
    for f in sorted(facturas, key=lambda x: (x["cups"], x["inicio"], str(x["factura"]))):
        base = {"CUPS": f["cups"], "Factura": f["factura"],
                "Comercializadora": f["comercializadora"], "Emisión": f["emision"],
                "Periodo inicio": f["inicio"], "Periodo fin": f["fin"],
                "Verificada": "Sí" if f["verificada"] else "No"}
        suministros = por_cups.get(f["cups"], [])
        if f["descartada"]:
            listado.append({**base, "Resultado": "%s (%s)" % (DESCARTADA, f["descartada"])})
            continue
        if not suministros:
            listado.append({**base, "Resultado": NO_EN_CUADRO})
            continue
        meses = [(a, m) for a, m in D.meses_cubiertos(f["inicio"], f["fin"])
                 if dt.date(a, m, 1) >= inicio]
        for s in suministros:
            for anio, mes in meses:
                clave = (s["id"], anio, mes)
                fila = {**base, "Denominación suministro": s["datos"].get(
                    "Denominación suministro"), "Mes": C.columna_mes(anio, mes),
                    "Estado anterior": (estados.get(clave) or {}).get("estado")}
                if not _es_syc(s):
                    fila["Resultado"] = NO_SYC
                elif mejor.get(clave) is not f:
                    fila["Resultado"] = OTRA_FACTURA
                elif clave in nuevos:
                    fila["Resultado"], fila["Estado nuevo"] = CAMBIA, nuevos[clave]
                else:
                    fila["Resultado"] = SIN_CAMBIOS
                listado.append(fila)
    return cambios, listado


def actualizar(simular=False, hoy=None, cliente=None):
    """Lee Gemweb y actualiza el cuadro. Devuelve dict con inicio (fecha desde la que se
    buscan facturas), facturas (las encontradas), cambios (hechos o, con simular, los que se
    harian) y listado (una fila por factura y mes con su resultado)."""
    hoy = hoy or dt.date.today()
    if cliente is None:
        import gemweb
        credenciales, _ = gemweb.cargar_credenciales()
        if not credenciales:
            raise RuntimeError("No hay credenciales de Gemweb.")
        cliente = gemweb.obtener_api(*credenciales)
    inicio = desde(hoy)
    facturas = facturas_gemweb(cliente, inicio)
    estados = {}
    for anio in range(inicio.year, hoy.year + 1):
        estados.update(almacen.listar_sgc_estados(anio))
    cambios, listado = revisar(facturas, D.suministros_por_cups(), estados, inicio)
    LOG.info("%d facturas de Gemweb desde %s; %d cambios%s", len(facturas), inicio,
             len(cambios), " (simulación)" if simular else "")
    resultado = {"inicio": inicio, "facturas": facturas, "cambios": cambios,
                 "listado": listado}
    if simular or not cambios:
        return resultado
    almacen.guardar_sgc_estados(
        [{"id_suministro": c["id_suministro"], "anio": c["anio"], "mes": c["mes"],
          "estado": c["despues"], "obs": c["obs"], "origen": "Gemweb"} for c in cambios],
        USUARIO, origen="Gemweb")
    almacen.guardar_cambios_gemweb(cambios)
    almacen.registrar_actividad(USUARIO, "Seguimiento y Control", "Estados desde Gemweb",
                                "%d cambios" % len(cambios))
    return resultado


def main(argv=None):
    p = argparse.ArgumentParser(description="Actualiza los estados del cuadro de seguimiento "
                                            "con las facturas de Gemweb.")
    p.add_argument("--simular", action="store_true", help="solo muestra lo que cambiaria")
    a = p.parse_args(argv)
    config.CARPETA_DATOS.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(config.CARPETA_DATOS / "estados_gemweb.log",
                                      encoding="utf-8"), logging.StreamHandler()])
    try:
        cambios = actualizar(simular=a.simular)["cambios"]
    except Exception:                                     # noqa: BLE001
        LOG.exception("Error al actualizar los estados desde Gemweb")
        return 1
    for c in cambios:
        LOG.info("%s %s %s: %s -> %s (factura %s)", c["cups"], c["denominacion"],
                 C.columna_mes(c["anio"], c["mes"]), c["antes"] or "vacío", c["despues"],
                 c["factura"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
