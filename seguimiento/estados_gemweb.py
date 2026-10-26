# -*- coding: utf-8 -*-
"""
Estados del cuadro de seguimiento a partir de las facturas de Gemweb
====================================================================
Se ejecuta a las 7:00 y a las 12:00 (tarea programada del servidor) y con el boton
«Actualizar ahora desde Gemweb» del cuadro:

    python -m seguimiento.estados_gemweb             (actualiza)
    python -m seguimiento.estados_gemweb --simular   (solo dice que cambiaria)

1. Una consulta con todas las facturas cuyo periodo termina desde el dia 1 de hace 3 meses
   (sin anuladas, simuladas, abonos, abonadas ni complementarias) y otra con los suministros,
   para pasar de su id al CUPS.
2. Cada factura cuenta en los meses que cubre en mas de la mitad de sus dias (una bimestral,
   en dos meses).
3. Regla, que nunca hace retroceder un estado:
     vacio o «Sin fra»  -> «Recibida» (factura sin verificar) o «Pdte inf» (verificada)
     «Recibida»         -> «Pdte inf» si esta verificada
   «Inf enviado», «Inf no enviado», «Temporal», «Cadencia» y cualquier otro no se tocan.
Solo afecta a los suministros activos «Sgto. y Control» del cuadro cuyo CUPS esta en Gemweb
(los de asesoramiento no llevan el seguimiento mensual de facturas).
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
EXCLUIR = ("anulada", "simulada", "abonament", "abonada", "complementaria")
USUARIO = "Gemweb automático"
LOG = logging.getLogger("estados_gemweb")


def _fecha(texto):
    try:
        return dt.datetime.strptime((texto or "").strip()[:10], "%d/%m/%Y").date()
    except ValueError:
        try:
            return dt.date.fromisoformat((texto or "").strip()[:10])
        except ValueError:
            return None


def desde(hoy):
    """Dia 1 de hace MESES_ATRAS meses."""
    total = hoy.year * 12 + hoy.month - 1 - MESES_ATRAS
    return dt.date(total // 12, total % 12 + 1, 1)


def facturas_gemweb(cliente, inicio):
    """[(cups de 20, numero, inicio, fin, verificada)] de Gemweb."""
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
        if any(str(d.get(m) or "0").strip() not in ("0", "") for m in EXCLUIR):
            continue
        ini, fin = _fecha(d.get("periode_ini")), _fecha(d.get("periode_fin"))
        if d.get("id_subministraments") not in cups or not ini or not fin:
            continue
        salida.append((cups[d["id_subministraments"]], d.get("num_factura"), ini, fin,
                       str(d.get("verificada") or "0").strip() == "1"))
    return salida


def nuevo_estado(actual, verificada):
    """Estado que corresponde con una factura (None si no cambia)."""
    if actual in (None, "", C.SIN_FRA):
        return C.PDTE_INF if verificada else C.RECIBIDA
    if actual == C.RECIBIDA and verificada:
        return C.PDTE_INF
    return None


def calcular_cambios(facturas, por_cups, estados, inicio):
    """Cambios a guardar: dicts id_suministro, cups, denominacion, anio, mes, antes, despues,
    factura y obs (la que ya tenia la celda)."""
    mejor = {}                                   # (id, anio, mes) -> (verificada, factura)
    for cups, numero, ini, fin, verificada in facturas:
        for anio, mes in D.meses_cubiertos(ini, fin):
            if dt.date(anio, mes, 1) < inicio:
                continue
            for s in por_cups.get(cups, []):
                clave = (s["id"], anio, mes)
                if clave not in mejor or (verificada and not mejor[clave][0]):
                    mejor[clave] = (verificada, numero, s)
    cambios = []
    for (id_, anio, mes), (verificada, numero, s) in sorted(mejor.items(),
                                                             key=lambda x: x[0]):
        previo = estados.get((id_, anio, mes)) or {}
        despues = nuevo_estado(previo.get("estado"), verificada)
        if despues:
            cambios.append({"id_suministro": id_, "cups": s["cups"],
                            "denominacion": s["datos"].get("Denominación suministro"),
                            "anio": anio, "mes": mes, "antes": previo.get("estado"),
                            "despues": despues, "factura": numero, "obs": previo.get("obs")})
    return cambios


def actualizar(simular=False, hoy=None, cliente=None):
    """Lee Gemweb y actualiza el cuadro. Devuelve la lista de cambios (hechos o, con
    simular, los que se harian)."""
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
    por_cups = {cups: [s for s in lista
                       if (s["datos"].get("Tipo contratación") or "").strip() == C.SGTO_Y_CONTROL]
                for cups, lista in D.suministros_por_cups().items()}
    cambios = calcular_cambios(facturas, por_cups, estados, inicio)
    LOG.info("%d facturas de Gemweb desde %s; %d cambios%s", len(facturas), inicio,
             len(cambios), " (simulación)" if simular else "")
    if simular or not cambios:
        return cambios
    almacen.guardar_sgc_estados(
        [{"id_suministro": c["id_suministro"], "anio": c["anio"], "mes": c["mes"],
          "estado": c["despues"], "obs": c["obs"], "origen": "Gemweb"} for c in cambios],
        USUARIO, origen="Gemweb")
    almacen.guardar_cambios_gemweb(cambios)
    almacen.registrar_actividad(USUARIO, "Seguimiento y Control", "Estados desde Gemweb",
                                "%d cambios" % len(cambios))
    return cambios


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
        cambios = actualizar(simular=a.simular)
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
