# -*- coding: utf-8 -*-
"""
COMPARAR EL MOTOR DEL ANALISIS DE OFERTAS CON LA PLANTILLA EXCEL
================================================================
    python -m herramientas.comparar_plantilla "Estudio Multipunto CLIENTE.xlsx"

Lee las entradas de un estudio hecho con la plantilla multipunto (hojas Cliente, Consumos,
Parámetros y las hojas de oferta de la lista del Resumen), construye el Estudio de la
aplicación, lo calcula y compara año a año y concepto a concepto con los importes que Excel
dejó guardados en cada hoja de oferta (G5:J14). El Excel tiene que estar guardado desde Excel
(con los valores calculados).

Solo lectura y solo en local: los estudios reales NO se suben a git.
"""

import datetime as dt
import sys
import warnings
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

from analisis_ofertas import calculo
from analisis_ofertas import modelo as mo

# filas y columnas de la plantilla (generar_plantilla.py)
TARIFAS = mo.TARIFAS
SISTEMAS = mo.SISTEMAS
CR0, NCUPS = 18, 80
TR0, ANIOS_PLANTILLA = 19, [2026, 2027, 2028, 2029, 2030, 2031]
LS0 = 182
KR0 = 99
AP_OM_EFF, AP_SS_EFF = 474, 598
G0, NSLOT, NCUST = 47, 12, 4
PT0, NPG = 68, 44
CI0 = 204
AO0 = 244
CU0 = 606
FILA_CFG = {"comercializadora": 5, "producto": 6, "modalidad": 7, "duracion": 8, "unidad": 9,
            "atr": 10, "potencia": 11, "descuento": 12, "base_b": 13, "pc": 15, "os": 16,
            "om": 17, "srad": 18, "desvios": 19, "fee": 20, "base_fee": 21, "k": 22,
            "k_ssaa": 23, "gdo": 24, "fnee": 25, "otros_fijos": 26, "ssaa_modo": 28,
            "banda_min": 29, "banda_max": 30, "techo": 31, "ssaa_perfil": 32,
            "ssaa_perdidas": 33, "ssaa_pct_perdidas": 34, "ssaa_f1": 35, "ssaa_f2": 36,
            "cierres_tipo": 38, "coste_click": 39}
RESULTADOS = {"consumo": 5, "potencia": 6, "energia": 7, "peajes": 8, "ssaa": 9, "otros": 10,
              "k": 11, "iee": 12, "servicios": 13, "total": 14}


def _n(v, defecto=0.0):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else defecto


def _fila(ws, r, c0, n=6):
    return [ws.cell(r, c0 + i).value for i in range(n)]


def _fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v


def aprow(r0, si, ci, m):
    return r0 + (si * 2 + ci) * 12 + m - 1


def regulados_de(pa) -> dict:
    """Datos regulados de cada año a partir de la hoja Parámetros (peajes+cargos juntos)."""
    generales = {"iee": _n(pa["C5"].value), "iee_minimo_general": _n(pa["C6"].value),
                 "iee_minimo_reducido": _n(pa["C7"].value),
                 "impuesto_municipal": _n(pa["C8"].value), "fnee": _n(pa["C9"].value),
                 "om": _n(pa["C10"].value), "os": _n(pa["C11"].value)}
    salida = {}
    for yi, anio in enumerate(ANIOS_PLANTILLA):
        tarifas = {}
        for ti, t in enumerate(TARIFAS):
            r = TR0 + ti * len(ANIOS_PLANTILLA) + yi
            n = mo.n_periodos(t)
            pen = LS0 + ti * len(SISTEMAS)
            tarifas[t] = {"peajes_potencia": [_n(v) for v in _fila(pa, r, 6)],
                          "cargos_potencia": [0.0] * 6,
                          "peajes_energia": [_n(v) for v in _fila(pa, r, 12)][:n],
                          "cargos_energia": [0.0] * n,
                          "pagos_capacidad": [_n(v) for v in _fila(pa, r, 18)][:n],
                          "srad": [_n(v) for v in _fila(pa, r, 24)][:n],
                          "perdidas": [_n(v) for v in _fila(pa, pen, 4)][:n]}
        perd_sis = {}
        for si, s in enumerate(SISTEMAS[1:], start=1):
            for ti, t in enumerate(TARIFAS):
                r = LS0 + ti * len(SISTEMAS) + si
                perd_sis.setdefault(s, {})[t] = [_n(v) for v in _fila(pa, r, 4)][:mo.n_periodos(t)]
        salida[anio] = {"generales": generales, "tarifas": tarifas, "perdidas_sistema": perd_sis,
                        "mercado": {"pct_omip": _n(pa["C12"].value), "anios_media_omie": 3}}
    return salida


def apuntamientos_de(pa, r0) -> dict:
    salida = {}
    for si, s in enumerate(SISTEMAS):
        for ci, cal in enumerate(["6P", "3P"]):
            salida[mo.clave_ap(s, cal)] = {
                m: [pa.cell(aprow(r0, si, ci, m), 3 + p).value for p in range(6)]
                for m in range(1, 13)}
    return salida


def oferta_de(ws, cups_por_n: dict) -> mo.Oferta:
    c = {k: ws.cell(r, 3).value for k, r in FILA_CFG.items()}
    o = mo.Oferta(nombre=ws.title, comercializadora=c["comercializadora"] or "",
                  producto=c["producto"] or "", modalidad=c["modalidad"] or mo.FIJO,
                  duracion=int(_n(c["duracion"], 12)), unidad=c["unidad"] or "€/kWh",
                  atr_incluido=c["atr"] == "Sí", potencia=c["potencia"] or mo.POT_REGULADA,
                  descuento=_n(c["descuento"]), base_b=c["base_b"] or mo.NETO,
                  pagos_capacidad=c["pc"] == "Sí", os=c["os"] == "Sí", om=c["om"] == "Sí",
                  srad=c["srad"] == "Sí", desvios=_n(c["desvios"]), fee=_n(c["fee"]),
                  base_fee=c["base_fee"] or mo.NETO, k=_n(c["k"]), k_ssaa=c["k_ssaa"] == "Sí",
                  gdo=_n(c["gdo"]), fnee=c["fnee"] or mo.FNEE_INCLUIDO,
                  otros_fijos=_n(c["otros_fijos"]), ssaa_modo=c["ssaa_modo"] or mo.SSAA_INCLUIDOS,
                  banda_min=_n(c["banda_min"]), banda_max=_n(c["banda_max"]),
                  techo=_n(c["techo"]), ssaa_perfil=c["ssaa_perfil"] == "Sí",
                  ssaa_perdidas=c["ssaa_perdidas"] or mo.PERD_NO,
                  ssaa_pct_perdidas=_n(c["ssaa_pct_perdidas"]), ssaa_f1=_n(c["ssaa_f1"], 1.0),
                  ssaa_f2=_n(c["ssaa_f2"], 1.0), cierres_tipo=c["cierres_tipo"] or mo.SIN_CIERRES,
                  coste_click=_n(c["coste_click"]))
    for r in range(G0 + NSLOT, G0 + NSLOT + NCUST):
        nombre, tar, sis = (ws.cell(r, c).value for c in (3, 4, 5))
        if nombre and tar:
            o.grupos_propios.append(mo.GrupoPropio(str(nombre), tar, sis or "Península"))
    for pg in range(NPG):
        nombre = ws.cell(PT0 + pg * 3, 2).value
        if not nombre:
            continue
        o.precios[str(nombre)] = [mo.PreciosAnio(
            energia=_fila(ws, PT0 + pg * 3 + y, 4), termino_b=_fila(ws, PT0 + pg * 3 + y, 10),
            ci=[_n(v, 1.0) for v in _fila(ws, PT0 + pg * 3 + y, 16)],
            potencia=_fila(ws, PT0 + pg * 3 + y, 22)) for y in range(3)]
    for k in range(36):
        r = CI0 + k
        pct, mw, precio, omip = (ws.cell(r, c).value for c in (6, 7, 8, 5))
        if pct is not None or mw is not None or (isinstance(precio, (int, float)) and precio != omip):
            o.cierres.append(mo.Cierre(k + 1, pct=pct, mw=mw,
                                       precio=precio if isinstance(precio, (int, float)) else None))
    for si, s in enumerate(SISTEMAS):
        for ci, cal in enumerate(["6P", "3P"]):
            propios = {m: _fila(ws, aprow(AO0, si, ci, m), 3) for m in range(1, 13)}
            propios = {m: v for m, v in propios.items() if any(x is not None for x in v)}
            if propios:
                o.apuntamientos_propios[mo.clave_ap(s, cal)] = propios
    for n in range(NCUPS):
        r = CU0 + n
        cups = cups_por_n.get(n + 1)
        if not cups:
            continue
        if ws.cell(r, 5).value != "Sí":
            o.excluidos.append(cups)
        if ws.cell(r, 6).value:
            o.asignacion[cups] = str(ws.cell(r, 6).value)
    return o


def estudio_de(ruta: Path):
    """(Estudio, {hoja: resultados guardados por Excel}, nombre de la hoja del actual)."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(ruta, data_only=True)
    cl, co, pa, rs = wb["Cliente"], wb["Consumos"], wb["Parámetros"], wb["Resumen"]
    e = mo.Estudio(cliente=cl["C7"].value or "", cif=cl["C8"].value or "",
                   inicio=_fecha(cl["C10"].value),
                   crecimiento=[_n(cl["C11"].value), _n(cl["C12"].value)])
    cups_por_n = {}
    for n in range(NCUPS):
        r = CR0 + n
        cups = cl.cell(r, 2).value
        if not cups:
            continue
        cups_por_n[n + 1] = cups
        e.suministros.append(mo.Suministro(
            cups=cups, alias=cl.cell(r, 3).value or "", tarifa=cl.cell(r, 5).value,
            sistema=cl.cell(r, 16).value or "Península",
            potencias=[_n(v) for v in _fila(cl, r, 6)], reduccion_iee=_n(cl.cell(r, 12).value),
            alquiler=_n(cl.cell(r, 13).value)))
    for r in range(6, 6 + NCUPS * 12):
        n, mes = co.cell(r, 1).value, co.cell(r, 4).value
        if n in cups_por_n and isinstance(mes, (int, float)):
            s = e.suministro(cups_por_n[n])
            s.consumo[int(mes)] = [_n(v) for v in _fila(co, r, 6)]
    e.regulados = regulados_de(pa)
    e.pct_omip = _n(pa["C12"].value)
    for k in range(36):
        r = KR0 + k
        fecha = _fecha(pa.cell(r, 3).value)
        if not fecha:
            continue
        clave = f"{fecha.year}-{fecha.month:02d}"
        e.mercado[clave] = {"omip": _n(pa.cell(r, 9).value, None),
                            "omie_media": _n(pa.cell(r, 10).value, None)}
        e.ssaa_manual[clave] = _n(pa.cell(r, 14).value)
    e.ap_omie = apuntamientos_de(pa, AP_OM_EFF)
    e.ap_ssaa = apuntamientos_de(pa, AP_SS_EFF)
    hojas = [rs.cell(r, 2).value for r in range(9, 30) if rs.cell(r, 2).value]
    guardados = {}
    for i, h in enumerate(hojas):
        if h not in wb.sheetnames:
            continue
        ws = wb[h]
        o = oferta_de(ws, cups_por_n)
        guardados[h] = {k: [ws.cell(r, 7 + y).value for y in range(4)]
                        for k, r in RESULTADOS.items()}
        if i == 0:
            o.es_actual = True
            e.actual = o
        else:
            e.ofertas.append(o)
    return e, guardados, hojas[0]


def comparar(ruta: Path, tolerancia: float = 0.01) -> list:
    """Diferencias (oferta, año, concepto, Excel, app) mayores que la tolerancia (€)."""
    e, guardados, actual = estudio_de(Path(ruta))
    ctx = calculo.Contexto(e)
    resultados = {o.nombre: calculo.calcular_oferta(e, o, ctx) for o in e.ofertas}
    resultados[actual] = calculo.calcular_oferta(e, e.actual, ctx)
    diferencias = []
    for hoja, res in resultados.items():
        anios = res.por_anio()
        for concepto, valores in guardados[hoja].items():
            for y in range(3):
                excel = valores[y]
                if not isinstance(excel, (int, float)):
                    continue
                if y < len(anios):
                    app = anios[y]["consumo"] / 1000 if concepto == "consumo" else anios[y][concepto]
                else:
                    app = 0.0
                if abs(app - excel) > tolerancia:
                    diferencias.append((hoja, y + 1, concepto, excel, app))
    return diferencias, resultados


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print(__doc__)
        return 2
    total = 0
    for ruta in argv:
        dif, resultados = comparar(Path(ruta))
        print(f"\n{Path(ruta).name}: {len(resultados)} hojas calculadas")
        for hoja, r in resultados.items():
            print(f"  {hoja:18} {r.oferta.comercializadora or '':22} {r.n_meses:3} meses "
                  f"{r.total:14,.2f} €  {r.eur_mwh:8.2f} €/MWh")
        if dif:
            print(f"  {len(dif)} DIFERENCIAS de más de 0,01 €:")
            for hoja, y, c, excel, app in dif[:60]:
                print(f"    {hoja:18} año {y} {c:10} Excel {excel:14,.4f}  app {app:14,.4f}  "
                      f"dif {app - excel:+12,.4f}")
        else:
            print("  Sin diferencias (tolerancia 0,01 €).")
        total += len(dif)
    return 1 if total else 0


if __name__ == "__main__":
    raise SystemExit(main())
