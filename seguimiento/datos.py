# -*- coding: utf-8 -*-
"""
Datos de Seguimiento y Control (sin Streamlit)
==============================================
- Importar el Excel combinado de cuadros de seguimiento (CLIENTES, BAJAS, CLAVES, CONTACTOS,
  SMAC, TELEMEDIDA y Lista) y el Excel de contactos de envio a la base de datos.
- Pasar las tablas de la base de datos a DataFrame para editarlas y guardar solo lo cambiado.
- Exportar a Excel con las columnas y colores del cuadro original.

Las columnas se reconocen por su titulo (columnas.py), no por su letra; los meses, por la
fecha de su cabecera. Lo que no encaja en ninguna columna se guarda en «Otros» con su titulo.
"""

import datetime as dt
import io
import math
import re

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import almacen
from seguimiento import columnas as C


# ------------------------------------------------------------------- valores
def _vacio(v):
    return v is None or (isinstance(v, float) and math.isnan(v)) or (
        isinstance(v, str) and not v.strip()) or v is pd.NaT


def _texto(v):
    if _vacio(v):
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    # solo los espacios de los extremos: los interiores se conservan porque los envios cruzan
    # contactos y facturas por texto exacto, como los programas originales
    t = str(v).strip()
    return t or None


def numero(v):
    """Numero de una celda; acepta texto con formato español (1.425.738 / 23,393). None si
    no lo es."""
    if _vacio(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("\xa0", "").replace(" ", "").replace("€", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1 or re.fullmatch(r"-?\d{1,3}\.\d{3}", t):
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


def fecha(v):
    """Fecha ISO (AAAA-MM-DD) de una celda; None si no lo es."""
    if _vacio(v):
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,5})", str(v).strip())
    if m:
        d, mes, a = (int(x) for x in m.groups())
        a = a + 2000 if a < 100 else a
        try:
            return dt.date(a, mes, d).isoformat()
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2}).*", str(v).strip())
    return "%s-%s-%s" % m.groups() if m else None


def convertir(columna, v):
    """(valor para guardar, texto que no se ha podido convertir o None)."""
    if columna["tipo"] == C.NUMERO:
        n = numero(v)
        return n, (None if n is not None or _vacio(v) or _texto(v) == "-" else _texto(v))
    if columna["tipo"] == C.FECHA:
        f = fecha(v)
        return f, (None if f is not None or _vacio(v) or _texto(v) == "-" else _texto(v))
    return _texto(v), None


def estado_canonico(v):
    """«inf enviado» -> «Inf enviado»; los que no son estados conocidos se dejan como estan."""
    t = _texto(v)
    if t is None:
        return None
    for valor, _, _ in C.ESTADOS:
        if C.normalizar(valor) == C.normalizar(t):
            return valor
    return t


def total_cuota(datos):
    valores = [numero(datos.get(c)) for c in C.CUOTAS_SUMADAS]
    valores = [v for v in valores if v is not None]
    return round(sum(valores), 2) if valores else None


# ------------------------------------------------------------------- lectura del Excel
def _mapear(cabecera, definiciones):
    """{indice de columna: definicion} reconociendo titulos: primero los exactos y despues
    sin acentos ni mayusculas; cada columna del Excel se asigna una sola vez."""
    asignadas, mapa = set(), {}
    for exacto in (True, False):
        for d in definiciones:
            if d["clave"] in mapa.values():
                continue
            for i, titulo in enumerate(cabecera):
                if i in asignadas or titulo is None or isinstance(titulo, dt.datetime):
                    continue
                t = str(titulo).strip()
                if any((t == e) if exacto else (C.normalizar(t) == C.normalizar(e))
                       for e in d["excel"]):
                    mapa[i] = d["clave"]
                    asignadas.add(i)
                    break
    return {i: next(d for d in definiciones if d["clave"] == k) for i, k in mapa.items()}


def _leer_hoja(ws, definiciones, meses=False):
    """Filas de una hoja: lista de (datos, {(anio, mes): (estado, obs)}) y avisos."""
    filas_excel = list(ws.iter_rows(values_only=True))
    if not filas_excel:
        return [], []
    cabecera = list(filas_excel[0])
    mapa = _mapear(cabecera, definiciones)
    columnas_mes = {}
    if meses:
        for i, t in enumerate(cabecera):
            if isinstance(t, dt.datetime):
                obs = i + 1 if i + 1 < len(cabecera) and C.normalizar(cabecera[i + 1]) == "obs" \
                    else None
                columnas_mes[i] = (t.year, t.month, obs)
    usadas = set(mapa) | set(columnas_mes) | {o for _, _, o in columnas_mes.values() if o}
    avisos, salida = [], []
    for n, fila in enumerate(filas_excel[1:], start=2):
        if all(_vacio(v) for v in fila):
            continue
        datos, otros = {}, []
        for i, d in mapa.items():
            v = fila[i] if i < len(fila) else None
            valor, sobra = convertir(d, v)
            if d.get("lista") and valor:
                valor = " ".join(valor.split())
            datos[d["clave"]] = valor
            if sobra:
                otros.append("%s: %s" % (d["clave"], sobra))
        for i, v in enumerate(fila):
            if i not in usadas and not _vacio(v):
                titulo = _texto(cabecera[i]) if i < len(cabecera) else None
                otros.append("%s: %s" % (titulo, _texto(v)) if titulo else _texto(v))
        if otros:
            previo = datos.get(C.OTROS)
            datos[C.OTROS] = " | ".join(([previo] if previo else []) + otros)
        estados = {}
        for i, (anio, mes, obs) in columnas_mes.items():
            est = estado_canonico(fila[i] if i < len(fila) else None)
            ob = _texto(fila[obs]) if obs is not None and obs < len(fila) else None
            if est or ob:
                estados[(anio, mes)] = (est, ob)
        if not any(datos.values()) and not estados:
            continue
        salida.append((datos, estados))
    return salida, avisos


def _leer_listas(ws):
    """Pestaña «Lista»: (listas {lista: [dicts]}, estados [dicts])."""
    listas, estados, actual = {}, [], None
    for fila in ws.iter_rows(values_only=True):
        a, b = (fila + (None, None, None, None))[:2]
        c, d = (fila + (None, None, None, None))[2:4]
        if _vacio(a) and not _vacio(b) and actual is None:
            estados.append({"valor": estado_canonico(b), "descripcion": _texto(c),
                            "color": _texto(d)})
        elif not _vacio(a) and C.normalizar(a) in {C.normalizar(k) for k in C.LISTAS_EXCEL}:
            actual = next(v for k, v in C.LISTAS_EXCEL.items()
                          if C.normalizar(k) == C.normalizar(a))
            listas.setdefault(actual, [])
        elif actual and not _vacio(a):
            listas[actual].append({"valor": " ".join(_texto(a).split()),
                                   "descripcion": _texto(b)})
    return listas, estados


def leer_cuadro(contenido):
    """Lee el Excel combinado (bytes). Devuelve dict con suministros, estados, filas (tablas
    auxiliares), listas, equipo y recuento por pestaña."""
    wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
    hojas = {C.normalizar(n): wb[n] for n in wb.sheetnames}
    if "clientes" not in hojas:
        raise ValueError("El Excel no tiene la pestaña CLIENTES.")
    suministros, estados, recuento = [], [], {}
    for nombre, baja in (("clientes", False), ("bajas", True)):
        if nombre not in hojas:
            continue
        filas, _ = _leer_hoja(hojas[nombre], C.CUADRO, meses=True)
        recuento[nombre.upper()] = len(filas)
        for datos, est in filas:
            clave = len(suministros)
            if baja and C.normalizar(datos.get("Baja en GEYPE")) != "si":
                datos["Baja en GEYPE"] = "SI"
            suministros.append({"clave": clave, "datos": datos, "baja": baja})
            estados += [{"clave": clave, "anio": a, "mes": m, "estado": e, "obs": o}
                        for (a, m), (e, o) in est.items()]
    filas_aux = {}
    for tabla, t in C.TABLAS.items():
        hoja = hojas.get(C.normalizar(t.get("hoja", "")))
        if hoja is None:
            continue
        filas, _ = _leer_hoja(hoja, t["columnas"])
        filas_aux[tabla] = [d for d, _ in filas]
        recuento[t["hoja"]] = len(filas)
    listas, estados_lista = _leer_listas(hojas["lista"]) if "lista" in hojas else ({}, [])
    listas["estados"] = estados_lista or [{"valor": v, "descripcion": d, "color": c}
                                          for v, d, c in C.ESTADOS]
    return {"suministros": suministros, "estados": estados, "filas": filas_aux,
            "listas": listas, "recuento": recuento}


def leer_contactos_envio(contenido):
    """Excel de contactos de envio (bytes): lista de datos de la primera hoja."""
    wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
    filas, _ = _leer_hoja(wb.worksheets[0], C.TABLAS["contactos_envio"]["columnas"])
    if filas and not any(d.get("Para") for d, _ in filas):
        raise ValueError("El Excel no tiene la columna «Para».")
    return [d for d, _ in filas]


def _equipo(leido, contactos):
    """Tecnicos y auxiliares: los de la pestaña «Lista» y los que aparecen en las tablas."""
    tecnicos = [v["valor"] for v in leido["listas"].pop("tecnicos", [])]
    auxiliares = [v["valor"] for v in leido["listas"].pop("auxiliares", [])]
    filas = [s["datos"] for s in leido["suministros"]] + list(contactos or [])
    for lista in leido["filas"].values():
        filas += lista
    no_nombres = {"tecnico", "auxiliar", "-"}       # cabeceras repetidas y guiones
    for f in filas:
        if f.get("Técnico") and C.normalizar(f["Técnico"]) not in no_nombres:
            tecnicos.append(f["Técnico"])
        if f.get("Auxiliar") and C.normalizar(f["Auxiliar"]) not in no_nombres:
            auxiliares.append(f["Auxiliar"])
    nombres = sorted(set(tecnicos) | set(auxiliares))
    return [{"nombre": n, "tecnico": n in tecnicos, "auxiliar": n in auxiliares}
            for n in nombres]


def _completar_listas(leido):
    """Añade a cada lista los valores que ya se usan en el cuadro y no estaban en ella, para
    que los desplegables los muestren."""
    listas = leido["listas"]
    for c in C.CUADRO:
        nombre = c["lista"]
        if not nombre or nombre in ("tecnicos", "auxiliares"):
            continue
        actuales = listas.setdefault(nombre, [])
        conocidos = {v["valor"] for v in actuales}
        for s in leido["suministros"]:
            v = s["datos"].get(c["clave"])
            if v and v not in conocidos:
                actuales.append({"valor": v, "descripcion": None})
                conocidos.add(v)
    estados = listas["estados"]
    conocidos = {v["valor"] for v in estados}
    for e in leido["estados"]:
        if e["estado"] and e["estado"] not in conocidos:
            estados.append({"valor": e["estado"], "descripcion": None, "color": None})
            conocidos.add(e["estado"])


def importar(leido, contactos, usuario):
    """Sustituye el contenido de Seguimiento y Control por lo leido. contactos: lista de
    contactos de envio o None para conservar los que ya hay."""
    if contactos is None:
        contactos = [f["datos"] for f in almacen.listar_sgc_filas("contactos_envio")]
    filas = dict(leido["filas"])
    filas["contactos_envio"] = contactos
    equipo = _equipo(leido, contactos)
    _completar_listas(leido)
    resultado = almacen.reemplazar_sgc(leido["suministros"], leido["estados"], filas,
                                       leido["listas"], equipo, usuario)
    # los suministros cambian de id: los cambios de Gemweb ya buscados no valen
    from seguimiento import estados_gemweb
    almacen.guardar_ajuste_sgc(estados_gemweb.PROPUESTA, "")
    return resultado


# ------------------------------------------------------------------- periodos
def meses_cubiertos(inicio, fin):
    """Meses (anio, mes) de un periodo de factura o informe: los que cubre en mas de la mitad
    de sus dias. Si ninguno llega (periodo corto que cae entre dos meses), el de mas dias."""
    if not inicio or not fin or fin < inicio:
        return []
    dias = {}
    d = inicio
    while d <= fin:
        dias[(d.year, d.month)] = dias.get((d.year, d.month), 0) + 1
        d += dt.timedelta(days=1)
    def del_mes(a, m):
        siguiente = dt.date(a + (m == 12), m % 12 + 1, 1)
        return (siguiente - dt.date(a, m, 1)).days
    meses = [k for k, n in dias.items() if n * 2 > del_mes(*k)]
    return sorted(meses) or [max(dias, key=lambda k: (dias[k], k))]


def suministros_por_cups():
    """{CUPS de 20 caracteres: [suministro activo]} del cuadro."""
    salida = {}
    for s in almacen.listar_sgc_suministros(bajas=False):
        cups = (s["cups"] or "").strip().upper()[:20]
        if cups:
            salida.setdefault(cups, []).append(s)
    return salida


def nombres_por_cups():
    """{CUPS de 20 caracteres: {"denominacion", "titular"}} del cuadro, para mostrar el nombre de
    un suministro en otros modulos. Los activos mandan sobre las bajas."""
    salida = {}
    for s in sorted(almacen.listar_sgc_suministros(), key=lambda x: bool(x["baja"])):
        cups = (s["cups"] or "").strip().upper()[:20]
        if cups and cups not in salida:
            salida[cups] = {"denominacion": s["datos"].get("Denominación suministro"),
                            "titular": s["datos"].get("Titular de contrato")}
    return salida


PASAN_A_ENVIADO = (C.RECIBIDA, C.PDTE_INF, C.INF_NO_ENVIADO)


def marcar_enviados(meses, usuario):
    """Pone «Inf enviado» en (cups, anio, mes) de los suministros activos del cuadro cuyo estado
    es «Recibida», «Pdte inf» o «Inf no enviado». Devuelve (marcados, sin marcar): listas de
    dicts cups, denominacion, mes y estado (el que tenia)."""
    por_cups = suministros_por_cups()
    anios = {a for _, a, _ in meses}
    estados = {}
    for a in anios:
        estados.update(almacen.listar_sgc_estados(a))
    cambios, marcados, sin_marcar = [], [], []
    for cups, anio, mes in dict.fromkeys(meses):
        for s in por_cups.get(cups, []):
            previo = estados.get((s["id"], anio, mes)) or {}
            fila = {"cups": cups, "denominacion": s["datos"].get("Denominación suministro"),
                    "mes": C.columna_mes(anio, mes), "estado": previo.get("estado")}
            if previo.get("estado") in PASAN_A_ENVIADO:
                cambios.append({"id_suministro": s["id"], "anio": anio, "mes": mes,
                                "estado": C.INF_ENVIADO, "obs": previo.get("obs")})
                marcados.append(fila)
            elif previo.get("estado") != C.INF_ENVIADO:
                sin_marcar.append(fila)
        if not por_cups.get(cups):
            sin_marcar.append({"cups": cups, "denominacion": None,
                               "mes": C.columna_mes(anio, mes), "estado": "no está en el cuadro"})
    almacen.guardar_sgc_estados(cambios, usuario, origen="envío")
    return marcados, sin_marcar


# ------------------------------------------------------------------- listas y equipo
def opciones(listas, equipo, nombre):
    """Valores de un desplegable."""
    if nombre == "tecnicos":
        return [e["nombre"] for e in equipo if e["tecnico"]]
    if nombre == "auxiliares":
        return [e["nombre"] for e in equipo if e["auxiliar"]]
    return [v["valor"] for v in listas.get(nombre, [])]


def colores_estados(listas):
    colores = dict(C.COLOR_ESTADO)
    for v in listas.get("estados", []):
        if v.get("color"):
            colores[v["valor"]] = v["color"]
    return colores


# ------------------------------------------------------------------- cuadro <-> DataFrame
ID = "_id"


def _valor_df(columna, v):
    if columna["tipo"] == C.FECHA:
        return dt.date.fromisoformat(v) if v else None
    if columna["tipo"] == C.NUMERO:
        return numero(v)
    return v


def columnas_cuadro(anio):
    """Claves de las columnas del cuadro en orden, con los meses del año en su sitio."""
    claves = [c["clave"] for c in C.CUADRO]
    meses = []
    for m in range(1, 13):
        meses += [C.columna_mes(anio, m), C.columna_obs(anio, m)]
    return claves[:C.POSICION_MESES] + meses + claves[C.POSICION_MESES:]


def cuadro_df(anio, bajas=False):
    """DataFrame del cuadro de un año (una fila por suministro) con la columna _id."""
    suministros = almacen.listar_sgc_suministros(bajas=bajas)
    estados = almacen.listar_sgc_estados(anio)
    filas = []
    for s in suministros:
        f = {ID: s["id"]}
        for c in C.CUADRO:
            f[c["clave"]] = (total_cuota(s["datos"]) if c["clave"] == C.TOTAL_CUOTA
                             else _valor_df(c, s["datos"].get(c["clave"])))
        for m in range(1, 13):
            e = estados.get((s["id"], anio, m), {})
            f[C.columna_mes(anio, m)] = e.get("estado")
            f[C.columna_obs(anio, m)] = e.get("obs")
        filas.append(f)
    df = pd.DataFrame(filas, columns=[ID] + columnas_cuadro(anio))
    for c in C.CUADRO:
        if c["tipo"] in (C.NUMERO, C.CALCULADA):
            df[c["clave"]] = pd.to_numeric(df[c["clave"]], errors="coerce")
        else:
            df[c["clave"]] = df[c["clave"]].astype(object)
    return df


def _guardable(columna, v):
    if _vacio(v):
        return None
    if columna["tipo"] == C.FECHA:
        return fecha(v)
    if columna["tipo"] == C.NUMERO:
        return numero(v)
    return _texto(v)


def cambios_cuadro(original, editado, anio):
    """Compara dos DataFrames del cuadro (el leido y el editado). Devuelve (suministros
    cambiados o nuevos: dicts id, datos; estados cambiados: dicts para guardar_sgc_estados;
    ids de las filas borradas)."""
    datos_previos = {s["id"]: s["datos"] for s in almacen.listar_sgc_suministros()}
    orig = {int(r[ID]): r for r in original.to_dict("records") if not _vacio(r.get(ID))}
    suministros, estados, vistos = [], [], set()
    definiciones = [c for c in C.CUADRO if c["tipo"] != C.CALCULADA]
    for nueva, r in enumerate(editado.to_dict("records")):
        id_ = None if _vacio(r.get(ID)) else int(r[ID])
        vistos.add(id_)
        previa = orig.get(id_, {})
        datos = dict(datos_previos.get(id_, {}))
        cambio = id_ is None
        for c in definiciones:
            v = _guardable(c, r.get(c["clave"]))
            if v != _guardable(c, previa.get(c["clave"])) or id_ is None:
                if v != datos.get(c["clave"]):
                    cambio = True
                datos[c["clave"]] = v
        if id_ is None and not any(datos.values()):
            continue                                   # fila nueva vacia
        if cambio:
            suministros.append({"id": id_, "datos": datos, "_nueva": nueva})
        for m in range(1, 13):
            col_e, col_o = C.columna_mes(anio, m), C.columna_obs(anio, m)
            e, o = _texto(r.get(col_e)), _texto(r.get(col_o))
            if id_ is None or e != _texto(previa.get(col_e)) or o != _texto(previa.get(col_o)):
                if id_ is not None or e or o:
                    estados.append({"id_suministro": id_, "anio": anio, "mes": m,
                                    "estado": e, "obs": o, "_nueva": nueva})
    borrados = [i for i in orig if i not in vistos]
    return suministros, estados, borrados


def guardar_cuadro(original, editado, anio, usuario):
    """Guarda lo cambiado en el editor. Las filas borradas pasan a baja con la fecha de hoy.
    Devuelve (nº de suministros cambiados, nº de celdas de meses, nº de bajas)."""
    suministros, estados, borrados = cambios_cuadro(original, editado, anio)
    ids = almacen.guardar_sgc_suministros(
        [{"id": s["id"], "datos": s["datos"]} for s in suministros], usuario)
    nuevo_id = {s["_nueva"]: i for s, i in zip(suministros, ids) if s["id"] is None}
    for e in estados:
        if e["id_suministro"] is None:
            e["id_suministro"] = nuevo_id.get(e["_nueva"])
    estados = [{k: v for k, v in e.items() if k != "_nueva"} for e in estados
               if e["id_suministro"] is not None]
    almacen.guardar_sgc_estados(estados, usuario)
    dar_de_baja(borrados, usuario)
    return len(suministros), len(estados), len(borrados)


def dar_de_baja(ids, usuario, baja=True):
    """Pasa suministros a baja (con la fecha de hoy y «Baja en GEYPE» = SI) o los reactiva."""
    if not ids:
        return
    previos = {s["id"]: s["datos"] for s in almacen.listar_sgc_suministros()}
    cambios = []
    for i in ids:
        datos = dict(previos.get(i, {}))
        if baja:
            datos["Fecha de baja"] = datos.get("Fecha de baja") or dt.date.today().isoformat()
            datos["Baja en GEYPE"] = "SI"
        else:
            datos["Fecha de baja"] = None
            datos["Baja en GEYPE"] = "NO"
        cambios.append({"id": i, "datos": datos, "baja": baja})
    almacen.guardar_sgc_suministros(cambios, usuario)


# ------------------------------------------------------------------- tablas auxiliares
def tabla_df(tabla):
    definiciones = C.TABLAS[tabla]["columnas"]
    filas = []
    for f in almacen.listar_sgc_filas(tabla):
        filas.append({ID: f["id"], **{c["clave"]: _valor_df(c, f["datos"].get(c["clave"]))
                                      for c in definiciones}})
    df = pd.DataFrame(filas, columns=[ID] + [c["clave"] for c in definiciones])
    for c in definiciones:
        df[c["clave"]] = df[c["clave"]].astype(object)
    return df


def guardar_tabla(tabla, original, editado, usuario):
    """Guarda lo cambiado de una tabla auxiliar. Devuelve (cambiadas o nuevas, borradas)."""
    definiciones = C.TABLAS[tabla]["columnas"]
    previos = {f["id"]: f["datos"] for f in almacen.listar_sgc_filas(tabla)}
    orig_ids = {int(i) for i in original[ID].dropna()}
    cambios, vistos = [], set()
    for r in editado.to_dict("records"):
        id_ = None if _vacio(r.get(ID)) else int(r[ID])
        vistos.add(id_)
        datos = dict(previos.get(id_, {}))
        nuevos = {c["clave"]: _guardable(c, r.get(c["clave"])) for c in definiciones}
        if id_ is None and not any(nuevos.values()):
            continue
        if id_ is None or any(datos.get(k) != v for k, v in nuevos.items()):
            datos.update(nuevos)
            cambios.append({"id": id_, "datos": datos})
    borrados = [i for i in orig_ids if i not in vistos]
    almacen.guardar_sgc_filas(tabla, cambios, borrados, usuario)
    return len(cambios), len(borrados)


# ------------------------------------------------------------------- exportar
_CABECERA = PatternFill("solid", fgColor="970000")


def _hoja_df(ws, df, colores=None, columnas_color=()):
    """Escribe un DataFrame en una hoja con cabecera granate, filtros y paneles fijos."""
    columnas = [c for c in df.columns if c != ID]
    ws.append(columnas)
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF", name="Arial")
        celda.fill = _CABECERA
        celda.alignment = Alignment(wrap_text=True, vertical="center")
    for r in df[columnas].itertuples(index=False):
        ws.append([None if _vacio(v) else v for v in r])
    for j, c in enumerate(columnas, start=1):
        ancho = max([len(str(c))] + [len(str(v)) for v in df[c].head(300) if not _vacio(v)])
        ws.column_dimensions[get_column_letter(j)].width = min(max(ancho + 2, 8), 45)
        muestra = df[c].dropna()
        if len(muestra) and isinstance(muestra.iloc[0], dt.date):
            for x in next(ws.iter_cols(min_col=j, max_col=j, min_row=2), ()):
                x.number_format = "dd/mm/yyyy"
    if colores:
        for j, c in enumerate(columnas, start=1):
            if c not in columnas_color:
                continue
            for x in next(ws.iter_cols(min_col=j, max_col=j, min_row=2)):
                color = colores.get(x.value)
                if color:
                    x.fill = PatternFill("solid", fgColor=color.lstrip("#").upper())
                    if color.upper() in ("#FF0000", "#7030A0"):
                        x.font = Font(color="FFFFFF", name="Arial")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


def excel_tablas(hojas, colores=None, anio=None):
    """Excel (bytes) con una hoja por DataFrame: hojas = [(nombre, df)]."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    columnas_color = [C.columna_mes(anio, m) for m in range(1, 13)] if anio else []
    for nombre, df in hojas:
        _hoja_df(wb.create_sheet(nombre[:31]), df, colores, columnas_color)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


def excel_completo(anio):
    """Todo Seguimiento y Control en un Excel, con las pestañas del cuadro original."""
    listas = almacen.listar_sgc_listas()
    hojas = [("CLIENTES", cuadro_df(anio)), ("BAJAS", cuadro_df(anio, bajas=True))]
    hojas += [(t.get("hoja") or t["titulo"].upper(), tabla_df(k)) for k, t in C.TABLAS.items()
              if k != "contactos_envio"]
    hojas.append(("CONTACTOS ENVÍO", tabla_df("contactos_envio")))
    return excel_tablas(hojas, colores_estados(listas), anio)
