# -*- coding: utf-8 -*-
"""
MI ESPACIO: DATOS (sin Streamlit)
=================================
Todo sale de lo que ya guarda la aplicación (no se pide nada nuevo a Gemweb):

    Fuente                         datos de una vez: suministros de Seguimiento y Control,
                                   estados del año anterior y actual, equipo, solicitudes,
                                   análisis, catálogo, contactos de envío, envíos y la propuesta
                                   de cambios de Gemweb (cargar() los lee de la base de datos)
    umbrales(usuario)              días del semáforo: los generales (Administración) y, por
                                   encima, los propios del usuario
    mis_nombres(usuario)           nombres de técnico y auxiliar del usuario (tabla del equipo)
    licitaciones(f, hoy, ...)      una fila por suministro con su fecha clave, días y semáforo
    agrupar_licitaciones(df)       una fila por licitación (cliente + fecha clave) con su estado
    carga_por_mes(df)              licitaciones por mes (y por técnico) para las gráficas
    informes_pendientes(f, hoy)    bloques de informes por vencer o con RETRASO y facturas
                                   pendientes, según la periodicidad (regla de retraso.py)
    eventos(f, desde, hasta)       calendario de vencimientos
    buscar(f, texto) / ficha(...)  todo lo de un cliente o CUPS
    BLOQUES, config_panel(...)     bloques del panel y la configuración de cada usuario
    datos_bloque(...)              filas de cada bloque
    avisos(...)                    contadores de lo urgente
"""

import datetime as dt
import re
from dataclasses import dataclass, field

import pandas as pd

from seguimiento import columnas as C
from seguimiento import retraso as R

# ------------------------------------------------------------------- semáforo y umbrales
ROJO, AMBAR, VERDE, GRIS, RETRASO = "Rojo", "Ámbar", "Verde", "Gris", "RETRASO"
COLOR = {ROJO: "#C00000", AMBAR: "#ED7D31", VERDE: "#70AD47", GRIS: "#A5A5A5",
         RETRASO: "#7B1E3A"}
ICONO = {ROJO: "🔴", AMBAR: "🟠", VERDE: "🟢", GRIS: "⚪", RETRASO: "⛔"}

DEFECTOS = {
    "rojo": 30, "ambar": 90,                 # licitaciones: días hasta la fecha clave
    "inf_rojo": 15, "inf_ambar": 30,         # informes y facturas: días hasta la fecha de control
    "margen": 3,                             # meses antes del fin de contrato si no hay otra fecha
    "catalogo": 15,                          # días antes de que caduque una oferta de catálogo
    "recepcion": 3,                          # días antes del fin de recepción de una solicitud
}
NOMBRES_UMBRAL = {
    "rojo": "Licitaciones: rojo si faltan menos de (días)",
    "ambar": "Licitaciones: ámbar si faltan hasta (días)",
    "inf_rojo": "Informes y facturas: rojo si faltan para el control hasta (días)",
    "inf_ambar": "Informes y facturas: ámbar si faltan para el control hasta (días)",
    "margen": "Fecha clave sin lanzamiento ni denuncia: fin de contrato menos (meses)",
    "catalogo": "Catálogo: avisar si caduca en (días)",
    "recepcion": "Solicitudes: avisar si la recepción acaba en (días)",
}

# columnas del cuadro
TECNICO, AUXILIAR, CLIENTE, CUPS = "Técnico", "Auxiliar", "Cliente/Grupo", "CUPS"
DENOMINACION, TITULAR, CIF = "Denominación suministro", "Titular de contrato", "CIF"
F_DENUNCIA, F_FIN = "Fecha denuncia contrato", "Fecha fin contrato"
F_LANZAMIENTO = "Fecha prevista lanzamiento licitación"
TIPOS_FECHA = ["Automática", "Lanzamiento", "Denuncia", "Fin de contrato"]


def umbrales(generales: dict = None, propios: dict = None) -> dict:
    """Los de DEFECTOS, sustituidos por los generales (Administración) y por los propios del
    usuario (los vacíos no cuentan)."""
    u = dict(DEFECTOS)
    for capa in (generales or {}, propios or {}):
        for k, v in capa.items():
            if k in u and v not in (None, ""):
                u[k] = int(v)
    return u


def semaforo(dias: int, u: dict, hay_trabajo: bool = False) -> str:
    """Licitaciones: rojo < u["rojo"] días, ámbar hasta u["ambar"], verde el resto; gris si la
    fecha ya ha pasado y no hay solicitud ni análisis."""
    if dias is None:
        return GRIS
    if dias < 0:
        return ROJO if hay_trabajo else GRIS
    if dias < u["rojo"]:
        return ROJO
    if dias <= u["ambar"]:
        return AMBAR
    return VERDE


def semaforo_control(dias: int, u: dict) -> str:
    """Informes y facturas: según los días hasta la fecha de control del bloque."""
    if dias < 0:
        return RETRASO
    if dias <= u["inf_rojo"]:
        return ROJO
    if dias <= u["inf_ambar"]:
        return AMBAR
    return VERDE


# ------------------------------------------------------------------- utilidades
def fecha(v):
    """date de un ISO (AAAA-MM-DD…), de dd/mm/aaaa [HH:MM] o de un date/datetime; None si no."""
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    t = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", t)
    try:
        if m:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        return dt.date.fromisoformat(t[:10])
    except ValueError:
        return None


def cups20(c) -> str:
    return (c or "").strip().upper()[:20]


def lista_cups(texto) -> set:
    return {cups20(c) for c in re.split(r"[,;\s|]+", texto or "") if c.strip()}


def restar_meses(f: dt.date, n: int) -> dt.date:
    total = f.year * 12 + f.month - 1 - n
    a, m = divmod(total, 12)
    ultimo = [31, 29 if a % 4 == 0 and (a % 100 or a % 400 == 0) else 28, 31, 30, 31, 30,
              31, 31, 30, 31, 30, 31][m]
    return dt.date(a, m + 1, min(f.day, ultimo))


def _texto(v) -> str:
    return "" if v is None else str(v).strip()


def _si(v) -> bool:
    return _texto(v).lower() in ("sí", "si", "s", "x", "true", "1")


# ------------------------------------------------------------------- fuente de datos
@dataclass
class Fuente:
    suministros: list = field(default_factory=list)      # listar_sgc_suministros(bajas=False)
    estados: dict = field(default_factory=dict)          # {(id, anio, mes): {estado...}}
    equipo: list = field(default_factory=list)           # listar_sgc_equipo()
    solicitudes: list = field(default_factory=list)      # listar_solicitudes()
    analisis: list = field(default_factory=list)         # listar_analisis()
    catalogo: list = field(default_factory=list)         # listar_catalogo()
    contactos_envio: list = field(default_factory=list)  # listar_sgc_filas("contactos_envio")
    envios: list = field(default_factory=list)           # listar_envios_sgc(desde)
    cambios_gemweb: list = field(default_factory=list)   # propuesta()["cambios"]
    estudios: list = field(default_factory=list)         # listar_estudios_potencia()
    revisiones: list = field(default_factory=list)       # listar_revisiones()
    fichas: dict = field(default_factory=dict)           # info_contratos()


def cargar(hoy: dt.date = None, completa: bool = False) -> Fuente:
    """Lee de la base de datos lo que necesitan el panel y sus páginas. completa: también
    estudios, revisiones y fichas de contrato (para la ficha de cliente)."""
    import almacen
    hoy = hoy or dt.date.today()
    estados = {}
    for anio in (hoy.year - 1, hoy.year):
        estados.update(almacen.listar_sgc_estados(anio))
    try:
        from seguimiento import estados_gemweb as G
        p = G.propuesta()
        cambios = p["cambios"] if p else []
    except Exception:                                       # noqa: BLE001
        cambios = []
    f = Fuente(
        suministros=almacen.listar_sgc_suministros(bajas=False), estados=estados,
        equipo=almacen.listar_sgc_equipo(), solicitudes=almacen.listar_solicitudes(),
        analisis=almacen.listar_analisis(), catalogo=almacen.listar_catalogo(),
        contactos_envio=almacen.listar_sgc_filas("contactos_envio"),
        envios=almacen.listar_envios_sgc((hoy - dt.timedelta(days=60)).isoformat()),
        cambios_gemweb=cambios)
    if completa:
        f.estudios = almacen.listar_estudios_potencia()
        f.revisiones = almacen.listar_revisiones(limite=5000)
        f.fichas = almacen.info_contratos()
    return f


# ------------------------------------------------------------------- de quién es cada cosa
def mis_nombres(usuario: str, equipo: list) -> dict:
    """{"tecnico": {nombres}, "auxiliar": {nombres}} del usuario en la tabla del equipo."""
    salida = {"tecnico": set(), "auxiliar": set()}
    for e in equipo:
        if (e.get("usuario") or "").strip().lower() == (usuario or "").strip().lower():
            nombre = _texto(e.get("nombre"))
            if e.get("tecnico"):
                salida["tecnico"].add(nombre)
            if e.get("auxiliar"):
                salida["auxiliar"].add(nombre)
            if not e.get("tecnico") and not e.get("auxiliar"):
                salida["tecnico"].add(nombre)
                salida["auxiliar"].add(nombre)
    return salida


def es_mio(datos: dict, nombres: dict, como: str = None) -> bool:
    """El suministro es del usuario como técnico, como auxiliar o como cualquiera de los dos."""
    tec = _texto(datos.get(TECNICO)) in nombres.get("tecnico", set())
    aux = _texto(datos.get(AUXILIAR)) in nombres.get("auxiliar", set())
    return tec if como == "tecnico" else aux if como == "auxiliar" else tec or aux


def tiene_nombres(nombres: dict) -> bool:
    return bool(nombres.get("tecnico") or nombres.get("auxiliar"))


# ------------------------------------------------------------------- licitaciones
def fecha_clave(datos: dict, tipo: str = "Automática", margen: int = 3):
    """(fecha, origen) de la licitación del suministro. Automática: lanzamiento previsto; si
    falta, denuncia; si falta, fin de contrato menos `margen` meses."""
    lan, den, fin = (fecha(datos.get(F_LANZAMIENTO)), fecha(datos.get(F_DENUNCIA)),
                     fecha(datos.get(F_FIN)))
    if tipo == "Lanzamiento":
        return lan, "Lanzamiento"
    if tipo == "Denuncia":
        return den, "Denuncia"
    if tipo == "Fin de contrato":
        return fin, "Fin de contrato"
    if lan:
        return lan, "Lanzamiento"
    if den:
        return den, "Denuncia"
    if fin:
        return restar_meses(fin, margen), f"Fin de contrato − {margen} meses"
    return None, ""


def estado_licitacion(cups: set, f_clave: dt.date, solicitudes: list, analisis: list):
    """(estado, origen, id) del trabajo hecho para esos CUPS desde 12 meses antes de la fecha
    clave: el análisis más reciente manda sobre la solicitud."""
    desde = (f_clave - dt.timedelta(days=365)) if f_clave else None
    cups = {cups20(c) for c in cups}

    def reciente(filas):
        validas = [r for r in filas if lista_cups(r.get("cups")) & cups
                   and (not desde or (fecha(r.get("fecha")) or dt.date.min) >= desde)]
        return max(validas, key=lambda r: r.get("fecha") or "", default=None)

    an = reciente(analisis)
    if an:
        est = an.get("estado") or "En curso"
        return ("Adjudicada" if est == "Adjudicada" else "Análisis descartado"
                if est == "Descartada" else "Análisis en curso"), "analisis", an["id"]
    so = reciente(solicitudes)
    if so:
        est = so.get("estado") or "Preparada"
        texto = {"Preparada": "Solicitud preparada", "Enviada": "Solicitud enviada",
                 "Ofertas recibidas": "Ofertas recibidas", "Adjudicada": "Adjudicada",
                 "Cancelada": "Solicitud cancelada"}.get(est, "Solicitud " + est.lower())
        return texto, "solicitud", so["id"]
    return "Sin solicitud", "", None


def licitaciones(f: Fuente, hoy: dt.date, u: dict, tipo: str = "Automática") -> pd.DataFrame:
    """Una fila por suministro activo con fecha clave."""
    filas = []
    for s in f.suministros:
        d = s["datos"]
        clave, origen = fecha_clave(d, tipo, u["margen"])
        if not clave:
            continue
        filas.append({
            "_id": s["id"], CLIENTE: _texto(d.get(CLIENTE)) or _texto(d.get(TITULAR)),
            CUPS: _texto(d.get(CUPS)), DENOMINACION: _texto(d.get(DENOMINACION)),
            "Tarifa": _texto(d.get("Tarifa de acceso")),
            "Comercializadora": _texto(d.get("Comercializadora")),
            "Consumo (MWh)": (float(d.get("Consumo anual (kWh)") or 0) or 0) / 1000,
            TECNICO: _texto(d.get(TECNICO)), AUXILIAR: _texto(d.get(AUXILIAR)),
            "Fecha clave": clave, "Origen de la fecha": origen,
            "Denuncia": fecha(d.get(F_DENUNCIA)), "Fin de contrato": fecha(d.get(F_FIN)),
            "Lanzamiento": fecha(d.get(F_LANZAMIENTO)),
            "Renovación abierta": _si(d.get("Renovación abierta")),
            "Días": (clave - hoy).days})
    return pd.DataFrame(filas)


def agrupar_licitaciones(df: pd.DataFrame, f: Fuente, u: dict) -> pd.DataFrame:
    """Una fila por licitación (mismo cliente y misma fecha clave) con su estado y semáforo."""
    if df.empty:
        return pd.DataFrame()
    filas = []
    for (cliente, clave), g in df.groupby([CLIENTE, "Fecha clave"], sort=False):
        cups = {cups20(c) for c in g[CUPS]}
        estado, origen, id_ = estado_licitacion(cups, clave, f.solicitudes, f.analisis)
        dias = int(g["Días"].iloc[0])
        hay = estado not in ("Sin solicitud", "Solicitud cancelada")
        tecs = sorted({t for t in g[TECNICO] if t})
        filas.append({
            "Semáforo": semaforo(dias, u, hay), CLIENTE: cliente, "Fecha clave": clave,
            "Días": dias, "Origen de la fecha": g["Origen de la fecha"].iloc[0],
            "Suministros": len(g), "Consumo (MWh)": round(float(g["Consumo (MWh)"].sum()), 1),
            "Tarifas": ", ".join(sorted({t for t in g["Tarifa"] if t})),
            "Comercializadora": ", ".join(sorted({c for c in g["Comercializadora"] if c})),
            TECNICO: ", ".join(tecs) or "(sin técnico)",
            AUXILIAR: ", ".join(sorted({a for a in g[AUXILIAR] if a})),
            "Estado": estado, "_origen": origen, "_id_trabajo": id_,
            "_cups": sorted(c for c in g[CUPS] if c), "_ids": list(g["_id"])})
    return pd.DataFrame(filas).sort_values(["Fecha clave", CLIENTE]).reset_index(drop=True)


def carga_por_mes(lic: pd.DataFrame, por_tecnico: bool = False) -> pd.DataFrame:
    """Nº de licitaciones por mes de la fecha clave (y por técnico)."""
    if lic.empty:
        return pd.DataFrame(columns=["Mes", "Licitaciones"] + ([TECNICO] if por_tecnico else []))
    d = lic.assign(Mes=[dt.date(x.year, x.month, 1) for x in lic["Fecha clave"]])
    claves = ["Mes", TECNICO] if por_tecnico else ["Mes"]
    return d.groupby(claves).size().reset_index(name="Licitaciones")


# ------------------------------------------------------------------- informes y facturas
ENVIADO_TXT = C.INF_ENVIADO.lower()
CON_FACTURA = {C.RECIBIDA.lower(), C.PDTE_INF.lower(), C.INF_ENVIADO.lower(),
               C.INF_NO_ENVIADO.lower()}
SIN_FACTURA = {"", C.SIN_FRA.lower()}


def informes_pendientes(f: Fuente, hoy: dt.date, u: dict, nombres: dict = None) -> tuple:
    """(informes, facturas) pendientes del auxiliar según la periodicidad de cada suministro
    «Sgto. y Control» y la regla de retraso de seguimiento/retraso.py:

    - el año se parte en bloques naturales de 1, 3, 6 o 12 meses;
    - fecha de control = día 1 del segundo mes tras acabar el bloque (M+2);
    - informes: meses del bloque que esperan informe (estado distinto de vacío, «Sin fra»,
      «Cadencia» y «Temporal») y no tienen «Inf enviado». Pasada la fecha de control el bloque
      está en RETRASO; antes, semáforo según los días que faltan;
    - facturas: meses ya acabados del bloque con «Sin fra», o vacíos en un año en que el
      suministro ya tiene estados (un año sin ningún estado no se ha llevado en el cuadro),
      ordenados por la fecha de control de su bloque (la factura hace falta para el informe).
    Se revisan los bloques del año anterior y del actual. nombres: solo los del auxiliar."""
    informes, facturas = [], []
    for s in f.suministros:
        d = s["datos"]
        if _texto(d.get("Tipo contratación")) != C.SGTO_Y_CONTROL:
            continue
        if nombres is not None and not es_mio(d, nombres, "auxiliar"):
            continue
        tam = R.tam_bloque(d.get("Cadencia informe"))
        if not tam:
            continue
        base = {"_id": s["id"], CLIENTE: _texto(d.get(CLIENTE)), CUPS: _texto(d.get(CUPS)),
                DENOMINACION: _texto(d.get(DENOMINACION)), AUXILIAR: _texto(d.get(AUXILIAR)),
                TECNICO: _texto(d.get(TECNICO)), "Periodicidad": _texto(d.get("Cadencia informe"))}
        for anio in (hoy.year - 1, hoy.year):
            con_estados = any((f.estados.get((s["id"], anio, m)) or {}).get("estado")
                              for m in range(1, 13))
            for bloque in range(1, -(-12 // tam) + 1):
                ini, fin = (bloque - 1) * tam + 1, min(bloque * tam, 12)
                if dt.date(anio, ini, 1) > hoy:
                    continue                              # bloque que aún no ha empezado
                control = R.fecha_control(anio, fin)
                dias = (control - hoy).days
                if dias < -400:
                    continue
                pend_inf, pend_fra = [], []
                for mes in range(ini, fin + 1):
                    v = _texto((f.estados.get((s["id"], anio, mes)) or {}).get("estado")).lower()
                    if v not in R.NO_ESPERAN and v != ENVIADO_TXT:
                        pend_inf.append(mes)
                    acabado = dt.date(anio, mes, 1) < dt.date(hoy.year, hoy.month, 1)
                    if acabado and (v == C.SIN_FRA.lower() or (v == "" and con_estados)):
                        pend_fra.append(mes)
                periodo = (f"{_mes(anio, ini)}" if ini == fin else
                           f"{_mes(anio, ini)} - {_mes(anio, fin)}")
                comun = dict(base, Bloque=periodo, **{"Fecha de control": control,
                                                      "Días": dias,
                                                      "Semáforo": semaforo_control(dias, u)})
                if pend_inf:
                    informes.append(dict(comun, **{"Meses pendientes": ", ".join(
                        _mes(anio, m) for m in pend_inf), "Pendientes": len(pend_inf)}))
                if pend_fra:
                    facturas.append(dict(comun, **{"Meses sin factura": ", ".join(
                        _mes(anio, m) for m in pend_fra), "Pendientes": len(pend_fra)}))
    orden = lambda x: (x["Fecha de control"], x[CLIENTE], x[CUPS])
    return sorted(informes, key=orden), sorted(facturas, key=orden)


MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov",
                "dic"]


def _mes(anio: int, mes: int) -> str:
    return f"{MESES_CORTOS[mes - 1]}-{anio % 100:02d}"


# ------------------------------------------------------------------- calendario
TIPOS_EVENTO = {"Denuncia": "#C00000", "Fin de contrato": "#7B1E3A",
                "Lanzamiento previsto": "#ED7D31", "Fin de recepción de ofertas": "#2E75B6",
                "Caduca oferta de catálogo": "#70AD47", "Control de informe": "#7030A0"}


def eventos(f: Fuente, desde: dt.date, hasta: dt.date, u: dict = None, hoy: dt.date = None,
            nombres: dict = None) -> list:
    """Vencimientos entre desde y hasta (ambos incluidos). nombres: solo lo del usuario."""
    u = u or DEFECTOS
    salida = []

    def dentro(x):
        return x and desde <= x <= hasta

    for s in f.suministros:
        d = s["datos"]
        if nombres is not None and not es_mio(d, nombres):
            continue
        titulo = " · ".join(x for x in (_texto(d.get(CLIENTE)), _texto(d.get(DENOMINACION)))
                            if x) or _texto(d.get(CUPS))
        for col, tipo in ((F_DENUNCIA, "Denuncia"), (F_FIN, "Fin de contrato"),
                          (F_LANZAMIENTO, "Lanzamiento previsto")):
            x = fecha(d.get(col))
            if dentro(x):
                salida.append({"Fecha": x, "Tipo": tipo, "Título": titulo,
                               CUPS: _texto(d.get(CUPS)), TECNICO: _texto(d.get(TECNICO)),
                               AUXILIAR: _texto(d.get(AUXILIAR)), "_origen": "suministro",
                               "_id": s["id"]})
    for r in f.solicitudes:
        x = fecha(r.get("recepcion_hasta"))
        if dentro(x) and r.get("estado") in ("Preparada", "Enviada"):
            salida.append({"Fecha": x, "Tipo": "Fin de recepción de ofertas",
                           "Título": " · ".join(y for y in (r.get("ref"), r.get("cliente")) if y),
                           CUPS: r.get("cups") or "", TECNICO: r.get("usuario") or "",
                           AUXILIAR: "", "_origen": "solicitud", "_id": r["id"]})
    for r in f.catalogo:
        x = fecha(r.get("vigencia_hasta"))
        if dentro(x):
            salida.append({"Fecha": x, "Tipo": "Caduca oferta de catálogo",
                           "Título": f"{r.get('comercializadora')} · {r.get('nombre')}",
                           CUPS: "", TECNICO: r.get("usuario") or "", AUXILIAR: "",
                           "_origen": "catalogo", "_id": r["id"]})
    informes, _ = informes_pendientes(f, hoy or desde, u, nombres)
    for i in informes:
        if dentro(i["Fecha de control"]):
            salida.append({"Fecha": i["Fecha de control"], "Tipo": "Control de informe",
                           "Título": f"{i[CLIENTE]} · {i[DENOMINACION]} ({i['Bloque']})",
                           CUPS: i[CUPS], TECNICO: i[TECNICO], AUXILIAR: i[AUXILIAR],
                           "_origen": "suministro", "_id": i["_id"]})
    return sorted(salida, key=lambda e: (e["Fecha"], e["Tipo"], e["Título"]))


# ------------------------------------------------------------------- ficha de cliente
def opciones_busqueda(f: Fuente) -> list:
    """Textos para el buscador: clientes, titulares, CIF, CUPS y denominaciones."""
    vistos = {}
    for s in f.suministros:
        d = s["datos"]
        for col, tipo in ((CLIENTE, "Cliente"), (TITULAR, "Titular"), (CIF, "CIF"),
                          (CUPS, "CUPS"), (DENOMINACION, "Suministro")):
            v = _texto(d.get(col))
            if len(v) > 1 and v.strip("-—. ") != "":
                vistos.setdefault(f"{v}  ·  {tipo}", (tipo, v))
    for r in f.solicitudes + f.analisis:
        if r.get("cliente"):
            vistos.setdefault(f"{r['cliente']}  ·  Cliente", ("Cliente", r["cliente"]))
    return sorted(vistos.items(), key=lambda x: x[0].lower())


def buscar(f: Fuente, tipo: str, valor: str) -> dict:
    """{"suministros": [...], "cups": set, "clientes": set} del cliente, CIF, titular, CUPS o
    suministro elegido. Un cliente se amplía a todos sus suministros."""
    valor_n = _texto(valor).lower()
    col = {"Cliente": CLIENTE, "Titular": TITULAR, "CIF": CIF, "CUPS": CUPS,
           "Suministro": DENOMINACION}.get(tipo, CLIENTE)
    sums = [s for s in f.suministros if _texto(s["datos"].get(col)).lower() == valor_n]
    clientes = {_texto(s["datos"].get(CLIENTE)) for s in sums if s["datos"].get(CLIENTE)}
    if tipo == "Cliente":
        clientes.add(_texto(valor))
    if tipo in ("Titular", "CIF") and clientes:
        sums = [s for s in f.suministros if _texto(s["datos"].get(CLIENTE)) in clientes]
    cups = {cups20(s["datos"].get(CUPS)) for s in sums if s["datos"].get(CUPS)}
    if tipo == "CUPS":
        cups.add(cups20(valor))
    return {"suministros": sums, "cups": cups, "clientes": {c for c in clientes if c}}


def ficha(f: Fuente, encontrado: dict) -> dict:
    """Todo lo guardado de esos CUPS o clientes en cada módulo."""
    cups, clientes = encontrado["cups"], {c.lower() for c in encontrado["clientes"]}

    def toca(r, campo_cliente="cliente"):
        return bool(lista_cups(r.get("cups")) & cups) or \
            _texto(r.get(campo_cliente)).lower() in clientes

    return {
        "Solicitudes": [r for r in f.solicitudes if toca(r)],
        "Análisis": [r for r in f.analisis if toca(r)],
        "Estudios de potencia": [r for r in f.estudios if toca(r, "titular")],
        "Revisiones de SSAA": [r for r in f.revisiones if cups20(r.get("cups")) in cups],
        "Fichas de contrato": [{"cups": c, "actualizado": a, "usuario": u}
                               for c, (a, u) in f.fichas.items() if cups20(c) in cups],
    }


# ------------------------------------------------------------------- bloques del panel
@dataclass
class Bloque:
    clave: str
    titulo: str
    rol: str                  # "Técnico", "Auxiliar" o "" (común)
    filas: int = 5
    aviso: str = ""           # texto del contador (vacío = no suma avisos)
    pagina: str = ""          # ruta de «Ver todo»


BLOQUES = [
    Bloque("licitaciones", "Mis próximas licitaciones", "Técnico", 8,
           "Licitaciones en rojo", "licitaciones"),
    Bloque("solicitudes", "Solicitudes con el plazo de recepción cerca o vencido", "Técnico", 5,
           "Solicitudes que vencen", "solicitudes-guardadas"),
    Bloque("analisis", "Mis análisis en curso", "Técnico", 5, "", "analisis-guardados"),
    Bloque("catalogo", "Ofertas de catálogo que caducan pronto", "Técnico", 5,
           "Catálogo que caduca", "catalogo-ofertas"),
    Bloque("renovaciones", "Renovaciones abiertas", "Técnico", 5, "", "licitaciones"),
    Bloque("sin_fechas", "Mis suministros sin fecha de fin o de denuncia", "Técnico", 5,
           "Suministros sin fechas", "seguimiento"),
    Bloque("informes", "Informes por vencer y con retraso", "Auxiliar", 8,
           "Informes en rojo o con retraso", "seguimiento-retraso"),
    Bloque("facturas", "Facturas pendientes de descargar", "Auxiliar", 8,
           "Facturas en rojo o con retraso", "seguimiento"),
    Bloque("gemweb", "Cambios de Gemweb por aceptar", "Auxiliar", 5,
           "Cambios de Gemweb", "seguimiento"),
    Bloque("sin_contacto", "Mis suministros sin contacto de envío", "Auxiliar", 5,
           "Suministros sin contacto de envío", "seguimiento-contactos-envio"),
    Bloque("envios", "Mis últimos envíos", "Auxiliar", 5, "", "seguimiento-envios"),
    Bloque("ultimo", "Seguir donde lo dejaste", "", 6, "", "mi-trabajo"),
    Bloque("favoritos", "Mis favoritos", "", 10, "", ""),
]
POR_CLAVE = {b.clave: b for b in BLOQUES}


def config_por_defecto(rol: str) -> list:
    """Bloques del rol (Técnico, Auxiliar o los dos) y los comunes, en el orden de BLOQUES."""
    import almacen
    roles = almacen.roles_de(rol)
    return [{"bloque": b.clave, "visible": not b.rol or b.rol in roles, "filas": b.filas}
            for b in BLOQUES]


def config_panel(rol: str, guardada: list = None) -> list:
    """La configuración del usuario (orden, visibles, filas) completada con los bloques nuevos
    que no tuviera; la de su rol si no ha guardado ninguna."""
    if not guardada:
        return config_por_defecto(rol)
    salida, vistos = [], set()
    for c in guardada:
        if c.get("bloque") in POR_CLAVE and c["bloque"] not in vistos:
            vistos.add(c["bloque"])
            salida.append({"bloque": c["bloque"], "visible": bool(c.get("visible", True)),
                           "filas": max(1, min(50, int(c.get("filas") or
                                                       POR_CLAVE[c["bloque"]].filas)))})
    for c in config_por_defecto(rol):
        if c["bloque"] not in vistos:
            salida.append(dict(c, visible=False))
    return salida


def contador_por_defecto(rol: str) -> list:
    import almacen
    roles = almacen.roles_de(rol)
    return [b.clave for b in BLOQUES if b.aviso and b.rol in roles]


def _mios(f: Fuente, nombres: dict, como: str, todos: bool) -> list:
    return [s for s in f.suministros if todos or es_mio(s["datos"], nombres, como)]


def datos_bloque(clave: str, f: Fuente, hoy: dt.date, u: dict, usuario: str, nombres: dict,
                 todos: bool = False) -> list:
    """Filas (dicts) de un bloque, de más a menos urgente. todos: no filtrar por el usuario."""
    if clave == "licitaciones":
        sub = Fuente(**{**f.__dict__, "suministros": _mios(f, nombres, "tecnico", todos)})
        lic = agrupar_licitaciones(licitaciones(sub, hoy, u), f, u)
        if lic.empty:
            return []
        lic = lic[(lic["Días"] >= -30) & (lic["Estado"] != "Adjudicada")]
        return lic.sort_values("Días").to_dict("records")
    if clave == "solicitudes":
        filas = []
        for r in f.solicitudes:
            if not todos and r.get("usuario") != usuario:
                continue
            hasta = fecha(r.get("recepcion_hasta"))
            if not hasta or r.get("estado") not in ("Preparada", "Enviada"):
                continue
            dias = (hasta - hoy).days
            if dias <= u["recepcion"]:
                filas.append({"Semáforo": ROJO if dias <= 0 else AMBAR, "Ref.": r.get("ref"),
                              CLIENTE: r.get("cliente"), "Recepción hasta": hasta,
                              "Días": dias, "Estado": r.get("estado"), "_origen": "solicitud",
                              "_id": r["id"]})
        return sorted(filas, key=lambda x: x["Días"])
    if clave == "analisis":
        return [{"Nº": r["id"], CLIENTE: r.get("cliente"), "Ref.": r.get("ref"),
                 "Inicio": fecha(r.get("inicio")), "Ofertas": r.get("ofertas"),
                 "Actualizado": fecha(r.get("actualizado") or r.get("fecha")),
                 "_origen": "analisis", "_id": r["id"]}
                for r in f.analisis if r.get("estado") == "En curso"
                and (todos or usuario in (r.get("usuario"), r.get("actualizado_por")))]
    if clave == "catalogo":
        filas = []
        for r in f.catalogo:
            hasta = fecha(r.get("vigencia_hasta"))
            if hasta and 0 <= (hasta - hoy).days <= u["catalogo"]:
                filas.append({"Semáforo": ROJO if (hasta - hoy).days <= 7 else AMBAR,
                              "Comercializadora": r.get("comercializadora"),
                              "Nombre": r.get("nombre"), "Caduca": hasta,
                              "Días": (hasta - hoy).days, "_origen": "catalogo", "_id": r["id"]})
        return sorted(filas, key=lambda x: x["Días"])
    if clave == "renovaciones":
        return [{CLIENTE: _texto(s["datos"].get(CLIENTE)),
                 DENOMINACION: _texto(s["datos"].get(DENOMINACION)),
                 CUPS: _texto(s["datos"].get(CUPS)),
                 "Fin de contrato": fecha(s["datos"].get(F_FIN)), "_origen": "suministro",
                 "_id": s["id"]}
                for s in _mios(f, nombres, "tecnico", todos)
                if _si(s["datos"].get("Renovación abierta"))]
    if clave == "sin_fechas":
        return [{CLIENTE: _texto(s["datos"].get(CLIENTE)),
                 DENOMINACION: _texto(s["datos"].get(DENOMINACION)),
                 CUPS: _texto(s["datos"].get(CUPS)),
                 "Falta": " y ".join(x for x, c in (("fin", F_FIN), ("denuncia", F_DENUNCIA))
                                     if not fecha(s["datos"].get(c))),
                 "_origen": "suministro", "_id": s["id"]}
                for s in _mios(f, nombres, "tecnico", todos)
                if not fecha(s["datos"].get(F_FIN)) or not fecha(s["datos"].get(F_DENUNCIA))]
    if clave in ("informes", "facturas"):
        informes, facturas = informes_pendientes(f, hoy, u, None if todos else nombres)
        return [dict(x, _origen="suministro") for x in (informes if clave == "informes"
                                                         else facturas)]
    if clave == "gemweb":
        mios = {s["id"] for s in _mios(f, nombres, "auxiliar", todos)}
        return [{CUPS: c.get("cups"), DENOMINACION: c.get("denominacion"),
                 "Mes": _mes(int(c["anio"]), int(c["mes"])) if c.get("anio") else "",
                 "Antes": c.get("antes"), "Después": c.get("despues"),
                 "_origen": "gemweb", "_id": c.get("id_suministro")}
                for c in f.cambios_gemweb if c.get("id_suministro") in mios]
    if clave == "sin_contacto":
        con_contacto = set()
        for r in f.contactos_envio:
            d = r["datos"]
            if _texto(d.get("Para")) and not fecha(d.get("Fecha de baja")):
                if d.get(CUPS):
                    con_contacto.add(cups20(d.get(CUPS)))
                if d.get(CLIENTE):
                    con_contacto.add("cliente:" + _texto(d.get(CLIENTE)).lower())
        return [{CLIENTE: _texto(s["datos"].get(CLIENTE)),
                 DENOMINACION: _texto(s["datos"].get(DENOMINACION)),
                 CUPS: _texto(s["datos"].get(CUPS)), "_origen": "suministro", "_id": s["id"]}
                for s in _mios(f, nombres, "auxiliar", todos)
                if _texto(s["datos"].get("Tipo contratación")) == C.SGTO_Y_CONTROL
                and cups20(s["datos"].get(CUPS)) not in con_contacto
                and "cliente:" + _texto(s["datos"].get(CLIENTE)).lower() not in con_contacto]
    if clave == "envios":
        return [{"Fecha": r.get("fecha", "")[:16].replace("T", " "), "Tipo": r.get("tipo"),
                 CLIENTE: r.get("cliente"), "Asunto": r.get("asunto"),
                 "Resultado": r.get("resultado"), "_origen": "envio", "_id": r["id"]}
                for r in f.envios if (todos or r.get("usuario") == usuario)
                and not r.get("prueba")]
    return []


def avisos(f: Fuente, hoy: dt.date, u: dict, usuario: str, nombres: dict, claves: list,
           todos: bool = False) -> dict:
    """{bloque: nº de avisos} de los bloques elegidos para el contador: lo urgente (rojo o
    RETRASO) de cada uno."""
    salida = {}
    for clave in claves:
        filas = datos_bloque(clave, f, hoy, u, usuario, nombres, todos)
        if clave in ("licitaciones", "informes", "facturas", "catalogo", "solicitudes"):
            filas = [x for x in filas if x.get("Semáforo") in (ROJO, RETRASO)]
        salida[clave] = len(filas)
    return salida
