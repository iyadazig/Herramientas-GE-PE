# -*- coding: utf-8 -*-
"""
MODELO DEL ANALISIS DE OFERTAS
==============================
Estudio (cliente, suministros, mercado, SSAA y apuntamientos), contrato actual y ofertas, con
las mismas opciones que la plantilla Excel multipunto. Se guardan como JSON (a_dict/de_dict).
Sin Streamlit.

Unidades: consumos en kWh, potencias en kW, precios de mercado, SSAA, fees, GdO, FNEE y
término B en €/MWh, potencia en €/kW·año, energía fija en la unidad de la oferta (€/kWh,
c€/kWh o €/MWh), porcentajes en tanto por uno.
"""

import copy
import datetime as dt
from dataclasses import asdict, dataclass, field, fields

PERIODOS = ["P1", "P2", "P3", "P4", "P5", "P6"]
TARIFAS = ["2.0TD", "3.0TD", "3.0TDVE", "6.1TD", "6.1TDVE", "6.2TD", "6.3TD", "6.4TD"]
SISTEMAS = ["Península", "Baleares", "Canarias", "Ceuta", "Melilla"]

# listas de opciones (las de la plantilla)
FIJO, PASS_POOL, PASS_THROUGH = "Fijo", "Indexado Pass-Pool", "Indexado Pass-Through"
MODALIDADES = [FIJO, PASS_POOL, PASS_THROUGH]
UNIDADES = {"€/kWh": 1, "c€/kWh": 100, "€/MWh": 1000}
POT_REGULADA, POT_PROPIA = "Regulada (ATR)", "Propia de la oferta"
NETO, BARRAS = "Consumo neto", "Barras central"
BASES = [NETO, BARRAS]
FNEE_INCLUIDO = "Incluido"
MODOS_FNEE = [FNEE_INCLUIDO, NETO, BARRAS]
SSAA_INCLUIDOS, SSAA_PERIODO, SSAA_PLANO, SSAA_BANDA, SSAA_TECHO = (
    "Incluidos", "Indexado por periodo", "Indexado plano", "Banda", "Techo")
MODOS_SSAA = [SSAA_INCLUIDOS, SSAA_PERIODO, SSAA_PLANO, SSAA_BANDA, SSAA_TECHO]
PERD_TARIFA, PERD_FIJO, PERD_NO = "Pérdidas tarifa", "% fijo", "No"
PERDIDAS_SSAA = [PERD_TARIFA, PERD_FIJO, PERD_NO]
SIN_CIERRES, CIERRE_PCT, CIERRE_MW = "Sin cierres", "Por %", "Por MW (carga base)"
TIPOS_CIERRE = [SIN_CIERRES, CIERRE_PCT, CIERRE_MW]
REDUCCIONES_IEE = [0.0, 0.85, 1.0]
ESTIMADO, CLIENTE = "Estimado (a mano)", "Del cliente (curva)"


def seis(valores, defecto=None) -> list:
    """Lista de 6 valores (rellena con `defecto`)."""
    valores = list(valores or [])[:6]
    return valores + [defecto] * (6 - len(valores))


def clave_ap(sistema: str, calendario: str) -> str:
    """«Península|6P»."""
    return f"{sistema}|{calendario}"


def calendario_de(tarifa: str) -> str:
    return "3P" if tarifa == "2.0TD" else "6P"


def n_periodos(tarifa: str) -> int:
    return 3 if tarifa == "2.0TD" else 6


def grupo_automatico(tarifa: str, sistema: str) -> str:
    """Grupo de precios de una tarifa en un sistema: «6.1TD» (Península), «6.1TD-Canarias»."""
    return tarifa if sistema == "Península" else f"{tarifa}-{sistema}"


def sumar_meses(fecha: dt.date, meses: int) -> dt.date:
    a, m = divmod(fecha.month - 1 + meses, 12)
    return dt.date(fecha.year + a, m + 1, 1)


# ------------------------------------------------------------------- suministros
@dataclass
class Suministro:
    cups: str
    tarifa: str = "6.1TD"
    sistema: str = "Península"
    alias: str = ""
    razon: str = ""
    direccion: str = ""
    potencias: list = field(default_factory=lambda: [0.0] * 6)          # kW
    consumo: dict = field(default_factory=dict)       # {mes natural 1..12: [6 kWh]}
    reduccion_iee: float = 0.0                         # 0, 0,85 o 1
    alquiler: float = 0.0                              # €/mes (equipo de medida, servicios)
    activo: bool = True                                # entra en el estudio
    apuntamiento: str = ESTIMADO                       # para los indexados, si tiene curva

    def consumo_mes(self, mes: int) -> list:
        return [float(v or 0) for v in seis(self.consumo.get(mes) or self.consumo.get(str(mes)),
                                            0.0)]

    def consumo_anual(self) -> float:
        return sum(sum(self.consumo_mes(m)) for m in range(1, 13))


# ------------------------------------------------------------------- ofertas
@dataclass
class PreciosAnio:
    """Precios de un grupo en un año de contrato."""
    energia: list = field(default_factory=lambda: [None] * 6)      # fijo, unidad de la oferta
    termino_b: list = field(default_factory=lambda: [None] * 6)    # indexados, €/MWh
    ci: list = field(default_factory=lambda: [1.0] * 6)            # coeficiente Pass-Pool
    potencia: list = field(default_factory=lambda: [None] * 6)     # propia, €/kW·año


@dataclass
class GrupoPropio:
    nombre: str
    tarifa: str
    sistema: str = "Península"


@dataclass
class Cierre:
    """Cierre en OMIP de un mes de contrato (1 = primer mes)."""
    mes: int
    pct: float = None             # «Por %»: parte del consumo cerrada (tanto por uno)
    mw: float = None              # «Por MW»: MW de carga base
    precio: float = None          # €/MWh; vacío = OMIP del mes


@dataclass
class Oferta:
    nombre: str = ""
    comercializadora: str = ""
    gestor: str = ""                       # gestor de la comercializadora
    email_gestor: str = ""
    producto: str = ""
    ronda: int = 1
    es_actual: bool = False
    modalidad: str = FIJO
    duracion: int = 12                     # meses, cualquiera
    inicio: dt.date = None                 # vacío = el del estudio
    unidad: str = "€/kWh"
    atr_incluido: bool = True              # el precio de energía incluye peajes y cargos
    potencia: str = POT_REGULADA
    descuento: float = 0.0                 # sobre el término de energía (fijo)
    base_b: str = NETO                     # base del término B (indexados)
    pagos_capacidad: bool = False
    os: bool = False
    om: bool = False
    srad: bool = False
    desvios: float = 0.0                   # €/MWh en barras
    fee: float = 0.0                       # €/MWh
    base_fee: str = NETO
    k: float = 0.0                         # K / impuesto municipal
    k_ssaa: bool = False
    gdo: float = 0.0                       # €/MWh
    fnee: str = FNEE_INCLUIDO
    otros_fijos: float = 0.0               # €/CUPS·mes
    ssaa_modo: str = SSAA_INCLUIDOS
    banda_min: float = 0.0
    banda_max: float = 0.0
    techo: float = 0.0
    ssaa_perfil: bool = False
    ssaa_perdidas: str = PERD_NO
    ssaa_pct_perdidas: float = 0.0
    ssaa_f1: float = 1.0
    ssaa_f2: float = 1.0
    cierres_tipo: str = SIN_CIERRES
    coste_click: float = 0.0               # €/MWh cerrado
    cierres: list = field(default_factory=list)              # [Cierre]
    precios: dict = field(default_factory=dict)              # {grupo: [PreciosAnio por año]}
    grupos_propios: list = field(default_factory=list)       # [GrupoPropio]
    asignacion: dict = field(default_factory=dict)           # {cups: grupo propio}
    excluidos: list = field(default_factory=list)            # CUPS que no entran
    apuntamientos_propios: dict = field(default_factory=dict)   # {"Península|6P": {mes: [6]}}
    condiciones: dict = field(default_factory=dict)          # validez, pago, garantías...
    notas: str = ""

    @property
    def indexada(self) -> bool:
        return self.modalidad != FIJO

    def precios_grupo(self, grupo: str, anio_contrato: int, tarifa: str = None) -> PreciosAnio:
        """Precios del grupo en el año de contrato (1, 2...). Un año sin precios propios usa
        los del anterior; un grupo territorial sin precios, los de su tarifa (como la
        plantilla)."""
        lista = self.precios.get(grupo)
        if not lista and tarifa and grupo != tarifa:
            lista = self.precios.get(tarifa)
        if not lista:
            return PreciosAnio()
        return lista[min(anio_contrato, len(lista)) - 1]

    def resumen_ssaa(self) -> str:
        return {SSAA_INCLUIDOS: "SSAA incluidos", SSAA_PERIODO: "SSAA index.",
                SSAA_PLANO: "SSAA index. plano", SSAA_BANDA: "SSAA banda",
                SSAA_TECHO: "SSAA techo"}.get(self.ssaa_modo, self.ssaa_modo)


# ------------------------------------------------------------------- estudio
@dataclass
class Estudio:
    cliente: str = ""
    cif: str = ""
    fecha: dt.date = None
    inicio: dt.date = None                 # inicio de suministro (primer día de mes)
    crecimiento: list = field(default_factory=list)          # año 2, 3... (tanto por uno)
    suministros: list = field(default_factory=list)          # [Suministro]
    # mercado: {"AAAA-MM": {"omip", "contrato", "omie_media", "anios_omie"}}; pct_omip
    mercado: dict = field(default_factory=dict)
    pct_omip: float = 0.95
    sesion_omip: str = ""
    omip_manual: dict = field(default_factory=dict)          # {"AAAA-MM": €/MWh}
    ssaa_mes: list = field(default_factory=lambda: [None] * 12)   # €/MWh por mes natural
    ssaa_manual: dict = field(default_factory=dict)          # {"AAAA-MM": €/MWh}
    ap_omie: dict = field(default_factory=dict)              # {"Península|6P": {mes: [6]}}
    ap_ssaa: dict = field(default_factory=dict)              # vacío = el de OMIE
    ap_origen: dict = field(default_factory=dict)            # {"Península|6P": {mes: texto}}
    ap_cliente: dict = field(default_factory=dict)           # {cups: {"omie": {mes: [6]},
    #                                                           "ssaa": {...}, "fuentes": ...}}
    regulados: dict = field(default_factory=dict)            # {año: datos regulados congelados}
    actual: Oferta = None
    ofertas: list = field(default_factory=list)              # [Oferta]
    id_solicitud: int = None
    duraciones_solicitadas: list = field(default_factory=list)   # meses pedidos (informativo)

    def suministro(self, cups: str) -> Suministro:
        return next((s for s in self.suministros if s.cups == cups), None)

    def horizonte(self) -> int:
        """Meses que hay que calcular: la oferta que más dura (desde el inicio del estudio)."""
        fin = 0
        for o in self.ofertas + ([self.actual] if self.actual else []):
            desfase = 0
            if o.inicio and self.inicio:
                desfase = (o.inicio.year - self.inicio.year) * 12 + o.inicio.month - \
                    self.inicio.month
            fin = max(fin, desfase + int(o.duracion or 0))
        return max(fin, 12)


# ------------------------------------------------------------------- JSON
def _fecha_a_texto(v):
    return v.isoformat() if isinstance(v, dt.date) else v


def _texto_a_fecha(v):
    if isinstance(v, str) and v:
        return dt.date.fromisoformat(v[:10])
    return v if isinstance(v, dt.date) else None


def _meses_int(d: dict) -> dict:
    """Claves de mes '1'..'12' (JSON) a enteros."""
    return {int(k): v for k, v in (d or {}).items()}


def oferta_a_dict(o: Oferta) -> dict:
    d = asdict(o)
    d["inicio"] = _fecha_a_texto(o.inicio)
    return d


def oferta_de_dict(d: dict) -> Oferta:
    if d is None:
        return None
    d = copy.deepcopy(d)
    validos = {f.name for f in fields(Oferta)}
    d = {k: v for k, v in d.items() if k in validos}
    d["inicio"] = _texto_a_fecha(d.get("inicio"))
    d["cierres"] = [Cierre(**c) for c in d.get("cierres") or []]
    d["precios"] = {g: [PreciosAnio(**p) for p in lista]
                    for g, lista in (d.get("precios") or {}).items()}
    d["grupos_propios"] = [GrupoPropio(**g) for g in d.get("grupos_propios") or []]
    d["apuntamientos_propios"] = {k: _meses_int(v)
                                  for k, v in (d.get("apuntamientos_propios") or {}).items()}
    return Oferta(**d)


def estudio_a_dict(e: Estudio) -> dict:
    d = asdict(e)
    d["fecha"] = _fecha_a_texto(e.fecha)
    d["inicio"] = _fecha_a_texto(e.inicio)
    d["actual"] = oferta_a_dict(e.actual) if e.actual else None
    d["ofertas"] = [oferta_a_dict(o) for o in e.ofertas]
    d["regulados"] = {str(a): v for a, v in e.regulados.items()}
    return d


def estudio_de_dict(d: dict) -> Estudio:
    d = copy.deepcopy(d)
    validos = {f.name for f in fields(Estudio)}
    d = {k: v for k, v in d.items() if k in validos}
    d["fecha"] = _texto_a_fecha(d.get("fecha"))
    d["inicio"] = _texto_a_fecha(d.get("inicio"))
    sv = {f.name for f in fields(Suministro)}
    d["suministros"] = [Suministro(**{k: v for k, v in s.items() if k in sv})
                        for s in d.get("suministros") or []]
    for s in d["suministros"]:
        s.consumo = _meses_int(s.consumo)
    for clave in ("ap_omie", "ap_ssaa", "ap_origen"):
        d[clave] = {k: _meses_int(v) for k, v in (d.get(clave) or {}).items()}
    d["ap_cliente"] = {c: {k: (_meses_int(v) if isinstance(v, dict) else v)
                           for k, v in a.items()}
                       for c, a in (d.get("ap_cliente") or {}).items()}
    d["regulados"] = {int(a): v for a, v in (d.get("regulados") or {}).items()}
    d["actual"] = oferta_de_dict(d.get("actual"))
    d["ofertas"] = [oferta_de_dict(o) for o in d.get("ofertas") or []]
    return Estudio(**d)


def duplicar(o: Oferta, nombre: str = None, nueva_ronda: bool = False) -> Oferta:
    """Copia de una oferta (para «Duplicar» o «Nueva ronda»)."""
    nueva = oferta_de_dict(oferta_a_dict(o))
    nueva.es_actual = False
    if nueva_ronda:
        nueva.ronda = o.ronda + 1
    nueva.nombre = nombre or (f"{o.nombre} (ronda {nueva.ronda})" if nueva_ronda
                              else f"{o.nombre} (copia)")
    return nueva
