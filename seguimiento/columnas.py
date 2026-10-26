# -*- coding: utf-8 -*-
"""
Columnas de las tablas de Seguimiento y Control
===============================================
Cada columna: clave (la que se guarda en el JSON de la base de datos y se ve en la app),
titulos con los que aparece en el Excel (el primero es el que se usa al exportar), tipo
(texto, numero, fecha, calculada) y, si sus valores salen de una lista, el nombre de la lista
(sgc_listas). Ningun dato de clientes vive aqui: solo la estructura.
"""

import unicodedata

TEXTO, NUMERO, FECHA, CALCULADA = "texto", "numero", "fecha", "calculada"
OTROS = "Otros"           # columnas del Excel sin titulo que traen datos


def col(clave, tipo=TEXTO, lista=None, excel=(), grupo=None):
    return {"clave": clave, "tipo": tipo, "lista": lista, "excel": (clave,) + tuple(excel),
            "grupo": grupo}


# ------------------------------------------------------------------- cuadro (CLIENTES)
CLIENTE, CONTRATO, CUOTAS, OBSERVACIONES = "Cliente", "Contrato", "Cuotas", "Observaciones"

CUOTAS_SUMADAS = ("Servicio de SyC (€/mes)", "Servicio SyC FV (€/mes)", "Servicio SGE (€/mes)",
                  "Servicio Seg. Mercados (€/mes)")
TOTAL_CUOTA = "Total cuota SyC (€/mes)"

CUADRO = [
    col("Técnico", lista="tecnicos", grupo=CLIENTE),
    col("Auxiliar", lista="auxiliares", grupo=CLIENTE),
    col("Otra información", lista="caracteristicas", grupo=CLIENTE),
    col("Tipo contratación", lista="tipo_contratacion", excel=("Tipo Contratación",),
        grupo=CLIENTE),
    col("Cadencia informe", lista="periodicidad", grupo=CLIENTE),
    col("Cliente/Grupo", grupo=CLIENTE),
    col("CIF", grupo=CLIENTE),
    col("Titular de contrato", grupo=CLIENTE),
    col("CUPS", grupo=CLIENTE),
    col("Denominación suministro", grupo=CLIENTE),
    col("Tipo suministro", lista="tipos_suministro",
        excel=("Tipo suministro (Electricidad / Gas)",), grupo=CONTRATO),
    col("Tarifa de acceso", lista="tarifas", grupo=CONTRATO),
    col("Max. potencia P3/P6 (kW)", NUMERO, grupo=CONTRATO),
    col("Consumo anual (kWh)", NUMERO, grupo=CONTRATO),
    col("Fecha denuncia contrato", FECHA, grupo=CONTRATO),
    col("Fecha fin contrato", FECHA, grupo=CONTRATO),
    col("Fecha prevista lanzamiento licitación", FECHA, grupo=CONTRATO),
    col("Comercializadora", lista="comercializadoras", grupo=CONTRATO),
    col("Tipo de contrato", lista="tipos_contrato",
        excel=("Tipo de contrato (Precio fijo / Indexado / Indexado+coberturas)",),
        grupo=CONTRATO),
    col("Provincia", lista="provincias", grupo=CONTRATO),
    col("Gestor Comercial", lista="gestores", grupo=CONTRATO),
    col("Servicio de SyC (€/mes)", NUMERO, grupo=CUOTAS),
    col("Servicio SyC FV (€/mes)", NUMERO, grupo=CUOTAS),
    col("Servicio SGE (€/mes)", NUMERO, grupo=CUOTAS),
    col("Servicio Seg. Mercados (€/mes)", NUMERO, grupo=CUOTAS),
    col(TOTAL_CUOTA, CALCULADA, grupo=CUOTAS),
    col("Servicio SMAC (€/mes)", NUMERO, grupo=CUOTAS),
    col("Renovación abierta", lista="si_no", excel=("RENOVACIÓN ABIERTA",), grupo=CONTRATO),
    # aqui van los 12 meses (estado + Obs), que se guardan en sgc_estados
    col("Observaciones", grupo=OBSERVACIONES),
    col("Observaciones Extra", grupo=OBSERVACIONES),
    col("Validación Gemweb", lista="si_no", grupo=OBSERVACIONES),
    col("Baja en GEYPE", lista="si_no", excel=("BAJA EN GEYPE",), grupo=OBSERVACIONES),
    col("Estado en Gemweb", excel=("ESTADO EN GEMWEB",), grupo=OBSERVACIONES),
    col("Informes y avisos fuera de GW", excel=("INFORMES Y AVISOS FUERA DE GW",),
        grupo=OBSERVACIONES),
    col("CNAE", grupo=OBSERVACIONES),
    col("Referencia catastral", excel=("REFERENCIA CATASTRAL",), grupo=OBSERVACIONES),
    col("Fecha de baja", FECHA, grupo=OBSERVACIONES),
    col(OTROS, grupo=OBSERVACIONES),
]
POSICION_MESES = [c["clave"] for c in CUADRO].index("Renovación abierta") + 1
GRUPOS_CUADRO = [CLIENTE, CONTRATO, CUOTAS, "Meses", OBSERVACIONES]

# ------------------------------------------------------------------- tablas auxiliares
_COMUNES = [col("Técnico", lista="tecnicos"), col("Auxiliar", lista="auxiliares")]

TABLAS = {
    "contactos_envio": {
        "titulo": "Contactos de envío",
        "columnas": _COMUNES + [
            col("Cliente/Grupo"), col("Denominación suministro"), col("CUPS"),
            col("Fecha de baja", FECHA), col("Nombre cliente"), col("Para"), col("CC")]},
    "contactos": {
        "titulo": "Contactos de clientes", "hoja": "CONTACTOS",
        "columnas": _COMUNES + [
            col("Cliente/Grupo"), col("Denominación suministro"), col("Contacto"),
            col("Puesto"), col("email"), col("Teléfono"), col("Comentarios"), col(OTROS)]},
    "claves": {
        "titulo": "Claves de clientes", "hoja": "CLAVES",
        "columnas": _COMUNES + [
            col("Cliente/Grupo"), col("Denominación suministro"), col("Usuario"), col("Clave"),
            col("Comercializadora"), col("Enlace web"), col("Gestor"), col("email"),
            col("Teléfono"), col(OTROS)]},
    "smac": {
        "titulo": "SMAC", "hoja": "SMAC",
        "columnas": _COMUNES + [
            col("Código Pedido"), col("Cliente/Grupo"), col("Denominación suministro"),
            col("Usuario"), col("Clave"), col("Último dato", FECHA), col("Observaciones")]},
    "telemedida": {
        "titulo": "Telemedida", "hoja": "TELEMEDIDA",
        "columnas": _COMUNES + [
            col("Cliente/Grupo"), col("Denominación suministro"),
            col("Teléfono", excel=("TELÉFONO",)), col("D. enlace", excel=("D. ENLACE",)),
            col("P.M."), col("C.L."), col("Módem", excel=("MODEM",)), col("IP"),
            col("Paridad", excel=("PARIDAD",)), col("Velocidad", excel=("VELOCIDAD",)),
            col("Técnico telemedida", excel=("TÉCNICO",)),
            col("Observaciones", excel=("OBSERVACIONES",))]},
}

# ------------------------------------------------------------------- estados de cada mes
SIN_FRA, RECIBIDA, PDTE_INF = "Sin fra", "Recibida", "Pdte inf"
INF_ENVIADO, INF_NO_ENVIADO, CADENCIA, TEMPORAL = (
    "Inf enviado", "Inf no enviado", "Cadencia", "Temporal")
ESTADOS = [  # valor, descripcion, color (los de la pestaña «Lista» del Excel)
    (SIN_FRA, "Factura sin registrar o recibir.", "#FFFFFF"),
    (RECIBIDA, "Factura recibida, pero sin validar.", "#ED7D31"),
    (PDTE_INF, "Factura pendiente de informe.", "#FF0000"),
    (INF_ENVIADO, "Informe realizado y enviado.", "#A5A5A5"),
    (INF_NO_ENVIADO, "Informe realizado pero no enviado.", "#7030A0"),
    (CADENCIA, "Recordatorio cadencia informe (no mensuales).", "#FFFF00"),
    (TEMPORAL, "Temporal o aún no ha comenzado el SyC.", "#70AD47"),
]
COLOR_ESTADO = {v: c for v, _, c in ESTADOS}

SGTO_Y_CONTROL = "Sgto. y Control"

# Bloques de la pestaña «Lista» del Excel: titulo del bloque -> nombre de la lista
LISTAS_EXCEL = {
    "CARACTERÍSTICAS DESTACABLES": "caracteristicas", "TÉCNICOS": "tecnicos",
    "AUXILIARES": "auxiliares", "TIPOS DE SUMINISTRO": "tipos_suministro",
    "TARIFAS DE ACCESO": "tarifas", "COMERCIALIZADORAS": "comercializadoras",
    "TIPOS DE CONTRATO": "tipos_contrato", "PROVINCIAS": "provincias",
    "GESTOR COMERCIAL": "gestores", "PERIODICIDAD": "periodicidad",
    "RENOVACIÓN ABIERTA": "si_no", "TIPO CONTRATACIÓN": "tipo_contratacion",
    "OBSERVACIONES": "observaciones",
}
TITULOS_LISTAS = {
    "caracteristicas": "Características destacables", "tecnicos": "Técnicos",
    "auxiliares": "Auxiliares", "tipos_suministro": "Tipos de suministro",
    "tarifas": "Tarifas de acceso", "comercializadoras": "Comercializadoras",
    "tipos_contrato": "Tipos de contrato", "provincias": "Provincias",
    "gestores": "Gestores comerciales", "periodicidad": "Cadencia del informe",
    "si_no": "Sí / No", "tipo_contratacion": "Tipo de contratación",
    "observaciones": "Observaciones habituales", "estados": "Estados de cada mes",
}

MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov",
                "dic"]


def normalizar(texto):
    """Titulo sin acentos, espacios repetidos ni mayusculas, para reconocer columnas."""
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(t.lower().split())


def columna_mes(anio, mes):
    return "%s-%02d" % (MESES_CORTOS[mes - 1], anio % 100)


def columna_obs(anio, mes):
    return "Obs %s" % columna_mes(anio, mes)
