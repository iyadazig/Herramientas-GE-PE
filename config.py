# -*- coding: utf-8 -*-
"""Rutas y parametros comunes del revisor de SSAA."""

import os
import sys
from pathlib import Path

# Como ejecutable (PyInstaller) el codigo se descomprime en una carpeta temporal que
# se borra al cerrar: los DATOS (fichas, revisiones...) van siempre junto al .exe y los
# RECURSOS (logo, icono) se leen del paquete.
EJECUTABLE = getattr(sys, "frozen", False)
if EJECUTABLE:
    CARPETA = Path(os.environ.get("SSAA_CARPETA_PROGRAMA", Path(sys.executable).resolve().parent))
    RECURSOS = Path(getattr(sys, "_MEIPASS", CARPETA))
else:
    CARPETA = Path(__file__).resolve().parent
    RECURSOS = CARPETA

# Excel historicos de ESIOS: los de la carpeta de red, que actualizan los scripts de
# ESIOS_Scraper (Programador de tareas del PC de la oficina). Si no se llega a la red, los
# locales de Descarga_datos_ESIOS. Se puede cambiar con SSAA_CARPETA_ESIOS (--carpeta-esios).
CARPETA_ESIOS_RED = Path(r"\\Svplaura-sys\energia 1\Tecnicos\Datos de mercado\ESIOS")
CARPETA_ESIOS_LOCAL = CARPETA.parent / "Descarga_datos_ESIOS"


def _carpeta_esios():
    if os.environ.get("SSAA_CARPETA_ESIOS"):
        return Path(os.environ["SSAA_CARPETA_ESIOS"])
    try:
        if CARPETA_ESIOS_RED.is_dir():
            return CARPETA_ESIOS_RED
    except OSError:
        pass
    return CARPETA_ESIOS_LOCAL


CARPETA_ESIOS = _carpeta_esios()

# Datos de mercado para el analisis de ofertas (solo lectura): OMIP diario (OMIP_Scraper) y
# medias mensuales de OMIE (Excel mantenido a mano). Hora a hora del mercado diario: los
# Excel mensuales de la carpeta Omie (respaldo de ESIOS para los apuntamientos).
CARPETA_MERCADO = Path(os.environ.get(
    "SSAA_CARPETA_MERCADO", CARPETA_ESIOS_RED.parent))
FICHERO_OMIP = CARPETA_MERCADO / "OMIP" / "OMIP_Scraper" / "data" / "omip_history.xlsx"
FICHERO_OMIE_MEDIO = CARPETA_MERCADO / "precios omel" / "PRECIO MEDIO OMIE.xls"
CARPETA_OMIE_HORARIO = CARPETA_MERCADO / "Omie"

# Base de datos compartida (usuarios, fichas, historial). En el servidor se puede llevar a
# otra carpeta con SSAA_CARPETA_DATOS.
CARPETA_DATOS = Path(os.environ.get("SSAA_CARPETA_DATOS", CARPETA))
# fichero de fichas de la version anterior: se importa a la base de datos la primera vez
FICHERO_CONTRATOS = CARPETA / "contratos_ssaa.json"
CARPETA_REVISIONES = CARPETA / "revisiones_ssaa"
CARPETA_EJEMPLOS = CARPETA / "facturas_ejemplo"

# Servidor para todo el equipo (lanzador.py --servidor): sin boton de cerrar la app
MODO_SERVIDOR = os.environ.get("SSAA_MODO_SERVIDOR") == "1"

# Tolerancia por defecto para dar una factura por correcta (% sobre lo recalculado).
TOLERANCIA_PCT = 0.5
# Un SSAA horario por encima de esto se marca como sospechoso (dato de avance roto).
UMBRAL_ANOMALO = 200.0
