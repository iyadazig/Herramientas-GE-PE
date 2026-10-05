# -*- coding: utf-8 -*-
"""
HERRAMIENTAS GE&PE — entrada unica de la plataforma
===================================================
    python lanzador.py --servidor          (PC de la oficina, para todo el equipo)
    streamlit run inicio.py                (pruebas en local)

Acceso con usuario y contrasena (paginas.acceso), imagen corporativa (estilo) y menu por
secciones (st.navigation). Cada pagina es un fichero o una funcion; las de Gemweb estan
en gemweb_extractor/paginas.py.
"""

import functools

import streamlit as st

import estilo
import paginas
from gemweb_extractor import paginas as gw

st.set_page_config(page_title="Herramientas GE&PE",
                   page_icon=str(estilo.ICONO) if estilo.ICONO.exists() else None,
                   layout="wide", menu_items={})
estilo.aplicar()
usuario = paginas.acceso()          # detiene la pagina hasta que se inicia sesion

with st.sidebar:
    paginas.barra_usuario()


def _pagina(funcion, titulo, subtitulo, ruta):
    """Pagina de funcion con la cabecera corporativa comun."""
    @functools.wraps(funcion)
    def mostrar():
        estilo.cabecera(titulo, subtitulo)
        funcion()
    return st.Page(mostrar, title=titulo, url_path=ruta)


SUB_SSAA = "Comprobación del concepto de SSAA con los datos publicados por REE (ESIOS)"
SUB_GEMWEB = "Consultas a la API de Gemweb"
menu = {
    "Revisión de SSAA": [
        st.Page("app_revision_ssaa.py", title="Revisar factura", url_path="revisar",
                default=True),
        _pagina(paginas.pagina_historial, "Historial de revisiones", SUB_SSAA, "historial"),
        _pagina(paginas.pagina_fichas, "Fichas de contrato", SUB_SSAA, "fichas"),
    ],
    "Gemweb": [
        _pagina(funcion, titulo, SUB_GEMWEB, ruta) for funcion, titulo, ruta in (
            (gw.pagina_inventario, "Inventario", "gemweb-inventario"),
            (gw.pagina_consumo_periodo, "Consumo por periodo", "gemweb-consumo-periodo"),
            (gw.pagina_telelecturas, "Telelecturas", "gemweb-telelecturas"),
            (gw.pagina_descarga_masiva, "Descarga masiva por CUPS", "gemweb-descarga-masiva"),
            (gw.pagina_coste_consumo, "Coste y consumo mensual", "gemweb-coste-consumo"),
            (gw.pagina_optimizacion, "Optimización de potencia", "gemweb-optimizacion"),
            (gw.pagina_incoherencias, "Incoherencias en facturas", "gemweb-incoherencias"),
            (gw.pagina_factura_calculada, "Factura calculada", "gemweb-factura-calculada"))
    ],
}
if usuario["admin"]:
    menu["Administración"] = [
        _pagina(paginas.pagina_usuarios, "Usuarios", "Altas, contraseñas y permisos",
                "usuarios"),
        _pagina(paginas.pagina_credenciales, "Credenciales de Gemweb",
                "Acceso a la API de Gemweb para toda la plataforma", "credenciales"),
    ]

st.navigation(menu, expanded=True).run()
estilo.pie()
