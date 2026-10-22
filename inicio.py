# -*- coding: utf-8 -*-
"""
HERRAMIENTAS GE&PE — entrada unica de la plataforma
===================================================
    python lanzador.py --servidor          (PC de la oficina, para todo el equipo)
    streamlit run inicio.py                (pruebas en local)

Acceso con usuario y contrasena (paginas.acceso), imagen corporativa (estilo) y menu por
secciones (st.navigation). Cada pagina es un fichero o una funcion; las de Gemweb estan
en gemweb_extractor/paginas.py y las de la
solicitud de ofertas en ofertas/paginas.py.
"""

import functools

import streamlit as st

import estilo
import paginas
from gemweb_extractor import paginas as gw
from ofertas import paginas as ofertas

st.set_page_config(page_title="Herramientas GE&PE",
                   page_icon=str(estilo.ICONO) if estilo.ICONO.exists() else None,
                   layout="wide", menu_items={})
estilo.aplicar()
usuario = paginas.acceso()          # detiene la pagina hasta que se inicia sesion
estilo.logo_menu()                  # no sale en la pagina de inicio de sesion

with st.sidebar:
    paginas.barra_usuario()
    if gw.AVISOS_GEMWEB:
        with st.expander(":material/warning: Fallos de la API de Gemweb"):
            st.caption("Estas opciones no funcionan del todo por falta de implementación o "
                       "mejora de la API de Gemweb (no de esta aplicación).")
            for _titulo, _texto in gw.AVISOS_GEMWEB.items():
                st.markdown("**%s**: %s" % (_titulo, _texto))


def _pagina(funcion, titulo, subtitulo, ruta):
    """Pagina de funcion con la cabecera corporativa comun. Si la API de Gemweb tiene un fallo
    conocido en esa opcion, se marca en el menu y se avisa al abrirla."""
    aviso = gw.AVISOS_GEMWEB.get(titulo)

    @functools.wraps(funcion)
    def mostrar():
        estilo.cabecera(titulo, subtitulo)
        if aviso:
            st.warning("Fallo de la API de Gemweb: " + aviso, icon=":material/warning:")
        funcion()
    return st.Page(mostrar, title=titulo, url_path=ruta,
                   icon=":material/warning:" if aviso else None)


SUB_SSAA = "Comprobación del concepto de SSAA con los datos publicados por REE (ESIOS)"
SUB_GEMWEB = "Consultas a la API de Gemweb"
menu = {
    "Revisión de SSAA": [
        st.Page("app_revision_ssaa.py", title="Revisar factura", url_path="revisar",
                default=True),
        _pagina(paginas.pagina_historial, "Historial de revisiones", SUB_SSAA, "historial"),
        _pagina(paginas.pagina_fichas, "Fichas de contrato", SUB_SSAA, "fichas"),
    ],
    "API de Gemweb": [
        _pagina(funcion, titulo, SUB_GEMWEB, ruta) for funcion, titulo, ruta in (
            (gw.pagina_inventario, "Inventario", "gemweb-inventario"),
            (gw.pagina_consumo_periodo, "Consumo por periodo", "gemweb-consumo-periodo"),
            (gw.pagina_telelecturas, "Telelecturas", "gemweb-telelecturas"),
            (gw.pagina_descarga_masiva, "Descarga masiva por CUPS", "gemweb-descarga-masiva"),
            (gw.pagina_coste_consumo, "Coste y consumo mensual", "gemweb-coste-consumo"),
            (gw.pagina_optimizacion, "Optimización de potencia", "gemweb-optimizacion"),
            (gw.pagina_incoherencias, "Incoherencias en facturas", "gemweb-incoherencias"),
            (gw.pagina_factura_calculada, "Factura calculada", "gemweb-factura-calculada"),
            (gw.pagina_simulacion_contrato, "Simulación de contrato",
             "gemweb-simulacion-contrato"))
    ],
    "Solicitud de ofertas": [
        _pagina(ofertas.pagina_solicitud_ofertas, "Crear nueva solicitud",
                "Excel de solicitud de ofertas con los datos de Gemweb", "solicitud-ofertas"),
        _pagina(ofertas.pagina_solicitudes_guardadas, "Solicitudes guardadas",
                "Solicitudes de ofertas guardadas por el equipo y su seguimiento",
                "solicitudes-guardadas"),
    ],
}
ofertas.PAGINA_FORMULARIO = menu["Solicitud de ofertas"][0]   # para «Cargar en el formulario»
if usuario["admin"]:
    menu["Administración"] = [
        _pagina(paginas.pagina_usuarios, "Usuarios", "Altas, contraseñas y permisos",
                "usuarios"),
        _pagina(paginas.pagina_credenciales, "Credenciales de Gemweb",
                "Acceso a la API de Gemweb para toda la plataforma", "credenciales"),
    ]

st.navigation(menu, expanded=True).run()
estilo.pie()
