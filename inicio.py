# -*- coding: utf-8 -*-
"""
HERRAMIENTAS GE&PE — entrada unica de la plataforma
===================================================
    python lanzador.py --servidor          (PC de la oficina, para todo el equipo)
    streamlit run inicio.py                (pruebas en local)

Acceso con usuario y contrasena (paginas.acceso), imagen corporativa (estilo) y menu por
secciones (st.navigation). Cada pagina es un fichero o una funcion; las de Gemweb estan
en gemweb_extractor/paginas.py y las de la
solicitud de ofertas en ofertas/paginas.py y las del analisis de ofertas en
analisis_ofertas/paginas.py.
"""

import functools

import streamlit as st

import almacen
import estilo
import paginas
from gemweb_extractor import paginas as gw
from analisis_ofertas import paginas as analisis
from analisis_ofertas import paginas_regulados as regulados
from estudio_potencia import paginas as potencia
from ofertas import paginas as ofertas
from mi_espacio import paginas as mi_espacio
from seguimiento import envios
from seguimiento import paginas as seguimiento
from supervision import paginas as supervision

st.set_page_config(page_title="Herramientas GE&PE",
                   page_icon=str(estilo.ICONO) if estilo.ICONO.exists() else None,
                   layout="wide", menu_items={})
estilo.aplicar()
usuario = paginas.acceso()          # detiene la pagina hasta que se inicia sesion
estilo.logo_menu()                  # no sale en la pagina de inicio de sesion


def _pagina(funcion, titulo, subtitulo, ruta, modulo, default=False):
    """Pagina de funcion con la cabecera corporativa comun. Si la API de Gemweb tiene un fallo
    conocido en esa opcion, se marca en el menu y se avisa al abrirla. Al abrirla queda
    anotado en la actividad (panel de supervision)."""
    aviso = gw.AVISOS_GEMWEB.get(titulo)

    @functools.wraps(funcion)
    def mostrar():
        c1, c2 = st.columns([12, 1], vertical_alignment="center")
        with c1:
            estilo.cabecera(titulo.split(" · ")[0], subtitulo)
        _estrella(c2, ruta)
        estilo.registrar_pagina(titulo.split(" · ")[0], modulo)
        if aviso:
            st.warning("Fallo de la API de Gemweb: " + aviso, icon=":material/warning:")
        funcion()
    return st.Page(mostrar, title=titulo, url_path=ruta, default=default,
                   icon=":material/warning:" if aviso else None)


def _estrella(contenedor, ruta):
    """☆/★ para añadir o quitar la página de los favoritos del usuario."""
    favs = almacen.leer_preferencia(usuario["usuario"], "favoritos", []) or []
    es = ruta in favs
    if contenedor.button("★" if es else "☆", key="fav_" + ruta,
                         help="Quitar de mis favoritos" if es else "Añadir a mis favoritos"):
        favs = [r for r in favs if r != ruta] if es else favs + [ruta]
        almacen.guardar_preferencia(usuario["usuario"], "favoritos", favs)
        st.rerun()


SUB_SSAA = "Comprobación del concepto de SSAA con los datos publicados por REE (ESIOS)"
SUB_GEMWEB = "Consultas a la API de Gemweb"
MI_ESPACIO = "Mi espacio"
SEGUIMIENTO, SSAA, GEMWEB, OFERTAS, POTENCIA, SUPERVISION, ADMIN = (
    "Seguimiento y Control", "Revisión de SSAA", "API de Gemweb", "Ofertas",
    "Estudio de potencia", "Supervisión", "Administración")
# modulos de la actividad (supervision) de las paginas de «Ofertas»
SOLICITUD, ANALISIS = "Solicitud de ofertas", "Análisis de ofertas"
SUB_ANALISIS = "Valoración de las ofertas recibidas frente al contrato actual"
SUB_SGC = "Seguimiento de facturas e informes de los clientes"
_avisos = mi_espacio.contador()
menu = {
    MI_ESPACIO: [
        _pagina(mi_espacio.pagina_mi_panel, "Mi panel" + (f" · {_avisos}" if _avisos else ""),
                "Lo tuyo de un vistazo: lo urgente, lo pendiente y tus atajos", "mi-panel",
                MI_ESPACIO, default=True),
        _pagina(mi_espacio.pagina_licitaciones, "Próximas licitaciones",
                "Contratos por fecha de lanzamiento, denuncia o fin, con su estado y la carga "
                "de trabajo por meses", "licitaciones", MI_ESPACIO),
        _pagina(mi_espacio.pagina_ficha, "Ficha de cliente",
                "Todo lo de un cliente o CUPS en la aplicación", "ficha-cliente", MI_ESPACIO),
        _pagina(mi_espacio.pagina_calendario, "Calendario de vencimientos",
                "Denuncias, fines de contrato, plazos de ofertas, catálogo e informes",
                "calendario", MI_ESPACIO),
    ],
    SEGUIMIENTO: [
        _pagina(seguimiento.pagina_cuadro, "Cuadro de seguimiento", SUB_SGC, "seguimiento",
                SEGUIMIENTO),
    ] + [
        _pagina(funcion, titulo, subtitulo, ruta, SEGUIMIENTO)
        for funcion, titulo, subtitulo, ruta in (
            # cuelgan de «Cuadro de seguimiento» (sangradas en el menu, estilo.GRUPOS_MENU)
            (seguimiento.pagina_contactos, "Contactos de clientes", SUB_SGC,
             "seguimiento-contactos"),
            (seguimiento.pagina_claves, "Claves de clientes", SUB_SGC, "seguimiento-claves"),
            (seguimiento.pagina_smac, "SMAC", SUB_SGC, "seguimiento-smac"),
            (seguimiento.pagina_telemedida, "Telemedida", SUB_SGC, "seguimiento-telemedida"),
            (seguimiento.pagina_bajas, "Bajas", "Suministros dados de baja en el seguimiento",
             "seguimiento-bajas"),
            # cuelgan de «Enviar Informes Sgto.» las tres siguientes (estilo.GRUPOS_MENU)
            (envios.pagina_informes, "Enviar Informes Sgto.",
             "Informes de seguimiento por correo desde tu cuenta de Outlook",
             "seguimiento-informes"),
            (envios.pagina_incidencias, "Avisos de incidencias",
             "Avisos de incidencias en facturas por correo desde tu cuenta de Outlook",
             "seguimiento-incidencias"),
            (envios.pagina_historial, "Historial de envíos",
             "Informes y avisos enviados por el equipo", "seguimiento-envios"),
            (seguimiento.pagina_contactos_envio, "Contactos de envío",
             "Destinatarios de los informes y de los avisos de incidencias",
             "seguimiento-contactos-envio"),
            (seguimiento.pagina_retraso, "Retraso de informes",
             "Cumplimiento del envío de los informes de seguimiento", "seguimiento-retraso"))
    ],
    SSAA: [
        st.Page("app_revision_ssaa.py", title="Revisar factura", url_path="revisar"),
        _pagina(paginas.pagina_historial, "Historial de revisiones", SUB_SSAA, "historial",
                SSAA),
        _pagina(paginas.pagina_fichas, "Fichas de contrato", SUB_SSAA, "fichas", SSAA),
    ],
    GEMWEB: [
        # las que tienen un fallo conocido de la API (con su triangulo) van al final
        _pagina(funcion, titulo, SUB_GEMWEB, ruta, GEMWEB) for funcion, titulo, ruta in sorted((
            (gw.pagina_inventario, "Inventario", "gemweb-inventario"),
            (gw.pagina_consumo_periodo, "Consumo por periodo", "gemweb-consumo-periodo"),
            (gw.pagina_telelecturas, "Telelecturas", "gemweb-telelecturas"),
            (gw.pagina_descarga_masiva, "Descarga masiva por CUPS", "gemweb-descarga-masiva"),
            (gw.pagina_coste_consumo, "Coste y consumo mensual", "gemweb-coste-consumo"),
            (gw.pagina_optimizacion, "Optimización de potencia", "gemweb-optimizacion"),
            (gw.pagina_incoherencias, "Incoherencias en facturas", "gemweb-incoherencias"),
            (gw.pagina_factura_calculada, "Factura calculada", "gemweb-factura-calculada"),
            (gw.pagina_simulacion_contrato, "Simulación de contrato",
             "gemweb-simulacion-contrato")), key=lambda p: p[1] in gw.AVISOS_GEMWEB)
    ],
    OFERTAS: [
        _pagina(ofertas.pagina_solicitud_ofertas, "Crear nueva solicitud",
                "Excel de solicitud de ofertas con los datos de Gemweb", "solicitud-ofertas",
                SOLICITUD),
        _pagina(ofertas.pagina_solicitudes_guardadas, "Solicitudes guardadas",
                "Solicitudes de ofertas guardadas por el equipo y su seguimiento",
                "solicitudes-guardadas", SOLICITUD),
        _pagina(analisis.pagina_nuevo_analisis, "Nuevo análisis", SUB_ANALISIS,
                "analisis-ofertas", ANALISIS),
        _pagina(analisis.pagina_analisis_guardados, "Análisis guardados",
                "Análisis de ofertas guardados por el equipo", "analisis-guardados", ANALISIS),
        _pagina(analisis.pagina_catalogo, "Catálogo de ofertas",
                "Ofertas de catálogo del equipo: búscalas, revísalas y cárgalas en un análisis",
                "catalogo-ofertas", ANALISIS),
        _pagina(analisis.pagina_parametros, "Parámetros utilizados",
                "Términos regulados, mercado y apuntamientos con su fuente y fecha",
                "parametros-utilizados", ANALISIS),
    ],
    POTENCIA: [
        _pagina(potencia.pagina_nuevo_estudio, "Nuevo estudio",
                "Optimización de la potencia contratada (6.1TD y 3.0TD) con la curva de carga",
                "estudio-potencia", POTENCIA),
        _pagina(potencia.pagina_estudios_guardados, "Estudios guardados",
                "Estudios de potencia guardados por el equipo", "estudios-potencia-guardados",
                POTENCIA),
    ],
}
ofertas.PAGINA_FORMULARIO = menu[OFERTAS][0]   # para «Cargar en el formulario»
analisis.PAGINA_FORMULARIO = menu[OFERTAS][2]  # para «Cargar en Nuevo análisis»
potencia.PAGINA_FORMULARIO = menu[POTENCIA][0]
if almacen.puede_supervisar(usuario):
    menu[SUPERVISION] = [
        _pagina(supervision.pagina_panel, "Panel del equipo",
                "Trabajo guardado y actividad de cada usuario", "supervision", SUPERVISION),
    ]
if usuario["admin"]:
    menu[ADMIN] = [
        _pagina(paginas.pagina_usuarios, "Usuarios", "Altas, contraseñas y permisos",
                "usuarios", ADMIN),
        _pagina(paginas.pagina_credenciales, "Credenciales de Gemweb",
                "Acceso a la API de Gemweb para toda la plataforma", "credenciales", ADMIN),
        _pagina(seguimiento.pagina_admin, "Seguimiento y Control",
                "Importar el cuadro, equipo, listas de valores y Outlook", "admin-seguimiento", ADMIN),
        _pagina(regulados.pagina_datos_regulados, "Datos regulados",
                "Peajes, cargos, impuestos y parámetros de mercado de cada año",
                "datos-regulados", ADMIN),
        _pagina(mi_espacio.seccion_ajustes, "Ajustes de Mi espacio",
                "Umbrales generales del semáforo y de los avisos", "admin-mi-espacio", ADMIN),
    ]
# «Mi trabajo»: el panel con lo del propio usuario. Va sin titulo de modulo y su enlace se
# oculta (estilo.CSS); se abre con el boton de la barra lateral
MI_TRABAJO = _pagina(supervision.pagina_mi_trabajo, "Mi trabajo",
                     "Tu trabajo guardado y tu actividad en la plataforma", "mi-trabajo",
                     supervision.MI_TRABAJO)
menu[""] = [MI_TRABAJO]
# páginas por ruta: para abrirlas desde «Mi espacio» y para los favoritos
mi_espacio.PAGINAS.clear()
mi_espacio.PAGINAS.update({p.url_path: p for paginas_ in menu.values() for p in paginas_})

with st.sidebar:
    _favs = [r for r in (almacen.leer_preferencia(usuario["usuario"], "favoritos", []) or [])
             if r in mi_espacio.PAGINAS]
    if _favs:
        st.caption("Mis favoritos")
        for _r in _favs:
            st.page_link(mi_espacio.PAGINAS[_r], icon=":material/star:")
    if st.button("Mi trabajo", icon=":material/person:", use_container_width=True,
                 help="Todo lo que has hecho en la plataforma: trabajo guardado y actividad"):
        st.switch_page(MI_TRABAJO)
    paginas.barra_usuario()
    if gw.AVISOS_GEMWEB:
        with st.expander(":material/warning: Fallos de la API de Gemweb"):
            st.caption("Estas opciones no funcionan del todo por falta de implementación o "
                       "mejora de la API de Gemweb (no de esta aplicación).")
            for _titulo, _texto in gw.AVISOS_GEMWEB.items():
                st.markdown("**%s**: %s" % (_titulo, _texto))

st.navigation(menu, expanded=True).run()
estilo.pie()
