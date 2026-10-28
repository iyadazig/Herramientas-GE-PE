# -*- coding: utf-8 -*-
"""
Imagen corporativa GE&PE para la app: colores, tipografia, logo y textos en
espanol de los elementos que Streamlit muestra en ingles.
"""

import base64

import streamlit as st

import almacen
import config

GRANATE = "#970000"          # rojo del logo
GRANATE_OSCURO = "#8B1A1A"   # plantilla del informe diario
NEGRO = "#1E1E1E"
GRIS = "#5A534C"
GRIS_CLARO = "#8A837C"
BEIS = "#FAF7F4"
BEIS_BORDE = "#E6E0DA"
ROSA = "#F3D9D9"
VERDE = "#2E7D32"
AMBAR = "#B26A00"

LOGO = config.RECURSOS / "assets" / "logo_geype.png"
LOGO_PEQ = config.RECURSOS / "assets" / "logo_geype_peq.png"
ICONO = config.RECURSOS / "assets" / "icono.ico"

COLOR_VEREDICTO = {"CORRECTO": VERDE, "FACTURADO DE MÁS": GRANATE,
                   "FACTURADO DE MENOS": AMBAR}

MODULOS_MENU = {"Seguimiento y Control": "seguimiento", "Revisión de SSAA": "ssaa",
                "API de Gemweb": "gemweb", "Solicitud de ofertas": "ofertas",
                "Estudio de potencia": "potencia",
                "Supervisión": "supervision", "Administración": "admin"}
LINEA_MENU = "#D9B8B8"        # granate muy claro: linea bajo cada modulo del menu


def _icono(trazos):
    """Icono de linea (24x24) en granate como imagen SVG en base64 (la forma que entienden
    todos los navegadores; si no se cargara, no se ve nada en vez de un recuadro)."""
    svg = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' "
           "stroke='%s' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>"
           % GRANATE + trazos + "</svg>")
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()


# icono de cada modulo del menu. El script del navegador marca cada titulo de grupo de
# inicio.py con su clave (MODULOS_MENU) segun su texto, no su posicion: la estructura del
# menu cambia entre versiones de Streamlit. Sin marca no hay icono (nunca un recuadro).
ICONOS_MODULO = {
    "seguimiento": _icono("<rect x='3' y='5' width='18' height='16' rx='2'/>"
                          "<path d='M3 10h18M8 3v4M16 3v4M9 15l2 2 4-4'/>"),
    "ssaa": _icono("<path d='M5 21V3h14v18l-2.5-1.5L14 21l-2-1.5L10 21l-2.5-1.5z'/>"
                   "<path d='M9 8h6M9 12h6M9 16h3'/>"),
    "gemweb": _icono("<ellipse cx='12' cy='6' rx='7' ry='3'/>"
                     "<path d='M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6'/>"
                     "<path d='M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6'/>"),
    "ofertas": _icono("<path d='M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z'/>"
                      "<path d='M14 3v5h5M9 13h6M9 17h6'/>"),
    "potencia": _icono("<path d='M13 2L4 14h7l-1 8 9-12h-7z'/>"),
    "supervision": _icono("<path d='M3 3v18h18'/><path d='M8 17v-5M12 17V8M16 17v-8M20 17v-3'/>"),
    "admin": _icono("<path d='M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z'/>"
                    "<path d='M9 12l2 2 4-4'/>"),
}

CSS = """
<style>
html, body, [class*="css"], .stMarkdown, .stText, button, input, textarea, select {
    font-family: Arial, Helvetica, sans-serif !important;
}
/* barra superior de Streamlit: sin texto en ingles */
[data-testid="stStatusWidget"], [data-testid="stDecoration"], [data-testid="stMainMenu"],
[data-testid="stToolbar"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 1.2rem; max-width: 1400px; }

/* cabecera corporativa */
.gp-cabecera {
    display: flex; align-items: center; gap: 1.6rem;
    border-bottom: 4px solid %(granate)s; padding: 0.4rem 0 1rem 0; margin-bottom: 1.4rem;
}
.gp-cabecera img { height: 74px; }
.gp-cabecera h1 {
    font-size: 1.75rem; font-weight: 700; color: %(negro)s; margin: 0; padding: 0;
    letter-spacing: 0.01em;
}
.gp-cabecera p { color: %(gris)s; margin: 0.25rem 0 0 0; font-size: 0.95rem; }

/* titulos de seccion numerados */
.gp-seccion {
    display: flex; align-items: center; gap: 0.7rem;
    margin: 1.8rem 0 0.8rem 0; padding-bottom: 0.45rem; border-bottom: 1px solid %(borde)s;
}
.gp-seccion .num {
    background: %(granate)s; color: #fff; font-weight: 700; border-radius: 50%%;
    width: 1.9rem; height: 1.9rem; display: flex; align-items: center; justify-content: center;
    font-size: 1rem; flex: none;
}
.gp-seccion .tit { font-size: 1.3rem; font-weight: 700; color: %(negro)s; }
.gp-seccion .sub { color: %(gris_claro)s; font-size: 0.9rem; margin-left: auto; }

/* subtitulos */
h3, .stMarkdown h3 {
    color: %(granate_oscuro)s !important; font-weight: 700 !important; font-size: 1.12rem !important;
    padding-top: 0.6rem !important;
}

/* tarjetas de metricas */
[data-testid="stMetric"] {
    background: %(beis)s; border: 1px solid %(borde)s; border-left: 4px solid %(granate)s;
    border-radius: 6px; padding: 0.7rem 0.9rem;
}
[data-testid="stMetricLabel"] { color: %(gris)s; }
[data-testid="stMetricValue"] { color: %(negro)s; font-weight: 700; }

/* veredicto */
.gp-veredicto {
    display: inline-block; padding: 0.35rem 0.9rem; border-radius: 4px; color: #fff;
    font-weight: 700; letter-spacing: 0.03em; font-size: 0.95rem;
}

/* botones */
.stButton > button, .stDownloadButton > button {
    border-radius: 4px; font-weight: 600; border: 1px solid %(granate)s; color: %(granate)s;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background: %(rosa)s; color: %(granate_oscuro)s; border-color: %(granate_oscuro)s;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: %(granate)s; color: #fff;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
    background: %(granate_oscuro)s; color: #fff;
}

/* barra lateral */
[data-testid="stSidebar"] { background: %(beis)s; border-right: 1px solid %(borde)s; }
[data-testid="stSidebar"] h2 { color: %(granate_oscuro)s; font-size: 1.05rem; }

/* menu: cada modulo con titulo granate, icono y linea debajo; los submodulos sangrados */
[data-testid="stNavSectionHeader"] {
    display: flex; align-items: center; gap: 0.5rem; color: %(granate_oscuro)s;
    font-size: 0.95rem; font-weight: 600; letter-spacing: 0; text-transform: none;
    margin: 1.1rem 1rem 0.3rem 1rem; padding: 0 0.25rem 0.35rem 0.25rem;
    border-bottom: 1px solid %(linea_menu)s;
}
[data-testid="stSidebarNavItems"] > [data-testid="stNavSectionHeader"]:first-child {
    margin-top: 0.3rem;
}
[data-testid="stNavSectionHeader"]::before {
    content: ""; flex: none; width: 18px; height: 18px;
    background: no-repeat center / contain;
}
[data-gp-modulo="seguimiento"]::before { background-image: url("%(ico_seguimiento)s"); }
[data-gp-modulo="ssaa"]::before { background-image: url("%(ico_ssaa)s"); }
[data-gp-modulo="gemweb"]::before { background-image: url("%(ico_gemweb)s"); }
[data-gp-modulo="ofertas"]::before { background-image: url("%(ico_ofertas)s"); }
[data-gp-modulo="potencia"]::before { background-image: url("%(ico_potencia)s"); }
[data-gp-modulo="supervision"]::before { background-image: url("%(ico_supervision)s"); }
[data-gp-modulo="admin"]::before { background-image: url("%(ico_admin)s"); }
/* flecha para recoger o desplegar los submodulos (traducir() pone data-gp-plegado y oculta
   los enlaces del modulo con data-gp-oculto) */
[data-gp-modulo] { cursor: pointer; user-select: none; }
[data-gp-modulo]::after {
    content: ""; flex: none; width: 7px; height: 7px; margin: -4px 4px 0 auto;
    border-right: 2px solid %(granate_oscuro)s; border-bottom: 2px solid %(granate_oscuro)s;
    transform: rotate(45deg); transition: transform 0.15s;
}
[data-gp-modulo][data-gp-plegado]::after { transform: rotate(-45deg); margin-top: 0; }
[data-gp-oculto] { display: none !important; }
/* «Mi trabajo» se abre con el boton de la barra lateral, no desde el menu */
[data-testid="stSidebarNavItems"] li:has(a[href$="/mi-trabajo"]) { display: none; }
[data-testid="stSidebarNavLink"] { padding-left: 2.1rem; margin-left: 0.5rem; }
[data-testid="stSidebarNavLink"] span { color: %(gris)s; font-size: 0.9rem; }
[data-testid="stSidebarNavLink"]:hover { background: %(rosa)s; }
[data-testid="stSidebarNavLink"][data-gp-activa] { background: %(rosa)s; }
[data-testid="stSidebarNavLink"][data-gp-activa] span {
    color: %(granate_oscuro)s; font-weight: 600;
}
/* paginas que cuelgan de «Cuadro de seguimiento»: sangradas bajo ella, con una linea */
[data-testid="stSidebarNavLink"][href$="/seguimiento-contactos"],
[data-testid="stSidebarNavLink"][href$="/seguimiento-claves"],
[data-testid="stSidebarNavLink"][href$="/seguimiento-smac"],
[data-testid="stSidebarNavLink"][href$="/seguimiento-telemedida"],
[data-testid="stSidebarNavLink"][href$="/seguimiento-bajas"] {
    padding-left: 1rem; margin-left: 2.6rem; border-left: 2px solid %(linea_menu)s;
    border-radius: 0 6px 6px 0;
}

/* expanders */
[data-testid="stExpander"] details { border: 1px solid %(borde)s; border-radius: 6px; }
[data-testid="stExpander"] summary:hover { color: %(granate)s; }

/* subida de ficheros: textos en espanol */
[data-testid="stFileUploaderDropzoneInstructions"] > div > span { display: none; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::before {
    content: "Arrastra el archivo aquí"; display: block; color: %(negro)s;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > small { display: none; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::after {
    content: "Hasta 200 MB por archivo"; display: block; font-size: 0.8rem; color: %(gris_claro)s;
}
[data-testid="stFileUploaderDropzone"] button { font-size: 0 !important; }
[data-testid="stFileUploaderDropzone"] button::after {
    content: "Buscar archivo"; font-size: 0.9rem;
}
[data-testid="stFileUploaderDropzone"] { background: %(beis)s; border: 1px dashed %(borde)s; }

/* pie */
.gp-pie {
    margin-top: 3rem; padding-top: 0.8rem; border-top: 1px solid %(borde)s;
    color: %(gris_claro)s; font-size: 0.8rem; text-align: center;
}
</style>
""" % {"granate": GRANATE, "granate_oscuro": GRANATE_OSCURO, "negro": NEGRO, "gris": GRIS,
       "gris_claro": GRIS_CLARO, "beis": BEIS, "borde": BEIS_BORDE, "rosa": ROSA,
       "linea_menu": LINEA_MENU, **{"ico_" + k: v for k, v in ICONOS_MODULO.items()}}


# Textos fijos de los componentes de Streamlit (en ingles y sin opcion de idioma).
# Solo se sustituyen nodos cuyo texto coincide EXACTAMENTE, para no tocar datos.
TRADUCCIONES = {
    "Browse files": "Buscar archivo", "Drag and drop file here": "Arrastra el archivo aquí",
    "Add row": "Añadir fila", "Delete row": "Borrar fila", "Delete rows": "Borrar filas",
    "Search": "Buscar", "Download as CSV": "Descargar CSV", "Fullscreen": "Pantalla completa",
    "Close fullscreen": "Salir de pantalla completa", "Show/hide columns": "Mostrar u ocultar columnas",
    "No results": "Sin resultados", "Choose an option": "Elige una opción",
    "Choose options": "Elige opciones", "Clear all": "Borrar todo",
    "No options to select.": "No hay opciones.", "Remove": "Quitar",
    "Sort ascending": "Orden ascendente", "Sort descending": "Orden descendente",
    "Format": "Formato", "Autosize": "Ajustar ancho", "Pin column": "Fijar columna",
    "Unpin column": "Soltar columna", "Hide column": "Ocultar columna",
    "Press the down arrow key to interact with the calendar and select a date. Press the escape "
    "button to close the calendar.": "Pulsa la flecha abajo para abrir el calendario y elegir una "
    "fecha. Pulsa Escape para cerrarlo.",
    "Select a date.": "Elige una fecha.", "Running...": "Calculando…",
    "Save as SVG": "Guardar como SVG", "Save as PNG": "Guardar como PNG",
    "Show password text": "Mostrar contraseña", "Hide password text": "Ocultar contraseña",
    "Press Enter to submit form": "Pulsa Intro para enviar",
    "Press Enter to apply": "Pulsa Intro para aplicar",
    "View more": "Ver más", "View less": "Ver menos",
    "Click to view actions": "Opciones de la gráfica",
    "Mo": "Lu", "Tu": "Ma", "We": "Mi", "Th": "Ju", "Fr": "Vi", "Sa": "Sá", "Su": "Do",
}
# opciones del menu de las graficas que se ocultan: son de programador y "Open in Vega
# Editor" enviaria los datos de la grafica (de clientes) a una web externa
OCULTAR = ["View Source", "View Compiled Vega", "Open in Vega Editor"]
MESES_EN = ["January", "February", "March", "April", "May", "June", "July", "August",
            "September", "October", "November", "December"]
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
            "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def traducir():
    """Traduce en el navegador los textos fijos en ingles de los widgets (calendario,
    subida de ficheros, tabla editable, desplegables) y el formato de fecha."""
    import json
    import streamlit.components.v1 as components
    components.html("""<script>
    const T = %s, EN = %s, ES = %s, OCULTAR = %s, MODULOS = %s;
    const doc = window.parent.document;
    // estado de cada modulo del menu en este navegador: {modulo: true si esta recogido}. Lo que
    // no se ha tocado sale recogido, salvo el modulo de la pagina abierta
    function estado() {
        try { return JSON.parse(window.parent.localStorage.getItem("gpMenuEstado")) || {}; }
        catch (e) { return window.parent.__gpMenuEstado || {}; }
    }
    // clic en el titulo de un modulo: lo recoge o despliega y se recuerda en este navegador.
    // Lo que registra este iframe deja de funcionar cuando Streamlit lo sustituye en la
    // siguiente ejecucion, asi que el manejador se mete como <script> en la propia pagina,
    // una sola vez. Escucha en captura, antes de que nadie pueda cortar el clic.
    if (!window.parent.__gpPlegar) {
        window.parent.__gpPlegar = true;
        const s = doc.createElement("script");
        s.textContent = `document.addEventListener("click", e => {
            const h = e.target.closest("[data-gp-modulo]");
            if (!h) return;
            let p;
            try { p = JSON.parse(localStorage.getItem("gpMenuEstado")) || {}; }
            catch (x) { p = window.__gpMenuEstado || {}; }
            p[h.dataset.gpModulo] = !h.hasAttribute("data-gp-plegado");
            try { localStorage.setItem("gpMenuEstado", JSON.stringify(p)); }
            catch (x) { /* sin almacenamiento: no se recuerda, pero se pliega igual */ }
            window.__gpMenuEstado = p;
            if (window.__gpMenu) window.__gpMenu();
        }, true);`;
        doc.head.appendChild(s);
    }
    function menu() {
        const aqui = window.parent.location.pathname.replace(/\\/$/, "");
        doc.querySelectorAll('[data-testid="stSidebarNavLink"]').forEach(a => {
            const activa = new URL(a.href, window.parent.location.href).pathname
                .replace(/\\/$/, "") === aqui;
            if (activa !== a.hasAttribute("data-gp-activa"))
                activa ? a.setAttribute("data-gp-activa", "") : a.removeAttribute("data-gp-activa");
        });
        const guardado = estado();
        // el titulo puede traer mas texto segun la version de Streamlit: se busca el nombre
        doc.querySelectorAll('[data-testid="stNavSectionHeader"]').forEach(h => {
            const texto = h.textContent.trim();
            const nombre = Object.keys(MODULOS).find(n => texto.startsWith(n));
            if (nombre && h.dataset.gpModulo !== MODULOS[nombre]) h.dataset.gpModulo = MODULOS[nombre];
            if (!h.dataset.gpModulo) return;
            const enlaces = [];
            for (let s = h.nextElementSibling; s && s.dataset.testid !== "stNavSectionHeader";
                 s = s.nextElementSibling) enlaces.push(s);
            // «Mi trabajo» (enlace oculto tras el ultimo modulo) no abre ningun modulo
            const abierta = enlaces.some(s => s.querySelector("[data-gp-activa]")
                                         && !s.querySelector('a[href$="/mi-trabajo"]'));
            const m = h.dataset.gpModulo;
            const plegado = m in guardado ? guardado[m] : !abierta;
            if (plegado !== h.hasAttribute("data-gp-plegado"))
                plegado ? h.setAttribute("data-gp-plegado", "") : h.removeAttribute("data-gp-plegado");
            enlaces.forEach(s => {
                if (plegado !== s.hasAttribute("data-gp-oculto"))
                    plegado ? s.setAttribute("data-gp-oculto", "") : s.removeAttribute("data-gp-oculto");
            });
        });
    }
    // la pagina ya esta en espanol: que el navegador no la traduzca (altera contrasenas y datos)
    doc.documentElement.lang = "es";
    doc.documentElement.setAttribute("translate", "no");
    doc.documentElement.classList.add("notranslate");
    if (!doc.querySelector('meta[name="google"]')) {
        const m = doc.createElement("meta");
        m.name = "google"; m.content = "notranslate"; doc.head.appendChild(m);
    }
    const reMes = new RegExp("^(" + EN.join("|") + ")( \\\\d{4})?$");
    function nodo(n) {
        const t = n.nodeValue.trim();
        if (!t) return;
        if (OCULTAR.includes(t) && n.parentElement) {
            const e = n.parentElement.closest("a") || n.parentElement;
            e.style.display = "none"; return;
        }
        if (T[t] !== undefined) { n.nodeValue = n.nodeValue.replace(t, T[t]); return; }
        const m = t.match(reMes);
        if (m) n.nodeValue = n.nodeValue.replace(m[1], ES[EN.indexOf(m[1])]);
    }
    function pasar(raiz) {
        const w = doc.createTreeWalker(raiz, NodeFilter.SHOW_TEXT);
        let n; while ((n = w.nextNode())) nodo(n);
        raiz.querySelectorAll && raiz.querySelectorAll("input[placeholder], [aria-label]").forEach(e => {
            if (e.placeholder === "DD/MM/YYYY") e.placeholder = "DD/MM/AAAA";
            const a = e.getAttribute("aria-label");
            if (a && T[a] !== undefined) e.setAttribute("aria-label", T[a]);
        });
    }
    pasar(doc.body);
    menu();
    // al pulsar en el menu, la pagina abierta se marca enseguida (sin esperar a la recarga);
    // se llama a la funcion de la ultima ejecucion (el iframe de las anteriores puede no estar)
    window.parent.__gpMenu = menu;
    if (!window.parent.__gpMenuClic) {
        window.parent.__gpMenuClic = true;
        doc.addEventListener("click", e => {
            if (e.target.closest('[data-testid="stSidebarNavLink"]'))
                setTimeout(() => window.parent.__gpMenu(), 50);
        });
    }
    // en cada ejecucion se sustituye el observador anterior (su iframe puede haberse borrado)
    if (window.parent.__gpTraductor) window.parent.__gpTraductor.disconnect();
    window.parent.__gpTraductor = new MutationObserver(ms => {
        ms.forEach(m => {
            m.addedNodes.forEach(n => n.nodeType === 3 ? nodo(n) : (n.nodeType === 1 && pasar(n)));
            if (m.type === "characterData") nodo(m.target);
        });
        menu();
    });
    window.parent.__gpTraductor.observe(doc.body, {childList: true, subtree: true,
                                                   characterData: true});
    </script>""" % (json.dumps(TRADUCCIONES, ensure_ascii=False), json.dumps(MESES_EN),
                    json.dumps(MESES_ES, ensure_ascii=False), json.dumps(OCULTAR),
                    json.dumps(MODULOS_MENU, ensure_ascii=False)), height=0)


def _b64(ruta):
    return base64.b64encode(ruta.read_bytes()).decode("ascii")


def aplicar():
    """CSS corporativo. Llamar justo tras set_page_config."""
    st.markdown(CSS, unsafe_allow_html=True)


def logo_menu():
    """Logo pequeno arriba a la izquierda, sobre el menu (solo con la sesion iniciada)."""
    if LOGO_PEQ.exists():
        st.logo(str(LOGO_PEQ), size="large")


def cabecera(titulo, subtitulo=""):
    logo = '<img src="data:image/png;base64,%s" alt="GE&PE">' % _b64(LOGO_PEQ) \
        if LOGO_PEQ.exists() else ""
    st.markdown('<div class="gp-cabecera">%s<div><h1>%s</h1><p>%s</p></div></div>'
                % (logo, titulo, subtitulo), unsafe_allow_html=True)


def _usuario():
    return (st.session_state.get("usuario") or {}).get("usuario")


def registrar_pagina(titulo, modulo):
    """Anota en la actividad que el usuario ha abierto una pagina: una vez por sesion y pagina,
    no en cada recarga."""
    vistas = st.session_state.setdefault("_paginas_registradas", set())
    if titulo not in vistas:
        vistas.add(titulo)
        almacen.registrar_actividad(_usuario(), modulo, "Abrir página", titulo)


def boton_descarga(modulo, contenedor, etiqueta, data, file_name, mime=None, **kw):
    """contenedor.download_button (st o una columna) que anota la descarga en la actividad,
    para el panel de supervision. Los demas argumentos, como los de st.download_button."""
    return contenedor.download_button(
        etiqueta, data, file_name, mime, on_click=almacen.registrar_actividad,
        args=(_usuario(), modulo, "Descarga", file_name), **kw)


def borrar_mios(modulo, filas, describir, borrar, clave, que="registros"):
    """Desplegable para que el usuario conectado borre lo que ha creado el mismo (nunca lo de
    otros). filas: dicts con id y usuario (los que se ven con los filtros); describir(fila):
    texto de cada uno; borrar(id, usuario): True si lo borra (almacen.borrar_...)."""
    usuario = _usuario()
    mios = {f["id"]: f for f in filas if f.get("usuario") == usuario}
    hecho = st.session_state.pop(clave + "_hecho", None)
    if hecho:
        st.success(hecho)
    with st.expander("Borrar %s creados por mí" % que, icon=":material/delete:"):
        if not mios:
            st.caption("Con estos filtros no hay %s creados por ti. Solo puedes borrar los "
                       "tuyos." % que)
            return
        elegidos = st.multiselect("Elige los que quieres borrar", list(mios),
                                  format_func=lambda i: describir(mios[i]), key=clave + "_sel")
        confirmo = st.checkbox("Confirmo que quiero borrarlos: no se puede deshacer",
                               key=clave + "_ok")
        if st.button("Borrar %d" % len(elegidos), icon=":material/delete:", key=clave + "_btn",
                     disabled=not (elegidos and confirmo)):
            n = sum(bool(borrar(i, usuario)) for i in elegidos)
            almacen.registrar_actividad(usuario, modulo, "Borrar",
                                        "%d %s: %s" % (n, que, ", ".join(map(str, elegidos))))
            st.session_state[clave + "_hecho"] = "Borrados %d %s." % (n, que)
            for k in (clave + "_sel", clave + "_ok"):
                st.session_state.pop(k, None)
            st.rerun()


def seccion(numero, titulo, subtitulo=""):
    st.markdown('<div class="gp-seccion"><span class="num">%s</span><span class="tit">%s</span>'
                '<span class="sub">%s</span></div>' % (numero, titulo, subtitulo),
                unsafe_allow_html=True)


def veredicto(texto):
    color = COLOR_VEREDICTO.get(texto, GRIS_CLARO)
    return '<span class="gp-veredicto" style="background:%s">%s</span>' % (color, texto)


def pie():
    traducir()
    st.markdown('<div class="gp-pie">Herramientas GE&amp;PE · Ingeniería y Gestión Energética. '
                'Uso interno: los datos se procesan en los equipos de GE&amp;PE.</div>',
                unsafe_allow_html=True)
