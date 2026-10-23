# -*- coding: utf-8 -*-
"""
PAGINAS DE «ESTUDIO DE POTENCIA» de Herramientas GE&PE
======================================================
Optimizacion de la potencia contratada (6.1TD y 3.0TD > 50 kW) a partir de la curva de carga,
como el programa de escritorio EstudioPotencia: individual o multipunto, curva de ficheros o de
Gemweb, propuestas, anexos, informe PDF y estudios guardados para todo el equipo.

La zona (Peninsula, Baleares, Canarias, Ceuta o Melilla) no se elige: sale del codigo postal de
cada suministro en el inventario de Gemweb (se puede cambiar a mano). Asi un multipunto puede
mezclar suministros de zonas distintas. La logica esta en estudio_potencia/servicio.py.
"""

import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import almacen
import gemweb
from estudio_potencia import servicio as sv
from gemweb_extractor import paginas as gw
from potencia import graficos
from potencia.calculo import N_P
from potencia.calendario import ZONAS
from potencia.lector_curva import CONVENIOS, UNIDADES, LectorCurva

# la registra inicio.py: «Estudios guardados» vuelve a ella al cargar un estudio
PAGINA_FORMULARIO = None

INDIVIDUAL, MULTIPUNTO = "individual", "multipunto"
PERIODOS = [f"P{i}" for i in range(1, N_P + 1)]
ANIO = "Año completo"


# ------------------------------------------------------------------- estado
def _estado() -> dict:
    """Estado del estudio en la sesion (el lector de la curva, los suministros, el PDF...)."""
    return st.session_state.setdefault("pot", {
        "tipo": INDIVIDUAL, "lector": None, "ajustes": {}, "nombre": "", "suministros": [],
        "errores": [], "avisos_descarga": [], "calculado": False, "cargado": None,
        "pdf": None, "pdf_firma": None, "version": 0, "ediciones": {}, "pendiente": {}})


def _reiniciar(tipo=INDIVIDUAL):
    for k in [k for k in st.session_state if str(k).startswith("pot")]:
        del st.session_state[k]
    _estado()["tipo"] = tipo
    st.session_state.pot["pendiente"] = {"pot_tipo": tipo}


def _aplicar_pendiente(pot):
    """Valores para los controles que hay que cambiar desde el codigo (se aplican antes de
    dibujarlos: Streamlit no deja cambiar un control ya dibujado)."""
    for clave, valor in pot.pop("pendiente", {}).items():
        st.session_state[clave] = valor
    pot["pendiente"] = {}


def _edicion(pot, i):
    """Contador para volver a empezar las tablas de propuestas de un suministro."""
    return pot["ediciones"].get(i, 0)


# ------------------------------------------------------------------- Gemweb
def _cliente_gemweb():
    cred, _origen = gemweb.cargar_credenciales()
    return gemweb.ClienteGemweb(*cred) if cred else None


def _inventario(cups) -> tuple:
    """(suministro, centro de consumo, hay credenciales) de Gemweb, guardado en la sesion."""
    cache = st.session_state.setdefault("pot_inventario", {})
    if cups not in cache:
        cliente = _cliente_gemweb()
        sumi = centro = None
        if cliente is not None:
            try:
                sumi = cliente.buscar_suministro(cups)
                if sumi and not sumi.get("codi_postal") and sumi.get("id_centres_consum"):
                    centro = gw.registro_inventario(cliente.api, "centres_consum",
                                                    "centres_consum.id",
                                                    sumi["id_centres_consum"])
            except Exception:            # sin conexion o fallo de la API: se revisa a mano
                sumi = None
        cache[cups] = (sumi, centro, cliente is not None)
    return cache[cups]


def _zona_auto(s) -> tuple:
    """(zona, de donde sale) del suministro segun su codigo postal en Gemweb."""
    if not s.cups:
        return sv.PENINSULA, "sin CUPS: revisa la zona"
    sumi, centro, hay = _inventario(s.cups)
    if not hay:
        return sv.PENINSULA, "sin credenciales de Gemweb: revisa la zona"
    return sv.zona_de(sumi, centro)


# ------------------------------------------------------------------- carga de curvas
def _nuevos_suministros(pot, lector, ajustes, nombre, info=None, avisos=(), rellenar=False):
    """Un suministro por CUPS de la curva, con su zona de Gemweb, y la curva leida con ella."""
    info = info or {}
    multipunto = pot["tipo"] == MULTIPUNTO
    try:
        primera = sv.interpretar(lector, sv.PENINSULA, ajustes)     # detecta los CUPS
    except Exception as e:
        st.error(f"No se ha podido interpretar la curva: {e}")
        return
    if multipunto:
        cups_lista = list(lector.cups_disponibles)
        if not cups_lista:
            st.error("Para un estudio multipunto la curva tiene que traer el CUPS de cada "
                     "registro (una columna con el CUPS, o el CUPS en el nombre de cada fichero).")
            return
    else:
        cups_lista = [primera.cups or st.session_state.get("pot_cups", "").strip().upper()]
    suministros, errores = [], []
    for cups in cups_lista:
        s = sv.Suministro(cups=cups or "")
        sumi = info.get(cups)
        if sumi:
            s.avisos_extra = sv.avisos_gemweb(cups, list(avisos), sumi)
            if rellenar:
                sv.rellenar_desde_gemweb(s, sumi)
        s.zona, s.zona_origen = _zona_auto(s)
        try:
            sv.leer_suministro(s, lector, ajustes, multipunto)
            suministros.append(s)
        except Exception as e:
            errores.append(f"{cups}: {e}")
    if not suministros:
        st.error("No se ha podido interpretar la curva de ningún CUPS.\n\n" + "\n".join(errores))
        return
    pot.update(lector=lector, ajustes=ajustes, nombre=nombre, suministros=suministros,
               errores=errores, avisos_descarga=list(avisos), calculado=False, cargado=None,
               pdf=None, pdf_firma=None, version=pot["version"] + 1, ediciones={})
    if not multipunto:
        s = suministros[0]
        pendiente = {}
        if s.cups and not st.session_state.get("pot_cups", "").strip():
            pendiente["pot_cups"] = s.cups
        if rellenar and info:
            pendiente.update(pot_cups=s.cups, pot_instalacion=s.denominacion,
                             pot_direccion=s.direccion, pot_tarifa=s.tarifa)
        pot["pendiente"] = pendiente
    st.rerun()


def _cargar_ficheros(pot, subidos):
    carpeta = Path(tempfile.mkdtemp(prefix="potencia_"))
    rutas = sv.guardar_subidos([(f.name, f.getvalue()) for f in subidos], carpeta)
    try:
        lector = sv.abrir(rutas)
    except Exception as e:
        st.error(f"No se ha podido leer {'el fichero' if len(rutas) == 1 else 'los ficheros'}: {e}")
        return
    nombre = rutas[0].name if len(rutas) == 1 else f"{len(rutas)} ficheros de curva"
    _nuevos_suministros(pot, lector, {"hoja": lector.hoja_defecto, "unidad": UNIDADES[0],
                                      "convenio": CONVENIOS[0], "cups": None}, nombre)


def _descargar_gemweb(pot, lista, desde, hasta, rellenar):
    cliente = _cliente_gemweb()
    if cliente is None:
        st.error("No hay credenciales de Gemweb: un administrador las configura en "
                 "Administración → Credenciales de Gemweb.")
        return
    barra = st.progress(0.0, text="Buscando en el inventario de Gemweb…")
    try:
        ruta, info, avisos = sv.descargar_gemweb(
            cliente, lista, desde, hasta, tempfile.mkdtemp(prefix="potencia_"),
            al_avanzar=lambda fraccion, texto: barra.progress(min(fraccion, 1.0), text=texto))
    except gemweb.GemwebError as e:
        barra.empty()
        st.error(f"Error de Gemweb: {e}")
        return
    barra.empty()
    if ruta is None:
        st.warning("No se ha descargado la curva de ningún CUPS.\n\n"
                   + "\n".join(f"- {a}" for a in avisos[:10]))
        return
    cache = st.session_state.setdefault("pot_inventario", {})
    for cups, sumi in info.items():
        cache[cups] = (sumi, None, True)
    _nuevos_suministros(pot, LectorCurva(ruta),
                        {"hoja": None, "unidad": UNIDADES[0], "convenio": "Fin de periodo",
                         "cups": None},
                        f"Gemweb · descargada el {date.today():%d/%m/%Y}", info, avisos, rellenar)


def _releer(pot, s):
    """Vuelve a leer la curva del suministro (otra zona u otros ajustes de lectura)."""
    if pot["lector"] is None:
        s.invalidar()                    # estudio guardado: se recalcula con la curva guardada
        return
    try:
        sv.leer_suministro(s, pot["lector"], pot["ajustes"], pot["tipo"] == MULTIPUNTO)
    except Exception as e:
        st.error(f"{s.cups}: no se ha podido volver a leer la curva: {e}")


# ------------------------------------------------------------------- pasos del formulario
def _datos(pot) -> dict:
    st.subheader("1. Datos del estudio")
    c1, c2 = st.columns(2)
    st.session_state.setdefault("pot_fecha", date.today().strftime("%d/%m/%Y"))
    titular = c1.text_input("Titular", key="pot_titular")
    fecha = c2.text_input("Fecha", key="pot_fecha")
    datos = {"titular": titular.strip(), "fecha": fecha.strip()}
    if pot["tipo"] == INDIVIDUAL:
        c1, c2, c3 = st.columns([2, 2, 1])
        datos["cups"] = c1.text_input("CUPS", key="pot_cups").strip().upper()
        datos["instalacion"] = c2.text_input("Instalación", key="pot_instalacion").strip()
        datos["tarifa"] = c3.selectbox("Tarifa", sv.TARIFAS, key="pot_tarifa")
        datos["direccion"] = st.text_input("Dirección", key="pot_direccion").strip()
    return datos


def _curva(pot):
    st.subheader("2. Curva de carga")
    multipunto = pot["tipo"] == MULTIPUNTO
    t_fich, t_gem = st.tabs(["Ficheros", "Descargar de Gemweb"])
    with t_fich:
        st.caption("CSV, TXT o Excel con la curva cuartohoraria u horaria (kW o kWh). Se pueden "
                   "subir varios a la vez, p. ej. uno por mes"
                   + (" y por CUPS: el CUPS se toma de su columna o del nombre del fichero."
                      if multipunto else "."))
        subidos = st.file_uploader("Ficheros de curva", type=sv.EXTENSIONES,
                                   accept_multiple_files=True,
                                   key="pot_subidos_%d" % pot["version"])
        if st.button("Leer los ficheros", disabled=not subidos, type="primary"):
            _cargar_ficheros(pot, subidos)
    with t_gem:
        st.caption("Curva cuartohoraria de la telemedida en Gemweb. Por defecto, los 12 últimos "
                   "meses completos.")
        texto = st.text_area("CUPS" + (" (uno por línea)" if multipunto else ""),
                             key="pot_gw_cups", height=110 if multipunto else 68,
                             placeholder="ES0000000000000000XX")
        ini, fin = sv.periodo_defecto()
        c1, c2 = st.columns(2)
        desde = c1.date_input("Desde", value=ini, format="DD/MM/YYYY", key="pot_gw_desde")
        hasta = c2.date_input("Hasta", value=fin, format="DD/MM/YYYY", key="pot_gw_hasta")
        rellenar = st.checkbox("Rellenar tarifa, potencias contratadas, denominación y dirección "
                               "con los datos de Gemweb", value=True, key="pot_gw_rellenar")
        lista = sv.lista_cups(texto)
        if lista and not multipunto and len(lista) > 1:
            st.warning("En un estudio individual se descarga un solo CUPS.")
        if st.button("Descargar de Gemweb", disabled=not lista or hasta <= desde or (
                not multipunto and len(lista) > 1)):
            _descargar_gemweb(pot, lista, desde, hasta, rellenar)


def _interpretacion(pot):
    """Lo que se ha entendido de la curva y los ajustes para corregirlo."""
    lector, sums = pot["lector"], pot["suministros"]
    multipunto = pot["tipo"] == MULTIPUNTO
    with st.expander("Interpretación de la curva de carga", expanded=not pot["calculado"]):
        st.caption(f"Origen: {pot['nombre']}")
        if lector is not None:
            a = dict(pot["ajustes"])
            c1, c2, c3, c4 = st.columns(4)
            v = pot["version"]
            if len(lector.hojas) > 1:
                hojas = [str(h) for h in lector.hojas]
                actual = str(a.get("hoja") or lector.hoja_defecto)
                a["hoja"] = c1.selectbox("Hoja", hojas, key="pot_hoja_%d" % v,
                                         index=hojas.index(actual) if actual in hojas else 0)
            a["unidad"] = c2.selectbox("Unidad de los valores", UNIDADES, key="pot_unidad_%d" % v,
                                       index=UNIDADES.index(a.get("unidad") or UNIDADES[0]))
            a["convenio"] = c3.selectbox("La hora de cada registro marca", CONVENIOS,
                                         key="pot_convenio_%d" % v,
                                         index=CONVENIOS.index(a.get("convenio") or CONVENIOS[0]))
            if not multipunto and len(lector.cups_disponibles) > 1:
                opciones = list(lector.cups_disponibles)
                elegido = sums[0].curva.cups if sums[0].curva.cups in opciones else opciones[0]
                a["cups"] = c4.selectbox("CUPS", opciones, key="pot_cups_curva_%d" % v,
                                         index=opciones.index(a.get("cups") or elegido))
            if a != pot["ajustes"]:
                pot["ajustes"] = a
                for s in sums:
                    if not multipunto and a.get("cups"):
                        s.cups = a["cups"]
                    _releer(pot, s)
                pot["pdf"] = None
        primera = sums[0].curva
        st.markdown("\n".join(f"- {l}" for l in sv.resumen_interpretacion(primera, lector,
                                                                            multipunto)))
        for s in sums:
            for aviso in s.curva.avisos:
                st.warning((f"{s.cups}: " if multipunto else "") + aviso)
        for e in pot["errores"]:
            st.error(e)
        st.caption("Primeros registros interpretados" + (" (del primer CUPS)" if multipunto else ""))
        st.dataframe(sv.primeros_registros(primera), hide_index=True, use_container_width=True)


def _zona_individual(pot, s):
    """Zona del suministro individual: la de Gemweb (segun el CUPS) salvo que se cambie."""
    if not s.zona_manual:
        zona, origen = _zona_auto(s)
        if (zona, origen) != (s.zona, s.zona_origen):
            s.zona, s.zona_origen = zona, origen
            if s.zona_curva != zona:
                _releer(pot, s)
    c1, c2 = st.columns([1, 3])
    elegida = c1.selectbox("Zona", ZONAS, index=ZONAS.index(s.zona),
                           key="pot_zona_%d_%s" % (pot["version"], s.zona_origen))
    c2.caption(f"Según el código postal del suministro: {s.zona_origen}" if not s.zona_manual
               else "Zona elegida a mano.")
    if elegida != s.zona:
        s.zona, s.zona_manual, s.zona_origen = elegida, True, "elegida a mano"
        _releer(pot, s)
        st.rerun()


def _suministros(pot, datos):
    st.subheader("3. Suministros y potencias contratadas actuales")
    sums = pot["suministros"]
    v = pot["version"]
    if pot["tipo"] == INDIVIDUAL:
        s = sums[0]
        if datos["cups"] and datos["cups"] != s.cups:
            s.cups = datos["cups"]
        s.denominacion, s.direccion = datos["instalacion"], datos["direccion"]
        if s.tarifa != datos["tarifa"]:
            s.tarifa = datos["tarifa"]
            s.invalidar()
        _zona_individual(pot, s)
        editada = st.data_editor(
            pd.DataFrame([dict(zip(PERIODOS, s.actual))], dtype=float), hide_index=True,
            key="pot_actual_%d" % v,
            column_config={p: st.column_config.NumberColumn(f"{p} (kW)", min_value=0,
                                                            format="%g") for p in PERIODOS})
        s.actual = [None if pd.isna(x) else float(x) for x in editada.iloc[0][PERIODOS]]
        return
    st.caption("La zona sale del código postal de cada suministro en el inventario de Gemweb; "
               "si la cambias, su curva se vuelve a leer con esa zona.")
    tabla = pd.DataFrame([dict({"CUPS": s.cups, "Denominación": s.denominacion,
                                "Dirección": s.direccion, "Tarifa": s.tarifa, "Zona": s.zona,
                                "Origen de la zona": s.zona_origen},
                               **dict(zip(PERIODOS, s.actual))) for s in sums])
    for p in PERIODOS:
        tabla[p] = pd.to_numeric(tabla[p])
    editada = st.data_editor(
        tabla, hide_index=True, use_container_width=True, key="pot_sums_%d" % v,
        disabled=["CUPS", "Origen de la zona"],
        column_config={"Tarifa": st.column_config.SelectboxColumn(options=sv.TARIFAS,
                                                                  required=True),
                       "Zona": st.column_config.SelectboxColumn(options=ZONAS, required=True),
                       **{p: st.column_config.NumberColumn(f"{p} (kW)", min_value=0, format="%g")
                          for p in PERIODOS}})
    for s, (_, fila) in zip(sums, editada.iterrows()):
        s.denominacion = str(fila["Denominación"] or "")
        s.direccion = str(fila["Dirección"] or "")
        if fila["Tarifa"] != s.tarifa:
            s.tarifa = fila["Tarifa"]
            s.invalidar()
        if fila["Zona"] != s.zona:
            s.zona, s.zona_manual, s.zona_origen = fila["Zona"], True, "elegida a mano"
            _releer(pot, s)
        s.actual = [None if pd.isna(fila[p]) else float(fila[p]) for p in PERIODOS]


def _propuestas(pot, s, i):
    """Propuestas editables (la 1 es la optima) y la inversion de cada una."""
    v, ed = pot["version"], _edicion(pot, i)
    b1, b2, b3, _ = st.columns([1, 1, 1, 2])
    if b1.button("Añadir propuesta", key=f"pot_add_{v}_{i}",
                 disabled=len(s.propuestas) >= sv.MAX_PROPUESTAS):
        s.propuestas.append({"pc": None, "opciones": None})
        pot["ediciones"][i] = ed + 1
        st.rerun()
    if b2.button("Quitar la última", key=f"pot_del_{v}_{i}", disabled=len(s.propuestas) <= 2):
        s.propuestas.pop()
        pot["ediciones"][i] = ed + 1
        st.rerun()
    if b3.button("Copiar la óptima", key=f"pot_copia_{v}_{i}",
                 help="Pone la potencia óptima en las propuestas editables."):
        for p in s.propuestas[1:]:
            p["pc"] = None
        pot["ediciones"][i] = ed + 1
        st.rerun()
    optimo = [float(x) for x in s.optimo]
    filas = [dict({"Propuesta": f"Propuesta {k}"}, **dict(zip(PERIODOS, p["pc"] or optimo)))
             for k, p in enumerate(s.propuestas[1:], start=2)]
    editada = st.data_editor(
        pd.DataFrame(filas), hide_index=True, key=f"pot_prop_{v}_{i}_{ed}",
        disabled=["Propuesta"],
        column_config={p: st.column_config.NumberColumn(f"{p} (kW)", min_value=0, format="%g")
                       for p in PERIODOS})
    st.caption("La Propuesta 1 es la potencia óptima. Las demás se pueden cambiar; si se dejan "
               "igual que la óptima, se recalculan con ella.")
    for p, (_, fila) in zip(s.propuestas[1:], editada.iterrows()):
        pc = [None if pd.isna(fila[q]) else float(fila[q]) for q in PERIODOS]
        p["pc"] = None if pc == optimo else pc
    # inversion: por defecto, todos los derechos si se amplia la potencia maxima
    defecto = []
    for k, p in enumerate(s.propuestas, start=1):
        pc = optimo if k == 1 or not p["pc"] or None in p["pc"] else p["pc"]
        defecto.append(sv.opciones_inversion_defecto(pc, s.actual)
                       if all(x for x in s.actual) else {"enganche": True})
    inv = pd.DataFrame([dict({"Propuesta": f"Propuesta {k}"},
                             **{t: bool((p["opciones"] or d).get(c)) for c, t in
                                sv.CONCEPTOS_INVERSION})
                        for k, (p, d) in enumerate(zip(s.propuestas, defecto), start=1)])
    with st.popover("Inversión de cada propuesta"):        # cabe dentro del expander del CUPS
        editada = st.data_editor(
            inv, hide_index=True, key=f"pot_inv_{v}_{i}_{ed}", disabled=["Propuesta"],
            column_config={t: st.column_config.CheckboxColumn() for _, t in
                           sv.CONCEPTOS_INVERSION})
        st.caption("Por defecto: si la potencia máxima de la propuesta supera la actual, "
                   "derechos de acceso y extensión sobre el aumento y enganche; si no, solo "
                   "enganche. Acceso y extensión solo cuentan si se modifica la acometida.")
    for p, d, (_, fila) in zip(s.propuestas, defecto, editada.iterrows()):
        opciones = {c: bool(fila[t]) for c, t in sv.CONCEPTOS_INVERSION}
        p["opciones"] = None if all(opciones[c] == bool(d.get(c)) for c in opciones) else opciones


def _resultados(s):
    tabla = sv.tabla_resultados(s.escenarios)

    def rojo(valor):
        return "color: #C00000" if str(valor).startswith("-") and str(valor) != "-" else ""
    st.dataframe(tabla.style.map(rojo, subset=["Ahorro (€) *"]), hide_index=True,
                 use_container_width=True)
    st.caption(sv.texto_notas(s.estudio, s.escenarios))


def _grafico(dibujar, *args, figsize=(10, 4), **kwargs):
    st.pyplot(sv.figura(dibujar, *args, figsize=figsize, **kwargs), use_container_width=True)


def _calculo(pot, precios):
    """Calcula todo; devuelve el resumen multipunto (o None) si no hay errores."""
    sums, multipunto = pot["suministros"], pot["tipo"] == MULTIPUNTO
    errores = []
    for s in sums:
        if s.zona_curva != s.zona and pot["lector"] is not None:
            _releer(pot, s)
        if s.curva is None:
            errores.append(f"{s.nombre}: falta la curva de carga.")
        elif any(p is None or p <= 0 for p in s.actual):
            errores.append(f"{s.nombre}: faltan las 6 potencias contratadas actuales.")
        else:
            sv.preparar(s, precios)
    if errores:
        st.warning("\n\n".join(errores))
        return False
    return True


def _propuestas_y_resultados(pot, precios):
    st.subheader("4. Propuestas y resultados")
    sums, multipunto = pot["suministros"], pot["tipo"] == MULTIPUNTO
    if not pot["calculado"]:
        if st.button("Calcular", type="primary"):
            pot["calculado"] = True
            st.rerun()
        return None
    if not _calculo(pot, precios):
        return None
    errores = []
    if multipunto:
        for i, s in enumerate(sums):
            with st.expander(" · ".join(x for x in (s.cups, s.denominacion, s.tarifa, s.zona)
                                        if x)):
                _propuestas(pot, s, i)
                err = sv.calcular(s, precios)
                if err:
                    errores.append(f"{s.cups}: {err}")
                else:
                    _resultados(s)
    else:
        s = sums[0]
        _propuestas(pot, s, 0)
        err = sv.calcular(s, precios)
        if err:
            errores.append(err)
    if errores:
        st.warning("\n\n".join(errores))
        return None
    avisos = [(f"{s.nombre}: " if multipunto else "") + a for s in sums
              for a in sv.avisos_calculo(s)]
    res = sv.resumen(sums) if multipunto else None
    if multipunto and res.periodos_distintos:
        avisos.append("Las curvas de los CUPS no cubren los mismos meses: el resumen mensual suma "
                      "los meses disponibles de cada uno.")
    for a in avisos:
        st.warning(a)
    if multipunto:
        st.markdown("#### Resumen conjunto")
        st.dataframe(sv.tabla_cups(sv.filas_cups(sums), res), hide_index=True,
                     use_container_width=True)
        st.caption("Coste de la potencia sin impuesto eléctrico. Ahorro y % respecto al coste "
                   "actual; inversión y PRS de la propuesta óptima (Propuesta 1) de cada CUPS.")
        c1, c2 = st.columns([2, 3])
        c1.dataframe(sv.tabla_mensual(res.etiquetas_meses, res.escenarios), hide_index=True,
                     use_container_width=True)
        with c2:
            _grafico(graficos.dibujar_costes, res, res.escenarios,
                     titulo="Coste de la potencia facturada – conjunto (€)", figsize=(7, 4))
    else:
        s = sums[0]
        _resultados(s)
        c1, c2 = st.columns([2, 3])
        c1.dataframe(sv.tabla_mensual(s.estudio.etiquetas_meses, s.escenarios), hide_index=True,
                     use_container_width=True)
        with c2:
            _grafico(graficos.dibujar_costes, s.estudio, s.escenarios, figsize=(7, 4))
    return res if multipunto else True


def _anexos(pot, res) -> tuple:
    """Anexos (energia, maximos, detalle de costes, curva, registro). Devuelve la seleccion de
    la curva (periodo, mes) para el PDF."""
    st.subheader("5. Anexos")
    sums, multipunto = pot["suministros"], pot["tipo"] == MULTIPUNTO
    s = sums[0]
    if multipunto:
        elegido = st.selectbox("CUPS de los anexos", [sv.TODOS] + [x.cups for x in sums],
                               key="pot_anexo_cups")
        s = next((x for x in sums if x.cups == elegido), None)
    est, escs = (s.estudio, s.escenarios) if s else (res, res.escenarios)
    t_e, t_m, t_d, t_c, t_r = st.tabs(["Energía (kWh)", "Potencias máximas (kW)",
                                       "Detalle de costes", "Curva de carga",
                                       "Registro de carga"])
    with t_e:
        st.dataframe(sv.tabla_energia(est), hide_index=True, use_container_width=True)
    with t_m:
        if s:
            st.dataframe(sv.tabla_maximos(est), hide_index=True, use_container_width=True)
            _grafico(graficos.dibujar_maximos, est, escs[0].pc, figsize=(10, 3.5))
        else:
            st.caption("Potencia máxima anual de cada CUPS por periodo.")
            st.dataframe(sv.tabla_maximos_cups(sums), hide_index=True, use_container_width=True)
    with t_d:
        nombres = [e.nombre for e in escs]
        nombre = st.selectbox("Escenario", nombres, key="pot_detalle")
        st.dataframe(sv.tabla_detalle(est, escs[nombres.index(nombre)]), hide_index=True,
                     use_container_width=True)
    with t_c:
        c1, c2 = st.columns(2)
        periodo = c1.selectbox("Periodo", ["Todos"] + PERIODOS, key="pot_sel_periodo")
        meses = [ANIO] + (list(est.etiquetas_meses) if s else [])
        mes = c2.selectbox("Mes", meses, key="pot_sel_mes_%d" % pot["version"])
        seleccion = (0 if periodo == "Todos" else int(periodo[1]), None if mes == ANIO else mes)
        if s:
            _grafico(graficos.dibujar_curva_selector, est, escs, seleccion[0],
                     None if mes == ANIO else meses.index(mes) - 1, figsize=(10, 4.5))
            st.caption("El informe PDF lleva el gráfico de la curva con esta misma selección.")
        else:
            st.info("Elige un CUPS concreto para ver su curva de carga.")
    with t_r:
        st.code(sv.registro_carga(sums, multipunto, pot["errores"]) or "(sin curva)",
                language=None)
    return seleccion


def _informe_y_guardar(pot, datos, res, seleccion, precios):
    st.subheader("6. Informe y guardado")
    sums, multipunto = pot["suministros"], pot["tipo"] == MULTIPUNTO
    st.session_state.setdefault("pot_pdf_anexos", True)
    anexos = st.checkbox("Incluir los anexos en el PDF (energía, potencias máximas, coste "
                         "mensual y curva de carga" + (" de cada CUPS" if multipunto else "")
                         + ")", key="pot_pdf_anexos")
    firma = (repr(datos), anexos, seleccion,
             tuple((s.cups, s.zona, s.tarifa, tuple(round(e.coste.total, 2) for e in s.escenarios),
                    s.denominacion, s.direccion) for s in sums))
    if pot.get("pdf_firma") != firma:
        pot["pdf"] = None
    nombre = sv.nombre_pdf(datos, multipunto)
    c1, c2, c3 = st.columns(3)
    if c1.button("Generar el informe PDF", type="primary", disabled=pot["pdf"] is not None):
        with st.spinner("Generando el PDF…"):
            pot["pdf"] = _generar_pdf(pot, datos, res, anexos, seleccion, precios)
            pot["pdf_firma"] = firma
        st.rerun()
    c2.download_button("Descargar el PDF", data=pot["pdf"] or b"", file_name=nombre,
                       mime="application/pdf", disabled=pot["pdf"] is None)
    if c3.button("Guardar estudio",
                 help="Guarda el estudio (datos, curvas, propuestas y PDF) en el ordenador de la "
                      "plataforma. Se consulta en «Estudios guardados»."):
        if pot["pdf"] is None:
            with st.spinner("Generando el PDF…"):
                pot["pdf"] = _generar_pdf(pot, datos, res, anexos, seleccion, precios)
                pot["pdf_firma"] = firma
        usuario = (st.session_state.get("usuario") or {}).get("usuario", "")
        guardar = {"tipo": pot["tipo"], "datos": datos, "ajustes": pot["ajustes"],
                   "nombre": pot["nombre"], "anexos": anexos}
        pot["guardado"] = almacen.guardar_estudio_potencia(
            usuario, sv.resumen_guardar(pot["tipo"], datos, sums), guardar,
            sv.empaquetar(sums), pot["pdf"])
    if pot.get("guardado"):
        st.success(f"Estudio guardado con el nº {pot['guardado']}. Está en «Estudio de potencia → "
                   "Estudios guardados».")


def _generar_pdf(pot, datos, res, anexos, seleccion, precios) -> bytes:
    if pot["tipo"] == MULTIPUNTO:
        return sv.pdf_multipunto(datos, pot["suministros"], res, precios, anexos, seleccion)
    return sv.pdf_individual(datos, pot["suministros"][0], precios, anexos, seleccion)


# ------------------------------------------------------------------- paginas
def pagina_nuevo_estudio():
    pot = _estado()
    _aplicar_pendiente(pot)
    cargado = pot.get("cargado")
    if cargado:
        st.info(f"Estudio nº {cargado['id']} cargado (guardado por {cargado['usuario']} el "
                f"{cargado['fecha']}). Cambia lo que necesites; al guardar se crea un estudio "
                "nuevo y el original no se modifica. Para cambiar la zona o los ajustes de "
                "lectura de la curva, vuelve a cargarla.")
        if st.button("Empezar un estudio nuevo"):
            _reiniciar()
            st.rerun()
    st.markdown("Optimización de la potencia contratada de suministros **6.1TD** y **3.0TD "
                "> 50 kW** a partir de la curva de carga de 12 meses (Circular CNMC 3/2020): coste "
                "actual, potencia óptima, otras propuestas, inversión y periodo de retorno.")
    tipos = [INDIVIDUAL, MULTIPUNTO]
    tipo = st.radio("Tipo de estudio", tipos, horizontal=True, key="pot_tipo",
                    format_func=lambda t: "Individual (un suministro)" if t == INDIVIDUAL
                    else "Multipunto (varios CUPS a la vez)", disabled=bool(cargado))
    if tipo != pot["tipo"]:
        _reiniciar(tipo)
        st.rerun()
    precios = sv.precios()
    datos = _datos(pot)
    if not cargado:
        _curva(pot)
    if not pot["suministros"]:
        return
    _interpretacion(pot)
    _suministros(pot, datos)
    res = _propuestas_y_resultados(pot, precios)
    if res is None:
        return
    seleccion = _anexos(pot, res if pot["tipo"] == MULTIPUNTO else None)
    _informe_y_guardar(pot, datos, res if pot["tipo"] == MULTIPUNTO else None, seleccion,
                       precios)


def cargar_en_formulario(estudio: dict) -> None:
    """Deja el estudio guardado listo en «Nuevo estudio» (sin la curva original: se usan las
    curvas ya leidas que se guardaron)."""
    datos = estudio["datos"]
    tipo = datos.get("tipo", INDIVIDUAL)
    _reiniciar(tipo)
    pot = st.session_state.pot
    d = datos.get("datos", {})
    pot.update(suministros=sv.desempaquetar(estudio["curvas"]), ajustes=datos.get("ajustes", {}),
               nombre=datos.get("nombre", ""), calculado=True,
               cargado={"id": estudio["id"], "usuario": estudio.get("usuario"),
                        "fecha": _fecha_hora(estudio.get("fecha"))})
    pot["pendiente"].update(pot_titular=d.get("titular", ""), pot_fecha=d.get("fecha", ""),
                            pot_pdf_anexos=bool(datos.get("anexos", True)))
    if tipo == INDIVIDUAL:
        s = pot["suministros"][0]
        pot["pendiente"].update(pot_cups=s.cups, pot_instalacion=s.denominacion,
                                pot_direccion=s.direccion, pot_tarifa=s.tarifa)


def _fecha_hora(iso):
    from datetime import datetime
    return datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M") if iso else ""


def pagina_estudios_guardados():
    estudios = almacen.listar_estudios_potencia()
    if not estudios:
        st.info("Todavía no hay estudios guardados. Se guardan con el botón «Guardar estudio», al "
                "final de «Nuevo estudio».")
        return
    nombres = {u["usuario"]: u["nombre"] for u in almacen.listar_usuarios()}
    buscar = st.text_input("Buscar por titular, CUPS o usuario").strip().lower()
    filtrados = [e for e in estudios if not buscar or buscar in " ".join(
        str(e.get(k) or "") for k in ("titular", "cups", "usuario")).lower()
        or buscar in nombres.get(e["usuario"], "").lower()]
    st.dataframe(pd.DataFrame([{
        "Nº": e["id"], "Guardado": _fecha_hora(e["fecha"]),
        "Usuario": nombres.get(e["usuario"], e["usuario"]),
        "Tipo": "Multipunto" if e["tipo"] == MULTIPUNTO else "Individual",
        "Titular": e["titular"], "CUPS": e["cups"], "Tarifa": e["tarifa"],
        "Ahorro anual (€)": e["ahorro"], "Inversión (€)": e["inversion"]} for e in filtrados]),
        hide_index=True, use_container_width=True,
        column_config={c: st.column_config.NumberColumn(format="localized")
                       for c in ("Ahorro anual (€)", "Inversión (€)")})
    if not filtrados:
        return
    st.subheader("Detalle")
    por_id = {e["id"]: e for e in filtrados}
    elegido = st.selectbox("Estudio", list(por_id), format_func=lambda i: "Nº %d · %s · %s" % (
        i, _fecha_hora(por_id[i]["fecha"]), por_id[i]["titular"] or por_id[i]["cups"] or "—"))
    estudio = almacen.estudio_potencia(elegido)
    c1, c2, c3 = st.columns(3)
    c1.metric("Suministros", estudio["suministros"])
    c2.metric("Ahorro anual (propuesta óptima)",
              f"{estudio['ahorro'] or 0:,.0f} €".replace(",", "."))
    c3.metric("Inversión", f"{estudio['inversion'] or 0:,.2f} €".replace(",", "X")
              .replace(".", ",").replace("X", "."))
    st.caption(f"Guardado por {nombres.get(estudio['usuario'], estudio['usuario'])} el "
               f"{_fecha_hora(estudio['fecha'])}.")
    with st.expander(f"CUPS ({estudio['suministros']})"):
        st.code((estudio["cups"] or "").replace(", ", "\n"), language=None)
    datos = estudio["datos"]
    b1, b2 = st.columns(2)
    b1.download_button("Descargar el PDF", data=estudio["pdf"] or b"",
                       file_name=sv.nombre_pdf(dict(datos.get("datos", {}),
                                                    instalacion=_instalacion(estudio)),
                                               datos.get("tipo") == MULTIPUNTO),
                       mime="application/pdf", type="primary", disabled=not estudio["pdf"])
    if b2.button("Cargar en el formulario",
                 help="Abre «Nuevo estudio» con este estudio para cambiar potencias o "
                      "propuestas. Al guardar se crea uno nuevo."):
        cargar_en_formulario(estudio)
        if PAGINA_FORMULARIO is not None:
            st.switch_page(PAGINA_FORMULARIO)
        st.success("Cargado: ve a «Nuevo estudio».")


def _instalacion(estudio):
    try:
        return sv.desempaquetar(estudio["curvas"])[0].denominacion
    except Exception:
        return ""
