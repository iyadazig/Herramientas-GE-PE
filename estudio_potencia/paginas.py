# -*- coding: utf-8 -*-
"""
PAGINAS DE «ESTUDIO DE POTENCIA» de Herramientas GE&PE
======================================================
Optimizacion de la potencia contratada (6.1TD y 3.0TD > 50 kW) a partir de la curva de carga,
como el programa de escritorio EstudioPotencia: propuestas, anexos, informe PDF y estudios
guardados para todo el equipo.

Lo primero son los CUPS: se buscan en el inventario de Gemweb, que rellena titular (razon
social), denominacion, direccion, tarifa, potencias contratadas y zona (por el codigo postal),
y se descarga su curva de carga. Los que no estan en Gemweb se completan a mano y su curva se
sube en un fichero (tambien se puede subir para sustituir la de Gemweb). Con 1 CUPS el estudio
es individual y con varios, multipunto. La logica esta en estudio_potencia/servicio.py.
"""

import tempfile
from datetime import date, datetime
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
AJUSTES_GEMWEB = {"hoja": None, "unidad": UNIDADES[0], "convenio": "Fin de periodo",
                  "cups": None}


# ------------------------------------------------------------------- estado
def _estado() -> dict:
    """Estado del estudio en la sesion: los suministros, de donde sale la curva de cada uno
    (lectores: {cups: {lector, ajustes, cups del fichero}}), el PDF..."""
    return st.session_state.setdefault("pot", {
        "suministros": [], "lectores": {}, "errores": [], "calculado": False, "cargado": None,
        "pdf": None, "pdf_firma": None, "version": 0, "ediciones": {}, "pendiente": {}})


def _tipo(pot) -> str:
    return MULTIPUNTO if len(pot["suministros"]) > 1 else INDIVIDUAL


def _reiniciar():
    for k in [k for k in st.session_state if str(k).startswith("pot")]:
        del st.session_state[k]
    _estado()


def _aplicar_pendiente(pot):
    """Valores para los controles que hay que cambiar desde el codigo (se aplican antes de
    dibujarlos: Streamlit no deja cambiar un control ya dibujado)."""
    for clave, valor in pot.pop("pendiente", {}).items():
        st.session_state[clave] = valor
    pot["pendiente"] = {}


def _edicion(pot, i):
    """Contador para volver a empezar las tablas de propuestas de un suministro."""
    return pot["ediciones"].get(i, 0)


def _base(pot, clave, construir) -> pd.DataFrame:
    """Datos de partida de una tabla editable, construidos una sola vez por clave. Si se
    reconstruyeran en cada recarga con lo ya editado, Streamlit reiniciaria la tabla y la
    primera edicion se perderia (habria que hacerla dos veces)."""
    bases = pot.setdefault("bases", {})
    if clave not in bases:
        bases[clave] = construir()
    return bases[clave]


# ------------------------------------------------------------------- Gemweb
def _cliente_gemweb():
    cred, _origen = gemweb.cargar_credenciales()
    return gemweb.ClienteGemweb(*cred) if cred else None


def _centro(cliente, sumi):
    """Centro de consumo del suministro, solo si el suministro no trae su codigo postal."""
    if not sumi or sv.zona_por_codigo_postal(sumi.get("codi_postal")) \
            or not sumi.get("id_centres_consum"):
        return None
    return gw.registro_inventario(cliente.api, "centres_consum", "centres_consum.id",
                                  sumi["id_centres_consum"])


def _inventario(cups) -> tuple:
    """(suministro, centro de consumo, entidad, hay credenciales) de Gemweb, guardado en la
    sesion."""
    cache = st.session_state.setdefault("pot_inventario", {})
    if cups not in cache:
        cliente = _cliente_gemweb()
        sumi = centro = entidad = None
        if cliente is not None:
            try:
                sumi = cliente.buscar_suministro(cups)
                centro = _centro(cliente, sumi)
                entidad = gw.registro_inventario(cliente.api, "entitats", "entitats.id",
                                                 (sumi or {}).get("id_entitats"))
            except Exception:            # sin conexion o fallo de la API: se completa a mano
                sumi = None
        cache[cups] = (sumi, centro, entidad, cliente is not None)
    return cache[cups]


def _suministro_de_gemweb(cups) -> sv.Suministro:
    """Suministro con los datos del inventario de Gemweb (o vacio si no esta)."""
    s = sv.Suministro(cups=cups)
    sumi, centro, entidad, hay = _inventario(cups)
    if sumi:
        s.en_gemweb = True
        sv.rellenar_desde_gemweb(s, sumi)
        s.titular = (entidad or {}).get("nom") or ""
        s.grupo = (entidad or {}).get("divisio") or ""
        s.zona, s.zona_origen = sv.zona_de(sumi, centro)
    else:
        s.zona_origen = ("no está en Gemweb: elige la zona" if hay
                         else "sin credenciales de Gemweb: elige la zona")
    return s


# ------------------------------------------------------------------- curvas
def _leer(pot, s):
    """(Re)lee la curva del suministro con su zona desde su origen (Gemweb o fichero)."""
    origen = pot["lectores"].get(s.cups)
    if origen is None:
        s.invalidar()                    # estudio guardado: se recalcula con la curva guardada
        return
    try:
        sv.leer_suministro(s, origen["lector"], origen["ajustes"], True, origen["cups"])
    except Exception as e:
        s.curva = None
        st.error(f"{s.cups}: no se ha podido leer la curva: {e}")


def _buscar(pot, lista, desde, hasta):
    """Busca los CUPS en Gemweb, rellena sus datos y descarga la curva de los que estan."""
    with st.spinner("Buscando los CUPS en el inventario de Gemweb…"):
        sums = [_suministro_de_gemweb(c) for c in lista]
    pot.update(suministros=sums, lectores={}, errores=[], calculado=False, cargado=None,
               pdf=None, pdf_firma=None, version=pot["version"] + 1, ediciones={})
    en_gemweb = [s.cups for s in sums if s.en_gemweb]
    if en_gemweb:
        _descargar(pot, en_gemweb, desde, hasta)
    pot["pendiente"] = {"pot_titular": sv.titular_comun(sums)}
    st.rerun()


def _descargar(pot, lista, desde, hasta):
    cliente = _cliente_gemweb()
    barra = st.progress(0.0, text="Descargando la curva de carga de Gemweb…")
    try:
        ruta, info, avisos = sv.descargar_gemweb(
            cliente, lista, desde, hasta, tempfile.mkdtemp(prefix="potencia_"),
            al_avanzar=lambda fraccion, texto: barra.progress(min(fraccion, 1.0), text=texto))
    except gemweb.GemwebError as e:
        pot["errores"].append(f"Error de Gemweb al descargar las curvas: {e}")
        return
    finally:
        barra.empty()
    pot["errores"] += [a for a in avisos if not any(a.startswith(c) for c in info)]
    if ruta is None:
        return
    lector = LectorCurva(ruta)
    for s in pot["suministros"]:
        if s.cups in info:
            s.avisos_extra = sv.avisos_gemweb(s.cups, avisos, info[s.cups])
            pot["lectores"][s.cups] = {"lector": lector, "ajustes": dict(AJUSTES_GEMWEB),
                                       "cups": s.cups}
            s.fuente = f"Gemweb ({desde:%d/%m/%Y} – {hasta:%d/%m/%Y})"
            _leer(pot, s)


def _cargar_ficheros(pot, subidos, destino):
    """Lee los ficheros subidos y los asigna a sus CUPS (por la columna de CUPS o el nombre
    del fichero; con un solo CUPS en el estudio, a ese). destino: CUPS a los que se aplica."""
    carpeta = Path(tempfile.mkdtemp(prefix="potencia_"))
    rutas = sv.guardar_subidos([(f.name, f.getvalue()) for f in subidos], carpeta)
    nombre = rutas[0].name if len(rutas) == 1 else f"{len(rutas)} ficheros"
    try:
        lector = sv.abrir(rutas)
        ajustes = {"hoja": lector.hoja_defecto, "unidad": UNIDADES[0],
                   "convenio": CONVENIOS[0], "cups": None}
        sv.interpretar(lector, sv.PENINSULA, ajustes)              # detecta los CUPS
    except Exception as e:
        st.error(f"No se ha podido leer {nombre}: {e}")
        return
    disponibles = list(lector.cups_disponibles)
    asignados, sin_curva = [], []
    for s in pot["suministros"]:
        if s.cups not in destino:
            continue
        cups_fichero = sv.cups_del_fichero(s.cups, disponibles)
        if cups_fichero is None and len(pot["suministros"]) == 1 and len(disponibles) <= 1:
            cups_fichero = disponibles[0] if disponibles else None
        elif cups_fichero is None:
            sin_curva.append(s.cups)
            continue
        pot["lectores"][s.cups] = {"lector": lector, "ajustes": dict(ajustes),
                                   "cups": cups_fichero}
        s.fuente, s.avisos_extra = f"fichero: {nombre}", []
        _leer(pot, s)
        asignados.append(s.cups)
    if sin_curva:
        st.warning("Los ficheros no traen la curva de: " + ", ".join(sin_curva)
                   + ". Su CUPS tiene que estar en una columna o en el nombre del fichero.")
    if asignados:
        pot.update(calculado=False, pdf=None, version=pot["version"] + 1)
        st.rerun()


COMO_SUBIR = """**Cómo tienen que ser los ficheros de curva**
- CSV, TXT o Excel (.xlsx, .xls) con la curva de carga **cuartohoraria** (o horaria) de los
  **12 últimos meses**.
- Una columna con la **fecha y hora** (o una de fecha y otra de hora) y otra con los **valores en
  kW o kWh**. Mejor si el título de la columna dice la unidad, p. ej. «Consumo (kWh)» o
  «Potencia (kW)»; si no, se puede indicar después en «Interpretación de la curva de carga».
- Valen los formatos habituales: P5D de la distribuidora, Datadis, el Excel de GE&PE, una fila
  por día con 24 columnas…
- **Varios suministros**: o una columna con el **CUPS** de cada registro, o un fichero por
  suministro con el **CUPS en el nombre** (p. ej. `ES0000000000000000XX_curva.xlsx`).
- Se pueden subir **varios ficheros del mismo suministro** (p. ej. uno por mes): se unen en una
  sola curva.
- Un solo suministro sin CUPS en el fichero: escribe su CUPS después, en la tabla de datos."""


def _crear_desde_ficheros(pot, subidos, titular):
    """Estudio de CUPS que no estan en Gemweb: un suministro por CUPS de las curvas subidas,
    con el titular indicado; el resto de datos se rellena a mano en la tabla."""
    carpeta = Path(tempfile.mkdtemp(prefix="potencia_"))
    rutas = sv.guardar_subidos([(f.name, f.getvalue()) for f in subidos], carpeta)
    nombre = rutas[0].name if len(rutas) == 1 else f"{len(rutas)} ficheros"
    try:
        lector = sv.abrir(rutas)
        ajustes = {"hoja": lector.hoja_defecto, "unidad": UNIDADES[0],
                   "convenio": CONVENIOS[0], "cups": None}
        sv.interpretar(lector, sv.PENINSULA, ajustes)              # detecta los CUPS
    except Exception as e:
        st.error(f"No se ha podido leer {nombre}: {e}")
        return
    sums, pot["lectores"] = [], {}
    for cups in list(lector.cups_disponibles) or [""]:
        s = sv.Suministro(cups=cups, titular=titular, fuente=f"fichero: {nombre}",
                          zona_origen="no está en Gemweb: elige la zona")
        pot["lectores"][cups] = {"lector": lector, "ajustes": dict(ajustes),
                                 "cups": cups or None}
        sums.append(s)
    pot.update(suministros=sums, errores=[], calculado=False, cargado=None, pdf=None,
               pdf_firma=None, version=pot["version"] + 1, ediciones={})
    for s in sums:
        _leer(pot, s)
    pot["pendiente"] = {"pot_titular": titular}
    st.rerun()


# ------------------------------------------------------------------- pasos del formulario
def _paso_cups(pot):
    st.subheader("1. CUPS")
    en_gemweb = st.radio("¿Los CUPS del estudio están en Gemweb?", [True, False],
                         horizontal=True, key="pot_modo",
                         format_func=lambda x: "Sí: buscarlos en Gemweb" if x
                         else "No: introducir los datos y subir las curvas")
    if not en_gemweb:
        st.caption("Escribe el titular y sube las curvas de carga: se crea un suministro por "
                   "cada CUPS de las curvas (con uno es un estudio individual y con varios, "
                   "multipunto). Después rellenas a mano el resto de sus datos.")
        titular = st.text_input("Titular", key="pot_titular_manual").strip()
        st.info(COMO_SUBIR)
        subidos = st.file_uploader("Ficheros de curva", type=sv.EXTENSIONES,
                                   accept_multiple_files=True,
                                   key="pot_subidos_inicio_%d" % pot["version"])
        if not titular:
            st.caption("Falta el titular.")
        if st.button("Leer las curvas", type="primary", disabled=not subidos or not titular):
            _crear_desde_ficheros(pot, subidos, titular)
        return
    st.caption("Pega los CUPS del estudio: con uno es un estudio individual y con varios, "
               "multipunto. Se buscan en el inventario de Gemweb, que rellena sus datos y "
               "descarga su curva de carga.")
    texto = st.text_area("CUPS (uno por línea, o separados por comas o espacios)",
                         key="pot_cups_texto", height=110, placeholder="ES0000000000000000XX")
    ini, fin = sv.periodo_defecto()
    c1, c2, _ = st.columns([1, 1, 2])
    desde = c1.date_input("Curva desde", value=ini, format="DD/MM/YYYY", key="pot_desde")
    hasta = c2.date_input("hasta", value=fin, format="DD/MM/YYYY", key="pot_hasta")
    lista = sv.lista_cups(texto)
    if st.button("Buscar en Gemweb y descargar las curvas", type="primary",
                 disabled=not lista or hasta <= desde):
        _buscar(pot, lista, desde, hasta)


def _paso_datos(pot) -> dict:
    """Titular del estudio y tabla de datos de los suministros (de Gemweb o a mano)."""
    st.subheader("2. Datos de los suministros")
    sums, v = pot["suministros"], pot["version"]
    multipunto = _tipo(pot) == MULTIPUNTO
    st.session_state.setdefault("pot_fecha", date.today().strftime("%d/%m/%Y"))
    c1, c2 = st.columns([3, 1])
    datos = {"titular": c1.text_input("Titular del estudio" if multipunto else "Titular",
                                      key="pot_titular").strip(),
             "fecha": c2.text_input("Fecha", key="pot_fecha").strip()}
    manual = all(not s.en_gemweb for s in sums)
    fuera = [s.cups or "(sin CUPS)" for s in sums if not s.en_gemweb and (
        not s.cups or not s.titular.strip() or any(p is None for p in s.actual)
        or s.curva is None)]
    if fuera and manual:
        st.warning("Completa en la tabla los datos de " + ", ".join(fuera) + ": CUPS, "
                   "denominación, dirección, tarifa, zona y potencias contratadas.")
    elif fuera:
        st.warning("No están en Gemweb: " + ", ".join(fuera) + ". Completa sus datos en la tabla "
                   "(titular, denominación, dirección, tarifa, zona y potencias) y sube su curva "
                   "de carga en el paso 3.")
    st.caption("Los datos de Gemweb se pueden cambiar. La zona sale del código postal del "
               "suministro; si la cambias, su curva se vuelve a leer con ella."
               if not manual else "Rellena los datos de cada suministro (y su CUPS si el fichero "
               "no lo traía). Pulsa Enter o haz clic fuera de la celda para aplicar cada cambio.")

    def construir():
        tabla = pd.DataFrame([dict({"CUPS": s.cups, "En Gemweb": "Sí" if s.en_gemweb else "No",
                                    "Titular": s.titular, "Denominación": s.denominacion,
                                    "Dirección": s.direccion, "Tarifa": s.tarifa, "Zona": s.zona,
                                    "Origen de la zona": s.zona_origen},
                                   **dict(zip(PERIODOS, s.actual))) for s in sums])
        for p in PERIODOS:
            tabla[p] = pd.to_numeric(tabla[p])
        return tabla
    editada = st.data_editor(
        _base(pot, "datos_%d" % v, construir), hide_index=True, use_container_width=True,
        key="pot_datos_%d" % v,
        disabled=(["En Gemweb", "Origen de la zona"] if manual
                  else ["CUPS", "En Gemweb", "Origen de la zona"]),
        column_config={"Tarifa": st.column_config.SelectboxColumn(options=sv.TARIFAS,
                                                                  required=True),
                       "Zona": st.column_config.SelectboxColumn(options=ZONAS, required=True),
                       **{p: st.column_config.NumberColumn(f"{p} (kW)", min_value=0, format="%g")
                          for p in PERIODOS}})
    for s, (_, fila) in zip(sums, editada.iterrows()):
        cups = "" if pd.isna(fila["CUPS"]) else str(fila["CUPS"]).strip().upper()
        if cups != s.cups:                  # CUPS escrito a mano: su curva sigue con el
            if s.cups in pot["lectores"]:
                pot["lectores"][cups] = pot["lectores"].pop(s.cups)
            s.cups = cups
        for campo, columna in (("titular", "Titular"), ("denominacion", "Denominación"),
                               ("direccion", "Dirección")):
            setattr(s, campo, "" if pd.isna(fila[columna]) else str(fila[columna]))
        if fila["Tarifa"] != s.tarifa:
            s.tarifa = fila["Tarifa"]
            s.invalidar()
        if fila["Zona"] != s.zona:
            s.zona, s.zona_manual, s.zona_origen = fila["Zona"], True, "elegida a mano"
            _leer(pot, s)
        s.actual = [None if pd.isna(fila[p]) else float(fila[p]) for p in PERIODOS]
    if not multipunto:
        s = sums[0]
        datos.update(cups=s.cups, instalacion=s.denominacion.strip(),
                     direccion=s.direccion.strip())
    return datos


def _paso_curvas(pot):
    st.subheader("3. Curva de carga")
    sums = pot["suministros"]
    st.dataframe(pd.DataFrame([{
        "CUPS": s.cups, "Curva": s.fuente if s.curva is not None else "FALTA",
        "Periodo": (f"{s.curva.fecha_inicio:%d/%m/%Y} – "
                    f"{s.curva.fecha_fin - pd.Timedelta(days=1):%d/%m/%Y}") if s.curva is not None
        else "", "Avisos": str(len(s.curva.avisos)) if s.curva is not None else ""}
        for s in sums]),
        hide_index=True, use_container_width=True)
    for e in pot["errores"]:
        st.warning(e)
    sin_curva = [s.cups for s in sums if s.curva is None]
    if pot["cargado"]:
        return
    with st.expander("Subir ficheros de curva" + (f" ({len(sin_curva)} CUPS sin curva)"
                                                  if sin_curva else ""),
                     expanded=bool(sin_curva)):
        st.markdown(COMO_SUBIR)
        opciones = [s.cups for s in sums]
        destino = st.multiselect("CUPS a los que se aplican", opciones,
                                 default=sin_curva or opciones[:1],
                                 key="pot_destino_%d" % pot["version"],
                                 help="Los que ya tienen la curva de Gemweb la sustituyen por la "
                                      "del fichero.")
        subidos = st.file_uploader("Ficheros de curva", type=sv.EXTENSIONES,
                                   accept_multiple_files=True,
                                   key="pot_subidos_%d" % pot["version"])
        if st.button("Leer los ficheros", disabled=not subidos or not destino):
            _cargar_ficheros(pot, subidos, destino)


def _interpretacion(pot):
    """Lo que se ha entendido de la curva de un suministro y los ajustes para corregirlo."""
    sums = [s for s in pot["suministros"] if s.curva is not None]
    if not sums:
        return
    multipunto = _tipo(pot) == MULTIPUNTO
    with st.expander("Interpretación de la curva de carga", expanded=False):
        s = sums[0]
        if len(sums) > 1:
            cups = st.selectbox("CUPS", [x.cups for x in sums], key="pot_interp_cups")
            s = next(x for x in sums if x.cups == cups)
        origen = pot["lectores"].get(s.cups)
        st.caption(f"Origen: {s.fuente or 'estudio guardado'}")
        if origen is not None and not s.fuente.startswith("Gemweb"):
            lector, a = origen["lector"], dict(origen["ajustes"])
            c1, c2, c3 = st.columns(3)
            v = f"{pot['version']}_{s.cups}"
            if len(lector.hojas) > 1:
                hojas = [str(h) for h in lector.hojas]
                actual = str(a.get("hoja") or lector.hoja_defecto)
                a["hoja"] = c1.selectbox("Hoja", hojas, key=f"pot_hoja_{v}",
                                         index=hojas.index(actual) if actual in hojas else 0)
            a["unidad"] = c2.selectbox("Unidad de los valores", UNIDADES, key=f"pot_unidad_{v}",
                                       index=UNIDADES.index(a.get("unidad") or UNIDADES[0]))
            a["convenio"] = c3.selectbox("La hora de cada registro marca", CONVENIOS,
                                         key=f"pot_convenio_{v}",
                                         index=CONVENIOS.index(a.get("convenio") or CONVENIOS[0]))
            if a != origen["ajustes"]:
                # los demas CUPS del mismo fichero se leen con los mismos ajustes
                for otro in pot["suministros"]:
                    o = pot["lectores"].get(otro.cups)
                    if o is not None and o["lector"] is lector:
                        o["ajustes"] = dict(a)
                        _leer(pot, otro)
                pot["pdf"] = None
        lector = origen["lector"] if origen else None
        st.markdown("\n".join(f"- {l}" for l in sv.resumen_interpretacion(s.curva, lector,
                                                                            False)))
        for aviso in s.curva.avisos:
            st.warning(aviso)
        st.caption("Primeros registros interpretados")
        st.dataframe(sv.primeros_registros(s.curva), hide_index=True, use_container_width=True)


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
    clave = f"pot_prop_{v}_{i}_{ed}_{abs(hash(tuple(optimo)))}"     # otra optima: de nuevo
    editada = st.data_editor(
        _base(pot, clave, lambda: pd.DataFrame([
            dict({"Propuesta": f"Propuesta {k}"}, **dict(zip(PERIODOS, p["pc"] or optimo)))
            for k, p in enumerate(s.propuestas[1:], start=2)])),
        hide_index=True, key=clave, disabled=["Propuesta"],
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
    # si cambian las potencias, cambian las opciones por defecto: la tabla vuelve a empezar
    clave = f"pot_inv_{v}_{i}_{ed}_{abs(hash(repr(defecto)))}"
    inv = _base(pot, clave, lambda: pd.DataFrame([
        dict({"Propuesta": f"Propuesta {k}"},
             **{t: bool((p["opciones"] or d).get(c)) for c, t in sv.CONCEPTOS_INVERSION})
        for k, (p, d) in enumerate(zip(s.propuestas, defecto), start=1)]))
    with st.popover("Inversión de cada propuesta"):        # cabe dentro del expander del CUPS
        editada = st.data_editor(
            inv, hide_index=True, key=clave, disabled=["Propuesta"],
            column_config={t: st.column_config.CheckboxColumn() for _, t in
                           sv.CONCEPTOS_INVERSION})
        st.caption("Por defecto: si la potencia máxima de la propuesta supera la actual, "
                   "derechos de acceso y extensión sobre el aumento y enganche; si no, solo "
                   "enganche. Acceso y extensión solo cuentan si se modifica la acometida.")
    for p, d, (_, fila) in zip(s.propuestas, defecto, editada.iterrows()):
        opciones = {c: bool(fila[t]) for c, t in sv.CONCEPTOS_INVERSION}
        p["opciones"] = None if all(opciones[c] == bool(d.get(c)) for c in opciones) else opciones


# ------------------------------------------------------------------- tablas con color
# como en el programa de escritorio: cabecera azul oscuro, filas alternas, fila de totales
# destacada y columnas de totales con su color (fondo, fondo de la fila de totales)
FONDO_PAR, FONDO_TOTAL, ROJO = "#F4F7FB", "#DCE6F1", "#C00000"
COLOR_FIJO, COLOR_EXCESO, COLOR_TOTAL = ("#FFF7DC", "#FBE7A8"), ("#FDE9E7", "#F6C9C4"), \
    ("#E3EEF9", "#BFD7EE")
CSS_TABLA = """<style>
.gp-tabla {overflow-x: auto; margin: 0.2rem 0 0.8rem 0;}
.gp-tabla table {border-collapse: collapse; font-size: 0.82rem; width: auto;}
.gp-tabla th {background: %s; color: white; font-weight: 700; padding: 6px 9px;
              text-align: center; border: 1px solid white; white-space: nowrap;}
.gp-tabla td {padding: 4px 9px; text-align: right; border: 1px solid white;
              white-space: nowrap; color: #1E1E1E;}
.gp-banda {background: %s; color: white; font-weight: 700; text-align: center;
           padding: 6px; margin-top: 0.6rem; font-size: 0.95rem;}
</style>""" % (graficos.AZUL_OSCURO, graficos.AZUL_OSCURO)


def _tabla(df, totales=1, columnas=None, izquierda=(), centro=None, celdas=None):
    """Tabla HTML con el estilo del programa. totales: nº de filas de totales al final;
    columnas: {columna: (fondo, fondo en la fila de totales)}; izquierda: columnas de texto;
    centro: columnas centradas (por defecto, la primera); celdas(i, columna, valor) -> css extra."""
    df = df.reset_index(drop=True)
    columnas, n = columnas or {}, len(df)
    centro = [df.columns[0]] if centro is None else centro

    def estilo(fila):
        i, total = fila.name, fila.name >= n - totales
        res = []
        for col in df.columns:
            css = []
            if col in columnas:
                css += [f"background-color: {columnas[col][1 if total else 0]}",
                        "font-weight: 700"]
            elif total:
                css += [f"background-color: {FONDO_TOTAL}", "font-weight: 700"]
            elif i % 2:
                css.append(f"background-color: {FONDO_PAR}")
            if col in izquierda:
                css.append("text-align: left")
            elif col in centro:
                css.append("text-align: center")
            if celdas:
                css.append(celdas(i, col, fila[col]) or "")
            res.append("; ".join(c for c in css if c))
        return res
    html = df.style.apply(estilo, axis=1).hide(axis="index").to_html()
    st.markdown(CSS_TABLA + f'<div class="gp-tabla">{html}</div>', unsafe_allow_html=True)


def _negativo(valor) -> bool:
    return str(valor).strip().startswith("-") and str(valor).strip() != "-"


def _resultados(s):
    """Tabla de optimizacion: cada escenario con su color, como en el programa."""
    tabla = sv.tabla_resultados(s.escenarios)

    def celdas(i, col, valor):
        if col == "Tipo":
            return (f"background-color: {graficos.color(i)}; color: white; font-weight: 700; "
                    "text-align: center")
        css = [f"background-color: {graficos.COLORES_SUAVES[i % len(graficos.COLORES_SUAVES)]}"]
        if col in ("Total (€)", "Ahorro (€) *"):
            css.append("font-weight: 700")
        if col == "Ahorro (€) *" and _negativo(valor):
            css.append(f"color: {ROJO}")
        return "; ".join(css)
    st.markdown('<div class="gp-banda">OPTIMIZACIÓN DE POTENCIA (sin i.e.)</div>',
                unsafe_allow_html=True)
    _tabla(tabla, totales=0, celdas=celdas, centro=[f"P{i} (kW)" for i in range(1, N_P + 1)])
    st.caption(sv.texto_notas(s.estudio, s.escenarios))


def _tabla_mensual(etiquetas, escenarios):
    """Coste mensual por escenario y ahorros (en rojo los negativos)."""
    def celdas(i, col, valor):
        return f"color: {ROJO}" if col.startswith("Ahorro") and _negativo(valor) else ""
    _tabla(sv.tabla_mensual(etiquetas, escenarios), celdas=celdas)


def _grafico(dibujar, *args, figsize=(10, 4), **kwargs):
    st.pyplot(sv.figura(dibujar, *args, figsize=figsize, **kwargs), use_container_width=True)


def _calculo(pot, precios) -> bool:
    """Prepara el estudio de cada suministro; False si falta algo."""
    errores = []
    for s in pot["suministros"]:
        if s.curva is not None and s.zona_curva != s.zona:
            _leer(pot, s)
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
    sums, multipunto = pot["suministros"], _tipo(pot) == MULTIPUNTO
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
        _tabla(sv.tabla_cups(sv.filas_cups(sums), res),
               columnas={"Coste actual": ("#E2EFDA", "#C6E0B4"),
                         "Coste óptimo": ("#FCE4D6", "#F8CBAD"), "Ahorro": COLOR_TOTAL},
               izquierda=("CUPS", "Denominación"), centro=["Nº", "Tarifa", "Zona"])
        st.caption("Coste de la potencia sin impuesto eléctrico. Ahorro y % respecto al coste "
                   "actual; inversión y PRS de la propuesta óptima (Propuesta 1) de cada CUPS.")
        c1, c2 = st.columns([2, 3])
        with c1:
            _tabla_mensual(res.etiquetas_meses, res.escenarios)
        with c2:
            _grafico(graficos.dibujar_costes, res, res.escenarios,
                     titulo="Coste de la potencia facturada – conjunto (€)", figsize=(7, 4))
    else:
        s = sums[0]
        _resultados(s)
        c1, c2 = st.columns([2, 3])
        with c1:
            _tabla_mensual(s.estudio.etiquetas_meses, s.escenarios)
        with c2:
            _grafico(graficos.dibujar_costes, s.estudio, s.escenarios, figsize=(7, 4))
    return res if multipunto else True


def _anexos(pot, res) -> tuple:
    """Anexos (energia, maximos, detalle de costes, curva, registro). Devuelve la seleccion de
    la curva (periodo, mes) para el PDF."""
    st.subheader("5. Anexos")
    sums, multipunto = pot["suministros"], _tipo(pot) == MULTIPUNTO
    s = sums[0]
    if multipunto:
        etiquetas = {x.cups: (f"{x.denominacion.strip()} ({x.cups})" if x.denominacion.strip()
                              else x.cups) for x in sums}
        elegido = st.selectbox("Suministro de los anexos", [sv.TODOS] + [x.cups for x in sums],
                               key="pot_anexo_cups", format_func=lambda c: etiquetas.get(c, c))
        s = next((x for x in sums if x.cups == elegido), None)
    est, escs = (s.estudio, s.escenarios) if s else (res, res.escenarios)
    t_e, t_m, t_d, t_c, t_r = st.tabs(["Energía (kWh)", "Potencias máximas (kW)",
                                       "Detalle de costes", "Curva de carga",
                                       "Registro de carga"])
    with t_e:
        _tabla(sv.tabla_energia(est), totales=2)
    with t_m:
        if s:
            _tabla(sv.tabla_maximos(est))
            _grafico(graficos.dibujar_maximos, est, escs[0].pc, figsize=(10, 3.5))
        else:
            st.caption("Potencia máxima anual de cada CUPS por periodo.")
            _tabla(sv.tabla_maximos_cups(sums), totales=0, izquierda=("CUPS",))
    with t_d:
        nombres = [e.nombre for e in escs]
        c1, c2 = st.columns([1, 2])
        nombre = c1.selectbox("Escenario", nombres, key="pot_detalle")
        c2.caption("Importes en € sin impuesto eléctrico.")
        _tabla(sv.tabla_detalle(est, escs[nombres.index(nombre)]), centro=["Mes", "Días"],
               columnas={"Total fijo": COLOR_FIJO, "Total excesos": COLOR_EXCESO,
                         "TOTAL": COLOR_TOTAL})
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
    sums, multipunto = pot["suministros"], _tipo(pot) == MULTIPUNTO
    st.session_state.setdefault("pot_pdf_anexos", True)
    anexos = st.checkbox("Incluir los anexos en el PDF (energía, potencias máximas, coste "
                         "mensual y curva de carga" + (" de cada CUPS" if multipunto else "")
                         + ")", key="pot_pdf_anexos")
    firma = (repr(datos), anexos, seleccion,
             tuple((s.cups, s.zona, s.tarifa, tuple(round(e.coste.total, 2) for e in s.escenarios),
                    s.titular, s.denominacion, s.direccion) for s in sums))
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
        guardar = {"tipo": _tipo(pot), "datos": datos, "anexos": anexos}
        pot["guardado"] = almacen.guardar_estudio_potencia(
            usuario, sv.resumen_guardar(_tipo(pot), datos, sums), guardar,
            sv.empaquetar(sums), pot["pdf"])
    if pot.get("guardado"):
        st.success(f"Estudio guardado con el nº {pot['guardado']}. Está en «Estudio de potencia → "
                   "Estudios guardados».")


def _generar_pdf(pot, datos, res, anexos, seleccion, precios) -> bytes:
    if _tipo(pot) == MULTIPUNTO:
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
                "nuevo y el original no se modifica. Para cambiar los ajustes de lectura de la "
                "curva, empieza un estudio nuevo.")
        if st.button("Empezar un estudio nuevo"):
            _reiniciar()
            st.rerun()
    st.markdown("Optimización de la potencia contratada de suministros **6.1TD** y **3.0TD "
                "> 50 kW** a partir de la curva de carga de 12 meses (Circular CNMC 3/2020): coste "
                "actual, potencia óptima, otras propuestas, inversión y periodo de retorno.")
    precios = sv.precios()
    if not cargado:
        _paso_cups(pot)
    if not pot["suministros"]:
        return
    datos = _paso_datos(pot)
    _paso_curvas(pot)
    _interpretacion(pot)
    res = _propuestas_y_resultados(pot, precios)
    if res is None:
        return
    multipunto = _tipo(pot) == MULTIPUNTO
    seleccion = _anexos(pot, res if multipunto else None)
    _informe_y_guardar(pot, datos, res if multipunto else None, seleccion, precios)


def cargar_en_formulario(estudio: dict) -> None:
    """Deja el estudio guardado listo en «Nuevo estudio» (sin la curva original: se usan las
    curvas ya leidas que se guardaron)."""
    datos = estudio["datos"]
    _reiniciar()
    pot = st.session_state.pot
    d = datos.get("datos", {})
    sums = sv.desempaquetar(estudio["curvas"])
    for s in sums:
        s.fuente = s.fuente or "estudio guardado"
    pot.update(suministros=sums, calculado=True,
               cargado={"id": estudio["id"], "usuario": estudio.get("usuario"),
                        "fecha": _fecha_hora(estudio.get("fecha"))})
    pot["pendiente"].update(pot_titular=d.get("titular", ""), pot_fecha=d.get("fecha", ""),
                            pot_pdf_anexos=bool(datos.get("anexos", True)),
                            pot_cups_texto="\n".join(s.cups for s in sums))


def _fecha_hora(iso):
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
                       file_name=sv.nombre_pdf(datos.get("datos", {}),
                                               datos.get("tipo") == MULTIPUNTO),
                       mime="application/pdf", type="primary", disabled=not estudio["pdf"])
    if b2.button("Cargar en el formulario",
                 help="Abre «Nuevo estudio» con este estudio para cambiar potencias o "
                      "propuestas. Al guardar se crea uno nuevo."):
        cargar_en_formulario(estudio)
        if PAGINA_FORMULARIO is not None:
            st.switch_page(PAGINA_FORMULARIO)
        st.success("Cargado: ve a «Nuevo estudio».")
