# -*- coding: utf-8 -*-
"""
PAGINAS DE GEMWEB de Herramientas GE&PE
========================================
Convertidas de API_Gemweb/app.py: una funcion por operacion (las registra inicio.py),
con el mismo comportamiento y parametros. El cliente de la API es el comun de la
plataforma (gemweb.obtener_api) con las credenciales de gemweb.cargar_credenciales().
"""

import io
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import gemweb
from gemweb_extractor.cliente import (
    GemwebError,
    inventory_a_dataframe,
    consumo_periodos_a_dataframe,
    cost_consumption_a_dataframe,
    power_optimization_a_dataframe,
)


# Opciones que no funcionan del todo por fallos de la API de Gemweb (comprobado contra la API
# real y el manual, oct-2026). inicio.py las marca en el menu y lo explica en la barra lateral;
# quitar la entrada cuando Gemweb lo corrija.
AVISOS_GEMWEB = {
    "Consumo por periodo": "la agrupación **diaria** no funciona: Gemweb devuelve la respuesta "
                           "rota (sin año ni mes). La mensual funciona.",
    "Telelecturas": "las agrupaciones **diaria, mensual, anual y total** pueden salir **a 0 en "
                    "todo el rango** si este incluye ciertos días (p. ej. el 18/02/2026 en "
                    "algunos suministros), aunque haya datos. La **cuartohoraria** y la "
                    "**horaria** son fiables; la descarga masiva de consumos ya las usa.",
    "Coste y consumo mensual": "con origen **telelecturas** Gemweb devuelve la respuesta rota "
                               "(consumo) o un error (coste). El consumo se saca de las "
                               "telelecturas mensuales; el coste, solo desde facturas.",
}


def obtener_cliente():
    """Cliente comun de la API; si no hay credenciales se avisa y se detiene la pagina."""
    cred, _origen = gemweb.cargar_credenciales()
    if cred is None:
        st.warning("No hay credenciales de Gemweb. Un administrador las configura en "
                   "Administración → Credenciales de Gemweb.")
        st.stop()
    return gemweb.obtener_api(*cred)


def _num(v, d=2):
    """1234.5 -> '1.234,50'"""
    try:
        return ("{:,.%df}" % d).format(v).replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(v)


def df_to_excel_bytes(df: pd.DataFrame, nombre_hoja: str = "Datos") -> bytes:
    """Convierte un DataFrame a bytes de Excel para descarga."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    return buffer.getvalue()


def mostrar_resultado(df: pd.DataFrame, nombre_descarga: str) -> None:
    """Muestra un DataFrame con métricas, tabla y botones de descarga."""
    if df.empty:
        st.warning("La consulta no ha devuelto registros.")
        return

    st.success(f"{len(df)} registros obtenidos")

    # Métricas rápidas si hay columnas numéricas relevantes
    cols_num = df.select_dtypes(include="number").columns.tolist()

    # Filtrar columnas técnicas que no aportan como métrica (IDs, índices, etc.)
    cols_excluir = {"id", "id_suministro", "anio", "mes"}
    cols_num = [c for c in cols_num if c.lower() not in cols_excluir]

    # Máximo 8 tarjetas para evitar que la fila se rompa visualmente
    # (cubre tarifas de 6 periodos: P1-P6 + total)
    MAX_METRICAS = 8
    if cols_num:
        cols_mostrar = cols_num[:MAX_METRICAS]
        cols = st.columns(len(cols_mostrar))
        for col, nombre_col in zip(cols, cols_mostrar):
            with col:
                st.metric(
                    nombre_col,
                    f"{_num(df[nombre_col].sum())}",
                    help=f"Suma total · Media: {_num(df[nombre_col].mean())}",
                )

    st.dataframe(df, use_container_width=True, height=400)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Descargar Excel",
            data=df_to_excel_bytes(df),
            file_name=f"{nombre_descarga}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with c2:
        st.download_button(
            "Descargar CSV",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{nombre_descarga}.csv",
            mime="text/csv",
        )



# Helper común para fechas
def selector_fechas(default_dias_atras: int = 90, inicio: date = None, fin: date = None):
    hoy = date.today()
    c1, c2 = st.columns(2)
    with c1:
        f_ini = st.date_input(
            "Fecha inicio",
            value=inicio or hoy - timedelta(days=default_dias_atras),
            format="DD/MM/YYYY",
        )
    with c2:
        f_fin = st.date_input(
            "Fecha fin",
            value=fin or hoy,
            format="DD/MM/YYYY",
        )
    return f_ini.isoformat(), f_fin.isoformat()


def mes_anterior(hoy: date = None):
    """(primer dia, ultimo dia) del mes anterior completo."""
    hoy = hoy or date.today()
    fin = hoy.replace(day=1) - timedelta(days=1)
    return fin.replace(day=1), fin


# Agrupaciones de las telelecturas. Gemweb no calcula el CO2 por debajo de la hora.
PERIODOS_METERING = ["diari", "horari", "quart-horari", "cinc-min", "un-min",
                     "mensual", "anual", "total"]
PERIODOS_SIN_CO2 = ("quart-horari", "cinc-min", "un-min")


def periodos_para(field: str) -> list:
    return [p for p in PERIODOS_METERING if not (field == "co2" and p in PERIODOS_SIN_CO2)]


AVISO_CEROS = ("Todos los valores son 0. Si en Gemweb ves datos para este suministro, es el "
               "fallo de su API en las agrupaciones diaria, mensual, anual y total: pide la "
               "agrupación cuartohoraria u horaria, o un rango más corto.")


def todo_a_cero_sospechoso(df: pd.DataFrame, period: str) -> bool:
    """Respuesta con todos los valores a 0 en una agrupacion que Gemweb calcula mal."""
    return (period in ("diari", "mensual", "anual", "total") and not df.empty
            and "valor" in df.columns and bool((df["valor"] == 0).all()))


def error_api(e: Exception, alternativa: str = None) -> None:
    """Error de Gemweb con la explicacion corta delante y el detalle plegado."""
    texto = str(e)
    if "mal formada" in texto:
        st.error("Gemweb ha devuelto una respuesta rota para esta consulta. Es un fallo de su "
                 "servidor, no de esta aplicación." + (" " + alternativa if alternativa else ""))
        with st.expander("Detalle técnico"):
            st.code(texto, language=None)
    else:
        st.error(f"Error de la API: {texto}" + (" " + alternativa if alternativa else ""))


def errores_tramos(fallidos: list) -> list:
    """Mensajes distintos de los tramos fallidos ('ini → fin: peticion: mensaje')."""
    return list(dict.fromkeys(f.split(": ", 2)[-1] for f in fallidos))


def ejecutar(nombre_operacion: str, fn, parser, archivo: str, alternativa: str = None):
    """Ejecuta la llamada API, parsea y muestra resultados con manejo de errores."""
    with st.spinner("Consultando Gemweb..."):
        try:
            xml = fn()
            df = parser(xml)
            mostrar_resultado(df, archivo)
        except GemwebError as e:
            error_api(e, alternativa)
        except Exception as e:
            st.error(f"Error inesperado: {e}")
            st.exception(e)


# ===== INVENTARIO =====
def pagina_inventario():
    cliente = obtener_cliente()
    st.markdown("Lista entidades, suministros, facturas, etc.")

    # Mapeo nombre amigable → nombre real de la API (en catalán)
    CATEGORIAS = {
        "Suministros": "subministraments",
        "Entidades": "entitats",
        "Centros de consumo": "centres_consum",
        "Instalaciones solares": "instalacions_solars",
        "Facturas": "factures",
    }

    c1, c2 = st.columns([1, 2])
    with c1:
        category_label = st.selectbox("Categoría", list(CATEGORIAS.keys()))
        category = CATEGORIAS[category_label]
    with c2:
        # Filtros sugeridos según la categoría elegida (todos en catalán, que es lo que acepta la API)
        FILTROS_SUGERIDOS = {
            "subministraments": ["", "subministraments.cups", "subministraments.polissa", "subministraments.id"],
            "entitats": ["", "entitats.id", "entitats.pais"],
            "centres_consum": ["", "centres_consum.id"],
            "instalacions_solars": ["", "instalacions_solars.id", "instalacions_solars.polissa"],
            "factures": ["", "subministraments.cups", "factures.id", "factures.num_factura",
                         "factures.data_emissio", "factures.periode_ini", "factures.periode_fin",
                         "factures.fecha_mod"],
        }
        filtro_libre = st.selectbox(
            "Filtrar por (opcional)",
            FILTROS_SUGERIDOS[category],
            help="Elige el campo por el que quieres filtrar. Deja en blanco para listar todo.",
        )

    operador = None
    if filtro_libre:
        f1, f2 = st.columns([1, 2])
        with f1:
            # search_operator del manual: por defecto "="
            operador = st.selectbox("Comparación", ["=", ">", ">=", "<", "<=", "!="],
                                    help="Por ejemplo, «>» con factures.data_emissio = "
                                         "2026-09-01 lista las facturas emitidas después.")
        with f2:
            valor_filtro = st.text_input(
                "Valor del filtro",
                placeholder="Fechas como AAAA-MM-DD" if "data" in filtro_libre or "periode" in
                filtro_libre or "fecha" in filtro_libre else "Ej: ES0000000000000001XX0F")
    else:
        valor_filtro = None

    o1, o2, o3 = st.columns([2, 1, 1])
    with o1:
        orden = st.selectbox("Ordenar por (opcional)",
                             [c for c in FILTROS_SUGERIDOS[category] if c != "subministraments.cups"
                              or category == "subministraments"],
                             help="Por ejemplo, factures.periode_ini descendente para ver primero "
                                  "las facturas más recientes.")
    with o2:
        sentido = st.radio("Sentido", ["ASC", "DESC"], horizontal=True, disabled=not orden,
                           format_func=lambda x: "Ascendente" if x == "ASC" else "Descendente")
    with o3:
        limit = st.number_input("Límite de registros (0 = sin límite)", min_value=0, value=100)

    if st.button("Ejecutar consulta", type="primary"):
        ejecutar(
            "inventario",
            lambda: cliente.get_inventory(
                category=category,
                search_by=filtro_libre or None,
                search_values=valor_filtro or None,
                search_operator=operador if operador and operador != "=" else None,
                order_by=f"{orden} {sentido}" if orden else None,
                limit=limit if limit > 0 else None,
            ),
            inventory_a_dataframe,
            f"inventario_{category}",
        )


# ===== CONSUMO POR PERIODO =====
def pagina_consumo_periodo():
    cliente = obtener_cliente()
    st.markdown("Consumo desagregado por periodos tarifarios (P1, P2, P3…).")
    st.caption("Gemweb saca estos consumos de las **facturas** introducidas, no de las "
               "telelecturas del contador (coinciden al kWh con lo facturado). Para la medida "
               "del contador, usa «Telelecturas».")

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    period = st.radio("Agrupación", ["mensual", "diari"], horizontal=True,
                      format_func=lambda x: "Diaria" if x == "diari" else "Mensual")
    f_ini, f_fin = selector_fechas(default_dias_atras=30)

    if st.button("Ejecutar consulta", type="primary"):
        ejecutar(
            "consumo",
            lambda: cliente.get_consumption_by_period(
                id=id_sum, period=period, date_from=f_ini, date_to=f_fin
            ),
            consumo_periodos_a_dataframe,
            f"consumo_{id_sum}_{f_ini}_{f_fin}",
            alternativa="Prueba con la agrupación mensual o con «Telelecturas» diarias.",
        )


# ===== TELELECTURAS =====
def pagina_telelecturas():
    cliente = obtener_cliente()
    st.markdown("Datos de telelectura (consumo, potencia, reactiva…).")

    c1, c2, c3 = st.columns(3)
    with c1:
        id_sum = st.number_input("ID suministro", min_value=1, step=1)
    with c2:
        field = st.selectbox(
            "Magnitud",
            ["consum", "potencia_activa", "potencia_maxima", "reactiva", "co2"],
        )
    with c3:
        period = st.selectbox(
            "Agrupación temporal", periodos_para(field),
            help="Gemweb no calcula el CO2 cuartohorario, de 5 minutos ni de 1 minuto.",
        )

    data_source = st.radio(
        "Origen", ["comptador", "instantanis"], horizontal=True,
        help="comptador: lecturas oficiales (15min). instantanis: medidas minuto a minuto "
             "(solo si el equipo de telemedida las registra).",
    )
    f_ini, f_fin = selector_fechas(default_dias_atras=15)

    if st.button("Ejecutar consulta", type="primary"):
        with st.spinner("Consultando Gemweb..."):
            aviso = st.empty()
            try:
                df, fallidos = cliente.get_metering_por_tramos(
                    id=id_sum, date_from=f_ini, date_to=f_fin,
                    period=period, field=field, data_source=data_source,
                    al_avanzar=lambda n, t, ini, fin: aviso.caption(
                        f"Tramo {n}/{t}: {ini} → {fin}"),
                )
                aviso.empty()
                if df.empty and fallidos:
                    st.error("Gemweb no ha devuelto datos: " + " · ".join(errores_tramos(fallidos)))
                    return
                for f in fallidos:
                    st.warning(f"Tramo sin descargar: {f}")
                if todo_a_cero_sospechoso(df, period):
                    st.warning(AVISO_CEROS)
                mostrar_resultado(df, f"telelectura_{id_sum}_{field}_{f_ini}_{f_fin}")
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)


# ===== ANÁLISIS COSTE / CONSUMO =====
def pagina_coste_consumo():
    cliente = obtener_cliente()
    st.markdown("Análisis mensual de coste (€) o consumo (kWh).")

    c1, c2 = st.columns(2)
    with c1:
        id_sum = st.number_input("ID suministro", min_value=1, step=1)
    with c2:
        tipo = st.radio("Tipo de análisis", ["consum", "cost"], horizontal=True,
                        format_func=lambda x: "Consumo (kWh)" if x == "consum" else "Coste (€)")

    source = st.radio("Origen de datos", ["factures", "telelectures"], horizontal=True)
    por_fecha_factura = False
    if source == "factures":
        por_fecha_factura = st.checkbox(
            "Asignar cada factura al mes de su fecha de facturación",
            help="Si no se marca, cada factura se reparte según su periodo facturado "
                 "(parámetro billing_date de Gemweb).")
    f_ini, f_fin = selector_fechas(default_dias_atras=365)

    if st.button("Ejecutar consulta", type="primary"):
        with st.spinner("Consultando Gemweb..."):
            try:
                try:
                    xml = cliente.get_cost_consumption(
                        id=id_sum, type=tipo, date_from=f_ini, date_to=f_fin, source=source,
                        billing_date=1 if por_fecha_factura else 0,
                    )
                    df = cost_consumption_a_dataframe(xml)
                except GemwebError:
                    if not (tipo == "consum" and source == "telelectures"):
                        raise
                    df = consumo_mensual_telelecturas(cliente, id_sum, f_ini, f_fin)
                    st.info("Gemweb ha devuelto mal el análisis desde telelecturas: se muestra "
                            "el consumo mensual de las telelecturas (mismo dato, otra consulta).")

                if df.empty:
                    st.warning("La consulta no ha devuelto registros.")
                else:
                    # Etiquetas dinámicas según tipo de análisis
                    es_consumo = (tipo == "consum")
                    unidad = "kWh" if es_consumo else "€"
                    nombre_valor = f"Consumo ({unidad})" if es_consumo else f"Coste ({unidad})"

                    st.success(f"{len(df)} meses obtenidos")

                    # Métricas relevantes para análisis temporal
                    total = df["valor"].sum()
                    media = df["valor"].mean()
                    maximo = df["valor"].max()
                    completitud = df["completitud_%"].iloc[0] if "completitud_%" in df.columns else "—"
                    if completitud is None or pd.isna(completitud):
                        completitud = "—"

                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        st.metric(f"Total ({unidad})", f"{_num(total)}")
                    with c2:
                        st.metric(f"Media mensual ({unidad})", f"{_num(media)}")
                    with c3:
                        st.metric(f"Máximo mensual ({unidad})", f"{_num(maximo)}")
                    with c4:
                        st.metric("Completitud", f"{completitud}%" if completitud != "—" else "—",
                                  help="Porcentaje de datos disponibles en el periodo")

                    # Renombrar columnas para que la tabla sea legible
                    df_vista = df.rename(columns={
                        "id_suministro": "ID suministro",
                        "anio": "Año",
                        "mes": "Mes",
                        "fecha": "Periodo",
                        "valor": nombre_valor,
                        "unidades": "Unidades",
                        "completitud_%": "Completitud (%)",
                    })
                    st.dataframe(df_vista, use_container_width=True, height=400)

                    # Descargas (con columnas ya renombradas)
                    nombre_archivo = f"analisis_{tipo}_{id_sum}_{f_ini}_{f_fin}"
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            "Descargar Excel",
                            data=df_to_excel_bytes(df_vista),
                            file_name=f"{nombre_archivo}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    with c2:
                        st.download_button(
                            "Descargar CSV",
                            data=df_vista.to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"{nombre_archivo}.csv",
                            mime="text/csv",
                        )
            except GemwebError as e:
                error_api(e, "Prueba con el origen «factures»." if source == "telelectures"
                          else None)
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)


def consumo_mensual_telelecturas(cliente, id_sum, f_ini, f_fin) -> pd.DataFrame:
    """Consumo mensual desde get_metering (mismas columnas que cost_consumption_a_dataframe)."""
    df, fallidos = cliente.get_metering_por_tramos(
        id=int(id_sum), date_from=f_ini, date_to=f_fin, period="mensual", field="consum")
    if df.empty:
        if fallidos:
            raise GemwebError(" · ".join(errores_tramos(fallidos)))
        return df
    fecha = df["fecha"].astype(str).str[:7]
    return pd.DataFrame({
        "id_suministro": df["id_suministro"], "anio": fecha.str[:4].astype(int),
        "mes": fecha.str[5:7].astype(int), "fecha": fecha, "valor": df["valor"],
        "unidades": df["unidades"], "completitud_%": None})


# ===== OPTIMIZACIÓN DE POTENCIA =====
def pagina_optimizacion():
    cliente = obtener_cliente()
    st.markdown("Cálculo de las potencias óptimas que habrían minimizado el coste.")
    st.caption("Gemweb la calcula con las lecturas horarias o cuartohorarias de la "
               "distribuidora cargadas en el suministro. Si responde que no dispone de ellas, "
               "falta cargarlas en Gemweb (las telelecturas del contador no le sirven).")

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    f_ini, f_fin = selector_fechas(default_dias_atras=365)

    if st.button("Ejecutar consulta", type="primary"):
        ejecutar(
            "optimizacion_potencia",
            lambda: cliente.get_power_optimization(
                id=id_sum, date_from=f_ini, date_to=f_fin,
            ),
            power_optimization_a_dataframe,
            f"optimizacion_potencia_{id_sum}",
        )


# ===== INCOHERENCIAS =====
def pagina_incoherencias():
    cliente = obtener_cliente()
    st.markdown("Detección de incoherencias en facturas.")

    c1, c2 = st.columns(2)
    with c1:
        category = st.selectbox(
            "Buscar por", ["subministraments", "factures", "instalacions_solars"],
        )
    with c2:
        id_elemento = st.number_input(f"ID de {category}", min_value=1, step=1)

    tipos_dispo = [
        "lectures_anteriors", "lectures_consums", "consum_total",
        "potencia_facturada", "consums_facturats_activa",
        "consums_facturats_reactiva", "imports", "preus",
    ]
    tipos = st.multiselect("Tipos de incoherencia (vacío = todas)", tipos_dispo)
    porcentajes = []
    if tipos:
        # El manual exige un porcentaje (margen de error) por cada tipo elegido, vacío para
        # los que no lo usan; si no: "El número de incoherencias y porcentajes ... no coinciden"
        st.caption("Margen de error de cada tipo (%). Déjalo vacío en los tipos que no usan "
                   "porcentaje, por ejemplo la comparación de lecturas.")
        columnas = st.columns(min(len(tipos), 4))
        for n, tipo in enumerate(tipos):
            with columnas[n % len(columnas)]:
                porcentajes.append(st.number_input(
                    tipo, min_value=0.0, value=None, step=1.0, format="%g",
                    placeholder="vacío", key=f"pct_{tipo}"))

    if st.button("Ejecutar consulta", type="primary"):
        # El parser para incoherencias es más enrevesado; lo hacemos inline
        with st.spinner("Consultando Gemweb..."):
            try:
                xml = cliente.get_invoice_inconsistencies(
                    category=category,
                    id=id_elemento,
                    inconsistencies=",".join(tipos) if tipos else None,
                    percentages=",".join("" if p is None else "%g" % p for p in porcentajes)
                    if tipos else None,
                )
                filas = []
                for fac in xml.findall(".//factura"):
                    fac_id = fac.findtext("id")
                    for inc in fac.findall("inconsistency"):
                        for desc in inc.findall(".//description"):
                            filas.append({
                                "factura_id": fac_id,
                                "tipo": inc.findtext("type"),
                                "porcentaje_%": inc.findtext("percentage"),
                                "descripcion": (desc.text or "").strip(),
                            })
                df = pd.DataFrame(filas)
                if df.empty:
                    st.success("No se han detectado incoherencias.")
                else:
                    mostrar_resultado(df, f"incoherencias_{category}_{id_elemento}")
            except GemwebError as e:
                error_api(e)
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)


# ===== FACTURA CALCULADA =====
def pagina_factura_calculada():
    cliente = obtener_cliente()
    st.markdown(
        "Genera una factura calculada a partir de las telelecturas del periodo. "
        "Devuelve el XML completo de la factura."
    )

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    st.caption("Necesita telelecturas de todo el periodo y que el suministro tenga en Gemweb "
               "un lote con tarifa (precios de energía y potencia) para esas fechas. Los "
               "últimos días suelen no estar cargados todavía: por defecto, el mes anterior.")
    desde_ultima = st.checkbox(
        "Empezar donde acaba la última factura del suministro",
        help="Sin fecha de inicio, Gemweb empieza en la fecha final de la última factura "
             "introducida (o el día 1 del mes actual si no hay ninguna). Útil para anticipar "
             "la próxima factura.")
    if desde_ultima:
        # solo cuenta la fecha final: tiene que ser posterior al final de la ultima factura
        f_fin = st.date_input("Fecha fin", value=date.today() - timedelta(days=3),
                              format="DD/MM/YYYY", key="fin_desde_ultima").isoformat()
        f_ini = None
    else:
        f_ini, f_fin = selector_fechas(*((0,) + mes_anterior()))

    if st.button("Ejecutar consulta", type="primary"):
        with st.spinner("Generando factura..."):
            try:
                xml = cliente.get_calculated_invoice(
                    id=id_sum, date_from=f_ini, date_to=f_fin,
                )
                # La factura tiene una estructura libre; la mostramos como dict plano
                datos = {hijo.tag: hijo.text for hijo in xml if hijo.text}
                df = pd.DataFrame([datos]).T.reset_index()
                df.columns = ["campo", "valor"]
                mostrar_resultado(df, f"factura_calculada_{id_sum}_{f_ini or 'ultima'}_{f_fin}")

                import xml.etree.ElementTree as ET_  # noqa
                with st.expander("Ver XML completo"):
                    st.code(ET_.tostring(xml, encoding="unicode"), language="xml")
            except GemwebError as e:
                error_api(e)
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)


# ===== SIMULACIÓN DE CONTRATO =====
PERIODOS_TARIFA = ["p1", "p2", "p3", "p4", "p5", "p6"]


def campos_simulacion(precios_ene: dict, precios_pot: dict) -> tuple:
    """(contract_fields, contract_values) con los precios informados; los de potencia
    vacios no se envian (Gemweb usa entonces el peaje)."""
    campos, valores = [], []
    for prefijo, precios in (("preu_ene", precios_ene), ("preu_pot", precios_pot)):
        for p in PERIODOS_TARIFA:
            if precios.get(p) is not None:
                campos.append(f"{prefijo}_{p}")
                valores.append("%g" % precios[p])         # decimales con punto
    return ",".join(campos), ",".join(valores)


def simulacion_a_dataframe(root) -> tuple:
    """get_contract_simulation -> (tabla periodo × energía/potencia/reactiva, datos generales:
    fechas de los consumos usados e importes totales)."""
    sim = root.find(".//simulation")
    if sim is None:
        return pd.DataFrame(), {}
    datos = {h.tag: h.text for h in sim}
    filas = []
    for p in PERIODOS_TARIFA:
        fila = {"Periodo": p.upper()}
        for clave, nombre in (("import_ene", "Energía €"), ("import_pot", "Potencia €"),
                              ("import_react", "Reactiva €")):
            v = datos.get(f"{clave}_{p}")
            fila[nombre] = float(v) if v not in (None, "") else None
        if any(fila[c] is not None for c in ("Energía €", "Potencia €", "Reactiva €")):
            filas.append(fila)
    numero = lambda k: float(datos[k]) if datos.get(k) not in (None, "") else None
    generales = {"desde": datos.get("data_ini_sim"), "hasta": datos.get("data_fin_sim"),
                 "energia": numero("import_ene"), "total": numero("import_tot")}
    return pd.DataFrame(filas), generales


def pagina_simulacion_contrato():
    cliente = obtener_cliente()
    st.markdown("Simula la facturación de un suministro con los precios de energía y potencia "
                "que indiques (por ejemplo, para comparar una oferta).")
    st.caption("Gemweb simula a partir de las **facturas** del periodo: si en las fechas "
               "elegidas no hay ninguna factura completa, devuelve 0 €. Si el periodo es "
               "futuro, usa los consumos del mismo periodo del año anterior.")

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    st.subheader("Vigencia del contrato simulado")
    hoy = date.today()
    c1, c2 = st.columns(2)
    with c1:
        c_ini = st.date_input("Inicio del contrato", value=date(hoy.year, 1, 1), format="DD/MM/YYYY")
    with c2:
        c_fin = st.date_input("Fin del contrato", value=date(hoy.year, 12, 31), format="DD/MM/YYYY")
    st.subheader("Periodo a simular")
    f_ini, f_fin = selector_fechas(inicio=date(hoy.year, 1, 1), fin=mes_anterior()[1])

    st.subheader("Precios")
    st.caption("Energía en €/kWh. Potencia opcional: si se deja vacía, Gemweb usa el peaje "
               "que corresponda.")
    precios_ene, precios_pot = {}, {}
    columnas = st.columns(6)
    for col, p in zip(columnas, PERIODOS_TARIFA):
        with col:
            precios_ene[p] = st.number_input(f"Energía {p.upper()}", min_value=0.0, value=None,
                                             format="%.6f", placeholder="vacío", key=f"ene_{p}")
            precios_pot[p] = st.number_input(f"Potencia {p.upper()}", min_value=0.0, value=None,
                                             format="%.6f", placeholder="vacío", key=f"pot_{p}")
    campos, valores = campos_simulacion(precios_ene, precios_pot)

    if st.button("Simular", type="primary", disabled=not campos):
        with st.spinner("Simulando en Gemweb..."):
            try:
                xml = cliente.get_contract_simulation(
                    id=id_sum, contract_date_from=c_ini.isoformat(),
                    contract_date_to=c_fin.isoformat(), contract_fields=campos,
                    contract_values=valores, date_from=f_ini, date_to=f_fin)
                df, gen = simulacion_a_dataframe(xml)
                if gen.get("desde"):
                    st.caption(f"Consumos usados por Gemweb: del {gen['desde']} al "
                               f"{gen['hasta']}.")
                if df.empty and not gen.get("total"):
                    st.warning("Gemweb devuelve 0 € sin desglose: no hay ninguna factura completa "
                               "en el periodo elegido. Amplía las fechas a meses con factura.")
                else:
                    m1, m2 = st.columns(2)
                    m1.metric("Total energía", f"{_num(gen['energia'])} €"
                              if gen.get("energia") is not None else "—")
                    m2.metric("Total simulado", f"{_num(gen['total'])} €"
                              if gen.get("total") is not None else "—")
                    mostrar_resultado(df, f"simulacion_{id_sum}_{f_ini}_{f_fin}")
            except GemwebError as e:
                error_api(e)
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)
    elif not campos:
        st.caption("Indica al menos un precio para simular.")


# ===== DESCARGA MASIVA: CONSUMO MENSUAL POR PERIODOS =====
COLS_P = ["P1 kWh", "P2 kWh", "P3 kWh", "P4 kWh", "P5 kWh", "P6 kWh"]
# facturas que no se cuentan: no son consumo real o duplicarian un mes (abono + rectificativa)
MARCAS_EXCLUIR = {"anulada": "anulada", "simulada": "simulada", "abonament": "abono",
                  "abonada": "abonada"}


def _fecha_gemweb(texto):
    """'31/08/2026' (o '2026-08-31') -> date; None si no se entiende."""
    from datetime import datetime
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime((texto or "").strip()[:10], formato).date()
        except ValueError:
            pass
    return None


def facturas_por_periodo(cups: str, root) -> tuple:
    """Facturas de get_inventory(category=factures) -> (filas validas, nº descartadas)."""
    filas, descartadas = [], 0
    for f in root:
        d = {h.tag: h.text for h in f}
        if any(str(d.get(m) or "0").strip() not in ("0", "") for m in MARCAS_EXCLUIR):
            descartadas += 1
            continue
        ini, fin = _fecha_gemweb(d.get("periode_ini")), _fecha_gemweb(d.get("periode_fin"))
        if fin is None:
            descartadas += 1
            continue
        fila = {"CUPS": cups, "Inicio factura": ini, "Fin factura": fin,
                "Tarifa": d.get("tarifa"), "Comercializadora": d.get("comercialitzadora")}
        for n, col in enumerate(COLS_P, start=1):
            fila[col] = float(d.get(f"consum_p{n}") or 0)
        fila["Total kWh"] = float(d.get("consum_tot") or 0) or sum(fila[c] for c in COLS_P)
        # potencias contratadas de la factura (no salen en las hojas; para la solicitud)
        fila["_potencias"] = [float(d[f"pot_contract_p{n}"]) if d.get(f"pot_contract_p{n}")
                              not in (None, "") else None for n in range(1, 7)]
        # gas: caudal diario aplicado (kWh/dia)
        fila["_qd"] = float(d["qd_aplicada"]) if d.get("qd_aplicada") not in (None, "") \
            else None
        filas.append(fila)
    return filas, descartadas


def _mayoria_dentro(f: dict, desde: date, hasta: date) -> bool:
    """La mayor parte de los dias de la factura cae entre desde y hasta (sin fecha de inicio,
    cuenta su fecha de fin)."""
    fin = f["Fin factura"]
    ini = f.get("Inicio factura") or fin
    dentro = (min(fin, hasta) - max(ini, desde)).days + 1
    return dentro * 2 > (fin - ini).days + 1


def elegir_facturas(filas: list, modo: str, ini: date = None, fin: date = None) -> list:
    """modo 'ultimos12': las facturas del año anterior al final de la ultima (suelen ser 12;
    alguna mas si un mes se facturo en dos, p. ej. por un alta a mitad de mes).
    modo 'periodo': las del periodo de ini a fin (incluidos). En los dos, una factura entra
    solo si la mayor parte de sus dias cae dentro: las fechas de lectura se mueven unos dias y
    una factura que acaba justo dentro pero empezo un mes antes sumaria un mes de mas."""
    filas = sorted(filas, key=lambda f: f["Fin factura"], reverse=True)
    if not filas:
        return []
    if modo == "ultimos12":
        ultima = filas[0]["Fin factura"]
        dia = 28 if (ultima.month, ultima.day) == (2, 29) else ultima.day
        desde, hasta = ultima.replace(year=ultima.year - 1, day=dia) + timedelta(days=1), ultima
    else:
        desde, hasta = ini, fin
    elegidas = [f for f in filas if _mayoria_dentro(f, desde, hasta)]
    return sorted(elegidas, key=lambda f: f["Fin factura"])


def variantes_cups(cups: str) -> list:
    """Gemweb guarda a veces el CUPS de 20 caracteres (sin el '0F' final de 22)."""
    cups = cups.strip().upper()
    return [cups] + ([cups[:20]] if len(cups) > 20 else [])


def meses_naturales(facturas: list, modo: str, ini: date = None, fin: date = None) -> tuple:
    """Rango de meses naturales de las telelecturas: los 12 meses hasta el de la ultima
    factura, o el periodo elegido."""
    if modo != "ultimos12":
        return ini, fin
    if not facturas:
        return None, None
    ultima = facturas[-1]["Fin factura"]
    fin_mes = (ultima.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    anio, mes = ultima.year, ultima.month - 11
    if mes <= 0:
        anio, mes = anio - 1, mes + 12
    return date(anio, mes, 1), fin_mes


def telelecturas_por_mes(cups: str, df_qh: pd.DataFrame) -> list:
    """get_metering (consum, quart-horari, comptador) sumado por mes natural -> filas CUPS |
    Mes | Total kWh | Días con datos | Días del mes | Completa.

    Se suman los cuartos de hora porque las agrupaciones diaria, mensual, anual y total de
    Gemweb devuelven 0 en TODO el rango pedido si incluye ciertos dias (p. ej. el 18/02/2026
    en algunos suministros, aunque ese dia tiene datos cuartohorarios; comprobado oct-2026).
    get_consumption_by_period tampoco sirve: devuelve los consumos de las facturas. Las
    telelecturas solo dan el total, sin reparto P1-P6."""
    if df_qh is None or df_qh.empty:
        return []
    import calendar
    # Gemweb etiqueta cada cuarto con su hora de FIN: el de 23:45-24:00 lleva el dia siguiente
    fin = pd.to_datetime(df_qh["fecha"].astype(str).str.split().str.join(" "), errors="coerce")
    df = df_qh.assign(dia=fin - pd.Timedelta(minutes=15)).dropna(subset=["dia"])
    filas = []
    for mes, grupo in df.groupby(df["dia"].dt.strftime("%Y-%m")):
        anio, m = int(mes[:4]), int(mes[5:7])
        dias_mes = calendar.monthrange(anio, m)[1]
        con_datos = int(grupo.loc[grupo["valor"] != 0, "dia"].dt.date.nunique())
        filas.append({"CUPS": cups, "Mes": mes, "Total kWh": float(grupo["valor"].sum()),
                      "Días con datos": con_datos, "Días del mes": dias_mes,
                      "Completa": "Sí" if con_datos >= dias_mes else "No"})
    return filas


def _pct(dif, base):
    return round(dif / base * 100, 2) if base else None


def diferencias_por_mes(filas_fact: list, filas_mes: list) -> list:
    """Facturado (cada factura en el mes de su fecha de fin) frente a las telelecturas del
    mes natural, por CUPS y mes. Diferencia = facturado - telelecturas, solo en los meses con
    telelectura completa (si faltan dias, la diferencia seria un hueco de datos, no de
    facturacion)."""
    meses = {}
    nuevo = lambda: {"facturas": 0, "fact": 0.0, "nat": None, "completa": None}
    for f in filas_fact:
        d = meses.setdefault((f["CUPS"], f["Fin factura"].strftime("%Y-%m")), nuevo())
        d["facturas"] += 1
        d["fact"] += f["Total kWh"]
    for m in filas_mes:
        d = meses.setdefault((m["CUPS"], m["Mes"]), nuevo())
        d["nat"] = m["Total kWh"] if m.get("Días con datos", 1) else None
        d["completa"] = m.get("Completa", "Sí")
    filas = []
    for (cups, mes), d in sorted(meses.items()):
        fact = d["fact"] if d["facturas"] else None
        comparable = fact is not None and d["nat"] is not None and d["completa"] == "Sí"
        dif = fact - d["nat"] if comparable else None
        nota = ""
        if fact is not None and d["nat"] is None:
            nota = "Sin telelecturas"
        elif fact is not None and d["completa"] == "No":
            nota = "Telelecturas incompletas"
        elif fact is None and d["nat"] is not None:
            nota = "Sin factura"
        filas.append({"CUPS": cups, "Mes": mes, "Facturas": d["facturas"],
                      "Facturado kWh": fact, "Telelecturas kWh": d["nat"],
                      "Diferencia kWh": dif,
                      "Diferencia %": _pct(dif, d["nat"]) if dif is not None else None,
                      "Nota": nota})
    return filas


def excel_consumo_periodos(df_fact, df_mes, df_resumen, df_dif=None) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_fact.to_excel(writer, sheet_name="Facturas por periodo", index=False)
        df_mes.to_excel(writer, sheet_name="Telelecturas por mes", index=False)
        if df_dif is not None:
            df_dif.to_excel(writer, sheet_name="Diferencias", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
        for hoja in writer.sheets.values():
            for col in hoja.columns:
                ancho = max(len(str(c.value or "")) for c in col[:200])
                hoja.column_dimensions[col[0].column_letter].width = min(max(ancho + 2, 10), 40)
            hoja.freeze_panes = "B2"
    return buffer.getvalue()


def elegir_periodo():
    """Radio «Ultimos 12 meses con factura» / «Un periodo concreto» -> (modo, ini, fin)."""
    modo = st.radio(
        "¿Qué facturas?", ["ultimos12", "periodo"], horizontal=True,
        format_func=lambda m: ("Últimos 12 meses con factura" if m == "ultimos12"
                               else "Un periodo concreto"))
    ini = fin = None
    if modo == "ultimos12":
        st.caption("Para cada CUPS se toman las facturas de los últimos 12 meses, contados hacia "
                   "atrás desde el final de su factura más reciente; cada una entra si la mayor "
                   "parte de sus días cae dentro. Si algún mes se facturó en dos facturas, se "
                   "incluyen las dos.")
    else:
        st.caption("Facturas con la mayor parte de sus días dentro del periodo.")
        hoy = date.today()
        ini_s, fin_s = selector_fechas(inicio=date(hoy.year - 1, hoy.month, 1),
                                       fin=mes_anterior()[1])
        ini, fin = date.fromisoformat(ini_s), date.fromisoformat(fin_s)
    return modo, ini, fin


def descarga_consumo_periodos(cliente, cups_list):
    """Opcion «Consumo mensual por periodos» de la descarga masiva."""
    st.subheader("2. Periodo")
    modo, ini, fin = elegir_periodo()
    st.caption("El Excel trae lo **facturado** (una fila por factura, con P1–P6 y su periodo), "
               "las **telelecturas** del contador por mes natural (solo el total: Gemweb no las "
               "reparte por periodos) y sus **diferencias** mes a mes (cada factura cuenta en el "
               "mes de su fecha de fin). No se cuentan las facturas anuladas, simuladas, abonos "
               "ni las abonadas.")

    st.subheader("3. Ejecutar")
    clave = (tuple(cups_list), modo, ini, fin)
    if st.button("Descargar consumos", type="primary", disabled=not cups_list):
        st.session_state.consumo_periodos = dict(
            calcular_consumo_periodos(cliente, cups_list, modo, ini, fin), clave=clave)
    res = st.session_state.get("consumo_periodos")
    if res and res["clave"] == clave:
        mostrar_consumo_periodos(res)


def _primer_registro(cliente, categoria, campo, valor):
    """Primer registro del inventario como dict, o None si no existe."""
    try:
        raiz = cliente.get_inventory(category=categoria, search_by=campo,
                                     search_values=str(valor), limit=1)
    except GemwebError as e:
        if "no se han encontrado" in str(e).lower():
            return None
        raise
    return {h.tag: h.text for h in raiz[0]} if len(raiz) else None


def registro_inventario(cliente, categoria, campo, valor) -> dict:
    """Primer registro del inventario, guardado en la sesion ({} si no existe o falla)."""
    if not valor:
        return {}
    cache = st.session_state.setdefault("cache_inventario", {})
    clave = (categoria, str(valor))
    if clave not in cache:
        try:
            cache[clave] = _primer_registro(cliente, categoria, campo, valor) or {}
        except GemwebError:
            cache[clave] = {}
    return cache[clave]


def suministro_de(cliente, cups) -> dict:
    """Suministro de Gemweb de un CUPS (prueba sus variantes), guardado en la sesion.
    {} si no existe. Los errores de la API se propagan y no se guardan."""
    cache = st.session_state.setdefault("cache_cups_suministro", {})
    if cups not in cache:
        sumi = {}
        for variante in variantes_cups(cups):
            sumi = _primer_registro(cliente, "subministraments", "subministraments.cups",
                                    variante) or {}
            if sumi:
                break
        cache[cups] = sumi
    return cache[cups]


def calcular_consumo_periodos(cliente, cups_list, modo, ini, fin, telelecturas=True,
                              fichas_solicitud=False) -> dict:
    """Facturas, telelecturas y datos del suministro de cada CUPS (con la barra de progreso).
    La solicitud de ofertas no necesita las telelecturas, pero si las fichas."""
    progreso = st.progress(0.0, text="Preparando descarga...")
    filas_fact, filas_mes, resumen, fichas = [], [], [], []
    total = len(cups_list)
    for i, cups in enumerate(cups_list, start=1):
        progreso.progress(i / total, text=f"Procesando {i}/{total}: {cups}")
        estado, elegidas, mensuales, descartadas = "OK", [], [], 0
        validas, hay_facturas, sumi = [], False, None
        try:
            for variante in variantes_cups(cups):
                try:
                    xml_f = cliente.get_inventory(
                        category="factures", search_by="subministraments.cups",
                        search_values=variante, order_by="factures.periode_fin DESC", limit=60)
                except GemwebError as e:
                    if "no se han encontrado" in str(e).lower():
                        continue
                    raise
                if len(xml_f):
                    hay_facturas = True
                    validas, descartadas = facturas_por_periodo(cups, xml_f)
                    break
            elegidas = elegir_facturas(validas, modo, ini, fin)
        except GemwebError as e:
            estado = f"Facturas: {str(e)[:200]}"
        try:
            sumi = suministro_de(cliente, cups)
            id_sum = sumi.get("id") if sumi else None
            m_ini, m_fin = meses_naturales(elegidas, modo, ini, fin)
            if telelecturas and id_sum and m_ini:
                df_tel, fallidos = cliente.get_metering_por_tramos(
                    id=int(id_sum), date_from=m_ini.isoformat(), date_to=m_fin.isoformat(),
                    period="quart-horari", field="consum", data_source="comptador")
                mensuales = telelecturas_por_mes(cups, df_tel)
                if fallidos and not mensuales:
                    raise GemwebError(" · ".join(errores_tramos(fallidos)))
        except GemwebError as e:
            id_sum = sumi.get("id") if sumi else None
            estado = (estado + " · " if estado != "OK" else "") + f"Telelecturas: {str(e)[:200]}"
        if estado == "OK" and id_sum is None and not hay_facturas:
            estado = "CUPS no encontrado en Gemweb"
        elif estado == "OK" and not hay_facturas:
            estado = "El suministro no tiene facturas en Gemweb"
        elif estado == "OK" and not elegidas:
            estado = "Sin facturas en el periodo"
        filas_fact += elegidas
        filas_mes += mensuales
        # la diferencia solo con los meses que tienen factura y telelectura a la vez
        comunes = [d for d in diferencias_por_mes(elegidas, mensuales)
                   if d["Diferencia kWh"] is not None]
        c_fact = sum(d["Facturado kWh"] for d in comunes)
        c_tel = sum(d["Telelecturas kWh"] for d in comunes)
        resumen.append({
            "CUPS": cups, "ID": id_sum, "Facturas": len(elegidas),
            "Desde": elegidas[0]["Inicio factura"] if elegidas else None,
            "Hasta": elegidas[-1]["Fin factura"] if elegidas else None,
            "Total facturado kWh": sum(f["Total kWh"] for f in elegidas),
            "Meses con telelecturas": len(mensuales),
            "Total telelecturas kWh": sum(f["Total kWh"] for f in mensuales),
            "Meses comparados": len(comunes),
            "Facturado en esos meses kWh": c_fact if comunes else None,
            "Telelecturas en esos meses kWh": c_tel if comunes else None,
            "Diferencia kWh": c_fact - c_tel if comunes else None,
            "Diferencia %": _pct(c_fact - c_tel, c_tel) if comunes else None,
            "Facturas descartadas": descartadas, "Estado": estado})
        if fichas_solicitud and (sumi or elegidas):
            fichas.append(ficha_solicitud(cliente, cups, sumi, validas, elegidas))
    progreso.progress(1.0, text="Descarga completada")

    df_fact = pd.DataFrame(filas_fact, columns=["CUPS", "Inicio factura", "Fin factura", "Tarifa",
                                                "Comercializadora"] + COLS_P + ["Total kWh"])
    df_mes = pd.DataFrame(filas_mes, columns=["CUPS", "Mes", "Total kWh", "Días con datos",
                                              "Días del mes", "Completa"])
    df_dif = pd.DataFrame(diferencias_por_mes(filas_fact, filas_mes),
                          columns=["CUPS", "Mes", "Facturas", "Facturado kWh", "Telelecturas kWh",
                                   "Diferencia kWh", "Diferencia %", "Nota"])
    nombre = "consumo_periodos_%s.xlsx" % ("ultimos12" if modo == "ultimos12"
                                            else f"{ini.isoformat()}_{fin.isoformat()}")
    return {"fact": df_fact, "mes": df_mes, "resumen": pd.DataFrame(resumen), "dif": df_dif,
            "nombre": nombre, "fichas": fichas}


def ficha_solicitud(cliente, cups, sumi, validas, elegidas) -> dict:
    """Datos de un suministro para la solicitud de ofertas: Gemweb + potencias de la ultima
    factura + perfil por meses. Razon social y CIF de la entidad; direccion del centro. En gas,
    ademas, el perfil de gas (con el Qd) y los aspectos tecnicos por defecto (Qd de la ultima
    factura que lo trae)."""
    from ofertas import solicitud as so
    sumi = sumi or {}
    entidad = registro_inventario(cliente, "entitats", "entitats.id", sumi.get("id_entitats"))
    centro = registro_inventario(cliente, "centres_consum", "centres_consum.id",
                                 sumi.get("id_centres_consum"))
    ultima = max(validas, key=lambda f: f["Fin factura"]) if validas else None
    potencias = ultima["_potencias"] if ultima else [None] * 6
    if not any(p is not None for p in potencias):          # sin potencias en la factura
        potencias = [float(sumi[f"pot_contract_p{n}"]) if sumi.get(f"pot_contract_p{n}")
                     else None for n in range(1, 7)]
    perfil, aviso = so.perfil_mensual(elegidas)
    tipo = so.producto_de(sumi.get("tipus_submin"))
    gas = {}
    if tipo == so.GAS:
        perfil_g = so.perfil_gas(elegidas)
        con_qd = [f for f in validas if f.get("_qd")]
        qd = max(con_qd, key=lambda f: f["Fin factura"])["_qd"] if con_qd else None
        gas = {"perfil_gas": perfil_g, "tecnicos": so.tecnicos_por_defecto(
            sum(p["kwh"] for p in perfil_g.values()), qd)}
    return {"cups": cups, "alias": so.alias_de(sumi.get("nom")), "tipo": tipo, **gas,
            "grupo": entidad.get("divisio") or "",
            "razon": entidad.get("nom") or "", "cif": so.cif_limpio(entidad.get("cif")),
            "direccion": centro.get("direccio") or entidad.get("direccio") or "",
            "tarifa": sumi.get("tarifa_acces") or (ultima or {}).get("Tarifa") or "",
            "fecha_inicio": "", "potencias": potencias, "perfil": perfil, "aviso": aviso,
            "aviso_grave": so.aviso_perfil(elegidas)[1],
            "facturas": len(elegidas),
            "ultima_factura": ultima["Fin factura"] if ultima else None}


def mostrar_consumo_periodos(res: dict) -> None:
    df_fact, df_mes, df_resumen, df_dif = res["fact"], res["mes"], res["resumen"], res["dif"]
    st.subheader("Resumen")
    st.caption("Diferencia = facturado − telelecturas del contador, sumando solo los meses que "
               "tienen las dos cosas («Meses comparados»). Positiva: "
               "se ha facturado más consumo del que registran las telelecturas.")
    c1, c2, c3 = st.columns(3)
    c1.metric("CUPS con facturas", int((df_resumen["Facturas"] > 0).sum()))
    c2.metric("Facturas", len(df_fact))
    c3.metric("Con problemas", int((df_resumen["Estado"] != "OK").sum()))
    st.dataframe(df_resumen, use_container_width=True, hide_index=True)
    if df_fact.empty and df_mes.empty:
        st.warning("No se ha descargado ningún dato. Revisa el resumen para ver los errores.")
        return
    t1, t2, t3 = st.tabs(["Facturas por periodo", "Telelecturas por mes", "Diferencias"])
    with t1:
        st.dataframe(df_fact, use_container_width=True, hide_index=True, height=350)
    with t2:
        st.dataframe(df_mes, use_container_width=True, hide_index=True, height=350)
    with t3:
        st.caption("Cada factura cuenta en el mes de su fecha de fin. Si el periodo facturado no "
                   "coincide con el mes natural, parte de la diferencia es solo de calendario. "
                   "Los meses con telelecturas incompletas no se comparan (ver «Nota»).")
        st.dataframe(df_dif, use_container_width=True, hide_index=True, height=350)
    st.download_button("Descargar Excel",
                       data=excel_consumo_periodos(df_fact, df_mes, df_resumen, df_dif),
                       file_name=res["nombre"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")


def datos_de_cups(cups_list) -> pd.DataFrame:
    """CUPS con la division (de su entidad) y el nombre del suministro en el inventario."""
    cliente = obtener_cliente()
    filas = []
    with st.spinner("Buscando los CUPS en el inventario de Gemweb..."):
        for cups in cups_list:
            try:
                sumi = suministro_de(cliente, cups)
            except GemwebError as e:
                filas.append({"CUPS": cups, "División": "", "Nombre": f"Error: {str(e)[:120]}"})
                continue
            entidad = registro_inventario(cliente, "entitats", "entitats.id",
                                          sumi.get("id_entitats"))
            filas.append({"CUPS": cups, "División": entidad.get("divisio") or "",
                          "Nombre": sumi.get("nom") or ("" if sumi else
                                                        "No encontrado en Gemweb")})
    return pd.DataFrame(filas, columns=["CUPS", "División", "Nombre"])


def pedir_cups() -> list:
    """CUPS pegados o leidos de un Excel/CSV, sin duplicados y como mucho 150."""
    metodo = st.radio(
        "¿Cómo quieres introducir los CUPS?",
        ["Pegar lista de CUPS", "Subir Excel/CSV"],
        horizontal=True,
    )

    cups_list = []

    if metodo == "Pegar lista de CUPS":
        texto_cups = st.text_area(
            "Pega los CUPS (uno por línea, o separados por comas/espacios)",
            height=140,
            placeholder="ES0000000000000000XX\nES0000000000000001XX0F\n...",
        )
        if texto_cups:
            # Acepta separadores: salto de línea, coma, punto y coma, espacio, tabulador
            import re
            cups_list = [c.strip().upper() for c in re.split(r"[\s,;]+", texto_cups) if c.strip()]

    else:  # Subir fichero
        archivo = st.file_uploader(
            "Sube un Excel (.xlsx) o CSV con una columna que contenga los CUPS",
            type=["xlsx", "csv"],
        )
        if archivo is not None:
            try:
                if archivo.name.endswith(".csv"):
                    df_in = pd.read_csv(archivo, dtype=str)
                else:
                    df_in = pd.read_excel(archivo, dtype=str)

                # Buscar la columna de CUPS automáticamente, o dejar que el usuario elija
                cols_candidatas = [c for c in df_in.columns if "cups" in c.lower()]
                if cols_candidatas:
                    col_cups = st.selectbox(
                        "Columna que contiene los CUPS",
                        cols_candidatas + [c for c in df_in.columns if c not in cols_candidatas],
                    )
                else:
                    col_cups = st.selectbox("Columna que contiene los CUPS", df_in.columns.tolist())

                cups_list = [str(c).strip().upper() for c in df_in[col_cups].dropna().tolist() if str(c).strip()]
                st.caption(f"Fichero leído: {len(df_in)} filas, {len(cups_list)} CUPS válidos")
            except Exception as e:
                st.error(f"No se ha podido leer el fichero: {e}")

    # Eliminar duplicados conservando el orden
    cups_list = list(dict.fromkeys(cups_list))

    if cups_list:
        st.success(f"{len(cups_list)} CUPS únicos preparados para descarga")
        with st.expander(f"Ver los {len(cups_list)} CUPS"):
            st.dataframe(datos_de_cups(cups_list[:150]), hide_index=True,
                         use_container_width=True)

        if len(cups_list) > 150:
            st.error(
                f"Has introducido {len(cups_list)} CUPS. El límite máximo absoluto es 150 "
                "para evitar saturar la API de Gemweb. Se procesarán solo los primeros 150."
            )
            cups_list = cups_list[:150]
        elif len(cups_list) > 80:
            st.warning(
                f"Has introducido {len(cups_list)} CUPS. Es una descarga grande "
                "y puede tardar varios minutos. Se procesará igualmente."
            )
    return cups_list


# ===== DESCARGA MASIVA POR CUPS =====
def pagina_descarga_masiva():
    cliente = obtener_cliente()
    st.markdown(
        "Descarga telelecturas o el **consumo mensual por periodos** de **varios suministros a "
        "la vez** identificándolos por su CUPS. Límite recomendado: 80 CUPS por descarga "
        "(máximo absoluto: 150)."
    )

    st.subheader("1. Suministros a descargar")
    cups_list = pedir_cups()

    que = st.radio("¿Qué quieres descargar?", ["telelecturas", "periodos"], horizontal=True,
                   format_func=lambda x: ("Telelecturas (curva)" if x == "telelecturas"
                                          else "Consumo mensual por periodos (P1–P6)"))
    if que == "periodos":
        descarga_consumo_periodos(cliente, cups_list)
        return

    # --- 2. Parámetros de la descarga ---
    st.subheader("2. Datos a descargar")

    c1, c2 = st.columns(2)
    with c1:
        field = st.selectbox(
            "Magnitud",
            ["consum", "potencia_activa", "potencia_maxima", "reactiva", "co2"],
            help="consum = energía activa en kWh",
        )
    with c2:
        opciones = periodos_para(field)
        period = st.selectbox(
            "Agrupación temporal", opciones,
            # quart-horari por defecto (horari para el CO2, que no es cuartohorario)
            index=opciones.index("quart-horari" if "quart-horari" in opciones else "horari"),
            help="Gemweb no calcula el CO2 cuartohorario, de 5 minutos ni de 1 minuto.",
        )

    data_source = st.radio(
        "Origen", ["comptador", "instantanis"], horizontal=True,
        help="comptador = lecturas oficiales (15 min). instantanis = medidas minuto a minuto "
             "(solo si el equipo de telemedida las registra).",
    )

    # --- 3. Rango de fechas (libre) ---
    st.subheader("3. Rango de fechas")
    f_ini, f_fin = selector_fechas(default_dias_atras=30)

    # --- 4. Ejecución ---
    st.subheader("4. Ejecutar")

    if st.button("Descargar todo", type="primary", disabled=not cups_list):
        # Inicializar caché de CUPS→ID si no existe (persiste durante la sesión)
        if "cache_cups_id" not in st.session_state:
            st.session_state.cache_cups_id = {}

        progreso = st.progress(0.0, text="Preparando descarga...")
        log_box = st.empty()
        log_lineas = []

        def log(msg):
            log_lineas.append(msg)
            log_box.text("\n".join(log_lineas[-8:]))  # Mostrar las últimas 8 líneas

        resultados = []   # Filas para la tabla resumen
        df_global = []    # DataFrames de telelecturas a concatenar

        total = len(cups_list)
        for i, cups in enumerate(cups_list, start=1):
            progreso.progress(i / total, text=f"Procesando {i}/{total}: {cups}")

            # Paso 1: resolver CUPS → ID (con caché)
            id_sum = st.session_state.cache_cups_id.get(cups)
            if id_sum is None:
                try:
                    log(f"[{i}/{total}] Buscando ID de {cups}...")
                    xml_inv = cliente.get_inventory(
                        category="subministraments",
                        search_by="subministraments.cups",
                        search_values=cups,
                        limit=1,
                    )
                    df_inv = inventory_a_dataframe(xml_inv)
                    if df_inv.empty or "id" not in df_inv.columns:
                        resultados.append({
                            "cups": cups, "id": None, "registros": 0,
                            "estado": "CUPS no encontrado en inventario",
                        })
                        continue
                    id_sum = str(df_inv["id"].iloc[0])
                    st.session_state.cache_cups_id[cups] = id_sum
                except GemwebError as e:
                    resultados.append({
                        "cups": cups, "id": None, "registros": 0,
                        "estado": f"Error al buscar CUPS: {e}",
                    })
                    continue
                except Exception as e:
                    resultados.append({
                        "cups": cups, "id": None, "registros": 0,
                        "estado": f"Error inesperado: {e}",
                    })
                    continue

            # Paso 2: pedir telelecturas (partidas en tramos si el rango es largo)
            try:
                def al_avanzar(n, n_tramos, ini, fin, i=i, cups=cups, id_sum=id_sum):
                    log(f"[{i}/{total}] {cups} (ID {id_sum}) · tramo {n}/{n_tramos}: {ini} → {fin}")

                df_met, fallidos = cliente.get_metering_por_tramos(
                    id=int(id_sum),
                    date_from=f_ini, date_to=f_fin,
                    period=period, field=field, data_source=data_source,
                    al_avanzar=al_avanzar,
                )
                if df_met.empty:
                    estado = (f"Error API: {' · '.join(errores_tramos(fallidos))[:300]}"
                              if fallidos else "Sin datos en el periodo")
                    resultados.append({
                        "cups": cups, "id": id_sum, "registros": 0, "estado": estado,
                    })
                else:
                    # Añadir columna CUPS al principio
                    df_met.insert(0, "cups", cups)
                    df_global.append(df_met)
                    estado = "OK"
                    if fallidos:
                        estado = (f"Parcial: fallaron {len(fallidos)} tramo(s) → "
                                  + " | ".join(fallidos)[:300])
                    elif todo_a_cero_sospechoso(df_met, period):
                        estado = ("Todo a 0: probable fallo de Gemweb en esta agrupación; "
                                  "prueba la cuartohoraria")
                    resultados.append({
                        "cups": cups, "id": id_sum, "registros": len(df_met),
                        "estado": estado,
                    })
            except GemwebError as e:
                resultados.append({
                    "cups": cups, "id": id_sum, "registros": 0,
                    "estado": f"Error API: {str(e)[:300]}",
                })
            except Exception as e:
                resultados.append({
                    "cups": cups, "id": id_sum, "registros": 0,
                    "estado": f"Error: {str(e)[:300]}",
                })

        progreso.progress(1.0, text="Descarga completada")

        # --- 5. Mostrar resultados ---
        st.markdown("---")
        st.subheader("Resumen de la descarga")
        df_resumen = pd.DataFrame(resultados)
        n_ok = (df_resumen["estado"] == "OK").sum()
        n_fail = len(df_resumen) - n_ok
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Suministros OK", n_ok)
        with c2: st.metric("Con problemas", n_fail)
        with c3:
            total_filas = sum(len(d) for d in df_global)
            st.metric("Filas totales", f"{_num(total_filas, 0)}")
        st.dataframe(df_resumen, use_container_width=True)

        # --- 6. Resultado final y descarga ---
        if df_global:
            st.subheader("Datos descargados")
            df_final = pd.concat(df_global, ignore_index=True)
            st.dataframe(df_final.head(500), use_container_width=True, height=300)
            if len(df_final) > 500:
                st.caption(f"Mostrando 500 de {_num(len(df_final), 0)} filas. La descarga incluye todo.")

            # Generar Excel con dos hojas: datos + resumen
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_final.to_excel(writer, sheet_name="Telelecturas", index=False)
                df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
            nombre = f"descarga_masiva_{field}_{f_ini}_{f_fin}.xlsx"

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "Descargar Excel completo",
                    data=buffer.getvalue(),
                    file_name=nombre,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with c2:
                st.download_button(
                    "Descargar CSV (solo datos)",
                    data=df_final.to_csv(index=False).encode("utf-8-sig"),
                    file_name=nombre.replace(".xlsx", ".csv"),
                    mime="text/csv",
                )
        else:
            st.warning("No se ha descargado ningún dato. Revisa el resumen para ver los errores.")
