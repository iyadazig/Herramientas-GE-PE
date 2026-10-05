# -*- coding: utf-8 -*-
"""
PAGINAS DE GEMWEB de Herramientas GE&PE
========================================
Convertidas de API_Gemweb/app.py: una funcion por operacion (las registra inicio.py),
con el mismo comportamiento y parametros. El cliente de la API es el comun de la
plataforma (gemweb.obtener_api) con las credenciales de gemweb.cargar_credenciales().
"""

import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import gemweb
from gemweb_extractor.cliente import (
    GemwebError,
    inventory_a_dataframe,
    consumo_periodos_a_dataframe,
    metering_a_dataframe,
    cost_consumption_a_dataframe,
    power_optimization_a_dataframe,
)


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
def selector_fechas(default_dias_atras: int = 90):
    hoy = date.today()
    c1, c2 = st.columns(2)
    with c1:
        f_ini = st.date_input(
            "Fecha inicio",
            value=hoy - timedelta(days=default_dias_atras),
            format="DD/MM/YYYY",
        )
    with c2:
        f_fin = st.date_input(
            "Fecha fin",
            value=hoy,
            format="DD/MM/YYYY",
        )
    return f_ini.isoformat(), f_fin.isoformat()


def ejecutar(nombre_operacion: str, fn, parser, archivo: str):
    """Ejecuta la llamada API, parsea y muestra resultados con manejo de errores."""
    with st.spinner(f"Consultando Gemweb..."):
        try:
            xml = fn()
            df = parser(xml)
            mostrar_resultado(df, archivo)
        except GemwebError as e:
            st.error(f"Error de la API: {e}")
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
            "factures": ["", "subministraments.cups", "factures.num_factura",
                         "factures.data_emissio", "factures.periode_ini", "factures.periode_fin"],
        }
        filtro_libre = st.selectbox(
            "Filtrar por (opcional)",
            FILTROS_SUGERIDOS[category],
            help="Elige el campo por el que quieres filtrar. Deja en blanco para listar todo.",
        )

    if filtro_libre:
        valor_filtro = st.text_input("Valor del filtro", placeholder="Ej: ES0000000000000001XX0F")
    else:
        valor_filtro = None

    limit = st.number_input("Límite de registros (0 = sin límite)", min_value=0, value=100)

    if st.button("Ejecutar consulta", type="primary"):
        ejecutar(
            "inventario",
            lambda: cliente.get_inventory(
                category=category,
                search_by=filtro_libre or None,
                search_values=valor_filtro or None,
                limit=limit if limit > 0 else None,
            ),
            inventory_a_dataframe,
            f"inventario_{category}",
        )


# ===== CONSUMO POR PERIODO =====
def pagina_consumo_periodo():
    cliente = obtener_cliente()
    st.markdown("Consumo desagregado por periodos tarifarios (P1, P2, P3…).")

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    period = st.radio("Agrupación", ["diari", "mensual"], horizontal=True,
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
            "Agrupación temporal",
            ["diari", "horari", "quart-horari", "cinc-min", "un-min",
             "mensual", "anual", "total"],
        )

    data_source = st.radio(
        "Origen", ["comptador", "instantanis"], horizontal=True,
        help="comptador: lecturas oficiales (15min). instantanis: medidas minuto a minuto.",
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
                for f in fallidos:
                    st.warning(f"Tramo sin descargar: {f}")
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
    f_ini, f_fin = selector_fechas(default_dias_atras=365)

    if st.button("Ejecutar consulta", type="primary"):
        with st.spinner("Consultando Gemweb..."):
            try:
                xml = cliente.get_cost_consumption(
                    id=id_sum, type=tipo, date_from=f_ini, date_to=f_fin, source=source,
                )
                df = cost_consumption_a_dataframe(xml)

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
                st.error(f"Error de la API: {e}")
            except Exception as e:
                st.error(f"Error inesperado: {e}")
                st.exception(e)


# ===== OPTIMIZACIÓN DE POTENCIA =====
def pagina_optimizacion():
    cliente = obtener_cliente()
    st.markdown("Cálculo de las potencias óptimas que habrían minimizado el coste.")

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

    if st.button("Ejecutar consulta", type="primary"):
        # El parser para incoherencias es más enrevesado; lo hacemos inline
        with st.spinner("Consultando Gemweb..."):
            try:
                xml = cliente.get_invoice_inconsistencies(
                    category=category,
                    id=id_elemento,
                    inconsistencies=",".join(tipos) if tipos else None,
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
                st.error(f"Error de la API: {e}")


# ===== FACTURA CALCULADA =====
def pagina_factura_calculada():
    cliente = obtener_cliente()
    st.markdown(
        "Genera una factura calculada a partir de las telelecturas del periodo. "
        "Devuelve el XML completo de la factura."
    )

    id_sum = st.number_input("ID suministro", min_value=1, step=1)
    f_ini, f_fin = selector_fechas(default_dias_atras=30)

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
                mostrar_resultado(df, f"factura_calculada_{id_sum}_{f_ini}_{f_fin}")

                import xml.etree.ElementTree as ET_  # noqa
                with st.expander("Ver XML completo"):
                    st.code(ET_.tostring(xml, encoding="unicode"), language="xml")
            except GemwebError as e:
                st.error(f"Error de la API: {e}")


# ===== DESCARGA MASIVA POR CUPS =====
def pagina_descarga_masiva():
    cliente = obtener_cliente()
    st.markdown(
        "Descarga telelecturas de **varios suministros a la vez** identificándolos por su CUPS. "
        "Ideal para extracciones mensuales o por rango libre. Límite recomendado: 80 CUPS por descarga (máximo absoluto: 150)."
    )

    # --- 1. Recogida de CUPS: dos métodos ---
    st.subheader("1. Suministros a descargar")
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
            st.code("\n".join(cups_list))

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
        period = st.selectbox(
            "Agrupación temporal",
            ["diari", "horari", "quart-horari", "cinc-min", "un-min",
             "mensual", "anual", "total"],
            index=2,  # quart-horari por defecto
        )

    data_source = st.radio(
        "Origen", ["comptador", "instantanis"], horizontal=True,
        help="comptador = lecturas oficiales (15 min). instantanis = medidas minuto a minuto.",
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
                    estado = (f"Error API: {fallidos[0][:300]}" if fallidos
                              else "Sin datos en el periodo")
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
