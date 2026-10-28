# Herramientas GE&PE — contexto del proyecto

Plataforma web interna de GE&PE (Streamlit) para el equipo, con acceso por usuario y
contraseña. Secciones:
- **Seguimiento y Control** (primer módulo; su cuadro es la página de inicio): ver abajo.
- **Revisión de SSAA**: revisa el concepto de servicios de ajuste de facturas eléctricas
  (lee la factura en PDF o a mano, aplica la cláusula del contrato con los datos de ESIOS y
  dice si lo facturado es correcto). Viene del proyecto `..\Comprobación_SSAA_facturas`.
- **Gemweb**: consultas a la API de Gemweb (viene de `..\API_Gemweb`).

**Coste cero**: todo corre en local. Nada de Streamlit Cloud, APIs de IA ni servicios de
pago. La factura y la curva no salen del PC.

## Estructura de la plataforma

`inicio.py` es la entrada única: configuración de página, estilo, acceso
(`paginas.acceso`), barra de usuario y menú `st.navigation` por grupos:
- **Seguimiento y Control** (`seguimiento/`, primer grupo del menú; «Cuadro de seguimiento»
  es la página por defecto). Sustituye al Excel combinado de cuadros de seguimiento, al Excel de
  contactos de envío y a los programas de escritorio de envío de informes y de avisos de
  incidencias. **Todos sus datos son confidenciales**: solo en la base de datos local, nunca en
  el código ni en GitHub; las pruebas usan un Excel ficticio (`tests/test_seguimiento.py`).
  - `columnas.py`: columnas de cada tabla (clave, títulos del Excel, tipo, lista) y estados
    con sus colores. Los datos de cada suministro se guardan en JSON con esas claves.
  - `datos.py`: importar los Excel (columnas por título, meses por la fecha de la cabecera; lo
    que no encaja va a «Otros» con su título), DataFrame del cuadro por año, guardar solo lo
    cambiado, bajas, marcar «Inf enviado» y exportar a Excel. Los textos se guardan tal cual
    (solo sin espacios en los extremos): los envíos cruzan por texto exacto.
  - `retraso.py`: la macro `GenerarDashboardRetraso` del Excel de retraso (bloques naturales
    por cadencia, control el día 1 de M+2, solo «Sgto. y Control»). Comprobado con los datos
    reales: mismos bloques y resultados que la macro.
  - `informes.py` e `incidencias.py`: la lógica de los dos programas de envío, igual que ellos
    (comprobado con datos reales: mismas filas del modo prueba). Contactos desde «Contactos de
    envío»; técnicos y auxiliares con su email en la tabla Equipo (Administración).
  - `correo.py`: envío desde la cuenta de Microsoft 365 de quien envía (Microsoft Graph,
    permisos delegados, código de dispositivo, sin secreto; refresh token cifrado con DPAPI en
    `sgc_outlook`). Registro de la aplicación: `servidor/GUIA_OUTLOOK.md`; sus IDs se guardan en
    Administración → «Seguimiento y Control» → Outlook.
  - `estados_gemweb.py`: comprueba que las facturas de Gemweb con periodo de este año (las
    del año anterior, para más adelante) estén señaladas en el cuadro y **propone** los
    cambios; no se aplica nada hasta aceptarlos en el cuadro (todos o los seleccionados), y
    no se pisa una celda cambiada después de la búsqueda. Busca con el botón «Buscar cambios
    en Gemweb» y a las 7:00 y a las 12:00 (tarea del servidor, deja la propuesta en
    `sgc_ajustes`). Las facturas cuentan en los meses que cubren en más de la mitad de sus
    días (sin anuladas, simuladas, abonos ni complementarias). **Único cambio que se recoge:**
    el cuadro dice que no hay factura (vacío o «Sin fra») y en Gemweb sí la hay → «Recibida»;
    ningún otro estado se toca. Además, si la Obs del mes está vacía (o «-»), se proponen las
    siglas de las incidencias de sus facturas: P (`import_pot_exces` > 0), R (`import_react` > 0)
    y LE (`tipus_lectura` = E), p. ej. «P/R»; una Obs escrita a mano no se toca. Al importar
    el Excel se anula la propuesta (cambian los ids). Solo suministros «Sgto. y Control».
    Cambios propuestos y facturas encontradas se ven y se descargan en Excel. Otro
    botón saca los suministros del cuadro cuyo CUPS no está en Gemweb (es normal que haya).
    `--aplicar` aplica todo desde la línea de comandos.
  - `paginas.py` (cuadro, retraso, contactos de envío, contactos y claves de clientes, SMAC,
    telemedida, bajas y administración) y `envios.py` (enviar informes, avisos de incidencias,
    historial y conexión con Outlook). Envíos con modo prueba por defecto; tras enviar informes
    se marca «Inf enviado» en los meses del periodo que estaban en «Recibida», «Pdte inf» o
    «Inf no enviado».
  - Tablas `sgc_*` de `almacen` (suministros, estados por año y mes, tablas auxiliares, listas,
    equipo, ajustes, envíos, cuentas de Outlook y cambios de Gemweb).
  - **Mejoras pendientes** (aplazadas por la usuaria; recordárselas si pregunta por mejoras):
    historial de cambios del cuadro (quién, cuándo y valor anterior) y avisos de incidencias con
    las facturas directamente de la API de Gemweb, sin el Excel y recordando las ya avisadas.
- **Revisión de SSAA**: `app_revision_ssaa.py` (página de fichero, por defecto),
  `paginas.pagina_historial`, `paginas.pagina_fichas`.
- **API de Gemweb**: `gemweb_extractor/paginas.py`, convertido de `API_Gemweb/app.py` (una función por
  operación: inventario, consumo por periodo, telelecturas, descarga masiva por CUPS,
  coste/consumo mensual, optimización de potencia, incoherencias, factura calculada).
- **Solicitud de ofertas**: `ofertas/paginas.py` (consulta Gemweb con las funciones de
  `gemweb_extractor/paginas.py`), `ofertas/solicitud.py` (formato del Excel) y `ofertas/correo.py`
  (texto del correo a las comercializadoras). Sección propia
  porque se le van a añadir más funciones. «Solicitudes guardadas»
  (`ofertas.paginas.pagina_solicitudes_guardadas`): las solicitudes se guardan en la base de
  datos (`almacen`, tabla `solicitudes_ofertas`: Excel, texto, datos del formulario y
  seguimiento); todos las ven y pueden volver a cargarlas en el formulario.
  **Electricidad o gas natural** (desplegable «Producto», preseleccionado con `tipus_submin` de
  Gemweb; una solicitud no mezcla los dos). Gas:
  - El perfil sale de `consum_tot` y `qd_aplicada` de las facturas.
  - Aspectos técnicos editables: Qd, telemedida, término fijo «Cuota fija» (por cliente) /
    «Variable» (por capacidad), presión, cogeneración, planta satélite, días de uso e
    interrumpible.
  - Tipologías MIBGAS (PVB), TTF y Brent, cada una sin y con coberturas.
  - Impuesto sobre Hidrocarburos (reducido, general o exento, con %) y GdO de gas renovable.
  - Reglas normativas en `ofertas/solicitud.py`:
    - telemedida obligatoria y término variable con más de 5 GWh/año (Orden IET/2446/2013);
    - con cogeneración no cabe el tipo reducido del impuesto;
    - avisos: RL.x según el consumo (Circular CNMC 6/2020), factor de carga mayor del 80 %,
      plazos de menos de un año sin telemedida y requisitos del interrumpible.
  - La telemedida de gas no está en Gemweb: Datadis no tiene gas.
- **Estudio de potencia**: optimización de la potencia contratada (6.1TD y 3.0TD > 50 kW)
  con la curva de carga, individual o multipunto.
  - `potencia/` es copia del cálculo del programa de escritorio EstudioPotencia
    (`..\estudio_potencia_cuartohorario`, sin su interfaz tkinter): `calculo`, `calendario`,
    `lector_curva`, `graficos` (matplotlib) e `informe_pdf` (reportlab).
  - Sus precios regulados están en `potencia/precios.json`. **Cada año se actualizan aquí y
    en `config/precios.json` del programa de escritorio**, mientras siga existiendo.
  - `estudio_potencia/servicio.py`: lógica sin Streamlit (curvas subidas o de Gemweb, cálculo,
    tablas, PDF, empaquetar para guardar).
  - `estudio_potencia/paginas.py`: «Nuevo estudio» y «Estudios guardados».
  - Lo primero son los CUPS: con 1 es individual y con varios, multipunto. Se buscan en el
    inventario de Gemweb, que rellena titular (razón social), denominación, dirección,
    tarifa, potencias y zona, y se descarga su curva. Los que no están se completan a mano y
    su curva se sube en un fichero (también para sustituir la de Gemweb).
  - La zona sale del código postal de cada suministro en Gemweb (`servicio.zona_de`: el del
    suministro, el del centro, uno escrito en la dirección o la población) y se puede cambiar
    a mano. Así un multipunto mezcla zonas.
  - Los estudios se guardan en `almacen` (tabla `estudios_potencia`: datos, suministros con
    sus curvas ya leídas en pickle comprimido y PDF), solo en local.
  - Pruebas en `tests/test_potencia.py` (las del programa) y `tests/test_estudio_potencia.py`,
    con curvas sintéticas de `tests/generador_curvas.py`.
- **Supervisión** (`supervision/`, solo supervisores y administradores,
  `almacen.puede_supervisar`): «Panel del equipo», de solo consulta.
  - Filtros: periodo, usuarios y módulos.
  - Indicadores frente al periodo anterior de la misma duración.
  - Resumen por usuario con gráficos por módulo y por mes.
  - Tabla única del trabajo guardado: revisiones, altas y cambios de fichas, solicitudes y su
    seguimiento, estudios y consultas de «Descarga datos API Gemweb» (de la actividad). Al
    pulsar una fila se abre y se descarga su informe, Excel o PDF.
  - Pestaña de actividad. Los intentos de acceso con un usuario que no existe se anotan sin
    usuario (el nombre tecleado va en el detalle), así no aparecen como alguien del equipo.
  - Exportación a Excel con tres hojas (Resumen, Trabajo y Actividad).
  - La lógica está en `supervision/datos.py`, sin Streamlit. Pruebas en
    `tests/test_supervision.py`.
- **Mi trabajo** (`supervision.pagina_mi_trabajo`, para todos): el mismo panel limitado al
  usuario conectado, sin filtro de usuarios ni datos de otros.
  - Se abre con el botón «Mi trabajo» de la barra lateral, encima del nombre.
  - Su página va en la sección sin título (`menu[""]`) y su enlace se oculta por CSS.
- **Equipo de Seguimiento:** la primera importación crea la tabla con los nombres de todas las
  pestañas; al reimportar se respeta tal como se ha dejado en Administración y solo se añaden
  nombres nuevos que lleven suministros activos en el cuadro (`datos._equipo_al_reimportar`).
- **Borrar:** cada usuario puede borrar solo lo que ha creado él (revisiones de SSAA,
  solicitudes de ofertas, estudios de potencia y envíos de Seguimiento), con
  `estilo.borrar_mios` y `almacen.borrar_*` (comprueban el usuario en la consulta).
- **Fichas e historial de SSAA:** la denominación y el titular de cada CUPS salen del cuadro de
  Seguimiento y Control (`seguimiento.datos.nombres_por_cups`).
- **Menú plegable:** cada módulo tiene una flecha.
  - Al entrar, todos están recogidos salvo el de la página abierta.
  - Lo que cada uno recoge o despliega se guarda en el `localStorage` del navegador
    (`gpMenuEstado`).
  - Script en `estilo.traducir`. El manejador del clic se inserta como `<script>` en la
    página, porque lo que registra el iframe de `components.html` deja de funcionar cuando
    Streamlit lo sustituye.
- **Administración** (solo administradores): `paginas.pagina_usuarios`,
  `paginas.pagina_credenciales` (credenciales de Gemweb comunes) y
  `seguimiento.paginas.pagina_admin` (importar los Excel de seguimiento, equipo, listas de
  valores y los IDs de la aplicación de Outlook).
Las páginas de función reciben la cabecera corporativa con `_pagina()` de `inicio.py`, que
además anota en la actividad que se ha abierto la página (una vez por sesión).

**Registro de actividad** (tabla `actividad`). Recoge lo que no queda en otra tabla:
- inicios de sesión y accesos fallidos, en `almacen.comprobar`;
- páginas abiertas;
- consultas a Gemweb, con `registrar_consulta` en `gemweb_extractor/paginas.py`;
- descargas.

**Toda descarga nueva debe usar `estilo.boton_descarga(modulo, contenedor, ...)`** en lugar de
`st.download_button`, para que quede anotada. El detalle es solo texto breve (nombre de
fichero, nº de suministro, periodo). Empezó a registrarse en octubre de 2026.
Arranque: `python lanzador.py --servidor` (puerto 8501) o `streamlit run inicio.py`.

## Uso en equipo (servidor)

Pensado para 9 compañeros en la oficina y por VPN: la app corre en un servidor
(`python lanzador.py --servidor`, puerto 8501, cortafuegos limitado a oficina + VPN) y se
usa desde el navegador. Instalación: `servidor/GUIA_SERVIDOR.md` y
`servidor/instalar_servidor.ps1` (cortafuegos, arranque automático, descargas de ESIOS y
copia diaria de la base de datos).

- `almacen.py`: SQLite `revisor_ssaa.db` en `config.CARPETA_DATOS` (variable
  `SSAA_CARPETA_DATOS`), modo WAL y escrituras con `BEGIN IMMEDIATE` para varios usuarios a
  la vez. Tablas: usuarios (PBKDF2-SHA256 con sal, bloqueo 15 min tras 5 fallos, siempre
  queda un admin), contratos + contratos_historial (quién y cuándo), revisiones (cada
  «Revisar SSAA» con su informe Excel). La primera vez importa `contratos_ssaa.json`.
  Copias: `python almacen.py --copia CARPETA`.
- `paginas.py`: inicio de sesión (la primera vez crea el administrador), cambio de
  contraseña obligatorio con la provisional, Historial de revisiones, Fichas de contrato
  (con sus cambios) y Usuarios (solo administradores: altas con contraseña provisional,
  restablecer, desactivar, admin, supervisor). El rol supervisor (columna `supervisor`,
  añadida por `_migrar`) solo da acceso al panel del equipo.
- Las credenciales de Gemweb solo las ve y cambia un administrador.
- La base de datos y sus copias tienen datos de clientes: no van a git (`*.db`).

## Ejecutable

Ya no se genera `.exe`: la plataforma corre en el PC de la oficina y se usa desde el
navegador. `lanzador.py` conserva el soporte de PyInstaller por si se recupera.

## Imagen corporativa

`estilo.py` (CSS, cabecera con logo, secciones numeradas, pie y traductor de los textos
fijos de Streamlit al español) y `.streamlit/config.toml` (tema: granate del logo #970000,
texto #1E1E1E, fondos beis, Arial; sin menú ni botón «Deploy»). Logo e icono en `assets/`
(copiados de `estudio_potencia_cuartohorario/assets`). El informe Excel usa los mismos
colores, Arial, el logo en «Resumen» y páginas ajustadas al ancho. Toda la interfaz en
español: si se añade un widget con textos propios de Streamlit, añadirlos a
`estilo.TRADUCCIONES` (solo se sustituyen textos que coinciden exactamente).

## Datos

Lee, sin modificarlos, los Excel históricos del proyecto hermano
`..\Descarga_datos_ESIOS` (ruta en `config.CARPETA_ESIOS`, variable `SSAA_CARPETA_ESIOS`):
`AAAA_Historico_PVPC_horario.xlsx` (Total SAH, PERD PVPC) y
`AAAA_Historico_componentes_ESIOS.xlsx` (SSAA QH por componente, `Pérdidas mmm-aa`,
`Pérdidas QH mmm-aa`). Se compara siempre con el último dato publicado que haya en ellos.
Para actualizarlos se usan los scripts de aquel proyecto, no se duplica la descarga.

Claves internas con la convención ESIOS: `(fecha, hora 1..25)` y `(fecha, hora, cuarto)`;
así los días de cambio de hora no necesitan zona horaria.

## Ficheros

| Fichero | Qué hace |
|---|---|
| `inicio.py` | Entrada de la plataforma (acceso, menú, estilo) |
| `app_revision_ssaa.py` | Página «Revisar factura» |
| `ssaa_motor.py` | Cálculo: `Contrato` (índice, agregación, mecanismo, pérdidas, factor), `revisar()`, `diagnostico()` |
| `ssaa_datos_esios.py` | Lectura de los Excel de ESIOS a series |
| `curva_consumo.py` | Lectura de curvas CSV/XLSX/XLS (horaria o QH) y reparto de los días de cambio de hora |
| `gemweb.py` | API de Gemweb (adaptado de `estudio_potencia_cuartohorario/potencia/gemweb.py`): CUPS → id, curva cuartohoraria |
| `lectores_factura/` | `base.py`: `DatosFactura`, `LineaSSAA` y lector genérico (PyMuPDF + regex). Lectores `endesa.py` y `naturgy.py` (grandes cuentas), elegidos por CIF de la comercializadora y registrados en `LECTORES` |
| `graficos.py` | Gráficos Altair en español: serie horaria/QH del índice con la media (SSAA reales) y las referencias del contrato en la leyenda, etiqueta del punto más cercano (fecha, hora y valor); barras de consumo |
| `informe.py` | Excel de revisión y fichas de contrato (`contratos_ssaa.json`, clave CUPS). Hojas: Resumen (factura, cláusula, fórmulas, líneas), «Cálculo N» por línea (parámetros + pasos con fórmula, sustitución, **fórmula viva de Excel** y valor del programa), «Detalle N» (datos horarios; en cláusulas hora a hora cada fila con su fórmula) y Diagnóstico |
| `tests/` | `python -m unittest discover tests`. Incluye las medias reales del Total SAH abr-jun 2026 y una API de Gemweb simulada |
| `facturas_ejemplo/` | Facturas reales de muestra y `esperado.json` con lo que debe leer cada una (fuera de git; lo usa `tests/test_lectores.py`) |
| `revisiones_ssaa/` | Informes guardados `AAAA-MM_CUPS_nºfactura.xlsx` |

## Gemweb

Un solo cliente para toda la plataforma: `gemweb_extractor/cliente.py` (copia de
`API_Gemweb/gemweb_client.py` con un candado al renovar el token). `gemweb.obtener_api()`
devuelve un único objeto por credenciales; `gemweb.ClienteGemweb` (CUPS → suministro y curva
cuartohoraria de la revisión de SSAA) trabaja sobre él. Verificado: mismas respuestas que el
cliente original (inventario y telelecturas).

`https://api.gemweb.es`, XML. Fechas con la hora de FIN del cuarto y siempre 96 cuartos
por día: en marzo lo de las 02:xx se suma a la hora 3 de ESIOS y en octubre se reparte a
medias entre las horas 3 y 4 (`curva_consumo.curva_reloj_a_esios`).

Credenciales, nunca en el repositorio, por este orden: variables `GEMWEB_CLIENT_ID` /
`GEMWEB_CLIENT_SECRET`; `%APPDATA%\ValidadorSSAA\gemweb.json` (guardadas desde la app,
cifradas con DPAPI); `%APPDATA%\EstudioPotencia\gemweb.json`;
`..\API_Gemweb\.streamlit\secrets.toml`.

## GitHub

`origin` = github.com/iyadazig/Herramientas-GE-PE. Se sube cada commit. **Nunca**
datos de clientes (facturas, curvas, CUPS reales, condiciones de contrato, informes) ni
credenciales: están en `.gitignore`. Las pruebas usan CUPS ficticios y bandas genéricas.

## Cálculo paso a paso

`ssaa_motor.formula_clausula(c)` da las fórmulas en texto y `pasos_calculo(r)` los pasos
(clave, fórmula, sustitución, valor). El informe convierte cada clave en una fórmula de
Excel sobre «Detalle N» y los parámetros; verificado recalculando en Excel que coincide
con el programa en techo, banda, media ponderada e indexado cuartohorario.

## Facturas

Una factura puede traer varias líneas de SSAA de periodos distintos (Naturgy regulariza
en la factura de julio los meses de abril, mayo y junio, cada uno con su consumo). Cada
`LineaSSAA` lleva su periodo, kWh, precio €/kWh e importe y se revisa por separado.
Al añadir una factura de muestra, añadir su entrada en `facturas_ejemplo/esperado.json`.
Las pruebas que van a git usan textos sintéticos con la maqueta y cifras inventadas.

## Mecanismos soportados

Regularizaciones sobre un SSAA ya incluido en el precio (las de Endesa grandes cuentas):
- **techo**: cobertura hasta una Referencia de SSAA; cargo = MWh × (SSAA reales − ref.)
  × (1+perd) × 1,015 solo si se supera. Sin abono.
- **banda**: cargo por encima de la ref. superior y abono por debajo de la inferior.

Naturgy — regularización trimestral (banda): Σ meses n [Dif. SSAA n × (1+pérdidas) ×
Ap × Consumo n × HL]. SSAA reales = media aritmética del PFMHORAS_COM (C2_PrecioFinal)
de cada mes; la ficha dice si el contrato **indica la suma de componentes** (por defecto
Restricciones + Procesos OS + Desvíos + REER + Importe participación servicios) o **no
la especifica**. En ambos casos `comprobar_componentes()` prueba todas las combinaciones
de términos del fichero y da la referencia implícita que haría cuadrar cada línea.
Pérdidas estándar 7 % AT (6.xTD) / 17 % BT, Ap 1,02, HL 1,015. Cada mes natural se calcula
aparte (`periodo_calculo="mensual"`: una línea de varios meses se parte; consumo del mes de
la curva o prorrateado por días) y se suma por trimestre natural.
Los resultados de las facturas de muestra no se apuntan aquí (datos de clientes).
El script de descarga guarda la C2 aparte en «PFMHORAS_COM C2 mmm-aa» (no se sustituye
cuando sale la C5); si el contrato fija la C2 se lee esa pestaña, y si falta, la principal
con aviso de la liquidación que tiene.

Además: indexado (+prima), indexado con precio máximo, con mínimo y máximo, y fijo.
Agregación: media aritmética, media ponderada por consumo u hora a hora (QH si curva e
índice son QH). `importe = MWh × precio × (1+perd/100) × factor`.

Cláusulas tipo en `ssaa_motor.PLANTILLAS` (la ficha de cada CUPS guarda cuál es, el tipo
techo/banda y sus referencias; nombres antiguos en `ALIAS_PLANTILLAS`):
- «Endesa grandes cuentas — techo (PVPC)» y «— banda (PVPC)»: Total SAH del
  PVPC_DETALLE_DD, media aritmética, PERD de la tarifa del liquicomún (media aritmética) y
  1,015. Comprobado con una factura real de Endesa 6.2TD: cuadra al céntimo.
- «Endesa grandes cuentas — banda (componentes OS, C2)»: NO usa el PVPC. SSAA reales =
  media aritmética de RT3 [806] + RT6 [807] + BS3 [811] + BALX [1368] + CFP [1286] del
  Excel de componentes en liquidación C2 (copia «Componentes C2 mmm-aa»), cálculo por mes,
  mismas pérdidas y 1,015. Si la factura es anterior a la publicación de la C2, la app avisa
  de que Endesa usó una publicación anterior y debe regularizar en la factura siguiente
  (`revisar(..., fecha_emision=)`). Los meses que ya pasaron a C3+ antes de guardar la copia
  (ene–jun 2026) solo tienen la última liquidación: se avisa.

## Pendiente

- Más comercializadoras y tipologías de factura según lleguen muestras.
- Condiciones de SSAA de los contratos de las dos facturas de muestra: con variantes
  simples (SAH/SSAA ESIOS, pérdidas, 1,015) no sale una referencia constante en Naturgy.
- El 30/09/2026 del Excel de componentes (avance A2) trae valores imposibles (Total SSAA
  ~588 €/MWh, RT3 ~21.645): el motor lo avisa como anómalo.
