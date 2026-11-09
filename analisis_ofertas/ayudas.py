# -*- coding: utf-8 -*-
"""
AYUDAS DE «OFERTAS» (textos del ? de cada campo, avisos de coherencia y resumen de la fórmula)
============================================================================================
AYUDA[clave]           texto (markdown) que aparece al pasar el ratón por el ? del campo
avisos_coherencia(o)   combinaciones de la ficha de una oferta que suelen ser un error
resumen_formula(o)     la fórmula de la oferta en una línea, con sus valores
GUIA                   pasos para meter una oferta y dónde buscar cada dato en el PDF
GUIA_GRUPOS            explicación de los grupos de precios propios y de «Incluido»
GUIA_CIERRES           explicación de los cierres en OMIP (por % y por MW base y pico)
Sin Streamlit.
"""

from analisis_ofertas import modelo as mo

AYUDA = {
    # ----------------------------------------------------------------- ficha de la oferta
    "modalidad": (
        "**Fijo**: un precio de energía por periodo que no cambia con el mercado.  \n"
        "**Indexado Pass-Pool**: precio = coeficiente (Ci) × precio de OMIE + término B por "
        "periodo. Los servicios de ajuste suelen ir dentro del término B.  \n"
        "**Indexado Pass-Through**: se paga el precio de OMIE con todos sus costes (servicios de "
        "ajuste, pagos por capacidad, OS, OM…) a coste real, con pérdidas, más un fee de "
        "gestión."),
    "duracion": "Meses del contrato (cualquier número: 12, 24, 33, 36…). No tiene que coincidir "
                "con las duraciones pedidas en la solicitud.",
    "ronda": "1.ª ronda: la oferta inicial. Usa «Nueva ronda» para copiar la oferta cuando la "
             "comercializadora la mejore y comparar las rondas en Resultados.",
    "inicio": "Márcalo solo si la oferta empieza en otro mes que el resto del estudio.",
    "unidad": "Unidad en que la comercializadora da el precio fijo de la energía. La aplicación "
              "lo convierte: 0,1520 €/kWh = 15,20 c€/kWh = 152,0 €/MWh.",
    "descuento": "Solo en precio fijo: % de descuento sobre el término de energía.",
    "atr": "Márcalo si el precio de energía ya incluye los peajes y cargos de energía (el ATR). "
           "Si no, se suman los del BOE de cada año. Fíjate en el PDF: «con ATR», «ATR "
           "incluido» o «sin ATR / peajes aparte».",
    "potencia": "**Regulada (ATR)**: el término de potencia del BOE (peajes y cargos), lo "
                "habitual.  \n**Propia de la oferta**: la comercializadora da su propio precio "
                "de potencia (€/kW·año); se escribe en la pestaña Precios.",
    "pc": "Pagos por capacidad (€/kWh regulados por periodo). Márcalo si la fórmula de la oferta "
          "los suma aparte (PC, PPC, «pagos por capacidad»). Si van dentro del término B o del "
          "precio, déjalo sin marcar.",
    "os": "Retribución al Operador del Sistema (REE), unos 0,2 €/MWh. Márcalo si la fórmula la "
          "suma aparte (OS, R_OS).",
    "om": "Retribución al Operador del Mercado (OMIE), unos 0,04 €/MWh. Márcalo si la fórmula "
          "la suma aparte (OM, R_OMIE).",
    "srad": "Servicio de respuesta activa de la demanda (sustituye a la interrumpibilidad). "
            "Márcalo si la fórmula lo suma aparte (SRAD, «Int» o «interrumpibilidad»).",
    "base_b": "**Consumo neto**: el término B se cobra sobre el consumo del contador.  \n"
              "**Barras central**: se cobra sobre el consumo con pérdidas (el término B queda "
              "dentro del paréntesis que se multiplica por (1 + pérdidas)).",
    "fee": "Coste de gestión de la comercializadora en €/MWh (fee, CG, FOMIE…). Si el PDF lo da "
           "en c€/kWh, multiplica por 10 (0,25 c€/kWh = 2,5 €/MWh).",
    "base_fee": "**Consumo neto**: el fee se cobra sobre el consumo del contador.  \n**Barras "
                "central**: el fee va dentro del paréntesis de las pérdidas.",
    "gdo": "Precio de las garantías de origen (energía renovable) en €/MWh. Ponlo solo si la "
           "oferta las cobra aparte; si ya van en el precio, deja 0.",
    "fnee": "Fondo Nacional de Eficiencia Energética (unos 2,4-2,7 €/MWh).  \n**Incluido**: ya va "
            "en el precio o en el fee.  \n**Consumo neto**: se cobra aparte sobre el consumo.  \n"
            "**Barras central**: se cobra aparte con pérdidas.",
    "k": "Impuesto municipal o «K» (tasa municipal por ocupación del dominio público): un % "
         "sobre la energía y los otros conceptos. Si la fórmula multiplica por 1,015, pon 1,5. "
         "En precio fijo normalmente va incluida en el precio: deja 0 salvo que la fórmula la "
         "aplique aparte.",
    "k_ssaa": "Márcalo si la K (× 1,015) también multiplica a los servicios de ajuste.",
    "otros_fijos": "Cargos fijos al mes por CUPS que cobre la oferta (servicios, alquileres no "
                   "regulados…). El alquiler del contador va en la tabla de suministros.",
    "desvios": "Coste de desvíos en €/MWh, en barras (con pérdidas), si la oferta lo cobra aparte "
               "de los servicios de ajuste.",
    "ssaa_modo": (
        "Cómo cobra la oferta los servicios de ajuste (restricciones, banda secundaria, desvíos "
        "del sistema…):  \n"
        "**Incluidos**: ya van en el precio (fijo o término B). No se suma nada.  \n"
        "**Indexado por periodo**: se paga el SSAA real del mes, más caro en punta y más barato "
        "en valle (apuntamiento de SSAA). Lo habitual en Pass-Through.  \n"
        "**Indexado plano**: el SSAA del mes, igual en todos los periodos.  \n"
        "**Banda**: incluidos entre un mínimo y un máximo; se paga lo que supere el máximo y se "
        "devuelve lo que quede por debajo del mínimo.  \n"
        "**Techo**: incluidos hasta un umbral; se paga solo lo que lo supere."),
    "banda_min": "Referencia inferior de la banda (€/MWh). Si el SSAA real baja de aquí, la "
                 "comercializadora devuelve la diferencia.",
    "banda_max": "Referencia superior de la banda (€/MWh). Si el SSAA real la supera, se paga la "
                 "diferencia.",
    "techo": "Umbral (€/MWh) hasta el que los SSAA van incluidos; se paga solo lo que lo supere.",
    "ssaa_perfil": "Compara la banda o el techo con el SSAA de cada periodo (apuntado) en vez de "
                   "con la media del mes.",
    "ssaa_perdidas": "**Pérdidas tarifa**: los SSAA se multiplican por (1 + pérdidas) de cada "
                     "periodo, lo habitual.  \n**% fijo**: un % que da la oferta.  \n**No**: sin "
                     "pérdidas.",
    "ssaa_f": "Factores extra que la fórmula aplica solo a los SSAA (por ejemplo, un "
              "apuntamiento de 1,02). 1 = sin factor. Si ya has marcado «La K se aplica a los "
              "SSAA», no pongas aquí el 1,015: se cobraría dos veces.",
    "cierres": "Si la oferta permite fijar precio en OMIP:  \n**Por %**: un % del consumo de un "
               "mes a un precio cerrado.  \n**Por MW (carga base)**: MW cerrados las 24 h (y, "
               "si se quiere, MW de carga pico 8-20 h L-V); se liquida la diferencia con el "
               "precio previsto.",
    "click": "Lo que cobra la comercializadora por cada MWh cerrado (trade fee o prima de cierre).",
    "cat_nombre": "Nombre con el que se busca en el catálogo: comercializadora, producto y "
                  "mes del catálogo (p. ej. «Endesa Empresas Fijo nov-26»).",
    "cat_vigencia": "Fechas en las que la comercializadora admite contratar esta oferta (validez "
                    "del catálogo). Fuera de ellas sale como caducada.",
    "cat_inicio": "Fechas entre las que puede empezar el suministro con estos precios. Al "
                  "cargarla se avisa si el inicio del análisis cae fuera.",
    "ph": "Precio horario de adquisición de la demanda en Canarias (Orden TEC/1172/2018, ESIOS). "
          "Ya incluye los costes de ajuste: en los CUPS de Canarias el mercado es (previsto + "
          "Ph − diario) × apuntamiento del Ph. No añadas SSAA si la oferta no los cobra aparte.",
    "cimm": "Precio de mercado mensual × Ci, con el Ci que ya incluye el perfil (y las "
            "pérdidas): no se aplica el apuntamiento. Desmárcalo si la oferta aplica el Ci "
            "sobre el precio horario (ponderado por el consumo).",
    # ----------------------------------------------------------------- resto del análisis
    "origen": "**Solicitud guardada**: toma los CUPS, potencias y consumos de una solicitud de "
              "ofertas de la aplicación.  \n**Plantilla Excel**: para suministros que no están "
              "en una solicitud; se descarga, se rellena y se sube.",
    "crecimiento": "Variación del consumo de cada año de contrato frente al primero (año 2, año "
                   "3…). Por ejemplo «3, 2» = el año 2 consume un 3 % más y el año 3 un 2 % más "
                   "que el 2.",
    "actualizar_mercado": "Vuelve a leer la última sesión de OMIP y las medias de OMIE para "
                          "todos los meses del horizonte. El estudio usa los datos congelados al "
                          "crearlo, para que los resultados no cambien solos.",
    "apuntamiento": "Relación entre el precio de cada periodo y el precio medio del mes. Un "
                    "consumo concentrado en horas caras (punta) paga más que el precio medio. "
                    "**Estimado**: el de los 12 últimos meses del mercado diario. **Del "
                    "cliente**: el calculado con su curva horaria de un año completo (≥ 360 días "
                    "y los 12 meses).",
    "estado": "Estado del análisis. Al guardarlo como «Adjudicado» con la oferta elegida, la "
              "solicitud de origen pasa también a «Adjudicada».",
    "adjudicada": "La oferta que el cliente ha elegido.",
    "referencia": "Texto que aparece en los informes y en el correo (por ejemplo la referencia "
                  "de la solicitud).",
    "escenarios": "Recalcula las ofertas indexadas y el contrato actual con el OMIP de todos los "
                  "meses subido o bajado un %. Sirve para ver el riesgo de cada oferta si el "
                  "mercado se mueve.",
    "mismo_periodo": "El coste del contrato actual en exactamente los mismos meses que la oferta "
                     "(si el actual dura menos, se repiten sus precios). Así se compara lo que se "
                     "pagaría manteniéndolo.",
    # ----------------------------------------------------------------- solicitud de ofertas
    "sol_indexado": (
        "**Pass Through**: OMIE + todos los costes del mercado a coste real + fee.  \n"
        "**Pass Pool**: OMIE × coeficiente + término fijo que incluye los servicios de ajuste.  \n"
        "**+ Omip (carga base)**: permite fijar precio en futuros por MW de carga base.  \n"
        "**+ Omip (% consumo)**: permite fijar precio de un % del consumo.  \n"
        "**PVPC**: tarifa regulada (solo 2.0TD de hasta 10 kW)."),
    "sol_duracion": "Fecha fin de cada alternativa de duración que se pide (o meses desde el "
                    "inicio). Se pueden pedir varias.",
    "sol_gdo": "Pide que el precio incluya garantías de origen (electricidad 100 % renovable).",
    "sol_exencion": "Si el suministro tiene exención o reducción del impuesto eléctrico (usos "
                    "del art. 98.1 de la Ley 38/1992 o del art. 94): % del consumo que está "
                    "exento y % de la exención.",
    "sol_recepcion": "Plazo en que las comercializadoras deben enviar las ofertas. La fecha "
                     "límite aparece en el Excel y en el correo de la solicitud.",
}

# ayuda de cada tipología de precio de la solicitud de ofertas
TIPOLOGIAS = {
    "PVPC": "Precio regulado (Precio Voluntario para el Pequeño Consumidor): solo 2.0TD de "
            "hasta 10 kW y con una comercializadora de referencia.",
    "Indexada Omie Pass Through": "OMIE horario + todos los costes del mercado a coste real "
                                  "(servicios de ajuste, capacidad, OS, OM…) con pérdidas, más "
                                  "un fee de gestión.",
    "Indexada Omie Pass Pool": "OMIE × coeficiente por periodo + un término fijo por periodo "
                               "que suele incluir los servicios de ajuste y la gestión.",
    "Indexada Omie Pass Through + Omip (carga base)": "Pass Through que además permite fijar "
                                                      "precio en futuros (OMIP) por MW de "
                                                      "carga base (las 24 h).",
    "Indexada Omie Pass Through + Omip (% consumo)": "Pass Through que además permite fijar "
                                                     "precio de un % del consumo de un periodo "
                                                     "(clicks).",
    "Indexada Omie Pass Pool + Omip (carga base)": "Pass Pool que además permite fijar precio "
                                                   "en futuros (OMIP) por MW de carga base.",
    "Indexada Omie Pass Pool + Omip (% consumo)": "Pass Pool que además permite fijar precio "
                                                  "de un % del consumo (clicks).",
}

GUIA = """\
1. **Comercializadora, gestor y email**: al elegir la comercializadora se proponen el gestor y
   el email de la última oferta suya.
2. **Modalidad y duración**: «Precio fijo», «Indexado Pass-Through» o «Pass-Pool» suele estar
   en el título o en las condiciones económicas; la duración, en las fechas de inicio y fin.
3. **¿Con o sin ATR?**: busca «peajes y cargos», «ATR» o «tarifa de acceso» junto a los
   precios.
4. **Precios** (pestaña Precios): fijo, el precio de cada periodo P1-P6; Pass-Pool, el
   coeficiente (Ci, Di, B…) y el término fijo (A, CMF, Bi…); Pass-Through, el fee.
5. **Qué suma aparte la fórmula**: marca pagos por capacidad, OS, OM, SRAD, desvíos, FNEE, GdO
   y K solo si la fórmula los nombra fuera del precio.
6. **Servicios de ajuste**: incluidos, a coste real (por periodo) o en banda o techo.
7. **Condiciones**: validez, pago, garantías, volumen, rescisión y GdO (pestaña Condiciones).
8. Revisa la línea **«Fórmula de esta oferta»** y los avisos: deben coincidir con la fórmula
   del PDF.
"""


GUIA_GRUPOS = """\
**¿Qué es un grupo de precios?** Es el conjunto de CUPS que pagan **los mismos precios** en
esta oferta. La aplicación crea sola un grupo por cada **tarifa y sistema** de los suministros
(por ejemplo «6.1TD», «3.0TD», «3.0TD-Canarias») y cada CUPS va a su grupo automáticamente. Los
precios de cada grupo se meten en la pestaña **Precios**.

**¿Cuándo hace falta un grupo propio?** Solo cuando la comercializadora da a **algunos CUPS un
precio distinto** del resto de su misma tarifa. Casos típicos:
- un CUPS grande (o con un perfil muy concreto) con precio especial;
- la oferta separa precios por zona, sociedad del grupo o tipo de centro (oficinas, naves…);
- precios distintos según el consumo (por ejemplo, los 3.0TD de más de 100 MWh).

**Cómo se hace, paso a paso**
1. En la tabla **Grupos propios** añade una fila: **Grupo** (un nombre corto, por ejemplo
   «3.0TD grandes»), **Tarifa** y **Sistema**. La tarifa y el sistema deben ser los de los CUPS
   que vas a meter en él: un grupo de 3.0TD solo puede llevar CUPS 3.0TD.
2. En la tabla de **CUPS**, columna **Grupo propio**, elige ese grupo en los CUPS afectados. Los
   que dejes en blanco siguen en el grupo de su tarifa.
3. Ve a **Precios**: aparece una fila nueva para el grupo propio; mete ahí sus precios (por año
   de contrato, igual que los demás).

**Columna «Incluido»**: desmárcala en los CUPS que **no entran en esta oferta** (la
comercializadora no los oferta o se han quitado del concurso). Esos CUPS no suman coste en
esta oferta, pero el contrato actual sí los incluye: **la comparación deja de ser homogénea**
(la oferta parece más barata de lo que es). Úsalo solo si de verdad se van a quedar fuera, y
coméntalo al cliente.

**Ejemplo**: 20 CUPS 3.0TD; la oferta da 0,125 €/kWh en P1 a todos salvo a los 3 hoteles, a
0,118 €/kWh. Crea el grupo «3.0TD hoteles» (3.0TD, Península), asígnalo a los 3 hoteles y mete
sus precios: los otros 17 siguen con los del grupo «3.0TD».

Si un CUPS tiene un grupo con otra tarifa u otro sistema, aparece en **Controles** y no se
calcula hasta que se corrija.
"""

GUIA_CIERRES = """\
**¿Qué es un cierre?** En una oferta indexada el precio de la energía sigue al mercado diario
(OMIE). Un **cierre** fija hoy, en el mercado de futuros (**OMIP**), el precio de una parte de
la energía de un mes: esa parte deja de depender de OMIE. Sirve para quitar riesgo o
aprovechar un buen precio de futuros. La comercializadora cobra un **coste por click** (€ por
MWh cerrado) que se mete en Configuración.

**Qué muestra la tabla**: una fila por **mes de contrato**, con el **Previsto** (el precio de
mercado que usa el estudio para ese mes: % OMIP + % media de OMIE) y el **OMIP** de la última
sesión. Solo rellenas los meses que quieras cerrar; los demás quedan a mercado.

---
**Cierres por % del consumo** (lo habitual en Pass-Pool y en muchas Pass-Through)
- **% cerrado**: parte del consumo del mes que se cierra, **en tanto por uno** (0,5 = 50 %).
- **Precio de cierre**: el precio pactado. **Vacío = OMIP del mes** (lo normal para simular).
- Cálculo del mercado del mes:

  `mercado = (Previsto × (1 − %) + Precio de cierre × %) × apuntamiento + click × %`

  Ejemplo: previsto 70 €/MWh, cierras el 40 % a 62 €/MWh → 70 × 0,6 + 62 × 0,4 =
  **66,8 €/MWh** antes del apuntamiento; con un click de 0,30 €/MWh se suman 0,12 €/MWh.

---
**Cierres por MW (carga base y pico)** (típico de grandes consumos en Pass-Through)
- **MW**: potencia cerrada **las 24 horas de todos los días** del mes (producto *baseload*).
  Energía cerrada = MW × horas del mes (por ejemplo, 1 MW en un mes de 30 días = 720 MWh).
- **Precio de cierre**: precio del producto base. **Vacío = OMIP del mes.**
- **MW pico** y **Precio pico**: producto *peakload*, de **8 a 20 h de lunes a viernes**
  (12 h por día laborable). Precio pico vacío = Previsto × factor de carga pico (tabla de
  Parámetros).
- **Cómo se liquida**: el consumo se sigue pagando al precio de mercado y, aparte, se liquida
  la diferencia de lo cerrado:

  `liquidación (€) = MW × horas × (Precio de cierre − Previsto)`
  `                + MW pico × horas pico × (Precio pico − Previsto pico)`

  Se reparte entre la energía en barras de toda la cartera del mes (€/MWh) y se suma al
  precio de mercado de todos los CUPS de la oferta, junto con el click por cada MWh cerrado.
  Si cierras **por debajo** del previsto, la liquidación es negativa: **ahorras**; si cierras
  **por encima**, pagas la diferencia.

  Ejemplo: 0,5 MW en un mes de 720 h, cerrado a 60 €/MWh con un previsto de 70 €/MWh →
  0,5 × 720 × (60 − 70) = **−3.600 €** ese mes.

"""


def avisos_coherencia(o: mo.Oferta) -> list:
    """Combinaciones de la ficha que suelen ser un error al meter la oferta."""
    a = []
    if o.modalidad == mo.PASS_THROUGH and not (o.pagos_capacidad or o.os or o.om):
        if o.ssaa_modo != mo.SSAA_INCLUIDOS:
            a.append("Pass-Through sin pagos por capacidad, OS ni OM: revisa si la fórmula los "
                     "suma aparte.")
    if o.modalidad == mo.PASS_THROUGH and o.ssaa_modo == mo.SSAA_INCLUIDOS and o.pagos_capacidad:
        a.append("SSAA «Incluidos» pero pagos por capacidad aparte: comprueba que la oferta no "
                 "los lleve también en su término fijo.")
    if o.modalidad == mo.FIJO and (o.pagos_capacidad or o.os or o.om or o.srad):
        a.append("Precio fijo con componentes de mercado marcados: en fijo no se usan.")
    if o.modalidad == mo.FIJO and o.k:
        a.append("En precio fijo la K (impuesto municipal) suele ir incluida en el precio: "
                 "comprueba que el PDF la cobre aparte.")
    if o.ssaa_modo == mo.SSAA_INCLUIDOS and o.srad and o.modalidad == mo.PASS_POOL:
        a.append("Pass-Pool con SRAD aparte: comprueba que no esté ya en el término B.")
    if o.ssaa_modo == mo.SSAA_BANDA and o.banda_max <= o.banda_min:
        a.append("Banda de SSAA: la referencia máxima debe ser mayor que la mínima.")
    if o.ssaa_modo == mo.SSAA_TECHO and o.techo <= 0:
        a.append("Techo de SSAA a 0: se pagarían todos los SSAA. Escribe el umbral.")
    con_b = any(v for pr in o.precios.values() for p in pr for v in (p.termino_b or []))
    if o.modalidad == mo.PASS_THROUGH and o.fee == 0 and not con_b:
        a.append("Pass-Through sin fee de gestión: la mayoría de ofertas cobran uno.")
    if o.k and o.k > 0.05:
        a.append(f"K del {o.k * 100:.2f} %".replace(".", ",") + ": suele ser del 1,5 %. "
                 "Escríbela en %, por ejemplo 1,5 (no 150).")
    if o.fee > 50:
        a.append(f"Fee de {o.fee:g} €/MWh: muy alto. ¿Venía en c€/kWh (× 10) o es un término B?")
    if o.gdo > 10:
        a.append(f"GdO de {o.gdo:g} €/MWh: muy alto; suelen costar 0,5-3 €/MWh.")
    if o.modalidad == mo.FIJO and o.atr_incluido and o.unidad == "€/MWh":
        precios = [v for pr in o.precios.values() for p in pr for v in (p.energia or []) if v]
        if precios and max(precios) < 60:
            a.append("Precio fijo «con ATR» por debajo de 60 €/MWh: ¿seguro que incluye peajes "
                     "y cargos?")
    if o.ssaa_modo != mo.SSAA_INCLUIDOS and o.k_ssaa and o.k:
        dobles = [f for f in (o.ssaa_f1, o.ssaa_f2) if f and abs(f - (1 + o.k)) < 0.0005]
        if dobles:
            a.append("La K ya se aplica a los SSAA y además hay un factor de SSAA de "
                     f"{dobles[0]:.4f}".replace(".", ",") + ": el impuesto municipal se "
                     "cobraría dos veces. Deja el factor en 1 o desmarca «La K se aplica a "
                     "los SSAA».")
    if o.indice_ph and o.ssaa_modo != mo.SSAA_INCLUIDOS:
        a.append("Indexada al Ph de Canarias con SSAA aparte: el Ph ya incluye los costes de "
                 "ajuste (posible doble cobro).")
    return a


def resumen_formula(o: mo.Oferta) -> str:
    """La fórmula de la oferta en una línea, para compararla con la del PDF."""
    def n(v, d=3):
        return f"{v:.{d}f}".rstrip("0").rstrip(".").replace(".", ",")
    regul = [x for x, si in (("pagos por capacidad", o.pagos_capacidad), ("OS", o.os),
                             ("OM", o.om), ("SRAD", o.srad)) if si]
    if o.desvios:
        regul.append(f"desvíos {n(o.desvios)}")
    reg = (" + " + " + ".join(regul)) if regul else ""
    b = "término B" + (" × (1 + pérdidas)" if o.base_b == mo.BARRAS else "")
    if o.modalidad == mo.FIJO:
        ener = f"Energía = precio fijo ({o.unidad})" + (f" − {n(o.descuento * 100, 2)} %"
                                                       if o.descuento else "")
        ener += "" if o.atr_incluido else " + peajes y cargos"
    elif o.modalidad == mo.PASS_POOL:
        merc = "OMIE medio del mes" if o.ci_media_mensual else "OMIE × apuntamiento"
        ener = f"Energía = Ci × ({merc}{reg}) + {b}"
        ener += "" if o.atr_incluido else " + peajes y cargos"
    else:
        ener = f"Energía = (OMIE × apuntamiento{reg}) × (1 + pérdidas) + {b}"
        ener += "" if o.atr_incluido else " + peajes y cargos"
    partes = [ener]
    ss = {mo.SSAA_INCLUIDOS: "SSAA incluidos", mo.SSAA_PERIODO: "SSAA reales por periodo",
          mo.SSAA_PLANO: "SSAA reales planos",
          mo.SSAA_BANDA: f"SSAA en banda {n(o.banda_min)}-{n(o.banda_max)} €/MWh",
          mo.SSAA_TECHO: f"SSAA con techo {n(o.techo)} €/MWh"}[o.ssaa_modo]
    if o.ssaa_modo != mo.SSAA_INCLUIDOS and o.ssaa_perdidas == mo.PERD_TARIFA:
        ss += " con pérdidas"
    partes.append(ss)
    otros = []
    if o.fee:
        otros.append(f"fee {n(o.fee)} €/MWh" + (" en barras" if o.base_fee == mo.BARRAS else ""))
    if o.fnee != mo.FNEE_INCLUIDO:
        otros.append("FNEE" + (" en barras" if o.fnee == mo.BARRAS else ""))
    if o.gdo:
        otros.append(f"GdO {n(o.gdo)} €/MWh")
    if otros:
        partes.append(" + ".join(otros))
    if o.k:
        partes.append(f"K {n(o.k * 100, 4)} %" + (" (también sobre SSAA)" if o.k_ssaa else ""))
    partes.append("potencia " + ("del BOE" if o.potencia == mo.POT_REGULADA else "propia"))
    return " · ".join(partes)


# ------------------------------------------------------------------- fórmula completa (LaTeX)
def _n(v, d=3):
    t = f"{v:.{d}f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
    return t.replace(".", "{,}")


def _t(texto):
    return r"\text{" + texto + "}"


def formula_latex(o: mo.Oferta):
    """(fórmula en LaTeX, leyenda) del precio de la energía de la oferta en €/MWh de cada
    periodo, con sus paréntesis y corchetes, igual que la calcula el motor:
        { [ (lo que va con pérdidas) × (1 + Pérdidas) + lo que va sin pérdidas ] × (1 + K)
          + SSAA (si la K no los incluye) + peajes y cargos } × (1 + I.E.)"""
    ley = []
    P = r"(1+" + _t("Pérdidas") + ")"
    if o.indice_ph:
        omie = r"(" + _t("OMIE") + "+" + _t("Ph−diario") + r")\times " + _t("Apunt.Ph")
        ley.append("Ph − diario: lo que el Ph demanda de Canarias lleva encima del mercado diario.")
    elif o.cierres_tipo != mo.SIN_CIERRES:
        omie = _t("Mercado con cierres") + r"\times " + _t("Apunt.")
        ley.append("Mercado con cierres: previsto en la parte abierta y precio cerrado en la parte "
                   "cerrada; los MW cerrados se liquidan aparte.")
    else:
        omie = _t("OMIE") + r"\times " + _t("Apunt.")
    regul = [_t(x) for x, si in (("Capacidad", o.pagos_capacidad), ("OS", o.os), ("OM", o.om),
                                  ("SRAD", o.srad)) if si]
    if o.desvios:
        regul.append(_t("Desvíos ") + _n(o.desvios))
    con_b = any(v for pr in o.precios.values() for p in pr for v in (p.termino_b or []))
    con_perdidas, sin_perdidas = [], []          # términos que van con (1 + Pérdidas) y los que no
    # --- energía
    if o.modalidad == mo.FIJO:
        sin_perdidas.append(_t("Precio fijo") + (r"\times(1-" + _n(o.descuento * 100, 2) + r"\%)"
                                                 if o.descuento else ""))
        ley.append(f"Precio fijo: el de cada periodo de la pestaña Precios ({o.unidad}).")
    elif o.modalidad == mo.PASS_POOL:
        merc = _t("OMIE medio del mes") if o.ci_media_mensual and not o.indice_ph else omie
        sin_perdidas.append(_t("Ci") + r"\times" + (r"\left(" + "+".join([merc] + regul) + r"\right)"
                                                    if regul else r"\left(" + merc + r"\right)"))
        (con_perdidas if o.base_b == mo.BARRAS else sin_perdidas).append(_t("B"))
        ley.append("Ci y B: coeficiente y término fijo de cada periodo (pestaña Precios).")
    else:
        con_perdidas += [omie] + regul
        if con_b:
            (con_perdidas if o.base_b == mo.BARRAS else sin_perdidas).append(_t("B"))
            ley.append("B: término fijo de cada periodo (pestaña Precios).")
    # --- otros conceptos
    if o.fee:
        (con_perdidas if o.base_fee == mo.BARRAS else sin_perdidas).append(_t("Fee ") + _n(o.fee))
    if o.fnee == mo.BARRAS:
        con_perdidas.append(_t("FNEE"))
    elif o.fnee == mo.NETO:
        sin_perdidas.append(_t("FNEE"))
    if o.gdo:
        sin_perdidas.append(_t("GdO ") + _n(o.gdo))
    # --- SSAA
    ssaa, ssaa_perd = "", False
    if o.ssaa_modo != mo.SSAA_INCLUIDOS:
        sp = _t("SSAA") + (r"\times " + _t("Apunt.SSAA") if o.ssaa_perfil else "")
        ssaa = {mo.SSAA_PERIODO: _t("SSAA") + r"\times " + _t("Apunt.SSAA"),
                mo.SSAA_PLANO: _t("SSAA del mes"),
                mo.SSAA_BANDA: _t("Ajuste banda") + r"\left(" + sp + r",\ " + _n(o.banda_min)
                + "-" + _n(o.banda_max) + r"\right)",
                mo.SSAA_TECHO: r"\max\left(0,\ " + sp + "-" + _n(o.techo) + r"\right)"}[o.ssaa_modo]
        for f in (o.ssaa_f1, o.ssaa_f2):
            if f and abs(f - 1) > 1e-9:
                ssaa += r"\times " + _n(f, 4)
        if o.ssaa_perdidas == mo.PERD_TARIFA:
            ssaa_perd = True
        elif o.ssaa_perdidas == mo.PERD_FIJO:
            ssaa += r"\times(1+" + _n(o.ssaa_pct_perdidas * 100, 2) + r"\%)"
        ley.append({mo.SSAA_PERIODO: "SSAA × Apunt.SSAA: servicios de ajuste del mes a coste real, "
                                     "repartidos por periodo.",
                    mo.SSAA_PLANO: "SSAA del mes: servicios de ajuste a coste real, iguales en "
                                   "todos los periodos.",
                    mo.SSAA_BANDA: f"Ajuste banda: se paga lo que pase de {o.banda_max:g} €/MWh y "
                                   f"se descuenta lo que quede por debajo de {o.banda_min:g}.",
                    mo.SSAA_TECHO: f"Se paga solo lo que pase de {o.techo:g} €/MWh."}[o.ssaa_modo])
    else:
        ley.append("SSAA incluidos en el precio: no se suman.")
    ssaa_en_k = bool(ssaa and o.k and o.k_ssaa)
    if ssaa_en_k:
        (con_perdidas if ssaa_perd else sin_perdidas).append(ssaa)

    def grupo(terminos, por_linea=4):
        """Paréntesis × (1 + Pérdidas); si hay muchos términos se parte en varias líneas para
        que la fórmula no se salga del recuadro."""
        if len(terminos) <= por_linea:
            return r"\left(" + "+".join(terminos) + r"\right)\times " + P
        trozos = ["+".join(terminos[i:i + por_linea]) for i in range(0, len(terminos), por_linea)]
        return (r"\Big(" + r" \\ &\qquad+".join(trozos) + r"\Big)\times " + P)
    # --- líneas: lo que va con pérdidas en un paréntesis y, debajo, lo que va sin pérdidas
    lineas = []
    if con_perdidas:
        lineas.append(grupo(con_perdidas))
    if sin_perdidas:
        lineas.append(("+" if lineas else "") + "+".join(sin_perdidas))
    if not lineas:
        lineas.append("0")
    if o.k:
        lineas[0] = r"\Big[\," + lineas[0]
        lineas[-1] += r"\,\Big]\times(1+" + _n(o.k * 100, 6) + r"\%)"
        ley.append(f"{o.k * 100:g} %".replace(".", ",") + ": K / impuesto municipal"
                   + (" (también sobre los SSAA)." if o.k_ssaa else "."))
    cola = []
    if ssaa and not ssaa_en_k:
        cola.append(grupo([ssaa]) if ssaa_perd else ssaa)
    if not o.atr_incluido:
        cola.append(_t("Peajes y cargos"))
        ley.append("Peajes y cargos: los de energía del BOE de cada periodo y año.")
    if cola:
        lineas.append("+" + "+".join(cola))
    # impuesto eléctrico sobre todo
    lineas[0] = r"\bigg\{" + lineas[0]
    lineas[-1] += r"\bigg\}\times(1+" + _t("I.E.") + ")"
    expr = " ".join(lineas)
    if "Apunt." in expr:
        ley.append("Apunt.: apuntamiento (precio del periodo ÷ precio medio del mes).")
    if "Pérdidas" in expr:
        ley.append("Pérdidas: coeficiente de pérdidas de la tarifa en cada periodo; el paréntesis "
                   "reúne todo lo que va elevado a pérdidas (en barras de central).")
    ley.append("I.E.: impuesto eléctrico (5,11 % sobre toda la factura, también la potencia).")
    ley.append("Aparte: término de potencia (" + ("BOE" if o.potencia == mo.POT_REGULADA
                                                  else "precio propio") +
               "), que también lleva I.E., y alquiler del contador.")
    cuerpo = r" \\ &".join(lineas)
    return (r"\begin{aligned}" + _t("Precio energía (€/MWh)") + r"=\;&" + cuerpo
            + r"\end{aligned}"), ley
