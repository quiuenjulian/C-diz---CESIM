# -*- coding: utf-8 -*-
"""
metric_crosswalk.py -- Tabla de cruce entre las dos fuentes de datos del control de gestión:

  - PROYECTADO: CADIZ_Gestion_v2.xlsx!DATA_EXPORT, leído directamente por gap_analysis.load_proyeccion()
    (columnas: round/scenario/status/team/region/technology/metric/value/unit/...).
  - REAL: el DataFrame tidy que arma cesim_parser.build_historico() a partir de los .xls de RDOS
    (columnas: Ronda/Tipo_Ronda/Ronda_Numero/Ronda_Orden/Modulo/Estado/Seccion/Subgrupo/Metrica/
    Empresa/Valor).

Cada entrada fue VERIFICADA numéricamente contra Ronda 1 (dato real conocido en ambos lados) antes de
incorporarse acá -- no se listó ninguna combinación a ciegas. Ver Informe_V2.md / README de la
integración web para el detalle de la verificación.

Alcance de este primer corte (Adenda 8): los 8 KPIs de cabecera de la sección Finanzas del tablero
(seccion_finanzas() en app.py), Global (no desglosado por país/tecnología). Se amplía a otras
secciones en cortes siguientes, con la misma disciplina de verificación fila a fila.

Cada entrada:
    "clave_kpi": {
        "label": "Nombre para mostrar en el tablero",
        "proyeccion": {"metric": <valor de la columna 'metric' en cadiz_proyeccion.csv>,
                        "region": <valor de la columna 'region'>},
        "real": {"estado": <valor de 'Estado' en el df de cesim_parser>,
                  "metrica": <valor de 'Metrica'>,
                  "seccion": <valor de 'Seccion', o None si no aplica>},
        "tipo": "usd" | "ratio" | "texto",   # cómo formatear e interpretar el gap
        "gap_favorable": "real_mayor" | "real_menor" | None,  # para colorear el gap (None = texto)
    }
"""

CROSSWALK_FINANZAS = {
    "ebitda": {
        "label": "EBITDA",
        "proyeccion": {"metric": "EBITDA", "region": "Global"},
        "real": {"estado": "Cuenta de resultados, miles USD, Global",
                 "metrica": "Beneficio operativo antes de depreciación (EBITDA)", "seccion": None},
        "tipo": "usd",
        "gap_favorable": "real_mayor",
    },
    "margen_bruto": {
        "label": "Margen bruto",
        "proyeccion": {"metric": "Indicadores Financieros Claves — Margen bruto", "region": "Global"},
        "real": {"estado": "Ratios e indicadores financieros clave", "metrica": "Margen bruto", "seccion": None},
        "tipo": "ratio",
        "gap_favorable": "real_mayor",
    },
    "ros": {
        "label": "ROS",
        "proyeccion": {"metric": "ROS", "region": "Global"},
        "real": {"estado": "Ratios e indicadores financieros clave",
                 "metrica": "Rentabilidad de las ventas (ROS)", "seccion": None},
        "tipo": "ratio",
        "gap_favorable": "real_mayor",
    },
    "caja_final": {
        "label": "Caja final",
        "proyeccion": {"metric": "Efectivo y equivalentes", "region": "Global"},
        "real": {"estado": "Hoja de Balance, miles USD, Global",
                 "metrica": "Efectivo y equivalentes de efectivo", "seccion": None},
        "tipo": "usd",
        "gap_favorable": "real_mayor",
    },
    "deuda_cp_no_planificada": {
        "label": "Deuda CP no planificada",
        "proyeccion": {"metric": "Deudas a corto plazo (automática)", "region": "Global"},
        "real": {"estado": "Hoja de Balance, miles USD, Global",
                 "metrica": "Deudas a corto plazo (no planificadas)", "seccion": None},
        "tipo": "usd",
        "gap_favorable": "real_menor",
    },
    "deuda_lp": {
        "label": "Deuda LP",
        "proyeccion": {"metric": "Deudas a largo plazo", "region": "Global"},
        "real": {"estado": "Hoja de Balance, miles USD, Global",
                 "metrica": "Deudas a largo plazo", "seccion": None},
        "tipo": "usd",
        "gap_favorable": None,   # ni más ni menos deuda LP es "mejor" per se -- es una decisión, no un resultado a evaluar
    },
    "calificacion_crediticia": {
        "label": "Calificación crediticia",
        "proyeccion": {"metric": "Indicadores Financieros Claves — Calificación crediticia", "region": "Global"},
        "real": {"estado": "Ratios e indicadores financieros clave",
                 "metrica": "Calificación crediticia", "seccion": None},
        "tipo": "texto",
        "gap_favorable": None,
    },
    "ingresos_por_ventas": {
        "label": "Ingresos por ventas",
        "proyeccion": {"metric": "Ingresos por ventas", "region": "Global"},
        "real": {"estado": "Cuenta de resultados, miles USD, Global",
                 "metrica": "Ingresos por ventas", "seccion": None},
        "tipo": "usd",
        "gap_favorable": "real_mayor",
    },
}

# Alcance ampliado (Adenda 11): Mercado y Operaciones, Global/por-región donde corresponde. Mismo
# criterio -- cada entrada se verificó contra Ronda 1 real antes de incorporarse, y solo se listaron
# combinaciones donde la definición de ambos lados (Plan/Real) es efectivamente la misma magnitud.
#
# NOTA IMPORTANTE sobre alcance temporal: tanto Mercado como Operaciones dependen de CADIZ_Gestion_v2
# .xlsx!DATA_EXPORT, que solo proyecta (status=PLAN) desde Ronda 2 en adelante -- Ronda 0 y Ronda 1 ya
# se jugaron sin un plan cargado de antemano en este modelo, así que ahí no corresponde ningún gap
# (calcular_gaps ya devuelve 'sin_proyeccion' para esas rondas, correctamente). El primer gap real
# aparece en Ronda 2 en cuanto CESIM publique su RDOS.
#
# "Cuota de mercado CADIZ (ponderada por volumen)" (fila de 03_RATIOS, antes "promedio simple, 3
# mercados, Combustión+Híbrido"): la fórmula del Excel dividía por 12 (incluía los 6 casilleros de
# Eléctrico/Hidrógeno donde CADIZ no compite, sumando ceros de más en vez de ponderar por volumen
# como hace CESIM) -- confirmado contra Ronda 1 real (CESIM: 18,03% ponderado por ventas; Excel
# viejo: 4,63% en Ronda 2 con esos mismos supuestos). CORREGIDO en build_gestion_v2.py (ver Adenda
# de Integracion_Excel_Web_Control_Gestion.md): ahora pondera la cuota de cada país por su tamaño
# de mercado (Sección A1 del motor), igual convención que el dato REAL. Sigue sin agregarse a este
# crosswalk (Control de Gestión) -- eso es una decisión aparte, a confirmar con el equipo.
CROSSWALK_MERCADO = {
    "demanda_estimada_eeuu": {
        "label": "Demanda estimada — EE.UU.",
        "proyeccion": {"metric": "Demanda estimada CADIZ (unidades)", "region": "EE.UU."},
        "real": {"estado": "Informe de mercado, EE.UU.", "metrica": "Demanda, miles unidades", "seccion": None},
        "tipo": "unidades", "gap_favorable": None,
    },
    "demanda_estimada_china": {
        "label": "Demanda estimada — China",
        "proyeccion": {"metric": "Demanda estimada CADIZ (unidades)", "region": "China"},
        "real": {"estado": "Informe de mercado, China", "metrica": "Demanda, miles unidades", "seccion": None},
        "tipo": "unidades", "gap_favorable": None,
    },
    "demanda_estimada_europa": {
        "label": "Demanda estimada — Europa",
        "proyeccion": {"metric": "Demanda estimada CADIZ (unidades)", "region": "Europa"},
        "real": {"estado": "Informe de mercado, Europa", "metrica": "Demanda, miles unidades", "seccion": None},
        "tipo": "unidades", "gap_favorable": None,
    },
}

# "Utilización de capacidad": el RDOS NO publica un único % por área (EE.UU./China) -- lo publica
# desglosado por (área, tecnología) en 'Detalles de fabricación' -> Seccion='Capacidad empleada, %',
# Subgrupo=área, Metrica=tecnología (ej. R1 real CADIZ: EE.UU./Combustión=30%, EE.UU./Híbrido=55%).
# Un primer intento de crosswalk (bug ya corregido) apuntaba mal a Metrica='Capacidad empleada, %'
# -- eso no existe como tal, siempre salía vacío. En vez de eso se RECONSTRUYE el agregado real con
# la misma lógica que usa el propio modelo Excel para su lado Plan: (Producción interna real, sumada
# todas las tecnologías del área) / Capacidad operativa (cierre de ronda) -- ver 'real_calc' más abajo,
# resuelto en gap_analysis._valor_real_utilizacion_capacidad(). La capacidad NO se hardcodea como
# constante en la web: se lee de CADIZ_Gestion_v2.xlsx!DATA_EXPORT (metric='Capacidad operativa',
# status=PLAN), que el motor ya actualiza solo si CADIZ invierte/desinvierte (decisión D2) -- así que
# si la capacidad cambia en una ronda futura, el real se recalcula solo, sin tocar la web. Verificado
# EXACTO contra Ronda 1 real: (336+616)/1.400 = 68,0% EE.UU., (240+100)/500 = 68,0% China -- coincide
# con el 68% ya confirmado por el equipo (no el 85% que se había calculado mal en un intento anterior).
#  "Capacidad empleada" (antes "Utilización de capacidad" acá) -- mismo término que usa la tarjeta
#  nativa de Operaciones > Capacidad y Costos ("Capacidad empleada — EE.UU./China", desglosada por
#  tecnología); el "(total)" aclara que este KPI es el agregado de planta, no el desglose por
#  tecnología que se ve en esa otra pestaña -- mismo dato, mismo nombre base, alcance distinto.
CROSSWALK_OPERACIONES = {
    "utilizacion_capacidad_eeuu": {
        "label": "Capacidad empleada — EE.UU. (total)",
        "proyeccion": {"metric": "Utilización de capacidad", "region": "EE.UU."},
        "real": {"estado": "Detalles de fabricación", "metrica": "Capacidad empleada, %", "seccion": None},
        "real_calc": "utilizacion_capacidad", "real_calc_region": "EE.UU.",
        "tipo": "ratio", "gap_favorable": None,
    },
    "utilizacion_capacidad_china": {
        "label": "Capacidad empleada — China (total)",
        "proyeccion": {"metric": "Utilización de capacidad", "region": "China"},
        "real": {"estado": "Detalles de fabricación", "metrica": "Capacidad empleada, %", "seccion": None},
        "real_calc": "utilizacion_capacidad", "real_calc_region": "China",
        "tipo": "ratio", "gap_favorable": None,
    },
}

# Resultados (Adenda 11). EPS: verificado exacto contra Ronda 1 real. FCF: CESIM no publica esta línea
# en el RDOS -- se deja en el crosswalk con 'real' apuntando a un campo inexistente a propósito, así
# calcular_gaps siempre devuelve 'sin_real' y el panel cae al gráfico de evolución (solo Proyectado),
# nunca a un gap inventado. "Retorno total acumulado del accionista": el dato real SÍ existe en el
# RDOS, pero el lado Plan es un PROXY explícito (no una réplica del algoritmo real de CESIM, que no es
# público) -- incluido por pedido expreso del equipo, ya que es el objetivo estratégico central del
# juego (manual cap. 2.1: "el ganador se determina por el retorno total de los accionistas"). Se
# etiqueta como PROXY en el propio label para no confundirlo con una regla CESIM verificada.
CROSSWALK_RESULTADOS = {
    "eps": {
        # "Ganancias por Acción (EPS)" -- mismo nombre que el campo real de CESIM ("Ganancias por
        # acción (EPS), USD"), antes acá decía solo la sigla en inglés sin la traducción.
        "label": "Ganancias por Acción (EPS)",
        "proyeccion": {"metric": "EPS", "region": "Global"},
        "real": {"estado": "Ratios e indicadores financieros clave",
                 "metrica": "Ganancias por acción (EPS), USD", "seccion": None},
        "tipo": "usd_accion", "gap_favorable": "real_mayor",
    },
    "fcf": {
        "label": "FCF (flujo de caja libre)",
        "proyeccion": {"metric": "FCF", "region": "Global"},
        "real": {"estado": "NO PUBLICADO POR CESIM", "metrica": "NO PUBLICADO POR CESIM", "seccion": None},
        "tipo": "usd", "gap_favorable": "real_mayor",
        # CESIM no publica una línea de FCF en el RDOS -- nunca va a haber 'real' para comparar (no es
        # un 'real pendiente' temporal como Utilización de capacidad en una ronda sin RDOS publicado
        # todavía). Por pedido del equipo, se muestra como evolución de la PROYECCIÓN propia de CADIZ
        # a través de las rondas, sin intentar un gap contra algo que no existe.
        "real_no_publicado": True,
    },
    "retorno_accionista": {
        # Acortado y sin corchetes en mayúscula ("Retorno acumulado del accionista (Proxy)") -- mismo
        # término base que el KPI nativo de Resultados ("Retorno acum. del accionista"), antes decía
        # acá "Retorno TOTAL ACUMULADO del accionista [PROXY]", una variante de nombre distinta.
        "label": "Retorno acumulado del accionista (Proxy)",
        "proyeccion": {"metric": "Retorno total acumulado del accionista (PROXY)", "region": "Global"},
        "real": {"estado": "Ratios e indicadores financieros clave",
                 "metrica": "Retorno total acumulado del accionista (p.a.), %", "seccion": None},
        "tipo": "ratio", "gap_favorable": "real_mayor",
    },
}
