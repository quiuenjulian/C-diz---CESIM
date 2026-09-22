# -*- coding: utf-8 -*-
"""
gap_analysis.py -- Control de gestión: compara la PROYECCIÓN de CADIZ (leída DIRECTAMENTE de
CADIZ_Gestion_v2.xlsx!DATA_EXPORT -- ver Adenda 10 / hotfix README) contra el dato REAL (el df tidy
que ya arma cesim_parser.build_historico() a partir de los RDOS oficiales) para los KPIs listados en
metric_crosswalk.py.

DISEÑO (corregido): ya no hay un CSV intermedio que mantener sincronizado a mano. El usuario sube el
Excel de gestión directamente a la raíz del repo (siempre con el mismo nombre, `CADIZ_Gestion_v2.xlsx`)
cada vez que hay una versión nueva; `load_proyeccion()` lo lee de ahí en cada carga de la app, así que
la versión presente en el repo es automáticamente "la más actual" -- sin exportar nada a mano.

Solo tiene sentido para RONDAS OFICIALES con número de ronda (el modelo de Excel de CADIZ proyecta
rondas oficiales, no las de práctica) -- para una ronda de práctica, o una ronda oficial que el
modelo todavía no cubre, calcular_gaps() devuelve estado='sin_proyeccion' para cada KPI en vez de
inventar un cruce.

"Real" puede no estar disponible todavía aunque haya proyección (CESIM no publicó los RDOS de esa
ronda) -- ese caso es distinto y se reporta como estado='sin_real', no como error.
"""
import os
import re

import pandas as pd

from export_proyeccion import read_dataframe, DEFAULT_EXCEL
from metric_crosswalk import CROSSWALK_FINANZAS, CROSSWALK_MERCADO, CROSSWALK_OPERACIONES, CROSSWALK_RESULTADOS

_RONDA_OFICIAL_RE = re.compile(r"^Ronda\s+(\d+)$")

# Mismo archivo que lee export_proyeccion.py -- vive en la raíz del repo, al lado de app.py. El
# usuario lo reemplaza (mismo nombre) cada vez que hay una versión nueva del modelo de gestión.
DEFAULT_PROYECCION_EXCEL = DEFAULT_EXCEL


# ------------------------------------------------------------------------------------------------
# NOTA (retirado parche de escala): hasta esta versión existía acá un parche `_corregir_escala_
# plan_usd()` que dividía /1000 todo valor status=='PLAN' + unit=='USD'. Se apoyaba en un benchmark
# "REAL" equivocado -- 'Ingresos por ventas' real de CADIZ Ronda 1 = 46,479,518.40 USD -- que en
# realidad es el valor CRUDO del RDOS ("miles USD" sin multiplicar), no el valor real en USD
# absolutos (que es 46,479,518,404.93 USD, ~46.48 MIL MILLONES, verificado bottom-up: precio de
# venta x volumen vendido x tipo de cambio sumado en las 3 áreas y 2 tecnologías activas, contra el
# total de "Ingresos por ventas" de la Cuenta de Resultados, 0,003% de diferencia). Es decir: el
# lado REAL estaba subvaluado 1000x (bug de lectura, ya corregido en el Excel de gestión --
# rdos_parser.py / build_historico_equipos() / _canonicalize_real() aplican el x1000 de "miles USD"
# correctamente), y el lado PLAN estaba en realidad BIEN escalado desde siempre. El parche
# "corregía" comparando contra el lado equivocado. Con el Excel ya corregido, ambos lados (PLAN y
# REAL) llegan en USD absolutos consistentes -- este parche queda retirado, no actualizado, porque
# aplicarlo ahora volvería a introducir el mismo desfasaje de 1000x pero en sentido inverso.
# ------------------------------------------------------------------------------------------------


def load_proyeccion(path=DEFAULT_PROYECCION_EXCEL, team="CADIZ"):
    """Lee la proyección de CADIZ directamente de CADIZ_Gestion_v2.xlsx!DATA_EXPORT. Devuelve None
    si el archivo todavía no fue subido al repo, o si algo falla al leerlo (hoja faltante, encabezado
    desalineado, archivo corrupto) -- el tablero debe poder arrancar igual, mostrando la sección de
    Control de Gestión como no disponible en vez de romper toda la app."""
    if not os.path.exists(path):
        return None
    try:
        df = read_dataframe(path, team=team)
    except Exception:
        return None
    return df


def _valor_proyeccion(df_proy, ronda_num, spec, team="CADIZ"):
    if df_proy is None or ronda_num is None:
        return None, None
    sub = df_proy[(df_proy["round"] == ronda_num) & (df_proy["team"] == team) &
                  (df_proy["metric"] == spec["metric"]) & (df_proy["region"] == spec["region"])]
    # spec["tech"] es opcional -- solo lo necesitan las métricas nuevas (Adenda 12) que traen varias
    # filas por (región, tecnología) con el MISMO nombre de métrica (ej. "Precio de venta" tiene una
    # fila por cada una de las 4 tecnologías en cada mercado). Sin este filtro, .iloc[0] más abajo
    # devolvería SIEMPRE la primera fila que aparece en DATA_EXPORT para esa métrica/región -- un
    # error silencioso. Las métricas viejas (sin desglose por tecnología, technology="NA") no pasan
    # "tech" y siguen funcionando exactamente igual que antes.
    if "tech" in spec:
        sub = sub[sub["technology"] == spec["tech"]]
    if sub.empty:
        return None, None
    row = sub.iloc[0]
    val = row["value"]
    # La columna 'value' de DATA_EXPORT mezcla números y texto (p.ej. la calificación crediticia, o
    # el enfoque de marketing en otras filas) en una sola columna. Se intenta convertir a número acá;
    # si falla (texto real, como la calificación crediticia) se deja tal cual.
    try:
        val = float(val)
    except (TypeError, ValueError):
        pass
    return val, row["status"]   # status: REAL (histórico) o PLAN (proyección propia)


def _valor_real(df_real, ronda_nombre, spec, team="CADIZ"):
    if df_real is None or ronda_nombre is None:
        return None
    sub = df_real[(df_real["Ronda"] == ronda_nombre) & (df_real["Empresa"] == team) &
                  (df_real["Estado"] == spec["estado"]) & (df_real["Metrica"] == spec["metrica"])]
    if spec.get("seccion"):
        sub = sub[sub["Seccion"] == spec["seccion"]]
    if sub.empty:
        return None
    # Adenda 26 (mismo bug de "Comparativa Plan vs. Real" de Mercado): sin filtro de 'seccion', antes
    # se tomaba .iloc[0] -- correcto para las métricas de un solo renglón (Global, sin desglose), pero
    # 'Demanda, miles unidades' viene desglosada por Seccion=tecnología (Combustión/Híbrido/Eléctrico/
    # Hidrógeno, ver 'Informe de mercado, {país}') SIN una fila de total ya sumado -- .iloc[0] se
    # quedaba con una sola tecnología (ej. solo Combustión) en vez de la demanda total. Se cambia a
    # sumar todas las filas que matchean -- no cambia nada para las métricas de un solo renglón
    # (verificado: EBITDA/Margen bruto/Efectivo/EPS/Retorno del accionista, Global, siempre 1 fila),
    # y corrige las que sí vienen desglosadas sin filtro de Seccion.
    val = pd.to_numeric(sub["Valor"], errors="coerce")
    if val.notna().any():
        return float(val.sum())
    return sub.iloc[0]["Valor"]   # texto (p.ej. calificación crediticia) -- no summable


def _valor_real_utilizacion_capacidad(df_real, df_proy, ronda_nombre, ronda_num, region, team="CADIZ"):
    """Reconstruye el % de utilización de capacidad REAL de un área (EE.UU./China) que el RDOS no
    publica como un único dato -- lo publica desglosado por tecnología ('Detalles de fabricación' ->
    Seccion='Capacidad empleada, %', Subgrupo=área, Metrica=tecnología). Se reconstruye con la MISMA
    lógica que usa el motor Excel para su propio lado Plan: (Producción interna real, sumada todas
    las tecnologías del área) / Capacidad operativa (cierre de ronda, leída de DATA_EXPORT -- no una
    constante hardcodeada acá, para que se actualice sola si CADIZ invierte/desinvierte capacidad en
    una ronda futura). Verificado EXACTO contra Ronda 1 real: 68,0% en ambas áreas.

    Devuelve el resultado en escala 0-100 (mismo convención que el resto de los campos 'real' de
    cesim_parser, p.ej. ROE=19.07 no 0.1907) para que _to_absoluto() lo procese sin necesitar un caso
    especial -- 'ratio' ya divide /100 más abajo en el flujo normal."""
    if df_real is None or df_proy is None or ronda_num is None:
        return None
    sub = df_real[(df_real["Ronda"] == ronda_nombre) & (df_real["Empresa"] == team) &
                  (df_real["Estado"] == "Detalles de fabricación") &
                  (df_real["Seccion"] == "Producción interna, miles unidades") &
                  (df_real["Subgrupo"] == region)]
    if sub.empty:
        return None
    prod_total_miles = pd.to_numeric(sub["Valor"], errors="coerce").sum()
    if pd.isna(prod_total_miles):
        return None
    prod_total = prod_total_miles * 1000.0   # "miles unidades" -> unidades absolutas

    cap_sub = df_proy[(df_proy["round"] == ronda_num) & (df_proy["team"] == team) &
                       (df_proy["metric"] == "Capacidad operativa") & (df_proy["region"] == region)]
    if cap_sub.empty:
        return None
    try:
        cap_val = float(cap_sub.iloc[0]["value"])
    except (TypeError, ValueError):
        return None
    if not cap_val:
        return None
    return prod_total / cap_val * 100.0


TECNOLOGIAS = ["Combustión", "Híbrido", "Eléctrico", "Hidrógeno"]
MERCADOS = ["EE.UU.", "China", "Europa"]
AREAS = ["EE.UU.", "China"]
MONEDA_MERCADO = {"EE.UU.": "USD", "China": "RMB", "Europa": "EUR"}


def _valor_real_grano(df_real, ronda_nombre, estado, seccion, metrica, subgrupo=None, empresa="CADIZ"):
    """Como _valor_real, pero además puede filtrar por 'Subgrupo' -- necesario para los campos reales
    (Adenda 12) donde Estado+Seccion+Metrica NO alcanza para identificar una sola fila (ej. 'Informe
    de costos': Seccion='Costo de producción interna por unidad, USD', Metrica=tecnología, y el ÁREA
    vive en Subgrupo, no en el nombre de la métrica -- a diferencia de 'Ventas en {mercado}', que sí
    desambigua el mercado en el propio nombre)."""
    if df_real is None or ronda_nombre is None:
        return None
    sub = df_real[(df_real["Ronda"] == ronda_nombre) & (df_real["Empresa"] == empresa) &
                  (df_real["Estado"] == estado) & (df_real["Seccion"] == seccion) & (df_real["Metrica"] == metrica)]
    if subgrupo is not None:
        sub = sub[sub["Subgrupo"] == subgrupo]
    if sub.empty:
        return None
    try:
        return float(sub.iloc[0]["Valor"])
    except (TypeError, ValueError):
        return None


def precio_volumen_mercado(df_real, df_proy, ronda_nombre, ronda_num, mercado, team="CADIZ"):
    """Precio y volumen (Plan y Real) de `team` en un mercado, por tecnología -- insumo del Análisis
    de Desvíos de Ingresos (Precio/Volumen/Mix, Adenda 12). El precio queda en la moneda NATIVA del
    mercado (MONEDA_MERCADO) en ambos lados -- ni CESIM ni este cálculo convierten a USD, así que no
    hace falta asumir un tipo de cambio (dato que además no es públicamente verificable). Devuelve
    solo las tecnologías con volumen Plan o Real > 0 -- las demás son tecnologías donde `team`
    directamente no compite en este mercado."""
    out = {}
    for tech in TECNOLOGIAS:
        precio_plan, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": "Precio de venta", "region": mercado, "tech": tech}, team)
        vol_plan, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": "Ventas efectivas (mercado)", "region": mercado, "tech": tech}, team)
        precio_real = _valor_real_grano(df_real, ronda_nombre, "Precio de venta", mercado, f"{tech}, {MONEDA_MERCADO[mercado]}", empresa=team)
        # OJO signo: "Detalles de logística" -> "Ventas en {mercado}" viene NEGATIVO en el RDOS (es
        # una salida en un ledger de inventario: Total disponible − Ventas − Exportado = Inventario
        # final) -- verificado cruzando contra 'Informe de mercado, {mercado}' -> 'Ventas, miles
        # unidades' (positivo, mismo valor absoluto: 667,398 en ambos para CADIZ/Combustión/EE.UU.
        # Ronda 1). Se usa directamente este segundo campo -- ya viene con el signo correcto y es el
        # mismo que usa el resto de la app (chart Demanda vs. Ventas de la sección Mercado).
        vol_real = _valor_real_grano(df_real, ronda_nombre, f"Informe de mercado, {mercado}", tech, "Ventas, miles unidades", empresa=team)
        vol_real = vol_real * 1000.0 if vol_real is not None else None   # "miles unidades" -> absoluto
        if not (vol_plan or vol_real):
            continue
        out[tech] = {"precio_plan": precio_plan, "precio_real": precio_real,
                     "vol_plan": vol_plan or 0.0, "vol_real": vol_real or 0.0}
    return out


def variacion_precio_volumen_mix(datos):
    """Análisis de Desvíos clásico de Contabilidad Gerencial (Sales Price / Volume / Mix Variance):
    descompone la variación de Ingresos (Real vs. Plan) de un conjunto de tecnologías (ya filtrado a
    UN mercado, en su moneda nativa -- ver precio_volumen_mercado) en 3 componentes. Fórmulas
    estándar (Horngren et al.), para cada tecnología i:
        Precio_i  = (P_real_i − P_plan_i) × Q_real_i
        Volumen_i = (ΣQ_real − ΣQ_plan) × Mix_plan_i × P_plan_i
        Mix_i     = (Mix_real_i − Mix_plan_i) × ΣQ_real × P_plan_i
    La suma de los 3 componentes reconcilia EXACTO con Ingresos_Real − Ingresos_Plan (identidad
    algebraica, no una aproximación) -- se verifica antes de devolver; si no reconcilia (no debería
    pasar nunca) se marca 'reconciliacion_ok': False en vez de mostrar números que no cierran."""
    q_plan_tot = sum(d["vol_plan"] for d in datos.values())
    q_real_tot = sum(d["vol_real"] for d in datos.values())
    ing_plan = sum((d["precio_plan"] or 0) * d["vol_plan"] for d in datos.values())
    ing_real = sum((d["precio_real"] or 0) * d["vol_real"] for d in datos.values())
    var_precio = var_volumen = var_mix = 0.0
    for d in datos.values():
        pp, pr = d["precio_plan"] or 0, d["precio_real"] or 0
        qp, qr = d["vol_plan"], d["vol_real"]
        mix_plan = (qp / q_plan_tot) if q_plan_tot else 0.0
        mix_real = (qr / q_real_tot) if q_real_tot else 0.0
        var_precio += (pr - pp) * qr
        var_volumen += (q_real_tot - q_plan_tot) * mix_plan * pp
        var_mix += (mix_real - mix_plan) * q_real_tot * pp
    reconciliado = ing_plan + var_precio + var_volumen + var_mix
    return {
        "ingresos_plan": ing_plan, "ingresos_real": ing_real,
        "var_precio": var_precio, "var_volumen": var_volumen, "var_mix": var_mix,
        "reconciliacion_ok": abs(reconciliado - ing_real) < max(1.0, abs(ing_real) * 1e-6),
    }


def costo_unitario_area(df_real, df_proy, ronda_nombre, ronda_num, area, team="CADIZ"):
    """Costo unitario de producción (propia y tercerizada), Plan vs. Real, por tecnología, en un
    ÁREA de producción (EE.UU./China) -- insumo del desvío de costos de fabricación en Operaciones
    (Adenda 12). Mismo grano (área × tecnología) en Plan y Real, en USD/u. absoluto en ambos lados --
    no hace falta ponderar por mix de origen. OJO alcance: esto es solo el costo de FABRICACIÓN
    (propia + contratada); transporte/aranceles y promoción se reportan por MERCADO de destino en el
    RDOS (no por área de origen), y repartirlos entre áreas de origen requeriría reconstruir la
    asignación de exportaciones del motor (Sección E) -- fuera de este primer corte. Por eso el
    'desvío de Unit Economics' en la web se limita a fabricación, y se aclara explícitamente que no
    reconcilia 100% con la Contribución Marginal unitaria completa."""
    out = {}
    for tech in TECNOLOGIAS:
        cu_propia_plan, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": "Costo unitario de producción propia", "region": area, "tech": tech}, team)
        cu_terc_plan, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": "Costo unitario de producción tercerizada", "region": area, "tech": tech}, team)
        cu_propia_real = _valor_real_grano(df_real, ronda_nombre, "Informe de costos", "Costo de producción interna por unidad, USD", tech, subgrupo=area, empresa=team)
        cu_terc_real = _valor_real_grano(df_real, ronda_nombre, "Informe de costos", "Costos de fabricación contratada por unidad, USD", tech, subgrupo=area, empresa=team)
        prod_propia_real = _valor_real_grano(df_real, ronda_nombre, "Detalles de fabricación", "Producción interna, miles unidades", tech, subgrupo=area, empresa=team)
        prod_terc_real = _valor_real_grano(df_real, ronda_nombre, "Detalles de fabricación", "Producción contratada, miles unidades", tech, subgrupo=area, empresa=team)
        # Filtra por producción REAL (no por el costo Plan): el motor calcula un costo unitario Plan
        # con la misma curva de costos para las 4 tecnologías aunque CADIZ no produzca ahí todavía
        # (el costo "existe" en la fórmula, pero no es una tecnología en la que CADIZ compita) -- si
        # se filtrara por cu_propia_plan!=0 quedarían Eléctrico/Hidrógeno mostrados por error.
        if not ((prod_propia_real or 0) or (prod_terc_real or 0)):
            continue
        out[tech] = {"cu_propia_plan": cu_propia_plan, "cu_propia_real": cu_propia_real,
                     "cu_terc_plan": cu_terc_plan, "cu_terc_real": cu_terc_real,
                     "prod_propia_real": (prod_propia_real or 0) * 1000.0, "prod_terc_real": (prod_terc_real or 0) * 1000.0}
    return out


def cuota_mercado_objetivo_vs_real(df_real, df_proy, ronda_nombre, ronda_num, mercado, team="CADIZ"):
    """Cuota de mercado por tecnología: Objetivo (Plan, decisión D1·DEMANDA) vs. Real, en la MISMA
    convención declarada por el propio input del Excel: "% del mercado regional TOTAL (no se
    multiplica por mix tecnológico)". OJO -- esto es DISTINTO del campo que el RDOS publica directo
    en 'Informe de mercado, {mercado}' -> Seccion='{mercado} cuotas de mercado, %' -> Metrica=
    tecnología: se verificó que ESE campo usa otra convención (Ventas del equipo en esa tecnología /
    Σ Ventas de los 7 equipos EN ESA TECNOLOGÍA -- ej. Ronda 1 Combustión EE.UU.: 667,398/4.146,85 =
    16,09%, exacto). Compararlo directo contra el Objetivo sería mezclar dos definiciones distintas
    de "cuota" -- por eso el Real se reconstruye ACÁ con la misma convención que el Objetivo: Ventas
    reales de `team` en esa tecnología / tamaño total del mercado (Σ Ventas reales, los 7 equipos,
    las 4 tecnologías)."""
    sub_ventas = df_real[(df_real["Ronda"] == ronda_nombre) &
                          (df_real["Estado"] == f"Informe de mercado, {mercado}") &
                          (df_real["Metrica"] == "Ventas, miles unidades")].copy()
    if sub_ventas.empty:
        return {}
    sub_ventas["Valor"] = pd.to_numeric(sub_ventas["Valor"], errors="coerce")
    tamano_total = sub_ventas["Valor"].sum()
    out = {}
    for tech in TECNOLOGIAS:
        objetivo, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": "Cuota de mercado objetivo CADIZ", "region": mercado, "tech": tech}, team)
        ventas_team = sub_ventas[(sub_ventas["Seccion"] == tech) & (sub_ventas["Empresa"] == team)]["Valor"].sum()
        real_cuota = (ventas_team / tamano_total) if tamano_total else None
        if not ((objetivo or 0) or ventas_team):
            continue
        out[tech] = {"objetivo": objetivo, "real": real_cuota}
    return out


def flujo_caja_plan_real_global(df_real, df_proy, ronda_nombre, ronda_num, team="CADIZ"):
    """Composición del Flujo de Caja (CFO/CFI/CFF), Plan vs. Real, a nivel Global -- insumo del
    gráfico de Composición de Flujo de Caja en la Comparativa Plan vs. Real de Finanzas (Adenda 12).

    Plan: ya está en DATA_EXPORT como 3 líneas Global directas ("Total actividades operativas (CFO)",
    "...de inversión (CFI)", "...financieras (CFF)"), en USD absoluto (misma convención canónica que
    el resto del lado PLAN -- Cambio 3, ver Integracion_Excel_Web_Control_Gestion.md) -- no requieren
    ninguna transformación acá.

    Real: CESIM NO publica un único "Estado de flujo de efectivo, Global" en el RDOS -- lo publica en
    3 Estados separados: 'Flujo de efectivo de casa matriz, miles USD' (HQ + operación en EE.UU.) +
    'Estado de flujo de efectivo, miles USD, China' + '...Europa'. Se reconstruye sumando el
    CFO/CFI/CFF de los tres, CON el x1000 de "miles USD" (ver Adenda 27 en _to_absoluto() -- revierte
    la Adenda 25: cesim_parser entrega el valor CRUDO tal cual la celda, sin ningún x1000, así que la
    sección rotulada "miles USD" hay que escalarla acá, igual que en el resto de los campos 'usd').
    Los movimientos INTERCOMPAÑÍA (préstamos internos entre casa matriz y filiales, dividendos que las
    filiales giran a casa matriz) se CANCELAN naturalmente al sumar los tres lados -- no hace falta
    identificarlos ni restarlos a mano."""
    plan = {}
    for clave, metric in [("cfo", "Total actividades operativas (CFO)"),
                           ("cfi", "Total actividades de inversión (CFI)"),
                           ("cff", "Total actividades financieras (CFF)")]:
        val, _ = _valor_proyeccion(df_proy, ronda_num, {"metric": metric, "region": "Global"}, team)
        plan[clave] = val

    ESTADOS_CF = ["Flujo de efectivo de casa matriz, miles USD",
                  "Estado de flujo de efectivo, miles USD, China",
                  "Estado de flujo de efectivo, miles USD, Europa"]
    real = {"cfo": 0.0, "cfi": 0.0, "cff": 0.0}
    encontrado = False
    for estado in ESTADOS_CF:
        v_cfo = _valor_real_grano(df_real, ronda_nombre, estado, "Efectivo proveniente de actividades operativas", "Total", empresa=team)
        v_cfi = _valor_real_grano(df_real, ronda_nombre, estado, "Efectivo proveniente de actividades de inversión",
                                   "Inversiones en fábricas (-) / desinversiones (+)", empresa=team)
        v_cff = _valor_real_grano(df_real, ronda_nombre, estado, "Efectivo proveniente de actividades financieras", "Total", empresa=team)
        if v_cfo is not None or v_cfi is not None or v_cff is not None:
            encontrado = True
        real["cfo"] += (v_cfo * 1000.0) if v_cfo is not None else 0.0
        real["cfi"] += (v_cfi * 1000.0) if v_cfi is not None else 0.0
        real["cff"] += (v_cff * 1000.0) if v_cff is not None else 0.0
    if not encontrado:
        real = {"cfo": None, "cfi": None, "cff": None}
    return {"plan": plan, "real": real}


def _to_absoluto(valor, spec_tipo, fuente):
    """cesim_parser reporta % en escala 0-100 (p.ej. 19.99, no 0.1999) -- se normaliza a fracción acá
    para poder restar directo contra el lado 'real'. 'usd_accion' (p.ej. EPS) ya viene absoluto de
    origen en el RDOS -- sin conversión.

    Historial de la conversión x1000 en USD/unidades (para que no se repita el vaivén):
    - Adenda 23: se concluyó que faltaba un x1000 y se lo agregó. La evidencia era una reconciliación
      (precio x volumen vs. 'Ingresos de mercados' publicado) que resultó ser matemáticamente
      INVARIANTE a la escala -- multiplicar todo por cualquier factor k sigue "reconciliando", así que
      esa prueba nunca pudo haber distinguido nada.
    - Adenda 24: se revirtió lo anterior con una prueba cruzada (salario mensual x headcount de I+D)
      que parecía confirmar el x1000. Esa prueba asumía sin verificar que el headcount de I+D
      coincidía con la dotación total relevante para "Salarios y costos laborales" -- un supuesto
      propio no confirmado, no un dato.
    - Adenda 25: se revirtió el x1000 de USD porque "Julián confirmó directamente contra el RDOS" que
      Ingresos Global de CADIZ Ronda 1 ronda los USD 46 millones, no los USD 46 mil millones.
    - Adenda 27 (definitiva -- revierte la Adenda 25): esa "confirmación" leía el número crudo que
      exporta la plataforma ("miles USD" en el rótulo de sección) SIN aplicar el propio rótulo, es
      decir, tomaba el valor ya en miles como si fuera absoluto. Verificado bottom-up de forma
      independiente (precio de venta x volumen vendido x tipo de cambio, sumado en las 3 áreas y las
      2 tecnologías activas de CADIZ, Ronda 2): USD 73.716.619.198 contra el "Ingresos por ventas"
      Global publicado (73.714.297,44 "miles USD") x1000 = USD 73.714.297.440,62 -- 0,003% de
      diferencia, dentro del redondeo de "miles unidades" a 3 decimales. Es decir, "Ingresos por
      ventas" Global de CADIZ SÍ ronda los USD 46 MIL MILLONES en Ronda 1 (no los USD 46 millones) --
      el rótulo "miles USD" de CESIM es literal y aplica también acá, igual que ya se corrigió en el
      lado PLAN (ver rdos_parser.py / build_historico_equipos() del Excel de gestión). cesim_parser NO
      aplica ningún x1000 (ver cesim_parser.py -- 'Valor' es siempre el crudo de la celda), así que el
      x1000 que "se sacó" en la Adenda 25 nunca debió sacarse: se vuelve a aplicar acá, esta vez con
      evidencia verificada de forma independiente (no una lectura de pantalla). La contradicción
      pendiente de la Adenda 24 (prueba de nómina) queda resuelta en el mismo sentido: CON el x1000,
      "Salarios y costos laborales" es plausible frente a un headcount total (no solo I+D), que es la
      lectura correcta del dato -- confirma que el x1000 sí correspondía.
      NOTA (ver Categoría de dato): esta es la MISMA corrección de escala que _corregir_escala_plan_usd
      aplicaba (mal) al lado PLAN antes de retirarse de load_proyeccion() -- aquella dividía por 1000
      un valor PLAN que ya estaba bien, comparándolo contra este mismo REAL mal escalado. Con esta
      corrección, PLAN (ya correcto) y REAL (corregido acá) quedan en la MISMA escala absoluta.
    """
    if fuente != "real" or not isinstance(valor, (int, float)):
        return valor
    if spec_tipo == "ratio":
        return valor / 100.0
    if spec_tipo == "unidades":
        # Adenda 26 (bug reportado: "Comparativa Plan vs. Real" de Mercado se rompió -- Real de
        # Demanda estimada salía ~1000x menor que el Proyectado). La métrica real detrás de este tipo
        # ('Demanda, miles unidades', Estado 'Informe de mercado, {país}') SÍ está en miles de unidades
        # de verdad -- confirmado: CADIZ Ronda 1, EE.UU., Combustión = 667.398 sin este x1000, un
        # mercado automotor de ~667 unidades totales no es plausible; con el x1000 (667,398) queda en
        # el mismo orden de magnitud que el Proyectado de Ronda 2 (~1,041,456 unidades). Mismo x1000
        # que ya se aplica a mano en otros lugares de este archivo para "miles unidades" (ver
        # prod_total_miles/vol_real más arriba) -- acá se generaliza al tipo 'unidades' de
        # calcular_gaps() en vez de dejarlo sin convertir.
        return valor * 1000.0
    if spec_tipo == "usd":
        # Adenda 27: ver docstring de esta función -- mismo rótulo "miles USD" que 'unidades' tiene
        # "miles unidades", misma corrección (x1000). Cubre Ingresos, Costos, EBITDA, EBIT, Beneficio,
        # Activos/Pasivos/Patrimonio y componentes de Balance -- los "totales" agregados que trae
        # metric_crosswalk.py con tipo="usd" (ver ahí). NO cubre precios unitarios ("Precio de venta"),
        # costos unitarios ("Costo de producción... por unidad") ni EPS (tipo="usd_accion") -- esos NO
        # están rotulados "miles" en el RDOS y no se tocan (ver precio_volumen_mercado, sin cambios).
        return valor * 1000.0
    return valor


def ronda_a_num(ronda_nombre):
    """CADIZ_Gestion_v2.xlsx proyecta únicamente rondas OFICIALES ('Ronda N', round=N en
    DATA_EXPORT) -- una ronda de Práctica no tiene equivalente ahí, así que para esas
    devuelve None a propósito (fuerza 'sin_proyeccion'/'sin_datos' en calcular_gaps, y el panel cae
    al gráfico de evolución en vez de mostrar un gap inexistente). Se parsea del NOMBRE de la ronda
    (convención ya usada en app.py: rondas_timeline = ['Ronda 1', ..., 'Ronda 12']), no de si ya hay
    un RDOS real cargado para ella -- el selector de ronda deja elegir 'Ronda 2' aunque CESIM todavía
    no haya publicado ese resultado, que es exactamente el caso que el control de gestión necesita
    poder mostrar ('proyectado ya cargado, real pendiente')."""
    m = _RONDA_OFICIAL_RE.match(ronda_nombre or "")
    return int(m.group(1)) if m else None


def calcular_gaps(df_real, df_proy, ronda_nombre, ronda_num, team="CADIZ", crosswalk=None):
    """Devuelve {clave_kpi: {label, tipo, gap_favorable, proyectado, proyectado_status, real,
    gap_abs, gap_pct, estado}} para cada entrada del crosswalk.

    estado: 'ok' (hay proyección y real), 'sin_proyeccion' (no hay fila en DATA_EXPORT para
    esta ronda/KPI -- p.ej. ronda de práctica, o ronda oficial que el modelo Excel todavía no cubre),
    'sin_real' (hay proyección pero CESIM no publicó todavía el RDOS real de esta ronda), 'sin_datos'
    (no hay ninguna de las dos)."""
    crosswalk = crosswalk or CROSSWALK_FINANZAS
    out = {}
    for clave, spec in crosswalk.items():
        proy_val, proy_status = _valor_proyeccion(df_proy, ronda_num, spec["proyeccion"], team)
        if spec.get("real_calc") == "utilizacion_capacidad":
            real_val = _valor_real_utilizacion_capacidad(
                df_real, df_proy, ronda_nombre, ronda_num, spec["real_calc_region"], team)
        else:
            real_val = _valor_real(df_real, ronda_nombre, spec["real"], team)
        real_val = _to_absoluto(real_val, spec["tipo"], "real")

        if proy_val is None and real_val is None:
            estado = "sin_datos"
        elif proy_val is None:
            estado = "sin_proyeccion"
        elif real_val is None:
            estado = "sin_real"
        else:
            estado = "ok"

        gap_abs = gap_pct = None
        if estado == "ok" and spec["tipo"] in ("usd", "ratio", "unidades", "usd_accion") and isinstance(real_val, (int, float)) and isinstance(proy_val, (int, float)):
            gap_abs = real_val - proy_val
            if proy_val not in (0, None):
                gap_pct = gap_abs / abs(proy_val) * 100.0

        out[clave] = {
            "label": spec["label"], "tipo": spec["tipo"], "gap_favorable": spec["gap_favorable"],
            "proyectado": proy_val, "proyectado_status": proy_status, "real": real_val,
            "gap_abs": gap_abs, "gap_pct": gap_pct, "estado": estado,
        }
    return out
