import glob
import os
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from cesim_parser import build_historico
from gap_analysis import (calcular_gaps, load_proyeccion, ronda_a_num, DEFAULT_PROYECCION_EXCEL,
                          precio_volumen_mercado, variacion_precio_volumen_mix, costo_unitario_area,
                          cuota_mercado_objetivo_vs_real, flujo_caja_plan_real_global,
                          TECNOLOGIAS as _TECNOLOGIAS_GAP, MERCADOS as _MERCADOS_GAP, AREAS as _AREAS_GAP,
                          MONEDA_MERCADO as _MONEDA_MERCADO_GAP)
from metric_crosswalk import CROSSWALK_FINANZAS, CROSSWALK_MERCADO, CROSSWALK_OPERACIONES, CROSSWALK_RESULTADOS
# Motor de proyección (build_gestion_v2.py, ver README del motor): genera el Excel de gestión
# ("CADIZ_Gestion_v2.xlsx") al vuelo a partir de los RDOS de cada ronda jugada, sin tocar el disco
# del servidor (io.BytesIO) -- ver bloque "Generar Excel de gestión" en la sidebar, más abajo.
# Ronda N.N (Fase 4 -- arquitectura basada en repo, sin file_uploader): la fuente de RDOS oficiales y
# de decisiones previas ya no se sube a mano en cada corrida -- se lee de rutas fijas DENTRO del
# propio repo de GitHub (data/raw/oficial/ y data/decisiones/), commiteadas como parte del repo (igual
# que ya se commitea CADIZ_Gestion_v2.xlsx hoy). generar_excel_desde_repo() hace todo: detecta la
# frontera REAL/PLAN, arma el Excel completo, rescata decisiones futuras cargadas en el archivo de
# decisiones anterior y congela el plan de la ronda que acaba de pasar a REAL.
from build_gestion_v2 import generar_excel_desde_repo, descubrir_ultima_proyeccion
# --- IDENTIDAD Y PALETA SEMÁNTICA ---
MY_COMPANY = 'CADIZ'
COMPANIES = ['CADIZ', 'CEOS', 'CHIEF', 'CLAVE', 'CUORE', 'FOCUS', 'TOKIO']
# Adenda 25 (a pedido del equipo, mockup aprobado -- Opción B "Ciruela"): antes un solo rojo
# (BRAND_ACCENT) hacía DOBLE trabajo -- identificar a CÁDIZ Y marcar negativo/crítico -- por eso
# casi todo el tablero "se veía mal" con solo ver el color. Se separan en dos constantes con un
# solo uso cada una:
#   COLOR_CADIZ    -- SOLO identidad ("esto es CADIZ", o el equipo en foco cuando coincide con
#                     CADIZ). Nunca se usa para positivo/negativo.
#   COLOR_NEGATIVE -- SOLO semántica negativo/crítico, universal para cualquier equipo (antes
#                     BRAND_ACCENT). COLOR_POSITIVE (verde) es su contraparte, sin cambios.
# BRAND_ACCENT queda retirado del código -- si aparece en algún lugar es una referencia vieja sin
# migrar, no una tercera opción válida.
COLOR_CADIZ = '#7C5CA8'        # Lila Ciruela -- identidad CADIZ
COLOR_NEGATIVE = '#B3261E'     # Rojo -- semántica negativo/crítico (mismo rojo que antes, ahora sin
                                # doble uso)
COLOR_POSITIVE = '#94D02D'     # Verde Lima
BRAND_DARK = '#1A1714'         # Negro Grafito
BRAND_LIGHT = '#F5F2ED'        # Crema
# Tonos apagados pero CON matiz (no gris puro) para distinguir competidores de un vistazo,
# sin competir visualmente con el lila de identidad de CADIZ.
MUTED_PALETTE = ['#8C97A6', '#A68C6E', '#7E9E8C', '#9E8CA0', '#A69B6E', '#7E8C9E']
# Adenda 25 (a pedido del equipo, "sacá ese marrón, ese tipo de cosas"): variante de MUTED_PALETTE
# SIN el marrón (índice 1) -- para usar en `color_discrete_sequence=` de gráficos con categorías
# genéricas (sin identidad ni sentimiento) donde Plotly asigna el color por ORDEN de aparición, no
# por nombre -- ahí un MUTED_PALETTE[1] "se cuela" igual aunque no se lo nombre a mano (fue el caso
# de "Mix tecnológico": Plotly le tocó el marrón a "Híbrido" solo por ser la 2ª tecnología de la
# lista). No se reemplaza por lila: siguen siendo categorías sin identidad de CADIZ.
# Fix adicional (feedback "ya casi"): sacar el marrón no alcanzaba -- MUTED_PALETTE[5] y [0] son dos
# grises azulados casi idénticos (distancia RGB ~19.5, muy por debajo del resto de la paleta) y
# MUTED_PALETTE[4] (kaki) está a solo ~15 de distancia del marrón baneado, o sea que seguía leyéndose
# como una variante de marrón. Se reemplazaron por '#58759D' (azul acero) y '#B67C8F' (rosa apagado).
# Adenda 28 (feedback "poca diferencia entre colores, seguí una paleta única"): la versión anterior
# de MUTED_SIN_MARRON (gris azulado + sage + rosa apagado + lila grisáceo + azul acero) fue auditada
# con el validador de contraste de la skill dataviz (validate_palette.js, método OKLab/CVD) y
# reprobó: el par gris/sage medía apenas ΔE 5.3 en visión normal (piso mínimo aceptable: 15) y el
# par ámbar/sage del gráfico de WACC medía ΔE 13.4, también por debajo del piso. Paleta nueva,
# verificada con el mismo validador (5 colores, modo claro, lista de pares adyacentes -- el uso real
# es en barras/torta/chips, no en dispersión, así que ese es el criterio correcto de la skill):
# TODOS los checks pasan (Lightness band, Chroma floor, CVD separation, Normal-vision floor). Los
# primeros dos colores son deliberadamente los mismos hex que COLOR_METRICA['riesgo'] (ámbar) y
# COLOR_METRICA['dinero'] (azul, definido más abajo -- no se puede referenciar el dict acá porque
# todavía no existe en este punto del archivo) -- ya usados en el resto de la app -- para que el
# gráfico de Mix tecnológico y el de Deuda/Equity del WACC compartan una paleta única y reconocible,
# en vez de dos sets de colores distintos sin relación entre sí.
MUTED_SIN_MARRON = ['#C9922E', '#3E7CB1', '#1BAF7A', '#E87BA4', '#4A3AA7']
# Color por CONCEPTO, no por orden de aparición. El lila de marca queda reservado para
# identificar a CÁDIZ entre los equipos; las métricas usan colores con significado propio
# y estable en toda la app (antes el rojo era "Salario" en un gráfico y "Rotación" en el de al lado).
COLOR_METRICA = {
    'dinero':      '#3E7CB1',   # azul  — plata: costos, presupuestos, salarios
    'personas':    '#8C97A6',   # gris azulado — headcount, contrataciones
    'eficiencia':  '#4E9A4E',   # verde — indicadores donde más es mejor
    'riesgo':      '#C9922E',   # ámbar — rotación, deuda, alertas blandas
    'critico':     COLOR_NEGATIVE # rojo  — problemas (ya no comparte color con "es CADIZ")
}
COLOR_MAP = {MY_COMPANY: COLOR_CADIZ}
for i, c in enumerate([c for c in COMPANIES if c != MY_COMPANY]):
    COLOR_MAP[c] = MUTED_PALETTE[i % len(MUTED_PALETTE)]

def _hex_a_rgba(hex_color, alpha):
    """Convierte '#RRGGBB' a 'rgba(r,g,b,alpha)' -- para fills semitransparentes que tienen que
    matchear el color de línea que sea (antes varios fills estaban hardcodeados al rojo/lila de
    marca sin importar qué color se pasara)."""
    h = (hex_color or '').lstrip('#')
    if len(h) != 6:
        return f'rgba(140,151,166,{alpha})'
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f'rgba({r},{g},{b},{alpha})'
CHART_HEIGHT = 260      # alto del área de ploteo de referencia
ALTURA_TARJETA = 370    # alto FIJO de toda tarjeta de gráfico, con o sin leyenda abajo:
                        # es lo que garantiza que dos gráficos en columnas queden parejos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data', 'raw')
DIR_DECISIONES = os.path.join(BASE_DIR, 'data', 'decisiones')
st.set_page_config(page_title='CÁDIZ | Tablero Directivo', layout='wide')
# ---------------- Carga de datos y Helpers ----------------
def get_pais(estado: str) -> str:
    if re.search(r'\bGlobal\b', estado) or 'casa matriz' in estado: return 'Global'
    if 'EE.UU.' in estado: return 'EE.UU.'
    if 'China' in estado: return 'China'
    if 'Europa' in estado: return 'Europa'
    return 'General'
@st.cache_data(show_spinner='Procesando rondas...')
def cargar_historico(files_signature: tuple) -> pd.DataFrame:
    data = build_historico(list(files_signature))
    data['Pais'] = data['Estado'].apply(get_pais)
    return data
def get_data(tipo_ronda) -> pd.DataFrame:
    xls_files = sorted(glob.glob(os.path.join(DATA_DIR, '**', '*.xls'), recursive=True) +
                       glob.glob(os.path.join(DATA_DIR, '**', '*.XLS'), recursive=True))
    if not xls_files:
        st.warning(f'No se encontraron archivos .xls en {DATA_DIR} ni en subcarpetas.')
        st.stop()
    df_raw = cargar_historico(tuple(xls_files))
    return df_raw[df_raw['Tipo_Ronda'] == tipo_ronda].copy()
@st.cache_data(show_spinner='Leyendo proyección de CADIZ...')
def _cargar_proyeccion_cached(path: str, mtime: float):
    # `mtime` en la firma de cache -- no se usa dentro de la función, pero cambia cada vez que se
    # reemplaza CADIZ_Gestion_v2.xlsx en el repo (mismo nombre de archivo), así Streamlit invalida
    # el caché y relee el Excel en vez de servir una versión vieja. Devuelve un DataFrame o None.
    return load_proyeccion(path)
def get_proyeccion():
    """Ronda N.N (Fase 4): la proyección de CADIZ ya NO se lee de un archivo fijo en la raíz del
    repo ('CADIZ_Gestion_v2.xlsx', que había que reemplazar a mano con ese mismo nombre) -- se
    autodetecta el 'Cadiz_proyeccion_R{N}.xlsx' de N más alto en data/decisiones/ (el mismo que arma
    y descarga el botón 'Generar Excel de gestión' de la sidebar): ese archivo ya trae todo el
    histórico REAL migrado + la proyección PLAN de la ronda en curso + los planes congelados de
    rondas previas, así que ninguna otra fuente hace falta. Si todavía no se commiteó ningún
    Cadiz_proyeccion_R{N}.xlsx (repo recién migrado a esta arquitectura), se cae al viejo
    CADIZ_Gestion_v2.xlsx en la raíz por compatibilidad -- si tampoco existe, no hay proyección
    para mostrar (la app debe poder arrancar igual)."""
    path = descubrir_ultima_proyeccion(DIR_DECISIONES)
    if path is None and os.path.exists(DEFAULT_PROYECCION_EXCEL):
        path = DEFAULT_PROYECCION_EXCEL
    if path is None:
        return None
    return _cargar_proyeccion_cached(path, os.path.getmtime(path))
def num(series):
    return pd.to_numeric(series, errors='coerce')
def format_num(val, dec=1):
    """Al menos un decimal en todos los rangos (antes el corte de miles truncaba a entero, ej.
    '638k' en vez de '638.2k' -- perdía precisión visible justo en el rango donde más se usa)."""
    if pd.isna(val) or val is None: return ""
    try:
        val = float(val)
        dec = max(dec, 1)
        if abs(val) >= 1_000_000: return f"{val/1_000_000:,.1f}M"
        if abs(val) >= 1_000: return f"{val/1_000:,.1f}k"
        return f"{val:,.{dec}f}"
    except (ValueError, TypeError):
        return ""
def _texto_cascada(valores, base=None):
    """Texto para las barras de un gráfico en cascada (Waterfall): valor absoluto + % sobre la barra
    de referencia -- por defecto la PRIMERA barra de la cascada (el ancla desde la que se arma el
    resto, ej. 'Ingresos Proyectados', 'Costo Proyectado', 'Precio'). A pedido del equipo: además
    del valor absoluto, mostrar también el %, en todos los gráficos en cascada de la app. Si la base
    es 0/None (o el valor no es numérico), se muestra solo el absoluto -- dividir por 0 no aporta."""
    base = valores[0] if base is None else base
    out = []
    for v in valores:
        if base not in (None, 0) and pd.notna(v) and pd.notna(base):
            out.append(f'{format_num(v)} ({v / base * 100:.1f}%)')
        else:
            out.append(format_num(v))
    return out
def valor_de(sub_df, metrica, empresa=None):
    d = sub_df[sub_df['Metrica'] == metrica]
    if empresa: d = d[d['Empresa'] == empresa]
    return pd.to_numeric(d['Valor'].iloc[0], errors='coerce') if not d.empty else None
def valor_texto(sub_df, metrica, empresa=None):
    """Como valor_de, pero para campos de texto (ej. calificación crediticia 'A+') —
    no fuerza conversión a número, así que no se rompe en NaN."""
    d = sub_df[sub_df['Metrica'] == metrica]
    if empresa: d = d[d['Empresa'] == empresa]
    return d['Valor'].iloc[0] if not d.empty else None
def valor_fuzzy(sub_df, keyword, empresa=None):
    d = sub_df[sub_df['Metrica'].str.contains(rf'{keyword}', case=False, na=False)]
    if empresa: d = d[d['Empresa'] == empresa]
    return pd.to_numeric(d['Valor'].iloc[0], errors='coerce') if not d.empty else None
def es_modo_oscuro():
    """Fuente de verdad única: el tema REAL de Streamlit (nativo, el que el usuario elige en
    Settings o 'usar el del sistema'), no un toggle casero desconectado del resto de la UI."""
    try:
        return st.context.theme.type == 'dark'
    except Exception:
        return True
def mostrar(fig, ocultar_eje_valores=None, en_card=True, altura=None, margen_b=None, **kwargs):
    oscuro = es_modo_oscuro()
    # Si la figura ya trae leyenda propia posicionada abajo (y<0), necesita más alto/margen
    # para que la leyenda no quede tapando el gráfico.
    leyenda_abajo = False
    for ln in ('legend', 'legend2'):
        leg = getattr(fig.layout, ln, None)
        if leg is not None and leg.y is not None and leg.y < 0:
            leyenda_abajo = True
    # Altura ÚNICA para todos los gráficos del tablero. Antes cada figura podía traer la suya
    # (CHART_HEIGHT, +20, +60, 280) y las que tenían leyenda abajo crecían: dos gráficos en
    # columnas contiguas terminaban de distinto alto. Ahora la tarjeta siempre mide lo mismo y
    # lo único que cambia es cuánto de ese alto se reserva abajo para la leyenda.
    # Adenda 23: 'altura' opcional para paneles que agrupan varios KPIs en una sola figura (ej.
    # chart_bullet_panel) -- necesitan más alto que ALTURA_TARJETA cuando tienen muchas filas, y
    # menos cuando tienen pocas; el resto de los llamados no pasa este parámetro y sigue con el
    # alto fijo de siempre (ALTURA_TARJETA), sin cambio de comportamiento.
    altura = altura or ALTURA_TARJETA
    # Adenda 24: 'margen_b' opcional -- BUG REPORTADO (leyenda y anotaciones tapando las barras en
    # chart_bullet_panel): el margen inferior fijo (110px) se restaba de un 'altura' que en paneles
    # de pocas filas quedaba demasiado chica, dejando un área de ploteo casi nula -- ahí cualquier
    # offset de leyenda (que es una FRACCIÓN del área de ploteo, no de la figura entera) terminaba en
    # unos pocos píxeles, es decir prácticamente pegado al eje X, encima de las barras y anotaciones.
    # Quien llama con pocas filas ahora puede pedir un margen distinto sin afectar los ~40 llamados
    # existentes que no lo pasan (siguen con 110/20 de siempre).
    if margen_b is None:
        margen_b = 110 if leyenda_abajo else 20
    color_linea_eje = '#4A4642' if oscuro else '#D8D3CC'
    fig.update_layout(template='plotly_dark' if oscuro else 'plotly_white',
                       paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                       font=dict(color=BRAND_LIGHT if oscuro else BRAND_DARK, family='Plus Jakarta Sans, sans-serif'),
                       title_font=dict(family='Plus Jakarta Sans, sans-serif', size=14, weight=600),
                       margin=dict(l=20, r=20, t=45, b=margen_b),
                       height=altura,
                       bargap=0.4)
    fig.update_xaxes(showgrid=False, zeroline=False, showline=True, linewidth=1, linecolor=color_linea_eje)
    fig.update_yaxes(showgrid=False, zeroline=False, showline=True, linewidth=1, linecolor=color_linea_eje)

    if ocultar_eje_valores == 'y': fig.update_yaxes(showticklabels=False, title=None)
    elif ocultar_eje_valores == 'x': fig.update_xaxes(showticklabels=False, title=None)
    
    if en_card:
        with st.container(border=True): st.plotly_chart(fig, use_container_width=True, **kwargs)
    else: st.plotly_chart(fig, use_container_width=True, **kwargs)
def linea_media(fig, valor, eje='y', etiqueta='Promedio'):
    if pd.isna(valor): return
    color_ref = 'rgba(255,255,255,0.35)' if es_modo_oscuro() else 'rgba(26,23,20,0.35)'
    kwargs = dict(line_dash='dash', line_color=color_ref, line_width=1.5,
                  annotation_text=etiqueta, annotation_font_size=10, annotation_font_color=color_ref)
    if eje == 'y': fig.add_hline(y=valor, **kwargs)
    else: fig.add_vline(x=valor, **kwargs)
def chart_comparacion_equipos(sub: pd.DataFrame, titulo: str, ronda=None):
    ronda = ronda or ronda_ultima
    d = sub[sub['Ronda'] == ronda].copy()
    d['Valor'] = num(d['Valor'])
    d = d.dropna(subset=['Valor']).sort_values('Valor', ascending=False)
    if d.empty: return st.info('Sin datos numéricos.')
    d['Etiqueta'] = d['Valor'].apply(format_num)
    fig = px.bar(d, x='Empresa', y='Valor', color='Empresa', color_discrete_map=COLOR_MAP, title=f'{titulo} — {ronda}', text='Etiqueta')
    fig.update_traces(textposition='outside', cliponaxis=False, showlegend=False)
    mostrar(fig, ocultar_eje_valores='y')
def chart_evolucion(sub: pd.DataFrame, titulo: str):
    ev = sub.copy()
    ev['Valor'] = num(ev['Valor'])
    ev = ev.dropna(subset=['Valor'])
    if ev.empty: return st.info('Sin datos para evolución.')
    fig = go.Figure()
    for comp in COMPANIES:
        d = ev[ev['Empresa'] == comp].sort_values('Ronda_Orden')
        if d.empty: continue
        es_cadiz = comp == MY_COMPANY
        fig.add_trace(go.Scatter(x=d['Ronda'], y=d['Valor'], mode='lines+markers', name=comp,
                                  line=dict(color=COLOR_MAP[comp], width=3 if es_cadiz else 1),
                                  marker=dict(size=6 if es_cadiz else 4), opacity=1.0 if es_cadiz else 0.5))
    promedio_x_ronda = ev.groupby('Ronda_Orden').agg(Ronda=('Ronda', 'first'), Valor=('Valor', 'mean')).sort_index()
    if len(promedio_x_ronda) > 0:
        color_ref = 'rgba(255,255,255,0.35)' if es_modo_oscuro() else 'rgba(26,23,20,0.35)'
        fig.add_trace(go.Scatter(x=promedio_x_ronda['Ronda'], y=promedio_x_ronda['Valor'], mode='lines+markers',
                                  name='Promedio', line=dict(color=color_ref, width=1, dash='dash'), marker=dict(size=4)))
    fig.update_layout(title=f'Evolución — {titulo}')
    mostrar(fig)
def chart_evolucion_proyeccion(df_proy: pd.DataFrame, metric: str, region: str, titulo: str, team='CADIZ'):
    """Evolución de una métrica que CADIZ proyecta pero que CESIM nunca publica en el RDOS (ej. FCF) --
    a diferencia de chart_evolucion (que grafica el dato REAL de los 7 equipos), acá solo hay una
    serie: la propia proyección de CADIZ, ronda a ronda. Se muestra igual, sin comparación, para no
    perder de vista la tendencia de un indicador que de otro modo quedaría sin ningún gráfico."""
    if df_proy is None:
        return st.info('Sin proyección cargada.')
    sub = df_proy[(df_proy['team'] == team) & (df_proy['metric'] == metric) & (df_proy['region'] == region)].copy()
    sub['value'] = pd.to_numeric(sub['value'], errors='coerce')
    sub = sub.dropna(subset=['value']).sort_values('round')
    if sub.empty:
        return st.info('Sin datos para evolución.')
    fig = go.Figure(go.Scatter(x=sub['round'], y=sub['value'], mode='lines+markers', name=team,
                                line=dict(color=COLOR_MAP.get(team, COLOR_CADIZ), width=3), marker=dict(size=6)))
    fig.update_layout(title=f'Evolución de la proyección — {titulo}', xaxis_title='Ronda')
    mostrar(fig)
def _techo_apilado(y):
    """Techo del eje con ~15% de aire arriba del máximo real -- ver nota en chart_dos_metricas_apiladas
    sobre por qué no alcanza con rangemode='tozero' solo para barras."""
    y_vals = [v for v in y if v is not None and not pd.isna(v)]
    y_max = max(y_vals) if y_vals else 0
    return y_max * 1.15 if y_max > 0 else 1
def chart_dos_metricas_apiladas(titulo, x_a, y_a, nombre_a, color_a, tipo_a,
                                 x_b, y_b, nombre_b, color_b, tipo_b):
    """UN solo gráfico, eje X compartido (misma Ronda), con la primera métrica en el eje Y
    izquierdo y la segunda en el eje Y derecho (doble eje Y superpuesto) -- a pedido explícito del
    equipo, que prefirió esto a la versión anterior (dos paneles apilados, cada uno con su propio
    eje). OJO: esto reintroduce a propósito el patrón de "doble eje Y" que se había evitado acá
    mismo (y que se sacó del todo en Mercado -> Evolución, "Trayectoria de precio y
    características", por el riesgo de sugerir una correlación entre las dos series que no está
    probada) -- queda anotado por si conviene revisar el criterio más adelante, pero el pedido fue
    explícito y puntual para ESTOS gráficos (RRHH y Beneficio vs. Deuda), no una vuelta atrás
    general de criterio."""
    fig = go.Figure()
    def _trace(x, y, nombre, color, tipo, yaxis):
        if tipo == 'bar':
            fig.add_trace(go.Bar(x=x, y=y, name=nombre, marker_color=color, yaxis=yaxis, opacity=0.85))
        else:
            fig.add_trace(go.Scatter(x=x, y=y, name=nombre, mode='lines+markers',
                                      line=dict(color=color, width=3), yaxis=yaxis))
    _trace(x_a, y_a, nombre_a, color_a, tipo_a, 'y1')
    _trace(x_b, y_b, nombre_b, color_b, tipo_b, 'y2')
    fig.update_layout(
        title=titulo,
        yaxis=dict(title=nombre_a, title_font=dict(size=10, color=color_a), tickfont=dict(color=color_a),
                   range=[0, _techo_apilado(y_a)], side='left'),
        yaxis2=dict(title=nombre_b, title_font=dict(size=10, color=color_b), tickfont=dict(color=color_b),
                    range=[0, _techo_apilado(y_b)], overlaying='y', side='right', showgrid=False),
        legend=dict(orientation='h', yanchor='top', y=-0.25, xanchor='center', x=0.5))
    mostrar(fig)
def sparkline(valores, color=None, invertir=False):
    """Minigráfico de tendencia para meter dentro de una tarjeta de KPI.
    Con 12-15 rondas, un número solo no dice nada: la forma de la serie sí.

    Adenda 25: el default ya NO es el lila/rojo de marca -- sparkline() es un helper genérico que no
    sabe de qué equipo es la serie (los dos llamadores lo usan para el equipo EN FOCO, que puede no
    ser CADIZ), así que asumir un color de identidad acá era el mismo bug de fondo que motivó esta
    limpieza de paleta. El default es un gris neutro; quien llama y sabe de qué equipo se trata debe
    pasar `color=COLOR_MAP.get(ese_equipo, ...)` explícito (ver _seccion_resultado_resumen)."""
    serie = [v for v in valores if v is not None and not pd.isna(v)]
    if len(serie) < 2:
        return None
    color = color or MUTED_PALETTE[0]
    fig = go.Figure(go.Scatter(y=serie, mode='lines', line=dict(color=color, width=2),
                                fill='tozeroy', fillcolor=_hex_a_rgba(color, 0.12), hoverinfo='skip'))
    fig.update_layout(height=46, margin=dict(l=0, r=0, t=0, b=0),
                       paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                       xaxis=dict(visible=False), yaxis=dict(visible=False, autorange='reversed' if invertir else True),
                       showlegend=False)
    return fig
def chart_bullet(titulo, valor_fondo, valor_frente, tipo, nombre_fondo='Proyectado', nombre_frente='Real',
                  color_excedente=None):
    """Barra de progreso con desborde apilado: `valor_fondo` (Proyectado, o Punto de Equilibrio) define
    el 100% = el largo de la pista de referencia. `valor_frente` (Real) se dibuja como relleno DENTRO
    de esa pista; si la supera, el excedente se apila como un segmento aparte que sobresale del final
    de la pista en vez de superponer dos barras independientes (diseño anterior) -- así se ve de un
    vistazo si el Real se pasó del Proyectado y por cuánto, no solo que son distintos.

    `color_excedente` tiñe ese segmento de sobra: quien llama puede pasar el color que corresponda a
    si pasarse del Proyectado es favorable o no para esa métrica puntual (ej. superar el Punto de
    Equilibrio es bueno; superar una Deuda CP no planificada, no) -- por defecto un ámbar neutro
    ('se pasó de la referencia', sin prejuzgar si es bueno o malo) para cuando no se sabe.

    Reutilizado por la Fila 2 de la Comparativa Plan vs. Real (Proyectado vs. Real) y por el gráfico
    de Punto de Equilibrio de Operaciones (Volumen de Equilibrio vs. Volumen Real Vendido)."""
    if valor_fondo is None or valor_frente is None:
        return st.info('Sin dato para graficar.')
    plan, real = float(valor_fondo), float(valor_frente)
    color_excedente = color_excedente or COLOR_METRICA['riesgo']
    color_ref = 'rgba(255,255,255,0.5)' if es_modo_oscuro() else 'rgba(26,23,20,0.5)'
    fig = go.Figure()
    if plan > 0:
        # Normalizado a % del Proyectado: la pista siempre mide 100 = Plan, así el "se pasó de largo"
        # es directamente visual (el segmento de excedente empieza justo donde termina la pista) y
        # comparable entre KPIs de escalas muy distintas (USD, unidades, %) en la misma fila de columnas.
        pct_real = real / plan * 100
        dentro = min(pct_real, 100)
        excedente = max(0.0, pct_real - 100)
        fig.add_trace(go.Bar(x=[100], y=[''], orientation='h', name=nombre_fondo, base=0, width=0.55,
                              marker_color='rgba(140,151,166,0.30)', hoverinfo='skip'))
        # COLOR_CADIZ (identidad, no semántica): esta barra "Real" en chart_bullet() siempre es de
        # CADIZ -- es el único equipo con Plan/Proyección cargado, así que Plan vs. Real acá nunca
        # compara a otro equipo (Adenda 25).
        fig.add_trace(go.Bar(x=[dentro], y=[''], orientation='h', name=nombre_frente, base=0, width=0.55,
                              marker_color=COLOR_CADIZ,
                              hovertemplate=f'{nombre_frente}: {_fmt_valor_cg(real, tipo)}<extra></extra>'))
        if excedente > 0:
            fig.add_trace(go.Bar(x=[excedente], y=[''], orientation='h', name='Excedente sobre el plan',
                                  base=100, width=0.55, marker_color=color_excedente,
                                  hovertemplate=f'Excedente: +{pct_real - 100:,.1f} p.p. sobre el plan<extra></extra>'))
        # BUG REPORTADO Y CONFIRMADO (Playwright, caso real "Punto de Equilibrio" con Ronda 1 de
        # CADIZ, donde Real = 302.8% del plan): la pista de referencia (nombre_fondo, gris muy claro)
        # queda TOTALMENTE tapada en cuanto dentro llega a 100 -- porque barmode='overlay' dibuja la
        # barra de "Real" exactamente encima, mismo ancho y misma fila. El resultado: solo se ven los
        # colores de "Real" y "Excedente", la referencia desaparece del todo salvo por su nombre en la
        # leyenda -- que es justo el reporte del equipo ("el punto de equilibrio no se ve"). La línea
        # punteada de acá abajo YA marcaba el 100% pero sin ninguna etiqueta -- ahora lleva el nombre
        # de la referencia (nombre_fondo) escrito directamente sobre el gráfico, así el dato clave
        # (dónde está el equilibrio) queda visible pase lo que pase con las barras de abajo.
        fig.add_vline(x=100, line_dash='dot', line_width=1.5, line_color=color_ref,
                      annotation_text=nombre_fondo, annotation_position='top',
                      annotation_font_size=11, annotation_font_color=color_ref)
        fig.update_layout(barmode='overlay', title=titulo, xaxis_title=None,
                           xaxis=dict(ticksuffix='%', range=[0, max(100, pct_real) * 1.15]),
                           legend=dict(orientation='h', yanchor='top', y=-0.25, xanchor='center', x=0.5))
        mostrar(fig, ocultar_eje_valores='y')
        cumplimiento = f' ({pct_real:,.1f}% del plan)'
    else:
        # Proyectado <= 0: expresarlo como % del plan no tiene sentido (división por ~0) -- se cae a
        # una barra simple en valor absoluto, sin pista de referencia, y se avisa en el caption.
        fig.add_trace(go.Bar(x=[real], y=[''], orientation='h', name=nombre_frente, marker_color=COLOR_CADIZ, width=0.55))
        fig.update_layout(title=titulo, xaxis_title=None, showlegend=False)
        mostrar(fig, ocultar_eje_valores='y')
        cumplimiento = ' — Proyectado ≤ 0, no expresable como % de avance'
    st.caption(f'{nombre_fondo}: {_fmt_valor_cg(plan, tipo)} · {nombre_frente}: {_fmt_valor_cg(real, tipo)}{cumplimiento}')

def chart_bullet_panel(items):
    """Adenda 23 (a pedido del equipo): versión COMPACTA de chart_bullet() para cuando hay VARIOS
    KPIs para comparar Plan vs. Real a la vez (ej. la Fila 2 de 'Comparativa Plan vs. Real' de
    Finanzas -- EBITDA, Margen bruto, ROS, etc.). Antes cada KPI dibujaba su PROPIO chart_bullet()
    en una grilla de columnas -- con 8 KPIs eso son 2 filas de gráficos de altura completa, mucho
    espacio vertical para escanear poco. Acá es UN solo gráfico con una fila horizontal por KPI
    (mismo lenguaje visual: pista de Plan normalizada a 100%, relleno de Real, excedente apilado si
    se pasa), todas comparten el mismo eje X (% del Plan) y se leen de arriba hacia abajo.

    `items`: lista de dicts {'label', 'plan', 'real', 'tipo', 'color_excedente'} en el orden en que
    deben listarse (de arriba hacia abajo). Los que tienen Plan<=0 (no expresable como % de avance,
    ver chart_bullet) se devuelven aparte para que el llamador los grafique con chart_bullet() normal
    -- son la excepción, no vale la pena forzarlos al panel compacto."""
    validos = [it for it in items if it.get('plan') not in (None, 0) and it['plan'] > 0 and it.get('real') is not None]
    if not validos:
        return
    color_ref = 'rgba(255,255,255,0.5)' if es_modo_oscuro() else 'rgba(26,23,20,0.5)'
    fig = go.Figure()
    max_pct = 100.0
    filas = [it['label'] for it in validos]
    for i, it in enumerate(validos):
        plan, real, tipo = float(it['plan']), float(it['real']), it['tipo']
        pct = real / plan * 100
        max_pct = max(max_pct, pct)
        dentro = min(pct, 100)
        excedente = max(0.0, pct - 100)
        color_exc = it.get('color_excedente') or COLOR_METRICA['riesgo']
        fig.add_trace(go.Bar(x=[100], y=[it['label']], orientation='h', base=0, width=0.55,
                              name='Plan (100%)', marker_color='rgba(140,151,166,0.30)',
                              showlegend=(i == 0), hoverinfo='skip'))
        fig.add_trace(go.Bar(x=[dentro], y=[it['label']], orientation='h', base=0, width=0.55,
                              name='Real', marker_color=COLOR_CADIZ, showlegend=(i == 0),
                              hovertemplate=f"Real: {_fmt_valor_cg(real, tipo)} ({pct:,.0f}% del plan)<extra></extra>"))
        if excedente > 0:
            fig.add_trace(go.Bar(x=[excedente], y=[it['label']], orientation='h', base=100, width=0.55,
                                  name='Excedente sobre el plan', marker_color=color_exc, showlegend=(i == 0),
                                  hovertemplate=f'Excedente: +{pct - 100:,.1f} p.p. sobre el plan<extra></extra>'))
        fig.add_annotation(x=max(pct, 100) + max_pct * 0.03, y=it['label'], xanchor='left', showarrow=False,
                            text=f"{_fmt_valor_cg(real, tipo)} · {pct:,.0f}%", font=dict(size=10, color=color_ref))
    fig.add_vline(x=100, line_dash='dot', line_width=1.5, line_color=color_ref,
                  annotation_text='Plan', annotation_position='top',
                  annotation_font_size=11, annotation_font_color=color_ref)
    # Adenda 24 -- BUG REPORTADO Y CONFIRMADO (leyenda y anotaciones superpuestas a las barras, con
    # apenas 1-2 filas): el alto total se calculaba SIN separar "área de ploteo" de "margen para la
    # leyenda" -- mostrar() reserva un margen inferior FIJO (110px, ver esa función) para la leyenda,
    # pero con pocas filas el 'altura' de acá (70+36xn) apenas superaba ese margen fijo, dejando un
    # área de ploteo casi nula. Como la posición de la leyenda (fig.layout.legend.y) es una FRACCIÓN
    # del área de ploteo (no de la figura completa), un área de ploteo de pocos píxeles convertía
    # cualquier offset negativo en apenas 1-2 píxeles reales -- la leyenda terminaba pegada al eje X,
    # justo encima de las barras y las anotaciones de valor. Ahora se separan explícitamente: un área
    # de ploteo mínima garantizada (para que las filas no queden apretadas) + un margen inferior propio
    # (en vez del fijo de mostrar()) dimensionado para la leyenda, y la posición de la leyenda se
    # calcula en función del área de ploteo real (no de un % arbitrario) para que caiga siempre dentro
    # de ese margen, sea cual sea la cantidad de filas.
    n = len(validos)
    plot_alto = max(120, 34 * n)              # área de ploteo: nunca menos de 120px, aunque haya 1 fila
    margen_b = 70                              # espacio fijo para la leyenda horizontal de abajo
    altura = min(620, 45 + margen_b + plot_alto)
    y_leyenda = -34 / plot_alto                # ~34px por debajo del eje X, constante en píxeles
    fig.update_layout(barmode='overlay', title='Proyectado vs. Real',
                       xaxis=dict(ticksuffix='%', range=[0, max_pct * 1.4]), xaxis_title=None,
                       yaxis=dict(categoryorder='array', categoryarray=list(reversed(filas))),
                       legend=dict(orientation='h', yanchor='top', y=y_leyenda, xanchor='center', x=0.5))
    mostrar(fig, altura=altura, margen_b=margen_b)

# --- COMPARATIVA PLAN VS. REAL: Proyectado (CADIZ_Gestion_v2.xlsx, vía export_proyeccion.py) vs.
# Real (RDOS de CESIM, ya parseados más arriba por cesim_parser) ---
def _fmt_valor_cg(v, tipo):
    if v is None:
        return '—'
    if tipo == 'usd':
        return format_num(v)
    if tipo == 'unidades':
        return f'{format_num(v)} u.'
    if tipo == 'usd_accion':
        return f'USD {v:,.2f}'
    if tipo == 'ratio':
        return f'{v * 100:,.1f}%'
    return str(v)
def _fmt_delta_cg(v):
    if v['estado'] != 'ok' or v['gap_abs'] is None:
        return None
    if v['tipo'] == 'usd':
        # Ronda con status=REAL en ambos lados (histórico ya migrado) da un gap de centavos por
        # redondeo de punto flotante entre el motor Python y el motor Excel -- no es un gap real y
        # mostrarlo como "-0" confundiría; se redondea al dólar y se omite si queda en 0.
        gap_redondeado = round(v['gap_abs'])
        return format_num(gap_redondeado) if gap_redondeado != 0 else None
    if v['tipo'] == 'unidades':
        gap_redondeado = round(v['gap_abs'])
        return f'{gap_redondeado:+,.0f} u.' if gap_redondeado != 0 else None
    if v['tipo'] == 'usd_accion':
        return f"{v['gap_abs']:+,.2f}"
    if v['tipo'] == 'ratio':
        return f"{v['gap_abs'] * 100:+.1f} p.p."
    return None
_DELTA_COLOR_CG = {'real_mayor': 'normal', 'real_menor': 'inverse', None: 'off'}
# Mismo mapeo de favorabilidad que _DELTA_COLOR_CG, pero como color de relleno para el segmento de
# excedente de chart_bullet(): si más Real es mejor (real_mayor), pasarse del plan es favorable
# (verde). Para 'real_menor' (ej. deuda no planificada) se probó primero con COLOR_NEGATIVE (rojo) para
# marcarlo como desfavorable, pero el relleno "Real" del propio bullet ya usa un color fijo (COLOR_CADIZ) --
# el segmento de excedente quedaba muy parecido, poco distinguible (ver test visual). Ámbar neutro
# funciona para ambos casos sin esa colisión: sigue leyéndose como "atención, se pasó de la
# referencia" sea o no favorable (la flecha de color del delta, un renglón más arriba, ya dice si eso
# es bueno o malo).
_COLOR_EXCEDENTE_CG = {'real_mayor': COLOR_POSITIVE, 'real_menor': COLOR_METRICA['riesgo'], None: COLOR_METRICA['riesgo']}
def panel_comparativa_plan_real(df_todas_rondas, ronda_snapshot, crosswalk=None, key_suffix='', mostrar_directo=False, mostrar_metricas=True, mostrar_extras=True):
    """Botón 'Comparativa Plan vs. Real': Proyectado (nuestro modelo) vs. Real (RDOS de CESIM) para los
    KPIs del crosswalk dado. Un KPI sin proyección para esta ronda (el modelo no lo cubre, o es una
    ronda de Práctica que el modelo no proyecta) o sin dato real todavía (CESIM no publicó esta
    ronda) no se omite en silencio: cae al gráfico de evolución de esa variable.

    mostrar_directo=True se salta el toggle y renderiza directo -- para cuando esta es la ÚNICA
    razón de estar en la página (p.ej. CESIM todavía no publicó ningún RDOS de esta ronda: no hay
    nada más que mostrar en el resto de las secciones, así que no tiene sentido esconder esto detrás
    de un clic extra).

    mostrar_metricas=False (Adenda 26, a pedido del equipo solo para Finanzas: "dejar como elemento
    protagonista exclusivo el panel de barras horizontales apiladas") -- oculta la Fila 1 de tarjetas
    st.metric (Proyectado/Real por KPI) y deja directo el panel compacto de barras (chart_bullet_panel).

    mostrar_extras=False (Adenda 27, mismo pedido pero llevado más lejos: "eliminar todo esos
    gráficos de comparativa vs real en finanzas, solo nos quedamos con la de comparativa vs real de
    barras horizontales") -- además de lo anterior, oculta el chart_bullet() suelto que se grafica
    aparte para los KPIs con Plan<=0 (ej. "Deuda CP no planificada") y todo el bloque de abajo (los
    gráficos de evolución de KPIs sin gap posible, ej. "Evolución — Margen bruto", y su caption). Deja
    ÚNICAMENTE el panel chart_bullet_panel ("Proyectado vs. Real") como elemento en pantalla.

    Ambos parámetros solo se pasaron en False desde el llamado de Finanzas: Resultados/Mercado/
    Operaciones siguen mostrando todo como antes -- no se tocó su comportamiento porque no fue lo que
    pidió el equipo (si se quiere el mismo cambio ahí, avisar)."""
    crosswalk = crosswalk or CROSSWALK_FINANZAS
    if not mostrar_directo:
        activo = st.toggle('📊 Comparativa Plan vs. Real', key=f'cg_toggle_{key_suffix}')
        if not activo:
            return
    df_proy = get_proyeccion()
    if df_proy is None:
        st.info('Todavía no hay ningún `Cadiz_proyeccion_R{N}.xlsx` en `data/decisiones/` (o no se '
                'pudo leer la hoja `DATA_EXPORT`) — generalo con el botón "Generar Excel de gestión" '
                'de la sidebar y commiteá el resultado en esa carpeta para ver la Comparativa Plan vs. Real.')
        return
    ronda_num = ronda_a_num(ronda_snapshot)
    claves_no_publicadas = {k for k, spec in crosswalk.items() if spec.get('real_no_publicado')}
    gaps_todos = calcular_gaps(df_todas_rondas, df_proy, ronda_nombre=ronda_snapshot, ronda_num=ronda_num, crosswalk=crosswalk)
    gaps = {k: v for k, v in gaps_todos.items() if k not in claves_no_publicadas}
    sin_publicar = {k: v for k, v in gaps_todos.items() if k in claves_no_publicadas}
    con_gap = {k: v for k, v in gaps.items() if v['estado'] in ('ok', 'sin_real')}
    sin_gap = {k: v for k, v in gaps.items() if v['estado'] in ('sin_datos', 'sin_proyeccion')}
    with st.container(border=True):
        st.markdown(f'**Comparativa Plan vs. Real — {ronda_snapshot}**')
        if not con_gap:
            st.caption('CADIZ no tiene una proyección cargada para esta ronda en el modelo de gestión — '
                       'ver la evolución de cada indicador más abajo.')
        else:
            if mostrar_metricas:
                cols = st.columns(min(4, len(con_gap)))
                for i, (clave, v) in enumerate(con_gap.items()):
                    with cols[i % len(cols)]:
                        if v['estado'] == 'sin_real':
                            st.metric(v['label'], _fmt_valor_cg(v['proyectado'], v['tipo']))
                            st.caption('Proyectado (CADIZ) — real de CESIM pendiente')
                        else:
                            st.metric(v['label'], _fmt_valor_cg(v['real'], v['tipo']),
                                       delta=_fmt_delta_cg(v), delta_color=_DELTA_COLOR_CG[v['gap_favorable']])
                            st.caption(f"Real — proyectado {_fmt_valor_cg(v['proyectado'], v['tipo'])}")
            # Fila 2: bullet charts (Proyectado vs. Real superpuestos) -- solo para los KPIs que
            # tienen AMBOS valores (estado='ok'); 'sin_real' ya se ve en la tarjeta de arriba, no hay
            # nada que superponer todavía.
            # tipo='texto' (ej. Calificación crediticia) no tiene sentido como barra -- ya se ve
            # completo en la tarjeta de Fila 1 (valor + "Real — proyectado ...").
            con_ambos = {k: v for k, v in con_gap.items() if v['estado'] == 'ok' and v['tipo'] != 'texto'}
            if con_ambos:
                st.markdown('###### Proyectado vs. Real')
                # Adenda 23 (a pedido del equipo, storytelling de Finanzas): panel COMPACTO de una
                # sola figura con una fila horizontal por KPI, en vez de un chart_bullet() por
                # columna -- mucho menos alto para escanear varios KPIs (EBITDA, ROS, etc.) de
                # arriba hacia abajo. Mismo criterio de favorabilidad que ya colorea la flecha del
                # delta de la tarjeta de arriba (_DELTA_COLOR_CG), pasado como color del excedente.
                items_panel = [
                    {'label': v['label'], 'plan': v['proyectado'], 'real': v['real'], 'tipo': v['tipo'],
                     'color_excedente': _COLOR_EXCEDENTE_CG.get(v['gap_favorable'])}
                    for v in con_ambos.values()
                ]
                chart_bullet_panel(items_panel)
                # Plan<=0 no es expresable como % de avance (ver chart_bullet) -- caso raro, se
                # grafica aparte con el chart_bullet() de siempre en vez de forzarlo al panel.
                # Adenda 27: mostrar_extras=False (Finanzas) lo saca -- el equipo pidió dejar SOLO
                # el panel de barras de arriba, sin excepciones sueltas.
                if mostrar_extras:
                    for clave, v in con_ambos.items():
                        if v['proyectado'] is not None and v['proyectado'] <= 0:
                            chart_bullet(v['label'], v['proyectado'], v['real'], v['tipo'],
                                         color_excedente=_COLOR_EXCEDENTE_CG.get(v['gap_favorable']))
        # Adenda 27: mostrar_extras=False (Finanzas) saca también todo este bloque de abajo (los
        # gráficos de evolución de KPIs sin gap posible + su caption) -- mismo pedido de "dejar solo
        # la de barras horizontales".
        if mostrar_extras and (sin_gap or sin_publicar):
            st.caption('Sin comparación posible para estos indicadores (no forman parte de la '
                       'proyección de CADIZ, es una ronda de práctica, o CESIM no publica ese dato en '
                       'el RDOS) — se muestra su evolución:')
    if mostrar_extras and (sin_gap or sin_publicar):
        total = len(sin_gap) + len(sin_publicar)
        cols_ev = st.columns(min(2, total))
        i = 0
        for clave, v in sin_gap.items():
            spec = crosswalk[clave]['real']
            sub = df_todas_rondas[(df_todas_rondas['Estado'] == spec['estado']) & (df_todas_rondas['Metrica'] == spec['metrica'])]
            if spec.get('seccion'):
                sub = sub[sub['Seccion'] == spec['seccion']]
            with cols_ev[i % len(cols_ev)]:
                chart_evolucion(sub, v['label'])
            i += 1
        for clave, v in sin_publicar.items():
            spec = crosswalk[clave]['proyeccion']
            with cols_ev[i % len(cols_ev)]:
                chart_evolucion_proyeccion(df_proy, spec['metric'], spec['region'], v['label'])
                st.caption('CESIM no publica este dato en el RDOS — evolución de la proyección propia de CADIZ.')
            i += 1

# --- Fila 3 de la Comparativa Plan vs. Real: gráficos de GAP/varianza específicos por sección
# (Adenda 12). Cada uno se llama desde la pestaña "Comparativa Plan vs. Real" de su propia sección
# (no desde panel_comparativa_plan_real, que es genérico y no conoce estas métricas de grano fino
# por mercado/tecnología/área) -- y solo tiene sentido con team=CADIZ (son cruces contra SU propia
# proyección en CADIZ_Gestion_v2.xlsx). ---
def fila3_resultados_ingresos(df_all, ronda_snapshot, ronda_num, df_proy):
    st.markdown('###### Análisis de Desvíos de Ingresos — Precio / Volumen / Mix')
    mercado_sel = st.selectbox('Mercado', _MERCADOS_GAP, key='sel_cg_resultados_mercado')
    datos = precio_volumen_mercado(df_all, df_proy, ronda_snapshot, ronda_num, mercado_sel, team=MY_COMPANY)
    if not datos:
        st.info(f'CADIZ no tiene datos de precio/volumen (Plan o Real) en {mercado_sel} para {ronda_snapshot}.')
        return
    if all(d['precio_plan'] is None for d in datos.values()):
        # Sin ESTO, el waterfall tomaría Plan=0 como línea base y le atribuiría el 100% de los
        # Ingresos Reales al "Desvío por Precio" -- un artefacto de la falta de dato, no un desvío
        # real. El modelo de CADIZ solo proyecta desde Ronda 2 (Ronda 0/1 no tienen Plan cargado).
        st.info(f'CADIZ no tiene una proyección de Precio/Volumen cargada para {mercado_sel} en {ronda_snapshot} '
                '— el modelo de gestión proyecta recién desde Ronda 2, no hay Plan con el que comparar.')
        return
    var = variacion_precio_volumen_mix(datos)
    if not var['reconciliacion_ok']:
        st.warning('El desglose Precio/Volumen/Mix no reconcilia exactamente con la variación de Ingresos — revisar.')
    etapas = [('Ingresos Proyectados', var['ingresos_plan']), ('Desvío por Precio', var['var_precio']),
              ('Desvío por Volumen', var['var_volumen']), ('Desvío por Mix', var['var_mix']),
              ('Ingresos Reales', var['ingresos_real'])]
    fig = go.Figure(go.Waterfall(
        orientation='v', measure=['absolute', 'relative', 'relative', 'relative', 'total'],
        x=[e[0] for e in etapas], y=[e[1] for e in etapas],
        text=_texto_cascada([v for _, v in etapas]), textposition='outside',
        # Adenda 25 (a pedido del equipo): TODOS los gráficos en cascada de la app usan ahora la
        # misma convención universal -- verde = lo que suma (increasing), rojo = lo que resta
        # (decreasing) -- en vez de favorable/desfavorable por caso. Antes acá era verde/ámbar.
        increasing={'marker': {'color': COLOR_POSITIVE}}, decreasing={'marker': {'color': COLOR_NEGATIVE}},
        totals={'marker': {'color': COLOR_CADIZ}}))  # total = "Ingresos Reales" -- identidad CADIZ, no sentimiento
    fig.update_layout(title=f'Ingresos — Proyectado vs. Real, {mercado_sel} ({_MONEDA_MERCADO_GAP[mercado_sel]})')
    mostrar(fig, ocultar_eje_valores='y')
    st.caption(f'Verde = desvío que sumó Ingresos; rojo = desvío que restó. % entre paréntesis = sobre los Ingresos '
               f'Proyectados. En moneda nativa de {mercado_sel} ({_MONEDA_MERCADO_GAP[mercado_sel]}) — no se '
               'convierte a USD para no asumir un tipo de cambio que el simulador no publica.')

def fila3_mercado_cuota_objetivo(df_all, ronda_snapshot, ronda_num, df_proy):
    st.markdown('###### Cuota de mercado — Proyectado vs. Real, por tecnología')
    mercado_sel = st.selectbox('Mercado', _MERCADOS_GAP, key='sel_cg_mercado_cuota')
    datos = cuota_mercado_objetivo_vs_real(df_all, df_proy, ronda_snapshot, ronda_num, mercado_sel, team=MY_COMPANY)
    if not datos:
        st.info(f'Sin datos de cuota objetivo/real en {mercado_sel} para {ronda_snapshot}.')
        return
    filas = []
    for tech, d in datos.items():
        # "Proyectado" en todos lados (antes decía "Objetivo (Plan)" acá, distinto del resto de la
        # Comparativa) -- mismo término que chart_bullet, la Fila 1 de KPIs y las demás Filas 3.
        if d['objetivo'] is not None:
            filas.append({'Tecnología': tech, 'Tipo': 'Proyectado', 'Cuota': d['objetivo'] * 100})
        if d['real'] is not None:
            filas.append({'Tecnología': tech, 'Tipo': 'Real', 'Cuota': d['real'] * 100})
    if not filas:
        return st.info('Sin datos suficientes.')
    dfc = pd.DataFrame(filas)
    fig = px.bar(dfc, x='Tecnología', y='Cuota', color='Tipo', barmode='group',
                 color_discrete_map={'Proyectado': MUTED_PALETTE[0], 'Real': COLOR_CADIZ},
                 text=dfc['Cuota'].apply(lambda v: f'{v:.1f}%'), title=f'Cuota de mercado — {mercado_sel}, {ronda_snapshot}')
    fig.update_traces(textposition='outside', cliponaxis=False)
    fig.update_layout(yaxis_title='% del mercado total')
    mostrar(fig)
    st.caption('Cuota = ventas de CADIZ en esa tecnología / tamaño TOTAL del mercado (las 4 tecnologías) — '
               'misma convención en Plan y Real (distinta de la que el RDOS publica directo por tecnología, '
               'ver nota metodológica en gap_analysis.cuota_mercado_objetivo_vs_real). '
               'Ojo con leerla junto al gráfico de "Demanda estimada" de más arriba: el Proyectado de ESTA '
               'cuota se calculó contra el tamaño de mercado que el modelo había asumido al planificar; el '
               'Real se calcula contra el tamaño que terminó publicando CESIM. Si la demanda real vino más '
               'chica que la proyectada, la cuota puede salir MÁS ALTA que el objetivo aunque el volumen '
               'propio en unidades haya sido menor al planeado — no es una contradicción entre los dos '
               'gráficos, es la misma torta más chica repartida distinto.')

def _costo_fabricacion_ponderado(datos_area):
    """A partir de costo_unitario_area(): costo unitario de fabricación PONDERADO por producción REAL
    (propia + contratada), Plan vs. Real, y el GAP que aporta cada tecnología al total — usa los
    MISMOS pesos (producción real) para ponderar Plan y Real, así el desglose por tecnología
    reconcilia EXACTO con la diferencia total (no es una aproximación)."""
    filas = []
    prod_total = 0.0
    for tech, d in datos_area.items():
        prod_p, prod_t = d['prod_propia_real'], d['prod_terc_real']
        prod_tech = prod_p + prod_t
        if prod_tech <= 0:
            continue
        plan_tech = ((d['cu_propia_plan'] or 0) * prod_p + (d['cu_terc_plan'] or 0) * prod_t) / prod_tech
        real_tech = ((d['cu_propia_real'] or 0) * prod_p + (d['cu_terc_real'] or 0) * prod_t) / prod_tech
        filas.append({'tech': tech, 'prod': prod_tech, 'plan': plan_tech, 'real': real_tech})
        prod_total += prod_tech
    if prod_total <= 0:
        return None
    plan_pond = sum(f['prod'] * f['plan'] for f in filas) / prod_total
    real_pond = sum(f['prod'] * f['real'] for f in filas) / prod_total
    gaps = [{'tech': f['tech'], 'gap': (f['real'] - f['plan']) * f['prod'] / prod_total} for f in filas]
    return {'plan_pond': plan_pond, 'real_pond': real_pond, 'gaps': gaps}

def fila3_operaciones_gap_fabricacion(df_all, ronda_snapshot, ronda_num, df_proy):
    st.markdown('###### Desvío en Costo Unitario de Fabricación (Proyectado vs. Real)')
    st.caption('Alcance: solo costo de FABRICACIÓN (propia + contratada), ponderado por producción real. '
               'Transporte/aranceles y promoción se reportan por mercado de destino (no por área de origen) '
               'y no están incluidos acá — por eso este desvío no reconcilia el 100% de la Contribución '
               'Marginal unitaria completa.')
    area_sel = st.selectbox('Área de producción', _AREAS_GAP, key='sel_cg_operaciones_area')
    datos = costo_unitario_area(df_all, df_proy, ronda_snapshot, ronda_num, area_sel, team=MY_COMPANY)
    if not datos:
        return st.info(f'Sin datos de costo unitario de fabricación en {area_sel} para {ronda_snapshot}.')
    if all(d['cu_propia_plan'] is None and d['cu_terc_plan'] is None for d in datos.values()):
        # Mismo motivo que en el waterfall de Ingresos: sin esto, Proyectado=0 le atribuiría el 100%
        # del costo real a un "Desvío" que en realidad es solo ausencia de proyección (Ronda 0/1,
        # antes de que el modelo de CADIZ empezara a proyectar en Ronda 2).
        return st.info(f'CADIZ no tiene una proyección de costo unitario cargada para {area_sel} en {ronda_snapshot} '
                        '— el modelo de gestión proyecta recién desde Ronda 2, no hay Plan con el que comparar.')
    pond = _costo_fabricacion_ponderado(datos)
    if not pond:
        return st.info(f'Sin datos de producción real en {area_sel} para {ronda_snapshot}.')
    # "Desvío {tecnología}" (antes "GAP {tecnología}", en inglés y sin explicar qué significa): cuánto
    # empujó ESA tecnología al costo unitario PONDERADO total, de Proyectado a Real -- no es "cuánto le
    # costó de más esa tecnología en el vacío", es su aporte a la diferencia total, ponderado por su
    # propia producción real. La suma de todos los "Desvío X" + Costo Proyectado da EXACTO el Costo
    # Real (reconciliación por construcción, ver _costo_fabricacion_ponderado) -- por eso tiene sentido
    # como cascada/waterfall y no como barras sueltas.
    etapas = [('Costo Proyectado', pond['plan_pond'])]
    for g in pond['gaps']:
        if abs(g['gap']) > 1e-9:
            etapas.append((f"Desvío {g['tech']}", g['gap']))
    etapas.append(('Costo Real', pond['real_pond']))
    fig = go.Figure(go.Waterfall(
        orientation='v', measure=['absolute'] + ['relative'] * (len(etapas) - 2) + ['total'],
        x=[e[0] for e in etapas], y=[e[1] for e in etapas],
        text=_texto_cascada([v for _, v in etapas]), textposition='outside',
        # Adenda 25 (a pedido del equipo): convención universal para TODOS los waterfalls de la
        # app -- verde = lo que suma (increasing), rojo = lo que resta (decreasing), siempre, sin
        # excepción por favorable/desfavorable. Antes acá era al revés (ámbar para el costo que
        # SUBE, verde para el que BAJA) porque se leía como favorable/desfavorable para el costo
        # -- pero eso rompía la consistencia con el resto de los waterfalls de la app, que es lo
        # que el equipo pidió priorizar. OJO: como consecuencia, acá una "Desvío X" que aumenta el
        # costo unitario (malo para CADIZ) se ve en VERDE (porque suma a la barra), y una que lo
        # reduce (bueno) se ve en ROJO (porque resta) -- es la convención universal pedida, ya no
        # favorable/desfavorable para esta métrica en particular.
        increasing={'marker': {'color': COLOR_POSITIVE}}, decreasing={'marker': {'color': COLOR_NEGATIVE}},
        totals={'marker': {'color': COLOR_CADIZ}}))  # total = "Costo Real" -- identidad CADIZ, no sentimiento
    fig.update_layout(title=f'Costo unitario de fabricación — {area_sel}, {ronda_snapshot}')
    mostrar(fig, ocultar_eje_valores='y')
    st.caption('Cada barra "Desvío {tecnología}" es cuánto empujó esa tecnología el costo unitario ponderado '
               'total, de Proyectado a Real (ponderado por su propia producción real) — no el costo de esa '
               'tecnología en sí. Verde = suma al costo (lo aumenta); rojo = resta (lo reduce) — convención '
               'universal de color de la app, no indica si es favorable o no para CADIZ. % entre paréntesis = '
               'sobre el Costo Proyectado. Costo Proyectado + todos los desvíos = Costo Real, exacto.')

def fila3_finanzas_flujo_caja(df_all, ronda_snapshot, ronda_num, df_proy):
    # Adenda 23 (a pedido del equipo): reemplaza las barras agrupadas (Proyectado vs. Real, un grupo
    # de barras por componente) por un puente/Cascada (Waterfall) que cuenta la historia de cómo se
    # llega de CFO a CFI a CFF hasta la Variación Neta de Caja de la ronda -- una Cascada por
    # Proyectado y otra por Real, para poder seguir mirando ambas lado a lado.
    st.markdown('###### Composición del Flujo de Caja — Puente CFO → CFI → CFF (Global)')
    res = flujo_caja_plan_real_global(df_all, df_proy, ronda_snapshot, ronda_num, team=MY_COMPANY)
    plan, real = res['plan'], res['real']

    def _waterfall_flujo(titulo, valores):
        if all(valores.get(k) is None for k in ('cfo', 'cfi', 'cff')):
            return False
        cfo, cfi, cff = valores.get('cfo') or 0.0, valores.get('cfi') or 0.0, valores.get('cff') or 0.0
        total = cfo + cfi + cff
        etapas = [('CFO', cfo), ('CFI', cfi), ('CFF', cff), ('= Variación Neta de Caja', total)]
        fig = go.Figure(go.Waterfall(
            orientation='v', measure=['absolute', 'relative', 'relative', 'total'],
            x=[e[0] for e in etapas], y=[v for _, v in etapas],
            text=_texto_cascada([v for _, v in etapas], base=cfo if cfo not in (0, None) else None), textposition='outside',
            # Adenda 25: convención universal de color de la app -- verde = suma (increasing),
            # rojo = resta (decreasing), en todos los waterfalls, sin excepción por favorable/
            # desfavorable. Antes decreasing usaba COLOR_METRICA['riesgo'] (ámbar).
            increasing={'marker': {'color': COLOR_POSITIVE}}, decreasing={'marker': {'color': COLOR_NEGATIVE}},
            totals={'marker': {'color': COLOR_POSITIVE if total >= 0 else COLOR_NEGATIVE}}))
        fig.update_layout(title=titulo)
        mostrar(fig, ocultar_eje_valores='y')
        return True

    col_p, col_r = st.columns(2)
    with col_p:
        if not _waterfall_flujo(f'Proyectado — {ronda_snapshot}', plan):
            st.info('CADIZ no proyecta Flujo de Caja para esta ronda en el modelo de gestión (recién desde Ronda 2).')
    with col_r:
        if not _waterfall_flujo(f'Real — {ronda_snapshot}', real):
            st.info('CESIM todavía no publicó el RDOS real de esta ronda.')
    st.caption('Cada barra es cuánta caja aportó o consumió ese bloque de actividades (Operativas / Inversión / '
               'Financieras) hasta llegar a la Variación Neta de Caja de la ronda — no es el saldo de caja del '
               'Balance, es el cambio de esta ronda puntual.')

def serie_metrica(estado, metrica, empresa=None, hasta_orden=None, seccion=None):
    """Serie histórica de una métrica para un equipo, ordenada por ronda.
    OJO con 'seccion': hay métricas cuyo nombre se repite dentro del mismo estado
    (ej. 'Total' aparece en Accionistas, Acreedores, Gobierno, Personal y Proveedores
    dentro de Creación de Valor). Sin filtrar por sección se mezclan cinco series distintas."""
    d = df_all[(df_all['Estado'] == estado) & (df_all['Metrica'] == metrica)].copy()
    if seccion: d = d[d['Seccion'] == seccion]
    if empresa: d = d[d['Empresa'] == empresa]
    if d.empty: return []
    d['Valor'] = num(d['Valor'])
    if hasta_orden is not None:
        d = d[d['Ronda_Orden'] <= hasta_orden]
    return d.sort_values('Ronda_Orden')['Valor'].tolist()

def ronda_anterior_de(ronda):
    """Adenda 24 -- helper genérico para el patrón 'variación vs. ronda anterior' que ya usaba
    _seccion_resultado_resumen() de forma manual (ver más abajo) -- se extrae acá para reusarlo en
    todos los KPIs del tablero que necesiten ese delta, sin repetir la lógica de Ronda_Orden en cada
    lugar. Devuelve el NOMBRE de la ronda anterior (ej. 'Ronda 1'), o None si `ronda` es la primera
    ronda disponible en los datos cargados (no hay anterior con la cual comparar)."""
    ordenes = sorted(df['Ronda_Orden'].dropna().unique())
    sub_act = df[df['Ronda'] == ronda]['Ronda_Orden']
    if sub_act.empty or sub_act.iloc[0] not in ordenes:
        return None
    idx = ordenes.index(sub_act.iloc[0])
    if idx == 0:
        return None
    sub_prev = df[df['Ronda_Orden'] == ordenes[idx - 1]]['Ronda']
    return sub_prev.iloc[0] if not sub_prev.empty else None

def delta_vs_ronda_anterior(estado, metrica, empresa=None, ronda_actual=None, seccion=None):
    """Adenda 25 -- delta % de una métrica para `empresa` entre `ronda_actual` (default:
    ronda_snapshot) y la ronda INMEDIATAMENTE ANTERIOR (vía ronda_anterior_de) -- generaliza el
    patrón 'vs. ronda anterior' para poder pedirlo de cualquier Estado/Métrica, igual que delta_raw
    (definido dentro de cada sección) generaliza 'vs. promedio de industria'. None si no hay ronda
    anterior en los datos cargados, o si falta el dato en cualquiera de las dos rondas.
    Devuelve (delta_pct, ronda_prev) -- se necesita el nombre de la ronda anterior para el texto
    ('vs Ronda 1'), así que se devuelve junto con el número en vez de forzar un segundo llamado a
    ronda_anterior_de()."""
    ronda_actual = ronda_actual or ronda_snapshot
    empresa = empresa or empresa_analisis
    ronda_prev = ronda_anterior_de(ronda_actual)
    if not ronda_prev:
        return None, None
    d_act = df[(df['Estado'] == estado) & (df['Ronda'] == ronda_actual)]
    d_prev = df[(df['Estado'] == estado) & (df['Ronda'] == ronda_prev)]
    if seccion:
        d_act = d_act[d_act['Seccion'] == seccion]
        d_prev = d_prev[d_prev['Seccion'] == seccion]
    val_act, val_prev = valor_de(d_act, metrica, empresa), valor_de(d_prev, metrica, empresa)
    if val_act is None or val_prev is None or pd.isna(val_act) or pd.isna(val_prev) or val_prev == 0:
        return None, ronda_prev
    return (val_act - val_prev) / abs(val_prev) * 100, ronda_prev

def kpi_con_tendencia(col, label, valor_txt, serie, delta=None, color=None, invertir=False):
    """Tarjeta de KPI + sparkline debajo, para ver nivel y tendencia sin cambiar de sección."""
    with col:
        st.metric(label, valor_txt, delta=delta)
        fig = sparkline(serie, color=color, invertir=invertir)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

def _icono_delta(texto):
    """Adenda 25: ícono 🔺/🔻 según el signo LITERAL del texto de delta (ya formateado con signo,
    ej. '+4.3% vs Prom' / '-2.0% vs Ronda 1') -- el color (verde/rojo) sigue viniendo de
    'favorable', que puede estar invertido para métricas donde bajar es bueno (Deuda, costos).
    Así un delta de Deuda que BAJA se ve con 🔻 (bajó, literal) pero en VERDE (es favorable)."""
    t = texto.strip()
    if t.startswith('+'): return '🔺 '
    if t.startswith('-') or t.startswith('−'): return '🔻 '
    return ''

def kpi_banda_oscura(items):
    """Banda oscura para los 2-3 KPIs de valor MÁS importantes de la ronda (los que le interesan
    a un accionista) -- los separa visualmente del resto de tarjetas claras en vez de competir al
    mismo nivel, estilo el bloque "Monto Pendiente / A vencer / Ya pagado" de la referencia. Es
    HTML de una sola pieza (no columnas + st.metric) porque no hay forma segura de "abrir" un div
    oscuro con un st.markdown y "cerrarlo" varios st.* después -- cada items[i] es un dict con
    'label', 'valor' (ya formateado) y opcionalmente 'delta' (texto) + 'favorable' (True/False/None
    para pintar el delta verde/rojo; None lo deja neutro).
    Adenda 25 (a pedido del equipo: "no quiero ver solamente la variación vs el promedio, que quede
    la variación con ronda anterior también, en todos los casos"): cada item admite además 'delta2'
    + 'favorable2' para una SEGUNDA línea de delta -- por convención delta = vs. Ronda Anterior,
    delta2 = vs. Promedio Industria (o al revés, según arme el caller; esta función no asume cuál es
    cuál, solo las apila en el orden en que llegan)."""
    piezas = []
    for it in items:
        delta_html = ''
        for clave_delta, clave_fav in (('delta', 'favorable'), ('delta2', 'favorable2')):
            if not it.get(clave_delta):
                continue
            # OJO: "favorable" suele venir de comparar floats de pandas/numpy (ej. delta > 0), que da
            # numpy.bool_ -- "is False" falla por identidad contra ese tipo aunque el valor sea
            # correcto (numpy.bool_(False) is False → False). Sin comparación de identidad.
            fav = it.get(clave_fav)
            clase = '' if fav is None else ('up' if fav else 'down')
            delta_html += f'<div class="kpi-band-delta {clase}">{_icono_delta(it[clave_delta])}{it[clave_delta]}</div>'
        piezas.append(f'<div class="kpi-band-item"><div class="kpi-band-label">{it["label"]}</div>'
                       f'<div class="kpi-band-value">{it["valor"]}</div>{delta_html}</div>')
    # En modo oscuro nativo de Streamlit la banda necesita distinguirse por color, no por
    # oscuridad -- ver el comentario junto a ".kpi-band-oscura.tema-oscuro" en style.css.
    clase_tema = ' tema-oscuro' if es_modo_oscuro() else ''
    st.markdown(f'<div class="kpi-band-oscura{clase_tema}">{"".join(piezas)}</div>', unsafe_allow_html=True)

def kpi_doble_delta(col, label, valor_txt, d_ronda=None, ronda_prev_nombre=None, favorable_ronda=None,
                     d_industria=None, favorable_industria=None, unidad='%'):
    """Adenda 25 (a pedido del equipo -- reemplaza st.metric(..., delta=...) en las tarjetas de KPI
    'claras' que antes mostraban un solo delta, 'X% vs Prom'): tarjeta con DOS deltas apilados, vs.
    Ronda Anterior y vs. Promedio Industria, siempre los dos que haya dato. st.metric no soporta dos
    deltas nativos, así que se arma en HTML -- el bloque '.kpi-card-doble' en style.css calca el
    estilo de las cards st.metric nativas (mismo borde lila, fondo y sombra) para que conviva sin
    desentonar al lado de tarjetas que todavía usan st.metric con un solo valor (ej. las que no
    tienen delta, como 'Calificación crediticia').
    - d_ronda / d_industria: delta YA calculado en %, como float (ej. de delta_vs_ronda_anterior() y
      delta_raw() respectivamente) -- None si no hay dato.
    - favorable_ronda / favorable_industria: True/False para pintar ESE delta verde/rojo; si se deja
      en None se asume signo positivo = favorable (sirve para la mayoría de las métricas; para
      'menos es mejor' -- Deuda, costos -- hay que pasarlo explícito, igual que ya hace el resto
      de la app con delta_color='inverse' / 'favorable' invertido).
    - ronda_prev_nombre: nombre de la ronda anterior (ronda_anterior_de(...)) para el texto de esa
      línea ('vs Ronda 1'); si no se pasa, queda 'vs ronda anterior' genérico.
    El ícono (🔺/🔻) sigue el signo LITERAL del delta (subió/bajó); el color sigue 'favorable' --
    mismo criterio que ya usa el resto del tablero (ver _icono_delta / kpi_banda_oscura)."""
    def _linea(valor, etiqueta, favorable):
        if valor is None or pd.isna(valor):
            return ''
        icono = '🔺' if valor > 0 else ('🔻' if valor < 0 else '➖')
        fav = favorable if favorable is not None else (valor > 0)
        clase = 'up' if fav else 'down'
        return f'<div class="kpi-card-delta {clase}">{icono} {valor:+.1f}{unidad} {etiqueta}</div>'
    etiqueta_ronda = f'vs {ronda_prev_nombre}' if ronda_prev_nombre else 'vs ronda anterior'
    html = (f'<div class="kpi-card-doble"><div class="kpi-card-label">{label}</div>'
            f'<div class="kpi-card-value">{valor_txt}</div>'
            f'{_linea(d_industria, "vs Industria", favorable_industria)}'
            f'{_linea(d_ronda, etiqueta_ronda, favorable_ronda)}'
            f'</div>')
    (col if col is not None else st).markdown(html, unsafe_allow_html=True)

# ---------------- Panel de alertas ----------------
def evaluar_alertas():
    """Corre un set de reglas sobre la ronda en foco y devuelve solo lo que se está prendiendo.
    La idea es no tener que recorrer las 5 secciones para enterarse de que algo se rompió."""
    alertas = []
    bal = df[(df['Estado'] == 'Hoja de Balance, miles USD, Global') & (df['Ronda'] == ronda_snapshot)]
    ratios = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Ronda'] == ronda_snapshot)]

    # --- Liquidez: el sobregiro automático de Cesim es un hecho del reporte, no un umbral nuestro
    deuda_cp = valor_de(bal, 'Deudas a corto plazo (no planificadas)', empresa_analisis)
    if deuda_cp and deuda_cp > 0:
        alertas.append(('critico', 'Sobregiro automático',
                        f'{format_num(deuda_cp)} USD de deuda de corto plazo NO planificada: la caja no alcanzó '
                        'para cubrir obligaciones. Suele venir con tasa de interés penal.'))

    orden_hoy = df[df['Ronda'] == ronda_snapshot]['Ronda_Orden'].iloc[0] if not df[df['Ronda'] == ronda_snapshot].empty else None

    # NOTA: hubo acá una alerta "I+D fuera de lo común" (Categoría 3, supuesto propio de CADIZ) que
    # vigilaba gasto de I+D récord + por encima del promedio de rivales, como indicio temprano de
    # tecnología nueva en camino. Se sacó a pedido del equipo (ensuciaba el panel y no aportaba
    # suficiente valor accionable) -- la detección CONFIRMADA de tecnología nueva (Categoría 2, un
    # hecho, no una estimación) sigue más abajo en "Entrada a tecnología nueva de la competencia".

    # --- Tendencias: dos rondas seguidas en la misma dirección. No hay número inventado acá,
    #     es la propia serie del reporte la que define si viene cayendo o subiendo.
    def dos_rondas_seguidas(estado, metrica, etiqueta, detalle, subiendo=False, seccion=None):
        serie = serie_metrica(estado, metrica, empresa_analisis, hasta_orden=orden_hoy, seccion=seccion)
        serie = [v for v in serie if v is not None and not pd.isna(v)]
        if len(serie) < 3: return
        d1, d2 = serie[-1] - serie[-2], serie[-2] - serie[-3]
        if (d1 > 0 and d2 > 0) if subiendo else (d1 < 0 and d2 < 0):
            alertas.append(('aviso', etiqueta, detalle.format(a=serie[-3], b=serie[-1])))

    dos_rondas_seguidas('Ratios e indicadores financieros clave', 'Retorno total acumulado del accionista (p.a.), %',
                        'Retorno del accionista en baja', 'Cayó dos rondas seguidas: de {a:,.1f}% a {b:,.1f}%.')
    dos_rondas_seguidas('Hoja de Balance, miles USD, Global', 'Inventario',
                        'Inventario acumulándose', 'Creció dos rondas seguidas — capital inmovilizado.', subiendo=True)

    # --- Cambios vs. la ronda anterior: se compara el dato de una ronda contra el de la otra,
    #     sin definir qué valor es "bueno" o "malo".
    ordenes = sorted(df['Ronda_Orden'].dropna().unique())
    orden_prev = None
    if orden_hoy in ordenes:
        i = ordenes.index(orden_hoy)
        orden_prev = ordenes[i - 1] if i > 0 else None

    if orden_prev is not None:
        ronda_prev = df[df['Ronda_Orden'] == orden_prev]['Ronda'].iloc[0]
        ratios_prev = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Ronda_Orden'] == orden_prev)]

        # Posición en el ranking de retorno del accionista
        def puesto(tabla):
            vals = {e: valor_de(tabla, 'Retorno total acumulado del accionista (p.a.), %', e) for e in COMPANIES}
            rk = sorted([e for e in vals if vals.get(e) is not None], key=vals.get, reverse=True)
            return rk.index(empresa_analisis) + 1 if empresa_analisis in rk else None
        p_hoy, p_ant = puesto(ratios), puesto(ratios_prev)
        if p_hoy is not None and p_ant is not None and p_hoy != p_ant:
            nivel = 'aviso' if p_hoy > p_ant else 'ok'
            verbo = 'Bajó' if p_hoy > p_ant else 'Subió'
            alertas.append((nivel, f'{verbo} del {p_ant}° al {p_hoy}° puesto',
                            f'Ranking de retorno del accionista, contra {ronda_prev}.'))

        # Calificación crediticia: se reporta el cambio, sin juzgar qué letra es aceptable
        calif_hoy = valor_texto(ratios, 'Calificación crediticia', empresa_analisis)
        calif_ant = valor_texto(ratios_prev, 'Calificación crediticia', empresa_analisis)
        if calif_hoy and calif_ant and str(calif_hoy).strip() != str(calif_ant).strip():
            alertas.append(('aviso', 'Cambió la calificación crediticia',
                            f'De {str(calif_ant).strip()} a {str(calif_hoy).strip()} respecto de {ronda_prev}.'))

        # --- Entrada CONFIRMADA de LA COMPETENCIA a una tecnología nueva: a diferencia del aviso de
        #     I+D de arriba (estimativo, y ciego a la vía de licencia), esto es un HECHO -- Categoría
        #     2, dato real de CESIM, sin ninguna estimación de nuestra parte. El manual describe DOS
        #     caminos para sumar tecnología (I+D propio o comprar una licencia, disponible de
        #     inmediato) y esta alerta los cubre a los dos por igual: no le importa CÓMO la consiguió
        #     el rival, solo que efectivamente ya la tiene y la está vendiendo -- verificado con la
        #     cuota de mercado real que CESIM publica por (país, tecnología) para los 7 equipos
        #     (`Informe de mercado, {país} → Seccion='{país} cuotas de mercado, %' → Metrica=
        #     tecnología`, el mismo campo que ya usa `cuota_mercado_objetivo_vs_real` en
        #     gap_analysis.py). Se dispara cuando un rival pasa de 0% en la ronda anterior a >0% en
        #     esta, en una tecnología/país donde antes no vendía nada -- vigila a todos los rivales
        #     siempre, igual que las otras alertas de competencia.
        #
        #     Aparece recién cuando ya hubo ventas reales (una ronda más tarde que la inversión en
        #     I+D si fue por esa vía, o la misma ronda si fue por licencia) -- llega después que el
        #     aviso de I+D, pero sin ninguna ambigüedad sobre qué tecnología es.
        #
        #     Se consolida en UN solo aviso (mismo patrón que "Movimientos de capacidad de la
        #     competencia" más abajo) en vez de una tarjeta por cada combinación: cuando una
        #     tecnología recién se habilita para toda la industria, pueden entrar 3-4 equipos juntos
        #     en la misma ronda, y una tarjeta idéntica repetida 4 veces satura el panel sin agregar
        #     información nueva en cada una.
        entradas_tech = []
        for pais in _MERCADOS_GAP:
            cuota_hoy = df[(df['Estado'] == f'Informe de mercado, {pais}') &
                           (df['Seccion'] == f'{pais} cuotas de mercado, %') &
                           (df['Ronda_Orden'] == orden_hoy) & (df['Empresa'] != MY_COMPANY)].copy()
            cuota_prev = df[(df['Estado'] == f'Informe de mercado, {pais}') &
                             (df['Seccion'] == f'{pais} cuotas de mercado, %') &
                             (df['Ronda_Orden'] == orden_prev) & (df['Empresa'] != MY_COMPANY)].copy()
            if cuota_hoy.empty: continue
            cuota_hoy['Valor'] = num(cuota_hoy['Valor'])
            cuota_prev['Valor'] = num(cuota_prev['Valor'])
            for tech in _TECNOLOGIAS_GAP:
                hoy_t = cuota_hoy[cuota_hoy['Metrica'] == tech].dropna(subset=['Valor']).set_index('Empresa')['Valor']
                prev_t = cuota_prev[cuota_prev['Metrica'] == tech].dropna(subset=['Valor']).set_index('Empresa')['Valor']
                for rival in [e for e in COMPANIES if e != MY_COMPANY]:
                    v_hoy = hoy_t.get(rival)
                    if v_hoy is None or v_hoy <= 0: continue
                    v_prev = prev_t.get(rival, 0.0)
                    if v_prev == 0:
                        entradas_tech.append(f'{rival} en {tech} ({pais}, 0% → {v_hoy:.1f}%)')
        if entradas_tech:
            detalle = (' · '.join(entradas_tech) + '. Antes no vendían nada ahí en esa tecnología — '
                       'confirmado con dato real, sin importar si la consiguieron con I+D propio o '
                       'comprando una licencia.')
            alertas.append(('aviso', 'Entrada a tecnología nueva de la competencia', detalle))

    # --- Movimientos de capacidad de LA COMPETENCIA: Cesim publica las fábricas que va a haber
    #     después de la próxima ronda, así que se sabe de antemano quién está por agrandarse o
    #     achicarse. La idea de esta alerta es específicamente vigilar a los rivales -- CADIZ ya
    #     sabe sus propias decisiones -- y hacerlo siempre, sin depender de a quién tengamos
    #     seleccionado en "Equipo en foco" (por eso se compara contra MY_COMPANY, no contra
    #     empresa_analisis: cambiar el equipo en foco para mirar otra sección no debe apagar esta
    #     vigilancia de la competencia).
    # OJO con la forma del reporte: Cesim pone el PAÍS en 'Metrica' (EE.UU. / China) y el
    # HORIZONTE en 'Subgrupo' (Ronda actual / Próxima ronda / Después de la próxima ronda).
    # Filtrando al revés no matchea nada y el conteo actual daba 0.
    #
    # Historial de este chequeo: la v1 mezclaba los 7 equipos en un solo mensaje ambiguo y solo
    # miraba subas (nunca bajas, como la reducción real de CEOS en EE.UU. 7->5 en Ronda 1) -- se
    # corrigió de más filtrando a empresa_analisis, lo cual apagaba el aviso de competencia cuando
    # el foco está en CADIZ, que es exactamente el caso de uso principal. Esta versión vuelve a
    # cubrir a todos los rivales (todo el que no sea MY_COMPANY), agrega reducciones, y desglosa
    # por área (Metrica) en vez de sumar EE.UU.+China en un solo número, que puede esconder un
    # movimiento real en un área si la otra se mueve al revés.
    #
    # Se mantiene el aviso de Ronda de práctica: verificado con datos reales que una decisión
    # cargada ahí (ensayo, ej. CADIZ China 2->3 en Práctica 2) puede no trasladarse nunca a la
    # competencia oficial (Ronda 1 mostró a CADIZ sin cambios) -- no es un compromiso real.
    #
    # FIX (reportado: "en Ronda 1 dice que CEOS expande China en las próximas 2 rondas, y en
    # Ronda 2 vuelve a decir lo mismo"): es el MISMO evento real, visto dos veces porque antes solo
    # se comparaba 'Ronda actual' contra 'Después de la próxima ronda' (fijo a 2 rondas vista) y el
    # título decía siempre "(próximas 2 rondas)" sin importar cuánto faltaba en realidad. Verificado
    # con los datos: en R1, CEOS China = 2 (actual) / 2 (próxima) / 3 (después) -> el cambio todavía
    # está a 2 rondas. En R2, CEOS China = 2 (actual) / 3 (próxima) / 3 (después) -> el MISMO cambio
    # ya está a 1 ronda -- el horizonte se acorta, pero el mensaje no lo reflejaba y parecía una
    # alerta repetida sin sentido. Ahora se mira primero 'Próxima ronda' (1 ronda vista); si ahí no
    # hay cambio, recién se usa 'Después de la próxima ronda' (2 rondas vista) -- así el aviso
    # cuenta regresiva (2 rondas -> 1 ronda) en vez de repetirse idéntico, y deja de aparecer solo
    # cuando el cambio ya se concretó en 'Ronda actual'.
    fab = df[(df['Estado'] == 'Detalles de fabricación') & (df['Seccion'] == 'Número de fábricas') &
             (df['Ronda'] == ronda_snapshot) & (df['Empresa'] != MY_COMPANY)].copy()
    if not fab.empty:
        fab['Valor'] = num(fab['Valor'])
        act = fab[fab['Subgrupo'] == 'Ronda actual'].groupby(['Empresa', 'Metrica'])['Valor'].sum()
        prox = fab[fab['Subgrupo'] == 'Próxima ronda'].groupby(['Empresa', 'Metrica'])['Valor'].sum()
        desp = fab[fab['Subgrupo'] == 'Después de la próxima ronda'].groupby(['Empresa', 'Metrica'])['Valor'].sum()
        movimientos = []
        for clave in act.index:
            empresa, area = clave
            a = act.get(clave, 0)
            p = prox.get(clave, a)
            d = desp.get(clave, p)
            # Prioridad: el cambio más cercano en el tiempo es el que importa mostrar ahora.
            if p != a:
                destino, horizonte = p, 'la próxima ronda'
            elif d != a:
                destino, horizonte = d, 'dentro de 2 rondas'
            else:
                continue
            verbo = 'expande' if destino > a else 'reduce'
            movimientos.append(f'{empresa} {verbo} {area} ({a:.0f} → {destino:.0f}, {horizonte})')
        if movimientos:
            detalle = ' · '.join(movimientos)
            alertas.append(('aviso', 'Movimientos de capacidad de la competencia', detalle))
    return alertas

def panel_alertas():
    alertas = evaluar_alertas()
    if not alertas:
        st.markdown('<div class="stat-segment-card"><div class="stat-segment-bar"><div class="seg ok" '
                    'style="width:100%"></div></div><div class="stat-segment-legend"><span class="ok">'
                    '<span class="pt"></span>Sin alertas activas</span></div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="alerta-fila ok">✅ <b>Sin alertas</b> '
                    f'<span class="detalle">— nada fuera de rango en {ronda_snapshot} para {empresa_analisis}.</span></div>',
                    unsafe_allow_html=True)
        return
    # Barra de composición por severidad -- el total es la cantidad de alertas ACTIVAS ahora
    # (no un universo fijo de "reglas evaluadas": una sola regla puede dispararse varias veces,
    # ver el comentario en assets/style.css junto a .stat-segment-card).
    n_critico = sum(1 for a in alertas if a[0] == 'critico')
    n_aviso = sum(1 for a in alertas if a[0] == 'aviso')
    n_ok = sum(1 for a in alertas if a[0] == 'ok')
    total_sev = n_critico + n_aviso + n_ok
    segs = ''.join(f'<div class="seg {niv}" style="width:{cnt/total_sev*100:.1f}%"></div>'
                   for niv, cnt in [('critico', n_critico), ('aviso', n_aviso), ('ok', n_ok)] if cnt)
    partes_legend = []
    if n_critico: partes_legend.append(f'<span class="critico"><span class="pt"></span>{n_critico} crítica{"s" if n_critico != 1 else ""}</span>')
    if n_aviso: partes_legend.append(f'<span class="aviso"><span class="pt"></span>{n_aviso} aviso{"s" if n_aviso != 1 else ""}</span>')
    if n_ok: partes_legend.append(f'<span class="ok"><span class="pt"></span>{n_ok} mejora{"s" if n_ok != 1 else ""}</span>')
    st.markdown(f'<div class="stat-segment-card"><div class="stat-segment-label">'
                f'Estado de alertas — {total_sev} activa{"s" if total_sev != 1 else ""} en {ronda_snapshot}</div>'
                f'<div class="stat-segment-bar">{segs}</div>'
                f'<div class="stat-segment-legend">{"".join(partes_legend)}</div></div>', unsafe_allow_html=True)
    # Orden: primero lo que hay que resolver, después lo informativo, al final las mejoras.
    orden = {'critico': 0, 'aviso': 1, 'ok': 2}
    iconos = {'critico': '🔴', 'aviso': '🟠', 'ok': '🟢'}
    clases = {'critico': 'alerta-fila', 'aviso': 'alerta-fila aviso', 'ok': 'alerta-fila ok'}
    for nivel, titulo, detalle in sorted(alertas, key=lambda a: orden.get(a[0], 1)):
        icono = iconos.get(nivel, '🟠')
        clase = clases.get(nivel, 'alerta-fila aviso')
        st.markdown(f'<div class="{clase}">{icono} <b>{titulo}</b> <span class="detalle">— {detalle}</span></div>',
                    unsafe_allow_html=True)

# ---------------- Sidebar ----------------
st.sidebar.markdown('### CÁDIZ AUTOMOTIVE')
# Las rondas de Práctica (el ensayo previo al arranque de la competencia oficial) se sacaron de la
# navegación a pedido del equipo -- ya arrancó la competencia Oficial (Ronda 0 y Ronda 1 jugadas) y
# esos ensayos "no suman" al análisis de gestión. Quedan como dato histórico en el repo
# (data/raw/practicas/*.xls) por si hace falta revisarlos, pero la app ya no los ofrece para elegir.
filtro_tipo = 'Oficial'
rondas_timeline = [f'Ronda {i}' for i in range(1, 13)]
ronda_snapshot = st.sidebar.select_slider('Ronda de análisis', options=rondas_timeline, value=rondas_timeline[0], key='slider_rondas')
# Indicador de estado (estilo "● Conectado" de la referencia) -- acá confirma de un vistazo qué
# ronda está en foco, en vez de ser puramente decorativo.
st.sidebar.markdown(f'<div class="sidebar-status"><span class="dot"></span>{ronda_snapshot}</div>',
                     unsafe_allow_html=True)
empresa_analisis = st.sidebar.selectbox('Equipo en foco', COMPANIES, index=0, key='select_equipo')
st.sidebar.divider()

# ---------------- Generar Excel de gestión (motor de proyección, build_gestion_v2.py) ----------------
# Ronda N.N (Fase 4): ya NO se suben archivos a mano -- todo se resuelve leyendo rutas fijas DEL PROPIO
# REPO de GitHub (mismas que usa el resto de la app / el modelo):
#   - data/raw/practicas/oficial/ : RDOS oficiales de CESIM ya commiteados (nombre de archivo libre --
#     descubrir_rdos_oficiales() detecta la ronda por el nombre O, si no matchea, abriendo el propio
#     RDOS y leyendo el título 'Ronda N' de la hoja Results). La última ronda con RDOS ahí define la
#     frontera REAL; la ronda a decidir es automáticamente N+1.
#   - data/decisiones/       : Excels de trabajo con las proyecciones/decisiones de CADIZ, nombrados
#     Cadiz_proyeccion_R{N}.xlsx. El botón toma el más reciente relevante (misma ronda en curso, o la
#     ronda que acaba de pasar a REAL), rescata ahí las decisiones futuras ya cargadas y congela el
#     plan de la ronda que acaba de cerrarse, todo dentro de generar_excel_desde_repo().
# El archivo nuevo se arma en memoria (io.BytesIO, nunca se escribe en el disco del servidor) y se
# ofrece para descargar con su nombre auto-detectado (Cadiz_proyeccion_R{N+1}.xlsx) -- el usuario lo
# commitea a mano en data/decisiones/, mismo flujo manual que ya usa hoy para CADIZ_Gestion_v2.xlsx.
with st.sidebar.expander('🔧 Generar Excel de gestión (próxima ronda)'):
    st.caption('Lee los RDOS oficiales de `data/raw/practicas/oficial/` y las decisiones previas de '
               '`data/decisiones/` -- ya commiteados en el repo, no hace falta subir nada acá.')
    if st.button('Generar Excel de la próxima ronda', key='btn_generar_excel'):
        try:
            with st.spinner('Generando Excel...'):
                buf, nombre_salida, info = generar_excel_desde_repo()
            rondas_reales_txt = ', '.join(f'R{r}' for r in info['rondas_reales'])
            st.success(f'{nombre_salida} generado -- rondas reales detectadas: {rondas_reales_txt} · '
                       f'ronda a decidir: R{info["ronda_a_decidir"]}.')
            detalle = []
            if info.get('archivo_previo_usado'):
                detalle.append(f'Decisiones previas rescatadas de `{os.path.basename(info["archivo_previo_usado"])}` '
                                f'({info["overrides_aplicados"]} valores aplicados'
                                + (f', {info["overrides_sin_match"]} sin match' if info.get('overrides_sin_match') else '')
                                + ').')
            else:
                detalle.append('No había un archivo de decisiones previo -- primera corrida (bootstrap).')
            if info.get('plan_congelado_filas'):
                detalle.append(f'Plan Congelado: {info["plan_congelado_filas"]} filas inyectadas en DATA_EXPORT.')
            if not info.get('ok', True):
                detalle.append(f'⚠️ No se pudo recalcular con LibreOffice ({info.get("mensaje")}) -- el '
                                'Excel se generó igual, pero Excel recalculará las fórmulas recién al abrirlo.')
            for d in detalle:
                st.caption(d)
            st.download_button(f'⬇️ Descargar {nombre_salida}', data=buf,
                                file_name=nombre_salida,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                key='dl_excel_generado')
        except FileNotFoundError as e:
            st.error(f'No se pudo generar el Excel: {e}')
        except Exception as e:
            st.error(f'No se pudo generar el Excel: {e}')
st.sidebar.divider()
# BUG REPORTADO Y CONFIRMADO -- es una limitación de la plataforma, no de este código: Streamlit
# expone el tema elegido (Settings > claro/oscuro/uso del sistema) vía st.context.theme.type, pero
# esa lectura se toma UNA VEZ por corrida del script, y alternar el tema en el menú de Settings no
# siempre dispara una corrida nueva por sí solo. Resultado: si tocás el toggle de tema, los colores
# que dependen de es_modo_oscuro() (banda de KPIs, sidebar, gráficos) pueden quedar con el tema
# VIEJO hasta que algo más fuerce un rerun -- cambiar de Ronda, de Equipo o de Sección, como ya
# venía notando el equipo. Este botón es el atajo más chico y confiable para lo mismo, sin tener
# que tocar otro control de la app.
if st.sidebar.button('🔄 Actualizar tema', help='Usalo si cambiaste entre modo claro/oscuro en Settings y los colores de la app no se actualizaron solos.'):
    st.rerun()
SECCIONES = ['Resultados', 'Mercado', 'Operaciones', 'Finanzas', 'RRHH y Sostenibilidad']
seccion = st.sidebar.radio('Sección', SECCIONES, key='select_seccion_router')
df_all = get_data(filtro_tipo)
if df_all.empty or ronda_snapshot not in df_all['Ronda'].unique():
    # CASO ESPECIAL: CESIM todavía no publicó ningún RDOS de esta ronda (nada que mostrar en
    # Resultados/Mercado/Operaciones/RRHH), pero CADIZ ya cargó su propia proyección en el modelo de
    # gestión para esa ronda -- en vez de un callejón sin salida, se muestra directamente el Control
    # de Gestión (lo único que SÍ hay para ofrecer) en la sección Finanzas. Fuera de Finanzas, o si
    # no hay proyección tampoco, se mantiene el aviso original sin cambios.
    ronda_num_bypass = ronda_a_num(ronda_snapshot)
    df_proy_bypass = get_proyeccion() if ronda_num_bypass is not None else None
    hay_proyeccion_cadiz = (df_proy_bypass is not None and not df_proy_bypass[
        (df_proy_bypass['round'] == ronda_num_bypass) & (df_proy_bypass['team'] == MY_COMPANY)].empty)
    _CROSSWALK_POR_SECCION_BYPASS = {'Finanzas': CROSSWALK_FINANZAS, 'Mercado': CROSSWALK_MERCADO,
                                      'Operaciones': CROSSWALK_OPERACIONES, 'Resultados': CROSSWALK_RESULTADOS}
    if hay_proyeccion_cadiz and seccion in _CROSSWALK_POR_SECCION_BYPASS and empresa_analisis == MY_COMPANY:
        st.info(f"📁 CESIM todavía no publicó los RDOS de **{ronda_snapshot}** — el resto del tablero "
                "no tiene datos para mostrar todavía, pero CADIZ ya cargó su proyección para esta "
                "ronda en el modelo de gestión:")
        panel_comparativa_plan_real(df_all.copy(), ronda_snapshot, crosswalk=_CROSSWALK_POR_SECCION_BYPASS[seccion],
                               key_suffix=f'{seccion.lower()}_sin_real', mostrar_directo=True)
    else:
        st.info(f"📁 Faltan datos: No se encontraron archivos para **{ronda_snapshot}** en el entorno **{filtro_tipo}**.")
    st.stop()
df = df_all.copy()
ronda_ultima = ronda_snapshot
try:
    st.markdown(f'<style>{open(os.path.join(BASE_DIR, "assets", "style.css"), encoding="utf-8").read()}</style>', unsafe_allow_html=True)
except FileNotFoundError:
    # Antes fallaba en silencio (pass): si el archivo no está en el despliegue, la app se veía
    # sin ningún estilo custom (sidebar claro, sin banda oscura, sin barra segmentada) y no había
    # ninguna pista de por qué. Ahora al menos avisa.
    st.warning('No se encontró assets/style.css — el estilo visual (sidebar oscuro, banda de KPIs, '
               'barras segmentadas) no se aplicó en este despliegue. Verificá que la carpeta '
               '"assets" se haya subido junto con app.py.')
# =================================================================
# SECCIÓN 1 — RESULTADOS
# =================================================================
def seccion_resultado():
    tab_resumen, tab_cg = st.tabs(['Resumen', 'Comparativa Plan vs. Real'])
    with tab_resumen:
        _seccion_resultado_resumen()
    with tab_cg:
        if empresa_analisis == MY_COMPANY:
            panel_comparativa_plan_real(df_all.copy(), ronda_snapshot, crosswalk=CROSSWALK_RESULTADOS, key_suffix='resultados', mostrar_directo=True)
            st.divider()
            fila3_resultados_ingresos(df_all.copy(), ronda_snapshot, ronda_a_num(ronda_snapshot), get_proyeccion())
        else:
            st.caption('Cambiá "Equipo en foco" a CADIZ en la barra lateral para ver la Comparativa Plan vs. Real (es sobre la proyección propia de CADIZ).')
def _seccion_resultado_resumen():
    val_ronda = df[(df['Estado'] == 'Valuación - Global') & (df['Ronda'] == ronda_snapshot)]
    ratios_ronda_r1 = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Ronda'] == ronda_snapshot)]

    # "Accionistas, Total" en la tabla de Creación de Valor = Beneficio de la ronda: es el valor
    # generado PARA EL ACCIONISTA en esa ronda puntual (no confundir con "Valor total creado", que
    # suma también lo pagado a Proveedores/Personal/Gobierno — no es plata del accionista).
    acc_ronda = df[(df['Modulo'] == 'Creación de valor') & (df['Seccion'] == 'Accionistas') &
                   (df['Metrica'] == 'Total') & (df['Ronda'] == ronda_snapshot)]
    cv_vals = {emp: valor_de(acc_ronda, 'Total', emp) for emp in COMPANIES}

    # Acumulado: Cesim ya reporta este campo pre-acumulado desde el inicio del juego — no hay que
    # sumarlo nosotros ronda a ronda. Es el criterio real de "creación de valor para el accionista".
    retorno_acum_vals = {emp: valor_de(ratios_ronda_r1, 'Retorno total acumulado del accionista (p.a.), %', emp) for emp in COMPANIES}

    # Retorno de ESTA ronda puntual: variación simple del precio de la acción vs. la ronda anterior.
    # (El campo "acumulado" es *per annum*, no es aditivo entre rondas — restarlo directo da un
    # número que no representa lo que pasó en la ronda. Esto sí es directamente comparable ronda a ronda.)
    precio_hist = df[(df['Estado'] == 'Ratios e indicadores financieros clave') &
                      (df['Metrica'] == 'Precio de la acción al final de la ronda, USD')].copy()
    precio_hist['Valor'] = num(precio_hist['Valor'])
    ordenes_disp = sorted(df['Ronda_Orden'].dropna().unique())
    orden_actual = df[df['Ronda'] == ronda_snapshot]['Ronda_Orden'].iloc[0] if not df[df['Ronda'] == ronda_snapshot].empty else None
    idx_orden = ordenes_disp.index(orden_actual) if orden_actual in ordenes_disp else None
    orden_anterior = ordenes_disp[idx_orden - 1] if idx_orden and idx_orden > 0 else None
    retorno_ronda_vals = {}
    if orden_anterior is not None:
        for emp in COMPANIES:
            p_act = valor_de(precio_hist[precio_hist['Ronda_Orden'] == orden_actual], 'Precio de la acción al final de la ronda, USD', emp)
            p_ant = valor_de(precio_hist[precio_hist['Ronda_Orden'] == orden_anterior], 'Precio de la acción al final de la ronda, USD', emp)
            retorno_ronda_vals[emp] = ((p_act - p_ant) / p_ant * 100) if (p_act is not None and p_ant not in (None, 0)) else None

    cap_vals = {emp: valor_de(val_ronda, 'Capitalización de mercado, miles USD', emp) for emp in COMPANIES}
    with st.container(border=True):
        st.markdown(f'**Alertas — {empresa_analisis}, {ronda_snapshot}**')
        panel_alertas()
    st.write('')
    st.subheader('KPIs de Valor')
    prom_ret_acum = np.nanmean([v for v in retorno_acum_vals.values() if v is not None]) if any(v is not None for v in retorno_acum_vals.values()) else None
    val_ret_acum = retorno_acum_vals.get(empresa_analisis)
    delta_ret_acum = ((val_ret_acum - prom_ret_acum) / abs(prom_ret_acum) * 100) if prom_ret_acum and val_ret_acum is not None else None

    ranking_acum = sorted([e for e in retorno_acum_vals if retorno_acum_vals.get(e) is not None], key=retorno_acum_vals.get, reverse=True)
    pos = ranking_acum.index(empresa_analisis) + 1 if empresa_analisis in ranking_acum else '-'
    ret_hist = df_all[(df_all['Estado'] == 'Ratios e indicadores financieros clave') &
                       (df_all['Metrica'] == 'Retorno total acumulado del accionista (p.a.), %')].copy()
    ret_hist['Valor'] = num(ret_hist['Valor'])
    puestos_hist = ret_hist.dropna(subset=['Valor']).copy()
    puestos_hist['Puesto'] = puestos_hist.groupby('Ronda')['Valor'].rank(ascending=False, method='min')
    serie_puesto = puestos_hist[puestos_hist['Empresa'] == empresa_analisis].sort_values('Ronda_Orden')['Puesto'].tolist()

    prom_cv = np.nanmean([v for v in cv_vals.values() if v is not None]) if any(v is not None for v in cv_vals.values()) else None
    val_cv = cv_vals.get(empresa_analisis)
    delta_cv = ((val_cv - prom_cv)/prom_cv*100) if prom_cv and val_cv is not None else None

    prom_cap = np.nanmean([v for v in cap_vals.values() if v is not None]) if any(v is not None for v in cap_vals.values()) else None
    val_cap = cap_vals.get(empresa_analisis)
    delta_cap = ((val_cap - prom_cap)/prom_cap*100) if prom_cap and val_cap else None

    # Adenda 25 (a pedido del equipo: "no quiero ver solamente la variación vs el promedio, que
    # quede la variación con ronda anterior también, en todos los casos"): cada KPI de la banda
    # ahora trae TAMBIÉN su delta vs. la ronda anterior, además del vs. Promedio Industria que ya
    # tenía. "Beneficio del accionista" no tiene un Estado/Metrica directo en `df` (sale de
    # Modulo == 'Creación de valor', no de un Estado) -- por eso se arma a mano en vez de con
    # delta_vs_ronda_anterior(), que solo sabe filtrar por Estado.
    d_ret_acum_ronda, ronda_prev_kpi = delta_vs_ronda_anterior('Ratios e indicadores financieros clave', 'Retorno total acumulado del accionista (p.a.), %')
    d_cap_ronda, _ = delta_vs_ronda_anterior('Valuación - Global', 'Capitalización de mercado, miles USD')
    d_cv_ronda = None
    if ronda_prev_kpi:
        acc_prev = df[(df['Modulo'] == 'Creación de valor') & (df['Seccion'] == 'Accionistas') &
                      (df['Metrica'] == 'Total') & (df['Ronda'] == ronda_prev_kpi)]
        val_cv_prev = valor_de(acc_prev, 'Total', empresa_analisis)
        if val_cv_prev not in (None, 0) and val_cv is not None and pd.notna(val_cv_prev) and pd.notna(val_cv):
            d_cv_ronda = (val_cv - val_cv_prev) / abs(val_cv_prev) * 100

    def _fmt_delta(d, sufijo):
        return f'{d:+.1f}% {sufijo}' if d is not None else None

    sufijo_ronda = f'vs {ronda_prev_kpi}' if ronda_prev_kpi else 'vs ronda anterior'
    # Los 3 números que más le importan a un accionista van en la banda oscura, separados del
    # resto (ver kpi_banda_oscura) -- "Retorno acum. del accionista", mismo término base que usa
    # el crosswalk de Resultados; "Capitalización de mercado" (antes "Market Cap", en inglés) usa
    # el mismo nombre que el propio campo de CESIM.
    kpi_banda_oscura([
        {'label': 'Retorno acum. del accionista', 'valor': f'{val_ret_acum:,.1f}%' if val_ret_acum is not None else '—',
         'delta': _fmt_delta(d_ret_acum_ronda, sufijo_ronda), 'favorable': (d_ret_acum_ronda > 0) if d_ret_acum_ronda is not None else None,
         'delta2': _fmt_delta(delta_ret_acum, 'vs Industria'), 'favorable2': (delta_ret_acum > 0) if delta_ret_acum is not None else None},
        {'label': 'Beneficio del accionista', 'valor': format_num(val_cv),
         'delta': _fmt_delta(d_cv_ronda, sufijo_ronda), 'favorable': (d_cv_ronda > 0) if d_cv_ronda is not None else None,
         'delta2': _fmt_delta(delta_cv, 'vs Industria'), 'favorable2': (delta_cv > 0) if delta_cv is not None else None},
        {'label': 'Capitalización de mercado (USD)', 'valor': format_num(val_cap),
         'delta': _fmt_delta(d_cap_ronda, sufijo_ronda), 'favorable': (d_cap_ronda > 0) if d_cap_ronda is not None else None,
         'delta2': _fmt_delta(delta_cap, 'vs Industria'), 'favorable2': (delta_cap > 0) if delta_cap else None},
    ])

    # Posición en el ranking y Retorno de la ronda son más de contexto que de "número que decide
    # la creación de valor" -- se quedan como tarjetas claras con sparkline, más chicas.
    c2, c5 = st.columns(2)
    # color=COLOR_MAP.get(empresa_analisis): antes el sparkline caía al default de la función, que
    # (Adenda 25) dejó de asumir el lila/rojo de marca -- acá SÍ corresponde el color de identidad del
    # equipo en foco (COLOR_MAP ya resuelve lila para CADIZ y el tono muted que le toque a cualquier
    # otro equipo), así que se pasa explícito.
    # invertir: en el ranking, "para arriba" en el gráfico tiene que ser mejorar de puesto
    # OJO -- excepción deliberada al "doble delta en todos los casos": "Posición en el ranking" es
    # un ORDINAL (1°, 2°...), no un monto ni un %, así que un "+X% vs Industria/Ronda Anterior" no
    # tiene una lectura clara acá (¿+50% de qué, de puesto?) -- se deja con su sparkline (que ya
    # muestra la tendencia del puesto) y sin delta numérico, en vez de forzar un % engañoso.
    kpi_con_tendencia(c2, 'Posición en el ranking', f'{pos}° de {len(COMPANIES)}', serie_puesto,
                       color=COLOR_MAP.get(empresa_analisis), invertir=True)

    prom_ret_ronda = np.nanmean([v for v in retorno_ronda_vals.values() if v is not None]) if any(v is not None for v in retorno_ronda_vals.values()) else None
    val_ret_ronda = retorno_ronda_vals.get(empresa_analisis)
    delta_ret_ronda = ((val_ret_ronda - prom_ret_ronda) / abs(prom_ret_ronda) * 100) if prom_ret_ronda and val_ret_ronda is not None else None
    # OJO -- segunda excepción deliberada: "Retorno de la acción" YA ES, por definición, la variación
    # de precio de ESTA ronda vs. la ronda anterior (ver comentario más arriba) -- agregarle un
    # segundo delta "vs. Ronda Anterior" sería comparar un cambio contra sí mismo, confuso. Se deja
    # con un solo delta (vs. Industria), que sí es una comparación genuinamente distinta.
    kpi_con_tendencia(c5, 'Retorno de la acción',
                       f'{val_ret_ronda:+,.1f}%' if val_ret_ronda is not None else '—',
                       serie_metrica('Ratios e indicadores financieros clave', 'Precio de la acción al final de la ronda, USD', empresa_analisis),
                       delta=f'{delta_ret_ronda:+.1f}% vs Industria' if delta_ret_ronda is not None else None,
                       color=COLOR_MAP.get(empresa_analisis))
    st.caption('El retorno acumulado es per-annum y no es aditivo entre rondas — para ver cómo fue *esta* ronda usá '
               '"Retorno de la acción", que es variación simple de precio. En la primera ronda del ecosistema no hay '
               'ronda previa con la cual compararla. "Retorno de la acción" ya es en sí una comparación vs. ronda '
               'anterior, por eso trae un solo delta (vs. Industria) y no dos.')
    st.divider()
    vista_ranking = st.radio('Vista del ranking (columna derecha)', ['Acumulado (Retorno del Accionista, %)', f'Solo {ronda_snapshot} (USD)'], horizontal=True, key='vista_ranking_cv')
    col_a, col_b = st.columns(2)
    with col_a:
        pl = df[(df['Estado'] == 'Cuenta de resultados, miles USD, Global') & (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot)]
        def g(metrica): return valor_de(pl, metrica) or 0.0
        ingresos = g('Ingresos por ventas')
        if ingresos > 0:
            costos_prod = g('Costos de fabricación interna') + g('Costos de la característica') + g('Costos de fabricación contratada')
            costos_op = g('Costos de transporte y aranceles') + g('I+D') + g('Promoción') + g('Administración')
            depr, ints, imp, ben = g('Depreciación de Activos Fijos'), g('Gastos financieros netos'), g('Impuesto sobre el beneficio'), g('Beneficio de la ronda')
            _valores_puente = [ingresos, -costos_prod, -costos_op, -depr, -ints, -imp, ben]
            fig = go.Figure(go.Waterfall(
                orientation='v', measure=['absolute', 'relative', 'relative', 'relative', 'relative', 'relative', 'total'],
                x=['Ingresos', '- Prod', '- Op/Admin', '- Depr', '- Int', '- Imp', '= Neto'],
                y=_valores_puente,
                # Adenda 25: % entre paréntesis sobre Ingresos (primera barra, el ancla del puente).
                text=_texto_cascada(_valores_puente), textposition='outside',
                # Adenda 25 (a pedido del equipo): convención universal -- verde = suma (increasing),
                # rojo = resta (decreasing), en todos los waterfalls de la app. Antes decreasing usaba
                # MUTED_PALETTE[2] (un verde grisáceo apagado, no rojo).
                decreasing={'marker': {'color': COLOR_NEGATIVE}}, increasing={'marker': {'color': COLOR_POSITIVE}},
                totals={'marker': {'color': COLOR_POSITIVE if ben > 0 else COLOR_NEGATIVE}}
            ))
            fig.update_layout(title=f'Puente de Beneficio Neto — {empresa_analisis}')
            mostrar(fig, ocultar_eje_valores='y')
    with col_b:
        es_acumulado = vista_ranking.startswith('Acumulado')
        datos_ranking = retorno_acum_vals if es_acumulado else cv_vals
        titulo_ranking = 'Ranking: Retorno Acumulado del Accionista, %' if es_acumulado else f'Ranking: Beneficio del Accionista — {ronda_snapshot}'
        ranking = pd.DataFrame(list(datos_ranking.items()), columns=['Empresa', 'Valor']).sort_values('Valor', ascending=True).dropna()
        ranking['Etiqueta'] = ranking['Valor'].apply(lambda v: f'{v:,.1f}%') if es_acumulado else ranking['Valor'].apply(format_num)
        fig = px.bar(ranking, x='Valor', y='Empresa', orientation='h', color='Empresa', color_discrete_map=COLOR_MAP, text='Etiqueta')
        fig.update_traces(textposition='outside', cliponaxis=False, showlegend=False)
        fig.update_layout(title=titulo_ranking, xaxis=dict(range=[0, ranking['Valor'].max() * 1.25]))
        mostrar(fig, ocultar_eje_valores='x')
    st.divider()
    st.subheader('Cuota de mercado')
    # Antes: 5 gráficos de barra casi idénticos (global, por valor, EE.UU., China, Europa) -- mismo
    # tipo de gráfico repetido 5 veces es exactamente la queja de "muchos gráficos de barra". Acá se
    # reemplazan por DOS formas distintas, cada una elegida por el trabajo que hace mejor:
    #   1) Dumbbell (dos puntos + línea): compara cuota por UNIDADES vs. cuota por VALOR por equipo en
    #      un solo gráfico -- la distancia entre los dos puntos de cada equipo ES el dato interesante
    #      (vender una porción de unidades distinta a la de ingresos implica un mix de precio propio).
    #   2) Heatmap: cuota por país (EE.UU./China/Europa) x equipo en una sola grilla -- reemplaza 3
    #      barras por 1 sola lectura de matriz, con la tabla de datos completa abajo por accesibilidad.
    sub_global = df[(df['Estado'] == 'Informe de mercado, global') &
                     (df['Seccion'] == 'Cuotas de mercado globales, %') & (df['Metrica'] == 'Total') &
                     (df['Ronda'] == ronda_snapshot)].copy()
    sub_global['Valor'] = num(sub_global['Valor'])
    cuota_unidades = sub_global.set_index('Empresa')['Valor'].to_dict()
    # Cuota por valor ($): CESIM no publica esta cifra directamente (los precios están en moneda local
    # por mercado -- USD/RMB/EUR -- y sumarlos sin convertir daría un número sin sentido). Se calcula
    # sobre "Ingresos por ventas, Global" (Cuenta de resultados), que CESIM SÍ publica ya convertido a
    # USD para los 7 equipos: cuota por valor = ingresos de la empresa / Σ ingresos de los 7 equipos.
    ingresos_glob = df[(df['Estado'] == 'Cuenta de resultados, miles USD, Global') &
                        (df['Metrica'] == 'Ingresos por ventas') & (df['Ronda'] == ronda_snapshot)].copy()
    ingresos_glob['Valor'] = num(ingresos_glob['Valor'])
    total_ing = ingresos_glob['Valor'].sum()
    cuota_valor = (ingresos_glob.set_index('Empresa')['Valor'] / total_ing * 100).to_dict() if total_ing else {}
    filas_dumbbell = [{'Empresa': e, 'Unidades': cuota_unidades.get(e), 'Valor': cuota_valor.get(e)} for e in COMPANIES]
    dfd = pd.DataFrame(filas_dumbbell).dropna(subset=['Unidades', 'Valor'])
    if not dfd.empty:
        dfd = dfd.sort_values('Valor')
        fig_d = go.Figure()
        for _, r in dfd.iterrows():
            es_cadiz = r['Empresa'] == MY_COMPANY
            fig_d.add_trace(go.Scatter(x=[r['Unidades'], r['Valor']], y=[r['Empresa'], r['Empresa']], mode='lines',
                                        line=dict(color=COLOR_CADIZ if es_cadiz else 'rgba(140,151,166,0.45)',
                                                  width=2.5 if es_cadiz else 1.5),
                                        showlegend=False, hoverinfo='skip'))
        # OJO: estos dos marcadores son por TIPO de medida (unidades vs. valor), para los 7 equipos a
        # la vez -- no son identidad de CADIZ (esa ya se marca en la línea de arriba, por es_cadiz),
        # así que van con tonos muted genéricos, no con el lila de marca.
        fig_d.add_trace(go.Scatter(x=dfd['Unidades'], y=dfd['Empresa'], mode='markers', name='Cuota por unidades, %',
                                    marker=dict(size=11, color=MUTED_PALETTE[0]),
                                    hovertemplate='%{y} — Unidades: %{x:.1f}%<extra></extra>'))
        # Adenda 25: se saca el marrón (MUTED_PALETTE[1]) de acá -- no pasa a lila porque seguiría sin
        # ser identidad (ver comentario de arriba), y el lila se reserva para "esto es CADIZ".
        # Fix "ya casi": MUTED_PALETTE[5] quedaba casi idéntico al MUTED_PALETTE[0] del marcador de
        # arriba (mismo gráfico) -- se reemplaza por el azul acero nuevo, bien separado del otro.
        fig_d.add_trace(go.Scatter(x=dfd['Valor'], y=dfd['Empresa'], mode='markers', name='Cuota por valor ($), %',
                                    marker=dict(size=11, color='#58759D', symbol='diamond'),
                                    hovertemplate='%{y} — Valor: %{x:.1f}%<extra></extra>'))
        fig_d.update_layout(title=f'Cuota de mercado global — unidades vs. valor, {ronda_snapshot}', xaxis_title='%',
                             legend=dict(orientation='h', yanchor='top', y=-0.2, xanchor='center', x=0.5))
        mostrar(fig_d)
    else:
        st.info('Sin datos suficientes de cuota global (unidades y valor) para esta ronda.')
    st.caption('Cada línea conecta la cuota de UN equipo en dos monedas de medida: unidades vendidas y valor en '
               'USD. Cuando el punto de Valor queda a la derecha del de Unidades, ese equipo vende más caro que '
               'el promedio (o con mejor mix); si queda a la izquierda, más barato.')
    paises = ['EE.UU.', 'China', 'Europa']
    filas_hm = []
    for pais in paises:
        sub_pais = df[(df['Estado'] == f'Informe de mercado, {pais}') &
                      (df['Seccion'] == f'{pais} cuotas de mercado, %') & (df['Metrica'] == 'Total') &
                      (df['Ronda'] == ronda_snapshot)].copy()
        sub_pais['Valor'] = num(sub_pais['Valor'])
        for _, row in sub_pais.dropna(subset=['Valor']).iterrows():
            filas_hm.append({'Empresa': row['Empresa'], 'País': pais, 'Cuota': row['Valor']})
    if filas_hm:
        dfh = pd.DataFrame(filas_hm)
        piv = dfh.pivot(index='Empresa', columns='País', values='Cuota').reindex(columns=paises)
        piv = piv.reindex(piv.mean(axis=1).sort_values(ascending=False).index)  # ranking, no alfabético
        fig_hm = go.Figure(go.Heatmap(
            z=piv.values, x=list(piv.columns), y=list(piv.index),
            # Secuencial de un solo matiz, sin sentido de identidad ni sentimiento (esto muestra a los
            # 7 equipos, no solo a CADIZ) -- se usa el azul de COLOR_METRICA['dinero'] en vez del lila/
            # rojo de marca (Adenda 25).
            colorscale=[[0, _hex_a_rgba(COLOR_METRICA['dinero'], 0.06)], [1, COLOR_METRICA['dinero']]],
            text=[[f'{v:.1f}%' if pd.notna(v) else '' for v in fila] for fila in piv.values],
            texttemplate='%{text}', textfont=dict(size=12),
            hovertemplate='%{y} — %{x}: %{z:.1f}%<extra></extra>', showscale=False,
            xgap=3, ygap=3))  # gap entre celdas -- separador de superficie, no una grilla de líneas
        fig_hm.update_layout(title=f'Cuota de mercado por país, % — {ronda_snapshot}')
        mostrar(fig_hm)
        with st.expander('Ver como tabla'):
            st.dataframe(piv.style.format('{:.1f}%', na_rep='—'), use_container_width=True)
    else:
        st.info('Sin datos de cuota de mercado por país para esta ronda.')
    st.divider()
    cap_sub = df[(df['Estado'] == 'Valuación - Global') & (df['Metrica'] == 'Capitalización de mercado, miles USD')].copy()
    # chart_evolucion() ya antepone "Evolución — " al titulo -- pasarle "Evolución de la ..." de nuevo
    # duplicaba la palabra ("Evolución — Evolución de..."). Solo el sustantivo acá.
    chart_evolucion(cap_sub, 'Capitalización de Mercado (USD)')
# =================================================================
# SECCIÓN 2 — MERCADO
# =================================================================
def seccion_mercado():
    tab_pos, tab_pan, tab_evo, tab_cg = st.tabs(['Posicionamiento', 'Panorama Competitivo', 'Evolución', 'Comparativa Plan vs. Real'])
    tecnologias = ['Combustión', 'Híbrido', 'Eléctrico', 'Hidrógeno']

    with tab_cg:
        if empresa_analisis == MY_COMPANY:
            panel_comparativa_plan_real(df_all.copy(), ronda_snapshot, crosswalk=CROSSWALK_MERCADO, key_suffix='mercado', mostrar_directo=True)
            st.divider()
            fila3_mercado_cuota_objetivo(df_all.copy(), ronda_snapshot, ronda_a_num(ronda_snapshot), get_proyeccion())
        else:
            st.caption('Cambiá "Equipo en foco" a CADIZ en la barra lateral para ver la Comparativa Plan vs. Real (es sobre la proyección propia de CADIZ).')

    with tab_pos:
        c1, c2 = st.columns(2)
        pais_sel = c1.selectbox('Mercado', ['EE.UU.', 'China', 'Europa'], key='sel_mercado_pais')
        tech_sel = c2.selectbox('Tecnología', tecnologias, key='sel_mercado_tech')
        estado_pais = f'Informe de mercado, {pais_sel}'
        sub = df[(df['Estado'] == estado_pais) & (df['Seccion'] == tech_sel) & (df['Ronda'] == ronda_snapshot)]

        ventas_data = [{'Empresa': emp, 'Volumen': valor_fuzzy(sub[sub['Empresa'] == emp], 'Ventas')} for emp in COMPANIES]
        vol_df = pd.DataFrame(ventas_data).dropna()

        # Adenda 22 (a pedido del equipo): antes había DOS scatter separados (Precio vs Volumen,
        # Características vs Volumen) -- se reemplazan por UNA sola Matriz de Valor con Precio y
        # Características en los dos ejes y el Volumen como tamaño de burbuja, para ver la
        # propuesta de valor completa de cada equipo de un vistazo.
        st.subheader('Matriz de Valor')
        st.caption('Precio vs. Características, tamaño de burbuja = Volumen de ventas. Arriba a la derecha = propuesta '
                   'de valor alta (mucho producto, precio alto); abajo a la izquierda = jugador de precio bajo / pocas '
                   'características. Sirve para ver en qué cuadrante de la propuesta de valor compite cada equipo.')
        precio_df = sub[sub['Metrica'].str.contains(r'^Precio', case=False, na=False)][['Empresa', 'Valor']].rename(columns={'Valor': 'Precio'})
        carac_df = sub[sub['Metrica'].str.contains(r'^Cantidad de características', case=False, na=False)][['Empresa', 'Valor']].rename(columns={'Valor': 'Características'})
        if not vol_df.empty and not precio_df.empty and not carac_df.empty:
            mv = precio_df.merge(carac_df, on='Empresa').merge(vol_df, on='Empresa').dropna()
            mv['Precio'] = num(mv['Precio']); mv['Características'] = num(mv['Características'])
            mv = mv[mv['Volumen'] > 0]
            if not mv.empty:
                fig_mv = px.scatter(mv, x='Precio', y='Características', size='Volumen', color='Empresa',
                                     color_discrete_map=COLOR_MAP, text='Empresa',
                                     title=f'Matriz de Valor — {tech_sel}, {pais_sel}, {ronda_snapshot}')
                fig_mv.update_traces(textposition='top center', showlegend=False)
                linea_media(fig_mv, mv['Precio'].mean(), eje='x', etiqueta='Precio Prom')
                linea_media(fig_mv, mv['Características'].mean(), eje='y', etiqueta='Caract. Prom')
                mostrar(fig_mv)
            else:
                st.info('Sin volumen de ventas positivo para graficar en esta combinación.')
        else:
            st.info('Sin datos suficientes de Precio/Características/Volumen para esta combinación.')

        st.divider()
        # Adenda 22 (a pedido del equipo): reemplaza el scatter "Eficiencia Comercial" (Marketing
        # vs. Volumen) por la matriz "Trampa del Volumen" (Margen de Contribución Unitario vs.
        # Volume Market Share), pensada para detectar equipos que ganan participación sacrificando
        # rentabilidad por unidad.
        st.subheader('"Trampa del Volumen" — Margen vs. Participación')
        st.caption('Eje X = Volume Market Share (% de las unidades vendidas de esta tecnología en este mercado que son de '
                   'cada equipo, dato que publica CESIM). Eje Y = Margen de Contribución Unitario (USD/unidad) = Margen de '
                   'contribución total de la tecnología / unidades vendidas. Abajo a la derecha = mucho volumen pero poco '
                   'margen por unidad — la "trampa del volumen". Arriba a la izquierda = poco volumen pero alto margen unitario.')
        share_tv = df[(df['Estado'] == estado_pais) & (df['Seccion'] == f'{pais_sel} cuotas de mercado, %') &
                      (df['Metrica'] == tech_sel) & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Valor']].rename(columns={'Valor': 'Share'}).copy()
        share_tv['Share'] = num(share_tv['Share'])
        margen_tot_tv = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {pais_sel}') & (df['Seccion'] == tech_sel) &
                           (df['Metrica'] == 'Margen de contribución') & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Valor']].rename(columns={'Valor': 'MargenTotal'}).copy()
        margen_tot_tv['MargenTotal'] = num(margen_tot_tv['MargenTotal'])
        ventas_tv = df[(df['Estado'] == estado_pais) & (df['Seccion'] == tech_sel) &
                       (df['Metrica'] == 'Ventas, miles unidades') & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Valor']].rename(columns={'Valor': 'VentasMiles'}).copy()
        ventas_tv['VentasMiles'] = num(ventas_tv['VentasMiles'])
        tv = share_tv.merge(margen_tot_tv, on='Empresa').merge(ventas_tv, on='Empresa').dropna()
        tv = tv[tv['VentasMiles'] > 0]
        if not tv.empty:
            tv['MargenUnitario'] = tv['MargenTotal'] / tv['VentasMiles']  # miles USD / miles u. = USD/u.
            fig_tv = px.scatter(tv, x='Share', y='MargenUnitario', color='Empresa', color_discrete_map=COLOR_MAP,
                                 text='Empresa', title=f'Margen Unitario vs. Share — {tech_sel}, {pais_sel}, {ronda_snapshot}')
            fig_tv.update_traces(textposition='top center', showlegend=False, marker=dict(size=14))
            fig_tv.update_layout(xaxis_title='Volume Market Share, %', yaxis_title='Margen de Contribución Unitario, USD/u.')
            linea_media(fig_tv, tv['Share'].mean(), eje='x', etiqueta='Share Prom')
            linea_media(fig_tv, tv['MargenUnitario'].mean(), eje='y', etiqueta='Margen Prom')
            mostrar(fig_tv)
        else:
            st.info('Sin datos suficientes de Share/Margen/Ventas para esta combinación.')

        st.divider()
        st.subheader('Share of Voice (SOV) vs. Share of Market (SOM)')
        st.caption(f'{tech_sel}, {pais_sel}, {ronda_snapshot}. SOV = Promoción del equipo / Promoción total de la '
                   'tecnología (los 7 equipos). SOM = cuota de mercado real que publica CESIM para esa tecnología '
                   '(Ventas del equipo / Ventas totales DE ESA TECNOLOGÍA, no del mercado completo) — mismo grano '
                   'que el SOV, para que la comparación sea de manzanas con manzanas.')
        promo_sov = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {pais_sel}') & (df['Seccion'] == tech_sel) &
                       (df['Metrica'] == 'Promoción') & (df['Ronda'] == ronda_snapshot)].copy()
        promo_sov['Valor'] = num(promo_sov['Valor']).abs()
        som_sov = df[(df['Estado'] == f'Informe de mercado, {pais_sel}') & (df['Seccion'] == f'{pais_sel} cuotas de mercado, %') &
                     (df['Metrica'] == tech_sel) & (df['Ronda'] == ronda_snapshot)].copy()
        som_sov['Valor'] = num(som_sov['Valor'])
        promo_total_sov = promo_sov['Valor'].sum()
        if promo_total_sov > 0 and not som_sov.empty:
            sov_df = promo_sov[['Empresa', 'Valor']].rename(columns={'Valor': 'Promo'})
            sov_df['SOV'] = sov_df['Promo'] / promo_total_sov * 100
            som_df = som_sov[['Empresa', 'Valor']].rename(columns={'Valor': 'SOM'})
            comp_sov = sov_df.merge(som_df, on='Empresa', how='outer')
            comp_sov_long = comp_sov.melt(id_vars='Empresa', value_vars=['SOV', 'SOM'], var_name='Indicador', value_name='Pct').dropna(subset=['Pct'])
            # Adenda 25 (pedido explícito del equipo, citado como ejemplo): se saca el marrón de acá.
            fig_sov = px.bar(comp_sov_long, x='Empresa', y='Pct', color='Indicador', barmode='group',
                              color_discrete_map={'SOV': MUTED_PALETTE[0], 'SOM': COLOR_CADIZ},
                              text=comp_sov_long['Pct'].apply(lambda v: f'{v:,.1f}%'),
                              title=f'SOV vs. SOM — {tech_sel}, {pais_sel}, {ronda_snapshot}')
            fig_sov.update_traces(textposition='outside', cliponaxis=False)
            fig_sov.update_layout(yaxis_title='%')
            mostrar(fig_sov)
        else:
            st.info('Sin datos suficientes de Promoción o Cuota de mercado para esta combinación.')

    with tab_pan:
        c1p, c2p = st.columns(2)
        pais_pan = c1p.selectbox('Mercado', ['EE.UU.', 'China', 'Europa'], key='sel_pan_pais')
        tech_pan = c2p.selectbox('Tecnología', tecnologias, key='sel_pan_tech')
        estado_pan = f'Informe de mercado, {pais_pan}'

        st.markdown('**Enfoque de estrategia de marketing — los 7 equipos**')
        est = df[(df['Estado'] == estado_pan) & (df['Seccion'] == tech_pan) &
                 (df['Metrica'] == 'Enfoque de la estrategia de marketing') & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Valor']]
        if not est.empty:
            estrategias = ['Precio bajo', 'Equilibrado', 'Marca', 'Características', 'Precio alto']
            # Mapa de color FIJO por estrategia (antes usaba el color automático de Plotly, que
            # reasigna colores según el orden de aparición en cada gráfico — por eso "Marca"
            # salía de un color acá y de otro allá).
            # 5 categorías genéricas (no hay identidad ni sentimiento acá) -- las 5 usan tonos muted.
            # Adenda 25: se saca el marrón (MUTED_PALETTE[1]) -- no pasa a lila para no sugerir que
            # una de las 5 estrategias "es CADIZ" (esto es la estrategia de LOS 7 equipos a la vez).
            # Fix "ya casi": también se saca MUTED_PALETTE[5] (casi idéntico a [0]) y MUTED_PALETTE[4]
            # (kaki, a solo ~15 de distancia RGB del marrón baneado) -- mismo set de 5 tonos separados
            # que MUTED_SIN_MARRON.
            estrategia_colores = dict(zip(estrategias, MUTED_SIN_MARRON))
            est = est.copy().sort_values('Empresa')
            # Antes era un gráfico de barras todas de la misma altura: el eje Y no codificaba nada
            # y el texto iba rotado adentro de la barra. Peor todavía, px.bar parte el dataframe en
            # un trace por estrategia y el update_traces le pasaba la columna ENTERA a cada trace,
            # así que el texto de la barra no correspondía a su color. Con chips no hay ambigüedad
            # posible: el color y el texto salen de la misma fila.
            with st.container(border=True):
                st.markdown(f'**Estrategia elegida — {tech_pan}, {pais_pan}, {ronda_snapshot}**')
                chips = []
                for _, fila in est.iterrows():
                    color = estrategia_colores.get(fila['Valor'], MUTED_PALETTE[0])
                    clase = 'chip-estrategia destacado' if fila['Empresa'] == MY_COMPANY else 'chip-estrategia'
                    chips.append(f'<span class="{clase}" style="background-color:{color}">'
                                 f'{fila["Empresa"]} · {fila["Valor"]}</span>')
                st.markdown(' '.join(chips), unsafe_allow_html=True)
        else:
            st.info('Sin datos de estrategia para esta combinación.')

        st.divider()
        st.markdown('**Mix tecnológico**')
        col_mix1, col_mix2 = st.columns(2)
        mix_rows = []
        for tech in tecnologias:
            v = df[(df['Estado'] == estado_pan) & (df['Seccion'] == tech) & (df['Metrica'] == 'Ventas, miles unidades') & (df['Ronda'] == ronda_snapshot)]['Valor']
            mix_rows.append({'Tecnología': tech, 'Ventas': num(v).sum()})
        mix_df = pd.DataFrame(mix_rows)
        mix_df = mix_df[mix_df['Ventas'] > 0]
        with col_mix1:
            if not mix_df.empty:
                # Adenda 25: MUTED_SIN_MARRON en vez de MUTED_PALETTE -- acá se detectó que el marrón
                # SÍ se colaba (Plotly lo asigna por orden de aparición, no por nombre de tecnología).
                fig_mix = px.pie(mix_df, names='Tecnología', values='Ventas', hole=0.5,
                                  color_discrete_sequence=MUTED_SIN_MARRON,  # 4 tecnologías, genérico
                                  title=f'Toda la industria — {pais_pan}, {ronda_snapshot}')
                mostrar(fig_mix)
            else:
                st.info('Sin ventas registradas en esta combinación.')
        with col_mix2:
            mix_emp_rows = []
            for tech in tecnologias:
                sub_tech = df[(df['Estado'] == estado_pan) & (df['Seccion'] == tech) & (df['Metrica'] == 'Ventas, miles unidades') & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Valor']]
                sub_tech['Valor'] = num(sub_tech['Valor'])
                for _, r in sub_tech.iterrows():
                    mix_emp_rows.append({'Empresa': r['Empresa'], 'Tecnología': tech, 'Ventas': r['Valor']})
            mix_emp_df = pd.DataFrame(mix_emp_rows).dropna(subset=['Ventas'])
            mix_emp_df = mix_emp_df[mix_emp_df['Ventas'] > 0]
            if not mix_emp_df.empty:
                totales = mix_emp_df.groupby('Empresa')['Ventas'].transform('sum')
                mix_emp_df['Pct'] = mix_emp_df['Ventas'] / totales * 100
                fig_mix_emp = px.bar(mix_emp_df, x='Empresa', y='Pct', color='Tecnología', barmode='stack',
                                      color_discrete_sequence=MUTED_SIN_MARRON,  # 4 tecnologías, genérico
                                      title=f'Por equipo — {pais_pan}, {ronda_snapshot}')
                fig_mix_emp.update_layout(yaxis_title='% de ventas')
                mostrar(fig_mix_emp)
            else:
                st.info('Sin ventas por equipo en esta combinación.')

    with tab_evo:
        pais_evo = st.selectbox('Mercado', ['EE.UU.', 'China', 'Europa'], key='sel_evo_pais')
        estado_evo = f'Informe de mercado, {pais_evo}'

        st.markdown('**Demanda total vs. ventas totales de la industria, por tecnología**')
        st.caption('Toda la torta del mercado (7 equipos sumados): la brecha entre demanda y ventas es oportunidad que nadie capturó. '
                   'Un panel por tecnología para ver dónde está esa brecha, en vez de un solo total que mezcla las cuatro — '
                   'se omite la tecnología que todavía no tiene ningún dato en este mercado.')
        paneles_tech = []
        for tech in tecnologias:
            dem = df[(df['Estado'] == estado_evo) & (df['Seccion'] == tech) & (df['Metrica'] == 'Demanda, miles unidades')].copy()
            ven = df[(df['Estado'] == estado_evo) & (df['Seccion'] == tech) & (df['Metrica'] == 'Ventas, miles unidades')].copy()
            dem['Valor'] = num(dem['Valor']); ven['Valor'] = num(ven['Valor'])
            dem_piv = dem.groupby(['Ronda', 'Ronda_Orden'], as_index=False)['Valor'].sum().assign(Tipo='Demanda')
            ven_piv = ven.groupby(['Ronda', 'Ronda_Orden'], as_index=False)['Valor'].sum().assign(Tipo='Ventas')
            piv = pd.concat([dem_piv, ven_piv], ignore_index=True)
            if piv.empty or piv['Valor'].fillna(0).abs().sum() == 0:
                continue  # tecnología sin ningún dato (>0) en este mercado -- no se muestra el panel
            paneles_tech.append((tech, piv.sort_values('Ronda_Orden')))
        if paneles_tech:
            cols_dv = st.columns(2)
            for i, (tech, piv) in enumerate(paneles_tech):
                fig_dv = go.Figure()
                # Adenda 25 (pedido explícito del equipo, override de la lectura anterior): "Ventas"
                # pasa a lila acá también, igual que en SOV vs. SOM -- aunque en rigor sea de los 7
                # equipos sumados y no identidad estricta de CADIZ, el equipo prefiere el lila como
                # color fijo de "Ventas" en los paneles de panorama competitivo.
                for tipo, color in [('Demanda', MUTED_PALETTE[0]), ('Ventas', COLOR_CADIZ)]:
                    d_t = piv[piv['Tipo'] == tipo]
                    fig_dv.add_trace(go.Bar(x=d_t['Ronda'], y=d_t['Valor'], name=tipo, marker_color=color))
                fig_dv.update_layout(barmode='group', title=f'{tech} — {pais_evo}')
                with cols_dv[i % 2]:
                    mostrar(fig_dv)
        else:
            st.info('Sin datos suficientes.')

        st.divider()
        st.markdown(f'**Evolución de la cuota de mercado — {pais_evo}**')
        st.caption('Los 7 equipos, CADIZ resaltado.')
        share_hist = df[(df['Estado'] == estado_evo) & (df['Seccion'] == f'{pais_evo} cuotas de mercado, %') &
                         (df['Metrica'].str.strip() == 'Total')].copy()
        chart_evolucion(share_hist, f'Cuota de mercado, % — {pais_evo}')
        # NOTA: hasta acá había un gráfico "Trayectoria de {empresa}: precio y características en el
        # tiempo" (Precio, USD vs. Cantidad de características, cada uno en su propio eje Y superpuesto
        # en el mismo plano -- yaxis / yaxis2 con overlaying='y'). Era el único gráfico de doble eje Y
        # de verdad que quedaba en la app (el resto ya se había migrado a paneles apilados, ver
        # chart_dos_metricas_apiladas) -- señalado como pendiente en la Adenda 18 y sacado a pedido
        # explícito del equipo.
# =================================================================
# SECCIÓN 3 — OPERACIONES
# =================================================================
def seccion_operaciones():
    bloque1, bloque2, tab_cg = st.tabs(['Capacidad y Costos', 'Inventario y Logística', 'Comparativa Plan vs. Real'])
    with tab_cg:
        if empresa_analisis == MY_COMPANY:
            panel_comparativa_plan_real(df_all.copy(), ronda_snapshot, crosswalk=CROSSWALK_OPERACIONES, key_suffix='operaciones', mostrar_directo=True)
            st.divider()
            fila3_operaciones_gap_fabricacion(df_all.copy(), ronda_snapshot, ronda_a_num(ronda_snapshot), get_proyeccion())
        else:
            st.caption('Cambiá "Equipo en foco" a CADIZ en la barra lateral para ver la Comparativa Plan vs. Real (es sobre la proyección propia de CADIZ).')
    with bloque1:
        cap = df[(df['Estado'] == 'Detalles de fabricación') & (df['Seccion'] == 'Capacidad empleada, %') & (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot) & (df['Subgrupo'].isin(['EE.UU.', 'China']))].copy()
        cap['Valor'] = num(cap['Valor'])
        # Cesim reporta la capacidad usada POR TECNOLOGÍA dentro de cada país (ej. Combustión 45% +
        # Híbrido 40% en EE.UU. = 85% de la planta) — hay que sumarlas para tener el % real de la planta,
        # si no, cada tecnología aparecía como una gauge separada.
        cap = cap.dropna(subset=['Valor']).groupby('Subgrupo', as_index=False)['Valor'].sum()
        if not cap.empty:
            cols = st.columns(len(cap))
            for col, (_, row) in zip(cols, cap.iterrows()):
                titulo = row['Subgrupo'] if pd.notna(row['Subgrupo']) else 'Capacidad'
                uso = float(row['Valor'])
                libre = max(100 - uso, 0)
                # Barra segmentada, mismo lenguaje visual que "Estado de Alertas" (panel_alertas) --
                # pero acá SÍ hay un universo fijo real (100% = capacidad instalada de la planta),
                # a diferencia de las alertas, donde no existe un total fijo de "reglas evaluadas".
                # Sin semáforo: no hay un rango "sano" publicado por Cesim, así que poner umbrales
                # propios sería inventar un criterio que el reporte no da.
                with col:
                    st.markdown(
                        f'<div class="stat-segment-card">'
                        f'<div class="stat-segment-label">Capacidad empleada — {titulo}</div>'
                        f'<div class="stat-segment-value">{uso:,.1f}%</div>'
                        f'<div class="stat-segment-bar">'
                        f'<div class="seg uso" style="width:{min(uso, 100):.1f}%"></div>'
                        f'<div class="seg libre" style="width:{libre:.1f}%"></div></div>'
                        f'<div class="stat-segment-legend">'
                        f'<span class="uso"><span class="pt"></span>Usado</span>'
                        f'<span class="libre"><span class="pt"></span>Libre</span></div></div>',
                        unsafe_allow_html=True)
        st.divider()
        st.markdown('**Producción: propia vs. contratada, y fábricas**')
        prod = df[(df['Estado'] == 'Detalles de fabricación') &
                  (df['Seccion'].isin(['Producción interna, miles unidades', 'Producción contratada, miles unidades'])) &
                  (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot) &
                  (df['Subgrupo'].isin(['EE.UU.', 'China']))].copy()
        prod['Valor'] = num(prod['Valor'])
        prod = prod.dropna(subset=['Valor']).groupby(['Subgrupo', 'Seccion'], as_index=False)['Valor'].sum()
        prod['Tipo'] = prod['Seccion'].map({'Producción interna, miles unidades': 'Interna', 'Producción contratada, miles unidades': 'Contratada'})

        col_pp, col_ff = st.columns(2)
        with col_pp:
            if not prod.empty:
                # Adenda 25: se saca el marrón. No pasa a lila -- este gráfico es del EQUIPO EN FOCO
                # (empresa_analisis), que puede no ser CADIZ; "Interna" no es identidad de CADIZ.
                # Fix "ya casi": MUTED_PALETTE[5] era casi idéntico a [0] -- se reemplaza por el sage
                # ya usado en otros lados de la app (MUTED_PALETTE[2]), bien separado del azul-grisáceo.
                fig_prod = px.bar(prod, x='Subgrupo', y='Valor', color='Tipo', barmode='group',
                                   color_discrete_map={'Interna': MUTED_PALETTE[2], 'Contratada': MUTED_PALETTE[0]},
                                   text=prod['Valor'].apply(lambda v: f'{v:,.0f}'),
                                   title=f'Producción propia vs. contratada — {ronda_snapshot}')
                fig_prod.update_traces(textposition='outside', cliponaxis=False)
                mostrar(fig_prod, ocultar_eje_valores='y')
            else:
                st.info('Sin datos de producción para esta combinación.')
        with col_ff:
            fab_all = df[(df['Estado'] == 'Detalles de fabricación') & (df['Seccion'] == 'Número de fábricas') &
                         (df['Ronda'] == ronda_snapshot) & (df['Subgrupo'] == 'Ronda actual')].copy()
            fab_all['Valor'] = num(fab_all['Valor'])
            fab_piv = fab_all.groupby(['Empresa', 'Metrica'], as_index=False)['Valor'].sum()
            if not fab_piv.empty:
                # Adenda 25: se saca el marrón. No pasa a lila -- son los 7 equipos a la vez, "EE.UU."
                # es un mercado, no identidad de CADIZ.
                # Fix "ya casi": MUTED_PALETTE[5] era casi idéntico a [0] -- mismo reemplazo por sage.
                fig_fab = px.bar(fab_piv, x='Empresa', y='Valor', color='Metrica', barmode='stack',
                                  color_discrete_map={'EE.UU.': MUTED_PALETTE[2], 'China': MUTED_PALETTE[0]},
                                  text=fab_piv['Valor'].apply(lambda v: f'{v:,.0f}'),
                                  title=f'Fábricas por equipo — {ronda_snapshot}')
                fig_fab.update_traces(textposition='inside')
                mostrar(fig_fab, ocultar_eje_valores='y')
            else:
                st.info('Sin datos de fábricas para esta ronda.')

        # La alerta de expansión de capacidad se movió al panel de alertas de Resultados,
        # para no tener avisos importantes desparramados por el tablero.

        st.divider()
        pl = df[(df['Estado'] == 'Cuenta de resultados, miles USD, Global') & (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot)]
        def g(metrica): return valor_de(pl, metrica) or 0.0
        ingresos = g('Ingresos por ventas')
        if ingresos > 0:
            # Adenda 24 (a pedido del equipo): las etiquetas anteriores ("- Fab. Interna", "- Caract.",
            # etc.) nombraban el COSTO que se resta en cada escalón, pero el valor graficado en ESE
            # escalón es el MARGEN REMANENTE después de restarlo (ingresos acumulados menos todos los
            # costos restados hasta ahí) -- no el costo en sí. Eso es lo que confundía: la etiqueta y
            # el número no describían la misma cosa. Se renombra cada escalón para que diga el estado
            # del margen que efectivamente muestra la barra.
            # De paso, el diseño anterior tenía un escalón "- Op/Admin" y el escalón final "= EBITDA"
            # con el MISMO valor (el EBITDA se calculaba como el valor del escalón anterior, sin restar
            # nada más) -- dos barras idénticas seguidas, sin una razón que lo justifique. Se lo
            # corrige acá restando I+D/Promoción/Administración directamente en el cálculo del EBITDA,
            # sin un escalón intermedio duplicado.
            etapas = [('Ingresos por Ventas', ingresos)]
            etapas.append(('Margen tras Fab. Interna', etapas[-1][1] - g('Costos de fabricación interna')))
            etapas.append(('Margen tras Características', etapas[-1][1] - g('Costos de la característica')))
            etapas.append(('Margen tras Fab. Contratada', etapas[-1][1] - g('Costos de fabricación contratada')))
            etapas.append(('Margen tras Logística', etapas[-1][1] - g('Costos de transporte y aranceles')))
            ebitda = etapas[-1][1] - g('I+D') - g('Promoción') - g('Administración')
            etapas.append(('= EBITDA', ebitda))
            colores = [COLOR_POSITIVE] + [MUTED_PALETTE[2]]*(len(etapas)-2) + [COLOR_POSITIVE if ebitda > 0 else COLOR_NEGATIVE]
            # textinfo='value+...' usa el formateo automático de Plotly, que muestra decimales sin
            # redondear ("45.91583M") -- inconsistente con format_num() (1 decimal) que se usa en el
            # resto del tablero, incluida la waterfall de "Puente de Beneficio Neto" con estos mismos
            # datos unas filas más abajo. Se arma el texto a mano con format_num().
            fig = go.Figure(go.Funnel(y=[e[0] for e in etapas], x=[e[1] for e in etapas],
                                       text=[format_num(e[1]) for e in etapas], textinfo='text+percent initial',
                                       marker={'color': colores}))
            fig.update_layout(title='Estructura Macro de Costos (Funnel) — margen remanente en cada etapa')
            mostrar(fig, ocultar_eje_valores='x')

        st.divider()
        st.markdown('**Costos de proveedores — composición**')
        st.caption('No incluye Fabricación contratada — eso ya se ve arriba, en Producción propia vs. contratada.')
        prov = df[(df['Modulo'] == 'Creación de valor') & (df['Seccion'] == 'Proveedores') &
                  (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot) &
                  (~df['Metrica'].isin(['Total', 'Valor total creado', 'Costos de fabricación contratada']))].copy()
        prov['Valor'] = num(prov['Valor'])
        prov = prov.dropna(subset=['Valor'])
        if not prov.empty:
            # Adenda 25: MUTED_SIN_MARRON -- mismo motivo que Mix tecnológico (Plotly asigna el color
            # por orden de aparición; con 2+ categorías el marrón podía tocarle a cualquiera).
            fig_prov = px.bar(prov, x='Metrica', y='Valor', color='Metrica', color_discrete_sequence=MUTED_SIN_MARRON,
                               text=prov['Valor'].apply(format_num), title=f'Composición de costos de proveedores — {ronda_snapshot}')
            fig_prov.update_traces(textposition='outside', cliponaxis=False, showlegend=False)
            mostrar(fig_prov, ocultar_eje_valores='y')
        else:
            st.info('Sin datos de proveedores para esta combinación.')
        st.divider()
        c3, c4 = st.columns(2)
        pais_ue = c3.selectbox('País — Margen Unitario', ['EE.UU.', 'China', 'Europa'], key='sel_op_pais')
        tech_ue = c4.selectbox('Tecnología — Margen Unitario', ['Combustión', 'Híbrido', 'Eléctrico', 'Hidrógeno'], key='sel_op_tech')
        margen = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {pais_ue}') & (df['Seccion'] == tech_ue) & (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot)]
        mercado = df[(df['Estado'] == f'Informe de mercado, {pais_ue}') & (df['Seccion'] == tech_ue) & (df['Ronda'] == ronda_snapshot)]
        unidades = valor_fuzzy(mercado, '^Ventas', empresa=empresa_analisis)
        def gm(metrica): return valor_de(margen, metrica) or 0.0
        if unidades and unidades > 0:
            p_venta = gm('Ingresos por ventas') / unidades
            c_prod = -gm('Fabricación propia y por contrato') / unidades
            c_flete = -gm('Transporte y aranceles') / unidades
            c_caract = -gm('Costos de la característica') / unidades
            m_bruto = gm('Beneficio bruto') / unidades
            _valores_margen_ue = [p_venta, c_prod, c_flete, c_caract, m_bruto]
            fig_ue = go.Figure(go.Waterfall(
                orientation='v', measure=['absolute', 'relative', 'relative', 'relative', 'total'],
                x=['Precio', '- Prod.', '- Logística', '- Caract.', '= Margen Unitario'],
                y=_valores_margen_ue,
                # Adenda 25: % entre paréntesis sobre el Precio (primera barra).
                text=_texto_cascada(_valores_margen_ue), textposition='outside',
                # Adenda 25: convención universal de color -- verde = suma, rojo = resta. Antes
                # decreasing usaba MUTED_PALETTE[2] (verde grisáceo apagado, no rojo).
                decreasing={'marker': {'color': COLOR_NEGATIVE}}, increasing={'marker': {'color': COLOR_POSITIVE}},
                totals={'marker': {'color': COLOR_POSITIVE if m_bruto > 0 else COLOR_NEGATIVE}}
            ))
            fig_ue.update_layout(title=f'Margen Unitario — {tech_ue}, {pais_ue}')
            mostrar(fig_ue, ocultar_eje_valores='y')
            costo_total_unit = -(c_prod + c_flete + c_caract)
            markup_pct = (m_bruto / costo_total_unit * 100) if costo_total_unit else None
            margen_pct = (m_bruto / p_venta * 100) if p_venta else None

            # Adenda 24 (a pedido del equipo): la tarjeta de Contribución Marginal Unitaria dejó de
            # mostrar el valor absoluto (USD/u.) -- ese desglose ya se ve arriba, en el Waterfall --
            # y pasa a mostrar solo el % (Contribución Marginal / Precio de Venta). Ambas tarjetas
            # (Margen % y Mark-up %) ahora traen como delta la variación en puntos porcentuales vs.
            # la ronda anterior, para el mismo país/tecnología elegidos, usando ronda_anterior_de().
            ronda_prev_ue = ronda_anterior_de(ronda_snapshot)
            margen_pct_prev = markup_pct_prev = None
            if ronda_prev_ue:
                margen_prev = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {pais_ue}') &
                                  (df['Seccion'] == tech_ue) & (df['Empresa'] == empresa_analisis) &
                                  (df['Ronda'] == ronda_prev_ue)]
                mercado_prev = df[(df['Estado'] == f'Informe de mercado, {pais_ue}') & (df['Seccion'] == tech_ue) &
                                   (df['Ronda'] == ronda_prev_ue)]
                unidades_prev = valor_fuzzy(mercado_prev, '^Ventas', empresa=empresa_analisis)
                if unidades_prev and unidades_prev > 0:
                    def gm_prev(metrica): return valor_de(margen_prev, metrica) or 0.0
                    p_venta_prev = gm_prev('Ingresos por ventas') / unidades_prev
                    m_bruto_prev = gm_prev('Beneficio bruto') / unidades_prev
                    costo_total_unit_prev = -(-gm_prev('Fabricación propia y por contrato') / unidades_prev
                                               - gm_prev('Transporte y aranceles') / unidades_prev
                                               - gm_prev('Costos de la característica') / unidades_prev)
                    margen_pct_prev = (m_bruto_prev / p_venta_prev * 100) if p_venta_prev else None
                    markup_pct_prev = (m_bruto_prev / costo_total_unit_prev * 100) if costo_total_unit_prev else None
            delta_margen_pct = (margen_pct - margen_pct_prev) if (margen_pct is not None and margen_pct_prev is not None) else None
            delta_markup_pct = (markup_pct - markup_pct_prev) if (markup_pct is not None and markup_pct_prev is not None) else None

            # Adenda 25 (a pedido del equipo: doble delta en todos los KPIs): además de vs. ronda
            # anterior, se suma vs. Promedio Industria -- para eso hay que recalcular Margen %/
            # Mark-up % de LOS 7 EQUIPOS en este mismo país/tecnología (antes solo se calculaba para
            # el equipo en foco). `mercado` ya no está filtrado por empresa, se reusa tal cual.
            margen_pct_todos, markup_pct_todos = {}, {}
            for _emp in COMPANIES:
                _margen_emp = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {pais_ue}') &
                                  (df['Seccion'] == tech_ue) & (df['Empresa'] == _emp) & (df['Ronda'] == ronda_snapshot)]
                _unid_emp = valor_fuzzy(mercado, '^Ventas', empresa=_emp)
                if _unid_emp and _unid_emp > 0:
                    def _gme(metrica): return valor_de(_margen_emp, metrica) or 0.0
                    _p_venta_e = _gme('Ingresos por ventas') / _unid_emp
                    _m_bruto_e = _gme('Beneficio bruto') / _unid_emp
                    _costo_e = -(-_gme('Fabricación propia y por contrato') / _unid_emp
                                 - _gme('Transporte y aranceles') / _unid_emp
                                 - _gme('Costos de la característica') / _unid_emp)
                    margen_pct_todos[_emp] = (_m_bruto_e / _p_venta_e * 100) if _p_venta_e else None
                    markup_pct_todos[_emp] = (_m_bruto_e / _costo_e * 100) if _costo_e else None
                else:
                    margen_pct_todos[_emp] = markup_pct_todos[_emp] = None

            def _delta_ind_pp(vals_dict, val_propio):
                # Delta en PUNTOS PORCENTUALES vs. promedio de industria (no % relativo) -- mismo
                # criterio que ya se usa acá para "vs. ronda anterior" (también en p.p., porque
                # margen_pct/markup_pct ya son porcentajes).
                _vals = [v for v in vals_dict.values() if v is not None]
                if not _vals or val_propio is None:
                    return None
                _prom = np.nanmean(_vals)
                return (val_propio - _prom) if pd.notna(_prom) else None

            d_margen_ind_pp = _delta_ind_pp(margen_pct_todos, margen_pct)
            d_markup_ind_pp = _delta_ind_pp(markup_pct_todos, markup_pct)

            col_cm, col_mk = st.columns(2)
            with col_cm:
                kpi_doble_delta(None, 'Margen Unitario (%)', f'{margen_pct:,.1f}%' if margen_pct is not None else '—',
                                 d_ronda=delta_margen_pct, ronda_prev_nombre=ronda_prev_ue,
                                 d_industria=d_margen_ind_pp, unidad=' p.p.')
            with col_mk:
                kpi_doble_delta(None, 'Mark-up aplicado', f'{markup_pct:,.1f}%' if markup_pct is not None else '—',
                                 d_ronda=delta_markup_pct, ronda_prev_nombre=ronda_prev_ue,
                                 d_industria=d_markup_ind_pp, unidad=' p.p.')
            st.caption('Margen % = Contribución Marginal Unitaria / Precio de Venta (el desglose en USD/u. está en el '
                       'Waterfall de arriba). Mark-up = Contribución Marginal Unitaria / Costo unitario total '
                       '(fabricación + logística + características) — mismo numerador, denominador distinto (precio '
                       'vs. costo), no confundir uno con otro. Deltas en puntos porcentuales (p.p.), no % relativo: '
                       'vs. la ronda inmediatamente anterior y vs. el promedio de los 7 equipos, mismo país/tecnología.')

        st.divider()
        st.markdown(f'**Punto de equilibrio — {empresa_analisis}, {ronda_snapshot}**')
        cm_total_miles = 0.0
        vol_total_miles = 0.0
        for mercado_be in ['EE.UU.', 'China', 'Europa']:
            margen_be = df[(df['Estado'] == f'Desglose de margen por tec, miles USD, {mercado_be}') &
                           (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot) &
                           (df['Metrica'] == 'Margen de contribución')].copy()
            cm_total_miles += num(margen_be['Valor']).sum()
            ventas_be = df[(df['Estado'] == f'Informe de mercado, {mercado_be}') & (df['Empresa'] == empresa_analisis) &
                          (df['Ronda'] == ronda_snapshot) & (df['Metrica'] == 'Ventas, miles unidades')].copy()
            vol_total_miles += num(ventas_be['Valor']).sum()
        pl_be = df[(df['Estado'] == 'Cuenta de resultados, miles USD, Global') & (df['Empresa'] == empresa_analisis) & (df['Ronda'] == ronda_snapshot)]
        def gbe(metrica): return valor_de(pl_be, metrica) or 0.0
        costos_fijos_miles = gbe('Depreciación de Activos Fijos') + gbe('I+D') + gbe('Administración')
        if vol_total_miles > 0 and cm_total_miles:
            cm_unitaria = cm_total_miles / vol_total_miles  # miles USD / miles u. = USD/u. -- la escala se cancela
            volumen_real = vol_total_miles * 1000.0
            volumen_equilibrio = (costos_fijos_miles * 1000.0) / cm_unitaria if cm_unitaria else None
            if volumen_equilibrio is not None and volumen_equilibrio > 0:
                chart_bullet(f'Punto de Equilibrio — {empresa_analisis}, {ronda_snapshot}', volumen_equilibrio, volumen_real,
                             'unidades', nombre_fondo='Volumen de Equilibrio', nombre_frente='Volumen Real Vendido',
                             color_excedente=COLOR_POSITIVE)  # vender por encima del equilibrio es favorable
                st.caption('Costos Fijos = Depreciación + I+D + Administración (real, Global). Contribución Marginal '
                           'Unitaria Ponderada = Margen de contribución total / unidades totales vendidas, sumado en '
                           'los 3 mercados donde el equipo vende.')
            else:
                st.info('No se pudo calcular el punto de equilibrio para esta combinación.')
        else:
            st.info('Sin datos suficientes de ventas/margen para calcular el punto de equilibrio.')
    with bloque2:
        c1, c2 = st.columns(2)
        pais_sel = c1.selectbox('País Inventario', ['EE.UU.', 'China', 'Europa'], key='sel_inv_pais')
        tech_sel = c2.selectbox('Tecnología — Inventario', ['Combustión', 'Híbrido', 'Eléctrico', 'Hidrógeno'], key='sel_inv_tech')
        log = df[(df['Estado'] == 'Detalles de logística') & (df['Empresa'] == empresa_analisis) & (df['Seccion'] == f'{tech_sel}, miles unidades') & (df['Subgrupo'] == pais_sel) & (df['Ronda'] == ronda_snapshot)].copy()
        log['Valor'] = num(log['Valor'])
        d = log.set_index('Metrica')['Valor']
        if not d.empty:
            inv_ini = d.get('Inventario inicial', 0) or 0
            prod = (d.get('Producción interna', 0) or 0) + (d.get('Producción contratada', 0) or 0)
            imp = sum(v for k, v in d.items() if k.startswith('Importado desde') and pd.notna(v))
            ventas = abs(d.get(f'Ventas en {pais_sel}', 0) or 0)
            exp = sum(abs(v) for k, v in d.items() if k.startswith('Exportado a') and pd.notna(v))
            inv_fin = d.get('Inventario final', 0) or 0
            _valores_inv = [inv_ini, prod, imp, -ventas, -exp, inv_fin]
            fig_inv = go.Figure(go.Waterfall(
                orientation='v', measure=['absolute', 'relative', 'relative', 'relative', 'relative', 'total'],
                x=['Inv Inicial', '+ Prod', '+ Import', '- Ventas', '- Export', '= Inv Final'],
                y=_valores_inv,
                # Adenda 25: % entre paréntesis sobre el Inventario Inicial (primera barra).
                text=_texto_cascada(_valores_inv), textposition='outside',
                # Adenda 25: convención universal de color -- verde = suma, rojo = resta. Antes
                # decreasing usaba MUTED_PALETTE[2] (verde grisáceo apagado, no rojo). Totals se
                # mantiene neutro (MUTED_PALETTE[0]): "Inv Final" acá no es identidad CADIZ ni
                # sentimiento (más/menos inventario final no es en sí bueno o malo), es un dato físico.
                decreasing={'marker': {'color': COLOR_NEGATIVE}}, increasing={'marker': {'color': COLOR_POSITIVE}},
                totals={'marker': {'color': MUTED_PALETTE[0]}}
            ))
            fig_inv.update_layout(title='Puente de Inventario Físico')
            mostrar(fig_inv, ocultar_eje_valores='y')

        st.divider()
        st.markdown('**Gap de pronóstico: demanda insatisfecha vs. inventario que sobró**')
        st.caption('Evolución de las dos formas de errar la estimación de demanda: faltante (no llegaste a vender lo que te pedían) y sobrante (produjiste de más y quedó en depósito).')
        log_hist = df[(df['Estado'] == 'Detalles de logística') & (df['Empresa'] == empresa_analisis) &
                      (df['Seccion'] == f'{tech_sel}, miles unidades') & (df['Subgrupo'] == pais_sel) &
                      (df['Metrica'].isin(['Demanda insatisfecha', 'Inventario final']))].copy()
        log_hist['Valor'] = num(log_hist['Valor'])
        log_hist = log_hist.dropna(subset=['Valor']).sort_values('Ronda_Orden')
        if not log_hist.empty:
            fig_gap = go.Figure()
            for metrica, color, nombre in [('Demanda insatisfecha', COLOR_NEGATIVE, 'Faltante (demanda insatisfecha)'),
                                            ('Inventario final', MUTED_PALETTE[0], 'Sobrante (inventario final)')]:
                d_m = log_hist[log_hist['Metrica'] == metrica]
                fig_gap.add_trace(go.Bar(x=d_m['Ronda'], y=d_m['Valor'], name=nombre, marker_color=color))
            fig_gap.update_layout(barmode='group', title=f'Faltante vs. sobrante — {empresa_analisis}, {tech_sel}, {pais_sel}',
                                   legend=dict(orientation="h", yanchor="top", y=-0.30, xanchor="center", x=0.5))
            mostrar(fig_gap)
        else:
            st.info('Sin datos suficientes para esta combinación.')

        st.divider()
        log_tech = df[(df['Estado'] == 'Detalles de logística') & (df['Empresa'] == empresa_analisis) & (df['Seccion'] == f'{tech_sel}, miles unidades') & (df['Ronda'] == ronda_snapshot)]
        def val_log(planta, metrica):
            r = log_tech[(log_tech['Subgrupo'] == planta) & (log_tech['Metrica'] == metrica)]['Valor']
            return abs(pd.to_numeric(r.iloc[0], errors='coerce')) if len(r) else 0.0
        destinos = ['EE.UU.', 'China', 'Europa']
        matriz = pd.DataFrame(0.0, index=['EE.UU.', 'China', 'Subcontratado'], columns=destinos)
        for planta in ['EE.UU.', 'China']:
            interna, contratada = val_log(planta, 'Producción interna'), val_log(planta, 'Producción contratada')
            tot = interna + contratada
            flows = {dst: val_log(planta, f'Ventas en {planta}' if dst == planta else f'Exportado a {dst}') for dst in destinos}
            if tot > 0:
                for dst in destinos:
                    matriz.loc[planta, dst] += flows[dst] * (interna/tot)
                    matriz.loc['Subcontratado', dst] += flows[dst] * (contratada/tot)
        if matriz.sum().sum() > 0:
            # Heatmap de magnitud (flujo logístico), sin identidad ni sentimiento -- azul de
            # COLOR_METRICA['dinero'] en vez del lila/rojo de marca (Adenda 25).
            fig3 = px.imshow(matriz.values, x=destinos, y=matriz.index, text_auto='.0f', aspect='auto', color_continuous_scale=[[0, BRAND_LIGHT], [1, COLOR_METRICA['dinero']]])
            fig3.update_coloraxes(showscale=False)
            fig3.update_xaxes(title_text='Destino (Mercado)')
            fig3.update_yaxes(title_text='Origen (Planta)')
            fig3.update_traces(xgap=3, ygap=3) 
            fig3.update_layout(title='Matriz Logística (Heatmap Origen -> Destino)')
            mostrar(fig3)
        else:
            st.info(f"Sin flujos logísticos de {tech_sel} para mostrar en esta ronda.")
# =================================================================
# SECCIÓN 4 — FINANZAS (Corto y Largo Plazo)
# =================================================================
def seccion_finanzas():
    # =================================================================
    # Adenda 25 -- RESTRUCTURACIÓN DE FINANZAS (a pedido del equipo). Cambios:
    #  1) KPIs principales (banda oscura): antes EBITDA/Caja final/Deuda CP no planificada -> ahora
    #     Precio Acción/Capitalización de mercado/EPS (la "foto" de una acción cotizante).
    #  2) KPIs secundarios: antes Margen bruto/ROS/Deuda LP/Calificación -> ahora Ingresos por
    #     ventas/EBITDA/Beneficio neto/Calificación. Margen bruto y ROS no se pierden: pasan al
    #     Gráfico 1 de más abajo.
    #  3) Gráfico 1 (nuevo): desglose Ingresos/Margen bruto/EBITDA/EBIT/Margen Neto en un solo eje.
    #  4) Gráfico 2 (nuevo), al lado: evolución del Precio de la Acción (reusa chart_evolucion, los
    #     7 equipos + promedio, como el resto de la app).
    #  5) Estructura del Balance se reposiciona debajo de Gráfico 1/2 (antes vivía en la pestaña
    #     "Largo Plazo").
    #  6) Debajo del Balance: 3 KPIs -- WACC, ROCE, Deuda LP (reemplazan la tarjeta única de
    #     "Spread de Creación de Valor").
    #  7) Rango de Industria se recorta a WACC/Apalancamiento/ROA/ROE/ROCE (se saca "Spread
    #     ROCE-WACC", que no lo pidió el equipo).
    #  8) Se eliminan 3 gráficos: "Costo de la deuda por mercado", "Matriz Riesgo/Retorno" y
    #     "Beneficio Neto vs. Nivel de Deuda". ("Evolución del margen bruto" -- el 4° de la lista
    #     original -- no se encontró como gráfico propio en esta sección; puede haberse eliminado ya
    #     en una etapa anterior, o el equipo puede estar refiriéndose al Gráfico 1 nuevo, que sí
    #     incluye Margen bruto. Señalado acá para no asumir en silencio.)
    #  9) En Comparativa Plan vs. Real se saca el llamado a fila3_finanzas_flujo_caja() (el puente
    #     CFO->CFI->CFF) -- queda solo el comparativo lateral Proyectado vs. Real. La función sigue
    #     definida más arriba, sin uso, por si se retoma.
    #
    # Adenda 26 -- AJUSTES SOBRE LA RESTRUCTURACIÓN DE ARRIBA (a pedido del equipo, feedback sobre
    # la Adenda 25). Cambios:
    #  A) Pestañas al TOPE de la sección (debajo del título "Finanzas", antes de las tarjetas de
    #     cabecera) y recortadas a 2: "Finanzas" (todo lo que antes vivía suelto: cabecera, Gráfico
    #     1/2, Balance, KPIs WACC/ROCE/Deuda LP, Rango de Industria, y "Deuda y liquidez" que antes
    #     vivía en la pestaña "Corto Plazo") y "Comparativa Plan vs. Real".
    #  B) Gráfico 1 (desglose de Ingresos y Márgenes): ahora SÍ muestra las etiquetas de categoría en
    #     el eje Y (antes se ocultaban con ocultar_eje_valores='y', quedaba sin texto a la izquierda)
    #     -- se saca ese parámetro del llamado a mostrar(). Etiquetas de categoría renombradas para
    #     que se lean como conceptos de estado de resultados en vez de "Margen X" repetido:
    #     "Margen bruto"->"Margen Bruto", "Margen EBITDA"->"EBITDA", "Margen EBIT"->"EBIT / Resultado
    #     Operativo", "Margen Neto (ROS)"->"Beneficio Neto" (el valor y el % que muestran son los
    #     MISMOS ratios de antes, solo cambia el texto de la etiqueta).
    #  C) "Evolución — Deuda CP no planificada" (el gráfico de línea temporal) se elimina. La lógica
    #     de alerta sigue viva: ya existía en evaluar_alertas()/panel_alertas() (regla "Sobregiro
    #     automático", mostrada en la sección Resultados) y ahora se agrega ADEMÁS un st.toast() acá
    #     mismo en Finanzas para no perder la señal visual que daba el gráfico que se sacó.
    #  D) Se reincorpora un badge de "Spread de Creación de Valor" (ROCE − WACC) y se agrega un
    #     gráfico nuevo de "Composición de Estructura de Capital" (ponderación Deuda/Patrimonio del
    #     propio cálculo de WACC) -- los dos, debajo de los 3 KPIs de WACC/ROCE/Deuda LP.
    #  E) panel_comparativa_plan_real(): se llama con mostrar_metricas=False solo desde acá -- saca
    #     la Fila 1 de tarjetas st.metric y deja solo el panel de barras horizontales apiladas como
    #     elemento protagonista (ver comentario en la definición de esa función).
    #  DECISIÓN DE LAYOUT PROPIA (no especificada por el equipo, a flagear): con Balance, Rango de
    #  Industria y los KPIs de Largo Plazo promovidos al bloque fijo de arriba, y con "Costo de la
    #  deuda"/"Matriz Riesgo-Retorno"/"Beneficio Neto vs Deuda" eliminados, la pestaña "Largo Plazo:
    #  Estructura, Retorno y Competencia" se quedaba vacía -- se elimina esa pestaña y "Deuda y
    #  liquidez" (lo único que quedaba en "Corto Plazo") pasa a vivir dentro de la pestaña "Finanzas".
    # =================================================================
    pl_ronda = df[(df['Estado'] == 'Cuenta de resultados, miles USD, Global') & (df['Ronda'] == ronda_snapshot)]
    ratios_ronda = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Ronda'] == ronda_snapshot)]
    val_ronda = df[(df['Estado'] == 'Valuación - Global') & (df['Ronda'] == ronda_snapshot)]
    bal_ronda = df[(df['Estado'] == 'Hoja de Balance, miles USD, Global') & (df['Ronda'] == ronda_snapshot)]

    def wacc(emp, val_df=None):
        val_df = val_ronda if val_df is None else val_df
        de = valor_de(val_df, 'Deuda a patrimonio', emp)
        re_ = valor_de(val_df, 'Rendimiento esperado del patrimonio, %', emp)
        rd = valor_de(val_df, 'Costo de la deuda después de impuestos, %', emp)
        if None in (de, re_, rd): return None
        return (1/(1+de))*re_ + (de/(1+de))*rd

    def delta_raw(vals, empresa=empresa_analisis):
        # Mismo cálculo que el resto de la app: % de distancia contra el promedio de los 7 equipos.
        prom = np.nanmean([v for v in vals.values() if v is not None]) if any(v is not None for v in vals.values()) else None
        val = vals.get(empresa)
        if prom in (None, 0) or val is None:
            return None
        return (val - prom) / abs(prom) * 100

    # Adenda 26: pestañas al TOPE de la sección, antes de cualquier tarjeta -- ver punto (A) arriba.
    tab_finanzas, tab_cg = st.tabs(['Finanzas', 'Comparativa Plan vs. Real'])

    with tab_finanzas:
        # ---- KPIs principales: Precio Acción / Capitalización / EPS ----
        precio_vals = {e: valor_de(ratios_ronda, 'Precio de la acción al final de la ronda, USD', e) for e in COMPANIES}
        cap_vals = {e: valor_de(val_ronda, 'Capitalización de mercado, miles USD', e) for e in COMPANIES}
        eps_vals = {e: valor_de(ratios_ronda, 'Ganancias por acción (EPS), USD', e) for e in COMPANIES}
        d_precio_ronda, ronda_prev_fin = delta_vs_ronda_anterior('Ratios e indicadores financieros clave', 'Precio de la acción al final de la ronda, USD')
        d_cap_ronda, _ = delta_vs_ronda_anterior('Valuación - Global', 'Capitalización de mercado, miles USD')
        d_eps_ronda, _ = delta_vs_ronda_anterior('Ratios e indicadores financieros clave', 'Ganancias por acción (EPS), USD')
        d_precio_ind, d_cap_ind, d_eps_ind = delta_raw(precio_vals), delta_raw(cap_vals), delta_raw(eps_vals)
        sufijo_ronda_fin = f'vs {ronda_prev_fin}' if ronda_prev_fin else 'vs ronda anterior'
        kpi_banda_oscura([
            {'label': 'Precio Acción (USD)', 'valor': f"{precio_vals.get(empresa_analisis):,.2f}" if pd.notna(precio_vals.get(empresa_analisis)) else '—',
             'delta': f'{d_precio_ronda:+.1f}% {sufijo_ronda_fin}' if d_precio_ronda is not None else None, 'favorable': (d_precio_ronda > 0) if d_precio_ronda is not None else None,
             'delta2': f'{d_precio_ind:+.1f}% vs Industria' if d_precio_ind is not None else None, 'favorable2': (d_precio_ind > 0) if d_precio_ind is not None else None},
            {'label': 'Capitalización de mercado (USD)', 'valor': format_num(cap_vals.get(empresa_analisis)) if pd.notna(cap_vals.get(empresa_analisis)) else '—',
             'delta': f'{d_cap_ronda:+.1f}% {sufijo_ronda_fin}' if d_cap_ronda is not None else None, 'favorable': (d_cap_ronda > 0) if d_cap_ronda is not None else None,
             'delta2': f'{d_cap_ind:+.1f}% vs Industria' if d_cap_ind is not None else None, 'favorable2': (d_cap_ind > 0) if d_cap_ind is not None else None},
            {'label': 'EPS (USD)', 'valor': f"{eps_vals.get(empresa_analisis):,.2f}" if pd.notna(eps_vals.get(empresa_analisis)) else '—',
             'delta': f'{d_eps_ronda:+.1f}% {sufijo_ronda_fin}' if d_eps_ronda is not None else None, 'favorable': (d_eps_ronda > 0) if d_eps_ronda is not None else None,
             'delta2': f'{d_eps_ind:+.1f}% vs Industria' if d_eps_ind is not None else None, 'favorable2': (d_eps_ind > 0) if d_eps_ind is not None else None},
        ])

        # ---- KPIs secundarios: Ingresos por ventas / EBITDA / Beneficio neto / Calificación ----
        ingresos_vals = {e: valor_de(pl_ronda, 'Ingresos por ventas', e) for e in COMPANIES}
        ebitda_vals = {e: valor_de(pl_ronda, 'Beneficio operativo antes de depreciación (EBITDA)', e) for e in COMPANIES}
        beneficio_neto_vals = {e: valor_de(pl_ronda, 'Beneficio de la ronda', e) for e in COMPANIES}
        # Margen bruto, ROS y Deuda CP/LP se siguen calculando -- Margen bruto/ROS alimentan el
        # Gráfico 1 de más abajo, Deuda CP sigue en "Deuda y liquidez", Deuda LP en los 3 KPIs junto
        # al Balance.
        margen_vals = {e: valor_de(ratios_ronda, 'Margen bruto', e) for e in COMPANIES}
        ros_vals = {e: valor_de(ratios_ronda, 'Rentabilidad de las ventas (ROS)', e) for e in COMPANIES}
        ebitda_pct_vals = {e: valor_de(ratios_ronda, 'Beneficio operativo antes de depreciación ( EBITDA )', e) for e in COMPANIES}
        ebit_pct_vals = {e: valor_de(ratios_ronda, 'BENEFICIO OPERATIVO (EBIT)', e) for e in COMPANIES}
        ebit_usd_vals = {e: valor_de(pl_ronda, 'Beneficio operativo (EBIT)', e) for e in COMPANIES}
        deuda_cp_vals = {e: valor_de(bal_ronda, 'Deudas a corto plazo (no planificadas)', e) for e in COMPANIES}
        deuda_lp_vals = {e: valor_de(bal_ronda, 'Deudas a largo plazo', e) for e in COMPANIES}
        calif_val = valor_texto(ratios_ronda, 'Calificación crediticia', empresa_analisis)

        d_ingresos_ronda, _ = delta_vs_ronda_anterior('Cuenta de resultados, miles USD, Global', 'Ingresos por ventas')
        d_ebitda_ronda, _ = delta_vs_ronda_anterior('Cuenta de resultados, miles USD, Global', 'Beneficio operativo antes de depreciación (EBITDA)')
        d_bn_ronda, _ = delta_vs_ronda_anterior('Cuenta de resultados, miles USD, Global', 'Beneficio de la ronda')
        f2, f3, f6, f7 = st.columns(4)
        with f2:
            kpi_doble_delta(None, 'Ingresos por ventas (USD)', format_num(ingresos_vals.get(empresa_analisis)),
                             d_ronda=d_ingresos_ronda, ronda_prev_nombre=ronda_prev_fin, d_industria=delta_raw(ingresos_vals))
        with f3:
            kpi_doble_delta(None, 'EBITDA (USD)', format_num(ebitda_vals.get(empresa_analisis)),
                             d_ronda=d_ebitda_ronda, ronda_prev_nombre=ronda_prev_fin, d_industria=delta_raw(ebitda_vals))
        with f6:
            kpi_doble_delta(None, 'Beneficio neto (USD)', format_num(beneficio_neto_vals.get(empresa_analisis)),
                             d_ronda=d_bn_ronda, ronda_prev_nombre=ronda_prev_fin, d_industria=delta_raw(beneficio_neto_vals))
        with f7: st.metric('Calificación crediticia', calif_val if calif_val else '—')
        st.write('')

        # ---- Gráfico 1 (desglose común-tamaño) + Gráfico 2 (evolución del precio) ----
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            ing_emp = ingresos_vals.get(empresa_analisis)
            if ing_emp:
                # Adenda 26: etiquetas renombradas a pedido del equipo (mismo dato/ratio de antes,
                # solo cambia el texto que se lee en el eje Y -- ver punto (B) arriba).
                filas_cs = [
                    ('Ingresos por ventas', 100.0, ing_emp),
                    ('Margen Bruto', margen_vals.get(empresa_analisis), None),
                    ('EBITDA', ebitda_pct_vals.get(empresa_analisis), ebitda_vals.get(empresa_analisis)),
                    ('EBIT / Resultado Operativo', ebit_pct_vals.get(empresa_analisis), ebit_usd_vals.get(empresa_analisis)),
                    ('Beneficio Neto', ros_vals.get(empresa_analisis), beneficio_neto_vals.get(empresa_analisis)),
                ]
                filas_cs = [(n, p, u) for n, p, u in filas_cs if pd.notna(p)]
                if filas_cs:
                    cs_df = pd.DataFrame(filas_cs, columns=['Concepto', 'Pct', 'USD'])
                    cs_df['Etiqueta'] = cs_df.apply(lambda r: f"{r['Pct']:.1f}% ({format_num(r['USD'])})" if pd.notna(r['USD']) else f"{r['Pct']:.1f}%", axis=1)
                    # Color = identidad del EQUIPO EN FOCO (COLOR_MAP ya resuelve lila solo si es CADIZ).
                    fig_cs = px.bar(cs_df.iloc[::-1], x='Pct', y='Concepto', orientation='h', text='Etiqueta',
                                     color_discrete_sequence=[COLOR_MAP.get(empresa_analisis, COLOR_CADIZ)],
                                     title=f'Desglose de Ingresos y Márgenes (% de Ingresos) — {empresa_analisis}, {ronda_snapshot}')
                    fig_cs.update_traces(textposition='outside', cliponaxis=False)
                    # Adenda 27: se sacan los títulos de los dos ejes (a pedido del equipo) -- las
                    # etiquetas de categoría (eje Y) y los valores de % (eje X, con el texto de cada
                    # barra) ya se entienden solos, sin necesidad de un título de eje aparte.
                    fig_cs.update_layout(xaxis_title=None, yaxis_title=None, showlegend=False)
                    # Adenda 26: se saca ocultar_eje_valores='y' -- a pedido del equipo, ahora se
                    # fuerza a mostrar el nombre de cada concepto sobre el eje vertical (antes
                    # quedaba en blanco a la izquierda, solo con la barra y el texto de afuera).
                    mostrar(fig_cs)
                    # Adenda 27: se saca el caption de abajo (a pedido del equipo, "no es necesario").
                else:
                    st.info('Sin datos de márgenes para esta combinación.')
            else:
                st.info('Sin datos de Ingresos por ventas para esta combinación.')
        with col_g2:
            precio_sub = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Metrica'] == 'Precio de la acción al final de la ronda, USD')]
            chart_evolucion(precio_sub, 'Precio de la Acción (USD)')
        st.write('')

        # ---- Estructura del Balance (reposicionada -- antes vivía en la pestaña "Largo Plazo") ----
        st.markdown('**Estructura del Balance: Activo vs. Pasivo + Patrimonio Neto**')

        def gb(metrica):
            return valor_de(bal_ronda, metrica, empresa_analisis) or 0.0

        activo_items = {'Efectivo y equivalentes': gb('Efectivo y equivalentes de efectivo'),
                         'Cuentas por cobrar': gb('Cuentas por Cobrar'), 'Inventario': gb('Inventario'),
                         'Activo fijo': gb('Activo fijo')}
        pasivo_pn_items = {'Cuentas por pagar': gb('Cuentas por pagar'),
                            'Deudas CP no planificadas': gb('Deudas a corto plazo (no planificadas)'),
                            'Deudas LP': gb('Deudas a largo plazo'),
                            'Capital social + adicional': gb('Capital social') + gb('Capital adicional desembolsado'),
                            'Ganancias acumuladas + de la ronda': gb('Ganancias acumuladas') + gb('Beneficio de la ronda')}
        if sum(activo_items.values()) > 0:
            fig_bal = go.Figure()
            # Paletas separadas por lado: antes el verde era "Activo fijo" a la izquierda y
            # "Deudas LP" a la derecha, y el tan era "Inventario" y "Ganancias acumuladas".
            # Con la misma paleta de los dos lados parecía que un color significaba lo mismo.
            colores_activo = ['#3E7CB1', '#5C9BC9', '#8FBEDC', '#C3DCEC']          # azules  = qué tengo
            colores_pasivo = ['#C9922E', '#B3261E', '#8C6D3F', '#6E8C6E', '#A8A29A']  # cálidos = quién lo financia
            for (nombre, val), color in zip(activo_items.items(), colores_activo):
                fig_bal.add_trace(go.Bar(x=['Activo'], y=[val], name=nombre, marker_color=color,
                                          text=format_num(val), textposition='inside', showlegend=False))
            for (nombre, val), color in zip(pasivo_pn_items.items(), colores_pasivo):
                fig_bal.add_trace(go.Bar(x=['Pasivo + PN'], y=[val], name=nombre, marker_color=color,
                                          text=format_num(val), textposition='inside', showlegend=False))
            fig_bal.update_layout(
                barmode='stack', title=f'Estructura del Balance — {empresa_analisis}, {ronda_snapshot}',
                # Dos leyendas separadas, cada una pegada a la barra que le corresponde —
                # antes una sola leyenda combinada hacía difícil saber qué color era de qué lado.
                showlegend=False)
            mostrar(fig_bal, ocultar_eje_valores='y')
            # La leyenda de Plotly, aun puesta afuera, mezclaba los conceptos de los dos lados en
            # una sola tira. Se dibuja a mano en dos columnas, cada una bajo su barra, para que
            # se vea de una qué compone el Activo y qué compone el Pasivo + PN.
            def bloque_leyenda(titulo, items, colores):
                filas = ''.join(
                    f'<div style="display:flex;align-items:center;gap:8px;margin:3px 0;font-size:0.82rem">'
                    f'<span style="width:11px;height:11px;border-radius:3px;background:{c};flex:none"></span>'
                    f'<span style="flex:1">{n}</span>'
                    f'<span style="opacity:0.65">{format_num(v)}</span></div>'
                    for (n, v), c in zip(items.items(), colores))
                return (f'<div style="font-weight:600;font-size:0.8rem;opacity:0.7;text-transform:uppercase;'
                        f'letter-spacing:0.02em;margin-bottom:4px">{titulo}</div>{filas}')
            col_leg_a, col_leg_p = st.columns(2)
            with col_leg_a:
                st.markdown(bloque_leyenda('Activo', activo_items, colores_activo), unsafe_allow_html=True)
            with col_leg_p:
                st.markdown(bloque_leyenda('Pasivo + PN', pasivo_pn_items, colores_pasivo), unsafe_allow_html=True)
            st.caption('Los dos lados deben dar la misma altura (el Balance siempre cierra) — Activo total = Pasivo + Patrimonio Neto.')
        else:
            st.info('Sin datos de balance para esta combinación.')

        # ---- 3 KPIs junto al Balance: WACC, ROCE, Deuda LP ----
        # Adenda 23 (a pedido del equipo): ROA no lo publica CESIM como ratio propio (verificado -- no
        # está en 'Ratios e indicadores financieros clave') así que se calcula acá con la fórmula
        # estándar Beneficio Neto / Activos Totales (Global) -- es una métrica de análisis financiero de
        # manual de cátedra, no una regla CESIM, así que se documenta como tal.
        def _roa(emp):
            ben = valor_de(pl_ronda, 'Beneficio de la ronda', emp)
            act = valor_de(bal_ronda, 'Activos Totales', emp)
            if ben is None or not act:
                return None
            return ben / act * 100

        # Orden pedido por el equipo para el Rango de Industria: WACC, Apalancamiento, ROA, ROE, ROCE.
        datos_lp = {
            'WACC': {e: wacc(e) for e in COMPANIES},
            'Apalancamiento': {e: valor_de(ratios_ronda, 'Endeudamiento neto/patrimonio (apalancamiento)', e) for e in COMPANIES},
            'ROA': {e: _roa(e) for e in COMPANIES},
            'ROE': {e: valor_de(ratios_ronda, 'Rendimiento de los Fondos Propios (ROE)', e) for e in COMPANIES},
            'ROCE': {e: valor_fuzzy(ratios_ronda, 'Rentabilidad del capital empleado', empresa=e) for e in COMPANIES},
        }
        ronda_prev_lp = ronda_anterior_de(ronda_snapshot)
        d_wacc_ronda = d_roce_ronda = None
        if ronda_prev_lp:
            val_ronda_prev = df[(df['Estado'] == 'Valuación - Global') & (df['Ronda'] == ronda_prev_lp)]
            ratios_ronda_prev = df[(df['Estado'] == 'Ratios e indicadores financieros clave') & (df['Ronda'] == ronda_prev_lp)]
            wacc_prev, wacc_act = wacc(empresa_analisis, val_ronda_prev), wacc(empresa_analisis)
            if wacc_prev not in (None, 0) and wacc_act is not None and pd.notna(wacc_prev) and pd.notna(wacc_act):
                d_wacc_ronda = (wacc_act - wacc_prev) / abs(wacc_prev) * 100
            roce_prev = valor_fuzzy(ratios_ronda_prev, 'Rentabilidad del capital empleado', empresa=empresa_analisis)
            roce_act = valor_fuzzy(ratios_ronda, 'Rentabilidad del capital empleado', empresa=empresa_analisis)
            if roce_prev not in (None, 0) and roce_act is not None and pd.notna(roce_prev) and pd.notna(roce_act):
                d_roce_ronda = (roce_act - roce_prev) / abs(roce_prev) * 100
        d_wacc_ind, d_roce_ind = delta_raw(datos_lp['WACC']), delta_raw(datos_lp['ROCE'])
        d_deuda_lp_ronda, _ = delta_vs_ronda_anterior('Hoja de Balance, miles USD, Global', 'Deudas a largo plazo')
        d_deuda_lp_ind = delta_raw(deuda_lp_vals)
        st.write('')
        col_w1, col_w2, col_w3 = st.columns(3)
        with col_w1:
            # WACC: costo de financiarse -- menos es mejor, favorable invertido.
            kpi_doble_delta(None, 'WACC', f"{datos_lp['WACC'].get(empresa_analisis):,.1f}%" if pd.notna(datos_lp['WACC'].get(empresa_analisis)) else '—',
                             d_ronda=d_wacc_ronda, ronda_prev_nombre=ronda_prev_fin, favorable_ronda=(d_wacc_ronda < 0) if d_wacc_ronda is not None else None,
                             d_industria=d_wacc_ind, favorable_industria=(d_wacc_ind < 0) if d_wacc_ind is not None else None)
        with col_w2:
            kpi_doble_delta(None, 'ROCE', f"{datos_lp['ROCE'].get(empresa_analisis):,.1f}%" if pd.notna(datos_lp['ROCE'].get(empresa_analisis)) else '—',
                             d_ronda=d_roce_ronda, ronda_prev_nombre=ronda_prev_fin, d_industria=d_roce_ind)
        with col_w3:
            # Deuda LP: menos es mejor, favorable invertido.
            kpi_doble_delta(None, 'Deuda LP (USD)', format_num(deuda_lp_vals.get(empresa_analisis)),
                             d_ronda=d_deuda_lp_ronda, ronda_prev_nombre=ronda_prev_fin, favorable_ronda=(d_deuda_lp_ronda < 0) if d_deuda_lp_ronda is not None else None,
                             d_industria=d_deuda_lp_ind, favorable_industria=(d_deuda_lp_ind < 0) if d_deuda_lp_ind is not None else None)

        # ---- Spread de Creación de Valor + Composición de Estructura de Capital (Adenda 26, a ----
        # ---- pedido del equipo: "debajo o cerca de los kpis de wacc, roce y deuda lp") ---------
        st.write('')
        roce_focal, wacc_focal = datos_lp['ROCE'].get(empresa_analisis), datos_lp['WACC'].get(empresa_analisis)
        spread = (roce_focal - wacc_focal) if (pd.notna(roce_focal) and pd.notna(wacc_focal)) else None
        de_focal = valor_de(val_ronda, 'Deuda a patrimonio', empresa_analisis)
        col_sp, col_wc = st.columns(2)
        with col_sp:
            # Spread = ROCE − WACC: fórmula de manual de cátedra (Categoría 3, no un ratio propio de
            # CESIM) a partir de dos ratios que CESIM sí publica y que ya están arriba (WACC, ROCE).
            kpi_doble_delta(None, 'Spread de Creación de Valor (ROCE − WACC)',
                             f"{spread:+.1f} p.p." if spread is not None else '—', unidad=' p.p.')
            st.caption('ROCE > WACC = la ronda creó valor para el accionista (el capital empleado rindió más de lo '
                       'que costó financiarlo); ROCE < WACC = lo destruyó.')
        with col_wc:
            if de_focal is not None and pd.notna(de_focal) and de_focal >= 0:
                # Mismos pesos que usa wacc() de arriba: de/(1+de) pondera el costo de la deuda,
                # 1/(1+de) pondera el retorno esperado del patrimonio -- se visualiza esa relación
                # como % Deuda / % Patrimonio Neto (Equity) de la estructura de capital.
                peso_deuda = de_focal / (1 + de_focal) * 100
                peso_equity = 100 - peso_deuda
                fig_wacc_c = go.Figure()
                fig_wacc_c.add_trace(go.Bar(x=[peso_deuda], y=['Estructura'], orientation='h', name='Deuda',
                                             marker_color=COLOR_METRICA['riesgo'],
                                             text=f'Deuda {peso_deuda:.0f}%', textposition='inside'))
                # Adenda 28 (feedback "más contraste, paleta única"): Equity pasa de MUTED_PALETTE[2]
                # (sage -- reprobó el validador de contraste, ΔE 13.4 contra el ámbar de Deuda, por
                # debajo del piso de 15) a COLOR_METRICA['dinero'] (azul), el mismo azul que ahora abre
                # MUTED_SIN_MARRON -- así este gráfico y el de Mix tecnológico comparten paleta.
                fig_wacc_c.add_trace(go.Bar(x=[peso_equity], y=['Estructura'], orientation='h', name='Patrimonio Neto (Equity)',
                                             marker_color=COLOR_METRICA['dinero'],
                                             text=f'Equity {peso_equity:.0f}%', textposition='inside'))
                fig_wacc_c.update_layout(barmode='stack', showlegend=False,
                                          title=f'Composición de Estructura de Capital (ponderación del WACC) — {empresa_analisis}',
                                          xaxis=dict(range=[0, 100], ticksuffix='%'))
                # Adenda 27 (a pedido del equipo: "que quede del mismo tamaño que el kpi de creación
                # de valor"): altura chica en vez de la ALTURA_TARJETA fija (370px) que usa el resto
                # de los gráficos -- acá se lo compara con una tarjeta kpi_doble_delta de al lado
                # (col_sp), mucho más baja, y sin leyenda abajo no hace falta el margen extra.
                mostrar(fig_wacc_c, ocultar_eje_valores='y', altura=150, margen_b=10)
                st.caption('Ponderación usada en el cálculo de WACC de arriba: Deuda a patrimonio / (1 + Deuda a '
                           'patrimonio) para la Deuda, y 1 / (1 + Deuda a patrimonio) para el Patrimonio Neto.')
            else:
                st.info('Sin datos de Deuda a patrimonio para esta combinación.')

        # ---- Rango de Industria (recortado: WACC, Apalancamiento, ROA, ROE, ROCE -- sin Spread) ----
        ejes_validos = {k: v for k, v in datos_lp.items() if len([x for x in v.values() if pd.notna(x)]) >= 2}
        if ejes_validos:
            color_ref = 'rgba(255,255,255,0.5)' if es_modo_oscuro() else 'rgba(26,23,20,0.5)'
            fig_rango = go.Figure()
            for i, (nombre, vals) in enumerate(ejes_validos.items()):
                valores = sorted(v for v in vals.values() if pd.notna(v))
                vmin, vmax, vmed = valores[0], valores[-1], np.median(valores)
                vcadiz = vals.get(empresa_analisis)
                rango = (vmax - vmin) or 1
                pos = lambda x: (x - vmin) / rango * 100
                suf = "%" if nombre in ['ROCE', 'ROE', 'WACC', 'ROA'] else "x"

                fig_rango.add_trace(go.Scatter(x=[0, 100], y=[i, i], mode='lines', line=dict(color=MUTED_PALETTE[3], width=6), showlegend=False))
                # Adenda 25: se saca el marrón de la marca de mediana. No pasa a lila -- esta marca
                # convive con el marcador del equipo en foco (unas líneas más abajo), que puede ser
                # lila si el foco es CADIZ; usar lila acá también las confundiría.
                # Fix "ya casi": se reemplaza MUTED_PALETTE[5] por el azul acero nuevo, por consistencia
                # con el resto de la paleta "sin marrón" (mismo tono ya usado en el dumbbell de mercado).
                fig_rango.add_trace(go.Scatter(x=[pos(vmed)], y=[i], mode='markers', marker=dict(symbol='line-ns', size=16, color='#58759D', line_width=2), showlegend=False))

                if vcadiz is not None:
                    val_str = f"{vcadiz:,.1f}{suf}"
                    # Bug de identidad (Adenda 25, mismo caso que la tasa de interés más arriba):
                    # "vcadiz" es en realidad el valor del EQUIPO EN FOCO (empresa_analisis), que puede
                    # no ser CADIZ -- se resuelve con COLOR_MAP en vez de asumir siempre el color de
                    # marca.
                    color_foco = COLOR_MAP.get(empresa_analisis, COLOR_CADIZ)
                    fig_rango.add_trace(go.Scatter(
                        x=[pos(vcadiz)], y=[i],
                        mode='markers+text',
                        text=[val_str],
                        textposition="top center",
                        textfont=dict(color=color_foco, size=12, family="JetBrains Mono"),
                        marker=dict(size=14, color=color_foco),
                        showlegend=False
                    ))

                fig_rango.add_annotation(x=0, y=i, text=f'{vmin:,.1f}{suf}', showarrow=False, xshift=-30, font=dict(size=11, color=color_ref))
                fig_rango.add_annotation(x=100, y=i, text=f'{vmax:,.1f}{suf}', showarrow=False, xshift=30, font=dict(size=11, color=color_ref))

            fig_rango.update_layout(yaxis=dict(tickmode='array', tickvals=list(range(len(ejes_validos))), ticktext=list(ejes_validos.keys())), xaxis=dict(range=[-15, 115]), title='Rango de Industria (Mín / Mediana / CÁDIZ / Máx)')
            mostrar(fig_rango, ocultar_eje_valores='x')

        # ---- Deuda y liquidez (Adenda 26: antes vivía en la pestaña "Corto Plazo", ahora acá) ----
        st.write('')
        st.subheader('Deuda y liquidez')
        # Adenda 26 (a pedido del equipo: "remover el gráfico de Evolución -- Deuda CP no
        # planificada, mantener únicamente la lógica condicional en segundo plano para disparar un
        # toast/banner de alerta solo si el valor llega a ser mayor a cero"): el gráfico de línea
        # temporal se saca. La regla "Sobregiro automático" YA existía en evaluar_alertas() y se ve
        # en el panel de alertas de Resultados -- acá se agrega el mismo disparo como st.toast() para
        # no perder la señal en Finanzas ahora que se sacó el gráfico que cumplía ese rol visual.
        deuda_cp_focal = deuda_cp_vals.get(empresa_analisis)
        if deuda_cp_focal and deuda_cp_focal > 0:
            st.toast(f'Sobregiro automático en {empresa_analisis}: {format_num(deuda_cp_focal)} USD de deuda CP '
                     f'no planificada en {ronda_snapshot} (detalle en el panel de Alertas, sección Resultados).',
                     icon='⚠️')
        ranking_deuda = pd.DataFrame(list(deuda_cp_vals.items()), columns=['Empresa', 'Valor']).dropna().sort_values('Valor', ascending=False)
        if not ranking_deuda.empty and ranking_deuda['Valor'].sum() > 0:
            ranking_deuda['Etiqueta'] = ranking_deuda['Valor'].apply(format_num)
            fig_dcp = px.bar(ranking_deuda, x='Empresa', y='Valor', color='Empresa', color_discrete_map=COLOR_MAP,
                              title=f'Deuda CP no planificada — {ronda_snapshot}', text='Etiqueta')
            fig_dcp.update_traces(textposition='outside', cliponaxis=False, showlegend=False)
            mostrar(fig_dcp, ocultar_eje_valores='y')
        else:
            st.info('Ningún equipo tomó deuda de corto plazo no planificada en esta ronda.')

    # Control de Gestión: es inherentemente sobre CADIZ (es nuestra propia proyección, no la de
    # "Equipo en foco") -- si se está mirando otro equipo, se avisa en vez de mostrar el gap de
    # CADIZ sin aclarar de quién es. Convertido de toggle suelto a pestaña (Adenda 11) por
    # consistencia con Resultados/Mercado/Operaciones, que ahora tienen la misma pestaña.
    with tab_cg:
        if empresa_analisis == MY_COMPANY:
            # Adenda 26/27 (a pedido del equipo: "eliminar todo esos gráficos de comparativa vs
            # real en finanzas, solo nos quedamos con la de comparativa vs real de barras
            # horizontales"): mostrar_metricas=False + mostrar_extras=False dejan ÚNICAMENTE el
            # panel chart_bullet_panel -- ver comentario en panel_comparativa_plan_real().
            panel_comparativa_plan_real(df, ronda_snapshot, key_suffix='finanzas', mostrar_directo=True,
                                         mostrar_metricas=False, mostrar_extras=False)
            # Adenda 25 (a pedido del equipo: "en comparativa vs real: solo dejamos las barras
            # laterales de proyectado vs real por ahora"): se saca el llamado a
            # fila3_finanzas_flujo_caja() (el puente CFO->CFI->CFF) que iba acá.
        else:
            st.caption('Cambiá "Equipo en foco" a CADIZ en la barra lateral para ver la Comparativa Plan vs. Real.')
    # =================================================================
    # SECCIÓN 5 — RRHH Y SOSTENIBILIDAD
    # =================================================================
def seccion_rrhh_sostenibilidad():
    bloque_rrhh, bloque_sost = st.tabs(['Personal y Talento', 'Sostenibilidad'])

    with bloque_rrhh:
        rrhh = df_all[(df_all['Estado'] == 'Informe de RRHH') & (df_all['Empresa'] == empresa_analisis)].copy()
        rrhh['Valor'] = num(rrhh['Valor'])

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader('Salario y rotación')
            salario = rrhh[rrhh['Metrica'] == 'Salario mensual, USD'].sort_values('Ronda_Orden')
            rotacion = rrhh[rrhh['Metrica'] == 'Rotación de personal, %'].sort_values('Ronda_Orden')
            if not salario.empty and not rotacion.empty:
                # Antes doble eje Y (Salario en USD y Rotación en % superpuestos en el mismo plano) --
                # dos paneles apilados, cada métrica con su propia escala.
                chart_dos_metricas_apiladas('Evolución: Salario vs Rotación',
                                             salario['Ronda'], salario['Valor'], 'Salario (USD)', COLOR_METRICA['dinero'], 'line',
                                             rotacion['Ronda'], rotacion['Valor'], 'Rotación (%)', COLOR_METRICA['riesgo'], 'line')
        with col_b:
            st.subheader('Rotación y contrataciones')
            contrat = rrhh[rrhh['Metrica'] == 'Contrataciones + / despidos -'].sort_values('Ronda_Orden')
            if not salario.empty and not contrat.empty and not rotacion.empty:
                chart_dos_metricas_apiladas('Rotación vs. Contrataciones netas',
                                             contrat['Ronda'], contrat['Valor'], 'Contrataciones netas (personas)', COLOR_METRICA['personas'], 'bar',
                                             rotacion['Ronda'], rotacion['Valor'], 'Rotación (%)', COLOR_METRICA['riesgo'], 'line')
                st.caption('Cuántas contrataciones netas hizo falta hacer, en la misma ronda en que se dio la rotación.')

        st.divider()
        col_c, col_d = st.columns(2)
        with col_c:
            st.subheader('Inversión en I+D')
            idn = rrhh[rrhh['Metrica'] == 'Número de personal de I+D, esta ronda'].sort_values('Ronda_Orden')
            idc = rrhh[rrhh['Metrica'] == 'Otros costos variables de I + D'].sort_values('Ronda_Orden')
            if not idn.empty and not idc.empty:
                chart_dos_metricas_apiladas('Inversión en I+D: costo vs. dotación',
                                             idc['Ronda'], idc['Valor'], 'Costo variable I+D (USD)', COLOR_METRICA['dinero'], 'bar',
                                             idn['Ronda'], idn['Valor'], 'Personal I+D (headcount)', COLOR_METRICA['personas'], 'line')
            else:
                st.info('Sin datos de I+D para este equipo.')
        with col_d:
            st.subheader('Capacitación e impacto')
            capac = rrhh[rrhh['Metrica'] == 'Presupuesto mensual para capacitación, USD'].sort_values('Ronda_Orden')
            efic = rrhh[rrhh['Metrica'] == 'Multiplicador de la eficiencia de RRHH'].sort_values('Ronda_Orden')
            if not capac.empty and not efic.empty:
                chart_dos_metricas_apiladas('Capacitación vs. eficiencia de RRHH',
                                             capac['Ronda'], capac['Valor'], 'Presupuesto capacitación (USD)', COLOR_METRICA['dinero'], 'bar',
                                             efic['Ronda'], efic['Valor'], 'Multiplicador eficiencia RRHH', COLOR_METRICA['eficiencia'], 'line')
                st.caption('Ojo con leer una relación directa: el efecto de la capacitación no es instantáneo, así que el '
                           'presupuesto de una ronda y el multiplicador de esa MISMA ronda no se explican entre sí. '
                           'Lo que hay que mirar es la pendiente del multiplicador en las rondas siguientes a un aumento de presupuesto.')
            else:
                st.info('Sin datos de capacitación para este equipo.')

    with bloque_sost:
        st.subheader('Impacto ambiental')
        c1, c2 = st.columns(2)
        pais_esg = c1.selectbox('País', ['EE.UU.', 'China'], key='pais_esg_rrhh')
        ind = c2.selectbox('Indicador', ['Emisiones de CO2', 'Consumo de energía', 'Consumo de agua'], key='ind_esg_rrhh')
        dicc = {'Emisiones de CO2': 'Total, toneladas métricas', 'Consumo de energía': 'Total, MWh', 'Consumo de agua': 'Total, miles de m3'}
        sub_amb = df[(df['Estado'] == 'Informe ESG') & (df['Seccion'] == f'Impacto ambiental, {pais_esg}') & (df['Metrica'] == dicc[ind])]
        chart_comparacion_equipos(sub_amb, f'{ind} — {pais_esg}')

        st.divider()
        # Adenda 22 (a pedido del equipo): reemplaza el scatter "Reputación ESG y cuota de mercado"
        # (un único score agregado por País vs. Share) por dos vistas más ricas construidas sobre
        # 'Informe ESG' -> Sección 'Impactos en la demanda', que trae el detalle línea por línea de
        # CADA variable Ambiental/Social/Gobernanza y su % de impacto sobre la demanda -- filtrando
        # SIEMPRE Subgrupo != 'Puntuación final' porque esas 4 filas (Ambiental (E), Social (S),
        # Gobernanza (G), Reputación ESG) son un score en una escala absoluta (~2,5-3,7) que NO es la
        # suma de los % de abajo (verificado: para CADIZ Ronda 1 los 5 ítems de Ambiental (E) suman
        # -11,8%, muy distinto del score 2,52 de 'Ambiental (E)' en Puntuación final) -- mezclar
        # ambas unidades en el mismo gráfico sería engañoso.
        def _pct(series):
            return pd.to_numeric(series.astype(str).str.rstrip('%'), errors='coerce')

        imp_esg = df[(df['Estado'] == 'Informe ESG') & (df['Seccion'] == 'Impactos en la demanda') &
                     (df['Subgrupo'] != 'Puntuación final') & (df['Ronda'] == ronda_snapshot)][['Empresa', 'Subgrupo', 'Metrica', 'Valor']].copy()
        imp_esg['Valor'] = _pct(imp_esg['Valor'])
        imp_esg = imp_esg.dropna(subset=['Valor'])

        st.subheader('ESG Impact Heatmap')
        st.caption('% de impacto de cada variable ESG sobre la demanda, ronda ' + str(ronda_snapshot) +
                   ', los 7 equipos. Verde = impacto positivo sobre la demanda, rojo = negativo. Filas agrupadas '
                   'por Ambiental (E) / Social (S) / Gobernanza (G) -- NO incluye los scores agregados de '
                   '"Puntuación final" (otra escala, ver nota).')
        if not imp_esg.empty:
            orden_metricas = (imp_esg[['Subgrupo', 'Metrica']].drop_duplicates()
                               .sort_values(['Subgrupo', 'Metrica'])['Metrica'].tolist())
            piv_heat = imp_esg.pivot_table(index='Metrica', columns='Empresa', values='Valor', aggfunc='mean')
            piv_heat = piv_heat.reindex(orden_metricas)
            max_abs = max(1.0, float(piv_heat.abs().max().max()) if not piv_heat.empty else 1.0)
            fig_heat = px.imshow(piv_heat, color_continuous_scale='RdYlGn', color_continuous_midpoint=0,
                                  zmin=-max_abs, zmax=max_abs, text_auto='.1f', aspect='auto',
                                  labels=dict(color='% impacto'))
            fig_heat.update_xaxes(side='top')
            fig_heat.update_layout(title=f'Impactos en la demanda, % — Ronda {ronda_snapshot}', height=420)
            mostrar(fig_heat)
        else:
            st.info('Sin datos de "Impactos en la demanda" para esta ronda.')

        st.divider()
        st.subheader('Balance Neto ESG')
        st.caption('Suma de los % de impacto sobre la demanda de cada bloque (Ambiental, Social, Gobernanza) por equipo '
                   '-- el total de la barra es el Impacto Neto en la demanda de todo el informe ESG, no un promedio.')
        if not imp_esg.empty:
            balance = imp_esg.groupby(['Empresa', 'Subgrupo'], as_index=False)['Valor'].sum()
            netos = balance.groupby('Empresa', as_index=False)['Valor'].sum().rename(columns={'Valor': 'Neto'})
            orden_emp = netos.sort_values('Neto', ascending=False)['Empresa'].tolist()
            # 3 categorías genéricas (E/S/G, para los 7 equipos) -- sin identidad ni sentimiento,
            # tonos muted (Adenda 25; el marrón que tenía "Social (S)" se saca en la misma Adenda,
            # sin pasar a lila porque tampoco es identidad de CADIZ).
            # Fix "ya casi": "Social (S)" en MUTED_PALETTE[5] quedaba casi idéntico a "Gobernanza (G)"
            # en MUTED_PALETTE[0] (mismo gráfico, barras apiladas) -- se reemplaza por el rosa apagado
            # nuevo, bien separado de los otros dos tonos.
            colores_subgrupo = {'Ambiental (E)': MUTED_PALETTE[2],
                                 'Social (S)': '#B67C8F',
                                 'Gobernanza (G)': MUTED_PALETTE[0]}
            fig_bal = px.bar(balance, x='Empresa', y='Valor', color='Subgrupo', barmode='relative',
                              category_orders={'Empresa': orden_emp},
                              color_discrete_map=colores_subgrupo,
                              title=f'Balance Neto ESG (Ambiental + Social + Gobernanza) — Ronda {ronda_snapshot}')
            fig_bal.update_layout(yaxis_title='% impacto en la demanda (suma)')
            fig_bal.add_trace(go.Scatter(x=netos['Empresa'], y=netos['Neto'], mode='markers+text',
                                          text=netos['Neto'].apply(lambda v: f'{v:+.1f}%'), textposition='top center',
                                          marker=dict(size=1, color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'))
            mostrar(fig_bal)
        else:
            st.info('Sin datos de "Impactos en la demanda" para esta ronda.')
# ---------------- Router ----------------
st.title(seccion)
if seccion == SECCIONES[0]: seccion_resultado()
elif seccion == SECCIONES[1]: seccion_mercado()
elif seccion == SECCIONES[2]: seccion_operaciones()
elif seccion == SECCIONES[3]: seccion_finanzas()
elif seccion == SECCIONES[4]: seccion_rrhh_sostenibilidad()
