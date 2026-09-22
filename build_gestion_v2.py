# -*- coding: utf-8 -*-
"""
Construye /home/claude/cadiz_gestion/CADIZ_Gestion_v1.xlsx desde cero.
Ver Informe_Construccion_CADIZ_Gestion_v1.md y CADIZ_Gestion_v1_README.md para el detalle de
arquitectura, simplificaciones y supuestos. Este script NUNCA lee ni escribe MOTOR_VALIDADO_R2.xlsx
(fuente de verdad de solo lectura) -- los valores semilla (historicos R0/R1, catalogo D1-D9, Escenario
A R2) se copian a mano abajo, ya extraidos y verificados por lectura separada del motor viejo.

Convencion de columnas (igual en TODAS las hojas de datos por ronda): columna G = Ronda 0,
columna H = Ronda 1, ..., columna S = Ronda 12 (col = 7+ronda), igual criterio que el motor anterior.
"""
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

import gestion_styles as S

# ======================================================================================
# Cambio 4 (v1.1) -- vocabulario corto y estandarizado para la columna "Fuente" de 01_INPUTS.
# Clasifica el texto largo original de "nota" (heredado de v1) en una etiqueta corta; el texto
# completo se conserva como comentario de Excel en la misma celda (ver put() en build_inputs).
# No cambia ningún valor ni fórmula -- es puramente presentación de 01_INPUTS.
# ======================================================================================
def _classify_fuente(tipo, nota):
    nota = (nota or "").strip()
    if not nota:
        return "", None
    nl = nota.lower()
    m = re.search(r"manual\s+cap\.?\s*([0-9]+(?:\.[0-9]+)?)", nl)
    if "no determinado" in nl or "ambigüedad no resuelta" in nl or "no se encontró" in nl:
        short = "No determinado"
    elif "no verificado en manual" in nl or "no verificado" in nl:
        short = "No determinado"
    elif nl.startswith("alerta") or "[alerta" in nl:
        short = "Alerta modelo"
    elif "rdos r1" in nl or "rdos r0" in nl or "verificado exacto" in nl or "calibrado contra rdos" in nl:
        short = "RDOS R1"
    elif "contexto - ronda 2" in nl or "condición oficial" in nl or "cesim" in nl:
        short = "CESIM R2"
    elif "simplificación" in nl or "[simplificación" in nl:
        short = "Supuesto propio"
    elif m:
        short = f"Manual cap. {m.group(1)}"
    elif nl.startswith("nota:"):
        short = "Nota modelo"
    elif "sin costo/interés" in nl or "definición cerrada del equipo" in nl:
        short = "Supuesto propio"
    else:
        short = "Ver comentario"
    return short, nota

PAISES = ["EE.UU.", "China", "Europa"]
AREAS = ["EE.UU.", "China"]          # areas productoras
MERCADOS = ["EE.UU.", "China", "Europa"]
TECNOLOGIAS = ["Combustión", "Híbrido", "Eléctrico", "Hidrógeno"]
TECH_ACTIVAS = ["Combustión", "Híbrido"]   # unicas comercializadas por CADIZ a R0-R2
RONDAS = list(range(0, 13))
ENFOQUES_MKT = ["Equilibrado", "Precio bajo", "Características", "Sostenibilidad", "Marca"]
MONEDA_MERCADO = {"EE.UU.": "USD", "China": "RMB", "Europa": "EUR"}

INPUTS_SHEET = "01_INPUTS"
ESTADOS_SHEET = "02_ESTADOS_PROYECTADOS"
RATIOS_SHEET = "03_RATIOS"
CONTROL_SHEET = "04_CONTROL_MODELO"
HIST_SHEET = "05_HISTORICO_EQUIPOS"
ENGINE_MKT_SHEET = "_ENGINE_MERCADO"
ENGINE_PROD_SHEET = "_ENGINE_PRODUCCION"
ENGINE_FIN_SHEET = "_ENGINE_FINANCIERO"
DATA_EXPORT_SHEET = "DATA_EXPORT"

OUT_PATH = "/home/claude/cadiz_gestion/CADIZ_Gestion_v3.xlsx"

# Fase 3: los históricos por país (P&L/Balance/CF) y el histórico tidy de 7 equipos ya NO se leen
# de JSON pre-parseados ni de MOTOR_VALIDADO_R2.xlsx -- se parsean al vuelo desde los RDOS de cada
# ronda jugada (rdos_files, ver extract_hist_pais()/build_historico_equipos() y generar_excel()),
# vía el parser único rdos_parser.py. Dato histórico real, no lógica propia; una sola fuente/lógica
# de escala para todas las rondas.
import xlrd
import rdos_parser
from rdos_parser import parse_rdos_workbook, detectar_ronda_desde_nombre


def extract_hist_pais(rdos_files):
    """RDOS por país (EE.UU./China/Europa) -- P&L, Balance y Flujo de Fondos reales, en USD.
    rdos_files: dict {ronda_int: path_o_fileobj_del_RDOS} para TODAS las rondas ya jugadas (antes
    solo R0/R1 desde ronda0.json/ronda1.json; ahora cualquier ronda, parseada al vuelo con
    rdos_parser.parse_rdos_workbook() -- una sola fuente/lógica de escala para todas las rondas,
    reemplaza a _load_json_build()). Usa el valor 'scaled' de cada fila (RDOS está en miles --
    x1000 verificado bottom-up: precio de venta x volumen vendido x tipo de cambio, sumado en
    las 3 áreas, coincide con el total de "Ingresos por ventas" con 0,003% de diferencia). Dato
    histórico real, no supuesto ni lógica propia."""
    out = {rn: {} for rn in rdos_files}
    cf_labels = {"Flujo de efectivo de casa matriz, miles USD": "EE.UU.",
                 "Estado de flujo de efectivo, miles USD, China": "China",
                 "Estado de flujo de efectivo, miles USD, Europa": "Europa"}
    for rn, f in rdos_files.items():
        rows = parse_rdos_workbook(f)
        cf_seccion_actual = {}
        for row in rows:
            sec, lbl = row["section"], row["label"]
            if row["type"] == "subheader":
                if sec in cf_labels:
                    cf_seccion_actual[sec] = lbl
                continue
            for pais in PAISES + ["Global"]:
                if sec == f"Cuenta de resultados, miles USD, {pais}":
                    out[rn].setdefault(pais, {}).setdefault("pl", {})[lbl] = row["scaled"].get("CADIZ")
                if sec == f"Hoja de Balance, miles USD, {pais}":
                    out[rn].setdefault(pais, {}).setdefault("bal", {})[lbl] = row["scaled"].get("CADIZ")
            if sec in cf_labels:
                lbl_out = lbl
                if lbl == "Total" and cf_seccion_actual.get(sec):
                    lbl_out = f"Total ({cf_seccion_actual[sec]})"
                out[rn].setdefault(cf_labels[sec], {}).setdefault("cf", {})[lbl_out] = row["scaled"].get("CADIZ")
            # DATO HISTÓRICO REAL (verificado, ver Informe Fase 4): "Acciones en circulación al final
            # de la ronda, m acciones" -- CESIM la publica dentro de la sección de flujo de efectivo de
            # Europa (quirk del propio RDOS, es un dato GLOBAL de casa matriz, no de ese país) y con la
            # unidad abreviada "m acciones" (miles de acciones) en vez de "miles unidades" -- rdos_parser
            # NO reconoce esa abreviatura (_MILES_RE exige la palabra completa "miles"), así que esta
            # fila llega en "values" (crudo, sin escalar) y hace falta el x1000 a mano acá. Verificado
            # contra R0-R3: da 1.625.000.000 acciones en las 4 rondas (CADIZ no emitió/recompró todavía),
            # consistente con el valor que antes estaba hardcodeado como "DATO HISTÓRICO REAL" para R0/R1.
            if sec == "Estado de flujo de efectivo, miles USD, Europa" and lbl == "Acciones en circulación al final de la ronda, m acciones":
                v_nacc = row["values"].get("CADIZ")
                if isinstance(v_nacc, (int, float)):
                    out[rn]["_nacc"] = v_nacc * 1000
    return out


def extract_mercado_real(rdos_files):
    """Tamaño de mercado total y Demanda estimada CADIZ, por mercado/tecnología, para CUALQUIER
    ronda con RDOS (reemplaza los diccionarios tam_hist/dem_real de _ENGINE_MERCADO, que solo
    cubrían R0/R1 a mano). Dato histórico real, no supuesto ni lógica propia.

    Verificado bottom-up contra R0/R1 (0 diferencia a 10 decimales en Tamaño de mercado, 0
    mismatches en las 24 combinaciones mercado×tecnología de Demanda) y contra R2/R3 (valores
    nuevos, plausibles, primera vez que se calculan) -- ver test_mercado_extractor.py / Informe
    Fase 4.

    tam[mercado][ronda] = Σ("Ventas, miles unidades" x1000, todas las tecnologías, CADIZ, sección
        "Informe de mercado, {mercado}") / ("Total" de "{mercado} cuotas de mercado, %", CADIZ, /100).
    dem[(mercado, tecnologia)][ronda] = "Demanda, miles unidades" x1000 de CADIZ, dentro del bloque
        de esa tecnología en "Informe de mercado, {mercado}" -- 0.0 si CADIZ no tiene fila para esa
        tecnología (no se inventa un valor).
    """
    tam, dem_out = {}, {}
    secciones_informe = {f"Informe de mercado, {m}": m for m in MERCADOS}
    for rn, f in rdos_files.items():
        rows = parse_rdos_workbook(f)
        ventas_por_mercado = {m: 0.0 for m in MERCADOS}
        cuota_por_mercado = {m: None for m in MERCADOS}
        demanda_por_mercado_tech = {}
        current_mercado, current_tech = None, None
        for row in rows:
            sec, lbl, typ = row["section"], row["label"], row["type"]
            if sec in secciones_informe:
                current_mercado = secciones_informe[sec]
                if typ == "subheader" and lbl in TECNOLOGIAS:
                    current_tech = lbl
                elif typ == "data" and lbl == "Demanda, miles unidades" and current_tech:
                    demanda_por_mercado_tech[(current_mercado, current_tech)] = row["scaled"].get("CADIZ")
                elif typ == "data" and lbl == "Ventas, miles unidades" and current_mercado:
                    v = row["scaled"].get("CADIZ")
                    if isinstance(v, (int, float)):
                        ventas_por_mercado[current_mercado] += v
            elif sec and sec.endswith("cuotas de mercado, %") and sec != "Cuotas de mercado globales, %":
                m = sec.replace(" cuotas de mercado, %", "")
                if typ == "data" and lbl == "Total" and m in MERCADOS:
                    v = row["values"].get("CADIZ")
                    if isinstance(v, (int, float)):
                        cuota_por_mercado[m] = v / 100.0
        for m in MERCADOS:
            tam.setdefault(m, {})[rn] = (ventas_por_mercado[m] / cuota_por_mercado[m]) if cuota_por_mercado[m] else None
        for tech in TECNOLOGIAS:
            for m in MERCADOS:
                dem_out.setdefault((m, tech), {})[rn] = demanda_por_mercado_tech.get((m, tech), 0.0)
    return tam, dem_out


# Hojas cuyo NOMBRE empieza con un dígito -- por gramática OOXML, cualquier referencia de fórmula
# a estas hojas DEBE ir entre comillas simples ('01_INPUTS'!A1), o el archivo es inválido para un
# parser estricto (Excel real). openpyxl ya escribe SIEMPRE con comillas (ver cref()/iref() abajo) --
# verificado por auditoría de build_gestion.py: no existe ningún lugar del script que arme una
# referencia de fórmula a estas hojas sin pasar por cref()/iref(). El bug real de v1 NO está en este
# script: aparece DESPUÉS, cuando recalc.py (LibreOffice) recalcula y re-guarda el .xlsx -- el
# exportador OOXML de LibreOffice no aplica la misma regla de quoting que Excel para nombres de hoja
# que empiezan con dígito, y al reescribir el archivo elimina las comillas de estas 5 hojas
# específicamente (reproducido y confirmado con un caso mínimo). Por eso hace falta el paso de
# postproceso fix_digit_sheet_quoting() (ver más abajo), que se corre UNA sola vez, DESPUÉS del único
# recalc.py final, directamente sobre el XML interno del .xlsx ya recalculado -- no es un parche
# manual "a ojo": es una reinserción quirúrgica y determinística de comillas alrededor de estos 5
# nombres de hoja exactos, dentro de <f> (fórmulas) y definedNames únicamente, sin tocar <v>
# (valores cacheados) ni sharedStrings.xml. Ver Informe_V1_1.md para el detalle y la reproducción.
DIGIT_LEAD_SHEETS = ["01_INPUTS", "02_ESTADOS_PROYECTADOS", "03_RATIOS", "04_CONTROL_MODELO", "05_HISTORICO_EQUIPOS"]


def col(ronda):
    return get_column_letter(7 + ronda)


def cref(sheet, row, ronda):
    """Referencia entre hojas: 'sheet'!<col><row>. SIEMPRE con comillas simples (obligatorio para
    hojas cuyo nombre empieza con dígito; inofensivo para el resto -- Excel acepta comillas en
    cualquier nombre de hoja)."""
    return f"'{sheet}'!{col(ronda)}{row}"


def fifo_layer_terms(row_prod_propia, row_costo_propia, row_terc_u, row_costo_terc, row_ventas_u,
                      heritage_qty, heritage_cost, r_min, r_max):
    """FIFO multicapa (Fase 2, cambio 1) -- replica la lógica de MOTOR_VALIDADO_R2 / build_engine_e2.py
    ::layer_valor_formula() (ver docstring allí): cada capa (ronda de origen × propia/tercerizada)
    conserva su costo unitario de origen hasta agotarse, consumida en orden estrictamente FIFO
    (más antigua primero; dentro de la misma ronda, propia antes que tercerizada). Devuelve el valor
    del INVENTARIO FINAL remanente (a costo de origen, sin fusionar capas) al cierre de r_max.

    A diferencia del motor viejo (que arranca en Ronda 0), este workbook usa R0/R1 como HISTÓRICO
    REAL migrado (no recalculado -- ver Sección C, semillas 'seed()'), así que la cadena de capas
    FIFO arranca en una única "capa heredada" sintética = el inventario final REAL de Ronda 1
    (heritage_qty unidades a heritage_cost USD/u., ambos ya conocidos y auditados), seguida por
    las rondas r_min..r_max (r_min=2) con sus propias sub-capas propia/tercerizada, exactamente
    igual que el motor viejo a partir de ahí. cumsold solo cuenta ventas de r_min..r_max porque las
    ventas de R0/R1 ya están netas dentro de heritage_qty (viene del inventario final REAL de R1)."""
    terms = []
    offset_terms = [f"{heritage_qty}"]
    cumsold = f"SUM({col(r_min)}{row_ventas_u}:{col(r_max)}{row_ventas_u})"
    terms.append(f"MAX(0,MIN({heritage_qty},{heritage_qty}-{cumsold}))*{heritage_cost}")
    for v in range(r_min, r_max + 1):
        cv = col(v)
        size_p = f"{cv}{row_prod_propia}"
        cost_p = f"{cv}{row_costo_propia}"
        offset_expr = "+".join(offset_terms)
        terms.append(f"MAX(0,MIN({size_p},{offset_expr}+{size_p}-{cumsold}))*{cost_p}")
        offset_terms.append(size_p)
        size_t = f"{cv}{row_terc_u}"
        cost_t = f"{cv}{row_costo_terc}"
        offset_expr = "+".join(offset_terms)
        terms.append(f"MAX(0,MIN({size_t},{offset_expr}+{size_t}-{cumsold}))*{cost_t}")
        offset_terms.append(size_t)
    return "+".join(terms)


def fix_digit_sheet_quoting(path, sheet_names=None):
    """POSTPROCESO OBLIGATORIO -- correr UNA sola vez, después del recalc.py final (nunca antes, y
    nunca volver a correr recalc.py después de este paso: LibreOffice volvería a quitar las
    comillas). Reinserta comillas simples alrededor de los nombres de hoja de sheet_names en:
      - los textos de fórmula (<f>...</f>) de xl/worksheets/*.xml
      - los rangos de xl/workbook.xml (<definedName>...)
    No toca <v> (valores cacheados: se preservan intactos) ni xl/sharedStrings.xml (texto de
    celdas/comentarios que pueda mencionar el nombre de hoja como texto, no como referencia)."""
    import zipfile
    import shutil
    import re as _re
    sheet_names = sheet_names or DIGIT_LEAD_SHEETS
    tmp_path = path + ".tmp_fixquote"
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        targets = [n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")] + ["xl/workbook.xml"]
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name in targets:
                    text = data.decode("utf-8")
                    for sheet in sheet_names:
                        # Reemplaza SHEETNAME! por 'SHEETNAME'! salvo que ya esté precedido de comilla.
                        pattern = _re.compile(r"(?<!')" + _re.escape(sheet) + r"!")
                        text = pattern.sub(f"'{sheet}'!", text)
                    data = text.encode("utf-8")
                zout.writestr(name, data)
    shutil.move(tmp_path, path)


# ======================================================================================
# 01_INPUTS
# ======================================================================================
def build_inputs(wb, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(INPUTS_SHEET)
    ws.sheet_view.showGridLines = False
    ncols = 6 + len(RONDAS)
    S.set_col_widths(ws, [16, 42, 22, 12, 16, 8] + [13] * len(RONDAS))
    S.style_title_row(ws, 1, "01_INPUTS — Única interfaz de carga manual: Identificación, Condiciones de Ronda y Decisiones D1–D8", ncols)
    ws.row_dimensions[1].height = 26
    ws.cell(row=2, column=1,
            value=("Azul = celda editable (decisión CADIZ o condición de ronda publicada por CESIM). Negro = fórmula de arrastre (edítese solo si CESIM publica un valor nuevo). "
                   "Gris = dato histórico real (Ronda 0/1, no editar). Para agregar una ronda nueva: completar la columna de esa ronda; el resto de las hojas ya está preparado hasta la Ronda 12."))
    ws.cell(row=2, column=1).font = S.f_normal(9, italic=True, color=S.COLOR_HIST)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.row_dimensions[2].height = 28

    header_row = 4
    headers = ["Bloque", "Variable", "Dimensión", "Unidad", "Tipo", "Fuente"] + [f"Ronda {r}" for r in RONDAS]
    S.style_header_row(ws, header_row, headers)
    ws.freeze_panes = "G5"
    r = [header_row + 1]
    row_idx = {}

    def put(bloque, variable, dimension, unidad, tipo, nota=""):
        rr = r[0]
        row_idx[f"{bloque}||{variable}||{dimension}"] = rr
        S.apply_cell(ws, rr, 1, value=bloque, kind="calculo", align=S.ALIGN_LEFT, bold=True, size=9)
        S.apply_cell(ws, rr, 2, value=variable, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=9)
        S.apply_cell(ws, rr, 3, value=dimension, kind="calculo", align=S.ALIGN_LEFT, size=9)
        S.apply_cell(ws, rr, 4, value=unidad, kind="calculo", align=S.ALIGN_CENTER, size=8)
        S.apply_cell(ws, rr, 5, value=tipo, kind="calculo", align=S.ALIGN_CENTER, size=8, bold=True)
        # Cambio 4 (v1.1): la columna "Fuente" muestra SOLO una etiqueta corta y estandarizada
        # (vocabulario fijo, ver README); el texto largo original -- si lo había -- se preserva
        # ÍNTEGRO como comentario de Excel de la misma celda (no se pierde trazabilidad, solo deja
        # de ocupar espacio visible en la grilla, que era la causa de las filas altísimas de v1).
        short_fuente, full_note = _classify_fuente(tipo, nota)
        c = S.apply_cell(ws, rr, 6, value=short_fuente, kind="calculo", align=S.ALIGN_LEFT, size=8, italic=True)
        if full_note:
            cm = Comment(full_note, "CADIZ Gestión v1.1")
            cm.width, cm.height = 340, 180
            c.comment = cm
        r[0] += 1
        return rr

    def section(title):
        S.style_section_row(ws, r[0], title, ncols)
        r[0] += 1

    def fill_row(row, values, kinds, numfmt=None, align=S.ALIGN_RIGHT):
        for rn in RONDAS:
            v = values.get(rn, None)
            k = kinds.get(rn, "input")
            if v is None:
                S.apply_cell(ws, row, 7 + rn, value=None, kind=k, numfmt=numfmt, align=align, size=9)
            else:
                S.apply_cell(ws, row, 7 + rn, value=v, kind=k, numfmt=numfmt, align=align, size=9)

    def arrastre_rows(row, from_round, to_round=12, numfmt=None):
        """Rondas from_round+1..to_round = '=celda anterior' (arrastre visible, editable)."""
        for rn in range(from_round + 1, to_round + 1):
            f = f"={col(rn-1)}{row}"
            S.apply_cell(ws, row, 7 + rn, value=f, kind="calculo", numfmt=numfmt, align=S.ALIGN_RIGHT, size=9)

    NA_PLAIN = {"kind": "plain"}

    # ---------------------------------------------------------------------------------
    # SECCIÓN [SUPUESTO PROPIO] · MATRIZ DE SENSIBILIDAD (Fase 2, cambio 4) -- herramienta de
    # planificación propia del equipo CADIZ, NO es una regla CESIM. Un único "Escenario Activo"
    # (celda escalar, no por ronda) multiplica el crecimiento de mercado (_ENGINE_MERCADO Sección
    # A1) y la cuota efectiva de CADIZ (Sección A2) en TODAS las rondas simultáneamente. El
    # escenario BASE tiene factores en 0% por diseño (no negociable, ver instrucción del usuario) --
    # con Escenario Activo=BASE, el multiplicador es exactamente 1 y Ronda 2 (Escenario A, ya
    # jugada) no cambia. Los valores Pesimista/Optimista son EDITABLES (celda azul) y sirven para
    # planificar R3 en adelante.
    # ---------------------------------------------------------------------------------
    # Hotfix regional (post-Fase 2): esta sección NO debe pisar las columnas de ronda (G:S) -- ni
    # con valores ni con el selector "Escenario Activo" -- porque rompe la lectura vertical de la
    # grilla (esas columnas son Ronda 0..12 en el resto de la hoja).
    #
    # Hotfix 2 (verificación independiente post-hotfix-1): con Pesimista/Base/Optimista en C/D/E,
    # "Optimista" (col. E) quedaba oculto -- E es la columna "Tipo", oculta en toda la hoja por el
    # hide de metadatos. Solución: solo B (Variable, con el nombre completo de la variable -- no se
    # puede reemplazar por un valor, es la única forma de saber qué fila es cada una), C y D quedan
    # visibles en esta hoja; por lo tanto los 3 valores Pesimista/Base/Optimista NO entran completos
    # en columnas visibles sin sacrificar la etiqueta. Como "Base" es fija en 0% por diseño (no
    # negociable, no se edita en la operación normal), se ubica en la columna oculta E -- ocultarla
    # no le quita nada, es justamente la que no hay que tocar. Pesimista (editable) y Optimista
    # (editable) quedan en C y D, ambas visibles. El selector "Escenario Activo" y los 2 factores
    # calculados son escalares únicos y reutilizan la columna C (en sus propias filas, sin
    # colisión con Pesimista de las filas 7-8).
    SENS_COL_PESIM, SENS_COL_OPTIM, SENS_COL_BASE = 3, 4, 5   # C (visible), D (visible), E (oculta)
    SENS_COL_SCALAR = SENS_COL_PESIM                          # C -- Escenario Activo y factores calculados
    LIST_COL = 21  # U -- columna oculta de listas para Data Validation / MATCH (fix de separador regional)

    section("SECCIÓN [SUPUESTO PROPIO] · MATRIZ DE SENSIBILIDAD (planificación propia, NO es una regla CESIM — no afecta Ronda 2 con Escenario Activo=BASE)")
    row_sens_label = r[0]
    S.apply_cell(ws, row_sens_label, 1, value="Pesimista/Optimista en C/D → (Base fija en 0%, col. E oculta a propósito)", kind="plain", align=S.ALIGN_LEFT, size=8, italic=True)
    S.apply_cell(ws, row_sens_label, SENS_COL_PESIM, value="Pesimista", kind="plain", align=S.ALIGN_CENTER, bold=True, size=8)
    S.apply_cell(ws, row_sens_label, SENS_COL_OPTIM, value="Optimista", kind="plain", align=S.ALIGN_CENTER, bold=True, size=8)
    S.apply_cell(ws, row_sens_label, SENS_COL_BASE, value="Base (fija, oculta)", kind="plain", align=S.ALIGN_CENTER, bold=True, size=8)
    r[0] += 1
    row_sens_crec = put("SENS · Supuesto propio", "Desvío en Crecimiento de Mercado (vs. condición publicada)", "Pesimista / Base / Optimista", "%", "SUPUESTO PROPIO",
                         "[SUPUESTO PROPIO] Herramienta de planificación del equipo, no una regla CESIM. Multiplica el crecimiento de mercado de _ENGINE_MERCADO Sección A1: Demanda_ajustada = Demanda_base × (1+Factor). BASE = 0% por diseño. Aplica ESTRICTAMENTE desde Ronda 3 en adelante -- R0/R1/R2 son histórico cerrado y no se recalculan con este mecanismo.")
    S.apply_cell(ws, row_sens_crec, SENS_COL_PESIM, value=-0.03, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)   # Pesimista, col C (visible)
    S.apply_cell(ws, row_sens_crec, SENS_COL_OPTIM, value=0.03, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)   # Optimista, col D (visible)
    S.apply_cell(ws, row_sens_crec, SENS_COL_BASE, value=0.00, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)    # Base, col E (oculta, fija por diseño)
    row_sens_comp = put("SENS · Supuesto propio", "Agresividad de la Competencia (recorte a la cuota efectiva de CADIZ)", "Pesimista / Base / Optimista", "%", "SUPUESTO PROPIO",
                         "[SUPUESTO PROPIO] Herramienta de planificación del equipo, no una regla CESIM. Recorta la cuota efectiva de CADIZ en _ENGINE_MERCADO Sección A2: Cuota_efectiva = Cuota_objetivo × (1−Factor). BASE = 0% por diseño. Aplica ESTRICTAMENTE desde Ronda 3 en adelante -- R0/R1/R2 son histórico cerrado y no se recalculan con este mecanismo.")
    S.apply_cell(ws, row_sens_comp, SENS_COL_PESIM, value=0.05, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_sens_comp, SENS_COL_OPTIM, value=0.00, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_sens_comp, SENS_COL_BASE, value=0.00, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)

    # Lista oculta para Data Validation y MATCH (columna U) -- reemplaza la cadena de texto separada
    # por comas (rompía en Excel es-AR/es-LA, donde el separador de listas es ";") y el array en
    # línea {"PESIMISTA","BASE","OPTIMISTA"} (rompe igual por el separador de matrices regional).
    # Orden PESIMISTA/OPTIMISTA/BASE (no alfabético): debe coincidir con el orden físico de columnas
    # C/D/E de arriba para que INDEX(C:E,MATCH(...)) resuelva la posición correcta -- el orden en que
    # aparecen las opciones en el desplegable de "Escenario Activo" es cosmético, no afecta la lógica.
    S.apply_cell(ws, 2, LIST_COL, value="PESIMISTA", kind="plain", size=8)
    S.apply_cell(ws, 3, LIST_COL, value="OPTIMISTA", kind="plain", size=8)
    S.apply_cell(ws, 4, LIST_COL, value="BASE", kind="plain", size=8)
    list_col_letter = get_column_letter(LIST_COL)
    sens_list_range = f"${list_col_letter}$2:${list_col_letter}$4"

    row_sens_activo = put("SENS · Supuesto propio", "Escenario Activo", "-", "texto", "DECISIÓN CADIZ (planificación)",
                           "[NO NEGOCIABLE] Aplica ESTRICTAMENTE desde Ronda 3 en adelante -- R0/R1/R2 son histórico cerrado, no se recalculan con este mecanismo pase lo que pase acá. Cambiar a PESIMISTA/OPTIMISTA para planificar R3+ (afecta TODAS las rondas activas por igual desde R3, es una celda única, no por ronda).")
    S.apply_cell(ws, row_sens_activo, SENS_COL_SCALAR, value="BASE", kind="input", align=S.ALIGN_CENTER, bold=True, size=9)
    dv_sens = DataValidation(type="list", formula1=sens_list_range, allow_blank=False)
    ws.add_data_validation(dv_sens)
    dv_sens.add(f"{get_column_letter(SENS_COL_SCALAR)}{row_sens_activo}")

    row_sens_factor_crec = put("SENS · Supuesto propio", "Factor de Crecimiento Activo (calculado)", "-", "%", "CÁLCULO",
                                "=INDEX(fila Pesimista/Optimista/Base de Crecimiento, MATCH(Escenario Activo)). Alimenta _ENGINE_MERCADO Sección A1, solo Ronda 3+.")
    f_fc = f'=INDEX({get_column_letter(SENS_COL_PESIM)}{row_sens_crec}:{get_column_letter(SENS_COL_BASE)}{row_sens_crec},MATCH({sens_ref(row_sens_activo)},{sens_list_range},0))'
    S.apply_cell(ws, row_sens_factor_crec, SENS_COL_SCALAR, value=f_fc, kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, bold=True, size=9)
    row_sens_factor_comp = put("SENS · Supuesto propio", "Factor de Competencia Activo (calculado)", "-", "%", "CÁLCULO",
                                "=INDEX(fila Pesimista/Optimista/Base de Competencia, MATCH(Escenario Activo)). Alimenta _ENGINE_MERCADO Sección A2, solo Ronda 3+.")
    f_fk = f'=INDEX({get_column_letter(SENS_COL_PESIM)}{row_sens_comp}:{get_column_letter(SENS_COL_BASE)}{row_sens_comp},MATCH({sens_ref(row_sens_activo)},{sens_list_range},0))'
    S.apply_cell(ws, row_sens_factor_comp, SENS_COL_SCALAR, value=f_fk, kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, bold=True, size=9)
    r[0] += 1
    sens_rows = dict(crec=row_sens_crec, comp=row_sens_comp, activo=row_sens_activo,
                      factor_crec=row_sens_factor_crec, factor_comp=row_sens_factor_comp)

    # ---------------------------------------------------------------------------------
    # SECCIÓN A · IDENTIFICACIÓN
    # ---------------------------------------------------------------------------------
    section("SECCIÓN A · IDENTIFICACIÓN")
    row_ronda = put("A. Identificación", "Número de ronda", "-", "n°", "CÁLCULO")
    for rn in RONDAS:
        S.apply_cell(ws, row_ronda, 7 + rn, value=rn, kind="calculo", align=S.ALIGN_CENTER, bold=True, size=9)
    row_escenario = put("A. Identificación", "Escenario", "-", "texto", "DECISIÓN CADIZ", "Nombre del escenario de decisión (p.ej. 'A'). n/a en rondas históricas.")
    # CORRECCIÓN (mismo hotfix que Estado, ver comentario abajo): R2 pasa a estilo azul "input" igual
    # que el resto de las rondas PLAN (R3-R12) -- el gris "histórico" quedaba reservado por el propio
    # encabezado de esta hoja (fila 2) a "dato histórico real (Ronda 0/1, no editar)", y usarlo también
    # en R2 sugería visualmente que R2 ya tenía RDOS reales cargados, lo cual es falso.
    esc_vals = {0: "n/a (histórico)", 1: "n/a (histórico)", 2: "A"}
    for rn in RONDAS:
        v = esc_vals.get(rn)
        S.apply_cell(ws, row_escenario, 7 + rn, value=v, kind=("plain" if rn < 2 else "input"), align=S.ALIGN_CENTER, size=9, italic=(v is None or rn < 2))
    row_estado = put("A. Identificación", "Estado", "-", "REAL/PLAN/SIM", "DECISIÓN CADIZ", "REAL = ronda ya jugada, con RDOS oficiales de CESIM cargados. PLAN = decisión cargada/enviada, aún sin RDOS oficiales confirmados. SIM = corrida de sensibilidad, no es el Plan oficial.")
    # CORRECCIÓN (hotfix integración web, confirmado con el usuario): la versión anterior de este
    # script marcaba Estado(R2)="REAL" para señalar que la DECISIÓN de R2 (Escenario A) ya estaba
    # cerrada/enviada al simulador -- pero esa lectura de "REAL" mezclaba dos cosas distintas del
    # proyecto: (2) Dato histórico real (RDOS publicados por CESIM) vs (4) Decisión/objetivo de CADIZ.
    # Los valores de R2 en este archivo NUNCA fueron una migración de RDOS publicados -- siguen siendo
    # la PROYECCIÓN del modelo Fase 2 (validada contra MOTOR_VALIDADO_R2, ver Informe_V2.md). Al
    # integrar este workbook con el tablero web (control de gestión Proyectado vs. Real), un
    # Estado="REAL" en R2 se exporta tal cual a DATA_EXPORT!status y la web lo toma como si Cesim ya
    # hubiese publicado el resultado real de R2 -- anulando exactamente el gap que se quiere medir.
    # Corregido: Estado(R2)="PLAN" (decisión ya enviada/cerrada, pero sin RDOS oficiales todavía).
    # Cuando el profesor publique los RDOS reales de R2, ese dato se migra aparte (mismo proceso que
    # R0/R1) y recién ahí corresponde pasar Estado(R2) a "REAL". R3 en adelante sigue como ronda activa
    # de planificación (Estado=PLAN, arrastra hasta que se confirme en CESIM).
    # Ronda N.N (Fase 4, frontera dinámica): Estado(rn)="REAL" para TODA ronda con RDOS oficial
    # cargado (rondas_reales = frozenset(rdos_files.keys()), calculado en generar_excel() a partir de
    # data/raw/oficial/) -- ya NO un boundary fijo (0,1). La ronda siguiente a la última real
    # (n_next = max(rondas_reales)+1) arranca en "PLAN" como default de bootstrap; de ahí en más queda
    # en blanco (kind="input") para que el usuario la cargue o para que el mecanismo de rescate de
    # decisiones (leer_decisiones_previas/aplicar_overrides_inputs, ver generar_excel) la restaure
    # desde el último Cadiz_proyeccion_R{N}.xlsx sin perder lo ya planificado a futuro.
    n_next_estado = (max(rondas_reales) + 1) if rondas_reales else 0
    estado_vals = {rn: "REAL" for rn in rondas_reales}
    if n_next_estado <= 12:
        estado_vals[n_next_estado] = "PLAN"
    for rn in RONDAS:
        v = estado_vals.get(rn)
        S.apply_cell(ws, row_estado, 7 + rn, value=v, kind=("historico" if rn in rondas_reales else "input"), align=S.ALIGN_CENTER, size=9, bold=True)
    row_ronda_activa = put(
        "A. Identificación", "RONDA_ACTIVA (calculado)", "-", "VERDADERO/FALSO", "CÁLCULO",
        "SUPUESTO DE MODELO (Cambio 2, v1.1): VERDADERO si Estado de esta ronda ∈ {PLAN,SIM,REAL} Y el "
        "input mínimo D1 'Cuota de mercado objetivo SOBRE LA TECNOLOGÍA' (EE.UU./Combustión) está cargado. R0 y R1 "
        "quedan fijos en VERDADERO (históricas, no dependen de esta lógica). Usado por "
        "02_ESTADOS_PROYECTADOS, 03_RATIOS, 04_CONTROL_MODELO y DATA_EXPORT para no calcular/exportar "
        "una ronda con inputs vacíos. El valor se completa más abajo, una vez creada la fila de D1.")
    r[0] += 1
    # (las celdas de row_ronda_activa se completan después de construir la Sección C.D1, porque su
    #  fórmula necesita la fila de "Cuota de mercado objetivo SOBRE LA TECNOLOGÍA" -- referencia hacia adelante
    #  dentro de la misma hoja, válida en Excel/OOXML sin restricción de orden físico de filas.)

    # ---------------------------------------------------------------------------------
    # SECCIÓN B · CONDICIONES DE RONDA (publicadas por CESIM, no son decisiones de CADIZ)
    # ---------------------------------------------------------------------------------
    section("SECCIÓN B · CONDICIONES DE RONDA (valores publicados por CESIM en 'Condiciones' de cada ronda — R3-R12 arrastran el último valor conocido, EDITABLE en cuanto se publique el real)")

    # -- Crecimiento de mercado esperado (D1, condición) --
    growth_hist = {}  # R0/R1: no publicado como condicion previa retroactiva
    growth_r2 = {"EE.UU.": 0.05, "China": 0.085, "Europa": 0.03}
    for mercado in MERCADOS:
        row = put("B. Condiciones", "Crecimiento de mercado esperado (vs. ronda anterior)", mercado, "%", "CONDICIÓN DE RONDA",
                   "Fuente: 'Contexto - Ronda 2.pdf' (China publicada como rango 8-9%; se usa 8,5% como punto medio, documentado).")
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0, sin ronda anterior)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value="n/a (no cargado retroactivamente)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=growth_r2[mercado], kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    # -- Tasa de impuesto corporativo (D7, condición) --
    tasa_hist = {"EE.UU.": {0: 0.22, 1: 0.25}, "China": {0: 0.25, 1: 0.28}, "Europa": {0: 0.21, 1: 0.23}}
    tasa_r2 = {"EE.UU.": 0.22, "China": 0.25, "Europa": 0.21}
    for pais in PAISES:
        row = put("B. Condiciones", "Tasa de impuesto corporativo", pais, "%", "CONDICIÓN DE RONDA", "DATO HISTÓRICO REAL R0/R1 (RDOS); R2 = Condición oficial ('Contexto - Ronda 2.pdf').")
        S.apply_cell(ws, row, 7 + 0, value=tasa_hist[pais][0], kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 1, value=tasa_hist[pais][1], kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=tasa_r2[pais], kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    # -- FX --
    fx_hist = {"China (RMB)": {1: 0.1620}, "Europa (EUR)": {1: 1.1460}}
    fx_r2 = {"China (RMB)": 0.1507, "Europa (EUR)": 1.1126}
    row_fx = {}
    for dim in ("China (RMB)", "Europa (EUR)"):
        row = put("B. Condiciones", "Tipo de cambio (USD por 1 unidad de moneda local)", dim, "USD/u.moneda", "CONDICIÓN DE RONDA")
        row_fx[dim] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=fx_hist[dim][1], kind="historico", numfmt='0.0000', align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=fx_r2[dim], kind="input", numfmt='0.0000', align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt='0.0000')

    # -- Efectivo mínimo --
    row_efmin = put("B. Condiciones", "Efectivo mínimo de fin de ronda", "Todos los países (por país)", "USD", "CONDICIÓN DE RONDA",
                     "REGLA VERIFICADA (manual cap.11): USD 2.000.000/país en R1-R2. Fase 2: el motor de este workbook usa Balance/Caja POR PAÍS -- este valor se aplica directamente a cada país (ya no se poolea ×3, esa simplificación quedó reemplazada).")
    S.apply_cell(ws, row_efmin, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_efmin, 7 + 1, value=2_000_000.0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_efmin, 7 + 2, value=2_000_000.0, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_efmin, 2, numfmt=S.NUM_MONEY)

    # -- Tasas de interés --
    row_tlp = put("B. Condiciones", "Tasa de interés — deuda a largo plazo", "Casa matriz EE.UU.", "%/ronda", "CONDICIÓN DE RONDA")
    S.apply_cell(ws, row_tlp, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_tlp, 7 + 1, value=0.08, kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_tlp, 7 + 2, value=0.05, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_tlp, 2, numfmt=S.NUM_PCT1)

    tasa_cp_hist = {"EE.UU.": 0.15, "China": 0.10, "Europa": 0.085}
    tasa_cp_r2 = {"EE.UU.": 0.10, "China": 0.07, "Europa": 0.055}
    row_tcp = {}
    for pais in PAISES:
        row = put("B. Condiciones", "Tasa de interés — deuda a corto plazo", pais, "%/ronda", "CONDICIÓN DE RONDA")
        row_tcp[pais] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=tasa_cp_hist[pais], kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=tasa_cp_r2[pais], kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    tasa_ef_hist = {"EE.UU.": 0.06, "China": 0.065, "Europa": 0.055}
    tasa_ef_r2 = {"EE.UU.": 0.04, "China": 0.055, "Europa": 0.03}
    row_tef = {}
    for pais in PAISES:
        row = put("B. Condiciones", "Tasa de interés sobre efectivo", pais, "%/ronda", "CONDICIÓN DE RONDA")
        row_tef[pais] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=tasa_ef_hist[pais], kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=tasa_ef_r2[pais], kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    # -- Costo de la característica --
    car_hist = {"EE.UU.": 700.0, "China": 650.0, "Europa": 750.0}
    row_car = {}
    for pais in PAISES:
        row = put("B. Condiciones", "Costo de la característica", pais, "USD/u.", "CONDICIÓN DE RONDA", "Verificado exacto contra RDOS R1 (desglose de margen por tecnología). Es un costo VARIABLE de OFRECER la característica (se carga al costo de venta por unidad vendida, manual cap. 8) -- distinto de generarla o adquirirla (ver las 2 filas siguientes).")
        row_car[pais] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=car_hist[pais], kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=car_hist[pais], kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PRICE)

    # -- Costo de adquirir tecnología/característica nueva (Adenda 7 / hotfix README) -- CONDICIÓN
    # DE RONDA (parámetro que publica CESIM en "Parámetros" de "Estimaciones", no un supuesto propio),
    # distinta del "Costo de la característica" de arriba: esa es el costo VARIABLE de OFRECER una
    # característica ya disponible (por unidad vendida, en Marketing); esto es el costo de AMPLIAR lo
    # disponible -- por tecnología, no por país/mercado (manual cap. 7) -- vía desarrollo interno
    # (I+D propio, fila "Jornadas propias asignadas a desarrollo" en D4) o compra directa (fila
    # "Compra de licencia (USD)" en D4). No se inventa el monto: queda vacío/editable hasta
    # confirmarlo contra CESIM.
    row_costo_tec_nueva = {}
    for tech in TECNOLOGIAS:
        row = put("B. Condiciones", "Costo de adquirir tecnología nueva", tech, "USD", "CONDICIÓN DE RONDA",
                   "No confirmado contra CESIM todavía (ver Parámetros de Estimaciones). Aplica solo si CADIZ "
                   "incorpora esta tecnología (Combustión/Híbrido ya están habilitadas).")
        row_costo_tec_nueva[tech] = row
        if tech in TECH_ACTIVAS:
            for rn in RONDAS:
                S.apply_cell(ws, row, 7 + rn, value="n/a (ya habilitada para CADIZ)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        else:
            S.apply_cell(ws, row, 7 + 0, value="n/a (no verificado)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 1, value="n/a (no verificado)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 2, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
            arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    row_costo_caract_nueva = {}
    for tech in TECNOLOGIAS:
        row = put("B. Condiciones", "Costo de generar/adquirir característica nueva", tech, "USD", "CONDICIÓN DE RONDA",
                   "No confirmado contra CESIM todavía (ver Parámetros de Estimaciones). Cubre tanto la ruta de I+D "
                   "propio como la de compra de licencia (ver D4).")
        row_costo_caract_nueva[tech] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (no verificado)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value="n/a (no verificado)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    # -- Administración: fijo / planta / % --
    admin_fijo_hist = {"EE.UU.": 35_000_000.0, "China": 71_000_000.0, "Europa": 20_000_000.0}
    admin_fijo_r2 = {"EE.UU.": 35_000_000.0, "China": 74_000_000.0, "Europa": 20_000_000.0}
    row_admin_fijo = {}
    for pais in PAISES:
        row = put("B. Condiciones", "Administración — componente fijo", pais, "USD", "CONDICIÓN DE RONDA",
                   "ALERTA (ver 04_CONTROL_MODELO): fórmula lineal fijo+planta×N°fábricas+%ingresos subestima fuerte el real R1 (~3,2x en EE.UU.) — el manual confirma economías de escala por fábrica no capturadas aquí.")
        row_admin_fijo[pais] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=admin_fijo_hist[pais], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=admin_fijo_r2[pais], kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    admin_planta_hist = {"EE.UU.": 4_000_000.0, "China": 12_000_000.0}
    admin_planta_r2 = {"EE.UU.": 4_000_000.0, "China": 13_000_000.0}
    row_admin_planta = {}
    for area in AREAS:
        row = put("B. Condiciones", "Administración — componente por planta (aplicado linealmente, ver ALERTA)", area, "USD/fábrica", "CONDICIÓN DE RONDA")
        row_admin_planta[area] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=admin_planta_hist[area], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=admin_planta_r2[area], kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    row_admin_pct = {}
    for pais in PAISES:
        row = put("B. Condiciones", "Administración — componente % sobre ingresos", pais, "%", "CONDICIÓN DE RONDA")
        row_admin_pct[pais] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=0.0, kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=0.0, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    # -- RRHH: % sueldo bruto, otros costos, contratación, despido --
    row_sueldopct = put("B. Condiciones", "% Sueldo bruto sobre costo laboral total (RRHH)", "Global", "%", "CONDICIÓN DE RONDA", "Verificado exacto contra RDOS R1.")
    S.apply_cell(ws, row_sueldopct, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_sueldopct, 7 + 1, value=0.70, kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_sueldopct, 7 + 2, value=0.70, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_sueldopct, 2, numfmt=S.NUM_PCT1)

    row_otrosid = put("B. Condiciones", "Otros costos de I+D por empleado/mes", "Global", "USD/empleado/mes", "CONDICIÓN DE RONDA", "Calibrado contra RDOS R1 con 0,12% de diferencia (aceptado, no forzado).")
    S.apply_cell(ws, row_otrosid, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_otrosid, 7 + 1, value=680.0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_otrosid, 7 + 2, value=670.0, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_otrosid, 2, numfmt=S.NUM_PRICE)

    row_contrat = put("B. Condiciones", "Costo de contratación por persona", "Global", "USD/persona", "CONDICIÓN DE RONDA")
    S.apply_cell(ws, row_contrat, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_contrat, 7 + 1, value=4_100.0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_contrat, 7 + 2, value=4_000.0, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_contrat, 2, numfmt=S.NUM_PRICE)

    despido_hist = {"tramo 5%-20% de reducción": 18_600.0, "tramo 20%-50% de reducción": 23_600.0, "tramo >50% de reducción": 30_600.0}
    despido_r2 = {"tramo 5%-20% de reducción": 18_400.0, "tramo 20%-50% de reducción": 23_400.0, "tramo >50% de reducción": 30_400.0}
    row_despido = {}
    for tramo in despido_hist:
        row = put("B. Condiciones", "Costo de despido por persona", tramo, "USD/persona", "CONDICIÓN DE RONDA")
        row_despido[tramo] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=despido_hist[tramo], kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=despido_r2[tramo], kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PRICE)

    # -- Valor de referencia de la acción (proxy, PENDIENTE) --
    row_accion = put("B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", "USD/acción", "NO DETERMINADO",
                      "[NO DETERMINADO] Manual cap.11: la emisión/recompra se valúa 'de acuerdo con la valuación del mercado al inicio de la ronda', no a un valor de libro. PROXY = último precio de mercado real conocido (cierre R1). No es una predicción del algoritmo real de CESIM.")
    S.apply_cell(ws, row_accion, 7 + 0, value=33.59236009933182, kind="historico", numfmt='0.00', align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_accion, 7 + 1, value=38.65179490691809, kind="historico", numfmt='0.00', align=S.ALIGN_RIGHT, size=9)
    S.apply_cell(ws, row_accion, 7 + 2, value=38.65179490691809, kind="input", numfmt='0.00', align=S.ALIGN_RIGHT, size=9)
    arrastre_rows(row_accion, 2, numfmt='0.00')

    # -- Transporte / aranceles --
    PARES = [("EE.UU.", "China"), ("EE.UU.", "Europa"), ("China", "EE.UU."), ("China", "Europa")]
    transp_hist = {("EE.UU.", "China"): 550.0, ("EE.UU.", "Europa"): 300.0, ("China", "EE.UU."): 450.0, ("China", "Europa"): 400.0}
    transp_r2 = {("EE.UU.", "China"): 350.0, ("EE.UU.", "Europa"): 200.0, ("China", "EE.UU."): 300.0, ("China", "Europa"): 250.0}
    row_transp = {}
    for par in PARES:
        dim = f"{par[0]} → {par[1]}"
        row = put("B. Condiciones", "Costo de transporte por unidad", dim, "USD/u.", "CONDICIÓN DE RONDA")
        row_transp[par] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=transp_hist[par], kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=transp_r2[par], kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PRICE)

    row_arancel_esp = {}
    for par in PARES:
        dim = f"{par[0]} → {par[1]}"
        row = put("B. Condiciones", "Arancel específico por unidad", dim, "USD/u.", "CONDICIÓN DE RONDA")
        row_arancel_esp[par] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=0.0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=0.0, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PRICE)

    arancel_adval_hist = {("EE.UU.", "China"): 0.30, ("EE.UU.", "Europa"): 0.20, ("China", "EE.UU."): 0.30, ("China", "Europa"): 0.20}
    arancel_adval_r2 = {("EE.UU.", "China"): 0.20, ("EE.UU.", "Europa"): 0.10, ("China", "EE.UU."): 0.025, ("China", "Europa"): 0.10}
    row_arancel_adval = {}
    for par in PARES:
        dim = f"{par[0]} → {par[1]}"
        row = put("B. Condiciones", "Arancel ad valorem (%)", dim, "%", "CONDICIÓN DE RONDA",
                   "[SIMPLIFICACIÓN, ver README] Base = precio de transferencia únicamente. El manual (cap. 9.1) define la base como precio de transferencia + costos de característica; no modelado por complejidad — carry-over documentado del motor anterior.")
        row_arancel_adval[par] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=arancel_adval_hist[par], kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=arancel_adval_r2[par], kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_PCT1)

    # -- Costo de inversión / precio de venta por fábrica --
    costo_fab_hist = {"EE.UU.": 1_207_500_000.0, "China": 1_057_500_000.0}
    costo_fab_r2 = {"EE.UU.": 1_202_000_000.0, "China": 1_052_000_000.0}
    row_costo_fab = {}
    for area in AREAS:
        row = put("B. Condiciones", "Costo de inversión por fábrica", area, "USD", "CONDICIÓN DE RONDA")
        row_costo_fab[area] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=costo_fab_hist[area], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=costo_fab_r2[area], kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    row_precio_fab = {}
    for area in AREAS:
        row = put("B. Condiciones", "Precio de venta por fábrica (desinversión)", area, "USD", "CONDICIÓN DE RONDA")
        row_precio_fab[area] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=costo_fab_hist[area], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=costo_fab_r2[area], kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_MONEY)

    # -- Costo unitario de producción tercerizada --
    costo_terc_r0 = {"EE.UU.": {"Combustión": 11425.6408888889}, "China": {"Combustión": 10264.4712345679}}
    costo_terc_r1 = {"EE.UU.": {"Combustión": 11142.1602938292, "Híbrido": 17630.5229121633},
                      "China": {"Combustión": 10358.738390622, "Híbrido": 16285.8239718646}}
    costo_terc_r2 = {"China": {"Combustión": 10106.74, "Híbrido": 15649.15}}
    row_costo_terc = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row = put("B. Condiciones", "Costo unitario de producción tercerizada", dim, "USD/u.", "NO DETERMINADO / PENDIENTE DE CARGA",
                       "[NO DETERMINADO desde R2] El manual (cap. 5) confirma que este costo varía según el volumen fabricado y se revela recién en la pantalla de decisión — NO es una cifra publicada de antemano en Condiciones. No arrastra: completar manualmente antes de tercerizar. Ver ALERTA bloqueante en 04_CONTROL_MODELO.")
            row_costo_terc[(area, tech)] = row
            v0 = costo_terc_r0.get(area, {}).get(tech)
            v1 = costo_terc_r1.get(area, {}).get(tech, v0)
            v2 = costo_terc_r2.get(area, {}).get(tech)
            S.apply_cell(ws, row, 7 + 0, value=v0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 1, value=v1, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 2, value=v2, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)

    # -- Costos de gestión de inventario (fijo/variable, por país) y parámetros de costo de --
    # sostenibilidad (energía, agua, CO2). Adenda 29 (a pedido explícito del equipo): pasan de
    # NO DETERMINADO a CONDICIÓN DE RONDA estándar -- ya están WIRED a fórmula real en
    # build_engine_financiero (Costos de gestión de inventario / Energía / Agua / Carbono), leídas
    # con iref() como cualquier otra Condición de Ronda. Siguen en blanco hasta que el equipo cargue
    # la tasa: la nota de Fuente indica dónde confirmar el monto antes de tipearlo.
    row_ginv_fijo, row_ginv_var = {}, {}
    for pais in ["EE.UU.", "China"]:
        rowf = put("B. Condiciones", "Costos fijos de gestión de inventario", pais, "miles USD", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.2 (promedio inventario apertura+cierre). Confirmar monto contra pantalla CESIM antes de cargar.")
        row_ginv_fijo[pais] = rowf
        rowv = put("B. Condiciones", "Costos variables de gestión de inventario", pais, "USD/unidad", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.2. Confirmar monto contra pantalla CESIM antes de cargar.")
        row_ginv_var[pais] = rowv
        for rn in RONDAS:
            S.apply_cell(ws, rowf, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, rowv, 7 + rn, value=None, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)

    # Tarifas de energía/agua/carbono, por país productor -- WIRED a Costos de Energía/Agua/Carbono
    # en build_engine_financiero. % de utilización de energía renovable se muda acá (Sección B,
    # antes vivía en un bloque D9 aparte) porque alimenta directamente la tarifa promedio ponderada
    # de Energía, junto al resto de las condiciones de esta sección.
    row_costo_erenov, row_costo_efosil, row_costo_agua, row_costo_carbono = {}, {}, {}, {}
    row_pct_renov = {}
    for area in AREAS:
        row_costo_erenov[area] = put("B. Condiciones", "Costo de la energía renovable", area, "USD/kWh", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.1. Confirmar monto por país antes de cargar.")
        row_costo_efosil[area] = put("B. Condiciones", "Costo de la energía fósil", area, "USD/kWh", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.1. Confirmar monto por país antes de cargar.")
        row_costo_agua[area] = put("B. Condiciones", "Costo del agua", area, "USD/m3", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.1. Confirmar monto por país antes de cargar.")
        row_costo_carbono[area] = put("B. Condiciones", "Costo del Carbono", area, "USD/ton", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.1. Confirmar monto por país antes de cargar.")
        row_pct_renov[area] = put("B. Condiciones", "% de utilización de energía renovable (vs. fósil)", area, "%", "CONDICIÓN DE RONDA",
                    "Manual cap. 5.1 ('puede establecer la proporción de su utilización de energía entre renovable y fósil'). "
                    "Confirmar mecánica/pantalla CESIM antes de cargar.")
        for row in (row_costo_erenov[area], row_costo_efosil[area], row_costo_agua[area], row_costo_carbono[area]):
            for rn in RONDAS:
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        for rn in RONDAS:
            S.apply_cell(ws, row_pct_renov[area], 7 + rn, value=None, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)

    # Factores físicos de consumo (fabricación) por tecnología × país productor -- WIRED a Costos de
    # Energía/Agua/Carbono en build_engine_financiero.
    row_consumo_kwh, row_consumo_m3, row_emision_co2 = {}, {}, {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_consumo_kwh[(area, tech)] = put("B. Condiciones", "Consumo de energía (kWh por unidad producida)", dim, "kWh/u.", "CONDICIÓN DE RONDA",
                        "Manual cap. 5.1. No publicado por tecnología -- confirmar antes de cargar.")
            row_consumo_m3[(area, tech)] = put("B. Condiciones", "Consumo de agua (m3 por unidad producida)", dim, "m3/u.", "CONDICIÓN DE RONDA",
                        "Manual cap. 5.1. No publicado por tecnología -- confirmar antes de cargar.")
            row_emision_co2[(area, tech)] = put("B. Condiciones", "Emisión de CO2 (kg por unidad producida)", dim, "kg/u.", "CONDICIÓN DE RONDA",
                        "Manual cap. 5.1. No publicado por tecnología -- confirmar antes de cargar.")
            for row in (row_consumo_kwh[(area, tech)], row_consumo_m3[(area, tech)], row_emision_co2[(area, tech)]):
                for rn in RONDAS:
                    S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt='#,##0.0000', align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    cond = dict(fx=row_fx, efmin=row_efmin, tlp=row_tlp, tcp=row_tcp, tef=row_tef, car=row_car,
                admin_fijo=row_admin_fijo, admin_planta=row_admin_planta, admin_pct=row_admin_pct,
                sueldopct=row_sueldopct, otrosid=row_otrosid, contrat=row_contrat, despido=row_despido,
                accion=row_accion, transp=row_transp, arancel_esp=row_arancel_esp, arancel_adval=row_arancel_adval,
                costo_fab=row_costo_fab, precio_fab=row_precio_fab, costo_terc=row_costo_terc)

    # ---------------------------------------------------------------------------------
    # SECCIÓN C · DECISIONES D1-D8 (catálogo verificado, ver README para exclusiones D1/D2/D9 no migradas)
    # ---------------------------------------------------------------------------------
    # ---------------------------------------------------------------------------------
    # SECCIÓN C.D1a-c · DEMANDA — MIX DE INDUSTRIA (TAM) [Adenda 22]
    # ---------------------------------------------------------------------------------
    # [ARQUITECTURA A PEDIDO EXPLÍCITO DEL EQUIPO] El manual CESIM (cap. 4.2) es taxativo: "el
    # porcentaje estimado [D1] se refiere a TODA el área del mercado, NO a la demanda de esa
    # tecnología específica" -- es decir, el input D1 histórico ("Cuota de mercado objetivo CADIZ")
    # YA es % del mercado TOTAL, no % de la tecnología. Multiplicar ese valor por un Mix de Industria
    # adicional duplicaría/descontaría la demanda: es el mismo bug que la etapa anterior detectó y
    # revirtió (ver claude/Prueba_Operacional_R2_Escenario_A_v2.md: "la cuota ya viene expresada
    # como % del mercado regional total"). Para evitar reintroducirlo, el equipo definió una
    # arquitectura de 3 bloques que REDEFINE el input D1 en vez de multiplicarlo directamente:
    #   D1a) Mix de Industria (TAM)                              = % del TAM que representa cada tecnología.
    #   D1b) Cuota de mercado objetivo SOBRE LA TECNOLOGÍA        = % que CADIZ apunta a capturar
    #        (antes "Cuota de mercado objetivo CADIZ")              DENTRO del segmento de esa tecnología (NUEVA semántica).
    #   D1c) Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)     = D1a × D1b, fila a fila -- ÉSTA
    #        es la que alimenta _ENGINE_MERCADO (Sección A2, Demanda estimada CADIZ) y reproduce la
    #        semántica "% del mercado TOTAL" que exige el manual.
    # R0/R1: "n/a" en los tres bloques -- la demanda de esas rondas es histórico real hardcodeado
    # (Sección A2 de _ENGINE_MERCADO no lee este bloque para R0/R1). R2: Estado(R2)="PLAN" (decisión
    # ya cerrada, RDOS de R2 aún no publicados) -- el Mix de Industria R2 se estima con el ÚNICO dato
    # real disponible (Mix de Industria R1, DATO HISTÓRICO REAL calculado de 'Informes de mercado'
    # RDOS Ronda 1: demanda de las 7 empresas por tecnología, sumada, sobre el tamaño total de esa
    # área) y la Cuota sobre la Tecnología R2 se retro-calculó (Resultante_R2_vieja / Mix_R2) para
    # que la Cuota Resultante R2 reproduzca EXACTO el valor ya decidido -- no se altera ninguna
    # proyección de una ronda cerrada. R3-R12: ambos inputs quedan en blanco para carga real del
    # equipo, ronda a ronda, con la mejor estimación de Mix de Industria disponible en cada momento.
    # NO incluye "Factores de estrés": el equipo confirmó (respuesta explícita) que ese término viene
    # de material de cátedra/profesor, no del manual CESIM, y decidió no incorporarlo por ahora al no
    # contar con su definición/fórmula concreta -- queda pendiente, no debe inferirse ni inventarse.
    # ---------------------------------------------------------------------------------
    section("SECCIÓN C.D1a · DEMANDA — Mix de Industria (TAM) [SUPUESTO PROPIO — no es un input directo publicado por CESIM]")
    # Mix de Industria R1: DATO HISTÓRICO REAL, derivado de 'Informes de mercado' RDOS Ronda 1
    # (demanda de las 7 empresas por tecnología, sumada, sobre el tamaño total de mercado de esa
    # área). Verificado: la suma da 5.192.959 (EE.UU.), 3.808.260 (China) y 6.646.784 (Europa),
    # reproduciendo exactamente BASE_R1_AJUSTADA usada más abajo para proyectar Ronda 2. Eléctrico e
    # Hidrógeno = 0% (no comercializados por ninguna empresa en R1).
    MIX_INDUSTRIA_R1 = {
        ("EE.UU.", "Combustión"): 0.809504, ("EE.UU.", "Híbrido"): 0.190496,
        ("EE.UU.", "Eléctrico"): 0.0, ("EE.UU.", "Hidrógeno"): 0.0,
        ("China", "Combustión"): 0.814027, ("China", "Híbrido"): 0.185973,
        ("China", "Eléctrico"): 0.0, ("China", "Hidrógeno"): 0.0,
        ("Europa", "Combustión"): 0.862040, ("Europa", "Híbrido"): 0.137960,
        ("Europa", "Eléctrico"): 0.0, ("Europa", "Hidrógeno"): 0.0,
    }
    row_mix = {}
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row = put("D1a · MIX INDUSTRIA", "Mix de Industria (TAM)", dim, "% del TAM", "SUPUESTO PROPIO",
                       "% que esta tecnología representa del tamaño TOTAL de mercado de esta área (todas las empresas, todas las "
                       "tecnologías = 100%). R2 = Mix de Industria R1 (DATO HISTÓRICO REAL, ver nota de la sección) usado como "
                       "estimación de partida porque aún no hay RDOS de R2 publicado -- para R2 este valor es SUPUESTO, no dato real. "
                       "R3-R12: a cargar por el equipo con la mejor estimación disponible en cada ronda.")
            row_mix[(mercado, tech)] = row
            S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            mix_r2 = MIX_INDUSTRIA_R1.get((mercado, tech), 0.0)
            S.apply_cell(ws, row, 7 + 2, value=mix_r2, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    section("SECCIÓN C.D1b · DEMANDA — Cuota de mercado objetivo SOBRE LA TECNOLOGÍA")
    # Retro-calculado para R2 = (vieja "Cuota de mercado objetivo CADIZ" R2, % del TOTAL) / (Mix de
    # Industria R2 de arriba) -- así la Cuota Resultante (D1c, más abajo) reproduce EXACTO el valor
    # ya decidido para R2 (0,137 / 0,054 / 0,162 / 0,056 / 0,100 / 0,047), sin alterar la proyección
    # de una ronda ya cerrada (Estado R2 = "PLAN"). Verificación: 0,169239×0,809504=0,137000;
    # 0,283471×0,190496=0,054000; 0,199011×0,814027=0,162000; 0,301118×0,185973=0,056000;
    # 0,116004×0,862040=0,100000; 0,340680×0,137960=0,047000 (redondeo a 6 decimales).
    cuota_tec_r2 = {
        ("EE.UU.", "Combustión"): 0.169239, ("EE.UU.", "Híbrido"): 0.283471,
        ("China", "Combustión"): 0.199011, ("China", "Híbrido"): 0.301118,
        ("Europa", "Combustión"): 0.116004, ("Europa", "Híbrido"): 0.340680,
    }
    row_cuota = {}
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row = put("D1 · DEMANDA", "Cuota de mercado objetivo SOBRE LA TECNOLOGÍA", dim, "% de la tecnología", "DECISIÓN CADIZ",
                       "[NUEVA SEMÁNTICA, Adenda 22 — antes 'Cuota de mercado objetivo CADIZ', % del mercado TOTAL] Ahora es el % que "
                       "CADIZ apunta a capturar DENTRO del segmento de esta tecnología (se multiplica por 'Mix de Industria (TAM)' de "
                       "arriba para obtener la 'Cuota Resultante SOBRE EL TOTAL' que efectivamente se carga en CESIM). R2: retro-"
                       "calculado para no alterar la proyección ya decidida — ver nota de la sección.")
            row_cuota[(mercado, tech)] = row
            S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            v2 = cuota_tec_r2.get((mercado, tech), 0.0)
            S.apply_cell(ws, row, 7 + 2, value=v2, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    # -- Completa RONDA_ACTIVA (reservada en Sección A) ahora que existe row_cuota --
    for rn in RONDAS:
        if rn in rondas_reales:
            S.apply_cell(ws, row_ronda_activa, 7 + rn, value=True, kind="calculo", align=S.ALIGN_CENTER, bold=True, size=9)
        else:
            estado_cell = f"{col(rn)}{row_estado}"
            cuota_cell = f"{col(rn)}{row_cuota[('EE.UU.', 'Combustión')]}"
            f_activa = f'=AND(OR({estado_cell}="PLAN",{estado_cell}="SIM",{estado_cell}="REAL"),ISNUMBER({cuota_cell}))'
            S.apply_cell(ws, row_ronda_activa, 7 + rn, value=f_activa, kind="calculo", align=S.ALIGN_CENTER, bold=True, size=9)

    section("SECCIÓN C.D1c · DEMANDA — Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM) [CALCULADO = Mix de Industria × Cuota sobre la Tecnología]")
    row_cuota_resultante = {}
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row = put("D1c · DEMANDA (RESULTANTE)", "Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)", dim, "% del mercado", "CÁLCULO",
                       "= Mix de Industria (TAM) × Cuota de mercado objetivo SOBRE LA TECNOLOGÍA, fila a fila. Este es el valor que "
                       "efectivamente se carga en CESIM y el que alimenta la Demanda estimada CADIZ de _ENGINE_MERCADO (Sección A2) -- "
                       "ya expresado como % del mercado TOTAL, conforme al manual CESIM cap. 4.2. R0/R1: n/a (demanda histórica "
                       "hardcodeada). R2: reproduce EXACTO el valor ya decidido para la ronda (Estado=PLAN, cerrada) por construcción "
                       "de los dos bloques de arriba.")
            row_cuota_resultante[(mercado, tech)] = row
            S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            for rn in range(2, 13):
                mix_cell = f"{col(rn)}{row_mix[(mercado, tech)]}"
                tec_cell = f"{col(rn)}{row_cuota[(mercado, tech)]}"
                f = f"={mix_cell}*{tec_cell}"
                S.apply_cell(ws, row, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    section("SECCIÓN C.D2 · PRODUCCIÓN")
    prod_propia_hist = {
        ("EE.UU.", "Combustión"): {0: 980, 1: 336}, ("EE.UU.", "Híbrido"): {0: 0, 1: 616},
        ("EE.UU.", "Eléctrico"): {0: 0, 1: 0}, ("EE.UU.", "Hidrógeno"): {0: 0, 1: 0},
        ("China", "Combustión"): {0: 400, 1: 240}, ("China", "Híbrido"): {0: 0, 1: 100},
        ("China", "Eléctrico"): {0: 0, 1: 0}, ("China", "Hidrógeno"): {0: 0, 1: 0},
    }
    prod_propia_r2 = {("EE.UU.", "Combustión"): 1200, ("EE.UU.", "Híbrido"): 60, ("EE.UU.", "Eléctrico"): 0, ("EE.UU.", "Hidrógeno"): 0,
                       ("China", "Combustión"): 400, ("China", "Híbrido"): 50, ("China", "Eléctrico"): 0, ("China", "Hidrógeno"): 0}
    row_prod_propia = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row = put("D2 · PRODUCCIÓN", "Producción propia", dim, "miles u.", "DECISIÓN CADIZ")
            row_prod_propia[(area, tech)] = row
            hv = prod_propia_hist[(area, tech)]
            S.apply_cell(ws, row, 7 + 0, value=hv[0], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 1, value=hv[1], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 2, value=prod_propia_r2[(area, tech)], kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)

    prod_terc_hist = {
        ("EE.UU.", "Combustión"): {0: 250, 1: 480}, ("EE.UU.", "Híbrido"): {0: 0, 1: 200},
        ("EE.UU.", "Eléctrico"): {0: 0, 1: 0}, ("EE.UU.", "Hidrógeno"): {0: 0, 1: 0},
        ("China", "Combustión"): {0: 400, 1: 360}, ("China", "Híbrido"): {0: 0, 1: 170},
        ("China", "Eléctrico"): {0: 0, 1: 0}, ("China", "Hidrógeno"): {0: 0, 1: 0},
    }
    prod_terc_r2 = {("EE.UU.", "Combustión"): 0, ("EE.UU.", "Híbrido"): 0, ("EE.UU.", "Eléctrico"): 0, ("EE.UU.", "Hidrógeno"): 0,
                     ("China", "Combustión"): 536, ("China", "Híbrido"): 343, ("China", "Eléctrico"): 0, ("China", "Hidrógeno"): 0}
    row_prod_terc = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row = put("D2 · PRODUCCIÓN", "Producción tercerizada", dim, "miles u.", "DECISIÓN CADIZ",
                       "China/Combustión y China/Híbrido en R2 incluyen +35k/+15k de safety stock (decisión estratégica ya consolidada del equipo, ver histórico R2 del motor anterior).")
            row_prod_terc[(area, tech)] = row
            hv = prod_terc_hist[(area, tech)]
            S.apply_cell(ws, row, 7 + 0, value=hv[0], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 1, value=hv[1], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row, 7 + 2, value=prod_terc_r2[(area, tech)], kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)

    row_inv_fab = {}
    for area in AREAS:
        row = put("D2 · PRODUCCIÓN", "Inversión (+) / Desinversión (−) en fábrica", area, "u. de capacidad", "DECISIÓN CADIZ",
                   "Inversión: capacidad +2 rondas, pago +1 ronda. Desinversión: capacidad −1 ronda, cobro +1 ronda.")
        row_inv_fab[area] = row
        S.apply_cell(ws, row, 7 + 0, value=0, kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 1, value=0, kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=0, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    section("SECCIÓN C.D3 · RRHH (decisión única, GLOBAL — manual cap. 6; costo se reparte por N° de fábricas)")
    d3_r2 = {"Cantidad de personal de I+D": 5300, "Salario mensual": 6300, "Presupuesto mensual de capacitación": 1100}
    row_d3 = {}
    for var, unidad in [("Cantidad de personal de I+D", "personas"), ("Salario mensual", "USD"), ("Presupuesto mensual de capacitación", "USD")]:
        row = put("D3 · RRHH", var, "Global", unidad, "DECISIÓN CADIZ")
        row_d3[var] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (RDOS no publica headcount por ronda)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value="n/a (RDOS no publica headcount por ronda)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=d3_r2[var], kind="input", numfmt=S.NUM_MONEY if unidad == "USD" else S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY if unidad == "USD" else S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    section("SECCIÓN C.D4 · I+D / TECNOLOGÍA")
    jornadas_r2 = {"Combustión": 350961, "Híbrido": 591076, "Eléctrico": 0, "Hidrógeno": 0}
    row_jornadas = {}
    row_licencia = {}
    for tech in TECNOLOGIAS:
        row = put("D4 · I+D / TECNOLOGÍA", "Jornadas propias asignadas a desarrollo", tech, "jornadas", "DECISIÓN CADIZ", "1 ronda de rezago hasta estar disponible (manual cap. 7) — no modelado como gatillo de desbloqueo en este workbook de gestión (ver README).")
        row_jornadas[tech] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=jornadas_r2[tech], kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        row2 = put("D4 · I+D / TECNOLOGÍA", "Compra de licencia (USD)", tech, "USD", "DECISIÓN CADIZ")
        row_licencia[tech] = row2
        S.apply_cell(ws, row2, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row2, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row2, 7 + 2, value=0, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row2, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)

    # CORRECCIÓN (Adenda 9, reemplaza el enfoque de Adenda 6/7 -- feedback explícito del usuario: "sigue
    # mostrando el total, no el diferencial nuevo que se va a desarrollar... me va a multiplicar por el
    # total y no por las generadas en esa ronda nada más"). El manual (cap. 7 vs. cap. 8, ya verificado)
    # distingue dos cosas que antes vivían mezcladas en una sola fila:
    #   1) Cuántas características tiene DISPONIBLES cada TECNOLOGÍA en total (stock acumulado, global,
    #      generado por I+D propio -fila "Jornadas..."- o comprado por licencia -fila "Compra de
    #      licencia (USD)"-, ambas ya existentes arriba en esta misma sección D4).
    #   2) De esas disponibles, cuántas se OFRECEN en cada MERCADO (decisión de Marketing, manual cap.
    #      8, tope 10) -- eso se movió a la Sección D5 más abajo ("Características ofrecidas").
    # Además, el costo de GENERAR/ADQUIRIR una característica nueva (Sección B) debe multiplicarse por
    # el DIFERENCIAL de este stock (las generadas ESTA ronda), nunca por el stock total acumulado --
    # si no, se recobraría cada ronda el costo de características ya generadas en rondas anteriores.
    # Por eso se agregan acá dos filas de CÁLCULO (diferencial y costo de referencia), no solo el stock.
    #
    # Baseline Ronda 1 (dato aportado por el usuario, no inventado ni heredado del motor anterior): el
    # usuario explicó el mecanismo con "si tengo 7 y genero 1 característica extra... paso a tener 8
    # disponibles", y confirmó para Ronda 2 "combustión: 8 características ya que desarrollamos una
    # extra" / "híbrido: 3 características ya que se generaron nuevas" -- de donde se deduce que el
    # stock ANTES de Ronda 2 (= Ronda 1) era Combustión=7, Híbrido=2 (Eléctrico/Hidrógeno=0, no
    # habilitadas). Es un dato histórico real aportado por el usuario (categoría 2 del marco del
    # proyecto), no un supuesto propio de este modelo.
    disponible_r1 = {"Combustión": 7, "Híbrido": 2, "Eléctrico": 0, "Hidrógeno": 0}
    disponible_r2 = {"Combustión": 8, "Híbrido": 3, "Eléctrico": 0, "Hidrógeno": 0}
    row_disponible = {}
    row_diferencial = {}
    row_costo_caract_calc = {}
    for tech in TECNOLOGIAS:
        row = put("D4 · I+D / TECNOLOGÍA", "Características disponibles (acumulado, stock por tecnología)", tech, "cantidad", "DECISIÓN CADIZ",
                   "Stock TOTAL acumulado de características de esta tecnología (I+D propio o licencia, manual cap. 7) -- GLOBAL, no por mercado. Ronda 1: dato aportado por el usuario (ver Adenda 9 en Informe_V2.md), no un supuesto propio. Distinto de 'Características ofrecidas' (Sección D5, Marketing): esa es cuántas de este stock se muestran en cada mercado.")
        row_disponible[tech] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=disponible_r1[tech], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=disponible_r2[tech], kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
        arrastre_rows(row, 2, numfmt=S.NUM_UNITS)

        row_d = put("D4 · I+D / TECNOLOGÍA", "Características nuevas generadas esta ronda (diferencial)", tech, "cantidad", "CÁLCULO",
                     "= stock disponible esta ronda − stock disponible ronda anterior (fila de arriba). Esta es la cantidad que corresponde multiplicar por 'Costo de generar/adquirir característica nueva' (Sección B) -- NUNCA el stock total acumulado.")
        row_diferencial[tech] = row_d
        S.apply_cell(ws, row_d, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        for rn in range(1, 13):
            cur_cell = f"{col(rn)}{row}"
            prev_cell = f"{col(rn - 1)}{row}"
            f_diff = f'=IF(AND(ISNUMBER({cur_cell}),ISNUMBER({prev_cell})),{cur_cell}-{prev_cell},"")'
            S.apply_cell(ws, row_d, 7 + rn, value=f_diff, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)

        row_cc = put("D4 · I+D / TECNOLOGÍA", "Costo estimado de generar/adquirir características nuevas (referencia)", tech, "USD", "CÁLCULO",
                      "= diferencial (fila de arriba) × 'Costo de generar/adquirir característica nueva' (Sección B, condición de ronda publicada por CESIM). Es una ESTIMACIÓN de referencia para dimensionar la decisión 'Compra de licencia (USD)' de esta misma sección -- NO se suma automáticamente a los Costos y gastos del P&L, para no duplicar con lo que ya se carga manualmente ahí (ver README, hotfix Adenda 9).")
        row_costo_caract_calc[tech] = row_cc
        costo_ref_row = row_costo_caract_nueva[tech]
        S.apply_cell(ws, row_cc, 7 + 0, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        for rn in range(1, 13):
            diff_cell = f"{col(rn)}{row_d}"
            costo_cell = f"{col(rn)}{costo_ref_row}"
            f_costo = f'=IF(AND(ISNUMBER({diff_cell}),ISNUMBER({costo_cell})),{diff_cell}*{costo_cell},"")'
            S.apply_cell(ws, row_cc, 7 + rn, value=f_costo, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
    r[0] += 1

    section("SECCIÓN C.D5 · MARKETING")
    # CORRECCIÓN (confirmada con el usuario, ver Adenda 6 / hotfix README) -- precio y promoción de
    # Ronda 2 actualizados con la decisión real de CADIZ. Promoción: el usuario entregó los montos en
    # "miles USD" (unidad nativa de la pantalla de decisión de CESIM para este campo, igual que el
    # resto de las cifras "miles USD" del proyecto) -- se multiplican x1000 acá para quedar en la
    # unidad declarada de esta fila (USD absoluto), confirmado explícitamente antes de aplicar (600.000
    # miles USD = USD 600.000.000, no USD 600.000 -- una diferencia de 1000x que se verificó en vez de
    # asumirse, dado que a simple vista el monto parecía ~700 veces más chico que el anterior).
    precio_r2 = {("EE.UU.", "Combustión"): 22500, ("EE.UU.", "Híbrido"): 28000,
                 ("China", "Combustión"): 140000, ("China", "Híbrido"): 175000,
                 ("Europa", "Combustión"): 27000, ("Europa", "Híbrido"): 31000}
    promo_r2 = {("EE.UU.", "Combustión"): 600_000_000, ("EE.UU.", "Híbrido"): 500_000_000,
                ("China", "Combustión"): 500_000_000, ("China", "Híbrido"): 400_000_000,
                ("Europa", "Combustión"): 500_000_000, ("Europa", "Híbrido"): 450_000_000}
    enfoque_r2 = {("EE.UU.", "Combustión"): "Marca", ("EE.UU.", "Híbrido"): "Marca",
                  ("China", "Combustión"): "Características", ("China", "Híbrido"): "Características",
                  ("Europa", "Combustión"): "Marca", ("Europa", "Híbrido"): "Marca"}
    # CORRECCIÓN (Adenda 9): "Características ofrecidas" es una decisión de MARKETING (manual cap. 8,
    # tope 10 por producto/mercado) -- de las características DISPONIBLES por tecnología (stock
    # global, Sección D4 más arriba), cuántas se muestran/ofrecen en ESTE mercado. Antes vivía
    # mezclada bajo el bloque "D4 · I+D / TECNOLOGÍA" (confundía el stock de I+D con la decisión de
    # Marketing); se reubica acá. Estrategia CADIZ confirmada por el usuario: siempre el máximo
    # disponible -> por eso replica en R2 el mismo valor de "Características disponibles" (D4) en
    # los 3 mercados (no es una coincidencia ni un dato hardcodeado two-veces: es la decisión real).
    caract_of_r2 = {(mercado, tech): disponible_r2.get(tech, 0) for mercado in MERCADOS for tech in TECNOLOGIAS}
    row_precio = {}
    row_promo = {}
    row_enfoque = {}
    row_caract = {}
    # Hotfix regional: lista oculta en columna U (misma columna que la Matriz de Sensibilidad,
    # filas 6-10) en vez de una cadena separada por comas -- rompía en Excel es-AR/es-LA.
    S.apply_cell(ws, 6, LIST_COL, value="Equilibrado", kind="plain", size=8)
    S.apply_cell(ws, 7, LIST_COL, value="Precio bajo", kind="plain", size=8)
    S.apply_cell(ws, 8, LIST_COL, value="Características", kind="plain", size=8)
    S.apply_cell(ws, 9, LIST_COL, value="Sostenibilidad", kind="plain", size=8)
    S.apply_cell(ws, 10, LIST_COL, value="Marca", kind="plain", size=8)
    dv_enfoque = DataValidation(type="list", formula1=f"${list_col_letter}$6:${list_col_letter}$10", allow_blank=True)
    ws.add_data_validation(dv_enfoque)
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row_p = put("D5 · MARKETING", "Precio de venta", dim, MONEDA_MERCADO[mercado], "DECISIÓN CADIZ")
            row_precio[(mercado, tech)] = row_p
            row_pr = put("D5 · MARKETING", "Presupuesto de promoción", dim, "USD", "DECISIÓN CADIZ")
            row_promo[(mercado, tech)] = row_pr
            row_e = put("D5 · MARKETING", "Enfoque de estrategia de marketing", dim, "categoría", "DECISIÓN CADIZ")
            row_enfoque[(mercado, tech)] = row_e
            row_c = put("D5 · MARKETING", "Características ofrecidas (de las disponibles, tope 10)", dim, "cantidad", "DECISIÓN CADIZ",
                        "De las características DISPONIBLES de esta tecnología (stock global, Sección D4), cuántas se ofrecen en ESTE mercado -- manual cap. 8, tope 10. No puede superar el stock disponible de D4 para esa tecnología. Es la cantidad que ya usaba (correctamente) la fórmula de 'Costos de la característica' del P&L -- ver iref más abajo.")
            row_caract[(mercado, tech)] = row_c
            dv_enfoque.add(f"{col(0)}{row_e}:{col(12)}{row_e}")
            for rr, valmap, numfmt in ((row_p, precio_r2, S.NUM_PRICE), (row_pr, promo_r2, S.NUM_MONEY), (row_e, enfoque_r2, None), (row_c, caract_of_r2, S.NUM_UNITS)):
                S.apply_cell(ws, rr, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
                S.apply_cell(ws, rr, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
                v2 = valmap.get((mercado, tech))
                S.apply_cell(ws, rr, 7 + 2, value=v2, kind="input", numfmt=numfmt, align=(S.ALIGN_CENTER if numfmt is None else S.ALIGN_RIGHT), size=9)
                for rn in range(3, 13):
                    S.apply_cell(ws, rr, 7 + rn, value=None, kind="input", numfmt=numfmt, align=(S.ALIGN_CENTER if numfmt is None else S.ALIGN_RIGHT), size=9)
    r[0] += 1

    section("SECCIÓN C.D6 · LOGÍSTICA — Orden de prioridad de entrega (1=primero, 3=último; ranking completo y sin empates por área/tecnología)")
    ranking_r2 = {
        ("EE.UU.", "Combustión", "EE.UU."): '2', ("EE.UU.", "Combustión", "China"): '3', ("EE.UU.", "Combustión", "Europa"): '1',
        ("EE.UU.", "Híbrido", "EE.UU."): '2', ("EE.UU.", "Híbrido", "China"): '3', ("EE.UU.", "Híbrido", "Europa"): '1',
        ("China", "Combustión", "EE.UU."): '3', ("China", "Combustión", "China"): '2', ("China", "Combustión", "Europa"): '1',
        ("China", "Híbrido", "EE.UU."): '2', ("China", "Híbrido", "China"): '3', ("China", "Híbrido", "Europa"): '1',
    }
    row_ranking = {}
    # Hotfix regional: lista oculta en columna U (filas 12-14) en vez de cadena separada por comas.
    S.apply_cell(ws, 12, LIST_COL, value="1", kind="plain", size=8)
    S.apply_cell(ws, 13, LIST_COL, value="2", kind="plain", size=8)
    S.apply_cell(ws, 14, LIST_COL, value="3", kind="plain", size=8)
    dv_rank = DataValidation(type="list", formula1=f"${list_col_letter}$12:${list_col_letter}$14", allow_blank=True)
    ws.add_data_validation(dv_rank)
    for area in AREAS:
        for tech in TECNOLOGIAS:
            for mercado in MERCADOS:
                dim = f"Área {area} / {tech} → mercado {mercado}"
                row = put("D6 · LOGÍSTICA", "Orden de prioridad de entrega (1=primero, 3=último)", dim, "ranking", "DECISIÓN CADIZ")
                row_ranking[(area, tech, mercado)] = row
                dv_rank.add(f"{col(0)}{row}:{col(12)}{row}")
                S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
                S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
                v2 = ranking_r2.get((area, tech, mercado))
                S.apply_cell(ws, row, 7 + 2, value=v2, kind="input", align=S.ALIGN_CENTER, size=9)
                for rn in range(3, 13):
                    S.apply_cell(ws, row, 7 + rn, value=None, kind="input", align=S.ALIGN_CENTER, size=9)
    r[0] += 1

    section("SECCIÓN C.D7 · IMPUESTOS / TRANSFER PRICING — Multiplicador de precio de transferencia")
    row_mult_tp = {}
    for par in PARES:
        dim = f"{par[0]} → {par[1]}"
        row = put("D7 · IMPUESTOS", "Multiplicador de precio de transferencia", dim, "x costo directo (1.0-2.0)", "DECISIÓN CADIZ",
                   "Rango [1,2] (manual cap. 10). Recomendación del profesor: no planificar, ajustar jugando.")
        row_mult_tp[par] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=1.0, kind="input", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    section("SECCIÓN C.D8 · FINANZAS (montos ABSOLUTOS en USD — la deuda de corto plazo NO es un input: es automática, ver 04_CONTROL_MODELO)")
    row_d8 = {}
    d8_simple_r2 = {
        "Variación de deuda de largo plazo": 0,
        "Emisión de acciones": 0,
        "Recompra de acciones": 0,
        "Dividendos pagados a accionistas (monto)": 3_000_000_000,
    }
    for var in d8_simple_r2:
        row = put("D8 · FINANZAS", var, "Casa matriz EE.UU.", "USD", "DECISIÓN CADIZ")
        row_d8[var] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        hv1 = {"Dividendos pagados a accionistas (monto)": 2_000_000_000}.get(var)
        if hv1 is not None:
            S.apply_cell(ws, row, 7 + 1, value=hv1, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        else:
            S.apply_cell(ws, row, 7 + 1, value="n/a (no cargado retroactivamente)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 2, value=d8_simple_r2[var], kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)

    row_transf = {}
    for origen in PAISES:
        for destino in PAISES:
            if origen == destino:
                continue
            dim = f"{origen} → {destino}"
            row = put("D8 · FINANZAS", "Transferencia de fondos entre países (préstamo interno)", dim, "USD", "DECISIÓN CADIZ", "Sin costo/interés (definición cerrada del equipo).")
            row_transf[(origen, destino)] = row
            S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 1, value="n/a (no cargado retroactivamente)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, row, 7 + 2, value=0, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
            for rn in range(3, 13):
                S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)

    row_div_filial = {}
    for filial in ("China", "Europa"):
        row = put("D8 · FINANZAS", "Dividendos pagados a casa matriz (filial)", filial, "USD", "NO DETERMINADO",
                   "[AMBIGÜEDAD NO RESUELTA] No confirmado si es una decisión separada de la transferencia de fondos o el mismo mecanismo clasificado distinto por CESIM. Default 0 hasta confirmar.")
        row_div_filial[filial] = row
        S.apply_cell(ws, row, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row, 7 + 1, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row, 7 + 2, value=0, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        for rn in range(3, 13):
            S.apply_cell(ws, row, 7 + rn, value=None, kind="input", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    ws.auto_filter.ref = f"A{header_row}:F{r[0]-1}"

    # Hotfix regional: columna U con las listas ocultas de Data Validation/MATCH (Escenario Activo,
    # Enfoque de marketing, Ranking de logística) -- oculta, no borrada.
    ws.column_dimensions[list_col_letter].hidden = True

    decisiones = dict(cuota=row_cuota, prod_propia=row_prod_propia, prod_terc=row_prod_terc, inv_fab=row_inv_fab,
                       d3=row_d3, jornadas=row_jornadas, licencia=row_licencia, caract=row_caract,
                       disponible=row_disponible, diferencial=row_diferencial, costo_caract_calc=row_costo_caract_calc,
                       precio=row_precio, promo=row_promo, enfoque=row_enfoque, ranking=row_ranking,
                       mult_tp=row_mult_tp, d8=row_d8, transf=row_transf, div_filial=row_div_filial)

    return ws, row_idx, {"ronda": row_ronda, "escenario": row_escenario, "estado": row_estado, "sens": sens_rows}, cond, decisiones


PARES_EXPORT = [("EE.UU.", "China"), ("EE.UU.", "Europa"), ("China", "EE.UU."), ("China", "Europa")]


def iref(inputs_idx, bloque, variable, dimension, ronda):
    row = inputs_idx[f"{bloque}||{variable}||{dimension}"]
    return f"'{INPUTS_SHEET}'!{col(ronda)}{row}"


def sens_ref(row):
    """Referencia absoluta (celda única, no por ronda) a una fila de la Matriz de Sensibilidad de
    01_INPUTS (Fase 2, cambio 4). Columna C (no G): hotfix regional -- la matriz ya no ocupa las
    columnas de ronda G:S; columna C porque D quedó reservada para "Optimista" visible y "Escenario
    Activo"/los factores calculados reutilizan la columna C en sus propias filas (hotfix 2, ver
    build_inputs())."""
    return f"'{INPUTS_SHEET}'!$C${row}"


# ======================================================================================
# _ENGINE_PRODUCCION — Parte 1: capacidad, curva de aprendizaje, costo unitario propio,
# disponibilidad (oferta previa a ventas), D&A. (Motor E1 + E4 del proyecto anterior, simplificado.)
# ======================================================================================
def build_engine_produccion_part1(wb, inputs_idx, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(ENGINE_PROD_SHEET)
    ncols = 6 + len(RONDAS)
    S.set_col_widths(ws, [18, 44, 20, 10, 14, 30] + [13] * len(RONDAS))
    S.style_title_row(ws, 1, "_ENGINE_PRODUCCION — Capacidad, curva de aprendizaje, costo unitario propio, disponibilidad e inventario (costeo FIFO multicapa), depreciación", ncols)
    header_row = 3
    headers = ["Bloque", "Variable", "Dimensión", "Unidad", "Tipo", "Nota"] + [f"R{r}" for r in RONDAS]
    S.style_header_row(ws, header_row, headers)
    r = [header_row + 1]
    row_idx = {}

    def put(bloque, variable, dimension, unidad="", tipo="CÁLCULO", nota=""):
        rr = r[0]
        row_idx[f"{bloque}||{variable}||{dimension}"] = rr
        S.apply_cell(ws, rr, 1, value=bloque, kind="calculo", align=S.ALIGN_LEFT, size=8, bold=True)
        S.apply_cell(ws, rr, 2, value=variable, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=8)
        S.apply_cell(ws, rr, 3, value=dimension, kind="calculo", align=S.ALIGN_LEFT, size=8)
        S.apply_cell(ws, rr, 4, value=unidad, kind="calculo", align=S.ALIGN_CENTER, size=7)
        S.apply_cell(ws, rr, 5, value=tipo, kind="calculo", align=S.ALIGN_CENTER, size=7, bold=True)
        S.apply_cell(ws, rr, 6, value=nota, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=6, italic=True)
        r[0] += 1
        return rr

    def section(title):
        S.style_section_row(ws, r[0], title, ncols)
        r[0] += 1

    # -------- PARÁMETROS DE MODELIZACIÓN (constantes, no dependen de la ronda) --------
    section("PARÁMETROS DEL MOTOR (supuestos de modelización — SUPUESTO PROPIO, no son campos de decisión por ronda)")
    p_opt = put("PARAM", "Utilización óptima de capacidad (mínimo de la curva en U)", "Todas las áreas", "%", "SUPUESTO")
    S.apply_cell(ws, p_opt, 7, value=0.85, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    p_k = put("PARAM", "Penalización por desvío de la utilización óptima (k)", "Todas las áreas", "-", "SUPUESTO")
    S.apply_cell(ws, p_k, 7, value=0.50, kind="input", numfmt='0.00', align=S.ALIGN_RIGHT, size=9)
    p_tasa_ap = put("PARAM", "Tasa de reducción de costo por duplicación de volumen (Wright's Law)", "Todas las tecnologías", "%", "SUPUESTO")
    S.apply_cell(ws, p_tasa_ap, 7, value=0.10, kind="input", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    p_capfab = {}
    for area, v in (("EE.UU.", 200000), ("China", 250000)):
        row = put("PARAM", "Capacidad por fábrica", area, "u.", "SUPUESTO")
        p_capfab[area] = row
        S.apply_cell(ws, row, 7, value=v, kind="input", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)
    p_tasadep = put("PARAM", "Tasa de depreciación (saldo decreciente)", "EE.UU. y China", "%", "REGLA ESTRUCTURAL CESIM",
                     "Fija en 10%, validada exacta contra RDOS R0→R1 (manual cap. 5.3, metodología de saldo decreciente).")
    S.apply_cell(ws, p_tasadep, 7, value=0.10, kind="historico", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    def P(row):
        return f"$G${row}"

    # -------- SECCIÓN A: CAPACIDAD OPERATIVA POR ÁREA --------
    section("SECCIÓN A · CAPACIDAD OPERATIVA POR ÁREA")
    cap_real0 = {"EE.UU.": 1_400_000, "China": 500_000}
    row_cap = {}
    row_prodtot = {}
    row_util = {}
    row_mult = {}
    row_costobasico = {}
    costo_basico_r0 = {"EE.UU.": 14721.1837038412, "China": 14033.8333357063}
    for area in AREAS:
        row_inv = put("A. Capacidad", "Inversión(+)/Desinversión(−) decidida (D2)", area, "u. capacidad", "DECISIÓN CADIZ (link)")
        row_entra = put("A. Capacidad", "Capacidad que entra en operación esta ronda (rezago +2)", area, "u.")
        row_sale = put("A. Capacidad", "Capacidad que sale de operación esta ronda (rezago +1)", area, "u.")
        row_c = put("A. Capacidad", "Capacidad operativa (cierre de ronda)", area, "u.", "OUTPUT")
        row_cap[area] = row_c
        row_pt = put("A. Capacidad", "Producción propia total del área (Σ tecnologías)", area, "u.")
        row_prodtot[area] = row_pt
        row_u = put("A. Capacidad", "% Utilización", area, "%", "OUTPUT")
        row_util[area] = row_u
        row_m = put("A. Capacidad", "Multiplicador de costo por utilización (curva en U)", area, "x")
        row_mult[area] = row_m
        row_cb = put("A. Capacidad", "Costo básico de producción del área", area, "USD/u.", "CÁLCULO (calibrado R0)",
                     "Ronda 0: calibrado = costo unitario real RDOS / multiplicador de la curva en U a la utilización real de R0. R1-R12: arrastra (editable si CESIM publica un costo básico nuevo).")
        row_costobasico[area] = row_cb

        for rn in RONDAS:
            c = col(rn)
            inv_ref = iref(inputs_idx, "D2 · PRODUCCIÓN", "Inversión (+) / Desinversión (−) en fábrica", area, rn)
            S.apply_cell(ws, row_inv, 7 + rn, value=f"=IF(ISNUMBER({inv_ref}),{inv_ref},0)", kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            if rn >= 2:
                c2 = col(rn - 2)
                f = f"=IF({c2}{row_inv}>0,{c2}{row_inv},0)"
            else:
                f = 0
            S.apply_cell(ws, row_entra, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            if rn >= 1:
                c1 = col(rn - 1)
                f = f"=IF({c1}{row_inv}<0,-{c1}{row_inv},0)"
            else:
                f = 0
            S.apply_cell(ws, row_sale, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            if rn == 0:
                S.apply_cell(ws, row_c, 7 + rn, value=cap_real0[area], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
            else:
                cprev = col(rn - 1)
                f = f"={cprev}{row_c}+{c}{row_entra}-{c}{row_sale}"
                S.apply_cell(ws, row_c, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
            partes = [iref(inputs_idx, "D2 · PRODUCCIÓN", "Producción propia", f"{area} / {t}", rn) for t in TECNOLOGIAS]
            f = "=(" + "+".join(f"IF(ISNUMBER({p}),{p},0)" for p in partes) + ")*1000"
            S.apply_cell(ws, row_pt, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            f = f"=IFERROR({c}{row_pt}/{c}{row_c},0)"
            S.apply_cell(ws, row_u, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=8)
            f = f"=1+{P(p_k)}*({c}{row_u}-{P(p_opt)})^2"
            S.apply_cell(ws, row_m, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=8)
            if rn == 0:
                S.apply_cell(ws, row_cb, 7 + rn, value=costo_basico_r0[area], kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            else:
                cprev = col(rn - 1)
                S.apply_cell(ws, row_cb, 7 + rn, value=f"={cprev}{row_cb}", kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
        r[0] += 1
    r[0] += 1

    # -------- SECCIÓN B: COSTO UNITARIO PROPIO POR ÁREA/TECNOLOGÍA (curva de aprendizaje) --------
    section("SECCIÓN B · COSTO UNITARIO DE PRODUCCIÓN PROPIA (curva de aprendizaje, Wright's Law sobre producción propia acumulada de CADIZ)")
    row_prod_u = {}
    row_acum = {}
    row_factor = {}
    row_costo_propio = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_p = put("B. Costo propio", "Producción propia de la ronda", dim, "u.")
            row_prod_u[(area, tech)] = row_p
            row_a = put("B. Costo propio", "Producción propia acumulada de CADIZ (cierre)", dim, "u.")
            row_acum[(area, tech)] = row_a
            row_f = put("B. Costo propio", "Factor de curva de aprendizaje", dim, "x")
            row_factor[(area, tech)] = row_f
            row_cu = put("B. Costo propio", "Costo unitario de producción propia", dim, "USD/u.", "OUTPUT")
            row_costo_propio[(area, tech)] = row_cu
            for rn in RONDAS:
                c = col(rn)
                prod_ref = iref(inputs_idx, "D2 · PRODUCCIÓN", "Producción propia", dim, rn)
                S.apply_cell(ws, row_p, 7 + rn, value=f"=IF(ISNUMBER({prod_ref}),{prod_ref},0)*1000", kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                if rn == 0:
                    f = f"={c}{row_p}"
                else:
                    cprev = col(rn - 1)
                    f = f"={cprev}{row_a}+{c}{row_p}"
                S.apply_cell(ws, row_a, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                base = col(0)
                f = (f"=IF(OR({base}${row_a}=0,{c}{row_a}=0),1,"
                     f"({c}{row_a}/{base}${row_a})^(LN(1-{P(p_tasa_ap)})/LN(2)))")
                S.apply_cell(ws, row_f, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_costobasico[area]}*{c}{row_mult[area]}*{c}{row_f}"
                S.apply_cell(ws, row_cu, 7 + rn, value=f, kind="output", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, bold=True, size=8)
            r[0] += 1
        r[0] += 1

    # -------- SECCIÓN C: DISPONIBILIDAD E INVENTARIO (costeo FIFO multicapa, Fase 2 cambio 1) --------
    section("SECCIÓN C · DISPONIBILIDAD E INVENTARIO — costeo FIFO multicapa (REGLA CESIM VERIFICADA, manual sección 5.2; replica MOTOR_VALIDADO_R2/build_engine_e2.py::layer_valor_formula()). 'Ventas' se completa después de _ENGINE_MERCADO.")
    # SEMILLA HISTÓRICA R0/R1 (DATO HISTÓRICO REAL, migrado sin alterar desde el motor validado anterior
    # -- 11_INVENTARIOS/E2 -- NO recalculado, para no alterar R0/R1). El inventario final real de R1 se
    # usa como "capa heredada" (heritage layer) de la que arranca el FIFO multicapa desde R2 en adelante
    # (ver nota de la Sección C y fifo_layer_terms()). Zero-dict por defecto para área/tecnología sin
    # producción/venta real.
    ZERO2 = {0: 0.0, 1: 0.0}
    ventas_hist_u = {t: dict(ZERO2) for t in [(a, te) for a in AREAS for te in TECNOLOGIAS]}
    disp_u_hist = {t: dict(ZERO2) for t in ventas_hist_u}
    disp_v_hist = {t: dict(ZERO2) for t in ventas_hist_u}
    cogs_hist = {t: dict(ZERO2) for t in ventas_hist_u}
    finu_hist = {t: dict(ZERO2) for t in ventas_hist_u}
    finv_hist = {t: dict(ZERO2) for t in ventas_hist_u}

    def seed(area, tech, r0=None, r1=None):
        k = (area, tech)
        if r0:
            disp_u_hist[k][0], disp_v_hist[k][0], ventas_hist_u[k][0], finu_hist[k][0], finv_hist[k][0] = r0
        if r1:
            disp_u_hist[k][1], disp_v_hist[k][1], ventas_hist_u[k][1], finu_hist[k][1], finv_hist[k][1] = r1

    # (disponible_u, disponible_v, ventas_u, inv_final_u, inv_final_v)
    seed("EE.UU.", "Combustión", r0=(1_230_000.0, 17_445_471_302.3214, 1_230_000.0, 0.0, 0.0),
                                   r1=(816_000.0, 10_146_142_867.2938, 816_000.0, 0.0, 0.0))
    seed("EE.UU.", "Híbrido", r1=(816_000.0, 12_725_389_944.3835, 520_789.0, 295_211.0, 4_947_976_593.14158))
    seed("China", "Combustión", r0=(800_000.0, 9_726_338_744.77752, 800_000.0, 0.0, 0.0),
                                  r1=(600_000.0, 6_910_348_459.1074, 600_000.0, 0.0, 0.0))
    seed("China", "Híbrido", r1=(270_000.0, 4_192_252_297.9577, 156_005.0, 113_995.0, 1_856_502_503.67271))
    row_terc_u = {}
    row_costo_terc = {}
    row_inv_ini_u = {}
    row_inv_ini_v = {}
    row_disp_u = {}
    row_disp_v = {}
    row_wac = {}
    row_ventas_u = {}
    row_cogs = {}
    row_inv_fin_u = {}
    row_inv_fin_v = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_t = put("C. Disponibilidad", "Producción tercerizada de la ronda (unidades)", dim, "u.")
            row_terc_u[(area, tech)] = row_t
            row_ct = put("C. Disponibilidad", "Costo unitario de producción tercerizada", dim, "USD/u.", "CONDICIÓN DE RONDA (link)")
            row_costo_terc[(area, tech)] = row_ct
            row_ini_u = put("C. Disponibilidad", "Inventario inicial (unidades)", dim, "u.")
            row_inv_ini_u[(area, tech)] = row_ini_u
            row_ini_v = put("C. Disponibilidad", "Inventario inicial (valor)", dim, "USD")
            row_inv_ini_v[(area, tech)] = row_ini_v
            row_du = put("C. Disponibilidad", "Total disponible (unidades)", dim, "u.", "OUTPUT")
            row_disp_u[(area, tech)] = row_du
            row_dv = put("C. Disponibilidad", "Total disponible (valor)", dim, "USD", "OUTPUT")
            row_disp_v[(area, tech)] = row_dv
            row_w = put("C. Disponibilidad", "Costo unitario promedio de venta (memo, resultante de FIFO)", dim, "USD/u.", "OUTPUT",
                        "Fase 2, cambio 1: MEMO informativo = COGS FIFO / unidades vendidas (promedio implícito de lectura rápida). Ya NO alimenta el cálculo -- el Costo de ventas y el Inventario final se calculan con FIFO multicapa por ronda de origen (ver 'Inventario final (valor)'), replicando MOTOR_VALIDADO_R2 / build_engine_e2.py::layer_valor_formula().")
            row_wac[(area, tech)] = row_w
            row_v = put("C. Disponibilidad", "Ventas (unidades)", dim, "u.", "OUTPUT (link _ENGINE_MERCADO)",
                        "R0/R1: histórico real (RDOS). R2-R12: link a _ENGINE_MERCADO, Sección C 'Ventas efectivas VALIDADAS' — completado por build_engine_produccion_part2().")
            row_ventas_u[(area, tech)] = row_v
            row_cg = put("C. Disponibilidad", "Costo de ventas de la ronda (COGS, FIFO multicapa)", dim, "USD", "OUTPUT")
            row_cogs[(area, tech)] = row_cg
            row_fu = put("C. Disponibilidad", "Inventario final (unidades)", dim, "u.", "OUTPUT")
            row_inv_fin_u[(area, tech)] = row_fu
            row_fv = put("C. Disponibilidad", "Inventario final (valor)", dim, "USD", "OUTPUT")
            row_inv_fin_v[(area, tech)] = row_fv

            for rn in RONDAS:
                c = col(rn)
                terc_ref = iref(inputs_idx, "D2 · PRODUCCIÓN", "Producción tercerizada", dim, rn)
                S.apply_cell(ws, row_t, 7 + rn, value=f"=IF(ISNUMBER({terc_ref}),{terc_ref},0)*1000", kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                ct_ref = iref(inputs_idx, "B. Condiciones", "Costo unitario de producción tercerizada", dim, rn)
                S.apply_cell(ws, row_ct, 7 + rn, value=f"=IF(ISNUMBER({ct_ref}),{ct_ref},0)", kind="link", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)

                k_at = (area, tech)
                if rn in (0, 1):
                    # DATO HISTÓRICO REAL (migrado sin alterar desde el motor validado anterior) --
                    # no se recalcula con el método de costeo simplificado (promedio ponderado) de
                    # este workbook, que rige SOLO desde Ronda 2 en adelante.
                    S.apply_cell(ws, row_ini_u, 7 + rn, value=0.0, kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_ini_v, 7 + rn, value=0.0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_du, 7 + rn, value=disp_u_hist[k_at][rn], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
                    S.apply_cell(ws, row_dv, 7 + rn, value=disp_v_hist[k_at][rn], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    wac_v = disp_v_hist[k_at][rn] / disp_u_hist[k_at][rn] if disp_u_hist[k_at][rn] else 0.0
                    S.apply_cell(ws, row_w, 7 + rn, value=wac_v, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_v, 7 + rn, value=ventas_hist_u[k_at][rn], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_cg, 7 + rn, value=cogs_hist[k_at][rn], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    S.apply_cell(ws, row_fu, 7 + rn, value=finu_hist[k_at][rn], kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_fv, 7 + rn, value=finv_hist[k_at][rn], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                else:
                    cprev = col(rn - 1)
                    S.apply_cell(ws, row_ini_u, 7 + rn, value=f"={cprev}{row_fu}", kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, row_ini_v, 7 + rn, value=f"={cprev}{row_fv}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    f = f"={c}{row_ini_u}+{c}{row_prod_u[(area,tech)]}+{c}{row_t}"
                    S.apply_cell(ws, row_du, 7 + rn, value=f, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
                    f = f"={c}{row_ini_v}+{c}{row_prod_u[(area,tech)]}*{c}{row_costo_propio[(area,tech)]}+{c}{row_t}*{c}{row_ct}"
                    S.apply_cell(ws, row_dv, 7 + rn, value=f, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    # Ventas (unidades): completado por build_engine_produccion_part2() con el link a
                    # _ENGINE_MERCADO. Se deja en 0 acá (placeholder), pero el FIFO de abajo referencia
                    # la CELDA (no el valor python), así que recalcula bien una vez parcheada.
                    S.apply_cell(ws, row_v, 7 + rn, value=0, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    # Fase 2, cambio 1: FIFO multicapa (reemplaza WAC). Capa heredada = Inventario final
                    # REAL de Ronda 1 (histórico, no recalculado); capas R2..rn = producción propia y
                    # tercerizada de cada ronda a su propio costo de origen, consumidas en orden FIFO.
                    her_qty = finu_hist[k_at][1]
                    her_cost = (finv_hist[k_at][1] / her_qty) if her_qty else 0.0
                    f_fifo = fifo_layer_terms(row_prod_u[(area, tech)], row_costo_propio[(area, tech)],
                                               row_t, row_ct, row_v, her_qty, her_cost, 2, rn)
                    S.apply_cell(ws, row_fv, 7 + rn, value="=" + f_fifo, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    f = f"={c}{row_dv}-{c}{row_fv}"
                    S.apply_cell(ws, row_cg, 7 + rn, value=f, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    f = f"={c}{row_du}-{c}{row_v}"
                    S.apply_cell(ws, row_fu, 7 + rn, value=f, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    f = f"=IFERROR({c}{row_cg}/{c}{row_v},0)"
                    S.apply_cell(ws, row_w, 7 + rn, value=f, kind="output", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            r[0] += 1
        r[0] += 1

    # -------- SECCIÓN D: DEPRECIACIÓN Y ACTIVO FIJO (independiente de ventas) --------
    section("SECCIÓN D · ACTIVO FIJO Y DEPRECIACIÓN (rollforward, saldo decreciente 10%)")
    seed_af = {
        "EE.UU.": {"bruto0": 9_045_615_600.0, "dep0": 1_005_068_400.0, "pago1": 14_000_000.0, "bruto1": 9_059_615_600.0, "dep1": 904_561_560.0, "depacum1": 904_561_560.0, "resid1": 8_155_054_040.0},
        "China": {"bruto0": 1_782_588_600.0, "dep0": 198_065_400.0, "pago1": 4_000_000.0, "bruto1": 1_786_588_600.0, "dep1": 178_258_860.0, "depacum1": 178_258_860.0, "resid1": 1_608_329_740.0},
    }
    row_af_bruto = {}
    row_af_depacum = {}
    row_af_resid = {}
    row_af_dep_ronda = {}
    row_pago_inv = {}
    row_cobro_desinv = {}
    for area in AREAS:
        seed = seed_af[area]
        row_costofab = put("D. Activo fijo", "Costo de inversión por fábrica", area, "USD", "CONDICIÓN DE RONDA (link)")
        row_preciofab = put("D. Activo fijo", "Precio de venta por fábrica", area, "USD", "CONDICIÓN DE RONDA (link)")
        row_nfab_dec = put("D. Activo fijo", "N° equivalente de fábricas invertidas(+)/desinvertidas(−)", area, "fábricas")
        row_montoinv = put("D. Activo fijo", "Monto de inversión decidida esta ronda (memo)", area, "USD")
        row_montodesinv = put("D. Activo fijo", "Monto de desinversión decidida esta ronda (memo)", area, "USD")
        row_pago = put("D. Activo fijo", "Pago por inversión (impacto de caja, rezago +1 ronda)", area, "USD", "OUTPUT")
        row_pago_inv[area] = row_pago
        row_cobro = put("D. Activo fijo", "Cobro por desinversión (impacto de caja, rezago +1 ronda)", area, "USD", "OUTPUT")
        row_cobro_desinv[area] = row_cobro
        row_bruto_ap = put("D. Activo fijo", "Valor bruto (apertura)", area, "USD")
        row_altas = put("D. Activo fijo", "Altas brutas de la ronda", area, "USD")
        row_bajas_b = put("D. Activo fijo", "Bajas brutas de la ronda", area, "USD")
        row_bruto_ci = put("D. Activo fijo", "Valor bruto (cierre)", area, "USD")
        row_af_bruto[area] = row_bruto_ci
        row_dep = put("D. Activo fijo", "Depreciación de la ronda", area, "USD", "OUTPUT")
        row_af_dep_ronda[area] = row_dep
        row_depacum_ap = put("D. Activo fijo", "Depreciación acumulada (apertura)", area, "USD")
        row_bajadep = put("D. Activo fijo", "Baja de depreciación acumulada por desinversión", area, "USD")
        row_depacum_ci = put("D. Activo fijo", "Depreciación acumulada (cierre)", area, "USD")
        row_af_depacum[area] = row_depacum_ci
        row_resid_ap = put("D. Activo fijo", "Valor residual / neto (apertura)", area, "USD")
        row_bajas_n = put("D. Activo fijo", "Bajas netas de la ronda", area, "USD")
        row_resid_ci = put("D. Activo fijo", "Valor residual / neto (cierre)", area, "USD", "OUTPUT")
        row_af_resid[area] = row_resid_ci

        row_decision = row_idx[f"A. Capacidad||Inversión(+)/Desinversión(−) decidida (D2)||{area}"]
        for rn in RONDAS:
            c = col(rn)
            cf_ref = iref(inputs_idx, "B. Condiciones", "Costo de inversión por fábrica", area, rn)
            pf_ref = iref(inputs_idx, "B. Condiciones", "Precio de venta por fábrica (desinversión)", area, rn)
            S.apply_cell(ws, row_costofab, 7 + rn, value=f"=IF(ISNUMBER({cf_ref}),{cf_ref},0)", kind="link", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, row_preciofab, 7 + rn, value=f"=IF(ISNUMBER({pf_ref}),{pf_ref},0)", kind="link", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            f = f"=IFERROR({c}{row_decision}/{P(p_capfab[area])},0)"
            S.apply_cell(ws, row_nfab_dec, 7 + rn, value=f, kind="calculo", numfmt='#,##0.00', align=S.ALIGN_RIGHT, size=8)

            if rn == 0:
                S.apply_cell(ws, row_montoinv, 7 + rn, value="n/a (R0)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_montodesinv, 7 + rn, value="n/a (R0)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
            else:
                f = f"=IF({c}{row_nfab_dec}>0,{c}{row_nfab_dec}*{c}{row_costofab},0)"
                S.apply_cell(ws, row_montoinv, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"=IF({c}{row_nfab_dec}<0,-{c}{row_nfab_dec}*{c}{row_preciofab},0)"
                S.apply_cell(ws, row_montodesinv, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if rn == 0:
                S.apply_cell(ws, row_pago, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_cobro, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            elif rn == 1:
                S.apply_cell(ws, row_pago, 7 + rn, value=seed["pago1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8,
                             bold=False)
                S.apply_cell(ws, row_cobro, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            else:
                c1 = col(rn - 1)
                S.apply_cell(ws, row_pago, 7 + rn, value=f"={c1}{row_montoinv}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_cobro, 7 + rn, value=f"={c1}{row_montodesinv}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if rn == 0:
                S.apply_cell(ws, row_bruto_ap, 7 + rn, value="n/a (dato inicial)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_altas, 7 + rn, value="n/a (R0)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_bajas_b, 7 + rn, value="n/a (R0)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_bruto_ci, 7 + rn, value=seed["bruto0"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                S.apply_cell(ws, row_dep, 7 + rn, value=seed["dep0"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_depacum_ap, 7 + rn, value="n/a (R0)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_bajadep, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_depacum_ci, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_resid_ap, 7 + rn, value="n/a (dato inicial)", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_bajas_n, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_resid_ci, 7 + rn, value=seed["bruto0"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            elif rn == 1:
                S.apply_cell(ws, row_bruto_ap, 7 + rn, value=seed["bruto0"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_altas, 7 + rn, value=seed["pago1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_bajas_b, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_bruto_ci, 7 + rn, value=seed["bruto1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                S.apply_cell(ws, row_dep, 7 + rn, value=seed["dep1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_depacum_ap, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_bajadep, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_depacum_ci, 7 + rn, value=seed["depacum1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_resid_ap, 7 + rn, value=seed["bruto0"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_bajas_n, 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_resid_ci, 7 + rn, value=seed["resid1"], kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            else:
                cprev = col(rn - 1)
                f = f"={cprev}{row_bruto_ci}"
                S.apply_cell(ws, row_bruto_ap, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_pago}"
                S.apply_cell(ws, row_altas, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                c1 = col(rn - 1)
                cap_row = row_cap[area]
                frac_baja = f'IF({c1}{row_decision}<0,-{c1}{row_decision}/{c1}{cap_row},0)'
                f = f"={c}{row_bruto_ap}*({frac_baja})"
                S.apply_cell(ws, row_bajas_b, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_bruto_ap}+{c}{row_altas}-{c}{row_bajas_b}"
                S.apply_cell(ws, row_bruto_ci, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                f = f"={cprev}{row_resid_ci}"
                S.apply_cell(ws, row_resid_ap, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_resid_ap}*{P(p_tasadep)}"
                S.apply_cell(ws, row_dep, 7 + rn, value=f, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                f = f"={c}{row_resid_ap}*({frac_baja})"
                S.apply_cell(ws, row_bajas_n, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_bajas_b}-{c}{row_bajas_n}"
                S.apply_cell(ws, row_bajadep, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={cprev}{row_depacum_ci}"
                S.apply_cell(ws, row_depacum_ap, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_depacum_ap}+{c}{row_dep}-{c}{row_bajadep}"
                S.apply_cell(ws, row_depacum_ci, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                f = f"={c}{row_bruto_ci}-{c}{row_depacum_ci}"
                S.apply_cell(ws, row_resid_ci, 7 + rn, value=f, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        r[0] += 1
    return ws, row_idx, r, dict(cap=row_cap, prodtot=row_prodtot, util=row_util, mult=row_mult,
                                  costobasico=row_costobasico, costo_propio=row_costo_propio,
                                  disp_u=row_disp_u, disp_v=row_disp_v, wac=row_wac, ventas_u=row_ventas_u,
                                  cogs=row_cogs, inv_fin_u=row_inv_fin_u, inv_fin_v=row_inv_fin_v,
                                  costo_terc=row_costo_terc, prod_u=row_prod_u,
                                  terc_u=row_terc_u, af_bruto=row_af_bruto, af_depacum=row_af_depacum,
                                  af_resid=row_af_resid, af_dep_ronda=row_af_dep_ronda,
                                  pago_inv=row_pago_inv, cobro_desinv=row_cobro_desinv,
                                  params=dict(opt=p_opt, k=p_k, tasa_ap=p_tasa_ap, capfab=p_capfab, tasadep=p_tasadep))


# ======================================================================================
# _ENGINE_MERCADO — Mercado, demanda, asignación multi-origen por prioridad logística,
# ventas efectivas, exportaciones, transporte/aranceles y transfer pricing. (Motor E3, MISMA
# lógica de asignación multi-origen del proyecto anterior — ver README, no se simplificó.)
# ======================================================================================
def build_engine_mercado(wb, inputs_idx, prod_out, sens_rows=None, rdos_files=None, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(ENGINE_MKT_SHEET)
    ncols = 6 + len(RONDAS)
    S.set_col_widths(ws, [18, 50, 26, 10, 10, 8] + [12] * len(RONDAS))
    S.style_title_row(ws, 1, "_ENGINE_MERCADO — Mercado, demanda, asignación multi-origen por prioridad logística (D6), ventas efectivas, exportaciones y transfer pricing (D7)", ncols)
    header_row = 3
    headers = ["Bloque", "Variable", "Dimensión", "Unidad", "Tipo", "Nota"] + [f"R{r}" for r in RONDAS]
    S.style_header_row(ws, header_row, headers)
    r = [header_row + 1]
    row_idx = {}

    def put(bloque, variable, dimension, unidad="", tipo="CÁLCULO", nota=""):
        rr = r[0]
        row_idx[f"{bloque}||{variable}||{dimension}"] = rr
        S.apply_cell(ws, rr, 1, value=bloque, kind="calculo", align=S.ALIGN_LEFT, size=8, bold=True)
        S.apply_cell(ws, rr, 2, value=variable, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=8)
        S.apply_cell(ws, rr, 3, value=dimension, kind="calculo", align=S.ALIGN_LEFT, size=8)
        S.apply_cell(ws, rr, 4, value=unidad, kind="calculo", align=S.ALIGN_CENTER, size=7)
        S.apply_cell(ws, rr, 5, value=tipo, kind="calculo", align=S.ALIGN_CENTER, size=7, bold=True)
        S.apply_cell(ws, rr, 6, value=nota, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=6, italic=True)
        r[0] += 1
        return rr

    def section(title):
        S.style_section_row(ws, r[0], title, ncols)
        r[0] += 1

    # Ronda N.N (Fase 4, frontera dinámica): Tamaño de mercado / Demanda estimada CADIZ para
    # CUALQUIER ronda con RDOS -- ver extract_mercado_real() (verificado bottom-up, reemplaza los
    # diccionarios tam_hist/dem_real que antes solo cubrían R0/R1 a mano).
    tam_real, dem_real = extract_mercado_real(rdos_files or {})
    # [DATO EXTERNO NO VERIFICADO EN ARCHIVO PROPIO] Observado directamente por el equipo en la
    # plataforma CESIM al momento de decidir Ronda 2 (no reconstruible desde los archivos del
    # proyecto). Se conserva solo como registro histórico -- YA NO se usa como base de cálculo
    # (ver nota extendida más abajo, Sección A1).
    BASE_R1_AJUSTADA = {"EE.UU.": 5_193_000.0, "China": 3_808_000.0, "Europa": 6_647_000.0}

    # -------- SECCIÓN A1: TAMAÑO DE MERCADO TOTAL --------
    section("SECCIÓN A1 · TAMAÑO DE MERCADO TOTAL (todas las empresas, por mercado)")
    # Ronda N.N (Fase 4, frontera dinámica): tam_real/dem_real ahora vienen de extract_mercado_real()
    # (verificado bottom-up contra R0-R3: Tamaño de mercado = Σ Ventas CADIZ / Cuota total CADIZ,
    # exacto a 10 decimales contra los valores antes hardcodeados a mano para R0/R1; Demanda estimada
    # CADIZ = fila "Demanda, miles unidades" propia de CADIZ dentro de cada tecnología -- 0 mismatches
    # en las 24 combinaciones mercado×tecnología×ronda de R0/R1). Reemplaza los diccionarios
    # tam_hist/dem_real (solo R0/R1, copiados a mano) -- ahora cualquier ronda con RDOS entra igual.
    #
    # BASE_R1_AJUSTADA (la fila "R1 — BASE AJUSTADA" de más abajo, hoy retirada del cálculo): era un
    # [DATO EXTERNO NO VERIFICADO] observado a mano en la plataforma CESIM al decidir Ronda 2 (5.193.000
    # vs. el Tamaño de mercado R1 derivado del RDOS propio, 5.017.355 -- ~3,5% de diferencia), no
    # reconstruible desde ningún archivo. Esa corrección puntual NO se generaliza (no es automatizable:
    # requería mirar la plataforma en vivo en un momento específico) -- de acá en más, la base de
    # crecimiento de CUALQUIER transición Ronda real→Ronda siguiente usa el Tamaño de mercado derivado
    # de RDOS puro. LIMITACIÓN DOCUMENTADA: la proyección de Ronda 2 en este archivo pudo diferir en
    # ~3,5% del tamaño de mercado por este motivo puntual; no vuelve a aplicar desde acá en más.
    n_next = (max(rondas_reales) + 1) if rondas_reales else 0
    row_tam = {}
    row_tam_base = {}
    for mercado in MERCADOS:
        row_b = put("A1. Tamaño de mercado", "Tamaño de mercado R1 — BASE AJUSTADA (histórico, ya no usada en el cálculo)", mercado, "u.", "DATO EXTERNO (no verificado en archivo propio)",
                    "[RETIRADA DEL CÁLCULO, Fase 4] Observado por el equipo en la plataforma CESIM al decidir R2 -- no reconstruible ni generalizable a otras rondas. Se conserva solo como registro histórico.")
        row_tam_base[mercado] = row_b
        row_t = put("A1. Tamaño de mercado", "Tamaño de mercado total estimado", mercado, "u.", "CÁLCULO",
                    "Ronda con RDOS: derivado de RDOS real (ventas CADIZ / cuota total CADIZ). Ronda sin RDOS: ronda anterior × (1+crecimiento D1) -- ya sea que la ronda anterior sea real o proyectada.")
        row_tam[mercado] = row_t
        for rn in RONDAS:
            c = col(rn)
            S.apply_cell(ws, row_b, 7 + rn, value=(BASE_R1_AJUSTADA[mercado] if rn == 1 else "n/a"), kind=("plain" if rn != 1 else "historico"), align=S.ALIGN_CENTER, italic=(rn != 1), size=8 if rn == 1 else 7, numfmt=(S.NUM_UNITS if rn == 1 else None))
            growth_ref = iref(inputs_idx, "B. Condiciones", "Crecimiento de mercado esperado (vs. ronda anterior)", mercado, rn)
            # [SUPUESTO PROPIO] Matriz de sensibilidad: aplica desde la primera ronda a decidir en
            # adelante (n_next) -- antes fijo a "desde R3" porque R2 ya estaba cerrada/decidida cuando
            # se creó esta herramienta. Con Escenario Activo=BASE (default) el factor es 0%, sin efecto.
            sens_mult = f"*(1+{sens_ref(sens_rows['factor_crec'])})" if (sens_rows and rn >= n_next) else ""
            if rn in rondas_reales:
                v_tam = tam_real.get(mercado, {}).get(rn)
                S.apply_cell(ws, row_t, 7 + rn, value=(round(v_tam) if v_tam is not None else "n/d"), kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            else:
                cprev = col(rn - 1)
                f = f"={cprev}{row_t}*(1+{growth_ref}){sens_mult}"
                S.apply_cell(ws, row_t, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
    r[0] += 1

    # -------- SECCIÓN A2: DEMANDA ESTIMADA CADIZ --------
    section("SECCIÓN A2 · DEMANDA ESTIMADA DE CADIZ POR MERCADO Y TECNOLOGÍA (proyección propia, no garantizada por CESIM)")
    row_demanda = {}
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row_d = put("A2. Demanda CADIZ", "Demanda estimada CADIZ", dim, "u.", "CÁLCULO",
                        "Ronda con RDOS: histórico real (RDOS). Ronda sin RDOS: Tamaño de mercado (A1) × Cuota Resultante SOBRE EL TOTAL (D1c = Mix de "
                        "Industria D1a × Cuota sobre la Tecnología D1b) — SUPUESTO, no garantizado por CESIM. Ver Adenda 22.")
            row_demanda[(mercado, tech)] = row_d
            for rn in RONDAS:
                c = col(rn)
                if rn in rondas_reales:
                    v_dem = dem_real.get((mercado, tech), {}).get(rn, 0.0)
                    S.apply_cell(ws, row_d, 7 + rn, value=v_dem, kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                else:
                    cuota_ref = iref(inputs_idx, "D1c · DEMANDA (RESULTANTE)", "Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)", dim, rn)
                    # [SUPUESTO PROPIO] Cuota_efectiva = Cuota_objetivo × (1−Factor_Competencia_Activo),
                    # misma guarda "desde n_next en adelante" que sens_mult arriba.
                    comp_mult = f"*(1-{sens_ref(sens_rows['factor_comp'])})" if (sens_rows and rn >= n_next) else ""
                    f = f"={c}{row_tam[mercado]}*{cuota_ref}{comp_mult}"
                    S.apply_cell(ws, row_d, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
    r[0] += 1

    # -------- SECCIÓN B: OFERTA DISPONIBLE (link a _ENGINE_PRODUCCION) --------
    section("SECCIÓN B · OFERTA DISPONIBLE POR ÁREA Y TECNOLOGÍA (link a _ENGINE_PRODUCCION, Total disponible)")
    row_oferta = {}
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_o = put("B. Oferta", "Oferta disponible (total disponible, previo a ventas)", dim, "u.")
            row_oferta[(area, tech)] = row_o
            for rn in RONDAS:
                ref = cref(ENGINE_PROD_SHEET, prod_out["disp_u"][(area, tech)], rn)
                S.apply_cell(ws, row_o, 7 + rn, value=f"={ref}", kind="link", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
    r[0] += 1

    # -------- SECCIÓN C: ASIGNACIÓN MULTI-ORIGEN POR PRIORIDAD LOGÍSTICA --------
    section("SECCIÓN C · ASIGNACIÓN MULTI-ORIGEN POR PRIORIDAD LOGÍSTICA (prorrateo proporcional si compiten 2 orígenes — SUPUESTO PROPIO, manual cap. 9.1; MISMA lógica del motor anterior)")
    slot_mercado_row, slot_ofrem_row, slot_comb_row, slot_asig_row = {}, {}, {}, {}
    dem_rem_row = {}
    row_ventas_at, row_excl_at = {}, {}

    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            for slot in (1, 2, 3):
                row_m = put("C. Asignación", f"Mercado en prioridad {slot}", dim, "texto")
                slot_mercado_row[(area, tech, slot)] = row_m
            for rn in RONDAS:
                c = col(rn)
                rank_cell = {m: iref(inputs_idx, "D6 · LOGÍSTICA", "Orden de prioridad de entrega (1=primero, 3=último)", f"Área {area} / {tech} → mercado {m}", rn) for m in MERCADOS}
                for slot in (1, 2, 3):
                    row_m = slot_mercado_row[(area, tech, slot)]
                    f_m = (f'=IF({rank_cell["EE.UU."]}="{slot}","EE.UU.",'
                           f'IF({rank_cell["China"]}="{slot}","China",'
                           f'IF({rank_cell["Europa"]}="{slot}","Europa","(sin definir)")))')
                    S.apply_cell(ws, row_m, 7 + rn, value=f_m, kind="calculo", align=S.ALIGN_CENTER, size=7)

    other_area = {AREAS[0]: AREAS[1], AREAS[1]: AREAS[0]}
    for tech in TECNOLOGIAS:
        for mercado in MERCADOS:
            for slot in (2, 3, 4):
                label = ("Demanda insatisfecha estimada (remanente final)" if slot == 4 else f"Demanda remanente de este mercado antes de la prioridad {slot}")
                row_dr = put("C. Asignación", label, f"{mercado} / {tech}", "u.")
                dem_rem_row[(mercado, tech, slot)] = row_dr

        for slot in (1, 2, 3):
            for area in AREAS:
                dim = f"{area} / {tech}"
                row_of = put("C. Asignación", f"Oferta propia remanente al iniciar prioridad {slot}", dim, "u.")
                row_cb = put("C. Asignación", f"Oferta combinada compitiendo por el mercado de prioridad {slot}", dim, "u.")
                row_as = put("C. Asignación", f"Unidades asignadas en prioridad {slot}", dim, "u.")
                slot_ofrem_row[(area, tech, slot)] = row_of
                slot_comb_row[(area, tech, slot)] = row_cb
                slot_asig_row[(area, tech, slot)] = row_as

            for rn in RONDAS:
                c = col(rn)
                for area in AREAS:
                    oarea = other_area[area]
                    row_m = slot_mercado_row[(area, tech, slot)]
                    row_m_o = slot_mercado_row[(oarea, tech, slot)]
                    row_of = slot_ofrem_row[(area, tech, slot)]
                    row_of_o = slot_ofrem_row[(oarea, tech, slot)]
                    row_cb = slot_comb_row[(area, tech, slot)]
                    row_as = slot_asig_row[(area, tech, slot)]
                    if slot == 1:
                        f_of = f"={c}{row_oferta[(area, tech)]}"
                    else:
                        prev_of = slot_ofrem_row[(area, tech, slot - 1)]
                        prev_as = slot_asig_row[(area, tech, slot - 1)]
                        f_of = f"=MAX({c}{prev_of}-{c}{prev_as},0)"
                    S.apply_cell(ws, row_of, 7 + rn, value=f_of, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    f_cb = f"={c}{row_of}+IF({c}{row_m}={c}{row_m_o},{c}{row_of_o},0)"
                    S.apply_cell(ws, row_cb, 7 + rn, value=f_cb, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                    if slot == 1:
                        demrem = "+".join(f'IF({c}{row_m}="{m}",{c}{row_demanda[(m, tech)]},0)' for m in MERCADOS)
                    else:
                        demrem = "+".join(f'IF({c}{row_m}="{m}",{c}{dem_rem_row[(m, tech, slot)]},0)' for m in MERCADOS)
                    f_as = f"=IF({c}{row_cb}<=({demrem}),{c}{row_of},IFERROR({c}{row_of}/{c}{row_cb}*({demrem}),0))"
                    S.apply_cell(ws, row_as, 7 + rn, value=f_as, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)

            if slot < 3:
                for mercado in MERCADOS:
                    row_next = dem_rem_row[(mercado, tech, slot + 1)]
                    for rn in RONDAS:
                        c = col(rn)
                        base = f"{c}{row_demanda[(mercado, tech)]}" if slot == 1 else f"{c}{dem_rem_row[(mercado, tech, slot)]}"
                        asignado = "+".join(f'IF({c}{slot_mercado_row[(area, tech, slot)]}="{mercado}",{c}{slot_asig_row[(area, tech, slot)]},0)' for area in AREAS)
                        f = f"=MAX({base}-({asignado}),0)"
                        S.apply_cell(ws, row_next, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)

        for mercado in MERCADOS:
            row_final = dem_rem_row[(mercado, tech, 4)]
            for rn in RONDAS:
                c = col(rn)
                base = f"{c}{dem_rem_row[(mercado, tech, 3)]}"
                asignado = "+".join(f'IF({c}{slot_mercado_row[(area, tech, 3)]}="{mercado}",{c}{slot_asig_row[(area, tech, 3)]},0)' for area in AREAS)
                f = f"=MAX({base}-({asignado}),0)"
                S.apply_cell(ws, row_final, 7 + rn, value=f, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)

    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_vt = put("C. Asignación", "Ventas efectivas VALIDADAS de esta área/tecnología", dim, "u.", "OUTPUT",
                        "Σ de lo asignado en las 3 prioridades, excluyendo unidades cuyo mercado quedó en estado ERROR (Sección D). Alimenta _ENGINE_PRODUCCION (ventas/COGS/inventario).")
            row_excl = put("C. Asignación", "Unidades excluidas por sobre-asignación sin regla de reparto (auditoría)", dim, "u.", "ALERTA",
                          "Debería ser SIEMPRE 0 por construcción algebraica del prorrateo; un valor ≠0 indica una inconsistencia a investigar.")
            row_ventas_at[(area, tech)] = row_vt
            row_excl_at[(area, tech)] = row_excl
    r[0] += 1

    # -------- SECCIÓN D: VENTAS EFECTIVAS AGREGADAS POR MERCADO x TECNOLOGÍA (control) --------
    section("SECCIÓN D · VENTAS EFECTIVAS AGREGADAS POR MERCADO Y TECNOLOGÍA (control bloqueante de sobre-asignación entre áreas)")
    row_estado_d = {}
    row_vef = {}
    row_insat = {}
    for mercado in MERCADOS:
        for tech in TECNOLOGIAS:
            dim = f"{mercado} / {tech}"
            row_tot = put("D. Ventas por mercado", "Oferta total asignada a este mercado (todas las áreas)", dim, "u.")
            row_alerta = put("D. Ventas por mercado", "Control: oferta asignada > demanda estimada", dim, "u.", "ALERTA")
            row_estado = put("D. Ventas por mercado", "Estado de validez (control bloqueante)", dim, "texto", "ALERTA")
            row_v = put("D. Ventas por mercado", "Ventas efectivas del mercado (capadas por demanda; 0 si ERROR)", dim, "u.", "OUTPUT")
            row_i = put("D. Ventas por mercado", "Demanda insatisfecha estimada", dim, "u.", "OUTPUT")
            row_estado_d[(mercado, tech)] = row_estado
            row_vef[(mercado, tech)] = row_v
            row_insat[(mercado, tech)] = row_i
            for rn in RONDAS:
                c = col(rn)
                terms = []
                for area in AREAS:
                    terms.append("+".join(f'IF({c}{slot_mercado_row[(area, tech, s)]}="{mercado}",{c}{slot_asig_row[(area, tech, s)]},0)' for s in (1, 2, 3)))
                f = "=" + "+".join(f"({t})" for t in terms)
                S.apply_cell(ws, row_tot, 7 + rn, value=f, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                f = f"=MAX({c}{row_tot}-{c}{row_demanda[(mercado, tech)]},0)"
                S.apply_cell(ws, row_alerta, 7 + rn, value=f, kind="alerta", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
                f_estado = f'=IF({c}{row_alerta}>0.5,"ERROR — sobre-asignación","OK")'
                S.apply_cell(ws, row_estado, 7 + rn, value=f_estado, kind="alerta", align=S.ALIGN_CENTER, size=7, bold=True)
                f = f'=IF({c}{row_estado}="OK",MIN({c}{row_tot},{c}{row_demanda[(mercado, tech)]}),0)'
                S.apply_cell(ws, row_v, 7 + rn, value=f, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
                f = f'=IF({c}{row_estado}="OK",MAX({c}{row_demanda[(mercado, tech)]}-{c}{row_tot},0),0)'
                S.apply_cell(ws, row_i, 7 + rn, value=f, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
    r[0] += 1

    # -------- Pasada 2 de Sección C: completar Ventas VALIDADAS / Excluidas --------
    for area in AREAS:
        for tech in TECNOLOGIAS:
            row_vt = row_ventas_at[(area, tech)]
            row_excl = row_excl_at[(area, tech)]
            for rn in RONDAS:
                c = col(rn)
                valid_terms = []
                for s in (1, 2, 3):
                    row_m = slot_mercado_row[(area, tech, s)]
                    row_a = slot_asig_row[(area, tech, s)]
                    ok_lookup = "+".join(f'IF({c}{row_m}="{m}",IF({c}{row_estado_d[(m, tech)]}="OK",1,0),0)' for m in MERCADOS)
                    valid_terms.append(f'({c}{row_a}*({ok_lookup}))')
                f_vt = "=" + "+".join(valid_terms)
                S.apply_cell(ws, row_vt, 7 + rn, value=f_vt, kind="output", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, bold=True, size=8)
                raw_sum = "+".join(f"{c}{slot_asig_row[(area, tech, s)]}" for s in (1, 2, 3))
                f_excl = f"=({raw_sum})-{c}{row_vt}"
                S.apply_cell(ws, row_excl, 7 + rn, value=f_excl, kind="alerta", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)

    # -------- SECCIÓN E: EXPORTACIÓN, TRANSPORTE, ARANCELES, TRANSFER PRICING --------
    section("SECCIÓN E · EXPORTACIÓN, TRANSPORTE, ARANCELES Y PRECIO DE TRANSFERENCIA (por par área origen → mercado destino)")
    row_export_u, row_costo_dir, row_pt, row_costo_tot_u, row_costo_tot = {}, {}, {}, {}, {}
    for (area, mercado) in PARES_EXPORT:
        dim = f"{area} → {mercado}"
        row_u = put("E. Exportación", "Unidades exportadas (todas las tecnologías)", dim, "u.")
        row_cd = put("E. Exportación", "Costo unitario de origen (promedio ponderado por mix exportado)", dim, "USD/u.")
        row_mu = put("E. Exportación", "Multiplicador de precio de transferencia (D7)", dim, "x")
        row_p = put("E. Exportación", "Precio de transferencia", dim, "USD/u.")
        row_car_u = put("E. Exportación", "Costo de la característica por unidad (base arancel, promedio ponderado por mix exportado)", dim, "USD/u.",
                        nota="Manual: 'las tarifas ad valorem se basan en el valor del producto definido por el precio de transferencia "
                             "y los costos de la característica' -- características/costo del país DESTINO (mercado), ponderado por "
                             "el mismo mix tecnológico exportado que Costo unitario de origen.")
        row_esp = put("E. Exportación", "Arancel específico por unidad", dim, "USD/u.")
        row_adv_pct = put("E. Exportación", "Arancel ad valorem (%)", dim, "%")
        row_adv_usd = put("E. Exportación", "Arancel ad valorem, USD/u.", dim, "USD/u.",
                        nota="REGLA VERIFICADA (manual): base = Precio de transferencia + Costo de la característica por unidad, "
                             "no solo Precio de transferencia.")
        row_ctu = put("E. Exportación", "Costo total transporte + arancel por unidad", dim, "USD/u.")
        row_ct = put("E. Exportación", "Costo total transporte + arancel de la ronda", dim, "USD", "OUTPUT")
        row_export_u[(area, mercado)] = row_u
        row_pt[(area, mercado)] = row_p
        row_costo_tot_u[(area, mercado)] = row_ctu
        row_costo_tot[(area, mercado)] = row_ct

        for rn in RONDAS:
            c = col(rn)
            terms, costo_terms = [], []
            for tech in TECNOLOGIAS:
                raw = "+".join(f'IF({c}{slot_mercado_row[(area, tech, s)]}="{mercado}",{c}{slot_asig_row[(area, tech, s)]},0)' for s in (1, 2, 3))
                ok = f'IF({c}{row_estado_d[(mercado, tech)]}="OK",1,0)'
                terms.append(f'(({raw})*({ok}))')
                costo_terms.append((terms[-1], cref(ENGINE_PROD_SHEET, prod_out["costo_propio"][(area, tech)], rn)))
            f_u = "=" + "+".join(f"({t})" for t in terms)
            S.apply_cell(ws, row_u, 7 + rn, value=f_u, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
            num = "+".join(f"({t})*{cref_}" for t, cref_ in costo_terms)
            S.apply_cell(ws, row_cd, 7 + rn, value=f"=IFERROR(({num})/{c}{row_u},0)", kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            # Base del arancel ad valorem (manual): Precio de transferencia + Costo de la característica
            # por unidad -- características/costo del país DESTINO (mercado), ponderado por el mismo mix
            # exportado (terms[i], en el mismo orden de TECNOLOGIAS) que ya usa Costo unitario de origen.
            costo_car_dest_ref = iref(inputs_idx, "B. Condiciones", "Costo de la característica", mercado, rn)
            car_num_terms = []
            for i, tech in enumerate(TECNOLOGIAS):
                feat_ref = iref(inputs_idx, "D5 · MARKETING", "Características ofrecidas (de las disponibles, tope 10)", f"{mercado} / {tech}", rn)
                car_num_terms.append(f'(({terms[i]})*IF(ISNUMBER({feat_ref}),{feat_ref},0)*IF(ISNUMBER({costo_car_dest_ref}),{costo_car_dest_ref},0))')
            f_car_u = f"=IFERROR(({'+'.join(car_num_terms)})/{c}{row_u},0)"
            S.apply_cell(ws, row_car_u, 7 + rn, value=f_car_u, kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            mult_ref = iref(inputs_idx, "D7 · IMPUESTOS", "Multiplicador de precio de transferencia", dim, rn)
            if rn in (0, 1):
                S.apply_cell(ws, row_mu, 7 + rn, value="n/a", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
                S.apply_cell(ws, row_p, 7 + rn, value="n/a", kind="plain", align=S.ALIGN_CENTER, italic=True, size=7)
            else:
                S.apply_cell(ws, row_mu, 7 + rn, value=f"={mult_ref}", kind="calculo", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_p, 7 + rn, value=f"={c}{row_cd}*{c}{row_mu}", kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            esp_ref = iref(inputs_idx, "B. Condiciones", "Arancel específico por unidad", dim, rn)
            adv_ref = iref(inputs_idx, "B. Condiciones", "Arancel ad valorem (%)", dim, rn)
            S.apply_cell(ws, row_esp, 7 + rn, value=f"=IF(ISNUMBER({esp_ref}),{esp_ref},0)", kind="link", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, row_adv_pct, 7 + rn, value=f"=IF(ISNUMBER({adv_ref}),{adv_ref},0)", kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=8)
            if rn in (0, 1):
                S.apply_cell(ws, row_adv_usd, 7 + rn, value=0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, row_ctu, 7 + rn, value=0, kind="historico", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            else:
                S.apply_cell(ws, row_adv_usd, 7 + rn, value=f"=IFERROR(({c}{row_p}+{c}{row_car_u})*{c}{row_adv_pct},0)", kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
                transp_ref = iref(inputs_idx, "B. Condiciones", "Costo de transporte por unidad", dim, rn)
                f = f"=IFERROR(IF(ISNUMBER({transp_ref}),{transp_ref},0),0)+IFERROR({c}{row_esp},0)+IFERROR({c}{row_adv_usd},0)"
                S.apply_cell(ws, row_ctu, 7 + rn, value=f, kind="output", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, row_ct, 7 + rn, value=f"={c}{row_u}*{c}{row_ctu}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
    r[0] += 1

    ws.sheet_state = "hidden"
    return ws, row_idx, r, dict(demanda=row_demanda, tam=row_tam, ventas_at=row_ventas_at, excl_at=row_excl_at,
                                  vef=row_vef, insat=row_insat, estado_d=row_estado_d, export_u=row_export_u,
                                  pt=row_pt, costo_tot=row_costo_tot, costo_tot_u=row_costo_tot_u)


def build_engine_produccion_part2(ws_prod, ridx_prod, ridx_mkt):
    """Completa, en _ENGINE_PRODUCCION (Sección C), las celdas 'Ventas (unidades)' de Ronda 2-12 con
    el link a _ENGINE_MERCADO (Sección C, 'Ventas efectivas VALIDADAS'). R0/R1 ya quedaron con el dato
    histórico real en build_engine_produccion_part1() y no se tocan."""
    for area in AREAS:
        for tech in TECNOLOGIAS:
            dim = f"{area} / {tech}"
            row_v = ridx_prod[f"C. Disponibilidad||Ventas (unidades)||{dim}"]
            for rn in range(2, 13):
                ref = cref(ENGINE_MKT_SHEET, ridx_mkt[f"C. Asignación||Ventas efectivas VALIDADAS de esta área/tecnología||{dim}"], rn)
                S.apply_cell(ws_prod, row_v, 7 + rn, value=f"={ref}", kind="link", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
    ws_prod.sheet_state = "hidden"


# ======================================================================================
# _ENGINE_FINANCIERO — P&L por país (transfer pricing + impuestos con carry-forward propio),
# consolidado GLOBAL de Estado de Resultados, Balance y Cash Flow (Motor E5, simplificado:
# balance/CF SOLO a nivel Global, ver README).
# ======================================================================================
def build_engine_financiero(wb, inputs_idx, prod_out, mkt_out, rdos_files, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(ENGINE_FIN_SHEET)
    ncols = 6 + len(RONDAS)
    S.set_col_widths(ws, [16, 48, 14, 10, 14, 34] + [13] * len(RONDAS))
    S.style_title_row(ws, 1, "_ENGINE_FINANCIERO — P&L, BALANCE y FLUJO DE FONDOS POR PAÍS (Fase 2: caja/deuda/interés reales por país, ya no pool Global) + consolidado GLOBAL = suma de los 3 países", ncols)
    header_row = 3
    headers = ["Bloque", "Variable", "Dimensión", "Unidad", "Tipo", "Nota"] + [f"R{r}" for r in RONDAS]
    S.style_header_row(ws, header_row, headers)
    r = [header_row + 1]
    row_idx = {}

    def put(bloque, variable, dimension, unidad="", tipo="CÁLCULO", nota=""):
        rr = r[0]
        row_idx[f"{bloque}||{variable}||{dimension}"] = rr
        S.apply_cell(ws, rr, 1, value=bloque, kind="calculo", align=S.ALIGN_LEFT, size=8, bold=True)
        S.apply_cell(ws, rr, 2, value=variable, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=8)
        S.apply_cell(ws, rr, 3, value=dimension, kind="calculo", align=S.ALIGN_LEFT, size=8)
        S.apply_cell(ws, rr, 4, value=unidad, kind="calculo", align=S.ALIGN_CENTER, size=7)
        S.apply_cell(ws, rr, 5, value=tipo, kind="calculo", align=S.ALIGN_CENTER, size=7, bold=True)
        S.apply_cell(ws, rr, 6, value=nota, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=6, italic=True)
        r[0] += 1
        return rr

    def section(title):
        S.style_section_row(ws, r[0], title, ncols)
        r[0] += 1

    def fx_of(pais, rn):
        if pais == "EE.UU.":
            return "1"
        dim = "China (RMB)" if pais == "China" else "Europa (EUR)"
        return f"IF(ISNUMBER({iref(inputs_idx,'B. Condiciones','Tipo de cambio (USD por 1 unidad de moneda local)',dim,rn)}),{iref(inputs_idx,'B. Condiciones','Tipo de cambio (USD por 1 unidad de moneda local)',dim,rn)},1)"

    def rutas_hacia(pais):
        return [o for (o, d) in PARES_EXPORT if d == pais]

    def rutas_desde(pais):
        return [d for (o, d) in PARES_EXPORT if o == pais]

    def PF(row):
        return f"$G${row}"

    # ============ SECCIÓN A: P&L POR PAÍS (interno, para impuestos con transfer pricing) ============
    section("SECCIÓN A · CUENTA DE RESULTADOS POR PAÍS (interno — transfer pricing e impuestos con tasa/carry-forward propios; el OUTPUT visible de gestión es SOLO GLOBAL, ver 02_ESTADOS_PROYECTADOS)")
    country_rows = {}
    for pais in PAISES:
        produce = pais in AREAS
        cr = {}
        row_mercado = put("A. P&L país", "Ingresos de mercados", pais, "USD")
        row_transf = put("A. P&L país", "Ingresos de transferencias internas", pais, "USD")
        row_ing = put("A. P&L país", "Ingresos totales", pais, "USD", "OUTPUT")
        row_cogs = put("A. P&L país", "Costos de fabricación (interna+contratada)", pais, "USD")
        row_import = put("A. P&L país", "Costos de productos importados", pais, "USD")
        row_transp = put("A. P&L país", "Costos de transporte y aranceles", pais, "USD")
        row_nfab = put("A. P&L país", "N° de fábricas operativas (memo)", pais, "fábricas")
        row_car = put("A. P&L país", "Costos de la característica", pais, "USD")
        row_ginv = put("A. P&L país", "Costos de gestión de inventario", pais, "USD", "CÁLCULO",
                       "Fijo(miles USD, 01_INPUTS)×1000 + Variable(USD/u.)×Σ Inventario final (unidades, _ENGINE_PRODUCCION). "
                       "Da 0 hasta que se carguen los inputs de 01_INPUTS (Condición de Ronda, hoy en blanco).")
        row_energia = put("A. P&L país", "Costos de Energía", pais, "USD", "CÁLCULO",
                       "Σ_tecnología Unidades Producidas×Consumo(kWh/u.)×TarifaPromedio, TarifaPromedio=%Renovable×TarifaRenovable"
                       "+(1-%Renovable)×TarifaFósil (todo 01_INPUTS, Condición de Ronda). Da 0 hasta cargar los inputs.")
        row_agua = put("A. P&L país", "Costos de Agua", pais, "USD", "CÁLCULO",
                       "Σ_tecnología Unidades Producidas×Consumo(m3/u.)×TarifaAgua (01_INPUTS, Condición de Ronda). "
                       "Da 0 hasta cargar los inputs.")
        row_carbono = put("A. P&L país", "Costo del Carbono", pais, "USD", "CÁLCULO",
                       "Σ_tecnología (Unidades Producidas×Emisión(kg/u.)/1000)×TarifaCarbono USD/ton (01_INPUTS, Condición de "
                       "Ronda). Da 0 hasta cargar los inputs.")
        row_rrhh = put("A. P&L país", "I+D — RRHH (prorrateado por N° de fábricas)", pais, "USD")
        row_lic = put("A. P&L país", "I+D — Licencias (D4, prorrateado)", pais, "USD")
        row_id = put("A. P&L país", "I+D total", pais, "USD")
        row_promo = put("A. P&L país", "Promoción", pais, "USD")
        row_admin = put("A. P&L país", "Administración", pais, "USD")
        row_gyc = put("A. P&L país", "Costos y gastos totales", pais, "USD", "OUTPUT")
        row_ebitda = put("A. P&L país", "EBITDA", pais, "USD", "OUTPUT")
        row_dep = put("A. P&L país", "Depreciación", pais, "USD")
        row_ebit = put("A. P&L país", "EBIT", pais, "USD", "OUTPUT")
        row_gfn = put("A. P&L país", "Gastos financieros netos (real, por país)", pais, "USD", "OUTPUT",
                      "Fase 2: Deuda LP apertura(solo EE.UU.)×tasa LP + Deuda CP apertura×tasa CP propia del país − Caja apertura×tasa de interés efectivo propia del país. Sobre saldos de APERTURA (ronda anterior) para evitar circularidad con el ajuste automático de la MISMA ronda -- ya NO es un pro-rateo de un pool Global (reemplaza la simplificación de v1.2).")
        row_ebt = put("A. P&L país", "Beneficio antes de impuestos", pais, "USD")
        row_perd_ap = put("A. P&L país", "Pérdida acumulada (apertura, memo)", pais, "USD")
        row_base_imp = put("A. P&L país", "Base imponible", pais, "USD")
        row_perd_ci = put("A. P&L país", "Pérdida acumulada (cierre, memo)", pais, "USD")
        row_tasa = put("A. P&L país", "Tasa de impuesto corporativo", pais, "%")
        row_imp = put("A. P&L país", "Impuesto sobre el beneficio", pais, "USD")
        row_ben = put("A. P&L país", "Beneficio de la ronda", pais, "USD", "OUTPUT")
        cr.update(dict(mercado=row_mercado, transf=row_transf, ing=row_ing, cogs=row_cogs, imp_cost=row_import,
                       transp=row_transp, nfab=row_nfab, car=row_car, ginv=row_ginv, energia=row_energia, agua=row_agua,
                       carbono=row_carbono, rrhh=row_rrhh, lic=row_lic, id_tot=row_id,
                       promo=row_promo, admin=row_admin, gyc=row_gyc, ebitda=row_ebitda, dep=row_dep, ebit=row_ebit,
                       gfn=row_gfn, ebt=row_ebt, perd_ap=row_perd_ap, base_imp=row_base_imp, perd_ci=row_perd_ci,
                       tasa=row_tasa, imp=row_imp, ben=row_ben))

        # ============ Fase 2, cambio 2: BALANCE POR PAÍS (reemplaza el pool GLOBAL de v1.2) ============
        b_af = put("BAL país", "Activo fijo (neto)", pais, "USD", "OUTPUT", "Link _ENGINE_PRODUCCION." if produce else "Europa no produce -- siempre 0.")
        b_inv = put("BAL país", "Inventario (FIFO)", pais, "USD", "OUTPUT", "Σ Inventario final FIFO por tecnología (_ENGINE_PRODUCCION)." if produce else "Europa no mantiene inventario -- siempre 0.")
        b_cxc = put("BAL país", "Cuentas por Cobrar", pais, "USD", "OUTPUT", "[SUPUESTO CALIBRADO] Ingresos totales × tasa CxC/Ingresos propia del país (calibrada contra RDOS R1, ver PARAM) -- MISMA tasa que usaba MOTOR_VALIDADO_R2/build_engine_e5.py.")
        b_caja = put("BAL país", "Efectivo y equivalentes", pais, "USD", "OUTPUT", "Efectivo de cierre del Flujo de Fondos de este país (más abajo).")
        b_activos = put("BAL país", "Activos Totales", pais, "USD", "OUTPUT")
        b_capsoc = put("BAL país", "Capital social", pais, "USD", "OUTPUT",
                        "Casa matriz: apertura + Emisión(D8) − Recompra(D8)." if pais == "EE.UU." else "Filiales: todo el financiamiento se decide en casa matriz (manual cap.11, REGLA VERIFICADA) -- constante en el valor real de Ronda 1 (RDOS).")
        b_capad = put("BAL país", "Capital adicional desembolsado", pais, "USD", "OUTPUT", "Constante (=0, sin regla CESIM que lo modifique en este modelo).")
        b_ga = put("BAL país", "Ganancias acumuladas", pais, "USD", "OUTPUT",
                   "Apertura + Beneficio de la ronda − Dividendos a accionistas(D8) + Dividendos recibidos de filiales(D8)." if pais == "EE.UU." else "Apertura + Beneficio de la ronda − Dividendos pagados a casa matriz (D8).")
        b_pat = put("BAL país", "Total patrimonio neto", pais, "USD", "OUTPUT")
        b_dlp = put("BAL país", "Deudas a largo plazo", pais, "USD", "OUTPUT",
                    "Apertura + Variación de deuda LP (D8)." if pais == "EE.UU." else "Solo Casa matriz EE.UU. emite deuda LP (manual cap.11, REGLA VERIFICADA) -- siempre 0.")
        b_dcp = put("BAL país", "Deudas a corto plazo (ajuste automático)", pais, "USD", "OUTPUT",
                    "REGLA VERIFICADA (manual cap.11): repone/repaga cada país individualmente hasta su propio efectivo mínimo (Condición, USD 2.000.000/país). Interés sobre saldos de APERTURA (ver Gastos financieros netos, Sección A) -- evita circularidad con el ajuste automático de la MISMA ronda.")
        b_prest = put("BAL país", "Préstamos internos", pais, "USD", "OUTPUT", "Apertura + Σ transferencias recibidas(D8) − Σ transferencias enviadas(D8), sin interés. REGLA VERIFICADA: netea a 0 entre los 3 países cada ronda (ver 04_CONTROL_MODELO).")
        b_cxp = put("BAL país", "Cuentas por pagar", pais, "USD", "OUTPUT", "[SUPUESTO CALIBRADO] Costos y gastos totales × tasa CxP/Costos propia del país (calibrada contra RDOS R1, ver PARAM) -- MISMA tasa que usaba MOTOR_VALIDADO_R2/build_engine_e5.py.")
        b_pas = put("BAL país", "Pasivos Totales", pais, "USD", "OUTPUT")
        b_ctrl = put("BAL país", "Control: Activo Total − (Patrimonio + Pasivo)", pais, "USD", "CONTROL", "Debe ser ≈0 (identidad contable) por país.")
        cr.update(af=b_af, inv=b_inv, cxc=b_cxc, caja=b_caja, activos=b_activos, capsoc=b_capsoc, capad=b_capad,
                  ga=b_ga, pat=b_pat, dlp=b_dlp, dcp=b_dcp, prest=b_prest, cxp=b_cxp, pasivos=b_pas, bal_ctrl=b_ctrl)

        # ============ Fase 2, cambio 2: FLUJO DE FONDOS POR PAÍS (método indirecto) ============
        f_cfo = put("CF país", "Total actividades operativas", pais, "USD", "OUTPUT", "Beneficio + Depreciación − ΔInventario − ΔCxC + ΔCxP (método indirecto).")
        f_cfi = put("CF país", "Total actividades de inversión", pais, "USD", "OUTPUT", "−Pago por inversión + Cobro por desinversión (_ENGINE_PRODUCCION)." if produce else "Europa no produce -- siempre 0.")
        f_cff = put("CF país", "Total actividades financieras (discrecional, D8)", pais, "USD", "OUTPUT",
                   "Emisión−Recompra−Dividendos a accionistas+Dividendos de filiales+ΔDeuda LP+Préstamos internos netos (D8)." if pais == "EE.UU." else "−Dividendos pagados a casa matriz(D8)+Préstamos internos netos(D8).")
        f_cfaj = put("CF país", "Ajuste automático de deuda de corto plazo", pais, "USD", "OUTPUT", "REGLA VERIFICADA (manual cap.11): repone/repaga hasta el efectivo mínimo de ESTE país (Condición).")
        f_cfcambio = put("CF país", "Cambio en efectivo y equivalentes", pais, "USD", "OUTPUT")
        f_cfap = put("CF país", "Efectivo de apertura", pais, "USD", "OUTPUT")
        f_cfci = put("CF país", "Efectivo de cierre", pais, "USD", "OUTPUT")
        cr.update(cfo=f_cfo, cfi=f_cfi, cff=f_cff, cfaj=f_cfaj, cfcambio=f_cfcambio, cfap=f_cfap, cfci=f_cfci)

        country_rows[pais] = cr
    r[0] += 1

    # (rellenado por país; requiere nfab de EE.UU./China para prorratear RRHH y licencias -> loop 2 pasadas)
    for pais in PAISES:
        cr = country_rows[pais]
        produce = pais in AREAS
        for rn in RONDAS:
            c = col(rn)
            fx = fx_of(pais, rn)
            terms_m = []
            for tech in TECNOLOGIAS:
                vef_ref = cref(ENGINE_MKT_SHEET, mkt_out["vef"][(pais, tech)], rn)
                precio_ref = iref(inputs_idx, "D5 · MARKETING", "Precio de venta", f"{pais} / {tech}", rn)
                terms_m.append(f'({vef_ref}*(IF(ISNUMBER({precio_ref}),{precio_ref},0)*({fx})))')
            S.apply_cell(ws, cr["mercado"], 7 + rn, value="=" + "+".join(terms_m), kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if produce:
                terms_t = [f'({cref(ENGINE_MKT_SHEET, mkt_out["export_u"][(pais, d)], rn)}*{cref(ENGINE_MKT_SHEET, mkt_out["pt"][(pais, d)], rn)})' for d in rutas_desde(pais) if rn not in (0, 1)]
                f_t = ("=" + "+".join(terms_t)) if (terms_t and rn not in (0, 1)) else "0"
            else:
                f_t = "0"
            S.apply_cell(ws, cr["transf"], 7 + rn, value=f_t, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["ing"], 7 + rn, value=f"={c}{cr['mercado']}+{c}{cr['transf']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

            if produce:
                f_cogs = "=" + "+".join(cref(ENGINE_PROD_SHEET, prod_out["cogs"][(pais, t)], rn) for t in TECNOLOGIAS)
            else:
                f_cogs = "0"
            S.apply_cell(ws, cr["cogs"], 7 + rn, value=f_cogs, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            origenes = rutas_hacia(pais)
            if origenes and rn not in (0, 1):
                f_imp = "=" + "+".join(f'({cref(ENGINE_MKT_SHEET, mkt_out["export_u"][(o, pais)], rn)}*{cref(ENGINE_MKT_SHEET, mkt_out["pt"][(o, pais)], rn)})' for o in origenes)
                f_transp = "=" + "+".join(cref(ENGINE_MKT_SHEET, mkt_out["costo_tot"][(o, pais)], rn) for o in origenes)
            else:
                f_imp, f_transp = "0", "0"
            S.apply_cell(ws, cr["imp_cost"], 7 + rn, value=f_imp, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["transp"], 7 + rn, value=f_transp, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if produce:
                cap_ref = cref(ENGINE_PROD_SHEET, prod_out["cap"][pais], rn)
                capfab_ref = f"'{ENGINE_PROD_SHEET}'!{PF(prod_out['params']['capfab'][pais])}"
                f_nfab = f"=IFERROR({cap_ref}/{capfab_ref},0)"
            else:
                f_nfab = "0"
            S.apply_cell(ws, cr["nfab"], 7 + rn, value=f_nfab, kind="calculo", numfmt='#,##0.00', align=S.ALIGN_RIGHT, size=8)

            costo_car_ref = iref(inputs_idx, "B. Condiciones", "Costo de la característica", pais, rn)
            terms_car = []
            for tech in TECNOLOGIAS:
                vef_ref = cref(ENGINE_MKT_SHEET, mkt_out["vef"][(pais, tech)], rn)
                car_ref = iref(inputs_idx, "D5 · MARKETING", "Características ofrecidas (de las disponibles, tope 10)", f"{pais} / {tech}", rn)
                terms_car.append(f'({vef_ref}*IF(ISNUMBER({car_ref}),{car_ref},0)*IF(ISNUMBER({costo_car_ref}),{costo_car_ref},0))')
            S.apply_cell(ws, cr["car"], 7 + rn, value="=" + "+".join(terms_car), kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if produce:
                ginv_fijo_ref = iref(inputs_idx, "B. Condiciones", "Costos fijos de gestión de inventario", pais, rn)
                ginv_var_ref = iref(inputs_idx, "B. Condiciones", "Costos variables de gestión de inventario", pais, rn)
                inv_fin_terms = "+".join(cref(ENGINE_PROD_SHEET, prod_out["inv_fin_u"][(pais, t)], rn) for t in TECNOLOGIAS)
                # "miles USD" en 01_INPUTS -> *1000 (misma convención que el resto de filas "miles X" del modelo).
                f_ginv = (f"=IF(ISNUMBER({ginv_fijo_ref}),{ginv_fijo_ref}*1000,0)"
                          f"+IF(ISNUMBER({ginv_var_ref}),{ginv_var_ref},0)*({inv_fin_terms})")
            else:
                f_ginv = "0"
            S.apply_cell(ws, cr["ginv"], 7 + rn, value=f_ginv, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            if produce:
                pct_ren_ref = iref(inputs_idx, "B. Condiciones", "% de utilización de energía renovable (vs. fósil)", pais, rn)
                tarifa_ren_ref = iref(inputs_idx, "B. Condiciones", "Costo de la energía renovable", pais, rn)
                tarifa_fos_ref = iref(inputs_idx, "B. Condiciones", "Costo de la energía fósil", pais, rn)
                tarifa_agua_ref = iref(inputs_idx, "B. Condiciones", "Costo del agua", pais, rn)
                tarifa_co2_ref = iref(inputs_idx, "B. Condiciones", "Costo del Carbono", pais, rn)
                pct_ren = f"IF(ISNUMBER({pct_ren_ref}),{pct_ren_ref},0)"
                tarifa_prom = (f"(({pct_ren})*IF(ISNUMBER({tarifa_ren_ref}),{tarifa_ren_ref},0)"
                               f"+(1-({pct_ren}))*IF(ISNUMBER({tarifa_fos_ref}),{tarifa_fos_ref},0))")
                terms_kwh, terms_m3, terms_co2 = [], [], []
                for tech in TECNOLOGIAS:
                    prod_ref = cref(ENGINE_PROD_SHEET, prod_out["prod_u"][(pais, tech)], rn)
                    kwh_ref = iref(inputs_idx, "B. Condiciones", "Consumo de energía (kWh por unidad producida)", f"{pais} / {tech}", rn)
                    m3_ref = iref(inputs_idx, "B. Condiciones", "Consumo de agua (m3 por unidad producida)", f"{pais} / {tech}", rn)
                    co2_ref = iref(inputs_idx, "B. Condiciones", "Emisión de CO2 (kg por unidad producida)", f"{pais} / {tech}", rn)
                    terms_kwh.append(f"({prod_ref}*IF(ISNUMBER({kwh_ref}),{kwh_ref},0))")
                    terms_m3.append(f"({prod_ref}*IF(ISNUMBER({m3_ref}),{m3_ref},0))")
                    terms_co2.append(f"({prod_ref}*IF(ISNUMBER({co2_ref}),{co2_ref},0))")
                f_energia = f"=({'+'.join(terms_kwh)})*({tarifa_prom})"
                f_agua = f"=({'+'.join(terms_m3)})*IF(ISNUMBER({tarifa_agua_ref}),{tarifa_agua_ref},0)"
                f_carbono = f"=(({'+'.join(terms_co2)})/1000)*IF(ISNUMBER({tarifa_co2_ref}),{tarifa_co2_ref},0)"
            else:
                f_energia, f_agua, f_carbono = "0", "0", "0"
            S.apply_cell(ws, cr["energia"], 7 + rn, value=f_energia, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["agua"], 7 + rn, value=f_agua, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["carbono"], 7 + rn, value=f_carbono, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            S.apply_cell(ws, cr["rrhh"], 7 + rn, value=0, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["lic"], 7 + rn, value=0, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["id_tot"], 7 + rn, value=f"={c}{cr['rrhh']}+{c}{cr['lic']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            f_promo = "=" + "+".join(f'IF(ISNUMBER({iref(inputs_idx,"D5 · MARKETING","Presupuesto de promoción",f"{pais} / {t}",rn)}),{iref(inputs_idx,"D5 · MARKETING","Presupuesto de promoción",f"{pais} / {t}",rn)},0)' for t in TECNOLOGIAS)
            S.apply_cell(ws, cr["promo"], 7 + rn, value=f_promo, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            fijo_ref = iref(inputs_idx, "B. Condiciones", "Administración — componente fijo", pais, rn)
            pct_ref = iref(inputs_idx, "B. Condiciones", "Administración — componente % sobre ingresos", pais, rn)
            fijo = f"IF(ISNUMBER({fijo_ref}),{fijo_ref},0)"
            pctv = f"IF(ISNUMBER({pct_ref}),{pct_ref},0)"
            if produce:
                planta_ref = iref(inputs_idx, "B. Condiciones", "Administración — componente por planta (aplicado linealmente, ver ALERTA)", pais, rn)
                planta = f"IF(ISNUMBER({planta_ref}),{planta_ref},0)"
                f_admin = f"=({fijo})+({planta})*{c}{cr['nfab']}+({pctv})*{c}{cr['ing']}"
            else:
                f_admin = f"=({fijo})+({pctv})*{c}{cr['ing']}"
            S.apply_cell(ws, cr["admin"], 7 + rn, value=f_admin, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            f_gyc = f"={c}{cr['cogs']}+{c}{cr['imp_cost']}+{c}{cr['transp']}+{c}{cr['car']}+{c}{cr['ginv']}+{c}{cr['energia']}+{c}{cr['agua']}+{c}{cr['carbono']}+{c}{cr['id_tot']}+{c}{cr['promo']}+{c}{cr['admin']}"
            S.apply_cell(ws, cr["gyc"], 7 + rn, value=f_gyc, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            S.apply_cell(ws, cr["ebitda"], 7 + rn, value=f"={c}{cr['ing']}-{c}{cr['gyc']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            f_dep = cref(ENGINE_PROD_SHEET, prod_out["af_dep_ronda"][pais], rn) if produce else "0"
            S.apply_cell(ws, cr["dep"], 7 + rn, value=f"={f_dep}" if produce else "0", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["ebit"], 7 + rn, value=f"={c}{cr['ebitda']}-{c}{cr['dep']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        r[0] += 1
    r[0] += 1

    # ============ PATCH: reparto de RRHH (global, D3) y licencias (D4) por N° de fábricas ============
    # REGLA CONFIRMADA (manual cap. 12, verbatim, ya usada en el motor anterior): "cuando también
    # tenga producción en Asia, la I+D se dividirá proporcionalmente entre EE.UU. y Asia según el
    # número de fábricas en cada una de las áreas". Se completa DESPUÉS del loop principal porque
    # necesita las filas 'nfab' de AMBOS países productores.
    nfab_eeuu = country_rows["EE.UU."]["nfab"]
    nfab_china = country_rows["China"]["nfab"]
    for pais in AREAS:
        cr = country_rows[pais]
        row_nfab_this = cr["nfab"]
        for rn in RONDAS:
            c = col(rn)
            lic_terms = "+".join(f'IF(ISNUMBER({iref(inputs_idx,"D4 · I+D / TECNOLOGÍA","Compra de licencia (USD)",tech,rn)}),{iref(inputs_idx,"D4 · I+D / TECNOLOGÍA","Compra de licencia (USD)",tech,rn)},0)' for tech in TECNOLOGIAS)
            f_lic = f"=IFERROR(({lic_terms})*{c}{row_nfab_this}/({c}{nfab_eeuu}+{c}{nfab_china}),0)"
            S.apply_cell(ws, cr["lic"], 7 + rn, value=f_lic, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            if rn not in (0, 1):
                pct_bruto = iref(inputs_idx, "B. Condiciones", "% Sueldo bruto sobre costo laboral total (RRHH)", "Global", rn)
                otros_id = iref(inputs_idx, "B. Condiciones", "Otros costos de I+D por empleado/mes", "Global", rn)
                costo_contrat = iref(inputs_idx, "B. Condiciones", "Costo de contratación por persona", "Global", rn)
                t1 = iref(inputs_idx, "B. Condiciones", "Costo de despido por persona", "tramo 5%-20% de reducción", rn)
                t2 = iref(inputs_idx, "B. Condiciones", "Costo de despido por persona", "tramo 20%-50% de reducción", rn)
                t3 = iref(inputs_idx, "B. Condiciones", "Costo de despido por persona", "tramo >50% de reducción", rn)
                sal_ref = iref(inputs_idx, "D3 · RRHH", "Salario mensual", "Global", rn)
                cap_ref = iref(inputs_idx, "D3 · RRHH", "Presupuesto mensual de capacitación", "Global", rn)
                hc_ref = iref(inputs_idx, "D3 · RRHH", "Cantidad de personal de I+D", "Global", rn)
                hc_prev_raw = iref(inputs_idx, "D3 · RRHH", "Cantidad de personal de I+D", "Global", rn - 1)
                hc_prev_ref = f'IF(ISNUMBER({hc_prev_raw}),{hc_prev_raw},{hc_ref})'
                salarios = f'(({sal_ref}/{pct_bruto})*{hc_ref}*12)'
                capacitacion = f'({cap_ref}*{hc_ref}*12)'
                otros = f'({otros_id}*{hc_ref}*12)'
                contrat = f'(MAX(0,{hc_ref}-{hc_prev_ref})*{costo_contrat})'
                reduccion = f'MAX(0,{hc_prev_ref}-{hc_ref})'
                pct_red = f'IFERROR({reduccion}/{hc_prev_ref},0)'
                despido = f'({reduccion}*IF({pct_red}>0.5,{t3},IF({pct_red}>0.2,{t2},IF({pct_red}>=0.05,{t1},0))))'
                f_rrhh_global = "(" + "+".join([salarios, capacitacion, otros, contrat, despido]) + ")"
                f_rrhh = f"=IFERROR(({f_rrhh_global})*{c}{row_nfab_this}/({c}{nfab_eeuu}+{c}{nfab_china}),0)"
                S.apply_cell(ws, cr["rrhh"], 7 + rn, value=f_rrhh, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            else:
                S.apply_cell(ws, cr["rrhh"], 7 + rn, value=0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["id_tot"], 7 + rn, value=f"={c}{cr['rrhh']}+{c}{cr['lic']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

    # ============ Fase 2, cambio 2: PARÁMETROS CxC/CxP CALIBRADOS POR PAÍS (reemplaza el pool blended GLOBAL de v1.2) ============
    section("PARÁMETROS · CxC/CxP calibrados POR PAÍS (Fase 2 — MISMAS 3 tasas por país que usaba MOTOR_VALIDADO_R2/build_engine_e5.py, calibradas contra RDOS Ronda 1 real; ver 01_INPUTS)")
    CXC_PCT = {"EE.UU.": 0.0271, "China": 0.0358, "Europa": 0.0385}
    CXP_PCT = {"EE.UU.": 0.0359, "China": 0.0363, "Europa": 0.0075}
    row_cxc_pct = {}
    row_cxp_pct = {}
    for pais in PAISES:
        row = put("PARAM", "CxC / Ingresos totales de la ronda (calibrado RDOS R1)", pais, "%", "SUPUESTO CALIBRADO",
                   "[SUPUESTO CALIBRADO] = Cuentas por Cobrar / Ingresos, RDOS Ronda 1 real de este país. El manual no publica una regla de días de cobro; MISMO valor que usaba MOTOR_VALIDADO_R2 (build_inputs.py, sección CxC/CxP), no una regla CESIM.")
        row_cxc_pct[pais] = row
        S.apply_cell(ws, row, 7, value=CXC_PCT[pais], kind="input", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9)
    for pais in PAISES:
        row = put("PARAM", "CxP / Costos y gastos totales de la ronda (calibrado RDOS R1)", pais, "%", "SUPUESTO CALIBRADO",
                   "[SUPUESTO CALIBRADO] = Cuentas por Pagar / Costos y gastos totales, RDOS Ronda 1 real de este país. MISMO valor que usaba MOTOR_VALIDADO_R2, no una regla CESIM.")
        row_cxp_pct[pais] = row
        S.apply_cell(ws, row, 7, value=CXP_PCT[pais], kind="input", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9)
    r[0] += 1

    # ============ SECCIÓN B: GLOBAL — Estado de Resultados consolidado ============
    section("SECCIÓN B · CUENTA DE RESULTADOS GLOBAL (consolidado — eliminaciones verificadas: transferencias internas y costos de productos importados se cancelan al sumar)")
    g_ing = put("B. P&L GLOBAL", "Ingresos por ventas", "Global", "USD", "OUTPUT", "REGLA VERIFICADA: Σ 'Ingresos de mercados' de los 3 países (transferencias internas NO se suman, se eliminan).")
    g_costos = put("B. P&L GLOBAL", "Costos y gastos totales", "Global", "USD", "OUTPUT", "REGLA VERIFICADA: Σ (Fabricación+Característica+Gestión de inventario+Energía+Agua+Carbono+Transporte/aranceles+I+D+Promoción+Administración). EXCLUYE 'Costos de productos importados' (se elimina, es la contrapartida de la transferencia interna).")
    g_ebitda = put("B. P&L GLOBAL", "EBITDA", "Global", "USD", "OUTPUT")
    g_dep = put("B. P&L GLOBAL", "Depreciación", "Global", "USD", "OUTPUT")
    g_ebit = put("B. P&L GLOBAL", "EBIT", "Global", "USD", "OUTPUT")
    g_gfn = put("B. P&L GLOBAL", "Gastos financieros netos", "Global", "USD", "OUTPUT", "Fase 2: Σ Gastos financieros netos REALES de los 3 países (cada uno sobre sus propios saldos de apertura de Caja/Deuda LP/Deuda CP -- ver Sección A).")
    g_ebt = put("B. P&L GLOBAL", "Beneficio antes de impuestos", "Global", "USD", "OUTPUT")
    g_imp = put("B. P&L GLOBAL", "Impuesto sobre el beneficio", "Global", "USD", "OUTPUT", "Σ Impuesto de los 3 países (cada uno tributa localmente con su propia tasa y carry-forward, ver Sección A).")
    g_ben = put("B. P&L GLOBAL", "Beneficio de la ronda", "Global", "USD", "OUTPUT")
    g_ctrl = put("B. P&L GLOBAL", "Control: Beneficio Global − Σ Beneficio países", "Global", "USD", "CONTROL", "Debe ser ≈0 (identidad de consolidación).")
    r[0] += 1

    # ============ SECCIÓN C: GLOBAL — Balance resumido (pool único de caja/deuda) ============
    section("SECCIÓN C · BALANCE GLOBAL (Fase 2: Σ Balance por país -- EE.UU., China, Europa -- ver 'BAL país' en Sección A arriba; ya NO es un pool único simplificado)")
    g_af = put("C. BAL GLOBAL", "Activo fijo (neto)", "Global", "USD", "OUTPUT", "Σ Activo fijo neto por país (EE.UU.+China; Europa=0).")
    g_inv = put("C. BAL GLOBAL", "Inventario", "Global", "USD", "OUTPUT", "Σ Inventario FIFO por país (EE.UU.+China; Europa=0).")
    g_cxc = put("C. BAL GLOBAL", "Cuentas por Cobrar", "Global", "USD", "OUTPUT", "Fase 2: Σ CxC por país, cada uno con su propia tasa calibrada (PARAM) -- ya NO una tasa blended única.")
    g_caja = put("C. BAL GLOBAL", "Efectivo y equivalentes", "Global", "USD", "OUTPUT", "Σ Efectivo de cierre por país (Sección D de cada país).")
    g_activos = put("C. BAL GLOBAL", "Activos Totales", "Global", "USD", "OUTPUT")
    g_capsoc = put("C. BAL GLOBAL", "Capital social", "Global", "USD", "OUTPUT", "Apertura + Emisión − Recompra (D8, Casa matriz EE.UU. — único nivel del modelo, ver README).")
    g_nacc = put("C. BAL GLOBAL", "Acciones en circulación al cierre (memo)", "Global", "acciones", "HISTÓRICO/CÁLCULO",
                 "DATO HISTÓRICO REAL (RDOS): R0=R1=1.625.000.000 acciones (CADIZ no emitió ni recompró en R0/R1). Desde R2: Apertura + (Emisión D8 − Recompra D8, en USD) ÷ Valor de referencia de la acción (Condición, PROXY). Usado para EPS/P-E/Dividend yield.")
    g_capad = put("C. BAL GLOBAL", "Capital adicional desembolsado", "Global", "USD", "OUTPUT", "Constante (=0, sin regla CESIM que lo modifique en este modelo).")
    g_gan = put("C. BAL GLOBAL", "Ganancias acumuladas", "Global", "USD", "OUTPUT", "Apertura + Beneficio Global de la ronda − Dividendos pagados a accionistas (D8, monto).")
    g_pat = put("C. BAL GLOBAL", "Patrimonio neto", "Global", "USD", "OUTPUT")
    g_dlp = put("C. BAL GLOBAL", "Deudas a largo plazo", "Global", "USD", "OUTPUT", "Exclusivamente a nivel Casa matriz EE.UU. (manual cap.11, REGLA VERIFICADA) -- Σ por país (China/Europa=0).")
    g_dcp = put("C. BAL GLOBAL", "Deudas a corto plazo (ajuste automático)", "Global", "USD", "OUTPUT", "Fase 2: Σ Deuda CP por país -- REGLA VERIFICADA (manual cap.11), cada país se ajusta a SU PROPIO efectivo mínimo (Condición, USD 2.000.000/país).")
    g_cxp = put("C. BAL GLOBAL", "Cuentas por Pagar", "Global", "USD", "OUTPUT", "Fase 2: Σ CxP por país, cada uno con su propia tasa calibrada (PARAM).")
    g_pasivos = put("C. BAL GLOBAL", "Pasivos Totales", "Global", "USD", "OUTPUT")
    g_bal_ctrl = put("C. BAL GLOBAL", "Control: Activo Total − (Patrimonio + Pasivo)", "Global", "USD", "CONTROL", "Debe ser ≈0 (identidad contable).")
    g_prest_ctrl = put("C. BAL GLOBAL", "Control: Σ Préstamos internos (3 países)", "Global", "USD", "CONTROL", "REGLA VERIFICADA: debe ser exactamente 0 -- los préstamos internos entre países netean a 0 por construcción (Fase 2, cambio 2). Ver 04_CONTROL_MODELO.")
    r[0] += 1

    # ============ SECCIÓN D: GLOBAL — Flujo de fondos (método indirecto) ============
    section("SECCIÓN D · FLUJO DE FONDOS GLOBAL — Fase 2: Σ Flujo de Fondos por país (cada país con su propio ajuste automático de deuda CP); préstamos internos y dividendos intercompañía netean a 0 por construcción (ver 04_CONTROL_MODELO)")
    g_cfo = put("D. CF GLOBAL", "Total actividades operativas", "Global", "USD", "OUTPUT", "Σ Total operativo por país (método indirecto, cada uno con su propio Beneficio/Depreciación/ΔInventario/ΔCxC/ΔCxP).")
    g_cfi = put("D. CF GLOBAL", "Total actividades de inversión", "Global", "USD", "OUTPUT", "Σ Total de inversión por país (pagos de inversión en fábrica − cobros por desinversión, _ENGINE_PRODUCCION).")
    g_cff = put("D. CF GLOBAL", "Total actividades financieras (discrecional, D8)", "Global", "USD", "OUTPUT", "Σ Total financiero por país. Las transferencias entre países y dividendos a casa matriz (filial) netean a 0 GLOBAL por construcción (ver control de préstamos internos).")
    g_cfaj = put("D. CF GLOBAL", "Ajuste automático de deuda de corto plazo", "Global", "USD", "OUTPUT", "Fase 2: Σ ajuste automático por país -- REGLA VERIFICADA (manual cap.11), cada país repone/repaga hasta SU PROPIO efectivo mínimo.")
    g_cfcambio = put("D. CF GLOBAL", "Cambio en efectivo y equivalentes", "Global", "USD", "OUTPUT")
    g_cfap = put("D. CF GLOBAL", "Efectivo de apertura", "Global", "USD", "OUTPUT")
    g_cfci = put("D. CF GLOBAL", "Efectivo de cierre", "Global", "USD", "OUTPUT")
    g_cf_ctrl = put("D. CF GLOBAL", "Control: Efectivo cierre CF − Efectivo Balance", "Global", "USD", "CONTROL", "Debe ser ≈0.")
    r[0] += 1

    glob = dict(ing=g_ing, costos=g_costos, ebitda=g_ebitda, dep=g_dep, ebit=g_ebit, gfn=g_gfn, ebt=g_ebt, imp=g_imp,
                ben=g_ben, ctrl=g_ctrl, af=g_af, inv=g_inv, cxc=g_cxc, caja=g_caja, activos=g_activos, capsoc=g_capsoc,
                nacc=g_nacc, capad=g_capad, gan=g_gan, pat=g_pat, dlp=g_dlp, dcp=g_dcp, cxp=g_cxp, pasivos=g_pasivos,
                bal_ctrl=g_bal_ctrl, prest_ctrl=g_prest_ctrl, cfo=g_cfo, cfi=g_cfi, cff=g_cff, cfaj=g_cfaj,
                cfcambio=g_cfcambio, cfap=g_cfap, cfci=g_cfci, cf_ctrl=g_cf_ctrl)

    # ============ Fase 2, cambio 2: GFN real por país (saldos de APERTURA) → EBT → carry-forward →
    # Impuesto → Beneficio → BALANCE por país → FLUJO DE FONDOS por país (secuencia no circular,
    # replica MOTOR_VALIDADO_R2/build_engine_e5.py) ============
    hist_pais = extract_hist_pais(rdos_files)

    def hv(rn_, pais_, sec_, label_):
        # Nota (corrección de escala): extract_hist_pais() ya devuelve el valor escalado
        # (rdos_parser aplica el x1000 de "miles USD"/"miles unidades" internamente) -- este
        # helper YA NO vuelve a multiplicar por 1000 (antes lo hacía sobre un valor crudo de
        # ronda0.json/ronda1.json que no traía el x1000 aplicado; con ambos x1000 el resultado
        # habría quedado 1.000.000x en vez de 1.000x).
        return hist_pais.get(rn_, {}).get(pais_, {}).get(sec_, {}).get(label_)

    def prest_terms(pais_, rn_):
        recibidos = "+".join(f'IF(ISNUMBER({iref(inputs_idx,"D8 · FINANZAS","Transferencia de fondos entre países (préstamo interno)",f"{o} → {pais_}",rn_)}),{iref(inputs_idx,"D8 · FINANZAS","Transferencia de fondos entre países (préstamo interno)",f"{o} → {pais_}",rn_)},0)' for o in PAISES if o != pais_)
        enviados = "+".join(f'IF(ISNUMBER({iref(inputs_idx,"D8 · FINANZAS","Transferencia de fondos entre países (préstamo interno)",f"{pais_} → {d}",rn_)}),{iref(inputs_idx,"D8 · FINANZAS","Transferencia de fondos entre países (préstamo interno)",f"{pais_} → {d}",rn_)},0)' for d in PAISES if d != pais_)
        return recibidos, enviados

    for pais in PAISES:
        cr = country_rows[pais]
        produce = pais in AREAS
        for rn in RONDAS:
            c = col(rn)
            if rn in rondas_reales:
                # -------- P&L (cola): DATO HISTÓRICO REAL (RDOS), no fórmula --------
                # Ronda N.N (Fase 4, frontera dinámica): antes fijo a (0,1) -- ahora CUALQUIER ronda
                # con RDOS oficial cargado (rondas_reales) usa el dato real vía hv()/extract_hist_pais,
                # bypasseando por completo las fórmulas de la Sección A de más arriba (que para una
                # ronda real quedan como celdas informativas sin efecto en GFN/EBT/Impuesto/Balance/CF,
                # igual que ya pasaba con R0/R1 antes de esta generalización).
                v_gfn = hv(rn, pais, "pl", "Gastos financieros netos")
                v_ebt = hv(rn, pais, "pl", "Beneficio antes de impuestos")
                v_imp = hv(rn, pais, "pl", "Impuesto sobre el beneficio")
                v_ben = hv(rn, pais, "pl", "Beneficio de la ronda")
                S.apply_cell(ws, cr["gfn"], 7 + rn, value=(v_gfn if v_gfn is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["ebt"], 7 + rn, value=(v_ebt if v_ebt is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                # SUPUESTO PROPIO (no verificado contra manual/RDOS -- el RDOS no publica "pérdida
                # trasladable acumulada" por país): se asume perd_ci=0.0 para TODA ronda real (antes
                # solo R0/R1). perd_ap se lee de la celda anterior (0.0 si rn==0, sin ronda previa) --
                # esto reemplaza el diccionario fijo hist_perd_ap[pais][0/1] para que la cadena de
                # arrastre funcione igual sin importar cuántas rondas sean ya reales.
                v_perd_ap = 0.0 if rn == 0 else f"={col(rn - 1)}{cr['perd_ci']}"
                S.apply_cell(ws, cr["perd_ap"], 7 + rn, value=v_perd_ap, kind=("historico" if rn == 0 else "calculo"), numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["base_imp"], 7 + rn, value="n/a", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
                S.apply_cell(ws, cr["perd_ci"], 7 + rn, value=0.0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["tasa"], 7 + rn, value=f"={iref(inputs_idx,'B. Condiciones','Tasa de impuesto corporativo',pais,rn)}", kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["imp"], 7 + rn, value=(v_imp if v_imp is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["ben"], 7 + rn, value=(v_ben if v_ben is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

                # -------- BALANCE: DATO HISTÓRICO REAL POR PAÍS (RDOS), auto-consistente --------
                v_af = hv(rn, pais, "bal", "Activo fijo") or 0.0
                v_inv = hv(rn, pais, "bal", "Inventario") or 0.0
                v_cxc = hv(rn, pais, "bal", "Cuentas por Cobrar") or 0.0
                v_caja = hv(rn, pais, "bal", "Efectivo y equivalentes de efectivo") or 0.0
                S.apply_cell(ws, cr["af"], 7 + rn, value=v_af, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["inv"], 7 + rn, value=v_inv, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["cxc"], 7 + rn, value=v_cxc, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["caja"], 7 + rn, value=v_caja, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["activos"], 7 + rn, value=f"={c}{cr['af']}+{c}{cr['inv']}+{c}{cr['cxc']}+{c}{cr['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                v_capsoc = hv(rn, pais, "bal", "Capital social") or 0.0
                v_capad = hv(rn, pais, "bal", "Capital adicional desembolsado") or 0.0
                v_ga = (hv(rn, pais, "bal", "Ganancias acumuladas") or 0.0) + (v_ben or 0.0)
                S.apply_cell(ws, cr["capsoc"], 7 + rn, value=v_capsoc, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["capad"], 7 + rn, value=v_capad, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["ga"], 7 + rn, value=v_ga, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["pat"], 7 + rn, value=f"={c}{cr['capsoc']}+{c}{cr['capad']}+{c}{cr['ga']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                v_dlp = hv(rn, pais, "bal", "Deudas a largo plazo") or 0.0
                v_dcp = hv(rn, pais, "bal", "Deudas a corto plazo (no planificadas)") or 0.0
                v_prest = hv(rn, pais, "bal", "Préstamos internos") or 0.0
                v_cxp = hv(rn, pais, "bal", "Cuentas por pagar") or 0.0
                S.apply_cell(ws, cr["dlp"], 7 + rn, value=v_dlp, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["dcp"], 7 + rn, value=v_dcp, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["prest"], 7 + rn, value=v_prest, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["cxp"], 7 + rn, value=v_cxp, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["pasivos"], 7 + rn, value=f"={c}{cr['dlp']}+{c}{cr['dcp']}+{c}{cr['prest']}+{c}{cr['cxp']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                S.apply_cell(ws, cr["bal_ctrl"], 7 + rn, value=f"={c}{cr['activos']}-({c}{cr['pat']}+{c}{cr['pasivos']})", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

                # -------- FLUJO DE FONDOS: DATO HISTÓRICO REAL POR PAÍS (RDOS) --------
                if rn == 0:
                    for kk in ("cfo", "cfi", "cff", "cfaj"):
                        S.apply_cell(ws, cr[kk], 7 + rn, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
                    S.apply_cell(ws, cr["cfcambio"], 7 + rn, value=0.0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    v_cfap0 = hv(rn, pais, "cf", "Efectivo 1.1.")
                    S.apply_cell(ws, cr["cfap"], 7 + rn, value=(v_cfap0 if v_cfap0 is not None else "n/a"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, cr["cfci"], 7 + rn, value=f"={c}{cr['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                else:
                    v_cfo = hv(rn, pais, "cf", "Total (Efectivo proveniente de actividades operativas)")
                    v_cfi = hv(rn, pais, "cf", "Inversiones en fábricas (-) / desinversiones (+)") if produce else 0.0
                    v_cff = hv(rn, pais, "cf", "Total (Efectivo proveniente de actividades financieras)")
                    v_cfcambio = hv(rn, pais, "cf", "Cambios en efectivo y equivalentes de efectivo")
                    v_cfap = hv(rn, pais, "cf", "Efectivo 1.1.")
                    S.apply_cell(ws, cr["cfo"], 7 + rn, value=(v_cfo if v_cfo is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, cr["cfi"], 7 + rn, value=(v_cfi if v_cfi is not None else 0.0), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, cr["cff"], 7 + rn, value=(v_cff if v_cff is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, cr["cfaj"], 7 + rn, value="n/a (histórico, ver RDOS -- Deuda CP no se movió en R1)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
                    S.apply_cell(ws, cr["cfcambio"], 7 + rn, value=(v_cfcambio if v_cfcambio is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                    S.apply_cell(ws, cr["cfap"], 7 + rn, value=(v_cfap if v_cfap is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                    S.apply_cell(ws, cr["cfci"], 7 + rn, value=f"={c}{cr['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
                continue

            # ==================== rn >= 2: FÓRMULAS (Escenario A R2 en adelante) ====================
            cprev = col(rn - 1)
            tasa_lp = iref(inputs_idx, "B. Condiciones", "Tasa de interés — deuda a largo plazo", "Casa matriz EE.UU.", rn)
            tasa_cp = iref(inputs_idx, "B. Condiciones", "Tasa de interés — deuda a corto plazo", pais, rn)
            tasa_ef = iref(inputs_idx, "B. Condiciones", "Tasa de interés sobre efectivo", pais, rn)
            dlp_ap = f"{cprev}{cr['dlp']}" if pais == "EE.UU." else "0"
            # Fase 2, cambio 2: interés REAL de este país sobre saldos de APERTURA (ronda anterior) de
            # Deuda LP (solo EE.UU.), Deuda CP y Caja propias -- ya NO es un pro-rateo de un pool GLOBAL.
            f_gfn = f"=({dlp_ap}*{tasa_lp})+({cprev}{cr['dcp']}*{tasa_cp})-({cprev}{cr['caja']}*{tasa_ef})"
            S.apply_cell(ws, cr["gfn"], 7 + rn, value=f_gfn, kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["ebt"], 7 + rn, value=f"={c}{cr['ebit']}-{c}{cr['gfn']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            # Ronda N.N: ya no hace falta un caso especial para "la primera ronda proyectada" (antes
            # hardcodeado a rn==2, sembrando desde hist_perd_ap[pais][1]) -- la celda perd_ci de la
            # ronda anterior YA está poblada (0.0) para cualquier ronda histórica (ver arriba), así que
            # la misma fórmula de arrastre sirve sea la ronda anterior histórica o ya proyectada.
            S.apply_cell(ws, cr["perd_ap"], 7 + rn, value=f"={cprev}{cr['perd_ci']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["base_imp"], 7 + rn, value=f"=MAX(0,{c}{cr['ebt']}-{c}{cr['perd_ap']})", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["perd_ci"], 7 + rn, value=f"=MAX(0,{c}{cr['perd_ap']}-{c}{cr['ebt']})", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["tasa"], 7 + rn, value=f"={iref(inputs_idx,'B. Condiciones','Tasa de impuesto corporativo',pais,rn)}", kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["imp"], 7 + rn, value=f"={c}{cr['base_imp']}*{c}{cr['tasa']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["ben"], 7 + rn, value=f"={c}{cr['ebt']}-{c}{cr['imp']}", kind="output", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

            # -------- BALANCE por país --------
            f_af = cref(ENGINE_PROD_SHEET, prod_out["af_resid"][pais], rn) if produce else "0"
            S.apply_cell(ws, cr["af"], 7 + rn, value=f"={f_af}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            f_inv = ("=" + "+".join(cref(ENGINE_PROD_SHEET, prod_out["inv_fin_v"][(pais, t)], rn) for t in TECNOLOGIAS)) if produce else "0"
            S.apply_cell(ws, cr["inv"], 7 + rn, value=f_inv, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["cxc"], 7 + rn, value=f"={c}{cr['ing']}*$G${row_cxc_pct[pais]}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["caja"], 7 + rn, value=f"={c}{cr['cfci']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["activos"], 7 + rn, value=f"={c}{cr['af']}+{c}{cr['inv']}+{c}{cr['cxc']}+{c}{cr['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

            recibidos, enviados = prest_terms(pais, rn)
            if pais == "EE.UU.":
                emis = iref(inputs_idx, "D8 · FINANZAS", "Emisión de acciones", "Casa matriz EE.UU.", rn)
                recom = iref(inputs_idx, "D8 · FINANZAS", "Recompra de acciones", "Casa matriz EE.UU.", rn)
                S.apply_cell(ws, cr["capsoc"], 7 + rn, value=f"={cprev}{cr['capsoc']}+IF(ISNUMBER({emis}),{emis},0)-IF(ISNUMBER({recom}),{recom},0)", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                divacc = iref(inputs_idx, "D8 · FINANZAS", "Dividendos pagados a accionistas (monto)", "Casa matriz EE.UU.", rn)
                div_perc_terms = "+".join(f'IF(ISNUMBER({iref(inputs_idx,"D8 · FINANZAS","Dividendos pagados a casa matriz (filial)",filial,rn)}),{iref(inputs_idx,"D8 · FINANZAS","Dividendos pagados a casa matriz (filial)",filial,rn)},0)' for filial in ("China", "Europa"))
                S.apply_cell(ws, cr["ga"], 7 + rn, value=f"={cprev}{cr['ga']}+{c}{cr['ben']}-IF(ISNUMBER({divacc}),{divacc},0)+({div_perc_terms})", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                vardlp = iref(inputs_idx, "D8 · FINANZAS", "Variación de deuda de largo plazo", "Casa matriz EE.UU.", rn)
                S.apply_cell(ws, cr["dlp"], 7 + rn, value=f"={cprev}{cr['dlp']}+IF(ISNUMBER({vardlp}),{vardlp},0)", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            else:
                S.apply_cell(ws, cr["capsoc"], 7 + rn, value=f"={cprev}{cr['capsoc']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                divpag = iref(inputs_idx, "D8 · FINANZAS", "Dividendos pagados a casa matriz (filial)", pais, rn)
                S.apply_cell(ws, cr["ga"], 7 + rn, value=f"={cprev}{cr['ga']}+{c}{cr['ben']}-IF(ISNUMBER({divpag}),{divpag},0)", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, cr["dlp"], 7 + rn, value="0", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["capad"], 7 + rn, value=f"={cprev}{cr['capad']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["pat"], 7 + rn, value=f"={c}{cr['capsoc']}+{c}{cr['capad']}+{c}{cr['ga']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            S.apply_cell(ws, cr["prest"], 7 + rn, value=f"={cprev}{cr['prest']}+({recibidos})-({enviados})", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["cxp"], 7 + rn, value=f"={c}{cr['gyc']}*$G${row_cxp_pct[pais]}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["dcp"], 7 + rn, value=f"={cprev}{cr['dcp']}+{c}{cr['cfaj']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["pasivos"], 7 + rn, value=f"={c}{cr['dlp']}+{c}{cr['dcp']}+{c}{cr['prest']}+{c}{cr['cxp']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            S.apply_cell(ws, cr["bal_ctrl"], 7 + rn, value=f"={c}{cr['activos']}-({c}{cr['pat']}+{c}{cr['pasivos']})", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

            # -------- FLUJO DE FONDOS por país (método indirecto, secuencia no circular) --------
            f_cfo = f"=({c}{cr['ben']}+{c}{cr['dep']})-({c}{cr['inv']}-{cprev}{cr['inv']})-({c}{cr['cxc']}-{cprev}{cr['cxc']})+({c}{cr['cxp']}-{cprev}{cr['cxp']})"
            S.apply_cell(ws, cr["cfo"], 7 + rn, value=f_cfo, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            f_cfi = f"=-{cref(ENGINE_PROD_SHEET, prod_out['pago_inv'][pais], rn)}+{cref(ENGINE_PROD_SHEET, prod_out['cobro_desinv'][pais], rn)}" if produce else "0"
            S.apply_cell(ws, cr["cfi"], 7 + rn, value=f_cfi, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            if pais == "EE.UU.":
                f_cff = (f"=-IF(ISNUMBER({divacc}),{divacc},0)+IF(ISNUMBER({emis}),{emis},0)-IF(ISNUMBER({recom}),{recom},0)"
                         f"+({div_perc_terms})+IF(ISNUMBER({vardlp}),{vardlp},0)+({recibidos})-({enviados})")
            else:
                f_cff = f"=-IF(ISNUMBER({divpag}),{divpag},0)+({recibidos})-({enviados})"
            S.apply_cell(ws, cr["cff"], 7 + rn, value=f_cff, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            efmin_raw = iref(inputs_idx, "B. Condiciones", "Efectivo mínimo de fin de ronda", "Todos los países (por país)", rn)
            efmin_ref = f"IF(ISNUMBER({efmin_raw}),{efmin_raw},2000000)"
            efectivo_antes = f"({cprev}{cr['caja']}+{c}{cr['cfo']}+{c}{cr['cfi']}+{c}{cr['cff']})"
            f_cfaj = f"=IF({efectivo_antes}<{efmin_ref},{efmin_ref}-{efectivo_antes},-MIN({cprev}{cr['dcp']},MAX(0,{efectivo_antes}-{efmin_ref})))"
            S.apply_cell(ws, cr["cfaj"], 7 + rn, value=f_cfaj, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["cfcambio"], 7 + rn, value=f"={c}{cr['cfo']}+{c}{cr['cfi']}+{c}{cr['cff']}+{c}{cr['cfaj']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
            S.apply_cell(ws, cr["cfap"], 7 + rn, value=f"={cprev}{cr['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, cr["cfci"], 7 + rn, value=f"={c}{cr['cfap']}+{c}{cr['cfcambio']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)

    # ============ Rellenar SECCIÓN B (GLOBAL P&L) ============
    # Ronda N.N (Fase 4, frontera dinámica): reemplaza el diccionario hist_pl_g (copia manual, a mano,
    # de los 9 valores Global de R0 y "n/a"/parcial para R1) por lectura DIRECTA de la sección propia
    # "Cuenta de resultados, miles USD, Global" del RDOS -- extract_hist_pais() YA la captura bajo
    # pais="Global" (no hacía falta un extractor nuevo). Verificado bottom-up con los 4 RDOS reales
    # (R0-R3): Global.Ingresos/Costos/EBITDA/Depreciación/EBIT/GFN/EBT/Impuesto/Beneficio están
    # publicados directamente (no hay que sumarlos desde los países) -- y la suma de "Beneficio de la
    # ronda" de los 3 países coincide EXACTO (diff=0.00 en las 4 rondas) con el Beneficio Global
    # publicado, así que el control de consistencia (glob['ctrl']) ya no necesita quedar en "n/a": se
    # activa con la misma fórmula que usan las rondas proyectadas.
    GLOB_PL_LABELS = {"ing": "Ingresos por ventas", "costos": "Costos y gastos totales",
                       "ebitda": "Beneficio operativo antes de depreciación (EBITDA)",
                       "dep": "Depreciación de Activos Fijos", "ebit": "Beneficio operativo (EBIT)",
                       "gfn": "Gastos financieros netos", "ebt": "Beneficio antes de impuestos",
                       "imp": "Impuesto sobre el beneficio", "ben": "Beneficio de la ronda"}
    for rn in RONDAS:
        c = col(rn)
        if rn in rondas_reales:
            for k, lbl in GLOB_PL_LABELS.items():
                v = hv(rn, "Global", "pl", lbl)
                S.apply_cell(ws, glob[k], 7 + rn, value=(v if v is not None else "n/d"), kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            f_sum_ben = "+".join(f"{c}{country_rows[p]['ben']}" for p in PAISES)
            S.apply_cell(ws, glob["ctrl"], 7 + rn, value=f"={c}{glob['ben']}-({f_sum_ben})", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            continue
        f_ing = "=" + "+".join(f"{c}{country_rows[p]['mercado']}" for p in PAISES)
        S.apply_cell(ws, glob["ing"], 7 + rn, value=f_ing, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_costos = "=" + "+".join(f"({c}{country_rows[p]['cogs']}+{c}{country_rows[p]['transp']}+{c}{country_rows[p]['car']}+{c}{country_rows[p]['ginv']}+{c}{country_rows[p]['energia']}+{c}{country_rows[p]['agua']}+{c}{country_rows[p]['carbono']}+{c}{country_rows[p]['id_tot']}+{c}{country_rows[p]['promo']}+{c}{country_rows[p]['admin']})" for p in PAISES)
        S.apply_cell(ws, glob["costos"], 7 + rn, value=f_costos, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        S.apply_cell(ws, glob["ebitda"], 7 + rn, value=f"={c}{glob['ing']}-{c}{glob['costos']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_dep = "=" + "+".join(f"{c}{country_rows[p]['dep']}" for p in PAISES)
        S.apply_cell(ws, glob["dep"], 7 + rn, value=f_dep, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["ebit"], 7 + rn, value=f"={c}{glob['ebitda']}-{c}{glob['dep']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_gfn = "=" + "+".join(f"{c}{country_rows[p]['gfn']}" for p in PAISES)
        S.apply_cell(ws, glob["gfn"], 7 + rn, value=f_gfn, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["ebt"], 7 + rn, value=f"={c}{glob['ebit']}-{c}{glob['gfn']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_imp = "=" + "+".join(f"{c}{country_rows[p]['imp']}" for p in PAISES)
        S.apply_cell(ws, glob["imp"], 7 + rn, value=f_imp, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["ben"], 7 + rn, value=f"={c}{glob['ebt']}-{c}{glob['imp']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_sum_ben = "+".join(f"{c}{country_rows[p]['ben']}" for p in PAISES)
        S.apply_cell(ws, glob["ctrl"], 7 + rn, value=f"={c}{glob['ben']}-({f_sum_ben})", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

    # ============ Fase 2, cambio 2: Rellenar SECCIÓN C (GLOBAL BALANCE) = SUMA REAL de los 3 países
    # (TODAS las rondas, R0-R12 -- ya NO hay pool ni historicos Global hardcodeados aparte: la Sección
    # A de cada país arriba YA es 100% consistente con RDOS para R0/R1, así que el Global sumado
    # también lo es -- esto ELIMINA la inconsistencia heredada de v1.2 (~1.508,8 MM USD en R1,
    # originada en mezclar el Inventario RDOS por país con la cifra "por tecnología" de motor viejo;
    # ver Informe_V2.md). nacc (acciones en circulación) sigue como serie propia, no es una suma de
    # países (concepto de casa matriz único). ============
    for rn in RONDAS:
        c = col(rn)
        cprev = col(rn - 1) if rn > 0 else None
        f_af = "=" + "+".join(f"{c}{country_rows[p]['af']}" for p in PAISES)
        S.apply_cell(ws, glob["af"], 7 + rn, value=f_af, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_inv = "=" + "+".join(f"{c}{country_rows[p]['inv']}" for p in PAISES)
        S.apply_cell(ws, glob["inv"], 7 + rn, value=f_inv, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_cxc = "=" + "+".join(f"{c}{country_rows[p]['cxc']}" for p in PAISES)
        S.apply_cell(ws, glob["cxc"], 7 + rn, value=f_cxc, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_caja = "=" + "+".join(f"{c}{country_rows[p]['caja']}" for p in PAISES)
        S.apply_cell(ws, glob["caja"], 7 + rn, value=f_caja, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["activos"], 7 + rn, value=f"={c}{glob['af']}+{c}{glob['inv']}+{c}{glob['cxc']}+{c}{glob['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_capsoc = "=" + "+".join(f"{c}{country_rows[p]['capsoc']}" for p in PAISES)
        S.apply_cell(ws, glob["capsoc"], 7 + rn, value=f_capsoc, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        if rn in rondas_reales:
            v_nacc_hist = hist_pais.get(rn, {}).get("_nacc")
            S.apply_cell(ws, glob["nacc"], 7 + rn, value=(v_nacc_hist if v_nacc_hist is not None else "n/d"), kind="historico", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
        else:
            emis = iref(inputs_idx, "D8 · FINANZAS", "Emisión de acciones", "Casa matriz EE.UU.", rn)
            recom = iref(inputs_idx, "D8 · FINANZAS", "Recompra de acciones", "Casa matriz EE.UU.", rn)
            precio_ref_nacc = iref(inputs_idx, "B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", rn)
            f_nacc = f"={cprev}{glob['nacc']}+IFERROR((IF(ISNUMBER({emis}),{emis},0)-IF(ISNUMBER({recom}),{recom},0))/{precio_ref_nacc},0)"
            S.apply_cell(ws, glob["nacc"], 7 + rn, value=f_nacc, kind="calculo", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=8)
        f_capad = "=" + "+".join(f"{c}{country_rows[p]['capad']}" for p in PAISES)
        S.apply_cell(ws, glob["capad"], 7 + rn, value=f_capad, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_gan = "=" + "+".join(f"{c}{country_rows[p]['ga']}" for p in PAISES)
        S.apply_cell(ws, glob["gan"], 7 + rn, value=f_gan, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["pat"], 7 + rn, value=f"={c}{glob['capsoc']}+{c}{glob['capad']}+{c}{glob['gan']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        f_dlp = "=" + "+".join(f"{c}{country_rows[p]['dlp']}" for p in PAISES)
        S.apply_cell(ws, glob["dlp"], 7 + rn, value=f_dlp, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_dcp = "=" + "+".join(f"{c}{country_rows[p]['dcp']}" for p in PAISES)
        S.apply_cell(ws, glob["dcp"], 7 + rn, value=f_dcp, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_cxp = "=" + "+".join(f"{c}{country_rows[p]['cxp']}" for p in PAISES)
        S.apply_cell(ws, glob["cxp"], 7 + rn, value=f_cxp, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_prest_ctrl = "=" + "+".join(f"{c}{country_rows[p]['prest']}" for p in PAISES)
        S.apply_cell(ws, glob["prest_ctrl"], 7 + rn, value=f_prest_ctrl, kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["pasivos"], 7 + rn, value=f"={c}{glob['dlp']}+{c}{glob['dcp']}+{c}{glob['cxp']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        S.apply_cell(ws, glob["bal_ctrl"], 7 + rn, value=f"={c}{glob['activos']}-({c}{glob['pat']}+{c}{glob['pasivos']})", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
    ws.cell(row=glob["bal_ctrl"], column=6).value = ("Fase 2: el Global es ahora la SUMA de los 3 Balances por país (100% consistentes con RDOS en "
        "R0/R1 -- ver Sección A/BAL país arriba). Esto corrige la inconsistencia heredada de v1.2 (~-1.508,8 MM USD en R1, que mezclaba el "
        "Inventario RDOS agregado con la cifra 'por tecnología' de motor viejo). Los Pasivos Totales de este renglón NO incluyen Préstamos "
        "internos por diseño de v1.2 (ver fila 'Pasivos Totales' arriba) -- el control de préstamos internos está separado en la fila "
        "'Control: Σ Préstamos internos' (debe ser exactamente 0).")

    # ============ Fase 2, cambio 2: Rellenar SECCIÓN D (GLOBAL CASH FLOW) = SUMA REAL de los 3 países ============
    for rn in RONDAS:
        c = col(rn)
        if rn == 0:
            for kk in ("cfo", "cfi", "cff", "cfaj"):
                S.apply_cell(ws, glob[kk], 7 + rn, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
            S.apply_cell(ws, glob["cfcambio"], 7 + rn, value=0.0, kind="historico", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, glob["cfap"], 7 + rn, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
            S.apply_cell(ws, glob["cfci"], 7 + rn, value=f"={c}{glob['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, glob["cf_ctrl"], 7 + rn, value="n/a (Ronda 0)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
            continue
        cprev = col(rn - 1)
        f_cfo = "=" + "+".join(f"{c}{country_rows[p]['cfo']}" for p in PAISES)
        S.apply_cell(ws, glob["cfo"], 7 + rn, value=f_cfo, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_cfi = "=" + "+".join(f"{c}{country_rows[p]['cfi']}" for p in PAISES)
        S.apply_cell(ws, glob["cfi"], 7 + rn, value=f_cfi, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_cff = "=" + "+".join(f"{c}{country_rows[p]['cff']}" for p in PAISES)
        S.apply_cell(ws, glob["cff"], 7 + rn, value=f_cff, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        if rn in rondas_reales:
            # Ronda N.N (Fase 4): generaliza el boundary fijo (antes solo rn==1) -- cr["cfaj"] por
            # país es texto "n/a" para CUALQUIER ronda histórica (rn>0) desde antes de este cambio
            # (el "ajuste automático de deuda CP" es una cifra derivada del propio modelo, no un dato
            # que el RDOS publique), así que sumarlo con "+" para una ronda real más allá de R1 (p.ej.
            # R2/R3 al pasar a REAL) daba #VALUE! antes de este fix.
            S.apply_cell(ws, glob["cfaj"], 7 + rn, value="n/a (histórico, ver RDOS)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
        else:
            f_cfaj = "=" + "+".join(f"{c}{country_rows[p]['cfaj']}" for p in PAISES)
            S.apply_cell(ws, glob["cfaj"], 7 + rn, value=f_cfaj, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        f_cfcambio = "=" + "+".join(f"{c}{country_rows[p]['cfcambio']}" for p in PAISES)
        S.apply_cell(ws, glob["cfcambio"], 7 + rn, value=f_cfcambio, kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        S.apply_cell(ws, glob["cfap"], 7 + rn, value=f"={cprev}{glob['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)
        S.apply_cell(ws, glob["cfci"], 7 + rn, value=f"={c}{glob['caja']}", kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, bold=True, size=8)
        S.apply_cell(ws, glob["cf_ctrl"], 7 + rn, value=f"={c}{glob['cfci']}-{c}{glob['caja']}", kind="alerta", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=8)

    ws.sheet_state = "hidden"
    return ws, row_idx, r, country_rows, glob, dict(cxc_pct=row_cxc_pct, cxp_pct=row_cxp_pct)


# ======================================================================================
# 02_ESTADOS_PROYECTADOS — vista GLOBAL de gestión, solo lectura (links a _ENGINE_FINANCIERO)
# ======================================================================================
def build_estados_proyectados(wb, glob, inputs_idx, country_rows, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(ESTADOS_SHEET)
    # Hotfix visual: ncols y set_col_widths deben cubrir TODAS las columnas físicas de la hoja
    # (A,B + C-F espaciador oculto + G-S rondas = 2+4+13=19), no solo hasta la última ronda contada
    # sin el espaciador (15) -- si no, el fill de título/sección y el ancho de columna quedan sin
    # definir a partir de la columna P (R9-R12), viéndose "corridos"/con formato roto.
    ncols = 2 + 4 + len(RONDAS)
    S.set_col_widths(ws, [42, 10] + [13] * 4 + [13] * len(RONDAS))
    S.style_title_row(ws, 1, "02_ESTADOS_PROYECTADOS — Estados financieros con selector GLOBAL/EE.UU./CHINA/EUROPA (B2) y detalle por país agrupado ([+]/[-] a la izquierda)", ncols)
    ws.sheet_properties.outlinePr.summaryBelow = False

    S.apply_cell(ws, 2, 1, value="SELECTOR DE VISTA (celda B2) →", kind="calculo", align=S.ALIGN_LEFT, bold=True, size=10)
    S.apply_cell(ws, 2, 2, value="GLOBAL", kind="input", align=S.ALIGN_CENTER, bold=True, size=10)
    # Hotfix regional: lista oculta en columna U (21), en vez de la cadena separada por comas --
    # rompía en Excel es-AR/es-LA (separador de listas ";"). El orden (filas U1:U5) debe coincidir
    # con el orden posicional que usa CHOOSE en choose_formula() más abajo.
    VISTA_LIST_COL = 21  # U
    vista_list_letter = get_column_letter(VISTA_LIST_COL)
    for i, opt in enumerate(["GLOBAL", "EE.UU.", "CHINA", "EUROPA", "VISTA DESPLEGADA"], start=1):
        S.apply_cell(ws, i, VISTA_LIST_COL, value=opt, kind="plain", size=8)
    vista_list_range = f"${vista_list_letter}$1:${vista_list_letter}$5"
    ws.column_dimensions[vista_list_letter].hidden = True
    dv_vista = DataValidation(type="list", formula1=vista_list_range, allow_blank=False)
    ws.add_data_validation(dv_vista)
    dv_vista.add("B2")
    ws.cell(row=2, column=3,
            value=("GLOBAL = consolidado (Σ 3 países, con eliminaciones). EE.UU./CHINA/EUROPA = ese país únicamente. VISTA DESPLEGADA = igual a GLOBAL en la fila principal "
                   "(pensada para cuando se quiere ver todo el detalle por país expandido de entrada). El detalle por país de cada línea está SIEMPRE presente debajo, agrupado con el "
                   "outline nativo de Excel/LibreOffice -- usar [+]/[-] a la izquierda de la grilla para expandir/colapsar, independientemente de lo que diga B2.")
            ).font = S.f_normal(8, italic=True, color=S.COLOR_HIST)
    ws.cell(row=2, column=3).alignment = S.ALIGN_LEFT_WRAP
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=ncols)
    ws.row_dimensions[2].height = 30
    S.apply_cell(ws, 3, 1, value="Detalle por país: Ingresos = 'Ingresos de mercados' (externo, sin transferencias internas). Costos = Costos y gastos totales EXCLUYENDO costos de productos importados (misma base de eliminación que GLOBAL). El resto de las líneas suma exacto a GLOBAL. Celdas en VERDE = link, no editar aquí.",
                 kind="plain", align=S.ALIGN_LEFT_WRAP, size=8, italic=True)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    ws.row_dimensions[3].height = 26

    header_row = 5
    S.style_header_row(ws, header_row, ["Concepto", "Unidad", "", "", "", ""] + [f"R{r}" for r in RONDAS])
    r = header_row + 1

    def sec(title):
        nonlocal r
        S.style_section_row(ws, r, title, ncols)
        r += 1

    def activo_ref(rn):
        return iref(inputs_idx, "A. Identificación", "RONDA_ACTIVA (calculado)", "-", rn)

    def guarded(rn, inner_formula_no_eq):
        """Cambio 2 (v1.1): en rondas no activas (sin Estado+inputs mínimos cargados), la celda
        queda VACÍA ("") en vez de mostrar un cálculo sin sentido con inputs en blanco.
        Nota (Fase 4, frontera dinámica): el boundary de acá abajo se deja fijo en (0,1) a propósito
        -- NO hace falta generalizarlo a rondas_reales, porque RONDA_ACTIVA (01_INPUTS, ya
        generalizada) es un valor LITERAL =VERDADERO para cualquier ronda en rondas_reales, así que
        IF(RONDA_ACTIVA,inner,"") ya se reduce a "inner" para esas rondas de todas formas -- ambas
        ramas son equivalentes, esto es solo una simplificación cosmética de la fórmula escrita."""
        if rn in (0, 1):
            return f"={inner_formula_no_eq}"
        return f'=IF({activo_ref(rn)},{inner_formula_no_eq},"")'

    def choose_formula(g_ref, us_ref, cn_ref, eu_ref):
        # Hotfix regional: MATCH contra el rango oculto $U$1:$U$5 (no un array en línea {"A","B",...})
        # -- el array en línea rompía en Excel es-AR/es-LA por el separador de matrices regional.
        return f'CHOOSE(MATCH($B$2,{vista_list_range},0),{g_ref},{us_ref},{cn_ref},{eu_ref},{g_ref})'

    def line_country(label, key_global, key_country=None, bold=False, combo=None, numfmt=S.NUM_MONEY):
        """Fila PRINCIPAL (selector CHOOSE/MATCH B2) + 3 filas de detalle por país SIEMPRE presentes,
        agrupadas con outline nativo (colapsadas por defecto). Fase 2, cambio 3.
        combo(pais, rn) -> fórmula (sin '=') para países cuya línea de detalle no es un simple link
        1:1 a country_rows (p.ej. Costos, que excluye 'Costos de productos importados' para sumar
        exacto a GLOBAL, igual criterio de eliminación de consolidación)."""
        nonlocal r
        key_country = key_country or key_global
        S.apply_cell(ws, r, 1, value=label, kind="calculo", align=S.ALIGN_LEFT, bold=bold, size=9)
        S.apply_cell(ws, r, 2, value="USD", kind="calculo", align=S.ALIGN_CENTER, size=8)
        for rn in RONDAS:
            g_ref = cref(ENGINE_FIN_SHEET, glob[key_global], rn)
            if combo:
                us_ref, cn_ref, eu_ref = combo("EE.UU.", rn), combo("China", rn), combo("Europa", rn)
            else:
                us_ref = cref(ENGINE_FIN_SHEET, country_rows["EE.UU."][key_country], rn)
                cn_ref = cref(ENGINE_FIN_SHEET, country_rows["China"][key_country], rn)
                eu_ref = cref(ENGINE_FIN_SHEET, country_rows["Europa"][key_country], rn)
            S.apply_cell(ws, r, 7 + rn, value=guarded(rn, choose_formula(g_ref, us_ref, cn_ref, eu_ref)), kind="link", numfmt=numfmt, align=S.ALIGN_RIGHT, bold=bold, size=9)
        row_principal = r
        r += 1
        for pais in PAISES:
            S.apply_cell(ws, r, 1, value=f"      {pais}", kind="calculo", align=S.ALIGN_LEFT, italic=True, size=8)
            S.apply_cell(ws, r, 2, value="USD", kind="calculo", align=S.ALIGN_CENTER, size=7)
            for rn in RONDAS:
                ref = combo(pais, rn) if combo else cref(ENGINE_FIN_SHEET, country_rows[pais][key_country], rn)
                S.apply_cell(ws, r, 7 + rn, value=guarded(rn, ref), kind="link", numfmt=numfmt, align=S.ALIGN_RIGHT, size=8, italic=True)
            ws.row_dimensions[r].outline_level = 1
            ws.row_dimensions[r].hidden = True
            r += 1
        return row_principal

    def combo_costos(pais, rn):
        gyc_ref = cref(ENGINE_FIN_SHEET, country_rows[pais]["gyc"], rn)
        imp_ref = cref(ENGINE_FIN_SHEET, country_rows[pais]["imp_cost"], rn)
        return f"({gyc_ref}-{imp_ref})"

    sec("ESTADO DE RESULTADOS")
    row_ing = line_country("Ingresos por ventas", "ing", key_country="mercado", bold=True)
    row_costos = line_country("Costos y gastos totales", "costos", combo=combo_costos)
    row_ebitda = line_country("EBITDA", "ebitda", bold=True)
    S.apply_cell(ws, r, 1, value="Margen EBITDA (%) — sigue la vista de B2 (sin detalle por país, métrica derivada)", kind="calculo", align=S.ALIGN_LEFT, size=9, italic=True)
    S.apply_cell(ws, r, 2, value="%", kind="calculo", align=S.ALIGN_CENTER, size=8)
    for rn in RONDAS:
        c = col(rn)
        S.apply_cell(ws, r, 7 + rn, value=guarded(rn, f'IFERROR({c}{row_ebitda}/{c}{row_ing},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9, italic=True)
    r += 1
    row_dep = line_country("Depreciación", "dep")
    row_ebit = line_country("EBIT", "ebit", bold=True)
    row_gfn = line_country("Gastos financieros netos", "gfn")
    row_ebt = line_country("Beneficio antes de impuestos", "ebt", bold=True)
    row_imp = line_country("Impuesto sobre el beneficio", "imp")
    row_ben = line_country("Beneficio de la ronda (neto)", "ben", bold=True)
    r += 1

    sec("BALANCE")
    row_af = line_country("Activo fijo (neto)", "af")
    row_inv = line_country("Inventario", "inv")
    row_cxc = line_country("Cuentas por Cobrar", "cxc")
    row_caja = line_country("Efectivo y equivalentes", "caja")
    row_activos = line_country("Activos Totales", "activos", bold=True)
    row_dlp = line_country("Deudas a largo plazo", "dlp")
    row_dcp = line_country("Deudas a corto plazo (automática)", "dcp")
    row_cxp = line_country("Cuentas por Pagar", "cxp")
    row_pasivos = line_country("Pasivos Totales", "pasivos", bold=True)
    row_capsoc = line_country("Capital social", "capsoc")
    row_capad = line_country("Capital adicional desembolsado", "capad")
    row_gan = line_country("Ganancias acumuladas", "gan", key_country="ga")
    row_pat = line_country("Patrimonio neto", "pat", bold=True)
    r += 1

    sec("FLUJO DE FONDOS (método indirecto)")
    row_cfo = line_country("Total actividades operativas", "cfo")
    row_cfi = line_country("Total actividades de inversión", "cfi")
    row_cff = line_country("Total actividades financieras (D8)", "cff")
    row_cfaj = line_country("Ajuste automático de deuda de corto plazo", "cfaj")
    row_cfcambio = line_country("Cambio en efectivo y equivalentes", "cfcambio", bold=True)
    row_cfap = line_country("Efectivo de apertura", "cfap")
    row_cfci = line_country("Efectivo de cierre", "cfci", bold=True)

    # Hotfix visual: columnas C-F se ocultan más abajo (main()) -- el freeze pasa a G (primera
    # columna visible después de "Unidad"), para que el límite de inmovilizado coincida con lo que
    # realmente se ve en pantalla.
    ws.freeze_panes = "G6"
    return ws


# ======================================================================================
# 03_RATIOS — ratios de gestión, cortos y útiles para decisión (por ronda)
# ======================================================================================
def build_ratios(wb, glob, inputs_idx, cond, decisiones, prod_out, mkt_out, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(RATIOS_SHEET)
    # Hotfix visual: idem 02_ESTADOS_PROYECTADOS -- ncols/anchos deben cubrir A,B + C-F (espaciador
    # oculto) + G-S (rondas) = 2+4+13=19, no solo 15.
    ncols = 2 + 4 + len(RONDAS)
    S.set_col_widths(ws, [46, 10] + [13] * 4 + [13] * len(RONDAS))
    S.style_title_row(ws, 1, "03_RATIOS — Ratios de crecimiento, rentabilidad, financieros y operativos (GLOBAL, por ronda)", ncols)
    header_row = 3
    S.style_header_row(ws, header_row, ["Concepto", "Unidad", "", "", "", ""] + [f"R{r}" for r in RONDAS])
    r = header_row + 1

    def sec(title):
        nonlocal r
        S.style_section_row(ws, r, title, ncols)
        r += 1

    def newrow(label, unidad="", note=None, italic=False):
        nonlocal r
        S.apply_cell(ws, r, 1, value=label, kind="calculo", align=S.ALIGN_LEFT, size=9, italic=italic)
        S.apply_cell(ws, r, 2, value=unidad, kind="calculo", align=S.ALIGN_CENTER, size=8)
        rr = r
        r += 1
        return rr

    fin = lambda rn, key: cref(ENGINE_FIN_SHEET, glob[key], rn)

    def activo_ref(rn):
        return iref(inputs_idx, "A. Identificación", "RONDA_ACTIVA (calculado)", "-", rn)

    def guarded(rn, inner_no_eq):
        """Cambio 2 (v1.1): idéntico criterio que en 02_ESTADOS_PROYECTADOS -- R0/R1 sin cambios,
        R2 (activa) sin cambios de valor, R3-R12 no activas quedan en blanco (""). Nota (Fase 4):
        boundary fijo a propósito, ver comentario equivalente en build_estados_proyectados -- es
        cosmético (RONDA_ACTIVA ya es =VERDADERO literal para cualquier ronda real)."""
        if rn in (0, 1):
            return f"={inner_no_eq}"
        return f'=IF({activo_ref(rn)},{inner_no_eq},"")'

    sec("CRECIMIENTO")
    row_growth = newrow("Crecimiento de Ingresos vs. ronda anterior", "%")
    sec("RENTABILIDAD")
    row_margen_ebitda = newrow("Margen EBITDA", "%")
    row_ros = newrow("ROS — Beneficio neto / Ingresos", "%")
    row_roe = newrow("ROE — Beneficio neto / Patrimonio promedio (apertura,cierre)", "%", note="REGLA VERIFICADA manual cap.13")
    row_roce = newrow("ROCE — EBIT / Capital empleado promedio (apertura,cierre)", "%", note="REGLA VERIFICADA manual cap.13")
    row_wacc = newrow("WACC (proxy simplificado)", "%", italic=True)
    row_roce_wacc = newrow("ROCE − WACC (creación de valor)", "p.p.", italic=True)
    row_eps = newrow("EPS — Beneficio neto / N° acciones", "USD/acción")
    sec("CAJA Y DEUDA")
    row_fcf = newrow("Flujo de caja libre (FCF) = CFO − CAPEX", "USD")
    row_deuda_fin = newrow("Deuda financiera total (LP+CP)", "USD")
    row_deuda_pat = newrow("Deuda financiera / Patrimonio", "x")
    row_payout = newrow("Payout ratio = Dividendos / Beneficio neto", "%")
    sec("VALUACIÓN (PROXY — no es una predicción del algoritmo real de CESIM)")
    row_precio_proxy = newrow("Precio de acción de referencia (Condición, PROXY)", "USD", italic=True)
    row_pe_proxy = newrow("P/E proxy = Precio referencia / EPS [PROXY]", "x", italic=True)
    row_divyield_proxy = newrow("Dividend yield proxy = Dividendo por acción / Precio [PROXY]", "%", italic=True)
    row_retorno_periodo = newrow(
        "Retorno del accionista del período (PROXY) = ΔPrecio referencia × (1+Dividend yield) [PROXY]", "%", italic=True,
        note="Fórmula verificada EXACTA contra Ronda 1 real (RDOS 'Retorno total acumulado del accionista (p.a.), %' = 18,725103%, "
             "reproducido con (Precio_R1/Precio_R0−1 +1)×(1+Rendimiento de dividendos R1) = 18,7251035%). El manual (cap. 2.1) confirma "
             "cualitativamente que el ratio combina evolución del precio de la acción + dividendos pagados; la fórmula exacta de combinación "
             "(multiplicativa) fue inferida y verificada numéricamente contra ese único dato real disponible, no publicada como tal por CESIM.")
    row_retorno_acum_idx = newrow("Índice acumulado del accionista (PROXY, base Ronda 0 = 1,00) [memo]", "x", italic=True)
    row_retorno_acum = newrow(
        "Retorno total acumulado del accionista (PROXY, anualizado desde R0) [PROXY]", "%", italic=True,
        note="= Índice acumulado^(1/ronda) − 1. En R1 coincide EXACTO con el dato real (ver nota de la fila de período). "
             "De R2 en adelante depende de 'Precio de acción de referencia' (una Condición dada por ronda, no una predicción propia del "
             "precio de mercado) y del dividendo efectivamente decidido -- es la mejor aproximación posible sin el algoritmo real de "
             "valuación de CESIM (no publicado), y así debe leerse: PROXY, no una réplica de la fórmula oficial.")
    sec("OPERATIVO — CAPACIDAD, INVENTARIO, DEMANDA")
    row_util_us = newrow("Utilización de capacidad — EE.UU.", "%")
    row_util_cn = newrow("Utilización de capacidad — China", "%")
    row_insatisfecha = newrow("Demanda insatisfecha CADIZ (total, todas tech/mercados)", "u.")
    row_cuota_prom = newrow("Cuota de mercado CADIZ (ponderada por volumen, 3 mercados)", "%")
    sec("POR VEHÍCULO")
    row_ing_veh = newrow("Ingresos por vehículo vendido", "USD/u.")
    row_ebitda_veh = newrow("EBITDA por vehículo vendido", "USD/u.")

    for rn in RONDAS:
        c = col(rn)
        cprev = col(rn - 1) if rn > 0 else None
        ing_ref, ing_prev = fin(rn, "ing"), (fin(rn - 1, "ing") if rn > 0 else None)
        # PENDIENTE (Fase 4, NO generalizado en este pase): este bloque muestra "n/a" para R0/R1 sin
        # tocar rondas_reales -- ahora que _ENGINE_FINANCIERO sí calcula P&L/Balance real vía RDOS
        # para CUALQUIER ronda real (no solo R0/R1), en principio estos ratios podrían calcularse por
        # fórmula también para R1 en adelante (fin(rn,*) ya tiene dato real). No se generalizó acá por
        # riesgo/tiempo -- requiere verificar que los inputs de "Condición" que usan estas fórmulas
        # (tasa de interés, "Valor de referencia de la acción", etc.) estén poblados para la ronda en
        # cuestión antes de confiar en el resultado. Ver informe de este pase.
        if rn == 0 or rn == 1:
            for rr in [row_growth, row_margen_ebitda, row_ros, row_roe, row_roce, row_wacc, row_roce_wacc, row_eps,
                       row_fcf, row_payout, row_pe_proxy, row_divyield_proxy]:
                S.apply_cell(ws, rr, 7 + rn, value="n/a (ver 02_ESTADOS_PROYECTADOS)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
        else:
            S.apply_cell(ws, row_growth, 7 + rn, value=guarded(rn, f'IFERROR(({ing_ref}-{ing_prev})/{ing_prev},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            ebitda_ref = fin(rn, "ebitda")
            S.apply_cell(ws, row_margen_ebitda, 7 + rn, value=guarded(rn, f'IFERROR({ebitda_ref}/{ing_ref},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            ben_ref = fin(rn, "ben")
            S.apply_cell(ws, row_ros, 7 + rn, value=guarded(rn, f'IFERROR({ben_ref}/{ing_ref},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            pat_prev, pat_now = fin(rn - 1, "pat"), fin(rn, "pat")
            S.apply_cell(ws, row_roe, 7 + rn, value=guarded(rn, f'IFERROR({ben_ref}/(({pat_prev}+{pat_now})/2),"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            ebit_ref = fin(rn, "ebit")
            tasa_ref = iref(inputs_idx, "B. Condiciones", "Tasa de impuesto corporativo", "EE.UU.", rn)
            dlp_prev, dcp_prev = fin(rn - 1, "dlp"), fin(rn - 1, "dcp")
            dlp_now, dcp_now = fin(rn, "dlp"), fin(rn, "dcp")
            cap_emp_prev = f"({dlp_prev}+{dcp_prev}+{pat_prev})"
            cap_emp_now = f"({dlp_now}+{dcp_now}+{pat_now})"
            cap_emp_avg = f"(({cap_emp_prev}+{cap_emp_now})/2)"
            S.apply_cell(ws, row_roce, 7 + rn, value=guarded(rn, f'IFERROR({ebit_ref}/{cap_emp_avg},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            tlp_ref = iref(inputs_idx, "B. Condiciones", "Tasa de interés — deuda a largo plazo", "Casa matriz EE.UU.", rn)
            S.apply_cell(ws, row_wacc, 7 + rn,
                          value=guarded(rn, f'IFERROR(({pat_now}/{cap_emp_now})*0.12+(({dlp_now}+{dcp_now})/{cap_emp_now})*{tlp_ref}*(1-{tasa_ref}),"")'),
                          kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row_roce_wacc, 7 + rn, value=guarded(rn, f'IFERROR({c}{row_roce}-{c}{row_wacc},"")'), kind="calculo", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9)
            precio_ref_cell = iref(inputs_idx, "B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", rn)
            n_acc = fin(rn, "nacc")
            S.apply_cell(ws, row_eps, 7 + rn, value=guarded(rn, f'IFERROR({ben_ref}/({n_acc}),"")'), kind="calculo", numfmt='0.00', align=S.ALIGN_RIGHT, size=9)
            cfo_ref = fin(rn, "cfo")
            capex_ref = "+".join(cref(ENGINE_PROD_SHEET, prod_out["pago_inv"][p], rn) for p in AREAS)
            S.apply_cell(ws, row_fcf, 7 + rn, value=guarded(rn, f"{cfo_ref}-({capex_ref})"), kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
            divacc_ref = iref(inputs_idx, "D8 · FINANZAS", "Dividendos pagados a accionistas (monto)", "Casa matriz EE.UU.", rn)
            S.apply_cell(ws, row_payout, 7 + rn, value=guarded(rn, f'IFERROR(IF(ISNUMBER({divacc_ref}),{divacc_ref},0)/{ben_ref},"")'), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row_pe_proxy, 7 + rn, value=guarded(rn, f'IFERROR({precio_ref_cell}/{c}{row_eps},"")'), kind="calculo", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=9, italic=True)
            S.apply_cell(ws, row_divyield_proxy, 7 + rn,
                          value=guarded(rn, f'IFERROR((IF(ISNUMBER({divacc_ref}),{divacc_ref},0)/({n_acc}))/{precio_ref_cell},"")'),
                          kind="calculo", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9, italic=True)

        # Retorno total acumulado del accionista (PROXY) -- ver nota de la fila para la verificación
        # contra R1 real. R0 = base del índice (1,00); R1 = valor histórico hardcodeado (coincide
        # exacto con el dato real de RDOS); R2 en adelante = fórmula viva (Δ precio proxy × (1+div yield proxy)).
        if rn == 0:
            S.apply_cell(ws, row_retorno_periodo, 7 + rn, value="n/a (Ronda 0, sin ronda anterior)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
            S.apply_cell(ws, row_retorno_acum_idx, 7 + rn, value=1, kind="historico", numfmt='0.0000', align=S.ALIGN_RIGHT, size=9, italic=True)
            S.apply_cell(ws, row_retorno_acum, 7 + rn, value="n/a (Ronda 0, sin ronda anterior)", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
        else:
            if rn == 1:
                # Histórico real (RDOS Ronda 1, "Retorno total acumulado del accionista (p.a.), %") --
                # NO se recalcula por fórmula porque row_divyield_proxy está en blanco en R1 (ver bloque
                # "n/a" más arriba, idéntico criterio que EPS/FCF/etc.: R0/R1 migrados, no recalculados).
                S.apply_cell(ws, row_retorno_periodo, 7 + rn, value=0.18725103, kind="historico", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9, italic=True)
            else:
                precio_now_ref = iref(inputs_idx, "B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", rn)
                precio_prev_ref = iref(inputs_idx, "B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", rn - 1)
                f_periodo = f'IFERROR(({precio_now_ref}/{precio_prev_ref})*(1+{c}{row_divyield_proxy})-1,"")'
                S.apply_cell(ws, row_retorno_periodo, 7 + rn, value=guarded(rn, f_periodo), kind="calculo", numfmt=S.NUM_PCT2, align=S.ALIGN_RIGHT, size=9, italic=True)
            f_idx = f'{cprev}{row_retorno_acum_idx}*(1+{c}{row_retorno_periodo})'
            S.apply_cell(ws, row_retorno_acum_idx, 7 + rn, value=guarded(rn, f_idx), kind="calculo", numfmt='0.0000', align=S.ALIGN_RIGHT, size=9, italic=True)
            f_acum = f'IFERROR({c}{row_retorno_acum_idx}^(1/{rn})-1,"")'
            S.apply_cell(ws, row_retorno_acum, 7 + rn, value=guarded(rn, f_acum), kind="calculo", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9, italic=True)

        dlp_ref, dcp_ref = fin(rn, "dlp"), fin(rn, "dcp")
        S.apply_cell(ws, row_deuda_fin, 7 + rn, value=guarded(rn, f"{dlp_ref}+{dcp_ref}"), kind="calculo", numfmt=S.NUM_MONEY, align=S.ALIGN_RIGHT, size=9)
        pat_ref = fin(rn, "pat")
        S.apply_cell(ws, row_deuda_pat, 7 + rn, value=guarded(rn, f'IFERROR(({dlp_ref}+{dcp_ref})/{pat_ref},"")'), kind="calculo", numfmt=S.NUM_X, align=S.ALIGN_RIGHT, size=9)
        precio_ref_cell2 = iref(inputs_idx, "B. Condiciones", "Valor de referencia de la acción (proxy para EPS/emisión-recompra)", "Casa matriz EE.UU.", rn)
        S.apply_cell(ws, row_precio_proxy, 7 + rn, value=guarded(rn, f"{precio_ref_cell2}"), kind="link", numfmt='0.00', align=S.ALIGN_RIGHT, size=9, italic=True)

        cap_us_ref = cref(ENGINE_PROD_SHEET, prod_out["cap"]["EE.UU."], rn)
        prodtot_us_ref = cref(ENGINE_PROD_SHEET, prod_out["prodtot"]["EE.UU."], rn) if "prodtot" in prod_out else None
        util_us_ref = cref(ENGINE_PROD_SHEET, prod_out["util"]["EE.UU."], rn)
        util_cn_ref = cref(ENGINE_PROD_SHEET, prod_out["util"]["China"], rn)
        S.apply_cell(ws, row_util_us, 7 + rn, value=guarded(rn, util_us_ref), kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)
        S.apply_cell(ws, row_util_cn, 7 + rn, value=guarded(rn, util_cn_ref), kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)

        f_insat = "+".join(cref(ENGINE_MKT_SHEET, mkt_out["insat"][(m, t)], rn) for m in MERCADOS for t in TECNOLOGIAS)
        S.apply_cell(ws, row_insatisfecha, 7 + rn, value=guarded(rn, f_insat), kind="link", numfmt=S.NUM_UNITS, align=S.ALIGN_RIGHT, size=9)

        # BUG REPORTADO Y CORREGIDO (pendiente desde la Adenda 12 de Integracion_Excel_Web_Control_
        # Gestion.md): la versión anterior promediaba los 12 casilleros (3 mercados x 4 tecnologías)
        # con el MISMO peso, incluyendo Eléctrico/Hidrógeno -- donde CADIZ pone 0 a propósito porque
        # no compite ahí. Sumar esos 6 ceros al promedio (y contarlos en el denominador, porque
        # ISNUMBER(0) es VERDADERO) diluía el resultado -- daba ~4,6% en Ronda 2 en vez de un número
        # representativo (confirmado contra Ronda 1 real: CESIM reporta 18,03% ponderado por
        # ventas). Corregido para ponderar por volumen, misma convención que usa CESIM en el dato
        # REAL que publica: la cuota de CADIZ de cada país (suma de sus 4 celdas de tecnología, ya
        # expresada como % del mercado TOTAL de ese país) se pondera por el tamaño de mercado de ese
        # país (Sección A1 del motor) -- el resultado es exactamente "volumen que CADIZ proyecta
        # vender en los 3 mercados / volumen total de los 3 mercados", no un promedio simple de
        # celdas sueltas.
        cuota_terms_todas = []
        for m in MERCADOS:
            for t in TECNOLOGIAS:
                try:
                    cuota_terms_todas.append(iref(inputs_idx, "D1c · DEMANDA (RESULTANTE)", "Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)", f"{m} / {t}", rn))
                except KeyError:
                    pass
        if cuota_terms_todas:
            # Guarda para que R0/R1 (sin ninguna celda numérica todavía, son "n/a") sigan en blanco
            # en vez de mostrar 0% -- mismo criterio que antes, solo que ahora el denominador del
            # cálculo principal ya no puede dar 0/0 por sí solo (ver más abajo).
            cuenta = "+".join(f'IF(ISNUMBER({x}),1,0)' for x in cuota_terms_todas)
            numer_partes, denom_partes = [], []
            for m in MERCADOS:
                celdas_m = []
                for t in TECNOLOGIAS:
                    try:
                        celdas_m.append(iref(inputs_idx, "D1c · DEMANDA (RESULTANTE)", "Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)", f"{m} / {t}", rn))
                    except KeyError:
                        pass
                if not celdas_m or m not in mkt_out["tam"]:
                    continue
                tam_ref = cref(ENGINE_MKT_SHEET, mkt_out["tam"][m], rn)
                suma_m = "+".join(f'IF(ISNUMBER({x}),{x},0)' for x in celdas_m)
                numer_partes.append(f'({suma_m})*{tam_ref}')
                denom_partes.append(tam_ref)
            if numer_partes:
                numer, denom = "+".join(numer_partes), "+".join(denom_partes)
                f_cuota = f'IF(({cuenta})=0,"",IFERROR(({numer})/({denom}),""))'
                S.apply_cell(ws, row_cuota_prom, 7 + rn, value=guarded(rn, f_cuota), kind="link", numfmt=S.NUM_PCT1, align=S.ALIGN_RIGHT, size=9)

        vef_terms = "+".join(cref(ENGINE_MKT_SHEET, mkt_out["vef"][(m, t)], rn) for m in MERCADOS for t in TECNOLOGIAS)
        if rn >= 2:
            S.apply_cell(ws, row_ing_veh, 7 + rn, value=guarded(rn, f'IFERROR({ing_ref}/({vef_terms}),"")'), kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
            S.apply_cell(ws, row_ebitda_veh, 7 + rn, value=guarded(rn, f'IFERROR({fin(rn,"ebitda")}/({vef_terms}),"")'), kind="calculo", numfmt=S.NUM_PRICE, align=S.ALIGN_RIGHT, size=9)
        else:
            S.apply_cell(ws, row_ing_veh, 7 + rn, value="n/a", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)
            S.apply_cell(ws, row_ebitda_veh, 7 + rn, value="n/a", kind="plain", align=S.ALIGN_CENTER, size=7, italic=True)

    # Hotfix visual: columnas C-F se ocultan más abajo (main()) -- freeze pasa a G.
    ws.freeze_panes = "G4"
    ratios_rows = dict(growth=row_growth, margen_ebitda=row_margen_ebitda, ros=row_ros, roe=row_roe, roce=row_roce,
                        wacc=row_wacc, roce_wacc=row_roce_wacc, eps=row_eps, fcf=row_fcf, deuda_fin=row_deuda_fin,
                        deuda_pat=row_deuda_pat, payout=row_payout, precio_proxy=row_precio_proxy, pe_proxy=row_pe_proxy,
                        divyield_proxy=row_divyield_proxy, util_us=row_util_us, util_cn=row_util_cn,
                        insatisfecha=row_insatisfecha, cuota_prom=row_cuota_prom, ing_veh=row_ing_veh,
                        ebitda_veh=row_ebitda_veh, retorno_acum=row_retorno_acum)
    return ws, ratios_rows


# ======================================================================================
# 04_CONTROL_MODELO — controles técnicos de salud del modelo (NO es Control de Gestión)
# ======================================================================================
def build_control_modelo(wb, inputs_idx, prod_out, mkt_out, glob, cond, decisiones, country_rows=None, rondas_reales=frozenset({0, 1})):
    ws = wb.create_sheet(CONTROL_SHEET)
    # Hotfix visual: idem 02/03 -- ncols/anchos deben cubrir A,B,C + D-F (espaciador oculto) + G-S
    # (rondas) = 3+3+13=19, no solo 16.
    ncols = 3 + 3 + len(RONDAS)
    S.set_col_widths(ws, [46, 10, 40] + [12] * 3 + [12] * len(RONDAS))
    S.style_title_row(ws, 1, "04_CONTROL_MODELO — Controles técnicos de salud del modelo (NO es Control de Gestión frente a profesores/accionistas)", ncols)
    S.apply_cell(ws, 2, 1, value="Verifica consistencia interna e integridad de datos. OK = check pasa; ERROR = revisar antes de usar la proyección. No evalúa desempeño estratégico (eso es 03_RATIOS).", kind="plain", align=S.ALIGN_LEFT_WRAP, size=8, italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    header_row = 4
    S.style_header_row(ws, header_row, ["Control", "Tipo", "Nota", "", "", ""] + [f"R{r}" for r in RONDAS])
    r = header_row + 1

    def sec(title):
        nonlocal r
        S.style_section_row(ws, r, title, ncols)
        r += 1

    def ctrl_row(label, nota=""):
        nonlocal r
        S.apply_cell(ws, r, 1, value=label, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=9)
        S.apply_cell(ws, r, 2, value="CONTROL", kind="calculo", align=S.ALIGN_CENTER, size=8, bold=True)
        S.apply_cell(ws, r, 3, value=nota, kind="calculo", align=S.ALIGN_LEFT_WRAP, size=7, italic=True)
        rr = r
        r += 1
        return rr

    sec("A · IDENTIDADES CONTABLES Y DE CONSOLIDACIÓN (Fase 2: por país Y Global)")
    row_bal_us = ctrl_row("Balance: Activo Total = Patrimonio + Pasivo — EE.UU.")
    row_bal_cn = ctrl_row("Balance: Activo Total = Patrimonio + Pasivo — China")
    row_bal_eu = ctrl_row("Balance: Activo Total = Patrimonio + Pasivo — Europa")
    row_bal = ctrl_row("Balance: Activo Total = Patrimonio + Pasivo (Global = Σ países)", "Fase 2: el Global ahora es la SUMA de los 3 Balances por país (arriba), 100% consistentes con RDOS en R0/R1 -- esto corrige la inconsistencia heredada de v1.2 (~1.508,8 MM USD en R1). Ver Informe_V2.md.")
    row_prest_ctrl = ctrl_row("Préstamos internos: Σ EE.UU.+China+Europa = 0", "REGLA VERIFICADA: los préstamos internos entre países (D8) deben netear a 0 exactamente cada ronda (Fase 2: ahora trackeados explícitamente por país, no solo colapsados a nivel Global).")
    row_cf = ctrl_row("Cash Flow: Efectivo cierre (CF) = Efectivo (Balance)")
    row_cons = ctrl_row("Consolidación: Beneficio Global = Σ Beneficio países")
    # -------------------------------------------------------------------------------------------
    # Cambio 1 (v1.2): el control de Balance ya NO usa una tolerancia amplia (~1.600 MM USD) que
    # disfrazaba de "OK" la diferencia heredada real (~-1.508.776.236,64 USD, inconsistencia
    # documentada entre 11_INVENTARIOS y 14_ESTADOS_FINANCIEROS de MOTOR_VALIDADO_R2 -- NO se
    # resuelve en este patch, ver nota de la fila y el Informe de Construcción punto 14). La
    # lógica de CÁLCULO de la diferencia (glob['bal_ctrl'] = Activo - (Patrimonio+Pasivo)) no se
    # toca; solo cambia el MENSAJE que se muestra, con 3 tramos genéricos (misma fórmula para
    # TODAS las rondas, no un valor fijo tipeado para R1/R2):
    #   - |diff| < TOL_REDONDEO (USD 1.000): "OK" -- diferencia de redondeo real, no la heredada.
    #   - |  |diff| - DIFERENCIA_HEREDADA_CONOCIDA  | < MARGEN_HEREDADA (±USD 1.000.000): mensaje de
    #     ADVERTENCIA explícito. Si la inconsistencia heredada se sigue arrastrando a R3+ cuando
    #     haya más rondas activas, esta misma fórmula la va a detectar automáticamente (no depende
    #     de qué ronda sea).
    #   - cualquier otra diferencia material: "ERROR" (deja de haber un techo de tolerancia laxo
    #     que esconda un descuadre real distinto al heredado).
    # -------------------------------------------------------------------------------------------
    TOL_REDONDEO_BALANCE = 1000  # USD -- redondeo real, no la inconsistencia heredada
    DIFERENCIA_HEREDADA_CONOCIDA = 1508776236.64  # USD -- ver nota de la fila / Informe de Construcción pto.14 (v1.2)
    MARGEN_HEREDADA = 1000000  # USD -- ±1MM de tolerancia alrededor de la cifra heredada conocida (v1.2; Fase 2 la corrige, esta rama queda como red de seguridad si algo la reintrodujera)

    def _bal_msg(diff_ref):
        return (
            f'=IF(ABS({diff_ref})<{TOL_REDONDEO_BALANCE},"OK",'
            f'IF(ABS(ABS({diff_ref})-{DIFERENCIA_HEREDADA_CONOCIDA})<{MARGEN_HEREDADA},'
            f'"ADVERTENCIA — diferencia heredada R1 ≈ USD 1.508,8 M (v1.2)","ERROR"))'
        )

    bal_rows_pais = {"EE.UU.": row_bal_us, "China": row_bal_cn, "Europa": row_bal_eu}
    for rn in RONDAS:
        c = col(rn)
        if country_rows:
            for pais, rr in bal_rows_pais.items():
                diff_ref_p = cref(ENGINE_FIN_SHEET, country_rows[pais]['bal_ctrl'], rn)
                S.apply_cell(ws, rr, 7 + rn, value=_bal_msg(diff_ref_p), kind="calculo", align=S.ALIGN_CENTER, size=8)
            prest_ref = cref(ENGINE_FIN_SHEET, glob['prest_ctrl'], rn)
            S.apply_cell(ws, row_prest_ctrl, 7 + rn, value=f'=IF(ABS({prest_ref})<1,"OK","ERROR")', kind="calculo", align=S.ALIGN_CENTER, size=8)
        diff_ref = cref(ENGINE_FIN_SHEET, glob['bal_ctrl'], rn)
        S.apply_cell(ws, row_bal, 7 + rn, value=_bal_msg(diff_ref), kind="calculo", align=S.ALIGN_CENTER, size=8)
        cfc = cref(ENGINE_FIN_SHEET, glob['cf_ctrl'], rn)
        S.apply_cell(ws, row_cf, 7 + rn, value=f'=IF(ISTEXT({cfc}),"OK",IF(ABS({cfc})<1,"OK","ERROR"))', kind="calculo", align=S.ALIGN_CENTER, size=8)
        cgc = cref(ENGINE_FIN_SHEET, glob['ctrl'], rn)
        S.apply_cell(ws, row_cons, 7 + rn, value=f'=IF(ISTEXT({cgc}),"OK",IF(ABS({cgc})<1,"OK","ERROR"))', kind="calculo", align=S.ALIGN_CENTER, size=8)
    r += 1

    sec("B · MERCADO, VENTAS E INVENTARIO")
    row_ventas_disp = ctrl_row("Ventas ≤ Disponibilidad (por área/tecnología)")
    row_inv_neg = ctrl_row("Inventario final ≥ 0 (todas las áreas/tecnología)")
    row_insat = ctrl_row("Demanda insatisfecha ≥ 0 (todos los mercados/tecnología)")
    row_prod_cap = ctrl_row("Producción propia ≤ Capacidad operativa (por área)")
    row_estado_d = ctrl_row("Asignación multi-origen: estado OK en Sección D de _ENGINE_MERCADO")
    for rn in RONDAS:
        c = col(rn)
        terms_vd = [f'({cref(ENGINE_PROD_SHEET, prod_out["ventas_u"][(a,t)], rn)}<={cref(ENGINE_PROD_SHEET, prod_out["disp_u"][(a,t)], rn)}+0.5)' for a in AREAS for t in TECNOLOGIAS]
        S.apply_cell(ws, row_ventas_disp, 7 + rn, value="=IF(AND(" + ",".join(terms_vd) + "),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        terms_inv = [f'({cref(ENGINE_PROD_SHEET, prod_out["inv_fin_u"][(a,t)], rn)}>=-0.5)' for a in AREAS for t in TECNOLOGIAS]
        S.apply_cell(ws, row_inv_neg, 7 + rn, value="=IF(AND(" + ",".join(terms_inv) + "),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        terms_ins = [f'({cref(ENGINE_MKT_SHEET, mkt_out["insat"][(m,t)], rn)}>=-0.5)' for m in MERCADOS for t in TECNOLOGIAS]
        S.apply_cell(ws, row_insat, 7 + rn, value="=IF(AND(" + ",".join(terms_ins) + "),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        terms_pc = [f'({cref(ENGINE_PROD_SHEET, prod_out["prodtot"][a], rn)}<={cref(ENGINE_PROD_SHEET, prod_out["cap"][a], rn)}+0.5)' for a in AREAS]
        S.apply_cell(ws, row_prod_cap, 7 + rn, value="=IF(AND(" + ",".join(terms_pc) + "),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        terms_ed = [f'({cref(ENGINE_MKT_SHEET, mkt_out["estado_d"][(m,t)], rn)}="OK")' for m in MERCADOS for t in TECNOLOGIAS] if isinstance(mkt_out.get("estado_d"), dict) else []
        if terms_ed:
            S.apply_cell(ws, row_estado_d, 7 + rn, value="=IF(AND(" + ",".join(terms_ed) + "),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        else:
            S.apply_cell(ws, row_estado_d, 7 + rn, value="n/d", kind="plain", align=S.ALIGN_CENTER, size=8)
    r += 1

    sec("C · TERCERIZACIÓN, DECISIONES Y CONDICIONES (D2/D8/D9)")
    row_terc_costo = ctrl_row("BLOQUEANTE: si Tercerización > 0, Costo unitario de tercerización debe estar cargado", "Manual cap.5: el costo de tercerización varía por volumen y no se publica de antemano -- ver 01_INPUTS 'B. Condiciones'.")
    row_d8_consist = ctrl_row("D8: Variación de deuda LP + Emisión − Recompra no deja Patrimonio/Deuda LP negativos")
    row_inputs_vacios = ctrl_row("Sin inputs de decisión (D1-D8) vacíos en la ronda activa")
    row_maxtech_region = ctrl_row("Máx. tecnologías producidas por área productora", "NO VERIFICADO EN MANUAL -- no se encontró un límite explícito en el texto disponible del manual CESIM. No se fuerza ningún tope; ver Informe de Construcción punto 14.")
    row_maxtech_mercado = ctrl_row("Máx. tecnologías comercializadas por mercado", "NO VERIFICADO EN MANUAL -- no se encontró un límite explícito en el texto disponible del manual CESIM. No se fuerza ningún tope; ver Informe de Construcción punto 14.")
    for rn in RONDAS:
        c = col(rn)
        terms_bt = []
        for a in AREAS:
            for t in TECNOLOGIAS:
                terc_ref = iref(inputs_idx, "D2 · PRODUCCIÓN", "Producción tercerizada", f"{a} / {t}", rn)
                costo_ref = iref(inputs_idx, "B. Condiciones", "Costo unitario de producción tercerizada", f"{a} / {t}", rn)
                terms_bt.append(f'(NOT(AND(ISNUMBER({terc_ref}),{terc_ref}>0,NOT(ISNUMBER({costo_ref})))))')
                terms_bt.append(f'(NOT(AND(ISNUMBER({terc_ref}),{terc_ref}>0,ISNUMBER({costo_ref}),{costo_ref}<=0)))')
        S.apply_cell(ws, row_terc_costo, 7 + rn, value="=IF(AND(" + ",".join(terms_bt) + "),\"OK\",\"ERROR — FALTA COSTO DE TERCERIZACIÓN — PROYECCIÓN NO VÁLIDA\")", kind="alerta", align=S.ALIGN_CENTER, size=8, bold=True)
        if rn in rondas_reales:
            S.apply_cell(ws, row_d8_consist, 7 + rn, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
        else:
            S.apply_cell(ws, row_d8_consist, 7 + rn, value=f"=IF(AND({cref(ENGINE_FIN_SHEET, glob['pat'], rn)}>=0,{cref(ENGINE_FIN_SHEET, glob['dlp'], rn)}>=0),\"OK\",\"ERROR\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
        S.apply_cell(ws, row_maxtech_region, 7 + rn, value="NO VERIFICADO", kind="alerta", align=S.ALIGN_CENTER, size=8, italic=True)
        S.apply_cell(ws, row_maxtech_mercado, 7 + rn, value="NO VERIFICADO", kind="alerta", align=S.ALIGN_CENTER, size=8, italic=True)
    for rn in range(2, 13):
        c = col(rn)
        # Solo se exige carga para TECH_ACTIVAS (Combustión/Híbrido): Eléctrico/Hidrógeno aún no
        # comercializados por CADIZ, sus celdas quedan en blanco legítimamente (no es un error).
        d1_terms = ([iref(inputs_idx, "D1a · MIX INDUSTRIA", "Mix de Industria (TAM)", f"{m} / {t}", rn) for m in MERCADOS for t in TECH_ACTIVAS] +
                    [iref(inputs_idx, "D1 · DEMANDA", "Cuota de mercado objetivo SOBRE LA TECNOLOGÍA", f"{m} / {t}", rn) for m in MERCADOS for t in TECH_ACTIVAS])
        d5_terms = [iref(inputs_idx, "D5 · MARKETING", "Precio de venta", f"{m} / {t}", rn) for m in MERCADOS for t in TECH_ACTIVAS]
        chk = "+".join(f'IF(ISNUMBER({x}),0,1)' for x in d1_terms + d5_terms)
        S.apply_cell(ws, row_inputs_vacios, 7 + rn, value=f"=IF(({chk})=0,\"OK\",\"REVISAR — hay celdas de decisión sin cargar\")", kind="calculo", align=S.ALIGN_CENTER, size=8)
    S.apply_cell(ws, row_inputs_vacios, 7 + 0, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_inputs_vacios, 7 + 1, value="n/a (ronda ya jugada)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    r += 1

    sec("D · INTEGRIDAD DEL MODELO (fórmulas, histórico)")
    row_hist_r0 = ctrl_row("Integridad histórica: Ronda 0 no fue alterada (Ingresos Global R0 = 39.489.000.000, dato migrado)")
    row_hist_r1 = ctrl_row("Integridad histórica: Ronda 1 no fue alterada (Beneficio Global R1 = 3.609.336.474,27, dato migrado)")
    S.apply_cell(ws, row_hist_r0, 7, value="OK (verificar manualmente contra RDOS si se sospecha edición)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    S.apply_cell(ws, row_hist_r1, 7, value="OK (verificar manualmente contra RDOS si se sospecha edición)", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
    r += 1

    S.apply_cell(ws, r, 1, value="NOTA: 'Costo de gestión de inventario' (manual cap.5.2) está declarado en 01_INPUTS pero en blanco -- NO se encontró la tasa/monto publicado en los archivos del proyecto. No alimenta ninguna fórmula de este workbook (ni EBITDA ni Balance). Ver Informe de Construcción punto 14 (pendiente, no es una decisión estratégica de CADIZ).", kind="alerta", align=S.ALIGN_LEFT_WRAP, size=8, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    ws.row_dimensions[r].height = 30

    # -------------------------------------------------------------------------------------------
    # Cambio 2 (v1.1): para toda ronda NO activa (R3-R12 sin Estado+inputs mínimos cargados), los
    # controles de esta hoja deben mostrar literalmente "NO ACTIVA" en vez de evaluar OK/ERROR con
    # inputs en blanco. Se aplica de forma genérica sobre TODAS las celdas-fórmula ya escritas en
    # las columnas R2-R12 de los controles tipo "CONTROL" (columnas R0/R1 y R2 no se tocan: R2
    # queda con el mismo resultado que en v1 porque RONDA_ACTIVA(R2)=VERDADERO). Los controles que
    # son texto fijo informativo (p.ej. "NO VERIFICADO") no son fórmulas y no se tocan.
    # -------------------------------------------------------------------------------------------
    def _activo_ref(rn):
        return iref(inputs_idx, "A. Identificación", "RONDA_ACTIVA (calculado)", "-", rn)

    ctrl_rows_guardadas = [row_bal_us, row_bal_cn, row_bal_eu, row_bal, row_prest_ctrl, row_cf, row_cons,
                            row_ventas_disp, row_inv_neg, row_insat,
                            row_prod_cap, row_estado_d, row_terc_costo, row_d8_consist,
                            row_inputs_vacios]
    for rr in ctrl_rows_guardadas:
        for rn in range(2, 13):
            cell = ws.cell(row=rr, column=7 + rn)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                inner = v[1:]
                cell.value = f'=IF({_activo_ref(rn)},{inner},"NO ACTIVA")'

    # Hotfix visual: B/C (Tipo/Nota) y D-F (relleno vacío) se ocultan en main() -- freeze pasa a G,
    # igual que 01_INPUTS/02_ESTADOS_PROYECTADOS/03_RATIOS.
    ws.freeze_panes = "G5"
    return ws


# ======================================================================================
# 05_HISTORICO_EQUIPOS — histórico tidy (Ronda 0/1, los 7 equipos) migrado de
# MOTOR_VALIDADO_R2.xlsx / 24_HISTORICO_RONDAS + append de CADIZ Escenario A Ronda 2 (PLAN)
# ======================================================================================
EQUIPOS = ["CADIZ", "CEOS", "CHIEF", "CLAVE", "CUORE", "FOCUS", "TOKIO"]


def _parse_seccion(sec):
    parts = [p.strip() for p in sec.split(",")]
    region = None
    tipo_parts = []
    for p in parts:
        if p in ("China", "EE.UU.", "Europa", "Global"):
            region = p
        elif p in ("miles USD", "%"):
            continue
        else:
            tipo_parts.append(p)
    tipo = ", ".join(tipo_parts)
    for pref in ("China ", "EE.UU. ", "Europa "):
        if tipo.startswith(pref):
            region = pref.strip()
            tipo = tipo[len(pref):]
    if region is None and "global" in tipo.lower():
        region = "Global"
    if region is None and "casa matriz" in tipo.lower():
        # VERIFICADO por dato: "Efectivo 31.12." de 'Flujo de efectivo de casa matriz' (R0, CADIZ)
        # = 9.274.082,79, IDÉNTICO a 'Hoja de Balance, miles USD, EE.UU.' — Efectivo y equivalentes
        # (mismo valor), NO al Global consolidado (10.817.319,05 = suma EE.UU.+China+Europa). La
        # "casa matriz" en 24_HISTORICO_RONDAS es la entidad EE.UU. (solo esa filial), no el
        # consolidado de 3 países -- región EE.UU. es correcta, no un supuesto.
        region = "EE.UU."
    if region is None and tipo.strip() == "Indicadores Financieros Claves":
        # VERIFICADO por dato: "Rentabilidad de las ventas (ROS)" (R0, CADIZ) = 4.410388%% =
        # Beneficio de la ronda Global / Ingresos por ventas Global (1.741.617,98 / 39.489.000),
        # y "Beneficio operativo antes de depreciación ( EBITDA )" = 10.512170% = EBITDA Global /
        # Ingresos Global (4.151.150,78 / 39.489.000). Sección sin país explícito = GLOBAL.
        region = "Global"
    return tipo.strip(), (region or "")


# --------------------------------------------------------------------------------------
# Mapeo canónico de nombres de 'metric' -- SOLO para conceptos de 24_HISTORICO_RONDAS (REAL)
# que tienen un equivalente EXACTO y VERIFICADO en 02_ESTADOS_PROYECTADOS/03_RATIOS (PLAN),
# ambos a nivel GLOBAL (consolidado 3 países) -- la única granularidad que existe en el PLAN.
# Ver README "Convención de nombres de metric" para la tabla completa y la verificación
# valor-a-valor. Fuera de esta lista, un rubro de 24_HISTORICO_RONDAS que NO tiene equivalente
# claro en el modelo nuevo conserva su nombre original "{Sección} — {Rubro}" (ya unívoco porque
# la Sección actúa de prefijo distintivo) -- NO se inventa una equivalencia que no existe.
CANON_GLOBAL_MAP = {
    ("Cuenta de resultados", "Ingresos por ventas"): "Ingresos por ventas",
    ("Cuenta de resultados", "Costos y gastos totales"): "Costos y gastos totales",
    ("Cuenta de resultados", "Beneficio operativo antes de depreciación (EBITDA)"): "EBITDA",
    ("Cuenta de resultados", "Depreciación de Activos Fijos"): "Depreciación",
    ("Cuenta de resultados", "Beneficio operativo (EBIT)"): "EBIT",
    ("Cuenta de resultados", "Gastos financieros netos"): "Gastos financieros netos",
    ("Cuenta de resultados", "Beneficio antes de impuestos"): "Beneficio antes de impuestos",
    ("Cuenta de resultados", "Impuesto sobre el beneficio"): "Impuesto sobre el beneficio",
    ("Cuenta de resultados", "Beneficio de la ronda"): "Beneficio de la ronda (neto)",
    ("Hoja de Balance", "Activo fijo"): "Activo fijo (neto)",
    ("Hoja de Balance", "Inventario"): "Inventario",
    ("Hoja de Balance", "Cuentas por Cobrar"): "Cuentas por Cobrar",
    ("Hoja de Balance", "Efectivo y equivalentes de efectivo"): "Efectivo y equivalentes",
    ("Hoja de Balance", "Activos Totales"): "Activos Totales",
    ("Hoja de Balance", "Deudas a largo plazo"): "Deudas a largo plazo",
    ("Hoja de Balance", "Deudas a corto plazo (no planificadas)"): "Deudas a corto plazo (automática)",
    ("Hoja de Balance", "Cuentas por pagar"): "Cuentas por Pagar",
    ("Hoja de Balance", "Pasivos Totales"): "Pasivos Totales",
    ("Hoja de Balance", "Total patrimonio neto"): "Patrimonio neto",
    # "Beneficio de la ronda" en Hoja de Balance es el MISMO importe que en Cuenta de resultados
    # (verificado: idéntico valor en R0 para los 7 equipos) -- mismo canónico; el dedup de más
    # abajo evita escribir la fila duplicada dos veces con el mismo valor.
    ("Hoja de Balance", "Beneficio de la ronda"): "Beneficio de la ronda (neto)",
    ("Indicadores Financieros Claves", "Beneficio operativo antes de depreciación ( EBITDA )"): "Margen EBITDA",
    ("Indicadores Financieros Claves", "Rentabilidad de las ventas (ROS)"): "ROS",
    ("Indicadores Financieros Claves", "Rendimiento de los Fondos Propios (ROE)"): "ROE",
    ("Indicadores Financieros Claves", "Rentabilidad del capital empleado (ROCE)"): "ROCE",
    ("Indicadores Financieros Claves", "Ganancias por acción (EPS), USD"): "EPS",
}
# Estados de flujo de efectivo (casa matriz=EE.UU., y regionales China/Europa): el rubro fuente
# repite el texto "Total" DOS VECES por sección/ronda (subtotal de operación y subtotal de
# financiación) -- verificado con datos R0 (fila 76 vs 85 de 24_HISTORICO_RONDAS): son valores
# DISTINTOS bajo el mismo (Sección, Rubro), lo que ya rompía el contrato de identificación por
# atributos ANTES de este parche. Se renombran a los mismos canónicos que usa el PLAN.
CF_TIPOS = {"Flujo de efectivo de casa matriz", "Estado de flujo de efectivo"}
CF_CFI_RUBRO = "Inversiones en fábricas (-) / desinversiones (+)"


def _values_equal(a, b):
    if a is None or b is None:
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < max(1.0, abs(a) * 1e-9)
    return a == b


def build_historico_equipos(wb, rdos_files, glob, ratios_rows, prod_out, mkt_out, inputs_idx):
    ws = wb.create_sheet(HIST_SHEET)
    S.set_col_widths(ws, [8, 10, 10, 12, 12, 46, 20, 12])
    S.style_title_row(ws, 1, "05_HISTORICO_EQUIPOS — Histórico tidy (7 equipos, parseado directo de los RDOS de cada ronda jugada -- rdos_parser.py) + PLAN/SIM CADIZ Ronda 2-12 (rondas activas)", 8)
    S.apply_cell(ws, 2, 1, value="DATO HISTÓRICO REAL para Ronda 0/1 (Estado=REAL, todos los equipos); Tecnología=NA salvo en 'Cuota de mercado' (dato real desagregado por tecnología, ver README) -- el resto de secciones del histórico mezcla varias tecnologías/sub-tablas sin columna propia y NO se reconstruye por inferencia (ver README 'Limitaciones conocidas'); filas con misma clave (ronda,equipo,región,métrica) pero valor distinto quedan marcadas '(variante N)' en vez de colisionar en silencio. Región=NA cuando la fila fuente no identifica un país/Global explícito (no se inventa). Filas PLAN/SIM (Equipo=CADIZ, Ronda 2-12) se generan SOLO para rondas con RONDA_ACTIVA=VERDADERO (Cambio 2, v1.2) y están LINKEADAS por fórmula al motor de este workbook y a 01_INPUTS (Estado); metric/region/technology siguen la misma convención canónica que REAL (ver README).", kind="plain", align=S.ALIGN_LEFT_WRAP, size=8, italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws.row_dimensions[2].height = 28
    header_row = 4
    S.style_header_row(ws, header_row, ["Ronda", "Equipo", "Estado", "Región", "Tecnología", "Métrica", "Valor", "Unidad"])
    r = header_row + 1

    # Fase 3 (corrección de escala): se lee directamente cada RDOS de ronda jugada (rdos_files =
    # {ronda_int: path_o_fileobj}) vía el parser único rdos_parser.parse_rdos_workbook(), en vez de
    # la copia pre-parseada 24_HISTORICO_RONDAS de MOTOR_VALIDADO_R2.xlsx (retirado como fuente).
    # Se preserva SIN CAMBIOS toda la lógica de canonicalización de nombres/CFO-CFF/dedup/variantes
    # que ya existía (_parse_seccion, CANON_GLOBAL_MAP, CF_TIPOS, CF_CFI_RUBRO, _values_equal) --
    # solo cambia la FUENTE de (Sección, Rubro, Unidad, valores por equipo), que ahora es el propio
    # RDOS (dato histórico real primario) en vez de una copia intermedia. El valor escrito es el
    # CRUDO de cada fila (row["values"], no row["scaled"]) + su unidad resuelta en texto (row["unit"],
    # ej. "miles USD"/"miles unidades"/"%"/"sin_unidad") -- igual que antes, la conversión x1000 de
    # "miles USD" se aplica más abajo en build_data_export()/_canonicalize_real(), no acá.
    n_hist = 0
    cf_total_count = {}      # (ronda,tipo,region) -> nº de 'Total' vistos (CFO=1º, CFF=2º)
    real_seen = {}           # (ronda,equipo,region,metrica_base) -> [valores ya escritos]
    n_dedup_skipped = 0
    n_variant_flagged = 0
    n_tech_cuota = 0
    for ronda in sorted(rdos_files):
        rows = parse_rdos_workbook(rdos_files[ronda])
        for row in rows:
            if row["type"] != "data":
                continue
            seccion, rubro, unidad = row["section"], row["label"], row["unit"]
            if seccion is None:
                continue
            tipo, region = _parse_seccion(seccion)
            tecnologia = "NA"

            canon = CANON_GLOBAL_MAP.get((tipo, rubro)) if region == "Global" else None
            tipo_low = tipo.lower()
            if canon:
                metrica = canon
            elif tipo in CF_TIPOS and rubro == "Total":
                key = (ronda, tipo, region)
                cf_total_count[key] = cf_total_count.get(key, 0) + 1
                metrica = ("Total actividades operativas (CFO)" if cf_total_count[key] == 1
                           else "Total actividades financieras (CFF)")
            elif tipo in CF_TIPOS and rubro == CF_CFI_RUBRO:
                metrica = "Total actividades de inversión (CFI)"
            elif rubro in TECNOLOGIAS and "cuota" in tipo_low and "mercado" in tipo_low:
                # "Cuotas de mercado globales, %" / "{Región} cuotas de mercado, %": el rubro fuente
                # YA ES la tecnología (Combustión/Híbrido/Eléctrico/Hidrógeno) -- dato real desagregado
                # por tecnología que existía pero no se exportaba en la columna Tecnología. Verificado
                # limpio: sin filas duplicadas (round,región,rubro) en esta sección (a diferencia de
                # otras secciones del histórico, ver README "Limitaciones conocidas").
                metrica = "Cuota de mercado"
                tecnologia = rubro
                n_tech_cuota += 1
            else:
                metrica = f"{tipo} — {rubro}" if rubro else tipo

            for equipo in EQUIPOS:
                val = row["values"].get(equipo)
                if val is None or val == "":
                    continue
                basekey = (ronda, equipo, (region or "NA"), tecnologia, metrica)
                vals_list = real_seen.setdefault(basekey, [])
                dup_idx = next((i for i, v in enumerate(vals_list) if _values_equal(v, val)), None)
                if dup_idx is not None:
                    n_dedup_skipped += 1
                    continue
                vals_list.append(val)
                variant_n = len(vals_list)
                if variant_n == 1:
                    metrica_row = metrica
                else:
                    # Misma clave (ronda,equipo,región,métrica) pero valor DISTINTO: colisión real
                    # preexistente en el propio RDOS (secciones que mezclan varias tecnologías/
                    # sub-tablas sin columna propia -- ver README). Se marca como variante en vez de
                    # sobrescribir en silencio o descartar el dato.
                    metrica_row = f"{metrica} (variante {variant_n})"
                    n_variant_flagged += 1
                S.apply_cell(ws, r, 1, value=ronda, kind="historico", align=S.ALIGN_CENTER, size=8)
                S.apply_cell(ws, r, 2, value=equipo, kind="historico", align=S.ALIGN_CENTER, size=8)
                S.apply_cell(ws, r, 3, value="REAL", kind="historico", align=S.ALIGN_CENTER, size=8)
                S.apply_cell(ws, r, 4, value=(region or "NA"), kind="historico", align=S.ALIGN_CENTER, size=8)
                S.apply_cell(ws, r, 5, value=tecnologia, kind="historico", align=S.ALIGN_CENTER, size=8)
                S.apply_cell(ws, r, 6, value=metrica_row, kind="historico", align=S.ALIGN_LEFT, size=8)
                S.apply_cell(ws, r, 7, value=val, kind="historico", numfmt='#,##0.####', align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, r, 8, value=(unidad or ""), kind="historico", align=S.ALIGN_CENTER, size=8)
                r += 1
                n_hist += 1
    print(f"  [05_HISTORICO_EQUIPOS] dedup (mismo valor, fila no repetida): {n_dedup_skipped} | "
          f"variantes marcadas (misma clave, valor distinto -- colisión preexistente en el RDOS): {n_variant_flagged} | "
          f"filas de Cuota de mercado con Tecnología poblada: {n_tech_cuota}")

    # -------------------------------------------------------------------------------------------
    # Cambio 2 (v1.2): generalización de la sección PLAN/SIM CADIZ de R2 (único caso hardcodeado
    # en v1.1) a un LOOP genérico R2-R12, guardado por el mismo indicador RONDA_ACTIVA que ya usan
    # 02_ESTADOS_PROYECTADOS/03_RATIOS/04_CONTROL_MODELO desde v1.1. Para una ronda rn con
    # RONDA_ACTIVA(rn)=FALSO no se escribe NINGUNA fila (ni vacía ni con placeholder): el loop
    # directamente hace `continue`. R0/R1 (histórico REAL) no pasan por este bloque -- ya fueron
    # migradas más arriba, sin cambios.
    #
    # RONDA_ACTIVA no puede leerse como fórmula ya evaluada en este punto (openpyxl no calcula
    # fórmulas -- recién se recalculan con recalc.py, al final): _ronda_activa_buildtime() replica
    # en Python la MISMA condición que la fórmula de 01_INPUTS!'RONDA_ACTIVA (calculado)'
    # (Estado ∈ {PLAN,SIM,REAL} Y Cuota de mercado objetivo SOBRE LA TECNOLOGÍA EE.UU./Combustión cargada),
    # leyendo los valores literales que build_inputs() ya escribió en esas celdas. No es una
    # fórmula nueva ni un criterio distinto -- es el mismo criterio, evaluado en Python porque acá
    # hace falta decidir en tiempo de construcción si la fila se escribe o no.
    # -------------------------------------------------------------------------------------------
    estado_row = inputs_idx["A. Identificación||Estado||-"]
    escenario_row = inputs_idx["A. Identificación||Escenario||-"]
    cuota_check_row = inputs_idx["D1 · DEMANDA||Cuota de mercado objetivo SOBRE LA TECNOLOGÍA||EE.UU. / Combustión"]

    def _ronda_activa_buildtime(rn_chk):
        estado_val = ws_in.cell(row=estado_row, column=7 + rn_chk).value
        cuota_val = ws_in.cell(row=cuota_check_row, column=7 + rn_chk).value
        return (estado_val in ("PLAN", "SIM", "REAL")
                and isinstance(cuota_val, (int, float)) and not isinstance(cuota_val, bool))

    ws_in = wb[INPUTS_SHEET]
    plan_start = r  # si ninguna ronda R2-R12 está activa, no se escribe nada y r no avanza -> 0 filas

    for rn in range(2, 13):
        if rn in rdos_files:
            # Ronda N.N (Fase 4, frontera dinámica): esta ronda YA fue migrada como histórico REAL
            # en el loop de arriba (for ronda in sorted(rdos_files)) -- antes esto era imposible para
            # rn>=2 (rondas_reales fijo a {0,1}), pero ahora que Estado(rn)="REAL" generaliza a
            # cualquier ronda con RDOS, sin este guard _ronda_activa_buildtime(rn) daría VERDADERO
            # también acá (Estado="REAL" ∈ {PLAN,SIM,REAL}) y se escribiría una fila DUPLICADA
            # (misma clave ronda/CADIZ/Global/métrica) con fórmulas PLAN en vez del valor REAL ya
            # migrado. Se salta explícitamente.
            continue
        if not _ronda_activa_buildtime(rn):
            continue
        estado_ref = f"'{INPUTS_SHEET}'!{col(rn)}{estado_row}"
        estado_val = ws_in.cell(row=estado_row, column=7 + rn).value
        esc_val = ws_in.cell(row=escenario_row, column=7 + rn).value
        esc_label = esc_val if esc_val not in (None, "") else "(sin escenario)"
        r += 1
        S.style_section_row(ws, r, f"CADIZ — ESCENARIO {esc_label}, RONDA {rn} ({estado_val} — linkeado por fórmula a _ENGINE_FINANCIERO / _ENGINE_PRODUCCION / _ENGINE_MERCADO / 03_RATIOS / 01_INPUTS de este workbook)", 8)
        r += 1

        def plan_row(metrica, ref, unidad, region="Global", tech="NA"):
            nonlocal r
            S.apply_cell(ws, r, 1, value=rn, kind="plain", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, r, 2, value="CADIZ", kind="plain", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, r, 3, value=f"={estado_ref}", kind="link", align=S.ALIGN_CENTER, size=8, bold=True)
            S.apply_cell(ws, r, 4, value=region, kind="plain", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, r, 5, value=tech, kind="plain", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, r, 6, value=metrica, kind="plain", align=S.ALIGN_LEFT, size=8)
            S.apply_cell(ws, r, 7, value=f"={ref}", kind="link", numfmt='#,##0.####', align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, r, 8, value=unidad, kind="plain", align=S.ALIGN_CENTER, size=8)
            r += 1

        for label, key, unidad in [
            ("Ingresos por ventas", "ing", "USD"), ("Costos y gastos totales", "costos", "USD"),
            ("EBITDA", "ebitda", "USD"), ("Depreciación", "dep", "USD"), ("EBIT", "ebit", "USD"),
            ("Gastos financieros netos", "gfn", "USD"), ("Beneficio antes de impuestos", "ebt", "USD"),
            ("Impuesto sobre el beneficio", "imp", "USD"), ("Beneficio de la ronda (neto)", "ben", "USD"),
            ("Activo fijo (neto)", "af", "USD"), ("Inventario", "inv", "USD"), ("Cuentas por Cobrar", "cxc", "USD"),
            ("Efectivo y equivalentes", "caja", "USD"), ("Activos Totales", "activos", "USD"),
            ("Deudas a largo plazo", "dlp", "USD"), ("Deudas a corto plazo (automática)", "dcp", "USD"),
            ("Cuentas por Pagar", "cxp", "USD"), ("Pasivos Totales", "pasivos", "USD"),
            ("Patrimonio neto", "pat", "USD"), ("Total actividades operativas (CFO)", "cfo", "USD"),
            ("Total actividades de inversión (CFI)", "cfi", "USD"), ("Total actividades financieras (CFF)", "cff", "USD"),
        ]:
            plan_row(label, cref(ENGINE_FIN_SHEET, glob[key], rn), unidad)

        # NOTA UNIDADES: 03_RATIOS guarda los % como fracción (0.105 = 10,5%, formato de celda '0.0%').
        # 24_HISTORICO_RONDAS (REAL) guarda TODOS sus valores en "%" ya multiplicados x100 (verificado:
        # p.ej. 'Cuotas de mercado globales, %' — Combustión R0 = 14.2857142857143, no 0.142857...; ROE
        # de 'Indicadores Financieros Claves' R0 = 19.07, no 0.1907). Para que un mismo metric="ROE" (u
        # otro ratio en %) sea numéricamente comparable entre status=REAL y status=PLAN, TODO ratio "%"
        # o "p.p." que va a 05_HISTORICO_EQUIPOS/DATA_EXPORT se multiplica x100 aquí (silo NUEVO, no
        # afecta el formato/valor de pantalla de 03_RATIOS).
        for label, key, unidad, mult in [
            ("Margen EBITDA", "margen_ebitda", "%", 100), ("ROS", "ros", "%", 100),
            ("ROE", "roe", "%", 100), ("ROCE", "roce", "%", 100),
            ("ROCE − WACC", "roce_wacc", "p.p.", 100), ("EPS", "eps", "USD/acción", 1), ("FCF", "fcf", "USD", 1),
            ("Payout ratio", "payout", "%", 100),
            ("Demanda insatisfecha CADIZ (total)", "insatisfecha", "u.", 1),
            ("Cuota de mercado CADIZ (ponderada por volumen)", "cuota_prom", "%", 100),
            ("Retorno total acumulado del accionista (PROXY)", "retorno_acum", "%", 100),
        ]:
            ref = cref(RATIOS_SHEET, ratios_rows[key], rn)
            # Bug real, corregido: 03_RATIOS devuelve IFERROR(...,"") cuando el ratio no es calculable
            # (p.ej. Ingresos=0 en una ronda con decisiones incompletas) -- multiplicar ese "" x100
            # directamente (fórmula ="..."&""*100) da #VALUE! en vez de propagar el blank. Se envuelve
            # en IFERROR acá también para que el blank se propague como blank, no como error.
            f_ref = f"IFERROR({ref}*{mult},\"\")" if mult != 1 else f"IFERROR({ref},\"\")"
            plan_row(label, f_ref, unidad)
        # Región va en la columna Región, NO embebida en el texto de la métrica (contrato DATA_EXPORT).
        plan_row("Utilización de capacidad", f"{cref(RATIOS_SHEET, ratios_rows['util_us'], rn)}*100", "%", region="EE.UU.")
        plan_row("Utilización de capacidad", f"{cref(RATIOS_SHEET, ratios_rows['util_cn'], rn)}*100", "%", region="China")
        # Capacidad operativa (cierre de ronda), unidades absolutas, por área (Adenda 11 hotfix): se
        # exporta para que la web pueda reconstruir la "Utilización de capacidad" REAL (el RDOS no
        # publica un único % por área -- lo publica desglosado por tecnología, ver metric_crosswalk.py)
        # dividiendo Producción interna real (real, sumada por tecnología) / esta capacidad, en vez de
        # asumir una constante fija a mano en la web. Si CADIZ invierte/desinvierte capacidad (D2) en
        # una ronda futura, este número ya viene actualizado desde el motor -- no hay que tocar la web.
        for area in AREAS:
            plan_row("Capacidad operativa", cref(ENGINE_PROD_SHEET, prod_out["cap"][area], rn), "u.", region=area)

        # Precio / Promoción / Ventas efectivas por (mercado, tecnología) y Costo unitario por
        # (área, tecnología) -- Adenda 12 hotfix: se exportan para la "Comparativa Plan vs. Real"
        # (waterfall de Ingresos por Precio/Volumen/Mix en Resultados, GAP de fabricación en
        # Operaciones). Ya se calculaban internamente (D5·MARKETING es una DECISIÓN CADIZ, _ENGINE_
        # PRODUCCION/_ENGINE_MERCADO ya las computan) -- esto solo las manda a DATA_EXPORT, no cambia
        # ninguna fórmula existente. Moneda: Precio de venta queda en la moneda nativa del mercado
        # (MONEDA_MERCADO) igual que el RDOS real (que tampoco convierte) -- comparable sin FX.
        for mercado in MERCADOS:
            for tech in TECNOLOGIAS:
                dim = f"{mercado} / {tech}"
                precio_ref = iref(inputs_idx, "D5 · MARKETING", "Precio de venta", dim, rn)
                plan_row("Precio de venta", precio_ref, MONEDA_MERCADO[mercado], region=mercado, tech=tech)
                cuota_obj_ref = iref(inputs_idx, "D1c · DEMANDA (RESULTANTE)", "Cuota Resultante SOBRE EL TOTAL (Copiar a CESIM)", dim, rn)
                # Se exporta como fracción 0-1 (mismo valor nativo del input, ej. 0,137 = 13,7%) con
                # unidad "ratio" -- NO se multiplica x100 (a diferencia de los ratios calculados del
                # motor migrados vía 05_HISTORICO_EQUIPOS/_RATIO_METRICS_PLAN) porque este valor sale
                # directo de 01_INPUTS, no de ese camino de canonicalización.
                plan_row("Cuota de mercado objetivo CADIZ", cuota_obj_ref, "ratio", region=mercado, tech=tech)
                promo_ref = iref(inputs_idx, "D5 · MARKETING", "Presupuesto de promoción", dim, rn)
                plan_row("Presupuesto de promoción", promo_ref, "USD", region=mercado, tech=tech)
                plan_row("Ventas efectivas (mercado)", cref(ENGINE_MKT_SHEET, mkt_out["vef"][(mercado, tech)], rn),
                         "u.", region=mercado, tech=tech)
        for area in AREAS:
            for tech in TECNOLOGIAS:
                plan_row("Costo unitario de producción propia", cref(ENGINE_PROD_SHEET, prod_out["costo_propio"][(area, tech)], rn),
                         "USD/u.", region=area, tech=tech)
                plan_row("Costo unitario de producción tercerizada", cref(ENGINE_PROD_SHEET, prod_out["costo_terc"][(area, tech)], rn),
                         "USD/u.", region=area, tech=tech)

        # Ventas totales y Demanda estimada: fila agregada (todas las tecnologías, Tecnología=NA) +
        # filas desagregadas por tecnología (dato real del motor, no inventado -- _ENGINE_PRODUCCION/
        # _ENGINE_MERCADO ya calculan cada (área,tecnología)/(mercado,tecnología) por separado).
        for area in AREAS:
            plan_row("Ventas totales (unidades)", "+".join(cref(ENGINE_PROD_SHEET, prod_out["ventas_u"][(area, t)], rn) for t in TECNOLOGIAS), "u.", region=area)
            for t in TECNOLOGIAS:
                plan_row("Ventas totales (unidades)", cref(ENGINE_PROD_SHEET, prod_out["ventas_u"][(area, t)], rn), "u.", region=area, tech=t)
        for m in MERCADOS:
            plan_row("Demanda estimada CADIZ (unidades)", "+".join(cref(ENGINE_MKT_SHEET, mkt_out["demanda"][(m, t)], rn) for t in TECNOLOGIAS), "u.", region=m)
            for t in TECNOLOGIAS:
                plan_row("Demanda estimada CADIZ (unidades)", cref(ENGINE_MKT_SHEET, mkt_out["demanda"][(m, t)], rn), "u.", region=m, tech=t)

    ws.auto_filter.ref = f"A{header_row}:H{max(r - 1, header_row)}"
    ws.freeze_panes = f"A{header_row+1}"
    return ws, n_hist, (r - plan_start), plan_start


# ======================================================================================
# DATA_EXPORT (oculta) — contrato tabular para una futura app Python. Espeja 05_HISTORICO_EQUIPOS
# fila a fila por fórmula (histórico y PLAN CADIZ ya vienen correctamente sembrados/linkeados allí).
# ======================================================================================
# ---------------------------------------------------------------------------------------------
# Cambio 3 (v1.1) -- CONVENCIÓN CANÓNICA DE UNIDADES en DATA_EXPORT (ver README, tabla completa).
# La transformación se decide EXCLUSIVAMENTE por la unidad de origen ya escrita en 05_HISTORICO_
# EQUIPOS (columna Unidad) -- nunca se toca 05_HISTORICO_EQUIPOS en sí. Regla:
#   "miles USD" -> "USD"        (valor x1000)
#   "USD"       -> "USD"        (sin cambio, ya es absoluto -- caso de filas PLAN)
#   "%"         -> "ratio"      (valor /100, decimal puro: 33,39% -> 0.3339)
#   "p.p."      -> "ratio"      (idem, es una diferencia de dos fracciones)
#   "u."        -> "units"      (sin cambio de escala -- ya son unidades absolutas en este motor)
# Excepción por nombre de métrica (no por unidad): "EPS" trae en el histórico REAL migrado la
# unidad "%" mal etiquetada en el archivo de origen (bug de datos de MOTOR_VALIDADO_R2, el valor
# en sí YA es USD/acción, no un porcentaje) -- se corrige la ETIQUETA a "USD/acción" sin tocar el
# valor. Cualquier unidad no listada arriba (incluye unit=None -- 770 filas del histórico cuya
# columna Unidad viene vacía de origen, dato no inventado) se deja SIN CONVERTIR y así se reporta
# (columna source_unit permite auditar qué filas quedan fuera de esta normalización).
# ---------------------------------------------------------------------------------------------
_UNIT_CANON_MAP = {
    "miles USD": ("USD", 1000.0),
    "USD": ("USD", 1.0),
    "%": ("ratio", 0.01),
    "p.p.": ("ratio", 0.01),
    "u.": ("units", 1.0),
    "USD/acción": ("USD/acción", 1.0),
}
# Metrics PLAN cuyo valor en 05_HISTORICO_EQUIPOS!G viene con un factor x100 respecto a la fracción
# nativa del motor (ver build_historico_equipos, sección "NOTA UNIDADES") -- para canonicalizar a
# decimal puro hay que dividir /100 en vez de solo copiar. Metrics PLAN de unidades físicas ("u.")
# no requieren cambio de valor, solo de etiqueta ("u." -> "units").
_RATIO_METRICS_PLAN = {"Margen EBITDA", "ROS", "ROE", "ROCE", "ROCE − WACC", "Payout ratio",
                        "Cuota de mercado CADIZ (ponderada por volumen)", "Utilización de capacidad",
                        "Retorno total acumulado del accionista (PROXY)"}
_UNITS_METRICS_PLAN = {"Demanda insatisfecha CADIZ (total)", "Ventas totales (unidades)",
                        "Demanda estimada CADIZ (unidades)", "Capacidad operativa",
                        "Ventas efectivas (mercado)"}

_VARIANT_RE = re.compile(r"^(.*)\s\(variante \d+\)$")


def _strip_variant(metrica):
    m = _VARIANT_RE.match(metrica or "")
    return m.group(1) if m else (metrica or "")


def _canonicalize_real(metric_base, raw_value, raw_unit):
    """Cambio 3 -- transforma (value, unit) de una fila REAL/histórico estático a la convención
    canónica. Devuelve (valor_canonico, unit_canonico)."""
    if metric_base == "EPS":
        return raw_value, "USD/acción"
    if not isinstance(raw_value, (int, float)) or isinstance(raw_value, bool):
        return raw_value, raw_unit
    mapped = _UNIT_CANON_MAP.get(raw_unit)
    if mapped is None:
        return raw_value, raw_unit
    unit, factor = mapped
    return raw_value * factor, unit


def build_data_export(wb, ws_hist, header_row, ids, plan_start):
    ws = wb.create_sheet(DATA_EXPORT_SHEET)
    S.set_col_widths(ws, [8, 14, 10, 10, 10, 12, 46, 20, 12, 20, 12, 12])
    headers = ["round", "scenario", "status", "team", "region", "technology", "metric", "value", "unit",
               "source_value", "source_unit", "is_canonical"]
    S.style_header_row(ws, 1, headers)
    escenario_row = ids["escenario"]

    # ---- Cambio 5 (v1.1) -- Pre-pase: tamaño de cada grupo (ronda,equipo,región,tecnología,
    # métrica-base) entre las filas REAL, para saber qué filas son ambiguas ("(variante N)" o su
    # contraparte sin sufijo -- ambas comparten el mismo grupo de colisión real, ver README). No
    # se toca 05_HISTORICO_EQUIPOS: esto solo lee lo que ya está escrito ahí. ----
    group_count = {}
    for src_row in range(header_row + 1, ws_hist.max_row + 1):
        equipo = ws_hist.cell(row=src_row, column=2).value
        if equipo is None:
            continue
        if ws_hist.cell(row=src_row, column=3).value != "REAL":
            continue
        key = (ws_hist.cell(row=src_row, column=1).value, equipo,
               ws_hist.cell(row=src_row, column=4).value, ws_hist.cell(row=src_row, column=5).value,
               _strip_variant(ws_hist.cell(row=src_row, column=6).value))
        group_count[key] = group_count.get(key, 0) + 1

    out_row = 2
    n_canonical = 0
    n_no_canonical = 0
    for src_row in range(header_row + 1, ws_hist.max_row + 1):
        equipo = ws_hist.cell(row=src_row, column=2).value
        if equipo is None:
            continue
        # Cambio 2 (v1.2): la clasificación PLAN/SIM vs. histórico/competidores YA NO se decide
        # leyendo el texto de la columna Estado de 05_HISTORICO_EQUIPOS (esa columna ahora es una
        # fórmula linkeada a 01_INPUTS, no el texto fijo "PLAN" -- ver build_historico_equipos), sino
        # por RANGO DE FILAS: toda fila desde plan_start en adelante es CADIZ PLAN/SIM (generada por
        # el loop genérico R2-R12), y toda fila anterior es histórico/competidores migrado de
        # MOTOR_VALIDADO_R2 (R0/R1, sin cambios). Si ninguna ronda R2-R12 está activa, plan_start
        # queda en una fila posterior a la última escrita y esta rama nunca se ejecuta -- 0 filas
        # PLAN, tal como exige el contrato.
        es_plan = src_row >= plan_start
        if es_plan:
            # CADIZ Plan/Sim: LINKEADO POR FÓRMULA (no texto estático) -- 05_HISTORICO ya está
            # linkeado al motor, así que esta fórmula queda linkeada transitivamente al motor.
            # El número de ronda de esta fila específica (columna A de 05_HISTORICO_EQUIPOS) ya se
            # conoce en Python (se escribió como valor literal en build_historico_equipos, dentro
            # del loop por ronda) -- se usa para apuntar la columna de Escenario correcta en
            # 01_INPUTS, en vez de asumir siempre la Ronda 2 (Cambio 2, v1.2).
            metrica_plan = ws_hist.cell(row=src_row, column=6).value
            rn_fila = ws_hist.cell(row=src_row, column=1).value
            S.apply_cell(ws, out_row, 1, value=f"='{HIST_SHEET}'!A{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 2, value=f"='{INPUTS_SHEET}'!{col(rn_fila)}{escenario_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 3, value=f"='{HIST_SHEET}'!C{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 4, value=f"='{HIST_SHEET}'!B{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 5, value=f"='{HIST_SHEET}'!D{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 6, value=f"='{HIST_SHEET}'!E{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 7, value=f"='{HIST_SHEET}'!F{src_row}", kind="link", align=S.ALIGN_LEFT, size=8)
            # Cambio 3: override puntual de value/unit para metrics PLAN cuya unidad de origen en
            # 05_HISTORICO_EQUIPOS (% x100, o "u.") no es todavía la canónica -- el resto (dinero
            # ya en USD absoluto, EPS ya en USD/acción) se deja como link directo sin cambios.
            if metrica_plan in _RATIO_METRICS_PLAN:
                # Bug real, corregido (mismo motivo que el IFERROR agregado en 05_HISTORICO_EQUIPOS/
                # PLAN/SIM CADIZ): si la celda origen viene en blanco ("" de un IFERROR previo, p.ej.
                # Ingresos=0 en una ronda con decisiones incompletas), dividir "" /100 da #VALUE! en
                # vez de propagar el blank -- se envuelve en IFERROR para que el blank se propague.
                S.apply_cell(ws, out_row, 8, value=f"=IFERROR('{HIST_SHEET}'!G{src_row}/100,\"\")", kind="link", numfmt='0.0000', align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, out_row, 9, value="ratio", kind="plain", align=S.ALIGN_CENTER, size=8)
            elif metrica_plan in _UNITS_METRICS_PLAN:
                S.apply_cell(ws, out_row, 8, value=f"='{HIST_SHEET}'!G{src_row}", kind="link", numfmt='#,##0.####', align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, out_row, 9, value="units", kind="plain", align=S.ALIGN_CENTER, size=8)
            else:
                S.apply_cell(ws, out_row, 8, value=f"='{HIST_SHEET}'!G{src_row}", kind="link", numfmt='#,##0.####', align=S.ALIGN_RIGHT, size=8)
                S.apply_cell(ws, out_row, 9, value=f"='{HIST_SHEET}'!H{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 10, value=f"='{HIST_SHEET}'!G{src_row}", kind="link", numfmt='#,##0.####', align=S.ALIGN_RIGHT, size=8)
            S.apply_cell(ws, out_row, 11, value=f"='{HIST_SHEET}'!H{src_row}", kind="link", align=S.ALIGN_CENTER, size=8)
            S.apply_cell(ws, out_row, 12, value=True, kind="plain", align=S.ALIGN_CENTER, size=8)
            n_canonical += 1
        else:
            # Histórico/competidores: COPIADO (valor estático) desde 05_HISTORICO_EQUIPOS, que ya
            # es el dato migrado inmutable de MOTOR_VALIDADO_R2 -- evita ~70.000 fórmulas redundantes
            # para mantener este workbook chico (objetivo del proyecto), sin perder trazabilidad
            # (05_HISTORICO_EQUIPOS documenta la fuente exacta de cada fila).
            vals = [ws_hist.cell(row=src_row, column=c).value for c in (1, 2, 3, 4, 5, 6, 7, 8)]
            aligns = [S.ALIGN_CENTER, S.ALIGN_CENTER, S.ALIGN_CENTER, S.ALIGN_CENTER, S.ALIGN_CENTER, S.ALIGN_LEFT, S.ALIGN_RIGHT, S.ALIGN_CENTER]
            metric_full = vals[5]
            metric_base = _strip_variant(metric_full)
            raw_value, raw_unit = vals[6], vals[7]
            canon_value, canon_unit = _canonicalize_real(metric_base, raw_value, raw_unit)
            key = (vals[0], equipo, vals[3], vals[4], metric_base)
            is_canon = group_count.get(key, 1) <= 1
            S.apply_cell(ws, out_row, 1, value=vals[0], kind="historico", align=aligns[0], size=8)
            S.apply_cell(ws, out_row, 2, value="n/a", kind="plain", align=S.ALIGN_CENTER, size=8, italic=True)
            S.apply_cell(ws, out_row, 3, value=vals[2], kind="historico", align=aligns[2], size=8)
            S.apply_cell(ws, out_row, 4, value=vals[1], kind="historico", align=aligns[1], size=8)
            S.apply_cell(ws, out_row, 5, value=vals[3], kind="historico", align=aligns[3], size=8)
            S.apply_cell(ws, out_row, 6, value=vals[4], kind="historico", align=aligns[4], size=8)
            S.apply_cell(ws, out_row, 7, value=vals[5], kind="historico", align=aligns[5], size=8)
            S.apply_cell(ws, out_row, 8, value=canon_value, kind="historico", numfmt='#,##0.####', align=aligns[6], size=8)
            S.apply_cell(ws, out_row, 9, value=canon_unit, kind="historico", align=aligns[7], size=8)
            S.apply_cell(ws, out_row, 10, value=raw_value, kind="historico", numfmt='#,##0.####', align=aligns[6], size=8)
            S.apply_cell(ws, out_row, 11, value=raw_unit, kind="historico", align=aligns[7], size=8)
            S.apply_cell(ws, out_row, 12, value=is_canon, kind="historico", align=S.ALIGN_CENTER, size=8)
            if is_canon:
                n_canonical += 1
            else:
                n_no_canonical += 1
        out_row += 1
    ws.auto_filter.ref = f"A1:L{out_row-1}"
    ws.freeze_panes = "A2"
    ws.sheet_state = "hidden"
    print(f"  [DATA_EXPORT] is_canonical=TRUE: {n_canonical} | is_canonical=FALSE: {n_no_canonical}")
    return ws, out_row - 2


def _detectar_ronda_desde_titulo(path):
    """Respaldo de detectar_ronda_desde_nombre() para cuando el nombre del ARCHIVO no dice 'rondaN'
    (Bug real, corregido -- descubierto contra el repo real de CADIZ: sus RDOS oficiales están
    nombrados 'results-r01.xls', etc., que no matchea ningún patrón de nombre razonable). CESIM
    siempre exporta el RDOS con la hoja 'Results' titulada 'Ronda N' en la celda superior izquierda
    (A1) -- mismo criterio que usaba el viejo flujo de carga manual (file_uploader) de app.py antes
    de esta migración a rutas fijas. Devuelve None si no se puede leer el archivo o si el título no
    matchea (nunca levanta excepción -- un archivo no reconocible se ignora, no rompe el descubrimiento
    de los demás)."""
    try:
        fmt = rdos_parser._sniff_formato(path)
        if fmt == "xlsx":
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            titulo = wb["Results"].cell(row=1, column=1).value
            wb.close()
        else:
            wb = xlrd.open_workbook(path)
            titulo = wb.sheet_by_name("Results").cell(0, 0).value
    except Exception:
        return None
    m = re.search(r"Ronda\s*(\d+)", str(titulo), re.IGNORECASE)
    return int(m.group(1)) if m else None


def descubrir_rdos_oficiales(base_dir):
    """Escanea base_dir de forma RECURSIVA (cualquier subcarpeta debajo, ej. 'data/raw/practicas/
    oficial/' -- el repo real de CADIZ tiene esta anidación un poco rara, y no hace falta forzar una
    reorganización del repo para que este descubrimiento funcione) buscando RDOS oficiales (.xls/
    .xlsx) y arma {ronda_int: path}. Detecta la ronda en dos pasos: primero por el NOMBRE del archivo
    (detectar_ronda_desde_nombre -- 'ronda0.xlsx', 'Ronda_2.xls', 'RDOS RONDA 3.xls'), y si el nombre
    no la delata (ej. 'results-r01.xls'), abre el archivo y lee el título de la hoja 'Results'
    ('Ronda N', ver _detectar_ronda_desde_titulo) -- así funciona sin importar cómo se llame el
    archivo, siempre que sea un RDOS real de CESIM. Si dos archivos matchean la misma ronda, gana el
    de modificación más reciente (mtime) -- no se inventa otro criterio de desempate; se imprime un
    aviso para que quede visible en los logs de la app. Devuelve {} si el directorio no existe o no
    hay ningún archivo reconocible (no es un error -- puede ser la primera corrida)."""
    out = {}
    if not os.path.isdir(base_dir):
        return out
    paths_encontrados = []
    for carpeta_actual, _subcarpetas, nombres in os.walk(base_dir):
        for nombre in sorted(nombres):
            if nombre.lower().endswith((".xls", ".xlsx")):
                paths_encontrados.append(os.path.join(carpeta_actual, nombre))
    for path in sorted(paths_encontrados):
        nombre = os.path.basename(path)
        rn = detectar_ronda_desde_nombre(nombre)
        if rn is None:
            rn = _detectar_ronda_desde_titulo(path)
        if rn is None:
            continue
        if rn in out and os.path.getmtime(path) <= os.path.getmtime(out[rn]):
            print(f"  [descubrir_rdos_oficiales] Ronda {rn}: se ignora '{path}' (más viejo que '{out[rn]}').")
            continue
        out[rn] = path
    return out


_PROYECCION_FILENAME_RE = re.compile(r"^Cadiz_proyeccion_R(\d+)\.xlsx?$", re.IGNORECASE)


def descubrir_ultima_proyeccion(dir_decisiones="data/decisiones"):
    """Escanea dir_decisiones (p.ej. 'data/decisiones/') buscando archivos 'Cadiz_proyeccion_R{N}.xlsx'
    y devuelve el path del de N más alto -- ese es siempre el más actual (contiene todo el histórico
    REAL ya migrado + la proyección PLAN de la ronda en curso + los planes congelados de rondas
    previas en DATA_EXPORT, ver extraer_plan_congelado/inyectar_plan_congelado), así que no hace
    falta ningún otro criterio de "más reciente" (ni mtime ni orden de commit).

    Bug real, corregido (Fase 4 -- integración repo/tablero): antes de esta función, el tablero web
    (app.py -> gap_analysis.load_proyeccion() -> export_proyeccion.DEFAULT_EXCEL) leía un archivo
    DISTINTO y completamente desacoplado de esta arquitectura -- 'CADIZ_Gestion_v2.xlsx' en la raíz
    del repo, que el usuario tenía que reemplazar A MANO con el mismo nombre fijo cada vez. Subir
    'Cadiz_proyeccion_R3.xlsx' a data/decisiones/ (el flujo que pide esta arquitectura) no tenía
    ningún efecto sobre el tablero, que seguía mostrando lo último que hubiera en CADIZ_Gestion_v2.xlsx
    (desactualizado, o inexistente) -- de ahí que el Plan vs. Real apareciera vacío. Ver
    app.get_proyeccion(), que ahora llama a esta función en lugar de usar DEFAULT_PROYECCION_EXCEL.

    Devuelve None si el directorio no existe o no hay ningún archivo que matchee el patrón exacto
    'Cadiz_proyeccion_R{N}.xlsx' (no es un error -- puede ser la primera corrida, antes de generar
    cualquier proyección)."""
    if not os.path.isdir(dir_decisiones):
        return None
    mejor_rn, mejor_path = None, None
    for nombre in os.listdir(dir_decisiones):
        m = _PROYECCION_FILENAME_RE.match(nombre)
        if not m:
            continue
        rn = int(m.group(1))
        if mejor_rn is None or rn > mejor_rn:
            mejor_rn, mejor_path = rn, os.path.join(dir_decisiones, nombre)
    return mejor_path


# Claves de 01_INPUTS que NUNCA se restauran por rescate de decisiones -- se recalculan siempre
# desde rondas_reales (frontera REAL/PLAN), aplicado DESPUÉS del rescate para garantizar que ganan
# ellas y no un valor viejo heredado del archivo anterior (ver aplicar_overrides_inputs()).
_INPUTS_KEYS_FORZADAS = {
    "A. Identificación||Estado||-",
    "A. Identificación||RONDA_ACTIVA (calculado)||-",
    "A. Identificación||Número de ronda||-",
}


def leer_decisiones_previas(path_excel_anterior, desde_ronda):
    """Lee el 01_INPUTS de un Cadiz_proyeccion_R{N}.xlsx ANTERIOR (data_only=False, para poder
    distinguir un valor literal tipeado por el usuario de una fórmula de arrastre '=G10' -- las
    fórmulas de arrastre NO se rescatan porque build_inputs()/arrastre_rows() ya las regenera solas,
    correctamente, en el archivo nuevo) y devuelve {(bloque||variable||dimension): {ronda: valor}}
    para TODAS las columnas de ronda >= desde_ronda que tengan un valor literal cargado -- esto es
    lo que el usuario pidió como "rescatar las decisiones futuras cargadas para que no se pierdan".
    Devuelve {} si el archivo no existe (primera corrida / bootstrap, nada que rescatar)."""
    if not path_excel_anterior or not os.path.exists(path_excel_anterior):
        return {}
    wb = openpyxl.load_workbook(path_excel_anterior, data_only=False)
    if INPUTS_SHEET not in wb.sheetnames:
        return {}
    ws = wb[INPUTS_SHEET]
    overrides = {}
    # Bug real, corregido: build_inputs() pone el encabezado de 01_INPUTS en la fila 4 (ver
    # header_row=4 en build_inputs) -- iterar desde min_row=1 incluía esa fila de encabezado
    # ("Bloque"/"Variable"/"Dimensión" literales en A4/B4/C4, "Ronda N" en las columnas de ronda)
    # como si fuera una fila de datos real, generando una clave espuria
    # "Bloque||Variable||Dimensión" -> {3: "Ronda 3", ...} que después aparecía como "sin match"
    # en los diagnósticos de overrides (inofensiva porque nunca matcheaba una key real de
    # row_idx, pero de todos modos incorrecta). Se arranca después del encabezado.
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        bloque, variable, dimension = row[0].value, row[1].value, row[2].value
        if not variable:
            continue
        key = f"{bloque}||{variable}||{dimension}"
        if key in _INPUTS_KEYS_FORZADAS:
            continue
        vals = {}
        for rn in range(desde_ronda, 13):
            col_idx = 7 + rn
            if col_idx - 1 >= len(row):
                continue
            v = row[col_idx - 1].value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                continue  # fórmula (arrastre u otra) -- se regenera sola, no se rescata
            vals[rn] = v
        if vals:
            overrides[key] = vals
    wb.close()
    return overrides


def aplicar_overrides_inputs(ws_in, row_idx, overrides, desde_ronda):
    """Overlay POST-build_inputs(): escribe los valores literales rescatados por
    leer_decisiones_previas() directamente en las celdas ya construidas de 01_INPUTS, emparejando
    por CLAVE DE TEXTO (bloque||variable||dimension) contra row_idx -- no por número de fila, así que
    tolera reordenamientos de secciones entre versiones del script. Nunca pisa una celda que ya tenga
    una fórmula (el motor nuevo, p.ej. RONDA_ACTIVA o un arrastre) -- solo pisa celdas vacías o con
    otro valor literal. Devuelve (aplicados, sin_match) para loggear en la UI."""
    aplicados, sin_match = 0, 0
    for key, valores_por_ronda in overrides.items():
        row = row_idx.get(key)
        if row is None:
            sin_match += 1
            continue
        for rn, val in valores_por_ronda.items():
            if rn < desde_ronda:
                continue
            cell = ws_in.cell(row=row, column=7 + rn)
            actual = cell.value
            if isinstance(actual, str) and actual.startswith("="):
                continue
            cell.value = val
            aplicados += 1
    return aplicados, sin_match


def extraer_plan_congelado(path_excel_anterior, ronda_actual_a_decidir):
    """Lee el DATA_EXPORT de un Cadiz_proyeccion_R{N}.xlsx ANTERIOR *ya recalculado* (data_only=True
    -- necesita valores cacheados; ver _recalcular_con_libreoffice) y devuelve TODAS las filas CADIZ
    con status="PLAN" cuya ronda sea DISTINTA de `ronda_actual_a_decidir` -- es decir, cualquier
    proyección ya congelada de una ronda que en el archivo nuevo va a ser REAL o ya lo era (rondas <
    ronda_actual_a_decidir), tanto si se congeló recién (la ronda que acaba de pasar a REAL con este
    mismo cambio de frontera) como si ya venía arrastrada de una congelación anterior. Se excluye
    deliberadamente la ronda `ronda_actual_a_decidir` en sí: esa es la ronda TODAVÍA en curso (sin
    RDOS todavía) -- si el archivo anterior ya tenía algo cargado para ella, el archivo nuevo la va a
    recalcular con fórmulas vivas propias (no hace falta, y sería redundante, copiarla como estática).

    Por qué "TODAS" y no solo la última: si no se arrastran las filas YA congeladas en corridas
    anteriores (p.ej. al simplemente re-generar sin RDOS nuevo, para afinar decisiones), se
    perderían silenciosamente en cada regeneración -- el dashboard de Plan vs. Real necesita
    conservarlas indefinidamente. Devuelve None si no hay nada para congelar (bootstrap, o la
    ronda actual nunca tuvo RDOS/decisiones)."""
    if not path_excel_anterior or not os.path.exists(path_excel_anterior):
        return None
    wb = openpyxl.load_workbook(path_excel_anterior, data_only=True)
    if DATA_EXPORT_SHEET not in wb.sheetnames:
        return None
    ws = wb[DATA_EXPORT_SHEET]
    filas = [row for row in ws.iter_rows(min_row=2, values_only=True)
             if row[3] == "CADIZ" and row[2] == "PLAN" and row[0] != ronda_actual_a_decidir]
    wb.close()
    return filas or None


def inyectar_plan_congelado(ws_export, filas_congeladas):
    """Agrega filas_congeladas (de extraer_plan_congelado) al FINAL del DATA_EXPORT ya construido,
    como valores ESTÁTICOS (no fórmulas) -- es la "foto" del Plan Congelado que pidió el usuario:
    coexiste con las filas REAL de la misma ronda (status distinto: "PLAN" congelado vs. "REAL"
    migrado del RDOS) para que el dashboard de Control de Gestión pueda comparar Plan vs. Real sin
    perder lo que se había proyectado antes de tener el resultado real. Devuelve la cantidad de filas
    agregadas."""
    if not filas_congeladas:
        return 0
    out_row = ws_export.max_row + 1
    aligns = [S.ALIGN_CENTER] * 6 + [S.ALIGN_RIGHT, S.ALIGN_CENTER, S.ALIGN_RIGHT, S.ALIGN_CENTER, S.ALIGN_CENTER]
    for row_vals in filas_congeladas:
        # metric (índice 6, columna 7) va alineado a la izquierda; el resto sigue aligns[] de arriba.
        for i, v in enumerate(row_vals, start=1):
            align = S.ALIGN_LEFT if i == 7 else (aligns[i - 1] if i - 1 < len(aligns) else S.ALIGN_CENTER)
            numfmt = '#,##0.####' if i in (8, 10) else None
            S.apply_cell(ws_export, out_row, i, value=v, kind="historico", numfmt=numfmt, align=align, size=8)
        out_row += 1
    return len(filas_congeladas)


def _recalcular_con_libreoffice(path_xlsx, timeout=120):
    """Recalcula TODAS las fórmulas de path_xlsx in-place, usando LibreOffice headless (mismo
    approach validado varias veces en este proyecto: 0 errores de fórmula, identidades contables
    exactas -- ver informes de fases anteriores). Hace falta porque openpyxl NUNCA evalúa fórmulas:
    sin este paso, un archivo recién generado no tiene valores cacheados para sus celdas de fórmula,
    y el mecanismo de "Plan Congelado" (leer el DATA_EXPORT de un Cadiz_proyeccion_R{N}.xlsx anterior
    con data_only=True) devolvería None para todo.

    Devuelve (True, None) si recalculó OK, o (False, mensaje) si LibreOffice no está disponible o
    falla -- nunca lanza una excepción, para que generar_excel() pueda seguir devolviendo el archivo
    (sin recalcular) en vez de romper toda la generación por un problema del entorno de despliegue.
    """
    if shutil.which("soffice") is None:
        return False, ("LibreOffice ('soffice') no está disponible en este entorno -- el archivo se "
                        "generó SIN recalcular (las fórmulas no tienen valor cacheado hasta que se abra "
                        "una vez en Excel/LibreOffice). Agregar 'libreoffice' a packages.txt si se corre "
                        "en Streamlit Community Cloud.")
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            result = subprocess.run(
                ["soffice", "--headless", "--norestore", "--convert-to",
                 "xlsx:Calc MS Excel 2007 XML", "--outdir", tmpdir, path_xlsx],
                capture_output=True, text=True, timeout=timeout,
                env={**os.environ, "HOME": tmpdir},
            )
        except subprocess.TimeoutExpired:
            return False, f"LibreOffice tardó más de {timeout}s recalculando -- se devuelve el archivo sin recalcular."
        except Exception as e:
            return False, f"Error ejecutando LibreOffice: {e} -- se devuelve el archivo sin recalcular."
        out_path = os.path.join(tmpdir, os.path.basename(path_xlsx))
        if result.returncode != 0 or not os.path.exists(out_path):
            return False, f"LibreOffice devolvió error (code={result.returncode}): {result.stderr[-500:]}"
        shutil.copyfile(out_path, path_xlsx)
    return True, None


# ======================================================================================
# GENERAR_EXCEL — ensambla el workbook completo, fija orden/visibilidad de hojas y devuelve un
# buffer en memoria (io.BytesIO), listo para servir como descarga (Streamlit st.download_button)
# sin escribir nada en disco. rdos_files: dict {ronda_int: path_o_fileobj_del_RDOS} -- TODAS las
# rondas ya jugadas que se quieran migrar como histórico REAL (R0, R1, R2, R3, ...). Reemplaza a
# main(), que queda como wrapper de línea de comandos para pruebas locales (ver más abajo).
# ======================================================================================
def generar_excel(rdos_files, out_buffer=None, recalcular=True, overrides=None, plan_congelado=None):
    """overrides: dict de leer_decisiones_previas() -- decisiones/condiciones literales rescatadas
    del Cadiz_proyeccion_R{N}.xlsx anterior, para no tener que volver a tipear todo cada ronda.
    plan_congelado: lista de extraer_plan_congelado() -- foto de las filas CADIZ de la última ronda
    real, tal como estaban proyectadas ANTES de tener el RDOS real, para el dashboard Plan vs. Real."""
    wb = Workbook()
    wb.remove(wb.active)

    # Ronda N.N (Fase 4, frontera dinámica): rondas_reales = TODAS las rondas con RDOS oficial
    # cargado en rdos_files (data/raw/oficial/ en el flujo de repo) -- ya no un boundary fijo (0,1).
    # "La última ronda con RDOS define la frontera REAL" (spec del usuario). RDOS se suben en orden
    # (0,1,2,...) así que en la práctica es siempre un rango contiguo {0..N}, pero se pasa como
    # frozenset (no un simple N) para que cada función generalice con "rn in rondas_reales" sin
    # asumir contigüidad.
    rondas_reales = frozenset(rdos_files.keys())
    n_next = (max(rondas_reales) + 1) if rondas_reales else 0

    ws_in, ridx_in, ids, cond, dec = build_inputs(wb, rondas_reales=rondas_reales)
    n_overrides_aplicados = n_overrides_sin_match = 0
    if overrides:
        # Rescate de decisiones (spec del usuario: "no quiero tener que modificar las rondas futuras
        # cada vez") -- overlay POST-build_inputs, nunca pisa fórmulas del motor nuevo. Estado/
        # RONDA_ACTIVA quedan siempre excluidos (ver _INPUTS_KEYS_FORZADAS) porque ya los fijó
        # build_inputs() arriba, a partir de rondas_reales -- esa es la fuente de verdad, no el
        # archivo anterior.
        n_overrides_aplicados, n_overrides_sin_match = aplicar_overrides_inputs(ws_in, ridx_in, overrides, desde_ronda=n_next)
    ws_p, ridx_p, r_p, prod_out = build_engine_produccion_part1(wb, ridx_in, rondas_reales=rondas_reales)
    ws_m, ridx_m, r_m, mkt_out = build_engine_mercado(wb, ridx_in, prod_out, sens_rows=ids.get("sens"), rdos_files=rdos_files, rondas_reales=rondas_reales)
    build_engine_produccion_part2(ws_p, ridx_p, ridx_m)
    ws_f, ridx_f, r_f, country_rows, glob, finparams = build_engine_financiero(wb, ridx_in, prod_out, mkt_out, rdos_files, rondas_reales=rondas_reales)
    ws_e = build_estados_proyectados(wb, glob, ridx_in, country_rows, rondas_reales=rondas_reales)
    ws_r, ratios_rows = build_ratios(wb, glob, ridx_in, cond, dec, prod_out, mkt_out, rondas_reales=rondas_reales)
    ws_c = build_control_modelo(wb, ridx_in, prod_out, mkt_out, glob, cond, dec, country_rows=country_rows, rondas_reales=rondas_reales)
    ws_h, n_hist, n_plan, plan_start = build_historico_equipos(
        wb, rdos_files, glob, ratios_rows, prod_out, mkt_out, ridx_in)
    ws_d, n_export = build_data_export(wb, ws_h, 4, ids, plan_start)
    n_congeladas = inyectar_plan_congelado(ws_d, plan_congelado) if plan_congelado else 0

    # Orden final: 5 hojas visibles + hojas ocultas (motores + DATA_EXPORT)
    order = [INPUTS_SHEET, ESTADOS_SHEET, RATIOS_SHEET, CONTROL_SHEET, HIST_SHEET,
             DATA_EXPORT_SHEET, ENGINE_PROD_SHEET, ENGINE_MKT_SHEET, ENGINE_FIN_SHEET]
    wb._sheets = [wb[name] for name in order]
    for name in order:
        wb[name].sheet_view.showGridLines = False
    wb.active = 0

    # ------------------------------------------------------------------------------------
    # Fase 2, cambio 5: ocultar columnas de metadatos (Bloque/Tipo/Nota-Fuente) en todas las hojas
    # de datos EXCEPTO 05_HISTORICO_EQUIPOS (que no tiene esas columnas) -- Dimensión (col. C) queda
    # SIEMPRE VISIBLE (confirmación explícita del usuario). Ocultar, no borrar: el dato sigue
    # existiendo y el lookup dinámico por etiqueta de build_gestion_v2.py sigue funcionando igual.
    # 01_INPUTS: A=Bloque, E=Tipo, F=Fuente. Hojas _ENGINE_*: A=Bloque, E=Tipo, F=Nota.
    # ------------------------------------------------------------------------------------
    for name in (INPUTS_SHEET, ENGINE_PROD_SHEET, ENGINE_MKT_SHEET, ENGINE_FIN_SHEET):
        wsx = wb[name]
        for letter in ("A", "E", "F"):
            wsx.column_dimensions[letter].hidden = True

    # 04_CONTROL_MODELO: A=Control (visible, es la etiqueta principal), B=Tipo, C=Nota -- mismo
    # criterio (ocultar metadatos de estructura, no la etiqueta ni los valores). Se había omitido
    # de este bucle en la primera pasada -- corregido tras verificación independiente.
    ws_control = wb[CONTROL_SHEET]
    for letter in ("B", "C"):
        ws_control.column_dimensions[letter].hidden = True

    # Hotfix visual (reportado por el usuario): en 02_ESTADOS_PROYECTADOS y 03_RATIOS, las columnas
    # C-F nunca tuvieron contenido -- son relleno del header (["Concepto","Unidad","","","",""]) para
    # que "R0" caiga en la columna G, coincidiendo con la convención col(ronda)=7+ronda usada en TODO
    # el workbook (01_INPUTS, _ENGINE_*, etc. -- cambiarla ahí rompería miles de fórmulas, no es lo
    # que se pidió). El hueco entre "Unidad" (B) y "R0" (G) se veía como un bloque vacío sin sentido.
    # Se ocultan esas 4 columnas (mismo criterio de "ocultar, no borrar/mover" del resto del archivo):
    # ninguna fórmula ni referencia cambia, R9-R12 siguen exactamente en P/Q/R/S como corresponde.
    # En 04_CONTROL_MODELO, D-F tienen el mismo relleno vacío (B/C ya ocultas arriba por Tipo/Nota).
    for letter in ("C", "D", "E", "F"):
        wb[ESTADOS_SHEET].column_dimensions[letter].hidden = True
        wb[RATIOS_SHEET].column_dimensions[letter].hidden = True
    for letter in ("D", "E", "F"):
        ws_control.column_dimensions[letter].hidden = True

    recalc_info = {"ok": False, "mensaje": "No solicitado (recalcular=False)."}
    if recalcular:
        # Recalcula con LibreOffice headless (necesario para que las fórmulas -- en particular
        # DATA_EXPORT!status="PLAN" de la ronda activa -- tengan valor cacheado, indispensable para
        # que un futuro "Plan Congelado" pueda leerlas con openpyxl(data_only=True)) y arregla el bug
        # de quoting de hojas con nombre que empieza en dígito que introduce ese mismo recálculo (ver
        # fix_digit_sheet_quoting() y DIGIT_LEAD_SHEETS más arriba). Se hace sobre un archivo temporal
        # en disco porque LibreOffice necesita un path, no un buffer en memoria.
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = os.path.join(tmpdir, "CADIZ_Gestion_tmp.xlsx")
            wb.save(tmp_path)
            ok, msg = _recalcular_con_libreoffice(tmp_path)
            if ok:
                fix_digit_sheet_quoting(tmp_path)
                with open(tmp_path, "rb") as fh:
                    data = fh.read()
                recalc_info = {"ok": True, "mensaje": None}
            else:
                # No se pudo recalcular (LibreOffice ausente o falló): se sirve igual el archivo SIN
                # recalcular (openpyxl, sin quoting-fix porque no hizo falta) en vez de bloquear la
                # generación -- mejor un Excel usable (que Excel/LibreOffice recalculan solos al
                # abrirlo) que ningún archivo.
                buf0 = io.BytesIO()
                wb.save(buf0)
                data = buf0.getvalue()
                recalc_info = {"ok": False, "mensaje": msg}
        buf = out_buffer if out_buffer is not None else io.BytesIO()
        buf.write(data)
        buf.seek(0)
    else:
        buf = out_buffer if out_buffer is not None else io.BytesIO()
        wb.save(buf)
        buf.seek(0)
    print(f"Filas históricas migradas: {n_hist} | Filas PLAN/SIM CADIZ (rondas activas R2-R12): {n_plan} | Filas DATA_EXPORT: {n_export}")
    if overrides:
        print(f"Rescate de decisiones: {n_overrides_aplicados} valores aplicados, {n_overrides_sin_match} claves sin match en 01_INPUTS")
    if plan_congelado:
        print(f"Plan Congelado: {n_congeladas} filas inyectadas en DATA_EXPORT")
    print(f"Recálculo LibreOffice: {'OK' if recalc_info['ok'] else 'NO -- ' + str(recalc_info['mensaje'])}")
    info = dict(recalc_info)
    info.update(rondas_reales=sorted(rondas_reales), ronda_a_decidir=n_next,
                overrides_aplicados=n_overrides_aplicados, overrides_sin_match=n_overrides_sin_match,
                plan_congelado_filas=n_congeladas)
    return buf, info


# ======================================================================================
# GENERAR_EXCEL_DESDE_REPO — punto de entrada del flujo "todo en rutas fijas del repo" (spec del
# usuario, Fase 4): sin file_uploader en Streamlit. Detecta automáticamente la frontera REAL a
# partir de dir_oficial, arma el nombre del archivo de salida Cadiz_proyeccion_R{N+1}.xlsx, rescata
# decisiones futuras + Plan Congelado del último Cadiz_proyeccion_R{N}.xlsx en dir_decisiones (si
# existe -- si no, es la primera corrida / bootstrap y arranca de cero, sin overrides ni plan
# congelado, que es un resultado válido, no un error) y genera el workbook ya recalculado.
# ======================================================================================
def generar_excel_desde_repo(dir_oficial="data/raw/practicas/oficial", dir_decisiones="data/decisiones", recalcular=True):
    rdos_files = descubrir_rdos_oficiales(dir_oficial)
    if not rdos_files:
        raise FileNotFoundError(
            f"No se encontró ningún RDOS oficial reconocible en '{dir_oficial}' (se esperaba algo "
            f"como 'ronda0.xlsx', 'ronda1.xlsx', ... -- ver detectar_ronda_desde_nombre()). No se "
            f"puede determinar la frontera REAL sin al menos un RDOS.")
    # Bug real, corregido -- Ronda 0 es la ronda FUNDACIONAL de todo el modelo (saldos de apertura,
    # semilla FIFO de inventario, etc. -- toda la cadena de balance del resto de las rondas arranca
    # de ahí). Si su RDOS no está entre los descubiertos, "Estado" y "RONDA_ACTIVA" de Ronda 0 quedan
    # sin definir (ninguna ronda R1+ puede migrarse como histórica sin que Ronda 0 también lo esté,
    # ver build_inputs/build_historico_equipos) y el workbook queda con cientos de #VALUE! en cascada
    # en vez de fallar con un mensaje claro -- verificado con una corrida de prueba deliberadamente
    # sin RDOS de Ronda 0 (312 errores de fórmula, Ronda 0 ausente de DATA_EXPORT). Antes se
    # confiaba en que Ronda 0 fuera la primera SIEMPRE presente; ahora se lo exige explícitamente en
    # vez de permitir un workbook roto en silencio.
    if 0 not in rdos_files:
        raise FileNotFoundError(
            f"Se encontraron RDOS de las rondas {sorted(rdos_files)} en '{dir_oficial}', pero falta "
            f"el de Ronda 0. Ronda 0 es la base de todo el modelo (saldos de apertura, inventario "
            f"inicial, etc.) -- sin su RDOS ahí, el Excel se genera roto (decenas de #VALUE! en "
            f"cascada desde Ronda 1 en adelante) en vez de fallar de entrada. Agregá el RDOS oficial "
            f"de Ronda 0 a esa carpeta (con nombre que contenga 'ronda0', o con la hoja 'Results' "
            f"titulada 'Ronda 0' -- ver descubrir_rdos_oficiales/_detectar_ronda_desde_titulo) antes "
            f"de generar de nuevo.")
    n_real = max(rdos_files)
    n_next = n_real + 1
    nombre_salida = f"Cadiz_proyeccion_R{n_next}.xlsx"

    # Ronda N.N: el archivo con las decisiones a rescatar puede ser uno de dos, según si esta corrida
    # AVANZA la frontera (llegó RDOS nuevo) o simplemente REGENERA la misma ronda en curso (el usuario
    # afinó decisiones y corre de nuevo, sin RDOS nuevo todavía):
    #   1. Cadiz_proyeccion_R{n_next}.xlsx -- MISMA ronda que se está por generar de nuevo (existe si
    #      ya se había generado antes para esta misma frontera). Se prioriza porque es la versión más
    #      reciente de las decisiones de la ronda en curso.
    #   2. Cadiz_proyeccion_R{n_real}.xlsx -- el archivo de la corrida ANTERIOR, de cuando n_real
    #      todavía era la ronda a decidir (recién ahora pasa a REAL con el RDOS nuevo).
    # En ambos casos se rescatan columnas >= n_next (incluye la propia ronda n_next: sus decisiones
    # ya cargadas no deben perderse) y se arrastran las filas ya congeladas del Plan (ver
    # extraer_plan_congelado) excluyendo la ronda n_next (todavía no tiene RDOS, no se congela).
    path_mismo_next = os.path.join(dir_decisiones, f"Cadiz_proyeccion_R{n_next}.xlsx")
    path_anterior = os.path.join(dir_decisiones, f"Cadiz_proyeccion_R{n_real}.xlsx")
    path_previo = path_mismo_next if os.path.exists(path_mismo_next) else (path_anterior if os.path.exists(path_anterior) else None)

    overrides, plan_congelado = {}, None
    if path_previo:
        overrides = leer_decisiones_previas(path_previo, desde_ronda=n_next)
        plan_congelado = extraer_plan_congelado(path_previo, ronda_actual_a_decidir=n_next)

    buf, info = generar_excel(rdos_files, recalcular=recalcular, overrides=overrides, plan_congelado=plan_congelado)
    info.update(nombre_salida=nombre_salida, archivo_previo_usado=path_previo,
                rdos_detectados={rn: os.path.basename(p) for rn, p in sorted(rdos_files.items())})
    return buf, nombre_salida, info


# ------------------------------------------------------------------------------------------------
# Wrapper de línea de comandos, solo para pruebas locales (guarda en disco en vez de devolver el
# buffer). El flujo real de producción es generar_excel() llamado desde app.py (Streamlit), que
# nunca toca el disco del servidor. Rutas por defecto = los 4 RDOS ya jugados (R0-R3) en este
# entorno de trabajo -- ajustar CADIZ_BUILD_DIR/nombres si cambian de ubicación.
# ------------------------------------------------------------------------------------------------
def main():
    rdos_files_default = {
        0: "/mnt/user-data/uploads/CESIM/RDOS RONDA 0.xls",
        1: "/mnt/user-data/uploads/CESIM/RDOS RONDA 1.xls",
        2: "/root/.claude/uploads/09e8cf51-b7ac-568d-b78d-5c25fdf40d52/82948c81-1789935829915_RDOS_RONDA_2.xls",
        3: "/root/.claude/uploads/09e8cf51-b7ac-568d-b78d-5c25fdf40d52/3fa98713-1789935781896_RDOS_RONDA_3.xls",
    }
    buf, recalc_info = generar_excel(rdos_files_default)
    with open(OUT_PATH, "wb") as f:
        f.write(buf.getvalue())
    print(f"Guardado: {OUT_PATH}")
    if not recalc_info["ok"]:
        print(f"ADVERTENCIA: {recalc_info['mensaje']}")
    return OUT_PATH


if __name__ == "__main__":
    main()
