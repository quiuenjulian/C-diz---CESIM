"""
cesim_parser.py
Modulo de parseo de exports de Cesim. Se importa directo desde app.py:
no requiere correrse por separado ni generar archivos intermedios.
"""
import re
import pandas as pd
import xlrd

COMPANIES = ['CADIZ', 'CEOS', 'CHIEF', 'CLAVE', 'CUORE', 'FOCUS', 'TOKIO']
N_PRACTICE_ROUNDS = 3  # Cesim: 3 rondas de practica + 12 oficiales

MODULE_MAP = {
    'Cuenta de resultados, miles USD, Global': 'Estados financieros',
    'Hoja de Balance, miles USD, Global': 'Estados financieros',
    'Cuenta de resultados, miles USD, EE.UU.': 'Estados financieros',
    'Hoja de Balance, miles USD, EE.UU.': 'Estados financieros',
    'Flujo de efectivo de casa matriz, miles USD': 'Estados financieros',
    'Cuenta de resultados, miles USD, China': 'Estados financieros',
    'Hoja de Balance, miles USD, China': 'Estados financieros',
    'Estado de flujo de efectivo, miles USD, China': 'Estados financieros',
    'Cuenta de resultados, miles USD, Europa': 'Estados financieros',
    'Hoja de Balance, miles USD, Europa': 'Estados financieros',
    'Estado de flujo de efectivo, miles USD, Europa': 'Estados financieros',
    'Ratios e indicadores financieros clave': 'Ratios',
    'Informe de mercado, global': 'Informes de mercado',
    'Informe de mercado, EE.UU.': 'Informes de mercado',
    'Informe de mercado, China': 'Informes de mercado',
    'Informe de mercado, Europa': 'Informes de mercado',
    'Informe de RRHH': 'Informe de RRHH',
    'Informe ESG': 'Sostenibilidad',
    'Informe del proveedor de componentes': 'Informes de producción',
    'Detalles de fabricación': 'Informes de producción',
    'Detalles de logística': 'Informes de producción',
    'Informe de costos': 'Informes de costos',
    'Desglose de margen por tec, miles USD, EE.UU.': 'Informes de costos',
    'Desglose de margen por tec, miles USD, China': 'Informes de costos',
    'Desglose de margen por tec, miles USD, Europa': 'Informes de costos',
    'Valuación - Global': 'Valuación',
    'Valuación - EE.UU.': 'Valuación',
    'Valuación - China': 'Valuación',
    'Valuación - Europa': 'Valuación',
    'Creación de valor, miles USD': 'Creación de valor',
    'Demanda estimada, miles unidades': 'Informes de mercado',
    'Precio de venta': 'Informes de mercado',
}


def detect_round(title: str):
    title = title.strip()
    m = re.search(r'Ronda de pr[aá]ctica\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f'Práctica {n}', 'Práctica', n, n - N_PRACTICE_ROUNDS
    m = re.search(r'Ronda\s*(\d+)', title, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f'Ronda {n}', 'Oficial', n, n
    return title, 'Desconocido', None, None


def parse_cesim_xls(path_or_buffer) -> pd.DataFrame:
    """Parsea UN archivo .xls de Cesim (ruta o file-like buffer) a formato tidy."""
    book = xlrd.open_workbook(file_contents=path_or_buffer.read()) if hasattr(path_or_buffer, 'read') \
        else xlrd.open_workbook(path_or_buffer, formatting_info=True)
    if hasattr(path_or_buffer, 'read'):
        # necesitamos formatting_info tambien en modo buffer
        path_or_buffer.seek(0)
        book = xlrd.open_workbook(file_contents=path_or_buffer.read(), formatting_info=True)

    sheet = book.sheet_by_index(0)
    xf_list = book.xf_list
    font_list = book.font_list

    title_row = sheet.cell(0, 0).value
    ronda, tipo_ronda, numero, orden = detect_round(title_row)

    raw = []
    for r in range(1, sheet.nrows):
        label = sheet.cell(r, 0).value
        vals = [sheet.cell(r, c).value for c in range(1, 8)]
        if label == '' and all(v == '' for v in vals):
            continue
        xf = xf_list[sheet.cell_xf_index(r, 0)]
        font = font_list[xf.font_index]
        size = int(font.height / 20)
        raw.append(dict(label=label.strip(), size=size, vals=vals))

    records = []
    cur_statement = cur_section = cur_subgroup = None

    def is_company_header(vals):
        return vals == COMPANIES

    for item in raw:
        label, size, vals = item['label'], item['size'], item['vals']
        if is_company_header(vals):
            continue
        all_empty = all(v == '' for v in vals)

        if all_empty:
            if size == 14:
                cur_statement, cur_section, cur_subgroup = label, None, None
            elif size == 12:
                cur_section, cur_subgroup = label, None
            else:
                cur_subgroup = label
            continue

        for company, v in zip(COMPANIES, vals):
            records.append({
                'Ronda': ronda,
                'Tipo_Ronda': tipo_ronda,
                'Ronda_Numero': numero,
                'Ronda_Orden': orden,
                'Modulo': MODULE_MAP.get(cur_statement, 'Sin clasificar'),
                'Estado': cur_statement or '',
                'Seccion': cur_section or '',
                'Subgrupo': cur_subgroup or '',
                'Metrica': label,
                'Empresa': company,
                'Valor': v,
            })

    return pd.DataFrame(records)


def build_historico(file_paths) -> pd.DataFrame:
    """Parsea una lista de rutas .xls y devuelve el dataset historico consolidado,
    ordenado cronologicamente. Si dos archivos corresponden a la misma Ronda
    (ej. se resubio corregido), se queda con el ULTIMO archivo de esa ronda
    completo (no mezcla filas fila-por-fila entre ambos, evita perder datos
    legitimos con nombres de metrica repetidos dentro de una misma seccion)."""
    por_ronda = {}  # Ronda -> DataFrame (el ultimo archivo visto para esa ronda gana)
    for p in file_paths:
        frame = parse_cesim_xls(p)
        if frame.empty:
            continue
        ronda = frame['Ronda'].iloc[0]
        por_ronda[ronda] = frame  # sobreescribe si la ronda ya existia

    if not por_ronda:
        return pd.DataFrame()

    df = pd.concat(por_ronda.values(), ignore_index=True)
    df = df.sort_values(['Ronda_Orden', 'Modulo', 'Estado', 'Seccion', 'Metrica', 'Empresa'])
    return df
