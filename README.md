# Tablero CADIZ — Cesim Global Automotive

## Cómo cargar una ronda nueva (RDOS real de CESIM)
1. Descargá el `.xls` de resultados de la ronda desde Cesim.
2. Subilo a `data/raw/` en este repo (drag & drop en github.com, o `git add` + commit + push).
3. Listo — la app lee todos los `.xls` de `data/raw/` en cada carga y arma el histórico sola.
   No hace falta correr ningún script aparte.

## Cómo actualizar la proyección de CADIZ (Control de Gestión: Proyectado vs. Real)
1. Reemplazá `CADIZ_Gestion_v2.xlsx` en la **raíz** de este repo por la versión más nueva del
   modelo de gestión — **siempre el mismo nombre de archivo** (drag & drop en github.com sobre el
   archivo existente, o `git add` + commit + push).
2. Listo — la próxima carga de la app lee `DATA_EXPORT` directamente de ese Excel y la toma como
   "la proyección más actual". No hace falta exportar ningún CSV a mano ni correr ningún script
   aparte (el script `export_proyeccion.py` sigue existiendo solo como utilidad opcional de
   diagnóstico/respaldo, no es parte de este flujo).

## Deploy
1. Subir esta carpeta como repo a GitHub.
2. En share.streamlit.io: New app → apuntar a este repo → `app.py` como entry point.
3. Cada `git push` (incluido subir una ronda nueva o una versión nueva del Excel) redeploya la app
   sola en 1-2 min.

## Estructura
- `cesim_parser.py` — toda la lógica de parseo del .xls de Cesim (reutilizable).
- `app.py` — la app Streamlit (sidebar con los módulos, lee data/raw/ automáticamente).
- `data/raw/` — los .xls de cada ronda, tal cual los bajás de Cesim.
- `CADIZ_Gestion_v2.xlsx` — el modelo de proyección de CADIZ (raíz del repo, mismo nombre siempre);
  `gap_analysis.py` lee su hoja `DATA_EXPORT` directamente para el panel de Control de Gestión.
- `metric_crosswalk.py` / `gap_analysis.py` — cruce de métricas y cálculo de gaps Proyectado vs. Real.
- `export_proyeccion.py` — utilidad opcional para volcar `DATA_EXPORT` a un CSV de respaldo manual;
  no es necesaria para que la app funcione.
