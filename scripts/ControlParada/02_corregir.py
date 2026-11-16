"""
PASO 2 · Aplicar las correcciones y generar el Excel corregido.

Entrada : data/raw/Matricules_Control_Parada.xlsx
          data/processed/control_parada_revision.xlsx  (del paso 1)
Salida  : data/processed/Matricules_Control_Parada_corregido.xlsx
          data/processed/control_parada_cambios.xlsx   (registro de cada cambio)

Reglas:
  - Toda la columna MATRICULA queda en formato 0000XXX.
  - Las que están mal se sustituyen por CORRECTA (si la has rellenado)
    o, si no, por la SUGERENCIA.
  - El resto del Excel (otras columnas, colores, anchos, hojas) no se toca.
"""

import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from comun import (DICCIONARIO, LOG_CAMBIOS, REVISION, SALIDA, buscar_entrada,
                   limpiar, normalizar, recorrer_matriculas)

print("⚙️ Aplicando correcciones...")

ENTRADA = buscar_entrada()
if not os.path.exists(REVISION):
    raise FileNotFoundError("Falta la revisión. Ejecuta antes 01_detectar.py")

# ── leer las correcciones decididas ─────────────────────────
rev = load_workbook(REVISION, data_only=True).active
cab = [str(c.value).strip().upper() if c.value else "" for c in rev[1]]
i_norm, i_sug, i_ok = cab.index("NORMALIZADA"), cab.index("SUGERENCIA"), cab.index("CORRECTA")

correcciones = {}
avisos = []
for fila in rev.iter_rows(min_row=2, values_only=True):
    norm = limpiar(fila[i_norm])
    if not norm:
        continue
    manual = limpiar(fila[i_ok])
    final = normalizar(manual) if manual else limpiar(fila[i_sug])
    if not final:
        continue  # sin sugerencia ni corrección manual: se queda como está
    if final not in DICCIONARIO:
        avisos.append(f"{norm} → {final} (no está en el diccionario, se aplica igualmente)")
    correcciones[norm] = final

# ── corregir el Excel original (manteniendo su formato) ─────
wb = load_workbook(ENTRADA)

cambios = []
pendientes = []
total = formato = corregidas = 0

for ws, celda in recorrer_matriculas(wb):
    total += 1
    original = celda.value
    texto = str(original).strip()
    norm = normalizar(texto)

    if norm in DICCIONARIO:
        nuevo, tipo = norm, "FORMATO"
    elif norm in correcciones:
        nuevo, tipo = correcciones[norm], "CORREGIDA"
    else:
        nuevo, tipo = norm, "PENDIENTE"
        pendientes.append(f"{ws.title}!{celda.row}: {texto}")

    if nuevo != original:
        celda.value = nuevo
        if tipo == "FORMATO":
            formato += 1
        elif tipo == "CORREGIDA":
            corregidas += 1
        cambios.append((ws.title, celda.row, texto, nuevo, tipo))

wb.save(SALIDA)

# ── registro de cambios (archivo aparte, no toca el tuyo) ───
log = Workbook()
wl = log.active
wl.title = "Cambios"
wl.append(["HOJA", "FILA", "ANTES", "DESPUES", "TIPO"])
for c in cambios:
    wl.append(list(c))
for c in wl[1]:
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="305496")
for i, ancho in enumerate([16, 8, 20, 14, 13], start=1):
    wl.column_dimensions[get_column_letter(i)].width = ancho
wl.freeze_panes = "A2"
log.save(LOG_CAMBIOS)

# ── resumen ─────────────────────────────────────────────────
print(f"\n   Matrículas procesadas:   {total}")
print(f"   ✏️  Solo formato:          {formato}")
print(f"   🔧 Corregidas:            {corregidas}")
print(f"   ⚠️  Sin corregir:          {len(pendientes)}")
for p in pendientes:
    print("      -", p)
for a in avisos:
    print("   ⚠️ ", a)

print("\n✅ Excel corregido:", SALIDA)
print("   Registro de cambios:", LOG_CAMBIOS)
