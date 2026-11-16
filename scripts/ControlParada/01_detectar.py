"""
PASO 1 · Detectar matrículas que NO están en el diccionario.

Entrada : data/raw/Matricules_Control_Parada.xlsx  (columna MATRICULA)
Salida  : data/processed/control_parada_revision.xlsx

En la revisión sale una fila por cada matrícula mal escrita con:
  - cómo venía escrita, cuántas veces aparece y en qué filas
  - SUGERENCIA (la del diccionario que más se parece) y su SIMILITUD
  - ALTERNATIVAS si hay otras casi igual de probables
  - ESTADO: FIABLE / REVISAR / SIN SUGERENCIA
  - CORRECTA (vacía): escribe aquí la buena SOLO si la sugerencia no te vale
"""

from collections import OrderedDict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from comun import (DICCIONARIO, REVISION, buscar_entrada, normalizar,
                   recorrer_matriculas, sugerir)

print("🔎 Buscando matrículas fuera del diccionario...")

ENTRADA = buscar_entrada()
print("   Archivo:", ENTRADA)

wb = load_workbook(ENTRADA, data_only=True)

total = 0
correctas = 0
solo_formato = 0
malas = OrderedDict()  # normalizada -> info

for ws, celda in recorrer_matriculas(wb):
    total += 1
    original = str(celda.value).strip()
    norm = normalizar(original)

    if norm in DICCIONARIO:
        correctas += 1
        if norm != original:
            solo_formato += 1
        continue

    info = malas.setdefault(norm, {"originales": [], "filas": []})
    if original not in info["originales"]:
        info["originales"].append(original)
    etiqueta = f"{ws.title}!{celda.row}" if len(wb.worksheets) > 1 else str(celda.row)
    info["filas"].append(etiqueta)

# ── construir el Excel de revisión ──────────────────────────
filas_salida = []
for norm, info in malas.items():
    sug, sim, alts, estado = sugerir(info["originales"][0])
    filas_salida.append({
        "MATRICULA_ORIGINAL": " | ".join(info["originales"]),
        "NORMALIZADA": norm,
        "APARICIONES": len(info["filas"]),
        "FILAS_EXCEL": ", ".join(info["filas"]),
        "SUGERENCIA": sug,
        "SIMILITUD": sim,
        "ALTERNATIVAS": ", ".join(alts),
        "ESTADO": estado,
        "CORRECTA": "",
    })

orden_estado = {"SIN SUGERENCIA": 0, "REVISAR": 1, "FIABLE": 2}
filas_salida.sort(key=lambda r: (orden_estado[r["ESTADO"]], r["NORMALIZADA"]))

out = Workbook()
ws_out = out.active
ws_out.title = "Revision"
cabeceras = ["MATRICULA_ORIGINAL", "NORMALIZADA", "APARICIONES", "FILAS_EXCEL",
             "SUGERENCIA", "SIMILITUD", "ALTERNATIVAS", "ESTADO", "CORRECTA"]
ws_out.append(cabeceras)
for r in filas_salida:
    ws_out.append([r[c] for c in cabeceras])

# formato de la hoja de revisión
negrita = Font(bold=True, color="FFFFFF")
fondo_cab = PatternFill("solid", start_color="305496")
colores = {
    "FIABLE": PatternFill("solid", start_color="C6EFCE"),
    "REVISAR": PatternFill("solid", start_color="FFEB9C"),
    "SIN SUGERENCIA": PatternFill("solid", start_color="FFC7CE"),
}
amarillo_claro = PatternFill("solid", start_color="FFF2CC")

for c in ws_out[1]:
    c.font = negrita
    c.fill = fondo_cab
    c.alignment = Alignment(horizontal="center")

col_estado = cabeceras.index("ESTADO") + 1
col_correcta = cabeceras.index("CORRECTA") + 1
col_sim = cabeceras.index("SIMILITUD") + 1
for fila in range(2, ws_out.max_row + 1):
    estado = ws_out.cell(row=fila, column=col_estado).value
    ws_out.cell(row=fila, column=col_estado).fill = colores[estado]
    ws_out.cell(row=fila, column=col_correcta).fill = amarillo_claro
    ws_out.cell(row=fila, column=col_sim).number_format = "0%"

anchos = [24, 14, 13, 30, 14, 11, 26, 16, 14]
for i, ancho in enumerate(anchos, start=1):
    ws_out.column_dimensions[get_column_letter(i)].width = ancho
ws_out.freeze_panes = "A2"
if ws_out.max_row > 1:
    ws_out.auto_filter.ref = ws_out.dimensions

out.save(REVISION)

# ── resumen por consola ─────────────────────────────────────
n_malas = sum(r["APARICIONES"] for r in filas_salida)
print(f"\n   Matrículas leídas:          {total}")
print(f"   ✔ En el diccionario:        {correctas}  (de ellas {solo_formato} solo cambian de formato)")
print(f"   ✖ Fuera del diccionario:    {n_malas}  ({len(filas_salida)} distintas)")

if filas_salida:
    print("\n   ORIGINAL            →  SUGERENCIA   SIMIL.  ESTADO")
    for r in filas_salida:
        extra = f"  (o {r['ALTERNATIVAS']})" if r["ALTERNATIVAS"] else ""
        print(f"   {r['MATRICULA_ORIGINAL'][:18]:<18}  →  {r['SUGERENCIA'] or '—':<11}  "
              f"{r['SIMILITUD']:>5.0%}  {r['ESTADO']}{extra}  x{r['APARICIONES']}")

print("\n✅ Revisión creada:", REVISION)
print("   👉 Revisa las amarillas/rojas. Si una sugerencia no vale, escribe la buena en CORRECTA.")
print("   👉 Después ejecuta 02_corregir.py")
