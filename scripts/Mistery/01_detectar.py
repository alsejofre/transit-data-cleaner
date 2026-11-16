"""
PASO 1 · Analizar la columna MATRICULA_CALCA del Excel Mistery.

Entrada : data/raw/Mistery.xlsx
Salida  : data/processed/mistery_revision.xlsx  con dos hojas:

  · Diccionario_Calcas -> la calca de cada matrícula, sacada de los datos
                          (la que más se repite). Si alguna está mal,
                          escribe la buena en CALCA_CORRECTA.
  · Revision           -> cada valor distinto que se va a cambiar:
                          matrículas corregidas, calcas corregidas,
                          calcas solas a las que se les añade matrícula
                          y matrículas solas a las que se les añade calca.
                          Si un RESULTADO no vale, escribe el bueno en CORRECTA.
"""

from collections import OrderedDict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import mistery_comun as mc

print("🔎 Analizando MATRICULA_CALCA de Mistery...")

ENTRADA = mc.buscar_entrada()
print("   Archivo:", ENTRADA)

wb = load_workbook(ENTRADA, data_only=True)

valores = OrderedDict()  # clave -> {original, filas}
for ws, celda in mc.recorrer_celdas(wb):
    k = mc.clave(celda.value)
    info = valores.setdefault(k, {"original": str(celda.value).strip(), "filas": []})
    info["filas"].append(f"{ws.title}!{celda.row}" if len(wb.worksheets) > 1 else str(celda.row))

todas = []
for k, info in valores.items():
    todas.extend([k] * len(info["filas"]))

# ── 1) diccionario de calcas ────────────────────────────────
dic = mc.construir_dic_calcas(todas)
mat_calca, calca_mat = mc.mapas_calcas(dic)

# ── 2) resolver cada valor distinto ─────────────────────────
filas_rev = []
resumen = {"MATRICULA + CALCA": 0, "SOLO MATRICULA": 0, "SOLO CALCA": 0, "ILEGIBLE": 0}
estados = {mc.FIABLE: 0, mc.REVISAR: 0, mc.SIN_RESOLVER: 0}
sin_cambio = 0

for k, info in valores.items():
    r = mc.resolver(k, mat_calca, calca_mat)
    n = len(info["filas"])
    resumen[r["TIPO"]] += n
    estados[r["ESTADO"]] += n
    if r["RESULTADO"] == k and r["ESTADO"] == mc.FIABLE:
        sin_cambio += n
        continue
    filas_rev.append({
        "ORIGINAL": info["original"],
        "TIPO": r["TIPO"],
        "RESULTADO": r["RESULTADO"],
        "ESTADO": r["ESTADO"],
        "MOTIVO": r["MOTIVO"],
        "APARICIONES": n,
        "FILAS_EXCEL": ", ".join(info["filas"]),
        "CORRECTA": "",
    })

filas_rev.sort(key=lambda x: (mc.ORDEN_ESTADO[x["ESTADO"]], x["TIPO"], x["ORIGINAL"]))

# ── 3) Excel de revisión ────────────────────────────────────
out = Workbook()
cab_fill = PatternFill("solid", start_color="305496")
cab_font = Font(bold=True, color="FFFFFF")
verde = PatternFill("solid", start_color="C6EFCE")
amarillo = PatternFill("solid", start_color="FFEB9C")
rojo = PatternFill("solid", start_color="FFC7CE")
gris = PatternFill("solid", start_color="EDEDED")
editable = PatternFill("solid", start_color="FFF2CC")


def hoja(ws, cabeceras, filas, anchos, col_estado, colores, col_editable, pct=None):
    ws.append(cabeceras)
    for f in filas:
        ws.append([f[c] for c in cabeceras])
    for c in ws[1]:
        c.font, c.fill = cab_font, cab_fill
        c.alignment = Alignment(horizontal="center")
    ie = cabeceras.index(col_estado) + 1
    ied = cabeceras.index(col_editable) + 1
    ip = cabeceras.index(pct) + 1 if pct else None
    for fila in range(2, ws.max_row + 1):
        est = ws.cell(row=fila, column=ie).value
        ws.cell(row=fila, column=ie).fill = colores.get(est, amarillo)
        ws.cell(row=fila, column=ied).fill = editable
        if ip:
            ws.cell(row=fila, column=ip).number_format = "0%"
    for i, a in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


ws_rev = out.active
ws_rev.title = mc.HOJA_REVISION
hoja(ws_rev,
     ["ORIGINAL", "TIPO", "RESULTADO", "ESTADO", "MOTIVO", "APARICIONES", "FILAS_EXCEL", "CORRECTA"],
     filas_rev, [22, 19, 18, 14, 60, 13, 30, 18],
     "ESTADO", {mc.FIABLE: verde, mc.REVISAR: amarillo, mc.SIN_RESOLVER: rojo}, "CORRECTA")

ws_dic = out.create_sheet(mc.HOJA_CALCAS)
orden_dic = {"CALCA REPETIDA": 0, "EMPATE": 1, "REVISAR": 2, "POCOS DATOS": 3, "FIABLE": 4, "SIN DATOS": 5}
dic_ordenado = sorted(dic, key=lambda f: (orden_dic[f["ESTADO"]], f["MATRICULA"]))
hoja(ws_dic,
     ["MATRICULA", "CALCA", "APARICIONES", "FRECUENCIA", "CONFIANZA", "OTRAS_CALCAS", "ESTADO", "CALCA_CORRECTA"],
     dic_ordenado, [13, 9, 13, 12, 12, 40, 16, 16],
     "ESTADO", {"FIABLE": verde, "POCOS DATOS": amarillo, "REVISAR": amarillo, "EMPATE": rojo,
                "CALCA REPETIDA": rojo, "SIN DATOS": gris}, "CALCA_CORRECTA", pct="CONFIANZA")

out.save(mc.REVISION)

# ── 4) resumen por consola ──────────────────────────────────
total = sum(resumen.values())
print(f"\n   Celdas leídas: {total}")
for t, n in resumen.items():
    print(f"     {t:<18} {n}")
print(f"\n   Ya correctas (sin cambios):  {sin_cambio}")
print(f"   🟩 Se arreglan solas:          {estados[mc.FIABLE] - sin_cambio}")
print(f"   🟨 Revisar:                    {estados[mc.REVISAR]}")
print(f"   🟥 Sin resolver:               {estados[mc.SIN_RESOLVER]}")

n_calcas = sum(1 for f in dic if f["CALCA"])
fiables = sum(1 for f in dic if f["ESTADO"] == "FIABLE")
print(f"\n   Diccionario de calcas: {n_calcas} de {len(dic)} matrículas con calca ({fiables} fiables)")
for f in dic_ordenado:
    if f["ESTADO"] in ("CALCA REPETIDA", "EMPATE", "REVISAR"):
        print(f"     ⚠️ {f['MATRICULA']} → {f['CALCA']}  [{f['ESTADO']}]  {f['OTRAS_CALCAS']}")

pendientes = [f for f in filas_rev if f["ESTADO"] != mc.FIABLE]
if pendientes:
    print("\n   A revisar:")
    for f in pendientes:
        print(f"     {f['ORIGINAL'][:20]:<20} → {f['RESULTADO']:<14} [{f['ESTADO']}] {f['MOTIVO']}  x{f['APARICIONES']}")

print("\n✅ Revisión creada:", mc.REVISION)
print("   👉 Mira la hoja Diccionario_Calcas y la hoja Revision.")
print("   👉 Corrige lo que haga falta en CALCA_CORRECTA / CORRECTA y ejecuta 02_corregir.py")
