"""
PASO 2 · Aplicar las correcciones y generar el Excel Mistery corregido.

Entrada : data/raw/Mistery.xlsx
          data/processed/mistery_revision.xlsx  (del paso 1, con tus cambios)
Salida  : data/processed/Mistery_corregido.xlsx
          data/processed/mistery_cambios.xlsx   (registro de cada cambio)

Cada celda de MATRICULA_CALCA queda como '0000ABC 0000'.
  - Usa el Diccionario_Calcas (con CALCA_CORRECTA si la has rellenado).
  - Si en la hoja Revision has puesto algo en CORRECTA, eso manda.
  - El resto del Excel no se toca.
"""

import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import mistery_comun as mc

print("⚙️ Aplicando correcciones Mistery...")

ENTRADA = mc.buscar_entrada()
if not os.path.exists(mc.REVISION):
    raise FileNotFoundError("Falta la revisión. Ejecuta antes 01_detectar.py")

rev_wb = load_workbook(mc.REVISION, data_only=True)


def leer_hoja(ws):
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    cab = [str(c).strip().upper() if c else "" for c in filas[0]]
    return [dict(zip(cab, f)) for f in filas[1:]]


# ── diccionario de calcas (con tus correcciones) ────────────
dic = leer_hoja(rev_wb[mc.HOJA_CALCAS])
for f in dic:
    f["MATRICULA"] = mc.cp.limpiar(f.get("MATRICULA"))
    f["ESTADO"] = f.get("ESTADO") or ""
mat_calca, calca_mat = mc.mapas_calcas(dic)

# ── correcciones manuales por valor ─────────────────────────
manuales = {}
for f in leer_hoja(rev_wb[mc.HOJA_REVISION]):
    correcta = mc.clave(f.get("CORRECTA"))
    if not correcta:
        continue
    mat_l, calca_l = mc.separar(correcta)
    mat = mc.cp.normalizar(mat_l) if mat_l else ""
    calca = mc.normalizar_calca(calca_l)
    manuales[mc.clave(f.get("ORIGINAL"))] = " ".join(x for x in (mat, calca) if x)

# ── corregir el Excel original (mantiene su formato) ────────
wb = load_workbook(ENTRADA)
cambios, pendientes = [], []
total = 0
cache = {}

for ws, celda in mc.recorrer_celdas(wb):
    total += 1
    k = mc.clave(celda.value)
    if k in manuales:
        nuevo, estado, motivo = manuales[k], "MANUAL", "corrección manual"
    else:
        if k not in cache:
            cache[k] = mc.resolver(k, mat_calca, calca_mat)
        r = cache[k]
        nuevo, estado, motivo = r["RESULTADO"], r["ESTADO"], r["MOTIVO"]
        if not r["MATRICULA"] and not r["CALCA"]:
            nuevo = celda.value  # ilegible: se deja tal cual
        if estado == mc.SIN_RESOLVER:
            pendientes.append(f"{ws.title}!{celda.row}: {celda.value}  ({motivo})")

    if nuevo != celda.value:
        cambios.append((ws.title, celda.row, str(celda.value), nuevo, estado, motivo))
        celda.value = nuevo

wb.save(mc.SALIDA)

# ── registro de cambios ─────────────────────────────────────
log = Workbook()
wl = log.active
wl.title = "Cambios"
wl.append(["HOJA", "FILA", "ANTES", "DESPUES", "ESTADO", "MOTIVO"])
for c in cambios:
    wl.append(list(c))
for c in wl[1]:
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="305496")
for i, a in enumerate([14, 8, 22, 18, 14, 60], start=1):
    wl.column_dimensions[get_column_letter(i)].width = a
wl.freeze_panes = "A2"
log.save(mc.LOG_CAMBIOS)

# ── resumen ─────────────────────────────────────────────────
print(f"\n   Celdas procesadas:  {total}")
print(f"   ✏️  Cambiadas:        {len(cambios)}")
print(f"   ✋ Manuales:          {sum(1 for c in cambios if c[4] == 'MANUAL')}")
print(f"   ⚠️  Sin resolver:     {len(pendientes)}")
for p in pendientes:
    print("      -", p)

print("\n✅ Excel corregido:", mc.SALIDA)
print("   Registro de cambios:", mc.LOG_CAMBIOS)
