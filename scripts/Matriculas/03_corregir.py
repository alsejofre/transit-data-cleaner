import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ORIGINAL = os.path.join(DATA_RAW, "Aforo_en_Parada_Madrid.xlsx")
MAPA = os.path.join(DATA_PROCESSED, "matriculas_agrupadas.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "resultado.xlsx")

COLUMNA = "MATRICULA"

print("⚙️ Corrigiendo datos...")

df = pd.read_excel(ORIGINAL)
df_map = pd.read_excel(MAPA, dtype=str)  # texto siempre, para no perder ceros ni mezclar tipos

df[COLUMNA] = df[COLUMNA].fillna("").astype(str).str.upper().str.replace(" ", "")
df_map[COLUMNA] = df_map[COLUMNA].fillna("").astype(str).str.upper().str.replace(" ", "")
df_map["SUGERENCIA"] = df_map["SUGERENCIA"].fillna("").astype(str).str.upper().str.replace(" ", "")
df_map["CORRECTA"] = df_map["CORRECTA"].fillna("").astype(str).str.upper().str.replace(" ", "")

# la corrección manual (CORRECTA), si existe, manda por encima de la sugerencia automática
df_map["FINAL"] = df_map["CORRECTA"].where(df_map["CORRECTA"] != "", df_map["SUGERENCIA"])

mapa = dict(zip(df_map[COLUMNA], df_map["FINAL"]))


def corregir(m):
    final = mapa.get(m, "")
    if final:
        return final, (m != final)
    return m, True  # no se encontró ninguna sugerencia fiable: queda marcada para revisión manual


res = df[COLUMNA].apply(corregir)

df["CORREGIDA"] = res.apply(lambda x: x[0])
df["REVISAR"] = res.apply(lambda x: x[1])

df.to_excel(SALIDA, index=False)

# 🎨 pintar rojo las filas que han cambiado o que no se han podido corregir
wb = load_workbook(SALIDA)
ws = wb.active

rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

col_idx = list(df.columns).index("CORREGIDA") + 1

for i, revisar in enumerate(df["REVISAR"], start=2):
    if revisar:
        ws.cell(row=i, column=col_idx).fill = rojo

wb.save(SALIDA)

print("✅ Resultado final:", SALIDA)
