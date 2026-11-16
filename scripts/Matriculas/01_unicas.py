import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ARCHIVO = os.path.join(DATA_RAW, "Aforo_en_Parada_Madrid.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "matriculas_unicas.xlsx")

COLUMNA = "MATRICULA"

print("📥 Generando matrículas únicas...")

df = pd.read_excel(ARCHIVO)

df[COLUMNA] = df[COLUMNA].fillna("").astype(str).str.upper().str.replace(" ", "")
df = df[df[COLUMNA] != ""]

unicas = df[[COLUMNA]].drop_duplicates().sort_values(by=COLUMNA)

unicas.to_excel(SALIDA, index=False)

print("✅ OK:", SALIDA)