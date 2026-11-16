import pandas as pd
import os
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ARCHIVO = os.path.join(DATA_PROCESSED, "lugares_unicos.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "lugares_agrupados.xlsx")

df = pd.read_excel(ARCHIVO)

# 🔥 limpiar fuerte
def normalizar(texto):
    texto = str(texto).upper()
    texto = re.sub(r"\bEL\b|\bLA\b|\bLOS\b|\bLAS\b", "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto

df["NORMALIZADO"] = df["LUGAR"].apply(normalizar)

# agrupar por normalizado
df["GRUPO"] = df["NORMALIZADO"]

# sugerencia automática (la versión más corta)
df["SUGERENCIA"] = df.groupby("GRUPO")["LUGAR"].transform(lambda x: min(x, key=len))

df["CORRECTO"] = ""

df = df.sort_values(by=["GRUPO", "LUGAR"])

df.to_excel(SALIDA, index=False)

print("✅ Agrupados:", SALIDA)