import pandas as pd
import re
import os
from collections import Counter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ARCHIVO = os.path.join(DATA_RAW, "Aforo_en_Parada_Madrid.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "diccionario.xlsx")

COLUMNA = "MATRICULA"

print("📥 Generando diccionario...")

df = pd.read_excel(ARCHIVO)

df[COLUMNA] = df[COLUMNA].fillna("").astype(str).str.upper().str.replace(" ", "")

def separar(m):
    nums = "".join(re.findall(r"\d", m))
    letras = "".join(re.findall(r"[A-Z]", m))
    return nums, letras

df[["NUM", "LET"]] = df[COLUMNA].apply(lambda x: pd.Series(separar(x)))

resultado = []

for num in df["NUM"].unique():
    subset = df[df["NUM"] == num]
    conteo = Counter(subset["LET"])
    
    if conteo:
        letras, freq = conteo.most_common(1)[0]
        total = len(subset)

        resultado.append({
            "NUMERO": num,
            "SUGERENCIA": num + letras,
            "FRECUENCIA": freq,
            "TOTAL": total,
            "CONFIANZA": round(freq / total, 2)
        })

pd.DataFrame(resultado).to_excel(SALIDA, index=False)

print("✅ Diccionario creado:", SALIDA)