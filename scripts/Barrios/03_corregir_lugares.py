import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ORIGINAL = os.path.join(DATA_RAW, "Encuestas_Guadalajara.xlsx")
MAPA = os.path.join(DATA_PROCESSED, "lugares_agrupados.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "viajes_corregidos.xlsx")

df = pd.read_excel(ORIGINAL)
df_map = pd.read_excel(MAPA)

# limpiar
df["ORIGEN"] = df["ORIGEN"].fillna("").astype(str).str.upper().str.strip()
df["DESTINO"] = df["DESTINO"].fillna("").astype(str).str.upper().str.strip()

# fillna("") antes de astype(str): si no, las celdas vacías se convierten en "NAN"
# y TODOS los lugares acabarían "corregidos" a "NAN"
df_map["LUGAR"] = df_map["LUGAR"].fillna("").astype(str).str.upper().str.strip()
df_map["CORRECTO"] = df_map["CORRECTO"].fillna("").astype(str).str.upper().str.strip()

# usar solo los que has definido
df_map = df_map[df_map["CORRECTO"] != ""]

mapa = dict(zip(df_map["LUGAR"], df_map["CORRECTO"]))

def corregir(x):
    return mapa.get(x, x)

df["ORIGEN_CORREGIDO"] = df["ORIGEN"].apply(corregir)
df["DESTINO_CORREGIDO"] = df["DESTINO"].apply(corregir)

df.to_excel(SALIDA, index=False)

print("✅ Resultado:", SALIDA)