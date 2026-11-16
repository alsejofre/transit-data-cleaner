import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

ARCHIVO = os.path.join(DATA_RAW, "Encuestas_Guadalajara.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "lugares_unicos.xlsx")

COLUMNAS = ["ORIGEN", "DESTINO"]

df = pd.read_excel(ARCHIVO)

# juntar columnas
lugares = pd.concat([df["ORIGEN"], df["DESTINO"]])

# limpiar básico
lugares = (
    lugares.fillna("")
    .astype(str)
    .str.upper()
    .str.strip()
)

lugares = lugares[lugares != ""].drop_duplicates().sort_values()

df_out = pd.DataFrame({"LUGAR": lugares})

df_out.to_excel(SALIDA, index=False)

print("✅ Lugares únicos:", SALIDA)