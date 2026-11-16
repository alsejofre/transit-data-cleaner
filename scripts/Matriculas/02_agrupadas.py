import pandas as pd
import os
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_PROCESSED = os.path.join(BASE_DIR, "data", "processed")

UNICAS = os.path.join(DATA_PROCESSED, "matriculas_unicas.xlsx")
DICCIONARIO = os.path.join(DATA_PROCESSED, "diccionario.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, "matriculas_agrupadas.xlsx")

COLUMNA = "MATRICULA"
UMBRAL_CONFIANZA = 0.6  # por debajo de esto no nos fiamos del diccionario para autocorregir

print("📊 Generando agrupadas...")

df = pd.read_excel(UNICAS)

# 🔒 IMPORTANTE: forzamos NUMERO como texto al leer.
# Si no, pandas detecta la columna como "todo números" y borra los ceros
# de la izquierda (ej. "0223" -> 223.0), y entonces NINGUNA matrícula
# vuelve a coincidir con el diccionario. Este era el bug que hacía que
# la sugerencia saliera siempre vacía.
dic = pd.read_excel(DICCIONARIO, dtype={"NUMERO": str})

df[COLUMNA] = df[COLUMNA].fillna("").astype(str).str.upper().str.replace(" ", "")

# solo usamos como candidatas las entradas fiables del diccionario
dic_fiable = dic.dropna(subset=["CONFIANZA"])
dic_fiable = dic_fiable[dic_fiable["CONFIANZA"] > UMBRAL_CONFIANZA]

mapa = dict(zip(dic_fiable["NUMERO"].astype(str), dic_fiable["SUGERENCIA"]))
confianzas = dict(zip(dic_fiable["NUMERO"].astype(str), dic_fiable["CONFIANZA"]))


def extraer_num(m):
    return "".join(re.findall(r"\d", m))


def mejor_coincidencia(num):
    """
    Busca en el diccionario el NUMERO que más se parece al extraído,
    dígito a dígito (tolera que el inspector se haya equivocado en
    un dígito, o que falte/sobre alguno). Esto es lo que hacían los
    scripts originales y es lo que permitía corregir automáticamente
    la mayoría de las matrículas sin revisión manual.
    """
    mejor_num, mejor_score = None, -1
    for n in mapa.keys():
        num2 = num.zfill(len(n))
        score = sum(a == b for a, b in zip(n, num2))
        if score > mejor_score:
            mejor_score, mejor_num = score, n
    return mejor_num, mejor_score


df["NUMERO"] = df[COLUMNA].apply(extraer_num)

sugerencias, confianza_match = [], []
for num in df["NUMERO"]:
    if not num:
        # matrícula sin ningún dígito legible (ej. ".", "XXXXXX"):
        # no tiene sentido buscar coincidencia, va directa a revisión
        sugerencias.append("")
        confianza_match.append(0)
        continue

    mejor_num, _ = mejor_coincidencia(num)
    sugerencias.append(mapa.get(mejor_num, ""))
    confianza_match.append(confianzas.get(mejor_num, 0))

df["SUGERENCIA"] = sugerencias
df["CONFIANZA"] = confianza_match
df["CORRECTA"] = ""  # ← opcional: rellena esto a mano solo si quieres forzar otra corrección distinta a la sugerida

df = df.sort_values(by=["NUMERO", COLUMNA])

df.to_excel(SALIDA, index=False)

print("✅ Agrupadas listas:", SALIDA)
