"""
Funciones y configuración compartidas por los scripts de ControlParada.

Aquí está:
  - las rutas (data/raw, data/processed) siguiendo el mismo flujo que Matriculas
  - el DICCIONARIO de matrículas buenas (config/matriculas.txt)
  - la normalización al formato 0000XXX
  - el buscador de la corrección más probable para una matrícula mal escrita
"""

import os
import re
import unicodedata

# ─────────────────────────────────────────────────────────────
# 📁 RUTAS (mismo esquema que scripts/Matriculas)
# ─────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_RAW = os.path.join(DATA_DIR, "raw")
DATA_PROCESSED = os.path.join(DATA_DIR, "processed")

NOMBRE_ENTRADA = "Matricules_Control_Parada"

REVISION = os.path.join(DATA_PROCESSED, "control_parada_revision.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, f"{NOMBRE_ENTRADA}_corregido.xlsx")
LOG_CAMBIOS = os.path.join(DATA_PROCESSED, "control_parada_cambios.xlsx")

COLUMNA = "MATRICULA"

# Por debajo de esta similitud la sugerencia se marca para revisar a mano
UMBRAL_FIABLE = 0.75

PATRON_OK = re.compile(r"^\d{4}[A-Z]{3}$")

# ─────────────────────────────────────────────────────────────
# 📖 DICCIONARIO DE MATRÍCULAS BUENAS
#    Las matrículas del diccionario son SIEMPRE las correctas.
#    Si una matrícula no está aquí, se considera mal escrita.
#
#    Se leen de config/matriculas.txt (una por línea, '#' = comentario).
#    Ese archivo NO se sube a git: cada uno tiene la flota de su cliente.
#    Si no existe, se usa config/matriculas.example.txt (matrículas ficticias).
# ─────────────────────────────────────────────────────────────
CONFIG_DIR = os.path.join(BASE_DIR, "config")
ARCHIVO_MATRICULAS = os.path.join(CONFIG_DIR, "matriculas.txt")
ARCHIVO_MATRICULAS_EJEMPLO = os.path.join(CONFIG_DIR, "matriculas.example.txt")


def cargar_matriculas(ruta=None):
    """Lee el diccionario de matrículas (sin duplicados, en el orden del archivo)."""
    if ruta is None:
        ruta = os.environ.get("MATRICULAS_FILE")  # permite forzar otro archivo (tests)
    if ruta is None:
        ruta = ARCHIVO_MATRICULAS if os.path.exists(ARCHIVO_MATRICULAS) else ARCHIVO_MATRICULAS_EJEMPLO
        if ruta == ARCHIVO_MATRICULAS_EJEMPLO:
            print("⚠️  No existe config/matriculas.txt: usando el diccionario de EJEMPLO.")
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No encuentro el diccionario de matrículas: {ruta}")

    matriculas = []
    with open(ruta, encoding="utf-8") as f:
        for n, linea in enumerate(f, start=1):
            linea = linea.split("#", 1)[0].strip()
            if not linea:
                continue
            for trozo in re.split(r"[,;]", linea):
                junto = re.sub(r"[^A-Z0-9]", "", trozo.upper())
                if not junto:
                    continue
                # '0791 FTR' es una matrícula; '0791FTR 1186FTR' son dos
                partes = [junto] if PATRON_OK.match(junto) else \
                    [re.sub(r"[^A-Z0-9]", "", p.upper()) for p in trozo.split()]
                for m in partes:
                    if not PATRON_OK.match(m):
                        raise ValueError(f"{ruta}:{n} -> '{trozo.strip()}' no tiene formato 0000ABC")
                    if m not in matriculas:
                        matriculas.append(m)
    if not matriculas:
        raise ValueError(f"El diccionario {ruta} está vacío")
    return matriculas


MATRICULAS = cargar_matriculas()
DICCIONARIO = set(MATRICULAS)


# ─────────────────────────────────────────────────────────────
# 🔤 NORMALIZACIÓN  ->  4 números + 3 letras, mayúsculas, sin espacios
# ─────────────────────────────────────────────────────────────
def limpiar(valor):
    """Mayúsculas y solo letras/números (quita espacios, guiones, puntos...)."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", texto.upper())


def normalizar(valor):
    """
    Pasa la matrícula al formato 0000XXX.
      '0791 ftr' -> '0791FTR'     'FTR-0791' -> '0791FTR'
      '791FTR'   -> '0791FTR'     (rellena con ceros a la izquierda)
    Si no se puede dejar en ese formato, devuelve lo que haya (sin espacios)
    para que salga en la revisión.
    """
    t = limpiar(valor)
    if not t:
        return ""
    if PATRON_OK.match(t):
        return t
    # letras delante: FTR0791 -> 0791FTR
    m = re.match(r"^([A-Z]{3})(\d{1,4})$", t)
    if m:
        return m.group(2).zfill(4) + m.group(1)
    # números cortos: 791FTR -> 0791FTR
    m = re.match(r"^(\d{1,3})([A-Z]{3})$", t)
    if m:
        return m.group(1).zfill(4) + m.group(2)
    # cualquier otra cosa: números delante y letras detrás
    nums = "".join(re.findall(r"\d", t))
    letras = "".join(re.findall(r"[A-Z]", t))
    if len(nums) < 4 and nums:
        nums = nums.zfill(4)
    return nums + letras


# ─────────────────────────────────────────────────────────────
# 🔍 BÚSQUEDA DE LA CORRECCIÓN MÁS PROBABLE
#    Distancia de edición (como Levenshtein) donde los errores típicos
#    de transcripción cuestan menos: O↔0, S↔5, B↔8, Z↔2, I/L↔1, G↔6...
#    y también el intercambio de dos caracteres seguidos (1850 ↔ 1580).
# ─────────────────────────────────────────────────────────────
_PARES_CONFUSOS = [
    ("0", "O"), ("0", "D"), ("0", "Q"), ("1", "I"), ("1", "L"), ("1", "7"),
    ("2", "Z"), ("5", "S"), ("8", "B"), ("6", "G"), ("7", "T"), ("4", "A"),
    ("3", "8"), ("6", "8"), ("9", "8"), ("0", "8"), ("6", "5"), ("3", "5"),
    ("M", "N"), ("V", "Y"), ("V", "W"), ("K", "X"), ("C", "G"), ("H", "N"),
    ("J", "I"), ("P", "R"), ("B", "R"),
]
CONFUSOS = {frozenset(p) for p in _PARES_CONFUSOS}

COSTE_CONFUSO = 0.4
COSTE_TRANSPOSICION = 0.6


def _coste(a, b):
    if a == b:
        return 0.0
    if frozenset((a, b)) in CONFUSOS:
        return COSTE_CONFUSO
    return 1.0


def distancia(a, b):
    n, m = len(a), len(b)
    d = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        d[i][0] = float(i)
    for j in range(m + 1):
        d[0][j] = float(j)
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            d[i][j] = min(
                d[i - 1][j] + 1,                               # sobra un carácter
                d[i][j - 1] + 1,                               # falta un carácter
                d[i - 1][j - 1] + _coste(a[i - 1], b[j - 1]),  # carácter cambiado
            )
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + COSTE_TRANSPOSICION)
    return d[n][m]


def similitud(a, b):
    if not a or not b:
        return 0.0
    return max(0.0, 1 - distancia(a, b) / max(len(a), len(b)))


def sugerir(valor_original):
    """
    Devuelve (sugerencia, similitud, alternativas, estado) para una matrícula
    que NO está en el diccionario.
      estado = 'FIABLE'  -> una sola candidata clara
               'REVISAR' -> similitud baja o empate entre varias candidatas
               'SIN SUGERENCIA' -> no se parece a nada
    Compara tanto la versión normalizada como la original sin espacios,
    y se queda con la mejor (así 'O791FTR' encuentra '0791FTR').
    """
    formas = {normalizar(valor_original), limpiar(valor_original)} - {""}
    if not formas:
        return "", 0.0, [], "SIN SUGERENCIA"

    puntuaciones = []
    for buena in MATRICULAS:
        s = max(similitud(f, buena) for f in formas)
        puntuaciones.append((s, buena))
    puntuaciones.sort(key=lambda x: (-x[0], x[1]))

    mejor_s, mejor = puntuaciones[0]
    if mejor_s < 0.5:
        return "", round(mejor_s, 2), [], "SIN SUGERENCIA"

    # candidatas casi igual de buenas = posible ambigüedad
    alternativas = [b for s, b in puntuaciones[1:4] if mejor_s - s <= 0.05]

    estado = "FIABLE" if (mejor_s >= UMBRAL_FIABLE and not alternativas) else "REVISAR"
    return mejor, round(mejor_s, 2), alternativas, estado


# ─────────────────────────────────────────────────────────────
# 📄 LOCALIZAR EL EXCEL DE ENTRADA Y LA COLUMNA MATRICULA
# ─────────────────────────────────────────────────────────────
def buscar_entrada():
    """Busca Matricules_Control_Parada.xlsx en data/raw (o en data/)."""
    for carpeta in (DATA_RAW, DATA_DIR):
        if not os.path.isdir(carpeta):
            continue
        for f in sorted(os.listdir(carpeta)):
            nombre, ext = os.path.splitext(f)
            if (nombre.lower() == NOMBRE_ENTRADA.lower()
                    and ext.lower() in (".xlsx", ".xlsm")
                    and not f.startswith("~$")):
                return os.path.join(carpeta, f)
    raise FileNotFoundError(
        f"No encuentro '{NOMBRE_ENTRADA}.xlsx'. Ponlo en {DATA_RAW}"
    )


def _texto_cabecera(v):
    t = unicodedata.normalize("NFKD", str(v or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.strip().upper()


def buscar_columna(ws, nombre=COLUMNA, max_filas=20):
    """
    Devuelve (fila_cabecera, columna) donde está la cabecera MATRICULA
    (acepta 'Matrícula', 'matricula', etc.). None si la hoja no la tiene.
    """
    objetivo = _texto_cabecera(nombre)
    for fila in ws.iter_rows(min_row=1, max_row=min(max_filas, ws.max_row)):
        for celda in fila:
            if _texto_cabecera(celda.value) == objetivo:
                return celda.row, celda.column
    return None


def recorrer_matriculas(wb):
    """
    Genera (hoja, celda) para cada matrícula no vacía de todas las hojas
    que tengan la columna MATRICULA.
    """
    encontrada = False
    for ws in wb.worksheets:
        pos = buscar_columna(ws)
        if not pos:
            continue
        encontrada = True
        fila_cab, col = pos
        for fila in range(fila_cab + 1, ws.max_row + 1):
            celda = ws.cell(row=fila, column=col)
            if celda.value is None or str(celda.value).strip() == "":
                continue
            yield ws, celda
    if not encontrada:
        raise ValueError(f"Ninguna hoja del Excel tiene la columna '{COLUMNA}'")
