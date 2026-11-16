"""
Lógica compartida de los scripts de Mistery.

Columna MATRICULA_CALCA. En cada celda puede haber:
    '0791FTR'        -> solo matrícula
    '1234'           -> solo calca
    '0791FTR 1234'   -> las dos
El resultado final siempre es '0791FTR 1234' (matrícula + espacio + calca).

Las matrículas se corrigen con el MISMO diccionario que ControlParada
(scripts/ControlParada/comun.py -> MATRICULAS).
Las calcas no tienen diccionario: se construye uno a partir de las celdas
que traen las dos cosas (la calca que más se repite para cada matrícula).
"""

import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

# ── reutilizamos el diccionario y el buscador de ControlParada ──
_SCRIPTS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_SCRIPTS, "ControlParada"))
import comun as cp  # noqa: E402

DICCIONARIO = cp.DICCIONARIO
MATRICULAS = cp.MATRICULAS
DATA_DIR, DATA_RAW, DATA_PROCESSED = cp.DATA_DIR, cp.DATA_RAW, cp.DATA_PROCESSED

# ─────────────────────────────────────────────────────────────
# 📁 RUTAS
# ─────────────────────────────────────────────────────────────
NOMBRE_ENTRADA = "Mistery"
COLUMNA = "MATRICULA_CALCA"

REVISION = os.path.join(DATA_PROCESSED, "mistery_revision.xlsx")
SALIDA = os.path.join(DATA_PROCESSED, f"{NOMBRE_ENTRADA}_corregido.xlsx")
LOG_CAMBIOS = os.path.join(DATA_PROCESSED, "mistery_cambios.xlsx")

HOJA_CALCAS = "Diccionario_Calcas"
HOJA_REVISION = "Revision"

# Para dar por buena la calca de una matrícula:
MIN_APARICIONES = 2    # al menos 2 celdas con matrícula + calca
MIN_CONFIANZA = 0.6    # y que la calca ganadora sea >= 60 % de ellas

FIABLE, REVISAR, SIN_RESOLVER = "FIABLE", "REVISAR", "SIN RESOLVER"
ORDEN_ESTADO = {SIN_RESOLVER: 0, REVISAR: 1, FIABLE: 2}


# ─────────────────────────────────────────────────────────────
# 🔤 LEER UNA CELDA  ->  (matrícula_leída, calca_leída)
# ─────────────────────────────────────────────────────────────
def clave(valor):
    """Texto de la celda en mayúsculas y con espacios simples (sirve de id)."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    t = unicodedata.normalize("NFKD", str(valor))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t.upper()).strip()


def normalizar_calca(texto):
    nums = "".join(re.findall(r"\d", str(texto or "")))
    if not nums:
        return ""
    return nums.zfill(4) if len(nums) < 4 else nums


def separar(valor):
    """
    Separa la celda en matrícula y calca.
      '0791FTR 1234'   -> ('0791FTR', '1234')
      '0791 ftr-1234'  -> ('0791FTR', '1234')
      '0791FTR1234'    -> ('0791FTR', '1234')
      '1234 0791FTR'   -> ('0791FTR', '1234')
      '1234'           -> ('',        '1234')
      '0791FTR'        -> ('0791FTR', '')
    """
    t = clave(valor)
    if not t:
        return "", ""
    tokens = re.sub(r"[^A-Z0-9]", " ", t).split()
    if not tokens:
        return "", ""

    # une '0791' + 'FTR' (matrícula escrita con espacio)
    unidos = []
    i = 0
    while i < len(tokens):
        tk = tokens[i]
        if tk.isdigit() and i + 1 < len(tokens) and tokens[i + 1].isalpha():
            unidos.append(tk + tokens[i + 1])
            i += 2
        else:
            unidos.append(tk)
            i += 1

    con_letras = [x for x in unidos if re.search(r"[A-Z]", x)]
    solo_num = [x for x in unidos if x.isdigit()]

    # todo junto en un solo bloque: '0791FTR1234'
    if len(con_letras) == 1 and not solo_num:
        m = re.match(r"^(\d*[A-Z]+)(\d{3,5})$", con_letras[0])
        if m and len(re.sub(r"\D", "", m.group(1))) >= 3:
            return m.group(1), m.group(2)

    mat = "".join(con_letras)
    calca = solo_num[0] if solo_num else ""
    # solo números: si hay varios bloques, el resto se ignora (sale a revisión)
    return mat, calca


# ─────────────────────────────────────────────────────────────
# 🔧 CORREGIR MATRÍCULA (diccionario de ControlParada)
# ─────────────────────────────────────────────────────────────
_cache_mat = {}


def corregir_matricula(mat_leida):
    """-> (matricula_buena, similitud, alternativas, estado)"""
    if mat_leida in _cache_mat:
        return _cache_mat[mat_leida]
    norm = cp.normalizar(mat_leida)
    if norm in DICCIONARIO:
        res = (norm, 1.0, [], FIABLE)
    else:
        sug, sim, alts, est = cp.sugerir(mat_leida)
        estado = {"FIABLE": FIABLE, "REVISAR": REVISAR}.get(est, SIN_RESOLVER)
        res = (sug, sim, alts, estado)
    _cache_mat[mat_leida] = res
    return res


# ─────────────────────────────────────────────────────────────
# 📖 DICCIONARIO DE CALCAS (se construye con los datos)
# ─────────────────────────────────────────────────────────────
def construir_dic_calcas(valores):
    """
    valores: lista de textos de la columna.
    Cuenta, para cada matrícula (ya corregida con fiabilidad), qué calcas
    aparecen a su lado. La más repetida es la calca de esa matrícula.
    Devuelve lista de dicts (una fila por matrícula del diccionario).
    """
    conteo = defaultdict(Counter)
    for v in valores:
        mat_l, calca_l = separar(v)
        calca = normalizar_calca(calca_l)
        if not mat_l or not calca:
            continue
        mat, _, _, est = corregir_matricula(mat_l)
        if est == FIABLE and mat:
            conteo[mat][calca] += 1

    filas = []
    for mat in MATRICULAS:
        c = conteo.get(mat)
        if not c:
            filas.append({"MATRICULA": mat, "CALCA": "", "APARICIONES": 0,
                          "FRECUENCIA": 0, "CONFIANZA": 0.0, "OTRAS_CALCAS": "",
                          "ESTADO": "SIN DATOS", "CALCA_CORRECTA": ""})
            continue
        (calca, freq), total = c.most_common(1)[0], sum(c.values())
        empate = len(c) > 1 and c.most_common(2)[1][1] == freq
        conf = round(freq / total, 2)
        otras = ", ".join(f"{k} (x{n})" for k, n in c.most_common()[1:])
        if empate:
            estado = "EMPATE"
        elif total < MIN_APARICIONES:
            estado = "POCOS DATOS"
        elif conf < MIN_CONFIANZA:
            estado = "REVISAR"
        else:
            estado = "FIABLE"
        filas.append({"MATRICULA": mat, "CALCA": calca, "APARICIONES": total,
                      "FRECUENCIA": freq, "CONFIANZA": conf, "OTRAS_CALCAS": otras,
                      "ESTADO": estado, "CALCA_CORRECTA": ""})

    # una calca no puede ser de dos matrículas a la vez
    por_calca = defaultdict(list)
    for f in filas:
        if f["CALCA"]:
            por_calca[f["CALCA"]].append(f["MATRICULA"])
    for f in filas:
        dup = [m for m in por_calca.get(f["CALCA"], []) if m != f["MATRICULA"]]
        if dup:
            f["ESTADO"] = "CALCA REPETIDA"
            f["OTRAS_CALCAS"] = (f["OTRAS_CALCAS"] + "  |  también en " + ", ".join(dup)).strip(" |")
    return filas


def mapas_calcas(filas):
    """
    A partir del diccionario de calcas (con CALCA_CORRECTA si la has puesto)
    devuelve:
      mat_calca : matrícula -> (calca, fiable?)
      calca_mat : calca -> matrícula   (solo calcas que no están repetidas)
    """
    mat_calca = {}
    for f in filas:
        manual = normalizar_calca(f.get("CALCA_CORRECTA"))
        calca = manual or normalizar_calca(f.get("CALCA"))
        if not calca:
            continue
        fiable = bool(manual) or f.get("ESTADO") == "FIABLE"
        mat_calca[f["MATRICULA"]] = (calca, fiable)

    por_calca = defaultdict(list)
    for m, (c, _) in mat_calca.items():
        por_calca[c].append(m)
    calca_mat = {c: ms[0] for c, ms in por_calca.items() if len(ms) == 1}
    return mat_calca, calca_mat


def _calca_mas_parecida(calca, calca_mat):
    puntos = sorted(((cp.similitud(calca, c), c) for c in calca_mat), reverse=True)
    if not puntos:
        return None, 0.0, False
    mejor_s, mejor = puntos[0]
    empate = len(puntos) > 1 and mejor_s - puntos[1][0] <= 0.05
    return mejor, mejor_s, empate


# ─────────────────────────────────────────────────────────────
# 🧩 RESOLVER UNA CELDA
# ─────────────────────────────────────────────────────────────
def resolver(valor, mat_calca, calca_mat):
    """
    Devuelve dict con: TIPO, MATRICULA_LEIDA, CALCA_LEIDA,
    MATRICULA, CALCA, RESULTADO, ESTADO, MOTIVO
    """
    mat_l, calca_l = separar(valor)
    calca = normalizar_calca(calca_l)
    r = {"MATRICULA_LEIDA": mat_l, "CALCA_LEIDA": calca_l,
         "MATRICULA": "", "CALCA": "", "ESTADO": FIABLE, "MOTIVO": ""}
    motivos = []

    def peor(e):
        if ORDEN_ESTADO[e] < ORDEN_ESTADO[r["ESTADO"]]:
            r["ESTADO"] = e

    if mat_l and calca:
        r["TIPO"] = "MATRICULA + CALCA"
    elif mat_l:
        r["TIPO"] = "SOLO MATRICULA"
    elif calca:
        r["TIPO"] = "SOLO CALCA"
    else:
        r["TIPO"] = "ILEGIBLE"

    # 1) matrícula
    mat = ""
    if mat_l:
        mat, sim, alts, est = corregir_matricula(mat_l)
        if not mat:
            motivos.append(f"matrícula {mat_l} no se parece a ninguna")
            # si hay calca, se intenta sacar la matrícula por la calca
            peor(REVISAR if calca else SIN_RESOLVER)
        elif mat != cp.normalizar(mat_l):
            txt = f"matrícula {mat_l}→{mat} ({sim:.0%})"
            if alts:
                txt += f" o {', '.join(alts)}"
            motivos.append(txt)
            peor(est)

    # 2) calca
    if mat and calca:
        esperada, fiable = mat_calca.get(mat, ("", False))
        if not esperada or calca == esperada:
            r["MATRICULA"], r["CALCA"] = mat, calca
            if not esperada:
                motivos.append("matrícula sin calca en el diccionario")
                peor(REVISAR)
        else:
            otra = calca_mat.get(calca)
            mat_era_buena = cp.normalizar(mat_l) in DICCIONARIO
            if otra and not mat_era_buena and cp.similitud(cp.limpiar(mat_l), otra) >= 0.5:
                # la calca es de otra matrícula parecida: la matrícula estaba mal
                r["MATRICULA"], r["CALCA"] = otra, calca
                motivos.append(f"la calca {calca} es de {otra}: se usa esa matrícula")
                peor(REVISAR)
            else:
                r["MATRICULA"], r["CALCA"] = mat, esperada
                motivos.append(f"calca {calca}→{esperada}")
                if otra:
                    motivos.append(f"({calca} es la calca de {otra})")
                    peor(REVISAR)
                elif not fiable:
                    peor(REVISAR)
                elif cp.similitud(calca, esperada) < 0.75:
                    peor(REVISAR)

    elif mat:  # solo matrícula
        esperada, fiable = mat_calca.get(mat, ("", False))
        r["MATRICULA"], r["CALCA"] = mat, esperada
        if esperada:
            motivos.append(f"añadida calca {esperada}")
            if not fiable:
                peor(REVISAR)
        else:
            motivos.append("sin calca conocida para esta matrícula")
            peor(SIN_RESOLVER)

    elif calca:  # solo calca
        otra = calca_mat.get(calca)
        if otra:
            r["MATRICULA"], r["CALCA"] = otra, calca
            motivos.append(f"añadida matrícula {otra}")
            if not mat_calca[otra][1]:
                peor(REVISAR)
        else:
            parecida, sim, empate = _calca_mas_parecida(calca, calca_mat)
            if parecida and sim >= 0.75 and not empate:
                otra = calca_mat[parecida]
                r["MATRICULA"], r["CALCA"] = otra, parecida
                motivos.append(f"calca {calca}→{parecida} ({sim:.0%}), matrícula {otra}")
                peor(REVISAR)
            else:
                r["CALCA"] = calca
                motivos.append(f"calca {calca} no corresponde a ninguna matrícula")
                peor(SIN_RESOLVER)
    else:
        if not motivos:
            motivos.append("no se puede leer")
        peor(SIN_RESOLVER)

    if r["TIPO"] == "MATRICULA + CALCA" and not r["MATRICULA"]:
        r["CALCA"] = calca

    r["RESULTADO"] = " ".join(x for x in (r["MATRICULA"], r["CALCA"]) if x) or clave(valor)
    r["MOTIVO"] = "; ".join(motivos)
    return r


# ─────────────────────────────────────────────────────────────
# 📄 EXCEL DE ENTRADA
# ─────────────────────────────────────────────────────────────
def buscar_entrada():
    for carpeta in (DATA_RAW, DATA_DIR):
        if not os.path.isdir(carpeta):
            continue
        for f in sorted(os.listdir(carpeta)):
            nombre, ext = os.path.splitext(f)
            if (nombre.lower() == NOMBRE_ENTRADA.lower()
                    and ext.lower() in (".xlsx", ".xlsm")
                    and not f.startswith("~$")):
                return os.path.join(carpeta, f)
    raise FileNotFoundError(f"No encuentro '{NOMBRE_ENTRADA}.xlsx'. Ponlo en {DATA_RAW}")


def recorrer_celdas(wb):
    """(hoja, celda) de cada valor no vacío de la columna MATRICULA_CALCA."""
    encontrada = False
    for ws in wb.worksheets:
        pos = cp.buscar_columna(ws, COLUMNA)
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
