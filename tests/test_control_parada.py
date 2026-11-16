import pytest

import comun as cp

BUENA = cp.MATRICULAS[0]          # p. ej. 5305FTR
NUM, LET = BUENA[:4], BUENA[4:]


# ── normalización al formato 0000ABC ───────────────────────
@pytest.mark.parametrize("entrada, esperado", [
    ("0791FTR", "0791FTR"),
    ("0791 ftr", "0791FTR"),
    ("  0791-FTR ", "0791FTR"),
    ("FTR0791", "0791FTR"),
    ("FTR 791", "0791FTR"),
    ("791FTR", "0791FTR"),
    ("", ""),
    (None, ""),
])
def test_normalizar(entrada, esperado):
    assert cp.normalizar(entrada) == esperado


def test_normalizar_numero_de_excel():
    # Excel a veces guarda 150 como 150.0
    assert cp.limpiar(150.0) == "150"


# ── distancia con errores típicos ──────────────────────────
def test_confusiones_tipicas_cuestan_menos_que_un_cambio_cualquiera():
    assert cp.distancia("0791FTR", "O791FTR") < cp.distancia("0791FTR", "9791FTR")


def test_transposicion_cuesta_menos_que_dos_cambios():
    assert cp.distancia("1542GKP", "5142GKP") < 2


# ── sugerencias contra el diccionario ──────────────────────
def test_letra_o_en_vez_de_cero():
    m = cp.MATRICULAS[[x[0] for x in cp.MATRICULAS].index("0")]  # una que empiece por 0
    sug, sim, _, estado = cp.sugerir("O" + m[1:])
    assert sug == m
    assert estado == "FIABLE"


def test_digitos_intercambiados():
    mal = NUM[1] + NUM[0] + NUM[2:] + LET
    sug, _, _, _ = cp.sugerir(mal)
    assert sug == BUENA


def test_letra_cambiada():
    mal = NUM + LET[:2] + ("X" if LET[2] != "X" else "Y")
    sug, _, _, _ = cp.sugerir(mal)
    assert sug == BUENA


def test_basura_no_tiene_sugerencia():
    sug, _, _, estado = cp.sugerir("XXXX")
    assert sug == ""
    assert estado == "SIN SUGERENCIA"


# ── carga del diccionario ──────────────────────────────────
def test_cargar_matriculas(tmp_path):
    f = tmp_path / "m.txt"
    f.write_text("# comentario\n0791 ftr, 1186FTR\n\n0791FTR  # repetida\n", encoding="utf-8")
    assert cp.cargar_matriculas(str(f)) == ["0791FTR", "1186FTR"]


def test_cargar_matriculas_formato_incorrecto(tmp_path):
    f = tmp_path / "m.txt"
    f.write_text("12345\n", encoding="utf-8")
    with pytest.raises(ValueError):
        cp.cargar_matriculas(str(f))
