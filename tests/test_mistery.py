import pytest

import mistery_comun as mc

M1, M2, M3 = mc.MATRICULAS[0], mc.MATRICULAS[1], mc.MATRICULAS[6]
C1, C2, C3 = "2201", "2202", "2207"


# ── separar matrícula y calca ──────────────────────────────
@pytest.mark.parametrize("entrada, esperado", [
    ("0791FTR 1234", ("0791FTR", "1234")),
    ("0791 ftr-1234", ("0791FTR", "1234")),
    ("0791FTR1234", ("0791FTR", "1234")),
    ("1234 0791FTR", ("0791FTR", "1234")),
    ("1234", ("", "1234")),
    (1234, ("", "1234")),
    ("0791FTR", ("0791FTR", "")),
    ("", ("", "")),
])
def test_separar(entrada, esperado):
    assert mc.separar(entrada) == esperado


def test_normalizar_calca():
    assert mc.normalizar_calca("220") == "0220"
    assert mc.normalizar_calca("A-2201") == "2201"
    assert mc.normalizar_calca(None) == ""


# ── diccionario de calcas construido con los datos ─────────
@pytest.fixture
def mapas():
    valores = (
        [f"{M1} {C1}"] * 3 + [f"{M1} 2209"]      # 2209 es un error puntual
        + [f"{M2} {C2}"] * 2
        + [f"{M3} {C3}"]                           # solo una aparición
    )
    dic = mc.construir_dic_calcas(valores)
    por_mat = {f["MATRICULA"]: f for f in dic}
    assert por_mat[M1]["CALCA"] == C1 and por_mat[M1]["ESTADO"] == "FIABLE"
    assert por_mat[M2]["ESTADO"] == "FIABLE"
    assert por_mat[M3]["ESTADO"] == "POCOS DATOS"
    return mc.mapas_calcas(dic)


def test_ya_correcta(mapas):
    r = mc.resolver(f"{M1} {C1}", *mapas)
    assert r["RESULTADO"] == f"{M1} {C1}"
    assert r["ESTADO"] == mc.FIABLE


def test_calca_mal_escrita_se_corrige(mapas):
    r = mc.resolver(f"{M1} 2209", *mapas)
    assert r["RESULTADO"] == f"{M1} {C1}"


def test_matricula_sola_recibe_su_calca(mapas):
    r = mc.resolver(M2.lower(), *mapas)
    assert r["RESULTADO"] == f"{M2} {C2}"
    assert r["TIPO"] == "SOLO MATRICULA"


def test_calca_sola_recibe_su_matricula(mapas):
    r = mc.resolver(C1, *mapas)
    assert r["RESULTADO"] == f"{M1} {C1}"
    assert r["TIPO"] == "SOLO CALCA"


def test_calca_poco_fiable_va_a_revisar(mapas):
    r = mc.resolver(M3, *mapas)
    assert r["RESULTADO"] == f"{M3} {C3}"
    assert r["ESTADO"] == mc.REVISAR


def test_calca_desconocida_sin_resolver(mapas):
    r = mc.resolver("9999", *mapas)
    assert r["ESTADO"] == mc.SIN_RESOLVER


def test_calca_repetida_en_dos_matriculas():
    dic = mc.construir_dic_calcas([f"{M1} {C1}"] * 2 + [f"{M2} {C1}"] * 2)
    estados = {f["MATRICULA"]: f["ESTADO"] for f in dic}
    assert estados[M1] == estados[M2] == "CALCA REPETIDA"
