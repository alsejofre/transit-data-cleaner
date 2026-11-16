"""
Tests de punta a punta: copia el proyecto a una carpeta temporal, genera los
datos de ejemplo y ejecuta cada pipeline completo comprobando que:
  - termina sin errores
  - crea el Excel de salida
  - no toca las demás columnas ni el formato de la cabecera
"""

import os
import shutil
import subprocess
import sys

import pytest
from openpyxl import load_workbook

from conftest import RAIZ

PIPELINES = {
    "Matriculas": ("Aforo_en_Parada_Madrid.xlsx", "resultado.xlsx"),
    "Barrios": ("Encuestas_Guadalajara.xlsx", "viajes_corregidos.xlsx"),
    "ControlParada": ("Matricules_Control_Parada.xlsx", "Matricules_Control_Parada_corregido.xlsx"),
    "Mistery": ("Mistery.xlsx", "Mistery_corregido.xlsx"),
}


@pytest.fixture(scope="module")
def proyecto(tmp_path_factory):
    dst = tmp_path_factory.mktemp("proyecto")
    for carpeta in ("scripts", "config", "examples"):
        shutil.copytree(os.path.join(RAIZ, carpeta), dst / carpeta,
                        ignore=shutil.ignore_patterns("__pycache__", "matriculas.txt"))
    (dst / "data" / "raw").mkdir(parents=True)
    (dst / "data" / "processed").mkdir(parents=True)
    subprocess.run([sys.executable, str(dst / "examples" / "generar_datos_ejemplo.py")],
                   check=True, capture_output=True)
    return dst


def ejecutar(proyecto, modulo):
    return subprocess.run(
        [sys.executable, str(proyecto / "scripts" / modulo / "run_pipeline.py"), "--todo"],
        capture_output=True, text=True, encoding="utf-8",
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )


@pytest.mark.parametrize("modulo", PIPELINES)
def test_pipeline_completo(proyecto, modulo):
    entrada, salida = PIPELINES[modulo]
    r = ejecutar(proyecto, modulo)
    assert r.returncode == 0, r.stdout + r.stderr
    assert (proyecto / "data" / "processed" / salida).exists()


@pytest.mark.parametrize("modulo, columna", [("ControlParada", "MATRICULA"),
                                             ("Mistery", "MATRICULA_CALCA")])
def test_salida_respeta_el_excel_original(proyecto, modulo, columna):
    entrada, salida = PIPELINES[modulo]
    ejecutar(proyecto, modulo)
    orig = load_workbook(proyecto / "data" / "raw" / entrada).active
    nuevo = load_workbook(proyecto / "data" / "processed" / salida).active

    assert orig.max_row == nuevo.max_row
    cab = [c.value for c in orig[1]]
    assert cab == [c.value for c in nuevo[1]]
    assert orig["A1"].fill.start_color.rgb == nuevo["A1"].fill.start_color.rgb

    col = cab.index(columna)
    for fo, fn in zip(orig.iter_rows(min_row=2, values_only=True),
                      nuevo.iter_rows(min_row=2, values_only=True)):
        assert [v for i, v in enumerate(fo) if i != col] == [v for i, v in enumerate(fn) if i != col]


def test_control_parada_deja_todo_en_el_diccionario(proyecto):
    import comun as cp
    ejecutar(proyecto, "ControlParada")
    ws = load_workbook(proyecto / "data" / "processed" / PIPELINES["ControlParada"][1]).active
    valores = [r[0] for r in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True)]
    fuera = [v for v in valores if v not in cp.DICCIONARIO]
    assert fuera == ["XXXX"]  # solo la basura metida a propósito


def test_mistery_formato_final(proyecto):
    import re
    ejecutar(proyecto, "Mistery")
    ws = load_workbook(proyecto / "data" / "processed" / PIPELINES["Mistery"][1]).active
    valores = [r[0] for r in ws.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True)]
    ok = [v for v in valores if re.fullmatch(r"\d{4}[A-Z]{3} \d{4}", str(v))]
    assert len(ok) / len(valores) > 0.95
