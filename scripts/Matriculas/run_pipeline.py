import subprocess
import os
import sys

print("🚀 INICIANDO PIPELINE...\n")

# 📍 carpeta donde está este script (Matriculas)
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

scripts = [
    "00_diccionario.py",
    "01_unicas.py",
    "02_agrupadas.py",
    "03_corregir.py"
]

for script in scripts:
    path = os.path.join(CURRENT_DIR, script)  # 🔥 aquí está la clave
    print(f"▶ Ejecutando {script}...")

    subprocess.run([sys.executable, path], check=True)

    print("✔ OK\n")

print("🎯 PIPELINE TERMINADO")