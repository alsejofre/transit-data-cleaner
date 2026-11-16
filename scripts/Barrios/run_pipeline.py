import subprocess
import os
import sys

# Uso:
#   python run_pipeline.py            -> genera lugares únicos y agrupados, y para para que revises
#   python run_pipeline.py --corregir -> solo aplica las correcciones (paso 3)
#   python run_pipeline.py --todo     -> los tres pasos seguidos

print("🚀 INICIANDO PIPELINE BARRIOS...\n")

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if "--corregir" in sys.argv:
    scripts = ["03_corregir_lugares.py"]
elif "--todo" in sys.argv:
    scripts = ["01_lugares_unicos.py", "02_lugares_agrupados.py", "03_corregir_lugares.py"]
else:
    scripts = ["01_lugares_unicos.py", "02_lugares_agrupados.py"]

for script in scripts:
    path = os.path.join(CURRENT_DIR, script)
    print(f"▶ Ejecutando {script}...")
    subprocess.run([sys.executable, path], check=True, cwd=CURRENT_DIR)
    print("✔ OK\n")

if "03_corregir_lugares.py" not in scripts:
    print("🛑 Rellena la columna CORRECTO en data/processed/lugares_agrupados.xlsx")
    print("   y luego ejecuta:  python run_pipeline.py --corregir")
else:
    print("🎯 PIPELINE TERMINADO")
