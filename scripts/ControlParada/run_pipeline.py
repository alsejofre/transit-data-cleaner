import subprocess
import os
import sys

# Uso:
#   python run_pipeline.py            -> detecta (paso 1) y para para que revises
#   python run_pipeline.py --todo     -> detecta y corrige directamente con las sugerencias
#   python run_pipeline.py --corregir -> solo aplica las correcciones (paso 2)

print("🚀 INICIANDO PIPELINE CONTROL PARADA...\n")

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if "--corregir" in sys.argv:
    scripts = ["02_corregir.py"]
elif "--todo" in sys.argv:
    scripts = ["01_detectar.py", "02_corregir.py"]
else:
    scripts = ["01_detectar.py"]

for script in scripts:
    path = os.path.join(CURRENT_DIR, script)
    print(f"▶ Ejecutando {script}...")
    subprocess.run([sys.executable, path], check=True, cwd=CURRENT_DIR)
    print("✔ OK\n")

if scripts == ["01_detectar.py"]:
    print("🛑 Revisa data/processed/control_parada_revision.xlsx")
    print("   y luego ejecuta:  python run_pipeline.py --corregir")
else:
    print("🎯 PIPELINE TERMINADO")
