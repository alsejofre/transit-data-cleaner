import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Los tests usan SIEMPRE el diccionario de ejemplo (ficticio), aunque en local
# exista config/matriculas.txt con la flota real.
os.environ["MATRICULAS_FILE"] = os.path.join(RAIZ, "config", "matriculas.example.txt")

sys.path.insert(0, os.path.join(RAIZ, "scripts", "ControlParada"))
sys.path.insert(0, os.path.join(RAIZ, "scripts", "Mistery"))
