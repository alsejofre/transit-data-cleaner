"""
Genera Excels de EJEMPLO (datos 100 % sintéticos) en data/raw para probar
los cuatro pipelines sin necesitar datos reales:

    Aforo_en_Parada_Madrid.xlsx       -> scripts/Matriculas
    Encuestas_Guadalajara.xlsx        -> scripts/Barrios
    Matricules_Control_Parada.xlsx    -> scripts/ControlParada
    Mistery.xlsx                      -> scripts/Mistery

Las matrículas salen de config/matriculas.example.txt y se les meten los
errores típicos de transcripción (O↔0, S↔5, dígitos cambiados, espacios...).

Uso:
    python examples/generar_datos_ejemplo.py           # no pisa archivos existentes
    python examples/generar_datos_ejemplo.py --force   # sobrescribe
"""

import os
import random
import sys
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_RAW = os.path.join(BASE_DIR, "data", "raw")
EJEMPLO = os.path.join(BASE_DIR, "config", "matriculas.example.txt")

SEMILLA = 2026
CONFUSIONES = {"0": "O", "1": "I", "5": "S", "8": "B", "2": "Z", "6": "G",
               "M": "N", "B": "8", "S": "5", "G": "6", "Z": "2", "K": "X"}


def leer_ejemplo():
    out = []
    with open(EJEMPLO, encoding="utf-8") as f:
        for linea in f:
            linea = linea.split("#", 1)[0]
            out += [x.strip() for x in linea.replace(";", ",").split(",") if x.strip()]
    return out


def ensuciar(m, rnd, prob=0.3):
    """Mete un error típico de transcripción con probabilidad `prob`."""
    if rnd.random() > prob:
        return rnd.choice([m, m, m.lower(), f"{m[:4]} {m[4:]}"])
    tipo = rnd.choice(["confusion", "digito", "transposicion", "espacio", "guion"])
    s = list(m)
    if tipo == "confusion":
        idx = [i for i, c in enumerate(s) if c in CONFUSIONES]
        if idx:
            i = rnd.choice(idx)
            s[i] = CONFUSIONES[s[i]]
    elif tipo == "digito":
        i = rnd.randrange(4)
        s[i] = str((int(s[i]) + rnd.choice([1, 9])) % 10)
    elif tipo == "transposicion":
        i = rnd.randrange(3)
        s[i], s[i + 1] = s[i + 1], s[i]
    elif tipo == "espacio":
        return f"{m[:4]} {m[4:].lower()}"
    else:
        return f"{m[:4]}-{m[4:]}"
    return "".join(s)


def libro(cabeceras, color):
    wb = Workbook()
    ws = wb.active
    ws.append(cabeceras)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color=color)
    return wb, ws


def guardar(wb, nombre, force):
    ruta = os.path.join(DATA_RAW, nombre)
    if os.path.exists(ruta) and not force:
        print(f"⏭️  {nombre} ya existe (usa --force para sobrescribir)")
        return
    wb.save(ruta)
    print(f"✅ {nombre}")


def main(force=False):
    os.makedirs(DATA_RAW, exist_ok=True)
    rnd = random.Random(SEMILLA)
    flota = leer_ejemplo()
    inicio = datetime(2026, 10, 27, 6, 30)

    # ── Aforo en parada (Matriculas) ────────────────────────
    wb, ws = libro(["FECHA", "HORA", "PARADA", "LINEA", "MATRICULA", "SUBEN", "BAJAN"], "DDEBF7")
    for i in range(300):
        t = inicio + timedelta(minutes=4 * i)
        ws.append([t.date(), t.time(), f"P{rnd.randint(1, 40):03d}", f"L{rnd.randint(1, 12)}",
                   ensuciar(rnd.choice(flota), rnd), rnd.randint(0, 25), rnd.randint(0, 25)])
    guardar(wb, "Aforo_en_Parada_Madrid.xlsx", force)

    # ── Encuesta origen/destino (Barrios) ───────────────────
    lugares = ["LA ESTACIÓN", "ESTACIÓN", "Estacion", "EL HOSPITAL", "HOSPITAL", "CENTRO",
               "el centro", "LOS MANANTIALES", "MANANTIALES", "UNIVERSIDAD", "LA UNIVERSIDAD",
               "POLÍGONO", "EL POLÍGONO", "AGUAS VIVAS", "Aguas Vivas"]
    wb, ws = libro(["ID", "ORIGEN", "DESTINO", "MOTIVO"], "E2EFDA")
    for i in range(1, 121):
        o, d = rnd.sample(lugares, 2)
        ws.append([i, o, d, rnd.choice(["TRABAJO", "ESTUDIOS", "COMPRAS", "OCIO"])])
    guardar(wb, "Encuestas_Guadalajara.xlsx", force)

    # ── Control en parada (ControlParada) ───────────────────
    wb, ws = libro(["FECHA", "PARADA", "MATRICULA", "LINEA"], "DDEBF7")
    for i in range(120):
        t = inicio + timedelta(minutes=7 * i)
        ws.append([t.date(), f"P{rnd.randint(1, 40):03d}", ensuciar(rnd.choice(flota), rnd),
                   f"L{rnd.randint(1, 12)}"])
    ws.append([inicio.date(), "P999", "XXXX", "L1"])  # basura: debe quedar sin resolver
    guardar(wb, "Matricules_Control_Parada.xlsx", force)

    # ── Mistery (MATRICULA_CALCA) ───────────────────────────
    calcas = {m: str(2200 + i) for i, m in enumerate(flota)}
    wb, ws = libro(["ID", "FECHA", "LINEA", "MATRICULA_CALCA", "PUNTUACION"], "FCE4D6")
    for i in range(1, 201):
        m = rnd.choice(flota)
        c = calcas[m]
        if rnd.random() < 0.1:  # calca mal escrita
            c = c[:3] + str((int(c[3]) + 1) % 10)
        forma = rnd.random()
        if forma < 0.6:
            valor = f"{ensuciar(m, rnd, 0.2)} {c}"
        elif forma < 0.8:
            valor = ensuciar(m, rnd, 0.2)
        else:
            valor = c
        t = inicio + timedelta(minutes=9 * i)
        ws.append([f"{i:05d}", t, f"L{rnd.randint(1, 12)}", valor, rnd.randint(1, 10)])
    guardar(wb, "Mistery.xlsx", force)


if __name__ == "__main__":
    main(force="--force" in sys.argv)
