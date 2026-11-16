# Control Parada · Corrector de matrículas

Compara las matrículas del Excel `Matricules_Control_Parada.xlsx` con el
diccionario de matrículas buenas (`config/matriculas.txt`).
Las del diccionario son siempre las correctas: si una matrícula no está
en el diccionario, está mal y se propone una corrección.

## Flujo

```
data/raw/Matricules_Control_Parada.xlsx
        │
        ▼  01_detectar.py
data/processed/control_parada_revision.xlsx   ← revisas aquí
        │
        ▼  02_corregir.py
data/processed/Matricules_Control_Parada_corregido.xlsx
data/processed/control_parada_cambios.xlsx    (registro de cada cambio)
```

1. Pon `Matricules_Control_Parada.xlsx` en `data/raw` (debe tener una columna `MATRICULA`).
2. Ejecuta el paso 1:
   ```bash
   python scripts/ControlParada/run_pipeline.py
   ```
   Saca por pantalla y en `control_parada_revision.xlsx` las matrículas que no
   están en el diccionario, con su `SUGERENCIA`, `SIMILITUD` y `ESTADO`:
   - 🟩 **FIABLE**: una única candidata clara.
   - 🟨 **REVISAR**: similitud baja o empate entre varias (mira `ALTERNATIVAS`).
   - 🟥 **SIN SUGERENCIA**: no se parece a ninguna del diccionario.
3. Si una sugerencia no vale, escribe la buena en la columna `CORRECTA`
   (con o sin espacios, en minúsculas da igual). Si la dejas vacía se usa la `SUGERENCIA`.
4. Ejecuta el paso 2:
   ```bash
   python scripts/ControlParada/run_pipeline.py --corregir
   ```

Si te fías de todas las sugerencias: `python scripts/ControlParada/run_pipeline.py --todo`

## Qué hace con cada matrícula

- Toda la columna queda en formato `0000XXX` (4 números + 3 letras, mayúsculas, sin espacios):
  `0791 ftr`, `FTR-0791`, `791FTR` → `0791FTR`.
- Las que no están en el diccionario se cambian por `CORRECTA` o, si está vacía, por `SUGERENCIA`.
- Las que no tienen sugerencia ni corrección manual se dejan como están y se listan al final.
- El resto del Excel (otras columnas, hojas, colores, anchos) no se toca.

## Cómo se elige la sugerencia

Busca la matrícula del diccionario más parecida, teniendo en cuenta los errores
típicos de transcripción: `O↔0`, `S↔5`, `B↔8`, `Z↔2`, `I/L↔1`, `G↔6`, `M↔N`…,
números intercambiados (`1580` ↔ `1850`) y caracteres que faltan o sobran.

## Añadir o quitar matrículas del diccionario

Edita `config/matriculas.txt` (una matrícula por línea, `#` para comentarios).
Ese archivo no se sube a git; si no existe se usa `config/matriculas.example.txt`.
