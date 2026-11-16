# Mistery · Corrector de MATRICULA_CALCA

Corrige la columna `MATRICULA_CALCA` del Excel `Mistery.xlsx`.
En esa columna puede haber:

| Lo que hay       | Ejemplo        |
|------------------|----------------|
| Solo matrícula   | `0791FTR`      |
| Solo calca       | `1001`         |
| Las dos          | `0791FTR 1001` |

Al final **todas las celdas quedan como `0000ABC 0000`** (matrícula + espacio + calca).

Usa el **mismo diccionario de matrículas** que ControlParada
(`config/matriculas.txt`). Si añades una matrícula allí,
sirve para los dos.

## Flujo

```
data/raw/Mistery.xlsx
        │
        ▼  01_detectar.py
data/processed/mistery_revision.xlsx   ← revisas aquí (2 hojas)
        │
        ▼  02_corregir.py
data/processed/Mistery_corregido.xlsx
data/processed/mistery_cambios.xlsx    (registro de cada cambio)
```

1. Pon `Mistery.xlsx` en `data/raw` (debe tener la columna `MATRICULA_CALCA`).
2. `python scripts/Mistery/run_pipeline.py`
3. Revisa `mistery_revision.xlsx`:
   - **Diccionario_Calcas**: la calca de cada matrícula, sacada de las celdas que
     traen las dos cosas (la que más se repite). Estados:
     - 🟩 `FIABLE`: 2 o más apariciones y la calca ganadora es ≥ 60 %.
     - 🟨 `POCOS DATOS` (solo 1 aparición) / `REVISAR` (confianza baja).
     - 🟥 `EMPATE` (dos calcas igual de repetidas) / `CALCA REPETIDA` (la misma calca en dos matrículas).
     - ⬜ `SIN DATOS`: esa matrícula nunca aparece con calca.
     Si alguna calca está mal o falta, escríbela en `CALCA_CORRECTA`.
   - **Revision**: cada valor distinto que se va a cambiar, con el `RESULTADO`
     y el `MOTIVO`. Si un resultado no vale, escribe el bueno en `CORRECTA`.
4. `python scripts/Mistery/run_pipeline.py --corregir`

Si te fías de todo: `python scripts/Mistery/run_pipeline.py --todo`

## Qué hace con cada celda

1. **Matrícula**: se pasa a `0000ABC` y se corrige con el diccionario.
2. **Calca con su matrícula**: si no coincide con la calca de esa matrícula, se corrige.
   Si la calca es exactamente la de *otra* matrícula parecida, se entiende que la mal
   escrita era la matrícula y se marca para revisar.
3. **Calca sola**: se le añade su matrícula. Si la calca no existe, se busca la más
   parecida (marcada para revisar).
4. **Matrícula sola**: se le añade su calca.
5. Lo que no se puede resolver se deja como está y se lista al final.

El resto del Excel (otras columnas, hojas, colores, anchos) no se toca.
