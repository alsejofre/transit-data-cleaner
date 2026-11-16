# Barrios · Normalización de origen y destino

Unifica los nombres de lugares (barrios, calles, municipios) que los encuestadores
escriben de formas distintas en las columnas `ORIGEN` y `DESTINO` de una encuesta.

Entrada: `data/raw/Encuestas_Guadalajara.xlsx` (columnas `ORIGEN` y `DESTINO`).

## Flujo

```
data/raw/Encuestas_Guadalajara.xlsx
        │
        ▼  01_lugares_unicos.py     -> lista de lugares distintos
        ▼  02_lugares_agrupados.py  -> agrupa variantes ("LA ESTACIÓN" / "ESTACIÓN")
data/processed/lugares_agrupados.xlsx   ← rellenas la columna CORRECTO
        │
        ▼  03_corregir_lugares.py
data/processed/viajes_corregidos.xlsx   (ORIGEN_CORREGIDO / DESTINO_CORREGIDO)
```

```bash
python scripts/Barrios/run_pipeline.py              # pasos 1 y 2
python scripts/Barrios/run_pipeline.py --corregir   # paso 3
```

Solo se cambian los lugares que tengan algo escrito en `CORRECTO`; el resto se queda igual.
La columna `SUGERENCIA` propone la variante más corta de cada grupo como ayuda.
