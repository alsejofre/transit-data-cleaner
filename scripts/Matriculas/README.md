# Matriculas · Corrector sin diccionario previo

Sistema automático para limpiar, agrupar y corregir matrículas a partir de datos reales
(aforos en parada). No necesita una lista de matrículas buenas: el diccionario se
deduce de los propios datos.

Entrada: `data/raw/Aforo_en_Parada_Madrid.xlsx` (columna `MATRICULA`).

# ¿Cómo funciona?
El proyecto sigue este flujo:

1. Generación de diccionario automático (NUMERO -> letras más frecuentes para ese número)
2. Creación de matrículas únicas
3. Agrupación inteligente + sugerencia automática (tolera errores de OCR en los dígitos)
4. Corrección final automática
5. (Opcional) Revisión / corrección manual de casos puntuales

El sistema corrige automáticamente la gran mayoría de las matrículas sin intervención
humana. Solo se marcan en rojo en resultado.xlsx las filas que han cambiado respecto
al original o para las que no hay ninguna sugerencia fiable (columna REVISAR = True).

Si quieres forzar manualmente una corrección distinta a la sugerida para una matrícula
concreta, abre matriculas_agrupadas.xlsx y escribe el valor correcto en la columna
CORRECTA de esa fila; al volver a ejecutar 03_corregir.py (o el pipeline completo) esa
corrección manual tendrá prioridad sobre la sugerencia automática.

---

# 🚀 Uso del sistema

## Ejecutar pipeline completo (automático)

```bash
python scripts/Matriculas/run_pipeline.py
```

Esto genera, en orden:

- diccionario.xlsx
- matriculas_unicas.xlsx
- matriculas_agrupadas.xlsx (con columna SUGERENCIA ya calculada)
- resultado.xlsx (resultado final, con las filas a revisar en rojo)

## Solo corregir de nuevo (tras editar CORRECTA a mano)

```bash
python scripts/Matriculas/03_corregir.py
```
