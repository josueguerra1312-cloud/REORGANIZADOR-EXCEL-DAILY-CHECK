# Ordenador de Programa AC/WO

App en Python + Streamlit para reacomodar un Excel de programa de mantenimiento.

## Qué hace

- Lee archivos `.xls` y `.xlsx`.
- Agrupa por matrícula `AC` y orden de trabajo `WO`.
- Combina visualmente las celdas de `AC` y `WO` por grupo.
- Ordena las tareas dentro de cada grupo así:
  1. `WEEKLY CHECK`
  2. `DAILY CHECK`
  3. EO normales
  4. Defectos / non-routines
- Usa siempre el texto completo de `EO_DESCRIPTION`.
- Agrega `P/N` y `S/N` en la descripción solo si ambos campos existen.
- Genera un `.xlsx` con formato profesional:
  - encabezados,
  - anchos ajustados,
  - textos ajustados,
  - panel congelado,
  - colores por categoría,
  - AC/WO combinados.

## Estructura del repositorio

```text
.
├── app.py
├── processor.py
├── requirements.txt
├── README.md
├── .gitignore
└── .streamlit/
    └── config.toml
