import tempfile
from pathlib import Path

import streamlit as st


st.set_page_config(
    page_title="Ordenador de Programa AC/WO",
    page_icon="✈️",
    layout="centered",
)


st.title("✈️ Ordenador de Programa AC/WO")

st.write(
    "Sube tu archivo Excel y la app generará un nuevo archivo ordenado por "
    "matrícula AC y WO, con WEEKLY CHECK, DAILY CHECK, EO normales y defectos."
)

st.markdown(
    """
### Orden aplicado dentro de cada AC / WO

1. **WEEKLY CHECK**
2. **DAILY CHECK**
3. **EO normales**
4. **Defectos / Non-Routines**

La descripción usa siempre el texto completo de `EO_DESCRIPTION`.
