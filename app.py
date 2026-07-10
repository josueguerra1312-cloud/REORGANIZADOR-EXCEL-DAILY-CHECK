from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from processor import process_file

st.set_page_config(
    page_title="Ordenador de Programa AC/WO",
    page_icon="✈️",
    layout="centered",
)

st.title("✈️ Ordenador de Programa AC/WO")
st.caption(
    "Agrupa por matrícula y WO, ordena WEEKLY, DAILY, EO normales y defectos, "
    "y genera un Excel profesional."
)

st.markdown(
    """
### Instrucciones
1. Sube tu archivo **.xls** o **.xlsx** exportado del sistema.
2. Presiona **Generar Excel ordenado**.
3. Descarga el archivo final.

La descripción se arma con el texto completo de `EO_DESCRIPTION` y solo agrega `P/N` + `S/N` cuando ambos existen.
"""
)

uploaded_file = st.file_uploader("Sube el Excel a ordenar", type=["xls", "xlsx"])

if uploaded_file is None:
    st.warning("Sube un archivo para comenzar.")
else:
    st.info(f"Archivo cargado: {uploaded_file.name}")

    if st.button("Generar Excel ordenado", type="primary"):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                input_path = tmpdir_path / uploaded_file.name
                output_path = tmpdir_path / f"{Path(uploaded_file.name).stem}_ordenado.xlsx"

                input_path.write_bytes(uploaded_file.getvalue())

                process_file(
                    input_file=input_path,
                    output_path=output_path,
                    filename=uploaded_file.name,
                )

                st.success("Listo. Tu archivo fue generado correctamente.")

                st.download_button(
                    label="Descargar Excel ordenado",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as exc:
            st.error(
                "No se pudo procesar el archivo. "
                "Revisa que el Excel tenga las columnas requeridas."
            )
            st.exception(exc)
