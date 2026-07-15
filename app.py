import streamlit as st
from procesador import procesar_excel


st.set_page_config(
    page_title="Ordenador Daily Check",
    page_icon="✈️",
    layout="centered",
)


st.title("✈️ Ordenador de Programa de Mantenimiento")

st.write(
    """
Carga tu archivo Excel original en formato **.xls** o **.xlsx**.
La aplicación generará automáticamente un nuevo Excel agrupado por **AC** y **WO**,
con las tareas ordenadas y con formato profesional.
"""
)

archivo = st.file_uploader(
    "Sube el Excel sin ordenar",
    type=["xls", "xlsx"],
)

nombre_salida = st.text_input(
    "Nombre del archivo de salida",
    value="PROGRAMA_PROCESADO.xlsx",
)

if archivo is not None:
    st.info(f"Archivo cargado: {archivo.name}")

    if st.button("Generar Excel ordenado", type="primary"):
        try:
            with st.spinner("Procesando archivo..."):
                excel_procesado = procesar_excel(archivo)

            st.success("Excel procesado correctamente.")

            if not nombre_salida.lower().endswith(".xlsx"):
                nombre_descarga = nombre_salida + ".xlsx"
            else:
                nombre_descarga = nombre_salida

            st.download_button(
                label="Descargar Excel ordenado",
                data=excel_procesado,
                file_name=nombre_descarga,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("Ocurrió un error al procesar el archivo.")
            st.exception(e)
else:
    st.warning("Carga un archivo para comenzar.")
