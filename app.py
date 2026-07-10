from io import BytesIO
from datetime import datetime

import streamlit as st

from procesador_programa import procesar_programa, exportar_excel

st.set_page_config(
    page_title="Ordenador Programa Mantenimiento",
    page_icon="✈️",
    layout="centered",
)

st.title("✈️ Ordenador de Programa de Mantenimiento")
st.write(
    "Sube tu Excel original (.xls o .xlsx) y la app generará un nuevo archivo agrupado por "
    "matrícula (AC) y WO, con WEEKLY CHECK, DAILY CHECK, EO normales y defectos al final."
)

archivo = st.file_uploader("Archivo Excel original", type=["xls", "xlsx"])

with st.expander("Opciones avanzadas"):
    incluir_nr = st.checkbox(
        "Incluir defectos/N/R sin EO",
        value=False,
        help="Déjalo desactivado si quieres replicar el ejemplo procesado, donde se omiten N/R sin EO.",
    )
    ordenar = st.selectbox(
        "Orden de grupos AC/WO",
        options=["ac_wo", "aparicion"],
        format_func=lambda x: "Por AC y WO" if x == "ac_wo" else "Mantener orden de aparición",
    )

if archivo is not None:
    try:
        df_final = procesar_programa(
            archivo,
            incluir_defectos_sin_eo=incluir_nr,
            ordenar_grupos_por=ordenar,
        )

        st.success(f"Archivo procesado correctamente: {len(df_final)} tareas encontradas.")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        salida = BytesIO()
        exportar_excel(df_final, salida)

        fecha = datetime.now().strftime("%Y%m%d_%H%M")
        nombre_salida = f"programa_procesado_{fecha}.xlsx"

        st.download_button(
            label="⬇️ Descargar Excel procesado",
            data=salida.getvalue(),
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error("No se pudo procesar el archivo.")
        st.exception(e)
else:
    st.info("Carga un archivo para iniciar.")
