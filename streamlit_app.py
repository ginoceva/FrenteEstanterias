# streamlit_app.py
import streamlit as st
import pandas as pd
import io
from etiquqtasfrentedepositos import generate_label_pdf

st.set_page_config(page_title="FRENTE-ESTANTERIAS", layout="wide")
st.title("📦 Generador de etiquetas — Frente Depósitos")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])
if uploaded:
    excel_file_content = io.BytesIO(uploaded.read())
    xls = pd.ExcelFile(excel_file_content)
    df = pd.read_excel(excel_file_content, sheet_name=xls.sheet_names[0])

    if "Ubicaciones" not in df.columns:
        st.error("El archivo no contiene la columna 'Ubicaciones'. Revisa tu Excel.")
    else:
        st.write("Vista previa del archivo:")
        st.dataframe(df)

        pdf_buffer = io.BytesIO()
        generate_label_pdf(df, output_filename=pdf_buffer)
        pdf_buffer.seek(0)

        st.download_button(
            label="📥 Descargar PDF con etiquetas",
            data=pdf_buffer,
            file_name="etiquetas_ubicacion.pdf",
            mime="application/pdf"
        )
