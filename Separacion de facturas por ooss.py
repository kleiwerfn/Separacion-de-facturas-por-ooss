import streamlit as st
import pandas as pd
import re
from io import BytesIO
import zipfile
import os
import traceback

# Columnas a eliminar completamente
columns_to_drop = [
    'FECHA REND', 'IMPORTE REND.HC', 'ALIC.IVA', 'QUIEN FAC.', 'HORA',
    'PANTALLA', 'ADMIS', 'TIPO DE MARCA', 'PROTOCOLO 1', 'PROTOCOLO 2',
    'PROTOCOLO 3', 'PROTOCOLO 4', 'PROTOCOLO 5', 'COD.MA'
]

# Orden deseado de columnas
column_order = [
    'H.CLINICA', 'HC UNICA', 'APELLIDO Y NOMBRE', 'AFILIADO', 'PERIODO',
    'COD.OBRA', 'COBERTURA', 'PLAN', 'NRO.FACTURA', 'FECHA PRES',
    'TIP.NOM', 'COD.NOM', 'PRESTACION', 'CANTID.', 'IMPORTE UNIT.',
    'IMPORTE PREST.', 'ORIGEN'
]

# Columnas que deben convertirse a numérico
numeric_columns = [
    'H.CLINICA', 'HC UNICA', 'AFILIADO', 'TIP.NOM',
    'COD.NOM', 'CANTID.', 'IMPORTE UNIT.', 'IMPORTE PREST.'
]

def clean_and_format_dataframe(df):
    df = df.drop(columns=[col for col in columns_to_drop if col in df.columns], errors='ignore')
    existing_columns = [col for col in column_order if col in df.columns]
    df = df[existing_columns + [col for col in df.columns if col not in existing_columns]]

    for col in numeric_columns:
        if col in df.columns:
            if pd.api.types.is_string_dtype(df[col]):
                df[col] = df[col].str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df

def generate_zip_with_summary(df, folder_base):
    zip_buffer = BytesIO()
    safe_base = re.sub(r'\W+', '_', folder_base.strip()) or "Facturas"

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        grouped = df.groupby(['COBERTURA', 'NRO.FACTURA'])

        for (cobertura, factura), group in grouped:
            safe_cobertura = re.sub(r'\W+', '', str(cobertura))[:20]
            safe_factura = re.sub(r'\W+', '', str(factura))[:20]
            filename = f"{safe_base}/{safe_cobertura}/Factura_{safe_factura}_{safe_cobertura}.xlsx"

            group = clean_and_format_dataframe(group)
            excel_buffer = BytesIO()
            group.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            zipf.writestr(filename, excel_buffer.read())

        # Resumen
        summary_df = (
            df.groupby(['COBERTURA', 'NRO.FACTURA', 'APELLIDO Y NOMBRE'], as_index=False)['IMPORTE PREST.']
            .sum(numeric_only=True)
        )
        summary_buffer = BytesIO()
        summary_df.to_excel(summary_buffer, index=False, engine='openpyxl')
        summary_buffer.seek(0)
        zipf.writestr(f"{safe_base}/resumen_facturas.xlsx", summary_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

def process_file(file, folder_base):
    try:
        df = pd.read_csv(file, delimiter='|', dtype=str)
        df.columns = df.columns.str.strip()

        required_columns = ['NRO.FACTURA', 'COBERTURA']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"Faltan las siguientes columnas requeridas: {', '.join(missing)}")
            return

        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df.dropna(how='all', inplace=True)
        df.sort_values(by='NRO.FACTURA', inplace=True)

        df_clean = clean_and_format_dataframe(df)

        output = BytesIO()
        df_clean.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        unique_invoices = df_clean['NRO.FACTURA'].nunique()
        st.info(f"Se generarán {unique_invoices} archivos únicos por número de factura.")

        zip_output = generate_zip_with_summary(df, folder_base)

        st.success("Archivo convertido y listo para descargar.")
        st.download_button("📥 Descargar archivo Excel completo", data=output, file_name="archivo_completo.xlsx")
        st.download_button("📦 Descargar ZIP con facturas y resumen", data=zip_output, file_name="facturas_por_cobertura.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
        st.text(traceback.format_exc())

# Interfaz de usuario
st.title("📄 Convertidor TXT a Excel con separación por COBERTURA y resumen")

uploaded_files = st.file_uploader("Selecciona uno o más archivos .txt para convertir a Excel", type="txt", accept_multiple_files=True)
folder_base = st.text_input("📁 Nombre de la carpeta raíz para los archivos generados", value="Facturas")

if st.button("🚀 Convertir"):
    if uploaded_files:
        with st.spinner("Procesando archivos..."):
            for file in uploaded_files:
                st.subheader(f"Procesando: {file.name}")
                process_file(file, folder_base)
    else:
        st.error("Por favor, sube al menos un archivo válido.")
