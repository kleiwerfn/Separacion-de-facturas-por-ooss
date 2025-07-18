import streamlit as st
import pandas as pd
import re
from io import BytesIO
import zipfile
import os
from tempfile import TemporaryDirectory
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
    df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)
    df = df[[col for col in column_order if col in df.columns]]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='coerce')
    existing_columns = [col for col in column_order if col in df.columns]
    df = df[existing_columns + [col for col in df.columns if col not in existing_columns]]
    return df

def generate_zip_with_summary(df, folder_base):
    zip_buffer = BytesIO()
    safe_base = re.sub(r'\W+', '_', folder_base.strip()) or "Facturas"

    with TemporaryDirectory() as temp_dir:
        for cobertura, cobertura_group in df.groupby('COBERTURA'):
            safe_cobertura = re.sub(r'\W+', '', str(cobertura))[:20]
            cobertura_dir = os.path.join(temp_dir, safe_cobertura)
            os.makedirs(cobertura_dir, exist_ok=True)

            for factura, factura_group in cobertura_group.groupby('NRO.FACTURA'):
                safe_factura = re.sub(r'\W+', '', str(factura))[:20]
                filename = f"Factura_{safe_factura}_{safe_cobertura}.xlsx"
                filepath = os.path.join(cobertura_dir, filename)
                factura_group = clean_and_format_dataframe(factura_group)
                factura_group.to_excel(filepath, index=False, engine='openpyxl')

        summary_df = (
            df.groupby(['COBERTURA', 'NRO.FACTURA', 'APELLIDO Y NOMBRE'], as_index=False)['IMPORTE PREST.']
            .sum(numeric_only=True)
        )
        summary_path = os.path.join(temp_dir, "resumen_facturas.xlsx")
        summary_df.to_excel(summary_path, index=False, engine='openpyxl')

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.join(safe_base, os.path.relpath(full_path, temp_dir))
                    zipf.write(full_path, arcname)

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

        for col in df.select_dtypes(include='object'):
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

        df.dropna(how='all', inplace=True)
        df.sort_values(by='NRO.FACTURA', inplace=True)
        df = clean_and_format_dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        unique_invoices = df['NRO.FACTURA'].nunique()
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
