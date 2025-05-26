import streamlit as st
import pandas as pd
import re
from io import BytesIO
import zipfile
import os
from tempfile import TemporaryDirectory

def process_file(file, folder_base):
    try:
        df = pd.read_csv(file, delimiter='|', dtype=str)
        df.columns = df.columns.str.strip()

        if 'NRO.FACTURA' not in df.columns:
            st.error("La columna 'NRO.FACTURA' no existe en el archivo.")
            return

        # Limpieza de espacios en columnas de texto
        for col in df.select_dtypes(include='object'):
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

        df.dropna(how='all', inplace=True)
        df.sort_values(by='NRO.FACTURA', inplace=True)

        # Archivo Excel completo
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        unique_invoices = df['NRO.FACTURA'].nunique()
        st.info(f"Se generarán {unique_invoices} archivos únicos por número de factura.")

        # Generar ZIP con carpetas por COBERTURA
        zip_output = generate_zip_by_coverage(df, folder_base)

        st.success("Archivo convertido y listo para descargar.")
        st.download_button(
            label="Descargar archivo Excel completo",
            data=output,
            file_name="archivo_completo.xlsx"
        )

        st.download_button(
            label="Descargar archivos por COBERTURA en ZIP",
            data=zip_output,
            file_name="facturas_por_cobertura.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")

def generate_zip_by_coverage(df, folder_base):
    zip_buffer = BytesIO()
    safe_base = re.sub(r'\W+', '_', folder_base.strip()) or "Facturas"

    with TemporaryDirectory() as temp_dir:
        for cobertura, cobertura_group in df.groupby('COBERTURA'):
            safe_cobertura = re.sub(r'\W+', '', str(cobertura))
            cobertura_dir = os.path.join(temp_dir, safe_cobertura)
            os.makedirs(cobertura_dir, exist_ok=True)

            for factura, factura_group in cobertura_group.groupby('NRO.FACTURA'):
                safe_factura = re.sub(r'\W+', '', str(factura))
                filename = f"Factura_{safe_factura}_{safe_cobertura}.xlsx"
                filepath = os.path.join(cobertura_dir, filename)

                factura_group.to_excel(filepath, index=False, engine='openpyxl')

        # Crear el ZIP
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.join(safe_base, os.path.relpath(full_path, temp_dir))
                    zipf.write(full_path, arcname)

    zip_buffer.seek(0)
    return zip_buffer

# Interfaz de usuario
st.title("Convertidor TXT a Excel con separación por COBERTURA")

uploaded_file = st.file_uploader("Selecciona un archivo .txt para convertir a Excel", type="txt")
folder_base = st.text_input("Nombre de la carpeta raíz para los archivos generados", value="Facturas")

if st.button("Convertir"):
    if uploaded_file:
        process_file(uploaded_file, folder_base)
    else:
        st.error("Por favor, sube un archivo válido.")
