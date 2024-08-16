import io
import os
import pandas as pd
import streamlit as st
import dropbox
from dotenv import load_dotenv
from datetime import datetime
import calendar

# Load environment variables from .env file
load_dotenv()

# Get the Dropbox access token from the .env file
access_token = os.getenv('DROPBOX_ACCESS_TOKEN')

# Set up Dropbox access
dbx = dropbox.Dropbox(access_token)

# Function to validate columns
def validate_columns(df, required_columns):
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return False, missing_columns
    else:
        return True, None

# Function to get the folder name based on the next month
def get_next_month_folder():
    today = datetime.today()
    next_month = today.month % 12 + 1
    year = today.year
    month_name = calendar.month_name[next_month]
    folder_name = f"{month_name} - {year}"
    return folder_name

# Function to share a folder
def share_folder(dropbox_folder_path):
    try:
        shared_link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_folder_path)
        return shared_link_metadata.url
    except dropbox.exceptions.ApiError as e:
        st.error(f"Error al compartir la carpeta: {e}")
        return None

# Display the link to the format type
st.title("Validación y Carga de Archivo BAES")

st.markdown("### Formato Tipo")
st.markdown("[Descargar formato tipo](https://docs.google.com/spreadsheets/d/1zjK2javlZ0mIYfyLHJym5rXSjgObNwRM/edit?usp=sharing&ouid=100328272349448439738&rtpof=true&sd=true)")

# Upload the file
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Validate the columns
    required_columns = ['RUTc', 'DV', 'SEDE', 'VIGENCIA', 'NOMBRES',
    'Estado-mar', 'Pago-mar', 'Estado-abr', 'Pago-abr',
                        'Estado-may', 'Pago-may', 'Estado-jun', 'Pago-jun',
                        'Estado-jul', 'Pago-jul', 'Estado-ago', 'Pago-ago',
                        'Estado-sep', 'Pago-sep', 'Estado-oct', 'Pago-oct',
                        'Estado-nov', 'Pago-nov', 'Estado-dic', 'Pago-dic']
    
    is_valid, missing_columns = validate_columns(df, required_columns)
    
    if is_valid:
        st.success("El archivo es válido y todas las columnas requeridas están presentes.")
        
        # Get the file name based on the first entry in the SEDE column
        sede_name = df.loc[0, 'SEDE']
        file_name = f"{sede_name}.xlsx"
        
        # Create the folder in Dropbox if it doesn't exist
        folder_name = get_next_month_folder()
        dropbox_folder_path = f'/{folder_name}'
        
        try:
            dbx.files_create_folder_v2(dropbox_folder_path)
            st.write(f"Carpeta {folder_name} creada exitosamente en Dropbox.")
        except dropbox.exceptions.ApiError as e:
            if isinstance(e.error, dropbox.files.CreateFolderError) and e.error.is_path() and e.error.get_path().is_conflict():
                st.write(f"La carpeta {folder_name} ya existe en Dropbox.")
            else:
                raise
        
        # Save the file in Dropbox
        dropbox_file_path = f'{dropbox_folder_path}/{file_name}'
        
        with st.spinner("Cargando archivo a Dropbox..."):
            with io.BytesIO() as buffer:
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                dbx.files_upload(buffer.read(), dropbox_file_path, mode=dropbox.files.WriteMode("overwrite"))
        
        st.success(f"El archivo {file_name} ha sido guardado exitosamente en la ruta: {dropbox_file_path}")
        
    
    else:
        st.error(f"El archivo no es válido. Faltan las siguientes columnas: {missing_columns}")
else:
    st.write("Por favor, sube un archivo Excel para validar.")
