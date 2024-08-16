import pandas as pd
import dropbox
import streamlit as st
from dotenv import load_dotenv
import os
from datetime import datetime
import calendar

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# Obtener el token de acceso de Dropbox desde el archivo .env
access_token = os.getenv('DROPBOX_ACCESS_TOKEN')

# Configurar acceso a Dropbox
dbx = dropbox.Dropbox(access_token)

# Función para validar las columnas
def validate_columns(df, required_columns):
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return False, missing_columns
    else:
        return True, None

# Función para obtener el nombre de la carpeta basado en el mes siguiente
def get_next_month_folder():
    today = datetime.today()
    next_month = today.month % 12 + 1
    year = today.year
    month_name = calendar.month_name[next_month]
    folder_name = f"{month_name} - {year}"
    return folder_name

# Subir el archivo
st.title("Validación y Carga de Archivo a Dropbox")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Validar las columnas
    required_columns = ['RUTc', 'DV', 'SEDE', 'VIGENCIA', 'NOMBRES', 'ESTADO MZ', 
                        'PAG MARZO', 'ESTADO AB', 'PAG ABRIL', 'ESTADO MY', 'PAG MAYO', 
                        'TOTAL', 'ESTADO JUN', 'PAG JUNIO', 'ESTADO JUL', 'PAG JUL RENOV', 
                        'PAG JUL ASIG', 'TOTAL JUN - JUL', 'ESTADO AGO', 'RETRO']
    
    is_valid, missing_columns = validate_columns(df, required_columns)
    
    if is_valid:
        st.success("El archivo es válido y todas las columnas requeridas están presentes.")
        
        # Obtener el nombre del archivo basado en la primera entrada de la columna SEDE
        sede_name = df.loc[0, 'SEDE']
        file_name = f"{sede_name}.xlsx"
        
        # Crear la carpeta en Dropbox si no existe
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
        
        # Guardar el archivo en Dropbox
        dropbox_file_path = f'{dropbox_folder_path}/{file_name}'
        
        with st.spinner("Cargando archivo a Dropbox..."):
            with io.BytesIO() as buffer:
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                dbx.files_upload(buffer.read(), dropbox_file_path, mode=dropbox.files.WriteMode("overwrite"))
        
        st.success(f"El archivo {file_name} ha sido guardado exitosamente en Dropbox en la ruta: {dropbox_file_path}")
    
    else:
        st.error(f"El archivo no es válido. Faltan las siguientes columnas: {missing_columns}")
else:
    st.write("Por favor, sube un archivo Excel para validar.")
