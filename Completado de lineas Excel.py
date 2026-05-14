import streamlit as st
import pandas as pd
import io

# Configuración de la aplicación
st.set_page_config(page_title="Procesador Contable", page_icon="📈")

st.title("📈 Procesador de Reportes e Información")
st.write("Esta herramienta limpia filas de totales, encabezados sueltos, completa datos faltantes (ffill) y acomoda las columnas de Periodo y Total.")

# 1. Subida del archivo
archivo_subido = st.file_uploader("Subí tu archivo de Excel (.xls o .xlsx)", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        # Identificamos el formato para elegir el motor correcto
        nombre_archivo = archivo_subido.name.lower()
        if nombre_archivo.endswith('.xls'):
            motor = 'xlrd'
        else:
            motor = 'openpyxl'
            
        # Leemos el archivo
        df = pd.read_excel(archivo_subido, engine=motor)
        st.success(f"Archivo cargado: {nombre_archivo}")
        
        st.subheader("Configuración del Proceso")
        
        # Opciones de limpieza
        limpiar_vacios = st.checkbox("Eliminar filas totalmente vacías", value=True)
        limpiar_totales = st.checkbox("Eliminar filas con palabras 'Total' o 'Subtotal'", value=True)
        
        # Selector de columna para el ffill
        columnas = df.columns.tolist()
        columna_a_rellenar = st.selectbox("¿Qué columna querés completar hacia abajo? (Ej: Localidad)", columnas)

        # Selección de Periodo
        st.markdown("**Información del Periodo**")
        col_mes, col_anio = st.columns(2)
        with col_mes:
            mes = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                                        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
        with col_anio:
            anio = st.selectbox("Año",
