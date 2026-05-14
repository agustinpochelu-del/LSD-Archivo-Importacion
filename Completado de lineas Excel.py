import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Configuración de la aplicación
st.set_page_config(page_title="Procesador Contable Profesional", page_icon="📈")

st.title("📈 Procesador de Reportes e Información")
st.write("Esta herramienta limpia filas, completa datos y genera columnas con formato de fecha real para Excel.")

# 1. Subida del archivo
archivo_subido = st.file_uploader("Subí tu archivo de Excel (.xls o .xlsx)", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        # Identificamos el formato para elegir el motor correcto
        nombre_archivo = archivo_subido.name.lower()
        motor = 'xlrd' if nombre_archivo.endswith('.xls') else 'openpyxl'
            
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

        # --- SELECCIÓN DE FECHA MEDIANTE CALENDARIO ---
        st.markdown("**Seleccioná el Periodo (Mes y Año)**")
        fecha_calendario = st.date_input("Elegí una fecha del mes que necesitás procesar", value=datetime.now())
        
        # Forzamos a que sea el día 1 de ese mes para estandarizar el periodo
        fecha_periodo = datetime(fecha_calendario.year, fecha_calendario.month, 1)
        # ----------------------------------------------

        if st.button("Procesar y Limpiar"):
            df_temp = df.copy()

            # A. Eliminamos filas vacías
            if limpiar_vacios:
                df_temp = df_temp.dropna(how='all')

            # B. Filtro para Totales y Subtotales
            if limpiar_totales:
                regex_busqueda = 'total|subtotal'
                mask = df_temp.apply(lambda row: row.astype(str).str.contains(regex_busqueda, case=False).any(), axis=1)
                df_temp = df_temp[~mask]

            # C. Completamos los datos hacia abajo (ffill)
            df_temp[columna_a_rellenar] = df_temp[columna_a_rellenar].ffill()

            # D. Eliminar filas que solo tienen dato en la columna que rellenamos 
            columnas_restantes = [col for col in df_temp.columns if col != columna_a_rellenar]
            df_temp = df_temp.dropna(subset=columnas_restantes, how='all')

            # E. Eliminamos la columna "Total" completa
            if 'Total' in df_temp.columns:
                df_temp = df_temp.drop(columns=['Total'])

            # F. Insertamos la columna "Periodo" como OBJETO FECHA
            # Esto permite que Excel lo reconozca como fecha automáticamente
            df_temp.insert(2, "Periodo", fecha_periodo)

            st.success("¡Proceso finalizado con éxito!")
            st.dataframe(df_temp.head(20))

            # Preparación del archivo para descarga con formato de fecha específico
            output = io.BytesIO()
            # Usamos un motor de escritura que nos permite dar formato a las fechas
            with pd.ExcelWriter(output, engine='openpyxl', datetime_format='MMM/YYYY') as writer:
                df_temp.to_excel(writer, index=False, sheet_name='Reporte_Procesado')
            
            st.download_button(
                label="📥 Descargar Excel con Formato Fecha",
                data=output.getvalue(),
                file_name="reporte_finalizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Se produjo un error al procesar el archivo: {e}")
