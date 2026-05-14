import streamlit as st
import pandas as pd
import io

# Configuración de la aplicación
st.set_page_config(page_title="Procesador Contable Agustín", page_icon="📈")

st.title("📈 Procesador de Reportes e Información")
st.write("Esta herramienta limpia filas de totales, encabezados de localidad y completa datos faltantes (ffill).")

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
        
        st.subheader("Opciones de Limpieza")
        
        # Checkboxes para que decidas qué limpiar
        limpiar_vacios = st.checkbox("Eliminar filas totalmente vacías", value=True)
        limpiar_totales = st.checkbox("Eliminar filas con 'Total' o 'Subtotal'", value=True)
        limpiar_localidad = st.checkbox("Eliminar filas que mencionen 'Localidad'", value=True)
        
        # Selector de columna para el ffill
        columnas = df.columns.tolist()
        columna_a_rellenar = st.selectbox("¿Qué columna querés completar hacia abajo?", columnas)

        if st.button("Procesar y Limpiar"):
            df_temp = df.copy()

            # A. Eliminamos filas vacías
            if limpiar_vacios:
                df_temp = df_temp.dropna(how='all')

            # B. Filtro inteligente para Totales, Subtotales y Localidad
            # Creamos una lista de palabras a buscar
            palabras_ruido = []
            if limpiar_totales:
                palabras_ruido.extend(['total', 'subtotal'])
            if limpiar_localidad:
                palabras_ruido.append('localidad')
            
            if palabras_ruido:
                regex_busqueda = '|'.join(palabras_ruido)
                # Buscamos en todas las columnas de la fila
                mask = df_temp.apply(lambda row: row.astype(str).str.contains(regex_busqueda, case=False).any(), axis=1)
                df_temp = df_temp[~mask]

# C. Completamos los datos hacia abajo (ffill)
            df_temp[columna_a_rellenar] = df_temp[columna_a_rellenar].ffill()

            # D. Eliminar filas que solo tienen dato en la columna que rellenamos
            # (Ej: cuando dice "BARILOCHE" pero el resto de las columnas están vacías)
            columnas_restantes = [col for col in df_temp.columns if col != columna_a_rellenar]
            df_temp = df_temp.dropna(subset=columnas_restantes, how='all')

            st.success("¡Proceso finalizado!")
            st.write(f"Filas resultantes: {len(df_temp)}")
            st.dataframe(df_temp.head(20))

            # Preparación del archivo para descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_temp.to_excel(writer, index=False, sheet_name='Reporte_Procesado')
            
            st.download_button(
                label="📥 Descargar Excel Limpio",
                data=output.getvalue(),
                file_name="reporte_finalizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Se produjo un error al procesar el archivo: {e}")
