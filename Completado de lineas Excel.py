import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Limpieza Contable Pro", page_icon="🧾")

st.title("🧾 Procesador de Reportes Contables")
st.write("Subí tu reporte (XLS o XLSX), limpiá filas innecesarias y completá los datos faltantes.")

# 1. Selector de archivo (ahora acepta ambos formatos)
archivo_subido = st.file_uploader("Subí tu archivo de Excel", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        # Pandas detecta automáticamente si es .xls o .xlsx si tenés instaladas las librerías
        df = pd.read_excel(archivo_subido)
        
        st.subheader("Configuración de limpieza")
        
        # Opciones de limpieza mediante checkboxes
        eliminar_vacios = st.checkbox("Eliminar filas totalmente vacías", value=True)
        eliminar_totales = st.checkbox("Eliminar filas de 'Totales' y 'Subtotales'", value=True)
        
        columnas = df.columns.tolist()
        columna_a_completar = st.selectbox("Columna a completar (ffill)", columnas)

        if st.button("Procesar Archivo"):
            df_final = df.copy()

            # A. Eliminar filas totalmente vacías
            if eliminar_vacios:
                df_final = df_final.dropna(how='all')

            # B. Eliminar filas que contengan "Total" o "Subtotal"
            if eliminar_totales:
                # Buscamos las palabras clave en todas las celdas de la fila (sin importar mayúsculas)
                mask = df_final.apply(lambda row: row.astype(str).str.contains('total|subtotal', case=False).any(), axis=1)
                df_final = df_final[~mask]

            # C. Completar datos hacia abajo
            df_final[columna_a_completar] = df_final[columna_a_completar].ffill()

            st.success(f"¡Listo! Se procesaron {len(df_final)} filas.")
            st.dataframe(df_final.head(15))

            # Preparar descarga (siempre en .xlsx por comodidad)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Reporte_Limpio')
            
            st.download_button(
                label="📥 Descargar Reporte Limpio (.xlsx)",
                data=buffer.getvalue(),
                file_name="reporte_procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.info("Asegurate de que el archivo no esté protegido con contraseña.")
