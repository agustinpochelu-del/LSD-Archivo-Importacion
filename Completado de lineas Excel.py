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
            mes = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
            
        with col_anio:
            anio = st.selectbox("Año", list(range(2024, 2031)))
        
        # Tomamos las primeras 3 letras del mes elegido, lo pasamos a mayúscula y armamos el formato
        periodo_seleccionado = f"{mes[:3].upper()}/{anio}"

        if st.button("Procesar y Limpiar"):
            df_temp = df.copy()

            # A. Eliminamos filas vacías
            if limpiar_vacios:
                df_temp = df_temp.dropna(how='all')

            # B. Filtro para Totales y Subtotales (busca en toda la fila)
            if limpiar_totales:
                regex_busqueda = 'total|subtotal'
                mask = df_temp.apply(lambda row: row.astype(str).str.contains(regex_busqueda, case=False).any(), axis=1)
                df_temp = df_temp[~mask]

            # C. Completamos los datos hacia abajo (ffill)
            df_temp[columna_a_rellenar] = df_temp[columna_a_rellenar].ffill()

            # D. Eliminar filas que solo tienen dato en la columna que rellenamos 
            # (Ej: la fila que solo dice "BARILOCHE" y el resto está vacío)
            columnas_restantes = [col for col in df_temp.columns if col != columna_a_rellenar]
            df_temp = df_temp.dropna(subset=columnas_restantes, how='all')

            # E. Eliminamos la columna "Total" completa
            if 'Total' in df_temp.columns:
                df_temp = df_temp.drop(columns=['Total'])

            # F. Insertamos la columna "Periodo" en la posición 2 (entre la columna B y C)
            df_temp.insert(2, "Periodo", periodo_seleccionado)

            st.success("¡Proceso finalizado con éxito!")
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
