import streamlit as st
import pandas as pd
import io

# Configuración de la página web
st.set_page_config(page_title="Procesador de Tendencias", layout="centered")

st.title("Mi Primer Tablero de Tendencias")
st.write("Subí tu archivo Excel para generar la solapa de 'Resultado' cada 10 minutos fijo para todo el año.")

# 1. Selector de archivos
uploaded_file = st.file_uploader("Seleccioná el archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 2. Leer la solapa 'Fuente'
        df_fuente = pd.read_excel(uploaded_file, sheet_name="Fuente")
        st.success("¡Archivo cargado con éxito! Iniciando procesamiento...")

        # 3. Limpieza de nombres de columnas (quitar espacios y comillas)
        df_fuente.columns = df_fuente.columns.str.strip().str.replace('"', '')
        
        # Forzar a que la primera columna de la izquierda sea nuestra columna de tiempo
        df_fuente.rename(columns={df_fuente.columns[0]: 'Fecha'}, inplace=True)

        # 4. Convertir la columna Fecha de manera flexible (ideal para detectar medianoches y formatos latinos)
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
            
        # Eliminamos filas donde la fecha REALMENTE sea imposible de leer o esté vacía
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if len(df_fuente) == 0:
            st.error("Error: Python no pudo interpretar las fechas de la primera columna. Asegurate de que tengan un formato válido de Fecha/Hora.")
            st.stop()
            
        # Ordenamos cronológicamente la fuente antes de procesar los sensores
        df_fuente = df_fuente.sort_values('Fecha')
        
        # Rellenamos los vacíos internos de los sensores (0 y 1) en las filas originales de la fuente
        df_fuente.ffill(inplace=True)
        df_fuente.fillna(0, inplace=True)

        # 5. REDONDEO COMPATIBLE A 10 MINUTOS 
        # Convertimos a formato de período de 10 minutos y volvemos a timestamp. 
        # Soluciona el error anterior y es 100% compatible con cualquier versión de Pandas.
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.to_period('10T').dt.to_timestamp()
        
        # Agrupamos por minuto quedándonos con la última actualización de cada bloque
        df_fuente = df_fuente.groupby('Fecha').last().reset_index()

        # 6. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        an_datos = int(df_fuente['Fecha'].dt.year.max())
        
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        df_resultado = pd.DataFrame({'Fecha': grilla_temporal})
        
        # 7. CRUCE DE DATOS E INTERPOLACIÓN FINAL (BuscarV en memoria)
        df_resultado = pd.merge(df_resultado, df_fuente, on='Fecha', how='left')
        
        # Volvemos a aplicar ffill para cubrir las nuevas filas de 10 min creadas por la grilla
        df_resultado.ffill(inplace=True)
        df_resultado.fillna(0, inplace=True)
        
        # 8. Forzar que los valores de los sensores queden como enteros puros (0 o 1) sin decimales
        for col in df_resultado.columns:
            if col != 'Fecha':
                df_resultado[col] = df_resultado[col].astype(float).astype(int)
        
        # Convertir las fechas a texto legible en formato LATINO para la exportación limpia a Excel
        df_resultado['Fecha'] = df_resultado['Fecha'].dt.strftime('%d/%m/%Y %H:%M:%S')

        # 9. Mostrar vista previa en la pantalla de Streamlit
        st.subheader("Vista previa del Resultado (Solapa Resultado generada):")
        st.dataframe(df_resultado.head(20))
        st.info(f"Total de filas generadas para el año {an_datos}: {len(df_resultado):,}")
        
        # 10. Crear el archivo Excel de salida con la solapa 'Resultado'
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, sheet_name='Resultado', index=False)
        bytes_data = output.getvalue()
        
        # 11. Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Resultado",
            data=bytes_data,
            file_name=f"Resultado_Cadencia_10min_{an_datos}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error inesperado al procesar: {e}")
        
