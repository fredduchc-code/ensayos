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

        # 4. TRATAMIENTO ULTRA-SEGURO DE FECHAS
        df_fuente['Fecha'] = df_fuente['Fecha'].astype(str).str.strip()
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
            
        # Eliminamos filas donde la fecha esté vacía o rota
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if len(df_fuente) == 0:
            st.error("Error crítico: No se pudieron interpretar las fechas del archivo.")
            st.stop()
            
        # Ordenamos cronológicamente la fuente de forma estricta
        df_fuente = df_fuente.sort_values('Fecha').reset_index(drop=True)
        
        # !!! REESTRUCTURACIÓN DEL ARRASTRE: INDIVIDUAL POR COLUMNA !!!
        for col in df_fuente.columns:
            if col != 'Fecha':
                # Aislamos la columna como texto limpio y sin espacios invisibles
                df_fuente[col] = df_fuente[col].astype(str).str.strip()
                # Convertimos textos vacíos o nulos falsos en verdaderos NaN numéricos
                df_fuente[col] = df_fuente[col].replace(['', 'nan', 'NaN', 'None', ','], pd.NA)
                df_fuente[col] = pd.to_numeric(df_fuente[col], errors='coerce')
                
                # Hacemos el arrastre vertical INMEDIATAMENTE en esta columna antes de pasar a la siguiente
                df_fuente[col].ffill(inplace=True)
                # Si el primer registro del año vino vacío en este sensor, arranca en 0
                df_fuente[col].fillna(0, inplace=True)

        # 5. REDONDEO DE TIEMPO UNIVERSAL (Cada 10 minutos)
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.to_period('10min').dt.to_timestamp()
        
        # Eliminamos duplicados de fecha dejando el ÚLTIMO registro (el cambio más reciente de ese bloque)
        df_fuente = df_fuente.drop_duplicates(subset=['Fecha'], keep='last')

        # 6. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        an_datos = int(df_fuente['Fecha'].dt.year.max())
        
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        df_resultado = pd.DataFrame({'Fecha': grilla_temporal})
        
        # 7. CRUCE DE DATOS E INTERPOLACIÓN FINAL (Mapeo estilo BuscarV)
        df_resultado = pd.merge(df_resultado, df_fuente, on='Fecha', how='left')
        
        # Volvemos a aplicar ffill para cubrir las nuevas filas de 10 min creadas por la grilla vacía
        df_resultado.ffill(inplace=True)
        df_resultado.fillna(0, inplace=True)
        
        # 8. Forzar que los valores de los sensores queden como enteros puros (0 o 1) sin decimales
        for col in df_resultado.columns:
            if col != 'Fecha':
                df_resultado[col] = pd.to_numeric(df_resultado[col], errors='coerce').fillna(0).astype(int)
        
        # 9. LÓGICA DE DETECCIÓN DE CAMBIOS DE ESTADO (TESTIGO)
        columnas_sensores = [col for col in df_resultado.columns if col != 'Fecha']
        hubo_cambio = df_resultado[columnas_sensores].diff().fillna(0) != 0
        df_resultado['Cambio'] = hubo_cambio.any(axis=1).map({True: 1, False: ""})
        
        # Limpiamos la primera fila por el inicio de la grilla
        df_resultado.loc[0, 'Cambio'] = ""

        # 10. Convertir las fechas a texto legible en formato LATINO para la exportación limpia a Excel
        df_resultado['Fecha'] = df_resultado['Fecha'].dt.strftime('%d/%m/%Y %H:%M:%S')

        # 11. Mostrar vista previa en la pantalla de Streamlit
        st.subheader("Vista previa del Resultado (Sensores Regularizados):")
        st.dataframe(df_resultado.head(20))
        st.info(f"Total de filas generadas para el año {an_datos}: {len(df_resultado):,}")
        
        # 12. Crear el archivo Excel de salida con la solapa 'Resultado'
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, sheet_name='Resultado', index=False)
        bytes_data = output.getvalue()
        
        # 13. Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Resultado",
            data=bytes_data,
            file_name=f"Resultado_Cadencia_10min_{an_datos}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error inesperado al procesar: {e}")
        
