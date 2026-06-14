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

        # 4. TRATAMIENTO SEGURO DE FECHAS EN LA FUENTE
        df_fuente['Fecha'] = df_fuente['Fecha'].astype(str).str.strip()
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if len(df_fuente) == 0:
            st.error("Error crítico: No se pudieron interpretar las fechas del archivo.")
            st.stop()
            
        # 5. LIMPIEZA DE TEXTOS VACÍOS EN SENSORES
        for col in df_fuente.columns:
            if col != 'Fecha':
                df_fuente[col] = df_fuente[col].astype(str).str.strip()
                df_fuente[col] = df_fuente[col].replace(['', 'nan', 'NaN', 'None', ','], pd.NA)
                df_fuente[col] = pd.to_numeric(df_fuente[col], errors='coerce')

        # Ordenamos cronológicamente y eliminamos duplicados exactos en la fuente
        df_fuente = df_fuente.sort_values('Fecha')
        df_fuente = df_fuente.drop_duplicates(subset=['Fecha'], keep='last')

        # 6. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        an_datos = int(df_fuente['Fecha'].dt.year.max())
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        
        # 7. MAPEO ULTRA-ROBUSTO CON REINDEX
        df_fuente.set_index('Fecha', inplace=True)
        indice_completo = grilla_temporal.union(df_fuente.index)
        df_procesado = df_fuente.reindex(indice_completo)
        
        # Arrastre vertical de los sensores
        columnas_sensores = df_procesado.columns.tolist()
        df_procesado[columnas_sensores] = df_procesado[columnas_sensores].ffill().fillna(0)
        
        # 8. FILTRADO FINAL A LA CADENCIA DE 10 MINUTOS
        df_final = df_procesado.loc[grilla_temporal].reset_index()
        df_final.rename(columns={'index': 'Fecha'}, inplace=True)

        # 9. Forzar valores a enteros limpios (0 o 1)
        for col in columnas_sensores:
            df_final[col] = df_final[col].astype(float).astype(int)
        
        # 10. LÓGICA DE DETECCIÓN DE CAMBIOS DE ESTADO (TESTIGO)
        hubo_cambio = df_final[columnas_sensores].diff().fillna(0) != 0
        df_final['Cambio'] = hubo_cambio.any(axis=1).map({True: 1, False: ""})
        df_final.loc[0, 'Cambio'] = "" 

        # 11. !!! FORMATO LATINO CON SEGUNDOS EN 00 FOSFATADOS !!!
        df_final['Fecha'] = df_final['Fecha'].dt.strftime('%d/%m/%Y %H:%M:00')

        # 12. Mostrar vista previa en Streamlit
        st.subheader("Vista previa del Resultado Regularizado (Segundos :00):")
        st.dataframe(df_final.head(20))
        st.info(f"Total de filas generadas para el año {an_datos}: {len(df_final):,}")
        
        # 13. Crear el archivo Excel de salida
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Resultado', index=False)
        bytes_data = output.getvalue()
        
        # 14. Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Resultado",
            data=bytes_data,
            file_name=f"Resultado_Cadencia_10min_{an_datos}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error inesperado al procesar: {e}")
        
