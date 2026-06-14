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

        # 4. TRATAMIENTO SEGURO Y REDONDEO DIRECTO A CADENCIA DE 10 MINUTOS
        df_fuente['Fecha'] = df_fuente['Fecha'].astype(str).str.strip()
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if len(df_fuente) == 0:
            st.error("Error crítico: No se pudieron interpretar las fechas del archivo.")
            st.stop()
            
        # Redondeamos directamente a los 10 minutos más cercanos
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.round('10min')
            
        # 5. LIMPIEZA DE TEXTOS VACÍOS EN SENSORES
        for col in df_fuente.columns:
            if col != 'Fecha':
                df_fuente[col] = df_fuente[col].astype(str).str.strip()
                df_fuente[col] = df_fuente[col].replace(['', 'nan', 'NaN', 'None', ','], pd.NA)
                df_fuente[col] = pd.to_numeric(df_fuente[col], errors='coerce')

        # Ordenamos cronológicamente de forma estricta la fuente
        df_fuente = df_fuente.sort_values('Fecha').reset_index(drop=True)
        
        # !!! EL CAMBIO CLAVE !!!
        # Hacemos el arrastre (ffill) dentro de la FUENTE primero, para que las filas del 13/01 
        # recuerden los estados en los que venían los sensores el 12/01 antes de mezclarse con el calendario general.
        columnas_sensores = [col for col in df_fuente.columns if col != 'Fecha']
        df_fuente[columnas_sensores] = df_fuente[columnas_sensores].ffill().fillna(0)
        
        # Agrupamos por si el redondeo generó duplicados, manteniendo la última actualización real
        df_fuente = df_fuente.groupby('Fecha').last().reset_index()

        # 6. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        an_datos = int(df_fuente['Fecha'].dt.year.max())
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        df_resultado = pd.DataFrame({'Fecha': grilla_temporal})
        
        # 7. CRUCE DIRECTO CON LA GRILLA ANUAL (BuscarV Perfecto)
        df_final = pd.merge(df_resultado, df_fuente, on='Fecha', how='left')
        
        # Arrastre vertical final sobre la grilla para rellenar los baches de 10 minutos vacíos
        df_final[columnas_sensores] = df_final[columnas_sensores].ffill().fillna(0)

        # 8. Forzar valores a enteros limpios (0 o 1)
        for col in columnas_sensores:
            df_final[col] = df_final[col].astype(float).astype(int)
        
        # 9. LÓGICA DE DETECCIÓN DE CAMBIOS DE ESTADO (TESTIGO)
        hubo_cambio = df_final[columnas_sensores].diff().fillna(0) != 0
        df_final['Cambio'] = hubo_cambio.any(axis=1).map({True: 1, False: ""})
        df_final.loc[0, 'Cambio'] = "" # Limpiar fila inicial

        # 10. FORMATO LATINO CON SEGUNDOS EN 00
        df_final['Fecha'] = df_final['Fecha'].dt.strftime('%d/%m/%Y %H:%M:00')

        # 11. Mostrar vista previa en Streamlit
        st.subheader("Vista previa del Resultado Regularizado (Flujo de Memoria Corregido):")
        st.dataframe(df_final.head(20))
        st.info(f"Total de filas generadas para el año {an_datos}: {len(df_final):,}")
        
        # 12. Crear el archivo Excel de salida
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Resultado', index=False)
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
        
