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
        st.success("¡Archivo cargado con éxito! Iniciando procesamiento temporal...")

        # 3. Limpieza de nombres de columnas (quitar espacios y comillas)
        df_fuente.columns = df_fuente.columns.str.strip().str.replace('"', '')
        
        # Forzar a que la primera columna se llame 'Fecha' explícitamente
        primera_columna = df_fuente.columns[0]
        df_fuente.rename(columns={primera_columna: 'Fecha'}, inplace=True)

        # 4. Convertir la columna Fecha a Datetime (ideal para tu formato personalizado)
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
        
        # Eliminar filas donde la fecha esté rota o realmente vacía
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if df_fuente.empty:
            st.error("Error crítico: No se pudieron detectar fechas válidas en la primera columna.")
            st.stop()
            
        # Ordenamos cronológicamente la fuente antes de procesar
        df_fuente = df_fuente.sort_values('Fecha')
        
        # !!! TRUCO CLAVE PARA TU EXCEL !!! 
        # Rellenamos los vacíos internos de la fuente original ANTES de tocar las horas.
        # Así, si EB7B1 venía en 1, se mantiene en 1 en las filas siguientes de la fuente aunque estén en blanco.
        df_fuente.ffill(inplace=True)
        df_fuente.fillna(0, inplace=True) # Lo que quede al inicio de todo va con 0

        # 5. Redondear las fechas de la fuente al bloque de 10 minutos más cercano hacia abajo
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.floor('10min')
        
        # Agrupamos por minuto quedándonos con la última actualización de ese bloque de 10 min
        df_fuente = df_fuente.groupby('Fecha').last().reset_index()

        # 6. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        an_datos = int(df_fuente['Fecha'].dt.year.iloc[0])
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        df_resultado = pd.DataFrame({'Fecha': grilla_temporal})
        
        # 7. CRUCE DE DATOS E INTERPOLACIÓN FINAL
        # Unimos la grilla perfecta de 10 min con nuestra fuente pre-procesada
        df_resultado = pd.merge(df_resultado, df_fuente, on='Fecha', how='left')
        
        # Volvemos a aplicar ffill para cubrir los minutos nuevos generados por la grilla
        df_resultado.ffill(inplace=True)
        df_resultado.fillna(0, inplace=True)
        
        # 8. Forzar que los valores de los sensores queden como enteros puros (0 o 1)
        for col in df_resultado.columns:
            if col != 'Fecha':
                df_resultado[col] = df_resultado[col].astype(int)
        
        # Convertir las fechas a texto legible para que Excel no las rompa al descargar
        df_resultado['Fecha'] = df_resultado['Fecha'].dt.strftime('%Y-%m-%d %H:%M:%S')

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
        
