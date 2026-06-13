import streamlit as st
import pandas as pd
import io

# Configuración de la página web
st.set_page_config(page_title="Procesador de Tendencias", layout="centered")

st.title("Mi Primer Tablero de Tendencias")
st.write("Subí tu archivo Excel para regularizar la cadencia temporal a 10 minutos fijo para todo el año.")

# 1. Selector de archivos
uploaded_file = st.file_uploader("Seleccioná el archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 2. Leer la solapa 'Fuente'
        df_fuente = pd.read_excel(uploaded_file, sheet_name="Fuente")
        st.success("¡Archivo cargado con éxito! Procesando grilla temporal...")

        # 3. Limpieza profunda de nombres de columnas (sacamos espacios y comillas dobles)
        df_fuente.columns = df_fuente.columns.str.strip().str.replace('"', '')
        
        # 4. Convertir la columna Fecha a formato de tiempo real (Datetime)
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True, errors='coerce')
        
        # Eliminamos filas que hayan quedado con la fecha vacía o rota
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        # 5. Redondear las fechas reales al bloque de 10 minutos más cercano hacia abajo
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.floor('10min')
        
        # Si hay más de un cambio en el mismo bloque de 10 min, nos quedamos con el último
        df_fuente = df_fuente.drop_duplicates(subset=['Fecha'], keep='last')
        
        # 6. Definir la Fecha como el Índice para trabajar la serie de tiempo
        df_fuente.set_index('Fecha', inplace=True)
        
        # 7. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        # Usamos 'an_datos' (sin la Ñ) de forma consistente para evitar el NameError
        an_datos = df_fuente.index.year[0]
        inicio_ano = f"{an_datos}-01-01 00:00:00"
        fin_ano = f"{an_datos}-12-31 23:50:00"
        
        # Generamos el rango perfecto de filas cada 10 minutos
        grilla_temporal = pd.date_range(start=inicio_ano, end=fin_ano, freq='10min')
        
        # 8. REINDEXACIÓN E INTERPOLACIÓN (Forward Fill)
        df_resultado = df_fuente.reindex(grilla_temporal)
        
        # Arrastramos el último estado conocido (0 o 1) hacia abajo en los baches vacíos
        df_resultado.ffill(inplace=True)
        
        # Las filas del principio del año (antes del primer cambio registrado) se completan con 0
        df_resultado.fillna(0, inplace=True)
        
        # Nos aseguramos de que todos los valores de los sensores queden como enteros (0 o 1) sin decimales
        columnas_sensores = [col for col in df_resultado.columns if col != 'Fecha']
        df_resultado[columnas_sensores] = df_resultado[columnas_sensores].astype(int)
        
        # Reacomodamos el índice para que vuelva a ser una columna llamada 'Fecha'
        df_resultado.index.name = 'Fecha'
        df_resultado.reset_index(inplace=True)
        
        # Convertimos las fechas a un formato de texto limpio para que Excel no las rompa
        df_resultado['Fecha'] = df_resultado['Fecha'].dt.strftime('%d/%m/%Y %H:%M')

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

    except ValueError as ve:
        st.error(f"Error: Asegurate de que el archivo subido tenga una solapa llamada 'Fuente'. Detalle: {ve}")
    except Exception as e:
        st.error(f"Ocurrió un error inesperado al procesar: {e}")
        
