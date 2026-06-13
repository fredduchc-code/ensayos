import streamlit as st
import pandas as pd
import io

# Configuración de la página web
st.set_page_config(page_title="Procesador de Tendencias", layout="centered")

st.title("Mi Primer Tablero de Tendencias")
st.write("Subí tu archivo Excel para regularizar la cadencia temporal a 10 minutos fijó para todo el año.")

# 1. Selector de archivos
uploaded_file = st.file_uploader("Seleccioná el archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 2. Leer la solapa 'Fuente'
        df_fuente = pd.read_excel(uploaded_file, sheet_name="Fuente")
        st.success("¡Archivo cargado con éxito! Procesando grilla temporal...")

        # 3. Limpieza de nombres de columnas por seguridad
        df_fuente.columns = df_fuente.columns.str.strip()
        
        # 4. Convertir la columna Fecha a un formato que Python entienda (Datetime)
        # dayfirst=True le avisa a Python que nuestro formato es Latino (Día/Mes/Año)
        df_fuente['Fecha'] = pd.to_datetime(df_fuente['Fecha'], dayfirst=True)
        
        # 5. Redondear las fechas reales al bloque de 10 minutos más cercano
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.floor('10min')
        
        # Si hubiera más de un cambio en el mismo bloque de 10 min, nos quedamos con el último valor
        df_fuente = df_fuente.drop_duplicates(subset=['Fecha'], keep='last')
        
        # 6. Definir la Fecha como el "Índice" para poder trabajar la serie de tiempo
        df_fuente.set_index('Fecha', inplace=True)
        
        # 7. CREAR LA GRILLA COMPLETA DE TODO EL AÑO (Cada 10 minutos)
        # Detectamos el año del archivo dinámicamente según el primer registro
        año_datos = df_fuente.index.year[0]
        inicio_año = f"{año_datos}-01-01 00:00:00"
        fin_año = f"{año_datos}-12-31 23:50:00"
        
        # Generamos el rango perfecto de filas cada 10 minutos ('10min')
        grilla_temporal = pd.date_range(start=inicio_año, end=fin_año, freq='10min')
        
        # 8. INTERPOLACIÓN (Resample y Forward Fill)
        # Reindexamos el archivo a la grilla completa (se crean filas vacías donde no había datos)
        df_resultado = df_fuente.reindex(grilla_temporal)
        
        # Arrastramos el último estado conocido (0 o 1) hacia abajo para rellenar los vacíos
        df_resultado.ffill(inplace=True)
        
        # Los baches que queden al principio de todo el año (antes del primer evento) se llenan con 0
        df_resultado.fillna(0, inplace=True)
        
        # Reacomodamos el índice para que vuelva a ser una columna llamada 'Fecha'
        df_resultado.index.name = 'Fecha'
        df_resultado.reset_index(inplace=True)
        
        # Convertimos las fechas a un formato de texto limpio y amigable para Excel
        df_resultado['Fecha'] = df_resultado['Fecha'].dt.strftime('%d/%m/%Y %H:%M')

        # 9. Mostrar vista previa en la pantalla de Streamlit
        st.subheader("Vista previa del Resultado (Cadencia 10 min):")
        st.dataframe(df_resultado.head(20))
        st.info(f"Total de filas generadas para el año: {len(df_resultado):,}")
        
        # 10. Crear el archivo Excel de descarga en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, sheet_name='Resultado', index=False)
        bytes_data = output.getvalue()
        
        # 11. Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Resultado",
            data=bytes_data,
            file_name=f"Resultado_Cadencia_10min_{año_datos}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ValueError as ve:
        st.error(f"Error en los datos: Verificá que la columna se llame 'Fecha' y la solapa 'Fuente'. Detalle: {ve}")
    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")
      
