import streamlit as st
import pandas as pd
import io

# Configuración de la página web
st.set_page_config(page_title="Procesador de Tendencias", layout="centered")

st.title("Mi Primer Tablero de Tendencias")
st.write("Subí tu archivo Excel para generar la solapa de 'Resultado' cada 10 minutos fijo para el mes correspondiente.")

# 1. Selector de archivos
uploaded_file = st.file_uploader("Seleccioná el archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 2. Leer la solapa 'Fuente'
        df_fuente = pd.read_excel(uploaded_file, sheet_name="Fuente")
        st.success("¡Archivo cargado con éxito! Iniciando procesamiento...")

        # 3. Limpieza de nombres de columnas
        df_fuente.columns = df_fuente.columns.str.strip().str.replace('"', '')
        df_fuente.rename(columns={df_fuente.columns[0]: 'Fecha'}, inplace=True)

        # 4. LEER LA FUENTE EN FORMATO LATINO Y FIJAR SEGUNDOS EN 00
        df_fuente['Fecha'] = df_fuente['Fecha'].astype(str).str.strip()
        df_fuente['Fecha'] = pd.to_datetime(
            df_fuente['Fecha'], 
            dayfirst=True, 
            errors='coerce', 
            format='mixed'
        )
        df_fuente = df_fuente.dropna(subset=['Fecha'])
        
        if len(df_fuente) == 0:
            st.error("Error crítico: No se pudieron interpretar las fechas del archivo.")
            st.stop()
            
        # !!! TU PROPUESTA: Llevamos los segundos a :00 truncando al bloque de 10 min de forma fija !!!
        df_fuente['Fecha'] = df_fuente['Fecha'].dt.floor('10min')
            
        # 5. LIMPIEZA DE TEXTOS VACÍOS EN SENSORES
        for col in df_fuente.columns:
            if col != 'Fecha':
                df_fuente[col] = df_fuente[col].astype(str).str.strip()
                df_fuente[col] = df_fuente[col].replace(['', 'nan', 'NaN', 'None', ','], pd.NA)
                df_fuente[col] = pd.to_numeric(df_fuente[col], errors='coerce')

        # Ordenamos cronológicamente antes de agrupar
        df_fuente = df_fuente.sort_values('Fecha').reset_index(drop=True)
        
        # Hacemos un arrastre inicial en la fuente para que no se pierdan estados previos
        columnas_sensores = [col for col in df_fuente.columns if col != 'Fecha']
        df_fuente[columnas_sensores] = df_fuente[columnas_sensores].ffill().fillna(0)
        
        # Agrupamos por la nueva fecha fija de 10 minutos, quedándonos con el último cambio real
        df_fuente = df_fuente.groupby('Fecha').last().reset_index()

        # 6. ACOTACIÓN DINÁMICA AL MES Y AÑO DE LA FUENTE
        fecha_base = df_fuente['Fecha'].iloc[0]
        an_datos = fecha_base.year
        mes_datos = fecha_base.month
        
        inicio_mes = pd.Timestamp(year=an_datos, month=mes_datos, day=1, hour=0, minute=0, second=0)
        fin_mes = inicio_mes + pd.offsets.MonthEnd(1) + pd.Timedelta(hours=23, minutes=50)
        
        grilla_temporal = pd.date_range(start=inicio_mes, end=fin_mes, freq='10min')
        df_resultado = pd.DataFrame({'Fecha': grilla_temporal})
        
        meses_nombres = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 
                         7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        nombre_mes_actual = meses_nombres.get(mes_datos, "Mes Detectado")

        # 7. CRUCE DIRECTO CON LA GRILLA ANUAL (BuscarV Perfecto sobre horas fijas)
        df_final = pd.merge(df_resultado, df_fuente, on='Fecha', how='left')
        
        # Arrastre vertical final sobre la grilla para cubrir los baches de 10 minutos intermedios
        df_final[columnas_sensores] = df_final[columnas_sensores].ffill().fillna(0)

        # 8. Forzar valores a enteros limpios (0 o 1)
        for col in columnas_sensores:
            df_final[col] = df_final[col].astype(float).astype(int)
        
        # 9. LÓGICA DE DETECCIÓN DE CAMBIOS DE ESTADO (TESTIGO)
        hubo_cambio = df_final[columnas_sensores].diff().fillna(0) != 0
        df_final['Cambio'] = hubo_cambio.any(axis=1).map({True: 1, False: ""})
        df_final.loc[0, 'Cambio'] = "" 

        # 10. FORMATO LATINO CON SEGUNDOS EN 00 PARA EXCEL
        df_final['Fecha'] = df_final['Fecha'].dt.strftime('%d/%m/%Y %H:%M:00')

        # 11. Mostrar vista previa en Streamlit
        st.subheader(f"Vista previa del Resultado (Acotado a {nombre_mes_actual} {an_datos}):")
        st.dataframe(df_final.head(20))
        st.info(f"Total de filas generadas para {nombre_mes_actual}: {len(df_final):,}")
        
        # 12. Crear el archivo Excel de salida
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Resultado', index=False)
        bytes_data = output.getvalue()
        
        # 13. Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Resultado",
            data=bytes_data,
            file_name=f"Resultado_Cadencia_10min_{nombre_mes_actual}_{an_datos}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error inesperado al procesar: {e}")
        
