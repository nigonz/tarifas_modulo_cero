import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="TTR - Módulo 0", layout="wide")

def truncar_a_dos_decimales(serie):
    """Truncamiento matemático estricto a 2 decimales para evitar desbordes binarios."""
    s_num = pd.to_numeric(serie, errors='coerce')
    return np.trunc(s_num * 100.0 + 1e-9) / 100.0

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    df_clean = df_hist.copy()
    
    # Normalizar nombres de columnas
    df_clean.columns = [str(c).strip() for c in df_clean.columns]
    
    if 'CONCAT' not in df_clean.columns or 'TS' not in df_clean.columns or 'Nominalizacion' not in df_clean.columns:
        raise ValueError("El Excel no encontró las columnas requeridas (CONCAT, TS, Nominalizacion).")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_clean.iterrows():
        concat = str(row['CONCAT']).strip()
        ts = str(row['TS']).strip().upper()
        nom = str(row['Nominalizacion']).strip().upper()

        # 1. Identificar estrictamente la base correspondiente según la jerarquía
        base_key = "1SCN"
        if "1-4KM" in concat:
            if concat.endswith("2") or "60" in concat or "75" in concat or "90" in concat:
                base_key = "5KPCN"
            else:
                base_key = "5KPCN" if "2" in concat else "1SCN" # o mapeo directo a 5KPCN para los 1-4KM especiales
        elif "SC" in concat:
            if "1SC" in concat: base_key = "1SCN"
            elif "2SC" in concat: base_key = "2SCN"
            elif "3SC" in concat: base_key = "3SCN"
            elif "4SC" in concat: base_key = "4SCN"
            elif "5SC" in concat: base_key = "5SCN"
        elif "KP" in concat:
            if "5KP" in concat: base_key = "5KPCN"
            elif "6KP" in concat: base_key = "6KPCN"
            elif "7KP" in concat: base_key = "7KPCN"
            elif "8KP" in concat: base_key = "8KPCN"
            elif "9KP" in concat: base_key = "9KPCN"

        # Extraer los valores base tipeados
        base_inf = float(dict_bases_inf.get(base_key, 0))
        base_sup = float(dict_bases_sup.get(base_key, 0))

        # 2. Aplicar multiplicadores exactos de la TTR
        mult_ts = 1.0
        if ts == "EA": 
            mult_ts = 1.75
        elif ts == "E": 
            mult_ts = 1.25

        mult_nom = 2.0 if nom == "SN" else 1.0

        # 3. Cálculo final aplicando factores y truncando a 2 decimales
        val_inf = truncar_a_dos_decimales(base_inf * mult_ts * mult_nom)
        val_sup = truncar_a_dos_decimales(base_sup * mult_ts * mult_nom)

        nuevos_limites_inf.append(float(val_inf))
        nuevos_limites_sup.append(float(val_sup))

    # Asignar las nuevas columnas al DataFrame
    df_clean[f'{mes_nuevo_nombre} - Límite Inferior'] = nuevos_limites_inf
    df_clean[f'{mes_nuevo_nombre} - Límite Superior'] = nuevos_limites_sup

    return df_clean

st.title("🚜 TTR_ARIA - Pipeline de Liquidación Tarifaria")
st.markdown("Cálculo estructurado por bases tipeadas y aplicación estricta de factores (1.25, 1.75 y 2.0).")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])
fila_header = st.sidebar.number_input("Fila de los títulos en el Excel", min_value=0, max_value=5, value=1)

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Julio")

st.subheader(f"Ingreso de Bases Tarifarias (Tipeo Manual): {mes_act}")
st.info("Ingrese los 10 valores base exactos. Las categorías Expreso (1.25), EA (1.75) y Nominalizadas SN (x2) se calcularán automáticamente por fórmula.")

datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 1266.10, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Calcular TTR por Fórmulas", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    else:
        with st.spinner("Calculando aplicando factores y bases exactas..."):
            try:
                df_hist = pd.read_excel(archivo_historico, header=fila_header, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Cálculo completado respetando las bases tipeadas y los factores de la TTR!")
                st.dataframe(df_actualizado.head(15))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_actualizado.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                st.download_button(
                    label="📥 Descargar Matriz Calculada (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Matriz_TTR_Calculada_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
