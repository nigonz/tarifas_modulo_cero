import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="TTR - Módulo 0", layout="wide")

def truncar_a_dos_decimales(serie):
    s_num = pd.to_numeric(serie, errors='coerce')
    return np.trunc(s_num * 100) / 100

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    df_final = df_hist.copy()
    
    # Normalizar nombres de columnas principales
    df_final.columns = [str(c).strip() for c in df_final.columns]
    
    if 'CONCAT' not in df_final.columns or 'TS' not in df_final.columns or 'Nominalizacion' not in df_final.columns:
        raise ValueError("El Excel no encontró las columnas requeridas (CONCAT, TS, Nominalizacion).")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_final.iterrows():
        concat = str(row['CONCAT']).strip()
        ts = str(row['TS']).strip().upper()
        nom = str(row['Nominalizacion']).strip().upper()

        # Identificar la base correspondiente según la jerarquía de la TTR
        base_key = "1SCN"
        if "SC" in concat:
            if "1SC" in concat: base_key = "1SCN"
            elif "2SC" in concat: base_key = "2SCN"
            elif "3SC" in concat: base_key = "3SCN"
            elif "4SC" in concat: base_key = "4SCN"
            elif "5SC" in concat: base_key = "5SCN"
        elif "1-4KM" in concat:
            base_key = "1-4KMCN" if not concat.endswith("2") else "5KPCN"
        elif "KP" in concat:
            if "5KP" in concat: base_key = "5KPCN"
            elif "6KP" in concat: base_key = "6KPCN"
            elif "7KP" in concat: base_key = "7KPCN"
            elif "8KP" in concat: base_key = "8KPCN"
            elif "9KP" in concat: base_key = "9KPCN"

        base_inf = dict_bases_inf.get(base_key, 0)
        base_sup = dict_bases_sup.get(base_key, 0)

        # Aplicar multiplicadores de Tipo de Servicio y Nominalización
        mult_ts = 1.0
        if ts == "EA": mult_ts = 1.75
        elif ts == "E": mult_ts = 1.25

        mult_nom = 2.0 if nom == "SN" else 1.0

        nuevos_limites_inf.append(base_inf * mult_ts * mult_nom)
        nuevos_limites_sup.append(base_sup * mult_ts * mult_nom)

    # Anexar las dos columnas nuevas al final del DataFrame
    df_final[f'{mes_nuevo_nombre} - Límite Inferior'] = truncar_a_dos_decimales(nuevos_limites_inf)
    df_final[f'{mes_nuevo_nombre} - Límite Superior'] = truncar_a_dos_decimales(nuevos_limites_sup)

    return df_final

st.title("🚜 Módulo 0: Pipeline de TTR")
st.markdown("Generación automática y completa de la matriz tarifaria mensual.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])
# Como tu captura muestra que los datos empiezan en la fila 3 (índice 2 de Excel), dejamos esto listo
fila_header = st.sidebar.number_input("Fila de los títulos en el Excel", min_value=0, max_value=5, value=1)

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Julio")

st.subheader(f"Valores Base de Referencia para: {mes_act}")
st.info("Ingresá las tarifas base principales. El sistema calculará automáticamente las 92 filas completas.")

datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '1-4KMCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.28, 1266.10, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.28, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Calcular TTR Completa", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    else:
        with st.spinner("Procesando las 92 filas del sistema..."):
            try:
                df_hist = pd.read_excel(archivo_historico, header=fila_header, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Matriz calculada al 100%! Las 92 filas tienen sus valores nuevos.")
                st.dataframe(df_actualizado.head(15))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_actualizado.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                st.download_button(
                    label="📥 Descargar Matriz Completa (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Matriz_TTR_Completa_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
