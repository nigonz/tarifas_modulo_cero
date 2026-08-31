import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="TTR - Módulo 0", layout="wide")

def truncar_a_dos_decimales(serie):
    s_num = pd.to_numeric(serie, errors='coerce')
    return np.trunc(s_num * 100.0 + 1e-9) / 100.0

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    df_clean = df_hist.copy()
    
    # Normalizar columnas: eliminar espacios, pasar a minúsculas para buscar sin errores de tildes o mayúsculas
    mapa_columnas = {str(c).strip().lower(): str(c).strip() for c in df_clean.columns}
    
    # Buscar nombres de columnas de forma flexible
    col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
    col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
    col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)

    if not col_concat or not col_ts or not col_nom:
        raise ValueError(f"No se pudieron identificar las columnas clave. Columnas leídas: {list(df_clean.columns)}")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_clean.iterrows():
        concat = str(row[col_concat]).strip()
        ts = str(row[col_ts]).strip().upper()
        nom = str(row[col_nom]).strip().upper()

        # Identificar la base exacta según las reglas de la TTR
        base_key = "1SCN"
        if "1-4KM" in concat:
            base_key = "5KPCN"
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

        base_inf = float(dict_bases_inf.get(base_key, 0))
        base_sup = float(dict_bases_sup.get(base_key, 0))

        # Aplicar multiplicadores estrictos
        mult_ts = 1.0
        if ts == "EA": 
            mult_ts = 1.75
        elif ts == "E": 
            mult_ts = 1.25

        mult_nom = 2.0 if "SN" in nom else 1.0

        val_inf = truncar_a_dos_decimales(base_inf * mult_ts * mult_nom)
        val_sup = truncar_a_dos_decimales(base_sup * mult_ts * mult_nom)

        nuevos_limites_inf.append(float(val_inf))
        nuevos_limites_sup.append(float(val_sup))

    df_clean[f'{mes_nuevo_nombre} - Límite Inferior'] = nuevos_limites_inf
    df_clean[f'{mes_nuevo_nombre} - Límite Superior'] = nuevos_limites_sup

    return df_clean

st.title("🚜 TTR_ARIA - Pipeline de Liquidación Tarifaria")
st.markdown("Cálculo estructurado por bases tipeadas y aplicación estricta de factores.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])
fila_header = st.sidebar.number_input("Fila de los títulos en el Excel (0 para la primera fila)", min_value=0, max_value=5, value=0)

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Julio")

st.subheader(f"Ingreso de Bases Tarifarias (Tipeo Manual): {mes_act}")
st.info("Ingrese los 10 valores base exactos. Las categorías se calcularán automáticamente por fórmula.")

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
        with st.spinner("Procesando y calculando..."):
            try:
                df_hist = pd.read_excel(archivo_historico, header=fila_header, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Cálculo completado exitosamente!")
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
