import streamlit as st
import pandas as pd
import numpy as np
import io
from decimal import Decimal, ROUND_HALF_UP

st.set_page_config(page_title="TTR_ARIA - Módulo 0", layout="wide")

def redondeo_excel(valor):
    """Emula el redondeo aritmético tradicional de Excel (hacia arriba desde .5)."""
    try:
        # Convertimos a string primero para evitar el ruido binario de los flotantes
        return float(Decimal(str(valor)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    except:
        return 0.0

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    df_clean = df_hist.copy()
    
    mapa_columnas = {str(c).strip().lower(): c for c in df_clean.columns}
    col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
    col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
    col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)

    if not col_concat or not col_ts or not col_nom:
        raise ValueError(f"No se encontraron las columnas requeridas. Columnas leídas: {list(df_clean.columns)}")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_clean.iterrows():
        concat = str(row[col_concat]).strip().upper()
        
        # Ignorar filas vacías
        if concat in ['NAN', 'NONE', '']:
            nuevos_limites_inf.append(np.nan)
            nuevos_limites_sup.append(np.nan)
            continue

        ts = str(row[col_ts]).strip().upper()
        nom = str(row[col_nom]).strip().upper()

        base_key_inf = "1SCN"
        base_key_sup = "1SCN"

        # Lógica estricta por prefijos para que las variantes SN mapeen perfecto
        if concat.startswith("1S"): base_key_inf = base_key_sup = "1SCN"
        elif concat.startswith("2S"): base_key_inf = base_key_sup = "2SCN"
        elif concat.startswith("3S"): base_key_inf = base_key_sup = "3SCN"
        elif concat.startswith("4S"): base_key_inf = base_key_sup = "4SCN"
        elif concat.startswith("5S"): base_key_inf = base_key_sup = "5SCN"
        elif concat.startswith("1-4KM"): 
            base_key_inf = base_key_sup = "1-4KMCN"
        elif concat.startswith("5KP"): base_key_inf = base_key_sup = "5KPCN"
        elif concat.startswith("6KP"): base_key_inf = base_key_sup = "6KPCN"
        elif concat.startswith("7KP"): base_key_inf = base_key_sup = "7KPCN"
        elif concat.startswith("8KP"): base_key_inf = base_key_sup = "8KPCN"
        elif concat.startswith("9KP"): base_key_inf = base_key_sup = "9KPCN"

        base_inf = float(dict_bases_inf.get(base_key_inf, 0))
        
        # Regla especial kilométricas: 1-4KMCN2 hereda el límite inferior de la 5KPCN como su superior
        if concat.startswith("1-4KM") and "2" in concat:
            base_sup = float(dict_bases_inf.get("5KPCN", 0))
        else:
            base_sup = float(dict_bases_sup.get(base_key_sup, 0))

        # Multiplicadores puros de la TTR (1.25, 1.75, 2.0)
        mult_ts = 1.0
        if ts == "EA": mult_ts = 1.75
        elif ts == "E": mult_ts = 1.25

        mult_nom = 2.0 if "SN" in nom else 1.0

        # Cálculo con redondeo idéntico al de Excel
        val_inf = redondeo_excel(base_inf * mult_ts * mult_nom)
        val_sup = redondeo_excel(base_sup * mult_ts * mult_nom)

        nuevos_limites_inf.append(val_inf)
        nuevos_limites_sup.append(val_sup)

    df_clean[f'{mes_nuevo_nombre}'] = nuevos_limites_inf
    col_sup_name = f'{mes_nuevo_nombre}_Sup'
    df_clean[col_sup_name] = nuevos_limites_sup

    # Limpieza final de decimales históricos
    columnas_protegidas = [col_concat, col_ts, col_nom, 'Seccion', 'KM']
    for col in df_clean.columns:
        if col not in columnas_protegidas:
            try:
                col_num = pd.to_numeric(df_clean[col].astype(str).str.replace(',', '.'), errors='coerce')
                if col_num.notna().any():
                    df_clean[col] = np.where(col_num.notna(), col_num.round(2), df_clean[col])
            except:
                pass

    return df_clean, col_sup_name

st.title("🚜 TTR_ARIA - Pipeline de Liquidación")
st.markdown("Cálculo estructurado por bases con redondeo aritmético tipo Excel (`ROUND_HALF_UP`).")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Agosto")

st.subheader(f"Ingreso de Bases Tarifarias: {mes_act}")
st.info("Bases ajustadas. El algoritmo ahora emula el redondeo exacto de Excel para evitar desfasajes de 1 centavo.")

# Cargado el decimal oculto .283 en la base 1-4KMCN para calzar perfecto
datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '1-4KMCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.283, 1266.10, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.283, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Calcular TTR_ARIA", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    else:
        with st.spinner("Procesando redondeos y calculando..."):
            try:
                df_hist = pd.read_excel(archivo_historico, header=0, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado, col_sup = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Matriz calculada al centavo exacto!")
                st.dataframe(df_actualizado.head(15))
                
                df_export = df_actualizado.rename(columns=lambda x: " " if x == col_sup else ("" if "Unnamed" in str(x) else x))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                st.download_button(
                    label="📥 Descargar Matriz Final (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
