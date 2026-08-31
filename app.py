import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="TTR_ARIA - Módulo 0", layout="wide")

def truncar_estricto(valor):
    """Fuerza el truncamiento matemático estricto a 2 decimales."""
    try:
        return float(f"{float(valor):.2f}")
    except:
        return 0.0

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    df_clean = df_hist.copy()
    
    # Mapear columnas de forma segura
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
        
        # Si encuentra una fila vacía perdida, la saltea con valores nulos
        if concat in ['NAN', 'NONE', '']:
            nuevos_limites_inf.append(np.nan)
            nuevos_limites_sup.append(np.nan)
            continue

        ts = str(row[col_ts]).strip().upper()
        nom = str(row[col_nom]).strip().upper()

        # Identificación estricta de la base
        base_key = "1SCN"
        if "1-4KM" in concat: base_key = "5KPCN"
        elif "1SC" in concat: base_key = "1SCN"
        elif "2SC" in concat: base_key = "2SCN"
        elif "3SC" in concat: base_key = "3SCN"
        elif "4SC" in concat: base_key = "4SCN"
        elif "5SC" in concat: base_key = "5SCN"
        elif "1SEN" in concat or "1SEAN" in concat: base_key = "1SCN"
        elif "2SEN" in concat or "2SEAN" in concat: base_key = "2SCN"
        elif "3SEN" in concat or "3SEAN" in concat: base_key = "3SCN"
        elif "4SEN" in concat or "4SEAN" in concat: base_key = "4SCN"
        elif "5SEN" in concat or "5SEAN" in concat: base_key = "5SCN"
        elif "5KP" in concat: base_key = "5KPCN"
        elif "6KP" in concat: base_key = "6KPCN"
        elif "7KP" in concat: base_key = "7KPCN"
        elif "8KP" in concat: base_key = "8KPCN"
        elif "9KP" in concat: base_key = "9KPCN"

        base_inf = float(dict_bases_inf.get(base_key, 0))
        base_sup = float(dict_bases_sup.get(base_key, 0))

        # Multiplicadores de TTR
        mult_ts = 1.0
        if ts == "EA": mult_ts = 1.75
        elif ts == "E": mult_ts = 1.25

        mult_nom = 2.0 if "SN" in nom else 1.0

        val_inf = truncar_estricto(base_inf * mult_ts * mult_nom)
        val_sup = truncar_estricto(base_sup * mult_ts * mult_nom)

        nuevos_limites_inf.append(val_inf)
        nuevos_limites_sup.append(val_sup)

    # Pegamos las nuevas columnas con nombres temporales
    df_clean[f'{mes_nuevo_nombre}'] = nuevos_limites_inf
    col_sup_name = f'{mes_nuevo_nombre}_Sup'
    df_clean[col_sup_name] = nuevos_limites_sup

    # Barrido para redondear todo a 2 decimales sin romper los textos
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
st.markdown("Cálculo estructurado. Estructura simplificada de un solo encabezado.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Agosto")

st.subheader(f"Ingreso de Bases Tarifarias: {mes_act}")
st.info("Ingrese las 10 bases exactas. Expreso (1.25), EA (1.75) y Nominalizadas SN (x2) se calculan solas.")

datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 1266.10, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Calcular TTR_ARIA", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    else:
        with st.spinner("Procesando matriz simplificada..."):
            try:
                # Leemos directo con header=0
                df_hist = pd.read_excel(archivo_historico, header=0, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado, col_sup = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Matriz calculada perfectamente sobre la nueva estructura!")
                st.dataframe(df_actualizado.head(15))
                
                # Blanquea los títulos de las columnas de límite superior y los Unnamed para la exportación
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
