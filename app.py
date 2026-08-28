import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="TTR - Módulo 0", layout="wide")

def truncar_a_dos_decimales(serie):
    """Fuerza la conversión numérica y trunca estrictamente a 2 decimales."""
    s_num = pd.to_numeric(serie, errors='coerce')
    return np.trunc(s_num * 100) / 100

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    """Calcula el nuevo mes aplicando límites inferiores y superiores por separado."""
    df_final = df_hist.copy()
    
    # Normalizar nombres de columnas del Excel subido (eliminar espacios y pasar a mayúsculas)
    df_final.columns = [str(c).strip().upper() for c in df_final.columns]
    
    # Validar que existan las columnas obligatorias
    if 'CONCAT' not in df_final.columns or 'TS' not in df_final.columns or 'NOMINALIZACION' not in df_final.columns:
        raise ValueError("El Excel debe contener las columnas: CONCAT, TS y Nominalizacion.")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_final.iterrows():
        concat = str(row['CONCAT']).strip()
        ts = str(row['TS']).strip().upper()
        nom = str(row['NOMINALIZACION']).strip().upper()

        # 1. Identificar la llave de la tarifa base correspondiente
        base_key = None
        if "1SC" in concat: base_key = "1SCN"
        elif "2SC" in concat: base_key = "2SCN"
        elif "3SC" in concat: base_key = "3SCN"
        elif "4SC" in concat: base_key = "4SCN"
        elif "5SC" in concat: base_key = "5SCN"
        elif "1-4KM" in concat:
            if concat.endswith("2"):
                base_key = "5KPCN"
            else:
                base_key = "1SCN" # fallback seguro si no encuentra
        elif "5KP" in concat: base_key = "5KPCN"
        elif "6KP" in concat: base_key = "6KPCN"
        elif "7KP" in concat: base_key = "7KPCN"
        elif "8KP" in concat: base_key = "8KPCN"
        elif "9KP" in concat: base_key = "9KPCN"

        # Extraer ambos valores
        base_inf = dict_bases_inf.get(base_key, 0)
        base_sup = dict_bases_sup.get(base_key, 0)

        # 2. Leer multiplicadores
        mult_ts = 1.0
        if ts == "EA": mult_ts = 1.75
        elif ts == "E": mult_ts = 1.25

        mult_nom = 2.0 if nom == "SN" else 1.0

        # 3. Calcular bandas separadas
        val_inf = base_inf * mult_ts * mult_nom
        val_sup = base_sup * mult_ts * mult_nom
        
        nuevos_limites_inf.append(val_inf)
        nuevos_limites_sup.append(val_sup)

    # 4. Anexar al final del DataFrame
    df_final[f'{mes_nuevo_nombre}_Limite_Inferior'] = truncar_a_dos_decimales(nuevos_limites_inf)
    df_final[f'{mes_nuevo_nombre}_Limite_Superior'] = truncar_a_dos_decimales(nuevos_limites_sup)

    return df_final

# --- UI ---

st.title("🚜 Módulo 0: Pipeline de TTR")
st.markdown("Generación automática de matriz tarifaria consolidada.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Etiqueta Mes Nuevo", "Julio")

st.subheader(f"Ingreso de Bases Tarifarias: {mes_act}")
datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.8, 1151.36, 1337.06, 1266.1, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.8, 1151.36, 1337.06, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Calcular TTR y Anexar", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el menú lateral.")
    else:
        with st.spinner("Calculando y anexando mes..."):
            try:
                df_hist = pd.read_excel(archivo_historico, decimal=',', thousands='.')

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ Matriz calculada. Las nuevas columnas se agregaron a la derecha.")

                st.write("**Vista Previa:**")
                st.dataframe(df_actualizado.head(10))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_actualizado.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                st.download_button(
                    label="📥 Descargar Matriz Consolidada (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Matriz_TTR_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
