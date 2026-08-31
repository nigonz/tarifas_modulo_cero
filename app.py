import streamlit as st
import pandas as pd
import numpy as np
import io
from decimal import Decimal, ROUND_DOWN

st.set_page_config(page_title="TTR_ARIA - Módulo 0", layout="wide")

def calcular_tarifa_truncada(base, mult_ts, mult_nom):
    """Cálculo matricial puro con truncamiento estricto a 2 decimales (ROUND_DOWN)."""
    try:
        d_base = Decimal(str(base).replace(',', '.'))
        d_mult_ts = Decimal(str(mult_ts))
        d_mult_nom = Decimal(str(mult_nom))
        res = d_base * d_mult_ts * d_mult_nom
        return float(res.quantize(Decimal('0.01'), rounding=ROUND_DOWN))
    except:
        return 0.0

def procesar_nueva_matriz(df_hist, mes_nuevo_nombre, dict_bases_inf, dict_bases_sup):
    # Copiamos el DataFrame histórico tal cual viene, sin tocar ni recalcular sus columnas existentes
    df_clean = df_hist.copy()
    
    mapa_columnas = {str(c).strip().lower(): c for c in df_clean.columns}
    col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
    col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
    col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)
    col_km = next((mapa_columnas[c] for c in mapa_columnas if c == 'km'), 'KM')

    if not col_concat or not col_ts or not col_nom:
        raise ValueError(f"Faltan columnas clave en el Excel. Leídas: {list(df_clean.columns)}")

    nuevos_limites_inf = []
    nuevos_limites_sup = []

    for _, row in df_clean.iterrows():
        concat = str(row[col_concat]).strip().upper()
        
        # Saltamos celdas vacías o subtítulos/SR
        if concat in ['NAN', 'NONE', ''] or "SR" in concat:
            nuevos_limites_inf.append(np.nan)
            nuevos_limites_sup.append(np.nan)
            continue

        ts = str(row[col_ts]).strip().upper()
        nom = str(row[col_nom]).strip().upper()
        km_str = str(row.get(col_km, '')).strip()

        base_key_inf = "1SCN"
        base_key_sup = "1SCN"
        es_caso_especial_km2 = False

        if concat.startswith("1S"): base_key_inf = base_key_sup = "1SCN"
        elif concat.startswith("2S"): base_key_inf = base_key_sup = "2SCN"
        elif concat.startswith("3S"): base_key_inf = base_key_sup = "3SCN"
        elif concat.startswith("4S"): base_key_inf = base_key_sup = "4SCN"
        elif concat.startswith("5S"): base_key_inf = base_key_sup = "5SCN"
        elif concat.startswith("1-4KM"): 
            if "2" in concat:
                es_caso_especial_km2 = True
            else:
                base_key_inf = base_key_sup = "1-4KMCN"
        elif concat.startswith("5KP"): base_key_inf = base_key_sup = "5KPCN"
        elif concat.startswith("6KP"): base_key_inf = base_key_sup = "6KPCN"
        elif concat.startswith("7KP"): base_key_inf = base_key_sup = "7KPCN"
        elif concat.startswith("8KP"): base_key_inf = base_key_sup = "8KPCN"
        elif concat.startswith("9KP"): base_key_inf = base_key_sup = "9KPCN"

        if es_caso_especial_km2:
            base_inf = float(str(dict_bases_inf.get("1-4KMCN", "977.28")).replace(',', '.'))
            base_sup = float(str(dict_bases_inf.get("5KPCN", "1266.10")).replace(',', '.'))
            
            if km_str == '45-60': mult_ts, mult_nom = 1.0, 1.0
            elif km_str == '60-75': mult_ts, mult_nom = 1.25, 1.0
            elif km_str == '75-90': mult_ts, mult_nom = 1.75, 1.0
            elif km_str == '90-150': mult_ts, mult_nom = 2.0, 1.0
            elif km_str == '0-3': mult_ts, mult_nom = 1.25, 2.0
            elif km_str == '3-6': mult_ts, mult_nom = 1.75, 2.0
            else: mult_ts, mult_nom = 1.0, 1.0
        else:
            base_inf = float(str(dict_bases_inf.get(base_key_inf, "0")).replace(',', '.'))
            base_sup = float(str(dict_bases_sup.get(base_key_sup, "0")).replace(',', '.'))
            
            if concat.startswith("5KP") and ts == "E":
                base_sup = base_inf
            
            mult_ts = 1.0
            if ts == "EA": mult_ts = 1.75
            elif ts == "E": mult_ts = 1.25
            
            mult_nom = 2.0 if "SN" in nom else 1.0

        val_inf = calcular_tarifa_truncada(base_inf, mult_ts, mult_nom)
        val_sup = calcular_tarifa_truncada(base_sup, mult_ts, mult_nom)

        nuevos_limites_inf.append(val_inf)
        nuevos_limites_sup.append(val_sup)

    # Inserción de las columnas nuevas del mes con formato limpio
    df_clean[f'{mes_nuevo_nombre}'] = nuevos_limites_inf
    col_sup_name = f'{mes_nuevo_nombre}_Sup'
    df_clean[col_sup_name] = nuevos_limites_sup

    return df_clean, col_sup_name

st.title("🚜 TTR_ARIA - Pipeline de Liquidación")
st.markdown("Motor TTR ajustado: El histórico se respeta intacto y los nuevos cálculos se truncan estrictamente a 2 decimales.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Agosto")

st.subheader(f"Ingreso de Bases Tarifarias Puras: {mes_act}")
st.info("Ingresá las 11 bases. El histórico no se modifica; solo se calculan los nuevos límites truncados.")

datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '1-4KMCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.28, 1266.10, 1945.42, 2511.52, 3077.62, 3643.72],
    "Límite Superior": [742.81, 861.66, 1002.80, 1151.36, 1337.06, 977.28, 1945.42, 2511.52, 3077.62, 3643.72, 5908.12]
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Generar TTR_ARIA", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    else:
        with st.spinner("Procesando matriz..."):
            try:
                # Leemos como texto puro para preservar exactas las celdas originales del Excel maestro
                df_hist = pd.read_excel(archivo_historico, header=0, dtype=str)

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                df_actualizado, col_sup = procesar_nueva_matriz(df_hist, mes_act, dict_bases_inf, dict_bases_sup)

                st.success("✅ ¡Matriz generada manteniendo el histórico intacto y calculando el nuevo mes!")
                st.dataframe(df_actualizado.head(15))
                
                df_export = df_actualizado.rename(columns=lambda x: " " if x == col_sup else ("" if "Unnamed" in str(x) else x))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                st.download_button(
                    label="📥 Descargar Matriz Definitiva (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
