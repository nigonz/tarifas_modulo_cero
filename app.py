import streamlit as st
import pandas as pd
import numpy as np
import io
from decimal import Decimal, ROUND_DOWN
import openpyxl

st.set_page_config(page_title="TTR_ARIA - Módulo 0", layout="wide")

def calcular_tarifa(base, mult_ts, mult_nom):
    """Cálculo matricial puro con truncamiento estricto a 2 decimales."""
    try:
        d_base = Decimal(str(base))
        d_mult_ts = Decimal(str(mult_ts))
        d_mult_nom = Decimal(str(mult_nom))
        res = d_base * d_mult_ts * d_mult_nom
        return float(res.quantize(Decimal('0.01'), rounding=ROUND_DOWN))
    except:
        return 0.0

st.title("🚜 TTR_ARIA - Pipeline de Liquidación")
st.markdown("Motor TTR optimizado: Ingesta numérica, formato 0.00 y lienzo en blanco para nuevos períodos.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
# Lo seteamos por defecto en Septiembre como punto de partida
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Septiembre")

st.subheader(f"Ingreso de Bases Tarifarias Puras: {mes_act}")
st.info("Ingresá las 11 bases maestras del mes en curso.")

# Inicializar el DataFrame vacío para obligar el ingreso manual
datos_base = pd.DataFrame({
    "CONCAT Base": ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '1-4KMCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN'],
    "Límite Inferior": pd.Series([None] * 11, dtype=float),
    "Límite Superior": pd.Series([None] * 11, dtype=float)
})

tarifas_editadas = st.data_editor(datos_base, hide_index=True, use_container_width=True)

if st.button("Generar TTR_ARIA", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    elif tarifas_editadas.isnull().values.any():
        st.error("⚠️ Faltan cargar valores. Por favor, completá todos los límites inferiores y superiores antes de generar la matriz.")
    else:
        with st.spinner("Procesando matriz y formateando celdas..."):
            try:
                # 1. Lectura normal
                df_hist = pd.read_excel(archivo_historico, header=0)

                mapa_columnas = {str(c).strip().lower(): c for c in df_hist.columns}
                col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
                col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
                col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)
                col_km = next((mapa_columnas[c] for c in mapa_columnas if c == 'km'), 'KM')

                if not col_concat or not col_ts or not col_nom:
                    raise ValueError(f"Faltan columnas clave. Leídas: {list(df_hist.columns)}")

                # 2. Limpieza de floats binarios históricos (truncado visual a 2 decimales reales)
                columnas_protegidas = [col_concat, col_ts, col_nom, col_km, 'Seccion', 'TIPO SECCION']
                for col in df_hist.columns:
                    if col not in columnas_protegidas and not str(col).startswith('Unnamed'):
                        df_hist[col] = pd.to_numeric(df_hist[col], errors='coerce').round(2)

                nuevos_limites_inf = []
                nuevos_limites_sup = []

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                for _, row in df_hist.iterrows():
                    concat = str(row[col_concat]).strip().upper()
                    
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
                        # Ya no hay valores fijos. Todo depende exclusivamente del data_editor
                        base_inf = float(dict_bases_inf.get("1-4KMCN", 0))
                        base_sup = float(dict_bases_inf.get("5KPCN", 0))
                        
                        if km_str == '45-60': mult_ts, mult_nom = 1.0, 1.0
                        elif km_str == '60-75': mult_ts, mult_nom = 1.25, 1.0
                        elif km_str == '75-90': mult_ts, mult_nom = 1.75, 1.0
                        elif km_str == '90-150': mult_ts, mult_nom = 2.0, 1.0
                        elif km_str == '0-3': mult_ts, mult_nom = 1.25, 2.0
                        elif km_str == '3-6': mult_ts, mult_nom = 1.75, 2.0
                        else: mult_ts, mult_nom = 1.0, 1.0
                    else:
                        base_inf = float(dict_bases_inf.get(base_key_inf, 0))
                        base_sup = float(dict_bases_sup.get(base_key_sup, 0))
                        
                        if concat.startswith("5KP") and ts == "E":
                            base_sup = base_inf
                        
                        mult_ts = 1.0
                        if ts == "EA": mult_ts = 1.75
                        elif ts == "E": mult_ts = 1.25
                        
                        mult_nom = 2.0 if "SN" in nom else 1.0

                    val_inf = calcular_tarifa(base_inf, mult_ts, mult_nom)
                    val_sup = calcular_tarifa(base_sup, mult_ts, mult_nom)

                    nuevos_limites_inf.append(val_inf)
                    nuevos_limites_sup.append(val_sup)

                df_hist[f'{mes_act}'] = nuevos_limites_inf
                col_sup_name = f'{mes_act}_Sup'
                df_hist[col_sup_name] = nuevos_limites_sup

                # Generar espacios invisibles únicos para burlar la restricción de nombres duplicados de Pandas
                rename_dict = {}
                espacios = 1
                for col in df_hist.columns:
                    if "Unnamed" in str(col) or str(col) == col_sup_name:
                        rename_dict[col] = " " * espacios
                        espacios += 1
                
                df_export = df_hist.rename(columns=rename_dict)

                # 3. Exportar con formato 0.00
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Matriz_ARIA')

                buffer.seek(0)
                wb = openpyxl.load_workbook(buffer)
                ws = wb.active

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '0.00'

                final_buffer = io.BytesIO()
                wb.save(final_buffer)
                final_buffer.seek(0)

                st.success(f"✅ ¡Matriz de {mes_act} generada con éxito!")
                st.dataframe(df_export.head(15))

                st.download_button(
                    label=f"📥 Descargar Matriz Definitiva {mes_act} (.xlsx)",
                    data=final_buffer.getvalue(),
                    file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
