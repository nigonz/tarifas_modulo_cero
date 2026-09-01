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
st.markdown("Motor TTR JN: Indexación automática y arquitectura de espejos exactos para la familia desagregada.")

st.sidebar.header("1. Cargar Historial")
archivo_historico = st.sidebar.file_uploader("Subir Archivo Histórico (.xlsx)", type=['xlsx'])

st.sidebar.header("2. Nuevo Período")
mes_act = st.sidebar.text_input("Nombre del Mes Nuevo", "Septiembre")

st.subheader(f"Ingreso de Bases Tarifarias Puras: {mes_act}")
st.info("💡 Ingresá tus 10 bases manuales. La base 1-4KMCN se indexa sola y las desagregadas espejan los resultados.")

llaves = ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']

if 'tabla_bases' not in st.session_state:
    st.session_state.tabla_bases = pd.DataFrame({
        "CONCAT Base": llaves,
        "Límite Inferior": pd.Series([None] * 10, dtype=float),
        "Límite Superior": pd.Series([None] * 10, dtype=float)
    })

tarifas_editadas = st.data_editor(st.session_state.tabla_bases, hide_index=True, use_container_width=True)

if st.button("🪄 Auto-Completar Límites Superiores"):
    df_temp = tarifas_editadas.copy()
    
    for i in range(5):
        if pd.notna(df_temp.loc[i, "Límite Inferior"]):
            df_temp.loc[i, "Límite Superior"] = df_temp.loc[i, "Límite Inferior"]
            
    for i in range(5, 9):
        if pd.notna(df_temp.loc[i+1, "Límite Inferior"]):
            df_temp.loc[i, "Límite Superior"] = df_temp.loc[i+1, "Límite Inferior"]
            
    st.session_state.tabla_bases = df_temp
    st.rerun()

st.markdown("---")

if st.button("Generar TTR_ARIA", type="primary"):
    if archivo_historico is None:
        st.warning("⚠️ Subí la matriz histórica en el panel lateral.")
    elif tarifas_editadas.isnull().values.any():
        st.error("⚠️ Faltan cargar valores. Revisá que la tabla esté completa.")
    else:
        with st.spinner("Procesando matriz..."):
            try:
                df_hist = pd.read_excel(archivo_historico, header=0)

                mapa_columnas = {str(c).strip().lower(): c for c in df_hist.columns}
                col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
                col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
                col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)
                col_km = next((mapa_columnas[c] for c in mapa_columnas if c == 'km'), 'KM')

                if not col_concat or not col_ts or not col_nom:
                    raise ValueError(f"Faltan columnas clave. Leídas: {list(df_hist.columns)}")

                columnas_protegidas = [col_concat, col_ts, col_nom, col_km, 'Seccion', 'TIPO SECCION']
                cols_historicas_meses = []
                
                for col in df_hist.columns:
                    if col not in columnas_protegidas and not str(col).startswith('Unnamed'):
                        df_hist[col] = pd.to_numeric(df_hist[col], errors='coerce').round(2)
                        cols_historicas_meses.append(col)

                dict_bases_inf = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior']))
                dict_bases_sup = dict(zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior']))

                # --- INDEXACIÓN 1-4KMCN ---
                base_1_4kmcn_dinamica = 0.0
                if len(cols_historicas_meses) > 0:
                    col_mes_anterior = cols_historicas_meses[0]
                    
                    idx_1scn = df_hist[df_hist[col_concat].astype(str).str.strip().str.upper() == '1SCN'].index
                    idx_1_4kmcn = df_hist[df_hist[col_concat].astype(str).str.strip().str.upper() == '1-4KMCN'].index
                    
                    if len(idx_1scn) > 0 and len(idx_1_4kmcn) > 0:
                        val_1scn_viejo = df_hist.loc[idx_1scn[0], col_mes_anterior]
                        val_1_4kmcn_viejo = df_hist.loc[idx_1_4kmcn[0], col_mes_anterior]
                        
                        val_1scn_nuevo = float(dict_bases_inf.get('1SCN', 0))
                        
                        if pd.notna(val_1scn_viejo) and val_1scn_viejo != 0:
                            factor_aumento = val_1scn_nuevo / float(val_1scn_viejo)
                        else:
                            factor_aumento = 1.0
                            
                        if pd.notna(val_1_4kmcn_viejo):
                            base_1_4kmcn_dinamica = float(val_1_4kmcn_viejo) * factor_aumento

                nuevos_limites_inf = []
                nuevos_limites_sup = []
                
                # Memoria caché para guardar resultados y espejarlos luego en la familia "2"
                mapa_resultados = {}

                for _, row in df_hist.iterrows():
                    concat = str(row[col_concat]).strip().upper()
                    
                    if concat in ['NAN', 'NONE', ''] or "SR" in concat:
                        nuevos_limites_inf.append(np.nan)
                        nuevos_limites_sup.append(np.nan)
                        continue

                    ts = str(row[col_ts]).strip().upper()
                    nom = str(row[col_nom]).strip().upper()
                    km_str = str(row.get(col_km, '')).strip()

                    es_caso_especial_km2 = False
                    if concat.startswith("1-4KM") and "2" in concat:
                        es_caso_especial_km2 = True

                    if es_caso_especial_km2:
                        # ARQUITECTURA DE ESPEJOS: Toma valores de la memoria en lugar de calcular
                        if km_str == '45-60':   ref_inf, ref_sup = '1-4KMCN', '5KPCN'
                        elif km_str == '60-75': ref_inf, ref_sup = '1-4KMEN', '5KPEN'
                        elif km_str == '75-90': ref_inf, ref_sup = '1-4KMEAN', '5KPEAN'
                        elif km_str == '90-150':ref_inf, ref_sup = '1-4KMCSN', '5KPCSN'
                        elif km_str == '0-3':   ref_inf, ref_sup = '1-4KMESN', '5KPESN'
                        elif km_str == '3-6':   ref_inf, ref_sup = '1-4KMEASN', '5KPEASN'
                        else:                   ref_inf, ref_sup = '1-4KMCN', '5KPCN'
                        
                        val_inf = mapa_resultados.get(ref_inf, (0, 0))[0]
                        val_sup = mapa_resultados.get(ref_sup, (0, 0))[1]
                        
                    else:
                        # CÁLCULO NORMAL PARA EL RESTO DE LA MATRIZ
                        base_key_inf = base_key_sup = "1SCN"
                        if concat.startswith("1S"): base_key_inf = base_key_sup = "1SCN"
                        elif concat.startswith("2S"): base_key_inf = base_key_sup = "2SCN"
                        elif concat.startswith("3S"): base_key_inf = base_key_sup = "3SCN"
                        elif concat.startswith("4S"): base_key_inf = base_key_sup = "4SCN"
                        elif concat.startswith("5S"): base_key_inf = base_key_sup = "5SCN"
                        elif concat.startswith("1-4KM"): base_key_inf = base_key_sup = "1-4KMCN"
                        elif concat.startswith("5KP"): base_key_inf = base_key_sup = "5KPCN"
                        elif concat.startswith("6KP"): base_key_inf = base_key_sup = "6KPCN"
                        elif concat.startswith("7KP"): base_key_inf = base_key_sup = "7KPCN"
                        elif concat.startswith("8KP"): base_key_inf = base_key_sup = "8KPCN"
                        elif concat.startswith("9KP"): base_key_inf = base_key_sup = "9KPCN"

                        if base_key_inf == "1-4KMCN":
                            base_inf = base_sup = base_1_4kmcn_dinamica
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
                        
                        # Guardamos en caché para cuando pasen las desagregadas "...2" al final de la hoja
                        if concat not in mapa_resultados:
                            mapa_resultados[concat] = (val_inf, val_sup)

                    nuevos_limites_inf.append(val_inf)
                    nuevos_limites_sup.append(val_sup)

                df_hist[f'{mes_act}'] = nuevos_limites_inf
                col_sup_name = f'{mes_act}_Sup'
                df_hist[col_sup_name] = nuevos_limites_sup

                rename_dict = {}
                espacios = 1
                for col in df_hist.columns:
                    if "Unnamed" in str(col) or str(col) == col_sup_name:
                        rename_dict[col] = " " * espacios
                        espacios += 1
                
                df_export = df_hist.rename(columns=rename_dict)

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

                st.success(f"✅ ¡Matriz de {mes_act} generada! Arquitectura de referencias y espejos aplicada.")
                st.dataframe(df_export.head(15))

                st.download_button(
                    label=f"📥 Descargar Matriz Definitiva {mes_act} (.xlsx)",
                    data=final_buffer.getvalue(),
                    file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la matriz: {e}")
