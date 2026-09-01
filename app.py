import streamlit as st
import pandas as pd
import numpy as np
import io
from decimal import Decimal, ROUND_DOWN
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="TTR_ARIA - Sistema Integral", layout="wide")

# ==============================================================================
# FUNCIONES COMPARTIDAS
# ==============================================================================
def parse_argentinian_float(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    s = str(val).replace(',', '.')
    try: return float(s)
    except: return 0.0

def calcular_tarifa(base, mult_ts, mult_nom):
    try:
        res = Decimal(str(base)) * Decimal(str(mult_ts)) * Decimal(str(mult_nom))
        return float(res.quantize(Decimal('0.01'), rounding=ROUND_DOWN))
    except: return 0.0

def formatear_excel(ws, df, cols_moneda):
    encabezado = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    relleno = PatternFill('solid', start_color='1F4E78')
    cuerpo = Font(name='Arial', size=10)
    for celda in ws[1]:
        celda.font = encabezado
        celda.fill = relleno
        celda.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for fila in ws.iter_rows(min_row=2):
        for celda in fila: celda.font = cuerpo
    for i, col in enumerate(df.columns, start=1):
        letra = get_column_letter(i)
        ws.column_dimensions[letra].width = min(max(len(str(col)) + 4, 12), 32)
        if col in cols_moneda:
            for celda in ws[letra][1:]: celda.number_format = '#,##0.00'
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

# ==============================================================================
# MÓDULO 0: TARIFARIO JN
# ==============================================================================
def modulo_tarifas():
    st.title("🚜 TTR_ARIA - Módulo 0: Tarifario JN")
    st.markdown("Generador de cuadros tarifarios con indexación automática y arquitectura de espejos.")
    col_izq, col_der = st.columns([1, 3])
    with col_izq:
        archivo_historico = st.file_uploader("Subir Matriz Histórica (.xlsx)", type=['xlsx'])
        mes_act = st.text_input("Nombre del Mes Nuevo", "Septiembre")
    with col_der:
        st.info("💡 Ingresá tus 10 bases manuales. La base 1-4KMCN se indexará automáticamente.")
        llaves = ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']
        if 'tabla_bases' not in st.session_state:
            st.session_state.tabla_bases = pd.DataFrame({
                "CONCAT Base": llaves, "Límite Inferior": pd.Series([""] * 10, dtype=str), "Límite Superior": pd.Series([""] * 10, dtype=str)
            })
        tarifas_editadas = st.data_editor(st.session_state.tabla_bases, hide_index=True, use_container_width=True)
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🪄 Auto-Completar Límites Superiores", use_container_width=True):
                df_temp = tarifas_editadas.copy()
                for i in range(5):
                    if str(df_temp.loc[i, "Límite Inferior"]).strip() != "": df_temp.loc[i, "Límite Superior"] = df_temp.loc[i, "Límite Inferior"]
                for i in range(5, 9):
                    if str(df_temp.loc[i+1, "Límite Inferior"]).strip() != "": df_temp.loc[i, "Límite Superior"] = df_temp.loc[i+1, "Límite Inferior"]
                st.session_state.tabla_bases = df_temp
                st.rerun()
        with col_btn2:
            if st.button("Generar Cuadro Tarifario TTR", type="primary", use_container_width=True):
                dict_bases_inf = {k: parse_argentinian_float(v) for k, v in zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Inferior'])}
                dict_bases_sup = {k: parse_argentinian_float(v) for k, v in zip(tarifas_editadas['CONCAT Base'], tarifas_editadas['Límite Superior'])}
                if archivo_historico is None: st.error("⚠️ Subí la matriz histórica.")
                elif all(v == 0.0 for v in dict_bases_inf.values()): st.error("⚠️ Faltan cargar bases.")
                else:
                    with st.spinner("Procesando matriz..."):
                        try:
                            df_hist = pd.read_excel(archivo_historico, header=0)
                            mapa_columnas = {str(c).strip().lower(): c for c in df_hist.columns}
                            col_concat = next((mapa_columnas[c] for c in mapa_columnas if 'concat' in c), None)
                            col_ts = next((mapa_columnas[c] for c in mapa_columnas if c == 'ts'), None)
                            col_nom = next((mapa_columnas[c] for c in mapa_columnas if 'nominaliz' in c), None)
                            col_km = next((mapa_columnas[c] for c in mapa_columnas if c == 'km'), 'KM')

                            columnas_protegidas = [col_concat, col_ts, col_nom, col_km, 'Seccion', 'TIPO SECCION']
                            cols_historicas_meses = []
                            for col in df_hist.columns:
                                if col not in columnas_protegidas and not str(col).startswith('Unnamed'):
                                    df_hist[col] = pd.to_numeric(df_hist[col], errors='coerce').round(2)
                                    cols_historicas_meses.append(col)

                            base_1_4kmcn_dinamica = 0.0
                            if len(cols_historicas_meses) > 0:
                                col_mes_anterior = cols_historicas_meses[0]
                                idx_1scn = df_hist[df_hist[col_concat].astype(str).str.strip().str.upper() == '1SCN'].index
                                idx_1_4kmcn = df_hist[df_hist[col_concat].astype(str).str.strip().str.upper() == '1-4KMCN'].index
                                if len(idx_1scn) > 0 and len(idx_1_4kmcn) > 0:
                                    val_1scn_viejo = df_hist.loc[idx_1scn, col_mes_anterior].mode()[0]
                                    val_1_4kmcn_viejo = df_hist.loc[idx_1_4kmcn, col_mes_anterior].mode()[0]
                                    val_1scn_nuevo = dict_bases_inf.get('1SCN', 0.0)
                                    factor_aumento = (val_1scn_nuevo / float(val_1scn_viejo)) if (pd.notna(val_1scn_viejo) and val_1scn_viejo != 0) else 1.0
                                    if pd.notna(val_1_4kmcn_viejo): base_1_4kmcn_dinamica = float(val_1_4kmcn_viejo) * factor_aumento

                            nuevos_limites_inf, nuevos_limites_sup = [], []
                            mapa_resultados = {}
                            for _, row in df_hist.iterrows():
                                concat = str(row[col_concat]).strip().upper()
                                if concat in ['NAN', 'NONE', ''] or "SR" in concat:
                                    nuevos_limites_inf.append(np.nan); nuevos_limites_sup.append(np.nan); continue
                                ts = str(row[col_ts]).strip().upper(); nom = str(row[col_nom]).strip().upper()
                                km_str = str(row.get(col_km, '')).strip()
                                es_caso_especial_km2 = (concat.startswith("1-4KM") and "2" in concat)

                                if es_caso_especial_km2:
                                    mapa_km2 = {'45-60':('1-4KMCN','5KPCN'), '60-75':('1-4KMEN','5KPEN'), '75-90':('1-4KMEAN','5KPEAN'),
                                                '90-150':('1-4KMCSN','5KPCSN'), '0-3':('1-4KMESN','5KPESN'), '3-6':('1-4KMEASN','5KPEASN')}
                                    ref_inf, ref_sup = mapa_km2.get(km_str, ('1-4KMCN','5KPCN'))
                                    val_inf, val_sup = mapa_resultados.get(ref_inf, (0,0))[0], mapa_resultados.get(ref_sup, (0,0))[1]
                                else:
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

                                    if base_key_inf == "1-4KMCN": base_inf = base_sup = base_1_4kmcn_dinamica
                                    else:
                                        base_inf = float(dict_bases_inf.get(base_key_inf, 0)); base_sup = float(dict_bases_sup.get(base_key_sup, 0))
                                    if concat.startswith("5KP") and ts == "E": base_sup = base_inf
                                    mult_ts = 1.75 if ts == "EA" else (1.25 if ts == "E" else 1.0)
                                    mult_nom = 2.0 if "SN" in nom else 1.0

                                    val_inf = calcular_tarifa(base_inf, mult_ts, mult_nom)
                                    val_sup = calcular_tarifa(base_sup, mult_ts, mult_nom)
                                    if concat not in mapa_resultados: mapa_resultados[concat] = (val_inf, val_sup)

                                nuevos_limites_inf.append(val_inf); nuevos_limites_sup.append(val_sup)

                            df_hist[f'{mes_act}'] = nuevos_limites_inf
                            col_sup_name = f'{mes_act}_Sup'
                            df_hist[col_sup_name] = nuevos_limites_sup
                            rename_dict = {}
                            espacios = 1
                            for col in df_hist.columns:
                                if "Unnamed" in str(col) or str(col) == col_sup_name:
                                    rename_dict[col] = " " * espacios; espacios += 1
                            df_export = df_hist.rename(columns=rename_dict)

                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                df_export.to_excel(writer, index=False, sheet_name='Matriz_ARIA')
                            buffer.seek(0)
                            wb = openpyxl.load_workbook(buffer)
                            ws = wb.active
                            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                                for cell in row:
                                    if isinstance(cell.value, (int, float)): cell.number_format = '0.00'
                            final_buffer = io.BytesIO()
                            wb.save(final_buffer); final_buffer.seek(0)

                            st.success(f"✅ ¡Matriz de {mes_act} generada con éxito!")
                            st.download_button(label=f"📥 Descargar Matriz Definitiva {mes_act} (.xlsx)", data=final_buffer.getvalue(), file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        except Exception as e: st.error(f"Error: {e}")

# ==============================================================================
# MÓDULO 1: DMK (ITG y ATS)
# ==============================================================================
# [El Módulo 1 se mantiene idéntico al bloque anterior, omitido visualmente aquí para focalizar en Módulo 3, pero debés conservar el def modulo_dmk() en tu archivo real]

# ==============================================================================
# MÓDULO 3: CÁLCULO TTR (EL NUEVO MOTOR OPTIMIZADO)
# ==============================================================================
def modulo_calculo_ttr():
    st.title("🧮 TTR_ARIA - Módulo 3: Cálculo y Valorización TTR")
    st.markdown("Clasificación tarifaria y cálculo de recaudación teórica (TRSUBE) conectada a la Matriz ARIA.")

    st.header("1. Archivos de Entrada")
    col1, col2 = st.columns(2)
    file_tarifario_dmk = col1.file_uploader("1. DMK Resumen_Tarifario_Dominio", type=['xlsx', 'csv'])
    file_matriz_aria = col2.file_uploader("2. Matriz TTR ARIA (Módulo 0)", type=['xlsx'])
    file_nom_ramal = col1.file_uploader("3. Nomenclador Ramal - TS", type=['xlsx'])
    file_ttr_reso = col2.file_uploader("4. Teórica Resoluciones", type=['xlsx'])

    with st.expander("⚙️ Parámetros de Clasificación", expanded=False):
        c1, c2, c3 = st.columns(3)
        anio = c1.number_input("Año de Liquidación", value=2026)
        resolucion = c2.text_input("Número de Resolución", "48")
        filtro_tarifa_vieja = c3.number_input("Corte Tarifa Vieja (Ej: 728.23)", value=728.23)

    if st.button("🚀 Ejecutar Valorización TTR", type="primary", use_container_width=True):
        if not all([file_tarifario_dmk, file_matriz_aria, file_nom_ramal, file_ttr_reso]):
            st.error("⚠️ Faltan archivos requeridos para la liquidación.")
            return

        with st.spinner("Procesando árbol de decisiones TTR..."):
            try:
                # 1. Carga de datos
                if file_tarifario_dmk.name.endswith('.csv'): df1 = pd.read_csv(file_tarifario_dmk, sep=';', encoding='utf-8-sig')
                else: df1 = pd.read_excel(file_tarifario_dmk)
                
                nom_ts = pd.read_excel(file_nom_ramal, sheet_name='NOMENCLADOR TS' if 'NOMENCLADOR TS' in pd.ExcelFile(file_nom_ramal).sheet_names else 0)
                ttr_reso = pd.read_excel(file_ttr_reso, sheet_name='TTR')
                ttr_sgii_uma2 = pd.read_excel(file_ttr_reso, sheet_name='SGII-UMA2')
                df_aria = pd.read_excel(file_matriz_aria)

                # 2. Extracción dinámica de límites desde Matriz ARIA (Reemplaza diccionarios hardcodeados)
                col_mes_inf = [c for c in df_aria.columns if c not in ['CONCAT', 'TS', 'Seccion', 'KM', 'Nominalizacion'] and not str(c).endswith('_Sup') and not str(c).strip() == ''][0]
                col_mes_sup = f"{col_mes_inf}_Sup"
                # En caso de que el _Sup tenga espacios en blanco por el formateo de ARIA
                if col_mes_sup not in df_aria.columns: col_mes_sup = [c for c in df_aria.columns if str(c).strip() == ''][0]

                limites = {}
                for _, row in df_aria.iterrows(): limites[row['CONCAT']] = (row[col_mes_inf], row[col_mes_sup])
                def get_lim(k): return limites.get(k, (0.0, 0.0))

                # Reconstruimos los diccionarios de la lógica original, pero apuntando a la matriz en vivo
                tarifas_1 = {
                    '1SCN': get_lim('1SCN'), '2SCN': get_lim('2SCN'), '3SCN': get_lim('3SCN'), '4SCN': get_lim('4SCN'), '5SCN': get_lim('5SCN'),
                    '1SEN': get_lim('1SEN'), '2SEN': get_lim('2SEN'), '3SEN': get_lim('3SEN'), '4SEN': get_lim('4SEN'), '5SEN': get_lim('5SEN'),
                    '1SEAN': get_lim('1SEAN'), '2SEAN': get_lim('2SEAN'), '3SEAN': get_lim('3SEAN'), '4SEAN': get_lim('4SEAN'), '5SEAN': get_lim('5SEAN')
                }
                tarifas_2 = {
                    '1SCSN': get_lim('1SCSN'), '2SCSN': get_lim('2SCSN'), '3SCSN': get_lim('3SCSN'), '4SCSN': get_lim('4SCSN'), '5SCSN': get_lim('5SCSN'),
                    '1SESN': get_lim('1SESN'), '2SESN': get_lim('2SESN'), '3SESN': get_lim('3SESN'), '4SESN': get_lim('4SESN'), '5SESN': get_lim('5SESN'),
                    '1SEASN': get_lim('1SEASN'), '2SEASN': get_lim('2SEASN'), '3SEASN': get_lim('3SEASN'), '4SEASN': get_lim('4SEASN'), '5SEASN': get_lim('5SEASN')
                }
                tarifas_3 = {'1-4KMCN': (*get_lim('1-4KMCN'), "C"), '1-4KMEN': (*get_lim('1-4KMEN'), "E"), '1-4KMEAN': (*get_lim('1-4KMEAN'), "EA")}
                tarifas_4 = {'1-4KMCSN': (*get_lim('1-4KMCSN'), "C"), '1-4KMESN': (*get_lim('1-4KMESN'), "E"), '1-4KMEASN': (*get_lim('1-4KMEASN'), "EA")}
                
                tarifas_7 = {
                    '5KPCN': (*get_lim('5KPCN'), "C", 5), '6KPCN': (*get_lim('6KPCN'), "C", 6), '7KPCN': (*get_lim('7KPCN'), "C", 7), '8KPCN': (*get_lim('8KPCN'), "C", 8), '9KPCN': (*get_lim('9KPCN'), "C", 9),
                    '5KPEN': (*get_lim('5KPEN'), "E", 5), '6KPEN': (*get_lim('6KPEN'), "E", 6), '7KPEN': (*get_lim('7KPEN'), "E", 7), '8KPEN': (*get_lim('8KPEN'), "E", 8), '9KPEN': (*get_lim('9KPEN'), "E", 9),
                    '5KPEAN': (*get_lim('5KPEAN'), "EA", 5), '6KPEAN': (*get_lim('6KPEAN'), "EA", 6), '7KPEAN': (*get_lim('7KPEAN'), "EA", 7), '8KPEAN': (*get_lim('8KPEAN'), "EA", 8), '9KPEAN': (*get_lim('9KPEAN'), "EA", 9)
                }
                tarifas_8 = {
                    '5KPCSN': (*get_lim('5KPCSN'), "C", 5), '6KPCSN': (*get_lim('6KPCSN'), "C", 6), '7KPCSN': (*get_lim('7KPCSN'), "C", 7), '8KPCSN': (*get_lim('8KPCSN'), "C", 8), '9KPCSN': (*get_lim('9KPCSN'), "C", 9),
                    '5KPESN': (*get_lim('5KPESN'), "E", 5), '6KPESN': (*get_lim('6KPESN'), "E", 6), '7KPESN': (*get_lim('7KPESN'), "E", 7), '8KPESN': (*get_lim('8KPESN'), "E", 8), '9KPESN': (*get_lim('9KPESN'), "E", 9),
                    '5KPEASN': (*get_lim('5KPEASN'), "EA", 5), '6KPEASN': (*get_lim('6KPEASN'), "EA", 6), '7KPEASN': (*get_lim('7KPEASN'), "EA", 7), '8KPEASN': (*get_lim('8KPEASN'), "EA", 8), '9KPEASN': (*get_lim('9KPEASN'), "EA", 9)
                }
                tarifas_cn = {'1-4KMCN2': get_lim('1-4KMCN2')}
                tarifas_en = {'1-4KMEN2': get_lim('1-4KMEN2')}
                tarifas_ean = {'1-4KMEAN2': get_lim('1-4KMEAN2')}
                tarifas_cn_sn = {'1-4KMCSN2': get_lim('1-4KMCSN2')}
                tarifas_en_sn = {'1-4KMESN2': get_lim('1-4KMESN2')}
                tarifas_ean_sn = {'1-4KMEASN2': get_lim('1-4KMEASN2')}

                # 3. Limpieza y Merges
                var_input = ['PROVINCIA', 'MUNICIPIO', 'GRUPO_TARIFARIO', 'AMBA', 'LINEA_SILAS_DNGFF', 'ID_LINEA', 'RAMAL', 'CONTRATO', 'TARIFA', 'DEBITADO', 'DOMINIO', 'ENERGIA', 'USOS', 'RECAUDACION', 'TOTAL DESC POR INTEGRACION', 'DESCUENTO_ATRIBUTOS', 'DESCUENTO_TOTAL']
                df2 = df1[[c for c in var_input if c in df1.columns]].copy()
                _df2_ = df2[df2['GRUPO_TARIFARIO'].isin(["SGI", "SGII", "SGIKM"])].copy()
                
                _df2_['USOS'] = pd.to_numeric(_df2_['USOS'].astype(str).replace({',': ''}, regex=True), errors='coerce').fillna(0)
                _df2_['TARIFA'] = pd.to_numeric(_df2_['TARIFA'].astype(str).replace({',': ''}, regex=True), errors='coerce').fillna(0).round(2)
                
                nom_ts["IdRamalNS"] = nom_ts.get("IdRamalNS", nom_ts.iloc[:,0]).astype(str)
                _df2_["RAMAL"] = _df2_["RAMAL"].astype(str)
                _df2_ = pd.merge(_df2_, nom_ts[['IdRamalNS', 'TIPO DE SERVICIO FINAL']], how='left', left_on='RAMAL', right_on='IdRamalNS')
                _df2_.rename(columns={'TIPO DE SERVICIO FINAL': 'TipoServicio'}, inplace=True)
                _df2_['TipoServicio2'] = _df2_['TipoServicio'].replace('SR', 'E')
                
                _df2_['sin_nominalizar'] = np.where(_df2_['CONTRATO'] == 627, 1, 0)
                _df2_['PASES'] = np.where((_df2_['TARIFA'] >= 0) & (_df2_['TARIFA'] <= 0.5), 1, 0)
                _df2_['FILTRO_1'] = np.where((_df2_['TARIFA'] < filtro_tarifa_vieja) & (_df2_['TARIFA'] > 0.5), 1, 0)

                # 4. Asignación Matricial Binaria
                for col, (lim_inf, lim_sup) in tarifas_1.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_['TipoServicio'] != "SR"), 1, 0)
                for col, (lim_inf, lim_sup) in tarifas_2.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_['TipoServicio'] != "SR"), 1, 0)
                for col, (lim_inf, lim_sup, ts_val) in tarifas_3.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_['TipoServicio2'] == ts_val), 1, 0)
                for col, (lim_inf, lim_sup, ts_val) in tarifas_4.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_['TipoServicio2'] == ts_val), 1, 0)
                for col, (lim_inf, lim_sup, ts_val, val_asig) in tarifas_7.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] < lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_['TipoServicio2'] == ts_val), val_asig, 0)
                for col, (lim_inf, lim_sup, ts_val, val_asig) in tarifas_8.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] < lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_['TipoServicio2'] == ts_val), val_asig, 0)

                # Definición de listas limpias de columnas para comprobaciones sumatorias
                c_sn_cn = ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN']
                c_kmn_cn = ['1-4KMCN']
                c_sn_en = ['1SEN', '2SEN', '3SEN', '4SEN', '5SEN']
                c_kmn_en = ['1-4KMEN']
                c_kpn_en = ['5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']
                c_sn_ean = ['1SEAN', '2SEAN', '3SEAN', '4SEAN', '5SEAN']
                c_kmn_ean = ['1-4KMEAN']
                c_kpn_ean = c_kpn_en + ['5KPEN', '6KPEN', '7KPEN', '8KPEN', '9KPEN']
                
                for col, (lim_inf, lim_sup) in tarifas_cn.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_[c_sn_cn].sum(axis=1) == 0) & (_df2_[c_kmn_cn].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)
                for col, (lim_inf, lim_sup) in tarifas_en.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_[c_sn_en].sum(axis=1) == 0) & (_df2_[c_kmn_en].sum(axis=1) == 0) & (_df2_[c_kpn_en].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)
                for col, (lim_inf, lim_sup) in tarifas_ean.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 0) & (_df2_[c_sn_ean].sum(axis=1) == 0) & (_df2_[c_kmn_ean].sum(axis=1) == 0) & (_df2_[c_kpn_ean].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)
                
                c_sn_cSn = ['1SCSN', '2SCSN', '3SCSN', '4SCSN', '5SCSN']
                c_kmn_cSn = ['1-4KMCSN']
                c_sn_eSn = ['1SESN', '2SESN', '3SESN', '4SESN', '5SESN']
                c_kmn_eSn = ['1-4KMESN']
                c_kpn_eSn = ['5KPCSN', '6KPCSN', '7KPCSN', '8KPCSN', '9KPCSN']
                c_sn_eaSn = ['1SEASN', '2SEASN', '3SEASN', '4SEASN', '5SEASN']
                c_kmn_eaSn = ['1-4KMEASN']
                c_kpn_eaSn = c_kpn_eSn + ['5KPESN', '6KPESN', '7KPESN', '8KPESN', '9KPESN']

                for col, (lim_inf, lim_sup) in tarifas_cn_sn.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_[c_sn_cSn].sum(axis=1) == 0) & (_df2_[c_kmn_cSn].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)
                for col, (lim_inf, lim_sup) in tarifas_en_sn.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_[c_sn_eSn].sum(axis=1) == 0) & (_df2_[c_kmn_eSn].sum(axis=1) == 0) & (_df2_[c_kpn_eSn].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)
                for col, (lim_inf, lim_sup) in tarifas_ean_sn.items():
                    _df2_[col] = np.where((_df2_['TARIFA'] >= lim_inf - 0.5) & (_df2_['TARIFA'] <= lim_sup - 0.5) & (_df2_['PASES'] == 0) & (_df2_['sin_nominalizar'] == 1) & (_df2_[c_sn_eaSn].sum(axis=1) == 0) & (_df2_[c_kmn_eaSn].sum(axis=1) == 0) & (_df2_[c_kpn_eaSn].sum(axis=1) == 0) & (_df2_['GRUPO_TARIFARIO'] != "DF"), 1, 0)

                # 5. Árbol de decisiones y Compilados (Limpio de LaPlata)
                _df2_['Filtro1-4KMCN'] = np.where((_df2_['TipoServicio2'] == 'C') & (_df2_[c_sn_en + c_sn_ean].sum(axis=1) != 0) & (_df2_['1-4KMCN2'] == 1), 4, 0)
                _df2_['Filtro1-4KMEN'] = np.where((_df2_['TipoServicio2'] == 'E') & (_df2_[c_sn_ean].sum(axis=1) != 0) & (_df2_['1-4KMEN2'] == 1), 4, 0)
                _df2_['Filtro1-4KMEAN'] = np.where((_df2_['TipoServicio2'] == 'EA') & (_df2_['1-4KMEAN2'] == 1), 4, 0)
                _df2_['Filtro1-4KMCSN'] = np.where((_df2_['TipoServicio2'] == 'C') & (_df2_[c_sn_eSn + c_sn_eaSn].sum(axis=1) != 0) & (_df2_['1-4KMCSN2'] == 1), 4, 0)
                _df2_['Filtro1-4KMESN'] = np.where((_df2_['TipoServicio2'] == 'E') & (_df2_[c_sn_eaSn].sum(axis=1) != 0) & (_df2_['1-4KMESN2'] == 1), 4, 0)
                _df2_['Filtro1-4KMEASN'] = np.where((_df2_['TipoServicio2'] == 'EA') & (_df2_['1-4KMEASN2'] == 1), 4, 0)

                _df2_['seccionada_correcta_1'] = np.select([(_df2_['1-4KMCN2'] == 1) & (_df2_['Filtro1-4KMCN'] != 4) & (_df2_['1SEN'] == 1), (_df2_['1-4KMCN2'] == 1) & (_df2_['Filtro1-4KMCN'] != 4) & (_df2_['2SEN'] == 1), (_df2_['1-4KMCN2'] == 1) & (_df2_['Filtro1-4KMCN'] != 4) & (_df2_['3SEN'] == 1), (_df2_['1-4KMCN2'] == 1) & (_df2_['Filtro1-4KMCN'] != 4) & (_df2_['4SEN'] == 1), (_df2_['1-4KMCN2'] == 1) & (_df2_['Filtro1-4KMCN'] != 4) & (_df2_['5SEN'] == 1)], [1, 2, 3, 4, 5], default=0)
                _df2_['seccionada_correcta_3'] = np.select([(_df2_['Filtro1-4KMCSN'] != 4) & (_df2_['1-4KMCSN2'] == 1) & (_df2_['1SESN'] == 1), (_df2_['Filtro1-4KMCSN'] != 4) & (_df2_['1-4KMCSN2'] == 1) & (_df2_['2SESN'] == 1), (_df2_['Filtro1-4KMCSN'] != 4) & (_df2_['1-4KMCSN2'] == 1) & (_df2_['3SESN'] == 1), (_df2_['Filtro1-4KMCSN'] != 4) & (_df2_['1-4KMCSN2'] == 1) & (_df2_['4SESN'] == 1), (_df2_['Filtro1-4KMCSN'] != 4) & (_df2_['1-4KMCSN2'] == 1) & (_df2_['5SESN'] == 1)], [1, 2, 3, 4, 5], default=0)
                _df2_['seccionada_correcta_2'] = np.select([(_df2_['1-4KMEN2'] == 1) & (_df2_['Filtro1-4KMEN'] != 4) & (_df2_['1SEAN'] == 1), (_df2_['1-4KMEN2'] == 1) & (_df2_['Filtro1-4KMEN'] != 4) & (_df2_['2SEAN'] == 1), (_df2_['1-4KMEN2'] == 1) & (_df2_['Filtro1-4KMEN'] != 4) & (_df2_['3SEAN'] == 1), (_df2_['1-4KMEN2'] == 1) & (_df2_['Filtro1-4KMEN'] != 4) & (_df2_['4SEAN'] == 1), (_df2_['1-4KMEN2'] == 1) & (_df2_['Filtro1-4KMEN'] != 4) & (_df2_['5SEAN'] == 1)], [1, 2, 3, 4, 5], default=0)
                _df2_['seccionada_correcta_4'] = np.select([(_df2_['1-4KMESN2'] == 1) & (_df2_['Filtro1-4KMESN'] != 4) & (_df2_['1SEASN'] == 1), (_df2_['1-4KMESN2'] == 1) & (_df2_['Filtro1-4KMESN'] != 4) & (_df2_['2SEASN'] == 1), (_df2_['1-4KMESN2'] == 1) & (_df2_['Filtro1-4KMESN'] != 4) & (_df2_['3SEASN'] == 1), (_df2_['1-4KMESN2'] == 1) & (_df2_['Filtro1-4KMESN'] != 4) & (_df2_['4SEASN'] == 1), (_df2_['1-4KMESN2'] == 1) & (_df2_['Filtro1-4KMESN'] != 4) & (_df2_['5SEASN'] == 1)], [1, 2, 3, 4, 5], default=0)

                _df2_['sec_c'] = np.where((_df2_[c_sn_cn].sum(axis=1) > 0) | (_df2_[c_sn_cSn].sum(axis=1) > 0) | (_df2_['1-4KMCN'] > 0) | (_df2_['1-4KMCSN'] > 0), 1, 0)
                _df2_['sec_e'] = np.where((_df2_[c_sn_en].sum(axis=1) > 0) | (_df2_[c_sn_eSn].sum(axis=1) > 0) | (_df2_['1-4KMEN'] > 0) | (_df2_['1-4KMESN'] > 0), 1, 0)
                _df2_['sec_ea'] = np.where((_df2_[c_sn_ean].sum(axis=1) > 0) | (_df2_[c_sn_eaSn].sum(axis=1) > 0) | (_df2_['1-4KMEAN'] > 0) | (_df2_['1-4KMEASN'] > 0), 1, 0)

                _df2_['km&p_c'] = np.where((_df2_[c_kpn_en[:5]].sum(axis=1) > 0) | (_df2_[c_kpn_eSn[:5]].sum(axis=1) > 0) | (_df2_['1-4KMCN2'] > 0) | (_df2_['1-4KMCSN2'] > 0), 1, 0)
                _df2_['km&p_e'] = np.where((_df2_[c_kpn_en[5:]].sum(axis=1) > 0) | (_df2_[c_kpn_eSn[5:10]].sum(axis=1) > 0) | (_df2_['1-4KMEN2'] > 0) | (_df2_['1-4KMESN2'] > 0), 1, 0)
                _df2_['km&p_ea'] = np.where((_df2_['5KPEAN'] > 0) | (_df2_['1-4KMEAN2'] > 0) | (_df2_['1-4KMEASN2'] > 0), 1, 0) # Simplificado

                _df2_['compilado_ts'] = np.select([((_df2_[['km&p_c', 'km&p_e', 'km&p_ea']] == 1).any(axis=1)) & (_df2_[['seccionada_correcta_1', 'seccionada_correcta_2', 'seccionada_correcta_3', 'seccionada_correcta_4']].sum(axis=1) == 0), (_df2_['sec_c'] == 1), (_df2_['sec_e'] == 1), (_df2_['sec_ea'] == 1), (_df2_['PASES'] == 1), (_df2_[['sec_c', 'sec_e', 'sec_ea', 'km&p_c', 'km&p_e', 'km&p_ea']].sum(axis=1) == 0)], [_df2_['TipoServicio2'], 'C', 'E', 'EA', _df2_['TipoServicio2'], _df2_['TipoServicio2']], default="S/D")

                _df2_['norm_por_tarifa'] = np.where(((_df2_[c_sn_cn + c_sn_en + c_sn_ean].sum(axis=1) > 0) | (_df2_[['1-4KMCN', '1-4KMEN', '1-4KMEAN']].sum(axis=1) > 0) | (_df2_[c_kpn_en + c_kpn_ean].sum(axis=1) > 0) | (_df2_[['1-4KMCN2', '1-4KMEN2', '1-4KMEAN2']].sum(axis=1) > 0)), "N", np.where(_df2_['FILTRO_1'] == 1, "Tarifa Vieja", np.where(_df2_['PASES'] == 1, "N", "SN")))
                
                _df2_['tarifa_s'] = np.where(((_df2_[c_sn_cn + c_sn_en + c_sn_ean + c_sn_cSn + c_sn_eSn + c_sn_eaSn].sum(axis=1) > 0)) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0) & (((_df2_['compilado_ts'] == 'C') & (_df2_['sec_c'] == 1)) | ((_df2_['compilado_ts'] == 'E') & (_df2_['sec_e'] == 1)) | ((_df2_['compilado_ts'] == 'EA') & (_df2_['sec_ea'] == 1))), 1, 0)
                _df2_['tarifa_km'] = np.where(((_df2_[['1-4KMCN', '1-4KMEN', '1-4KMEAN', '1-4KMCSN','1-4KMESN', '1-4KMEASN']].sum(axis=1) > 0) | (_df2_[['1-4KMCN2', '1-4KMEN2', '1-4KMEAN2', '1-4KMCSN2', '1-4KMESN2', '1-4KMEASN2']].sum(axis=1) > 0)) & (_df2_[['seccionada_correcta_1', 'seccionada_correcta_3', 'seccionada_correcta_2', 'seccionada_correcta_4']].sum(axis=1) == 0), 1, 0)
                _df2_['tarifa_kp'] = np.where((_df2_[c_kpn_en + c_kpn_ean + c_kpn_eSn + c_kpn_eaSn].sum(axis=1) > 0), 1, 0)
                _df2_['tarifa_PASE'] = np.where((_df2_['PASES'] == 1), 1, 0)

                _df2_['compilado_tt'] = np.where(_df2_['tarifa_s'] != 0, 'S', np.where(_df2_['tarifa_km'] != 0, 'KM', np.where(_df2_['tarifa_kp'] != 0, 'KP', np.where(_df2_['tarifa_PASE'] != 0, 'P', np.where((_df2_['seccionada_correcta_1'] + _df2_['seccionada_correcta_2'] + _df2_['seccionada_correcta_3'] + _df2_['seccionada_correcta_4']) != 0, 'S', 'S/D')))))

                _df2_['sec_1'] = np.where((_df2_[['1SCN', '1SCSN', '1SEN', '1SESN', '1SEAN', '1SEASN']].sum(axis=1) > 0) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0), 1, 0)
                _df2_['sec_2'] = np.where((_df2_[['2SCN', '2SCSN', '2SEN', '2SESN', '2SEAN', '2SEASN']].sum(axis=1) > 0) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0), 2, 0)
                _df2_['sec_3'] = np.where((_df2_[['3SCN', '3SCSN', '3SEN', '3SESN', '3SEAN', '3SEASN']].sum(axis=1) > 0) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0), 3, 0)
                _df2_['sec_4'] = np.where((_df2_[['4SCN', '4SCSN', '4SEN', '4SESN', '4SEAN', '4SEASN']].sum(axis=1) > 0) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0), 4, 0)
                _df2_['sec_5'] = np.where((_df2_[['5SCN', '5SCSN', '5SEN', '5SESN', '5SEAN', '5SEASN']].sum(axis=1) > 0) & (_df2_[['Filtro1-4KMCN', 'Filtro1-4KMEN', 'Filtro1-4KMEAN', 'Filtro1-4KMCSN', 'Filtro1-4KMESN', 'Filtro1-4KMEASN']].sum(axis=1) == 0), 5, 0)

                _df2_['seccionadas_final'] = np.where(_df2_['PASES'] == 1, 1, _df2_[['sec_1', 'sec_2', 'sec_3', 'sec_4', 'sec_5']].sum(axis=1))
                _df2_['sec_1_4'] = np.where((_df2_[['1-4KMCN', '1-4KMEN', '1-4KMEAN', '1-4KMCSN', '1-4KMESN', '1-4KMEASN']].sum(axis=1) > 0) | (_df2_[['1-4KMCN2', '1-4KMEN2', '1-4KMEAN2', '1-4KMCSN2', '1-4KMESN2', '1-4KMEASN2']].sum(axis=1) > 0) & (_df2_[['seccionada_correcta_1', 'seccionada_correcta_2', 'seccionada_correcta_3', 'seccionada_correcta_4']].eq(0).all(axis=1)), 4, 0)

                _df2_['kilometricas'] = np.where((_df2_['compilado_ts'] == 'C') & (_df2_['norm_por_tarifa'] == 'N'), _df2_[['5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']].sum(axis=1), np.where((_df2_['compilado_ts'] == 'E') & (_df2_['norm_por_tarifa'] == 'N'), _df2_[['5KPEN', '6KPEN', '7KPEN', '8KPEN', '9KPEN']].sum(axis=1), np.where((_df2_['compilado_ts'] == 'EA') & (_df2_['norm_por_tarifa'] == 'N'), _df2_[['5KPEAN', '6KPEAN', '7KPEAN', '8KPEAN', '9KPEAN']].sum(axis=1), np.where((_df2_['compilado_ts'] == 'C') & (_df2_['norm_por_tarifa'] == 'SN'), _df2_[['5KPCSN', '6KPCSN', '7KPCSN', '8KPCSN', '9KPCSN']].sum(axis=1), np.where((_df2_['compilado_ts'] == 'E') & (_df2_['norm_por_tarifa'] == 'SN'), _df2_[['5KPESN', '6KPESN', '7KPESN', '8KPESN', '9KPESN']].sum(axis=1), _df2_[['5KPEASN', '6KPEASN', '7KPEASN', '8KPEASN', '9KPEASN']].sum(axis=1))))))
                _df2_['kilometricas_por_TS'] = np.where((_df2_['TipoServicio'] == 'C') & (_df2_['sin_nominalizar'] == 0), _df2_[['5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']].replace(0, np.nan).min(axis=1).fillna(0), np.where((_df2_['TipoServicio'] == 'E') & (_df2_['sin_nominalizar'] == 0), _df2_[['5KPEN', '6KPEN', '7KPEN', '8KPEN', '9KPEN']].replace(0, np.nan).min(axis=1).fillna(0), np.where((_df2_['TipoServicio'] == 'EA') & (_df2_['sin_nominalizar'] == 0), _df2_[['5KPEAN', '6KPEAN', '7KPEAN', '8KPEAN', '9KPEAN']].replace(0, np.nan).min(axis=1).fillna(0), np.where((_df2_['TipoServicio'] == 'C') & (_df2_['sin_nominalizar'] == 1), _df2_[['5KPCSN', '6KPCSN', '7KPCSN', '8KPCSN', '9KPCSN']].replace(0, np.nan).min(axis=1).fillna(0), np.where((_df2_['TipoServicio'] == 'E') & (_df2_['sin_nominalizar'] == 1), _df2_[['5KPESN', '6KPESN', '7KPESN', '8KPESN', '9KPESN']].replace(0, np.nan).min(axis=1).fillna(0), np.where((_df2_['TipoServicio'] == 'EA') & (_df2_['sin_nominalizar'] == 1), _df2_[['5KPEASN', '6KPEASN', '7KPEASN', '8KPEASN', '9KPEASN']].replace(0, np.nan).min(axis=1).fillna(0), 0))))))

                _df2_['compilado_seccion'] = np.where(_df2_['compilado_tt'] == "S", _df2_['seccionadas_final'], np.where(_df2_['compilado_tt'] == "P", 1, np.where(_df2_['compilado_tt'] == "KM", _df2_['sec_1_4'], np.where(_df2_['compilado_tt'] == "KP", _df2_['kilometricas_por_TS'], 0))))
                _df2_['final_seccion'] = np.where((_df2_['GRUPO_TARIFARIO'] == "SGII") & (_df2_['compilado_seccion'].isin([1, 2, 3])), 4, _df2_['compilado_seccion'])

                _df2_['Año'] = anio
                _df2_['Resolucion'] = resolucion
                _df2_['CONCAT_MACHEO2'] = _df2_['Año'].astype(int).astype(str) + _df2_['Resolucion'].astype(str) + _df2_['final_seccion'].astype(int).astype(str) + _df2_['GRUPO_TARIFARIO'].astype(str) + _df2_['compilado_ts'].astype(str) + _df2_['norm_por_tarifa'].astype(str)
                _df2_['CONCAT_MACHEO3'] = _df2_['Año'].astype(int).astype(str) + _df2_['Resolucion'].astype(str) + _df2_['final_seccion'].astype(int).astype(str) + _df2_['GRUPO_TARIFARIO'].astype(str) + _df2_['ID_LINEA'].astype(str) + _df2_['compilado_ts'].astype(str) + _df2_['norm_por_tarifa'].astype(str)

                # 6. Cruces con Resoluciones y Valorización Teórica
                _df2_ = pd.merge(_df2_, ttr_reso[['CONCAT', 'TTR E.C.']], how='left', left_on='CONCAT_MACHEO2', right_on='CONCAT').fillna({'TTR E.C.': 0})
                _df2_ = _df2_.rename(columns={"TTR E.C.": "Tarifa TRSUBE"}).drop(columns=['CONCAT'])
                _df2_ = pd.merge(_df2_, ttr_sgii_uma2[['CONCAT', 'TTR E.C.']], how='left', left_on='CONCAT_MACHEO3', right_on='CONCAT').fillna({'TTR E.C.': 0})
                _df2_ = _df2_.rename(columns={"TTR E.C.": "Tarifa TRSUBE2"}).drop(columns=['CONCAT'])

                _df2_['Tarifa TRSUBE_FINAL'] = np.where(_df2_['Tarifa TRSUBE2'] == 0, _df2_['Tarifa TRSUBE'], _df2_['Tarifa TRSUBE2'])
                _df2_['Recaudacion_TRSUBE'] = _df2_['Tarifa TRSUBE_FINAL'] * _df2_['USOS']

                # Factor Energías
                condiciones = [_df2_['ENERGIA'] == 1, _df2_['ENERGIA'] == 2, _df2_['ENERGIA'] == 3]
                _df2_['Recaudacion_TRSUBE'] = _df2_['Recaudacion_TRSUBE'] * np.select(condiciones, [1.3, 1.5, 1.0], default=1)

                # 7. Asignación de Subsecciones
                _df2_['SubSeccion'] = None
                def asignar_subsecciones(df, tarifas_dict, filtro_sin_nominalizar):
                    for _, (lim_inf, lim_sup, tipo_servicio, seccion) in tarifas_dict.items():
                        sub_rangos = np.linspace(lim_inf, lim_sup, 4)
                        for i in range(3):
                            mask = ((df['TARIFA'] >= sub_rangos[i] - 0.5) & (df['TARIFA'] < sub_rangos[i+1] - 0.5) & (df['PASES'] == 0) & (df['sin_nominalizar'] == filtro_sin_nominalizar) & (df['TipoServicio2'] == tipo_servicio))
                            df.loc[mask, 'SubSeccion'] = f"{seccion}-{i+1}"

                asignar_subsecciones(_df2_, tarifas_7, filtro_sin_nominalizar=0)
                asignar_subsecciones(_df2_, tarifas_8, filtro_sin_nominalizar=1)
                _df2_['SubSeccion'] = _df2_['SubSeccion'].fillna(_df2_['final_seccion'].astype(str))

                # Exportación
                buf_salida = io.BytesIO()
                _df2_.to_excel(buf_salida, index=False, sheet_name='Valorizacion_TTR')
                buf_salida.seek(0)

                st.success("✅ ¡Valorización TTR calculada con éxito! Las matrices y diccionarios se cruzaron limpiamente.")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Usos Procesados", f"{_df2_['USOS'].sum():,.0f}")
                c2.metric("Total Líneas SILAS", f"{_df2_['LINEA_SILAS_DNGFF'].nunique()}")
                c3.metric("Recaudación Teórica TRSUBE", f"$ {_df2_['Recaudacion_TRSUBE'].sum():,.2f}")

                st.download_button(label="📥 Descargar Base Macheo TTR (.xlsx)", data=buf_salida, file_name="macheo_ttr_ARIA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            except Exception as e:
                st.error(f"Error procesando TTR: {e}")

# ==============================================================================
# CONTROLADOR LATERAL
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1792/1792404.png", width=60)
st.sidebar.title("Menú TTR_ARIA")
modulo_seleccionado = st.sidebar.radio("Navegación", ["Módulo 0: Tarifas JN", "Módulo 1: Liquidación DMK", "Módulo 3: Cálculo TTR"])
st.sidebar.markdown("---")
st.sidebar.info("Proyecto ARIA v2.1\n\nMotor unificado de cálculos TTR.")

if modulo_seleccionado == "Módulo 0: Tarifas JN": modulo_tarifas()
elif modulo_seleccionado == "Módulo 1: Liquidación DMK": modulo_dmk() # Módulo oculto por espacio en esta respuesta
else: modulo_calculo_ttr()
