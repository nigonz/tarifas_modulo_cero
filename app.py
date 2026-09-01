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
# FUNCIONES COMPARTIDAS Y MÓDULO 0 (TARIFAS TTR JN)
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

def modulo_tarifas():
    st.title("🚜 TTR_ARIA - Módulo 0: Tarifario JN")
    st.markdown("Cálculo de cuadros tarifarios con indexación automática y arquitectura de espejos.")

    col_izq, col_der = st.columns([1, 3])
    
    with col_izq:
        st.header("1. Cargar Historial")
        archivo_historico = st.file_uploader("Subir Matriz Histórica (.xlsx)", type=['xlsx'], key="mod0_hist")
        st.header("2. Nuevo Período")
        mes_act = st.text_input("Nombre del Mes Nuevo", "Septiembre")

    with col_der:
        st.subheader(f"Ingreso de Bases Tarifarias Puras: {mes_act}")
        st.info("💡 Ingresá tus 10 bases manuales. La base 1-4KMCN se indexará automáticamente.")

        llaves = ['1SCN', '2SCN', '3SCN', '4SCN', '5SCN', '5KPCN', '6KPCN', '7KPCN', '8KPCN', '9KPCN']

        if 'tabla_bases' not in st.session_state:
            st.session_state.tabla_bases = pd.DataFrame({
                "CONCAT Base": llaves,
                "Límite Inferior": pd.Series([""] * 10, dtype=str),
                "Límite Superior": pd.Series([""] * 10, dtype=str)
            })

        tarifas_editadas = st.data_editor(st.session_state.tabla_bases, hide_index=True, use_container_width=True)

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🪄 Auto-Completar Límites Superiores", use_container_width=True):
                df_temp = tarifas_editadas.copy()
                for i in range(5):
                    if str(df_temp.loc[i, "Límite Inferior"]).strip() != "":
                        df_temp.loc[i, "Límite Superior"] = df_temp.loc[i, "Límite Inferior"]
                for i in range(5, 9):
                    if str(df_temp.loc[i+1, "Límite Inferior"]).strip() != "":
                        df_temp.loc[i, "Límite Superior"] = df_temp.loc[i+1, "Límite Inferior"]
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
                                    if pd.notna(val_1_4kmcn_viejo):
                                        base_1_4kmcn_dinamica = float(val_1_4kmcn_viejo) * factor_aumento

                            nuevos_limites_inf, nuevos_limites_sup = [], []
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
                                        base_inf = float(dict_bases_inf.get(base_key_inf, 0))
                                        base_sup = float(dict_bases_sup.get(base_key_sup, 0))
                                    
                                    if concat.startswith("5KP") and ts == "E": base_sup = base_inf
                                    
                                    mult_ts = 1.75 if ts == "EA" else (1.25 if ts == "E" else 1.0)
                                    mult_nom = 2.0 if "SN" in nom else 1.0

                                    val_inf = calcular_tarifa(base_inf, mult_ts, mult_nom)
                                    val_sup = calcular_tarifa(base_sup, mult_ts, mult_nom)
                                    if concat not in mapa_resultados: mapa_resultados[concat] = (val_inf, val_sup)

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
                                    if isinstance(cell.value, (int, float)): cell.number_format = '0.00'

                            final_buffer = io.BytesIO()
                            wb.save(final_buffer)
                            final_buffer.seek(0)

                            st.success(f"✅ ¡Matriz de {mes_act} generada con éxito!")
                            st.download_button(label=f"📥 Descargar Matriz Definitiva {mes_act} (.xlsx)",
                                               data=final_buffer.getvalue(), file_name=f"Matriz_TTR_ARIA_{mes_act}.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        except Exception as e: st.error(f"Error procesando: {e}")

# ==============================================================================
# MÓDULO 1: DMK (ITG y ATS) - EL PIPELINE DE COLAB
# ==============================================================================
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
            for celda in ws[letra][1:]:
                celda.number_format = '#,##0.00'
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

def resumir(df, claves, agg_estandar, **extra):
    return df.groupby(claves, dropna=False).agg(**agg_estandar, **extra).reset_index()

def modulo_dmk():
    st.title("🧮 TTR_ARIA - Módulo 1: Liquidación DMK (ITG/ATS)")
    st.markdown("Pipeline integral de enriquecimiento, control y generación de reportes de compensaciones.")

    with st.expander("⚙️ Parámetros de Negocio", expanded=False):
        col1, col2, col3 = st.columns(3)
        IVA = col1.number_format = col1.number_input("Factor IVA (Ej: 1.105)", value=1.105, format="%.3f")
        str_contratos_ats = col2.text_input("Contratos ATS (separados por coma)", "621")
        str_contratos_est = col3.text_input("Contratos Estudiantiles", "830, 831, 832, 833")
        
        CONTRATOS_ATS = [int(x.strip()) for x in str_contratos_ats.split(',')]
        CONTRATOS_ESTUDIANTILES = [int(x.strip()) for x in str_contratos_est.split(',')]
        ENERGIA_DEFECTO = 3
        GRUPOS_AMBA = ['DF', 'SGI', 'SGII', 'SGIKM', 'UMA1', 'UMA2', 'UPA', 'UPAKM']
        GRUPO_INP = 'INP'
        SIN_DATO = 'S/D'

    st.header("1. Carga de Archivos de Entrada")
    col1, col2 = st.columns(2)
    file_dggi = col1.file_uploader("1. Base DGGI (CSV o Excel)", type=['csv', 'xlsx'])
    file_nom_univ = col2.file_uploader("2. Nomenclador Líneas (Universo)", type=['xlsx'])
    file_nom_ramal = col1.file_uploader("3. Nomenclador Ramal - TS", type=['xlsx'])
    file_pme = col2.file_uploader("4. Parque Móvil - Energías", type=['xlsx'])

    if st.button("🚀 Ejecutar Liquidación DMK", type="primary", use_container_width=True):
        if not all([file_dggi, file_nom_univ, file_nom_ramal, file_pme]):
            st.error("⚠️ Faltan cargar archivos. Asegurate de subir los 4 requeridos.")
            return
            
        with st.spinner("Procesando pipeline de datos DMK..."):
            try:
                # 3. Carga de datos
                df_raw = pd.read_csv(file_dggi, encoding='ISO-8859-1', delimiter=';') if file_dggi.name.endswith('.csv') else pd.read_excel(file_dggi)
                nom_lineas_raw = pd.read_excel(file_nom_univ, sheet_name='Nomenclador_Interior')
                nom_ramal_raw = pd.read_excel(file_nom_ramal, sheet_name='NOMENCLADOR TS')
                pme_raw = pd.read_excel(file_pme, sheet_name='Nomenclador_PM_E')
                tipo_energia_raw = pd.read_excel(file_pme, sheet_name='Tipo_Energia')

                # 4. Normalización
                df_base = df_raw.copy()
                df_base.columns = df_base.columns.str.strip()
                df_base = df_base.rename(columns={'MONTO': 'RECAUDACION'})
                for c in ['DOMINIO', 'MK', 'VIAJE INTEGRADO', 'MEDIOS_DE_PAGO']:
                    if c in df_base.columns: df_base[c] = df_base[c].astype('string').str.strip().str.upper()
                for c in ['TARIFA', 'DEBITADO', 'DESCUENTO X INTEGRACION', 'CANTIDAD_USOS', 'RECAUDACION', 'TOTAL DESC POR INTEGRACION', 'DESCUENTO_TOTAL', 'DESCUENTO_ATRIBUTOS']:
                    if c in df_base.columns: df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0)
                for c in ['ID_EMPRESA', 'ID_LINEA', 'RAMAL', 'CONTRATO', 'INTERNO']:
                    if c in df_base.columns: df_base[c] = pd.to_numeric(df_base[c], errors='coerce').astype('Int64')

                # 5. Nom Lineas
                n_lin = nom_lineas_raw[['ID_LINEA', 'GT', 'SILAS - AMBA', 'ID_EMPRESA', 'Razon social', 'Jurisdiccion', 'Provincia', 'Localidad', 'Departamento']].copy()
                n_lin = n_lin.rename(columns={'GT': 'GRUPO_TARIFARIO', 'SILAS - AMBA': 'LINEA_SILAS_DNGFF', 'ID_EMPRESA': 'ID_EMPRESA_NOM', 'Razon social': 'RAZON_SOCIAL', 'Jurisdiccion': 'JURISDICCION', 'Provincia': 'PROVINCIA', 'Localidad': 'MUNICIPIO', 'Departamento': 'DEPARTAMENTO'})
                n_lin['ID_LINEA'] = pd.to_numeric(n_lin['ID_LINEA'], errors='coerce').astype('Int64')
                n_lin['ID_EMPRESA_NOM'] = pd.to_numeric(n_lin['ID_EMPRESA_NOM'], errors='coerce').astype('Int64')
                for c in ['GRUPO_TARIFARIO', 'LINEA_SILAS_DNGFF', 'RAZON_SOCIAL', 'JURISDICCION', 'PROVINCIA', 'MUNICIPIO', 'DEPARTAMENTO']:
                    n_lin[c] = n_lin[c].astype('string').str.strip()
                nom_lineas = n_lin.dropna(subset=['ID_LINEA'])

                # 6. Nom Ramal y PME
                n_ram = nom_ramal_raw[['IdRamalNS', 'TIPO DE SERVICIO FINAL']].copy()
                n_ram = n_ram.rename(columns={'IdRamalNS': 'RAMAL', 'TIPO DE SERVICIO FINAL': 'TIPO_SERVICIO'})
                n_ram['RAMAL'] = pd.to_numeric(n_ram['RAMAL'], errors='coerce').astype('Int64')
                n_ram['TIPO_SERVICIO'] = n_ram['TIPO_SERVICIO'].astype('string').str.strip().str.upper()
                nom_ramal = n_ram.dropna(subset=['RAMAL']).drop_duplicates(subset='RAMAL')

                p = pme_raw[['DOMINIO', 'ENERGIA']].copy()
                p['DOMINIO'] = p['DOMINIO'].astype('string').str.strip().str.upper()
                p['ENERGIA'] = pd.to_numeric(p['ENERGIA'], errors='coerce').astype('Int64')
                pme = p.dropna(subset=['DOMINIO']).drop_duplicates(subset='DOMINIO')
                MAPA_ENERGIA = dict(zip(pd.to_numeric(tipo_energia_raw['ENERGIA'], errors='coerce'), tipo_energia_raw['CONCEPTO'].astype(str).str.strip().str.upper()))

                # 7. Enriquecer
                d = df_base.merge(nom_lineas, on='ID_LINEA', how='left', validate='m:1')
                d = d.merge(nom_ramal, on='RAMAL', how='left', validate='m:1')
                d = d.merge(pme, on='DOMINIO', how='left', validate='m:1')
                
                d['ES_BENEFICIARIA'] = np.where(d['GRUPO_TARIFARIO'].notna(), 'SI', 'NO')
                gt_upper = d['GRUPO_TARIFARIO'].astype('string').str.strip().str.upper()
                es_amba = gt_upper.isin(GRUPOS_AMBA).fillna(False).to_numpy()
                es_inp = gt_upper.eq(GRUPO_INP).fillna(False).to_numpy()
                d['AMBA'] = np.select([es_amba, es_inp], ['SI', 'AMBA - INP'], default='NO')
                
                d['EN_PARQUE_MOVIL'] = np.where(d['ENERGIA'].notna(), 'SI', 'NO')
                d['TIPO_ENERGIA'] = d['ENERGIA'].fillna(ENERGIA_DEFECTO).astype('Int64')
                d['ENERGIA_DESC'] = d['TIPO_ENERGIA'].map(MAPA_ENERGIA).astype('string')
                d = d.drop(columns='ENERGIA')
                for c in ['JURISDICCION', 'PROVINCIA', 'MUNICIPIO', 'DEPARTAMENTO', 'GRUPO_TARIFARIO', 'RAZON_SOCIAL', 'LINEA_SILAS_DNGFF', 'TIPO_SERVICIO']:
                    d[c] = d[c].astype('string').fillna(SIN_DATO)
                
                sin_empresa_nom = d['ID_EMPRESA_NOM'].isna().to_numpy()
                misma_empresa = (d['ID_EMPRESA'] == d['ID_EMPRESA_NOM']).fillna(False).to_numpy()
                d['COINCIDE_EMPRESA'] = np.select([sin_empresa_nom, misma_empresa], [SIN_DATO, 'SI'], default='NO')

                # 8. Marcas beneficio
                d['ES_INTEGRADO'] = np.where(d['VIAJE INTEGRADO'].astype('string').eq('SI'), 'SI', 'NO')
                es_ats = d['CONTRATO'].isin(CONTRATOS_ATS).to_numpy()
                es_est = d['CONTRATO'].isin(CONTRATOS_ESTUDIANTILES).to_numpy()
                tiene_desc = (d['DESCUENTO_ATRIBUTOS'] > 0).to_numpy()
                d['ES_ATS'] = np.where(es_ats, 'SI', 'NO')
                d['ES_ESTUDIANTIL'] = np.where(es_est, 'SI', 'NO')
                d['ES_OTRO_BENEFICIO'] = np.where(tiene_desc & ~es_ats & ~es_est, 'SI', 'NO')
                d['TIPO_BENEFICIO'] = np.select([es_ats, es_est, tiene_desc & ~es_ats & ~es_est], ['ATS', 'ESTUDIANTIL', 'OTRO BENEFICIO'], default='SIN BENEFICIO')

                # 9. Compensaciones
                u = d['CANTIDAD_USOS']
                d['COMP. ITG'] = d['TOTAL DESC POR INTEGRACION']
                d['COMP. ITG s/IVA'] = d['COMP. ITG'] / IVA
                d['COMP. ATS'] = np.where(es_ats, d['DESCUENTO_ATRIBUTOS'], 0.0)
                d['COMP. ATS s/IVA'] = d['COMP. ATS'] / IVA
                d['DESCUENTO_TOTAL s/IVA'] = d['DESCUENTO_TOTAL'] / IVA
                d['COMP. TOTAL s/IVA'] = d['COMP. ITG s/IVA'] + d['COMP. ATS s/IVA']
                
                d['RECAUDACION_CALC'] = d['DEBITADO'] * u
                d['DESC_TOTAL_CALC'] = (d['TARIFA'] - d['DEBITADO']) * u
                d['COMP_ITG_CALC'] = d['DESCUENTO X INTEGRACION'] * u
                d['COMP_ATS_CALC'] = np.where(es_ats, (d['TARIFA'] - d['DEBITADO'] - d['DESCUENTO X INTEGRACION']) * u, 0.0)
                d['DIF_RECAUDACION'] = (d['RECAUDACION_CALC'] - d['RECAUDACION']).round(2)
                d['DIF_DESC_TOTAL'] = (d['DESC_TOTAL_CALC'] - d['DESCUENTO_TOTAL']).round(2)
                d['DIF_ITG'] = (d['COMP_ITG_CALC'] - d['COMP. ITG']).round(2)
                d['DIF_ATS'] = (d['COMP_ATS_CALC'] - d['COMP. ATS']).round(2)

                # 10. Separación
                df_final = d[d['ES_BENEFICIARIA'] == 'SI'].copy()
                df_no_benef = d[d['ES_BENEFICIARIA'] == 'NO'].copy()

                ORDEN_BASE = [
                    'ID_EMPRESA', 'ID_EMPRESA_NOM', 'COINCIDE_EMPRESA', 'RAZON_SOCIAL', 'ID_LINEA', 'LINEA_SILAS_DNGFF', 'RAMAL', 'TIPO_SERVICIO', 'INTERNO', 'DOMINIO', 'MK',
                    'JURISDICCION', 'PROVINCIA', 'MUNICIPIO', 'DEPARTAMENTO', 'GRUPO_TARIFARIO', 'AMBA', 'ES_BENEFICIARIA',
                    'EN_PARQUE_MOVIL', 'TIPO_ENERGIA', 'ENERGIA_DESC',
                    'CONTRATO', 'MEDIOS_DE_PAGO', 'VIAJE INTEGRADO', 'ES_INTEGRADO', 'TIPO_BENEFICIO', 'ES_ATS', 'ES_ESTUDIANTIL', 'ES_OTRO_BENEFICIO',
                    'TARIFA', 'DEBITADO', 'DESCUENTO X INTEGRACION', 'CANTIDAD_USOS',
                    'RECAUDACION', 'DESCUENTO_TOTAL', 'DESCUENTO_TOTAL s/IVA', 'TOTAL DESC POR INTEGRACION', 'DESCUENTO_ATRIBUTOS',
                    'COMP. ITG', 'COMP. ITG s/IVA', 'COMP. ATS', 'COMP. ATS s/IVA', 'COMP. TOTAL s/IVA',
                    'RECAUDACION_CALC', 'DESC_TOTAL_CALC', 'COMP_ITG_CALC', 'COMP_ATS_CALC', 'DIF_RECAUDACION', 'DIF_DESC_TOTAL', 'DIF_ITG', 'DIF_ATS',
                ]
                
                df_final = df_final[[c for c in ORDEN_BASE if c in df_final.columns] + [c for c in df_final.columns if c not in ORDEN_BASE]]
                df_no_benef = df_no_benef[[c for c in ORDEN_BASE if c in df_no_benef.columns] + [c for c in df_no_benef.columns if c not in ORDEN_BASE]]

                # 12. Tablas
                AGG_ESTANDAR = dict(RECAUDACION=('RECAUDACION', 'sum'), USOS=('CANTIDAD_USOS', 'sum'), DESCUENTO_TOTAL=('DESCUENTO_TOTAL', 'sum'), DESCUENTO_TOTAL_sIVA=('DESCUENTO_TOTAL s/IVA', 'sum'), COMP_ITG=('COMP. ITG', 'sum'), COMP_ITG_sIVA=('COMP. ITG s/IVA', 'sum'), COMP_ATS=('COMP. ATS', 'sum'), COMP_ATS_sIVA=('COMP. ATS s/IVA', 'sum'))
                
                resumen_compensacion = resumir(df_final, ['JURISDICCION', 'PROVINCIA', 'MUNICIPIO', 'GRUPO_TARIFARIO', 'AMBA'], AGG_ESTANDAR)
                resumen_unicos = df_final.groupby(['PROVINCIA', 'GRUPO_TARIFARIO', 'AMBA'], dropna=False).agg(LINEAS_SILAS_UNICAS=('LINEA_SILAS_DNGFF', 'nunique'), ID_LINEAS_UNICAS=('ID_LINEA', 'nunique'), RAMALES_UNICOS=('RAMAL', 'nunique'), EMPRESAS_UNICAS=('ID_EMPRESA', 'nunique'), INTERNOS_UNICOS=('INTERNO', 'nunique'), DOMINIOS_UNICOS=('DOMINIO', 'nunique')).reset_index()
                resumen_energia = resumir(df_final[df_final['EN_PARQUE_MOVIL'] == 'SI'], ['ID_LINEA', 'LINEA_SILAS_DNGFF', 'PROVINCIA', 'AMBA', 'ENERGIA_DESC'], AGG_ESTANDAR)
                resumen_contrato = df_final.groupby(['CONTRATO', 'AMBA'], dropna=False).agg(RECAUDACION=('RECAUDACION', 'sum'), USOS=('CANTIDAD_USOS', 'sum'), DESCUENTO_TOTAL_sIVA=('DESCUENTO_TOTAL s/IVA', 'sum'), COMP_ITG_sIVA=('COMP. ITG s/IVA', 'sum'), COMP_ATS_sIVA=('COMP. ATS s/IVA', 'sum'), ATRIBUTO_EN_BASE=('DESCUENTO_ATRIBUTOS', 'sum')).reset_index()
                resumen_medio_pago = resumir(df_final, ['MEDIOS_DE_PAGO', 'PROVINCIA', 'MUNICIPIO', 'GRUPO_TARIFARIO', 'AMBA'], AGG_ESTANDAR)
                control_cobertura = d.groupby(['ES_BENEFICIARIA', 'JURISDICCION', 'AMBA'], dropna=False).agg(**AGG_ESTANDAR, LINEAS=('ID_LINEA', 'nunique')).reset_index()

                CLAVES_TARIFARIO = ['PROVINCIA', 'MUNICIPIO', 'GRUPO_TARIFARIO', 'AMBA', 'LINEA_SILAS_DNGFF', 'ID_LINEA', 'RAMAL', 'CONTRATO', 'TARIFA', 'DEBITADO']
                resumen_tarifario = df_final.groupby(CLAVES_TARIFARIO, dropna=False).agg(USOS=('CANTIDAD_USOS', 'sum'), RECAUDACION=('RECAUDACION', 'sum'), **{'TOTAL DESC POR INTEGRACION': ('TOTAL DESC POR INTEGRACION', 'sum')}, DESCUENTO_ATRIBUTOS=('DESCUENTO_ATRIBUTOS', 'sum'), DESCUENTO_TOTAL=('DESCUENTO_TOTAL', 'sum')).reset_index()

                def _listar(s): return ', '.join(str(v) for v in sorted(s.dropna().unique()))
                no_benef_detalle = df_no_benef.groupby(['ID_LINEA', 'ID_EMPRESA'], dropna=False).agg(RAMALES=('RAMAL', _listar), CONTRATOS=('CONTRATO', _listar), RECAUDACION=('RECAUDACION', 'sum'), USOS=('CANTIDAD_USOS', 'sum'), DESCUENTO_TOTAL_sIVA=('DESCUENTO_TOTAL s/IVA', 'sum'), COMP_ITG_sIVA=('COMP. ITG s/IVA', 'sum'), COMP_ATS_sIVA=('COMP. ATS s/IVA', 'sum'), ATRIBUTO_EN_BASE=('DESCUENTO_ATRIBUTOS', 'sum')).reset_index().sort_values('USOS', ascending=False)

                df_tarifario_dominio = df_final.copy()
                es_gasoil = (df_tarifario_dominio['TIPO_ENERGIA'] == 3).fillna(False).to_numpy()
                df_tarifario_dominio['DOMINIO'] = np.where(es_gasoil, 'NO', df_tarifario_dominio['DOMINIO'])
                df_tarifario_dominio['ENERGIA'] = df_tarifario_dominio['TIPO_ENERGIA']
                resumen_tarifario_dominio = df_tarifario_dominio.groupby(CLAVES_TARIFARIO + ['DOMINIO', 'ENERGIA'], dropna=False).agg(USOS=('CANTIDAD_USOS', 'sum'), RECAUDACION=('RECAUDACION', 'sum'), **{'TOTAL DESC POR INTEGRACION': ('TOTAL DESC POR INTEGRACION', 'sum')}, DESCUENTO_ATRIBUTOS=('DESCUENTO_ATRIBUTOS', 'sum'), DESCUENTO_TOTAL=('DESCUENTO_TOTAL', 'sum')).reset_index()

                HOJAS = {
                    'Resumen_Compensacion': resumen_compensacion, 'Resumen_Unicos': resumen_unicos, 'Resumen_Energia': resumen_energia,
                    'Resumen_Contrato': resumen_contrato, 'Resumen_MedioPago': resumen_medio_pago, 'Control_Cobertura': control_cobertura,
                    'Resumen_Tarifario': resumen_tarifario, 'Resumen_Tarifario_Dominio': resumen_tarifario_dominio, 'NoBenef_Detalle': no_benef_detalle,
                }

                COLS_MONEDA = {
                    'TARIFA', 'DEBITADO', 'DESCUENTO X INTEGRACION', 'RECAUDACION', 'RECAUDACION_CALC', 'DESCUENTO_TOTAL', 'DESCUENTO_TOTAL s/IVA', 'DESCUENTO_TOTAL_sIVA', 'DESC_TOTAL_CALC',
                    'TOTAL DESC POR INTEGRACION', 'DESCUENTO_ATRIBUTOS', 'ATRIBUTO_EN_BASE', 'COMP. ITG', 'COMP. ITG s/IVA', 'COMP. ATS', 'COMP. ATS s/IVA', 'COMP. TOTAL s/IVA',
                    'COMP_ITG', 'COMP_ITG_sIVA', 'COMP_ATS', 'COMP_ATS_sIVA', 'COMP_ITG_CALC', 'COMP_ATS_CALC', 'DIF_RECAUDACION', 'DIF_DESC_TOTAL', 'DIF_ITG', 'DIF_ATS', 'MONTO TOTAL COBRADO', 'DESCUENTO TOTAL ITG', 'DESCUENTO TOTAL ATS'
                }

                # 14. Excel ATS 621
                base_621 = df_final[df_final['CONTRATO'].isin(CONTRATOS_ATS)]
                resumen_621 = base_621.groupby(['JURISDICCION', 'PROVINCIA', 'MUNICIPIO', 'ID_EMPRESA', 'RAZON_SOCIAL', 'ID_LINEA', 'RAMAL'], dropna=False).agg(**{'MONTO TOTAL COBRADO': ('RECAUDACION', 'sum')}, **{'CANTIDAD DE TRANSACCIONES': ('CANTIDAD_USOS', 'sum')}, **{'DESCUENTO TOTAL ITG': ('COMP. ITG', 'sum')}, **{'DESCUENTO TOTAL ATS': ('COMP. ATS', 'sum')}, AMBA=('AMBA', 'first')).reset_index().rename(columns={'ID_LINEA': 'LINEA', 'RAZON_SOCIAL': 'RAZON SOCIAL'})

                # EXPORTACION EN MEMORIA
                buf_resumenes = io.BytesIO()
                with pd.ExcelWriter(buf_resumenes, engine='openpyxl') as writer:
                    for nombre, tabla in HOJAS.items():
                        tabla.to_excel(writer, index=False, sheet_name=nombre)
                        formatear_excel(writer.sheets[nombre], tabla, COLS_MONEDA)
                buf_resumenes.seek(0)

                buf_621 = io.BytesIO()
                with pd.ExcelWriter(buf_621, engine='openpyxl') as writer:
                    resumen_621.to_excel(writer, index=False, sheet_name='ATS_621')
                    formatear_excel(writer.sheets['ATS_621'], resumen_621, COLS_MONEDA)
                buf_621.seek(0)

                buf_csv = io.BytesIO()
                df_final.to_csv(buf_csv, index=False, sep=';', encoding='utf-8-sig')
                buf_csv.seek(0)

                # DASHBOARD FINAL
                st.success("✅ ¡Liquidación procesada con éxito!")
                st.markdown("### Resumen de la Corrida")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Líneas Beneficiarias", f"{df_final['ID_LINEA'].nunique():,}")
                c2.metric("Usos Totales", f"{df_final['CANTIDAD_USOS'].sum():,}")
                c3.metric("Recaudación", f"$ {df_final['RECAUDACION'].sum():,.2f}")
                c4.metric("Compensación Total s/IVA", f"$ {df_final['COMP. TOTAL s/IVA'].sum():,.2f}")
                
                st.markdown("### Descargas Disponibles")
                d1, d2, d3 = st.columns(3)
                d1.download_button("📥 Descargar Resúmenes (.xlsx)", data=buf_resumenes, file_name="DGGI_ITG_ATS_Resumenes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                d2.download_button("📥 Descargar Reporte ATS 621 (.xlsx)", data=buf_621, file_name="DGGI_ATS_621.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                d3.download_button("📥 Descargar Base Detalle (.csv)", data=buf_csv, file_name="DGGI_Base_Detalle.csv", mime="text/csv", use_container_width=True)

            except Exception as e:
                st.error(f"Error procesando DMK: {e}")

# ==============================================================================
# CONTROLADOR LATERAL
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1792/1792404.png", width=60)
st.sidebar.title("Menú TTR_ARIA")
modulo_seleccionado = st.sidebar.radio("Navegación", ["Módulo 0: Tarifas JN", "Módulo 1: Liquidación DMK"])

st.sidebar.markdown("---")
st.sidebar.info("Proyecto ARIA v2.0\n\nMotor unificado de cálculos TTR.")

if modulo_seleccionado == "Módulo 0: Tarifas JN":
    modulo_tarifas()
else:
    modulo_dmk()
