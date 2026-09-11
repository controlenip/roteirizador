import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, fundir_super_pontos
from modules.export_analise import gerar_excel_analise, gerar_kml_analise

st.set_page_config(page_title="Análise Cruzada", page_icon="🔍", layout="wide")

st.markdown("""
<style>
    .stMultiSelect [data-baseweb="select"] > div:first-child { flex-wrap: wrap !important; }
    .stMultiSelect [data-baseweb="select"] > div:first-child > div:last-child { display: none !important; }
    .stMultiSelect [data-baseweb="tag"] { max-width: 100% !important; }
</style>
""", unsafe_allow_html=True)

try:
    from modules.export_analise import injetar_logo
    injetar_logo()
except: pass

def corrigir_coord(val, limite):
    if pd.isna(val): return np.nan
    v = float(val)
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v

st.markdown("<h1 class='brand-title'>🔍 Análise Cruzada e Auditoria</h1>", unsafe_allow_html=True)
st.info("💡 Cruza bases de Saneamento e Levantamento, isola notas bloqueadas (SAP FINL/CANC) e define os colaboradores mais próximos independente da origem.")

if "df_final_analise" not in st.session_state: st.session_state.df_final_analise = pd.DataFrame()
if "is_done_analise" not in st.session_state: st.session_state.is_done_analise = False

with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    raio_prox = st.slider("Distância p/ agrupar obras (Metros)", 10, 1000, 100, 10)
    
    st.markdown("---")
    
    if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
        df_fin = st.session_state.df_final_analise.copy()
        
        if 'COR_NOME' not in df_fin.columns:
            st.session_state.is_done_analise = False
            st.session_state.df_final_analise = pd.DataFrame()
            st.rerun()
        
        st.markdown("### 🎨 Filtro de Cores (Mapa e Export)")
        opcoes_cores = sorted(df_fin['COR_NOME'].unique().tolist())
        cores_selecionadas = st.multiselect("Selecione os dados para visualizar:", opcoes_cores, default=opcoes_cores)
        
        if not cores_selecionadas:
            st.warning("Selecione pelo menos uma cor para gerar o mapa e os relatórios.")
            st.stop()
            
        df_view = df_fin[df_fin['COR_NOME'].isin(cores_selecionadas)].copy()
        
        st.markdown("---")
        d_fmt = datetime.now().strftime("%d.%m.%Y_%H%M")
        
        # EXCEL COM 3 ABAS
        bu_xl = io.BytesIO()
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
            dict_dfs = {
                'Consolidado (Todas)': df_view,
                'Apenas Saneamento': df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'],
                'Apenas Levantamento': df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO']
            }
            excel_bytes = gerar_excel_analise(dict_dfs)
            zx.writestr(f"Planilha_Analise_Cruzada.xlsx", excel_bytes)
            
        st.download_button("🌐 Baixar Planilha Excel (ZIP)", data=bu_xl.getvalue(), file_name=f"Analise_Planilhas_{d_fmt}.zip", use_container_width=True)
        
        # KML
        kml_str = gerar_kml_analise(df_view)
        bu_kml = io.BytesIO()
        with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
            zk.writestr("Mapa_Analise_Cruzada.kml", kml_str.encode('utf-8'))
            
        st.download_button("🗺️ Baixar Mapa (KML)", data=bu_kml.getvalue(), file_name=f"Analise_Mapa_{d_fmt}.zip", use_container_width=True)
        
        if st.button("🧹 Nova Análise", type="primary", use_container_width=True):
            st.session_state.is_done_analise = False
            st.session_state.df_final_analise = pd.DataFrame()
            st.rerun()

if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
    
    st.markdown("### 🗺️ Mapa Analítico")
    mapa = folium.Map(location=[df_view['LATITUDE'].mean(), df_view['LONGITUDE'].mean()], zoom_start=8) if not df_view.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    
    m_clust = MarkerCluster(name="📍 Obras Filtradas").add_to(mapa)
    
    for cid, grp in df_view.groupby('CLUSTER_ID'):
        lat = grp['LATITUDE'].iloc[0]
        lon = grp['LONGITUDE'].iloc[0]
        c_names = grp['COR_NOME'].tolist()
        
        if any('Preto' in c for c in c_names) or any('Inválidas' in c for c in c_names): c_i = 'black'
        elif any('Vermelho' in c for c in c_names) or any('Duplicadas' in c for c in c_names): c_i = 'red'
        elif len(grp) > 1: c_i = 'orange'
        else: c_i = grp['COR_MAPA'].iloc[0]
        
        titulo_card = f"📍 Obras no Local ({len(grp)})"
        
        pop_html = f'''
        <div style="font-family:sans-serif; width:280px; max-height:280px; overflow-y:auto; border-radius:8px; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
            <div style="background:#0D256C; color:#ffffff; padding:8px; font-size:13px; font-weight:bold; text-align:center; position:sticky; top:0;">{titulo_card}</div>
            <div style="padding:10px; background:#fafafa; font-size:12px;">
        '''
        
        for _, r in grp.iterrows():
            n = html.escape(str(r.get('NOTA', '')))
            mun = html.escape(str(r.get('MUNICIPIO', '')))
            o = html.escape(str(r.get('ORIGEM_BASE', '')))
            s = html.escape(str(r.get('SITUACAO SAP', '')))
            col = html.escape(str(r.get('COLABORADORES MAIS PROXIMOS', '')))
            dup = html.escape(str(r.get('DUPLICADA', '')))
            
            aviso_gps = ""
            if dup == 'SIM' and len(grp) == 1:
                aviso_gps = f"<br><span style='color:red; font-size:10px;'>⚠️ A cópia desta nota está em outro ponto geográfico.</span>"
            
            pop_html += f'''
            <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
                <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                <tr><td style="padding:2px;"><b>Município:</b></td><td style="padding:2px;">{mun}</td></tr>
                <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                <tr><td style="padding:2px;"><b>SAP:</b></td><td style="padding:2px;">{s}</td></tr>
                <tr><td style="padding:2px;"><b>Equipes Perto:</b></td><td style="padding:2px;">{col}</td></tr>
                <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup}{aviso_gps}</td></tr>
            </table>
            <hr style="margin:4px 0; border:0; border-top:1px solid #ccc;">
            '''
        pop_html += '</div></div>'
        folium.Marker([lat, lon], icon=folium.Icon(color=c_i, icon='info-sign'), popup=folium.Popup(pop_html, max_width=320)).add_to(m_clust)
        
    folium.LayerControl().add_to(mapa)
    st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📈 Resumo da Volumetria")
    total_san = len(df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'])
    total_lev = len(df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO'])
    
    col_a, col_b, col_c = st.columns(3)
    col_a.info(f"**🟣 Saneamento (Validado):** {total_san} obras")
    col_b.success(f"**🟢 Levantamento (Validado):** {total_lev} obras")
    col_c.warning(f"**🎯 Total Geral:** {len(df_view)} obras")

    st.markdown(f"### 📊 Tabela Consolidada Detalhada")
    st.data_editor(df_view.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME', 'CLUSTER_ID'], errors='ignore'), use_container_width=True)

else:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### 🧹 Base Saneamento")
        file_san = st.file_uploader("Upload Saneamento", type=["xlsx", "xls", "csv"], key="san")
    with c2:
        st.markdown("#### 📋 Base Levantamento")
        file_lev = st.file_uploader("Upload Levantamento", type=["xlsx", "xls", "csv"], key="lev")
    with c3:
        st.markdown("#### 👥 Localidade Equipes")
        file_loc = st.file_uploader("Upload Localidades", type=["xlsx", "xls", "csv"], key="loc")

    if not (file_san and file_lev and file_loc):
        st.stop()

    if st.button("🚀 Processar Análise Cruzada", type="primary", use_container_width=True):
        st_run = time.time()
        pb = st.progress(0.0)
        tmp = st.empty()
        sgt = st.empty()
        
        def render_t(pct, msg):
            e = time.time() - st_run
            f = pct
            if f > 0:
                rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.05 else "Calculando..."
            else:
                rs = "Calculando..."
            es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
            tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center; box-shadow:0 2px 5px rgba(0,0,0,0.05);"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center; box-shadow:0 2px 5px rgba(0,0,0,0.05);"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)
            if msg: sgt.info(msg)
            pb.progress(pct)

        render_t(0.1, "Lendo planilhas de Saneamento e Levantamento...")
        
        # BLINDAGEM DO PANDAS (Usando io.BytesIO para evitar TypeError)
        file_san_buffer = io.BytesIO(file_san.getvalue())
        file_lev_buffer = io.BytesIO(file_lev.getvalue())
        file_loc_buffer = io.BytesIO(file_loc.getvalue())
        
        df_san = pd.read_excel(file_san_buffer) if not file_san.name.lower().endswith('.csv') else pd.read_csv(file_san_buffer)
        df_lev = pd.read_excel(file_lev_buffer) if not file_lev.name.lower().endswith('.csv') else pd.read_csv(file_lev_buffer)
        
        df_san.columns = normalize_cols(df_san.columns)
        df_lev.columns = normalize_cols(df_lev.columns)
        
        # UNIFICAÇÃO DE "MUNICÍPIO" E "MUNICIPIO"
        for c in df_san.columns:
            if 'MUNI' in c: df_san.rename(columns={c: 'MUNICIPIO'}, inplace=True); break
        for c in df_lev.columns:
            if 'MUNI' in c: df_lev.rename(columns={c: 'MUNICIPIO'}, inplace=True); break
        
        # IDENTIFICAÇÃO ABSOLUTA DA LATITUDE SANEAMENTO
        for c in df_san.columns:
            if 'PROJETO' in c.upper() and 'LAT' in c.upper(): df_san.rename(columns={c: 'LATITUDE'}, inplace=True); break
        for c in df_san.columns:
            if 'PROJETO' in c.upper() and 'LON' in c.upper(): df_san.rename(columns={c: 'LONGITUDE'}, inplace=True); break
            
        if 'LATITUDE' not in df_san.columns:
            for c in df_san.columns:
                if 'LAT' in c.upper(): df_san.rename(columns={c: 'LATITUDE'}, inplace=True); break
        if 'LONGITUDE' not in df_san.columns:
            for c in df_san.columns:
                if 'LON' in c.upper(): df_san.rename(columns={c: 'LONGITUDE'}, inplace=True); break
                
        # IDENTIFICAÇÃO LEVANTAMENTO
        for c in df_lev.columns:
            if 'LAT' in c.upper() and 'LATITUDE' not in df_lev.columns: df_lev.rename(columns={c: 'LATITUDE'}, inplace=True)
        for c in df_lev.columns:
            if 'LON' in c.upper() and 'LONGITUDE' not in df_lev.columns: df_lev.rename(columns={c: 'LONGITUDE'}, inplace=True)
        
        # IDENTIFICAÇÃO DE NOTA
        for pref in ['NOTA', 'PROTOCOLO', 'OS', 'ID_SISCO']:
            if pref in df_san.columns: df_san.rename(columns={pref: 'NOTA'}, inplace=True); break
        for pref in ['NOTA', 'PROTOCOLO', 'OS', 'ID_SISCO']:
            if pref in df_lev.columns: df_lev.rename(columns={pref: 'NOTA'}, inplace=True); break
        
        df_san['ORIGEM_BASE'] = 'SANEAMENTO'
        df_lev['ORIGEM_BASE'] = 'LEVANTAMENTO'
        
        if 'NOTA' in df_san.columns: df_san['NOTA'] = df_san['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
        if 'NOTA' in df_lev.columns: df_lev['NOTA'] = df_lev['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        render_t(0.3, "Lendo Localidades...")
        
        # SOLUÇÃO BLINDADA PARA CSV OU EXCEL MULTI-ABAS
        dfs_loc = []
        if file_loc.name.lower().endswith('.csv'):
            df_temp = pd.read_csv(file_loc_buffer)
            df_temp.columns = normalize_cols(df_temp.columns)
            for c in df_temp.columns:
                if 'NOME' in c: df_temp.rename(columns={c: 'NOME_COLAB'}, inplace=True)
                if 'LAT' in c: df_temp.rename(columns={c: 'LAT_LOC'}, inplace=True)
                if 'LON' in c: df_temp.rename(columns={c: 'LON_LOC'}, inplace=True)
            df_temp['TIPO_EQUIPE'] = 'Equipe (CSV)'
            dfs_loc.append(df_temp)
        else:
            xls_loc = pd.ExcelFile(file_loc_buffer)
            for sheet in xls_loc.sheet_names:
                df_temp = pd.read_excel(xls_loc, sheet_name=sheet)
                df_temp.columns = normalize_cols(df_temp.columns)
                for c in df_temp.columns:
                    if 'NOME' in c: df_temp.rename(columns={c: 'NOME_COLAB'}, inplace=True)
                    if 'LAT' in c: df_temp.rename(columns={c: 'LAT_LOC'}, inplace=True)
                    if 'LON' in c: df_temp.rename(columns={c: 'LON_LOC'}, inplace=True)
                
                tipo = "Saneamento" if "SAN" in sheet.upper() else "Levantamento"
                df_temp['TIPO_EQUIPE'] = tipo
                dfs_loc.append(df_temp)
                
        df_loc = pd.concat(dfs_loc, ignore_index=True)
            
        for c in df_loc.columns:
            if 'NOME' in c and 'NOME_COLAB' not in df_loc.columns: df_loc.rename(columns={c: 'NOME_COLAB'}, inplace=True)
            if 'LAT' in c and 'LAT_LOC' not in df_loc.columns: df_loc.rename(columns={c: 'LAT_LOC'}, inplace=True)
            if 'LON' in c and 'LON_LOC' not in df_loc.columns: df_loc.rename(columns={c: 'LON_LOC'}, inplace=True)
            
        df_loc['LAT_LOC'] = pd.to_numeric(df_loc['LAT_LOC'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_loc['LON_LOC'] = pd.to_numeric(df_loc['LON_LOC'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_loc = df_loc.dropna(subset=['LAT_LOC', 'LON_LOC'])
        
        if 'NOME_COLAB' not in df_loc.columns: 
            st.error("A planilha de Localidades não tem a coluna de Nomes.")
            st.stop()
            
        lat_locs = df_loc['LAT_LOC'].values
        lon_locs = df_loc['LON_LOC'].values
        nomes_locs = df_loc['NOME_COLAB'].astype(str).values
        tipos_locs = df_loc['TIPO_EQUIPE'].astype(str).values
        
        render_t(0.4, "Validando Status SAP (Bloqueios)...")
        status_col = 'STATUS_SAP' if 'STATUS_SAP' in df_lev.columns else ('STATUS SAP' if 'STATUS SAP' in df_lev.columns else ('STATUS' if 'STATUS' in df_lev.columns else None))
        status_dict = df_lev.set_index('NOTA')[status_col].to_dict() if status_col else {}

        def get_situacao(nota):
            s = str(status_dict.get(nota, '')).strip().upper()
            if s in ['FINL', 'CANC']: return f"BLOQUEADO ({s})"
            return "APTO"

        if 'NOTA' in df_san.columns: df_san['SITUACAO SAP'] = df_san['NOTA'].apply(get_situacao)
        if 'NOTA' in df_lev.columns: df_lev['SITUACAO SAP'] = df_lev['NOTA'].apply(get_situacao)
        
        render_t(0.5, "Verificando Duplicidades nas Bases...")
        notas_san = set(df_san['NOTA'].dropna()) if 'NOTA' in df_san.columns else set()
        notas_lev = set(df_lev['NOTA'].dropna()) if 'NOTA' in df_lev.columns else set()
        duplicadas = notas_san.intersection(notas_lev)
        
        if 'NOTA' in df_san.columns: df_san['DUPLICADA'] = df_san['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
        if 'NOTA' in df_lev.columns: df_lev['DUPLICADA'] = df_lev['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
        
        render_t(0.6, "Iniciando Cruzamento Espacial (Colaboradores)...")
        
        def get_closest_teams(lat, lon):
            if pd.isna(lat) or pd.isna(lon): return "DESCONHECIDO"
            dists = haversine_vectorized(lat, lon, lat_locs, lon_locs)
            if len(dists) == 0: return "DESCONHECIDO"
            
            sorted_idx = np.argsort(dists)
            res = []
            for i in sorted_idx[:2]:
                d_km = dists[i]
                if len(res) == 1 and d_km > 100: break
                nome = nomes_locs[i].title()
                tipo = tipos_locs[i]
                res.append(f"{nome} ({tipo}) - {d_km:.1f}km")
                
            return " | ".join(res)

        df_master = pd.concat([df_san, df_lev], ignore_index=True)
        df_master['LAT_NUM'] = pd.to_numeric(df_master['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_master['LON_NUM'] = pd.to_numeric(df_master['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_master['LAT_NUM'] = df_master['LAT_NUM'].apply(lambda x: corrigir_coord(x, 90))
        df_master['LON_NUM'] = df_master['LON_NUM'].apply(lambda x: corrigir_coord(x, 180))
        df_master['LATITUDE'] = df_master['LAT_NUM']
        df_master['LONGITUDE'] = df_master['LON_NUM']
        
        df_valid = df_master.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
        
        render_t(0.8, "Agrupando Obras Vizinhas e Processando Cores...")
        df_clustered, _ = fundir_super_pontos(df_valid, raio_metros=st.session_state.get('raio_prox', 50), agrupar_por_levantador=False)
        
        expanded = []
        c_id = 0
        for _, r in df_clustered.iterrows():
            is_prox = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
            if is_prox:
                for orig in r['_ORIGINAL_ROWS']:
                    nr = dict(orig)
                    nr['PROXIMA'] = 'SIM'
                    nr['CLUSTER_ID'] = c_id
                    expanded.append(nr)
            else:
                nr = r.to_dict()
                nr['PROXIMA'] = 'NÃO'
                nr['CLUSTER_ID'] = c_id
                expanded.append(nr)
            c_id += 1
                
        df_final = pd.DataFrame(expanded)
        df_final['COLABORADORES MAIS PROXIMOS'] = df_final.apply(lambda x: get_closest_teams(x['LATITUDE'], x['LONGITUDE']), axis=1)

        def determinar_cor(linha):
            dupl = str(linha.get('DUPLICADA', ''))
            prox = str(linha.get('PROXIMA', ''))
            orig = str(linha.get('ORIGEM_BASE', ''))
            
            is_black = False
            
            st_sap = str(linha.get('SITUACAO SAP', '')).upper()
            if 'BLOQUEADO' in st_sap or 'FINL' in st_sap or 'CANC' in st_sap:
                is_black = True
                
            if orig == 'LEVANTAMENTO':
                st_list = str(linha.get('STATUS_LIST', linha.get('STATUS LIST', '0'))).strip().upper().replace('.0', '')
                if st_list in ['NAN', 'NONE', '']: st_list = '0'
                
                st_sisco = str(linha.get('STATUS_SISCO', linha.get('STATUS SISCO', '0'))).strip().upper().replace('.0', '')
                if st_sisco in ['NAN', 'NONE', '']: st_sisco = '0'
                
                v_list = ['0', 'EM LEVANTAMENTO', 'CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO']
                v_sisco = ['0', 'PRÉ ANALISE', 'PRE ANALISE', 'LIBERADO PARA LEVANTAMENTO', 'LIBERADO P/ LEVANTAMENTO']
                
                if st_list not in v_list or st_sisco not in v_sisco:
                    is_black = True
            
            if is_black: return 'black', '⚫ Notas Inválidas'
            if dupl == 'SIM': return 'red', '🔴 Notas Duplicadas'
            if prox == 'SIM': return 'orange', '🟠 Notas Próximas'
            if orig == 'LEVANTAMENTO': return 'green', '🟢 Notas Levantamento Solitárias'
            if orig == 'SANEAMENTO': return 'purple', '🟣 Notas Saneamento Solitárias'
            return 'blue', '🔵 Outras'
            
        cores_calculadas = [determinar_cor(r) for _, r in df_final.iterrows()]
        df_final['COR_MAPA'] = [c[0] for c in cores_calculadas]
        df_final['COR_NOME'] = [c[1] for c in cores_calculadas]

        front_cols = ['NOTA', 'ORIGEM_BASE', 'SITUACAO SAP', 'DUPLICADA', 'PROXIMA', 'COLABORADORES MAIS PROXIMOS', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'CLUSTER_ID', 'COR_MAPA', 'COR_NOME']
        rest_cols = [c for c in df_final.columns if c not in front_cols and not c.startswith('_')]
        df_final = df_final[front_cols + rest_cols]
        
        render_t(1.0, "✅ Análise Concluída!")
        time.sleep(1)
        
        st.session_state.df_final_analise = df_final
        st.session_state.is_done_analise = True
        st.rerun()
