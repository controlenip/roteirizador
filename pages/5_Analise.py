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

# CSS para o multiselect e visualização
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
st.info("💡 Este módulo cruza as bases de **Saneamento** e **Levantamento**, bloqueia notas com SAP 'FINL' ou 'CANC' e identifica obras de diferentes bases no mesmo endereço geográfico.")

if "df_final_analise" not in st.session_state: st.session_state.df_final_analise = pd.DataFrame()
if "is_done_analise" not in st.session_state: st.session_state.is_done_analise = False

# ==========================================
# BARRA LATERAL E FILTROS DINÂMICOS
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    raio_prox = st.slider("Distância p/ agrupar obras (Metros)", 10, 1000, 100, 10)
    
    st.markdown("---")
    
    if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
        df_fin = st.session_state.df_final_analise.copy()
        
        # Trava de Segurança do Cache
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
        
        # GERAR EXCEL COM 3 ABAS
        bu_xl = io.BytesIO()
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
            dict_dfs = {
                'Consolidado (Todas)': df_view,
                'Apenas Saneamento': df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'],
                'Apenas Levantamento': df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO']
            }
            excel_bytes = gerar_excel_analise(dict_dfs)
            zx.writestr("Planilha_Analise_Cruzada.xlsx", excel_bytes)
            
        st.download_button("🌐 Baixar Planilha Excel (ZIP)", data=bu_xl.getvalue(), file_name=f"Analise_Planilha_{d_fmt}.zip", use_container_width=True)
        
        # GERAR KML
        kml_str = gerar_kml_analise(df_view)
        bu_kml = io.BytesIO()
        with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
            zk.writestr("Mapa_Analise_Cruzada.kml", kml_str.encode('utf-8'))
            
        st.download_button("🗺️ Baixar Mapa (KML)", data=bu_kml.getvalue(), file_name=f"Analise_Mapa_{d_fmt}.zip", use_container_width=True)
        
        if st.button("🧹 Nova Análise", type="primary", use_container_width=True):
            st.session_state.is_done_analise = False
            st.session_state.df_final_analise = pd.DataFrame()
            st.rerun()

# ==========================================
# EXIBIÇÃO DE RESULTADOS (MAPA AGRUPADO)
# ==========================================
if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
    
    st.markdown("### 🗺️ Mapa Analítico")
    mapa = folium.Map(location=[df_view['LATITUDE'].mean(), df_view['LONGITUDE'].mean()], zoom_start=8) if not df_view.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    
    m_clust = MarkerCluster(name="📍 Obras Filtradas").add_to(mapa)
    
    # Agrupa por CLUSTER_ID para renderizar 1 único pino por terreno com os dados das N notas nele
    for cid, grp in df_view.groupby('CLUSTER_ID'):
        lat = grp['LATITUDE'].iloc[0]
        lon = grp['LONGITUDE'].iloc[0]
        
        c_names = grp['COR_NOME'].tolist()
        
        # Decide a cor do pino agrupado
        if any('Preto' in c for c in c_names): c_i = 'black'
        elif any('Vermelho' in c for c in c_names): c_i = 'red'
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
            o = html.escape(str(r.get('ORIGEM_BASE', '')))
            s = html.escape(str(r.get('SITUACAO SAP', '')))
            col = html.escape(str(r.get('COLABORADOR MAIS PROXIMO', '')))
            dup = html.escape(str(r.get('DUPLICADA', '')))
            
            pop_html += f'''
            <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
                <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                <tr><td style="padding:2px;"><b>SAP:</b></td><td style="padding:2px;">{s}</td></tr>
                <tr><td style="padding:2px;"><b>Colab:</b></td><td style="padding:2px;">{col}</td></tr>
                <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup}</td></tr>
            </table>
            <hr style="margin:4px 0; border:0; border-top:1px solid #ccc;">
            '''
        
        pop_html += '</div></div>'
        
        folium.Marker([lat, lon], icon=folium.Icon(color=c_i, icon='info-sign'), popup=folium.Popup(pop_html, max_width=320)).add_to(m_clust)
        
    folium.LayerControl().add_to(mapa)
    st_folium(mapa, use_container_width=True, height=550)

    # --- PAINEL DE CONTABILIDADE (Obras por Base) ---
    st.markdown("### 📈 Resumo da Volumetria")
    total_san = len(df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'])
    total_lev = len(df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO'])
    total_geral = len(df_view)
    
    col_a, col_b, col_c = st.columns(3)
    col_a.info(f"**🟣 Total Saneamento:** {total_san} obras")
    col_b.success(f"**🟢 Total Levantamento:** {total_lev} obras")
    col_c.warning(f"**🎯 Total Geral:** {total_geral} obras")

    st.markdown(f"### 📊 Tabela Consolidada Detalhada")
    st.data_editor(df_view.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME', 'CLUSTER_ID'], errors='ignore'), use_container_width=True)

# ==========================================
# PAINEL DE UPLOAD E CRUZAMENTO
# ==========================================
else:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### 🧹 Base Saneamento")
        file_san = st.file_uploader("Upload Saneamento", type=["xlsx", "xls", "csv"], key="san")
    with c2:
        st.markdown("#### 📋 Base Levantamento")
        file_lev = st.file_uploader("Upload Levantamento", type=["xlsx", "xls", "csv"], key="lev")
    with c3:
        st.markdown("#### 👥 Localidade Levantadores")
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

        render_t(0.1, "Lendo planilhas e limpando colunas...")
        df_san = ler_planilha_cached(file_san.getvalue()) if not file_san.name.endswith('.csv') else pd.read_csv(file_san)
        df_lev = ler_planilha_cached(file_lev.getvalue()) if not file_lev.name.endswith('.csv') else pd.read_csv(file_lev)
        df_loc = ler_planilha_cached(file_loc.getvalue()) if not file_loc.name.endswith('.csv') else pd.read_csv(file_loc)
        
        df_san.columns = normalize_cols(df_san.columns)
        df_lev.columns = normalize_cols(df_lev.columns)
        df_loc.columns = normalize_cols(df_loc.columns)
        
        # --- RADAR DE COORDENADAS BLINDADO ---
        # Saneamento (Prioridade para LATITUDE PROJETO)
        for c in df_san.columns:
            if 'PROJETO' in c.upper() and 'LAT' in c.upper(): df_san.rename(columns={c: 'LATITUDE'}, inplace=True); break
        for c in df_san.columns:
            if 'PROJETO' in c.upper() and 'LON' in c.upper(): df_san.rename(columns={c: 'LONGITUDE'}, inplace=True); break
            
        # Se falhar o Projeto, procura genérico
        if 'LATITUDE' not in df_san.columns:
            for c in df_san.columns:
                if 'LAT' in c.upper(): df_san.rename(columns={c: 'LATITUDE'}, inplace=True); break
        if 'LONGITUDE' not in df_san.columns:
            for c in df_san.columns:
                if 'LON' in c.upper(): df_san.rename(columns={c: 'LONGITUDE'}, inplace=True); break
                
        # Levantamento
        for c in df_lev.columns:
            if 'LAT' in c.upper() and 'LATITUDE' not in df_lev.columns: df_lev.rename(columns={c: 'LATITUDE'}, inplace=True)
        for c in df_lev.columns:
            if 'LON' in c.upper() and 'LONGITUDE' not in df_lev.columns: df_lev.rename(columns={c: 'LONGITUDE'}, inplace=True)
        
        # --- IDENTIFICAÇÃO DE NOTA ---
        for pref in ['NOTA', 'PROTOCOLO', 'OS', 'ID SISCO']:
            if pref in df_san.columns:
                df_san.rename(columns={pref: 'NOTA'}, inplace=True); break
                
        for pref in ['NOTA', 'PROTOCOLO', 'OS', 'ID SISCO']:
            if pref in df_lev.columns:
                df_lev.rename(columns={pref: 'NOTA'}, inplace=True); break
        
        df_san['ORIGEM_BASE'] = 'SANEAMENTO'
        df_lev['ORIGEM_BASE'] = 'LEVANTAMENTO'
        
        if 'NOTA' in df_san.columns: df_san['NOTA'] = df_san['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
        if 'NOTA' in df_lev.columns: df_lev['NOTA'] = df_lev['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        render_t(0.3, "Validando Status SAP (Bloqueios)...")
        status_col = 'STATUS_SAP' if 'STATUS_SAP' in df_lev.columns else ('STATUS SAP' if 'STATUS SAP' in df_lev.columns else ('STATUS' if 'STATUS' in df_lev.columns else None))
        status_dict = df_lev.set_index('NOTA')[status_col].to_dict() if status_col else {}

        def get_situacao(nota):
            s = str(status_dict.get(nota, '')).strip().upper()
            if s in ['FINL', 'CANC']: return f"BLOQUEADO ({s})"
            return "APTO"

        if 'NOTA' in df_san.columns: df_san['SITUACAO SAP'] = df_san['NOTA'].apply(get_situacao)
        if 'NOTA' in df_lev.columns: df_lev['SITUACAO SAP'] = df_lev['NOTA'].apply(get_situacao)
        
        render_t(0.5, "Verificando Duplicidades...")
        notas_san = set(df_san['NOTA'].dropna()) if 'NOTA' in df_san.columns else set()
        notas_lev = set(df_lev['NOTA'].dropna()) if 'NOTA' in df_lev.columns else set()
        duplicadas = notas_san.intersection(notas_lev)
        
        if 'NOTA' in df_san.columns: df_san['DUPLICADA'] = df_san['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
        if 'NOTA' in df_lev.columns: df_lev['DUPLICADA'] = df_lev['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
        
        df_loc['LATITUDE'] = pd.to_numeric(df_loc['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_loc['LONGITUDE'] = pd.to_numeric(df_loc['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_loc = df_loc.dropna(subset=['LATITUDE', 'LONGITUDE'])
        
        nome_col_loc = next((c for c in df_loc.columns if c in ['NOME', 'LEVANTADOR', 'FISCAL']), None)
        if not nome_col_loc: 
            st.error("Planilha de Localidades não tem coluna NOME.")
            st.stop()
            
        lat_locs = df_loc['LATITUDE'].values
        lon_locs = df_loc['LONGITUDE'].values
        nomes_locs = df_loc[nome_col_loc].values
        
        render_t(0.7, "Analisando Colaborador mais próximo...")
        def get_closest(lat, lon):
            if pd.isna(lat) or pd.isna(lon): return "DESCONHECIDO"
            dists = haversine_vectorized(lat, lon, lat_locs, lon_locs)
            if len(dists) == 0: return "DESCONHECIDO"
            min_idx = np.argmin(dists)
            return f"{nomes_locs[min_idx]}"

        df_master = pd.concat([df_san, df_lev], ignore_index=True)
        df_master['LAT_NUM'] = pd.to_numeric(df_master['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_master['LON_NUM'] = pd.to_numeric(df_master['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_master['LAT_NUM'] = df_master['LAT_NUM'].apply(lambda x: corrigir_coord(x, 90))
        df_master['LON_NUM'] = df_master['LON_NUM'].apply(lambda x: corrigir_coord(x, 180))
        df_master['LATITUDE'] = df_master['LAT_NUM']
        df_master['LONGITUDE'] = df_master['LON_NUM']
        
        df_valid = df_master.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
        
        render_t(0.9, "Agrupando Obras Vizinhas (Super Pontos)...")
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
        df_final['COLABORADOR MAIS PROXIMO'] = df_final.apply(lambda x: get_closest(x['LATITUDE'], x['LONGITUDE']), axis=1)

        def determinar_cor(linha):
            dupl = str(linha.get('DUPLICADA', ''))
            prox = str(linha.get('PROXIMA', ''))
            orig = str(linha.get('ORIGEM_BASE', ''))
            
            is_black = False
            if orig == 'LEVANTAMENTO':
                st_list = str(linha.get('STATUS LIST', '0')).strip().upper().replace('.0', '')
                if st_list in ['NAN', 'NONE', '']: st_list = '0'
                st_sisco = str(linha.get('STATUS SISCO', '0')).strip().upper().replace('.0', '')
                if st_sisco in ['NAN', 'NONE', '']: st_sisco = '0'
                
                v_list = ['0', 'EM LEVANTAMENTO', 'CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO']
                v_sisco = ['0', 'PRÉ ANALISE', 'PRE ANALISE', 'LIBERADO PARA LEVANTAMENTO', 'LIBERADO P/ LEVANTAMENTO']
                
                if st_list not in v_list or st_sisco not in v_sisco:
                    is_black = True
            
            if dupl == 'SIM': return 'red', '🔴 Vermelho (Duplicadas)'
            if prox == 'SIM': return 'orange', '🟠 Laranja (Próximas)'
            if is_black: return 'black', '⚫ Preto (Levant. Inválido)'
            if orig == 'LEVANTAMENTO': return 'green', '🟢 Verde (Levant. Solitário)'
            if orig == 'SANEAMENTO': return 'purple', '🟣 Magenta (Saneamento Solitário)'
            return 'blue', '🔵 Azul (Outros)'
            
        cores_calculadas = [determinar_cor(r) for _, r in df_final.iterrows()]
        df_final['COR_MAPA'] = [c[0] for c in cores_calculadas]
        df_final['COR_NOME'] = [c[1] for c in cores_calculadas]

        front_cols = ['NOTA', 'ORIGEM_BASE', 'SITUACAO SAP', 'DUPLICADA', 'PROXIMA', 'COLABORADOR MAIS PROXIMO', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'CLUSTER_ID', 'COR_MAPA', 'COR_NOME']
        rest_cols = [c for c in df_final.columns if c not in front_cols and not c.startswith('_')]
        df_final = df_final[front_cols + rest_cols]
        
        render_t(1.0, "✅ Análise Concluída!")
        time.sleep(1)
        
        st.session_state.df_final_analise = df_final
        st.session_state.is_done_analise = True
        st.rerun()
