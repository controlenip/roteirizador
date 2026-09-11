import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formata_campo_html, normalize_cols
from modules.geospatial import haversine_vectorized, fundir_super_pontos
from modules.export_analise import gerar_excel_analise, gerar_kml_analise

st.set_page_config(page_title="Análise Cruzada", page_icon="🔍", layout="wide")

# ==========================================
# FUNÇÕES VISUAIS E DE APOIO
# ==========================================
def corrigir_coord(val, limite):
    if pd.isna(val): return np.nan
    v = float(val)
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v

st.markdown("<h1 class='brand-title'>🔍 Análise Cruzada e Auditoria</h1>", unsafe_allow_html=True)
st.info("💡 Este módulo cruza as bases de **Saneamento** e **Levantamento**, identifica obras próximas ou duplicadas, e atrela o Colaborador mais perto geograficamente.")

if "df_final_analise" not in st.session_state: st.session_state.df_final_analise = pd.DataFrame()
if "is_done_analise" not in st.session_state: st.session_state.is_done_analise = False

# --- BARRA LATERAL (CONFIGS E DOWNLOAD) ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    raio_prox = st.slider("Distância p/ agrupar obras (Metros)", 10, 1000, 100, 10)
    
    st.markdown("---")
    
    # Se a análise estiver pronta, exibe os botões de Download igual às outras páginas
    if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
        df_fin = st.session_state.df_final_analise
        d_fmt = datetime.now().strftime("%d.%m.%Y_%H%M")
        
        bu_xl = io.BytesIO()
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
            out_fin = io.BytesIO()
            dfg = df_fin.drop(columns=['_ORIGINAL_ROWS'], errors='ignore')
            dfg.to_excel(out_fin, index=False)
            zx.writestr("Planilha_Analise_Cruzada.xlsx", out_fin.getvalue())
            
        st.download_button("🌐 Baixar Planilha Excel (ZIP)", data=bu_xl.getvalue(), file_name=f"Analise_Planilha_{d_fmt}.zip", use_container_width=True)
        
        kml_str = gerar_kml_analise(df_fin)
        bu_kml = io.BytesIO()
        with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
            zk.writestr("Mapa_Analise_Cruzada.kml", kml_str.encode('utf-8'))
            
        st.download_button("🗺️ Baixar Mapa (KML)", data=bu_kml.getvalue(), file_name=f"Analise_Mapa_{d_fmt}.zip", use_container_width=True)
        
        if st.button("🧹 Nova Análise", type="primary", use_container_width=True):
            st.session_state.is_done_analise = False
            st.session_state.df_final_analise = pd.DataFrame()
            st.rerun()

# ==========================================
# EXIBIÇÃO DE RESULTADOS
# ==========================================
if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
    df_final = st.session_state.df_final_analise
    
    st.markdown("### 🗺️ Mapa Analítico")
    
    st.markdown("""
    **Legenda de Cores:**
    *   🔴 **Vermelho:** Obras Duplicadas (Mesma NOTA nas duas planilhas).
    *   🟠 **Amarelo (Laranja):** Obras Próximas.
    *   🟢 **Verde:** Obra Solitária do Levantamento.
    *   🟣 **Magenta (Roxo):** Obra Solitária do Saneamento.
    *   🔵 **Azul:** Obra sem identificação (Fallback).
    """)
    
    mapa = folium.Map(location=[df_final['LATITUDE'].mean(), df_final['LONGITUDE'].mean()], zoom_start=8) if not df_final.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    
    fg = folium.FeatureGroup(name="Obras Analisadas").add_to(mapa)
    
    for _, r in df_final.iterrows():
        lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
        
        duplicada = str(r.get('DUPLICADA', ''))
        proxima = str(r.get('PROXIMA', ''))
        origem = str(r.get('ORIGEM_BASE', ''))
        
        if duplicada == 'SIM': c_i, ic = 'red', 'info-sign'
        elif proxima == 'SIM': c_i, ic = 'orange', 'info-sign'
        elif origem == 'LEVANTAMENTO': c_i, ic = 'green', 'info-sign'
        elif origem == 'SANEAMENTO': c_i, ic = 'purple', 'info-sign'
        else: c_i, ic = 'blue', 'info-sign'
        
        nota = str(r.get('NOTA', ''))
        status_sap = str(r.get('SITUACAO SAP', ''))
        colab = str(r.get('COLABORADOR MAIS PROXIMO', ''))
        mun = str(r.get('MUNICIPIO', ''))
        
        pop_html = f'''<![CDATA[
        <div style="width:250px; font-size:12px;">
            <b>Nota:</b> {html.escape(nota)}<br>
            <b>Origem:</b> {html.escape(origem)}<br>
            <b>Status:</b> {html.escape(status_sap)}<br>
            <b>Colab Perto:</b> {html.escape(colab)}<br>
            <b>Duplicada:</b> {html.escape(duplicada)} | <b>Próxima:</b> {html.escape(proxima)}
        </div>
        ]]>'''
        pop_html = pop_html.replace("{", "&#123;").replace("}", "&#125;")
        
        folium.Marker([lat, lon], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(fg)
        
    folium.LayerControl().add_to(mapa)
    st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📊 Tabela Consolidada (Obras e Colaboradores)")
    st.data_editor(df_final.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM'], errors='ignore'), use_container_width=True)

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
        with st.spinner("Analisando e cruzando bases de dados..."):
            df_san = ler_planilha_cached(file_san.getvalue()) if not file_san.name.endswith('.csv') else pd.read_csv(file_san)
            df_lev = ler_planilha_cached(file_lev.getvalue()) if not file_lev.name.endswith('.csv') else pd.read_csv(file_lev)
            df_loc = ler_planilha_cached(file_loc.getvalue()) if not file_loc.name.endswith('.csv') else pd.read_csv(file_loc)
            
            df_san.columns = normalize_cols(df_san.columns)
            df_lev.columns = normalize_cols(df_lev.columns)
            df_loc.columns = normalize_cols(df_loc.columns)
            
            if 'LATITUDE_PROJETO' in df_san.columns: df_san.rename(columns={'LATITUDE_PROJETO': 'LATITUDE'}, inplace=True)
            if 'LONGITUDE_PROJETO' in df_san.columns: df_san.rename(columns={'LONGITUDE_PROJETO': 'LONGITUDE'}, inplace=True)
            
            # Map NOTA
            col_n_san = next((c for c in df_san.columns if c in ['NOTA', 'PROTOCOLO', 'ID SISCO']), None)
            if col_n_san: df_san.rename(columns={col_n_san: 'NOTA'}, inplace=True)
            col_n_lev = next((c for c in df_lev.columns if c in ['NOTA', 'PROTOCOLO', 'ID SISCO']), None)
            if col_n_lev: df_lev.rename(columns={col_n_lev: 'NOTA'}, inplace=True)
            
            df_san['ORIGEM_BASE'] = 'SANEAMENTO'
            df_lev['ORIGEM_BASE'] = 'LEVANTAMENTO'
            
            df_san['NOTA'] = df_san['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
            df_lev['NOTA'] = df_lev['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            # --- IDENTIFICAÇÃO DO STATUS SAP (CANC / FINL) ---
            status_col = 'STATUS SAP' if 'STATUS SAP' in df_lev.columns else ('STATUS' if 'STATUS' in df_lev.columns else None)
            status_dict = df_lev.set_index('NOTA')[status_col].to_dict() if status_col else {}

            def get_situacao(nota):
                s = str(status_dict.get(nota, '')).strip().upper()
                if s in ['FINL', 'CANC']: return f"BLOQUEADO ({s})"
                return "APTO"

            df_san['SITUACAO SAP'] = df_san['NOTA'].apply(get_situacao)
            df_lev['SITUACAO SAP'] = df_lev['NOTA'].apply(get_situacao)
            
            # --- DUPLICIDADE CRUZADA ---
            notas_san = set(df_san['NOTA'].dropna())
            notas_lev = set(df_lev['NOTA'].dropna())
            duplicadas = notas_san.intersection(notas_lev)
            
            df_san['DUPLICADA'] = df_san['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
            df_lev['DUPLICADA'] = df_lev['NOTA'].apply(lambda x: 'SIM' if x in duplicadas else 'NÃO')
            
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
            
            # --- COLABORADOR MAIS PROXIMO ---
            def get_closest(lat, lon):
                if pd.isna(lat) or pd.isna(lon): return "DESCONHECIDO"
                dists = haversine_vectorized(lat, lon, lat_locs, lon_locs)
                if len(dists) == 0: return "DESCONHECIDO"
                min_idx = np.argmin(dists)
                return f"{nomes_locs[min_idx]} - EQUIPE LEVANTAMENTO"

            df_master = pd.concat([df_san, df_lev], ignore_index=True)
            df_master['LAT_NUM'] = pd.to_numeric(df_master['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
            df_master['LON_NUM'] = pd.to_numeric(df_master['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
            df_master['LAT_NUM'] = df_master['LAT_NUM'].apply(lambda x: corrigir_coord(x, 90))
            df_master['LON_NUM'] = df_master['LON_NUM'].apply(lambda x: corrigir_coord(x, 180))
            df_master['LATITUDE'] = df_master['LAT_NUM']
            df_master['LONGITUDE'] = df_master['LON_NUM']
            
            df_valid = df_master.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
            
            # --- PROXIMIDADE (SUPER PONTOS) ---
            df_clustered, _ = fundir_super_pontos(df_valid, raio_metros=st.session_state.get('raio_prox', 50), agrupar_por_levantador=False)
            
            expanded = []
            for _, r in df_clustered.iterrows():
                is_prox = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
                if is_prox:
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = orig.copy()
                        nr['PROXIMA'] = 'SIM'
                        expanded.append(nr)
                else:
                    nr = r.copy()
                    nr['PROXIMA'] = 'NÃO'
                    expanded.append(nr)
                    
            df_final = pd.DataFrame(expanded)
            df_final['COLABORADOR MAIS PROXIMO'] = df_final.apply(lambda x: get_closest(x['LATITUDE'], x['LONGITUDE']), axis=1)

            front_cols = ['NOTA', 'ORIGEM_BASE', 'SITUACAO SAP', 'DUPLICADA', 'PROXIMA', 'COLABORADOR MAIS PROXIMO', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE']
            rest_cols = [c for c in df_final.columns if c not in front_cols and not c.startswith('_')]
            df_final = df_final[front_cols + rest_cols]
            
            st.session_state.df_final_analise = df_final
            st.session_state.is_done_analise = True
            st.rerun()
