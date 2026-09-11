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
from modules.geospatial import haversine_vectorized, haversine_scalar, fundir_super_pontos, obter_coordenadas_municipio_cached
from modules.routing_engine import resolver_tsp_ortools

from modules.export_analise import injetar_logo, gerar_excel_analise, gerar_kml_analise, gerar_gpx_simples

st.set_page_config(page_title="Análise Cruzada", page_icon="🔍", layout="wide")

# CSS para o multiselect e visualização
st.markdown("""
<style>
    .stMultiSelect [data-baseweb="select"] > div:first-child { flex-wrap: wrap !important; }
    .stMultiSelect [data-baseweb="select"] > div:first-child > div:last-child { display: none !important; }
    .stMultiSelect [data-baseweb="tag"] { max-width: 100% !important; }
</style>
""", unsafe_allow_html=True)

injetar_logo()

def corrigir_coord(val, limite):
    if pd.isna(val): return np.nan
    v = float(val)
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} KM"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

st.markdown("<h1 class='brand-title'>🔍 Análise Cruzada e Auditoria</h1>", unsafe_allow_html=True)
st.info("💡 Este módulo cruza as bases de **Saneamento** e **Levantamento**, bloqueia notas com SAP 'FINL' ou 'CANC' e identifica obras de diferentes bases no mesmo endereço geográfico.")

# ==========================================
# 1. ENTRADA DOS 3 ARQUIVOS
# ==========================================
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

# ==========================================
# 2. PROCESSAMENTO E CRUZAMENTO DAS BASES
# ==========================================
with st.spinner("Processando e cruzando bases de dados..."):
    # Lê os 3 arquivos
    df_san = ler_planilha_cached(file_san.getvalue())
    df_lev = ler_planilha_cached(file_lev.getvalue())
    df_loc = ler_planilha_cached(file_loc.getvalue())
    
    # Normaliza colunas
    df_san.columns = normalize_cols(df_san.columns)
    df_lev.columns = normalize_cols(df_lev.columns)
    df_loc.columns = normalize_cols(df_loc.columns)
    
    # Padroniza nomes de colunas cruciais
    if 'LATITUDE_PROJETO' in df_san.columns: df_san.rename(columns={'LATITUDE_PROJETO': 'LATITUDE'}, inplace=True)
    if 'LONGITUDE_PROJETO' in df_san.columns: df_san.rename(columns={'LONGITUDE_PROJETO': 'LONGITUDE'}, inplace=True)
    if 'PROTOCOLO' in df_lev.columns: df_lev.rename(columns={'PROTOCOLO': 'NOTA'}, inplace=True)
    
    df_san['ORIGEM_BASE'] = 'SANEAMENTO'
    df_lev['ORIGEM_BASE'] = 'LEVANTAMENTO'
    
    # Limpa as colunas de NOTA para string exata
    df_san['NOTA'] = df_san['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
    df_lev['NOTA'] = df_lev['NOTA'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    # --- REGRA DE NEGÓCIO: BLOQUEIO STATUS SAP ---
    # Captura o status SAP do Levantamento
    if 'STATUS_SAP' in df_lev.columns:
        status_sap_dict = df_lev.set_index('NOTA')['STATUS_SAP'].to_dict()
    elif 'STATUS' in df_lev.columns:
        status_sap_dict = df_lev.set_index('NOTA')['STATUS'].to_dict()
    else:
        status_sap_dict = {}

    def avaliar_status(nota, origem, dict_sap):
        status_atual = str(dict_sap.get(nota, '')).strip().upper()
        if status_atual in ['FINL', 'CANC']:
            return f"⛔ BLOQUEADA - SAP {status_atual}"
        return "✅ APTO PARA CAMPO"

    # Aplica a trava nas duas planilhas
    df_san['SITUACAO_DE_CAMPO'] = df_san['NOTA'].apply(lambda x: avaliar_status(x, 'SAN', status_sap_dict))
    df_lev['SITUACAO_DE_CAMPO'] = df_lev['NOTA'].apply(lambda x: avaliar_status(x, 'LEV', status_sap_dict))
    
    # Identificando notas exatas duplicadas em ambas as bases
    notas_comuns = set(df_san['NOTA']).intersection(set(df_lev['NOTA']))
    df_san['DUPLICIDADE_CRUZADA'] = df_san['NOTA'].apply(lambda x: 'SIM (Nas 2 Bases)' if x in notas_comuns else 'NÃO')
    df_lev['DUPLICIDADE_CRUZADA'] = df_lev['NOTA'].apply(lambda x: 'SIM (Nas 2 Bases)' if x in notas_comuns else 'NÃO')

# ==========================================
# 3. CONSOLIDAÇÃO E FILTRAGEM GEOGRÁFICA
# ==========================================
st.markdown("---")
st.markdown("### 🚦 Filtro Geográfico e Montagem")

raio_prox = st.slider("Distância para agrupar obras vizinhas (Metros)", 10, 500, 50, 10)

if st.button("🚀 Processar Rotas e Vizinhança", type="primary", use_container_width=True):
    # Separa apenas os APTOS para jogar no mapa
    df_san_apto = df_san[df_san['SITUACAO_DE_CAMPO'] == "✅ APTO PARA CAMPO"].copy()
    df_lev_apto = df_lev[df_lev['SITUACAO_DE_CAMPO'] == "✅ APTO PARA CAMPO"].copy()
    
    # Funde as duas bases num Master Dataframe
    df_master = pd.concat([df_san_apto, df_lev_apto], ignore_index=True)
    
    if df_master.empty:
        st.error("Nenhuma obra Apta sobrou após o filtro do SAP.")
        st.stop()

    # Arruma as coordenadas
    df_master['LATITUDE'] = pd.to_numeric(df_master['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_master['LONGITUDE'] = pd.to_numeric(df_master['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_master['LATITUDE'] = df_master['LATITUDE'].apply(lambda x: corrigir_coord(x, 90))
    df_master['LONGITUDE'] = df_master['LONGITUDE'].apply(lambda x: corrigir_coord(x, 180))
    df_master = df_master.dropna(subset=['LATITUDE', 'LONGITUDE'])

    # --- REGRA DE NEGÓCIO: AGRUPAMENTO (Obras Próximas) ---
    # Usamos o algoritmo de Super Pontos para encontrar quem está grudado
    df_master, _ = fundir_super_pontos(df_master, raio_metros=raio_prox, agrupar_por_levantador=False)
    
    # Descobre os levantadores e suas bases
    bases_dict = {}
    col_nome = next((c for c in df_loc.columns if c in ['NOME', 'LEVANTADOR', 'FISCAL']), None)
    if col_nome:
        for _, r in df_loc.iterrows():
            nome = str(r[col_nome]).strip().upper()
            lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
            if pd.notna(lat) and pd.notna(lon):
                bases_dict[nome] = (float(lat), float(lon))
                
    # Atribuição por Proximidade ao Levantador
    assigned_tasks = []
    for r in df_master.to_dict('records'):
        la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
        best_f, best_d = None, float('inf')
        
        for f_name, coords in bases_dict.items():
            d = haversine_scalar(la, lo, coords[0], coords[1])
            if d < best_d:
                best_d = d
                best_f = f_name
                
        r['BASE_ATRIBUIDA'] = best_f if best_f else "NÃO ALOCADO"
        assigned_tasks.append(r)
        
    df_final = pd.DataFrame(assigned_tasks)
    
    # ==========================================
    # 4. GERAÇÃO DE RELATÓRIOS E DOWNLOADS
    # ==========================================
    st.success("✅ Cruzamento e Análise Espacial Concluídos com Sucesso!")
    
    c_res1, c_res2, c_res3 = st.columns(3)
    c_res1.markdown(f"**Total Saneamento:** {len(df_san)}")
    c_res2.markdown(f"**Total Levantamento:** {len(df_lev)}")
    c_res3.markdown(f"**Bloqueadas (FINL/CANC):** {len(df_san[df_san['SITUACAO_DE_CAMPO'].str.contains('BLOQUEADA')]) + len(df_lev[df_lev['SITUACAO_DE_CAMPO'].str.contains('BLOQUEADA')])}")

    # Criação do ZIP
    bu_xl = io.BytesIO()
    with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
        # Exporta a Master Planilha de Validação
        out_san = io.BytesIO(); df_san.to_excel(out_san, index=False); zx.writestr("Auditoria_Base_Saneamento.xlsx", out_san.getvalue())
        out_lev = io.BytesIO(); df_lev.to_excel(out_lev, index=False); zx.writestr("Auditoria_Base_Levantamento.xlsx", out_lev.getvalue())
        
        # Exporta Obras Liberadas
        dfg = df_final.drop(columns=['_ORIGINAL_ROWS', 'COORD_KEY'], errors='ignore')
        out_fin = io.BytesIO(); dfg.to_excel(out_fin, index=False); zx.writestr("Obras_Liberadas_Agrupadas.xlsx", out_fin.getvalue())

    st.download_button("📥 Baixar Relatórios de Auditoria e Obras Liberadas (ZIP)", data=bu_xl.getvalue(), file_name=f"Auditoria_Cruzada_{datetime.now().strftime('%d%m%Y')}.zip", type="primary", use_container_width=True)

    # Exibe visualmente o mapa das obras agrupadas
    st.markdown("### 🗺️ Mapa de Convergência (Obras Válidas)")
    mapa = folium.Map(location=[df_final['LATITUDE'].mean(), df_final['LONGITUDE'].mean()], zoom_start=8)
    m_clust = MarkerCluster(name="📍 Obras Liberadas").add_to(mapa)
    
    for _, r in df_final.iterrows():
        lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
        
        is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
        nota = str(r.get('NOTA', ''))
        
        if is_sp:
            origens = [str(o.get('ORIGEM_BASE', '')) for o in r['_ORIGINAL_ROWS']]
            if 'SANEAMENTO' in origens and 'LEVANTAMENTO' in origens:
                c_i = 'orange'
                titulo = f"⚠️ CONFLITO GEOGRÁFICO ({len(origens)} Obras)"
            else:
                c_i = 'purple'
                titulo = f"🏢 Obras Sobrepostas ({len(origens)})"
        else:
            c_i = 'blue' if r.get('ORIGEM_BASE') == 'SANEAMENTO' else 'green'
            titulo = f"📍 {r.get('ORIGEM_BASE')} - {nota}"
            
        pop_html = f'<div style="width:250px;"><b>{titulo}</b><br>Atribuído a: {r.get("BASE_ATRIBUIDA")}<br>Nota: {nota}</div>'
        folium.Marker([lat, lon], icon=folium.Icon(color=c_i), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        
    st_folium(mapa, use_container_width=True, height=500)
