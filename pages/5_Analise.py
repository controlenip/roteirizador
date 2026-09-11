import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
import requests
import gc
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formata_campo_html, normalize_cols
from modules.geospatial import haversine_vectorized, haversine_scalar
from modules.routing_engine import resolver_tsp_ortools

# IMPORTAÇÃO DA EXPORTAÇÃO DE ANÁLISE
from modules.export_analise import injetar_logo, gerar_excel_analise, gerar_excel_resumo_analise, gerar_gpx_simples, gerar_kml_analise, limpar_colunas_analise, gerar_txt_analise

st.set_page_config(page_title="Análise de Rotas", page_icon="📊", layout="wide")

# CSS para melhorar o multiselect
st.markdown("""
<style>
    .stMultiSelect [data-baseweb="select"] { min-height: 42px !important; }
    .stMultiSelect [data-baseweb="select"] > div:first-child {
        display: flex !important; flex-wrap: wrap !important; height: auto !important;
        max-height: none !important; overflow-y: visible !important; padding-bottom: 5px !important;
    }
    .stMultiSelect [data-baseweb="select"] > div:first-child > div:last-child { display: none !important; }
    .stMultiSelect [data-baseweb="tag"] { max-width: 100% !important; margin-bottom: 4px !important; margin-top: 4px !important; }
    .stMultiSelect [data-baseweb="tag"] span { white-space: normal !important; }
</style>
""", unsafe_allow_html=True)

injetar_logo()

# ==========================================
# FUNÇÕES DE CORREÇÃO E APOIO
# ==========================================
def corrigir_coord(val, limite):
    if pd.isna(val): return np.nan
    v = float(val)
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v

def obter_rota_osrm_robusta(lat1, lon1, lat2, lon2, base_url):
    url = f"{base_url}/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson"
    headers = {'User-Agent': 'NIP-Roteirizador/3.0'}
    for _ in range(3):
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                if data.get('code') == 'Ok' and len(data.get('routes', [])) > 0:
                    return data['routes'][0]['geometry']['coordinates'], data['routes'][0]['duration']
            elif resp.status_code == 429: time.sleep(2)
        except Exception:
            time.sleep(1)
    return None, None

def render_metric_card(title, value, icon, border_color, bg_color):
    return f"""
    <div style="background-color: #ffffff; border-left: 5px solid {border_color}; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); display: flex; align-items: center; margin-bottom: 10px;">
        <div style="background-color: {bg_color}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; font-size: 20px; margin-right: 15px;">{icon}</div>
        <div><p style="margin: 0; font-size: 11px; color: #666; text-transform: uppercase; font-weight: bold;">{title}</p><p style="margin: 0; font-size: 22px; color: #333; font-weight: bold;">{value}</p></div>
    </div>
    """

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} KM"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_analise': False, 'vrp_status_analise': "IDLE", 'vrp_state_analise': {}, 'df_routed_analise': pd.DataFrame(), 'colunas_exibir_analise': [], 'colunas_originais_analise': []})
    for k in ['bytes_zip_xl_analise', 'bytes_zip_kml_analise', 'bytes_zip_gpx_analise', 'bytes_zip_txt_analise', 'start_time_run_analise', 'df_correcao_analise']: st.session_state.pop(k, None)
    ler_planilha_cached.clear(); tentar_rerun()

# ==========================================
# CONTROLE DE SESSÃO
# ==========================================
if "roteamento_concluido_analise" not in st.session_state: st.session_state.roteamento_concluido_analise = False
if "vrp_status_analise" not in st.session_state: st.session_state.vrp_status_analise = "IDLE"

status_exec = st.session_state.vrp_status_analise
is_done = st.session_state.roteamento_concluido_analise
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📊 Análise Espacial de Rotas</h1>", unsafe_allow_html=True)
st.info("💡 Este módulo é um laboratório genérico de roteamento. Ele aceita planilhas de qualquer formato, limpa as coordenadas e roteia em fluxo contínuo. Ideal para estudos de viabilidade.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    with st.expander("Parâmetros do Laboratório", expanded=True):
        trava_global = st.number_input("Travar qtd. de Obras", min_value=0, value=0, step=50, disabled=is_locked)
        sentido_rota = st.radio("Sentido:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        vel_kmh = st.slider("Velocidade (km/h)", 10, 80, 30, 5, disabled=is_locked)
        st.markdown("---")
        
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)

    st.markdown("---")
    
    if is_done and not st.session_state.df_routed_analise.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y_%H%M")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_analise', b"vazio"), file_name=f"Analise_Planilhas_{d_fmt}.zip", use_container_width=True)
        st.download_button("📝 Baixar Relatórios (TXT)", data=st.session_state.get('bytes_zip_txt_analise', b"vazio"), file_name=f"Analise_TXT_{d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_analise', b"vazio"), file_name=f"Analise_Mapas_{d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_analise', b"vazio"), file_name=f"Analise_GPS_{d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Análise", type="primary", use_container_width=True): limpar_roteirizador()

# ==========================================
# EXIBIÇÃO DE RESULTADOS (SE CONCLUÍDO)
# ==========================================
if is_done and not st.session_state.df_routed_analise.empty:
    st.markdown("## 🎯 Resultado da Análise Espacial")
    
    df_c = st.session_state.get('df_correcao_analise', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Linhas Retidas (Coordenadas Inválidas)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>Foram separadas no ZIP de Correção.</p>
        </div>
        """, unsafe_allow_html=True)

    dfr = st.session_state.df_routed_analise.copy()
    
    tr = len(dfr)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Pontos Analisados", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Grupos / Fiscais", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("Status", "Concluído", "✅", "#28a745", "rgba(40,167,69,0.15)"), unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Analítico")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        
        bn_safe = str(bn).replace("{", "[").replace("}", "]")
        fg = folium.FeatureGroup(name=f"Rota: {bn_safe}", show=False)
        
        pts = [p for _, r in db.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
        folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
        folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
        
        for r in db.to_dict('records'):
            c_i = 'blue'
            er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_analise if c not in ['BASE_ATRIBUIDA']])
            pop_html = f'<div style="width:250px;"><b>Grupo/Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
            pop_html = pop_html.replace("{", "&#123;").replace("}", "&#125;")
            folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon='info-sign'), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📊 Dados Tabulares")
    st.data_editor(st.session_state.df_routed_analise.drop(columns=['ROTA_GEOMETRIA', 'PERIODO'], errors='ignore'), use_container_width=True)

# ==========================================
# START DA APLICAÇÃO (UPLOAD DE DADOS)
# ==========================================
elif status_exec == "IDLE":
    st.markdown("### 📁 Fonte de Dados (Genérica)")
    st.info("💡 Faça upload de qualquer planilha. Identificaremos as colunas de Latitude e Longitude automaticamente. Você poderá escolher como dividir as equipes.")
    task_files = st.file_uploader("Suba a planilha para Análise", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if not task_files: st.stop()
    
    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        if not dfs: st.session_state.colunas_originais_analise = dft.columns.tolist()
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)

    st.markdown("---")
    
    # Identifica colunas do dataframe para o usuário mapear se precisar
    colunas_disp = df_tasks.columns.tolist()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        col_lat = st.selectbox("🌎 Coluna de Latitude", ["(Auto-Detectar)"] + colunas_disp, index=0)
    with col2:
        col_lon = st.selectbox("🌎 Coluna de Longitude", ["(Auto-Detectar)"] + colunas_disp, index=0)
    with col3:
        col_grupo = st.selectbox("👥 Dividir por Grupo/Equipe", ["(Equipe Única)"] + colunas_disp, index=0)
        
    # Auto-detect LAT/LON se o usuário pedir
    if col_lat == "(Auto-Detectar)":
        for c in colunas_disp:
            if 'LAT' in c.upper(): col_lat = c; break
    if col_lon == "(Auto-Detectar)":
        for c in colunas_disp:
            if 'LON' in c.upper(): col_lon = c; break

    if col_lat not in df_tasks.columns or col_lon not in df_tasks.columns:
        st.error("❌ Não foi possível encontrar as colunas de Latitude e Longitude. Mapeie manualmente acima.")
        st.stop()
        
    # Renomeia para o motor
    df_tasks = df_tasks.rename(columns={col_lat: 'LATITUDE', col_lon: 'LONGITUDE'})
    
    if col_grupo != "(Equipe Única)":
        df_tasks = df_tasks.rename(columns={col_grupo: 'BASE_ATRIBUIDA'})
    else:
        df_tasks['BASE_ATRIBUIDA'] = "EQUIPE_UNICA_ANALISE"
        
    df_tasks['BASE_ATRIBUIDA'] = df_tasks['BASE_ATRIBUIDA'].astype(str).str.strip().str.upper()
    df_tasks = df_tasks[~df_tasks['BASE_ATRIBUIDA'].isin(['NAN', 'NONE', ''])]

    st.markdown("#### 🌍 Limpeza Geográfica Automática")
    if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    
    df_tasks['LAT_NUM'] = df_tasks['LAT_NUM'].apply(lambda x: corrigir_coord(x, 90))
    df_tasks['LON_NUM'] = df_tasks['LON_NUM'].apply(lambda x: corrigir_coord(x, 180))

    m_na, m_0 = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna(), (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    df_tasks.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df_tasks.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)
    df_tasks.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    m_inv = abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])
    df_tasks.loc[m_inv & ~m_na & ~m_0 & ~m_pos, 'MOTIVO_REJEICAO'] = 'Coordenada Invertida'
    
    mc = m_na | m_0 | m_pos | m_inv
    if mc.sum() > 0: df_rej = pd.concat([df_rej, df_tasks[mc].copy()], ignore_index=True); df_tasks = df_tasks[~mc].copy()
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']; df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    st.session_state.df_correcao_analise = df_rej
    
    if not df_rej.empty: 
        st.warning(f"⚠️ {len(df_rej)} linhas foram ignoradas por coordenadas inválidas.")

    if df_tasks.empty: st.error("🚨 Nenhuma linha válida restou."); st.stop()

    if trava_global > 0: df_tasks = df_tasks.head(trava_global)

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_tasks.columns if not c.startswith('_')]
        colunas_exibir = st.multiselect("Colunas Visíveis no Mapa:", tc, default=tc[:7] if len(tc)>7 else tc)

    if st.button("🚀 Iniciar Análise de Rota Espacial", type="primary", use_container_width=True):
        st.session_state.update({'colunas_exibir_analise': colunas_exibir})
        st.session_state.vrp_state_analise = {
            'config': {
                'velocidade_media_kmh': vel_kmh, 
                'sentido_rota': sentido_rota, 
                'url_osrm_base': url_osrm, 
                'tracado_real': usa_osrm
            }, 
            'b_names': list(set(df_tasks['BASE_ATRIBUIDA'].unique())), 
            'b_idx': 0, 
            'unvisited': df_tasks.copy(), 
            'routed_data': [], 
            'current_geoms': []
        }
        st.session_state.vrp_status_analise = "RUNNING"; tentar_rerun()

# ==========================================
# CÁLCULO VRP CONTÍNUO
# ==========================================
if status_exec == "RUNNING":
    st.markdown("## 🚀 Executando Motor de Análise")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_analise', time.time())
    if 'start_time_run_analise' not in st.session_state: st.session_state.start_time_run_analise = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_analise; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Analisando Cluster: **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            bl = sum(float(x['LATITUDE']) for x in oe) / len(oe)
            bL = sum(float(x['LONGITUDE']) for x in oe) / len(oe)
            
            ot = []
            if oe:
                if "Varredura Reversa" in cfg.get('sentido_rota', "Lógica Padrão"):
                    max_idx = max(range(len(oe)), key=lambda i: haversine_scalar(bl, bL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                    p_longe = oe.pop(max_idx)
                    ot.append(p_longe)
                    cl, cL = float(p_longe['LATITUDE']), float(p_longe['LONGITUDE'])
                else:
                    cl, cL = bl, bL
                    
                while oe:
                    closest_idx = min(range(len(oe)), key=lambda i: haversine_scalar(cl, cL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                    nx = oe.pop(closest_idx)
                    ot.append(nx)
                    cl, cL = float(nx['LATITUDE']), float(nx['LONGITUDE'])
            
            rf = []
            c_l, c_L = bl, bL
            for o in ot:
                vkr = haversine_vectorized(c_l, c_L, o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                rf.append({'o': o, 'la': c_l, 'La': c_L, 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 'dk': vk})
                c_l, c_L = o['LATITUDE'], o['LONGITUDE']
                
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_analise = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (30 if cfg.get('tracado_real') else len(rf)), len(rf)) 
            
            for i in range(oi, ei):
                it = rf[i]
                fallback = ([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600)
                
                if not cfg.get('tracado_real'):
                    gd.append(fallback)
                else:
                    if i % 5 == 0: sgt.info(f"🛣️ Traçando arruamento **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf))
                    
                    sucesso_rota = False
                    for tentativa in range(3):
                        try:
                            time.sleep(0.3)
                            coords, dur = obter_rota_osrm_robusta(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'])
                            if coords:
                                gd.append((coords, dur))
                                sucesso_rota = True
                                break
                        except Exception:
                            time.sleep(1.0)
                            
                    if not sucesso_rota:
                        gd.append(fallback)
                
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_analise = st_v; tentar_rerun(); st.stop()
            
            rdf, og = [], 1
            for it, (g, ds) in zip(rf, gd):
                ob = it['o']; ob['ORDEM'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, round(it['dk'], 2)
                ob['ROTA_GEOMETRIA'], ob['PERIODO'] = g, "Único"
                rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state_analise = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Análise Finalizada!"); pb.progress(1.0)
        st.session_state.df_routed_analise = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status_analise = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    
    df_routed = st.session_state.df_routed_analise.copy()

    d_fmt = datetime.now().strftime("%d.%m.%Y_%H%M")
    bu_xl, bu_kml, bu_gpx, bu_txt = io.BytesIO(), io.BytesIO(), io.BytesIO(), io.BytesIO()
    
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, \
             zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, \
             zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg, \
             zipfile.ZipFile(bu_txt, 'w', zipfile.ZIP_DEFLATED) as zt:
            
            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b)]
                res.append({
                    'Grupo / Equipe': b, 
                    'Qtd Pontos': len(db), 
                    'KM Total Previsto': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)
                })
            zx.writestr(f"Resumo_Analise_{d_fmt}.xlsx", gerar_excel_resumo_analise(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_analise', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy()
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Linhas_Correcao_{d_fmt}.xlsx", out_e.getvalue())
            
            dfg_total = limpar_colunas_analise(df_routed.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), st.session_state.colunas_originais_analise)
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Analise_Total_{d_fmt}.xlsx", gerar_excel_analise(dfg_total, st.session_state.colunas_originais_analise))
            
            txt_total = gerar_txt_analise(df_routed)
            zt.writestr(f"Analise_Total_{d_fmt}.txt", txt_total.encode('utf-8'))
            
            dfk_total = df_routed.copy()
            if not dfk_total.empty:
                col_exibir = st.session_state.colunas_exibir_analise.copy()
                ks_tot = gerar_kml_analise(dfk_total, "ANALISE TOTAL", col_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), formatar_valor_coluna)
                zk.writestr(f"ANALISE_TOTAL_{d_fmt}.kml", ks_tot.encode('utf-8'))
                zg.writestr(f"GPS_TOTAL_{d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ANALISE TOTAL").encode('utf-8'))

            for base in df_routed['BASE_ATRIBUIDA'].unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                df_base_excel = df_routed[df_routed['BASE_ATRIBUIDA'] == base]
                
                if not df_base_excel.empty:
                    dfg = limpar_colunas_analise(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), st.session_state.colunas_originais_analise)
                    for cc in dfg.columns:
                        if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                    zx.writestr(f"Rotas/Rota_{b_safe}.xlsx", gerar_excel_analise(dfg, st.session_state.colunas_originais_analise))
                    
                    txt_ind = gerar_txt_analise(df_base_excel)
                    if txt_ind: zt.writestr(f"Relatorios_TXT/Relatorio_{b_safe}.txt", txt_ind.encode('utf-8'))
                    
                    ks = gerar_kml_analise(df_base_excel, f"Rota {b_safe}", col_exibir, [base], formatar_valor_coluna)
                    zk.writestr(f"KML/Rota_{b_safe}.kml", ks.encode('utf-8'))
                    zg.writestr(f"GPX/Rota_{b_safe}.gpx", gerar_gpx_simples(df_base_excel, f"Rota {b_safe}").encode('utf-8'))

        st.session_state.bytes_zip_xl_analise = bu_xl.getvalue()
        st.session_state.bytes_zip_kml_analise = bu_kml.getvalue()
        st.session_state.bytes_zip_gpx_analise = bu_gpx.getvalue()
        st.session_state.bytes_zip_txt_analise = bu_txt.getvalue()
        st.session_state.roteamento_concluido_analise = True; st.session_state.vrp_status_analise = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_analise = "IDLE"
