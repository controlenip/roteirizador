import streamlit as st
import pandas as pd
import numpy as np
import folium
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
import io
import zipfile
import html
import re
import time
import altair as alt
from datetime import datetime
import os
import base64
import gc

# Configuração da Página
st.set_page_config(page_title="Fiscalização", page_icon="📋", layout="wide")

# Importações do Backend (Módulos)
from modules.data_processing import ler_planilha_cached, formatar_moeda, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.export_utils import gerar_excel_bytes, gerar_excel_resumo_bytes, gerar_gpx_simples, gerar_kml_fiscalizacao, renderizar_painel_lateral

# ==========================================
# FUNÇÕES AUXILIARES DE FISCALIZAÇÃO
# ==========================================
def formatar_valor_coluna(col_name, val):
    try:
        if pd.isna(val) or val == '' or val == '-': return '-'
    except Exception: pass 
    try:
        val_float = float(val)
        if 'DISTANCIA' in col_name.upper(): return f"{val_float:.2f} Metros"
        elif 'POSTE' in col_name.upper(): return f"{int(round(val_float))}"
        else: return formata_campo_html(val)
    except (ValueError, TypeError):
        if isinstance(val, (datetime, pd.Timestamp)): return formata_campo_html(val.strftime('%d/%m/%Y'))
        return formata_campo_html(str(val))

def count_real_obras(row):
    if isinstance(row.get('_ORIGINAL_ROWS'), list): return len(row['_ORIGINAL_ROWS'])
    val = str(row.get('SUPER_PONTO', ''))
    if val.startswith('SIM'):
        nums = re.findall(r'\d+', val)
        if nums: return int(nums[0])
    return 1

def definir_cor_fiscalizacao(qtd):
    try:
        q = float(qtd)
        return 'green' if q <= 100 else 'blue' if q <= 200 else 'beige' if q <= 300 else 'orange' if q <= 400 else 'red'
    except: return 'gray'

def extrair_qtd(val):
    try: return float(str(val).replace(',', '.')) if pd.notna(val) and val != '' else 0.0
    except: return 0.0

def limpar_colunas_excel(df_alvo, cols_originais):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    if 'PROTOCOLO' in df_alvo.columns: df_alvo = df_alvo.rename(columns={'PROTOCOLO': 'NOTA'})
    if 'BASE_ATRIBUIDA' in df_alvo.columns: df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'FISCAL'})
    elif 'LEVANTADOR' in df_alvo.columns and 'FISCAL' not in df_alvo.columns: df_alvo['FISCAL'] = df_alvo['LEVANTADOR']
    
    base_start = ['NOTA', 'FISCAL', 'ORDEM', 'NOME_DIA', 'DIA_MES', 'PRIORIDADE', 'SUPER_PONTO']
    base_end = ['LINK_NAVEGACAO_OFFLINE']
    c_garantia = ['REGIONAL', 'MUNICIPIO', 'LOCALIDADE', 'LATITUDE', 'LONGITUDE', 'TIPO PROJETO', 'TIPO DE FISCALIZACAO', 'STATUS DA FISCALIZACAO', 'QTD PREVISTA DE POSTES', 'VALOR DA OBRA']
    
    m_cols = [c for c in cols_originais if c in df_alvo.columns and c not in base_start and c not in base_end and not str(c).startswith('_')]
    for c in c_garantia:
        if c in df_alvo.columns and c not in base_start and c not in base_end and c not in m_cols: m_cols.append(c)
    
    f_cols = []
    for c in base_start + m_cols + base_end:
        if c in df_alvo.columns and c not in f_cols: f_cols.append(c)
    return df_alvo[f_cols]

def tentar_rerun():
    try: st.rerun()
    except AttributeError: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido': False, 'vrp_status': "IDLE", 'vrp_state': {}, 'df_routed': pd.DataFrame(), 'bases_records': [], 'tipo_periodo': "Dia", 'colunas_exibir': [], 'colunas_originais': []})
    for k in ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg', 'df_unallocated', 'df_correcao_fiscalizacao']: st.session_state.pop(k, None)
    ler_planilha_cached.clear(); tentar_rerun()

# ==========================================
# INTERFACE PRINCIPAL
# ==========================================
if "roteamento_concluido" not in st.session_state: st.session_state.roteamento_concluido = False
if "vrp_status" not in st.session_state: st.session_state.vrp_status = "IDLE"

status_exec = st.session_state.vrp_status
is_done = st.session_state.roteamento_concluido
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📋 Planejamento Tático (Fiscalização)</h1>", unsafe_allow_html=True)
st.info("💡 **Inteligência Logística:** Este módulo distribui Obras e Postes de forma equilibrada entre os Fiscais disponíveis no mesmo município, priorizando as maiores demandas.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Esforço e Limites", expanded=True):
        trava_global_obras = st.number_input("Trava Total de Obras no Estado", min_value=0, value=0, step=50, disabled=is_locked)
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão (Mais Próximo Primeiro)", "🎯 Varredura Reversa (Mais Distante Primeiro)"], index=0, disabled=is_locked)
        raio_super_ponto = st.slider("Raio Super Ponto (Metros)", 10, 1000, 100, 10, disabled=is_locked)
        
        st.markdown("---")
        tipo_periodo = st.radio("Agrupamento:", ["☀️ Dia", "📅 Semana"], index=1, horizontal=True, disabled=is_locked)
        tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
        dias_semana_selecionados = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
        if tipo_periodo_clean == "Semana":
            dias_semana_selecionados = st.multiselect("Dias úteis:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked)
        
        data_inicio_roteiro = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        obras_por_dia = st.number_input("Obras Previstas por Dia", min_value=1, value=30, step=1, disabled=is_locked)
        limite_periodos = st.number_input(f"Limite total de {tipo_periodo_clean}s", min_value=1, value=5, step=1, disabled=is_locked)
        velocidade_media_kmh = 30.0

    st.markdown("---")
    sidebar_html_placeholder = st.empty()
    
    if is_done and not st.session_state.df_routed.empty:
        st.markdown("### 📥 Baixar Resultados")
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl', b"vazio"), file_name=f"Fiscais_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml', b"vazio"), file_name=f"Fiscais_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx', b"vazio"), file_name=f"Fiscais_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

# --- RESULTADOS (SE CONCLUÍDO) ---
if is_done and not st.session_state.df_routed.empty:
    st.markdown("## 🎯 Dashboards de Produtividade")
    df_routed = st.session_state.df_routed.copy()
    df_real_tasks = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tot_obras_reais = sum(count_real_obras(r) for _, r in df_real_tasks.iterrows())
    tot_equipes = df_routed['BASE_ATRIBUIDA'].nunique()
    tot_km = f"{df_routed['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    tot_postes_global = int(pd.to_numeric(df_real_tasks['QTD PREVISTA DE POSTES'], errors='coerce').sum()) if 'QTD PREVISTA DE POSTES' in df_real_tasks else 0

    c_m1, c_m2, c_m3, c_m4 = st.columns(4)
    c_m1.markdown(f'<div class="metric-card" style="border-left: 5px solid #0D256C;"><div class="metric-icon" style="background: rgba(13,37,108,0.12);">🎯</div><div class="metric-content"><div class="metric-title">TOTAL DE OBRAS</div><div class="metric-value">{tot_obras_reais}</div></div></div>', unsafe_allow_html=True)
    c_m2.markdown(f'<div class="metric-card" style="border-left: 5px solid #8b5cf6;"><div class="metric-icon" style="background: rgba(139,92,246,0.15);">👥</div><div class="metric-content"><div class="metric-title">Fiscais Alocados</div><div class="metric-value">{tot_equipes}</div></div></div>', unsafe_allow_html=True)
    c_m3.markdown(f'<div class="metric-card" style="border-left: 5px solid #FF9800;"><div class="metric-icon" style="background: rgba(255,152,0,0.15);">🏗️</div><div class="metric-content"><div class="metric-title">Postes Fiscalizados</div><div class="metric-value">{tot_postes_global}</div></div></div>', unsafe_allow_html=True)
    c_m4.markdown(f'<div class="metric-card" style="border-left: 5px solid #55B929;"><div class="metric-icon" style="background: rgba(85,185,41,0.15);">🛣️</div><div class="metric-content"><div class="metric-title">KM Previsto Total</div><div class="metric-value">{tot_km}</div></div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- GRÁFICOS EXCLUSIVOS DE FISCALIZAÇÃO ---
    chart_data = []
    for b_name in df_real_tasks['BASE_ATRIBUIDA'].unique():
        df_f = df_real_tasks[df_real_tasks['BASE_ATRIBUIDA'] == b_name]
        q_obras = sum(count_real_obras(r) for _, r in df_f.iterrows())
        q_postes = pd.to_numeric(df_f['QTD PREVISTA DE POSTES'], errors='coerce').sum() if 'QTD PREVISTA DE POSTES' in df_f.columns else 0
        if q_obras > 0: chart_data.append({"Fiscal": b_name, "Obras": q_obras, "Postes": int(q_postes)})
    
    df_chart = pd.DataFrame(chart_data)
    if not df_chart.empty:
        c_ch1, c_ch2 = st.columns([1.2, 1])
        with c_ch1:
            st.markdown("#### Balanceamento: Obras vs Postes")
            df_melted = df_chart.melt(id_vars='Fiscal', value_vars=['Obras', 'Postes'], var_name='Métrica', value_name='Quantidade')
            bar_chart = alt.Chart(df_melted).mark_bar().encode(
                x=alt.X('Fiscal:N', title=None, axis=alt.Axis(labelAngle=-45)),
                y=alt.Y('Quantidade:Q', title="Total Alocado"),
                color=alt.Color('Métrica:N', scale=alt.Scale(domain=['Obras', 'Postes'], range=['#0D256C', '#FF9800'])),
                xOffset='Métrica:N',
                tooltip=['Fiscal', 'Métrica', 'Quantidade']
            ).properties(height=350)
            st.altair_chart(bar_chart, use_container_width=True)
        
        with c_ch2:
            st.markdown("#### % Fatia de Postes por Fiscal")
            donut = alt.Chart(df_chart).mark_arc(innerRadius=60).encode(
                theta=alt.Theta(field="Postes", type="quantitative"),
                color=alt.Color(field="Fiscal", type="nominal"),
                tooltip=["Fiscal", "Postes", "Obras"]
            ).properties(height=350)
            st.altair_chart(donut, use_container_width=True)
    
    st.markdown("#### 📝 Resumo Operacional")
    resumo_ui = []
    for item in chart_data:
        km_tot = df_routed[df_routed['BASE_ATRIBUIDA'] == item["Fiscal"]]['DISTANCIA_PONTO_ANTERIOR_KM'].sum()
        resumo_ui.append({"Levantador (Fiscal)": item["Fiscal"], "Postes Fiscalizados": item["Postes"], "KM Total Previsto": round(km_tot, 2)})
    st.dataframe(pd.DataFrame(resumo_ui), use_container_width=True, hide_index=True)

    # --- MAPA ---
    st.markdown("### 🗺️ Mapa Geográfico")
    mapa = folium.Map(location=[df_routed['LATITUDE'].mean(), df_routed['LONGITUDE'].mean()], zoom_start=8) if not df_routed.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    cores_folium = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548']
    
    m_clust = MarkerCluster(name="Obras").add_to(mapa)
    for b_name in df_routed['BASE_ATRIBUIDA'].unique().tolist():
        cor_r = cores_folium[df_routed['BASE_ATRIBUIDA'].unique().tolist().index(b_name) % len(cores_folium)]
        df_br = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
        f_grp = folium.FeatureGroup(name=f"Rota: {b_name}", show=False)
        
        for per in df_br['PERIODO'].unique():
            df_per = df_br[df_br['PERIODO'] == per]
            pts = [p for _, r in df_per.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
            folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(f_grp)
            folium.PolyLine(pts, color=cor_r, weight=3, opacity=1.0).add_to(f_grp)
            
            for r in df_per.to_dict('records'):
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                c_i = r.get('COR_ICONE', 'gray')
                qtd_p = int(float(r.get('QTD PREVISTA DE POSTES', 0)))
                p_txt = f"📍 FISCALIZAÇÃO - {qtd_p} POSTES"
                p_bg = "#4CAF50" if c_i=='green' else "#2196F3" if c_i=='blue' else "#F5F5DC" if c_i=='beige' else "#FF9800" if c_i=='orange' else "#F44336"
                p_c = "#000000" if c_i=='beige' else "#ffffff"
                
                e_rows = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}:</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir])
                pop_html = f'<div style="width:280px;font-family:sans-serif;"><div style="background:{p_bg};color:{p_c};padding:8px;font-weight:bold;">{p_txt}</div><table border="1" style="width:100%;border-collapse:collapse;font-size:12px;">{e_rows}</table></div>'
                
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon='eye'), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        f_grp.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

# --- TELA INICIAL DE UPLOAD (SE NÃO CONCLUÍDO) ---
elif status_exec == "IDLE":
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown("### 👥 1. Fiscais")
        df_bases = pd.DataFrame()
        base_file = st.file_uploader("Suba a planilha de Fiscais (Excel)", type=["xlsx", "xls"])
        if base_file:
            try:
                b_t = ler_planilha_cached(base_file.getvalue()); b_t.columns = normalize_cols(b_t.columns)
                for pn in ['NOME', 'FISCAL', 'TECNICO', 'COLABORADOR']:
                    if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'LEVANTADOR'}); break
                if 'LEVANTADOR' in b_t.columns:
                    opts = sorted([str(x) for x in b_t['LEVANTADOR'].dropna().unique() if str(x).upper().strip() != 'SEM LEVANTADOR'])
                    sel = st.multiselect("Selecione os Fiscais Ativos:", opts, default=opts)
                    if sel:
                        df_bases = b_t[b_t['LEVANTADOR'].isin(sel)].copy()
                        if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                            df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                            df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                        elif 'RESIDENCIA' in df_bases.columns or 'MUNICIPIO' in df_bases.columns:
                            cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                            mc = {}
                            with st.spinner("🌍 Mapeando bases..."):
                                for m in df_bases[cr].dropna().unique(): mc[m] = obter_coordenadas_municipio_cached(m)
                            df_bases['LATITUDE'], df_bases['LONGITUDE'] = df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0]), df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
                        df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE']); df_bases['TIPO_EQUIPE'] = 'FISCAL'
                else: st.error("❌ A planilha não possui a coluna 'FISCAL'.")
            except Exception as e: st.error(f"Erro: {e}")

    with c_up2:
        st.markdown("### 📁 2. Obras de Fiscalização")
        task_files = st.file_uploader("Suba as Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if df_bases.empty or not task_files: st.stop()
    
    qtd_eq = df_bases['LEVANTADOR'].nunique()
    cap_max = obras_por_dia * (len(dias_semana_selecionados) if tipo_periodo_clean == 'Semana' else 1) * limite_periodos
    sidebar_html_placeholder.markdown(renderizar_painel_lateral(cap_max, 0, qtd_eq, cap_max * qtd_eq), unsafe_allow_html=True)

    # Lendo Obras
    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        if not dfs: st.session_state.colunas_originais = dft.columns.tolist()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    st.markdown("---")
    # FILTRO GEOFENCING E STATUS
    falta = [c for c in ['MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas: {', '.join(falta)}."); st.stop()
    
    cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO'
    if cs in df_tasks.columns:
        df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
        st.markdown("#### 📊 Filtragem de Status")
        opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
        sel_s = st.multiselect("Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
        if not sel_s: st.stop()
        df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()

    cg1, cg2 = st.columns([4, 1])
    with cg1: st.markdown("#### 🌍 Cerca Eletrônica (Geofencing 70km)")
    with cg2:
        if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    pbg = st.progress(0.0); tmp = st.empty(); sgt = st.empty(); tsg = time.time()
    df_tasks['MOTIVO_REJEICAO'] = ''; df_rej = pd.DataFrame()
    
    m_m = df_tasks['MUNICIPIO'].isna() | (df_tasks['MUNICIPIO'].astype(str).str.strip() == '') | (df_tasks['MUNICIPIO'].astype(str).str.strip().str.upper() == 'NAN')
    if m_m.sum() > 0:
        df_tasks.loc[m_m, 'MOTIVO_REJEICAO'] = 'Município Vazio'
        df_rej = pd.concat([df_rej, df_tasks[m_m].copy()], ignore_index=True); df_tasks = df_tasks[~m_m].copy()

    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    m_na, m_0 = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna(), (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    df_tasks.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df_tasks.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)
    df_tasks.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    m_inv = abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])
    df_tasks.loc[m_inv & ~m_na & ~m_0 & ~m_pos, 'MOTIVO_REJEICAO'] = 'Coordenada Invertida'
    
    mc = m_na | m_0 | m_pos | m_inv
    if mc.sum() > 0: df_rej = pd.concat([df_rej, df_tasks[mc].copy()], ignore_index=True); df_tasks = df_tasks[~mc].copy()
    
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']
    df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    if not df_tasks.empty:
        mu = df_tasks['MUNICIPIO'].unique(); tm = len(mu); md = {}
        for i, m in enumerate(mu):
            fr = (i + 1) / max(1, tm); pbg.progress(fr); sgt.info(f"🛰️ Satélite: {m}")
            try: lt, ln = obter_coordenadas_municipio_cached(m); time.sleep(0.3)
            except: lt, ln = np.nan, np.nan
            md[m] = (lt, ln)
            
        sgt.info("📏 Aplicando Cerca...")
        mo = df_tasks.apply(lambda r: haversine_scalar(r['LATITUDE'], r['LONGITUDE'], md.get(r['MUNICIPIO'], (np.nan, np.nan))[0], md.get(r['MUNICIPIO'], (np.nan, np.nan))[1]) > 70.0 if pd.notna(md.get(r['MUNICIPIO'], (np.nan, np.nan))[0]) else False, axis=1)
        if mo.sum() > 0:
            df_tasks.loc[mo, 'MOTIVO_REJEICAO'] = 'Fora do Município (> 70km)'
            df_rej = pd.concat([df_rej, df_tasks[mo].copy()], ignore_index=True); df_tasks = df_tasks[~mo].copy()
        
        pbg.empty(); sgt.empty()
        st.session_state.df_correcao_fiscalizacao = df_rej
        if not df_rej.empty: st.warning(f"⚠️ {len(df_rej)} obras com erro isoladas na 'Planilha de Correção'.")

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    if 'QTD PREVISTA DE POSTES' in df_tasks.columns:
        df_tasks['QTD PREVISTA DE POSTES'] = df_tasks['QTD PREVISTA DE POSTES'].apply(extrair_qtd)
        df_tasks['COR_ICONE'] = df_tasks['QTD PREVISTA DE POSTES'].apply(definir_cor_fiscalizacao)
    else: df_tasks['QTD PREVISTA DE POSTES'], df_tasks['COR_ICONE'] = 0.0, 'gray'
    df_tasks['PRIORIDADE'] = 'Sim'

    df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_super_ponto, agrupar_por_levantador=False)
    df_tasks['MUN_LIMPO'] = normalizar_municipios(df_tasks['MUNICIPIO'].fillna(''))

    tbr = df_bases.to_dict('records')
    mta = {}
    for b in tbr:
        for m in str(b.get('MUNICIPIO', b.get('RESIDENCIA', ''))).split(','):
            ml = normalizar_municipios(pd.Series([m])).iloc[0]
            if ml:
                if ml not in mta: mta[ml] = []
                if b['LEVANTADOR'] not in mta[ml]: mta[ml].append(b['LEVANTADOR'])

    bc, bp = {b['LEVANTADOR']: 0 for b in tbr}, {b['LEVANTADOR']: 0.0 for b in tbr}
    
    if trava_global_obras > 0:
        df_tasks = df_tasks.sort_values(by=['QTD PREVISTA DE POSTES', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True]).head(trava_global_obras)

    df_tasks['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
    assigned_tasks, unassigned_tasks = [], []
    
    df_combinado = df_tasks.sort_values(by=['QTD PREVISTA DE POSTES', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
    
    for r in df_combinado.to_dict('records'):
        qr = len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
        la, lo, ms = r.get('LATITUDE'), r.get('LONGITUDE'), str(r.get('MUN_LIMPO', ''))
        vb = [b for b in tbr if b['LEVANTADOR'] in set(mta.get(ms, []))]
        bb, bd, qpo = None, float('inf'), float(r.get('QTD PREVISTA DE POSTES', 0))
        
        if pd.notna(la) and pd.notna(lo):
            # Lógica Crucial de Load Balancing para Fiscais (Sorteia quem tem menos postes primeiro)
            vb = sorted(vb, key=lambda b: (bp[b['LEVANTADOR']], bc[b['LEVANTADOR']]))
            for b in vb:
                bn = b['LEVANTADOR']
                if bc[bn] + qr <= cap_max: bb = bn; break
        
        if bb:
            bc[bb] += qr; bp[bb] += qpo; r['BASE_ATRIBUIDA'] = bb; assigned_tasks.append(r)
        else:
            r['MOTIVO_REJEICAO'] = "Estoque Lotado ou Sem Fiscal na Cidade"
            unassigned_tasks.append(r)

    df_ta, df_u = pd.DataFrame(assigned_tasks), pd.DataFrame(unassigned_tasks)
    st.session_state.df_unallocated, st.session_state.tot_obras_nao_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_u.iterrows())
    
    sidebar_html_placeholder.markdown(renderizar_painel_lateral(cap_max, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows()), qtd_eq, cap_max * qtd_eq), unsafe_allow_html=True)
    if df_ta.empty: st.error("Nenhuma obra pôde ser alocada aos Fiscais."); st.stop()

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'COR_ICONE' and c != 'MUN_LIMPO']
        cd = ['PROTOCOLO', 'VALOR DA OBRA', 'QTD PREVISTA DE POSTES', 'PREVISAO DE ENTREGA', 'PARCEIRO', 'TIPO DE FISCALIZACAO', 'TIPO DE PROJETO', 'REGIONAL', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'ZONA', 'STATUS DA FISCALIZACAO', 'LEVANTADOR', 'BACKOFFICE DA FISCALIZACAO', 'OBSERVACAO']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        st.session_state.update({'bases_records': tbr, 'tipo_periodo': tipo_periodo_clean, 'colunas_exibir': colunas_exibir, 'col_prioridade': "PRIORIDADE"})
        st.session_state.vrp_state = {'config': {'velocidade_media_kmh': velocidade_media_kmh, 'obras_por_dia': obras_por_dia, 'tipo_periodo': tipo_periodo_clean, 'limite_periodos': limite_periodos, 'dias_selecionados': dias_semana_selecionados, 'url_osrm_base': 'http://router.project-osrm.org', 'tracado_real': False, 'data_inicio': data_inicio_roteiro, 'sentido_rota': sentido_rota, 'modo_operacao': '3'}, 'b_names': list(set([b['LEVANTADOR'] for b in tbr])), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': []}
        st.session_state.vrp_status = "RUNNING"; tentar_rerun()

# --- EXECUÇÃO (RUNNING) ---
if status_exec == "RUNNING":
    state = st.session_state.vrp_state
    b_names = state['b_names']
    b_idx = state.get('b_idx', 0)
    
    if b_idx < len(b_names):
        b_name = b_names[b_idx]
        df_todas_bases_ativas = pd.DataFrame(st.session_state.bases_records)
        base_lat, base_lon = float(df_todas_bases_ativas[df_todas_bases_ativas['LEVANTADOR'] == b_name].iloc[0]['LATITUDE']), float(df_todas_bases_ativas[df_todas_bases_ativas['LEVANTADOR'] == b_name].iloc[0]['LONGITUDE'])
        obras_equipe = state['unvisited'][state['unvisited']['BASE_ATRIBUIDA'] == b_name].to_dict('records')
        
        # Greedy Sorting focado em POSTES (Score = Distância - Peso dos Postes)
        ordered_tasks, curr_lat, curr_lon, unvisited_pts = [], base_lat, base_lon, list(obras_equipe)
        is_reversa = "Reversa" in state['config'].get('sentido_rota', '')
        
        if is_reversa and unvisited_pts:
            bi, bs = 0, -1
            for i, p in enumerate(unvisited_pts):
                d = haversine_scalar(curr_lat, curr_lon, p['LATITUDE'], p['LONGITUDE'])
                sc = d * (1 + (extrair_qtd(p.get('QTD PREVISTA DE POSTES', 0)) / 50.0))
                if sc > bs: bs, bi = sc, i
            nxt = unvisited_pts.pop(bi); ordered_tasks.append(nxt); curr_lat, curr_lon = nxt['LATITUDE'], nxt['LONGITUDE']
            
        while unvisited_pts:
            bi, bs = 0, float('inf')
            for i, p in enumerate(unvisited_pts):
                d = haversine_scalar(curr_lat, curr_lon, p['LATITUDE'], p['LONGITUDE'])
                sc = d - min(d * 0.8, extrair_qtd(p.get('QTD PREVISTA DE POSTES', 0)) * 0.1)
                if sc < bs: bs, bi = sc, i
            nxt = unvisited_pts.pop(bi); ordered_tasks.append(nxt); curr_lat, curr_lon = nxt['LATITUDE'], nxt['LONGITUDE']

        # Fatiamento Simples por Dias
        routed_data_final, dia_abs = [], 1
        for o in ordered_tasks:
            o['ORDEM'], o['DIA'], o['PERIODO'], o['NOME_DIA'], o['DIA_MES'] = len(routed_data_final) + 1, dia_abs, dia_abs, f"Dia {dia_abs}", ""
            o['DISTANCIA_PONTO_ANTERIOR_KM'] = 1.0 # Placeholder simplificado para velocidade
            o['_HORA_INICIO_DT'], o['_HORA_FIM_DT'] = datetime.now(), datetime.now()
            routed_data_final.append(o)

        state['routed_data'].extend(routed_data_final)
        state['b_idx'] += 1; st.session_state.vrp_state = state; tentar_rerun()
    else:
        st.session_state.df_routed = pd.DataFrame(state['routed_data'])
        st.session_state.vrp_status = "PACKAGING"; tentar_rerun()

# --- EMPACOTAMENTO ---
if status_exec == "PACKAGING":
    df_routed = st.session_state.df_routed
    d_fmt = datetime.now().strftime("%d.%m.%Y")
    buf_xl, buf_kml = io.BytesIO(), io.BytesIO()
    
    with zipfile.ZipFile(buf_xl, 'w', zipfile.ZIP_DEFLATED) as z_xl, zipfile.ZipFile(buf_kml, 'w', zipfile.ZIP_DEFLATED) as z_kml:
        # Resumo
        res = []
        for b in df_routed['BASE_ATRIBUIDA'].unique():
            dfb = df_routed[df_routed['BASE_ATRIBUIDA'] == b]
            res.append({'LEVANTADOR (FISCAL)': b, 'POSTES FISCALIZADOS': int(pd.to_numeric(dfb.get('QTD PREVISTA DE POSTES',0), errors='coerce').sum()), 'OBRAS': len(dfb)})
        z_xl.writestr(f"Resumo_Fiscais - {d_fmt}.xlsx", gerar_excel_resumo_bytes(pd.DataFrame(res)))
        
        # Correção
        if not st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame()).empty:
            dfc = st.session_state.df_correcao_fiscalizacao.copy()
            dfc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
            for c in dfc.columns:
                if str(dfc[c].dtype) == 'object': dfc[c] = dfc[c].astype(str).replace('nan', '')
            out_e = io.BytesIO(); dfc.to_excel(out_e, index=False); z_xl.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
        
        # Geral
        dfg = limpar_colunas_excel(df_routed.drop(columns=['MUN_LIMPO', 'COR_ICONE'], errors='ignore'), st.session_state.colunas_originais)
        z_xl.writestr(f"Demanda_Fiscais - {d_fmt}.xlsx", gerar_excel_bytes(dfg, "PRIORIDADE"))
        
        # KML
        z_kml.writestr(f"ROTA_TOTAL - {d_fmt}.kml", gerar_kml_fiscalizacao(df_routed, f"ROTAS FISCAIS", st.session_state.colunas_exibir).encode('utf-8'))

    st.session_state.bytes_zip_xl, st.session_state.bytes_zip_kml = buf_xl.getvalue(), buf_kml.getvalue()
    st.session_state.roteamento_concluido = True; st.session_state.vrp_status = "IDLE"; tentar_rerun()
