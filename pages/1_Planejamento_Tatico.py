import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
import gc
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Módulos Separados
from modules.data_processing import ler_planilha_cached, formata_campo_html, formatar_moeda, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas
from modules.export_tatica import injetar_logo, identificar_icone_folium, gerar_excel_tatica, limpar_colunas_tatica, gerar_kml_tatica, gerar_gpx_simples, gerar_excel_resumo_tatica

st.set_page_config(page_title="Roteirizador Tático", page_icon="🗺️", layout="wide")
injetar_logo()

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} KM"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Cota p/ Equipe:</b> <span style="color: #d9534f; font-weight: bold;">{limite_por_equipe}</span> obras</p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Capacidade Total:</b> <span style="color: #55B929; font-weight: bold;">{total_capacidade}</span> obras</p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras_prontas}</span></p>
    </div>
    """

def render_metric_card(title, value, icon, border_color, bg_color):
    return f"""
    <div style="background-color: #ffffff; border-left: 5px solid {border_color}; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); display: flex; align-items: center; margin-bottom: 10px;">
        <div style="background-color: {bg_color}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; font-size: 20px; margin-right: 15px;">
            {icon}
        </div>
        <div>
            <p style="margin: 0; font-size: 11px; color: #666; text-transform: uppercase; font-weight: bold;">{title}</p>
            <p style="margin: 0; font-size: 22px; color: #333; font-weight: bold;">{value}</p>
        </div>
    </div>
    """

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido': False, 'vrp_status': "IDLE", 'vrp_state': {}, 'df_routed': pd.DataFrame(), 'bases_records': [], 'colunas_exibir': [], 'colunas_originais_tat': []})
    for k in ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg', 'df_unallocated', 'df_correcao_tatica']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido" not in st.session_state: st.session_state.roteamento_concluido = False
if "vrp_status" not in st.session_state: st.session_state.vrp_status = "IDLE"

status_exec = st.session_state.vrp_status
is_done = st.session_state.roteamento_concluido
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>🗺️ Planejamento Tático (Operação)</h1>", unsafe_allow_html=True)
st.info("💡 Distribui as obras equitativamente entre as equipes, criando rotas circulares diárias e priorizando notas urgentes.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Capacidade e Prazos", expanded=True):
        obras_dia = st.number_input("Cota Diária por Equipe:", min_value=1, value=6, disabled=is_locked)
        tpc = st.radio("Visão de Trabalho:", ["Dia", "Semana"], index=1, disabled=is_locked)
        limite_per = st.number_input(f"Qtd de {tpc}s de Rota:", min_value=1, value=1, disabled=is_locked)
        dias_sel = st.multiselect("Dias Úteis na Semana:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked)
        data_ini = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        
        st.markdown("---")
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)

    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl', b"vazio"), file_name=f"Rotas_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml', b"vazio"), file_name=f"Rotas_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx', b"vazio"), file_name=f"Rotas_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    df_c = st.session_state.get('df_correcao_tatica', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica:</b> Estas obras apresentaram coordenadas em branco, zeradas, invertidas ou caíram fora da <b>Cerca Eletrônica de 70km</b> do município de origem.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = len(dfr_t)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    
    # NOVO: Conta a quantidade de Super Pontos (Onde _ORIGINAL_ROWS é uma lista com mais de 1 item)
    tsp = sum(1 for _, r in dfr_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Planejadas", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(tsp), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    # NOVO: Mensagem de alocação movida para cá e sem abas inferiores
    if not st.session_state.get('df_unallocated', pd.DataFrame()).empty:
        st.warning(f"⚠️ {len(st.session_state.df_unallocated)} obras não couberam na cota de dias/equipes ou ficaram distantes demais (Verifique o arquivo ZIP gerado).")
    else:
        st.success("✅ 100% das obras foram alocadas com sucesso.")

    st.markdown("### 🗺️ Mapa Operacional")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        for pe in db['PERIODO'].unique():
            dp = db[db['PERIODO'] == pe]
            pts = [p for _, r in dp.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
            folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
            folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
            
            for r in dp.to_dict('records'):
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                c_i = 'red' if str(r.get('PRIORIDADE')) == 'Sim' else 'blue'
                ic = identificar_icone_folium(r, dfr.columns)
                er = "".join([f"<tr><td><b>{html.escape(c)}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
                pop_html = f'<div style="width:250px;"><b>Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

elif status_exec == "IDLE":
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown("### 👥 1. Bases de Equipes")
        df_bases = pd.DataFrame()
        bf = st.file_uploader("Suba a planilha de Equipes (Excel)", type=["xlsx", "xls"])
        if bf:
            b_t = ler_planilha_cached(bf.getvalue()); b_t.columns = normalize_cols(b_t.columns)
            b_t = b_t.loc[:, ~b_t.columns.duplicated()].copy()
            
            # NOVO: Priorizando as colunas de "NOME" do Levantador antes de bater com "EQUIPE" numérico
            for pn in ['LEVANTADOR', 'FISCAL', 'NOME_FISCAL', 'NOME', 'TECNICO', 'COLABORADOR', 'EQUIPE']:
                if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'BASE_NOME'}); break
                
            if 'BASE_NOME' in b_t.columns:
                
                b_t['BASE_NOME'] = b_t['BASE_NOME'].astype(str).str.split(r'\s*\|\s*')
                b_t = b_t.explode('BASE_NOME').reset_index(drop=True)
                b_t['BASE_NOME'] = b_t['BASE_NOME'].str.strip().str.upper()
                
                opts = sorted([str(x) for x in b_t['BASE_NOME'].dropna().unique() if str(x) not in ['SEM EQUIPE', 'NAN', 'NONE', '']])
                sel = st.multiselect("Selecione as Equipes Ativas:", opts, default=opts)
                if sel:
                    df_bases = b_t[b_t['BASE_NOME'].isin(sel)].copy()
                    if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                        df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                        df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                    elif 'RESIDENCIA' in df_bases.columns or 'MUNICIPIO' in df_bases.columns:
                        cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                        mc = {}
                        with st.spinner("🌍 Mapeando bases via IBGE..."):
                            for m in df_bases[cr].dropna().unique(): mc[m] = obter_coordenadas_municipio_cached(m)
                        df_bases['LATITUDE'], df_bases['LONGITUDE'] = df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0]), df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
                    df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE'])
            else: st.error("❌ A planilha não possui coluna de nome da Equipe/Levantador.")

        st.markdown("##### 📍 Regra de Atribuição")
        ta = st.radio("Como amarrar as notas aos técnicos?", ["Por Proximidade Espacial", "Por Município Base"], index=0, label_visibility="collapsed")
        st.caption("A proximidade empurra as obras para a equipe mais próxima no raio geral. Por Município isola a equipe.")

    with c_up2:
        st.markdown("### 📁 2. Demandas (Obras)")
        task_files = st.file_uploader("Suba as planilhas de Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if df_bases.empty or not task_files: st.stop()
    
    qtd_eq = df_bases['BASE_NOME'].nunique()
    cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per
    sb_html.markdown(render_sidebar_card(cm, 0, qtd_eq, cm * qtd_eq), unsafe_allow_html=True)

    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        if not dfs: st.session_state.colunas_originais_tat = dft.columns.tolist()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    st.markdown("---")
    falta = [c for c in ['MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas: {', '.join(falta)}."); st.stop()
    
    c1_f, c2_f = st.columns([1, 1])
    with c1_f:
        cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO'
        if cs in df_tasks.columns:
            df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
            opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
            sel_s = st.multiselect("1. Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
            if not sel_s: st.stop()
            df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()
            
    with c2_f:
        if 'TIPO NOTA' in df_tasks.columns:
            df_tasks['TIPO NOTA'] = df_tasks['TIPO NOTA'].astype(str).str.strip().str.upper()
            opts_n = sorted([str(x) for x in df_tasks['TIPO NOTA'].unique() if str(x) != 'NAN'])
            sel_p = st.multiselect("🚨 2. Obras de Alta Prioridade:", options=opts_n, default=[n for n in opts_n if n in ['ASC', 'CCF', 'DIF', 'MGD', 'MTP', 'SID']])
            df_tasks['PRIORIDADE'] = df_tasks['TIPO NOTA'].apply(lambda x: 'Sim' if str(x) in sel_p else 'Não')
        else: df_tasks['PRIORIDADE'] = 'Não'

    cg1, cg2 = st.columns([4, 1])
    with cg1: st.markdown("#### 🌍 Cerca Eletrônica Municipal (Anti-Fuga)")
    with cg2:
        if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    pbg = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
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
    
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']; df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    if not df_tasks.empty:
        mu = df_tasks['MUNICIPIO'].unique(); tm = len(mu); md = {}
        st_run_geo = time.time()
        
        def render_t_geo(curr, total):
            e = time.time() - st_run_geo; f = curr / max(1, total)
            rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
            es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
            html_timer = f"""
            <div style="display:flex; gap:15px; margin-top: 10px; margin-bottom: 10px;">
                <div style="flex:1; padding:10px; border-radius:8px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;">
                    <div style="font-size:0.8rem; color:#6c757d; font-weight:bold; margin-bottom:2px;">⏱️ Decorrido</div>
                    <div style="font-size:1.5rem; font-weight:bold; color:#0D256C;">{es}</div>
                </div>
                <div style="flex:1; padding:10px; border-radius:8px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;">
                    <div style="font-size:0.8rem; color:#2e7d32; font-weight:bold; margin-bottom:2px;">🎯 Restante</div>
                    <div style="font-size:1.5rem; font-weight:bold; color:#1b5e20;">{rs}</div>
                </div>
            </div>
            """
            tmp.markdown(html_timer, unsafe_allow_html=True)

        for i, m in enumerate(mu):
            pbg.progress((i + 1) / max(1, tm)); sgt.info(f"🛰️ Satélite IBGE: {m}")
            render_t_geo(i + 1, tm)
            try: lt, ln = obter_coordenadas_municipio_cached(m); time.sleep(0.1)
            except: lt, ln = np.nan, np.nan
            md[m] = (lt, ln)
            
        sgt.info("📏 Aplicando régua na Cerca...")
        mo = df_tasks.apply(lambda r: haversine_scalar(r['LATITUDE'], r['LONGITUDE'], md.get(r['MUNICIPIO'], (np.nan, np.nan))[0], md.get(r['MUNICIPIO'], (np.nan, np.nan))[1]) > 70.0 if pd.notna(md.get(r['MUNICIPIO'], (np.nan, np.nan))[0]) else False, axis=1)
        if mo.sum() > 0:
            df_tasks.loc[mo, 'MOTIVO_REJEICAO'] = 'Fora do Município (> 70km)'
            df_rej = pd.concat([df_rej, df_tasks[mo].copy()], ignore_index=True); df_tasks = df_tasks[~mo].copy()
        
        pbg.empty(); tmp.empty(); sgt.empty()
        st.session_state.df_correcao_tatica = df_rej

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_sp, agrupar_por_levantador=False)

    tbr = df_bases.to_dict('records')
    fiscal_anchors = {b['BASE_NOME']: (float(b.get('LATITUDE',0)), float(b.get('LONGITUDE',0))) for b in tbr}
    assigned_tasks, unassigned_tasks = [], []
    
    # ------------------ ISOLAMENTO DE FISCAIS AQUI ------------------
    # Antes, a trava_global cortava os dados antes de fazer o sorteio de equipe, 
    # o que poderia fazer a linha falhar se 'df_tasks' ficasse muito curto
    
    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])

    for r in df_tasks.to_dict('records'):
        la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
        ms = normalizar_municipios(pd.Series([str(r.get('MUNICIPIO', ''))])).iloc[0]
        
        if "Município" in ta: vb = [b for b in tbr if ms in str(b.get('MUNICIPIO', b.get('RESIDENCIA', ''))).upper()]
        else: vb = tbr
            
        best_f, best_d = None, float('inf')
        
        if pd.notna(la) and pd.notna(lo) and vb:
            for b in vb:
                f_name = b['BASE_NOME']
                d = haversine_scalar(la, lo, fiscal_anchors[f_name][0], fiscal_anchors[f_name][1])
                if d < best_d:
                    best_d = d; best_f = f_name
                        
        if best_f:
            r['BASE_ATRIBUIDA'], r['MUN_LIMPO'] = best_f, ms
            assigned_tasks.append(r)
            fiscal_anchors[best_f] = (la, lo)
        else:
            r['MOTIVO_REJEICAO'], r['BASE_ATRIBUIDA'] = "Fora de Área (Sem Fiscal)", "NÃO ALOCADO"
            unassigned_tasks.append(r)

    df_ta, df_u = pd.DataFrame(assigned_tasks), pd.DataFrame(unassigned_tasks)
    
    st.session_state.df_unallocated, total_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    
    sb_html.markdown(render_sidebar_card(cm, total_alocadas, qtd_eq, cm * qtd_eq), unsafe_allow_html=True)
    if df_ta.empty: st.error("Nenhuma obra pôde ser alocada aos Fiscais."); st.stop()

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = ['PROTOCOLO', 'VALOR DA OBRA', 'QTD PREVISTA DE POSTES', 'PREVISAO DE ENTREGA', 'PARCEIRO', 'TIPO DE FISCALIZACAO', 'TIPO DE PROJETO', 'REGIONAL', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'ZONA', 'STATUS DA FISCALIZACAO', 'LEVANTADOR', 'BACKOFFICE DA FISCALIZACAO', 'OBSERVACAO']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        st.session_state.update({'bases_records': tbr, 'colunas_exibir': colunas_exibir})
        # Removemos vel_kmh e tmp_obra da UI, então forçamos aqui 30 e 45 respectivamente
        st.session_state.vrp_state = {'config': {'velocidade_media_kmh': 30.0, 'obras_por_dia': obras_dia, 'tipo_periodo': tpc, 'limite_periodos': limite_per, 'dias_selecionados': dias_sel, 'url_osrm_base': url_osrm, 'tracado_real': usa_osrm, 'data_inicio': data_ini, 'tempo_medio_obra': 45.0 / 60.0, 'sentido_rota': sentido_rota}, 'b_names': list(set([b['BASE_NOME'] for b in tbr])), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': [], 'current_geoms': []}
        st.session_state.vrp_status = "RUNNING"; tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor VRP Tático")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run', time.time())
    if 'start_time_run' not in st.session_state: st.session_state.start_time_run = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            br = pd.DataFrame(st.session_state.bases_records)
            br = br[br['BASE_NOME'] == bn].iloc[0]
            if pd.isna(br.get('LATITUDE')): st_v['b_idx'] += 1; st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
            bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            
            # IMPLEMENTAÇÃO: Sentido do Roteamento (Varredura Reversa ou Lógica Padrão)
            if "Varredura Reversa" in cfg.get('sentido_rota', "Lógica Padrão"):
                # Varredura Reversa: Inicia do ponto mais DISTANTE e vem voltando pra base
                ot = []
                if oe:
                    # Encontra a obra mais longe da base
                    max_idx = max(range(len(oe)), key=lambda i: haversine_scalar(bl, bL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                    p_longe = oe.pop(max_idx)
                    ot.append(p_longe)
                    
                    # A partir dela, vai pegando a mais próxima para ir construindo o caminho de volta
                    cl, cL = float(p_longe['LATITUDE']), float(p_longe['LONGITUDE'])
                    while oe:
                        closest_idx = min(range(len(oe)), key=lambda i: haversine_scalar(cl, cL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                        nx = oe.pop(closest_idx)
                        ot.append(nx)
                        cl, cL = float(nx['LATITUDE']), float(nx['LONGITUDE'])
            else:
                ot = resolver_tsp_ortools(oe, bl, bL, cfg['url_osrm_base']) if oe else []
                if not ot: ot = oe
            
            rf, da, sa, dds = [], 1, 1, 1
            dtb = datetime.combine(cfg['data_inicio'], datetime.min.time()).replace(hour=8, minute=0)
            def gi(da):
                c, d_ok = dtb, [0,1,2,3,4,5,6] if not cfg['dias_selecionados'] else [{"Segunda":0,"Terça":1,"Quarta":2,"Quinta":3,"Sexta":4,"Sábado":5,"Domingo":6}[d] for d in cfg['dias_selecionados']]
                while c.weekday() not in d_ok: c += pd.Timedelta(days=1)
                ct = 1
                while ct < da:
                    c += pd.Timedelta(days=1)
                    if c.weekday() in d_ok: ct += 1
                return {'l': bl, 'L': bL, 't': c, 'd': c, 'oh': 0, 'lu': False}
            es = gi(da)

            for o in ot:
                if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                    st.session_state.df_unallocated = pd.concat([st.session_state.get('df_unallocated', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue

                qr = len(o.get('_ORIGINAL_ROWS', [1])) if isinstance(o.get('_ORIGINAL_ROWS'), list) else 1
                vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                if vkr < 0.05 and es['oh'] > 0: vm, em = 0.0, 30.0
                else: vm, em = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60, cfg['tempo_medio_obra']*60
                
                cp = es['t'] + pd.Timedelta(minutes=vm)
                if cp.hour >= 12 and not es['lu']:
                    ls = max(es['t'], es['d'].replace(hour=12)); le = ls + pd.Timedelta(hours=1)
                    rf.append({'o': None, 'il': True, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': es['l'], 'Lt': es['L'], 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': ls, 'hf': le, 'vm': 0.0, 'dk': 0.0})
                    es['t'], es['lu'] = le, True; cp = es['t'] + pd.Timedelta(minutes=vm)
                fp = cp + pd.Timedelta(minutes=em)
                
                if es['oh'] > 0 and (es['oh'] + qr > cfg['obras_por_dia']):
                    dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                    rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
                    da += 1
                    if cfg['tipo_periodo'] == "Semana":
                        dds += 1
                        if dds > len(cfg['dias_selecionados']): sa += 1; dds = 1
                    es = gi(da)
                    if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                        st.session_state.df_unallocated = pd.concat([st.session_state.get('df_unallocated', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue
                    vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE']); vk = vkr * 1.3
                    if vkr < 0.05 and es['oh'] > 0: vm, em = 0.0, 30.0
                    else: vm, em = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60, cfg['tempo_medio_obra']*60
                    cp = es['t'] + pd.Timedelta(minutes=vm); fp = cp + pd.Timedelta(minutes=em)
                
                rf.append({'o': o, 'il': False, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': cp, 'hf': fp, 'vm': vm, 'dk': vk})
                es['l'], es['L'], es['t'], es['oh'] = o['LATITUDE'], o['LONGITUDE'], fp, es['oh'] + qr
                
            if es['oh'] > 0 and not ((cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos'])):
                dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
            
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (30 if cfg['tracado_real'] else len(rf)), len(rf))
            for i in range(oi, ei):
                it = rf[i]
                if not cfg['tracado_real']: gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
                else:
                    if i%5==0: sgt.info(f"🛣️ Traçando arruamento **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf))
                    time.sleep(0.15)
                    try: gd.append(obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh']))
                    except: gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
            
            bl, bL = float(pd.DataFrame(st.session_state.bases_records)[pd.DataFrame(st.session_state.bases_records)['BASE_NOME']==bn].iloc[0]['LATITUDE']), float(pd.DataFrame(st.session_state.bases_records)[pd.DataFrame(st.session_state.bases_records)['BASE_NOME']==bn].iloc[0]['LONGITUDE'])
            rdf, og, dp = [], 1, ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            for it, (g, ds) in zip(rf, gd):
                pv = it['s'] if cfg['tipo_periodo']=="Semana" else it['d']
                dn = dp[datetime.strptime(it['dm'], '%d/%m/%Y').weekday()] if cfg['tipo_periodo']=="Semana" else f"Dia {it['d']}"
                if it['il']: rdf.append({'PROTOCOLO': 'PAUSA_ALMOCO', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': 0.0, 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']})
                elif it['ir']: rdf.append({'PROTOCOLO': 'RETORNO_BASE', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': round(it['dk'], 2), 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']})
                else:
                    ob = it['o']; ob['ORDEM'], ob['NOME_DIA'], ob['DIA_MES'], ob['SEMANA'], ob['DIA'], ob['PERIODO'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, dn, it['dm'], it['s'], it['d'], pv, round(it['dk'], 2)
                    ob['ROTA_GEOMETRIA'] = [[it['Lt'], it['lt']], [it['Lt'], it['lt']]] if it['la']==bl and it['La']==bL else g
                    ob['HORA_INICIO'], ob['HORA_FIM'], ob['_HORA_INICIO_DT'], ob['_HORA_FIM_DT'] = it['hi'].strftime('%H:%M'), it['hf'].strftime('%H:%M'), it['hi'], it['hf']
                    rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        st.session_state.df_routed = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento Tático")
    df_routed, d_fmt = st.session_state.df_routed, datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx = io.BytesIO(), io.BytesIO(), io.BytesIO()
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg:
            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                res.append({'Equipe': b, 'Obras Roteirizadas': sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in db.iterrows()), 'Super Pontos': qs, 'Prioridades Atendidas': len(db[db['PRIORIDADE']=='Sim']), 'KM Total Previsto': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)})
            zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_tatica(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_tatica', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
            
            linhas_gerais = []
            for _, r in df_routed.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                if isinstance(r.get('_ORIGINAL_ROWS'), list):
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = r.copy()
                        for k, v in orig.items(): 
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'ROTA_GEOMETRIA', 'PERIODO']: nr[k] = v
                        linhas_gerais.append(nr)
                else: linhas_gerais.append(r)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            dfg = limpar_colunas_tatica(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), st.session_state.colunas_originais_tat)
            dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
            for cc in dfg.columns:
                if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Tatica - {d_fmt}.xlsx", gerar_excel_tatica(dfg, st.session_state.colunas_originais_tat))
            
            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            ks = gerar_kml_tatica(dfk_total, "ROTA_TOTAL", st.session_state.colunas_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), tpc, formatar_valor_coluna)
            zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks.encode('utf-8'))
            zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

        st.session_state.bytes_zip_xl, st.session_state.bytes_zip_kml, st.session_state.bytes_zip_gpx = bu_xl.getvalue(), bu_kml.getvalue(), bu_gpx.getvalue()
        st.session_state.roteamento_concluido = True; st.session_state.vrp_status = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status = "IDLE"
