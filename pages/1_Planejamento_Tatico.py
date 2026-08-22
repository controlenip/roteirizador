import streamlit as st, pandas as pd, numpy as np, folium, io, zipfile, html, re, time, os, gc
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime
from modules.data_processing import ler_planilha_cached, formatar_moeda, formata_campo_html, normalize_cols, normalizar_municipios, atualizar_status_via_df
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas
from modules.export_utils import identificar_icone_folium, renderizar_painel_lateral, gerar_excel_bytes, gerar_excel_resumo_bytes, gerar_kml_agrupado, gerar_gpx_simples

st.set_page_config(page_title="Modo Tático", page_icon="🎯", layout="wide")
STATUS_PADRAO = ['EM LEVANTAMENTO', '0', 'SEM INFORMAÇÕES', 'SEM INFORMACOES', 'CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO', 'PRÉ ANÁLISE', 'PRE ANALISE']
TIPOS_PRIORITARIOS = ["CCF", "DIF", "MGD", "MTP", "ASC", "SID"]

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} Metros"
        if 'POSTE' in c.upper(): return f"{int(round(vf))}"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def count_real_obras(r):
    if isinstance(r.get('_ORIGINAL_ROWS'), list): return len(r['_ORIGINAL_ROWS'])
    v = str(r.get('SUPER_PONTO', ''))
    return int(re.findall(r'\d+', v)[0]) if v.startswith('SIM') and re.findall(r'\d+', v) else 1

def limpar_colunas_excel(df_alvo, cols_originais):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    if 'BASE_ATRIBUIDA' in df_alvo.columns: df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'LEVANTADOR_RESPONSAVEL'})
    elif 'LEVANTADOR' in df_alvo.columns and 'LEVANTADOR_RESPONSAVEL' not in df_alvo.columns: df_alvo['LEVANTADOR_RESPONSAVEL'] = df_alvo['LEVANTADOR']
    bs = ['PROTOCOLO', 'LEVANTADOR_RESPONSAVEL', 'ORDEM', 'NOME_DIA', 'DIA_MES', 'PRIORIDADE', 'SUPER_PONTO']
    cg = ['REGIONAL', 'MUNICIPIO', 'LOCALIDADE', 'LATITUDE', 'LONGITUDE', 'TIPO NOTA', 'FASE', 'INFORMACOES EXTRAS']
    mc = [c for c in cols_originais if c in df_alvo.columns and c not in bs and c != 'LINK_NAVEGACAO_OFFLINE' and not str(c).startswith('_')]
    for c in cg:
        if c in df_alvo.columns and c not in bs and c != 'LINK_NAVEGACAO_OFFLINE' and c not in mc: mc.append(c)
    fc = []
    for c in bs + mc + ['LINK_NAVEGACAO_OFFLINE']:
        if c in df_alvo.columns and c not in fc: fc.append(c)
    return df_alvo[fc]

def tentar_rerun():
    try: st.rerun()
    except: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_tat': False, 'vrp_status_tat': "IDLE", 'vrp_state_tat': {}, 'df_routed_tat': pd.DataFrame(), 'bases_records_tat': [], 'colunas_exibir_tat': [], 'colunas_originais_tat': []})
    for k in ['bytes_zip_xl_tat', 'bytes_zip_kml_tat', 'bytes_zip_gpx_tat', 'start_time_run_tat', 'start_time_pkg_tat']: st.session_state.pop(k, None)
    ler_planilha_cached.clear(); tentar_rerun()

if "roteamento_concluido_tat" not in st.session_state: st.session_state.roteamento_concluido_tat = False
if "vrp_status_tat" not in st.session_state: st.session_state.vrp_status_tat = "IDLE"

status_exec = st.session_state.vrp_status_tat
is_done = st.session_state.roteamento_concluido_tat
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>🎯 Planejamento Tático (Automático)</h1>", unsafe_allow_html=True)
st.info("💡 A IA analisa 100% da base e distribui as obras. Se a equipe esgotar a cidade, o radar de 100km é acionado.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Esforço e Limites", expanded=True):
        trava_global = st.number_input("Trava Total de Operação", min_value=0, value=0, step=50, disabled=is_locked)
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (Metros)", 10, 1000, 100, 10, disabled=is_locked)
        modo_prod = st.checkbox("🔥 Alta Densidade", value=False, disabled=is_locked)
        min_vizinhos = st.slider("Mínimo obras próximas (2km):", 2, 50, 10, 1, disabled=is_locked) if modo_prod else 0
        st.markdown("---")
        tipo_periodo = st.radio("Agrupamento:", ["☀️ Dia", "📅 Semana"], index=1, horizontal=True, disabled=is_locked)
        tpc = "Semana" if "Semana" in tipo_periodo else "Dia"
        dias_sel = st.multiselect("Dias úteis:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked) if tpc == "Semana" else ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
        data_ini = st.date_input("📅 Data Início:", value=datetime.today(), disabled=is_locked)
        obras_dia = st.number_input("Obras por Dia", 1, 30, step=1, disabled=is_locked)
        limite_per = st.number_input(f"Limite total de {tpc}s", 1, 5, step=1, disabled=is_locked)
        vel_kmh, t_obra = 30.0, 1.5

    with st.expander("💰 Gestão Financeira", expanded=False):
        c_comb = st.number_input("Custo Combustível (R$/L)", min_value=0.0, value=0.0, step=0.1, disabled=is_locked)
        c_veic = st.number_input("Consumo (Km/L)", min_value=0.0, value=0.0, step=0.5, disabled=is_locked)
        c_hora = st.number_input("Hora-Homem (R$)", min_value=0.0, value=0.0, step=1.0, disabled=is_locked)
        
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=False, disabled=is_locked)
        
    st.markdown("---")
    sb_html = st.empty()
    if is_done and not st.session_state.df_routed_tat.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_tat', b""), file_name=f"Tatico_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_tat', b""), file_name=f"Tatico_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_tat', b""), file_name=f"Tatico_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed_tat.empty:
    st.markdown("## 🎯 Resultados da Otimização")
    st.session_state.df_routed_tat['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_tat.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed_tat.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = sum(count_real_obras(r) for _, r in dfr_t.iterrows())
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    tp = sum(count_real_obras(r) for _, r in dfr_t[dfr_t['PRIORIDADE'] == 'Sim'].iterrows()) if 'PRIORIDADE' in dfr_t else 0
    
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="metric-card" style="border-left:5px solid #0D256C;"><div class="metric-icon" style="background:rgba(13,37,108,0.12);">🎯</div><div class="metric-content"><div class="metric-title">TOTAL ROTEIRIZADAS</div><div class="metric-value">{tr}</div></div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="metric-card" style="border-left:5px solid #8b5cf6;"><div class="metric-icon" style="background:rgba(139,92,246,0.15);">👥</div><div class="metric-content"><div class="metric-title">Equipes Alocadas</div><div class="metric-value">{te}</div></div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="metric-card" style="border-left:5px solid #55B929;"><div class="metric-icon" style="background:rgba(85,185,41,0.15);">🛣️</div><div class="metric-content"><div class="metric-title">KM Previsto</div><div class="metric-value">{tk}</div></div></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="metric-card" style="border-left:5px solid #ef4444;"><div class="metric-icon" style="background:rgba(239,68,68,0.15);">🚨</div><div class="metric-content"><div class="metric-title">Prioridades</div><div class="metric-value">{tp}</div></div></div>', unsafe_allow_html=True)

    c_at = st.session_state.vrp_state_tat.get('config', {})
    m_eq = obras_dia * (len(dias_sel) if tpc == "Semana" else 1) * limite_per
    m_gb = m_eq * len(set(b['LEVANTADOR'] for b in st.session_state.bases_records_tat))
    if trava_global > 0: m_gb = min(m_gb, trava_global)
    f_obr = m_gb - tr
    
    if f_obr > 0: st.markdown(f'<div style="background:#fff3cd;color:#856404;padding:15px;border-left:5px solid #ffeeba;border-radius:4px;"><h4 style="margin:0;">⚠️ Alerta de Estoque: Faltaram {f_obr} obras para atingir a meta global.</h4><p style="margin:0;">O sistema roteirizou o máximo de obras compatíveis ({tr}). Verifique a aba de Déficit.</p></div>', unsafe_allow_html=True)
    else: st.markdown(f'<div style="background:#d4edda;color:#155724;padding:15px;border-left:5px solid #c3e6cb;border-radius:4px;"><h4 style="margin:0;">✅ Meta de Despacho 100% Atingida ({m_gb} Obras)!</h4></div>', unsafe_allow_html=True)

    if c_comb > 0 or c_hora > 0:
        km_v = dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum()
        cc_t = (km_v / c_veic * c_comb) if c_veic > 0 else 0
        df_fin = dfr.copy(); df_fin['_HORA_INICIO_DT'], df_fin['_HORA_FIM_DT'] = pd.to_datetime(df_fin['_HORA_INICIO_DT']), pd.to_datetime(df_fin['_HORA_FIM_DT'])
        ch_t = sum([(g['_HORA_FIM_DT'].max() - g['_HORA_INICIO_DT'].min()).total_seconds()/3600.0 for _, g in df_fin.groupby(['BASE_ATRIBUIDA', 'PERIODO'])]) * c_hora
        cf1, cf2, cf3, cf4 = st.columns(4)
        cf1.markdown(f'<div class="metric-card" style="border-left:5px solid #f59e0b;"><div class="metric-content"><div class="metric-title">⛽ Combustível Estimado</div><div class="metric-value">R$ {formatar_moeda(cc_t)}</div></div></div>', unsafe_allow_html=True)
        cf2.markdown(f'<div class="metric-card" style="border-left:5px solid #8b5cf6;"><div class="metric-content"><div class="metric-title">👷 Mão de Obra</div><div class="metric-value">R$ {formatar_moeda(ch_t)}</div></div></div>', unsafe_allow_html=True)
        cf3.markdown(f'<div class="metric-card" style="border-left:5px solid #ef4444;"><div class="metric-content"><div class="metric-title">💲 Custo Total</div><div class="metric-value">R$ {formatar_moeda(cc_t + ch_t)}</div></div></div>', unsafe_allow_html=True)
        cf4.markdown(f'<div class="metric-card" style="border-left:5px solid #55B929;"><div class="metric-content"><div class="metric-title">📊 Custo / Obra</div><div class="metric-value">R$ {formatar_moeda((cc_t+ch_t)/tr if tr>0 else 0)}</div></div></div>', unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Geográfico")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    HeatMap([[r['LATITUDE'], r['LONGITUDE']] for _, r in dfr_t.iterrows()], radius=15, blur=10).add_to(mapa)
    mc = MarkerCluster().add_to(mapa)
    
    for bn in dfr['BASE_ATRIBUIDA'].unique():
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
                ic = identificar_icone_folium(r, dfr.columns)
                if str(r.get('SUPER_PONTO', '')).startswith('SIM'):
                    ci, p_txt, p_bg, p_c = 'orange', f"🏢 SUPER PONTO {str(r.get('SUPER_PONTO')).replace('SIM','').strip()}", "#FFD700", "#000000"
                else:
                    ci, p_txt, p_bg, p_c = ('red', "🚨 OBRA PRIORITÁRIA", "#d9534f", "#ffffff") if r.get('PRIORIDADE')=='Sim' else ('blue', "📍 Atendimento Padrão", "#0D256C", "#ffffff")
                er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}:</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_tat if c not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=ci, icon=ic), popup=folium.Popup(f'<div style="width:280px;"><div style="background:{p_bg};color:{p_c};padding:8px;">{p_txt}</div><table border="1" style="width:100%;font-size:12px;"><tr><td><b>Ordem:</b></td><td>{r.get("ORDEM",0)}</td></tr>{er}</table></div>', max_width=300)).add_to(mc)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    t1, t2, t3 = st.tabs(["📊 Tabela", "📉 Déficit", "🏨 Hotéis"])
    with t1: st.data_editor(dfr.drop(columns=['ROTA_GEOMETRIA','_HORA_INICIO_DT','_HORA_FIM_DT','_ORIGINAL_ROWS','_ORIGEM_BASE','PERIODO','ALERTA_TOPOLOGIA','TEMPO_VIAGEM_MINUTOS'], errors='ignore'), use_container_width=True)
    with t2:
        df_dfc = pd.DataFrame([{"Levantador": b['LEVANTADOR'], "Obras": sum(count_real_obras(r) for _, r in dfr_t[dfr_t['BASE_ATRIBUIDA']==b['LEVANTADOR']].iterrows()), "Meta": m_eq, "Falta": max(0, m_eq - sum(count_real_obras(r) for _, r in dfr_t[dfr_t['BASE_ATRIBUIDA']==b['LEVANTADOR']].iterrows()))} for b in st.session_state.bases_records_tat]).sort_values(by="Falta", ascending=False)
        st.dataframe(df_dfc, use_container_width=True)
    with t3:
        hs = False
        for b in st.session_state.bases_records_tat:
            n, lt, ln = b['LEVANTADOR'], b.get('LATITUDE'), b.get('LONGITUDE')
            dt = dfr_t[dfr_t['BASE_ATRIBUIDA']==n]
            if not dt.empty and pd.notna(lt) and pd.notna(ln):
                cl, cL = dt['LATITUDE'].mean(), dt['LONGITUDE'].mean()
                d = haversine_scalar(float(lt), float(ln), cl, cL)
                if d > 60:
                    hs = True
                    st.markdown(f'<div style="background:#f8f9fa;border-left:5px solid #0D256C;padding:15px;margin-bottom:15px;"><h4 style="margin:0;">👨‍🔧 {n}</h4><p>⚠️ Pernoite Sugerido: Polo a {d:.1f} KM da base.</p><a href="https://www.google.com/maps/search/hoteis+pousadas/@{cl:.6f},{cL:.6f},12z" target="_blank">🏨 Buscar Hotéis no Polo</a></div>', unsafe_allow_html=True)
        if not hs: st.success("✅ Logística Segura: Nenhuma equipe com pernoite detectado (>60km).")

elif status_exec == "IDLE":
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 👥 1. Levantadores")
        df_bases, df_bases_temp = pd.DataFrame(), pd.DataFrame()
        bf = st.file_uploader("Principais (Excel)", type=["xlsx", "xls"])
        if bf:
            b_t = ler_planilha_cached(bf.getvalue()); b_t.columns = normalize_cols(b_t.columns)
            for pn in ['NOME', 'TECNICO', 'EQUIPE']:
                if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'LEVANTADOR'}); break
            if 'LEVANTADOR' in b_t.columns:
                opts = sorted([str(x) for x in b_t['LEVANTADOR'].dropna().unique() if str(x).upper().strip() != 'SEM LEVANTADOR'])
                sel = st.multiselect("Equipes Ativas:", opts, default=opts)
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
                    df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE']); df_bases['TIPO_EQUIPE'] = 'PRINCIPAL'
        
        ta = st.radio("Atribuição", ["Por Município", "Por Proximidade"], index=0, label_visibility="collapsed")
    with c2:
        st.markdown("### 📁 2. Demandas (Obras)")
        t_f = st.file_uploader("1️⃣ Levantamento", type=["xlsx", "xls"], accept_multiple_files=True)
        s_f = st.file_uploader("2️⃣ Saneamento", type=["xlsx", "xls"], accept_multiple_files=True)
        g_f = st.file_uploader("3️⃣ Genérica / Livre", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
        st_f = st.file_uploader("4️⃣ SharePoint (Opcional)", type=["xlsx", "xls"])
        df_su, cs_sel = pd.DataFrame(), None
        if st_f:
            df_su = ler_planilha_cached(st_f.getvalue()); cs_sel = st.selectbox("📌 Coluna Status?", df_su.columns.tolist(), index=4 if len(df_su.columns)>=5 else 0)
        
        with st.expander("🧑‍🤝‍🧑 Equipes de Apoio (Temporários)"):
            tb_f = st.file_uploader("Apoio", type=["xlsx", "xls"], accept_multiple_files=True)
            if tb_f:
                dbf = pd.concat([ler_planilha_cached(f.getvalue()).rename(columns=lambda c: 'LEVANTADOR' if c in ['NOME', 'TECNICO', 'EQUIPE'] else c) for f in tb_f], ignore_index=True)
                dbf.columns = normalize_cols(dbf.columns)
                if 'LEVANTADOR' in dbf.columns:
                    sel = st.multiselect("Equipes de Apoio:", sorted([str(x) for x in dbf['LEVANTADOR'].dropna().unique()]), default=sorted([str(x) for x in dbf['LEVANTADOR'].dropna().unique()]))
                    if sel:
                        df_bases_temp = dbf[dbf['LEVANTADOR'].isin(sel)].copy()
                        df_bases_temp['LATITUDE'] = pd.to_numeric(df_bases_temp.get('LATITUDE', np.nan), errors='coerce')
                        df_bases_temp['LONGITUDE'] = pd.to_numeric(df_bases_temp.get('LONGITUDE', np.nan), errors='coerce')
                        df_bases_temp = df_bases_temp.dropna(subset=['LATITUDE', 'LONGITUDE']); df_bases_temp['TIPO_EQUIPE'] = 'TEMPORARIA'

        ig_d = st.checkbox("Filtro: Ignorar obras já despachadas?", value=False)

    qe = df_bases.get('LEVANTADOR', pd.Series()).nunique() + df_bases_temp.get('LEVANTADOR', pd.Series()).nunique()
    cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per
    sb_html.markdown(renderizar_painel_lateral(cm, 0, qe, cm * max(1, qe)), unsafe_allow_html=True)

    if not t_f and not s_f and not g_f: st.info("Aguardando planilhas..."); st.stop()

    dfs = []
    for f in (t_f or []):
        d = ler_planilha_cached(f.getvalue()); d.columns = normalize_cols(d.columns)
        if not dfs: st.session_state.colunas_originais_tat = d.columns.tolist()
        d['_ORIGEM_BASE'] = 'LEVANTAMENTO'
        if 'PRIORIDADE' in d.columns: d['_PRIORIDADE_ORIGINAL'] = d['PRIORIDADE']
        for cc in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
            if cc in d.columns: d['PROTOCOLO'] = d[cc]; break
        dfs.append(d)
    for f in (s_f or []):
        d = ler_planilha_cached(f.getvalue()); d.columns = normalize_cols(d.columns)
        d['_ORIGEM_BASE'] = 'SANEAMENTO'
        if 'PRIORIDADE' in d.columns: d['_PRIORIDADE_ORIGINAL'] = d['PRIORIDADE']
        d['LATITUDE'] = d.get('LATITUDE PROJETO', d.get('LATITUDE'))
        d['LONGITUDE'] = d.get('LONGITUDE PROJETO', d.get('LONGITUDE'))
        for cc in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
            if cc in d.columns: d['PROTOCOLO'] = d[cc]; break
        dfs.append(d)
    for f in (g_f or []):
        d = pd.read_csv(f) if f.name.endswith('.csv') else ler_planilha_cached(f.getvalue()); d.columns = normalize_cols(d.columns)
        d['_ORIGEM_BASE'] = 'GENERICA'
        if 'PRIORIDADE' in d.columns: d['_PRIORIDADE_ORIGINAL'] = d['PRIORIDADE']
        if 'LATITUDE' not in d.columns or 'LONGITUDE' not in d.columns: continue
        fid = False
        for cc in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS', 'ID', 'CODIGO']:
            if cc in d.columns: d['PROTOCOLO'] = d[cc]; fid = True; break
        if not fid: d['PROTOCOLO'] = [f"GEN-{i+1}" for i in range(len(d))]
        dfs.append(d)

    if not dfs: st.stop()
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()
    for cc in ['CONTA CONTRATO', 'INSTALACAO', 'PROTOCOLO']:
        if cc in df_tasks.columns: df_tasks[cc] = df_tasks[cc].astype(str).replace(r'\.0$', '', regex=True).replace('nan', '-')

    st.markdown("##### 🗂️ Triagem Dinâmica")
    if 'TIPO NOTA' in df_tasks.columns:
        tc = df_tasks['TIPO NOTA'].value_counts()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("UNR", tc.get('UNR',0)); c2.metric("MGD", tc.get('MGD',0)); c3.metric("ASC", tc.get('ASC',0)); c4.metric("DIF", tc.get('DIF',0))
        trej = st.multiselect("🗑️ DESCARTAR Tipos:", options=[str(x) for x in df_tasks['TIPO NOTA'].dropna().unique()], default=[])
        if trej: df_tasks = df_tasks[~df_tasks['TIPO NOTA'].isin(trej)]
    if not df_su.empty and cs_sel: df_tasks = atualizar_status_via_df(df_tasks, df_su, cs_sel)
    if ig_d and 'DATA DESPACHO CAMPO' in df_tasks.columns:
        md = df_tasks['DATA DESPACHO CAMPO'].notna() & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip() != '') & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip().str.lower() != 'nan')
        df_tasks = df_tasks[~md]

    st.markdown("### 🎯 Escopo da Operação")
    ce1, ce2 = st.columns(2)
    cr = 'Regional' if 'Regional' in df_tasks.columns else 'REGIONAL' if 'REGIONAL' in df_tasks.columns else None
    if cr:
        sr = ce1.multiselect("🌍 REGIONAL:", options=[str(x) for x in df_tasks[cr].dropna().unique()], default=[])
        if sr: df_tasks = df_tasks[df_tasks[cr].isin(sr)]
    if 'PAT' in df_tasks.columns:
        sp = ce2.multiselect("🏗️ PAT:", options=[str(x) for x in df_tasks['PAT'].dropna().unique()], default=[])
        if sp: df_tasks = df_tasks[df_tasks['PAT'].isin(sp)]

    dfls = []
    if 'LEVANTAMENTO' in df_tasks['_ORIGEM_BASE'].values:
        with st.expander("🛠️ Filtros - LEVANTAMENTO", expanded=True):
            dfl = df_tasks[df_tasks['_ORIGEM_BASE'] == 'LEVANTAMENTO'].copy()
            cf1, cf2, cf3 = st.columns(3)
            if 'STATUS LIST' in dfl.columns:
                op = sorted(list(set([str(x).strip().upper() for x in dfl['STATUS LIST'].dropna() if str(x).lower() != 'nan'])))
                sl = cf1.multiselect("📌 Status de Início:", op, default=[s for s in op if s in STATUS_PADRAO])
                if sl: dfl = dfl[dfl['STATUS LIST'].astype(str).str.strip().str.upper().isin(sl)]
            cv = [c for c in dfl.columns if not c.startswith('_')]
            id_d = next((i + 1 for i, c in enumerate(cv) if c in ['TIPO NOTA', 'TIPO DE NOTA']), 0)
            cp = cf2.selectbox("📌 1. Coluna de prioridade?", ["Nenhuma"] + cv, index=id_d)
            if cp != "Nenhuma":
                op = sorted(list(set(dfl[cp].fillna('SEM TIPO').astype(str).str.strip().str.upper().unique())))
                sv = cf2.multiselect(f"🏷️ 2. Filtrar '{cp}':", op, default=op)
                if sv: dfl = dfl[dfl[cp].astype(str).str.strip().str.upper().isin(sv)]
                sp = cf3.multiselect(f"🚨 3. PRIORIDADE:", sv, default=[x for x in sv if x in TIPOS_PRIORITARIOS])
                dfl['PRIORIDADE'] = dfl[cp].astype(str).str.strip().str.upper().apply(lambda x: 'Sim' if x in sp else 'Não')
                st.session_state.col_prioridade_tat = cp
            else: dfl['PRIORIDADE'] = 'Não'
            if 'STATUS SAP' in dfl.columns: dfl = dfl[~dfl['STATUS SAP'].astype(str).str.strip().str.upper().isin(['CANC', 'FINL'])]
            dfls.append(dfl)
            
    if 'SANEAMENTO' in df_tasks['_ORIGEM_BASE'].values:
        dfs = df_tasks[df_tasks['_ORIGEM_BASE'] == 'SANEAMENTO'].copy(); dfs['PRIORIDADE'] = 'Não'; dfls.append(dfs)
        
    if 'GENERICA' in df_tasks['_ORIGEM_BASE'].values:
        with st.expander("🛠️ Filtros - GENÉRICA", expanded=True):
            dfg = df_tasks[df_tasks['_ORIGEM_BASE'] == 'GENERICA'].copy()
            c1, c2 = st.columns(2)
            cv = [c for c in dfg.columns if not c.startswith('_')]
            id_d = cv.index('TIPO NOTA') + 1 if 'TIPO NOTA' in cv else 0
            cp = c1.selectbox("📌 Prioridade?", ["Nenhuma"] + cv, index=id_d)
            if cp != "Nenhuma":
                op = [x for x in dfg[cp].fillna('SEM TIPO').astype(str).str.strip().unique() if x.lower() != 'nan']
                sp = c2.multiselect(f"🚨 Obras PRIORITÁRIAS:", op, default=[x for x in op if x in TIPOS_PRIORITARIOS] if cp=='TIPO NOTA' else [])
                dfg['PRIORIDADE'] = dfg[cp].astype(str).str.strip().apply(lambda x: 'Sim' if x in sp else 'Não')
            else: dfg['PRIORIDADE'] = 'Não'
            dfls.append(dfg)

    if not dfls: st.stop()
    df_tasks = pd.concat(dfls, ignore_index=True)
    
    if 'STATUS LIST' in df_tasks.columns: df_tasks.loc[df_tasks['STATUS LIST'].astype(str).str.strip().str.upper().isin(['CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO']), 'PRIORIDADE'] = 'Sim'
    if '_PRIORIDADE_ORIGINAL' in df_tasks.columns:
        m_nz = df_tasks['_PRIORIDADE_ORIGINAL'].notna() & (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '') & (~df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper().isin(['0','0.0','NAN','NÃO','NAO','FALSE']))
        df_tasks.loc[m_nz, 'PRIORIDADE'] = 'Sim'

    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    ec = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna() | (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    if ec.sum() > 0: st.warning(f"⚠️ {ec.sum()} obras ignoradas (coordenadas inválidas)."); df_tasks = df_tasks[~ec]
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']
    
    if 'NOME' not in df_tasks.columns: df_tasks['NOME'] = "SEM NOME"
    for cn in ['NOME DO SOLICITANTE', 'CLIENTE', 'RAZAO SOCIAL', 'DESCRICAO', 'ENDERECO', 'LOCAL']:
        if cn in df_tasks.columns: df_tasks['NOME'] = df_tasks['NOME'].fillna(df_tasks[cn])
    df_tasks = df_tasks[~(df_tasks['_ORIGEM_BASE'].isin(['LEVANTAMENTO', 'SANEAMENTO']) & (df_tasks['NOME'].isna() | (df_tasks['NOME'].astype(str).str.strip() == '') | (df_tasks['NOME'].astype(str).str.strip().str.lower() == 'nan')))]

    df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_sp, agrupar_por_levantador=False)
    
    tbr = (df_bases.to_dict('records') if not df_bases.empty else []) + (df_bases_temp.to_dict('records') if not df_bases_temp.empty else [])
    cmn = 'MUNICIPIO' if 'MUNICIPIO' in df_tasks.columns else 'CIDADE' if 'CIDADE' in df_tasks.columns else None
    df_tasks['MUN_LIMPO'] = normalizar_municipios(df_tasks[cmn].fillna('')) if cmn else ''
    
    mtm, mta, mag = {}, {}, set()
    for b in tbr:
        for m in str(b.get('MUNICIPIO', b.get('RESIDENCIA', ''))).split(','):
            ml = normalizar_municipios(pd.Series([m])).iloc[0]
            if ml:
                mag.add(ml)
                if ml not in mta: mta[ml] = []
                if b['LEVANTADOR'] not in mta[ml]: mta[ml].append(b['LEVANTADOR'])
                if b.get('TIPO_EQUIPE') == 'PRINCIPAL':
                    if ml not in mtm: mtm[ml] = []
                    if b['LEVANTADOR'] not in mtm[ml]: mtm[ml].append(b['LEVANTADOR'])

    if trava_global > 0 and len(df_tasks) > trava_global:
        df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
        oaf, cr = [], 0
        for _, r in df_tasks.iterrows():
            q = len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
            if cr + q <= trava_global: oaf.append(r); cr += q
            else: break
        df_tasks = pd.DataFrame(oaf)

    df_sp = pd.DataFrame()
    if modo_prod:
        with st.spinner("Mapeando bolsões de alta densidade..."):
            lt, ln, pm, mv = df_tasks['LATITUDE'].values, df_tasks['LONGITUDE'].values, (df_tasks['PRIORIDADE'] == 'Sim').values, df_tasks['MUN_LIMPO'].values
            km = np.copy(pm)
            for i in range(len(df_tasks)):
                if mag and mv[i] not in mag: km[i] = False; continue
                if not km[i]:
                    if np.sum(haversine_vectorized(lt[i], ln[i], lt, ln) <= 2.0) >= min_vizinhos: km[i] = True
            df_sp = df_tasks[~km].copy()
            if not df_sp.empty: df_sp['BASE_ATRIBUIDA'], df_sp['MOTIVO_REJEICAO'] = "NÃO ALOCADO", "Obra Isolada"
            df_tasks = df_tasks[km].copy()

    df_ta = pd.DataFrame()
    if tbr:
        df_tasks['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
        df_tasks['COORD_KEY'] = df_tasks['LATITUDE'].astype(str) + "_" + df_tasks['LONGITUDE'].astype(str)
        df_pp, df_cp = df_tasks[df_tasks['COORD_KEY'].isin(df_tasks[df_tasks['PRIORIDADE'] == 'Sim']['COORD_KEY'].unique())].copy(), df_tasks[~df_tasks['COORD_KEY'].isin(df_tasks[df_tasks['PRIORIDADE'] == 'Sim']['COORD_KEY'].unique())].copy()
        
        bc = {b['LEVANTADOR']: 0 for b in tbr}
        at, ut = [], []
        for r in pd.concat([df_pp, df_cp], ignore_index=True).sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True]).to_dict('records'):
            qr = len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
            la, lo, ms, ip, p4 = r.get('LATITUDE'), r.get('LONGITUDE'), str(r.get('MUN_LIMPO', '')), r.get('PRIORIDADE') == 'Sim', str(r.get('TIPO VEICULO', '')).strip().upper() == '4X4'
            vn = set(mtm.get(ms, [])) if ip else set(mta.get(ms, []))
            vb = sorted([b for b in tbr if b['LEVANTADOR'] in vn and (not p4 or str(b.get('VEICULO', '')).upper() == '4X4')], key=lambda x: bc[x['LEVANTADOR']])
            bb, bd = None, float('inf')
            if pd.notna(la) and pd.notna(lo):
                for b in vb:
                    bn = b['LEVANTADOR']
                    if bc[bn] + qr <= cm:
                        d = haversine_scalar(la, lo, float(b.get('LATITUDE', 0)), float(b.get('LONGITUDE', 0)))
                        if d < bd: bd, bb = d, bn
            if bb: bc[bb] += qr; r['BASE_ATRIBUIDA'] = bb; at.append(r)
            else: ut.append(r)
                
        su = []
        for r in ut:
            qr, la, lo, p4 = len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1, r.get('LATITUDE'), r.get('LONGITUDE'), str(r.get('TIPO VEICULO', '')).strip().upper() == '4X4'
            vbf = sorted([b for b in tbr if not p4 or str(b.get('VEICULO', '')).upper() == '4X4'], key=lambda x: bc[x['LEVANTADOR']])
            bbf = None
            if pd.notna(la) and pd.notna(lo):
                for b in vbf:
                    bn = b['LEVANTADOR']
                    if bc[bn] + qr <= cm and haversine_scalar(la, lo, float(b.get('LATITUDE', 0)), float(b.get('LONGITUDE', 0))) <= 100.0: bbf = bn; break
            if bbf: bc[bbf] += qr; r['BASE_ATRIBUIDA'] = bbf; at.append(r)
            else: r['BASE_ATRIBUIDA'], r['MOTIVO_REJEICAO'] = "NÃO ALOCADO", "Estoque Lotado ou > 100km"; su.append(r)

        df_ta, df_u = pd.DataFrame(at), pd.DataFrame(su)
        if not df_sp.empty: df_u = pd.concat([df_u, df_sp], ignore_index=True)
        st.session_state.df_unallocated, st.session_state.tot_obras_nao_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_u.iterrows())
        sb_html.markdown(renderizar_painel_lateral(cm, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows()), qe, cm * qe), unsafe_allow_html=True)
        
        if df_ta.empty: st.error("Nenhuma obra alocada."); st.stop()
        
        with st.expander("🛠️ Configuração de Saída", expanded=True):
            tc = [c for c in df_ta.columns if not c.startswith('_')]
            cd = ['PROTOCOLO', 'CONTA CONTRATO', 'INSTALACAO', 'NOME', 'ENDERECO', 'INFORMACOES EXTRAS', 'LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'TIPO NOTA', 'FASE']
            st.session_state.colunas_exibir_tat = st.multiselect("Colunas Visíveis:", tc, default=[c for c in cd if c in tc])
            
        if st.button("🚀 Iniciar Motor VRP", type="primary", use_container_width=True):
            if tpc == "Semana" and not dias_sel: st.error("Selecione os dias da semana."); st.stop()
            st.session_state.update({'bases_records_tat': tbr, 'col_prioridade_tat': st.session_state.get('col_prioridade_lev', "PRIORIDADE")})
            st.session_state.vrp_state_tat = {'config': {'velocidade_media_kmh': vel_kmh, 'tempo_medio_obra': t_obra, 'obras_por_dia': obras_dia, 'tipo_periodo': tpc, 'limite_periodos': limite_per, 'dias_selecionados': dias_sel, 'url_osrm_base': url_osrm, 'tracado_real': usa_osrm, 'data_inicio': data_ini, 'sentido_rota': sentido_rota, 'trava_global_obras': trava_global}, 'b_names': list(set([b['LEVANTADOR'] for b in tbr])), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': [], 'current_geoms': []}
            st.session_state.vrp_status_tat = "RUNNING"; tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor de Inteligência (OR-Tools VRP)")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_tat', time.time())
    if 'start_time_run_tat' not in st.session_state: st.session_state.start_time_run_tat = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_tat; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex;gap:10px;margin-bottom:15px;"><div style="flex:1;padding:15px;background:#f8f9fa;border:1px solid #dee2e6;text-align:center;"><div style="font-size:0.85rem;color:#6c757d;font-weight:bold;">⏱️ Decorrido</div><div style="font-size:1.8rem;color:#0D256C;">{es}</div></div><div style="flex:1;padding:15px;background:#e8f5e9;border:1px solid #a5d6a7;text-align:center;"><div style="font-size:0.85rem;color:#2e7d32;font-weight:bold;">🎯 Restante</div><div style="font-size:1.8rem;color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Analisando **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            dfb = pd.DataFrame(st.session_state.bases_records_tat)
            br = dfb[dfb['LEVANTADOR'] == bn].iloc[0]
            if pd.isna(br.get('LATITUDE')): st_v['b_idx'] += 1; st.session_state.vrp_state_tat = st_v; tentar_rerun(); st.stop()
            bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            
            ot = []
            if oe:
                cd = {}
                for o in oe:
                    k = (round(float(o['LATITUDE']), 4), round(float(o['LONGITUDE']), 4))
                    if k not in cd: cd[k] = []
                    cd[k].append(o)
                mo = []
                for k, l in cd.items():
                    tp = any(x.get('PRIORIDADE') == 'Sim' for x in l)
                    rp = l[0].copy(); rp['PRIORIDADE'] = 'Sim' if tp else 'Não'
                    ml = normalizar_municipios(pd.Series([rp.get('MUNICIPIO', rp.get('CIDADE', ''))])).iloc[0]
                    rp['MUN_LIMPO'] = ml if ml else 'DESCONHECIDO'
                    rp['_so'] = l; mo.append(rp)
                
                mbm = {}
                for m in mo:
                    mu = m['MUN_LIMPO']
                    if mu not in mbm: mbm[mu] = []
                    mbm[mu].append(m)
                
                ms = sorted([{'m': mu, 'pc': sum(1 for x in ml if x['PRIORIDADE']=='Sim'), 'tc': len(ml)} for mu, ml in mbm.items()], key=lambda x: (x['pc']>0, x['pc'], x['tc']), reverse=True)
                
                om = []
                for s in ms:
                    ml = mbm[s['m']]
                    pm, cm = [m for m in ml if m['PRIORIDADE']=='Sim'], [m for m in ml if m['PRIORIDADE']!='Sim']
                    if pm:
                        tr = resolver_tsp_ortools(pm, bl, bL, cfg['url_osrm_base'])
                        if "Reversa" in cfg['sentido_rota'] and tr:
                            fi = max(range(len(tr)), key=lambda i: haversine_scalar(bl, bL, float(tr[i]['LATITUDE']), float(tr[i]['LONGITUDE'])))
                            tr = tr[fi:] + tr[:fi]
                        om.extend(tr)
                    if cm:
                        tr = resolver_tsp_ortools(cm, bl, bL, cfg['url_osrm_base'])
                        if "Reversa" in cfg['sentido_rota'] and tr:
                            fi = max(range(len(tr)), key=lambda i: haversine_scalar(bl, bL, float(tr[i]['LATITUDE']), float(tr[i]['LONGITUDE'])))
                            tr = tr[fi:] + tr[:fi]
                        om.extend(tr)
                for m in om:
                    for s in sorted(m['_so'], key=lambda x: 0 if x.get('PRIORIDADE')=='Sim' else 1):
                        s['MLC'] = m['MUN_LIMPO']; ot.append(s)
            
            rf, da, sa, dds, ma = [], 1, 1, 1, None
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
                mu, qr = o.get('MLC', ''), len(o.get('_ORIGINAL_ROWS', [1])) if isinstance(o.get('_ORIGINAL_ROWS'), list) else 1
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
                
                vd = False
                if es['oh'] > 0 and (es['oh'] + qr > cfg['obras_por_dia']): vd = True
                elif ma is not None and mu != ma and es['oh'] > 0 and vkr > 100.0: vd = True
                
                if vd:
                    dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                    rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
                    da += 1
                    if cfg['tipo_periodo'] == "Semana":
                        dds += 1
                        if dds > len(cfg['dias_selecionados']): sa += 1; dds = 1
                    es = gi(da)
                    vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE']); vk = vkr * 1.3
                    if vkr < 0.05 and es['oh'] > 0: vm, em = 0.0, 30.0
                    else: vm, em = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60, cfg['tempo_medio_obra']*60
                    cp = es['t'] + pd.Timedelta(minutes=vm); fp = cp + pd.Timedelta(minutes=em)
                
                rf.append({'o': o, 'il': False, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': cp, 'hf': fp, 'vm': vm, 'dk': vk})
                es['l'], es['L'], es['t'], es['oh'] = o['LATITUDE'], o['LONGITUDE'], fp, es['oh'] + qr; ma = mu
                
            if es['oh'] > 0:
                dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
            
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_tat = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (30 if cfg['tracado_real'] else len(rf)), len(rf))
            for i in range(oi, ei):
                it = rf[i]
                if not cfg['tracado_real']: gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
                else:
                    if i%5==0: sgt.info(f"🛣️ Traçado real **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf)); time.sleep(0.05)
                    try: gd.append(obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh']))
                    except: gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_tat = st_v; tentar_rerun(); st.stop()
            
            bl, bL = float(pd.DataFrame(st.session_state.bases_records_tat)[pd.DataFrame(st.session_state.bases_records_tat)['LEVANTADOR']==bn].iloc[0]['LATITUDE']), float(pd.DataFrame(st.session_state.bases_records_tat)[pd.DataFrame(st.session_state.bases_records_tat)['LEVANTADOR']==bn].iloc[0]['LONGITUDE'])
            rdf, og, dp = [], 1, ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            for it, (g, ds) in zip(rf, gd):
                pv = it['s'] if cfg['tipo_periodo']=="Semana" else it['d']
                dn = dp[datetime.strptime(it['dm'], '%d/%m/%Y').weekday()] if cfg['tipo_periodo']=="Semana" else f"Dia {it['d']}"
                if it['il']: rdf.append({'PROTOCOLO': 'PAUSA_ALMOCO', 'NOME': '🍔 ALMOÇO', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': 0.0, 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']})
                elif it['ir']: rdf.append({'PROTOCOLO': 'RETORNO_BASE', 'NOME': 'BASE_RETORNO', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': round(it['dk'], 2), 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']})
                else:
                    ob = it['o']; ob['ORDEM'], ob['NOME_DIA'], ob['DIA_MES'], ob['SEMANA'], ob['DIA'], ob['PERIODO'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, dn, it['dm'], it['s'], it['d'], pv, round(it['dk'], 2)
                    ob['ROTA_GEOMETRIA'] = [[it['Lt'], it['lt']], [it['Lt'], it['lt']]] if it['la']==bl and it['La']==bL else g
                    ob['HORA_INICIO'], ob['HORA_FIM'], ob['_HORA_INICIO_DT'], ob['_HORA_FIM_DT'] = it['hi'].strftime('%H:%M'), it['hf'].strftime('%H:%M'), it['hi'], it['hf']
                    rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state_tat = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Matrizes Resolvidas!"); pb.progress(1.0)
        st.session_state.df_routed_tat = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status_tat = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    df_routed, d_fmt = st.session_state.df_routed_tat, datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx = io.BytesIO(), io.BytesIO(), io.BytesIO()
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg:
            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                br = next((x for x in st.session_state.bases_records_tat if x['LEVANTADOR']==b), None)
                qc, qp, qs = sum(count_real_obras(r) for _, r in db[db['PRIORIDADE']=='Não'].iterrows()), sum(count_real_obras(r) for _, r in db[db['PRIORIDADE']=='Sim'].iterrows()), len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                pt, pm = pd.to_numeric(db['POSTE PREVISTO BT'], errors='coerce').replace(0, np.nan) if 'POSTE PREVISTO BT' in db.columns else pd.Series(dtype=float), pd.to_numeric(db['POSTE PREVISTO MT'], errors='coerce').replace(0, np.nan) if 'POSTE PREVISTO MT' in db.columns else pd.Series(dtype=float)
                pms = pd.concat([pt, pm], axis=1).min(axis=1).fillna(0).round().astype(int) if not pt.empty or not pm.empty else pd.to_numeric(db['POSTES PREVISTOS'], errors='coerce').fillna(0).round().astype(int) if 'POSTES PREVISTOS' in db.columns else pd.Series([0]*len(db))
                du, su = db['DIA_MES'].nunique() if 'DIA_MES' in db.columns else db['DIA'].nunique(), db['SEMANA'].nunique()
                res.append({'LEVANTADOR': b, 'TIPO EQUIPE': br.get('TIPO_EQUIPE', 'PRINCIPAL') if br else 'DESCONHECIDO', 'OBRAS COMUNS': qc, 'OBRAS PRIORITARIAS': qp, 'SUPER PONTOS': qs, 'TOTAL OBRAS': qc+qp, 'POSTES PREVISTOS TOTAIS': int(pms.sum()), 'POSTES PREVISTOS / DIA': int(round(int(pms.sum())/du)) if du>0 else 0, 'POSTES PREVISTOS / SEMANA': int(round(int(pms.sum())/su)) if su>0 else 0, 'KM TOTAL PREVISTO': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)})
            zx.writestr(f"Resumo_Levantadores - {d_fmt}.xlsx", gerar_excel_resumo_bytes(pd.DataFrame(res)))
            
            dfg = limpar_colunas_excel(df_routed.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS'], errors='ignore'), st.session_state.colunas_originais_tat)
            for cc in dfg.columns:
                if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Tatica - {d_fmt}.xlsx", gerar_excel_bytes(dfg, "PRIORIDADE"))
            
            dfk = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            ks = gerar_kml_agrupado(dfk, st.session_state.bases_records_tat, f"ROTA_TOTAL", st.session_state.colunas_exibir_tat, df_routed['BASE_ATRIBUIDA'].unique().tolist(), st.session_state.vrp_state_tat['config']['tipo_periodo'], formatar_valor_coluna)
            zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", re.sub(r'<Placemark>(?:(?!</Placemark>).)*?<name>(?:(?!</name>).)*?BASE:(?:(?!</name>).)*?</name>(?:(?!</Placemark>).)*?</Placemark>', '', ks, flags=re.IGNORECASE | re.DOTALL).encode('utf-8'))
            zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk, "ROTA TOTAL").encode('utf-8'))
            
            df_u = st.session_state.get('df_unallocated', pd.DataFrame())
            if not df_u.empty:
                ku = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document><name>OBRAS NÃO ALOCADAS</name>', '<Style id="wp"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href></Icon></IconStyle></Style>']
                for _, r in df_u.iterrows():
                    if pd.notna(r.get('LATITUDE')) and pd.notna(r.get('LONGITUDE')): ku.append(f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO", "Rejeitado")))}</name><styleUrl>#wp</styleUrl><Point><coordinates>{r.get("LONGITUDE")},{r.get("LATITUDE")}</coordinates></Point></Placemark>')
                ku.append('</Document></kml>'); zk.writestr(f"OBRAS_NAO_ALOCADAS - {d_fmt}.kml", "\n".join(ku).encode('utf-8'))
            
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                ns = re.sub(r'[^A-Za-z0-9_ ]', '', str(b)).replace(" ", "_").upper()
                db = df_routed[df_routed['BASE_ATRIBUIDA'] == b]; dk = db[~db['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
                dx = limpar_colunas_excel(db.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS'], errors='ignore'), st.session_state.colunas_originais_tat)
                for c in dx.columns:
                    if str(dx[c].dtype) == 'object': dx[c] = dx[c].astype(str).replace('nan', '')
                zx.writestr(f"ROTA_{ns} - {d_fmt}.xlsx", gerar_excel_bytes(dx, "PRIORIDADE"))
                kl = gerar_kml_agrupado(dk, st.session_state.bases_records_tat, f"ROTA_{ns}", st.session_state.colunas_exibir_tat, [b], st.session_state.vrp_state_tat['config']['tipo_periodo'], formatar_valor_coluna)
                zk.writestr(f"ROTA_{ns} - {d_fmt}.kml", re.sub(r'<Placemark>(?:(?!</Placemark>).)*?<name>(?:(?!</name>).)*?BASE:(?:(?!</name>).)*?</name>(?:(?!</Placemark>).)*?</Placemark>', '', kl, flags=re.IGNORECASE | re.DOTALL).encode('utf-8'))
                zg.writestr(f"GPS_{ns} - {d_fmt}.gpx", gerar_gpx_simples(dk, f"ROTA_{ns}").encode('utf-8'))

        st.session_state.bytes_zip_xl_tat, st.session_state.bytes_zip_kml_tat, st.session_state.bytes_zip_gpx_tat = bu_xl.getvalue(), bu_kml.getvalue(), bu_gpx.getvalue()
        st.session_state.roteamento_concluido_tat = True; st.session_state.vrp_status_tat = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_tat = "IDLE"
