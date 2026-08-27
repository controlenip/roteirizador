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
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

from modules.data_processing import ler_planilha_cached, formata_campo_html, formatar_moeda, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas

from modules.export_lista import injetar_logo, identificar_icone_folium, gerar_excel_lista, gerar_excel_resumo_lista, gerar_gpx_simples, gerar_kml_lista, limpar_colunas_lista, gerar_txt_lista

st.set_page_config(page_title="Roteirizador - Lista Contínua", page_icon="📜", layout="wide")
injetar_logo()

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        if 'POSTE' in c.upper(): return str(int(float(v)))
        
        vf = float(v)
        if c.upper() in ['DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM']: 
            return f"{vf:.2f} KM"
        elif 'DISTANCIA' in c.upper(): 
            return f"{vf:.2f} Metros"
            
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(total_obras, qtd_equipes_ativas):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Modo:</b> <span style="color: #d9534f; font-weight: bold;">Lista Contínua</span></p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras}</span></p>
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
    st.session_state.update({'roteamento_concluido_lista': False, 'vrp_status_lista': "IDLE", 'vrp_state_lista': {}, 'df_routed_lista': pd.DataFrame(), 'colunas_exibir_lista': [], 'colunas_originais_lista': []})
    for k in ['bytes_zip_xl_lista', 'bytes_zip_kml_lista', 'bytes_zip_gpx_lista', 'bytes_zip_txt_lista', 'start_time_run_lista', 'start_time_pkg_lista', 'df_unallocated_lista', 'df_correcao_lista']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido_lista" not in st.session_state: st.session_state.roteamento_concluido_lista = False
if "vrp_status_lista" not in st.session_state: st.session_state.vrp_status_lista = "IDLE"

status_exec = st.session_state.vrp_status_lista
is_done = st.session_state.roteamento_concluido_lista
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📜 Lista Contínua</h1>", unsafe_allow_html=True)
st.info("💡 Gera uma lista de execução contínua com base na atribuição já existente na planilha de obras.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Parâmetros de Rota", expanded=True):
        st.success("📦 **Modo Contínuo:** Todas as obras serão alocadas numa lista única.")
        trava_global = st.number_input("Trava Total de Obras", min_value=0, value=0, step=50, disabled=is_locked)
        data_ini = st.date_input("📅 Data Base (Início):", value=datetime.today(), disabled=is_locked)
        
        st.markdown("---")
        obras_por_dia_est = st.number_input("Meta de Obras/Dia (Cálculo de Postes):", min_value=1, value=4, disabled=is_locked)
        
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)
        
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_lista.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_lista', b"vazio"), file_name=f"ListaContinua_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("📝 Baixar Relatórios (TXT)", data=st.session_state.get('bytes_zip_txt_lista', b"vazio"), file_name=f"ListaContinua_TXT - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_lista', b"vazio"), file_name=f"ListaContinua_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_lista', b"vazio"), file_name=f"ListaContinua_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed_lista.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    df_c = st.session_state.get('df_correcao_lista', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica:</b> Estas obras apresentaram coordenadas em branco, zeradas ou invertidas. Elas foram isoladas para não corromper o mapa.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.df_routed_lista['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_lista.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed_lista.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = len(dfr_t)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    
    tsp = sum(1 for _, r in dfr_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Roteirizadas", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(tsp), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    html_msg = f"""
    <div style='background-color: #d4edda; border-left: 8px solid #28a745; padding: 20px; border-radius: 8px; margin-bottom: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
        <h2 style='color: #155724; margin-top: 0; font-weight: 900; letter-spacing: 1px; font-size: 24px;'>🎉 100% DAS OBRAS ROTEIRIZADAS!</h2>
        <p style='color: #155724; font-size: 16px; margin-bottom: 10px;'>Todas as <b>{tr} tarefas válidas</b> presentes na sua planilha foram organizadas e processadas com sucesso!</p>
        <hr style='border-top: 1px solid #c3e6cb; margin: 15px 0;'>
        <p style='color: #155724; font-size: 14px; margin-bottom: 0;'>
            <b>💡 Resumo Técnico:</b> O módulo de <b>Lista Contínua</b> opera sem limites de jornada ou cota de dias. A Inteligência Artificial respeitou a atribuição que você já fez na planilha e traçou a melhor rota possível (do ponto A ao Z) para <b>todas</b> as obras de cada levantador, sem deixar absolutamente nada para trás.
        </p>
    </div>
    """
    st.markdown(html_msg, unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Operacional")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        pts = [p for _, r in db.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
        folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
        folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
        
        for r in db.to_dict('records'):
            if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
            c_i = 'red' if str(r.get('PRIORIDADE')) == 'Sim' else 'blue'
            ic = identificar_icone_folium(r, dfr.columns)
            er = "".join([f"<tr><td><b>{html.escape(c)}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_lista if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
            pop_html = f'<div style="width:250px;"><b>Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
            folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📊 Dados Tabulares")
    st.data_editor(st.session_state.df_routed_lista.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM'], errors='ignore'), use_container_width=True)

elif status_exec == "IDLE":
    st.markdown("### 📁 Demandas (Obras e Equipes)")
    st.info("💡 Apenas faça o upload da sua planilha de Obras. A IA identificará o Levantador/Fiscal Responsável automaticamente.")
    task_files = st.file_uploader("Suba a planilha com a Lista Contínua", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if not task_files: st.stop()
    
    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        if not dfs: st.session_state.colunas_originais_lista = dft.columns.tolist()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    agent_col = None
    for col in ['LEVANTADOR_RESPONSAVEL', 'LEVANTADOR', 'FISCAL', 'EQUIPE', 'NOME_FISCAL']:
        if col in df_tasks.columns:
            agent_col = col; break
            
    if not agent_col:
        st.error("❌ A planilha de Obras não possui uma coluna de responsável. Insira uma coluna chamada 'LEVANTADOR' ou 'FISCAL'.")
        st.stop()
        
    df_tasks = df_tasks.rename(columns={agent_col: 'BASE_ATRIBUIDA'})
    df_tasks['BASE_ATRIBUIDA'] = df_tasks['BASE_ATRIBUIDA'].astype(str).str.strip().str.upper()
    df_tasks = df_tasks[~df_tasks['BASE_ATRIBUIDA'].isin(['NAN', 'NONE', ''])]

    qtd_eq = df_tasks['BASE_ATRIBUIDA'].nunique()
    sb_html.markdown(render_sidebar_card("Ilimitada", qtd_eq), unsafe_allow_html=True)

    st.markdown("---")
    falta = [c for c in ['LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas vitais: {', '.join(falta)}."); st.stop()
    
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
        if 'PRIORIDADE' in df_tasks.columns:
            df_tasks['PRIORIDADE'] = df_tasks['PRIORIDADE'].apply(
                lambda x: 'Sim' if pd.notna(x) and str(x).strip() not in ['', '0', 'NÃO', 'NAO', 'FALSE'] else 'Não'
            )
            st.success("🎯 **Prioridade Ativada:** As obras marcadas na coluna 'PRIORIDADE' serão destacadas no roteiro e no KML.")
        elif 'TIPO NOTA' in df_tasks.columns:
            df_tasks['TIPO NOTA'] = df_tasks['TIPO NOTA'].astype(str).str.strip().str.upper()
            opts_n = sorted([str(x) for x in df_tasks['TIPO NOTA'].unique() if str(x) != 'NAN'])
            sel_p = st.multiselect("🚨 2. Obras de Alta Prioridade:", options=opts_n, default=[n for n in opts_n if n in ['ASC', 'CCF', 'DIF', 'MGD', 'MTP', 'SID']])
            df_tasks['PRIORIDADE'] = df_tasks['TIPO NOTA'].apply(lambda x: 'Sim' if str(x) in sel_p else 'Não')
        else: 
            df_tasks['PRIORIDADE'] = 'Não'

    st.markdown("#### 🌍 Limpeza Geográfica")
    if st.button("⏹️ Abortar Roteamento", use_container_width=True): limpar_roteirizador(); st.stop()
    
    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
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
    
    st.session_state.df_correcao_lista = df_rej
    if not df_rej.empty: 
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_rej)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa:</b> Apresentaram coordenadas zeradas, invertidas ou em branco. Foram removidas da roteirização.
            </p>
        </div>
        """, unsafe_allow_html=True)

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    dfs_fundidos = []
    for base in df_tasks['BASE_ATRIBUIDA'].unique():
        df_base = df_tasks[df_tasks['BASE_ATRIBUIDA'] == base].copy()
        df_base_f, _ = fundir_super_pontos(df_base, raio_metros=raio_sp, agrupar_por_levantador=True)
        dfs_fundidos.append(df_base_f)
    df_tasks = pd.concat(dfs_fundidos, ignore_index=True)

    if trava_global > 0: df_tasks = df_tasks.head(trava_global)
    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])

    df_ta = df_tasks.copy()
    df_u = pd.DataFrame()

    st.session_state.df_unallocated_lista, total_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    
    sb_html.markdown(render_sidebar_card(total_alocadas, qtd_eq), unsafe_allow_html=True)

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = ['ID SISCO', 'FASE', 'PRIORIDADE', 'TIPO NOTA', 'PROTOCOLO', 'CONTA CONTRATO', 'INSTALACAO', 'NOME', 'ENDERECO', 'LATITUDE', 'LONGITUDE', 'LOCALIDADE', 'MUNICIPIO', 'INFORMACOES EXTRAS', 'DISTANCIA BT', 'DISTANCIA MT', 'DISTANCIA TRAFO', 'POSTE PREVISTO BT', 'POSTE PREVISTO MT']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        st.session_state.update({'colunas_exibir_lista': colunas_exibir})
        st.session_state.vrp_state_lista = {'config': {'velocidade_media_kmh': 30.0, 'sentido_rota': sentido_rota, 'url_osrm_base': url_osrm, 'tracado_real': usa_osrm, 'data_inicio': data_ini, 'obras_por_dia_est': obras_por_dia_est}, 'b_names': list(set(df_ta['BASE_ATRIBUIDA'].unique())), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': [], 'current_geoms': []}
        st.session_state.vrp_status_lista = "RUNNING"; tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor Contínuo")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_lista', time.time())
    if 'start_time_run_lista' not in st.session_state: st.session_state.start_time_run_lista = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_lista; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            bl = sum(float(x['LATITUDE']) for x in oe) / len(oe)
            bL = sum(float(x['LONGITUDE']) for x in oe) / len(oe)
            
            if "Varredura Reversa" in cfg.get('sentido_rota', "Lógica Padrão"):
                ot = []
                if oe:
                    max_idx = max(range(len(oe)), key=lambda i: haversine_scalar(bl, bL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                    p_longe = oe.pop(max_idx)
                    ot.append(p_longe)
                    cl, cL = float(p_longe['LATITUDE']), float(p_longe['LONGITUDE'])
                    while oe:
                        closest_idx = min(range(len(oe)), key=lambda i: haversine_scalar(cl, cL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                        nx = oe.pop(closest_idx)
                        ot.append(nx)
                        cl, cL = float(nx['LATITUDE']), float(nx['LONGITUDE'])
            else:
                ot = resolver_tsp_ortools(oe, bl, bL, cfg['url_osrm_base'] if cfg.get('tracado_real') else "") if oe else []
                if not ot: ot = oe
            
            rf = []
            c_l, c_L = bl, bL
            for o in ot:
                vkr = haversine_vectorized(c_l, c_L, o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                rf.append({'o': o, 'la': c_l, 'La': c_L, 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 'dk': vk})
                c_l, c_L = o['LATITUDE'], o['LONGITUDE']
                
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_lista = st_v; tentar_rerun(); st.stop()
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
                    time.sleep(0.15)
                    try: 
                        res = obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'])
                        if not res or len(res) == 0 or len(res[0]) == 0: gd.append(fallback)
                        else: gd.append(res)
                    except: gd.append(fallback)
                
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_lista = st_v; tentar_rerun(); st.stop()
            
            rdf, og = [], 1
            for it, (g, ds) in zip(rf, gd):
                ob = it['o']; ob['ORDEM'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, round(it['dk'], 2)
                ob['ROTA_GEOMETRIA'], ob['PERIODO'] = g, "Único"
                rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state_lista = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        st.session_state.df_routed_lista = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status_lista = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    
    df_routed = st.session_state.df_routed_lista.copy()
    df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = df_routed.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)

    d_fmt = datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx, bu_txt = io.BytesIO(), io.BytesIO(), io.BytesIO(), io.BytesIO()
    
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, \
             zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, \
             zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg, \
             zipfile.ZipFile(bu_txt, 'w', zipfile.ZIP_DEFLATED) as zt:
            
            data_ini = st.session_state.vrp_state_lista.get('config', {}).get('data_inicio', datetime.today())
            obras_por_dia_est = st.session_state.vrp_state_lista.get('config', {}).get('obras_por_dia_est', 4.0)
            
            if isinstance(data_ini, datetime): data_ini = data_ini.date()
            dia_mes_str = data_ini.strftime("%d/%m/%Y")
            dias_semana_pt = {0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA", 3: "QUINTA-FEIRA", 4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO"}
            dia_semana_str = dias_semana_pt[data_ini.weekday()]

            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                
                qtd_obras = 0
                qtd_postes = 0
                for _, r in db.iterrows():
                    if isinstance(r.get('_ORIGINAL_ROWS'), list):
                        qtd_obras += len(r['_ORIGINAL_ROWS'])
                        for orig in r['_ORIGINAL_ROWS']:
                            pv = []
                            for k, v in orig.items():
                                if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                    try: 
                                        val = float(v)
                                        if val > 0: pv.append(val)
                                    except: pass
                            if pv: qtd_postes += min(pv)
                    else:
                        qtd_obras += 1
                        pv = []
                        for k, v in r.items():
                            if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                try: 
                                    val = float(v)
                                    if val > 0: pv.append(val)
                                except: pass
                        if pv: qtd_postes += min(pv)

                postes_dia = (qtd_postes / (qtd_obras / float(obras_por_dia_est))) if qtd_obras > 0 else 0
                postes_semana = postes_dia * 5.0

                res.append({
                    'Equipe': b, 
                    'Obras Roteirizadas': qtd_obras, 
                    'Postes/Dia (Est.)': int(round(postes_dia)),
                    'Postes/Semana (Est.)': int(round(postes_semana)),
                    'Postes Total': int(round(qtd_postes)),
                    'Super Pontos': qs, 
                    'Prioridades Atendidas': len(db[db['PRIORIDADE']=='Sim']), 
                    'KM Total Previsto': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)
                })
            zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_lista(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_lista', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
            
            linhas_gerais = []
            for _, r in df_routed.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
                sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else "NÃO"
                
                if is_sp:
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = r.copy()
                        for k, v in orig.items(): 
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'ROTA_GEOMETRIA', 'PERIODO']: nr[k] = v
                        nr['DIA_SEMANA'] = dia_semana_str
                        nr['DIA_MES'] = dia_mes_str
                        nr['SUPER_PONTO'] = sp_text
                        linhas_gerais.append(nr)
                else:
                    rn = r.copy()
                    rn['DIA_SEMANA'] = dia_semana_str
                    rn['DIA_MES'] = dia_mes_str
                    rn['SUPER_PONTO'] = sp_text
                    linhas_gerais.append(rn)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            
            for c in df_excel_full.columns:
                if 'POSTE' in c.upper():
                    df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

            col_exibir = st.session_state.colunas_exibir_lista.copy()
            if 'DIA_SEMANA' not in col_exibir: col_exibir.insert(0, 'DIA_SEMANA')
            if 'DIA_MES' not in col_exibir: col_exibir.insert(1, 'DIA_MES')
            if 'SUPER_PONTO' not in col_exibir: col_exibir.insert(2, 'SUPER_PONTO')

            dfg_total = limpar_colunas_lista(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
            dfg_total = dfg_total.loc[:, ~dfg_total.columns.duplicated()].copy()
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.xlsx", gerar_excel_lista(dfg_total, st.session_state.colunas_originais_lista))
            
            # --- TXT TOTAL ---
            txt_total = gerar_txt_lista(df_excel_full)
            zt.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.txt", txt_total.encode('utf-8'))
            
            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                dfk_total['DIA_SEMANA'] = dia_semana_str
                dfk_total['DIA_MES'] = dia_mes_str
                dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                
                ks_tot = gerar_kml_lista(dfk_total, "ROTA TOTAL", col_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), formatar_valor_coluna)
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks_tot.encode('utf-8'))
                zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

            for base in df_routed['BASE_ATRIBUIDA'].unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == base]
                if not df_base_excel.empty:
                    dfg = limpar_colunas_lista(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
                    dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
                    for cc in dfg.columns:
                        if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                    zx.writestr(f"Rotas_{d_fmt}/Rota_{b_safe}.xlsx", gerar_excel_lista(dfg, st.session_state.colunas_originais_lista))
                    
                    # --- TXT INDIVIDUAL ---
                    txt_ind = gerar_txt_lista(df_base_excel)
                    if txt_ind:
                        zt.writestr(f"Relatorios_TXT_{d_fmt}/Relatorio_{b_safe}.txt", txt_ind.encode('utf-8'))
                    
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == base) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if not dfk_base.empty:
                    dfk_base['DIA_SEMANA'] = dia_semana_str
                    dfk_base['DIA_MES'] = dia_mes_str
                    dfk_base['SUPER_PONTO'] = dfk_base.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                    
                    ks = gerar_kml_lista(dfk_base, f"Rota {b_safe}", col_exibir, [base], formatar_valor_coluna)
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kml", ks.encode('utf-8'))
                    zg.writestr(f"GPX_{d_fmt}/Rota_{b_safe}.gpx", gerar_gpx_simples(dfk_base, f"Rota {b_safe}").encode('utf-8'))

        st.session_state.bytes_zip_xl_lista = bu_xl.getvalue()
        st.session_state.bytes_zip_kml_lista = bu_kml.getvalue()
        st.session_state.bytes_zip_gpx_lista = bu_gpx.getvalue()
        st.session_state.bytes_zip_txt_lista = bu_txt.getvalue()
        st.session_state.roteamento_concluido_lista = True; st.session_state.vrp_status_lista = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_lista = "IDLE"
