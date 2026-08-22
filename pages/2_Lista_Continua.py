import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime

# Importações do Backend
from modules.data_processing import ler_planilha_cached, formatar_moeda, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.export_utils import injetar_logo, identificar_icone_folium, gerar_excel_bytes, gerar_excel_resumo_bytes, gerar_gpx_simples, gerar_kml_agrupado

st.set_page_config(page_title="Lista Contínua", page_icon="♾️", layout="wide")
injetar_logo()

# ==========================================
# FUNÇÕES VISUAIS E AUXILIARES
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
    except:
        if isinstance(val, (datetime, pd.Timestamp)): return formata_campo_html(val.strftime('%d/%m/%Y'))
        return formata_campo_html(str(val))

def count_real_obras(row):
    if isinstance(row.get('_ORIGINAL_ROWS'), list): return len(row['_ORIGINAL_ROWS'])
    val = str(row.get('SUPER_PONTO', ''))
    if val.startswith('SIM'):
        nums = re.findall(r'\d+', val)
        if nums: return int(nums[0])
    return 1

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

def render_sidebar_card(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade):
    return f"""
    <div style="background-color: #ffffff; padding: 25px; border-radius: 10px; border: 1px solid #e0e0e0; box-shadow: 0 4px 8px rgba(0,0,0,0.05); margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 18px; border-bottom: 2px solid #55B929; padding-bottom: 10px; margin-bottom: 15px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Cota p/ Equipe:</b> <span style="color: #d9534f; font-weight: bold;">{limite_por_equipe}</span> obras</p>
        <p style="margin-bottom: 15px; font-size: 15px;"><b>Capacidade Total:</b> <span style="color: #55B929; font-weight: bold;">{total_capacidade}</span> obras</p>
        <hr style="margin: 15px 0; border: 0; border-top: 1px solid #eee;">
        <div style="text-align: center; margin-top: 15px;">
            <p style="margin-bottom: 5px; font-size: 16px; color: #555;"><b>Obras Validadas:</b></p>
            <span style="font-size: 36px; color: #0D256C; font-weight: 900;">{total_obras_prontas}</span>
        </div>
    </div>
    """

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_lc': False, 'vrp_status_lc': "IDLE", 'vrp_state_lc': {}, 'df_routed_lc': pd.DataFrame(), 'bases_records_lc': [], 'tipo_periodo_lc': "Dia", 'colunas_exibir_lc': [], 'colunas_originais_lc': []})
    for k in ['bytes_zip_xl_lc', 'bytes_zip_kml_lc', 'bytes_zip_gpx_lc', 'start_time_run_lc', 'start_time_pkg_lc']: st.session_state.pop(k, None)
    ler_planilha_cached.clear(); tentar_rerun()


# ==========================================
# INÍCIO DA PÁGINA
# ==========================================
if "roteamento_concluido_lc" not in st.session_state: st.session_state.roteamento_concluido_lc = False
if "vrp_status_lc" not in st.session_state: st.session_state.vrp_status_lc = "IDLE"

status_exec = st.session_state.vrp_status_lc
is_done = st.session_state.roteamento_concluido_lc
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>♾️ Lista Contínua (Técnico Fixo)</h1>", unsafe_allow_html=True)
st.info("💡 **Como funciona:** O sistema respeita estritamente a coluna 'LEVANTADOR' da sua planilha. A IA ignora o limite de dias e roteiriza 100% da lista de cada técnico em sequência contínua.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Esforço e Limites", expanded=True):
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão (Mais Próximo)", "🎯 Varredura Reversa (Longe -> Perto)"], index=0, disabled=is_locked)
        raio_super_ponto = st.slider("Raio Super Ponto (Metros)", 10, 1000, 100, 10, disabled=is_locked)
        st.markdown("---")
        tipo_periodo = st.radio("Agrupamento:", ["☀️ Dia", "📅 Semana"], index=1, horizontal=True, disabled=is_locked)
        tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
        dias_semana_selecionados = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
        if tipo_periodo_clean == "Semana":
            dias_semana_selecionados = st.multiselect("Dias úteis:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked)
        
        data_inicio_roteiro = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        obras_por_dia = st.number_input("Obras Previstas por Dia", min_value=1, value=30, step=1, disabled=is_locked)
        velocidade_media_kmh = 30.0

    st.markdown("---")
    sidebar_html_placeholder = st.empty()
    
    if is_done and not st.session_state.df_routed_lc.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_lc', b"vazio"), file_name=f"Continua_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_lc', b"vazio"), file_name=f"Continua_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_lc', b"vazio"), file_name=f"Continua_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

# ==========================================
# EXIBIÇÃO DE RESULTADOS (SE CONCLUÍDO)
# ==========================================
if is_done and not st.session_state.df_routed_lc.empty:
    st.markdown("## 🎯 Resultados da Otimização Contínua")
    st.session_state.df_routed_lc['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_lc.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    df_routed = st.session_state.df_routed_lc.copy()
    df_real_tasks = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tot_paradas = len(df_real_tasks)
    tot_obras_reais = sum(count_real_obras(r) for _, r in df_real_tasks.iterrows())
    tot_equipes = df_routed['BASE_ATRIBUIDA'].nunique()
    tot_km = f"{df_routed['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    tot_super_pontos = len(df_real_tasks[df_real_tasks['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in df_real_tasks.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("TOTAL DE OBRAS", f"{tot_obras_reais} <span style='font-size:12px;color:#888;'>(Em {tot_paradas} Pontos)</span>", "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", tot_equipes, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("KM Total Projetado", tot_km, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("Pontos Agrupados", tot_super_pontos, "🏢", "#eab308", "rgba(234,179,8,0.15)"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<div style="background-color: #d4edda; color: #155724; padding: 15px; border-left: 5px solid #c3e6cb; margin-bottom: 20px; border-radius: 4px;"><h4 style="margin-top: 0; margin-bottom: 5px;">✅ Processamento Contínuo Concluído!</h4><p style="margin: 0;">100% da planilha compatível (<b>{tot_obras_reais} obras</b>) foi roteirizada para os levantadores já definidos no seu arquivo Excel.</p></div>', unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Geográfico")
    mapa = folium.Map(location=[df_routed['LATITUDE'].mean(), df_routed['LONGITUDE'].mean()], zoom_start=8) if not df_routed.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Agrupamento de Obras").add_to(mapa)
    for bn in df_routed['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(df_routed['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = df_routed[df_routed['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        for pe in db['PERIODO'].unique():
            dp = db[db['PERIODO'] == pe]
            pts = [p for _, r in dp.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
            folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
            folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
            
            for r in dp.to_dict('records'):
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                ic = identificar_icone_folium(r, df_routed.columns)
                if str(r.get('SUPER_PONTO', '')).startswith('SIM'):
                    ci, p_txt, p_bg, p_c = 'orange', f"🏢 SUPER PONTO {str(r.get('SUPER_PONTO')).replace('SIM','').strip()}", "#FFD700", "#000000"
                else:
                    ci, p_txt, p_bg, p_c = ('red', "🚨 OBRA PRIORITÁRIA", "#d9534f", "#ffffff") if r.get('PRIORIDADE')=='Sim' else ('blue', "📍 Atendimento Padrão", "#0D256C", "#ffffff")
                er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}:</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_lc if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA','COR_ICONE']])
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=ci, icon=ic), popup=folium.Popup(f'<div style="width:280px;"><div style="background:{p_bg};color:{p_c};padding:8px;">{p_txt}</div><table border="1" style="width:100%;font-size:12px;"><tr><td style="padding:3px;"><b>Ordem:</b></td><td style="padding:3px;">{r.get("ORDEM",0)}</td></tr>{er}</table></div>', max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    t1, t2 = st.tabs(["📊 Dados Tabulares", "📉 Resumo por Técnico"])
    with t1:
        st.data_editor(st.session_state.df_routed_lc.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'BASE_ATRIBUIDA', 'COR_ICONE'], errors='ignore'), use_container_width=True)
    with t2:
        dr = pd.DataFrame([{"Levantador": b['LEVANTADOR'], "Obras Roteirizadas": sum(count_real_obras(r) for _, r in df_real_tasks[df_real_tasks['BASE_ATRIBUIDA']==b['LEVANTADOR']].iterrows())} for b in st.session_state.bases_records_lc]).reset_index(drop=True)
        st.dataframe(dr, use_container_width=True)

# ==========================================
# START DA APLICAÇÃO (UPLOAD DE DADOS)
# ==========================================
elif status_exec == "IDLE":
    st.markdown("### 📥 1. Planilha de Demanda")
    with st.container(border=True): p_f = st.file_uploader("Upload da Planilha Obras", type=["xlsx", "xls", "csv"], help="A planilha DEVE conter as colunas LEVANTADOR, MUNICIPIO, LATITUDE e LONGITUDE.")
    ig_d = st.checkbox("Filtro: Ignorar obras despachadas?", value=False)
    
    if p_f:
        df_tasks = ler_planilha_cached(p_f.getvalue()) if not p_f.name.endswith('.csv') else pd.read_csv(p_f)
        df_tasks.columns = normalize_cols(df_tasks.columns)
        st.session_state.colunas_originais_lc = df_tasks.columns.tolist()
        df_tasks['_ORIGEM_BASE'] = 'LISTA_CONTINUA'
        if 'PRIORIDADE' in df_tasks.columns: df_tasks['_PRIORIDADE_ORIGINAL'] = df_tasks['PRIORIDADE']
        
        falta = [c for c in ['LEVANTADOR', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE'] if c not in df_tasks.columns and c not in ['NOME DO LEVANTADOR', 'CIDADE']]
        if falta: st.error(f"🚨 A planilha precisa das colunas: {', '.join(falta)}."); st.stop()
        
        if 'NOME DO LEVANTADOR' in df_tasks.columns and 'LEVANTADOR' not in df_tasks.columns: df_tasks.rename(columns={'NOME DO LEVANTADOR': 'LEVANTADOR'}, inplace=True)
        if 'CIDADE' in df_tasks.columns and 'MUNICIPIO' not in df_tasks.columns: df_tasks.rename(columns={'CIDADE': 'MUNICIPIO'}, inplace=True)
            
        if 'PROTOCOLO' not in df_tasks.columns:
            for cc in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
                if cc in df_tasks.columns: df_tasks['PROTOCOLO'] = df_tasks[cc]; break
            if 'PROTOCOLO' not in df_tasks.columns: df_tasks['PROTOCOLO'] = [f"LC-{i+1}" for i in range(len(df_tasks))]
        
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()
        for cc in ['CONTA CONTRATO', 'INSTALACAO', 'PROTOCOLO']:
            if cc in df_tasks.columns: df_tasks[cc] = df_tasks[cc].astype(str).replace(r'\.0$', '', regex=True).replace('nan', '-')
            
        if ig_d and 'DATA DESPACHO CAMPO' in df_tasks.columns:
            md = df_tasks['DATA DESPACHO CAMPO'].notna() & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip() != '') & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip().str.lower() != 'nan')
            if md.sum() > 0: st.info(f"⏭️ {md.sum()} ignoradas (Despachadas)."); df_tasks = df_tasks[~md]
            
        if 'STATUS SAP' in df_tasks.columns:
            ms = df_tasks['STATUS SAP'].astype(str).str.strip().str.upper().isin(['CANC', 'FINL'])
            if ms.sum() > 0: st.info(f"🗑️ {ms.sum()} ignoradas (CANC/FINL)."); df_tasks = df_tasks[~ms]

        df_tasks['LATITUDE'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df_tasks['LONGITUDE'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        ec = df_tasks['LATITUDE'].isna() | df_tasks['LONGITUDE'].isna() | (df_tasks['LATITUDE'] == 0.0) | (df_tasks['LONGITUDE'] == 0.0)
        df_tasks = df_tasks[~ec]
        
        if 'NOME' not in df_tasks.columns: df_tasks['NOME'] = "SEM NOME"
        df_tasks['LEVANTADOR'] = df_tasks['LEVANTADOR'].astype(str).str.strip().str.upper()
        df_tasks = df_tasks[~df_tasks['LEVANTADOR'].isin(['NAN', 'NONE', '', '-', 'SEM LEVANTADOR', '0', '0.0', 'N/A', 'NULO'])]
        
        df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_super_ponto, agrupar_por_levantador=True)
        
        if 'PRIORIDADE' not in df_tasks.columns: df_tasks['PRIORIDADE'] = 'Não'
        else: df_tasks['PRIORIDADE'] = df_tasks['PRIORIDADE'].astype(str).str.strip().str.upper().apply(lambda x: 'Sim' if x == 'SIM' else 'Não')
        if 'STATUS LIST' in df_tasks.columns: df_tasks.loc[df_tasks['STATUS LIST'].astype(str).str.strip().str.upper().isin(['CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO']), 'PRIORIDADE'] = 'Sim'
        if '_PRIORIDADE_ORIGINAL' in df_tasks.columns:
            m_nz = df_tasks['_PRIORIDADE_ORIGINAL'].notna() & (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '') & (~df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper().isin(['0','0.0','NAN','NÃO','NAO','FALSE']))
            df_tasks.loc[m_nz, 'PRIORIDADE'] = 'Sim'

        if df_tasks.empty: st.error("🚨 Nenhuma obra restou na base."); st.stop()
        
        df_tasks['BASE_ATRIBUIDA'] = df_tasks['LEVANTADOR']; df_tasks_alocadas = df_tasks.copy()
        bases_records = []
        for lev in df_tasks_alocadas['LEVANTADOR'].unique():
            dl = df_tasks_alocadas[df_tasks_alocadas['LEVANTADOR'] == lev]
            mb = dl['MUNICIPIO'].mode().iloc[0] if not dl['MUNICIPIO'].dropna().empty else "DESCONHECIDO"
            rb = dl['REGIONAL'].iloc[0] if 'REGIONAL' in dl.columns else "DESCONHECIDO"
            lt, ln = obter_coordenadas_municipio_cached(mb)
            if pd.isna(lt) or pd.isna(ln): lt, ln = dl['LATITUDE'].iloc[0], dl['LONGITUDE'].iloc[0]
            bases_records.append({'LEVANTADOR': lev, 'RESIDENCIA': mb, 'MUNICIPIO': mb, 'REGIONAL': rb, 'LATITUDE': lt, 'LONGITUDE': ln, 'TIPO_EQUIPE': 'LISTA_CONTINUA'})
        
        tot_obras_prontas = sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_tasks_alocadas.iterrows())
        sidebar_html_placeholder.markdown(render_sidebar_card("Ilimitado", tot_obras_prontas, len(bases_records), "Ilimitado"), unsafe_allow_html=True)
        st.success(f"✅ Planilha validadas! {len(df_tasks_alocadas)} paradas encontradas para {len(bases_records)} técnicos.")

        with st.expander("🛠️ Configuração de Saída", expanded=True):
            tc = [c for c in df_tasks_alocadas.columns if not c.startswith('_') and c != 'COR_ICONE']
            cd = ['PROTOCOLO', 'CONTA CONTRATO', 'INSTALACAO', 'NOME', 'ENDERECO', 'INFORMACOES EXTRAS', 'LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'TIPO NOTA', 'FASE']
            cp = [c for c in cd if c in tc]
            colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
            ro = cd + [c for c in tc if c not in cd]; colunas_exibir.sort(key=lambda x: ro.index(x) if x in ro else 999)

        if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
            if tipo_periodo_clean == "Semana" and not dias_semana_selecionados: st.error("Selecione os dias da semana na barra lateral."); st.stop()
            st.session_state.update({'bases_records_lc': bases_records, 'tipo_periodo_lc': tipo_periodo_clean, 'colunas_exibir_lc': colunas_exibir, 'col_prioridade': "PRIORIDADE"})
            st.session_state.vrp_state_lc = {'config': {'obras_por_dia': obras_por_dia, 'tipo_periodo': tipo_periodo_clean, 'dias_selecionados': dias_semana_selecionados, 'sentido_rota': sentido_rota}, 'b_names': list(set([b['LEVANTADOR'] for b in bases_records])), 'b_idx': 0, 'unvisited': df_tasks_alocadas.copy(), 'routed_data': []}
            st.session_state.vrp_status_lc = "RUNNING"; tentar_rerun()

# ==========================================
# CÁLCULO VRP
# ==========================================
if status_exec == "RUNNING":
    state = st.session_state.vrp_state_lc
    cfg = state['config']
    is_reversa = "Reversa" in cfg.get('sentido_rota', '')
    
    st.markdown("## 🚀 Motor Sequencial Rápido")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
        
    b_names = state['b_names']
    b_idx = state.get('b_idx', 0)
    
    st_run = st.session_state.get('start_time_run_lc', time.time())
    if 'start_time_run_lc' not in st.session_state: st.session_state.start_time_run_lc = st_run
    
    pb = st.progress(0.0)
    tmp = st.empty()
    sgt = st.empty()

    if b_idx < len(b_names):
        b_name = b_names[b_idx]
        pb.progress(b_idx / max(1, len(b_names)))
        sgt.info(f"🧠 Sequenciando obras de **{b_name}**... ({b_idx + 1}/{len(b_names)})")
        
        df_todas_bases = pd.DataFrame(st.session_state.bases_records_lc)
        base_ref = df_todas_bases[df_todas_bases['LEVANTADOR'] == b_name].iloc[0]
        base_lat, base_lon = float(base_ref['LATITUDE']), float(base_ref['LONGITUDE'])
        obras_equipe = state['unvisited'][state['unvisited']['BASE_ATRIBUIDA'] == b_name].to_dict('records')
        
        ordered_tasks = []
        if obras_equipe:
            mg = {}
            for o in obras_equipe:
                ml = normalizar_municipios(pd.Series([o.get('MUNICIPIO', 'DESCONHECIDO')])).iloc[0]
                o['MUN_LIMPO_CALC'] = ml
                if ml not in mg: mg[ml] = []
                mg[ml].append(o)
                
            for mun, obs in mg.items():
                def greedy_sort(pts, s_lat, s_lon, rev=False):
                    if not pts: return []
                    s_pts = []; cl, cL = s_lat, s_lon; up = list(pts)
                    if rev and up:
                        bi, bd = 0, -1
                        for i, p in enumerate(up):
                            d = haversine_scalar(cl, cL, p['LATITUDE'], p['LONGITUDE'])
                            if d > bd: bd, bi = d, i
                        nx = up.pop(bi); s_pts.append(nx); cl, cL = nx['LATITUDE'], nx['LONGITUDE']
                    while up:
                        bi, bd = 0, float('inf')
                        for i, p in enumerate(up):
                            d = haversine_scalar(cl, cL, p['LATITUDE'], p['LONGITUDE'])
                            if d < bd: bd, bi = d, i
                        nx = up.pop(bi); s_pts.append(nx); cl, cL = nx['LATITUDE'], nx['LONGITUDE']
                    return s_pts

                ordered_tasks.extend(greedy_sort(obs, base_lat, base_lon, is_reversa))
        
        routed_data_final, dia_abs, o_hoje = [], 1, 0
        cl, cL = base_lat, base_lon
        
        for o in ordered_tasks:
            qr = len(o.get('_ORIGINAL_ROWS', [1])) if isinstance(o.get('_ORIGINAL_ROWS'), list) else 1
            if o_hoje > 0 and (o_hoje + qr > cfg['obras_por_dia']):
                routed_data_final.append({'PROTOCOLO': 'RETORNO_BASE', 'NOME': 'BASE_RETORNO', 'LATITUDE': base_lat, 'LONGITUDE': base_lon, 'BASE_ATRIBUIDA': b_name, 'ORDEM': len(routed_data_final)+1, 'NOME_DIA': f"Dia {dia_abs}", 'DIA_MES': "", 'SEMANA': 1, 'DIA': dia_abs, 'PERIODO': dia_abs, 'DISTANCIA_PONTO_ANTERIOR_KM': haversine_scalar(cl, cL, base_lat, base_lon), 'PRIORIDADE': 'Não'})
                dia_abs += 1; o_hoje = 0; cl, cL = base_lat, base_lon
            
            d_km = haversine_scalar(cl, cL, o['LATITUDE'], o['LONGITUDE'])
            o['ORDEM'], o['DIA'], o['PERIODO'], o['NOME_DIA'], o['DIA_MES'] = len(routed_data_final) + 1, dia_abs, dia_abs, f"Dia {dia_abs}", ""
            o['DISTANCIA_PONTO_ANTERIOR_KM'] = round(d_km, 2)
            routed_data_final.append(o)
            cl, cL = o['LATITUDE'], o['LONGITUDE']; o_hoje += qr
            
        if o_hoje > 0:
            routed_data_final.append({'PROTOCOLO': 'RETORNO_BASE', 'NOME': 'BASE_RETORNO', 'LATITUDE': base_lat, 'LONGITUDE': base_lon, 'BASE_ATRIBUIDA': b_name, 'ORDEM': len(routed_data_final)+1, 'NOME_DIA': f"Dia {dia_abs}", 'DIA_MES': "", 'SEMANA': 1, 'DIA': dia_abs, 'PERIODO': dia_abs, 'DISTANCIA_PONTO_ANTERIOR_KM': haversine_scalar(cl, cL, base_lat, base_lon), 'PRIORIDADE': 'Não'})

        state['routed_data'].extend(routed_data_final)
        state['b_idx'] += 1; st.session_state.vrp_state_lc = state; tentar_rerun()
    else:
        sgt.success("✅ Roteamento Concluído! Preparando empacotamento...")
        pb.progress(1.0)
        st.session_state.df_routed_lc = pd.DataFrame(state['routed_data'])
        st.session_state.vrp_status_lc = "PACKAGING"; time.sleep(1); tentar_rerun()

# ==========================================
# PACOTES E DOWNLOAD
# ==========================================
if status_exec == "PACKAGING":
    st.markdown("## 📦 Etapa Final: Construção de Arquivos (Excel e KML)")
    df_routed = st.session_state.df_routed_lc
    d_fmt = datetime.now().strftime("%d.%m.%Y")
    buf_xl, buf_kml, buf_gpx = io.BytesIO(), io.BytesIO(), io.BytesIO()
    
    try:
        with zipfile.ZipFile(buf_xl, 'w', zipfile.ZIP_DEFLATED) as z_xl, zipfile.ZipFile(buf_kml, 'w', zipfile.ZIP_DEFLATED) as z_kml, zipfile.ZipFile(buf_gpx, 'w', zipfile.ZIP_DEFLATED) as z_gpx:
            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                dfb = df_routed[(df_routed['BASE_ATRIBUIDA'] == b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                res.append({'LEVANTADOR': b, 'TOTAL OBRAS': len(dfb), 'KM TOTAL PREVISTO': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)})
            z_xl.writestr(f"Resumo_Continua - {d_fmt}.xlsx", gerar_excel_resumo_bytes(pd.DataFrame(res)))
            
            dfg = limpar_colunas_excel(df_routed.drop(columns=['MUN_LIMPO', 'COR_ICONE'], errors='ignore'), st.session_state.colunas_originais_lc)
            z_xl.writestr(f"Demanda_Continua - {d_fmt}.xlsx", gerar_excel_bytes(dfg, "PRIORIDADE"))
            
            dfk = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            k_str = gerar_kml_agrupado(dfk, st.session_state.bases_records_lc, f"ROTA_TOTAL - {d_fmt}", st.session_state.colunas_exibir_lc, df_routed['BASE_ATRIBUIDA'].unique().tolist(), st.session_state.tipo_periodo_lc, formatar_valor_coluna)
            z_kml.writestr(f"ROTA_TOTAL - {d_fmt}.kml", re.sub(r'<Placemark>(?:(?!</Placemark>).)*?<name>(?:(?!</name>).)*?BASE:(?:(?!</name>).)*?</name>(?:(?!</Placemark>).)*?</Placemark>', '', k_str, flags=re.IGNORECASE | re.DOTALL).encode('utf-8'))
            z_gpx.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk, "ROTA TOTAL").encode('utf-8'))
            
            for b_name in df_routed['BASE_ATRIBUIDA'].unique():
                ns = re.sub(r'[^A-Za-z0-9_ ]', '', str(b_name)).replace(" ", "_").upper()
                dfl = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
                dflk = dfl[~dfl['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
                dx = limpar_colunas_excel(dfl.drop(columns=['MUN_LIMPO'], errors='ignore'), st.session_state.colunas_originais_lc)
                for c in dx.columns:
                    if str(dx[c].dtype) == 'object': dx[c] = dx[c].astype(str).replace('nan', '')
                z_xl.writestr(f"ROTA_{ns} - {d_fmt}.xlsx", gerar_excel_bytes(dx, "PRIORIDADE"))
                kl = gerar_kml_agrupado(dflk, st.session_state.bases_records_lc, f"ROTA_{ns}", st.session_state.colunas_exibir_lc, [b_name], st.session_state.tipo_periodo_lc, formatar_valor_coluna)
                z_kml.writestr(f"ROTA_{ns} - {d_fmt}.kml", re.sub(r'<Placemark>(?:(?!</Placemark>).)*?<name>(?:(?!</name>).)*?BASE:(?:(?!</name>).)*?</name>(?:(?!</Placemark>).)*?</Placemark>', '', kl, flags=re.IGNORECASE | re.DOTALL).encode('utf-8'))
                z_gpx.writestr(f"GPS_{ns} - {d_fmt}.gpx", gerar_gpx_simples(dflk, f"ROTA_{ns}").encode('utf-8'))

        st.session_state.bytes_zip_xl_lc, st.session_state.bytes_zip_kml_lc, st.session_state.bytes_zip_gpx_lc = buf_xl.getvalue(), buf_kml.getvalue(), buf_gpx.getvalue()
        st.session_state.roteamento_concluido_lc = True; st.session_state.vrp_status_lc = "IDLE"; tentar_rerun()
    except Exception as e:
        st.error(f"🚨 ERRO NO EMPACOTAMENTO: {e}"); st.session_state.vrp_status_lc = "IDLE"
