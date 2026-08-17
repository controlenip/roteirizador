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
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import os
import base64
import gc

# ==========================================
# 1. CONFIGURAÇÕES INICIAIS DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Roteirizador NIP v3.0",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. IMPORTAÇÕES DOS MÓDULOS DIVIDIDOS
# ==========================================
from modules.data_processing import (
    ler_planilha_cached, formatar_moeda, formata_campo_html, 
    normalize_cols, normalizar_municipios, atualizar_status_via_df
)
from modules.geospatial import (
    haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, 
    resgatar_coordenadas, fundir_super_pontos
)
from modules.routing_engine import (
    resolver_tsp_ortools, obter_rota_ruas, calcular_matriz_distancias_numpy
)
from modules.export_utils import (
    identificar_icone_folium, renderizar_painel_lateral, 
    gerar_excel_bytes, gerar_excel_resumo_bytes, gerar_kml_agrupado
)

LOGO_PATH = "assets/LOGO_NIP.png"
STATUS_PADRAO = ['EM LEVANTAMENTO', '0', 'SEM INFORMAÇÕES', 'SEM INFORMACOES', 'CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO', 'PRÉ ANÁLISE', 'PRE ANALISE']
TIPOS_PRIORITARIOS = ["CCF", "DIF", "MGD", "MTP", "ASC", "SID"]

try:
    with open("assets/style.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
except FileNotFoundError:
    pass 

def tentar_rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.roteamento_concluido = False
    st.session_state.vrp_status = "IDLE"
    st.session_state.vrp_state = {}
    st.session_state.df_routed = pd.DataFrame()
    st.session_state.bases_records = []
    st.session_state.tipo_periodo = "Dia"
    st.session_state.colunas_exibir = []
    st.session_state.col_prioridade = "TIPO NOTA"
    st.session_state.colunas_originais = []
    
    # Resetando as variáveis de tempo para o relógio
    keys_to_clear = ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg', 'tempo_processamento', 'df_unallocated']
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
            
    ler_planilha_cached.clear()
    tentar_rerun()

# ==========================================
# FUNÇÕES AUXILIARES DE ROTEAMENTO
# ==========================================
def formatar_valor_coluna(col_name, val):
    if pd.isna(val) or val == '' or val == '-':
        return '-'
    try:
        val_float = float(val)
        if 'DISTANCIA' in col_name.upper():
            return f"{val_float:.2f} Metros"
        elif 'POSTE PREVISTO' in col_name.upper() or 'POSTES PREVISTOS' in col_name.upper():
            return f"{int(round(val_float))}"
        else:
            return formata_campo_html(val)
    except ValueError:
        return formata_campo_html(val)

def count_real_obras(row):
    if isinstance(row.get('_ORIGINAL_ROWS'), list):
        return len(row['_ORIGINAL_ROWS'])
    val = str(row.get('SUPER_PONTO', ''))
    if val.startswith('SIM'):
        nums = re.findall(r'\d+', val)
        if nums: return int(nums[0])
    return 1

def limpar_colunas_excel(df_alvo, cols_originais):
    base_start = ['LEVANTADOR_RESPONSAVEL', 'ORDEM', 'NOME_DIA', 'DIA_MES', 'PRIORIDADE', 'SUPER_PONTO']
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'LEVANTADOR_RESPONSAVEL'})
    elif 'LEVANTADOR' in df_alvo.columns and 'LEVANTADOR_RESPONSAVEL' not in df_alvo.columns:
        df_alvo['LEVANTADOR_RESPONSAVEL'] = df_alvo['LEVANTADOR']
        
    base_end = ['LINK_NAVEGACAO_OFFLINE']
    middle_cols = [c for c in cols_originais if c in df_alvo.columns and c not in base_start and c not in base_end]
    
    final_cols = []
    for c in base_start + middle_cols + base_end:
        if c in df_alvo.columns and c not in final_cols:
            final_cols.append(c)
            
    return df_alvo[final_cols]

def gerar_gpx_simples(df_kml, nome_rota):
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>']
    gpx.append('<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">')
    gpx.append(f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>')
    
    for idx, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat = row.get('LATITUDE')
        lon = row.get('LONGITUDE')
        nome = str(row.get('PROTOCOLO', 'Ponto'))
        if pd.notna(lat) and pd.notna(lon):
            gpx.append(f'  <wpt lat="{lat}" lon="{lon}">')
            gpx.append(f'    <name>{html.escape(nome)}</name>')
            gpx.append(f'  </wpt>')
            
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append('  <trk>')
        gpx.append(f'    <name>Traçado - {html.escape(str(nome_rota))}</name>')
        gpx.append('    <trkseg>')
        for idx, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom:
                    gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg>')
        gpx.append('  </trk>')
        
    gpx.append('</gpx>')
    return "\n".join(gpx)

# ==========================================
# 3. LÓGICA DO ROTEIRIZADOR (MOTOR PRINCIPAL)
# ==========================================
def app_roteirizador():
    if "roteamento_concluido" not in st.session_state: st.session_state.roteamento_concluido = False
    if "vrp_status" not in st.session_state: st.session_state.vrp_status = "IDLE"
    if "vrp_state" not in st.session_state: st.session_state.vrp_state = {}
    if "df_routed" not in st.session_state: st.session_state.df_routed = pd.DataFrame()
    if "bases_records" not in st.session_state: st.session_state.bases_records = []
    if "colunas_exibir" not in st.session_state: st.session_state.colunas_exibir = []
    if "col_prioridade" not in st.session_state: st.session_state.col_prioridade = "TIPO NOTA"
    if "colunas_originais" not in st.session_state: st.session_state.colunas_originais = []
    if "config_financeira" not in st.session_state: st.session_state.config_financeira = {}
    if "cache_coords" not in st.session_state: st.session_state.cache_coords = {}

    status_exec = st.session_state.vrp_status
    is_done = st.session_state.roteamento_concluido

    st.markdown("<h1 class='brand-title'>Plataforma Roteirizadora NIP v3.0</h1>", unsafe_allow_html=True)

    s1_class = "step-item done" if (status_exec != "IDLE" or is_done) else "step-item active"
    s2_class = "step-item done" if (status_exec != "IDLE" or is_done) else "step-item active"
    s3_class = "step-item active" if status_exec in ["RUNNING", "PACKAGING"] else ("step-item done" if is_done else "step-item")
    s4_class = "step-item active" if is_done else "step-item"
    
    st.markdown(f"""
    <div class="stepper-container">
        <div class="{s1_class}">📁 1. Dados e Profiling</div>
        <div class="{s2_class}">⚙️ 2. Filtros Dinâmicos</div>
        <div class="{s3_class}">🚀 3. IA VRP OR-Tools</div>
        <div class="{s4_class}">🎯 4. Resultados e Custos</div>
    </div>
    """, unsafe_allow_html=True)

    is_locked = status_exec != "IDLE" or is_done
    
    with st.sidebar:
        with st.expander("⚙️ Esforço e Limites Diários", expanded=True):
            trava_global_obras = st.number_input("Trava Total de Operação (Bolo Geral da Empresa)", min_value=0, value=0, step=50, disabled=is_locked)
            st.caption("⚠️ **Aviso:** Se o valor for **0**, o sistema levará em consideração **todas** as obras compatíveis. Se ficar em qualquer outro valor, apenas esta exata quantidade será roteirizada.")
            
            sentido_rota = st.radio("Sentido do Roteamento Diário:", ["📍 Lógica Padrão (Mais Próximo Primeiro)", "🎯 Varredura Reversa (Mais Distante Primeiro)"], index=0, disabled=is_locked)
            raio_super_ponto = st.slider("Raio do Super Ponto (Metros)", min_value=10, max_value=1000, value=100, step=10, disabled=is_locked, help="Agrupa obras que estiverem dentro desta distância em um único pino.")
            
            st.markdown("---")
            modo_produtividade = st.checkbox("🔥 Focar em Alta Densidade (Produtividade)", value=False, disabled=is_locked)
            st.caption("⚠️ **Aviso:** Se ativado, a IA rejeita obras isoladas e foca a equipe apenas em 'bolsões' com muitas notas.")
            if modo_produtividade:
                min_vizinhos = st.slider("Mínimo de obras próximas (Raio 2km):", min_value=2, max_value=50, value=10, step=1, disabled=is_locked)
            else:
                min_vizinhos = 0
            st.markdown("---")

            tipo_periodo = st.radio("Agrupamento de percurso:", ["☀️ Dia", "📅 Semana"], index=1, horizontal=True, disabled=is_locked)
            tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
            
            if tipo_periodo_clean == "Dia":
                st.caption("A IA encerra o roteiro no fim do dia.")
            else:
                st.caption("A IA monta jornadas contínuas de segunda a sexta.")
            
            dias_semana_selecionados = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
            if tipo_periodo_clean == "Semana":
                dias_semana_selecionados = st.multiselect(
                    "Dias úteis na semana:",
                    ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"],
                    default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"],
                    disabled=is_locked
                )
                if not dias_semana_selecionados:
                    st.warning("⚠️ Selecione pelo menos 1 dia da semana para o cálculo.")
                else:
                    st.caption(f"ℹ️ Cada semana terá **{len(dias_semana_selecionados)} dias** alocados.")
            
            data_inicio_roteiro = st.date_input("📅 Data de Início do Roteiro:", value=datetime.today(), disabled=is_locked)
            obras_por_dia = st.number_input("Obras Previstas por Dia", min_value=1, value=30, step=1, disabled=is_locked)
            limite_periodos = st.number_input(f"Limite total de {tipo_periodo_clean}s", min_value=1, value=5, step=1, disabled=is_locked)
            tempo_medio_obra = 1.5
            velocidade_media_kmh = 30.0

        with st.expander("💰 Custos e Gestão Financeira", expanded=False):
            custo_combustivel = st.number_input("Custo Combustível (R$/L)", min_value=0.0, value=0.0, step=0.1, disabled=is_locked)
            consumo_veiculo = st.number_input("Consumo Frota (Km/L)", min_value=0.0, value=0.0, step=0.5, disabled=is_locked)
            custo_hora_equipe = st.number_input("Hora-Homem da Equipe (R$)", min_value=0.0, value=0.0, step=1.0, disabled=is_locked)
            
        with st.expander("📡 Conexão de Rede (Avançado)", expanded=False):
            url_osrm_base = st.text_input("Endpoint OSRM ⚠️ (NÃO APAGUE OU EDITE):", value="http://router.project-osrm.org", disabled=is_locked)
            st.caption("Este link conecta o sistema à malha viária real de ruas do mundo.")
            tracado_real = st.checkbox("🛣️ Desenhar Traçado de Ruas no Mapa (Lento)", value=False, disabled=is_locked, help="Desmarque para usar o modo Vetorial Rápido (Altamente recomendado para bases grandes).")
            
        st.markdown("---")
        sidebar_html_placeholder = st.empty()
        
        st.markdown("### 📥 Ações e Arquivos")
        data_atual_formatada = datetime.now().strftime("%d.%m.%Y")
        bytes_zip_xl = st.session_state.get('bytes_zip_xl', b"")
        bytes_zip_kml = st.session_state.get('bytes_zip_kml', b"")
        bytes_zip_gpx = st.session_state.get('bytes_zip_gpx', b"")
        
        botoes_desabilitados = not is_done or st.session_state.df_routed.empty
        
        st.download_button("🌐 1. Baixar Planilhas (ZIP)", data=bytes_zip_xl if bytes_zip_xl else b"vazio", file_name=f"Planilhas_Equipes - {data_atual_formatada}.zip", mime="application/zip", use_container_width=True, disabled=botoes_desabilitados)
        st.download_button("🗺️ 2. Baixar Mapas (KML)", data=bytes_zip_kml if bytes_zip_kml else b"vazio", file_name=f"Mapas_Rotas - {data_atual_formatada}.zip", mime="application/zip", use_container_width=True, disabled=botoes_desabilitados)
        st.download_button("🛰️ 3. Baixar GPS Offline (GPX)", data=bytes_zip_gpx if bytes_zip_gpx else b"vazio", file_name=f"GPS_Offline_GPX - {data_atual_formatada}.zip", mime="application/zip", use_container_width=True, disabled=botoes_desabilitados)
        
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True, disabled=botoes_desabilitados): 
            limpar_roteirizador()

    # RESULTADOS FINAIS
    if is_done and not st.session_state.df_routed.empty:
        st.markdown("## 🎯 Resultados da Otimização")

        st.session_state.df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)

        df_routed = st.session_state.df_routed.copy()
        bases_records = st.session_state.bases_records
        colunas_exibir = st.session_state.colunas_exibir
        df_real_tasks = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
        
        tot_paradas = len(df_real_tasks)
        tot_obras_reais = sum(count_real_obras(r) for _, r in df_real_tasks.iterrows())
        
        tot_equipes = df_routed['BASE_ATRIBUIDA'].nunique()
        tot_km = f"{df_routed['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
        tot_prio = sum(count_real_obras(r) for _, r in df_real_tasks[df_real_tasks['PRIORIDADE'] == 'Sim'].iterrows()) if 'PRIORIDADE' in df_real_tasks else 0
        tot_super_pontos = len(df_real_tasks[df_real_tasks['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in df_real_tasks.columns else 0

        is_saneamento_puro = False
        if '_ORIGEM_BASE' in df_routed.columns:
            origens = df_routed['_ORIGEM_BASE'].unique()
            if 'SANEAMENTO' in origens and 'LEVANTAMENTO' not in origens:
                is_saneamento_puro = True

        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        c_m1.markdown(f'<div class="metric-card" style="border-left: 5px solid #0D256C;"><div class="metric-icon" style="background: rgba(13, 37, 108, 0.12);">🎯</div><div class="metric-content"><div class="metric-title">TOTAL DE OBRAS ROTEIRIZADAS</div><div class="metric-value">{tot_obras_reais} <span style="font-size:12px;color:#888;">(Em {tot_paradas} Pontos)</span></div></div></div>', unsafe_allow_html=True)
        c_m2.markdown(f'<div class="metric-card" style="border-left: 5px solid #8b5cf6;"><div class="metric-icon" style="background: rgba(139, 92, 246, 0.15);">👥</div><div class="metric-content"><div class="metric-title">Equipes Alocadas</div><div class="metric-value">{tot_equipes}</div></div></div>', unsafe_allow_html=True)
        c_m3.markdown(f'<div class="metric-card" style="border-left: 5px solid #55B929;"><div class="metric-icon" style="background: rgba(85, 185, 41, 0.15);">🛣️</div><div class="metric-content"><div class="metric-title">KM Total Projetado</div><div class="metric-value">{tot_km}</div></div></div>', unsafe_allow_html=True)
        
        if is_saneamento_puro:
            c_m4.markdown(f'<div class="metric-card" style="border-left: 5px solid #eab308;"><div class="metric-icon" style="background: rgba(234, 179, 8, 0.15);">🏢</div><div class="metric-content"><div class="metric-title">Pontos Agrupados (Super Pontos)</div><div class="metric-value">{tot_super_pontos}</div></div></div>', unsafe_allow_html=True)
        else:
            c_m4.markdown(f'<div class="metric-card" style="border-left: 5px solid #ef4444;"><div class="metric-icon" style="background: rgba(239, 68, 68, 0.15);">🚨</div><div class="metric-content"><div class="metric-title">Prioridades</div><div class="metric-value">{tot_prio}</div></div></div>', unsafe_allow_html=True)

        cfg_atual = st.session_state.vrp_state.get('config', {})
        obras_dia_meta = cfg_atual.get('obras_por_dia', 30)
        limite_periodos_meta = cfg_atual.get('limite_periodos', 5)
        tipo_periodo_meta = cfg_atual.get('tipo_periodo', 'Dia')
        roteirizar_tudo_meta = cfg_atual.get('roteirizar_tudo', False)
        
        dias_multiplicador = len(cfg_atual.get('dias_selecionados', [])) if tipo_periodo_meta == "Semana" else 1
        tot_equipes_cadastradas = len(set(b['LEVANTADOR'] for b in st.session_state.bases_records))
        
        if roteirizar_tudo_meta:
            meta_exata_por_equipe = float('inf')
            meta_global_exata = tot_obras_reais + st.session_state.get('tot_obras_nao_alocadas', 0)
        else:
            meta_exata_por_equipe = obras_dia_meta * dias_multiplicador * limite_periodos_meta
            meta_global_exata = meta_exata_por_equipe * tot_equipes_cadastradas
            
        # Adicionando a trava global matemática nos relatórios se ativada
        trava_global = cfg_atual.get('trava_global_obras', 0)
        if trava_global > 0:
            meta_global_exata = min(meta_global_exata, trava_global)
        
        obras_por_equipe = {b['LEVANTADOR']: 0 for b in st.session_state.bases_records}
            
        for _, r in df_real_tasks.iterrows():
            b_name = r['BASE_ATRIBUIDA']
            qtd = count_real_obras(r)
            if b_name in obras_por_equipe:
                obras_por_equipe[b_name] += qtd
                
        obras_faltantes = meta_global_exata - tot_obras_reais
        obras_sobrando_na_planilha = st.session_state.get('tot_obras_nao_alocadas', 0)
        
        if roteirizar_tudo_meta:
             if obras_sobrando_na_planilha > 0:
                 st.markdown(f'''
                 <div style="background-color: #fff3cd; color: #856404; padding: 20px; border-left: 6px solid #ffeeba; margin-bottom: 20px; border-radius: 8px;">
                     <h3 style="margin-top: 0; color: #856404;">⚠️ Modo Lista Contínua: {tot_obras_reais} Obras Roteirizadas</h3>
                     <p>O sistema processou a lista ignorando limites de tempo. No entanto, <b>{obras_sobrando_na_planilha} obras</b> ficaram de fora pois pertencem a municípios que nenhum levantador atende de forma explícita. Faça o download do KML "OBRAS NÃO ALOCADAS" para visualizar as rejeitadas.</p>
                 </div>
                 ''', unsafe_allow_html=True)
             else:
                 st.markdown(f'''
                 <div style="background-color: #d4edda; color: #155724; padding: 15px; border-left: 5px solid #c3e6cb; margin-bottom: 20px; border-radius: 4px;">
                     <h4 style="margin-top: 0; margin-bottom: 5px;">✅ Modo Lista Contínua Concluído!</h4>
                     <p style="margin: 0;">100% da sua planilha compatível (<b>{tot_obras_reais} obras</b>) foi roteirizada para os levantadores definidos no arquivo. O limite de semanas foi desativado e os cronogramas estendidos automaticamente.</p>
                 </div>
                 ''', unsafe_allow_html=True)
        else:
            if obras_faltantes > 0:
                if obras_sobrando_na_planilha > 0:
                    dica_extra = f"<li><b>Falta de Obras nos Municípios Atendidos:</b> O sistema detectou que sobraram <b>{obras_sobrando_na_planilha} obras</b> na planilha, mas elas pertencem a cidades que os seus levantadores atuais não atendem (ou foram rejeitadas pelo filtro de densidade). A meta de {meta_global_exata} não foi atingida. O sistema roteirizou o máximo possível ({tot_obras_reais} obras) com base na disponibilidade real. Faça o download do <b>KML de Obras Não Alocadas</b> para inspecionar.</li>"
                else:
                    dica_extra = f"<li><b>Falta de Obras na Planilha Geral:</b> O estoque total de obras válidas esgotou antes de fechar a meta operacional. Faltaram obras nos municípios de atuação. O sistema roteirizou a quantidade máxima encontrada ({tot_obras_reais} obras).</li>"
                
                st.markdown(f'''
                <div style="background-color: #fff3cd; color: #856404; padding: 20px; border-left: 6px solid #ffeeba; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <h3 style="margin-top: 0; color: #856404; display: flex; align-items: center;"><span style="font-size: 24px; margin-right: 10px;">⚠️</span> Quadro de Aviso: Quantidade de Obras Limitada pelo Estoque</h3>
                    <p style="font-size: 16px;">Você configurou o sistema para roteirizar <b>{meta_global_exata} obras</b> no total <i>({obras_dia_meta} obras/dia × {dias_multiplicador * limite_periodos_meta} dias × {tot_equipes_cadastradas} equipes)</i>.</p>
                    <p style="font-size: 16px;">No entanto, <b>não havia obras suficientes nos municípios e raio de 100km</b>. O algoritmo roteirizou a quantidade máxima encontrada e compatível: <b>{tot_obras_reais} obras</b>. <br>
                    <span style="color: #d9534f; font-weight: bold; font-size: 18px;">❌ Faltaram {obras_faltantes} obras para atingir a meta global escolhida.</span></p>
                    <hr style="border-top: 1px solid #ffeeba; margin: 15px 0;">
                    <h4 style="margin-bottom: 10px; color: #856404;">🔍 Resumo do Cenário:</h4>
                    <ul style="font-size: 14px; line-height: 1.6;">
                        {dica_extra}
                        <li>Verifique a aba <b>"Relatório de Déficit por Levantador"</b> abaixo para ver exatamente quantos e quais técnicos ficaram ociosos e em quais cidades você precisa adicionar mais notas no Excel.</li>
                    </ul>
                </div>
                ''', unsafe_allow_html=True)
            else:
                st.markdown(f'''
                <div style="background-color: #d4edda; color: #155724; padding: 15px; border-left: 5px solid #c3e6cb; margin-bottom: 20px; border-radius: 4px;">
                    <h4 style="margin-top: 0; margin-bottom: 5px;">✅ Meta de Despacho 100% Atingida!</h4>
                    <p style="margin: 0;">O sistema logístico preencheu perfeitamente a meta exata de <b>{meta_global_exata} obras</b>.</p>
                </div>
                ''', unsafe_allow_html=True)

        cf_comb = st.session_state.config_financeira.get('custo_combustivel', 0.0)
        cf_cons = st.session_state.config_financeira.get('consumo_veiculo', 0.0)
        cf_hora = st.session_state.config_financeira.get('custo_hora_equipe', 0.0)
        
        mostrar_financeiro = (cf_comb > 0) or (cf_hora > 0)

        if mostrar_financeiro:
            tot_km_val = df_routed['DISTANCIA_PONTO_ANTERIOR_KM'].sum()
            litros_gastos = tot_km_val / cf_cons if cf_cons > 0 else 0
            custo_total_combustivel = litros_gastos * cf_comb
            
            df_financeiro = df_routed.copy()
            df_financeiro['_HORA_INICIO_DT'] = pd.to_datetime(df_financeiro['_HORA_INICIO_DT'])
            df_financeiro['_HORA_FIM_DT'] = pd.to_datetime(df_financeiro['_HORA_FIM_DT'])
            
            custo_total_mao_de_obra = 0.0
            horas_totais = 0.0
            
            for (eq, periodo_f), group in df_financeiro.groupby(['BASE_ATRIBUIDA', 'PERIODO']):
                h_inicio = group['_HORA_INICIO_DT'].min()
                h_fim = group['_HORA_FIM_DT'].max()
                h_trab = (h_fim - h_inicio).total_seconds() / 3600.0
                horas_totais += h_trab
                custo_total_mao_de_obra += (h_trab * cf_hora)
                
            custo_operacao_total = custo_total_combustivel + custo_total_mao_de_obra
            custo_por_obra = custo_operacao_total / tot_obras_reais if tot_obras_reais > 0 else 0.0

            c_fin1, c_fin2, c_fin3, c_fin4 = st.columns(4)
            c_fin1.markdown(f'<div class="metric-card" style="border-left: 5px solid #f59e0b;"><div class="metric-icon" style="background: rgba(245, 158, 11, 0.15);">⛽</div><div class="metric-content"><div class="metric-title">Combustível Estimado</div><div class="metric-value">R$ {formatar_moeda(custo_total_combustivel)}</div></div></div>', unsafe_allow_html=True)
            c_fin2.markdown(f'<div class="metric-card" style="border-left: 5px solid #8b5cf6;"><div class="metric-icon" style="background: rgba(139, 92, 246, 0.15);">👷</div><div class="metric-content"><div class="metric-title">Mão de Obra ({horas_totais:.1f}h)</div><div class="metric-value">R$ {formatar_moeda(custo_total_mao_de_obra)}</div></div></div>', unsafe_allow_html=True)
            c_fin3.markdown(f'<div class="metric-card" style="border-left: 5px solid #ef4444;"><div class="metric-icon" style="background: rgba(239, 68, 68, 0.15);">💲</div><div class="metric-content"><div class="metric-title">Custo Total Operação</div><div class="metric-value">R$ {formatar_moeda(custo_operacao_total)}</div></div></div>', unsafe_allow_html=True)
            c_fin4.markdown(f'<div class="metric-card" style="border-left: 5px solid #55B929;"><div class="metric-icon" style="background: rgba(85, 185, 41, 0.15);">📊</div><div class="metric-content"><div class="metric-title">Custo Médio por Obra</div><div class="metric-value">R$ {formatar_moeda(custo_por_obra)}</div></div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("### 🗺️ Mapa Geográfico de Rotas")
        mapa = folium.Map(location=[df_routed['LATITUDE'].mean(), df_routed['LONGITUDE'].mean()], zoom_start=8) if not df_routed.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        
        cores_folium = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#ff9800', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548']
        lista_bases_mapa = df_routed['BASE_ATRIBUIDA'].unique().tolist()
        
        heat_data = [[r['LATITUDE'], r['LONGITUDE']] for _, r in df_real_tasks.iterrows()]
        HeatMap(heat_data, name="🔥 Mapa de Calor (Demandas)", radius=15, blur=10).add_to(mapa)
        marker_cluster = MarkerCluster(name="Obras (Agrupadas)").add_to(mapa)
        
        for base_nome in lista_bases_mapa:
            idx_cor = lista_bases_mapa.index(base_nome)
            cor_rota = cores_folium[idx_cor % len(cores_folium)]
            df_base_rota = df_routed[df_routed['BASE_ATRIBUIDA'] == base_nome]
            
            fg_linhas = folium.FeatureGroup(name=f"Rota: {base_nome}", show=False)
            
            for periodo_val in df_base_rota['PERIODO'].unique():
                df_periodo = df_base_rota[df_base_rota['PERIODO'] == periodo_val]
                
                pontos_linha_folium = []
                for _, r in df_periodo.iterrows():
                    if isinstance(r.get('ROTA_GEOMETRIA'), list):
                        for lon, lat in r['ROTA_GEOMETRIA']: pontos_linha_folium.append([lat, lon]) 
                            
                folium.PolyLine(pontos_linha_folium, color='black', weight=7, opacity=0.9).add_to(fg_linhas)
                folium.PolyLine(pontos_linha_folium, color=cor_rota, weight=3, opacity=1.0).add_to(fg_linhas)
                
                for r in df_periodo.to_dict('records'):
                    if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                    icone = identificar_icone_folium(r, df_routed.columns)
                    
                    is_super = str(r.get('SUPER_PONTO', '')).startswith('SIM')
                    if is_super:
                        cor_icone = 'orange'
                        qtd_str = str(r.get('SUPER_PONTO')).replace('SIM', '').strip()
                        pop_header_bg, pop_header_color = "#FFD700", "#000000"
                        pop_prio_txt = f"🏢 SUPER PONTO {qtd_str}"
                    else:
                        cor_icone = 'red' if r.get('PRIORIDADE') == "Sim" else 'blue'
                        pop_header_bg, pop_header_color = ("#d9534f", "#ffffff") if r.get('PRIORIDADE') == "Sim" else ("#0D256C", "#ffffff")
                        pop_prio_txt = "🚨 OBRA PRIORITÁRIA" if r.get('PRIORIDADE') == "Sim" else "📍 Atendimento Padrão"
                    
                    dist_prox = r.get('DISTANCIA_PROXIMO_PONTO_KM', 0.0)
                    
                    extra_rows_list = []
                    for c in colunas_exibir:
                        if c.upper() not in ['NOME_DIA', 'DIA_MES', 'SEMANA', 'HORA_INICIO', 'HORA_FIM', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'BASE_ATRIBUIDA']:
                            val_html = formatar_valor_coluna(c, r.get(c, ''))
                            extra_rows_list.append(f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;'>{html.escape(str(c))}:</td><td style='padding:3px 6px; color:#333;'>{val_html}</td></tr>")

                    extra_rows = "".join(extra_rows_list)
                    dia_mes_str = f" - {r.get('DIA_MES', '')}" if r.get('DIA_MES') else ""

                    popup_html = f"""
                    <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                        <div style="background:{pop_header_bg}; color:{pop_header_color}; padding:8px 10px; font-size:13px; font-weight:bold;">{pop_prio_txt}</div>
                        <div style="padding:10px; background:#fafafa; font-size:12px;">
                            <table style="width:100%; border-collapse:collapse;">
                                <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Ordem Logística:</td><td style="padding:3px 6px; color:#333;"><b>{r.get('ORDEM', 0)}</b> ({r.get('NOME_DIA', f'Dia {r.get("DIA", 0)}')}{dia_mes_str})</td></tr>
                                <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Ant.:</td><td style="padding:3px 6px; color:#333;">{r.get('DISTANCIA_PONTO_ANTERIOR_KM', 0)} KM</td></tr>
                                <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Próx.:</td><td style="padding:3px 6px; color:#333;">{dist_prox} KM</td></tr>
                                <tr><td colspan="2"><hr style="margin: 5px 0; border: 0; border-top: 1px solid #ddd;"></td></tr>
                                {extra_rows}
                            </table>
                        </div>
                    </div>"""
                    folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=cor_icone, icon=icone), popup=folium.Popup(popup_html, max_width=300)).add_to(marker_cluster)
            
            fg_linhas.add_to(mapa)
            
        folium.LayerControl().add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550, returned_objects=[])

        st.markdown("<br>", unsafe_allow_html=True)
        tab_dados, tab_relatorio, tab_hospedagem = st.tabs(["📊 Dados Tabulares", "📉 Relatório de Déficit", "🏨 Apoio Logístico (Hotéis)"])

        with tab_dados:
            st.markdown("#### Detalhamento de Rotas")
            df_display = st.session_state.df_routed.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'BASE_ATRIBUIDA'], errors='ignore')
            
            df_editado_ui = st.data_editor(
                df_display, use_container_width=True, height=400,
                column_config={ 
                    "LATITUDE": st.column_config.NumberColumn(disabled=True), "LONGITUDE": st.column_config.NumberColumn(disabled=True),
                    "DISTANCIA_PONTO_ANTERIOR_KM": st.column_config.ProgressColumn("Dist. Anterior (KM)", format="%.2f", min_value=0, max_value=30), 
                    "LINK_NAVEGACAO_OFFLINE": st.column_config.LinkColumn("Link GPS", display_text="📍 Abrir no Maps")
                }
            )

        with tab_relatorio:
            st.markdown("#### Análise de Ociosidade e Falta de Obras por Técnico")
            
            dados_relatorio = []
            for b_record in st.session_state.bases_records:
                nome_lev = b_record['LEVANTADOR']
                muns_atendidos = str(b_record.get('MUNICIPIO', b_record.get('RESIDENCIA', 'DESCONHECIDO')))
                
                qtd_roteirizada = obras_por_equipe.get(nome_lev, 0)
                
                if roteirizar_tudo_meta:
                    meta_exibicao = "Ilimitada"
                    deficit = 0
                    status_meta = "✅ Processamento Contínuo"
                else:
                    meta_exibicao = meta_exata_por_equipe
                    deficit = meta_exata_por_equipe - qtd_roteirizada
                    status_meta = "✅ Meta Atingida" if deficit <= 0 else "❌ Faltam Obras"
                
                dados_relatorio.append({
                    "Levantador": nome_lev,
                    "Municípios de Atuação": muns_atendidos,
                    "Meta (Obras)": meta_exibicao,
                    "Roteirizadas": qtd_roteirizada,
                    "Faltantes (Déficit)": deficit if deficit > 0 else 0,
                    "Status": status_meta
                })
            
            df_relatorio = pd.DataFrame(dados_relatorio)
            df_relatorio = df_relatorio.sort_values(by="Faltantes (Déficit)", ascending=False).reset_index(drop=True)
            
            if roteirizar_tudo_meta:
                st.info("O modo 'Lista Contínua Direta' foi utilizado. Nenhuma equipe teve limite de obras ou dias; os roteiros foram criados para abraçar 100% da planilha fornecida.")
            else:
                st.info("Abaixo estão os levantadores que não atingiram a quantidade de obras solicitada porque o estoque de notas da sua cidade de atuação esgotou na planilha e num raio de 100km.")
            
            max_v = int(meta_exata_por_equipe) if not roteirizar_tudo_meta else int(max(df_relatorio["Roteirizadas"].max(), 1))
            st.dataframe(
                df_relatorio,
                use_container_width=True,
                column_config={
                    "Faltantes (Déficit)": st.column_config.NumberColumn(
                        "Faltantes (Déficit)",
                        help="Quantas obras faltaram para bater a meta deste técnico.",
                        format="%d ⚠️"
                    ),
                    "Roteirizadas": st.column_config.ProgressColumn(
                        "Obras Roteirizadas",
                        format="%d",
                        min_value=0,
                        max_value=max_v
                    )
                }
            )

        with tab_hospedagem:
            st.markdown("#### 🏨 Análise de Distância Extrema e Pernoite")
            st.markdown("O sistema calcula o **Centro de Gravidade** (Centroid) do bloco de obras atribuído a cada levantador. Se a massa de trabalho estiver concentrada a mais de **60 km** da base do técnico, o sistema sugere opções de hospedagem na região para evitar o desgaste diário com deslocamento em rodovias.")
            
            hospedagens_sugeridas = False
            for base_record in st.session_state.bases_records:
                nome_tec = base_record['LEVANTADOR']
                lat_base = base_record.get('LATITUDE')
                lon_base = base_record.get('LONGITUDE')
                
                df_tec = df_real_tasks[df_real_tasks['BASE_ATRIBUIDA'] == nome_tec]
                if not df_tec.empty and pd.notna(lat_base) and pd.notna(lon_base):
                    centro_lat = df_tec['LATITUDE'].mean()
                    centro_lon = df_tec['LONGITUDE'].mean()
                    
                    distancia_base_polo = haversine_scalar(float(lat_base), float(lon_base), centro_lat, centro_lon)
                    
                    if distancia_base_polo > 60:
                        hospedagens_sugeridas = True
                        mun_polo = df_tec['MUNICIPIO'].mode().iloc[0] if 'MUNICIPIO' in df_tec.columns else "Região Afastada"
                        qtd_obras_polo = len(df_tec)
                        
                        link_hoteis = f"https://www.google.com/maps/search/hoteis+pousadas/@{centro_lat:.6f},{centro_lon:.6f},12z"
                        
                        st.markdown(f"""
                        <div style="background-color: #f8f9fa; border-left: 5px solid #0D256C; padding: 15px; margin-bottom: 15px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                            <h4 style="margin-top: 0; color: #0D256C;">👨‍🔧 {nome_tec}</h4>
                            <p style="margin-bottom: 5px;"><b>Polo de Obras:</b> {mun_polo} ({qtd_obras_polo} paradas programadas)</p>
                            <p style="margin-bottom: 5px; color: #d9534f;"><b>⚠️ Distância da Base:</b> {distancia_base_polo:.1f} KM</p>
                            <a href="{link_hoteis}" target="_blank" style="display: inline-block; margin-top: 10px; padding: 8px 15px; background-color: #55B929; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">🏨 Buscar Hotéis no Centro da Operação</a>
                        </div>
                        """, unsafe_allow_html=True)
            
            if not hospedagens_sugeridas:
                st.success("✅ Logística Segura: Nenhuma equipe recebeu um pacote de obras cujo centro de gravidade fique a mais de 60km de sua residência atual.")

        sidebar_html_placeholder.markdown(renderizar_painel_lateral(meta_exata_por_equipe if not roteirizar_tudo_meta else "Ilimitado", tot_obras_reais, tot_equipes_cadastradas, meta_global_exata if not roteirizar_tudo_meta else "Ilimitado"), unsafe_allow_html=True)
        return 

    if status_exec == "IDLE" and not is_done:
        st.markdown("### ⚙️ Selecione a Estratégia de Roteirização")
        modo_selecionado = st.radio(
            "Modo:",
            ["🎯 1. Planejamento Tático (IA Automática)", "♾️ 2. Lista Contínua (Técnico Fixo)"],
            horizontal=True,
            label_visibility="collapsed"
        )

        if "Tático" in modo_selecionado:
            st.info("💡 **Como funciona o Planejamento Tático:** A Inteligência Artificial assume o controle. Ela analisa todas as obras pendentes e as distribui de forma estratégica entre as equipes disponíveis, agrupando-as pela melhor rota geográfica e acionando o raio de 100km caso falte obras na cidade base.")
            modo_operacao = "1"
        else:
            st.info("💡 **Como funciona a Lista Contínua:** O sistema respeita estritamente a coluna 'LEVANTADOR' da sua planilha. A IA apenas calcula as distâncias e roteiriza 100% da lista de cada técnico, gerando quantos dias forem necessários.")
            modo_operacao = "2"

        st.markdown("---")
        
        df_tasks_alocadas = pd.DataFrame()
        bases_records = []
        colunas_originais = []
        col_prioridade = "Nenhuma"
        
        if modo_operacao == "1":
            col_up_1, col_up_2 = st.columns(2)
            with col_up_1:
                st.markdown("### 👥 1. Levantadores Principais")
                df_bases = pd.DataFrame()
                with st.container(border=True):
                    base_file = st.file_uploader("Suba a planilha Levantadores_MA", type=["xlsx", "xls"])
                
                if base_file:
                    try:
                        df_bases_temp_ui = ler_planilha_cached(base_file.getvalue())
                        df_bases_temp_ui.columns = normalize_cols(df_bases_temp_ui.columns)
                        if 'LEVANTADOR' not in df_bases_temp_ui.columns:
                            for p_nome in ['NOME', 'TECNICO', 'EQUIPE', 'COLABORADOR']:
                                if p_nome in df_bases_temp_ui.columns:
                                    df_bases_temp_ui = df_bases_temp_ui.rename(columns={p_nome: 'LEVANTADOR'})
                                    break
                        if 'LEVANTADOR' in df_bases_temp_ui.columns:
                            opcoes_levs = sorted([str(x) for x in df_bases_temp_ui['LEVANTADOR'].dropna().unique().tolist() if str(x).upper().strip() != 'SEM LEVANTADOR'])
                            levs_selecionados = st.multiselect("Selecione as Equipes Principais:", opcoes_levs, default=opcoes_levs)
                            
                            if levs_selecionados:
                                df_bases = df_bases_temp_ui[df_bases_temp_ui['LEVANTADOR'].isin(levs_selecionados)].copy()
                                if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                                    df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                                    df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                                elif 'RESIDENCIA' in df_bases.columns or 'MUNICIPIO' in df_bases.columns:
                                    col_ref = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                                    muns_unicos = df_bases[col_ref].dropna().unique()
                                    mapa_coords = {}
                                    with st.spinner("🌍 Buscando coordenadas base no satélite..."):
                                        for mun in muns_unicos:
                                            lat, lon = obter_coordenadas_municipio_cached(mun)
                                            mapa_coords[mun] = (lat, lon)
                                    df_bases['LATITUDE'] = df_bases[col_ref].map(lambda x: mapa_coords.get(x, (np.nan, np.nan))[0])
                                    df_bases['LONGITUDE'] = df_bases[col_ref].map(lambda x: mapa_coords.get(x, (np.nan, np.nan))[1])
                                else:
                                    df_bases['LATITUDE'] = np.nan
                                    df_bases['LONGITUDE'] = np.nan
                                df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE'])
                                df_bases['TIPO_EQUIPE'] = 'PRINCIPAL'
                        else: st.error("❌ A planilha não possui a coluna 'LEVANTADOR'.")
                    except Exception as e: st.error(f"Erro ao ler a planilha: {e}")

                st.markdown("##### Regra de Atribuição Territorial")
                tipo_atribuicao = st.radio("Regra", ["Por Municípios Atendidos (Lê texto da planilha)", "Por Proximidade Geográfica das Coordenadas (Ignora texto)", "Clusterização Inteligente por IA (K-Means)"], index=0, label_visibility="collapsed")

            with col_up_2:
                st.markdown("### 📁 2. Upload de Demandas (Obras)")
                with st.container(border=True): task_files = st.file_uploader("1️⃣ Base Levantamento", type=["xlsx", "xls"], accept_multiple_files=True, key="lev_uploader")
                with st.container(border=True): saneamento_files = st.file_uploader("2️⃣ Base Saneamento", type=["xlsx", "xls"], accept_multiple_files=True, key="san_uploader")
                with st.container(border=True): generica_files = st.file_uploader("3️⃣ Base Genérica / Livre (Qualquer Planilha)", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="gen_uploader")
                with st.container(border=True): status_file = st.file_uploader("4️⃣ Planilha Atualizada SharePoint (Opcional)", type=["xlsx", "xls"])
                
                df_status_upload = pd.DataFrame()
                coluna_status_selecionada = None
                if status_file:
                    try:
                        df_status_upload = ler_planilha_cached(status_file.getvalue())
                        cols_status = df_status_upload.columns.tolist()
                        coluna_status_selecionada = st.selectbox("📌 Coluna Status?", cols_status, index=4 if len(cols_status) >= 5 else 0)
                    except Exception as e: st.error(f"Erro ao ler status: {e}")
                
                st.markdown("##### 🧑‍🤝‍🧑 3. Equipes de Apoio (Temporários)")
                with st.container(border=True): temp_bases_files = st.file_uploader("Suba a(s) planilha(s) de Apoio", type=["xlsx", "xls"], accept_multiple_files=True)
                df_bases_temp = pd.DataFrame()
                if temp_bases_files:
                    try:
                        dfs_temp = []
                        for f in temp_bases_files:
                            df_t = ler_planilha_cached(f.getvalue())
                            df_t.columns = normalize_cols(df_t.columns)
                            if 'LEVANTADOR' not in df_t.columns:
                                for p_nome in ['NOME', 'TECNICO', 'EQUIPE']:
                                    if p_nome in df_t.columns: df_t = df_t.rename(columns={p_nome: 'LEVANTADOR'}); break
                            dfs_temp.append(df_t)
                        df_bases_temp_full = pd.concat(dfs_temp, ignore_index=True)
                        if 'LEVANTADOR' in df_bases_temp_full.columns:
                            opcoes_levs_temp = sorted([str(x) for x in df_bases_temp_full['LEVANTADOR'].dropna().unique().tolist() if str(x).upper().strip() != 'SEM LEVANTADOR'])
                            levs_temp_selecionados = st.multiselect("Selecione as Equipes:", opcoes_levs_temp, default=opcoes_levs_temp)
                            if levs_temp_selecionados:
                                df_bases_temp = df_bases_temp_full[df_bases_temp_full['LEVANTADOR'].isin(levs_temp_selecionados)].copy()
                                if 'LATITUDE' in df_bases_temp.columns and 'LONGITUDE' in df_bases_temp.columns:
                                    df_bases_temp['LATITUDE'] = pd.to_numeric(df_bases_temp['LATITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                                    df_bases_temp['LONGITUDE'] = pd.to_numeric(df_bases_temp['LONGITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                                elif 'RESIDENCIA' in df_bases_temp.columns or 'MUNICIPIO' in df_bases_temp.columns:
                                    col_ref_temp = 'RESIDENCIA' if 'RESIDENCIA' in df_bases_temp.columns else 'MUNICIPIO'
                                    muns_unicos_temp = df_bases_temp[col_ref_temp].dropna().unique()
                                    mapa_coords_temp = {}
                                    with st.spinner("🌍 Buscando coordenadas bases temporárias..."):
                                        for mun in muns_unicos_temp:
                                            lat, lon = obter_coordenadas_municipio_cached(mun)
                                            mapa_coords_temp[mun] = (lat, lon)
                                    df_bases_temp['LATITUDE'] = df_bases_temp[col_ref_temp].map(lambda x: mapa_coords_temp.get(x, (np.nan, np.nan))[0])
                                    df_bases_temp['LONGITUDE'] = df_bases_temp[col_ref_temp].map(lambda x: mapa_coords_temp.get(x, (np.nan, np.nan))[1])
                                else:
                                    df_bases_temp['LATITUDE'] = np.nan
                                    df_bases_temp['LONGITUDE'] = np.nan
                                df_bases_temp = df_bases_temp.dropna(subset=['LATITUDE', 'LONGITUDE'])
                                df_bases_temp['TIPO_EQUIPE'] = 'TEMPORARIA'
                    except Exception as e: st.error(f"Erro: {e}")

                qtd_eq_princ = df_bases['LEVANTADOR'].nunique() if 'df_bases' in locals() and not df_bases.empty else 0
                qtd_eq_temp = df_bases_temp['LEVANTADOR'].nunique() if 'df_bases_temp' in locals() and not df_bases_temp.empty else 0
                qtd_eq_atual_live = qtd_eq_princ + qtd_eq_temp
                st.session_state.qtd_equipes_ativas = qtd_eq_atual_live
                
                tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
                dias_multiplier = len(dias_semana_selecionados) if tipo_periodo_clean == 'Semana' else 1
                cap_por_eq_live = obras_por_dia * dias_multiplier * limite_periodos
                cap_total_estimada_live = cap_por_eq_live * (qtd_eq_atual_live if qtd_eq_atual_live > 0 else 1)
                
                sidebar_html_placeholder.markdown(renderizar_painel_lateral(cap_por_eq_live, 0, qtd_eq_atual_live, cap_total_estimada_live), unsafe_allow_html=True)

                if not task_files and not saneamento_files and not generica_files: 
                    st.info("Aguardando upload de obras para iniciar o roteamento.")
                    return

                try:
                    dfs = []
                    if task_files:
                        for f in task_files:
                            df_temp = ler_planilha_cached(f.getvalue())
                            if len(dfs) == 0: st.session_state.colunas_originais_lev = df_temp.columns.tolist()
                            df_temp.columns = normalize_cols(df_temp.columns)
                            df_temp['_ORIGEM_BASE'] = 'LEVANTAMENTO'
                            if 'PRIORIDADE' in df_temp.columns: df_temp['_PRIORIDADE_ORIGINAL'] = df_temp['PRIORIDADE']
                            if 'PROTOCOLO' not in df_temp.columns:
                                for col_candidata in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
                                    if col_candidata in df_temp.columns:
                                        df_temp['PROTOCOLO'] = df_temp[col_candidata]
                                        break
                            dfs.append(df_temp)
                            
                    if saneamento_files:
                        for f in saneamento_files:
                            df_temp = ler_planilha_cached(f.getvalue())
                            if len(dfs) == 0: st.session_state.colunas_originais_san = df_temp.columns.tolist()
                            df_temp.columns = normalize_cols(df_temp.columns)
                            df_temp['_ORIGEM_BASE'] = 'SANEAMENTO'
                            if 'PRIORIDADE' in df_temp.columns: df_temp['_PRIORIDADE_ORIGINAL'] = df_temp['PRIORIDADE']
                            if 'LATITUDE PROJETO' in df_temp.columns and 'LATITUDE' not in df_temp.columns: df_temp['LATITUDE'] = df_temp['LATITUDE PROJETO']
                            if 'LONGITUDE PROJETO' in df_temp.columns and 'LONGITUDE' not in df_temp.columns: df_temp['LONGITUDE'] = df_temp['LONGITUDE PROJETO']
                            if 'PROTOCOLO' not in df_temp.columns:
                                for col_candidata in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
                                    if col_candidata in df_temp.columns:
                                        df_temp['PROTOCOLO'] = df_temp[col_candidata]
                                        break
                            dfs.append(df_temp)

                    if generica_files:
                        for f in generica_files:
                            if f.name.endswith('.csv'): df_temp = pd.read_csv(f)
                            else: df_temp = ler_planilha_cached(f.getvalue())
                            if len(dfs) == 0: st.session_state.colunas_originais_gen = df_temp.columns.tolist()
                            df_temp.columns = normalize_cols(df_temp.columns)
                            df_temp['_ORIGEM_BASE'] = 'GENERICA'
                            if 'PRIORIDADE' in df_temp.columns: df_temp['_PRIORIDADE_ORIGINAL'] = df_temp['PRIORIDADE']
                            if 'LATITUDE' not in df_temp.columns or 'LONGITUDE' not in df_temp.columns:
                                st.error(f"🚨 A planilha '{f.name}' foi ignorada: É obrigatório conter colunas chamadas 'LATITUDE' e 'LONGITUDE'.")
                                continue
                            if 'PROTOCOLO' not in df_temp.columns:
                                id_cols = ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS', 'ID', 'CODIGO', 'CHAMADO', 'CHAVE']
                                found_id = False
                                for c in id_cols:
                                    if c in df_temp.columns:
                                        df_temp['PROTOCOLO'] = df_temp[c]
                                        found_id = True
                                        break
                                if not found_id: df_temp['PROTOCOLO'] = [f"GEN-{i+1}" for i in range(len(df_temp))]
                            dfs.append(df_temp)
                            
                    if not dfs: return

                    df_tasks = pd.concat(dfs, ignore_index=True)
                    total_obras_inicial = len(df_tasks)
                    cols_orig_lev = st.session_state.get('colunas_originais_lev', [])
                    cols_orig_san = st.session_state.get('colunas_originais_san', [])
                    cols_orig_gen = st.session_state.get('colunas_originais_gen', [])
                    st.session_state.colunas_originais = list(dict.fromkeys(cols_orig_lev + cols_orig_san + cols_orig_gen))
                    
                    for c_nome in ['CONTA CONTRATO', 'INSTALACAO', 'PROTOCOLO']:
                        if c_nome in df_tasks.columns: df_tasks[c_nome] = df_tasks[c_nome].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '-')
                            
                except Exception as e: st.error(f"Erro ao unificar planilhas: {e}"); return

                # ABA DE DASHBOARD INTERATIVO (IMPLEMENTAÇÃO #2)
                st.markdown("##### 🗂️ Triagem Dinâmica de Notas (Bolo Geral)")
                col_dash1, col_dash2, col_dash3, col_dash4 = st.columns(4)
                
                if 'TIPO NOTA' in df_tasks.columns:
                    tipos_counts = df_tasks['TIPO NOTA'].value_counts()
                    t_unr = tipos_counts.get('UNR', 0)
                    t_mgd = tipos_counts.get('MGD', 0)
                    t_asc = tipos_counts.get('ASC', 0)
                    t_dif = tipos_counts.get('DIF', 0)
                    
                    col_dash1.metric("Notas UNR", f"{t_unr}")
                    col_dash2.metric("Notas MGD", f"{t_mgd}")
                    col_dash3.metric("Notas ASC", f"{t_asc}")
                    col_dash4.metric("Notas DIF", f"{t_dif}")
                    
                    tipos_unicos_na_base = [str(x) for x in df_tasks['TIPO NOTA'].dropna().unique()]
                    tipos_rejeitados = st.multiselect("🗑️ Selecione Tipos de Nota para DESCARTAR temporariamente da operação:", options=tipos_unicos_na_base, default=[])
                    if tipos_rejeitados:
                        df_tasks = df_tasks[~df_tasks['TIPO NOTA'].isin(tipos_rejeitados)]
                        st.info(f"Omitindo obras do tipo {tipos_rejeitados}.")

                if not df_status_upload.empty and coluna_status_selecionada:
                    df_tasks = atualizar_status_via_df(df_tasks, df_status_upload, coluna_status_selecionada)

                ignorar_despacho_m1 = st.checkbox("Filtro: Ignorar obras já despachadas (com DATA DESPACHO CAMPO)?", value=False, help="Se marcado, o sistema não roteirizará obras que já tenham data preenchida.")
                if ignorar_despacho_m1 and 'DATA DESPACHO CAMPO' in df_tasks.columns:
                    mask_despacho = df_tasks['DATA DESPACHO CAMPO'].notna() & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip() != '') & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip().str.lower() != 'nan')
                    obras_despachadas = mask_despacho.sum()
                    if obras_despachadas > 0:
                        st.info(f"⏭️ {obras_despachadas} obras ignoradas (DATA DESPACHO CAMPO preenchida).")
                        df_tasks = df_tasks[~mask_despacho]

            st.markdown("---")
            has_levantamento = 'LEVANTAMENTO' in df_tasks['_ORIGEM_BASE'].values
            has_saneamento = 'SANEAMENTO' in df_tasks['_ORIGEM_BASE'].values
            has_generica = 'GENERICA' in df_tasks['_ORIGEM_BASE'].values
            df_list = []
            
            # MATRIZ MULTI-FILTRO (ESCOPO DA OPERAÇÃO / IMPLEMENTAÇÃO #3)
            st.markdown("### 🎯 Escopo da Operação (Multi-Filtro)")
            col_escopo1, col_escopo2 = st.columns(2)
            
            if 'Regional' in df_tasks.columns or 'REGIONAL' in df_tasks.columns:
                col_reg_real = 'Regional' if 'Regional' in df_tasks.columns else 'REGIONAL'
                regs_unicas = [str(x) for x in df_tasks[col_reg_real].dropna().unique()]
                regs_selecionadas = col_escopo1.multiselect("🌍 Filtrar por REGIONAL (Deixe vazio para roteirizar todas):", options=regs_unicas, default=[])
                if regs_selecionadas:
                    df_tasks = df_tasks[df_tasks[col_reg_real].isin(regs_selecionadas)]
                    
            if 'PAT' in df_tasks.columns:
                pats_unicos = [str(x) for x in df_tasks['PAT'].dropna().unique()]
                pats_selecionados = col_escopo2.multiselect("🏗️ Filtrar por PAT (Deixe vazio para roteirizar todos):", options=pats_unicos, default=[])
                if pats_selecionados:
                    df_tasks = df_tasks[df_tasks['PAT'].isin(pats_selecionados)]
                    
            st.markdown("---")
            
            if has_levantamento:
                with st.expander("🛠️ 4A. Filtros Iniciais - Base LEVANTAMENTO", expanded=True):
                    df_lev = df_tasks[df_tasks['_ORIGEM_BASE'] == 'LEVANTAMENTO'].copy()
                    c_filt1, c_filt2, c_filt3 = st.columns(3)
                    if 'STATUS LIST' in df_lev.columns:
                        status_brutos = [str(x).strip().upper() for x in df_lev['STATUS LIST'].unique() if pd.notna(x) and str(x).lower() != 'nan']
                        status_unicos = sorted(list(set(status_brutos)))
                        padroes_ativos = [s for s in status_unicos if s in STATUS_PADRAO]
                        status_selecionados = c_filt1.multiselect("📌 Filtrar Status de Início:", options=status_unicos, default=padroes_ativos)
                        if not status_selecionados: st.warning("Selecione um status."); return
                        df_lev['STATUS_LIMPO'] = df_lev['STATUS LIST'].astype(str).str.strip().str.upper()
                        df_lev = df_lev[df_lev['STATUS_LIMPO'].isin(status_selecionados)].drop(columns=['STATUS_LIMPO'])

                    colunas_validas = [c for c in df_lev.columns if not c.startswith('_')]
                    idx_default = 0
                    for i, col in enumerate(colunas_validas):
                        if col in ['TIPO NOTA', 'TIPO DE NOTA', 'TIPO DEMANDA', 'TIPO', 'SERVICO']: idx_default = i + 1; break
                            
                    coluna_prio = c_filt2.selectbox("📌 1. Qual coluna define a prioridade?", ["Nenhuma"] + colunas_validas, index=idx_default, key='prio_col_lev_din')
                    
                    if coluna_prio != "Nenhuma":
                        df_lev[coluna_prio] = df_lev[coluna_prio].fillna('SEM TIPO').astype(str).str.strip().str.upper()
                        valores_unicos = sorted(list(set(df_lev[coluna_prio].unique())))
                        tipos_selecionados = c_filt2.multiselect(f"🏷️ 2. Filtrar dados na coluna '{coluna_prio}':", valores_unicos, default=valores_unicos, key='filt_prio_lev')
                        if not tipos_selecionados: st.warning(f"Selecione valores em {coluna_prio}."); return
                        df_lev = df_lev[df_lev[coluna_prio].isin(tipos_selecionados)]
                        default_prio = [x for x in tipos_selecionados if x in TIPOS_PRIORITARIOS]
                        valores_prio = c_filt3.multiselect(f"🚨 3. Definir PRIORIDADE em '{coluna_prio}':", tipos_selecionados, default=default_prio, key='def_prio_lev')
                        
                        if valores_prio: df_lev['PRIORIDADE'] = df_lev[coluna_prio].apply(lambda x: 'Sim' if x in valores_prio else 'Não')
                        else: df_lev['PRIORIDADE'] = 'Não'
                        st.session_state.col_prioridade_lev = coluna_prio
                    else:
                        df_lev['PRIORIDADE'] = 'Não'
                        st.session_state.col_prioridade_lev = "Nenhuma"

                    if 'STATUS SAP' in df_lev.columns: df_lev = df_lev[~df_lev['STATUS SAP'].astype(str).str.strip().str.upper().isin(['CANC', 'FINL'])]
                    df_list.append(df_lev)
                    
            if has_saneamento:
                with st.expander("🛠️ 4B. Filtros Iniciais - Base SANEAMENTO", expanded=True):
                    st.info("✅ Base de Saneamento detectada: Todas as obras foram aprovadas automaticamente para o roteamento (sem filtros de área).")
                    df_san = df_tasks[df_tasks['_ORIGEM_BASE'] == 'SANEAMENTO'].copy()
                    df_san['PRIORIDADE'] = 'Não'
                    df_list.append(df_san)

            if has_generica:
                with st.expander("🛠️ 4C. Filtros Iniciais - Base GENÉRICA", expanded=True):
                    df_gen = df_tasks[df_tasks['_ORIGEM_BASE'] == 'GENERICA'].copy()
                    st.info("💡 A base Genérica é flexível. O sistema tenta adivinhar a prioridade, mas você pode alterar as regras abaixo.")
                    col_c1, col_c2 = st.columns(2)
                    colunas_validas = [c for c in df_gen.columns if not c.startswith('_')]
                    idx_default = 0
                    if 'TIPO NOTA' in colunas_validas: idx_default = colunas_validas.index('TIPO NOTA') + 1
                    coluna_prio = col_c1.selectbox("📌 1. Qual coluna define a prioridade?", ["Nenhuma"] + colunas_validas, index=idx_default, key='prio_col_gen')
                    
                    if coluna_prio != "Nenhuma":
                        df_gen[coluna_prio] = df_gen[coluna_prio].fillna('SEM TIPO').astype(str).str.strip()
                        valores_unicos = [x for x in df_gen[coluna_prio].unique() if str(x).lower() != 'nan']
                        default_prio = []
                        if coluna_prio == 'TIPO NOTA': default_prio = [x for x in valores_unicos if x in TIPOS_PRIORITARIOS]
                        valores_prio = col_c2.multiselect(f"🚨 2. Definir Obras PRIORITÁRIAS em '{coluna_prio}':", valores_unicos, default=default_prio, key='prio_val_gen')
                        if valores_prio: df_gen['PRIORIDADE'] = df_gen[coluna_prio].apply(lambda x: 'Sim' if str(x).strip() in valores_prio else 'Não')
                        else: df_gen['PRIORIDADE'] = 'Não'
                        st.session_state.col_prioridade_gen = coluna_prio
                    else:
                        df_gen['PRIORIDADE'] = 'Não'
                        st.session_state.col_prioridade_gen = "Nenhuma"
                    df_list.append(df_gen)

            if not df_list: return
            df_tasks = pd.concat(df_list, ignore_index=True)

            # --- REGRA DE OURO 1: CORREÇÃO DE LEVANTAMENTO É SEMPRE PRIORIDADE ---
            if 'STATUS LIST' in df_tasks.columns:
                mask_correcao = df_tasks['STATUS LIST'].astype(str).str.strip().str.upper().isin(['CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO'])
                df_tasks.loc[mask_correcao, 'PRIORIDADE'] = 'Sim'
            # -------------------------------------------------------------------
            
            # --- REGRA DE OURO 2: COLUNA "PRIORIDADE" DIFERENTE DE ZERO ---
            if '_PRIORIDADE_ORIGINAL' in df_tasks.columns:
                mask_not_zero = (
                    df_tasks['_PRIORIDADE_ORIGINAL'].notna() & 
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '') & 
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '0') & 
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '0.0') & 
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NAN') &
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NÃO') &
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NAO') &
                    (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'FALSE')
                )
                df_tasks.loc[mask_not_zero, 'PRIORIDADE'] = 'Sim'
            # -------------------------------------------------------------------

            if 'LATITUDE' not in df_tasks.columns or 'LONGITUDE' not in df_tasks.columns:
                st.error("🚨 ERRO GRAVE: As planilhas carregadas não possuem as colunas obrigatórias 'LATITUDE' e/ou 'LONGITUDE'. Verifique o cabeçalho dos arquivos enviados e certifique-se de que o sistema de mapeamento não foi comprometido na origem.")
                st.stop()

            df_tasks['LATITUDE'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
            df_tasks['LONGITUDE'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
            df_tasks = resgatar_coordenadas(df_tasks)
            
            erros_nome = 0
            if 'NOME' not in df_tasks.columns: df_tasks['NOME'] = "SEM NOME"
            for col_nome in ['NOME DO SOLICITANTE', 'CLIENTE', 'RAZAO SOCIAL', 'DESCRICAO', 'ENDERECO', 'LOCAL']:
                if col_nome in df_tasks.columns: df_tasks['NOME'] = df_tasks['NOME'].fillna(df_tasks[col_nome])

            mask_origin_strict = df_tasks['_ORIGEM_BASE'].isin(['LEVANTAMENTO', 'SANEAMENTO'])
            mask_invalid = df_tasks['NOME'].isna() | (df_tasks['NOME'].astype(str).str.strip() == '') | (df_tasks['NOME'].astype(str).str.strip().str.lower() == 'nan')
            drop_mask = mask_origin_strict & mask_invalid
            erros_nome += drop_mask.sum()
            df_tasks = df_tasks[~drop_mask]

            df_tasks, qtd_condensada = fundir_super_pontos(df_tasks, raio_metros=st.session_state.get('vrp_state', {}).get('config', {}).get('raio_super_ponto', 100) if 'vrp_state' in st.session_state else 100, agrupar_por_levantador=False)
            if qtd_condensada > 0: st.toast(f"✅ Inteligência condensou {qtd_condensada} obras repetidas no mesmo endereço em 'Super Pontos'.")

            st.markdown("#### 📊 Raio-X da Base de Dados Carregada")
            tot_obras_aprovadas = sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_tasks.iterrows())
            st.markdown(f"""
            <div class="profiling-box">
                <b>Análise Estrutural:</b> Das {total_obras_inicial} linhas encontradas, o sistema aprovou <b style="color: #0D256C;">{tot_obras_aprovadas} obras reais</b> (compactadas em {len(df_tasks)} paradas físicas no mapa). <br>
            </div>
            """, unsafe_allow_html=True)
            if df_tasks.empty: return
            
            # --- LIMITE GLOBAL DA DEMANDA (TRAVA DE ESTOQUE / IMPLEMENTAÇÃO #4) ---
            if trava_global_obras > 0 and modo_operacao == "1":
                if tot_obras_aprovadas > trava_global_obras:
                    st.info(f"✂️ A Trava Global de Operação foi ativada. O sistema limitará o roteamento às primeiras {trava_global_obras} obras prioritárias/próximas e descartará o excedente ({tot_obras_aprovadas - trava_global_obras} obras).")
                    
                    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
                    
                    obras_aceitas_fisicas = []
                    contador_real = 0
                    for idx, r in df_tasks.iterrows():
                        qtd = len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
                        if contador_real + qtd <= trava_global_obras:
                            obras_aceitas_fisicas.append(r)
                            contador_real += qtd
                        else:
                            break
                            
                    df_tasks = pd.DataFrame(obras_aceitas_fisicas)
            # ------------------------------------------------------------------------

            # --- FILTRO ALTA DENSIDADE (MODO PRODUTIVIDADE) ---
            df_sparse_global = pd.DataFrame()
            if modo_produtividade and modo_operacao == "1":
                with st.spinner("🔥 Modo Produtividade: Mapeando bolsões e isolando obras esparsas..."):
                    lats = df_tasks['LATITUDE'].values
                    lons = df_tasks['LONGITUDE'].values
                    prio_mask = (df_tasks['PRIORIDADE'] == 'Sim').values
                    keep_mask = np.copy(prio_mask)
                    
                    for i in range(len(df_tasks)):
                        if not keep_mask[i]:
                            dists = haversine_vectorized(lats[i], lons[i], lats, lons)
                            if np.sum(dists <= 2.0) >= min_vizinhos:
                                keep_mask[i] = True
                                
                    df_sparse_global = df_tasks[~keep_mask].copy()
                    if not df_sparse_global.empty:
                        df_sparse_global['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
                        df_sparse_global['MOTIVO_REJEICAO'] = "Alta Densidade (Obra Isolada)"
                    df_tasks = df_tasks[keep_mask].copy()
                    
                    if not df_sparse_global.empty:
                        st.toast(f"🔥 {len(df_sparse_global)} obras isoladas foram ignoradas para maximizar a produtividade da equipe!")
            # ----------------------------------------------------

            df_tasks_alocadas = pd.DataFrame()
            bases_principais_records = df_bases.to_dict('records') if not df_bases.empty else []
            bases_temporarias_records = df_bases_temp.to_dict('records') if not df_bases_temp.empty else []
            todas_bases_records = bases_principais_records + bases_temporarias_records
            
            if len(todas_bases_records) > 0:
                df_tasks['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
                df_tasks['COORD_KEY'] = df_tasks['LATITUDE'].astype(str) + "_" + df_tasks['LONGITUDE'].astype(str)
                coords_com_prio = df_tasks[df_tasks['PRIORIDADE'] == 'Sim']['COORD_KEY'].unique()
                df_tasks['PRECISA_PRINCIPAL'] = df_tasks['COORD_KEY'].isin(coords_com_prio)
                
                df_prio_e_agregadas = df_tasks[df_tasks['PRECISA_PRINCIPAL']].copy()
                df_comum_puro = df_tasks[~df_tasks['PRECISA_PRINCIPAL']].copy()
                col_mun_name = 'MUNICIPIO' if 'MUNICIPIO' in df_tasks.columns else ('CIDADE' if 'CIDADE' in df_tasks.columns else None)
                if col_mun_name:
                    df_prio_e_agregadas['MUN_LIMPO'] = normalizar_municipios(df_prio_e_agregadas[col_mun_name].fillna(''))
                    df_comum_puro['MUN_LIMPO'] = normalizar_municipios(df_comum_puro[col_mun_name].fillna(''))
                else:
                    df_prio_e_agregadas['MUN_LIMPO'] = ''
                    df_comum_puro['MUN_LIMPO'] = ''
                
                mun_to_main = {}
                mun_to_all = {}
                for b in todas_bases_records:
                    muns_str = str(b.get('MUNICIPIO', b.get('RESIDENCIA', '')))
                    for m in muns_str.split(','):
                        m_limpo = normalizar_municipios(pd.Series([m])).iloc[0]
                        if m_limpo:
                            if m_limpo not in mun_to_all: mun_to_all[m_limpo] = []
                            if b['LEVANTADOR'] not in mun_to_all[m_limpo]: mun_to_all[m_limpo].append(b['LEVANTADOR'])
                            if b.get('TIPO_EQUIPE') == 'PRINCIPAL':
                                if m_limpo not in mun_to_main: mun_to_main[m_limpo] = []
                                if b['LEVANTADOR'] not in mun_to_main[m_limpo]: mun_to_main[m_limpo].append(b['LEVANTADOR'])

                base_counts = {b['LEVANTADOR']: 0 for b in todas_bases_records}
                
                tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
                dias_multiplier = len(dias_semana_selecionados) if tipo_periodo_clean == 'Semana' else 1
                max_capacity = obras_por_dia * dias_multiplier * limite_periodos

                def assign_load_balanced_strict_and_fallback(df_sub_prio, df_sub_comum, allowed_bases):
                    if df_sub_prio.empty and df_sub_comum.empty: return pd.DataFrame(), pd.DataFrame()
                    
                    assigned_tasks = []
                    unassigned_tasks = []
                    
                    df_combinado = pd.concat([df_sub_prio, df_sub_comum], ignore_index=True)
                    df_combinado = df_combinado.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
                    
                    lista_obras = df_combinado.to_dict('records')
                    
                    for row in lista_obras:
                        qtd_real = len(row.get('_ORIGINAL_ROWS', [1])) if isinstance(row.get('_ORIGINAL_ROWS'), list) else 1
                        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
                        mun_str = str(row.get('MUN_LIMPO', ''))
                        is_prio = row.get('PRIORIDADE') == 'Sim'
                        
                        precisa_4x4 = str(row.get('TIPO VEICULO', '')).strip().upper() == '4X4'
                        valid_names = set(mun_to_main.get(mun_str, [])) if is_prio else set(mun_to_all.get(mun_str, []))
                        valid_bases = [b for b in allowed_bases if b['LEVANTADOR'] in valid_names]
                        
                        if precisa_4x4:
                            valid_bases = [b for b in valid_bases if str(b.get('VEICULO', '')).upper() == '4X4']
                            
                        best_base = None
                        best_dist = float('inf')
                        
                        if pd.notna(lat) and pd.notna(lon):
                            for b in valid_bases:
                                b_name = b['LEVANTADOR']
                                if base_counts[b_name] + qtd_real <= max_capacity:
                                    b_lat, b_lon = b.get('LATITUDE'), b.get('LONGITUDE')
                                    if pd.notna(b_lat) and pd.notna(b_lon):
                                        d = haversine_scalar(lat, lon, float(b_lat), float(b_lon))
                                        if d < best_dist:
                                            best_dist = d
                                            best_base = b_name
                                            
                        if best_base:
                            base_counts[best_base] += qtd_real
                            row['BASE_ATRIBUIDA'] = best_base
                            assigned_tasks.append(row)
                        else:
                            unassigned_tasks.append(row)
                            
                    still_unassigned = []
                    for row in unassigned_tasks:
                        qtd_real = len(row.get('_ORIGINAL_ROWS', [1])) if isinstance(row.get('_ORIGINAL_ROWS'), list) else 1
                        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
                        
                        precisa_4x4 = str(row.get('TIPO VEICULO', '')).strip().upper() == '4X4'
                        valid_bases_fallback = allowed_bases
                        
                        if precisa_4x4:
                            valid_bases_fallback = [b for b in valid_bases_fallback if str(b.get('VEICULO', '')).upper() == '4X4']
                            
                        best_base_fb = None
                        valid_bases_fallback = sorted(valid_bases_fallback, key=lambda b: base_counts[b['LEVANTADOR']])
                        
                        if pd.notna(lat) and pd.notna(lon):
                            for b in valid_bases_fallback:
                                b_name = b['LEVANTADOR']
                                if base_counts[b_name] + qtd_real <= max_capacity:
                                    b_lat, b_lon = b.get('LATITUDE'), b.get('LONGITUDE')
                                    if pd.notna(b_lat) and pd.notna(b_lon):
                                        d = haversine_scalar(lat, lon, float(b_lat), float(b_lon))
                                        if d <= 100.0: 
                                            best_base_fb = b_name
                                            break
                                            
                        if best_base_fb:
                            base_counts[best_base_fb] += qtd_real
                            row['BASE_ATRIBUIDA'] = best_base_fb
                            assigned_tasks.append(row)
                        else:
                            row['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
                            row['MOTIVO_REJEICAO'] = "Estoque do Técnico Lotado ou > 100km"
                            still_unassigned.append(row)

                    return pd.DataFrame(assigned_tasks), pd.DataFrame(still_unassigned)

                df_tasks_alocadas, df_unallocated = assign_load_balanced_strict_and_fallback(df_prio_e_agregadas, df_comum_puro, todas_bases_records)
                
                # ADICIONANDO AS OBRAS REJEITADAS PELA DENSIDADE (MODO PRODUTIVIDADE) NA LISTA DE ÓRFÃS
                if not df_sparse_global.empty:
                    df_unallocated = pd.concat([df_unallocated, df_sparse_global], ignore_index=True)
                
                df_tasks_alocadas = df_tasks_alocadas.drop(columns=['COORD_KEY', 'PRECISA_PRINCIPAL', 'MUN_LIMPO'], errors='ignore')
                st.session_state.df_unallocated = df_unallocated
                st.session_state.tot_obras_nao_alocadas = sum(len(r['_ORIGINAL_ROWS']) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_unallocated.iterrows())

                tot_obras_prontas = sum(len(r['_ORIGINAL_ROWS']) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_tasks_alocadas.iterrows())
                sidebar_html_placeholder.markdown(renderizar_painel_lateral(cap_por_eq_live, tot_obras_prontas, qtd_eq_atual_live, cap_total_estimada_live), unsafe_allow_html=True)

                if df_tasks_alocadas.empty: 
                    st.error("Nenhuma obra encontrou equipes com cobertura geográfica (num raio de 100km) ou com limite diário disponível.")
                    return
                bases_records = todas_bases_records 

            if has_generica: col_prioridade = st.session_state.get('col_prioridade_gen', "Nenhuma")
            elif has_levantamento: col_prioridade = st.session_state.get('col_prioridade_lev', "Nenhuma")
            else: col_prioridade = "Nenhuma"

        else:
            st.markdown("### 📥 1. Planilha de Demanda (Lista Contínua)")
            st.info("Neste modo, o sistema apenas lê as colunas **LEVANTADOR**, **REGIONAL** e **MUNICIPIO** da sua planilha. Nenhuma equipe receberá obras de outro levantador. A IA vai roteirizar 100% da lista ignorando o limite de dias.")
            
            with st.container(border=True):
                pre_file = st.file_uploader("Planilha de Obras", type=["xlsx", "xls", "csv"], help="A planilha deve conter LEVANTADOR, MUNICIPIO, LATITUDE e LONGITUDE.")
            
            ignorar_despacho = st.checkbox("Filtro: Ignorar obras já despachadas (com DATA DESPACHO CAMPO)?", value=False, help="Se marcado, o sistema não roteirizará obras que já tenham data preenchida. (Deixe desmarcado para roteirizar tudo).")

            if pre_file:
                df_tasks = ler_planilha_cached(pre_file.getvalue()) if not pre_file.name.endswith('.csv') else pd.read_csv(pre_file)
                st.session_state.colunas_originais = df_tasks.columns.tolist()
                df_tasks.columns = normalize_cols(df_tasks.columns)
                df_tasks['_ORIGEM_BASE'] = 'LISTA_CONTINUA'
                if 'PRIORIDADE' in df_tasks.columns: df_tasks['_PRIORIDADE_ORIGINAL'] = df_tasks['PRIORIDADE']
                
                if 'LEVANTADOR' not in df_tasks.columns:
                    if 'NOME DO LEVANTADOR' in df_tasks.columns: df_tasks.rename(columns={'NOME DO LEVANTADOR': 'LEVANTADOR'}, inplace=True)
                    else: st.error("🚨 A planilha precisa da coluna 'LEVANTADOR'."); st.stop()
                if 'MUNICIPIO' not in df_tasks.columns:
                    if 'CIDADE' in df_tasks.columns: df_tasks.rename(columns={'CIDADE': 'MUNICIPIO'}, inplace=True)
                    else: st.error("🚨 A planilha precisa da coluna 'MUNICIPIO'."); st.stop()
                if 'LATITUDE' not in df_tasks.columns or 'LONGITUDE' not in df_tasks.columns:
                    st.error("🚨 A planilha precisa das colunas 'LATITUDE' e 'LONGITUDE'."); st.stop()
                    
                if 'PROTOCOLO' not in df_tasks.columns:
                    for col_candidata in ['NOTA', 'NOTA CCS', 'NOTA SGO', 'ID SISCO', 'OS']:
                        if col_candidata in df_tasks.columns:
                            df_tasks['PROTOCOLO'] = df_tasks[col_candidata]
                            break
                    if 'PROTOCOLO' not in df_tasks.columns:
                        df_tasks['PROTOCOLO'] = [f"LC-{i+1}" for i in range(len(df_tasks))]
                        
                for c_nome in ['CONTA CONTRATO', 'INSTALACAO', 'PROTOCOLO']:
                    if c_nome in df_tasks.columns: df_tasks[c_nome] = df_tasks[c_nome].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '-')
                
                if ignorar_despacho and 'DATA DESPACHO CAMPO' in df_tasks.columns:
                    mask_despacho = df_tasks['DATA DESPACHO CAMPO'].notna() & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip() != '') & (df_tasks['DATA DESPACHO CAMPO'].astype(str).str.strip().str.lower() != 'nan')
                    obras_despachadas = mask_despacho.sum()
                    if obras_despachadas > 0:
                        st.info(f"⏭️ {obras_despachadas} obras foram ignoradas (DATA DESPACHO CAMPO preenchida).")
                        df_tasks = df_tasks[~mask_despacho]
                        
                df_tasks['LATITUDE'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                df_tasks['LONGITUDE'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).str.replace(',', '.'), errors='coerce')
                
                erros_coords_mask = df_tasks['LATITUDE'].isna() | df_tasks['LONGITUDE'].isna() | (df_tasks['LATITUDE'] == 0.0) | (df_tasks['LONGITUDE'] == 0.0)
                qtd_erros_coords_finais = erros_coords_mask.sum()
                df_tasks = df_tasks[~erros_coords_mask]
                
                if qtd_erros_coords_finais > 0:
                    st.toast(f"⚠️ {qtd_erros_coords_finais} obras ignoradas por falta de coordenadas válidas (vazias ou 0.0).")
                
                if 'NOME' not in df_tasks.columns: df_tasks['NOME'] = "SEM NOME"
                
                df_tasks['LEVANTADOR'] = df_tasks['LEVANTADOR'].astype(str).str.strip().str.upper()
                lixos_lev = ['NAN', 'NONE', '', '-', 'SEM LEVANTADOR', '0', '0.0', 'N/A', 'NULO']
                df_tasks = df_tasks[~df_tasks['LEVANTADOR'].isin(lixos_lev)]

                df_tasks, qtd_condensada = fundir_super_pontos(df_tasks, raio_metros=st.session_state.get('vrp_state', {}).get('config', {}).get('raio_super_ponto', 100) if 'vrp_state' in st.session_state else 100, agrupar_por_levantador=True)
                if qtd_condensada > 0: st.toast(f"✅ {qtd_condensada} obras repetidas no mesmo endereço viraram 'Super Pontos'.")
                
                if 'PRIORIDADE' not in df_tasks.columns:
                    df_tasks['PRIORIDADE'] = 'Não'
                else:
                    df_tasks['PRIORIDADE'] = df_tasks['PRIORIDADE'].astype(str).str.strip().str.upper().apply(lambda x: 'Sim' if x == 'SIM' else 'Não')
                
                # --- REGRA DE OURO 1: CORREÇÃO DE LEVANTAMENTO É SEMPRE PRIORIDADE ---
                if 'STATUS LIST' in df_tasks.columns:
                    mask_correcao = df_tasks['STATUS LIST'].astype(str).str.strip().str.upper().isin(['CORREÇÃO DE LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO'])
                    df_tasks.loc[mask_correcao, 'PRIORIDADE'] = 'Sim'
                # -------------------------------------------------------------------
                
                # --- REGRA DE OURO 2: COLUNA "PRIORIDADE" DIFERENTE DE ZERO ---
                if '_PRIORIDADE_ORIGINAL' in df_tasks.columns:
                    mask_not_zero = (
                        df_tasks['_PRIORIDADE_ORIGINAL'].notna() & 
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '') & 
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '0') & 
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip() != '0.0') & 
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NAN') &
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NÃO') &
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'NAO') &
                        (df_tasks['_PRIORIDADE_ORIGINAL'].astype(str).str.strip().str.upper() != 'FALSE')
                    )
                    df_tasks.loc[mask_not_zero, 'PRIORIDADE'] = 'Sim'
                # -------------------------------------------------------------------

                if df_tasks.empty:
                    st.error("🚨 Nenhuma obra restou após os filtros de coordenadas. Verifique sua planilha.")
                else:
                    df_tasks['BASE_ATRIBUIDA'] = df_tasks['LEVANTADOR']
                    df_tasks_alocadas = df_tasks.copy()
                    
                    st.session_state.df_unallocated = pd.DataFrame()
                    st.session_state.tot_obras_nao_alocadas = 0
                    
                    bases_records = []
                    for lev in df_tasks_alocadas['LEVANTADOR'].unique():
                        df_lev = df_tasks_alocadas[df_tasks_alocadas['LEVANTADOR'] == lev]
                        mun_base = df_lev['MUNICIPIO'].mode().iloc[0] if not df_lev['MUNICIPIO'].dropna().empty else "DESCONHECIDO"
                        reg_base = df_lev['REGIONAL'].iloc[0] if 'REGIONAL' in df_lev.columns else "DESCONHECIDO"
                        lat, lon = obter_coordenadas_municipio_cached(mun_base)
                        
                        if pd.isna(lat) or pd.isna(lon):
                            lat = df_lev['LATITUDE'].iloc[0]
                            lon = df_lev['LONGITUDE'].iloc[0]
                            
                        bases_records.append({
                            'LEVANTADOR': lev,
                            'RESIDENCIA': mun_base,
                            'MUNICIPIO': mun_base,
                            'REGIONAL': reg_base,
                            'LATITUDE': lat,
                            'LONGITUDE': lon,
                            'TIPO_EQUIPE': 'LISTA_CONTINUA'
                        })
                    
                    tot_obras_prontas = sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_tasks_alocadas.iterrows())
                    sidebar_html_placeholder.markdown(renderizar_painel_lateral("Ilimitado", tot_obras_prontas, len(bases_records), "Ilimitado"), unsafe_allow_html=True)
                    st.success(f"✅ Planilha carregada! {len(df_tasks_alocadas)} paradas identificadas para {len(bases_records)} levantadores.")
                    
                    col_prioridade = "PRIORIDADE"

        if not df_tasks_alocadas.empty:
            with st.expander("🛠️ 5. Configuração de Saída", expanded=True):
                todas_cols = df_tasks_alocadas.columns.tolist()
                todas_cols_limpas = [c for c in todas_cols if not c.startswith('_')]
                
                cols_desejadas = ['PROTOCOLO', 'CONTA CONTRATO', 'INSTALACAO', 'NOME', 'ENDERECO', 'INFORMACOES EXTRAS', 'LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'TIPO NOTA', 'FASE']

                cols_padrao = [c for c in cols_desejadas if c in todas_cols_limpas]
                
                colunas_exibir = st.multiselect("Colunas Visíveis nos Cartões (KML/Mapa)", todas_cols_limpas, default=cols_padrao)
                
                reference_order = cols_desejadas + [c for c in todas_cols_limpas if c not in cols_desejadas]
                colunas_exibir.sort(key=lambda x: reference_order.index(x) if x in reference_order else 999)
                
                st.info("⚡ **Deduplicação Ativa:** Obras baseadas no raio definido no menu lateral foram condensadas para otimização.")

            if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
                tipo_periodo_clean = "Semana" if "Semana" in tipo_periodo else "Dia"
                if tipo_periodo_clean == "Semana" and not dias_semana_selecionados:
                    st.error("Selecione os dias da semana na barra lateral antes de continuar.")
                    return

                st.session_state.tarefas_alocadas_inicialmente = len(df_tasks_alocadas)
                st.session_state.bases_records = bases_records
                st.session_state.tipo_periodo = tipo_periodo_clean
                st.session_state.colunas_exibir = colunas_exibir
                st.session_state.col_prioridade = col_prioridade
                
                st.session_state.config_financeira = {
                    'custo_combustivel': custo_combustivel,
                    'consumo_veiculo': consumo_veiculo,
                    'custo_hora_equipe': custo_hora_equipe
                }
                
                keys_to_clear = ['start_time_run', 'tempo_processamento', 'start_time_pkg']
                for k in keys_to_clear:
                    if k in st.session_state: del st.session_state[k]
                
                st.session_state.vrp_state = {
                    'config': {
                        'velocidade_media_kmh': velocidade_media_kmh,
                        'tempo_medio_obra': tempo_medio_obra, 
                        'obras_por_dia': obras_por_dia, 
                        'tipo_periodo': tipo_periodo_clean, 
                        'limite_periodos': limite_periodos,
                        'roteirizar_tudo': True if modo_operacao == "2" else False,
                        'is_lista_continua': True if modo_operacao == "2" else False,
                        'dias_selecionados': dias_semana_selecionados,
                        'url_osrm_base': url_osrm_base,
                        'tracado_real': tracado_real,
                        'data_inicio': data_inicio_roteiro,
                        'sentido_rota': sentido_rota,
                        'raio_super_ponto': raio_super_ponto,
                        'trava_global_obras': trava_global_obras
                    },
                    'b_names': list(set([b['LEVANTADOR'] for b in bases_records])),
                    'b_idx': 0, 'unvisited': df_tasks_alocadas.copy(), 'routed_data': [],
                }
                st.session_state.vrp_status = "RUNNING"
                tentar_rerun()

    if status_exec in ["RUNNING"]:
        state = st.session_state.vrp_state
        cfg = state['config']
        is_lista_continua = cfg.get('is_lista_continua', False)
        is_reversa = "Reversa" in cfg.get('sentido_rota', '')
        
        titulo_motor = "🚀 Execução do Motor Leve Vetorial (Lista Contínua)" if is_lista_continua else "🚀 Execução do Motor de Inteligência (OR-Tools VRP)"
        st.markdown(f"## {titulo_motor}")
        st.markdown("Calculando Matrizes Vetoriais e Otimizando Rotas...")
        
        if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
            
        b_names = state['b_names']
        total_equipes = len(b_names)
        b_idx = state.get('b_idx', 0)
        
        if 'start_time_run' not in st.session_state:
            st.session_state.start_time_run = time.time()
        if 'tempo_processamento' not in st.session_state:
            st.session_state.tempo_processamento = 0.0
            
        global_start_time = st.session_state.start_time_run
        
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        timer_placeholder = st.empty()

        def update_running_timer(b_index, inner_idx, inner_total):
            elapsed = time.time() - global_start_time
            fraction = (b_index + (inner_idx / max(1, inner_total))) / max(1, total_equipes)
            
            if fraction > 0.02: 
                est_total = elapsed / fraction
                rem = max(0, est_total - elapsed)
                r_m, r_s = divmod(int(rem), 60)
                r_str = f"{r_m:02d}m {r_s:02d}s"
            else:
                r_str = "Calculando..."
                
            e_m, e_s = divmod(int(elapsed), 60)
            e_str = f"{e_m:02d}m {e_s:02d}s"
            
            html_timer = f"""
            <div style="display: flex; gap: 10px; margin-bottom: 15px;">
                <div style="flex: 1; padding: 15px; border-radius: 8px; background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border: 1px solid #dee2e6; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="font-size: 0.85rem; color: #6c757d; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">⏱️ Tempo Decorrido</div>
                    <div style="font-size: 1.8rem; font-weight: 800; color: #0D256C; margin-top: 5px; font-variant-numeric: tabular-nums;">{e_str}</div>
                </div>
                <div style="flex: 1; padding: 15px; border-radius: 8px; background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border: 1px solid #a5d6a7; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="font-size: 0.85rem; color: #2e7d32; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">🎯 Estimativa Restante</div>
                    <div style="font-size: 1.8rem; font-weight: 800; color: #1b5e20; margin-top: 5px; font-variant-numeric: tabular-nums;">{r_str}</div>
                </div>
            </div>
            """
            timer_placeholder.markdown(html_timer, unsafe_allow_html=True)

        if b_idx < total_equipes:
            b_name = b_names[b_idx]
            start_iter = time.time()
            
            progresso = b_idx / max(1, total_equipes)
            progress_bar.progress(progresso)
            
            texto_acao = "Mapeando e sequenciando" if is_lista_continua else "IA Analisando nós e traçando rotas para"
            status_text.info(f"🧠 {texto_acao} **{b_name}**... ({b_idx + 1}/{total_equipes})")
            
            update_running_timer(b_idx, 0, 1)
            
            if 'current_rotas_flat' not in state:
                df_todas_bases_ativas = pd.DataFrame(st.session_state.bases_records)
                unvisited = state['unvisited']
                
                base_ref = df_todas_bases_ativas[df_todas_bases_ativas['LEVANTADOR'] == b_name].iloc[0]
                if pd.isna(base_ref.get('LATITUDE')): 
                    state['b_idx'] += 1
                    st.session_state.vrp_state = state
                    tentar_rerun()
                    return
                    
                base_lat, base_lon = float(base_ref['LATITUDE']), float(base_ref['LONGITUDE'])
                obras_equipe = unvisited[unvisited['BASE_ATRIBUIDA'] == b_name].to_dict('records')
                
                ordered_tasks = []
                if obras_equipe:
                    if is_lista_continua:
                        mun_groups = {}
                        for o in obras_equipe:
                            mun_raw = o.get('MUNICIPIO', o.get('CIDADE', 'DESCONHECIDO'))
                            mun_limpo = normalizar_municipios(pd.Series([mun_raw])).iloc[0] if pd.notna(mun_raw) else 'DESCONHECIDO'
                            o['MUN_LIMPO_CALC'] = mun_limpo
                            if mun_limpo not in mun_groups: mun_groups[mun_limpo] = []
                            mun_groups[mun_limpo].append(o)
                            
                        for mun, obs in mun_groups.items():
                            prio_sim = [o for o in obs if str(o.get('PRIORIDADE')).upper() == 'SIM']
                            prio_nao = [o for o in obs if str(o.get('PRIORIDADE')).upper() != 'SIM']
                            
                            def greedy_sort(pts, start_lat, start_lon, reverse_first=False):
                                if not pts: return []
                                sorted_pts = []
                                curr_lat, curr_lon = start_lat, start_lon
                                unvisited_pts = list(pts)
                                
                                if reverse_first and unvisited_pts:
                                    best_idx = 0
                                    best_d = -1
                                    for i, p in enumerate(unvisited_pts):
                                        d = haversine_scalar(curr_lat, curr_lon, p['LATITUDE'], p['LONGITUDE'])
                                        if d > best_d:
                                            best_d = d; best_idx = i
                                    nxt = unvisited_pts.pop(best_idx)
                                    sorted_pts.append(nxt)
                                    curr_lat, curr_lon = nxt['LATITUDE'], nxt['LONGITUDE']
                                    
                                while unvisited_pts:
                                    best_idx = 0
                                    best_d = float('inf')
                                    for i, p in enumerate(unvisited_pts):
                                        d = haversine_scalar(curr_lat, curr_lon, p['LATITUDE'], p['LONGITUDE'])
                                        if d < best_d:
                                            best_d = d; best_idx = i
                                    nxt = unvisited_pts.pop(best_idx)
                                    sorted_pts.append(nxt)
                                    curr_lat, curr_lon = nxt['LATITUDE'], nxt['LONGITUDE']
                                return sorted_pts

                            sub_sim = greedy_sort(prio_sim, base_lat, base_lon, is_reversa)
                            ordered_tasks.extend(sub_sim)
                            if ordered_tasks:
                                last = ordered_tasks[-1]
                                ordered_tasks.extend(greedy_sort(prio_nao, last['LATITUDE'], last['LONGITUDE'], False))
                            else:
                                ordered_tasks.extend(greedy_sort(prio_nao, base_lat, base_lon, is_reversa))

                    else:
                        coords_dict = {}
                        for o in obras_equipe:
                            k = (round(float(o['LATITUDE']), 4), round(float(o['LONGITUDE']), 4))
                            if k not in coords_dict: coords_dict[k] = []
                            coords_dict[k].append(o)
                            
                        macro_obras = []
                        for k, lista in coords_dict.items():
                            tem_prio = any(x.get('PRIORIDADE') == 'Sim' for x in lista)
                            rep = lista[0].copy()
                            rep['PRIORIDADE'] = 'Sim' if tem_prio else 'Não'
                            mun_raw = rep.get('MUNICIPIO', rep.get('CIDADE', 'DESCONHECIDO'))
                            rep['MUN_LIMPO'] = normalizar_municipios(pd.Series([mun_raw])).iloc[0] if pd.notna(mun_raw) and str(mun_raw).strip() != '' else 'DESCONHECIDO'
                            rep['_sub_obras'] = lista
                            macro_obras.append(rep)

                        macros_by_mun = {}
                        for m in macro_obras:
                            mun = m['MUN_LIMPO']
                            if mun not in macros_by_mun: macros_by_mun[mun] = []
                            macros_by_mun[mun].append(m)

                        mun_stats = []
                        for mun, m_list in macros_by_mun.items():
                            prio_count = sum(1 for x in m_list if x['PRIORIDADE'] == 'Sim')
                            mun_stats.append({'mun': mun, 'prio_count': prio_count, 'total_count': len(m_list)})

                        mun_stats.sort(key=lambda x: (x['prio_count'] > 0, x['prio_count'], x['total_count']), reverse=True)

                        ordered_macros = []
                        for stat in mun_stats:
                            mun = stat['mun']
                            m_list = macros_by_mun[mun]
                            prio_macros = [m for m in m_list if m['PRIORIDADE'] == 'Sim']
                            comum_macros = [m for m in m_list if m['PRIORIDADE'] != 'Sim']
                            
                            if prio_macros: 
                                tsp_res = resolver_tsp_ortools(prio_macros, base_lat, base_lon, cfg['url_osrm_base'])
                                if is_reversa and tsp_res:
                                    farthest_idx = max(range(len(tsp_res)), key=lambda i: haversine_scalar(base_lat, base_lon, float(tsp_res[i]['LATITUDE']), float(tsp_res[i]['LONGITUDE'])))
                                    tsp_res = tsp_res[farthest_idx:] + tsp_res[:farthest_idx]
                                ordered_macros.extend(tsp_res)
                                
                            if comum_macros: 
                                tsp_res = resolver_tsp_ortools(comum_macros, base_lat, base_lon, cfg['url_osrm_base'])
                                if is_reversa and tsp_res:
                                    farthest_idx = max(range(len(tsp_res)), key=lambda i: haversine_scalar(base_lat, base_lon, float(tsp_res[i]['LATITUDE']), float(tsp_res[i]['LONGITUDE'])))
                                    tsp_res = tsp_res[farthest_idx:] + tsp_res[:farthest_idx]
                                ordered_macros.extend(tsp_res)
                            
                        for macro in ordered_macros:
                            subs = sorted(macro['_sub_obras'], key=lambda x: 0 if x.get('PRIORIDADE') == 'Sim' else 1)
                            for s in subs: s['MUN_LIMPO_CALC'] = macro['MUN_LIMPO']
                            ordered_tasks.extend(subs)
                
                rotas_flat = []
                dia_absoluto = 1
                semana_atual = 1
                dia_da_semana = 1
                obras_no_periodo_macro = 0
                mun_anterior = None
                
                agora_dt = datetime.combine(cfg['data_inicio'], datetime.min.time())
                data_base_inicio = agora_dt.replace(hour=8, minute=0, second=0, microsecond=0)
                
                def get_workday_date(start_dt, dia_abs, valid_days_list):
                    pt_to_idx = {"Segunda": 0, "Terça": 1, "Quarta": 2, "Quinta": 3, "Sexta": 4, "Sábado": 5, "Domingo": 6}
                    allowed = [pt_to_idx[d] for d in valid_days_list] if valid_days_list else [0,1,2,3,4]
                    curr = start_dt
                    while curr.weekday() not in allowed:
                        curr += pd.Timedelta(days=1)
                    count = 1
                    while count < dia_abs:
                        curr += pd.Timedelta(days=1)
                        if curr.weekday() in allowed:
                            count += 1
                    return curr

                def iniciar_dia(dia_abs):
                    data_atual = get_workday_date(data_base_inicio, dia_abs, cfg['dias_selecionados'])
                    return {
                        'lat': base_lat, 'lon': base_lon,
                        'time': data_atual, 'date_obj': data_atual,
                        'obras_hoje': 0, 'km_hoje': 0.0, 'lunch': False
                    }
                    
                estado = iniciar_dia(dia_absoluto)
                
                for obra in ordered_tasks:
                    mun_atual = obra.get('MUN_LIMPO_CALC', 'DESCONHECIDO')
                    qtd_real = len(obra.get('_ORIGINAL_ROWS', [1])) if isinstance(obra.get('_ORIGINAL_ROWS'), list) else 1
                    
                    viagem_km_reta = haversine_vectorized(estado['lat'], estado['lon'], obra['LATITUDE'], obra['LONGITUDE'])
                    viagem_km = viagem_km_reta * 1.3 
                    obra['ALERTA_TOPOLOGIA'] = 'OK'
                    
                    if viagem_km_reta < 0.05 and estado['obras_hoje'] > 0:
                        viagem_min = 0.0
                        exec_min = 30.0 
                    else:
                        vel_dinamica = cfg['velocidade_media_kmh'] * 1.5 if viagem_km > 20 else cfg['velocidade_media_kmh']
                        viagem_min = (viagem_km / vel_dinamica) * 60
                        exec_min = cfg['tempo_medio_obra'] * 60
                    
                    chegada_prevista = estado['time'] + pd.Timedelta(minutes=viagem_min)
                    
                    if chegada_prevista.hour >= 12 and not estado['lunch']:
                        lunch_start = max(estado['time'], estado['date_obj'].replace(hour=12))
                        lunch_end = lunch_start + pd.Timedelta(hours=1)
                        rotas_flat.append({
                            'obra': None, 'is_lunch': True, 'is_retorno': False,
                            'lat_ant': estado['lat'], 'lon_ant': estado['lon'],
                            'lat_atual': estado['lat'], 'lon_atual': estado['lon'], 
                            'semana': semana_atual, 'dia': dia_absoluto, 'dia_semana_idx': dia_da_semana,
                            'dia_mes': estado['date_obj'].strftime('%d/%m/%Y'),
                            'hora_inicio': lunch_start, 'hora_fim': lunch_end,
                            'viagem_min': 0.0, 'dist_km': 0.0
                        })
                        estado['time'] = lunch_end
                        estado['lunch'] = True
                        chegada_prevista = estado['time'] + pd.Timedelta(minutes=viagem_min)
                        
                    fim_previsto = chegada_prevista + pd.Timedelta(minutes=exec_min)
                    
                    virar_dia = False
                    limite_diario_atual = cfg['obras_por_dia']
                    
                    if estado['obras_hoje'] > 0 and (estado['obras_hoje'] + qtd_real > limite_diario_atual):
                        virar_dia = True
                    elif mun_anterior is not None and mun_atual != mun_anterior and estado['obras_hoje'] > 0:
                        if viagem_km_reta > 100.0:
                            virar_dia = True 
                            
                    if virar_dia:
                        dist_ret = haversine_vectorized(estado['lat'], estado['lon'], base_lat, base_lon)
                        viagem_ret = (dist_ret / cfg['velocidade_media_kmh']) * 60
                        ret_fim = estado['time'] + pd.Timedelta(minutes=viagem_ret)
                        rotas_flat.append({
                            'obra': None, 'is_lunch': False, 'is_retorno': True,
                            'lat_ant': estado['lat'], 'lon_ant': estado['lon'],
                            'lat_atual': base_lat, 'lon_atual': base_lon,
                            'semana': semana_atual, 'dia': dia_absoluto, 'dia_semana_idx': dia_da_semana,
                            'dia_mes': estado['date_obj'].strftime('%d/%m/%Y'),
                            'hora_inicio': estado['time'], 'hora_fim': ret_fim,
                            'viagem_min': viagem_ret, 'dist_km': dist_ret
                        })
                        
                        dia_absoluto += 1
                        
                        if cfg['tipo_periodo'] == "Semana":
                            dia_da_semana += 1
                            if dia_da_semana > len(cfg['dias_selecionados']):
                                semana_atual += 1
                                dia_da_semana = 1
                                
                        estado = iniciar_dia(dia_absoluto)
                        
                        viagem_km_reta = haversine_vectorized(estado['lat'], estado['lon'], obra['LATITUDE'], obra['LONGITUDE'])
                        viagem_km = viagem_km_reta * 1.3
                        
                        if viagem_km_reta < 0.05 and estado['obras_hoje'] > 0:
                            viagem_min = 0.0; exec_min = 30.0 
                        else:
                            vel_dinamica = cfg['velocidade_media_kmh'] * 1.5 if viagem_km > 20 else cfg['velocidade_media_kmh']
                            viagem_min = (viagem_km / vel_dinamica) * 60; exec_min = cfg['tempo_medio_obra'] * 60
                        
                        chegada_prevista = estado['time'] + pd.Timedelta(minutes=viagem_min)
                        fim_previsto = chegada_prevista + pd.Timedelta(minutes=exec_min)
                        
                    rotas_flat.append({
                        'obra': obra, 'is_lunch': False, 'is_retorno': False,
                        'lat_ant': estado['lat'], 'lon_ant': estado['lon'],
                        'lat_atual': obra['LATITUDE'], 'lon_atual': obra['LONGITUDE'],
                        'semana': semana_atual, 'dia': dia_absoluto, 'dia_semana_idx': dia_da_semana,
                        'dia_semana_nome': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][estado['date_obj'].weekday()],
                        'dia_mes': estado['date_obj'].strftime('%d/%m/%Y'),
                        'hora_inicio': chegada_prevista, 'hora_fim': fim_previsto,
                        'viagem_min': viagem_min, 'dist_km': viagem_km
                    })
                    estado['lat'] = obra['LATITUDE']
                    estado['lon'] = obra['LONGITUDE']
                    estado['time'] = fim_previsto
                    estado['obras_hoje'] += qtd_real
                    estado['km_hoje'] += viagem_km
                    mun_anterior = mun_atual 

                if estado['obras_hoje'] > 0:
                    dist_ret = haversine_vectorized(estado['lat'], estado['lon'], base_lat, base_lon)
                    viagem_ret = (dist_ret / cfg['velocidade_media_kmh']) * 60
                    ret_fim = estado['time'] + pd.Timedelta(minutes=viagem_ret)
                    rotas_flat.append({
                        'obra': None, 'is_lunch': False, 'is_retorno': True,
                        'lat_ant': estado['lat'], 'lon_ant': estado['lon'],
                        'lat_atual': base_lat, 'lon_atual': base_lon,
                        'semana': semana_atual, 'dia': dia_absoluto, 'dia_semana_idx': dia_da_semana,
                        'dia_semana_nome': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][estado['date_obj'].weekday()],
                        'dia_mes': estado['date_obj'].strftime('%d/%m/%Y'),
                        'hora_inicio': estado['time'], 'hora_fim': ret_fim,
                        'viagem_min': viagem_ret, 'dist_km': dist_ret
                    })

                state['current_rotas_flat'] = rotas_flat
                state['current_osrm_idx'] = 0
                state['current_geoms'] = []
                st.session_state.vrp_state = state
                tentar_rerun()
                return

            else:
                rotas_flat = state['current_rotas_flat']
                osrm_idx = state['current_osrm_idx']
                geoms_and_durs = state['current_geoms']
                
                batch_size = 30 if cfg.get('tracado_real', False) else len(rotas_flat)
                end_idx = min(osrm_idx + batch_size, len(rotas_flat))
                
                for i in range(osrm_idx, end_idx):
                    item = rotas_flat[i]
                    if not cfg.get('tracado_real', False):
                        dist_m = item['dist_km'] * 1000
                        dur_sec = (dist_m / 1000.0 / cfg['velocidade_media_kmh']) * 3600
                        geom = [[item['lon_ant'], item['lat_ant']], [item['lon_atual'], item['lat_atual']]]
                        geoms_and_durs.append((geom, dur_sec))
                    else:
                        if i % 5 == 0:
                            status_text.info(f"🛣️ Desenhando curvas reais para **{b_name}**... (Trecho {i}/{len(rotas_flat)})")
                        update_running_timer(b_idx, i, len(rotas_flat))
                        time.sleep(0.05) 
                        try:
                            geom, dur_sec = obter_rota_ruas(item['lat_ant'], item['lon_ant'], item['lat_atual'], item['lon_atual'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'])
                        except:
                            dist_m = item['dist_km'] * 1000
                            geom = [[item['lon_ant'], item['lat_ant']], [item['lon_atual'], item['lat_atual']]]
                            dur_sec = (dist_m / 1000.0 / cfg['velocidade_media_kmh']) * 3600
                        geoms_and_durs.append((geom, dur_sec))

                state['current_osrm_idx'] = end_idx
                state['current_geoms'] = geoms_and_durs
                
                if end_idx < len(rotas_flat):
                    st.session_state.vrp_state = state
                    tentar_rerun()
                    return
                
                routed_data_final_team = []
                ordem_global = 1
                dias_pt = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
                for item, (geom, dur_sec) in zip(rotas_flat, geoms_and_durs):
                    periodo_val = item['semana'] if cfg['tipo_periodo'] == "Semana" else item['dia']
                    
                    data_do_item = datetime.strptime(item['dia_mes'], '%d/%m/%Y')
                    dia_nome_str = dias_pt[data_do_item.weekday()] if cfg['tipo_periodo'] == "Semana" else f"Dia {item['dia']}"
                    
                    if item['is_lunch']:
                        routed_data_final_team.append({
                            'PROTOCOLO': 'PAUSA_ALMOCO', 'NOME': '🍔 ALMOÇO DA EQUIPE', 
                            'LATITUDE': item['lat_atual'], 'LONGITUDE': item['lon_atual'],
                            'BASE_ATRIBUIDA': b_name, 'ORDEM': ordem_global, 
                            'NOME_DIA': dia_nome_str, 'DIA_MES': item['dia_mes'],
                            'SEMANA': item['semana'], 'DIA': item['dia'], 
                            'PERIODO': periodo_val,
                            'DISTANCIA_PONTO_ANTERIOR_KM': 0.0, 'TEMPO_VIAGEM_MINUTOS': 0.0,
                            'ROTA_GEOMETRIA': geom,
                            'PRIORIDADE': 'Não',
                            'HORA_INICIO': item['hora_inicio'].strftime('%H:%M'),
                            'HORA_FIM': item['hora_fim'].strftime('%H:%M'),
                            '_HORA_INICIO_DT': item['hora_inicio'], '_HORA_FIM_DT': item['hora_fim']
                        })
                    elif item['is_retorno']:
                        routed_data_final_team.append({
                            'PROTOCOLO': 'RETORNO_BASE', 'NOME': 'BASE_RETORNO', 
                            'LATITUDE': item['lat_atual'], 'LONGITUDE': item['lon_atual'],
                            'BASE_ATRIBUIDA': b_name, 'ORDEM': ordem_global, 
                            'NOME_DIA': dia_nome_str, 'DIA_MES': item['dia_mes'],
                            'SEMANA': item['semana'], 'DIA': item['dia'], 
                            'PERIODO': periodo_val,
                            'DISTANCIA_PONTO_ANTERIOR_KM': round(item['dist_km'], 2), 
                            'TEMPO_VIAGEM_MINUTOS': round(item['viagem_min'], 1),
                            'ROTA_GEOMETRIA': geom,
                            'PRIORIDADE': 'Não',
                            'HORA_INICIO': item['hora_inicio'].strftime('%H:%M'),
                            'HORA_FIM': item['hora_fim'].strftime('%H:%M'),
                            '_HORA_INICIO_DT': item['hora_inicio'], '_HORA_FIM_DT': item['hora_fim']
                        })
                    else:
                        obra = item['obra']
                        obra['ORDEM'] = ordem_global
                        obra['NOME_DIA'] = dia_nome_str
                        obra['DIA_MES'] = item['dia_mes']
                        obra['SEMANA'] = item['semana']
                        obra['DIA'] = item['dia']
                        obra['PERIODO'] = periodo_val
                        obra['DISTANCIA_PONTO_ANTERIOR_KM'] = round(item['dist_km'], 2)
                        obra['TEMPO_VIAGEM_MINUTOS'] = round(item['viagem_min'], 1)
                        obra['ROTA_GEOMETRIA'] = geom
                        obra['HORA_INICIO'] = item['hora_inicio'].strftime('%H:%M')
                        obra['HORA_FIM'] = item['hora_fim'].strftime('%H:%M')
                        obra['_HORA_INICIO_DT'] = item['hora_inicio']
                        obra['_HORA_FIM_DT'] = item['hora_fim']
                        
                        routed_data_final_team.append(obra)
                    ordem_global += 1

                state['routed_data'].extend(routed_data_final_team)
                del state['current_rotas_flat']
                del state['current_osrm_idx']
                del state['current_geoms']
                state['b_idx'] += 1
                
                st.session_state.tempo_processamento += (time.time() - start_iter)
                st.session_state.vrp_state = state
                gc.collect() 
                tentar_rerun()
                return

        else:
            status_text.success("✅ Matrizes Resolvidas! Preparando empacotamento...")
            progress_bar.progress(1.0)
            
            df_final_route = pd.DataFrame(state['routed_data'])
            if not df_final_route.empty:
                df_final_route['DISTANCIA_PROXIMO_PONTO_KM'] = df_final_route.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
                
            st.session_state.df_routed = df_final_route
            st.session_state.vrp_status = "PACKAGING"
            time.sleep(1)
            tentar_rerun()

    if status_exec == "PACKAGING":
        st.markdown("## 📦 Etapa Final: Construção de Arquivos (Excel e KML)")
        st.markdown("A inteligência já finalizou as rotas. Compilando os dados para o download...")
        
        if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
            
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        timer_placeholder = st.empty()
        
        df_routed = st.session_state.df_routed
        data_atual_formatada = datetime.now().strftime("%d.%m.%Y")
        
        bases_unicas = df_routed['BASE_ATRIBUIDA'].unique().tolist()
        
        total_steps = len(bases_unicas) * 2 + 3
        current_step = 0
        
        if 'start_time_pkg' not in st.session_state:
            st.session_state.start_time_pkg = time.time()
        start_time = st.session_state.start_time_pkg
        
        buf_zip_xl = io.BytesIO()
        buf_zip_kml = io.BytesIO()
        buf_zip_gpx = io.BytesIO()
        
        tipo_periodo_atual = st.session_state.vrp_state.get('config', {}).get('tipo_periodo', 'Dia')
        
        try:
            with zipfile.ZipFile(buf_zip_xl, 'w', zipfile.ZIP_DEFLATED) as zip_xl, \
                 zipfile.ZipFile(buf_zip_kml, 'w', zipfile.ZIP_DEFLATED) as zip_kml, \
                 zipfile.ZipFile(buf_zip_gpx, 'w', zipfile.ZIP_DEFLATED) as zip_gpx:
                 
                def update_ui(msg):
                    nonlocal current_step
                    current_step += 1
                    progresso = min(current_step / total_steps, 1.0)
                    progress_bar.progress(progresso)
                    status_text.info(f"⏳ {msg}")
                    
                    elapsed = time.time() - start_time
                    avg_time = elapsed / current_step if current_step > 0 else 0
                    rem_time = max(0, avg_time * (total_steps - current_step))
                    
                    e_m, e_s = divmod(int(elapsed), 60)
                    r_m, r_s = divmod(int(rem_time), 60)
                    
                    html_timer = f"""
                    <div style="display: flex; gap: 10px; margin-bottom: 15px;">
                        <div style="flex: 1; padding: 15px; border-radius: 8px; background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border: 1px solid #dee2e6; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                            <div style="font-size: 0.85rem; color: #6c757d; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">⏱️ Tempo Decorrido</div>
                            <div style="font-size: 1.8rem; font-weight: 800; color: #0D256C; margin-top: 5px; font-variant-numeric: tabular-nums;">{e_m:02d}m {e_s:02d}s</div>
                        </div>
                        <div style="flex: 1; padding: 15px; border-radius: 8px; background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border: 1px solid #a5d6a7; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                            <div style="font-size: 0.85rem; color: #2e7d32; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">🗂️ Empacotando (Restante)</div>
                            <div style="font-size: 1.8rem; font-weight: 800; color: #1b5e20; margin-top: 5px; font-variant-numeric: tabular-nums;">{r_m:02d}m {r_s:02d}s</div>
                        </div>
                    </div>
                    """
                    timer_placeholder.markdown(html_timer, unsafe_allow_html=True)

                update_ui("Gerando Painel de Resumo Operacional...")
                resumo_levantadores = []
                for base in bases_unicas:
                    df_base = df_routed[df_routed['BASE_ATRIBUIDA'] == base]
                    df_base_real = df_base[~df_base['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
                    base_ref = next((b for b in st.session_state.bases_records if b['LEVANTADOR'] == base), None)
                    tipo_eq = base_ref.get('TIPO_EQUIPE', 'PRINCIPAL') if base_ref else 'DESCONHECIDO'
                    
                    qtd_comum = sum(count_real_obras(r) for _, r in df_base_real[df_base_real['PRIORIDADE'] == 'Não'].iterrows())
                    qtd_prio = sum(count_real_obras(r) for _, r in df_base_real[df_base_real['PRIORIDADE'] == 'Sim'].iterrows())
                    qtd_super = len(df_base_real[df_base_real['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in df_base_real.columns else 0
                    
                    p_bt = pd.to_numeric(df_base_real['POSTE PREVISTO BT'], errors='coerce').replace(0, np.nan) if 'POSTE PREVISTO BT' in df_base_real.columns else pd.Series(dtype=float)
                    p_mt = pd.to_numeric(df_base_real['POSTE PREVISTO MT'], errors='coerce').replace(0, np.nan) if 'POSTE PREVISTO MT' in df_base_real.columns else pd.Series(dtype=float)
                    if not p_bt.empty or not p_mt.empty:
                        poste_min_series = pd.concat([p_bt, p_mt], axis=1).min(axis=1).fillna(0).round().astype(int)
                    else:
                        poste_min_series = pd.to_numeric(df_base_real['POSTES PREVISTOS'], errors='coerce').fillna(0).round().astype(int) if 'POSTES PREVISTOS' in df_base_real.columns else pd.Series([0]*len(df_base_real))
                        
                    qtd_postes_min_sum = int(poste_min_series.sum())
                    
                    dias_unicos = df_base_real['DIA_MES'].nunique() if 'DIA_MES' in df_base_real.columns else df_base_real['DIA'].nunique()
                    semanas_unicas = df_base_real['SEMANA'].nunique()
                    postes_por_dia = int(round(qtd_postes_min_sum / dias_unicos)) if dias_unicos > 0 else 0
                    postes_por_semana = int(round(qtd_postes_min_sum / semanas_unicas)) if semanas_unicas > 0 else 0
                    
                    resumo_levantadores.append({
                        'LEVANTADOR': base, 'TIPO EQUIPE': tipo_eq, 'OBRAS COMUNS': qtd_comum,
                        'OBRAS PRIORITARIAS': qtd_prio, 'SUPER PONTOS': qtd_super, 'TOTAL OBRAS': qtd_comum + qtd_prio,
                        'POSTES PREVISTOS TOTAIS': qtd_postes_min_sum,
                        'POSTES PREVISTOS / DIA': postes_por_dia,
                        'POSTES PREVISTOS / SEMANA': postes_por_semana,
                        'KM TOTAL PREVISTO': round(df_base['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)
                    })
                df_resumo = pd.DataFrame(resumo_levantadores)
                zip_xl.writestr(f"Resumo_Levantadores - {data_atual_formatada}.xlsx", gerar_excel_resumo_bytes(df_resumo))
                
                cols_to_drop_excel = ['PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'ROTA_GEOMETRIA', '_ORIGINAL_ROWS', '_ORIGEM_BASE']
                cols_to_hide_popup = ['HORA_INICIO', 'HORA_FIM', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'BASE_ATRIBUIDA', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS']
                cols_to_drop_kml_df = ['_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS']
                
                update_ui("Gerando Arquivo Excel de Demanda Geral e Pacote GPX Offline...")
                df_demanda_geral = df_routed.drop(columns=[c for c in cols_to_drop_excel if c in df_routed.columns and c != 'BASE_ATRIBUIDA'], errors='ignore')
                for col in ['POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'POSTES PREVISTOS']:
                    if col in df_demanda_geral.columns:
                        df_demanda_geral[col] = pd.to_numeric(df_demanda_geral[col], errors='coerce').round().fillna(0).astype(int)
                for col in ['DISTANCIA BT', 'DISTANCIA MT', 'DISTANCIA TRAFO']:
                    if col in df_demanda_geral.columns:
                        df_demanda_geral[col] = pd.to_numeric(df_demanda_geral[col], errors='coerce').round(2)
                        
                df_demanda_geral = limpar_colunas_excel(df_demanda_geral, st.session_state.colunas_originais)
                cols_originais_hack = df_demanda_geral.columns.tolist()
                zip_xl.writestr(f"Demanda_Geral - {data_atual_formatada}.xlsx", gerar_excel_bytes(df_demanda_geral, st.session_state.col_prioridade, cols_originais_hack))
                
                update_ui("Gerando Mapa KML Consolidado de todas as rotas...")
                df_routed_kml = df_routed.drop(columns=[c for c in cols_to_drop_kml_df if c in df_routed.columns], errors='ignore')
                for col in ['POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'POSTES PREVISTOS']:
                    if col in df_routed_kml.columns:
                        df_routed_kml[col] = pd.to_numeric(df_routed_kml[col], errors='coerce').round().fillna(0).astype(int)
                
                colunas_exibir_kml = [c for c in st.session_state.colunas_exibir if c not in cols_to_hide_popup]
                
                kml_geral_str = gerar_kml_agrupado(df_routed_kml, st.session_state.bases_records, f"ROTA TOTAL LEVANTADORES - {data_atual_formatada}", colunas_exibir_kml, bases_unicas, tipo_periodo_atual)
                kml_geral_str = re.sub(r'<tr[^>]*>(?:(?!<tr).)*?Horário:(?:(?!</tr>).)*?</tr>', '', kml_geral_str, flags=re.IGNORECASE | re.DOTALL)
                zip_kml.writestr(f"ROTA TOTAL LEVANTADORES - {data_atual_formatada}.kml", kml_geral_str.encode('utf-8'))
                
                # GERADOR DO KML DE OBRAS NAO ALOCADAS
                df_unallocated = st.session_state.get('df_unallocated', pd.DataFrame())
                if not df_unallocated.empty:
                    kml_u = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document><name>OBRAS NÃO ALOCADAS</name>']
                    kml_u.append('<Style id="white_pin"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href></Icon></IconStyle></Style>')
                    for _, r in df_unallocated.iterrows():
                        lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                        if pd.notna(lat) and pd.notna(lon):
                            nome = html.escape(str(r.get('PROTOCOLO', 'Ponto Rejeitado')))
                            kml_u.append(f'<Placemark><name>{nome}</name><styleUrl>#white_pin</styleUrl><Point><coordinates>{lon},{lat}</coordinates></Point></Placemark>')
                    kml_u.append('</Document></kml>')
                    zip_kml.writestr(f"OBRAS_NAO_ALOCADAS - {data_atual_formatada}.kml", "\n".join(kml_u).encode('utf-8'))
                
                gpx_geral_str = gerar_gpx_simples(df_routed_kml, f"ROTA TOTAL LEVANTADORES - {data_atual_formatada}")
                zip_gpx.writestr(f"GPS_ROTA_TOTAL - {data_atual_formatada}.gpx", gpx_geral_str.encode('utf-8'))
                
                for i, base_nome in enumerate(bases_unicas):
                    nome_seguro = re.sub(r'[^A-Za-z0-9_ ]', '', str(base_nome)).replace(" ", "_").upper()
                    nome_seguro = re.sub(r'_+', '_', nome_seguro)
                    df_lev = df_routed[df_routed['BASE_ATRIBUIDA'] == base_nome].copy()
                    
                    update_ui(f"Formatando rotas para: {base_nome}...")
                    df_lev_xl = df_lev.drop(columns=[c for c in cols_to_drop_excel if c in df_lev.columns] + ['BASE_ATRIBUIDA'], errors='ignore')
                    for col in ['POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'POSTES PREVISTOS']:
                        if col in df_lev_xl.columns:
                            df_lev_xl[col] = pd.to_numeric(df_lev_xl[col], errors='coerce').round().fillna(0).astype(int)
                    for col in ['DISTANCIA BT', 'DISTANCIA MT', 'DISTANCIA TRAFO']:
                        if col in df_lev_xl.columns:
                            df_lev_xl[col] = pd.to_numeric(df_lev_xl[col], errors='coerce').round(2)
                            
                    df_lev_xl = limpar_colunas_excel(df_lev_xl, st.session_state.colunas_originais)
                    zip_xl.writestr(f"ROTA_{nome_seguro} - {data_atual_formatada}.xlsx", gerar_excel_bytes(df_lev_xl, st.session_state.col_prioridade, df_lev_xl.columns.tolist()))
                    
                    df_lev_kml = df_lev.drop(columns=[c for c in cols_to_drop_kml_df if c in df_lev.columns], errors='ignore')
                    for col in ['POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'POSTES PREVISTOS']:
                        if col in df_lev_kml.columns:
                            df_lev_kml[col] = pd.to_numeric(df_lev_kml[col], errors='coerce').round().fillna(0).astype(int)
                            
                    kml_lev_str = gerar_kml_agrupado(df_lev_kml, st.session_state.bases_records, f"ROTA_{nome_seguro} - {data_atual_formatada}", colunas_exibir_kml, bases_unicas, tipo_periodo_atual)
                    kml_lev_str = re.sub(r'<tr[^>]*>(?:(?!<tr).)*?Horário:(?:(?!</tr>).)*?</tr>', '', kml_lev_str, flags=re.IGNORECASE | re.DOTALL)
                    zip_kml.writestr(f"ROTA_{nome_seguro} - {data_atual_formatada}.kml", kml_lev_str.encode('utf-8'))
                    
                    gpx_lev_str = gerar_gpx_simples(df_lev_kml, f"ROTA_{nome_seguro} - {data_atual_formatada}")
                    zip_gpx.writestr(f"GPS_ROTA_{nome_seguro} - {data_atual_formatada}.gpx", gpx_lev_str.encode('utf-8'))
                    
            st.session_state.bytes_zip_xl = buf_zip_xl.getvalue()
            st.session_state.bytes_zip_kml = buf_zip_kml.getvalue()
            st.session_state.bytes_zip_gpx = buf_zip_gpx.getvalue()
            
            status_text.success("✅ Pacotes gerados com sucesso! (Rotas extraídas integralmente para KML e GPX).")
            time.sleep(1.5)
            st.session_state.roteamento_concluido = True
            st.session_state.vrp_status = "IDLE"
            tentar_rerun()
            
        except Exception as e:
            st.error(f"🚨 ERRO NO EMPACOTAMENTO: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.session_state.vrp_status = "IDLE"
            if st.button("⬅️ Voltar"): limpar_roteirizador()
            return

# ==========================================
# 4. CONTEÚDO DA PÁGINA FAQ (MANUAL COMPLETO)
# ==========================================
def gerar_excel_modelo(df_modelo):
    output = io.BytesIO()
    df_modelo.to_excel(output, index=False, sheet_name='Modelo')
    return output.getvalue()

def renderizar_faq():
    st.markdown("<h1 class='brand-title' style='margin-bottom: 20px;'>📖 Central de Ajuda e Manual de Operação</h1>", unsafe_allow_html=True)
    
    st.markdown("""
    Bem-vindo ao manual do **Roteirizador NIP v3.0**. Esta página detalha como o Cérebro Logístico (IA) toma decisões, os filtros ocultos e o fluxo de todos os dados do sistema.
    """)
    
    st.markdown("---")
    st.markdown("### 1. As Duas Estratégias de Despacho (Modos)")
    
    col_faq1, col_faq2 = st.columns(2)
    with col_faq1:
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #0D256C; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <h4 style='color: #0D256C; margin-top: 0;'>🎯 1. Planejamento Tático (IA Automática)</h4>
            <b>A IA no Comando:</b> Você sobe a planilha de técnicos e joga milhares de obras brutas. O sistema lê onde cada técnico atua, calcula todas as distâncias cruzadas e distribui as obras do zero de forma otimizada.<br><br>
            <b>A Regra dos 100km (Pulo Logístico):</b> Se o técnico terminar a cota do município dele antes de bater a meta diária, a IA usa o "radar" de 100km em linha reta para acionar e agrupar obras ociosas de cidades vizinhas no mesmo dia, garantindo produtividade máxima.
        </div>
        """, unsafe_allow_html=True)
        
    with col_faq2:
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #55B929; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>
            <h4 style='color: #2e7d32; margin-top: 0;'>♾️ 2. Lista Contínua (Técnico Fixo)</h4>
            <b>O Usuário no Comando:</b> A IA respeita estritamente o que você definiu. Sua planilha já deve ter a coluna <b>LEVANTADOR</b> preenchida.<br><br>
            <b>Processamento Contínuo:</b> A ferramenta ignora as travas de fim de expediente e desenha o caminho mais curto para conectar 100% da lista do profissional. Nenhuma nota é transferida de um técnico para outro, independentemente da distância ou tempo.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("### 2. Baixar Modelos de Planilhas (Templates)")
    st.markdown("Para garantir que a Inteligência consiga rastrear o endereço de cada nota perfeitamente, utilize os modelos padrão abaixo para formatar sua base de dados antes do upload.")

    df_equipes = pd.DataFrame([{
        'MunicIpio': 'FORTUNA', 'Estado': 'Maranhão', 'Levantador': 'NOME DO TECNICO', 
        'Regional': 'CENTRO', 'Longitude': -44.0264, 'Latitude': -5.7335, 'Equipe': 'EQUIPE 17'
    }])

    df_levantamento = pd.DataFrame([{
        'ID SISCO': 1982315, 'PROTOCOLO': 1081945188, 'TIPO NOTA': 'UNR', 'PRIORIDADE': 0, 
        'STATUS SAP': 'ATIV', 'STATUS SISCO': 'Liberado para Levantamento', 'STATUS LIST': 'Em levantamento', 
        'FASE': 'MO', 'PAT': 'PAT1', 'Regional': 'LESTE', 'Município': 'CODÓ', 
        'DISTANCIA BT': 248.78, 'DISTANCIA MT': 241.99, 'DISTANCIA TRAFO': 243.4, 
        'POSTE PREVISTO BT': 6, 'POSTE PREVISTO MT': 2, 'NOME': 'NOME DO CLIENTE', 
        'CONTA CONTRATO': 3019326160, 'INSTALAÇÃO': 2000876176, 'ENDEREÇO': 'ENDERECO COMPLETO', 
        'LOCALIDADE': 'RURAL', 'LATITUDE': -4.459156, 'LONGITUDE': -44.150418, 
        'INFORMAÇÕES EXTRAS': 'INFORMACOES ADICIONAIS', 'TEXTO': 'TEXTO DESCRITIVO', 'TEXTO_GERAL': 'TEXTO GERAL'
    }])

    df_continua = pd.DataFrame([{
        'LEVANTADOR': 'NOME DO TECNICO', 'DATA DESPACHO': '17/08/2026', 'ID SISCO': 1982149, 
        'PROTOCOLO': 1076894592, 'TIPO NOTA': 'UNR', 'PRIORIDADE': 0, 'STATUS SAP': 'ATIV', 
        'STATUS SISCO': 'Pré Análise', 'STATUS LIST': 'Em levantamento', 'FASE': 'MO', 'PAT': 'PAT1', 
        'Regional': 'LESTE', 'Município': 'CAXIAS', 'DISTANCIA BT': 285.18, 'DISTANCIA MT': 212.99, 
        'DISTANCIA TRAFO': 224.66, 'POSTE PREVISTO BT': 7, 'POSTE PREVISTO MT': 2, 
        'NOME': 'NOME DO CLIENTE', 'CONTA CONTRATO': 3018299797, 'INSTALAÇÃO': 2000835041, 
        'ENDEREÇO': 'ENDERECO COMPLETO', 'LOCALIDADE': 'RURAL', 'LATITUDE': -5.060694, 
        'LONGITUDE': -43.438846, 'INFORMAÇÕES EXTRAS': 'INFORMACOES ADICIONAIS'
    }])

    col_dl1, col_dl2, col_dl3 = st.columns(3)
    
    with col_dl1:
        st.markdown("**👥 Planilha de Levantadores**")
        st.caption("Apenas para o Modo Tático. Define onde cada técnico mora/atua.")
        st.download_button("📥 Baixar Modelo Equipes", data=gerar_excel_modelo(df_equipes), file_name="MODELO_LEVANTADORES.xlsx", use_container_width=True)

    with col_dl2:
        st.markdown("**📁 Base Levantamento / Saneamento**")
        st.caption("Apenas para o Modo Tático. A IA distribuirá estas obras automaticamente.")
        st.download_button("📥 Baixar Modelo Obras Livres", data=gerar_excel_modelo(df_levantamento), file_name="MODELO_BASE_LEVANTAMENTO.xlsx", use_container_width=True)

    with col_dl3:
        st.markdown("**♾️ Base Lista Contínua**")
        st.caption("Apenas para o Modo Lista Contínua. Já exige a coluna 'LEVANTADOR'.")
        st.download_button("📥 Baixar Modelo Lista Contínua", data=gerar_excel_modelo(df_continua), file_name="MODELO_LISTA_CONTINUA.xlsx", use_container_width=True)

    st.markdown("---")
    st.markdown("### 3. Filtros Inteligentes e Controle de Escopo")
    
    c_flt1, c_flt2 = st.columns(2)
    with c_flt1:
        st.markdown("**🗂️ Triagem Dinâmica de Notas (Bolo Geral)**")
        st.markdown("Assim que você sobe a planilha, um Mini-Dashboard mostra os totais dos principais tipos de notas (UNR, MGD, ASC, DIF). Abaixo dele, você pode escolher tipos específicos (ex: 'UNR') para **descartar temporariamente**, limpando a base sem precisar editar o arquivo Excel.")

        st.markdown("**🎯 Matriz Multi-Filtro (Regional e PAT)**")
        st.markdown("Na seção 'Escopo da Operação', o sistema lê todas as Regionais e PATs presentes no seu arquivo. Você pode afunilar o roteamento para processar *apenas* a 'REGIONAL LESTE' e *apenas* o 'PAT1', bloqueando o restante do Estado instantaneamente.")
        
        st.markdown("**⚡ Super Pontos (Deduplicação Espacial)**")
        st.markdown("Se a IA encontrar notas sobrepostas em um pequeno raio de ação, ela funde tudo num mega-ícone laranja (`SUPER PONTO`), garantindo que o técnico visite o local apenas uma vez. **Você pode ajustar o raio desse agrupamento (em metros) na barra lateral.**")

    with c_flt2:
        st.markdown("**🔥 Alta Densidade (Modo Produtividade Máxima)**")
        st.markdown("Uma trava que foca apenas no que dá lucro de tempo. Quando ativada na barra lateral, a IA varre o mapa e joga fora as obras isoladas ou esparsas na zona rural. A equipe é enviada apenas para os 'Bolsões de Densidade', e las obras isoladas vão para o arquivo de rejeições para tratativa futura.")
        
        st.markdown("**🚨 Tripla Checagem de Prioridade (Fura Fila)**")
        st.markdown("O sistema exige urgência. A obra fura a fila do roteiro e fica vermelha no mapa se: 1) Você selecioná-la manualmente nos Filtros Dinâmicos; 2) For detectado o status 'CORREÇÃO DE LEVANTAMENTO'; 3) A coluna nativa `PRIORIDADE` no Excel tiver marcações urgentes (ex: 'GIRO NO PRAZO').")

        st.markdown("**🛡️ Ignorar Obras já Despachadas**")
        st.markdown("Se a caixa 'Ignorar obras já despachadas' for marcada, a IA lê a coluna `DATA DESPACHO CAMPO` do seu Excel. Qualquer linha que tenha algo escrito ali é sumariamente ignorada para evitar o retrabalho de rotas já ativas.")


    st.markdown("---")
    st.markdown("### 4. Esforços, Limites e Avisos Gerenciais")
    st.markdown("""
    * **🛑 Trava Total de Operação (O Limite Global da Empresa):** Localizado na barra lateral, define um teto absoluto. Se você digitar **300**, a IA vai garimpar as 300 melhores/mais prioritárias notas do estado e rejeitar todo el resto, poupando a equipe de backoffice. **Se deixar no valor '0' (Zero)**, a trava é desligada e o sistema roteiriza 100% da base que encontrar.
    * **O Paredão Diário (Corte Rígido):** Se a meta for **6 Obras Previstas por Dia**, na hora que a IA montar a 6ª obra, ela aborta o cálculo, traça a linha de "Retorno para a Base" e a 7ª obra cai para a Terça-Feira, blindando o técnico de sobrecarga.
    * **Varredura Reversa (Longe -> Perto):** Permite inverter a lógica da rota diária. Em vez de começar pelas obras da esquina, a IA manda o técnico cedo para a fazenda mais distante do mapa e vem puxando ele de volta obra por obra, para que o fim do expediente seja feito a poucos minutos de casa.
    * **Cálculo de Postes e Malhas:** Nos relatórios, o sistema soma as colunas `POSTE PREVISTO BT` e `POSTE PREVISTO MT`. Para evitar contagem em dobro (pois Alta e Baixa tensão costumam dividir o mesmo poste físico), ele usa matematicamente o Menor Valor entre as duas.
    * **🏨 Alerta de Pernoite / Hotel:** Durante o cálculo, a IA avalia o "Centro de Gravidade" do lote de obras. Se essa mancha de trabalho ficar concentrada a mais de **60 km** de distância da residência do técnico, a aba gerencial sugere **Hospedagem** com um link de busca de pousadas integrado ao Google Maps.
    """)
    
    st.markdown("---")
    st.markdown("### 5. Configurações Avançadas e Saídas (O que você baixa)")
    st.info("""
    * **Traçado de Ruas Real (OSRM) vs. Vetorial Rapido:** Nas configurações de Conexão de Rede, a opção de *Traçado de Ruas Lento* usa uma API global para curvar a linha exatamente pelas rodovias e asfaltos. Se desmarcado (Vetorial Rápido), ele liga as obras em linha reta (padrão satélite), acelerando o tempo de geração de 10 minutos para apenas alguns segundos.
    * **Demanda_Geral.xlsx:** Uma compilação cristalina. A planilha exportada contém **exatamente** as colunas originais do seu projeto, blindadas contra lixo de programação.
    * **Pacote KML e KML de Rejeições:** O KML principal roda em Google Earth (limpo de caixas de textos desnecessárias). Além dele, se alguma obra for isolada pela Alta Densidade ou esgotar a cota da Trava Global, a IA gera o arquivo **`OBRAS_NAO_ALOCADAS.kml`** (pinos brancos) para você visualizar exatamente o que sobrou.
    * **Pacote GPX:** O GPX é o **GPS Offline de Alta Precisão** – feito para o técnico importar em apps como *OsmAnd* ou *Wikiloc* para navegar no sertão e em áreas rurais mesmo quando estiver com 0% de sinal de operadora móvel.
    """)

    st.markdown("---")
    st.markdown("### 6. Bastidores Técnicos e Motores de IA (Arquitetura Interna)")
    st.markdown("""
    Para engenheiros e planejadores que desejam entender a matemática por trás do sistema, a v3.0 executa rotinas automatizadas robustas no motor Python:
    * **Alocação com Fallback de 100km (`assign_load_balanced_strict_and_fallback`):** No modo tático, o motor tenta encaixar a obra estritamente no município do técnico. Caso o estoque acabe, ele ativa um raio de tolerância de até 100 km em linha reta para buscar equipes vizinhas ociosas.
    * **Contagem Real de Ativos (`count_real_obras`):** Identifica automaticamente quando uma linha de dados representa um *Super Ponto* fundido, extraindo a quantidade numérica exata de notas para que os relatórios de postes e metas não fiquem distorcidos.
    * **Heurística Gulosa de Proximidade (`greedy_sort`):** Algoritmo de ordenação sequencial que calcula o vizinho mais próximo a cada parada, podendo ser invertido quando a **Varredura Reversa** está ativada.
    * **Gestão de Dias Úteis Lógicos (`get_workday_date` e `iniciar_dia`):** Mapeiam os dias da semana selecionados pelo usuário (ex: pulando fins de semana ou dias específicos) para garantir que a linha de cronograma de campo respeite o calendário real da empresa.
    """)

# ==========================================
# 6. ESTRUTURA PRINCIPAL E NAVEGAÇÃO
# ==========================================
def main():
    # MENU LATERAL SUPERIOR (ROTEAMENTO ENTRE PÁGINAS)
    with st.sidebar:
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as f:
                encoded_logo = base64.b64encode(f.read()).decode()
            st.markdown(
                f'<div style="text-align: center; margin-bottom: 25px;">'
                f'<img src="data:image/png;base64,{encoded_logo}" style="width: 70%; max-width: 180px; pointer-events: none;">'
                f'</div>',
                unsafe_allow_html=True
            )
            
        st.markdown("### 🧭 Navegação")
        menu_selecionado = st.radio("Selecione a Página:", ["🚀 Roteirizador Logístico", "📖 FAQ & Guia de Uso"], label_visibility="collapsed")
        st.markdown("---")

    if menu_selecionado == "📖 FAQ & Guia de Uso":
        renderizar_faq()
    else:
        app_roteirizador()

if __name__ == "__main__":
    main()
