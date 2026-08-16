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
    page_title="Roteirizador NIP v2.0",
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
    
    keys_to_clear = ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg', 'tempo_processamento']
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
# 4. CONTEÚDO DA PÁGINA FAQ
# ==========================================
def renderizar_faq():
    st.markdown("<h1 class='brand-title' style='margin-bottom: 30px;'>📖 Central de Ajuda e FAQ</h1>", unsafe_allow_html=True)
    
    st.markdown("""
    Bem-vindo ao manual oficial do **Roteirizador NIP v2.0**. Abaixo você encontrará explicações detalhadas sobre como o "Cérebro" do sistema funciona, as diferenças cruciais entre os modos de operação, e o que cada arquivo faz.
    
    ---
    """)
    
    st.markdown("### 1. A Grande Diferença: Modos de Roteirização")
    st.markdown("A ferramenta oferece duas formas fundamentais de trabalhar. Entender a diferença entre elas é a chave para o sucesso logístico:")
    
    col_faq1, col_faq2 = st.columns(2)
    with col_faq1:
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #0D256C; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>
            <h4 style='color: #0D256C; margin-top: 0;'>🎯 1. Planejamento Tático (IA Automática)</h4>
            <b>Como funciona:</b> A Inteligência Artificial assume o controle total da logística. Você insere uma base com "N" técnicos e joga milhares de obras genéricas no sistema. A IA cruza a localização dos técnicos com a localização das obras e <b>decide sozinha quem vai fazer o quê</b>, buscando a rota mais curta e barata.<br><br>
            <b>Limites e Regras:</b> Respeita rigorosamente a meta diária (ex: 6 obras/dia).<br><br>
            <b>A Regra dos 100km:</b> Se um técnico atende a cidade de São Luís e as obras lá acabam, a IA não deixa ele ocioso. Ela ativa o radar de 100km e "puxa" obras de cidades vizinhas (ex: Raposa) para completar o dia de trabalho dele, tudo de forma automática.
        </div>
        """, unsafe_allow_html=True)
        
    with col_faq2:
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #55B929; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>
            <h4 style='color: #2e7d32; margin-top: 0;'>♾️ 2. Lista Contínua (Técnico Fixo)</h4>
            <b>Como funciona:</b> Você é quem manda! O sistema ignora a criatividade de distribuição da IA e atua apenas como um "Calculador de Rotas". Ele exige que sua planilha já possua a coluna <b>LEVANTADOR</b> preenchida para cada obra.<br><br>
            <b>Limites e Regras:</b> Não há limite de obras por dia nem de semanas. O sistema vai pegar o bloco de obras do <i>João</i>, criar o KML e desenhar a melhor rota para ele visitar 100% daquela lista, custe o que custar.<br><br>
            <b>Bloqueio de Troca:</b> A IA jamais passará uma obra do <i>João</i> para a <i>Maria</i> neste modo, mesmo que a Maria esteja mais perto. É o modo perfeito de usar quando a área técnica já fechou a distribuição dos pacotes.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    
    st.markdown("### 2. O que enviar para a máquina? (Arquivos de Upload)")
    st.markdown("O Roteirizador é flexível, mas possui regras de leitura para as planilhas:")
    
    st.markdown("""
    *   **👥 Planilhas de Equipes (Principais e Temporários):** Usadas *apenas no Planejamento Tático*. A planilha precisa ter uma coluna chamada **`LEVANTADOR`** (ou NOME, TECNICO) e as colunas de **`LATITUDE` / `LONGITUDE`** (ou `MUNICIPIO` para o sistema descobrir as coordenadas base do profissional).
    *   **📁 Base Levantamento:** Planilha bruta com as notas. O sistema vai ler a aba de `STATUS LIST` e só deixará passar as obras que estiverem com status iniciais de operação ou as temidas **"Correções de Levantamento"** (que ganham prioridade máxima imediata).
    *   **📁 Base Saneamento:** Planilha que costuma vir pré-aprovada e com postes definidos. O sistema aprova 100% das linhas que caem aqui (não filtra status).
    *   **📁 Base Genérica / Livre:** Pode ser qualquer CSV ou Excel do mundo, desde que contenha `LATITUDE` e `LONGITUDE`. Excelente para projetos ad-hoc.
    *   **🔄 Planilha de Status (SharePoint):** Se você inserir este arquivo junto, a IA varre suas obras base e "mata" (ignora) todas aquelas que constarem como já concluídas no SharePoint mais recente, evitando despachar a mesma obra duas vezes.
    """)

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    
    st.markdown("### 3. O que sai da máquina? (Arquivos de Download)")
    st.markdown("Após a Inteligência trabalhar, você baixará três pacotes na barra lateral:")

    st.markdown("""
    *   **🌐 1. Baixar Planilhas (ZIP):**
        *   **`Demanda_Geral.xlsx`:** A visão do Gerente. Todas as obras do projeto ordenadas de 1 a N, com a coluna `LEVANTADOR_RESPONSAVEL` no início.
        *   **`ROTA_Nome_Tecnico.xlsx`:** A visão do Campo. Planilhas limpas, apenas com as colunas originais do seu projeto. Todo o "lixo" de processamento de máquina (distâncias algorítmicas de retas, coordenadas do folium, etc.) é destruído antes da exportação para não poluir a tela do levantador.
        *   **`Resumo_Levantadores.xlsx`:** O painel logístico mostrando quantos postes (BT/MT) totais cada técnico deve prever. O cálculo é inteligente: ele pega o menor número entre BT e MT (para evitar somar o mesmo poste duas vezes) e arredonda para números inteiros (porque não existe meio poste no mundo real).
    
    *   **🗺️ 2. Baixar Mapas (KML):** 
        *   Arquivo nativo para visualizar no **Google Earth** ou Meu Maps. Exibe ícones coloridos: Laranja para *Super Pontos*, Vermelho para *Obras Prioritárias* e Azul para comuns. Ele foi higienizado com expressões regulares pesadas para não exibir campos de "Horário" internos e nem ícones indesejados. As linhas coloridas conectam o melhor caminho passo a passo.
    
    *   **🛰️ 3. Baixar GPS Offline (GPX):**
        *   O arquivo que salva vidas no interior! Formato puro de satélite (`.gpx`). Pode ser importado em aplicativos de celular como **OsmAnd**, **Wikiloc** ou **GPX Viewer**. Se o levantador perder o sinal de internet/4G no mato ou na estrada de terra, o GPS continuará lendo a rota por cima via satélite.
    """)
    
    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    
    st.markdown("### 4. Os 'Poderes Invisíveis' (Regras que rodam nos bastidores)")
    st.markdown("""
    *   **⚡ Super Pontos (Deduplicação):** A IA odeia ineficiência. Se a sua planilha possuir 4 notas geradas com coordenadas muito próximas (raio de 100 metros) - indicando que é tudo na mesma rua ou no mesmo poste - a IA as amarra invisivelmente e as trata como um "SUPER PONTO". A meta do técnico não cai, mas o mapa fica muito mais leve e inteligente.
    *   **🚨 Força Bruta de Prioridade:** Se houver uma coluna `PRIORIDADE` na sua planilha e estiver escrita com qualquer valor diferente de zero ou vazio (ex: "Giro Fora do Prazo", "URGENTE", etc), ou se a nota estiver como `CORREÇÃO DE LEVANTAMENTO`, a IA fura a fila por conta própria e grifa isso em vermelho para a equipe despachar primeiro.
    *   **🏨 Modo Pernoite / Acampamento:** Se o "Centro de Gravidade" do bloco de obras despachado para um levantador ficar a mais de **60 KM** da casa dele, a UI exibe um alerta vermelho na aba "Apoio Logístico", sugerindo que o técnico seja hospedado em um hotel para não estourar os custos de gasolina indo e voltando rodovias todo dia.
    """)

# ==========================================
# 5. ESTRUTURA PRINCIPAL E NAVEGAÇÃO
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
