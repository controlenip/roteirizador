import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import zipfile
import re
import numpy as np

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Visualizador de Malha Elétrica",
    page_icon="⚡",
    layout="wide"
)

# ==========================================
# 2. MOTOR DE EXTRAÇÃO LEVE (SEM TRAVAR)
# ==========================================
@st.cache_data(show_spinner=False)
def processar_arquivos_kmz(arquivos):
    dados_extraidos = []
    
    for f in arquivos:
        nome_arquivo = f.name.upper().replace('.KMZ', '').replace('.KML', '')
        
        # 1. Leitura bruta do conteúdo
        conteudo_kml = ""
        if f.name.lower().endswith('.kmz'):
            try:
                with zipfile.ZipFile(f, 'r') as z:
                    for item in z.namelist():
                        if item.lower().endswith('.kml'):
                            conteudo_kml = z.read(item).decode('utf-8', errors='ignore')
                            break
            except Exception:
                continue
        else:
            conteudo_kml = f.read().decode('utf-8', errors='ignore')
            
        # 2. Tentativa de extrair Regional e Município dos metadados do CEMAR/Equatorial
        mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        municipio = mun_match.group(1).strip().upper() if mun_match else "NÃO IDENTIFICADO"

        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        regional = reg_match.group(1).strip().upper() if reg_match else "NÃO IDENTIFICADO"
        
        # Fallback: Se não achar a regional no KML, tenta extrair a sigla do nome do arquivo (ex: [GDI], [AAM])
        if regional == "NÃO IDENTIFICADO":
            sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
            if sigla_match:
                regional = sigla_match.group(1)

        # 3. Extrair apenas as geometrias de linha (LineStrings) para desenhar o traçado
        linhas_mapa = []
        coords_matches = re.findall(r'<coordinates>(.*?)</coordinates>', conteudo_kml, re.DOTALL)
        
        for match in coords_matches:
            pontos = []
            # Limpa quebras de linha e espaços
            coordenadas_brutas = match.strip().split()
            for coord in coordenadas_brutas:
                partes = coord.split(',')
                if len(partes) >= 2:
                    try:
                        # Folium exige o formato [Latitude, Longitude]
                        lon = float(partes[0].strip())
                        lat = float(partes[1].strip())
                        pontos.append([lat, lon])
                    except:
                        continue
            
            # Só adiciona se for uma linha válida
            if len(pontos) > 1:
                linhas_mapa.append(pontos)
                
        if linhas_mapa:
            dados_extraidos.append({
                'ALIMENTADOR': nome_arquivo,
                'REGIONAL': regional,
                'MUNICIPIO': municipio,
                'LINHAS': linhas_mapa
            })
            
    return pd.DataFrame(dados_extraidos)

# ==========================================
# 3. INTERFACE E LÓGICA DE FILTROS
# ==========================================
def view_visualizador():
    st.markdown("<h2 style='color: #0D256C;'>🗺️ Visualizador Leve de Rede Existente</h2>", unsafe_allow_html=True)
    st.markdown("Faça o upload dos seus arquivos KMZ/KML. O sistema extrairá a malha e permitirá o filtro rápido para análise.")

    # Estado da sessão para armazenar os dados e não recarregar toda hora
    if 'df_rede' not in st.session_state:
        st.session_state.df_rede = pd.DataFrame()

    with st.sidebar:
        st.markdown("### 📥 1. Upload de Malha")
        arquivos_upados = st.file_uploader("Selecione os KMZs", type=["kmz", "kml"], accept_multiple_files=True)
        
        if st.button("⚙️ Processar Arquivos", type="primary", use_container_width=True):
            if arquivos_upados:
                with st.spinner("Extraindo linhas de rede..."):
                    df_extraido = processar_arquivos_kmz(arquivos_upados)
                    st.session_state.df_rede = df_extraido
                st.success(f"{len(df_extraido)} Alimentadores processados!")
            else:
                st.warning("Suba ao menos um arquivo KMZ.")
                
        st.markdown("---")
        
        df = st.session_state.df_rede
        
        # Filtros Dinâmicos
        regioes_selecionadas = []
        municipios_selecionados = []
        alimentadores_selecionados = []
        
        if not df.empty:
            st.markdown("### 🔍 2. Filtros de Exibição")
            
            # Filtro 1: Regional
            lista_regioes = sorted(df['REGIONAL'].unique().tolist())
            regioes_selecionadas = st.multiselect("📍 Regional:", lista_regioes)
            
            # Aplica filtro para o próximo dropdown
            df_filt1 = df[df['REGIONAL'].isin(regioes_selecionadas)] if regioes_selecionadas else df
            
            # Filtro 2: Município
            lista_municipios = sorted(df_filt1['MUNICIPIO'].unique().tolist())
            municipios_selecionados = st.multiselect("🏙️ Município:", lista_municipios)
            
            # Aplica filtro para o próximo dropdown
            df_filt2 = df_filt1[df_filt1['MUNICIPIO'].isin(municipios_selecionados)] if municipios_selecionados else df_filt1
            
            # Filtro 3: Alimentador
            lista_alimentadores = sorted(df_filt2['ALIMENTADOR'].unique().tolist())
            alimentadores_selecionados = st.multiselect("⚡ Alimentador (Arquivo):", lista_alimentadores)

    # ==========================================
    # 4. RENDERIZAÇÃO DO MAPA
    # ==========================================
    if not st.session_state.df_rede.empty:
        # Pega a base filtrada ou a base completa se nada for filtrado
        df_mapa = df.copy()
        if regioes_selecionadas:
            df_mapa = df_mapa[df_mapa['REGIONAL'].isin(regioes_selecionadas)]
        if municipios_selecionados:
            df_mapa = df_mapa[df_mapa['MUNICIPIO'].isin(municipios_selecionados)]
        if alimentadores_selecionados:
            df_mapa = df_mapa[df_mapa['ALIMENTADOR'].isin(alimentadores_selecionados)]

        if df_mapa.empty:
            st.info("Nenhuma rede encontrada para os filtros selecionados.")
            return

        # Centralização inteligente do mapa
        todas_lats = []
        todas_lons = []
        for linhas in df_mapa['LINHAS']:
            for segmento in linhas:
                for ponto in segmento:
                    todas_lats.append(ponto[0])
                    todas_lons.append(ponto[1])
                    
        if todas_lats and todas_lons:
            centro_lat = np.mean(todas_lats)
            centro_lon = np.mean(todas_lons)
            mapa = folium.Map(location=[centro_lat, centro_lon], zoom_start=11, tiles="CartoDB positron")
            
            # Desenhando as linhas
            cores = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe']
            
            for idx, row in df_mapa.iterrows():
                cor_atual = cores[idx % len(cores)]
                alimentador_nome = row['ALIMENTADOR']
                
                # Criar um grupo para cada alimentador (permite ligar/desligar no mapa)
                fg = folium.FeatureGroup(name=alimentador_nome)
                
                for segmento in row['LINHAS']:
                    folium.PolyLine(
                        locations=segmento,
                        color=cor_atual,
                        weight=3,
                        opacity=0.8,
                        tooltip=f"<b>Alimentador:</b> {alimentador_nome}<br><b>Mun:</b> {row['MUNICIPIO']}"
                    ).add_to(fg)
                
                fg.add_to(mapa)

            # Adiciona controle de camadas (Layer Control)
            folium.LayerControl().add_to(mapa)
            
            # Renderiza o mapa na tela inteira
            st_folium(mapa, use_container_width=True, height=650, returned_objects=[])
            
            # Tabela de Resumo abaixo do mapa
            st.markdown("### 📋 Tabela Resumo dos Alimentadores Visíveis")
            df_resumo = df_mapa[['REGIONAL', 'MUNICIPIO', 'ALIMENTADOR']].reset_index(drop=True)
            st.dataframe(df_resumo, use_container_width=True)

if __name__ == "__main__":
    view_visualizador()
