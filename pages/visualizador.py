import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import zipfile
import re
import numpy as np
import io

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Visualizador de Malha Elétrica",
    page_icon="⚡",
    layout="wide"
)

# ==========================================
# 2. MOTOR DE EXTRAÇÃO LEVE E ANTILIXO
# ==========================================
@st.cache_data(show_spinner=False)
def processar_arquivos_kmz(arquivos):
    dados_extraidos = []
    
    for f in arquivos:
        nome_arquivo = f.name.upper().replace('.KMZ', '').replace('.KML', '')
        conteudo_kml = ""
        
        if f.name.lower().endswith('.kmz'):
            try:
                # Usa BytesIO para leitura ultrarrápida da memória RAM (não trava)
                with zipfile.ZipFile(io.BytesIO(f.getvalue()), 'r') as z:
                    for item in z.namelist():
                        if item.lower().endswith('.kml'):
                            conteudo_kml = z.read(item).decode('utf-8', errors='ignore')
                            break
            except Exception:
                continue
        else:
            try:
                conteudo_kml = f.getvalue().decode('utf-8', errors='ignore')
            except:
                continue
                
        # Tenta extrair a cidade e regional de dentro do arquivo CEMAR
        mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        municipio = mun_match.group(1).strip().upper() if mun_match else "N/A"

        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        regional = reg_match.group(1).strip().upper() if reg_match else "N/A"
        
        if regional == "N/A":
            sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
            if sigla_match:
                regional = sigla_match.group(1)

        linhas_mapa = []
        coords_matches = re.findall(r'<coordinates>(.*?)</coordinates>', conteudo_kml, re.DOTALL)
        
        for match in coords_matches:
            pontos = []
            coordenadas_brutas = match.strip().split()
            for coord in coordenadas_brutas:
                partes = coord.split(',')
                if len(partes) >= 2:
                    try:
                        lon = float(partes[0].strip())
                        lat = float(partes[1].strip())
                        
                        # 🛡️ FILTRO GEOGRÁFICO RESTRITO (O Fim das linhas loucas!)
                        # Ignora os 0.0 e só aceita coordenadas que estejam dentro do BRASIL (Lat: -35 a 5, Lon: -75 a -30)
                        if lat != 0.0 and lon != 0.0 and -35.0 <= lat <= 5.0 and -75.0 <= lon <= -30.0:
                            pontos.append([lat, lon])
                    except:
                        continue
            
            # Só adiciona no mapa se a linha tiver pelo menos 2 pontos válidos conectados
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
# 3. INTERFACE E LÓGICA DO MENU
# ==========================================
def view_visualizador():
    st.markdown("<h2 style='color: #0D256C;'>🗺️ Visualizador Leve de Malha Elétrica (KMZ/KML)</h2>", unsafe_allow_html=True)
    st.markdown("Faça o upload dos seus arquivos de rede. O sistema removerá erros geográficos de projeto e permitirá o filtro rápido.")

    if 'df_rede_vis' not in st.session_state:
        st.session_state.df_rede_vis = pd.DataFrame()

    with st.sidebar:
        st.markdown("### 📥 1. Upload de Malha")
        # Atenção: Fazer upload apenas dos arquivos .KMZ ou .KML
        arquivos_upados = st.file_uploader("Selecione os Alimentadores", type=["kmz", "kml"], accept_multiple_files=True)
        
        if st.button("⚙️ Processar Arquivos", type="primary", use_container_width=True):
            if arquivos_upados:
                with st.spinner("Limpando lixo geográfico e extraindo a rede..."):
                    df_extraido = processar_arquivos_kmz(arquivos_upados)
                    st.session_state.df_rede_vis = df_extraido
                if not df_extraido.empty:
                    st.success(f"✅ {len(df_extraido)} Alimentadores processados com sucesso!")
                else:
                    st.error("Nenhuma coordenada válida encontrada nos arquivos.")
            else:
                st.warning("Suba ao menos um arquivo KMZ ou KML.")
                
        st.markdown("---")
        
        df = st.session_state.df_rede_vis
        
        regioes_sel = []
        municipios_sel = []
        alim_sel = []
        
        if not df.empty:
            st.markdown("### 🔍 2. Filtros de Exibição")
            
            # Filtro de Regional
            lista_regioes = sorted(df['REGIONAL'].unique().tolist())
            regioes_sel = st.multiselect("📍 Regional:", lista_regioes)
            
            df_filt1 = df[df['REGIONAL'].isin(regioes_sel)] if regioes_sel else df
            
            # Filtro de Município (depende do anterior)
            lista_municipios = sorted(df_filt1['MUNICIPIO'].unique().tolist())
            municipios_sel = st.multiselect("🏙️ Município:", lista_municipios)
            
            df_filt2 = df_filt1[df_filt1['MUNICIPIO'].isin(municipios_sel)] if municipios_sel else df_filt1
            
            # Filtro de Alimentador (depende do anterior)
            lista_alimentadores = sorted(df_filt2['ALIMENTADOR'].unique().tolist())
            alim_sel = st.multiselect("⚡ Alimentador (Nome do Arquivo):", lista_alimentadores)

    # ==========================================
    # 4. RENDERIZAÇÃO DO MAPA
    # ==========================================
    if not st.session_state.df_rede_vis.empty:
        df_mapa = df.copy()
        if regioes_sel: df_mapa = df_mapa[df_mapa['REGIONAL'].isin(regioes_sel)]
        if municipios_sel: df_mapa = df_mapa[df_mapa['MUNICIPIO'].isin(municipios_sel)]
        if alim_sel: df_mapa = df_mapa[df_mapa['ALIMENTADOR'].isin(alim_sel)]

        if df_mapa.empty:
            st.info("Nenhuma rede encontrada para os filtros selecionados.")
            return

        # Busca o centro para abrir a câmera do mapa no lugar exato
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
            
            # Paleta de cores para diferenciar os alimentadores quando exibidos juntos
            cores = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe']
            
            for idx, row in df_mapa.iterrows():
                cor_atual = cores[idx % len(cores)]
                alimentador_nome = row['ALIMENTADOR']
                
                # Agrupa tudo em um Layer Control (Menu flutuante do mapa)
                fg = folium.FeatureGroup(name=alimentador_nome)
                for segmento in row['LINHAS']:
                    folium.PolyLine(
                        locations=segmento,
                        color=cor_atual,
                        weight=3,
                        opacity=0.9,
                        tooltip=f"<b>Alimentador:</b> {alimentador_nome}<br><b>Município:</b> {row['MUNICIPIO']}"
                    ).add_to(fg)
                fg.add_to(mapa)

            folium.LayerControl().add_to(mapa)
            st_folium(mapa, use_container_width=True, height=650, returned_objects=[])
            
            st.markdown("### 📋 Tabela Resumo dos Alimentadores Visíveis")
            df_resumo = df_mapa[['REGIONAL', 'MUNICIPIO', 'ALIMENTADOR']].reset_index(drop=True)
            st.dataframe(df_resumo, use_container_width=True)

if __name__ == "__main__":
    view_visualizador()
