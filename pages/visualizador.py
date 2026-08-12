import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import zipfile
import re
import numpy as np
import io
import xml.etree.ElementTree as ET
import html

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Visualizador de Malha",
    page_icon="🗺️",
    layout="wide"
)

# ==========================================
# 2. MOTOR DE EXTRAÇÃO DE KML (COM NOMES)
# ==========================================
def extrair_coordenadas_vis(texto_coords):
    """Filtro Antilixo: Converte texto de coordenadas e ignora a Ilha Nula e coordenadas fora do Brasil"""
    pontos = []
    coordenadas_brutas = texto_coords.strip().split()
    for coord in coordenadas_brutas:
        partes = coord.split(',')
        if len(partes) >= 2:
            try:
                lon = float(partes[0].strip())
                lat = float(partes[1].strip())
                if lat != 0.0 and lon != 0.0 and -35.0 <= lat <= 5.0 and -75.0 <= lon <= -30.0:
                    pontos.append([lat, lon]) 
            except:
                continue
    return pontos

@st.cache_data(show_spinner=False)
def processar_arquivos_kmz_estruturado(arquivos):
    dados_extraidos = []
    
    for f in arquivos:
        nome_arquivo = f.name.upper().replace('.KMZ', '').replace('.KML', '')
        conteudo_kml = ""
        
        if f.name.lower().endswith('.kmz'):
            try:
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
        
        # Limpeza agressiva de namespaces XML para não quebrar a leitura
        conteudo_kml = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', conteudo_kml)
        
        try:
            root = ET.fromstring(conteudo_kml)
        except Exception:
            continue 

        municipio = "N/A"
        regional = "N/A"
        mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if mun_match: municipio = mun_match.group(1).strip().upper()
        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if reg_match: regional = reg_match.group(1).strip().upper()
        
        if regional == "N/A":
            sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
            if sigla_match: regional = sigla_match.group(1)

        camadas_do_alimentador = {}
        
        for folder in root.findall('.//Folder'):
            name_tag = folder.find('name')
            if name_tag is not None and name_tag.text:
                nome_pasta = name_tag.text.strip().upper()
                
                # Ignora pastas raiz
                if "CEMAR" in nome_pasta or nome_arquivo in nome_pasta:
                    continue

                linhas = []
                pontos = []
                
                for placemark in folder.findall('.//Placemark'):
                    # Captura o NOME DO POSTE/TRAFO/REDE da tag <name>
                    pm_name_tag = placemark.find('name')
                    nome_elemento = pm_name_tag.text.strip() if pm_name_tag is not None and pm_name_tag.text else "S/N"
                    
                    for ls in placemark.findall('.//LineString/coordinates'):
                        if ls.text:
                            coords_linha = extrair_coordenadas_vis(ls.text)
                            if len(coords_linha) > 1:
                                linhas.append({'nome': nome_elemento, 'coords': coords_linha})
                            
                    for pt in placemark.findall('.//Point/coordinates'):
                        if pt.text:
                            coords_ponto = extrair_coordenadas_vis(pt.text)
                            if len(coords_ponto) > 0:
                                pontos.append({'nome': nome_elemento, 'coords': coords_ponto[0]}) 
                
                if linhas or pontos:
                    camadas_do_alimentador[nome_pasta] = {
                        'linhas': linhas,
                        'pontos': pontos
                    }
        
        if camadas_do_alimentador:
            dados_extraidos.append({
                'ALIMENTADOR': nome_arquivo,
                'REGIONAL': regional,
                'MUNICIPIO': municipio,
                'CAMADAS': camadas_do_alimentador
            })
            
    return pd.DataFrame(dados_extraidos)

# ==========================================
# 3. INTERFACE E APLICATIVO
# ==========================================
st.markdown("<h2 style='color: #0D256C;'>🗺️ Inspeção de Malha Elétrica (KMZ)</h2>", unsafe_allow_html=True)
st.markdown("Faça o upload dos arquivos KMZ. O sistema recriará as camadas e permitirá pesquisa avançada de equipamentos.")

if 'df_rede_vis' not in st.session_state:
    st.session_state.df_rede_vis = pd.DataFrame()
    st.session_state.processed_files = set()

with st.sidebar:
    st.markdown("### 📥 1. Upload de Malha")
    arquivos_upados = st.file_uploader("Selecione os Alimentadores", type=["kmz", "kml"], accept_multiple_files=True)
    
    nomes_arquivos_atuais = set([f.name for f in arquivos_upados]) if arquivos_upados else set()
    
    if nomes_arquivos_atuais != st.session_state.processed_files:
        if arquivos_upados:
            with st.spinner("Lendo metadados e processando estrutura..."):
                df_extraido = processar_arquivos_kmz_estruturado(arquivos_upados)
                st.session_state.df_rede_vis = df_extraido
                st.session_state.processed_files = nomes_arquivos_atuais
        else:
            st.session_state.df_rede_vis = pd.DataFrame()
            st.session_state.processed_files = set()

    df = st.session_state.df_rede_vis
    
    alim_sel = []
    camadas_ativas = {}
    termo_pesquisa = ""
    
    if not df.empty:
        st.markdown("---")
        st.markdown("### 🔎 2. Pesquisa de Equipamento")
        termo_pesquisa = st.text_input("Nome/Num. Poste ou Trafo:", placeholder="Ex: 554930, 201...", help="Ao pesquisar, o mapa focará diretamente neste equipamento.").strip().upper()
        
        st.markdown("---")
        st.markdown("### 🔍 3. Filtros Geográficos")
        
        lista_regioes = sorted(df['REGIONAL'].unique().tolist())
        regioes_sel = st.multiselect("📍 Escolha a Regional:", lista_regioes)
        df_filt1 = df[df['REGIONAL'].isin(regioes_sel)] if regioes_sel else df
        
        lista_municipios = sorted(df_filt1['MUNICIPIO'].unique().tolist())
        municipios_sel = st.multiselect("🏙️ Escolha o Município:", lista_municipios)
        df_filt2 = df_filt1[df_filt1['MUNICIPIO'].isin(municipios_sel)] if municipios_sel else df_filt1
        
        lista_alimentadores = sorted(df_filt2['ALIMENTADOR'].unique().tolist())
        alim_sel = st.multiselect("⚡ Selecione o Alimentador:", lista_alimentadores)
        
        alimentadores_visiveis = alim_sel if alim_sel else lista_alimentadores

        st.markdown("---")
        st.markdown("### 🗂️ 4. Camadas (Google Earth)")
        
        for alim in alimentadores_visiveis:
            st.markdown(f"**{alim}**")
            dict_camadas = df[df['ALIMENTADOR'] == alim]['CAMADAS'].iloc[0]
            lista_camadas_alim = sorted(list(dict_camadas.keys()))
            
            # 🛡️ BUSCA INTELIGENTE POR PALAVRA-CHAVE PARA MARCAR POR PADRÃO
            camadas_default = []
            for c in lista_camadas_alim:
                c_upper = str(c).upper()
                if 'PRIM' in c_upper or 'SECUND' in c_upper or 'TRANSF' in c_upper or 'POSTE' in c_upper:
                    camadas_default.append(c)
            
            camadas_ativas[alim] = st.multiselect(
                "Ligar/Desligar Visibilidade:", 
                lista_camadas_alim, 
                default=camadas_default,
                key=f"ms_{alim}"
            )

# ==========================================
# 4. RENDERIZAÇÃO DO MAPA (SEMPRE ATIVO)
# ==========================================

# 1. Mapa nasce sempre criado (mesmo sem arquivos)
mapa = folium.Map(location=[-5.2, -45.0], zoom_start=7, tiles=None)

folium.TileLayer(
    tiles='https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}',
    attr='Google',
    name='Satélite (Google)',
    overlay=False,
    control=True
).add_to(mapa)

folium.TileLayer(
    tiles='CartoDB positron',
    name='Mapa Vetorial (Limpo)',
    overlay=False,
    control=True
).add_to(mapa)

# 2. Popula o mapa se houver dados
if not df.empty:
    df_mapa = df.copy()
    if regioes_sel: df_mapa = df_mapa[df_mapa['REGIONAL'].isin(regioes_sel)]
    if municipios_sel: df_mapa = df_mapa[df_mapa['MUNICIPIO'].isin(municipios_sel)]
    
    alimentadores_visiveis = alim_sel if alim_sel else df_mapa['ALIMENTADOR'].unique().tolist()
    df_mapa = df_mapa[df_mapa['ALIMENTADOR'].isin(alimentadores_visiveis)]

    dict_cores = {
        'REDE PRIMÁRIA': '#e6194b', 'REDE PRIMARIA': '#e6194b',
        'REDE SECUNDÁRIA': '#4363d8', 'REDE SECUNDARIA': '#4363d8',
        'POSTE': '#a9a9a9', 'TRANSFORMADOR': '#f58231', 
        'CHAVE': '#3cb44b', 'REGULADOR': '#911eb4', 
        'RELIGADOR': '#46f0f0', 'SUBESTAÇÃO': '#000000', 'SUBESTACAO': '#000000'
    }

    todas_lats, todas_lons = [], []
    busca_lats, busca_lons = [], []
    
    for alim in alimentadores_visiveis:
        if alim not in df_mapa['ALIMENTADOR'].values:
            continue
            
        dict_camadas = df_mapa[df_mapa['ALIMENTADOR'] == alim]['CAMADAS'].iloc[0]
        camadas_permitidas = camadas_ativas.get(alim, [])
        
        fg_alim = folium.FeatureGroup(name=f"⚡ {alim}", show=True)
        
        for nome_pasta, conteudos in dict_camadas.items():
            if nome_pasta not in camadas_permitidas:
                continue 
                
            cor_elemento = dict_cores.get(nome_pasta, '#333333') 
            
            # --- PROCESSA AS LINHAS ---
            for linha in conteudos['linhas']:
                nome_item = str(linha['nome'])
                coords_item = linha['coords']
                
                pesquisado = termo_pesquisa != "" and termo_pesquisa in nome_item.upper()
                cor_final = "#FF00FF" if pesquisado else cor_elemento # Magenta se pesquisado
                peso = 8 if pesquisado else (3 if 'PRIM' in nome_pasta else 2)
                
                if pesquisado:
                    for pt in coords_item:
                        busca_lats.append(pt[0]); busca_lons.append(pt[1])

                html_popup = f"""
                <div style="min-width: 200px; font-family: sans-serif;">
                    <h4 style="margin-top: 0; color: {cor_elemento}; border-bottom: 1px solid #ccc; padding-bottom: 5px;">{nome_pasta}</h4>
                    <b style="color: #555;">IDENTIFICAÇÃO:</b> {html.escape(nome_item)}<br>
                    <b style="color: #555;">ALIMENTADOR:</b> {html.escape(alim)}
                </div>
                """
                
                folium.PolyLine(
                    locations=coords_item,
                    color=cor_final,
                    weight=peso,
                    opacity=0.9 if not pesquisado else 1.0,
                    tooltip=f"<b>{nome_pasta}</b><br>{html.escape(nome_item)}",
                    popup=folium.Popup(html_popup, max_width=300)
                ).add_to(fg_alim)
                
                for pt in coords_item:
                    todas_lats.append(pt[0]); todas_lons.append(pt[1])
            
            # --- PROCESSA OS PONTOS (Postes, Trafos) ---
            for ponto in conteudos['pontos']:
                nome_item = str(ponto['nome'])
                coords_item = ponto['coords']
                
                pesquisado = termo_pesquisa != "" and termo_pesquisa in nome_item.upper()
                cor_final = "#FF00FF" if pesquisado else cor_elemento
                raio = 10 if pesquisado else (5 if 'TRANSFORMADOR' in nome_pasta else (3 if 'POSTE' in nome_pasta else 4))
                
                if pesquisado:
                    busca_lats.append(coords_item[0]); busca_lons.append(coords_item[1])
                    folium.Marker(
                        location=coords_item,
                        icon=folium.Icon(color='purple', icon='star'),
                        tooltip=f"ALVO ENCONTRADO: {nome_item}"
                    ).add_to(fg_alim)

                html_popup = f"""
                <div style="min-width: 200px; font-family: sans-serif;">
                    <h4 style="margin-top: 0; color: {cor_elemento}; border-bottom: 1px solid #ccc; padding-bottom: 5px;">{nome_pasta}</h4>
                    <b style="color: #555;">IDENTIFICAÇÃO:</b> {html.escape(nome_item)}<br>
                    <b style="color: #555;">ALIMENTADOR:</b> {html.escape(alim)}<br>
                    <b style="color: #555;">GPS:</b> {coords_item[0]:.5f}, {coords_item[1]:.5f}
                </div>
                """

                folium.CircleMarker(
                    location=coords_item,
                    radius=raio,
                    color=cor_final,
                    fill=True,
                    fill_color=cor_final,
                    fill_opacity=1.0 if pesquisado else 0.8,
                    tooltip=f"<b>{nome_pasta}</b><br>{html.escape(nome_item)}",
                    popup=folium.Popup(html_popup, max_width=300)
                ).add_to(fg_alim)
                
                todas_lats.append(coords_item[0]); todas_lons.append(coords_item[1])
                
        fg_alim.add_to(mapa)

    # 3. Foco da Câmera
    if busca_lats and busca_lons:
        mapa.fit_bounds([[min(busca_lats), min(busca_lons)], [max(busca_lats), max(busca_lons)]], max_zoom=19)
    elif todas_lats and todas_lons:
        mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])

# 4. Renderiza na tela
folium.LayerControl(position='topright').add_to(mapa)
st_folium(mapa, use_container_width=True, height=750, returned_objects=[])
