import streamlit as st
import pandas as pd
import io
import zipfile
import re
import time
import os
import xml.etree.ElementTree as ET
import pydeck as pdk
import math
import requests
import unicodedata
import copy

st.set_page_config(page_title="Visualizador de Malha", page_icon="🗺️", layout="wide")

if not os.path.exists("database/redes"):
    os.makedirs("database/redes", exist_ok=True)

# ==========================================
# 1. FUNÇÕES MATEMÁTICAS E DE EXTRAÇÃO
# ==========================================
def haversine(lat1, lon1, lat2, lon2):
    """Calcula a distância real em KM entre duas coordenadas geográficas"""
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

def extrair_coordenadas_pdk(texto_coords):
    pontos = []
    for coord in texto_coords.strip().split():
        partes = coord.split(',')
        if len(partes) >= 2:
            try:
                lon = float(partes[0].strip())
                lat = float(partes[1].strip())
                if lat != 0.0 and lon != 0.0 and -35.0 <= lat <= 5.0 and -75.0 <= lon <= -30.0:
                    pontos.append([lon, lat]) 
            except:
                continue
    return pontos

def processar_e_salvar_kmz(arquivos):
    dict_cores_rgb = {
        'REDE PRIMÁRIA': [230, 25, 75], 'REDE PRIMARIA': [230, 25, 75],
        'REDE SECUNDÁRIA': [67, 99, 216], 'REDE SECUNDARIA': [67, 99, 216],
        'POSTE': [150, 150, 150], 'TRANSFORMADOR': [245, 130, 49], 
        'CHAVE': [60, 180, 75], 'REGULADOR': [145, 30, 180], 
        'RELIGADOR': [70, 240, 240], 'SUBESTAÇÃO': [255, 255, 255], 'SUBESTACAO': [255, 255, 255]
    }

    novos_processados = 0
    for f in arquivos:
        nome_arquivo = f.name.upper().replace('.KMZ', '').replace('.KML', '')
        caminho_db = f"database/redes/{nome_arquivo}.pkl"
        
        if os.path.exists(caminho_db):
            continue
            
        conteudo_kml = ""
        if f.name.lower().endswith('.kmz'):
            try:
                with zipfile.ZipFile(io.BytesIO(f.getvalue()), 'r') as z:
                    for item in z.namelist():
                        if item.lower().endswith('.kml'):
                            conteudo_kml = z.read(item).decode('utf-8', errors='ignore')
                            break
            except Exception: continue
        else:
            try: conteudo_kml = f.getvalue().decode('utf-8', errors='ignore')
            except: continue
        
        conteudo_kml = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', conteudo_kml)
        try: root = ET.fromstring(conteudo_kml)
        except Exception: continue 

        municipio = "N/A"
        regional = "N/A"
        mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if mun_match: municipio = mun_match.group(1).strip().upper()
        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if reg_match: regional = reg_match.group(1).strip().upper()
        if regional == "N/A":
            sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
            if sigla_match: regional = sigla_match.group(1)

        registros_flat = []
        
        for folder in root.findall('.//Folder'):
            name_tag = folder.find('name')
            if name_tag is not None and name_tag.text:
                nome_pasta = name_tag.text.strip().upper()
                if "CEMAR" in nome_pasta or nome_arquivo in nome_pasta: continue

                cor_elemento = dict_cores_rgb.get(nome_pasta, [100, 100, 100])
                
                for placemark in folder.findall('.//Placemark'):
                    pm_name_tag = placemark.find('name')
                    nome_elemento = pm_name_tag.text.strip() if pm_name_tag is not None and pm_name_tag.text else "S/N"
                    
                    for ls in placemark.findall('.//LineString/coordinates'):
                        if ls.text:
                            coords = extrair_coordenadas_pdk(ls.text)
                            if len(coords) > 1:
                                peso = 4 if 'PRIM' in nome_pasta else 2
                                registros_flat.append({
                                    'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio,
                                    'TIPO_GEOMETRIA': 'Linha', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento,
                                    'COORDS': coords, 'COR': cor_elemento, 'RAIO': 0, 'PESO': peso
                                })
                            
                    for pt in placemark.findall('.//Point/coordinates'):
                        if pt.text:
                            coords = extrair_coordenadas_pdk(pt.text)
                            if len(coords) > 0:
                                raio = 5 if 'TRANSFORMADOR' in nome_pasta else (2 if 'POSTE' in nome_pasta else 3)
                                registros_flat.append({
                                    'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio,
                                    'TIPO_GEOMETRIA': 'Ponto', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento,
                                    'COORDS': coords[0], 'COR': cor_elemento, 'RAIO': raio, 'PESO': 0
                                })
        
        if registros_flat:
            df_alimentador = pd.DataFrame(registros_flat)
            df_alimentador.to_pickle(caminho_db)
            novos_processados += 1
            
    return novos_processados

@st.cache_data(show_spinner=False)
def carregar_banco_redes():
    dfs = []
    arquivos = [f for f in os.listdir("database/redes") if f.endswith('.pkl')]
    for f in arquivos:
        try: dfs.append(pd.read_pickle(f"database/redes/{f}"))
        except: pass
    if dfs: return pd.concat(dfs, ignore_index=True)
    return pd.DataFrame()

# ==========================================
# NOVO: MAPEAMENTO DO MARANHÃO (GEOJSON + BASE.XLSX)
# ==========================================
@st.cache_data(show_spinner=False)
def get_base_geojson():
    def remove_accents(input_str):
        nfkd_form = unicodedata.normalize('NFKD', str(input_str))
        return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

    mun_to_reg = {}
    try:
        if os.path.exists("BASE.xlsx"):
            df_base = pd.read_excel("BASE.xlsx")
            if 'MunicIpio' in df_base.columns and 'Regional' in df_base.columns:
                df_base['MunicIpio_norm'] = df_base['MunicIpio'].apply(lambda x: remove_accents(x).upper().strip())
                mun_to_reg = dict(zip(df_base['MunicIpio_norm'], df_base['Regional'].str.upper().str.strip()))
    except:
        pass

    url_geojson = "https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-21-mun.json"
    try:
        resp = requests.get(url_geojson)
        geo_data = resp.json()
    except:
        return None

    for feature in geo_data['features']:
        mun_name = feature['properties']['name']
        mun_name_norm = remove_accents(mun_name).upper().strip()
        reg = mun_to_reg.get(mun_name_norm, "DESCONHECIDO")
        
        feature['properties']['REGIONAL'] = reg
        feature['properties']['NOME'] = mun_name
        feature['properties']['ALIMENTADOR'] = f"Regional: {reg}"
        feature['properties']['TIPO_REDE'] = "MUNICÍPIO"
        
    return geo_data


# ==========================================
# 2. INTERFACE DE MENU E MAPA
# ==========================================
st.markdown("<h2 style='color: #0D256C;'>🗺️ Visualizador Avançado de Malha (Alta Performance)</h2>", unsafe_allow_html=True)
st.markdown("O sistema usa **Aceleração 3D (GPU)** e **Banco de Dados** para aguentar o estado inteiro sem travar. O que você upar fica salvo e disponível na hora.")

df = carregar_banco_redes()

with st.sidebar:
    st.markdown("### 📥 1. Upload para o Banco de Dados")
    arquivos_upados = st.file_uploader("Arraste novos KMZs aqui", type=["kmz", "kml"], accept_multiple_files=True)
    
    if arquivos_upados:
        if st.button("💾 Processar e Salvar no Banco", type="primary", use_container_width=True):
            with st.spinner("Decodificando, limpando lixos e gravando em disco..."):
                qtd = processar_e_salvar_kmz(arquivos_upados)
            if qtd > 0:
                st.success(f"✅ {qtd} novos Alimentadores salvos!")
                carregar_banco_redes.clear()
                time.sleep(1)
                st.rerun()
            else:
                st.info("Os arquivos upados já existiam no banco ou não continham redes válidas.")

    st.markdown("---")
    
    termo_pesquisa = ""
    busca_lat = None
    busca_lon = None

    st.markdown("### 🔎 2. Pesquisas Inteligentes")
    tab_nome, tab_coord = st.tabs(["📝 Por Nome/ID", "📍 Por Coordenada"])
    
    with tab_nome:
        termo_pesquisa = st.text_input("Nome/Num. Poste ou Trafo:", placeholder="Ex: 554930...").strip().upper()
        
    with tab_coord:
        c_lat, c_lon = st.columns(2)
        with c_lat:
            lat_input = st.text_input("Latitude:", placeholder="Ex: -5.532", help="Pode usar ponto ou vírgula")
        with c_lon:
            lon_input = st.text_input("Longitude:", placeholder="Ex: -47.432")
            
        if lat_input and lon_input:
            try:
                b_lat = float(lat_input.replace(',', '.').strip())
                b_lon = float(lon_input.replace(',', '.').strip())
                if -35.0 <= b_lat <= 5.0 and -75.0 <= b_lon <= -30.0:
                    busca_lat = b_lat
                    busca_lon = b_lon
                else:
                    st.warning("⚠️ Coordenada fora do Brasil.")
            except:
                st.warning("⚠️ Formato inválido. Digite apenas números.")

    st.markdown("---")
    st.markdown("### 🔍 3. Filtros Geográficos")
    lista_regioes = sorted(df['REGIONAL'].unique().tolist()) if not df.empty else ["SUL", "LESTE", "NORTE", "NOROESTE", "CENTRO"]
    regioes_sel = st.multiselect("📍 Regional:", lista_regioes)
    df_filt1 = df[df['REGIONAL'].isin(regioes_sel)] if regioes_sel and not df.empty else df
    
    lista_municipios = sorted(df_filt1['MUNICIPIO'].unique().tolist()) if not df.empty else []
    municipios_sel = st.multiselect("🏙️ Município:", lista_municipios)
    df_filt2 = df_filt1[df_filt1['MUNICIPIO'].isin(municipios_sel)] if municipios_sel and not df.empty else df_filt1
    
    lista_alimentadores = sorted(df_filt2['ALIMENTADOR'].unique().tolist()) if not df.empty else []
    alim_sel = st.multiselect("⚡ Alimentador:", lista_alimentadores)
    
    alimentadores_visiveis = alim_sel if alim_sel else lista_alimentadores

    camadas_ativas = {}
    if not df.empty:
        st.markdown("---")
        st.markdown("### 🗂️ 4. Camadas (Desempenho)")
        
        for alim in alimentadores_visiveis:
            st.markdown(f"**{alim}**")
            lista_camadas_alim = sorted(df[df['ALIMENTADOR'] == alim]['TIPO_REDE'].unique().tolist())
            
            camadas_essenciais = ['REDE PRIMÁRIA', 'REDE PRIMARIA', 'REDE SECUNDÁRIA', 'REDE SECUNDARIA', 'TRANSFORMADOR', 'POSTE']
            camadas_default = [c for c in lista_camadas_alim if c in camadas_essenciais]
            
            camadas_ativas[alim] = st.multiselect(
                "Ligado/Desligado:", 
                lista_camadas_alim, 
                default=camadas_default,
                key=f"ms_{alim}"
            )
            
        st.markdown("---")
        st.markdown("### 🗑️ Gerenciar Malha Local")
        alim_para_deletar = st.selectbox("Apagar Alimentador do Banco:", ["Selecione..."] + sorted(df['ALIMENTADOR'].unique().tolist()))
        if alim_para_deletar != "Selecione...":
            if st.button("❌ Excluir Permanentemente", use_container_width=True):
                caminho_del = f"database/redes/{alim_para_deletar}.pkl"
                if os.path.exists(caminho_del):
                    os.remove(caminho_del)
                    carregar_banco_redes.clear()
                    st.success("Excluído!")
                    time.sleep(1)
                    st.rerun()

    st.markdown("---")
    tipo_mapa = st.radio("Visual do Mapa:", ["🗺️ Satélite (Real)", "🛣️ Vetorial (Rápido)"])
    map_style_pdk = "mapbox://styles/mapbox/satellite-v9" if "Satélite" in tipo_mapa else "mapbox://styles/mapbox/light-v9"

# ==========================================
# 3. MOTOR GPU / BUSCA E RENDERIZAÇÃO
# ==========================================
layers = []

# --- 1. CAMADA BASE IBGE (MARANHÃO DIVIDIDO POR REGIONAIS) ---
geo_data = get_base_geojson()
if geo_data:
    geo_render = copy.deepcopy(geo_data)
    reg_colors = {
        'LESTE': [67, 99, 216],     # Azul
        'CENTRO': [230, 25, 75],    # Vermelho
        'NOROESTE': [255, 225, 25], # Amarelo
        'NORTE': [245, 130, 49],    # Laranja
        'SUL': [60, 180, 75]        # Verde
    }
    
    for feature in geo_render['features']:
        reg = feature['properties']['REGIONAL']
        base_color = reg_colors.get(reg, [150, 150, 150]) # Cinza padrão
        
        # Lógica de Relevo e Destaque
        if regioes_sel:
            if reg in regioes_sel:
                feature['properties']['fill_color'] = base_color + [150] # Fica forte
                feature['properties']['elevation'] = 3000 # Salta 3KM pra fora da tela
            else:
                feature['properties']['fill_color'] = base_color + [15]  # Quase transparente
                feature['properties']['elevation'] = 0
        else:
            feature['properties']['fill_color'] = base_color + [60] # Opacidade leve para ver o satélite por baixo
            feature['properties']['elevation'] = 0

    geojson_layer = pdk.Layer(
        "GeoJsonLayer",
        geo_render,
        opacity=1,
        stroked=True,
        filled=True,
        extruded=True,
        wireframe=True,
        get_elevation="properties.elevation",
        get_fill_color="properties.fill_color",
        get_line_color=[255, 255, 255, 80],
        pickable=True
    )
    layers.append(geojson_layer)

# --- 2. CAMADAS DE REDE (KMZ) ---
if not df.empty:
    df_mapa = df.copy()
    if regioes_sel: df_mapa = df_mapa[df_mapa['REGIONAL'].isin(regioes_sel)]
    if municipios_sel: df_mapa = df_mapa[df_mapa['MUNICIPIO'].isin(municipios_sel)]
    df_mapa = df_mapa[df_mapa['ALIMENTADOR'].isin(alimentadores_visiveis)]

    mask_camadas = pd.Series(False, index=df_mapa.index)
    for alim in alimentadores_visiveis:
        if alim in camadas_ativas:
            mask_camadas = mask_camadas | ((df_mapa['ALIMENTADOR'] == alim) & (df_mapa['TIPO_REDE'].isin(camadas_ativas[alim])))
    df_mapa = df_mapa[mask_camadas]

    # --- PROCESSADOR DE BUSCAS ---
    df_busca = pd.DataFrame()
    nearest_idx = None

    if busca_lat is not None and busca_lon is not None and not df_mapa.empty:
        def calc_min_dist(row):
            if row['TIPO_GEOMETRIA'] == 'Ponto':
                return haversine(busca_lat, busca_lon, row['COORDS'][1], row['COORDS'][0])
            else:
                return min([haversine(busca_lat, busca_lon, pt[1], pt[0]) for pt in row['COORDS']])
        
        df_mapa['DISTANCIA_KM'] = df_mapa.apply(calc_min_dist, axis=1)
        nearest_idx = df_mapa['DISTANCIA_KM'].idxmin()
        dist_metros = df_mapa.loc[nearest_idx, 'DISTANCIA_KM'] * 1000
        
        elem_prox = df_mapa.loc[nearest_idx]
        st.sidebar.success(f"🎯 **Alvo mais próximo:** {elem_prox['TIPO_REDE']} ({elem_prox['NOME']}) a {dist_metros:.1f} metros.")
        
        df_busca = df_mapa.loc[[nearest_idx]]
        df_mapa = df_mapa.drop(nearest_idx)

    elif termo_pesquisa != "":
        mask_nome = df_mapa['NOME'].str.contains(termo_pesquisa, case=False, na=False)
        df_busca = df_mapa[mask_nome]
        df_mapa = df_mapa[~mask_nome]

    df_linhas = df_mapa[df_mapa['TIPO_GEOMETRIA'] == 'Linha']
    df_pontos = df_mapa[df_mapa['TIPO_GEOMETRIA'] == 'Ponto']

    if not df_linhas.empty:
        layers.append(pdk.Layer("PathLayer", data=df_linhas, pickable=True, get_color="COR", width_scale=1, width_min_pixels=1, get_path="COORDS", get_width="PESO"))
    
    if not df_pontos.empty:
        layers.append(pdk.Layer("ScatterplotLayer", data=df_pontos, pickable=True, get_position="COORDS", get_color="COR", get_radius="RAIO", radius_scale=1, radius_min_pixels=2, radius_max_pixels=6))
        
    if not df_busca.empty:
        df_busca_linhas = df_busca[df_busca['TIPO_GEOMETRIA'] == 'Linha']
        df_busca_pontos = df_busca[df_busca['TIPO_GEOMETRIA'] == 'Ponto']
        
        if not df_busca_linhas.empty:
            layers.append(pdk.Layer("PathLayer", data=df_busca_linhas, pickable=True, get_color=[255, 0, 255], width_scale=1, width_min_pixels=4, get_path="COORDS", get_width=8))
        if not df_busca_pontos.empty:
            layers.append(pdk.Layer("ScatterplotLayer", data=df_busca_pontos, pickable=True, get_position="COORDS", get_color=[255, 0, 255], radius_min_pixels=8, radius_max_pixels=15))

    # --- PINO DOURADO DA COORDENADA PESQUISADA ---
    if busca_lat is not None and busca_lon is not None:
        target_df = pd.DataFrame([{"COORDS": [busca_lon, busca_lat], "NOME": "Sua Pesquisa GPS", "TIPO_REDE": "ALVO", "ALIMENTADOR": "N/A"}])
        layers.append(pdk.Layer("ScatterplotLayer", data=target_df, pickable=True, get_position="COORDS", get_color=[255, 215, 0], get_radius=15, radius_min_pixels=10, radius_max_pixels=20))

# --- MOVIMENTAÇÃO DE CÂMERA ---
if busca_lat is not None and busca_lon is not None:
    view_state = pdk.ViewState(latitude=busca_lat, longitude=busca_lon, zoom=18, pitch=45)
elif not df.empty and 'df_busca' in locals() and not df_busca.empty and not df_busca[df_busca['TIPO_GEOMETRIA'] == 'Ponto'].empty:
    alvo = df_busca[df_busca['TIPO_GEOMETRIA'] == 'Ponto'].iloc[0]['COORDS']
    view_state = pdk.ViewState(latitude=alvo[1], longitude=alvo[0], zoom=18, pitch=45)
elif not df.empty and 'df_pontos' in locals() and not df_pontos.empty:
    centro_lon = df_pontos['COORDS'].apply(lambda x: x[0]).mean()
    centro_lat = df_pontos['COORDS'].apply(lambda x: x[1]).mean()
    view_state = pdk.ViewState(latitude=centro_lat, longitude=centro_lon, zoom=12, pitch=0)
else:
    view_state = pdk.ViewState(latitude=-5.2, longitude=-45.0, zoom=6, pitch=10)

tooltip_html = {
    "html": "<b>{TIPO_REDE}</b><br/><b>Identificação:</b> {NOME}<br/><b>Base:</b> {ALIMENTADOR}",
    "style": {"backgroundColor": "steelblue", "color": "white", "fontFamily": "sans-serif"}
}

r = pdk.Deck(
    layers=layers,
    initial_view_state=view_state,
    map_style=map_style_pdk,
    tooltip=tooltip_html
)

st.pydeck_chart(r)
