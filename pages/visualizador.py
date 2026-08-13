import streamlit as st
import pandas as pd
import io
import zipfile
import re
import time
import os
import xml.etree.ElementTree as ET
import math
import requests
import unicodedata
import folium
from streamlit_folium import st_folium
import html

st.set_page_config(page_title="Visualizador de Malha", page_icon="🗺️", layout="wide")

if not os.path.exists("database/redes"):
    os.makedirs("database/redes", exist_ok=True)

# ==========================================
# 1. FUNÇÕES BASE E PLANILHA
# ==========================================
def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

@st.cache_data(show_spinner=False)
def load_base_mapping():
    """Lê a BASE.xlsx para descobrir exatamente qual município pertence a qual regional"""
    mun_to_reg = {}
    try:
        if os.path.exists("BASE.xlsx"):
            df_base = pd.read_excel("BASE.xlsx")
            for _, row in df_base.iterrows():
                mun = remove_accents(row.get('MunicIpio', '')).upper().strip()
                reg = str(row.get('Regional', '')).strip().upper()
                if mun and reg and reg != 'NAN':
                    mun_to_reg[mun] = reg
    except:
        pass
    
    # Forçar Overrides (Correções Manuais solicitadas)
    overrides_centro = [
        'SANTA LUZIA', 'CONCEICAO DO LAGO-ACU', 'CONCEICAO DO LAGO ACU', 
        'PINDARE-MIRIM', 'PINDARE MIRIM', 'OLHO DAGUA DAS CUNHAS', 
        'OLHO D\'AGUA DAS CUNHAS', 'GOVERNADOR LUIZ ROCHA'
    ]
    for mun in overrides_centro:
        mun_to_reg[mun] = 'CENTRO'
        
    return mun_to_reg

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

def extrair_coordenadas_vis(texto_coords):
    pontos = []
    for coord in texto_coords.strip().split():
        partes = coord.split(',')
        if len(partes) >= 2:
            try:
                lon = float(partes[0].strip().replace(',', '.'))
                lat = float(partes[1].strip().replace(',', '.'))
                if lat != 0.0 and lon != 0.0 and -35.0 <= lat <= 5.0 and -75.0 <= lon <= -30.0:
                    pontos.append([lat, lon]) 
            except:
                continue
    return pontos

def processar_e_salvar_kmz(arquivos):
    base_map = load_base_mapping()
    dict_cores = {
        'REDE PRIMÁRIA': '#e6194b', 'REDE PRIMARIA': '#e6194b',
        'REDE SECUNDÁRIA': '#4363d8', 'REDE SECUNDARIA': '#4363d8',
        'POSTE': '#808080', 'TRANSFORMADOR': '#f58231', 
        'CHAVE': '#3cb44b', 'REGULADOR': '#911eb4', 
        'RELIGADOR': '#46f0f0', 'SUBESTAÇÃO': '#000000', 'SUBESTACAO': '#000000'
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
        if mun_match: 
            municipio = mun_match.group(1).strip().upper()
            mun_norm = remove_accents(municipio)
            if mun_norm in base_map:
                regional = base_map[mun_norm]

        if regional == "N/A":
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

                cor_elemento = dict_cores.get(nome_pasta, '#333333')
                
                for placemark in folder.findall('.//Placemark'):
                    pm_name_tag = placemark.find('name')
                    nome_elemento = pm_name_tag.text.strip() if pm_name_tag is not None and pm_name_tag.text else "S/N"
                    
                    for ls in placemark.findall('.//LineString/coordinates'):
                        if ls.text:
                            coords = extrair_coordenadas_vis(ls.text)
                            if len(coords) > 1:
                                registros_flat.append({
                                    'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio,
                                    'TIPO_GEOMETRIA': 'Linha', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento,
                                    'COORDS': coords, 'COR': cor_elemento
                                })
                            
                    for pt in placemark.findall('.//Point/coordinates'):
                        if pt.text:
                            coords = extrair_coordenadas_vis(pt.text)
                            if len(coords) > 0:
                                registros_flat.append({
                                    'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio,
                                    'TIPO_GEOMETRIA': 'Ponto', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento,
                                    'COORDS': coords[0], 'COR': cor_elemento
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

@st.cache_data(show_spinner=False)
def get_base_geojson():
    mun_to_reg = load_base_mapping()
    url_geojson = "https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-21-mun.json"
    try:
        resp = requests.get(url_geojson, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
        geo_data = resp.json()
    except:
        return None

    # Cores Exatas da sua Imagem
    reg_colors = {
        'LESTE': '#1f77b4',     # Azul
        'CENTRO': '#d62728',    # Vermelho
        'NOROESTE': '#ffed6f',  # Amarelo
        'NORTE': '#ff7f0e',     # Laranja
        'SUL': '#8fbc8f',       # Verde Musgo
        'DESCONHECIDO': '#cccccc'
    }

    for feature in geo_data['features']:
        mun_name = feature['properties']['name']
        mun_name_norm = remove_accents(mun_name).upper().strip()
        reg = mun_to_reg.get(mun_name_norm, "DESCONHECIDO")
        
        feature['properties']['REGIONAL'] = reg
        feature['properties']['MUNICIPIO'] = mun_name_norm
        feature['properties']['fillColor'] = reg_colors.get(reg, '#cccccc')
        
    return geo_data

# ==========================================
# 2. INTERFACE E FILTROS DO SISTEMA
# ==========================================
st.markdown("<h2 style='color: #0D256C;'>🗺️ Visualizador Oficial de Malha (Satélite Integrado)</h2>", unsafe_allow_html=True)
st.markdown("O sistema cruza as regionais oficiais com a sua planilha `BASE.xlsx`. Use o zoom para ver o satélite em alta definição!")

df = carregar_banco_redes()
base_map = load_base_mapping()

with st.sidebar:
    st.markdown("### 📥 1. Upload de Redes")
    arquivos_upados = st.file_uploader("Arraste novos KMZs aqui", type=["kmz", "kml"], accept_multiple_files=True)
    if arquivos_upados:
        if st.button("💾 Processar e Salvar no Banco", type="primary", use_container_width=True):
            with st.spinner("Lendo propriedades XML e gravando em disco..."):
                qtd = processar_e_salvar_kmz(arquivos_upados)
            if qtd > 0:
                st.success(f"✅ {qtd} novos Alimentadores salvos!")
                carregar_banco_redes.clear()
                time.sleep(1)
                st.rerun()

    st.markdown("---")
    st.markdown("### 🔎 2. Pesquisas Inteligentes")
    tab_nome, tab_coord = st.tabs(["📝 Por Nome/ID", "📍 Por Coordenada"])
    
    termo_pesquisa = ""
    busca_lat = None
    busca_lon = None
    
    with tab_nome:
        termo_pesquisa = st.text_input("Nome/Num. Poste ou Trafo:", placeholder="Ex: 554930...").strip().upper()
        
    with tab_coord:
        c_lat, c_lon = st.columns(2)
        with c_lat:
            lat_input = st.text_input("Latitude:", placeholder="Ex: -5.532")
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
                st.warning("⚠️ Formato inválido. Use números.")

    st.markdown("---")
    st.markdown("### 🔍 3. Filtros Geográficos")
    
    # Lista de Regionais baseada na BASE.xlsx
    lista_regioes = sorted(list(set(base_map.values()))) if base_map else ["CENTRO", "LESTE", "NOROESTE", "NORTE", "SUL"]
    regioes_sel = st.multiselect("📍 Regional:", lista_regioes)
    
    # Filtra municípios da BASE.xlsx pela Regional selecionada
    lista_municipios = []
    for mun, reg in base_map.items():
        if not regioes_sel or reg in regioes_sel:
            lista_municipios.append(mun)
    lista_municipios = sorted(lista_municipios)
    
    municipios_sel = st.multiselect("🏙️ Município (Foco e Contorno):", lista_municipios)
    
    df_filt = df.copy()
    if not df.empty:
        if regioes_sel: df_filt = df_filt[df_filt['REGIONAL'].isin(regioes_sel)]
        if municipios_sel: df_filt = df_filt[df_filt['MUNICIPIO'].isin(municipios_sel)]
    
    lista_alimentadores = sorted(df_filt['ALIMENTADOR'].unique().tolist()) if not df_filt.empty else []
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
            camadas_ativas[alim] = st.multiselect("Visibilidade:", lista_camadas_alim, default=camadas_default, key=f"ms_{alim}")
            
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

# ==========================================
# 3. CONSTRUÇÃO DO MAPA FOLIUM
# ==========================================
mapa = folium.Map(location=[-5.2, -45.0], zoom_start=6, tiles=None)

# Camada Híbrida do Google (Satélite + Ruas + Nomes)
folium.TileLayer(
    tiles='https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}',
    attr='Google',
    name='Satélite (Google Maps)',
    overlay=False,
    control=True,
    max_zoom=20
).add_to(mapa)

folium.TileLayer(
    tiles='CartoDB positron',
    name='Mapa Base (Limpo)',
    overlay=False,
    control=True,
    max_zoom=20
).add_to(mapa)

geo_data = get_base_geojson()
if geo_data:
    def style_function(feature):
        reg_mun = feature['properties'].get('MUNICIPIO', '')
        reg_name = feature['properties'].get('REGIONAL', '')
        
        if municipios_sel:
            if reg_mun in municipios_sel:
                # Município Focado: Transparente com Borda Grossa Magenta (Para ver o satélite dentro)
                return {'fillColor': 'transparent', 'color': '#FF00FF', 'weight': 4, 'fillOpacity': 0}
            else:
                return {'fillColor': 'transparent', 'color': 'transparent', 'weight': 0}
                
        elif regioes_sel:
            if reg_name in regioes_sel:
                return {'fillColor': feature['properties']['fillColor'], 'color': '#000000', 'weight': 1, 'fillOpacity': 0.6}
            else:
                return {'fillColor': 'transparent', 'color': 'transparent', 'weight': 0}
        
        # Padrão: Estado inteiro colorido
        return {'fillColor': feature['properties']['fillColor'], 'color': '#000000', 'weight': 1, 'fillOpacity': 0.5}

    # Adiciona a camada IBGE
    folium.GeoJson(
        geo_data,
        name="Divisão IBGE (Maranhão)",
        style_function=style_function,
        tooltip=folium.features.GeoJsonTooltip(fields=['name', 'REGIONAL'], aliases=['Município:', 'Regional:'], style="background-color: white; color: #333; font-family: arial; font-size: 12px; padding: 10px;"),
        zoom_on_click=False,
        show=True
    ).add_to(mapa)

    # 🚀 SCRIPT NINJA: Faz a cor do estado sumir se der zoom maior que 9!
    map_id = mapa.get_name()
    js_zoom_hide = f"""
    <script>
        setTimeout(function() {{
            var ibge_layer_{map_id} = null;
            {map_id}.eachLayer(function(layer) {{
                if (layer.options && layer.options.name === 'Divisão IBGE (Maranhão)') {{
                    ibge_layer_{map_id} = layer;
                }}
            }});
            {map_id}.on('zoomend', function() {{
                if (ibge_layer_{map_id}) {{
                    if ({map_id}.getZoom() > 9) {{
                        if ({map_id}.hasLayer(ibge_layer_{map_id})) {{
                            {map_id}.removeLayer(ibge_layer_{map_id});
                        }}
                    }} else {{
                        if (!{map_id}.hasLayer(ibge_layer_{map_id})) {{
                            {map_id}.addLayer(ibge_layer_{map_id});
                        }}
                    }}
                }}
            }});
        }}, 500);
    </script>
    """
    mapa.get_root().html.add_child(folium.Element(js_zoom_hide))


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

    df_busca = pd.DataFrame()
    nearest_idx = None
    todas_lats, todas_lons = [], []

    if busca_lat is not None and busca_lon is not None and not df_mapa.empty:
        def calc_min_dist(row):
            if row['TIPO_GEOMETRIA'] == 'Ponto': return haversine(busca_lat, busca_lon, row['COORDS'][0], row['COORDS'][1])
            else: return min([haversine(busca_lat, busca_lon, pt[0], pt[1]) for pt in row['COORDS']])
        
        df_mapa['DISTANCIA_KM'] = df_mapa.apply(calc_min_dist, axis=1)
        nearest_idx = df_mapa['DISTANCIA_KM'].idxmin()
        dist_metros = df_mapa.loc[nearest_idx, 'DISTANCIA_KM'] * 1000
        
        elem_prox = df_mapa.loc[nearest_idx]
        st.sidebar.success(f"🎯 **Alvo mais próximo:** {elem_prox['TIPO_REDE']} ({elem_prox['NOME']}) a {dist_metros:.1f} metros.")
        
        df_busca = df_mapa.loc[[nearest_idx]]
        df_mapa = df_mapa.drop(nearest_idx)

    elif termo_pesquisa != "":
        mask_nome = df_mapa['NOME'].astype(str).str.contains(termo_pesquisa, case=False, na=False)
        df_busca = df_mapa[mask_nome]
        df_mapa = df_mapa[~mask_nome]

    fg_rede = folium.FeatureGroup(name="Rede Elétrica (Malha)")
    
    def criar_popup(row):
        coord_txt = f"{row['COORDS'][0]:.5f}, {row['COORDS'][1]:.5f}" if row['TIPO_GEOMETRIA'] == 'Ponto' else "Linha de Múltiplos Pontos"
        html_popup = f"""
        <div style="min-width: 250px; font-family: sans-serif;">
            <h4 style="margin-top: 0; color: {row['COR']}; border-bottom: 2px solid {row['COR']}; padding-bottom: 5px;">{row['TIPO_REDE']}</h4>
            <table style="width:100%;">
                <tr><td style="color: #555; padding: 2px;"><b>IDENTIFICAÇÃO:</b></td><td>{html.escape(str(row['NOME']))}</td></tr>
                <tr><td style="color: #555; padding: 2px;"><b>ALIMENTADOR:</b></td><td>{html.escape(str(row['ALIMENTADOR']))}</td></tr>
                <tr><td style="color: #555; padding: 2px;"><b>LOCAL:</b></td><td>{html.escape(str(row['MUNICIPIO']))} - {row['REGIONAL']}</td></tr>
                <tr><td style="color: #555; padding: 2px;"><b>GPS:</b></td><td>{coord_txt}</td></tr>
            </table>
        </div>
        """
        return folium.Popup(html_popup, max_width=350)

    for _, row in df_mapa.iterrows():
        if row['TIPO_GEOMETRIA'] == 'Linha':
            folium.PolyLine(
                locations=row['COORDS'],
                color=row['COR'],
                weight=3 if 'PRIM' in row['TIPO_REDE'] else 2,
                opacity=0.8,
                popup=criar_popup(row),
                tooltip=f"<b>{row['TIPO_REDE']}</b><br>{html.escape(str(row['NOME']))}"
            ).add_to(fg_rede)
            for pt in row['COORDS']: todas_lats.append(pt[0]); todas_lons.append(pt[1])
        else:
            folium.CircleMarker(
                location=row['COORDS'],
                radius=4 if 'TRANSFORMADOR' in row['TIPO_REDE'] else 2,
                color=row['COR'],
                fill=True,
                fillOpacity=1.0,
                popup=criar_popup(row),
                tooltip=f"<b>{row['TIPO_REDE']}</b><br>{html.escape(str(row['NOME']))}"
            ).add_to(fg_rede)
            todas_lats.append(row['COORDS'][0]); todas_lons.append(row['COORDS'][1])

    busca_lats, busca_lons = [], []
    for _, row in df_busca.iterrows():
        if row['TIPO_GEOMETRIA'] == 'Linha':
            folium.PolyLine(
                locations=row['COORDS'],
                color='#FF00FF', weight=8, opacity=1.0,
                popup=criar_popup(row),
                tooltip=f"ALVO ENCONTRADO: {html.escape(str(row['NOME']))}"
            ).add_to(fg_rede)
            for pt in row['COORDS']: busca_lats.append(pt[0]); busca_lons.append(pt[1])
        else:
            folium.Marker(
                location=row['COORDS'],
                icon=folium.Icon(color='purple', icon='star'),
                popup=criar_popup(row),
                tooltip=f"ALVO ENCONTRADO: {html.escape(str(row['NOME']))}"
            ).add_to(fg_rede)
            busca_lats.append(row['COORDS'][0]); busca_lons.append(row['COORDS'][1])

    if busca_lat is not None and busca_lon is not None:
        folium.Marker(
            location=[busca_lat, busca_lon],
            icon=folium.Icon(color='orange', icon='map-pin', prefix='fa'),
            tooltip="Sua Pesquisa GPS"
        ).add_to(fg_rede)

    fg_rede.add_to(mapa)

    # Lógica Automática de Foco de Câmera
    if busca_lat is not None and busca_lon is not None:
        mapa.fit_bounds([[busca_lat - 0.001, busca_lon - 0.001], [busca_lat + 0.001, busca_lon + 0.001]])
    elif busca_lats and busca_lons:
        mapa.fit_bounds([[min(busca_lats), min(busca_lons)], [max(busca_lats), max(busca_lons)]])
    elif municipios_sel and geo_data:
        mun_foco_lats, mun_foco_lons = [], []
        for feature in geo_data['features']:
            if feature['properties'].get('MUNICIPIO') in municipios_sel:
                geom = feature['geometry']
                if geom['type'] == 'Polygon':
                    for pt in geom['coordinates'][0]:
                        mun_foco_lats.append(pt[1]); mun_foco_lons.append(pt[0])
                elif geom['type'] == 'MultiPolygon':
                    for poly in geom['coordinates']:
                        for pt in poly[0]:
                            mun_foco_lats.append(pt[1]); mun_foco_lons.append(pt[0])
        if mun_foco_lats and mun_foco_lons:
            mapa.fit_bounds([[min(mun_foco_lats), min(mun_foco_lons)], [max(mun_foco_lats), max(mun_foco_lons)]])
        elif todas_lats and todas_lons:
            mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])
    elif todas_lats and todas_lons:
        mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])

folium.LayerControl(position='topright').add_to(mapa)

# Tamanho Gigante para Ocupar a Tela Inteira
st_folium(mapa, use_container_width=True, height=850, returned_objects=[])
