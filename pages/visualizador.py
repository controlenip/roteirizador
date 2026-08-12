import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import zipfile
import re
import numpy as np
import io
import xml.etree.ElementTree as ET

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Visualizador de Malha Elétrica",
    page_icon="⚡",
    layout="wide"
)

# ==========================================
# 2. MOTOR DE EXTRAÇÃO ESTRUTURADA (COM PASTAS)
# ==========================================
def extrair_coordenadas(texto_coords):
    """Filtro Antilixo: Converte texto de coordenadas e joga fora zeros e Ásia/África"""
    pontos = []
    coordenadas_brutas = texto_coords.strip().split()
    for coord in coordenadas_brutas:
        partes = coord.split(',')
        if len(partes) >= 2:
            try:
                lon = float(partes[0].strip())
                lat = float(partes[1].strip())
                # Filtro Brasil: Lat -35 a 5, Lon -75 a -30
                if lat != 0.0 and lon != 0.0 and -35.0 <= lat <= 5.0 and -75.0 <= lon <= -30.0:
                    pontos.append([lat, lon]) # Folium usa [Lat, Lon]
            except:
                continue
    return pontos

@st.cache_data(show_spinner=False)
def processar_arquivos_kmz(arquivos):
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
        
        # Limpeza para facilitar o parsing do XML
        conteudo_kml = re.sub(r'\sxmlns="[^"]+"', '', conteudo_kml, count=1)
        
        try:
            root = ET.fromstring(conteudo_kml)
        except Exception:
            continue # Pula se o KML estiver corrompido

        # Extrair Metadados
        municipio = "N/A"
        regional = "N/A"
        mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if mun_match: municipio = mun_match.group(1).strip().upper()
        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if reg_match: regional = reg_match.group(1).strip().upper()
        
        if regional == "N/A":
            sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
            if sigla_match: regional = sigla_match.group(1)

        # ----------------------------------------------------
        # NAVEGAÇÃO NAS PASTAS DO KML (O SEGREDO DO GOOGLE EARTH)
        # ----------------------------------------------------
        camadas_do_alimentador = {}
        
        # Procura por todas as pastas dentro do KML
        for folder in root.findall('.//Folder'):
            name_tag = folder.find('name')
            if name_tag is not None and name_tag.text:
                nome_pasta = name_tag.text.strip().upper()
                
                # Vamos ignorar pastas raiz que só servem de container
                if "CEMAR" in nome_pasta or nome_pasta == nome_arquivo:
                    continue

                linhas = []
                pontos = []
                
                # Lê os Placemarks dentro desta pasta específica
                for placemark in folder.findall('.//Placemark'):
                    # Busca Linhas (Ex: Rede Primária, Secundária)
                    for ls in placemark.findall('.//LineString/coordinates'):
                        coords_linha = extrair_coordenadas(ls.text)
                        if len(coords_linha) > 1:
                            linhas.append(coords_linha)
                            
                    # Busca Pontos (Ex: Poste, Chave, Trafo)
                    for pt in placemark.findall('.//Point/coordinates'):
                        coords_ponto = extrair_coordenadas(pt.text)
                        if len(coords_ponto) > 0:
                            pontos.append(coords_ponto[0]) # Ponto único
                
                # Se a pasta tiver conteúdo, salva no dicionário deste alimentador
                if linhas or pontos:
                    camadas_do_alimentador[nome_pasta] = {
                        'linhas': linhas,
                        'pontos': pontos
                    }
        
        # Se achou camadas, salva o alimentador
        if camadas_do_alimentador:
            dados_extraidos.append({
                'ALIMENTADOR': nome_arquivo,
                'REGIONAL': regional,
                'MUNICIPIO': municipio,
                'CAMADAS': camadas_do_alimentador
            })
            
    return pd.DataFrame(dados_extraidos)

# ==========================================
# 3. INTERFACE E LÓGICA DO MENU
# ==========================================
def view_visualizador():
    st.markdown("<h2 style='color: #0D256C;'>🗺️ Inspeção de Malha Elétrica (KMZ)</h2>", unsafe_allow_html=True)
    st.markdown("Faça o upload dos arquivos KMZ. O sistema recriará as camadas (Rede, Postes, Chaves) exatamente como no Google Earth.")

    if 'df_rede_vis' not in st.session_state:
        st.session_state.df_rede_vis = pd.DataFrame()

    with st.sidebar:
        st.markdown("### 📥 1. Upload de Malha")
        arquivos_upados = st.file_uploader("Selecione os Alimentadores", type=["kmz", "kml"], accept_multiple_files=True)
        
        if st.button("⚙️ Processar Arquivos", type="primary", use_container_width=True):
            if arquivos_upados:
                with st.spinner("Decodificando camadas e equipamentos..."):
                    df_extraido = processar_arquivos_kmz(arquivos_upados)
                    st.session_state.df_rede_vis = df_extraido
                if not df_extraido.empty:
                    st.success(f"✅ {len(df_extraido)} Alimentadores processados!")
                else:
                    st.error("Nenhuma rede válida encontrada.")
            else:
                st.warning("Suba ao menos um arquivo KMZ.")
                
        df = st.session_state.df_rede_vis
        
        alim_sel = []
        camadas_ativas = {}
        
        if not df.empty:
            st.markdown("---")
            st.markdown("### ⚡ 2. Seleção de Alimentadores")
            
            lista_alimentadores = sorted(df['ALIMENTADOR'].unique().tolist())
            alim_sel = st.multiselect("Selecione os alimentadores para exibir no mapa:", lista_alimentadores, default=lista_alimentadores[0] if lista_alimentadores else None)
            
            st.markdown("---")
            if alim_sel:
                st.markdown("### 🗂️ 3. Camadas (Google Earth)")
                st.caption("Ligue ou desligue elementos específicos da rede.")
                
                # CRIA UM MENU DINÂMICO PARA CADA ALIMENTADOR SELECIONADO
                for alim in alim_sel:
                    st.markdown(f"**{alim}**")
                    dict_camadas = df[df['ALIMENTADOR'] == alim]['CAMADAS'].iloc[0]
                    lista_camadas_alim = sorted(list(dict_camadas.keys()))
                    
                    # Seleciona quais pastas exibir (padrão: mostrar tudo)
                    camadas_ativas[alim] = st.multiselect(
                        f"Ocultar/Exibir em {alim}:", 
                        lista_camadas_alim, 
                        default=lista_camadas_alim,
                        key=f"ms_{alim}"
                    )

    # ==========================================
    # 4. RENDERIZAÇÃO INTELIGENTE DO MAPA
    # ==========================================
    if not st.session_state.df_rede_vis.empty and alim_sel:
        # Define cores para elementos comuns para manter o mapa legível
        dict_cores = {
            'REDE PRIMÁRIA': '#ff0000', # Vermelho
            'REDE PRIMARIA': '#ff0000',
            'REDE SECUNDÁRIA': '#0000ff', # Azul
            'REDE SECUNDARIA': '#0000ff',
            'POSTE': '#808080', # Cinza
            'TRANSFORMADOR': '#ff9900', # Laranja
            'CHAVE': '#00ff00' # Verde
        }

        todas_lats = []
        todas_lons = []
        
        mapa = folium.Map(tiles="CartoDB positron")
        
        for alim in alim_sel:
            dict_camadas = df[df['ALIMENTADOR'] == alim]['CAMADAS'].iloc[0]
            camadas_permitidas = camadas_ativas.get(alim, [])
            
            # FeatureGroup master para o Alimentador
            fg_alim = folium.FeatureGroup(name=f"⚡ {alim}", show=True)
            
            for nome_pasta, conteudos in dict_camadas.items():
                if nome_pasta not in camadas_permitidas:
                    continue # Pula se o usuário desmarcou no menu lateral
                    
                cor_elemento = dict_cores.get(nome_pasta, '#333333') # Preto por padrão se for algo novo
                
                # 1. Desenha as Linhas (Cables/Rede)
                for linha in conteudos['linhas']:
                    folium.PolyLine(
                        locations=linha,
                        color=cor_elemento,
                        weight=3 if 'PRIM' in nome_pasta else 2, # Rede primária mais grossa
                        opacity=0.9,
                        tooltip=f"{alim}<br><b>{nome_pasta}</b>"
                    ).add_to(fg_alim)
                    for pt in linha:
                        todas_lats.append(pt[0])
                        todas_lons.append(pt[1])
                
                # 2. Desenha os Pontos (Postes, Trafos, Chaves)
                for ponto in conteudos['pontos']:
                    # Usa CircleMarker para não pesar o navegador com milhares de ícones grandes
                    folium.CircleMarker(
                        location=ponto,
                        radius=4 if 'TRANSFORMADOR' in nome_pasta else 2,
                        color=cor_elemento,
                        fill=True,
                        fill_color=cor_elemento,
                        fill_opacity=1.0,
                        tooltip=f"{alim}<br><b>{nome_pasta}</b>"
                    ).add_to(fg_alim)
                    todas_lats.append(ponto[0])
                    todas_lons.append(ponto[1])
                    
            fg_alim.add_to(mapa)

        # Centraliza e ajusta o Zoom magicamente
        if todas_lats and todas_lons:
            mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])
            
        folium.LayerControl().add_to(mapa)
        st_folium(mapa, use_container_width=True, height=650, returned_objects=[])

if __name__ == "__main__":
    view_visualizador()
