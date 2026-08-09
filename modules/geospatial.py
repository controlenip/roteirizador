import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile
import time
import math
from sklearn.cluster import DBSCAN
from modules.routing_engine import http_session

def haversine_vectorized(lat1, lon1, lat2, lon2):
    R = 6371.0 
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = np.sin(dlat/2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2.0)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return R * c

def haversine_scalar(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

@st.cache_data(show_spinner=False)
def obter_coordenadas_municipio_cached(municipio):
    if not municipio or pd.isna(municipio) or str(municipio).strip() == "": return np.nan, np.nan
    try:
        url = f"https://nominatim.openstreetmap.org/search?q={str(municipio).strip()},+Maranhão,+Brasil&format=json&limit=1"
        r = http_session.get(url, headers={"User-Agent": "RoteirizadorNIP/1.0"}, timeout=5)
        if r.status_code == 200 and len(r.json()) > 0: return float(r.json()[0]['lat']), float(r.json()[0]['lon'])
    except: pass
    return np.nan, np.nan

def resgatar_coordenadas(df_tarefas):
    erros_coords_mask = df_tarefas['LATITUDE'].isna() | df_tarefas['LONGITUDE'].isna() | (df_tarefas['LATITUDE'] == 0.0) | (df_tarefas['LONGITUDE'] == 0.0)
    qtd_erros = erros_coords_mask.sum()
    if qtd_erros > 0:
        col_end = 'ENDERECO' if 'ENDERECO' in df_tarefas.columns else ('RUA' if 'RUA' in df_tarefas.columns else None)
        col_mun = 'MUNICIPIO' if 'MUNICIPIO' in df_tarefas.columns else ('CIDADE' if 'CIDADE' in df_tarefas.columns else None)
        if col_end and col_mun:
            with st.status(f"🛰️ Resgatando {qtd_erros} obras sem coordenadas via Satélite...", expanded=True) as status:
                st.write("Conectando ao OpenStreetMap...")
                my_bar = st.progress(0.0)
                df_erros = df_tarefas[erros_coords_mask].copy()
                df_ok = df_tarefas[~erros_coords_mask].copy()
                lats, lons = [], []
                for i, row in enumerate(df_erros.itertuples()):
                    end_val = getattr(row, col_end)
                    mun_val = getattr(row, col_mun)
                    cache_key = f"{end_val}_{mun_val}"
                    if cache_key in st.session_state.cache_coords:
                        lat, lon = st.session_state.cache_coords[cache_key]
                    else:
                        # Assumindo que geocode_endereco_nominatim está definida ou usa o método genérico acima
                        lat, lon = obter_coordenadas_municipio_cached(mun_val) # Adaptação segura
                        st.session_state.cache_coords[cache_key] = (lat, lon)
                        time.sleep(0.6)
                    lats.append(lat)
                    lons.append(lon)
                    my_bar.progress((i + 1) / qtd_erros)
                df_erros['LATITUDE'] = lats
                df_erros['LONGITUDE'] = lons
                my_bar.empty()
                ainda_com_erro = df_erros['LATITUDE'].isna() | df_erros['LONGITUDE'].isna() | (df_erros['LATITUDE'] == 0.0)
                resgatadas = (~ainda_com_erro).sum()
                if resgatadas > 0:
                    status.update(label=f"✅ {resgatadas} coordenadas recuperadas!", state="complete", expanded=False)
                else:
                    status.update(label="Falha ao resgatar coordenadas.", state="error", expanded=False)
                df_tarefas = pd.concat([df_ok, df_erros[~ainda_com_erro]])
    
    final_mask = df_tarefas['LATITUDE'].isna() | df_tarefas['LONGITUDE'].isna() | (df_tarefas['LATITUDE'] == 0.0) | (df_tarefas['LONGITUDE'] == 0.0)
    return df_tarefas[~final_mask]

def extrair_lon_lat_kml(kml_text):
    coords = []
    padrao_alvo = re.compile(r'prim[aá]ri|secund[aá]ri|trafo|transformador|_pri|_sec|\bpri\b|\bsec\b|\bmt\b|\bbt\b', re.IGNORECASE)
    target_styles = set()
    for style_id, style_content in re.findall(r'<Style\s+id="([^"]+)"(.*?</Style>)', kml_text, re.DOTALL | re.IGNORECASE):
        if padrao_alvo.search(style_id) or padrao_alvo.search(style_content):
            target_styles.add(style_id)
            
    for stylemap_id, stylemap_content in re.findall(r'<StyleMap\s+id="([^"]+)"(.*?</StyleMap>)', kml_text, re.DOTALL | re.IGNORECASE):
        if padrao_alvo.search(stylemap_id) or padrao_alvo.search(stylemap_content):
            target_styles.add(stylemap_id)

    placemarks = re.findall(r'<Placemark.*?</Placemark>', kml_text, re.DOTALL | re.IGNORECASE)
    for pm in placemarks:
        is_target = False
        if padrao_alvo.search(pm):
            is_target = True
        else:
            style_urls = re.findall(r'<styleUrl>#?(.*?)</styleUrl>', pm, re.IGNORECASE)
            for url in style_urls:
                if url in target_styles or padrao_alvo.search(url):
                    is_target = True
                    break
        if is_target:
            matches = re.findall(r'<coordinates>\s*(.*?)\s*</coordinates>', pm, re.DOTALL)
            for match in matches:
                points = match.strip().split()
                for pt in points:
                    parts = pt.split(',')
                    if len(parts) >= 2:
                        try:
                            lon = float(parts[0])
                            lat = float(parts[1])
                            if lat != 0.0 and lon != 0.0: coords.append((lon, lat))
                        except: pass
    return list(set(coords)) 

def extrair_coordenadas_rede(uploaded_files):
    coords_list = []
    for f in uploaded_files:
        file_bytes = f.getvalue()
        if f.name.lower().endswith('.kmz'):
            try:
                with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                    for zinfo in z.namelist():
                        if zinfo.lower().endswith('.kml'):
                            kml_data = z.read(zinfo).decode('utf-8', errors='ignore')
                            coords_list.extend(extrair_lon_lat_kml(kml_data))
            except Exception: pass
        elif f.name.lower().endswith('.kml'):
            kml_data = file_bytes.decode('utf-8', errors='ignore')
            coords_list.extend(extrair_lon_lat_kml(kml_data))
    
    if not coords_list: return pd.DataFrame()
    df_rede = pd.DataFrame(coords_list, columns=['LONGITUDE', 'LATITUDE'])
    df_rede = df_rede.dropna().drop_duplicates()
    return df_rede

def encontrar_rede_mais_proxima(df_tasks, df_rede, vao_medio):
    if df_rede.empty or df_tasks.empty: return df_tasks
    
    rede_lats = df_rede['LATITUDE'].values
    rede_lons = df_rede['LONGITUDE'].values
    nearest_lats, nearest_lons, nearest_dists, nearest_postes = [], [], [], []
    
    for _, row in df_tasks.iterrows():
        t_lat, t_lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.isna(t_lat) or pd.isna(t_lon):
            nearest_lats.append(np.nan); nearest_lons.append(np.nan); nearest_dists.append(np.nan); nearest_postes.append(np.nan)
            continue
            
        dists = haversine_vectorized(t_lat, t_lon, rede_lats, rede_lons)
        min_idx = np.argmin(dists)
        
        dist_metros = dists[min_idx] * 1000 
        postes = int(dist_metros // vao_medio) 
        
        nearest_dists.append(dist_metros)
        nearest_postes.append(postes)
        nearest_lats.append(rede_lats[min_idx])
        nearest_lons.append(rede_lons[min_idx])
        
    df_tasks['DISTANCIA_REDE_METROS'] = nearest_dists
    df_tasks['POSTES PREVISTOS'] = nearest_postes
    df_tasks['LATITUDE_REDE'] = nearest_lats
    df_tasks['LONGITUDE_REDE'] = nearest_lons
    return df_tasks

def fundir_super_pontos(df_tasks, raio_metros=5, agrupar_por_levantador=False):
    if df_tasks.empty or 'LATITUDE' not in df_tasks.columns or 'LONGITUDE' not in df_tasks.columns: return df_tasks, 0
    df_valid = df_tasks.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
    if df_valid.empty: return df_tasks, 0

    coords = np.radians(df_valid[['LATITUDE', 'LONGITUDE']].values)
    eps_rad = raio_metros / 6371000.0
    
    db = DBSCAN(eps=eps_rad, min_samples=1, algorithm='ball_tree', metric='haversine').fit(coords)
    df_valid['CLUSTER_ID'] = db.labels_
    
    if agrupar_por_levantador and 'LEVANTADOR' in df_valid.columns:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str) + "_" + df_valid['LEVANTADOR'].astype(str)
    else:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str)
    
    cluster_counts = df_valid['CLUSTER_GRP'].value_counts()
    single_clusters = cluster_counts[cluster_counts == 1].index
    multi_clusters = cluster_counts[cluster_counts > 1].index
    agrupado = []
    
    if len(multi_clusters) > 0:
        df_multi = df_valid[df_valid['CLUSTER_GRP'].isin(multi_clusters)]
        for c_id, group in df_multi.groupby('CLUSTER_GRP'):
            row_base = group.iloc[0].copy()
            qtd = len(group)
            orig_rows = group.to_dict('records')
            row_base['_ORIGINAL_ROWS'] = orig_rows
            
            def safe_list_join(col_name):
                itens = [str(x).strip() for x in group[col_name] if pd.notna(x) and str(x).lower() != 'nan']
                itens_unicos = list(dict.fromkeys(itens))
                return " | ".join(itens_unicos) if itens_unicos else "-"
            
            for col in group.columns:
                if col not in ['LATITUDE', 'LONGITUDE', 'CLUSTER_ID', 'CLUSTER_GRP', '_ORIGEM_BASE', 'PRIORIDADE']:
                    row_base[col] = safe_list_join(col)
                
            row_base['LATITUDE'] = group['LATITUDE'].mean()
            row_base['LONGITUDE'] = group['LONGITUDE'].mean()
            row_base['SUPER_PONTO'] = f"SIM ({qtd} un.)"
            if 'PRIORIDADE' in group.columns and 'Sim' in group['PRIORIDADE'].values:
                row_base['PRIORIDADE'] = 'Sim'
            agrupado.append(row_base)
            
    df_final_multi = pd.DataFrame(agrupado)
    df_single = df_valid[df_valid['CLUSTER_GRP'].isin(single_clusters)].copy()
    if not df_single.empty:
        df_single['SUPER_PONTO'] = "NÃO"
        dict_records = df_single.to_dict('records')
        df_single['_ORIGINAL_ROWS'] = [[r] for r in dict_records]
        
    df_final = pd.concat([df_single, df_final_multi], ignore_index=True).drop(columns=['CLUSTER_ID', 'CLUSTER_GRP'], errors='ignore')
    
    df_nan = df_tasks[df_tasks['LATITUDE'].isna() | df_tasks['LONGITUDE'].isna()].copy()
    if not df_nan.empty:
        df_nan['SUPER_PONTO'] = "NÃO"
        dict_records_nan = df_nan.to_dict('records')
        df_nan['_ORIGINAL_ROWS'] = [[r] for r in dict_records_nan]
        df_final = pd.concat([df_final, df_nan], ignore_index=True)
        
    return df_final, len(df_tasks) - len(df_final)
