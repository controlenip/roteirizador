def extrair_coordenadas_rede(rede_files):
    import zipfile
    import re
    import pandas as pd
    
    pontos_rede = []
    for f in rede_files:
        # Extrai o nome do alimentador a partir do nome do arquivo KMZ
        nome_alimentador = f.name.replace('.kmz', '').replace('.kml', '').replace('.KMZ', '').replace('.KML', '')
        
        kml_content = ""
        if f.name.lower().endswith('.kmz'):
            try:
                with zipfile.ZipFile(f, 'r') as kmz:
                    for item in kmz.namelist():
                        if item.lower().endswith('.kml'):
                            kml_content = kmz.read(item).decode('utf-8', errors='ignore')
                            break
            except Exception as e:
                continue
        else:
            kml_content = f.read().decode('utf-8', errors='ignore')
            
        coords_matches = re.findall(r'<coordinates>(.*?)</coordinates>', kml_content, re.DOTALL)
        for match in coords_matches:
            coords_str = match.strip().split()
            for c in coords_str:
                parts = c.split(',')
                if len(parts) >= 2:
                    try:
                        lon = float(parts[0].strip())
                        lat = float(parts[1].strip())
                        pontos_rede.append({
                            'LATITUDE_REDE': lat,
                            'LONGITUDE_REDE': lon,
                            'ALIMENTADOR': nome_alimentador # NOVO: Guarda o nome do alimentador
                        })
                    except:
                        pass
                        
    return pd.DataFrame(pontos_rede)

def encontrar_rede_mais_proxima(df_tasks, df_rede, vao_medio):
    import numpy as np
    import pandas as pd
    
    if df_rede.empty or df_tasks.empty:
        return df_tasks
        
    lat_rede = df_rede['LATITUDE_REDE'].values
    lon_rede = df_rede['LONGITUDE_REDE'].values
    nomes_rede = df_rede['ALIMENTADOR'].values # NOVO: Puxa a lista de nomes dos alimentadores
    
    distancias_rede = []
    postes_previstos = []
    lats_prox = []
    lons_prox = []
    alimentadores_prox = [] # NOVO: Lista que vai guardar o alimentador de cada obra
    
    for _, row in df_tasks.iterrows():
        lat_obra = row.get('LATITUDE')
        lon_obra = row.get('LONGITUDE')
        
        if pd.isna(lat_obra) or pd.isna(lon_obra):
            distancias_rede.append(np.nan)
            postes_previstos.append(np.nan)
            lats_prox.append(np.nan)
            lons_prox.append(np.nan)
            alimentadores_prox.append(np.nan)
            continue
            
        # Cálculo vetorizado ultra-rápido de Haversine
        lat1 = np.radians(lat_obra)
        lon1 = np.radians(lon_obra)
        lat2 = np.radians(lat_rede)
        lon2 = np.radians(lon_rede)
        
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        
        a = np.sin(dlat/2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2.0)**2
        c = 2 * np.arcsin(np.sqrt(a))
        dist_m = 6371000 * c
        
        idx_min = np.argmin(dist_m)
        min_dist = dist_m[idx_min]
        
        distancias_rede.append(min_dist)
        postes_previstos.append(np.ceil(min_dist / vao_medio) + 1)
        lats_prox.append(df_rede.iloc[idx_min]['LATITUDE_REDE'])
        lons_prox.append(df_rede.iloc[idx_min]['LONGITUDE_REDE'])
        alimentadores_prox.append(df_rede.iloc[idx_min]['ALIMENTADOR']) # NOVO: Salva o Alimentador vencedor
        
    df_tasks['DISTANCIA_REDE_METROS'] = distancias_rede
    df_tasks['POSTES PREVISTOS'] = postes_previstos
    df_tasks['LATITUDE_REDE'] = lats_prox
    df_tasks['LONGITUDE_REDE'] = lons_prox
    df_tasks['ALIMENTADOR_PROXIMO'] = alimentadores_prox # NOVO: Injeta a coluna final no dataframe
    
    return df_tasks
