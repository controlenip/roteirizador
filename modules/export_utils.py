import pandas as pd
import io
import html
import folium

# ==========================================
# GERAÇÃO DE EXCEL (DEMANDA GERAL E RESUMO)
# ==========================================

def gerar_excel_bytes(df, col_prio, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Obras Roteirizadas')
    return output.getvalue()

def gerar_excel_resumo_bytes(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
    return output.getvalue()

# ==========================================
# GERAÇÃO DE GPS OFFLINE (GPX) E KML
# ==========================================

def gerar_gpx_simples(df_kml, nome_rota):
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>', '<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">', f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>']
    for _, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.notna(lat) and pd.notna(lon):
            gpx.append(f'  <wpt lat="{lat}" lon="{lon}"><name>{html.escape(str(row.get("PROTOCOLO", "Ponto")))}</name></wpt>')
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append(f'  <trk><name>Traçado - {html.escape(str(nome_rota))}</name><trkseg>')
        for _, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom: gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg></trk>')
    gpx.append('</gpx>')
    return "\n".join(gpx)

def gerar_kml_fiscalizacao(df_kml, nome_rota, colunas_exibir, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '  <Document>', f'    <name>{html.escape(str(nome_rota))}</name>']
    styles = {
        'green': 'http://maps.google.com/mapfiles/kml/paddle/grn-blank.png',
        'blue': 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png',
        'beige': 'http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png', 
        'orange': 'http://maps.google.com/mapfiles/kml/paddle/orange-blank.png',
        'red': 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png',
        'gray': 'http://maps.google.com/mapfiles/kml/paddle/wht-blank.png'
    }
    for color, url in styles.items():
        kml.extend([f'    <Style id="style_{color}">', f'      <IconStyle><Icon><href>{url}</href></Icon></IconStyle>', '    </Style>'])

    coords_linha = []
    for _, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.isna(lat) or pd.isna(lon): continue
        coords_linha.append(f"{lon},{lat},0")
        
        nome, qtd, cor = str(row.get('PROTOCOLO', 'Ponto')), float(row.get('QTD PREVISTA DE POSTES', 0)), row.get('COR_ICONE', 'gray')
        desc = '<table border="1" style="border-collapse:collapse; width:100%;">' + "".join([f'<tr><td style="padding:3px;"><b>{html.escape(c)}</b></td><td style="padding:3px;">{funcao_formatadora(c, row.get(c, ""))}</td></tr>' for c in colunas_exibir]) + '</table>'

        kml.extend(['    <Placemark>', f'      <name>[{int(qtd)} Postes] {html.escape(nome)}</name>', f'      <styleUrl>#style_{cor}</styleUrl>', f'      <description><![CDATA[{desc}]]></description>', f'      <Point><coordinates>{lon},{lat},0</coordinates></Point>', '    </Placemark>'])

    if coords_linha:
        kml.extend(['    <Placemark>', '      <name>Traçado da Rota</name>', '      <Style><LineStyle><color>ff00ffff</color><width>3</width></LineStyle></Style>', '      <LineString><tessellate>1</tessellate><coordinates>', ' '.join(coords_linha), '      </coordinates></LineString>', '    </Placemark>'])

    kml.extend(['  </Document>', '</kml>'])
    return '\n'.join(kml)

def gerar_kml_agrupado(df_kml, bases_records, nome_arquivo, colunas_exibir, bases_ativas, tipo_periodo, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    kml.append('<Style id="s_blue"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon></IconStyle></Style>')
    kml.append('<Style id="s_red"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon></IconStyle></Style>')
    kml.append('<Style id="s_line"><LineStyle><color>ff0000ff</color><width>3</width></LineStyle></Style>')

    for b in bases_ativas:
        pasta = f'  <Folder><name>Base: {b}</name>'
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        for p in df_b['PERIODO'].unique():
            df_p = df_b[df_b['PERIODO'] == p]
            pasta += f'<Folder><name>Período {p}</name>'
            coords = []
            for _, r in df_p.iterrows():
                if isinstance(r.get('ROTA_GEOMETRIA'), list): coords.extend([f"{lon},{lat},0" for lon, lat in r['ROTA_GEOMETRIA']])
            if coords:
                pasta += '<Placemark><name>Rota</name><styleUrl>#s_line</styleUrl><LineString><coordinates>' + ' '.join(coords) + '</coordinates></LineString></Placemark>'
            for _, r in df_p.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                desc = '<table border="1">' + "".join([f'<tr><td>{c}</td><td>{funcao_formatadora(c, r.get(c,""))}</td></tr>' for c in colunas_exibir]) + '</table>'
                pasta += f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO")))}</name><styleUrl>#{"s_red" if r.get("PRIORIDADE")=="Sim" else "s_blue"}</styleUrl><description><![CDATA[{desc}]]></description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>'
            pasta += '</Folder>'
        pasta += '</Folder>'
        kml.append(pasta)
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)

# ==========================================
# UI E AUXILIARES
# ==========================================

def identificar_icone_folium(row, colunas_disponiveis):
    if 'TIPO NOTA' in colunas_disponiveis:
        t = str(row.get('TIPO NOTA', '')).upper()
        if t in ['UNR', 'ASC']: return 'bolt'
        elif t in ['MGD', 'MTP']: return 'industry'
        elif t == 'DIF': return 'exclamation-triangle'
    return 'map-marker'

def renderizar_painel_lateral(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Cota p/ Equipe:</b> <span style="color: #d9534f; font-weight: bold;">{limite_por_equipe}</span> obras</p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Capacidade Total:</b> <span style="color: #55B929; font-weight: bold;">{total_capacidade}</span> obras</p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras_prontas}</span></p>
    </div>
    """
