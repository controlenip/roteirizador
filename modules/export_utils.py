import pandas as pd
import html
import re

def gerar_excel_bytes(df, col_prio, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df_saida.to_excel(writer, index=False, sheet_name='Obras')
    return output.getvalue()

def gerar_excel_resumo_bytes(df_resumo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df_resumo.to_excel(writer, index=False, sheet_name='Resumo')
    return output.getvalue()

def gerar_gpx_simples(df_kml, nome_rota):
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>', '<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">', f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>']
    for _, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.notna(lat) and pd.notna(lon): gpx.append(f'  <wpt lat="{lat}" lon="{lon}"><name>{html.escape(str(row.get("PROTOCOLO", "Ponto")))}</name></wpt>')
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append(f'  <trk><name>Traçado - {html.escape(str(nome_rota))}</name><trkseg>')
        for _, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom: gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg></trk>')
    gpx.append('</gpx>')
    return "\n".join(gpx)

def gerar_kml_agrupado(df_kml, bases_records, nome_arquivo, colunas_exibir, bases_ativas, tipo_periodo, funcao_formatadora):
    """Gera KML puro sem precisar de biblioteca externa."""
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    # Styles
    kml.append('<Style id="s_blue"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon></IconStyle></Style>')
    kml.append('<Style id="s_red"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon></IconStyle></Style>')
    kml.append('<Style id="s_line"><LineStyle><color>ff0000ff</color><width>3</width></LineStyle></Style>')

    for b in bases_ativas:
        pasta = f'  <Folder><name>Base: {b}</name>'
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        for p in df_b['PERIODO'].unique():
            df_p = df_b[df_b['PERIODO'] == p]
            pasta += f'<Folder><name>Período {p}</name>'
            # Linha
            coords = []
            for _, r in df_p.iterrows():
                if isinstance(r.get('ROTA_GEOMETRIA'), list): coords.extend([f"{lon},{lat},0" for lon, lat in r['ROTA_GEOMETRIA']])
            pasta += '<Placemark><name>Rota</name><styleUrl>#s_line</styleUrl><LineString><coordinates>' + ' '.join(coords) + '</coordinates></LineString></Placemark>'
            # Pontos
            for _, r in df_p.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                desc = '<table border="1">' + "".join([f'<tr><td>{c}</td><td>{funcao_formatadora(c, r.get(c,""))}</td></tr>' for c in colunas_exibir]) + '</table>'
                pasta += f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO")))}</name><styleUrl>#{"s_red" if r.get("PRIORIDADE")=="Sim" else "s_blue"}</styleUrl><description><![CDATA[{desc}]]></description><Point><coordinates>{r.get("LONGITUDE")},{r.get("LATITUDE")},0</coordinates></Point></Placemark>'
            pasta += '</Folder>'
        pasta += '</Folder>'
        kml.append(pasta)
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)

def renderizar_painel_lateral(limite, total, equipes, capacidade):
    return f"""
    <div style="background-color: #ffffff; padding: 25px; border-radius: 10px; border: 1px solid #e0e0e0; box-shadow: 0 4px 8px rgba(0,0,0,0.05); margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 18px; border-bottom: 2px solid #55B929; padding-bottom: 10px; margin-bottom: 15px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{equipes}</span></p>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Cota p/ Equipe:</b> <span style="color: #d9534f; font-weight: bold;">{limite}</span> obras</p>
        <p style="margin-bottom: 15px; font-size: 15px;"><b>Capacidade Total:</b> <span style="color: #55B929; font-weight: bold;">{capacidade}</span> obras</p>
        <hr style="margin: 15px 0; border: 0; border-top: 1px solid #eee;">
        <div style="text-align: center; margin-top: 15px;">
            <p style="margin-bottom: 5px; font-size: 16px; color: #555;"><b>Obras Validadas:</b></p>
            <span style="font-size: 36px; color: #0D256C; font-weight: 900;">{total}</span>
        </div>
    </div>
    """
