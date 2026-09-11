import pandas as pd
import io
import html
import os
import re
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

def formatar_planilha_analise(writer, sheet_name):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri')
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    worksheet.auto_filter.ref = worksheet.dimensions
    
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        worksheet.column_dimensions[column].width = min(max_length + 2, 60)

def gerar_excel_analise(df, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Obras Roteirizadas')
        formatar_planilha_analise(writer, 'Obras Roteirizadas')
    return output.getvalue()

def gerar_excel_resumo_analise(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
        formatar_planilha_analise(writer, 'Resumo Operacional')
    return output.getvalue()

def limpar_colunas_analise(df_alvo, cols_originais):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    # Assegura colunas mínimas de roteamento na frente
    final_cols = ['BASE_ATRIBUIDA', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM']
    
    if cols_originais is not None:
        for c in cols_originais:
            if c in df_alvo.columns and c not in final_cols:
                final_cols.append(c)
                
    colunas_lixo = ['LINK_NAVEGACAO_OFFLINE', 'ROTA_GEOMETRIA', 'COORD_KEY', 'MUN_LIMPO', 'COR_ICONE', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'CLUSTER_ID', 'CLUSTER_GRP', 'MLC']
    for c in df_alvo.columns:
        if c not in final_cols and not str(c).startswith('_') and c not in colunas_lixo: 
            final_cols.append(c)
            
    return df_alvo[[c for c in final_cols if c in df_alvo.columns]]

def gerar_txt_analise(df, colunas_originais=None):
    """Gera um arquivo TXT de texto puro de forma genérica para as notas analisadas."""
    linhas_txt = []
    
    # Tenta descobrir colunas lógicas
    col_nota = next((c for c in df.columns if c.upper() in ['NOTA', 'PROTOCOLO', 'OS', 'ID']), 'N/A')
    
    for _, r in df.iterrows():
        nota = str(r.get(col_nota, '-')).strip()
        lat = str(r.get('LATITUDE', '')).strip()
        lon = str(r.get('LONGITUDE', '')).strip()

        bloco = []
        bloco.append(f"IDENTIFICADOR: {nota}")
        
        # Puxa até 10 colunas relevantes para o TXT genérico
        for col in df.columns:
            if col not in [col_nota, 'LATITUDE', 'LONGITUDE', 'ROTA_GEOMETRIA', 'BASE_ATRIBUIDA', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM'] and not col.startswith('_'):
                val = str(r.get(col, '')).strip()
                if val.lower() not in ['nan', 'none', '']:
                    bloco.append(f"{col}: {val}")
        
        if lat and lon and lat.lower() != 'nan' and lon.lower() != 'nan':
            bloco.append(f"https://www.google.com.br/maps/place/{lat},{lon}")
        
        linhas_txt.append("\n".join(bloco))
        linhas_txt.append("\n-------------------------------------------------------------------------------------------------------------------------\n")
        
    return "".join(linhas_txt)

def gerar_kml_analise(df_kml, nome_arquivo, colunas_exibir, bases_ativas, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    kml.append('<Style id="linha-rota-contorno"><LineStyle><color>ff000000</color><width>8</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')
    kml.append('<Style id="style-blue-n"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')

    cores_kml = ['ff4b19e6', 'ffd4bc00', 'ffb5513f', 'ff889600', 'ff0098ff', 'ffb0279c', 'ff39dccd', 'ff148000', 'ffeb004b', 'ff1f618d', 'ffd35400', 'ff16a085', 'ff8e44ad', 'ff27ae60', 'ffe67e22']

    for idx, b in enumerate(bases_ativas):
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        b_safe = re.sub(r'[^A-Za-z0-9]', '', str(b))
        style_id = f"rota-centro-{b_safe}"
        cor = cores_kml[idx % len(cores_kml)]
        kml.append(f'<Style id="{style_id}"><LineStyle><color>{cor}</color><width>5</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')

    for b in bases_ativas:
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        pasta = [f'<Folder><name>Equipe: {html.escape(str(b))}</name>']
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        b_safe = re.sub(r'[^A-Za-z0-9]', '', str(b))
        
        for p in df_b['PERIODO'].unique():
            df_p = df_b[df_b['PERIODO'] == p]
            pasta.append(f'<Folder><name>Rota Principal</name>')
            
            coords_linha = []
            for _, r in df_p.iterrows():
                geom = r.get('ROTA_GEOMETRIA')
                if isinstance(geom, list) and len(geom) > 0:
                    for pt in geom:
                        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                            coords_linha.append(f"{pt[0]},{pt[1]},0")
                else:
                    lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                    if pd.notna(lat) and pd.notna(lon): coords_linha.append(f"{lon},{lat},0")
            
            if coords_linha:
                str_coords = '\n'.join(coords_linha)
                pasta.append('<Placemark><name>Contorno Rota</name><styleUrl>#linha-rota-contorno</styleUrl><LineString><tessellate>1</tessellate><coordinates>\n' + str_coords + '\n</coordinates></LineString></Placemark>')
                pasta.append(f'<Placemark><name>Traçado Rota</name><styleUrl>#rota-centro-{b_safe}</styleUrl><LineString><tessellate>1</tessellate><coordinates>\n' + str_coords + '\n</coordinates></LineString></Placemark>')

            for _, r in df_p.iterrows():
                if str(r.get('BASE_ATRIBUIDA', '')) in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                
                nome_ponto = f"[{r.get('ORDEM', '')}] Ponto Roteado"
                
                er = "".join([f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555;'>{html.escape(c)}:</td><td style='padding:3px 6px; color:#333;'>{funcao_formatadora(c, r.get(c, ''))}</td></tr>" for c in colunas_exibir if c not in ['BASE_ATRIBUIDA', 'ROTA_GEOMETRIA']])
                desc = f'''<![CDATA[
                <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                    <div style="background:#0D256C; color:#ffffff; padding:8px 10px; font-size:13px; font-weight:bold;">📍 Análise de Ponto</div>
                    <div style="padding:10px; background:#fafafa; font-size:12px;">
                        <table style="width:100%; border-collapse:collapse;">
                            <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Ordem:</td><td style="padding:3px 6px; color:#333;">{r.get('ORDEM', '')}</td></tr>
                            {er}
                        </table>
                    </div>
                </div>]]>'''
                desc = desc.replace("{", "&#123;").replace("}", "&#125;")

                pasta.append(f'<Placemark><name>{html.escape(nome_ponto)}</name><styleUrl>#style-blue-n</styleUrl><description>{desc}</description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')
                
            pasta.append('</Folder>')
        pasta.append('</Folder>')
        kml.extend(pasta)
        
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)

def gerar_gpx_simples(df_kml, nome_rota):
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>', '<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">', f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>']
    for _, row in df_kml.iterrows():
        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.notna(lat) and pd.notna(lon):
            gpx.append(f'  <wpt lat="{lat}" lon="{lon}"><name>Ponto {row.get("ORDEM", "")}</name></wpt>')
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append(f'  <trk><name>Traçado - {html.escape(str(nome_rota))}</name><trkseg>')
        for _, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom: gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg></trk>')
    gpx.append('</gpx>')
    return "\n".join(gpx)

def injetar_logo():
    if os.path.exists("LOGO_NIP.png"): st.logo("LOGO_NIP.png", icon_image=None)
