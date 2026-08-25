import pandas as pd
import io
import html
import os
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime

# ==========================================
# EXPORTADORES DO SANEAMENTO
# ==========================================

def formatar_planilha_openpyxl(writer, sheet_name):
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
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        worksheet.column_dimensions[column].width = min(max_length + 2, 60)

def gerar_excel_saneamento(df, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    
    for col_name in ['PROTOCOLO', 'NOTA']:
        if col_name in df_saida.columns:
            df_saida = df_saida[~df_saida[col_name].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY', 'LAT_NUM', 'LON_NUM']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Saneamento Roteirizado')
        formatar_planilha_openpyxl(writer, 'Saneamento Roteirizado')
        
    return output.getvalue()

def gerar_excel_resumo_saneamento(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Saneamento')
        formatar_planilha_openpyxl(writer, 'Resumo Saneamento')
    return output.getvalue()

def limpar_colunas_saneamento(df_alvo, cols_originais):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    if 'PROTOCOLO' in df_alvo.columns:
        if 'NOTA' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['NOTA'])
        df_alvo = df_alvo.rename(columns={'PROTOCOLO': 'NOTA'})
        
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        if 'FISCAL' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['FISCAL'])
        df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'FISCAL'})
        
    # Colunas bases da roteirização
    final_cols = ['FISCAL', 'NOME_DIA', 'DIA_MES', 'SEMANA', 'DIA', 'PERIODO', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'HORA_INICIO', 'HORA_FIM']
    
    # As colunas Obrigatórias solicitadas pelo usuário (Mapeamento normalizado)
    req_cols = ['NOTA', 'STATUS CLIENTE', 'NOME', 'TIPO DEMANDA', 'MUNICIPIO', 'ENDERECO', 'BAIRRO', 'PONTO REFERENCIA', 'COMPLEMENTO', 'LATITUDE PROJETO', 'LONGITUDE PROJETO', 'CLASSIFICACAO AREA', 'TEL FIXO', 'TEL MOVEL', 'GRUPO TENSAO']
    
    # Adicionando os Requisitos
    for req in req_cols:
        if req in df_alvo.columns and req not in final_cols:
            final_cols.append(req)
            
    # Adicionando o que sobrou de forma orgânica
    if cols_originais is not None:
        for c in cols_originais:
            if c in df_alvo.columns and c not in final_cols and c not in ['LEVANTADOR', 'NOME DO LEVANTADOR', 'LEVANTADOR_RESPONSAVEL']:
                final_cols.append(c)
                
    colunas_lixo = ['LINK_NAVEGACAO_OFFLINE', 'ROTA_GEOMETRIA', 'COORD_KEY', 'MUN_LIMPO', 'COR_ICONE', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'CLUSTER_ID', 'CLUSTER_GRP', 'MLC', 'LAT_NUM', 'LON_NUM']
    for c in df_alvo.columns:
        if c not in final_cols and not str(c).startswith('_') and c not in colunas_lixo and c not in ['LATITUDE', 'LONGITUDE']:
            final_cols.append(c)
            
    return df_alvo[[c for c in final_cols if c in df_alvo.columns]]

def gerar_kml_saneamento(df_kml, nome_arquivo, colunas_exibir, bases_ativas, tipo_periodo, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    kml.append('<Style id="s_blue"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon></IconStyle></Style>')
    kml.append('<Style id="s_red"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon></IconStyle></Style>')
    
    colors = ['ff0000ff', 'ff00ff00', 'ffff0000', 'ff00ffff', 'ffffff00', 'ffff00ff']
    
    for b in bases_ativas:
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        pasta = [f'<Folder><name>Equipe: {html.escape(str(b))}</name>']
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        
        for p_idx, p in enumerate(df_b['PERIODO'].unique()):
            df_p = df_b[df_b['PERIODO'] == p]
            
            c_line = colors[p_idx % len(colors)]
            kml.append(f'<Style id="s_line_{p_idx}"><LineStyle><color>{c_line}</color><width>4</width></LineStyle></Style>')
            
            nome_pasta_per = f"{tipo_periodo} {p}" if tipo_periodo == "Dia" else f"Semana {p}"
            pasta.append(f'<Folder><name>{nome_pasta_per}</name>')
            
            coords_linha = []
            for _, r in df_p.iterrows():
                geom = r.get('ROTA_GEOMETRIA')
                if isinstance(geom, list) and len(geom) > 0:
                    for pt in geom:
                        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                            coords_linha.append(f"{pt[0]},{pt[1]},0")
                else:
                    lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                    if pd.notna(lat) and pd.notna(lon):
                        coords_linha.append(f"{lon},{lat},0")
            
            if coords_linha:
                pasta.append(f'<Placemark><name>Traçado da Rota</name><styleUrl>#s_line_{p_idx}</styleUrl><LineString><tessellate>1</tessellate><coordinates>' + ' '.join(coords_linha) + '</coordinates></LineString></Placemark>')

            for _, r in df_p.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                
                is_prio = r.get('PRIORIDADE') == 'Sim'
                cor = "s_red" if is_prio else "s_blue"
                
                desc = '<table border="1">' + "".join([f'<tr><td>{html.escape(c)}</td><td>{funcao_formatadora(c, r.get(c,""))}</td></tr>' for c in colunas_exibir]) + '</table>'
                pasta.append(f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO")))}</name><styleUrl>#{cor}</styleUrl><description><![CDATA[{desc}]]></description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')
                
            pasta.append('</Folder>')
        pasta.append('</Folder>')
        kml.extend(pasta)
        
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)

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

def injetar_logo():
    if os.path.exists("LOGO_NIP.png"): st.logo("LOGO_NIP.png", icon_image=None)

def identificar_icone_folium(row, colunas_disponiveis):
    if 'TIPO DEMANDA' in colunas_disponiveis:
        t = str(row.get('TIPO DEMANDA', '')).upper()
        if 'RELIGA' in t or 'LIGACAO' in t: return 'bolt'
    return 'map-marker'
