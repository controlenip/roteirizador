import pandas as pd
import io
import html
import os
import re
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# EXPORTADORES DA LISTA CONTÍNUA
# ==========================================

def formatar_planilha_openpyxl(writer, sheet_name):
    """Aplica o padrão NIP (Azul Escuro) no arquivo Excel gerado."""
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

def gerar_excel_lista(df, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    
    for col_name in ['PROTOCOLO', 'NOTA']:
        if col_name in df_saida.columns:
            df_saida = df_saida[~df_saida[col_name].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Obras Roteirizadas')
        formatar_planilha_openpyxl(writer, 'Obras Roteirizadas')
        
    return output.getvalue()

def gerar_excel_resumo_lista(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
        formatar_planilha_openpyxl(writer, 'Resumo Operacional')
    return output.getvalue()

def limpar_colunas_lista(df_alvo, cols_originais=None):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    if 'PROTOCOLO' in df_alvo.columns:
        if 'NOTA' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['NOTA'])
        df_alvo = df_alvo.rename(columns={'PROTOCOLO': 'NOTA'})
        
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        if 'FISCAL' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['FISCAL'])
        df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'FISCAL'})
        
    # Colunas obrigatórias para roteirização e identificação
    final_cols = ['FISCAL', 'DIA_SEMANA', 'DIA_MES', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM']
    if 'NOTA' in df_alvo.columns: final_cols.append('NOTA')
    
    if cols_originais is not None:
        for c in cols_originais:
            nome_coluna = 'NOTA' if c == 'PROTOCOLO' else c
            nome_coluna = 'FISCAL' if c == 'BASE_ATRIBUIDA' else nome_coluna
            
            if nome_coluna in df_alvo.columns and nome_coluna not in final_cols and nome_coluna not in ['LEVANTADOR', 'NOME DO LEVANTADOR', 'LEVANTADOR_RESPONSAVEL']:
                final_cols.append(nome_coluna)
    else:
        colunas_lixo = ['LINK_NAVEGACAO_OFFLINE', 'ROTA_GEOMETRIA', 'COORD_KEY', 'MUN_LIMPO', 'COR_ICONE', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'CLUSTER_ID', 'CLUSTER_GRP', 'MLC', 'DISTANCIA_PROXIMO_PONTO_KM']
        for c in df_alvo.columns:
            if c not in final_cols and not str(c).startswith('_') and c not in colunas_lixo:
                final_cols.append(c)
                
    return df_alvo[[c for c in final_cols if c in df_alvo.columns]]

def gerar_kml_lista(df_kml, nome_arquivo, colunas_exibir, bases_ativas, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    # Injetando Estilos de Ícones Padrões e Contorno (Baseado nas especificações)
    kml.append('<Style id="linha-rota-contorno"><LineStyle><color>ff000000</color><width>8</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')
    kml.append('<Style id="linha-ligacao-rede"><LineStyle><color>8800ffff</color><width>2</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')

    kml.append('''<Style id="style-blue-n"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-blue-h"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-blue"><Pair><key>normal</key><styleUrl>#style-blue-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-blue-h</styleUrl></Pair></StyleMap>''')

    kml.append('''<Style id="style-red-n"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-red-h"><IconStyle><scale>1.5</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-red"><Pair><key>normal</key><styleUrl>#style-red-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-red-h</styleUrl></Pair></StyleMap>''')

    kml.append('''<Style id="style-yellow-n"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png</href></Icon></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-yellow-h"><IconStyle><scale>1.5</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png</href></Icon></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-yellow"><Pair><key>normal</key><styleUrl>#style-yellow-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-yellow-h</styleUrl></Pair></StyleMap>''')

    cores_kml = ['ff4b19e6', 'ffd4bc00', 'ffb5513f', 'ff889600', 'ff0098ff', 'ffb0279c', 'ff39dccd', 'ff148000', 'ffeb004b', 'ff1f618d', 'ffd35400', 'ff16a085', 'ff8e44ad', 'ff27ae60', 'ffe67e22']

    # Gerando os estilos das linhas de rota dinamicamente para cada levantador
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
            
            # --- DESENHO DO ARRUAMENTO / TRAÇADO ---
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

            # --- GERAÇÃO DOS MARCADORES E CARDS (POP-UP) ---
            for _, r in df_p.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
                is_prio = str(r.get('PRIORIDADE')).upper() == 'SIM'
                
                # Identificando a Cor / Formato pelo Status da Obra
                if is_sp:
                    bg_color, text_color = "#FFD700", "#000000"
                    qty = len(r.get('_ORIGINAL_ROWS', []))
                    header_txt = f"🏢 SUPER PONTO ({qty} un.)"
                    icon = "icon-yellow"
                    nome_ponto = f"[{r.get('ORDEM', '')}] 🏢 SUPER PONTO ({qty} un.)"
                elif is_prio:
                    bg_color, text_color = "#d9534f", "#ffffff"
                    header_txt = "🚨 OBRA PRIORITÁRIA"
                    icon = "icon-red"
                    nome_ponto = f"[PRIORIDADE] [{r.get('ORDEM', '')}] Doc: {r.get('PROTOCOLO', '')}"
                else:
                    bg_color, text_color = "#0D256C", "#ffffff"
                    header_txt = "📍 Atendimento Padrão"
                    icon = "icon-blue"
                    nome_ponto = f"[{r.get('ORDEM', '')}] Doc: {r.get('PROTOCOLO', '')}"

                # Cabeçalho HTML do Pop-Up
                desc = f'''<![CDATA[
                <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                    <div style="background:{bg_color}; color:{text_color}; padding:8px 10px; font-size:13px; font-weight:bold;">{header_txt}</div>
                    <div style="padding:10px; background:#fafafa; font-size:12px;">
                        <table style="width:100%; border-collapse:collapse;">
                '''
                
                ordem_txt = f"{r.get('ORDEM', '')}"
                if r.get('DIA_SEMANA'):
                    dia_sm = str(r.get('DIA_SEMANA')).title()
                    ordem_txt += f" ({dia_sm})"

                dist_ant = f"{r.get('DISTANCIA_PONTO_ANTERIOR_KM', 0.0)} KM"
                dist_prox = f"{r.get('DISTANCIA_PROXIMO_PONTO_KM', 0.0)} KM"

                # Bloco do Protocolo
                if is_sp:
                    prot_list = [orig.get('PROTOCOLO', orig.get('NOTA', '')) for orig in r['_ORIGINAL_ROWS']]
                    prot_html = "<div style='max-height:95px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>" + "".join([f"<li style='margin-bottom:3px;'><b>[{i+1}]</b> {html.escape(str(p))}</li>" for i, p in enumerate(prot_list)]) + "</ul></div>"
                else:
                    prot_html = html.escape(str(r.get('PROTOCOLO', r.get('NOTA', ''))))

                desc += f'''
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;">Nota/Protocolo:</td><td style="padding:3px 6px; color:#333;">{prot_html}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Ordem:</td><td style="padding:3px 6px; color:#333;">{html.escape(str(ordem_txt))}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Ant.:</td><td style="padding:3px 6px; color:#333;">{dist_ant}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Próx.:</td><td style="padding:3px 6px; color:#333;">{dist_prox}</td></tr>
                '''

                # Iterando as Colunas Visíveis Solicitadas (Apenas o que o usuário marcou nos botões azuis)
                for c in colunas_exibir:
                    if c.upper() in ['PROTOCOLO', 'NOTA', 'DIA_SEMANA', 'DIA_MES']: continue
                    
                    if is_sp and c.upper() not in ['LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'ZONA', 'REGIONAL']:
                        vals = [orig.get(c, '') for orig in r['_ORIGINAL_ROWS']]
                        val_html = "<div style='max-height:95px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>" + "".join([f"<li style='margin-bottom:3px;'><b>[{i+1}]</b> {funcao_formatadora(c, v)}</li>" for i, v in enumerate(vals)]) + "</ul></div>"
                    else:
                        val_html = funcao_formatadora(c, r.get(c, ''))
                        
                    desc += f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;'>{html.escape(c)}:</td><td style='padding:3px 6px; color:#333;'>{val_html}</td></tr>"

                desc += "</table></div></div>]]>"

                pasta.append(f'<Placemark><name>{html.escape(nome_ponto)}</name><styleUrl>#{icon}</styleUrl><description>{desc}</description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')
                
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
    if 'TIPO NOTA' in colunas_disponiveis:
        t = str(row.get('TIPO NOTA', '')).upper()
        if t in ['UNR', 'ASC']: return 'bolt'
        elif t in ['MGD', 'MTP']: return 'industry'
        elif t == 'DIF': return 'exclamation-triangle'
    return 'map-marker'
