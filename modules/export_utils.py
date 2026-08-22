import pandas as pd
import io
import html
import folium
import os
import base64
import streamlit as st

# ==========================================
# GERAÇÃO DE EXCEL (RESPEITANDO A FORMATAÇÃO ORIGINAL)
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

def limpar_colunas_excel(df_alvo, cols_originais):
    """Nova lógica que respeita 100% as colunas originais do arquivo subido"""
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    rename_map = {}
    if 'PROTOCOLO' in df_alvo.columns and 'PROTOCOLO' not in cols_originais:
        rename_map['PROTOCOLO'] = 'NOTA'
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        rename_map['BASE_ATRIBUIDA'] = 'FISCAL'
    df_alvo = df_alvo.rename(columns=rename_map)
    
    # Priorizamos as colunas do Roteirizador, depois as Originais exatas
    final_cols = ['FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM']
    for c in cols_originais:
        if c in df_alvo.columns and c not in final_cols:
            final_cols.append(c)
        elif c == 'NOTA' and 'PROTOCOLO' in df_alvo.columns and 'PROTOCOLO' not in final_cols:
            final_cols.append('PROTOCOLO')
            
    # Adiciona colunas que sobraram e não são lixo de programação
    for c in df_alvo.columns:
        if c not in final_cols and not str(c).startswith('_') and c not in ['LINK_NAVEGACAO_OFFLINE', 'ROTA_GEOMETRIA', 'COORD_KEY', 'MUN_LIMPO', 'COR_ICONE', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM']:
            final_cols.append(c)
            
    return df_alvo[[c for c in final_cols if c in df_alvo.columns]]

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

def gerar_kml_fiscalizacao_agrupado(df_kml, nome_arquivo, colunas_exibir, bases_ativas, funcao_formatadora):
    """Gera KML idêntico ao Tático (Pastas por Fiscal) mas com Pop-ups coloridos de Fiscalização."""
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    styles = {
        'green': 'http://maps.google.com/mapfiles/kml/paddle/grn-blank.png',
        'blue': 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png',
        'beige': 'http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png', 
        'orange': 'http://maps.google.com/mapfiles/kml/paddle/orange-blank.png',
        'red': 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png',
        'gray': 'http://maps.google.com/mapfiles/kml/paddle/wht-blank.png'
    }
    for color, url in styles.items():
        kml.extend([f'<Style id="style_{color}"><IconStyle><Icon><href>{url}</href></Icon></IconStyle></Style>'])
        
    kml.append('<Style id="s_line"><LineStyle><color>ff0000ff</color><width>4</width></LineStyle></Style>')

    for b in bases_ativas:
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        pasta = [f'<Folder><name>Fiscal: {html.escape(str(b))}</name>']
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        
        coords_linha = []
        for _, r in df_b.iterrows():
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
            pasta.append('<Placemark><name>Traçado da Rota</name><styleUrl>#s_line</styleUrl><LineString><tessellate>1</tessellate><coordinates>' + ' '.join(coords_linha) + '</coordinates></LineString></Placemark>')

        for _, r in df_b.iterrows():
            if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
            lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
            if pd.isna(lat) or pd.isna(lon): continue
            
            qtd = float(r.get('QTD PREVISTA DE POSTES', 0))
            cor = r.get('COR_ICONE', 'gray')
            nome = str(r.get('PROTOCOLO', 'Ponto'))
            
            bg_colors = {'green': '#4CAF50', 'blue': '#2196F3', 'beige': '#FFC107', 'orange': '#FF9800', 'red': '#F44336', 'gray': '#9E9E9E'}
            txt_colors = {'beige': '#000000', 'orange': '#000000', 'green': '#ffffff', 'blue': '#ffffff', 'red': '#ffffff', 'gray': '#ffffff'}
            p_bg = bg_colors.get(cor, '#9E9E9E')
            p_c = txt_colors.get(cor, '#ffffff')
            p_txt = f"📋 FISCALIZAÇÃO - {int(qtd)} POSTES"

            er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}:</b></td><td style='padding:3px;'>{funcao_formatadora(c, r.get(c, ''))}</td></tr>" for c in colunas_exibir if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA','COR_ICONE']])
            desc = f'<div style="width:280px;"><div style="background:{p_bg};color:{p_c};padding:8px;font-weight:bold;">{p_txt}</div><table border="1" style="width:100%;font-size:12px;"><tr><td style="padding:3px;"><b>Ordem:</b></td><td style="padding:3px;">{r.get("ORDEM",0)}</td></tr>{er}</table></div>'

            pasta.append(f'<Placemark><name>[{int(qtd)} Postes] {html.escape(nome)}</name><styleUrl>#style_{cor}</styleUrl><description><![CDATA[{desc}]]></description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')
            
        pasta.append('</Folder>')
        kml.extend(pasta)
        
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)

# ==========================================
# UI E AUXILIARES
# ==========================================
def injetar_logo():
    import os
    if os.path.exists("LOGO_NIP.png"):
        st.logo("LOGO_NIP.png", icon_image=None)

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
