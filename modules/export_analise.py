import pandas as pd
import io
import html
import os
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

def gerar_excel_analise(dict_dfs):
    """Gera um arquivo Excel com múltiplas abas baseadas no dicionário enviado"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dict_dfs.items():
            if df.empty: continue
            
            # Remove colunas auxiliares do sistema
            df_saida = df.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME', 'CLUSTER_ID'], errors='ignore')
            
            # Limita o nome da aba a 31 caracteres (Regra do Excel)
            safe_sheet_name = sheet_name[:31]
            df_saida.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            
            worksheet = writer.sheets[safe_sheet_name]
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
                
    return output.getvalue()

def gerar_kml_analise(df):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', '<name>Análise Cruzada NIP</name>']
    
    styles = {
        'red': 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png',
        'orange': 'http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png',
        'green': 'http://maps.google.com/mapfiles/kml/paddle/grn-blank.png',
        'purple': 'http://maps.google.com/mapfiles/kml/paddle/purple-blank.png',
        'blue': 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png'
    }
    
    for color, url in styles.items():
        kml.append(f'<Style id="style_{color}"><IconStyle><Icon><href>{url}</href></Icon></IconStyle></Style>')

    # Estilo Preto Customizado
    kml.append('''<Style id="style_black"><IconStyle><color>ff000000</color><Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon></IconStyle></Style>''')

    # Agrupa por CLUSTER_ID para montar os agrupamentos geográficos (pastas)
    clusters = []
    if 'CLUSTER_ID' not in df.columns:
        df['CLUSTER_ID'] = range(len(df))
        
    for cid, grp in df.groupby('CLUSTER_ID'):
        lat = str(grp['LATITUDE'].iloc[0]).strip()
        lon = str(grp['LONGITUDE'].iloc[0]).strip()
        if not lat or not lon or lat.lower() == 'nan': continue
        
        c_names = grp['COR_NOME'].tolist()
        
        # Define a pasta dominante do terreno
        if any('Preto' in c for c in c_names) or any('Inválidas' in c for c in c_names): c_nome = next(c for c in c_names if 'Inválidas' in c or 'Preto' in c)
        elif any('Vermelho' in c for c in c_names) or any('Duplicadas' in c for c in c_names): c_nome = next(c for c in c_names if 'Duplicadas' in c or 'Vermelho' in c)
        elif len(grp) > 1: c_nome = '🟠 Notas Próximas'
        else: c_nome = c_names[0]

        if 'Inválidas' in c_nome or 'Preto' in c_nome: c_mapa = 'black'
        elif 'Duplicadas' in c_nome or 'Vermelho' in c_nome: c_mapa = 'red'
        elif 'Próximas' in c_nome or 'Laranja' in c_nome: c_mapa = 'orange'
        else: c_mapa = grp['COR_MAPA'].iloc[0]

        titulo_card = f"📍 Obras no Local ({len(grp)})"
        
        desc = f'''<![CDATA[
        <div style="font-family:sans-serif; width:300px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
            <div style="background:#0D256C; color:#ffffff; padding:8px 10px; font-size:13px; font-weight:bold; text-align:center;">{titulo_card}</div>
            <div style="padding:10px; background:#fafafa; font-size:12px;">
        '''
        
        nomes_notas = []
        for _, r in grp.iterrows():
            n = html.escape(str(r.get('NOTA', '')))
            nomes_notas.append(n)
            mun = html.escape(str(r.get('MUNICIPIO', '')))
            o = html.escape(str(r.get('ORIGEM_BASE', '')))
            s = html.escape(str(r.get('SITUACAO SAP', '')))
            col = html.escape(str(r.get('COLABORADORES MAIS PROXIMOS', '')))
            dup = html.escape(str(r.get('DUPLICADA', '')))
            
            desc += f'''
            <table style="width:100%; border-collapse:collapse; margin-bottom:8px;">
                <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                <tr><td style="padding:2px;"><b>Município:</b></td><td style="padding:2px;">{mun}</td></tr>
                <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                <tr><td style="padding:2px;"><b>Status SAP:</b></td><td style="padding:2px;">{s}</td></tr>
                <tr><td style="padding:2px;"><b>Equipes Perto:</b></td><td style="padding:2px;">{col}</td></tr>
                <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup}</td></tr>
            </table>
            <hr style="margin:4px 0; border:0; border-top:1px solid #ddd;">
            '''
        
        desc += '</div></div>]]>'
        
        clusters.append({
            'COR_NOME': c_nome,
            'COR_MAPA': c_mapa,
            'LAT': lat,
            'LON': lon,
            'DESC': desc,
            'NOTAS': " | ".join(nomes_notas)
        })

    df_clusters = pd.DataFrame(clusters)
    
    if not df_clusters.empty:
        for nome_grupo, df_grupo in df_clusters.groupby('COR_NOME'):
            kml.append(f'<Folder><name>{html.escape(str(nome_grupo))}</name>')
            for _, r in df_grupo.iterrows():
                kml.append(f'<Placemark><name>{r["NOTAS"]}</name><styleUrl>#style_{r["COR_MAPA"]}</styleUrl><description>{r["DESC"]}</description><Point><coordinates>{r["LON"]},{r["LAT"]},0</coordinates></Point></Placemark>')
            kml.append('</Folder>')

    kml.append('</Document></kml>')
    return "\n".join(kml)
