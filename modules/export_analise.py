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
            df_saida = df.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME'], errors='ignore')
            
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

    # Estilo Preto Customizado (pois o Google não tem paddle preto por padrão)
    kml.append('''<Style id="style_black"><IconStyle><color>ff000000</color><Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon></IconStyle></Style>''')

    for _, r in df.iterrows():
        lat = str(r.get('LATITUDE', '')).strip()
        lon = str(r.get('LONGITUDE', '')).strip()
        if not lat or not lon or lat.lower() == 'nan': continue
        
        nota = str(r.get('NOTA', ''))
        origem = str(r.get('ORIGEM_BASE', ''))
        duplicada = str(r.get('DUPLICADA', ''))
        proxima = str(r.get('PROXIMA', ''))
        status_sap = str(r.get('SITUACAO SAP', ''))
        colab = str(r.get('COLABORADOR MAIS PROXIMO', ''))
        mun = str(r.get('MUNICIPIO', ''))
        
        # Puxa a cor exata gerada no filtro do Streamlit
        color = str(r.get('COR_MAPA', 'blue'))

        desc = f'''<![CDATA[
        <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
            <div style="background:#0D256C; color:#ffffff; padding:8px 10px; font-size:13px; font-weight:bold;">📍 Análise de Ponto</div>
            <div style="padding:10px; background:#fafafa; font-size:12px;">
                <table style="width:100%; border-collapse:collapse;">
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Nota/Protocolo:</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(nota)}</td></tr>
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Município:</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(mun)}</td></tr>
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Origem:</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(origem)}</td></tr>
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Status SAP:</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(status_sap)}</td></tr>
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Duplicada (Bases):</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(duplicada)}</td></tr>
                    <tr><td style="padding:3px; border-bottom:1px solid #ddd;"><b>Próxima a outra:</b></td><td style="padding:3px; border-bottom:1px solid #ddd;">{html.escape(proxima)}</td></tr>
                    <tr><td style="padding:3px;"><b>Colab Mais Perto:</b></td><td style="padding:3px;">{html.escape(colab)}</td></tr>
                </table>
            </div>
        </div>
        ]]>'''
        
        desc = desc.replace("{", "&#123;").replace("}", "&#125;")
        kml.append(f'<Placemark><name>{html.escape(nota)}</name><styleUrl>#style_{color}</styleUrl><description>{desc}</description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')

    kml.append('</Document></kml>')
    return "\n".join(kml)
