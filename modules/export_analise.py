import pandas as pd
import io
import html
import os
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

def gerar_excel_analise(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Analise Cruzada')
        worksheet = writer.sheets['Analise Cruzada']
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
        
        if duplicada == 'SIM': color = 'red'
        elif proxima == 'SIM': color = 'orange'
        elif origem == 'LEVANTAMENTO': color = 'green'
        elif origem == 'SANEAMENTO': color = 'purple'
        else: color = 'blue'

        desc = f'''<![CDATA[
        <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
            <div style="background:#0D256C; color:#ffffff; padding:8px 10px; font-size:13px; font-weight:bold;">📍 Análise de Ponto</div>
            <div style="padding:10px; background:#fafafa; font-size:12px;">
                <table style="width:100%; border-collapse:collapse;">
                    <tr><td style="padding:3px;"><b>Nota:</b></td><td style="padding:3px;">{html.escape(nota)}</td></tr>
                    <tr><td style="padding:3px;"><b>Município:</b></td><td style="padding:3px;">{html.escape(mun)}</td></tr>
                    <tr><td style="padding:3px;"><b>Origem:</b></td><td style="padding:3px;">{html.escape(origem)}</td></tr>
                    <tr><td style="padding:3px;"><b>Status SAP:</b></td><td style="padding:3px;">{html.escape(status_sap)}</td></tr>
                    <tr><td style="padding:3px;"><b>Duplicada:</b></td><td style="padding:3px;">{html.escape(duplicada)}</td></tr>
                    <tr><td style="padding:3px;"><b>Próxima a outra:</b></td><td style="padding:3px;">{html.escape(proxima)}</td></tr>
                    <tr><td style="padding:3px;"><b>Colab Mais Perto:</b></td><td style="padding:3px;">{html.escape(colab)}</td></tr>
                </table>
            </div>
        </div>
        ]]>'''
        
        desc = desc.replace("{", "&#123;").replace("}", "&#125;")
        kml.append(f'<Placemark><name>{html.escape(nota)}</name><styleUrl>#style_{color}</styleUrl><description>{desc}</description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')

    kml.append('</Document></kml>')
    return "\n".join(kml)
