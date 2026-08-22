import pandas as pd
import io
import html
import folium
from folium import plugins

# ==========================================
# GERAÇÃO DE EXCEL (DEMANDA GERAL E RESUMO)
# ==========================================

def gerar_excel_bytes(df, col_prio, colunas_originais=None):
    """
    Exporta o dataframe roteirizado garantindo que não existam colunas duplicadas.
    """
    output = io.BytesIO()
    
    df_saida = df.copy()
    df_saida = df_saida.loc[:, ~df_saida.columns.duplicated()]

    # Tenta remover colunas de controle interno, se existirem
    colunas_remover = [
        'ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', 
        '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY'
    ]
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Obras Roteirizadas')
    
    return output.getvalue()

def gerar_excel_resumo_bytes(df_resumo):
    """
    Gera a planilha de Resumo Operacional dos Levantadores/Fiscais.
    """
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
    return output.getvalue()

# ==========================================
# GERAÇÃO DE GPS OFFLINE (GPX) E KML
# ==========================================

def gerar_gpx_simples(df_kml, nome_rota):
    """
    Gera arquivo GPX para navegação offline rural (OsmAnd, Wikiloc).
    """
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>']
    gpx.append('<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">')
    gpx.append(f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>')
    
    for idx, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat = row.get('LATITUDE')
        lon = row.get('LONGITUDE')
        nome = str(row.get('PROTOCOLO', 'Ponto'))
        if pd.notna(lat) and pd.notna(lon):
            gpx.append(f'  <wpt lat="{lat}" lon="{lon}">')
            gpx.append(f'    <name>{html.escape(nome)}</name>')
            gpx.append(f'  </wpt>')
            
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append('  <trk>')
        gpx.append(f'    <name>Traçado - {html.escape(str(nome_rota))}</name>')
        gpx.append('    <trkseg>')
        for idx, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom:
                    gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg>')
        gpx.append('  </trk>')
        
    gpx.append('</gpx>')
    return "\n".join(gpx)

def gerar_kml_fiscalizacao(df_kml, nome_rota, colunas_exibir, funcao_formatadora):
    """
    Módulo exclusivo para gerar KML de Fiscalização (com cores por volume de postes).
    """
    kml = ['<?xml version="1.0" encoding="UTF-8"?>']
    kml.append('<kml xmlns="http://www.opengis.net/kml/2.2">')
    kml.append('  <Document>')
    kml.append(f'    <name>{html.escape(str(nome_rota))}</name>')

    # Definição Estrita de Cores Customizadas
    styles = {
        'green': 'http://maps.google.com/mapfiles/kml/paddle/grn-blank.png',
        'blue': 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png',
        'beige': 'http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png', 
        'orange': 'http://maps.google.com/mapfiles/kml/paddle/orange-blank.png',
        'red': 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png',
        'gray': 'http://maps.google.com/mapfiles/kml/paddle/wht-blank.png'
    }
    for color, url in styles.items():
        kml.append(f'    <Style id="style_{color}">')
        kml.append(f'      <IconStyle><Icon><href>{url}</href></Icon></IconStyle>')
        kml.append('    </Style>')

    coords_linha = []
    
    for idx, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        
        lat = row.get('LATITUDE')
        lon = row.get('LONGITUDE')
        if pd.isna(lat) or pd.isna(lon): continue
        
        coords_linha.append(f"{lon},{lat},0")
        
        nome = str(row.get('PROTOCOLO', 'Ponto'))
        qtd = float(row.get('QTD PREVISTA DE POSTES', 0))
        cor = row.get('COR_ICONE', 'gray')
        
        desc = '<table border="1" style="border-collapse:collapse; width:100%;">'
        for c in colunas_exibir:
            # Usa a função formatadora que será passada do arquivo principal
            val = funcao_formatadora(c, row.get(c, ''))
            desc += f'<tr><td style="padding:3px;"><b>{html.escape(c)}</b></td><td style="padding:3px;">{val}</td></tr>'
        desc += '</table>'

        kml.append('    <Placemark>')
        kml.append(f'      <name>[{int(qtd)} Postes] {html.escape(nome)}</name>')
        kml.append(f'      <styleUrl>#style_{cor}</styleUrl>')
        kml.append(f'      <description><![CDATA[{desc}]]></description>')
        kml.append('      <Point>')
        kml.append(f'        <coordinates>{lon},{lat},0</coordinates>')
        kml.append('      </Point>')
        kml.append('    </Placemark>')

    if coords_linha:
        kml.append('    <Placemark>')
        kml.append('      <name>Traçado da Rota</name>')
        kml.append('      <Style><LineStyle><color>ff00ffff</color><width>3</width></LineStyle></Style>') 
        kml.append('      <LineString><tessellate>1</tessellate><coordinates>')
        kml.append(' '.join(coords_linha))
        kml.append('      </coordinates></LineString>')
        kml.append('    </Placemark>')

    kml.append('  </Document>')
    kml.append('</kml>')
    return '\n'.join(kml)

def gerar_kml_agrupado(df_kml, bases_records, nome_arquivo, colunas_exibir, bases_ativas, tipo_periodo, funcao_formatadora):
    """
    Gera o KML Clássico (Tático e Contínua Comum) usando a biblioteca simples do Python.
    (O código original aqui já tratava o kml.SimpleKML, mantivemos a estrutura lógica).
    """
    try:
        import simplekml
        kml = simplekml.Kml()
        kml.document.name = nome_arquivo

        fol_colors = ['ff1493', '00ffff', 'ff0000', '008000', 'ffa500', '800080', 'ffff00', '0000ff', 'ffc0cb', 'a52a2a']
        
        for idx_b, base_nome in enumerate(bases_ativas):
            cor_linha = fol_colors[idx_b % len(fol_colors)]
            df_base_kml = df_kml[df_kml['BASE_ATRIBUIDA'] == base_nome]
            pasta_base = kml.newfolder(name=f"Base: {base_nome}")
            
            for periodo_val in df_base_kml['PERIODO'].unique():
                df_periodo = df_base_kml[df_base_kml['PERIODO'] == periodo_val]
                p_nome = f"Semana {periodo_val}" if tipo_periodo == 'Semana' else f"Dia {periodo_val}"
                pasta_periodo = pasta_base.newfolder(name=p_nome)
                
                coords_linha = []
                for _, row in df_periodo.iterrows():
                    geom = row.get('ROTA_GEOMETRIA')
                    if isinstance(geom, list):
                        coords_linha.extend([(lon, lat) for lon, lat in geom])
                
                if coords_linha:
                    ls = pasta_periodo.newlinestring(name=f"Traçado - {p_nome}", coords=coords_linha)
                    ls.style.linestyle.color = cor_linha
                    ls.style.linestyle.width = 4

                for _, row in df_periodo.iterrows():
                    if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                    lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
                    if pd.isna(lat) or pd.isna(lon): continue

                    nome = str(row.get('PROTOCOLO', 'Ponto'))
                    is_super = str(row.get('SUPER_PONTO', '')).startswith('SIM')
                    is_prio = row.get('PRIORIDADE') == 'Sim'

                    pnt = pasta_periodo.newpoint(name=nome, coords=[(lon, lat)])
                    
                    if is_super:
                        pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/orange-blank.png'
                    elif is_prio:
                        pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png'
                    else:
                        pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png'

                    desc_html = '<table border="1" style="border-collapse:collapse; width:100%;">'
                    for col in colunas_exibir:
                        val = funcao_formatadora(col, row.get(col, ''))
                        desc_html += f'<tr><td style="padding:3px;"><b>{html.escape(col)}</b></td><td style="padding:3px;">{val}</td></tr>'
                    desc_html += '</table>'
                    pnt.description = desc_html

        return kml.kml()
    except ImportError:
        # Fallback caso a biblioteca simplekml falhe
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<kml xmlns=\"http://www.opengis.net/kml/2.2\"><Document><name>Erro</name><description>Biblioteca simplekml ausente no servidor.</description></Document></kml>"

# ==========================================
# PAINEL LATERAL (UI)
# ==========================================
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

def identificar_icone_folium(row, colunas_disponiveis):
    # Função simples para ícones do mapa folium baseada no tipo de nota
    if 'TIPO NOTA' in colunas_disponiveis:
        t = str(row.get('TIPO NOTA', '')).upper()
        if t in ['UNR', 'ASC']: return 'bolt'
        elif t in ['MGD', 'MTP']: return 'industry'
        elif t == 'DIF': return 'exclamation-triangle'
    return 'map-marker'
