import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
import gc
import os
import requests
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime




# Adaptacao exclusiva da leitura da planilha. O motor e os exportadores abaixo
# sao os mesmos do modulo original Lista Continua.
import unicodedata
from sklearn.cluster import DBSCAN
from openpyxl.styles import PatternFill, Font, Alignment

CAMPOS_VISIVEIS_NIP = [
    'NOTA','STATUS','TIPO DEMANDA','MUNICIPIO','REGIONAL','ENDEREÇO',
    'BAIRRO','LATITUDE','LONGITUDE','COLABORADOR','MATRICULA'
]

def normalize_cols(columns):
    def clean(x):
        x=str(x).strip().upper()
        x=unicodedata.normalize('NFKD',x)
        return ''.join(c for c in x if not unicodedata.combining(c)).replace('_',' ')
    return [clean(x) for x in columns]

def formata_campo_html(v):
    if pd.isna(v): return '-'
    return html.escape(str(v))

def formatar_moeda(v):
    try: return f'R$ {float(v):,.2f}'
    except (TypeError,ValueError): return '-'

def normalizar_municipios(df):
    return df

@st.cache_data(show_spinner=False)
def ler_planilha_cached(raw):
    return pd.read_excel(io.BytesIO(raw),engine='openpyxl')

def injetar_logo():
    for path in ['LOGO_NIP.png','logo_nip.png','assets/logo_nip.png']:
        if os.path.exists(path):
            st.logo(path)
            break

def adaptar_planilha_nip(df):
    """Adapta tanto a planilha Saneamento quanto a Lista Continua original.

    O campo COLABORADOR PLANEJADO (K) vale somente para Saneamento;
    a Lista Continua original utiliza LEVANTADOR (A).
    """
    df = df.copy()
    df.columns = normalize_cols(df.columns)
    cols = set(df.columns)

    if 'COLABORADOR PLANEJADO' in cols:
        # Modelo Saneamento: a coluna K identifica o colaborador planejado.
        if len(df.columns) < 11 or df.columns[10] != 'COLABORADOR PLANEJADO':
            raise ValueError('Na base de Saneamento, a coluna K deve ser COLABORADOR PLANEJADO.')
        rename = {
            'MUNICIPIO': 'MUNICIPIO', 'REGIONAL SERVICO': 'REGIONAL',
            'ENDERECO': 'ENDEREÇO', 'MATRICULA': 'MATRICULA',
            'COLABORADOR PLANEJADO': 'COLABORADOR',
        }
        df = df.rename(columns=rename)
        needed = ['NOTA', 'STATUS', 'TIPO DEMANDA', 'MUNICIPIO', 'REGIONAL',
                  'ENDEREÇO', 'BAIRRO', 'LATITUDE', 'LONGITUDE', 'COLABORADOR']
        missing = [c for c in needed if c not in df.columns]
        if missing:
            raise ValueError('Colunas ausentes na base Saneamento: ' + ', '.join(missing))
        if 'MATRICULA' not in df.columns:
            df['MATRICULA'] = ''
        df['PRIORIDADE'] = 'Não'
    elif {'LEVANTADOR', 'PROTOCOLO', 'LATITUDE', 'LONGITUDE'}.issubset(cols):
        # Modelo original Lista Continua: nao exigir COLABORADOR PLANEJADO na K.
        df['NOTA'] = df['PROTOCOLO']
        df['STATUS'] = df['STATUS LIST'] if 'STATUS LIST' in df.columns else (
            df['STATUS SISCO'] if 'STATUS SISCO' in df.columns else '')
        df['TIPO DEMANDA'] = df['TIPO NOTA'] if 'TIPO NOTA' in df.columns else ''
        df['COLABORADOR'] = df['LEVANTADOR']
        df['BAIRRO'] = df['LOCALIDADE'] if 'LOCALIDADE' in df.columns else ''
        if 'ENDERECO' in df.columns and 'ENDEREÇO' not in df.columns:
            df = df.rename(columns={'ENDERECO': 'ENDEREÇO'})
        if 'MATRICULA' not in df.columns:
            df['MATRICULA'] = ''
        if 'REGIONAL' not in df.columns:
            df['REGIONAL'] = ''
        if 'PRIORIDADE' not in df.columns:
            df['PRIORIDADE'] = 'Não'
    else:
        raise ValueError(
            'Formato nao reconhecido. Envie a base Saneamento com COLABORADOR PLANEJADO '
            'na coluna K ou a Lista Continua com LEVANTADOR e PROTOCOLO.'
        )

    # Mantem os dados operacionais originais para os cards e exportadores.
    for c in CAMPOS_VISIVEIS_NIP:
        if c not in df.columns:
            df[c] = ''
    df['PROTOCOLO'] = df['NOTA'].astype('string').str.replace(r'\.0$', '', regex=True)
    df['BASE_ATRIBUIDA'] = df['COLABORADOR'].fillna('').astype(str).str.strip().str.upper()
    if 'TIPO NOTA' not in df.columns:
        df['TIPO NOTA'] = df['TIPO DEMANDA']
    if 'LOCALIDADE' not in df.columns:
        df['LOCALIDADE'] = df['BAIRRO']
    return df


def haversine_vectorized(lat1, lon1, lat2, lon2):
    R = 6371.0 
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = np.sin(dlat/2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2.0)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return R * c


def haversine_scalar(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c


def fundir_super_pontos(df_tasks, raio_metros=5, agrupar_por_levantador=False):
    if df_tasks.empty or 'LATITUDE' not in df_tasks.columns or 'LONGITUDE' not in df_tasks.columns: return df_tasks, 0
    df_valid = df_tasks.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
    if df_valid.empty: return df_tasks, 0

    coords = np.radians(df_valid[['LATITUDE', 'LONGITUDE']].values)
    eps_rad = raio_metros / 6371000.0
    
    db = DBSCAN(eps=eps_rad, min_samples=1, algorithm='ball_tree', metric='haversine').fit(coords)
    df_valid['CLUSTER_ID'] = db.labels_
    
    if agrupar_por_levantador and 'LEVANTADOR' in df_valid.columns:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str) + "_" + df_valid['LEVANTADOR'].astype(str)
    else:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str)
    
    cluster_counts = df_valid['CLUSTER_GRP'].value_counts()
    single_clusters = cluster_counts[cluster_counts == 1].index
    multi_clusters = cluster_counts[cluster_counts > 1].index
    agrupado = []
    
    if len(multi_clusters) > 0:
        df_multi = df_valid[df_valid['CLUSTER_GRP'].isin(multi_clusters)]
        for c_id, group in df_multi.groupby('CLUSTER_GRP'):
            row_base = group.iloc[0].copy()
            qtd = len(group)
            orig_rows = group.to_dict('records')
            row_base['_ORIGINAL_ROWS'] = orig_rows
            
            def safe_list_join(col_name):
                itens = [str(x).strip() for x in group[col_name] if pd.notna(x) and str(x).lower() != 'nan']
                itens_unicos = list(dict.fromkeys(itens))
                return " | ".join(itens_unicos) if itens_unicos else "-"
            
            for col in group.columns:
                if col not in ['LATITUDE', 'LONGITUDE', 'CLUSTER_ID', 'CLUSTER_GRP', '_ORIGEM_BASE', 'PRIORIDADE']:
                    row_base[col] = safe_list_join(col)
                
            row_base['LATITUDE'] = group['LATITUDE'].mean()
            row_base['LONGITUDE'] = group['LONGITUDE'].mean()
            row_base['SUPER_PONTO'] = f"SIM ({qtd} un.)"
            if 'PRIORIDADE' in group.columns and 'Sim' in group['PRIORIDADE'].values:
                row_base['PRIORIDADE'] = 'Sim'
            agrupado.append(row_base)
            
    df_final_multi = pd.DataFrame(agrupado)
    df_single = df_valid[df_valid['CLUSTER_GRP'].isin(single_clusters)].copy()
    if not df_single.empty:
        df_single['SUPER_PONTO'] = "NÃO"
        dict_records = df_single.to_dict('records')
        df_single['_ORIGINAL_ROWS'] = [[r] for r in dict_records]
        
    df_final = pd.concat([df_single, df_final_multi], ignore_index=True).drop(columns=['CLUSTER_ID', 'CLUSTER_GRP'], errors='ignore')
    
    df_nan = df_tasks[df_tasks['LATITUDE'].isna() | df_tasks['LONGITUDE'].isna()].copy()
    if not df_nan.empty:
        df_nan['SUPER_PONTO'] = "NÃO"
        dict_records_nan = df_nan.to_dict('records')
        df_nan['_ORIGINAL_ROWS'] = [[r] for r in dict_records_nan]
        df_final = pd.concat([df_final, df_nan], ignore_index=True)
        
    return df_final, len(df_tasks) - len(df_final)


def formatar_planilha_openpyxl(writer, sheet_name, df_orig):
    """Aplica o padrão NIP (Azul Escuro) no arquivo Excel gerado."""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri')
    fill_superponto = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    font_prioridade = Font(color='FF0000', bold=True, name='Calibri')
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    worksheet.auto_filter.ref = worksheet.dimensions
    
    col_prio_idx = df_orig.columns.get_loc('PRIORIDADE') + 1 if 'PRIORIDADE' in df_orig.columns else None
    col_sp_idx = df_orig.columns.get_loc('SUPER_PONTO') + 1 if 'SUPER_PONTO' in df_orig.columns else None
    
    for r_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
        is_prio = col_prio_idx and str(row[col_prio_idx - 1].value).strip().upper() == 'SIM'
        is_sp = col_sp_idx and str(row[col_sp_idx - 1].value).strip().upper().startswith('SIM')
        for cell in row:
            if is_sp: cell.fill = fill_superponto
            if is_prio: cell.font = font_prioridade
    
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
        formatar_planilha_openpyxl(writer, 'Obras Roteirizadas', df_saida)
        
    return output.getvalue()


def gerar_excel_resumo_lista(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
        
        # Formata o header do resumo
        workbook = writer.book
        worksheet = writer.sheets['Resumo Operacional']
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
            worksheet.column_dimensions[column].width = min(max_length + 2, 40)
            
    return output.getvalue()


def limpar_colunas_lista(df_alvo, cols_originais=None):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    if 'PROTOCOLO' in df_alvo.columns:
        if 'NOTA' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['NOTA'])
        df_alvo = df_alvo.rename(columns={'PROTOCOLO': 'NOTA'})
        
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        if 'FISCAL' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['FISCAL'])
        df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'FISCAL'})
        
    # Colunas obrigatórias fixadas no início
    final_cols = ['FISCAL', 'DIA_SEMANA', 'DIA_MES', 'SUPER_PONTO', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM']
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


def gerar_txt_lista(df):
    """Gera um arquivo de texto com blocos contendo detalhes de cada obra e link do Google Maps."""
    linhas_txt = []
    for _, r in df.iterrows():
        nota = str(r.get('NOTA', r.get('PROTOCOLO', ''))).strip()
        if nota.lower() in ['nan', 'none', '']: nota = '-'
        
        nome = str(r.get('NOME', '')).strip()
        if nome.lower() in ['nan', 'none']: nome = ''
        
        # --- NOVOS CAMPOS: CONTA CONTRATO, INSTALAÇÃO E FASE ---
        conta_contrato = str(r.get('CONTA CONTRATO', '')).strip()
        if conta_contrato.endswith('.0'): conta_contrato = conta_contrato[:-2]
        if conta_contrato.lower() in ['nan', 'none', '']: conta_contrato = '-'
            
        instalacao = str(r.get('INSTALAÇÃO', r.get('INSTALACAO', ''))).strip()
        if instalacao.endswith('.0'): instalacao = instalacao[:-2]
        if instalacao.lower() in ['nan', 'none', '']: instalacao = '-'
            
        fase_raw = str(r.get('FASE', '')).strip().upper()
        if fase_raw == 'MO': fase_txt = '(MO) MONOFÁSICA'
        elif fase_raw == 'B': fase_txt = '(B) BIFÁSICA'
        elif fase_raw == 'TR': fase_txt = '(TR) TRIFÁSICA'
        elif fase_raw in ['NAN', 'NONE', '']: fase_txt = '-'
        else: fase_txt = fase_raw
        
        # --- ENDEREÇO E BAIRRO ---
        endereco = str(r.get('ENDEREÇO', r.get('ENDERECO', ''))).strip()
        if endereco.lower() in ['nan', 'none']: endereco = ''
        
        bairro = str(r.get('BAIRRO', r.get('LOCALIDADE', ''))).strip()
        if bairro.lower() in ['nan', 'none']: bairro = ''
        
        if bairro and endereco:
            endereco_completo = f"{endereco} / {bairro}"
        elif endereco:
            endereco_completo = endereco
        elif bairro:
            endereco_completo = bairro
        else:
            endereco_completo = '-'

        municipio = str(r.get('MUNICIPIO', '')).strip()
        if municipio.lower() in ['nan', 'none']: municipio = '-'
        
        # --- INFORMAÇÕES EXTRAS DA COLUNA AA ---
        info_extra = r.get('INFORMAÇÕES')
        if pd.isna(info_extra): 
            info_extra = r.get('INFORMACOES')
            
        if pd.isna(info_extra) or str(info_extra).strip().lower() in ['nan', 'none', '']:
            info_extra = ''
        else:
            info_extra = str(info_extra).strip()
        
        tipo_nota = str(r.get('TIPO NOTA', '')).strip()
        if tipo_nota.lower() in ['nan', 'none']: tipo_nota = ''

        lat = str(r.get('LATITUDE', '')).strip()
        lon = str(r.get('LONGITUDE', '')).strip()

        # --- MONTAGEM DO BLOCO NO FORMATO EXIGIDO ---
        bloco = []
        bloco.append(f"NOTA CCS: {nota}")
        bloco.append(f"NOME DO CLIENTE: {nome}")
        bloco.append(f"CONTA CONTRATO: {conta_contrato}")
        bloco.append(f"INSTALAÇÃO: {instalacao}")
        bloco.append(f"FASE: {fase_txt}")
        bloco.append(f"ENDEREÇO: {endereco_completo}")
        bloco.append(f"MUNICIPIO: {municipio}")
        bloco.append(f"INFORMAÇOES EXTRAS: {info_extra}")
        bloco.append(f"TIPO NOTA: {tipo_nota}")
        
        if lat and lon and lat.lower() != 'nan' and lon.lower() != 'nan':
            bloco.append(f"https://www.google.com.br/maps/place/{lat},{lon}")
        
        linhas_txt.append("\n".join(bloco))
        linhas_txt.append("\n-------------------------------------------------------------------------------------------------------------------------\n")
        
    return "".join(linhas_txt)


def gerar_kml_lista(df_kml, nome_arquivo, colunas_exibir, bases_ativas, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    # Injetando Estilos de Ícones Padrões e Contorno
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

    # Gerando os estilos das linhas dinamicamente para cada levantador ter uma cor diferente
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
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
                is_prio = str(r.get('PRIORIDADE')).upper() == 'SIM'
                
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

                if is_sp:
                    prot_list = [orig.get('PROTOCOLO', orig.get('NOTA', '')) for orig in r['_ORIGINAL_ROWS']]
                    prot_html = "<div style='max-height:95px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>" + "".join([f"<li style='margin-bottom:3px;'><b>[{i+1}]</b> {html.escape(str(pr))}</li>" for i, pr in enumerate(prot_list)]) + "</ul></div>"
                else:
                    prot_html = html.escape(str(r.get('PROTOCOLO', r.get('NOTA', ''))))

                desc += f'''
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;">Nota/Protocolo:</td><td style="padding:3px 6px; color:#333;">{prot_html}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Ordem:</td><td style="padding:3px 6px; color:#333;">{html.escape(str(ordem_txt))}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Ant.:</td><td style="padding:3px 6px; color:#333;">{dist_ant}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Próx.:</td><td style="padding:3px 6px; color:#333;">{dist_prox}</td></tr>
                '''

                for c in colunas_exibir:
                    if c.upper() in ['PROTOCOLO', 'NOTA', 'DIA_SEMANA', 'DIA_MES']: continue
                    
                    if c.upper() == 'SUPER_PONTO':
                        val_html = f"SIM ({qty} Obras)" if is_sp else "NÃO"
                        desc += f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;'>{html.escape(c)}:</td><td style='padding:3px 6px; color:#333;'>{val_html}</td></tr>"
                        continue

                    if is_sp and c.upper() not in ['LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'ZONA', 'REGIONAL']:
                        vals = [orig.get(c, '') for orig in r.get('_ORIGINAL_ROWS', [])]
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


def identificar_icone_folium(row, colunas_disponiveis):
    if 'TIPO NOTA' in colunas_disponiveis:
        t = str(row.get('TIPO NOTA', '')).upper()
        if t in ['UNR', 'ASC']: return 'bolt'
        elif t in ['MGD', 'MTP']: return 'industry'
        elif t == 'DIF': return 'exclamation-triangle'
    return 'map-marker'


injetar_logo()

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        if 'POSTE' in c.upper(): return str(int(float(v)))
        
        vf = float(v)
        if c.upper() in ['DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM']: 
            return f"{vf:.2f} KM"
        elif 'DISTANCIA' in c.upper(): 
            return f"{vf:.2f} Metros"
            
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(total_obras, qtd_equipes_ativas):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Modo:</b> <span style="color: #d9534f; font-weight: bold;">Lista Contínua</span></p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras}</span></p>
    </div>
    """

def render_metric_card(title, value, icon, border_color, bg_color):
    return f"""
    <div style="background-color: #ffffff; border-left: 5px solid {border_color}; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); display: flex; align-items: center; margin-bottom: 10px;">
        <div style="background-color: {bg_color}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; font-size: 20px; margin-right: 15px;">
            {icon}
        </div>
        <div>
            <p style="margin: 0; font-size: 11px; color: #666; text-transform: uppercase; font-weight: bold;">{title}</p>
            <p style="margin: 0; font-size: 22px; color: #333; font-weight: bold;">{value}</p>
        </div>
    </div>
    """



def ordenar_rota_rapida(obras, lat_inicial, lon_inicial, reversa=False):
    """Ordenação gulosa vetorizada, sem chamadas externas e sem risco de travar no solver."""
    if not obras:
        return []

    coords = np.array([[float(o['LATITUDE']), float(o['LONGITUDE'])] for o in obras], dtype=float)
    restantes = np.ones(len(obras), dtype=bool)
    ordem = []

    lat_atual, lon_atual = float(lat_inicial), float(lon_inicial)

    if reversa and len(obras) > 1:
        # Começa pelo ponto mais distante do centro e depois segue sempre ao mais próximo.
        d0 = (coords[:, 0] - lat_atual) ** 2 + ((coords[:, 1] - lon_atual) * np.cos(np.radians(lat_atual))) ** 2
        idx = int(np.argmax(d0))
        ordem.append(obras[idx])
        restantes[idx] = False
        lat_atual, lon_atual = coords[idx]

    while restantes.any():
        idxs = np.flatnonzero(restantes)
        cand = coords[idxs]
        cos_lat = np.cos(np.radians(lat_atual))
        dist2 = (cand[:, 0] - lat_atual) ** 2 + ((cand[:, 1] - lon_atual) * cos_lat) ** 2
        idx = int(idxs[int(np.argmin(dist2))])
        ordem.append(obras[idx])
        restantes[idx] = False
        lat_atual, lon_atual = coords[idx]

    return ordem


def obter_rota_ruas_segura(lat1, lon1, lat2, lon2, url_osrm_base, velocidade_media_kmh, timeout_s=5):
    """Consulta OSRM com timeout curto; em qualquer falha retorna linha reta e tempo estimado."""
    dist_km = haversine_scalar(lat1, lon1, lat2, lon2) * 1.3
    fallback = ([[lon1, lat1], [lon2, lat2]], (dist_km / max(float(velocidade_media_kmh), 1.0)) * 3600.0)

    if not url_osrm_base:
        return fallback

    try:
        base = str(url_osrm_base).rstrip('/')
        url = f"{base}/route/v1/driving/{lon1},{lat1};{lon2},{lat2}"
        resp = requests.get(
            url,
            params={'overview': 'full', 'geometries': 'geojson', 'steps': 'false'},
            timeout=timeout_s,
            headers={'User-Agent': 'ListaContinua/1.0'}
        )
        resp.raise_for_status()
        data = resp.json()
        rotas = data.get('routes') or []
        if not rotas:
            return fallback
        rota = rotas[0]
        geom = ((rota.get('geometry') or {}).get('coordinates')) or []
        if len(geom) < 2:
            return fallback
        duracao = float(rota.get('duration') or fallback[1])
        return geom, duracao
    except Exception:
        return fallback

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_lista': False, 'vrp_status_lista': "IDLE", 'vrp_state_lista': {}, 'df_routed_lista': pd.DataFrame(), 'colunas_exibir_lista': [], 'colunas_originais_lista': []})
    for k in ['bytes_zip_xl_lista', 'bytes_zip_kml_lista', 'bytes_zip_gpx_lista', 'bytes_zip_txt_lista', 'start_time_run_lista', 'start_time_pkg_lista', 'df_unallocated_lista', 'df_correcao_lista']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido_lista" not in st.session_state: st.session_state.roteamento_concluido_lista = False
if "vrp_status_lista" not in st.session_state: st.session_state.vrp_status_lista = "IDLE"

status_exec = st.session_state.vrp_status_lista
is_done = st.session_state.roteamento_concluido_lista
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📜 Roteirização</h1>", unsafe_allow_html=True)
st.info("💡 Gera uma lista de execução contínua com base na atribuição já existente na planilha de obras.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Parâmetros de Rota", expanded=True):
        st.success("📦 **Modo Contínuo:** Todas as obras serão alocadas numa lista única.")
        trava_global = st.number_input("Trava Total de Obras", min_value=0, value=0, step=50, disabled=is_locked)
        data_ini = st.date_input("📅 Data Base (Início):", value=datetime.today(), disabled=is_locked)
        
        st.markdown("---")
        obras_por_dia_est = st.number_input("Meta de Obras/Dia (Cálculo de Postes):", min_value=1, value=4, disabled=is_locked)
        
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)
        
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=False, disabled=is_locked, help="Ative apenas quando precisar do traçado viário real. Desativado, o motor usa distância geográfica e fica muito mais rápido.")

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_lista.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_lista', b"vazio"), file_name=f"ListaContinua_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("📝 Baixar Relatórios (TXT)", data=st.session_state.get('bytes_zip_txt_lista', b"vazio"), file_name=f"ListaContinua_TXT - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_lista', b"vazio"), file_name=f"ListaContinua_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_lista', b"vazio"), file_name=f"ListaContinua_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed_lista.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    df_c = st.session_state.get('df_correcao_lista', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica:</b> Estas obras apresentaram coordenadas em branco, zeradas ou invertidas. Elas foram isoladas para não corromper o mapa.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.df_routed_lista['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_lista.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed_lista.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = len(dfr_t)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    
    tsp = sum(1 for _, r in dfr_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Roteirizadas", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(tsp), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    html_msg = f"""
    <div style='background-color: #d4edda; border-left: 8px solid #28a745; padding: 20px; border-radius: 8px; margin-bottom: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
        <h2 style='color: #155724; margin-top: 0; font-weight: 900; letter-spacing: 1px; font-size: 24px;'>🎉 100% DAS OBRAS ROTEIRIZADAS!</h2>
        <p style='color: #155724; font-size: 16px; margin-bottom: 10px;'>Todas as <b>{tr} tarefas válidas</b> presentes na sua planilha foram organizadas e processadas com sucesso!</p>
        <hr style='border-top: 1px solid #c3e6cb; margin: 15px 0;'>
        <p style='color: #155724; font-size: 14px; margin-bottom: 0;'>
            <b>💡 Resumo Técnico:</b> O módulo de <b>Lista Contínua</b> opera sem limites de jornada ou cota de dias. A Inteligência Artificial respeitou a atribuição que você já fez na planilha e traçou a melhor rota possível (do ponto A ao Z) para <b>todas</b> as obras de cada levantador, sem deixar absolutamente nada para trás.
        </p>
    </div>
    """
    st.markdown(html_msg, unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Operacional")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        pts = [p for _, r in db.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
        folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
        folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
        
        for r in db.to_dict('records'):
            if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
            c_i = 'red' if str(r.get('PRIORIDADE')) == 'Sim' else 'blue'
            ic = identificar_icone_folium(r, dfr.columns)
            er = "".join([f"<tr><td><b>{html.escape(c)}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_lista if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
            pop_html = f'<div style="width:250px;"><b>Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
            folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📊 Dados Tabulares")
    st.data_editor(st.session_state.df_routed_lista.reindex(columns=CAMPOS_VISIVEIS_NIP), use_container_width=True)

elif status_exec == "IDLE":
    st.markdown("### 📁 Demandas (Obras e Equipes)")
    st.info("💡 Apenas faça o upload da sua planilha de Obras. A IA identificará o Levantador/Fiscal Responsável automaticamente.")
    task_files = st.file_uploader("Suba a planilha com a Lista Contínua", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if not task_files: st.stop()
    
    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        dft = adaptar_planilha_nip(dft)
        if not dfs: st.session_state.colunas_originais_lista = CAMPOS_VISIVEIS_NIP.copy()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    agent_col = None
    for col in ['BASE_ATRIBUIDA', 'LEVANTADOR_RESPONSAVEL', 'LEVANTADOR', 'FISCAL', 'EQUIPE', 'NOME_FISCAL']:
        if col in df_tasks.columns:
            agent_col = col; break
            
    if not agent_col:
        st.error("❌ A planilha de Obras não possui uma coluna de responsável. Insira uma coluna chamada 'LEVANTADOR' ou 'FISCAL'.")
        st.stop()
        
    if agent_col != 'BASE_ATRIBUIDA': df_tasks = df_tasks.rename(columns={agent_col: 'BASE_ATRIBUIDA'})
    df_tasks['BASE_ATRIBUIDA'] = df_tasks['BASE_ATRIBUIDA'].astype(str).str.strip().str.upper()
    df_tasks = df_tasks[~df_tasks['BASE_ATRIBUIDA'].isin(['NAN', 'NONE', ''])]

    qtd_eq = df_tasks['BASE_ATRIBUIDA'].nunique()
    sb_html.markdown(render_sidebar_card("Ilimitada", qtd_eq), unsafe_allow_html=True)

    st.markdown("---")
    falta = [c for c in ['LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas vitais: {', '.join(falta)}."); st.stop()
    
    c1_f, c2_f = st.columns([1, 1])
    with c1_f:
        cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO'
        if cs in df_tasks.columns:
            df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
            opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
            sel_s = st.multiselect("1. Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
            if not sel_s: st.stop()
            df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()
            
    with c2_f:
        if 'PRIORIDADE' in df_tasks.columns:
            df_tasks['PRIORIDADE'] = df_tasks['PRIORIDADE'].apply(
                lambda x: 'Sim' if pd.notna(x) and str(x).strip() not in ['', '0', 'NÃO', 'NAO', 'FALSE'] else 'Não'
            )
            st.success("🎯 **Prioridade Ativada:** As obras marcadas na coluna 'PRIORIDADE' serão destacadas no roteiro e no KML.")
        elif 'TIPO NOTA' in df_tasks.columns:
            df_tasks['TIPO NOTA'] = df_tasks['TIPO NOTA'].astype(str).str.strip().str.upper()
            opts_n = sorted([str(x) for x in df_tasks['TIPO NOTA'].unique() if str(x) != 'NAN'])
            sel_p = st.multiselect("🚨 2. Obras de Alta Prioridade:", options=opts_n, default=[n for n in opts_n if n in ['ASC', 'CCF', 'DIF', 'MGD', 'MTP', 'SID']])
            df_tasks['PRIORIDADE'] = df_tasks['TIPO NOTA'].apply(lambda x: 'Sim' if str(x) in sel_p else 'Não')
        else: 
            df_tasks['PRIORIDADE'] = 'Não'

    st.markdown("#### 🌍 Limpeza Geográfica")
    if st.button("⏹️ Abortar Roteamento", use_container_width=True): limpar_roteirizador(); st.stop()
    
    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    m_na, m_0 = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna(), (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    df_tasks.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df_tasks.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)
    df_tasks.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    m_inv = abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])
    df_tasks.loc[m_inv & ~m_na & ~m_0 & ~m_pos, 'MOTIVO_REJEICAO'] = 'Coordenada Invertida'
    
    mc = m_na | m_0 | m_pos | m_inv
    if mc.sum() > 0: df_rej = pd.concat([df_rej, df_tasks[mc].copy()], ignore_index=True); df_tasks = df_tasks[~mc].copy()
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']; df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    st.session_state.df_correcao_lista = df_rej
    if not df_rej.empty: 
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_rej)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa:</b> Apresentaram coordenadas zeradas, invertidas ou em branco. Foram removidas da roteirização.
            </p>
        </div>
        """, unsafe_allow_html=True)

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    dfs_fundidos = []
    for base in df_tasks['BASE_ATRIBUIDA'].unique():
        df_base = df_tasks[df_tasks['BASE_ATRIBUIDA'] == base].copy()
        df_base_f, _ = fundir_super_pontos(df_base, raio_metros=raio_sp, agrupar_por_levantador=True)
        dfs_fundidos.append(df_base_f)
    df_tasks = pd.concat(dfs_fundidos, ignore_index=True)

    if trava_global > 0: df_tasks = df_tasks.head(trava_global)
    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])

    df_ta = df_tasks.copy()
    df_u = pd.DataFrame()

    st.session_state.df_unallocated_lista, total_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    
    sb_html.markdown(render_sidebar_card(total_alocadas, qtd_eq), unsafe_allow_html=True)

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = CAMPOS_VISIVEIS_NIP.copy()
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis:", CAMPOS_VISIVEIS_NIP, default=CAMPOS_VISIVEIS_NIP)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        st.session_state.update({'colunas_exibir_lista': colunas_exibir})
        st.session_state.vrp_state_lista = {'config': {'velocidade_media_kmh': 30.0, 'sentido_rota': sentido_rota, 'url_osrm_base': url_osrm, 'tracado_real': usa_osrm, 'data_inicio': data_ini, 'obras_por_dia_est': obras_por_dia_est}, 'b_names': list(set(df_ta['BASE_ATRIBUIDA'].unique())), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': [], 'current_geoms': []}
        st.session_state.vrp_status_lista = "RUNNING"; tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor Contínuo")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_lista', time.time())
    if 'start_time_run_lista' not in st.session_state: st.session_state.start_time_run_lista = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_lista; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            if not oe:
                st_v['b_idx'] += 1
                st.session_state.vrp_state_lista = st_v
                tentar_rerun()
                st.stop()

            bl = sum(float(x['LATITUDE']) for x in oe) / len(oe)
            bL = sum(float(x['LONGITUDE']) for x in oe) / len(oe)

            # O solver antigo podia ficar preso aguardando matriz/OSRM.
            # A ordenação abaixo é local, determinística e não depende de rede.
            ot = ordenar_rota_rapida(
                oe, bl, bL,
                reversa=("Varredura Reversa" in cfg.get('sentido_rota', "Lógica Padrão"))
            )
            
            rf = []
            c_l, c_L = bl, bL
            for o in ot:
                vkr = haversine_vectorized(c_l, c_L, o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                rf.append({'o': o, 'la': c_l, 'La': c_L, 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 'dk': vk})
                c_l, c_L = o['LATITUDE'], o['LONGITUDE']
                
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_lista = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (8 if cfg.get('tracado_real') else len(rf)), len(rf)) 
            
            for i in range(oi, ei):
                it = rf[i]
                fallback = ([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600)
                
                if not cfg.get('tracado_real'):
                    gd.append(fallback)
                else:
                    if i % 5 == 0: sgt.info(f"🛣️ Traçando arruamento **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf))
                    time.sleep(0.02)
                    try:
                        res = obter_rota_ruas_segura(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'], timeout_s=5)
                        if not res or len(res) == 0 or len(res[0]) == 0: gd.append(fallback)
                        else: gd.append(res)
                    except: gd.append(fallback)
                
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_lista = st_v; tentar_rerun(); st.stop()
            
            rdf, og = [], 1
            for it, (g, ds) in zip(rf, gd):
                ob = it['o']; ob['ORDEM'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, round(it['dk'], 2)
                ob['ROTA_GEOMETRIA'], ob['PERIODO'] = g, "Único"
                rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state_lista = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        st.session_state.df_routed_lista = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status_lista = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    
    df_routed = st.session_state.df_routed_lista.copy()
    df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = df_routed.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)

    d_fmt = datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx, bu_txt = io.BytesIO(), io.BytesIO(), io.BytesIO(), io.BytesIO()
    
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, \
             zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, \
             zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg, \
             zipfile.ZipFile(bu_txt, 'w', zipfile.ZIP_DEFLATED) as zt:
            
            data_ini = st.session_state.vrp_state_lista.get('config', {}).get('data_inicio', datetime.today())
            obras_por_dia_est = st.session_state.vrp_state_lista.get('config', {}).get('obras_por_dia_est', 4.0)
            
            if isinstance(data_ini, datetime): data_ini = data_ini.date()
            dia_mes_str = data_ini.strftime("%d/%m/%Y")
            dias_semana_pt = {0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA", 3: "QUINTA-FEIRA", 4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO"}
            dia_semana_str = dias_semana_pt[data_ini.weekday()]

            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                
                qtd_obras = 0
                qtd_postes = 0
                for _, r in db.iterrows():
                    if isinstance(r.get('_ORIGINAL_ROWS'), list):
                        qtd_obras += len(r['_ORIGINAL_ROWS'])
                        for orig in r['_ORIGINAL_ROWS']:
                            pv = []
                            for k, v in orig.items():
                                if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                    try: 
                                        val = float(v)
                                        if val > 0: pv.append(val)
                                    except: pass
                            if pv: qtd_postes += min(pv)
                    else:
                        qtd_obras += 1
                        pv = []
                        for k, v in r.items():
                            if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                try: 
                                    val = float(v)
                                    if val > 0: pv.append(val)
                                except: pass
                        if pv: qtd_postes += min(pv)

                postes_dia = (qtd_postes / (qtd_obras / float(obras_por_dia_est))) if qtd_obras > 0 else 0
                postes_semana = postes_dia * 5.0

                res.append({
                    'Equipe': b, 
                    'Obras Roteirizadas': qtd_obras, 
                    'Postes/Dia (Est.)': int(round(postes_dia)),
                    'Postes/Semana (Est.)': int(round(postes_semana)),
                    'Postes Total': int(round(qtd_postes)),
                    'Super Pontos': qs, 
                    'Prioridades Atendidas': len(db[db['PRIORIDADE']=='Sim']), 
                    'KM Total Previsto': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)
                })
            zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_lista(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_lista', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
            
            linhas_gerais = []
            for _, r in df_routed.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
                sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else "NÃO"
                
                if is_sp:
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = r.copy()
                        for k, v in orig.items(): 
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'ROTA_GEOMETRIA', 'PERIODO']: nr[k] = v
                        nr['DIA_SEMANA'] = dia_semana_str
                        nr['DIA_MES'] = dia_mes_str
                        nr['SUPER_PONTO'] = sp_text
                        linhas_gerais.append(nr)
                else:
                    rn = r.copy()
                    rn['DIA_SEMANA'] = dia_semana_str
                    rn['DIA_MES'] = dia_mes_str
                    rn['SUPER_PONTO'] = sp_text
                    linhas_gerais.append(rn)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            
            for c in df_excel_full.columns:
                if 'POSTE' in c.upper():
                    df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

            col_exibir = st.session_state.colunas_exibir_lista.copy()
            if 'DIA_SEMANA' not in col_exibir: col_exibir.insert(0, 'DIA_SEMANA')
            if 'DIA_MES' not in col_exibir: col_exibir.insert(1, 'DIA_MES')
            if 'SUPER_PONTO' not in col_exibir: col_exibir.insert(2, 'SUPER_PONTO')

            dfg_total = limpar_colunas_lista(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
            dfg_total = dfg_total.loc[:, ~dfg_total.columns.duplicated()].copy()
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.xlsx", gerar_excel_lista(dfg_total, st.session_state.colunas_originais_lista))
            
            # --- TXT TOTAL ---
            txt_total = gerar_txt_lista(df_excel_full)
            zt.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.txt", txt_total.encode('utf-8'))
            
            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                dfk_total['DIA_SEMANA'] = dia_semana_str
                dfk_total['DIA_MES'] = dia_mes_str
                dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                
                ks_tot = gerar_kml_lista(dfk_total, "ROTA TOTAL", col_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), formatar_valor_coluna)
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks_tot.encode('utf-8'))
                zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

            for base in df_routed['BASE_ATRIBUIDA'].unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == base]
                if not df_base_excel.empty:
                    dfg = limpar_colunas_lista(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
                    dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
                    for cc in dfg.columns:
                        if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                    zx.writestr(f"Rotas_{d_fmt}/Rota_{b_safe}.xlsx", gerar_excel_lista(dfg, st.session_state.colunas_originais_lista))
                    
                    # --- TXT INDIVIDUAL ---
                    txt_ind = gerar_txt_lista(df_base_excel)
                    if txt_ind:
                        zt.writestr(f"Relatorios_TXT_{d_fmt}/Relatorio_{b_safe}.txt", txt_ind.encode('utf-8'))
                    
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == base) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if not dfk_base.empty:
                    dfk_base['DIA_SEMANA'] = dia_semana_str
                    dfk_base['DIA_MES'] = dia_mes_str
                    dfk_base['SUPER_PONTO'] = dfk_base.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                    
                    ks = gerar_kml_lista(dfk_base, f"Rota {b_safe}", col_exibir, [base], formatar_valor_coluna)
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kml", ks.encode('utf-8'))
                    zg.writestr(f"GPX_{d_fmt}/Rota_{b_safe}.gpx", gerar_gpx_simples(dfk_base, f"Rota {b_safe}").encode('utf-8'))

        st.session_state.bytes_zip_xl_lista = bu_xl.getvalue()
        st.session_state.bytes_zip_kml_lista = bu_kml.getvalue()
        st.session_state.bytes_zip_gpx_lista = bu_gpx.getvalue()
        st.session_state.bytes_zip_txt_lista = bu_txt.getvalue()
        st.session_state.roteamento_concluido_lista = True; st.session_state.vrp_status_lista = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_lista = "IDLE"
