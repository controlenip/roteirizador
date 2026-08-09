import pandas as pd
import io
import re
import html
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from modules.data_processing import formata_campo_html

def identificar_icone_folium(row, colunas):
    tipo_str = str(row.get('TIPO LIGACAO', '')) + str(row.get('SERVICO', '')) + str(row.get('TIPO NOTA', ''))
    tipo_str = tipo_str.upper()
    if row.get('PROTOCOLO') == 'RETORNO_BASE': return 'home'
    if row.get('PROTOCOLO') == 'PAUSA_ALMOCO': return 'cutlery'
    if 'NOVA' in tipo_str or 'LIGACAO' in tipo_str or 'UNI' in tipo_str or 'UNR' in tipo_str: return 'bolt'
    if 'MANUT' in tipo_str or 'REPARO' in tipo_str: return 'wrench'
    if 'INSP' in tipo_str or 'VISTORIA' in tipo_str: return 'eye-open'
    return 'info-sign'

def renderizar_painel_lateral(cap_eq, obras_prontas, eq_sel, cap_tot):
    return f'''
    <label style="font-size: 13.5px; font-weight: 700; color: #0D256C; margin-bottom: 4px; display: block; margin-top: 10px;">Capacidade da Rota (Por Equipe):</label>
    <div style="background-color: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; margin-bottom: 12px; color: #d9534f; font-weight: 800; font-size: 15px; text-align: center;">
        {cap_eq} Obras
    </div>
    
    <label style="font-size: 13.5px; font-weight: 700; color: #0D256C; margin-bottom: 4px; display: block;">Obras Prontas p/ Roteirizar:</label>
    <div style="background-color: #edf7ed; border: 1px solid #c8e6c9; border-radius: 6px; padding: 8px 12px; margin-bottom: 12px; color: #2e7d32; font-weight: 800; font-size: 15px; text-align: center;">
        {obras_prontas} Obras
    </div>
    
    <label style="font-size: 13.5px; font-weight: 700; color: #0D256C; margin-bottom: 4px; display: block;">Equipes Selecionadas:</label>
    <div style="background-color: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; margin-bottom: 12px; color: #d9534f; font-weight: 800; font-size: 15px; text-align: center;">
        {eq_sel} Equipes
    </div>
    
    <label style="font-size: 13.5px; font-weight: 700; color: #0D256C; margin-bottom: 4px; display: block;">Quantidade Total Projetada:</label>
    <div style="background-color: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; margin-bottom: 5px; color: #d9534f; font-weight: 800; font-size: 15px; text-align: center;">
        {cap_tot} Obras Max.
    </div>
    '''

def gerar_excel_bytes(df, col_prioridade, colunas_originais=None):
    df_export = df.copy()
    if 'PROTOCOLO' in df_export.columns: 
        df_export = df_export[~df_export['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
        
    unpacked_rows = []
    for _, row in df_export.iterrows():
        if '_ORIGINAL_ROWS' in row and isinstance(row['_ORIGINAL_ROWS'], list):
            for orig in row['_ORIGINAL_ROWS']:
                new_row = orig.copy()
                for vrp_col in ['NOME_DIA', 'ORDEM', 'SEMANA', 'DIA', 'PERIODO', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'SUPER_PONTO', 'BASE_ATRIBUIDA', 'PRIORIDADE', 'DISTANCIA_REDE_METROS', 'POSTES PREVISTOS', 'LATITUDE_REDE', 'LONGITUDE_REDE']:
                    if vrp_col in row:
                        new_row[vrp_col] = row[vrp_col]
                unpacked_rows.append(new_row)
        else:
            unpacked_rows.append(row.to_dict())
            
    df_export = pd.DataFrame(unpacked_rows)
    for col in ['ROTA_GEOMETRIA', 'STATUS LIST', 'INICIO AVARIA', 'STATUS ATUAL (LEVANTAMENTO)', 'DESCRICAO', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'MUN_LIMPO_CALC']:
        if col in df_export.columns: df_export = df_export.drop(columns=[col], errors='ignore')

    if colunas_originais:
        cols_atuais = df_export.columns.tolist()
        cols_originais_validas = [c for c in colunas_originais if c in cols_atuais]
        cols_novas_geradas = [c for c in cols_atuais if c not in colunas_originais]
        df_export = df_export[cols_originais_validas + cols_novas_geradas]
        
    buf_xl = io.BytesIO()
    with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Roteiro')
        ws = writer.sheets['Roteiro']
        header_fill = PatternFill(start_color='0D256C', end_color='0D256C', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            
        col_types = {}
        for col_idx, col_name in enumerate(df_export.columns, 1):
            col_letter = get_column_letter(col_idx)
            col_name_upper = str(col_name).upper()
            if any(x in col_name_upper for x in ['NOME', 'CLIENTE', 'ENDEREÇO', 'INFORMAÇ']): ws.column_dimensions[col_letter].width = 45.0
            elif any(x in col_name_upper for x in ['PROTOCOLO', 'MUNICIPIO', 'BASE', 'LOCALIDADE']): ws.column_dimensions[col_letter].width = 25.0
            else: ws.column_dimensions[col_letter].width = 18.0
                
            if col_name_upper in ['NOME_DIA', 'ORDEM', 'SEMANA', 'DIA', 'PERIODO', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'TEMPO_VIAGEM_MINUTOS', 'PRIORIDADE', 'HORA_INICIO', 'HORA_FIM', 'DISTANCIA_REDE_METROS', 'POSTES PREVISTOS', 'LATITUDE_REDE', 'LONGITUDE_REDE']:
                col_types[col_idx] = center_align
            else: col_types[col_idx] = left_align

        red_font = Font(color="FF0000", bold=True)
        prio_idx = df_export.columns.get_loc('PRIORIDADE') + 1 if 'PRIORIDADE' in df_export.columns else None
        prio_target_idx = df_export.columns.get_loc(col_prioridade) + 1 if col_prioridade in df_export.columns else None

        for row_idx in range(2, len(df_export) + 2):
            for col_idx in range(1, len(df_export.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = col_types.get(col_idx, left_align)
            if prio_idx and ws.cell(row=row_idx, column=prio_idx).value == "Sim":
                ws.cell(row=row_idx, column=prio_idx).font = red_font
                if prio_target_idx and col_prioridade != "Nenhuma":
                    try: ws.cell(row=row_idx, column=prio_target_idx).font = red_font
                    except: pass
                    
        if 'SUPER_PONTO' in df_export.columns:
            try:
                idx_super = df_export.columns.get_loc('SUPER_PONTO') + 1
                yellow_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                for row_idx in range(2, len(df_export) + 2):
                    val = str(ws.cell(row=row_idx, column=idx_super).value)
                    if val.startswith("SIM"):
                        for col_idx_f in range(1, len(df_export.columns) + 1):
                            ws.cell(row=row_idx, column=col_idx_f).fill = yellow_fill
            except: pass
                    
    return buf_xl.getvalue()

def gerar_excel_resumo_bytes(df):
    buf_xl = io.BytesIO()
    with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resumo')
        ws = writer.sheets['Resumo']
        header_fill = PatternFill(start_color='0D256C', end_color='0D256C', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')
        for cell in ws[1]:
            cell.fill = header_fill; cell.font = header_font; cell.alignment = center_align
        for col_idx, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            if str(col_name).upper() == 'LEVANTADOR':
                ws.column_dimensions[col_letter].width = 45.0
                col_align = left_align
            else:
                ws.column_dimensions[col_letter].width = 22.0
                col_align = center_align
            for row_idx in range(2, len(df) + 2): ws.cell(row=row_idx, column=col_idx).alignment = col_align
    return buf_xl.getvalue()

def gerar_kml_agrupado(df_rota, bases_records, doc_name, cols_exibir, lista_todas_bases=None, tipo_periodo="Dia"):
    # Código KML gigante abstraído para não cortar a resposta.
    # [COLE AQUI O CONTEÚDO ORIGINAL DA FUNÇÃO `gerar_kml_agrupado` DO SEU SCRIPT]
    # Linha 441 até a linha 613 do seu script original.
    pass
