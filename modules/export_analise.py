import pandas as pd
import io
import html
import os
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment


def _valor_alias(row, aliases, default='-'):
    for col in aliases:
        val = row.get(col)
        if pd.notna(val) and str(val).strip().lower() not in ['', 'nan', 'none']:
            return str(val).strip()
    return default


def gerar_excel_analise(dict_dfs):
    """Gera um arquivo Excel com múltiplas abas, preservando os dados e otimizando a formatação."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        escreveu_aba = False
        for sheet_name, df in dict_dfs.items():
            if df is None or df.empty:
                continue

            escreveu_aba = True
            df_saida = df.drop(
                columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME'],
                errors='ignore'
            ).copy()

            safe_sheet_name = str(sheet_name)[:31]
            df_saida.to_excel(writer, index=False, sheet_name=safe_sheet_name)

            worksheet = writer.sheets[safe_sheet_name]
            header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True, name='Calibri')

            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

            worksheet.auto_filter.ref = worksheet.dimensions
            worksheet.freeze_panes = 'A2'

            # Em bases grandes, medir todas as células para largura pode ser caro.
            # Amostra até 750 linhas sem alterar qualquer conteúdo exportado.
            max_linhas_largura = min(worksheet.max_row, 751)
            for col_cells in worksheet.iter_cols(min_row=1, max_row=max_linhas_largura):
                max_length = 0
                column = col_cells[0].column_letter
                for cell in col_cells:
                    try:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                worksheet.column_dimensions[column].width = min(max_length + 2, 60)

        # Proteção para cenários onde todos os dataframes recebidos estejam vazios.
        if not escreveu_aba:
            pd.DataFrame({'INFORMACAO': ['Nenhum registro para exportar']}).to_excel(
                writer, index=False, sheet_name='Sem Dados'
            )

    return output.getvalue()


def gerar_kml_analise(df, nome_documento='Análise Cruzada NIP'):
    # Trabalha sempre em cópia para não alterar o DataFrame usado pela tela/exportações.
    df = df.copy()

    kml = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        '<Document>',
        f'<name>{html.escape(str(nome_documento))}</name>'
    ]

    styles = {
        'red': 'http://maps.google.com/mapfiles/kml/paddle/red-blank.png',
        'orange': 'http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png',
        'green': 'http://maps.google.com/mapfiles/kml/paddle/grn-blank.png',
        'purple': 'http://maps.google.com/mapfiles/kml/paddle/purple-blank.png',
        'blue': 'http://maps.google.com/mapfiles/kml/paddle/blu-blank.png'
    }

    for color, url in styles.items():
        kml.append(
            f'<Style id="style_{color}"><IconStyle><scale>1.1</scale><Icon><href>{url}</href></Icon></IconStyle>'
            '<LabelStyle><scale>0</scale></LabelStyle></Style>'
        )

    kml.append(
        '<Style id="style_black"><IconStyle><scale>1.1</scale><color>ff000000</color>'
        '<Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon></IconStyle>'
        '<LabelStyle><scale>0</scale></LabelStyle></Style>'
    )

    if 'CLUSTER_ID' not in df.columns:
        df['CLUSTER_ID'] = range(len(df))

    clusters = []
    for cid, grp in df.groupby('CLUSTER_ID'):
        lat_val = grp['LAT_CENTRO_CLUSTER'].iloc[0] if 'LAT_CENTRO_CLUSTER' in grp.columns else grp['LATITUDE'].mean()
        lon_val = grp['LONG_CENTRO_CLUSTER'].iloc[0] if 'LONG_CENTRO_CLUSTER' in grp.columns else grp['LONGITUDE'].mean()
        lat = str(lat_val).strip()
        lon = str(lon_val).strip()
        if not lat or not lon or lat.lower() == 'nan' or lon.lower() == 'nan':
            continue

        c_names = grp['COR_NOME'].astype(str).tolist()

        if any('Preto' in c for c in c_names) or any('Inválidas' in c for c in c_names):
            c_nome = next(c for c in c_names if 'Inválidas' in c or 'Preto' in c)
        elif any('Vermelho' in c for c in c_names) or any('Duplicadas' in c for c in c_names):
            c_nome = next(c for c in c_names if 'Duplicadas' in c or 'Vermelho' in c)
        elif len(grp) > 1:
            c_nome = '🟠 Notas Próximas'
        else:
            c_nome = c_names[0]

        if 'Inválidas' in c_nome or 'Preto' in c_nome:
            c_mapa = 'black'
        elif 'Duplicadas' in c_nome or 'Vermelho' in c_nome:
            c_mapa = 'red'
        elif 'Próximas' in c_nome or 'Laranja' in c_nome:
            c_mapa = 'orange'
        else:
            c_mapa = str(grp['COR_MAPA'].iloc[0])
            if c_mapa not in styles and c_mapa != 'black':
                c_mapa = 'blue'

        total_cluster = int(grp['QTD_OBRAS_CLUSTER'].iloc[0]) if 'QTD_OBRAS_CLUSTER' in grp.columns and pd.notna(grp['QTD_OBRAS_CLUSTER'].iloc[0]) else len(grp)
        titulo_card = f"📍 Obras no Local ({len(grp)})" if len(grp) == total_cluster else f"📍 {len(grp)} visíveis de {total_cluster} obras no local"
        desc = f'''<![CDATA[
        <div style="font-family:sans-serif; width:300px; max-height:360px; overflow-y:auto; border-radius:8px; overflow-x:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
            <div style="background:#0D256C; color:#ffffff; padding:8px 10px; font-size:13px; font-weight:bold; text-align:center; position:sticky; top:0;">{titulo_card}</div>
            <div style="padding:10px; background:#fafafa; font-size:12px;">
        '''
        if 'DISTANCIA_MAX_CLUSTER_M' in grp.columns:
            desc += f"<div style='margin-bottom:6px;color:#555;'><b>Cluster:</b> {total_cluster} obra(s) | diâmetro máx.: {grp['DISTANCIA_MAX_CLUSTER_M'].iloc[0]} m</div>"

        nomes_notas = []
        for _, r in grp.iterrows():
            n = html.escape(str(r.get('NOTA', '')))
            nomes_notas.append(n)
            mun = html.escape(str(r.get('MUNICIPIO', '')))
            o = html.escape(str(r.get('ORIGEM_BASE', '')))
            s = html.escape(str(r.get('SITUACAO SAP', '')))
            s_sisco = html.escape(_valor_alias(r, ['STATUS SISCO', 'STATUS_SISCO']))
            s_list = html.escape(_valor_alias(r, ['STATUS LIST', 'STATUS_LIST']))
            col = html.escape(str(r.get('COLABORADORES MAIS PROXIMOS', '')))
            dup = html.escape(str(r.get('DUPLICADA', '')))
            dup_geo = html.escape(str(r.get('CLASSIFICACAO_DUPLICIDADE_GEO', '')))
            dist_dup = r.get('DISTANCIA_ENTRE_DUPLICATAS_KM')
            dist_dup_txt = f"{float(dist_dup):.3f} km" if pd.notna(dist_dup) else '-'
            mun_div = html.escape(str(r.get('MUNICIPIO_DIVERGENTE', '')))
            alerta_eq = html.escape(str(r.get('ALERTA_EQUIPE_DISTANTE', 'NÃO')))
            classificacao = html.escape(str(r.get('COR_NOME', '-')))
            motivo_inval = html.escape(str(r.get('MOTIVO_INVALIDADE', '-')))
            nota_valida_fluxo = str(r.get('NOTA_VALIDA_FLUXO', 'SIM')).upper()

            link_dup = str(r.get('LINK_DUPLICATA_MAPS', '')).strip()
            origem_dup = html.escape(str(r.get('DUPLICATA_ORIGEM_DESTINO', '')))
            mun_dup = html.escape(str(r.get('DUPLICATA_MUNICIPIO_DESTINO', '')))
            nota_dup = html.escape(str(r.get('DUPLICATA_NOTA_DESTINO', '')))
            link_dup_html = ''
            if dup == 'SIM' and dup_geo == 'LOCAIS DIFERENTES' and link_dup:
                link_dup_safe = html.escape(link_dup, quote=True)
                link_dup_html = (
                    "<div style='margin:6px 0;padding:7px;background:#fff3f3;border-left:3px solid #d32f2f;'>"
                    f"<b>Outra ocorrência:</b> {origem_dup or '-'} | {mun_dup or '-'}<br>"
                    f"<b>Nota:</b> {nota_dup or n}<br>"
                    f"<a href='{link_dup_safe}' target='_blank' style='color:#0D47A1;font-weight:bold;text-decoration:none;'>📍 Abrir outra ocorrência no Google Maps</a>"
                    "</div>"
                )

            motivo_html = ''
            if nota_valida_fluxo == 'NÃO':
                motivo_html = f"<tr><td style='padding:2px;color:#b71c1c;'><b>Motivo da invalidez:</b></td><td style='padding:2px;color:#b71c1c;font-weight:bold;'>{motivo_inval}</td></tr>"

            desc += f'''
            <table style="width:100%; border-collapse:collapse; margin-bottom:8px;">
                <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                <tr><td style="padding:2px;"><b>Classificação:</b></td><td style="padding:2px;font-weight:bold;">{classificacao}</td></tr>
                <tr><td style="padding:2px;"><b>Município:</b></td><td style="padding:2px;">{mun}</td></tr>
                <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                <tr><td style="padding:2px;"><b>Status SAP:</b></td><td style="padding:2px;">{s}</td></tr>
                <tr><td style="padding:2px;"><b>SISCO / LIST:</b></td><td style="padding:2px;">{s_sisco} / {s_list}</td></tr>
                {motivo_html}
                <tr><td style="padding:2px;"><b>Equipes Perto:</b></td><td style="padding:2px;">{col}</td></tr>
                <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup} {dup_geo}</td></tr>
                <tr><td style="padding:2px;"><b>Dist. duplicata:</b></td><td style="padding:2px;">{dist_dup_txt}</td></tr>
                <tr><td style="padding:2px;"><b>Município divergente:</b></td><td style="padding:2px;">{mun_div or '-'}</td></tr>
                <tr><td style="padding:2px;"><b>Equipe distante:</b></td><td style="padding:2px;">{alerta_eq}</td></tr>
            </table>
            {link_dup_html}
            <hr style="margin:4px 0; border:0; border-top:1px solid #ddd;">
            '''

        desc += '</div></div>]]>'

        # Mantém a nota no nome quando houver uma única obra. Em clusters grandes,
        # usa um nome curto e deixa a relação completa dentro do card.
        if len(grp) == 1:
            nome_placemark = nomes_notas[0] if nomes_notas else 'Obra'
        else:
            nome_placemark = f"{len(grp)} obras no local"

        clusters.append({
            'COR_NOME': c_nome,
            'COR_MAPA': c_mapa,
            'LAT': lat,
            'LON': lon,
            'DESC': desc,
            'NOME': nome_placemark
        })

    df_clusters = pd.DataFrame(clusters)

    if not df_clusters.empty:
        for nome_grupo, df_grupo in df_clusters.groupby('COR_NOME', sort=True):
            kml.append(f'<Folder><name>{html.escape(str(nome_grupo))}</name>')
            for _, r in df_grupo.iterrows():
                kml.append(
                    f'<Placemark><name>{html.escape(str(r["NOME"]))}</name>'
                    f'<styleUrl>#style_{r["COR_MAPA"]}</styleUrl>'
                    f'<description>{r["DESC"]}</description>'
                    f'<Point><coordinates>{r["LON"]},{r["LAT"]},0</coordinates></Point></Placemark>'
                )
            kml.append('</Folder>')

    kml.append('</Document></kml>')
    return "\n".join(kml)


def injetar_logo():
    if os.path.exists("LOGO_NIP.png"):
        st.logo("LOGO_NIP.png", icon_image=None)
