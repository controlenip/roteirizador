
"""Análise cruzada de saneamento e levantamento, com alocação geográfica de equipes.
Executar: streamlit run Analise_Cruzada_Atualizada.py
Requisitos: streamlit, pandas, numpy, openpyxl, folium, streamlit-folium.
"""
import io
import hashlib
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Análise de Obras Pendentes', page_icon='📍', layout='wide')

CONTRATOS_VALIDOS = {'0', 'NIP GLOBAL LTDA - EQTL MARANHAO'}
STATUS_LIST_SEM_CONTAGEM = {'EM LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO'}
STATUS_SAP_BLOQUEADOS = {'FINL', 'CANC'}
CATEGORIAS = ('OBRA SANEAMENTO', 'OBRA LEVANTAMENTO', 'OBRA SANEAMENTO E LEVANTAMENTO')


def norm(v):
    if pd.isna(v):
        return ''
    t = unicodedata.normalize('NFKD', str(v)).encode('ascii', 'ignore').decode().upper().strip()
    return re.sub(r'\s+', ' ', t)


def status(v):
    t = norm(v)
    return re.sub(r'\.0+$', '', t) if re.fullmatch(r'\d+\.0+', t) else t


def chave(v):
    t = str(v).strip() if pd.notna(v) else ''
    if re.fullmatch(r'\d+\.0+', t):
        t = t.split('.')[0]
    return '' if norm(t) in {'', 'NAN', 'NONE', 'NULL', '0'} else t


def ler_excel(arquivo, aba=None):
    if arquivo.name.lower().endswith('.csv'):
        for enc in ['utf-8-sig', 'latin-1']:
            try:
                return pd.read_csv(io.BytesIO(arquivo.getvalue()), sep=None, engine='python', encoding=enc, dtype=str)
            except (UnicodeError, pd.errors.ParserError):
                pass
        raise ValueError(f'Não foi possível ler o CSV {arquivo.name}.')
    return pd.read_excel(io.BytesIO(arquivo.getvalue()), sheet_name=aba if aba else 0, dtype=str, engine='openpyxl')


def achar(df, nomes, obrigatorio=True):
    indice = {norm(c): c for c in df.columns}
    for nome in nomes:
        if norm(nome) in indice:
            return indice[norm(nome)]
    if obrigatorio:
        raise ValueError('Coluna obrigatória não encontrada: ' + ' / '.join(nomes))
    return None


def coordenadas(df, lat, lon):
    def parse(s):
        return pd.to_numeric(s.astype(str).str.replace(',', '.', regex=False), errors='coerce')
    la, lo = parse(df[lat]), parse(df[lon])
    # Evita interpretar 0 como coordenada válida.
    valido = la.between(-35, 6) & lo.between(-75, -30) & la.ne(0) & lo.ne(0)
    return la.where(valido), lo.where(valido)


def ler_base_san(arquivo):
    df = ler_excel(arquivo)
    nota = achar(df, ['NOTA'])
    municipio = achar(df, ['MUNICÍPIO', 'MUNICIPIO', 'CIDADE'])
    lat = achar(df, ['LATITUDE PROJETO', 'LATITUDE', 'LATITUDE CAMPO'])
    lon = achar(df, ['LONGITUDE PROJETO', 'LONGITUDE', 'LONGITUDE CAMPO'])
    df['CHAVE_NOTA'] = df[nota].map(chave)
    df['MUNICIPIO_OBRA'] = df[municipio].fillna('').astype(str).str.strip()
    df['LAT_OBRA'], df['LON_OBRA'] = coordenadas(df, lat, lon)
    return df


def ler_base_lev(arquivo):
    df = ler_excel(arquivo, 'NOTAS')
    requeridas = {
        'PROTOCOLO': achar(df, ['PROTOCOLO']),
        'STATUS SAP': achar(df, ['STATUS SAP', 'STATUS_SAP']),
        'STATUS LIST': achar(df, ['STATUS LIST', 'STATUS_LIST']),
        'CONTRATO': achar(df, ['CONTRATO']),
        'MUNICIPIO': achar(df, ['MUNICIPIO', 'MUNICÍPIO']),
        'LATITUDE': achar(df, ['LATITUDE']),
        'LONGITUDE': achar(df, ['LONGITUDE']),
    }
    if len(df.columns) < 14:
        raise ValueError('A aba NOTAS não contém as colunas M e N esperadas.')
    # M e N são verificadas tanto por posição quanto pelo título da base recebida.
    m, n = df.columns[12], df.columns[13]
    if not ('ORCAMENTO MODULAR' in norm(m) and 'PLA ALVOS' in norm(n)):
        raise ValueError(f'Colunas M e N inesperadas: {m!r} / {n!r}. Confira a estrutura da aba NOTAS.')
    df['CHAVE_NOTA'] = df[requeridas['PROTOCOLO']].map(chave)
    df['MUNICIPIO_OBRA'] = df[requeridas['MUNICIPIO']].fillna('').astype(str).str.strip()
    df['LAT_OBRA'], df['LON_OBRA'] = coordenadas(df, requeridas['LATITUDE'], requeridas['LONGITUDE'])
    df['SAP_NORM'] = df[requeridas['STATUS SAP']].map(status)
    df['LIST_NORM'] = df[requeridas['STATUS LIST']].map(status)
    df['CONTRATO_NORM'] = df[requeridas['CONTRATO']].map(status)
    df['COLUNA_M_NORM'] = df[m].map(status)
    df['COLUNA_N_NORM'] = df[n].map(status)
    df['MOTIVOS_EXCLUSAO'] = df.apply(motivos_lev, axis=1)
    df['APTA_CONTAGEM'] = df['MOTIVOS_EXCLUSAO'].eq('')
    return df


def motivos_lev(r):
    motivos = []
    if r['SAP_NORM'] in STATUS_SAP_BLOQUEADOS:
        motivos.append('SAP ' + r['SAP_NORM'])
    if r['LIST_NORM'] in STATUS_LIST_SEM_CONTAGEM:
        motivos.append('LIST ' + r['LIST_NORM'] + ' (sinalizada, fora da contagem)')
    elif r['LIST_NORM'] != '0':
        motivos.append('LIST diferente de 0: ' + (r['LIST_NORM'] or 'vazio'))
    if r['CONTRATO_NORM'] not in CONTRATOS_VALIDOS:
        motivos.append('Contrato fora dos permitidos: ' + (r['CONTRATO_NORM'] or 'vazio'))
    if r['COLUNA_M_NORM'] != '0':
        motivos.append('Coluna M diferente de 0: ' + (r['COLUNA_M_NORM'] or 'vazio'))
    if r['COLUNA_N_NORM'] != '0':
        motivos.append('Coluna N diferente de 0: ' + (r['COLUNA_N_NORM'] or 'vazio'))
    return ' | '.join(motivos)


def ler_equipes(arquivo):
    if arquivo.name.lower().endswith('.csv'):
        raise ValueError('O arquivo de localidades precisa ser XLSX, com abas LEVANTADORES, SANEAMENTO e FISCALIZAÇÃO.')
    xls = pd.ExcelFile(io.BytesIO(arquivo.getvalue()), engine='openpyxl')
    partes = []
    for tipo, nome_aba in [('LEVANTAMENTO', 'LEVANTADORES'), ('SANEAMENTO', 'SANEAMENTO'), ('FISCALIZAÇÃO', 'FISCALIZACAO')]:
        nome = next((x for x in xls.sheet_names if norm(x) == nome_aba), None)
        if nome is None:
            raise ValueError(f'Aba obrigatória não localizada no arquivo de localidades: {nome_aba}')
        d = pd.read_excel(xls, sheet_name=nome, dtype=str)
        col_nome = achar(d, ['NOME', 'NOME_COLAB', 'EQUIPE'])
        col_cidade = achar(d, ['CIDADES', 'CIDADE', 'MUNICIPIO'])
        la = achar(d, ['LATITUDE', 'LAT'])
        lo = achar(d, ['LONGITUDE', 'LON'])
        y, x = coordenadas(d, la, lo)
        partes.append(pd.DataFrame({'EQUIPE': d[col_nome], 'CIDADE_BASE': d[col_cidade], 'LAT_EQUIPE': y,
                                     'LON_EQUIPE': x, 'TIPO_EQUIPE': tipo}))
    result = pd.concat(partes, ignore_index=True).dropna(subset=['EQUIPE', 'LAT_EQUIPE', 'LON_EQUIPE'])
    result = result[result['EQUIPE'].astype(str).str.strip().ne('')].copy()
    if result.empty:
        raise ValueError('O arquivo de localidades não contém equipes com coordenadas válidas.')
    return result


def primeiro(grupo):
    # Prefere uma ocorrência com coordenadas completas, preservando os dados para auditoria.
    indice = grupo['LAT_OBRA'].notna() & grupo['LON_OBRA'].notna()
    return grupo.loc[indice].iloc[0] if indice.any() else grupo.iloc[0]


def consolidar(san, lev):
    san = san[san['CHAVE_NOTA'] != ''].copy()
    lev = lev[lev['CHAVE_NOTA'] != ''].copy()
    san_por_nota = {k: primeiro(g) for k, g in san.groupby('CHAVE_NOTA', sort=False)}
    lev_por_nota = {k: primeiro(g) for k, g in lev.groupby('CHAVE_NOTA', sort=False)}
    lev_grupos = {k: g for k, g in lev.groupby('CHAVE_NOTA', sort=False)}
    san_counts = san['CHAVE_NOTA'].value_counts().to_dict()
    linhas = []
    for nota in dict.fromkeys(list(san_por_nota) + list(lev_por_nota)):
        a, b = san_por_nota.get(nota), lev_por_nota.get(nota)
        dupla = a is not None and b is not None
        categoria = CATEGORIAS[2] if dupla else CATEGORIAS[0] if a is not None else CATEGORIAS[1]
        # Se qualquer ocorrência do levantamento tiver bloqueio, a nota toda fica fora da contagem.
        bloqueios = []
        if b is not None:
            bloqueios = [x for x in lev_grupos[nota]['MOTIVOS_EXCLUSAO'].tolist() if x]
        origem = b if b is not None and pd.notna(b['LAT_OBRA']) else a if a is not None else b
        if origem is None:
            continue
        motivo = ' | '.join(dict.fromkeys(' | '.join(bloqueios).split(' | '))) if bloqueios else ''
        info = {
            'NOTA': nota, 'LISTA': categoria, 'DUPLICADA': 'SIM' if dupla else 'NÃO',
            'PENDENTE_CONTAGEM': 'NÃO' if bloqueios else 'SIM',
            'MOTIVO_EXCLUSAO': motivo or '-', 'MUNICIPIO': origem.get('MUNICIPIO_OBRA', ''),
            'LATITUDE': origem.get('LAT_OBRA'), 'LONGITUDE': origem.get('LON_OBRA'),
            'STATUS_SAP': b['SAP_NORM'] if b is not None else 'SEM REGISTRO LEVANTAMENTO',
            'STATUS_LIST': b['LIST_NORM'] if b is not None else 'SEM REGISTRO LEVANTAMENTO',
            'CONTRATO': b['CONTRATO_NORM'] if b is not None else 'SEM REGISTRO LEVANTAMENTO',
            'COLUNA_M': b['COLUNA_M_NORM'] if b is not None else '-',
            'COLUNA_N': b['COLUNA_N_NORM'] if b is not None else '-',
            'OCORRENCIAS_SANEAMENTO': int(san_counts.get(nota, 0)),
            'OCORRENCIAS_LEVANTAMENTO': int(len(lev_grupos[nota])) if b is not None else 0,
        }
        if a is not None and b is not None:
            info['MUNICIPIO_SANEAMENTO'] = a['MUNICIPIO_OBRA']
            info['MUNICIPIO_LEVANTAMENTO'] = b['MUNICIPIO_OBRA']
        linhas.append(info)
    return pd.DataFrame(linhas)


def haversine(lat, lon, lat2, lon2):
    r1 = np.radians(lat)
    r2 = np.radians(lat2)
    delta_lat = r2 - r1
    delta_lon = np.radians(lon2 - lon)
    a = np.sin(delta_lat / 2)**2 + np.cos(r1)*np.cos(r2)*np.sin(delta_lon / 2)**2
    return 6371.0088 * 2 * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


def adicionar_equipes(df, equipes, limite=200, quantidade=3):
    """Apura apenas equipes habilitadas ao tipo de trabalho, com limite estrito configurável."""
    saida = df.copy()
    equipes_por_tipo = {
        'SANEAMENTO': equipes[equipes.TIPO_EQUIPE == 'SANEAMENTO'],
        'LEVANTAMENTO': equipes[equipes.TIPO_EQUIPE == 'LEVANTAMENTO']
    }
    registros = []
    for r in saida.itertuples(index=False):
        cat = r.LISTA
        tipos = ['SANEAMENTO', 'LEVANTAMENTO'] if cat == CATEGORIAS[2] else ['SANEAMENTO'] if cat == CATEGORIAS[0] else ['LEVANTAMENTO']
        detalhe = {}
        for tipo in tipos:
            prefixo = tipo
            candidatos = equipes_por_tipo[tipo]
            if pd.isna(r.LATITUDE) or pd.isna(r.LONGITUDE) or candidatos.empty:
                proximas = candidatos.iloc[:0].copy()
                menor = np.nan
            else:
                d = haversine(float(r.LATITUDE), float(r.LONGITUDE), candidatos.LAT_EQUIPE.to_numpy(float), candidatos.LON_EQUIPE.to_numpy(float))
                menor = float(np.min(d)) if len(d) else np.nan
                proximas = candidatos.assign(DISTANCIA_KM=d).loc[lambda q: q.DISTANCIA_KM <= limite].sort_values('DISTANCIA_KM').head(quantidade)
            detalhe[f'DISTANCIA_MINIMA_{prefixo}_KM'] = round(menor, 2) if pd.notna(menor) else np.nan
            detalhe[f'{prefixo}_SEM_EQUIPE_NO_RAIO'] = 'SIM' if proximas.empty else 'NÃO'
            detalhe[f'EQUIPES_{prefixo}'] = ' | '.join(f'{x.EQUIPE} ({x.CIDADE_BASE}) - {x.DISTANCIA_KM:.1f} km' for x in proximas.itertuples(index=False)) if not proximas.empty else 'SEM EQUIPE NO RAIO'
            for i in range(quantidade):
                prefix = f'{prefixo}_EQUIPE_{i+1}'
                if i < len(proximas):
                    x = proximas.iloc[i]
                    detalhe[prefix] = x['EQUIPE']
                    detalhe[prefix + '_CIDADE'] = x['CIDADE_BASE']
                    detalhe[prefix + '_KM'] = round(float(x['DISTANCIA_KM']), 2)
                else:
                    detalhe[prefix] = ''
                    detalhe[prefix + '_CIDADE'] = ''
                    detalhe[prefix + '_KM'] = np.nan
        registros.append(detalhe)
    return pd.concat([saida.reset_index(drop=True), pd.DataFrame(registros)], axis=1)


def excel_bytes(ativos, auditoria, san, lev, equipes, distancia):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        contagens = ativos['LISTA'].value_counts() if not ativos.empty else pd.Series(dtype=int)
        resumo = pd.DataFrame([
            ['Obras Saneamento', int(contagens.get(CATEGORIAS[0], 0))],
            ['Obras Levantamento', int(contagens.get(CATEGORIAS[1], 0))],
            ['Obras Saneamento e Levantamento', int(contagens.get(CATEGORIAS[2], 0))],
            ['Total único pendente', len(ativos)],
            ['Sinalizadas / excluídas', int((auditoria['PENDENTE_CONTAGEM'] == 'NÃO').sum())],
            ['Distância máxima (km)', distancia],
            ['Gerado em', datetime.now().strftime('%d/%m/%Y %H:%M')]
        ], columns=['INDICADOR', 'VALOR'])
        resumo.to_excel(writer, sheet_name='RESUMO', index=False)
        for cat, aba in zip(CATEGORIAS, ['OBRAS SANEAMENTO', 'OBRAS LEVANTAMENTO', 'OBRAS DUPLICADAS']):
            ativos.loc[ativos.LISTA.eq(cat)].to_excel(writer, sheet_name=aba, index=False)
        auditoria.loc[auditoria.PENDENTE_CONTAGEM.eq('NÃO')].to_excel(writer, sheet_name='SINALIZADAS EXCLUIDAS', index=False)
        auditoria.to_excel(writer, sheet_name='AUDITORIA COMPLETA', index=False)
        equipes.to_excel(writer, sheet_name='EQUIPES LOCALIDADES', index=False)
        for ws in writer.book:
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import Font, PatternFill
            for cell in ws[1]:
                cell.fill = PatternFill('solid', fgColor='14366F')
                cell.font = Font(bold=True, color='FFFFFF')
            for column in ws.columns:
                letter = column[0].column_letter
                largura = min(55, max(13, max(len(str(c.value or '')) for c in list(column)[:150]) + 2))
                ws.column_dimensions[letter].width = largura
    return buf.getvalue()


st.title('📍 Análise Cruzada — Obras Pendentes e Equipes Próximas')
st.caption('Uma nota = uma obra na volumetria. Dados de SAP, LIST, contrato e colunas M/N são validados na aba NOTAS da base de Levantamento.')

with st.expander('📘 Regras de contagem e critérios', expanded=False):
    st.markdown('''A nota entra na contagem quando **não** houver bloqueios registrados na base Levantamento. Para notas presentes em Levantamento são obrigatórios: SAP diferente de FINL/CANC, LIST exatamente 0, contrato 0 ou NIP GLOBAL LTDA - EQTL MARANHÃO, M=0 e N=0. "Em levantamento" e "Correção de levantamento" ficam sinalizadas e fora da contagem. Notas exclusivas de Saneamento permanecem pendentes por não haver evidência de bloqueio em Levantamento. O raio é calculado em linha reta (Haversine) entre a obra e a base da equipe, não é distância rodoviária.''')

c1, c2, c3 = st.columns(3)
with c1:
    san_file = st.file_uploader('1. BASE_SANEAMENTO', type=['xlsx', 'csv'], key='san_upload')
with c2:
    lev_file = st.file_uploader('2. BASE_LEVANTAMENTO_ATUALIZADA', type=['xlsx', 'csv'], key='lev_upload')
with c3:
    loc_file = st.file_uploader('3. LOCALIDADE LEVANTADORES-SANEAMENTO-FISCALIZAÇÃO', type=['xlsx'], key='loc_upload')

if not (san_file and lev_file and loc_file):
    st.warning('Envie os três arquivos obrigatórios para habilitar o processamento.')
    st.stop()

quantidade = st.selectbox('Equipes próximas por atividade', [1, 2, 3, 4, 5], index=2)
if st.button('🚀 Processar três bases', type='primary'):
    try:
        with st.spinner('Cruzando notas, auditando os status e calculando distâncias...'):
            san = ler_base_san(san_file)
            lev = ler_base_lev(lev_file)
            equipes = ler_equipes(loc_file)
            consolidado = consolidar(san, lev)
            st.session_state['dados_analise'] = (san, lev, equipes, consolidado)
            st.session_state['qtd_equipes'] = quantidade
            st.session_state['assinatura_arquivos'] = tuple(hashlib.sha256(f.getvalue()).hexdigest() for f in (san_file, lev_file, loc_file))
        st.success('Bases validadas e cruzadas. Selecione o raio para ver as equipes habilitadas.')
    except Exception as e:
        st.error(f'Não foi possível processar: {e}')
        st.stop()

if 'dados_analise' not in st.session_state:
    st.stop()

san, lev, equipes, auditoria = st.session_state['dados_analise']
if st.session_state.get('assinatura_arquivos') != tuple(hashlib.sha256(f.getvalue()).hexdigest() for f in (san_file, lev_file, loc_file)):
    st.warning('Os arquivos foram alterados. Clique em Processar três bases novamente para atualizar os resultados.')
    st.stop()

ativos = auditoria.loc[auditoria.PENDENTE_CONTAGEM.eq('SIM')].copy()

raio = st.select_slider('📏 Distância máxima da equipe até a obra (km)', options=list(range(200, 501, 25)), value=200)
st.caption('Sem equipe até 200 km? Aumente gradualmente para 225, 250, 275, 300 ... até 500 km. Apenas equipes dentro do limite escolhido são mostradas.')
ativos = adicionar_equipes(ativos, equipes, raio, quantidade)
for cat in CATEGORIAS:
    cont = int(ativos.LISTA.eq(cat).sum())
    st.session_state['cnt_' + cat] = cont

met = st.columns(5)
met[0].metric('Saneamento', st.session_state['cnt_' + CATEGORIAS[0]])
met[1].metric('Levantamento', st.session_state['cnt_' + CATEGORIAS[1]])
met[2].metric('Duplicadas (ambas)', st.session_state['cnt_' + CATEGORIAS[2]])
met[3].metric('Total único pendente', len(ativos))
met[4].metric('Sinalizadas fora da conta', int(auditoria.PENDENTE_CONTAGEM.eq('NÃO').sum()))

municipios = sorted(x for x in ativos.MUNICIPIO.dropna().unique() if str(x).strip())
filtra_municipio = st.multiselect('Filtrar municípios (opcional)', municipios)
nota_busca = st.text_input('Pesquisar nota (opcional)').strip()
view = ativos.copy()
if filtra_municipio:
    view = view.loc[view.MUNICIPIO.isin(filtra_municipio)]
if nota_busca:
    view = view.loc[view.NOTA.astype(str).str.contains(re.escape(nota_busca), case=False, na=False)]

abas = st.tabs(['🟣 OBRAS SANEAMENTO', '🟢 OBRAS LEVANTAMENTO', '🔵 SANEAMENTO + LEVANTAMENTO', '⚠️ SINALIZADAS / EXCLUÍDAS', '🗺️ MAPA'])
for aba, categoria in zip(abas[:3], CATEGORIAS):
    with aba:
        tabela = view.loc[view.LISTA.eq(categoria)].copy()
        st.subheader(f'{categoria} — {len(tabela)} obras')
        if categoria == CATEGORIAS[2]:
            sem = tabela.loc[tabela['SANEAMENTO_SEM_EQUIPE_NO_RAIO'].eq('SIM') | tabela['LEVANTAMENTO_SEM_EQUIPE_NO_RAIO'].eq('SIM')]
        else:
            tipo = 'SANEAMENTO' if categoria == CATEGORIAS[0] else 'LEVANTAMENTO'
            sem = tabela.loc[tabela[f'{tipo}_SEM_EQUIPE_NO_RAIO'].eq('SIM')]
        if not sem.empty:
            st.warning(f'{len(sem)} obra(s) sem equipe de uma ou mais atividades no raio de {raio} km. Aumente o seletor de distância para consultar alternativas.')
        st.dataframe(tabela, hide_index=True, use_container_width=True)

with abas[3]:
    bloqueadas = auditoria.loc[auditoria.PENDENTE_CONTAGEM.eq('NÃO')].copy()
    st.subheader(f'Notas sinalizadas, fora da contagem — {len(bloqueadas)}')
    st.dataframe(bloqueadas, hide_index=True, use_container_width=True)
    st.caption('Motivos completos são mantidos por nota, inclusive quando existem várias linhas na mesma base.')

with abas[4]:
    st.subheader('Mapa das obras pendentes')
    if st.checkbox('Carregar mapa interativo'):
        try:
            import folium
            from folium.plugins import MarkerCluster
            from streamlit_folium import st_folium
            geos = view.dropna(subset=['LATITUDE', 'LONGITUDE'])
            centro = [float(geos.LATITUDE.mean()), float(geos.LONGITUDE.mean())] if len(geos) else [-5.0, -45.0]
            mapa = folium.Map(location=centro, zoom_start=7)
            paleta = dict(zip(CATEGORIAS, ['purple', 'green', 'blue']))
            for categoria in CATEGORIAS:
                camada = folium.FeatureGroup(name=categoria)
                cluster = MarkerCluster().add_to(camada)
                for r in geos.loc[geos.LISTA.eq(categoria)].itertuples(index=False):
                    conteudo = f'<b>Nota:</b> {r.NOTA}<br><b>Município:</b> {r.MUNICIPIO}<br><b>Tipo:</b> {categoria}'
                    for tipo in (['SANEAMENTO', 'LEVANTAMENTO'] if categoria == CATEGORIAS[2] else ['SANEAMENTO'] if categoria == CATEGORIAS[0] else ['LEVANTAMENTO']):
                        conteudo += f'<br><b>Equipes {tipo}:</b> {getattr(r, "EQUIPES_" + tipo, "")}'
                    folium.Marker([r.LATITUDE, r.LONGITUDE], popup=folium.Popup(conteudo, max_width=450), icon=folium.Icon(color=paleta[categoria])).add_to(cluster)
                camada.add_to(mapa)
            folium.LayerControl().add_to(mapa)
            st_folium(mapa, use_container_width=True, height=550)
        except ImportError:
            st.warning('Para usar o mapa, instale folium e streamlit-folium.')

st.download_button('📥 Baixar Excel com 3 listas, auditoria e equipes', data=excel_bytes(ativos, auditoria, san, lev, equipes, raio), file_name=f'Obras_Pendentes_{datetime.now():%Y%m%d_%H%M}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption(f'Base de origem: {len(san):,} linhas de saneamento e {len(lev):,} linhas de levantamento. O total reflete notas únicas, não linhas repetidas.')
