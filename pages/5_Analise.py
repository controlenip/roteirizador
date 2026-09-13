import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
import unicodedata
import uuid
import traceback
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, fundir_super_pontos
from modules.export_analise import gerar_excel_analise, gerar_kml_analise

st.set_page_config(page_title="Análise Cruzada", page_icon="🔍", layout="wide")

st.markdown("""
<style>
    .stMultiSelect [data-baseweb="select"] > div:first-child { flex-wrap: wrap !important; }
    .stMultiSelect [data-baseweb="select"] > div:first-child > div:last-child { display: none !important; }
    .stMultiSelect [data-baseweb="tag"] { max-width: 100% !important; }
</style>
""", unsafe_allow_html=True)

try:
    from modules.export_analise import injetar_logo
    injetar_logo()
except Exception:
    pass


# ==============================================================
# FUNÇÕES DE APOIO
# ==============================================================
def corrigir_coord(val, limite):
    if pd.isna(val):
        return np.nan
    try:
        v = float(val)
    except Exception:
        return np.nan
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v


def remover_acentos_str(texto):
    if not isinstance(texto, str):
        return str(texto)
    return "".join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))


def _norm_col(c):
    return re.sub(r'\s+', ' ', remover_acentos_str(str(c)).upper().strip())


def encontrar_coluna(df, candidatos_exatos, contem=None):
    """Localiza uma coluna com prioridade determinística, evitando renomear a coluna errada."""
    if df is None or df.empty and len(df.columns) == 0:
        return None
    mapa = {_norm_col(c): c for c in df.columns}
    for cand in candidatos_exatos:
        chave = _norm_col(cand)
        if chave in mapa:
            return mapa[chave]
    if contem:
        tokens = [_norm_col(x) for x in contem]
        for c in df.columns:
            nc = _norm_col(c)
            if all(t in nc for t in tokens):
                return c
    return None


def renomear_seguro(df, origem, destino):
    if origem is None or origem == destino:
        return df
    if destino in df.columns:
        return df
    return df.rename(columns={origem: destino})


def valor_alias(row, aliases, default='-'):
    # Primeiro tenta os nomes exatos; depois compara nomes normalizados para tolerar
    # espaços duplicados, acentos e pequenas variações vindas das planilhas.
    for col in aliases:
        try:
            val = row.get(col)
        except Exception:
            val = None
        if pd.notna(val) and str(val).strip().lower() not in ['', 'nan', 'none']:
            return str(val).strip()
    try:
        mapa = {_norm_col(c): c for c in row.index}
        for alias in aliases:
            real = mapa.get(_norm_col(alias))
            if real is not None:
                val = row.get(real)
                if pd.notna(val) and str(val).strip().lower() not in ['', 'nan', 'none']:
                    return str(val).strip()
    except Exception:
        pass
    return default


def nota_valida(v):
    s = str(v).strip()
    return s.upper() not in ['', 'NAN', 'NONE', 'NULL', '-']


def limpar_nota_serie(s):
    # Remove apenas o sufixo decimal criado pelo Excel (ex.: 12345.0 -> 12345).
    # Não altera identificadores legítimos como 123.01.
    return s.astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)


def normalizar_status_fluxo(valor, default='0'):
    """Normaliza status de SAP/SISCO/LIST sem depender de acentos, pontuação ou espaços duplicados."""
    if pd.isna(valor):
        return default
    s = remover_acentos_str(str(valor)).upper().strip()
    if s in ['', 'NAN', 'NONE', 'NULL', '-']:
        return default
    # Remove apenas sufixo decimal artificial de valores numéricos (0.0 -> 0).
    s = re.sub(r'(?<=\d)\.0+$', '', s)
    # Torna equivalentes formas como PRÉ-ANÁLISE, PRE_ANALISE, PRE  ANALISE e P/.
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s or default


# Versão explícita das regras usadas na auditoria. É gravada nas exportações para rastreabilidade.
VERSAO_REGRAS_ANALISE = '2026.09'
REGRAS_STATUS_ANALISE = {
    'STATUS_LIST_VALIDOS': ('0', 'EM LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO'),
    'STATUS_SISCO_VALIDOS': ('0', 'PRE ANALISE', 'LIBERADO PARA LEVANTAMENTO', 'LIBERADO P LEVANTAMENTO'),
}
STATUS_LIST_VALIDOS = set(REGRAS_STATUS_ANALISE['STATUS_LIST_VALIDOS'])
STATUS_SISCO_VALIDOS = set(REGRAS_STATUS_ANALISE['STATUS_SISCO_VALIDOS'])


def _coluna_status_disponivel(linha, aliases, flag_coluna):
    """Distingue coluna inexistente de valor 0 real, inclusive após concatenação das bases."""
    flag = linha.get(flag_coluna, None)
    if pd.notna(flag) and str(flag).strip().upper() in {'SIM', 'NÃO', 'NAO'}:
        return str(flag).strip().upper() == 'SIM'
    try:
        nomes = {_norm_col(c) for c in linha.index}
        return any(_norm_col(a) in nomes for a in aliases)
    except Exception:
        return False


def avaliar_validade_fluxo(linha):
    """Retorna validade e motivo usando ausência de status como ausência, nunca como 0 implícito."""
    motivos = []
    situacao_sap = str(linha.get('SITUACAO SAP', '')).strip().upper()
    if 'BLOQUEADO' in situacao_sap or 'FINL' in situacao_sap or 'CANC' in situacao_sap:
        motivos.append(f"Status SAP bloqueado: {situacao_sap or '-'}")

    origem = str(linha.get('ORIGEM_BASE', '')).strip().upper()
    tem_list = _coluna_status_disponivel(linha, ['STATUS_LIST', 'STATUS LIST'], 'COLUNA_STATUS_LIST_LOCALIZADA')
    tem_sisco = _coluna_status_disponivel(linha, ['STATUS_SISCO', 'STATUS SISCO'], 'COLUNA_STATUS_SISCO_LOCALIZADA')

    if tem_list:
        st_list = normalizar_status_fluxo(valor_alias(linha, ['STATUS_LIST', 'STATUS LIST'], 'NAO INFORMADO'), 'NAO INFORMADO')
    else:
        st_list = 'COLUNA NAO LOCALIZADA'
    if tem_sisco:
        st_sisco = normalizar_status_fluxo(valor_alias(linha, ['STATUS_SISCO', 'STATUS SISCO'], 'NAO INFORMADO'), 'NAO INFORMADO')
    else:
        st_sisco = 'COLUNA NAO LOCALIZADA'

    if origem == 'LEVANTAMENTO':
        if not tem_list:
            motivos.append('Coluna STATUS LIST não localizada')
        elif st_list not in STATUS_LIST_VALIDOS:
            motivos.append(f"Status LIST não permitido: {st_list}")
        if not tem_sisco:
            motivos.append('Coluna STATUS SISCO não localizada')
        elif st_sisco not in STATUS_SISCO_VALIDOS:
            motivos.append(f"Status SISCO não permitido: {st_sisco}")

    return {
        'valida': not motivos,
        'motivo': ' | '.join(motivos) if motivos else '-',
        'status_list_normalizado': st_list,
        'status_sisco_normalizado': st_sisco,
    }


def classificar_situacao_sap(nota, status_dict, status_sap_localizado, notas_lev_cadastradas):
    """Nunca transforma ausência de informação SAP em APTO."""
    nota = str(nota).strip()
    if not status_sap_localizado or nota not in notas_lev_cadastradas:
        return 'SEM REGISTRO SAP'
    s_status = str(status_dict.get(nota, '')).strip().upper()
    if not s_status:
        return 'SEM STATUS SAP'
    if 'FINL' in s_status or 'CANC' in s_status:
        return f"BLOQUEADO ({s_status})"
    return 'APTO'


def fonte_validacao_sap(nota, status_sap_localizado, notas_lev_cadastradas):
    nota = str(nota).strip()
    if not status_sap_localizado:
        return 'Coluna de Status SAP não localizada'
    if nota not in notas_lev_cadastradas:
        return 'Nota não encontrada na base Levantamento'
    return 'Base Levantamento'


def determinar_classificacao_analise(linha):
    """Classificação única da análise, usando a validade já auditada como fonte de verdade."""
    if str(linha.get('NOTA_VALIDA_FLUXO', 'SIM')).strip().upper() == 'NÃO':
        return 'black', '⚫ Notas Inválidas'
    if str(linha.get('DUPLICADA', '')).strip().upper() == 'SIM':
        return 'red', '🔴 Notas Duplicadas'
    if str(linha.get('PROXIMA', '')).strip().upper() == 'SIM':
        return 'orange', '🟠 Notas Próximas'
    origem = str(linha.get('ORIGEM_BASE', '')).strip().upper()
    if origem == 'LEVANTAMENTO':
        return 'green', '🟢 Notas Levantamento Solitárias'
    if origem == 'SANEAMENTO':
        return 'purple', '🟣 Notas Saneamento Solitárias'
    return 'blue', '🔵 Outras'


def ler_csv_resiliente(file_bytes):
    """Lê CSVs corporativos com separador/encoding variáveis sem mudar o conteúdo."""
    ultimo_erro = None
    for encoding in ['utf-8-sig', 'utf-8', 'latin-1']:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding=encoding)
        except Exception as exc:
            ultimo_erro = exc
    # Fallback explícito para exportações que usam ponto e vírgula.
    for encoding in ['utf-8-sig', 'latin-1']:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=';', encoding=encoding)
        except Exception as exc:
            ultimo_erro = exc
    raise ultimo_erro if ultimo_erro else ValueError('Não foi possível ler o CSV.')


def criar_id_analise():
    return f"ANL-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"


def montar_config_txt(config, cores=None, filtros=None):
    linhas = [
        "CONFIGURAÇÃO DA ANÁLISE CRUZADA",
        f"ID da análise: {config.get('id_analise', '-')}",
        f"Versão das regras: {config.get('versao_regras', VERSAO_REGRAS_ANALISE)}",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Arquivo Saneamento: {config.get('arquivo_saneamento', '-')}",
        f"Arquivo Levantamento: {config.get('arquivo_levantamento', '-')}",
        f"Arquivo Localidades: {config.get('arquivo_localidades', '-')}",
        f"Raio de agrupamento: {config.get('raio_proximidade_m', '-')} m",
        f"Quantidade de equipes próximas: {config.get('qtd_equipes_proximas', '-')}",
        f"Distância máxima para equipes adicionais: {config.get('distancia_max_equipe_km', '-')} km",
        f"Alerta de equipe distante: {config.get('distancia_alerta_equipe_km', '-')} km",
        f"Linhas Saneamento recebidas: {config.get('linhas_saneamento', '-')}",
        f"Linhas Levantamento recebidas: {config.get('linhas_levantamento', '-')}",
        f"Linhas válidas espacialmente: {config.get('linhas_validas', '-')}",
        f"Coordenadas de obras rejeitadas: {config.get('coordenadas_rejeitadas', '-')}",
        f"Coordenadas de obras corrigidas: {config.get('coordenadas_corrigidas', '-')}",
        f"Localidades rejeitadas: {config.get('localidades_rejeitadas', '-')}",
        f"Localidades corrigidas: {config.get('localidades_corrigidas', '-')}",
        f"Localidades com alertas: {config.get('localidades_alertas', '-')}",
        f"Status SAP localizado: {'SIM' if config.get('status_sap_localizado', False) else 'NÃO'}",
        f"Registros sem confirmação SAP: {config.get('sem_registro_sap', '-')}",
        f"Duplicadas entre bases: {config.get('duplicadas_interbase', '-')}",
        f"Tempo total de processamento: {config.get('tempo_processamento_s', '-')} s",
    ]
    tempos = config.get('tempos_etapas', {}) or {}
    if tempos:
        linhas.append("TEMPOS POR ETAPA:")
        for etapa, segundos in tempos.items():
            linhas.append(f"- {etapa}: {segundos} s")
    if cores is not None:
        linhas.append(f"Filtros de cores exportados: {', '.join(cores)}")
    if filtros:
        linhas.append("FILTROS DE EXPORTAÇÃO:")
        for k, v in filtros.items():
            linhas.append(f"- {k}: {v}")
    return "\n".join(linhas) + "\n"


def preparar_base_obras(df, origem):
    df = df.copy()
    df.columns = normalize_cols(df.columns)
    df = df.loc[:, ~df.columns.duplicated()].copy()

    col_mun = encontrar_coluna(df, ['MUNICIPIO', 'MUNICÍPIO', 'CIDADE'], ['MUNI'])
    df = renomear_seguro(df, col_mun, 'MUNICIPIO')

    # Saneamento prioriza coordenadas de projeto quando existirem, preservando a lógica anterior.
    if origem == 'SANEAMENTO':
        col_lat = encontrar_coluna(
            df,
            ['LATITUDE PROJETO', 'LATITUDE_PROJETO', 'PROJETO LATITUDE', 'LAT PROJETO', 'LATITUDE', 'LAT'],
            ['PROJETO', 'LAT']
        )
        col_lon = encontrar_coluna(
            df,
            ['LONGITUDE PROJETO', 'LONGITUDE_PROJETO', 'PROJETO LONGITUDE', 'LON PROJETO', 'LONGITUDE', 'LON'],
            ['PROJETO', 'LON']
        )
        if col_lat is None:
            col_lat = encontrar_coluna(df, ['LATITUDE', 'LAT'], ['LAT'])
        if col_lon is None:
            col_lon = encontrar_coluna(df, ['LONGITUDE', 'LON'], ['LON'])
    else:
        col_lat = encontrar_coluna(df, ['LATITUDE', 'LAT', 'LATITUDE PROJETO', 'LATITUDE_PROJETO'], ['LAT'])
        col_lon = encontrar_coluna(df, ['LONGITUDE', 'LON', 'LONGITUDE PROJETO', 'LONGITUDE_PROJETO'], ['LON'])

    df = renomear_seguro(df, col_lat, 'LATITUDE')
    df = renomear_seguro(df, col_lon, 'LONGITUDE')

    col_nota = encontrar_coluna(df, ['NOTA', 'PROTOCOLO', 'OS', 'ID_SISCO', 'ID SISCO'])
    df = renomear_seguro(df, col_nota, 'NOTA')

    df['ORIGEM_BASE'] = origem
    if 'NOTA' in df.columns:
        df['NOTA'] = limpar_nota_serie(df['NOTA'])
    return df


def preparar_localidades(file_loc):
    dfs_loc = []

    def tratar_aba(df_temp, tipo):
        df_temp = df_temp.copy()
        df_temp.columns = normalize_cols(df_temp.columns)
        df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()].copy()

        col_nome = encontrar_coluna(
            df_temp,
            ['NOME_COLAB', 'NOME COLAB', 'NOME', 'COLABORADOR', 'LEVANTADOR', 'FISCAL', 'TECNICO', 'TÉCNICO', 'EQUIPE']
        )
        col_lat = encontrar_coluna(df_temp, ['LAT_LOC', 'LAT LOC', 'LATITUDE', 'LAT'], ['LAT'])
        col_lon = encontrar_coluna(df_temp, ['LON_LOC', 'LON LOC', 'LONGITUDE', 'LON'], ['LON'])

        df_temp = renomear_seguro(df_temp, col_nome, 'NOME_COLAB')
        df_temp = renomear_seguro(df_temp, col_lat, 'LAT_LOC')
        df_temp = renomear_seguro(df_temp, col_lon, 'LON_LOC')
        df_temp['TIPO_EQUIPE'] = tipo
        return df_temp

    if file_loc.name.lower().endswith('.csv'):
        df_temp = ler_csv_resiliente(file_loc.getvalue())
        dfs_loc.append(tratar_aba(df_temp, 'Equipe (CSV)'))
    else:
        file_loc_buffer = io.BytesIO(file_loc.getvalue())
        xls_loc = pd.ExcelFile(file_loc_buffer)
        for sheet in xls_loc.sheet_names:
            df_temp = pd.read_excel(xls_loc, sheet_name=sheet)
            tipo = "Saneamento" if "SAN" in sheet.upper() else "Levantamento"
            dfs_loc.append(tratar_aba(df_temp, tipo))

    return pd.concat(dfs_loc, ignore_index=True) if dfs_loc else pd.DataFrame()


def validar_colunas(df, obrigatorias, nome_base):
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        st.error(f"❌ {nome_base}: faltam as colunas obrigatórias: {', '.join(faltantes)}.")
        return False
    return True


def colaboradores_proximos_em_lote(df, lat_locs, lon_locs, nomes_locs, tipos_locs, qtd_equipes, distancia_max_km, distancia_alerta_km=100.0):
    """Calcula uma vez por coordenada única e devolve texto + colunas estruturadas."""
    qtd_equipes = max(1, int(qtd_equipes))
    max_cols = max(5, qtd_equipes)
    if len(lat_locs) == 0:
        base = pd.DataFrame(index=df.index)
        base['COLABORADORES MAIS PROXIMOS'] = 'DESCONHECIDO'
        base['DISTANCIA_EQUIPE_MAIS_PROXIMA_KM'] = np.nan
        base['ALERTA_EQUIPE_DISTANTE'] = 'NÃO'
        return base

    coords_unicas = df[['LATITUDE', 'LONGITUDE']].drop_duplicates().copy()
    cache = {}
    for _, rr in coords_unicas.iterrows():
        lat, lon = float(rr['LATITUDE']), float(rr['LONGITUDE'])
        chave = (lat, lon)
        dists = haversine_vectorized(lat, lon, lat_locs, lon_locs)
        if len(dists) == 0:
            cache[chave] = {'COLABORADORES MAIS PROXIMOS': 'DESCONHECIDO', 'DISTANCIA_EQUIPE_MAIS_PROXIMA_KM': np.nan, 'ALERTA_EQUIPE_DISTANTE': 'NÃO'}
            continue

        k = min(qtd_equipes, len(dists))
        if k == len(dists):
            idxs = np.argsort(dists)
        else:
            idxs = np.argpartition(dists, k - 1)[:k]
            idxs = idxs[np.argsort(dists[idxs])]

        item = {}
        partes = []
        for pos in range(max_cols):
            item[f'EQUIPE_{pos+1}'] = ''
            item[f'TIPO_EQUIPE_{pos+1}'] = ''
            item[f'DISTANCIA_EQUIPE_{pos+1}_KM'] = np.nan

        adicionados = 0
        for pos, i in enumerate(idxs):
            d_km = float(dists[i])
            # A primeira equipe continua sempre sendo exibida. O limite vale só para adicionais.
            if pos >= 1 and d_km > float(distancia_max_km):
                break
            nome = str(nomes_locs[i]).strip()
            tipo = str(tipos_locs[i]).strip()
            item[f'EQUIPE_{pos+1}'] = nome.title()
            item[f'TIPO_EQUIPE_{pos+1}'] = tipo
            item[f'DISTANCIA_EQUIPE_{pos+1}_KM'] = round(d_km, 3)
            partes.append(f"{nome.title()} ({tipo}) - {d_km:.1f}km")
            adicionados += 1

        d1 = float(dists[idxs[0]]) if len(idxs) else np.nan
        item['COLABORADORES MAIS PROXIMOS'] = " | ".join(partes) if partes else 'DESCONHECIDO'
        item['DISTANCIA_EQUIPE_MAIS_PROXIMA_KM'] = round(d1, 3) if pd.notna(d1) else np.nan
        item['ALERTA_EQUIPE_DISTANTE'] = 'SIM' if pd.notna(d1) and d1 > float(distancia_alerta_km) else 'NÃO'
        cache[chave] = item

    registros = []
    for _, rr in df[['LATITUDE', 'LONGITUDE']].iterrows():
        registros.append(cache.get((float(rr['LATITUDE']), float(rr['LONGITUDE'])), {'COLABORADORES MAIS PROXIMOS': 'DESCONHECIDO'}))
    return pd.DataFrame(registros, index=df.index)


def auditar_duplicidades_geograficas(df_valid, duplicadas_inter, raio_metros):
    """Audita duplicidades, conta ocorrências e aponta a contraparte mais próxima na outra base."""
    df_valid = df_valid.copy()
    colunas = {
        'DISTANCIA_ENTRE_DUPLICATAS_KM': np.nan,
        'DUPLICATA_MESMO_LOCAL': '',
        'COORDENADA_DIVERGENTE': '',
        'MUNICIPIO_DIVERGENTE': '',
        'CLASSIFICACAO_DUPLICIDADE_GEO': '',
        'DUPLICATA_NOTA_DESTINO': '',
        'DUPLICATA_ORIGEM_DESTINO': '',
        'DUPLICATA_MUNICIPIO_DESTINO': '',
        'DUPLICATA_LAT_DESTINO': np.nan,
        'DUPLICATA_LON_DESTINO': np.nan,
        'LINK_DUPLICATA_MAPS': '',
        'QTD_OCORRENCIAS_NOTA': 0,
        'QTD_OCORRENCIAS_SANEAMENTO': 0,
        'QTD_OCORRENCIAS_LEVANTAMENTO': 0,
        'QTD_OCORRENCIAS_OUTRA_BASE': 0,
    }
    for c, default in colunas.items():
        df_valid[c] = default

    # Contagens são úteis mesmo quando existem mais de duas ocorrências da mesma NOTA.
    notas_validas_mask = df_valid['NOTA'].apply(nota_valida)
    totais = df_valid.loc[notas_validas_mask].groupby('NOTA').size().to_dict()
    san_counts = df_valid.loc[notas_validas_mask & df_valid['ORIGEM_BASE'].astype(str).str.upper().eq('SANEAMENTO')].groupby('NOTA').size().to_dict()
    lev_counts = df_valid.loc[notas_validas_mask & df_valid['ORIGEM_BASE'].astype(str).str.upper().eq('LEVANTAMENTO')].groupby('NOTA').size().to_dict()
    df_valid['QTD_OCORRENCIAS_NOTA'] = df_valid['NOTA'].map(totais).fillna(0).astype(int)
    df_valid['QTD_OCORRENCIAS_SANEAMENTO'] = df_valid['NOTA'].map(san_counts).fillna(0).astype(int)
    df_valid['QTD_OCORRENCIAS_LEVANTAMENTO'] = df_valid['NOTA'].map(lev_counts).fillna(0).astype(int)

    raio_km = float(raio_metros) / 1000.0

    for nota in duplicadas_inter:
        idxs = df_valid.index[df_valid['NOTA'] == nota].tolist()
        if not idxs:
            continue

        for idx in idxs:
            atual = df_valid.loc[idx]
            origem_atual = str(atual.get('ORIGEM_BASE', '')).strip().upper()
            candidatos = df_valid[(df_valid['NOTA'] == nota) & (df_valid['ORIGEM_BASE'].astype(str).str.upper() != origem_atual)]
            df_valid.at[idx, 'QTD_OCORRENCIAS_OUTRA_BASE'] = int(len(candidatos))
            if candidatos.empty:
                continue

            lat = float(atual['LATITUDE'])
            lon = float(atual['LONGITUDE'])
            dists = haversine_vectorized(
                lat, lon,
                candidatos['LATITUDE'].to_numpy(dtype=float),
                candidatos['LONGITUDE'].to_numpy(dtype=float)
            )
            if len(dists) == 0:
                continue

            pos = int(np.argmin(dists))
            destino = candidatos.iloc[pos]
            dist_km = float(dists[pos])
            mesmo = 'SIM' if dist_km <= raio_km else 'NÃO'
            coord_div = 'NÃO' if mesmo == 'SIM' else 'SIM'
            classe = 'MESMO LOCAL' if mesmo == 'SIM' else 'LOCAIS DIFERENTES'

            mun_atual = normalizar_municipios(pd.Series([str(atual.get('MUNICIPIO', ''))])).iloc[0]
            mun_dest = normalizar_municipios(pd.Series([str(destino.get('MUNICIPIO', ''))])).iloc[0]
            mun_div = 'SIM' if mun_atual and mun_dest and mun_atual != mun_dest else 'NÃO'

            lat_dest = float(destino['LATITUDE'])
            lon_dest = float(destino['LONGITUDE'])
            link = f"https://www.google.com/maps?q={lat_dest:.8f},{lon_dest:.8f}"

            df_valid.at[idx, 'DISTANCIA_ENTRE_DUPLICATAS_KM'] = round(dist_km, 3)
            df_valid.at[idx, 'DUPLICATA_MESMO_LOCAL'] = mesmo
            df_valid.at[idx, 'COORDENADA_DIVERGENTE'] = coord_div
            df_valid.at[idx, 'MUNICIPIO_DIVERGENTE'] = mun_div
            df_valid.at[idx, 'CLASSIFICACAO_DUPLICIDADE_GEO'] = classe
            df_valid.at[idx, 'DUPLICATA_NOTA_DESTINO'] = str(destino.get('NOTA', nota))
            df_valid.at[idx, 'DUPLICATA_ORIGEM_DESTINO'] = str(destino.get('ORIGEM_BASE', ''))
            df_valid.at[idx, 'DUPLICATA_MUNICIPIO_DESTINO'] = str(destino.get('MUNICIPIO', ''))
            df_valid.at[idx, 'DUPLICATA_LAT_DESTINO'] = lat_dest
            df_valid.at[idx, 'DUPLICATA_LON_DESTINO'] = lon_dest
            df_valid.at[idx, 'LINK_DUPLICATA_MAPS'] = link

    return df_valid


def _max_distancia_haversine_km(lats, lons, block_size=512):
    """Maior distância exata do cluster usando NumPy em blocos, sem matriz N x N completa."""
    lats = np.asarray(lats, dtype=float)
    lons = np.asarray(lons, dtype=float)
    n = len(lats)
    if n <= 1:
        return 0.0
    lat_r = np.radians(lats)
    lon_r = np.radians(lons)
    maior = 0.0
    r_terra = 6371.0
    for i0 in range(0, n, block_size):
        i1 = min(n, i0 + block_size)
        la = lat_r[i0:i1][:, None]
        loa = lon_r[i0:i1][:, None]
        for j0 in range(i0, n, block_size):
            j1 = min(n, j0 + block_size)
            lb = lat_r[j0:j1][None, :]
            lob = lon_r[j0:j1][None, :]
            dlat = lb - la
            dlon = lob - loa
            a = np.sin(dlat / 2.0) ** 2 + np.cos(la) * np.cos(lb) * np.sin(dlon / 2.0) ** 2
            a = np.clip(a, 0.0, 1.0)
            dist = 2.0 * r_terra * np.arcsin(np.sqrt(a))
            if dist.size:
                maior = max(maior, float(np.nanmax(dist)))
    return maior


def adicionar_metricas_clusters(df):
    """Acrescenta métricas do cluster sem alterar coordenadas originais das obras."""
    df = df.copy()
    if df.empty or 'CLUSTER_ID' not in df.columns:
        return df
    metricas = {}
    for cid, grp in df.groupby('CLUSTER_ID'):
        lats = grp['LATITUDE'].astype(float).to_numpy()
        lons = grp['LONGITUDE'].astype(float).to_numpy()
        max_km = _max_distancia_haversine_km(lats, lons)
        origens = sorted(grp['ORIGEM_BASE'].dropna().astype(str).unique().tolist()) if 'ORIGEM_BASE' in grp.columns else []
        metricas[cid] = {
            'QTD_OBRAS_CLUSTER': int(len(grp)),
            'QTD_SANEAMENTO_CLUSTER': int((grp['ORIGEM_BASE'] == 'SANEAMENTO').sum()) if 'ORIGEM_BASE' in grp.columns else 0,
            'QTD_LEVANTAMENTO_CLUSTER': int((grp['ORIGEM_BASE'] == 'LEVANTAMENTO').sum()) if 'ORIGEM_BASE' in grp.columns else 0,
            'DISTANCIA_MAX_CLUSTER_M': round(max_km * 1000.0, 1),
            'ORIGENS_CLUSTER': ' | '.join(origens),
            'LAT_CENTRO_CLUSTER': float(np.mean(lats)),
            'LONG_CENTRO_CLUSTER': float(np.mean(lons)),
        }
    for col in ['QTD_OBRAS_CLUSTER','QTD_SANEAMENTO_CLUSTER','QTD_LEVANTAMENTO_CLUSTER','DISTANCIA_MAX_CLUSTER_M','ORIGENS_CLUSTER','LAT_CENTRO_CLUSTER','LONG_CENTRO_CLUSTER']:
        df[col] = df['CLUSTER_ID'].map({k: v[col] for k, v in metricas.items()})
    return df


def montar_resumo_executivo(df_view, config, rejeitadas_obras=0, rejeitadas_localidades=0):
    itens = [
        ('ID da análise', config.get('id_analise', '-')),
        ('Versão das regras', config.get('versao_regras', VERSAO_REGRAS_ANALISE)),
        ('Total filtrado', len(df_view)),
        ('Saneamento', int((df_view.get('ORIGEM_BASE', pd.Series(dtype='object')) == 'SANEAMENTO').sum())),
        ('Levantamento', int((df_view.get('ORIGEM_BASE', pd.Series(dtype='object')) == 'LEVANTAMENTO').sum())),
        ('Notas inválidas', int(df_view.get('COR_NOME', pd.Series(dtype='object')).astype(str).str.contains('Inválidas', na=False).sum())),
        ('Notas duplicadas', int(df_view.get('DUPLICADA', pd.Series(dtype='object')).astype(str).eq('SIM').sum())),
        ('Notas próximas', int(df_view.get('PROXIMA', pd.Series(dtype='object')).astype(str).eq('SIM').sum())),
        ('Clusters', int(df_view['CLUSTER_ID'].nunique()) if 'CLUSTER_ID' in df_view.columns else len(df_view)),
        ('Municípios', int(df_view['MUNICIPIO'].dropna().nunique()) if 'MUNICIPIO' in df_view.columns else 0),
        ('Alertas equipe distante', int(df_view.get('ALERTA_EQUIPE_DISTANTE', pd.Series(dtype='object')).astype(str).eq('SIM').sum())),
        ('Sem registro SAP', int(df_view.get('SITUACAO SAP', pd.Series(dtype='object')).astype(str).eq('SEM REGISTRO SAP').sum())),
        ('Coordenadas de obras rejeitadas', int(rejeitadas_obras)),
        ('Localidades rejeitadas', int(rejeitadas_localidades)),
        ('Tempo total (s)', config.get('tempo_processamento_s', '-')),
    ]
    if 'EQUIPE_1' in df_view.columns:
        top = df_view['EQUIPE_1'].replace('', np.nan).dropna().value_counts().head(5)
        for pos, (nome, qtd) in enumerate(top.items(), start=1):
            itens.append((f'Top equipe próxima {pos}', f'{nome} ({int(qtd)} obras)'))
    return pd.DataFrame(itens, columns=['INDICADOR', 'VALOR'])


def executar_autoteste_core_analise():
    erros = []
    try:
        vals = limpar_nota_serie(pd.Series(['123.0', '123.01', '456']))
        if vals.tolist() != ['123', '123.01', '456']:
            erros.append('normalização de NOTA')
    except Exception:
        erros.append('normalização de NOTA')

    try:
        linha_ok = pd.Series({
            'ORIGEM_BASE': 'LEVANTAMENTO', 'SITUACAO SAP': 'APTO',
            'STATUS SISCO': 'Pré Análise', 'STATUS LIST': 0,
            'COLUNA_STATUS_SISCO_LOCALIZADA': 'SIM', 'COLUNA_STATUS_LIST_LOCALIZADA': 'SIM'
        })
        v = avaliar_validade_fluxo(linha_ok)
        if not v['valida'] or v['status_sisco_normalizado'] != 'PRE ANALISE' or v['status_list_normalizado'] != '0':
            erros.append('validação APTO/Pré Análise/0')
        linha_sem_sisco = pd.Series({
            'ORIGEM_BASE': 'LEVANTAMENTO', 'SITUACAO SAP': 'APTO', 'STATUS LIST': 0,
            'COLUNA_STATUS_SISCO_LOCALIZADA': 'NÃO', 'COLUNA_STATUS_LIST_LOCALIZADA': 'SIM'
        })
        vv = avaliar_validade_fluxo(linha_sem_sisco)
        if vv['valida'] or 'não localizada' not in vv['motivo'].lower():
            erros.append('ausência de coluna SISCO não pode virar 0')
        sem_sap = classificar_situacao_sap('999', {}, True, {'1'})
        if sem_sap != 'SEM REGISTRO SAP':
            erros.append('SAP inexistente não pode virar APTO')
        cor_t, nome_t = determinar_classificacao_analise(pd.Series({
            'NOTA_VALIDA_FLUXO': 'SIM', 'DUPLICADA': 'SIM', 'PROXIMA': 'NÃO', 'ORIGEM_BASE': 'LEVANTAMENTO'
        }))
        if cor_t != 'red' or 'Duplicadas' not in nome_t:
            erros.append('prioridade válida + duplicada = vermelho')
    except Exception:
        erros.append('validação de fluxo')

    try:
        df_t = pd.DataFrame({'LATITUDE': [-2.5, -2.5], 'LONGITUDE': [-44.2, -44.2]})
        det = colaboradores_proximos_em_lote(df_t, np.array([-2.5, -3.0]), np.array([-44.2, -44.0]), np.array(['Equipe A','Equipe B']), np.array(['Saneamento','Levantamento']), 2, 100, 100)
        if det.iloc[0]['EQUIPE_1'] != 'Equipe A' or det.iloc[1]['EQUIPE_1'] != 'Equipe A':
            erros.append('equipes próximas/cache por coordenada')
    except Exception:
        erros.append('equipes próximas/cache por coordenada')

    try:
        d = pd.DataFrame({
            'NOTA':['1','1','1'], 'ORIGEM_BASE':['SANEAMENTO','LEVANTAMENTO','LEVANTAMENTO'], 'MUNICIPIO':['A','A','A'],
            'LATITUDE':[-2.5,-3.5,-2.5001], 'LONGITUDE':[-44.2,-45.2,-44.2001],
            'COR_NOME':['🔴 Notas Duplicadas']*3, 'COR_MAPA':['red']*3, 'CLUSTER_ID':[0,1,2],
            'SITUACAO SAP':['APTO']*3, 'NOTA_VALIDA_FLUXO':['SIM']*3, 'MOTIVO_INVALIDADE':['-']*3,
            'DUPLICADA':['SIM']*3, 'COLABORADORES MAIS PROXIMOS':['Equipe A','Equipe B','Equipe C'],
            'ALERTA_EQUIPE_DISTANTE':['NÃO']*3
        })
        a = auditar_duplicidades_geograficas(d, {'1'}, 100)
        if not (a['QTD_OCORRENCIAS_NOTA'] == 3).all():
            erros.append('contagem de múltiplas ocorrências')
        # A ocorrência de Saneamento deve escolher a contraparte mais próxima, não a primeira arbitrária.
        if a.iloc[0]['DUPLICATA_MESMO_LOCAL'] != 'SIM':
            erros.append('seleção da contraparte mais próxima')
        if not a['LINK_DUPLICATA_MAPS'].astype(str).str.startswith('https://www.google.com/maps?q=').all():
            erros.append('link bidirecional de duplicidade')
        mesmo_mun_distante = a[(a['ORIGEM_BASE'] == 'LEVANTAMENTO') & (a['LATITUDE'] == -3.5)].iloc[0]
        if mesmo_mun_distante['COORDENADA_DIVERGENTE'] != 'SIM' or mesmo_mun_distante['MUNICIPIO_DIVERGENTE'] != 'NÃO':
            erros.append('município igual com coordenada distante')
        k = gerar_kml_analise(adicionar_metricas_clusters(a))
        if '<kml' not in k or '<Placemark>' not in k or 'Abrir outra ocorrência no Google Maps' not in k or 'Ligações de Duplicadas' not in k:
            erros.append('exportação KML/link/linha duplicidade')
    except Exception:
        erros.append('auditoria/KML')
    return {'ok': not erros, 'erros': erros}


def limpar_estado_analise():
    for chave in [
        'df_final_analise', 'is_done_analise', 'df_coord_rejeitadas_analise',
        'df_coord_corrigidas_analise', 'config_analise', 'bytes_excel_analise',
        'bytes_kml_analise', 'export_sig_excel_analise', 'export_sig_kml_analise',
        'df_loc_rejeitadas_analise', 'df_loc_corrigidas_analise', 'df_loc_alertas_analise', 'autoteste_core_analise',
        'filtro_cores_analise', 'filtro_origem_analise', 'filtro_municipio_analise',
        'filtro_nota_analise', 'filtro_colab_analise', 'mostrar_mapa_analise'
    ]:
        st.session_state.pop(chave, None)
    st.session_state.df_final_analise = pd.DataFrame()
    st.session_state.is_done_analise = False


# ==============================================================
# ESTADO / CABEÇALHO
# ==============================================================
st.markdown("<h1 class='brand-title'>🔍 Análise Cruzada e Auditoria</h1>", unsafe_allow_html=True)
st.info("💡 Cruza bases de Saneamento e Levantamento, isola notas bloqueadas (SAP FINL/CANC) e define os colaboradores mais próximos independente da origem.")

if "df_final_analise" not in st.session_state:
    st.session_state.df_final_analise = pd.DataFrame()
if "is_done_analise" not in st.session_state:
    st.session_state.is_done_analise = False
if "autoteste_core_analise" not in st.session_state:
    st.session_state.autoteste_core_analise = executar_autoteste_core_analise()


# ==============================================================
# SIDEBAR
# ==============================================================
with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    cfg_locked = st.session_state.is_done_analise and not st.session_state.df_final_analise.empty
    raio_prox = st.slider("Distância p/ agrupar obras (Metros)", 10, 1000, 100, 10, key='raio_prox_analise', disabled=cfg_locked)
    qtd_equipes_prox = st.number_input("Qtd. de equipes mais próximas", min_value=1, max_value=5, value=2, step=1, key='qtd_equipes_prox_analise', disabled=cfg_locked)
    distancia_max_equipe = st.number_input("Distância máx. p/ equipes adicionais (km)", min_value=1.0, max_value=1000.0, value=100.0, step=10.0, key='dist_max_equipe_analise', disabled=cfg_locked)
    distancia_alerta_equipe = st.number_input("Alerta: equipe mais próxima acima de (km)", min_value=1.0, max_value=2000.0, value=100.0, step=10.0, key='dist_alerta_equipe_analise', disabled=cfg_locked)

    teste_core = st.session_state.get('autoteste_core_analise', {'ok': True, 'erros': []})
    if teste_core.get('ok'):
        st.caption("✅ Testes internos de regressão: OK")
    else:
        with st.expander("⚠️ Falha em teste interno", expanded=False):
            for err in teste_core.get('erros', []):
                st.write(f"• {err}")

    if cfg_locked:
        cfg_usada = st.session_state.get('config_analise', {})
        with st.expander("🔒 Configuração utilizada nesta análise", expanded=False):
            st.write(f"Raio: **{cfg_usada.get('raio_proximidade_m', '-')} m**")
            st.write(f"Equipes próximas: **{cfg_usada.get('qtd_equipes_proximas', '-')}**")
            st.write(f"Distância adicional: **{cfg_usada.get('distancia_max_equipe_km', '-')} km**")
            st.write(f"Alerta equipe distante: **{cfg_usada.get('distancia_alerta_equipe_km', '-')} km**")
            st.write(f"Versão das regras: **{cfg_usada.get('versao_regras', VERSAO_REGRAS_ANALISE)}**")

    st.markdown("---")

    if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
        df_fin = st.session_state.df_final_analise.copy()

        if 'COR_NOME' not in df_fin.columns:
            limpar_estado_analise()
            st.rerun()

        config_analise = st.session_state.get('config_analise', {})
        st.markdown("### 🎨 Filtros (Mapa e Export)")

        contagem_cores = df_fin['COR_NOME'].value_counts().to_dict()
        opcoes_cores = sorted(df_fin['COR_NOME'].dropna().unique().tolist())
        cores_selecionadas = st.multiselect(
            "Classificação:",
            opcoes_cores,
            default=opcoes_cores,
            format_func=lambda x: f"{x} ({int(contagem_cores.get(x, 0))})",
            key='filtro_cores_analise'
        )
        if not cores_selecionadas:
            st.warning("Selecione pelo menos uma classificação.")
            st.stop()

        opcoes_origem = sorted(df_fin['ORIGEM_BASE'].dropna().astype(str).unique().tolist()) if 'ORIGEM_BASE' in df_fin.columns else []
        origens_selecionadas = st.multiselect("Origem:", opcoes_origem, default=opcoes_origem, key='filtro_origem_analise')

        opcoes_mun = sorted([x for x in df_fin.get('MUNICIPIO', pd.Series(dtype='object')).dropna().astype(str).unique().tolist() if x.strip()])
        municipios_selecionados = st.multiselect("Municípios (vazio = todos):", opcoes_mun, default=[], key='filtro_municipio_analise')
        busca_nota = st.text_input("Pesquisar NOTA:", value='', key='filtro_nota_analise').strip()
        busca_colaborador = st.text_input("Pesquisar colaborador próximo:", value='', key='filtro_colab_analise').strip()

        df_view = df_fin[df_fin['COR_NOME'].isin(cores_selecionadas)].copy()
        if opcoes_origem:
            df_view = df_view[df_view['ORIGEM_BASE'].astype(str).isin(origens_selecionadas)].copy()
        if municipios_selecionados:
            df_view = df_view[df_view['MUNICIPIO'].astype(str).isin(municipios_selecionados)].copy()
        if busca_nota:
            df_view = df_view[df_view['NOTA'].astype(str).str.contains(re.escape(busca_nota), case=False, na=False)].copy()
        if busca_colaborador:
            serie_colab = df_view.get('COLABORADORES MAIS PROXIMOS', pd.Series(index=df_view.index, dtype='object'))
            df_view = df_view[serie_colab.astype(str).str.contains(re.escape(busca_colaborador), case=False, na=False)].copy()

        st.caption(f"Registros após filtros: **{len(df_view)}** de **{len(df_fin)}**")

        filtros_export = {
            'Classificações': ', '.join(cores_selecionadas),
            'Origens': ', '.join(origens_selecionadas) if origens_selecionadas else '-',
            'Municípios': ', '.join(municipios_selecionados) if municipios_selecionados else 'TODOS',
            'Busca NOTA': busca_nota or '-',
            'Busca colaborador': busca_colaborador or '-',
        }
        filtro_sig = (
            tuple(sorted(cores_selecionadas)), tuple(sorted(origens_selecionadas)),
            tuple(sorted(municipios_selecionados)), busca_nota.upper(), busca_colaborador.upper()
        )

        if st.session_state.get('export_sig_excel_analise') != filtro_sig:
            st.session_state.pop('bytes_excel_analise', None)
        if st.session_state.get('export_sig_kml_analise') != filtro_sig:
            st.session_state.pop('bytes_kml_analise', None)

        st.markdown("---")
        id_analise = config_analise.get('id_analise', criar_id_analise())
        id_safe = re.sub(r'[^A-Za-z0-9_-]', '_', str(id_analise))

        if 'bytes_excel_analise' not in st.session_state:
            if st.button("⚙️ Gerar Planilha Excel", use_container_width=True):
                with st.spinner("Gerando planilhas..."):
                    df_corrigidas = st.session_state.get('df_coord_corrigidas_analise', pd.DataFrame()).copy()
                    df_rejeitadas = st.session_state.get('df_coord_rejeitadas_analise', pd.DataFrame()).copy()
                    df_loc_corrigidas = st.session_state.get('df_loc_corrigidas_analise', pd.DataFrame()).copy()
                    df_loc_rejeitadas = st.session_state.get('df_loc_rejeitadas_analise', pd.DataFrame()).copy()
                    df_loc_alertas = st.session_state.get('df_loc_alertas_analise', pd.DataFrame()).copy()

                    resumo_exec = montar_resumo_executivo(
                        df_view, config_analise,
                        rejeitadas_obras=len(df_rejeitadas),
                        rejeitadas_localidades=len(df_loc_rejeitadas)
                    )
                    classificacao_final = (
                        df_view['COR_NOME'].value_counts().rename_axis('CLASSIFICACAO_FINAL').reset_index(name='QUANTIDADE')
                        if 'COR_NOME' in df_view.columns else pd.DataFrame()
                    )

                    classe_dup = df_view.get('CLASSIFICACAO_DUPLICIDADE_GEO', pd.Series(index=df_view.index, dtype='object')).astype(str)
                    dict_dfs = {
                        'Resumo Executivo': resumo_exec,
                        'Classificacao Final': classificacao_final,
                        'Consolidado (Todas)': df_view,
                        'Apenas Saneamento': df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'],
                        'Apenas Levantamento': df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO'],
                        'Notas Inválidas': df_view[df_view['COR_NOME'].astype(str).str.contains('Inválidas', na=False)],
                        'Notas Duplicadas': df_view[df_view['DUPLICADA'].astype(str).eq('SIM')] if 'DUPLICADA' in df_view.columns else pd.DataFrame(),
                        'Duplicadas Locais Diferentes': df_view[classe_dup.eq('LOCAIS DIFERENTES')],
                        'Divergencias Duplicadas': df_view[(df_view.get('COORDENADA_DIVERGENTE', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM')) | (df_view.get('MUNICIPIO_DIVERGENTE', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM'))],
                        'Notas Próximas': df_view[df_view['PROXIMA'].astype(str).eq('SIM')] if 'PROXIMA' in df_view.columns else pd.DataFrame(),
                        'Notas Solitárias': df_view[df_view['PROXIMA'].astype(str).eq('NÃO')] if 'PROXIMA' in df_view.columns else pd.DataFrame(),
                        'Coordenadas Corrigidas': df_corrigidas,
                        'Coordenadas Rejeitadas': df_rejeitadas,
                        'Localidades Corrigidas': df_loc_corrigidas,
                        'Localidades Rejeitadas': df_loc_rejeitadas,
                        'Alertas Localidades': df_loc_alertas,
                    }
                    excel_bytes = gerar_excel_analise(dict_dfs)
                    bu_xl = io.BytesIO()
                    with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
                        zx.writestr(f"Analise_Cruzada_{id_safe}.xlsx", excel_bytes)
                        zx.writestr(f"Configuracao_Analise_{id_safe}.txt", montar_config_txt(config_analise, cores_selecionadas, filtros_export).encode('utf-8'))
                    st.session_state.bytes_excel_analise = bu_xl.getvalue()
                    st.session_state.export_sig_excel_analise = filtro_sig
                st.rerun()
        else:
            st.download_button(
                "🌐 Baixar Planilha Excel (ZIP)",
                data=st.session_state.bytes_excel_analise,
                file_name=f"Analise_Planilhas_{id_safe}.zip",
                use_container_width=True
            )

        if 'bytes_kml_analise' not in st.session_state:
            if st.button("⚙️ Gerar Mapa KML", use_container_width=True):
                with st.spinner("Gerando KML..."):
                    kml_str = gerar_kml_analise(df_view, nome_documento=f"Análise Cruzada {id_analise}")
                    bu_kml = io.BytesIO()
                    with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
                        zk.writestr(f"Mapa_Analise_Cruzada_{id_safe}.kml", kml_str.encode('utf-8'))
                        zk.writestr(f"Configuracao_Analise_{id_safe}.txt", montar_config_txt(config_analise, cores_selecionadas, filtros_export).encode('utf-8'))
                    st.session_state.bytes_kml_analise = bu_kml.getvalue()
                    st.session_state.export_sig_kml_analise = filtro_sig
                st.rerun()
        else:
            st.download_button(
                "🗺️ Baixar Mapa (KML)",
                data=st.session_state.bytes_kml_analise,
                file_name=f"Analise_Mapa_{id_safe}.zip",
                use_container_width=True
            )

        if st.button("🧹 Nova Análise", type="primary", use_container_width=True):
            limpar_estado_analise()
            st.rerun()


# ==============================================================
# RESULTADOS
# ==============================================================
if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
    config_analise = st.session_state.get('config_analise', {})
    id_analise = config_analise.get('id_analise', '-')
    st.caption(f"ID da análise: **{id_analise}**")

    if not config_analise.get('status_sap_localizado', True):
        st.warning("⚠️ A coluna de Status SAP não foi localizada na base de Levantamento. As notas aparecem como SEM REGISTRO SAP; o sistema não presume APTO.")

    tempos = config_analise.get('tempos_etapas', {}) or {}
    if tempos:
        with st.expander("⏱️ Tempo por etapa", expanded=False):
            cols_tempo = st.columns(min(4, max(1, len(tempos))))
            for i, (etapa, seg) in enumerate(tempos.items()):
                cols_tempo[i % len(cols_tempo)].metric(etapa, f"{seg:.2f}s")

    st.markdown("### 📈 Resumo da Volumetria")
    total_san = len(df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO']) if 'ORIGEM_BASE' in df_view.columns else 0
    total_lev = len(df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO']) if 'ORIGEM_BASE' in df_view.columns else 0

    col_a, col_b, col_c = st.columns(3)
    col_a.info(f"**🟣 Saneamento (Validado):** {total_san} obras")
    col_b.success(f"**🟢 Levantamento (Validado):** {total_lev} obras")
    col_c.warning(f"**🎯 Total Geral:** {len(df_view)} obras")

    invalidas = int(df_view.get('COR_NOME', pd.Series(index=df_view.index, dtype='object')).astype(str).str.contains('Inválidas', na=False).sum())
    duplicadas = int(df_view.get('DUPLICADA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    proximas = int(df_view.get('PROXIMA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    corrigidas = int(df_view.get('COORDENADA_CORRIGIDA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    clusters = int(df_view['CLUSTER_ID'].nunique()) if 'CLUSTER_ID' in df_view.columns else len(df_view)
    municipios = int(df_view['MUNICIPIO'].dropna().nunique()) if 'MUNICIPIO' in df_view.columns else 0
    alertas_geo = int(df_view.get('COORDENADA_ALERTA', pd.Series(index=df_view.index, dtype='object')).astype(str).str.strip().ne('').sum())
    alertas_equipe = int(df_view.get('ALERTA_EQUIPE_DISTANTE', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    dup_coord_div = int(df_view.get('COORDENADA_DIVERGENTE', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    rejeitadas = len(st.session_state.get('df_coord_rejeitadas_analise', pd.DataFrame()))
    loc_rejeitadas = len(st.session_state.get('df_loc_rejeitadas_analise', pd.DataFrame()))

    st.markdown("#### 🧪 Qualidade da Análise")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Notas inválidas", invalidas)
    q2.metric("Notas duplicadas", duplicadas)
    q3.metric("Duplicadas em locais distintos", dup_coord_div)
    q4.metric("Clusters", clusters)
    q5, q6, q7, q8 = st.columns(4)
    q5.metric("Municípios", municipios)
    q6.metric("Coords. corrigidas", corrigidas)
    q7.metric("Coords. rejeitadas", rejeitadas)
    q8.metric("Equipe distante", alertas_equipe)
    sem_sap = int(df_view.get('SITUACAO SAP', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SEM REGISTRO SAP').sum())
    if sem_sap > 0:
        st.info(f"ℹ️ {sem_sap} registro(s) sem confirmação SAP. Eles não são tratados automaticamente como APTO.")

    if rejeitadas > 0:
        with st.expander(f"⚠️ {rejeitadas} obras com coordenadas rejeitadas", expanded=False):
            st.dataframe(st.session_state.df_coord_rejeitadas_analise, use_container_width=True, hide_index=True)
    if loc_rejeitadas > 0:
        with st.expander(f"👥 {loc_rejeitadas} localidades rejeitadas", expanded=False):
            st.dataframe(st.session_state.df_loc_rejeitadas_analise, use_container_width=True, hide_index=True)
    loc_corrigidas = len(st.session_state.get('df_loc_corrigidas_analise', pd.DataFrame()))
    loc_alertas = len(st.session_state.get('df_loc_alertas_analise', pd.DataFrame()))
    if loc_corrigidas > 0:
        with st.expander(f"🧭 {loc_corrigidas} localidades com coordenadas corrigidas", expanded=False):
            st.dataframe(st.session_state.df_loc_corrigidas_analise, use_container_width=True, hide_index=True)
    if loc_alertas > 0:
        with st.expander(f"🧭 {loc_alertas} localidades com alertas geográficos", expanded=False):
            st.dataframe(st.session_state.df_loc_alertas_analise, use_container_width=True, hide_index=True)
    if alertas_geo > 0:
        with st.expander(f"🌍 {alertas_geo} registros com coordenadas geograficamente atípicas", expanded=False):
            cols_alerta = [c for c in ['NOTA', 'ORIGEM_BASE', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'COORDENADA_ALERTA'] if c in df_view.columns]
            st.dataframe(df_view[df_view['COORDENADA_ALERTA'].astype(str).str.strip().ne('')][cols_alerta], use_container_width=True, hide_index=True)
    if alertas_equipe > 0:
        with st.expander(f"🚗 {alertas_equipe} obras com equipe mais próxima acima do limite de alerta", expanded=False):
            cols_eq = [c for c in ['NOTA','MUNICIPIO','EQUIPE_1','TIPO_EQUIPE_1','DISTANCIA_EQUIPE_1_KM','ALERTA_EQUIPE_DISTANTE'] if c in df_view.columns]
            st.dataframe(df_view[df_view['ALERTA_EQUIPE_DISTANTE'].astype(str).eq('SIM')][cols_eq], use_container_width=True, hide_index=True)

    st.markdown("### 🗺️ Mapa Analítico")
    mostrar_mapa = st.checkbox("Exibir mapa analítico", value=False, key='mostrar_mapa_analise')
    if mostrar_mapa:
        mapa = folium.Map(location=[df_view['LATITUDE'].mean(), df_view['LONGITUDE'].mean()], zoom_start=8) if not df_view.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        layers = {}
        linhas_dup_layer = folium.FeatureGroup(name="🔗 Ligações de Duplicadas", show=True).add_to(mapa)
        pares_dup_desenhados = set()

        # Liga visualmente ocorrências duplicadas em locais diferentes, sem duplicar a mesma linha.
        for _, rr in df_view.iterrows():
            if str(rr.get("CLASSIFICACAO_DUPLICIDADE_GEO", "")) != "LOCAIS DIFERENTES":
                continue
            if pd.isna(rr.get("DUPLICATA_LAT_DESTINO")) or pd.isna(rr.get("DUPLICATA_LON_DESTINO")):
                continue
            a = (round(float(rr["LATITUDE"]), 6), round(float(rr["LONGITUDE"]), 6))
            b = (round(float(rr["DUPLICATA_LAT_DESTINO"]), 6), round(float(rr["DUPLICATA_LON_DESTINO"]), 6))
            chave_par = (str(rr.get("NOTA", "")), tuple(sorted([a, b])))
            if chave_par in pares_dup_desenhados:
                continue
            pares_dup_desenhados.add(chave_par)
            folium.PolyLine(
                [a, b], color="red", weight=2, opacity=0.75, dash_array="6,6",
                tooltip=f"Duplicata {rr.get('NOTA', '')} - {rr.get('DISTANCIA_ENTRE_DUPLICATAS_KM', '-')} km"
            ).add_to(linhas_dup_layer)

        def obter_layer(nome):
            if nome not in layers:
                fg = folium.FeatureGroup(name=str(nome), show=True)
                mc = MarkerCluster().add_to(fg)
                fg.add_to(mapa)
                layers[nome] = mc
            return layers[nome]

        for cid, grp in df_view.groupby('CLUSTER_ID'):
            lat = grp['LAT_CENTRO_CLUSTER'].iloc[0] if 'LAT_CENTRO_CLUSTER' in grp.columns else grp['LATITUDE'].mean()
            lon = grp['LONG_CENTRO_CLUSTER'].iloc[0] if 'LONG_CENTRO_CLUSTER' in grp.columns else grp['LONGITUDE'].mean()
            c_names = grp['COR_NOME'].astype(str).tolist()

            if any('Inválidas' in c or 'Preto' in c for c in c_names):
                c_i, layer_nome = 'black', '⚫ Notas Inválidas'
            elif any('Duplicadas' in c or 'Vermelho' in c for c in c_names):
                c_i, layer_nome = 'red', '🔴 Notas Duplicadas'
            elif len(grp) > 1 or any('Próximas' in c for c in c_names):
                c_i, layer_nome = 'orange', '🟠 Notas Próximas'
            else:
                c_i = str(grp['COR_MAPA'].iloc[0])
                layer_nome = str(grp['COR_NOME'].iloc[0])

            total_cluster = int(grp['QTD_OBRAS_CLUSTER'].iloc[0]) if 'QTD_OBRAS_CLUSTER' in grp.columns and pd.notna(grp['QTD_OBRAS_CLUSTER'].iloc[0]) else len(grp)
            titulo_card = f"📍 Obras no Local ({len(grp)})" if len(grp) == total_cluster else f"📍 {len(grp)} visíveis de {total_cluster} obras no local"
            pop_html = f'''
            <div style="font-family:sans-serif; width:300px; max-height:340px; overflow-y:auto; border-radius:8px; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                <div style="background:#0D256C; color:#ffffff; padding:8px; font-size:13px; font-weight:bold; text-align:center; position:sticky; top:0;">{titulo_card}</div>
                <div style="padding:10px; background:#fafafa; font-size:12px;">
            '''

            if 'DISTANCIA_MAX_CLUSTER_M' in grp.columns:
                pop_html += f"<div style='margin-bottom:6px;color:#555;'><b>Cluster:</b> {total_cluster} obra(s) | diâmetro máx.: {grp['DISTANCIA_MAX_CLUSTER_M'].iloc[0]} m</div>"

            for _, r in grp.iterrows():
                n = html.escape(str(r.get('NOTA', '')))
                mun = html.escape(str(r.get('MUNICIPIO', '')))
                origem_raw = str(r.get('ORIGEM_BASE', ''))
                o = html.escape(origem_raw)
                sap = html.escape(str(r.get('SITUACAO SAP', '')))
                sap_fonte = html.escape(str(r.get('FONTE_VALIDACAO_SAP', '-')))
                s_sisco_raw = valor_alias(r, ['STATUS SISCO', 'STATUS_SISCO'], '-')
                s_list_raw = valor_alias(r, ['STATUS LIST', 'STATUS_LIST'], '-')
                if origem_raw.strip().upper() == 'LEVANTAMENTO' and s_sisco_raw == '-':
                    s_sisco_raw = str(r.get('STATUS_SISCO_NORMALIZADO', 'NÃO INFORMADO')).replace('NAO ', 'NÃO ')
                if origem_raw.strip().upper() == 'LEVANTAMENTO' and s_list_raw == '-':
                    s_list_raw = str(r.get('STATUS_LIST_NORMALIZADO', 'NÃO INFORMADO')).replace('NAO ', 'NÃO ')
                s_sisco = html.escape(s_sisco_raw)
                s_list = html.escape(s_list_raw)
                col = html.escape(str(r.get('COLABORADORES MAIS PROXIMOS', '')))
                dup = html.escape(str(r.get('DUPLICADA', '')))
                dup_geo = html.escape(str(r.get('CLASSIFICACAO_DUPLICIDADE_GEO', '')))
                dist_dup = r.get('DISTANCIA_ENTRE_DUPLICATAS_KM', np.nan)
                dist_dup_txt = f"{float(dist_dup):.3f} km" if pd.notna(dist_dup) else '-'
                alert_eq = html.escape(str(r.get('ALERTA_EQUIPE_DISTANTE', 'NÃO')))
                classificacao = html.escape(str(r.get('COR_NOME', '-')))
                motivo_inval = html.escape(str(r.get('MOTIVO_INVALIDADE', '-')))
                nota_valida_fluxo = str(r.get('NOTA_VALIDA_FLUXO', 'SIM')).upper()
                qtd_ocorr = int(r.get('QTD_OCORRENCIAS_NOTA', 0) or 0)
                qtd_san = int(r.get('QTD_OCORRENCIAS_SANEAMENTO', 0) or 0)
                qtd_lev = int(r.get('QTD_OCORRENCIAS_LEVANTAMENTO', 0) or 0)
                link_atual = f"https://www.google.com/maps?q={float(r.get('LATITUDE')):.8f},{float(r.get('LONGITUDE')):.8f}"
                link_atual_safe = html.escape(link_atual, quote=True)

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

                pop_html += f'''
                <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
                    <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                    <tr><td style="padding:2px;"><b>Classificação:</b></td><td style="padding:2px;font-weight:bold;">{classificacao}</td></tr>
                    <tr><td style="padding:2px;"><b>Município:</b></td><td style="padding:2px;">{mun}</td></tr>
                    <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                    <tr><td style="padding:2px;"><b>SAP:</b></td><td style="padding:2px;">{sap}</td></tr>
                    <tr><td style="padding:2px;"><b>Validação SAP:</b></td><td style="padding:2px;">{sap_fonte}</td></tr>
                    <tr><td style="padding:2px;"><b>SISCO / LIST:</b></td><td style="padding:2px;">{s_sisco} / {s_list}</td></tr>
                    {motivo_html}
                    <tr><td style="padding:2px;"><b>Equipes Perto:</b></td><td style="padding:2px;">{col}</td></tr>
                    <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup} {dup_geo}</td></tr>
                    <tr><td style="padding:2px;"><b>Ocorrências:</b></td><td style="padding:2px;">{qtd_ocorr} (Saneamento: {qtd_san} | Levantamento: {qtd_lev})</td></tr>
                    <tr><td style="padding:2px;"><b>Dist. duplicata:</b></td><td style="padding:2px;">{dist_dup_txt}</td></tr>
                    <tr><td style="padding:2px;"><b>Equipe distante:</b></td><td style="padding:2px;">{alert_eq}</td></tr>
                </table>
                <div style='margin:6px 0;'><a href='{link_atual_safe}' target='_blank' style='color:#0D47A1;font-weight:bold;text-decoration:none;'>📍 Abrir este ponto no Google Maps</a></div>
                {link_dup_html}
                <hr style="margin:4px 0; border:0; border-top:1px solid #ccc;">
                '''

            pop_html += '</div></div>'
            folium.Marker([lat, lon], icon=folium.Icon(color=c_i, icon='info-sign'), popup=folium.Popup(pop_html, max_width=340)).add_to(obter_layer(layer_nome))

        folium.LayerControl(collapsed=False).add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550)
    else:
        st.caption("O mapa é carregado apenas quando solicitado para manter a tela mais rápida.")

    st.markdown("### 📊 Tabela Consolidada Detalhada")
    st.dataframe(
        df_view.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME'], errors='ignore'),
        use_container_width=True,
        hide_index=True
    )


# ==============================================================
# ENTRADA / PROCESSAMENTO
# ==============================================================
else:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### 🧹 Base Saneamento")
        file_san = st.file_uploader("Upload Saneamento", type=["xlsx", "xls", "csv"], key="san")
    with c2:
        st.markdown("#### 📋 Base Levantamento")
        file_lev = st.file_uploader("Upload Levantamento", type=["xlsx", "xls", "csv"], key="lev")
    with c3:
        st.markdown("#### 👥 Localidade Equipes")
        file_loc = st.file_uploader("Upload Localidades", type=["xlsx", "xls", "csv"], key="loc")

    if not (file_san and file_lev and file_loc):
        st.stop()

    if st.button("🚀 Processar Análise Cruzada", type="primary", use_container_width=True):
        st_run = time.time()
        pb = st.progress(0.0)
        tmp = st.empty()
        sgt = st.empty()

        def render_t(pct, msg):
            e = time.time() - st_run
            f = pct
            if f > 0:
                rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.05 else "Calculando..."
            else:
                rs = "Calculando..."
            es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
            tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom:20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)
            if msg:
                sgt.info(msg)
            pb.progress(pct)

        render_t(0.08, "Lendo e padronizando planilhas...")

        try:
            tempos_etapas = {}
            t_etapa = time.time()

            if file_san.name.lower().endswith('.csv'):
                df_san_raw = ler_csv_resiliente(file_san.getvalue())
            else:
                df_san_raw = ler_planilha_cached(file_san.getvalue())

            if file_lev.name.lower().endswith('.csv'):
                df_lev_raw = ler_csv_resiliente(file_lev.getvalue())
            else:
                df_lev_raw = ler_planilha_cached(file_lev.getvalue())

            df_san = preparar_base_obras(df_san_raw, 'SANEAMENTO')
            df_lev = preparar_base_obras(df_lev_raw, 'LEVANTAMENTO')

            col_sisco_lev = encontrar_coluna(df_lev, ['STATUS_SISCO', 'STATUS SISCO'])
            col_list_lev = encontrar_coluna(df_lev, ['STATUS_LIST', 'STATUS LIST'])
            df_lev['COLUNA_STATUS_SISCO_LOCALIZADA'] = 'SIM' if col_sisco_lev is not None else 'NÃO'
            df_lev['COLUNA_STATUS_LIST_LOCALIZADA'] = 'SIM' if col_list_lev is not None else 'NÃO'
            tempos_etapas['Leitura/Padronização'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.20, "Validando estrutura das bases...")
            ok = True
            ok &= validar_colunas(df_san, ['NOTA', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE'], 'Base Saneamento')
            ok &= validar_colunas(df_lev, ['NOTA', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE'], 'Base Levantamento')
            if not ok:
                st.stop()

            render_t(0.28, "Lendo e validando Localidades...")
            df_loc = preparar_localidades(file_loc)
            if not validar_colunas(df_loc, ['NOME_COLAB', 'LAT_LOC', 'LON_LOC'], 'Base Localidades'):
                st.stop()

            # Mesma auditoria/correção aplicada às coordenadas das obras.
            df_loc['LAT_LOC_ORIGINAL'] = df_loc['LAT_LOC']
            df_loc['LON_LOC_ORIGINAL'] = df_loc['LON_LOC']
            lat_loc_orig = pd.to_numeric(df_loc['LAT_LOC'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            lon_loc_orig = pd.to_numeric(df_loc['LON_LOC'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            df_loc['LAT_LOC'] = lat_loc_orig.apply(lambda x: corrigir_coord(x, 90))
            df_loc['LON_LOC'] = lon_loc_orig.apply(lambda x: corrigir_coord(x, 180))
            mudou_lat_loc = lat_loc_orig.notna() & df_loc['LAT_LOC'].notna() & ((lat_loc_orig - df_loc['LAT_LOC']).abs() > 1e-10)
            mudou_lon_loc = lon_loc_orig.notna() & df_loc['LON_LOC'].notna() & ((lon_loc_orig - df_loc['LON_LOC']).abs() > 1e-10)
            df_loc['COORDENADA_LOCALIDADE_CORRIGIDA'] = np.where(mudou_lat_loc | mudou_lon_loc, 'SIM', 'NÃO')

            nomes_limpos = df_loc['NOME_COLAB'].astype(str).str.strip()
            m_nome_invalido = nomes_limpos.str.upper().isin(['', 'NAN', 'NONE', 'NULL'])
            m_coord_loc_invalida = df_loc['LAT_LOC'].isna() | df_loc['LON_LOC'].isna()
            m_loc_rej = m_nome_invalido | m_coord_loc_invalida
            df_loc_rej = df_loc[m_loc_rej].copy()
            if not df_loc_rej.empty:
                motivos = []
                for idx in df_loc_rej.index:
                    mm = []
                    if m_nome_invalido.loc[idx]:
                        mm.append('Nome do colaborador vazio')
                    if pd.isna(df_loc_rej.loc[idx, 'LAT_LOC']):
                        mm.append('Latitude inválida')
                    if pd.isna(df_loc_rej.loc[idx, 'LON_LOC']):
                        mm.append('Longitude inválida')
                    motivos.append(' | '.join(mm))
                df_loc_rej['MOTIVO_REJEICAO'] = motivos

            df_loc = df_loc[~m_loc_rej].copy()
            df_loc['NOME_COLAB'] = df_loc['NOME_COLAB'].astype(str).str.strip()
            alertas_loc = []
            for _, rr in df_loc.iterrows():
                al = []
                la, lo = float(rr['LAT_LOC']), float(rr['LON_LOC'])
                if la == 0.0 or lo == 0.0:
                    al.append('Coordenada zerada')
                if la > 0 or lo > 0:
                    al.append('Coordenada positiva')
                if abs(la) > abs(lo):
                    al.append('Possível LAT/LON invertida')
                alertas_loc.append(' | '.join(al))
            df_loc['COORDENADA_LOCALIDADE_ALERTA'] = alertas_loc
            df_loc_corr = df_loc[df_loc['COORDENADA_LOCALIDADE_CORRIGIDA'] == 'SIM'].copy()
            st.session_state.df_loc_rejeitadas_analise = df_loc_rej
            st.session_state.df_loc_corrigidas_analise = df_loc_corr
            st.session_state.df_loc_alertas_analise = df_loc[df_loc['COORDENADA_LOCALIDADE_ALERTA'].astype(str).str.strip().ne('')].copy()

            if df_loc.empty:
                st.error("❌ A base de Localidades não possui colaboradores com nome e coordenadas válidas.")
                st.stop()

            lat_locs = df_loc['LAT_LOC'].to_numpy(dtype=float)
            lon_locs = df_loc['LON_LOC'].to_numpy(dtype=float)
            nomes_locs = df_loc['NOME_COLAB'].astype(str).to_numpy()
            tipos_locs = df_loc['TIPO_EQUIPE'].astype(str).to_numpy()
            tempos_etapas['Validação/Localidades'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.36, "Validando Status SAP (Bloqueios)...")
            status_col = encontrar_coluna(df_lev, ['STATUS_SAP', 'STATUS SAP', 'STATUS'])
            status_sap_localizado = status_col is not None
            notas_lev_cadastradas = set(df_lev.loc[df_lev['NOTA'].apply(nota_valida), 'NOTA'].astype(str))
            status_dict = {}
            if status_col:
                tmp_status = df_lev[['NOTA', status_col]].copy()
                tmp_status = tmp_status[tmp_status['NOTA'].apply(nota_valida)]
                tmp_status['_STATUS'] = tmp_status[status_col].apply(lambda v: normalizar_status_fluxo(v, ''))

                def consolidar_status(vals):
                    vals = [normalizar_status_fluxo(v, '') for v in vals]
                    vals = [v for v in vals if v]
                    bloqueados = []
                    if 'FINL' in vals:
                        bloqueados.append('FINL')
                    if 'CANC' in vals:
                        bloqueados.append('CANC')
                    if bloqueados:
                        return '/'.join(bloqueados)
                    return vals[-1] if vals else ''

                status_dict = tmp_status.groupby('NOTA')['_STATUS'].agg(consolidar_status).to_dict()

            df_san['SITUACAO SAP'] = df_san['NOTA'].apply(lambda n: classificar_situacao_sap(n, status_dict, status_sap_localizado, notas_lev_cadastradas))
            df_lev['SITUACAO SAP'] = df_lev['NOTA'].apply(lambda n: classificar_situacao_sap(n, status_dict, status_sap_localizado, notas_lev_cadastradas))
            df_san['FONTE_VALIDACAO_SAP'] = df_san['NOTA'].apply(lambda n: fonte_validacao_sap(n, status_sap_localizado, notas_lev_cadastradas))
            df_lev['FONTE_VALIDACAO_SAP'] = df_lev['NOTA'].apply(lambda n: fonte_validacao_sap(n, status_sap_localizado, notas_lev_cadastradas))

            render_t(0.44, "Verificando duplicidades nas bases...")
            notas_san_validas = df_san.loc[df_san['NOTA'].apply(nota_valida), 'NOTA']
            notas_lev_validas = df_lev.loc[df_lev['NOTA'].apply(nota_valida), 'NOTA']
            set_san = set(notas_san_validas)
            set_lev = set(notas_lev_validas)
            duplicadas_inter = set_san.intersection(set_lev)
            rep_san = set(notas_san_validas[notas_san_validas.duplicated(keep=False)])
            rep_lev = set(notas_lev_validas[notas_lev_validas.duplicated(keep=False)])

            df_san['DUPLICADA'] = df_san['NOTA'].apply(lambda x: 'SIM' if nota_valida(x) and x in duplicadas_inter else 'NÃO')
            df_lev['DUPLICADA'] = df_lev['NOTA'].apply(lambda x: 'SIM' if nota_valida(x) and x in duplicadas_inter else 'NÃO')
            df_san['DUPLICADA_INTERBASE'] = df_san['DUPLICADA']
            df_lev['DUPLICADA_INTERBASE'] = df_lev['DUPLICADA']
            df_san['REPETIDA_NA_ORIGEM'] = df_san['NOTA'].apply(lambda x: 'SIM' if nota_valida(x) and x in rep_san else 'NÃO')
            df_lev['REPETIDA_NA_ORIGEM'] = df_lev['NOTA'].apply(lambda x: 'SIM' if nota_valida(x) and x in rep_lev else 'NÃO')
            tempos_etapas['SAP/Duplicidades'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.53, "Validando e auditando coordenadas...")
            df_master = pd.concat([df_san, df_lev], ignore_index=True)
            df_master['LATITUDE_ORIGINAL'] = df_master['LATITUDE']
            df_master['LONGITUDE_ORIGINAL'] = df_master['LONGITUDE']

            lat_original_num = pd.to_numeric(df_master['LATITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            lon_original_num = pd.to_numeric(df_master['LONGITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            df_master['LAT_NUM'] = lat_original_num.apply(lambda x: corrigir_coord(x, 90))
            df_master['LON_NUM'] = lon_original_num.apply(lambda x: corrigir_coord(x, 180))
            df_master['LATITUDE'] = df_master['LAT_NUM']
            df_master['LONGITUDE'] = df_master['LON_NUM']

            mudou_lat = lat_original_num.notna() & df_master['LAT_NUM'].notna() & ((lat_original_num - df_master['LAT_NUM']).abs() > 1e-10)
            mudou_lon = lon_original_num.notna() & df_master['LON_NUM'].notna() & ((lon_original_num - df_master['LON_NUM']).abs() > 1e-10)
            df_master['COORDENADA_CORRIGIDA'] = np.where(mudou_lat | mudou_lon, 'SIM', 'NÃO')

            m_coord_invalida = df_master['LATITUDE'].isna() | df_master['LONGITUDE'].isna()
            df_rej_coord = df_master[m_coord_invalida].copy()
            if not df_rej_coord.empty:
                df_rej_coord['MOTIVO_REJEICAO'] = np.where(
                    df_rej_coord['LATITUDE'].isna() & df_rej_coord['LONGITUDE'].isna(),
                    'Latitude e Longitude inválidas',
                    np.where(df_rej_coord['LATITUDE'].isna(), 'Latitude inválida', 'Longitude inválida')
                )

            df_valid = df_master[~m_coord_invalida].copy()
            alertas = []
            for _, rr in df_valid.iterrows():
                al = []
                lat = float(rr['LATITUDE'])
                lon = float(rr['LONGITUDE'])
                if lat == 0.0 or lon == 0.0:
                    al.append('Coordenada zerada')
                if lat > 0 or lon > 0:
                    al.append('Coordenada positiva')
                if abs(lat) > abs(lon):
                    al.append('Possível LAT/LON invertida')
                alertas.append(' | '.join(al))
            df_valid['COORDENADA_ALERTA'] = alertas

            # Complementa a duplicidade com distância, município e divergência geográfica.
            df_valid = auditar_duplicidades_geograficas(df_valid, duplicadas_inter, raio_prox)

            df_corrigidas = df_valid[df_valid['COORDENADA_CORRIGIDA'] == 'SIM'].copy()
            st.session_state.df_coord_rejeitadas_analise = df_rej_coord
            st.session_state.df_coord_corrigidas_analise = df_corrigidas
            if df_valid.empty:
                st.error("❌ Nenhum registro possui coordenadas válidas para a análise espacial.")
                st.stop()
            tempos_etapas['Coordenadas'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.63, "Agrupando Obras Vizinhas e Processando Cores...")
            df_clustered, _ = fundir_super_pontos(df_valid, raio_metros=raio_prox, agrupar_por_levantador=False)

            expanded = []
            c_id = 0
            for _, r in df_clustered.iterrows():
                is_prox = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
                if is_prox:
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = dict(orig)
                        nr['PROXIMA'] = 'SIM'
                        nr['CLUSTER_ID'] = c_id
                        expanded.append(nr)
                else:
                    nr = r.to_dict()
                    nr['PROXIMA'] = 'NÃO'
                    nr['CLUSTER_ID'] = c_id
                    expanded.append(nr)
                c_id += 1

            df_final = adicionar_metricas_clusters(pd.DataFrame(expanded))
            tempos_etapas['Clusters'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.74, "Calculando colaboradores mais próximos...")
            detalhe_equipes = colaboradores_proximos_em_lote(
                df_final,
                lat_locs, lon_locs, nomes_locs, tipos_locs,
                qtd_equipes_prox,
                distancia_max_equipe,
                distancia_alerta_equipe
            )
            for c in detalhe_equipes.columns:
                df_final[c] = detalhe_equipes[c].values
            tempos_etapas['Equipes Próximas'] = round(time.time() - t_etapa, 3)

            t_etapa = time.time()
            render_t(0.86, "Classificando regras e cores da auditoria...")

            validacoes_fluxo = [avaliar_validade_fluxo(r) for _, r in df_final.iterrows()]
            df_final['STATUS_LIST_NORMALIZADO'] = [v['status_list_normalizado'] for v in validacoes_fluxo]
            df_final['STATUS_SISCO_NORMALIZADO'] = [v['status_sisco_normalizado'] for v in validacoes_fluxo]
            df_final['NOTA_VALIDA_FLUXO'] = ['SIM' if v['valida'] else 'NÃO' for v in validacoes_fluxo]
            df_final['MOTIVO_INVALIDADE'] = [v['motivo'] for v in validacoes_fluxo]

            cores_calculadas = [determinar_classificacao_analise(r) for _, r in df_final.iterrows()]
            df_final['COR_MAPA'] = [c[0] for c in cores_calculadas]
            df_final['COR_NOME'] = [c[1] for c in cores_calculadas]
            tempos_etapas['Classificação'] = round(time.time() - t_etapa, 3)

            id_analise = criar_id_analise()
            df_final['ID_ANALISE'] = id_analise

            front_cols = [
                'ID_ANALISE', 'NOTA', 'ORIGEM_BASE', 'SITUACAO SAP', 'FONTE_VALIDACAO_SAP', 'NOTA_VALIDA_FLUXO',
                'MOTIVO_INVALIDADE', 'STATUS_SISCO_NORMALIZADO', 'STATUS_LIST_NORMALIZADO',
                'COLUNA_STATUS_SISCO_LOCALIZADA', 'COLUNA_STATUS_LIST_LOCALIZADA', 'DUPLICADA',
                'DUPLICADA_INTERBASE', 'REPETIDA_NA_ORIGEM', 'CLASSIFICACAO_DUPLICIDADE_GEO',
                'DISTANCIA_ENTRE_DUPLICATAS_KM', 'DUPLICATA_MESMO_LOCAL', 'COORDENADA_DIVERGENTE',
                'MUNICIPIO_DIVERGENTE', 'QTD_OCORRENCIAS_NOTA', 'QTD_OCORRENCIAS_SANEAMENTO',
                'QTD_OCORRENCIAS_LEVANTAMENTO', 'QTD_OCORRENCIAS_OUTRA_BASE',
                'DUPLICATA_NOTA_DESTINO', 'DUPLICATA_ORIGEM_DESTINO',
                'DUPLICATA_MUNICIPIO_DESTINO', 'DUPLICATA_LAT_DESTINO', 'DUPLICATA_LON_DESTINO',
                'LINK_DUPLICATA_MAPS', 'PROXIMA', 'COLABORADORES MAIS PROXIMOS',
                'EQUIPE_1', 'TIPO_EQUIPE_1', 'DISTANCIA_EQUIPE_1_KM',
                'DISTANCIA_EQUIPE_MAIS_PROXIMA_KM', 'ALERTA_EQUIPE_DISTANTE',
                'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'LATITUDE_ORIGINAL', 'LONGITUDE_ORIGINAL',
                'COORDENADA_CORRIGIDA', 'COORDENADA_ALERTA', 'CLUSTER_ID', 'QTD_OBRAS_CLUSTER',
                'QTD_SANEAMENTO_CLUSTER', 'QTD_LEVANTAMENTO_CLUSTER', 'DISTANCIA_MAX_CLUSTER_M',
                'ORIGENS_CLUSTER', 'LAT_CENTRO_CLUSTER', 'LONG_CENTRO_CLUSTER', 'COR_MAPA', 'COR_NOME'
            ]
            front_cols = [c for c in front_cols if c in df_final.columns]
            rest_cols = [c for c in df_final.columns if c not in front_cols and not c.startswith('_')]
            df_final = df_final[front_cols + rest_cols]

            config = {
                'id_analise': id_analise,
                'versao_regras': VERSAO_REGRAS_ANALISE,
                'arquivo_saneamento': file_san.name,
                'arquivo_levantamento': file_lev.name,
                'arquivo_localidades': file_loc.name,
                'raio_proximidade_m': int(raio_prox),
                'qtd_equipes_proximas': int(qtd_equipes_prox),
                'distancia_max_equipe_km': float(distancia_max_equipe),
                'distancia_alerta_equipe_km': float(distancia_alerta_equipe),
                'linhas_saneamento': int(len(df_san)),
                'linhas_levantamento': int(len(df_lev)),
                'linhas_validas': int(len(df_final)),
                'coordenadas_rejeitadas': int(len(df_rej_coord)),
                'coordenadas_corrigidas': int(len(df_corrigidas)),
                'localidades_rejeitadas': int(len(df_loc_rej)),
                'localidades_corrigidas': int(len(df_loc_corr)),
                'localidades_alertas': int(len(st.session_state.get('df_loc_alertas_analise', pd.DataFrame()))),
                'status_sap_localizado': bool(status_sap_localizado),
                'sem_registro_sap': int(df_final['SITUACAO SAP'].astype(str).eq('SEM REGISTRO SAP').sum()),
                'duplicadas_interbase': int(len(duplicadas_inter)),
                'linhas_duplicadas_interbase': int(df_final['DUPLICADA'].astype(str).eq('SIM').sum()),
                'repetidas_saneamento': int(len(rep_san)),
                'repetidas_levantamento': int(len(rep_lev)),
                'tempos_etapas': tempos_etapas,
                'tempo_processamento_s': round(time.time() - st_run, 2),
            }

            render_t(1.0, "✅ Análise Concluída!")
            st.session_state.df_final_analise = df_final
            st.session_state.config_analise = config
            st.session_state.is_done_analise = True
            st.session_state.pop('bytes_excel_analise', None)
            st.session_state.pop('bytes_kml_analise', None)
            time.sleep(0.5)
            st.rerun()

        except Exception as e:
            codigo_erro = f"ANLERR-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
            detalhe_erro = traceback.format_exc()
            st.session_state.ultimo_erro_analise = {
                'codigo': codigo_erro, 'tipo': type(e).__name__, 'mensagem': str(e),
                'data_hora': datetime.now().isoformat(timespec='seconds')
            }
            print(f"[ANALISE CRUZADA][{codigo_erro}]\n{detalhe_erro}")
            st.error(f"🚨 Não foi possível concluir a análise. Código de diagnóstico: {codigo_erro}")
            st.caption("O detalhe técnico foi registrado no log da aplicação para diagnóstico, sem expor informações internas na tela.")
