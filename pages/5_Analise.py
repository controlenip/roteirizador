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
    for col in aliases:
        try:
            val = row.get(col)
        except Exception:
            val = None
        if pd.notna(val) and str(val).strip().lower() not in ['', 'nan', 'none']:
            return str(val).strip()
    return default


def nota_valida(v):
    s = str(v).strip()
    return s.upper() not in ['', 'NAN', 'NONE', 'NULL', '-']


def limpar_nota_serie(s):
    return s.astype(str).str.replace('.0', '', regex=False).str.strip()


def criar_id_analise():
    return f"ANL-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"


def montar_config_txt(config, cores=None):
    linhas = [
        "CONFIGURAÇÃO DA ANÁLISE CRUZADA",
        f"ID da análise: {config.get('id_analise', '-')}",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Arquivo Saneamento: {config.get('arquivo_saneamento', '-')}",
        f"Arquivo Levantamento: {config.get('arquivo_levantamento', '-')}",
        f"Arquivo Localidades: {config.get('arquivo_localidades', '-')}",
        f"Raio de agrupamento: {config.get('raio_proximidade_m', '-')} m",
        f"Quantidade de equipes próximas: {config.get('qtd_equipes_proximas', '-')}",
        f"Distância máxima para equipes adicionais: {config.get('distancia_max_equipe_km', '-')} km",
        f"Linhas Saneamento recebidas: {config.get('linhas_saneamento', '-')}",
        f"Linhas Levantamento recebidas: {config.get('linhas_levantamento', '-')}",
        f"Linhas válidas espacialmente: {config.get('linhas_validas', '-')}",
        f"Coordenadas rejeitadas: {config.get('coordenadas_rejeitadas', '-')}",
        f"Coordenadas corrigidas: {config.get('coordenadas_corrigidas', '-')}",
    ]
    if cores is not None:
        linhas.append(f"Filtros de cores exportados: {', '.join(cores)}")
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
        df_temp = pd.read_csv(io.BytesIO(file_loc.getvalue()))
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


def colaboradores_proximos_em_lote(df, lat_locs, lon_locs, nomes_locs, tipos_locs, qtd_equipes, distancia_max_km):
    """Mantém a mesma regra de distância, reduzindo overhead de DataFrame.apply()."""
    resultados = []
    qtd_equipes = max(1, int(qtd_equipes))
    n_locs = len(lat_locs)
    if n_locs == 0:
        return ['DESCONHECIDO'] * len(df)

    coords = df[['LATITUDE', 'LONGITUDE']].to_numpy(dtype=float)
    for inicio in range(0, len(coords), 500):
        bloco = coords[inicio:inicio + 500]
        for lat, lon in bloco:
            if pd.isna(lat) or pd.isna(lon):
                resultados.append('DESCONHECIDO')
                continue

            dists = haversine_vectorized(lat, lon, lat_locs, lon_locs)
            if len(dists) == 0:
                resultados.append('DESCONHECIDO')
                continue

            k = min(qtd_equipes, len(dists))
            if k == len(dists):
                idxs = np.argsort(dists)
            else:
                idxs = np.argpartition(dists, k - 1)[:k]
                idxs = idxs[np.argsort(dists[idxs])]

            res = []
            for pos, i in enumerate(idxs):
                d_km = float(dists[i])
                # Preserva a regra anterior: a equipe mais próxima sempre aparece;
                # somente equipes adicionais obedecem ao limite de distância.
                if pos >= 1 and d_km > float(distancia_max_km):
                    break
                res.append(f"{str(nomes_locs[i]).title()} ({tipos_locs[i]}) - {d_km:.1f}km")
            resultados.append(" | ".join(res) if res else 'DESCONHECIDO')
    return resultados


def limpar_estado_analise():
    for chave in [
        'df_final_analise', 'is_done_analise', 'df_coord_rejeitadas_analise',
        'df_coord_corrigidas_analise', 'config_analise', 'bytes_excel_analise',
        'bytes_kml_analise', 'export_sig_excel_analise', 'export_sig_kml_analise'
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


# ==============================================================
# SIDEBAR
# ==============================================================
with st.sidebar:
    st.markdown("### ⚙️ Configurações da Análise")
    raio_prox = st.slider("Distância p/ agrupar obras (Metros)", 10, 1000, 100, 10, key='raio_prox_analise')
    qtd_equipes_prox = st.number_input("Qtd. de equipes mais próximas", min_value=1, max_value=5, value=2, step=1, key='qtd_equipes_prox_analise')
    distancia_max_equipe = st.number_input("Distância máx. p/ equipes adicionais (km)", min_value=1.0, max_value=1000.0, value=100.0, step=10.0, key='dist_max_equipe_analise')

    st.markdown("---")

    if st.session_state.is_done_analise and not st.session_state.df_final_analise.empty:
        df_fin = st.session_state.df_final_analise.copy()

        if 'COR_NOME' not in df_fin.columns:
            limpar_estado_analise()
            st.rerun()

        st.markdown("### 🎨 Filtro de Cores (Mapa e Export)")
        contagem_cores = df_fin['COR_NOME'].value_counts().to_dict()
        opcoes_cores = sorted(df_fin['COR_NOME'].dropna().unique().tolist())
        cores_selecionadas = st.multiselect(
            "Selecione os dados para visualizar:",
            opcoes_cores,
            default=opcoes_cores,
            format_func=lambda x: f"{x} ({int(contagem_cores.get(x, 0))})"
        )

        if not cores_selecionadas:
            st.warning("Selecione pelo menos uma cor para gerar o mapa e os relatórios.")
            st.stop()

        df_view = df_fin[df_fin['COR_NOME'].isin(cores_selecionadas)].copy()
        config_analise = st.session_state.get('config_analise', {})
        filtro_sig = tuple(sorted(cores_selecionadas))

        # Se o filtro mudar, os arquivos antigos deixam de representar a tela atual.
        if st.session_state.get('export_sig_excel_analise') != filtro_sig:
            st.session_state.pop('bytes_excel_analise', None)
        if st.session_state.get('export_sig_kml_analise') != filtro_sig:
            st.session_state.pop('bytes_kml_analise', None)

        st.markdown("---")
        d_fmt = datetime.now().strftime("%d.%m.%Y_%H%M")

        if 'bytes_excel_analise' not in st.session_state:
            if st.button("⚙️ Gerar Planilha Excel", use_container_width=True):
                with st.spinner("Gerando planilhas..."):
                    df_corrigidas = st.session_state.get('df_coord_corrigidas_analise', pd.DataFrame()).copy()
                    df_rejeitadas = st.session_state.get('df_coord_rejeitadas_analise', pd.DataFrame()).copy()

                    dict_dfs = {
                        'Consolidado (Todas)': df_view,
                        'Apenas Saneamento': df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'],
                        'Apenas Levantamento': df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO'],
                        'Notas Inválidas': df_view[df_view['COR_NOME'].astype(str).str.contains('Inválidas', na=False)],
                        'Notas Duplicadas': df_view[df_view['DUPLICADA'].astype(str).eq('SIM')] if 'DUPLICADA' in df_view.columns else pd.DataFrame(),
                        'Notas Próximas': df_view[df_view['PROXIMA'].astype(str).eq('SIM')] if 'PROXIMA' in df_view.columns else pd.DataFrame(),
                        'Notas Solitárias': df_view[df_view['PROXIMA'].astype(str).eq('NÃO')] if 'PROXIMA' in df_view.columns else pd.DataFrame(),
                        'Coordenadas Corrigidas': df_corrigidas,
                        'Coordenadas Rejeitadas': df_rejeitadas,
                    }
                    excel_bytes = gerar_excel_analise(dict_dfs)
                    bu_xl = io.BytesIO()
                    with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
                        zx.writestr("Planilha_Analise_Cruzada.xlsx", excel_bytes)
                        zx.writestr("Configuracao_Analise.txt", montar_config_txt(config_analise, cores_selecionadas).encode('utf-8'))
                    st.session_state.bytes_excel_analise = bu_xl.getvalue()
                    st.session_state.export_sig_excel_analise = filtro_sig
                st.rerun()
        else:
            st.download_button(
                "🌐 Baixar Planilha Excel (ZIP)",
                data=st.session_state.bytes_excel_analise,
                file_name=f"Analise_Planilhas_{d_fmt}.zip",
                use_container_width=True
            )

        if 'bytes_kml_analise' not in st.session_state:
            if st.button("⚙️ Gerar Mapa KML", use_container_width=True):
                with st.spinner("Gerando KML..."):
                    kml_str = gerar_kml_analise(df_view)
                    bu_kml = io.BytesIO()
                    with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
                        zk.writestr("Mapa_Analise_Cruzada.kml", kml_str.encode('utf-8'))
                        zk.writestr("Configuracao_Analise.txt", montar_config_txt(config_analise, cores_selecionadas).encode('utf-8'))
                    st.session_state.bytes_kml_analise = bu_kml.getvalue()
                    st.session_state.export_sig_kml_analise = filtro_sig
                st.rerun()
        else:
            st.download_button(
                "🗺️ Baixar Mapa (KML)",
                data=st.session_state.bytes_kml_analise,
                file_name=f"Analise_Mapa_{d_fmt}.zip",
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

    st.markdown("### 📈 Resumo da Volumetria")
    total_san = len(df_view[df_view['ORIGEM_BASE'] == 'SANEAMENTO'])
    total_lev = len(df_view[df_view['ORIGEM_BASE'] == 'LEVANTAMENTO'])

    col_a, col_b, col_c = st.columns(3)
    col_a.info(f"**🟣 Saneamento (Validado):** {total_san} obras")
    col_b.success(f"**🟢 Levantamento (Validado):** {total_lev} obras")
    col_c.warning(f"**🎯 Total Geral:** {len(df_view)} obras")

    # Painel de qualidade sem interferir em nenhuma classificação.
    invalidas = int(df_view['COR_NOME'].astype(str).str.contains('Inválidas', na=False).sum())
    duplicadas = int(df_view.get('DUPLICADA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    proximas = int(df_view.get('PROXIMA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    corrigidas = int(df_view.get('COORDENADA_CORRIGIDA', pd.Series(index=df_view.index, dtype='object')).astype(str).eq('SIM').sum())
    clusters = int(df_view['CLUSTER_ID'].nunique()) if 'CLUSTER_ID' in df_view.columns else len(df_view)
    municipios = int(df_view['MUNICIPIO'].dropna().nunique()) if 'MUNICIPIO' in df_view.columns else 0
    alertas_geo = int(df_view.get('COORDENADA_ALERTA', pd.Series(index=df_view.index, dtype='object')).astype(str).str.strip().ne('').sum())
    rejeitadas = len(st.session_state.get('df_coord_rejeitadas_analise', pd.DataFrame()))

    st.markdown("#### 🧪 Qualidade da Análise")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Notas inválidas", invalidas)
    q2.metric("Notas duplicadas", duplicadas)
    q3.metric("Notas próximas", proximas)
    q4.metric("Clusters", clusters)
    q5, q6, q7, q8 = st.columns(4)
    q5.metric("Municípios", municipios)
    q6.metric("Coords. corrigidas", corrigidas)
    q7.metric("Coords. rejeitadas", rejeitadas)
    q8.metric("Alertas geográficos", alertas_geo)

    if rejeitadas > 0:
        with st.expander(f"⚠️ {rejeitadas} registros com coordenadas rejeitadas", expanded=False):
            st.dataframe(st.session_state.df_coord_rejeitadas_analise, use_container_width=True, hide_index=True)

    if alertas_geo > 0:
        with st.expander(f"🌍 {alertas_geo} registros com coordenadas geograficamente atípicas", expanded=False):
            cols_alerta = [c for c in ['NOTA', 'ORIGEM_BASE', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'COORDENADA_ALERTA'] if c in df_view.columns]
            st.dataframe(df_view[df_view['COORDENADA_ALERTA'].astype(str).str.strip().ne('')][cols_alerta], use_container_width=True, hide_index=True)

    st.markdown("### 🗺️ Mapa Analítico")
    mostrar_mapa = st.checkbox("Exibir mapa analítico", value=False, key='mostrar_mapa_analise')
    if mostrar_mapa:
        mapa = folium.Map(location=[df_view['LATITUDE'].mean(), df_view['LONGITUDE'].mean()], zoom_start=8) if not df_view.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        m_clust = MarkerCluster(name="📍 Obras Filtradas").add_to(mapa)

        for cid, grp in df_view.groupby('CLUSTER_ID'):
            lat = grp['LATITUDE'].iloc[0]
            lon = grp['LONGITUDE'].iloc[0]
            c_names = grp['COR_NOME'].tolist()

            if any('Preto' in c for c in c_names) or any('Inválidas' in c for c in c_names):
                c_i = 'black'
            elif any('Vermelho' in c for c in c_names) or any('Duplicadas' in c for c in c_names):
                c_i = 'red'
            elif len(grp) > 1:
                c_i = 'orange'
            else:
                c_i = grp['COR_MAPA'].iloc[0]

            titulo_card = f"📍 Obras no Local ({len(grp)})"
            pop_html = f'''
            <div style="font-family:sans-serif; width:280px; max-height:280px; overflow-y:auto; border-radius:8px; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                <div style="background:#0D256C; color:#ffffff; padding:8px; font-size:13px; font-weight:bold; text-align:center; position:sticky; top:0;">{titulo_card}</div>
                <div style="padding:10px; background:#fafafa; font-size:12px;">
            '''

            for _, r in grp.iterrows():
                n = html.escape(str(r.get('NOTA', '')))
                mun = html.escape(str(r.get('MUNICIPIO', '')))
                o = html.escape(str(r.get('ORIGEM_BASE', '')))
                s = html.escape(str(r.get('SITUACAO SAP', '')))
                s_sisco = html.escape(valor_alias(r, ['STATUS SISCO', 'STATUS_SISCO']))
                s_list = html.escape(valor_alias(r, ['STATUS LIST', 'STATUS_LIST']))
                col = html.escape(str(r.get('COLABORADORES MAIS PROXIMOS', '')))
                dup = html.escape(str(r.get('DUPLICADA', '')))

                aviso_gps = ""
                if dup == 'SIM' and len(grp) == 1:
                    aviso_gps = "<br><span style='color:red; font-size:10px;'>⚠️ A cópia desta nota está em outro ponto geográfico.</span>"

                pop_html += f'''
                <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
                    <tr><td style="padding:2px;"><b>Nota:</b></td><td style="padding:2px;">{n}</td></tr>
                    <tr><td style="padding:2px;"><b>Município:</b></td><td style="padding:2px;">{mun}</td></tr>
                    <tr><td style="padding:2px;"><b>Origem:</b></td><td style="padding:2px;">{o}</td></tr>
                    <tr><td style="padding:2px;"><b>SAP:</b></td><td style="padding:2px;">{s}</td></tr>
                    <tr><td style="padding:2px;"><b>SISCO / LIST:</b></td><td style="padding:2px;">{s_sisco} / {s_list}</td></tr>
                    <tr><td style="padding:2px;"><b>Equipes Perto:</b></td><td style="padding:2px;">{col}</td></tr>
                    <tr><td style="padding:2px;"><b>Duplicada:</b></td><td style="padding:2px;">{dup}{aviso_gps}</td></tr>
                </table>
                <hr style="margin:4px 0; border:0; border-top:1px solid #ccc;">
                '''

            pop_html += '</div></div>'
            folium.Marker([lat, lon], icon=folium.Icon(color=c_i, icon='info-sign'), popup=folium.Popup(pop_html, max_width=320)).add_to(m_clust)

        folium.LayerControl().add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550)
    else:
        st.caption("O mapa é carregado apenas quando solicitado para manter a tela mais rápida.")

    st.markdown("### 📊 Tabela Consolidada Detalhada")
    st.dataframe(
        df_view.drop(columns=['_ORIGINAL_ROWS', 'LAT_NUM', 'LON_NUM', 'COR_MAPA', 'COR_NOME', 'CLUSTER_ID'], errors='ignore'),
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
            if file_san.name.lower().endswith('.csv'):
                df_san_raw = pd.read_csv(io.BytesIO(file_san.getvalue()))
            else:
                df_san_raw = ler_planilha_cached(file_san.getvalue())

            if file_lev.name.lower().endswith('.csv'):
                df_lev_raw = pd.read_csv(io.BytesIO(file_lev.getvalue()))
            else:
                df_lev_raw = ler_planilha_cached(file_lev.getvalue())

            df_san = preparar_base_obras(df_san_raw, 'SANEAMENTO')
            df_lev = preparar_base_obras(df_lev_raw, 'LEVANTAMENTO')

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

            df_loc['LAT_LOC'] = pd.to_numeric(df_loc['LAT_LOC'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            df_loc['LON_LOC'] = pd.to_numeric(df_loc['LON_LOC'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            df_loc = df_loc.dropna(subset=['LAT_LOC', 'LON_LOC']).copy()
            if df_loc.empty:
                st.error("❌ A base de Localidades não possui coordenadas válidas.")
                st.stop()

            lat_locs = df_loc['LAT_LOC'].to_numpy(dtype=float)
            lon_locs = df_loc['LON_LOC'].to_numpy(dtype=float)
            nomes_locs = df_loc['NOME_COLAB'].astype(str).to_numpy()
            tipos_locs = df_loc['TIPO_EQUIPE'].astype(str).to_numpy()

            render_t(0.36, "Validando Status SAP (Bloqueios)...")
            status_col = encontrar_coluna(df_lev, ['STATUS_SAP', 'STATUS SAP', 'STATUS'])
            status_dict = {}
            if status_col:
                tmp_status = df_lev[['NOTA', status_col]].copy()
                tmp_status = tmp_status[tmp_status['NOTA'].apply(nota_valida)]
                tmp_status['_STATUS'] = tmp_status[status_col].astype(str).str.strip().str.upper().str.replace('.0', '', regex=False)

                def consolidar_status(vals):
                    vals = [v for v in vals if v not in ['', 'NAN', 'NONE']]
                    bloqueados = []
                    if 'FINL' in vals:
                        bloqueados.append('FINL')
                    if 'CANC' in vals:
                        bloqueados.append('CANC')
                    if bloqueados:
                        return '/'.join(bloqueados)
                    return vals[-1] if vals else ''

                status_dict = tmp_status.groupby('NOTA')['_STATUS'].agg(consolidar_status).to_dict()

            def get_situacao(nota):
                s = str(status_dict.get(nota, '')).strip().upper()
                if 'FINL' in s or 'CANC' in s:
                    return f"BLOQUEADO ({s})"
                return "APTO"

            df_san['SITUACAO SAP'] = df_san['NOTA'].apply(get_situacao)
            df_lev['SITUACAO SAP'] = df_lev['NOTA'].apply(get_situacao)

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

            # Apenas auditoria: nenhuma dessas situações é excluída automaticamente.
            alertas = []
            for _, rr in df_valid.iterrows():
                a = []
                lat = float(rr['LATITUDE'])
                lon = float(rr['LONGITUDE'])
                if lat == 0.0 or lon == 0.0:
                    a.append('Coordenada zerada')
                if lat > 0 or lon > 0:
                    a.append('Coordenada positiva')
                if abs(lat) > abs(lon):
                    a.append('Possível LAT/LON invertida')
                alertas.append(' | '.join(a))
            df_valid['COORDENADA_ALERTA'] = alertas

            df_corrigidas = df_valid[df_valid['COORDENADA_CORRIGIDA'] == 'SIM'].copy()
            st.session_state.df_coord_rejeitadas_analise = df_rej_coord
            st.session_state.df_coord_corrigidas_analise = df_corrigidas

            if df_valid.empty:
                st.error("❌ Nenhum registro possui coordenadas válidas para a análise espacial.")
                st.stop()

            render_t(0.63, "Agrupando Obras Vizinhas e Processando Cores...")
            # Usa diretamente o valor selecionado no slider.
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

            df_final = pd.DataFrame(expanded)

            render_t(0.74, "Calculando colaboradores mais próximos...")
            df_final['COLABORADORES MAIS PROXIMOS'] = colaboradores_proximos_em_lote(
                df_final,
                lat_locs, lon_locs, nomes_locs, tipos_locs,
                qtd_equipes_prox,
                distancia_max_equipe
            )

            render_t(0.86, "Classificando regras e cores da auditoria...")

            def determinar_cor(linha):
                dupl = str(linha.get('DUPLICADA', ''))
                prox = str(linha.get('PROXIMA', ''))
                orig = str(linha.get('ORIGEM_BASE', ''))
                is_black = False

                st_sap = str(linha.get('SITUACAO SAP', '')).upper()
                if 'BLOQUEADO' in st_sap or 'FINL' in st_sap or 'CANC' in st_sap:
                    is_black = True

                if orig == 'LEVANTAMENTO':
                    st_list_raw = valor_alias(linha, ['STATUS_LIST', 'STATUS LIST'], '0').strip().upper().replace('.0', '')
                    if st_list_raw in ['NAN', 'NONE', '']:
                        st_list_raw = '0'
                    st_list = remover_acentos_str(st_list_raw)

                    st_sisco_raw = valor_alias(linha, ['STATUS_SISCO', 'STATUS SISCO'], '0').strip().upper().replace('.0', '')
                    if st_sisco_raw in ['NAN', 'NONE', '']:
                        st_sisco_raw = '0'
                    st_sisco = remover_acentos_str(st_sisco_raw)

                    v_list = ['0', 'EM LEVANTAMENTO', 'CORRECAO DE LEVANTAMENTO']
                    v_sisco = ['0', 'PRE ANALISE', 'LIBERADO PARA LEVANTAMENTO', 'LIBERADO P/ LEVANTAMENTO']

                    if st_list not in v_list or st_sisco not in v_sisco:
                        is_black = True

                if is_black:
                    return 'black', '⚫ Notas Inválidas'
                if dupl == 'SIM':
                    return 'red', '🔴 Notas Duplicadas'
                if prox == 'SIM':
                    return 'orange', '🟠 Notas Próximas'
                if orig == 'LEVANTAMENTO':
                    return 'green', '🟢 Notas Levantamento Solitárias'
                if orig == 'SANEAMENTO':
                    return 'purple', '🟣 Notas Saneamento Solitárias'
                return 'blue', '🔵 Outras'

            cores_calculadas = [determinar_cor(r) for _, r in df_final.iterrows()]
            df_final['COR_MAPA'] = [c[0] for c in cores_calculadas]
            df_final['COR_NOME'] = [c[1] for c in cores_calculadas]

            id_analise = criar_id_analise()
            df_final['ID_ANALISE'] = id_analise

            front_cols = [
                'ID_ANALISE', 'NOTA', 'ORIGEM_BASE', 'SITUACAO SAP', 'DUPLICADA',
                'DUPLICADA_INTERBASE', 'REPETIDA_NA_ORIGEM', 'PROXIMA',
                'COLABORADORES MAIS PROXIMOS', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE',
                'LATITUDE_ORIGINAL', 'LONGITUDE_ORIGINAL', 'COORDENADA_CORRIGIDA',
                'COORDENADA_ALERTA', 'CLUSTER_ID', 'COR_MAPA', 'COR_NOME'
            ]
            front_cols = [c for c in front_cols if c in df_final.columns]
            rest_cols = [c for c in df_final.columns if c not in front_cols and not c.startswith('_')]
            df_final = df_final[front_cols + rest_cols]

            config = {
                'id_analise': id_analise,
                'arquivo_saneamento': file_san.name,
                'arquivo_levantamento': file_lev.name,
                'arquivo_localidades': file_loc.name,
                'raio_proximidade_m': int(raio_prox),
                'qtd_equipes_proximas': int(qtd_equipes_prox),
                'distancia_max_equipe_km': float(distancia_max_equipe),
                'linhas_saneamento': int(len(df_san)),
                'linhas_levantamento': int(len(df_lev)),
                'linhas_validas': int(len(df_final)),
                'coordenadas_rejeitadas': int(len(df_rej_coord)),
                'coordenadas_corrigidas': int(len(df_corrigidas)),
                'duplicadas_interbase': int(df_final['DUPLICADA'].astype(str).eq('SIM').sum()),
                'repetidas_saneamento': int(len(rep_san)),
                'repetidas_levantamento': int(len(rep_lev)),
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
            st.error(f"🚨 Erro durante a análise: {e}")
