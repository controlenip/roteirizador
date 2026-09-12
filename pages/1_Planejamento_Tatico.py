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
import inspect
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Módulos Separados
from modules.data_processing import ler_planilha_cached, formata_campo_html, formatar_moeda, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas
from modules.export_tatica import injetar_logo, identificar_icone_folium, gerar_excel_tatica, limpar_colunas_tatica, gerar_kml_tatica, gerar_gpx_simples, gerar_excel_resumo_tatica

st.set_page_config(page_title="Roteirizador Tático", page_icon="🗺️", layout="wide")
injetar_logo()

@st.cache_data(show_spinner=False)
def carregar_arquivo_normalizado_tatica(file_bytes, file_name):
    """Lê e normaliza um arquivo apenas uma vez por conteúdo/nome durante a sessão/cache do Streamlit."""
    nome = str(file_name).lower()
    if nome.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = ler_planilha_cached(file_bytes)
    df = df.copy()
    df.columns = normalize_cols(df.columns)
    return df.loc[:, ~df.columns.duplicated()].copy()


@st.cache_data(show_spinner=False)
def preparar_demanda_arquivo_tatica(file_bytes, file_name, criar_id_generico=False):
    """Cacheia normalização, identificação de protocolo e explosão de notas compostas."""
    dft = carregar_arquivo_normalizado_tatica(file_bytes, file_name).copy()
    for cc in ['NOTA', 'PROTOCOLO', 'OS', 'ID']:
        if cc in dft.columns:
            dft['PROTOCOLO'] = dft[cc]
            break
    if 'PROTOCOLO' not in dft.columns and criar_id_generico:
        dft['PROTOCOLO'] = "GEN_" + dft.index.astype(str)
    if 'PROTOCOLO' in dft.columns:
        dft['PROTOCOLO'] = dft['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        dft = dft.explode('PROTOCOLO').reset_index(drop=True)
        dft['PROTOCOLO'] = dft['PROTOCOLO'].str.strip()
    return dft


@st.cache_data(show_spinner=False)
def preparar_bases_tatica_cache(b_t, selecionadas):
    """Guarda a base já filtrada, geocodificada e normalizada para evitar refazer o preparo a cada rerun."""
    df = b_t[b_t['BASE_NOME'].isin(list(selecionadas))].copy()
    if 'LATITUDE' in df.columns and 'LONGITUDE' in df.columns:
        df['LATITUDE'] = pd.to_numeric(df['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
        df['LONGITUDE'] = pd.to_numeric(df['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    elif 'RESIDENCIA' in df.columns or 'MUNICIPIO' in df.columns:
        cr = 'RESIDENCIA' if 'RESIDENCIA' in df.columns else 'MUNICIPIO'
        mc = {m: obter_coordenadas_municipio_cached(m) for m in df[cr].dropna().unique()}
        df['LATITUDE'] = df[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0])
        df['LONGITUDE'] = df[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
    df = df.dropna(subset=['LATITUDE', 'LONGITUDE'])
    cr_mun = 'MUNICIPIO' if 'MUNICIPIO' in df.columns else ('RESIDENCIA' if 'RESIDENCIA' in df.columns else None)
    if cr_mun:
        try:
            df['MUN_LIMPO_BASE'] = normalizar_municipios(df[cr_mun].astype(str)).astype(str).str.strip().str.upper()
        except Exception:
            df['MUN_LIMPO_BASE'] = df[cr_mun].astype(str).str.strip().str.upper()
    return df


def chamar_osrm_compativel(lat1, lon1, lat2, lon2, url_base, velocidade, timeout_s=12.0):
    """Usa timeout quando a implementação de routing_engine oferecer esse parâmetro, sem quebrar versões antigas."""
    try:
        params = inspect.signature(obter_rota_ruas).parameters
    except Exception:
        params = {}
    if 'timeout' in params:
        return obter_rota_ruas(lat1, lon1, lat2, lon2, url_base, velocidade, timeout=timeout_s)
    if 'timeout_s' in params:
        return obter_rota_ruas(lat1, lon1, lat2, lon2, url_base, velocidade, timeout_s=timeout_s)
    return obter_rota_ruas(lat1, lon1, lat2, lon2, url_base, velocidade)

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        if 'POSTE' in c.upper(): return str(int(float(v)))
        vf = float(v)
        if c.upper() in ['DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'DISTANCIA_RODOVIARIA_KM', 'DISTANCIA_ESTIMADA_KM']: 
            return f"{vf:.2f} KM"
        elif 'DISTANCIA' in c.upper(): 
            return f"{vf:.2f} Metros"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade):
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

def contar_obras_registro(registro):
    """Quantidade real de obras representadas por uma linha (inclui Super Ponto)."""
    orig = registro.get('_ORIGINAL_ROWS') if isinstance(registro, dict) else registro.get('_ORIGINAL_ROWS', None)
    return len(orig) if isinstance(orig, list) and len(orig) > 0 else 1


def chave_rota_cache(lat1, lon1, lat2, lon2):
    """Chave estável para reaproveitar consultas OSRM durante a mesma roteirização."""
    return (round(float(lat1), 5), round(float(lon1), 5), round(float(lat2), 5), round(float(lon2), 5))


def distancia_geometria_km(geom):
    """Calcula a distância do traçado retornado pelo OSRM, em vez de linha reta x fator."""
    if not isinstance(geom, list) or len(geom) < 2:
        return 0.0
    total = 0.0
    for a, b in zip(geom[:-1], geom[1:]):
        try:
            lon1, lat1 = float(a[0]), float(a[1])
            lon2, lat2 = float(b[0]), float(b[1])
            total += float(haversine_scalar(lat1, lon1, lat2, lon2))
        except Exception:
            continue
    return total


def municipio_normalizado(valor):
    """Aplica exatamente o mesmo normalizador usado nas demandas e nas bases."""
    try:
        return str(normalizar_municipios(pd.Series([str(valor)])).iloc[0]).strip().upper()
    except Exception:
        return str(valor).strip().upper()


def ordenar_bolsoes_diarios(obras, base_lat, base_lon, capacidade_dia, sentido, url_osrm):
    """
    Agrupa espacialmente antes de otimizar a ordem. Mantém a mesma cota diária e o mesmo
    motor TSP já usado pelo app; apenas evita que o corte do dia aconteça no meio de um bolsão.
    """
    if not obras:
        return []
    capacidade_dia = max(1, int(capacidade_dia))
    pendentes = [dict(o) for o in obras]
    saida = []
    reversa = "Varredura Reversa" in str(sentido)

    while pendentes:
        # Mantém prioridades no início e respeita o sentido operacional existente.
        prioritarias = [o for o in pendentes if str(o.get('PRIORIDADE', '')).strip().upper() == 'SIM']
        universo_semente = prioritarias if prioritarias else pendentes
        if reversa:
            semente = max(universo_semente, key=lambda o: haversine_scalar(base_lat, base_lon, float(o['LATITUDE']), float(o['LONGITUDE'])))
        else:
            semente = min(universo_semente, key=lambda o: haversine_scalar(base_lat, base_lon, float(o['LATITUDE']), float(o['LONGITUDE'])))

        grupo = [semente]
        pendentes.remove(semente)
        ocupacao = contar_obras_registro(semente)
        cl, cL = float(semente['LATITUDE']), float(semente['LONGITUDE'])

        while pendentes and ocupacao < capacidade_dia:
            candidatos = sorted(
                pendentes,
                key=lambda o: (
                    0 if str(o.get('PRIORIDADE', '')).strip().upper() == 'SIM' else 1,
                    haversine_scalar(cl, cL, float(o['LATITUDE']), float(o['LONGITUDE']))
                )
            )
            escolhido = None
            for cand in candidatos:
                peso = contar_obras_registro(cand)
                if ocupacao == 0 or ocupacao + peso <= capacidade_dia:
                    escolhido = cand
                    break
            if escolhido is None:
                break
            grupo.append(escolhido)
            pendentes.remove(escolhido)
            ocupacao += contar_obras_registro(escolhido)
            cl, cL = float(escolhido['LATITUDE']), float(escolhido['LONGITUDE'])

        # Dentro de cada bolsão, reaproveita a lógica original de ordenação.
        if reversa:
            restante = list(grupo)
            ordenado = []
            if restante:
                idx = max(range(len(restante)), key=lambda i: haversine_scalar(base_lat, base_lon, float(restante[i]['LATITUDE']), float(restante[i]['LONGITUDE'])))
                atual = restante.pop(idx)
                ordenado.append(atual)
                cl, cL = float(atual['LATITUDE']), float(atual['LONGITUDE'])
                while restante:
                    idx = min(range(len(restante)), key=lambda i: haversine_scalar(cl, cL, float(restante[i]['LATITUDE']), float(restante[i]['LONGITUDE'])))
                    atual = restante.pop(idx)
                    ordenado.append(atual)
                    cl, cL = float(atual['LATITUDE']), float(atual['LONGITUDE'])
            grupo = ordenado
        else:
            try:
                otimizado = resolver_tsp_ortools(grupo, base_lat, base_lon, url_osrm) if grupo else []
                if otimizado:
                    grupo = otimizado
            except Exception:
                pass
        saida.extend(grupo)
    return saida


def calcular_metricas_resultado_tatica(df):
    """Calcula os indicadores uma vez; pode ser refeito apenas quando as rotas mudarem."""
    if df is None or df.empty:
        return {'obras': 0, 'equipes': 0, 'super_pontos': 0, 'km_total': 0.0, 'osrm': 0, 'falhas': 0, 'estimados': 0}
    d = df.copy()
    d_t = d[~d['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])] if 'PROTOCOLO' in d.columns else d
    if 'DISTANCIA_RODOVIARIA_KM' in d.columns:
        km_real = pd.to_numeric(d['DISTANCIA_RODOVIARIA_KM'], errors='coerce')
        km_est = pd.to_numeric(d.get('DISTANCIA_PONTO_ANTERIOR_KM', 0), errors='coerce').fillna(0)
        km_total = float(km_real.fillna(km_est).fillna(0).sum())
    else:
        km_total = float(pd.to_numeric(d.get('DISTANCIA_PONTO_ANTERIOR_KM', 0), errors='coerce').fillna(0).sum())
    tsp = sum(1 for _, r in d_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)
    status = d.get('STATUS_ROTA', pd.Series(dtype=str)).astype(str)
    return {
        'obras': int(len(d_t)),
        'equipes': int(d['BASE_ATRIBUIDA'].nunique()) if 'BASE_ATRIBUIDA' in d.columns else 0,
        'super_pontos': int(tsp),
        'km_total': km_total,
        'osrm': int(status.str.startswith('OSRM').sum()) if not status.empty else 0,
        'falhas': int(status.str.contains('SEM_ROTA|TIMEOUT|ERRO', regex=True).sum()) if not status.empty else 0,
        'estimados': int(status.str.contains('ESTIMADA|LINHA_RETA', regex=True).sum()) if not status.empty else 0,
    }


def registrar_excedente_capacidade(registro):
    """Mantém compatibilidade com df_unallocated e separa excedentes de capacidade das falhas de atribuição."""
    r = dict(registro)
    r['MOTIVO_REJEICAO'] = 'Fora da Capacidade de Dias/Equipes'
    linha = pd.DataFrame([r])
    atual_cap = st.session_state.get('df_unallocated_capacity', pd.DataFrame())
    st.session_state.df_unallocated_capacity = pd.concat([atual_cap, linha], ignore_index=True)
    atual = st.session_state.get('df_unallocated', pd.DataFrame())
    st.session_state.df_unallocated = pd.concat([atual, linha], ignore_index=True)


def _montar_contexto_exportacao_tatica(df_routed):
    """Prepara dados comuns aos três pacotes somente quando algum download for solicitado."""
    df = df_routed.copy()
    if 'DISTANCIA_PROXIMO_PONTO_KM' not in df.columns:
        df['DISTANCIA_PROXIMO_PONTO_KM'] = df.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)

    linhas_gerais = []
    for _, r in df.iterrows():
        if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
            continue
        is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
        sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else "NÃO"
        if is_sp:
            for orig in r['_ORIGINAL_ROWS']:
                nr = r.copy()
                for k, v in orig.items():
                    if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'ROTA_GEOMETRIA', 'PERIODO', 'NOME_DIA', 'DIA_MES']:
                        nr[k] = v
                nr['SUPER_PONTO'] = sp_text
                linhas_gerais.append(nr)
        else:
            rn = r.copy(); rn['SUPER_PONTO'] = sp_text; linhas_gerais.append(rn)

    df_excel_full = pd.DataFrame(linhas_gerais)
    for c in df_excel_full.columns:
        if 'POSTE' in str(c).upper():
            df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

    col_exibir = list(st.session_state.get('colunas_exibir', []))
    if 'NOME_DIA' not in col_exibir: col_exibir.insert(0, 'NOME_DIA')
    if 'DIA_MES' not in col_exibir: col_exibir.insert(1, 'DIA_MES')
    if 'SUPER_PONTO' not in col_exibir: col_exibir.insert(2, 'SUPER_PONTO')
    tpc_exp = st.session_state.get('vrp_state', {}).get('config', {}).get('tipo_periodo', 'Semana')
    return df, df_excel_full, col_exibir, tpc_exp


def gerar_zip_planilhas_sob_demanda(df_routed):
    df, df_excel_full, col_exibir, _ = _montar_contexto_exportacao_tatica(df_routed)
    d_fmt = st.session_state.get('d_fmt_resultado_tatica', datetime.now().strftime('%d.%m.%Y'))
    bu = io.BytesIO()
    with zipfile.ZipFile(bu, 'w', zipfile.ZIP_DEFLATED) as zx:
        obras_por_dia_est = st.session_state.get('vrp_state', {}).get('config', {}).get('obras_por_dia', 4.0)
        res = []
        for b in df['BASE_ATRIBUIDA'].dropna().unique():
            db = df[(df['BASE_ATRIBUIDA'] == b) & (~df['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
            qtd_obras, qtd_postes = 0, 0.0
            for _, r in db.iterrows():
                origs = r.get('_ORIGINAL_ROWS') if isinstance(r.get('_ORIGINAL_ROWS'), list) else None
                registros = origs if origs else [r]
                qtd_obras += len(registros)
                for orig in registros:
                    vals = []
                    for k, v in orig.items():
                        if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                            try:
                                vf = float(v)
                                if vf > 0: vals.append(vf)
                            except Exception:
                                pass
                    if vals: qtd_postes += min(vals)
            postes_dia = (qtd_postes / (qtd_obras / float(obras_por_dia_est))) if qtd_obras > 0 else 0
            postes_semana = postes_dia * 5.0
            db_all = df[df['BASE_ATRIBUIDA'] == b]
            if 'DISTANCIA_RODOVIARIA_KM' in db_all.columns:
                km_total = pd.to_numeric(db_all['DISTANCIA_RODOVIARIA_KM'], errors='coerce').fillna(pd.to_numeric(db_all['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce')).fillna(0).sum()
            else:
                km_total = pd.to_numeric(db_all['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce').fillna(0).sum()
            status = db_all.get('STATUS_ROTA', pd.Series(dtype=str)).astype(str)
            res.append({
                'Equipe': b,
                'Obras Roteirizadas': qtd_obras,
                'Postes/Dia (Est.)': int(round(postes_dia)),
                'Postes/Semana (Est.)': int(round(postes_semana)),
                'Postes Total': int(round(qtd_postes)),
                'Super Pontos': sum(1 for _, r_sp in db.iterrows() if isinstance(r_sp.get('_ORIGINAL_ROWS'), list) and len(r_sp.get('_ORIGINAL_ROWS')) > 1),
                'Prioridades Atendidas': len(db[db['PRIORIDADE'] == 'Sim']) if 'PRIORIDADE' in db.columns else 0,
                'KM Total Previsto': round(float(km_total), 2),
                'Trechos OSRM': int(status.str.startswith('OSRM').sum()),
                'Trechos sem rota': int(status.str.contains('SEM_ROTA', regex=False).sum())
            })
        zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_tatica(pd.DataFrame(res)))

        dfc = st.session_state.get('df_correcao_tatica', pd.DataFrame())
        if not dfc.empty:
            dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
            dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
            for cc in dfcc.columns:
                if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
            out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())

        if 'STATUS_ROTA' in df.columns:
            dff = df[df['STATUS_ROTA'].astype(str).str.contains('SEM_ROTA|TIMEOUT|ERRO', regex=True)].copy()
            if not dff.empty:
                cols_falha = [c for c in ['BASE_ATRIBUIDA', 'PROTOCOLO', 'ORDEM', 'DIA_MES', 'LATITUDE', 'LONGITUDE', 'DISTANCIA_ESTIMADA_KM', 'STATUS_ROTA'] if c in dff.columns]
                out_f = io.BytesIO(); dff[cols_falha].to_excel(out_f, index=False); zx.writestr(f"Trechos_Falha_OSRM - {d_fmt}.xlsx", out_f.getvalue())

        for estado_key, nome in [('df_unallocated_area', 'Obras_Nao_Alocadas_Area'), ('df_unallocated_capacity', 'Obras_Excedentes_Capacidade')]:
            dfx = st.session_state.get(estado_key, pd.DataFrame())
            if not dfx.empty:
                out_e = io.BytesIO(); dfx.to_excel(out_e, index=False); zx.writestr(f"{nome} - {d_fmt}.xlsx", out_e.getvalue())

        if not df_excel_full.empty:
            dfg_total = limpar_colunas_tatica(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
            dfg_total = dfg_total.loc[:, ~dfg_total.columns.duplicated()].copy()
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Tatica_Total - {d_fmt}.xlsx", gerar_excel_tatica(dfg_total, st.session_state.get('colunas_originais_tat', [])))
            for base in df['BASE_ATRIBUIDA'].dropna().unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == base]
                if df_base_excel.empty: continue
                dfg = limpar_colunas_tatica(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
                dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
                for cc in dfg.columns:
                    if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                zx.writestr(f"Rotas_{d_fmt}/Rota_{b_safe}.xlsx", gerar_excel_tatica(dfg, st.session_state.get('colunas_originais_tat', [])))
    return bu.getvalue()


def gerar_zip_kml_sob_demanda(df_routed):
    df, _, col_exibir, tpc_exp = _montar_contexto_exportacao_tatica(df_routed)
    d_fmt = st.session_state.get('d_fmt_resultado_tatica', datetime.now().strftime('%d.%m.%Y'))
    bu = io.BytesIO()
    with zipfile.ZipFile(bu, 'w', zipfile.ZIP_DEFLATED) as zk:
        dfk_total = df[~df['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
        if not dfk_total.empty:
            dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS']) > 1 else "NÃO", axis=1)
            bases = df['BASE_ATRIBUIDA'].dropna().unique().tolist()
            ks_tot = gerar_kml_tatica(dfk_total, "ROTA TOTAL", col_exibir, bases, tpc_exp, formatar_valor_coluna)
            zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks_tot.encode('utf-8'))
            for base in bases:
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                dfk_base = dfk_total[dfk_total['BASE_ATRIBUIDA'] == base].copy()
                if not dfk_base.empty:
                    ks = gerar_kml_tatica(dfk_base, f"Rota {b_safe}", col_exibir, [base], tpc_exp, formatar_valor_coluna)
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kml", ks.encode('utf-8'))
    return bu.getvalue()


def gerar_zip_gpx_sob_demanda(df_routed):
    df, _, _, _ = _montar_contexto_exportacao_tatica(df_routed)
    d_fmt = st.session_state.get('d_fmt_resultado_tatica', datetime.now().strftime('%d.%m.%Y'))
    bu = io.BytesIO()
    with zipfile.ZipFile(bu, 'w', zipfile.ZIP_DEFLATED) as zg:
        dfk_total = df[~df['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
        if not dfk_total.empty:
            zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))
            for base in df['BASE_ATRIBUIDA'].dropna().unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                dfk_base = dfk_total[dfk_total['BASE_ATRIBUIDA'] == base].copy()
                if not dfk_base.empty:
                    zg.writestr(f"GPX_{d_fmt}/Rota_{b_safe}.gpx", gerar_gpx_simples(dfk_base, f"Rota {b_safe}").encode('utf-8'))
    return bu.getvalue()


def reprocessar_falhas_osrm(df_routed):
    """Refaz somente trechos sinalizados como falha, sem recalcular a distribuição ou a ordem das obras."""
    if df_routed is None or df_routed.empty or 'STATUS_ROTA' not in df_routed.columns:
        return df_routed, 0, 0
    df = df_routed.copy()
    cfg = st.session_state.get('vrp_state', {}).get('config', {})
    bases_df = pd.DataFrame(st.session_state.get('bases_records', []))
    if bases_df.empty:
        return df, 0, int(df['STATUS_ROTA'].astype(str).str.contains('SEM_ROTA', regex=False).sum())
    cache = st.session_state.setdefault('osrm_cache_tatica_global', {})
    falhas_total = int(df['STATUS_ROTA'].astype(str).str.contains('SEM_ROTA', regex=False).sum())
    if falhas_total == 0:
        return df, 0, 0
    pb = st.progress(0.0); msg = st.empty(); processadas = 0; recuperadas = 0
    for equipe, idxs in df.groupby('BASE_ATRIBUIDA').groups.items():
        base_rows = bases_df[bases_df['BASE_NOME'] == equipe]
        if base_rows.empty: continue
        prev_lat, prev_lon = float(base_rows.iloc[0]['LATITUDE']), float(base_rows.iloc[0]['LONGITUDE'])
        ordem_idxs = sorted(list(idxs), key=lambda ix: float(df.at[ix, 'ORDEM']) if pd.notna(df.at[ix, 'ORDEM']) else 0)
        for ix in ordem_idxs:
            row = df.loc[ix]
            destino_lat, destino_lon = row.get('LATITUDE'), row.get('LONGITUDE')
            if pd.isna(destino_lat) or pd.isna(destino_lon):
                continue
            status = str(row.get('STATUS_ROTA', ''))
            if 'SEM_ROTA' in status and row.get('PROTOCOLO') != 'PAUSA_ALMOCO':
                processadas += 1
                msg.info(f"🔄 Reprocessando trechos com falha... {processadas}/{falhas_total}")
                chave = chave_rota_cache(prev_lat, prev_lon, destino_lat, destino_lon)
                cached = cache.get(chave)
                rota = cached[0] if cached is not None else None
                meta = dict(cached[1]) if cached is not None else None
                if rota is None or not rota[0]:
                    for tentativa in range(3):
                        try:
                            ini = time.time()
                            rr = chamar_osrm_compativel(prev_lat, prev_lon, destino_lat, destino_lon, cfg.get('url_osrm_base', 'http://router.project-osrm.org'), cfg.get('velocidade_media_kmh', 30.0), timeout_s=12.0)
                            if rr and len(rr) >= 2 and isinstance(rr[0], list) and len(rr[0]) >= 2:
                                geom, dur = rr[0], rr[1]
                                km_real = distancia_geometria_km(geom)
                                status_ok = 'OSRM_LENTO' if (time.time() - ini) > 12.0 else 'OSRM'
                                meta = {'status': status_ok, 'km_real': km_real if km_real > 0 else np.nan, 'km_estimado': float(row.get('DISTANCIA_ESTIMADA_KM', row.get('DISTANCIA_PONTO_ANTERIOR_KM', 0)) or 0), 'tempo_real_min': float(dur) / 60.0 if pd.notna(dur) else np.nan}
                                rota = (geom, dur); cache[chave] = (rota, dict(meta)); break
                        except Exception:
                            time.sleep(1.0 + tentativa)
                if rota is not None and rota[0] and meta:
                    df.at[ix, 'ROTA_GEOMETRIA'] = rota[0]
                    df.at[ix, 'STATUS_ROTA'] = meta.get('status', 'OSRM')
                    df.at[ix, 'DISTANCIA_RODOVIARIA_KM'] = round(float(meta.get('km_real')), 2) if pd.notna(meta.get('km_real')) else np.nan
                    if pd.notna(meta.get('km_real')): df.at[ix, 'DISTANCIA_PONTO_ANTERIOR_KM'] = round(float(meta.get('km_real')), 2)
                    df.at[ix, 'TEMPO_ROTA_REAL_MIN'] = round(float(meta.get('tempo_real_min')), 1) if pd.notna(meta.get('tempo_real_min')) else np.nan
                    recuperadas += 1
                pb.progress(min(1.0, processadas / max(1, falhas_total)))
            prev_lat, prev_lon = float(destino_lat), float(destino_lon)
    st.session_state.osrm_cache_tatica_global = cache
    restantes = int(df['STATUS_ROTA'].astype(str).str.contains('SEM_ROTA', regex=False).sum())
    pb.empty(); msg.empty()
    return df, recuperadas, restantes


def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    # O cache global de rotas OSRM é intencionalmente preservado durante a sessão.
    st.session_state.update({'roteamento_concluido': False, 'vrp_status': "IDLE", 'vrp_state': {}, 'df_routed': pd.DataFrame(), 'bases_records': [], 'colunas_exibir': [], 'colunas_originais_tat': []})
    for k in [
        'bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg',
        'df_unallocated', 'df_unallocated_area', 'df_unallocated_capacity', 'df_correcao_tatica',
        'qtd_coords_autocorrigidas', 'metricas_resultado_tatica', 'd_fmt_resultado_tatica',
        'prep_tasks_tatica', 'prep_bases_tatica', 'prep_params_tatica', 'inicio_em_processamento_tatica',
        'mostrar_mapa_tatico', 'checkpoint_tatico'
    ]:
        st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    gc.collect()
    tentar_rerun()

if "roteamento_concluido" not in st.session_state: st.session_state.roteamento_concluido = False
if "vrp_status" not in st.session_state: st.session_state.vrp_status = "IDLE"
if "osrm_cache_tatica_global" not in st.session_state: st.session_state.osrm_cache_tatica_global = {}
if "mostrar_mapa_tatico" not in st.session_state: st.session_state.mostrar_mapa_tatico = False

status_exec = st.session_state.vrp_status
is_done = st.session_state.roteamento_concluido
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>🗺️ Planejamento Tático (Operação)</h1>", unsafe_allow_html=True)
st.info("💡 Distribui as obras equitativamente entre as equipes, criando rotas circulares diárias e priorizando notas urgentes.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Capacidade e Prazos", expanded=True):
        obras_dia = st.number_input("Cota Diária por Equipe:", min_value=1, value=6, disabled=is_locked)
        tpc = st.radio("Visão de Trabalho:", ["Dia", "Semana"], index=1, disabled=is_locked)
        limite_per = st.number_input(f"Qtd de {tpc}s de Rota:", min_value=1, value=1, disabled=is_locked)
        dias_sel = st.multiselect("Dias Úteis na Semana:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked)
        data_ini = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        
        st.markdown("---")
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)

    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed.empty:
        d_fmt = st.session_state.get('d_fmt_resultado_tatica', datetime.now().strftime("%d.%m.%Y"))

        if st.session_state.get('bytes_zip_xl'):
            st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.bytes_zip_xl, file_name=f"Rotas_Planilhas - {d_fmt}.zip", use_container_width=True)
        else:
            if st.button("🌐 Preparar Planilhas (ZIP)", use_container_width=True):
                with st.spinner("Gerando planilhas somente agora..."):
                    st.session_state.bytes_zip_xl = gerar_zip_planilhas_sob_demanda(st.session_state.df_routed)
                tentar_rerun()

        if st.session_state.get('bytes_zip_kml'):
            st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.bytes_zip_kml, file_name=f"Rotas_Mapas - {d_fmt}.zip", use_container_width=True)
        else:
            if st.button("🗺️ Preparar Mapas (KML)", use_container_width=True):
                with st.spinner("Gerando KML somente agora..."):
                    st.session_state.bytes_zip_kml = gerar_zip_kml_sob_demanda(st.session_state.df_routed)
                tentar_rerun()

        if st.session_state.get('bytes_zip_gpx'):
            st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.bytes_zip_gpx, file_name=f"Rotas_GPS - {d_fmt}.zip", use_container_width=True)
        else:
            if st.button("🛰️ Preparar GPS (GPX)", use_container_width=True):
                with st.spinner("Gerando GPX somente agora..."):
                    st.session_state.bytes_zip_gpx = gerar_zip_gpx_sob_demanda(st.session_state.df_routed)
                tentar_rerun()

        st.caption("Os pacotes são gerados sob demanda para o resultado aparecer mais rápido.")
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    df_c = st.session_state.get('df_correcao_tatica', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica:</b> Estas obras apresentaram coordenadas em branco, zeradas ou invertidas.
            </p>
        </div>
        """, unsafe_allow_html=True)

    # Trabalha em cópia para não alterar o DataFrame oficial da sessão durante a visualização.
    dfr = st.session_state.df_routed.copy()
    dfr['DISTANCIA_PROXIMO_PONTO_KM'] = dfr.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]

    metricas = st.session_state.get('metricas_resultado_tatica')
    if not metricas:
        metricas = calcular_metricas_resultado_tatica(dfr)
        st.session_state.metricas_resultado_tatica = metricas
    tk = f"{metricas['km_total']:.1f} km"

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Planejadas", metricas['obras'], "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", metricas['equipes'], "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(metricas['super_pontos']), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    df_area = st.session_state.get('df_unallocated_area', pd.DataFrame())
    df_cap = st.session_state.get('df_unallocated_capacity', pd.DataFrame())
    if not df_area.empty:
        st.warning(f"⚠️ {len(df_area)} obras não foram atribuídas por área/município sem equipe compatível.")
    if not df_cap.empty:
        st.warning(f"⚠️ {len(df_cap)} obras excederam a capacidade configurada de dias/equipes.")
    if df_area.empty and df_cap.empty:
        st.success("✅ 100% das obras foram alocadas com sucesso.")

    # Diagnóstico adicional sem alterar o fluxo original da tela.
    with st.expander("🧠 Diagnóstico da Roteirização", expanded=False):
        status_series = dfr.get('STATUS_ROTA', pd.Series(dtype=str)).astype(str) if 'STATUS_ROTA' in dfr.columns else pd.Series(dtype=str)
        qtd_osrm = int(status_series.str.startswith('OSRM').sum()) if not status_series.empty else 0
        qtd_falha = int(status_series.str.contains('SEM_ROTA|TIMEOUT|ERRO', regex=True).sum()) if not status_series.empty else 0
        qtd_estimado = int(status_series.str.contains('ESTIMADA|LINHA_RETA', regex=True).sum()) if not status_series.empty else 0
        qtd_corr = int(st.session_state.get('qtd_coords_autocorrigidas', 0))
        cargas = dfr_t.groupby('BASE_ATRIBUIDA').size() if not dfr_t.empty else pd.Series(dtype=float)
        cdiag1, cdiag2, cdiag3, cdiag4 = st.columns(4)
        cdiag1.metric("Trechos OSRM", qtd_osrm)
        cdiag2.metric("Trechos sem rota", qtd_falha)
        cdiag3.metric("Trechos estimados", qtd_estimado)
        cdiag4.metric("Coords. autocorrigidas", qtd_corr)
        if not cargas.empty:
            st.caption(f"Balanceamento por equipe — mín.: {int(cargas.min())} | média: {cargas.mean():.1f} | máx.: {int(cargas.max())}")
        if qtd_falha > 0:
            st.warning("Existem trechos em que o servidor de arruamento não retornou geometria. Eles não são desenhados como linha reta; permanecem sinalizados para conferência.")
            if st.button("🔄 Reprocessar somente trechos com falha", use_container_width=True):
                with st.spinner("Tentando recuperar somente os trechos que falharam..."):
                    novo_df, recuperadas, restantes = reprocessar_falhas_osrm(st.session_state.df_routed)
                    st.session_state.df_routed = novo_df
                    st.session_state.metricas_resultado_tatica = calcular_metricas_resultado_tatica(novo_df)
                    for k in ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx']:
                        st.session_state.pop(k, None)
                st.success(f"Trechos recuperados: {recuperadas}. Falhas restantes: {restantes}.")
                time.sleep(0.6)
                tentar_rerun()

    st.markdown("### 🗺️ Mapa Operacional")
    mostrar_mapa = st.checkbox("Carregar/Exibir mapa operacional", key='mostrar_mapa_tatico')
    if not mostrar_mapa:
        st.info("O mapa ficou sob demanda para acelerar a abertura do resultado. Ative a opção acima quando quiser visualizá-lo.")
    else:
        mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
        equipes_ordem = dfr['BASE_ATRIBUIDA'].dropna().unique().tolist()
        for bn in equipes_ordem:
            cr = co_f[equipes_ordem.index(bn) % len(co_f)]
            db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
            bn_safe = str(bn).replace("{", "[").replace("}", "]")
            fg = folium.FeatureGroup(name=f"Equipe: {bn_safe}", show=False)
            m_clust_eq = MarkerCluster(name=f"Obras - {bn_safe}").add_to(fg)
            for pe in db['PERIODO'].unique():
                dp = db[db['PERIODO'] == pe]
                for _, r_linha in dp.iterrows():
                    geom = r_linha.get('ROTA_GEOMETRIA')
                    if isinstance(geom, list) and len(geom) >= 2:
                        pts = [[pt[1], pt[0]] for pt in geom if isinstance(pt, (list, tuple)) and len(pt) >= 2]
                        if len(pts) >= 2:
                            folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
                            folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
                for r in dp.to_dict('records'):
                    if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                    c_i = 'red' if str(r.get('PRIORIDADE')) == 'Sim' else 'blue'
                    ic = identificar_icone_folium(r, dfr.columns)
                    er = "".join([f"<tr><td><b>{html.escape(c)}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
                    status_rota = html.escape(str(r.get('STATUS_ROTA', '-')))
                    pop_html = f'<div style="width:250px;"><b>Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><b>Status rota:</b> {status_rota}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
                    pop_html = pop_html.replace("{", "&#123;").replace("}", "&#125;")
                    folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust_eq)
            fg.add_to(mapa)
        folium.LayerControl().add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550)

elif status_exec == "IDLE":
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown("### 👥 1. Bases de Equipes")
        df_bases = pd.DataFrame()
        bf = st.file_uploader("Suba a planilha de Equipes (Excel)", type=["xlsx", "xls"])
        if bf:
            b_t = carregar_arquivo_normalizado_tatica(bf.getvalue(), bf.name)
            
            for pn in ['LEVANTADOR', 'FISCAL', 'NOME_FISCAL', 'NOME', 'TECNICO', 'COLABORADOR', 'EQUIPE']:
                if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'BASE_NOME'}); break
            
            if 'BASE_NOME' in b_t.columns:
                b_t['BASE_NOME'] = b_t['BASE_NOME'].astype(str).str.split(r'\s*\|\s*')
                b_t = b_t.explode('BASE_NOME').reset_index(drop=True)
                b_t['BASE_NOME'] = b_t['BASE_NOME'].str.strip().str.upper()
                
                opts = sorted([str(x) for x in b_t['BASE_NOME'].dropna().unique() if str(x) not in ['SEM EQUIPE', 'NAN', 'NONE', '']])
                sel = st.multiselect("Selecione as Equipes Ativas:", opts, default=opts)
                if sel:
                    with st.spinner("🌍 Preparando bases..."):
                        df_bases = preparar_bases_tatica_cache(b_t, tuple(sel))
            else: st.error("❌ A planilha não possui coluna de nome da Equipe/Levantador.")

        st.markdown("##### 📍 Regra de Atribuição")
        ta = st.radio("Como amarrar as notas aos técnicos?", ["Por Proximidade Espacial", "Por Município Base"], index=1, label_visibility="collapsed")
        st.caption("A proximidade empurra as obras para a equipe mais próxima no raio geral. Por Município isola a equipe.")

    with c_up2:
        st.markdown("### 📁 2. Demandas (Obras)")
        task_files = st.file_uploader("Suba as planilhas de Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📁 3. Planilha Genérica (Opcional)")
        st.info("💡 Utilize este campo extra caso precise roteirizar planilhas auxiliares que não sigam o padrão do sistema. Você mapeará a prioridade e o status manualmente.")
        generic_files = st.file_uploader("Suba uma Planilha Genérica", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
        
    if df_bases.empty or (not task_files and not generic_files): st.stop()
    
    qtd_eq = df_bases['BASE_NOME'].nunique()
    cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per
    sb_html.markdown(render_sidebar_card(cm, 0, qtd_eq, cm * qtd_eq), unsafe_allow_html=True)

    # ------------------ BLOCO OBRAS PADRÃO ------------------
    df_tasks_padrao = pd.DataFrame()
    if task_files:
        dfs = []
        for f in task_files:
            raw_cols = carregar_arquivo_normalizado_tatica(f.getvalue(), f.name).columns.tolist()
            dft = preparar_demanda_arquivo_tatica(f.getvalue(), f.name, criar_id_generico=False)
            if not dfs: st.session_state.colunas_originais_tat = raw_cols
            dfs.append(dft)
        df_tasks_padrao = pd.concat(dfs, ignore_index=True)

        st.markdown("---")
        st.markdown("#### ⚙️ Filtros - Demandas (Obras Padrão)")
        c1_f, c2_f = st.columns([1, 1])
        with c1_f:
            cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks_padrao.columns else 'STATUS DA FISCALIZAÇÃO'
            if cs in df_tasks_padrao.columns:
                df_tasks_padrao[cs] = df_tasks_padrao[cs].astype(str).str.strip().str.upper()
                opts_s = sorted([str(x) for x in df_tasks_padrao[cs].unique() if str(x) != 'NAN'])
                sel_s = st.multiselect("1. Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
                if not sel_s: st.stop()
                df_tasks_padrao = df_tasks_padrao[df_tasks_padrao[cs].isin(sel_s)].copy()
                
        with c2_f:
            if 'TIPO NOTA' in df_tasks_padrao.columns:
                df_tasks_padrao['TIPO NOTA'] = df_tasks_padrao['TIPO NOTA'].astype(str).str.strip().str.upper()
                opts_n = sorted([str(x) for x in df_tasks_padrao['TIPO NOTA'].unique() if str(x) != 'NAN'])
                sel_p = st.multiselect("🚨 2. Obras de Alta Prioridade:", options=opts_n, default=[n for n in opts_n if n in ['ASC', 'CCF', 'DIF', 'MGD', 'MTP', 'SID']])
                df_tasks_padrao['PRIORIDADE'] = np.where(df_tasks_padrao['TIPO NOTA'].isin(sel_p), 'Sim', 'Não')
            else: df_tasks_padrao['PRIORIDADE'] = 'Não'

    # ------------------ BLOCO PLANILHA GENÉRICA ------------------
    df_tasks_gen = pd.DataFrame()
    if generic_files:
        dfs_gen = []
        for f in generic_files:
            raw_cols = carregar_arquivo_normalizado_tatica(f.getvalue(), f.name).columns.tolist()
            dft = preparar_demanda_arquivo_tatica(f.getvalue(), f.name, criar_id_generico=True)
            if 'colunas_originais_tat' not in st.session_state or not st.session_state.colunas_originais_tat:
                st.session_state.colunas_originais_tat = raw_cols
            else:
                for col in raw_cols:
                    if col not in st.session_state.colunas_originais_tat:
                        st.session_state.colunas_originais_tat.append(col)
            dfs_gen.append(dft)
        df_tasks_gen = pd.concat(dfs_gen, ignore_index=True)

        st.markdown("---")
        st.markdown("#### ⚙️ Filtros - Planilha Genérica")
        colunas_gen = df_tasks_gen.columns.tolist()
        c1_g, c2_g = st.columns(2)
        with c1_g:
            col_prio = st.selectbox("🎯 1. Escolha a coluna de Prioridade (Genérica):", ["Nenhuma"] + colunas_gen)
            if col_prio != "Nenhuma":
                df_tasks_gen[col_prio] = df_tasks_gen[col_prio].astype(str).str.strip().str.upper()
                opts_p = sorted([str(x) for x in df_tasks_gen[col_prio].unique() if str(x) != 'NAN'])
                sel_p_gen = st.multiselect("🚨 Escolha os Valores de Alta Prioridade:", options=opts_p)
                df_tasks_gen['PRIORIDADE'] = np.where(df_tasks_gen[col_prio].isin(sel_p_gen), 'Sim', 'Não')
            else:
                df_tasks_gen['PRIORIDADE'] = 'Não'
                
        with c2_g:
            col_status = st.selectbox("🚦 2. Escolha a coluna de Status (Filtro Opcional):", ["Nenhuma"] + colunas_gen)
            if col_status != "Nenhuma":
                df_tasks_gen[col_status] = df_tasks_gen[col_status].astype(str).str.strip().str.upper()
                opts_s = sorted([str(x) for x in df_tasks_gen[col_status].unique() if str(x) != 'NAN'])
                sel_s_gen = st.multiselect("✅ Valores Roteirizáveis (Aptos):", options=opts_s)
                if sel_s_gen:
                    df_tasks_gen = df_tasks_gen[df_tasks_gen[col_status].isin(sel_s_gen)].copy()
                else:
                    st.warning("Selecione os status aptos para roteirizar a planilha genérica.")
                    st.stop()

    # ------------------ FUSÃO DAS DUAS FONTES ------------------
    df_tasks = pd.concat([df_tasks_padrao, df_tasks_gen], ignore_index=True)
    if df_tasks.empty:
        st.error("🚨 Nenhuma obra ou linha válida restou após aplicar os filtros.")
        st.stop()

    st.markdown("---")
    cg1, cg2 = st.columns([4, 1])
    with cg1: st.markdown("#### 🌍 Limpeza Geográfica")
    with cg2:
        if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    falta = [c for c in ['MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas obrigatórias nas planilhas: {', '.join(falta)}."); st.stop()

    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    m_m = df_tasks['MUNICIPIO'].isna() | (df_tasks['MUNICIPIO'].astype(str).str.strip() == '') | (df_tasks['MUNICIPIO'].astype(str).str.strip().str.upper() == 'NAN')
    if m_m.sum() > 0:
        df_tasks.loc[m_m, 'MOTIVO_REJEICAO'] = 'Município Vazio'
        df_rej = pd.concat([df_rej, df_tasks[m_m].copy()], ignore_index=True); df_tasks = df_tasks[~m_m].copy()

    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['COORD_CORRIGIDA'] = 'NÃO'
    m_na = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna()
    m_0 = (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)

    # Autocorrige somente inversões inequivocamente plausíveis no território brasileiro.
    m_inv_candidata = (abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])) & ~m_na & ~m_0 & ~m_pos
    lat_trocada = df_tasks['LON_NUM']
    lon_trocada = df_tasks['LAT_NUM']
    m_swap_valido = m_inv_candidata & lat_trocada.between(-35.0, 6.0) & lon_trocada.between(-75.0, -30.0)
    if m_swap_valido.any():
        lat_ant = df_tasks.loc[m_swap_valido, 'LAT_NUM'].copy()
        df_tasks.loc[m_swap_valido, 'LAT_NUM'] = df_tasks.loc[m_swap_valido, 'LON_NUM'].values
        df_tasks.loc[m_swap_valido, 'LON_NUM'] = lat_ant.values
        df_tasks.loc[m_swap_valido, 'COORD_CORRIGIDA'] = 'SIM - LAT/LON INVERTIDAS'

    st.session_state.qtd_coords_autocorrigidas = int(m_swap_valido.sum())
    m_na = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna()
    m_0 = (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)
    m_inv = (abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])) & ~m_na & ~m_0 & ~m_pos
    df_tasks.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df_tasks.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    df_tasks.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    df_tasks.loc[m_inv, 'MOTIVO_REJEICAO'] = 'Coordenada Invertida'

    mc = m_na | m_0 | m_pos | m_inv
    if mc.sum() > 0: df_rej = pd.concat([df_rej, df_tasks[mc].copy()], ignore_index=True); df_tasks = df_tasks[~mc].copy()

    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']; df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    st.session_state.df_correcao_tatica = df_rej

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    # Normaliza municipios uma unica vez. Evita milhares de chamadas repetidas durante a atribuicao.
    try:
        df_tasks['MUN_LIMPO'] = normalizar_municipios(df_tasks['MUNICIPIO'].astype(str)).astype(str).str.strip().str.upper()
    except Exception:
        df_tasks['MUN_LIMPO'] = df_tasks['MUNICIPIO'].astype(str).str.strip().str.upper()

    # Pré-validação rápida antes de qualquer cálculo pesado.
    capacidade_total_teorica = int(cm * qtd_eq)
    qtd_validas = int(len(df_tasks))
    qtd_rejeitadas = int(len(df_rej))
    excesso_teorico = max(0, qtd_validas - capacidade_total_teorica)
    pv1, pv2, pv3, pv4 = st.columns(4)
    pv1.metric("Obras válidas", qtd_validas)
    pv2.metric("Retidas/correção", qtd_rejeitadas)
    pv3.metric("Equipes ativas", int(qtd_eq))
    pv4.metric("Excesso teórico", excesso_teorico)
    if not dias_sel:
        st.warning("Nenhum dia útil foi selecionado. O motor manterá a regra padrão interna, mas é recomendável escolher os dias para que a capacidade exibida corresponda ao planejamento.")

    # O formulário evita reruns ao ajustar apenas a configuração final de saída.
    # Uploads e filtros permanecem como antes para preservar as opções dinâmicas da tela.
    with st.form("form_inicio_motor_tatico", clear_on_submit=False):
        with st.expander("🛠️ Configuração de Saída", expanded=True):
            tc = [c for c in df_tasks.columns if not c.startswith('_') and c != 'MUN_LIMPO']
            for c_extra in ['BASE_ATRIBUIDA', 'SUPER_PONTO']:
                if c_extra not in tc:
                    tc.append(c_extra)

            cd = [
                'ID SISCO', 'PROTOCOLO', 'CONTA CONTRATO', 'INSTALACAO', 'NOME',
                'ENDERECO', 'LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE',
                'INFORMACOES EXTRAS', 'TIPO NOTA', 'FASE'
            ]
            cp = [c for c in cd if c in tc]
            colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
            colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)
        st.info("⚡ Distribuição, balanceamento e Super Pontos só serão processados depois da confirmação abaixo.")
        iniciar_motor = st.form_submit_button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True)

    if iniciar_motor and not st.session_state.get('inicio_em_processamento_tatica', False):
        st.session_state.inicio_em_processamento_tatica = True
        st.session_state.prep_tasks_tatica = df_tasks.copy()
        st.session_state.prep_bases_tatica = df_bases.copy()
        st.session_state.prep_params_tatica = {
            'ta': ta, 'raio_sp': raio_sp, 'cm': cm, 'qtd_eq': qtd_eq,
            'obras_dia': obras_dia, 'tpc': tpc, 'limite_per': limite_per,
            'dias_sel': list(dias_sel), 'url_osrm': url_osrm, 'usa_osrm': usa_osrm,
            'data_ini': data_ini, 'sentido_rota': sentido_rota,
            'colunas_exibir': list(colunas_exibir)
        }
        st.session_state.vrp_status = "PREPARING"
        tentar_rerun()

if status_exec == "PREPARING":
    st.markdown("## ⚙️ Preparando distribuição das obras")
    if st.button("⏹️ Abortar Preparação", use_container_width=True):
        limpar_roteirizador(); st.stop()

    df_tasks = st.session_state.get('prep_tasks_tatica', pd.DataFrame()).copy()
    df_bases = st.session_state.get('prep_bases_tatica', pd.DataFrame()).copy()
    pp = st.session_state.get('prep_params_tatica', {})

    if df_tasks.empty or df_bases.empty:
        st.error("🚨 Dados de preparação não encontrados. Inicie uma nova roteirização.")
        st.session_state.vrp_status = "IDLE"
        st.stop()

    ta_p = pp.get('ta', 'Por Município Base')
    raio_sp_p = pp.get('raio_sp', 50)
    cm_p = max(1, int(pp.get('cm', 1)))
    qtd_eq_p = max(1, int(pp.get('qtd_eq', df_bases['BASE_NOME'].nunique())))

    tbr = df_bases.to_dict('records')
    base_anchors = {b['BASE_NOME']: (float(b.get('LATITUDE', 0)), float(b.get('LONGITUDE', 0))) for b in tbr}
    fiscal_anchors = dict(base_anchors)
    carga_equipes = {b['BASE_NOME']: 0 for b in tbr}

    # Mapa municipio -> equipes, calculado uma unica vez.
    mun_to_bases = {}
    if "Município" in ta_p:
        for b in tbr:
            mun = str(b.get('MUN_LIMPO_BASE', '')).strip().upper()
            if not mun:
                raw_m = b.get('MUNICIPIO', b.get('RESIDENCIA', ''))
                mun = municipio_normalizado(raw_m)
            mun_to_bases.setdefault(mun, []).append(b)

    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
    registros = df_tasks.to_dict('records')
    assigned_tasks, unassigned_tasks = [], []

    pb_prep = st.progress(0.0)
    msg_prep = st.empty()
    total_reg = max(1, len(registros))

    for idx, r in enumerate(registros):
        la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
        ms = str(r.get('MUN_LIMPO', '')).strip().upper()
        vb = mun_to_bases.get(ms, []) if "Município" in ta_p else tbr

        best_f, best_score = None, float('inf')
        if pd.notna(la) and pd.notna(lo) and vb:
            # Caso comum de municipio rigido com uma unica equipe: atribuicao direta.
            if len(vb) == 1:
                best_f = vb[0]['BASE_NOME']
            else:
                metricas = []
                for b in vb:
                    f_name = b['BASE_NOME']
                    d_base = haversine_scalar(la, lo, base_anchors[f_name][0], base_anchors[f_name][1])
                    d_bolsao = haversine_scalar(la, lo, fiscal_anchors[f_name][0], fiscal_anchors[f_name][1])
                    metricas.append((b, d_base, d_bolsao, carga_equipes.get(f_name, 0)))
                max_db = max([m[1] for m in metricas] + [1.0])
                max_dbol = max([m[2] for m in metricas] + [1.0])
                max_carga = max([m[3] for m in metricas] + [1])
                for b, d_base, d_bolsao, carga in metricas:
                    f_name = b['BASE_NOME']
                    n_base = d_base / max_db
                    n_bolsao = d_bolsao / max_dbol
                    n_carga = carga / max(1, max_carga)
                    n_cap = min(1.0, carga / cm_p)
                    score = (0.40 * n_base) + (0.30 * n_carga) + (0.20 * n_bolsao) + (0.10 * n_cap)
                    if score < best_score:
                        best_score = score
                        best_f = f_name

        if best_f:
            r['BASE_ATRIBUIDA'] = best_f
            assigned_tasks.append(r)
            fiscal_anchors[best_f] = (la, lo)
            carga_equipes[best_f] = carga_equipes.get(best_f, 0) + contar_obras_registro(r)
        else:
            r['MOTIVO_REJEICAO'], r['BASE_ATRIBUIDA'] = "Fora de Área (Sem Fiscal)", "NÃO ALOCADO"
            unassigned_tasks.append(r)

        if idx % 250 == 0 or idx == total_reg - 1:
            frac = 0.65 * ((idx + 1) / total_reg)
            pb_prep.progress(min(frac, 0.65))
            msg_prep.info(f"📍 Distribuindo obras entre as equipes... {idx + 1}/{total_reg}")

    df_ta = pd.DataFrame(assigned_tasks)
    df_u = pd.DataFrame(unassigned_tasks)

    # Super Pontos continuam com a mesma funcao original, mas agora so rodam uma vez, apos o clique.
    dfs_fundidos = []
    if not df_ta.empty:
        bases_unicas = list(df_ta['BASE_ATRIBUIDA'].dropna().unique())
        total_bases = max(1, len(bases_unicas))
        for i_base, base in enumerate(bases_unicas):
            msg_prep.info(f"🏢 Consolidando Super Pontos: {base} ({i_base + 1}/{total_bases})")
            df_base = df_ta[df_ta['BASE_ATRIBUIDA'] == base].copy()
            df_base_f, _ = fundir_super_pontos(df_base, raio_metros=raio_sp_p, agrupar_por_levantador=True)
            dfs_fundidos.append(df_base_f)
            pb_prep.progress(0.65 + 0.30 * ((i_base + 1) / total_bases))
        df_ta = pd.concat(dfs_fundidos, ignore_index=True) if dfs_fundidos else pd.DataFrame()

    if df_ta.empty:
        st.error("Nenhuma obra pôde ser alocada aos Fiscais.")
        st.session_state.vrp_status = "IDLE"
        st.stop()

    st.session_state.df_unallocated_area = df_u.copy()
    st.session_state.df_unallocated_capacity = pd.DataFrame()
    st.session_state.df_unallocated = df_u.copy()
    total_alocadas = sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    sb_html.markdown(render_sidebar_card(cm_p, total_alocadas, qtd_eq_p, cm_p * qtd_eq_p), unsafe_allow_html=True)

    st.session_state.update({
        'bases_records': tbr,
        'colunas_exibir': pp.get('colunas_exibir', [])
    })
    st.session_state.vrp_state = {
        'config': {
            'velocidade_media_kmh': 30.0,
            'obras_por_dia': pp.get('obras_dia', 6),
            'tipo_periodo': pp.get('tpc', 'Semana'),
            'limite_periodos': pp.get('limite_per', 1),
            'dias_selecionados': pp.get('dias_sel', []),
            'url_osrm_base': pp.get('url_osrm', 'http://router.project-osrm.org'),
            'tracado_real': pp.get('usa_osrm', True),
            'data_inicio': pp.get('data_ini', datetime.today().date()),
            'tempo_medio_obra': 45.0 / 60.0,
            'sentido_rota': pp.get('sentido_rota', '📍 Lógica Padrão')
        },
        'b_names': list(dict.fromkeys([b['BASE_NOME'] for b in tbr])),
        'b_idx': 0,
        'unvisited': df_ta.copy(),
        'routed_data': [],
        'current_geoms': [],
        'route_cache': dict(st.session_state.get('osrm_cache_tatica_global', {}))
    }

    # Libera cópias intermediárias grandes antes de iniciar o roteamento.
    for k in ['prep_tasks_tatica', 'prep_bases_tatica']:
        st.session_state.pop(k, None)
    st.session_state.inicio_em_processamento_tatica = False
    gc.collect()

    pb_prep.progress(1.0)
    msg_prep.success("✅ Preparação concluída. Iniciando o motor de roteirização...")
    st.session_state.vrp_status = "RUNNING"
    time.sleep(0.3)
    tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor VRP Tático")
    cp_salvo = st.session_state.get('checkpoint_tatico')
    if cp_salvo:
        st.caption(f"Checkpoint: {cp_salvo.get('indice_proxima', 0)}/{cp_salvo.get('total_equipes', 0)} equipes processadas. Última: {cp_salvo.get('equipe_concluida', '-')}")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run', time.time())
    if 'start_time_run' not in st.session_state: st.session_state.start_time_run = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    bases_df_run = pd.DataFrame(st.session_state.get('bases_records', []))
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            br = bases_df_run[bases_df_run['BASE_NOME'] == bn].iloc[0]
            if pd.isna(br.get('LATITUDE')): st_v['b_idx'] += 1; st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
            bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            
            # Primeiro forma bolsões compatíveis com a cota diária; depois otimiza a ordem dentro de cada bolsão.
            ot = ordenar_bolsoes_diarios(
                oe, bl, bL, cfg['obras_por_dia'], cfg.get('sentido_rota', "Lógica Padrão"), cfg['url_osrm_base']
            ) if oe else []
            if not ot: ot = oe
            
            rf, da, sa, dds = [], 1, 1, 1
            dtb = datetime.combine(cfg['data_inicio'], datetime.min.time()).replace(hour=8, minute=0)
            def gi(da):
                c, d_ok = dtb, [0,1,2,3,4,5,6] if not cfg['dias_selecionados'] else [{"Segunda":0,"Terça":1,"Quarta":2,"Quinta":3,"Sexta":4,"Sábado":5,"Domingo":6}[d] for d in cfg['dias_selecionados']]
                while c.weekday() not in d_ok: c += pd.Timedelta(days=1)
                ct = 1
                while ct < da:
                    c += pd.Timedelta(days=1)
                    if c.weekday() in d_ok: ct += 1
                return {'l': bl, 'L': bL, 't': c, 'd': c, 'oh': 0, 'lu': False}
            es = gi(da)

            for o in ot:
                if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                    registrar_excedente_capacidade(o); continue

                qr = len(o.get('_ORIGINAL_ROWS', [1])) if isinstance(o.get('_ORIGINAL_ROWS'), list) else 1
                vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                if vkr < 0.05 and es['oh'] > 0: vm, em = 0.0, 30.0
                else: vm, em = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60, cfg['tempo_medio_obra']*60
                
                cp = es['t'] + pd.Timedelta(minutes=vm)
                if cp.hour >= 12 and not es['lu']:
                    ls = max(es['t'], es['d'].replace(hour=12)); le = ls + pd.Timedelta(hours=1)
                    rf.append({'o': None, 'il': True, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': es['l'], 'Lt': es['L'], 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': ls, 'hf': le, 'vm': 0.0, 'dk': 0.0})
                    es['t'], es['lu'] = le, True; cp = es['t'] + pd.Timedelta(minutes=vm)
                fp = cp + pd.Timedelta(minutes=em)
                
                if es['oh'] > 0 and (es['oh'] + qr > cfg['obras_por_dia']):
                    dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                    rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
                    da += 1
                    if cfg['tipo_periodo'] == "Semana":
                        dds += 1
                        if dds > len(cfg['dias_selecionados']): sa += 1; dds = 1
                    es = gi(da)
                    if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                        registrar_excedente_capacidade(o); continue
                    vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE']); vk = vkr * 1.3
                    if vkr < 0.05 and es['oh'] > 0: vm, em = 0.0, 30.0
                    else: vm, em = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60, cfg['tempo_medio_obra']*60
                    cp = es['t'] + pd.Timedelta(minutes=vm); fp = cp + pd.Timedelta(minutes=em)
                
                rf.append({'o': o, 'il': False, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': cp, 'hf': fp, 'vm': vm, 'dk': vk})
                es['l'], es['L'], es['t'], es['oh'] = o['LATITUDE'], o['LONGITUDE'], fp, es['oh'] + qr
                
            if es['oh'] > 0 and not ((cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos'])):
                dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
            
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st_v['current_meta'] = []; st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            gm = st_v.get('current_meta', [])
            cache = st_v.setdefault('route_cache', {})
            ei = min(oi + (10 if cfg['tracado_real'] else len(rf)), len(rf))

            for i in range(oi, ei):
                it = rf[i]
                if i % 5 == 0:
                    sgt.info(f"🛣️ Etapa de arruamento — **{bn}** | trecho {i + 1}/{len(rf)} | equipe {b_i + 1}/{len(b_n)}")
                render_t(b_i, i, len(rf))

                # Pausas e outros deslocamentos nulos não precisam consultar o OSRM.
                if abs(float(it['la']) - float(it['lt'])) < 1e-9 and abs(float(it['La']) - float(it['Lt'])) < 1e-9:
                    gd.append(([], 0.0))
                    gm.append({'status': 'SEM_DESLOCAMENTO', 'km_real': 0.0, 'km_estimado': 0.0, 'tempo_real_min': 0.0})
                    continue

                if not cfg['tracado_real']:
                    geom = [[it['La'], it['la']], [it['Lt'], it['lt']]]
                    dur = (it['dk'] / max(cfg['velocidade_media_kmh'], 1.0)) * 3600
                    gd.append((geom, dur))
                    gm.append({'status': 'LINHA_RETA_ESTIMADA', 'km_real': np.nan, 'km_estimado': float(it['dk']), 'tempo_real_min': dur / 60.0})
                    continue

                chave = chave_rota_cache(it['la'], it['La'], it['lt'], it['Lt'])
                cached = cache.get(chave)
                if cached is not None:
                    rota, meta = cached
                    gd.append(rota)
                    gm.append(dict(meta))
                    continue

                sucesso_rota = False
                ultimo_erro = None
                for tentativa in range(5):
                    try:
                        time.sleep(0.8 if tentativa == 0 else 1.2)
                        ini_osrm = time.time()
                        rota = chamar_osrm_compativel(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'], timeout_s=12.0)
                        tempo_chamada = time.time() - ini_osrm
                        if rota and len(rota) >= 2 and isinstance(rota[0], list) and len(rota[0]) >= 2:
                            geom, dur = rota[0], rota[1]
                            km_real = distancia_geometria_km(geom)
                            status_ok = 'OSRM_LENTO' if tempo_chamada > 12.0 else 'OSRM'
                            meta = {'status': status_ok, 'km_real': km_real if km_real > 0 else np.nan, 'km_estimado': float(it['dk']), 'tempo_real_min': float(dur) / 60.0 if pd.notna(dur) else np.nan}
                            gd.append((geom, dur))
                            gm.append(meta)
                            cache[chave] = ((geom, dur), dict(meta))
                            sucesso_rota = True
                            break
                    except Exception as exc:
                        ultimo_erro = exc
                        time.sleep(1.5 + tentativa * 0.5)

                if not sucesso_rota:
                    # Com arruamento real ativado, não inventa uma reta. Mantém o trecho vazio e sinalizado.
                    rota_vazia = ([], 0.0)
                    meta = {'status': 'SEM_ROTA_OSRM', 'km_real': np.nan, 'km_estimado': float(it['dk']), 'tempo_real_min': np.nan}
                    gd.append(rota_vazia)
                    gm.append(meta)
                    cache[chave] = (rota_vazia, dict(meta))

            st_v['c_idx'], st_v['current_geoms'], st_v['current_meta'], st_v['route_cache'] = ei, gd, gm, cache
            cache_global = st.session_state.get('osrm_cache_tatica_global', {})
            for ck, cv in cache.items():
                try:
                    if cv and cv[0] and isinstance(cv[0][0], list) and len(cv[0][0]) >= 2:
                        cache_global[ck] = cv
                except Exception:
                    pass
            st.session_state.osrm_cache_tatica_global = cache_global
            if ei < len(rf): st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
            
            base_row_run = bases_df_run[bases_df_run['BASE_NOME'] == bn].iloc[0]
            bl, bL = float(base_row_run['LATITUDE']), float(base_row_run['LONGITUDE'])
            rdf, og, dp = [], 1, ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            gm = st_v.get('current_meta', [{} for _ in gd])
            for idx_seg, (it, (g, ds)) in enumerate(zip(rf, gd)):
                meta = gm[idx_seg] if idx_seg < len(gm) else {}
                km_real_meta = meta.get('km_real', np.nan)
                km_estimado = float(meta.get('km_estimado', it.get('dk', 0.0)) or 0.0)
                km_rota = float(km_real_meta) if pd.notna(km_real_meta) else km_estimado
                status_rota = str(meta.get('status', 'ESTIMADA'))
                tempo_real = meta.get('tempo_real_min', np.nan)
                pv = it['s'] if cfg['tipo_periodo']=="Semana" else it['d']
                dn = dp[datetime.strptime(it['dm'], '%d/%m/%Y').weekday()] if cfg['tipo_periodo']=="Semana" else f"Dia {it['d']}"
                comum = {'DISTANCIA_RODOVIARIA_KM': round(float(km_real_meta), 2) if pd.notna(km_real_meta) else np.nan, 'DISTANCIA_ESTIMADA_KM': round(km_estimado, 2), 'STATUS_ROTA': status_rota, 'TEMPO_ROTA_REAL_MIN': round(float(tempo_real), 1) if pd.notna(tempo_real) else np.nan}
                if it['il']:
                    linha = {'PROTOCOLO': 'PAUSA_ALMOCO', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': 0.0, 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']}
                    linha.update(comum); rdf.append(linha)
                elif it['ir']:
                    linha = {'PROTOCOLO': 'RETORNO_BASE', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': round(km_rota, 2), 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']}
                    linha.update(comum); rdf.append(linha)
                else:
                    ob = it['o']; ob['ORDEM'], ob['NOME_DIA'], ob['DIA_MES'], ob['SEMANA'], ob['DIA'], ob['PERIODO'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, dn, it['dm'], it['s'], it['d'], pv, round(km_rota, 2)
                    # Preserva o traçado real também no primeiro deslocamento BASE -> primeira obra.
                    ob['ROTA_GEOMETRIA'] = g
                    ob.update(comum)
                    ob['HORA_INICIO'], ob['HORA_FIM'], ob['_HORA_INICIO_DT'], ob['_HORA_FIM_DT'] = it['hi'].strftime('%H:%M'), it['hf'].strftime('%H:%M'), it['hi'], it['hf']
                    rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']; st_v.pop('current_meta', None)
            st_v['b_idx'] += 1
            st.session_state.checkpoint_tatico = {'equipe_concluida': bn, 'indice_proxima': st_v['b_idx'], 'total_equipes': len(b_n), 'rotas_acumuladas': len(st_v.get('routed_data', []))}
            st.session_state.vrp_state = st_v
            gc.collect()
            tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        df_final = pd.DataFrame(st_v.get('routed_data', []))
        st.session_state.df_routed = df_final
        cache_global = st.session_state.get('osrm_cache_tatica_global', {})
        for ck, cv in st_v.get('route_cache', {}).items():
            try:
                if cv and cv[0] and isinstance(cv[0][0], list) and len(cv[0][0]) >= 2:
                    cache_global[ck] = cv
            except Exception:
                pass
        st.session_state.osrm_cache_tatica_global = cache_global
        st.session_state.metricas_resultado_tatica = calcular_metricas_resultado_tatica(df_final)
        st.session_state.d_fmt_resultado_tatica = datetime.now().strftime("%d.%m.%Y")
        # Downloads passam a ser preparados apenas quando solicitados.
        for k in ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx']:
            st.session_state.pop(k, None)
        # Mantém somente a configuração necessária para resultado/exportação e libera estruturas pesadas.
        st.session_state.vrp_state = {'config': dict(cfg)}
        st.session_state.roteamento_concluido = True
        st.session_state.vrp_status = "IDLE"
        gc.collect()
        time.sleep(0.3)
        tentar_rerun()

if status_exec == "PACKAGING":
    # Compatibilidade com sessões antigas: a etapa de empacotamento automático foi removida.
    if not st.session_state.get('df_routed', pd.DataFrame()).empty:
        st.session_state.metricas_resultado_tatica = calcular_metricas_resultado_tatica(st.session_state.df_routed)
        st.session_state.d_fmt_resultado_tatica = st.session_state.get('d_fmt_resultado_tatica', datetime.now().strftime("%d.%m.%Y"))
        st.session_state.roteamento_concluido = True
    st.session_state.vrp_status = "IDLE"
    tentar_rerun()
