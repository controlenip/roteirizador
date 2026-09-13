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
import uuid
import unicodedata
import traceback
from datetime import datetime, time as dt_time
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

from modules.data_processing import ler_planilha_cached, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas
from modules.export_saneamento import (
    injetar_logo,
    identificar_icone_folium,
    gerar_excel_saneamento,
    limpar_colunas_saneamento,
    gerar_kml_saneamento,
    gerar_gpx_simples,
    gerar_excel_resumo_saneamento,
    gerar_excel_generico,
)

st.set_page_config(page_title="Saneamento", page_icon="🧹", layout="wide")
injetar_logo()

VERSAO_REGRAS_SANEAMENTO = "2026.09"
DIAS_NOMES = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
DIAS_MAP = {nome: i for i, nome in enumerate(DIAS_NOMES)}

st.markdown('''
    <style>
        [data-testid="stSidebarHeader"] img, [data-testid="stLogo"] img {
            max-height: 1.5rem !important;
            height: auto !important;
            width: auto !important;
            padding: 0 !important;
            margin: 0 !important;
        }
    </style>
''', unsafe_allow_html=True)


# ==============================================================
# FUNÇÕES DE APOIO / AUDITORIA
# ==============================================================
def remover_acentos_str(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    return "".join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))


def normalizar_texto(valor):
    if pd.isna(valor):
        return ''
    s = remover_acentos_str(str(valor)).upper().strip()
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def limpar_protocolo_serie(serie):
    s = serie.astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
    invalidos = {'', 'NAN', 'NONE', 'NULL', '-'}
    return s.mask(s.str.upper().isin(invalidos), '')


def ler_csv_resiliente(file_bytes):
    ultimo = None
    for enc in ['utf-8-sig', 'utf-8', 'latin-1']:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding=enc)
        except Exception as exc:
            ultimo = exc
    for enc in ['utf-8-sig', 'latin-1']:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=';', encoding=enc)
        except Exception as exc:
            ultimo = exc
    raise ultimo if ultimo else ValueError('Não foi possível ler o CSV.')


def criar_id_execucao():
    return f"SAN-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"


def formatar_valor_coluna(c, v):
    if isinstance(v, (list, tuple, dict, set)):
        return html.escape(str(v))
    if pd.isna(v) or v in ['', '-']:
        return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper() or c.upper().endswith('_KM'):
            return f"{vf:.2f} KM"
        if 'TEMPO' in c.upper() and c.upper().endswith('_MIN'):
            return f"{vf:.1f} min"
        return formata_campo_html(v)
    except Exception:
        if isinstance(v, (datetime, pd.Timestamp)):
            return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))


def peso_tarefa(registro):
    orig = registro.get('_ORIGINAL_ROWS') if hasattr(registro, 'get') else None
    return len(orig) if isinstance(orig, list) and orig else 1


def distancia_geometria_km(geom):
    if not isinstance(geom, list) or len(geom) < 2:
        return np.nan
    total = 0.0
    anterior = None
    for pt in geom:
        if not isinstance(pt, (list, tuple)) or len(pt) < 2:
            continue
        atual = (float(pt[1]), float(pt[0]))  # lat, lon
        if anterior is not None:
            total += float(haversine_scalar(anterior[0], anterior[1], atual[0], atual[1]))
        anterior = atual
    return total if anterior is not None else np.nan


def estimar_deslocamento(lat1, lon1, lat2, lon2, velocidade_media_kmh):
    d_reta = float(haversine_scalar(float(lat1), float(lon1), float(lat2), float(lon2)))
    d_est = d_reta * 1.3
    vel = float(velocidade_media_kmh) * (1.5 if d_est > 20 else 1.0)
    tempo_min = (d_est / max(1.0, vel)) * 60.0
    return d_est, tempo_min


def aplicar_intervalo_almoco(inicio, duracao_min, cfg):
    if not cfg.get('usar_intervalo_almoco', True):
        return inicio
    data = inicio.date()
    a_ini = datetime.combine(data, cfg['almoco_inicio'])
    a_fim = datetime.combine(data, cfg['almoco_fim'])
    fim_prev = inicio + pd.Timedelta(minutes=float(duracao_min))
    if a_ini <= inicio < a_fim:
        return a_fim
    if inicio < a_ini < fim_prev:
        return a_fim
    return inicio


def data_trabalho_por_indice(data_inicio, indice, dias_selecionados):
    dias_ok = [DIAS_MAP[d] for d in dias_selecionados] if dias_selecionados else list(range(7))
    atual = pd.Timestamp(data_inicio).normalize()
    while atual.weekday() not in dias_ok:
        atual += pd.Timedelta(days=1)
    contador = 1
    while contador < indice:
        atual += pd.Timedelta(days=1)
        if atual.weekday() in dias_ok:
            contador += 1
    return atual.to_pydatetime()


def periodo_do_dia(indice_dia, tipo_periodo, dias_selecionados):
    qtd_dias_semana = max(1, len(dias_selecionados) if dias_selecionados else 7)
    semana = ((indice_dia - 1) // qtd_dias_semana) + 1
    dia_na_semana = ((indice_dia - 1) % qtd_dias_semana) + 1
    periodo = semana if tipo_periodo == 'Semana' else indice_dia
    return semana, dia_na_semana, periodo


def ordenar_bloco_rota(tarefas, lat_inicio, lon_inicio, sentido_rota, url_osrm_base):
    """Prioridade é preservada: otimiza primeiro alta prioridade, depois as demais."""
    tarefas = [dict(x) for x in tarefas]
    if not tarefas:
        return []

    altas = [x for x in tarefas if str(x.get('PRIORIDADE', '')).strip().upper() == 'SIM']
    normais = [x for x in tarefas if str(x.get('PRIORIDADE', '')).strip().upper() != 'SIM']
    saida = []
    atual_lat, atual_lon = float(lat_inicio), float(lon_inicio)

    def ordenar_grupo(grupo, la, lo):
        grupo = [dict(x) for x in grupo]
        if not grupo:
            return []
        if 'Varredura Reversa' in str(sentido_rota):
            primeiro = max(
                range(len(grupo)),
                key=lambda i: haversine_scalar(la, lo, float(grupo[i]['LATITUDE']), float(grupo[i]['LONGITUDE']))
            )
            ordem = [grupo.pop(primeiro)]
            cl, co = float(ordem[0]['LATITUDE']), float(ordem[0]['LONGITUDE'])
            while grupo:
                idx = min(
                    range(len(grupo)),
                    key=lambda i: haversine_scalar(cl, co, float(grupo[i]['LATITUDE']), float(grupo[i]['LONGITUDE']))
                )
                nx = grupo.pop(idx)
                ordem.append(nx)
                cl, co = float(nx['LATITUDE']), float(nx['LONGITUDE'])
            return ordem
        try:
            ordem = resolver_tsp_ortools(grupo, la, lo, url_osrm_base)
            return ordem if ordem else grupo
        except Exception:
            # Fallback determinístico por vizinho mais próximo; não altera o status OSRM dos segmentos.
            restantes = grupo[:]
            ordem = []
            cl, co = la, lo
            while restantes:
                idx = min(
                    range(len(restantes)),
                    key=lambda i: haversine_scalar(cl, co, float(restantes[i]['LATITUDE']), float(restantes[i]['LONGITUDE']))
                )
                nx = restantes.pop(idx)
                ordem.append(nx)
                cl, co = float(nx['LATITUDE']), float(nx['LONGITUDE'])
            return ordem

    for grupo in [altas, normais]:
        ord_g = ordenar_grupo(grupo, atual_lat, atual_lon)
        if ord_g:
            saida.extend(ord_g)
            atual_lat, atual_lon = float(ord_g[-1]['LATITUDE']), float(ord_g[-1]['LONGITUDE'])
    return saida


def aplicar_trava_global(df, limite):
    """Aplica a trava sobre tarefas reais (antes de Super Pontos) e devolve o excedente auditável."""
    if not limite or int(limite) <= 0 or len(df) <= int(limite):
        return df.copy(), pd.DataFrame(columns=df.columns)
    ordenado = df.copy()
    if '_ORDEM_ENTRADA' not in ordenado.columns:
        ordenado['_ORDEM_ENTRADA'] = np.arange(len(ordenado))
    prioridade_num = ordenado.get('PRIORIDADE', pd.Series('Não', index=ordenado.index)).astype(str).str.upper().eq('SIM').astype(int)
    ordenado = ordenado.assign(_PRIORIDADE_NUM=prioridade_num).sort_values(['_PRIORIDADE_NUM', '_ORDEM_ENTRADA'], ascending=[False, True])
    dentro = ordenado.head(int(limite)).drop(columns=['_PRIORIDADE_NUM'], errors='ignore').copy()
    fora = ordenado.iloc[int(limite):].drop(columns=['_PRIORIDADE_NUM'], errors='ignore').copy()
    if not fora.empty:
        fora['MOTIVO_NAO_ALOCACAO'] = 'FORA_DA_TRAVA_GLOBAL'
        fora['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
    return dentro, fora


def validar_coordenadas_obras(df_tasks):
    df = df_tasks.copy()
    df['LATITUDE_ORIGINAL'] = df['LATITUDE']
    df['LONGITUDE_ORIGINAL'] = df['LONGITUDE']
    df['LAT_NUM'] = pd.to_numeric(df['LATITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
    df['LON_NUM'] = pd.to_numeric(df['LONGITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
    m_na = df['LAT_NUM'].isna() | df['LON_NUM'].isna()
    m_0 = (df['LAT_NUM'] == 0.0) | (df['LON_NUM'] == 0.0)
    m_pos = (df['LAT_NUM'] > 0) | (df['LON_NUM'] > 0)
    m_inv = df['LAT_NUM'].abs() > df['LON_NUM'].abs()
    df['MOTIVO_REJEICAO'] = ''
    df.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    df.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    df.loc[m_inv & ~m_na & ~m_0 & ~m_pos, 'MOTIVO_REJEICAO'] = 'Possível Latitude/Longitude Invertida'
    rejeitar = m_na | m_0 | m_pos | m_inv
    rejeitadas = df[rejeitar].copy()
    validas = df[~rejeitar].copy()
    validas['LATITUDE'] = validas['LAT_NUM']
    validas['LONGITUDE'] = validas['LON_NUM']
    return validas, rejeitadas


def construir_plano_rota_equipe(tarefas, base_lat, base_lon, cfg):
    """Divide primeiro por dia e só então otimiza cada rota diária."""
    pending = [dict(x) for x in tarefas]
    # Alta prioridade vem primeiro na seleção do dia; ordem espacial é otimizada depois.
    pending.sort(key=lambda r: (0 if str(r.get('PRIORIDADE', '')).upper() == 'SIM' else 1, int(r.get('_ORDEM_ENTRADA', 10**9))))
    rotas = []
    nao_alocadas = []
    dia_idx = 1
    qtd_dias_semana = len(cfg.get('dias_selecionados') or []) or 7
    max_dias = None
    if not cfg.get('modo_continuo', False):
        max_dias = int(cfg['limite_periodos']) if cfg['tipo_periodo'] == 'Dia' else int(cfg['limite_periodos']) * qtd_dias_semana

    while pending:
        if max_dias is not None and dia_idx > max_dias:
            for o in pending:
                o['MOTIVO_NAO_ALOCACAO'] = 'LIMITE_DE_PERIODOS_ATINGIDO'
                o['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
                nao_alocadas.append(o)
            break

        data_dia = data_trabalho_por_indice(cfg['data_inicio'], dia_idx, cfg.get('dias_selecionados', []))
        inicio_jornada = datetime.combine(data_dia.date(), cfg['hora_inicio'])
        fim_jornada = datetime.combine(data_dia.date(), cfg['hora_fim'])
        total_min_jornada = max(1.0, (fim_jornada - inicio_jornada).total_seconds() / 60.0)
        if cfg.get('usar_intervalo_almoco', True):
            total_min_jornada -= max(0.0, (datetime.combine(data_dia.date(), cfg['almoco_fim']) - datetime.combine(data_dia.date(), cfg['almoco_inicio'])).total_seconds() / 60.0)

        # Seleção inicial do bloco respeita cota e tempo mínimo de atendimento.
        bloco, resto = [], []
        carga = 0
        servico = 0.0
        for o in pending:
            qr = peso_tarefa(o)
            serv = float(cfg['tempo_medio_obra_min']) * (qr if cfg.get('tempo_super_ponto_por_obra', True) else 1)
            if qr > int(cfg['obras_por_dia']):
                o['MOTIVO_NAO_ALOCACAO'] = 'SUPER_PONTO_EXCEDE_COTA_DIARIA'
                o['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
                nao_alocadas.append(o)
                continue
            if serv > total_min_jornada:
                o['MOTIVO_NAO_ALOCACAO'] = 'ATENDIMENTO_EXCEDE_JORNADA_DIARIA'
                o['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
                nao_alocadas.append(o)
                continue
            caberia = (carga + qr <= int(cfg['obras_por_dia'])) and (servico + serv <= total_min_jornada)
            if caberia:
                bloco.append(o)
                carga += qr
                servico += serv
            else:
                resto.append(o)
        pending = resto

        ordem = ordenar_bloco_rota(bloco, base_lat, base_lon, cfg['sentido_rota'], cfg['url_osrm_base'])
        semana, dds, periodo = periodo_do_dia(dia_idx, cfg['tipo_periodo'], cfg.get('dias_selecionados', []))
        atual_lat, atual_lon = float(base_lat), float(base_lon)
        relogio = inicio_jornada
        usados = 0
        aceitos = 0
        deferidos = []

        for pos, o in enumerate(ordem):
            qr = peso_tarefa(o)
            d_est, t_est = estimar_deslocamento(atual_lat, atual_lon, o['LATITUDE'], o['LONGITUDE'], cfg['velocidade_media_kmh'])
            serv_min = float(cfg['tempo_medio_obra_min']) * (qr if cfg.get('tempo_super_ponto_por_obra', True) else 1)
            chegada = relogio + pd.Timedelta(minutes=t_est)
            inicio_serv = aplicar_intervalo_almoco(chegada, serv_min, cfg)
            fim_serv = inicio_serv + pd.Timedelta(minutes=serv_min)
            d_ret, t_ret = estimar_deslocamento(o['LATITUDE'], o['LONGITUDE'], base_lat, base_lon, cfg['velocidade_media_kmh'])

            excede_cota = aceitos > 0 and (usados + qr > int(cfg['obras_por_dia']))
            excede_jornada = aceitos > 0 and (fim_serv + pd.Timedelta(minutes=t_ret) > fim_jornada)
            if excede_cota or excede_jornada:
                deferidos.extend(ordem[pos:])
                break

            if aceitos == 0 and (fim_serv + pd.Timedelta(minutes=t_ret) > fim_jornada):
                o['MOTIVO_NAO_ALOCACAO'] = 'TAREFA_NAO_CABE_NA_JORNADA'
                o['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
                nao_alocadas.append(o)
                continue

            rotas.append({
                'o': o, 'ir': False,
                'la': atual_lat, 'La': atual_lon, 'lt': float(o['LATITUDE']), 'Lt': float(o['LONGITUDE']),
                's': semana, 'd': dia_idx, 'ds': dds, 'periodo': periodo,
                'dn': DIAS_NOMES[data_dia.weekday()], 'dm': data_dia.strftime('%d/%m/%Y'),
                'hi_est': inicio_serv, 'hf_est': fim_serv, 'servico_min': serv_min,
                'tempo_estimado_min': t_est, 'distancia_estimada_km': d_est,
            })
            atual_lat, atual_lon = float(o['LATITUDE']), float(o['LONGITUDE'])
            relogio = fim_serv
            usados += qr
            aceitos += 1

        if deferidos:
            # Reotimiza os itens adiados no próximo dia, em vez de herdar a ordem do dia anterior.
            pending = deferidos + pending

        if aceitos > 0:
            d_ret, t_ret = estimar_deslocamento(atual_lat, atual_lon, base_lat, base_lon, cfg['velocidade_media_kmh'])
            rotas.append({
                'o': None, 'ir': True,
                'la': atual_lat, 'La': atual_lon, 'lt': float(base_lat), 'Lt': float(base_lon),
                's': semana, 'd': dia_idx, 'ds': dds, 'periodo': periodo,
                'dn': DIAS_NOMES[data_dia.weekday()], 'dm': data_dia.strftime('%d/%m/%Y'),
                'hi_est': relogio, 'hf_est': relogio + pd.Timedelta(minutes=t_ret), 'servico_min': 0.0,
                'tempo_estimado_min': t_ret, 'distancia_estimada_km': d_ret,
            })

        dia_idx += 1

    return rotas, nao_alocadas


def montar_manifesto(config, df_routed=None):
    linhas = [
        'MANIFESTO - ROTEIRIZADOR SANEAMENTO',
        f"ID execução: {config.get('id_execucao', '-')}",
        f"Versão das regras: {VERSAO_REGRAS_SANEAMENTO}",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Modo: {'CONTÍNUO' if config.get('modo_continuo') else 'PADRÃO'}",
        f"Atribuição: {config.get('regra_atribuicao', '-')}",
        f"Cota diária: {config.get('obras_por_dia', '-')}",
        f"Visão: {config.get('tipo_periodo', '-')}",
        f"Limite de períodos: {config.get('limite_periodos', '-')}",
        f"Dias: {', '.join(config.get('dias_selecionados', []))}",
        f"Início da jornada: {config.get('hora_inicio_str', '-')}",
        f"Fim da jornada: {config.get('hora_fim_str', '-')}",
        f"Tempo médio/obra: {config.get('tempo_medio_obra_min', '-')} min",
        f"Raio Super Ponto: {config.get('raio_super_ponto_m', '-')} m",
        f"Distância máx. atribuição por proximidade: {config.get('distancia_max_atribuicao_km', '-')} km",
        f"OSRM: {'SIM' if config.get('tracado_real') else 'NÃO'}",
        f"Endpoint OSRM: {config.get('url_osrm_base', '-')}",
        f"Remoção de duplicidades: {'SIM' if config.get('deduplicar_notas') else 'NÃO'}",
        f"Trava global: {config.get('trava_global', 0)}",
        f"Arquivos de demanda: {', '.join(config.get('arquivos_demanda', []))}",
        f"Arquivo equipes: {config.get('arquivo_equipes', '-')}",
        f"Entradas após filtros: {config.get('qtd_entrada_filtrada', '-')}",
        f"Duplicidades detectadas: {config.get('qtd_duplicidades', '-')}",
        f"Duplicidades removidas: {config.get('qtd_duplicidades_removidas', '-')}",
        f"Coordenadas de obras rejeitadas: {config.get('qtd_correcao_obras', '-')}",
        f"Bases rejeitadas/corrigidas: {config.get('qtd_correcao_bases', '-')}",
        f"Não alocadas: {config.get('qtd_nao_alocadas', '-')}",
    ]
    if df_routed is not None and not df_routed.empty:
        obras = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
        linhas.append(f"Obras roteirizadas: {sum(peso_tarefa(r) for _, r in obras.iterrows())}")
        linhas.append(f"Distância estimada total: {pd.to_numeric(df_routed.get('DISTANCIA_ESTIMADA_KM', 0), errors='coerce').fillna(0).sum():.2f} km")
        linhas.append(f"Distância rodoviária conhecida: {pd.to_numeric(df_routed.get('DISTANCIA_RODOVIARIA_KM', 0), errors='coerce').fillna(0).sum():.2f} km")
    return '\n'.join(linhas) + '\n'


def render_sidebar_card(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade, is_continuo):
    txt_modo = "Contínuo" if is_continuo else "Padrão (Por Cota)"
    cor_modo = "#FF9800" if is_continuo else "#55B929"
    return f"""
    <div style="background-color:#f8f9fa;padding:15px;border-radius:8px;border:1px solid #dee2e6;margin-bottom:20px;">
        <h4 style="margin-top:0;color:#0D256C;font-size:16px;border-bottom:2px solid #55B929;padding-bottom:5px;">📊 Resumo da Operação</h4>
        <p style="margin-bottom:5px;font-size:14px;"><b>Equipes Ativas:</b> <span style="color:#0D256C;font-weight:bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom:5px;font-size:14px;"><b>Modo do Motor:</b> <span style="color:{cor_modo};font-weight:bold;">{txt_modo}</span></p>
        <p style="margin-bottom:5px;font-size:14px;"><b>Cota p/ Equipe:</b> <span style="color:#d9534f;font-weight:bold;">{limite_por_equipe}</span> tarefas</p>
        <hr style="margin:10px 0;border:0;border-top:1px solid #ddd;">
        <p style="margin-bottom:0;font-size:15px;text-align:center;"><b>Tarefas Validadas:</b><br><span style="font-size:24px;color:#0D256C;font-weight:900;">{total_obras_prontas}</span></p>
    </div>
    """


def render_metric_card(title, value, icon, border_color, bg_color):
    return f"""
    <div style="background-color:#ffffff;border-left:5px solid {border_color};padding:15px;border-radius:5px;box-shadow:0 2px 4px rgba(0,0,0,0.1);display:flex;align-items:center;margin-bottom:10px;">
        <div style="background-color:{bg_color};width:40px;height:40px;border-radius:50%;display:flex;justify-content:center;align-items:center;font-size:20px;margin-right:15px;">{icon}</div>
        <div><p style="margin:0;font-size:11px;color:#666;text-transform:uppercase;font-weight:bold;">{title}</p><p style="margin:0;font-size:22px;color:#333;font-weight:bold;">{value}</p></div>
    </div>
    """


def tentar_rerun():
    if hasattr(st, 'rerun'):
        st.rerun()
    else:
        st.experimental_rerun()


def limpar_roteirizador():
    st.session_state.update({
        'roteamento_concluido_san': False,
        'vrp_status_san': 'IDLE',
        'vrp_state_san': {},
        'df_routed_san': pd.DataFrame(),
        'bases_records_san': [],
        'colunas_exibir_san': [],
        'colunas_originais_san': [],
    })
    for k in [
        'bytes_zip_xl_san', 'bytes_zip_kml_san', 'bytes_zip_gpx_san', 'start_time_run_san',
        'df_unallocated_san', 'df_correcao_san', 'df_bases_correcao_san', 'df_duplicadas_san',
        'df_duplicadas_removidas_san', 'config_execucao_san', 'mostrar_mapa_san'
    ]:
        st.session_state.pop(k, None)
    try:
        ler_planilha_cached.clear()
    except Exception:
        pass
    tentar_rerun()


def _expandir_super_pontos_para_excel(df):
    ld = []
    for _, r in df.iterrows():
        if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
            continue
        is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
        sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else 'NÃO'
        if is_sp:
            for orig in r['_ORIGINAL_ROWS']:
                nr = r.copy()
                for k, v in orig.items():
                    if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'ROTA_GEOMETRIA', 'PERIODO']:
                        nr[k] = v
                nr['SUPER_PONTO'] = sp_text
                ld.append(nr)
        else:
            nr = r.copy()
            nr['SUPER_PONTO'] = sp_text
            ld.append(nr)
    return pd.DataFrame(ld)


def gerar_zip_excel_atual():
    df_routed = st.session_state.df_routed_san.copy()
    cfg = st.session_state.get('config_execucao_san', {})
    id_exec = cfg.get('id_execucao', criar_id_execucao())
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zx:
        res = []
        for b in df_routed['BASE_ATRIBUIDA'].dropna().unique():
            db = df_routed[(df_routed['BASE_ATRIBUIDA'] == b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
            if db.empty:
                continue
            qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
            res.append({
                'Equipe': b,
                'Obras Roteirizadas': sum(peso_tarefa(r) for _, r in db.iterrows()),
                'Super Pontos': qs,
                'KM Estimado': round(pd.to_numeric(df_routed.loc[df_routed['BASE_ATRIBUIDA'] == b, 'DISTANCIA_ESTIMADA_KM'], errors='coerce').fillna(0).sum(), 2),
                'KM Rodoviário Conhecido': round(pd.to_numeric(df_routed.loc[df_routed['BASE_ATRIBUIDA'] == b, 'DISTANCIA_RODOVIARIA_KM'], errors='coerce').fillna(0).sum(), 2),
                'Segmentos sem OSRM': int((df_routed.loc[df_routed['BASE_ATRIBUIDA'] == b, 'STATUS_ROTA'].astype(str) == 'SEM_ROTA_OSRM').sum()),
            })
        zx.writestr(f'Resumo_Operacional_{id_exec}.xlsx', gerar_excel_resumo_saneamento(pd.DataFrame(res)))

        extras = [
            ('Obras_Correcao', st.session_state.get('df_correcao_san', pd.DataFrame())),
            ('Bases_Correcao', st.session_state.get('df_bases_correcao_san', pd.DataFrame())),
            ('Obras_Nao_Alocadas', st.session_state.get('df_unallocated_san', pd.DataFrame())),
            ('Auditoria_Duplicadas', st.session_state.get('df_duplicadas_san', pd.DataFrame())),
            ('Duplicadas_Removidas', st.session_state.get('df_duplicadas_removidas_san', pd.DataFrame())),
        ]
        for nome, dfx in extras:
            if dfx is not None and not dfx.empty:
                zx.writestr(f'{nome}_{id_exec}.xlsx', gerar_excel_generico(dfx, nome[:31]))

        col_sel = st.session_state.get('colunas_exibir_san', [])
        for b_name in [x for x in df_routed['BASE_ATRIBUIDA'].dropna().unique() if x != 'NÃO ALOCADO']:
            df_ind = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
            dx = _expandir_super_pontos_para_excel(df_ind)
            if dx.empty:
                continue
            dx = limpar_colunas_saneamento(dx, st.session_state.get('colunas_originais_san', []), col_sel)
            ns = re.sub(r'[^A-Za-z0-9_]+', '_', remover_acentos_str(str(b_name))).strip('_').upper()
            zx.writestr(f'ROTA_{ns}_{id_exec}.xlsx', gerar_excel_saneamento(dx, st.session_state.get('colunas_originais_san', [])))

        zx.writestr(f'Manifesto_{id_exec}.txt', montar_manifesto(cfg, df_routed).encode('utf-8'))
    return out.getvalue()


def gerar_zip_kml_atual():
    df_routed = st.session_state.df_routed_san.copy()
    cfg = st.session_state.get('config_execucao_san', {})
    id_exec = cfg.get('id_execucao', criar_id_execucao())
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zk:
        equipes = [x for x in df_routed['BASE_ATRIBUIDA'].dropna().unique() if x != 'NÃO ALOCADO']
        for b_name in equipes:
            df_ind = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
            ns = re.sub(r'[^A-Za-z0-9_]+', '_', remover_acentos_str(str(b_name))).strip('_').upper()
            kl = gerar_kml_saneamento(df_ind, f'ROTA_{ns}', st.session_state.get('colunas_exibir_san', []), [b_name], cfg.get('tipo_periodo', 'Semana'), formatar_valor_coluna)
            zk.writestr(f'ROTA_{ns}_{id_exec}.kml', kl.encode('utf-8'))
        kt = gerar_kml_saneamento(df_routed, 'ROTA_TOTAL', st.session_state.get('colunas_exibir_san', []), equipes, cfg.get('tipo_periodo', 'Semana'), formatar_valor_coluna)
        zk.writestr(f'ROTA_TOTAL_{id_exec}.kml', kt.encode('utf-8'))
        zk.writestr(f'Manifesto_{id_exec}.txt', montar_manifesto(cfg, df_routed).encode('utf-8'))
    return out.getvalue()


def gerar_zip_gpx_atual():
    df_routed = st.session_state.df_routed_san.copy()
    cfg = st.session_state.get('config_execucao_san', {})
    id_exec = cfg.get('id_execucao', criar_id_execucao())
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zg:
        equipes = [x for x in df_routed['BASE_ATRIBUIDA'].dropna().unique() if x != 'NÃO ALOCADO']
        for b_name in equipes:
            df_ind = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
            ns = re.sub(r'[^A-Za-z0-9_]+', '_', remover_acentos_str(str(b_name))).strip('_').upper()
            zg.writestr(f'GPS_{ns}_{id_exec}.gpx', gerar_gpx_simples(df_ind, f'ROTA_{ns}').encode('utf-8'))
        zg.writestr(f'GPS_TOTAL_{id_exec}.gpx', gerar_gpx_simples(df_routed, 'ROTA TOTAL').encode('utf-8'))
        zg.writestr(f'Manifesto_{id_exec}.txt', montar_manifesto(cfg, df_routed).encode('utf-8'))
    return out.getvalue()


# ==============================================================
# ESTADO / SIDEBAR
# ==============================================================
if 'roteamento_concluido_san' not in st.session_state:
    st.session_state.roteamento_concluido_san = False
if 'vrp_status_san' not in st.session_state:
    st.session_state.vrp_status_san = 'IDLE'
if 'df_routed_san' not in st.session_state:
    st.session_state.df_routed_san = pd.DataFrame()

status_exec = st.session_state.vrp_status_san
is_done = st.session_state.roteamento_concluido_san
is_locked = status_exec != 'IDLE' or is_done

st.markdown("<h1 class='brand-title'>🧹 Roteirizador Saneamento</h1>", unsafe_allow_html=True)
st.info("💡 Planeja Saneamento com auditoria de entrada, divisão diária, prioridade, jornada e rotas por equipe.")

with st.sidebar:
    st.markdown('### ⚙️ Configurações Logísticas')
    modo_operacao = st.radio('Modo de Operação do Motor:', ['📌 Padrão (Respeitar Limites)', '🚀 Contínuo (Roteirizar 100%)'], index=0, disabled=is_locked)
    is_continuo = 'Contínuo' in modo_operacao

    with st.expander('Capacidade e Prazos', expanded=True):
        trava_global = st.number_input('Trava Total de Tarefas:', min_value=0, value=0, step=50, disabled=is_locked)
        obras_dia = st.number_input('Cota Diária por Equipe:', min_value=1, value=25, disabled=is_locked)
        tempo_medio_obra = st.number_input('Tempo médio por obra (min):', min_value=1.0, max_value=480.0, value=45.0, step=5.0, disabled=is_locked)
        tempo_sp_por_obra = st.checkbox('Super Ponto consome tempo por obra agrupada', value=True, disabled=is_locked)
        tpc = st.radio('Visão de Trabalho:', ['Dia', 'Semana'], index=1, disabled=is_locked)
        limite_per = st.number_input(f'Qtd de {tpc}s de Rota:', min_value=1, value=1, disabled=is_locked or is_continuo)
        if is_continuo:
            st.caption('*(Limite de períodos ignorado no modo Contínuo; jornada e cota diária continuam válidas.)*')
        dias_sel = st.multiselect('Dias Úteis na Semana:', DIAS_NOMES, default=DIAS_NOMES[:5], disabled=is_locked)
        data_ini = st.date_input('📅 Data de Início:', value=datetime.today(), disabled=is_locked)

        c_hi, c_hf = st.columns(2)
        with c_hi:
            hora_inicio = st.time_input('Início jornada', value=dt_time(8, 0), disabled=is_locked)
        with c_hf:
            hora_fim = st.time_input('Fim jornada', value=dt_time(18, 0), disabled=is_locked)
        usar_almoco = st.checkbox('Considerar intervalo de almoço', value=True, disabled=is_locked)
        if usar_almoco:
            c_ai, c_af = st.columns(2)
            with c_ai:
                almoco_inicio = st.time_input('Início almoço', value=dt_time(12, 0), disabled=is_locked)
            with c_af:
                almoco_fim = st.time_input('Fim almoço', value=dt_time(13, 0), disabled=is_locked)
        else:
            almoco_inicio, almoco_fim = dt_time(12, 0), dt_time(13, 0)

        st.markdown('---')
        sentido_rota = st.radio('Sentido do Roteamento:', ['📍 Lógica Padrão', '🎯 Varredura Reversa'], index=0, disabled=is_locked)
        raio_sp = st.slider('Raio Super Ponto (m):', 10, 500, 50, 10, disabled=is_locked)
        deduplicar_notas = st.checkbox('Remover ocorrências duplicadas da mesma NOTA', value=True, disabled=is_locked)

    with st.expander('📡 Conexão de Rede', expanded=False):
        url_osrm = st.text_input('Endpoint OSRM:', value='http://router.project-osrm.org', disabled=is_locked)
        usa_osrm = st.checkbox('🛣️ Traçado de Ruas Real', value=False, disabled=is_locked)
        tentativas_osrm = st.number_input('Tentativas por segmento OSRM:', min_value=1, max_value=3, value=2, step=1, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_san.empty:
        cfg_res = st.session_state.get('config_execucao_san', {})
        id_exec = cfg_res.get('id_execucao', 'SAN')
        st.caption(f'Execução: **{id_exec}** | Regras: **{VERSAO_REGRAS_SANEAMENTO}**')

        if 'bytes_zip_xl_san' not in st.session_state:
            if st.button('⚙️ Gerar Planilhas Excel', use_container_width=True):
                with st.spinner('Gerando Excel...'):
                    st.session_state.bytes_zip_xl_san = gerar_zip_excel_atual()
                tentar_rerun()
        else:
            st.download_button('🌐 Baixar Planilhas (ZIP)', data=st.session_state.bytes_zip_xl_san, file_name=f'Saneamento_Planilhas_{id_exec}.zip', use_container_width=True)

        if 'bytes_zip_kml_san' not in st.session_state:
            if st.button('⚙️ Gerar Mapas KML', use_container_width=True):
                with st.spinner('Gerando KML...'):
                    st.session_state.bytes_zip_kml_san = gerar_zip_kml_atual()
                tentar_rerun()
        else:
            st.download_button('🗺️ Baixar Mapas (KML)', data=st.session_state.bytes_zip_kml_san, file_name=f'Saneamento_Mapas_{id_exec}.zip', use_container_width=True)

        if 'bytes_zip_gpx_san' not in st.session_state:
            if st.button('⚙️ Gerar GPS GPX', use_container_width=True):
                with st.spinner('Gerando GPX...'):
                    st.session_state.bytes_zip_gpx_san = gerar_zip_gpx_atual()
                tentar_rerun()
        else:
            st.download_button('🛰️ Baixar GPS (GPX)', data=st.session_state.bytes_zip_gpx_san, file_name=f'Saneamento_GPS_{id_exec}.zip', use_container_width=True)

        if st.button('🧹 Nova Roteirização', type='primary', use_container_width=True):
            limpar_roteirizador()


# ==============================================================
# RESULTADO
# ==============================================================
if is_done and not st.session_state.df_routed_san.empty:
    st.markdown('## 🎯 Resultado do Planejamento')
    cfg_res = st.session_state.get('config_execucao_san', {})
    st.caption(f"ID: **{cfg_res.get('id_execucao', '-')}** | Versão das regras: **{VERSAO_REGRAS_SANEAMENTO}**")

    df_c = st.session_state.get('df_correcao_san', pd.DataFrame())
    df_bases_c = st.session_state.get('df_bases_correcao_san', pd.DataFrame())
    if not df_c.empty:
        st.warning(f'⚠️ {len(df_c)} obras foram retidas por problemas de coordenadas.')
    if not df_bases_c.empty:
        st.info(f'🧭 {len(df_bases_c)} registros de bases/equipes tiveram correção ou rejeição geográfica auditada.')

    dfr = st.session_state.df_routed_san.copy()
    if 'DISTANCIA_PONTO_ANTERIOR_KM' in dfr.columns:
        dfr['DISTANCIA_PROXIMO_PONTO_KM'] = dfr.groupby(['BASE_ATRIBUIDA', 'DIA_MES'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
        st.session_state.df_routed_san['DISTANCIA_PROXIMO_PONTO_KM'] = dfr['DISTANCIA_PROXIMO_PONTO_KM']
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]

    te = dfr_t['BASE_ATRIBUIDA'].nunique()
    tr_real = sum(peso_tarefa(r) for _, r in dfr_t.iterrows())
    qs_total = int(dfr_t.get('SUPER_PONTO', pd.Series(index=dfr_t.index, dtype='object')).astype(str).str.startswith('SIM').sum())
    km_rod = pd.to_numeric(dfr.get('DISTANCIA_RODOVIARIA_KM', pd.Series(index=dfr.index, dtype=float)), errors='coerce').fillna(0).sum()
    km_est = pd.to_numeric(dfr.get('DISTANCIA_ESTIMADA_KM', pd.Series(index=dfr.index, dtype=float)), errors='coerce').fillna(0).sum()
    km_txt = f'{km_rod:.1f} km rod.' if km_rod > 0 else f'{km_est:.1f} km est.'

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card('Tarefas Planejadas', tr_real, '🎯', '#0D256C', 'rgba(13,37,108,0.12)'), unsafe_allow_html=True)
    c2.markdown(render_metric_card('Equipes Alocadas', te, '👥', '#8b5cf6', 'rgba(139,92,246,0.15)'), unsafe_allow_html=True)
    c3.markdown(render_metric_card('Super Pontos', qs_total, '🏢', '#FF9800', 'rgba(255,152,0,0.15)'), unsafe_allow_html=True)
    c4.markdown(render_metric_card('KM Total Previsto', km_txt, '🛣️', '#55B929', 'rgba(85,185,41,0.15)'), unsafe_allow_html=True)

    df_na = st.session_state.get('df_unallocated_san', pd.DataFrame())
    obras_na = sum(peso_tarefa(r) for _, r in df_na.iterrows()) if not df_na.empty else 0
    dup_rem = len(st.session_state.get('df_duplicadas_removidas_san', pd.DataFrame()))
    seg_falhos = int((dfr.get('STATUS_ROTA', pd.Series(index=dfr.index, dtype='object')).astype(str) == 'SEM_ROTA_OSRM').sum())
    alert_jornada = int((dfr.get('ALERTA_JORNADA', pd.Series(index=dfr.index, dtype='object')).astype(str) == 'SIM').sum())

    q1, q2, q3, q4 = st.columns(4)
    q1.metric('Não alocadas', obras_na)
    q2.metric('Duplicadas removidas', dup_rem)
    q3.metric('Segmentos sem OSRM', seg_falhos)
    q4.metric('Alertas de jornada', alert_jornada)

    if obras_na == 0:
        st.success(f'✅ 100% das {tr_real} tarefas elegíveis após auditorias/trava foram roteirizadas.')
    else:
        st.warning(f'⚠️ {obras_na} tarefas ficaram fora do planejamento. Consulte a aba “Obras Não Alocadas”.')

    st.markdown('### 🗺️ Mapa Operacional')
    mostrar_mapa = st.checkbox('Exibir mapa operacional', value=False, key='mostrar_mapa_san')
    if mostrar_mapa:
        mapa = folium.Map(location=[dfr_t['LATITUDE'].mean(), dfr_t['LONGITUDE'].mean()], zoom_start=8) if not dfr_t.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
        equipes_ordem = list(dict.fromkeys(dfr['BASE_ATRIBUIDA'].dropna().astype(str).tolist()))
        for idx_bn, bn in enumerate(equipes_ordem):
            if bn == 'NÃO ALOCADO':
                continue
            cr = co_f[idx_bn % len(co_f)]
            db = dfr[dfr['BASE_ATRIBUIDA'].astype(str) == bn]
            fg = folium.FeatureGroup(name=f'Rota: {bn}', show=False)
            markers = MarkerCluster(name=f'Obras: {bn}').add_to(fg)
            for _, r in db.iterrows():
                geom = r.get('ROTA_GEOMETRIA')
                if isinstance(geom, list) and len(geom) >= 2:
                    pts = [[float(pt[1]), float(pt[0])] for pt in geom if isinstance(pt, (list, tuple)) and len(pt) >= 2]
                    if len(pts) >= 2:
                        folium.PolyLine(pts, color='black', weight=6, opacity=0.75).add_to(fg)
                        folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
                    continue
                c_i = 'orange' if str(r.get('SUPER_PONTO', '')).startswith('SIM') else ('gray' if r.get('STATUS_ROTA') == 'SEM_ROTA_OSRM' else 'blue')
                ic = identificar_icone_folium(r, dfr.columns)
                er = ''.join([
                    f"<tr><td><b>{html.escape(str(c))}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>"
                    for c in st.session_state.get('colunas_exibir_san', [])
                    if c.upper() not in ['NOME_DIA', 'DIA_MES', 'SEMANA', 'BASE_ATRIBUIDA']
                ])
                extras = f"<tr><td><b>Status rota</b></td><td>{html.escape(str(r.get('STATUS_ROTA','-')))}</td></tr><tr><td><b>Hora</b></td><td>{r.get('HORA_INICIO','-')} - {r.get('HORA_FIM','-')}</td></tr>"
                pop_html = f'<div style="width:280px;"><b>Equipe:</b> {html.escape(str(bn))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{extras}{er}</table></div>'
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=330)).add_to(markers)
            fg.add_to(mapa)
        folium.LayerControl(collapsed=False).add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550)
    else:
        st.caption('O mapa é carregado apenas quando solicitado para manter a tela rápida.')

    t1, t2, t3, t4 = st.tabs(['📊 Dados Tabulares', '📉 Obras Não Alocadas', '🔁 Duplicidades', '🧭 Auditoria de Bases'])
    with t1:
        st.dataframe(dfr.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE'], errors='ignore'), use_container_width=True, hide_index=True)
    with t2:
        if not df_na.empty:
            st.dataframe(df_na, use_container_width=True, hide_index=True)
        else:
            st.success('✅ Nenhuma obra ficou sem alocação.')
    with t3:
        dfd = st.session_state.get('df_duplicadas_san', pd.DataFrame())
        if not dfd.empty:
            st.dataframe(dfd, use_container_width=True, hide_index=True)
        else:
            st.success('✅ Nenhuma NOTA duplicada detectada.')
    with t4:
        if not df_bases_c.empty:
            st.dataframe(df_bases_c, use_container_width=True, hide_index=True)
        else:
            st.success('✅ Nenhuma inconsistência registrada nas bases das equipes.')


# ==============================================================
# ENTRADA / PRÉ-PROCESSAMENTO
# ==============================================================
elif status_exec == 'IDLE':
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown('### 👥 1. Bases de Equipes')
        df_bases = pd.DataFrame()
        bf = st.file_uploader('Suba a planilha de Equipes (Excel)', type=['xlsx', 'xls'])
        if bf:
            b_t = ler_planilha_cached(bf.getvalue())
            b_t.columns = normalize_cols(b_t.columns)
            b_t = b_t.loc[:, ~b_t.columns.duplicated()].copy()
            for pn in ['NOME', 'EQUIPE', 'TECNICO', 'COLABORADOR', 'LEVANTADOR', 'FISCAL']:
                if pn in b_t.columns:
                    b_t = b_t.rename(columns={pn: 'BASE_NOME'})
                    break

            if 'BASE_NOME' in b_t.columns:
                b_t['BASE_NOME'] = b_t['BASE_NOME'].astype(str).str.split(r'\s*\|\s*')
                b_t = b_t.explode('BASE_NOME').reset_index(drop=True)
                b_t['BASE_NOME'] = b_t['BASE_NOME'].astype(str).str.strip().str.upper()
                opts = sorted([str(x) for x in b_t['BASE_NOME'].dropna().unique() if str(x) not in ['SEM EQUIPE', 'NAN', 'NONE', '']])
                sel = st.multiselect('Selecione as Equipes Ativas:', opts, default=opts)
                if sel:
                    df_bases = b_t[b_t['BASE_NOME'].isin(sel)].copy()
                    df_bases['BASE_COORD_FONTE'] = 'PLANILHA'
                    if 'LATITUDE' not in df_bases.columns:
                        df_bases['LATITUDE'] = np.nan
                    if 'LONGITUDE' not in df_bases.columns:
                        df_bases['LONGITUDE'] = np.nan
                    df_bases['LATITUDE_ORIGINAL'] = df_bases['LATITUDE']
                    df_bases['LONGITUDE_ORIGINAL'] = df_bases['LONGITUDE']
                    df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
                    df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).str.replace(',', '.', regex=False), errors='coerce')

                    inval = (
                        df_bases['LATITUDE'].isna() | df_bases['LONGITUDE'].isna() |
                        (df_bases['LATITUDE'] == 0) | (df_bases['LONGITUDE'] == 0) |
                        (df_bases['LATITUDE'] > 0) | (df_bases['LONGITUDE'] > 0) |
                        (df_bases['LATITUDE'].abs() > df_bases['LONGITUDE'].abs())
                    )
                    df_bases['STATUS_COORD_BASE'] = np.where(inval, 'PENDENTE_CORRECAO', 'OK')
                    cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else ('MUNICIPIO' if 'MUNICIPIO' in df_bases.columns else None)
                    audit_bases = []
                    if inval.any() and cr:
                        mapa_coord = {}
                        with st.spinner('🌍 Corrigindo bases inválidas via município...'):
                            for m in df_bases.loc[inval, cr].dropna().unique():
                                mapa_coord[m] = obter_coordenadas_municipio_cached(m)
                        for idx in df_bases.index[inval]:
                            mun = df_bases.at[idx, cr]
                            coord = mapa_coord.get(mun, (np.nan, np.nan))
                            la, lo = coord[0], coord[1]
                            ok_fb = pd.notna(la) and pd.notna(lo) and float(la) < 0 and float(lo) < 0 and abs(float(la)) <= abs(float(lo))
                            if ok_fb:
                                df_bases.at[idx, 'LATITUDE'] = float(la)
                                df_bases.at[idx, 'LONGITUDE'] = float(lo)
                                df_bases.at[idx, 'BASE_COORD_FONTE'] = 'IBGE_FALLBACK'
                                df_bases.at[idx, 'STATUS_COORD_BASE'] = 'CORRIGIDA'
                                rr = df_bases.loc[idx].copy()
                                rr['MOTIVO_AUDITORIA_BASE'] = 'Coordenada da planilha inválida; corrigida por município'
                                audit_bases.append(rr)
                            else:
                                df_bases.at[idx, 'STATUS_COORD_BASE'] = 'REJEITADA'
                                rr = df_bases.loc[idx].copy()
                                rr['MOTIVO_AUDITORIA_BASE'] = 'Coordenada inválida e fallback por município indisponível'
                                audit_bases.append(rr)
                    elif inval.any():
                        for idx in df_bases.index[inval]:
                            df_bases.at[idx, 'STATUS_COORD_BASE'] = 'REJEITADA'
                            rr = df_bases.loc[idx].copy()
                            rr['MOTIVO_AUDITORIA_BASE'] = 'Coordenada inválida e sem município/residência para fallback'
                            audit_bases.append(rr)
                    st.session_state.df_bases_correcao_san = pd.DataFrame(audit_bases)
                    df_bases = df_bases[df_bases['STATUS_COORD_BASE'] != 'REJEITADA'].copy()

                    if 'MUNICIPIO' in df_bases.columns:
                        df_bases['MUN_LIMPO_BASE'] = normalizar_municipios(df_bases['MUNICIPIO'].astype(str).fillna(''))
                    elif 'RESIDENCIA' in df_bases.columns:
                        df_bases['MUN_LIMPO_BASE'] = normalizar_municipios(df_bases['RESIDENCIA'].astype(str).fillna(''))
                    else:
                        df_bases['MUN_LIMPO_BASE'] = ''
            else:
                st.error('❌ A planilha não possui coluna de nome da Equipe/Levantador.')

        st.markdown('##### 📍 Regra de Atribuição')
        ta_index = 1 if is_continuo else 0
        ta = st.radio('Como amarrar as notas aos técnicos?', ['Por Proximidade Espacial', 'Por Município da Planilha'], index=ta_index, disabled=is_continuo, label_visibility='collapsed')
        if is_continuo:
            st.caption('🔒 No modo Contínuo, as obras são vinculadas aos municípios definidos para as equipes.')
        dist_max_atribuicao = st.number_input('Distância máxima por proximidade (km, 0 = ilimitado):', min_value=0.0, max_value=2000.0, value=0.0, step=10.0, disabled=('Município' in ta or is_continuo))

    with c_up2:
        st.markdown('### 📁 2. Demandas (Obras)')
        task_files = st.file_uploader('Suba as planilhas de Demandas', type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

    if df_bases.empty or not task_files:
        st.stop()
    if hora_fim <= hora_inicio:
        st.error('❌ O fim da jornada deve ser posterior ao início.')
        st.stop()
    if usar_almoco and almoco_fim <= almoco_inicio:
        st.error('❌ O fim do almoço deve ser posterior ao início do almoço.')
        st.stop()

    qtd_eq = df_bases['BASE_NOME'].nunique()
    cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per
    sb_html.markdown(render_sidebar_card(cm, 0, qtd_eq, cm * qtd_eq, is_continuo), unsafe_allow_html=True)

    dfs = []
    for ordem_arq, f in enumerate(task_files):
        dft = ler_csv_resiliente(f.getvalue()) if f.name.lower().endswith('.csv') else ler_planilha_cached(f.getvalue())
        dft.columns = normalize_cols(dft.columns)
        dft = dft.loc[:, ~dft.columns.duplicated()].copy()
        dft['ARQUIVO_ORIGEM'] = f.name
        dft['_ORDEM_ARQUIVO'] = ordem_arq
        if not dfs:
            st.session_state.colunas_originais_san = dft.columns.tolist()
        dfs.append(dft)

    df_tasks = pd.concat(dfs, ignore_index=True)
    df_tasks['_ORDEM_ENTRADA'] = np.arange(len(df_tasks))
    if 'LATITUDE PROJETO' in df_tasks.columns and 'LATITUDE' not in df_tasks.columns:
        df_tasks['LATITUDE'] = df_tasks['LATITUDE PROJETO']
    if 'LONGITUDE PROJETO' in df_tasks.columns and 'LONGITUDE' not in df_tasks.columns:
        df_tasks['LONGITUDE'] = df_tasks['LONGITUDE PROJETO']
    for cc in ['NOTA', 'PROTOCOLO', 'OS', 'ID']:
        if cc in df_tasks.columns:
            df_tasks['PROTOCOLO'] = df_tasks[cc]
            break
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True)
        df_tasks['PROTOCOLO'] = limpar_protocolo_serie(df_tasks['PROTOCOLO'])
        df_tasks = df_tasks[df_tasks['PROTOCOLO'] != ''].copy()

    falta = [c for c in ['MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta:
        st.error(f"🚨 Faltam colunas: {', '.join(falta)}. Certifique-se de que a planilha possui coordenadas de projeto.")
        st.stop()

    c1_f, c2_f = st.columns([1, 1])
    with c1_f:
        cs = 'STATUS CLIENTE' if 'STATUS CLIENTE' in df_tasks.columns else ('STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO')
        if cs in df_tasks.columns:
            df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
            opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
            default_s = [s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']] if 'FISCAL' in cs else opts_s
            sel_s = st.multiselect('1. Status Roteirizáveis:', options=opts_s, default=default_s)
            if not sel_s:
                st.stop()
            df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()

    with c2_f:
        if 'TIPO DEMANDA' in df_tasks.columns:
            df_tasks['TIPO DEMANDA'] = df_tasks['TIPO DEMANDA'].astype(str).str.strip()
            opts_n = sorted([str(x) for x in df_tasks['TIPO DEMANDA'].dropna().unique() if str(x).upper() != 'NAN'])
            default_prio = [n for n in opts_n if 'RELIGACAO' in normalizar_texto(n) or 'EMERGENCIA' in normalizar_texto(n)]
            sel_p = st.multiselect('🚨 2. Demandas de Alta Prioridade:', options=opts_n, default=default_prio)
            prio_norm = {normalizar_texto(x) for x in sel_p}
            df_tasks['PRIORIDADE'] = df_tasks['TIPO DEMANDA'].apply(lambda x: 'Sim' if normalizar_texto(x) in prio_norm else 'Não')
        else:
            df_tasks['PRIORIDADE'] = 'Não'

    # Auditoria de duplicidades antes de qualquer agrupamento espacial.
    df_tasks['QTD_OCORRENCIAS_NOTA'] = df_tasks.groupby('PROTOCOLO')['PROTOCOLO'].transform('size')
    df_tasks['QTD_ARQUIVOS_NOTA'] = df_tasks.groupby('PROTOCOLO')['ARQUIVO_ORIGEM'].transform('nunique')
    df_tasks['QTD_OCORRENCIAS_NA_ORIGEM'] = df_tasks.groupby(['PROTOCOLO', 'ARQUIVO_ORIGEM'])['PROTOCOLO'].transform('size')
    df_tasks['DUPLICADA_GERAL'] = np.where(df_tasks['QTD_OCORRENCIAS_NOTA'] > 1, 'SIM', 'NÃO')
    df_tasks['DUPLICADA_ENTRE_ARQUIVOS'] = np.where(df_tasks['QTD_ARQUIVOS_NOTA'] > 1, 'SIM', 'NÃO')
    df_tasks['REPETIDA_NA_ORIGEM'] = np.where(df_tasks['QTD_OCORRENCIAS_NA_ORIGEM'] > 1, 'SIM', 'NÃO')
    df_dup_audit = df_tasks[df_tasks['DUPLICADA_GERAL'] == 'SIM'].copy()
    st.session_state.df_duplicadas_san = df_dup_audit
    df_dup_rem = pd.DataFrame()
    if deduplicar_notas and not df_dup_audit.empty:
        prioridade_ord = df_tasks['PRIORIDADE'].astype(str).str.upper().eq('SIM').astype(int)
        temp = df_tasks.assign(_P=prioridade_ord).sort_values(['PROTOCOLO', '_P', '_ORDEM_ENTRADA'], ascending=[True, False, True])
        manter_idx = temp.drop_duplicates('PROTOCOLO', keep='first').index
        df_dup_rem = df_tasks[~df_tasks.index.isin(manter_idx)].copy()
        if not df_dup_rem.empty:
            df_dup_rem['MOTIVO_NAO_ALOCACAO'] = 'DUPLICIDADE_REMOVIDA'
        df_tasks = df_tasks[df_tasks.index.isin(manter_idx)].copy().sort_values('_ORDEM_ENTRADA')
    st.session_state.df_duplicadas_removidas_san = df_dup_rem

    # Trava global conta tarefas reais, antes de Super Pontos.
    df_tasks, df_fora_trava = aplicar_trava_global(df_tasks, trava_global)

    st.markdown('#### 🌍 Limpeza de Coordenadas Geográficas')
    if st.button('⏹️ Abortar', use_container_width=True):
        limpar_roteirizador()
        st.stop()

    df_tasks, df_rej = validar_coordenadas_obras(df_tasks)
    st.session_state.df_correcao_san = df_rej
    if df_tasks.empty:
        st.error('🚨 Nenhuma obra válida restou.')
        st.stop()

    df_tasks['MUN_LIMPO'] = normalizar_municipios(df_tasks['MUNICIPIO'].astype(str).fillna(''))

    # Super Pontos nunca atravessam município quando a atribuição é municipal.
    if 'Município' in ta:
        partes, qc = [], 0
        for _, grp in df_tasks.groupby('MUN_LIMPO', sort=False, dropna=False):
            agrupado, q = fundir_super_pontos(grp.copy(), raio_metros=raio_sp, agrupar_por_levantador=False)
            partes.append(agrupado)
            qc += int(q or 0)
        df_tasks = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    else:
        df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_sp, agrupar_por_levantador=False)

    if '_ORDEM_ENTRADA' not in df_tasks.columns:
        df_tasks['_ORDEM_ENTRADA'] = np.arange(len(df_tasks))
    else:
        df_tasks['_ORDEM_ENTRADA'] = pd.to_numeric(df_tasks['_ORDEM_ENTRADA'], errors='coerce')
        falt_ord = df_tasks['_ORDEM_ENTRADA'].isna()
        if falt_ord.any():
            df_tasks.loc[falt_ord, '_ORDEM_ENTRADA'] = np.arange(len(df_tasks))[falt_ord.to_numpy()]

    # Super Ponto herda a maior prioridade e a primeira ordem das obras que contém.
    if '_ORIGINAL_ROWS' in df_tasks.columns:
        for idx, rr in df_tasks.iterrows():
            origs = rr.get('_ORIGINAL_ROWS')
            if isinstance(origs, list) and origs:
                if any(str(o.get('PRIORIDADE', '')).strip().upper() == 'SIM' for o in origs):
                    df_tasks.at[idx, 'PRIORIDADE'] = 'Sim'
                ords = [pd.to_numeric(o.get('_ORDEM_ENTRADA'), errors='coerce') for o in origs]
                ords = [float(x) for x in ords if pd.notna(x)]
                if ords:
                    df_tasks.at[idx, '_ORDEM_ENTRADA'] = min(ords)

    tbr = df_bases.to_dict('records')
    fiscal_anchors = {}
    for b in tbr:
        fn = b['BASE_NOME']
        if fn not in fiscal_anchors:
            fiscal_anchors[fn] = (float(b['LATITUDE']), float(b['LONGITUDE']))

    team_mun_counts = {b['BASE_NOME']: {} for b in tbr}
    assigned_tasks, unassigned_tasks = [], []
    # Excedente da trava é explicitamente não alocado.
    if not df_fora_trava.empty:
        unassigned_tasks.extend(df_fora_trava.to_dict('records'))

    # Ordem estável e prioridade primeiro.
    df_tasks = df_tasks.assign(_P=df_tasks.get('PRIORIDADE', pd.Series('Não', index=df_tasks.index)).astype(str).str.upper().eq('SIM').astype(int))
    df_tasks = df_tasks.sort_values(['_P', 'LATITUDE', 'LONGITUDE', '_ORDEM_ENTRADA'], ascending=[False, True, True, True]).drop(columns=['_P'])

    for r in df_tasks.to_dict('records'):
        la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
        ms = r.get('MUN_LIMPO', '')
        qr = peso_tarefa(r)
        best_f, best_d = None, float('inf')

        if 'Município' in ta:
            candidatos = [b for b in tbr if b.get('MUN_LIMPO_BASE', '') == ms]
            melhor_score = None
            for b in candidatos:
                f_name = b['BASE_NOME']
                carga_atual = team_mun_counts.get(f_name, {}).get(ms, 0)
                carga_projetada = carga_atual + qr
                b_lat, b_lon = fiscal_anchors.get(f_name, (float(b['LATITUDE']), float(b['LONGITUDE'])))
                d = float(haversine_scalar(la, lo, b_lat, b_lon))
                score = (carga_projetada, d, f_name)
                if melhor_score is None or score < melhor_score:
                    melhor_score = score
                    best_f, best_d = f_name, d
        else:
            for b in tbr:
                f_name = b['BASE_NOME']
                b_lat, b_lon = fiscal_anchors.get(f_name, (float(b['LATITUDE']), float(b['LONGITUDE'])))
                d = float(haversine_scalar(la, lo, b_lat, b_lon))
                if d < best_d:
                    best_d, best_f = d, f_name
            if best_f and float(dist_max_atribuicao) > 0 and best_d > float(dist_max_atribuicao):
                best_f = None
                r['MOTIVO_NAO_ALOCACAO'] = f'DISTANCIA_ACIMA_LIMITE ({best_d:.1f} km)'

        if best_f:
            r['BASE_ATRIBUIDA'] = best_f
            r['DISTANCIA_ATRIBUICAO_KM'] = round(best_d, 3)
            assigned_tasks.append(r)
            fiscal_anchors[best_f] = (float(la), float(lo))
            team_mun_counts.setdefault(best_f, {})
            team_mun_counts[best_f][ms] = team_mun_counts[best_f].get(ms, 0) + qr
        else:
            r.setdefault('MOTIVO_NAO_ALOCACAO', 'SEM_EQUIPE_COMPATIVEL')
            r['BASE_ATRIBUIDA'] = 'NÃO ALOCADO'
            unassigned_tasks.append(r)

    df_ta = pd.DataFrame(assigned_tasks)
    df_u = pd.DataFrame(unassigned_tasks)
    st.session_state.df_unallocated_san = df_u
    total_alocadas = sum(peso_tarefa(r) for _, r in df_ta.iterrows()) if not df_ta.empty else 0
    sb_html.markdown(render_sidebar_card(cm, total_alocadas, qtd_eq, cm * qtd_eq, is_continuo), unsafe_allow_html=True)
    if df_ta.empty:
        st.error('Nenhuma obra pôde ser alocada às equipes. Verifique municípios, coordenadas e limite de distância.')
        st.stop()

    with st.expander('🛠️ Configuração de Saída (Colunas)', expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = ['NOTA', 'STATUS CLIENTE', 'NOME', 'TIPO DEMANDA', 'MUNICIPIO', 'ENDERECO', 'BAIRRO', 'PONTO REFERENCIA', 'COMPLEMENTO', 'LATITUDE PROJETO', 'LONGITUDE PROJETO', 'CLASSIFICACAO AREA', 'TEL FIXO', 'TEL MOVEL', 'GRUPO TENSAO', 'ARQUIVO_ORIGEM']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect('Colunas que vão aparecer no Mapa e Excel:', tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button('🚀 Iniciar Motor de Roteirização', type='primary', use_container_width=True):
        id_exec = criar_id_execucao()
        b_names = list(dict.fromkeys([b['BASE_NOME'] for b in tbr]))
        config_exec = {
            'id_execucao': id_exec,
            'versao_regras': VERSAO_REGRAS_SANEAMENTO,
            'velocidade_media_kmh': 30.0,
            'obras_por_dia': int(obras_dia),
            'tempo_medio_obra_min': float(tempo_medio_obra),
            'tempo_super_ponto_por_obra': bool(tempo_sp_por_obra),
            'tipo_periodo': tpc,
            'limite_periodos': int(limite_per),
            'dias_selecionados': list(dias_sel),
            'url_osrm_base': url_osrm,
            'tracado_real': bool(usa_osrm),
            'tentativas_osrm': int(tentativas_osrm),
            'data_inicio': data_ini,
            'hora_inicio': hora_inicio,
            'hora_fim': hora_fim,
            'hora_inicio_str': hora_inicio.strftime('%H:%M'),
            'hora_fim_str': hora_fim.strftime('%H:%M'),
            'usar_intervalo_almoco': bool(usar_almoco),
            'almoco_inicio': almoco_inicio,
            'almoco_fim': almoco_fim,
            'sentido_rota': sentido_rota,
            'modo_continuo': bool(is_continuo),
            'regra_atribuicao': ta,
            'distancia_max_atribuicao_km': float(dist_max_atribuicao),
            'raio_super_ponto_m': int(raio_sp),
            'deduplicar_notas': bool(deduplicar_notas),
            'trava_global': int(trava_global),
            'arquivo_equipes': bf.name if bf else '-',
            'arquivos_demanda': [f.name for f in task_files],
            'qtd_entrada_filtrada': int(len(df_tasks)),
            'qtd_duplicidades': int(df_dup_audit['PROTOCOLO'].nunique()) if not df_dup_audit.empty else 0,
            'qtd_linhas_duplicadas': int(len(df_dup_audit)),
            'qtd_duplicidades_removidas': int(len(df_dup_rem)),
            'qtd_correcao_obras': int(len(df_rej)),
            'qtd_correcao_bases': int(len(st.session_state.get('df_bases_correcao_san', pd.DataFrame()))),
        }
        st.session_state.config_execucao_san = config_exec
        st.session_state.update({'bases_records_san': tbr, 'colunas_exibir_san': colunas_exibir})
        st.session_state.vrp_state_san = {
            'config': config_exec,
            'b_names': b_names,
            'b_idx': 0,
            'unvisited': df_ta.copy(),
            'routed_data': [],
            'osrm_cache': {},
        }
        for k in ['bytes_zip_xl_san', 'bytes_zip_kml_san', 'bytes_zip_gpx_san']:
            st.session_state.pop(k, None)
        st.session_state.vrp_status_san = 'RUNNING'
        tentar_rerun()


# ==============================================================
# MOTOR
# ==============================================================
if status_exec == 'RUNNING':
    st.markdown('## 🚀 Execução do Motor Saneamento')
    if st.button('⏹️ Abortar Execução', use_container_width=True):
        limpar_roteirizador()

    st_run = st.session_state.get('start_time_run_san', time.time())
    if 'start_time_run_san' not in st.session_state:
        st.session_state.start_time_run_san = st_run

    pb = st.progress(0.0)
    tmp = st.empty()
    sgt = st.empty()
    st_v = st.session_state.vrp_state_san
    cfg = st_v['config']
    b_n = st_v['b_names']
    b_i = st_v.get('b_idx', 0)

    def render_t(bi, ii, it):
        e = time.time() - st_run
        f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else 'Calc...'
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex;gap:15px;margin-bottom:20px;"><div style="flex:1;padding:20px;border-radius:10px;background:#f8f9fa;border:1px solid #dee2e6;text-align:center;"><b>⏱️ Decorrido</b><div style="font-size:2rem;color:#0D256C;font-weight:bold;">{es}</div></div><div style="flex:1;padding:20px;border-radius:10px;background:#e8f5e9;border:1px solid #a5d6a7;text-align:center;"><b>🎯 Restante</b><div style="font-size:2rem;color:#1b5e20;font-weight:bold;">{rs}</div></div></div>', unsafe_allow_html=True)

    try:
        if b_i < len(b_n):
            bn = b_n[b_i]
            pb.progress(b_i / max(1, len(b_n)))
            sgt.info(f'🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})')
            render_t(b_i, 0, 1)

            if 'c_rotas' not in st_v:
                br_all = pd.DataFrame(st.session_state.bases_records_san)
                br_rows = br_all[br_all['BASE_NOME'] == bn]
                if br_rows.empty or pd.isna(br_rows.iloc[0].get('LATITUDE')):
                    st_v['b_idx'] += 1
                    st.session_state.vrp_state_san = st_v
                    tentar_rerun()
                    st.stop()
                br = br_rows.iloc[0]
                bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
                oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
                rf, nao_agendadas = construir_plano_rota_equipe(oe, bl, bL, cfg)
                if nao_agendadas:
                    st.session_state.df_unallocated_san = pd.concat([
                        st.session_state.get('df_unallocated_san', pd.DataFrame()), pd.DataFrame(nao_agendadas)
                    ], ignore_index=True)
                st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []
                st.session_state.vrp_state_san = st_v
                tentar_rerun()
                st.stop()

            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (20 if cfg['tracado_real'] else max(1, len(rf))), len(rf))
            cache = st_v.setdefault('osrm_cache', {})
            for i in range(oi, ei):
                it = rf[i]
                if not cfg['tracado_real']:
                    geom = [[it['La'], it['la']], [it['Lt'], it['lt']]]
                    gd.append({'geom': geom, 'duracao_s': float(it['tempo_estimado_min']) * 60.0, 'dist_rod_km': np.nan, 'status': 'ESTIMADA_SEM_OSRM'})
                    continue

                if i % 5 == 0:
                    sgt.info(f'🛣️ Traçando arruamento **{bn}**... ({i+1}/{len(rf)})')
                render_t(b_i, i, len(rf))
                chave = (round(float(it['la']), 5), round(float(it['La']), 5), round(float(it['lt']), 5), round(float(it['Lt']), 5), str(cfg['url_osrm_base']))
                if chave in cache:
                    gd.append(cache[chave])
                    continue

                resultado = None
                for tentativa in range(int(cfg.get('tentativas_osrm', 2))):
                    try:
                        geom, dur_s = obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'])
                        if isinstance(geom, list) and len(geom) >= 2:
                            dist_rod = distancia_geometria_km(geom)
                            resultado = {'geom': geom, 'duracao_s': float(dur_s), 'dist_rod_km': dist_rod, 'status': 'OK_OSRM'}
                            cache[chave] = resultado  # somente sucessos ficam no cache
                            break
                    except Exception:
                        if tentativa + 1 < int(cfg.get('tentativas_osrm', 2)):
                            time.sleep(0.15)
                if resultado is None:
                    resultado = {'geom': [], 'duracao_s': float(it['tempo_estimado_min']) * 60.0, 'dist_rod_km': np.nan, 'status': 'SEM_ROTA_OSRM'}
                gd.append(resultado)

            st_v['c_idx'], st_v['current_geoms'], st_v['osrm_cache'] = ei, gd, cache
            if ei < len(rf):
                st.session_state.vrp_state_san = st_v
                tentar_rerun()
                st.stop()

            # Converte plano + geometrias em saída final. A geometria do primeiro trecho é preservada.
            rdf = []
            ordem_por_dia = {}
            relogio_por_dia = {}
            fim_jornada_por_dia = {}
            for it, info_geo in zip(rf, gd):
                chave_dia = (it['d'], it['dm'])
                if chave_dia not in relogio_por_dia:
                    data_dt = datetime.strptime(it['dm'], '%d/%m/%Y').date()
                    relogio_por_dia[chave_dia] = datetime.combine(data_dt, cfg['hora_inicio'])
                    fim_jornada_por_dia[chave_dia] = datetime.combine(data_dt, cfg['hora_fim'])
                    ordem_por_dia[chave_dia] = 1

                tempo_rota_min = float(info_geo['duracao_s']) / 60.0
                inicio_desloc = relogio_por_dia[chave_dia]
                chegada = inicio_desloc + pd.Timedelta(minutes=tempo_rota_min)
                status_rota = info_geo['status']
                dist_rod = info_geo['dist_rod_km']

                if it.get('ir', False):
                    fim = chegada
                    alerta_j = 'SIM' if fim > fim_jornada_por_dia[chave_dia] else 'NÃO'
                    rdf.append({
                        'PROTOCOLO': 'RETORNO_BASE', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'],
                        'BASE_ATRIBUIDA': bn, 'ORDEM': ordem_por_dia[chave_dia], 'NOME_DIA': it['dn'],
                        'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': it['periodo'],
                        'DISTANCIA_PONTO_ANTERIOR_KM': round(float(dist_rod) if pd.notna(dist_rod) else float(it['distancia_estimada_km']), 2),
                        'DISTANCIA_ESTIMADA_KM': round(float(it['distancia_estimada_km']), 3),
                        'DISTANCIA_RODOVIARIA_KM': round(float(dist_rod), 3) if pd.notna(dist_rod) else np.nan,
                        'TEMPO_ROTA_MIN': round(tempo_rota_min, 2), 'TEMPO_ATENDIMENTO_MIN': 0.0, 'STATUS_ROTA': status_rota,
                        'ROTA_GEOMETRIA': info_geo['geom'], 'PRIORIDADE': 'Não',
                        'HORA_INICIO': inicio_desloc.strftime('%H:%M'), 'HORA_FIM': fim.strftime('%H:%M'),
                        '_HORA_INICIO_DT': inicio_desloc, '_HORA_FIM_DT': fim, 'ALERTA_JORNADA': alerta_j,
                    })
                    relogio_por_dia[chave_dia] = fim
                    ordem_por_dia[chave_dia] += 1
                    continue

                ob = dict(it['o'])
                inicio_serv = aplicar_intervalo_almoco(chegada, it['servico_min'], cfg)
                fim_serv = inicio_serv + pd.Timedelta(minutes=float(it['servico_min']))
                alerta_j = 'SIM' if fim_serv > fim_jornada_por_dia[chave_dia] else 'NÃO'
                ob.update({
                    'ORDEM': ordem_por_dia[chave_dia], 'NOME_DIA': it['dn'], 'DIA_MES': it['dm'],
                    'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': it['periodo'],
                    'DISTANCIA_PONTO_ANTERIOR_KM': round(float(dist_rod) if pd.notna(dist_rod) else float(it['distancia_estimada_km']), 2),
                    'DISTANCIA_ESTIMADA_KM': round(float(it['distancia_estimada_km']), 3),
                    'DISTANCIA_RODOVIARIA_KM': round(float(dist_rod), 3) if pd.notna(dist_rod) else np.nan,
                    'TEMPO_ROTA_MIN': round(tempo_rota_min, 2), 'TEMPO_ATENDIMENTO_MIN': round(float(it['servico_min']), 2), 'STATUS_ROTA': status_rota,
                    'ROTA_GEOMETRIA': info_geo['geom'],
                    'HORA_INICIO': inicio_serv.strftime('%H:%M'), 'HORA_FIM': fim_serv.strftime('%H:%M'),
                    '_HORA_INICIO_DT': inicio_serv, '_HORA_FIM_DT': fim_serv, 'ALERTA_JORNADA': alerta_j,
                })
                rdf.append(ob)
                relogio_por_dia[chave_dia] = fim_serv
                ordem_por_dia[chave_dia] += 1

            st_v['routed_data'].extend(rdf)
            for chave in ['c_rotas', 'c_idx', 'current_geoms']:
                st_v.pop(chave, None)
            st_v['b_idx'] += 1
            st.session_state.vrp_state_san = st_v
            gc.collect()
            tentar_rerun()
        else:
            sgt.success('✅ Rotas Finalizadas!')
            pb.progress(1.0)
            st.session_state.df_routed_san = pd.DataFrame(st_v['routed_data'])
            cfg_final = st.session_state.get('config_execucao_san', {}).copy()
            cfg_final['qtd_nao_alocadas'] = int(sum(peso_tarefa(r) for _, r in st.session_state.get('df_unallocated_san', pd.DataFrame()).iterrows())) if not st.session_state.get('df_unallocated_san', pd.DataFrame()).empty else 0
            cfg_final['tempo_processamento_s'] = round(time.time() - st_run, 2)
            st.session_state.config_execucao_san = cfg_final
            st.session_state.roteamento_concluido_san = True
            st.session_state.vrp_status_san = 'IDLE'
            st.session_state.pop('start_time_run_san', None)
            tentar_rerun()

    except Exception:
        erro_id = f"SANERR-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
        print(f'[{erro_id}] Falha no Roteirizador Saneamento:\n{traceback.format_exc()}')
        st.error(f'🚨 Não foi possível concluir esta etapa. Código para suporte: {erro_id}')
        st.session_state.vrp_status_san = 'IDLE'
