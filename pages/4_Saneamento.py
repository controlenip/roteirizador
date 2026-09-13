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

from modules.saneamento_engine import (
    VERSAO_REGRAS_SANEAMENTO, DIAS_NOMES, DIAS_MAP, remover_acentos_str, normalizar_texto,
    limpar_protocolo_serie, peso_tarefa, validar_coordenadas_obras, auditar_duplicidades,
    deduplicar_apenas_exatas, aplicar_trava_global, preparar_ids_equipes, atribuir_tarefas_equipes,
    aplicar_config_equipe, capacidade_total_equipe, tempo_servico_registro, distancia_geometria_km,
    estimar_deslocamento, aplicar_intervalo_almoco, data_trabalho_por_indice, periodo_do_dia,
    ordenar_bloco_rota, construir_plano_rota_equipe, detectar_reagendamento_osrm,
    calcular_reconciliacao, montar_dashboard_equipes, validar_par_coordenadas,
)

st.set_page_config(page_title="Saneamento", page_icon="🧹", layout="wide")
injetar_logo()


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


def parse_tempos_demanda_texto(texto):
    """Lê regras no formato TIPO=MIN; OUTRO=MIN, mantendo vazio como configuração padrão."""
    regras = {}
    for item in re.split(r'[;\n]+', str(texto or '')):
        item = item.strip()
        if not item or '=' not in item:
            continue
        chave, valor = item.split('=', 1)
        try:
            minutos = float(str(valor).strip().replace(',', '.'))
        except Exception:
            continue
        if minutos > 0 and normalizar_texto(chave):
            regras[normalizar_texto(chave)] = minutos
    return regras


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






















def montar_manifesto(config, df_routed=None):
    rec = config.get('reconciliacao', {}) or {}
    linhas = [
        'MANIFESTO - ROTEIRIZADOR SANEAMENTO',
        f"ID execução: {config.get('id_execucao', '-')}",
        f"Versão das regras: {VERSAO_REGRAS_SANEAMENTO}",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Modo: {'CONTÍNUO' if config.get('modo_continuo') else 'PADRÃO'}",
        f"Atribuição: {config.get('regra_atribuicao', '-')}",
        f"Modo de âncora: {config.get('modo_ancora', '-')}",
        f"Cota diária padrão: {config.get('obras_por_dia', '-')}",
        f"Visão: {config.get('tipo_periodo', '-')}",
        f"Limite de períodos: {config.get('limite_periodos', '-')}",
        f"Dias padrão: {', '.join(config.get('dias_selecionados', []))}",
        f"Início da jornada padrão: {config.get('hora_inicio_str', '-')}",
        f"Fim da jornada padrão: {config.get('hora_fim_str', '-')}",
        f"Tempo médio/obra padrão: {config.get('tempo_medio_obra_min', '-')} min",
        f"Tempos por tipo de demanda: {config.get('tempos_por_tipo_demanda', {})}",
        f"Raio Super Ponto: {config.get('raio_super_ponto_m', '-')} m",
        f"Distância máx. atribuição por proximidade: {config.get('distancia_max_atribuicao_km', '-')} km",
        f"OSRM: {'SIM' if config.get('tracado_real') else 'NÃO'}",
        f"Endpoint OSRM autorizado: {config.get('url_osrm_base', '-')}",
        f"Remoção automática somente de duplicatas exatas: {'SIM' if config.get('deduplicar_notas') else 'NÃO'}",
        f"Trava global: {config.get('trava_global', 0)}",
        f"Arquivos de demanda: {', '.join(config.get('arquivos_demanda', []))}",
        f"Arquivo equipes: {config.get('arquivo_equipes', '-')}",
        f"Entradas totais após explosão de protocolos: {config.get('qtd_entrada_total', '-')}",
        f"Registros sem nota: {config.get('qtd_sem_nota', '-')}",
        f"Registros fora do filtro: {config.get('qtd_fora_filtro', '-')}",
        f"Duplicidades detectadas: {config.get('qtd_duplicidades', '-')}",
        f"Duplicidades divergentes: {config.get('qtd_duplicidades_divergentes', '-')}",
        f"Duplicidades exatas removidas: {config.get('qtd_duplicidades_removidas', '-')}",
        f"Coordenadas de obras rejeitadas: {config.get('qtd_correcao_obras', '-')}",
        f"Bases rejeitadas/corrigidas: {config.get('qtd_correcao_bases', '-')}",
        f"Fora da trava global: {config.get('qtd_fora_trava', '-')}",
        f"Não alocadas operacionais: {config.get('qtd_nao_alocadas', '-')}",
        f"Status de elegibilidade localizado: {'SIM' if config.get('status_col_localizada', False) else 'NÃO'}",
        f"Reconciliação: {'OK' if rec.get('ok') else 'ERRO'} | diferença={rec.get('diferenca', '-')}",
    ]
    if df_routed is not None and not df_routed.empty:
        obras = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
        linhas.append(f"Obras roteirizadas: {sum(peso_tarefa(r) for _, r in obras.iterrows())}")
        linhas.append(f"Distância híbrida total: {pd.to_numeric(df_routed.get('DISTANCIA_HIBRIDA_KM', 0), errors='coerce').fillna(0).sum():.2f} km")
        seg = len(df_routed)
        ok = int((df_routed.get('STATUS_ROTA', pd.Series(index=df_routed.index, dtype='object')).astype(str) == 'OK_OSRM').sum())
        linhas.append(f"Cobertura OSRM: {(100.0 * ok / seg if seg else 0.0):.1f}%")
    if rec.get('componentes'):
        linhas.append('COMPONENTES DA RECONCILIAÇÃO:')
        for k, v in rec['componentes'].items():
            linhas.append(f'- {k}: {v}')
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
        'df_duplicadas_removidas_san', 'df_duplicadas_divergentes_san', 'df_fora_filtro_san', 'df_sem_nota_san', 'df_fora_trava_san', 'df_dashboard_equipes_san', 'config_execucao_san', 'mostrar_mapa_san'
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
            seg_b = df_routed[df_routed['BASE_ATRIBUIDA'] == b]
            ok_b = int((seg_b['STATUS_ROTA'].astype(str) == 'OK_OSRM').sum()) if 'STATUS_ROTA' in seg_b.columns else 0
            res.append({
                'Equipe': b,
                'Obras Roteirizadas': sum(peso_tarefa(r) for _, r in db.iterrows()),
                'Super Pontos': qs,
                'KM Híbrido Previsto': round(pd.to_numeric(seg_b.get('DISTANCIA_HIBRIDA_KM', pd.Series(index=seg_b.index, dtype=float)), errors='coerce').fillna(0).sum(), 2),
                'KM Estimado': round(pd.to_numeric(seg_b.get('DISTANCIA_ESTIMADA_KM', pd.Series(index=seg_b.index, dtype=float)), errors='coerce').fillna(0).sum(), 2),
                'KM Rodoviário Conhecido': round(pd.to_numeric(seg_b.get('DISTANCIA_RODOVIARIA_KM', pd.Series(index=seg_b.index, dtype=float)), errors='coerce').fillna(0).sum(), 2),
                'Cobertura OSRM (%)': round(100.0 * ok_b / len(seg_b), 1) if len(seg_b) else 0.0,
                'Segmentos sem OSRM': int((seg_b.get('STATUS_ROTA', pd.Series(index=seg_b.index, dtype='object')).astype(str) == 'SEM_ROTA_OSRM').sum()),
            })
        zx.writestr(f'Resumo_Operacional_{id_exec}.xlsx', gerar_excel_resumo_saneamento(pd.DataFrame(res)))

        rec = cfg.get('reconciliacao', {}) or {}
        rec_df = pd.DataFrame([
            {'INDICADOR': 'TOTAL_ENTRADA', 'VALOR': rec.get('total_entrada', 0)},
            *[{'INDICADOR': k, 'VALOR': v} for k, v in (rec.get('componentes', {}) or {}).items()],
            {'INDICADOR': 'SOMA_SAIDAS', 'VALOR': rec.get('soma_saidas', 0)},
            {'INDICADOR': 'DIFERENCA', 'VALOR': rec.get('diferenca', 0)},
            {'INDICADOR': 'STATUS', 'VALOR': 'OK' if rec.get('ok') else 'ERRO'},
        ]) if rec else pd.DataFrame()
        extras = [
            ('Reconciliacao', rec_df),
            ('Obras_Correcao', st.session_state.get('df_correcao_san', pd.DataFrame())),
            ('Bases_Correcao', st.session_state.get('df_bases_correcao_san', pd.DataFrame())),
            ('Obras_Nao_Alocadas', st.session_state.get('df_unallocated_san', pd.DataFrame())),
            ('Fora_Trava_Global', st.session_state.get('df_fora_trava_san', pd.DataFrame())),
            ('Fora_Filtro', st.session_state.get('df_fora_filtro_san', pd.DataFrame())),
            ('Sem_Nota', st.session_state.get('df_sem_nota_san', pd.DataFrame())),
            ('Auditoria_Duplicadas', st.session_state.get('df_duplicadas_san', pd.DataFrame())),
            ('Duplicadas_Divergentes', st.session_state.get('df_duplicadas_divergentes_san', pd.DataFrame())),
            ('Duplicadas_Exatas_Removidas', st.session_state.get('df_duplicadas_removidas_san', pd.DataFrame())),
            ('Utilizacao_Equipes', st.session_state.get('df_dashboard_equipes_san', pd.DataFrame())),
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
        obras_dia = st.number_input('Cota Diária Padrão por Equipe:', min_value=1, value=25, disabled=is_locked)
        tempo_medio_obra = st.number_input('Tempo médio padrão por obra (min):', min_value=1.0, max_value=480.0, value=45.0, step=5.0, disabled=is_locked)
        tempos_demanda_txt = st.text_area(
            'Tempos por tipo de demanda (opcional):',
            value='',
            placeholder='RELIGACAO=30; EMERGENCIA=30; INSPECAO=20',
            help='Regras por texto normalizado. Se vazio, usa o tempo médio padrão.',
            disabled=is_locked,
        )
        tempo_sp_por_obra = st.checkbox('Super Ponto consome tempo por obra agrupada', value=True, disabled=is_locked)
        tpc = st.radio('Visão de Trabalho:', ['Dia', 'Semana'], index=1, disabled=is_locked)
        limite_per = st.number_input(f'Qtd de {tpc}s de Rota:', min_value=1, value=1, disabled=is_locked or is_continuo)
        if is_continuo:
            st.caption('*(Limite de períodos ignorado no modo Contínuo; jornada e cota diária continuam válidas.)*')
        dias_sel = st.multiselect('Dias Úteis Padrão na Semana:', DIAS_NOMES, default=DIAS_NOMES[:5], disabled=is_locked)
        data_ini = st.date_input('📅 Data de Início:', value=datetime.today(), disabled=is_locked)

        c_hi, c_hf = st.columns(2)
        with c_hi:
            hora_inicio = st.time_input('Início jornada padrão', value=dt_time(8, 0), disabled=is_locked)
        with c_hf:
            hora_fim = st.time_input('Fim jornada padrão', value=dt_time(18, 0), disabled=is_locked)
        usar_almoco = st.checkbox('Considerar intervalo de almoço', value=True, disabled=is_locked)
        if usar_almoco:
            c_ai, c_af = st.columns(2)
            with c_ai:
                almoco_inicio = st.time_input('Início almoço', value=dt_time(12, 0), disabled=is_locked)
            with c_af:
                almoco_fim = st.time_input('Fim almoço', value=dt_time(13, 0), disabled=is_locked)
        else:
            almoco_inicio, almoco_fim = dt_time(12, 0), dt_time(13, 0)

        st.caption('A planilha de equipes pode sobrescrever por equipe: COTA_DIA, HORA_INICIO, HORA_FIM e DIAS_TRABALHO.')
        st.markdown('---')
        sentido_rota = st.radio('Sentido do Roteamento:', ['📍 Lógica Padrão', '🎯 Varredura Reversa'], index=0, disabled=is_locked)
        raio_sp = st.slider('Raio Super Ponto (m):', 10, 500, 50, 10, disabled=is_locked)
        deduplicar_notas = st.checkbox('Remover automaticamente apenas duplicatas EXATAS', value=True, disabled=is_locked)

    with st.expander('📡 Conexão de Rede', expanded=False):
        endpoints_osrm = {'Servidor público OSRM': 'http://router.project-osrm.org'}
        try:
            ep_secret = st.secrets.get('OSRM_ENDPOINT')
            if ep_secret:
                endpoints_osrm['Servidor corporativo configurado'] = str(ep_secret).rstrip('/')
        except Exception:
            pass
        endpoint_nome = st.selectbox('Endpoint OSRM autorizado:', list(endpoints_osrm.keys()), disabled=is_locked)
        url_osrm = endpoints_osrm[endpoint_nome]
        st.caption(url_osrm)
        usa_osrm = st.checkbox('🛣️ Traçado de Ruas Real', value=False, disabled=is_locked)
        tentativas_osrm = st.number_input('Tentativas por segmento OSRM:', min_value=1, max_value=3, value=2, step=1, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_san.empty:
        cfg_res = st.session_state.get('config_execucao_san', {})
        id_exec = cfg_res.get('id_execucao', 'SAN')
        st.caption(f'Execução: **{id_exec}** | Regras: **{VERSAO_REGRAS_SANEAMENTO}**')
        export_ok = bool((cfg_res.get('reconciliacao', {}) or {}).get('ok', False))
        if not export_ok:
            st.error('Exportações bloqueadas: a reconciliação de volumes não fechou.')

        if 'bytes_zip_xl_san' not in st.session_state:
            if st.button('⚙️ Gerar Planilhas Excel', use_container_width=True, disabled=not export_ok):
                with st.spinner('Gerando Excel...'):
                    st.session_state.bytes_zip_xl_san = gerar_zip_excel_atual()
                tentar_rerun()
        else:
            st.download_button('🌐 Baixar Planilhas (ZIP)', data=st.session_state.bytes_zip_xl_san, file_name=f'Saneamento_Planilhas_{id_exec}.zip', use_container_width=True, disabled=not export_ok)

        if 'bytes_zip_kml_san' not in st.session_state:
            if st.button('⚙️ Gerar Mapas KML', use_container_width=True, disabled=not export_ok):
                with st.spinner('Gerando KML...'):
                    st.session_state.bytes_zip_kml_san = gerar_zip_kml_atual()
                tentar_rerun()
        else:
            st.download_button('🗺️ Baixar Mapas (KML)', data=st.session_state.bytes_zip_kml_san, file_name=f'Saneamento_Mapas_{id_exec}.zip', use_container_width=True, disabled=not export_ok)

        if 'bytes_zip_gpx_san' not in st.session_state:
            if st.button('⚙️ Gerar GPS GPX', use_container_width=True, disabled=not export_ok):
                with st.spinner('Gerando GPX...'):
                    st.session_state.bytes_zip_gpx_san = gerar_zip_gpx_atual()
                tentar_rerun()
        else:
            st.download_button('🛰️ Baixar GPS (GPX)', data=st.session_state.bytes_zip_gpx_san, file_name=f'Saneamento_GPS_{id_exec}.zip', use_container_width=True, disabled=not export_ok)

        if st.button('🧹 Nova Roteirização', type='primary', use_container_width=True):
            limpar_roteirizador()


# ==============================================================
# RESULTADO
# ==============================================================
if is_done and not st.session_state.df_routed_san.empty:
    st.markdown('## 🎯 Resultado do Planejamento')
    cfg_res = st.session_state.get('config_execucao_san', {})
    st.caption(f"ID: **{cfg_res.get('id_execucao', '-')}** | Versão das regras: **{VERSAO_REGRAS_SANEAMENTO}**")

    if not cfg_res.get('status_col_localizada', True):
        st.warning('⚠️ A coluna de status/elegibilidade não foi localizada. Nenhuma obra foi excluída por status nesta execução.')

    rec = cfg_res.get('reconciliacao', {}) or {}
    if rec.get('ok'):
        st.success(f"✅ Reconciliação de volumes OK: {rec.get('total_entrada', 0)} registros de entrada foram integralmente conciliados.")
    else:
        st.error(f"🚨 ERRO DE RECONCILIAÇÃO: diferença de {rec.get('diferenca', '?')} registro(s). As exportações ficam bloqueadas até uma nova execução consistente.")

    df_c = st.session_state.get('df_correcao_san', pd.DataFrame())
    df_bases_c = st.session_state.get('df_bases_correcao_san', pd.DataFrame())
    if not df_c.empty:
        st.warning(f'⚠️ {len(df_c)} obras foram retidas por problemas de coordenadas.')
    if not df_bases_c.empty:
        st.info(f'🧭 {len(df_bases_c)} registros de bases/equipes tiveram correção, identidade separada ou rejeição geográfica auditada.')

    dfr = st.session_state.df_routed_san.copy()
    if 'DISTANCIA_HIBRIDA_KM' not in dfr.columns:
        rod = pd.to_numeric(dfr.get('DISTANCIA_RODOVIARIA_KM', pd.Series(index=dfr.index, dtype=float)), errors='coerce')
        est = pd.to_numeric(dfr.get('DISTANCIA_ESTIMADA_KM', pd.Series(index=dfr.index, dtype=float)), errors='coerce')
        dfr['DISTANCIA_HIBRIDA_KM'] = rod.where(rod.notna(), est)
    if 'DISTANCIA_PONTO_ANTERIOR_KM' in dfr.columns:
        dfr['DISTANCIA_PROXIMO_PONTO_KM'] = dfr.groupby(['BASE_ATRIBUIDA', 'DIA_MES'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
        st.session_state.df_routed_san['DISTANCIA_PROXIMO_PONTO_KM'] = dfr['DISTANCIA_PROXIMO_PONTO_KM']
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]

    te = dfr_t['BASE_ATRIBUIDA'].nunique()
    tr_real = sum(peso_tarefa(r) for _, r in dfr_t.iterrows())
    qs_total = int(dfr_t.get('SUPER_PONTO', pd.Series(index=dfr_t.index, dtype='object')).astype(str).str.startswith('SIM').sum())
    km_hibrido = pd.to_numeric(dfr['DISTANCIA_HIBRIDA_KM'], errors='coerce').fillna(0).sum()
    segmentos = len(dfr)
    osrm_ok = int((dfr.get('STATUS_ROTA', pd.Series(index=dfr.index, dtype='object')).astype(str) == 'OK_OSRM').sum())
    cobertura_osrm = (100.0 * osrm_ok / segmentos) if segmentos else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card('Tarefas Planejadas', tr_real, '🎯', '#0D256C', 'rgba(13,37,108,0.12)'), unsafe_allow_html=True)
    c2.markdown(render_metric_card('Equipes Alocadas', te, '👥', '#8b5cf6', 'rgba(139,92,246,0.15)'), unsafe_allow_html=True)
    c3.markdown(render_metric_card('Super Pontos', qs_total, '🏢', '#FF9800', 'rgba(255,152,0,0.15)'), unsafe_allow_html=True)
    c4.markdown(render_metric_card('KM Total Previsto', f'{km_hibrido:.1f} km', '🛣️', '#55B929', 'rgba(85,185,41,0.15)'), unsafe_allow_html=True)

    df_na = st.session_state.get('df_unallocated_san', pd.DataFrame())
    df_ft = st.session_state.get('df_fora_trava_san', pd.DataFrame())
    obras_na = sum(peso_tarefa(r) for _, r in df_na.iterrows()) if not df_na.empty else 0
    obras_ft = len(df_ft)
    dup_rem = len(st.session_state.get('df_duplicadas_removidas_san', pd.DataFrame()))
    dup_div = st.session_state.get('df_duplicadas_divergentes_san', pd.DataFrame())
    qtd_dup_div = int(dup_div['PROTOCOLO'].nunique()) if not dup_div.empty and 'PROTOCOLO' in dup_div.columns else 0
    seg_falhos = int((dfr.get('STATUS_ROTA', pd.Series(index=dfr.index, dtype='object')).astype(str) == 'SEM_ROTA_OSRM').sum())
    alert_jornada = int((dfr.get('ALERTA_JORNADA', pd.Series(index=dfr.index, dtype='object')).astype(str) == 'SIM').sum())

    q1, q2, q3, q4, q5 = st.columns(5)
    q1.metric('Não alocadas', obras_na)
    q2.metric('Fora da trava', obras_ft)
    q3.metric('Duplicadas divergentes', qtd_dup_div)
    q4.metric('Segmentos sem OSRM', seg_falhos)
    q5.metric('Cobertura OSRM', f'{cobertura_osrm:.1f}%')

    if alert_jornada:
        st.warning(f'⚠️ {alert_jornada} segmento(s) ainda apresentam alerta de jornada após as tentativas automáticas de replanejamento OSRM.')

    dash = st.session_state.get('df_dashboard_equipes_san', pd.DataFrame())
    if not dash.empty:
        with st.expander('📊 Utilização por equipe', expanded=True):
            st.dataframe(dash, use_container_width=True, hide_index=True)

    st.markdown('### 🗺️ Mapa Operacional')
    mostrar_mapa = st.checkbox('Exibir mapa operacional', value=False, key='mostrar_mapa_san')
    if mostrar_mapa:
        mapa = folium.Map(location=[dfr_t['LATITUDE'].mean(), dfr_t['LONGITUDE'].mean()], zoom_start=8) if not dfr_t.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ff9800', '#795548', '#607d8b']
        equipes_ordem = list(dict.fromkeys(dfr['BASE_ATRIBUIDA'].dropna().astype(str).tolist()))
        for idx_bn, bn in enumerate(equipes_ordem):
            if bn == 'NÃO ALOCADO':
                continue
            cr = co_f[idx_bn % len(co_f)]
            db = dfr[dfr['BASE_ATRIBUIDA'].astype(str) == bn]
            dias_ordem = list(dict.fromkeys(db.get('DIA_MES', pd.Series(index=db.index, dtype='object')).astype(str).tolist()))
            for dia in dias_ordem:
                dd = db[db.get('DIA_MES', pd.Series(index=db.index, dtype='object')).astype(str) == str(dia)]
                fg = folium.FeatureGroup(name=f'{bn} • {dia}', show=False)
                markers = MarkerCluster(name=f'Obras {bn} {dia}').add_to(fg)
                for _, r in dd.iterrows():
                    geom = r.get('ROTA_GEOMETRIA')
                    if isinstance(geom, list) and len(geom) >= 2:
                        pts = [[float(pt[1]), float(pt[0])] for pt in geom if isinstance(pt, (list, tuple)) and len(pt) >= 2]
                        if len(pts) >= 2:
                            folium.PolyLine(pts, color='black', weight=6, opacity=0.65).add_to(fg)
                            folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
                    if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
                        continue
                    ordem = int(r.get('ORDEM', 0) or 0)
                    cor_marker = '#777777' if r.get('STATUS_ROTA') == 'SEM_ROTA_OSRM' else ('#ff9800' if str(r.get('SUPER_PONTO', '')).startswith('SIM') else cr)
                    icon_html = f"<div style='background:{cor_marker};color:white;border:2px solid white;border-radius:50%;width:28px;height:28px;line-height:24px;text-align:center;font-weight:bold;box-shadow:0 1px 4px #333;'>{ordem}</div>"
                    er = ''.join([
                        f"<tr><td><b>{html.escape(str(c))}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>"
                        for c in st.session_state.get('colunas_exibir_san', [])
                        if c.upper() not in ['NOME_DIA', 'DIA_MES', 'SEMANA', 'BASE_ATRIBUIDA']
                    ])
                    extras = f"<tr><td><b>Status rota</b></td><td>{html.escape(str(r.get('STATUS_ROTA','-')))}</td></tr><tr><td><b>Hora</b></td><td>{r.get('HORA_INICIO','-')} - {r.get('HORA_FIM','-')}</td></tr>"
                    pop_html = f'<div style="width:290px;"><b>Equipe:</b> {html.escape(str(bn))}<br><b>Dia:</b> {html.escape(str(dia))}<br><b>Ordem:</b> {ordem}<br><table border="1" style="width:100%;font-size:11px;">{extras}{er}</table></div>'
                    folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.DivIcon(html=icon_html), popup=folium.Popup(pop_html, max_width=340)).add_to(markers)
                fg.add_to(mapa)
        folium.LayerControl(collapsed=False).add_to(mapa)
        st_folium(mapa, use_container_width=True, height=580)
    else:
        st.caption('O mapa é carregado apenas quando solicitado para manter a tela rápida.')

    tabs = st.tabs(['📊 Dados', '📉 Não Alocadas', '🧱 Fora da Trava', '🔁 Duplicidades', '🚫 Fora do Filtro', '❓ Sem Nota', '🧭 Auditoria de Bases'])
    with tabs[0]:
        st.dataframe(dfr.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE'], errors='ignore'), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(df_na, use_container_width=True, hide_index=True) if not df_na.empty else st.success('✅ Nenhuma obra ficou sem alocação operacional.')
    with tabs[2]:
        st.dataframe(df_ft, use_container_width=True, hide_index=True) if not df_ft.empty else st.success('✅ Nenhuma obra ficou fora da trava global.')
    with tabs[3]:
        dfd = st.session_state.get('df_duplicadas_san', pd.DataFrame())
        st.dataframe(dfd, use_container_width=True, hide_index=True) if not dfd.empty else st.success('✅ Nenhuma NOTA duplicada detectada.')
    with tabs[4]:
        dff = st.session_state.get('df_fora_filtro_san', pd.DataFrame())
        st.dataframe(dff, use_container_width=True, hide_index=True) if not dff.empty else st.success('✅ Nenhum registro ficou fora do filtro.')
    with tabs[5]:
        dfsn = st.session_state.get('df_sem_nota_san', pd.DataFrame())
        st.dataframe(dfsn, use_container_width=True, hide_index=True) if not dfsn.empty else st.success('✅ Nenhum registro sem NOTA/PROTOCOLO.')
    with tabs[6]:
        st.dataframe(df_bases_c, use_container_width=True, hide_index=True) if not df_bases_c.empty else st.success('✅ Nenhuma inconsistência registrada nas bases das equipes.')


# ==============================================================
# ENTRADA / PRÉ-PROCESSAMENTO
# ==============================================================
else:
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

                    motivos_base = []
                    inval = []
                    for _, rr in df_bases.iterrows():
                        okc, mot = validar_par_coordenadas(rr.get('LATITUDE'), rr.get('LONGITUDE'))
                        inval.append(not okc)
                        motivos_base.append(mot)
                    inval = pd.Series(inval, index=df_bases.index)
                    df_bases['STATUS_COORD_BASE'] = np.where(inval, 'PENDENTE_CORRECAO', 'OK')
                    df_bases['MOTIVO_COORD_BASE'] = motivos_base
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
                            ok_fb, mot_fb = validar_par_coordenadas(la, lo)
                            if ok_fb:
                                df_bases.at[idx, 'LATITUDE'] = float(la)
                                df_bases.at[idx, 'LONGITUDE'] = float(lo)
                                df_bases.at[idx, 'BASE_COORD_FONTE'] = 'IBGE_FALLBACK'
                                df_bases.at[idx, 'STATUS_COORD_BASE'] = 'CORRIGIDA'
                                rr = df_bases.loc[idx].copy()
                                rr['MOTIVO_AUDITORIA_BASE'] = f"{df_bases.at[idx, 'MOTIVO_COORD_BASE']}; corrigida por município"
                                audit_bases.append(rr)
                            else:
                                df_bases.at[idx, 'STATUS_COORD_BASE'] = 'REJEITADA'
                                rr = df_bases.loc[idx].copy()
                                rr['MOTIVO_AUDITORIA_BASE'] = f"{df_bases.at[idx, 'MOTIVO_COORD_BASE']}; fallback indisponível ({mot_fb})"
                                audit_bases.append(rr)
                    elif inval.any():
                        for idx in df_bases.index[inval]:
                            df_bases.at[idx, 'STATUS_COORD_BASE'] = 'REJEITADA'
                            rr = df_bases.loc[idx].copy()
                            rr['MOTIVO_AUDITORIA_BASE'] = f"{df_bases.at[idx, 'MOTIVO_COORD_BASE']}; sem município/residência para fallback"
                            audit_bases.append(rr)

                    df_bases = df_bases[df_bases['STATUS_COORD_BASE'] != 'REJEITADA'].copy()
                    if 'MUNICIPIO' in df_bases.columns:
                        df_bases['MUN_LIMPO_BASE'] = normalizar_municipios(df_bases['MUNICIPIO'].astype(str).fillna(''))
                    elif 'RESIDENCIA' in df_bases.columns:
                        df_bases['MUN_LIMPO_BASE'] = normalizar_municipios(df_bases['RESIDENCIA'].astype(str).fillna(''))
                    else:
                        df_bases['MUN_LIMPO_BASE'] = ''

                    if not df_bases.empty:
                        df_bases, audit_ident = preparar_ids_equipes(df_bases)
                        if not audit_ident.empty:
                            audit_bases.extend([r for _, r in audit_ident.iterrows()])
                    st.session_state.df_bases_correcao_san = pd.DataFrame(audit_bases)
                    overrides = [c for c in ['COTA_DIA', 'HORA_INICIO', 'HORA_FIM', 'DIAS_TRABALHO'] if c in df_bases.columns]
                    if overrides:
                        st.caption('⚙️ Configurações por equipe detectadas: ' + ', '.join(overrides))
            else:
                st.error('❌ A planilha não possui coluna de nome da Equipe/Levantador.')

        st.markdown('##### 📍 Regra de Atribuição')
        ta_index = 1 if is_continuo else 0
        ta = st.radio('Como amarrar as notas aos técnicos?', ['Por Proximidade Espacial', 'Por Município da Planilha'], index=ta_index, disabled=is_continuo, label_visibility='collapsed')
        if is_continuo:
            st.caption('🔒 No modo Contínuo, as obras são vinculadas aos municípios definidos para as equipes.')
        if 'Município' in ta:
            modo_ancora = 'Base fixa'
        else:
            modo_ancora = st.radio('Referência para proximidade:', ['Base fixa', 'Âncora dinâmica', 'Balanceada'], index=2, disabled=is_locked)
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

    qtd_eq = df_bases['EQUIPE_ID'].nunique() if 'EQUIPE_ID' in df_bases.columns else df_bases['BASE_NOME'].nunique()
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
    col_id = next((cc for cc in ['NOTA', 'PROTOCOLO', 'OS', 'ID'] if cc in df_tasks.columns), None)
    if col_id is None:
        st.error('🚨 Nenhuma coluna de NOTA/PROTOCOLO/OS/ID foi localizada nas demandas.')
        st.stop()
    df_tasks['PROTOCOLO'] = df_tasks[col_id].astype(str).str.split(r'\s*\|\s*')
    df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True)
    df_tasks['_ORDEM_ENTRADA'] = np.arange(len(df_tasks))
    qtd_entrada_total = len(df_tasks)
    df_tasks['PROTOCOLO'] = limpar_protocolo_serie(df_tasks['PROTOCOLO'])
    df_sem_nota = df_tasks[df_tasks['PROTOCOLO'] == ''].copy()
    if not df_sem_nota.empty:
        df_sem_nota['MOTIVO_EXCLUSAO'] = 'SEM_NOTA_PROTOCOLO'
    st.session_state.df_sem_nota_san = df_sem_nota
    df_tasks = df_tasks[df_tasks['PROTOCOLO'] != ''].copy()

    if 'MUNICIPIO' not in df_tasks.columns:
        st.error('🚨 Falta a coluna MUNICIPIO.')
        st.stop()
    possui_coord = (('LATITUDE PROJETO' in df_tasks.columns and 'LONGITUDE PROJETO' in df_tasks.columns) or ('LATITUDE' in df_tasks.columns and 'LONGITUDE' in df_tasks.columns))
    if not possui_coord:
        st.error('🚨 Não foi encontrado um par de coordenadas: LATITUDE/LONGITUDE ou LATITUDE PROJETO/LONGITUDE PROJETO.')
        st.stop()

    c1_f, c2_f = st.columns([1, 1])
    df_fora_filtro = pd.DataFrame(columns=df_tasks.columns)
    cs = next((c for c in ['STATUS CLIENTE', 'STATUS DA FISCALIZACAO', 'STATUS DA FISCALIZAÇÃO'] if c in df_tasks.columns), None)
    status_col_localizada = cs is not None
    with c1_f:
        if cs:
            df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
            opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
            default_s = [s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']] if 'FISCAL' in cs else opts_s
            sel_s = st.multiselect('1. Status Roteirizáveis:', options=opts_s, default=default_s)
            if not sel_s:
                st.stop()
            mask_status = df_tasks[cs].isin(sel_s)
            df_fora_filtro = df_tasks[~mask_status].copy()
            if not df_fora_filtro.empty:
                df_fora_filtro['MOTIVO_EXCLUSAO'] = 'STATUS_NAO_SELECIONADO'
            df_tasks = df_tasks[mask_status].copy()
        else:
            st.warning('⚠️ Coluna de status/elegibilidade não localizada. O sistema não aplicará filtro de status.')
    st.session_state.df_fora_filtro_san = df_fora_filtro

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

    # 1) Validação geográfica vem antes de duplicidade e trava, evitando perder uma ocorrência válida.
    st.markdown('#### 🌍 Auditoria Geográfica e de Duplicidades')
    df_tasks, df_rej = validar_coordenadas_obras(df_tasks)
    st.session_state.df_correcao_san = df_rej
    if df_tasks.empty:
        st.error('🚨 Nenhuma obra válida restou após a validação geográfica.')
        st.stop()

    # 2) Classifica duplicidades em exatas/divergentes; só exatas podem ser removidas automaticamente.
    df_tasks = auditar_duplicidades(df_tasks, tolerancia_exata_m=30.0)
    df_dup_audit = df_tasks[df_tasks['DUPLICADA_GERAL'] == 'SIM'].copy()
    df_dup_div = df_tasks[df_tasks['DUPLICATA_DIVERGENTE'] == 'SIM'].copy()
    st.session_state.df_duplicadas_san = df_dup_audit
    st.session_state.df_duplicadas_divergentes_san = df_dup_div
    df_tasks, df_dup_rem = deduplicar_apenas_exatas(df_tasks, habilitado=deduplicar_notas)
    st.session_state.df_duplicadas_removidas_san = df_dup_rem
    if not df_dup_div.empty:
        st.warning(f"⚠️ {df_dup_div['PROTOCOLO'].nunique()} NOTA(s) possuem ocorrências divergentes. Elas foram preservadas para auditoria e roteamento; não são removidas automaticamente.")

    # 3) Trava global após validação e deduplicação exata.
    df_tasks, df_fora_trava = aplicar_trava_global(df_tasks, trava_global)
    st.session_state.df_fora_trava_san = df_fora_trava

    df_tasks['MUN_LIMPO'] = normalizar_municipios(df_tasks['MUNICIPIO'].astype(str).fillna(''))

    # 4) Super Pontos nunca atravessam município quando a atribuição é municipal.
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
    df_tasks['_TASK_UID'] = [f'TASK-{i:07d}' for i in range(len(df_tasks))]

    tempos_demanda = parse_tempos_demanda_texto(tempos_demanda_txt)
    config_pre = {
        'versao_regras': VERSAO_REGRAS_SANEAMENTO,
        'velocidade_media_kmh': 30.0,
        'obras_por_dia': int(obras_dia),
        'tempo_medio_obra_min': float(tempo_medio_obra),
        'tempos_por_tipo_demanda': tempos_demanda,
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
        'modo_ancora': modo_ancora,
        'distancia_max_atribuicao_km': float(dist_max_atribuicao),
        'raio_super_ponto_m': int(raio_sp),
        'deduplicar_notas': bool(deduplicar_notas),
        'trava_global': int(trava_global),
    }

    # 5) Atribuição considera capacidade restante e o modo de âncora escolhido.
    df_ta, df_u = atribuir_tarefas_equipes(df_tasks, df_bases, ta, modo_ancora, dist_max_atribuicao, config_pre)
    st.session_state.df_unallocated_san = df_u
    total_alocadas = sum(peso_tarefa(r) for _, r in df_ta.iterrows()) if not df_ta.empty else 0
    sb_html.markdown(render_sidebar_card(cm, total_alocadas, qtd_eq, cm * qtd_eq, is_continuo), unsafe_allow_html=True)
    if df_ta.empty:
        st.error('Nenhuma obra pôde ser alocada às equipes. Verifique municípios, coordenadas, capacidade e limite de distância.')
        st.stop()

    with st.expander('🛠️ Configuração de Saída (Colunas)', expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = ['NOTA', 'STATUS CLIENTE', 'NOME', 'TIPO DEMANDA', 'MUNICIPIO', 'ENDERECO', 'BAIRRO', 'PONTO REFERENCIA', 'COMPLEMENTO', 'LATITUDE PROJETO', 'LONGITUDE PROJETO', 'FONTE_COORDENADA', 'CLASSIFICACAO AREA', 'TEL FIXO', 'TEL MOVEL', 'GRUPO TENSAO', 'ARQUIVO_ORIGEM']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect('Colunas que vão aparecer no Mapa e Excel:', tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button('🚀 Iniciar Motor de Roteirização', type='primary', use_container_width=True):
        id_exec = criar_id_execucao()
        tbr = df_bases.to_dict('records')
        b_names = list(dict.fromkeys(df_ta['BASE_ATRIBUIDA'].astype(str).tolist()))
        config_exec = dict(config_pre)
        config_exec.update({
            'id_execucao': id_exec,
            'arquivo_equipes': bf.name if bf else '-',
            'arquivos_demanda': [f.name for f in task_files],
            'qtd_entrada_total': int(qtd_entrada_total),
            'qtd_sem_nota': int(len(df_sem_nota)),
            'qtd_fora_filtro': int(len(df_fora_filtro)),
            'qtd_entrada_filtrada': int(len(df_tasks)),
            'qtd_duplicidades': int(df_dup_audit['PROTOCOLO'].nunique()) if not df_dup_audit.empty else 0,
            'qtd_linhas_duplicadas': int(len(df_dup_audit)),
            'qtd_duplicidades_divergentes': int(df_dup_div['PROTOCOLO'].nunique()) if not df_dup_div.empty else 0,
            'qtd_duplicidades_removidas': int(len(df_dup_rem)),
            'qtd_correcao_obras': int(len(df_rej)),
            'qtd_correcao_bases': int(len(st.session_state.get('df_bases_correcao_san', pd.DataFrame()))),
            'qtd_fora_trava': int(len(df_fora_trava)),
            'status_col_localizada': bool(status_col_localizada),
        })
        st.session_state.config_execucao_san = config_exec
        st.session_state.update({'bases_records_san': tbr, 'colunas_exibir_san': colunas_exibir})
        st.session_state.vrp_state_san = {
            'config': config_exec,
            'b_names': b_names,
            'b_idx': 0,
            'unvisited': df_ta.copy(),
            'routed_data': [],
            'osrm_cache': {},
            'replan_osrm_iter': 0,
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
                br_rows = br_all[br_all['BASE_NOME'].astype(str) == str(bn)]
                if br_rows.empty or pd.isna(br_rows.iloc[0].get('LATITUDE')):
                    st_v['b_idx'] += 1
                    st.session_state.vrp_state_san = st_v
                    tentar_rerun()
                    st.stop()
                br = br_rows.iloc[0].to_dict()
                bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
                cfg_eq = aplicar_config_equipe(cfg, br)
                oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'].astype(str) == str(bn)].to_dict('records')
                rf, nao_agendadas = construir_plano_rota_equipe(oe, bl, bL, cfg_eq)
                if nao_agendadas:
                    dfn = pd.DataFrame(nao_agendadas)
                    st.session_state.df_unallocated_san = pd.concat([
                        st.session_state.get('df_unallocated_san', pd.DataFrame()), dfn
                    ], ignore_index=True)
                    uids_nao = set(dfn.get('_TASK_UID', pd.Series(dtype='object')).dropna().astype(str))
                    oe = [x for x in oe if str(x.get('_TASK_UID')) not in uids_nao]
                st_v['team_tasks'] = oe
                st_v['cfg_equipe_atual'] = cfg_eq
                st_v['base_atual'] = (bl, bL)
                st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []
                st_v['replan_osrm_iter'] = 0
                st.session_state.vrp_state_san = st_v
                tentar_rerun()
                st.stop()

            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            cfg_eq = st_v.get('cfg_equipe_atual', cfg)
            ei = min(oi + (20 if cfg_eq['tracado_real'] else max(1, len(rf))), len(rf))
            cache = st_v.setdefault('osrm_cache', {})
            for i in range(oi, ei):
                it = rf[i]
                if not cfg_eq['tracado_real']:
                    geom = [[it['La'], it['la']], [it['Lt'], it['lt']]]
                    gd.append({'geom': geom, 'duracao_s': float(it['tempo_estimado_min']) * 60.0, 'dist_rod_km': np.nan, 'status': 'ESTIMADA_SEM_OSRM'})
                    continue

                if i % 5 == 0:
                    sgt.info(f'🛣️ Traçando arruamento **{bn}**... ({i+1}/{len(rf)})')
                render_t(b_i, i, len(rf))
                chave = (round(float(it['la']), 5), round(float(it['La']), 5), round(float(it['lt']), 5), round(float(it['Lt']), 5), str(cfg_eq['url_osrm_base']))
                if chave in cache:
                    gd.append(cache[chave])
                    continue

                resultado = None
                for tentativa in range(int(cfg_eq.get('tentativas_osrm', 2))):
                    try:
                        geom, dur_s = obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg_eq['url_osrm_base'], cfg_eq['velocidade_media_kmh'])
                        if isinstance(geom, list) and len(geom) >= 2:
                            dist_rod = distancia_geometria_km(geom)
                            resultado = {'geom': geom, 'duracao_s': float(dur_s), 'dist_rod_km': dist_rod, 'status': 'OK_OSRM'}
                            cache[chave] = resultado
                            break
                    except Exception:
                        if tentativa + 1 < int(cfg_eq.get('tentativas_osrm', 2)):
                            time.sleep(0.15)
                if resultado is None:
                    resultado = {'geom': [], 'duracao_s': float(it['tempo_estimado_min']) * 60.0, 'dist_rod_km': np.nan, 'status': 'SEM_ROTA_OSRM'}
                gd.append(resultado)

            st_v['c_idx'], st_v['current_geoms'], st_v['osrm_cache'] = ei, gd, cache
            if ei < len(rf):
                st.session_state.vrp_state_san = st_v
                tentar_rerun()
                st.stop()

            # Segunda passagem: tempos OSRM reais podem empurrar a última obra do dia para o dia seguinte.
            if cfg_eq.get('tracado_real', False):
                ajustes = detectar_reagendamento_osrm(rf, gd, cfg_eq)
                if ajustes and int(st_v.get('replan_osrm_iter', 0)) < 5:
                    tarefas = [dict(x) for x in st_v.get('team_tasks', [])]
                    mapa_aj = {str(a['task_uid']): int(a['novo_dia_min']) for a in ajustes}
                    for o in tarefas:
                        uid = str(o.get('_TASK_UID'))
                        if uid in mapa_aj:
                            o['_DIA_MINIMO'] = max(int(o.get('_DIA_MINIMO', 1) or 1), mapa_aj[uid])
                            o['REAGENDADA_OSRM'] = 'SIM'
                    bl, bL = st_v.get('base_atual', (0.0, 0.0))
                    novo_rf, novas_nao = construir_plano_rota_equipe(tarefas, bl, bL, cfg_eq)
                    if novas_nao:
                        dfn = pd.DataFrame(novas_nao)
                        antigos = st.session_state.get('df_unallocated_san', pd.DataFrame())
                        if '_TASK_UID' in dfn.columns and not antigos.empty and '_TASK_UID' in antigos.columns:
                            dfn = dfn[~dfn['_TASK_UID'].astype(str).isin(antigos['_TASK_UID'].astype(str))]
                        if not dfn.empty:
                            st.session_state.df_unallocated_san = pd.concat([antigos, dfn], ignore_index=True)
                        uids_nao = set(dfn.get('_TASK_UID', pd.Series(dtype='object')).dropna().astype(str))
                        tarefas = [x for x in tarefas if str(x.get('_TASK_UID')) not in uids_nao]
                    st_v['team_tasks'] = tarefas
                    st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = novo_rf, 0, []
                    st_v['replan_osrm_iter'] = int(st_v.get('replan_osrm_iter', 0)) + 1
                    st.session_state.vrp_state_san = st_v
                    sgt.warning(f"♻️ Ajustando jornada com tempos reais OSRM (passagem {st_v['replan_osrm_iter']}/5)...")
                    tentar_rerun()
                    st.stop()

            # Converte plano + geometrias em saída final usando tempos reais quando disponíveis.
            rdf = []
            ordem_por_dia = {}
            relogio_por_dia = {}
            fim_jornada_por_dia = {}
            for it, info_geo in zip(rf, gd):
                chave_dia = (it['d'], it['dm'])
                if chave_dia not in relogio_por_dia:
                    data_dt = datetime.strptime(it['dm'], '%d/%m/%Y').date()
                    relogio_por_dia[chave_dia] = datetime.combine(data_dt, cfg_eq['hora_inicio'])
                    fim_jornada_por_dia[chave_dia] = datetime.combine(data_dt, cfg_eq['hora_fim'])
                    ordem_por_dia[chave_dia] = 1

                tempo_rota_min = float(info_geo['duracao_s']) / 60.0
                inicio_desloc = relogio_por_dia[chave_dia]
                chegada = inicio_desloc + pd.Timedelta(minutes=tempo_rota_min)
                status_rota = info_geo['status']
                dist_rod = info_geo['dist_rod_km']
                dist_est = float(it['distancia_estimada_km'])
                dist_hib = float(dist_rod) if pd.notna(dist_rod) else dist_est

                if it.get('ir', False):
                    fim = chegada
                    alerta_j = 'SIM' if fim > fim_jornada_por_dia[chave_dia] else 'NÃO'
                    rdf.append({
                        'PROTOCOLO': 'RETORNO_BASE', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'],
                        'BASE_ATRIBUIDA': bn, 'ORDEM': ordem_por_dia[chave_dia], 'NOME_DIA': it['dn'],
                        'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': it['periodo'],
                        'DISTANCIA_PONTO_ANTERIOR_KM': round(dist_hib, 2),
                        'DISTANCIA_ESTIMADA_KM': round(dist_est, 3),
                        'DISTANCIA_RODOVIARIA_KM': round(float(dist_rod), 3) if pd.notna(dist_rod) else np.nan,
                        'DISTANCIA_HIBRIDA_KM': round(dist_hib, 3),
                        'TEMPO_ROTA_MIN': round(tempo_rota_min, 2), 'TEMPO_ATENDIMENTO_MIN': 0.0, 'STATUS_ROTA': status_rota,
                        'ROTA_GEOMETRIA': info_geo['geom'], 'PRIORIDADE': 'Não',
                        'HORA_INICIO': inicio_desloc.strftime('%H:%M'), 'HORA_FIM': fim.strftime('%H:%M'),
                        '_HORA_INICIO_DT': inicio_desloc, '_HORA_FIM_DT': fim, 'ALERTA_JORNADA': alerta_j,
                    })
                    relogio_por_dia[chave_dia] = fim
                    ordem_por_dia[chave_dia] += 1
                    continue

                ob = dict(it['o'])
                inicio_serv = aplicar_intervalo_almoco(chegada, it['servico_min'], cfg_eq)
                fim_serv = inicio_serv + pd.Timedelta(minutes=float(it['servico_min']))
                alerta_j = 'SIM' if fim_serv > fim_jornada_por_dia[chave_dia] else 'NÃO'
                ob.update({
                    'ORDEM': ordem_por_dia[chave_dia], 'NOME_DIA': it['dn'], 'DIA_MES': it['dm'],
                    'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': it['periodo'],
                    'DISTANCIA_PONTO_ANTERIOR_KM': round(dist_hib, 2),
                    'DISTANCIA_ESTIMADA_KM': round(dist_est, 3),
                    'DISTANCIA_RODOVIARIA_KM': round(float(dist_rod), 3) if pd.notna(dist_rod) else np.nan,
                    'DISTANCIA_HIBRIDA_KM': round(dist_hib, 3),
                    'TEMPO_ROTA_MIN': round(tempo_rota_min, 2), 'TEMPO_ATENDIMENTO_MIN': round(float(it['servico_min']), 2), 'STATUS_ROTA': status_rota,
                    'ROTA_GEOMETRIA': info_geo['geom'],
                    'HORA_INICIO': inicio_serv.strftime('%H:%M'), 'HORA_FIM': fim_serv.strftime('%H:%M'),
                    '_HORA_INICIO_DT': inicio_serv, '_HORA_FIM_DT': fim_serv, 'ALERTA_JORNADA': alerta_j,
                })
                rdf.append(ob)
                relogio_por_dia[chave_dia] = fim_serv
                ordem_por_dia[chave_dia] += 1

            st_v['routed_data'].extend(rdf)
            for chave in ['c_rotas', 'c_idx', 'current_geoms', 'team_tasks', 'cfg_equipe_atual', 'base_atual']:
                st_v.pop(chave, None)
            st_v['b_idx'] += 1
            st_v['replan_osrm_iter'] = 0
            st.session_state.vrp_state_san = st_v
            gc.collect()
            tentar_rerun()
        else:
            sgt.success('✅ Rotas Finalizadas!')
            pb.progress(1.0)
            df_final = pd.DataFrame(st_v['routed_data'])
            st.session_state.df_routed_san = df_final
            cfg_final = st.session_state.get('config_execucao_san', {}).copy()
            df_na = st.session_state.get('df_unallocated_san', pd.DataFrame())
            roteadas = sum(peso_tarefa(r) for _, r in df_final[~df_final['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].iterrows()) if not df_final.empty else 0
            nao_alocadas = sum(peso_tarefa(r) for _, r in df_na.iterrows()) if not df_na.empty else 0
            rec = calcular_reconciliacao(
                cfg_final.get('qtd_entrada_total', 0),
                roteadas,
                len(st.session_state.get('df_sem_nota_san', pd.DataFrame())),
                len(st.session_state.get('df_fora_filtro_san', pd.DataFrame())),
                len(st.session_state.get('df_correcao_san', pd.DataFrame())),
                len(st.session_state.get('df_duplicadas_removidas_san', pd.DataFrame())),
                len(st.session_state.get('df_fora_trava_san', pd.DataFrame())),
                nao_alocadas,
            )
            cfg_final['qtd_nao_alocadas'] = int(nao_alocadas)
            cfg_final['tempo_processamento_s'] = round(time.time() - st_run, 2)
            cfg_final['reconciliacao'] = rec
            st.session_state.config_execucao_san = cfg_final
            bases_df = pd.DataFrame(st.session_state.get('bases_records_san', []))
            st.session_state.df_dashboard_equipes_san = montar_dashboard_equipes(df_final, bases_df, cfg_final)
            st.session_state.roteamento_concluido_san = True
            st.session_state.vrp_status_san = 'IDLE'
            st.session_state.pop('start_time_run_san', None)
            # Libera estruturas intermediárias pesadas; resultados e cache exportável ficam fora do motor.
            st.session_state.vrp_state_san = {'config': cfg_final}
            gc.collect()
            tentar_rerun()

    except Exception:
        erro_id = f"SANERR-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
        print(f'[{erro_id}] Falha no Roteirizador Saneamento:\n{traceback.format_exc()}')
        st.error(f'🚨 Não foi possível concluir esta etapa. Código para suporte: {erro_id}')
        st.session_state.vrp_status_san = 'IDLE'
