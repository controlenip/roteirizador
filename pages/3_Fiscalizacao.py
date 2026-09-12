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
import altair as alt
import plotly.express as px
import unicodedata
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formatar_moeda, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas

# IMPORTAÇÃO DA EXPORTAÇÃO
from modules.export_fisc import injetar_logo, gerar_excel_fisc, gerar_excel_resumo_fisc, gerar_gpx_simples, gerar_kml_fisc, identificar_icone_folium, limpar_colunas_fisc, gerar_txt_fisc

st.set_page_config(page_title="Fiscalização", page_icon="📋", layout="wide")

# ==========================================
# CSS AGRESSIVO PARA CONSERTAR O MULTISELECT
# ==========================================
st.markdown("""
<style>
    /* Expande a caixa do multiselect verticalmente */
    .stMultiSelect [data-baseweb="select"] {
        min-height: 42px !important;
    }
    /* Força os itens a quebrarem de linha e remove a trava de altura */
    .stMultiSelect [data-baseweb="select"] > div:first-child {
        display: flex !important;
        flex-wrap: wrap !important;
        height: auto !important;
        max-height: none !important;
        overflow-y: visible !important;
        padding-bottom: 5px !important;
    }
    /* Remove o gradiente transparente que corta os nomes à direita */
    .stMultiSelect [data-baseweb="select"] > div:first-child > div:last-child {
        display: none !important;
    }
    /* Permite que os chips azuis usem todo o espaço necessário */
    .stMultiSelect [data-baseweb="tag"] {
        max-width: 100% !important;
        margin-bottom: 4px !important;
        margin-top: 4px !important;
    }
    .stMultiSelect [data-baseweb="tag"] span {
        white-space: normal !important;
        text-overflow: clip !important;
    }
</style>
""", unsafe_allow_html=True)

injetar_logo()

# ==========================================
# FUNÇÕES DE CORREÇÃO E APOIO
# ==========================================
def corrigir_coord(val, limite):
    if pd.isna(val): return np.nan
    v = float(val)
    iters = 0
    while abs(v) > limite and iters < 10:
        v /= 10.0
        iters += 1
    return v

def remover_acentos(text):
    if not isinstance(text, str): return text
    return "".join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))

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

def render_sidebar_card(limite_por_equipe, total_obras_prontas, qtd_equipes_ativas, total_capacidade):
    return f"""
    <div style="background-color: #ffffff; padding: 25px; border-radius: 10px; border: 1px solid #e0e0e0; box-shadow: 0 4px 8px rgba(0,0,0,0.05); margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 18px; border-bottom: 2px solid #55B929; padding-bottom: 10px; margin-bottom: 15px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Fiscais Ativos:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 10px; font-size: 15px;"><b>Cota p/ Fiscal:</b> <span style="color: #d9534f; font-weight: bold;">{limite_por_equipe}</span></p>
        <p style="margin-bottom: 15px; font-size: 15px;"><b>Capacidade Total:</b> <span style="color: #55B929; font-weight: bold;">{total_capacidade}</span></p>
        <hr style="margin: 15px 0; border: 0; border-top: 1px solid #eee;">
        <div style="text-align: center; margin-top: 15px;">
            <p style="margin-bottom: 5px; font-size: 16px; color: #555;"><b>Obras Validadas:</b></p>
            <span style="font-size: 36px; color: #0D256C; font-weight: 900;">{total_obras_prontas}</span>
        </div>
    </div>
    """

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        cu = c.upper()
        if cu in ['DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'DISTANCIA_RODOVIARIA_KM']:
            return f"{vf:.2f} KM"
        if 'DISTANCIA' in cu:
            return f"{vf:.2f} Metros"
        if 'TEMPO_ROTA_MIN' in cu:
            return f"{vf:.1f} min"
        if 'POSTE' in cu: return f"{int(round(vf))}"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def count_real_obras(row):
    if isinstance(row.get('_ORIGINAL_ROWS'), list): return len(row['_ORIGINAL_ROWS'])
    val = str(row.get('SUPER_PONTO', ''))
    if val.startswith('SIM'):
        nums = re.findall(r'\d+', val)
        if nums: return int(nums[0])
    return 1

def extrair_qtd(val):
    if pd.isna(val) or val == '': return 0.0
    if isinstance(val, (int, float)): return float(val)
    try:
        nums = re.findall(r'\d+\.?\d*', str(val).replace(',', '.'))
        return sum(float(n) for n in nums) if nums else 0.0
    except: return 0.0

def definir_cor_fiscalizacao(qtd):
    try:
        q = float(qtd)
        return 'green' if q <= 100 else 'blue' if q <= 200 else 'beige' if q <= 300 else 'orange' if q <= 400 else 'red'
    except: return 'gray'


def distancia_geometria_km(geom):
    """Calcula a distância percorrida ao longo da geometria sem alterar a lógica de alocação."""
    if not isinstance(geom, list) or len(geom) < 2:
        return np.nan
    total = 0.0
    pontos = []
    for pt in geom:
        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
            lon, lat = pt[0], pt[1]
            if pd.notna(lat) and pd.notna(lon):
                pontos.append((float(lat), float(lon)))
    if len(pontos) < 2:
        return np.nan
    for i in range(1, len(pontos)):
        total += haversine_scalar(pontos[i-1][0], pontos[i-1][1], pontos[i][0], pontos[i][1])
    return total


def chave_osrm_fisc(lat1, lon1, lat2, lon2, url, velocidade):
    """Chave estável do cache, incluindo endpoint e parâmetros relevantes."""
    return (
        str(url).strip(), round(float(velocidade), 2),
        round(float(lat1), 5), round(float(lon1), 5),
        round(float(lat2), 5), round(float(lon2), 5)
    )


def obter_rota_cache_fisc(lat1, lon1, lat2, lon2, url, velocidade):
    cache = st.session_state.setdefault('osrm_cache_fisc', {})
    chave = chave_osrm_fisc(lat1, lon1, lat2, lon2, url, velocidade)
    if chave in cache:
        return cache[chave], True

    rota = obter_rota_ruas(lat1, lon1, lat2, lon2, url, velocidade)
    if rota and len(rota) > 0 and isinstance(rota[0], list) and len(rota[0]) >= 2:
        cache[chave] = rota
        # Evita crescimento indefinido em sessões longas.
        while len(cache) > 10000:
            cache.pop(next(iter(cache)))
    return rota, False


def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_fisc': False, 'vrp_status_fisc': "IDLE", 'vrp_state_fisc': {}, 'df_routed_fisc': pd.DataFrame(), 'bases_records_fisc': [], 'colunas_exibir_fisc': [], 'colunas_originais_fisc': []})
    for k in ['bytes_zip_xl_fisc', 'bytes_zip_kml_fisc', 'bytes_zip_gpx_fisc', 'bytes_zip_txt_fisc', 'start_time_run_fisc', 'start_time_pkg_fisc', 'df_unallocated_fisc', 'df_correcao_fiscalizacao', 'diagnostico_fisc']: st.session_state.pop(k, None)
    # Mantém cache de leitura e cache OSRM para acelerar uma nova execução com os mesmos arquivos/trechos.
    tentar_rerun()

# ==========================================
# CONTROLE DE SESSÃO
# ==========================================
if "roteamento_concluido_fisc" not in st.session_state: st.session_state.roteamento_concluido_fisc = False
if "vrp_status_fisc" not in st.session_state: st.session_state.vrp_status_fisc = "IDLE"
if "osrm_cache_fisc" not in st.session_state: st.session_state.osrm_cache_fisc = {}

status_exec = st.session_state.vrp_status_fisc
is_done = st.session_state.roteamento_concluido_fisc
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📋 Planejamento de Fiscalização</h1>", unsafe_allow_html=True)
st.info("💡 **A Regra do Bolsão:** A IA ancora os Fiscais nas obras com MAIS POSTES primeiro. Em seguida, varre as obras menores e finaliza a rota exatamente no maior foco do mapa.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Esforço e Limites", expanded=True):
        
        is_modo_tatico = st.session_state.get('modo_operacao_fisc', 'Tático') == "Tático"
        
        trava_global = st.number_input("Trava Total de Obras no Estado", min_value=0, value=0, step=50, disabled=is_locked)
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (Metros)", 10, 1000, 100, 10, disabled=is_locked)
        st.markdown("---")
        
        if is_modo_tatico:
            obras_dia = st.number_input("Cota Diária por Equipe:", min_value=1, value=6, disabled=is_locked)
            tpc = st.radio("Visão de Trabalho:", ["Dia", "Semana"], index=1, disabled=is_locked)
            limite_per = st.number_input(f"Qtd de {tpc}s de Rota:", min_value=1, value=1, disabled=is_locked)
            dias_sel = st.multiselect("Dias Úteis na Semana:", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"], default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], disabled=is_locked)
        else:
            st.success("📦 **Modo Contínuo Ativado:** Limites e Jornadas desativados. Roteirização infinita.")
            obras_dia = 999999
            tpc = "Dia"
            limite_per = 1
            dias_sel = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            
        data_ini = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        vel_kmh = 30.0
    
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)

    st.markdown("---")
    sb_html = st.empty()
    
    if is_done and not st.session_state.df_routed_fisc.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_fisc', b"vazio"), file_name=f"Fiscais_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("📝 Baixar Relatórios (TXT)", data=st.session_state.get('bytes_zip_txt_fisc', b"vazio"), file_name=f"Fiscais_TXT - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_fisc', b"vazio"), file_name=f"Fiscais_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_fisc', b"vazio"), file_name=f"Fiscais_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

# ==========================================
# EXIBIÇÃO DE RESULTADOS (SE CONCLUÍDO)
# ==========================================
if is_done and not st.session_state.df_routed_fisc.empty:
    st.markdown("## 🎯 Dashboards de Produtividade")
    
    df_c = st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica Oficial:</b> As obras listadas no arquivo <b>"Obras_Correcao"</b> foram bloqueadas porque apresentaram coordenadas geográficas em branco, zeradas ou invertidas.
            </p>
        </div>
        """, unsafe_allow_html=True)

    # Trabalha em cópia para que a camada visual não altere o dataframe oficial da sessão.
    dfr = st.session_state.df_routed_fisc.copy()
    dfr['DISTANCIA_PROXIMO_PONTO_KM'] = dfr.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()

    if not dfr_t.empty:
        dfr_t['_QTD_OBRAS_REAL'] = [count_real_obras(r) for _, r in dfr_t.iterrows()]
    else:
        dfr_t['_QTD_OBRAS_REAL'] = pd.Series(dtype='int64')

    tr = int(dfr_t['_QTD_OBRAS_REAL'].sum()) if '_QTD_OBRAS_REAL' in dfr_t.columns else 0
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    tot_postes_global = int(pd.to_numeric(dfr_t['QTD PREVISTA DE POSTES'], errors='coerce').sum()) if 'QTD PREVISTA DE POSTES' in dfr_t else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("TOTAL DE OBRAS", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Fiscais Alocados", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Postes Auditados", tot_postes_global, "🏗️", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Previsto Total", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    # Diagnóstico de arruamento: apenas informativo, sem interferir nas rotas.
    if 'STATUS_ROTA' in dfr.columns:
        status_validos = dfr[dfr['PROTOCOLO'] != 'PAUSA_ALMOCO']['STATUS_ROTA'].fillna('SEM STATUS').astype(str)
        qtd_sem_rota = int(status_validos.str.contains('SEM_ROTA', na=False).sum())
        qtd_osrm = int(status_validos.str.startswith('OSRM', na=False).sum())
        qtd_estimado = int(status_validos.str.contains('ESTIMADA', na=False).sum())
        if qtd_sem_rota > 0:
            st.warning(f"⚠️ Arruamento: {qtd_osrm} trechos OSRM | {qtd_sem_rota} sem rota | {qtd_estimado} estimados.")
        else:
            st.caption(f"🛣️ Arruamento: {qtd_osrm} trechos OSRM | {qtd_estimado} estimados | nenhuma falha pendente.")

    st.markdown("<br>", unsafe_allow_html=True)

    # Agrega uma única vez para gráfico e resumo, evitando filtros repetidos por fiscal.
    if not dfr_t.empty:
        tmp_chart = dfr_t.copy()
        tmp_chart['_POSTES_NUM'] = pd.to_numeric(tmp_chart.get('QTD PREVISTA DE POSTES', 0), errors='coerce').fillna(0)
        df_chart = tmp_chart.groupby('BASE_ATRIBUIDA', as_index=False).agg(
            Obras=('_QTD_OBRAS_REAL', 'sum'),
            Postes=('_POSTES_NUM', 'sum')
        ).rename(columns={'BASE_ATRIBUIDA': 'Fiscal'})
        df_chart['Obras'] = df_chart['Obras'].astype(int)
        df_chart['Postes'] = df_chart['Postes'].round().astype(int)
    else:
        df_chart = pd.DataFrame(columns=['Fiscal', 'Obras', 'Postes'])

    if not df_chart.empty:
        total_postes_chart = df_chart['Postes'].sum()
        df_chart['Perc_Postes'] = (df_chart['Postes'] / total_postes_chart * 100).fillna(0) if total_postes_chart else 0.0
        df_chart['Perc_Text'] = df_chart['Perc_Postes'].apply(lambda x: f"{x:.1f}%")
        df_chart = df_chart.sort_values(by="Fiscal").reset_index(drop=True)

        c_ch1, c_ch2 = st.columns([1.2, 1])
        with c_ch1:
            st.markdown("#### Balanceamento: Obras vs Postes")
            df_melted = df_chart.melt(id_vars='Fiscal', value_vars=['Obras', 'Postes'], var_name='Métrica', value_name='Quantidade')

            bars = alt.Chart(df_melted).mark_bar().encode(
                x=alt.X('Fiscal:N', title=None, axis=alt.Axis(labelAngle=-45)),
                y=alt.Y('Quantidade:Q', title="Total Alocado"),
                color=alt.Color('Métrica:N', scale=alt.Scale(domain=['Obras', 'Postes'], range=['#0D256C', '#FF9800'])),
                xOffset='Métrica:N',
                tooltip=['Fiscal', 'Métrica', 'Quantidade']
            )
            text_bars = alt.Chart(df_melted).mark_text(align='center', baseline='bottom', dy=-5, fontSize=11).encode(
                x=alt.X('Fiscal:N'), y=alt.Y('Quantidade:Q'), xOffset='Métrica:N', text='Quantidade:Q'
            )
            st.altair_chart((bars + text_bars).properties(height=350), use_container_width=True)

        with c_ch2:
            st.markdown("#### % Fatia de Postes por Fiscal")
            fig_pie = px.pie(
                df_chart,
                values='Postes',
                names='Fiscal',
                hole=0.55,
                custom_data=['Obras']
            )
            fig_pie.update_traces(
                textposition='outside',
                textinfo='percent',
                hovertemplate="<b>%{label}</b><br>Postes: %{value}<br>Obras: %{customdata[0]}<br>Fatia: %{percent}<extra></extra>",
                marker=dict(line=dict(color='#ffffff', width=2))
            )
            fig_pie.update_layout(
                margin=dict(t=20, b=20, l=20, r=20),
                showlegend=True,
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0)
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    # O mapa passa a ser carregado sob demanda para não atrasar dashboards/tabelas.
    st.markdown("### 🗺️ Mapa Geográfico")
    mostrar_mapa = st.checkbox("Exibir mapa operacional", value=False, key="mostrar_mapa_fisc")
    if mostrar_mapa:
        mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
        co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']

        m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
        bases_unicas = dfr['BASE_ATRIBUIDA'].dropna().unique().tolist()
        for bn in bases_unicas:
            cr = co_f[bases_unicas.index(bn) % len(co_f)]
            db = dfr[dfr['BASE_ATRIBUIDA'] == bn]

            bn_safe = str(bn).replace("{", "[").replace("}", "]")
            fg = folium.FeatureGroup(name=f"Rota: {bn_safe}", show=False)

            for pe in db['PERIODO'].unique():
                dp = db[db['PERIODO'] == pe]

                for _, r_seg in dp.iterrows():
                    geom = r_seg.get('ROTA_GEOMETRIA')
                    if not isinstance(geom, list) or len(geom) < 2:
                        continue

                    pts_seg = []
                    for pt in geom:
                        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                            lon_pt, lat_pt = pt[0], pt[1]
                            if pd.notna(lat_pt) and pd.notna(lon_pt):
                                pts_seg.append([lat_pt, lon_pt])

                    if len(pts_seg) >= 2:
                        folium.PolyLine(pts_seg, color='black', weight=7, opacity=0.9).add_to(fg)
                        folium.PolyLine(pts_seg, color=cr, weight=3, opacity=1.0).add_to(fg)

                for r in dp.to_dict('records'):
                    if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
                        continue

                    c_i = r.get('COR_ICONE', 'gray')
                    try:
                        qtd_p = int(float(r.get('QTD PREVISTA DE POSTES', 0)))
                    except Exception:
                        qtd_p = 0
                    ic = identificar_icone_folium(r, dfr.columns)

                    bg_colors = {'green': '#4CAF50', 'blue': '#2196F3', 'beige': '#FFC107', 'orange': '#FF9800', 'red': '#F44336', 'gray': '#9E9E9E'}
                    txt_colors = {'beige': '#000000', 'orange': '#000000', 'green': '#ffffff', 'blue': '#ffffff', 'red': '#ffffff', 'gray': '#ffffff'}

                    if str(r.get('SUPER_PONTO', '')).startswith('SIM'):
                        p_bg, p_c = '#FFD700', '#000000'
                        p_txt = f"🏢 SUPER PONTO {str(r.get('SUPER_PONTO')).replace('SIM','').strip()}"
                    else:
                        p_bg, p_c = bg_colors.get(c_i, '#9E9E9E'), txt_colors.get(c_i, '#ffffff')
                        p_txt = f"📋 FISCALIZAÇÃO - {qtd_p} POSTES"

                    def formatar_popup(coluna):
                        if coluna == 'PROTOCOLO': return 'NOTA'
                        if coluna == 'BASE_ATRIBUIDA': return 'FISCAL'
                        return coluna

                    er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(formatar_popup(c))}</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_fisc if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA','COR_ICONE']])
                    pop_html = f'<div style="width:280px;"><div style="background:{p_bg};color:{p_c};padding:8px;font-weight:bold;">{p_txt}</div><table border="1" style="width:100%;font-size:12px;"><tr><td style="padding:3px;"><b>Ordem:</b></td><td style="padding:3px;">{r.get("ORDEM",0)}</td></tr>{er}</table></div>'
                    pop_html = pop_html.replace("{", "&#123;").replace("}", "&#125;")

                    folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
            fg.add_to(mapa)
        folium.LayerControl().add_to(mapa)
        st_folium(mapa, use_container_width=True, height=550)
    else:
        st.caption("O mapa é carregado somente quando solicitado para reduzir o tempo de renderização.")

    t1, t2 = st.tabs(["📊 Dados Tabulares", "📉 Resumo por Técnico"])
    with t1:
        st.data_editor(st.session_state.df_routed_fisc.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'COR_ICONE'], errors='ignore'), use_container_width=True)
    with t2:
        resumo_dict = dict(zip(df_chart['Fiscal'], df_chart['Obras'])) if not df_chart.empty else {}
        dr = pd.DataFrame([{"Fiscal": b['LEVANTADOR'], "Obras Roteirizadas": int(resumo_dict.get(b['LEVANTADOR'], 0))} for b in st.session_state.bases_records_fisc]).reset_index(drop=True)
        st.dataframe(dr, use_container_width=True)

# ==========================================
# START DA APLICAÇÃO (UPLOAD DE DADOS)
# ==========================================
elif status_exec == "IDLE":
    modo_operacao = st.radio("🎯 Selecione o Modo de Operação:", ["📍 Planejamento Tático (Cruzar Bases e Obras)", "📜 Lista Contínua (Fiscal da Própria Planilha)"], horizontal=True)
    st.session_state.modo_operacao_fisc = "Tático" if "Tático" in modo_operacao else "Contínua"
    st.markdown("---")
    
    is_tatico = st.session_state.modo_operacao_fisc == "Tático"
    
    df_bases = pd.DataFrame()
    task_files = None
    
    if is_tatico:
        c_up1, c_up2 = st.columns(2)
        with c_up1:
            st.markdown("### 👥 1. Fiscais")
            bf = st.file_uploader("Suba a planilha de Fiscais (Excel)", type=["xlsx", "xls"])
            if bf:
                b_t = ler_planilha_cached(bf.getvalue()); b_t.columns = normalize_cols(b_t.columns)
                b_t = b_t.loc[:, ~b_t.columns.duplicated()].copy()
                for pn in ['NOME', 'FISCAL', 'TECNICO', 'COLABORADOR', 'LEVANTADOR']:
                    if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'LEVANTADOR'}); break
                
                if 'LEVANTADOR' in b_t.columns:
                    b_t['LEVANTADOR'] = b_t['LEVANTADOR'].astype(str).str.split(r'\s*\|\s*')
                    b_t = b_t.explode('LEVANTADOR').reset_index(drop=True)
                    b_t['LEVANTADOR'] = b_t['LEVANTADOR'].str.strip().str.upper()
                    
                    opts = sorted([str(x) for x in b_t['LEVANTADOR'].dropna().unique() if str(x) not in ['SEM LEVANTADOR', 'NAN', 'NONE', '']])
                    sel = st.multiselect("Selecione os Fiscais Ativos:", opts, default=opts)
                    if sel:
                        df_bases = b_t[b_t['LEVANTADOR'].isin(sel)].copy()
                        
                        cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                        if cr in df_bases.columns:
                            df_bases['MUN_LIMPO_BASE'] = normalizar_municipios(df_bases[cr])

                        if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                            df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                            df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                            df_bases['LATITUDE'] = df_bases['LATITUDE'].apply(lambda x: corrigir_coord(x, 90))
                            df_bases['LONGITUDE'] = df_bases['LONGITUDE'].apply(lambda x: corrigir_coord(x, 180))
                            
                        elif cr in df_bases.columns:
                            mc = {}
                            with st.spinner("🌍 Mapeando bases..."):
                                for m in df_bases[cr].dropna().unique(): mc[m] = obter_coordenadas_municipio_cached(m)
                            df_bases['LATITUDE'], df_bases['LONGITUDE'] = df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0]), df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
                        df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE']); df_bases['TIPO_EQUIPE'] = 'FISCAL'
                else: st.error("❌ A planilha não possui a coluna 'FISCAL'.")

            st.markdown("##### 📍 Regra de Atribuição")
            ta = st.radio("Atribuição", ["Por Proximidade (Recomendado)", "Por Município Rígido"], index=0, label_visibility="collapsed")
            if "Proximidade" in ta: st.caption("A IA persegue os maiores Bolsões e puxa o Fiscal mais próximo.")
            else: st.caption("Trava o Fiscal rigorosamente à cidade informada na sua planilha.")

        with c_up2:
            st.markdown("### 📁 2. Obras de Fiscalização")
            task_files = st.file_uploader("Suba as Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
            
        if df_bases.empty or not task_files: st.stop()
        
    else:
        st.markdown("### 📁 Demandas (Obras de Fiscalização)")
        st.info("💡 Suba a planilha de Fiscalização. O sistema identificará a coluna **'FISCAL'** automaticamente e criará a lista contínua por técnico, varrendo todo o roteiro sem limites.")
        task_files = st.file_uploader("Suba as Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
        if not task_files: st.stop()

    dfs = []
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        if not dfs: st.session_state.colunas_originais_fisc = dft.columns.tolist()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    st.markdown("---")
    falta = [c for c in ['MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas: {', '.join(falta)}."); st.stop()
    
    cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO'
    if cs in df_tasks.columns:
        df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
        st.markdown("#### 📊 Filtragem de Status")
        opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
        sel_s = st.multiselect("Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
        if not sel_s: st.stop()
        df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()

    cg1, cg2 = st.columns([4, 1])
    with cg1: st.markdown("#### 🌍 Limpeza de Coordenadas Geográficas")
    with cg2:
        if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    pbg = st.progress(0.0); tmp = st.empty(); sgt = st.empty(); df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    m_m = df_tasks['MUNICIPIO'].isna() | (df_tasks['MUNICIPIO'].astype(str).str.strip() == '') | (df_tasks['MUNICIPIO'].astype(str).str.strip().str.upper() == 'NAN')
    if m_m.sum() > 0:
        df_tasks.loc[m_m, 'MOTIVO_REJEICAO'] = 'Município Vazio'
        df_rej = pd.concat([df_rej, df_tasks[m_m].copy()], ignore_index=True); df_tasks = df_tasks[~m_m].copy()

    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    
    df_tasks['LAT_NUM'] = df_tasks['LAT_NUM'].apply(lambda x: corrigir_coord(x, 90))
    df_tasks['LON_NUM'] = df_tasks['LON_NUM'].apply(lambda x: corrigir_coord(x, 180))

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
    
    pbg.empty(); tmp.empty(); sgt.empty()
    st.session_state.df_correcao_fiscalizacao = df_rej
    
    if not df_rej.empty: 
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_rej)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica Oficial:</b> As obras listadas no arquivo <b>"Obras_Correcao"</b> foram bloqueadas porque apresentaram <b>coordenadas geográficas em branco, zeradas ou invertidas</b>.
            </p>
        </div>
        """, unsafe_allow_html=True)

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    # ================================================================
    # PRÉ-VALIDAÇÃO LEVE (não altera nem exclui registros automaticamente)
    # ================================================================
    diag_alertas = []
    if 'PROTOCOLO' in df_tasks.columns:
        qtd_dup_prot = int(df_tasks['PROTOCOLO'].astype(str).duplicated(keep=False).sum())
        if qtd_dup_prot > 0:
            diag_alertas.append(f"{qtd_dup_prot} linhas com protocolo repetido")
    if all(c in df_tasks.columns for c in ['PROTOCOLO', 'LATITUDE', 'LONGITUDE']):
        qtd_dup_exato = int(df_tasks.duplicated(subset=['PROTOCOLO', 'LATITUDE', 'LONGITUDE'], keep=False).sum())
        if qtd_dup_exato > 0:
            diag_alertas.append(f"{qtd_dup_exato} linhas repetidas por protocolo + coordenada")
    if all(c in df_tasks.columns for c in ['LATITUDE', 'LONGITUDE']):
        qtd_coord_rep = int(df_tasks.duplicated(subset=['LATITUDE', 'LONGITUDE'], keep=False).sum())
        if qtd_coord_rep > 0:
            diag_alertas.append(f"{qtd_coord_rep} linhas compartilham coordenadas")
    if is_tatico and not df_bases.empty and 'LEVANTADOR' in df_bases.columns:
        qtd_fisc_dup = int(df_bases['LEVANTADOR'].duplicated(keep=False).sum())
        if qtd_fisc_dup > 0:
            diag_alertas.append(f"{qtd_fisc_dup} linhas de fiscais possuem nome repetido")

    if diag_alertas:
        with st.expander("🔎 Diagnóstico preventivo dos dados", expanded=False):
            st.info("Os itens abaixo são apenas alertas. Nenhum registro foi removido automaticamente.")
            for msg in diag_alertas:
                st.write(f"• {msg}")

    # A configuração visual passa a usar uma prévia das colunas. A alocação, a
    # consolidação de Super Pontos e a ordenação pesada só ocorrem após o clique.
    df_preview = df_tasks.copy()
    agent_col_preview = None
    if is_tatico:
        if 'BASE_ATRIBUIDA' not in df_preview.columns:
            df_preview['BASE_ATRIBUIDA'] = ''
        qtd_eq_preview = df_bases['LEVANTADOR'].nunique()
    else:
        for col in ['FISCAL', 'LEVANTADOR', 'NOME_FISCAL', 'EQUIPE', 'COLABORADOR']:
            if col in df_preview.columns:
                agent_col_preview = col
                break
        if not agent_col_preview:
            st.error("❌ A planilha de Obras não possui uma coluna de responsável (FISCAL).")
            st.stop()
        if 'BASE_ATRIBUIDA' not in df_preview.columns:
            df_preview['BASE_ATRIBUIDA'] = df_preview[agent_col_preview]
        qtd_eq_preview = df_preview['BASE_ATRIBUIDA'].astype(str).str.strip().replace(['NAN', 'NONE', ''], np.nan).dropna().nunique()

    cm_preview = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per if is_tatico else "Ilimitada"
    tot_cap_preview = cm_preview * qtd_eq_preview if is_tatico else "Ilimitada"
    sb_html.markdown(render_sidebar_card(cm_preview, len(df_tasks), qtd_eq_preview, tot_cap_preview), unsafe_allow_html=True)

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_preview.columns if not c.startswith('_') and c not in ['COR_ICONE', 'MUN_LIMPO']]

        cd_padrao = [
            'PROTOCOLO',
            'VALOR DA OBRA',
            'QTD PREVISTA DE POSTES',
            'PREVISAO DE ENTREGA',
            'PARCEIRO',
            'TIPO DE FISCALIZACAO',
            'TIPO DE PROJETO',
            'REGIONAL',
            'MUNICIPIO',
            'LATITUDE',
            'LONGITUDE',
            'ZONA',
            'BASE_ATRIBUIDA'
        ]

        tc_limpo = {remover_acentos(c).upper().replace("QTD DE PREVISTA", "QTD PREVISTA"): c for c in tc}
        cp_default = []
        for ideal in cd_padrao:
            if ideal in tc_limpo:
                cp_default.append(tc_limpo[ideal])

        def formatar_display(col):
            if col in ['PROTOCOLO', 'NOTA']:
                return 'NOTA'
            if col in ['BASE_ATRIBUIDA', 'FISCAL']:
                return 'FISCAL'
            return col

        colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp_default, format_func=formatar_display)

        def get_sort_index(col):
            c_upper = remover_acentos(col).upper().replace("QTD DE PREVISTA", "QTD PREVISTA")
            if c_upper in cd_padrao:
                return cd_padrao.index(c_upper)
            if c_upper == 'NOTA':
                return cd_padrao.index('PROTOCOLO')
            if c_upper == 'FISCAL':
                return cd_padrao.index('BASE_ATRIBUIDA')
            return 999

        colunas_exibir.sort(key=get_sort_index)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        prep_pb = st.progress(0.0)
        prep_msg = st.empty()
        prep_msg.info("⚙️ Preparando dados para o motor...")

        df_ta = pd.DataFrame()
        df_u = pd.DataFrame()
        tbr = []

        # Normalização em lote: mesma regra anterior, sem repetir uma chamada por linha.
        df_work = df_tasks.copy()
        df_work['MUN_LIMPO'] = normalizar_municipios(df_work['MUNICIPIO'].astype(str))
        prep_pb.progress(0.15)

        if is_tatico:
            tbr = df_bases.to_dict('records')
            fiscal_anchors = {
                b['LEVANTADOR']: (float(b.get('LATITUDE', 0)), float(b.get('LONGITUDE', 0)))
                for b in tbr
            }
            assigned_tasks = []
            unassigned_tasks = []

            df_work = df_work.sort_values(by=['LATITUDE', 'LONGITUDE'], ascending=[True, True])
            registros = df_work.to_dict('records')
            total_reg = max(1, len(registros))

            prep_msg.info("📍 Distribuindo obras entre os fiscais...")
            for idx, r in enumerate(registros):
                la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
                ms = r.get('MUN_LIMPO', '')

                if "Município" in ta:
                    vb = [b for b in tbr if str(b.get('MUN_LIMPO_BASE', '')) == str(ms)]
                else:
                    vb = tbr

                best_f, best_d = None, float('inf')
                if pd.notna(la) and pd.notna(lo) and vb:
                    for b in vb:
                        f_name = b['LEVANTADOR']
                        d = haversine_scalar(la, lo, fiscal_anchors[f_name][0], fiscal_anchors[f_name][1])
                        if d < best_d:
                            best_d = d
                            best_f = f_name

                if best_f:
                    r['BASE_ATRIBUIDA'] = best_f
                    assigned_tasks.append(r)
                    fiscal_anchors[best_f] = (la, lo)
                else:
                    r['MOTIVO_REJEICAO'] = "Fora de Área (Sem Fiscal)"
                    r['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
                    unassigned_tasks.append(r)

                if idx % max(1, total_reg // 20) == 0:
                    prep_pb.progress(min(0.50, 0.15 + 0.35 * ((idx + 1) / total_reg)))

            df_ta = pd.DataFrame(assigned_tasks)
            df_u = pd.DataFrame(unassigned_tasks)
            qtd_eq = df_bases['LEVANTADOR'].nunique()
        else:
            agent_col = agent_col_preview
            df_work = df_work.rename(columns={agent_col: 'BASE_ATRIBUIDA'}) if agent_col != 'BASE_ATRIBUIDA' else df_work
            df_work['BASE_ATRIBUIDA'] = df_work['BASE_ATRIBUIDA'].astype(str).str.strip().str.upper()
            df_work = df_work[~df_work['BASE_ATRIBUIDA'].isin(['NAN', 'NONE', ''])]
            df_ta = df_work
            df_u = pd.DataFrame()
            qtd_eq = df_ta['BASE_ATRIBUIDA'].nunique()
            tbr = [{'LEVANTADOR': b} for b in df_ta['BASE_ATRIBUIDA'].dropna().unique()]
            prep_pb.progress(0.50)

        if df_ta.empty:
            prep_pb.empty(); prep_msg.empty()
            st.error("Nenhuma obra pôde ser alocada aos Fiscais.")
            st.stop()

        if 'QTD PREVISTA DE POSTES' in df_ta.columns:
            df_ta['QTD PREVISTA DE POSTES'] = df_ta['QTD PREVISTA DE POSTES'].apply(extrair_qtd)
            df_ta['COR_ICONE'] = df_ta['QTD PREVISTA DE POSTES'].apply(definir_cor_fiscalizacao)
        else:
            df_ta['QTD PREVISTA DE POSTES'], df_ta['COR_ICONE'] = 0.0, 'gray'

        prep_msg.info("🧩 Consolidando Super Pontos...")
        dfs_fundidos = []
        bases_lista = list(df_ta['BASE_ATRIBUIDA'].dropna().unique())
        total_bases = max(1, len(bases_lista))
        for ib, base in enumerate(bases_lista):
            df_base = df_ta[df_ta['BASE_ATRIBUIDA'] == base].copy()
            df_base_f, _ = fundir_super_pontos(df_base, raio_metros=raio_sp, agrupar_por_levantador=False)
            dfs_fundidos.append(df_base_f)
            prep_pb.progress(min(0.90, 0.50 + 0.40 * ((ib + 1) / total_bases)))

        df_ta = pd.concat(dfs_fundidos, ignore_index=True) if dfs_fundidos else pd.DataFrame()
        if df_ta.empty:
            prep_pb.empty(); prep_msg.empty()
            st.error("Nenhuma obra pôde ser consolidada para roteirização.")
            st.stop()

        df_ta = df_ta.sort_values(by=['QTD PREVISTA DE POSTES', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
        if trava_global > 0:
            df_ta = df_ta.head(trava_global)

        st.session_state.df_unallocated_fisc = df_u
        st.session_state.tot_obras_nao_alocadas = sum(
            len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
            for _, r in df_u.iterrows()
        )

        cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per if is_tatico else "Ilimitada"
        tot_cap = cm * qtd_eq if is_tatico else "Ilimitada"
        total_validadas = sum(
            len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
            for _, r in df_ta.iterrows()
        )

        st.session_state.diagnostico_fisc = {
            'linhas_validas_preparacao': int(len(df_tasks)),
            'obras_validadas': int(total_validadas),
            'nao_alocadas': int(st.session_state.tot_obras_nao_alocadas),
            'alertas_duplicidade': diag_alertas.copy(),
        }

        st.session_state.update({'bases_records_fisc': tbr, 'colunas_exibir_fisc': colunas_exibir})

        cfg_obras_dia = obras_dia if is_tatico else 999999
        cfg_limite_per = limite_per if is_tatico else 1
        cfg_tpc = tpc if is_tatico else "Dia"

        st.session_state.vrp_state_fisc = {
            'config': {
                'velocidade_media_kmh': vel_kmh,
                'obras_por_dia': cfg_obras_dia,
                'tipo_periodo': cfg_tpc,
                'limite_periodos': cfg_limite_per,
                'dias_selecionados': dias_sel,
                'url_osrm_base': url_osrm,
                'tracado_real': usa_osrm,
                'data_inicio': data_ini,
                'sentido_rota': sentido_rota,
                'is_tatico': is_tatico,
                'modo_operacao': st.session_state.modo_operacao_fisc,
                'regra_atribuicao': ta if is_tatico else 'Fiscal da própria planilha',
                'raio_super_ponto_m': raio_sp,
                'trava_global': trava_global,
                'capacidade_por_fiscal': cm,
                'capacidade_total': tot_cap
            },
            # Remove duplicados sem perder a ordem original dos fiscais.
            'b_names': list(dict.fromkeys([b['LEVANTADOR'] for b in tbr])),
            'b_idx': 0,
            'unvisited': df_ta.copy(),
            'routed_data': [],
            'current_geoms': []
        }
        prep_pb.progress(1.0)
        prep_msg.success("✅ Preparação concluída. Iniciando motor de roteirização...")
        st.session_state.vrp_status_fisc = "RUNNING"
        time.sleep(0.3)
        tentar_rerun()

# ==========================================
# CÁLCULO VRP (LÓGICA: SMALL -> BIG)
# ==========================================
if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor VRP (Fiscalização de Bolsões)")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_fisc', time.time())
    if 'start_time_run_fisc' not in st.session_state: st.session_state.start_time_run_fisc = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_fisc; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center; box-shadow:0 2px 5px rgba(0,0,0,0.05);"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center; box-shadow:0 2px 5px rgba(0,0,0,0.05);"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Intercalando Obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            if not oe:
                st_v['b_idx'] += 1; st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
                
            if cfg['is_tatico']:
                br = pd.DataFrame(st.session_state.bases_records_fisc)
                br = br[br['LEVANTADOR'] == bn].iloc[0]
                if pd.isna(br.get('LATITUDE')): st_v['b_idx'] += 1; st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
                bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
            else:
                bl = sum(float(x['LATITUDE']) for x in oe) / len(oe)
                bL = sum(float(x['LONGITUDE']) for x in oe) / len(oe)
            
            ot = []
            if oe:
                max_idx = max(range(len(oe)), key=lambda i: extrair_qtd(oe[i].get('QTD PREVISTA DE POSTES', 0)))
                p_max = oe.pop(max_idx)
                
                cl, cL = bl, bL
                while oe:
                    closest_idx = min(range(len(oe)), key=lambda i: haversine_scalar(cl, cL, float(oe[i]['LATITUDE']), float(oe[i]['LONGITUDE'])))
                    nx = oe.pop(closest_idx)
                    ot.append(nx)
                    cl, cL = float(nx['LATITUDE']), float(nx['LONGITUDE'])
                ot.append(p_max)
            
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
                if cfg['is_tatico']:
                    if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                        st.session_state.df_unallocated_fisc = pd.concat([st.session_state.get('df_unallocated_fisc', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue

                qr = count_real_obras(o)
                vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                vm = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60
                
                cp = es['t'] + pd.Timedelta(minutes=vm)
                
                if cfg['is_tatico'] and cp.hour >= 12 and not es['lu']:
                    ls = max(es['t'], es['d'].replace(hour=12)); le = ls + pd.Timedelta(hours=1)
                    rf.append({'o': None, 'il': True, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': es['l'], 'Lt': es['L'], 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': ls, 'hf': le, 'vm': 0.0, 'dk': 0.0})
                    es['t'], es['lu'] = le, True; cp = es['t'] + pd.Timedelta(minutes=vm)
                    
                fp = cp + pd.Timedelta(minutes=60)
                
                if cfg['is_tatico'] and es['oh'] > 0 and (es['oh'] + qr > cfg['obras_por_dia']):
                    dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                    rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
                    da += 1
                    if cfg['tipo_periodo'] == "Semana":
                        dds += 1
                        if dds > len(cfg['dias_selecionados']): sa += 1; dds = 1
                    es = gi(da)
                    if (cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos']):
                        st.session_state.df_unallocated_fisc = pd.concat([st.session_state.get('df_unallocated_fisc', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue
                    vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE']); vk = vkr * 1.3
                    vm = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60
                    cp = es['t'] + pd.Timedelta(minutes=vm); fp = cp + pd.Timedelta(minutes=60)
                
                rf.append({'o': o, 'il': False, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': cp, 'hf': fp, 'vm': vm, 'dk': vk})
                es['l'], es['L'], es['t'], es['oh'] = o['LATITUDE'], o['LONGITUDE'], fp, es['oh'] + qr
                
            if cfg['is_tatico'] and es['oh'] > 0 and not ((cfg['tipo_periodo'] == "Semana" and sa > cfg['limite_periodos']) or (cfg['tipo_periodo'] == "Dia" and da > cfg['limite_periodos'])):
                dr = haversine_vectorized(es['l'], es['L'], bl, bL); vr = (dr/cfg['velocidade_media_kmh'])*60
                rf.append({'o': None, 'il': False, 'ir': True, 'la': es['l'], 'La': es['L'], 'lt': bl, 'Lt': bL, 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': es['t'], 'hf': es['t']+pd.Timedelta(minutes=vr), 'vm': vr, 'dk': dr})
            
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']

            # Com arruamento real ativo, processa menos trechos por ciclo para
            # reduzir a pressão sobre o servidor público do OSRM.
            ei = min(oi + (10 if cfg['tracado_real'] else len(rf)), len(rf))

            for i in range(oi, ei):
                it = rf[i]

                if not cfg['tracado_real']:
                    geom = [[it['La'], it['la']], [it['Lt'], it['lt']]]
                    dur_s = (it['dk'] / cfg['velocidade_media_kmh']) * 3600 if cfg['velocidade_media_kmh'] else 0.0
                    gd.append({
                        'geom': geom,
                        'dur_s': dur_s,
                        'status': 'LINHA_RETA_ESTIMADA',
                        'dist_km': float(it['dk']),
                        'cache': False
                    })
                    continue

                if i % 5 == 0:
                    sgt.info(f"🛣️ Traçando arruamento real **{bn}**... ({i}/{len(rf)})")
                render_t(b_i, i, len(rf))

                sucesso_rota = False
                rota_valida = None
                veio_cache = False

                # Primeiro consulta o cache persistente da sessão. Se não existir,
                # mantém as mesmas 5 tentativas com backoff já usadas pelo motor.
                try:
                    chave = chave_osrm_fisc(
                        it['la'], it['La'], it['lt'], it['Lt'],
                        cfg['url_osrm_base'], cfg['velocidade_media_kmh']
                    )
                    cache = st.session_state.setdefault('osrm_cache_fisc', {})
                    if chave in cache:
                        rota_valida = cache[chave]
                        sucesso_rota = True
                        veio_cache = True
                except Exception:
                    pass

                if not sucesso_rota:
                    for tentativa in range(5):
                        try:
                            time.sleep(0.8)
                            rota, _ = obter_rota_cache_fisc(
                                it['la'], it['La'],
                                it['lt'], it['Lt'],
                                cfg['url_osrm_base'],
                                cfg['velocidade_media_kmh']
                            )
                            if rota and len(rota) > 0:
                                geom_rota = rota[0]
                                if isinstance(geom_rota, list) and len(geom_rota) >= 2:
                                    rota_valida = rota
                                    sucesso_rota = True
                                    break
                        except Exception:
                            pass

                        time.sleep(1.5 + tentativa * 0.5)

                if sucesso_rota and rota_valida is not None:
                    geom_rota = rota_valida[0] if len(rota_valida) > 0 else []
                    dur_s = float(rota_valida[1]) if len(rota_valida) > 1 and pd.notna(rota_valida[1]) else 0.0
                    dist_real = distancia_geometria_km(geom_rota)
                    gd.append({
                        'geom': geom_rota,
                        'dur_s': dur_s,
                        'status': 'OSRM',
                        'dist_km': dist_real,
                        'cache': veio_cache
                    })
                else:
                    # Mantém correspondência 1:1 entre rf e gd sem desenhar reta fictícia.
                    gd.append({
                        'geom': [],
                        'dur_s': 0.0,
                        'status': 'SEM_ROTA_OSRM',
                        'dist_km': np.nan,
                        'cache': False
                    })

            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
            
            rdf, og, dp = [], 1, ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            for it, info_rota in zip(rf, gd):
                g = info_rota.get('geom', []) if isinstance(info_rota, dict) else info_rota[0]
                ds = info_rota.get('dur_s', 0.0) if isinstance(info_rota, dict) else (info_rota[1] if len(info_rota) > 1 else 0.0)
                status_rota = info_rota.get('status', 'OSRM' if cfg['tracado_real'] else 'LINHA_RETA_ESTIMADA') if isinstance(info_rota, dict) else ('OSRM' if cfg['tracado_real'] else 'LINHA_RETA_ESTIMADA')
                dist_real = info_rota.get('dist_km', np.nan) if isinstance(info_rota, dict) else distancia_geometria_km(g)
                tempo_rota_min = round(float(ds) / 60.0, 2) if pd.notna(ds) else np.nan
                dist_real_round = round(float(dist_real), 2) if pd.notna(dist_real) else np.nan

                pv = it['s'] if cfg['tipo_periodo']=="Semana" else it['d']
                dn = dp[datetime.strptime(it['dm'], '%d/%m/%Y').weekday()] if cfg['tipo_periodo']=="Semana" else f"Dia {it['d']}"
                campos_rota = {
                    'STATUS_ROTA': status_rota,
                    'DISTANCIA_RODOVIARIA_KM': dist_real_round,
                    'TEMPO_ROTA_MIN': tempo_rota_min
                }

                if it['il']:
                    linha = {'PROTOCOLO': 'PAUSA_ALMOCO', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': 0.0, 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']}
                    linha.update(campos_rota)
                    rdf.append(linha)
                elif it['ir']:
                    linha = {'PROTOCOLO': 'RETORNO_BASE', 'LATITUDE': it['lt'], 'LONGITUDE': it['Lt'], 'BASE_ATRIBUIDA': bn, 'ORDEM': og, 'NOME_DIA': dn, 'DIA_MES': it['dm'], 'SEMANA': it['s'], 'DIA': it['d'], 'PERIODO': pv, 'DISTANCIA_PONTO_ANTERIOR_KM': round(it['dk'], 2), 'ROTA_GEOMETRIA': g, 'PRIORIDADE': 'Não', 'HORA_INICIO': it['hi'].strftime('%H:%M'), 'HORA_FIM': it['hf'].strftime('%H:%M'), '_HORA_INICIO_DT': it['hi'], '_HORA_FIM_DT': it['hf']}
                    linha.update(campos_rota)
                    rdf.append(linha)
                else:
                    ob = it['o']
                    ob['ORDEM'], ob['NOME_DIA'], ob['DIA_MES'], ob['SEMANA'], ob['DIA'], ob['PERIODO'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, dn, it['dm'], it['s'], it['d'], pv, round(it['dk'], 2)
                    ob['ROTA_GEOMETRIA'] = g
                    ob.update(campos_rota)
                    ob['HORA_INICIO'], ob['HORA_FIM'], ob['_HORA_INICIO_DT'], ob['_HORA_FIM_DT'] = it['hi'].strftime('%H:%M'), it['hf'].strftime('%H:%M'), it['hi'], it['hf']
                    rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf); del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            st_v['b_idx'] += 1; st.session_state.vrp_state_fisc = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas de Fiscalização Traçadas!"); pb.progress(1.0)
        st.session_state.df_routed_fisc = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status_fisc = "PACKAGING"; time.sleep(1); tentar_rerun()

# ==========================================
# PACOTES E DOWNLOAD
# ==========================================
if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    if 'start_time_pkg_fisc' not in st.session_state:
        st.session_state.start_time_pkg_fisc = time.time()

    df_routed = st.session_state.df_routed_fisc.copy()
    d_fmt = datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx, bu_txt = io.BytesIO(), io.BytesIO(), io.BytesIO(), io.BytesIO()

    try:
        from modules.export_fisc import limpar_colunas_fisc, gerar_txt_fisc

        cfg_pack = st.session_state.vrp_state_fisc.get('config', {})
        data_ini = cfg_pack.get('data_inicio', datetime.today())
        if isinstance(data_ini, datetime):
            data_ini = data_ini.date()

        dias_semana_pt = {
            0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA",
            3: "QUINTA-FEIRA", 4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO"
        }

        def dia_semana_da_linha(row):
            data_txt = str(row.get('DIA_MES', '')).strip()
            try:
                dt_linha = datetime.strptime(data_txt, '%d/%m/%Y')
                return dias_semana_pt[dt_linha.weekday()]
            except Exception:
                nd = str(row.get('NOME_DIA', '')).strip()
                return nd.upper() if nd else dias_semana_pt[data_ini.weekday()]

        # Indicadores agora respeitam a configuração escolhida pelo usuário.
        obras_por_dia_est = cfg_pack.get('obras_por_dia', 4)
        dias_uteis_est = len(cfg_pack.get('dias_selecionados', []) or [])
        if not cfg_pack.get('is_tatico', True) or not isinstance(obras_por_dia_est, (int, float)) or obras_por_dia_est >= 999999:
            obras_por_dia_est = 4
        if dias_uteis_est <= 0 or not cfg_pack.get('is_tatico', True):
            dias_uteis_est = 5

        # Relatório de configuração/rastreabilidade sem interferir na roteirização.
        config_linhas = [
            "CONFIGURAÇÃO DA ROTEIRIZAÇÃO DE FISCALIZAÇÃO",
            f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            f"Modo de operação: {cfg_pack.get('modo_operacao', '-')}",
            f"Regra de atribuição: {cfg_pack.get('regra_atribuicao', '-')}",
            f"Data de início: {data_ini.strftime('%d/%m/%Y')}",
            f"Cota diária por equipe: {cfg_pack.get('obras_por_dia', '-')}",
            f"Visão de trabalho: {cfg_pack.get('tipo_periodo', '-')}",
            f"Quantidade de períodos: {cfg_pack.get('limite_periodos', '-')}",
            f"Dias selecionados: {', '.join(cfg_pack.get('dias_selecionados', []) or [])}",
            f"Raio Super Ponto: {cfg_pack.get('raio_super_ponto_m', '-')} m",
            f"Trava global: {cfg_pack.get('trava_global', 0)}",
            f"Traçado real OSRM: {'SIM' if cfg_pack.get('tracado_real') else 'NÃO'}",
            f"Endpoint OSRM: {cfg_pack.get('url_osrm_base', '-')}",
            f"Sentido da rota: {cfg_pack.get('sentido_rota', '-')}",
            f"Velocidade média configurada: {cfg_pack.get('velocidade_media_kmh', '-')} km/h",
        ]
        config_txt = "\n".join(config_linhas) + "\n"

        # ---------------- RESUMO OPERACIONAL ----------------
        res = []
        for b in df_routed['BASE_ATRIBUIDA'].dropna().unique():
            db = df_routed[(df_routed['BASE_ATRIBUIDA'] == b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
            br = next((x for x in st.session_state.bases_records_fisc if x['LEVANTADOR'] == b), None)
            qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
            pms = pd.to_numeric(db['QTD PREVISTA DE POSTES'], errors='coerce').fillna(0).round().astype(int) if 'QTD PREVISTA DE POSTES' in db.columns else pd.Series(dtype='int64')

            qtd_obras = sum(count_real_obras(r) for _, r in db.iterrows())
            postes_dia = (int(pms.sum()) / (qtd_obras / float(obras_por_dia_est))) if qtd_obras > 0 else 0
            postes_semana = postes_dia * float(dias_uteis_est)
            km_rod = pd.to_numeric(db.get('DISTANCIA_RODOVIARIA_KM', pd.Series(dtype='float64')), errors='coerce').sum(min_count=1)
            qtd_sem_rota = int(db.get('STATUS_ROTA', pd.Series(dtype='object')).astype(str).str.contains('SEM_ROTA', na=False).sum()) if 'STATUS_ROTA' in db.columns else 0

            item_res = {
                'FISCAL': b,
                'TIPO EQUIPE': br.get('TIPO_EQUIPE', 'PRINCIPAL') if br else 'DESCONHECIDO',
                'Obras Roteirizadas': qtd_obras,
                'Postes/Dia (Est.)': int(round(postes_dia)),
                'Postes/Semana (Est.)': int(round(postes_semana)),
                'POSTES AUDITADOS': int(pms.sum()),
                'SUPER PONTOS': qs,
                'KM TOTAL PREVISTO': round(df_routed[df_routed['BASE_ATRIBUIDA'] == b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2),
                'TRECHOS SEM ROTA OSRM': qtd_sem_rota
            }
            if pd.notna(km_rod):
                item_res['KM RODOVIÁRIO CALCULADO'] = round(float(km_rod), 2)
            res.append(item_res)
        df_resumo = pd.DataFrame(res)

        # ---------------- DATAFRAME EXPANDIDO ----------------
        linhas_gerais = []
        for _, r in df_routed.iterrows():
            if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']:
                continue

            is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
            sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else "NÃO"
            dia_semana_linha = dia_semana_da_linha(r)

            if is_sp:
                for orig in r['_ORIGINAL_ROWS']:
                    nr = r.copy()
                    for k, v in orig.items():
                        if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'ROTA_GEOMETRIA', 'PERIODO', 'NOME_DIA', 'DIA_MES']:
                            nr[k] = v
                    # Preserva DIA_MES calculado pelo motor e deriva o dia da semana dessa data.
                    nr['DIA_SEMANA'] = dia_semana_linha
                    nr['SUPER_PONTO'] = sp_text
                    linhas_gerais.append(nr)
            else:
                rn = r.copy()
                rn['DIA_SEMANA'] = dia_semana_linha
                rn['SUPER_PONTO'] = sp_text
                linhas_gerais.append(rn)

        df_excel_full = pd.DataFrame(linhas_gerais)

        # Garante que todos os arquivos exportados usem o fiscal realmente
        # atribuido pelo roteirizador. A coluna FISCAL original pode estar
        # vazia em algumas linhas da planilha de entrada.
        if 'BASE_ATRIBUIDA' in df_excel_full.columns:
            df_excel_full['FISCAL'] = df_excel_full['BASE_ATRIBUIDA']

        for c in df_excel_full.columns:
            if 'POSTE' in c.upper():
                df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

        col_exibir = st.session_state.colunas_exibir_fisc.copy()
        if 'NOME_DIA' not in col_exibir:
            col_exibir.insert(0, 'NOME_DIA')
        if 'DIA_MES' not in col_exibir:
            col_exibir.insert(1, 'DIA_MES')
        if 'SUPER_PONTO' not in col_exibir:
            col_exibir.insert(2, 'SUPER_PONTO')

        fiscais_reais = [f for f in df_routed['BASE_ATRIBUIDA'].dropna().unique() if f != "NÃO ALOCADO"]

        # ---------------- DIAGNÓSTICO DE FALHAS / INTEGRIDADE ----------------
        df_falhas_rota = pd.DataFrame()
        if 'STATUS_ROTA' in df_routed.columns:
            df_falhas_rota = df_routed[
                (~df_routed['PROTOCOLO'].isin(['PAUSA_ALMOCO'])) &
                (df_routed['STATUS_ROTA'].astype(str).str.contains('SEM_ROTA', na=False))
            ].copy()

        df_u = st.session_state.get('df_unallocated_fisc', pd.DataFrame())
        roteadas_reais = sum(count_real_obras(r) for _, r in df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].iterrows())
        nao_alocadas_reais = sum(
            len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1
            for _, r in df_u.iterrows()
        ) if not df_u.empty else 0
        integridade_txt = (
            "VERIFICAÇÃO DE INTEGRIDADE\n"
            f"Obras roteirizadas: {roteadas_reais}\n"
            f"Obras não alocadas/excedentes: {nao_alocadas_reais}\n"
            f"Trechos sem rota OSRM: {len(df_falhas_rota)}\n"
            f"Obras retidas para correção: {len(st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame()))}\n"
        )

        # ================================================================
        # PLANILHAS ZIP
        # ================================================================
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx:
            zx.writestr(f"Resumo_Fiscais - {d_fmt}.xlsx", gerar_excel_resumo_fisc(df_resumo))
            zx.writestr(f"Configuracao_Roteirizacao - {d_fmt}.txt", config_txt.encode('utf-8'))
            zx.writestr(f"Integridade_Roteirizacao - {d_fmt}.txt", integridade_txt.encode('utf-8'))

            dfc = st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy()
                dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object':
                        dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO()
                dfcc.to_excel(out_e, index=False)
                zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())

            if not df_falhas_rota.empty:
                falhas_saida = df_falhas_rota.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS'], errors='ignore')
                out_f = io.BytesIO()
                falhas_saida.to_excel(out_f, index=False)
                zx.writestr(f"Trechos_Sem_Rota_OSRM - {d_fmt}.xlsx", out_f.getvalue())

            dfg = limpar_colunas_fisc(
                df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'),
                st.session_state.colunas_originais_fisc
            )
            dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
            for cc in dfg.columns:
                if str(dfg[cc].dtype) == 'object':
                    dfg[cc] = dfg[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Fiscalizacao - {d_fmt}.xlsx", gerar_excel_fisc(dfg, st.session_state.colunas_originais_fisc))

            for b_name in fiscais_reais:
                ns = re.sub(r'[^A-Za-z0-9_ -]', '', str(b_name)).strip()
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == b_name]
                if df_base_excel.empty:
                    continue
                dfg_b = limpar_colunas_fisc(
                    df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'),
                    st.session_state.colunas_originais_fisc
                )
                dfg_b = dfg_b.loc[:, ~dfg_b.columns.duplicated()].copy()
                for cc in dfg_b.columns:
                    if str(dfg_b[cc].dtype) == 'object':
                        dfg_b[cc] = dfg_b[cc].astype(str).replace('nan', '')
                zx.writestr(f"Rotas_{d_fmt}/Rota_{ns}.xlsx", gerar_excel_fisc(dfg_b, st.session_state.colunas_originais_fisc))

        # ================================================================
        # TXT ZIP
        # ================================================================
        with zipfile.ZipFile(bu_txt, 'w', zipfile.ZIP_DEFLATED) as zt:
            zt.writestr(f"Configuracao_Roteirizacao - {d_fmt}.txt", config_txt.encode('utf-8'))
            zt.writestr(f"Integridade_Roteirizacao - {d_fmt}.txt", integridade_txt.encode('utf-8'))
            txt_total = gerar_txt_fisc(df_excel_full, st.session_state.colunas_originais_fisc)
            zt.writestr(f"Demanda_Fiscalizacao_Total - {d_fmt}.txt", txt_total.encode('utf-8'))

            if not df_falhas_rota.empty:
                falhas_txt = ["TRECHOS SEM ROTA OSRM"]
                for _, rr in df_falhas_rota.iterrows():
                    falhas_txt.append(f"{rr.get('BASE_ATRIBUIDA', '-')} | {rr.get('PROTOCOLO', '-')} | {rr.get('DIA_MES', '-')} | ordem {rr.get('ORDEM', '-')}")
                zt.writestr(f"Trechos_Sem_Rota_OSRM - {d_fmt}.txt", "\n".join(falhas_txt).encode('utf-8'))

            for b_name in fiscais_reais:
                ns = re.sub(r'[^A-Za-z0-9_ -]', '', str(b_name)).strip()
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == b_name]
                if not df_base_excel.empty:
                    txt_ind = gerar_txt_fisc(df_base_excel, st.session_state.colunas_originais_fisc)
                    zt.writestr(f"Relatorios_TXT_{d_fmt}/ROTA_{ns}.txt", txt_ind.encode('utf-8'))

        # ================================================================
        # KML ZIP
        # ================================================================
        with zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk:
            zk.writestr(f"Configuracao_Roteirizacao - {d_fmt}.txt", config_txt.encode('utf-8'))
            for b_name in fiscais_reais:
                ns = re.sub(r'[^A-Za-z0-9_ -]', '', str(b_name)).strip()
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == b_name) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if dfk_base.empty:
                    continue
                dfk_base['DIA_SEMANA'] = dfk_base.apply(dia_semana_da_linha, axis=1)
                dfk_base['SUPER_PONTO'] = dfk_base.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS']) > 1 else "NÃO", axis=1)
                kl = gerar_kml_fisc(dfk_base, f"Rota {ns}", col_exibir, [b_name], formatar_valor_coluna)
                zk.writestr(f"KML_{d_fmt}/Rota_{ns}.kml", kl.encode('utf-8'))

            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                dfk_total['DIA_SEMANA'] = dfk_total.apply(dia_semana_da_linha, axis=1)
                dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS']) > 1 else "NÃO", axis=1)
                ks = gerar_kml_fisc(dfk_total, "ROTA_TOTAL", col_exibir, fiscais_reais, formatar_valor_coluna)
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks.encode('utf-8'))

            if not df_u.empty:
                ku = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document><name>OBRAS NÃO ALOCADAS</name>', '<Style id="wp"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href></Icon></IconStyle></Style>']
                for _, r in df_u.iterrows():
                    if pd.notna(r.get('LATITUDE')) and pd.notna(r.get('LONGITUDE')):
                        ku.append(f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO", "Rejeitado")))}</name><styleUrl>#wp</styleUrl><Point><coordinates>{r.get("LONGITUDE")},{r.get("LATITUDE")}</coordinates></Point></Placemark>')
                ku.append('</Document></kml>')
                zk.writestr(f"OBRAS_NAO_ALOCADAS - {d_fmt}.kml", "\n".join(ku).encode('utf-8'))

        # ================================================================
        # GPX ZIP
        # ================================================================
        with zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg:
            zg.writestr(f"Configuracao_Roteirizacao - {d_fmt}.txt", config_txt.encode('utf-8'))
            for b_name in fiscais_reais:
                ns = re.sub(r'[^A-Za-z0-9_ -]', '', str(b_name)).strip()
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == b_name) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if not dfk_base.empty:
                    zg.writestr(f"GPX_{d_fmt}/Rota_{ns}.gpx", gerar_gpx_simples(dfk_base, f"Rota {ns}").encode('utf-8'))

            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

        st.session_state.bytes_zip_xl_fisc = bu_xl.getvalue()
        st.session_state.bytes_zip_kml_fisc = bu_kml.getvalue()
        st.session_state.bytes_zip_gpx_fisc = bu_gpx.getvalue()
        st.session_state.bytes_zip_txt_fisc = bu_txt.getvalue()
        st.session_state.roteamento_concluido_fisc = True
        st.session_state.vrp_status_fisc = "IDLE"
        tentar_rerun()

    except Exception as e:
        st.error(f"🚨 ERRO: {e}")
        st.session_state.vrp_status_fisc = "IDLE"
