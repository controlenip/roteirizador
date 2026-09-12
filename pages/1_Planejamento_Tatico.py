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


def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido': False, 'vrp_status': "IDLE", 'vrp_state': {}, 'df_routed': pd.DataFrame(), 'bases_records': [], 'colunas_exibir': [], 'colunas_originais_tat': []})
    for k in ['bytes_zip_xl', 'bytes_zip_kml', 'bytes_zip_gpx', 'start_time_run', 'start_time_pkg', 'df_unallocated', 'df_correcao_tatica', 'qtd_coords_autocorrigidas']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido" not in st.session_state: st.session_state.roteamento_concluido = False
if "vrp_status" not in st.session_state: st.session_state.vrp_status = "IDLE"

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
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl', b"vazio"), file_name=f"Rotas_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml', b"vazio"), file_name=f"Rotas_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx', b"vazio"), file_name=f"Rotas_GPS - {d_fmt}.zip", use_container_width=True)
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

    st.session_state.df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = len(dfr_t)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    if 'DISTANCIA_RODOVIARIA_KM' in dfr.columns:
        km_real_s = pd.to_numeric(dfr['DISTANCIA_RODOVIARIA_KM'], errors='coerce')
        km_fallback_s = pd.to_numeric(dfr['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce').fillna(0)
        km_total_apurado = km_real_s.fillna(km_fallback_s).fillna(0).sum()
    else:
        km_total_apurado = pd.to_numeric(dfr['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce').fillna(0).sum()
    tk = f"{km_total_apurado:.1f} km"
    
    tsp = sum(1 for _, r in dfr_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Planejadas", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(tsp), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    if not st.session_state.get('df_unallocated', pd.DataFrame()).empty:
        st.warning(f"⚠️ {len(st.session_state.df_unallocated)} obras não couberam na cota de dias/equipes (Verifique o arquivo ZIP gerado).")
    else:
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

    st.markdown("### 🗺️ Mapa Operacional")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        
        bn_safe = str(bn).replace("{", "[").replace("}", "]")
        fg = folium.FeatureGroup(name=f"Equipe: {bn_safe}", show=False)
        m_clust_eq = MarkerCluster(name=f"Obras - {bn_safe}").add_to(fg)
        
        for pe in db['PERIODO'].unique():
            dp = db[db['PERIODO'] == pe]
            # Cada geometria é desenhada separadamente: uma falha intermediária não liga dois trechos por uma reta artificial.
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
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

elif status_exec == "IDLE":
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown("### 👥 1. Bases de Equipes")
        df_bases = pd.DataFrame()
        bf = st.file_uploader("Suba a planilha de Equipes (Excel)", type=["xlsx", "xls"])
        if bf:
            b_t = ler_planilha_cached(bf.getvalue()); b_t.columns = normalize_cols(b_t.columns)
            b_t = b_t.loc[:, ~b_t.columns.duplicated()].copy()
            
            for pn in ['LEVANTADOR', 'FISCAL', 'NOME_FISCAL', 'NOME', 'TECNICO', 'COLABORADOR', 'EQUIPE']:
                if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'BASE_NOME'}); break
            
            if 'BASE_NOME' in b_t.columns:
                b_t['BASE_NOME'] = b_t['BASE_NOME'].astype(str).str.split(r'\s*\|\s*')
                b_t = b_t.explode('BASE_NOME').reset_index(drop=True)
                b_t['BASE_NOME'] = b_t['BASE_NOME'].str.strip().str.upper()
                
                opts = sorted([str(x) for x in b_t['BASE_NOME'].dropna().unique() if str(x) not in ['SEM EQUIPE', 'NAN', 'NONE', '']])
                sel = st.multiselect("Selecione as Equipes Ativas:", opts, default=opts)
                if sel:
                    df_bases = b_t[b_t['BASE_NOME'].isin(sel)].copy()
                    if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                        df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                        df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                    elif 'RESIDENCIA' in df_bases.columns or 'MUNICIPIO' in df_bases.columns:
                        cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                        mc = {}
                        with st.spinner("🌍 Mapeando bases via IBGE..."):
                            for m in df_bases[cr].dropna().unique(): mc[m] = obter_coordenadas_municipio_cached(m)
                        df_bases['LATITUDE'], df_bases['LONGITUDE'] = df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0]), df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
                    df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE'])
                    cr_mun = 'MUNICIPIO' if 'MUNICIPIO' in df_bases.columns else ('RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else None)
                    if cr_mun:
                        df_bases['MUN_LIMPO_BASE'] = df_bases[cr_mun].apply(municipio_normalizado)
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
            dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
            dft.columns = normalize_cols(dft.columns)
            if not dfs: st.session_state.colunas_originais_tat = dft.columns.tolist()
            for cc in ['NOTA', 'PROTOCOLO', 'OS', 'ID']:
                if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
            dfs.append(dft)
        
        df_tasks_padrao = pd.concat(dfs, ignore_index=True)
        if 'PROTOCOLO' in df_tasks_padrao.columns:
            df_tasks_padrao['PROTOCOLO'] = df_tasks_padrao['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
            df_tasks_padrao = df_tasks_padrao.explode('PROTOCOLO').reset_index(drop=True); df_tasks_padrao['PROTOCOLO'] = df_tasks_padrao['PROTOCOLO'].str.strip()

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
                df_tasks_padrao['PRIORIDADE'] = df_tasks_padrao['TIPO NOTA'].apply(lambda x: 'Sim' if str(x) in sel_p else 'Não')
            else: df_tasks_padrao['PRIORIDADE'] = 'Não'

    # ------------------ BLOCO PLANILHA GENÉRICA ------------------
    df_tasks_gen = pd.DataFrame()
    if generic_files:
        dfs_gen = []
        for f in generic_files:
            dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
            dft.columns = normalize_cols(dft.columns)
            
            if 'colunas_originais_tat' not in st.session_state or not st.session_state.colunas_originais_tat:
                st.session_state.colunas_originais_tat = dft.columns.tolist()
            else:
                for col in dft.columns:
                    if col not in st.session_state.colunas_originais_tat:
                        st.session_state.colunas_originais_tat.append(col)
                        
            for cc in ['NOTA', 'PROTOCOLO', 'OS', 'ID']:
                if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
            dfs_gen.append(dft)
            
        df_tasks_gen = pd.concat(dfs_gen, ignore_index=True)
        
        if 'PROTOCOLO' not in df_tasks_gen.columns:
            df_tasks_gen['PROTOCOLO'] = "GEN_" + df_tasks_gen.index.astype(str)
            
        if 'PROTOCOLO' in df_tasks_gen.columns:
            df_tasks_gen['PROTOCOLO'] = df_tasks_gen['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
            df_tasks_gen = df_tasks_gen.explode('PROTOCOLO').reset_index(drop=True); df_tasks_gen['PROTOCOLO'] = df_tasks_gen['PROTOCOLO'].str.strip()

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
                df_tasks_gen['PRIORIDADE'] = df_tasks_gen[col_prio].apply(lambda x: 'Sim' if str(x) in sel_p_gen else 'Não')
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

    # A parte pesada (balanceamento + Super Pontos) foi movida para depois do clique.
    # Assim o Streamlit nao recalcula milhares de obras toda vez que um widget muda.
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

    st.info("⚡ Os cálculos pesados de distribuição e Super Pontos só serão executados após clicar em **Iniciar Motor de Roteirização**.")

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
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

    st.session_state.df_unallocated = df_u
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
        'route_cache': {}
    }

    pb_prep.progress(1.0)
    msg_prep.success("✅ Preparação concluída. Iniciando o motor de roteirização...")
    st.session_state.vrp_status = "RUNNING"
    time.sleep(0.3)
    tentar_rerun()

if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor VRP Tático")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run', time.time())
    if 'start_time_run' not in st.session_state: st.session_state.start_time_run = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            br = pd.DataFrame(st.session_state.bases_records)
            br = br[br['BASE_NOME'] == bn].iloc[0]
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
                    st.session_state.df_unallocated = pd.concat([st.session_state.get('df_unallocated', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue

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
                        st.session_state.df_unallocated = pd.concat([st.session_state.get('df_unallocated', pd.DataFrame()), pd.DataFrame([o])], ignore_index=True); continue
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
                if not cfg['tracado_real']:
                    geom = [[it['La'], it['la']], [it['Lt'], it['lt']]]
                    dur = (it['dk'] / max(cfg['velocidade_media_kmh'], 1.0)) * 3600
                    gd.append((geom, dur))
                    gm.append({'status': 'LINHA_RETA_ESTIMADA', 'km_real': np.nan, 'km_estimado': float(it['dk']), 'tempo_real_min': dur / 60.0})
                else:
                    if i % 5 == 0: sgt.info(f"🛣️ Traçando arruamento **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf))
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
                            rota = obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'])
                            if rota and len(rota) >= 2 and isinstance(rota[0], list) and len(rota[0]) >= 2:
                                geom, dur = rota[0], rota[1]
                                km_real = distancia_geometria_km(geom)
                                meta = {'status': 'OSRM', 'km_real': km_real if km_real > 0 else np.nan, 'km_estimado': float(it['dk']), 'tempo_real_min': float(dur) / 60.0 if pd.notna(dur) else np.nan}
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
            if ei < len(rf): st.session_state.vrp_state = st_v; tentar_rerun(); st.stop()
            
            bl, bL = float(pd.DataFrame(st.session_state.bases_records)[pd.DataFrame(st.session_state.bases_records)['BASE_NOME']==bn].iloc[0]['LATITUDE']), float(pd.DataFrame(st.session_state.bases_records)[pd.DataFrame(st.session_state.bases_records)['BASE_NOME']==bn].iloc[0]['LONGITUDE'])
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
            st_v['b_idx'] += 1; st.session_state.vrp_state = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        st.session_state.df_routed = pd.DataFrame(st_v['routed_data'])
        st.session_state.vrp_status = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento Tático")
    df_routed = st.session_state.df_routed.copy()
    df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = df_routed.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    
    d_fmt = datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx = io.BytesIO(), io.BytesIO(), io.BytesIO()
    
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg:
            
            obras_por_dia_est = st.session_state.vrp_state.get('config', {}).get('obras_por_dia', 4.0)

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
                    'Super Pontos': sum(1 for _, r_sp in db.iterrows() if isinstance(r_sp.get('_ORIGINAL_ROWS'), list) and len(r_sp.get('_ORIGINAL_ROWS')) > 1), 
                    'Prioridades Atendidas': len(db[db['PRIORIDADE']=='Sim']), 
                    'KM Total Previsto': round((pd.to_numeric(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_RODOVIARIA_KM'], errors='coerce').fillna(pd.to_numeric(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce')).fillna(0).sum()) if 'DISTANCIA_RODOVIARIA_KM' in df_routed.columns else pd.to_numeric(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'], errors='coerce').fillna(0).sum(), 2),
                    'Trechos OSRM': int(df_routed[df_routed['BASE_ATRIBUIDA']==b].get('STATUS_ROTA', pd.Series(dtype=str)).astype(str).str.startswith('OSRM').sum()),
                    'Trechos sem rota': int(df_routed[df_routed['BASE_ATRIBUIDA']==b].get('STATUS_ROTA', pd.Series(dtype=str)).astype(str).str.contains('SEM_ROTA', regex=False).sum())
                })
            zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_tatica(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_tatica', pd.DataFrame())
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
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'ROTA_GEOMETRIA', 'PERIODO', 'NOME_DIA', 'DIA_MES']: nr[k] = v
                        nr['SUPER_PONTO'] = sp_text
                        linhas_gerais.append(nr)
                else:
                    rn = r.copy()
                    rn['SUPER_PONTO'] = sp_text
                    linhas_gerais.append(rn)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            
            for c in df_excel_full.columns:
                if 'POSTE' in c.upper():
                    df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

            col_exibir = st.session_state.colunas_exibir.copy()
            if 'NOME_DIA' not in col_exibir: col_exibir.insert(0, 'NOME_DIA')
            if 'DIA_MES' not in col_exibir: col_exibir.insert(1, 'DIA_MES')
            if 'SUPER_PONTO' not in col_exibir: col_exibir.insert(2, 'SUPER_PONTO')

            dfg_total = limpar_colunas_tatica(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
            dfg_total = dfg_total.loc[:, ~dfg_total.columns.duplicated()].copy()
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Tatica_Total - {d_fmt}.xlsx", gerar_excel_tatica(dfg_total, st.session_state.colunas_originais_tat))
            
            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                
                tpc = st.session_state.vrp_state.get('config', {}).get('tipo_periodo', 'Semana')
                ks_tot = gerar_kml_tatica(dfk_total, "ROTA TOTAL", col_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), tpc, formatar_valor_coluna)
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks_tot.encode('utf-8'))
                zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

            for base in df_routed['BASE_ATRIBUIDA'].unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == base]
                if not df_base_excel.empty:
                    dfg = limpar_colunas_tatica(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
                    dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
                    for cc in dfg.columns:
                        if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                    zx.writestr(f"Rotas_{d_fmt}/Rota_{b_safe}.xlsx", gerar_excel_tatica(dfg, st.session_state.colunas_originais_tat))
                    
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == base) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if not dfk_base.empty:
                    dfk_base['SUPER_PONTO'] = dfk_base.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                    
                    ks = gerar_kml_tatica(dfk_base, f"Rota {b_safe}", col_exibir, [base], tpc, formatar_valor_coluna)
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kml", ks.encode('utf-8'))
                    zg.writestr(f"GPX_{d_fmt}/Rota_{b_safe}.gpx", gerar_gpx_simples(dfk_base, f"Rota {b_safe}").encode('utf-8'))

        st.session_state.bytes_zip_xl, st.session_state.bytes_zip_kml, st.session_state.bytes_zip_gpx = bu_xl.getvalue(), bu_kml.getvalue(), bu_gpx.getvalue()
        st.session_state.roteamento_concluido = True; st.session_state.vrp_status = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status = "IDLE"
