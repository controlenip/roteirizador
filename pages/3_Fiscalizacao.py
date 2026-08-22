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
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formatar_moeda, formata_campo_html, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas

# IMPORTAÇÕES DA NOVA ARQUITETURA ISOLADA DE EXPORTAÇÃO
from modules.export_utils import injetar_logo, gerar_excel_fisc, gerar_excel_resumo_fisc, gerar_gpx_simples, gerar_kml_fisc, identificar_icone_folium, limpar_colunas_fisc

st.set_page_config(page_title="Fiscalização", page_icon="📋", layout="wide")
injetar_logo()

# ==========================================
# FUNÇÕES VISUAIS E AUXILIARES
# ==========================================
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
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} Metros"
        if 'POSTE' in c.upper(): return f"{int(round(vf))}"
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

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_fisc': False, 'vrp_status_fisc': "IDLE", 'vrp_state_fisc': {}, 'df_routed_fisc': pd.DataFrame(), 'bases_records_fisc': [], 'colunas_exibir_fisc': [], 'colunas_originais_fisc': []})
    for k in ['bytes_zip_xl_fisc', 'bytes_zip_kml_fisc', 'bytes_zip_gpx_fisc', 'start_time_run_fisc', 'start_time_pkg_fisc', 'df_unallocated_fisc', 'df_correcao_fiscalizacao']: st.session_state.pop(k, None)
    ler_planilha_cached.clear(); tentar_rerun()

# ==========================================
# INÍCIO DA PÁGINA
# ==========================================
if "roteamento_concluido_fisc" not in st.session_state: st.session_state.roteamento_concluido_fisc = False
if "vrp_status_fisc" not in st.session_state: st.session_state.vrp_status_fisc = "IDLE"

status_exec = st.session_state.vrp_status_fisc
is_done = st.session_state.roteamento_concluido_fisc
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📋 Planejamento de Fiscalização</h1>", unsafe_allow_html=True)
st.info("💡 **A Regra do Bolsão:** A IA ancora os Fiscais nas obras com MAIS POSTES primeiro. Em seguida, varre as obras menores e finaliza a rota exatamente no maior foco do mapa.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Esforço e Limites", expanded=True):
        trava_global = st.number_input("Trava Total de Obras no Estado", min_value=0, value=0, step=50, disabled=is_locked)
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (Metros)", 10, 1000, 100, 10, disabled=is_locked)
        st.markdown("---")
        
        st.success("📦 **Carga Total:** O sistema roteirizará 100% das obras da planilha (Modo Contínuo).")
        obras_dia = 999999
        limite_per = 1
        tpc = "Dia"
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
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_fisc', b"vazio"), file_name=f"Fiscais_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_fisc', b"vazio"), file_name=f"Fiscais_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

# ==========================================
# EXIBIÇÃO DE RESULTADOS (SE CONCLUÍDO)
# ==========================================
if is_done and not st.session_state.df_routed_fisc.empty:
    st.markdown("## 🎯 Dashboards de Produtividade")
    
    # JUSTIFICATIVA DOS ARQUIVOS DE CORREÇÃO (CARD EM DESTAQUE)
    df_c = st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica Oficial:</b> As obras listadas no arquivo <b>"Obras_Correcao"</b> foram bloqueadas e não roteirizadas porque apresentaram <b>coordenadas geográficas em branco, zeradas, invertidas</b> ou porque o GPS da obra apontou para um local que está <b>fora da Cerca Eletrônica de 70km</b> do município preenchido na planilha. O bloqueio é automático para garantir que a rota em campo não seja corrompida.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.df_routed_fisc['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_fisc.groupby(['BASE_ATRIBUIDA', 'PERIODO'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed_fisc.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = sum(count_real_obras(r) for _, r in dfr_t.iterrows())
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    tot_postes_global = int(pd.to_numeric(dfr_t['QTD PREVISTA DE POSTES'], errors='coerce').sum()) if 'QTD PREVISTA DE POSTES' in dfr_t else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("TOTAL DE OBRAS", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Fiscais Alocados", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Postes Auditados", tot_postes_global, "🏗️", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Previsto Total", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    chart_data = []
    for b_name in dfr_t['BASE_ATRIBUIDA'].unique():
        df_f = dfr_t[dfr_t['BASE_ATRIBUIDA'] == b_name]
        q_obras = sum(count_real_obras(r) for _, r in df_f.iterrows())
        q_postes = pd.to_numeric(df_f['QTD PREVISTA DE POSTES'], errors='coerce').sum() if 'QTD PREVISTA DE POSTES' in df_f.columns else 0
        if q_obras > 0: chart_data.append({"Fiscal": b_name, "Obras": q_obras, "Postes": int(q_postes)})
    
    df_chart = pd.DataFrame(chart_data)
    if not df_chart.empty:
        df_chart['Perc_Postes'] = (df_chart['Postes'] / df_chart['Postes'].sum() * 100).fillna(0)
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
            
            # SOLUÇÃO PROFISSIONAL: PLOTLY PARA O GRÁFICO DE ROSCA (CRIA SETAS E LINHAS DE CHAMADA AUTOMÁTICAS)
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
            
        # Botão de Exportação para Relatório Executivo (PDF Seguro)
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors

            def gerar_pdf_bytes(df_res):
                pdf_buf = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buf, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elements = []
                styles = getSampleStyleSheet()
                
                title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0D256C'), spaceAfter=15)
                elements.append(Paragraph("<b>Relatório Executivo de Fiscalização - NIP v3.0</b>", title_style))
                elements.append(Spacer(1, 10))
                
                table_data = [["Fiscal", "Total Obras", "Postes Auditados", "Fat. Postes"]]
                for _, row in df_res.iterrows():
                    table_data.append([str(row['Fiscal']), str(row['Obras']), str(row['Postes']), f"{row.get('Perc_Text', '')}"])
                    
                t = Table(table_data, colWidths=[200, 90, 100, 110])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0D256C')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0,0), (-1,0), 8),
                    ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f8f9fa')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dee2e6')),
                ]))
                elements.append(t)
                doc.build(elements)
                return pdf_buf.getvalue()

            st.download_button("📥 Baixar Relatório Executivo (PDF)", data=gerar_pdf_bytes(df_chart), file_name=f"Relatorio_Fiscalizacao - {datetime.now().strftime('%d.%m.%Y')}.pdf", mime="application/pdf", use_container_width=True)
        except Exception:
            st.info("💡 Dica: Você pode salvar os gráficos em alta resolução clicando no ícone de câmera no canto superior direito do gráfico.")

    st.markdown("### 🗺️ Mapa Geográfico")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        for pe in db['PERIODO'].unique():
            dp = db[db['PERIODO'] == pe]
            pts = [p for _, r in dp.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
            folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
            folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
            
            for r in dp.to_dict('records'):
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                
                c_i = r.get('COR_ICONE', 'gray')
                qtd_p = int(float(r.get('QTD PREVISTA DE POSTES', 0)))
                ic = identificar_icone_folium(r, dfr.columns)
                
                bg_colors = {'green': '#4CAF50', 'blue': '#2196F3', 'beige': '#FFC107', 'orange': '#FF9800', 'red': '#F44336', 'gray': '#9E9E9E'}
                txt_colors = {'beige': '#000000', 'orange': '#000000', 'green': '#ffffff', 'blue': '#ffffff', 'red': '#ffffff', 'gray': '#ffffff'}
                
                if str(r.get('SUPER_PONTO', '')).startswith('SIM'):
                    p_bg, p_c = '#FFD700', '#000000'
                    p_txt = f"🏢 SUPER PONTO {str(r.get('SUPER_PONTO')).replace('SIM','').strip()}"
                else:
                    p_bg, p_c = bg_colors.get(c_i, '#9E9E9E'), txt_colors.get(c_i, '#ffffff')
                    p_txt = f"📋 FISCALIZAÇÃO - {qtd_p} POSTES"
                
                er = "".join([f"<tr><td style='padding:3px;'><b>{html.escape(c)}:</b></td><td style='padding:3px;'>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_fisc if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA','COR_ICONE']])
                pop_html = f'<div style="width:280px;"><div style="background:{p_bg};color:{p_c};padding:8px;font-weight:bold;">{p_txt}</div><table border="1" style="width:100%;font-size:12px;"><tr><td style="padding:3px;"><b>Ordem:</b></td><td style="padding:3px;">{r.get("ORDEM",0)}</td></tr>{er}</table></div>'
                
                folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    t1, t2 = st.tabs(["📊 Dados Tabulares", "📉 Resumo por Técnico"])
    with t1:
        st.data_editor(st.session_state.df_routed_fisc.drop(columns=['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'PERIODO', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'BASE_ATRIBUIDA', 'COR_ICONE'], errors='ignore'), use_container_width=True)
    with t2:
        dr = pd.DataFrame([{"Fiscal": b['LEVANTADOR'], "Obras Roteirizadas": sum(count_real_obras(r) for _, r in dfr_t[dfr_t['BASE_ATRIBUIDA']==b['LEVANTADOR']].iterrows())} for b in st.session_state.bases_records_fisc]).reset_index(drop=True)
        st.dataframe(dr, use_container_width=True)

# ==========================================
# START DA APLICAÇÃO (UPLOAD DE DADOS)
# ==========================================
elif status_exec == "IDLE":
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.markdown("### 👥 1. Fiscais")
        df_bases = pd.DataFrame()
        bf = st.file_uploader("Suba a planilha de Fiscais (Excel)", type=["xlsx", "xls"])
        if bf:
            b_t = ler_planilha_cached(bf.getvalue()); b_t.columns = normalize_cols(b_t.columns)
            b_t = b_t.loc[:, ~b_t.columns.duplicated()].copy()
            for pn in ['NOME', 'FISCAL', 'TECNICO', 'COLABORADOR']:
                if pn in b_t.columns: b_t = b_t.rename(columns={pn: 'LEVANTADOR'}); break
            if 'LEVANTADOR' in b_t.columns:
                
                b_t['LEVANTADOR'] = b_t['LEVANTADOR'].astype(str).str.split(r'\s*\|\s*')
                b_t = b_t.explode('LEVANTADOR').reset_index(drop=True)
                b_t['LEVANTADOR'] = b_t['LEVANTADOR'].str.strip().str.upper()
                
                opts = sorted([str(x) for x in b_t['LEVANTADOR'].dropna().unique() if str(x) not in ['SEM LEVANTADOR', 'NAN', 'NONE', '']])
                sel = st.multiselect("Selecione os Fiscais Ativos:", opts, default=opts)
                if sel:
                    df_bases = b_t[b_t['LEVANTADOR'].isin(sel)].copy()
                    if 'LATITUDE' in df_bases.columns and 'LONGITUDE' in df_bases.columns:
                        df_bases['LATITUDE'] = pd.to_numeric(df_bases['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                        df_bases['LONGITUDE'] = pd.to_numeric(df_bases['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
                    elif 'RESIDENCIA' in df_bases.columns or 'MUNICIPIO' in df_bases.columns:
                        cr = 'RESIDENCIA' if 'RESIDENCIA' in df_bases.columns else 'MUNICIPIO'
                        mc = {}
                        with st.spinner("🌍 Mapeando bases..."):
                            for m in df_bases[cr].dropna().unique(): mc[m] = obter_coordenadas_municipio_cached(m)
                        df_bases['LATITUDE'], df_bases['LONGITUDE'] = df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[0]), df_bases[cr].map(lambda x: mc.get(x, (np.nan, np.nan))[1])
                    df_bases = df_bases.dropna(subset=['LATITUDE', 'LONGITUDE']); df_bases['TIPO_EQUIPE'] = 'FISCAL'
            else: st.error("❌ A planilha não possui a coluna 'FISCAL'.")

        st.markdown("##### 📍 Regra de Atribuição")
        ta = st.radio("Atribuição", ["Por Proximidade (Recomendado)", "Por Município Rígido"], index=0, label_visibility="collapsed")
        if "Proximidade" in ta: st.caption("A IA persegue os maiores Bolsões e puxa o Fiscal mais próximo.")
        else: st.caption("Trava o Fiscal apenas dentro da cidade informada na sua planilha.")

    with c_up2:
        st.markdown("### 📁 2. Obras de Fiscalização")
        task_files = st.file_uploader("Suba as Demandas", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if df_bases.empty or not task_files: st.stop()
    
    qtd_eq = df_bases['LEVANTADOR'].nunique()
    cm = obras_dia * (len(dias_sel) if tpc == 'Semana' else 1) * limite_per
    sb_html.markdown(render_sidebar_card("Ilimitada", 0, qtd_eq, "Ilimitada"), unsafe_allow_html=True)

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
    with cg1: st.markdown("#### 🌍 Cerca Eletrônica (Geofencing 70km)")
    with cg2:
        if st.button("⏹️ Abortar", use_container_width=True): limpar_roteirizador(); st.stop()
    
    pbg = st.progress(0.0); tmp = st.empty(); sgt = st.empty(); df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    m_m = df_tasks['MUNICIPIO'].isna() | (df_tasks['MUNICIPIO'].astype(str).str.strip() == '') | (df_tasks['MUNICIPIO'].astype(str).str.strip().str.upper() == 'NAN')
    if m_m.sum() > 0:
        df_tasks.loc[m_m, 'MOTIVO_REJEICAO'] = 'Município Vazio'
        df_rej = pd.concat([df_rej, df_tasks[m_m].copy()], ignore_index=True); df_tasks = df_tasks[~m_m].copy()

    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
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
    
    if not df_tasks.empty:
        mu = df_tasks['MUNICIPIO'].unique(); tm = len(mu); md = {}
        st_run_geo = time.time()
        
        def render_t_geo(curr, total):
            e = time.time() - st_run_geo; f = curr / max(1, total)
            rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
            es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
            html_timer = f"""
            <div style="display:flex; gap:15px; margin-top: 10px; margin-bottom: 10px;">
                <div style="flex:1; padding:10px; border-radius:8px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;">
                    <div style="font-size:0.8rem; color:#6c757d; font-weight:bold; margin-bottom:2px;">⏱️ Decorrido</div>
                    <div style="font-size:1.5rem; font-weight:bold; color:#0D256C;">{es}</div>
                </div>
                <div style="flex:1; padding:10px; border-radius:8px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;">
                    <div style="font-size:0.8rem; color:#2e7d32; font-weight:bold; margin-bottom:2px;">🎯 Restante</div>
                    <div style="font-size:1.5rem; font-weight:bold; color:#1b5e20;">{rs}</div>
                </div>
            </div>
            """
            tmp.markdown(html_timer, unsafe_allow_html=True)

        for i, m in enumerate(mu):
            pbg.progress((i + 1) / max(1, tm)); sgt.info(f"🛰️ Satélite: {m}")
            render_t_geo(i + 1, tm)
            try: lt, ln = obter_coordenadas_municipio_cached(m); time.sleep(0.1)
            except: lt, ln = np.nan, np.nan
            md[m] = (lt, ln)
            
        sgt.info("📏 Aplicando Cerca...")
        mo = df_tasks.apply(lambda r: haversine_scalar(r['LATITUDE'], r['LONGITUDE'], md.get(r['MUNICIPIO'], (np.nan, np.nan))[0], md.get(r['MUNICIPIO'], (np.nan, np.nan))[1]) > 70.0 if pd.notna(md.get(r['MUNICIPIO'], (np.nan, np.nan))[0]) else False, axis=1)
        if mo.sum() > 0:
            df_tasks.loc[mo, 'MOTIVO_REJEICAO'] = 'Fora do Município (> 70km)'
            df_rej = pd.concat([df_rej, df_tasks[mo].copy()], ignore_index=True); df_tasks = df_tasks[~mo].copy()
        
        pbg.empty(); tmp.empty(); sgt.empty()
        st.session_state.df_correcao_fiscalizacao = df_rej
        
        # JUSTIFICATIVA DOS ARQUIVOS DE CORREÇÃO (CARD EM DESTAQUE - TELA DE TRIAGEM)
        if not df_rej.empty: 
            st.markdown(f"""
            <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;'>
                <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_rej)} Obras Retidas para Correção</h4>
                <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                    <b>Justificativa:</b> Estas obras apresentaram coordenadas em branco, zeradas, invertidas ou caíram fora da <b>Cerca Eletrônica de 70km</b> do município de origem.
                    Elas foram isoladas automaticamente pelo sistema para não corromper o traçado dos Fiscais. Faça o download da <b>Planilha de Correção</b> na etapa de Empacotamento para verificar a falha de cada uma.
                </p>
            </div>
            """, unsafe_allow_html=True)

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    if 'QTD PREVISTA DE POSTES' in df_tasks.columns:
        df_tasks['QTD PREVISTA DE POSTES'] = df_tasks['QTD PREVISTA DE POSTES'].apply(extrair_qtd)
        df_tasks['COR_ICONE'] = df_tasks['QTD PREVISTA DE POSTES'].apply(definir_cor_fiscalizacao)
    else: df_tasks['QTD PREVISTA DE POSTES'], df_tasks['COR_ICONE'] = 0.0, 'gray'

    df_tasks, qc = fundir_super_pontos(df_tasks, raio_metros=raio_sp, agrupar_por_levantador=False)
    
    if 'QTD PREVISTA DE POSTES' in df_tasks.columns:
        df_tasks['QTD PREVISTA DE POSTES'] = df_tasks['QTD PREVISTA DE POSTES'].apply(extrair_qtd)
        df_tasks['COR_ICONE'] = df_tasks['QTD PREVISTA DE POSTES'].apply(definir_cor_fiscalizacao)

    tbr = df_bases.to_dict('records')
    fiscal_anchors = {b['LEVANTADOR']: (float(b.get('LATITUDE',0)), float(b.get('LONGITUDE',0))) for b in tbr}
    assigned_tasks = []
    unassigned_tasks = []
    
    df_tasks = df_tasks.sort_values(by=['QTD PREVISTA DE POSTES', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])
    if trava_global > 0: df_tasks = df_tasks.head(trava_global)
        
    for r in df_tasks.to_dict('records'):
        la, lo = r.get('LATITUDE'), r.get('LONGITUDE')
        ms = normalizar_municipios(pd.Series([str(r.get('MUNICIPIO', ''))])).iloc[0]
        
        if "Município" in ta: vb = [b for b in tbr if ms in str(b.get('MUNICIPIO', b.get('RESIDENCIA', ''))).upper()]
        else: vb = tbr
            
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
            r['MUN_LIMPO'] = ms
            assigned_tasks.append(r)
            fiscal_anchors[best_f] = (la, lo)
        else:
            r['MOTIVO_REJEICAO'] = "Fora de Área (Sem Fiscal)"
            r['BASE_ATRIBUIDA'] = "NÃO ALOCADO"
            unassigned_tasks.append(r)

    df_ta, df_u = pd.DataFrame(assigned_tasks), pd.DataFrame(unassigned_tasks)
    st.session_state.df_unallocated_fisc, st.session_state.tot_obras_nao_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_u.iterrows())
    
    total_validadas = sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    sb_html.markdown(render_sidebar_card("Ilimitada", total_validadas, qtd_eq, "Ilimitada"), unsafe_allow_html=True)
    
    if df_ta.empty: st.error("Nenhuma obra pôde ser alocada aos Fiscais."); st.stop()

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'COR_ICONE' and c != 'MUN_LIMPO']
        cd = ['PROTOCOLO', 'VALOR DA OBRA', 'QTD PREVISTA DE POSTES', 'PREVISAO DE ENTREGA', 'PARCEIRO', 'TIPO DE FISCALIZACAO', 'TIPO DE PROJETO', 'REGIONAL', 'MUNICIPIO', 'LATITUDE', 'LONGITUDE', 'ZONA', 'STATUS DA FISCALIZACAO', 'LEVANTADOR', 'BACKOFFICE DA FISCALIZACAO', 'OBSERVACAO']
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis:", tc, default=cp)
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True):
        st.session_state.update({'bases_records_fisc': tbr, 'colunas_exibir_fisc': colunas_exibir})
        st.session_state.vrp_state_fisc = {'config': {'velocidade_media_kmh': vel_kmh, 'obras_por_dia': 999999, 'tipo_periodo': tpc, 'limite_periodos': limite_per, 'dias_selecionados': dias_sel, 'url_osrm_base': url_osrm, 'tracado_real': usa_osrm, 'data_inicio': data_ini, 'tempo_medio_obra': 1.0, 'sentido_rota': sentido_rota}, 'b_names': list(set([b['LEVANTADOR'] for b in tbr])), 'b_idx': 0, 'unvisited': df_ta.copy(), 'routed_data': [], 'current_geoms': []}
        st.session_state.vrp_status_fisc = "RUNNING"; tentar_rerun()

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
            br = pd.DataFrame(st.session_state.bases_records_fisc)
            br = br[br['LEVANTADOR'] == bn].iloc[0]
            if pd.isna(br.get('LATITUDE')): st_v['b_idx'] += 1; st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
            bl, bL = float(br['LATITUDE']), float(br['LONGITUDE'])
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            
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
                vkr = haversine_vectorized(es['l'], es['L'], o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                vm = (vk / (cfg['velocidade_media_kmh']*1.5 if vk>20 else cfg['velocidade_media_kmh']))*60
                
                cp = es['t'] + pd.Timedelta(minutes=vm)
                fp = cp + pd.Timedelta(minutes=60)
                
                rf.append({'o': o, 'il': False, 'ir': False, 'la': es['l'], 'La': es['L'], 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 's': sa, 'd': da, 'ds': dds, 'dn': ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][es['d'].weekday()], 'dm': es['d'].strftime('%d/%m/%Y'), 'hi': cp, 'hf': fp, 'vm': vm, 'dk': vk})
                es['l'], es['L'], es['t'] = o['LATITUDE'], o['LONGITUDE'], fp
                
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []; st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (30 if cfg['tracado_real'] else len(rf)), len(rf))
            for i in range(oi, ei):
                it = rf[i]
                if not cfg['tracado_real']: 
                    gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
                else:
                    if i%5==0: sgt.info(f"🛣️ Traçando arruamento real **{bn}**... ({i}/{len(rf)})")
                    render_t(b_i, i, len(rf))
                    time.sleep(0.15) 
                    try: 
                        res_ruas = obter_rota_ruas(it['la'], it['La'], it['lt'], it['Lt'], cfg['url_osrm_base'], cfg['velocidade_media_kmh'])
                        gd.append(res_ruas)
                    except: 
                        gd.append(([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600))
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_fisc = st_v; tentar_rerun(); st.stop()
            
            rdf, og = [], 1
            for it, (g, ds) in zip(rf, gd):
                pv = it['s'] if cfg['tipo_periodo']=="Semana" else it['d']
                ob = it['o']; ob['ORDEM'], ob['NOME_DIA'], ob['DIA_MES'], ob['SEMANA'], ob['DIA'], ob['PERIODO'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, it['dn'], it['dm'], it['s'], it['d'], pv, round(it['dk'], 2)
                ob['ROTA_GEOMETRIA'] = g
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
    df_routed, d_fmt = st.session_state.df_routed_fisc, datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx = io.BytesIO(), io.BytesIO(), io.BytesIO()
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg:
            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                br = next((x for x in st.session_state.bases_records_fisc if x['LEVANTADOR']==b), None)
                qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                pms = pd.to_numeric(db['QTD PREVISTA DE POSTES'], errors='coerce').fillna(0).round().astype(int)
                res.append({'FISCAL': b, 'TIPO EQUIPE': br.get('TIPO_EQUIPE', 'PRINCIPAL') if br else 'DESCONHECIDO', 'TOTAL OBRAS': sum(count_real_obras(r) for _, r in db.iterrows()), 'SUPER PONTOS': qs, 'POSTES AUDITADOS': int(pms.sum()), 'KM TOTAL PREVISTO': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2)})
            zx.writestr(f"Resumo_Fiscais - {d_fmt}.xlsx", gerar_excel_resumo_fisc(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_fiscalizacao', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
            
            linhas_gerais = []
            for _, r in df_routed.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                if isinstance(r.get('_ORIGINAL_ROWS'), list):
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = r.copy()
                        for k, v in orig.items(): 
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'ROTA_GEOMETRIA', 'PERIODO']:
                                nr[k] = v
                        linhas_gerais.append(nr)
                else: linhas_gerais.append(r)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            dfg = limpar_colunas_fisc(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), st.session_state.colunas_originais_fisc)
            dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
            for cc in dfg.columns:
                if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_Fiscalizacao - {d_fmt}.xlsx", gerar_excel_fisc(dfg, st.session_state.colunas_originais_fisc))
            
            fiscais_reais = [f for f in df_routed['BASE_ATRIBUIDA'].unique() if f != "NÃO ALOCADO"]
            
            for b_name in fiscais_reais:
                ns = re.sub(r'[^A-Za-z0-9_ ]', '', str(b_name)).replace(" ", "_").upper()
                df_fisc_ind = df_routed[df_routed['BASE_ATRIBUIDA'] == b_name]
                dk = df_fisc_ind[~df_fisc_ind['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
                if dk.empty: continue
                
                ld = []
                for _, r in dk.iterrows():
                    if isinstance(r.get('_ORIGINAL_ROWS'), list):
                        for orig in r['_ORIGINAL_ROWS']:
                            nr = r.copy()
                            for k, v in orig.items(): 
                                if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'ROTA_GEOMETRIA', 'PERIODO']:
                                    nr[k] = v
                            ld.append(nr)
                    else: ld.append(r)
                if ld:
                    dx = pd.DataFrame(ld)
                    dx = limpar_colunas_fisc(dx.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), st.session_state.colunas_originais_fisc)
                    dx = dx.loc[:, ~dx.columns.duplicated()].copy()
                    for c in dx.columns:
                        if str(dx[c].dtype) == 'object': dx[c] = dx[c].astype(str).replace('nan', '')
                    zx.writestr(f"ROTA_{ns} - {d_fmt}.xlsx", gerar_excel_fisc(dx, st.session_state.colunas_originais_fisc))
                
                kl = gerar_kml_fisc(dk, f"ROTA_{ns}", st.session_state.colunas_exibir_fisc, [b_name], formatar_valor_coluna)
                zk.writestr(f"ROTA_{ns} - {d_fmt}.kml", kl.encode('utf-8'))
                zg.writestr(f"GPS_{ns} - {d_fmt}.gpx", gerar_gpx_simples(dk, f"ROTA_{ns}").encode('utf-8'))

            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            ks = gerar_kml_fisc(dfk_total, f"ROTA_TOTAL", st.session_state.colunas_exibir_fisc, fiscais_reais, formatar_valor_coluna)
            zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks.encode('utf-8'))
            zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))
            
            df_u = st.session_state.get('df_unallocated_fisc', pd.DataFrame())
            if not df_u.empty:
                ku = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document><name>OBRAS NÃO ALOCADAS</name>', '<Style id="wp"><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href></Icon></IconStyle></Style>']
                for _, r in df_u.iterrows():
                    if pd.notna(r.get('LATITUDE')) and pd.notna(r.get('LONGITUDE')): ku.append(f'<Placemark><name>{html.escape(str(r.get("PROTOCOLO", "Rejeitado")))}</name><styleUrl>#wp</styleUrl><Point><coordinates>{r.get("LONGITUDE")},{r.get("LATITUDE")}</coordinates></Point></Placemark>')
                ku.append('</Document></kml>'); zk.writestr(f"OBRAS_NAO_ALOCADAS - {d_fmt}.kml", "\n".join(ku).encode('utf-8'))

        st.session_state.bytes_zip_xl_fisc, st.session_state.bytes_zip_kml_fisc, st.session_state.bytes_zip_gpx_fisc = bu_xl.getvalue(), bu_kml.getvalue(), bu_gpx.getvalue()
        st.session_state.roteamento_concluido_fisc = True; st.session_state.vrp_status_fisc = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_fisc = "IDLE"
