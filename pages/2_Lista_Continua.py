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
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime

# Importações dos Motores Matemáticos
from modules.data_processing import ler_planilha_cached, formata_campo_html, formatar_moeda, normalize_cols, normalizar_municipios
from modules.geospatial import haversine_vectorized, haversine_scalar, obter_coordenadas_municipio_cached, fundir_super_pontos
from modules.routing_engine import resolver_tsp_ortools, obter_rota_ruas

# IMPORTAÇÕES DA NOVA ARQUITETURA ISOLADA (LISTA CONTÍNUA)
from modules.export_lista import injetar_logo, identificar_icone_folium, gerar_excel_lista, gerar_excel_resumo_lista, gerar_gpx_simples, gerar_kml_lista, limpar_colunas_lista

st.set_page_config(page_title="Roteirizador - Lista Contínua", page_icon="📜", layout="wide")
injetar_logo()

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    try:
        vf = float(v)
        if 'DISTANCIA' in c.upper(): return f"{vf:.2f} KM"
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(total_obras, qtd_equipes_ativas):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Modo:</b> <span style="color: #d9534f; font-weight: bold;">Lista Contínua</span></p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras}</span></p>
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

def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_lista': False, 'vrp_status_lista': "IDLE", 'vrp_state_lista': {}, 'df_routed_lista': pd.DataFrame(), 'bases_records_lista': [], 'colunas_exibir_lista': [], 'colunas_originais_lista': []})
    for k in ['bytes_zip_xl_lista', 'bytes_zip_kml_lista', 'bytes_zip_gpx_lista', 'start_time_run_lista', 'start_time_pkg_lista', 'df_unallocated_lista']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido_lista" not in st.session_state: st.session_state.roteamento_concluido_lista = False
if "vrp_status_lista" not in st.session_state: st.session_state.vrp_status_lista = "IDLE"

status_exec = st.session_state.vrp_status_lista
is_done = st.session_state.roteamento_concluido_lista
is_locked = status_exec != "IDLE" or is_done

st.markdown("<h1 class='brand-title'>📜 Lista Contínua</h1>", unsafe_allow_html=True)
st.info("💡 Distribui as obras equitativamente entre as equipes, criando uma lista de execução única e contínua.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Parâmetros de Rota", expanded=True):
        st.success("📦 **Modo Contínuo:** Todas as obras serão alocadas numa lista única, sem corte de dias.")
        trava_global = st.number_input("Trava Total de Obras no Estado", min_value=0, value=0, step=50, disabled=is_locked)
        data_ini = st.date_input("📅 Data de Início:", value=datetime.today(), disabled=is_locked)
        
        st.markdown("---")
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=True, disabled=is_locked)
        url_osrm = st.text_input("Endpoint OSRM:", value="http://router.project-osrm.org", disabled=is_locked)
        st.markdown("---")
        vel_kmh = st.slider("Velocidade Média (km/h)", 10, 100, 30, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_lista.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_lista', b"vazio"), file_name=f"ListaContinua_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🗺️ Baixar Mapas (KML)", data=st.session_state.get('bytes_zip_kml_lista', b"vazio"), file_name=f"ListaContinua_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_lista', b"vazio"), file_name=f"ListaContinua_GPS - {d_fmt}.zip", use_container_width=True)
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed_lista.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    # JUSTIFICATIVA DOS ARQUIVOS DE CORREÇÃO
    df_c = st.session_state.get('df_correcao_lista', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção (Verifique o ZIP)</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica Oficial:</b> As obras listadas no arquivo <b>"Obras_Correcao"</b> foram bloqueadas e não roteirizadas porque apresentaram <b>coordenadas geográficas em branco, zeradas, invertidas</b> ou porque o GPS da obra apontou para um local que está <b>fora da Cerca Eletrônica de 70km</b> do município preenchido na planilha.
            </pNão fui programado para fazer essas coisas.
