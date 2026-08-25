import streamlit as st
import pandas as pd
import io
from modules.export_tatica import injetar_logo

st.set_page_config(page_title="FAQ e Modelos", page_icon="❓", layout="centered")
injetar_logo()

st.title("❓ FAQ e Planilhas Modelo")

st.markdown("---")

st.markdown("### 📥 Baixar Planilhas Padrão")
st.markdown("Utilize estes modelos para garantir que o roteirizador leia seus dados corretamente. As colunas devem estar preenchidas com as formatações indicadas.")

c1, c2 = st.columns(2)

# Modelo Base (Equipes e Tático)
df_equipes = pd.DataFrame(columns=["EQUIPE", "MUNICIPIO", "LATITUDE", "LONGITUDE", "ATIVO"])
buf_eq = io.BytesIO()
df_equipes.to_excel(buf_eq, index=False)

df_obras = pd.DataFrame(columns=["PROTOCOLO", "MUNICIPIO", "LATITUDE", "LONGITUDE", "STATUS DA FISCALIZACAO", "TIPO NOTA"])
buf_obras = io.BytesIO()
df_obras.to_excel(buf_obras, index=False)

with c1:
    st.download_button("👥 Planilha Modelo: Equipes / Fiscais", data=buf_eq.getvalue(), file_name="Modelo_Equipes.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Colunas chave: EQUIPE/FISCAL, LATITUDE e LONGITUDE (ou MUNICIPIO).")

with c2:
    st.download_button("🏗️ Planilha Modelo: Obras Táticas", data=buf_obras.getvalue(), file_name="Modelo_Obras.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Colunas chave: PROTOCOLO, LATITUDE e LONGITUDE.")

st.markdown("<br>", unsafe_allow_html=True)

c3, c4 = st.columns(2)

# Modelo Fisc e Saneamento
df_fisc = pd.DataFrame(columns=["PROTOCOLO", "MUNICIPIO", "LATITUDE", "LONGITUDE", "QTD PREVISTA DE POSTES", "STATUS DA FISCALIZACAO"])
buf_fisc = io.BytesIO()
df_fisc.to_excel(buf_fisc, index=False)

df_san = pd.DataFrame(columns=["NOTA", "STATUS CLIENTE", "NOME", "TIPO DEMANDA", "MUNICIPIO", "ENDERECO", "BAIRRO", "PONTO REFERENCIA", "COMPLEMENTO", "LATITUDE PROJETO", "LONGITUDE PROJETO", "CLASSIFICACAO AREA", "TEL FIXO", "TEL MOVEL", "GRUPO TENSAO"])
buf_san = io.BytesIO()
df_san.to_excel(buf_san, index=False)

with c3:
    st.download_button("📋 Planilha Modelo: Fiscalização", data=buf_fisc.getvalue(), file_name="Modelo_Fiscalizacao.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Exige a coluna 'QTD PREVISTA DE POSTES' para ancoragem dos bolsões.")

with c4:
    st.download_button("🧹 Planilha Modelo: Saneamento", data=buf_san.getvalue(), file_name="Modelo_Saneamento.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Exige LATITUDE PROJETO e LONGITUDE PROJETO. Possui mais de 10 colunas obrigatórias.")

st.markdown("---")

st.markdown("### 💡 Perguntas Frequentes (FAQ)")

with st.expander("Por que algumas obras caem na planilha de 'Obras_Correcao'?"):
    st.markdown("""
    O sistema possui uma trava de segurança chamada **Filtro Geográfico**. Toda vez que uma obra apresentar:
    1. **Latitude ou Longitude em branco ou igual a Zero.**
    2. **Coordenadas Invertidas** (A Latitude digitada no lugar da Longitude).
    3. **Coordenadas Positivas** (No Maranhão/Brasil as coordenadas devem ser negativas).
    
    A inteligência artificial isola essas obras na planilha de Correção e não as coloca no mapa KML para evitar que um erro de digitação no GPS "puxe" uma equipe do Maranhão até a África ou Ásia no Google Earth, o que estragaria toda a roteirização.
    """)

with st.expander("Qual a diferença entre Lógica Padrão e Varredura Reversa?"):
    st.markdown("""
    - **Lógica Padrão:** O motor sai do ponto inicial (A casa do fiscal ou a obra central) e vai "costurando" os pontos de perto até o mais longe.
    - **Varredura Reversa:** O motor identifica qual é a obra mais distante na carga do fiscal (o extremo geográfico) e joga ele para lá primeiro. A partir daquele ponto distante, ele vem varrendo as notas voltando em direção ao centro. Ideal para quando você quer "limpar a periferia" primeiro.
    """)

with st.expander("O que significa a Atribuição Por Proximidade?"):
    st.markdown("""
    Na **Atribuição por Proximidade**, a coluna de município na planilha de Equipes/Fiscais serve **apenas para a IA saber onde o funcionário mora**. Na hora de distribuir as obras, o robô derruba os muros municipais. Se uma obra de São José de Ribamar ficar na divisa com São Luís, e o Fiscal de São Luís estiver mais perto, a obra vai para o Fiscal de São Luís para economizar tempo e combustível.
    Já na **Atribuição por Município Rígido**, a fronteira da cidade é respeitada como um muro de concreto.
    """)

with st.expander("Como funciona o motor do Saneamento?"):
    st.markdown("""
    A página de **Saneamento** possui uma arquitetura veloz voltada para volume massivo (Cota Padrão: 25 notas por dia). Ela é programada para encontrar as colunas `LATITUDE PROJETO` e `LONGITUDE PROJETO` da planilha NIP automaticamente. O foco deste motor é a eficiência e limpeza do mapa em linha reta, dispensando o cálculo de arruamento (OSRM).
    """)
