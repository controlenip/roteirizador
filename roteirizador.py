import streamlit as st
import os
from modules.export_tatica import injetar_logo

st.set_page_config(page_title="Roteirizador NIP", page_icon="📍", layout="centered")
injetar_logo()

st.title("📍 Hub de Roteirização NIP")
st.markdown("Bem-vindo ao sistema inteligente de planejamento logístico da NIP.")

st.info("👈 **Selecione o motor de roteirização no menu lateral esquerdo.**")

st.markdown("---")
st.markdown("### 🗺️ Planejamento Tático (Operação)")
st.markdown("Motor focado na distribuição circular de obras para Levantadores e Equipes de Campo. Utiliza a cota diária de produção e roteiriza com base no cruzamento espacial para criar mapas diários ou semanais eficientes.")

st.markdown("### 📜 Lista Contínua")
st.markdown("Gera uma lista de execução unificada e linear com base na atribuição já existente na planilha de Obras. Ideal para quando as tarefas já possuem um fiscal/levantador atrelado e o foco é apenas traçar a ordem de visitação ideal, sem corte de dias.")

st.markdown("### 📋 Planejamento de Fiscalização")
st.markdown("Motor especial para as equipes de Fiscalização. A inteligência artificial persegue os maiores focos (bolsões) de postes no mapa e ancora o Fiscal mais próximo na região, garantindo o maior volume de produção com o menor deslocamento possível.")

st.markdown("### 🧹 Saneamento")
st.markdown("Motor super-rápido focado em operações de Saneamento de Notas e Visitas Técnicas massivas. Exige o preenchimento exato da 'Latitude Projeto' e 'Longitude Projeto' e foca em resolver a maior quantidade de protocolos no menor espaço geográfico possível.")

st.markdown("---")
st.caption("Desenvolvido para NIP v3.0 | 2026")
