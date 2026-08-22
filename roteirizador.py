import streamlit as st
from modules.export_tatica import injetar_logo

st.set_page_config(
    page_title="Roteirizador NIP v3.0",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Aciona a Logo acima do Menu
injetar_logo()

def main():
    with st.sidebar:
        st.success("Selecione um dos Módulos Acima 👆")

    st.markdown("<h1 class='brand-title' style='text-align: center; margin-bottom: 30px;'>Bem-vindo à Plataforma Roteirizadora NIP v3.0 ⚡</h1>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style="background-color: #f8f9fa; padding: 30px; border-radius: 10px; border-left: 5px solid #0D256C; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h3 style="color: #0D256C; margin-top: 0;">Iniciando o Sistema</h3>
        <p style="font-size: 16px;">O sistema agora é dividido em módulos autônomos para garantir estabilidade, segurança e zero travamentos.</p>
        <p style="font-size: 16px;">👉 <b>Utilize o menu lateral esquerdo para navegar entre as operações:</b></p>
        <ul style="font-size: 16px; line-height: 1.8;">
            <li><b>🎯 Modo Tático:</b> Planejamento Inteligente com IA e agrupamento otimizado.</li>
            <li><b>♾️ Lista Contínua:</b> Roteamento sequencial direto, focado em alta velocidade.</li>
            <li><b>📋 Fiscalização:</b> Motor exclusivo para Volumetria de Postes e Cerca Eletrônica (Geofencing).</li>
            <li><b>📖 Ajuda e Modelos:</b> Download de planilhas de exemplo e manual de regras do sistema.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
