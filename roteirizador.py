import streamlit as st
from modules.export_tatica import injetar_logo

st.set_page_config(
    page_title="Hub de Roteirização NIP",
    page_icon="📍",
    layout="wide"
)

injetar_logo()

# ==============================================================
# ESTILO VISUAL DO HUB
# ==============================================================
st.markdown(
    """
    <style>
        .block-container {
            max-width: 1450px;
            padding-top: 2.2rem;
            padding-bottom: 3rem;
        }

        .hub-hero {
            background: linear-gradient(135deg, rgba(13,37,108,0.08), rgba(85,185,41,0.08));
            border: 1px solid rgba(13,37,108,0.12);
            border-radius: 18px;
            padding: 28px 30px;
            margin-bottom: 24px;
        }

        .hub-title {
            font-size: 2.25rem;
            font-weight: 800;
            color: #0D256C;
            margin: 0;
            line-height: 1.15;
        }

        .hub-subtitle {
            font-size: 1rem;
            color: #4b5563;
            margin-top: 8px;
            margin-bottom: 0;
            line-height: 1.55;
        }

        .status-row {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 18px;
        }

        .status-pill {
            display: inline-block;
            background: #ffffff;
            border: 1px solid #dbe3ef;
            color: #334155;
            border-radius: 999px;
            padding: 7px 12px;
            font-size: 0.82rem;
            font-weight: 600;
        }

        .module-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            padding: 22px;
            min-height: 245px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
            transition: transform .15s ease, box-shadow .15s ease;
            margin-bottom: 16px;
        }

        .module-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
        }

        .module-icon {
            font-size: 1.9rem;
            margin-bottom: 8px;
        }

        .module-title {
            font-size: 1.20rem;
            font-weight: 800;
            color: #111827;
            margin-bottom: 8px;
        }

        .module-desc {
            font-size: 0.92rem;
            color: #4b5563;
            line-height: 1.55;
            min-height: 88px;
        }

        .module-meta {
            margin-top: 14px;
            padding-top: 12px;
            border-top: 1px solid #eef2f7;
            font-size: 0.80rem;
            color: #64748b;
            line-height: 1.45;
        }

        .section-title {
            font-size: 1.45rem;
            font-weight: 800;
            color: #0f172a;
            margin-top: 12px;
            margin-bottom: 6px;
        }

        .section-subtitle {
            color: #64748b;
            font-size: 0.92rem;
            margin-bottom: 16px;
        }

        .flow-box {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            padding: 18px;
            text-align: center;
            min-height: 118px;
        }

        .flow-num {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            background: #0D256C;
            color: white;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .flow-title {
            font-weight: 800;
            color: #111827;
            margin-bottom: 4px;
        }

        .flow-text {
            font-size: 0.80rem;
            color: #64748b;
            line-height: 1.4;
        }

        .footer-nip {
            text-align: center;
            color: #94a3b8;
            font-size: 0.80rem;
            padding-top: 20px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================
# CABEÇALHO
# ==============================================================
st.markdown(
    """
    <div class="hub-hero">
        <div class="hub-title">📍 Hub de Roteirização NIP</div>
        <div class="hub-subtitle">
            Plataforma de planejamento operacional para distribuição de obras, roteirização de equipes,
            fiscalização, saneamento e auditoria cruzada de bases.
        </div>
        <div class="status-row">
            <span class="status-pill">🗺️ Planejamento Operacional</span>
            <span class="status-pill">📋 Fiscalização</span>
            <span class="status-pill">🧹 Saneamento</span>
            <span class="status-pill">🔍 Auditoria</span>
            <span class="status-pill">📦 Exportações Operacionais</span>
            <span class="status-pill">❓ FAQ e Modelos</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.info(
    "👈 **Escolha o módulo no menu lateral esquerdo.** "
    "Se for a primeira utilização, abra **FAQ e Modelos** para baixar a planilha correta antes de iniciar."
)

# ==============================================================
# MÓDULOS
# ==============================================================
st.markdown('<div class="section-title">🚀 Módulos do sistema</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Cada motor foi desenhado para um tipo diferente de operação. Use o módulo que corresponde ao seu fluxo.</div>',
    unsafe_allow_html=True,
)

cards = [
    {
        "icon": "🗺️",
        "title": "Planejamento Tático",
        "desc": (
            "Distribui as demandas entre levantadores/equipes, respeitando capacidade, período de trabalho, "
            "prioridades e regras de atribuição. Indicado quando ainda é necessário decidir quem receberá cada obra."
        ),
        "meta": "Entrada principal: BASE_LEVANTAMENTO_ATUALIZADA + LEVANTADORES_PRINCIPAIS"
    },
    {
        "icon": "📜",
        "title": "Lista Contínua",
        "desc": (
            "Organiza a melhor sequência de visitação quando a própria planilha já informa o levantador ou fiscal responsável. "
            "Não redistribui as obras entre equipes: ordena o roteiro de cada responsável."
        ),
        "meta": "Entrada principal: BASE_LISTA_CONTINUA_LEVANTAMENTO"
    },
    {
        "icon": "📋",
        "title": "Planejamento de Fiscalização",
        "desc": (
            "Motor específico para fiscalização de obras e postes. Pode trabalhar em modo Tático ou Contínuo, "
            "considerando a concentração de postes, proximidade, capacidade e sentido operacional da rota."
        ),
        "meta": "Destaque: Regra do Bolsão + indicadores de postes e produtividade"
    },
    {
        "icon": "🧹",
        "title": "Saneamento",
        "desc": (
            "Planeja grandes volumes de demandas com controle de cota, jornada, tempo médio por atendimento, "
            "intervalo de almoço, prioridades, Super Pontos e possibilidade de operação contínua."
        ),
        "meta": "Destaque: planejamento de alta produtividade com controle de jornada"
    },
    {
        "icon": "🔍",
        "title": "Análise Cruzada e Auditoria",
        "desc": (
            "Cruza bases de Saneamento e Levantamento para identificar duplicidades, notas próximas, divergências, "
            "situação SAP/SISCO/LIST e colaboradores mais próximos das demandas."
        ),
        "meta": "Destaque: auditoria de fluxo, proximidade, duplicidade e equipes próximas"
    },
    {
        "icon": "❓",
        "title": "FAQ e Modelos",
        "desc": (
            "Central de orientação do sistema. Reúne as planilhas oficiais, explicações das regras, motivos de rejeição, "
            "não alocação, OSRM, Super Pontos, exportações e diagnóstico das execuções."
        ),
        "meta": "Recomendado para novos usuários e para conferir o padrão correto dos arquivos"
    },
]

for inicio in range(0, len(cards), 3):
    cols = st.columns(3, gap="large")
    for coluna, card in zip(cols, cards[inicio:inicio + 3]):
        with coluna:
            st.markdown(
                f"""
                <div class="module-card">
                    <div class="module-icon">{card['icon']}</div>
                    <div class="module-title">{card['title']}</div>
                    <div class="module-desc">{card['desc']}</div>
                    <div class="module-meta"><b>Como usar:</b> {card['meta']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

# ==============================================================
# QUAL MÓDULO USAR
# ==============================================================
st.markdown("---")
st.markdown('<div class="section-title">🧭 Qual módulo devo usar?</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Use este quadro como decisão rápida antes de iniciar uma execução.</div>',
    unsafe_allow_html=True,
)

matriz = {
    "Situação": [
        "Tenho uma base de obras e preciso distribuir entre levantadores/equipes",
        "As obras já possuem levantador/fiscal definido",
        "Preciso planejar fiscalização considerando quantidade de postes",
        "Preciso organizar muitas demandas de saneamento por equipe e jornada",
        "Quero comparar Saneamento x Levantamento e auditar status/proximidade",
        "Não sei qual arquivo usar ou quero entender uma regra do sistema",
    ],
    "Módulo recomendado": [
        "🗺️ Planejamento Tático",
        "📜 Lista Contínua",
        "📋 Planejamento de Fiscalização",
        "🧹 Saneamento",
        "🔍 Análise Cruzada e Auditoria",
        "❓ FAQ e Modelos",
    ],
}

st.table(matriz)

# ==============================================================
# FLUXO RECOMENDADO
# ==============================================================
st.markdown("---")
st.markdown('<div class="section-title">🔄 Fluxo recomendado de utilização</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Seguir esta sequência reduz erros de entrada e facilita a conferência da roteirização.</div>',
    unsafe_allow_html=True,
)

passos = [
    ("1", "Baixe o modelo", "Use sempre a planilha oficial correspondente ao módulo."),
    ("2", "Prepare os dados", "Preencha responsáveis, municípios, coordenadas e demais campos necessários."),
    ("3", "Escolha o motor", "Abra o módulo correto pelo menu lateral."),
    ("4", "Configure", "Defina cotas, períodos, filtros, atribuição e parâmetros de rota."),
    ("5", "Execute e confira", "Valide rejeitadas, não alocadas, mapa, métricas e status de rota."),
    ("6", "Exporte", "Gere as planilhas e demais arquivos operacionais disponíveis no módulo."),
]

cols_fluxo = st.columns(6, gap="small")
for col, (numero, titulo, texto) in zip(cols_fluxo, passos):
    with col:
        st.markdown(
            f"""
            <div class="flow-box">
                <div class="flow-num">{numero}</div>
                <div class="flow-title">{titulo}</div>
                <div class="flow-text">{texto}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ==============================================================
# ORIENTAÇÕES IMPORTANTES
# ==============================================================
st.markdown("---")
st.markdown('<div class="section-title">⚠️ Antes de iniciar</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2, gap="large")
with c1:
    st.success(
        "**Use os modelos oficiais.**\n\n"
        "No Planejamento Tático utilize **BASE_LEVANTAMENTO_ATUALIZADA** junto com "
        "**LEVANTADORES_PRINCIPAIS**. Na Lista Contínua utilize **BASE_LISTA_CONTINUA_LEVANTAMENTO**."
    )

with c2:
    st.warning(
        "**Confira coordenadas e responsáveis.**\n\n"
        "Coordenadas inválidas, responsáveis ausentes, filtros restritivos ou capacidade insuficiente podem "
        "gerar demandas rejeitadas ou não alocadas. Consulte o FAQ quando houver divergência."
    )

# ==============================================================
# RODAPÉ
# ==============================================================
st.markdown(
    """
    <div class="footer-nip">
        Roteirizador NIP • Planejamento, Fiscalização, Saneamento e Auditoria • v3.0 | 2026
    </div>
    """,
    unsafe_allow_html=True,
)
