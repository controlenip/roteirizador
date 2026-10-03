import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.export_tatica import injetar_logo


st.set_page_config(page_title="FAQ e Modelos", page_icon="❓", layout="wide")
injetar_logo()

st.markdown("<h1 class='brand-title'>❓ FAQ, Manual Operacional e Planilhas Modelo</h1>", unsafe_allow_html=True)
st.info(
    "💡 Manual operacional do Roteirizador NIP. Aqui você encontra os modelos de planilha, "
    "as regras dos motores, as causas de rejeição/não alocação, a interpretação dos resultados "
    "e orientações para conferir se uma execução realmente funcionou."
)


# ==============================================================
# FUNÇÕES DE APOIO DO FAQ
# ==============================================================
def _norm(txt):
    return re.sub(r"\s+", " ", str(txt).lower()).strip()


def _to_excel_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


APP_DIR = Path(__file__).resolve().parent


def _carregar_modelo(nome_arquivo, fallback_df):
    """Carrega o modelo oficial completo quando ele estiver no projeto.
    Procura primeiro ao lado desta página e depois na pasta /modelos.
    Se o arquivo não existir no servidor, gera um modelo estrutural simples
    com as colunas essenciais para que o botão nunca fique quebrado.
    """
    candidatos = [
        APP_DIR / nome_arquivo,
        APP_DIR / "modelos" / nome_arquivo,
    ]
    for caminho in candidatos:
        if caminho.exists() and caminho.is_file():
            return caminho.read_bytes(), True
    return _to_excel_bytes(fallback_df), False


def render_faq(items, busca=""):
    termo = _norm(busca)
    exibidos = 0
    for item in items:
        alvo = _norm(item["titulo"] + " " + item["texto"] + " " + " ".join(item.get("tags", [])))
        if termo and termo not in alvo:
            continue
        exibidos += 1
        with st.expander(item["titulo"]):
            st.markdown(item["texto"])
    return exibidos


# ==============================================================
# 1. CENTRAL DE DOWNLOADS
# ==============================================================
st.markdown("---")
st.markdown("## 📥 1. Central de Downloads — Planilhas Modelo")
st.markdown(
    "Para Levantamento, os modelos oficiais abaixo substituem os antigos modelos genéricos de Equipes e Obras. "
    "O Planejamento Tático usa duas planilhas (demanda + levantadores), enquanto a Lista Contínua usa uma única base já atribuída."
)

# Modelos oficiais de Levantamento definidos para a operação.
# Quando os arquivos completos estiverem no repositório, o FAQ entrega exatamente esses arquivos.
cols_base_levantamento = [
    "ÁREA DE ATUAÇÃO (LEVANTADOR)", "LEVANTADOR", "DATA DESPACHO", "ID SISCO", "PROTOCOLO",
    "TIPO NOTA", "DATA ABERTURA", "PRIORIDADE", "STATUS SAP", "STATUS SISCO", "STATUS LIST",
    "ORÇAMENTO MODULAR (NÃO ENVIAR A CAMPO)", "PLA ALVOS LEVANTAMENTOS (NÃO ENVIAR A CAMPO)",
    "FASE", "PAT", "REGIONAL", "MUNICIPIO", "DISTANCIA BT", "DISTANCIA MT", "DISTANCIA TRAFO",
    "POSTE PREVISTO BT", "POSTE PREVISTO MT", "NOME", "CONTA CONTRATO", "INSTALAÇÃO", "ENDEREÇO",
    "LOCALIDADE", "LATITUDE", "LONGITUDE", "INFORMAÇÕES", "INFORMAÇÕES EXTRAS", "INCLUIR_DASH",
    "POSTES PREVISTOS TOTAL", "DIAS ABERTURA-DESPACHO", "FILA OPERACIONAL", "ETAPA OPERACIONAL",
    "DIAS EM LEVANTAMENTO", "FAIXA AGING", "SEM LEVANTADOR ATIVO", "PRIORIDADE FORA LIST",
    "DIVERGENCIA SISCO-LIST", "AÇÃO NECESSÁRIA", "SEVERIDADE", "INCLUIR_DASH_SEM_DATA", "PRONTO_PARA_LIST"
]
df_base_levantamento = pd.DataFrame(columns=cols_base_levantamento)

cols_levantadores_principais = [
    "MunicIpio", "Estado", "Levantador", "Regional", "Longitude", "Latitude", "Equipe"
]
df_levantadores_principais = pd.DataFrame(columns=cols_levantadores_principais)

cols_lista_continua_levantamento = [
    "LEVANTADOR", "DATA DESPACHO", "ID SISCO", "PROTOCOLO", "TIPO NOTA", "DATA ABERTURA",
    "PRIORIDADE", "STATUS SAP", "STATUS SISCO", "STATUS LIST", "FASE", "PAT", "REGIONAL",
    "MUNICIPIO", "DISTANCIA BT", "DISTANCIA MT", "DISTANCIA TRAFO", "POSTE PREVISTO BT",
    "POSTE PREVISTO MT", "NOME", "CONTA CONTRATO", "INSTALAÇÃO", "ENDEREÇO", "LOCALIDADE",
    "LATITUDE", "LONGITUDE", "INFORMAÇÕES"
]
df_lista_continua_levantamento = pd.DataFrame(columns=cols_lista_continua_levantamento)

cols_fisc = [
    "NOTA", "VALOR DA OBRA", "QTD PREVISTA DE POSTES", "REGIONAL", "MUNICIPIO",
    "LATITUDE", "LONGITUDE", "ZONA", "STATUS DA FISCALIZACAO", "FISCAL", "OBSERVACAO"
]
df_fisc = pd.DataFrame(columns=cols_fisc)

cols_san = [
    "NOTA", "STATUS CLIENTE", "NOME", "TIPO DEMANDA", "MUNICIPIO", "ENDERECO", "BAIRRO",
    "PONTO REFERENCIA", "COMPLEMENTO", "LATITUDE PROJETO", "LONGITUDE PROJETO",
    "CLASSIFICACAO AREA", "TEL FIXO", "TEL MOVEL", "GRUPO TENSAO"
]
df_san = pd.DataFrame(columns=cols_san)

cols_lev_analise = [
    "NOTA", "MUNICIPIO", "LATITUDE", "LONGITUDE", "STATUS SAP", "STATUS SISCO", "STATUS LIST",
    "LEVANTADOR", "TIPO NOTA", "PRIORIDADE"
]
df_lev_analise = pd.DataFrame(columns=cols_lev_analise)

cols_localidades = ["NOME_COLAB", "LAT_LOC", "LON_LOC", "MUNICIPIO"]
df_localidades = pd.DataFrame(columns=cols_localidades)

bytes_base_tatico, oficial_base_tatico = _carregar_modelo("BASE_LEVANTAMENTO_ATUALIZADA.xlsx", df_base_levantamento)
bytes_levantadores, oficial_levantadores = _carregar_modelo("LEVANTADORES_PRINCIPAIS.xlsx", df_levantadores_principais)
bytes_lista_continua, oficial_lista_continua = _carregar_modelo("BASE_LISTA_CONTINUA_LEVANTAMENTO.xlsx", df_lista_continua_levantamento)

r1c1, r1c2, r1c3 = st.columns(3)
with r1c1:
    st.markdown("#### 🗺️ Tático — Base de Levantamento")
    st.download_button(
        "📥 BASE_LEVANTAMENTO_ATUALIZADA",
        data=bytes_base_tatico,
        file_name="BASE_LEVANTAMENTO_ATUALIZADA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Planilha de demandas do Planejamento Tático. O arquivo oficial possui a aba NOTAS e também painéis/abas auxiliares de gestão."
        + ("" if oficial_base_tatico else " ⚠️ Arquivo completo não localizado no servidor; foi gerado um modelo estrutural simplificado.")
    )

with r1c2:
    st.markdown("#### 👥 Tático — Levantadores")
    st.download_button(
        "📥 LEVANTADORES_PRINCIPAIS",
        data=bytes_levantadores,
        file_name="LEVANTADORES_PRINCIPAIS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Base oficial de equipes do Planejamento Tático: município, levantador, regional, longitude, latitude e equipe."
        + ("" if oficial_levantadores else " ⚠️ Arquivo completo não localizado no servidor; foi gerado um modelo estrutural simplificado.")
    )

with r1c3:
    st.markdown("#### 📜 Lista Contínua — Levantamento")
    st.download_button(
        "📥 BASE_LISTA_CONTINUA_LEVANTAMENTO",
        data=bytes_lista_continua,
        file_name="BASE_LISTA_CONTINUA_LEVANTAMENTO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Modelo oficial da Lista Contínua. A coluna LEVANTADOR já define quem receberá cada nota antes da roteirização."
        + ("" if oficial_lista_continua else " ⚠️ Arquivo completo não localizado no servidor; foi gerado um modelo estrutural simplificado.")
    )

r2c1, r2c2, r2c3 = st.columns(3)
with r2c1:
    st.markdown("#### 📋 Fiscalização")
    st.download_button(
        "📥 Baixar Modelo: Fiscalização",
        data=_to_excel_bytes(df_fisc),
        file_name="Modelo_Fiscalizacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("A coluna QTD PREVISTA DE POSTES alimenta a Regra do Bolsão e os indicadores do módulo.")

with r2c2:
    st.markdown("#### 🧹 Saneamento")
    st.download_button(
        "📥 Baixar Modelo: Saneamento",
        data=_to_excel_bytes(df_san),
        file_name="Modelo_Saneamento.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("O motor aceita LATITUDE/LONGITUDE e prioriza LATITUDE PROJETO/LONGITUDE PROJETO quando presentes.")

with r2c3:
    st.markdown("#### 🔍 Análise Cruzada — Levantamento")
    st.download_button(
        "📥 Baixar Modelo: Levantamento",
        data=_to_excel_bytes(df_lev_analise),
        file_name="Modelo_Levantamento_Analise.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Inclui STATUS SAP, STATUS SISCO e STATUS LIST para a auditoria de fluxo da Análise Cruzada.")

r3c1, r3c2, r3c3 = st.columns(3)
with r3c1:
    st.markdown("#### 📍 Análise Cruzada — Localidades")
    st.download_button(
        "📥 Baixar Modelo: Localidades",
        data=_to_excel_bytes(df_localidades),
        file_name="Modelo_Localidades_Analise.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Use uma linha por colaborador/localidade, com nome e coordenadas válidas da referência operacional.")

with r3c2:
    st.info("**Planejamento Tático:** use sempre os dois arquivos oficiais acima: BASE_LEVANTAMENTO_ATUALIZADA + LEVANTADORES_PRINCIPAIS.")

with r3c3:
    st.info("**Lista Contínua:** use BASE_LISTA_CONTINUA_LEVANTAMENTO; não é necessário carregar uma base separada de levantadores.")


# ==============================================================
# 2. CHECKLIST RÁPIDO
# ==============================================================
st.markdown("---")
st.markdown("## ✅ 2. Como saber se a ferramenta funcionou?")

c1, c2, c3, c4 = st.columns(4)
c1.success("**1. Entrada validada**\n\nSem erro de colunas obrigatórias e sem bloqueio total por coordenadas.")
c2.success("**2. Resultado gerado**\n\nA tela exibe obras roteirizadas, equipes, KM e/ou indicadores do módulo.")
c3.success("**3. Pendências auditáveis**\n\nRejeitadas e não alocadas aparecem separadas, com motivo quando aplicável.")
c4.success("**4. Exportação coerente**\n\nExcel/KML/GPX/TXT refletem a mesma execução e os filtros selecionados.")

st.warning(
    "⚠️ Uma execução concluir sem erro não significa que 100% das entradas foram roteirizadas. "
    "Sempre confira Obras_Correcao, Não Alocadas/Excedentes, falhas de rota e filtros utilizados."
)


# ==============================================================
# 3. CONTEÚDO DO FAQ
# ==============================================================
faq_geral = [
    {
        "titulo": "🧭 1. Qual módulo devo usar?",
        "tags": ["módulos", "tático", "lista", "fiscalização", "saneamento", "análise"],
        "texto": """
**Planejamento Tático:** distribui demandas entre equipes e organiza a capacidade por dia/semana. Para o fluxo padrão de Levantamento, utilize **BASE_LEVANTAMENTO_ATUALIZADA.xlsx** como demanda e **LEVANTADORES_PRINCIPAIS.xlsx** como base dos levantadores/equipes.

**Lista Contínua:** usa o responsável já existente na própria planilha. Para Levantamento, o modelo oficial é **BASE_LISTA_CONTINUA_LEVANTAMENTO.xlsx**, que já possui a coluna **LEVANTADOR** e permite criar a sequência contínua sem uma planilha separada de equipes.

**Fiscalização:** trabalha com obras de fiscalização e utiliza a **Regra do Bolsão**, considerando a quantidade prevista de postes. Pode operar em modo Tático ou Contínuo.

**Saneamento:** possui planejamento por equipe, prioridade, jornada, almoço, tempo médio de atendimento, duplicidades e modos Padrão/Contínuo.

**Análise Cruzada:** não distribui rotas. Cruza Saneamento + Levantamento + Localidades para encontrar duplicidades, proximidades, inconsistências de fluxo e colaboradores próximos.
""",
    },
    {
        "titulo": "🔄 2. Qual é o fluxo normal de uma roteirização?",
        "tags": ["passo a passo", "fluxo"],
        "texto": """
De forma geral, o sistema segue esta sequência:

1. lê e padroniza as planilhas;
2. valida colunas obrigatórias;
3. aplica filtros de status/prioridade;
4. audita coordenadas;
5. separa registros rejeitados;
6. aplica regra de atribuição e capacidade;
7. forma Super Pontos quando aplicável;
8. ordena as rotas;
9. calcula/consulta os trechos de deslocamento;
10. apresenta indicadores, mapa e tabelas;
11. gera os pacotes de exportação.

Os detalhes variam por módulo, mas as rejeições e não alocações devem permanecer auditáveis.
""",
    },
    {
        "titulo": "🧠 3. O sistema é uma IA generativa?",
        "tags": ["ia", "algoritmo", "tsp"],
        "texto": """
O núcleo do sistema é um conjunto de **regras determinísticas de validação, geoprocessamento, agrupamento e roteirização**. Ele utiliza cálculos de distância, heurísticas e/ou solver de rota, além de OSRM quando o traçado viário real está ativado.

Por isso, o resultado deve ser interpretado como **planejamento algorítmico auditável**, e não como uma resposta de texto gerada livremente.
""",
    },
    {
        "titulo": "🔒 4. Por que alguns campos ficam bloqueados depois que eu executo?",
        "tags": ["bloqueado", "nova roteirização", "configuração"],
        "texto": """
Depois que uma execução começa ou termina, várias configurações ficam bloqueadas para impedir que o usuário altere parâmetros no meio do resultado.

Para mudar cota, período, sentido, OSRM ou outras regras, utilize **Nova Roteirização/Nova Análise** e execute novamente com a nova configuração.
""",
    },
]

faq_arquivos = [
    {
        "titulo": "📄 5. Quais colunas são realmente essenciais?",
        "tags": ["colunas", "obrigatórias", "protocolo"],
        "texto": """
No fluxo oficial de **Planejamento Tático de Levantamento**, use:

- **BASE_LEVANTAMENTO_ATUALIZADA.xlsx** para as demandas. A aba principal é **NOTAS** e contém, entre outros campos, PROTOCOLO, TIPO NOTA, PRIORIDADE, STATUS SAP, STATUS SISCO, STATUS LIST, REGIONAL, MUNICIPIO, LATITUDE e LONGITUDE.
- **LEVANTADORES_PRINCIPAIS.xlsx** para a distribuição entre equipes. O arquivo utiliza **MunicIpio, Estado, Levantador, Regional, Longitude, Latitude e Equipe**.

Na **Lista Contínua de Levantamento**, use **BASE_LISTA_CONTINUA_LEVANTAMENTO.xlsx**. Ela já contém **LEVANTADOR** junto com PROTOCOLO, status, regional, município e coordenadas; por isso não precisa de uma segunda planilha de levantadores.

A Fiscalização usa **QTD PREVISTA DE POSTES** para a lógica de bolsão.

A Análise Cruzada exige, em Saneamento e Levantamento, **NOTA, MUNICIPIO, LATITUDE e LONGITUDE**; a base de Localidades precisa fornecer **nome do colaborador e coordenadas da localidade**.
""",
    },
    {
        "titulo": "🌍 6. O que é a planilha Obras_Correcao?",
        "tags": ["Obras_Correcao", "rejeitadas", "coordenadas"],
        "texto": """
É a relação de registros que não puderam entrar normalmente na roteirização por problema de qualidade geográfica ou cadastral.

Entre os motivos tratados pelos módulos estão:

- município vazio;
- latitude/longitude inválida;
- coordenada zerada;
- coordenada positiva;
- latitude/longitude aparentemente invertidas.

No Planejamento Tático, inversões inequivocamente plausíveis dentro de uma faixa geográfica brasileira podem ser **autocorrigidas**. As demais inconsistências permanecem separadas para conferência.
""",
    },
    {
        "titulo": "🔁 7. O que acontece se latitude e longitude estiverem invertidas?",
        "tags": ["latitude", "longitude", "invertida", "autocorreção"],
        "texto": """
O comportamento depende do módulo.

No **Planejamento Tático**, o sistema tenta autocorrigir uma inversão apenas quando a troca resulta em coordenadas geograficamente plausíveis para o território brasileiro. Quando não há segurança para corrigir, o registro é rejeitado para revisão.

Em outros módulos, coordenadas invertidas podem ser tratadas diretamente como inconsistência/rejeição ou alerta. Portanto, a recomendação continua sendo corrigir a fonte de dados.
""",
    },
    {
        "titulo": "📚 8. Posso enviar vários arquivos de demanda de uma vez?",
        "tags": ["vários arquivos", "upload", "múltiplos"],
        "texto": """
Sim, vários módulos aceitam múltiplos arquivos de demanda e consolidam as linhas antes do processamento.

No Saneamento, o sistema ainda registra o arquivo de origem de cada demanda para permitir auditoria de duplicidades entre arquivos.
""",
    },
    {
        "titulo": "🧾 9. NOTA, PROTOCOLO, OS e identificadores com .0 são tratados?",
        "tags": ["nota", "protocolo", "os", ".0"],
        "texto": """
Os módulos procuram aliases conhecidos para identificar a demanda. Em rotinas de análise e saneamento, o sufixo decimal artificial criado pelo Excel, como `12345.0`, pode ser removido para preservar o identificador como `12345`.

Identificadores legítimos com parte decimal não devem ser alterados indiscriminadamente.
""",
    },
]

faq_planejamento = [
    {
        "titulo": "📆 10. O que significa Cota Diária por Equipe?",
        "tags": ["cota", "obras por dia", "capacidade"],
        "texto": """
É a quantidade de obras que o Planejamento Tático tenta reservar para cada equipe em cada dia de trabalho. No fluxo de Levantamento, a demanda vem de **BASE_LEVANTAMENTO_ATUALIZADA.xlsx** e as origens/equipes vêm de **LEVANTADORES_PRINCIPAIS.xlsx**.

A capacidade total depende também da quantidade de equipes, dos dias selecionados e da quantidade de períodos configurados. Se a demanda exceder essa capacidade, o excedente pode aparecer separado como **Fora da Capacidade de Dias/Equipes**.
""",
    },
    {
        "titulo": "🗓️ 11. Qual a diferença entre visão Dia e Semana?",
        "tags": ["dia", "semana", "período"],
        "texto": """
A visão define como o limite de planejamento será interpretado.

- **Dia:** o número de períodos representa dias de rota.
- **Semana:** o número de períodos representa semanas, combinadas com os dias úteis escolhidos.

Selecionar corretamente os dias úteis é importante para que a capacidade exibida corresponda ao planejamento real.
""",
    },
    {
        "titulo": "📍 12. Qual a diferença entre atribuição por proximidade e por município?",
        "tags": ["proximidade", "município", "atribuição"],
        "texto": """
**Por Proximidade Espacial:** o sistema compara a localização das obras com as bases/equipes e busca o responsável geograficamente mais próximo dentro das regras do módulo.

**Por Município:** a equipe só recebe demandas compatíveis com o município definido para ela. Se não houver equipe elegível para aquela área, a obra pode ficar como **Fora de Área/Sem Fiscal** ou equivalente.

A nomenclatura da opção varia entre os módulos, mas o princípio é esse.
""",
    },
    {
        "titulo": "🚨 13. Como funcionam as prioridades no Planejamento Tático?",
        "tags": ["prioridade", "tipo nota", "ASC", "CCF", "DIF", "MGD", "MTP", "SID"],
        "texto": """
Quando a planilha possui campos de prioridade/tipo, o usuário pode selecionar valores considerados urgentes. O Tático sugere, quando disponíveis, tipos como **ASC, CCF, DIF, MGD, MTP e SID**.

Em planilhas genéricas, o usuário pode mapear manualmente quais valores representam prioridade e quais status estão aptos para roteirização.
""",
    },
    {
        "titulo": "🧩 14. Para que serve o upload de Planilha Genérica no Tático?",
        "tags": ["genérica", "mapeamento", "status"],
        "texto": """
Serve para roteirizar uma base auxiliar que não segue integralmente o padrão principal do sistema.

Nesse fluxo, o usuário escolhe manualmente os valores que representam **alta prioridade** e os valores considerados **roteirizáveis/aptos**. Se nenhum status apto for escolhido, a base não deve seguir para roteirização.
""",
    },
    {
        "titulo": "⚠️ 15. Por que uma obra pode ficar Não Alocada no Tático?",
        "tags": ["não alocada", "capacidade", "sem fiscal", "área"],
        "texto": """
As duas causas operacionais mais importantes são:

- **Área/Município sem equipe compatível:** não existe responsável elegível segundo a regra de atribuição.
- **Capacidade insuficiente:** acabaram os dias/equipes disponíveis para a quantidade de obras configurada.

Essas situações são diferentes de uma rejeição geográfica. A obra pode estar perfeitamente válida, mas não caber no planejamento.
""",
    },
    {
        "titulo": "🗺️ 16. Por que o mapa do Tático não aparece automaticamente?",
        "tags": ["mapa", "desempenho", "carregar mapa"],
        "texto": """
O mapa é carregado sob demanda para acelerar a abertura do resultado e reduzir o custo de renderização de muitas rotas e marcadores.

Ative **Carregar/Exibir mapa operacional** quando precisar conferir visualmente a distribuição.
""",
    },
]

faq_rotas = [
    {
        "titulo": "🏢 17. O que é o Raio Super Ponto?",
        "tags": ["super ponto", "cluster", "raio"],
        "texto": """
Super Ponto é um agrupamento espacial de demandas muito próximas entre si. O raio define a distância máxima usada para considerar essas demandas como um mesmo agrupamento operacional.

O objetivo é evitar deslocamentos artificiais entre pontos praticamente coincidentes e melhorar a leitura do roteiro.

**Importante:** o agrupamento visual/espacial não significa que as obras deixam de existir. Alguns módulos continuam considerando a quantidade real de tarefas agrupadas para capacidade, tempo e indicadores.
""",
    },
    {
        "titulo": "🔄 18. Qual a diferença entre Lógica Padrão e Varredura Reversa?",
        "tags": ["lógica padrão", "varredura reversa", "sentido"],
        "texto": """
Nos módulos gerais, a **Lógica Padrão** tende a iniciar pela referência/base e evoluir espacialmente pela rota. A **Varredura Reversa** altera o sentido operacional e pode começar pelo extremo mais distante para depois retornar pela área.

Na **Fiscalização**, há uma regra específica: a sequência é organizada para que o **maior foco de postes fique no final na Lógica Padrão**. Na **Varredura Reversa**, a mesma lógica é percorrida no sentido oposto, fazendo o maior foco aparecer no início.
""",
    },
    {
        "titulo": "🛣️ 19. Qual a diferença entre distância estimada e Traçado de Ruas Real?",
        "tags": ["haversine", "osrm", "linha reta", "ruas"],
        "texto": """
Sem traçado real, o sistema utiliza distância geográfica entre coordenadas e, em alguns cálculos, um fator de estimativa para representar o deslocamento rodoviário.

Com **OSRM/Traçado de Ruas Real** ativado, o sistema tenta obter a geometria real de condução entre os pontos. Isso melhora a representação do percurso, mas depende de um serviço de rede e pode deixar a execução mais lenta.
""",
    },
    {
        "titulo": "📡 20. O que acontece quando o OSRM falha?",
        "tags": ["osrm", "falha", "timeout", "sem rota"],
        "texto": """
O comportamento varia por módulo, mas a falha não deve ser silenciosa.

Você pode encontrar situações como **SEM_ROTA, TIMEOUT, ERRO, ESTIMADA ou LINHA_RETA**. A Lista Contínua possui fallback para uma geometria/distância estimada quando o servidor não responde corretamente.

No Tático e na Fiscalização, trechos problemáticos podem ser contabilizados e destacados para conferência. Nunca interprete uma falha de OSRM como se fosse automaticamente uma rota viária confirmada.
""",
    },
    {
        "titulo": "✅ 21. Como saber se o OSRM realmente foi usado?",
        "tags": ["status rota", "trechos osrm", "diagnóstico"],
        "texto": """
Confira os indicadores/colunas de **STATUS_ROTA** e os resumos de arruamento.

Dependendo do módulo, o resultado informa quantos trechos foram obtidos via OSRM, quantos ficaram estimados e quantos ficaram sem rota. No Tático, o Resumo Operacional também pode registrar **Trechos OSRM** e **Trechos sem rota**.
""",
    },
]

faq_lista = [
    {
        "titulo": "📜 22. Como funciona a Lista Contínua?",
        "tags": ["lista contínua", "responsável"],
        "texto": """
A Lista Contínua não distribui a demanda entre pessoas. Ela parte da atribuição já existente na própria planilha e cria uma sequência contínua de execução para cada responsável.

Para o fluxo de Levantamento, utilize **BASE_LISTA_CONTINUA_LEVANTAMENTO.xlsx**. O modelo oficial já possui a coluna **LEVANTADOR**, além de PROTOCOLO, status, regional, município, latitude e longitude. Portanto, **não é necessário carregar LEVANTADORES_PRINCIPAIS nesse módulo**.

Em bases alternativas, o código também reconhece aliases de responsável como **FISCAL, EQUIPE, LEVANTADOR_RESPONSAVEL ou NOME_FISCAL**.
""",
    },
    {
        "titulo": "🎯 23. A Meta de Obras/Dia limita a Lista Contínua?",
        "tags": ["meta obras dia", "postes", "limite"],
        "texto": """
Não. Na Lista Contínua, o campo **Meta de Obras/Dia (Cálculo de Postes)** é utilizado como referência de produtividade/estimativa, e não como uma cota que interrompe a lista.

A lista permanece contínua para as obras elegíveis do responsável, sujeita aos filtros, trava global e validações de entrada.
""",
    },
    {
        "titulo": "🧱 24. Para que serve a Trava Total de Obras na Lista Contínua?",
        "tags": ["trava global", "lista contínua"],
        "texto": """
É um limitador global de volume processado naquela execução. Valor **0** significa que não há uma trava numérica definida pelo usuário.

Use a trava quando quiser testar uma amostra, limitar um lote operacional ou reduzir o volume da execução.
""",
    },
    {
        "titulo": "🚨 25. Como a Lista Contínua identifica prioridade?",
        "tags": ["prioridade", "tipo nota"],
        "texto": """
Se existir uma coluna **PRIORIDADE**, os registros marcados são preservados/destacados no roteiro e nas exportações. Quando existe **TIPO NOTA**, o módulo também permite selecionar tipos considerados de alta prioridade.
""",
    },
]

faq_fiscalizacao = [
    {
        "titulo": "📋 26. Qual a diferença entre Fiscalização Tática e Fiscalização Contínua?",
        "tags": ["fiscalização", "tático", "contínua"],
        "texto": """
**Fiscalização Tática:** recebe uma planilha de fiscais/bases e uma ou mais planilhas de obras. O sistema cruza as duas bases e atribui as obras de acordo com proximidade ou município.

**Fiscalização Contínua:** usa o fiscal já informado na própria planilha de obras e cria o roteiro contínuo por técnico, sem a divisão operacional normal de cotas/períodos do modo Tático.
""",
    },
    {
        "titulo": "📦 27. O que é exatamente a Regra do Bolsão?",
        "tags": ["bolsão", "postes", "fiscalização"],
        "texto": """
A Fiscalização utiliza a quantidade prevista de postes para orientar o sentido operacional dos focos de auditoria.

- **Lógica Padrão:** percorre os pontos por proximidade e **termina no maior foco de postes**.
- **Varredura Reversa:** percorre a mesma lógica no sentido oposto e **inicia pelo maior foco**.

Isso é diferente de afirmar que o maior bolsão sempre será o primeiro ponto da rota.
""",
    },
    {
        "titulo": "🧠 28. O que é o cache de rotas da Fiscalização?",
        "tags": ["cache", "osrm", "fiscalização"],
        "texto": """
O módulo mantém em memória trechos de rota OSRM já consultados durante a sessão. Isso evita repetir chamadas de rede para o mesmo par de coordenadas e pode acelerar novas etapas/execuções.

A barra lateral mostra a quantidade de trechos em cache e possui uma opção para **Limpar cache de rotas**.
""",
    },
    {
        "titulo": "📡 29. O que significa o teste de OSRM da Fiscalização?",
        "tags": ["pré-teste", "latência", "osrm"],
        "texto": """
Antes de executar o motor, a Fiscalização pode testar um trecho curto no endpoint configurado. Quando funciona, a tela informa que o OSRM está disponível e mostra a latência aproximada.

Se o teste falhar, o motor ainda pode continuar e registrar falhas por trecho. O teste é diagnóstico, não uma garantia de que todos os trajetos retornarão geometria.
""",
    },
    {
        "titulo": "📊 30. Quais resultados devo conferir na Fiscalização?",
        "tags": ["dashboard", "resumo técnico", "postes"],
        "texto": """
Confira pelo menos:

- obras roteirizadas;
- fiscais/equipes utilizados;
- KM previsto;
- total de postes;
- alertas de arruamento;
- obras retidas para correção;
- obras não alocadas/excedentes;
- resumo por técnico;
- mapa operacional quando necessário.
""",
    },
]

faq_saneamento = [
    {
        "titulo": "🧹 31. Qual a diferença entre modo Padrão e Contínuo do Saneamento?",
        "tags": ["saneamento", "padrão", "contínuo"],
        "texto": """
**Padrão:** respeita a quantidade de períodos configurada e trabalha com a regra de atribuição escolhida.

**Contínuo:** tenta roteirizar todas as tarefas elegíveis sem limitar a quantidade de períodos. Porém, **a cota diária e a jornada continuam válidas**. No modo Contínuo, a vinculação das obras fica amarrada aos municípios definidos para as equipes.
""",
    },
    {
        "titulo": "⏱️ 32. Como a jornada de trabalho interfere no Saneamento?",
        "tags": ["jornada", "horário", "almoço"],
        "texto": """
O Saneamento considera:

- horário de início da jornada;
- horário de fim;
- intervalo de almoço, quando ativado;
- tempo médio por obra;
- tempo estimado de deslocamento;
- retorno à base.

Se uma tarefa ou conjunto de tarefas não couber no tempo disponível, ela pode ser adiada ou classificada como não alocada, conforme a situação.
""",
    },
    {
        "titulo": "🕒 33. Para que serve o Tempo médio por obra?",
        "tags": ["tempo médio", "obra", "jornada"],
        "texto": """
É a estimativa de tempo de atendimento em campo usada para montar a agenda diária.

O motor soma deslocamento + atendimento + intervalo + retorno à base para decidir se a tarefa cabe na jornada configurada.
""",
    },
    {
        "titulo": "🏢 34. O Super Ponto do Saneamento consome tempo por obra agrupada?",
        "tags": ["super ponto", "tempo", "saneamento"],
        "texto": """
Existe uma configuração específica: **Super Ponto consome tempo por obra agrupada**.

Quando ativada, um Super Ponto com várias obras consome tempo de atendimento proporcional à quantidade real de tarefas agrupadas. Isso evita que um agrupamento grande pareça ocupar o mesmo tempo de uma única visita.
""",
    },
    {
        "titulo": "❌ 35. O que acontece se um Super Ponto for maior que a cota diária?",
        "tags": ["super ponto", "cota", "não alocada"],
        "texto": """
Se a quantidade real de tarefas contidas em um Super Ponto ultrapassar a cota diária da equipe, o Saneamento pode classificar o item como **SUPER_PONTO_EXCEDE_COTA_DIARIA** e deixá-lo fora do planejamento.
""",
    },
    {
        "titulo": "🔁 36. Como funciona a remoção de NOTAs duplicadas no Saneamento?",
        "tags": ["duplicidade", "deduplicar", "nota"],
        "texto": """
O sistema audita quantas vezes cada NOTA aparece e se a repetição ocorre no mesmo arquivo ou entre arquivos diferentes.

Quando **Remover ocorrências duplicadas da mesma NOTA** está ativado, uma ocorrência é preservada e as demais podem ser retiradas da roteirização com motivo **DUPLICIDADE_REMOVIDA**. A aba de duplicidades mantém a rastreabilidade.
""",
    },
    {
        "titulo": "📏 37. Para que serve a Distância máxima por proximidade no Saneamento?",
        "tags": ["distância máxima", "proximidade", "atribuição"],
        "texto": """
Quando a atribuição por proximidade está ativa, esse campo pode limitar a distância máxima entre obra e equipe elegível.

Valor **0 = ilimitado**. Se você definir um limite muito pequeno, demandas válidas podem ficar sem equipe elegível.
""",
    },
    {
        "titulo": "⚠️ 38. Quais motivos podem aparecer em Obras Não Alocadas no Saneamento?",
        "tags": ["não alocadas", "motivos", "saneamento"],
        "texto": """
Entre os motivos produzidos pelo motor estão:

- `FORA_DA_TRAVA_GLOBAL`;
- `LIMITE_DE_PERIODOS_ATINGIDO`;
- `SUPER_PONTO_EXCEDE_COTA_DIARIA`;
- `ATENDIMENTO_EXCEDE_JORNADA_DIARIA`;
- `TAREFA_NAO_CABE_NA_JORNADA`;
- `DUPLICIDADE_REMOVIDA`;
- impossibilidade de atribuição por município/proximidade.

Use a aba **Obras Não Alocadas** para diferenciar falta de capacidade de problema cadastral/geográfico.
""",
    },
]

faq_analise = [
    {
        "titulo": "🔍 39. O que a Análise Cruzada faz?",
        "tags": ["análise cruzada", "auditoria", "saneamento", "levantamento"],
        "texto": """
Ela cruza bases de **Saneamento** e **Levantamento**, compara ocorrências e proximidades geográficas, audita os status de fluxo e calcula os colaboradores/localidades mais próximos de cada obra.

Ela é uma ferramenta de **auditoria e decisão**, não um roteirizador de execução diária.
""",
    },
    {
        "titulo": "📂 40. Quais arquivos a Análise Cruzada exige?",
        "tags": ["arquivos análise", "localidades"],
        "texto": """
São três entradas:

1. **Base Saneamento**;
2. **Base Levantamento**;
3. **Base Localidades** dos colaboradores/equipes.

Saneamento e Levantamento precisam ter NOTA, município e coordenadas. Localidades precisa ter nome do colaborador e coordenadas válidas da localidade.
""",
    },
    {
        "titulo": "🚦 41. Quais STATUS LIST são considerados válidos na auditoria?",
        "tags": ["status list", "0", "levantamento"],
        "texto": """
Na versão de regras `2026.09`, a auditoria considera válidos:

- `0`;
- `EM LEVANTAMENTO`;
- `CORRECAO DE LEVANTAMENTO`.

O sistema normaliza acentos, espaços e pontuação antes da comparação.
""",
    },
    {
        "titulo": "🚦 42. Quais STATUS SISCO são considerados válidos?",
        "tags": ["status sisco", "pré análise", "liberado levantamento"],
        "texto": """
Na versão de regras `2026.09`, são aceitos:

- `0`;
- `PRE ANALISE`;
- `LIBERADO PARA LEVANTAMENTO`;
- `LIBERADO P LEVANTAMENTO`.

Formas com acentos e separadores diferentes são normalizadas antes da validação.
""",
    },
    {
        "titulo": "0️⃣ 43. O valor 0 no LIST/SISCO é o mesmo que coluna ausente?",
        "tags": ["zero", "coluna ausente", "status"],
        "texto": """
Não. Essa distinção é intencional.

`0` pode ser um **valor real e permitido**. Já uma coluna STATUS LIST ou STATUS SISCO que não existe na planilha é tratada como **coluna não localizada**, e isso pode invalidar o fluxo do registro.

O sistema não deve transformar automaticamente ausência de informação em `0`.
""",
    },
    {
        "titulo": "🧾 44. O que significam APTO, SEM REGISTRO SAP, SEM STATUS SAP e BLOQUEADO?",
        "tags": ["sap", "apto", "finl", "canc"],
        "texto": """
Na validação SAP da Análise Cruzada:

- **APTO:** há registro/status e ele não está nos bloqueios tratados;
- **SEM REGISTRO SAP:** a nota não pôde ser confirmada na base de Levantamento/SAP usada na análise;
- **SEM STATUS SAP:** o registro existe, mas o status está vazio;
- **BLOQUEADO (...):** o status contém situações como **FINL** ou **CANC**.

Falta de confirmação SAP **não é presumida como APTO**.
""",
    },
    {
        "titulo": "🎨 45. O que significam as cores da Análise Cruzada?",
        "tags": ["cores", "duplicadas", "próximas", "solitárias"],
        "texto": """
A classificação principal segue esta lógica:

- ⚫ **Notas Inválidas** — falharam na auditoria de fluxo;
- 🔴 **Notas Duplicadas** — mesma nota aparece em mais de uma ocorrência conforme a análise;
- 🟠 **Notas Próximas** — existem ocorrências geograficamente próximas dentro do raio configurado;
- 🟢 **Levantamento Solitárias** — registro de Levantamento sem proximidade/duplicidade prioritária;
- 🟣 **Saneamento Solitárias** — registro de Saneamento sem proximidade/duplicidade prioritária;
- 🔵 **Outras** — classificações residuais.
""",
    },
    {
        "titulo": "📏 46. O que significa Distância para agrupar obras na Análise?",
        "tags": ["raio", "proximidade", "cluster"],
        "texto": """
É o raio, em metros, usado para identificar obras/ocorrências próximas e formar a leitura de clusters espaciais.

Ele não altera a identidade da NOTA; apenas muda a interpretação de proximidade geográfica.
""",
    },
    {
        "titulo": "👥 47. Como funcionam as Equipes/Colaboradores mais próximos?",
        "tags": ["equipes próximas", "localidades", "distância"],
        "texto": """
Para cada coordenada analisada, o sistema calcula os colaboradores/localidades mais próximos.

Você pode configurar:

- quantidade de equipes exibidas;
- distância máxima para equipes adicionais;
- distância a partir da qual deve ser gerado um **alerta de equipe distante**.

A primeira equipe mais próxima continua sendo usada como referência mesmo quando as equipes adicionais ultrapassam o limite configurado.
""",
    },
    {
        "titulo": "🔎 48. Os filtros da Análise afetam também a exportação?",
        "tags": ["filtros", "excel", "kml", "exportação"],
        "texto": """
Sim. Depois da análise, você pode filtrar por **classificação, origem, município, NOTA e colaborador próximo**.

Os pacotes Excel/KML são regenerados quando a assinatura desses filtros muda, de forma que a exportação corresponda à visualização selecionada.
""",
    },
    {
        "titulo": "🧪 49. O que significa Testes internos de regressão: OK?",
        "tags": ["regressão", "teste interno", "diagnóstico"],
        "texto": """
A Análise Cruzada executa testes internos em regras críticas, como normalização de NOTA, validação LIST/SISCO/SAP, classificação de duplicidade, cálculo de equipes próximas e geração de KML.

Quando aparece **✅ Testes internos de regressão: OK**, esses testes básicos passaram naquela inicialização da página. Isso aumenta a confiança na integridade do núcleo, mas não substitui a conferência dos dados carregados pelo usuário.
""",
    },
]

faq_exportacao = [
    {
        "titulo": "📦 50. Qual a diferença entre Excel, TXT, KML e GPX?",
        "tags": ["excel", "txt", "kml", "gpx", "download"],
        "texto": """
**Excel:** resultado tabular, resumos, auditorias e listas operacionais.

**TXT:** relatório textual/manifesto/configuração em módulos que oferecem esse pacote.

**KML:** mapa para Google Earth e ferramentas compatíveis.

**GPX:** formato de intercâmbio para aplicativos/dispositivos de navegação compatíveis.

Os arquivos normalmente são empacotados em ZIP e podem ser separados por equipe/técnico.
""",
    },
    {
        "titulo": "🧾 51. O que é o manifesto/configuração da execução?",
        "tags": ["manifesto", "configuração", "rastreabilidade", "id"],
        "texto": """
Alguns módulos registram junto da exportação um resumo das configurações efetivamente utilizadas: ID da execução, versão das regras, arquivos, cota, períodos, OSRM, raio, duplicidades, quantidades rejeitadas/não alocadas e outros parâmetros.

Esse arquivo é importante para reproduzir ou auditar uma execução posteriormente.
""",
    },
    {
        "titulo": "🆔 52. Para que serve o ID da execução/análise?",
        "tags": ["id execução", "versão regras"],
        "texto": """
O ID diferencia uma execução das demais e permite relacionar tela, planilhas, mapas e arquivos de configuração ao mesmo processamento.

Na Análise Cruzada e no Saneamento, a versão das regras também é registrada para indicar qual conjunto de validações foi usado.
""",
    },
    {
        "titulo": "🧹 53. Quando devo usar Nova Roteirização/Nova Análise?",
        "tags": ["reset", "nova roteirização", "nova análise"],
        "texto": """
Use quando quiser:

- alterar parâmetros;
- substituir os arquivos de entrada;
- mudar equipes selecionadas;
- refazer filtros de entrada;
- testar outra regra de atribuição;
- trocar o sentido da rota;
- alterar cota/período/jornada.

O botão limpa o estado do resultado daquela página. Alguns caches técnicos podem ser preservados deliberadamente para desempenho, conforme o módulo.
""",
    },
]

faq_diagnostico = [
    {
        "titulo": "🚫 54. Qual a diferença entre Rejeitada, Não Alocada e Sem Rota?",
        "tags": ["rejeitada", "não alocada", "sem rota"],
        "texto": """
**Rejeitada:** falhou antes da roteirização, normalmente por problema de estrutura/cadastro/coordenadas.

**Não Alocada:** passou pela validação, mas não encontrou espaço ou responsável compatível segundo capacidade, município, proximidade, jornada ou outra regra operacional.

**Sem Rota:** a obra foi alocada, mas um trecho viário não retornou geometria válida no serviço de rotas. É um problema diferente de atribuição.
""",
    },
    {
        "titulo": "🧪 55. Como faço uma validação rápida antes de liberar o resultado para a equipe?",
        "tags": ["conferência", "checklist", "validação"],
        "texto": """
Antes de distribuir o resultado, confira:

1. quantidade de entradas x roteirizadas x rejeitadas x não alocadas;
2. equipes/fiscais realmente selecionados;
3. cota, período e dias úteis;
4. regra de atribuição;
5. prioridades;
6. quantidade de Super Pontos;
7. KM total e trechos OSRM/estimados/sem rota;
8. mapa de pelo menos algumas equipes;
9. planilha de correção;
10. arquivos exportados e ID/configuração da execução, quando disponíveis.
""",
    },
    {
        "titulo": "🐢 56. Por que algumas execuções ficam mais lentas?",
        "tags": ["lentidão", "desempenho", "osrm", "mapa"],
        "texto": """
Os maiores fatores de custo são quantidade de obras, quantidade de equipes, quantidade de trechos viários, uso de OSRM, geração de mapas e volume de exportações.

Por isso alguns módulos adotam estratégias como **mapa sob demanda, cache de rotas e geração de ZIP apenas quando o usuário solicita o download**.
""",
    },
]

TODOS = (
    faq_geral + faq_arquivos + faq_planejamento + faq_rotas + faq_lista +
    faq_fiscalizacao + faq_saneamento + faq_analise + faq_exportacao + faq_diagnostico
)

st.markdown("---")
st.markdown("## 🧠 3. FAQ Completo")

busca = st.text_input(
    "🔎 Pesquisar no FAQ",
    placeholder="Ex.: OSRM, não alocada, STATUS LIST, Super Ponto, Fiscalização...",
)

if busca:
    st.caption("Exibindo somente perguntas que contenham o termo pesquisado.")
    qtd = render_faq(TODOS, busca)
    if qtd == 0:
        st.warning("Nenhuma pergunta encontrada. Tente outro termo.")
else:
    abas = st.tabs([
        "🧭 Geral",
        "📄 Arquivos",
        "🗺️ Tático",
        "🛣️ Rotas",
        "📜 Lista",
        "📋 Fiscalização",
        "🧹 Saneamento",
        "🔍 Análise",
        "📦 Exportação",
        "🧪 Diagnóstico",
    ])

    grupos = [
        faq_geral,
        faq_arquivos,
        faq_planejamento,
        faq_rotas,
        faq_lista,
        faq_fiscalizacao,
        faq_saneamento,
        faq_analise,
        faq_exportacao,
        faq_diagnostico,
    ]

    for aba, itens in zip(abas, grupos):
        with aba:
            render_faq(itens)


st.markdown("---")
st.markdown("## 📌 4. Resumo das regras mais importantes")
resumo = pd.DataFrame([
    ["Planejamento Tático", "BASE_LEVANTAMENTO_ATUALIZADA + LEVANTADORES_PRINCIPAIS", "Cota + dias/períodos", "Proximidade ou município", "Excel/KML/GPX"],
    ["Lista Contínua", "BASE_LISTA_CONTINUA_LEVANTAMENTO", "Sequência contínua", "Responsável já atribuído", "Excel/TXT/KML/GPX"],
    ["Fiscalização", "Tático ou Contínuo", "Bolsão por postes", "Proximidade/município ou fiscal informado", "Excel/TXT/KML/GPX"],
    ["Saneamento", "Padrão ou Contínuo", "Cota + jornada + atendimento", "Proximidade/município", "Excel/KML/GPX + manifesto"],
    ["Análise Cruzada", "Auditoria", "Não se aplica", "Equipes mais próximas", "Excel/KML + configuração"],
], columns=["Módulo", "Objetivo", "Controle principal", "Atribuição", "Saídas"])
st.dataframe(resumo, use_container_width=True, hide_index=True)

st.markdown(
    "<br><center><p style='color:#888;'>Manual Operacional e FAQ | Roteirizador NIP | Atualizado para as regras e funções presentes nos módulos analisados</p></center>",
    unsafe_allow_html=True,
)
