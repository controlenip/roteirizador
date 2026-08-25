import streamlit as st
import pandas as pd
import io
from modules.export_tatica import injetar_logo

st.set_page_config(page_title="FAQ e Modelos", page_icon="❓", layout="wide")
injetar_logo()

st.markdown("<h1 class='brand-title'>❓ FAQ e Planilhas Modelo</h1>", unsafe_allow_html=True)
st.info("💡 Este é o manual definitivo da Inteligência Artificial do Roteirizador NIP. Baixe as planilhas padrão para garantir 100% de compatibilidade e tire todas as suas dúvidas sobre o funcionamento do motor logístico.")

st.markdown("---")

# ==========================================
# SESSÃO 1: DOWNLOAD DE MODELOS
# ==========================================
st.markdown("### 📥 1. Central de Downloads (Planilhas Padrão)")
st.markdown("Para que a IA consiga ler seus dados, processar as coordenadas e devolver as rotas sem erros, suas planilhas devem seguir as nomenclaturas de colunas exatas destes modelos abaixo.")

c1, c2 = st.columns(2)

# Modelo Base (Equipes)
df_equipes = pd.DataFrame(columns=["EQUIPE", "MUNICIPIO", "LATITUDE", "LONGITUDE", "ATIVO"])
buf_eq = io.BytesIO()
df_equipes.to_excel(buf_eq, index=False)

# Modelo Obras Táticas
df_obras = pd.DataFrame(columns=["PROTOCOLO", "MUNICIPIO", "LATITUDE", "LONGITUDE", "STATUS DA FISCALIZACAO", "TIPO NOTA", "VALOR DA OBRA", "PRIORIDADE"])
buf_obras = io.BytesIO()
df_obras.to_excel(buf_obras, index=False)

with c1:
    st.markdown("#### 👥 Equipes e Fiscais")
    st.download_button("📥 Baixar Modelo: Equipes / Fiscais", data=buf_eq.getvalue(), file_name="Modelo_Equipes.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("A coluna pode se chamar **EQUIPE**, **FISCAL** ou **LEVANTADOR**. Opcionalmente, adicione **LATITUDE** e **LONGITUDE** da residência do fiscal para rotas exatas, ou apenas preencha o **MUNICIPIO** para usar o centro da cidade como partida.")

with c2:
    st.markdown("#### 🗺️ Planejamento Tático e Lista Contínua")
    st.download_button("📥 Baixar Modelo: Obras Táticas", data=buf_obras.getvalue(), file_name="Modelo_Obras.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Colunas vitais: **PROTOCOLO**, **LATITUDE** e **LONGITUDE**. A coluna de status ajuda o sistema a ignorar obras canceladas. Para a *Lista Contínua*, adicione a coluna do Fiscal já atribuído.")

st.markdown("<br>", unsafe_allow_html=True)

c3, c4 = st.columns(2)

# Modelo Fisc
df_fisc = pd.DataFrame(columns=["PROTOCOLO", "MUNICIPIO", "LATITUDE", "LONGITUDE", "QTD PREVISTA DE POSTES", "STATUS DA FISCALIZACAO", "TIPO DE PROJETO"])
buf_fisc = io.BytesIO()
df_fisc.to_excel(buf_fisc, index=False)

# Modelo Saneamento
cols_san = ['NOTA', 'STATUS CLIENTE', 'NOME', 'TIPO DEMANDA', 'MUNICIPIO', 'ENDERECO', 'BAIRRO', 'PONTO REFERENCIA', 'COMPLEMENTO', 'LATITUDE PROJETO', 'LONGITUDE PROJETO', 'CLASSIFICACAO AREA', 'TEL FIXO', 'TEL MOVEL', 'GRUPO TENSAO']
df_san = pd.DataFrame(columns=cols_san)
buf_san = io.BytesIO()
df_san.to_excel(buf_san, index=False)

with c3:
    st.markdown("#### 📋 Fiscalização")
    st.download_button("📥 Baixar Modelo: Fiscalização", data=buf_fisc.getvalue(), file_name="Modelo_Fiscalizacao.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Além das coordenadas, **exige obrigatoriamente** a coluna 'QTD PREVISTA DE POSTES'. É através desse número que a IA encontra os 'bolsões' e baliza a carga de trabalho.")

with c4:
    st.markdown("#### 🧹 Saneamento")
    st.download_button("📥 Baixar Modelo: Saneamento", data=buf_san.getvalue(), file_name="Modelo_Saneamento.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    st.caption("Formato rigoroso NIP. Exige as colunas 'LATITUDE PROJETO' e 'LONGITUDE PROJETO', além dos dados cadastrais completos do cliente. Otimizado para altas cotas de produção (ex: 25+/dia).")

st.markdown("---")

# ==========================================
# SESSÃO 2: FAQ DETALHADO E PROFUNDO
# ==========================================
st.markdown("### 🧠 2. Entendendo a Inteligência Artificial (FAQ)")

with st.expander("🛠️ 1. O que é a planilha de 'Obras_Correcao' e por que minhas obras foram rejeitadas?"):
    st.markdown("""
    O Roteirizador possui uma trava de segurança rigorosa chamada **Filtro Geográfico e de Qualidade**. 
    Antes de criar o roteiro, a IA varre 100% das obras da sua planilha e expulsa sumariamente qualquer obra que possa corromper o cálculo matemático da rota ou o visual do mapa.

    **Sua obra vai cair no arquivo 'Obras_Correcao' se tiver qualquer um destes 4 erros fatais:**
    
    1. **Coordenadas Zeradas ou em Branco:** Se a Latitude ou Longitude estiver vazia ou for `0.0`.
    2. **Coordenadas Positivas (Fora do Brasil):** As coordenadas do Brasil são sempre negativas (ex: Latitude -5.0, Longitude -43.0). Se a planilha vier com valores positivos, o GPS tentará criar uma rota para a Europa ou Ásia, quebrando o sistema.
    3. **Coordenadas Invertidas:** É muito comum o humano digitar a Longitude na coluna de Latitude e vice-versa. A IA percebe que o cruzamento matemático está errado e isola a nota.
    4. **Fuga da Cerca Eletrônica (Geofencing 70km):** Nos módulos Tático e Lista Contínua, se a planilha diz que a obra é na cidade X, a IA pega as coordenadas exatas do centro da cidade X (via satélite do IBGE) e cria um raio de 70km. Se o GPS da obra apontar para um local a mais de 70km de distância (ex: Obra cadastrada em São Luís, mas o GPS aponta para Balsas), ela é sumariamente expulsa para evitar que o técnico dirija centenas de quilômetros por um erro de digitação.
    
    **Como resolver?** Abra o arquivo ZIP baixado, procure a planilha de Correção, corrija as coordenadas destas obras e suba novamente no sistema.
    """)

with st.expander("⚖️ 2. Qual a diferença entre Atribuição 'Por Proximidade' e 'Por Município Rígido'?"):
    st.markdown("""
    Esse é o coração logístico da distribuição de tarefas entre diferentes bases/equipes.
    
    *   **📍 Por Proximidade Espacial (Recomendado):**
        Nesta opção, a IA ignora completamente os limites e fronteiras desenhados no mapa do IBGE. O nome da cidade na planilha da equipe serve **apenas para a IA saber o endereço da casa (base) do fiscal**. 
        **Exemplo:** O Fiscal 'João' mora em *São Luís* e o Fiscal 'Maria' mora em *São José de Ribamar*. Uma nota cai no sistema com o município escrito "São José de Ribamar", porém as coordenadas geográficas mostram que a nota fica literalmente na divisa, a 2km da casa do João e a 15km da casa da Maria.
        **O que a IA faz?** Entrega a nota para o João! Ela prioriza economia extrema de combustível e tempo, foda-se o município escrito no papel.
        
    *   **🔒 Por Município Rígido:**
        Nesta opção, a IA constrói um "muro de concreto" ao redor de cada equipe. O Fiscal só pode atuar **dentro das cidades que foram vinculadas a ele na planilha**.
        Se uma nota for gerada para a cidade de *Açailândia*, a IA **só** vai procurar equipes que tenham *Açailândia* escrito na sua lista de municípios. Se não houver ninguém cadastrado para Açailândia, a obra cai como "Não Alocada", mesmo que exista um fiscal numa cidade a 5km de distância dali. Ideal para contratos amarrados regionalmente.
    """)

with st.expander("🔄 3. Qual a diferença entre 'Lógica Padrão' e 'Varredura Reversa'?"):
    st.markdown("""
    Isso altera a forma como o algoritmo do **TSP (Caixeiro Viajante)** se comporta ao costurar a sua rota na rua.

    *   **📍 Lógica Padrão:** 
        A IA sai do ponto inicial (A casa do Fiscal ou a obra mais pesada) e vai "comendo pelas beiradas". Ela vai pegando sempre a nota mais próxima da posição atual. Isso cria rotas circulares muito eficientes e de expansão gradual. Ideal para quando você tem tempo sobrando e quer fechar uma área completa.
        
    *   **🎯 Varredura Reversa:** 
        O "Efeito Bumerangue". A IA olha para o mapa de obras daquele fiscal, descobre qual é a obra **mais distante fisicamente** da base dele (o extremo do território) e "chuta" o fiscal direto para lá como a primeira tarefa do dia. A partir dessa ponta mais distante, o algoritmo vem agrupando as notas voltando em direção à base (para casa). 
        Ideal para limpar periferias isoladas ou garantir que a equipe não fique "presa" no fim do expediente longe de casa.
    """)

with st.expander("⚙️ 4. Quais são as diferenças matemáticas entre os 4 Motores (Páginas)?"):
    st.markdown("""
    O sistema possui 4 cérebros diferentes, cada um treinado para uma especialidade:
    
    1. **🗺️ Planejamento Tático:** Especialista em **distribuição de cotas**. Você diz que a equipe tem que fazer "6 obras por dia", e a IA agrupa as notas criando dias exatos (Segunda, Terça, Quarta...) com 6 obras cada, garantindo que elas fiquem grudadas umas nas outras. Se passarem das cotas diárias estipuladas, as notas extras ficam de fora.
    
    2. **📜 Lista Contínua:** Especialista em **velocidade e absorção massiva**. Ignora limites de "obras por dia". Ele pega um volume gigantesco (ex: 500 notas atreladas a um levantador) e traça a linha perfeita que liga todas essas notas. Ótimo para enviar listas gigantescas de uma vez e deixar a equipe matar ao longo do mês.
    
    3. **📋 Fiscalização:** O motor de **"Bolsões de Postes"**. Em vez de olhar apenas para distância, a IA da fiscalização olha para a coluna `QTD PREVISTA DE POSTES`. Ela caça no mapa onde está a obra mais "gorda" (ex: 200 postes) e manda o fiscal direto pra lá primeiro, pois é lá que está a maior receita/importância. Depois de ancorar no bolsão, ela puxa as obras miúdas em volta.
    
    4. **🧹 Saneamento:** O motor de **Tiro Rápido**. Configurado para ignorar bloqueios complexos e rodar em segundos. Projetado para cotas altas (ex: 25 a 40 notas por dia). Ele absorve planilhas com formato duro (exigindo colunas específicas como 'LATITUDE PROJETO') e monta o mapa logístico focado puramente em proximidade geométrica.
    """)

with st.expander("🏢 5. O que significa o 'Raio Super Ponto'?"):
    st.markdown("""
    O **Super Ponto** é uma tecnologia de *clustering* (agrupamento espacial). 
    
    Imagine que existem 5 notas separadas na planilha, mas todas caem no mesmo quarteirão ou no mesmo prédio (um condomínio com vários medidores, por exemplo). Se o sistema tratasse elas individualmente, ia calcular tempo de deslocamento entre elas e ocupar a grade da equipe com "viagens fantasmas".
    
    O Raio Super Ponto (que pode ser regulado de 10 a 500 metros) junta essas notas sobrepostas, transforma elas em uma "entidade única" (um Super Ponto), atribui a mesma equipe para todas elas, e zera a distância de deslocamento entre elas. Isso economiza cotas diárias e deixa a rota realística.
    """)

with st.expander("🛣️ 6. Como as distâncias são calculadas? (Linha Reta vs Traçado de Ruas)"):
    st.markdown("""
    O sistema possui duas marchas de cálculo:
    
    1. **Fórmula de Haversine (Linha Reta):** É o cálculo base e ultrarrápido do sistema. Ele mede a distância esférica do planeta Terra entre dois pontos, multiplicada por uma constante de desvio padrão (1.3x) para simular esquinas.
    
    2. **Integração OSRM (Traçado de Ruas Real):** Se a caixinha "Traçado de Ruas Real" estiver marcada na barra lateral, a IA vai acessar servidores de tráfego open-source para desenhar a geometria exata da rua (respeitando mão dupla, rotatórias e avenidas). 
    *Atenção:* Ligar essa opção aumenta a precisão e deixa o mapa com a linha desenhada por cima da rodovia (parecido com Waze ou Google Maps), mas **demora muito mais tempo** para carregar a rota.
    
    Por padrão, o motor assume uma **velocidade média de deslocamento de 30 km/h** em áreas urbanas/rurais para calcular a hora que o fiscal entra e sai da nota.
    """)

st.markdown("<br><center><p style='color: #888;'>Manual de Operações Logísticas e Georreferenciamento | Desenvolvido para Roteirizador NIP v3.0</p></center>", unsafe_allow_html=True)
