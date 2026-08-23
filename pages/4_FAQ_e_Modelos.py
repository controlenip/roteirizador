import streamlit as st
import pandas as pd
import io
import os

# Função genérica para instanciar a Logo nos manuais também
def injetar_logo():
    if os.path.exists("LOGO_NIP.png"): st.logo("LOGO_NIP.png", icon_image=None)

st.set_page_config(page_title="FAQ e Modelos", page_icon="📖", layout="wide")
injetar_logo()

def gerar_excel_modelo(df_modelo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_modelo.to_excel(writer, index=False, sheet_name='Modelo')
    return output.getvalue()

def renderizar_faq():
    with st.sidebar:
        st.info("Aqui você encontra os Manuais e Templates oficiais do sistema.")

    st.markdown("<h1 class='brand-title' style='margin-bottom: 20px;'>📖 Central de Ajuda e Manual de Operação</h1>", unsafe_allow_html=True)
    st.markdown("Bem-vindo ao manual do **Roteirizador NIP v3.0**. Esta página detalha como o Cérebro Logístico (IA) toma decisões, os filtros ocultos e o fluxo de todos os dados do sistema.")
    st.markdown("---")
    
    st.markdown("### 1. As Três Estratégias de Despacho (Modos do Sistema)")
    c_faq1, c_faq2, c_faq3 = st.columns(3)
    with c_faq1: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #0D256C; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <h4 style='color: #0D256C; margin-top: 0;'>🗺️ 1. Planejamento Tático</h4>
            <b>A IA no Comando:</b> Você sobe a planilha de Equipes/Bases e a planilha com milhares de obras brutas. O sistema lê onde cada técnico atua, calcula todas as distâncias cruzadas e distribui as obras do zero, de forma equitativa e otimizada.<br><br>
            <b>Cota Diária e Super Pontos:</b> Você define quantas obras o técnico fará no dia/semana. A IA funde endereços idênticos num único "Super Ponto" (um prédio, condomínio) e traça a rota ideal.<br><br>
            <b>Planilha Genérica:</b> Permite acoplar planilhas extra e customizar quais colunas definem prioridade e status, mesclando dados fora do padrão.
        </div>
        """, unsafe_allow_html=True)
    with c_faq2: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #d9534f; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>
            <h4 style='color: #d9534f; margin-top: 0;'>📜 2. Lista Contínua</h4>
            <b>O Usuário no Comando:</b> A IA respeita estritamente o que você definiu. Sua planilha já deve ter a coluna <b>LEVANTADOR</b> ou <b>FISCAL</b> preenchida indicando o dono da obra.<br><br>
            <b>Carga Ilimitada:</b> A ferramenta ignora as travas de fim de expediente e desenha o caminho mais curto para conectar 100% da lista do profissional em uma varredura única.<br><br>
            <b>Dicionário de Nomes:</b> Se as suas obras estiverem vinculadas a nomes numéricos (ex: "EQUIPE 01"), você pode fazer o upload da planilha de Fiscais e a IA fará o cruzamento/PROCV mágico, substituindo tudo para o nome real nos mapas.
        </div>
        """, unsafe_allow_html=True)
    with c_faq3: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #FF9800; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <h4 style='color: #FF9800; margin-top: 0;'>📋 3. Fiscalização</h4>
            <b>A Regra do Bolsão:</b> O motor matemático garimpa as obras com a MAIOR quantidade de postes. Em seguida, ancora o Fiscal mais próximo nesse "Bolsão de Densidade" e traça a rota puxando ele para o local de maior rentabilidade.<br><br>
            <b>Termografia Visual e Relatórios:</b> No KML, os pinos ganham cores pelo volume de postes. Gera automaticamente gráficos de Rosca/Barras e Relatório Executivo em PDF.<br><br>
            <b>Cerca Eletrônica 70km:</b> Uma trava restrita deste módulo bloqueia automaticamente obras cujo GPS escape em mais de 70km da cidade base do fiscal, retendo-as numa "Planilha de Correção".
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 2. Baixar Modelos de Planilhas (Templates)")
    
    # Gerando as estruturas com base nos arquivos enviados pelo usuário
    df_equipes_tat = pd.DataFrame(columns=['MunicIpio', 'Estado', 'Levantador', 'Regional', 'Longitude', 'Latitude', 'Equipe'])
    df_equipes_tat.loc[0] = ['SÃO LUIS', 'Maranhão', 'NOME DO TECNICO', 'CENTRO', -44.3028, -2.5297, 'EQUIPE 01']

    df_obras_tat = pd.DataFrame(columns=['ID SISCO', 'PROTOCOLO', 'TIPO NOTA', 'PRIORIDADE', 'STATUS SAP', 'STATUS SISCO', 'STATUS LIST', 'FASE', 'PAT', 'Regional', 'Município', 'DISTANCIA BT', 'DISTANCIA MT', 'DISTANCIA TRAFO', 'POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'NOME', 'CONTA CONTRATO', 'INSTALAÇÃO', 'ENDEREÇO', 'LOCALIDADE', 'LATITUDE', 'LONGITUDE', 'INFORMAÇÕES EXTRAS', 'TEXTO', 'TEXTO_GERAL'])
    df_obras_tat.loc[0] = [12345, 1081945188, 'UNR', 0, 'ATIV', 'Liberado', 'Em levantamento', 'MO', 'PAT1', 'LESTE', 'SAO LUIS', 248.78, 241.99, 243.4, 6, 2, 'CLIENTE TESTE', 3019326160, 2000876176, 'RUA TESTE, 123', 'URBANA', -2.5297, -44.3028, 'INFO', 'TXT', 'TXT_G']

    df_continua = pd.DataFrame(columns=['LEVANTADOR', 'DATA DESPACHO', 'ID SISCO', 'PROTOCOLO', 'TIPO NOTA', 'PRIORIDADE', 'STATUS SAP', 'STATUS SISCO', 'STATUS LIST', 'FASE', 'PAT', 'Regional', 'Município', 'DISTANCIA BT', 'DISTANCIA MT', 'DISTANCIA TRAFO', 'POSTE PREVISTO BT', 'POSTE PREVISTO MT', 'NOME', 'CONTA CONTRATO', 'INSTALAÇÃO', 'ENDEREÇO', 'LOCALIDADE', 'LATITUDE', 'LONGITUDE', 'INFORMAÇÕES EXTRAS'])
    df_continua.loc[0] = ['NOME DO TECNICO', '17/08/2026', 12346, 1076894592, 'UNR', 0, 'ATIV', 'Pré Análise', 'Em levantamento', 'MO', 'PAT1', 'LESTE', 'CAXIAS', 285.18, 212.99, 224.66, 7, 2, 'CLIENTE TESTE', 3018299797, 2000835041, 'RUA TESTE, 123', 'RURAL', -5.0606, -43.4388, 'INFO']

    df_fiscais = pd.DataFrame(columns=['Município', 'Estado', 'Regional', 'Longitude', 'Latitude', 'FISCAL'])
    df_fiscais.loc[0] = ['SÃO LUIS', 'Maranhão', 'CENTRO', -44.3028, -2.5297, 'NOME DO FISCAL']

    df_fiscalizacao = pd.DataFrame(columns=['Nota', 'Pasta', 'Descrição', 'PEP', 'Visita', 'Ordem', 'Valor da Obra', 'Qtd prevista de postes', 'Previsão de Entrega', 'Parceiro', 'Tipo de Fiscalização', 'Tipo de Projeto', 'Contrato', 'Empresa', 'Regional', 'Município', 'Latitude', 'Longitude', 'Zona', 'Data de entrada', 'Data de despacho P/Campo', 'Postes fiscalizados', 'Status da fiscalização', 'Fiscal', 'Backoffice de fiscalização', 'Data da fiscalização', 'Data entrega SISCO', 'Observação', 'Data de medição', 'Data de faturamento'])
    df_fiscalizacao.loc[0] = [1081945188, 'PASTA1', 'DESC', 'PEP1', 1, 1, 15000.50, 45, '10/10/2026', 'PARCEIRO', 'NORMAL', 'EXTENSAO', 'CONT1', 'EMP1', 'LESTE', 'SAO LUIS', -2.5297, -44.3028, 'URBANA', '01/01/2026', '02/01/2026', 0, 'APTO PARA CAMPO', 'NOME DO FISCAL', 'BACKOFFICE', '', '', 'ATENCAO', '', '']

    col_dl1, col_dl2, col_dl3, col_dl4, col_dl5 = st.columns(5)
    with col_dl1: st.download_button("📥 Baixar Tático - Equipes", data=gerar_excel_modelo(df_equipes_tat), file_name="LEVANTADORES_PRINCIPAIS.xlsx", use_container_width=True)
    with col_dl2: st.download_button("📥 Baixar Tático - Demandas", data=gerar_excel_modelo(df_obras_tat), file_name="BASE_LEVANTAMENTO.xlsx", use_container_width=True)
    with col_dl3: st.download_button("📥 Baixar Lista Contínua", data=gerar_excel_modelo(df_continua), file_name="BASE_LISTA_CONTINUA.xlsx", use_container_width=True)
    with col_dl4: st.download_button("📥 Baixar Fiscalização - Fiscais", data=gerar_excel_modelo(df_fiscais), file_name="FISCAIS_PRINCIPAIS.xlsx", use_container_width=True)
    with col_dl5: st.download_button("📥 Baixar Fiscalização - Obras", data=gerar_excel_modelo(df_fiscalizacao), file_name="FISCALIZACAO_TOTAIS.xlsx", use_container_width=True)

    st.markdown("---")
    st.markdown("### 3. Filtros Inteligentes e Controle de Escopo")
    c_flt1, c_flt2 = st.columns(2)
    with c_flt1:
        st.markdown("**🗂️ Triagem Dinâmica de Notas (Bolo Geral)**\nAssim que você sobe a planilha, o sistema abre uma interface para escolher exatamente quais tipos de Nota e quais Status devem passar pelo roteirizador. O resto é descartado instantaneamente, poupando edições prévias no Excel.")
        st.markdown("**⚡ Super Pontos (Deduplicação Espacial)**\nSe a IA encontrar notas sobrepostas num pequeno raio (ex: 50m), ela funde tudo num mega-ícone amarelo (`🏢 SUPER PONTO`), garantindo que o técnico saiba que no mesmo poste ou condomínio existem X obras agrupadas. A IA faz a contagem correta e desmembra os pacotes no Excel.")
        st.markdown("**📍 Atribuição Flexível (Tático)**\nVocê pode forçar a IA a amarrar as notas aos técnicos de duas formas: 'Por Município Base' (isola ele rigidamente na cidade da planilha) ou 'Por Proximidade Espacial' (a IA ignora as fronteiras municipais e empurra as obras para a equipe que estiver geograficamente mais perto).")
    with c_flt2:
        st.markdown("**🚨 Prioridades (Obras Fura Fila)**\nVocê pode selecionar quais siglas de obra são Críticas (ex: DIF, ASC, MGD). Elas receberão ícones e cores diferentes (Vermelho) no Google Earth, e serão levadas para o topo da matriz de decisão do sistema VRP.")
        st.markdown("**🎯 Planilhas Genéricas Customizadas**\nNo Planejamento Tático, um upload oculto e dinâmico permite que você suba arquivos sem nenhum padrão. Ao subir, o sistema pergunta: 'Qual coluna aqui dentro significa Prioridade?' e 'Qual coluna é o Status?'. Depois de escolher, ele funde e traça com a planilha principal.")
        st.markdown("**🛑 Trava Total de Obras**\nVocê pode limitar o esforço logístico (ex: Limitar a 500 notas num mar de 2.000). A IA fará a varredura das 500 notas prioritárias/melhores localizadas e rejeitará todo o resto sem sequer tentar rotear, poupando processamento e restringindo a equipe.")

    st.markdown("---")
    st.markdown("### 4. Esforços, Configurações e Cálculos Matemáticos")
    st.markdown("* **Varredura Reversa (Longe -> Perto):** Permite inverter a lógica da rota diária. Em vez de começar pelas obras da esquina da casa/base do técnico, a IA manda ele primeiro para o ponto extremo do mapa e vem roteando ele de volta, para que ele termine o expediente no quintal de casa.\n* **Cálculo de Postes vs Super Pontos:** Na Lista Contínua e Tático, a IA captura o menor valor numérico válido de Poste de MT/BT da obra (pois ambas podem compartilhar o poste físico) para fazer as projeções semanais. Na Fiscalização, a métrica oficial lida com a coluna `QTD PREVISTA DE POSTES` puramente.\n* **🎨 Termografia Visual (Fiscalização):** No mapa e no KML exportado, os pinos de fiscalização ganham cores pela carga horária (Volume de Postes): 🟢 Verde (Até 100), 🔵 Azul (Até 200), 🟡 Bege (Até 300), 🟠 Laranja (Até 400), e 🔴 Vermelho (Mais de 400 postes).")
    
    st.markdown("---")
    st.markdown("### 5. As Saídas (O que o Motor Exporta)")
    st.info("* **Traçado OSRM de Ruas vs Vetorial Rápido:** Ao marcar 'Traçado de Ruas Real (Lento)', a inteligência deforma as linhas retas acompanhando meticulosamente o asfalto das rodovias usando servidores OSRM externos. Sem ele, a visualização fica em formato teia de aranha (vetor) e roda super rápido.\n* **Planilhas Individuais (Rotas_Nome.xlsx):** Contêm **exatamente** as colunas que você marcou na interface, acrescidas de `SUPER_PONTO`, `DIA_SEMANA` e `DIA_MES` em posições fixas.\n* **Pacote KML Globais e Individuais:** Para o Google Earth. Ricamente HTML-formatados. Mostram os dados num pop-up colorido e estruturado ao clicar no pino da obra. O Tático divide as pastas do KML por Dias da Semana.\n* **Pacote GPX:** O GPS Offline puro. Usado para plugar as rotas no app Wikiloc ou OsmAnd. Essencial para áreas rurais sem sinal de rede 3G/4G.\n* **Obras_Correcao.xlsx:** O repositório das obras defeituosas. Se uma coordenada veio invertida, ausente, ou (no caso da fiscalização) se escapou para mais de 70km, o sistema barra e coloca nesta planilha indicando o `MOTIVO_REJEICAO` para que o Backoffice conserte o erro na SAP/SISCO.")

if __name__ == "__main__":
    renderizar_faq()
