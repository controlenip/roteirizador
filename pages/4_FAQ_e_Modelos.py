import streamlit as st
import pandas as pd
import io
from modules.export_utils import injetar_logo

st.set_page_config(page_title="FAQ e Modelos", page_icon="📖", layout="wide")

# Aciona a Logo acima do Menu
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
    
    st.markdown("### 1. As Três Estratégias de Despacho (Modos)")
    c_faq1, c_faq2, c_faq3 = st.columns(3)
    with c_faq1: 
        st.markdown("<div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #0D256C; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'><h4 style='color: #0D256C; margin-top: 0;'>🎯 1. Planejamento Tático</h4><b>A IA no Comando:</b> Você sobe a planilha de técnicos e joga milhares de obras brutas. O sistema lê onde cada técnico atua, calcula todas as distâncias cruzadas e distribui as obras do zero de forma otimizada.<br><br><b>A Regra dos 100km:</b> Se o técnico terminar a cota do município dele antes de bater a meta diária, a IA usa o radar de 100km para acionar obras de cidades vizinhas.</div>", unsafe_allow_html=True)
    with c_faq2: 
        st.markdown("<div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #55B929; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'><h4 style='color: #2e7d32; margin-top: 0;'>♾️ 2. Lista Contínua</h4><b>O Usuário no Comando:</b> A IA respeita estritamente o que você definiu. Sua planilha já deve ter a coluna <b>LEVANTADOR</b> preenchida.<br><br><b>Processamento:</b> A ferramenta ignora as travas de fim de expediente e desenha o caminho mais curto para conectar 100% da lista do profissional. Nenhuma nota é transferida de um técnico para outro.</div>", unsafe_allow_html=True)
    with c_faq3: 
        st.markdown("<div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #FF9800; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'><h4 style='color: #FF9800; margin-top: 0;'>📋 3. Fiscalização</h4><b>A Regra do Bolsão:</b> A IA garimpa as obras com a MAIOR quantidade de postes e ancora o Fiscal mais próximo nessa região. Em seguida, ela traça a rota intercalando as obras menores que estão no caminho, eliminando espaços vazios de deslocamento.<br><br><b>Cerca Eletrônica:</b> Uma trava de segurança bloqueia obras com coordenadas vazias ou fora da cidade (>70km), gerando a Planilha de Correção.</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 2. Baixar Modelos de Planilhas (Templates)")
    
    df_equipes = pd.DataFrame([{'MunicIpio': 'FORTUNA', 'Estado': 'Maranhão', 'Levantador': 'NOME DO TECNICO', 'Regional': 'CENTRO', 'Longitude': -44.0264, 'Latitude': -5.7335, 'Equipe': 'EQUIPE 17'}])
    df_levantamento = pd.DataFrame([{'ID SISCO': 1982315, 'PROTOCOLO': 1081945188, 'TIPO NOTA': 'UNR', 'PRIORIDADE': 0, 'STATUS SAP': 'ATIV', 'STATUS SISCO': 'Liberado para Levantamento', 'STATUS LIST': 'Em levantamento', 'FASE': 'MO', 'PAT': 'PAT1', 'Regional': 'LESTE', 'Município': 'CODÓ', 'DISTANCIA BT': 248.78, 'DISTANCIA MT': 241.99, 'DISTANCIA TRAFO': 243.4, 'POSTE PREVISTO BT': 6, 'POSTE PREVISTO MT': 2, 'NOME': 'NOME DO CLIENTE', 'CONTA CONTRATO': 3019326160, 'INSTALAÇÃO': 2000876176, 'ENDEREÇO': 'ENDERECO COMPLETO', 'LOCALIDADE': 'RURAL', 'LATITUDE': -4.459156, 'LONGITUDE': -44.150418, 'INFORMAÇÕES EXTRAS': 'INFORMACOES ADICIONAIS'}])
    df_continua = pd.DataFrame([{'LEVANTADOR': 'NOME DO TECNICO', 'DATA DESPACHO': '17/08/2026', 'ID SISCO': 1982149, 'PROTOCOLO': 1076894592, 'TIPO NOTA': 'UNR', 'PRIORIDADE': 0, 'STATUS SAP': 'ATIV', 'STATUS SISCO': 'Pré Análise', 'STATUS LIST': 'Em levantamento', 'FASE': 'MO', 'PAT': 'PAT1', 'Regional': 'LESTE', 'Município': 'CAXIAS', 'DISTANCIA BT': 285.18, 'DISTANCIA MT': 212.99, 'DISTANCIA TRAFO': 224.66, 'POSTE PREVISTO BT': 7, 'POSTE PREVISTO MT': 2, 'NOME': 'NOME DO CLIENTE', 'CONTA CONTRATO': 3018299797, 'INSTALAÇÃO': 2000835041, 'ENDEREÇO': 'ENDERECO COMPLETO', 'LOCALIDADE': 'RURAL', 'LATITUDE': -5.060694, 'LONGITUDE': -43.438846, 'INFORMAÇÕES EXTRAS': 'INFORMACOES ADICIONAIS'}])
    df_fiscalizacao = pd.DataFrame([{'NOTA': 1081945188, 'FISCAL': 'NOME DO FISCAL', 'STATUS DA FISCALIZACAO': 'APTO PARA CAMPO', 'QTD PREVISTA DE POSTES': 45, 'MUNICIPIO': 'SAO LUIS', 'LATITUDE': -2.5297, 'LONGITUDE': -44.3028, 'VALOR DA OBRA': 15000.50, 'PREVISAO DE ENTREGA': '10/10/2026', 'TIPO DE FISCALIZACAO': 'NORMAL', 'TIPO DE PROJETO': 'EXTENSAO', 'REGIONAL': 'LESTE', 'ZONA': 'URBANA', 'BACKOFFICE DA FISCALIZACAO': 'NOME BACKOFFICE', 'OBSERVACAO': 'Atenção ao prazo'}])

    col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)
    with col_dl1: st.download_button("📥 Baixar Modelo Equipes", data=gerar_excel_modelo(df_equipes), file_name="MODELO_LEVANTADORES.xlsx", use_container_width=True)
    with col_dl2: st.download_button("📥 Baixar Modelo Obras Livres", data=gerar_excel_modelo(df_levantamento), file_name="MODELO_BASE_LEVANTAMENTO.xlsx", use_container_width=True)
    with col_dl3: st.download_button("📥 Baixar Modelo Contínua", data=gerar_excel_modelo(df_continua), file_name="MODELO_LISTA_CONTINUA.xlsx", use_container_width=True)
    with col_dl4: st.download_button("📥 Baixar Modelo Fiscalização", data=gerar_excel_modelo(df_fiscalizacao), file_name="MODELO_FISCALIZACAO.xlsx", use_container_width=True)

    st.markdown("---")
    st.markdown("### 3. Filtros Inteligentes e Controle de Escopo")
    c_flt1, c_flt2 = st.columns(2)
    with c_flt1:
        st.markdown("**🗂️ Triagem Dinâmica de Notas (Bolo Geral)**\nAssim que você sobe a planilha, um Mini-Dashboard mostra os totais dos principais tipos de notas (UNR, MGD, ASC, DIF). Abaixo dele, você pode escolher tipos específicos (ex: 'UNR') para **descartar temporariamente**, limpando a base sem precisar editar o arquivo Excel.")
        st.markdown("**🎯 Matriz Multi-Filtro (Regional e PAT)**\nNa seção 'Escopo da Operação', o sistema lê todas as Regionais e PATs presentes no seu arquivo. Você pode afunilar o roteamento para processar *apenas* a 'REGIONAL LESTE' e *apenas* o 'PAT1', bloqueando o restante do Estado instantaneamente.")
        st.markdown("**⚡ Super Pontos (Deduplicação Espacial)**\nSe a IA encontrar notas sobrepostas em um pequeno raio de ação, ela funde tudo num mega-ícone laranja (`SUPER PONTO`), garantindo que o técnico visite o local apenas uma vez. **Você pode ajustar o raio desse agrupamento (em metros) na barra lateral.**")
    with c_flt2:
        st.markdown("**🔥 Alta Densidade (Modo Produtividade Máxima)**\nUma trava que foca apenas no que dá lucro de tempo. Quando ativada na barra lateral, a IA varre o mapa e joga fora as obras isoladas ou esparsas na zona rural, garantindo que ela não crie rotas para cidades que nenhum técnico atende. A equipe é enviada apenas para os 'Bolsões de Densidade', e as obras isoladas vão para o arquivo de rejeições para tratativa futura.")
        st.markdown("**🚨 Tripla Checagem de Prioridade (Fura Fila)**\nO sistema exige urgência. A obra fura a fila do roteiro e fica vermelha no mapa se: 1) Você selecioná-la manualmente nos Filtros Dinâmicos; 2) For detectado o status 'CORREÇÃO DE LEVANTAMENTO'; 3) A coluna nativa `PRIORIDADE` no Excel tiver marcações urgentes (ex: 'GIRO NO PRAZO').")
        st.markdown("**📍 Atribuição de Equipes (Proximidade vs. Município)**\nVocê pode forçar a IA a respeitar rigidamente o município escrito na planilha do técnico, ou usar a atribuição por proximidade, onde a IA ignora a cidade de cadastro e joga o técnico no maior bolsão de obras que estiver no raio de alcance do seu GPS (Ideal para Fiscais).")

    st.markdown("---")
    st.markdown("### 4. Esforços, Limites e Avisos Gerenciais")
    st.markdown("* **🛑 Trava Total de Operação (O Limite Global da Empresa):** Localizado na barra lateral, define um teto absoluto. Se você digitar **300**, a IA vai garimpar as 300 melhores/mais prioritárias notas do estado e rejeitar todo o resto, poupando a equipe de backoffice. **Se deixar no valor '0' (Zero)**, a trava é desligada e o sistema roteiriza 100% da base que encontrar.\n* **O Paredão Diário (Corte Rígido):** Se a meta for **6 Obras Previstas por Dia**, na hora que a IA montar a 6ª obra, ela aborta o cálculo, traça a linha de \"Retorno para a Base\" e a 7ª obra cai para a Terça-Feira, blindando o técnico de sobrecarga.\n* **Varredura Reversa (Longe -> Perto):** Permite inverter a lógica da rota diária. Em vez de começar pelas obras da esquina, a IA manda o técnico cedo para a fazenda mais distante do mapa e vem puxando ele de volta obra por obra, para que o fim do expediente seja feito a poucos minutos de casa.\n* **Cálculo de Postes e Malhas:** Nos relatórios, o sistema soma as colunas `POSTE PREVISTO BT` e `POSTE PREVISTO MT`. Para evitar contagem em dobro (pois Alta e Baixa tensão costumam dividir o mesmo poste físico), ele usa matematicamente o Menor Valor entre as duas.\n* **🎨 Termografia Visual (Fiscalização):** No mapa de fiscalização e no arquivo KML exportado, os pinos de obras ganham cores automáticas baseadas no volume de postes para facilitar a auditoria visual: 🟢 Verde (Até 100), 🔵 Azul (Até 200), 🟡 Bege (Até 300), 🟠 Laranja (Até 400), e 🔴 Vermelho (Mais de 400 postes).")
    
    st.markdown("---")
    st.markdown("### 5. Configurações Avançadas e Saídas (O que você baixa)")
    st.info("* **Traçado de Ruas Real (OSRM) vs. Vetorial Rapido:** Nas configurações de Conexão de Rede, a opção de *Traçado de Ruas Lento* usa uma API global para curvar a linha exatamente pelas rodovias e asfaltos. Se desmarcado (Vetorial Rápido), ele liga as obras em linha reta (padrão satélite), acelerando o tempo de geração de 10 minutos para apenas alguns segundos.\n* **Demanda_Geral.xlsx:** Uma compilação cristalina. A planilha exportada contém **exatamente** as colunas originais do seu projeto, blindadas contra lixo de programação. O sistema faz a autolimpeza com a função `limpar_colunas_excel()` garantindo que os identificadores primários (como PROTOCOLO, REGIONAL e LAT/LON) nunca sumam da entrega final.\n* **Pacote KML e KML de Rejeições:** O KML principal roda em Google Earth (limpo de caixas de textos desnecessárias). Além dele, se alguma obra for isolada pela Alta Densidade ou esgotar a cota da Trava Global, a IA gera o arquivo **`OBRAS_NAO_ALOCADAS.kml`** (pinos brancos) para você visualizar exatamente o que sobrou.\n* **Pacote GPX:** O GPX é o **GPS Offline de Alta Precisão** – feito para o técnico importar em apps como *OsmAnd* ou *Wikiloc* para navegar no sertão e em áreas rurais mesmo quando estiver com 0% de sinal de operadora móvel. Construído nativamente pela função interna `gerar_gpx_simples()`.\n* **Planilha de Correção (Cerca Eletrônica):** Se alguma obra apresentar coordenadas zeradas, invertidas ou a dezenas de quilômetros do centro da cidade cadastrada, ela é barrada. O sistema cria automaticamente o arquivo `Obras_Correcao.xlsx` detalhando o motivo exato do erro para o seu backoffice tratar.")

if __name__ == "__main__":
    renderizar_faq()
