import streamlit as st
import pandas as pd
import io
import os

def injetar_logo():
    if os.path.exists("LOGO_NIP.png"): st.logo("LOGO_NIP.png", icon_image=None)

st.set_page_config(page_title="FAQ e Manual do Usuário", page_icon="📖", layout="wide")
injetar_logo()

def gerar_excel_modelo(df_modelo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_modelo.to_excel(writer, index=False, sheet_name='Modelo')
    return output.getvalue()

def renderizar_faq():
    with st.sidebar:
        st.info("💡 **Dica:** Utilize esta central para treinar novos colaboradores no uso do Cérebro Logístico NIP.")

    st.markdown("<h1 class='brand-title' style='margin-bottom: 5px;'>📖 Manual de Operação e Inteligência Geográfica</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 16px; color: #555; margin-bottom: 30px;'>Bem-vindo ao manual completo do <b>Roteirizador NIP v3.0</b>. Entenda profundamente as lógicas matemáticas, as travas de segurança e o funcionamento interno do Cérebro Logístico que toma as decisões de roteamento.</p>", unsafe_allow_html=True)
    
    # ==========================================
    # SEÇÃO 1: ESTRATÉGIAS DE DESPACHO
    # ==========================================
    st.markdown("### 1. Entendendo os 3 Módulos do Sistema")
    st.markdown("O sistema foi desenhado para atender três momentos completamente diferentes da operação:")
    
    c_faq1, c_faq2, c_faq3 = st.columns(3)
    with c_faq1: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #0D256C; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <h4 style='color: #0D256C; margin-top: 0;'>🗺️ 1. Planejamento Tático</h4>
            <b>Para que serve:</b> É o distribuidor global. Usado quando você tem um bolo massivo de obras (ex: 5.000 notas) e precisa dividi-las do zero para uma equipe (ex: 20 técnicos), garantindo a mesma quantidade diária para todos.<br><br>
            <b>Como a IA pensa:</b> Ela fixa os técnicos em suas cidades bases, puxa as notas mais prioritárias da fila, agrupa as que estão próximas fisicamente e vai preenchendo a "mochila" diária de cada técnico. Se a cota for 6 obras por dia, ao rotear a 6ª obra a IA corta a rota, traça a linha de retorno para casa, e empurra a 7ª nota para o dia seguinte.
        </div>
        """, unsafe_allow_html=True)
    with c_faq2: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #d9534f; border-radius: 8px; height: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>
            <h4 style='color: #d9534f; margin-top: 0;'>📜 2. Lista Contínua</h4>
            <b>Para que serve:</b> O backlog já está definido. O seu backoffice ou cliente já definiu de quem é a obra (A coluna LEVANTADOR/FISCAL já veio preenchida na planilha original).<br><br>
            <b>Como a IA pensa:</b> O sistema atua puramente como um otimizador de caminho (Problema do Caixeiro Viajante). Ele não impõe limites de horário nem corta a rota; apenas liga todos os pontos pertencentes ao técnico da maneira mais curta possível. Ideal para pacotes fechados de serviços.<br><br>
            <b>Dicionário de Nomes:</b> Exclusivo deste módulo. Se as obras estiverem designadas para "EQUIPE 01", você pode subir a planilha de Equipes para que a IA faça o PROCV automático e troque "EQUIPE 01" pelo nome real "JOÃO DA SILVA".
        </div>
        """, unsafe_allow_html=True)
    with c_faq3: 
        st.markdown("""
        <div style='background: #f8f9fa; padding: 20px; border-left: 5px solid #FF9800; border-radius: 8px; height: 100%; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <h4 style='color: #FF9800; margin-top: 0;'>📋 3. Fiscalização</h4>
            <b>Para que serve:</b> Auditoria por produtividade focada em Volume de Esforço (Postes), e não em Obras rasas.<br><br>
            <b>A Regra do Bolsão de Densidade:</b> A IA não roteiriza pelo caminho mais próximo primeiro. Ela varre o estado e encontra a obra com a MAIOR quantidade de postes. Ela ancora o fiscal nessa mega obra. Depois, ela "puxa" as obras menores que estão fisicamente próximas a esse gigante, otimizando o lucro/rendimento do fiscal para aquela região.<br><br>
            <b>Termografia Visual:</b> O KML exportado é colorido pela densidade. Pino verde = poucos postes. Pino Vermelho = alta complexidade e muitos postes.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # ==========================================
    # SEÇÃO 2: DOWNLOAD DE MODELOS
    # ==========================================
    st.markdown("### 2. Padrões de Planilhas (Templates)")
    st.markdown("Para que a leitura automatizada funcione 100% livre de erros, os cabeçalhos das suas planilhas devem seguir os padrões estruturais abaixo. Você pode baixar os arquivos modelo em branco.")
    
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
    with col_dl1: st.download_button("📥 Tático - Equipes", data=gerar_excel_modelo(df_equipes_tat), file_name="LEVANTADORES_PRINCIPAIS.xlsx", use_container_width=True)
    with col_dl2: st.download_button("📥 Tático - Demandas", data=gerar_excel_modelo(df_obras_tat), file_name="BASE_LEVANTAMENTO.xlsx", use_container_width=True)
    with col_dl3: st.download_button("📥 Lista Contínua", data=gerar_excel_modelo(df_continua), file_name="BASE_LISTA_CONTINUA.xlsx", use_container_width=True)
    with col_dl4: st.download_button("📥 Fiscal. - Equipes", data=gerar_excel_modelo(df_fiscais), file_name="FISCAIS_PRINCIPAIS.xlsx", use_container_width=True)
    with col_dl5: st.download_button("📥 Fiscal. - Obras", data=gerar_excel_modelo(df_fiscalizacao), file_name="FISCALIZACAO_TOTAIS.xlsx", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # ==========================================
    # SEÇÃO 3: LÓGICAS E FILTROS DE DADOS
    # ==========================================
    st.markdown("### 3. Motores de Limpeza, Fusão e Filtros Dinâmicos")
    
    with st.expander("🌍 Limpeza Geográfica e a Planilha de Correções", expanded=True):
        st.markdown("""
        Um dos maiores custos operacionais é enviar o técnico para um endereço que caiu no mar ou em outro estado devido a erros de digitação no SAP/SISCO. O Roteirizador conta com três escudos:
        1. **Filtro de Coordenada Zerada/Branca:** Bloqueia automaticamente obras sem `LATITUDE` ou `LONGITUDE`.
        2. **Filtro de Inversão e Sinal:** Retém obras onde a Latitude é maior que a Longitude (erro comum de digitação) ou onde o sinal de menos (-) foi esquecido, indicando Hemisfério Norte ao invés do Brasil.
        3. **Cerca Eletrônica de 70km (Exclusivo da Fiscalização):** A ferramenta calcula um raio de 70km usando o IBGE como âncora a partir do nome do Município. Se as coordenadas cadastradas fugirem dessa barreira, a IA retém a nota para evitar deslocamentos massivos irreais.
        
        *Todas essas notas defeituosas são injetadas em um arquivo chamado `Obras_Correcao.xlsx` no momento da exportação, apontando a justificativa técnica para o Backoffice tratar, preservando 100% o KML da equipe limpo.*
        """)

    with st.expander("🏢 Super Pontos e Deduplicação Espacial", expanded=True):
        st.markdown("""
        O sistema possui um algoritmo que escaneia o raio de atuação de todas as obras (*O raio pode ser ajustado por você na barra lateral, em metros*).
        * **O que acontece:** Se a IA identificar 5 obras diferentes em coordenadas idênticas ou separadas por menos metros do que a trava estipulada (Ex: 5 apartamentos no mesmo prédio), ela as funde em uma única superestrutura.
        * **Visualização KML:** Ao invés de exibir 5 ícones sobrepostos que confundem o Google Earth, a IA coloca um mega-ícone dourado chamado **🏢 SUPER PONTO**. No pop-up gerado, haverá uma listagem *scrollável* contendo os 5 protocolos e dados daquelas notas empilhadas.
        * **Exportação no Excel:** No Excel exportado para o técnico, a IA é inteligente o suficiente para desfazer a fusão. As 5 obras aparecem em 5 linhas individuais, porém a coluna `SUPER_PONTO` fica marcada como "SIM (5 Obras)" e as linhas são pintadas de laranja para alertar que é um trabalho concentrado no mesmo lugar.
        """)

    with st.expander("🎯 Triagem Dinâmica, Prioridades e Planilha Genérica", expanded=True):
        st.markdown("""
        * **Filtro Nativo de Prioridade:** Se você selecionar tipos de nota de Alta Urgência (Ex: DIF, ASC, MGD) na interface, elas ganham preferência na fila do roteiro de cada técnico e são mapeadas com pinos e cores diferentes no KML final.
        * **O Fura-Fila (Triagem de Status):** Ao subir a planilha, os módulos varrem a coluna de Status. Você seleciona o que entra (ex: 'APTO PARA CAMPO'). O sistema limpa as obras "Concluídas", "Canceladas" ou "Suspensas" do motor sem você precisar editar o Excel original.
        * **Uso da Planilha Genérica:** Se sua Regional usa uma planilha com cabeçalhos bagunçados que não aderem aos modelos, você pode upá-la no campo especial **"3. Planilha Genérica"**. A interface abrirá seletores manuais para você apontar para a IA: *"Esta é a coluna que dita o Status"* e *"Esta é a coluna que dita a Prioridade"*. O sistema formatará e engolirá os dados mesclando-os à demanda padrão.
        """)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # ==========================================
    # SEÇÃO 4: CONFIGURAÇÕES E CÁLCULOS MATEMÁTICOS
    # ==========================================
    st.markdown("### 4. Algoritmos, Varredura Reversa e Cálculos de OSRM")
    
    c_mtr1, c_mtr2 = st.columns(2)
    with c_mtr1:
        st.markdown("#### Sentido do Roteamento")
        st.markdown("**📍 Lógica Padrão (Vizinho Mais Próximo):** A IA solta o técnico da base (sua cidade ou obra âncora) e constrói a rota pegando a próxima obra mais perto geograficamente. Isso faz com que o fim de tarde do técnico ocorra nos pontos mais remotos, longe de casa.")
        st.markdown("**🎯 Varredura Reversa (Tático e Contínua):** Esta inteligência vira o mapa de cabeça para baixo. A IA encontra a obra MAIS LONGE no raio de atuação e joga o técnico lá às 8h da manhã. O caminho é montado de lá para cá, garantindo que o técnico encerre o expediente quase na porta de casa/base, reduzindo horas extras com deslocamento noturno.")
    with c_mtr2:
        st.markdown("#### Motores de Mapa (Reta vs Asfalto)")
        st.markdown("**📡 Vetorial Rápido (Euclidiano):** A distância calculada e o traçado exportado em KML conectam as obras em linha reta sobre o mapa (modo voo de pássaro). Ideal para processamentos super massivos (> 2.000 notas) onde a agilidade é crucial. Roda em segundos.")
        st.markdown("**🛣️ Traçado de Ruas Real (OSRM):** A API OSRM puxa dados de satélite da infraestrutura rodoviária e asfáltica. A linha traçada nas exportações vai curvar, respeitar mão contramão, desviar de lagoas e desenhar a rota estritamente por onde o carro consegue passar. O processamento é lento (Pode levar vários minutos) e requer conexão estável de internet.")

    st.markdown("---")
    
    # ==========================================
    # SEÇÃO 5: EXPORTAÇÃO
    # ==========================================
    st.markdown("### 5. Entendendo os Arquivos de Exportação (Arquivos ZIP)")
    st.markdown("Sempre que você clica para Baixar, o sistema compacta toda a inteligência processada em pastas e zipa. Eis o que tem lá dentro:")
    
    st.info("""
    * **📊 Resumo_Operacional.xlsx:** Um Dashboard Gerencial instantâneo. Em vez de ler 5 mil linhas de obras, esta planilha resume: Nome do Técnico, Quantidade de Obras Alocadas, Postes Planejados por Dia/Semana, Quantidade de Super Pontos identificados e Total de KM da rota desenhada para ele.
    * **📁 Pastas Individuais (KML_Rotas / Excel_Rotas):** A ferramenta desmembra a base de dados de 50.000 obras em 20, 50 ou 100 planilhas Excel exclusivas (uma para cada Fiscal/Técnico), além de gerar um KML individual para cada um deles focado somente na rota diária que lhes compete.
    * **🗺️ ROTA_TOTAL.kml:** Um arquivo consolidado para gestores. Joga a atuação de toda a empresa no Google Earth, pintando o caminho de cada equipe com uma linha de cor viva e diferente, facilitando a visão de macro-áreas de sobreposição e conflito.
    * **🛰️ Arquivos GPX:** Exportados obrigatoriamente. São rastros GPS de Alta Precisão (Waypoints) que podem ser enviados via WhatsApp para a equipe importar em navegadores off-line, como *OsmAnd* ou *Wikiloc*, garantindo que achemistividade persista no mato profundo onde a internet não chega.
    * **Demanda_Tatica_Total.xlsx:** A base completa já roteirizada. Esta planilha **mantém 100% das suas colunas brutas originais** embutidas no meio das métricas calculadas pelo Roteirizador (`ORDEM`, `SUPER_PONTO`, `DIA_SEMANA`), prontas para subir nos sistemas SAP da empresa sem precisar recolar os dados.
    """)

if __name__ == "__main__":
    renderizar_faq()
