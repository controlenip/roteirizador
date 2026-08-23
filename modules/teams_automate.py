import requests
import base64
import streamlit as st
import time

def disparar_rotas_power_automate(fiscais_reais, bases_records, url_webhook, pacotes_bytes):
    """
    Função que envia os arquivos gerados via HTTP POST para o Webhook do Power Automate.
    pacotes_bytes: Dicionário contendo os arquivos em bytes { 'NOME_FISCAL': {'excel': bytes, 'kml': bytes} }
    """
    if not url_webhook or "http" not in url_webhook:
        st.error("🚨 ERRO: A URL do Webhook do Power Automate é inválida.")
        return

    sucessos = 0
    erros = 0

    barra_progresso = st.progress(0)
    status_texto = st.empty()

    for idx, fiscal in enumerate(fiscais_reais):
        # Atualiza UI
        barra_progresso.progress((idx + 1) / len(fiscais_reais))
        status_texto.text(f"📤 Enviando rota para {fiscal}...")

        # Busca o email do fiscal na planilha original
        registro = next((x for x in bases_records if x.get('BASE_NOME', x.get('LEVANTADOR', '')) == fiscal), None)
        email_teams = registro.get('EMAIL_TEAMS', '').strip() if registro else ''

        if not email_teams or str(email_teams) == 'nan':
            st.warning(f"⚠️ {fiscal} ignorado: Sem e-mail cadastrado na coluna EMAIL_TEAMS.")
            erros += 1
            continue

        arquivos_do_fiscal = pacotes_bytes.get(fiscal)
        if not arquivos_do_fiscal:
            continue

        # Codifica os arquivos em Base64 para trafegar via HTTP
        excel_b64 = base64.b64encode(arquivos_do_fiscal['excel']).decode('utf-8')
        kml_b64 = base64.b64encode(arquivos_do_fiscal['kml']).decode('utf-8')

        # Monta a "Carga" (Payload) JSON que o Power Automate está esperando
        payload = {
            "nome_fiscal": fiscal,
            "email_destino": email_teams,
            "nome_arquivo_excel": f"Rota_{fiscal.replace(' ', '_')}.xlsx",
            "arquivo_excel_b64": excel_b64,
            "nome_arquivo_kml": f"Rota_{fiscal.replace(' ', '_')}.kml",
            "arquivo_kml_b64": kml_b64
        }

        # Dispara para o Power Automate
        try:
            resposta = requests.post(url_webhook, json=payload, headers={"Content-Type": "application/json"})
            if resposta.status_code in [200, 202]:
                sucessos += 1
            else:
                st.error(f"❌ Falha ao enviar para {fiscal}. Código HTTP: {resposta.status_code}")
                erros += 1
        except Exception as e:
            st.error(f"❌ Erro de conexão ao enviar para {fiscal}: {e}")
            erros += 1
            
        time.sleep(1) # Pausa de 1 segundo para não sobrecarregar a API da Microsoft

    status_texto.empty()
    st.success(f"✅ Disparo Concluído! **{sucessos}** Fiscais notificados via Teams. ({erros} erros ou sem e-mail).")
