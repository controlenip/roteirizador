import pandas as pd
import streamlit as st
import io
import re
import html

@st.cache_data(show_spinner=False)
def ler_planilha_cached(file_content):
    return pd.read_excel(io.BytesIO(file_content))

def formatar_moeda(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formata_campo_html(val):
    val_str = str(val)
    if val_str.lower() == 'nan': return '-'
    if '|' in val_str:
        itens = val_str.split('|')
        if len(itens) > 1:
            lis = "".join([f"<li style='margin-bottom:3px;'><b>[{idx+1}]</b> {html.escape(i.strip())}</li>" for idx, i in enumerate(itens)])
            return f"<div style='max-height:95px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>{lis}</ul></div>"
    return html.escape(val_str)

def normalize_cols(cols):
    new_cols = []
    for c in cols:
        c = str(c).strip().upper()
        c = re.sub(r'[ÁÀÂÃÄ]', 'A', c)
        c = re.sub(r'[ÉÈÊË]', 'E', c)
        c = re.sub(r'[ÍÌÎÏ]', 'I', c)
        c = re.sub(r'[ÓÒÔÕÖ]', 'O', c)
        c = re.sub(r'[ÚÙÛÜ]', 'U', c)
        c = re.sub(r'Ç', 'C', c)
        new_cols.append(c)
    return new_cols

def normalizar_municipios(series_mun):
    s = series_mun.astype(str).str.upper()
    s = s.str.replace(r'[ÁÀÂÃÄ]', 'A', regex=True)
    s = s.str.replace(r'[ÉÈÊË]', 'E', regex=True)
    s = s.str.replace(r'[ÍÌÎÏ]', 'I', regex=True)
    s = s.str.replace(r'[ÓÒÔÕÖ]', 'O', regex=True)
    s = s.str.replace(r'[ÚÙÛÜ]', 'U', regex=True)
    s = s.str.replace(r'Ç', 'C', regex=True)
    return s.str.split('-').str[0].str.strip()

def atualizar_status_via_df(df_principal, df_status, coluna_alvo):
    try:
        chave_nome = df_status.columns[0]
        df_status[chave_nome] = df_status[chave_nome].astype(str).str.strip()
        df_status_map = df_status.set_index(chave_nome)[coluna_alvo].to_dict()
        if 'PROTOCOLO' in df_principal.columns:
            df_principal['PROTOCOLO_STR'] = df_principal['PROTOCOLO'].astype(str).str.strip()
            df_principal['STATUS LIST'] = df_principal['PROTOCOLO_STR'].map(df_status_map).fillna(df_principal.get('STATUS LIST', 'SEM INFORMAÇÕES'))
            df_principal = df_principal.drop(columns=['PROTOCOLO_STR'])
            st.success(f"✅ Status Sincronizados: {len(df_status_map)} registros atualizados!")
        else:
            st.warning("⚠️ Coluna 'PROTOCOLO' não encontrada na base principal.")
    except Exception as e:
        st.error(f"Erro na sincronização: {e}")
    return df_principal
