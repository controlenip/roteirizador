
import io
import re
import time
import math
import html
import hashlib
import zipfile
import unicodedata
import xml.etree.ElementTree as ET

from datetime import datetime, timedelta
from collections import defaultdict

import numpy as np
import pandas as pd
import requests
import streamlit as st
import folium

from sklearn.cluster import DBSCAN
from sklearn.neighbors import BallTree
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from openpyxl.styles import Font, PatternFill, Alignment


# ============================================================
# CONFIGURACAO GERAL
# ============================================================

st.set_page_config(
    page_title="NIP | Analise Cruzada",
    page_icon="📍",
    layout="wide"
)

EARTH_KM = 6371.0088
OSRM_DEFAULT = "https://router.project-osrm.org"

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_DUP = "OBRA SANEAMENTO E LEVANTAMENTO"

CATEGORIAS = [CAT_SAN, CAT_LEV, CAT_DUP]

SAP_EXCLUIDOS = {"FINL", "CANC"}

LIST_PERMITIDOS = {
    "0",
    "EM LEVANTAMENTO",
    "CORRECAO DE LEVANTAMENTO"
}

CONTRATOS_PERMITIDOS = {
    "0",
    "NIP GLOBAL LTDA - EQTL MARANHAO"
}

CORES = {
    CAT_SAN: "#2563EB",
    CAT_LEV: "#16A34A",
    CAT_DUP: "#9333EA"
}

CORES_FOLIUM = {
    CAT_SAN: "blue",
    CAT_LEV: "green",
    CAT_DUP: "purple"
}

DIAS_PT = {
    0: "SEGUNDA-FEIRA",
    1: "TERCA-FEIRA",
    2: "QUARTA-FEIRA",
    3: "QUINTA-FEIRA",
    4: "SEXTA-FEIRA",
    5: "SABADO",
    6: "DOMINGO"
}


# ============================================================
# VISUAL
# ============================================================

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 100%;
}

.nip-title {
    font-size: 2rem;
    font-weight: 800;
    color: #0D256C;
    margin-bottom: 0.2rem;
}

.nip-subtitle {
    font-size: 0.93rem;
    color: #64748b;
    margin-bottom: 1rem;
}

.nip-card {
    background: #FFFFFF;
    padding: 17px;
    border-radius: 11px;
    border: 1px solid #e2e8f0;
    border-left: 5px solid var(--cor);
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    margin-bottom: 12px;
}

.nip-card-label {
    color: #64748b;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
}

.nip-card-value {
    color: #172554;
    font-weight: 800;
    font-size: 28px;
}

.nip-card-desc {
    color: #64748b;
    font-size: 11px;
}

.section-title {
    color: #0D256C;
    font-size: 1.3rem;
    font-weight: 750;
    margin-top: 15px;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)


def card(titulo, valor, cor="#0D256C", descricao=""):
    return f"""
    <div class="nip-card" style="--cor:{cor}">
        <div class="nip-card-label">
            {html.escape(str(titulo))}
        </div>
        <div class="nip-card-value">
            {html.escape(str(valor))}
        </div>
        <div class="nip-card-desc">
            {html.escape(str(descricao))}
        </div>
    </div>
    """


def mostrar_cards(itens):
    colunas = st.columns(len(itens))
    for coluna, item in zip(colunas, itens):
        with coluna:
            st.markdown(
                card(*item),
                unsafe_allow_html=True
            )


def secao(titulo):
    st.markdown(
        f'<div class="section-title">{html.escape(titulo)}</div>',
        unsafe_allow_html=True
    )


# ============================================================
# CRONOMETRO
# ============================================================

def formatar_tempo(segundos):
    segundos = max(0, int(segundos))
    horas, resto = divmod(segundos, 3600)
    minutos, segundos = divmod(resto, 60)

    if horas:
        return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

    return f"{minutos:02d}:{segundos:02d}"


def iniciar_cronometro(chave):
    st.session_state[chave] = {
        "inicio": time.monotonic(),
        "decorrido": 0,
        "processados": 0,
        "total": 0,
        "concluido": False
    }


def atualizar_cronometro(
    chave,
    processados,
    total,
    container=None
):
    relogio = st.session_state.get(chave)

    if not relogio:
        return

    decorrido = (
        relogio["decorrido"]
        + time.monotonic()
        - relogio["inicio"]
    )

    relogio["processados"] = processados
    relogio["total"] = total

    if processados > 0 and total > processados:
        media = decorrido / processados
        restante = media * (total - processados)
    elif total == processados:
        restante = 0
    else:
        restante = None

    if container is not None:
        texto_restante = (
            formatar_tempo(restante)
            if restante is not None
            else "Calculando..."
        )

        container.markdown(
            f"""
            **⏱️ Decorrido:** {formatar_tempo(decorrido)}
            &nbsp;&nbsp;&nbsp;
            **🎯 Restante estimado:** {texto_restante}
            &nbsp;&nbsp;&nbsp;
            **📍 Progresso:** {processados}/{total}
            """
        )


def finalizar_cronometro(chave):
    relogio = st.session_state.get(chave)

    if relogio:
        relogio["decorrido"] += (
            time.monotonic()
            - relogio["inicio"]
        )
        relogio["concluido"] = True


# ============================================================
# NORMALIZACAO
# ============================================================

def norm(valor):
    if pd.isna(valor):
        return ""

    texto = unicodedata.normalize(
        "NFKD", str(valor)
    ).encode(
        "ascii", "ignore"
    ).decode().upper().strip()

    return re.sub(r"\s+", " ", texto)


def status(valor):
    texto = norm(valor)

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".")[0]

    return texto


def chave_nota(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]

    if norm(texto) in {
        "", "NAN", "NONE", "NULL", "0"
    }:
        return ""

    return texto


def texto_seguro(valor):
    if valor is None:
        return ""

    if isinstance(valor, (list, dict)):
        return str(valor)

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    return str(valor)


def procurar_coluna(df, nomes, obrigatoria=True):
    mapa = {
        norm(c): c for c in df.columns
    }

    for nome in nomes:
        if norm(nome) in mapa:
            return mapa[norm(nome)]

    if obrigatoria:
        raise ValueError(
            "Coluna obrigatoria nao encontrada: "
            + " / ".join(nomes)
        )

    return None


def converter_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(
            ",", ".", regex=False
        ),
        errors="coerce"
    )


def validar_coordenadas(df, c_lat, c_lon):
    valido = (
        df[c_lat].between(-35, 6)
        & df[c_lon].between(-75, -30)
        & df[c_lat].ne(0)
        & df[c_lon].ne(0)
    )

    df.loc[~valido, c_lat] = np.nan
    df.loc[~valido, c_lon] = np.nan


def escolher_aba(arquivo, grupos, preferida=None):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    abas = excel.sheet_names[:]

    if preferida:
        abas.sort(
            key=lambda a: (
                0 if norm(a) == norm(preferida) else 1
            )
        )

    for aba in abas:
        for header in range(5):
            try:
                amostra = pd.read_excel(
                    excel,
                    sheet_name=aba,
                    header=header,
                    nrows=3,
                    dtype=str
                )

                existentes = {
                    norm(c) for c in amostra.columns
                }

                valido = all(
                    any(
                        norm(alternativa) in existentes
                        for alternativa in grupo
                    )
                    for grupo in grupos
                )

                if valido:
                    df = pd.read_excel(
                        excel,
                        sheet_name=aba,
                        header=header,
                        dtype=str
                    )
                    return df, aba, header

            except Exception:
                continue

    raise ValueError(
        "Nenhuma aba compativel localizada. "
        "Abas existentes: "
        + ", ".join(excel.sheet_names)
    )


# ============================================================
# DATAS E COLUNAS B / C
# ============================================================

def converter_data(valor):
    """
    Aceita datas do Excel, texto dd/mm/aaaa,
    aaaa-mm-dd e numeros seriais do Excel.
    """
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor)

    texto = str(valor).strip()

    if norm(texto) in {
        "", "NAN", "NONE", "NULL", "0", "-"
    }:
        return pd.NaT

    # Numero serial de data do Excel.
    try:
        numero = float(texto.replace(",", "."))

        if 20000 <= numero <= 80000:
            return pd.Timestamp(
                "1899-12-30"
            ) + pd.Timedelta(days=numero)

    except (ValueError, OverflowError):
        pass

    return pd.to_datetime(
        valor,
        errors="coerce",
        dayfirst=True
    )


def responsavel_valido(valor):
    texto = norm(valor)

    return texto not in {
        "",
        "SEM LEVANTADOR",
        "NAN",
        "NONE",
        "NULL",
        "0",
        "-"
    }


# ============================================================
# BASE SANEAMENTO
# ============================================================

def ler_saneamento(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        dados = None

        for codificacao in ["utf-8-sig", "latin-1"]:
            try:
                dados = pd.read_csv(
                    io.BytesIO(arquivo.getvalue()),
                    sep=None,
                    engine="python",
                    encoding=codificacao,
                    dtype=str
                )
                break
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                continue

        if dados is None:
            raise ValueError("CSV de saneamento invalido.")

        df = dados
        aba = "CSV"

    else:
        df, aba, _ = escolher_aba(
            arquivo,
            [
                ["NOTA"],
                ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
            ],
            preferida="Clientes Existentes"
        )

    c_nota = procurar_coluna(df, ["NOTA"])

    c_municipio = procurar_coluna(
        df,
        ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
    )

    c_regional = procurar_coluna(
        df, ["REGIONAL"], False
    )

    c_lat = procurar_coluna(
        df,
        [
            "LATITUDE PROJETO",
            "LATITUDE",
            "LATITUDE CAMPO"
        ]
    )

    c_lon = procurar_coluna(
        df,
        [
            "LONGITUDE PROJETO",
            "LONGITUDE",
            "LONGITUDE CAMPO"
        ]
    )

    df["CHAVE_NOTA"] = df[c_nota].map(chave_nota)

    df["MUNICIPIO_OBRA"] = (
        df[c_municipio].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None else ""
    )

    df["LAT_OBRA"] = converter_numero(df[c_lat])
    df["LON_OBRA"] = converter_numero(df[c_lon])

    validar_coordenadas(
        df, "LAT_OBRA", "LON_OBRA"
    )

    return df, aba


# ============================================================
# BASE LEVANTAMENTO
# ============================================================

def motivos_levantamento(linha):
    motivos = []

    sap = linha["SAP_NORM"]
    lista = linha["LIST_NORM"]
    contrato = linha["CONTRATO_NORM"]
    m = linha["M_NORM"]
    n = linha["N_NORM"]

    if sap in SAP_EXCLUIDOS:
        motivos.append(f"SAP {sap}")

    if lista not in LIST_PERMITIDOS:
        motivos.append(
            "LIST NAO ACEITO: "
            + (lista or "VAZIO")
        )

    if contrato not in CONTRATOS_PERMITIDOS:
        motivos.append(
            "CONTRATO NAO ACEITO: "
            + (contrato or "VAZIO")
        )

    if m != "0":
        motivos.append(
            "ORCAMENTO MODULAR DIFERENTE DE 0"
        )

    if n != "0":
        motivos.append(
            "PLA ALVOS DIFERENTE DE 0"
        )

    # NOVA REGRA:
    # Coluna B com nome valido E coluna C com data valida.
    if linha["JA_EM_CAMPO"]:
        motivos.append(
            "JA EM CAMPO - EQUIPE E DATA PREENCHIDAS"
        )

    return " | ".join(motivos)


def ler_levantamento(arquivo):
    df, aba, header = escolher_aba(
        arquivo,
        [
            ["PROTOCOLO"],
            ["STATUS SAP", "STATUS_SAP"],
            ["STATUS LIST", "STATUS_LIST"]
        ],
        preferida="NOTAS"
    )

    c_nota = procurar_coluna(
        df, ["PROTOCOLO"]
    )

    c_sap = procurar_coluna(
        df, ["STATUS SAP", "STATUS_SAP"]
    )

    c_list = procurar_coluna(
        df, ["STATUS LIST", "STATUS_LIST"]
    )

    c_contrato = procurar_coluna(
        df, ["CONTRATO"]
    )

    c_municipio = procurar_coluna(
        df,
        ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
    )

    c_regional = procurar_coluna(
        df, ["REGIONAL"], False
    )

    c_lat = procurar_coluna(df, ["LATITUDE"])
    c_lon = procurar_coluna(df, ["LONGITUDE"])

    c_m = next(
        (
            c for c in df.columns
            if "ORCAMENTO MODULAR" in norm(c)
        ),
        None
    )

    c_n = next(
        (
            c for c in df.columns
            if "PLA ALVOS" in norm(c)
        ),
        None
    )

    if c_m is None or c_n is None:
        raise ValueError(
            "Colunas ORCAMENTO MODULAR ou "
            "PLA ALVOS nao encontradas."
        )

    if len(df.columns) < 3:
        raise ValueError(
            "Base Levantamento precisa ter "
            "pelo menos tres colunas."
        )

    # B e C sao posicoes fisicas da planilha
    # considerando o cabecalho detectado.
    coluna_b = df.columns[1]
    coluna_c = df.columns[2]

    c_prioridade = procurar_coluna(
        df, ["PRIORIDADE"], False
    )

    c_abertura = procurar_coluna(
        df,
        ["DATA ABERTURA", "DATA DE ABERTURA"],
        False
    )

    df["CHAVE_NOTA"] = df[c_nota].map(chave_nota)

    df["MUNICIPIO_OBRA"] = (
        df[c_municipio].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None else ""
    )

    df["LAT_OBRA"] = converter_numero(df[c_lat])
    df["LON_OBRA"] = converter_numero(df[c_lon])

    validar_coordenadas(
        df, "LAT_OBRA", "LON_OBRA"
    )

    df["SAP_NORM"] = df[c_sap].map(status)
    df["LIST_NORM"] = df[c_list].map(status)

    df["CONTRATO_NORM"] = (
        df[c_contrato].map(status)
    )

    df["M_NORM"] = df[c_m].map(status)
    df["N_NORM"] = df[c_n].map(status)

    df["PRIORIDADE_NORM"] = (
        df[c_prioridade].fillna("")
        if c_prioridade is not None else ""
    )

    df["DATA_ABERTURA_NORM"] = (
        df[c_abertura].fillna("")
        if c_abertura is not None else ""
    )

    # NOVAS COLUNAS B E C
    df["RESPONSAVEL_CAMPO"] = (
        df[coluna_b].fillna("").astype(str).str.strip()
    )

    df["DATA_CAMPO"] = (
        df[coluna_c].map(converter_data)
    )

    df["JA_EM_CAMPO"] = (
        df["RESPONSAVEL_CAMPO"].map(
            responsavel_valido
        )
        & df["DATA_CAMPO"].notna()
    )

    df["MOTIVOS_EXCLUSAO"] = df.apply(
        motivos_levantamento,
        axis=1
    )

    informacoes = {
        "ABA": aba,
        "LINHA_CABECALHO": header + 1,
        "COLUNA_B": str(coluna_b),
        "COLUNA_C": str(coluna_c),
        "JA_EM_CAMPO_LINHAS": int(df["JA_EM_CAMPO"].sum())
    }

    return df, informacoes


# ============================================================
# EQUIPES
# ============================================================

def ler_equipes(arquivo):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    partes = []

    for tipo, nome_aba in [
        ("LEVANTAMENTO", "LEVANTADORES"),
        ("SANEAMENTO", "SANEAMENTO")
    ]:
        aba = next(
            (
                a for a in excel.sheet_names
                if norm(a) == nome_aba
            ),
            None
        )

        if aba is None:
            raise ValueError(
                "Aba obrigatoria ausente: " + nome_aba
            )

        df = pd.read_excel(
            excel,
            sheet_name=aba,
            dtype=str
        )

        c_nome = procurar_coluna(
            df,
            ["NOME", "EQUIPE", "NOME_COLAB"]
        )

        c_cidade = procurar_coluna(
            df,
            ["CIDADES", "CIDADE", "MUNICIPIO"]
        )

        c_lat = procurar_coluna(
            df, ["LATITUDE", "LAT"]
        )

        c_lon = procurar_coluna(
            df, ["LONGITUDE", "LON"]
        )

        parte = pd.DataFrame({
            "EQUIPE": df[c_nome],
            "CIDADE_BASE": df[c_cidade],
            "LAT_EQUIPE": converter_numero(df[c_lat]),
            "LON_EQUIPE": converter_numero(df[c_lon]),
            "TIPO_EQUIPE": tipo
        })

        validar_coordenadas(
            parte, "LAT_EQUIPE", "LON_EQUIPE"
        )

        partes.append(parte)

    equipes = pd.concat(
        partes,
        ignore_index=True
    )

    equipes = equipes.dropna(
        subset=[
            "EQUIPE",
            "LAT_EQUIPE",
            "LON_EQUIPE"
        ]
    ).copy()

    equipes = equipes[
        equipes["EQUIPE"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    equipes["EQUIPE"] = (
        equipes["EQUIPE"].astype(str).str.strip()
    )

    equipes["ID_EQUIPE"] = (
        equipes["TIPO_EQUIPE"]
        + "|"
        + equipes["EQUIPE"]
        + "|"
        + equipes["CIDADE_BASE"].astype(str)
    )

    return equipes.drop_duplicates(
        subset=["ID_EQUIPE"]
    ).reset_index(drop=True)


# ============================================================
# CRUZAMENTO E AUDITORIA
# ============================================================

def primeira_valida(grupo):
    valido = (
        grupo["LAT_OBRA"].notna()
        & grupo["LON_OBRA"].notna()
    )

    if valido.any():
        return grupo.loc[valido].iloc[0]

    return grupo.iloc[0]


def consolidar(san, lev):
    san = san[
        san["CHAVE_NOTA"].ne("")
    ].copy()

    lev = lev[
        lev["CHAVE_NOTA"].ne("")
    ].copy()

    grupos_san = dict(
        tuple(san.groupby("CHAVE_NOTA", sort=False))
    )

    grupos_lev = dict(
        tuple(lev.groupby("CHAVE_NOTA", sort=False))
    )

    notas = dict.fromkeys(
        list(grupos_san) + list(grupos_lev)
    )

    registros = []

    for nota in notas:
        gs = grupos_san.get(nota)
        gl = grupos_lev.get(nota)

        tem_san = gs is not None
        tem_lev = gl is not None

        a = primeira_valida(gs) if tem_san else None
        b = primeira_valida(gl) if tem_lev else None

        if tem_san and tem_lev:
            categoria = CAT_DUP
        elif tem_san:
            categoria = CAT_SAN
        else:
            categoria = CAT_LEV

        motivos = []

        if tem_lev:
            for texto in gl["MOTIVOS_EXCLUSAO"]:
                if texto:
                    motivos.extend(texto.split(" | "))

        motivos = list(dict.fromkeys(motivos))

        ja_campo = bool(
            gl["JA_EM_CAMPO"].any()
        ) if tem_lev else False

        responsaveis = []
        datas_campo = []

        if tem_lev:
            em_campo = gl[gl["JA_EM_CAMPO"]]

            responsaveis = [
                str(x)
                for x in em_campo[
                    "RESPONSAVEL_CAMPO"
                ].dropna().unique()
            ]

            datas_campo = [
                x.strftime("%d/%m/%Y")
                for x in em_campo[
                    "DATA_CAMPO"
                ].dropna().unique()
            ]

        origem = a if a is not None else b

        if (
            b is not None
            and pd.notna(b["LAT_OBRA"])
            and pd.notna(b["LON_OBRA"])
        ):
            origem = b

        registros.append({
            "NOTA": nota,
            "LISTA": categoria,
            "MUNICIPIO": origem["MUNICIPIO_OBRA"],
            "REGIONAL": origem["REGIONAL_OBRA"],
            "LATITUDE": origem["LAT_OBRA"],
            "LONGITUDE": origem["LON_OBRA"],
            "STATUS_SAP": (
                b["SAP_NORM"] if tem_lev else ""
            ),
            "STATUS_LIST": (
                b["LIST_NORM"] if tem_lev else ""
            ),
            "CONTRATO": (
                b["CONTRATO_NORM"] if tem_lev else ""
            ),
            "COLUNA_M": (
                b["M_NORM"] if tem_lev else ""
            ),
            "COLUNA_N": (
                b["N_NORM"] if tem_lev else ""
            ),
            "PRIORIDADE_ORIGINAL": (
                b["PRIORIDADE_NORM"] if tem_lev else ""
            ),
            "DATA_ABERTURA": (
                b["DATA_ABERTURA_NORM"] if tem_lev else ""
            ),
            "DUPLICADA": (
                "SIM" if tem_san and tem_lev else "NÃO"
            ),
            "JA_EM_CAMPO": (
                "SIM" if ja_campo else "NÃO"
            ),
            "EQUIPE_CAMPO": " | ".join(responsaveis),
            "DATA_CAMPO": " | ".join(datas_campo),
            "PENDENTE_CONTAGEM": (
                "NÃO" if motivos else "SIM"
            ),
            "MOTIVO_EXCLUSAO": (
                " | ".join(motivos) if motivos else "-"
            ),
            "OCORRENCIAS_SANEAMENTO": (
                len(gs) if tem_san else 0
            ),
            "OCORRENCIAS_LEVANTAMENTO": (
                len(gl) if tem_lev else 0
            )
        })

    return pd.DataFrame(registros)


# ============================================================
# GEOGRAFIA
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    lat1 = np.radians(lat1)
    lon1 = np.radians(lon1)
    lat2 = np.radians(lat2)
    lon2 = np.radians(lon2)

    a = (
        np.sin((lat2 - lat1) / 2) ** 2
        + np.cos(lat1)
        * np.cos(lat2)
        * np.sin((lon2 - lon1) / 2) ** 2
    )

    return (
        2 * EARTH_KM
        * np.arcsin(np.sqrt(np.clip(a, 0, 1)))
    )


def atividades(categoria):
    if categoria == CAT_DUP:
        return ["SANEAMENTO", "LEVANTAMENTO"]

    if categoria == CAT_SAN:
        return ["SANEAMENTO"]

    return ["LEVANTAMENTO"]


def gerar_candidatos(obras, equipes, quantidade=3):
    saida = obras.copy().reset_index(drop=True)

    saida["EQUIPES_SANEAMENTO"] = "NAO APLICAVEL"
    saida["EQUIPES_LEVANTAMENTO"] = "NAO APLICAVEL"

    registros = []

    for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
        eq = equipes[
            equipes["TIPO_EQUIPE"].eq(tipo)
        ].reset_index(drop=True)

        if eq.empty:
            continue

        arvore = BallTree(
            np.radians(
                eq[
                    ["LAT_EQUIPE", "LON_EQUIPE"]
                ].to_numpy(float)
            ),
            metric="haversine"
        )

        indices_obras = [
            i for i, r in saida.iterrows()
            if (
                tipo in atividades(r["LISTA"])
                and pd.notna(r["LATITUDE"])
                and pd.notna(r["LONGITUDE"])
            )
        ]

        if not indices_obras:
            continue

        pontos = np.radians(
            saida.loc[
                indices_obras,
                ["LATITUDE", "LONGITUDE"]
            ].to_numpy(float)
        )

        distancias, indices = arvore.query(
            pontos,
            k=min(quantidade, len(eq))
        )

        for posicao, indice_obra in enumerate(
            indices_obras
        ):
            textos = []

            for j in range(indices.shape[1]):
                equipe = eq.iloc[
                    indices[posicao, j]
                ]

                km = float(
                    distancias[posicao, j] * EARTH_KM
                )

                registros.append({
                    "NOTA": saida.at[
                        indice_obra, "NOTA"
                    ],
                    "ATIVIDADE": tipo,
                    "ID_EQUIPE": equipe["ID_EQUIPE"],
                    "EQUIPE": equipe["EQUIPE"],
                    "CIDADE_BASE": equipe["CIDADE_BASE"],
                    "LAT_EQUIPE": equipe["LAT_EQUIPE"],
                    "LON_EQUIPE": equipe["LON_EQUIPE"],
                    "DISTANCIA_RETA_KM": round(km, 2)
                })

                textos.append(
                    f"{equipe['EQUIPE']} "
                    f"({equipe['CIDADE_BASE']}) "
                    f"- {km:.1f} km"
                )

            saida.at[
                indice_obra,
                f"EQUIPES_{tipo}"
            ] = " | ".join(textos)

    return saida, pd.DataFrame(registros)


# ============================================================
# PRIORIDADE
# ============================================================

def priorizar(obras, dias_media=14):
    df = obras.copy()

    datas = pd.to_datetime(
        df["DATA_ABERTURA"],
        errors="coerce",
        dayfirst=True
    )

    df["DIAS_ABERTURA"] = (
        pd.Timestamp.today().normalize() - datas
    ).dt.days

    def prioridade(r):
        p = norm(r["PRIORIDADE_ORIGINAL"])

        if any(
            x in p
            for x in ["URGENTE", "CRITICA", "ALTA"]
        ):
            return "ALTA"

        idade = r["DIAS_ABERTURA"]

        if pd.notna(idade):
            if idade >= dias_media * 4:
                return "ALTA"

            if idade >= dias_media:
                return "MEDIA"

        return "BAIXA"

    df["PRIORIDADE_PLANEJAMENTO"] = df.apply(
        prioridade,
        axis=1
    )

    return df


# ============================================================
# TAREFAS E PROGRAMACAO
# ============================================================

def gerar_tarefas(obras):
    registros = []

    for _, obra in obras.iterrows():
        for tipo in atividades(obra["LISTA"]):
            registro = obra.to_dict()
            registro["ATIVIDADE"] = tipo
            registros.append(registro)

    return pd.DataFrame(registros)


def gerar_datas(data_inicio, modo, periodos):
    inicio = pd.Timestamp(data_inicio)

    if modo == "POR DIA":
        fim = inicio + pd.Timedelta(
            days=periodos * 2 + 15
        )

    elif modo == "POR SEMANA":
        fim = inicio + pd.Timedelta(
            weeks=periodos
        )

    elif modo == "POR MES":
        fim = (
            inicio
            + pd.DateOffset(months=periodos)
            - pd.Timedelta(days=1)
        )

    else:
        fim = inicio + pd.Timedelta(days=365)

    datas = pd.date_range(
        inicio, fim, freq="D"
    )

    dias_uteis = [
        data.date()
        for data in datas
        if data.weekday() < 5
    ]

    if modo == "POR DIA":
        return dias_uteis[:periodos]

    return dias_uteis


def obter_periodo(data, modo):
    if modo == "POR SEMANA":
        iso = data.isocalendar()
        return f"{iso.year}-S{iso.week:02d}"

    if modo == "POR MES":
        return data.strftime("%Y-%m")

    return data.isoformat()


def programar(
    tarefas,
    candidatos,
    raio,
    data_inicio,
    modo,
    periodos,
    limite_periodo,
    limite_dia
):
    if tarefas.empty:
        return pd.DataFrame(), pd.DataFrame()

    ordem = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    tarefas = tarefas.copy()

    tarefas["_ORDEM"] = (
        tarefas["PRIORIDADE_PLANEJAMENTO"]
        .map(ordem)
        .fillna(3)
    )

    tarefas = tarefas.sort_values(
        ["_ORDEM", "MUNICIPIO", "NOTA"]
    )

    datas = gerar_datas(
        data_inicio, modo, periodos
    )

    candidatos_por_nota = defaultdict(list)

    for _, c in candidatos.iterrows():
        candidatos_por_nota[
            (c["NOTA"], c["ATIVIDADE"])
        ].append(c.to_dict())

    ocupacao_dia = defaultdict(int)
    ocupacao_periodo = defaultdict(int)

    resultados = []

    for _, tarefa in tarefas.iterrows():
        chave = (
            tarefa["NOTA"],
            tarefa["ATIVIDADE"]
        )

        alternativas = []

        for c in candidatos_por_nota.get(chave, []):
            km = float(c["DISTANCIA_RETA_KM"])

            if km > raio:
                continue

            for data in datas:
                periodo = obter_periodo(data, modo)

                chave_dia = (
                    c["ID_EQUIPE"], data
                )

                chave_periodo = (
                    c["ID_EQUIPE"], periodo
                )

                carga_dia = ocupacao_dia[chave_dia]
                carga_periodo = ocupacao_periodo[
                    chave_periodo
                ]

                if carga_dia >= limite_dia:
                    continue

                if carga_periodo >= limite_periodo:
                    continue

                alternativas.append({
                    "CANDIDATO": c,
                    "DATA": data,
                    "PERIODO": periodo,
                    "KM": km,
                    "CARGA_DIA": carga_dia,
                    "CARGA_PERIODO": carga_periodo
                })

                break

        registro = tarefa.drop(
            labels=["_ORDEM"]
        ).to_dict()

        if alternativas:
            melhor = min(
                alternativas,
                key=lambda x: (
                    x["CARGA_PERIODO"],
                    x["CARGA_DIA"],
                    x["KM"],
                    x["DATA"]
                )
            )

            c = melhor["CANDIDATO"]
            data = melhor["DATA"]
            periodo = melhor["PERIODO"]

            ocupacao_dia[
                (c["ID_EQUIPE"], data)
            ] += 1

            ocupacao_periodo[
                (c["ID_EQUIPE"], periodo)
            ] += 1

            iso = data.isocalendar()

            registro.update({
                "ID_EQUIPE": c["ID_EQUIPE"],
                "EQUIPE_PROGRAMADA": c["EQUIPE"],
                "CIDADE_BASE": c["CIDADE_BASE"],
                "DATA_PROGRAMADA": data.isoformat(),
                "DIA_SEMANA": DIAS_PT[
                    data.weekday()
                ],
                "DIA_MES": data.strftime(
                    "%d/%m/%Y"
                ),
                "SEMANA": (
                    f"{iso.year}-S{iso.week:02d}"
                ),
                "MES": data.strftime("%Y-%m"),
                "PERIODO": periodo,
                "DISTANCIA_REFERENCIA_KM": (
                    melhor["KM"]
                ),
                "TIPO_DISTANCIA": "LINHA RETA",
                "STATUS_PROGRAMACAO": "SUGESTAO"
            })

        else:
            registro.update({
                "ID_EQUIPE": "",
                "EQUIPE_PROGRAMADA": "NAO ALOCADA",
                "CIDADE_BASE": "",
                "DATA_PROGRAMADA": "",
                "DIA_SEMANA": "",
                "DIA_MES": "",
                "SEMANA": "",
                "MES": "",
                "PERIODO": "",
                "DISTANCIA_REFERENCIA_KM": np.nan,
                "TIPO_DISTANCIA": "",
                "STATUS_PROGRAMACAO": "SEM ALOCACAO"
            })

        resultados.append(registro)

    programacao = pd.DataFrame(resultados)

    alocadas = programacao[
        programacao["STATUS_PROGRAMACAO"].eq(
            "SUGESTAO"
        )
    ]

    carga = (
        alocadas.groupby(
            [
                "ATIVIDADE",
                "EQUIPE_PROGRAMADA",
                "PERIODO"
            ]
        )
        .size()
        .reset_index(name="TAREFAS")
    )

    return programacao, carga


# ============================================================
# MEDIA DE OBRAS POR EQUIPE
# ============================================================

def calcular_capacidade_media(
    obras,
    equipes,
    programacao
):
    resumo = []

    for tipo, descricao in [
        ("SANEAMENTO", "Saneamento"),
        ("LEVANTAMENTO", "Levantamento")
    ]:
        equipes_tipo = equipes[
            equipes["TIPO_EQUIPE"].eq(tipo)
        ]

        qtd_equipes = len(equipes_tipo)

        qtd_obras = int(
            obras["LISTA"].isin(
                [CAT_SAN, CAT_DUP]
                if tipo == "SANEAMENTO"
                else [CAT_LEV, CAT_DUP]
            ).sum()
        )

        media = (
            qtd_obras / qtd_equipes
            if qtd_equipes else 0
        )

        tarefas_tipo = (
            programacao[
                programacao["ATIVIDADE"].eq(tipo)
            ]
            if not programacao.empty
            else pd.DataFrame()
        )

        alocadas = (
            int(
                tarefas_tipo["STATUS_PROGRAMACAO"]
                .eq("SUGESTAO").sum()
            )
            if not tarefas_tipo.empty else 0
        )

        resumo.append({
            "ATIVIDADE": descricao,
            "EQUIPES": qtd_equipes,
            "OBRAS_DISPONIVEIS": qtd_obras,
            "MEDIA_OBRAS_POR_EQUIPE": round(media, 2),
            "TAREFAS_PROGRAMADAS": alocadas,
            "MEDIA_PROGRAMADA_POR_EQUIPE": round(
                alocadas / qtd_equipes, 2
            ) if qtd_equipes else 0
        })

    quadro = pd.DataFrame(resumo)

    contagem = (
        programacao[
            programacao["STATUS_PROGRAMACAO"].eq(
                "SUGESTAO"
            )
        ]
        .groupby("ID_EQUIPE")
        .size()
        .to_dict()
        if not programacao.empty else {}
    )

    detalhe = equipes[
        [
            "ID_EQUIPE",
            "TIPO_EQUIPE",
            "EQUIPE",
            "CIDADE_BASE"
        ]
    ].copy()

    detalhe["OBRAS_PROGRAMADAS"] = (
        detalhe["ID_EQUIPE"]
        .map(contagem)
        .fillna(0)
        .astype(int)
    )

    for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
        total = quadro.loc[
            quadro["ATIVIDADE"].eq(
                "Saneamento"
                if tipo == "SANEAMENTO"
                else "Levantamento"
            ),
            "MEDIA_OBRAS_POR_EQUIPE"
        ].iloc[0]

        detalhe.loc[
            detalhe["TIPO_EQUIPE"].eq(tipo),
            "MEDIA_DISPONIVEL_ATIVIDADE"
        ] = total

    return quadro, detalhe


# ============================================================
# SUPERPONTOS - DBSCAN DA LISTA CONTINUA
# ============================================================

def fundir_superpontos(df, raio_metros=50):
    if df.empty:
        return pd.DataFrame(), 0

    validas = df.dropna(
        subset=["LATITUDE", "LONGITUDE"]
    ).copy()

    invalidas = df[
        df["LATITUDE"].isna()
        | df["LONGITUDE"].isna()
    ].copy()

    saida = []

    if not validas.empty:
        coordenadas = np.radians(
            validas[
                ["LATITUDE", "LONGITUDE"]
            ].to_numpy(float)
        )

        modelo = DBSCAN(
            eps=raio_metros / 6371000,
            min_samples=1,
            metric="haversine",
            algorithm="ball_tree"
        ).fit(coordenadas)

        validas["CLUSTER_ID"] = modelo.labels_

        validas["CHAVE_GRUPO"] = (
            validas["CLUSTER_ID"].astype(str)
            + "|"
            + validas["ID_EQUIPE"].astype(str)
            + "|"
            + validas["DATA_PROGRAMADA"].astype(str)
            + "|"
            + validas["ATIVIDADE"].astype(str)
        )

        for _, grupo in validas.groupby(
            "CHAVE_GRUPO",
            sort=False
        ):
            base = grupo.iloc[0].copy()

            originais = grupo.to_dict("records")
            qtd = len(originais)

            base["_ORIGINAL_ROWS"] = originais

            base["LATITUDE"] = (
                grupo["LATITUDE"].mean()
            )

            base["LONGITUDE"] = (
                grupo["LONGITUDE"].mean()
            )

            base["SUPER_PONTO"] = (
                f"SIM ({qtd} un.)"
                if qtd > 1 else "NÃO"
            )

            base["NOTAS_AGRUPADAS"] = " | ".join(
                grupo["NOTA"].astype(str).tolist()
            )

            saida.append(base)

    if not invalidas.empty:
        invalidas["SUPER_PONTO"] = "NÃO"
        invalidas["NOTAS_AGRUPADAS"] = (
            invalidas["NOTA"].astype(str)
        )

        invalidas["_ORIGINAL_ROWS"] = [
            [r] for r in invalidas.to_dict("records")
        ]

        saida.extend(
            invalidas.to_dict("records")
        )

    final = pd.DataFrame(saida)

    final = final.drop(
        columns=["CLUSTER_ID", "CHAVE_GRUPO"],
        errors="ignore"
    )

    return final, len(df) - len(final)


# ============================================================
# OSRM
# ============================================================

@st.cache_data(
    ttl=86400,
    max_entries=3000,
    show_spinner=False
)
def consultar_osrm(
    lat1,
    lon1,
    lat2,
    lon2,
    servidor,
    timeout_s=4
):
    url = (
        servidor.rstrip("/")
        + "/route/v1/driving/"
        + f"{lon1:.6f},{lat1:.6f};"
        + f"{lon2:.6f},{lat2:.6f}"
    )

    try:
        resposta = requests.get(
            url,
            params={
                "overview": "full",
                "geometries": "geojson",
                "steps": "false"
            },
            timeout=(2, timeout_s),
            headers={
                "User-Agent": "NIP-Roteirizador/1.0"
            }
        )

        resposta.raise_for_status()
        dados = resposta.json()

        if dados.get("code") != "Ok":
            return {"OK": False}

        rotas = dados.get("routes", [])

        if not rotas:
            return {"OK": False}

        rota = rotas[0]

        geometria = (
            rota.get("geometry", {})
            .get("coordinates", [])
        )

        if len(geometria) < 2:
            return {"OK": False}

        return {
            "OK": True,
            "KM": round(
                float(rota["distance"]) / 1000, 2
            ),
            "MIN": round(
                float(rota["duration"]) / 60, 1
            ),
            "GEOMETRIA": geometria
        }

    except (
        requests.RequestException,
        ValueError,
        TypeError,
        KeyError
    ):
        return {"OK": False}


def preparar_trechos(programacao, equipes):
    if programacao.empty:
        return pd.DataFrame()

    linhas = []

    validas = programacao[
        programacao["STATUS_PROGRAMACAO"].eq(
            "SUGESTAO"
        )
    ].dropna(
        subset=["LATITUDE", "LONGITUDE"]
    )

    for (id_equipe, data), grupo in validas.groupby(
        ["ID_EQUIPE", "DATA_PROGRAMADA"]
    ):
        eq = equipes[
            equipes["ID_EQUIPE"].eq(id_equipe)
        ]

        if eq.empty:
            continue

        equipe = eq.iloc[0]

        lat_atual = float(
            equipe["LAT_EQUIPE"]
        )

        lon_atual = float(
            equipe["LON_EQUIPE"]
        )

        restantes = grupo.to_dict("records")
        ordem = 0

        while restantes:
            distancias = [
                float(
                    haversine(
                        lat_atual,
                        lon_atual,
                        float(r["LATITUDE"]),
                        float(r["LONGITUDE"])
                    )
                )
                for r in restantes
            ]

            indice = int(np.argmin(distancias))
            obra = restantes.pop(indice)
            ordem += 1

            linhas.append({
                "NOTA": obra["NOTA"],
                "ATIVIDADE": obra["ATIVIDADE"],
                "EQUIPE": obra["EQUIPE_PROGRAMADA"],
                "ID_EQUIPE": id_equipe,
                "DATA": data,
                "ORDEM": ordem,
                "LAT_ORIGEM": lat_atual,
                "LON_ORIGEM": lon_atual,
                "LAT_DESTINO": float(
                    obra["LATITUDE"]
                ),
                "LON_DESTINO": float(
                    obra["LONGITUDE"]
                ),
                "KM_RETA": round(
                    distancias[indice], 2
                )
            })

            lat_atual = float(obra["LATITUDE"])
            lon_atual = float(obra["LONGITUDE"])

    return pd.DataFrame(linhas)


def processar_lote_osrm(
    trechos,
    servidor,
    inicio,
    tamanho_lote,
    cronometro_chave
):
    novos = []

    fim = min(
        inicio + tamanho_lote,
        len(trechos)
    )

    progresso = st.progress(
        inicio / max(1, len(trechos))
    )

    cronometro_visual = st.empty()

    for indice in range(inicio, fim):
        r = trechos.iloc[indice]

        rota = consultar_osrm(
            round(float(r["LAT_ORIGEM"]), 6),
            round(float(r["LON_ORIGEM"]), 6),
            round(float(r["LAT_DESTINO"]), 6),
            round(float(r["LON_DESTINO"]), 6),
            servidor
        )

        registro = r.to_dict()

        if rota["OK"]:
            registro["TIPO_TRAJETO"] = "OSRM"
            registro["KM_RODOVIARIO"] = rota["KM"]
            registro["TEMPO_MIN"] = rota["MIN"]
            registro["GEOMETRIA"] = (
                rota["GEOMETRIA"]
            )
        else:
            registro["TIPO_TRAJETO"] = (
                "ESTIMATIVA - LINHA RETA"
            )
            registro["KM_RODOVIARIO"] = np.nan
            registro["TEMPO_MIN"] = np.nan
            registro["GEOMETRIA"] = [
                [
                    registro["LON_ORIGEM"],
                    registro["LAT_ORIGEM"]
                ],
                [
                    registro["LON_DESTINO"],
                    registro["LAT_DESTINO"]
                ]
            ]

        novos.append(registro)

        processados = indice + 1

        progresso.progress(
            processados / len(trechos)
        )

        atualizar_cronometro(
            cronometro_chave,
            processados,
            len(trechos),
            cronometro_visual
        )

    return novos, fim


# ============================================================
# POPUP
# ============================================================

def popup_html(obra):
    def esc(campo):
        return html.escape(
            texto_seguro(obra.get(campo, ""))
        )

    linhas = [
        ("Nota", esc("NOTA")),
        ("Municipio", esc("MUNICIPIO")),
        ("Tipo", esc("LISTA"))
    ]

    san = esc("EQUIPES_SANEAMENTO")
    lev = esc("EQUIPES_LEVANTAMENTO")

    if san and san != "NAO APLICAVEL":
        linhas.append(
            ("Equipes Saneamento", san)
        )

    if lev and lev != "NAO APLICAVEL":
        linhas.append(
            ("Equipes Levantamento", lev)
        )

    for campo, rotulo in [
        ("REGIONAL", "Regional"),
        ("PRIORIDADE_PLANEJAMENTO", "Prioridade"),
        ("EQUIPE_PROGRAMADA", "Equipe Programada"),
        ("DATA_PROGRAMADA", "Data"),
        ("SUPER_PONTO", "Superponto"),
        ("NOTAS_AGRUPADAS", "Notas agrupadas")
    ]:
        valor = esc(campo)

        if valor and valor.lower() != "nan":
            linhas.append(
                (rotulo, valor)
            )

    corpo = "".join(
        f"""
        <tr>
            <td style="
                padding:5px;
                width:115px;
                font-weight:bold;
                vertical-align:top;
            ">{rotulo}</td>
            <td style="padding:5px;">{valor}</td>
        </tr>
        """
        for rotulo, valor in linhas
    )

    cor = CORES.get(
        obra.get("LISTA"),
        "#0D256C"
    )

    return f"""
    <div style="
        font-family:Arial,sans-serif;
        width:340px;
        font-size:12px;
        color:#1f2937;
    ">
        <div style="
            background:{cor};
            color:white;
            padding:11px;
            font-weight:bold;
            border-radius:7px 7px 0 0;
        ">
            INFORMACOES DA OBRA
        </div>
        <table style="
            width:100%;
            border-collapse:collapse;
            background:white;
        ">
            {corpo}
        </table>
    </div>
    """


# ============================================================
# EXPORTACAO KML
# ============================================================

def gerar_kml(
    obras,
    superpontos=None,
    trajetos=None,
    nome="NIP - Obras Pendentes"
):
    namespace = "http://www.opengis.net/kml/2.2"
    ET.register_namespace("", namespace)

    def el(pai, tag, texto=None):
        item = ET.SubElement(
            pai,
            f"{{{namespace}}}{tag}"
        )

        if texto is not None:
            item.text = str(texto)

        return item

    raiz = ET.Element(
        f"{{{namespace}}}kml"
    )

    documento = el(raiz, "Document")
    el(documento, "name", nome)

    estilos = {
        CAT_SAN: (
            "saneamento",
            "http://maps.google.com/mapfiles/"
            "kml/paddle/blu-blank.png"
        ),
        CAT_LEV: (
            "levantamento",
            "http://maps.google.com/mapfiles/"
            "kml/paddle/grn-blank.png"
        ),
        CAT_DUP: (
            "duplicadas",
            "http://maps.google.com/mapfiles/"
            "kml/paddle/purple-blank.png"
        )
    }

    for categoria, (id_estilo, url) in estilos.items():
        estilo = el(documento, "Style")
        estilo.set("id", id_estilo)

        icone_estilo = el(
            estilo, "IconStyle"
        )

        el(icone_estilo, "scale", "1.2")

        icone = el(
            icone_estilo, "Icon"
        )

        el(icone, "href", url)

        rotulo = el(
            estilo, "LabelStyle"
        )

        el(rotulo, "scale", "0")

    for categoria in CATEGORIAS:
        pasta = el(documento, "Folder")
        el(pasta, "name", categoria)

        dados = obras[
            obras["LISTA"].eq(categoria)
        ].dropna(
            subset=["LATITUDE", "LONGITUDE"]
        )

        for _, obra in dados.iterrows():
            ponto = el(pasta, "Placemark")

            el(
                ponto, "name",
                f"NOTA {obra['NOTA']}"
            )

            el(
                ponto, "styleUrl",
                "#" + estilos[categoria][0]
            )

            el(
                ponto,
                "description",
                popup_html(obra)
            )

            geometria = el(
                ponto, "Point"
            )

            el(
                geometria,
                "coordinates",
                (
                    f"{float(obra['LONGITUDE'])},"
                    f"{float(obra['LATITUDE'])},0"
                )
            )

    if superpontos is not None and not superpontos.empty:
        pasta = el(documento, "Folder")

        el(
            pasta,
            "name",
            "SUPERPONTOS"
        )

        for _, r in superpontos.iterrows():
            if not str(
                r.get("SUPER_PONTO", "")
            ).startswith("SIM"):
                continue

            if (
                pd.isna(r["LATITUDE"])
                or pd.isna(r["LONGITUDE"])
            ):
                continue

            pm = el(pasta, "Placemark")

            el(
                pm,
                "name",
                f"SUPERPONTO - {r['SUPER_PONTO']}"
            )

            el(
                pm,
                "styleUrl",
                "#" + estilos.get(
                    r["LISTA"],
                    estilos[CAT_DUP]
                )[0]
            )

            el(
                pm,
                "description",
                popup_html(r)
            )

            geometria = el(pm, "Point")

            el(
                geometria,
                "coordinates",
                (
                    f"{float(r['LONGITUDE'])},"
                    f"{float(r['LATITUDE'])},0"
                )
            )

    if trajetos is not None and not trajetos.empty:
        pasta_rotas = el(
            documento, "Folder"
        )

        el(
            pasta_rotas,
            "name",
            "TRAJETOS DAS EQUIPES"
        )

        for _, trecho in trajetos.iterrows():
            geometria = trecho.get(
                "GEOMETRIA"
            )

            if not isinstance(geometria, list):
                continue

            if len(geometria) < 2:
                continue

            pm = el(
                pasta_rotas, "Placemark"
            )

            tipo = trecho.get(
                "TIPO_TRAJETO",
                "LINHA RETA"
            )

            el(
                pm,
                "name",
                (
                    f"{trecho['EQUIPE']} - "
                    f"{trecho['ORDEM']} - {tipo}"
                )
            )

            el(
                pm,
                "description",
                (
                    f"Equipe: {trecho['EQUIPE']}\n"
                    f"Nota: {trecho['NOTA']}\n"
                    f"Tipo: {tipo}\n"
                    f"KM rodoviario: "
                    f"{texto_seguro(trecho.get('KM_RODOVIARIO'))}"
                )
            )

            linha = el(pm, "LineString")
            el(linha, "tessellate", "1")

            el(
                linha,
                "coordinates",
                " ".join(
                    f"{pt[0]},{pt[1]},0"
                    for pt in geometria
                )
            )

            estilo = el(pm, "Style")
            linha_estilo = el(
                estilo, "LineStyle"
            )

            cor = (
                "ffff6600"
                if tipo == "OSRM"
                else "ff999999"
            )

            el(linha_estilo, "color", cor)
            el(linha_estilo, "width", "4")

    return ET.tostring(
        raiz,
        encoding="utf-8",
        xml_declaration=True
    )


# ============================================================
# EXPORTACAO EXCEL
# ============================================================

def limpar_excel(df):
    if df is None:
        return pd.DataFrame()

    saida = df.copy()

    saida = saida.drop(
        columns=[
            "_ORIGINAL_ROWS",
            "GEOMETRIA",
            "ROTA_GEOMETRIA",
            "CLUSTER_ID",
            "CHAVE_GRUPO"
        ],
        errors="ignore"
    )

    for coluna in saida.columns:
        if saida[coluna].dtype == "object":
            saida[coluna] = saida[coluna].map(
                lambda x: (
                    " | ".join(map(str, x))
                    if isinstance(x, list)
                    else str(x)
                    if isinstance(x, dict)
                    else x
                )
            )

    return saida


def formatar_planilha(writer, aba, df):
    ws = writer.sheets[aba]

    cabecalho = PatternFill(
        "solid",
        fgColor="002060"
    )

    super_cor = PatternFill(
        "solid",
        fgColor="FCE4D6"
    )

    fonte_cabecalho = Font(
        name="Calibri",
        color="FFFFFF",
        bold=True
    )

    fonte_prioridade = Font(
        name="Calibri",
        color="FF0000",
        bold=True
    )

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions

    for celula in ws[1]:
        celula.fill = cabecalho
        celula.font = fonte_cabecalho

        celula.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 26

    indice_sp = (
        df.columns.get_loc("SUPER_PONTO")
        if "SUPER_PONTO" in df.columns
        else None
    )

    indice_prioridade = (
        df.columns.get_loc(
            "PRIORIDADE_PLANEJAMENTO"
        )
        if "PRIORIDADE_PLANEJAMENTO"
        in df.columns
        else None
    )

    for linha in ws.iter_rows(min_row=2):
        superponto = (
            str(linha[indice_sp].value).startswith("SIM")
            if indice_sp is not None
            else False
        )

        prioridade = (
            str(
                linha[indice_prioridade].value
            ).upper() == "ALTA"
            if indice_prioridade is not None
            else False
        )

        for celula in linha:
            if superponto:
                celula.fill = super_cor

            if prioridade:
                celula.font = fonte_prioridade

    for coluna in ws.columns:
        letra = coluna[0].column_letter

        tamanhos = [
            len(str(c.value or ""))
            for c in list(coluna)[:120]
        ]

        ws.column_dimensions[letra].width = min(
            60,
            max(13, max(tamanhos, default=10) + 2)
        )


def gerar_excel(planilhas):
    memoria = io.BytesIO()

    with pd.ExcelWriter(
        memoria,
        engine="openpyxl"
    ) as writer:
        for nome, df in planilhas.items():
            aba = nome[:31]
            tabela = limpar_excel(df)

            tabela.to_excel(
                writer,
                sheet_name=aba,
                index=False
            )

            formatar_planilha(
                writer,
                aba,
                tabela
            )

    return memoria.getvalue()


def gerar_zip_excel_por_equipe(programacao):
    memoria = io.BytesIO()

    with zipfile.ZipFile(
        memoria,
        "w",
        zipfile.ZIP_DEFLATED
    ) as arquivo_zip:

        if not programacao.empty:
            for (tipo, equipe), grupo in programacao.groupby(
                ["ATIVIDADE", "EQUIPE_PROGRAMADA"]
            ):
                if equipe == "NAO ALOCADA":
                    continue

                nome_seguro = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    norm(equipe)
                )[:55]

                arquivo_zip.writestr(
                    (
                        f"{tipo}/"
                        f"Rota_{nome_seguro}.xlsx"
                    ),
                    gerar_excel({
                        "Obras Roteirizadas": grupo
                    })
                )

    return memoria.getvalue()


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.markdown(
    '<div class="nip-title">'
    '📍 NIP | Analise Cruzada e Planejamento'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="nip-subtitle">'
    'Obras Pendentes | Programacao | Superpontos | '
    'Auditoria | Excel e KML'
    '</div>',
    unsafe_allow_html=True
)

with st.expander(
    "📘 Regras de contagem e exclusao"
):
    st.markdown("""
    **STATUS SAP:** somente FINL e CANC são excluídos.

    **STATUS LIST:** são aceitos `0`,
    `Em levantamento` e `Correção de levantamento`.

    **Contrato:** `0` ou
    `NIP GLOBAL LTDA - EQTL MARANHÃO`.

    **Colunas M e N:** precisam estar iguais a zero.

    **Nova regra B/C:** quando B contém um responsável
    válido, diferente de `SEM LEVANTADOR`, e C contém
    uma data válida, a nota fica fora da contagem
    por já estar encaminhada ao campo.

    **Auditoria:** a nota continua registrada e
    pode apresentar múltiplos motivos de exclusão.

    **Duplicadas:** uma nota presente nas duas bases
    é contada uma vez, mas pode gerar duas tarefas,
    uma para cada atividade.

    **Distâncias:** o planejamento inicial utiliza
    distância geográfica. OSRM é opcional para
    traçado de ruas no KML.
    """)


# ============================================================
# UPLOAD
# ============================================================

secao("📂 Importacao das Bases")

c1, c2, c3 = st.columns(3)

with c1:
    arquivo_san = st.file_uploader(
        "BASE_SANEAMENTO",
        type=["xlsx", "csv"],
        key="san"
    )

with c2:
    arquivo_lev = st.file_uploader(
        "BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"],
        key="lev"
    )

with c3:
    arquivo_equipes = st.file_uploader(
        "LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"],
        key="equipes"
    )

if not all([
    arquivo_san,
    arquivo_lev,
    arquivo_equipes
]):
    st.info(
        "Envie os tres arquivos obrigatorios "
        "para iniciar a analise."
    )
    st.stop()

assinatura = tuple(
    hashlib.sha256(
        arquivo.getvalue()
    ).hexdigest()
    for arquivo in [
        arquivo_san,
        arquivo_lev,
        arquivo_equipes
    ]
)

if st.button(
    "🚀 PROCESSAR BASES",
    type="primary",
    use_container_width=True
):
    inicio = time.perf_counter()

    try:
        with st.spinner("Processando bases..."):
            san, aba_san = ler_saneamento(
                arquivo_san
            )

            lev, info_lev = ler_levantamento(
                arquivo_lev
            )

            equipes = ler_equipes(
                arquivo_equipes
            )

            auditoria = consolidar(
                san, lev
            )

            if auditoria.empty:
                raise ValueError(
                    "Nenhuma nota valida encontrada."
                )

            st.session_state["nip_base"] = {
                "assinatura": assinatura,
                "auditoria": auditoria,
                "equipes": equipes,
                "aba_san": aba_san,
                "info_lev": info_lev,
                "qtd_san": len(san),
                "qtd_lev": len(lev),
                "tempo_processamento": (
                    time.perf_counter() - inicio
                )
            }

            st.session_state.pop(
                "nip_osrm", None
            )

        st.success(
            "Bases processadas em "
            + formatar_tempo(
                time.perf_counter() - inicio
            )
        )

    except Exception as erro:
        st.error(
            f"Erro no processamento: {erro}"
        )

if "nip_base" not in st.session_state:
    st.stop()

dados = st.session_state["nip_base"]

if dados["assinatura"] != assinatura:
    st.warning(
        "Os arquivos foram alterados. "
        "Clique novamente em PROCESSAR BASES."
    )
    st.stop()

auditoria = dados["auditoria"]
equipes = dados["equipes"]

st.caption(
    f"Saneamento: aba {dados['aba_san']} | "
    f"Levantamento: aba "
    f"{dados['info_lev']['ABA']} | "
    f"Tempo de leitura: "
    f"{formatar_tempo(dados['tempo_processamento'])}"
)

with st.expander(
    "🔎 Conferencia das colunas B e C"
):
    info = dados["info_lev"]

    st.write(
        f"**Coluna B identificada:** "
        f"{info['COLUNA_B']}"
    )

    st.write(
        f"**Coluna C identificada:** "
        f"{info['COLUNA_C']}"
    )

    st.write(
        f"**Registros com equipe e data:** "
        f"{info['JA_EM_CAMPO_LINHAS']}"
    )

    st.caption(
        "A contagem final e por nota unica. "
        "O total de linhas acima pode ser maior "
        "que o numero de notas excluidas."
    )


# ============================================================
# PARAMETROS
# ============================================================

with st.sidebar:
    st.header("⚙️ Planejamento")

    raio_equipes = st.select_slider(
        "Raio de referencia das equipes (km)",
        options=list(range(200, 501, 25)),
        value=200
    )

    qtd_equipes = st.slider(
        "Equipes proximas por atividade",
        1, 5, 3
    )

    raio_super = st.slider(
        "Raio Superponto (metros)",
        10, 500, 50, 10
    )

    dias_prioridade = st.number_input(
        "Dias para prioridade media",
        min_value=1,
        value=14
    )

    st.divider()
    st.subheader("📅 Programacao")

    modo = st.selectbox(
        "Modo",
        [
            "POR QUANTIDADE",
            "POR DIA",
            "POR SEMANA",
            "POR MES"
        ]
    )

    limite_periodo = st.number_input(
        "Quantidade por equipe no periodo",
        min_value=1,
        value=10
    )

    limite_dia = st.number_input(
        "Limite diario por equipe",
        min_value=1,
        value=8
    )

    periodos = st.number_input(
        "Quantidade de periodos",
        min_value=1,
        max_value=52,
        value=5
    )

    data_inicio = st.date_input(
        "Data inicial",
        value=datetime.today().date()
    )

    st.divider()
    st.subheader("🛣️ OSRM")

    usar_osrm = st.checkbox(
        "Tracado real por ruas no KML",
        value=False
    )

    servidor_osrm = st.text_input(
        "Endpoint OSRM",
        value=OSRM_DEFAULT
    )

    lote_osrm = st.slider(
        "Consultas OSRM por lote",
        1, 8, 5
    )


# ============================================================
# INDICADORES E CONJUNTOS EXCLUIDOS
# ============================================================

obras = auditoria[
    auditoria["PENDENTE_CONTAGEM"].eq("SIM")
].copy()

duplicadas_total = auditoria[
    auditoria["DUPLICADA"].eq("SIM")
].copy()

duplicadas_pendentes = obras[
    obras["DUPLICADA"].eq("SIM")
].copy()

obras_finl = auditoria[
    auditoria["MOTIVO_EXCLUSAO"]
    .astype(str)
    .str.contains(r"(^|\s\|\s)SAP FINL($|\s\|\s)", regex=True)
].copy()

obras_canc = auditoria[
    auditoria["MOTIVO_EXCLUSAO"]
    .astype(str)
    .str.contains(r"(^|\s\|\s)SAP CANC($|\s\|\s)", regex=True)
].copy()

obras_campo = auditoria[
    auditoria["JA_EM_CAMPO"].eq("SIM")
].copy()

excluidas = auditoria[
    auditoria["PENDENTE_CONTAGEM"].eq("NÃO")
].copy()

obras = priorizar(
    obras,
    dias_prioridade
)

secao("📊 Indicadores Operacionais")

mostrar_cards([
    (
        "Saneamento",
        int(obras["LISTA"].eq(CAT_SAN).sum()),
        "#2563EB",
        "Exclusivas"
    ),
    (
        "Levantamento",
        int(obras["LISTA"].eq(CAT_LEV).sum()),
        "#16A34A",
        "Exclusivas"
    ),
    (
        "Duplicadas",
        len(duplicadas_pendentes),
        "#9333EA",
        "Pendentes nas duas bases"
    ),
    (
        "Total Pendente",
        len(obras),
        "#0D256C",
        "Notas unicas"
    ),
    (
        "Excluidas",
        len(excluidas),
        "#DC2626",
        "Fora da contagem"
    )
])

mostrar_cards([
    (
        "SAP FINL",
        len(obras_finl),
        "#64748B",
        "Notas com FINL"
    ),
    (
        "SAP CANC",
        len(obras_canc),
        "#F97316",
        "Notas com CANC"
    ),
    (
        "Ja Em Campo",
        len(obras_campo),
        "#0891B2",
        "Equipe + data nas colunas B/C"
    ),
    (
        "Duplicadas Totais",
        len(duplicadas_total),
        "#9333EA",
        "Incluindo excluidas"
    )
])


# ============================================================
# FILTROS
# ============================================================

secao("🔎 Filtrar Obras")

c1, c2, c3 = st.columns(3)

regionais = sorted(
    obras["REGIONAL"]
    .dropna()
    .astype(str)
    .unique()
)

municipios = sorted(
    obras["MUNICIPIO"]
    .dropna()
    .astype(str)
    .unique()
)

with c1:
    filtro_regional = st.multiselect(
        "Regional", regionais
    )

with c2:
    filtro_municipio = st.multiselect(
        "Municipio", municipios
    )

with c3:
    filtro_prioridade = st.multiselect(
        "Prioridade",
        ["ALTA", "MEDIA", "BAIXA"]
    )

pesquisa = st.text_input(
    "Pesquisar nota"
)

view = obras.copy()

if filtro_regional:
    view = view[
        view["REGIONAL"].isin(filtro_regional)
    ]

if filtro_municipio:
    view = view[
        view["MUNICIPIO"].isin(filtro_municipio)
    ]

if filtro_prioridade:
    view = view[
        view["PRIORIDADE_PLANEJAMENTO"]
        .isin(filtro_prioridade)
    ]

if pesquisa.strip():
    view = view[
        view["NOTA"]
        .astype(str)
        .str.contains(
            re.escape(pesquisa.strip()),
            case=False,
            na=False
        )
    ]


# ============================================================
# EQUIPES E PROGRAMACAO
# ============================================================

view, candidatos = gerar_candidatos(
    view,
    equipes,
    qtd_equipes
)

tarefas = gerar_tarefas(view)

programacao, carga = programar(
    tarefas,
    candidatos,
    raio_equipes,
    data_inicio,
    modo,
    int(periodos),
    int(limite_periodo),
    int(limite_dia)
)

quadro_capacidade, detalhe_capacidade = (
    calcular_capacidade_media(
        view,
        equipes,
        programacao
    )
)

secao("👥 Capacidade e Media de Obras por Equipe")

c1, c2 = st.columns(2)

san_resumo = quadro_capacidade[
    quadro_capacidade["ATIVIDADE"].eq(
        "Saneamento"
    )
].iloc[0]

lev_resumo = quadro_capacidade[
    quadro_capacidade["ATIVIDADE"].eq(
        "Levantamento"
    )
].iloc[0]

with c1:
    st.markdown(
        card(
            "Media por equipe - Saneamento",
            f"{san_resumo['MEDIA_OBRAS_POR_EQUIPE']:.2f}",
            "#2563EB",
            (
                f"{san_resumo['OBRAS_DISPONIVEIS']} obras / "
                f"{san_resumo['EQUIPES']} equipes"
            )
        ),
        unsafe_allow_html=True
    )

with c2:
    st.markdown(
        card(
            "Media por equipe - Levantamento",
            f"{lev_resumo['MEDIA_OBRAS_POR_EQUIPE']:.2f}",
            "#16A34A",
            (
                f"{lev_resumo['OBRAS_DISPONIVEIS']} obras / "
                f"{lev_resumo['EQUIPES']} equipes"
            )
        ),
        unsafe_allow_html=True
    )

st.dataframe(
    quadro_capacidade,
    hide_index=True,
    use_container_width=True
)

with st.expander(
    "📋 Distribuicao detalhada por equipe"
):
    st.dataframe(
        detalhe_capacidade,
        hide_index=True,
        use_container_width=True
    )

st.caption(
    "A media de obras disponiveis e uma divisao "
    "matematica entre tarefas e equipes. "
    "Nao representa capacidade real de execucao "
    "nem considera todas as restricoes de deslocamento."
)


# ============================================================
# SUPERPONTOS
# ============================================================

if not programacao.empty:
    alocadas = programacao[
        programacao["STATUS_PROGRAMACAO"].eq(
            "SUGESTAO"
        )
    ]

    superpontos, qtd_fundidas = fundir_superpontos(
        alocadas,
        raio_super
    )

else:
    superpontos = pd.DataFrame()
    qtd_fundidas = 0


# ============================================================
# OSRM E CRONOMETRO
# ============================================================

trechos = preparar_trechos(
    programacao,
    equipes
)

assinatura_rotas = hashlib.sha256(
    (
        str(
            trechos[
                ["NOTA", "ATIVIDADE", "DATA", "ID_EQUIPE"]
            ].to_json()
        )
        if not trechos.empty
        else ""
    ).encode()
    + servidor_osrm.encode()
).hexdigest()

if (
    "nip_osrm" not in st.session_state
    or st.session_state["nip_osrm"].get(
        "assinatura"
    ) != assinatura_rotas
):
    st.session_state["nip_osrm"] = {
        "assinatura": assinatura_rotas,
        "cursor": 0,
        "registros": [],
        "segundos_acumulados": 0.0
    }

estado_osrm = st.session_state["nip_osrm"]

if usar_osrm and not trechos.empty:
    secao("⏱️ Cronometro do Roteamento")

    total = len(trechos)
    processados = estado_osrm["cursor"]

    st.progress(
        processados / total
    )

    decorrido_acumulado = estado_osrm.get(
        "segundos_acumulados", 0.0
    )

    restante = (
        decorrido_acumulado / processados
        * (total - processados)
        if processados > 0
        else None
    )

    mostrar_cards([
        (
            "Tempo Decorrido",
            formatar_tempo(decorrido_acumulado),
            "#0D256C",
            "Consultas OSRM"
        ),
        (
            "Tempo Restante",
            formatar_tempo(restante)
            if restante is not None
            else "Calculando...",
            "#16A34A",
            "Estimativa"
        ),
        (
            "Trechos Processados",
            f"{processados}/{total}",
            "#9333EA",
            "Progresso geral"
        )
    ])

    if processados < total:
        if st.button(
            "🛣️ PROCESSAR PROXIMO LOTE OSRM",
            type="primary",
            use_container_width=True
        ):
            inicio_lote = time.monotonic()

            iniciar_cronometro("cronometro_lote")

            novos, proximo = processar_lote_osrm(
                trechos,
                servidor_osrm,
                processados,
                lote_osrm,
                "cronometro_lote"
            )

            estado_osrm["registros"].extend(novos)
            estado_osrm["cursor"] = proximo

            estado_osrm["segundos_acumulados"] += (
                time.monotonic() - inicio_lote
            )

            st.session_state["nip_osrm"] = (
                estado_osrm
            )

            st.rerun()

    else:
        st.success(
            "Todos os trechos foram processados."
        )

trajetos = pd.DataFrame(
    estado_osrm["registros"]
)

trajetos_exportar = (
    trajetos
    if usar_osrm
    else pd.DataFrame()
)


# ============================================================
# ABAS
# ============================================================

abas = st.tabs([
    "🟣 SANEAMENTO",
    "🟢 LEVANTAMENTO",
    "🔵 DUPLICADAS",
    "📅 PROGRAMAÇÃO",
    "👥 CAPACIDADE",
    "🏢 SUPERPONTOS",
    "⚠️ AUDITORIA",
    "🗺️ MAPA"
])

for aba, categoria in zip(
    abas[:3],
    CATEGORIAS
):
    with aba:
        tabela = view[
            view["LISTA"].eq(categoria)
        ]

        st.subheader(
            f"{categoria} - {len(tabela)} obras"
        )

        st.dataframe(
            tabela,
            hide_index=True,
            use_container_width=True
        )


with abas[3]:
    st.subheader("📅 Programacao")

    st.dataframe(
        programacao,
        hide_index=True,
        use_container_width=True
    )

    st.subheader("Carga por equipe")

    st.dataframe(
        carga,
        hide_index=True,
        use_container_width=True
    )


with abas[4]:
    st.subheader("👥 Capacidade por Equipe")

    st.dataframe(
        detalhe_capacidade,
        hide_index=True,
        use_container_width=True
    )

    if not detalhe_capacidade.empty:
        grafico = detalhe_capacidade.set_index(
            "EQUIPE"
        )["OBRAS_PROGRAMADAS"]

        st.bar_chart(grafico)


with abas[5]:
    st.subheader(
        "🏢 Superpontos"
    )

    if not superpontos.empty:
        st.dataframe(
            limpar_excel(superpontos),
            hide_index=True,
            use_container_width=True
        )


with abas[6]:
    subabas = st.tabs([
        "Excluidas",
        "Ja em campo",
        "SAP FINL",
        "SAP CANC",
        "Duplicadas totais"
    ])

    for subaba, tabela in zip(
        subabas,
        [
            excluidas,
            obras_campo,
            obras_finl,
            obras_canc,
            duplicadas_total
        ]
    ):
        with subaba:
            st.dataframe(
                tabela,
                hide_index=True,
                use_container_width=True
            )


with abas[7]:
    st.subheader("🗺️ Mapa Operacional")

    if st.checkbox("Carregar mapa"):
        geos = view.dropna(
            subset=["LATITUDE", "LONGITUDE"]
        )

        if geos.empty:
            st.warning(
                "Nenhuma coordenada valida."
            )

        else:
            mapa = folium.Map(
                location=[
                    float(geos["LATITUDE"].mean()),
                    float(geos["LONGITUDE"].mean())
                ],
                zoom_start=7
            )

            for categoria in CATEGORIAS:
                camada = folium.FeatureGroup(
                    name=categoria
                )

                cluster = MarkerCluster().add_to(
                    camada
                )

                subset = geos[
                    geos["LISTA"].eq(categoria)
                ]

                for _, obra in subset.iterrows():
                    folium.Marker(
                        [
                            float(obra["LATITUDE"]),
                            float(obra["LONGITUDE"])
                        ],
                        popup=folium.Popup(
                            popup_html(obra),
                            max_width=420
                        ),
                        icon=folium.Icon(
                            color=CORES_FOLIUM[
                                categoria
                            ]
                        )
                    ).add_to(cluster)

                camada.add_to(mapa)

            folium.LayerControl().add_to(mapa)

            st_folium(
                mapa,
                use_container_width=True,
                height=600
            )


# ============================================================
# EXPORTACOES
# ============================================================

secao("📥 Exportacao de Planilhas e KML")

resumo_geral = pd.DataFrame([
    ["Saneamento", int(obras["LISTA"].eq(CAT_SAN).sum())],
    ["Levantamento", int(obras["LISTA"].eq(CAT_LEV).sum())],
    ["Duplicadas pendentes", len(duplicadas_pendentes)],
    ["Total pendente", len(obras)],
    ["Total excluidas", len(excluidas)],
    ["Ja em campo", len(obras_campo)],
    ["SAP FINL", len(obras_finl)],
    ["SAP CANC", len(obras_canc)]
], columns=["INDICADOR", "VALOR"])

excel_geral = gerar_excel({
    "RESUMO": resumo_geral,
    "CONSOLIDADO": obras,
    "SANEAMENTO": obras[
        obras["LISTA"].eq(CAT_SAN)
    ],
    "LEVANTAMENTO": obras[
        obras["LISTA"].eq(CAT_LEV)
    ],
    "DUPLICADAS": duplicadas_pendentes,
    "DUPLICADAS TOTAL": duplicadas_total,
    "SAP FINL": obras_finl,
    "SAP CANC": obras_canc,
    "JA EM CAMPO": obras_campo,
    "AUDITORIA": auditoria,
    "PROGRAMACAO": programacao,
    "CAPACIDADE EQUIPES": detalhe_capacidade,
    "MEDIAS": quadro_capacidade,
    "SUPERPONTOS": superpontos,
    "ROTAS OSRM": trajetos_exportar
})

excel_duplicadas = gerar_excel({
    "DUPLICADAS PENDENTES": duplicadas_pendentes,
    "DUPLICADAS TOTAL": duplicadas_total
})

excel_finl = gerar_excel({
    "SAP FINL": obras_finl
})

excel_canc = gerar_excel({
    "SAP CANC": obras_canc
})

excel_campo = gerar_excel({
    "JA EM CAMPO": obras_campo
})

kml_geral = gerar_kml(
    view,
    superpontos,
    trajetos_exportar
)

kml_duplicadas = gerar_kml(
    duplicadas_total,
    nome="NIP - Obras Duplicadas"
)

zip_equipes = gerar_zip_excel_por_equipe(
    programacao
)

# Exportacao geral
st.subheader("Arquivos Gerais")

c1, c2, c3 = st.columns(3)

with c1:
    st.download_button(
        "📊 EXCEL GERAL",
        data=excel_geral,
        file_name="NIP_Analise_Geral.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

with c2:
    st.download_button(
        "🗺️ KML GERAL",
        data=kml_geral,
        file_name="NIP_Obras_Geral.kml",
        mime="application/vnd.google-earth.kml+xml",
        use_container_width=True
    )

with c3:
    st.download_button(
        "📦 EXCEL POR EQUIPE (ZIP)",
        data=zip_equipes,
        file_name="NIP_Programacao_Equipes.zip",
        mime="application/zip",
        use_container_width=True
    )

st.subheader("Exportacoes Especificas")

c1, c2, c3 = st.columns(3)

with c1:
    st.download_button(
        "🔵 OBRAS DUPLICADAS - EXCEL",
        data=excel_duplicadas,
        file_name="NIP_Obras_Duplicadas.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.download_button(
        "🗺️ DUPLICADAS - KML",
        data=kml_duplicadas,
        file_name="NIP_Obras_Duplicadas.kml",
        mime="application/vnd.google-earth.kml+xml",
        use_container_width=True
    )

with c2:
    st.download_button(
        "⚪ OBRAS FINL - EXCEL",
        data=excel_finl,
        file_name="NIP_Obras_FINL.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.download_button(
        "🟠 OBRAS CANC - EXCEL",
        data=excel_canc,
        file_name="NIP_Obras_CANC.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

with c3:
    st.download_button(
        "🚗 OBRAS JA EM CAMPO - EXCEL",
        data=excel_campo,
        file_name="NIP_Obras_Ja_Em_Campo.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

st.caption(
    f"Saneamento: {dados['qtd_san']} linhas | "
    f"Levantamento: {dados['qtd_lev']} linhas | "
    f"Notas unicas analisadas: {len(auditoria)}"
)
