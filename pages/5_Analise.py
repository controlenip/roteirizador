
import io
import re
import time
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

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment
)


# ============================================================
# CONFIGURACAO GERAL
# ============================================================

st.set_page_config(
    page_title="NIP | Analise Cruzada e Planejamento",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded"
)

EARTH_KM = 6371.0088

OSRM_DEFAULT = "https://router.project-osrm.org"

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_DUP = "OBRA SANEAMENTO E LEVANTAMENTO"

CATEGORIAS = [
    CAT_SAN,
    CAT_LEV,
    CAT_DUP
]

SAP_EXCLUIDOS = {
    "FINL",
    "CANC"
}

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
# ESTILO VISUAL - CORRECAO DO TITULO
# ============================================================

st.markdown("""
<style>

/* Espacamento superior corrigido */
.block-container {
    padding-top: 4.5rem !important;
    padding-bottom: 2.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
}

/* Titulo principal */
.nip-title {
    color: #0D256C !important;
    font-size: 2rem !important;
    font-weight: 800 !important;
    line-height: 1.45 !important;
    margin-top: 0 !important;
    margin-bottom: 0.4rem !important;
    padding-top: 10px !important;
    padding-bottom: 5px !important;
    overflow: visible !important;
    white-space: normal !important;
    word-break: normal !important;
}

/* Subtitulo */
.nip-subtitle {
    color: #64748B;
    font-size: 0.94rem;
    line-height: 1.5;
    margin-top: 0.2rem;
    margin-bottom: 1.5rem;
}

/* Secoes */
.section-title {
    color: #0D256C;
    font-size: 1.32rem;
    font-weight: 750;
    line-height: 1.4;
    margin-top: 1.3rem;
    margin-bottom: 0.9rem;
    padding-bottom: 4px;
}

/* Cards de indicadores */
.nip-card {
    background: #FFFFFF;
    padding: 18px;
    border-radius: 11px;
    border: 1px solid #E2E8F0;
    border-left: 5px solid var(--cor);
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    margin-bottom: 12px;
    min-height: 122px;
    box-sizing: border-box;
}

.nip-card-label {
    color: #64748B;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    line-height: 1.4;
}

.nip-card-value {
    color: #172554;
    font-size: clamp(20px, 2vw, 29px);
    font-weight: 800;
    line-height: 1.4;
    margin-top: 6px;
    margin-bottom: 4px;
    overflow-wrap: anywhere;
}

.nip-card-desc {
    color: #64748B;
    font-size: 11px;
    line-height: 1.4;
}

/* Uploads */
div[data-testid="stFileUploader"] {
    border-radius: 10px;
}

/* Botoes */
div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {
    border-radius: 8px;
    min-height: 42px;
}

/* Tabelas */
div[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
}

/* Adaptacao para telas menores */
@media (max-width: 900px) {
    .block-container {
        padding-top: 3.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    .nip-title {
        font-size: 1.55rem !important;
        line-height: 1.4 !important;
    }

    .nip-card {
        min-height: 105px;
        padding: 12px;
    }

    .nip-card-value {
        font-size: 22px;
    }
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# COMPONENTES VISUAIS
# ============================================================

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
    if not itens:
        return

    for inicio in range(0, len(itens), 5):
        lote = itens[inicio:inicio + 5]
        colunas = st.columns(len(lote))

        for coluna, item in zip(colunas, lote):
            with coluna:
                st.markdown(
                    card(*item),
                    unsafe_allow_html=True
                )


def secao(titulo):
    st.markdown(
        (
            '<div class="section-title">'
            + html.escape(titulo)
            + '</div>'
        ),
        unsafe_allow_html=True
    )


# ============================================================
# CRONOMETRO
# ============================================================

def formatar_tempo(segundos):
    if segundos is None:
        return "Calculando..."

    segundos = max(0, int(segundos))

    horas, resto = divmod(segundos, 3600)
    minutos, segundos = divmod(resto, 60)

    if horas:
        return (
            f"{horas:02d}:"
            f"{minutos:02d}:"
            f"{segundos:02d}"
        )

    return f"{minutos:02d}:{segundos:02d}"


def renderizar_cronometro(
    decorrido,
    processados,
    total
):
    if processados > 0 and total > processados:
        restante = (
            decorrido / processados
        ) * (total - processados)

    elif total == processados:
        restante = 0

    else:
        restante = None

    mostrar_cards([
        (
            "Tempo Decorrido",
            formatar_tempo(decorrido),
            "#0D256C",
            "Processamento OSRM"
        ),
        (
            "Tempo Restante",
            formatar_tempo(restante),
            "#16A34A",
            "Estimativa"
        ),
        (
            "Trechos Processados",
            f"{processados}/{total}",
            "#9333EA",
            "Progresso"
        )
    ])


# ============================================================
# NORMALIZACAO DOS DADOS
# ============================================================

def norm(valor):
    if pd.isna(valor):
        return ""

    texto = unicodedata.normalize(
        "NFKD",
        str(valor)
    ).encode(
        "ascii",
        "ignore"
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
        "",
        "NAN",
        "NONE",
        "NULL",
        "0"
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


def procurar_coluna(
    df,
    alternativas,
    obrigatoria=True
):
    mapa = {
        norm(coluna): coluna
        for coluna in df.columns
    }

    for alternativa in alternativas:
        if norm(alternativa) in mapa:
            return mapa[norm(alternativa)]

    if obrigatoria:
        raise ValueError(
            "Coluna obrigatoria nao encontrada: "
            + " / ".join(alternativas)
        )

    return None


def converter_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(
            ",",
            ".",
            regex=False
        ),
        errors="coerce"
    )


def validar_coordenadas(
    df,
    coluna_lat,
    coluna_lon
):
    valido = (
        df[coluna_lat].between(-35, 6)
        & df[coluna_lon].between(-75, -30)
        & df[coluna_lat].ne(0)
        & df[coluna_lon].ne(0)
    )

    df.loc[
        ~valido,
        coluna_lat
    ] = np.nan

    df.loc[
        ~valido,
        coluna_lon
    ] = np.nan


def escolher_aba(
    arquivo,
    grupos_colunas,
    preferida=None
):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    abas = list(excel.sheet_names)

    if preferida:
        abas.sort(
            key=lambda aba: (
                0 if norm(aba) == norm(preferida)
                else 1
            )
        )

    for aba in abas:
        for indice_cabecalho in range(5):
            try:
                amostra = pd.read_excel(
                    excel,
                    sheet_name=aba,
                    header=indice_cabecalho,
                    nrows=3,
                    dtype=str
                )

                nomes = {
                    norm(coluna)
                    for coluna in amostra.columns
                }

                encontrou = all(
                    any(
                        norm(alternativa) in nomes
                        for alternativa in grupo
                    )
                    for grupo in grupos_colunas
                )

                if encontrou:
                    df = pd.read_excel(
                        excel,
                        sheet_name=aba,
                        header=indice_cabecalho,
                        dtype=str
                    )

                    return (
                        df,
                        aba,
                        indice_cabecalho
                    )

            except Exception:
                continue

    raise ValueError(
        "Nenhuma aba compativel encontrada. "
        "Abas disponiveis: "
        + ", ".join(excel.sheet_names)
    )


# ============================================================
# CONVERSAO DE DATAS
# ============================================================

def converter_data(valor):
    if pd.isna(valor):
        return pd.NaT

    if isinstance(
        valor,
        (pd.Timestamp, datetime)
    ):
        return pd.Timestamp(valor)

    texto = str(valor).strip()

    if norm(texto) in {
        "",
        "NAN",
        "NONE",
        "NULL",
        "0",
        "-"
    }:
        return pd.NaT

    try:
        numero = float(
            texto.replace(",", ".")
        )

        if 20000 <= numero <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.Timedelta(days=numero)
            )

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
# LEITURA BASE SANEAMENTO
# ============================================================

def ler_saneamento(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        df = None

        for encoding in [
            "utf-8-sig",
            "latin-1"
        ]:
            try:
                df = pd.read_csv(
                    io.BytesIO(
                        arquivo.getvalue()
                    ),
                    sep=None,
                    engine="python",
                    encoding=encoding,
                    dtype=str
                )

                break

            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                continue

        if df is None:
            raise ValueError(
                "Nao foi possivel ler o CSV."
            )

        aba = "CSV"

    else:
        df, aba, _ = escolher_aba(
            arquivo,
            [
                ["NOTA"],
                [
                    "MUNICIPIO",
                    "MUNICÍPIO",
                    "CIDADE"
                ]
            ],
            preferida="Clientes Existentes"
        )

    c_nota = procurar_coluna(
        df,
        ["NOTA"]
    )

    c_municipio = procurar_coluna(
        df,
        [
            "MUNICIPIO",
            "MUNICÍPIO",
            "CIDADE"
        ]
    )

    c_regional = procurar_coluna(
        df,
        ["REGIONAL"],
        False
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

    df["CHAVE_NOTA"] = (
        df[c_nota].map(chave_nota)
    )

    df["MUNICIPIO_OBRA"] = (
        df[c_municipio].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None
        else ""
    )

    df["LAT_OBRA"] = (
        converter_numero(df[c_lat])
    )

    df["LON_OBRA"] = (
        converter_numero(df[c_lon])
    )

    validar_coordenadas(
        df,
        "LAT_OBRA",
        "LON_OBRA"
    )

    return df, aba


# ============================================================
# REGRAS BASE LEVANTAMENTO
# ============================================================

def motivos_levantamento(linha):
    motivos = []

    sap = linha["SAP_NORM"]
    lista = linha["LIST_NORM"]
    contrato = linha["CONTRATO_NORM"]
    m = linha["M_NORM"]
    n = linha["N_NORM"]

    if sap in SAP_EXCLUIDOS:
        motivos.append(
            f"SAP {sap}"
        )

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

    if linha["JA_EM_CAMPO"]:
        motivos.append(
            "JA EM CAMPO - EQUIPE E DATA PREENCHIDAS"
        )

    return " | ".join(motivos)


# ============================================================
# LEITURA BASE LEVANTAMENTO
# ============================================================

def ler_levantamento(arquivo):
    df, aba, header = escolher_aba(
        arquivo,
        [
            ["PROTOCOLO"],
            [
                "STATUS SAP",
                "STATUS_SAP"
            ],
            [
                "STATUS LIST",
                "STATUS_LIST"
            ]
        ],
        preferida="NOTAS"
    )

    c_nota = procurar_coluna(
        df,
        ["PROTOCOLO"]
    )

    c_sap = procurar_coluna(
        df,
        [
            "STATUS SAP",
            "STATUS_SAP"
        ]
    )

    c_list = procurar_coluna(
        df,
        [
            "STATUS LIST",
            "STATUS_LIST"
        ]
    )

    c_contrato = procurar_coluna(
        df,
        ["CONTRATO"]
    )

    c_municipio = procurar_coluna(
        df,
        [
            "MUNICIPIO",
            "MUNICÍPIO",
            "CIDADE"
        ]
    )

    c_regional = procurar_coluna(
        df,
        ["REGIONAL"],
        False
    )

    c_lat = procurar_coluna(
        df,
        ["LATITUDE"]
    )

    c_lon = procurar_coluna(
        df,
        ["LONGITUDE"]
    )

    c_m = next(
        (
            coluna
            for coluna in df.columns
            if "ORCAMENTO MODULAR"
            in norm(coluna)
        ),
        None
    )

    c_n = next(
        (
            coluna
            for coluna in df.columns
            if "PLA ALVOS"
            in norm(coluna)
        ),
        None
    )

    if c_m is None or c_n is None:
        raise ValueError(
            "Colunas ORCAMENTO MODULAR "
            "ou PLA ALVOS nao encontradas."
        )

    if len(df.columns) < 3:
        raise ValueError(
            "A base de Levantamento possui "
            "menos de tres colunas."
        )

    coluna_b = df.columns[1]
    coluna_c = df.columns[2]

    c_prioridade = procurar_coluna(
        df,
        ["PRIORIDADE"],
        False
    )

    c_abertura = procurar_coluna(
        df,
        [
            "DATA ABERTURA",
            "DATA DE ABERTURA"
        ],
        False
    )

    df["CHAVE_NOTA"] = (
        df[c_nota].map(chave_nota)
    )

    df["MUNICIPIO_OBRA"] = (
        df[c_municipio].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None
        else ""
    )

    df["LAT_OBRA"] = (
        converter_numero(df[c_lat])
    )

    df["LON_OBRA"] = (
        converter_numero(df[c_lon])
    )

    validar_coordenadas(
        df,
        "LAT_OBRA",
        "LON_OBRA"
    )

    df["SAP_NORM"] = (
        df[c_sap].map(status)
    )

    df["LIST_NORM"] = (
        df[c_list].map(status)
    )

    df["CONTRATO_NORM"] = (
        df[c_contrato].map(status)
    )

    df["M_NORM"] = (
        df[c_m].map(status)
    )

    df["N_NORM"] = (
        df[c_n].map(status)
    )

    df["PRIORIDADE_NORM"] = (
        df[c_prioridade].fillna("")
        if c_prioridade is not None
        else ""
    )

    df["DATA_ABERTURA_NORM"] = (
        df[c_abertura].fillna("")
        if c_abertura is not None
        else ""
    )

    # COLUNA B - RESPONSAVEL
    df["RESPONSAVEL_CAMPO"] = (
        df[coluna_b]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # COLUNA C - DATA
    df["DATA_CAMPO"] = (
        df[coluna_c].map(converter_data)
    )

    # EXCLUI APENAS QUANDO AMBAS AS CONDICOES
    # FOREM VERDADEIRAS
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

    info = {
        "ABA": aba,
        "LINHA_CABECALHO": header + 1,
        "COLUNA_B": str(coluna_b),
        "COLUNA_C": str(coluna_c),
        "JA_EM_CAMPO_LINHAS": int(
            df["JA_EM_CAMPO"].sum()
        )
    }

    return df, info


# ============================================================
# LEITURA DAS EQUIPES
# ============================================================

def ler_equipes(arquivo):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    partes = []

    configuracoes = [
        (
            "LEVANTAMENTO",
            "LEVANTADORES"
        ),
        (
            "SANEAMENTO",
            "SANEAMENTO"
        )
    ]

    for tipo, nome_aba in configuracoes:
        aba = next(
            (
                nome
                for nome in excel.sheet_names
                if norm(nome) == nome_aba
            ),
            None
        )

        if aba is None:
            raise ValueError(
                "Aba obrigatoria nao localizada: "
                + nome_aba
            )

        df = pd.read_excel(
            excel,
            sheet_name=aba,
            dtype=str
        )

        c_nome = procurar_coluna(
            df,
            [
                "NOME",
                "EQUIPE",
                "NOME_COLAB"
            ]
        )

        c_cidade = procurar_coluna(
            df,
            [
                "CIDADES",
                "CIDADE",
                "MUNICIPIO"
            ]
        )

        c_lat = procurar_coluna(
            df,
            [
                "LATITUDE",
                "LAT"
            ]
        )

        c_lon = procurar_coluna(
            df,
            [
                "LONGITUDE",
                "LON"
            ]
        )

        parte = pd.DataFrame({
            "EQUIPE": df[c_nome],
            "CIDADE_BASE": df[c_cidade],
            "LAT_EQUIPE": converter_numero(
                df[c_lat]
            ),
            "LON_EQUIPE": converter_numero(
                df[c_lon]
            ),
            "TIPO_EQUIPE": tipo
        })

        validar_coordenadas(
            parte,
            "LAT_EQUIPE",
            "LON_EQUIPE"
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
        equipes["EQUIPE"]
        .astype(str)
        .str.strip()
    )

    equipes["ID_EQUIPE"] = (
        equipes["TIPO_EQUIPE"]
        + "|"
        + equipes["EQUIPE"]
        + "|"
        + equipes["CIDADE_BASE"].astype(str)
    )

    equipes = equipes.drop_duplicates(
        subset=["ID_EQUIPE"]
    ).reset_index(drop=True)

    if equipes.empty:
        raise ValueError(
            "Nenhuma equipe valida encontrada."
        )

    return equipes


# ============================================================
# CRUZAMENTO DAS BASES
# ============================================================

def primeira_valida(grupo):
    coordenadas_validas = (
        grupo["LAT_OBRA"].notna()
        & grupo["LON_OBRA"].notna()
    )

    if coordenadas_validas.any():
        return grupo.loc[
            coordenadas_validas
        ].iloc[0]

    return grupo.iloc[0]


def consolidar(saneamento, levantamento):
    san = saneamento[
        saneamento["CHAVE_NOTA"].ne("")
    ].copy()

    lev = levantamento[
        levantamento["CHAVE_NOTA"].ne("")
    ].copy()

    grupos_san = dict(
        tuple(
            san.groupby(
                "CHAVE_NOTA",
                sort=False
            )
        )
    )

    grupos_lev = dict(
        tuple(
            lev.groupby(
                "CHAVE_NOTA",
                sort=False
            )
        )
    )

    notas = dict.fromkeys(
        list(grupos_san)
        + list(grupos_lev)
    )

    registros = []

    for nota in notas:
        grupo_san = grupos_san.get(nota)
        grupo_lev = grupos_lev.get(nota)

        tem_san = grupo_san is not None
        tem_lev = grupo_lev is not None

        a = (
            primeira_valida(grupo_san)
            if tem_san
            else None
        )

        b = (
            primeira_valida(grupo_lev)
            if tem_lev
            else None
        )

        if tem_san and tem_lev:
            categoria = CAT_DUP

        elif tem_san:
            categoria = CAT_SAN

        else:
            categoria = CAT_LEV

        motivos = []

        if tem_lev:
            for texto in grupo_lev[
                "MOTIVOS_EXCLUSAO"
            ]:
                if texto:
                    motivos.extend(
                        texto.split(" | ")
                    )

        motivos = list(
            dict.fromkeys(motivos)
        )

        ja_em_campo = (
            bool(
                grupo_lev[
                    "JA_EM_CAMPO"
                ].any()
            )
            if tem_lev
            else False
        )

        responsaveis = []
        datas_campo = []

        if tem_lev:
            linhas_campo = grupo_lev[
                grupo_lev["JA_EM_CAMPO"]
            ]

            responsaveis = [
                str(valor)
                for valor in linhas_campo[
                    "RESPONSAVEL_CAMPO"
                ].dropna().unique()
            ]

            datas_campo = []

            for valor in linhas_campo[
                "DATA_CAMPO"
            ].dropna().unique():
                data = pd.Timestamp(valor)
                datas_campo.append(
                    data.strftime("%d/%m/%Y")
                )

        origem = (
            a if a is not None else b
        )

        if (
            b is not None
            and pd.notna(b["LAT_OBRA"])
            and pd.notna(b["LON_OBRA"])
        ):
            origem = b

        registros.append({
            "NOTA": nota,
            "LISTA": categoria,
            "MUNICIPIO": origem[
                "MUNICIPIO_OBRA"
            ],
            "REGIONAL": origem[
                "REGIONAL_OBRA"
            ],
            "LATITUDE": origem[
                "LAT_OBRA"
            ],
            "LONGITUDE": origem[
                "LON_OBRA"
            ],
            "STATUS_SAP": (
                b["SAP_NORM"]
                if tem_lev else ""
            ),
            "STATUS_LIST": (
                b["LIST_NORM"]
                if tem_lev else ""
            ),
            "CONTRATO": (
                b["CONTRATO_NORM"]
                if tem_lev else ""
            ),
            "COLUNA_M": (
                b["M_NORM"]
                if tem_lev else ""
            ),
            "COLUNA_N": (
                b["N_NORM"]
                if tem_lev else ""
            ),
            "PRIORIDADE_ORIGINAL": (
                b["PRIORIDADE_NORM"]
                if tem_lev else ""
            ),
            "DATA_ABERTURA": (
                b["DATA_ABERTURA_NORM"]
                if tem_lev else ""
            ),
            "DUPLICADA": (
                "SIM"
                if tem_san and tem_lev
                else "NÃO"
            ),
            "JA_EM_CAMPO": (
                "SIM"
                if ja_em_campo
                else "NÃO"
            ),
            "EQUIPE_CAMPO": " | ".join(
                responsaveis
            ),
            "DATA_CAMPO": " | ".join(
                datas_campo
            ),
            "PENDENTE_CONTAGEM": (
                "NÃO"
                if motivos
                else "SIM"
            ),
            "MOTIVO_EXCLUSAO": (
                " | ".join(motivos)
                if motivos
                else "-"
            ),
            "OCORRENCIAS_SANEAMENTO": (
                len(grupo_san)
                if tem_san
                else 0
            ),
            "OCORRENCIAS_LEVANTAMENTO": (
                len(grupo_lev)
                if tem_lev
                else 0
            )
        })

    return pd.DataFrame(registros)


# ============================================================
# DISTANCIAS
# ============================================================

def haversine(
    lat1,
    lon1,
    lat2,
    lon2
):
    lat1 = np.radians(lat1)
    lon1 = np.radians(lon1)
    lat2 = np.radians(lat2)
    lon2 = np.radians(lon2)

    a = (
        np.sin(
            (lat2 - lat1) / 2
        ) ** 2

        + np.cos(lat1)
        * np.cos(lat2)

        * np.sin(
            (lon2 - lon1) / 2
        ) ** 2
    )

    return (
        2
        * EARTH_KM
        * np.arcsin(
            np.sqrt(
                np.clip(a, 0, 1)
            )
        )
    )


def atividades(categoria):
    if categoria == CAT_DUP:
        return [
            "SANEAMENTO",
            "LEVANTAMENTO"
        ]

    if categoria == CAT_SAN:
        return ["SANEAMENTO"]

    return ["LEVANTAMENTO"]


# ============================================================
# EQUIPES MAIS PROXIMAS
# ============================================================

def gerar_candidatos(
    obras,
    equipes,
    quantidade=3
):
    saida = obras.copy().reset_index(
        drop=True
    )

    saida["EQUIPES_SANEAMENTO"] = (
        "NAO APLICAVEL"
    )

    saida["EQUIPES_LEVANTAMENTO"] = (
        "NAO APLICAVEL"
    )

    registros = []

    for tipo in [
        "SANEAMENTO",
        "LEVANTAMENTO"
    ]:
        equipes_tipo = equipes[
            equipes["TIPO_EQUIPE"].eq(tipo)
        ].reset_index(drop=True)

        if equipes_tipo.empty:
            continue

        arvore = BallTree(
            np.radians(
                equipes_tipo[
                    [
                        "LAT_EQUIPE",
                        "LON_EQUIPE"
                    ]
                ].to_numpy(float)
            ),
            metric="haversine"
        )

        indices_obras = [
            i
            for i, obra in saida.iterrows()
            if (
                tipo in atividades(
                    obra["LISTA"]
                )
                and pd.notna(
                    obra["LATITUDE"]
                )
                and pd.notna(
                    obra["LONGITUDE"]
                )
            )
        ]

        if not indices_obras:
            continue

        pontos = np.radians(
            saida.loc[
                indices_obras,
                [
                    "LATITUDE",
                    "LONGITUDE"
                ]
            ].to_numpy(float)
        )

        distancias, indices = (
            arvore.query(
                pontos,
                k=min(
                    quantidade,
                    len(equipes_tipo)
                )
            )
        )

        for posicao, indice_obra in enumerate(
            indices_obras
        ):
            textos = []

            for j in range(
                indices.shape[1]
            ):
                equipe = equipes_tipo.iloc[
                    indices[posicao, j]
                ]

                km = float(
                    distancias[posicao, j]
                    * EARTH_KM
                )

                registros.append({
                    "NOTA": saida.at[
                        indice_obra,
                        "NOTA"
                    ],
                    "ATIVIDADE": tipo,
                    "ID_EQUIPE": equipe[
                        "ID_EQUIPE"
                    ],
                    "EQUIPE": equipe[
                        "EQUIPE"
                    ],
                    "CIDADE_BASE": equipe[
                        "CIDADE_BASE"
                    ],
                    "LAT_EQUIPE": equipe[
                        "LAT_EQUIPE"
                    ],
                    "LON_EQUIPE": equipe[
                        "LON_EQUIPE"
                    ],
                    "DISTANCIA_RETA_KM": round(
                        km,
                        2
                    )
                })

                textos.append(
                    f"{equipe['EQUIPE']} "
                    f"({equipe['CIDADE_BASE']})"
                    f" - {km:.1f} km"
                )

            saida.at[
                indice_obra,
                f"EQUIPES_{tipo}"
            ] = " | ".join(textos)

    candidatos = pd.DataFrame(
        registros,
        columns=[
            "NOTA",
            "ATIVIDADE",
            "ID_EQUIPE",
            "EQUIPE",
            "CIDADE_BASE",
            "LAT_EQUIPE",
            "LON_EQUIPE",
            "DISTANCIA_RETA_KM"
        ]
    )

    return saida, candidatos


# ============================================================
# PRIORIDADES
# ============================================================

def priorizar(
    obras,
    dias_media=14
):
    df = obras.copy()

    datas = pd.to_datetime(
        df["DATA_ABERTURA"],
        errors="coerce",
        dayfirst=True
    )

    df["DIAS_ABERTURA"] = (
        pd.Timestamp.today().normalize()
        - datas
    ).dt.days

    def classificar(linha):
        prioridade_original = norm(
            linha["PRIORIDADE_ORIGINAL"]
        )

        if any(
            palavra in prioridade_original
            for palavra in [
                "URGENTE",
                "CRITICA",
                "ALTA"
            ]
        ):
            return "ALTA"

        idade = linha["DIAS_ABERTURA"]

        if pd.notna(idade):
            if idade >= dias_media * 4:
                return "ALTA"

            if idade >= dias_media:
                return "MEDIA"

        return "BAIXA"

    df["PRIORIDADE_PLANEJAMENTO"] = (
        df.apply(
            classificar,
            axis=1
        )
    )

    return df


# ============================================================
# GERAR TAREFAS
# ============================================================

def gerar_tarefas(obras):
    registros = []

    for _, obra in obras.iterrows():
        for tipo in atividades(
            obra["LISTA"]
        ):
            registro = obra.to_dict()
            registro["ATIVIDADE"] = tipo

            registros.append(registro)

    return pd.DataFrame(registros)


# ============================================================
# DATAS DA PROGRAMACAO
# ============================================================

def gerar_datas(
    data_inicio,
    modo,
    periodos
):
    inicio = pd.Timestamp(
        data_inicio
    )

    if modo == "POR DIA":
        fim = (
            inicio
            + pd.Timedelta(
                days=periodos * 2 + 15
            )
        )

    elif modo == "POR SEMANA":
        fim = (
            inicio
            + pd.Timedelta(
                weeks=periodos
            )
        )

    elif modo == "POR MES":
        fim = (
            inicio
            + pd.DateOffset(
                months=periodos
            )
            - pd.Timedelta(days=1)
        )

    else:
        fim = (
            inicio
            + pd.Timedelta(days=365)
        )

    datas = pd.date_range(
        inicio,
        fim,
        freq="D"
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

        return (
            f"{iso.year}-S{iso.week:02d}"
        )

    if modo == "POR MES":
        return data.strftime("%Y-%m")

    return data.isoformat()


# ============================================================
# PROGRAMACAO
# ============================================================

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
        return (
            pd.DataFrame(),
            pd.DataFrame()
        )

    ordem_prioridade = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    tarefas = tarefas.copy()

    tarefas["_ORDEM"] = (
        tarefas[
            "PRIORIDADE_PLANEJAMENTO"
        ]
        .map(ordem_prioridade)
        .fillna(3)
    )

    tarefas = tarefas.sort_values(
        [
            "_ORDEM",
            "MUNICIPIO",
            "NOTA"
        ]
    )

    datas = gerar_datas(
        data_inicio,
        modo,
        periodos
    )

    candidatos_por_nota = defaultdict(list)

    for _, candidato in candidatos.iterrows():
        chave = (
            candidato["NOTA"],
            candidato["ATIVIDADE"]
        )

        candidatos_por_nota[
            chave
        ].append(
            candidato.to_dict()
        )

    ocupacao_dia = defaultdict(int)
    ocupacao_periodo = defaultdict(int)

    resultados = []

    for _, tarefa in tarefas.iterrows():
        chave = (
            tarefa["NOTA"],
            tarefa["ATIVIDADE"]
        )

        alternativas = []

        for candidato in candidatos_por_nota.get(
            chave,
            []
        ):
            km = float(
                candidato["DISTANCIA_RETA_KM"]
            )

            if km > raio:
                continue

            for data in datas:
                periodo = obter_periodo(
                    data,
                    modo
                )

                chave_dia = (
                    candidato["ID_EQUIPE"],
                    data
                )

                chave_periodo = (
                    candidato["ID_EQUIPE"],
                    periodo
                )

                carga_dia = ocupacao_dia[
                    chave_dia
                ]

                carga_periodo = (
                    ocupacao_periodo[
                        chave_periodo
                    ]
                )

                if carga_dia >= limite_dia:
                    continue

                if (
                    carga_periodo
                    >= limite_periodo
                ):
                    continue

                alternativas.append({
                    "CANDIDATO": candidato,
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
                key=lambda item: (
                    item["CARGA_PERIODO"],
                    item["CARGA_DIA"],
                    item["KM"],
                    item["DATA"]
                )
            )

            candidato = melhor["CANDIDATO"]
            data = melhor["DATA"]
            periodo = melhor["PERIODO"]

            ocupacao_dia[
                (
                    candidato["ID_EQUIPE"],
                    data
                )
            ] += 1

            ocupacao_periodo[
                (
                    candidato["ID_EQUIPE"],
                    periodo
                )
            ] += 1

            iso = data.isocalendar()

            registro.update({
                "ID_EQUIPE": candidato[
                    "ID_EQUIPE"
                ],
                "EQUIPE_PROGRAMADA": candidato[
                    "EQUIPE"
                ],
                "CIDADE_BASE": candidato[
                    "CIDADE_BASE"
                ],
                "DATA_PROGRAMADA": data.isoformat(),
                "DIA_SEMANA": DIAS_PT[
                    data.weekday()
                ],
                "DIA_MES": data.strftime(
                    "%d/%m/%Y"
                ),
                "SEMANA": (
                    f"{iso.year}-S"
                    f"{iso.week:02d}"
                ),
                "MES": data.strftime(
                    "%Y-%m"
                ),
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

    programacao = pd.DataFrame(
        resultados
    )

    alocadas = programacao[
        programacao[
            "STATUS_PROGRAMACAO"
        ].eq("SUGESTAO")
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
        .reset_index(
            name="TAREFAS"
        )
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
    resumos = []

    for tipo, descricao in [
        (
            "SANEAMENTO",
            "Saneamento"
        ),
        (
            "LEVANTAMENTO",
            "Levantamento"
        )
    ]:
        equipes_tipo = equipes[
            equipes[
                "TIPO_EQUIPE"
            ].eq(tipo)
        ]

        total_equipes = len(
            equipes_tipo
        )

        categorias = (
            [CAT_SAN, CAT_DUP]
            if tipo == "SANEAMENTO"
            else [CAT_LEV, CAT_DUP]
        )

        quantidade_obras = int(
            obras[
                "LISTA"
            ].isin(
                categorias
            ).sum()
        )

        media = (
            quantidade_obras / total_equipes
            if total_equipes
            else 0
        )

        tarefas_tipo = (
            programacao[
                programacao[
                    "ATIVIDADE"
                ].eq(tipo)
            ]
            if not programacao.empty
            else pd.DataFrame()
        )

        programadas = (
            int(
                tarefas_tipo[
                    "STATUS_PROGRAMACAO"
                ].eq(
                    "SUGESTAO"
                ).sum()
            )
            if not tarefas_tipo.empty
            else 0
        )

        resumos.append({
            "ATIVIDADE": descricao,
            "EQUIPES": total_equipes,
            "OBRAS_DISPONIVEIS": quantidade_obras,
            "MEDIA_OBRAS_POR_EQUIPE": round(
                media, 2
            ),
            "TAREFAS_PROGRAMADAS": programadas,
            "MEDIA_PROGRAMADA_POR_EQUIPE": (
                round(
                    programadas / total_equipes,
                    2
                )
                if total_equipes
                else 0
            )
        })

    quadro = pd.DataFrame(
        resumos
    )

    if not programacao.empty:
        contagem = (
            programacao[
                programacao[
                    "STATUS_PROGRAMACAO"
                ].eq("SUGESTAO")
            ]
            .groupby("ID_EQUIPE")
            .size()
            .to_dict()
        )

    else:
        contagem = {}

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

    for tipo, descricao in [
        (
            "SANEAMENTO",
            "Saneamento"
        ),
        (
            "LEVANTAMENTO",
            "Levantamento"
        )
    ]:
        media = quadro.loc[
            quadro[
                "ATIVIDADE"
            ].eq(descricao),
            "MEDIA_OBRAS_POR_EQUIPE"
        ].iloc[0]

        detalhe.loc[
            detalhe[
                "TIPO_EQUIPE"
            ].eq(tipo),
            "MEDIA_DISPONIVEL_ATIVIDADE"
        ] = media

    return quadro, detalhe


# ============================================================
# SUPERPONTOS
# ============================================================

def fundir_superpontos(
    df,
    raio_metros=50
):
    if df.empty:
        return (
            pd.DataFrame(),
            0
        )

    validas = df.dropna(
        subset=[
            "LATITUDE",
            "LONGITUDE"
        ]
    ).copy()

    invalidas = df[
        df["LATITUDE"].isna()
        | df["LONGITUDE"].isna()
    ].copy()

    saida = []

    if not validas.empty:
        coordenadas = np.radians(
            validas[
                [
                    "LATITUDE",
                    "LONGITUDE"
                ]
            ].to_numpy(float)
        )

        modelo = DBSCAN(
            eps=raio_metros / 6371000,
            min_samples=1,
            metric="haversine",
            algorithm="ball_tree"
        ).fit(coordenadas)

        validas["CLUSTER_ID"] = (
            modelo.labels_
        )

        validas["CHAVE_GRUPO"] = (
            validas[
                "CLUSTER_ID"
            ].astype(str)
            + "|"
            + validas[
                "ID_EQUIPE"
            ].astype(str)
            + "|"
            + validas[
                "DATA_PROGRAMADA"
            ].astype(str)
            + "|"
            + validas[
                "ATIVIDADE"
            ].astype(str)
        )

        for _, grupo in validas.groupby(
            "CHAVE_GRUPO",
            sort=False
        ):
            base = grupo.iloc[0].copy()

            originais = (
                grupo.to_dict(
                    "records"
                )
            )

            quantidade = len(
                originais
            )

            base["_ORIGINAL_ROWS"] = (
                originais
            )

            base["LATITUDE"] = (
                grupo[
                    "LATITUDE"
                ].mean()
            )

            base["LONGITUDE"] = (
                grupo[
                    "LONGITUDE"
                ].mean()
            )

            base["SUPER_PONTO"] = (
                f"SIM ({quantidade} un.)"
                if quantidade > 1
                else "NÃO"
            )

            base["NOTAS_AGRUPADAS"] = (
                " | ".join(
                    grupo[
                        "NOTA"
                    ].astype(str).tolist()
                )
            )

            saida.append(base)

    if not invalidas.empty:
        invalidas["SUPER_PONTO"] = (
            "NÃO"
        )

        invalidas["NOTAS_AGRUPADAS"] = (
            invalidas[
                "NOTA"
            ].astype(str)
        )

        invalidas["_ORIGINAL_ROWS"] = [
            [registro]
            for registro in invalidas.to_dict(
                "records"
            )
        ]

        saida.extend(
            invalidas.to_dict(
                "records"
            )
        )

    final = pd.DataFrame(
        saida
    )

    final = final.drop(
        columns=[
            "CLUSTER_ID",
            "CHAVE_GRUPO"
        ],
        errors="ignore"
    )

    return (
        final,
        len(df) - len(final)
    )


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
    if lat1 == lat2 and lon1 == lon2:
        return {
            "OK": True,
            "KM": 0.0,
            "MIN": 0.0,
            "GEOMETRIA": [
                [lon1, lat1],
                [lon2, lat2]
            ]
        }

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
            return {
                "OK": False
            }

        rotas = dados.get(
            "routes",
            []
        )

        if not rotas:
            return {
                "OK": False
            }

        rota = rotas[0]

        geometria = (
            rota.get(
                "geometry",
                {}
            ).get(
                "coordinates",
                []
            )
        )

        if len(geometria) < 2:
            return {
                "OK": False
            }

        return {
            "OK": True,
            "KM": round(
                float(
                    rota["distance"]
                ) / 1000,
                2
            ),
            "MIN": round(
                float(
                    rota["duration"]
                ) / 60,
                1
            ),
            "GEOMETRIA": geometria
        }

    except (
        requests.RequestException,
        ValueError,
        TypeError,
        KeyError
    ):
        return {
            "OK": False
        }


# ============================================================
# PREPARAR TRECHOS
# ============================================================

def preparar_trechos(
    programacao,
    equipes
):
    if programacao.empty:
        return pd.DataFrame()

    linhas = []

    validas = programacao[
        programacao[
            "STATUS_PROGRAMACAO"
        ].eq("SUGESTAO")
    ].dropna(
        subset=[
            "LATITUDE",
            "LONGITUDE"
        ]
    )

    for (
        id_equipe,
        data
    ), grupo in validas.groupby(
        [
            "ID_EQUIPE",
            "DATA_PROGRAMADA"
        ]
    ):
        equipe_df = equipes[
            equipes[
                "ID_EQUIPE"
            ].eq(id_equipe)
        ]

        if equipe_df.empty:
            continue

        equipe = equipe_df.iloc[0]

        lat_atual = float(
            equipe[
                "LAT_EQUIPE"
            ]
        )

        lon_atual = float(
            equipe[
                "LON_EQUIPE"
            ]
        )

        restantes = (
            grupo.to_dict(
                "records"
            )
        )

        ordem = 0

        while restantes:
            distancias = [
                float(
                    haversine(
                        lat_atual,
                        lon_atual,
                        float(
                            registro[
                                "LATITUDE"
                            ]
                        ),
                        float(
                            registro[
                                "LONGITUDE"
                            ]
                        )
                    )
                )
                for registro in restantes
            ]

            indice = int(
                np.argmin(
                    distancias
                )
            )

            obra = restantes.pop(
                indice
            )

            ordem += 1

            linhas.append({
                "NOTA": obra["NOTA"],
                "ATIVIDADE": obra[
                    "ATIVIDADE"
                ],
                "EQUIPE": obra[
                    "EQUIPE_PROGRAMADA"
                ],
                "ID_EQUIPE": id_equipe,
                "DATA": data,
                "ORDEM": ordem,
                "LAT_ORIGEM": lat_atual,
                "LON_ORIGEM": lon_atual,
                "LAT_DESTINO": float(
                    obra[
                        "LATITUDE"
                    ]
                ),
                "LON_DESTINO": float(
                    obra[
                        "LONGITUDE"
                    ]
                ),
                "KM_RETA": round(
                    distancias[indice],
                    2
                )
            })

            lat_atual = float(
                obra[
                    "LATITUDE"
                ]
            )

            lon_atual = float(
                obra[
                    "LONGITUDE"
                ]
            )

    return pd.DataFrame(
        linhas
    )


# ============================================================
# PROCESSAMENTO OSRM POR LOTES
# ============================================================

def processar_lote_osrm(
    trechos,
    servidor,
    inicio,
    tamanho_lote
):
    novos = []

    fim = min(
        inicio + tamanho_lote,
        len(trechos)
    )

    progresso = st.progress(
        inicio / max(
            1,
            len(trechos)
        )
    )

    texto_progresso = st.empty()

    inicio_lote = time.monotonic()

    for indice in range(
        inicio,
        fim
    ):
        trecho = trechos.iloc[
            indice
        ]

        rota = consultar_osrm(
            round(
                float(
                    trecho[
                        "LAT_ORIGEM"
                    ]
                ),
                6
            ),
            round(
                float(
                    trecho[
                        "LON_ORIGEM"
                    ]
                ),
                6
            ),
            round(
                float(
                    trecho[
                        "LAT_DESTINO"
                    ]
                ),
                6
            ),
            round(
                float(
                    trecho[
                        "LON_DESTINO"
                    ]
                ),
                6
            ),
            servidor
        )

        registro = trecho.to_dict()

        if rota["OK"]:
            registro[
                "TIPO_TRAJETO"
            ] = "OSRM"

            registro[
                "KM_RODOVIARIO"
            ] = rota["KM"]

            registro[
                "TEMPO_MIN"
            ] = rota["MIN"]

            registro[
                "GEOMETRIA"
            ] = rota["GEOMETRIA"]

        else:
            registro[
                "TIPO_TRAJETO"
            ] = "ESTIMATIVA - LINHA RETA"

            registro[
                "KM_RODOVIARIO"
            ] = np.nan

            registro[
                "TEMPO_MIN"
            ] = np.nan

            registro[
                "GEOMETRIA"
            ] = [
                [
                    registro[
                        "LON_ORIGEM"
                    ],
                    registro[
                        "LAT_ORIGEM"
                    ]
                ],
                [
                    registro[
                        "LON_DESTINO"
                    ],
                    registro[
                        "LAT_DESTINO"
                    ]
                ]
            ]

        novos.append(
            registro
        )

        processados = indice + 1

        progresso.progress(
            processados / len(
                trechos
            )
        )

        texto_progresso.markdown(
            f"**Processando trecho "
            f"{processados}/{len(trechos)}** "
            f"| Lote: "
            f"{formatar_tempo(time.monotonic() - inicio_lote)}"
        )

    return (
        novos,
        fim,
        time.monotonic() - inicio_lote
    )


# ============================================================
# POPUPS DOS MARCADORES
# ============================================================

def popup_html(obra):
    def esc(campo):
        return html.escape(
            texto_seguro(
                obra.get(
                    campo,
                    ""
                )
            )
        )

    linhas = [
        (
            "Nota",
            esc("NOTA")
        ),
        (
            "Municipio",
            esc("MUNICIPIO")
        ),
        (
            "Tipo",
            esc("LISTA")
        )
    ]

    saneamento = esc(
        "EQUIPES_SANEAMENTO"
    )

    levantamento = esc(
        "EQUIPES_LEVANTAMENTO"
    )

    if (
        saneamento
        and saneamento != "NAO APLICAVEL"
    ):
        linhas.append(
            (
                "Equipes Saneamento",
                saneamento
            )
        )

    if (
        levantamento
        and levantamento != "NAO APLICAVEL"
    ):
        linhas.append(
            (
                "Equipes Levantamento",
                levantamento
            )
        )

    campos_adicionais = [
        (
            "REGIONAL",
            "Regional"
        ),
        (
            "PRIORIDADE_PLANEJAMENTO",
            "Prioridade"
        ),
        (
            "EQUIPE_PROGRAMADA",
            "Equipe Programada"
        ),
        (
            "DATA_PROGRAMADA",
            "Data Programada"
        ),
        (
            "SUPER_PONTO",
            "Superponto"
        ),
        (
            "NOTAS_AGRUPADAS",
            "Notas Agrupadas"
        )
    ]

    for campo, rotulo in campos_adicionais:
        valor = esc(campo)

        if (
            valor
            and valor.lower() != "nan"
        ):
            linhas.append(
                (
                    rotulo,
                    valor
                )
            )

    corpo = "".join(
        f"""
        <tr>
            <td style="
                padding:6px;
                width:115px;
                font-weight:bold;
                vertical-align:top;
                border-bottom:1px solid #eee;
            ">
                {html.escape(rotulo)}
            </td>
            <td style="
                padding:6px;
                border-bottom:1px solid #eee;
                word-break:break-word;
            ">
                {valor}
            </td>
        </tr>
        """
        for rotulo, valor in linhas
    )

    cor = CORES.get(
        obra.get(
            "LISTA"
        ),
        "#0D256C"
    )

    return f"""
    <div style="
        font-family:Arial,sans-serif;
        width:350px;
        max-width:100%;
        font-size:12px;
        color:#1f2937;
    ">
        <div style="
            background:{cor};
            color:white;
            padding:12px;
            font-weight:bold;
            border-radius:8px 8px 0 0;
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
    namespace = (
        "http://www.opengis.net/kml/2.2"
    )

    ET.register_namespace(
        "",
        namespace
    )

    def elemento(
        pai,
        tag,
        texto=None
    ):
        item = ET.SubElement(
            pai,
            f"{{{namespace}}}{tag}"
        )

        if texto is not None:
            item.text = str(
                texto
            )

        return item

    raiz = ET.Element(
        f"{{{namespace}}}kml"
    )

    documento = elemento(
        raiz,
        "Document"
    )

    elemento(
        documento,
        "name",
        nome
    )

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

    for categoria, (
        id_estilo,
        url
    ) in estilos.items():
        estilo = elemento(
            documento,
            "Style"
        )

        estilo.set(
            "id",
            id_estilo
        )

        icone_estilo = elemento(
            estilo,
            "IconStyle"
        )

        elemento(
            icone_estilo,
            "scale",
            "1.2"
        )

        icone = elemento(
            icone_estilo,
            "Icon"
        )

        elemento(
            icone,
            "href",
            url
        )

        rotulo = elemento(
            estilo,
            "LabelStyle"
        )

        elemento(
            rotulo,
            "scale",
            "0"
        )

    # PONTOS DAS OBRAS
    for categoria in CATEGORIAS:
        pasta = elemento(
            documento,
            "Folder"
        )

        elemento(
            pasta,
            "name",
            categoria
        )

        if obras.empty:
            continue

        dados_categoria = obras[
            obras[
                "LISTA"
            ].eq(categoria)
        ].dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        for _, obra in dados_categoria.iterrows():
            ponto = elemento(
                pasta,
                "Placemark"
            )

            elemento(
                ponto,
                "name",
                f"NOTA {obra['NOTA']}"
            )

            elemento(
                ponto,
                "styleUrl",
                "#" + estilos[categoria][0]
            )

            elemento(
                ponto,
                "description",
                popup_html(obra)
            )

            geometria = elemento(
                ponto,
                "Point"
            )

            elemento(
                geometria,
                "coordinates",
                (
                    f"{float(obra['LONGITUDE'])},"
                    f"{float(obra['LATITUDE'])},0"
                )
            )

    # SUPERPONTOS
    if (
        superpontos is not None
        and not superpontos.empty
    ):
        pasta_super = elemento(
            documento,
            "Folder"
        )

        elemento(
            pasta_super,
            "name",
            "SUPERPONTOS"
        )

        for _, registro in superpontos.iterrows():
            if not str(
                registro.get(
                    "SUPER_PONTO",
                    ""
                )
            ).startswith("SIM"):
                continue

            if (
                pd.isna(
                    registro["LATITUDE"]
                )
                or pd.isna(
                    registro["LONGITUDE"]
                )
            ):
                continue

            ponto = elemento(
                pasta_super,
                "Placemark"
            )

            elemento(
                ponto,
                "name",
                (
                    "SUPERPONTO - "
                    + str(
                        registro[
                            "SUPER_PONTO"
                        ]
                    )
                )
            )

            categoria = registro[
                "LISTA"
            ]

            elemento(
                ponto,
                "styleUrl",
                "#" + estilos.get(
                    categoria,
                    estilos[CAT_DUP]
                )[0]
            )

            elemento(
                ponto,
                "description",
                popup_html(registro)
            )

            geometria = elemento(
                ponto,
                "Point"
            )

            elemento(
                geometria,
                "coordinates",
                (
                    f"{float(registro['LONGITUDE'])},"
                    f"{float(registro['LATITUDE'])},0"
                )
            )

    # TRAJETOS
    if (
        trajetos is not None
        and not trajetos.empty
    ):
        pasta_rotas = elemento(
            documento,
            "Folder"
        )

        elemento(
            pasta_rotas,
            "name",
            "TRAJETOS DAS EQUIPES"
        )

        for _, trecho in trajetos.iterrows():
            geometria = trecho.get(
                "GEOMETRIA"
            )

            if not isinstance(
                geometria,
                list
            ):
                continue

            if len(geometria) < 2:
                continue

            ponto = elemento(
                pasta_rotas,
                "Placemark"
            )

            tipo = trecho.get(
                "TIPO_TRAJETO",
                "LINHA RETA"
            )

            elemento(
                ponto,
                "name",
                (
                    f"{trecho['EQUIPE']} - "
                    f"{trecho['ORDEM']} - "
                    f"{tipo}"
                )
            )

            elemento(
                ponto,
                "description",
                (
                    f"Equipe: {trecho['EQUIPE']}\n"
                    f"Nota: {trecho['NOTA']}\n"
                    f"Tipo: {tipo}\n"
                    f"KM rodoviario: "
                    f"{texto_seguro(trecho.get('KM_RODOVIARIO'))}"
                )
            )

            linha = elemento(
                ponto,
                "LineString"
            )

            elemento(
                linha,
                "tessellate",
                "1"
            )

            elemento(
                linha,
                "coordinates",
                " ".join(
                    f"{coordenada[0]},"
                    f"{coordenada[1]},0"
                    for coordenada in geometria
                )
            )

            estilo = elemento(
                ponto,
                "Style"
            )

            linha_estilo = elemento(
                estilo,
                "LineStyle"
            )

            cor = (
                "ffff6600"
                if tipo == "OSRM"
                else "ff999999"
            )

            elemento(
                linha_estilo,
                "color",
                cor
            )

            elemento(
                linha_estilo,
                "width",
                "4"
            )

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
                lambda valor: (
                    " | ".join(
                        map(str, valor)
                    )
                    if isinstance(
                        valor,
                        list
                    )
                    else str(valor)
                    if isinstance(
                        valor,
                        dict
                    )
                    else valor
                )
            )

    return saida


def formatar_planilha(
    writer,
    aba,
    df
):
    ws = writer.sheets[
        aba
    ]

    cabecalho = PatternFill(
        "solid",
        fgColor="002060"
    )

    cor_super = PatternFill(
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

    indice_super = (
        df.columns.get_loc(
            "SUPER_PONTO"
        )
        if "SUPER_PONTO"
        in df.columns
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

    for linha in ws.iter_rows(
        min_row=2
    ):
        superponto = (
            str(
                linha[indice_super].value
            ).startswith("SIM")
            if indice_super is not None
            else False
        )

        prioridade = (
            str(
                linha[
                    indice_prioridade
                ].value
            ).upper() == "ALTA"
            if indice_prioridade is not None
            else False
        )

        for celula in linha:
            if superponto:
                celula.fill = cor_super

            if prioridade:
                celula.font = fonte_prioridade

    for coluna in ws.columns:
        letra = coluna[0].column_letter

        tamanhos = [
            len(
                str(
                    celula.value or ""
                )
            )
            for celula in list(
                coluna
            )[:120]
        ]

        ws.column_dimensions[
            letra
        ].width = min(
            60,
            max(
                13,
                max(
                    tamanhos,
                    default=10
                ) + 2
            )
        )


def gerar_excel(planilhas):
    memoria = io.BytesIO()

    with pd.ExcelWriter(
        memoria,
        engine="openpyxl"
    ) as writer:

        for nome, df in planilhas.items():
            aba = nome[:31]

            tabela = limpar_excel(
                df
            )

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


def gerar_zip_excel_por_equipe(
    programacao
):
    memoria = io.BytesIO()

    with zipfile.ZipFile(
        memoria,
        "w",
        zipfile.ZIP_DEFLATED
    ) as arquivo_zip:

        if not programacao.empty:
            grupos = programacao.groupby(
                [
                    "ATIVIDADE",
                    "EQUIPE_PROGRAMADA"
                ]
            )

            for (
                tipo,
                equipe
            ), grupo in grupos:

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


def gerar_zip_kml_por_equipe(
    obras,
    programacao,
    trajetos=None
):
    memoria = io.BytesIO()

    with zipfile.ZipFile(
        memoria,
        "w",
        zipfile.ZIP_DEFLATED
    ) as arquivo_zip:

        if not programacao.empty:
            alocadas = programacao[
                programacao[
                    "STATUS_PROGRAMACAO"
                ].eq("SUGESTAO")
            ]

            for (
                tipo,
                equipe
            ), grupo in alocadas.groupby(
                [
                    "ATIVIDADE",
                    "EQUIPE_PROGRAMADA"
                ]
            ):
                notas = set(
                    grupo[
                        "NOTA"
                    ].astype(str)
                )

                obras_equipe = obras[
                    obras[
                        "NOTA"
                    ].astype(str).isin(
                        notas
                    )
                ]

                rotas_equipe = (
                    trajetos[
                        trajetos[
                            "EQUIPE"
                        ].eq(equipe)
                    ]
                    if (
                        trajetos is not None
                        and not trajetos.empty
                    )
                    else pd.DataFrame()
                )

                nome_seguro = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    norm(equipe)
                )[:55]

                arquivo_zip.writestr(
                    (
                        f"{tipo}/"
                        f"Rota_{nome_seguro}.kml"
                    ),
                    gerar_kml(
                        obras_equipe,
                        trajetos=rotas_equipe,
                        nome=f"Rota - {equipe}"
                    )
                )

    return memoria.getvalue()


# ============================================================
# CABECALHO DO APLICATIVO
# ============================================================

st.markdown(
    """
    <div class="nip-title">
        📍 NIP | Análise Cruzada e Planejamento
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="nip-subtitle">
        Obras Pendentes | Programação |
        Superpontos | Auditoria | Excel e KML
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# REGRAS DE CONTAGEM
# ============================================================

with st.expander(
    "📘 Regras de contagem e exclusão",
    expanded=False
):
    st.markdown("""
    **STATUS SAP**

    - Somente FINL e CANC são excluídos.
    - Todos os demais status SAP são considerados.

    **STATUS LIST**

    - 0 é considerado.
    - Em levantamento é considerado.
    - Correção de levantamento é considerada.
    - Outros status são excluídos.

    **Contrato**

    - 0.
    - NIP GLOBAL LTDA - EQTL MARANHÃO.

    **Colunas M e N**

    - Precisam estar iguais a zero.

    **Colunas B e C**

    - Se B contém equipe diferente de SEM LEVANTADOR
      e C contém data válida, a nota é retirada
      da contagem pendente.
    - A nota permanece registrada na auditoria.

    **Duplicidade**

    - A nota presente nas duas bases conta uma vez.
    - Pode gerar duas tarefas, uma por atividade.

    **Distâncias**

    - Equipes são sugeridas inicialmente por
      proximidade geográfica.
    - OSRM é opcional para desenhar trajetos reais
      pelas ruas e estradas.
    """)


# ============================================================
# IMPORTACAO DOS ARQUIVOS
# ============================================================

secao("📂 Importação das Bases")

c1, c2, c3 = st.columns(3)

with c1:
    arquivo_san = st.file_uploader(
        "BASE_SANEAMENTO",
        type=[
            "xlsx",
            "csv"
        ],
        key="upload_saneamento"
    )

with c2:
    arquivo_lev = st.file_uploader(
        "BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"],
        key="upload_levantamento"
    )

with c3:
    arquivo_equipes = st.file_uploader(
        "LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"],
        key="upload_equipes"
    )

if not all([
    arquivo_san,
    arquivo_lev,
    arquivo_equipes
]):
    st.info(
        "Envie os três arquivos obrigatórios "
        "para iniciar a análise."
    )

    st.stop()


# ============================================================
# ASSINATURA DOS ARQUIVOS
# ============================================================

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


# ============================================================
# PROCESSAMENTO DAS BASES
# ============================================================

if st.button(
    "🚀 PROCESSAR BASES",
    type="primary",
    use_container_width=True
):
    inicio_processamento = (
        time.perf_counter()
    )

    try:
        with st.spinner(
            "Validando e cruzando as bases..."
        ):
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
                san,
                lev
            )

            if auditoria.empty:
                raise ValueError(
                    "Nenhuma nota válida encontrada."
                )

            tempo_total = (
                time.perf_counter()
                - inicio_processamento
            )

            st.session_state[
                "nip_base"
            ] = {
                "assinatura": assinatura,
                "auditoria": auditoria,
                "equipes": equipes,
                "aba_san": aba_san,
                "info_lev": info_lev,
                "qtd_san": len(san),
                "qtd_lev": len(lev),
                "tempo_processamento": tempo_total
            }

            st.session_state.pop(
                "nip_osrm",
                None
            )

        st.success(
            "Bases processadas em "
            + formatar_tempo(
                tempo_total
            )
        )

    except Exception as erro:
        st.error(
            f"Erro no processamento: {erro}"
        )


if "nip_base" not in st.session_state:
    st.stop()

dados = st.session_state[
    "nip_base"
]

if dados["assinatura"] != assinatura:
    st.warning(
        "Os arquivos foram alterados. "
        "Clique novamente em PROCESSAR BASES."
    )

    st.stop()

auditoria = dados[
    "auditoria"
]

equipes = dados[
    "equipes"
]

st.caption(
    f"Saneamento: aba {dados['aba_san']} | "
    f"Levantamento: aba "
    f"{dados['info_lev']['ABA']} | "
    f"Processamento: "
    f"{formatar_tempo(dados['tempo_processamento'])}"
)


# ============================================================
# CONFERENCIA DAS COLUNAS B E C
# ============================================================

with st.expander(
    "🔎 Conferência das colunas B e C"
):
    info = dados["info_lev"]

    st.write(
        "**Coluna B identificada:** "
        + info["COLUNA_B"]
    )

    st.write(
        "**Coluna C identificada:** "
        + info["COLUNA_C"]
    )

    st.write(
        "**Registros com equipe e data:** "
        + str(
            info["JA_EM_CAMPO_LINHAS"]
        )
    )

    st.caption(
        "A exclusão é aplicada somente quando "
        "as duas condições são atendidas. "
        "A volumetria final considera notas únicas."
    )


# ============================================================
# CONFIGURACOES LATERAIS
# ============================================================

with st.sidebar:
    st.header("⚙️ Planejamento")

    raio_equipes = st.select_slider(
        "Raio de referência das equipes (km)",
        options=list(
            range(200, 501, 25)
        ),
        value=200
    )

    quantidade_equipes = st.slider(
        "Equipes próximas por atividade",
        min_value=1,
        max_value=5,
        value=3
    )

    raio_super = st.slider(
        "Raio Superponto (metros)",
        min_value=10,
        max_value=500,
        value=50,
        step=10
    )

    dias_prioridade = st.number_input(
        "Dias para prioridade média",
        min_value=1,
        value=14
    )

    st.divider()

    st.subheader(
        "📅 Programação"
    )

    modo = st.selectbox(
        "Modo de programação",
        [
            "POR QUANTIDADE",
            "POR DIA",
            "POR SEMANA",
            "POR MES"
        ]
    )

    limite_periodo = st.number_input(
        "Quantidade por equipe no período",
        min_value=1,
        value=10
    )

    limite_dia = st.number_input(
        "Limite diário por equipe",
        min_value=1,
        value=8
    )

    periodos = st.number_input(
        "Quantidade de períodos",
        min_value=1,
        max_value=52,
        value=5
    )

    data_inicio = st.date_input(
        "Data inicial",
        value=datetime.today().date()
    )

    st.divider()

    st.subheader(
        "🛣️ OSRM"
    )

    usar_osrm = st.checkbox(
        "Traçado real por ruas no KML",
        value=False,
        help=(
            "Ative para consultar trajetos pelas "
            "ruas. Desativado, exporta somente "
            "os marcadores sem consultas OSRM."
        )
    )

    servidor_osrm = st.text_input(
        "Endpoint OSRM",
        value=OSRM_DEFAULT
    )

    lote_osrm = st.slider(
        "Consultas OSRM por lote",
        min_value=1,
        max_value=8,
        value=5
    )

    st.caption(
        "O OSRM é processado em pequenos lotes "
        "para evitar consultas longas."
    )


# ============================================================
# OBRAS PENDENTES E EXCLUIDAS
# ============================================================

obras = auditoria[
    auditoria[
        "PENDENTE_CONTAGEM"
    ].eq("SIM")
].copy()

duplicadas_total = auditoria[
    auditoria[
        "DUPLICADA"
    ].eq("SIM")
].copy()

duplicadas_pendentes = obras[
    obras[
        "DUPLICADA"
    ].eq("SIM")
].copy()

obras_finl = auditoria[
    auditoria[
        "MOTIVO_EXCLUSAO"
    ].astype(str).str.contains(
        r"(?:^|\s\|\s)SAP FINL(?:$|\s\|\s)",
        regex=True
    )
].copy()

obras_canc = auditoria[
    auditoria[
        "MOTIVO_EXCLUSAO"
    ].astype(str).str.contains(
        r"(?:^|\s\|\s)SAP CANC(?:$|\s\|\s)",
        regex=True
    )
].copy()

obras_campo = auditoria[
    auditoria[
        "JA_EM_CAMPO"
    ].eq("SIM")
].copy()

excluidas = auditoria[
    auditoria[
        "PENDENTE_CONTAGEM"
    ].eq("NÃO")
].copy()

obras = priorizar(
    obras,
    dias_prioridade
)


# ============================================================
# CARDS DE INDICADORES
# ============================================================

secao("📊 Indicadores Operacionais")

mostrar_cards([
    (
        "Saneamento",
        int(
            obras[
                "LISTA"
            ].eq(CAT_SAN).sum()
        ),
        "#2563EB",
        "Obras exclusivas"
    ),
    (
        "Levantamento",
        int(
            obras[
                "LISTA"
            ].eq(CAT_LEV).sum()
        ),
        "#16A34A",
        "Obras exclusivas"
    ),
    (
        "Duplicadas",
        len(
            duplicadas_pendentes
        ),
        "#9333EA",
        "Pendentes nas duas bases"
    ),
    (
        "Total Pendente",
        len(obras),
        "#0D256C",
        "Notas únicas"
    ),
    (
        "Excluídas",
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
        "Já em Campo",
        len(obras_campo),
        "#0891B2",
        "Equipe e data nas colunas B/C"
    ),
    (
        "Duplicadas Totais",
        len(duplicadas_total),
        "#9333EA",
        "Incluindo excluídas"
    )
])


# ============================================================
# FILTROS
# ============================================================

secao("🔎 Filtrar Obras")

c1, c2, c3 = st.columns(3)

regionais = sorted(
    obras[
        "REGIONAL"
    ].dropna().astype(str).unique()
)

municipios = sorted(
    obras[
        "MUNICIPIO"
    ].dropna().astype(str).unique()
)

with c1:
    filtro_regional = st.multiselect(
        "Regional",
        regionais
    )

with c2:
    filtro_municipio = st.multiselect(
        "Município",
        municipios
    )

with c3:
    filtro_prioridade = st.multiselect(
        "Prioridade",
        [
            "ALTA",
            "MEDIA",
            "BAIXA"
        ]
    )

pesquisa = st.text_input(
    "Pesquisar nota"
)

view = obras.copy()

if filtro_regional:
    view = view[
        view[
            "REGIONAL"
        ].isin(
            filtro_regional
        )
    ]

if filtro_municipio:
    view = view[
        view[
            "MUNICIPIO"
        ].isin(
            filtro_municipio
        )
    ]

if filtro_prioridade:
    view = view[
        view[
            "PRIORIDADE_PLANEJAMENTO"
        ].isin(
            filtro_prioridade
        )
    ]

if pesquisa.strip():
    view = view[
        view[
            "NOTA"
        ].astype(str).str.contains(
            re.escape(
                pesquisa.strip()
            ),
            case=False,
            na=False
        )
    ]


# ============================================================
# EQUIPES CANDIDATAS
# ============================================================

view, candidatos = gerar_candidatos(
    view,
    equipes,
    quantidade_equipes
)

tarefas = gerar_tarefas(
    view
)

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


# ============================================================
# QUADRO DE CAPACIDADE
# ============================================================

quadro_capacidade, detalhe_capacidade = (
    calcular_capacidade_media(
        view,
        equipes,
        programacao
    )
)

secao(
    "👥 Capacidade e Média de Obras por Equipe"
)

saneamento_resumo = quadro_capacidade[
    quadro_capacidade[
        "ATIVIDADE"
    ].eq("Saneamento")
].iloc[0]

levantamento_resumo = quadro_capacidade[
    quadro_capacidade[
        "ATIVIDADE"
    ].eq("Levantamento")
].iloc[0]

mostrar_cards([
    (
        "Média por Equipe - Saneamento",
        (
            f"{saneamento_resumo['MEDIA_OBRAS_POR_EQUIPE']:.2f}"
        ),
        "#2563EB",
        (
            f"{saneamento_resumo['OBRAS_DISPONIVEIS']} obras / "
            f"{saneamento_resumo['EQUIPES']} equipes"
        )
    ),
    (
        "Média por Equipe - Levantamento",
        (
            f"{levantamento_resumo['MEDIA_OBRAS_POR_EQUIPE']:.2f}"
        ),
        "#16A34A",
        (
            f"{levantamento_resumo['OBRAS_DISPONIVEIS']} obras / "
            f"{levantamento_resumo['EQUIPES']} equipes"
        )
    )
])

st.dataframe(
    quadro_capacidade,
    hide_index=True,
    use_container_width=True
)

with st.expander(
    "📋 Distribuição detalhada por equipe"
):
    st.dataframe(
        detalhe_capacidade,
        hide_index=True,
        use_container_width=True
    )


# ============================================================
# SUPERPONTOS
# ============================================================

if not programacao.empty:
    alocadas = programacao[
        programacao[
            "STATUS_PROGRAMACAO"
        ].eq("SUGESTAO")
    ].copy()

    superpontos, qtd_fundidas = (
        fundir_superpontos(
            alocadas,
            raio_super
        )
    )

else:
    superpontos = pd.DataFrame()
    qtd_fundidas = 0

qtd_superpontos = (
    int(
        superpontos[
            "SUPER_PONTO"
        ].astype(str).str.startswith(
            "SIM"
        ).sum()
    )
    if not superpontos.empty
    else 0
)


# ============================================================
# PREPARACAO DO OSRM
# ============================================================

trechos = preparar_trechos(
    programacao,
    equipes
)

if not trechos.empty:
    dados_assinatura = trechos[
        [
            "NOTA",
            "ATIVIDADE",
            "DATA",
            "ID_EQUIPE",
            "LAT_ORIGEM",
            "LON_ORIGEM",
            "LAT_DESTINO",
            "LON_DESTINO"
        ]
    ].to_json(
        orient="records"
    )

else:
    dados_assinatura = ""

assinatura_rotas = hashlib.sha256(
    (
        dados_assinatura
        + servidor_osrm
    ).encode()
).hexdigest()

if (
    "nip_osrm" not in st.session_state
    or st.session_state[
        "nip_osrm"
    ].get(
        "assinatura"
    ) != assinatura_rotas
):
    st.session_state[
        "nip_osrm"
    ] = {
        "assinatura": assinatura_rotas,
        "cursor": 0,
        "registros": [],
        "segundos_acumulados": 0.0
    }

estado_osrm = st.session_state[
    "nip_osrm"
]


# ============================================================
# CRONOMETRO OSRM
# ============================================================

if usar_osrm and not trechos.empty:
    secao(
        "⏱️ Tempo de Processamento OSRM"
    )

    total_trechos = len(
        trechos
    )

    processados = estado_osrm[
        "cursor"
    ]

    decorrido = estado_osrm.get(
        "segundos_acumulados",
        0.0
    )

    renderizar_cronometro(
        decorrido,
        processados,
        total_trechos
    )

    st.progress(
        processados / total_trechos
    )

    if processados < total_trechos:
        if st.button(
            "🛣️ PROCESSAR PRÓXIMO LOTE OSRM",
            type="primary",
            use_container_width=True
        ):
            novos, proximo, tempo_lote = (
                processar_lote_osrm(
                    trechos,
                    servidor_osrm,
                    processados,
                    lote_osrm
                )
            )

            estado_osrm[
                "registros"
            ].extend(novos)

            estado_osrm[
                "cursor"
            ] = proximo

            estado_osrm[
                "segundos_acumulados"
            ] += tempo_lote

            st.session_state[
                "nip_osrm"
            ] = estado_osrm

            st.rerun()

    else:
        st.success(
            "Todos os trechos foram processados."
        )

trajetos = pd.DataFrame(
    estado_osrm[
        "registros"
    ]
)

trajetos_exportar = (
    trajetos
    if usar_osrm
    else pd.DataFrame()
)


# ============================================================
# ABAS DE RESULTADOS
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


# ============================================================
# TABELAS DAS TRES CATEGORIAS
# ============================================================

for aba, categoria in zip(
    abas[:3],
    CATEGORIAS
):
    with aba:
        tabela = view[
            view[
                "LISTA"
            ].eq(categoria)
        ]

        st.subheader(
            f"{categoria} - "
            f"{len(tabela)} obras"
        )

        st.dataframe(
            tabela,
            hide_index=True,
            use_container_width=True
        )


# ============================================================
# PROGRAMACAO
# ============================================================

with abas[3]:
    st.subheader(
        "📅 Programação de Obras"
    )

    if not programacao.empty:
        quantidade_alocada = int(
            programacao[
                "STATUS_PROGRAMACAO"
            ].eq("SUGESTAO").sum()
        )

        quantidade_sem = int(
            programacao[
                "STATUS_PROGRAMACAO"
            ].eq("SEM ALOCACAO").sum()
        )

        mostrar_cards([
            (
                "Total de Tarefas",
                len(programacao),
                "#0D256C",
                "Saneamento e levantamento"
            ),
            (
                "Tarefas Alocadas",
                quantidade_alocada,
                "#16A34A",
                "Programação sugerida"
            ),
            (
                "Sem Alocação",
                quantidade_sem,
                "#DC2626",
                "Necessitam avaliação"
            )
        ])

    st.dataframe(
        programacao,
        hide_index=True,
        use_container_width=True
    )

    st.subheader(
        "Carga por Equipe"
    )

    st.dataframe(
        carga,
        hide_index=True,
        use_container_width=True
    )


# ============================================================
# CAPACIDADE
# ============================================================

with abas[4]:
    st.subheader(
        "👥 Capacidade por Equipe"
    )

    st.dataframe(
        detalhe_capacidade,
        hide_index=True,
        use_container_width=True
    )

    if not detalhe_capacidade.empty:
        grafico = detalhe_capacidade.set_index(
            "EQUIPE"
        )[
            "OBRAS_PROGRAMADAS"
        ]

        st.bar_chart(
            grafico
        )


# ============================================================
# SUPERPONTOS
# ============================================================

with abas[5]:
    st.subheader(
        "🏢 Superpontos"
    )

    mostrar_cards([
        (
            "Superpontos",
            qtd_superpontos,
            "#F59E0B",
            "Grupos com mais de uma tarefa"
        ),
        (
            "Tarefas Fundidas",
            qtd_fundidas,
            "#9333EA",
            "Redução de marcadores"
        ),
        (
            "Raio",
            f"{raio_super} metros",
            "#0D256C",
            "DBSCAN Haversine"
        )
    ])

    if not superpontos.empty:
        st.dataframe(
            limpar_excel(
                superpontos
            ),
            hide_index=True,
            use_container_width=True
        )


# ============================================================
# AUDITORIA
# ============================================================

with abas[6]:
    subabas = st.tabs([
        "Excluídas",
        "Já em Campo",
        "SAP FINL",
        "SAP CANC",
        "Duplicadas Totais"
    ])

    tabelas = [
        excluidas,
        obras_campo,
        obras_finl,
        obras_canc,
        duplicadas_total
    ]

    for subaba, tabela in zip(
        subabas,
        tabelas
    ):
        with subaba:
            st.dataframe(
                tabela,
                hide_index=True,
                use_container_width=True
            )


# ============================================================
# MAPA INTERATIVO
# ============================================================

with abas[7]:
    st.subheader(
        "🗺️ Mapa Operacional"
    )

    if st.checkbox(
        "Carregar mapa interativo"
    ):
        geos = view.dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        if geos.empty:
            st.warning(
                "Nenhuma obra com coordenadas válidas."
            )

        else:
            mapa = folium.Map(
                location=[
                    float(
                        geos[
                            "LATITUDE"
                        ].mean()
                    ),
                    float(
                        geos[
                            "LONGITUDE"
                        ].mean()
                    )
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
                    geos[
                        "LISTA"
                    ].eq(categoria)
                ]

                for _, obra in subset.iterrows():
                    folium.Marker(
                        [
                            float(
                                obra[
                                    "LATITUDE"
                                ]
                            ),
                            float(
                                obra[
                                    "LONGITUDE"
                                ]
                            )
                        ],
                        popup=folium.Popup(
                            popup_html(obra),
                            max_width=440
                        ),
                        icon=folium.Icon(
                            color=CORES_FOLIUM[
                                categoria
                            ]
                        )
                    ).add_to(cluster)

                camada.add_to(
                    mapa
                )

            # Trajetos OSRM, quando disponíveis.
            if not trajetos_exportar.empty:
                camada_rotas = folium.FeatureGroup(
                    name="Trajetos OSRM"
                )

                for _, trecho in trajetos_exportar.iterrows():
                    geometria = trecho.get(
                        "GEOMETRIA"
                    )

                    if not isinstance(
                        geometria,
                        list
                    ):
                        continue

                    if len(geometria) < 2:
                        continue

                    cor_rota = (
                        "blue"
                        if trecho.get(
                            "TIPO_TRAJETO"
                        ) == "OSRM"
                        else "gray"
                    )

                    folium.PolyLine(
                        [
                            [
                                ponto[1],
                                ponto[0]
                            ]
                            for ponto in geometria
                        ],
                        color=cor_rota,
                        weight=3
                    ).add_to(
                        camada_rotas
                    )

                camada_rotas.add_to(
                    mapa
                )

            folium.LayerControl().add_to(
                mapa
            )

            st_folium(
                mapa,
                use_container_width=True,
                height=600
            )


# ============================================================
# EXPORTACOES
# ============================================================

secao(
    "📥 Exportação de Planilhas e KML"
)

resumo_geral = pd.DataFrame([
    [
        "Saneamento",
        int(
            obras[
                "LISTA"
            ].eq(CAT_SAN).sum()
        )
    ],
    [
        "Levantamento",
        int(
            obras[
                "LISTA"
            ].eq(CAT_LEV).sum()
        )
    ],
    [
        "Duplicadas Pendentes",
        len(
            duplicadas_pendentes
        )
    ],
    [
        "Total Pendente",
        len(obras)
    ],
    [
        "Total Excluídas",
        len(excluidas)
    ],
    [
        "Já em Campo",
        len(obras_campo)
    ],
    [
        "SAP FINL",
        len(obras_finl)
    ],
    [
        "SAP CANC",
        len(obras_canc)
    ]
], columns=[
    "INDICADOR",
    "VALOR"
])


# ============================================================
# EXCEL GERAL
# ============================================================

excel_geral = gerar_excel({
    "RESUMO": resumo_geral,
    "CONSOLIDADO": obras,
    "SANEAMENTO": obras[
        obras[
            "LISTA"
        ].eq(CAT_SAN)
    ],
    "LEVANTAMENTO": obras[
        obras[
            "LISTA"
        ].eq(CAT_LEV)
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


# ============================================================
# EXPORTACOES ESPECIFICAS
# ============================================================

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


# ============================================================
# KML
# ============================================================

kml_geral = gerar_kml(
    view,
    superpontos,
    trajetos_exportar
)

kml_duplicadas = gerar_kml(
    duplicadas_total,
    nome="NIP - Obras Duplicadas"
)


# ============================================================
# ZIP POR EQUIPE
# ============================================================

zip_excel_equipes = gerar_zip_excel_por_equipe(
    programacao
)

zip_kml_equipes = gerar_zip_kml_por_equipe(
    view,
    programacao,
    trajetos_exportar
)


# ============================================================
# BOTOES - ARQUIVOS GERAIS
# ============================================================

st.subheader(
    "Arquivos Gerais"
)

c1, c2, c3, c4 = st.columns(4)

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
        "📦 EXCEL POR EQUIPE",
        data=zip_excel_equipes,
        file_name="NIP_Equipes_Excel.zip",
        mime="application/zip",
        use_container_width=True
    )

with c4:
    st.download_button(
        "📦 KML POR EQUIPE",
        data=zip_kml_equipes,
        file_name="NIP_Equipes_KML.zip",
        mime="application/zip",
        use_container_width=True
    )


# ============================================================
# BOTOES - ARQUIVOS ESPECIFICOS
# ============================================================

st.subheader(
    "Exportações Específicas"
)

c1, c2, c3 = st.columns(3)

with c1:
    st.download_button(
        "🔵 DUPLICADAS - EXCEL",
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
        "⚪ FINL - EXCEL",
        data=excel_finl,
        file_name="NIP_Obras_FINL.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.download_button(
        "🟠 CANC - EXCEL",
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
        "🚗 JÁ EM CAMPO - EXCEL",
        data=excel_campo,
        file_name="NIP_Obras_Ja_Em_Campo.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )


# ============================================================
# RODAPE
# ============================================================

st.divider()

st.caption(
    f"Base Saneamento: "
    f"{dados['qtd_san']} linhas | "
    f"Base Levantamento: "
    f"{dados['qtd_lev']} linhas | "
    f"Notas únicas analisadas: "
    f"{len(auditoria)}"
)
