
import io
import re
import html
import json
import hashlib
import zipfile
import unicodedata
import xml.etree.ElementTree as ET

from collections import defaultdict
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import requests
import streamlit as st
import folium

from sklearn.cluster import DBSCAN
from sklearn.neighbors import BallTree
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment
)

# ============================================================
# NIP - ANALISE CRUZADA E LISTA CONTINUA
# ============================================================

st.set_page_config(
    page_title="NIP | Analise e Planejamento",
    page_icon="📍",
    layout="wide"
)

EARTH_KM = 6371.0088

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_DUP = "OBRA SANEAMENTO E LEVANTAMENTO"

CATEGORIAS = [CAT_SAN, CAT_LEV, CAT_DUP]

CORES_FOLIUM = {
    CAT_SAN: "blue",
    CAT_LEV: "green",
    CAT_DUP: "purple"
}

CORES_HEX = {
    CAT_SAN: "#2563eb",
    CAT_LEV: "#16a34a",
    CAT_DUP: "#9333ea"
}

SAP_BLOQUEADOS = {"FINL", "CANC"}

LIST_VALIDOS = {
    "0",
    "EM LEVANTAMENTO",
    "CORRECAO DE LEVANTAMENTO"
}

CONTRATOS_VALIDOS = {
    "0",
    "NIP GLOBAL LTDA - EQTL MARANHAO"
}

OSRM_DEFAULT = "https://router.project-osrm.org"

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
# ESTILO VISUAL
# ============================================================

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 100%;
}

.nip-title {
    color: #0D256C;
    font-size: 2rem;
    font-weight: 800;
    margin-bottom: 0.2rem;
}

.nip-subtitle {
    color: #64748b;
    font-size: 0.95rem;
    margin-bottom: 1rem;
}

.nip-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-left: 5px solid var(--cor);
    border-radius: 10px;
    padding: 17px;
    margin-bottom: 10px;
    box-shadow: 0 2px 5px rgba(0,0,0,0.05);
}

.nip-card-label {
    color: #64748b;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
}

.nip-card-value {
    color: #172554;
    font-size: 29px;
    font-weight: 800;
}

.nip-card-caption {
    color: #64748b;
    font-size: 11px;
}

.nip-section {
    color: #0D256C;
    font-size: 1.3rem;
    font-weight: 750;
    margin-top: 12px;
    margin-bottom: 12px;
}

div[data-testid="stDownloadButton"] button {
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)


def card(titulo, valor, cor="#0D256C", legenda=""):
    return f"""
    <div class="nip-card" style="--cor:{cor}">
        <div class="nip-card-label">
            {html.escape(str(titulo))}
        </div>
        <div class="nip-card-value">
            {html.escape(str(valor))}
        </div>
        <div class="nip-card-caption">
            {html.escape(str(legenda))}
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


def titulo_secao(texto):
    st.markdown(
        f'<div class="nip-section">{html.escape(texto)}</div>',
        unsafe_allow_html=True
    )


# ============================================================
# NORMALIZACAO
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


def procurar_coluna(df, alternativas, obrigatoria=True):
    mapa = {
        norm(coluna): coluna
        for coluna in df.columns
    }

    for nome in alternativas:
        if norm(nome) in mapa:
            return mapa[norm(nome)]

    if obrigatoria:
        raise ValueError(
            "Coluna nao encontrada: "
            + " / ".join(alternativas)
        )

    return None


def escolher_aba(arquivo, grupos_colunas, preferida=None):
    """
    Identifica automaticamente a aba pelos cabecalhos.
    Nao depende exclusivamente do nome NOTAS.
    """
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
        for linha_cabecalho in [0, 1, 2, 3, 4]:
            try:
                amostra = pd.read_excel(
                    excel,
                    sheet_name=aba,
                    header=linha_cabecalho,
                    nrows=3,
                    dtype=str
                )

                nomes = {
                    norm(c) for c in amostra.columns
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
                        header=linha_cabecalho,
                        dtype=str
                    )

                    return df, aba

            except Exception:
                continue

    raise ValueError(
        "Nao foi encontrada uma aba compativel. "
        "Abas existentes: "
        + ", ".join(excel.sheet_names)
    )


def ler_saneamento(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        df = pd.read_csv(
            io.BytesIO(arquivo.getvalue()),
            sep=None,
            engine="python",
            dtype=str
        )

        aba = "CSV"

    else:
        df, aba = escolher_aba(
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

    validar_coordenadas_df(
        df,
        "LAT_OBRA",
        "LON_OBRA"
    )

    return df, aba


def converter_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(
            ",",
            ".",
            regex=False
        ),
        errors="coerce"
    )


def validar_coordenadas_df(df, c_lat, c_lon):
    valido = (
        df[c_lat].between(-35, 6)
        & df[c_lon].between(-75, -30)
        & df[c_lat].ne(0)
        & df[c_lon].ne(0)
    )

    df.loc[~valido, c_lat] = np.nan
    df.loc[~valido, c_lon] = np.nan


def motivos_levantamento(linha):
    motivos = []

    if linha["SAP_NORM"] in SAP_BLOQUEADOS:
        motivos.append(
            "SAP " + linha["SAP_NORM"]
        )

    if linha["LIST_NORM"] not in LIST_VALIDOS:
        motivos.append(
            "LIST NAO ACEITO: "
            + (linha["LIST_NORM"] or "VAZIO")
        )

    if linha["CONTRATO_NORM"] not in CONTRATOS_VALIDOS:
        motivos.append(
            "CONTRATO NAO ACEITO: "
            + (linha["CONTRATO_NORM"] or "VAZIO")
        )

    if linha["M_NORM"] != "0":
        motivos.append(
            "ORCAMENTO MODULAR DIFERENTE DE 0"
        )

    if linha["N_NORM"] != "0":
        motivos.append(
            "PLA ALVOS DIFERENTE DE 0"
        )

    return " | ".join(motivos)


def ler_levantamento(arquivo):
    df, aba = escolher_aba(
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
        df, ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
    )

    c_lat = procurar_coluna(df, ["LATITUDE"])
    c_lon = procurar_coluna(df, ["LONGITUDE"])

    c_regional = procurar_coluna(
        df, ["REGIONAL"], False
    )

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
            "Nao foram localizadas as colunas "
            "ORCAMENTO MODULAR e PLA ALVOS."
        )

    c_prioridade = procurar_coluna(
        df, ["PRIORIDADE"], False
    )

    c_data = procurar_coluna(
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

    validar_coordenadas_df(
        df,
        "LAT_OBRA",
        "LON_OBRA"
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
        df[c_data].fillna("")
        if c_data is not None else ""
    )

    df["MOTIVOS_EXCLUSAO"] = df.apply(
        motivos_levantamento,
        axis=1
    )

    return df, aba


def ler_equipes(arquivo):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    partes = []

    for tipo, aba_esperada in [
        ("LEVANTAMENTO", "LEVANTADORES"),
        ("SANEAMENTO", "SANEAMENTO")
    ]:
        aba = next(
            (
                a for a in excel.sheet_names
                if norm(a) == aba_esperada
            ),
            None
        )

        if aba is None:
            raise ValueError(
                "Aba nao encontrada: " + aba_esperada
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

        validar_coordenadas_df(
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
        equipes["EQUIPE"].astype(str).str.strip()
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
            "Nenhuma equipe com coordenadas validas."
        )

    return equipes


# ============================================================
# CRUZAMENTO
# ============================================================

def primeira_valida(grupo):
    valido = (
        grupo["LAT_OBRA"].notna()
        & grupo["LON_OBRA"].notna()
    )

    if valido.any():
        return grupo.loc[valido].iloc[0]

    return grupo.iloc[0]


def consolidar(saneamento, levantamento):
    san = saneamento[
        saneamento["CHAVE_NOTA"].ne("")
    ].copy()

    lev = levantamento[
        levantamento["CHAVE_NOTA"].ne("")
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

    linhas = []

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

        origem = a if a is not None else b

        if (
            b is not None
            and pd.notna(b["LAT_OBRA"])
            and pd.notna(b["LON_OBRA"])
        ):
            origem = b

        linhas.append({
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

    return pd.DataFrame(linhas)


# ============================================================
# DISTANCIAS E EQUIPES PROXIMAS
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    a1 = np.radians(lat1)
    b1 = np.radians(lon1)
    a2 = np.radians(lat2)
    b2 = np.radians(lon2)

    a = (
        np.sin((a2 - a1) / 2) ** 2
        + np.cos(a1)
        * np.cos(a2)
        * np.sin((b2 - b1) / 2) ** 2
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
    """
    Busca vetorizada das equipes geograficamente
    mais proximas, sem consultar rede.
    """
    saida = obras.copy().reset_index(drop=True)

    for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
        saida[f"EQUIPES_{tipo}"] = "NAO APLICAVEL"

    registros = []

    for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
        equipes_tipo = equipes[
            equipes["TIPO_EQUIPE"].eq(tipo)
        ].reset_index(drop=True)

        if equipes_tipo.empty:
            continue

        arvore = BallTree(
            np.radians(
                equipes_tipo[
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
            k=min(quantidade, len(equipes_tipo))
        )

        for posicao, i_obra in enumerate(indices_obras):
            partes = []

            for posicao_eq in range(indices.shape[1]):
                equipe = equipes_tipo.iloc[
                    indices[posicao, posicao_eq]
                ]

                km = float(
                    distancias[posicao, posicao_eq]
                    * EARTH_KM
                )

                registros.append({
                    "NOTA": saida.at[i_obra, "NOTA"],
                    "ATIVIDADE": tipo,
                    "ID_EQUIPE": equipe["ID_EQUIPE"],
                    "EQUIPE": equipe["EQUIPE"],
                    "CIDADE_BASE": equipe["CIDADE_BASE"],
                    "LAT_EQUIPE": equipe["LAT_EQUIPE"],
                    "LON_EQUIPE": equipe["LON_EQUIPE"],
                    "LAT_OBRA": saida.at[i_obra, "LATITUDE"],
                    "LON_OBRA": saida.at[i_obra, "LONGITUDE"],
                    "DISTANCIA_RETA_KM": round(km, 2)
                })

                partes.append(
                    f"{equipe['EQUIPE']} "
                    f"({equipe['CIDADE_BASE']}) "
                    f"- {km:.1f} km"
                )

            saida.at[
                i_obra, f"EQUIPES_{tipo}"
            ] = " | ".join(partes)

    return saida, pd.DataFrame(registros)


# ============================================================
# OSRM - REDE COM TIMEOUT CURTO E CACHE
# ============================================================

@st.cache_data(
    ttl=86400,
    show_spinner=False,
    max_entries=3000
)
def consultar_osrm(
    lat1,
    lon1,
    lat2,
    lon2,
    servidor,
    timeout_s=4
):
    """
    Retorna geometria real e distancia rodoviaria.
    Uma falha nao gera distancia rodoviaria ficticia.
    """
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
                float(rota["distance"]) / 1000,
                2
            ),
            "MIN": round(
                float(rota["duration"]) / 60,
                1
            ),
            "GEOMETRIA": geometria
        }

    except (
        requests.RequestException,
        ValueError,
        KeyError,
        TypeError
    ):
        return {"OK": False}


def preparar_trechos(programacao, equipes):
    """
    Ordena as visitas por proximidade a partir
    da cidade-base de cada equipe.
    A sequencia e heuristica, nao um TSP exato.
    """
    linhas = []

    if programacao.empty:
        return pd.DataFrame()

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
        equipe = equipes[
            equipes["ID_EQUIPE"].eq(id_equipe)
        ]

        if equipe.empty:
            continue

        eq = equipe.iloc[0]

        lat_atual = float(eq["LAT_EQUIPE"])
        lon_atual = float(eq["LON_EQUIPE"])

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
                "LAT_DESTINO": float(obra["LATITUDE"]),
                "LON_DESTINO": float(obra["LONGITUDE"]),
                "KM_RETA": round(distancias[indice], 2)
            })

            lat_atual = float(obra["LATITUDE"])
            lon_atual = float(obra["LONGITUDE"])

    return pd.DataFrame(linhas)


def processar_lote_osrm(trechos, servidor, inicio, lote):
    """
    Processa somente uma pequena quantidade por clique,
    evitando longas requisicoes bloqueantes.
    """
    registros = []

    fim = min(inicio + lote, len(trechos))

    for i in range(inicio, fim):
        r = trechos.iloc[i]

        resposta = consultar_osrm(
            round(float(r["LAT_ORIGEM"]), 6),
            round(float(r["LON_ORIGEM"]), 6),
            round(float(r["LAT_DESTINO"]), 6),
            round(float(r["LON_DESTINO"]), 6),
            servidor
        )

        registro = r.to_dict()

        if resposta["OK"]:
            registro["TIPO_TRAJETO"] = "OSRM"
            registro["KM_RODOVIARIO"] = resposta["KM"]
            registro["TEMPO_MIN"] = resposta["MIN"]
            registro["GEOMETRIA"] = resposta["GEOMETRIA"]

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

        registros.append(registro)

    return registros, fim


# ============================================================
# DISTRIBUICAO DE TAREFAS E PROGRAMACAO
# ============================================================

def gerar_tarefas(obras):
    linhas = []

    for _, obra in obras.iterrows():
        for tipo in atividades(obra["LISTA"]):
            r = obra.to_dict()
            r["ATIVIDADE"] = tipo
            linhas.append(r)

    return pd.DataFrame(linhas)


def priorizar(obras, dias_media):
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

    def prioridade(r):
        original = norm(r["PRIORIDADE_ORIGINAL"])

        if any(
            x in original
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
        d.date()
        for d in datas
        if d.weekday() < 5
    ]

    if modo == "POR DIA":
        return dias_uteis[:periodos]

    return dias_uteis


def periodo_de(data, modo):
    if modo == "POR SEMANA":
        iso = data.isocalendar()
        return f"{iso.year}-S{iso.week:02d}"

    if modo == "POR MES":
        return data.strftime("%Y-%m")

    return data.isoformat()


def programar_tarefas(
    tarefas,
    candidatos,
    raio_equipe,
    data_inicio,
    modo,
    periodos,
    limite_periodo,
    limite_dia,
    distancia_rodoviaria=None
):
    """
    Distribui tarefas considerando capacidade e
    candidatos proximos.

    A distancia inicial de triagem e geodesica.
    Distancias OSRM consultadas podem substituir
    a referencia quando disponiveis.
    """
    if tarefas.empty:
        return pd.DataFrame(), pd.DataFrame()

    ordem_prioridade = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    tarefas = tarefas.copy()

    tarefas["_ORDEM_PRIORIDADE"] = (
        tarefas["PRIORIDADE_PLANEJAMENTO"]
        .map(ordem_prioridade)
        .fillna(3)
    )

    tarefas = tarefas.sort_values(
        [
            "_ORDEM_PRIORIDADE",
            "MUNICIPIO",
            "NOTA"
        ]
    )

    datas = gerar_datas(
        data_inicio,
        modo,
        periodos
    )

    ocupacao_dia = defaultdict(int)
    ocupacao_periodo = defaultdict(int)

    candidatos_por_tarefa = defaultdict(list)

    for _, c in candidatos.iterrows():
        candidatos_por_tarefa[
            (c["NOTA"], c["ATIVIDADE"])
        ].append(c.to_dict())

    resultado = []

    for _, tarefa in tarefas.iterrows():
        chave = (
            tarefa["NOTA"],
            tarefa["ATIVIDADE"]
        )

        candidatos_tarefa = candidatos_por_tarefa.get(
            chave, []
        )

        alternativas = []

        for candidato in candidatos_tarefa:
            km = float(
                candidato["DISTANCIA_RETA_KM"]
            )

            if km > raio_equipe:
                continue

            for data in datas:
                periodo = periodo_de(data, modo)

                chave_dia = (
                    candidato["ID_EQUIPE"],
                    data
                )

                chave_periodo = (
                    candidato["ID_EQUIPE"],
                    periodo
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
                    "CANDIDATO": candidato,
                    "DATA": data,
                    "PERIODO": periodo,
                    "CARGA_DIA": carga_dia,
                    "CARGA_PERIODO": carga_periodo,
                    "KM": km
                })

                break

        registro = tarefa.drop(
            labels=["_ORDEM_PRIORIDADE"]
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
                "DIA_SEMANA": DIAS_PT[data.weekday()],
                "DIA_MES": data.strftime("%d/%m/%Y"),
                "SEMANA": (
                    f"{iso.year}-S{iso.week:02d}"
                ),
                "MES": data.strftime("%Y-%m"),
                "PERIODO": periodo,
                "DISTANCIA_REFERENCIA_KM": melhor["KM"],
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

        resultado.append(registro)

    programacao = pd.DataFrame(resultado)

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
# SUPERPONTOS - MESMA LOGICA DBSCAN DA LISTA CONTINUA
# ============================================================

def fundir_super_pontos(
    df_tarefas,
    raio_metros=50,
    agrupar_por_equipe=True
):
    """
    Adaptacao da funcao fundir_super_pontos
    recebida no modulo geospatial.py.

    DBSCAN Haversine:
    - eps em metros
    - min_samples=1
    - centro na media das coordenadas
    - preserva _ORIGINAL_ROWS
    """
    if df_tarefas.empty:
        return df_tarefas.copy(), 0

    validas = df_tarefas.dropna(
        subset=["LATITUDE", "LONGITUDE"]
    ).copy()

    invalidas = df_tarefas[
        df_tarefas["LATITUDE"].isna()
        | df_tarefas["LONGITUDE"].isna()
    ].copy()

    if validas.empty:
        saida = invalidas.copy()
        saida["SUPER_PONTO"] = "NÃO"
        saida["_ORIGINAL_ROWS"] = [
            [r] for r in saida.to_dict("records")
        ]
        return saida, 0

    # A Lista Continua utiliza DBSCAN Haversine.
    coords = np.radians(
        validas[
            ["LATITUDE", "LONGITUDE"]
        ].to_numpy(float)
    )

    db = DBSCAN(
        eps=raio_metros / 6371000.0,
        min_samples=1,
        algorithm="ball_tree",
        metric="haversine"
    ).fit(coords)

    validas["CLUSTER_ID"] = db.labels_

    if agrupar_por_equipe:
        # Impede agrupamentos entre equipes e dias
        # diferentes.
        validas["CLUSTER_GRP"] = (
            validas["CLUSTER_ID"].astype(str)
            + "|"
            + validas["ID_EQUIPE"].astype(str)
            + "|"
            + validas["DATA_PROGRAMADA"].astype(str)
            + "|"
            + validas["ATIVIDADE"].astype(str)
        )

    else:
        validas["CLUSTER_GRP"] = (
            validas["CLUSTER_ID"].astype(str)
        )

    saida = []

    for _, grupo in validas.groupby(
        "CLUSTER_GRP",
        sort=False
    ):
        base = grupo.iloc[0].copy()
        originais = grupo.to_dict("records")
        quantidade = len(originais)

        base["_ORIGINAL_ROWS"] = originais

        base["LATITUDE"] = (
            grupo["LATITUDE"].mean()
        )

        base["LONGITUDE"] = (
            grupo["LONGITUDE"].mean()
        )

        if quantidade > 1:
            base["SUPER_PONTO"] = (
                f"SIM ({quantidade} un.)"
            )

            base["NOTAS_AGRUPADAS"] = " | ".join(
                grupo["NOTA"].astype(str).tolist()
            )

        else:
            base["SUPER_PONTO"] = "NÃO"
            base["NOTAS_AGRUPADAS"] = (
                str(base["NOTA"])
            )

        saida.append(base)

    final = pd.DataFrame(saida)

    if not invalidas.empty:
        invalidas["SUPER_PONTO"] = "NÃO"
        invalidas["NOTAS_AGRUPADAS"] = (
            invalidas["NOTA"].astype(str)
        )
        invalidas["_ORIGINAL_ROWS"] = [
            [r] for r in invalidas.to_dict("records")
        ]

        final = pd.concat(
            [final, invalidas],
            ignore_index=True
        )

    final = final.drop(
        columns=["CLUSTER_ID", "CLUSTER_GRP"],
        errors="ignore"
    )

    return final, len(df_tarefas) - len(final)


# ============================================================
# POPUPS PARA MAPA E KML
# ============================================================

def popup_html(obra, titulo_extra=""):
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
        ("EQUIPE_PROGRAMADA", "Equipe programada"),
        ("DATA_PROGRAMADA", "Data programada"),
        ("SUPER_PONTO", "Superponto"),
        ("NOTAS_AGRUPADAS", "Notas agrupadas")
    ]:
        valor = esc(campo)

        if valor and valor.lower() != "nan":
            linhas.append((rotulo, valor))

    conteudo = "".join(
        f"""
        <tr>
            <td style="
                padding:5px;
                font-weight:bold;
                vertical-align:top;
                width:110px;
            ">{rotulo}</td>
            <td style="padding:5px;">{valor}</td>
        </tr>
        """
        for rotulo, valor in linhas
    )

    cor = CORES_HEX.get(
        obra.get("LISTA"),
        "#0D256C"
    )

    titulo = (
        titulo_extra
        if titulo_extra
        else "INFORMACOES DA OBRA"
    )

    return f"""
    <div style="
        width:340px;
        max-width:100%;
        font-family:Arial,sans-serif;
        font-size:12px;
        color:#1f2937;
    ">
        <div style="
            background:{cor};
            padding:11px;
            color:white;
            font-weight:bold;
            border-radius:7px 7px 0 0;
        ">
            {html.escape(titulo)}
        </div>
        <table style="
            width:100%;
            border-collapse:collapse;
            background:#ffffff;
        ">
            {conteudo}
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
    ns = "http://www.opengis.net/kml/2.2"

    ET.register_namespace("", ns)

    def elemento(pai, tag, texto=None):
        el = ET.SubElement(
            pai,
            f"{{{ns}}}{tag}"
        )

        if texto is not None:
            el.text = str(texto)

        return el

    raiz = ET.Element(f"{{{ns}}}kml")
    documento = elemento(raiz, "Document")

    elemento(documento, "name", nome)

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
        style = elemento(documento, "Style")
        style.set("id", id_estilo)

        icon_style = elemento(
            style, "IconStyle"
        )
        elemento(icon_style, "scale", "1.2")

        icone = elemento(icon_style, "Icon")
        elemento(icone, "href", url)

        label = elemento(
            style, "LabelStyle"
        )
        elemento(label, "scale", "0")

    # Obras individuais por categoria.
    for categoria in CATEGORIAS:
        pasta = elemento(documento, "Folder")
        elemento(pasta, "name", categoria)

        dados = obras[
            obras["LISTA"].eq(categoria)
        ].dropna(
            subset=["LATITUDE", "LONGITUDE"]
        )

        id_estilo = estilos[categoria][0]

        for _, obra in dados.iterrows():
            ponto = elemento(
                pasta, "Placemark"
            )

            elemento(
                ponto,
                "name",
                f"NOTA {obra['NOTA']}"
            )

            elemento(
                ponto,
                "styleUrl",
                "#" + id_estilo
            )

            elemento(
                ponto,
                "description",
                popup_html(obra)
            )

            geometria = elemento(
                ponto, "Point"
            )

            elemento(
                geometria,
                "coordinates",
                (
                    f"{float(obra['LONGITUDE']):.8f},"
                    f"{float(obra['LATITUDE']):.8f},0"
                )
            )

    # Pasta independente para superpontos.
    if superpontos is not None and not superpontos.empty:
        pasta_super = elemento(
            documento, "Folder"
        )

        elemento(
            pasta_super,
            "name",
            "SUPERPONTOS - PROGRAMACAO"
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

            pm = elemento(
                pasta_super, "Placemark"
            )

            elemento(
                pm,
                "name",
                f"SUPERPONTO - {r['SUPER_PONTO']}"
            )

            elemento(
                pm,
                "styleUrl",
                "#"
                + estilos.get(
                    r["LISTA"],
                    estilos[CAT_DUP]
                )[0]
            )

            elemento(
                pm,
                "description",
                popup_html(
                    r,
                    "SUPERPONTO - NOTAS AGRUPADAS"
                )
            )

            ponto = elemento(pm, "Point")

            elemento(
                ponto,
                "coordinates",
                (
                    f"{float(r['LONGITUDE'])},"
                    f"{float(r['LATITUDE'])},0"
                )
            )

    # Rotas calculadas ou estimadas.
    if trajetos is not None and not trajetos.empty:
        pasta_rotas = elemento(
            documento, "Folder"
        )

        elemento(
            pasta_rotas,
            "name",
            "TRAJETOS DAS EQUIPES"
        )

        for _, trecho in trajetos.iterrows():
            geometria = trecho.get("GEOMETRIA")

            if not isinstance(geometria, list):
                continue

            if len(geometria) < 2:
                continue

            pm = elemento(
                pasta_rotas,
                "Placemark"
            )

            tipo = trecho.get(
                "TIPO_TRAJETO",
                "LINHA RETA"
            )

            elemento(
                pm,
                "name",
                (
                    f"{trecho['EQUIPE']} - "
                    f"{trecho['ORDEM']} - {tipo}"
                )
            )

            descricao = (
                f"Equipe: {trecho['EQUIPE']}\n"
                f"Nota: {trecho['NOTA']}\n"
                f"Tipo: {tipo}\n"
                f"Distancia rodoviaria: "
                f"{texto_seguro(trecho.get('KM_RODOVIARIO'))}"
            )

            elemento(
                pm,
                "description",
                descricao
            )

            linha = elemento(
                pm,
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
                    f"{pt[0]},{pt[1]},0"
                    for pt in geometria
                )
            )

            estilo = elemento(
                pm, "Style"
            )

            estilo_linha = elemento(
                estilo, "LineStyle"
            )

            # KML usa ordem AABBGGRR.
            if tipo == "OSRM":
                cor = "ffff6600"
            else:
                cor = "ff999999"

            elemento(
                estilo_linha,
                "color",
                cor
            )

            elemento(
                estilo_linha,
                "width",
                "4"
            )

    return ET.tostring(
        raiz,
        encoding="utf-8",
        xml_declaration=True
    )


# ============================================================
# EXPORTACAO EXCEL - PADRAO LISTA CONTINUA
# ============================================================

def limpar_para_excel(df):
    if df is None or df.empty:
        return pd.DataFrame()

    saida = df.copy()

    remover = [
        "_ORIGINAL_ROWS",
        "ROTA_GEOMETRIA",
        "GEOMETRIA",
        "CLUSTER_ID",
        "CLUSTER_GRP"
    ]

    saida = saida.drop(
        columns=remover,
        errors="ignore"
    )

    for coluna in saida.columns:
        if saida[coluna].dtype == "object":
            saida[coluna] = saida[coluna].map(
                lambda v: (
                    " | ".join(map(str, v))
                    if isinstance(v, list)
                    else str(v)
                    if isinstance(v, dict)
                    else v
                )
            )

    return saida


def formatar_excel(writer, nome_aba, df):
    ws = writer.sheets[nome_aba]

    cor_cabecalho = PatternFill(
        "solid",
        fgColor="002060"
    )

    cor_super = PatternFill(
        "solid",
        fgColor="FCE4D6"
    )

    fonte_cabecalho = Font(
        name="Calibri",
        bold=True,
        color="FFFFFF"
    )

    fonte_prioridade = Font(
        name="Calibri",
        bold=True,
        color="FF0000"
    )

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions

    for celula in ws[1]:
        celula.fill = cor_cabecalho
        celula.font = fonte_cabecalho
        celula.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 26

    col_sp = (
        df.columns.get_loc("SUPER_PONTO")
        if "SUPER_PONTO" in df.columns
        else None
    )

    col_prioridade = (
        df.columns.get_loc(
            "PRIORIDADE_PLANEJAMENTO"
        )
        if "PRIORIDADE_PLANEJAMENTO"
        in df.columns
        else None
    )

    for linha_excel in ws.iter_rows(min_row=2):
        destacar_sp = False
        destacar_prio = False

        if col_sp is not None:
            destacar_sp = str(
                linha_excel[col_sp].value
            ).startswith("SIM")

        if col_prioridade is not None:
            destacar_prio = (
                str(
                    linha_excel[col_prioridade].value
                ).upper() == "ALTA"
            )

        for celula in linha_excel:
            if destacar_sp:
                celula.fill = cor_super

            if destacar_prio:
                celula.font = fonte_prioridade

    for coluna in ws.columns:
        letra = coluna[0].column_letter

        larguras = [
            len(str(c.value or ""))
            for c in list(coluna)[:120]
        ]

        ws.column_dimensions[letra].width = min(
            60,
            max(
                13,
                max(larguras, default=10) + 2
            )
        )


def excel_multiplas_abas(planilhas):
    memoria = io.BytesIO()

    with pd.ExcelWriter(
        memoria,
        engine="openpyxl"
    ) as writer:
        for nome, tabela in planilhas.items():
            nome_aba = nome[:31]
            limpa = limpar_para_excel(tabela)

            limpa.to_excel(
                writer,
                sheet_name=nome_aba,
                index=False
            )

            formatar_excel(
                writer,
                nome_aba,
                limpa
            )

    return memoria.getvalue()


def gerar_exportacoes(
    obras,
    auditoria,
    equipes,
    candidatos,
    programacao,
    superpontos,
    carga,
    trajetos
):
    resumo = pd.DataFrame([
        ["Saneamento", int(obras["LISTA"].eq(CAT_SAN).sum())],
        ["Levantamento", int(obras["LISTA"].eq(CAT_LEV).sum())],
        ["Duplicadas", int(obras["LISTA"].eq(CAT_DUP).sum())],
        ["Total de obras", len(obras)],
        [
            "Excluidas",
            int(
                auditoria["PENDENTE_CONTAGEM"]
                .eq("NÃO").sum()
            )
        ],
        ["Tarefas programadas", len(programacao)],
        ["Superpontos", int(
            superpontos["SUPER_PONTO"]
            .astype(str)
            .str.startswith("SIM")
            .sum()
        ) if not superpontos.empty else 0]
    ], columns=["INDICADOR", "VALOR"])

    abas = {
        "RESUMO OPERACIONAL": resumo,
        "CONSOLIDADO": obras,
        "OBRAS SANEAMENTO": obras[
            obras["LISTA"].eq(CAT_SAN)
        ],
        "OBRAS LEVANTAMENTO": obras[
            obras["LISTA"].eq(CAT_LEV)
        ],
        "OBRAS DUPLICADAS": obras[
            obras["LISTA"].eq(CAT_DUP)
        ],
        "AUDITORIA COMPLETA": auditoria,
        "PROGRAMACAO": programacao,
        "SUPERPONTOS": superpontos,
        "CARGA EQUIPES": carga,
        "EQUIPES": equipes,
        "CANDIDATOS": candidatos,
        "TRAJETOS": trajetos
    }

    excel_geral = excel_multiplas_abas(abas)

    zip_excel = io.BytesIO()
    zip_kml = io.BytesIO()

    data = datetime.now().strftime("%d.%m.%Y")

    with zipfile.ZipFile(
        zip_excel,
        "w",
        zipfile.ZIP_DEFLATED
    ) as z:
        z.writestr(
            f"Resumo_Operacional_{data}.xlsx",
            excel_multiplas_abas({
                "RESUMO OPERACIONAL": resumo
            })
        )

        z.writestr(
            f"Demanda_ListaContinua_Total_{data}.xlsx",
            excel_geral
        )

        if not programacao.empty:
            for equipe, grupo in programacao.groupby(
                "EQUIPE_PROGRAMADA"
            ):
                if equipe == "NAO ALOCADA":
                    continue

                nome = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    norm(equipe)
                )[:60]

                z.writestr(
                    f"Rotas_{data}/Rota_{nome}.xlsx",
                    excel_multiplas_abas({
                        "Obras Roteirizadas": grupo
                    })
                )

    with zipfile.ZipFile(
        zip_kml,
        "w",
        zipfile.ZIP_DEFLATED
    ) as z:
        z.writestr(
            f"ROTA_TOTAL_{data}.kml",
            gerar_kml(
                obras,
                superpontos,
                trajetos
            )
        )

        if not programacao.empty:
            for equipe, grupo in programacao.groupby(
                "EQUIPE_PROGRAMADA"
            ):
                if equipe == "NAO ALOCADA":
                    continue

                notas = set(
                    grupo["NOTA"].astype(str)
                )

                obras_eq = obras[
                    obras["NOTA"]
                    .astype(str)
                    .isin(notas)
                ]

                sp_eq = (
                    superpontos[
                        superpontos[
                            "EQUIPE_PROGRAMADA"
                        ].eq(equipe)
                    ]
                    if not superpontos.empty
                    else pd.DataFrame()
                )

                tr_eq = (
                    trajetos[
                        trajetos["EQUIPE"].eq(equipe)
                    ]
                    if not trajetos.empty
                    else pd.DataFrame()
                )

                nome = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    norm(equipe)
                )[:60]

                z.writestr(
                    f"KML_{data}/Rota_{nome}.kml",
                    gerar_kml(
                        obras_eq,
                        sp_eq,
                        tr_eq,
                        nome=f"Rota - {equipe}"
                    )
                )

    return (
        excel_geral,
        zip_excel.getvalue(),
        zip_kml.getvalue()
    )


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

st.markdown(
    '<div class="nip-title">📍 NIP | Analise Cruzada e Roteirizacao</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="nip-subtitle">'
    'Planejamento de Saneamento e Levantamento | '
    'Superpontos | Lista Continua | Excel e KML'
    '</div>',
    unsafe_allow_html=True
)

with st.expander(
    "📘 Regras e funcionamento",
    expanded=False
):
    st.markdown("""
    **STATUS SAP**
    - FINL e CANC são excluídos.
    - Todos os demais status são aceitos.

    **STATUS LIST**
    - `0`, `Em levantamento` e
      `Correção de levantamento` são aceitos.

    **Cruzamento**
    - NOTA de Saneamento com PROTOCOLO de Levantamento.
    - Notas das duas bases contam uma vez.

    **Equipes**
    - Somente LEVANTADORES e SANEAMENTO.

    **Superpontos**
    - DBSCAN Haversine em metros.
    - Agrupados por equipe, data e atividade.
    - As notas originais ficam preservadas.

    **Rotas**
    - Planejamento inicial por distância geográfica.
    - OSRM opcional para traçar ruas no KML.
    - Trajetos sem OSRM são identificados como
      estimativas em linha reta.
    """)

titulo_secao("📂 Importacao das bases")

c1, c2, c3 = st.columns(3)

with c1:
    arquivo_san = st.file_uploader(
        "BASE_SANEAMENTO",
        type=["xlsx", "csv"],
        key="arquivo_san"
    )

with c2:
    arquivo_lev = st.file_uploader(
        "BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"],
        key="arquivo_lev"
    )

with c3:
    arquivo_equipes = st.file_uploader(
        "LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"],
        key="arquivo_eq"
    )

if not all([
    arquivo_san,
    arquivo_lev,
    arquivo_equipes
]):
    st.info(
        "Envie os tres arquivos para habilitar "
        "o processamento."
    )
    st.stop()

assinatura = tuple(
    hashlib.sha256(
        f.getvalue()
    ).hexdigest()
    for f in [
        arquivo_san,
        arquivo_lev,
        arquivo_equipes
    ]
)

if st.button(
    "🚀 PROCESSAR AS TRES BASES",
    type="primary",
    use_container_width=True
):
    try:
        with st.spinner(
            "Validando colunas e cruzando notas..."
        ):
            san, aba_san = ler_saneamento(
                arquivo_san
            )

            lev, aba_lev = ler_levantamento(
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
                    "Nenhuma nota valida localizada."
                )

            st.session_state["nip_base"] = {
                "assinatura": assinatura,
                "auditoria": auditoria,
                "equipes": equipes,
                "aba_san": aba_san,
                "aba_lev": aba_lev,
                "qtd_san": len(san),
                "qtd_lev": len(lev)
            }

            st.session_state.pop(
                "nip_osrm", None
            )

        st.success(
            "Bases processadas com sucesso!"
        )

    except Exception as erro:
        st.error(
            f"Falha no processamento: {erro}"
        )

if "nip_base" not in st.session_state:
    st.stop()

dados = st.session_state["nip_base"]

if dados["assinatura"] != assinatura:
    st.warning(
        "Os arquivos foram alterados. "
        "Processe novamente."
    )
    st.stop()

auditoria = dados["auditoria"]
equipes = dados["equipes"]

st.caption(
    f"Abas reconhecidas: "
    f"Saneamento = {dados['aba_san']} | "
    f"Levantamento = {dados['aba_lev']}"
)

# ============================================================
# PARAMETROS
# ============================================================

with st.sidebar:
    st.header("⚙️ Planejamento")

    raio_equipes = st.select_slider(
        "Distancia maxima de referencia (km)",
        options=list(range(200, 501, 25)),
        value=200
    )

    quantidade_equipes = st.slider(
        "Equipes proximas para cada atividade",
        1, 5, 3
    )

    raio_super = st.slider(
        "Raio Superponto (metros)",
        min_value=10,
        max_value=500,
        value=50,
        step=10
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

    st.subheader("🛣️ Configuracao OSRM")

    usar_osrm_kml = st.checkbox(
        "Traçado de Ruas Real no KML (OSRM)",
        value=False,
        help=(
            "Desativado: exporta rapidamente os pontos. "
            "Ativado: consulta o traçado das estradas "
            "em pequenos lotes."
        )
    )

    servidor_osrm = st.text_input(
        "Endpoint OSRM",
        value=OSRM_DEFAULT
    )

    tamanho_lote_osrm = st.slider(
        "Trechos por consulta",
        1, 8, 5
    )

    st.caption(
        "Cada clique consulta poucos trechos, "
        "com timeout curto e cache. O servidor "
        "publico nao garante disponibilidade."
    )

# ============================================================
# INDICADORES
# ============================================================

obras = auditoria[
    auditoria["PENDENTE_CONTAGEM"].eq("SIM")
].copy()

obras = priorizar(
    obras, dias_prioridade
)

titulo_secao("📊 Indicadores Operacionais")

mostrar_cards([
    (
        "Saneamento",
        int(obras["LISTA"].eq(CAT_SAN).sum()),
        "#2563eb",
        "Exclusivas"
    ),
    (
        "Levantamento",
        int(obras["LISTA"].eq(CAT_LEV).sum()),
        "#16a34a",
        "Exclusivas"
    ),
    (
        "Duplicadas",
        int(obras["LISTA"].eq(CAT_DUP).sum()),
        "#9333ea",
        "Presentes nas duas bases"
    ),
    (
        "Total Pendente",
        len(obras),
        "#0D256C",
        "Notas unicas"
    ),
    (
        "Excluidas",
        int(
            auditoria["PENDENTE_CONTAGEM"]
            .eq("NÃO").sum()
        ),
        "#dc2626",
        "Fora da contagem"
    )
])

# ============================================================
# FILTROS OPERACIONAIS
# ============================================================

titulo_secao("🔎 Filtros Operacionais")

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
        "Regional",
        regionais
    )

with c2:
    filtro_municipio = st.multiselect(
        "Municipio",
        municipios
    )

with c3:
    filtro_prioridade = st.multiselect(
        "Prioridade",
        ["ALTA", "MEDIA", "BAIXA"]
    )

pesquisar = st.text_input(
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

if pesquisar.strip():
    view = view[
        view["NOTA"]
        .astype(str)
        .str.contains(
            re.escape(pesquisar.strip()),
            case=False,
            na=False
        )
    ]

mostrar_cards([
    (
        "Obras Selecionadas",
        len(view),
        "#0D256C",
        "Apos os filtros"
    ),
    (
        "Prioridade Alta",
        int(
            view["PRIORIDADE_PLANEJAMENTO"]
            .eq("ALTA").sum()
        ),
        "#dc2626",
        "Atendimento prioritario"
    ),
    (
        "Sem Coordenadas",
        int(
            view["LATITUDE"].isna().sum()
        ),
        "#ea580c",
        "Necessitam revisao"
    )
])

# ============================================================
# EQUIPES CANDIDATAS
# ============================================================

with st.spinner(
    "Identificando equipes mais proximas..."
):
    view, candidatos = gerar_candidatos(
        view,
        equipes,
        quantidade_equipes
    )

tarefas = gerar_tarefas(view)

programacao, carga = programar_tarefas(
    tarefas,
    candidatos,
    raio_equipes,
    data_inicio,
    modo,
    int(periodos),
    int(limite_periodo),
    int(limite_dia)
)

# Enriquecer lista original com a equipe sugerida.
if not programacao.empty:
    programacao_por_nota = programacao.groupby(
        "NOTA", sort=False
    )

    equipes_programadas = {}

    for nota, grupo in programacao_por_nota:
        equipes_programadas[nota] = " | ".join(
            f"{r['ATIVIDADE']}: "
            f"{r['EQUIPE_PROGRAMADA']}"
            for _, r in grupo.iterrows()
        )

    view["PROGRAMACAO_RESUMO"] = (
        view["NOTA"].map(equipes_programadas)
    )

# ============================================================
# SUPERPONTOS
# ============================================================

if not programacao.empty:
    tarefas_alocadas = programacao[
        programacao["STATUS_PROGRAMACAO"]
        .eq("SUGESTAO")
    ].copy()

    superpontos, quantidade_fundida = fundir_super_pontos(
        tarefas_alocadas,
        raio_metros=raio_super,
        agrupar_por_equipe=True
    )

else:
    superpontos = pd.DataFrame()
    quantidade_fundida = 0

qtd_superpontos = (
    int(
        superpontos["SUPER_PONTO"]
        .astype(str)
        .str.startswith("SIM")
        .sum()
    )
    if not superpontos.empty else 0
)

# ============================================================
# PREPARAR OSRM
# ============================================================

trechos = preparar_trechos(
    programacao,
    equipes
)

assinatura_rota = hashlib.sha256(
    (
        "|".join(
            trechos["NOTA"].astype(str)
            + ":"
            + trechos["ATIVIDADE"].astype(str)
            + ":"
            + trechos["DATA"].astype(str)
        )
        if not trechos.empty else ""
    ).encode()
    + str(servidor_osrm).encode()
).hexdigest()

if (
    "nip_osrm" not in st.session_state
    or st.session_state["nip_osrm"].get(
        "assinatura"
    ) != assinatura_rota
):
    st.session_state["nip_osrm"] = {
        "assinatura": assinatura_rota,
        "cursor": 0,
        "registros": []
    }

estado_osrm = st.session_state["nip_osrm"]

if usar_osrm_kml and not trechos.empty:
    titulo_secao("🛣️ Processamento OSRM")

    total_trechos = len(trechos)
    concluidos = estado_osrm["cursor"]

    st.progress(
        min(1.0, concluidos / total_trechos)
    )

    st.caption(
        f"Trechos processados: "
        f"{concluidos}/{total_trechos}"
    )

    if concluidos < total_trechos:
        if st.button(
            "🛣️ PROCESSAR PROXIMO LOTE OSRM",
            type="primary"
        ):
            novos, proximo_cursor = (
                processar_lote_osrm(
                    trechos,
                    servidor_osrm,
                    concluidos,
                    tamanho_lote_osrm
                )
            )

            estado_osrm["registros"].extend(
                novos
            )

            estado_osrm["cursor"] = (
                proximo_cursor
            )

            st.session_state["nip_osrm"] = estado_osrm

            st.rerun()

    else:
        st.success(
            "Todos os trechos foram consultados."
        )

    st.caption(
        "Os segmentos sem resposta do OSRM "
        "sao identificados separadamente."
    )

trajetos = pd.DataFrame(
    estado_osrm["registros"]
)

# ============================================================
# ABAS DE RESULTADOS
# ============================================================

abas = st.tabs([
    "🟣 SANEAMENTO",
    "🟢 LEVANTAMENTO",
    "🔵 DUPLICADAS",
    "📅 PROGRAMAÇÃO",
    "🏢 SUPERPONTOS",
    "🛣️ ROTAS",
    "⚠️ AUDITORIA",
    "🗺️ MAPA"
])

for aba, categoria in zip(
    abas[:3], CATEGORIAS
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

# ------------------------------------------------------------
# PROGRAMACAO
# ------------------------------------------------------------

with abas[3]:
    st.subheader(
        f"Programacao - {modo}"
    )

    if programacao.empty:
        st.info(
            "Nenhuma tarefa disponivel."
        )

    else:
        alocadas = int(
            programacao["STATUS_PROGRAMACAO"]
            .eq("SUGESTAO").sum()
        )

        nao_alocadas = int(
            programacao["STATUS_PROGRAMACAO"]
            .eq("SEM ALOCACAO").sum()
        )

        mostrar_cards([
            (
                "Tarefas",
                len(programacao),
                "#0D256C",
                "Total"
            ),
            (
                "Alocadas",
                alocadas,
                "#16a34a",
                "Programacao sugerida"
            ),
            (
                "Sem Alocacao",
                nao_alocadas,
                "#dc2626",
                "Requerem avaliacao"
            )
        ])

        st.dataframe(
            programacao,
            hide_index=True,
            use_container_width=True
        )

        st.subheader(
            "Carga por equipe"
        )

        st.dataframe(
            carga,
            hide_index=True,
            use_container_width=True
        )

# ------------------------------------------------------------
# SUPERPONTOS
# ------------------------------------------------------------

with abas[4]:
    st.subheader(
        "🏢 Superpontos - Lista Continua"
    )

    mostrar_cards([
        (
            "Superpontos",
            qtd_superpontos,
            "#f59e0b",
            "Grupos com mais de uma tarefa"
        ),
        (
            "Tarefas Fundidas",
            quantidade_fundida,
            "#9333ea",
            "Reducao de marcadores"
        ),
        (
            "Raio de Agrupamento",
            f"{raio_super} m",
            "#0D256C",
            "DBSCAN Haversine"
        )
    ])

    if not superpontos.empty:
        st.dataframe(
            limpar_para_excel(superpontos),
            hide_index=True,
            use_container_width=True
        )

# ------------------------------------------------------------
# ROTAS
# ------------------------------------------------------------

with abas[5]:
    st.subheader(
        "🛣️ Rotas e distancias"
    )

    st.caption(
        "Distancias geometricas sao referencia "
        "inicial. Somente OSRM identifica trajetos "
        "rodoviarios calculados."
    )

    st.dataframe(
        candidatos,
        hide_index=True,
        use_container_width=True
    )

    if not trajetos.empty:
        st.subheader(
            "Trechos OSRM"
        )

        st.dataframe(
            limpar_para_excel(trajetos),
            hide_index=True,
            use_container_width=True
        )

# ------------------------------------------------------------
# AUDITORIA
# ------------------------------------------------------------

with abas[6]:
    excluidas = auditoria[
        auditoria["PENDENTE_CONTAGEM"]
        .eq("NÃO")
    ]

    st.subheader(
        f"Notas excluidas - {len(excluidas)}"
    )

    st.dataframe(
        excluidas,
        hide_index=True,
        use_container_width=True
    )

    motivos = []

    for _, r in excluidas.iterrows():
        for motivo in str(
            r["MOTIVO_EXCLUSAO"]
        ).split(" | "):
            motivos.append({
                "NOTA": r["NOTA"],
                "MOTIVO": motivo
            })

    if motivos:
        resumo_motivos = (
            pd.DataFrame(motivos)["MOTIVO"]
            .value_counts()
        )

        st.bar_chart(resumo_motivos)

# ------------------------------------------------------------
# MAPA
# ------------------------------------------------------------

with abas[7]:
    st.subheader(
        "🗺️ Mapa das Obras"
    )

    if st.checkbox(
        "Carregar mapa interativo"
    ):
        geos = view.dropna(
            subset=["LATITUDE", "LONGITUDE"]
        )

        if geos.empty:
            st.warning(
                "Nenhuma coordenada valida."
            )

        else:
            centro = [
                float(geos["LATITUDE"].mean()),
                float(geos["LONGITUDE"].mean())
            ]

            mapa = folium.Map(
                location=centro,
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

                for _, r in subset.iterrows():
                    folium.Marker(
                        [
                            float(r["LATITUDE"]),
                            float(r["LONGITUDE"])
                        ],
                        popup=folium.Popup(
                            popup_html(r),
                            max_width=420
                        ),
                        icon=folium.Icon(
                            color=CORES_FOLIUM[categoria]
                        )
                    ).add_to(cluster)

                camada.add_to(mapa)

            if not trajetos.empty:
                camada_rotas = folium.FeatureGroup(
                    name="Trajetos OSRM"
                )

                for _, r in trajetos.iterrows():
                    geom = r.get("GEOMETRIA")

                    if not isinstance(geom, list):
                        continue

                    cor = (
                        "blue"
                        if r.get("TIPO_TRAJETO") == "OSRM"
                        else "gray"
                    )

                    folium.PolyLine(
                        [
                            [pt[1], pt[0]]
                            for pt in geom
                        ],
                        color=cor,
                        weight=3
                    ).add_to(camada_rotas)

                camada_rotas.add_to(mapa)

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

st.divider()

titulo_secao("📥 Exportacao - Padrao Lista Continua")

st.caption(
    "Planilhas no padrao NIP e mapas KML gerais "
    "ou individuais por equipe."
)

# Se a opcao estiver desligada, o KML nao inclui
# trajetos rodoviarios.
if usar_osrm_kml:
    trajetos_exportar = trajetos

    if not trechos.empty and len(trajetos) < len(trechos):
        st.warning(
            "O processamento OSRM ainda nao terminou. "
            "O KML exportado incluira apenas os "
            "trajetos ja processados."
        )

else:
    trajetos_exportar = pd.DataFrame()

excel_geral, zip_excel, zip_kml = gerar_exportacoes(
    view,
    auditoria,
    equipes,
    candidatos,
    programacao,
    superpontos,
    carga,
    trajetos_exportar
)

kml_geral = gerar_kml(
    view,
    superpontos,
    trajetos_exportar
)

c1, c2 = st.columns(2)

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

    st.download_button(
        "📦 PLANILHAS POR EQUIPE (ZIP)",
        data=zip_excel,
        file_name="NIP_Planilhas_ListaContinua.zip",
        mime="application/zip",
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

    st.download_button(
        "📦 KML POR EQUIPE (ZIP)",
        data=zip_kml,
        file_name="NIP_KML_ListaContinua.zip",
        mime="application/zip",
        use_container_width=True
    )

st.caption(
    f"Base Saneamento: {dados['qtd_san']} linhas | "
    f"Base Levantamento: {dados['qtd_lev']} linhas | "
    f"Notas analisadas: {len(auditoria)}"
)
