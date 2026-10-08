
import io
import re
import hashlib
import html
import unicodedata
import xml.etree.ElementTree as ET

from collections import defaultdict
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

from sklearn.neighbors import BallTree
from openpyxl.styles import Font, PatternFill


# ============================================================
# CONFIGURACAO
# ============================================================

st.set_page_config(
    page_title="NIP | Planejamento de Obras",
    page_icon="📍",
    layout="wide"
)

CONTRATOS_VALIDOS = {
    "0",
    "NIP GLOBAL LTDA - EQTL MARANHAO"
}

STATUS_SAP_BLOQUEADOS = {"FINL", "CANC"}

STATUS_LIST_ESPECIAIS = {
    "EM LEVANTAMENTO",
    "CORRECAO DE LEVANTAMENTO"
}

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_AMBAS = "OBRA SANEAMENTO E LEVANTAMENTO"

CATEGORIAS = (CAT_SAN, CAT_LEV, CAT_AMBAS)

CORES = {
    CAT_SAN: "purple",
    CAT_LEV: "green",
    CAT_AMBAS: "blue"
}

RAIO_TERRA_KM = 6371.0088


# ============================================================
# NORMALIZACAO E VALIDACOES
# ============================================================

def normalizar(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    texto = unicodedata.normalize(
        "NFKD", texto
    ).encode(
        "ascii", "ignore"
    ).decode("ascii")

    texto = texto.upper()
    return re.sub(r"\s+", " ", texto)


def normalizar_status(valor):
    texto = normalizar(valor)

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".")[0]

    return texto


def normalizar_nota(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]

    if normalizar(texto) in {
        "", "NAN", "NONE", "NULL", "0"
    }:
        return ""

    return texto


def encontrar_coluna(df, alternativas, obrigatoria=True):
    mapa = {
        normalizar(coluna): coluna
        for coluna in df.columns
    }

    for alternativa in alternativas:
        chave = normalizar(alternativa)

        if chave in mapa:
            return mapa[chave]

    if obrigatoria:
        raise ValueError(
            "Coluna nao encontrada: "
            + " / ".join(alternativas)
        )

    return None


def obter_valor(linha, coluna, padrao=""):
    if coluna is None:
        return padrao

    valor = linha.get(coluna, padrao)

    if pd.isna(valor):
        return padrao

    return valor


def ler_planilha(arquivo, aba=None):
    dados = io.BytesIO(arquivo.getvalue())

    if arquivo.name.lower().endswith(".csv"):
        for codificacao in ["utf-8-sig", "latin-1"]:
            try:
                return pd.read_csv(
                    io.BytesIO(arquivo.getvalue()),
                    sep=None,
                    engine="python",
                    encoding=codificacao,
                    dtype=str
                )
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                continue

        raise ValueError(
            "Nao foi possivel interpretar o CSV."
        )

    return pd.read_excel(
        dados,
        sheet_name=aba if aba else 0,
        dtype=str,
        engine="openpyxl"
    )


def converter_coordenadas(df, coluna_lat, coluna_lon):
    def converter(serie):
        return pd.to_numeric(
            serie.astype(str).str.replace(
                ",", ".", regex=False
            ),
            errors="coerce"
        )

    latitude = converter(df[coluna_lat])
    longitude = converter(df[coluna_lon])

    validas = (
        latitude.between(-35, 6)
        & longitude.between(-75, -30)
        & latitude.ne(0)
        & longitude.ne(0)
    )

    return (
        latitude.where(validas),
        longitude.where(validas)
    )


# ============================================================
# BASE DE SANEAMENTO
# ============================================================

def carregar_saneamento(arquivo):
    df = ler_planilha(arquivo)

    nota = encontrar_coluna(df, ["NOTA"])

    cidade = encontrar_coluna(
        df,
        ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
    )

    lat = encontrar_coluna(
        df,
        [
            "LATITUDE PROJETO",
            "LATITUDE",
            "LATITUDE CAMPO"
        ]
    )

    lon = encontrar_coluna(
        df,
        [
            "LONGITUDE PROJETO",
            "LONGITUDE",
            "LONGITUDE CAMPO"
        ]
    )

    regional = encontrar_coluna(
        df, ["REGIONAL"], False
    )

    df["CHAVE_NOTA"] = df[nota].map(normalizar_nota)

    df["MUNICIPIO_OBRA"] = (
        df[cidade].fillna("").astype(str).str.strip()
    )

    df["LAT_OBRA"], df["LON_OBRA"] = (
        converter_coordenadas(df, lat, lon)
    )

    df["REGIONAL_OBRA"] = (
        df[regional].fillna("")
        if regional is not None
        else ""
    )

    return df


# ============================================================
# BASE DE LEVANTAMENTO
# ============================================================

def identificar_exclusoes(linha):
    motivos = []

    sap = linha["SAP_NORM"]
    lista = linha["LIST_NORM"]
    contrato = linha["CONTRATO_NORM"]
    m = linha["COLUNA_M_NORM"]
    n = linha["COLUNA_N_NORM"]

    if sap in STATUS_SAP_BLOQUEADOS:
        motivos.append(f"SAP {sap}")

    if lista in STATUS_LIST_ESPECIAIS:
        motivos.append(
            f"LIST {lista} - sinalizada"
        )
    elif lista != "0":
        motivos.append(
            f"LIST diferente de 0: {lista or 'VAZIO'}"
        )

    if contrato not in CONTRATOS_VALIDOS:
        motivos.append(
            f"Contrato nao permitido: "
            f"{contrato or 'VAZIO'}"
        )

    if m != "0":
        motivos.append(
            f"Coluna M diferente de 0: {m or 'VAZIO'}"
        )

    if n != "0":
        motivos.append(
            f"Coluna N diferente de 0: {n or 'VAZIO'}"
        )

    return " | ".join(motivos)


def carregar_levantamento(arquivo):
    df = ler_planilha(arquivo, "NOTAS")

    protocolo = encontrar_coluna(
        df, ["PROTOCOLO"]
    )

    sap = encontrar_coluna(
        df, ["STATUS SAP", "STATUS_SAP"]
    )

    lista = encontrar_coluna(
        df, ["STATUS LIST", "STATUS_LIST"]
    )

    contrato = encontrar_coluna(
        df, ["CONTRATO"]
    )

    cidade = encontrar_coluna(
        df, ["MUNICIPIO", "MUNICÍPIO"]
    )

    lat = encontrar_coluna(df, ["LATITUDE"])
    lon = encontrar_coluna(df, ["LONGITUDE"])

    regional = encontrar_coluna(
        df, ["REGIONAL"], False
    )

    prioridade = encontrar_coluna(
        df, ["PRIORIDADE"], False
    )

    abertura = encontrar_coluna(
        df,
        [
            "DATA ABERTURA",
            "DATA DE ABERTURA"
        ],
        False
    )

    if len(df.columns) < 14:
        raise ValueError(
            "A aba NOTAS precisa possuir "
            "as colunas M e N."
        )

    coluna_m = df.columns[12]
    coluna_n = df.columns[13]

    if not (
        "ORCAMENTO MODULAR" in normalizar(coluna_m)
        and "PLA ALVOS" in normalizar(coluna_n)
    ):
        raise ValueError(
            "As colunas M e N nao correspondem "
            "a ORCAMENTO MODULAR e PLA ALVOS. "
            "Confira o layout da base."
        )

    df["CHAVE_NOTA"] = df[protocolo].map(
        normalizar_nota
    )

    df["MUNICIPIO_OBRA"] = (
        df[cidade].fillna("").astype(str).str.strip()
    )

    df["LAT_OBRA"], df["LON_OBRA"] = (
        converter_coordenadas(df, lat, lon)
    )

    df["SAP_NORM"] = df[sap].map(
        normalizar_status
    )

    df["LIST_NORM"] = df[lista].map(
        normalizar_status
    )

    df["CONTRATO_NORM"] = df[contrato].map(
        normalizar_status
    )

    df["COLUNA_M_NORM"] = df[coluna_m].map(
        normalizar_status
    )

    df["COLUNA_N_NORM"] = df[coluna_n].map(
        normalizar_status
    )

    df["REGIONAL_OBRA"] = (
        df[regional].fillna("")
        if regional is not None
        else ""
    )

    df["PRIORIDADE_ORIGINAL"] = (
        df[prioridade].fillna("")
        if prioridade is not None
        else ""
    )

    df["DATA_ABERTURA_ORIGINAL"] = (
        df[abertura].fillna("")
        if abertura is not None
        else ""
    )

    df["MOTIVOS_EXCLUSAO"] = df.apply(
        identificar_exclusoes,
        axis=1
    )

    return df


# ============================================================
# LOCALIDADES - SOMENTE DUAS ABAS
# ============================================================

def carregar_equipes(arquivo):
    excel = pd.ExcelFile(
        io.BytesIO(arquivo.getvalue()),
        engine="openpyxl"
    )

    partes = []

    configuracao = [
        ("LEVANTAMENTO", "LEVANTADORES"),
        ("SANEAMENTO", "SANEAMENTO")
    ]

    for tipo, aba_esperada in configuracao:
        nome_aba = next(
            (
                nome
                for nome in excel.sheet_names
                if normalizar(nome) == aba_esperada
            ),
            None
        )

        if nome_aba is None:
            raise ValueError(
                f"Aba obrigatoria ausente: {aba_esperada}"
            )

        df = pd.read_excel(
            excel,
            sheet_name=nome_aba,
            dtype=str
        )

        nome = encontrar_coluna(
            df,
            ["NOME", "NOME_COLAB", "EQUIPE"]
        )

        cidade = encontrar_coluna(
            df,
            ["CIDADES", "CIDADE", "MUNICIPIO"]
        )

        lat = encontrar_coluna(
            df, ["LATITUDE", "LAT"]
        )

        lon = encontrar_coluna(
            df, ["LONGITUDE", "LON"]
        )

        latitude, longitude = converter_coordenadas(
            df, lat, lon
        )

        partes.append(
            pd.DataFrame({
                "EQUIPE": df[nome],
                "CIDADE_BASE": df[cidade],
                "LAT_EQUIPE": latitude,
                "LON_EQUIPE": longitude,
                "TIPO_EQUIPE": tipo
            })
        )

    equipes = pd.concat(
        partes, ignore_index=True
    )

    equipes = equipes.dropna(
        subset=[
            "EQUIPE",
            "LAT_EQUIPE",
            "LON_EQUIPE"
        ]
    )

    equipes = equipes[
        equipes["EQUIPE"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    if equipes.empty:
        raise ValueError(
            "Nenhuma equipe com coordenadas "
            "validas foi encontrada."
        )

    equipes["ID_EQUIPE"] = (
        equipes["TIPO_EQUIPE"].astype(str)
        + "|"
        + equipes["EQUIPE"].astype(str)
        + "|"
        + equipes["CIDADE_BASE"].astype(str)
    )

    return equipes.reset_index(drop=True)


# ============================================================
# CRUZAMENTO DAS BASES
# ============================================================

def primeira_ocorrencia(grupo):
    validas = (
        grupo["LAT_OBRA"].notna()
        & grupo["LON_OBRA"].notna()
    )

    if validas.any():
        return grupo.loc[validas].iloc[0]

    return grupo.iloc[0]


def consolidar_bases(saneamento, levantamento):
    san = saneamento[
        saneamento["CHAVE_NOTA"].ne("")
    ].copy()

    lev = levantamento[
        levantamento["CHAVE_NOTA"].ne("")
    ].copy()

    san_grupos = dict(
        tuple(san.groupby("CHAVE_NOTA", sort=False))
    )

    lev_grupos = dict(
        tuple(lev.groupby("CHAVE_NOTA", sort=False))
    )

    todas_notas = dict.fromkeys(
        list(san_grupos) + list(lev_grupos)
    )

    registros = []

    for nota in todas_notas:
        grupo_san = san_grupos.get(nota)
        grupo_lev = lev_grupos.get(nota)

        tem_san = grupo_san is not None
        tem_lev = grupo_lev is not None

        a = (
            primeira_ocorrencia(grupo_san)
            if tem_san else None
        )

        b = (
            primeira_ocorrencia(grupo_lev)
            if tem_lev else None
        )

        if tem_san and tem_lev:
            categoria = CAT_AMBAS
        elif tem_san:
            categoria = CAT_SAN
        else:
            categoria = CAT_LEV

        motivos = []

        if tem_lev:
            for texto in grupo_lev["MOTIVOS_EXCLUSAO"]:
                if texto:
                    motivos.extend(texto.split(" | "))

        motivos = list(dict.fromkeys(motivos))

        bloqueada = bool(motivos)

        if b is not None and pd.notna(b["LAT_OBRA"]) \
                and pd.notna(b["LON_OBRA"]):
            origem = b
        elif a is not None:
            origem = a
        else:
            origem = b

        registro = {
            "NOTA": nota,
            "LISTA": categoria,
            "DUPLICADA": (
                "SIM" if tem_san and tem_lev else "NÃO"
            ),
            "PENDENTE_CONTAGEM": (
                "NÃO" if bloqueada else "SIM"
            ),
            "MOTIVO_EXCLUSAO": (
                " | ".join(motivos) if motivos else "-"
            ),
            "MUNICIPIO": origem.get(
                "MUNICIPIO_OBRA", ""
            ),
            "LATITUDE": origem.get("LAT_OBRA"),
            "LONGITUDE": origem.get("LON_OBRA"),
            "REGIONAL": origem.get(
                "REGIONAL_OBRA", ""
            ),
            "STATUS_SAP": (
                b["SAP_NORM"]
                if tem_lev else "SEM LEVANTAMENTO"
            ),
            "STATUS_LIST": (
                b["LIST_NORM"]
                if tem_lev else "SEM LEVANTAMENTO"
            ),
            "CONTRATO": (
                b["CONTRATO_NORM"]
                if tem_lev else "SEM LEVANTAMENTO"
            ),
            "COLUNA_M": (
                b["COLUNA_M_NORM"] if tem_lev else "-"
            ),
            "COLUNA_N": (
                b["COLUNA_N_NORM"] if tem_lev else "-"
            ),
            "PRIORIDADE_ORIGINAL": (
                b["PRIORIDADE_ORIGINAL"]
                if tem_lev else ""
            ),
            "DATA_ABERTURA": (
                b["DATA_ABERTURA_ORIGINAL"]
                if tem_lev else ""
            ),
            "OCORRENCIAS_SANEAMENTO": (
                len(grupo_san) if tem_san else 0
            ),
            "OCORRENCIAS_LEVANTAMENTO": (
                len(grupo_lev) if tem_lev else 0
            ),
            "MUNICIPIO_SANEAMENTO": (
                a["MUNICIPIO_OBRA"] if tem_san else ""
            ),
            "MUNICIPIO_LEVANTAMENTO": (
                b["MUNICIPIO_OBRA"] if tem_lev else ""
            )
        }

        registros.append(registro)

    return pd.DataFrame(registros)


# ============================================================
# DISTANCIAS
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    lat1 = np.radians(lat1)
    lon1 = np.radians(lon1)
    lat2 = np.radians(lat2)
    lon2 = np.radians(lon2)

    delta_lat = lat2 - lat1
    delta_lon = lon2 - lon1

    a = (
        np.sin(delta_lat / 2) ** 2
        + np.cos(lat1)
        * np.cos(lat2)
        * np.sin(delta_lon / 2) ** 2
    )

    return (
        2 * RAIO_TERRA_KM
        * np.arcsin(np.sqrt(np.clip(a, 0, 1)))
    )


def tipos_necessarios(categoria):
    if categoria == CAT_AMBAS:
        return ["SANEAMENTO", "LEVANTAMENTO"]

    if categoria == CAT_SAN:
        return ["SANEAMENTO"]

    return ["LEVANTAMENTO"]


def adicionar_equipes_proximas(
    obras, equipes, raio=200, quantidade=3
):
    saida = obras.copy()

    resultados = []

    for _, obra in saida.iterrows():
        registro = {}

        for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
            candidatos = equipes[
                equipes["TIPO_EQUIPE"].eq(tipo)
            ]

            col_equipe = f"EQUIPES_{tipo}"
            col_min = f"DISTANCIA_MINIMA_{tipo}_KM"
            col_sem = f"{tipo}_SEM_EQUIPE_NO_RAIO"

            if tipo not in tipos_necessarios(
                obra["LISTA"]
            ):
                registro[col_equipe] = "NAO APLICAVEL"
                registro[col_min] = np.nan
                registro[col_sem] = "NAO APLICAVEL"
                continue

            if (
                pd.isna(obra["LATITUDE"])
                or pd.isna(obra["LONGITUDE"])
                or candidatos.empty
            ):
                registro[col_equipe] = "SEM EQUIPE"
                registro[col_min] = np.nan
                registro[col_sem] = "SIM"
                continue

            distancias = haversine(
                float(obra["LATITUDE"]),
                float(obra["LONGITUDE"]),
                candidatos["LAT_EQUIPE"].to_numpy(float),
                candidatos["LON_EQUIPE"].to_numpy(float)
            )

            encontrados = candidatos.assign(
                DISTANCIA_KM=distancias
            ).sort_values("DISTANCIA_KM")

            registro[col_min] = round(
                float(encontrados["DISTANCIA_KM"].min()),
                2
            )

            encontrados = encontrados[
                encontrados["DISTANCIA_KM"].le(raio)
            ].head(quantidade)

            registro[col_sem] = (
                "SIM" if encontrados.empty else "NÃO"
            )

            if encontrados.empty:
                registro[col_equipe] = (
                    f"SEM EQUIPE EM {raio} KM"
                )
            else:
                registro[col_equipe] = " | ".join(
                    f"{r['EQUIPE']} "
                    f"({r['CIDADE_BASE']}) - "
                    f"{r['DISTANCIA_KM']:.1f} km"
                    for _, r in encontrados.iterrows()
                )

        resultados.append(registro)

    detalhes = pd.DataFrame(
        resultados, index=saida.index
    )

    return pd.concat([saida, detalhes], axis=1)


# ============================================================
# AGRUPAMENTO GEOGRAFICO
# ============================================================

def agrupar_obras(obras, raio_km=20):
    df = obras.copy().reset_index(drop=True)

    df["GRUPO_DESLOCAMENTO"] = ""
    df["QTD_OBRAS_GRUPO"] = 0

    validas = df.dropna(
        subset=["LATITUDE", "LONGITUDE"]
    )

    if validas.empty:
        return df

    pontos = np.radians(
        validas[
            ["LATITUDE", "LONGITUDE"]
        ].astype(float).to_numpy()
    )

    arvore = BallTree(
        pontos,
        metric="haversine"
    )

    vizinhos = arvore.query_radius(
        pontos,
        r=raio_km / RAIO_TERRA_KM
    )

    pais = list(range(len(validas)))

    def raiz(i):
        while pais[i] != i:
            pais[i] = pais[pais[i]]
            i = pais[i]
        return i

    def unir(i, j):
        ri = raiz(i)
        rj = raiz(j)

        if ri != rj:
            pais[rj] = ri

    for i, proximos in enumerate(vizinhos):
        for j in proximos:
            if i != int(j):
                unir(i, int(j))

    grupos = {}

    for posicao, indice in enumerate(validas.index):
        identificador = raiz(posicao)

        if identificador not in grupos:
            grupos[identificador] = (
                f"GRP-{len(grupos) + 1:04d}"
            )

        df.at[
            indice, "GRUPO_DESLOCAMENTO"
        ] = grupos[identificador]

    contagem = (
        df.loc[
            df["GRUPO_DESLOCAMENTO"].ne(""),
            "GRUPO_DESLOCAMENTO"
        ].value_counts()
    )

    df["QTD_OBRAS_GRUPO"] = (
        df["GRUPO_DESLOCAMENTO"]
        .map(contagem)
        .fillna(0)
        .astype(int)
    )

    return df


# ============================================================
# PRIORIDADES
# ============================================================

def adicionar_prioridades(obras, dias_media=14):
    df = obras.copy()

    datas = pd.to_datetime(
        df["DATA_ABERTURA"],
        errors="coerce",
        dayfirst=True
    )

    hoje = pd.Timestamp.today().normalize()

    df["DIAS_DESDE_ABERTURA"] = (
        hoje - datas
    ).dt.days

    def classificar(linha):
        original = normalizar(
            linha["PRIORIDADE_ORIGINAL"]
        )

        if any(
            termo in original
            for termo in [
                "URGENTE",
                "CRITICA",
                "ALTA"
            ]
        ):
            return "ALTA"

        idade = linha["DIAS_DESDE_ABERTURA"]

        if pd.notna(idade):
            if idade >= dias_media * 4:
                return "ALTA"

            if idade >= dias_media:
                return "MEDIA"

        return "BAIXA"

    df["PRIORIDADE_PLANEJAMENTO"] = df.apply(
        classificar,
        axis=1
    )

    return df


# ============================================================
# PROGRAMACAO EQUILIBRADA DAS EQUIPES
# ============================================================

def programar_equipes(
    obras,
    equipes,
    capacidade=8,
    dias=5,
    raio=200,
    candidatos_max=3
):
    tarefas = []

    for _, obra in obras.iterrows():
        for atividade in tipos_necessarios(
            obra["LISTA"]
        ):
            registro = obra.to_dict()
            registro["ATIVIDADE"] = atividade
            tarefas.append(registro)

    colunas_programacao = [
        "NOTA",
        "ATIVIDADE",
        "EQUIPE_PROGRAMADA",
        "DIA_PROGRAMADO",
        "DISTANCIA_PROGRAMADA_KM",
        "STATUS_PROGRAMACAO"
    ]

    if not tarefas:
        return (
            pd.DataFrame(columns=colunas_programacao),
            pd.DataFrame(
                columns=[
                    "ATIVIDADE",
                    "EQUIPE_PROGRAMADA",
                    "DIA_PROGRAMADO",
                    "TAREFAS"
                ]
            )
        )

    df = pd.DataFrame(tarefas)

    ordem = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    df["_PRIORIDADE"] = (
        df["PRIORIDADE_PLANEJAMENTO"]
        .map(ordem)
        .fillna(3)
    )

    df = df.sort_values(
        [
            "_PRIORIDADE",
            "GRUPO_DESLOCAMENTO",
            "NOTA"
        ]
    ).reset_index(drop=True)

    ocupacao = defaultdict(int)
    alocacoes = []

    for _, tarefa in df.iterrows():
        candidatos = equipes[
            equipes["TIPO_EQUIPE"].eq(
                tarefa["ATIVIDADE"]
            )
        ].copy()

        escolhido = "NÃO ALOCADA"
        dia_escolhido = ""
        distancia_escolhida = np.nan

        if (
            pd.notna(tarefa["LATITUDE"])
            and pd.notna(tarefa["LONGITUDE"])
            and not candidatos.empty
        ):
            distancias = haversine(
                float(tarefa["LATITUDE"]),
                float(tarefa["LONGITUDE"]),
                candidatos["LAT_EQUIPE"].to_numpy(float),
                candidatos["LON_EQUIPE"].to_numpy(float)
            )

            candidatos["DISTANCIA_KM"] = distancias

            candidatos = candidatos[
                candidatos["DISTANCIA_KM"].le(raio)
            ].sort_values(
                "DISTANCIA_KM"
            ).head(candidatos_max)

            alternativas = []

            for _, equipe in candidatos.iterrows():
                for dia in range(1, dias + 1):
                    chave = (
                        equipe["ID_EQUIPE"],
                        dia
                    )

                    carga = ocupacao[chave]

                    if carga < capacidade:
                        alternativas.append(
                            (
                                carga,
                                float(
                                    equipe["DISTANCIA_KM"]
                                ),
                                dia,
                                equipe["ID_EQUIPE"],
                                str(equipe["EQUIPE"])
                            )
                        )
                        break

            if alternativas:
                (
                    _,
                    distancia_escolhida,
                    dia_escolhido,
                    id_equipe,
                    escolhido
                ) = min(alternativas)

                ocupacao[
                    (id_equipe, dia_escolhido)
                ] += 1

        alocacoes.append({
            "EQUIPE_PROGRAMADA": escolhido,
            "DIA_PROGRAMADO": dia_escolhido,
            "DISTANCIA_PROGRAMADA_KM": (
                round(distancia_escolhida, 2)
                if pd.notna(distancia_escolhida)
                else np.nan
            ),
            "STATUS_PROGRAMACAO": (
                "SUGESTÃO"
                if escolhido != "NÃO ALOCADA"
                else "SEM ALOCACAO"
            )
        })

    df = pd.concat(
        [df, pd.DataFrame(alocacoes)],
        axis=1
    )

    df = df.drop(columns=["_PRIORIDADE"])

    alocadas = df[
        df["EQUIPE_PROGRAMADA"].ne(
            "NÃO ALOCADA"
        )
    ]

    carga = (
        alocadas.groupby(
            [
                "ATIVIDADE",
                "EQUIPE_PROGRAMADA",
                "DIA_PROGRAMADO"
            ],
            dropna=False
        )
        .size()
        .reset_index(name="TAREFAS")
    )

    return df, carga


# ============================================================
# COMPARACAO COM ANALISE ANTERIOR
# ============================================================

def comparar_historico(obras, arquivo_anterior):
    planilhas = pd.read_excel(
        io.BytesIO(arquivo_anterior.getvalue()),
        sheet_name=None,
        dtype=str,
        engine="openpyxl"
    )

    anterior = planilhas.get("AUDITORIA COMPLETA")

    if anterior is None:
        anterior = planilhas.get("CONSOLIDADO")

    if anterior is None:
        raise ValueError(
            "O arquivo anterior precisa conter "
            "AUDITORIA COMPLETA ou CONSOLIDADO."
        )

    if "NOTA" not in anterior.columns:
        raise ValueError(
            "A analise anterior nao possui "
            "a coluna NOTA."
        )

    if "PENDENTE_CONTAGEM" in anterior.columns:
        anterior = anterior[
            anterior["PENDENTE_CONTAGEM"].eq("SIM")
        ]

    antigas = set(
        anterior["NOTA"].map(normalizar_nota)
    ) - {""}

    atuais = set(
        obras["NOTA"].map(normalizar_nota)
    ) - {""}

    novas = pd.DataFrame({
        "NOTA": sorted(atuais - antigas)
    })

    removidas = pd.DataFrame({
        "NOTA": sorted(antigas - atuais)
    })

    novas["ALTERACAO"] = "NOVA"
    removidas["ALTERACAO"] = "SAIU DA PENDENCIA"

    return novas, removidas


# ============================================================
# EXPORTACAO KML - GOOGLE EARTH
# ============================================================

def exportar_kml(obras):
    namespace = (
        "http://www.opengis.net/kml/2.2"
    )

    ET.register_namespace("", namespace)

    def elemento(pai, nome, valor=None):
        item = ET.SubElement(
            pai,
            f"{{{namespace}}}{nome}"
        )

        if valor is not None:
            item.text = str(valor)

        return item

    raiz = ET.Element(
        f"{{{namespace}}}kml"
    )

    documento = elemento(
        raiz, "Document"
    )

    elemento(
        documento,
        "name",
        "NIP - Obras Pendentes"
    )

    for categoria in CATEGORIAS:
        pasta = elemento(
            documento, "Folder"
        )

        elemento(
            pasta, "name", categoria
        )

        dados = obras[
            obras["LISTA"].eq(categoria)
        ].dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        for _, obra in dados.iterrows():
            ponto = elemento(
                pasta, "Placemark"
            )

            elemento(
                ponto, "name", obra["NOTA"]
            )

            descricao = (
                f"Nota: {obra['NOTA']}\n"
                f"Categoria: {categoria}\n"
                f"Municipio: {obra['MUNICIPIO']}\n"
                f"Regional: {obra['REGIONAL']}\n"
                f"Prioridade: "
                f"{obra.get('PRIORIDADE_PLANEJAMENTO', '')}\n"
                f"Grupo: "
                f"{obra.get('GRUPO_DESLOCAMENTO', '')}\n"
                f"Equipes Saneamento: "
                f"{obra.get('EQUIPES_SANEAMENTO', '')}\n"
                f"Equipes Levantamento: "
                f"{obra.get('EQUIPES_LEVANTAMENTO', '')}"
            )

            elemento(
                ponto, "description", descricao
            )

            local = elemento(
                ponto, "Point"
            )

            elemento(
                local,
                "coordinates",
                (
                    f"{float(obra['LONGITUDE']):.8f},"
                    f"{float(obra['LATITUDE']):.8f},0"
                )
            )

    return ET.tostring(
        raiz,
        encoding="utf-8",
        xml_declaration=True
    )


# ============================================================
# EXPORTACAO EXCEL
# ============================================================

def exportar_excel(
    obras,
    auditoria,
    equipes,
    programacao,
    carga,
    novas,
    removidas,
    parametros
):
    excluidas = auditoria[
        auditoria["PENDENTE_CONTAGEM"].eq("NÃO")
    ]

    motivos = []

    for _, obra in excluidas.iterrows():
        for motivo in str(
            obra["MOTIVO_EXCLUSAO"]
        ).split(" | "):
            motivos.append({
                "NOTA": obra["NOTA"],
                "MOTIVO": motivo
            })

    resumo = [
        [
            "Obras Saneamento",
            int(obras["LISTA"].eq(CAT_SAN).sum())
        ],
        [
            "Obras Levantamento",
            int(obras["LISTA"].eq(CAT_LEV).sum())
        ],
        [
            "Obras Duplicadas",
            int(obras["LISTA"].eq(CAT_AMBAS).sum())
        ],
        [
            "Total unico pendente",
            len(obras)
        ],
        [
            "Total excluido",
            len(excluidas)
        ],
        [
            "Tarefas nao alocadas",
            int(
                programacao["EQUIPE_PROGRAMADA"]
                .eq("NÃO ALOCADA").sum()
            )
        ]
    ]

    for chave, valor in parametros.items():
        resumo.append([chave, valor])

    resumo.append([
        "Gerado em",
        datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        )
    ])

    planilhas = {
        "RESUMO": pd.DataFrame(
            resumo,
            columns=["INDICADOR", "VALOR"]
        ),
        "CONSOLIDADO": obras,
        "OBRAS SANEAMENTO": obras[
            obras["LISTA"].eq(CAT_SAN)
        ],
        "OBRAS LEVANTAMENTO": obras[
            obras["LISTA"].eq(CAT_LEV)
        ],
        "OBRAS DUPLICADAS": obras[
            obras["LISTA"].eq(CAT_AMBAS)
        ],
        "AUDITORIA COMPLETA": auditoria,
        "EXCLUIDAS": excluidas,
        "MOTIVOS EXCLUSAO": pd.DataFrame(
            motivos,
            columns=["NOTA", "MOTIVO"]
        ),
        "PROGRAMACAO": programacao,
        "CARGA EQUIPES": carga,
        "EQUIPES": equipes,
        "NOTAS NOVAS": novas,
        "NOTAS REMOVIDAS": removidas
    }

    memoria = io.BytesIO()

    with pd.ExcelWriter(
        memoria,
        engine="openpyxl"
    ) as writer:

        for nome, tabela in planilhas.items():
            tabela.to_excel(
                writer,
                sheet_name=nome[:31],
                index=False
            )

            ws = writer.sheets[nome[:31]]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False

            for celula in ws[1]:
                celula.fill = PatternFill(
                    "solid",
                    fgColor="14366F"
                )

                celula.font = Font(
                    color="FFFFFF",
                    bold=True
                )

            for coluna in ws.columns:
                letra = coluna[0].column_letter

                valores = [
                    len(str(c.value or ""))
                    for c in list(coluna)[:100]
                ]

                largura = min(
                    55,
                    max(
                        13,
                        max(valores, default=10) + 2
                    )
                )

                ws.column_dimensions[
                    letra
                ].width = largura

    return memoria.getvalue()


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.title(
    "📍 NIP | Planejamento e Auditoria de Obras"
)

st.caption(
    "Análise cruzada de Saneamento e Levantamento, "
    "com planejamento de equipes e exportação."
)

with st.expander(
    "📘 Regras de contagem e metodologia"
):
    st.markdown("""
    **Cruzamento:** NOTA da base Saneamento com
    PROTOCOLO da base Levantamento.

    **Exclusões:** SAP FINL/CANC, LIST diferente
    de zero, contratos não permitidos e
    colunas M/N diferentes de zero.

    **Em levantamento** e **Correção de levantamento**
    permanecem sinalizadas, mas fora da contagem.

    **Contrato válido:** 0 ou
    NIP GLOBAL LTDA - EQTL MARANHÃO.

    **Duplicidade:** uma nota presente nas duas
    bases aparece uma única vez na categoria
    Saneamento + Levantamento.

    **Equipes:** somente as abas LEVANTADORES
    e SANEAMENTO.

    **Distâncias:** calculadas em linha reta,
    não por estradas.

    **Agrupamento:** obras conectadas por
    proximidade podem formar grupos em cadeia.
    Portanto, dois extremos do mesmo grupo
    podem estar além do raio definido.

    **Programação:** sugestão de distribuição
    por capacidade, sem registrar execução real.
    """)


# ============================================================
# UPLOADS
# ============================================================

st.subheader("📂 Importar bases")

c1, c2, c3 = st.columns(3)

with c1:
    arquivo_san = st.file_uploader(
        "1. BASE_SANEAMENTO",
        type=["xlsx", "csv"],
        key="upload_san"
    )

with c2:
    arquivo_lev = st.file_uploader(
        "2. BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"],
        key="upload_lev"
    )

with c3:
    arquivo_equipes = st.file_uploader(
        "3. LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"],
        key="upload_equipes"
    )

if not all([
    arquivo_san,
    arquivo_lev,
    arquivo_equipes
]):
    st.warning(
        "Envie os três arquivos obrigatórios "
        "para iniciar o processamento."
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
    "🚀 Processar bases",
    type="primary",
    use_container_width=True
):
    try:
        with st.spinner(
            "Validando bases e cruzando notas..."
        ):
            saneamento = carregar_saneamento(
                arquivo_san
            )

            levantamento = carregar_levantamento(
                arquivo_lev
            )

            equipes = carregar_equipes(
                arquivo_equipes
            )

            auditoria = consolidar_bases(
                saneamento,
                levantamento
            )

            if auditoria.empty:
                raise ValueError(
                    "Nenhuma nota valida encontrada."
                )

            st.session_state["analise_nip"] = {
                "assinatura": assinatura,
                "auditoria": auditoria,
                "equipes": equipes,
                "linhas_san": len(saneamento),
                "linhas_lev": len(levantamento)
            }

        st.success(
            "Bases processadas com sucesso!"
        )

    except Exception as erro:
        st.error(
            f"Erro no processamento: {erro}"
        )

if "analise_nip" not in st.session_state:
    st.stop()

dados = st.session_state["analise_nip"]

if dados["assinatura"] != assinatura:
    st.warning(
        "Os arquivos foram alterados. "
        "Clique novamente em Processar bases."
    )
    st.stop()

auditoria = dados["auditoria"]
equipes = dados["equipes"]


# ============================================================
# PARAMETROS DO PLANEJAMENTO
# ============================================================

with st.sidebar:
    st.header("⚙️ Planejamento")

    raio = st.select_slider(
        "Raio maximo das equipes (km)",
        options=list(range(200, 501, 25)),
        value=200
    )

    quantidade_equipes = st.slider(
        "Equipes candidatas por atividade",
        1,
        5,
        3
    )

    raio_grupo = st.select_slider(
        "Agrupar obras proximas ate (km)",
        options=[5, 10, 20, 30, 50],
        value=20
    )

    capacidade = st.number_input(
        "Tarefas por equipe por dia",
        min_value=1,
        max_value=100,
        value=8
    )

    dias = st.number_input(
        "Dias de programacao",
        min_value=1,
        max_value=60,
        value=5
    )

    dias_prioridade = st.number_input(
        "Dias para prioridade MEDIA",
        min_value=1,
        max_value=365,
        value=14
    )

    st.divider()

    arquivo_anterior = st.file_uploader(
        "Analise anterior (opcional)",
        type=["xlsx"]
    )

    st.caption(
        "A base anterior é usada apenas "
        "para comparar entradas e saídas."
    )


# ============================================================
# PROCESSAMENTO OPERACIONAL
# ============================================================

obras = auditoria[
    auditoria["PENDENTE_CONTAGEM"].eq("SIM")
].copy()

obras = adicionar_prioridades(
    obras,
    dias_prioridade
)

obras = agrupar_obras(
    obras,
    raio_grupo
)

obras = adicionar_equipes_proximas(
    obras,
    equipes,
    raio,
    quantidade_equipes
)

programacao, carga = programar_equipes(
    obras,
    equipes,
    capacidade=capacidade,
    dias=dias,
    raio=raio,
    candidatos_max=quantidade_equipes
)

novas = pd.DataFrame(
    columns=["NOTA", "ALTERACAO"]
)

removidas = pd.DataFrame(
    columns=["NOTA", "ALTERACAO"]
)

if arquivo_anterior is not None:
    try:
        novas, removidas = comparar_historico(
            obras,
            arquivo_anterior
        )

    except Exception as erro:
        st.warning(
            f"Falha na comparação: {erro}"
        )


# ============================================================
# INDICADORES
# ============================================================

st.subheader("📊 Indicadores gerais")

metricas = st.columns(5)

indicadores = [
    (
        "Saneamento",
        int(obras["LISTA"].eq(CAT_SAN).sum())
    ),
    (
        "Levantamento",
        int(obras["LISTA"].eq(CAT_LEV).sum())
    ),
    (
        "Duplicadas",
        int(obras["LISTA"].eq(CAT_AMBAS).sum())
    ),
    (
        "Total pendente",
        len(obras)
    ),
    (
        "Excluidas",
        int(
            auditoria["PENDENTE_CONTAGEM"]
            .eq("NÃO")
            .sum()
        )
    )
]

for coluna, (titulo, valor) in zip(
    metricas,
    indicadores
):
    coluna.metric(titulo, valor)

st.divider()


# ============================================================
# FILTROS
# ============================================================

st.subheader("🔎 Filtros operacionais")

col1, col2, col3 = st.columns(3)

regionais = sorted(
    obras["REGIONAL"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

municipios = sorted(
    obras["MUNICIPIO"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

with col1:
    filtro_regional = st.multiselect(
        "Regional",
        regionais
    )

with col2:
    filtro_municipio = st.multiselect(
        "Municipio",
        municipios
    )

with col3:
    filtro_prioridade = st.multiselect(
        "Prioridade",
        ["ALTA", "MEDIA", "BAIXA"]
    )

busca_nota = st.text_input(
    "Pesquisar nota"
).strip()

visualizacao = obras.copy()

if filtro_regional:
    visualizacao = visualizacao[
        visualizacao["REGIONAL"].isin(
            filtro_regional
        )
    ]

if filtro_municipio:
    visualizacao = visualizacao[
        visualizacao["MUNICIPIO"].isin(
            filtro_municipio
        )
    ]

if filtro_prioridade:
    visualizacao = visualizacao[
        visualizacao[
            "PRIORIDADE_PLANEJAMENTO"
        ].isin(filtro_prioridade)
    ]

if busca_nota:
    visualizacao = visualizacao[
        visualizacao["NOTA"]
        .astype(str)
        .str.contains(
            re.escape(busca_nota),
            case=False,
            na=False
        )
    ]


# ============================================================
# ABAS
# ============================================================

abas = st.tabs([
    "🟣 SANEAMENTO",
    "🟢 LEVANTAMENTO",
    "🔵 DUPLICADAS",
    "📅 PROGRAMAÇÃO",
    "⚠️ AUDITORIA",
    "🗺️ MAPA",
    "📈 HISTÓRICO"
])


# ------------------------------------------------------------
# TRES LISTAS
# ------------------------------------------------------------

for aba, categoria in zip(
    abas[:3],
    CATEGORIAS
):
    with aba:
        tabela = visualizacao[
            visualizacao["LISTA"].eq(categoria)
        ]

        st.subheader(
            f"{categoria} - {len(tabela)} obras"
        )

        st.dataframe(
            tabela,
            use_container_width=True,
            hide_index=True
        )


# ------------------------------------------------------------
# PROGRAMACAO
# ------------------------------------------------------------

with abas[3]:
    st.subheader(
        "📅 Programacao sugerida das equipes"
    )

    st.caption(
        "Notas duplicadas geram duas tarefas: "
        "uma para Saneamento e outra para "
        "Levantamento."
    )

    programacao_filtrada = programacao[
        programacao["NOTA"].isin(
            visualizacao["NOTA"]
        )
    ]

    st.dataframe(
        programacao_filtrada,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Carga por equipe")

    st.dataframe(
        carga,
        use_container_width=True,
        hide_index=True
    )

    sem_alocacao = int(
        programacao_filtrada[
            "EQUIPE_PROGRAMADA"
        ].eq("NÃO ALOCADA").sum()
    )

    grupos = int(
        visualizacao[
            "GRUPO_DESLOCAMENTO"
        ].replace("", np.nan).nunique()
    )

    c1, c2 = st.columns(2)

    c1.metric(
        "Tarefas nao alocadas",
        sem_alocacao
    )

    c2.metric(
        "Grupos geograficos",
        grupos
    )


# ------------------------------------------------------------
# AUDITORIA
# ------------------------------------------------------------

with abas[4]:
    st.subheader(
        "⚠️ Auditoria das exclusoes"
    )

    excluidas = auditoria[
        auditoria[
            "PENDENTE_CONTAGEM"
        ].eq("NÃO")
    ]

    st.dataframe(
        excluidas,
        use_container_width=True,
        hide_index=True
    )

    lista_motivos = []

    for texto in excluidas[
        "MOTIVO_EXCLUSAO"
    ].astype(str):
        lista_motivos.extend(
            texto.split(" | ")
        )

    if lista_motivos:
        contagem_motivos = (
            pd.Series(lista_motivos)
            .value_counts()
        )

        st.bar_chart(contagem_motivos)

    st.caption(
        "Uma nota pode ter varios motivos de "
        "exclusao, mas conta apenas uma vez "
        "no total de obras excluidas."
    )


# ------------------------------------------------------------
# MAPA
# ------------------------------------------------------------

with abas[5]:
    st.subheader(
        "🗺️ Mapa das obras pendentes"
    )

    if st.checkbox(
        "Carregar mapa interativo"
    ):
        try:
            import folium

            from folium.plugins import MarkerCluster
            from streamlit_folium import st_folium

            pontos = visualizacao.dropna(
                subset=[
                    "LATITUDE",
                    "LONGITUDE"
                ]
            )

            if pontos.empty:
                st.warning(
                    "Nenhuma obra com coordenadas "
                    "validas para exibir."
                )

            else:
                centro = [
                    float(pontos["LATITUDE"].mean()),
                    float(pontos["LONGITUDE"].mean())
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

                    subset = pontos[
                        pontos["LISTA"].eq(categoria)
                    ]

                    for _, obra in subset.iterrows():
                        descricao = (
                            f"Nota: {obra['NOTA']}\n"
                            f"Municipio: {obra['MUNICIPIO']}\n"
                            f"Categoria: {categoria}\n"
                            f"Grupo: "
                            f"{obra['GRUPO_DESLOCAMENTO']}\n"
                            f"Prioridade: "
                            f"{obra['PRIORIDADE_PLANEJAMENTO']}"
                        )

                        folium.Marker(
                            location=[
                                float(obra["LATITUDE"]),
                                float(obra["LONGITUDE"])
                            ],
                            popup=folium.Popup(
                                html.escape(
                                    descricao
                                ).replace("\n", "<br>"),
                                max_width=450
                            ),
                            icon=folium.Icon(
                                color=CORES[categoria]
                            )
                        ).add_to(cluster)

                    camada.add_to(mapa)

                folium.LayerControl().add_to(mapa)

                st_folium(
                    mapa,
                    use_container_width=True,
                    height=550
                )

        except ImportError:
            st.error(
                "Instale folium e streamlit-folium "
                "para utilizar o mapa."
            )


# ------------------------------------------------------------
# HISTORICO
# ------------------------------------------------------------

with abas[6]:
    st.subheader(
        "📈 Comparacao com analise anterior"
    )

    if arquivo_anterior is None:
        st.info(
            "Envie a planilha de uma analise "
            "anterior na barra lateral."
        )

    else:
        c1, c2 = st.columns(2)

        c1.metric(
            "Notas novas",
            len(novas)
        )

        c2.metric(
            "Notas que sairam",
            len(removidas)
        )

        st.markdown("**Novas notas**")

        st.dataframe(
            novas,
            hide_index=True,
            use_container_width=True
        )

        st.markdown(
            "**Notas que sairam da pendencia**"
        )

        st.dataframe(
            removidas,
            hide_index=True,
            use_container_width=True
        )


# ============================================================
# EXPORTACOES EXCEL E KML
# ============================================================

st.divider()

st.subheader(
    "📥 Exportar resultados"
)

st.caption(
    "O Excel inclui o consolidado, as três "
    "categorias de obras, auditoria, "
    "programação e histórico. "
    "O KML inclui as obras filtradas com "
    "coordenadas válidas."
)

parametros = {
    "Raio maximo equipes (km)": raio,
    "Raio agrupamento (km)": raio_grupo,
    "Capacidade por equipe/dia": capacidade,
    "Dias de programacao": dias,
    "Equipes candidatas": quantidade_equipes
}

excel_resultado = exportar_excel(
    obras,
    auditoria,
    equipes,
    programacao,
    carga,
    novas,
    removidas,
    parametros
)

kml_resultado = exportar_kml(
    visualizacao
)

col_excel, col_kml = st.columns(2)

with col_excel:
    st.download_button(
        label="📊 Exportar planilha Excel",
        data=excel_resultado,
        file_name=(
            "Planejamento_NIP_"
            + datetime.now().strftime("%Y%m%d_%H%M")
            + ".xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

with col_kml:
    st.download_button(
        label="🗺️ Exportar arquivo KML",
        data=kml_resultado,
        file_name=(
            "Obras_NIP_"
            + datetime.now().strftime("%Y%m%d_%H%M")
            + ".kml"
        ),
        mime="application/vnd.google-earth.kml+xml",
        use_container_width=True
    )

st.caption(
    f"Linhas na base Saneamento: "
    f"{dados['linhas_san']:,} | "
    f"Linhas na base Levantamento: "
    f"{dados['linhas_lev']:,} | "
    f"Notas unicas analisadas: "
    f"{len(auditoria):,}"
)
