
import io
import re
import time
import hashlib
import unicodedata
import xml.etree.ElementTree as ET

from collections import defaultdict
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Font, PatternFill
from sklearn.neighbors import BallTree

# ============================================================
# CONFIGURACOES
# ============================================================

st.set_page_config(
    page_title="NIP | Planejamento de Obras",
    page_icon="📍",
    layout="wide"
)

EARTH_KM = 6371.0088

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_DUP = "OBRA SANEAMENTO E LEVANTAMENTO"

CATEGORIAS = [CAT_SAN, CAT_LEV, CAT_DUP]

SAP_EXCLUIR = {"FINL", "CANC"}

LIST_ACEITOS = {
    "0",
    "EM LEVANTAMENTO",
    "CORRECAO DE LEVANTAMENTO"
}

CONTRATOS_ACEITOS = {
    "0",
    "NIP GLOBAL LTDA - EQTL MARANHAO"
}

COLUNAS_FINAL = [
    "NOTA", "LISTA", "MUNICIPIO", "REGIONAL",
    "LATITUDE", "LONGITUDE", "STATUS_SAP",
    "STATUS_LIST", "CONTRATO", "COLUNA_M",
    "COLUNA_N", "PRIORIDADE", "DATA_ABERTURA",
    "DUPLICADA", "PENDENTE_CONTAGEM",
    "MOTIVO_EXCLUSAO"
]

OSRM_PADRAO = "https://router.project-osrm.org"

# ============================================================
# FUNCOES BASICAS
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


def nota_chave(valor):
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


def procurar_coluna(df, nomes, obrigatoria=True):
    mapa = {
        norm(c): c for c in df.columns
    }

    for nome in nomes:
        if norm(nome) in mapa:
            return mapa[norm(nome)]

    if obrigatoria:
        raise ValueError(
            "Coluna nao encontrada: "
            + " / ".join(nomes)
        )

    return None


def procurar_aba(arquivo, colunas_essenciais, preferida=None):
    """
    Procura automaticamente a aba com os cabecalhos
    necessarios, mesmo quando nao se chama NOTAS.
    """
    dados = io.BytesIO(arquivo.getvalue())
    excel = pd.ExcelFile(dados, engine="openpyxl")

    abas = list(excel.sheet_names)

    if preferida:
        abas.sort(
            key=lambda x: (
                0 if norm(x) == norm(preferida) else 1
            )
        )

    for aba in abas:
        try:
            amostra = pd.read_excel(
                excel,
                sheet_name=aba,
                nrows=5,
                dtype=str
            )

            colunas = {
                norm(c) for c in amostra.columns
            }

            if all(
                norm(c) in colunas
                for c in colunas_essenciais
            ):
                df = pd.read_excel(
                    excel,
                    sheet_name=aba,
                    dtype=str
                )

                return df, aba

        except Exception:
            continue

    raise ValueError(
        "Nenhuma aba compativel encontrada. "
        "Abas existentes: "
        + ", ".join(excel.sheet_names)
        + ". Colunas obrigatorias: "
        + ", ".join(colunas_essenciais)
    )


def ler_arquivo(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        for encoding in ["utf-8-sig", "latin-1"]:
            try:
                return pd.read_csv(
                    io.BytesIO(arquivo.getvalue()),
                    sep=None,
                    engine="python",
                    encoding=encoding,
                    dtype=str
                )
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                pass

        raise ValueError("CSV invalido.")

    return pd.read_excel(
        io.BytesIO(arquivo.getvalue()),
        dtype=str,
        engine="openpyxl"
    )


def coordenadas(df, c_lat, c_lon):
    def converter(serie):
        return pd.to_numeric(
            serie.astype(str).str.replace(
                ",", ".", regex=False
            ),
            errors="coerce"
        )

    lat = converter(df[c_lat])
    lon = converter(df[c_lon])

    validas = (
        lat.between(-35, 6)
        & lon.between(-75, -30)
        & lat.ne(0)
        & lon.ne(0)
    )

    return lat.where(validas), lon.where(validas)


def valor_linha(linha, coluna, padrao=""):
    if coluna is None:
        return padrao

    valor = linha.get(coluna, padrao)

    return padrao if pd.isna(valor) else valor


def haversine(lat1, lon1, lat2, lon2):
    a1 = np.radians(lat1)
    o1 = np.radians(lon1)
    a2 = np.radians(lat2)
    o2 = np.radians(lon2)

    x = (
        np.sin((a2 - a1) / 2) ** 2
        + np.cos(a1) * np.cos(a2)
        * np.sin((o2 - o1) / 2) ** 2
    )

    return 2 * EARTH_KM * np.arcsin(
        np.sqrt(np.clip(x, 0, 1))
    )


# ============================================================
# CARREGAMENTO DAS BASES
# ============================================================

def carregar_saneamento(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        df = ler_arquivo(arquivo)
        aba = "CSV"
    else:
        df, aba = procurar_aba(
            arquivo,
            ["NOTA"],
            preferida="Clientes Existentes"
        )

    c_nota = procurar_coluna(df, ["NOTA"])

    c_cidade = procurar_coluna(
        df, ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]
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

    df["CHAVE_NOTA"] = df[c_nota].map(nota_chave)

    df["MUNICIPIO_OBRA"] = df[c_cidade].fillna("")

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None else ""
    )

    df["LAT_OBRA"], df["LON_OBRA"] = coordenadas(
        df, c_lat, c_lon
    )

    return df, aba


def motivos_exclusao(linha):
    motivos = []

    sap = linha["SAP_NORM"]
    lista = linha["LIST_NORM"]
    contrato = linha["CONTRATO_NORM"]
    m = linha["M_NORM"]
    n = linha["N_NORM"]

    # SAP: somente FINL e CANC sao excluidos.
    if sap in SAP_EXCLUIR:
        motivos.append(f"STATUS SAP {sap}")

    # LIST: zero, Em levantamento e Correcao
    # de levantamento sao considerados.
    if lista not in LIST_ACEITOS:
        motivos.append(
            f"STATUS LIST NAO ACEITO: "
            f"{lista or 'VAZIO'}"
        )

    if contrato not in CONTRATOS_ACEITOS:
        motivos.append(
            f"CONTRATO NAO ACEITO: "
            f"{contrato or 'VAZIO'}"
        )

    if m != "0":
        motivos.append(
            f"COLUNA M DIFERENTE DE 0: "
            f"{m or 'VAZIO'}"
        )

    if n != "0":
        motivos.append(
            f"COLUNA N DIFERENTE DE 0: "
            f"{n or 'VAZIO'}"
        )

    return " | ".join(motivos)


def carregar_levantamento(arquivo):
    df, aba = procurar_aba(
        arquivo,
        ["PROTOCOLO", "STATUS SAP", "STATUS LIST"],
        preferida="NOTAS"
    )

    c_nota = procurar_coluna(df, ["PROTOCOLO"])
    c_sap = procurar_coluna(df, ["STATUS SAP"])
    c_list = procurar_coluna(df, ["STATUS LIST"])
    c_contrato = procurar_coluna(df, ["CONTRATO"])

    c_cidade = procurar_coluna(
        df, ["MUNICIPIO", "MUNICÍPIO"]
    )

    c_regional = procurar_coluna(
        df, ["REGIONAL"], False
    )

    c_lat = procurar_coluna(df, ["LATITUDE"])
    c_lon = procurar_coluna(df, ["LONGITUDE"])

    c_m = procurar_coluna(
        df,
        [
            "ORÇAMENTO MODULAR (NÃO ENVIAR A CAMPO)",
            "ORCAMENTO MODULAR (NAO ENVIAR A CAMPO)"
        ],
        False
    )

    c_n = procurar_coluna(
        df,
        [
            "PLA ALVOS LEVANTAMENTOS (NÃO ENVIAR A CAMPO)",
            "PLA ALVOS LEVANTAMENTOS (NAO ENVIAR A CAMPO)"
        ],
        False
    )

    if c_m is None or c_n is None:
        raise ValueError(
            "Nao foram encontradas as colunas "
            "ORCAMENTO MODULAR e PLA ALVOS."
        )

    c_prioridade = procurar_coluna(
        df, ["PRIORIDADE"], False
    )

    c_abertura = procurar_coluna(
        df, ["DATA ABERTURA", "DATA DE ABERTURA"],
        False
    )

    df["CHAVE_NOTA"] = df[c_nota].map(nota_chave)

    df["MUNICIPIO_OBRA"] = df[c_cidade].fillna("")

    df["REGIONAL_OBRA"] = (
        df[c_regional].fillna("")
        if c_regional is not None else ""
    )

    df["LAT_OBRA"], df["LON_OBRA"] = coordenadas(
        df, c_lat, c_lon
    )

    df["SAP_NORM"] = df[c_sap].map(status)
    df["LIST_NORM"] = df[c_list].map(status)
    df["CONTRATO_NORM"] = df[c_contrato].map(status)

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

    df["MOTIVOS_EXCLUSAO"] = df.apply(
        motivos_exclusao, axis=1
    )

    return df, aba


def carregar_equipes(arquivo):
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
                x for x in excel.sheet_names
                if norm(x) == norm(aba_esperada)
            ),
            None
        )

        if aba is None:
            raise ValueError(
                f"Aba ausente: {aba_esperada}"
            )

        df = pd.read_excel(
            excel,
            sheet_name=aba,
            dtype=str
        )

        c_nome = procurar_coluna(
            df, ["NOME", "EQUIPE"]
        )

        c_cidade = procurar_coluna(
            df, ["CIDADES", "CIDADE", "MUNICIPIO"]
        )

        c_lat = procurar_coluna(
            df, ["LATITUDE", "LAT"]
        )

        c_lon = procurar_coluna(
            df, ["LONGITUDE", "LON"]
        )

        lat, lon = coordenadas(
            df, c_lat, c_lon
        )

        partes.append(
            pd.DataFrame({
                "EQUIPE": df[c_nome],
                "CIDADE_BASE": df[c_cidade],
                "LAT_EQUIPE": lat,
                "LON_EQUIPE": lon,
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
    ).copy()

    equipes = equipes[
        equipes["EQUIPE"].astype(str).str.strip().ne("")
    ].copy()

    equipes["ID_EQUIPE"] = (
        equipes["TIPO_EQUIPE"].astype(str)
        + "|"
        + equipes["EQUIPE"].astype(str)
        + "|"
        + equipes["CIDADE_BASE"].astype(str)
    )

    if equipes.empty:
        raise ValueError(
            "Nenhuma equipe com coordenadas validas."
        )

    return equipes.reset_index(drop=True)


# ============================================================
# CRUZAMENTO E AUDITORIA
# ============================================================

def primeira_linha(grupo):
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

    gs = dict(
        tuple(san.groupby("CHAVE_NOTA", sort=False))
    )

    gl = dict(
        tuple(lev.groupby("CHAVE_NOTA", sort=False))
    )

    notas = dict.fromkeys(
        list(gs.keys()) + list(gl.keys())
    )

    registros = []

    for nota in notas:
        grupo_s = gs.get(nota)
        grupo_l = gl.get(nota)

        tem_s = grupo_s is not None
        tem_l = grupo_l is not None

        a = primeira_linha(grupo_s) if tem_s else None
        b = primeira_linha(grupo_l) if tem_l else None

        if tem_s and tem_l:
            categoria = CAT_DUP
        elif tem_s:
            categoria = CAT_SAN
        else:
            categoria = CAT_LEV

        motivos = []

        if tem_l:
            for texto in grupo_l["MOTIVOS_EXCLUSAO"]:
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

        registros.append({
            "NOTA": nota,
            "LISTA": categoria,
            "MUNICIPIO": origem["MUNICIPIO_OBRA"],
            "REGIONAL": origem["REGIONAL_OBRA"],
            "LATITUDE": origem["LAT_OBRA"],
            "LONGITUDE": origem["LON_OBRA"],
            "STATUS_SAP": (
                b["SAP_NORM"] if tem_l else ""
            ),
            "STATUS_LIST": (
                b["LIST_NORM"] if tem_l else ""
            ),
            "CONTRATO": (
                b["CONTRATO_NORM"] if tem_l else ""
            ),
            "COLUNA_M": (
                b["M_NORM"] if tem_l else ""
            ),
            "COLUNA_N": (
                b["N_NORM"] if tem_l else ""
            ),
            "PRIORIDADE": (
                b["PRIORIDADE_NORM"] if tem_l else ""
            ),
            "DATA_ABERTURA": (
                b["DATA_ABERTURA_NORM"] if tem_l else ""
            ),
            "DUPLICADA": (
                "SIM" if tem_s and tem_l else "NÃO"
            ),
            "PENDENTE_CONTAGEM": (
                "NÃO" if motivos else "SIM"
            ),
            "MOTIVO_EXCLUSAO": (
                " | ".join(motivos) if motivos else "-"
            ),
            "OCORRENCIAS_SANEAMENTO": (
                len(grupo_s) if tem_s else 0
            ),
            "OCORRENCIAS_LEVANTAMENTO": (
                len(grupo_l) if tem_l else 0
            )
        })

    return pd.DataFrame(
        registros,
        columns=COLUNAS_FINAL + [
            "OCORRENCIAS_SANEAMENTO",
            "OCORRENCIAS_LEVANTAMENTO"
        ]
    )


# ============================================================
# PRIORIDADE E AGRUPAMENTO
# ============================================================

def priorizar(obras, dias_media):
    df = obras.copy()

    datas = pd.to_datetime(
        df["DATA_ABERTURA"],
        dayfirst=True,
        errors="coerce"
    )

    df["DIAS_ABERTURA"] = (
        pd.Timestamp.today().normalize() - datas
    ).dt.days

    def classificar(r):
        p = norm(r["PRIORIDADE"])
        idade = r["DIAS_ABERTURA"]

        if any(
            palavra in p
            for palavra in ["URGENTE", "CRITICA", "ALTA"]
        ):
            return "ALTA"

        if pd.notna(idade):
            if idade >= dias_media * 4:
                return "ALTA"

            if idade >= dias_media:
                return "MEDIA"

        return "BAIXA"

    df["PRIORIDADE_PLANEJAMENTO"] = df.apply(
        classificar, axis=1
    )

    return df


def agrupar(obras, raio_km):
    df = obras.copy().reset_index(drop=True)

    df["GRUPO"] = ""
    df["QTD_GRUPO"] = 0

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
        pontos, metric="haversine"
    )

    vizinhos = arvore.query_radius(
        pontos,
        r=raio_km / EARTH_KM
    )

    pais = list(range(len(validas)))

    def raiz(i):
        while pais[i] != i:
            pais[i] = pais[pais[i]]
            i = pais[i]
        return i

    for i, proximos in enumerate(vizinhos):
        for j in proximos:
            ri = raiz(i)
            rj = raiz(int(j))
            if ri != rj:
                pais[rj] = ri

    grupos = {}

    for i, indice in enumerate(validas.index):
        r = raiz(i)

        if r not in grupos:
            grupos[r] = f"GRP-{len(grupos)+1:05d}"

        df.at[indice, "GRUPO"] = grupos[r]

    contagem = df[
        df["GRUPO"].ne("")
    ]["GRUPO"].value_counts()

    df["QTD_GRUPO"] = (
        df["GRUPO"]
        .map(contagem)
        .fillna(0)
        .astype(int)
    )

    return df


# ============================================================
# ROTAS REAIS POR RUAS E ESTRADAS
# ============================================================

def tipos_atividade(categoria):
    if categoria == CAT_DUP:
        return ["SANEAMENTO", "LEVANTAMENTO"]

    if categoria == CAT_SAN:
        return ["SANEAMENTO"]

    return ["LEVANTAMENTO"]


@st.cache_data(
    ttl=86400,
    show_spinner=False
)
def consultar_rota_osrm(
    lat_origem,
    lon_origem,
    lat_destino,
    lon_destino,
    servidor
):
    """
    Retorna distancia rodoviaria e tempo estimado.
    Nunca substitui falha por distancia em linha reta.
    """
    url = (
        servidor.rstrip("/")
        + "/route/v1/driving/"
        + f"{lon_origem},{lat_origem};"
        + f"{lon_destino},{lat_destino}"
    )

    try:
        resposta = requests.get(
            url,
            params={
                "overview": "false",
                "steps": "false"
            },
            timeout=18,
            headers={
                "User-Agent": "NIP-Planejamento/1.0"
            }
        )

        resposta.raise_for_status()
        dados = resposta.json()

        rotas = dados.get("routes", [])

        if dados.get("code") != "Ok" or not rotas:
            return None

        rota = rotas[0]

        return {
            "KM": round(
                float(rota["distance"]) / 1000,
                2
            ),
            "MINUTOS": round(
                float(rota["duration"]) / 60,
                1
            )
        }

    except (
        requests.RequestException,
        ValueError,
        KeyError,
        TypeError
    ):
        return None


def gerar_tarefas(obras):
    registros = []

    for _, obra in obras.iterrows():
        for tipo in tipos_atividade(obra["LISTA"]):
            r = obra.to_dict()
            r["ATIVIDADE"] = tipo
            registros.append(r)

    return pd.DataFrame(registros)


def calcular_rotas(
    tarefas,
    equipes,
    raio_km,
    qtd_candidatos,
    max_consultas,
    servidor
):
    """
    Seleciona candidatos usando linha reta apenas como
    pre-filtro e consulta a distancia rodoviaria real.
    """
    if tarefas.empty:
        return pd.DataFrame(), pd.DataFrame()

    lista_rotas = []
    registros = []

    consultas = 0

    progresso = st.progress(
        0, text="Consultando trajetos rodoviarios..."
    )

    total = len(tarefas)

    for posicao, (_, tarefa) in enumerate(
        tarefas.iterrows()
    ):
        numero = posicao + 1

        progresso.progress(
            numero / total,
            text=(
                f"Consultando rotas: "
                f"{numero}/{total} tarefas"
            )
        )

        candidatos = equipes[
            equipes["TIPO_EQUIPE"].eq(
                tarefa["ATIVIDADE"]
            )
        ].copy()

        resultado = tarefa.to_dict()

        resultado["EQUIPE_MAIS_PROXIMA"] = ""
        resultado["DISTANCIA_RODOVIARIA_KM"] = np.nan
        resultado["TEMPO_RODOVIARIO_MIN"] = np.nan
        resultado["STATUS_ROTA"] = "SEM ROTA"

        if (
            pd.isna(tarefa["LATITUDE"])
            or pd.isna(tarefa["LONGITUDE"])
        ):
            resultado["STATUS_ROTA"] = (
                "COORDENADAS INVALIDAS"
            )

            registros.append(resultado)
            continue

        if candidatos.empty:
            resultado["STATUS_ROTA"] = (
                "SEM EQUIPE DISPONIVEL"
            )

            registros.append(resultado)
            continue

        dist_reta = haversine(
            float(tarefa["LATITUDE"]),
            float(tarefa["LONGITUDE"]),
            candidatos["LAT_EQUIPE"].to_numpy(float),
            candidatos["LON_EQUIPE"].to_numpy(float)
        )

        candidatos["KM_RETA"] = dist_reta

        # Uma estrada nunca e mais curta que a distancia
        # geodesica entre origem e destino.
        candidatos = candidatos[
            candidatos["KM_RETA"].le(raio_km)
        ].sort_values(
            "KM_RETA"
        ).head(qtd_candidatos)

        if candidatos.empty:
            resultado["STATUS_ROTA"] = (
                "SEM EQUIPE NO RAIO"
            )

            registros.append(resultado)
            continue

        for _, equipe in candidatos.iterrows():
            if consultas >= max_consultas:
                break

            consultas += 1

            rota = consultar_rota_osrm(
                round(float(equipe["LAT_EQUIPE"]), 6),
                round(float(equipe["LON_EQUIPE"]), 6),
                round(float(tarefa["LATITUDE"]), 6),
                round(float(tarefa["LONGITUDE"]), 6),
                servidor
            )

            if rota is None:
                continue

            # O limite e aplicado na distancia rodoviaria.
            if rota["KM"] > raio_km:
                continue

            lista_rotas.append({
                "NOTA": tarefa["NOTA"],
                "ATIVIDADE": tarefa["ATIVIDADE"],
                "ID_EQUIPE": equipe["ID_EQUIPE"],
                "EQUIPE": equipe["EQUIPE"],
                "CIDADE_BASE": equipe["CIDADE_BASE"],
                "DISTANCIA_KM": rota["KM"],
                "TEMPO_MIN": rota["MINUTOS"]
            })

        rotas_tarefa = [
            r for r in lista_rotas
            if (
                r["NOTA"] == tarefa["NOTA"]
                and r["ATIVIDADE"] == tarefa["ATIVIDADE"]
            )
        ]

        if rotas_tarefa:
            melhor = min(
                rotas_tarefa,
                key=lambda x: x["DISTANCIA_KM"]
            )

            resultado["EQUIPE_MAIS_PROXIMA"] = melhor["EQUIPE"]
            resultado["DISTANCIA_RODOVIARIA_KM"] = (
                melhor["DISTANCIA_KM"]
            )
            resultado["TEMPO_RODOVIARIO_MIN"] = (
                melhor["TEMPO_MIN"]
            )
            resultado["STATUS_ROTA"] = "ROTA CALCULADA"

        elif consultas >= max_consultas:
            resultado["STATUS_ROTA"] = (
                "LIMITE DE CONSULTAS"
            )

        else:
            resultado["STATUS_ROTA"] = (
                "ROTA NAO ENCONTRADA"
            )

        registros.append(resultado)

    progresso.empty()

    return (
        pd.DataFrame(registros),
        pd.DataFrame(
            lista_rotas,
            columns=[
                "NOTA", "ATIVIDADE", "ID_EQUIPE",
                "EQUIPE", "CIDADE_BASE",
                "DISTANCIA_KM", "TEMPO_MIN"
            ]
        )
    )


# ============================================================
# PROGRAMACAO POR QUANTIDADE / DIA / SEMANA / MES
# ============================================================

def datas_programacao(inicio, modo, periodos):
    """
    Gera dias uteis para uma programacao operacional.
    Sem feriados municipais, estaduais ou nacionais.
    """
    datas = []

    data = pd.Timestamp(inicio).date()

    if modo == "POR QUANTIDADE":
        limite_dias = max(20, periodos * 31)
    elif modo == "POR DIA":
        limite_dias = periodos * 2 + 20
    elif modo == "POR SEMANA":
        limite_dias = periodos * 7
    else:
        limite_dias = (
            pd.Timestamp(data)
            + pd.DateOffset(months=periodos)
            - pd.Timedelta(days=1)
        ).date()

    if isinstance(limite_dias, int):
        fim = data + timedelta(days=limite_dias)
    else:
        fim = limite_dias

    while data <= fim:
        if data.weekday() < 5:
            datas.append(data)

        data += timedelta(days=1)

    if modo == "POR DIA":
        return datas[:periodos]

    return datas


def programar(
    tarefas,
    rotas,
    modo,
    limite,
    periodos,
    data_inicio,
    capacidade_dia
):
    """
    Aloca tarefas por equipe dentro de cada periodo.
    Se nao houver rota calculada, nao aloca a tarefa.
    """
    if tarefas.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = tarefas.copy()

    ordem = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    df["_PRIORIDADE"] = df[
        "PRIORIDADE_PLANEJAMENTO"
    ].map(ordem).fillna(3)

    df = df.sort_values(
        ["_PRIORIDADE", "GRUPO", "NOTA"]
    )

    datas = datas_programacao(
        data_inicio, modo, periodos
    )

    ocupacao_dia = defaultdict(int)
    ocupacao_periodo = defaultdict(int)

    rotas_por_tarefa = defaultdict(list)

    for _, rota in rotas.iterrows():
        chave = (
            rota["NOTA"],
            rota["ATIVIDADE"]
        )

        rotas_por_tarefa[chave].append(
            rota.to_dict()
        )

    resultado = []

    for _, tarefa in df.iterrows():
        chave_tarefa = (
            tarefa["NOTA"],
            tarefa["ATIVIDADE"]
        )

        candidatas = rotas_por_tarefa.get(
            chave_tarefa, []
        )

        escolha = None
        alternativas = []

        for rota in candidatas:
            for data in datas:
                if modo == "POR SEMANA":
                    iso = data.isocalendar()
                    periodo = (
                        f"{iso.year}-S{iso.week:02d}"
                    )

                    primeira_semana = (
                        pd.Timestamp(data_inicio)
                        .date().isocalendar()
                    )

                    semana_zero = (
                        pd.Timestamp.fromisocalendar(
                            primeira_semana.year,
                            primeira_semana.week,
                            1
                        ).date()
                    )

                    indice_semana = (
                        (
                            data
                            - semana_zero
                        ).days // 7
                    )

                    if indice_semana >= periodos:
                        continue

                elif modo == "POR MES":
                    periodo = data.strftime("%Y-%m")

                else:
                    periodo = data.isoformat()

                chave_dia = (
                    rota["ID_EQUIPE"],
                    data
                )

                chave_periodo = (
                    rota["ID_EQUIPE"],
                    periodo
                )

                if ocupacao_dia[chave_dia] >= capacidade_dia:
                    continue

                # Quantidade maxima por equipe no periodo.
                if ocupacao_periodo[chave_periodo] >= limite:
                    continue

                alternativas.append((
                    ocupacao_periodo[chave_periodo],
                    data,
                    float(rota["DISTANCIA_KM"]),
                    rota["EQUIPE"],
                    rota["ID_EQUIPE"],
                    periodo,
                    rota["TEMPO_MIN"]
                ))

        if alternativas:
            escolha = min(
                alternativas,
                key=lambda a: (
                    a[0], a[1], a[2]
                )
            )

        registro = tarefa.drop(
            labels=["_PRIORIDADE"]
        ).to_dict()

        if escolha is None:
            registro.update({
                "EQUIPE_PROGRAMADA": "NÃO ALOCADA",
                "DATA_PROGRAMADA": "",
                "SEMANA_PROGRAMADA": "",
                "MES_PROGRAMADO": "",
                "PERIODO_PROGRAMADO": "",
                "DISTANCIA_PROGRAMADA_KM": np.nan,
                "TEMPO_PROGRAMADO_MIN": np.nan,
                "STATUS_PROGRAMACAO": "SEM ALOCACAO"
            })

        else:
            (
                _,
                data,
                distancia,
                equipe,
                id_equipe,
                periodo,
                tempo
            ) = escolha

            ocupacao_dia[
                (id_equipe, data)
            ] += 1

            ocupacao_periodo[
                (id_equipe, periodo)
            ] += 1

            iso = data.isocalendar()

            registro.update({
                "EQUIPE_PROGRAMADA": equipe,
                "DATA_PROGRAMADA": data.isoformat(),
                "SEMANA_PROGRAMADA": (
                    f"{iso.year}-S{iso.week:02d}"
                ),
                "MES_PROGRAMADO": data.strftime("%Y-%m"),
                "PERIODO_PROGRAMADO": periodo,
                "DISTANCIA_PROGRAMADA_KM": distancia,
                "TEMPO_PROGRAMADO_MIN": tempo,
                "STATUS_PROGRAMACAO": "SUGESTAO"
            })

        resultado.append(registro)

    saida = pd.DataFrame(resultado)

    alocadas = saida[
        saida["STATUS_PROGRAMACAO"].eq("SUGESTAO")
    ]

    carga = (
        alocadas.groupby(
            [
                "ATIVIDADE",
                "EQUIPE_PROGRAMADA",
                "PERIODO_PROGRAMADO"
            ]
        ).size().reset_index(name="QUANTIDADE")
    )

    return saida, carga


# ============================================================
# HISTORICO
# ============================================================

def comparar_historico(obras, arquivo):
    planilhas = pd.read_excel(
        io.BytesIO(arquivo.getvalue()),
        sheet_name=None,
        dtype=str,
        engine="openpyxl"
    )

    anterior = None

    for nome in [
        "CONSOLIDADO",
        "AUDITORIA COMPLETA",
        "OBRAS PENDENTES"
    ]:
        if nome in planilhas:
            anterior = planilhas[nome]
            break

    if anterior is None:
        raise ValueError(
            "Nao foi encontrada aba CONSOLIDADO "
            "ou AUDITORIA COMPLETA."
        )

    if "NOTA" not in anterior.columns:
        raise ValueError(
            "A planilha anterior nao possui NOTA."
        )

    if "PENDENTE_CONTAGEM" in anterior.columns:
        anterior = anterior[
            anterior["PENDENTE_CONTAGEM"].eq("SIM")
        ]

    antigas = set(
        anterior["NOTA"].map(nota_chave)
    ) - {""}

    atuais = set(
        obras["NOTA"].map(nota_chave)
    ) - {""}

    novas = pd.DataFrame({
        "NOTA": sorted(atuais - antigas)
    })

    removidas = pd.DataFrame({
        "NOTA": sorted(antigas - atuais)
    })

    return novas, removidas


# ============================================================
# EXPORTACAO KML
# ============================================================

def exportar_kml(obras):
    namespace = "http://www.opengis.net/kml/2.2"
    ET.register_namespace("", namespace)

    def el(pai, nome, texto=None):
        item = ET.SubElement(
            pai,
            f"{{{namespace}}}{nome}"
        )

        if texto is not None:
            item.text = str(texto)

        return item

    raiz = ET.Element(
        f"{{{namespace}}}kml"
    )

    doc = el(raiz, "Document")

    el(doc, "name", "NIP - Obras Pendentes")

    for categoria in CATEGORIAS:
        pasta = el(doc, "Folder")
        el(pasta, "name", categoria)

        subset = obras[
            obras["LISTA"].eq(categoria)
        ].dropna(
            subset=["LATITUDE", "LONGITUDE"]
        )

        for _, r in subset.iterrows():
            ponto = el(pasta, "Placemark")
            el(ponto, "name", r["NOTA"])

            descricao = (
                f"NOTA: {r['NOTA']}\n"
                f"MUNICIPIO: {r['MUNICIPIO']}\n"
                f"REGIONAL: {r['REGIONAL']}\n"
                f"CATEGORIA: {r['LISTA']}\n"
                f"PRIORIDADE: "
                f"{r.get('PRIORIDADE_PLANEJAMENTO', '')}\n"
                f"GRUPO: {r.get('GRUPO', '')}"
            )

            el(ponto, "description", descricao)

            geometria = el(ponto, "Point")

            el(
                geometria,
                "coordinates",
                (
                    f"{float(r['LONGITUDE'])},"
                    f"{float(r['LATITUDE'])},0"
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
    tarefas,
    rotas,
    programacao,
    carga,
    equipes,
    novas,
    removidas,
    parametros
):
    excluidas = auditoria[
        auditoria["PENDENTE_CONTAGEM"].eq("NÃO")
    ]

    motivos = []

    for _, r in excluidas.iterrows():
        for motivo in str(
            r["MOTIVO_EXCLUSAO"]
        ).split(" | "):
            motivos.append({
                "NOTA": r["NOTA"],
                "MOTIVO": motivo
            })

    resumo = [
        [
            "Total obras pendentes",
            len(obras)
        ],
        [
            "Saneamento",
            int(obras["LISTA"].eq(CAT_SAN).sum())
        ],
        [
            "Levantamento",
            int(obras["LISTA"].eq(CAT_LEV).sum())
        ],
        [
            "Duplicadas",
            int(obras["LISTA"].eq(CAT_DUP).sum())
        ],
        [
            "Excluidas",
            len(excluidas)
        ],
        [
            "Tarefas na programacao",
            len(programacao)
        ],
        [
            "Tarefas alocadas",
            int(
                programacao["STATUS_PROGRAMACAO"]
                .eq("SUGESTAO").sum()
            ) if not programacao.empty else 0
        ],
        [
            "Gerado em",
            datetime.now().strftime("%d/%m/%Y %H:%M")
        ]
    ]

    for chave, valor in parametros.items():
        resumo.append([chave, valor])

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
            obras["LISTA"].eq(CAT_DUP)
        ],
        "AUDITORIA COMPLETA": auditoria,
        "EXCLUIDAS": excluidas,
        "MOTIVOS EXCLUSAO": pd.DataFrame(motivos),
        "TAREFAS ROTAS": tarefas,
        "ROTAS RODOVIARIAS": rotas,
        "PROGRAMACAO": programacao,
        "CARGA EQUIPES": carga,
        "EQUIPES": equipes,
        "NOTAS NOVAS": novas,
        "NOTAS REMOVIDAS": removidas
    }

    memoria = io.BytesIO()

    with pd.ExcelWriter(
        memoria, engine="openpyxl"
    ) as writer:
        for nome, df in planilhas.items():
            df.to_excel(
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
                    bold=True,
                    color="FFFFFF"
                )

            for coluna in ws.columns:
                valores = [
                    len(str(c.value or ""))
                    for c in list(coluna)[:100]
                ]

                largura = min(
                    55,
                    max(13, max(valores, default=10) + 2)
                )

                ws.column_dimensions[
                    coluna[0].column_letter
                ].width = largura

    return memoria.getvalue()


# ============================================================
# INTERFACE
# ============================================================

st.title("📍 NIP | Planejamento de Obras")

st.caption(
    "Cruzamento de Saneamento e Levantamento, "
    "rotas por estradas, programacao e exportacoes."
)

with st.expander("📘 Regras de contagem"):
    st.markdown("""
    **STATUS SAP**

    - Excluir somente FINL e CANC.
    - Considerar os demais status SAP.

    **STATUS LIST**

    - Considerar 0.
    - Considerar Em levantamento.
    - Considerar Correção de levantamento.
    - Outros valores ficam fora da contagem.

    **Demais filtros**

    - Contrato 0 ou NIP GLOBAL LTDA - EQTL MARANHÃO.
    - Colunas M e N iguais a 0.
    - Duplicadas aparecem uma única vez na volumetria.
    - Obras presentes somente no Saneamento continuam
      na contagem quando não há bloqueio conhecido.

    **Distâncias**

    Rotas calculadas por serviço rodoviário OSRM.
    Rotas não encontradas não recebem distâncias
    estimadas em linha reta.
    """)

# ============================================================
# UPLOADS
# ============================================================

st.subheader("📂 Bases obrigatorias")

c1, c2, c3 = st.columns(3)

with c1:
    arquivo_san = st.file_uploader(
        "1. BASE_SANEAMENTO",
        type=["xlsx", "csv"]
    )

with c2:
    arquivo_lev = st.file_uploader(
        "2. BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"]
    )

with c3:
    arquivo_equipes = st.file_uploader(
        "3. LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"]
    )

if not all([
    arquivo_san,
    arquivo_lev,
    arquivo_equipes
]):
    st.warning(
        "Envie os tres arquivos obrigatorios."
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
    "🚀 PROCESSAR BASES",
    type="primary",
    use_container_width=True
):
    try:
        with st.spinner("Processando bases..."):
            san, aba_san = carregar_saneamento(
                arquivo_san
            )

            lev, aba_lev = carregar_levantamento(
                arquivo_lev
            )

            equipes = carregar_equipes(
                arquivo_equipes
            )

            auditoria = consolidar(san, lev)

            if auditoria.empty:
                raise ValueError(
                    "Nenhuma nota valida encontrada."
                )

            st.session_state["nip_dados"] = {
                "assinatura": assinatura,
                "auditoria": auditoria,
                "equipes": equipes,
                "aba_san": aba_san,
                "aba_lev": aba_lev,
                "qtd_san": len(san),
                "qtd_lev": len(lev)
            }

            # Invalida rotas da analise anterior.
            st.session_state.pop(
                "nip_rotas", None
            )

        st.success(
            "Bases processadas com sucesso!"
        )

    except Exception as erro:
        st.error(
            f"Erro no processamento: {erro}"
        )

if "nip_dados" not in st.session_state:
    st.stop()

dados = st.session_state["nip_dados"]

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
    st.header("⚙️ Configuracoes")

    raio = st.select_slider(
        "Raio rodoviario maximo (km)",
        options=list(range(200, 501, 25)),
        value=200
    )

    qtd_candidatos = st.slider(
        "Equipes candidatas por obra",
        1, 5, 3
    )

    raio_grupo = st.select_slider(
        "Raio para agrupar obras (km)",
        options=[5, 10, 20, 30, 50],
        value=20
    )

    dias_prioridade = st.number_input(
        "Dias para prioridade MEDIA",
        min_value=1,
        value=14
    )

    st.divider()

    st.subheader("📅 Programacao")

    modo = st.selectbox(
        "Tipo de programacao",
        [
            "POR QUANTIDADE",
            "POR DIA",
            "POR SEMANA",
            "POR MES"
        ]
    )

    limite = st.number_input(
        "Quantidade maxima por equipe no periodo",
        min_value=1,
        value=10
    )

    capacidade_dia = st.number_input(
        "Capacidade maxima por equipe por dia",
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

    st.subheader("🛣️ Rotas rodoviarias")

    servidor = st.text_input(
        "Servidor OSRM",
        value=OSRM_PADRAO
    )

    max_consultas = st.number_input(
        "Maximo de consultas por processamento",
        min_value=1,
        max_value=2000,
        value=150
    )

    max_obras = st.number_input(
        "Maximo de obras por lote de rotas",
        min_value=1,
        max_value=2000,
        value=100
    )

    st.caption(
        "O servidor publico OSRM e demonstrativo "
        "e pode apresentar limites ou indisponibilidade. "
        "Para alto volume, use um servidor contratado "
        "ou proprio."
    )

    st.divider()

    historico = st.file_uploader(
        "Analise anterior (opcional)",
        type=["xlsx"]
    )


# ============================================================
# PENDENCIAS E INDICADORES
# ============================================================

obras = auditoria[
    auditoria["PENDENTE_CONTAGEM"].eq("SIM")
].copy()

obras = priorizar(
    obras, dias_prioridade
)

obras = agrupar(
    obras, raio_grupo
)

st.subheader("📊 Indicadores")

colunas = st.columns(5)

indicadores = [
    ("Saneamento", obras["LISTA"].eq(CAT_SAN).sum()),
    ("Levantamento", obras["LISTA"].eq(CAT_LEV).sum()),
    ("Duplicadas", obras["LISTA"].eq(CAT_DUP).sum()),
    ("Total pendente", len(obras)),
    (
        "Excluidas",
        auditoria[
            "PENDENTE_CONTAGEM"
        ].eq("NÃO").sum()
    )
]

for coluna, (titulo, valor) in zip(
    colunas, indicadores
):
    coluna.metric(titulo, int(valor))

# ============================================================
# FILTROS
# ============================================================

st.subheader("🔎 Filtrar obras")

c1, c2, c3 = st.columns(3)

regionais = sorted(
    obras["REGIONAL"].dropna().unique()
)

municipios = sorted(
    obras["MUNICIPIO"].dropna().unique()
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

busca = st.text_input("Pesquisar nota")

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
        view["PRIORIDADE_PLANEJAMENTO"].isin(
            filtro_prioridade
        )
    ]

if busca:
    view = view[
        view["NOTA"].astype(str).str.contains(
            re.escape(busca),
            case=False,
            na=False
        )
    ]

st.caption(
    f"Obras selecionadas pelos filtros: {len(view)}"
)

# ============================================================
# CALCULO DE ROTAS
# ============================================================

st.subheader("🛣️ Calculo de rotas por estradas")

st.info(
    "As rotas usam a malha viaria informada pelo "
    "servico de roteamento. A ferramenta nao "
    "substitui uma rota sem resposta por distancia "
    "em linha reta."
)

lote = view.head(int(max_obras)).copy()

st.caption(
    f"Lote preparado: {len(lote)} obras. "
    "Os filtros definem quais obras entram primeiro."
)

assinatura_lote = hashlib.sha256(
    (
        "|".join(lote["NOTA"].astype(str))
        + str(raio)
        + str(qtd_candidatos)
        + servidor
    ).encode()
).hexdigest()

if st.button(
    "🚗 CALCULAR ROTAS RODOVIARIAS",
    type="primary"
):
    tarefas_lote = gerar_tarefas(lote)

    if tarefas_lote.empty:
        st.warning("Nenhuma tarefa para calcular.")
    else:
        tarefas_rotas, rotas = calcular_rotas(
            tarefas_lote,
            equipes,
            raio,
            qtd_candidatos,
            int(max_consultas),
            servidor
        )

        st.session_state["nip_rotas"] = {
            "assinatura": assinatura_lote,
            "tarefas": tarefas_rotas,
            "rotas": rotas
        }

        st.success(
            f"Processamento concluido. "
            f"{len(rotas)} alternativas de rota "
            f"rodoviaria encontradas."
        )

rotas_validas = False

if "nip_rotas" in st.session_state:
    cache = st.session_state["nip_rotas"]

    rotas_validas = (
        cache["assinatura"] == assinatura_lote
    )

if rotas_validas:
    tarefas_rotas = cache["tarefas"]
    rotas = cache["rotas"]

else:
    tarefas_rotas = pd.DataFrame()
    rotas = pd.DataFrame(
        columns=[
            "NOTA", "ATIVIDADE", "ID_EQUIPE",
            "EQUIPE", "CIDADE_BASE",
            "DISTANCIA_KM", "TEMPO_MIN"
        ]
    )

    if "nip_rotas" in st.session_state:
        st.warning(
            "Os filtros ou parametros de rota "
            "mudaram. Calcule as rotas novamente."
        )


# ============================================================
# PROGRAMACAO
# ============================================================

if rotas_validas:
    programacao, carga = programar(
        tarefas_rotas,
        rotas,
        modo,
        int(limite),
        int(periodos),
        data_inicio,
        int(capacidade_dia)
    )

else:
    programacao = pd.DataFrame()
    carga = pd.DataFrame()


# ============================================================
# HISTORICO
# ============================================================

novas = pd.DataFrame(columns=["NOTA"])
removidas = pd.DataFrame(columns=["NOTA"])

if historico is not None:
    try:
        novas, removidas = comparar_historico(
            obras, historico
        )
    except Exception as erro:
        st.warning(
            f"Historico nao processado: {erro}"
        )


# ============================================================
# ABAS DE RESULTADOS
# ============================================================

abas = st.tabs([
    "🟣 SANEAMENTO",
    "🟢 LEVANTAMENTO",
    "🔵 DUPLICADAS",
    "📅 PROGRAMAÇÃO",
    "🛣️ ROTAS",
    "⚠️ AUDITORIA",
    "🗺️ MAPA",
    "📈 HISTÓRICO"
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


with abas[3]:
    st.subheader(
        f"📅 Programacao - {modo}"
    )

    if programacao.empty:
        st.info(
            "Calcule as rotas rodoviarias "
            "para gerar a programacao."
        )
    else:
        alocadas = int(
            programacao[
                "STATUS_PROGRAMACAO"
            ].eq("SUGESTAO").sum()
        )

        sem = int(
            programacao[
                "STATUS_PROGRAMACAO"
            ].eq("SEM ALOCACAO").sum()
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Tarefas programadas",
            alocadas
        )

        c2.metric(
            "Nao alocadas",
            sem
        )

        c3.metric(
            "Total tarefas",
            len(programacao)
        )

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
    st.subheader(
        "🛣️ Distancias pelas estradas"
    )

    if tarefas_rotas.empty:
        st.info(
            "Ainda nao foram calculadas rotas."
        )
    else:
        st.dataframe(
            tarefas_rotas,
            hide_index=True,
            use_container_width=True
        )

        st.subheader(
            "Alternativas por equipe"
        )

        st.dataframe(
            rotas,
            hide_index=True,
            use_container_width=True
        )


with abas[5]:
    excluidas = auditoria[
        auditoria["PENDENTE_CONTAGEM"].eq("NÃO")
    ]

    st.subheader(
        f"⚠️ Excluidas - {len(excluidas)}"
    )

    st.dataframe(
        excluidas,
        hide_index=True,
        use_container_width=True
    )

    motivos = []

    for _, r in excluidas.iterrows():
        for m in str(
            r["MOTIVO_EXCLUSAO"]
        ).split(" | "):
            motivos.append({
                "NOTA": r["NOTA"],
                "MOTIVO": m
            })

    if motivos:
        df_motivos = pd.DataFrame(motivos)

        resumo_motivos = (
            df_motivos["MOTIVO"]
            .value_counts()
        )

        st.bar_chart(resumo_motivos)

        st.caption(
            "Uma nota pode ter varios motivos, "
            "mas conta somente uma vez "
            "como obra excluida."
        )


with abas[6]:
    st.subheader("🗺️ Mapa das obras")

    if st.checkbox("Carregar mapa"):
        try:
            import folium
            from folium.plugins import MarkerCluster
            from streamlit_folium import st_folium

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

                cores = {
                    CAT_SAN: "purple",
                    CAT_LEV: "green",
                    CAT_DUP: "blue"
                }

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
                        popup = (
                            f"Nota: {r['NOTA']}\n"
                            f"Municipio: {r['MUNICIPIO']}\n"
                            f"Categoria: {categoria}\n"
                            f"Grupo: {r['GRUPO']}"
                        )

                        folium.Marker(
                            [
                                float(r["LATITUDE"]),
                                float(r["LONGITUDE"])
                            ],
                            popup=popup,
                            icon=folium.Icon(
                                color=cores[categoria]
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
                "Bibliotecas folium nao instaladas."
            )


with abas[7]:
    st.subheader("📈 Historico")

    if historico is None:
        st.info(
            "Envie uma analise anterior "
            "na barra lateral."
        )
    else:
        c1, c2 = st.columns(2)

        c1.metric(
            "Novas notas",
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

        st.markdown("**Notas removidas**")

        st.dataframe(
            removidas,
            hide_index=True,
            use_container_width=True
        )


# ============================================================
# EXPORTACAO
# ============================================================

st.divider()

st.subheader("📥 Exportacao Excel e KML")

st.caption(
    "Excel: inclui obras, auditoria e, quando "
    "calculadas, rotas e programacao. "
    "KML: inclui as obras filtradas "
    "com coordenadas validas."
)

parametros = {
    "Modo programacao": modo,
    "Limite por periodo": limite,
    "Capacidade por dia": capacidade_dia,
    "Periodos": periodos,
    "Raio rodoviario KM": raio,
    "Raio agrupamento KM": raio_grupo,
    "Lote obras": max_obras
}

excel_bytes = exportar_excel(
    obras,
    auditoria,
    tarefas_rotas,
    rotas,
    programacao,
    carga,
    equipes,
    novas,
    removidas,
    parametros
)

kml_bytes = exportar_kml(view)

c1, c2 = st.columns(2)

with c1:
    st.download_button(
        "📊 EXPORTAR EXCEL",
        data=excel_bytes,
        file_name=(
            "NIP_Planejamento_"
            + datetime.now().strftime("%Y%m%d_%H%M")
            + ".xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

with c2:
    st.download_button(
        "🗺️ EXPORTAR KML",
        data=kml_bytes,
        file_name=(
            "NIP_Obras_"
            + datetime.now().strftime("%Y%m%d_%H%M")
            + ".kml"
        ),
        mime="application/vnd.google-earth.kml+xml",
        use_container_width=True
    )

st.caption(
    f"Saneamento: {dados['qtd_san']} linhas | "
    f"Levantamento: {dados['qtd_lev']} linhas | "
    f"Notas unicas: {len(auditoria)}"
)
