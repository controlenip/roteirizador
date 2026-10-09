
import io
import re
import time
from concurrent.futures import ThreadPoolExecutor
import html
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

from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from sklearn.cluster import DBSCAN
from sklearn.neighbors import BallTree
from openpyxl.styles import Font, PatternFill, Alignment


# ============================================================
# CONFIGURACAO GERAL
# ============================================================

st.set_page_config(
    page_title="NIP | Análise Cruzada",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded"
)

CAT_SAN = "OBRA SANEAMENTO"
CAT_LEV = "OBRA LEVANTAMENTO"
CAT_DUP = "OBRA SANEAMENTO E LEVANTAMENTO"

CATS = [
    CAT_SAN,
    CAT_LEV,
    CAT_DUP
]

COLORS = {
    CAT_SAN: "#2563eb",
    CAT_LEV: "#16a34a",
    CAT_DUP: "#9333ea"
}

MAP_COLORS = {
    CAT_SAN: "blue",
    CAT_LEV: "green",
    CAT_DUP: "purple"
}

LIST_VALID = {
    "0",
    "EM LEVANTAMENTO",
    "CORRECAO DE LEVANTAMENTO"
}

CONTRACT_VALID = {
    "0",
    "NIP GLOBAL LTDA - EQTL MARANHAO"
}

R = 6371.0088

OSRM = "https://router.project-osrm.org"

DAY_NAMES = [
    "Segunda",
    "Terça",
    "Quarta",
    "Quinta",
    "Sexta",
    "Sábado",
    "Domingo"
]


# ============================================================
# ESTILO VISUAL
# ============================================================

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 4.5rem !important;
        padding-bottom: 2rem !important;
        max-width: 100% !important;
    }

    .nip-title {
        font-size: 2rem;
        color: #0d256c;
        font-weight: 800;
        line-height: 1.5;
        margin: 0 0 8px;
        overflow: visible;
        padding-top: 8px;
    }

    .nip-muted {
        color: #64748b;
        margin-bottom: 20px;
    }

    .nip-card {
        padding: 15px;
        border: 1px solid #e2e8f0;
        border-left: 5px solid var(--accent);
        border-radius: 10px;
        min-height: 110px;
        overflow-wrap: anywhere;
    }

    .nip-label {
        font-size: 12px;
        color: #64748b;
        font-weight: 600;
    }

    .nip-value {
        font-size: 26px;
        font-weight: 800;
        color: #172554;
    }

    .nip-small {
        font-size: 11px;
        color: #64748b;
    }

    @media(max-width:900px) {
        .block-container {
            padding-top: 3.3rem !important;
        }

        .nip-title {
            font-size: 1.45rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# NORMALIZACAO
# ============================================================

def n(v):
    if v is None:
        return ""

    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass

    texto = unicodedata.normalize(
        "NFKD",
        str(v)
    ).encode(
        "ascii",
        "ignore"
    ).decode().upper().strip()

    return re.sub(r"\s+", " ", texto)


def clean(v):
    s = n(v)

    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]

    return s


def nota(v):
    s = str(v).strip() if pd.notna(v) else ""

    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]

    if n(s) in (
        "",
        "NAN",
        "NONE",
        "NULL",
        "0"
    ):
        return ""

    return s


def col(df, variants, required=True):
    names = {
        n(c): c
        for c in df.columns
    }

    for x in variants:
        if n(x) in names:
            return names[n(x)]

    if required:
        raise ValueError(
            "Coluna ausente: "
            + " / ".join(variants)
        )

    return None


def to_num(series):
    return pd.to_numeric(
        series.astype(str).str.replace(
            ",",
            ".",
            regex=False
        ),
        errors="coerce"
    )


def valid_coords(df, lat, lon):
    la = to_num(df[lat])
    lo = to_num(df[lon])

    ok = (
        la.between(-35, 6)
        & lo.between(-75, -30)
        & la.ne(0)
        & lo.ne(0)
    )

    return la.where(ok), lo.where(ok)


# ============================================================
# IDENTIFICACAO AUTOMATICA DE ABAS
# ============================================================

def find_sheet(file, groups, prefer=""):
    xls = pd.ExcelFile(
        io.BytesIO(file.getvalue()),
        engine="openpyxl"
    )

    names = sorted(
        xls.sheet_names,
        key=lambda s: (
            0 if n(s) == n(prefer) else 1
        )
    )

    for name in names:
        for header in range(6):
            try:
                test = pd.read_excel(
                    xls,
                    sheet_name=name,
                    header=header,
                    nrows=3,
                    dtype=str
                )

                keys = {
                    n(x)
                    for x in test.columns
                }

                encontrado = all(
                    any(
                        n(v) in keys
                        for v in group
                    )
                    for group in groups
                )

                if encontrado:
                    df = pd.read_excel(
                        xls,
                        sheet_name=name,
                        header=header,
                        dtype=str
                    )

                    return df, name

            except (
                ValueError,
                TypeError,
                KeyError
            ):
                continue

    raise ValueError(
        "Não encontrei aba com as colunas exigidas. "
        "Abas: "
        + ", ".join(xls.sheet_names)
    )


# ============================================================
# DATAS
# ============================================================

def parse_date(v):
    if v is None or n(v) in (
        "",
        "NAN",
        "NONE",
        "NULL",
        "0",
        "-"
    ):
        return pd.NaT

    try:
        f = float(
            str(v).replace(",", ".")
        )

        if 20000 <= f <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.Timedelta(days=f)
            )

    except (
        ValueError,
        TypeError,
        OverflowError
    ):
        pass

    return pd.to_datetime(
        v,
        dayfirst=True,
        errors="coerce"
    )


# ============================================================
# LEITURA SANEAMENTO
# ============================================================

def load_san(file):
    if file.name.lower().endswith(".csv"):
        raw = file.getvalue()

        try:
            df = pd.read_csv(
                io.BytesIO(raw),
                sep=None,
                engine="python",
                dtype=str,
                encoding="utf-8-sig"
            )

        except UnicodeDecodeError:
            df = pd.read_csv(
                io.BytesIO(raw),
                sep=None,
                engine="python",
                dtype=str,
                encoding="latin-1"
            )

        sheet = "CSV"

    else:
        df, sheet = find_sheet(
            file,
            [
                ["NOTA"],
                [
                    "MUNICIPIO",
                    "MUNICÍPIO",
                    "CIDADE"
                ]
            ],
            "Clientes Existentes"
        )

    k = col(df, ["NOTA"])

    city = col(
        df,
        [
            "MUNICIPIO",
            "MUNICÍPIO",
            "CIDADE"
        ]
    )

    reg = col(
        df,
        ["REGIONAL"],
        False
    )

    lat = col(
        df,
        [
            "LATITUDE PROJETO",
            "LATITUDE",
            "LATITUDE CAMPO"
        ]
    )

    lon = col(
        df,
        [
            "LONGITUDE PROJETO",
            "LONGITUDE",
            "LONGITUDE CAMPO"
        ]
    )

    df["CHAVE"] = df[k].map(nota)

    df["CIDADE_OBRA"] = (
        df[city].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[reg].fillna("")
        if reg else ""
    )

    df["LAT_OBRA"], df["LON_OBRA"] = (
        valid_coords(
            df,
            lat,
            lon
        )
    )

    return df, sheet


# ============================================================
# LEITURA LEVANTAMENTO
# ============================================================

def load_lev(file):
    df, sheet = find_sheet(
        file,
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
        "NOTAS"
    )

    k = col(df, ["PROTOCOLO"])

    sap = col(
        df,
        [
            "STATUS SAP",
            "STATUS_SAP"
        ]
    )

    lst = col(
        df,
        [
            "STATUS LIST",
            "STATUS_LIST"
        ]
    )

    ctr = col(
        df,
        ["CONTRATO"]
    )

    city = col(
        df,
        [
            "MUNICIPIO",
            "MUNICÍPIO",
            "CIDADE"
        ]
    )

    reg = col(
        df,
        ["REGIONAL"],
        False
    )

    lat = col(df, ["LATITUDE"])
    lon = col(df, ["LONGITUDE"])

    m = next(
        (
            c for c in df.columns
            if "ORCAMENTO MODULAR" in n(c)
        ),
        None
    )

    nn = next(
        (
            c for c in df.columns
            if "PLA ALVOS" in n(c)
        ),
        None
    )

    if m is None or nn is None:
        raise ValueError(
            "Não localizei ORÇAMENTO MODULAR "
            "e PLA ALVOS."
        )

    if len(df.columns) < 3:
        raise ValueError(
            "A aba necessita das colunas físicas B e C."
        )

    b = df.columns[1]
    c = df.columns[2]

    priority = col(
        df,
        ["PRIORIDADE"],
        False
    )

    opening = col(
        df,
        [
            "DATA ABERTURA",
            "DATA DE ABERTURA"
        ],
        False
    )

    df["CHAVE"] = df[k].map(nota)

    df["CIDADE_OBRA"] = (
        df[city].fillna("")
    )

    df["REGIONAL_OBRA"] = (
        df[reg].fillna("")
        if reg else ""
    )

    df["LAT_OBRA"], df["LON_OBRA"] = (
        valid_coords(
            df,
            lat,
            lon
        )
    )

    df["SAP"] = df[sap].map(clean)
    df["LIST"] = df[lst].map(clean)

    df["CONTRATO_N"] = (
        df[ctr].map(clean)
    )

    df["M"] = df[m].map(clean)
    df["N"] = df[nn].map(clean)

    df["RESP_CAMPO"] = (
        df[b]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["DATA_CAMPO_DT"] = (
        df[c].map(parse_date)
    )

    df["EM_CAMPO"] = (
        df["RESP_CAMPO"].map(
            lambda x: n(x) not in (
                "",
                "SEM LEVANTADOR",
                "NAN",
                "NONE",
                "NULL",
                "0",
                "-"
            )
        )
        & df["DATA_CAMPO_DT"].notna()
    )

    df["PRIORIDADE_RAW"] = (
        df[priority].fillna("")
        if priority else ""
    )

    df["ABERTURA_RAW"] = (
        df[opening].fillna("")
        if opening else ""
    )

    def reasons(r):
        result = []

        if r["SAP"] in (
            "FINL",
            "CANC"
        ):
            result.append(
                "SAP " + r["SAP"]
            )

        if r["LIST"] not in LIST_VALID:
            result.append(
                "LIST "
                + (
                    r["LIST"]
                    or "VAZIO"
                )
            )

        if r["CONTRATO_N"] not in CONTRACT_VALID:
            result.append(
                "CONTRATO NÃO PERMITIDO"
            )

        if r["M"] != "0":
            result.append(
                "ORÇAMENTO MODULAR NÃO ZERO"
            )

        if r["N"] != "0":
            result.append(
                "PLA ALVOS NÃO ZERO"
            )

        if r["EM_CAMPO"]:
            result.append(
                "JÁ EM CAMPO (B+C)"
            )

        return result

    df["MOTIVOS"] = df.apply(
        reasons,
        axis=1
    )

    return df, {
        "aba": sheet,
        "B": b,
        "C": c,
        "linhas_campo": int(
            df["EM_CAMPO"].sum()
        )
    }


# ============================================================
# LEITURA EQUIPES
# ============================================================

def load_teams(file):
    xls = pd.ExcelFile(
        io.BytesIO(file.getvalue()),
        engine="openpyxl"
    )

    parts = []

    for kind, sheet in [
        (
            "SANEAMENTO",
            "SANEAMENTO"
        ),
        (
            "LEVANTAMENTO",
            "LEVANTADORES"
        )
    ]:
        name = next(
            (
                x for x in xls.sheet_names
                if n(x) == sheet
            ),
            None
        )

        if name is None:
            raise ValueError(
                "Aba de equipe ausente: "
                + sheet
            )

        df = pd.read_excel(
            xls,
            sheet_name=name,
            dtype=str
        )

        namecol = col(
            df,
            [
                "NOME",
                "EQUIPE",
                "NOME_COLAB"
            ]
        )

        city = col(
            df,
            [
                "CIDADES",
                "CIDADE",
                "MUNICIPIO"
            ]
        )

        lat = col(
            df,
            [
                "LATITUDE",
                "LAT"
            ]
        )

        lon = col(
            df,
            [
                "LONGITUDE",
                "LON"
            ]
        )

        la, lo = valid_coords(
            df,
            lat,
            lon
        )

        parts.append(
            pd.DataFrame({
                "TIPO_EQUIPE": kind,
                "EQUIPE": df[namecol],
                "CIDADE_BASE": df[city],
                "LAT_EQUIPE": la,
                "LON_EQUIPE": lo
            })
        )

    data = pd.concat(
        parts,
        ignore_index=True
    ).dropna(
        subset=[
            "EQUIPE",
            "LAT_EQUIPE",
            "LON_EQUIPE"
        ]
    )

    data = data[
        data["EQUIPE"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    data["EQUIPE"] = (
        data["EQUIPE"]
        .astype(str)
        .str.strip()
    )

    data["CIDADE_BASE"] = (
        data["CIDADE_BASE"].fillna("")
    )

    data["ID_EQUIPE"] = (
        data["TIPO_EQUIPE"]
        + "|"
        + data["EQUIPE"]
        + "|"
        + data["CIDADE_BASE"]
    )

    return data.drop_duplicates(
        "ID_EQUIPE"
    ).reset_index(drop=True)


# ============================================================
# CRUZAMENTO DAS BASES
# ============================================================

def consolidate(san, lev):
    gs = dict(
        tuple(
            san[
                san["CHAVE"].ne("")
            ].groupby(
                "CHAVE",
                sort=False
            )
        )
    )

    gl = dict(
        tuple(
            lev[
                lev["CHAVE"].ne("")
            ].groupby(
                "CHAVE",
                sort=False
            )
        )
    )

    rows = []

    for key in dict.fromkeys(
        [*gs, *gl]
    ):
        a = gs.get(key)
        b = gl.get(key)

        if a is not None:
            valid_a = (
                a["LAT_OBRA"].notna()
                & a["LON_OBRA"].notna()
            )

            sa = (
                a.loc[valid_a].iloc[0]
                if valid_a.any()
                else a.iloc[0]
            )
        else:
            sa = None

        if b is not None:
            valid_b = (
                b["LAT_OBRA"].notna()
                & b["LON_OBRA"].notna()
            )

            sb = (
                b.loc[valid_b].iloc[0]
                if valid_b.any()
                else b.iloc[0]
            )
        else:
            sb = None

        source = (
            sb
            if (
                sb is not None
                and pd.notna(
                    sb["LAT_OBRA"]
                )
            )
            else sa
            if sa is not None
            else sb
        )

        reasons = (
            list(
                dict.fromkeys(
                    v
                    for sub in b["MOTIVOS"]
                    for v in sub
                )
            )
            if b is not None
            else []
        )

        status_saps = (
            set(b["SAP"])
            if b is not None
            else set()
        )

        on_field = (
            b[b["EM_CAMPO"]]
            if b is not None
            else pd.DataFrame()
        )

        rows.append({
            "NOTA": key,
            "LISTA": (
                CAT_DUP
                if (
                    a is not None
                    and b is not None
                )
                else CAT_SAN
                if a is not None
                else CAT_LEV
            ),
            "MUNICIPIO": source[
                "CIDADE_OBRA"
            ],
            "REGIONAL": source[
                "REGIONAL_OBRA"
            ],
            "LATITUDE": source[
                "LAT_OBRA"
            ],
            "LONGITUDE": source[
                "LON_OBRA"
            ],
            "DUPLICADA": (
                "SIM"
                if (
                    a is not None
                    and b is not None
                )
                else "NÃO"
            ),
            "PENDENTE_CONTAGEM": (
                "NÃO"
                if reasons
                else "SIM"
            ),
            "MOTIVO_EXCLUSAO": (
                " | ".join(reasons)
                if reasons
                else "-"
            ),
            "SAP_FINL": (
                "SIM"
                if "FINL" in status_saps
                else "NÃO"
            ),
            "SAP_CANC": (
                "SIM"
                if "CANC" in status_saps
                else "NÃO"
            ),
            "JA_EM_CAMPO": (
                "SIM"
                if len(on_field)
                else "NÃO"
            ),
            "EQUIPE_CAMPO": (
                " | ".join(
                    on_field[
                        "RESP_CAMPO"
                    ].astype(str).unique()
                )
                if len(on_field)
                else ""
            ),
            "DATA_CAMPO": (
                " | ".join(
                    pd.Timestamp(x).strftime(
                        "%d/%m/%Y"
                    )
                    for x in on_field[
                        "DATA_CAMPO_DT"
                    ].dropna().unique()
                )
                if len(on_field)
                else ""
            ),
            "STATUS_SAP": (
                " | ".join(
                    dict.fromkeys(
                        b["SAP"]
                    )
                )
                if b is not None
                else ""
            ),
            "STATUS_LIST": (
                " | ".join(
                    dict.fromkeys(
                        b["LIST"]
                    )
                )
                if b is not None
                else ""
            ),
            "PRIORIDADE_ORIGINAL": (
                sb["PRIORIDADE_RAW"]
                if sb is not None
                else ""
            ),
            "DATA_ABERTURA": (
                sb["ABERTURA_RAW"]
                if sb is not None
                else ""
            ),
            "OCORRENCIAS_SANEAMENTO": (
                len(a)
                if a is not None
                else 0
            ),
            "OCORRENCIAS_LEVANTAMENTO": (
                len(b)
                if b is not None
                else 0
            ),
            "MUNICIPIO_SANEAMENTO": (
                sa["CIDADE_OBRA"]
                if sa is not None
                else ""
            ),
            "MUNICIPIO_LEVANTAMENTO": (
                sb["CIDADE_OBRA"]
                if sb is not None
                else ""
            )
        })

    return pd.DataFrame(rows)


# ============================================================
# TIPOS DE ATIVIDADE
# ============================================================

def activity(cat):
    if cat == CAT_DUP:
        return [
            "SANEAMENTO",
            "LEVANTAMENTO"
        ]

    if cat == CAT_SAN:
        return ["SANEAMENTO"]

    return ["LEVANTAMENTO"]


# ============================================================
# HAVERSINE
# ============================================================

def haversine(a, b, c, d):
    x = np.radians(a)
    y = np.radians(b)
    z = np.radians(c)
    w = np.radians(d)

    t = (
        np.sin(
            (z - x) / 2
        ) ** 2
        + np.cos(x)
        * np.cos(z)
        * np.sin(
            (w - y) / 2
        ) ** 2
    )

    return (
        2
        * R
        * np.arcsin(
            np.sqrt(
                np.clip(t, 0, 1)
            )
        )
    )


# ============================================================
# PRIORIDADE
# ============================================================

def priority(df, days):
    df = df.copy()

    dates = pd.to_datetime(
        df["DATA_ABERTURA"],
        errors="coerce",
        dayfirst=True
    )

    age = (
        pd.Timestamp.today().normalize()
        - dates
    ).dt.days

    df["DIAS_ABERTURA"] = age

    resultados = []

    for p, d in zip(
        df["PRIORIDADE_ORIGINAL"],
        age
    ):
        if (
            "ALTA" in n(p)
            or "URGENTE" in n(p)
            or (
                pd.notna(d)
                and d >= 4 * days
            )
        ):
            resultados.append(
                "ALTA"
            )

        elif (
            pd.notna(d)
            and d >= days
        ):
            resultados.append(
                "MEDIA"
            )

        else:
            resultados.append(
                "BAIXA"
            )

    df["PRIORIDADE_PLANEJAMENTO"] = (
        resultados
    )

    return df


# ============================================================
# CANDIDATOS MAIS PROXIMOS
# ============================================================

@st.cache_data(
    show_spinner=False
)
def team_candidates(
    obras,
    equipes,
    k
):
    rows = []
    labels = {}

    for kind in [
        "SANEAMENTO",
        "LEVANTAMENTO"
    ]:
        team = equipes[
            equipes[
                "TIPO_EQUIPE"
            ].eq(kind)
        ].reset_index(drop=True)

        eligible = obras[
            obras["LISTA"].map(
                lambda c: (
                    kind in activity(c)
                )
            )
        ].dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        if team.empty or eligible.empty:
            continue

        tree = BallTree(
            np.radians(
                team[
                    [
                        "LAT_EQUIPE",
                        "LON_EQUIPE"
                    ]
                ].to_numpy(float)
            ),
            metric="haversine"
        )

        dist, ids = tree.query(
            np.radians(
                eligible[
                    [
                        "LATITUDE",
                        "LONGITUDE"
                    ]
                ].to_numpy(float)
            ),
            k=min(
                k,
                len(team)
            )
        )

        for i, (_, work) in enumerate(
            eligible.iterrows()
        ):
            txt = []

            for j in range(
                ids.shape[1]
            ):
                eq = team.iloc[
                    ids[i, j]
                ]

                km = round(
                    float(
                        dist[i, j] * R
                    ),
                    2
                )

                rows.append({
                    "NOTA": work["NOTA"],
                    "ATIVIDADE": kind,
                    "ID_EQUIPE": eq["ID_EQUIPE"],
                    "EQUIPE": eq["EQUIPE"],
                    "CIDADE_BASE": eq["CIDADE_BASE"],
                    "LAT_EQUIPE": eq["LAT_EQUIPE"],
                    "LON_EQUIPE": eq["LON_EQUIPE"],
                    "LAT_OBRA": work["LATITUDE"],
                    "LON_OBRA": work["LONGITUDE"],
                    "KM_RETA": km
                })

                txt.append(
                    f"{eq['EQUIPE']} "
                    f"({eq['CIDADE_BASE']}) "
                    f"- {km:.1f} km (linha reta)"
                )

            labels[
                (
                    work["NOTA"],
                    kind
                )
            ] = " | ".join(txt)

    out = obras.copy()

    for kind in [
        "SANEAMENTO",
        "LEVANTAMENTO"
    ]:
        out[
            "EQUIPES_" + kind
        ] = out.apply(
            lambda r: labels.get(
                (
                    r["NOTA"],
                    kind
                ),
                (
                    "SEM COORDENADAS OU EQUIPE"
                    if kind in activity(
                        r["LISTA"]
                    )
                    else "NÃO APLICÁVEL"
                )
            ),
            axis=1
        )

    columns = [
        "NOTA",
        "ATIVIDADE",
        "ID_EQUIPE",
        "EQUIPE",
        "CIDADE_BASE",
        "LAT_EQUIPE",
        "LON_EQUIPE",
        "LAT_OBRA",
        "LON_OBRA",
        "KM_RETA"
    ]

    return out, pd.DataFrame(
        rows,
        columns=columns
    )


# ============================================================
# OSRM COM CACHE E TIMEOUT
# ============================================================

@st.cache_data(
    ttl=86400,
    max_entries=10000,
    show_spinner=False
)
def osrm_route(
    lat1,
    lon1,
    lat2,
    lon2,
    server,
    geometry=False
):
    if (
        lat1 == lat2
        and lon1 == lon2
    ):
        return {
            "ok": True,
            "km": 0.0,
            "min": 0.0,
            "geometry": [
                [lon1, lat1],
                [lon2, lat2]
            ]
        }

    url = (
        server.rstrip("/")
        + "/route/v1/driving/"
        + f"{lon1:.6f},{lat1:.6f};"
        + f"{lon2:.6f},{lat2:.6f}"
    )

    try:
        response = requests.get(
            url,
            params={
                "overview": (
                    "full"
                    if geometry
                    else "false"
                ),
                "geometries": "geojson",
                "steps": "false"
            },
            headers={
                "User-Agent": "NIP-Planejamento/1.0"
            },
            timeout=(2, 4)
        )

        response.raise_for_status()

        data = response.json()

        if (
            data.get("code") != "Ok"
            or not data.get("routes")
        ):
            return {
                "ok": False
            }

        route = data["routes"][0]

        geom = (
            route.get(
                "geometry",
                {}
            ).get(
                "coordinates",
                []
            )
            if geometry
            else []
        )

        return {
            "ok": True,
            "km": round(
                float(
                    route["distance"]
                ) / 1000,
                2
            ),
            "min": round(
                float(
                    route["duration"]
                ) / 60,
                1
            ),
            "geometry": geom
        }

    except (
        requests.RequestException,
        ValueError,
        TypeError,
        KeyError
    ):
        return {
            "ok": False
        }


# ============================================================
# PREPARACAO DE ROTAS
# ============================================================

def initial_routes(
    candidates,
    max_km
):
    data = candidates.copy()

    data["KM_RODOVIARIO"] = np.nan
    data["TEMPO_RODOVIARIO_MIN"] = np.nan

    data["ROTA_STATUS"] = (
        "NÃO CONSULTADA"
    )

    data["KM_ALOCACAO"] = np.nan

    data.loc[
        data["KM_RETA"].gt(max_km),
        "ROTA_STATUS"
    ] = "FORA RAIO GEOGRÁFICO"

    return data


# ============================================================
# CONSULTAS OSRM POR LOTES
# ============================================================

def compute_route_batch(
    routes,
    start,
    batch,
    server,
    placeholder=None
):
    df = routes.copy()

    eligible = df.index[
        df["ROTA_STATUS"].eq(
            "NÃO CONSULTADA"
        )
    ].tolist()

    end = min(
        start + batch,
        len(eligible)
    )

    t0 = time.monotonic()

    for pos in range(start, end):
        idx = eligible[pos]

        r = df.loc[idx]

        answer = osrm_route(
            round(
                float(
                    r["LAT_EQUIPE"]
                ),
                6
            ),
            round(
                float(
                    r["LON_EQUIPE"]
                ),
                6
            ),
            round(
                float(
                    r["LAT_OBRA"]
                ),
                6
            ),
            round(
                float(
                    r["LON_OBRA"]
                ),
                6
            ),
            server,
            False
        )

        if answer["ok"]:
            df.at[
                idx,
                "KM_RODOVIARIO"
            ] = answer["km"]

            df.at[
                idx,
                "TEMPO_RODOVIARIO_MIN"
            ] = answer["min"]

            df.at[
                idx,
                "ROTA_STATUS"
            ] = "OSRM CONFIRMADA"

        else:
            df.at[
                idx,
                "ROTA_STATUS"
            ] = "SEM ROTA OSRM"

        if placeholder is not None:
            placeholder.text(
                f"Lote: "
                f"{pos - start + 1}/{end - start}"
                f" | Tempo: "
                f"{time.monotonic() - t0:.1f}s"
            )

    return (
        df,
        end - start,
        time.monotonic() - t0
    )


# ============================================================
# GERAR TAREFAS
# ============================================================

def tasks_for(obras):
    rows = []

    for _, r in obras.iterrows():
        for kind in activity(
            r["LISTA"]
        ):
            rows.append({
                **r.to_dict(),
                "ATIVIDADE": kind
            })

    return pd.DataFrame(rows)


# ============================================================
# DATAS PARA PROGRAMACAO
# ============================================================

def schedule_days(
    start,
    mode,
    periods
):
    start = pd.Timestamp(start)

    if mode == "POR DIA":
        end = (
            start
            + pd.Timedelta(
                days=periods * 3 + 14
            )
        )

    elif mode == "POR SEMANA":
        end = (
            start
            + pd.Timedelta(
                weeks=periods
            )
        )

    elif mode == "POR MÊS":
        end = (
            start
            + pd.DateOffset(
                months=periods
            )
            - pd.Timedelta(days=1)
        )

    else:
        end = (
            start
            + pd.Timedelta(days=365)
        )

    days = [
        v.date()
        for v in pd.date_range(
            start,
            end,
            freq="D"
        )
        if v.weekday() < 5
    ]

    if mode == "POR DIA":
        return days[:periods]

    return days


def period_id(date, mode):
    iso = date.isocalendar()

    if mode == "POR SEMANA":
        return (
            f"{iso.year}-S"
            f"{iso.week:02d}"
        )

    if mode == "POR MÊS":
        return date.strftime("%Y-%m")

    return date.isoformat()


# ============================================================
# ALOCACAO COM DISTANCIA RODOVIARIA
# ============================================================

def schedule_tasks(
    tasks,
    routes,
    teams,
    capacities,
    mode,
    periods,
    start,
    period_cap,
    default_daily,
    limit_km,
    route_only
):
    cols = [
        "EQUIPE_PROGRAMADA",
        "ID_EQUIPE",
        "DATA_PROGRAMADA",
        "PERIODO",
        "STATUS_PROGRAMACAO"
    ]

    if tasks.empty:
        return (
            pd.DataFrame(
                columns=list(tasks.columns) + cols
            ),
            pd.DataFrame()
        )

    routes = routes.copy()

    routes["KM_ALOCACAO"] = (
        routes["KM_RODOVIARIO"]
    )

    route_map = defaultdict(list)

    for _, c in routes.iterrows():
        if (
            c["ROTA_STATUS"]
            == "OSRM CONFIRMADA"
            and pd.notna(
                c["KM_RODOVIARIO"]
            )
            and c["KM_RODOVIARIO"] <= limit_km
        ):
            route_map[
                (
                    c["NOTA"],
                    c["ATIVIDADE"]
                )
            ].append(
                c.to_dict()
            )

    capmap = {
        str(r["ID_EQUIPE"]): int(
            r["CAPACIDADE_DIA"]
        )
        for _, r in capacities.iterrows()
    }

    days = schedule_days(
        start,
        mode,
        periods
    )

    used_day = defaultdict(int)
    used_period = defaultdict(int)

    order = {
        "ALTA": 0,
        "MEDIA": 1,
        "BAIXA": 2
    }

    ordered = tasks.assign(
        _ord=tasks[
            "PRIORIDADE_PLANEJAMENTO"
        ].map(order).fillna(3)
    ).sort_values(
        [
            "_ord",
            "MUNICIPIO",
            "NOTA"
        ]
    )

    out = []

    for _, task in ordered.iterrows():
        candidates = route_map.get(
            (
                task["NOTA"],
                task["ATIVIDADE"]
            ),
            []
        )

        options = []

        for c in candidates:
            for day in days:
                period = period_id(
                    day,
                    mode
                )

                key = (
                    c["ID_EQUIPE"],
                    day
                )

                pk = (
                    c["ID_EQUIPE"],
                    period
                )

                cap = int(
                    capmap.get(
                        c["ID_EQUIPE"],
                        default_daily
                    )
                )

                if (
                    used_day[key] < cap
                    and used_period[pk] < period_cap
                ):
                    options.append(
                        (
                            used_period[pk],
                            used_day[key],
                            day,
                            float(
                                c["KM_RODOVIARIO"]
                            ),
                            c,
                            period
                        )
                    )

                    break

        r = task.drop(
            labels=["_ord"]
        ).to_dict()

        if options:
            (
                _,
                _,
                day,
                km,
                c,
                period
            ) = min(
                options,
                key=lambda v: v[:4]
            )

            used_day[
                (
                    c["ID_EQUIPE"],
                    day
                )
            ] += 1

            used_period[
                (
                    c["ID_EQUIPE"],
                    period
                )
            ] += 1

            r.update({
                "EQUIPE_PROGRAMADA": c["EQUIPE"],
                "ID_EQUIPE": c["ID_EQUIPE"],
                "DATA_PROGRAMADA": day.isoformat(),
                "DIA_SEMANA": DAY_NAMES[
                    day.weekday()
                ],
                "PERIODO": period,
                "KM_RODOVIARIO": km,
                "TEMPO_RODOVIARIO_MIN": c[
                    "TEMPO_RODOVIARIO_MIN"
                ],
                "STATUS_PROGRAMACAO": "SUGESTÃO"
            })

        else:
            r.update({
                "EQUIPE_PROGRAMADA": "NÃO ALOCADA",
                "ID_EQUIPE": "",
                "DATA_PROGRAMADA": "",
                "DIA_SEMANA": "",
                "PERIODO": "",
                "KM_RODOVIARIO": np.nan,
                "TEMPO_RODOVIARIO_MIN": np.nan,
                "STATUS_PROGRAMACAO": (
                    "SEM ROTA OU CAPACIDADE"
                )
            })

        out.append(r)

    plan = pd.DataFrame(out)

    load = (
        plan[
            plan[
                "STATUS_PROGRAMACAO"
            ].eq("SUGESTÃO")
        ]
        .groupby(
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

    return plan, load


# ============================================================
# SUPERPONTOS DBSCAN
# ============================================================

def merge_superpoints(
    plan,
    meters
):
    if plan.empty:
        return (
            pd.DataFrame(),
            0
        )

    valid = plan.dropna(
        subset=[
            "LATITUDE",
            "LONGITUDE"
        ]
    ).copy()

    invalid = plan[
        plan["LATITUDE"].isna()
        | plan["LONGITUDE"].isna()
    ].copy()

    if valid.empty:
        return invalid, 0

    grouped = []

    # Agrupamento independente por equipe,
    # data e atividade.
    for _, g in valid.groupby(
        [
            "ID_EQUIPE",
            "DATA_PROGRAMADA",
            "ATIVIDADE"
        ],
        dropna=False,
        sort=False
    ):
        coords = np.radians(
            g[
                [
                    "LATITUDE",
                    "LONGITUDE"
                ]
            ].to_numpy(float)
        )

        labels = DBSCAN(
            eps=meters / 6371000,
            min_samples=1,
            metric="haversine",
            algorithm="ball_tree"
        ).fit(coords).labels_

        clustered = g.assign(
            _cluster=labels
        )

        for _, cluster in clustered.groupby(
            "_cluster",
            sort=False
        ):
            r = cluster.iloc[0].drop(
                labels=["_cluster"]
            ).to_dict()

            original = cluster.drop(
                columns="_cluster"
            ).to_dict("records")

            r["_ORIGINAL_ROWS"] = original

            r["LATITUDE"] = (
                cluster["LATITUDE"].mean()
            )

            r["LONGITUDE"] = (
                cluster["LONGITUDE"].mean()
            )

            r["SUPER_PONTO"] = (
                f"SIM ({len(cluster)} un.)"
                if len(cluster) > 1
                else "NÃO"
            )

            r["NOTAS_AGRUPADAS"] = (
                " | ".join(
                    cluster[
                        "NOTA"
                    ].astype(str)
                )
            )

            grouped.append(r)

    for _, r in invalid.iterrows():
        grouped.append({
            **r.to_dict(),
            "SUPER_PONTO": "NÃO",
            "NOTAS_AGRUPADAS": r["NOTA"],
            "_ORIGINAL_ROWS": [
                r.to_dict()
            ]
        })

    return (
        pd.DataFrame(grouped),
        len(plan) - len(grouped)
    )


# ============================================================
# AUDITORIA DE INCONSISTENCIAS
# ============================================================

def anomalies(
    san,
    lev,
    consolidated,
    teams
):
    issues = []

    for origin, df in [
        ("SANEAMENTO", san),
        ("LEVANTAMENTO", lev)
    ]:
        counts = (
            df.loc[
                df["CHAVE"].ne(""),
                "CHAVE"
            ].value_counts()
        )

        repeated = counts[
            counts > 1
        ]

        for key, count in repeated.items():
            issues.append({
                "NOTA": key,
                "TIPO": (
                    "REPETIDA NA BASE "
                    + origin
                ),
                "DETALHE": f"{count} linhas"
            })

    for _, r in consolidated.iterrows():
        if (
            pd.isna(r["LATITUDE"])
            or pd.isna(r["LONGITUDE"])
        ):
            issues.append({
                "NOTA": r["NOTA"],
                "TIPO": "COORDENADA AUSENTE",
                "DETALHE": r["MUNICIPIO"]
            })

        municipio_san = n(
            r["MUNICIPIO_SANEAMENTO"]
        )

        municipio_lev = n(
            r["MUNICIPIO_LEVANTAMENTO"]
        )

        if (
            municipio_san
            and municipio_lev
            and municipio_san != municipio_lev
        ):
            issues.append({
                "NOTA": r["NOTA"],
                "TIPO": (
                    "DIVERGÊNCIA DE MUNICÍPIO"
                ),
                "DETALHE": (
                    f"{r['MUNICIPIO_SANEAMENTO']} / "
                    f"{r['MUNICIPIO_LEVANTAMENTO']}"
                )
            })

        if (
            r["SAP_FINL"] == "SIM"
            and r["SAP_CANC"] == "SIM"
        ):
            issues.append({
                "NOTA": r["NOTA"],
                "TIPO": (
                    "STATUS SAP CONFLITANTE"
                ),
                "DETALHE": (
                    "FINL e CANC em linhas diferentes"
                )
            })

    return pd.DataFrame(
        issues,
        columns=[
            "NOTA",
            "TIPO",
            "DETALHE"
        ]
    )


# ============================================================
# COMPARACAO HISTORICA
# ============================================================

def compare_history(
    current,
    file
):
    sheets = pd.read_excel(
        io.BytesIO(file.getvalue()),
        sheet_name=None,
        engine="openpyxl",
        dtype=str
    )

    previous = next(
        (
            sheets[x]
            for x in [
                "CONSOLIDADO",
                "AUDITORIA COMPLETA",
                "AUDITORIA"
            ]
            if x in sheets
        ),
        None
    )

    if (
        previous is None
        or "NOTA" not in previous.columns
    ):
        raise ValueError(
            "Histórico deve conter CONSOLIDADO "
            "ou AUDITORIA e coluna NOTA."
        )

    if (
        "PENDENTE_CONTAGEM"
        in previous.columns
    ):
        previous = previous[
            previous[
                "PENDENTE_CONTAGEM"
            ].eq("SIM")
        ]

    old = set(
        previous["NOTA"].map(nota)
    ) - {""}

    new = set(
        current["NOTA"].map(nota)
    ) - {""}

    return (
        pd.DataFrame({
            "NOTA": sorted(
                new - old
            )
        }),
        pd.DataFrame({
            "NOTA": sorted(
                old - new
            )
        })
    )


# ============================================================
# EXCEL
# ============================================================

def safe_excel(df):
    df = df.copy()

    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(
                lambda x: (
                    " | ".join(
                        map(str, x)
                    )
                    if isinstance(x, list)
                    else str(x)
                    if isinstance(x, dict)
                    else x
                )
            )

    return df.drop(
        columns=[
            "_ORIGINAL_ROWS",
            "GEOMETRIA"
        ],
        errors="ignore"
    )


def excel_bytes(sheets):
    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:

        for name, frame in sheets.items():
            df = safe_excel(frame)

            sheet = name[:31]

            df.to_excel(
                writer,
                sheet_name=sheet,
                index=False
            )

            ws = writer.sheets[sheet]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            ws.sheet_view.showGridLines = False

            for cell in ws[1]:
                cell.fill = PatternFill(
                    "solid",
                    fgColor="002060"
                )

                cell.font = Font(
                    color="FFFFFF",
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

            for row in ws.iter_rows(
                min_row=2
            ):
                if (
                    "SUPER_PONTO" in df.columns
                    and str(
                        row[
                            df.columns.get_loc(
                                "SUPER_PONTO"
                            )
                        ].value
                    ).startswith("SIM")
                ):
                    for cell in row:
                        cell.fill = PatternFill(
                            "solid",
                            fgColor="FCE4D6"
                        )

            for col_cells in ws.columns:
                letter = (
                    col_cells[0].column_letter
                )

                lengths = [
                    len(str(c.value or ""))
                    for c in list(
                        col_cells
                    )[:100]
                ]

                ws.column_dimensions[
                    letter
                ].width = min(
                    55,
                    max(
                        13,
                        max(
                            lengths,
                            default=10
                        ) + 2
                    )
                )

    return out.getvalue()


# ============================================================
# POPUP DOS MARCADORES
# ============================================================

def popup(r):
    """Cartao para Folium e Google Earth, no estilo de referencia do NIP.

    Mantem equipes candidatas e informa que a distancia e em linha reta.
    """
    def readable(value):
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
        return str(value).strip()

    def esc(value):
        return html.escape(readable(value), quote=True)

    categoria = readable(r.get("LISTA", ""))
    cor = COLORS.get(categoria, "#9333ea")
    numero = readable(r.get("NOTA", ""))

    # O titulo do Placemark continua "Nota XXXXX"; abaixo reproduzimos
    # o cabecalho para uso tanto no Google Earth quanto no Folium.
    itens = [
        ("Nota", r.get("NOTA", "")),
        ("Município", r.get("MUNICIPIO", "")),
        ("Tipo", r.get("LISTA", "")),
    ]

    for chave, titulo in [
        ("EQUIPES_SANEAMENTO", "Equipes SANEAMENTO"),
        ("EQUIPES_LEVANTAMENTO", "Equipes LEVANTAMENTO"),
    ]:
        valor = readable(r.get(chave, ""))
        if valor and n(valor) not in ("NAO APLICAVEL", "NAN"):
            itens.append((titulo, valor))

    for chave, titulo in [
        ("PRIORIDADE_PLANEJAMENTO", "Prioridade"),
        ("EQUIPE_PROGRAMADA", "Equipe programada"),
        ("DATA_PROGRAMADA", "Data programada"),
        ("TIPO_OBRA", "Classificação da obra"),
        ("BASE_ATRIBUIDA", "Equipe mais próxima"),
        ("ATIVIDADE_EQUIPE", "Atividade da equipe"),
        ("DISTANCIA_EQUIPE_KM", "Distância da equipe (km, linha reta)"),
        ("ORDEM", "Ordem da visita"),
        ("SUPER_PONTO", "Superponto"),
        ("NOTAS_AGRUPADAS", "Notas agrupadas"),
    ]:
        valor = readable(r.get(chave, ""))
        if valor and n(valor) not in ("NAN", "NAO APLICAVEL"):
            itens.append((titulo, valor))

    linhas = "".join(
        '<tr>'
        '<td style="padding:5px 7px 5px 3px;vertical-align:top;'
        'font-size:11px;font-weight:700;white-space:nowrap;width:94px;">'
        f'{esc(titulo)}</td>'
        '<td style="padding:5px 3px;vertical-align:top;'
        'font-size:11px;line-height:1.4;overflow-wrap:anywhere;">'
        f'{esc(valor)}</td>'
        '</tr>'
        for titulo, valor in itens
    )

    links = ""
    try:
        lat = float(r.get("LATITUDE"))
        lon = float(r.get("LONGITUDE"))
        if np.isfinite(lat) and np.isfinite(lon):
            destino = (
                'https://www.google.com/maps/dir/?api=1&destination='
                f'{lat:.7f}%2C{lon:.7f}'
            )
            origem = (
                'https://www.google.com/maps/dir/?api=1&origin='
                f'{lat:.7f}%2C{lon:.7f}'
            )
            links = (
                '<div style="font-size:11px;margin-top:12px;">'
                'Rotas: '
                f'<a href="{destino}" target="_blank">Até aqui</a>'
                ' - '
                f'<a href="{origem}" target="_blank">Daqui</a>'
                '</div>'
            )
    except (TypeError, ValueError):
        pass

    return (
        '<div style="background:#fff;color:#171717;'
        'font-family:Arial,Helvetica,sans-serif;font-size:11px;'
        'padding:4px 3px 6px;max-width:440px;min-width:260px;">'
        '<div style="font-weight:700;font-size:12px;'
        'margin:0 0 10px;">Nota '
        + esc(numero)
        + '</div>'
        + f'<div style="background:{cor};color:white;'
          'font-weight:700;font-size:11px;padding:8px 9px;'
          'margin:0 0 3px;">INFORMAÇÕES DA OBRA</div>'
        + '<table style="width:100%;border-collapse:collapse;'
          'table-layout:auto;">'
        + linhas
        + '</table>'
        + '<div style="font-size:10px;color:#697386;'
          'margin-top:7px;line-height:1.3;">'
          'Distâncias das equipes: linha reta (não OSRM).</div>'
        + links
        + '</div>'
    )


# ============================================================
# GERAR KML
# ============================================================

def kml_bytes(
    works,
    superpoints=None,
    paths=None,
    name="NIP - Obras"
):
    ns = (
        "http://www.opengis.net/kml/2.2"
    )

    ET.register_namespace("", ns)

    def add(parent, tag, value=None):
        x = ET.SubElement(
            parent,
            f"{{{ns}}}{tag}"
        )

        x.text = (
            str(value)
            if value is not None
            else None
        )

        return x

    root = ET.Element(
        f"{{{ns}}}kml"
    )

    doc = add(
        root,
        "Document"
    )

    add(
        doc,
        "name",
        name
    )

    icons = {
        CAT_SAN: "blu-blank.png",
        CAT_LEV: "grn-blank.png",
        CAT_DUP: "purple-blank.png"
    }

    for cat in CATS:
        style = add(
            doc,
            "Style"
        )

        style.set(
            "id",
            "cat" + str(
                CATS.index(cat)
            )
        )

        ic = add(
            add(
                style,
                "IconStyle"
            ),
            "Icon"
        )

        add(
            ic,
            "href",
            (
                "http://maps.google.com/mapfiles/"
                "kml/paddle/"
                + icons[cat]
            )
        )

        folder = add(
            doc,
            "Folder"
        )

        add(
            folder,
            "name",
            cat
        )

        if works.empty:
            continue

        subset = works[
            works[
                "LISTA"
            ].eq(cat)
        ].dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        for _, r in subset.iterrows():
            pm = add(
                folder,
                "Placemark"
            )

            add(
                pm,
                "name",
                (
                    "SUPERPONTO " + str(r["SUPER_PONTO"])
                    if str(r.get("SUPER_PONTO", "")).startswith("SIM")
                    else "Nota " + str(r["NOTA"])
                )
            )

            add(
                pm,
                "styleUrl",
                "#cat"
                + str(
                    CATS.index(cat)
                )
            )

            add(
                pm,
                "description",
                popup(r)
            )

            add(
                add(
                    pm,
                    "Point"
                ),
                "coordinates",
                (
                    f"{r['LONGITUDE']},"
                    f"{r['LATITUDE']},0"
                )
            )

    # SUPERPONTOS
    if (
        superpoints is not None
        and not superpoints.empty
    ):
        folder = add(
            doc,
            "Folder"
        )

        add(
            folder,
            "name",
            "SUPERPONTOS"
        )

        dados = superpoints[
            superpoints[
                "SUPER_PONTO"
            ].astype(str).str.startswith(
                "SIM"
            )
        ].dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        for _, r in dados.iterrows():
            pm = add(
                folder,
                "Placemark"
            )

            add(
                pm,
                "name",
                (
                    "SUPERPONTO "
                    + str(
                        r["SUPER_PONTO"]
                    )
                )
            )

            add(
                pm,
                "styleUrl",
                "#cat"
                + str(
                    CATS.index(
                        r["LISTA"]
                    )
                )
            )

            add(
                pm,
                "description",
                popup(r)
            )

            add(
                add(
                    pm,
                    "Point"
                ),
                "coordinates",
                (
                    f"{r['LONGITUDE']},"
                    f"{r['LATITUDE']},0"
                )
            )

    # TRAJETOS OSRM CONFIRMADOS
    if (
        paths is not None
        and not paths.empty
    ):
        folder = add(
            doc,
            "Folder"
        )

        add(
            folder,
            "name",
            "TRAJETOS OSRM CONFIRMADOS"
        )

        for _, r in paths.iterrows():
            coords = r.get(
                "GEOMETRIA",
                []
            )

            if (
                not isinstance(coords, list)
                or len(coords) < 2
            ):
                continue

            pm = add(
                folder,
                "Placemark"
            )

            add(
                pm,
                "name",
                (
                    f"{r['EQUIPE']} "
                    f"| {r['NOTA']}"
                )
            )

            line = add(
                pm,
                "LineString"
            )

            add(
                line,
                "tessellate",
                "1"
            )

            add(
                line,
                "coordinates",
                " ".join(
                    f"{x},{y},0"
                    for x, y in coords
                )
            )

    return ET.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True
    )


# ============================================================
# LISTA CONTINUA - EXPORTACAO POR EQUIPE MAIS PROXIMA
# ============================================================

def classificar_tipo_obra(categoria):
    """Classificacao original da nota, nunca altera sua contagem."""
    return {
        CAT_SAN: "SANEAMENTO",
        CAT_LEV: "LEVANTAMENTO",
        CAT_DUP: "SANEAMENTO E LEVANTAMENTO",
    }.get(categoria, str(categoria))


def escolher_equipes_proximas(obras, equipes):
    """Atribui UMA nota a UMA equipe mais proxima elegivel.

    A referencia e Haversine, nao distancia rodoviaria. Para obras
    de ambas as atividades, compara equipes dos dois grupos e preserva
    a classificacao original da nota no Excel e no KML.
    """
    saida = obras.copy().reset_index(drop=True)
    for campo in ["TIPO_OBRA", "BASE_ATRIBUIDA", "ID_EQUIPE", "ATIVIDADE_EQUIPE", "CIDADE_BASE", "METODO_DISTANCIA"]:
        saida[campo] = ""
    saida["DISTANCIA_EQUIPE_KM"] = np.nan
    saida["TIPO_OBRA"] = saida["LISTA"].map(classificar_tipo_obra)
    saida["STATUS_DISTRIBUICAO"] = "SEM COORDENADAS OU EQUIPE"
    for tipo in ["SANEAMENTO", "LEVANTAMENTO"]:
        equipe_tipo = equipes.loc[
            equipes["TIPO_EQUIPE"].eq(tipo),
            ["ID_EQUIPE", "EQUIPE", "CIDADE_BASE", "LAT_EQUIPE", "LON_EQUIPE"]
        ].dropna(subset=["LAT_EQUIPE", "LON_EQUIPE"]).reset_index(drop=True)
        if equipe_tipo.empty:
            continue
        mascara = (
            saida["LISTA"].map(lambda cat: tipo in activity(cat))
            & saida["LATITUDE"].notna()
            & saida["LONGITUDE"].notna()
        )
        indices = saida.index[mascara].to_numpy()
        if not len(indices):
            continue
        arvore = BallTree(
            np.radians(equipe_tipo[["LAT_EQUIPE", "LON_EQUIPE"]].to_numpy(float)),
            metric="haversine",
        )
        distancias, posicoes = arvore.query(
            np.radians(saida.loc[indices, ["LATITUDE", "LONGITUDE"]].to_numpy(float)),
            k=1,
        )
        for i, indice in enumerate(indices):
            km = float(distancias[i, 0] * R)
            anterior = saida.at[indice, "DISTANCIA_EQUIPE_KM"]
            if pd.isna(anterior) or km < float(anterior):
                equipe = equipe_tipo.iloc[int(posicoes[i, 0])]
                saida.at[indice, "DISTANCIA_EQUIPE_KM"] = round(km, 2)
                saida.at[indice, "BASE_ATRIBUIDA"] = equipe["EQUIPE"]
                saida.at[indice, "ID_EQUIPE"] = equipe["ID_EQUIPE"]
                saida.at[indice, "ATIVIDADE_EQUIPE"] = tipo
                saida.at[indice, "CIDADE_BASE"] = equipe["CIDADE_BASE"]
                saida.at[indice, "METODO_DISTANCIA"] = "LINHA RETA (HAVERSINE)"
                saida.at[indice, "STATUS_DISTRIBUICAO"] = "EQUIPE MAIS PROXIMA"
    saida["EQUIPE_PROGRAMADA"] = saida["BASE_ATRIBUIDA"]
    return saida


def agrupar_rota_equipe(df, metros):
    """DBSCAN por equipe + classificacao: uma linha por superponto KML;
    Excel mantem uma linha para cada nota e marca superponto e ordem.
    Segue a organizacao da Lista Continua sem alterar os modulos comuns.
    """
    if df.empty:
        return df.copy(), df.copy()
    kml_rows, excel_rows = [], []
    for tipo_obra, subset in df.groupby("TIPO_OBRA", sort=False):
        rad = np.radians(subset[["LATITUDE", "LONGITUDE"]].to_numpy(float))
        labels = DBSCAN(
            eps=float(metros) / 6371000.0,
            min_samples=1,
            metric="haversine",
            algorithm="ball_tree",
        ).fit_predict(rad)
        partes = []
        for _, cluster in subset.assign(_grupo=labels).groupby("_grupo", sort=False):
            primeira = cluster.iloc[0].drop(labels="_grupo").to_dict()
            primeira["LATITUDE"] = float(cluster["LATITUDE"].mean())
            primeira["LONGITUDE"] = float(cluster["LONGITUDE"].mean())
            primeira["SUPER_PONTO"] = f"SIM ({len(cluster)} Obras)" if len(cluster) > 1 else "NÃO"
            primeira["NOTAS_AGRUPADAS"] = " | ".join(cluster["NOTA"].astype(str))
            primeira["_ORIGINAL_ROWS"] = cluster.drop(columns="_grupo").to_dict("records")
            partes.append(primeira)
        # Uma rota heuristica de vizinho proximo por equipe, sem alegar OSRM.
        restante = partes[:]
        atual = None
        ordem = 0
        while restante:
            if atual is None:
                pos = min(range(len(restante)), key=lambda j: (str(restante[j]["MUNICIPIO"]), str(restante[j]["NOTA"])))
                km_anterior = 0.0
            else:
                ds = [float(haversine(atual["LATITUDE"], atual["LONGITUDE"], x["LATITUDE"], x["LONGITUDE"])) for x in restante]
                pos = int(np.argmin(ds))
                km_anterior = round(ds[pos], 2)
            item = restante.pop(pos)
            ordem += 1
            item["ORDEM"] = ordem
            item["DISTANCIA_PONTO_ANTERIOR_KM"] = km_anterior
            kml_rows.append(item)
            for original in item["_ORIGINAL_ROWS"]:
                linha = dict(original)
                linha["ORDEM"] = ordem
                linha["SUPER_PONTO"] = item["SUPER_PONTO"]
                linha["NOTAS_AGRUPADAS"] = item["NOTAS_AGRUPADAS"]
                linha["DISTANCIA_PONTO_ANTERIOR_KM"] = km_anterior
                linha["DIA_SEMANA"] = "NÃO PROGRAMADA"
                linha["DIA_MES"] = ""
                excel_rows.append(linha)
            atual = item
    return pd.DataFrame(excel_rows), pd.DataFrame(kml_rows)


def nome_seguro_equipe(equipe, identificador):
    base = n(equipe).replace(" ", "_")
    base = re.sub(r"[^A-Z0-9_-]", "", base)[:54].strip("_") or "EQUIPE"
    codigo = hashlib.sha1(str(identificador).encode("utf-8")).hexdigest()[:7]
    return f"{base}_{codigo}"


def gerar_zips_lista_continua_nip(obras, equipes, metros, kml_geral):
    """Retorna ZIP Excel/KML + resumo usando TODO o backlog pendente."""
    distribuicao = escolher_equipes_proximas(obras, equipes)
    com_eq = distribuicao.loc[
        distribuicao["STATUS_DISTRIBUICAO"].eq("EQUIPE MAIS PROXIMA")
    ].copy()
    sem_eq = distribuicao.loc[
        ~distribuicao["STATUS_DISTRIBUICAO"].eq("EQUIPE MAIS PROXIMA")
    ].copy()
    resumo = (
        com_eq.groupby(["ATIVIDADE_EQUIPE", "BASE_ATRIBUIDA", "ID_EQUIPE"], dropna=False)
        .agg(OBRAS=("NOTA", "nunique"), MUNICIPIOS=("MUNICIPIO", "nunique"))
        .reset_index()
    )
    data_fmt = datetime.now().strftime("%d.%m.%Y")
    excel_zip = io.BytesIO()
    kml_zip = io.BytesIO()
    with zipfile.ZipFile(excel_zip, "w", zipfile.ZIP_DEFLATED) as zx, zipfile.ZipFile(kml_zip, "w", zipfile.ZIP_DEFLATED) as zk:
        zx.writestr(f"Resumo_Operacional - {data_fmt}.xlsx", excel_bytes({"Resumo Operacional": resumo, "Sem Equipe": sem_eq}))
        zx.writestr(f"Demanda_ListaContinua_Total - {data_fmt}.xlsx", excel_bytes({"Obras Roteirizadas": distribuicao}))
        zk.writestr(f"ROTA_TOTAL - {data_fmt}.kml", kml_geral)
        for ident, grupo in com_eq.groupby("ID_EQUIPE", sort=True):
            info = grupo.iloc[0]
            atividade = info["ATIVIDADE_EQUIPE"]
            safe = nome_seguro_equipe(info["BASE_ATRIBUIDA"], ident)
            excel_rows, kml_rows = agrupar_rota_equipe(grupo, metros)
            # Uma linha por nota original, inclusive as agrupadas em superpontos.
            primeiras = ["BASE_ATRIBUIDA", "TIPO_OBRA", "ATIVIDADE_EQUIPE", "DIA_SEMANA", "DIA_MES", "SUPER_PONTO", "ORDEM", "DISTANCIA_PONTO_ANTERIOR_KM", "NOTA", "MUNICIPIO", "REGIONAL", "DISTANCIA_EQUIPE_KM", "METODO_DISTANCIA"]
            outras = [c for c in excel_rows.columns if c not in primeiras and not c.startswith("_")]
            excel_rows = excel_rows[[c for c in primeiras if c in excel_rows.columns] + outras]
            zx.writestr(f"Rotas_{data_fmt}/{atividade}/Rota_{safe}.xlsx", excel_bytes({"Obras Roteirizadas": excel_rows}))
            # O KML da equipe tem um marcador por superponto ou obra individual.
            zk.writestr(f"KML_{data_fmt}/{atividade}/Rota_{safe}.kml", kml_bytes(kml_rows, name=f"Rota {info['BASE_ATRIBUIDA']}"))
    return excel_zip.getvalue(), kml_zip.getvalue(), distribuicao, resumo, sem_eq


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

st.markdown(
    """
    <div class="nip-title">
        📍 NIP | Análise Cruzada e Planejamento
    </div>
    <div class="nip-muted">
        Backlog | Equipes | OSRM |
        Superpontos | Excel e KML
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# REGRAS
# ============================================================

with st.expander(
    "📘 Regras de contagem"
):
    st.markdown("""
    **STATUS SAP:** somente FINL e CANC excluídos.

    **STATUS LIST aceitos:**
    0, Em levantamento e Correção de levantamento.

    **Contratos aceitos:**
    0 ou NIP GLOBAL LTDA - EQTL MARANHÃO.

    **Colunas M e N:** precisam estar iguais a zero.

    **Colunas B e C:**
    se B contém equipe diferente de SEM LEVANTADOR
    e C contém data válida, a nota fica fora
    da contagem.

    **Duplicadas:** contam uma única vez na
    volumetria, podendo gerar tarefas distintas.

    **Distância:** a alocação exige rota
    rodoviária OSRM confirmada.

    **Sem rota confirmada:** a tarefa permanece
    sem alocação.
    """)


# ============================================================
# ZERAR ANALISE - NAO ALTERA OUTRAS PAGINAS DO ROTEIRIZADOR
# ============================================================

def zerar_analise_nip():
    # A callback ocorre ANTES de criar os widgets nesta execucao.
    versao_atual = st.session_state.get("nip_analise_versao", 0)
    for chave in list(st.session_state.keys()):
        if (
            chave.startswith("nip_analise_widget_")
            or chave in {
                "nipbase", "niproutes", "nipgeom", "capacity_editor",
            }
        ):
            del st.session_state[chave]
    # Novas chaves recriam os uploaders e todos os filtros vazios/padrao.
    st.session_state["nip_analise_versao"] = versao_atual + 1


versao_analise = st.session_state.get("nip_analise_versao", 0)

col_acoes, col_instrucoes = st.columns([1, 3], vertical_alignment="center")
with col_acoes:
    st.button(
        "🗑️ Zerar análise",
        key="nip_analise_botao_zerar",
        type="secondary",
        use_container_width=True,
        on_click=zerar_analise_nip,
        help="Remove os arquivos anexados, bases processadas, filtros, programação e progresso OSRM desta análise.",
    )
with col_instrucoes:
    st.caption("Limpa somente a Análise Cruzada; não apaga os arquivos do seu computador nem altera a Lista Contínua.")


# ============================================================
# UPLOAD DAS BASES
# ============================================================

st.subheader(
    "📂 Importação das bases"
)

a, b, c = st.columns(3)

with a:
    san_file = st.file_uploader(
        "BASE_SANEAMENTO",
        type=[
            "xlsx",
            "csv"
        ],
        key=f"nip_analise_widget_{versao_analise}_0"
)

with b:
    lev_file = st.file_uploader(
        "BASE_LEVANTAMENTO_ATUALIZADA",
        type=["xlsx"],
        key=f"nip_analise_widget_{versao_analise}_1"
)

with c:
    team_file = st.file_uploader(
        "LOCALIDADE LEVANTADORES-SANEAMENTO",
        type=["xlsx"],
        key=f"nip_analise_widget_{versao_analise}_2"
)

if not all([
    san_file,
    lev_file,
    team_file
]):
    st.info(
        "Envie os três arquivos obrigatórios."
    )
    st.stop()


# ============================================================
# ASSINATURA DOS ARQUIVOS
# ============================================================

sig = tuple(
    hashlib.sha256(
        f.getvalue()
    ).hexdigest()
    for f in [
        san_file,
        lev_file,
        team_file
    ]
)


# ============================================================
# PROCESSAMENTO - CRONOMETRO COM ATUALIZACAO EM TEMPO REAL
# ============================================================

PROCESS_STAGES = [
    ("Lendo a base de Saneamento", 0.13),
    ("Lendo a base de Levantamento", 0.73),
    ("Lendo a base de equipes", 0.05),
    ("Cruzando notas e classificando obras", 0.07),
    ("Verificando inconsistencias", 0.02),
]


def tempo_legivel(segundos):
    if segundos is None:
        return "Calculando..."
    total = max(0, int(round(segundos)))
    horas, resto = divmod(total, 3600)
    minutos, secs = divmod(resto, 60)
    return (
        f"{horas:02d}:{minutos:02d}:{secs:02d}"
        if horas else f"{minutos:02d}:{secs:02d}"
    )


def processar_bases_em_thread(san_bytes, lev_bytes, team_bytes, estado):
    # Nenhuma chamada st.* dentro da thread: apenas leitura/processamento.
    class ArquivoMemoria:
        def __init__(self, dados, nome):
            self._dados = dados
            self.name = nome

        def getvalue(self):
            return self._dados

    def etapa(indice):
        estado["indice"] = indice
        estado["inicio_etapa"] = time.monotonic()

    etapa(0)
    san, sa = load_san(ArquivoMemoria(*san_bytes))
    etapa(1)
    lev, info = load_lev(ArquivoMemoria(*lev_bytes))
    etapa(2)
    teams = load_teams(ArquivoMemoria(*team_bytes))
    etapa(3)
    audit = consolidate(san, lev)
    etapa(4)
    errors = anomalies(san, lev, audit, teams)
    estado["indice"] = len(PROCESS_STAGES)
    return {
        "audit": audit,
        "teams": teams,
        "san_name": sa,
        "lev_info": info,
        "issues": errors,
    }


if st.button(
    "🚀 Processar bases",
    type="primary",
    use_container_width=True
):
    inicio_processamento = time.monotonic()
    estado_progresso = {"indice": 0, "inicio_etapa": inicio_processamento}
    linha_status = st.empty()
    barra_status = st.empty()
    linha_tempo = st.empty()

    # Se os arquivos mudaram, nao utilizar resultados anteriores ao clicar.
    st.session_state.pop("nipbase", None)

    arquivos = (
        (san_file.getvalue(), san_file.name),
        (lev_file.getvalue(), lev_file.name),
        (team_file.getvalue(), team_file.name),
    )

    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(
                processar_bases_em_thread,
                *arquivos,
                estado_progresso
            )

            while not future.done():
                decorrido = time.monotonic() - inicio_processamento
                indice = min(estado_progresso["indice"], len(PROCESS_STAGES)-1)
                percentual = sum(peso for _, peso in PROCESS_STAGES[:indice])
                etapa_nome = PROCESS_STAGES[indice][0]

                # A ETA e uma projecao por etapa, nao uma promessa exata.
                # Primeira etapa: aguarda dados para estimar o restante.
                restante = (
                    decorrido / percentual * (1 - percentual)
                    if percentual >= 0.01 else None
                )
                linha_status.info(f"⏳ {etapa_nome}...")
                barra_status.progress(
                    max(0.01, min(0.99, percentual)),
                    text=f"Etapas concluídas: {indice}/{len(PROCESS_STAGES)}"
                )
                linha_tempo.markdown(
                    f"**⏱️ Tempo decorrido:** {tempo_legivel(decorrido)}"
                    f"　 |　 **⌛ Tempo restante estimado:** {tempo_legivel(restante)}"
                )
                time.sleep(0.4)

            # Propaga erros da thread para o tratamento abaixo.
            resultado = future.result()

        total = time.monotonic() - inicio_processamento
        resultado.update({
            "sig": sig,
            "seconds": total,
        })
        st.session_state["nipbase"] = resultado
        st.session_state.pop("niproutes", None)
        st.session_state.pop("nipgeom", None)

        linha_status.success("✅ Bases processadas com sucesso.")
        barra_status.progress(1.0, text="Processamento concluído")
        linha_tempo.markdown(
            f"**⏱️ Tempo total:** {tempo_legivel(total)}"
            "　 |　 **⌛ Tempo restante:** 00:00"
        )

    except Exception as e:
        linha_status.error(f"Erro no processamento: {e}")
        barra_status.empty()
        linha_tempo.markdown(
            f"**⏱️ Tempo decorrido:** "
            f"{tempo_legivel(time.monotonic() - inicio_processamento)}"
        )

if "nipbase" not in st.session_state:
    st.stop()

data = st.session_state[
    "nipbase"
]

if data["sig"] != sig:
    st.warning(
        "Os arquivos mudaram. Processe novamente."
    )
    st.stop()

audit = data["audit"]
teams = data["teams"]
issues = data["issues"]

st.caption(
    f"Abas: Saneamento = "
    f"{data['san_name']} | "
    f"Levantamento = "
    f"{data['lev_info']['aba']} | "
    f"B: {data['lev_info']['B']} | "
    f"C: {data['lev_info']['C']} | "
    f"Leitura: {data['seconds']:.1f}s"
)


# ============================================================
# CONFIGURACOES LATERAIS
# ============================================================

with st.sidebar:
    st.header(
        "⚙️ Planejamento"
    )

    radius = st.select_slider(
        "Raio máximo rodoviário (km)",
        options=list(
            range(
                200,
                501,
                25
            )
        ),
        value=200,
        key=f"nip_analise_widget_{versao_analise}_3"
)

    k = st.slider(
        "Equipes candidatas por atividade",
        1,
        5,
        3,
        key=f"nip_analise_widget_{versao_analise}_4"
)

    super_m = st.slider(
        "Raio de superponto (m)",
        5,
        500,
        50,
        5,
        key=f"nip_analise_widget_{versao_analise}_5"
)

    medium_days = st.number_input(
        "Dias para prioridade média",
        1,
        365,
        14,
        key=f"nip_analise_widget_{versao_analise}_6"
)

    st.divider()

    st.subheader(
        "📅 Capacidade"
    )

    mode = st.selectbox(
        "Programação",
        [
            "POR QUANTIDADE",
            "POR DIA",
            "POR SEMANA",
            "POR MÊS"
        ],
        key=f"nip_analise_widget_{versao_analise}_7"
)

    period_cap = st.number_input(
        "Tarefas por equipe no período",
        1,
        10000,
        20,
        key=f"nip_analise_widget_{versao_analise}_8"
)

    daily_cap = st.number_input(
        "Tarefas por dia (padrão)",
        1,
        100,
        8,
        key=f"nip_analise_widget_{versao_analise}_9"
)

    periods = st.number_input(
        "Quantidade de períodos",
        1,
        52,
        5,
        key=f"nip_analise_widget_{versao_analise}_10"
)

    first_day = st.date_input(
        "Data inicial",
        value=datetime.today().date(),
        key=f"nip_analise_widget_{versao_analise}_11"
)

    st.divider()

    st.subheader(
        "🛣️ Rotas reais"
    )

    server = st.text_input(
        "OSRM",
        OSRM,
        key=f"nip_analise_widget_{versao_analise}_12"
)

    batch = st.slider(
        "Consultas por lote",
        1,
        25,
        5,
        key=f"nip_analise_widget_{versao_analise}_13"
)

    kml_osrm = st.checkbox(
        "Incluir traçados reais OSRM no KML",
        value=False,
        key=f"nip_analise_widget_{versao_analise}_14"
)

    geom_batch = st.slider(
        "Traçados por lote",
        1,
        15,
        4,
        key=f"nip_analise_widget_{versao_analise}_15"
)

    st.divider()

    previous = st.file_uploader(
        "Análise anterior (opcional)",
        type=["xlsx"],
        key=f"nip_analise_widget_{versao_analise}_16"
)


# ============================================================
# DADOS OPERACIONAIS
# ============================================================

pending = priority(
    audit[
        audit[
            "PENDENTE_CONTAGEM"
        ].eq("SIM")
    ].copy(),
    medium_days
)

finl = audit[
    audit[
        "SAP_FINL"
    ].eq("SIM")
]

canc = audit[
    audit[
        "SAP_CANC"
    ].eq("SIM")
]

field = audit[
    audit[
        "JA_EM_CAMPO"
    ].eq("SIM")
]

dup = audit[
    audit[
        "DUPLICADA"
    ].eq("SIM")
]


# ============================================================
# INDICADORES
# ============================================================

metrics = [
    (
        "Saneamento",
        int(
            pending[
                "LISTA"
            ].eq(CAT_SAN).sum()
        ),
        "#2563eb"
    ),
    (
        "Levantamento",
        int(
            pending[
                "LISTA"
            ].eq(CAT_LEV).sum()
        ),
        "#16a34a"
    ),
    (
        "Duplicadas",
        int(
            pending[
                "LISTA"
            ].eq(CAT_DUP).sum()
        ),
        "#9333ea"
    ),
    (
        "Total pendente",
        len(pending),
        "#0d256c"
    ),
    (
        "Excluídas",
        int(
            audit[
                "PENDENTE_CONTAGEM"
            ].eq("NÃO").sum()
        ),
        "#dc2626"
    )
]

metric_groups = [
    metrics,
    [
        (
            "FINL",
            len(finl),
            "#64748b"
        ),
        (
            "CANC",
            len(canc),
            "#ea580c"
        ),
        (
            "Já em campo",
            len(field),
            "#0891b2"
        ),
        (
            "Inconsistências",
            len(issues),
            "#d97706"
        )
    ]
]

for block in metric_groups:
    cols = st.columns(
        len(block)
    )

    for cn, (
        label,
        value,
        color
    ) in zip(cols, block):

        cn.markdown(
            (
                '<div class="nip-card" '
                f'style="--accent:{color}">'
                '<div class="nip-label">'
                f'{label}'
                '</div>'
                '<div class="nip-value">'
                f'{value:,}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True
        )


# ============================================================
# FILTROS OPERACIONAIS
# ============================================================

st.subheader(
    "🔎 Filtros operacionais"
)

f1, f2, f3 = st.columns(3)

with f1:
    reg = st.multiselect(
        "Regional",
        sorted(
            pending[
                "REGIONAL"
            ].fillna("").unique()
        ),
        key=f"nip_analise_widget_{versao_analise}_17"
)

with f2:
    city = st.multiselect(
        "Município",
        sorted(
            pending[
                "MUNICIPIO"
            ].fillna("").unique()
        ),
        key=f"nip_analise_widget_{versao_analise}_18"
)

with f3:
    prio = st.multiselect(
        "Prioridade",
        [
            "ALTA",
            "MEDIA",
            "BAIXA"
        ],
        key=f"nip_analise_widget_{versao_analise}_19"
)

search = st.text_input(
    "Pesquisar nota",
        key=f"nip_analise_widget_{versao_analise}_20"
).strip()

view = pending.copy()

if reg:
    view = view[
        view[
            "REGIONAL"
        ].isin(reg)
    ]

if city:
    view = view[
        view[
            "MUNICIPIO"
        ].isin(city)
    ]

if prio:
    view = view[
        view[
            "PRIORIDADE_PLANEJAMENTO"
        ].isin(prio)
    ]

if search:
    view = view[
        view[
            "NOTA"
        ].astype(str).str.contains(
            re.escape(search),
            case=False,
            na=False
        )
    ]

view = view.reset_index(
    drop=True
)


# ============================================================
# CANDIDATOS
# ============================================================

view, candidates = team_candidates(
    view,
    teams,
    k
)

if candidates.empty:
    st.warning(
        "Nenhuma equipe candidata para "
        "as obras filtradas."
    )


# ============================================================
# METAS POR EQUIPE
# ============================================================

st.subheader(
    "👥 Metas por equipe"
)

capacity = teams[
    [
        "ID_EQUIPE",
        "TIPO_EQUIPE",
        "EQUIPE",
        "CIDADE_BASE"
    ]
].copy()

capacity["CAPACIDADE_DIA"] = int(
    daily_cap
)

capacity = st.data_editor(
    capacity,
    hide_index=True,
    use_container_width=True,
    disabled=[
        "ID_EQUIPE",
        "TIPO_EQUIPE",
        "EQUIPE",
        "CIDADE_BASE"
    ],
    key=f"nip_analise_widget_{versao_analise}_21"
)

capacity["CAPACIDADE_DIA"] = (
    pd.to_numeric(
        capacity["CAPACIDADE_DIA"],
        errors="coerce"
    )
    .fillna(daily_cap)
    .clip(lower=1)
    .astype(int)
)


# ============================================================
# ASSINATURA DAS ROTAS
# ============================================================

route_key = hashlib.sha256(
    (
        candidates.to_json(
            orient="records"
        )
        + str(radius)
        + server
    ).encode()
).hexdigest()

if (
    "niproutes" not in st.session_state
    or st.session_state[
        "niproutes"
    ]["key"] != route_key
):
    st.session_state[
        "niproutes"
    ] = {
        "key": route_key,
        "df": initial_routes(
            candidates,
            radius
        ),
        "elapsed": 0.0
    }

rs = st.session_state[
    "niproutes"
]

routes = rs["df"]

waiting = int(
    routes[
        "ROTA_STATUS"
    ].eq("NÃO CONSULTADA").sum()
)

outside = int(
    routes[
        "ROTA_STATUS"
    ].eq("FORA RAIO GEOGRÁFICO").sum()
)

done = (
    len(routes)
    - waiting
    - outside
)

total = done + waiting


# ============================================================
# CONSULTAS RODOVIARIAS
# ============================================================

st.subheader(
    "🛣️ Validação rodoviária antes da distribuição"
)

if total:
    remaining = (
        rs["elapsed"]
        / done
        * waiting
        if done
        else None
    )

    st.progress(
        done / total
    )

    texto_restante = (
        f"{remaining:.1f}s"
        if remaining is not None
        else "calculando"
    )

    st.caption(
        f"Consultas: {done}/{total} | "
        f"Decorrido: {rs['elapsed']:.1f}s | "
        f"Restante estimado: {texto_restante}"
    )

else:
    st.info(
        "Sem candidatos para consultar."
    )

if waiting:
    if st.button(
        "🚗 Consultar próximo lote OSRM",
        type="primary"
    ):
        msg = st.empty()

        new, processed, spent = (
            compute_route_batch(
                routes,
                0,
                batch,
                server,
                msg
            )
        )

        rs["df"] = new
        rs["elapsed"] += spent

        st.session_state[
            "niproutes"
        ] = rs

        st.rerun()

if not waiting and total:
    st.success(
        "Consultas rodoviárias concluídas."
    )

routes = rs["df"]


# ============================================================
# PROGRAMACAO
# ============================================================

filtered_tasks = tasks_for(
    view
)

plan, load = schedule_tasks(
    filtered_tasks,
    routes,
    teams,
    capacity,
    mode,
    periods,
    first_day,
    period_cap,
    daily_cap,
    radius,
    True
)

allocated = (
    plan[
        plan[
            "STATUS_PROGRAMACAO"
        ].eq("SUGESTÃO")
    ]
    if not plan.empty
    else pd.DataFrame()
)


# ============================================================
# SUPERPONTOS
# ============================================================

superpoints, reduction = (
    merge_superpoints(
        allocated,
        super_m
    )
)


# ============================================================
# MEDIA DE OBRAS POR ATIVIDADE
# ============================================================

st.subheader(
    "📊 Cobertura e médias por atividade"
)

summary = []

for kind, categories in [
    (
        "SANEAMENTO",
        [
            CAT_SAN,
            CAT_DUP
        ]
    ),
    (
        "LEVANTAMENTO",
        [
            CAT_LEV,
            CAT_DUP
        ]
    )
]:
    nteam = int(
        teams[
            "TIPO_EQUIPE"
        ].eq(kind).sum()
    )

    nwork = int(
        view[
            "LISTA"
        ].isin(categories).sum()
    )

    nalloc = (
        int(
            allocated[
                "ATIVIDADE"
            ].eq(kind).sum()
        )
        if not allocated.empty
        else 0
    )

    summary.append({
        "ATIVIDADE": kind,
        "EQUIPES": nteam,
        "OBRAS_DISPONÍVEIS": nwork,
        "MÉDIA_POR_EQUIPE": (
            round(
                nwork / nteam,
                2
            )
            if nteam
            else 0
        ),
        "TAREFAS_ALOCADAS": nalloc,
        "MÉDIA_PROGRAMADA": (
            round(
                nalloc / nteam,
                2
            )
            if nteam
            else 0
        )
    })

summary = pd.DataFrame(
    summary
)

st.dataframe(
    summary,
    hide_index=True,
    use_container_width=True
)

actual_load = (
    allocated.groupby(
        "ID_EQUIPE"
    ).size().to_dict()
    if not allocated.empty
    else {}
)

capacity["TAREFAS_PROGRAMADAS"] = (
    capacity[
        "ID_EQUIPE"
    ].map(actual_load)
    .fillna(0)
    .astype(int)
)


# ============================================================
# HISTORICO
# ============================================================

new = pd.DataFrame(
    columns=["NOTA"]
)

gone = pd.DataFrame(
    columns=["NOTA"]
)

if previous is not None:
    try:
        new, gone = compare_history(
            pending,
            previous
        )

    except Exception as e:
        st.warning(
            f"Histórico não comparado: {e}"
        )


# ============================================================
# FILTROS DE EXPORTACAO
# ============================================================

st.subheader(
    "📦 Escopo das exportações"
)

exp1, exp2, exp3 = st.columns(3)

with exp1:
    exp_kind = st.multiselect(
        "Atividade para exportar",
        [
            "SANEAMENTO",
            "LEVANTAMENTO"
        ],
        default=[
            "SANEAMENTO",
            "LEVANTAMENTO"
        ],
        key=f"nip_analise_widget_{versao_analise}_22"
)

with exp2:
    equipes_exportaveis = (
        sorted(
            allocated[
                "EQUIPE_PROGRAMADA"
            ].dropna().unique().tolist()
        )
        if not allocated.empty
        else []
    )

    exp_team = st.multiselect(
        "Equipe para exportar",
        equipes_exportaveis,
        key=f"nip_analise_widget_{versao_analise}_23"
)

with exp3:
    exp_dates = st.date_input(
        "Período da programação",
        value=(
            first_day,
            first_day
            + timedelta(days=30)
        ),
        key=f"nip_analise_widget_{versao_analise}_24"
)

exp_plan = plan.copy()

if not exp_plan.empty:
    if exp_kind:
        exp_plan = exp_plan[
            exp_plan[
                "ATIVIDADE"
            ].isin(exp_kind)
        ]

    else:
        exp_plan = (
            exp_plan.iloc[0:0]
        )

    if exp_team:
        exp_plan = exp_plan[
            exp_plan[
                "EQUIPE_PROGRAMADA"
            ].isin(exp_team)
        ]

    if (
        isinstance(
            exp_dates,
            (tuple, list)
        )
        and len(exp_dates) == 2
    ):
        # Normalize both sides to pandas Timestamp (no date/datetime64 mismatch).
        dat = pd.to_datetime(
            exp_plan["DATA_PROGRAMADA"],
            errors="coerce"
        ).dt.normalize()

        data_inicial = pd.Timestamp(exp_dates[0]).normalize()
        data_final = pd.Timestamp(exp_dates[1]).normalize()

        if data_final < data_inicial:
            data_inicial, data_final = data_final, data_inicial

        mascara_datas = dat.between(
            data_inicial,
            data_final,
            inclusive="both"
        ).fillna(False)

        exp_plan = exp_plan.loc[mascara_datas].copy()

exp_notas = (
    set(
        exp_plan[
            "NOTA"
        ].astype(str)
    )
    if not exp_plan.empty
    else set()
)

if (
    exp_team
    or (
        isinstance(
            exp_dates,
            (tuple, list)
        )
        and len(exp_dates) == 2
    )
):
    exp_works = view[
        view[
            "NOTA"
        ].astype(str).isin(
            exp_notas
        )
    ]

else:
    exp_works = view

exp_sp = (
    superpoints[
        superpoints[
            "NOTA"
        ].astype(str).isin(
            exp_notas
        )
    ]
    if not superpoints.empty
    else superpoints
)


# ============================================================
# PREPARACAO DE TRECHOS PARA KML
# ============================================================

def build_segments(p):
    rows = []

    if p.empty:
        return pd.DataFrame()

    validas = p[
        p[
            "STATUS_PROGRAMACAO"
        ].eq("SUGESTÃO")
    ]

    for (
        tid,
        day
    ), g in validas.groupby(
        [
            "ID_EQUIPE",
            "DATA_PROGRAMADA"
        ]
    ):
        eq = teams[
            teams[
                "ID_EQUIPE"
            ].eq(tid)
        ]

        if eq.empty:
            continue

        lat = float(
            eq.iloc[0][
                "LAT_EQUIPE"
            ]
        )

        lon = float(
            eq.iloc[0][
                "LON_EQUIPE"
            ]
        )

        rest = g.to_dict(
            "records"
        )

        order = 0

        while rest:
            distances = [
                float(
                    haversine(
                        lat,
                        lon,
                        float(
                            r["LATITUDE"]
                        ),
                        float(
                            r["LONGITUDE"]
                        )
                    )
                )
                for r in rest
            ]

            idx = int(
                np.argmin(
                    distances
                )
            )

            r = rest.pop(idx)

            order += 1

            rows.append({
                "NOTA": r["NOTA"],
                "ATIVIDADE": r["ATIVIDADE"],
                "EQUIPE": r["EQUIPE_PROGRAMADA"],
                "ID_EQUIPE": tid,
                "DATA": day,
                "ORDEM": order,
                "LAT1": lat,
                "LON1": lon,
                "LAT2": float(r["LATITUDE"]),
                "LON2": float(r["LONGITUDE"])
            })

            lat = float(
                r["LATITUDE"]
            )

            lon = float(
                r["LONGITUDE"]
            )

    return pd.DataFrame(rows)


segments = build_segments(
    exp_plan
)


# ============================================================
# ASSINATURA DO KML OSRM
# ============================================================

geom_key = hashlib.sha256(
    (
        segments.to_json(
            orient="records"
        )
        + server
    ).encode()
).hexdigest()

if (
    "nipgeom" not in st.session_state
    or st.session_state[
        "nipgeom"
    ]["key"] != geom_key
):
    st.session_state[
        "nipgeom"
    ] = {
        "key": geom_key,
        "cursor": 0,
        "items": [],
        "elapsed": 0.0
    }

geom = st.session_state[
    "nipgeom"
]


# ============================================================
# PROCESSAMENTO DOS TRACADOS OSRM
# ============================================================

if (
    kml_osrm
    and not segments.empty
):
    st.markdown(
        "**Traçados OSRM para KML**"
    )

    nseg = len(segments)

    processed = geom["cursor"]

    left = (
        geom["elapsed"]
        / processed
        * (nseg - processed)
        if processed
        else None
    )

    st.progress(
        processed / nseg
    )

    restante_texto = (
        f"{left:.1f}s"
        if left is not None
        else "calculando"
    )

    st.caption(
        f"{processed}/{nseg} trechos | "
        f"Decorrido: {geom['elapsed']:.1f}s | "
        f"Restante: {restante_texto}"
    )

    if (
        processed < nseg
        and st.button(
            "🗺️ Traçar próximo lote"
        )
    ):
        t0 = time.monotonic()

        fim = min(
            processed + geom_batch,
            nseg
        )

        for idx in range(
            processed,
            fim
        ):
            s = segments.iloc[idx]

            resp = osrm_route(
                round(
                    s["LAT1"],
                    6
                ),
                round(
                    s["LON1"],
                    6
                ),
                round(
                    s["LAT2"],
                    6
                ),
                round(
                    s["LON2"],
                    6
                ),
                server,
                True
            )

            if (
                resp["ok"]
                and resp["geometry"]
            ):
                geom["items"].append({
                    **s.to_dict(),
                    "GEOMETRIA": resp["geometry"],
                    "KM_RODOVIARIO": resp["km"]
                })

        geom["cursor"] = fim

        geom["elapsed"] += (
            time.monotonic()
            - t0
        )

        st.session_state[
            "nipgeom"
        ] = geom

        st.rerun()

paths = (
    pd.DataFrame(
        geom["items"]
    )
    if kml_osrm
    else pd.DataFrame()
)


# ============================================================
# ABAS
# ============================================================

names = [
    "🔵 SANEAMENTO",
    "🟢 LEVANTAMENTO",
    "🟣 DUPLICADAS",
    "📅 PROGRAMAÇÃO",
    "🏢 SUPERPONTOS",
    "⚠️ AUDITORIA",
    "📊 HISTÓRICO",
    "🗺️ MAPA"
]

tabs = st.tabs(
    names
)


# ============================================================
# CATEGORIAS
# ============================================================

for tab, kind in zip(
    tabs[:3],
    CATS
):
    with tab:
        x = view[
            view[
                "LISTA"
            ].eq(kind)
        ]

        st.write(
            f"**{len(x)} obras**"
        )

        st.dataframe(
            x,
            hide_index=True,
            use_container_width=True
        )


# ============================================================
# PROGRAMACAO
# ============================================================

with tabs[3]:
    st.metric(
        "Alocadas com OSRM confirmado",
        len(allocated)
    )

    st.metric(
        "Ainda sem alocação",
        len(plan) - len(allocated)
    )

    st.dataframe(
        plan,
        hide_index=True,
        use_container_width=True
    )

    st.dataframe(
        load,
        hide_index=True,
        use_container_width=True
    )


# ============================================================
# SUPERPONTOS
# ============================================================

with tabs[4]:
    quantidade_super = (
        int(
            superpoints[
                "SUPER_PONTO"
            ].astype(str).str.startswith(
                "SIM"
            ).sum()
        )
        if not superpoints.empty
        else 0
    )

    st.metric(
        "Superpontos com múltiplas tarefas",
        quantidade_super
    )

    st.metric(
        "Tarefas agrupadas",
        reduction
    )

    st.dataframe(
        safe_excel(
            superpoints
        ),
        hide_index=True,
        use_container_width=True
    )


# ============================================================
# AUDITORIA
# ============================================================

with tabs[5]:
    t = st.tabs([
        "Excluídas",
        "FINL",
        "CANC",
        "Já em campo",
        "Duplicadas totais",
        "Inconsistências"
    ])

    frames = [
        audit[
            audit[
                "PENDENTE_CONTAGEM"
            ].eq("NÃO")
        ],
        finl,
        canc,
        field,
        dup,
        issues
    ]

    for tab, frame in zip(
        t,
        frames
    ):
        with tab:
            st.dataframe(
                frame,
                hide_index=True,
                use_container_width=True
            )


# ============================================================
# HISTORICO
# ============================================================

with tabs[6]:
    st.metric(
        "Notas novas",
        len(new)
    )

    st.metric(
        "Notas que saíram",
        len(gone)
    )

    st.dataframe(
        new,
        hide_index=True
    )

    st.dataframe(
        gone,
        hide_index=True
    )


# ============================================================
# MAPA
# ============================================================

with tabs[7]:
    if st.checkbox(
        "Carregar mapa",
        key=f"nip_analise_widget_{versao_analise}_25"
):
        geo = view.dropna(
            subset=[
                "LATITUDE",
                "LONGITUDE"
            ]
        )

        if geo.empty:
            st.warning(
                "Sem coordenadas válidas."
            )

        else:
            m = folium.Map(
                location=[
                    geo[
                        "LATITUDE"
                    ].mean(),
                    geo[
                        "LONGITUDE"
                    ].mean()
                ],
                zoom_start=7
            )

            for cat in CATS:
                fg = folium.FeatureGroup(
                    name=cat
                )

                cl = MarkerCluster().add_to(
                    fg
                )

                subset = geo[
                    geo[
                        "LISTA"
                    ].eq(cat)
                ]

                for _, r in subset.iterrows():
                    folium.Marker(
                        [
                            r["LATITUDE"],
                            r["LONGITUDE"]
                        ],
                        popup=folium.Popup(
                            popup(r),
                            max_width=450
                        ),
                        icon=folium.Icon(
                            color=MAP_COLORS[
                                cat
                            ]
                        )
                    ).add_to(cl)

                fg.add_to(m)

            if not paths.empty:
                fg = folium.FeatureGroup(
                    name=(
                        "Traçados rodoviários OSRM"
                    )
                )

                for _, s in paths.iterrows():
                    folium.PolyLine(
                        [
                            [
                                v[1],
                                v[0]
                            ]
                            for v in s["GEOMETRIA"]
                        ],
                        color="blue",
                        weight=3
                    ).add_to(fg)

                fg.add_to(m)

            folium.LayerControl().add_to(
                m
            )

            st_folium(
                m,
                use_container_width=True,
                height=570
            )


# ============================================================
# EXPORTACOES
# ============================================================

st.subheader(
    "📥 Exportações"
)

book = excel_bytes({
    "RESUMO": summary,
    "CONSOLIDADO": view,
    "PROGRAMACAO": exp_plan,
    "CARGA": load,
    "SUPERPONTOS": exp_sp,
    "EQUIPES": capacity,
    "CANDIDATOS OSRM": routes,
    "AUDITORIA COMPLETA": audit,
    "INCONSISTENCIAS": issues,
    "DUPLICADAS": dup,
    "FINL": finl,
    "CANC": canc,
    "JA EM CAMPO": field,
    "NOVAS": new,
    "SAIRAM": gone
})

# KML GERAL: todas as obras PENDENTES, independentemente dos filtros
# de programacao, equipe, periodo e prioridade. Exporta apenas notas
# com coordenadas validas. Nao cria trajetos/rotas inventados.
# Enriquece TODAS as notas pendentes para o KML geral, não apenas
# as notas selecionadas nos filtros de programação. O cálculo inicial
# de proximidade é geográfico e não se confunde com distância OSRM.
kml_total_works, _ = team_candidates(
    pending.dropna(subset=["LATITUDE", "LONGITUDE"]).copy(),
    teams,
    k
)
kml = kml_bytes(
    kml_total_works,
    name="NIP - Todas as Obras Pendentes"
)

# KML FILTRADO: mantem o comportamento anterior, inclusive
# superpontos e trajetos calculados para a programacao selecionada.
kml_filtrado = kml_bytes(
    exp_works,
    exp_sp,
    paths,
    name="NIP - Obras Programadas / Filtradas"
)

st.caption(
    f"KML geral: {len(kml_total_works):,} notas pendentes com coordenadas | "
    f"Notas pendentes sem coordenadas: "
    f"{len(pending) - len(kml_total_works):,} | "
    f"KML filtrado: {len(exp_works.dropna(subset=['LATITUDE', 'LONGITUDE'])):,} notas"
)


# ============================================================
# ZIPS POR EQUIPE - PADRAO LISTA CONTINUA
# ============================================================
# Este ZIP independe do estado das consultas OSRM, do periodo e das
# equipes programadas. Cada nota pendente com coordenadas validas
# e atribuida a uma unica equipe fisicamente mais proxima, dentro da
# atividade aplicavel, com distancia identificada como LINHA RETA.
zip_excel_equipes, zip_kml_equipes, distribuicao_geral, resumo_equipes, sem_equipe = (
    gerar_zips_lista_continua_nip(
        kml_total_works,
        teams,
        super_m,
        kml,
    )
)
st.caption(
    f"Distribuição Lista Contínua: {len(distribuicao_geral) - len(sem_equipe):,} "
    f"notas atribuídas à equipe mais próxima | "
    f"{len(sem_equipe):,} sem equipe | "
    f"{resumo_equipes['ID_EQUIPE'].nunique() if not resumo_equipes.empty else 0} equipes. "
    "Uma nota classificada como SANEAMENTO E LEVANTAMENTO fica em apenas "
    "um arquivo de equipe nesta distribuição de proximidade. "
    "Distâncias indicadas são em linha reta."
)
with st.expander("👥 Conferir distribuição por equipe (Lista Contínua)"):
    st.dataframe(resumo_equipes, hide_index=True, use_container_width=True)
    if not sem_equipe.empty:
        st.warning(f"{len(sem_equipe)} notas sem coordenada ou equipe elegível; constam no Excel geral.")
    st.dataframe(
        distribuicao_geral[[c for c in ["NOTA", "TIPO_OBRA", "BASE_ATRIBUIDA", "ATIVIDADE_EQUIPE", "CIDADE_BASE", "DISTANCIA_EQUIPE_KM", "METODO_DISTANCIA", "STATUS_DISTRIBUICAO"] if c in distribuicao_geral.columns]],
        hide_index=True,
        use_container_width=True,
    )


# ============================================================
# DOWNLOADS GERAIS
# ============================================================

first, second, third, fourth = (
    st.columns(4)
)

with first:
    st.download_button(
        "📊 Excel geral",
        book,
        "NIP_Analise.xlsx",
        use_container_width=True
    )

with second:
    st.download_button(
        "🗺️ KML TODAS AS OBRAS PENDENTES",
        kml,
        "NIP_Obras.kml",
        use_container_width=True
    )

with third:
    st.download_button(
        "📦 ZIP Excel equipes (Lista Contínua)",
        zip_excel_equipes,
        "NIP_Excel_Equipes.zip",
        use_container_width=True
    )

with fourth:
    st.download_button(
        "📦 ZIP KML equipes (Lista Contínua)",
        zip_kml_equipes,
        "NIP_KML_Equipes.zip",
        use_container_width=True
    )

# Este download respeita filtros e programacao; o KML geral acima nao.
st.download_button(
    "🗺️ KML filtrado / programado",
    kml_filtrado,
    "NIP_Obras_Filtradas.kml",
    mime="application/vnd.google-earth.kml+xml",
    use_container_width=True
)


# ============================================================
# DOWNLOADS ESPECIFICOS
# ============================================================

c1, c2, c3, c4 = (
    st.columns(4)
)

with c1:
    st.download_button(
        "Duplicadas Excel",
        excel_bytes({
            "DUPLICADAS": dup
        }),
        "Duplicadas.xlsx",
        use_container_width=True
    )

with c2:
    st.download_button(
        "FINL Excel",
        excel_bytes({
            "FINL": finl
        }),
        "FINL.xlsx",
        use_container_width=True
    )

with c3:
    st.download_button(
        "CANC Excel",
        excel_bytes({
            "CANC": canc
        }),
        "CANC.xlsx",
        use_container_width=True
    )

with c4:
    st.download_button(
        "Já em campo Excel",
        excel_bytes({
            "JA EM CAMPO": field
        }),
        "Em_Campo.xlsx",
        use_container_width=True
    )

# Também inclui a relação de equipes no KML exclusivo de duplicadas.
dup_com_equipes, _ = team_candidates(
    dup,
    teams,
    k
)
st.download_button(
    "🟣 KML de duplicadas",
    kml_bytes(
        dup_com_equipes,
        name="NIP - Duplicadas"
    ),
    "NIP_Duplicadas.kml"
)


# ============================================================
# RODAPE
# ============================================================

st.caption(
    "Distâncias de alocação exigem confirmação "
    "OSRM. A programação é sugestiva, considera "
    "dias úteis (sem feriados) e requer "
    "revisão operacional."
)
