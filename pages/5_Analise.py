# -*- coding: utf-8 -*-
"""NIP | Analise cruzada V3: auditoria de coordenadas por nota e por registro.

Instalacao: pip install streamlit pandas numpy openpyxl folium streamlit-folium scikit-learn
Execucao: streamlit run NIP_Analise_Cruzada_Equipes_V2.py

As distancias sao GEOGRAFICAS (Haversine) entre a coordenada da obra e a
coordenada informada para a base da equipe. Nao sao quilometros por estrada.
"""
import base64
import hashlib
import html
import io
import math
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from collections import defaultdict

import numpy as np
import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from sklearn.cluster import DBSCAN
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import Workbook
import xlsxwriter
from openpyxl.utils import get_column_letter

EARTH_KM = 6371.0088
CATEGORIES = ("SANEAMENTO", "LEVANTAMENTO")
NAVY = "002060"
EXCEL_COLORS = {"DUPLICADAS": "EADCF6", "SEM EQUIPE PROXIMA": "FCE5E5", "MULTIPLAS EQUIPES": "FFF0C7", "SUPER_PONTO": "FCE4D6"}
# Icone KMZ incorporado para nao depender de arquivo externo na implantacao.
# Abaixo o PNG e um icone de casa compacto gerado em tempo de empacotamento
# quando a CASA 3D original nao estiver ao lado do modulo.
HOME_ICON_REMOTE = "https://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"
HOME_PNG_EMBEDDED = "iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFxUlEQVR4nO3YW0xTdxwH8G7Zlj2YJcuSJXtYsizLnvaypz0pGVCgpdALl1JKoSAthVKgDAUUoUMCIqBQQMcg6EBAaLi0KBMEuSOXiheUocigAjrinM5JuCh8l3McpECRFjiAC7/k+9J/L79Pz/md/8mh0XZrt3Zrt7aqyuLiPqC9bZWVVbYnNfVnRY4qtVETGjVRGhg+rXMTID0yvoe2UyspqejjxabDoh+3iiWvRjhcPHR2xl2eCyoF3uixs4OeTkenvQPiE7OjaTvpH69VpVTrPUXzY2w22bRxRp3ZqPLyRTvTkQQYp5brOp+YlvvVdhtolak5B/v8pNPLm19Ii89+1POFKwDGKfQLmtw2gCb7rF1n6MGJ1QC3xP6oDVC8EWCctMOJLVsKKDl55ouGI4lXR7k8k4AhvgCXQg6gy4FhNkJPp6PDgYGE5NNBlAMaGxvfu5ikzhkUCOdMAcbZbDQG/4A2rqtFAL1Ralzc549ml3xOGaIqJUd5yz9warXT6KosBFdEvusG6I2SL1M+3XSAJqvgu7YDR4ZWA/R7+eBSkHJTAHrjeYk51rApgIJTBZ/Wxh5reuDiahIw7OqG2pCD6DJxOd2MtDOYiE/O8V83IC4u7t1L8WlZ94SiV8aNj3G4uO0fiB5lNHoPxKBOIqcEoDdKtZtg7lBO+WcWIyoz8kS9gSHPRzlc3PGTQq+MwvWIGPyuCMe0lwjg8xczEhCMTkcW5Zg8ecRjixBtCWn9fRGHMS5TYE4gWNK0cSbFfhg4FI9pbx90KKMoh+jpdCTHplSZDVmt6fuRsWhQn0F+RTNU3aMIGJmH1ADkatvJ91yWh1MOaWWyEJucyzcL8iQgaHYgWjXanZReq667jojbT8mG35SOlFOYFHmjSSCiHKPlC+eIm1GLTrO1AAtRDE7joSIcw7Jgyq5eeqP8pIh8RAmEiKrrAWZE3mileF6andjIdfEAZRAi50ovv56XwNBNB3Q7MFAWEIIKhiOq9+6lFkKk81gm/vEWo4XvuXlzIZaijOdOAoicFQc8phyiuDeFP4KVGJSHWXzXq1+Weg8hisUBi4DzXPcZtSr1R4sQyyH+92fMxsR3GjAr9EKzMnpdgDZnDopkodB+b00CKuwZUB05jqy4rD0WI5ZD6Dnl8GgehGx4zixMaVEN5j08UC8LMX8O6HRopMEoZ7JeHwUrK6hDosBOL4S1utiyuVgNQnwREeIym1F3Y0nTomsTKyDEhtkbn4q/fXzR4i5YE1EtluI8X7h4GuX7yiA6cXbxdzcdsrDbL19jlTbDu39yyethAy/wpywYvynC0WNvbxLQ5OqOIn85dPv2kYAyZy4iEjJhk1G0BLFlECLM/jk4dU1g/9Ds4lpi6yBeegrRFLZ0f7nKdESxLAxaG1sSUGVLR8LhJDioVwK2HjIAMoybU0vWC4+ffj0vEjn5HKtcIoeGzSMBxJHIlkfA5WTBqoBtgxAxXidOkxtRcXjiux8lAp/FOfjFyw9+aWfWBFhvFcQmuwysoitwquwGs/f5CgjxOdfMYvwlkZEADYuN6KMZZgOsNwwZQaLUgKmFZmyyzuOl0GsFhHehF7yqHvC0enCrusG6cNPkEYtUFyIpKgGMN8yBtYkQv+usaemgbaQCDfhSYkA581z9OEfXa1WTXWR1N0o1vvxyK+57RiIIDBFTkPWEUVj3yK2m14pGVa22EXp2GMAp79wwxC5PN8PT9cRQBlgLQmbZ7m8JwPa0Zp5d0V7h0zj8IeWI/yAZkhG8NOd2xaw5yCwBq+RKn5P22tY/jfcfw9dSAy5sFMIs/PUJ7+J1Dm27S2KArcSAO5ZC7PK0sxxt1wnaTiqpHu9LDAiVGvBsLQix7zhpWpuENfc+ou3Uko/hE2J+pCN4tQJCzkHDELf6xje0t6Vkw/hWMoLmBYhD/sUXbJ1+/c9wt7WAd6QG8Lm6rnQa8O72NrNbu7VbtP9D/QvGWFY6Zjt3KQAAAABJRU5ErkJggg=="


def norm(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    t = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", t.strip().upper())


def ident(value):
    t = str(value).strip() if value is not None and not pd.isna(value) else ""
    if re.fullmatch(r"\d+\.0+", t):
        t = t.split(".")[0]
    if norm(t) in ("", "NAN", "NULL", "NONE", "0", "-"):
        return ""
    return t


def tidy(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return "" if norm(s) in ("NAN", "NONE", "NULL") else s


def find_col(df, options, required=False):
    names = {norm(x): x for x in df.columns}
    for opt in options:
        if norm(opt) in names:
            return names[norm(opt)]
    if required:
        raise ValueError("Coluna obrigatoria nao encontrada: " + " / ".join(options))
    return None


def get_col(df, options, fallback=""):
    k = find_col(df, options)
    return df[k].fillna("").astype(str).map(tidy) if k else pd.Series(fallback, index=df.index, dtype="object")


def nums(values):
    return pd.to_numeric(values.astype(str).str.strip().str.replace(",", ".", regex=False), errors="coerce")


def normalize_coords(lat, lon):
    la, lo = nums(lat), nums(lon)
    # Troca somente quando o par esta claramente invertido para a geografia brasileira.
    sw = la.between(-75, -34) & lo.between(-34, 6)
    la0, lo0 = la.copy(), lo.copy()
    la.loc[sw] = lo0.loc[sw]
    lo.loc[sw] = la0.loc[sw]
    valid = la.between(-35, 6) & lo.between(-75, -30) & la.ne(0) & lo.ne(0)
    return la.where(valid), lo.where(valid), sw.fillna(False)


def diagnosticar_par_coordenadas(latitude, longitude):
    """Classifica o problema ORIGINAL sem inventar localizacoes nem arredondar 0.

    Utilizado para indicar o motivo da rejeicao de coordenadas em cada base.
    """
    def converter(valor):
        texto = tidy(valor)
        if norm(texto) in ("", "-", "NAN", "NONE", "NULL"):
            return None, "VAZIO"
        try:
            numero = float(texto.replace(",", "."))
        except (ValueError, TypeError):
            return None, "FORMATO INVALIDO"
        if not math.isfinite(numero):
            return None, "FORMATO INVALIDO"
        return numero, "OK"

    la, status_la = converter(latitude)
    lo, status_lo = converter(longitude)
    if status_la == "VAZIO" and status_lo == "VAZIO":
        return "LATITUDE E LONGITUDE VAZIAS"
    if status_la == "VAZIO":
        return "LATITUDE VAZIA"
    if status_lo == "VAZIO":
        return "LONGITUDE VAZIA"
    if status_la != "OK" or status_lo != "OK":
        return "COORDENADA NAO NUMERICA"
    if la == 0 and lo == 0:
        return "LATITUDE E LONGITUDE ZERADAS"
    if la == 0:
        return "LATITUDE ZERADA"
    if lo == 0:
        return "LONGITUDE ZERADA"
    if (-35 <= la <= 6 and -75 <= lo <= -30) or (-75 <= la <= -34 and -34 <= lo <= 6):
        return "VALIDAS"
    return "FORA DA FAIXA GEOGRAFICA DO BRASIL"


def diagnosticar_origem(lat_final, lon_final, lat_original, lon_original, lat_campo, lon_campo, fonte):
    """Retorna situacao, motivo e fonte de cada registro, inclusive fallback de campo."""
    status, motivo, fontes = [], [], []
    for la, lo, lr, lor, lc, loc in zip(lat_final, lon_final,
                                         lat_original, lon_original, lat_campo, lon_campo):
        primary = diagnosticar_par_coordenadas(lr, lor)
        campo = diagnosticar_par_coordenadas(lc, loc) if fonte == "SANEAMENTO" else "NAO INFORMADAS"
        if pd.notna(la) and pd.notna(lo):
            status.append("VALIDAS")
            motivo.append("-")
            fontes.append("CAMPO" if primary != "VALIDAS" and campo == "VALIDAS" else "PLANILHA")
        else:
            status.append("SEM COORDENADAS")
            if fonte == "SANEAMENTO" and campo not in ("NAO INFORMADAS", "LATITUDE E LONGITUDE VAZIAS"):
                motivo.append(f"PROJETO: {primary} | CAMPO: {campo}")
            else:
                motivo.append(primary)
            fontes.append("NAO DISPONIVEL")
    return status, motivo, fontes


def workbook_from_bytes(contents, prefer_sheets, mandatory):
    if isinstance(contents, (str, Path)):
        contents = Path(contents).read_bytes()
    xls = pd.ExcelFile(io.BytesIO(contents), engine="openpyxl")
    order = sorted(xls.sheet_names, key=lambda sh: (0 if norm(sh) in {norm(n) for n in prefer_sheets} else 1))
    for sh in order:
        for header_row in range(4):
            preview = pd.read_excel(xls, sheet_name=sh, header=header_row, nrows=2, dtype=str)
            cols = {norm(c) for c in preview.columns}
            if all(any(norm(n) in cols for n in alternatives) for alternatives in mandatory):
                return pd.read_excel(xls, sheet_name=sh, header=header_row, dtype=str), sh
    raise ValueError(f"Aba com as colunas exigidas nao encontrada. Abas: {xls.sheet_names}")


def read_works(contents, origin):
    if origin == "SANEAMENTO":
        frame, sheet = workbook_from_bytes(contents, ["Clientes Existentes"], [["NOTA"], ["LATITUDE PROJETO", "LATITUDE"], ["LONGITUDE PROJETO", "LONGITUDE"]])
        key = get_col(frame, ["NOTA"])
        lat_raw = get_col(frame, ["LATITUDE PROJETO", "LATITUDE"])
        lon_raw = get_col(frame, ["LONGITUDE PROJETO", "LONGITUDE"])
        lat_campo_raw, lon_campo_raw = get_col(frame, ["LATITUDE CAMPO"]), get_col(frame, ["LONGITUDE CAMPO"])
        la1, lo1, sw1 = normalize_coords(lat_raw, lon_raw)
        if find_col(frame, ["LATITUDE CAMPO"]) and find_col(frame, ["LONGITUDE CAMPO"]):
            la2, lo2, sw2 = normalize_coords(lat_campo_raw, lon_campo_raw)
            fall = la1.isna() & la2.notna()
            la1.loc[fall], lo1.loc[fall] = la2.loc[fall], lo2.loc[fall]
            sw1 = sw1 | (sw2 & fall)
        status = get_col(frame, ["STATUS"])
        prior = pd.Series("", index=frame.index)
        sisco = get_col(frame, ["NOTA PROJ SISGB", "ID SISCO"])
        field_team = get_col(frame, ["CODIGO EQUIPE", "CÓDIGO EQUIPE"])
    else:
        frame, sheet = workbook_from_bytes(contents, ["OBRAS", "NOTAS"], [["PROTOCOLO"], ["LATITUDE"], ["LONGITUDE"]])
        key = get_col(frame, ["PROTOCOLO"])
        lat_raw, lon_raw = get_col(frame, ["LATITUDE"]), get_col(frame, ["LONGITUDE"])
        lat_campo_raw = pd.Series("", index=frame.index)
        lon_campo_raw = pd.Series("", index=frame.index)
        la1, lo1, sw1 = normalize_coords(lat_raw, lon_raw)
        status = get_col(frame, ["STATUS LIST", "STATUS SISCO", "STATUS SAP"])
        prior = get_col(frame, ["PRIORIDADE"])
        sisco = get_col(frame, ["ID SISCO", "ID ISCO"])
        field_team = get_col(frame, ["LEVANTADOR"])
    situacao, motivo, fonte_coord = diagnosticar_origem(
        la1, lo1, lat_raw, lon_raw, lat_campo_raw, lon_campo_raw, origin)
    result = pd.DataFrame({
        "NOTA": key.map(ident), "ORIGEM": origin, "ABA_ORIGEM": sheet,
        "LINHA_ORIGINAL": np.arange(2, len(frame) + 2),
        "NOME_CLIENTE": get_col(frame, ["NOME", "NOME DO CLIENTE", "CLIENTE"]),
        "ID_SISCO_OU_SISGB": sisco,
        "MUNICIPIO": get_col(frame, ["MUNICIPIO", "MUNICÍPIO", "CIDADE"]),
        "REGIONAL": get_col(frame, ["REGIONAL", "REGIONAL SERVICO"]),
        "ENDERECO": get_col(frame, ["ENDEREÇO", "ENDERECO"]),
        "LOCALIDADE": get_col(frame, ["LOCALIDADE", "BAIRRO"]),
        "STATUS_ORIGEM": status, "PRIORIDADE_ORIGEM": prior,
        "COLABORADOR_BASE": field_team,
        "LATITUDE": la1, "LONGITUDE": lo1,
        "LATITUDE_ORIGINAL": lat_raw, "LONGITUDE_ORIGINAL": lon_raw,
        "LATITUDE_CAMPO_ORIGINAL": lat_campo_raw, "LONGITUDE_CAMPO_ORIGINAL": lon_campo_raw,
        "SITUACAO_COORDENADAS_ORIGEM": situacao,
        "MOTIVO_SEM_COORDENADAS_ORIGEM": motivo,
        "FONTE_COORDENADAS_ORIGEM": fonte_coord,
        "COORDENADAS_INVERTIDAS_CORRIGIDAS": np.where(sw1, "SIM", "NÃO"),
    })
    result["CHAVE_MUNICIPIO"] = result["MUNICIPIO"].map(norm)
    return result, sheet


def read_teams(contents):
    if isinstance(contents, (str, Path)):
        contents = Path(contents).read_bytes()
    xls = pd.ExcelFile(io.BytesIO(contents), engine="openpyxl")
    parts = []
    for typ, options in [("SANEAMENTO", ["SANEAMENTO"]), ("LEVANTAMENTO", ["LEVANTADORES", "LEVANTAMENTO"])]:
        matched = next((name for name in xls.sheet_names if norm(name) in {norm(x) for x in options}), None)
        if not matched:
            raise ValueError(f"Falta aba de equipes de {typ}. Abas encontradas: {xls.sheet_names}")
        raw = pd.read_excel(xls, sheet_name=matched, dtype=str)
        name = get_col(raw, ["NOME", "EQUIPE", "COLABORADOR"])
        city = get_col(raw, ["CIDADES", "CIDADE", "MUNICIPIO", "MUNICÍPIO"])
        lat, lon, _ = normalize_coords(get_col(raw, ["LATITUDE", "LAT"]), get_col(raw, ["LONGITUDE", "LON"]))
        d = pd.DataFrame({"TIPO_EQUIPE": typ, "COLABORADOR": name, "MUNICIPIO_BASE": city, "LAT_BASE": lat, "LON_BASE": lon})
        d = d[d["COLABORADOR"].map(norm).ne("")].copy()
        d["CHAVE_MUNICIPIO"] = d["MUNICIPIO_BASE"].map(norm)
        d["ID_EQUIPE"] = d["TIPO_EQUIPE"] + "|" + d["COLABORADOR"].map(norm) + "|" + d["MUNICIPIO_BASE"].map(norm)
        parts.append(d)
    out = pd.concat(parts, ignore_index=True).drop_duplicates("ID_EQUIPE").reset_index(drop=True)
    return out


def nonempty_first(group, col):
    seq = group[col].map(tidy)
    valid = seq[seq.ne("")]
    return valid.iloc[0] if len(valid) else ""


def consolidate(san, pla):
    valid = pd.concat([san, pla], ignore_index=True)
    missing = valid[valid["NOTA"].eq("")].copy()  # identificadores vazios preservados separadamente
    valid = valid[valid["NOTA"].ne("")].copy()
    counts = valid.groupby(["NOTA", "ORIGEM"]).size().unstack(fill_value=0)
    result, conflicts = [], []
    for nota_id, group in valid.groupby("NOTA", sort=False):
        sa = group[group.ORIGEM == "SANEAMENTO"]
        pl = group[group.ORIGEM == "LEVANTAMENTO"]
        firsts = pd.concat([pl, sa], ignore_index=True)  # preferir OBRAS PLA, se houver
        geom = firsts[firsts.LATITUDE.notna() & firsts.LONGITUDE.notna()]
        src = geom.iloc[0] if len(geom) else firsts.iloc[0]
        existing = ("SANEAMENTO", "LEVANTAMENTO") if len(sa) and len(pl) else (("SANEAMENTO",) if len(sa) else ("LEVANTAMENTO",))
        origens_invalidas = firsts[firsts["SITUACAO_COORDENADAS_ORIGEM"].eq("SEM COORDENADAS")]
        faltando = pd.isna(src.LATITUDE) or pd.isna(src.LONGITUDE)
        razoes = " | ".join(dict.fromkeys(
            f"{r.ORIGEM} (linha {r.LINHA_ORIGINAL}): {r.MOTIVO_SEM_COORDENADAS_ORIGEM}"
            for r in origens_invalidas.itertuples()))
        row = {
            "NOTA": nota_id, "ORIGEM": " + ".join(existing),
            "DUPLICADA_ENTRE_BASES": "SIM" if len(existing) == 2 else "NÃO",
            "OCORRENCIAS_SANEAMENTO": len(sa), "OCORRENCIAS_PLA": len(pl),
            "NOME_CLIENTE": nonempty_first(firsts,"NOME_CLIENTE"),
            "ID_SISCO_OU_SISGB": nonempty_first(firsts,"ID_SISCO_OU_SISGB"),
            "MUNICIPIO": tidy(src.MUNICIPIO) or nonempty_first(firsts,"MUNICIPIO"),
            "REGIONAL": tidy(src.REGIONAL) or nonempty_first(firsts,"REGIONAL"),
            "ENDERECO": nonempty_first(firsts,"ENDERECO"), "LOCALIDADE": nonempty_first(firsts,"LOCALIDADE"),
            "LATITUDE": src.LATITUDE, "LONGITUDE": src.LONGITUDE,
            "FONTE_COORDENADAS": src.ORIGEM if pd.notna(src.LATITUDE) else "SEM COORDENADAS",
            "SITUACAO_COORDENADAS": "SEM COORDENADAS" if faltando else ("VALIDAS - PENDENCIA NA ORIGEM" if len(origens_invalidas) else "VALIDAS"),
            "MOTIVO_SEM_COORDENADAS": razoes if len(origens_invalidas) else "-",
            "REGISTROS_ORIGEM_SEM_COORDENADAS": len(origens_invalidas),
            "ACAO_COORDENADAS": "CORRIGIR/INFORMAR COORDENADAS NA BASE" if faltando else ("REVISAR REGISTROS DE ORIGEM" if len(origens_invalidas) else "SEM ACAO"),
            "STATUS_SANEAMENTO": " | ".join(dict.fromkeys(sa.STATUS_ORIGEM.map(tidy))) if len(sa) else "",
            "STATUS_PLA": " | ".join(dict.fromkeys(pl.STATUS_ORIGEM.map(tidy))) if len(pl) else "",
            "PRIORIDADE": nonempty_first(pl,"PRIORIDADE_ORIGEM") if len(pl) else "",
            "COLABORADOR_PLANILHA_PLA": nonempty_first(pl,"COLABORADOR_BASE") if len(pl) else "",
            "CODIGO_EQUIPE_SANEAMENTO": nonempty_first(sa,"COLABORADOR_BASE") if len(sa) else "",
            "LINHAS_SANEAMENTO": ", ".join(sa.LINHA_ORIGINAL.astype(str)) if len(sa) else "",
            "LINHAS_PLA": ", ".join(pl.LINHA_ORIGINAL.astype(str)) if len(pl) else "",
        }
        row["CHAVE_MUNICIPIO"] = norm(row["MUNICIPIO"])
        result.append(row)
        if len(sa) and len(pl):
            asa = sa.dropna(subset=["LATITUDE","LONGITUDE"])
            apl = pl.dropna(subset=["LATITUDE","LONGITUDE"])
            if len(asa) and len(apl):
                diff = float(haversine_km(float(asa.iloc[0].LATITUDE),float(asa.iloc[0].LONGITUDE),float(apl.iloc[0].LATITUDE),float(apl.iloc[0].LONGITUDE)))
                if diff > .1:
                    conflicts.append({"NOTA": nota_id,"TIPO": "COORDENADAS DIFERENTES ENTRE BASES", "DISTANCIA_ENTRE_COORDENADAS_KM": round(diff,2),"DETALHE": "PLA priorizada no mapa"})
            m1 = norm(nonempty_first(sa,"MUNICIPIO"))
            m2 = norm(nonempty_first(pl,"MUNICIPIO"))
            if m1 and m2 and m1 != m2:
                conflicts.append({"NOTA": nota_id, "TIPO": "MUNICIPIOS DIVERGENTES", "DISTANCIA_ENTRE_COORDENADAS_KM": np.nan,"DETALHE":f"SAN: {m1} | PLA: {m2}"})
    for nota_id,row in counts.iterrows():
        for typ, val in row.items():
            if int(val)>1:
                conflicts.append({"NOTA":nota_id,"TIPO":f"REPETICAO INTERNA {typ}","DISTANCIA_ENTRE_COORDENADAS_KM":np.nan,"DETALHE":f"{val} registros"})
    return pd.DataFrame(result), pd.DataFrame(conflicts,columns=["NOTA","TIPO","DISTANCIA_ENTRE_COORDENADAS_KM","DETALHE"]), missing


def haversine_km(lat1, lon1, lat2, lon2):
    a1, b1, a2, b2 = np.radians(lat1), np.radians(lon1), np.radians(lat2), np.radians(lon2)
    dlat, dlon = a2 - a1, b2 - b1
    t = np.sin(dlat/2)**2+np.cos(a1)*np.cos(a2)*np.sin(dlon/2)**2
    return EARTH_KM*2*np.arcsin(np.sqrt(np.clip(t,0,1)))


def add_superpoints(works, radius_m=50):
    result = works.copy()
    result["SUPER_PONTO"] = "NÃO"
    result["ID_SUPER_PONTO"] = ""
    good = result.LATITUDE.notna() & result.LONGITUDE.notna()
    if good.any():
        indexes = result.index[good]
        labels = DBSCAN(eps=max(1,radius_m)/1000/EARTH_KM,min_samples=1,algorithm="ball_tree",metric="haversine").fit_predict(np.radians(result.loc[indexes,["LATITUDE","LONGITUDE"]].to_numpy(float)))
        count = pd.Series(labels).value_counts()
        result.loc[indexes,"ID_SUPER_PONTO"] = [f"SP-{i+1:05d}" for i in labels]
        result.loc[indexes,"SUPER_PONTO"] = [f"SIM ({count[i]} obras)" if count[i]>1 else "NÃO" for i in labels]
    return result


def assess_coverage(works, teams, radius_km=50):
    """Uma nota mestre; uma linha por atividade aplicavel; candidatos de todas as equipes no raio.

    Um trabalho compartilhado origina duas linhas de atividade, mas uma so nota no mestre.
    A sugestao nunca e confundida com alocacao operacional confirmada.
    """
    team_sets = {}
    for typ in CATEGORIES:
        team_sets[typ] = teams[(teams.TIPO_EQUIPE==typ)&teams.LAT_BASE.notna()&teams.LON_BASE.notna()].copy().reset_index(drop=True)
    activity_rows, near_rows, suggestion_rows = [], [], []
    rows_by_note = defaultdict(list)
    for _, work in works.iterrows():
        for typ in CATEGORIES:
            if typ not in work.ORIGEM.split(" + "):
                continue
            t = team_sets[typ]
            record = {
                "NOTA":work.NOTA, "ATIVIDADE":typ,"NOME_CLIENTE":work.NOME_CLIENTE,
                "MUNICIPIO_OBRA":work.MUNICIPIO,"REGIONAL":work.REGIONAL,
                "LATITUDE":work.LATITUDE,"LONGITUDE":work.LONGITUDE,
                "SITUACAO_COORDENADAS":work.SITUACAO_COORDENADAS,
                "MOTIVO_SEM_COORDENADAS":work.MOTIVO_SEM_COORDENADAS,
                "DUPLICADA_ENTRE_BASES":work.DUPLICADA_ENTRE_BASES,
                "EQUIPE_MAIS_PROXIMA":"","TIPO_EQUIPE_MAIS_PROXIMA":typ,
                "MUNICIPIO_EQUIPE_MAIS_PROXIMA":"", "DISTANCIA_MAIS_PROXIMA_KM":np.nan,
                "QUANTIDADE_EQUIPES_PROXIMAS":0,"QUANTIDADE_EQUIPES_NO_MUNICIPIO":0,
                "EQUIPES_PROXIMAS":"","EQUIPES_NO_MUNICIPIO":"",
                "COBERTURA":"","SUGESTAO_FORA_RAIO":"NÃO",
                "METODO_DISTANCIA":"LINHA RETA (HAVERSINE)",
            }
            if pd.isna(work.LATITUDE) or pd.isna(work.LONGITUDE):
                record["COBERTURA"] = "SEM COORDENADAS DA OBRA"
            elif len(t)==0:
                record["COBERTURA"] = "SEM EQUIPE DA ATIVIDADE COM COORDENADAS"
            else:
                kms = haversine_km(float(work.LATITUDE),float(work.LONGITUDE),t.LAT_BASE.to_numpy(float),t.LON_BASE.to_numpy(float))
                ranked = t.copy()
                ranked["DISTANCIA_KM"] = kms
                ranked = ranked.sort_values(["DISTANCIA_KM","COLABORADOR"],kind="stable")
                close = ranked[ranked.DISTANCIA_KM.le(radius_km+1e-9)]
                in_city = ranked[ranked.CHAVE_MUNICIPIO.eq(work.CHAVE_MUNICIPIO)] if work.CHAVE_MUNICIPIO else ranked.iloc[:0]
                nearest = ranked.iloc[0]
                record.update({
                    "EQUIPE_MAIS_PROXIMA":nearest.COLABORADOR,
                    "MUNICIPIO_EQUIPE_MAIS_PROXIMA":nearest.MUNICIPIO_BASE,
                    "DISTANCIA_MAIS_PROXIMA_KM":round(float(nearest.DISTANCIA_KM),2),
                    "QUANTIDADE_EQUIPES_PROXIMAS":len(close),
                    "QUANTIDADE_EQUIPES_NO_MUNICIPIO":len(in_city),
                    "EQUIPES_NO_MUNICIPIO":" | ".join(in_city.COLABORADOR),
                    "EQUIPES_PROXIMAS":" | ".join(f"{c.COLABORADOR} ({c.DISTANCIA_KM:.1f} km)" for c in close.itertuples()),
                    "COBERTURA":"SEM EQUIPE PROXIMA" if len(close)==0 else "DUAS OU MAIS EQUIPES PROXIMAS" if len(close)>=2 else "UMA EQUIPE PROXIMA",
                    "SUGESTAO_FORA_RAIO":"SIM" if len(close)==0 else "NÃO",
                })
                if len(close)==0:
                    suggestion_rows.append({"NOTA":work.NOTA,"ATIVIDADE":typ,"MUNICIPIO_OBRA":work.MUNICIPIO,
                        "COLABORADOR_SUGERIDO":nearest.COLABORADOR,"MUNICIPIO_BASE":nearest.MUNICIPIO_BASE,
                        "DISTANCIA_GEOGRAFICA_KM":round(float(nearest.DISTANCIA_KM),2),
                        "RAIO_ADOTADO_KM":radius_km,"STATUS":"SUGESTAO FORA DO RAIO - NAO ALOCADA"})
                for rank, cand in enumerate(close.itertuples(),1):
                    near_rows.append({"NOTA":work.NOTA,"ATIVIDADE":typ,"RANK_PROXIMIDADE":rank,
                        "COLABORADOR":cand.COLABORADOR,"ID_EQUIPE":cand.ID_EQUIPE,
                        "MUNICIPIO_BASE":cand.MUNICIPIO_BASE,"MUNICIPIO_OBRA":work.MUNICIPIO,
                        "MESMO_MUNICIPIO":"SIM" if cand.CHAVE_MUNICIPIO == work.CHAVE_MUNICIPIO else "NÃO",
                        "DISTANCIA_GEOGRAFICA_KM":round(float(cand.DISTANCIA_KM),2),"LATITUDE":work.LATITUDE,"LONGITUDE":work.LONGITUDE})
            activity_rows.append(record)
            rows_by_note[work.NOTA].append(record)
    matrix = pd.DataFrame(activity_rows)
    closers = pd.DataFrame(near_rows, columns=["NOTA","ATIVIDADE","RANK_PROXIMIDADE","COLABORADOR","ID_EQUIPE","MUNICIPIO_BASE","MUNICIPIO_OBRA","MESMO_MUNICIPIO","DISTANCIA_GEOGRAFICA_KM","LATITUDE","LONGITUDE"])
    suggestions = pd.DataFrame(suggestion_rows, columns=["NOTA","ATIVIDADE","MUNICIPIO_OBRA","COLABORADOR_SUGERIDO","MUNICIPIO_BASE","DISTANCIA_GEOGRAFICA_KM","RAIO_ADOTADO_KM","STATUS"])
    master = works.copy()
    summaries = []
    for note in master.NOTA:
        sub = rows_by_note[note]
        eligible = [r for r in sub if pd.notna(r["DISTANCIA_MAIS_PROXIMA_KM"])]
        best = min(eligible,key=lambda r:r["DISTANCIA_MAIS_PROXIMA_KM"]) if eligible else None
        findings = {
            "COBERTURA_GERAL":" | ".join(f"{r['ATIVIDADE']}: {r['COBERTURA']}" for r in sub),
            "EQUIPE_MAIS_PROXIMA":best["EQUIPE_MAIS_PROXIMA"] if best else "",
            "TIPO_EQUIPE_MAIS_PROXIMA":best["ATIVIDADE"] if best else "",
            "DISTANCIA_GEOGRAFICA_KM":best["DISTANCIA_MAIS_PROXIMA_KM"] if best else np.nan,
            "TOTAL_EQUIPES_PROXIMAS":sum(r["QUANTIDADE_EQUIPES_PROXIMAS"] for r in sub),
            "EXISTE_ATIVIDADE_SEM_EQUIPE":"SIM" if any(r["COBERTURA"] in ("SEM EQUIPE PROXIMA","SEM EQUIPE DA ATIVIDADE COM COORDENADAS") for r in sub) else "NÃO",
            "EXISTE_ATIVIDADE_COM_MULTIPLAS":"SIM" if any(r["QUANTIDADE_EQUIPES_PROXIMAS"]>=2 for r in sub) else "NÃO",
            "EQUIPE_SANEAMENTO_MAIS_PROXIMA":next((r["EQUIPE_MAIS_PROXIMA"] for r in sub if r["ATIVIDADE"]=="SANEAMENTO"),""),
            "DISTANCIA_SANEAMENTO_KM":next((r["DISTANCIA_MAIS_PROXIMA_KM"] for r in sub if r["ATIVIDADE"]=="SANEAMENTO"),np.nan),
            "EQUIPE_LEVANTAMENTO_MAIS_PROXIMA":next((r["EQUIPE_MAIS_PROXIMA"] for r in sub if r["ATIVIDADE"]=="LEVANTAMENTO"),""),
            "DISTANCIA_LEVANTAMENTO_KM":next((r["DISTANCIA_MAIS_PROXIMA_KM"] for r in sub if r["ATIVIDADE"]=="LEVANTAMENTO"),np.nan),
            "EQUIPES_PROXIMAS_E_TIPOS":" | ".join(f"{r['ATIVIDADE']}: {r['EQUIPES_PROXIMAS']}" for r in sub if r['EQUIPES_PROXIMAS']),
        }
        summaries.append(findings)
    master = pd.concat([master.reset_index(drop=True),pd.DataFrame(summaries)],axis=1)
    return master, matrix, closers, suggestions


def rank_team_summary(teams, matrix, closers, suggestions):
    out = teams.copy()
    # Proximidade e sugestao sao indicadores, nao alocacoes firmes.
    if closers.empty:
        out["OBRAS_NO_RAIO"] = 0;out["OBRAS_COMO_MAIS_PROXIMA"] = 0
    else:
        within = closers.groupby("ID_EQUIPE").NOTA.nunique()
        leader = closers[closers.RANK_PROXIMIDADE.eq(1)].groupby("ID_EQUIPE").NOTA.nunique()
        out["OBRAS_NO_RAIO"] = out.ID_EQUIPE.map(within).fillna(0).astype(int)
        out["OBRAS_COMO_MAIS_PROXIMA"] = out.ID_EQUIPE.map(leader).fillna(0).astype(int)
    if len(suggestions):
        s = suggestions.groupby(["ATIVIDADE","COLABORADOR_SUGERIDO"]).NOTA.nunique()
        out["SUGESTOES_FORA_RAIO"] = [int(s.get((r.TIPO_EQUIPE,r.COLABORADOR),0)) for r in out.itertuples()]
    else:
        out["SUGESTOES_FORA_RAIO"] = 0
    return out


def analysis_from_bytes(san_bytes, pla_bytes, team_bytes, radius_km=50, superpoint_m=50):
    sane, sname = read_works(san_bytes,"SANEAMENTO")
    pla, pname = read_works(pla_bytes,"LEVANTAMENTO")
    teams = read_teams(team_bytes)
    works, problems, empty_ids = consolidate(sane,pla)
    works = add_superpoints(works,superpoint_m)
    works, matrix, near, suggestions = assess_coverage(works,teams,radius_km)
    team_summary = rank_team_summary(teams,matrix,near,suggestions)
    no_team = matrix[matrix.COBERTURA.isin(["SEM EQUIPE PROXIMA","SEM EQUIPE DA ATIVIDADE COM COORDENADAS"])].copy()
    many = matrix[matrix.QUANTIDADE_EQUIPES_PROXIMAS.ge(2)].copy()
    dup = works[works.DUPLICADA_ENTRE_BASES.eq("SIM")].copy()
    no_coords = works[works.SITUACAO_COORDENADAS.eq("SEM COORDENADAS")].copy()
    # Uma linha para cada registro com problema na base original. Se uma nota
    # tiver coordenada valida na outra base, ela continua aqui para saneamento.
    source_issues = pd.concat([sane, pla], ignore_index=True)
    source_issues = source_issues[source_issues.SITUACAO_COORDENADAS_ORIGEM.eq("SEM COORDENADAS")].copy()
    source_issues["ACAO_CORRETIVA"] = "REVISAR OU PREENCHER LATITUDE/LONGITUDE NA BASE DE ORIGEM"
    source_issues = source_issues.sort_values(["ORIGEM","NOTA","LINHA_ORIGINAL"],kind="stable")
    params = pd.DataFrame([
        ["Raio para equipe proxima (km)",radius_km],["Raio superponto (m)",superpoint_m],
        ["Distancias","Haversine em linha reta, nao trajeto rodoviario"],
        ["Local da equipe","Coordenadas da planilha LOCALIDADE; nao GPS atual"],
        ["Duplicidade","Mesmo numero de NOTA (Saneamento) e PROTOCOLO (PLA)"],
        ["Nota compartilhada","Uma linha em OBRAS, duas atividades em ATENDIMENTOS"],
        ["Ausencia de coordenadas", "Sem posicionamento valido em qualquer uma das duas bases"],
        ["Pendencias na origem", "Registros invalidos em SAN/PLA, mesmo que a outra base tenha coordenadas validas"],
        ["Equipe sem raio","A menor distancia aparece como sugestao; nao eh atribuicao"],
        ["Planilhas analisadas",f"Saneamento aba {sname}; PLA aba {pname}"],
        ["Processado em",datetime.now().strftime("%d/%m/%Y %H:%M")],
    ],columns=["PARAMETRO","VALOR"])
    return {"obras":works,"atendimentos":matrix,"proximas":near,"sugestoes":suggestions,
        "sem_equipe":no_team,"multiplas":many,"duplicadas":dup,"sem_coordenadas":no_coords,
        "pendencias_coordenadas_origem":source_issues,
        "divergencias":problems,"sem_id":empty_ids,"equipes":team_summary,"parametros":params,
        "brutas_saneamento":sane,"brutas_pla":pla}


def excel_safe(df):
    clean = df.copy().drop(columns=["CHAVE_MUNICIPIO"],errors="ignore")
    for c in clean.columns:
        if clean[c].dtype == "object":
            clean[c] = clean[c].map(lambda v: ("'"+v if isinstance(v,str) and v.startswith(("=","+","-","@")) else v))
    return clean


def make_excel(tables):
    """Modelo da Lista Continua com Excel rapido e formatacao padronizada."""
    out=io.BytesIO()
    wb=xlsxwriter.Workbook(out,{"in_memory":True,"strings_to_urls":False})
    head=wb.add_format({"bold":True,"font_name":"Calibri","font_size":10,"font_color":"white","bg_color":"#002060","align":"center","valign":"vcenter","text_wrap":True})
    number_fmt=wb.add_format({"font_name":"Calibri","font_size":10,"num_format":"0.00"})
    fill_p=wb.add_format({"bg_color":"#EEE5F4"})
    fill_r=wb.add_format({"bg_color":"#FFE8E8"})
    fill_y=wb.add_format({"bg_color":"#FFF2D0"})
    fill_coord=wb.add_format({"bg_color":"#FCE8E6","font_color":"#9C0006"})
    fill_review=wb.add_format({"bg_color":"#FFF2CC","font_color":"#854F0B"})
    default=None
    for name,raw in tables.items():
        if raw is None:continue
        df=raw.drop(columns=["CHAVE_MUNICIPIO"],errors="ignore")
        ws=wb.add_worksheet(name[:31])
        cols=list(df.columns)
        ws.hide_gridlines(2)
        ws.set_tab_color("#C62828" if name in ("SEM COORDENADAS", "PENDENCIAS NA ORIGEM") else "#002060")
        ws.freeze_panes(1,3 if len(cols)>=4 else 0)
        ws.set_row(0,35)
        widths=[min(30,max(13,len(str(c))+2)) for c in cols]
        for i,c in enumerate(cols):ws.write(0,i,str(c),head)
        idxdup=cols.index("DUPLICADA_ENTRE_BASES") if "DUPLICADA_ENTRE_BASES" in cols else -1
        idxstatus=cols.index("COBERTURA") if "COBERTURA" in cols else -1
        idxno=cols.index("EXISTE_ATIVIDADE_SEM_EQUIPE") if "EXISTE_ATIVIDADE_SEM_EQUIPE" in cols else -1
        idxcoord=cols.index("SITUACAO_COORDENADAS") if "SITUACAO_COORDENADAS" in cols else (cols.index("SITUACAO_COORDENADAS_ORIGEM") if "SITUACAO_COORDENADAS_ORIGEM" in cols else -1)
        kmcols=[("KM" in norm(c) or "DISTANCIA" in norm(c)) for c in cols]
        for ir,values in enumerate(df.itertuples(index=False,name=None),start=1):
            shade=None
            if idxdup>=0 and values[idxdup]=="SIM":shade=fill_p
            if (idxstatus>=0 and values[idxstatus]=="SEM EQUIPE PROXIMA") or (idxno>=0 and values[idxno]=="SIM"):shade=fill_r
            elif idxstatus>=0 and values[idxstatus]=="DUAS OU MAIS EQUIPES PROXIMAS":shade=fill_y
            # A ausencia de coordenadas tem prioridade sobre status de equipe
            # e duplicidade: nao sugerir deslocamento para um ponto desconhecido.
            if name in ("SEM COORDENADAS", "PENDENCIAS NA ORIGEM") or (idxcoord>=0 and str(values[idxcoord]).upper()=="SEM COORDENADAS"):
                shade=fill_coord
            elif idxcoord>=0 and "PENDENCIA" in str(values[idxcoord]).upper():
                shade=fill_review
            for i,value in enumerate(values):
                if value is None:continue
                if isinstance(value,(list,dict,set)):value=str(value)
                if isinstance(value,np.integer):value=int(value)
                if isinstance(value,np.floating):value=float(value)
                if isinstance(value,pd.Timestamp):value=value.to_pydatetime()
                try:
                    if pd.isna(value):continue
                except (ValueError,TypeError):pass
                if isinstance(value,str):
                    if len(value)>32700:value=value[:32690]+"..."
                    if value.startswith(("=","+","-","@")):value="'"+value
                if ir<=65:widths[i]=min(52,max(widths[i],len(str(value))+2))
                if isinstance(value,(int,float)) and math.isfinite(float(value)):
                    ws.write_number(ir,i,float(value),shade or (number_fmt if kmcols[i] else default))
                else:
                    ws.write_string(ir,i,str(value),shade)
        if cols:
            ws.autofilter(0,0,len(df),len(cols)-1)
        for i,width in enumerate(widths):ws.set_column(i,i,width)
    wb.close()
    return out.getvalue()


def ordered_master(df):
    front=["NOTA","SITUACAO_COORDENADAS","MOTIVO_SEM_COORDENADAS","ACAO_COORDENADAS","ORIGEM","DUPLICADA_ENTRE_BASES","SUPER_PONTO","NOME_CLIENTE","ID_SISCO_OU_SISGB","REGIONAL","MUNICIPIO","ENDERECO","LOCALIDADE","EQUIPE_MAIS_PROXIMA","TIPO_EQUIPE_MAIS_PROXIMA","DISTANCIA_GEOGRAFICA_KM","EXISTE_ATIVIDADE_SEM_EQUIPE","EXISTE_ATIVIDADE_COM_MULTIPLAS","TOTAL_EQUIPES_PROXIMAS","COBERTURA_GERAL","EQUIPE_SANEAMENTO_MAIS_PROXIMA","DISTANCIA_SANEAMENTO_KM","EQUIPE_LEVANTAMENTO_MAIS_PROXIMA","DISTANCIA_LEVANTAMENTO_KM","EQUIPES_PROXIMAS_E_TIPOS","LATITUDE","LONGITUDE"]
    return df[[c for c in front if c in df.columns]+[c for c in df.columns if c not in front]]


def ordered_missing(df):
    first=["NOTA", "SITUACAO_COORDENADAS", "MOTIVO_SEM_COORDENADAS", "ACAO_COORDENADAS",
           "ORIGEM", "NOME_CLIENTE", "ID_SISCO_OU_SISGB", "MUNICIPIO", "REGIONAL",
           "ENDERECO", "LOCALIDADE", "LINHAS_SANEAMENTO", "LINHAS_PLA", "STATUS_SANEAMENTO",
           "STATUS_PLA", "LATITUDE", "LONGITUDE", "FONTE_COORDENADAS"]
    return df[[c for c in first if c in df.columns]+[c for c in df.columns if c not in first]]


def ordered_source_issues(df):
    first=["NOTA","ORIGEM","ABA_ORIGEM","LINHA_ORIGINAL","SITUACAO_COORDENADAS_ORIGEM",
           "MOTIVO_SEM_COORDENADAS_ORIGEM","ACAO_CORRETIVA","NOME_CLIENTE","MUNICIPIO",
           "ENDERECO","LATITUDE_ORIGINAL","LONGITUDE_ORIGINAL",
           "LATITUDE_CAMPO_ORIGINAL","LONGITUDE_CAMPO_ORIGINAL","STATUS_ORIGEM"]
    return df[[c for c in first if c in df.columns]+[c for c in df.columns if c not in first]]


def make_complete_excel(r):
    ws = r["obras"]
    summary=pd.DataFrame([
        ("Notas unicas",len(ws)),("Saneamento (notas de origem)",int(ws.OCORRENCIAS_SANEAMENTO.gt(0).sum())),
        ("PLA (notas de origem)",int(ws.OCORRENCIAS_PLA.gt(0).sum())),
        ("Duplicadas entre bases",len(r["duplicadas"])),
        ("Atendimentos separados por atividade",len(r["atendimentos"])),
        ("Atividades sem equipe proxima",len(r["sem_equipe"])),
        ("Atividades com duas ou mais equipes",len(r["multiplas"])),
        ("Notas sem coordenadas",len(r["sem_coordenadas"])),
        ("Registros sem coordenadas nas bases originais",len(r["pendencias_coordenadas_origem"])),
        ("Equipes de levantamento",int(r["equipes"].TIPO_EQUIPE.eq("LEVANTAMENTO").sum())),
        ("Equipes de saneamento",int(r["equipes"].TIPO_EQUIPE.eq("SANEAMENTO").sum()))],columns=["INDICADOR","VALOR"])
    return make_excel({"RESUMO":summary,"SEM COORDENADAS":ordered_missing(r["sem_coordenadas"]),
        "PENDENCIAS NA ORIGEM":ordered_source_issues(r["pendencias_coordenadas_origem"]),
        "OBRAS CONSOLIDADAS":ordered_master(ws),"ATENDIMENTOS":r["atendimentos"],
        "DUPLICADAS":ordered_master(r["duplicadas"]),"SEM EQUIPE PERTO":r["sem_equipe"],
        "MULTIPLAS EQUIPES":r["multiplas"],"SUGESTOES FORA DO RAIO":r["sugestoes"],
        "TODAS EQUIPES NO RAIO":r["proximas"],"RESUMO POR EQUIPE":r["equipes"],
        "DIVERGENCIAS":r["divergencias"], "PARAMETROS":r["parametros"]})


def safe_name(name):
    code=hashlib.sha1(str(name).encode("utf-8")).hexdigest()[:6]
    return re.sub(r"[^A-Z0-9_-]","_",norm(name).replace(" ","_"))[:52]+"_"+code


def popup_work(row):
    v=lambda c: html.escape(tidy(row.get(c,"")),quote=True)
    lat,lon=row.get("LATITUDE"),row.get("LONGITUDE")
    link=""
    if pd.notna(lat) and pd.notna(lon):
        url=f"https://www.google.com/maps/search/?api=1&query={float(lat):.7f}%2C{float(lon):.7f}"
        link=f'<a href="{url}" target="_blank" style="display:block;padding:10px;background:#0D256C;color:white;text-align:center;text-decoration:none;border-radius:5px;margin-top:10px">Abrir no Google Maps</a>'
    fields=[("Cliente",v("NOME_CLIENTE")),("ID SISCO/SISGB",v("ID_SISCO_OU_SISGB")),("Origem",v("ORIGEM")),("Municipio",v("MUNICIPIO")),("Regional",v("REGIONAL")),("Endereco",v("ENDERECO")),("Duplicada",v("DUPLICADA_ENTRE_BASES")),("Cobertura",v("COBERTURA_GERAL")),("Equipe mais proxima",v("EQUIPE_MAIS_PROXIMA")),("Tipo",v("TIPO_EQUIPE_MAIS_PROXIMA")),("Distancia geografica",v("DISTANCIA_GEOGRAFICA_KM")+" km"),("Levantamento",v("EQUIPE_LEVANTAMENTO_MAIS_PROXIMA")),("Saneamento",v("EQUIPE_SANEAMENTO_MAIS_PROXIMA")),("Equipes no raio",v("EQUIPES_PROXIMAS_E_TIPOS")),("Superponto",v("SUPER_PONTO"))]
    body="".join(f'<div style="display:grid;grid-template-columns:38% 62%;padding:5px 0;border-bottom:1px solid #ddd"><span style="color:#536273">{k}</span><b style="overflow-wrap:anywhere">{a or "-"}</b></div>' for k,a in fields)
    return f'<div style="font:12px Arial;width:330px;max-height:480px;overflow-y:auto;background:#fff"><div style="background:#0D256C;color:white;padding:14px"><small>NOTA CCS</small><div style="font-size:19px;font-weight:bold">{v("NOTA")}</div></div><div style="padding:10px">{body}{link}</div></div>'


def marker_category(row):
    if row["DUPLICADA_ENTRE_BASES"]=="SIM":return "Duplicadas nas bases"
    if row["EXISTE_ATIVIDADE_SEM_EQUIPE"]=="SIM":return "Sem equipe proxima"
    if row["EXISTE_ATIVIDADE_COM_MULTIPLAS"]=="SIM":return "Duas ou mais equipes proximas"
    return "Obras Saneamento" if row["ORIGEM"]=="SANEAMENTO" else "Obras Levantamento"


def home_png_bytes():
    """Sempre inclui a casinha original dentro do KMZ, inclusive no Streamlit Cloud."""
    return base64.b64decode(HOME_PNG_EMBEDDED)


def build_kml(works, teams, title, included_teams=None, kmz=False):
    ns="http://www.opengis.net/kml/2.2"
    ET.register_namespace("",ns)
    root=ET.Element("{%s}kml"%ns)
    doc=ET.SubElement(root,"{%s}Document"%ns)
    def el(p,tag,value=None):
        x=ET.SubElement(p,"{%s}%s"%(ns,tag))
        if value is not None:x.text=str(value)
        return x
    el(doc,"name",title)
    icons={"Duplicadas nas bases":"purple-blank.png","Sem equipe proxima":"red-blank.png","Duas ou mais equipes proximas":"ylw-blank.png","Obras Saneamento":"blu-blank.png","Obras Levantamento":"grn-blank.png"}
    for n,icon in icons.items():
        style=el(doc,"Style");style.set("id","work_"+str(list(icons).index(n)))
        style_ic=el(style,"IconStyle");el(style_ic,"scale","1.1")
        el(el(style_ic,"Icon"),"href","https://maps.google.com/mapfiles/kml/paddle/"+icon)
        el(el(style,"LabelStyle"),"scale","0.6")
    hstyle=el(doc,"Style");hstyle.set("id","home")
    hicon=el(hstyle,"IconStyle");el(hicon,"scale","1.4")
    el(el(hicon,"Icon"),"href","assets/casinha.png" if kmz else HOME_ICON_REMOTE)
    for cat in icons:
        folder=el(doc,"Folder");el(folder,"name",cat)
        if len(works)==0:continue
        selected=works[works.apply(marker_category,axis=1).eq(cat)]
        for _,w in selected.dropna(subset=["LATITUDE","LONGITUDE"]).iterrows():
            pm=el(folder,"Placemark")
            el(pm,"name",f"{w['NOTA']} - {tidy(w['NOME_CLIENTE'])[:55]}")
            el(pm,"styleUrl","#work_"+str(list(icons).index(cat)))
            el(pm,"description",popup_work(w))
            el(el(pm,"Point"),"coordinates",f"{w['LONGITUDE']},{w['LATITUDE']},0")
    tf=el(doc,"Folder");el(tf,"name","CASINHAS - BASES DAS EQUIPES")
    eqs=teams if included_teams is None else teams[teams.ID_EQUIPE.isin(included_teams)]
    for _,e in eqs.dropna(subset=["LAT_BASE","LON_BASE"]).iterrows():
        pm=el(tf,"Placemark")
        el(pm,"name",f"INICIO DA EQUIPE - {e['COLABORADOR']}")
        el(pm,"styleUrl","#home")
        detail=(f"<b>Equipe:</b> {html.escape(e['COLABORADOR'])}<br><b>Atividade:</b> {html.escape(e['TIPO_EQUIPE'])}<br>"
                f"<b>Municipio-base:</b> {html.escape(e['MUNICIPIO_BASE'])}<br><small>Ponto aproximado informado na planilha LOCALIDADE.</small>")
        el(pm,"description",detail)
        el(el(pm,"Point"),"coordinates",f"{e['LON_BASE']},{e['LAT_BASE']},0")
    return ET.tostring(root,encoding="utf-8",xml_declaration=True)


def wrap_kmz(kml):
    kml_kmz = kml
    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml",kml_kmz)
        z.writestr("assets/casinha.png",home_png_bytes())
    return out.getvalue()


def make_all_exports(r):
    works,teams=r["obras"],r["equipes"]
    whole=make_complete_excel(r)
    kml=build_kml(works,teams,"NIP | Auditoria de Obras - Geral",kmz=False)
    kmz=wrap_kmz(build_kml(works,teams,"NIP | Auditoria de Obras - Geral",kmz=True))
    per_excel=io.BytesIO();per_maps=io.BytesIO()
    with zipfile.ZipFile(per_excel,"w",zipfile.ZIP_DEFLATED) as zex,zipfile.ZipFile(per_maps,"w",zipfile.ZIP_DEFLATED) as zmap:
        for _,e in teams.iterrows():
            eqid=e.ID_EQUIPE
            near=r["proximas"][r["proximas"].ID_EQUIPE.eq(eqid)]
            suggested=r["sugestoes"][(r["sugestoes"].COLABORADOR_SUGERIDO.eq(e.COLABORADOR)) & (r["sugestoes"].ATIVIDADE.eq(e.TIPO_EQUIPE))]
            # Incluir tambem sugestoes fora de raio, visualmente distinguiveis no Excel.
            notes=set(near.NOTA)|set(suggested.NOTA)
            selected=works[works.NOTA.isin(notes)].copy()
            if selected.empty:
                safe=safe_name(eqid)
                subdir=e.TIPO_EQUIPE
                empty_cols=["FISCAL","TIPO_EQUIPE_ALVO","NOTA","NOME_CLIENTE","SITUACAO_EQUIPE","KM_DA_BASE"]
                zex.writestr(f"{subdir}/Rota_{safe}.xlsx",make_excel({"Obras Roteirizadas":pd.DataFrame(columns=empty_cols)}))
                single_team=teams[teams.ID_EQUIPE.eq(eqid)]
                zmap.writestr(f"{subdir}/Rota_{safe}.kml",build_kml(selected,single_team,f"NIP - {e.COLABORADOR}",kmz=False))
                zmap.writestr(f"{subdir}/Rota_{safe}.kmz",wrap_kmz(build_kml(selected,single_team,f"NIP - {e.COLABORADOR}",kmz=True)))
                continue
            team_near=near[near.RANK_PROXIMIDADE.eq(1)].drop_duplicates("NOTA")
            selected["EQUIPE_ALVO"] = e.COLABORADOR
            selected["TIPO_EQUIPE_ALVO"] = e.TIPO_EQUIPE
            selected["SITUACAO_EQUIPE"] = selected.NOTA.map(lambda nt:"PROXIMA NO RAIO" if nt in set(near.NOTA) else "SUGESTAO FORA DO RAIO")
            distances = near.drop_duplicates("NOTA").set_index("NOTA").DISTANCIA_GEOGRAFICA_KM.to_dict()
            distances.update(suggested.set_index("NOTA").DISTANCIA_GEOGRAFICA_KM.to_dict())
            selected["KM_DA_BASE"] = selected.NOTA.map(distances)
            selected=selected.sort_values(["SITUACAO_EQUIPE","KM_DA_BASE","NOTA"],kind="stable")
            selected.insert(0,"ORDEM",np.arange(1,len(selected)+1))
            selected.insert(0,"FISCAL",e.COLABORADOR)
            selected["DIA_SEMANA"]="NÃO PROGRAMADO"
            selected["DIA_MES"]=""
            selected["DISTANCIA_PONTO_ANTERIOR_KM"]="NÃO CALCULADA (SEM ROTA VIARIA)"
            cols=["FISCAL","TIPO_EQUIPE_ALVO","DIA_SEMANA","DIA_MES","SUPER_PONTO","ORDEM","DISTANCIA_PONTO_ANTERIOR_KM","NOTA","SITUACAO_COORDENADAS","NOME_CLIENTE","MUNICIPIO","REGIONAL","SITUACAO_EQUIPE","KM_DA_BASE","DUPLICADA_ENTRE_BASES","ID_SISCO_OU_SISGB","ENDERECO","LOCALIDADE","LATITUDE","LONGITUDE"]
            selected=selected[[x for x in cols if x in selected.columns]+[x for x in selected.columns if x not in cols]]
            safe=safe_name(eqid)
            subdir=e.TIPO_EQUIPE
            zex.writestr(f"{subdir}/Rota_{safe}.xlsx",make_excel({"Obras Roteirizadas":selected}))
            # KML individual contem somente obras relacionadas a esta equipe.
            single_team=teams[teams.ID_EQUIPE.eq(eqid)]
            zmap.writestr(f"{subdir}/Rota_{safe}.kml",build_kml(selected,single_team,f"NIP - {e.COLABORADOR}",kmz=False))
            zmap.writestr(f"{subdir}/Rota_{safe}.kmz",wrap_kmz(build_kml(selected,single_team,f"NIP - {e.COLABORADOR}",kmz=True)))
    return {"excel":whole,"kml":kml,"kmz":kmz,"equipes_excel_zip":per_excel.getvalue(),"equipes_mapas_zip":per_maps.getvalue()}


def main():
    st.set_page_config(page_title="NIP | Analise Cruzada de Equipes",page_icon="📍",layout="wide",initial_sidebar_state="expanded")
    st.markdown("""<style>.block-container{max-width:100%!important;padding-top:2.6rem!important;padding-left:2.2rem;padding-right:2.2rem}
    h1,h2,h3{color:#0D256C} [data-testid='stMetric']{background:#f8fafd;border-left:4px solid #0D256C;padding:12px 15px;border-radius:8px}
    </style>""",unsafe_allow_html=True)
    st.title("📍 NIP | Análise Cruzada de Obras e Equipes — V3")
    st.caption("BASE_SANEAMENTO_ATUALIZADA + OBRAS PLA + LOCALIDADE | Duplicidades, cobertura geográfica, sugestões e KML/KMZ")
    with st.sidebar:
        st.header("⚙️ Configuração")
        radius=st.slider("Raio para equipe próxima (km)",5,300,50,5,help="O raio considera distância geográfica em linha reta da base do colaborador até a obra.")
        sp=st.slider("Raio de superponto (m)",5,200,50,5)
        st.info("As distâncias são em linha reta (Haversine). Não representam percurso por estrada ou tempo de viagem. Coordenadas das equipes são as bases informadas na planilha LOCALIDADE.")
    a,b,c=st.columns(3)
    with a:up_s=st.file_uploader("1. BASE_SANEAMENTO_ATUALIZADA",type=["xlsx"],key="nip_san")
    with b:up_p=st.file_uploader("2. OBRAS PLA",type=["xlsx"],key="nip_pla")
    with c:up_e=st.file_uploader("3. LOCALIDADE LEVANTADORES-SANEAMENTO",type=["xlsx"],key="nip_equipes")
    if not all((up_s,up_p,up_e)):
        st.info("Selecione as três planilhas para iniciar o cruzamento.")
        return
    sig=hashlib.sha256(up_s.getvalue()+up_p.getvalue()+up_e.getvalue()+str(radius).encode()+str(sp).encode()).hexdigest()
    if st.session_state.get("nip_sig")!=sig:
        st.session_state.pop("nip_result",None);st.session_state.pop("nip_exports",None)
        st.session_state["nip_sig"]=sig
    if st.button("🚀 Analisar bases e localizar equipes",type="primary",use_container_width=True):
        try:
            with st.spinner("Conferindo notas, coordenadas, bases de equipes e cobertura..."):
                st.session_state.nip_result=analysis_from_bytes(up_s.getvalue(),up_p.getvalue(),up_e.getvalue(),radius,sp)
                st.session_state.pop("nip_exports",None)
            st.success("Cruzamento concluído. As notas foram mantidas no consolidado, mesmo quando aparecem nas duas bases.")
        except Exception as exc:
            st.session_state.pop("nip_result",None)
            st.error(f"Erro na análise: {exc}")
            return
    if "nip_result" not in st.session_state:return
    r=st.session_state.nip_result
    x1,x2,x3,x4,x5,x6=st.columns(6)
    for box,(label,val) in zip([x1,x2,x3,x4,x5,x6],[
        ("Notas únicas",len(r["obras"])),("Duplicadas nas bases",len(r["duplicadas"])),
        ("Atividades sem equipe perto",len(r["sem_equipe"])),("Atividades com 2+ equipes",len(r["multiplas"])),
        ("Equipes mapeadas",len(r["equipes"])),("Sem coordenadas",len(r["sem_coordenadas"]))]):
        box.metric(label,f"{val:,}".replace(",","."))
    st.caption("Uma nota existente nas duas bases conta uma vez em 'Notas únicas' e duas vezes em 'Atendimentos', uma por atividade. 'Sem equipe' significa sem colaborador do tipo adequado dentro do raio configurado.")
    overview,missing,duplicates,no_team,many,roster,audit,maptab=st.tabs(["📊 Visão geral","📍 Sem coordenadas","🟣 Duplicadas","🔴 Sem equipe próxima","🟠 Duas ou mais equipes","👥 Equipes","📋 Auditoria","🗺️ Mapa"])
    with overview:
        filters=st.columns([1,2])
        with filters[0]: category=st.selectbox("Filtrar origem",["TODAS","SANEAMENTO","LEVANTAMENTO","SANEAMENTO + LEVANTAMENTO"])
        with filters[1]:search=st.text_input("Buscar nota, cliente ou município")
        show=r["obras"]
        if category!="TODAS":show=show[show.ORIGEM.eq(category)]
        if search.strip():
            ss=re.escape(search.strip())
            mask=show.NOTA.astype(str).str.contains(ss,case=False,na=False)|show.NOME_CLIENTE.astype(str).str.contains(ss,case=False,na=False)|show.MUNICIPIO.astype(str).str.contains(ss,case=False,na=False)
            show=show[mask]
        st.dataframe(ordered_master(show),hide_index=True,use_container_width=True,height=510)
        st.subheader("Atendimentos separados por atividade")
        st.dataframe(r["atendimentos"],hide_index=True,use_container_width=True,height=310)
    with missing:
        st.error(f"⚠️ {len(r['sem_coordenadas'])} nota(s) sem coordenadas válidas. Não entram no mapa nem na seleção de equipes por distância até a correção.")
        st.dataframe(ordered_missing(r["sem_coordenadas"]),hide_index=True,use_container_width=True,height=410)
        st.subheader("Pendências identificadas nas bases de origem")
        st.caption("Também inclui registro sem coordenadas em uma base quando a mesma nota possui localização válida na outra. Verifique a aba e a linha originais para correção.")
        st.dataframe(ordered_source_issues(r["pendencias_coordenadas_origem"]),hide_index=True,use_container_width=True,height=320)
        st.download_button("📥 Excel apenas com pendências de coordenadas",
                           make_excel({"SEM COORDENADAS": ordered_missing(r["sem_coordenadas"]),
                                       "PENDENCIAS NA ORIGEM": ordered_source_issues(r["pendencias_coordenadas_origem"])}),
                           file_name="NIP_Obras_Sem_Coordenadas.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with duplicates:
        st.caption("O mesmo número de NOTA na base Saneamento e PROTOCOLO em OBRAS PLA. Uma linha consolidada por nota.")
        st.dataframe(ordered_master(r["duplicadas"]),hide_index=True,use_container_width=True,height=530)
    with no_team:
        st.info("Mesmo sem uma equipe dentro do raio, a sugestão aponta o colaborador elegível de menor distância geográfica, inclusive quando estiver longe.")
        st.dataframe(r["sem_equipe"],hide_index=True,use_container_width=True,height=420)
        st.subheader("Sugestões externas ao raio")
        st.dataframe(r["sugestoes"],hide_index=True,use_container_width=True,height=300)
    with many:
        st.dataframe(r["multiplas"],hide_index=True,use_container_width=True,height=390)
        st.subheader("Todas as equipes próximas — nome, atividade e distância")
        st.dataframe(r["proximas"],hide_index=True,use_container_width=True,height=390)
    with roster:
        st.dataframe(r["equipes"],hide_index=True,use_container_width=True,height=500)
    with audit:
        st.subheader("Divergências entre bases")
        st.dataframe(r["divergencias"],hide_index=True,use_container_width=True)
        st.subheader("Notas sem coordenadas (auditoria)")
        st.dataframe(ordered_missing(r["sem_coordenadas"]),hide_index=True,use_container_width=True)
        st.subheader("Parâmetros utilizados")
        st.dataframe(r["parametros"],hide_index=True,use_container_width=True)
    with maptab:
        st.caption("Marcadores de obras por status, casinhas das bases das equipes. Os pontos representam localizações informadas nas planilhas.")
        if st.checkbox("Carregar mapa interativo",value=False):
            geo=r["obras"].dropna(subset=["LATITUDE","LONGITUDE"])
            if len(geo):
                mp=folium.Map(location=[float(geo.LATITUDE.mean()),float(geo.LONGITUDE.mean())],zoom_start=7,tiles="OpenStreetMap",prefer_canvas=True)
                iconcolors={"Duplicadas nas bases":"purple","Sem equipe proxima":"red","Duas ou mais equipes proximas":"orange","Obras Saneamento":"blue","Obras Levantamento":"green"}
                for cat,col in iconcolors.items():
                    layer=folium.FeatureGroup(name=cat,show=True)
                    cl=MarkerCluster().add_to(layer)
                    rows=geo[geo.apply(marker_category,axis=1).eq(cat)]
                    for _,w in rows.iterrows():
                        folium.Marker([float(w.LATITUDE),float(w.LONGITUDE)],icon=folium.Icon(color=col,icon="map-marker"),popup=folium.Popup(popup_work(w),max_width=370)).add_to(cl)
                    layer.add_to(mp)
                layer=folium.FeatureGroup(name="Casas / equipes",show=True)
                for _,e in r["equipes"].dropna(subset=["LAT_BASE","LON_BASE"]).iterrows():
                    folium.Marker([e.LAT_BASE,e.LON_BASE],icon=folium.Icon(icon="home",color="lightblue"),tooltip=f"{e.COLABORADOR} | {e.TIPO_EQUIPE} | {e.MUNICIPIO_BASE}").add_to(layer)
                layer.add_to(mp)
                folium.LayerControl(collapsed=False).add_to(mp)
                st_folium(mp,height=630,use_container_width=True)
    st.divider()
    st.subheader("📦 Exportar Excel e mapas — padrão Lista Contínua")
    st.caption("Excel com cabeçalhos azul-marinho, filtros e abas de auditoria; KML geral e KMZ com ícone de casinha; arquivos individuais por equipe.")
    if st.button("📦 Preparar todos os arquivos",type="primary"):
        try:
            with st.spinner("Gerando Excel consolidado, KML/KMZ e pacotes individuais por equipe..."):
                st.session_state.nip_exports=make_all_exports(r)
            st.success("Arquivos preparados para download.")
        except Exception as exc:
            st.error(f"Falha ao gerar exportações: {exc}")
    if "nip_exports" in st.session_state:
        exp=st.session_state.nip_exports
        one,two,three=st.columns(3)
        one.download_button("📊 Excel completo",exp["excel"],file_name="NIP_Analise_Cruzada_Equipes.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        two.download_button("🗺️ KML geral",exp["kml"],file_name="NIP_Analise_Geral.kml",mime="application/vnd.google-earth.kml+xml",use_container_width=True)
        three.download_button("🏠 KMZ geral",exp["kmz"],file_name="NIP_Analise_Geral.kmz",mime="application/vnd.google-earth.kmz",use_container_width=True)
        f,g=st.columns(2)
        f.download_button("📚 Excel individual por equipe (ZIP)",exp["equipes_excel_zip"],file_name="NIP_Excel_Individual_Equipes.zip",mime="application/zip",use_container_width=True)
        g.download_button("🏠 KML e KMZ individual por equipe (ZIP)",exp["equipes_mapas_zip"],file_name="NIP_KML_KMZ_Equipes.zip",mime="application/zip",use_container_width=True)


if __name__ == "__main__":
    main()
