
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
