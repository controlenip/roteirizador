import streamlit as st
import pandas as pd
import numpy as np
import folium
import io
import zipfile
import html
import re
import time
import gc
import os
import requests
import math
import hashlib
import base64
import json
import csv
from urllib.parse import urlencode
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from datetime import datetime




# Adaptacao exclusiva da leitura da planilha. O motor e os exportadores abaixo
# sao os mesmos do modulo original Lista Continua.
import unicodedata
from sklearn.cluster import DBSCAN
from openpyxl.styles import PatternFill, Font, Alignment

# Ordem solicitada para os detalhes das obras nos cards do KML.
CAMPOS_KML_SOLICITADOS = [
    'TIPO NOTA', 'ID SISCO', 'PROTOCOLO', 'CONTA CONTRATO', 'INSTALAÇÃO',
    'NOME', 'REGIONAL', 'MUNICIPIO', 'ENDEREÇO', 'LOCALIDADE',
    'LATITUDE', 'LONGITUDE', 'FASE'
]
# Preserva os campos anteriores como escolhas adicionais no seletor.
CAMPOS_VISIVEIS_NIP = CAMPOS_KML_SOLICITADOS + [
    'NOTA', 'STATUS', 'TIPO DEMANDA', 'BAIRRO', 'COLABORADOR', 'MATRICULA'
]

# O icone fornecido pelo usuario fica incorporado ao .py.
# Os mapas exportados incluem a imagem e nao dependem de arquivo externo.
ICONE_CASA_ARQUIVO = 'CASINHA_ORIGEM.png'
ICONE_CASA_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFxUlEQVR4nO3YW0xTdxwH8G7Zlj2YJcuSJXtYsizLnvaypz0pGVCgpdALl1JKoSAthVKgDAUUoUMCIqBQQMcg6EBAaLi0KBMEuSOXiheUocigAjrinM5JuCh8l3McpECRFjiAC7/k+9J/L79Pz/md/8mh0XZrt3Zrt7aqyuLiPqC9bZWVVbYnNfVnRY4qtVETGjVRGhg+rXMTID0yvoe2UyspqejjxabDoh+3iiWvRjhcPHR2xl2eCyoF3uixs4OeTkenvQPiE7OjaTvpH69VpVTrPUXzY2w22bRxRp3ZqPLyRTvTkQQYp5brOp+YlvvVdhtolak5B/v8pNPLm19Ii89+1POFKwDGKfQLmtw2gCb7rF1n6MGJ1QC3xP6oDVC8EWCctMOJLVsKKDl55ouGI4lXR7k8k4AhvgCXQg6gy4FhNkJPp6PDgYGE5NNBlAMaGxvfu5ikzhkUCOdMAcbZbDQG/4A2rqtFAL1Ralzc549ml3xOGaIqJUd5yz9warXT6KosBFdEvusG6I2SL1M+3XSAJqvgu7YDR4ZWA/R7+eBSkHJTAHrjeYk51rApgIJTBZ/Wxh5reuDiahIw7OqG2pCD6DJxOd2MtDOYiE/O8V83IC4u7t1L8WlZ94SiV8aNj3G4uO0fiB5lNHoPxKBOIqcEoDdKtZtg7lBO+WcWIyoz8kS9gSHPRzlc3PGTQq+MwvWIGPyuCMe0lwjg8xczEhCMTkcW5Zg8ecRjixBtCWn9fRGHMS5TYE4gWNK0cSbFfhg4FI9pbx90KKMoh+jpdCTHplSZDVmt6fuRsWhQn0F+RTNU3aMIGJmH1ADkatvJ91yWh1MOaWWyEJucyzcL8iQgaHYgWjXanZReq667jojbT8mG35SOlFOYFHmjSSCiHKPlC+eIm1GLTrO1AAtRDE7joSIcw7Jgyq5eeqP8pIh8RAmEiKrrAWZE3mileF6andjIdfEAZRAi50ovv56XwNBNB3Q7MFAWEIIKhiOq9+6lFkKk81gm/vEWo4XvuXlzIZaijOdOAoicFQc8phyiuDeFP4KVGJSHWXzXq1+Weg8hisUBi4DzXPcZtSr1R4sQyyH+92fMxsR3GjAr9EKzMnpdgDZnDopkodB+b00CKuwZUB05jqy4rD0WI5ZD6Dnl8GgehGx4zixMaVEN5j08UC8LMX8O6HRopMEoZ7JeHwUrK6hDosBOL4S1utiyuVgNQnwREeIym1F3Y0nTomsTKyDEhtkbn4q/fXzR4i5YE1EtluI8X7h4GuX7yiA6cXbxdzcdsrDbL19jlTbDu39yyethAy/wpywYvynC0WNvbxLQ5OqOIn85dPv2kYAyZy4iEjJhk1G0BLFlECLM/jk4dU1g/9Ds4lpi6yBeegrRFLZ0f7nKdESxLAxaG1sSUGVLR8LhJDioVwK2HjIAMoybU0vWC4+ffj0vEjn5HKtcIoeGzSMBxJHIlkfA5WTBqoBtgxAxXidOkxtRcXjiux8lAp/FOfjFyw9+aWfWBFhvFcQmuwysoitwquwGs/f5CgjxOdfMYvwlkZEADYuN6KMZZgOsNwwZQaLUgKmFZmyyzuOl0GsFhHehF7yqHvC0enCrusG6cNPkEYtUFyIpKgGMN8yBtYkQv+usaemgbaQCDfhSYkA581z9OEfXa1WTXWR1N0o1vvxyK+57RiIIDBFTkPWEUVj3yK2m14pGVa22EXp2GMAp79wwxC5PN8PT9cRQBlgLQmbZ7m8JwPa0Zp5d0V7h0zj8IeWI/yAZkhG8NOd2xaw5yCwBq+RKn5P22tY/jfcfw9dSAy5sFMIs/PUJ7+J1Dm27S2KArcSAO5ZC7PK0sxxt1wnaTiqpHu9LDAiVGvBsLQix7zhpWpuENfc+ou3Uko/hE2J+pCN4tQJCzkHDELf6xje0t6Vkw/hWMoLmBYhD/sUXbJ1+/c9wt7WAd6QG8Lm6rnQa8O72NrNbu7VbtP9D/QvGWFY6Zjt3KQAAAABJRU5ErkJggg=='


def imagem_casa_bytes():
    return base64.b64decode(ICONE_CASA_BASE64)


def gerar_kmz_lista(texto_kml):
    """Empacota um KML autossuficiente com a casinha personalizada para Google Earth."""
    buffer_kmz = io.BytesIO()
    with zipfile.ZipFile(buffer_kmz, 'w', zipfile.ZIP_DEFLATED) as kmz:
        kmz.writestr('doc.kml', texto_kml.encode('utf-8'))
        kmz.writestr(ICONE_CASA_ARQUIVO, imagem_casa_bytes())
    return buffer_kmz.getvalue()

# Google Maps: links universais oficiais (sem chave de API), CSV para My Maps e
# painel HTML local para reorganizar visitas sem alterar a rota central da NIP.
# Para garantir compatibilidade com navegadores de celular, cada link tem no
# maximo 4 obras (destino + ate 3 waypoints). O restante vai em outras etapas.
GOOGLE_MAPS_MAX_OBRAS_ETAPA = 4


def texto_seguro_maps(valor):
    """Texto legivel para etiquetas, sem NaN nem '123.0' para identificadores."""
    try:
        if pd.isna(valor):
            return ''
    except (ValueError, TypeError):
        pass
    if isinstance(valor, float) and math.isfinite(valor) and valor.is_integer():
        return str(int(valor))
    resultado = str(valor).strip()
    return '' if resultado.upper() in ('NAN', 'NONE', 'NULL') else resultado


def pontos_equipe_google_maps(df_equipe):
    """Converte pontos da equipe na ordem do motor, preservando superpontos.

    Cada registro representa uma parada fisica, e as notas agregadas de um
    superponto continuam visiveis no painel e no CSV de importacao.
    """
    pontos = []
    df_equipe = df_equipe.sort_values('ORDEM', kind='stable')
    for _, linha in df_equipe.iterrows():
        protocolo = texto_seguro_maps(linha.get('PROTOCOLO'))
        if protocolo in ('RETORNO_BASE', 'PAUSA_ALMOCO'):
            continue
        try:
            lat, lon = float(linha['LATITUDE']), float(linha['LONGITUDE'])
        except (ValueError, TypeError, KeyError):
            continue
        if not (math.isfinite(lat) and math.isfinite(lon) and -90 <= lat <= 90 and -180 <= lon <= 180):
            continue
        originais = linha.get('_ORIGINAL_ROWS')
        notas = []
        if isinstance(originais, list):
            for item in originais:
                if isinstance(item, dict):
                    codigo = texto_seguro_maps(item.get('PROTOCOLO', item.get('NOTA', '')))
                    if codigo and codigo not in notas:
                        notas.append(codigo)
        if not notas:
            notas = [protocolo] if protocolo else []
        nota_resumo = ' | '.join(notas)
        pontos.append({
            'ordem': int(linha['ORDEM']),
            'nota': nota_resumo,
            'municipio': texto_seguro_maps(linha.get('MUNICIPIO')),
            'localidade': texto_seguro_maps(linha.get('LOCALIDADE')),
            'endereco': texto_seguro_maps(linha.get('ENDEREÇO')),
            'prioridade': eh_prioridade(linha.get('PRIORIDADE')),
            'lat': lat, 'lon': lon,
            'qtd_obras': len(originais) if isinstance(originais, list) else 1,
        })
    return pontos


def origem_google_maps(df_equipe):
    """Le a origem municipal de uma equipe (a referencia e a sede municipal)."""
    if df_equipe.empty:
        raise ValueError('Equipe sem obras roteirizadas.')
    primeira = df_equipe.sort_values('ORDEM').iloc[0]
    lat, lon = float(primeira['LAT_ORIGEM']), float(primeira['LON_ORIGEM'])
    if not all(map(math.isfinite, (lat, lon))):
        raise ValueError('Origem municipal sem coordenadas validas.')
    return {
        'municipio': texto_seguro_maps(primeira.get('ORIGEM_MUNICIPIO')),
        'lat': lat, 'lon': lon,
        'km_iniciais': float(primeira.get('DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM') or 0.0),
        'min_iniciais': float(primeira.get('TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN') or 0.0),
        'metodo': texto_seguro_maps(primeira.get('TIPO_DESLOCAMENTO')),
    }


def coordenada_maps(ponto):
    return f"{float(ponto['lat']):.7f},{float(ponto['lon']):.7f}"


def link_google_maps_primeira_obra(origem, primeira, local_atual=False):
    """URL universal para navegar ate a primeira obra (com origem opcional)."""
    parametros = {'api': '1', 'destination': coordenada_maps(primeira), 'travelmode': 'driving'}
    if not local_atual:
        parametros['origin'] = coordenada_maps(origem)
    return 'https://www.google.com/maps/dir/?' + urlencode(parametros)


def link_google_maps_paradas(origem, obras, local_atual=False):
    """Rota editavel no Maps com ate 4 obras, mantendo a ordem escolhida."""
    if not obras:
        raise ValueError('E necessario selecionar ao menos uma obra.')
    if len(obras) > GOOGLE_MAPS_MAX_OBRAS_ETAPA:
        raise ValueError('A etapa excede o limite de obras para navegadores moveis.')
    parametros = {'api': '1', 'travelmode': 'driving'}
    if not local_atual:
        parametros['origin'] = coordenada_maps(origem)
    parametros['destination'] = coordenada_maps(obras[-1])
    if len(obras) > 1:
        parametros['waypoints'] = '|'.join(coordenada_maps(p) for p in obras[:-1])
    url = 'https://www.google.com/maps/dir/?' + urlencode(parametros)
    if len(url) > 2048:
        raise ValueError('URL do Google Maps excedeu 2.048 caracteres.')
    return url


def gerar_etapas_google_maps(origem, obras, tamanho=GOOGLE_MAPS_MAX_OBRAS_ETAPA):
    """Divisao sem perda de pontos: cada etapa termina onde comeca a proxima."""
    etapas = []
    origem_etapa = origem
    for inicio in range(0, len(obras), tamanho):
        grupo = obras[inicio:inicio + tamanho]
        if not grupo:
            continue
        etapas.append({
            'etapa': len(etapas) + 1,
            'ordens': [x['ordem'] for x in grupo],
            'url': link_google_maps_paradas(origem_etapa, grupo),
            'url_local_atual': link_google_maps_paradas(origem_etapa, grupo, local_atual=True),
            'primeira_nota': grupo[0]['nota'],
            'ultima_nota': grupo[-1]['nota'],
        })
        origem_etapa = grupo[-1]
    return etapas


def gerar_csv_google_mymaps(df_equipe, equipe, incluir_origem=True):
    """CSV UTF-8 BOM importavel em Google My Maps, com ordem e coordenadas."""
    buffer = io.StringIO(newline='')
    campos = ['Nome', 'Equipe', 'Ordem', 'Tipo', 'Prioridade', 'Nota CCS',
              'Municipio', 'Localidade', 'Endereco', 'Latitude', 'Longitude', 'Qtd Obras']
    writer = csv.DictWriter(buffer, fieldnames=campos, lineterminator='\r\n')
    writer.writeheader()
    if incluir_origem:
        ori = origem_google_maps(df_equipe)
        writer.writerow({'Nome': f"CASINHA - Origem {equipe} ({ori['municipio']})",
                         'Equipe': equipe, 'Ordem': 0, 'Tipo': 'ORIGEM / HOSPEDAGEM',
                         'Prioridade': '', 'Nota CCS': '', 'Municipio': ori['municipio'],
                         'Latitude': ori['lat'], 'Longitude': ori['lon'], 'Qtd Obras': 0})
    for p in pontos_equipe_google_maps(df_equipe):
        writer.writerow({'Nome': f"{p['ordem']:03d} - {p['nota']}",
                         'Equipe': equipe, 'Ordem': p['ordem'], 'Tipo': 'OBRA',
                         'Prioridade': 'SIM' if p['prioridade'] else 'NAO',
                         'Nota CCS': p['nota'], 'Municipio': p['municipio'],
                         'Localidade': p['localidade'], 'Endereco': p['endereco'],
                         'Latitude': p['lat'], 'Longitude': p['lon'],
                         'Qtd Obras': p['qtd_obras']})
    return buffer.getvalue()


def gerar_txt_links_google_maps(df_equipe, equipe):
    origem = origem_google_maps(df_equipe)
    pontos = pontos_equipe_google_maps(df_equipe)
    linhas = [f'EQUIPE: {equipe}', f"PARTIDA: {origem['municipio']} (sede municipal)",
              f"PRIMEIRA OBRA: {pontos[0]['nota'] if pontos else '-'}",
              '', 'ABRIR A PRIMEIRA OBRA NO GOOGLE MAPS:']
    if pontos:
        linhas.append(link_google_maps_primeira_obra(origem, pontos[0]))
        linhas.extend(['', 'ROTAS POR ETAPA (ate 4 obras por link; funciona em celular):'])
        for etapa in gerar_etapas_google_maps(origem, pontos):
            linhas.append(f"Etapa {etapa['etapa']} - notas {etapa['primeira_nota']} a {etapa['ultima_nota']}")
            linhas.append(etapa['url'])
    if pontos:
        linhas.extend(['', 'IPHONE / iOS - GOOGLE MAPS PRIMEIRA OBRA (copiar no Safari se necessario):',
                       link_google_maps_primeira_obra(origem, pontos[0]),
                       'IPHONE / iOS - GOOGLE MAPS USANDO GPS ATUAL:',
                       link_google_maps_primeira_obra(origem, pontos[0], local_atual=True),
                       'IPHONE / iOS - APPLE MAPS PRIMEIRA OBRA:',
                       link_apple_maps_obra(origem, pontos[0])])
        linhas.append('')
        linhas.append('IPHONE / iOS - ABRIR QUALQUER OBRA INDIVIDUALMENTE (GPS ATUAL):')
        for ponto in pontos:
            linhas.append(f"ORDEM {ponto['ordem']} - NOTA {ponto['nota']}")
            linhas.append(link_google_maps_primeira_obra(origem, ponto, local_atual=True))
    linhas.extend(['', 'Voce pode alterar a sequencia de visitas no painel HTML da sua equipe.',
                   'No Google Maps, confirme as ruas, tempo de viagem e viabilidade local.',
                   'A localizacao de moradia/hospedagem representa a sede municipal, nao um endereco.'])
    return '\n'.join(linhas) + '\n'


def gerar_painel_google_maps(df_equipe, equipe, pwa=False):
    """Planejador autossuficiente: lista, prioridades, edicao, backup e GPX offline.

    Todos os dados e scripts sao incorporados; nenhum CDN/servidor e exigido para
    consultar e editar no navegador compativel. No iOS, a pre-visualizacao de
    HTML do app Arquivos pode bloquear JS: use a versao hospedada em HTTPS/PWA.
    A navegacao rodoviaria offline cabe ao aplicativo de mapas previamente instalado.
    """
    origem = origem_google_maps(df_equipe)
    pontos = pontos_equipe_google_maps(df_equipe)
    assinatura = hashlib.sha256(json.dumps(pontos, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()[:16]
    dados = {'equipe': str(equipe), 'origem': origem, 'obras': pontos, 'versao': assinatura}
    primeira = pontos[0] if pontos else origem
    fallback_primeira = link_google_maps_primeira_obra(origem, primeira)
    fallback_gps = link_google_maps_primeira_obra(origem, primeira, local_atual=True)
    fallback_etapa = gerar_etapas_google_maps(origem, pontos)[0]['url'] if pontos else fallback_primeira
    dados_json = (json.dumps(dados, ensure_ascii=False, allow_nan=False)
                  .replace('<', '\\u003c').replace('>', '\\u003e')
                  .replace('&', '\\u0026'))
    linhas_simples = []
    for i, pt in enumerate(pontos, 1):
        nome = html.escape(pt['nota'])
        mun = html.escape(pt['municipio'])
        url = html.escape(link_google_maps_primeira_obra(origem, pt, local_atual=True), quote=True)
        linhas_simples.append(f'<li><strong>{i}. Nota {nome}</strong> • {mun} '
                             f'({pt["lat"]:.6f}, {pt["lon"]:.6f}) '
                             f'<a href="{url}">Maps (online)</a></li>')
    fallback_html = ''.join(linhas_simples) or '<li>Nenhuma obra valida nesta equipe</li>'
    pwa_head = ('<link rel="manifest" href="./manifest.webmanifest">'
                '<link rel="icon" href="./icon-192.png" type="image/png">'
                '<meta name="apple-mobile-web-app-capable" content="yes">'
                '<meta name="apple-mobile-web-app-title" content="NIP Campo">') if pwa else ''
    pwa_script = ("if('serviceWorker' in navigator && "
                  "(location.protocol==='https:' || location.hostname==='localhost')){"
                  "navigator.serviceWorker.register('./sw.js').catch(()=>{});}") if pwa else ''
    template = r"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#0d256c"><title>NIP Campo - planejamento offline</title>
__PWA_HEAD__
<style>
:root{font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#162a43;background:#f3f6fa;font-size:15px}
*{box-sizing:border-box}body{margin:0}button,input,select{font:inherit}.top{background:#0d256c;color:white;padding:17px 14px}
.wrap{max-width:900px;margin:auto;padding:12px}.top .wrap{padding:0 4px}h1{font-size:1.4rem;margin:6px 0}
h2{font-size:1.12rem;margin:0 0 10px}h3{font-size:1rem;margin:10px 0}.card{background:white;border:1px solid #dae2ec;border-radius:12px;margin:12px 0;padding:15px;box-shadow:0 2px 8px #1b2c450a}
.note{color:#526478;font-size:.87rem;line-height:1.5}.alert{border-left:4px solid #df9f28;padding:8px 10px;background:#fffaf0;border-radius:5px}
.btn{border:0;background:#0d256c;color:white;border-radius:9px;padding:11px 13px;cursor:pointer;display:inline-block;text-decoration:none;font-size:.90rem;font-weight:650;text-align:center}
.btn.green{background:#16814e}.btn.soft{background:#e6eef9;color:#0d256c}.btn.gray{background:#e9ebef;color:#243346}.btn.danger{background:#fff0ed;color:#a4252a}
.btn:disabled{opacity:.45}.actions{display:flex;flex-wrap:wrap;gap:8px;margin:9px 0}.kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}.kpi{background:#edf3ff;padding:11px;border-radius:9px;text-align:center}.kpi b{display:block;font-size:1.5rem;color:#0d256c}
.item{border:1px solid #dce4ec;border-left:5px solid #0d256c;border-radius:9px;padding:12px;margin:9px 0}.item.priority{border-left-color:#cf3030}.item.done{opacity:.67;background:#f0f5f0}.row{display:flex;gap:9px;align-items:center;justify-content:space-between;flex-wrap:wrap}
.tag{font-size:.74rem;background:#e8eeff;color:#0d256c;border-radius:100px;padding:4px 9px;font-weight:700}.tag.red{background:#fff0ef;color:#a22}.tag.green{background:#e4f6e9;color:#166e37}
.muted{color:#586779;font-size:.85rem}.small{font-size:.8rem}.filters{display:flex;gap:8px;flex-wrap:wrap}.filters input{flex:3;min-width:150px}.filters select{flex:1;min-width:140px}input[type=search],select{width:100%;padding:10px;border:1px solid #cbd7e5;border-radius:9px;background:white;color:#123}
.mode{display:flex;gap:8px;flex-wrap:wrap}.mode button{flex:1}.mode button.active{outline:3px solid #73b9fa}.hidden{display:none!important}
#toast{min-height:18px;color:#167844;font-weight:650}.links a{overflow-wrap:anywhere}.staticlist{max-height:430px;overflow:auto;padding-left:24px}.staticlist li{padding:5px 0}.coords{font-family:ui-monospace,monospace;font-size:.82rem;user-select:all}
@media(max-width:580px){.btn{flex:1;padding:12px 9px}.item .btn{flex:1 1 43%}.card{padding:12px}.top{padding:15px 8px}.kpi b{font-size:1.25rem}}
</style></head><body>
<header class="top"><div class="wrap"><div>🏠 NIP GLOBAL • Lista Contínua</div><h1 id="titulo">__EQUIPE_HTML__</h1><div>Partida: __ORIGEM_HTML__ / MA</div></div></header>
<main class="wrap">
<section class="card"><div class="row"><div><strong>📶 Situação do dispositivo:</strong> <span id="conexao">Verificando</span></div><span class="tag" id="modoTag">PLANEJAMENTO OFFLINE</span></div>
<p class="note">A lista e o planejamento funcionam sem sinal quando este HTML ja estiver salvo e puder ser aberto num navegador compativel. Os botões do Google Maps dependem do aplicativo e da disponibilidade dos mapas; para navegação offline use o GPX no Organic Maps/OsmAnd, com mapas baixados antes.</p>
<div class="mode"><button type="button" id="modoOffline" class="btn active">📴 Planejamento offline</button><button type="button" id="modoOnline" class="btn soft">🌐 Google Maps online</button></div>
<p id="alertaArmazenamento" class="note alert">Se o navegador nao permitir gravacao local, exporte o backup JSON ao fim do expediente.</p></section>
<section class="card"><div class="kpis"><div class="kpi"><span class="small">Paradas</span><b id="qtd">__QTD__</b></div><div class="kpi"><span class="small">Concluidas</span><b id="feitas">0</b></div><div class="kpi"><span class="small">Pendentes</span><b id="pend">__QTD__</b></div></div>
<p id="primeira" class="note">Primeira obra na ordem original: __NOTA_HTML__</p></section>
<section class="card"><h2>📴 Trabalhar sem internet</h2>
<p class="note">Depois de mudar a ordem ou concluir obras, exporte um GPX atualizado. Importe no <strong>Organic Maps</strong> ou <strong>OsmAnd</strong> para ver as paradas e navegar pelas estradas com os mapas da regiao previamente baixados. GPX contem <strong>pontos na ordem escolhida, nao uma estrada desenhada</strong>.</p>
<div class="actions"><button id="exportGPX" class="btn green" type="button">🧭 GPX das pendentes</button><button id="exportKML" class="btn soft" type="button">🗺️ KML atualizado</button><button id="exportCSV" class="btn soft" type="button">📊 CSV progresso</button></div>
<h3>💾 Salvar ou recuperar alterações</h3><div class="actions"><button id="exportJSON" class="btn" type="button">⬇ Salvar backup JSON</button><label for="importJSON" class="btn gray">⬆ Restaurar backup JSON</label><input id="importJSON" accept=".json,application/json" type="file" style="display:none"></div>
<p class="note">O arquivo JSON permite restaurar ordem e concluídas, inclusive em outro celular, desde que seja <strong>a mesma equipe e versao da rota</strong>. Sem sincronizacao automática com o escritorio.</p></section>
<section class="card" id="blocoOnline"><h2>🌐 Navegação no Google Maps</h2>
<p class="note">Cada link envia no máximo 4 obras, preservando a sequência de visita. As demais aparecem nas etapas seguintes. Para partir do GPS atual, habilite a opção abaixo.</p>
<label><input type="checkbox" id="localAtual"> Usar minha localização atual no primeiro trecho</label>
<div class="actions"><a class="btn green" id="linkPrimeira" href="__PRIMEIRA__">📍 Ir a primeira obra</a><a class="btn soft" id="rotaProximas" href="__ETAPA__">🚗 Abrir proximas 4 obras</a><a class="btn soft" href="__GPS__">GPS atual - link direto</a></div>
<div id="etapas"></div></section>
<section class="card"><h2>📋 Organizar atendimentos</h2><div class="filters"><input id="busca" type="search" placeholder="Buscar por nota, municipio ou endereco"><select id="filtro"><option value="todas">Todas as obras</option><option value="pendentes">Somente pendentes</option><option value="prioritarias">Prioritarias pendentes</option><option value="feitas">Concluidas</option></select></div>
<div class="actions"><button id="reset" class="btn danger" type="button">Restaurar ordem e status</button></div><div id="lista"></div><div class="actions"><button id="mais" class="btn soft hidden" type="button">Mostrar mais obras</button></div></section>
<section class="card"><h2>📌 Lista de emergencia (sem JavaScript)</h2><p class="note">Caso o iPhone abra uma visualizacao que bloqueia JavaScript, as obras abaixo continuam visiveis e podem ser consultadas offline. Para alterar o planejamento no iPhone, hospede a pasta "Offline_Instalavel" em HTTPS e abra pelo Safari.</p>
<details id="listaEstatica"><summary>Mostrar todas as notas e coordenadas</summary><ol class="staticlist">__LISTA_ESTATICA__</ol></details></section>
<section class="card"><h2>ℹ️ Preparacao para campo</h2><p class="note">1. Baixe os mapas do MA no Organic Maps/OsmAnd quando houver Wi-Fi.<br>2. Importe o GPX desta equipe no aplicativo de mapas.<br>3. Abra este HTML para reordenar e dar baixa nas obras.<br>4. Se mudar a lista, exporte novamente o GPX e importe a nova versão no aplicativo.<br>5. Para entregar o progresso ao escritorio, exporte CSV e backup JSON.</p><p id="toast" role="status" aria-live="polite"></p></section>
</main><script id="dados" type="application/json">__DADOS__</script><script>
(()=>{'use strict';
const data=JSON.parse(document.getElementById('dados').textContent);
const original=data.obras.map((p,i)=>({...p,id:i,feito:false}));
const k='nip_maps_v4_'+data.equipe+'_'+data.origem.municipio+'_'+data.versao;
const oldKey='nip_maps_v3_'+data.equipe+'_'+data.origem.municipio+'_'+data.versao;
const el=id=>document.getElementById(id);
let obras=original.map(p=>({...p})),modo='offline',visiveis=100,storeOK=false;
const coord=p=>Number(p.lat).toFixed(7)+','+Number(p.lon).toFixed(7);
const safeName=String(data.equipe).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9_-]+/g,'_').slice(0,48);
function toast(s){el('toast').textContent=s;}
function validState(s){if(!s||!Array.isArray(s.order)||s.order.length!==original.length)return false;
const ids=new Set(original.map(p=>p.id));return new Set(s.order).size===original.length&&s.order.every(x=>ids.has(x))&&Array.isArray(s.feitas)&&s.feitas.every(x=>ids.has(x));}
function applyState(s){if(!validState(s))return false;const dict=new Map(original.map(p=>[p.id,{...p}]));obras=s.order.map(id=>dict.get(id));const feitas=new Set(s.feitas);obras.forEach(p=>p.feito=feitas.has(p.id));return true;}
try{const v=localStorage.getItem(k)||localStorage.getItem(oldKey);if(v)applyState(JSON.parse(v));localStorage.setItem(k,JSON.stringify({order:obras.map(p=>p.id),feitas:obras.filter(p=>p.feito).map(p=>p.id)}));storeOK=true;}catch(e){storeOK=false;}
function persist(){if(storeOK){try{localStorage.setItem(k,JSON.stringify({order:obras.map(p=>p.id),feitas:obras.filter(p=>p.feito).map(p=>p.id)}));}catch(e){storeOK=false;}}el('alertaArmazenamento').textContent=storeOK?'✅ Alteracoes gravadas neste navegador. Exporte backup JSON para guardar uma copia e transferir entre dispositivos.':'⚠️ Este navegador nao permite salvar progresso automaticamente. Use Salvar backup JSON antes de fechar!';}
function pendentes(){return obras.filter(p=>!p.feito);}
function urlDirecoes(de,lista,gps=false){if(!lista.length)return '#';const a=new URLSearchParams({api:'1',travelmode:'driving',destination:coord(lista[lista.length-1])});if(!gps)a.set('origin',coord(de));if(lista.length>1)a.set('waypoints',lista.slice(0,-1).map(coord).join('|'));return 'https://www.google.com/maps/dir/?'+a.toString();}
function pin(p){return 'https://www.google.com/maps/search/?'+new URLSearchParams({api:'1',query:coord(p)});}
function routeGroups(){const p=pendentes(),g=[];let de=data.origem;for(let i=0;i<p.length;i+=4){const subset=p.slice(i,i+4);g.push({de,subset});de=subset[subset.length-1];}return g;}
function download(name,text,type){const blob=new Blob([text],{type});const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.style.display='none';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),30000);toast('Arquivo solicitado: '+name+'. Se o iPhone nao baixar, use Safari via HTTPS.');}
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function filename(suffix){return 'NIP_'+safeName+'_'+suffix;}
function exportGpx(){const p=pendentes();let t=['<?xml version="1.0" encoding="UTF-8"?>','<gpx version="1.1" creator="NIP Campo Offline" xmlns="http://www.topografix.com/GPX/1/1">','<metadata><name>'+esc(data.equipe)+' - pendentes</name><desc>Sequencia de paradas; NAO e tracado rodoviario.</desc></metadata>'];
const item=(tag,n,p,d)=>'<'+tag+' lat="'+Number(p.lat).toFixed(7)+'" lon="'+Number(p.lon).toFixed(7)+'"><name>'+esc(n)+'</name><desc>'+esc(d)+'</desc></'+tag+'>';
t.push(item('wpt','CASINHA - '+data.origem.municipio,data.origem,'Partida municipal aproximada'));
p.forEach((x,i)=>t.push(item('wpt',String(i+1).padStart(3,'0')+' - '+x.nota,x,x.municipio+' | '+(x.prioridade?'PRIORITARIA':'NORMAL'))));
t.push('<rte><name>'+esc(data.equipe)+' - ordem planejada</name><desc>Lista de paradas sem geometria de ruas; o app offline precisa calcular percurso.</desc>');
t.push(item('rtept','ORIGEM - '+data.origem.municipio,data.origem,'Ponto inicial municipal'));
p.forEach((x,i)=>t.push(item('rtept',String(i+1).padStart(3,'0')+' - '+x.nota,x,x.municipio)));
t.push('</rte></gpx>');download(filename('PENDENTES_OFFLINE.gpx'),t.join('\n'),'application/gpx+xml;charset=utf-8');}
function exportKml(){const p=pendentes();let t=['<?xml version="1.0" encoding="UTF-8"?>','<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>'+esc(data.equipe)+' - pendentes</name>'];const mark=(name,x,descr)=>'<Placemark><name>'+esc(name)+'</name><description>'+esc(descr)+'</description><Point><coordinates>'+Number(x.lon).toFixed(7)+','+Number(x.lat).toFixed(7)+',0</coordinates></Point></Placemark>';
t.push(mark('CASINHA - '+data.origem.municipio,data.origem,'Referencia municipal'));
p.forEach((x,i)=>t.push(mark((i+1)+' - '+x.nota,x,x.municipio+' | '+(x.prioridade?'PRIORITARIA':'NORMAL'))));t.push('</Document></kml>');download(filename('PENDENTES_OFFLINE.kml'),t.join('\n'),'application/vnd.google-earth.kml+xml;charset=utf-8');}
function csvVal(x){let s=String(x??'');if(/^[=+@\t\r]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
function exportCsv(){const rows=[['Equipe','Ordem atual','Nota CCS','Status','Prioridade','Municipio','Localidade','Endereco','Latitude','Longitude','Qtd Obras']];obras.forEach((p,i)=>rows.push([data.equipe,i+1,p.nota,p.feito?'CONCLUIDA':'PENDENTE',p.prioridade?'SIM':'NAO',p.municipio,p.localidade,p.endereco,p.lat,p.lon,p.qtd_obras]));download(filename('PROGRESSO.csv'),'\ufeff'+rows.map(r=>r.map(csvVal).join(';')).join('\r\n'),'text/csv;charset=utf-8');}
function state(){return {schema:'nip-offline-v1',equipe:data.equipe,versao:data.versao,geradoEm:new Date().toISOString(),order:obras.map(p=>p.id),feitas:obras.filter(p=>p.feito).map(p=>p.id)};}
function exportJson(){download(filename('BACKUP.json'),JSON.stringify(state(),null,2),'application/json;charset=utf-8');}
function mover(index,delta){const j=index+delta;if(j<0||j>=obras.length)return;[obras[index],obras[j]]=[obras[j],obras[index]];persist();render();}
function toggle(id){const p=obras.find(p=>p.id===id);if(p){p.feito=!p.feito;persist();render();}}
function button(text,cb,cl='soft'){const b=document.createElement('button');b.type='button';b.className='btn '+cl;b.textContent=text;b.addEventListener('click',cb);return b;}
function link(text,url,cl='soft'){const a=document.createElement('a');a.className='btn '+cl;a.textContent=text;a.href=url;return a;}
function filtradas(){const q=el('busca').value.trim().toLocaleLowerCase('pt-BR'),f=el('filtro').value;return obras.map((p,i)=>({p,i})).filter(({p})=>((f==='todas')||(f==='pendentes'&&!p.feito)||(f==='prioritarias'&&!p.feito&&p.prioridade)||(f==='feitas'&&p.feito))&&(!q||[p.nota,p.municipio,p.localidade,p.endereco].join(' ').toLocaleLowerCase('pt-BR').includes(q)));}
function renderList(){const box=el('lista');box.replaceChildren();const arr=filtradas();arr.slice(0,visiveis).forEach(({p,i})=>{const item=document.createElement('article');item.className='item'+(p.prioridade?' priority':'')+(p.feito?' done':'');const head=document.createElement('div');head.className='row';const label=document.createElement('strong');label.textContent=(i+1)+'. NOTA '+p.nota;head.append(label);const badge=document.createElement('span');badge.className='tag '+(p.feito?'green':p.prioridade?'red':'');badge.textContent=p.feito?'✓ Concluida':p.prioridade?'⚠ Prioritaria':'Pendente';head.append(badge);item.append(head);
const desc=document.createElement('p');desc.className='muted';desc.textContent=[p.municipio,p.localidade,p.endereco].filter(Boolean).join(' / ');item.append(desc);const gps=document.createElement('div');gps.className='coords';gps.textContent=coord(p);item.append(gps);const ctrl=document.createElement('div');ctrl.className='actions';const up=button('⬆ Subir',()=>mover(i,-1));up.disabled=i===0;ctrl.append(up);const down=button('⬇ Descer',()=>mover(i,1));down.disabled=i===obras.length-1;ctrl.append(down);ctrl.append(button(p.feito?'↩ Reabrir':'✓ Concluida',()=>toggle(p.id),p.feito?'soft':'green'));if(modo==='online'){ctrl.append(link('📍 Google Maps',urlDirecoes(data.origem,[p],el('localAtual').checked)));}item.append(ctrl);box.append(item);});
el('mais').classList.toggle('hidden',arr.length<=visiveis);if(!arr.length){box.textContent='Nenhuma obra com este filtro.';}}
function renderRoutes(){const b=el('etapas');b.replaceChildren();const groups=routeGroups();groups.forEach((g,i)=>{const p=document.createElement('div');p.className='item';const text=document.createElement('div');text.textContent='Etapa '+(i+1)+' | '+g.subset.length+' parada(s) | '+g.subset.map(x=>x.nota).join(' → ');p.append(text);p.append(link('Abrir etapa '+(i+1),urlDirecoes(g.de,g.subset,el('localAtual').checked&&i===0),'green'));b.append(p);});
el('rotaProximas').href=groups.length?urlDirecoes(groups[0].de,groups[0].subset,el('localAtual').checked):'#';el('rotaProximas').style.display=groups.length?'inline-block':'none';}
function render(){el('qtd').textContent=obras.length;el('feitas').textContent=obras.filter(p=>p.feito).length;el('pend').textContent=pendentes().length;const first=pendentes()[0];el('primeira').textContent=first?'Proxima obra: '+first.nota+' | '+first.municipio+(first.prioridade?' | PRIORITARIA':''):'Todas as obras concluidas!';if(first){el('linkPrimeira').href=urlDirecoes(data.origem,[first],el('localAtual').checked);el('linkPrimeira').style.display='';}else el('linkPrimeira').style.display='none';renderRoutes();renderList();persist();}
function setMode(m){modo=m;el('modoOffline').classList.toggle('active',m==='offline');el('modoOnline').classList.toggle('active',m==='online');el('blocoOnline').classList.toggle('hidden',m!=='online');el('modoTag').textContent=m==='offline'?'PLANEJAMENTO OFFLINE':'GOOGLE MAPS ONLINE';renderList();}
function statusConexao(){el('conexao').textContent=navigator.onLine?'Conexao possivelmente disponivel':'Sem conexao detectada';}
el('exportGPX').onclick=exportGpx;el('exportKML').onclick=exportKml;el('exportCSV').onclick=exportCsv;el('exportJSON').onclick=exportJson;
el('importJSON').addEventListener('change',async ev=>{const f=ev.target.files[0];if(!f)return;try{if(f.size>5*1024*1024)throw Error('Backup maior que 5 MB');const s=JSON.parse(await f.text());if(s.schema!=='nip-offline-v1'||s.equipe!==data.equipe||s.versao!==data.versao||!validState(s))throw Error('Backup invalido ou de outra equipe/versao');if(!confirm('Substituir ordem e concluidas pelo backup selecionado?'))return;applyState(s);render();toast('Backup restaurado com sucesso');}catch(e){alert('Falha ao restaurar: '+e.message);}finally{ev.target.value='';}});
el('reset').onclick=()=>{if(confirm('Restaurar a ordem original e desmarcar todas as concluidas?')){obras=original.map(p=>({...p}));visiveis=100;render();}};
el('busca').oninput=()=>{visiveis=100;renderList();};el('filtro').onchange=()=>{visiveis=100;renderList();};el('mais').onclick=()=>{visiveis+=100;renderList();};el('localAtual').onchange=renderRoutes;
el('modoOffline').onclick=()=>setMode('offline');el('modoOnline').onclick=()=>setMode('online');
window.addEventListener('online',statusConexao);window.addEventListener('offline',statusConexao);statusConexao();el('listaEstatica').classList.add('hidden');setMode('offline');render();
__PWA_JS__
})();</script></body></html>"""
    return (template.replace('__DADOS__', dados_json)
            .replace('__PWA_HEAD__', pwa_head)
            .replace('__PWA_JS__', pwa_script)
            .replace('__EQUIPE_HTML__', html.escape(str(equipe)))
            .replace('__ORIGEM_HTML__', html.escape(origem['municipio']))
            .replace('__QTD__', str(len(pontos)))
            .replace('__NOTA_HTML__', html.escape(pontos[0]['nota'] if pontos else '-'))
            .replace('__LISTA_ESTATICA__', fallback_html)
            .replace('__PRIMEIRA__', html.escape(fallback_primeira, quote=True))
            .replace('__GPS__', html.escape(fallback_gps, quote=True))
            .replace('__ETAPA__', html.escape(fallback_etapa, quote=True)))



def link_apple_maps_obra(origem, destino, local_atual=False):
    """Apple Maps route to a single work location, without an external service."""
    params = {'daddr': coordenada_maps(destino), 'dirflg': 'd'}
    if not local_atual:
        params['saddr'] = coordenada_maps(origem)
    return 'https://maps.apple.com/?' + urlencode(params)


def gerar_painel_iphone_estatico(df_equipe, equipe):
    """No-JS fallback: plain HTTPS hrefs that can be followed or copied on iOS.

    iOS Files/WhatsApp HTML previews can omit scripts. Static markup is
    intentionally present in the document before any JavaScript runs.
    """
    ori = origem_google_maps(df_equipe)
    pontos = pontos_equipe_google_maps(df_equipe)
    safe = lambda x: html.escape(str(x), quote=True)
    css = ("body{font-family:-apple-system,system-ui,sans-serif;margin:0;"
           "background:#f2f6fb;color:#18314f}header{background:#0d256c;"
           "color:white;padding:22px 14px}main{max-width:720px;margin:auto;"
           "padding:14px}.card{background:white;border:1px solid #d6e2ee;"
           "border-radius:10px;margin:12px 0;padding:15px}.b{display:inline-block;"
           "background:#1b8049;color:white;text-decoration:none;border-radius:8px;"
           "padding:12px;margin:5px 5px 5px 0;overflow-wrap:anywhere;font-weight:650}"
           ".b.alt{background:#0d256c}.note{color:#4b6078;font-size:14px}"
           "h1{font-size:24px}h2{font-size:18px;margin:0 0 9px}")
    b = ['<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">',
         '<meta name="viewport" content="width=device-width,initial-scale=1">',
         '<title>Rotas iPhone - NIP</title><style>'+css+'</style></head><body>',
         '<header><h1>Rotas iPhone / iOS</h1><strong>'+safe(equipe)+'</strong>',
         '<p>Origem: '+safe(ori['municipio'])+'</p></header><main>',
         '<p class="note">Links diretos: nao dependem de botoes JavaScript. '
         'Se o visor do app Arquivos nao abrir o link, copie o endereco do '
         'arquivo LINKS_Maps desta equipe e cole no Safari.</p>']
    if pontos:
        primeira = pontos[0]
        b.extend(['<section class="card"><h2>1. Comece por esta obra</h2>',
                  '<strong>Nota '+safe(primeira['nota'])+'</strong>',
                  '<p class="note">'+safe(primeira['municipio'])+'</p>',
                  '<a class="b" href="'+safe(link_google_maps_primeira_obra(ori,primeira))+'">Google Maps - desde municipio</a>',
                  '<a class="b" href="'+safe(link_google_maps_primeira_obra(ori,primeira,local_atual=True))+'">Google Maps - usar GPS atual</a>',
                  '<a class="b alt" href="'+safe(link_apple_maps_obra(ori,primeira))+'">Apple Maps - primeira obra</a>',
                  '<p class="note">Se o toque nao funcionar, copie este link para o Safari: '+safe(link_google_maps_primeira_obra(ori, primeira))+'</p>',
                  '</section>'])
        etapas = gerar_etapas_google_maps(ori, pontos)
        b.append('<section class="card"><h2>2. Abrir etapas no Google Maps</h2>')
        for e in etapas:
            b.extend(['<p><strong>Etapa '+str(e['etapa'])+'</strong> - notas '+safe(e['primeira_nota'])+' ate '+safe(e['ultima_nota'])+'</p>',
                      '<a class="b" href="'+safe(e['url'])+'">Abrir etapa '+str(e['etapa'])+'</a>',
                      '<a class="b alt" href="'+safe(e['url_local_atual'])+'">Etapa '+str(e['etapa'])+' a partir do GPS</a>',
                      '<p class="note">Link: '+safe(e['url'])+'</p>'])
        b.append('</section>')
        b.append('<section class="card"><h2>3. Navegar obra por obra</h2>')
        atual = ori
        for i, pt in enumerate(pontos, start=1):
            b.extend(['<p><strong>'+str(i)+'. Nota '+safe(pt['nota'])+'</strong> '
                      + ('(PRIORITARIA)' if pt['prioridade'] else '')+'</p>',
                      '<a class="b" href="'+safe(link_google_maps_primeira_obra(atual,pt))+'">Google Maps</a>',
                      '<a class="b alt" href="'+safe(link_apple_maps_obra(atual,pt))+'">Apple Maps</a>',
                      '<a class="b alt" href="'+safe(link_google_maps_primeira_obra(atual,pt,local_atual=True))+'">Do meu GPS</a>'])
            atual = pt
        b.append('</section>')
    else:
        b.append('<p>Nenhuma obra disponivel para esta equipe.</p>')
    b.append('<p class="note">A origem e a sede municipal de referencia, '
             'nao o endereco residencial. Para modificar a sequencia no iPhone, '
             'abra o planejador interativo em uma pagina HTTPS no Safari. '
             'Os links deste arquivo seguem a ordem original da NIP.</p>')
    b.append('</main></body></html>')
    return '\n'.join(b)


def gerar_leia_me_iphone():
    return ("ROTAS NO IPHONE / iOS - NIP\n\n"
            "1. Extraia o ZIP no app Arquivos do iPhone.\n"
            "2. Abra o arquivo Equipes/Roteiro_iPhone_*.html da sua equipe.\n"
            "3. Toque em Google Maps - desde municipio, ou Google Maps - usar GPS atual.\n"
            "4. Se a visualizacao do iPhone nao permitir abrir o link, abra\n"
            "   Equipes/LINKS_Maps_*.txt e COPIE o link https:// para o Safari.\n"
            "5. Apple Maps e uma alternativa para percursos individuais.\n"
            "6. Para mudar e salvar a ordem pelo painel interativo, use uma\n"
            "   versao hospedada em HTTPS no Safari (a visualizacao rapida do\n"
            "   app Arquivos pode nao executar JavaScript).\n"
            "7. A origem municipal e uma referencia, nao um endereco de casa.\n"
            "8. Para navegar, o celular precisa de internet ou mapas offline\n"
            "   previamente disponiveis no app de navegacao.\n")


def csv_google_mymaps_em_partes(texto_csv, max_dados=1999):
    """Limite do My Maps: ate 2000 linhas por CSV, incluindo o cabecalho."""
    if max_dados < 1 or max_dados > 1999:
        raise ValueError('max_dados deve ficar entre 1 e 1999')
    valores = list(csv.DictReader(io.StringIO(texto_csv.lstrip('\ufeff'))))
    if not valores:
        return []
    campos = list(valores[0])
    saidas = []
    for i in range(0, len(valores), max_dados):
        trecho = io.StringIO(newline='')
        escritor = csv.DictWriter(trecho, fieldnames=campos, lineterminator='\r\n')
        escritor.writeheader()
        escritor.writerows(valores[i:i+max_dados])
        saidas.append(trecho.getvalue())
    return saidas


def gerar_indice_google_maps(equipes):
    """Pagina inicial do ZIP para abrir o planejamento de qualquer equipe."""
    blocos=[]
    for equipe in equipes:
        nome=html.escape(equipe['equipe'])
        arq=html.escape(equipe['arquivo'], quote=True)
        local=html.escape(equipe['municipio'])
        nota=html.escape(equipe['primeira'])
        arq_ios=html.escape(equipe.get('arquivo_ios',equipe['arquivo']), quote=True)
        blocos.append(f'<article class="item"><strong>{nome}</strong><p>🏠 {local} | 1ª obra: {nota}</p><a href="{arq_ios}">iPhone: abrir rotas diretas</a> &nbsp; <a href="{arq}">Planejador interativo (Android / Safari HTTPS)</a></article>')
    return ("""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Google Maps - NIP</title><style>body{font-family:system-ui;background:#f3f6fd;margin:0;color:#123}header{background:#0d256c;color:white;padding:22px}main{max-width:900px;margin:auto;padding:15px}.item{background:white;border-radius:9px;padding:16px;margin:12px 0;box-shadow:0 2px 8px #1231}.item a{background:#238b46;color:white;display:inline-block;padding:11px;border-radius:7px;text-decoration:none;font-weight:bold}p{color:#567}</style></head><body><header><h1>📍 Planejamento Google Maps - NIP</h1><p style="color:white">Selecione a equipe para organizar o roteiro no celular</p></header><main>"""
            + ''.join(blocos) + """<p>O arquivo HTML de cada equipe permite reordenar e marcar visitas. Para importar os pontos num mapa, use os arquivos CSV em Google My Maps.</p></main></body></html>""")


def gerar_leia_me_google_maps():
    return ("""ROTAS GOOGLE MAPS - NIP / LISTA CONTINUA

1. Abra INICIAR_AQUI.html no navegador do computador ou do celular.
   NO IPHONE use Equipes/Roteiro_iPhone_*.html, com links estaticos.
   Se o app Arquivos nao abrir links, copie os URLs de LINKS_Maps_*.txt
   e cole diretamente no Safari. Leia LEIA_ME_IPHONE.txt.
2. Selecione a equipe; no HTML individual, a primeira obra sugerida ja aparece.
3. Use Subir/Descer para alterar a ordem, Concluida para controlar o realizado.
4. Clique em Abrir proximas 4 obras no Maps, ou abra cada etapa sugerida.
5. Para partir de onde o celular estiver, marque Usar minha localizacao atual.
6. Os links em LINKS_*.txt servem para envio direto por WhatsApp.
7. Para visualizar TODOS os marcadores de cada equipe num mesmo mapa no
   Google My Maps (https://www.google.com/maps/d/), importe o CSV correspondente.
   Escolha Latitude e Longitude como colunas de posicao e Nome como titulo.
   Se houver mais de 1999 registros por CSV, importe as partes _PARTE_002 etc.

IMPORTANTE:
- Google Maps normal nao importa KML/KMZ como trajeto editavel. Esses sao para Google Earth.
- HTML, CSV e TXT sao gerados separados por equipe; INICIAR_AQUI lista todas.
- Cada link inclui no maximo quatro obras (ate 3 intermediarias + 1 destino),
  para compatibilidade com navegadores de celular. Etapas seguintes continuam
  a partir da ultima obra da etapa anterior, se a equipe nao marcar origem GPS.
- A sequencia planejada pode ser ajustada no Google Maps ou no HTML da equipe.
- O Google Maps calcula percursos e tempos pela malha viaria disponivel.
- O primeiro trajeto do roteiro NIP usa a sede municipal como origem, nao
  o endereco exato da casa ou hospedagem.
- Sem sincronizacao online: a ordem editada e o status de concluido no HTML
  ficam apenas no navegador local, quando o armazenamento local e permitido.
- Em dispositivos que nao abrem HTML do ZIP diretamente, extraia o ZIP primeiro
  e abra o HTML no navegador; os arquivos TXT contem links alternativos.
""")



def gerar_gpx_offline_paradas(df_equipe, equipe):
    """GPX apenas de waypoints e route points (nao e desenho de estrada).

    Organic Maps e OsmAnd podem importar; o aplicativo recalcula as estradas
    com seus proprios mapas. Nao gerar trkpt ficticio ligando obras em reta.
    """
    from xml.sax.saxutils import escape as xml_escape
    origem = origem_google_maps(df_equipe)
    pontos = pontos_equipe_google_maps(df_equipe)
    x = lambda v: xml_escape(str(v))
    def item(tag, nome, obj, info):
        return (f'<{tag} lat="{float(obj["lat"]):.7f}" lon="{float(obj["lon"]):.7f}">'
                f'<name>{x(nome)}</name><desc>{x(info)}</desc></{tag}>')
    partes = ['<?xml version="1.0" encoding="UTF-8"?>',
              '<gpx version="1.1" creator="NIP Campo Offline" xmlns="http://www.topografix.com/GPX/1/1">',
              f'<metadata><name>{x(equipe)}</name><desc>Paradas em ordem, sem rota rodoviaria.</desc></metadata>',
              item('wpt', f'CASINHA - {origem["municipio"]}', origem, 'Origem: sede municipal aproximada')]
    for n, ponto in enumerate(pontos, 1):
        partes.append(item('wpt', f'{n:03d} - {ponto["nota"]}', ponto,
                            ponto['municipio'] + (' | PRIORIDADE' if ponto['prioridade'] else '')))
    partes.append(f'<rte><name>{x(equipe)} - paradas planejadas</name>'
                  '<desc>Sequencia logica. O app offline calcula o trajeto pelas ruas.</desc>')
    partes.append(item('rtept', f'ORIGEM - {origem["municipio"]}', origem, 'Origem municipal'))
    for n, ponto in enumerate(pontos, 1):
        partes.append(item('rtept', f'{n:03d} - {ponto["nota"]}', ponto, ponto['municipio']))
    partes.append('</rte></gpx>')
    return '\n'.join(partes)


def gerar_pwa_offline(equipe, versao, nome_arquivo_html, conteudo_html):
    """Devolve arquivos de uma PWA autocontida, para publicacao HTTPS opcional.

    O ZIP local funciona para consulta; instalar e guardar via service worker so
    funciona no navegador apos publicar esta pasta em HTTPS (ou localhost).
    """
    from PIL import Image
    nome = texto_seguro_maps(equipe)
    grupo = hashlib.sha256(nome.encode('utf-8')).hexdigest()[:8]
    prefixo = f'nip-campo-offline-{grupo}-'
    chave = hashlib.sha256((nome + versao).encode('utf-8')).hexdigest()[:16]
    manifest = {
        'name': f'NIP Campo - {nome}', 'short_name': 'NIP Campo',
        'description': f'Planejamento offline da equipe {nome}',
        'start_url': './index.html', 'scope': './', 'display': 'standalone',
        'background_color': '#f3f6fa', 'theme_color': '#0d256c',
        'icons': [{'src': './icon-192.png', 'sizes': '192x192', 'type': 'image/png'},
                  {'src': './icon-512.png', 'sizes': '512x512', 'type': 'image/png'}]
    }
    sw = f"""const CACHE='{prefixo}{chave}';
const PREFIX='{prefixo}';
const ASSETS=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)).then(()=>self.skipWaiting())));
self.addEventListener('activate',e=>e.waitUntil(Promise.all([caches.keys().then(a=>Promise.all(a.filter(n=>n.startsWith(PREFIX)&&n!==CACHE).map(n=>caches.delete(n))))),self.clients.claim()])));
self.addEventListener('fetch',e=>{{if(e.request.method!=='GET')return;const u=new URL(e.request.url);if(u.origin!==self.location.origin)return;
e.respondWith(caches.match(e.request).then(hit=>hit||fetch(e.request)));}});
"""
    img = Image.open(io.BytesIO(imagem_casa_bytes())).convert('RGBA')
    icons = {}
    for sz in [192, 512]:
        surface = Image.new('RGBA', (sz, sz), '#0d256c')
        casa = img.resize((int(sz*.84), int(sz*.84)), Image.Resampling.LANCZOS)
        surface.alpha_composite(casa, ((sz-casa.width)//2, (sz-casa.height)//2))
        buff = io.BytesIO(); surface.convert('RGB').save(buff, format='PNG'); icons[f'icon-{sz}.png'] = buff.getvalue()
    return {'index.html': conteudo_html.encode('utf-8'),
            'manifest.webmanifest': json.dumps(manifest, ensure_ascii=False, indent=2).encode('utf-8'),
            'sw.js': sw.encode('utf-8'), **icons}


def gerar_guia_modo_offline():
    return ("""NIP CAMPO - PLANEJAMENTO ONLINE E OFFLINE (ANDROID / IPHONE)

MODO OFFLINE (sem internet):
1. No escritorio: gere o ZIP Planejamento Offline ou Google Maps.
2. Baixe e EXTRAIA o ZIP no celular enquanto houver conexao.
3. Para consulta, abra Equipes/Planejador_*.html em um navegador que execute JavaScript.
4. No iPhone, a pre-visualizacao do app Arquivos pode BLOQUEAR JavaScript.
   Para edicao confiavel no iPhone, publique a pasta Offline_Instalavel/<equipe>
   em um servidor HTTPS. Abra index.html no Safari, use Adicionar a Tela de Inicio.
   A PWA usa cache de seus arquivos basicos; abra-a uma vez online antes de sair.
5. No Android, o arquivo HTML pode funcionar no Chrome/Firefox dependendo da
   permissao de abrir arquivo local e permitir downloads. Se houver bloqueio,
   tambem use a pasta HTTPS/PWA instalavel.
6. A lista, prioridades, ordem, conclusao e exportacao dos arquivos podem
   funcionar offline se o navegador executar JavaScript.
7. Exporte BACKUP.json para guardar a ordem e as concluidas, e CSV para o escritorio.
   Um backup so pode ser restaurado na mesma equipe e mesma versao da rota.

NAVEGACAO OFFLINE POR ESTRADAS:
1. Instale Organic Maps ou OsmAnd enquanto tiver internet.
2. Dentro desses aplicativos, baixe os MAPAS das regioes do Maranhao e Piaui em que
   havera deslocamento. Confira cobertura dos acessos/estradas previamente.
3. Em Equipes, importe Pontos_Offline_*.gpx no app de navegacao.
   Android: escolha abrir/importar pelo Organic Maps ou OsmAnd.
   iOS: abra Arquivos > Compartilhar > Importar com Organic Maps/OsmAnd.
4. O arquivo GPX contem coordenadas e SEQUENCIA de paradas, nao um trajeto
   calculado pelas ruas. O app offline cria a rota usando sua malha viaria.
5. Se mudar a ordem / baixar obras no HTML, clique em GPX das pendentes
   e importe a nova lista no app. Nao ha sincronizacao automatica.

MODO ONLINE (GOOGLE MAPS):
- Dentro do Planejador selecione Google Maps online.
- Abra a primeira obra ou etapas com ate quatro obras por link.
- Sao multiplas etapas, sem limite fixo de obras no planejamento.
- Dependendo do aplicativo e da conectividade, o link pode nao abrir offline.

PWA/INSTALACAO:
- O pacote NAO publica o app na internet automaticamente.
- ATENCAO: as notas e coordenadas sao dados operacionais; se publicar em
  HTTPS, use hospedagem privada com controle de acesso por equipe, nunca um
  site publico sem autenticacao.
- HTTPS e necessario para ativar cache/service worker da PWA no iPhone/Android.
- Os dados do HTML nao incluem mapa-base/estradas nem motor rodoviario embutido.
- Para navegacao totalmente sem rede, use aplicativo GPS offline com mapas
  previamente baixados. O GPS do telefone em si nao exige dados moveis.

Fontes oficiais (para instrucoes de importacao):
https://organicmaps.app/faq/bookmarks/how-to-import/
https://osmand.net/docs/user/personal/tracks/manage-tracks/
https://developers.google.com/maps/documentation/urls/get-started
""")


def normalize_cols(columns):
    def clean(x):
        x=str(x).strip().upper()
        x=unicodedata.normalize('NFKD',x)
        return ''.join(c for c in x if not unicodedata.combining(c)).replace('_',' ')
    return [clean(x) for x in columns]

def formata_campo_html(v):
    if pd.isna(v): return '-'
    return html.escape(str(v))

def formatar_moeda(v):
    try: return f'R$ {float(v):,.2f}'
    except (TypeError,ValueError): return '-'

def normalizar_municipios(df):
    return df

@st.cache_data(show_spinner=False)
def ler_planilha_cached(raw):
    return pd.read_excel(io.BytesIO(raw),engine='openpyxl')

def injetar_logo():
    for path in ['LOGO_NIP.png','logo_nip.png','assets/logo_nip.png']:
        if os.path.exists(path):
            st.logo(path)
            break

def adaptar_planilha_nip(df):
    """Adapta tanto a planilha Saneamento quanto a Lista Continua original.

    O campo COLABORADOR PLANEJADO (K) vale somente para Saneamento;
    a Lista Continua original utiliza LEVANTADOR (A).
    """
    df = df.copy()
    df.columns = normalize_cols(df.columns)
    cols = set(df.columns)

    # A normalizacao retira acentos dos cabecalhos Excel: INSTALAÇÃO -> INSTALACAO.
    # Recupera o nome usado no seletor de KML sem perder os dados da base.
    # Executado antes da fusao de superpontos, preservando cada instalacao individual.
    if 'INSTALACAO' in df.columns:
        if 'INSTALAÇÃO' not in df.columns:
            df['INSTALAÇÃO'] = df['INSTALACAO']
        else:
            sem_instalacao = df['INSTALAÇÃO'].isna() | df['INSTALAÇÃO'].astype(str).str.strip().str.upper().isin(['', 'NAN', 'NONE', 'NULL', '-'])
            df.loc[sem_instalacao, 'INSTALAÇÃO'] = df.loc[sem_instalacao, 'INSTALACAO']

    # Algumas bases fornecidas usam ID ISCO em vez de ID SISCO.
    # A coluna original e mantida; o nome padrao e usado nos cards/KML.
    if 'ID SISCO' not in df.columns:
        df['ID SISCO'] = df['ID ISCO'] if 'ID ISCO' in df.columns else ''
    elif 'ID ISCO' in df.columns:
        vazio_id = df['ID SISCO'].isna() | df['ID SISCO'].astype(str).str.strip().isin(['', 'nan', 'None'])
        df.loc[vazio_id, 'ID SISCO'] = df.loc[vazio_id, 'ID ISCO']

    # Padroniza o nome do cliente sem depender da posicao das colunas.
    # Mantem a coluna original e agrega a informacao em NOME para exportacoes.
    nomes_alternativos = ('NOME', 'NOME DO CLIENTE', 'NOME CLIENTE',
                          'CLIENTE', 'NOME SOLICITANTE', 'SOLICITANTE')
    if 'NOME' not in df.columns:
        df['NOME'] = pd.Series('', index=df.index, dtype='object')
    for coluna_nome in nomes_alternativos:
        if coluna_nome != 'NOME' and coluna_nome in cols:
            atual = df['NOME'].astype('string').fillna('').str.strip()
            vazio = atual.eq('') | atual.str.upper().isin(['NAN', 'NONE', 'NULL', '-'])
            df.loc[vazio, 'NOME'] = df.loc[vazio, coluna_nome]

    if 'COLABORADOR PLANEJADO' in cols:
        # Modelo Saneamento: a coluna K identifica o colaborador planejado.
        if len(df.columns) < 11 or df.columns[10] != 'COLABORADOR PLANEJADO':
            raise ValueError('Na base de Saneamento, a coluna K deve ser COLABORADOR PLANEJADO.')
        rename = {
            'MUNICIPIO': 'MUNICIPIO', 'REGIONAL SERVICO': 'REGIONAL',
            'ENDERECO': 'ENDEREÇO', 'MATRICULA': 'MATRICULA',
            'COLABORADOR PLANEJADO': 'COLABORADOR',
        }
        df = df.rename(columns=rename)
        needed = ['NOTA', 'STATUS', 'TIPO DEMANDA', 'MUNICIPIO', 'REGIONAL',
                  'ENDEREÇO', 'BAIRRO', 'LATITUDE', 'LONGITUDE', 'COLABORADOR']
        missing = [c for c in needed if c not in df.columns]
        if missing:
            raise ValueError('Colunas ausentes na base Saneamento: ' + ', '.join(missing))
        if 'MATRICULA' not in df.columns:
            df['MATRICULA'] = ''
        df['PRIORIDADE'] = 'Não'
    elif {'LEVANTADOR', 'PROTOCOLO', 'LATITUDE', 'LONGITUDE'}.issubset(cols):
        # Modelo original Lista Continua: nao exigir COLABORADOR PLANEJADO na K.
        df['NOTA'] = df['PROTOCOLO']
        df['STATUS'] = df['STATUS LIST'] if 'STATUS LIST' in df.columns else (
            df['STATUS SISCO'] if 'STATUS SISCO' in df.columns else '')
        df['TIPO DEMANDA'] = df['TIPO NOTA'] if 'TIPO NOTA' in df.columns else ''
        df['COLABORADOR'] = df['LEVANTADOR']
        df['BAIRRO'] = df['LOCALIDADE'] if 'LOCALIDADE' in df.columns else ''
        if 'ENDERECO' in df.columns and 'ENDEREÇO' not in df.columns:
            df = df.rename(columns={'ENDERECO': 'ENDEREÇO'})
        if 'MATRICULA' not in df.columns:
            df['MATRICULA'] = ''
        if 'REGIONAL' not in df.columns:
            df['REGIONAL'] = ''
        if 'PRIORIDADE' not in df.columns:
            df['PRIORIDADE'] = 'Não'
    else:
        raise ValueError(
            'Formato nao reconhecido. Envie a base Saneamento com COLABORADOR PLANEJADO '
            'na coluna K ou a Lista Continua com LEVANTADOR e PROTOCOLO.'
        )

    # A BASE_LISTA_CONTINUA_LEVANTAMENTO pode conter latitude e longitude
    # trocadas nos CABECALHOS. Corrige por linha apenas quando os valores
    # correspondem inequivocamente a uma localizacao no Brasil (sem presumir
    # que todas as planilhas possuem essa inversao).
    if 'LATITUDE' in df.columns and 'LONGITUDE' in df.columns:
        lat_num = pd.to_numeric(df['LATITUDE'].astype('string').str.replace(',', '.', regex=False), errors='coerce')
        lon_num = pd.to_numeric(df['LONGITUDE'].astype('string').str.replace(',', '.', regex=False), errors='coerce')
        # Latitude brasileira: -34 a +6. Longitude brasileira: -75 a -30.
        # Para evitar trocas ambiguas, exige latitude aparente fora da faixa
        # brasileira e longitude aparente dentro da faixa de latitude.
        invertidas = (lat_num.between(-75, -34, inclusive='both')
                      & lon_num.between(-34, 6, inclusive='both'))
        if invertidas.any():
            df.loc[invertidas, 'LATITUDE'] = lon_num.loc[invertidas].to_numpy()
            df.loc[invertidas, 'LONGITUDE'] = lat_num.loc[invertidas].to_numpy()
        df.attrs['coordenadas_invertidas_corrigidas'] = int(invertidas.sum())
    else:
        df.attrs['coordenadas_invertidas_corrigidas'] = 0

    # Mantem os dados operacionais originais para os cards e exportadores.
    for c in CAMPOS_VISIVEIS_NIP:
        if c not in df.columns:
            df[c] = ''
    df['PROTOCOLO'] = df['NOTA'].astype('string').str.replace(r'\.0$', '', regex=True)
    df['BASE_ATRIBUIDA'] = df['COLABORADOR'].fillna('').astype(str).str.strip().str.upper()
    if 'TIPO NOTA' not in df.columns:
        df['TIPO NOTA'] = df['TIPO DEMANDA']
    if 'LOCALIDADE' not in df.columns:
        df['LOCALIDADE'] = df['BAIRRO']
    return df


def haversine_vectorized(lat1, lon1, lat2, lon2):
    R = 6371.0 
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = np.sin(dlat/2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2.0)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return R * c


def haversine_scalar(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c



# ---------- Origem municipal para roteirizacao de equipes (MA + PI) ----------
# Fonte: dados abertos do repositorio kelvins/municipios-brasileiros,
# com referencias municipais e codigos de UF alinhados a divisao do IBGE.
# Sao 217 municipios maranhenses (codigo 21) e 224 piauienses (codigo 22).
# As coordenadas representam a sede municipal, nao o endereco da residencia.
FONTES_MUNICIPIOS = (
    'https://raw.githubusercontent.com/kelvins/municipios-brasileiros/master/csv/municipios.csv',
    'https://raw.githubusercontent.com/projectgithubrepository/Municipios-Brasileiros/master/csv/municipios.csv',
)


@st.cache_data(ttl=86400, show_spinner=False)
def carregar_municipios_ma_pi():
    """Retorna as 441 sedes municipais como 'Nome / MA' e 'Nome / PI'.

    Pode funcionar sem conexao caso o arquivo 'municipios_ma_pi.csv' seja
    colocado na mesma pasta deste .py, contendo as colunas: nome, latitude,
    longitude e codigo_uf (21 ou 22), ou codigo_ibge, ou uf (MA ou PI).
    """
    fontes = []
    csv_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'municipios_ma_pi.csv')
    if os.path.isfile(csv_local):
        fontes.append(('arquivo local', csv_local))
    fontes.extend(('online', url) for url in FONTES_MUNICIPIOS)
    erros = []

    for tipo, origem in fontes:
        try:
            if tipo == 'arquivo local':
                dados = pd.read_csv(origem, encoding='utf-8-sig')
            else:
                resposta = requests.get(
                    origem, timeout=(5, 20),
                    headers={'User-Agent': 'NIP-ListaContinua/2.1', 'Accept': 'text/csv'}
                )
                resposta.raise_for_status()
                dados = pd.read_csv(io.BytesIO(resposta.content), encoding='utf-8-sig')

            dados.columns = [str(c).strip().lower() for c in dados.columns]
            if not {'nome', 'latitude', 'longitude'}.issubset(dados.columns):
                raise ValueError('CSV sem as colunas nome, latitude ou longitude')

            if 'codigo_uf' in dados.columns:
                codigo = pd.to_numeric(dados['codigo_uf'], errors='coerce')
                dados['UF_PARTIDA'] = codigo.map({21: 'MA', 22: 'PI'})
            elif 'codigo_ibge' in dados.columns:
                codigos = dados['codigo_ibge'].astype(str).str.replace(r'\.0$', '', regex=True)
                dados['UF_PARTIDA'] = codigos.str[:2].map({'21': 'MA', '22': 'PI'})
            elif 'uf' in dados.columns:
                dados['UF_PARTIDA'] = dados['uf'].astype(str).str.strip().str.upper()
            else:
                raise ValueError('CSV sem codigo_uf, codigo_ibge ou uf para diferenciar os estados')

            dados = dados[dados['UF_PARTIDA'].isin(['MA', 'PI'])].copy()
            dados['latitude'] = pd.to_numeric(dados['latitude'], errors='coerce')
            dados['longitude'] = pd.to_numeric(dados['longitude'], errors='coerce')
            dados = dados.dropna(subset=['nome', 'latitude', 'longitude'])
            dados = dados[
                dados['latitude'].between(-12, 0)
                & dados['longitude'].between(-50, -39)
            ].copy()
            dados['nome'] = dados['nome'].astype(str).str.strip()
            dados = dados[dados['nome'].ne('')]
            dados = dados.drop_duplicates(subset=['UF_PARTIDA', 'nome'])

            quantidade_ma = int((dados['UF_PARTIDA'] == 'MA').sum())
            quantidade_pi = int((dados['UF_PARTIDA'] == 'PI').sum())
            if quantidade_ma != 217 or quantidade_pi != 224:
                raise ValueError(
                    f'Lista incompleta: MA={quantidade_ma}/217 e PI={quantidade_pi}/224'
                )

            municipios = {
                f'{row.nome} / {row.UF_PARTIDA}': (float(row.latitude), float(row.longitude))
                for row in dados.itertuples(index=False)
            }
            return dict(sorted(
                municipios.items(),
                key=lambda item: (normalizar_texto_municipio(item[0].rsplit(' / ', 1)[0]), item[0])
            ))
        except (OSError, ValueError, KeyError, requests.RequestException) as exc:
            erros.append(f'{tipo}: {exc}')

    raise RuntimeError(
        'Nao foi possivel carregar os municipios de MA e PI (217 + 224). '
        'Confira sua conexao ou coloque municipios_ma_pi.csv ao lado do .py. '
        + ' | '.join(erros)
    )


def normalizar_texto_municipio(texto):
    s = unicodedata.normalize('NFKD', str(texto).upper())
    return ''.join(c for c in s if not unicodedata.combining(c))


def eh_prioridade(valor):
    """Reconhece SIM/1/True, sem transformar vazios e NAO em prioridade."""
    if pd.isna(valor):
        return False
    texto = normalizar_texto_municipio(str(valor).strip())
    return texto not in ('', '0', '0.0', 'NAO', 'N', 'FALSE', 'F', 'NAN', 'NONE', 'NULL', '-')


def escolher_primeira_obra(obras, lat_origem, lon_origem):
    """Escolhe obra prioritária mais proxima; senao, mais proxima geral.

    Criterio de proximidade = distancia geodesica (Haversine). O tempo por
    estradas sera calculado no motor caso o OSRM esteja habilitado.
    """
    if not obras:
        return None, float('nan')
    prioritarias = [o for o in obras if eh_prioridade(o.get('PRIORIDADE'))]
    candidatas = prioritarias if prioritarias else obras
    lats = np.array([float(o['LATITUDE']) for o in candidatas], dtype=float)
    lons = np.array([float(o['LONGITUDE']) for o in candidatas], dtype=float)
    distancias = haversine_vectorized(float(lat_origem), float(lon_origem), lats, lons)
    idx = int(np.argmin(distancias))
    return candidatas[idx], float(distancias[idx])


def formatar_tempo_deslocamento(minutos):
    """Formata duracao para leitura de equipes sem confundir com hora de chegada."""
    try:
        total = max(0, int(round(float(minutos))))
    except (ValueError, TypeError):
        return '-'
    horas, mins = divmod(total, 60)
    return f'{horas}h {mins:02d}min' if horas else f'{mins} min'


def calcular_previa_origens(df_obras, origens_por_equipe, municipios, velocidade_kmh=30.0):
    """Tabela de conferencia do primeiro atendimento (estimativa sem OSRM)."""
    resultados = []
    for equipe in sorted(df_obras['BASE_ATRIBUIDA'].unique()):
        municipio = origens_por_equipe.get(equipe)
        if municipio not in municipios:
            continue
        lat, lon = municipios[municipio]
        lista = df_obras[df_obras['BASE_ATRIBUIDA'] == equipe].to_dict('records')
        primeira, km_diretos = escolher_primeira_obra(lista, lat, lon)
        if primeira is None:
            continue
        km_estimados = km_diretos * 1.30
        minutos = km_estimados / max(float(velocidade_kmh), 1) * 60
        resultados.append({
            'EQUIPE (COLUNA P)': equipe,
            'MUNICIPIO DE PARTIDA': municipio,
            '1a NOTA': primeira.get('PROTOCOLO', '-'),
            'PRIORITARIA': 'Sim' if eh_prioridade(primeira.get('PRIORIDADE')) else 'Nao',
            'DISTANCIA ESTIMADA (KM)': round(km_estimados, 1),
            'TEMPO ESTIMADO': formatar_tempo_deslocamento(minutos),
        })
    return pd.DataFrame(resultados)

def fundir_super_pontos(df_tasks, raio_metros=5, agrupar_por_levantador=False):
    if df_tasks.empty or 'LATITUDE' not in df_tasks.columns or 'LONGITUDE' not in df_tasks.columns: return df_tasks, 0
    df_valid = df_tasks.dropna(subset=['LATITUDE', 'LONGITUDE']).copy()
    if df_valid.empty: return df_tasks, 0

    coords = np.radians(df_valid[['LATITUDE', 'LONGITUDE']].values)
    eps_rad = raio_metros / 6371000.0
    
    db = DBSCAN(eps=eps_rad, min_samples=1, algorithm='ball_tree', metric='haversine').fit(coords)
    df_valid['CLUSTER_ID'] = db.labels_
    
    if agrupar_por_levantador and 'LEVANTADOR' in df_valid.columns:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str) + "_" + df_valid['LEVANTADOR'].astype(str)
    else:
        df_valid['CLUSTER_GRP'] = df_valid['CLUSTER_ID'].astype(str)
    
    cluster_counts = df_valid['CLUSTER_GRP'].value_counts()
    single_clusters = cluster_counts[cluster_counts == 1].index
    multi_clusters = cluster_counts[cluster_counts > 1].index
    agrupado = []
    
    if len(multi_clusters) > 0:
        df_multi = df_valid[df_valid['CLUSTER_GRP'].isin(multi_clusters)]
        for c_id, group in df_multi.groupby('CLUSTER_GRP'):
            row_base = group.iloc[0].copy()
            qtd = len(group)
            orig_rows = group.to_dict('records')
            row_base['_ORIGINAL_ROWS'] = orig_rows
            
            def safe_list_join(col_name):
                itens = [str(x).strip() for x in group[col_name] if pd.notna(x) and str(x).lower() != 'nan']
                itens_unicos = list(dict.fromkeys(itens))
                return " | ".join(itens_unicos) if itens_unicos else "-"
            
            for col in group.columns:
                if col not in ['LATITUDE', 'LONGITUDE', 'CLUSTER_ID', 'CLUSTER_GRP', '_ORIGEM_BASE', 'PRIORIDADE']:
                    row_base[col] = safe_list_join(col)
                
            row_base['LATITUDE'] = group['LATITUDE'].mean()
            row_base['LONGITUDE'] = group['LONGITUDE'].mean()
            row_base['SUPER_PONTO'] = f"SIM ({qtd} un.)"
            if 'PRIORIDADE' in group.columns and 'Sim' in group['PRIORIDADE'].values:
                row_base['PRIORIDADE'] = 'Sim'
            agrupado.append(row_base)
            
    df_final_multi = pd.DataFrame(agrupado)
    df_single = df_valid[df_valid['CLUSTER_GRP'].isin(single_clusters)].copy()
    if not df_single.empty:
        df_single['SUPER_PONTO'] = "NÃO"
        dict_records = df_single.to_dict('records')
        df_single['_ORIGINAL_ROWS'] = [[r] for r in dict_records]
        
    df_final = pd.concat([df_single, df_final_multi], ignore_index=True).drop(columns=['CLUSTER_ID', 'CLUSTER_GRP'], errors='ignore')
    
    df_nan = df_tasks[df_tasks['LATITUDE'].isna() | df_tasks['LONGITUDE'].isna()].copy()
    if not df_nan.empty:
        df_nan['SUPER_PONTO'] = "NÃO"
        dict_records_nan = df_nan.to_dict('records')
        df_nan['_ORIGINAL_ROWS'] = [[r] for r in dict_records_nan]
        df_final = pd.concat([df_final, df_nan], ignore_index=True)
        
    return df_final, len(df_tasks) - len(df_final)


def formatar_planilha_openpyxl(writer, sheet_name, df_orig):
    """Aplica o padrão NIP (Azul Escuro) no arquivo Excel gerado."""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri')
    fill_superponto = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    font_prioridade = Font(color='FF0000', bold=True, name='Calibri')
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    worksheet.auto_filter.ref = worksheet.dimensions
    
    col_prio_idx = df_orig.columns.get_loc('PRIORIDADE') + 1 if 'PRIORIDADE' in df_orig.columns else None
    col_sp_idx = df_orig.columns.get_loc('SUPER_PONTO') + 1 if 'SUPER_PONTO' in df_orig.columns else None
    
    for r_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
        is_prio = col_prio_idx and str(row[col_prio_idx - 1].value).strip().upper() == 'SIM'
        is_sp = col_sp_idx and str(row[col_sp_idx - 1].value).strip().upper().startswith('SIM')
        for cell in row:
            if is_sp: cell.fill = fill_superponto
            if is_prio: cell.font = font_prioridade
    
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        worksheet.column_dimensions[column].width = min(max_length + 2, 60)


def gerar_excel_lista(df, colunas_originais=None):
    output = io.BytesIO()
    df_saida = df.loc[:, ~df.columns.duplicated()].copy()
    
    for col_name in ['PROTOCOLO', 'NOTA']:
        if col_name in df_saida.columns:
            df_saida = df_saida[~df_saida[col_name].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
            
    colunas_remover = ['ROTA_GEOMETRIA', '_HORA_INICIO_DT', '_HORA_FIM_DT', '_ORIGINAL_ROWS', '_ORIGEM_BASE', 'COR_ICONE', 'MUN_LIMPO', 'COORD_KEY']
    df_saida = df_saida.drop(columns=[c for c in colunas_remover if c in df_saida.columns], errors='ignore')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='Obras Roteirizadas')
        formatar_planilha_openpyxl(writer, 'Obras Roteirizadas', df_saida)
        
    return output.getvalue()


def gerar_excel_resumo_lista(df_resumo):
    output = io.BytesIO()
    df_resumo = df_resumo.loc[:, ~df_resumo.columns.duplicated()].copy()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo Operacional')
        
        # Formata o header do resumo
        workbook = writer.book
        worksheet = writer.sheets['Resumo Operacional']
        header_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, name='Calibri')
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.auto_filter.ref = worksheet.dimensions
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[column].width = min(max_length + 2, 40)
            
    return output.getvalue()


def limpar_colunas_lista(df_alvo, cols_originais=None):
    df_alvo = df_alvo.loc[:, ~df_alvo.columns.duplicated()].copy()
    
    if 'PROTOCOLO' in df_alvo.columns:
        if 'NOTA' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['NOTA'])
        df_alvo = df_alvo.rename(columns={'PROTOCOLO': 'NOTA'})
        
    if 'BASE_ATRIBUIDA' in df_alvo.columns:
        if 'FISCAL' in df_alvo.columns: df_alvo = df_alvo.drop(columns=['FISCAL'])
        df_alvo = df_alvo.rename(columns={'BASE_ATRIBUIDA': 'FISCAL'})
        
    # Colunas obrigatórias fixadas no início
    final_cols = ['FISCAL', 'DIA_SEMANA', 'DIA_MES', 'SUPER_PONTO', 'ORDEM',
                  'ORIGEM_MUNICIPIO', 'PRIORIDADE', 'TIPO_DESLOCAMENTO',
                  'DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM', 'TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN',
                  'TEMPO_VIAGEM_MINUTOS', 'DISTANCIA_PONTO_ANTERIOR_KM']
    if 'NOTA' in df_alvo.columns: final_cols.append('NOTA')
    
    if cols_originais is not None:
        for c in cols_originais:
            nome_coluna = 'NOTA' if c == 'PROTOCOLO' else c
            nome_coluna = 'FISCAL' if c == 'BASE_ATRIBUIDA' else nome_coluna
            
            if nome_coluna in df_alvo.columns and nome_coluna not in final_cols and nome_coluna not in ['LEVANTADOR', 'NOME DO LEVANTADOR', 'LEVANTADOR_RESPONSAVEL']:
                final_cols.append(nome_coluna)
    else:
        colunas_lixo = ['LINK_NAVEGACAO_OFFLINE', 'ROTA_GEOMETRIA', 'COORD_KEY', 'MUN_LIMPO', 'COR_ICONE', 'ALERTA_TOPOLOGIA', 'TEMPO_VIAGEM_MINUTOS', 'HORA_INICIO', 'HORA_FIM', 'CLUSTER_ID', 'CLUSTER_GRP', 'MLC', 'DISTANCIA_PROXIMO_PONTO_KM']
        for c in df_alvo.columns:
            if c not in final_cols and not str(c).startswith('_') and c not in colunas_lixo:
                final_cols.append(c)
                
    return df_alvo[[c for c in final_cols if c in df_alvo.columns]]


def gerar_txt_lista(df):
    """Gera um arquivo de texto com blocos contendo detalhes de cada obra e link do Google Maps."""
    linhas_txt = []
    for _, r in df.iterrows():
        nota = str(r.get('NOTA', r.get('PROTOCOLO', ''))).strip()
        if nota.lower() in ['nan', 'none', '']: nota = '-'
        
        nome = str(r.get('NOME', '')).strip()
        if nome.lower() in ['nan', 'none']: nome = ''
        
        # --- NOVOS CAMPOS: CONTA CONTRATO, INSTALAÇÃO E FASE ---
        conta_contrato = str(r.get('CONTA CONTRATO', '')).strip()
        if conta_contrato.endswith('.0'): conta_contrato = conta_contrato[:-2]
        if conta_contrato.lower() in ['nan', 'none', '']: conta_contrato = '-'
            
        instalacao = str(r.get('INSTALAÇÃO', r.get('INSTALACAO', ''))).strip()
        if instalacao.endswith('.0'): instalacao = instalacao[:-2]
        if instalacao.lower() in ['nan', 'none', '']: instalacao = '-'
            
        fase_raw = str(r.get('FASE', '')).strip().upper()
        if fase_raw == 'MO': fase_txt = '(MO) MONOFÁSICA'
        elif fase_raw == 'B': fase_txt = '(B) BIFÁSICA'
        elif fase_raw == 'TR': fase_txt = '(TR) TRIFÁSICA'
        elif fase_raw in ['NAN', 'NONE', '']: fase_txt = '-'
        else: fase_txt = fase_raw
        
        # --- ENDEREÇO E BAIRRO ---
        endereco = str(r.get('ENDEREÇO', r.get('ENDERECO', ''))).strip()
        if endereco.lower() in ['nan', 'none']: endereco = ''
        
        bairro = str(r.get('BAIRRO', r.get('LOCALIDADE', ''))).strip()
        if bairro.lower() in ['nan', 'none']: bairro = ''
        
        if bairro and endereco:
            endereco_completo = f"{endereco} / {bairro}"
        elif endereco:
            endereco_completo = endereco
        elif bairro:
            endereco_completo = bairro
        else:
            endereco_completo = '-'

        municipio = str(r.get('MUNICIPIO', '')).strip()
        if municipio.lower() in ['nan', 'none']: municipio = '-'
        
        # --- INFORMAÇÕES EXTRAS DA COLUNA AA ---
        info_extra = r.get('INFORMAÇÕES')
        if pd.isna(info_extra): 
            info_extra = r.get('INFORMACOES')
            
        if pd.isna(info_extra) or str(info_extra).strip().lower() in ['nan', 'none', '']:
            info_extra = ''
        else:
            info_extra = str(info_extra).strip()
        
        tipo_nota = str(r.get('TIPO NOTA', '')).strip()
        if tipo_nota.lower() in ['nan', 'none']: tipo_nota = ''

        lat = str(r.get('LATITUDE', '')).strip()
        lon = str(r.get('LONGITUDE', '')).strip()

        # --- MONTAGEM DO BLOCO NO FORMATO EXIGIDO ---
        bloco = []
        bloco.append(f"NOTA CCS: {nota}")
        bloco.append(f"NOME DO CLIENTE: {nome}")
        bloco.append(f"CONTA CONTRATO: {conta_contrato}")
        bloco.append(f"INSTALAÇÃO: {instalacao}")
        bloco.append(f"FASE: {fase_txt}")
        bloco.append(f"ENDEREÇO: {endereco_completo}")
        bloco.append(f"MUNICIPIO: {municipio}")
        if r.get('ORDEM') == 1 or str(r.get('ORDEM')) == '1':
            origem = str(r.get('ORIGEM_MUNICIPIO', '')).strip()
            if origem:
                bloco.append(f"ORIGEM DA EQUIPE: {origem}")
                bloco.append(f"DISTANCIA ATE A PRIMEIRA OBRA: {r.get('DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM', '-')} km")
                bloco.append(f"TEMPO DE DESLOCAMENTO INICIAL: {formatar_tempo_deslocamento(r.get('TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN', 0))}")
                bloco.append(f"METODO: {r.get('TIPO_DESLOCAMENTO', '-')}")
        bloco.append(f"INFORMAÇOES EXTRAS: {info_extra}")
        bloco.append(f"TIPO NOTA: {tipo_nota}")
        
        if lat and lon and lat.lower() != 'nan' and lon.lower() != 'nan':
            bloco.append(f"https://www.google.com.br/maps/place/{lat},{lon}")
        
        linhas_txt.append("\n".join(bloco))
        linhas_txt.append("\n-------------------------------------------------------------------------------------------------------------------------\n")
        
    return "".join(linhas_txt)


def gerar_kml_lista(df_kml, nome_arquivo, colunas_exibir, bases_ativas, funcao_formatadora):
    kml = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>', f'<name>{html.escape(nome_arquivo)}</name>']
    
    # Injetando Estilos de Ícones Padrões e Contorno
    kml.append('<Style id="linha-rota-contorno"><LineStyle><color>ff000000</color><width>8</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')
    kml.append('<Style id="linha-ligacao-rede"><LineStyle><color>8800ffff</color><width>2</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')
    # Marcador de moradia/hospedagem, usando a casinha enviada pelo usuario.
    # Em KML: o PNG acompanha o arquivo na mesma pasta do ZIP.
    # Em KMZ: o PNG fica incorporado, pronto para abrir no Google Earth.
    kml.append(
        '<Style id="icon-casa-origem"><IconStyle><scale>1.15</scale>'
        f'<Icon><href>{ICONE_CASA_ARQUIVO}</href></Icon>'
        '<hotSpot x="0.5" y="0.10" xunits="fraction" yunits="fraction"/>'
        '</IconStyle><LabelStyle><scale>0.9</scale><color>ffffffff</color>'
        '</LabelStyle></Style>'
    )

    kml.append('''<Style id="style-blue-n"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-blue-h"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-blue"><Pair><key>normal</key><styleUrl>#style-blue-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-blue-h</styleUrl></Pair></StyleMap>''')

    kml.append('''<Style id="style-red-n"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-red-h"><IconStyle><scale>1.5</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-blank.png</href></Icon><hotSpot x="32" xunits="pixels" y="64" yunits="insetPixels"/></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-red"><Pair><key>normal</key><styleUrl>#style-red-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-red-h</styleUrl></Pair></StyleMap>''')

    kml.append('''<Style id="style-yellow-n"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png</href></Icon></IconStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>
<Style id="style-yellow-h"><IconStyle><scale>1.5</scale><Icon><href>http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png</href></Icon></IconStyle><LabelStyle><scale>1.0</scale><color>ffffffff</color></LabelStyle></Style>
<StyleMap id="icon-yellow"><Pair><key>normal</key><styleUrl>#style-yellow-n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#style-yellow-h</styleUrl></Pair></StyleMap>''')

    cores_kml = ['ff4b19e6', 'ffd4bc00', 'ffb5513f', 'ff889600', 'ff0098ff', 'ffb0279c', 'ff39dccd', 'ff148000', 'ffeb004b', 'ff1f618d', 'ffd35400', 'ff16a085', 'ff8e44ad', 'ff27ae60', 'ffe67e22']

    # Gerando os estilos das linhas dinamicamente para cada levantador ter uma cor diferente
    for idx, b in enumerate(bases_ativas):
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        b_safe = re.sub(r'[^A-Za-z0-9]', '', str(b))
        style_id = f"rota-centro-{b_safe}"
        cor = cores_kml[idx % len(cores_kml)]
        kml.append(f'<Style id="{style_id}"><LineStyle><color>{cor}</color><width>5</width></LineStyle><LabelStyle><scale>0</scale><color>00ffffff</color></LabelStyle></Style>')

    for b in bases_ativas:
        if pd.isna(b) or b == "NÃO ALOCADO": continue
        pasta = [f'<Folder><name>Equipe: {html.escape(str(b))}</name>']
        df_b = df_kml[df_kml['BASE_ATRIBUIDA'] == b]
        b_safe = re.sub(r'[^A-Za-z0-9]', '', str(b))

        if not df_b.empty:
            primeira = df_b.sort_values('ORDEM').iloc[0]
            lat_ini, lon_ini = primeira.get('LAT_ORIGEM'), primeira.get('LON_ORIGEM')
            if pd.notna(lat_ini) and pd.notna(lon_ini):
                mun_ini = html.escape(str(primeira.get('ORIGEM_MUNICIPIO', '')))
                nome_equipe = html.escape(str(b))
                nota_inicial = html.escape(str(primeira.get('PROTOCOLO', '-')))
                km_iniciais = primeira.get('DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM', 0)
                min_iniciais = primeira.get('TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN', 0)
                desc_casa = (
                    f'<b>CASINHA: MORADIA / HOSPEDAGEM</b><br/>'
                    f'<b>Equipe:</b> {nome_equipe}<br/>'
                    f'<b>Municipio escolhido:</b> {mun_ini}<br/>'
                    f'<b>Primeira obra:</b> {nota_inicial}<br/>'
                    f'<b>Distancia ate a primeira obra:</b> {km_iniciais} km<br/>'
                    f'<b>Tempo de deslocamento inicial:</b> {formatar_tempo_deslocamento(min_iniciais)}<br/>'
                    '<i>Posicao indicativa: sede do municipio selecionado. '
                    'Nao e o endereco exato da casa ou hospedagem.</i>'
                )
                pasta.append(
                    f'<Placemark><name>INICIO DA EQUIPE - {mun_ini} - CASINHA</name>'
                    f'<description><![CDATA[{desc_casa}]]></description>'
                    '<styleUrl>#icon-casa-origem</styleUrl>'
                    f'<Point><coordinates>{float(lon_ini)},{float(lat_ini)},0</coordinates></Point>'
                    '</Placemark>'
                )

        for p in df_b['PERIODO'].unique():
            df_p = df_b[df_b['PERIODO'] == p]
            pasta.append(f'<Folder><name>Rota Principal</name>')
            
            coords_linha = []
            for _, r in df_p.iterrows():
                geom = r.get('ROTA_GEOMETRIA')
                if isinstance(geom, list) and len(geom) > 0:
                    for pt in geom:
                        if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                            coords_linha.append(f"{pt[0]},{pt[1]},0")
                else:
                    lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                    if pd.notna(lat) and pd.notna(lon): coords_linha.append(f"{lon},{lat},0")
            
            if coords_linha:
                str_coords = '\n'.join(coords_linha)
                pasta.append('<Placemark><name>Contorno Rota</name><styleUrl>#linha-rota-contorno</styleUrl><LineString><tessellate>1</tessellate><coordinates>\n' + str_coords + '\n</coordinates></LineString></Placemark>')
                pasta.append(f'<Placemark><name>Traçado Rota</name><styleUrl>#rota-centro-{b_safe}</styleUrl><LineString><tessellate>1</tessellate><coordinates>\n' + str_coords + '\n</coordinates></LineString></Placemark>')

            for _, r in df_p.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                lat, lon = r.get('LATITUDE'), r.get('LONGITUDE')
                if pd.isna(lat) or pd.isna(lon): continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r['_ORIGINAL_ROWS']) > 1
                is_prio = str(r.get('PRIORIDADE')).upper() == 'SIM'
                
                if is_sp:
                    bg_color, text_color = "#FFD700", "#000000"
                    qty = len(r.get('_ORIGINAL_ROWS', []))
                    header_txt = f"🏢 SUPER PONTO ({qty} un.)"
                    icon = "icon-yellow"
                    nome_ponto = f"[{r.get('ORDEM', '')}] 🏢 SUPER PONTO ({qty} un.)"
                elif is_prio:
                    bg_color, text_color = "#d9534f", "#ffffff"
                    header_txt = "🚨 OBRA PRIORITÁRIA"
                    icon = "icon-red"
                    nome_ponto = f"[PRIORIDADE] [{r.get('ORDEM', '')}] Doc: {r.get('PROTOCOLO', '')}"
                else:
                    bg_color, text_color = "#0D256C", "#ffffff"
                    header_txt = "📍 Atendimento Padrão"
                    icon = "icon-blue"
                    nome_ponto = f"[{r.get('ORDEM', '')}] Doc: {r.get('PROTOCOLO', '')}"

                desc = f'''<![CDATA[
                <div style="font-family:sans-serif; width:280px; border-radius:8px; overflow:hidden; box-shadow:0 2px 5px rgba(0,0,0,0.15);">
                    <div style="background:{bg_color}; color:{text_color}; padding:8px 10px; font-size:13px; font-weight:bold;">{header_txt}</div>
                    <div style="padding:10px; background:#fafafa; font-size:12px;">
                        <table style="width:100%; border-collapse:collapse;">
                '''
                
                ordem_txt = f"{r.get('ORDEM', '')}"
                if r.get('DIA_SEMANA'):
                    dia_sm = str(r.get('DIA_SEMANA')).title()
                    ordem_txt += f" ({dia_sm})"

                dist_ant = f"{r.get('DISTANCIA_PONTO_ANTERIOR_KM', 0.0)} KM"
                dist_prox = f"{r.get('DISTANCIA_PROXIMO_PONTO_KM', 0.0)} KM"

                if is_sp:
                    # Identifica cada nota e o respectivo cliente no superponto.
                    entradas_super = []
                    for i, orig in enumerate(r['_ORIGINAL_ROWS'], start=1):
                        protocolo = orig.get('PROTOCOLO', orig.get('NOTA', ''))
                        cliente = orig.get('NOME', orig.get('NOME DO CLIENTE', orig.get('CLIENTE', '')))
                        cliente_txt = str(cliente).strip() if pd.notna(cliente) else ''
                        if cliente_txt.upper() in ('NAN', 'NONE', 'NULL'): cliente_txt = ''
                        entradas_super.append(
                            f"<li style='margin-bottom:5px;'><b>[{i}]</b> "
                            f"{html.escape(str(protocolo))}"
                            + (f"<br><span style='color:#22526b'>Cliente: {html.escape(cliente_txt)}</span>" if cliente_txt else '')
                            + "</li>"
                        )
                    prot_html = "<div style='max-height:160px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>" + "".join(entradas_super) + "</ul></div>"
                else:
                    prot_html = html.escape(str(r.get('PROTOCOLO', r.get('NOTA', ''))))

                desc += f'''
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;">PROTOCOLO:</td><td style="padding:3px 6px; color:#333;">{prot_html}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Ordem:</td><td style="padding:3px 6px; color:#333;">{html.escape(str(ordem_txt))}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Ant.:</td><td style="padding:3px 6px; color:#333;">{dist_ant}</td></tr>
                    <tr><td style="padding:3px 6px; font-weight:bold; color:#555;">Distância Próx.:</td><td style="padding:3px 6px; color:#333;">{dist_prox}</td></tr>
                '''

                for c in colunas_exibir:
                    # O protocolo ja aparece no topo junto a seu cliente nos superpontos.
                    if c.upper() in ['PROTOCOLO', 'NOTA', 'DIA_SEMANA', 'DIA_MES']: continue
                    
                    if c.upper() == 'SUPER_PONTO':
                        val_html = f"SIM ({qty} Obras)" if is_sp else "NÃO"
                        desc += f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;'>{html.escape(c)}:</td><td style='padding:3px 6px; color:#333;'>{val_html}</td></tr>"
                        continue

                    if is_sp and c.upper() not in ['LATITUDE', 'LONGITUDE', 'MUNICIPIO', 'LOCALIDADE', 'ZONA', 'REGIONAL']:
                        vals = [orig.get(c, '') for orig in r.get('_ORIGINAL_ROWS', [])]
                        val_html = "<div style='max-height:95px; overflow-y:auto; border:1px solid #ccc; padding:6px; background:#fff; border-radius:4px;'><ul style='margin:0; padding-left:0px; list-style-type:none; font-size:11px; color:#333;'>" + "".join([f"<li style='margin-bottom:3px;'><b>[{i+1}]</b> {funcao_formatadora(c, v)}</li>" for i, v in enumerate(vals)]) + "</ul></div>"
                    else:
                        val_html = funcao_formatadora(c, r.get(c, ''))
                        
                    desc += f"<tr><td style='padding:3px 6px; font-weight:bold; color:#555; vertical-align:top; width:35%;'>{html.escape(c)}:</td><td style='padding:3px 6px; color:#333;'>{val_html}</td></tr>"

                if r.get('ORDEM') == 1:
                    desc += (
                        f"<tr><td style='padding:3px 6px; font-weight:bold;'>Origem da equipe:</td>"
                        f"<td>{html.escape(str(r.get('ORIGEM_MUNICIPIO', '-')))}</td></tr>"
                        f"<tr><td style='padding:3px 6px; font-weight:bold;'>Tempo inicial:</td>"
                        f"<td>{formatar_tempo_deslocamento(r.get('TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN', 0))}</td></tr>"
                        f"<tr><td style='padding:3px 6px; font-weight:bold;'>Cálculo:</td>"
                        f"<td>{html.escape(str(r.get('TIPO_DESLOCAMENTO', '-')))}</td></tr>"
                    )
                desc += "</table></div></div>]]>"

                pasta.append(f'<Placemark><name>{html.escape(nome_ponto)}</name><styleUrl>#{icon}</styleUrl><description>{desc}</description><Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>')
                
            pasta.append('</Folder>')
        pasta.append('</Folder>')
        kml.extend(pasta)
        
    kml.extend(['</Document>', '</kml>'])
    return "\n".join(kml)


def gerar_gpx_simples(df_kml, nome_rota):
    gpx = ['<?xml version="1.0" encoding="UTF-8"?>', '<gpx version="1.1" creator="Roteirizador NIP" xmlns="http://www.topografix.com/GPX/1/1">', f'  <metadata><name>{html.escape(str(nome_rota))}</name></metadata>']
    if not df_kml.empty and {'LAT_ORIGEM', 'LON_ORIGEM', 'ORIGEM_MUNICIPIO'}.issubset(df_kml.columns):
        for b, db in df_kml.groupby('BASE_ATRIBUIDA'):
            primeira = db.sort_values('ORDEM').iloc[0]
            lat_i, lon_i = primeira.get('LAT_ORIGEM'), primeira.get('LON_ORIGEM')
            if pd.notna(lat_i) and pd.notna(lon_i):
                nome_origem = html.escape(str(primeira.get('ORIGEM_MUNICIPIO', 'Origem')))
                gpx.append(f'  <wpt lat="{lat_i}" lon="{lon_i}"><name>ORIGEM - {nome_origem}</name></wpt>')
    for _, row in df_kml.iterrows():
        if row.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
        lat, lon = row.get('LATITUDE'), row.get('LONGITUDE')
        if pd.notna(lat) and pd.notna(lon):
            gpx.append(f'  <wpt lat="{lat}" lon="{lon}"><name>{html.escape(str(row.get("PROTOCOLO", "Ponto")))}</name></wpt>')
    if 'ROTA_GEOMETRIA' in df_kml.columns:
        gpx.append(f'  <trk><name>Traçado - {html.escape(str(nome_rota))}</name><trkseg>')
        for _, row in df_kml.iterrows():
            geom = row.get('ROTA_GEOMETRIA')
            if isinstance(geom, list):
                for lon, lat in geom: gpx.append(f'      <trkpt lat="{lat}" lon="{lon}"></trkpt>')
        gpx.append('    </trkseg></trk>')
    gpx.append('</gpx>')
    return "\n".join(gpx)


def identificar_icone_folium(row, colunas_disponiveis):
    if 'TIPO NOTA' in colunas_disponiveis:
        t = str(row.get('TIPO NOTA', '')).upper()
        if t in ['UNR', 'ASC']: return 'bolt'
        elif t in ['MGD', 'MTP']: return 'industry'
        elif t == 'DIF': return 'exclamation-triangle'
    return 'map-marker'


injetar_logo()

def formatar_valor_coluna(c, v):
    if pd.isna(v) or v in ['', '-']: return '-'
    # Identificadores de instalacao/contrato nao sao quantidades ou medidas.
    if c.upper() in ['INSTALAÇÃO', 'INSTALACAO', 'CONTA CONTRATO', 'ID SISCO']:
        if isinstance(v, (float, np.floating)) and float(v).is_integer():
            return html.escape(str(int(v)))
        return html.escape(str(v).strip())
    try:
        if 'POSTE' in c.upper(): return str(int(float(v)))
        
        vf = float(v)
        if c.upper() in ['DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM']: 
            return f"{vf:.2f} KM"
        elif 'DISTANCIA' in c.upper(): 
            return f"{vf:.2f} Metros"
            
        return formata_campo_html(v)
    except:
        if isinstance(v, (datetime, pd.Timestamp)): return formata_campo_html(v.strftime('%d/%m/%Y'))
        return formata_campo_html(str(v))

def render_sidebar_card(total_obras, qtd_equipes_ativas):
    return f"""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #dee2e6; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #0D256C; font-size: 16px; border-bottom: 2px solid #55B929; padding-bottom: 5px;">📊 Resumo da Capacidade</h4>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Equipes Ativas:</b> <span style="color: #0D256C; font-weight: bold;">{qtd_equipes_ativas}</span></p>
        <p style="margin-bottom: 5px; font-size: 14px;"><b>Modo:</b> <span style="color: #d9534f; font-weight: bold;">Lista Contínua</span></p>
        <hr style="margin: 10px 0; border: 0; border-top: 1px solid #ddd;">
        <p style="margin-bottom: 0; font-size: 15px; text-align: center;"><b>Obras Validadas:</b> <br><span style="font-size: 24px; color: #0D256C; font-weight: 900;">{total_obras}</span></p>
    </div>
    """

def render_metric_card(title, value, icon, border_color, bg_color):
    return f"""
    <div style="background-color: #ffffff; border-left: 5px solid {border_color}; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); display: flex; align-items: center; margin-bottom: 10px;">
        <div style="background-color: {bg_color}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; font-size: 20px; margin-right: 15px;">
            {icon}
        </div>
        <div>
            <p style="margin: 0; font-size: 11px; color: #666; text-transform: uppercase; font-weight: bold;">{title}</p>
            <p style="margin: 0; font-size: 22px; color: #333; font-weight: bold;">{value}</p>
        </div>
    </div>
    """



def ordenar_rota_rapida(obras, lat_inicial, lon_inicial, reversa=False):
    """Ordenação gulosa vetorizada, sem chamadas externas e sem risco de travar no solver."""
    if not obras:
        return []

    coords = np.array([[float(o['LATITUDE']), float(o['LONGITUDE'])] for o in obras], dtype=float)
    restantes = np.ones(len(obras), dtype=bool)
    ordem = []

    lat_atual, lon_atual = float(lat_inicial), float(lon_inicial)

    if reversa and len(obras) > 1:
        # Começa pelo ponto mais distante do centro e depois segue sempre ao mais próximo.
        d0 = (coords[:, 0] - lat_atual) ** 2 + ((coords[:, 1] - lon_atual) * np.cos(np.radians(lat_atual))) ** 2
        idx = int(np.argmax(d0))
        ordem.append(obras[idx])
        restantes[idx] = False
        lat_atual, lon_atual = coords[idx]

    while restantes.any():
        idxs = np.flatnonzero(restantes)
        cand = coords[idxs]
        cos_lat = np.cos(np.radians(lat_atual))
        dist2 = (cand[:, 0] - lat_atual) ** 2 + ((cand[:, 1] - lon_atual) * cos_lat) ** 2
        idx = int(idxs[int(np.argmin(dist2))])
        ordem.append(obras[idx])
        restantes[idx] = False
        lat_atual, lon_atual = coords[idx]

    return ordem


class FalhaRoteamentoOSRM(Exception):
    """O servico viario nao conseguiu produzir um trajeto valido entre os pontos."""


def obter_rota_ruas_segura(lat1, lon1, lat2, lon2, url_osrm_base, velocidade_media_kmh, timeout_s=15):
    """Obtem geometria [longitude,latitude] e distancia real do OSRM.

    Nunca devolve linha reta em lugar de rota viaria: uma falha interrompe o
    motor para evitar que o mapa e o KML informem erroneamente um arruamento.
    Retorno: (geometria_lon_lat, duracao_segundos, distancia_km).
    """
    try:
        lat1, lon1 = float(lat1), float(lon1)
        lat2, lon2 = float(lat2), float(lon2)
    except (ValueError, TypeError) as exc:
        raise FalhaRoteamentoOSRM('Coordenadas de origem ou destino invalidas.') from exc

    distancia_direta_km = float(haversine_scalar(lat1, lon1, lat2, lon2))
    if distancia_direta_km < 0.003:
        # Pontos coincidentes nao necessitam de consulta; nenhum trecho ficticio.
        return [[lon1, lat1], [lon2, lat2]], 0.0, 0.0

    base = str(url_osrm_base or '').strip().rstrip('/')
    # O endpoint publico do OSRM precisa ser acessado por HTTPS no Streamlit Cloud.
    if base.lower().startswith('http://router.project-osrm.org'):
        base = 'https://' + base[7:]
    if not base.startswith(('https://', 'http://')):
        raise FalhaRoteamentoOSRM('Endpoint OSRM invalido. Utilize uma URL HTTPS valida.')

    url = f'{base}/route/v1/driving/{lon1:.7f},{lat1:.7f};{lon2:.7f},{lat2:.7f}'
    params = {'overview': 'full', 'geometries': 'geojson', 'steps': 'false'}
    last_error = None
    for tentativa in range(2):
        try:
            # Respeitar intervalo de 1 segundo no servidor publico de demonstracao.
            if 'router.project-osrm.org' in base.lower():
                ultima = st.session_state.get('nip_osrm_ultima_consulta', 0.0)
                pausa = 1.10 - (time.monotonic() - ultima)
                if pausa > 0:
                    time.sleep(pausa)
                st.session_state['nip_osrm_ultima_consulta'] = time.monotonic()
            resp = requests.get(
                url, params=params, timeout=(5, max(8, timeout_s)),
                headers={'User-Agent': 'NIP-Roteirizacao/1.0', 'Accept': 'application/json'}
            )
            resp.raise_for_status()
            data = resp.json()
            codigo = data.get('code')
            if codigo != 'Ok':
                descricao = data.get('message') or codigo or 'resposta sem codigo'
                raise FalhaRoteamentoOSRM(f'OSRM retornou {descricao}. Verifique a rede viaria dos pontos.')
            rotas = data.get('routes') or []
            if not rotas:
                raise FalhaRoteamentoOSRM('O OSRM nao encontrou uma rota rodoviaria entre os pontos.')
            rota = rotas[0]
            geom = (rota.get('geometry') or {}).get('coordinates') or []
            if len(geom) < 2 or not all(
                isinstance(p, (list, tuple)) and len(p) >= 2
                and all(math.isfinite(float(v)) for v in p[:2]) for p in geom
            ):
                raise FalhaRoteamentoOSRM('O OSRM nao retornou uma geometria viaria valida.')
            metros = float(rota.get('distance', -1))
            if not math.isfinite(metros) or metros < 0:
                raise FalhaRoteamentoOSRM('O OSRM nao informou uma distancia de rota valida.')
            duracao = float(rota.get('duration', 0) or 0)
            return [[float(p[0]), float(p[1])] for p in geom], duracao, metros / 1000.0
        except FalhaRoteamentoOSRM:
            raise
        except (requests.exceptions.RequestException, ValueError, TypeError, KeyError) as exc:
            last_error = exc
            # Repetir apenas falhas temporarias de rede/HTTP; nao falsificar trajetos.
            if tentativa == 0:
                time.sleep(1.0)
                continue
    raise FalhaRoteamentoOSRM(
        f'Falha de comunicacao com {base}: {last_error}. '
        'Confira o endpoint, o acesso HTTPS e a disponibilidade da rede do OSRM.'
    ) from last_error


def tentar_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def limpar_roteirizador():
    st.session_state.update({'roteamento_concluido_lista': False, 'vrp_status_lista': "IDLE", 'vrp_state_lista': {}, 'df_routed_lista': pd.DataFrame(), 'colunas_exibir_lista': [], 'colunas_originais_lista': []})
    for k in ['bytes_zip_xl_lista', 'bytes_zip_kml_lista', 'bytes_zip_gpx_lista', 'bytes_zip_txt_lista', 'bytes_zip_maps_lista', 'bytes_zip_offline_lista', 'maps_primeiras_lista', 'start_time_run_lista', 'start_time_pkg_lista', 'df_unallocated_lista', 'df_correcao_lista']: st.session_state.pop(k, None)
    ler_planilha_cached.clear()
    tentar_rerun()

if "roteamento_concluido_lista" not in st.session_state: st.session_state.roteamento_concluido_lista = False
if "vrp_status_lista" not in st.session_state: st.session_state.vrp_status_lista = "IDLE"

status_exec = st.session_state.vrp_status_lista
is_done = st.session_state.roteamento_concluido_lista
is_locked = status_exec != "IDLE" or is_done

# Expande apenas o conteudo principal; a barra lateral permanece inalterada.
st.markdown("""<style>
[data-testid="stMainBlockContainer"], section.main .block-container {
    max-width: 1480px !important;
    width: min(96%, 1480px) !important;
}
@media (max-width: 850px) {
    [data-testid="stMainBlockContainer"], section.main .block-container {
        width: 100% !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
}
</style>""", unsafe_allow_html=True)
st.markdown("<h1 class='brand-title'>📜 Roteirização</h1>", unsafe_allow_html=True)
st.info("💡 Gera uma lista de execução contínua com base na atribuição já existente na planilha de obras.")

with st.sidebar:
    st.markdown("### ⚙️ Configurações Logísticas")
    with st.expander("Parâmetros de Rota", expanded=True):
        st.success("📦 **Modo Contínuo:** Todas as obras serão alocadas numa lista única.")
        trava_global = st.number_input("Trava Total de Obras", min_value=0, value=0, step=50, disabled=is_locked)
        data_ini = st.date_input("📅 Data Base (Início):", value=datetime.today(), disabled=is_locked)
        velocidade_media = st.number_input("Velocidade média estimada (km/h):", min_value=5, max_value=110, value=30, step=5, disabled=is_locked,
                                           help="Usada apenas quando o traçado viário OSRM está desligado. Estimativa com fator viário de 1,30.")
        
        st.markdown("---")
        obras_por_dia_est = st.number_input("Meta de Obras/Dia (Cálculo de Postes):", min_value=1, value=4, disabled=is_locked)
        
        sentido_rota = st.radio("Sentido do Roteamento:", ["📍 Lógica Padrão", "🎯 Varredura Reversa"], index=0, disabled=is_locked)
        raio_sp = st.slider("Raio Super Ponto (m):", 10, 500, 50, 10, disabled=is_locked)
        
    with st.expander("📡 Conexão de Rede", expanded=False):
        url_osrm = st.text_input("Endpoint OSRM:", value="https://router.project-osrm.org", disabled=is_locked)
        usa_osrm = st.checkbox("🛣️ Traçado de Ruas Real (Lento)", value=False, disabled=is_locked, help="Ative apenas quando precisar do traçado viário real. Desativado, o motor usa distância geográfica e fica muito mais rápido.")
        if not is_locked and st.button("🔎 Testar conexão OSRM", use_container_width=True, help="Testa uma rota curta em São Luís/MA e confirma se o servidor devolve geometria de ruas."):
            try:
                caminho, segundos, km = obter_rota_ruas_segura(-2.532, -44.302, -2.529, -44.295, url_osrm, 30.0, timeout_s=15)
                st.success(f"Conexão OSRM OK: {len(caminho)} pontos viários; {km:.2f} km.")
            except FalhaRoteamentoOSRM as exc:
                st.error(f"Falha no teste do OSRM: {exc}")
        if usa_osrm:
            st.caption("Rotas por ruas exigem resposta válida do servidor. Se o OSRM falhar, a exportação será interrompida, sem substituir o percurso por uma reta.")

    sb_html = st.empty()

    if is_done and not st.session_state.df_routed_lista.empty:
        d_fmt = datetime.now().strftime("%d.%m.%Y")
        st.download_button("🌐 Baixar Planilhas (ZIP)", data=st.session_state.get('bytes_zip_xl_lista', b"vazio"), file_name=f"ListaContinua_Planilhas - {d_fmt}.zip", use_container_width=True)
        st.download_button("📝 Baixar Relatórios (TXT)", data=st.session_state.get('bytes_zip_txt_lista', b"vazio"), file_name=f"ListaContinua_TXT - {d_fmt}.zip", use_container_width=True)
        st.download_button("🏠 Baixar Mapas (KML e KMZ)", data=st.session_state.get('bytes_zip_kml_lista', b"vazio"), file_name=f"ListaContinua_Mapas - {d_fmt}.zip", use_container_width=True)
        st.download_button("🛰️ Baixar GPS (GPX)", data=st.session_state.get('bytes_zip_gpx_lista', b"vazio"), file_name=f"ListaContinua_GPS - {d_fmt}.zip", use_container_width=True)
        st.download_button("📍 Baixar Rotas Google Maps (ZIP)",
                           data=st.session_state.get('bytes_zip_maps_lista', b""),
                           file_name=f"ListaContinua_GoogleMaps - {d_fmt}.zip",
                           use_container_width=True,
                           help="Links de navegacao, HTML editavel por equipe e CSV para Google My Maps.")
        st.download_button("📴 Baixar Planejamento Offline (ZIP)",
                           data=st.session_state.get('bytes_zip_offline_lista', b""),
                           file_name=f"ListaContinua_Offline - {d_fmt}.zip",
                           use_container_width=True,
                           help="HTML offline por equipe, GPX para Organic Maps/OsmAnd, links e pacote PWA HTTPS para iPhone/Android.")
        st.caption("🏠 A casinha indica a sede do municipio de moradia/hospedagem. "
                   "Abra o KMZ no Google Earth para visualizar o icone sem arquivos separados. "
                   "Nos KMLs, mantenha CASINHA_ORIGEM.png na mesma pasta.")
        if st.button("🧹 Nova Roteirização", type="primary", use_container_width=True): limpar_roteirizador()

if is_done and not st.session_state.df_routed_lista.empty:
    st.markdown("## 🎯 Resultado do Planejamento")
    
    df_c = st.session_state.get('df_correcao_lista', pd.DataFrame())
    if not df_c.empty:
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_c)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa Técnica:</b> Estas obras apresentaram coordenadas em branco, zeradas ou invertidas. Elas foram isoladas para não corromper o mapa.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.df_routed_lista['DISTANCIA_PROXIMO_PONTO_KM'] = st.session_state.df_routed_lista.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)
    dfr = st.session_state.df_routed_lista.copy()
    dfr_t = dfr[~dfr['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])]
    
    tr = len(dfr_t)
    te = dfr['BASE_ATRIBUIDA'].nunique()
    tk = f"{dfr['DISTANCIA_PONTO_ANTERIOR_KM'].sum():.1f} km"
    
    tsp = sum(1 for _, r in dfr_t.iterrows() if isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(render_metric_card("Obras Roteirizadas", tr, "🎯", "#0D256C", "rgba(13,37,108,0.12)"), unsafe_allow_html=True)
    c2.markdown(render_metric_card("Equipes Alocadas", te, "👥", "#8b5cf6", "rgba(139,92,246,0.15)"), unsafe_allow_html=True)
    c3.markdown(render_metric_card("Super Pontos", str(tsp), "🏢", "#FF9800", "rgba(255,152,0,0.15)"), unsafe_allow_html=True)
    c4.markdown(render_metric_card("KM Total Previsto", tk, "🛣️", "#55B929", "rgba(85,185,41,0.15)"), unsafe_allow_html=True)

    st.markdown("### 🚐 Deslocamento inicial das equipes")
    primeiras = dfr[dfr['ORDEM'] == 1].copy()
    if not primeiras.empty:
        detalhes = primeiras[[
            'BASE_ATRIBUIDA', 'ORIGEM_MUNICIPIO', 'PROTOCOLO', 'PRIORIDADE',
            'DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM', 'TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN',
            'TIPO_DESLOCAMENTO'
        ]].rename(columns={
            'BASE_ATRIBUIDA': 'EQUIPE (COLUNA P)', 'ORIGEM_MUNICIPIO': 'PARTIDA',
            'PROTOCOLO': 'PRIMEIRA NOTA', 'PRIORIDADE': 'PRIORITARIA',
            'DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM': 'DISTANCIA INICIAL (KM)',
            'TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN': 'TEMPO INICIAL (MIN)',
            'TIPO_DESLOCAMENTO': 'CALCULO',
        })
        st.dataframe(detalhes, use_container_width=True, hide_index=True)
        tempo_inicial = float(primeiras['TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN'].sum())
        st.caption(f"Tempo de saída somado das equipes: {formatar_tempo_deslocamento(tempo_inicial)}. "
                   "Cada equipe inicia de seu município definido, não da média das obras.")

    html_msg = f"""
    <div style='background-color: #d4edda; border-left: 8px solid #28a745; padding: 20px; border-radius: 8px; margin-bottom: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
        <h2 style='color: #155724; margin-top: 0; font-weight: 900; letter-spacing: 1px; font-size: 24px;'>🎉 100% DAS OBRAS ROTEIRIZADAS!</h2>
        <p style='color: #155724; font-size: 16px; margin-bottom: 10px;'>Todas as <b>{tr} tarefas válidas</b> presentes na sua planilha foram organizadas e processadas com sucesso!</p>
        <hr style='border-top: 1px solid #c3e6cb; margin: 15px 0;'>
        <p style='color: #155724; font-size: 14px; margin-bottom: 0;'>
            <b>💡 Resumo Técnico:</b> O módulo de <b>Lista Contínua</b> opera sem limites de jornada ou cota de dias. O motor respeitou a equipe definida na planilha e a origem municipal selecionada, organizando os pontos por proximidade geográfica (do ponto A ao Z) para <b>todas</b> as obras de cada levantador, sem deixar absolutamente nada para trás.
        </p>
    </div>
    """
    st.markdown(html_msg, unsafe_allow_html=True)

    st.markdown("### 🗺️ Mapa Operacional")
    mapa = folium.Map(location=[dfr['LATITUDE'].mean(), dfr['LONGITUDE'].mean()], zoom_start=8) if not dfr.empty else folium.Map(location=[-5.2, -45.0], zoom_start=7)
    co_f = ['#e6194b', '#00bcd4', '#3f51b5', '#009688', '#9c27b0', '#cddc39', '#e91e63', '#ffeb3b', '#795548', '#FF9800']
    
    m_clust = MarkerCluster(name="📍 Obras").add_to(mapa)
    for bn in dfr['BASE_ATRIBUIDA'].unique().tolist():
        cr = co_f[list(dfr['BASE_ATRIBUIDA'].unique()).index(bn) % len(co_f)]
        db = dfr[dfr['BASE_ATRIBUIDA'] == bn]
        fg = folium.FeatureGroup(name=f"Rota: {bn}", show=False)
        
        if not db.empty:
            primeira = db.sort_values('ORDEM').iloc[0]
            if pd.notna(primeira.get('LAT_ORIGEM')) and pd.notna(primeira.get('LON_ORIGEM')):
                origem_label = html.escape(str(primeira.get('ORIGEM_MUNICIPIO', '-')))
                folium.Marker(
                    [float(primeira['LAT_ORIGEM']), float(primeira['LON_ORIGEM'])],
                    tooltip=f"🏠 Moradia / hospedagem - {origem_label} ({bn})",
                    popup=folium.Popup(
                        f"<b>🏠 BASE DA EQUIPE (MORADIA / HOSPEDAGEM):</b> {origem_label}"
                        f"<br><b>Equipe:</b> {html.escape(str(bn))}"
                        "<br><i>Marcador aproximado: sede municipal, nao endereco exato.</i>", max_width=340),
                    icon=folium.CustomIcon(
                        icon_image='data:image/png;base64,' + ICONE_CASA_BASE64,
                        icon_size=(42, 42), icon_anchor=(21, 38)
                    )
                ).add_to(m_clust)
        pts = [p for _, r in db.iterrows() for p in ([[l, L] for L, l in r['ROTA_GEOMETRIA']] if isinstance(r.get('ROTA_GEOMETRIA'), list) else [])]
        folium.PolyLine(pts, color='black', weight=7, opacity=0.9).add_to(fg)
        folium.PolyLine(pts, color=cr, weight=3, opacity=1.0).add_to(fg)
        
        for r in db.to_dict('records'):
            if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
            c_i = 'red' if str(r.get('PRIORIDADE')) == 'Sim' else 'blue'
            ic = identificar_icone_folium(r, dfr.columns)
            er = "".join([f"<tr><td><b>{html.escape(c)}</b></td><td>{formatar_valor_coluna(c, r.get(c, ''))}</td></tr>" for c in st.session_state.colunas_exibir_lista if c.upper() not in ['NOME_DIA','DIA_MES','SEMANA','BASE_ATRIBUIDA']])
            pop_html = f'<div style="width:250px;"><b>Equipe:</b> {html.escape(str(r.get("BASE_ATRIBUIDA")))}<br><b>Ordem:</b> {r.get("ORDEM")}<br><table border="1" style="width:100%;font-size:11px;">{er}</table></div>'
            folium.Marker([r['LATITUDE'], r['LONGITUDE']], icon=folium.Icon(color=c_i, icon=ic), popup=folium.Popup(pop_html, max_width=300)).add_to(m_clust)
        fg.add_to(mapa)
    folium.LayerControl().add_to(mapa); st_folium(mapa, use_container_width=True, height=550)

    st.markdown("### 📍 Google Maps - começar o levantamento")
    links_inicio_maps = st.session_state.get('maps_primeiras_lista', {})
    if links_inicio_maps:
        with st.expander("🚗 Abrir a primeira obra de cada equipe", expanded=False):
            st.caption("A origem é a sede do município escolhido. O Google Maps calculará o caminho rodoviário quando aberto.")
            for equipe_m, dados_m in sorted(links_inicio_maps.items()):
                url_m = html.escape(dados_m['url'], quote=True)
                equipe_nome = html.escape(str(equipe_m))
                nota_nome = html.escape(str(dados_m['nota']))
                st.markdown(
                    f'<a href="{url_m}" target="_blank" rel="noopener noreferrer">🚗 {equipe_nome} — primeira obra {nota_nome}</a>',
                    unsafe_allow_html=True
                )
    st.markdown("### 📊 Dados Tabulares")
    dados_tabela = list(dict.fromkeys([
        'BASE_ATRIBUIDA', 'ORIGEM_MUNICIPIO', 'ORDEM', 'PRIORIDADE',
        'DISTANCIA_PONTO_ANTERIOR_KM', 'TEMPO_VIAGEM_MINUTOS',
        'TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN', 'TIPO_DESLOCAMENTO'
    ] + CAMPOS_VISIVEIS_NIP))
    st.dataframe(st.session_state.df_routed_lista.reindex(columns=dados_tabela), use_container_width=True, hide_index=True)

elif status_exec == "IDLE":
    st.markdown("### 📁 Demandas (Obras e Equipes)")
    st.info("💡 Apenas faça o upload da sua planilha de Obras. O sistema identificará o responsável pela coluna P (LEVANTADOR).")
    task_files = st.file_uploader("Suba a planilha com a Lista Contínua", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    if not task_files: st.stop()
    
    # Uma nova planilha limpa configuracoes de origem de planilhas anteriores.
    assinatura_arquivos = hashlib.sha256(
        b'|'.join(hashlib.sha256(f.getvalue()).digest() for f in task_files)
    ).hexdigest()
    if st.session_state.get('assinatura_planilha_origens_lista') != assinatura_arquivos:
        st.session_state['assinatura_planilha_origens_lista'] = assinatura_arquivos
        st.session_state['origem_por_equipe_lista'] = {}
        st.session_state.pop('equipe_origem_selecionada_lista', None)
        st.session_state.pop('municipio_origem_selecionado_lista', None)

    dfs = []
    qtd_coords_corrigidas = 0
    for f in task_files:
        dft = ler_planilha_cached(f.getvalue()) if not f.name.endswith('.csv') else pd.read_csv(f)
        dft.columns = normalize_cols(dft.columns)
        dft = adaptar_planilha_nip(dft)
        qtd_coords_corrigidas += dft.attrs.get('coordenadas_invertidas_corrigidas', 0)
        if not dfs: st.session_state.colunas_originais_lista = CAMPOS_VISIVEIS_NIP.copy()
        for cc in ['NOTA', 'PROTOCOLO', 'OS']:
            if cc in dft.columns: dft['PROTOCOLO'] = dft[cc]; break
        dfs.append(dft)
    
    if qtd_coords_corrigidas:
        st.info(f"📍 {qtd_coords_corrigidas} obra(s) com LATITUDE/LONGITUDE invertidas na planilha tiveram as coordenadas corrigidas automaticamente para a roteirização e os arquivos exportados.")
    df_tasks = pd.concat(dfs, ignore_index=True)
    if 'PROTOCOLO' in df_tasks.columns:
        df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].astype(str).str.split(r'\s*\|\s*')
        df_tasks = df_tasks.explode('PROTOCOLO').reset_index(drop=True); df_tasks['PROTOCOLO'] = df_tasks['PROTOCOLO'].str.strip()

    agent_col = None
    for col in ['BASE_ATRIBUIDA', 'LEVANTADOR_RESPONSAVEL', 'LEVANTADOR', 'FISCAL', 'EQUIPE', 'NOME_FISCAL']:
        if col in df_tasks.columns:
            agent_col = col; break
            
    if not agent_col:
        st.error("❌ A planilha de Obras não possui uma coluna de responsável. Insira uma coluna chamada 'LEVANTADOR' ou 'FISCAL'.")
        st.stop()
        
    if agent_col != 'BASE_ATRIBUIDA': df_tasks = df_tasks.rename(columns={agent_col: 'BASE_ATRIBUIDA'})
    df_tasks['BASE_ATRIBUIDA'] = df_tasks['BASE_ATRIBUIDA'].astype(str).str.strip().str.upper()
    df_tasks = df_tasks[~df_tasks['BASE_ATRIBUIDA'].isin(['NAN', 'NONE', ''])]

    qtd_eq = df_tasks['BASE_ATRIBUIDA'].nunique()
    sb_html.markdown(render_sidebar_card("Ilimitada", qtd_eq), unsafe_allow_html=True)

    st.markdown("---")
    falta = [c for c in ['LATITUDE', 'LONGITUDE', 'PROTOCOLO'] if c not in df_tasks.columns]
    if falta: st.error(f"🚨 Faltam colunas vitais: {', '.join(falta)}."); st.stop()
    
    c1_f, c2_f = st.columns([1, 1])
    with c1_f:
        cs = 'STATUS DA FISCALIZACAO' if 'STATUS DA FISCALIZACAO' in df_tasks.columns else 'STATUS DA FISCALIZAÇÃO'
        if cs in df_tasks.columns:
            df_tasks[cs] = df_tasks[cs].astype(str).str.strip().str.upper()
            opts_s = sorted([str(x) for x in df_tasks[cs].unique() if str(x) != 'NAN'])
            sel_s = st.multiselect("1. Status Roteirizáveis:", options=opts_s, default=[s for s in opts_s if s in ['APTO PARA CAMPO', 'EM CAMPO']])
            if not sel_s: st.stop()
            df_tasks = df_tasks[df_tasks[cs].isin(sel_s)].copy()
            
    with c2_f:
        if 'PRIORIDADE' in df_tasks.columns:
            df_tasks['PRIORIDADE'] = df_tasks['PRIORIDADE'].apply(
                lambda x: 'Sim' if eh_prioridade(x) else 'Não'
            )
            st.success("🎯 **Prioridade Ativada:** As obras marcadas na coluna 'PRIORIDADE' serão destacadas no roteiro e no KML.")
        elif 'TIPO NOTA' in df_tasks.columns:
            df_tasks['TIPO NOTA'] = df_tasks['TIPO NOTA'].astype(str).str.strip().str.upper()
            opts_n = sorted([str(x) for x in df_tasks['TIPO NOTA'].unique() if str(x) != 'NAN'])
            sel_p = st.multiselect("🚨 2. Obras de Alta Prioridade:", options=opts_n, default=[n for n in opts_n if n in ['ASC', 'CCF', 'DIF', 'MGD', 'MTP', 'SID']])
            df_tasks['PRIORIDADE'] = df_tasks['TIPO NOTA'].apply(lambda x: 'Sim' if str(x) in sel_p else 'Não')
        else: 
            df_tasks['PRIORIDADE'] = 'Não'

    st.markdown("#### 🌍 Limpeza Geográfica")
    if st.button("⏹️ Abortar Roteamento", use_container_width=True): limpar_roteirizador(); st.stop()
    
    df_rej = pd.DataFrame(); df_tasks['MOTIVO_REJEICAO'] = ''
    
    df_tasks['LAT_NUM'] = pd.to_numeric(df_tasks['LATITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    df_tasks['LON_NUM'] = pd.to_numeric(df_tasks['LONGITUDE'].astype(str).replace(',', '.', regex=True), errors='coerce')
    m_na, m_0 = df_tasks['LAT_NUM'].isna() | df_tasks['LON_NUM'].isna(), (df_tasks['LAT_NUM'] == 0.0) | (df_tasks['LON_NUM'] == 0.0)
    df_tasks.loc[m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Inválida'
    df_tasks.loc[m_0 & ~m_na, 'MOTIVO_REJEICAO'] = 'Coordenada Zerada'
    m_pos = (df_tasks['LAT_NUM'] > 0) | (df_tasks['LON_NUM'] > 0)
    df_tasks.loc[m_pos & ~m_na & ~m_0, 'MOTIVO_REJEICAO'] = 'Coordenada Positiva'
    m_inv = abs(df_tasks['LAT_NUM']) > abs(df_tasks['LON_NUM'])
    df_tasks.loc[m_inv & ~m_na & ~m_0 & ~m_pos, 'MOTIVO_REJEICAO'] = 'Coordenada Invertida'
    
    mc = m_na | m_0 | m_pos | m_inv
    if mc.sum() > 0: df_rej = pd.concat([df_rej, df_tasks[mc].copy()], ignore_index=True); df_tasks = df_tasks[~mc].copy()
    df_tasks['LATITUDE'], df_tasks['LONGITUDE'] = df_tasks['LAT_NUM'], df_tasks['LON_NUM']; df_tasks.drop(columns=['LAT_NUM', 'LON_NUM'], inplace=True)
    
    st.session_state.df_correcao_lista = df_rej
    if not df_rej.empty: 
        st.markdown(f"""
        <div style='background-color: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-top: 0; margin-bottom: 10px;'>⚠️ {len(df_rej)} Obras Retidas para Correção</h4>
            <p style='color: #856404; font-size: 14px; margin-bottom: 0;'>
                <b>Justificativa:</b> Apresentaram coordenadas zeradas, invertidas ou em branco. Foram removidas da roteirização.
            </p>
        </div>
        """, unsafe_allow_html=True)

    if df_tasks.empty: st.error("🚨 Nenhuma obra válida restou."); st.stop()

    dfs_fundidos = []
    for base in df_tasks['BASE_ATRIBUIDA'].unique():
        df_base = df_tasks[df_tasks['BASE_ATRIBUIDA'] == base].copy()
        df_base_f, _ = fundir_super_pontos(df_base, raio_metros=raio_sp, agrupar_por_levantador=True)
        dfs_fundidos.append(df_base_f)
    df_tasks = pd.concat(dfs_fundidos, ignore_index=True)

    if trava_global > 0: df_tasks = df_tasks.head(trava_global)
    df_tasks = df_tasks.sort_values(by=['PRIORIDADE', 'LATITUDE', 'LONGITUDE'], ascending=[False, True, True])

    df_ta = df_tasks.copy()

    st.markdown("### 🚐 Município de partida de cada equipe")
    st.caption("1. Selecione a equipe lida da coluna P (LEVANTADOR). "
               "2. Escolha entre 217 municípios do Maranhão (MA) e 224 do Piauí (PI), total de 441. "
               "A escolha é salva automaticamente para esta planilha.")
    try:
        municipios_ma_pi = carregar_municipios_ma_pi()
    except RuntimeError as exc:
        st.error(str(exc))
        st.stop()

    equipes_disponiveis = sorted(df_ta['BASE_ATRIBUIDA'].dropna().unique().tolist())
    origens_municipais = st.session_state.setdefault('origem_por_equipe_lista', {})
    # Migrar antigas escolhas salvas sem UF (eram sempre do Maranhao).
    municipios_antigos_ma = {
        normalizar_texto_municipio(m.rsplit(' / ', 1)[0]): m
        for m in municipios_ma_pi if m.endswith(' / MA')
    }
    origens_convertidas = {}
    for equipe, municipio_antigo in origens_municipais.items():
        if equipe not in equipes_disponiveis:
            continue
        if municipio_antigo in municipios_ma_pi:
            origens_convertidas[equipe] = municipio_antigo
        else:
            antigo = str(municipio_antigo).strip()
            for sufixo in (' / MA', ' - MA', ' (MA)'):
                if antigo.upper().endswith(sufixo):
                    antigo = antigo[:-len(sufixo)].strip()
                    break
            convertido = municipios_antigos_ma.get(normalizar_texto_municipio(antigo))
            if convertido:
                origens_convertidas[equipe] = convertido
    origens_municipais = origens_convertidas
    st.session_state['origem_por_equipe_lista'] = origens_municipais

    def ao_trocar_equipe_origem():
        equipe_atual = st.session_state['equipe_origem_selecionada_lista']
        st.session_state['municipio_origem_selecionado_lista'] = (
            st.session_state['origem_por_equipe_lista'].get(equipe_atual, '-- Selecione --')
        )

    def ao_selecionar_municipio_origem():
        equipe_atual = st.session_state['equipe_origem_selecionada_lista']
        escolhido = st.session_state['municipio_origem_selecionado_lista']
        if escolhido in municipios_ma_pi:
            st.session_state['origem_por_equipe_lista'][equipe_atual] = escolhido
        else:
            st.session_state['origem_por_equipe_lista'].pop(equipe_atual, None)

    equipe_preselecionada = st.session_state.get('equipe_origem_selecionada_lista')
    if equipe_preselecionada not in equipes_disponiveis:
        st.session_state['equipe_origem_selecionada_lista'] = equipes_disponiveis[0]
        st.session_state['municipio_origem_selecionado_lista'] = origens_municipais.get(equipes_disponiveis[0], '-- Selecione --')

    col_eq, col_mu = st.columns(2)
    with col_eq:
        equipe_sel = st.selectbox('1. Selecione a equipe (coluna P):',
                                 equipes_disponiveis,
                                 key='equipe_origem_selecionada_lista',
                                 on_change=ao_trocar_equipe_origem)
    with col_mu:
        opcoes_municipios = ['-- Selecione --'] + list(municipios_ma_pi.keys())
        municipio_atual = st.session_state.get('municipio_origem_selecionado_lista')
        if municipio_atual not in opcoes_municipios:
            st.session_state['municipio_origem_selecionado_lista'] = origens_municipais.get(equipe_sel, '-- Selecione --')
        st.selectbox('2. Município de início do deslocamento (MA/PI):',
                     opcoes_municipios, key='municipio_origem_selecionado_lista',
                     on_change=ao_selecionar_municipio_origem)

    faltando_origem = [e for e in equipes_disponiveis if e not in origens_municipais]
    if faltando_origem:
        st.warning(f"Defina o município de partida para {len(faltando_origem)} equipe(s): " + ', '.join(faltando_origem))
    else:
        st.success(f"✅ Municípios definidos para todas as {len(equipes_disponiveis)} equipes.")

    if origens_municipais:
        conferencia = calcular_previa_origens(df_ta, origens_municipais, municipios_ma_pi, velocidade_media)
        st.markdown('#### 📍 Prévia da primeira obra por equipe')
        st.dataframe(conferencia, use_container_width=True, hide_index=True)
        st.caption('A prioridade é aplicada à PRIMEIRA obra. A distância e o tempo da prévia são estimativas ' 
                   '(Haversine × 1,30 / velocidade média). Com OSRM ativo, o roteiro final apresenta ' 
                   'a distância e o tempo reais do serviço viário para o trecho selecionado.')

    df_u = pd.DataFrame()

    st.session_state.df_unallocated_lista, total_alocadas = df_u, sum(len(r.get('_ORIGINAL_ROWS', [1])) if isinstance(r.get('_ORIGINAL_ROWS'), list) else 1 for _, r in df_ta.iterrows())
    
    sb_html.markdown(render_sidebar_card(total_alocadas, qtd_eq), unsafe_allow_html=True)

    with st.expander("🛠️ Configuração de Saída", expanded=True):
        tc = [c for c in df_ta.columns if not c.startswith('_') and c != 'MUN_LIMPO']
        cd = CAMPOS_VISIVEIS_NIP.copy()
        cp = [c for c in cd if c in tc]
        colunas_exibir = st.multiselect("Colunas Visíveis no Card do KML:", CAMPOS_VISIVEIS_NIP, default=CAMPOS_KML_SOLICITADOS, help="Escolha os campos dos detalhes de cada obra no mapa KML. Os 13 campos solicitados estão pré-selecionados.")
        colunas_exibir.sort(key=lambda x: cd.index(x) if x in cd else 999)

    if st.button("🚀 Iniciar Motor de Roteirização", type="primary", use_container_width=True,
                 disabled=bool(faltando_origem)):
        st.session_state.update({'colunas_exibir_lista': colunas_exibir})
        st.session_state.vrp_state_lista = {
            'config': {
                'velocidade_media_kmh': float(velocidade_media), 'sentido_rota': sentido_rota,
                'url_osrm_base': url_osrm, 'tracado_real': usa_osrm,
                'data_inicio': data_ini, 'obras_por_dia_est': obras_por_dia_est,
                'municipios_origem': {
                    e: {'municipio': origens_municipais[e],
                        'lat': municipios_ma_pi[origens_municipais[e]][0],
                        'lon': municipios_ma_pi[origens_municipais[e]][1]}
                    for e in equipes_disponiveis
                },
            },
            'b_names': equipes_disponiveis,
            'b_idx': 0, 'unvisited': df_ta.copy(),
            'routed_data': [], 'current_geoms': [],
        }
        st.session_state.vrp_status_lista = "RUNNING"; tentar_rerun()

if status_exec == "OSRM_ERROR":
    st.error("🚧 O traçado por ruas foi interrompido: o OSRM não conseguiu traçar um segmento. Nenhum KML/GPX em linha reta será exportado como rota viária.")
    st.warning(st.session_state.vrp_state_lista.get('osrm_error', 'Erro de roteamento não identificado.'))
    c_retry, c_abort = st.columns(2)
    with c_retry:
        if st.button("🔄 Tentar novamente com OSRM", type="primary", use_container_width=True):
            st.session_state.vrp_status_lista = "RUNNING"
            tentar_rerun()
    with c_abort:
        if st.button("⏹️ Abortar roteirização", use_container_width=True):
            limpar_roteirizador()
    st.stop()


if status_exec == "RUNNING":
    st.markdown("## 🚀 Execução do Motor Contínuo")
    if st.button("⏹️ Abortar Execução", use_container_width=True): limpar_roteirizador()
    
    st_run = st.session_state.get('start_time_run_lista', time.time())
    if 'start_time_run_lista' not in st.session_state: st.session_state.start_time_run_lista = st_run
    
    pb = st.progress(0.0); tmp = st.empty(); sgt = st.empty()
    st_v = st.session_state.vrp_state_lista; cfg = st_v['config']; b_n = st_v['b_names']; b_i = st_v.get('b_idx', 0)
    
    def render_t(bi, ii, it):
        e = time.time() - st_run; f = (bi + (ii / max(1, it))) / max(1, len(b_n))
        rs = f"{divmod(int(max(0, (e/f)-e)), 60)[0]:02d}m {divmod(int(max(0, (e/f)-e)), 60)[1]:02d}s" if f > 0.02 else "Calc..."
        es = f"{divmod(int(e), 60)[0]:02d}m {divmod(int(e), 60)[1]:02d}s"
        tmp.markdown(f'<div style="display:flex; gap:15px; margin-bottom: 20px;"><div style="flex:1; padding:20px; border-radius:10px; background-color:#f8f9fa; border:1px solid #dee2e6; text-align:center;"><div style="font-size:0.9rem; color:#6c757d; font-weight:bold; margin-bottom:5px;">⏱️ Decorrido</div><div style="font-size:2rem; font-weight:bold; color:#0D256C;">{es}</div></div><div style="flex:1; padding:20px; border-radius:10px; background-color:#e8f5e9; border:1px solid #a5d6a7; text-align:center;"><div style="font-size:0.9rem; color:#2e7d32; font-weight:bold; margin-bottom:5px;">🎯 Restante</div><div style="font-size:2rem; font-weight:bold; color:#1b5e20;">{rs}</div></div></div>', unsafe_allow_html=True)

    if b_i < len(b_n):
        bn = b_n[b_i]; pb.progress(b_i / max(1, len(b_n))); sgt.info(f"🧠 Roteirizando obras de **{bn}**... ({b_i+1}/{len(b_n)})")
        render_t(b_i, 0, 1)
        
        if 'c_rotas' not in st_v:
            oe = st_v['unvisited'][st_v['unvisited']['BASE_ATRIBUIDA'] == bn].to_dict('records')
            if not oe:
                st_v['b_idx'] += 1
                st.session_state.vrp_state_lista = st_v
                tentar_rerun()
                st.stop()

            origem_info = cfg.get('municipios_origem', {}).get(bn)
            if not origem_info:
                st_v['osrm_error'] = f'Municipio de partida nao definido para {bn}.'
                st.session_state.vrp_status_lista = 'OSRM_ERROR'
                st.session_state.vrp_state_lista = st_v
                tentar_rerun()
                st.stop()

            bl, bL = float(origem_info['lat']), float(origem_info['lon'])
            # Primeira parada: obra PRIORITARIA mais proxima da sede municipal.
            # Sem prioridade, a obra valida mais proxima entre todas da equipe.
            primeira, distancia_direta = escolher_primeira_obra(oe, bl, bL)
            demais = [o for o in oe if o is not primeira]
            restante = ordenar_rota_rapida(
                demais, primeira['LATITUDE'], primeira['LONGITUDE'],
                reversa=("Varredura Reversa" in cfg.get('sentido_rota', 'Lógica Padrão'))
            )
            ot = [primeira] + restante
            
            rf = []
            c_l, c_L = bl, bL
            for o in ot:
                vkr = haversine_vectorized(c_l, c_L, o['LATITUDE'], o['LONGITUDE'])
                vk = vkr * 1.3
                rf.append({'o': o, 'la': c_l, 'La': c_L, 'lt': o['LATITUDE'], 'Lt': o['LONGITUDE'], 'dk': vk})
                c_l, c_L = o['LATITUDE'], o['LONGITUDE']
                
            st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'] = rf, 0, []
            st_v['c_origem_info'] = origem_info
            st.session_state.vrp_state_lista = st_v
            tentar_rerun(); st.stop()
        else:
            rf, oi, gd = st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms']
            ei = min(oi + (8 if cfg.get('tracado_real') else len(rf)), len(rf)) 
            
            for i in range(oi, ei):
                it = rf[i]
                fallback = ([[it['La'], it['la']], [it['Lt'], it['lt']]], (it['dk']*1000/1000.0/cfg['velocidade_media_kmh'])*3600)
                
                if not cfg.get('tracado_real'):
                    gd.append(fallback)
                else:
                    if i % 5 == 0: sgt.info(f"🛣️ Traçando arruamento **{bn}**... ({i+1}/{len(rf)})")
                    render_t(b_i, i, len(rf))
                    try:
                        geometria, segundos, km_ruas = obter_rota_ruas_segura(
                            it['la'], it['La'], it['lt'], it['Lt'],
                            cfg['url_osrm_base'], cfg['velocidade_media_kmh'], timeout_s=15
                        )
                    except FalhaRoteamentoOSRM as exc:
                        # Guarda o progresso ate o trecho que falhou. Nao gera KML falso.
                        st_v['c_idx'] = i
                        st_v['current_geoms'] = gd
                        st_v['osrm_error'] = (
                            f"Colaborador: {bn} | Trecho {i+1}/{len(rf)} | "
                            f"NOTA: {it['o'].get('PROTOCOLO', '-')} | {exc}"
                        )
                        st.session_state.vrp_state_lista = st_v
                        st.session_state.vrp_status_lista = 'OSRM_ERROR'
                        tentar_rerun()
                        st.stop()
                    it['dk'] = km_ruas
                    gd.append((geometria, segundos))
                
            st_v['c_idx'], st_v['current_geoms'] = ei, gd
            if ei < len(rf): st.session_state.vrp_state_lista = st_v; tentar_rerun(); st.stop()
            
            rdf, og = [], 1
            origem_info = st_v['c_origem_info']
            for it, (g, ds) in zip(rf, gd):
                ob = it['o']
                ob['ORDEM'], ob['DISTANCIA_PONTO_ANTERIOR_KM'] = og, round(it['dk'], 2)
                ob['ROTA_GEOMETRIA'], ob['PERIODO'] = g, "Único"
                ob['ORIGEM_MUNICIPIO'] = origem_info['municipio']
                ob['LAT_ORIGEM'] = origem_info['lat']
                ob['LON_ORIGEM'] = origem_info['lon']
                ob['TIPO_DESLOCAMENTO'] = 'Rota viaria OSRM' if cfg.get('tracado_real') else 'Estimativa geografica x 1,30'
                minutos = float(ds) / 60.0
                ob['TEMPO_VIAGEM_MINUTOS'] = round(minutos, 1)
                ob['DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM'] = round(it['dk'], 2) if og == 1 else None
                ob['TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN'] = round(minutos, 1) if og == 1 else None
                rdf.append(ob)
                og += 1
            st_v['routed_data'].extend(rdf)
            del st_v['c_rotas'], st_v['c_idx'], st_v['current_geoms'], st_v['c_origem_info']
            st_v['b_idx'] += 1; st.session_state.vrp_state_lista = st_v; gc.collect(); tentar_rerun()
    else:
        sgt.success("✅ Rotas Finalizadas!"); pb.progress(1.0)
        st.session_state.df_routed_lista = pd.DataFrame(st_v['routed_data'])
        # Invariant: a rota viaria completa deve ter sido obtida, sem fallback.
        if cfg.get('tracado_real'):
            invalidas = st.session_state.df_routed_lista['ROTA_GEOMETRIA'].map(
                lambda g: not isinstance(g, list) or len(g) < 2
            )
            if invalidas.any():
                st_v['osrm_error'] = f'{int(invalidas.sum())} trechos sem geometria valida: exportacao impedida.'
                st.session_state.vrp_status_lista = 'OSRM_ERROR'
                tentar_rerun()
                st.stop()
        st.session_state.vrp_status_lista = "PACKAGING"; time.sleep(1); tentar_rerun()

if status_exec == "PACKAGING":
    st.markdown("## 📦 Empacotamento")
    
    df_routed = st.session_state.df_routed_lista.copy()
    df_routed['DISTANCIA_PROXIMO_PONTO_KM'] = df_routed.groupby(['BASE_ATRIBUIDA'])['DISTANCIA_PONTO_ANTERIOR_KM'].shift(-1).fillna(0.0)

    d_fmt = datetime.now().strftime("%d.%m.%Y")
    bu_xl, bu_kml, bu_gpx, bu_txt, bu_maps, bu_offline = (io.BytesIO() for _ in range(6))
    
    try:
        with zipfile.ZipFile(bu_xl, 'w', zipfile.ZIP_DEFLATED) as zx, \
             zipfile.ZipFile(bu_kml, 'w', zipfile.ZIP_DEFLATED) as zk, \
             zipfile.ZipFile(bu_gpx, 'w', zipfile.ZIP_DEFLATED) as zg, \
             zipfile.ZipFile(bu_txt, 'w', zipfile.ZIP_DEFLATED) as zt, \
             zipfile.ZipFile(bu_maps, 'w', zipfile.ZIP_DEFLATED) as zm, \
             zipfile.ZipFile(bu_offline, 'w', zipfile.ZIP_DEFLATED) as zo:
            
            data_ini = st.session_state.vrp_state_lista.get('config', {}).get('data_inicio', datetime.today())
            obras_por_dia_est = st.session_state.vrp_state_lista.get('config', {}).get('obras_por_dia_est', 4.0)
            
            if isinstance(data_ini, datetime): data_ini = data_ini.date()
            dia_mes_str = data_ini.strftime("%d/%m/%Y")
            dias_semana_pt = {0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA", 3: "QUINTA-FEIRA", 4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO"}
            dia_semana_str = dias_semana_pt[data_ini.weekday()]

            res = []
            for b in df_routed['BASE_ATRIBUIDA'].unique():
                db = df_routed[(df_routed['BASE_ATRIBUIDA']==b) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))]
                qs = len(db[db['SUPER_PONTO'].astype(str).str.startswith('SIM')]) if 'SUPER_PONTO' in db.columns else 0
                
                qtd_obras = 0
                qtd_postes = 0
                for _, r in db.iterrows():
                    if isinstance(r.get('_ORIGINAL_ROWS'), list):
                        qtd_obras += len(r['_ORIGINAL_ROWS'])
                        for orig in r['_ORIGINAL_ROWS']:
                            pv = []
                            for k, v in orig.items():
                                if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                    try: 
                                        val = float(v)
                                        if val > 0: pv.append(val)
                                    except: pass
                            if pv: qtd_postes += min(pv)
                    else:
                        qtd_obras += 1
                        pv = []
                        for k, v in r.items():
                            if 'POSTE' in str(k).upper() and pd.notna(v) and str(v).strip() != '':
                                try: 
                                    val = float(v)
                                    if val > 0: pv.append(val)
                                except: pass
                        if pv: qtd_postes += min(pv)

                postes_dia = (qtd_postes / (qtd_obras / float(obras_por_dia_est))) if qtd_obras > 0 else 0
                postes_semana = postes_dia * 5.0

                res.append({
                    'Equipe': b, 
                    'Obras Roteirizadas': qtd_obras, 
                    'Postes/Dia (Est.)': int(round(postes_dia)),
                    'Postes/Semana (Est.)': int(round(postes_semana)),
                    'Postes Total': int(round(qtd_postes)),
                    'Super Pontos': qs, 
                    'Prioridades Atendidas': len(db[db['PRIORIDADE']=='Sim']), 
                    'KM Total Previsto': round(df_routed[df_routed['BASE_ATRIBUIDA']==b]['DISTANCIA_PONTO_ANTERIOR_KM'].sum(), 2),
                    'Origem Municipal': str(db.iloc[0].get('ORIGEM_MUNICIPIO', '-')),
                    'Nota Inicial': str(db.iloc[0].get('PROTOCOLO', '-')),
                    'Prioridade Inicial': str(db.iloc[0].get('PRIORIDADE', '-')),
                    'KM Ate 1a Obra': db.iloc[0].get('DISTANCIA_ORIGEM_PRIMEIRA_OBRA_KM', 0),
                    'Min Ate 1a Obra': db.iloc[0].get('TEMPO_ORIGEM_PRIMEIRA_OBRA_MIN', 0),
                    'Metodo de Deslocamento': str(db.iloc[0].get('TIPO_DESLOCAMENTO', '-'))
                })
            zx.writestr(f"Resumo_Operacional - {d_fmt}.xlsx", gerar_excel_resumo_lista(pd.DataFrame(res)))
            
            dfc = st.session_state.get('df_correcao_lista', pd.DataFrame())
            if not dfc.empty:
                dfcc = dfc.copy(); dfcc.rename(columns={'LEVANTADOR': 'FISCAL', 'PROTOCOLO': 'NOTA'}, inplace=True)
                dfcc = dfcc.loc[:, ~dfcc.columns.duplicated()].copy()
                for cc in dfcc.columns:
                    if str(dfcc[cc].dtype) == 'object': dfcc[cc] = dfcc[cc].astype(str).replace('nan', '')
                out_e = io.BytesIO(); dfcc.to_excel(out_e, index=False); zx.writestr(f"Obras_Correcao - {d_fmt}.xlsx", out_e.getvalue())
            
            linhas_gerais = []
            for _, r in df_routed.iterrows():
                if r.get('PROTOCOLO') in ['RETORNO_BASE', 'PAUSA_ALMOCO']: continue
                
                is_sp = isinstance(r.get('_ORIGINAL_ROWS'), list) and len(r.get('_ORIGINAL_ROWS')) > 1
                sp_text = f"SIM ({len(r['_ORIGINAL_ROWS'])} Obras)" if is_sp else "NÃO"
                
                if is_sp:
                    for orig in r['_ORIGINAL_ROWS']:
                        nr = r.copy()
                        for k, v in orig.items(): 
                            if k not in ['BASE_ATRIBUIDA', 'LEVANTADOR', 'FISCAL', 'ORDEM', 'DISTANCIA_PONTO_ANTERIOR_KM', 'DISTANCIA_PROXIMO_PONTO_KM', 'ROTA_GEOMETRIA', 'PERIODO']: nr[k] = v
                        nr['DIA_SEMANA'] = dia_semana_str
                        nr['DIA_MES'] = dia_mes_str
                        nr['SUPER_PONTO'] = sp_text
                        linhas_gerais.append(nr)
                else:
                    rn = r.copy()
                    rn['DIA_SEMANA'] = dia_semana_str
                    rn['DIA_MES'] = dia_mes_str
                    rn['SUPER_PONTO'] = sp_text
                    linhas_gerais.append(rn)
            
            df_excel_full = pd.DataFrame(linhas_gerais)
            
            for c in df_excel_full.columns:
                if 'POSTE' in c.upper():
                    df_excel_full[c] = pd.to_numeric(df_excel_full[c], errors='coerce').apply(lambda x: str(int(x)) if pd.notna(x) else '')

            col_exibir = st.session_state.colunas_exibir_lista.copy()
            # O nome do cliente integra todas as exportacoes.
            if 'NOME' not in col_exibir: col_exibir.append('NOME')
            if 'DIA_SEMANA' not in col_exibir: col_exibir.insert(0, 'DIA_SEMANA')
            if 'DIA_MES' not in col_exibir: col_exibir.insert(1, 'DIA_MES')
            if 'SUPER_PONTO' not in col_exibir: col_exibir.insert(2, 'SUPER_PONTO')

            dfg_total = limpar_colunas_lista(df_excel_full.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
            dfg_total = dfg_total.loc[:, ~dfg_total.columns.duplicated()].copy()
            for cc in dfg_total.columns:
                if str(dfg_total[cc].dtype) == 'object': dfg_total[cc] = dfg_total[cc].astype(str).replace('nan', '')
            zx.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.xlsx", gerar_excel_lista(dfg_total, st.session_state.colunas_originais_lista))
            
            # --- TXT TOTAL ---
            txt_total = gerar_txt_lista(df_excel_full)
            zt.writestr(f"Demanda_ListaContinua_Total - {d_fmt}.txt", txt_total.encode('utf-8'))
            
            # Um unico PNG por pasta, referenciado pelos arquivos KML dessa pasta.
            zk.writestr(ICONE_CASA_ARQUIVO, imagem_casa_bytes())
            zk.writestr(f"KML_{d_fmt}/{ICONE_CASA_ARQUIVO}", imagem_casa_bytes())

            dfk_total = df_routed[~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO'])].copy()
            if not dfk_total.empty:
                dfk_total['DIA_SEMANA'] = dia_semana_str
                dfk_total['DIA_MES'] = dia_mes_str
                dfk_total['SUPER_PONTO'] = dfk_total.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                
                ks_tot = gerar_kml_lista(dfk_total, "ROTA TOTAL", col_exibir, df_routed['BASE_ATRIBUIDA'].unique().tolist(), formatar_valor_coluna)
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kml", ks_tot.encode('utf-8'))
                zk.writestr(f"ROTA_TOTAL - {d_fmt}.kmz", gerar_kmz_lista(ks_tot))
                zg.writestr(f"GPS_TOTAL - {d_fmt}.gpx", gerar_gpx_simples(dfk_total, "ROTA TOTAL").encode('utf-8'))

            entradas_maps = []
            mapas_primeiras = {}
            registros_mymaps_geral = []
            for base in df_routed['BASE_ATRIBUIDA'].unique():
                b_safe = re.sub(r'[^A-Za-z0-9_ -]', '', str(base)).strip()
                
                df_base_excel = df_excel_full[df_excel_full['BASE_ATRIBUIDA'] == base]
                if not df_base_excel.empty:
                    dfg = limpar_colunas_lista(df_base_excel.drop(columns=['MUN_LIMPO', 'COR_ICONE', 'COORD_KEY', 'ALERTA_TOPOLOGIA', 'ROTA_GEOMETRIA', 'PERIODO', '_HORA_INICIO_DT', '_HORA_FIM_DT', 'HORA_INICIO', 'HORA_FIM', 'TEMPO_VIAGEM_MINUTOS', '_ORIGINAL_ROWS'], errors='ignore'), col_exibir)
                    dfg = dfg.loc[:, ~dfg.columns.duplicated()].copy()
                    for cc in dfg.columns:
                        if str(dfg[cc].dtype) == 'object': dfg[cc] = dfg[cc].astype(str).replace('nan', '')
                    zx.writestr(f"Rotas_{d_fmt}/Rota_{b_safe}.xlsx", gerar_excel_lista(dfg, st.session_state.colunas_originais_lista))
                    
                    # --- TXT INDIVIDUAL ---
                    txt_ind = gerar_txt_lista(df_base_excel)
                    if txt_ind:
                        zt.writestr(f"Relatorios_TXT_{d_fmt}/Relatorio_{b_safe}.txt", txt_ind.encode('utf-8'))
                    
                dfk_base = df_routed[(df_routed['BASE_ATRIBUIDA'] == base) & (~df_routed['PROTOCOLO'].isin(['RETORNO_BASE', 'PAUSA_ALMOCO']))].copy()
                if not dfk_base.empty:
                    dfk_base['DIA_SEMANA'] = dia_semana_str
                    dfk_base['DIA_MES'] = dia_mes_str
                    dfk_base['SUPER_PONTO'] = dfk_base.apply(lambda row_k: f"SIM ({len(row_k['_ORIGINAL_ROWS'])} Obras)" if isinstance(row_k.get('_ORIGINAL_ROWS'), list) and len(row_k['_ORIGINAL_ROWS'])>1 else "NÃO", axis=1)
                    
                    ks = gerar_kml_lista(dfk_base, f"Rota {b_safe}", col_exibir, [base], formatar_valor_coluna)
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kml", ks.encode('utf-8'))
                    zk.writestr(f"KML_{d_fmt}/Rota_{b_safe}.kmz", gerar_kmz_lista(ks))
                    zg.writestr(f"GPX_{d_fmt}/Rota_{b_safe}.gpx", gerar_gpx_simples(dfk_base, f"Rota {b_safe}").encode('utf-8'))

                    # Exportacao Google Maps independente por equipe.
                    # HTML para ordenar as visitas; links TXT para WhatsApp;
                    # CSV para importar os marcadores no Google My Maps.
                    id_equipe_maps = hashlib.sha256(str(base).encode('utf-8')).hexdigest()[:8]
                    chave_arquivo = f'{b_safe}_{id_equipe_maps}'
                    nome_html = f'Equipes/Planejador_{chave_arquivo}.html'
                    nome_html_ios = f'Equipes/Roteiro_iPhone_{chave_arquivo}.html'
                    nome_csv = f'Equipes/Pontos_MyMaps_{chave_arquivo}.csv'
                    nome_links = f'Equipes/LINKS_Maps_{chave_arquivo}.txt'
                    html_campo = gerar_painel_google_maps(dfk_base, str(base))
                    html_ios = gerar_painel_iphone_estatico(dfk_base, str(base))
                    gpx_offline = gerar_gpx_offline_paradas(dfk_base, str(base))
                    nome_gpx_offline = f'Equipes/Pontos_Offline_{chave_arquivo}.gpx'
                    zm.writestr(nome_html, html_campo.encode('utf-8'))
                    zm.writestr(nome_gpx_offline, gpx_offline.encode('utf-8'))
                    zo.writestr(nome_html, html_campo.encode('utf-8'))
                    zo.writestr(nome_gpx_offline, gpx_offline.encode('utf-8'))
                    zo.writestr(nome_html_ios, html_ios.encode('utf-8'))
                    # Opcional: publicacao HTTPS para modo instalavel e cache no iOS.
                    pontos_versao = hashlib.sha256(json.dumps(
                        {'origem': origem_google_maps(dfk_base),
                         'obras': pontos_equipe_google_maps(dfk_base)},
                        ensure_ascii=False, sort_keys=True
                    ).encode('utf-8')).hexdigest()[:16]
                    html_pwa = gerar_painel_google_maps(dfk_base, str(base), pwa=True)
                    pasta_pwa = f'Offline_Instalavel/{chave_arquivo}/'
                    for relativo, conteudo in gerar_pwa_offline(str(base), pontos_versao, nome_html, html_pwa).items():
                        zo.writestr(pasta_pwa + relativo, conteudo)

                    zm.writestr(nome_html_ios, html_ios.encode('utf-8'))
                    csv_equipe = gerar_csv_google_mymaps(dfk_base, str(base))
                    for parte_idx, parte in enumerate(csv_google_mymaps_em_partes(csv_equipe), start=1):
                        nome_parte = nome_csv if parte_idx == 1 else nome_csv.replace('.csv', f'_PARTE_{parte_idx:03d}.csv')
                        zm.writestr(nome_parte, ('\ufeff' + parte).encode('utf-8'))
                    zm.writestr(nome_links, gerar_txt_links_google_maps(dfk_base, str(base)).encode('utf-8'))
                    registros_mymaps_geral.extend(list(csv.DictReader(io.StringIO(csv_equipe))))
                    orig_m = origem_google_maps(dfk_base)
                    pontos_m = pontos_equipe_google_maps(dfk_base)
                    entradas_maps.append({
                        'equipe': str(base), 'arquivo': nome_html,
                        'arquivo_ios': nome_html_ios,
                        'municipio': orig_m['municipio'],
                        'primeira': pontos_m[0]['nota'] if pontos_m else '-',
                    })
                    if pontos_m:
                        mapas_primeiras[str(base)] = {
                            'nota': pontos_m[0]['nota'],
                            'url': link_google_maps_primeira_obra(orig_m, pontos_m[0]),
                        }

            if entradas_maps:
                zm.writestr('INICIAR_AQUI.html', gerar_indice_google_maps(entradas_maps).encode('utf-8'))
                zo.writestr('INICIAR_AQUI.html', gerar_indice_google_maps(entradas_maps).encode('utf-8'))
                zo.writestr('LEIA_ME_OFFLINE.txt', gerar_guia_modo_offline().encode('utf-8'))
                zm.writestr('LEIA_ME_OFFLINE.txt', gerar_guia_modo_offline().encode('utf-8'))
                zm.writestr('LEIA_ME.txt', gerar_leia_me_google_maps().encode('utf-8'))
                zm.writestr('LEIA_ME_IPHONE.txt', gerar_leia_me_iphone().encode('utf-8'))
                csv_geral = io.StringIO(newline='')
                chaves_geral = ['Nome', 'Equipe', 'Ordem', 'Tipo', 'Prioridade',
                                'Nota CCS', 'Municipio', 'Localidade', 'Endereco',
                                'Latitude', 'Longitude', 'Qtd Obras']
                escritor = csv.DictWriter(csv_geral, fieldnames=chaves_geral, lineterminator='\r\n')
                escritor.writeheader()
                escritor.writerows(registros_mymaps_geral)
                for parte_idx, parte in enumerate(csv_google_mymaps_em_partes(csv_geral.getvalue()), start=1):
                    nome_parte = ('MAPA_GERAL_GoogleMyMaps.csv' if parte_idx == 1
                                  else f'MAPA_GERAL_GoogleMyMaps_PARTE_{parte_idx:03d}.csv')
                    zm.writestr(nome_parte, ('\ufeff' + parte).encode('utf-8'))

        st.session_state.bytes_zip_xl_lista = bu_xl.getvalue()
        st.session_state.bytes_zip_kml_lista = bu_kml.getvalue()
        st.session_state.bytes_zip_gpx_lista = bu_gpx.getvalue()
        st.session_state.bytes_zip_txt_lista = bu_txt.getvalue()
        st.session_state.bytes_zip_maps_lista = bu_maps.getvalue()
        st.session_state.bytes_zip_offline_lista = bu_offline.getvalue()
        st.session_state.maps_primeiras_lista = mapas_primeiras
        st.session_state.roteamento_concluido_lista = True; st.session_state.vrp_status_lista = "IDLE"; tentar_rerun()
    except Exception as e: st.error(f"🚨 ERRO: {e}"); st.session_state.vrp_status_lista = "IDLE"
