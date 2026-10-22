# Roteirização NIP - baseada em 2_Lista_Continua(6).py.
# K = COLABORADOR PLANEJADO
from __future__ import annotations
from pathlib import Path
from io import BytesIO
from collections import defaultdict
from datetime import datetime
import math
import re
import html
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import folium
import requests
import time
from folium.plugins import MarkerCluster, Fullscreen
from streamlit_folium import st_folium

ROOT = Path(__file__).resolve().parents[1]
COLUMNS = ['NOTA','STATUS','TIPO DEMANDA','MUNICIPIO','REGIONAL','ENDEREÇO','BAIRRO','LATITUDE','LONGITUDE','COLABORADOR','MATRICULA']
COLS_ORIG = {'NOTA':'NOTA','STATUS':'STATUS','TIPO DEMANDA':'TIPO DEMANDA','MUNICIPIO':'MUNICÍPIO','REGIONAL':'REGIONAL SERVIÇO','ENDEREÇO':'ENDEREÇO','BAIRRO':'BAIRRO','LATITUDE':'LATITUDE','LONGITUDE':'LONGITUDE','COLABORADOR':'COLABORADOR PLANEJADO','MATRICULA':'MATRÍCULA'}

def esc(x):
    if pd.isna(x): return ''
    return html.escape(str(x),quote=True)

def safe_filename(s):
    return re.sub(r'[^a-zA-Z0-9 _.-]+','_',str(s)).strip(' ._')[:90] or 'SEM_COLABORADOR'

@st.cache_data(show_spinner=False)
def load_file(raw:bytes):
    df=pd.read_excel(BytesIO(raw),dtype=str,engine='openpyxl')
    df.columns=[str(c).strip() for c in df.columns]
    missing=[v for v in COLS_ORIG.values() if v not in df.columns]
    if missing: raise ValueError('Colunas ausentes: '+', '.join(missing))
    # Verifica explicitamente a posicao solicitada, sem adivinhar o responsavel.
    if df.columns[10]!='COLABORADOR PLANEJADO':
        raise ValueError('A coluna K deve ser COLABORADOR PLANEJADO; confira o arquivo.')
    df=df.rename(columns={v:k for k,v in COLS_ORIG.items()})[COLUMNS].copy()
    for c in COLUMNS: df[c]=df[c].fillna('').astype(str).str.strip()
    for c in ['LATITUDE','LONGITUDE']:
        df[c+'_NUM']=pd.to_numeric(df[c].str.replace(',','.',regex=False),errors='coerce')
    df['VALIDA']=df['LATITUDE_NUM'].between(-90,90)&df['LONGITUDE_NUM'].between(-180,180)&(df['LATITUDE_NUM']!=0)&(df['LONGITUDE_NUM']!=0)
    return df

def haversine(a,b,c,d):
    r=6371.0088
    a,b,c,d=map(math.radians,[a,b,c,d]);dl=c-a;dn=d-b
    h=math.sin(dl/2)**2+math.cos(a)*math.cos(c)*math.sin(dn/2)**2
    return 2*r*math.asin(min(1,math.sqrt(h)))

def order_route(frame,reverse=False):
    rows=frame.to_dict('records')
    if not rows:return []
    la=sum(float(x['LATITUDE_NUM']) for x in rows)/len(rows)
    lo=sum(float(x['LONGITUDE_NUM']) for x in rows)/len(rows)
    if reverse:
        start=max(range(len(rows)),key=lambda i:haversine(la,lo,float(rows[i]['LATITUDE_NUM']),float(rows[i]['LONGITUDE_NUM'])))
        rows[0],rows[start]=rows[start],rows[0]
        result=[rows.pop(0)];la=float(result[-1]['LATITUDE_NUM']);lo=float(result[-1]['LONGITUDE_NUM'])
    else:result=[]
    while rows:
        i=min(range(len(rows)),key=lambda i:haversine(la,lo,float(rows[i]['LATITUDE_NUM']),float(rows[i]['LONGITUDE_NUM'])))
        x=rows.pop(i);x['DISTANCIA_PONTO_ANTERIOR_KM']=round(haversine(la,lo,float(x['LATITUDE_NUM']),float(x['LONGITUDE_NUM'])),3)
        result.append(x);la=float(x['LATITUDE_NUM']);lo=float(x['LONGITUDE_NUM'])
    for i,x in enumerate(result):
        x['ORDEM']=i+1
        if 'DISTANCIA_PONTO_ANTERIOR_KM' not in x:x['DISTANCIA_PONTO_ANTERIOR_KM']=0
    return result

def popup(row):
    """Card no estilo do KML original: cabeçalho azul, prioridade vermelha, superponto amarelo."""
    priority=str(row.get('PRIORIDADE','')).strip().upper() in ('SIM','S','1','TRUE')
    superpoint=bool(row.get('SUPER_PONTO'))
    label='SUPER PONTO' if superpoint else ('OBRA PRIORITÁRIA' if priority else 'ATENDIMENTO PADRÃO')
    bg='#ffd400' if superpoint else ('#df5151' if priority else '#102d76')
    fg='#222' if superpoint else '#fff'
    rows=''.join('<div style="display:grid;grid-template-columns:102px 1fr;gap:7px;padding:5px 7px;border-bottom:1px solid #edf0f3">'
        '<span style="font-size:10px;color:#536073;font-weight:700">'+esc(c)+':</span>'
        '<span style="font-size:11px;overflow-wrap:anywhere">'+esc(row.get(c,''))+'</span></div>' for c in COLUMNS)
    lat=esc(row.get('LATITUDE_NUM',''));lon=esc(row.get('LONGITUDE_NUM',''))
    return ('<div style="font-family:Arial,sans-serif;width:265px;color:#1b2737">'
        '<div style="font-weight:800;font-size:12px;margin-bottom:7px">['+esc(row.get('ORDEM',''))+'] Doc: '+esc(row.get('NOTA',''))+'</div>'
        '<div style="background:'+bg+';color:'+fg+';font-weight:800;font-size:12px;padding:9px;border-radius:5px 5px 0 0">'+label+'</div>'
        '<div style="border:1px solid #e3e6eb;border-radius:0 0 7px 7px;background:white">'
        '<div style="padding:7px 9px;font-size:11px"><b>Ordem:</b> '+esc(row.get('ORDEM',''))+
        ' &nbsp; <b>Distância ant.:</b> '+esc(row.get('DISTANCIA_PONTO_ANTERIOR_KM',0))+' km</div>'+rows+'</div>'
        '<div style="margin-top:8px;font-size:11px">Rotas: <a target="_blank" href="https://www.google.com/maps/dir/?api=1&amp;destination='+lat+','+lon+'">Até aqui</a> - '
        '<a target="_blank" href="https://www.google.com/maps/search/?api=1&amp;query='+lat+','+lon+'">Daqui</a></div></div>')

def kml_content(records,title):
    kml=ET.Element('kml',xmlns='http://www.opengis.net/kml/2.2');doc=ET.SubElement(kml,'Document');ET.SubElement(doc,'name').text=title
    for sid,href in [('normal','http://maps.google.com/mapfiles/kml/paddle/blu-blank.png'),('priority','http://maps.google.com/mapfiles/kml/paddle/red-blank.png'),('super','http://maps.google.com/mapfiles/kml/paddle/ylw-blank.png')]:
        style=ET.SubElement(doc,'Style',id=sid)
        iconstyle=ET.SubElement(style,'IconStyle')
        ET.SubElement(iconstyle,'scale').text='1.2'
        ET.SubElement(ET.SubElement(iconstyle,'Icon'),'href').text=href
    by=defaultdict(list)
    for r in records:by[r['COLABORADOR']].append(r)
    for name,rows in by.items():
        folder=ET.SubElement(doc,'Folder');ET.SubElement(folder,'name').text=name or 'SEM COLABORADOR'
        clustered=defaultdict(list)
        for r in rows:clustered[r.get('SUPER_GRUPO',str(r['NOTA']))].append(r)
        for group in clustered.values():
            r=group[0]
            point=ET.SubElement(folder,'Placemark')
            ET.SubElement(point,'name').text=(f"[{r['ORDEM']}] SUPER PONTO ({len(group)} obras)" if len(group)>1 else f"{r['ORDEM']:03d} | {r['NOTA']}")
            desc='<div style="font-family:Arial"><h3>SUPER PONTO ('+str(len(group))+' obras)</h3>'+''.join(popup(item)+'<hr>' for item in group)+'</div>' if len(group)>1 else popup(r)
            ET.SubElement(point,'description').text=desc
            ET.SubElement(point,'styleUrl').text=('#super' if len(group)>1 else ('#priority' if str(r.get('PRIORIDADE','')).upper() in ('SIM','S','1','TRUE') else '#normal'))
            ET.SubElement(ET.SubElement(point,'Point'),'coordinates').text=f"{r['LONGITUDE_NUM']},{r['LATITUDE_NUM']},0"
        if len(rows)>1:
            line=ET.SubElement(folder,'Placemark');ET.SubElement(line,'name').text='Sequência da rota (linha reta)'
            ET.SubElement(ET.SubElement(line,'LineString'),'coordinates').text=' '.join(f"{r['LONGITUDE_NUM']},{r['LATITUDE_NUM']},0" for r in rows)
    return ET.tostring(kml,encoding='utf-8',xml_declaration=True)

def gpx_content(records,title):
    root=ET.Element('gpx',version='1.1',creator='NIP Roteirizacao',xmlns='http://www.topografix.com/GPX/1/1')
    ET.SubElement(root,'name').text=title
    by=defaultdict(list)
    for r in records:
        w=ET.SubElement(root,'wpt',lat=str(r['LATITUDE_NUM']),lon=str(r['LONGITUDE_NUM']))
        ET.SubElement(w,'name').text=str(r['NOTA']);ET.SubElement(w,'desc').text=f"{r['COLABORADOR']} | {r['STATUS']} | {r['MUNICIPIO']}"
        by[r['COLABORADOR']].append(r)
    for name,rows in by.items():
        trk=ET.SubElement(root,'trk');ET.SubElement(trk,'name').text=name
        seg=ET.SubElement(trk,'trkseg')
        for r in rows:ET.SubElement(seg,'trkpt',lat=str(r['LATITUDE_NUM']),lon=str(r['LONGITUDE_NUM']))
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)

def txt_content(rows):
    return ('\n\n'.join('ORDEM: '+str(r['ORDEM'])+'\n'+'\n'.join(f'{c}: {r.get(c,"")}' for c in COLUMNS) for r in rows)).encode('utf-8-sig')

def excel_content(rows):
    out=BytesIO();pd.DataFrame(rows).reindex(columns=['ORDEM']+COLUMNS).to_excel(out,index=False,engine='openpyxl');return out.getvalue()

def export_zip(rows):
    b=BytesIO();today=datetime.now().strftime('%Y%m%d')
    with zipfile.ZipFile(b,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('ROTA_TOTAL.kml',kml_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.gpx',gpx_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.xlsx',excel_content(rows))
        z.writestr('ROTA_TOTAL.txt',txt_content(rows))
        kmz=BytesIO()
        with zipfile.ZipFile(kmz,'w',zipfile.ZIP_DEFLATED) as km:zkm=km.writestr('doc.kml',kml_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.kmz',kmz.getvalue())
        by=defaultdict(list)
        for r in rows:by[r['COLABORADOR']].append(r)
        for name,group in by.items():
            stem='POR_COLABORADOR/'+safe_filename(name)
            z.writestr(stem+'.kml',kml_content(group,name));z.writestr(stem+'.gpx',gpx_content(group,name))
            z.writestr(stem+'.xlsx',excel_content(group));z.writestr(stem+'.txt',txt_content(group))
            kmz=BytesIO()
            with zipfile.ZipFile(kmz,'w',zipfile.ZIP_DEFLATED) as km:km.writestr('doc.kml',kml_content(group,name))
            z.writestr(stem+'.kmz',kmz.getvalue())
    return b.getvalue()


# CONFIGURACAO LOGISTICA — interface da aba, isolada do HTML da Visao Geral.
st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#f5f7fa}
.block-container{max-width:100%;padding-top:.8rem;padding-bottom:2rem}
.nip-banner{background:#e4f0ff;color:#145b9a;padding:11px 16px;border-radius:5px;font-size:13px;margin-bottom:15px}
.nip-result{background:#d8f1dd;border-left:5px solid #19a04a;border-radius:5px;padding:13px;color:#145a2b}
</style>""",unsafe_allow_html=True)
st.title('📜 Lista Contínua')
st.markdown('<div class="nip-banner">💡 Gera uma lista de execução contínua com base na atribuição já existente na planilha de obras. Nenhuma obra é transferida de colaborador.</div>',unsafe_allow_html=True)
ROOT=Path(__file__).resolve().parents[1]

with st.sidebar:
    st.markdown('### ⚙️ Configurações Logísticas')
    with st.expander('Parâmetros de rota',expanded=True):
        st.success('📦 Modo Contínuo: as obras são organizadas numa lista única por colaborador.')
        cap=st.number_input('Trava total de obras (0 = sem limite)',min_value=0,max_value=200000,value=0,step=1)
        start_date=st.date_input('Data base (início)',value=datetime.now().date(),format='DD/MM/YYYY')
        daily=st.number_input('Meta de obras/dia (cálculo de postes)',min_value=1,max_value=1000,value=4)
        reverse=st.radio('Sentido do roteamento',['Lógica padrão','Varredura reversa'])=='Varredura reversa'
        radius=st.slider('Raio Super Ponto (m)',min_value=10,max_value=300,value=50,step=10)
    with st.expander('Conexão de rede',expanded=True):
        osrm_url=st.text_input('Endpoint OSRM',value='https://router.project-osrm.org',help='Serviço externo; sujeito a indisponibilidade e limites.')
        use_roads=st.checkbox('Traçado de ruas real (lento)',value=False)
        max_calls=st.number_input('Limite de trechos OSRM por execução',min_value=1,max_value=500,value=60)

st.markdown('### 📁 Demandas (Obras e Equipes)')
st.markdown('<div class="nip-banner">💡 A coluna K — COLABORADOR PLANEJADO — define o responsável. O sistema não redistribui obras.</div>',unsafe_allow_html=True)
file=next((p for p in [ROOT/'BASE_SANEAMENTO(1).xlsx',ROOT/'BASE_SANEAMENTO.xlsx',ROOT/'BASE_SANEAMENTO_ATUALIZADA.xlsx'] if p.is_file()),None)
with st.expander('Base de dados',expanded=True):
    if file: st.caption('Arquivo encontrado no repositório: '+file.name)
    up=st.file_uploader('Suba a planilha com a Lista Contínua (campo original)',type=['xlsx','xls','csv'],key='nip_route_upload_original',help='Campo original mantido. Para roteirizar a planilha de distribuição, utilize o novo campo abaixo.')
    st.markdown('#### 📥 Planilha padrão para roteirização')
    st.caption('Novo campo — planilha com as colunas A:L no padrão BASE_SANEAMENTO(1).xlsx. A coluna K é COLABORADOR PLANEJADO.')
    up_rota=st.file_uploader('Enviar planilha de distribuição / roteirização',type=['xlsx','xls','csv'],key='nip_route_upload_distribuicao')

# O segundo campo prevalece; o primeiro permanece disponível e pode servir como entrada
# quando estiver no mesmo padrão de distribuição. Arquivos com outro layout não são misturados.
chosen=up_rota or up or file
if chosen is None:
    st.warning('Carregue a planilha padrão no novo campo para continuar.');st.stop()
try:
    raw=chosen.getvalue() if hasattr(chosen,'getvalue') else chosen.read_bytes()
    if str(getattr(chosen,'name',chosen)).lower().endswith('.csv'):
        source=pd.read_csv(BytesIO(raw),dtype=str,sep=None,engine='python')
        temp=BytesIO()
        source.to_excel(temp,index=False,engine='openpyxl')
        raw=temp.getvalue()
    df=load_file(raw)
    st.success(f'Planilha de roteirização: {getattr(chosen,"name",str(chosen))} — {len(df):,} registros lidos.'.replace(',','.'))
except Exception as e:
    st.error(f'Falha na leitura da planilha de roteirização: {e}')
    st.info('Padrão obrigatório: NOTA, STATUS, TIPO DEMANDA, MUNICÍPIO, REGIONAL SERVIÇO, ENDEREÇO, BAIRRO, LATITUDE, LONGITUDE, COLABORADOR PLANEJADO (coluna K) e MATRÍCULA.')
    st.stop()

st.markdown('### 🌐 Limpeza Geográfica')
f1,f2,f3,f4=st.columns(4)
with f1: colaboradores=st.multiselect('Colaborador planejado',sorted(x for x in df['COLABORADOR'].unique() if x))
with f2: regionais=st.multiselect('Regional',sorted(x for x in df['REGIONAL'].unique() if x))
with f3: municipios=st.multiselect('Município',sorted(x for x in df['MUNICIPIO'].unique() if x))
with f4: status=st.multiselect('Status',sorted(x for x in df['STATUS'].unique() if x))
for key,sel in [('COLABORADOR',colaboradores),('REGIONAL',regionais),('MUNICIPIO',municipios),('STATUS',status)]:
    if sel:df=df[df[key].isin(sel)]
q=st.text_input('Buscar nota, endereço, bairro ou matrícula')
if q:
    pattern=re.escape(q)
    df=df[df[['NOTA','ENDEREÇO','BAIRRO','MATRICULA']].apply(lambda c:c.str.contains(pattern,case=False,na=False,regex=True)).any(axis=1)]
valid=df[df['VALIDA']].copy();invalid=df[~df['VALIDA']].copy()
if cap>0:valid=valid.head(cap).copy()
with st.expander('Configuração de saída — Colunas visíveis',expanded=True):
    st.caption('NOTA · STATUS · TIPO DEMANDA · MUNICIPIO · REGIONAL · ENDEREÇO · BAIRRO · LATITUDE · LONGITUDE · COLABORADOR · MATRICULA')
    st.caption('Colaborador obtido da coluna K (COLABORADOR PLANEJADO).')
if len(invalid):
    st.warning(f'{len(invalid)} registros não têm coordenadas válidas; não serão roteirizados.')
    with st.expander('Conferir registros sem coordenadas'):st.dataframe(invalid[COLUMNS],hide_index=True,use_container_width=True)
if valid.empty:
    st.info('Nenhuma obra válida selecionada.');st.stop()

# Agrupa pontos num raio aproximado em metros por colaborador, preservando todas as notas.
def tag_superpoints(rows,meters):
    from collections import defaultdict
    by=defaultdict(list)
    for r in rows:by[r['COLABORADOR']].append(r)
    super_count=0
    for _,group in by.items():
        for i,r in enumerate(group):
            if r.get('_clustered'):continue
            same=[x for x in group[i:] if not x.get('_clustered') and haversine(float(r['LATITUDE_NUM']),float(r['LONGITUDE_NUM']),float(x['LATITUDE_NUM']),float(x['LONGITUDE_NUM']))*1000<=meters]
            group_id=str(r['COLABORADOR'])+'|'+str(i)
            for x in same:x['_clustered']=True;x['SUPER_PONTO']=len(same)>1;x['SUPER_PONTO_QTD']=len(same);x['SUPER_GRUPO']=group_id
            if len(same)>1:super_count+=1
    for r in rows:r.pop('_clustered',None)
    return super_count

@st.cache_data(show_spinner=False,ttl=3600)
def osrm_segment(a,b,base):
    try:
        base=base.rstrip('/')
        if not base.startswith(('https://','http://')):return None
        url=f'{base}/route/v1/driving/{a[1]},{a[0]};{b[1]},{b[0]}'
        data=requests.get(url,params={'overview':'full','geometries':'geojson','steps':'false'},timeout=7).json()
        if data.get('code')=='Ok' and data.get('routes'):
            rt=data['routes'][0]
            return {'km':round(rt['distance']/1000,3),'points':[[c[1],c[0]] for c in rt['geometry']['coordinates']]}
    except (requests.RequestException,ValueError,KeyError,TypeError):pass
    return None

def generate_plan(valid,reverse,radius,use_roads,max_calls,osrm_url):
    rows=[]
    for _,group in valid.groupby('COLABORADOR',sort=True,dropna=False):
        rows.extend(order_route(group,reverse=reverse))
    n_super=tag_superpoints(rows,radius)
    paths={};total_km=0;real=0;fallback=0;calls=0
    by=defaultdict(list)
    for r in rows:by[r['COLABORADOR']].append(r)
    for name,group in by.items():
        pts=[]
        for i,r in enumerate(group):
            dest=[float(r['LATITUDE_NUM']),float(r['LONGITUDE_NUM'])]
            if i==0:pts.append(dest);r['DISTANCIA_PONTO_ANTERIOR_KM']=0;continue
            origin=[float(group[i-1]['LATITUDE_NUM']),float(group[i-1]['LONGITUDE_NUM'])]
            road=None
            if use_roads and calls<max_calls:
                calls+=1;road=osrm_segment(origin,dest,osrm_url)
            if road:
                real+=1;km=road['km'];pts.extend(road['points'][1:])
            else:
                fallback+=1;km=round(haversine(*origin,*dest),3);pts.append(dest)
            r['DISTANCIA_PONTO_ANTERIOR_KM']=km
            total_km+=km
        paths[name]=pts
    return rows,paths,round(total_km,2),n_super,real,fallback

st.divider()
if st.button('🚀 Iniciar Motor de Roteirização',type='primary',use_container_width=True):
    with st.spinner('Organizando as obras e consultando o OSRM quando habilitado...'):
        begin=time.monotonic()
        rows,paths,km,n_super,real,fallback=generate_plan(valid,reverse,radius,use_roads,int(max_calls),osrm_url)
        st.session_state['nip_plan']={'rows':rows,'paths':paths,'km':km,'super':n_super,'real':real,'fallback':fallback,'elapsed':time.monotonic()-begin}
        st.session_state['nip_plan_key']=(len(valid),tuple(valid['NOTA'].head(10)),tuple(colaboradores),tuple(regionais),tuple(municipios),tuple(status),q,cap,reverse,radius,use_roads,max_calls,osrm_url)
key=(len(valid),tuple(valid['NOTA'].head(10)),tuple(colaboradores),tuple(regionais),tuple(municipios),tuple(status),q,cap,reverse,radius,use_roads,max_calls,osrm_url)
if st.session_state.get('nip_plan_key')!=key:
    st.info('Configure os parâmetros e clique em Iniciar Motor de Roteirização.');st.stop()
plan=st.session_state['nip_plan'];rows=plan['rows'];paths=plan['paths']
st.markdown('### 🎯 Resultado do Planejamento')
c1,c2,c3,c4=st.columns(4)
c1.metric('🎯 Obras roteirizadas',len(rows))
c2.metric('👥 Equipes alocadas',len(paths))
c3.metric('📦 Super pontos',plan['super'])
c4.metric('🗺️ KM total estimado',f"{plan['km']:.2f} km")
st.markdown(f'<div class="nip-result"><b>🎉 100% DAS {len(rows)} OBRAS VÁLIDAS ROTEIRIZADAS!</b><br>Organizadas sem modificar a atribuição da coluna K.</div>',unsafe_allow_html=True)
st.caption(f"Tempo de processamento: {plan['elapsed']:.1f}s · Trechos rodoviários OSRM: {plan['real']} · Trechos em linha reta: {plan['fallback']}. A distância total combina quilômetros rodoviários obtidos e distâncias retas de fallback; não é uma estimativa integral de deslocamento por ruas.")
st.markdown('### 🗺️ Mapa Operacional')
if len(rows)>2000:st.info('Mapa limitado aos primeiros 2.000 pontos para desempenho. As exportações incluem todas as obras.')
visible=rows[:2000]
mp=folium.Map(location=[valid['LATITUDE_NUM'].mean(),valid['LONGITUDE_NUM'].mean()],zoom_start=7,tiles='OpenStreetMap',control_scale=True)
folium.TileLayer('CartoDB positron',name='Claro').add_to(mp)
folium.TileLayer('Esri.WorldImagery',name='Satélite',attr='Esri',overlay=False).add_to(mp)
cluster=MarkerCluster(name='Obras').add_to(mp)
colors=['blue','green','purple','orange','red','cadetblue','darkgreen','darkpurple']
for i,(name,group) in enumerate(pd.DataFrame(visible).groupby('COLABORADOR',sort=True)):
    color=colors[i%len(colors)]
    grouped=defaultdict(list)
    for r in group.to_dict('records'):
        grouped[r.get('SUPER_GRUPO',str(r['NOTA']))].append(r)
    for grouped_rows in grouped.values():
        r=grouped_rows[0]
        multi=len(grouped_rows)>1
        card=('<div style="font-family:Arial"><b>📦 SUPER PONTO ('+str(len(grouped_rows))+' obras)</b><hr>'+''.join(popup(item)+'<hr>' for item in grouped_rows)+'</div>') if multi else popup(r)
        folium.Marker([r['LATITUDE_NUM'],r['LONGITUDE_NUM']],popup=folium.Popup(card,max_width=420),tooltip=f"{r['ORDEM']} | {'SUPER PONTO ('+str(len(grouped_rows))+')' if multi else r['NOTA']} | {name}",icon=folium.Icon(color='orange' if multi else color,icon='info-sign')).add_to(cluster)
    if len(rows)<=2000 and len(paths.get(name,[]))>1:
        folium.PolyLine(paths[name],color=color,weight=4,opacity=.85,tooltip=f'Rota de {name} (ver aviso de distâncias)').add_to(mp)
Fullscreen().add_to(mp);folium.LayerControl().add_to(mp)
st_folium(mp,height=550,use_container_width=True,returned_objects=[])
st.markdown('### 📊 Dados Tabulares')
st.dataframe(pd.DataFrame(rows).reindex(columns=['ORDEM']+COLUMNS),use_container_width=True,hide_index=True,height=380)
st.markdown('### 📦 Exportações')
if st.button('Preparar KMZ, KML, GPX, Excel e TXT',type='primary'):
    with st.spinner('Preparando arquivos por colaborador...'):
        st.session_state['nip_export_blob']=export_zip(rows)
        st.session_state['nip_export_key']=key
if st.session_state.get('nip_export_key')==key:
    st.download_button('⬇️ Baixar pacote completo ZIP',st.session_state['nip_export_blob'],file_name='NIP_LISTA_CONTINUA_'+datetime.now().strftime('%Y%m%d')+'.zip',mime='application/zip')
st.caption('O KMZ e GPX mantêm a sequência e os cards, mas o traçado exportado entre pontos é reto. A geometria de ruas OSRM, quando consultada, aparece no mapa operacional; não é exportada como rota rodoviária neste pacote.')
