"""Observatorio CIS 7.0: explorador de resultados agregados publicados.
No incluye modelos de estimación privados ni ejecuta contenido de Excel.
"""
import html
import io
import re
import unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Observatorio CIS', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
ROOT = Path(__file__).resolve().parent
DATA_DIR = next((ROOT / n for n in ('Datos','datos') if (ROOT / n).is_dir()), ROOT / 'Datos')
DOC_SHEETS = {'Contenidos','Ficha técnica','Tabla ponderaciones','Notas','Nota estimación','Estimación de Voto'}

def norm(s):
    return ''.join(c for c in unicodedata.normalize('NFD',str(s).lower()) if unicodedata.category(c)!='Mn')

def clean(s):
    return re.sub(r'\s+', ' ', str(s).replace('_x000D_',' ')).strip()

@st.cache_data(max_entries=4, show_spinner='Leyendo el estudio oficial…')
def parse_study(raw: bytes):
    xls = pd.ExcelFile(io.BytesIO(raw), engine='openpyxl')
    output = {}
    for sheet in xls.sheet_names:
        if sheet in DOC_SHEETS:
            continue
        df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        starts = [i for i in range(len(df)) if isinstance(df.iat[i,0],str) and re.match(r'^Pregunta\s+\S+',df.iat[i,0].strip(),re.I)]
        blocks=[]
        for j,start in enumerate(starts):
            end=starts[j+1] if j+1<len(starts) else len(df)
            code=clean(df.iat[start,0]).replace('Pregunta ','').strip()
            title=next((clean(df.iat[k,0]) for k in range(start+1,min(start+6,end)) if isinstance(df.iat[k,0],str) and clean(df.iat[k,0])), '')
            if sheet=='Resultados':
                positions=[1]; names=['TOTAL']; first=start+2
            else:
                head=next((k for k in range(start+1,min(start+12,end)) if any(clean(v).upper()=='TOTAL' for v in df.iloc[k,1:].tolist() if pd.notna(v))),None)
                if head is None: continue
                positions=[c for c in range(1,len(df.columns)) if pd.notna(df.iat[head,c]) and clean(df.iat[head,c])]
                names=[clean(df.iat[head,c]) for c in positions]
                first=head+1
            records=[]; bases={}
            for k in range(first,end):
                label=df.iat[k,0]
                if not isinstance(label,str) or not clean(label):
                    if records: break
                    continue
                label=clean(label)
                if label=='(N)':
                    bases={n:pd.to_numeric(df.iat[k,c],errors='coerce') for c,n in zip(positions,names)}
                    break
                if label.startswith('Pregunta '):break
                row={'Respuesta':label}
                for c,n in zip(positions,names): row[n]=pd.to_numeric(df.iat[k,c],errors='coerce')
                if any(pd.notna(row[n]) for n in names): records.append(row)
            if records:blocks.append({'code':code,'title':title,'table':pd.DataFrame(records),'bases':bases})
        if blocks: output[sheet]=blocks
    return output

# Clasificación editorial por CÓDIGO de pregunta, no por coincidencia de palabras
# que pueden colocar una pregunta en un tema equivocado.
CATALOG = {
 'A1':('Perfil de los entrevistados','Datos personales','Sexo'),
 'A2':('Perfil de los entrevistados','Datos personales','Edad'),
 '0a':('Perfil de los entrevistados','Origen y nacionalidad','Nacionalidad'),
 '0b':('Perfil de los entrevistados','Origen y nacionalidad','Origen de la nacionalidad'),
 '0c':('Perfil de los entrevistados','Origen y nacionalidad','País de nacimiento'),
 '0d':('Perfil de los entrevistados','Origen y nacionalidad','País de origen (detalle)'),
 '1':('Vivienda','Caso Maricarmen','Conocimiento del desahucio'),
 '2':('Vivienda','Caso Maricarmen','Importancia del caso'),
 '3':('Vivienda','Caso Maricarmen','Prioridades ante un desahucio'),
 '4':('Vivienda','Medidas del Gobierno','Conocimiento de los decretos'),
 '4a':('Vivienda','Medidas del Gobierno','Valoración de la medida'),
 '5':('Vivienda','Medidas del Gobierno','Apoyo a los decretos'),
 '6a':('Vivienda','Competencias y soluciones','Administración competente'),
 '6b':('Vivienda','Competencias y soluciones','Quién debería gestionar la vivienda'),
 '7':('Vivienda','Competencias y soluciones','Prioridades de política de vivienda'),
 '8':('Medio ambiente y política exterior','Medio ambiente','Preocupación por el cambio climático'),
 '9':('Medio ambiente y política exterior','Conflictos internacionales','Preocupación por Ucrania'),
 '10':('Medio ambiente y política exterior','Conflictos internacionales','Preocupación por Oriente Próximo'),
 '11':('Economía','Situación económica','Economía personal'),
 '12':('Economía','Situación económica','Economía española'),
 '13R':('Problemas y preocupaciones','Problemas del país','Principales problemas de España'),
 '14R':('Problemas y preocupaciones','Problemas personales','Problemas que más afectan'),
 '15R':('Líderes políticos','Preferencia de presidente','Presidente preferido'),
 '15Ra':('Líderes políticos','Preferencia de presidente','Presidente preferido (voto definido)'),
 '16':('Elecciones y voto','Participación electoral','Probabilidad de votar'),
 '17R':('Elecciones y voto','Intención de voto','Intención directa de voto'),
 '18R':('Elecciones y voto','Intención de voto','Segunda opción de voto'),
 '19R':('Elecciones y voto','Simpatía política','Simpatía por partidos'),
 '19a':('Elecciones y voto','Simpatía política','Voto más simpatía'),
 '19aR':('Elecciones y voto','Simpatía política','Voto más simpatía (voto definido)'),
 '20':('Líderes políticos','Conocimiento y valoración','Conocimiento de líderes'),
 '20a':('Líderes políticos','Conocimiento y valoración','Valoración de líderes (1–10)'),
 '21':('Líderes políticos','Confianza','Confianza en Pedro Sánchez'),
 '22':('Líderes políticos','Confianza','Confianza en Alberto Núñez Feijóo'),
 '23':('Líderes políticos','Gobierno y ministros','Conocimiento de ministros'),
 '23a':('Líderes políticos','Gobierno y ministros','Valoración de ministros'),
 '24':('Elecciones y voto','Recuerdo electoral','Participación en 2023'),
 '24aR':('Elecciones y voto','Recuerdo electoral','Voto en las generales de 2023'),
 '24aa':('Elecciones y voto','Recuerdo electoral','Participación y voto en 2023'),
 '25':('Elecciones y voto','Fidelidad y cambio de voto','¿Repetiría su voto?'),
 '26R':('Elecciones y voto','Identificación partidista','Partido más cercano'),
 '26Ra':('Elecciones y voto','Identificación partidista','Partido más cercano (voto definido)'),
 '27':('Elecciones y voto','Fidelidad y cambio de voto','Fidelidad a un partido'),
 '28':('Elecciones y voto','Fidelidad y cambio de voto','Cuándo decide su voto'),
 '29R':('Elecciones y voto','Intención de voto','Intención de voto autonómico'),
 '30':('Ideología','Autoubicación ideológica','Posición izquierda–derecha'),
 '31':('Perfil de los entrevistados','Educación','Escolarización'),
 '31a':('Perfil de los entrevistados','Educación','Nivel educativo alcanzado'),
 '31aa':('Perfil de los entrevistados','Educación','Nivel de estudios (agrupado)'),
 '32':('Sociedad y valores','Religión','Identificación religiosa'),
 '32a':('Sociedad y valores','Religión','Práctica religiosa'),
 '33':('Perfil de los entrevistados','Datos personales','Estado civil'),
 '34':('Trabajo y clase social','Empleo','Situación laboral'),
 '34a':('Trabajo y clase social','Empleo','Tipo de trabajo'),
 '34b':('Trabajo y clase social','Empleo','Ocupación'),
 '34ba':('Trabajo y clase social','Empleo','Población activa e inactiva'),
 '35':('Economía','Ingresos','Ingresos del hogar'),
 '36':('Trabajo y clase social','Clase social','Clase social declarada'),
 '36a':('Trabajo y clase social','Clase social','Clase social (agrupada)'),
}

def meta(block):
    return CATALOG.get(block['code'],('Otros temas','Otras preguntas', clean(block['title'])[:60] or block['code']))

PALETTE=['#3563e9','#dd8080','#97a3b7','#22a699','#e8b455','#9c7ad7','#6c9bd1','#d26f9d','#74888b','#d4b58a']
PARTY_LIST=['PSOE','PP','VOX','Sumar','Podemos','ERC','Junts','EH Bildu','PNV']

def e(x): return html.escape(str(x),quote=True)

def pct(x):return f'{float(x):.1f}'.replace('.',',')+' %'

def bar_single(labels,values):
    parts=[]
    for i,(label,v) in enumerate(zip(labels,values)):
        if pd.isna(v):continue
        v=float(v)
        width=max(0,min(100,v))
        parts.append(f'<div class="barrow"><div class="barlabel"><span>{e(label)}</span><b>{pct(v)}</b></div><div class="track"><div class="fill" style="width:{width:.3f}%;background:{PALETTE[i%len(PALETTE)]}"></div></div></div>')
    return '<div class="viz">'+''.join(parts)+'</div>'

def bar_stacked(view,groups,answers):
    parts=[]
    for group in groups:
        vals=[pd.to_numeric(view.loc[view['Respuesta']==a,group],errors='coerce').iloc[0] for a in answers]
        total=sum(float(v) for v in vals if pd.notna(v) and v>=0)
        # Never silently renormalize published values. Show raw segments only if compatible.
        if not 98<=total<=102:
            parts.append(f'<div class="barrow"><div class="barlabel"><span>{e(group)}</span><small>No suma 100 % ({pct(total)}). Usa barras individuales.</small></div></div>')
            continue
        parts.append(f'<div class="barrow"><div class="barlabel"><span>{e(group)}</span><b>{pct(total)}</b></div><div class="stack">')
        for i,v in enumerate(vals):
            if pd.isna(v) or v<=0:continue
            w=float(v)
            txt=pct(w) if w>=12 else ''
            parts.append(f'<div class="segment" style="width:{w:.4f}%;background:{PALETTE[i%len(PALETTE)]}" title="{e(answers[i])}: {pct(w)}">{e(txt)}</div>')
        parts.append('</div></div>')
    legend=''.join(f'<span class="legenditem"><i style="background:{PALETTE[i%len(PALETTE)]}"></i>{e(a)}</span>' for i,a in enumerate(answers))
    return '<div class="viz">'+''.join(parts)+f'<div class="legend">{legend}</div></div>'

def heatmap(view,groups,answers):
    values=[]
    for a in answers:
        r=view[view['Respuesta']==a]
        if r.empty:continue
        values.append((a,[pd.to_numeric(r[g],errors='coerce').iloc[0] for g in groups]))
    maxv=max([float(v) for _,row in values for v in row if pd.notna(v)] or [100])
    maxv=max(maxv,1)
    head='<div class="heatgrid" style="--cols:'+str(len(answers))+'"><div class="heathead"></div>'+''.join('<div class="heathead">'+e(a)+'</div>' for a in answers)
    rows=[]
    for group in groups:
        rows.append('<div class="heatrowlabel">'+e(group)+'</div>')
        for a in answers:
            r=view[view['Respuesta']==a]
            v=pd.to_numeric(r[group],errors='coerce').iloc[0] if not r.empty else float('nan')
            if pd.isna(v):rows.append('<div class="heatcell empty">—</div>');continue
            alpha=.14+.78*max(0,min(1,float(v)/maxv))
            rows.append(f'<div class="heatcell" style="background:rgba(53,99,233,{alpha:.3f})" title="{e(group)} · {e(a)}: {pct(v)}">{pct(v)}</div>')
    return '<div class="heatwrap">'+head+''.join(rows)+'</div></div>'

st.markdown('''<style>
.landing{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin:22px 0}.landingcard{border:1px solid rgba(145,158,180,.28);border-radius:16px;padding:20px;display:flex;flex-direction:column;gap:8px}.landingcard span{color:#9ca3af;font-size:.85rem}.block-container{max-width:1020px;padding-top:1.4rem;padding-bottom:4rem}
h1{letter-spacing:-.045em}h2,h3{letter-spacing:-.025em}
[data-testid="stMetric"]{border:1px solid rgba(130,145,170,.2);border-radius:14px;padding:12px}
.viz{border:1px solid rgba(145,158,180,.24);border-radius:18px;padding:20px;margin:10px 0 18px}
.barrow{margin:0 0 20px}.barrow:last-child{margin-bottom:0}
.barlabel{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:7px;font-size:1rem}
.barlabel b{white-space:nowrap;font-size:1.05rem}.barlabel small{font-size:.75rem;color:#9ca3af}
.track{height:15px;background:rgba(150,160,175,.20);border-radius:7px;overflow:hidden}
.fill{height:100%;border-radius:7px}.stack{height:38px;display:flex;border-radius:8px;overflow:hidden;background:rgba(150,160,175,.13)}
.segment{display:flex;align-items:center;justify-content:center;color:white;font-size:.83rem;font-weight:700;white-space:nowrap;overflow:hidden}
.legend{display:flex;gap:10px 18px;flex-wrap:wrap;margin-top:18px;font-size:.8rem}.legenditem{display:flex;align-items:center;gap:6px}
.legenditem i{display:inline-block;width:11px;height:11px;border-radius:3px}
.heatwrap{overflow-x:auto;max-width:100%;padding-bottom:10px}.heatgrid{display:grid;grid-template-columns:110px repeat(var(--cols),minmax(90px,1fr));gap:5px;min-width:max-content}
.heathead{text-align:center;font-size:.76rem;font-weight:700;padding:8px 4px;max-width:135px;overflow-wrap:anywhere}
.heatrowlabel{display:flex;align-items:center;font-size:.84rem;padding:6px;max-width:110px;overflow-wrap:anywhere}
.heatcell{min-height:51px;display:flex;align-items:center;justify-content:center;border-radius:7px;color:#fff;font-size:.85rem;font-weight:700}
.heatcell.empty{background:rgba(150,160,175,.12);color:#aaa}
@media(max-width:640px){.block-container{padding-left:1rem;padding-right:1rem}.viz{padding:14px}.stack{height:34px}.barlabel{font-size:.92rem}.heatgrid{grid-template-columns:90px repeat(var(--cols),minmax(85px,1fr))}}
</style>''',unsafe_allow_html=True)

st.title('📊 Observatorio CIS')
st.caption('Datos oficiales · Panel temático · Versión 7.0')
if not DATA_DIR.exists():
    st.error('Falta la carpeta Datos junto a app.py.');st.stop()
files=sorted(p for p in DATA_DIR.glob('*.xlsx') if p.is_file() and not p.is_symlink() and not p.name.startswith('~$') and p.stat().st_size<=30*1024*1024)
if not files:
    st.error('No se encontraron estudios Excel en Datos/.');st.stop()
with st.sidebar:
    st.markdown('### Estudios')
    file=st.selectbox('Barómetro',files,format_func=lambda p:p.stem)
    st.caption('Los datos del repositorio privado se muestran solo a usuarios autorizados de la aplicación.')
try:
    data=parse_study(file.read_bytes())
except Exception:
    st.error('No se pudo interpretar el Excel. Consulta los registros de Streamlit.');st.stop()
if not data:
    st.warning('No hay preguntas reconocidas.');st.stop()
reference=data.get('Resultados') or next(iter(data.values()))

PAGES=[
    ('🏠 Vivienda','Vivienda'),('🗳️ Elecciones','Elecciones y voto'),
    ('👤 Líderes','Líderes políticos'),('💶 Economía','Economía'),
    ('⚖️ Ideología','Ideología'),('📋 Problemas','Problemas y preocupaciones'),
    ('👥 Sociedad','Sociedad y valores'),('💼 Trabajo','Trabajo y clase social'),
    ('🎓 Perfil','Perfil de los entrevistados'),('🌍 Internacional','Medio ambiente y política exterior'),
]

# Each topic is a genuine navigation page in the sidebar rather than a topic selector.
with st.sidebar:
    st.markdown('### Páginas')
    choices=['📊 Gráficos','🔄 Trasvases','🌡️ Mapas de calor','🏡 Inicio']+[label for label,topic in PAGES if any(meta(b)[0]==topic for b in reference)]+['📚 Estudios']
    page=st.radio('Navegación',choices,label_visibility='collapsed',key='navpage')

if page=='🏡 Inicio':
    st.header('Explora el CIS por temas')
    st.write('Elige una página en el menú lateral ☰. Dentro encontrarás subtemas e indicadores con nombres breves. La pregunta oficial está disponible al desplegar su ficha.')
    st.markdown('<div class="landing">'+''.join(f'<div class="landingcard"><b>{e(label)}</b><span>{sum(meta(b)[0]==topic for b in reference)} indicadores</span></div>' for label,topic in PAGES if any(meta(b)[0]==topic for b in reference))+'</div>',unsafe_allow_html=True)
    st.info('Los gráficos usan los porcentajes publicados. No se infieren entrevistas individuales ni se generan estimaciones propias.')
    st.stop()

if page=='📚 Estudios':
    st.header('Estudios disponibles')
    for f in files:st.write('📄',f.stem)
    st.write('Hojas reconocidas:',', '.join(data))
    st.info('Para añadir un estudio, sube otro .xlsx a Datos/ en GitHub. No incluyas modelos privados ni claves en el repositorio.')
    st.stop()

if page=='🔄 Trasvases':
    st.header('Trasvases electorales')
    st.caption('Intención declarada por recuerdo de voto de 2023. Son porcentajes condicionados publicados, no flujos absolutos de personas.')
    b=next((b for b in data.get('RV EG23',[]) if b['code']=='17R'),None)
    if b is None:
        st.warning('Este estudio no tiene la tabla cruzada necesaria.');st.stop()
    available=[c for c in b['table'].columns if c not in ('Respuesta','TOTAL')]
    origins=st.multiselect('Voto en 2023',available,default=[x for x in ['PSOE','PP','VOX','Sumar'] if x in available])
    destinations=b['table']['Respuesta'].astype(str).tolist()
    dests=st.multiselect('Intención actual',destinations,default=[x for x in destinations if norm(x) in ('psoe','pp','vox','sumar')])
    if origins and dests:
        st.markdown(heatmap(b['table'],origins,dests),unsafe_allow_html=True)
        st.download_button('Descargar matriz CSV',b['table'][b['table']['Respuesta'].isin(dests)][['Respuesta']+origins].to_csv(index=False).encode('utf-8-sig'),'trasvases.csv','text/csv')
    st.stop()

# Visualización por tipo: estas son las tres páginas principales del menú.
if page in ('📊 Gráficos','🌡️ Mapas de calor'):
    st.header(page)
    st.caption('Selecciona un tema, un subtema y un indicador. La pregunta original se muestra aparte.')
    available_topics=[(label,topic) for label,topic in PAGES if any(meta(b)[0]==topic for b in reference)]
    labels=[label for label,_ in available_topics]
    chosen_topic=st.selectbox('Tema',labels,key='viz_topic')
    topic_name=dict(available_topics)[chosen_topic]
    items=[b for b in reference if meta(b)[0]==topic_name]
    sub=st.selectbox('Subtema',sorted({meta(b)[1] for b in items}),key='viz_sub')
    items=[b for b in items if meta(b)[1]==sub]
    codes=[b['code'] for b in items]
    qcode=st.selectbox('Indicador',codes,format_func=lambda code:next((meta(b)[2] for b in items if b['code']==code),code),key='viz_indicator')
    selected=next(b for b in items if b['code']==qcode)
    st.subheader(meta(selected)[2])
    with st.expander('Leer pregunta original del CIS'):
        st.write(selected['title'] or 'Enunciado no disponible en el Excel')
        st.caption('Código '+qcode+' · Estudio '+file.stem)
    MODES_VIZ={'Total':'Resultados','Por partido (voto en 2023)':'RV EG23','Por edad':'Edad','Por sexo':'Sexo','Por estudios':'Nivel de estudios','Por ideología':'Escala de ideología','Por clase social':'Clase social','Por ocupación':'Ocupación','Por religiosidad':'Religiosidad','Por tamaño de municipio':'Tamaño de municipio'}
    valid={label:sheet for label,sheet in MODES_VIZ.items() if sheet in data and any(b['code']==qcode for b in data[sheet])}
    if not valid:
        st.info('No existe un desglose publicado para este indicador.');st.stop()
    mode=st.selectbox('Comparar por',list(valid),index=0 if len(valid)==1 else 1,key='viz_mode')
    b=next(b for b in data[valid[mode]] if b['code']==qcode)
    columns=[c for c in b['table'].columns if c!='Respuesta']
    if valid[mode]=='RV EG23':
        party=st.selectbox('Partido o bloque',['Todos','Izquierda (partidos)','Derecha (partidos)']+[c for c in PARTY_LIST if c in columns],key='viz_party')
        if party=='Todos': defaults=[c for c in PARTY_LIST if c in columns][:5]
        elif party.startswith('Izquierda'):defaults=[c for c in ['PSOE','Sumar','Podemos','ERC','EH Bildu'] if c in columns]
        elif party.startswith('Derecha'):defaults=[c for c in ['PP','VOX'] if c in columns]
        else:defaults=[party]
        st.caption('Los bloques son selecciones de partidos, no sumas estadísticas de porcentajes.')
    else:defaults=columns[:min(5,len(columns))]
    groups=st.multiselect('Grupos que quieres mostrar',columns,default=defaults,key='viz_groups')
    if not groups:
        st.info('Selecciona al menos un grupo.');st.stop()
    view=b['table'][['Respuesta']+groups].copy()
    answers=view['Respuesta'].astype(str).tolist()
    if not answers:
        st.info('No hay respuestas disponibles.');st.stop()
    if page=='📊 Gráficos':
        chart=st.radio('Tipo de gráfico',['Barras horizontales','Barras apiladas'],horizontal=True,key='viz_chart')
        if chart=='Barras horizontales':
            answer=st.selectbox('Respuesta que quieres comparar',answers,key='viz_answer')
            row=view.loc[view['Respuesta'].astype(str)==answer].iloc[0]
            st.markdown(bar_single(groups,[row[g] for g in groups]),unsafe_allow_html=True)
        else:
            chosen=st.multiselect('Respuestas que quieres representar',answers,default=answers[:min(5,len(answers))],key='viz_answers')
            if chosen:st.markdown(bar_stacked(view,groups,chosen),unsafe_allow_html=True)
            st.caption('Las barras apiladas solo representan el 100 % cuando las categorías son exhaustivas y excluyentes.')
    else:
        chosen=st.multiselect('Respuestas (columnas del mapa)',answers,default=answers[:min(4,len(answers))],key='viz_heat_answers')
        if chosen:st.markdown(heatmap(view,groups,chosen),unsafe_allow_html=True)
        st.caption('Desliza horizontalmente para ver todas las columnas en el móvil.')
    st.download_button('⬇️ Descargar datos CSV',view.to_csv(index=False).encode('utf-8-sig'),f'cis_{file.stem}_{qcode}.csv','text/csv',key='viz_csv')
    st.stop()

topic=dict(PAGES)[page]
st.header(page)
st.caption('Estudio '+file.stem+' · Resultados oficiales publicados')
items=[b for b in reference if meta(b)[0]==topic]
if not items:
    st.warning('No hay indicadores disponibles en este tema.');st.stop()
subtopics=sorted({meta(b)[1] for b in items})
sub=st.selectbox('Subtema',subtopics)
items=[b for b in items if meta(b)[1]==sub]
qcode=st.selectbox('Indicador',[b['code'] for b in items],format_func=lambda code:next((meta(b)[2] for b in items if b['code']==code),code))
selected=next(b for b in items if b['code']==qcode)
st.subheader(meta(selected)[2])
with st.expander('Ver la pregunta original del CIS'):
    st.write(selected['title'] or 'Enunciado no disponible en el Excel')
    st.caption('Código '+qcode+' · Estudio '+file.stem)

MODES={'Total':'Resultados','Por partido (voto en 2023)':'RV EG23','Por edad':'Edad','Por sexo':'Sexo','Por estudios':'Nivel de estudios','Por ideología':'Escala de ideología','Por clase social':'Clase social','Por ocupación':'Ocupación','Por religiosidad':'Religiosidad','Por tamaño de municipio':'Tamaño de municipio'}
valid={label:sheet for label,sheet in MODES.items() if sheet in data and any(b['code']==qcode for b in data[sheet])}
if not valid:
    st.warning('El Excel no ofrece un desglose para este indicador.');st.stop()
mode=st.selectbox('Comparar por',list(valid),index=0)
sheet=valid[mode]
b=next(b for b in data[sheet] if b['code']==qcode)
cols=[c for c in b['table'].columns if c!='Respuesta']
if sheet=='RV EG23':
    party_options=['Todos','Izquierda (partidos)','Derecha (partidos)']+[c for c in PARTY_LIST if c in cols]
    party=st.selectbox('Partido / bloque',party_options)
    if party=='Todos':default=[c for c in PARTY_LIST if c in cols][:5]
    elif party.startswith('Izquierda'):default=[c for c in ['PSOE','Sumar','Podemos','ERC','EH Bildu'] if c in cols]
    elif party.startswith('Derecha'):default=[c for c in ['PP','VOX'] if c in cols]
    else:default=[party]
    st.caption('Izquierda y derecha son selecciones de partidos para comparar, no porcentajes agregados.')
else:default=cols[:min(5,len(cols))]
groups=st.multiselect('Grupos visibles',cols,default=default)
if not groups:
    st.info('Selecciona al menos un grupo.');st.stop()
view=b['table'][['Respuesta']+groups].copy()
answers=view['Respuesta'].astype(str).tolist()
if not answers:
    st.info('No hay respuestas en esta pregunta.');st.stop()
chart=st.segmented_control('Visualización',['Barras','Apiladas','Mapa de calor','Tabla'],default='Barras')
if chart=='Barras':
    answer=st.selectbox('Respuesta',answers)
    row=view.loc[view['Respuesta'].astype(str)==answer].iloc[0]
    st.markdown(bar_single(groups,[row[g] for g in groups]),unsafe_allow_html=True)
elif chart=='Apiladas':
    chosen=st.multiselect('Respuestas',answers,default=answers[:min(5,len(answers))])
    if chosen:st.markdown(bar_stacked(view,groups,chosen),unsafe_allow_html=True)
    st.caption('Solo se apilan categorías que sumen aproximadamente el 100 %; no se modifican los porcentajes oficiales.')
elif chart=='Mapa de calor':
    chosen=st.multiselect('Respuestas',answers,default=answers[:min(4,len(answers))])
    if chosen:st.markdown(heatmap(view,groups,chosen),unsafe_allow_html=True)
else:
    st.dataframe(view,use_container_width=True,hide_index=True)
st.download_button('⬇️ Descargar tabla CSV',view.to_csv(index=False).encode('utf-8-sig'),f'cis_{file.stem}_{qcode}.csv','text/csv')
with st.expander('Notas metodológicas'):
    st.write('Los valores son los porcentajes publicados por el CIS. Las respuestas múltiples pueden sumar más del 100 %. No se calculan estimaciones propias ni pruebas estadísticas con estas tablas agregadas. Los nombres breves son etiquetas editoriales: consulta el enunciado original para interpretar cada indicador.')
