import io
from pathlib import Path
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Observatorio CIS', page_icon='📊', layout='wide')
st.title('📊 Observatorio CIS')
st.caption('Explorador de resultados publicados del CIS · Versión 3.0')
st.info('Los Excel de resultados agregados contienen porcentajes publicados, no entrevistas individuales. No se calculan frecuencias de filas ni significación estadística a partir de estas tablas.')

# Solo se cargan estudios incluidos por el administrador en el repositorio.
# Los visitantes no pueden subir archivos arbitrarios.
DATA_DIR = Path(__file__).resolve().parent / "Datos"
studies = sorted(p for p in DATA_DIR.glob("*.xlsx") if p.is_file() and not p.is_symlink() and p.stat().st_size <= 30 * 1024 * 1024)
if not studies:
    st.error("No hay estudios publicados. Añade archivos .xlsx a la carpeta Datos/ del repositorio.")
    st.stop()
study = st.selectbox("Selecciona un estudio del CIS", studies, format_func=lambda p: p.stem)
raw = study.read_bytes()


@st.cache_data(show_spinner='Leyendo el archivo…')
def load_xlsx(data):
    return pd.ExcelFile(io.BytesIO(data))

@st.cache_data(show_spinner='Extrayendo preguntas…')
def get_blocks(data, sheet):
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object)
    blocks=[]
    starts=[]
    for i in range(len(df)):
        x=df.iat[i,0]
        if isinstance(x,str) and re.match(r'^Pregunta\s+\S+',x.strip(),re.I):
            starts.append(i)
    for j,start in enumerate(starts):
        end=starts[j+1] if j+1<len(starts) else len(df)
        seg=df.iloc[start:end].reset_index(drop=True)
        code=str(seg.iat[0,0]).strip()
        title=''
        for k in range(1,min(7,len(seg))):
            v=seg.iat[k,0]
            if isinstance(v,str) and v.strip() and v.strip()!='(N)':
                title=v.strip().replace('_x000D_',' ')
                break
        header=None
        for k in range(1,min(12,len(seg))):
            if any(str(x).strip().upper()=='TOTAL' for x in seg.iloc[k,1:].tolist() if pd.notna(x)):
                header=k
                break
        if header is None and sheet=='Resultados':
            header=1
            while header<len(seg) and pd.notna(seg.iat[header,0]):
                header+=1
            header-=1
        if header is None:
            continue
        cols=[]
        if sheet=='Resultados':
            cols=['Respuesta','TOTAL']
            start_rows=1
            while start_rows<len(seg) and (pd.isna(seg.iat[start_rows,0]) or not pd.isna(seg.iat[start_rows,1])):
                start_rows+=1
            # The first row after the question text contains the first response.
            start_rows=next((k for k in range(1,len(seg)) if pd.notna(seg.iat[k,0]) and pd.notna(seg.iat[k,1])),len(seg))
            indexes=[0,1]
        else:
            indexes=[0]+[i for i in range(1,len(seg.columns)) if pd.notna(seg.iat[header,i]) and str(seg.iat[header,i]).strip()]
            cols=['Respuesta']+[str(seg.iat[header,i]).strip() for i in indexes[1:]]
            start_rows=header+1
        records=[]; nvals={}
        for k in range(start_rows,len(seg)):
            label=seg.iat[k,0]
            if not isinstance(label,str) or not label.strip():
                if records:break
                continue
            label=label.strip()
            if label=='(N)':
                for ci,c in zip(indexes[1:],cols[1:]):
                    nvals[c]=pd.to_numeric(seg.iat[k,ci],errors='coerce')
                break
            if label.startswith('Pregunta '):break
            record={'Respuesta':label}
            for ci,c in zip(indexes[1:],cols[1:]):
                record[c]=pd.to_numeric(seg.iat[k,ci],errors='coerce')
            if any(pd.notna(record[c]) for c in cols[1:]):records.append(record)
        if records:
            blocks.append({'code':code,'title':title,'table':pd.DataFrame(records),'n':nvals})
    return blocks

book=load_xlsx(raw)
st.success(f'Estudio disponible: {study.stem} · {len(book.sheet_names)} hojas.')
st.subheader('1. Selecciona un desglose')
sheet=st.selectbox('Hoja del CIS',book.sheet_names,index=book.sheet_names.index('RV EG23') if 'RV EG23' in book.sheet_names else 0)
if sheet in ['Contenidos','Ficha técnica','Tabla ponderaciones','Notas','Nota estimación','Estimación de Voto']:
    st.warning('Esta hoja es documentación o una tabla especial; se muestra sin calcular porcentajes adicionales.')
    df=pd.read_excel(io.BytesIO(raw),sheet_name=sheet,header=None)
    st.dataframe(df.fillna(''),use_container_width=True,height=500)
    st.stop()

blocks=get_blocks(raw,sheet)
if not blocks:
    st.warning('No se han reconocido preguntas en esta hoja. Vista original:')
    st.dataframe(pd.read_excel(io.BytesIO(raw),sheet_name=sheet,header=None),use_container_width=True)
    st.stop()
st.write(f'Preguntas o bloques reconocidos: **{len(blocks)}**')
st.subheader('2. Busca y selecciona una pregunta')
search=st.text_input('Buscar por código o texto (por ejemplo: vivienda, alquiler, problema, voto)')
options=[i for i,b in enumerate(blocks) if search.casefold() in (b['code']+' '+b['title']).casefold()]
if not options:
    st.warning('No hay coincidencias. Prueba otro término.')
    st.stop()
idx=st.selectbox('Pregunta',options,format_func=lambda i:f"{blocks[i]['code']} — {blocks[i]['title'][:145]}")
b=blocks[idx]
st.markdown(f"**{b['code']}** · {b['title']}")
table=b['table'].copy()
columns=[c for c in table.columns if c!='Respuesta']
if len(columns)>1:
    defaults=[c for c in ['TOTAL','PSOE','PP','VOX','Sumar'] if c in columns]
    selected=st.multiselect('Columnas/grupos que quieres comparar',columns,default=defaults or columns[:5])
else:selected=columns
if not selected:st.stop()
view=table[['Respuesta']+selected].copy()
for c in selected:view[c]=pd.to_numeric(view[c],errors='coerce').round(2)
st.subheader('3. Porcentajes publicados')
st.dataframe(view,use_container_width=True,hide_index=True)
if b['n']:
    st.caption('Bases (N) publicadas, cuando están disponibles: '+', '.join(f'{c}: {b["n"][c]:.1f}' for c in selected if c in b['n'] and pd.notna(b['n'][c])))
st.subheader('4. Gráfica comparativa')
chart=view.set_index('Respuesta')[selected].dropna(how='all')
if len(chart)>0:st.bar_chart(chart)
st.download_button('Descargar esta tabla en CSV',view.to_csv(index=False).encode('utf-8-sig'),file_name=f'cis_{sheet}_{b["code"].replace(" ","_")}.csv',mime='text/csv')
st.caption('Fuente: archivo oficial del CIS incorporado por el administrador. Los valores son porcentajes publicados en la tabla, no estimaciones recalculadas. Las bases pueden ser ponderadas y las preguntas de respuesta múltiple no tienen por qué sumar 100 %.')
