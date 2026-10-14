import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Observatorio CIS', page_icon='📊', layout='wide')
st.title('📊 Observatorio CIS')
st.caption('Explorador de datos de encuestas · Versión 1.0')
st.info('Esta herramienta calcula frecuencias de filas, sin ponderación ni inferencia estadística. No convierte tablas agregadas en microdatos.')

file = st.file_uploader('Sube un archivo Excel (.xlsx) o CSV (.csv)', type=['xlsx', 'csv'])
if file is None:
    st.write('Para empezar, sube un archivo. Los datos se procesan en la sesión de la aplicación.')
    st.stop()

try:
    if file.name.lower().endswith('.xlsx'):
        book = pd.ExcelFile(file)
        sheet = st.selectbox('Hoja de Excel', book.sheet_names)
        df = pd.read_excel(book, sheet_name=sheet, dtype=str)
    else:
        raw = file.getvalue()
        encoding = st.selectbox('Codificación del CSV', ['utf-8-sig', 'latin-1'])
        sep = st.selectbox('Separador', ['Detectar automáticamente', ';', ',', '\t', '|'])
        df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding=encoding, sep=None if sep == 'Detectar automáticamente' else sep, engine='python' if sep == 'Detectar automáticamente' else 'c')
except Exception as exc:
    st.error(f'No se pudo leer el archivo: {exc}')
    st.stop()

if df.empty or len(df.columns) == 0:
    st.warning('La hoja o archivo no contiene datos utilizables.')
    st.stop()

df.columns = [str(c).strip() for c in df.columns]
# Evitar nombres de columna repetidos.
if df.columns.duplicated().any():
    st.error('Hay columnas con el mismo nombre. Renómbralas antes de continuar.')
    st.stop()

st.subheader('1. Inventario de datos')
a, b = st.columns(2)
a.metric('Filas', f'{len(df):,}')
b.metric('Variables', len(df.columns))
with st.expander('Vista previa y nombres de variables', expanded=True):
    st.dataframe(df.head(30), use_container_width=True)
    inventory = pd.DataFrame({'variable': df.columns, 'valores_no_nulos': [df[c].notna().sum() for c in df.columns], 'categorias_distintas': [df[c].nunique(dropna=True) for c in df.columns]})
    st.dataframe(inventory, use_container_width=True)

st.subheader('2. Filtros')
st.caption('Se aplica un filtro opcional. Puedes cambiarlo para explorar distintos grupos.')
filter_col = st.selectbox('Variable para filtrar', ['(Sin filtro)'] + list(df.columns))
filtered = df.copy()
if filter_col != '(Sin filtro)':
    options = sorted(df[filter_col].dropna().unique().tolist())
    if len(options) > 1000:
        st.warning('Esta variable tiene más de 1.000 valores distintos; selecciona otra para filtrar.')
    else:
        selected = st.multiselect('Valores incluidos', options, default=options)
        filtered = df[df[filter_col].isin(selected)]
st.write(f'Filas seleccionadas: **{len(filtered):,}**')

st.subheader('3. Tabla y gráfica')
variable = st.selectbox('Variable que quieres analizar', df.columns)
include_missing = st.checkbox('Mostrar valores ausentes', value=False)
values = filtered[variable].fillna('(Sin respuesta)') if include_missing else filtered[variable].dropna()
if values.empty:
    st.warning('No hay observaciones con estos filtros.')
    st.stop()
counts = values.value_counts(dropna=False).rename_axis('respuesta').reset_index(name='frecuencia')
counts['porcentaje'] = (counts['frecuencia'] / counts['frecuencia'].sum() * 100).round(2)
counts['respuesta'] = counts['respuesta'].astype(str)
st.dataframe(counts, use_container_width=True, hide_index=True)
if len(counts) <= 100:
    st.bar_chart(counts.set_index('respuesta')['porcentaje'], x_label='Respuesta', y_label='Porcentaje (%)')
else:
    st.warning('Hay más de 100 categorías; se omite la gráfica para mantenerla legible.')
st.download_button('Descargar tabla CSV', counts.to_csv(index=False).encode('utf-8-sig'), file_name='resultados_observatorio_cis.csv', mime='text/csv')
st.caption('Los porcentajes usan como denominador las respuestas incluidas en la tabla. No se han aplicado pesos muestrales. Los códigos numéricos de las respuestas requieren el cuestionario o libro de códigos para interpretarlos.')
