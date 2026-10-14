"""Observatorio CIS 4.0 — exploration of official published aggregate tables.

No credentials, user uploads or proprietary estimation models are included.
"""
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Observatorio CIS", page_icon="📊", layout="wide")
DATA_DIR = Path(__file__).resolve().parent / "Datos"
if not DATA_DIR.exists():
    DATA_DIR = Path(__file__).resolve().parent / "datos"
DOC_SHEETS = {"Contenidos", "Ficha técnica", "Tabla ponderaciones", "Notas", "Nota estimación", "Estimación de Voto"}
COLORS = px.colors.qualitative.Safe

@st.cache_data(show_spinner="Leyendo tablas oficiales…", max_entries=6)
def parse_study(data: bytes):
    """Cache only serializable pandas frames/dicts, never ExcelFile objects."""
    source = io.BytesIO(data)
    xls = pd.ExcelFile(source)
    result = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        if sheet in DOC_SHEETS:
            result[sheet] = {"blocks": [], "raw": df}
            continue
        starts = [i for i in range(len(df)) if isinstance(df.iat[i, 0], str)
                  and re.match(r"^Pregunta\s+\S+", df.iat[i, 0].strip(), re.I)]
        blocks = []
        for j, start in enumerate(starts):
            end = starts[j + 1] if j + 1 < len(starts) else len(df)
            seg = df.iloc[start:end]
            code = str(df.iat[start, 0]).strip()
            title = ""
            for k in range(start + 1, min(start + 5, end)):
                v = df.iat[k, 0]
                if isinstance(v, str) and v.strip():
                    title = v.strip().replace("_x000D_", " ")
                    break
            if sheet == "Resultados":
                header = None
                col_positions = [1]
                col_names = ["TOTAL"]
                first = start + 2
            else:
                header = next((k for k in range(start + 1, min(start + 12, end))
                               if any(str(v).strip().upper() == "TOTAL" for v in df.iloc[k, 1:].tolist()
                                      if pd.notna(v))), None)
                if header is None:
                    continue
                col_positions = [c for c in range(1, len(df.columns))
                                 if pd.notna(df.iat[header, c]) and str(df.iat[header, c]).strip()]
                col_names = [str(df.iat[header, c]).strip() for c in col_positions]
                first = header + 1
            records, bases = [], {}
            for k in range(first, end):
                label = df.iat[k, 0]
                if not isinstance(label, str) or not label.strip():
                    if records:
                        break
                    continue
                label = label.strip()
                if label == "(N)":
                    bases = {name: pd.to_numeric(df.iat[k, c], errors="coerce")
                             for c, name in zip(col_positions, col_names)}
                    break
                if label.startswith("Pregunta "):
                    break
                record = {"Respuesta": label}
                for c, name in zip(col_positions, col_names):
                    record[name] = pd.to_numeric(df.iat[k, c], errors="coerce")
                if any(pd.notna(record[n]) for n in col_names):
                    records.append(record)
            if records:
                blocks.append({"code": code, "title": title,
                               "table": pd.DataFrame(records), "bases": bases})
        result[sheet] = {"blocks": blocks, "raw": None}
    return result


def find_question(blocks, term):
    term = term.casefold().strip()
    return [i for i, b in enumerate(blocks) if term in (b["code"] + " " + b["title"]).casefold()]


def selector(study_data, key):
    sheets = [s for s, value in study_data.items() if value["blocks"]]
    if not sheets:
        st.warning("No se han reconocido preguntas en este archivo.")
        st.stop()
    default = sheets.index("RV EG23") if "RV EG23" in sheets else 0
    sheet = st.selectbox("Desglose / hoja", sheets, index=default, key=f"{key}_sheet")
    blocks = study_data[sheet]["blocks"]
    query = st.text_input("Buscar pregunta (vivienda, voto, problemas, confianza…)", key=f"{key}_query")
    hits = find_question(blocks, query)
    if not hits:
        st.info("No hay coincidencias para esa búsqueda.")
        st.stop()
    chosen = st.selectbox("Pregunta", hits,
                          format_func=lambda i: f'{blocks[i]["code"]} — {blocks[i]["title"][:160]}',
                          key=f"{key}_question")
    b = blocks[chosen]
    st.caption(f'{b["code"]} · {b["title"]}')
    return sheet, b


def group_selector(b, key, limit=6):
    cols = [c for c in b["table"].columns if c != "Respuesta"]
    defaults = [c for c in ["TOTAL", "PSOE", "PP", "VOX", "Sumar"] if c in cols]
    return st.multiselect("Grupos que quieres comparar", cols, default=(defaults or cols[:min(limit, len(cols))]), key=key)


def numeric_view(b, cols):
    v = b["table"][["Respuesta"] + cols].copy()
    for c in cols:
        v[c] = pd.to_numeric(v[c], errors="coerce")
    return v


def plot_stacked(view, groups, question, normalize):
    labels = view["Respuesta"].astype(str).tolist()
    vals = view[groups].fillna(0).clip(lower=0)
    totals = vals.sum(axis=0)
    if normalize:
        # Normalize only if response categories plausibly form a single-choice partition.
        if not all(97 <= float(t) <= 103 for t in totals if t > 0):
            st.warning("Las respuestas no suman aproximadamente 100 %. Puede ser una pregunta múltiple o una tabla con otra base. No se normaliza automáticamente.")
            normalize = False
    fig = go.Figure()
    for i, label in enumerate(labels):
        y = [(float(vals.at[i, g]) / float(totals[g]) * 100 if normalize and totals[g] else float(vals.at[i, g])) for g in groups]
        fig.add_trace(go.Bar(name=label[:75], y=groups, x=y, orientation="h",
                             marker_color=COLORS[i % len(COLORS)],
                             hovertemplate="%{y}: %{x:.1f}%<extra>" + label[:75] + "</extra>"))
    fig.update_layout(barmode="stack", xaxis_title="Porcentaje publicado (%)", yaxis_title="",
                      legend_title="Respuesta", height=max(360, 75 * len(groups) + 140),
                      margin=dict(l=10, r=10, t=30, b=30))
    if normalize:
        fig.update_xaxes(range=[0, 100])
    st.plotly_chart(fig, use_container_width=True)


def plot_heatmap(view, groups):
    m = view.set_index("Respuesta")[groups].T
    m = m.loc[:, m.notna().any(axis=0)]
    if m.empty:
        st.warning("No hay porcentajes numéricos para mostrar.")
        return
    fig = px.imshow(m, aspect="auto", color_continuous_scale="Blues", text_auto=".1f",
                    labels={"x": "Respuesta", "y": "Grupo", "color": "%"})
    fig.update_layout(height=max(360, len(groups) * 65 + 180),
                      margin=dict(l=10, r=10, t=25, b=10))
    st.plotly_chart(fig, use_container_width=True)


def csv_button(df, name):
    st.download_button("⬇️ Descargar tabla CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=name, mime="text/csv")


# Navegación guiada por temas. La clasificación es orientativa y revisable.
TOPICS = {
    "Todos los temas": [],
    "Vivienda": ["viviend", "alquiler", "hipotec", "desahuci", "precio de la casa"],
    "Intención de voto y política": ["voto", "votar", "partido", "electoral", "eleccion", "elección", "gobierno", "polític", "politic", "presidente", "ideolog"],
    "Problemas de España": ["problema", "preocup", "situación general", "situacion general"],
    "Economía y empleo": ["econom", "empleo", "trabajo", "paro", "salario", "ingres", "inflaci", "impuesto", "pension"],
    "Sanidad": ["sanidad", "salud", "médic", "medic", "hospital", "atención primaria"],
    "Educación": ["educa", "estudi", "universidad", "enseñanza"],
    "Inmigración": ["inmigra", "extranjero", "migraci"],
    "Valoración de líderes": ["valoración", "valoracion", "líder", "lider", "conoce ud.", "conocimiento de"],
    "Sociedad y valores": ["relig", "igualdad", "género", "genero", "familia", "discrimin", "seguridad"],
    "Otros temas": [],
}
PARTIES = ["TOTAL", "PSOE", "PP", "VOX", "Sumar", "Podemos", "ERC", "Junts", "EH Bildu", "PNV"]
# Agrupaciones convencionales, NO equivalen a autoubicación ideológica.
PARTY_GROUPS = {"Izquierda (partidos)": ["PSOE", "Sumar", "Podemos", "ERC", "EH Bildu"],
                "Derecha (partidos)": ["PP", "VOX", "Junts", "PNV"]}

def topic_for(block):
    title = (block["code"] + " " + block["title"]).casefold()
    for topic, words in TOPICS.items():
        if topic not in ("Todos los temas", "Otros temas") and any(w in title for w in words):
            return topic
    return "Otros temas"

def available_questions(data, sheet, topic, search):
    blocks = data[sheet]["blocks"]
    hits = [(i, b) for i, b in enumerate(blocks)
            if (topic == "Todos los temas" or topic_for(b) == topic)
            and search.casefold().strip() in (b["code"] + " " + b["title"]).casefold()]
    return hits

def pick_columns(block, sheet, mode, party, ideology, extra):
    cols = [c for c in block["table"].columns if c != "Respuesta"]
    if mode == "Partidos (recuerdo de voto)" and sheet == "RV EG23":
        if party == "Todos los partidos":
            proposed = [x for x in PARTIES if x in cols]
        elif party in PARTY_GROUPS:
            proposed = [x for x in PARTY_GROUPS[party] if x in cols]
        else:
            proposed = [party] if party in cols else []
    elif mode == "Autoubicación ideológica (escala CIS)" and sheet == "Escala de ideología":
        import unicodedata
        def score(c):
            s = str(c).strip()
            m = re.match(r"^(\d{1,2})(?:\s|$|[.:-])", s)
            return int(m.group(1)) if m else None
        ranges = {"Todos": range(1,11), "Izquierda (1–4)": range(1,5),
                  "Centro (5–6)": range(5,7), "Derecha (7–10)": range(7,11)}
        proposed = [c for c in cols if score(c) in ranges[ideology]]
    else:
        proposed = cols[:8]
    return [c for c in proposed if c in cols]

st.title("📊 Observatorio CIS")
st.caption("Panel de análisis · Temas, partidos y gráficos · Versión 4.1")
st.info("Resultados agregados oficiales. Los porcentajes proceden de las tablas publicadas por el CIS; no son microdatos ni estimaciones propias.")
if not DATA_DIR.exists():
    st.error("No se encuentra la carpeta Datos/ del repositorio.")
    st.stop()
studies = sorted(p for p in DATA_DIR.glob("*.xlsx") if p.is_file() and not p.is_symlink()
                 and not p.name.startswith("~$") and p.stat().st_size <= 30 * 1024 * 1024)
if not studies:
    st.error("No hay archivos Excel .xlsx en la carpeta Datos/.")
    st.stop()
study = st.selectbox("📁 Estudio del CIS", studies, format_func=lambda p: p.stem)
try:
    content = parse_study(study.read_bytes())
except Exception:
    st.error("No se pudo leer el Excel. Consulta los registros privados de Streamlit.")
    st.stop()
st.success(f"Estudio disponible: {study.stem} · {len(content)} hojas")

st.subheader("Filtros de análisis")
topic = st.selectbox("1. Tema general", list(TOPICS))
search = st.text_input("Buscar dentro del tema (opcional)", placeholder="Ej.: vivienda, conoce, voto…")
mode = st.selectbox("2. ¿Cómo quieres desglosar las respuestas?", [
    "Partidos (recuerdo de voto)", "Autoubicación ideológica (escala CIS)",
    "Sexo", "Edad", "Nivel de estudios", "Tamaño de municipio", "Ocupación", "Clase social", "Religiosidad", "Total general"])
party = "Todos los partidos"
ideology = "Todos"
if mode == "Partidos (recuerdo de voto)":
    party = st.selectbox("3. Partido o bloque político", ["Todos los partidos", "PSOE", "PP", "VOX", "Sumar", "Podemos", "ERC", "Junts", "EH Bildu", "PNV", "Izquierda (partidos)", "Derecha (partidos)"])
    st.caption("Las agrupaciones por partidos son orientativas. No equivalen a la autoubicación ideológica de los entrevistados y no se suman sus porcentajes.")
elif mode == "Autoubicación ideológica (escala CIS)":
    ideology = st.selectbox("3. Bloque ideológico", ["Todos", "Izquierda (1–4)", "Centro (5–6)", "Derecha (7–10)"])
    st.caption("Se seleccionan columnas de la escala de ideología; no se agregan porcentajes de diferentes columnas como si fueran un único grupo.")

sheet = {"Partidos (recuerdo de voto)": "RV EG23", "Autoubicación ideológica (escala CIS)": "Escala de ideología",
         "Total general": "Resultados"}.get(mode, mode)
if sheet not in content or not content[sheet]["blocks"]:
    st.warning("Este estudio no tiene tablas utilizables para el desglose seleccionado.")
    st.stop()
hits = available_questions(content, sheet, topic, search)
st.caption(f"{len(hits)} preguntas o bloques encontrados · Desglose: {sheet}")
if not hits:
    st.warning("No hay preguntas de ese tema en el desglose seleccionado. Prueba «Todos los temas» o cambia la búsqueda.")
    st.stop()
idx = st.selectbox("4. Pregunta", [i for i, _ in hits],
                   format_func=lambda i: f'{content[sheet]["blocks"][i]["code"]} · {content[sheet]["blocks"][i]["title"][:145]}')
b = content[sheet]["blocks"][idx]
st.caption(f"Fuente: {sheet} · {b['code']} · {b['title']}")
proposed = pick_columns(b, sheet, mode, party, ideology, None)
if not proposed:
    st.warning("No hay columnas disponibles para ese filtro en esta pregunta. Selecciona otro partido o desglose.")
    st.stop()
cols = st.multiselect("5. Columnas a comparar", [c for c in b["table"].columns if c != "Respuesta"],
                      default=proposed[:10])
if not cols:
    st.info("Selecciona al menos una columna.")
    st.stop()
view = numeric_view(b, cols)
chart = st.selectbox("6. Visualización", ["Barras horizontales apiladas", "Mapa de calor", "Barras comparativas", "Tabla de porcentajes", "Trasvases de voto"])
if chart == "Barras horizontales apiladas":
    answers = st.multiselect("Respuestas", view["Respuesta"].astype(str).tolist(), default=view["Respuesta"].astype(str).tolist()[:12])
    selected = view[view["Respuesta"].astype(str).isin(answers)].reset_index(drop=True)
    if not selected.empty:
        st.caption("Los porcentajes son los publicados. En preguntas de respuesta múltiple las barras pueden sumar más del 100 %.")
        plot_stacked(selected, cols, b["title"], normalize=False)
elif chart == "Mapa de calor":
    answers = st.multiselect("Respuestas", view["Respuesta"].astype(str).tolist(), default=view["Respuesta"].astype(str).tolist()[:18])
    selected = view[view["Respuesta"].astype(str).isin(answers)].reset_index(drop=True)
    if not selected.empty:
        plot_heatmap(selected, cols)
elif chart == "Barras comparativas":
    long = view.melt(id_vars="Respuesta", value_vars=cols, var_name="Grupo", value_name="Porcentaje")
    fig = px.bar(long, x="Porcentaje", y="Respuesta", color="Grupo", barmode="group", orientation="h")
    fig.update_layout(height=max(440, len(view)*30 + 160))
    st.plotly_chart(fig, use_container_width=True)
elif chart == "Tabla de porcentajes":
    st.dataframe(view.round(2), use_container_width=True, hide_index=True)
else:
    st.caption("Trasvases calculados únicamente a partir de la pregunta 17R de la hoja RV EG23 cuando esté disponible.")
    candidates = [q for q in content.get("RV EG23", {}).get("blocks", []) if re.search(r"^Pregunta\s+17R\b", q["code"], re.I)]
    if not candidates:
        st.warning("No se ha encontrado el cruce de intención actual y recuerdo de voto necesario para el trasvase.")
    else:
        flow_b = candidates[0]
        table = flow_b["table"]
        origins = [x for x in ["PSOE", "PP", "VOX", "Sumar"] if x in table.columns and pd.notna(flow_b["bases"].get(x))]
        origins = st.multiselect("Partidos de origen (2023)", [c for c in table.columns if c != "Respuesta" and c != "TOTAL" and pd.notna(flow_b["bases"].get(c))], default=origins)
        dests = st.multiselect("Respuestas de destino", table["Respuesta"].astype(str).tolist(), default=[x for x in table["Respuesta"].astype(str) if x.casefold() in ("psoe", "pp", "vox", "sumar")])
        flows=[]
        for origin in origins:
            base=float(flow_b["bases"][origin])
            for _, r in table.iterrows():
                if str(r["Respuesta"]) not in dests: continue
                pct=pd.to_numeric(r[origin],errors="coerce")
                if pd.notna(pct) and pct>0:
                    flows.append((origin,str(r["Respuesta"]),base*float(pct)/100,float(pct)))
        if flows:
            names=["2023 · "+x for x in origins]+["Actual · "+x for x in dests]
            fig=go.Figure(go.Sankey(node=dict(label=names),link=dict(
                source=[origins.index(x[0]) for x in flows],target=[len(origins)+dests.index(x[1]) for x in flows],
                value=[x[2] for x in flows],customdata=[x[3] for x in flows],
                hovertemplate="Base ponderada aproximada: %{value:.1f}<br>% del origen: %{customdata:.1f}<extra></extra>")))
            fig.update_layout(height=560)
            st.plotly_chart(fig,use_container_width=True)
            st.warning("Flujos aproximados basados en bases ponderadas publicadas. Los destinos omitidos no se muestran. No es una predicción ni un modelo propio.")
        else:
            st.info("Selecciona al menos un origen y un destino con datos disponibles.")

csv_button(view, "observatorio_cis_seleccion.csv")
with st.expander("Metodología, límites y seguridad"):
    st.markdown("""Los temas se clasifican mediante palabras clave; revisa la pregunta exacta antes de interpretar. Las tablas son agregadas, no microdatos. No se suman porcentajes de partidos para obtener una ideología conjunta. Para un agregado ponderado válido harían falta bases compatibles y un cálculo específico. La aplicación solo lee archivos Excel del directorio autorizado, limita su tamaño y no ejecuta macros. Los modelos privados de estimación no deben guardarse en este repositorio.""")
