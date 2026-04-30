import re
import unicodedata
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Reporte Agenda Buzón",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# HELPERS
# =========================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x).strip().upper()
    x = unicodedata.normalize("NFKD", x)
    x = "".join(c for c in x if not unicodedata.combining(c))
    x = re.sub(r"\s+", " ", x)
    return x


def clean_column_name(col):
    col = normalize_text(col)
    col = col.replace(".", "")
    col = col.replace("/", " ")
    col = col.replace("-", " ")
    col = col.replace(":", " ")
    col = col.replace("\n", " ")
    col = re.sub(r"\s+", "_", col)
    return col


def parse_filename(filename):
    name = Path(filename).stem.lower()

    centro = "CC2" if "cc2" in name else "JV" if "jv" in name else "N/D"

    corte = "N/D"
    m_corte = re.search(r"_(1|2)_", name)
    if m_corte:
        corte = f"Corte {m_corte.group(1)}"

    fecha = pd.NaT
    m_fecha = re.search(r"(\d{4}-\d{1,2}-\d{1,2})", name)
    if m_fecha:
        fecha = pd.to_datetime(m_fecha.group(1), errors="coerce")

    return centro, corte, fecha


def pct(num, den):
    if den == 0:
        return 0
    return (num / den) * 100


# =========================================================
# SMART EXCEL READER
# =========================================================
def detect_header_row(file, sheet_name):
    preview = pd.read_excel(
        file,
        sheet_name=sheet_name,
        header=None,
        nrows=20,
        engine="openpyxl"
    )

    keywords = [
        "AGENTE", "EJECUTIVO", "ASESOR", "GESTOR", "NOMBRE",
        "SUPERVISOR", "JEFE",
        "ENVIO", "ENVÍO", "MENSAJE", "WHATSAPP",
        "EVIDENCIA", "RESULTADO", "OBSERVACION", "OBSERVACIÓN",
        "TELEFONO", "TELÉFONO", "CELULAR", "FOLIO", "CLIENTE"
    ]

    best_row = 0
    best_score = -1

    for idx, row in preview.iterrows():
        txt = " ".join([normalize_text(v) for v in row.values])
        score = sum(1 for k in keywords if normalize_text(k) in txt)

        non_empty = row.notna().sum()
        score += min(non_empty, 10) * 0.1

        if score > best_score:
            best_score = score
            best_row = idx

    return best_row


def read_one_sheet(file, sheet_name):
    try:
        header_row = detect_header_row(file, sheet_name)

        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            header=header_row,
            engine="openpyxl"
        )

        df = df.dropna(how="all")
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", case=False, regex=True)]

        if df.empty:
            return pd.DataFrame()

        df.columns = [clean_column_name(c) for c in df.columns]

        return df

    except Exception:
        return pd.DataFrame()


def read_uploaded_files(uploaded_files):
    all_data = []

    for file in uploaded_files:
        centro, corte, fecha_archivo = parse_filename(file.name)

        try:
            xls = pd.ExcelFile(file, engine="openpyxl")
        except Exception:
            continue

        for sheet in xls.sheet_names:
            df = read_one_sheet(file, sheet)

            if df.empty:
                continue

            df["ARCHIVO"] = file.name
            df["HOJA"] = sheet
            df["CENTRO_ARCHIVO"] = centro
            df["CORTE_ARCHIVO"] = corte
            df["FECHA_ARCHIVO"] = fecha_archivo

            all_data.append(df)

    if not all_data:
        return pd.DataFrame()

    return pd.concat(all_data, ignore_index=True, sort=False)




def get_repository_excel_files():
    """
    Lee archivos Excel desde el mismo repositorio para despliegue en GitHub/Streamlit.

    Estructura recomendada:
    - repo/
      - app.py
      - data/
        - agenda_buzon_cc2_1_2026-04-29_validado.xlsx
        - agenda_buzon_jv_1_2026-04-29_validado.xlsx

    También busca en la misma carpeta del script por si prefieres dejar los Excel junto al .py.
    """
    base_dir = Path(__file__).parent

    search_dirs = [
        base_dir / "Agenda buzon",
        base_dir / "datos",
        base_dir / "excel",
        base_dir / "archivos",
        base_dir,
    ]

    excel_files = []

    for folder in search_dirs:
        if not folder.exists():
            continue

        for file in sorted(folder.glob("*.xlsx")):
            name = file.name.lower()

            if name.startswith("~$"):
                continue

            if "agenda_buzon" not in name:
                continue

            excel_files.append(file)

    unique_files = []
    seen = set()

    for file in excel_files:
        key = str(file.resolve())
        if key not in seen:
            unique_files.append(file)
            seen.add(key)

    return unique_files


# =========================================================
# SMART COLUMN DETECTION
# =========================================================
def find_best_column(df, keyword_groups):
    scores = {}

    for col in df.columns:
        col_norm = normalize_text(col).replace("_", " ")
        score = 0

        for group in keyword_groups:
            group_score = 0
            for word in group:
                word_norm = normalize_text(word)
                if word_norm in col_norm:
                    group_score += 1

            if group_score > 0:
                score += group_score * len(group)

        if score > 0:
            scores[col] = score

    if not scores:
        return None

    return max(scores, key=scores.get)


def detect_columns(df):
    agente_col = find_best_column(df, [
        ["AGENTE"],
        ["EJECUTIVO"],
        ["ASESOR"],
        ["GESTOR"],
        ["NOMBRE", "AGENTE"],
        ["NOMBRE", "EJECUTIVO"],
    ])

    supervisor_col = find_best_column(df, [
        ["SUPERVISOR"],
        ["JEFE"],
        ["JEFE", "DIRECTO"],
        ["COORDINADOR"],
    ])

    respuesta_col = find_best_column(df, [
        ["ENVIO", "MENSAJE"],
        ["ENVIO", "WHATSAPP"],
        ["MENSAJE"],
        ["EVIDENCIA"],
        ["RESULTADO"],
        ["RESPUESTA"],
        ["OBSERVACION"],
        ["COMENTARIO"],
        ["VALIDACION"],
        ["ESTATUS"],
    ])

    telefono_col = find_best_column(df, [
        ["TELEFONO"],
        ["TEL"],
        ["CELULAR"],
        ["NUMERO"],
        ["NUMERO", "CLIENTE"],
    ])

    folio_col = find_best_column(df, [
        ["FOLIO"],
        ["ORDEN"],
        ["ID"],
        ["CUENTA"],
    ])

    cliente_col = find_best_column(df, [
        ["CLIENTE"],
        ["NOMBRE", "CLIENTE"],
        ["TITULAR"],
    ])

    return {
        "agente_col": agente_col,
        "supervisor_col": supervisor_col,
        "respuesta_col": respuesta_col,
        "telefono_col": telefono_col,
        "folio_col": folio_col,
        "cliente_col": cliente_col,
    }


# =========================================================
# CLASSIFICATION
# =========================================================
def classify_status(value):
    txt = normalize_text(value)

    if txt == "" or txt in ["NAN", "NONE", "NULL", "N/D", "NA", "SIN DATO"]:
        return "NO CONTACTADO / SIN EVIDENCIA"

    # Cumplimiento operativo
    if "HASTA QUE SE PIDIO" in txt or "HASTA QUE SE PIDIO LA EVIDENCIA" in txt:
        return "CONTACTADO TARDE"

    if txt in ["SI", "SÍ"] or txt.startswith("SI "):
        return "CONTACTADO EN TIEMPO"

    # Casos no contactables / informativos
    if "NO TIENE WHATSAPP" in txt or "SIN WHATSAPP" in txt:
        return "CLIENTE SIN WHATSAPP"

    if "LINEA SUSPENDIDA" in txt or "LÍNEA SUSPENDIDA" in txt:
        return "LÍNEA SUSPENDIDA"

    if "RESTRINGIDA" in txt or "CUENTA RESTRINGIDA" in txt:
        return "CUENTA RESTRINGIDA"

    if "CAPTURA DE OTRO NUMERO" in txt or "CAPTURA OTRO NUMERO" in txt or "OTRO NUMERO" in txt:
        return "CAPTURA DE OTRO NÚMERO"

    if "YA RENOVO" in txt or "RENOVO" in txt or "RENOVÓ" in txt:
        return "CLIENTE YA RENOVÓ"

    # Normalización de no viable
    if (
        "NO VIABLE" in txt
        or "N VIABLE" in txt
        or "INVIABLE" in txt
        or "NO GESTIONABLE" in txt
        or "NO GESTION" in txt
        or "ADEUDO" in txt
        or "DEUDA" in txt
        or "IMPROCEDENTE" in txt
    ):
        return "CLIENTE NO VIABLE"

    # Si aparece una respuesta nueva, se muestra tal cual ya normalizada,
    # en lugar de mandarla a una categoría genérica como REVISAR.
    return txt

def score_category(cat):
    if cat == "CONTACTADO EN TIEMPO":
        return 100
    if cat == "CONTACTADO TARDE":
        return 70
    if cat in [
        "CLIENTE SIN WHATSAPP",
        "CLIENTE YA RENOVÓ",
        "CLIENTE NO VIABLE",
    ]:
        return np.nan
    return 0


def is_excluido_contactacion(cat):
    return cat in [
        "CLIENTE SIN WHATSAPP",
        "CLIENTE YA RENOVÓ",
        "CLIENTE NO VIABLE",
    ]


def prepare_data(df):
    df = df.copy()

    cols = detect_columns(df)

    agente_col = cols["agente_col"]
    supervisor_col = cols["supervisor_col"]
    respuesta_col = cols["respuesta_col"]
    telefono_col = cols["telefono_col"]
    folio_col = cols["folio_col"]
    cliente_col = cols["cliente_col"]

    df["AGENTE_DETECTADO"] = (
        df[agente_col].apply(normalize_text)
        if agente_col else "SIN AGENTE"
    )

    df["SUPERVISOR_DETECTADO"] = (
        df[supervisor_col].apply(normalize_text)
        if supervisor_col else "SIN SUPERVISOR"
    )

    df["RESPUESTA_DETECTADA"] = (
        df[respuesta_col]
        if respuesta_col else ""
    )

    df["TELEFONO_DETECTADO"] = (
        df[telefono_col].astype(str)
        if telefono_col else ""
    )

    df["FOLIO_DETECTADO"] = (
        df[folio_col].astype(str)
        if folio_col else ""
    )

    df["CLIENTE_DETECTADO"] = (
        df[cliente_col].astype(str)
        if cliente_col else ""
    )

    df["COLUMNA_AGENTE_USADA"] = agente_col or "NO DETECTADA"
    df["COLUMNA_SUPERVISOR_USADA"] = supervisor_col or "NO DETECTADA"
    df["COLUMNA_RESPUESTA_USADA"] = respuesta_col or "NO DETECTADA"
    df["COLUMNA_TELEFONO_USADA"] = telefono_col or "NO DETECTADA"

    df["CATEGORIA"] = df["RESPUESTA_DETECTADA"].apply(classify_status)
    df["SCORE"] = df["CATEGORIA"].apply(score_category)

    df["FECHA"] = pd.to_datetime(df["FECHA_ARCHIVO"], errors="coerce")
    df["FECHA_TXT"] = df["FECHA"].dt.strftime("%Y-%m-%d")
    df["CENTRO"] = df["CENTRO_ARCHIVO"]
    df["CORTE"] = df["CORTE_ARCHIVO"]

    return df


# =========================================================
# SUMMARIES
# =========================================================
def build_summary(df, group_cols):
    base = df.copy()

    base["ES_CONTACTABLE"] = ~base["CATEGORIA"].apply(is_excluido_contactacion)
    base["ES_CONTACTADO"] = base["CATEGORIA"].isin([
        "CONTACTADO EN TIEMPO",
        "CONTACTADO TARDE"
    ])
    base["ES_EN_TIEMPO"] = base["CATEGORIA"].eq("CONTACTADO EN TIEMPO")
    base["ES_TARDE"] = base["CATEGORIA"].eq("CONTACTADO TARDE")
    base["ES_SIN_EVIDENCIA"] = base["CATEGORIA"].eq("NO CONTACTADO / SIN EVIDENCIA")
    base["ES_SIN_WHATSAPP"] = base["CATEGORIA"].eq("CLIENTE SIN WHATSAPP")
    base["ES_YA_RENOVO"] = base["CATEGORIA"].eq("CLIENTE YA RENOVÓ")
    base["ES_NO_VIABLE"] = base["CATEGORIA"].eq("CLIENTE NO VIABLE")
    base["ES_REVISAR"] = base["CATEGORIA"].eq("REVISAR")

    out = (
        base.groupby(group_cols, dropna=False)
        .agg(
            TOTAL=("CATEGORIA", "size"),
            CONTACTABLES=("ES_CONTACTABLE", "sum"),
            CONTACTADOS=("ES_CONTACTADO", "sum"),
            EN_TIEMPO=("ES_EN_TIEMPO", "sum"),
            TARDE=("ES_TARDE", "sum"),
            SIN_EVIDENCIA=("ES_SIN_EVIDENCIA", "sum"),
            SIN_WHATSAPP=("ES_SIN_WHATSAPP", "sum"),
            YA_RENOVO=("ES_YA_RENOVO", "sum"),
            NO_VIABLE=("ES_NO_VIABLE", "sum"),
            REVISAR=("ES_REVISAR", "sum"),
            SCORE_PROMEDIO=("SCORE", "mean"),
        )
        .reset_index()
    )

    out["% CONTACTACIÓN"] = out.apply(lambda r: pct(r["CONTACTADOS"], r["CONTACTABLES"]), axis=1)
    out["% EN TIEMPO"] = out.apply(lambda r: pct(r["EN_TIEMPO"], r["CONTACTABLES"]), axis=1)
    out["% TARDE"] = out.apply(lambda r: pct(r["TARDE"], r["CONTACTABLES"]), axis=1)
    out["% SIN EVIDENCIA"] = out.apply(lambda r: pct(r["SIN_EVIDENCIA"], r["CONTACTABLES"]), axis=1)

    out["SCORE_PROMEDIO"] = out["SCORE_PROMEDIO"].fillna(0).round(2)

    return out


def build_improvement(df):
    daily = build_summary(df, ["AGENTE_DETECTADO", "FECHA"])

    if daily.empty or daily["FECHA"].nunique() < 2:
        return pd.DataFrame()

    daily = daily.sort_values("FECHA")

    first = daily.groupby("AGENTE_DETECTADO").first().reset_index()
    last = daily.groupby("AGENTE_DETECTADO").last().reset_index()

    imp = first[[
        "AGENTE_DETECTADO",
        "% CONTACTACIÓN",
        "% EN TIEMPO",
        "SCORE_PROMEDIO"
    ]].merge(
        last[[
            "AGENTE_DETECTADO",
            "% CONTACTACIÓN",
            "% EN TIEMPO",
            "SCORE_PROMEDIO"
        ]],
        on="AGENTE_DETECTADO",
        suffixes=("_INICIAL", "_FINAL")
    )

    imp["MEJORA_CONTACTACIÓN"] = imp["% CONTACTACIÓN_FINAL"] - imp["% CONTACTACIÓN_INICIAL"]
    imp["MEJORA_EN_TIEMPO"] = imp["% EN TIEMPO_FINAL"] - imp["% EN TIEMPO_INICIAL"]
    imp["MEJORA_SCORE"] = imp["SCORE_PROMEDIO_FINAL"] - imp["SCORE_PROMEDIO_INICIAL"]

    imp["ESTATUS_MEJORA"] = np.where(
        imp["MEJORA_CONTACTACIÓN"] > 0,
        "MEJORÓ",
        np.where(imp["MEJORA_CONTACTACIÓN"] < 0, "BAJÓ", "SE MANTUVO")
    )

    return imp.sort_values("MEJORA_CONTACTACIÓN", ascending=False)


def to_excel(dataframes):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in dataframes.items():
            data.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    return output.getvalue()


# =========================================================
# APP
# =========================================================
st.title("📊 Reporte Agenda Buzón")
st.caption("Reporte flexible para diferentes formatos de archivos: CC2, JV, corte 1, corte 2 y layouts distintos.")

repo_excel_files = get_repository_excel_files()

st.sidebar.subheader("Fuente de datos")
st.sidebar.caption("Leyendo archivos .xlsx desde el repositorio.")
st.sidebar.write(f"Archivos detectados: {len(repo_excel_files)}")

with st.sidebar.expander("Ver archivos detectados", expanded=False):
    if repo_excel_files:
        for file in repo_excel_files:
            st.write(f"- {file.name}")
    else:
        st.write("No se encontraron archivos agenda_buzon*.xlsx.")

if not repo_excel_files:
    st.error(
        "No se encontraron archivos Excel en el repositorio. "
        "Coloca los archivos agenda_buzon*.xlsx en la carpeta data/ o junto al archivo .py."
    )
    st.stop()

raw = read_uploaded_files(repo_excel_files)

if raw.empty:
    st.error("No se pudo leer información válida de los archivos.")
    st.stop()

df = prepare_data(raw)

# Excluir supervisores que no deben entrar al reporte
EXCLUDE_SUPERVISORES = [
    "ENCUBADORA",
    "SIN SUPERVISOR",
    "SIN SUPERVISOR JV",
    "",
    "NAN",
    "NONE",
    "NULL",
    "N/D",
]

df["SUPERVISOR_DETECTADO"] = df["SUPERVISOR_DETECTADO"].apply(normalize_text)

df = df[
    ~df["SUPERVISOR_DETECTADO"].isin(EXCLUDE_SUPERVISORES)
].copy()

st.sidebar.header("Filtros")

centros = sorted(df["CENTRO"].dropna().unique())
centro_sel = st.sidebar.multiselect("Centro", centros, default=centros)

cortes = sorted(df["CORTE"].dropna().unique())
corte_sel = st.sidebar.multiselect("Corte", cortes, default=cortes)

fechas = sorted(df["FECHA_TXT"].dropna().unique())
fecha_sel = st.sidebar.multiselect("Fecha", fechas, default=fechas)

categorias = sorted(df["CATEGORIA"].dropna().unique())
cat_sel = st.sidebar.multiselect("Categoría", categorias, default=categorias)

df_f = df[
    df["CENTRO"].isin(centro_sel)
    & df["CORTE"].isin(corte_sel)
    & df["FECHA_TXT"].isin(fecha_sel)
    & df["CATEGORIA"].isin(cat_sel)
].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()


# =========================================================
# KPIS
# =========================================================
total = len(df_f)

contactables = df_f[
    ~df_f["CATEGORIA"].apply(is_excluido_contactacion)
]

total_contactables = len(contactables)

en_tiempo = (df_f["CATEGORIA"] == "CONTACTADO EN TIEMPO").sum()
tarde = (df_f["CATEGORIA"] == "CONTACTADO TARDE").sum()
contactados = en_tiempo + tarde
sin_evidencia = (df_f["CATEGORIA"] == "NO CONTACTADO / SIN EVIDENCIA").sum()

sin_whatsapp = (df_f["CATEGORIA"] == "CLIENTE SIN WHATSAPP").sum()
ya_renovo = (df_f["CATEGORIA"] == "CLIENTE YA RENOVÓ").sum()
no_viable = (df_f["CATEGORIA"] == "CLIENTE NO VIABLE").sum()

categorias_principales_metricas = [
    "CONTACTADO EN TIEMPO",
    "CONTACTADO TARDE",
    "NO CONTACTADO / SIN EVIDENCIA",
    "CLIENTE SIN WHATSAPP",
    "CLIENTE YA RENOVÓ",
    "CLIENTE NO VIABLE",
]

# Categorías reales que se muestran tal cual, por ejemplo:
# CUENTA RESTRINGIDA, LÍNEA SUSPENDIDA, CAPTURA DE OTRO NÚMERO, etc.
otros_casos = df_f[~df_f["CATEGORIA"].isin(categorias_principales_metricas)].shape[0]

tasa_contactacion = pct(contactados, total_contactables)
tasa_en_tiempo = pct(en_tiempo, total_contactables)
tasa_tarde = pct(tarde, total_contactables)
tasa_sin_evidencia = pct(sin_evidencia, total_contactables)

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Total registros", f"{total:,}")
k2.metric("Contactables", f"{total_contactables:,}")
k3.metric("Tasa contactación", f"{tasa_contactacion:.2f}%")
k4.metric("En tiempo", f"{tasa_en_tiempo:.2f}%")
k5.metric("Cumplió tarde", f"{tasa_tarde:.2f}%")

k6, k7, k8, k9, k10 = st.columns(5)

k6.metric("Sin evidencia", f"{tasa_sin_evidencia:.2f}%")
k7.metric("Sin WhatsApp", f"{sin_whatsapp:,}")
k8.metric("Ya renovó", f"{ya_renovo:,}")
k9.metric("No viable", f"{no_viable:,}")
k10.metric("Casos especiales", f"{otros_casos:,}")

st.divider()


# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📌 Resumen general",
    "👤 Agentes",
    "📈 Mejora",
    "🧑‍💼 Supervisores",
    "🧪 Diagnóstico columnas",
    "📥 Descargar reporte"
])


with tab1:
    c1, c2 = st.columns(2)

    with c1:
        cat_count = df_f["CATEGORIA"].value_counts().reset_index()
        cat_count.columns = ["CATEGORIA", "TOTAL"]

        fig = px.pie(
            cat_count,
            names="CATEGORIA",
            values="TOTAL",
            hole=0.45,
            title="Distribución general de resultados"
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        daily = build_summary(df_f, ["FECHA"])
        fig = px.line(
            daily,
            x="FECHA",
            y="% CONTACTACIÓN",
            markers=True,
            title="Tasa de contactación por día"
        )
        st.plotly_chart(fig, use_container_width=True)

    daily_cat = df_f.groupby(["FECHA", "CATEGORIA"]).size().reset_index(name="TOTAL")

    fig = px.bar(
        daily_cat,
        x="FECHA",
        y="TOTAL",
        color="CATEGORIA",
        barmode="group",
        title="Resultados diarios por categoría",
        text="TOTAL"
    )
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True)


with tab2:
    agente_summary = build_summary(df_f, ["AGENTE_DETECTADO"])
    agente_summary = agente_summary.sort_values("% CONTACTACIÓN", ascending=False)

    c1, c2 = st.columns(2)

    with c1:
        top = agente_summary.head(15)
        fig = px.bar(
            top,
            x="% CONTACTACIÓN",
            y="AGENTE_DETECTADO",
            orientation="h",
            title="Top agentes por contactación",
            text="% CONTACTACIÓN"
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        low = agente_summary.sort_values("% SIN EVIDENCIA", ascending=False).head(15)
        fig = px.bar(
            low,
            x="% SIN EVIDENCIA",
            y="AGENTE_DETECTADO",
            orientation="h",
            title="Mayor oportunidad: sin evidencia",
            text="% SIN EVIDENCIA"
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

    fig = px.scatter(
        agente_summary,
        x="CONTACTABLES",
        y="% CONTACTACIÓN",
        size="CONTACTABLES",
        color="% EN TIEMPO",
        hover_name="AGENTE_DETECTADO",
        title="Volumen contactable vs contactación"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(agente_summary, use_container_width=True)


with tab3:
    mejora = build_improvement(df_f)

    if mejora.empty:
        st.warning("Se necesitan al menos dos fechas para calcular mejora.")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("Agentes que mejoraron", int((mejora["ESTATUS_MEJORA"] == "MEJORÓ").sum()))
        m2.metric("Agentes que bajaron", int((mejora["ESTATUS_MEJORA"] == "BAJÓ").sum()))
        m3.metric("Agentes sin cambio", int((mejora["ESTATUS_MEJORA"] == "SE MANTUVO").sum()))

        fig = px.bar(
            mejora.head(20),
            x="MEJORA_CONTACTACIÓN",
            y="AGENTE_DETECTADO",
            orientation="h",
            color="ESTATUS_MEJORA",
            title="Mejora en contactación por agente",
            text="MEJORA_CONTACTACIÓN"
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(mejora, use_container_width=True)


with tab4:
    sup_summary = build_summary(df_f, ["SUPERVISOR_DETECTADO"])
    sup_summary = sup_summary.sort_values("% CONTACTACIÓN", ascending=False)

    fig = px.bar(
        sup_summary,
        x="SUPERVISOR_DETECTADO",
        y=["EN_TIEMPO", "TARDE", "SIN_EVIDENCIA"],
        title="Contactación por supervisor",
        barmode="stack"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(sup_summary, use_container_width=True)


with tab5:
    st.subheader("Diagnóstico de columnas detectadas")
    st.caption("Esta pestaña te ayuda a revisar si el archivo cambió de formato.")

    diag = df_f[[
        "ARCHIVO",
        "HOJA",
        "COLUMNA_AGENTE_USADA",
        "COLUMNA_SUPERVISOR_USADA",
        "COLUMNA_RESPUESTA_USADA",
        "COLUMNA_TELEFONO_USADA"
    ]].drop_duplicates()

    st.dataframe(diag, use_container_width=True)

    st.subheader("Valores detectados en respuesta")
    valores = (
        df_f["RESPUESTA_DETECTADA"]
        .astype(str)
        .value_counts()
        .reset_index()
    )
    valores.columns = ["RESPUESTA_DETECTADA", "TOTAL"]
    st.dataframe(valores, use_container_width=True)


with tab6:
    resumen_general = build_summary(df_f, ["CENTRO", "CORTE", "FECHA"])
    resumen_agentes = build_summary(df_f, ["CENTRO", "CORTE", "AGENTE_DETECTADO"])
    resumen_supervisores = build_summary(df_f, ["CENTRO", "CORTE", "SUPERVISOR_DETECTADO"])
    mejora = build_improvement(df_f)

    diagnostico = df_f[[
        "ARCHIVO",
        "HOJA",
        "COLUMNA_AGENTE_USADA",
        "COLUMNA_SUPERVISOR_USADA",
        "COLUMNA_RESPUESTA_USADA",
        "COLUMNA_TELEFONO_USADA"
    ]].drop_duplicates()

    excel_bytes = to_excel({
        "Base filtrada": df_f,
        "Resumen general": resumen_general,
        "Resumen agentes": resumen_agentes,
        "Resumen supervisores": resumen_supervisores,
        "Mejora": mejora,
        "Diagnostico columnas": diagnostico,
    })

    st.download_button(
        label="📥 Descargar reporte Excel",
        data=excel_bytes,
        file_name="reporte_agenda_buzon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Vista previa base filtrada")
    st.dataframe(df_f, use_container_width=True)
