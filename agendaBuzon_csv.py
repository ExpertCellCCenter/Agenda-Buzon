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
    name = Path(str(filename)).stem.lower()

    centro = "CC2" if "cc2" in name else "JV" if "jv" in name else "N/D"

    fecha = pd.NaT
    m_fecha = re.search(r"(\d{4}-\d{1,2}-\d{1,2})", name)
    if m_fecha:
        fecha = pd.to_datetime(m_fecha.group(1), errors="coerce")

    corte = "N/D"
    m_corte = re.search(r"_(1|2)_", name)
    if m_corte:
        corte = f"Corte {m_corte.group(1)}"
    elif pd.notna(fecha) and fecha.weekday() == 5:
        corte = "Sábado"
    else:
        # En los archivos revisados, el N/D corresponde al corte de sábado.
        corte = "Sábado"

    return centro, corte, fecha


def normalize_supervisor_name(value):
    txt = normalize_text(value)

    if txt in ["", "NAN", "NONE", "NULL", "N/D", "NA", "SIN DATO"]:
        return "SIN SUPERVISOR"

    if (
        "MARIA FERNANDA" in txt
        or txt == "FERNANDA"
        or " FERNANDA" in txt
        or "MARIA LUISA" in txt
        or txt == "LUISA"
        or " LUISA" in txt
    ):
        return "MARIA FERNANDA / MARIA LUISA"

    return txt


def pct(num, den):
    if den == 0:
        return 0
    return (num / den) * 100


def seek_start(file):
    try:
        file.seek(0)
    except Exception:
        pass


# =========================================================
# SMART EXCEL READER
# =========================================================
def detect_header_row(file, sheet_name):
    seek_start(file)

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
        "ENVIO", "ENVÍO", "MENSAJE", "WHATSAPP", "CONTACTADO",
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

        seek_start(file)
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
        file_name = getattr(file, "name", str(file))
        centro, corte, fecha_archivo = parse_filename(file_name)

        try:
            seek_start(file)
            xls = pd.ExcelFile(file, engine="openpyxl")
        except Exception:
            continue

        for sheet in xls.sheet_names:
            sheet_norm = normalize_text(sheet)

            # No tomar en cuenta hojas de Sin Supervisor / Encubadora desde el origen.
            if "SIN SUPERVISOR" in sheet_norm or "ENCUBADORA" in sheet_norm:
                continue

            df = read_one_sheet(file, sheet)

            if df.empty:
                continue

            columnas_hoja = set(df.columns.astype(str))

            df["ARCHIVO"] = Path(str(file_name)).name
            df["HOJA"] = sheet
            df["CENTRO_ARCHIVO"] = centro
            df["CORTE_ARCHIVO"] = corte
            df["FECHA_ARCHIVO"] = fecha_archivo
            df["COLUMNAS_HOJA"] = "|".join(sorted(columnas_hoja))

            all_data.append(df)

    if not all_data:
        return pd.DataFrame()

    return pd.concat(all_data, ignore_index=True, sort=False)



def get_repository_excel_files():
    """
    Lee archivos Excel directamente desde la carpeta del repositorio en GitHub/Streamlit.

    Estructura esperada:
    - repo/
      - app.py
      - Agenda buzon/
        - agenda_buzon_cc2_1_2026-04-29_validado.xlsx
        - agenda_buzon_jv_1_2026-04-29_validado.xlsx
    """
    base_dir = Path(__file__).parent if "__file__" in globals() else Path.cwd()
    folder = base_dir / "Agenda buzon"

    if not folder.exists():
        return []

    excel_files = []

    for file in sorted(folder.glob("*.xlsx")):
        name = file.name.lower()

        if name.startswith("~$"):
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


def find_contactado_whatsapp_column(df):
    for col in df.columns:
        col_norm = normalize_text(col).replace("_", " ")
        if "CONTACTADO" in col_norm and "WHATSAPP" in col_norm:
            return col
    return None


def find_respuesta_cliente_column(df):
    """Detecta la columna 'Repuesta/Respuesta por parte del cliente'.

    En los archivos viene como 'Repuesta por parte del cliente' (con typo),
    por eso se contempla RESPUESTA y REPUESTA.
    """
    for col in df.columns:
        col_norm = normalize_text(col).replace("_", " ")
        tiene_respuesta = "RESPUESTA" in col_norm or "REPUESTA" in col_norm
        if tiene_respuesta and "CLIENTE" in col_norm:
            return col
    return None


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

    respuesta_col = find_contactado_whatsapp_column(df) or find_best_column(df, [
        ["CONTACTADO", "WHATSAPP"],
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

    observacion_col = find_best_column(df, [
        ["OBSERVACION"],
        ["OBSERVACIONES"],
        ["COMENTARIO"],
        ["COMENTARIOS"],
        ["VALIDACION"],
        ["VALIDACIÓN"],
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

    respuesta_cliente_col = find_respuesta_cliente_column(df)

    return {
        "agente_col": agente_col,
        "supervisor_col": supervisor_col,
        "respuesta_col": respuesta_col,
        "observacion_col": observacion_col,
        "telefono_col": telefono_col,
        "folio_col": folio_col,
        "cliente_col": cliente_col,
        "respuesta_cliente_col": respuesta_cliente_col,
    }


# =========================================================
# CLASSIFICATION
# =========================================================
def classify_status(value, source=None):
    txt = normalize_text(value)
    source_norm = normalize_text(source)

    # Vacíos reales o celdas sin captura = no se entregó evidencia.
    if txt == "" or txt in ["NAN", "NONE", "NULL", "N/D", "NA", "SIN DATO"]:
        return "NO ENVIÓ EVIDENCIA"

    # Contactado tarde. Variantes con o sin "se", y con "la evidencia".
    if "HASTA QUE SE PIDIO" in txt or "HASTA QUE PIDIO" in txt or "HASTA QUE SE SOLICITO" in txt:
        return "CONTACTADO TARDE"

    # Contactado en tiempo.
    if txt in ["SI", "SÍ"] or txt.startswith("SI ") or txt.startswith("SÍ "):
        return "CONTACTADO EN TIEMPO"

    # En algunos archivos sin columna "Contactado por WhatsApp", Observaciones trae la evidencia.
    if "EN PROCESO DE VENTA" in txt or "PROCESO DE VENTA" in txt:
        return "CONTACTADO EN TIEMPO"

    if (
        "SE MANDA WHATS" in txt
        or "MANDA WHATS" in txt
        or "SE ENVIA WHATS" in txt
        or "ENVIA WHATS" in txt
        or "SE ENVIA MENSAJE" in txt
        or "ENVIO MENSAJE" in txt
        or "SE LE ENVIA MENSAJE" in txt
        or "MANDA MENSAJE" in txt
        or "SE MANDA MENSAJE" in txt
        or "SE COMPARTE INF" in txt
        or "SE VCOMPARTE INF" in txt
    ):
        return "CONTACTADO EN TIEMPO"

    # No contacto directo. Esto NO es lo mismo que no enviar evidencia.
    if (
        txt in ["NO", "NO.", "NO CONTACTADO"]
        or "NO CONTESTA" in txt
        or "NO CONTESTO" in txt
        or "NO RESPONDE" in txt
        or "NO LOCALIZADO" in txt
        or "CUELGA" in txt
    ):
        return "NO CONTACTADO"

    # Evidencias no enviadas o enviadas fuera del formato solicitado.
    if (
        "NO ENVIO" in txt
        or "NO SE ENVIO" in txt
        or "NO ENVIARON" in txt
        or "NO SE ENVIARON" in txt
        or "NO MANDO" in txt
        or "NO MANDARON" in txt
        or "NO SE MANDO" in txt
        or "NO SE MANDARON" in txt
        or "SIN EVIDENCIA" in txt
        or "FORMATO SOLICITADO" in txt
    ):
        return "NO ENVIÓ EVIDENCIA"

    if "NO TIENE WHATSAPP" in txt or "SIN WHATSAPP" in txt or "NO CUENTA CON WHATSAPP" in txt:
        return "CLIENTE SIN WHATSAPP"

    if "YA RENOVO" in txt or "RENOVO" in txt or "RENOVÓ" in txt:
        return "CLIENTE YA RENOVÓ"

    if (
        "NO VIABLE" in txt
        or "INVIABLE" in txt
        or "NO GESTIONABLE" in txt
        or "IMPROCEDENTE" in txt
        or "ADEUDO" in txt
        or "DEUDA" in txt
        or "NO TOMA DECISIONES" in txt
    ):
        return "CLIENTE NO VIABLE"

    if (
        "CUENTA RESTRINGIDA" in txt
        or "RESTRINGIDA" in txt
        or "RESTRINGIDO" in txt
        or "RESTRINGIERON" in txt
        or "RESTRINGIERON LA CUENTA" in txt
        or "LE RESTRINGIERON" in txt
        or "CUENTA RESTRING" in txt
    ):
        return "CUENTA RESTRINGIDA"

    if "LINEA SUSPENDIDA" in txt or "LÍNEA SUSPENDIDA" in txt or "SUSPENDIDA" in txt or "SUSPENDIDO" in txt:
        return "LÍNEA SUSPENDIDA"

    if "CAPTURA DE OTRO NUMERO" in txt or "CAPTURA OTRO NUMERO" in txt or "OTRO NUMERO" in txt or "OTRO NÚMERO" in txt:
        return "CAPTURA DE OTRO NÚMERO"

    # Si el valor viene de Observaciones como respaldo y no trae una señal clara,
    # no se convierte en una categoría libre para evitar mezclar planes/notas con categorías.
    if source_norm == "OBSERVACIONES":
        return "NO ENVIÓ EVIDENCIA"

    # Si aparece un valor nuevo en la columna formal de respuesta, se muestra como
    # categoría propia y normalizada, en lugar de mandarlo a "REVISAR".
    return txt.title()


def classify_respuesta_cliente(value):
    """Normaliza la columna 'Repuesta/Respuesta por parte del cliente'.

    Para el indicador de clientes que contestaron SOLO cuenta el valor "Si".
    Cualquier otro valor no se considera como cliente contestado.
    """
    txt = normalize_text(value)
    txt_simple = re.sub(r"[^A-Z0-9]+", " ", txt).strip()

    if txt_simple == "" or txt_simple in ["NAN", "NONE", "NULL", "N D", "NA", "SIN DATO"]:
        return "SIN RESPUESTA DEL CLIENTE"

    # ÚNICO valor que cuenta como cliente que sí contestó.
    if txt_simple == "SI":
        return "CLIENTE CONTESTÓ"

    # Se conserva como categoría informativa, pero NO cuenta como contestado.
    if txt_simple == "NO":
        return "CLIENTE NO CONTESTÓ"

    # Cualquier otro texto se muestra aparte y NO suma al porcentaje de respuesta del cliente.
    return "OTRA RESPUESTA / NO CUENTA COMO CONTESTÓ"


def cliente_contesto_desde_respuesta_cliente(value):
    """Solo cuenta como contestado cuando la normalización de la columna del cliente es 'CLIENTE CONTESTÓ'."""
    return value == "CLIENTE CONTESTÓ"

def score_category(cat):
    if cat == "CONTACTADO EN TIEMPO":
        return 100
    if cat == "CONTACTADO TARDE":
        return 70
    if cat in [
        "CLIENTE SIN WHATSAPP",
        "CLIENTE YA RENOVÓ",
        "CLIENTE NO VIABLE",
        "CUENTA RESTRINGIDA",
        "LÍNEA SUSPENDIDA",
        "CAPTURA DE OTRO NÚMERO",
    ]:
        return np.nan
    return 0


def is_excluido_contactacion(cat):
    return cat in [
        "CLIENTE SIN WHATSAPP",
        "CLIENTE YA RENOVÓ",
        "CLIENTE NO VIABLE",
        "CUENTA RESTRINGIDA",
        "LÍNEA SUSPENDIDA",
        "CAPTURA DE OTRO NÚMERO",
    ]


def get_row_response_value(row, respuesta_col, observacion_col):
    """Respeta Contactado por WhatsApp cuando existe en la hoja.

    Si la hoja no trae esa columna, usa Observaciones como respaldo.
    Si la columna sí existe pero la celda está vacía, se conserva vacía para
    clasificarla como NO ENVIÓ EVIDENCIA.
    """
    columnas_hoja = str(row.get("COLUMNAS_HOJA", "")).split("|")

    if respuesta_col and respuesta_col in columnas_hoja:
        return row.get(respuesta_col, ""), "CONTACTADO_POR_WHATSAPP"

    if observacion_col and observacion_col in columnas_hoja:
        return row.get(observacion_col, ""), "OBSERVACIONES"

    if respuesta_col:
        return row.get(respuesta_col, ""), "CONTACTADO_POR_WHATSAPP"

    return "", "NO DETECTADA"


def prepare_data(df):
    df = df.copy()

    cols = detect_columns(df)

    agente_col = cols["agente_col"]
    supervisor_col = cols["supervisor_col"]
    respuesta_col = cols["respuesta_col"]
    observacion_col = cols["observacion_col"]
    telefono_col = cols["telefono_col"]
    folio_col = cols["folio_col"]
    cliente_col = cols["cliente_col"]
    respuesta_cliente_col = cols["respuesta_cliente_col"]

    df["AGENTE_DETECTADO"] = (
        df[agente_col].apply(normalize_text)
        if agente_col else "SIN AGENTE"
    )
    df["AGENTE_DETECTADO"] = df["AGENTE_DETECTADO"].replace("", "SIN AGENTE")

    df["SUPERVISOR_DETECTADO"] = (
        df[supervisor_col].apply(normalize_supervisor_name)
        if supervisor_col else "SIN SUPERVISOR"
    )
    df["SUPERVISOR_DETECTADO"] = df["SUPERVISOR_DETECTADO"].apply(normalize_supervisor_name)

    respuesta_info = df.apply(
        lambda row: get_row_response_value(row, respuesta_col, observacion_col),
        axis=1,
        result_type="expand",
    )
    df["RESPUESTA_DETECTADA"] = respuesta_info[0]
    df["FUENTE_RESPUESTA"] = respuesta_info[1]

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

    df["RESPUESTA_CLIENTE_DETECTADA"] = (
        df[respuesta_cliente_col]
        if respuesta_cliente_col else ""
    )
    df["RESPUESTA_CLIENTE_NORMALIZADA"] = df["RESPUESTA_CLIENTE_DETECTADA"].apply(classify_respuesta_cliente)
    df["CLIENTE_CONTESTO"] = df["RESPUESTA_CLIENTE_NORMALIZADA"].apply(cliente_contesto_desde_respuesta_cliente)

    df["COLUMNA_AGENTE_USADA"] = agente_col or "NO DETECTADA"
    df["COLUMNA_SUPERVISOR_USADA"] = supervisor_col or "NO DETECTADA"
    df["COLUMNA_RESPUESTA_USADA"] = respuesta_col or "NO DETECTADA"
    df["COLUMNA_RESPUESTA_CLIENTE_USADA"] = respuesta_cliente_col or "NO DETECTADA"
    df["COLUMNA_OBSERVACION_RESPALDO"] = observacion_col or "NO DETECTADA"
    df["COLUMNA_TELEFONO_USADA"] = telefono_col or "NO DETECTADA"

    df["CATEGORIA"] = df.apply(
        lambda row: classify_status(row["RESPUESTA_DETECTADA"], row["FUENTE_RESPUESTA"]),
        axis=1,
    )
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
    base["ES_NO_CONTACTADO"] = base["CATEGORIA"].eq("NO CONTACTADO")
    base["ES_SIN_EVIDENCIA"] = base["CATEGORIA"].eq("NO ENVIÓ EVIDENCIA")
    base["ES_SIN_WHATSAPP"] = base["CATEGORIA"].eq("CLIENTE SIN WHATSAPP")
    base["ES_YA_RENOVO"] = base["CATEGORIA"].eq("CLIENTE YA RENOVÓ")
    base["ES_NO_VIABLE"] = base["CATEGORIA"].eq("CLIENTE NO VIABLE")
    base["ES_CUENTA_RESTRINGIDA"] = base["CATEGORIA"].eq("CUENTA RESTRINGIDA")
    base["ES_LINEA_SUSPENDIDA"] = base["CATEGORIA"].eq("LÍNEA SUSPENDIDA")
    base["ES_OTRO_NUMERO"] = base["CATEGORIA"].eq("CAPTURA DE OTRO NÚMERO")
    base["ES_CLIENTE_CONTESTO"] = base["CLIENTE_CONTESTO"].fillna(False).astype(bool)
    base["ES_CLIENTE_NO_CONTESTO"] = base["RESPUESTA_CLIENTE_NORMALIZADA"].eq("CLIENTE NO CONTESTÓ")
    base["ES_SIN_RESPUESTA_CLIENTE"] = base["RESPUESTA_CLIENTE_NORMALIZADA"].eq("SIN RESPUESTA DEL CLIENTE")

    out = (
        base.groupby(group_cols, dropna=False)
        .agg(
            TOTAL=("CATEGORIA", "size"),
            CONTACTABLES=("ES_CONTACTABLE", "sum"),
            CONTACTADOS=("ES_CONTACTADO", "sum"),
            EN_TIEMPO=("ES_EN_TIEMPO", "sum"),
            TARDE=("ES_TARDE", "sum"),
            NO_CONTACTADO=("ES_NO_CONTACTADO", "sum"),
            SIN_EVIDENCIA=("ES_SIN_EVIDENCIA", "sum"),
            SIN_WHATSAPP=("ES_SIN_WHATSAPP", "sum"),
            YA_RENOVO=("ES_YA_RENOVO", "sum"),
            NO_VIABLE=("ES_NO_VIABLE", "sum"),
            CUENTA_RESTRINGIDA=("ES_CUENTA_RESTRINGIDA", "sum"),
            LINEA_SUSPENDIDA=("ES_LINEA_SUSPENDIDA", "sum"),
            OTRO_NUMERO=("ES_OTRO_NUMERO", "sum"),
            CLIENTES_CONTESTARON=("ES_CLIENTE_CONTESTO", "sum"),
            CLIENTES_NO_CONTESTARON=("ES_CLIENTE_NO_CONTESTO", "sum"),
            SIN_RESPUESTA_CLIENTE=("ES_SIN_RESPUESTA_CLIENTE", "sum"),
            SCORE_PROMEDIO=("SCORE", "mean"),
        )
        .reset_index()
    )

    out["% CONTACTACIÓN"] = out.apply(lambda r: pct(r["CONTACTADOS"], r["CONTACTABLES"]), axis=1)
    out["% EN TIEMPO"] = out.apply(lambda r: pct(r["EN_TIEMPO"], r["CONTACTABLES"]), axis=1)
    out["% TARDE"] = out.apply(lambda r: pct(r["TARDE"], r["CONTACTABLES"]), axis=1)
    out["% NO CONTACTADO"] = out.apply(lambda r: pct(r["NO_CONTACTADO"], r["CONTACTABLES"]), axis=1)
    out["% SIN EVIDENCIA"] = out.apply(lambda r: pct(r["SIN_EVIDENCIA"], r["CONTACTABLES"]), axis=1)
    out["% CLIENTE CONTESTÓ"] = out.apply(lambda r: pct(r["CLIENTES_CONTESTARON"], r["TOTAL"]), axis=1)

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
st.sidebar.caption("Leyendo archivos .xlsx desde la carpeta Agenda buzon del repositorio.")
st.sidebar.write(f"Archivos detectados: {len(repo_excel_files)}")

with st.sidebar.expander("Ver archivos detectados", expanded=False):
    if repo_excel_files:
        for file in repo_excel_files:
            st.write(f"- {file.name}")
    else:
        st.write("No se encontraron archivos .xlsx en la carpeta Agenda buzon.")

if not repo_excel_files:
    st.error(
        "No se encontraron archivos Excel en el repositorio. "
        "Coloca los archivos .xlsx dentro de la carpeta Agenda buzon/."
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

df["SUPERVISOR_DETECTADO"] = df["SUPERVISOR_DETECTADO"].apply(normalize_supervisor_name)

df = df[
    ~df["SUPERVISOR_DETECTADO"].isin(EXCLUDE_SUPERVISORES)
].copy()

st.sidebar.header("Filtros")

centros = sorted(df["CENTRO"].dropna().unique())
centro_sel = st.sidebar.multiselect("Centro", centros, default=centros)

cortes = sorted(df["CORTE"].dropna().unique())
corte_sel = st.sidebar.multiselect("Corte", cortes, default=cortes)

supervisores = sorted(df["SUPERVISOR_DETECTADO"].dropna().unique())
supervisor_sel = st.sidebar.multiselect("Supervisor / Equipo", supervisores, default=supervisores)

categorias = sorted(df["CATEGORIA"].dropna().unique())
cat_sel = st.sidebar.multiselect("Categoría", categorias, default=categorias)

fechas_validas = df["FECHA"].dropna()
if fechas_validas.empty:
    st.sidebar.warning("No se detectaron fechas válidas en los nombres de los archivos.")
    fecha_inicio = None
    fecha_fin = None
else:
    fecha_min = fechas_validas.min().date()
    fecha_max = fechas_validas.max().date()

    fecha_inicio = st.sidebar.date_input(
        "Fecha inicio",
        value=fecha_min,
        min_value=fecha_min,
        max_value=fecha_max,
    )
    fecha_fin = st.sidebar.date_input(
        "Fecha fin",
        value=fecha_max,
        min_value=fecha_min,
        max_value=fecha_max,
    )

    if fecha_inicio > fecha_fin:
        st.sidebar.error("La fecha inicio no puede ser mayor que la fecha fin.")
        st.stop()

agentes_disponibles = sorted(df["AGENTE_DETECTADO"].dropna().unique())
agente_sel = st.sidebar.multiselect(
    "Agente",
    agentes_disponibles,
    default=agentes_disponibles,
    help="Por defecto se muestran todos los agentes. Puedes filtrar uno o varios si necesitas revisar un caso específico."
)

df_f = df[
    df["CENTRO"].isin(centro_sel)
    & df["CORTE"].isin(corte_sel)
    & df["SUPERVISOR_DETECTADO"].isin(supervisor_sel)
    & df["AGENTE_DETECTADO"].isin(agente_sel)
    & df["CATEGORIA"].isin(cat_sel)
].copy()

if fecha_inicio is not None and fecha_fin is not None:
    df_f = df_f[
        (df_f["FECHA"].dt.date >= fecha_inicio)
        & (df_f["FECHA"].dt.date <= fecha_fin)
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
no_contactado = (df_f["CATEGORIA"] == "NO CONTACTADO").sum()
sin_evidencia = (df_f["CATEGORIA"] == "NO ENVIÓ EVIDENCIA").sum()
cliente_contesto = df_f["CLIENTE_CONTESTO"].fillna(False).astype(bool).sum()

sin_whatsapp = (df_f["CATEGORIA"] == "CLIENTE SIN WHATSAPP").sum()
ya_renovo = (df_f["CATEGORIA"] == "CLIENTE YA RENOVÓ").sum()
no_viable = (df_f["CATEGORIA"] == "CLIENTE NO VIABLE").sum()
cuenta_restringida = (df_f["CATEGORIA"] == "CUENTA RESTRINGIDA").sum()
linea_suspendida = (df_f["CATEGORIA"] == "LÍNEA SUSPENDIDA").sum()
otro_numero = (df_f["CATEGORIA"] == "CAPTURA DE OTRO NÚMERO").sum()

tasa_contactacion = pct(contactados, total_contactables)
tasa_en_tiempo = pct(en_tiempo, total_contactables)
tasa_tarde = pct(tarde, total_contactables)
tasa_sin_evidencia = pct(sin_evidencia, total_contactables)
tasa_cliente_contesto = pct(cliente_contesto, total)

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Total registros", f"{total:,}")
k2.metric("Contactables", f"{total_contactables:,}")
k3.metric("Tasa contactación", f"{tasa_contactacion:.2f}%")
k4.metric("En tiempo", f"{tasa_en_tiempo:.2f}%")
k5.metric("Cumplió tarde", f"{tasa_tarde:.2f}%")

k6, k7, k8, k9, k10 = st.columns(5)

k6.metric("Sin evidencia", f"{tasa_sin_evidencia:.2f}%")
k7.metric("No contactado", f"{no_contactado:,}")
k8.metric("Sin WhatsApp", f"{sin_whatsapp:,}")
k9.metric("No viable", f"{no_viable:,}")
k10.metric("Cliente contestó", f"{tasa_cliente_contesto:.2f}%")

st.divider()


# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📌 Resumen general",
    "👤 Agentes",
    "📈 Mejora",
    "🧑‍💼 Supervisores"
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
            y=["% CONTACTACIÓN", "% EN TIEMPO", "% CLIENTE CONTESTÓ"],
            markers=True,
            title="Evolución diaria: contactación, cumplimiento y clientes que contestaron"
        )
        fig.update_layout(yaxis_title="Porcentaje", yaxis_ticksuffix="%", legend_title_text="Indicador")
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
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Registros")
    st.plotly_chart(fig, use_container_width=True)


with tab2:
    agente_summary = build_summary(df_f, ["SUPERVISOR_DETECTADO", "AGENTE_DETECTADO"])
    agente_summary = agente_summary.sort_values("% CONTACTACIÓN", ascending=False)

    st.subheader("Vista por agente")
    st.caption("Se muestran todos los agentes que cumplen con los filtros seleccionados.")

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(
            agente_summary.sort_values("% CONTACTACIÓN", ascending=True),
            x="% CONTACTACIÓN",
            y="AGENTE_DETECTADO",
            orientation="h",
            color="SUPERVISOR_DETECTADO",
            title="Contactación por agente",
            text="% CONTACTACIÓN",
            hover_data=[
                "SUPERVISOR_DETECTADO",
                "TOTAL",
                "CONTACTABLES",
                "CONTACTADOS",
                "EN_TIEMPO",
                "TARDE",
                "NO_CONTACTADO",
                "SIN_EVIDENCIA",
                "CLIENTES_CONTESTARON",
                "% CLIENTE CONTESTÓ",
            ],
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(
            yaxis_title="Agente",
            xaxis_title="% Contactación",
            xaxis_ticksuffix="%",
            height=max(520, 28 * len(agente_summary)),
            margin=dict(l=20, r=70, t=60, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.bar(
            agente_summary.sort_values("% SIN EVIDENCIA", ascending=True),
            x="% SIN EVIDENCIA",
            y="AGENTE_DETECTADO",
            orientation="h",
            color="SUPERVISOR_DETECTADO",
            title="Mayor oportunidad: sin evidencia",
            text="% SIN EVIDENCIA",
            hover_data=[
                "SUPERVISOR_DETECTADO",
                "TOTAL",
                "CONTACTABLES",
                "NO_CONTACTADO",
                "SIN_EVIDENCIA",
                "CLIENTES_CONTESTARON",
                "% CLIENTE CONTESTÓ",
            ],
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(
            yaxis_title="Agente",
            xaxis_title="% Sin evidencia",
            xaxis_ticksuffix="%",
            height=max(520, 28 * len(agente_summary)),
            margin=dict(l=20, r=70, t=60, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

    fig = px.scatter(
        agente_summary,
        x="CONTACTABLES",
        y="% CONTACTACIÓN",
        size="CONTACTABLES",
        color="% EN TIEMPO",
        hover_name="AGENTE_DETECTADO",
        hover_data=[
            "SUPERVISOR_DETECTADO",
            "TOTAL",
            "CONTACTADOS",
            "EN_TIEMPO",
            "TARDE",
            "NO_CONTACTADO",
            "SIN_EVIDENCIA",
            "CLIENTES_CONTESTARON",
            "% CLIENTE CONTESTÓ",
        ],
        title="Volumen contactable vs contactación"
    )
    fig.update_layout(yaxis_ticksuffix="%")
    st.plotly_chart(fig, use_container_width=True)

    agente_cat = (
        df_f.groupby(["SUPERVISOR_DETECTADO", "AGENTE_DETECTADO", "CATEGORIA"])
        .size()
        .reset_index(name="TOTAL")
    )
    orden_agentes = agente_summary.sort_values("% CONTACTACIÓN", ascending=False)["AGENTE_DETECTADO"].tolist()

    fig = px.bar(
        agente_cat,
        x="TOTAL",
        y="AGENTE_DETECTADO",
        color="CATEGORIA",
        orientation="h",
        title="Detalle de resultados por agente",
        text="TOTAL",
        category_orders={"AGENTE_DETECTADO": orden_agentes},
        hover_data=["SUPERVISOR_DETECTADO"],
    )
    fig.update_traces(textposition="inside")
    fig.update_layout(
        yaxis_title="Agente",
        xaxis_title="Registros",
        height=max(560, 30 * len(orden_agentes)),
        margin=dict(l=20, r=40, t=60, b=40),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(agente_summary, use_container_width=True, hide_index=True)


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
            mejora,
            x="MEJORA_CONTACTACIÓN",
            y="AGENTE_DETECTADO",
            orientation="h",
            color="ESTATUS_MEJORA",
            title="Mejora en contactación por agente",
            text="MEJORA_CONTACTACIÓN"
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_ticksuffix="%",
            height=max(520, 28 * len(mejora)),
            margin=dict(l=20, r=70, t=60, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(mejora, use_container_width=True, hide_index=True)


with tab4:
    sup_summary = build_summary(df_f, ["SUPERVISOR_DETECTADO"])
    sup_summary = sup_summary.sort_values("% CONTACTACIÓN", ascending=False)

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(
            sup_summary.sort_values("% CONTACTACIÓN", ascending=True),
            x="% CONTACTACIÓN",
            y="SUPERVISOR_DETECTADO",
            orientation="h",
            title="Contactación por supervisor / equipo",
            text="% CONTACTACIÓN",
            hover_data=[
                "TOTAL",
                "CONTACTABLES",
                "CONTACTADOS",
                "EN_TIEMPO",
                "TARDE",
                "NO_CONTACTADO",
                "SIN_EVIDENCIA",
                "CLIENTES_CONTESTARON",
                "% CLIENTE CONTESTÓ",
            ],
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(
            xaxis_ticksuffix="%",
            yaxis_title="Supervisor / Equipo",
            xaxis_title="% Contactación"
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.bar(
            sup_summary.sort_values("% SIN EVIDENCIA", ascending=True),
            x="% SIN EVIDENCIA",
            y="SUPERVISOR_DETECTADO",
            orientation="h",
            title="Sin evidencia por supervisor / equipo",
            text="% SIN EVIDENCIA",
            hover_data=[
                "TOTAL",
                "CONTACTABLES",
                "NO_CONTACTADO",
                "SIN_EVIDENCIA",
                "CLIENTES_CONTESTARON",
                "% CLIENTE CONTESTÓ",
            ],
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(
            xaxis_ticksuffix="%",
            yaxis_title="Supervisor / Equipo",
            xaxis_title="% Sin evidencia"
        )
        st.plotly_chart(fig, use_container_width=True)

    fig = px.bar(
        sup_summary,
        x="SUPERVISOR_DETECTADO",
        y=["EN_TIEMPO", "TARDE", "NO_CONTACTADO", "SIN_EVIDENCIA"],
        title="Contactación por supervisor",
        barmode="stack",
        text_auto=True,
    )
    fig.update_layout(xaxis_title="Supervisor / Equipo", yaxis_title="Registros")
    st.plotly_chart(fig, use_container_width=True)

    sup_cat = df_f.groupby(["SUPERVISOR_DETECTADO", "CATEGORIA"]).size().reset_index(name="TOTAL")
    fig = px.bar(
        sup_cat,
        x="TOTAL",
        y="SUPERVISOR_DETECTADO",
        color="CATEGORIA",
        orientation="h",
        title="Composición de resultados por supervisor / equipo",
        text="TOTAL",
    )
    fig.update_traces(textposition="inside")
    fig.update_layout(yaxis_title="Supervisor / Equipo", xaxis_title="Registros")
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(sup_summary, use_container_width=True, hide_index=True)
