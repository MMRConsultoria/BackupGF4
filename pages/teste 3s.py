import streamlit as st
import pandas as pd
import psycopg2
import uuid
import json
import ast
from datetime import datetime, date, timedelta
from io import BytesIO

st.set_page_config(page_title="3S - Dump RAW com Filtro de Data", layout="wide")

CERT_PATH = "aws-us-east-2-bundle.pem"

def ensure_cert_written():
    if "cert_written_dump" not in st.session_state:
        with open(CERT_PATH, "w", encoding="utf-8") as f:
            f.write(st.secrets["certs"]["aws_rds_us_east_2"])
        st.session_state["cert_written_dump"] = True

def get_db_conn():
    return psycopg2.connect(
        host=st.secrets["db"]["host"],
        port=st.secrets["db"]["port"],
        dbname=st.secrets["db"]["database"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        sslmode="verify-full",
        sslrootcert=CERT_PATH,
    )

def list_tables(conn, schema="public"):
    q = "SELECT table_name FROM information_schema.tables WHERE table_schema = %s AND table_type = 'BASE TABLE' ORDER BY table_name"
    return pd.read_sql(q, conn, params=(schema,))["table_name"].tolist()

def list_columns(conn, table_name, schema="public"):
    q = "SELECT column_name, data_type FROM information_schema.columns WHERE table_schema = %s AND table_name = %s ORDER BY ordinal_position"
    return pd.read_sql(q, conn, params=(schema, table_name))

def fetch_table_with_date(conn, table_name, date_col, start_date, limit=1000, schema="public"):
    if date_col and date_col != "Sem filtro de data":
        q = f'SELECT * FROM "{schema}"."{table_name}" WHERE "{date_col}" >= %s ORDER BY "{date_col}" DESC LIMIT %s'
        return pd.read_sql(q, conn, params=(start_date, int(limit)))
    else:
        q = f'SELECT * FROM "{schema}"."{table_name}" LIMIT %s'
        return pd.read_sql(q, conn, params=(int(limit),))

def _make_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    df_safe = df.copy()
    for col in df_safe.columns:
        if pd.api.types.is_datetime64_any_dtype(df_safe[col]):
            try:
                df_safe[col] = df_safe[col].dt.tz_localize(None)
            except:
                try:
                    df_safe[col] = df_safe[col].dt.tz_convert(None)
                except:
                    df_safe[col] = df_safe[col].astype(str)
        elif pd.api.types.is_timedelta64_dtype(df_safe[col]):
            df_safe[col] = df_safe[col].astype(str)
        elif df_safe[col].dtype == object:
            df_safe[col] = df_safe[col].apply(
                lambda x: str(x) if isinstance(x, (dict, list, bytes, uuid.UUID)) else x
            )
    return df_safe

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="data"):
    df_safe = _make_excel_safe(df)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_safe.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()

# ── NOVAS FUNÇÕES JSON ────────────────────────────────────────────────────────

def detect_json_columns(df: pd.DataFrame) -> list:
    """Detecta colunas que provavelmente contêm JSON (dict/string JSON)."""
    json_cols = []
    for col in df.columns:
        if df[col].dtype != object:
            continue
        sample = df[col].dropna().head(20)
        hits = 0
        for val in sample:
            if isinstance(val, dict):
                hits += 1
                continue
            if isinstance(val, str) and val.strip().startswith(("{", "[")):
                try:
                    json.loads(val)
                    hits += 1
                except:
                    pass
        if hits >= max(1, len(sample) * 0.3):  # pelo menos 30% da amostra é JSON
            json_cols.append(col)
    return json_cols

def expand_json_columns(df: pd.DataFrame, cols_to_expand: list) -> pd.DataFrame:
    """Expande colunas JSON em colunas separadas com prefixo 'coluna__chave'."""
    df_result = df.copy()

    for col in cols_to_expand:
        if col not in df_result.columns:
            continue

        def safe_parse(x):
            if pd.isna(x) or x == "":
                return {}
            try:
                if isinstance(x, dict):
                    return x
                if isinstance(x, str):
                    return json.loads(x)
            except:
                try:
                    return ast.literal_eval(x)
                except:
                    return {}
            return {}

        parsed = df_result[col].apply(safe_parse)

        # Só expande se tiver pelo menos 1 dict não vazio
        if parsed.apply(lambda x: bool(x)).any():
            expanded = pd.json_normalize(parsed)
            expanded.columns = [f"{col}__{c}" for c in expanded.columns]
            expanded.index = df_result.index

            # Insere as colunas expandidas logo após a coluna original
            pos = df_result.columns.get_loc(col) + 1
            for i, new_col in enumerate(expanded.columns):
                df_result.insert(pos + i, new_col, expanded[new_col])

    return df_result

# ─────────────────────────────────────────────────────────────────────────────

ensure_cert_written()
st.title("3S (Postgres) — Exportador RAW Excel com Filtro")

# Sidebar
with st.sidebar:
    st.header("📅 Filtro de Data")
    data_inicio = st.date_input("Desde quando?", value=date.today() - timedelta(days=7))
    st.info("O filtro será aplicado na coluna de data selecionada para cada tabela.")

    st.header("⚙️ Configurações")
    limit_global = st.number_input("Limite de linhas", min_value=1, max_value=1000000, value=100000)
    schema = st.text_input("Schema", value="public")

# Conexão
try:
    conn = get_db_conn()
    tables = list_tables(conn, schema=schema)
    st.success(f"Conectado! {len(tables)} tabelas encontradas.")
except Exception as e:
    st.error(f"Erro de conexão: {e}")
    st.stop()

# Seleção de Tabela
tbl = st.selectbox("Escolha a tabela para exportar:", tables)

if tbl:
    df_cols = list_columns(conn, tbl, schema=schema)

    cols_data_sugeridas = [
        c for c in df_cols["column_name"]
        if any(x in c.lower() for x in ["date", "dt", "at", "time"])
    ]
    cols_data_sugeridas = ["Sem filtro de data"] + cols_data_sugeridas

    col1, col2 = st.columns([1, 2])
    with col1:
        col_data_escolhida = st.selectbox("Coluna de data para filtrar:", cols_data_sugeridas)
    with col2:
        st.write(f"**Filtro:** `{col_data_escolhida}` >= `{data_inicio}`")

    if st.button("🚀 Carregar dados", type="primary"):
        with st.spinner(f"Buscando dados de `{tbl}`..."):
            try:
                df = fetch_table_with_date(
                    conn, tbl, col_data_escolhida, data_inicio,
                    limit=limit_global, schema=schema
                )

                if df.empty:
                    st.warning("Nenhum dado encontrado para este período.")
                else:
                    st.write(f"✅ {len(df)} linhas | {df.shape[1]} colunas carregadas.")

                    # ── Detectar colunas JSON ──────────────────────────────
                    json_cols_detectadas = detect_json_columns(df)

                    cols_para_expandir = []
                    if json_cols_detectadas:
                        st.markdown(
                            f"🔍 **Colunas JSON detectadas:** `{'`, `'.join(json_cols_detectadas)}`"
                        )
                        cols_para_expandir = st.multiselect(
                            "Selecione quais colunas JSON deseja expandir em colunas separadas:",
                            options=json_cols_detectadas,
                            default=json_cols_detectadas,  # pré-seleciona todas
                        )
                    else:
                        st.info("ℹ️ Nenhuma coluna JSON detectada nesta tabela.")

                    # ── Expandir se solicitado ─────────────────────────────
                    if cols_para_expandir:
                        df_expandido = expand_json_columns(df, cols_para_expandir)
                        st.success(
                            f"✅ JSON expandido! {df.shape[1]} → {df_expandido.shape[1]} colunas"
                        )
                    else:
                        df_expandido = df

                    # ── Preview ───────────────────────────────────────────
                    st.dataframe(df_expandido.head(100), use_container_width=True)

                    # ── Download Excel ────────────────────────────────────
                    xlsx_bytes = df_to_excel_bytes(df_expandido, sheet_name=tbl)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

                    st.download_button(
                        label="📥 Baixar Tabela em Excel",
                        data=xlsx_bytes,
                        file_name=f"3S_{tbl}_{data_inicio}_ate_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Erro ao processar tabela: {e}")

conn.close()
