import streamlit as st
import pandas as pd
import psycopg2
import uuid
import json
import ast
from datetime import datetime, date, timedelta
from io import BytesIO

st.set_page_config(page_title="3S - Dump RAW com Filtro", layout="wide")

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

def get_unique_stores(conn, table_name, schema="public"):
    """Busca códigos de loja únicos para o filtro."""
    try:
        q = f'SELECT DISTINCT "store_code" FROM "{schema}"."{table_name}" ORDER BY "store_code"'
        return pd.read_sql(q, conn)["store_code"].tolist()
    except:
        return []

def fetch_table_filtered(conn, table_name, date_col, start_date, end_date, stores, limit=1000, schema="public"):
    """Busca dados com filtros de Data e Store Code."""
    where_clauses = []
    params = []

    # Filtro de Data
    if date_col and date_col != "Sem filtro de data":
        where_clauses.append(f'"{date_col}" >= %s')
        params.append(start_date)
        where_clauses.append(f'"{date_col}" <= %s')
        params.append(end_date)

    # Filtro de Lojas
    if stores:
        where_clauses.append(f'"store_code" IN %s')
        params.append(tuple(stores))

    where_stmt = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""
    
    order_stmt = f' ORDER BY "{date_col}" DESC' if (date_col and date_col != "Sem filtro de data") else ""
    
    q = f'SELECT * FROM "{schema}"."{table_name}"{where_stmt}{order_stmt} LIMIT %s'
    params.append(int(limit))
    
    return pd.read_sql(q, conn, params=params)

def _make_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    df_safe = df.copy()
    for col in df_safe.columns:
        if pd.api.types.is_datetime64_any_dtype(df_safe[col]):
            try: df_safe[col] = df_safe[col].dt.tz_localize(None)
            except:
                try: df_safe[col] = df_safe[col].dt.tz_convert(None)
                except: df_safe[col] = df_safe[col].astype(str)
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

def detect_json_columns(df: pd.DataFrame) -> list:
    json_cols = []
    for col in df.columns:
        if df[col].dtype != object: continue
        sample = df[col].dropna().head(20)
        hits = 0
        for val in sample:
            if isinstance(val, dict): hits += 1; continue
            if isinstance(val, str) and val.strip().startswith(("{", "[")):
                try: json.loads(val); hits += 1
                except: pass
        if hits >= max(1, len(sample) * 0.3): json_cols.append(col)
    return json_cols

def expand_json_columns(df: pd.DataFrame, cols_to_expand: list) -> pd.DataFrame:
    df_result = df.copy()
    for col in cols_to_expand:
        if col not in df_result.columns: continue
        def safe_parse(x):
            if pd.isna(x) or x == "": return {}
            try:
                if isinstance(x, dict): return x
                if isinstance(x, str): return json.loads(x)
            except:
                try: return ast.literal_eval(x)
                except: return {}
            return {}
        parsed = df_result[col].apply(safe_parse)
        if parsed.apply(lambda x: bool(x)).any():
            expanded = pd.json_normalize(parsed)
            expanded.columns = [f"{col}__{c}" for c in expanded.columns]
            expanded.index = df_result.index
            pos = df_result.columns.get_loc(col) + 1
            for i, new_col in enumerate(expanded.columns):
                df_result.insert(pos + i, new_col, expanded[new_col])
    return df_result

# ── INÍCIO DO APP ─────────────────────────────────────────────────────────────

ensure_cert_written()
st.title("3S (Postgres) — Exportador RAW com Filtro de Loja")

# Sidebar
with st.sidebar:
    st.header("📅 Filtro de Data")
    data_inicio = st.date_input("Data início", value=date.today() - timedelta(days=7))
    data_fim = st.date_input("Data fim", value=date.today())
    if data_fim < data_inicio:
        st.error("⚠️ Data fim deve ser maior ou igual à data início.")

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
tbl = st.selectbox("Escolha a tabela:", tables)

if tbl:
    # 1. Buscar colunas e lojas únicas
    df_cols = list_columns(conn, tbl, schema=schema)
    lista_lojas = get_unique_stores(conn, tbl, schema=schema)
    
    # 2. Configurar Filtros na UI
    col1, col2 = st.columns(2)
    
    with col1:
        cols_data = [c for c in df_cols["column_name"] if any(x in c.lower() for x in ["date", "dt", "at", "time"])]
        col_data_escolhida = st.selectbox("Coluna de data:", ["Sem filtro de data"] + cols_data)
    
    with col2:
        if lista_lojas:
            lojas_selecionadas = st.multiselect("Filtrar por Store Code (Vazio = Todas):", options=lista_lojas)
        else:
            st.info("Coluna 'store_code' não encontrada nesta tabela.")
            lojas_selecionadas = []

    # 3. Botão de Ação
    if st.button("🚀 Carregar e Processar", type="primary"):
        with st.spinner("Consultando banco de dados..."):
            try:
                df = fetch_table_filtered(
                    conn, tbl, col_data_escolhida, data_inicio, data_fim,
                    lojas_selecionadas, limit=limit_global, schema=schema
                )

                if df.empty:
                    st.warning("Nenhum dado encontrado para os filtros aplicados.")
                else:
                    st.write(f"✅ {len(df)} linhas carregadas.")

                    # Expansão JSON
                    json_cols = detect_json_columns(df)
                    if json_cols:
                        st.markdown(f"🔍 **JSON detectado em:** `{'`, `'.join(json_cols)}`")
                        cols_expand = st.multiselect("Expandir colunas JSON:", json_cols, default=json_cols)
                        df_final = expand_json_columns(df, cols_expand)
                    else:
                        df_final = df

                    st.dataframe(df_final.head(100), use_container_width=True)

                    # Download
                    xlsx_bytes = df_to_excel_bytes(df_final, sheet_name=tbl)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📥 Baixar Excel",
                        data=xlsx_bytes,
                        file_name=f"3S_{tbl}_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Erro: {e}")

conn.close()


Tenho esse codigo quero poder escolher o periodo que quero cada tabela
