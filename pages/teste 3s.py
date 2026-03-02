import streamlit as st
import pandas as pd
import psycopg2
import uuid
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
    # Se houver coluna de data, filtra. Se não, faz select normal com limit.
    if date_col and date_col != "Sem filtro de data":
        q = f'SELECT * FROM "{schema}"."{table_name}" WHERE "{date_col}" >= %s ORDER BY "{date_col}" DESC LIMIT %s'
        return pd.read_sql(q, conn, params=(start_date, int(limit)))
    else:
        q = f'SELECT * FROM "{schema}"."{table_name}" LIMIT %s'
        return pd.read_sql(q, conn, params=(int(limit),))

def _make_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    df_safe = df.copy()
    for col in df_safe.columns:
        # 1) Datetime com timezone -> Naive
        if pd.api.types.is_datetime64_any_dtype(df_safe[col]):
            try:
                df_safe[col] = df_safe[col].dt.tz_localize(None)
            except:
                try: df_safe[col] = df_safe[col].dt.tz_convert(None)
                except: df_safe[col] = df_safe[col].astype(str)
        # 2) Timedelta -> String
        elif pd.api.types.is_timedelta64_dtype(df_safe[col]):
            df_safe[col] = df_safe[col].astype(str)
        # 3) Objetos (Dict, List, UUID) -> String
        elif df_safe[col].dtype == object:
            df_safe[col] = df_safe[col].apply(lambda x: str(x) if isinstance(x, (dict, list, bytes, uuid.UUID)) else x)
    return df_safe

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="data"):
    df_safe = _make_excel_safe(df)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_safe.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()

# ========================
ensure_cert_written()
st.title("3S (Postgres) — Exportador RAW Excel com Filtro")

# Sidebar de Configuração Global
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
    # Busca colunas para identificar qual é a de data
    df_cols = list_columns(conn, tbl, schema=schema)
    
    # Tenta sugerir colunas de data (que contenham 'date', 'dt', 'at', 'time')
    cols_data_sugeridas = [c for c in df_cols["column_name"] if any(x in c.lower() for x in ["date", "dt", "at", "time"])]
    cols_data_sugeridas = ["Sem filtro de data"] + cols_data_sugeridas
    
    col1, col2 = st.columns([1, 2])
    with col1:
        col_data_escolhida = st.selectbox("Coluna de data para filtrar:", cols_data_sugeridas)
    
    with col2:
        st.write(f"**Filtro:** `{col_data_escolhida}` >= `{data_inicio}`")

    if st.button("🚀 Carregar e Gerar Excel", type="primary"):
        with st.spinner(f"Buscando dados de `{tbl}`..."):
            try:
                df = fetch_table_with_date(conn, tbl, col_data_escolhida, data_inicio, limit=limit_global, schema=schema)
                
                if df.empty:
                    st.warning("Nenhum dado encontrado para este período.")
                else:
                    st.write(f"✅ {len(df)} linhas carregadas.")
                    st.dataframe(df.head(100), use_container_width=True)
                    
                    # Gerar Excel
                    xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    st.download_button(
                        label="📥 Baixar Tabela em Excel",
                        data=xlsx_bytes,
                        file_name=f"3S_{tbl}_{data_inicio}_ate_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Erro ao processar tabela: {e}")

conn.close()
