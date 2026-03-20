import streamlit as st
import pandas as pd
import psycopg2
import uuid
from datetime import datetime, date, timedelta
from io import BytesIO

st.set_page_config(page_title="3S - Query Builder", layout="wide")

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

def _make_excel_safe(df):
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
            df_safe[col] = df_safe[col].apply(lambda x: str(x) if isinstance(x, (dict, list, bytes, uuid.UUID)) else x)
    return df_safe

def df_to_excel_bytes(df, sheet_name="data"):
    df_safe = _make_excel_safe(df)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_safe.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()

ensure_cert_written()
st.title("3S (Postgres) — Query Builder")

with st.sidebar:
    st.header("⚙️ Configurações")
    schema = st.text_input("Schema", value="public")

try:
    conn = get_db_conn()
    tables = list_tables(conn, schema=schema)
    st.success(f"Conectado! {len(tables)} tabelas encontradas.")
except Exception as e:
    st.error(f"Erro de conexão: {e}")
    st.stop()

tbl = st.selectbox("1️⃣ Escolha a tabela:", tables)

if tbl:
    df_cols = list_columns(conn, tbl, schema=schema)

    with st.expander("📋 Colunas da tabela (clique para ver)", expanded=True):
        st.dataframe(df_cols, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("2️⃣ Monte sua query")

    modo = st.radio("Modo:", ["RAW (todas as colunas)", "Agregado (SUM por grupo)"], horizontal=True)

    cols_data = [c for c in df_cols["column_name"] if any(x in c.lower() for x in ["date", "dt", "at", "time"])]
    cols_todas = df_cols["column_name"].tolist()
    cols_numericas = df_cols[df_cols["data_type"].str.contains("numeric|integer|double|float", na=False)]["column_name"].tolist()
    cols_texto = df_cols[df_cols["data_type"].str.contains("text|char|varchar", na=False)]["column_name"].tolist()

    col1, col2, col3 = st.columns(3)

    with col1:
        usar_filtro_data = st.checkbox("Filtrar por data?", value=True)
        if usar_filtro_data:
            col_data = st.selectbox("Coluna de data:", cols_data if cols_data else cols_todas)
            data_inicio = st.date_input("De:", value=date.today() - timedelta(days=90))
            data_fim = st.date_input("Até:", value=date.today())

    with col2:
        if modo == "Agregado (SUM por grupo)":
            col_valor = st.selectbox("Coluna para somar (SUM):", cols_numericas if cols_numericas else cols_todas)
            col_grupo = st.selectbox("Agrupar por (GROUP BY):", ["Nenhum"] + cols_texto + cols_todas)

    with col3:
        if modo == "RAW (todas as colunas)":
            limit = st.number_input("Limite de linhas:", min_value=1, max_value=100000, value=5000)

    st.divider()

    if st.button("🚀 Executar Query", type="primary"):
        with st.spinner("Executando..."):
            try:
                params = []
                where = ""

                if usar_filtro_data:
                    where = f'WHERE "{col_data}" >= %s AND "{col_data}" < %s'
                    params += [data_inicio, data_fim + timedelta(days=1)]

                if modo == "RAW (todas as colunas)":
                    order = f'ORDER BY "{col_data}" DESC' if usar_filtro_data else ""
                    q = f'SELECT  FROM "{schema}"."{tbl}" {where} {order} LIMIT %s'
                    params.append(int(limit))

                else:
                    if col_grupo and col_grupo != "Nenhum":
                        q = f'''
                            SELECT
                                COALESCE(CAST("{col_grupo}" AS TEXT), 'TOTAL') AS "{col_grupo}",
                                SUM("{col_valor}") AS total_{col_valor}
                            FROM "{schema}"."{tbl}"
                            {where}
                            GROUP BY ROLLUP ("{col_grupo}")
                            ORDER BY ("{col_grupo}" IS NULL), "{col_grupo}"
                        '''
                    else:
                        q = f'SELECT SUM("{col_valor}") AS total_{col_valor} FROM "{schema}"."{tbl}" {where}'

                df = pd.read_sql(q, conn, params=params if params else None)

                if df.empty:
                    st.warning("Nenhum dado encontrado.")
                else:
                    st.write(f"✅ {len(df)} linhas retornadas.")
                    st.dataframe(df, use_container_width=True)

                    xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)
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
