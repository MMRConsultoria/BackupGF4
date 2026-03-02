import streamlit as st
import pandas as pd
import psycopg2
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="3S - Dump de Tabelas (Raw)", layout="wide")

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
    q = """
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = %s
          AND table_type = 'BASE TABLE'
        ORDER BY table_name
    """
    df = pd.read_sql(q, conn, params=(schema,))
    return df["table_name"].tolist()

def list_columns(conn, table_name, schema="public"):
    q = """
        SELECT column_name, data_type
        FROM information_schema.columns
        WHERE table_schema = %s
          AND table_name = %s
        ORDER BY ordinal_position
    """
    return pd.read_sql(q, conn, params=(schema, table_name))

def fetch_table_limit(conn, table_name, limit=1000, schema="public"):
    # Sem alteração: SELECT direto
    q = f'SELECT * FROM "{schema}"."{table_name}" LIMIT {int(limit)}'
    return pd.read_sql(q, conn)

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()

ensure_cert_written()

st.title("3S (Postgres) — Exportador RAW de Tabelas")

with st.sidebar:
    st.header("Config")
    schema = st.text_input("Schema", value="public")
    limit = st.number_input("LIMIT por tabela", min_value=1, max_value=500000, value=1000, step=1000)
    show_columns = st.checkbox("Mostrar colunas e tipos", value=True)
    show_preview = st.checkbox("Mostrar preview (dataframe)", value=True)

    st.caption("Nada aqui normaliza/limpa dados. É dump bruto com LIMIT.")

# Conecta
try:
    conn = get_db_conn()
except Exception as e:
    st.error(f"Falha ao conectar no banco: {e}")
    st.stop()

with conn:
    try:
        tables = list_tables(conn, schema=schema)
    except Exception as e:
        st.error(f"Falha listando tabelas: {e}")
        st.stop()

st.success(f"Conectado. {len(tables)} tabelas encontradas em `{schema}`.")

# Seleção
tab_mode = st.radio("Modo", ["Selecionar tabela", "Dump em lote (uma por vez)"], horizontal=True)

if tab_mode == "Selecionar tabela":
    tbl = st.selectbox("Tabela", tables)
    col1, col2 = st.columns([1, 1])

    with col1:
        if show_columns:
            st.subheader("Colunas")
            with conn:
                df_cols = list_columns(conn, tbl, schema=schema)
            st.dataframe(df_cols, use_container_width=True)

    with col2:
        st.subheader("Carregar dados RAW")
        if st.button("Carregar", type="primary"):
            with st.spinner(f"Carregando `{tbl}`..."):
                with conn:
                    df = fetch_table_limit(conn, tbl, limit=limit, schema=schema)

            st.write(f"Linhas: {len(df)} | Colunas: {df.shape[1]}")
            if show_preview:
                st.dataframe(df, use_container_width=True, height=520)

            # Downloads
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)

            st.download_button(
                "Baixar CSV",
                data=csv_bytes,
                file_name=f"{tbl}_{ts}.csv",
                mime="text/csv",
            )
            st.download_button(
                "Baixar Excel",
                data=xlsx_bytes,
                file_name=f"{tbl}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.subheader("Dump em lote (interativo)")
    st.write("Vai tabela por tabela, pra não estourar memória. Você escolhe e baixa.")

    selected = st.multiselect("Selecione tabelas", tables, default=tables[:5])
    if st.button("Carregar selecionadas (uma por vez)", type="primary"):
        for tbl in selected:
            st.markdown(f"### {tbl}")
            with conn:
                df = fetch_table_limit(conn, tbl, limit=limit, schema=schema)

            st.write(f"Linhas: {len(df)} | Colunas: {df.shape[1]}")
            if show_columns:
                with conn:
                    df_cols = list_columns(conn, tbl, schema=schema)
                st.dataframe(df_cols, use_container_width=True)

            if show_preview:
                st.dataframe(df, use_container_width=True, height=360)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)

            c1, c2 = st.columns(2)
            c1.download_button(
                "Baixar CSV",
                data=csv_bytes,
                file_name=f"{tbl}_{ts}.csv",
                mime="text/csv",
                key=f"csv_{tbl}_{ts}",
            )
            c2.download_button(
                "Baixar Excel",
                data=xlsx_bytes,
                file_name=f"{tbl}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx_{tbl}_{ts}",
            )

conn.close()
