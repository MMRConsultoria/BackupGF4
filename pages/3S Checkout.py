import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo só uma vez por sessão
if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w") as f:
        f.write(st.secrets["certs"]["aws_rds_us_east_2"])
    st.session_state["cert_written"] = True

def get_conn():
    conn = psycopg2.connect(
        host=st.secrets["db"]["host"],
        port=st.secrets["db"]["port"],
        dbname=st.secrets["db"]["database"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        sslmode="verify-full",
        sslrootcert=CERT_PATH,
    )
    return conn

def get_all_tables(conn):
    query = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_type = 'BASE TABLE' AND table_schema NOT IN ('pg_catalog', 'information_schema');
    """
    df = pd.read_sql(query, conn)
    return df

def fetch_table_data(conn, schema, table):
    query = f'SELECT * FROM "{schema}"."{table}"'
    df = pd.read_sql(query, conn)
    return df

st.title("Exportar todas as tabelas do banco para Excel")

if st.button("Gerar Excel"):
    try:
        conn = get_conn()
        tables_df = get_all_tables(conn)
        if tables_df.empty:
            st.warning("Nenhuma tabela encontrada no banco.")
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for idx, row in tables_df.iterrows():
                    schema = row['table_schema']
                    table = row['table_name']
                    st.write(f"Lendo tabela: {schema}.{table}")
                    df = fetch_table_data(conn, schema, table)
                    # Nome da aba: schema_table (máximo 31 caracteres)
                    sheet_name = f"{schema}_{table}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            conn.close()
            output.seek(0)
            st.success("Arquivo Excel gerado com sucesso!")
            st.download_button(
                label="Baixar Excel",
                data=output,
                file_name="banco_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
