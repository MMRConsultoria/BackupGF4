import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO
from datetime import datetime

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo só uma vez por sessão
if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w") as f:
        f.write(st.secrets["certs"]["aws_rds_us_east_2"])
    st.session_state["cert_written"] = True


def get_conn():
    return psycopg2.connect(
        host=st.secrets["db"]["host"],
        port=st.secrets["db"]["port"],
        dbname=st.secrets["db"]["database"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        sslmode="verify-full",
        sslrootcert=CERT_PATH,
    )


def get_all_tables(conn):
    query = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_type = 'BASE TABLE'
      AND table_schema NOT IN ('pg_catalog', 'information_schema');
    """
    return pd.read_sql(query, conn)


def fetch_table_data(conn, schema, table):
    query = f'SELECT * FROM "{schema}"."{table}"'
    return pd.read_sql(query, conn)


def sanitize_for_excel(df: pd.DataFrame, target_tz: str = "UTC") -> pd.DataFrame:
    """
    Excel não suporta datetimes com timezone.
    Converte colunas com timezone para target_tz e remove o tz.
    """
    df = df.copy()

    for col in df.columns:
        # Coluna datetime com timezone (datetime64[ns, tz])
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_convert(target_tz).dt.tz_localize(None)

        # Caso raro: dtype object com datetimes tz-aware
        elif df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (pd.Timestamp, datetime)) and getattr(x, "tzinfo", None) is not None:
                    ts = pd.Timestamp(x).tz_convert(target_tz)
                    return ts.tz_localize(None).to_pydatetime()
                return x

            df[col] = df[col].map(_fix)

    return df


st.title("Exportar todas as tabelas do banco para Excel")

# Você pode trocar para "UTC" se preferir
TARGET_TZ = "America/Sao_Paulo"

if st.button("Gerar Excel"):
    conn = None
    try:
        conn = get_conn()
        tables_df = get_all_tables(conn)

        if tables_df.empty:
            st.warning("Nenhuma tabela encontrada no banco.")
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for _, row in tables_df.iterrows():
                    schema = row["table_schema"]
                    table = row["table_name"]

                    st.write(f"Lendo tabela: {schema}.{table}")
                    df = fetch_table_data(conn, schema, table)

                    # Corrige datetimes com timezone para Excel
                    df = sanitize_for_excel(df, target_tz=TARGET_TZ)

                    # Nome da aba: máximo 31 caracteres
                    sheet_name = f"{schema}_{table}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            output.seek(0)
            st.success("Arquivo Excel gerado com sucesso!")

            st.download_button(
                label="Baixar Excel",
                data=output,
                file_name="banco_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")

    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
