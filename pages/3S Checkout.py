import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO
from datetime import datetime

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo só uma vez por sessão
if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w", encoding="utf-8") as f:
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
      AND table_schema NOT IN ('pg_catalog', 'information_schema')
    ORDER BY table_schema, table_name;
    """
    return pd.read_sql(query, conn)


def fetch_table_data(conn, schema, table):
    query = f'SELECT * FROM "{schema}"."{table}"'
    return pd.read_sql(query, conn)


def sanitize_for_excel(df: pd.DataFrame, target_tz: str = "America/Sao_Paulo") -> pd.DataFrame:
    df = df.copy()

    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_convert(target_tz).dt.tz_localize(None)

        elif df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (pd.Timestamp, datetime)) and getattr(x, "tzinfo", None) is not None:
                    ts = pd.Timestamp(x).tz_convert(target_tz)
                    return ts.tz_localize(None).to_pydatetime()
                return x
            df[col] = df[col].map(_fix)

    return df


def export_db_to_excel(target_tz: str = "America/Sao_Paulo"):
    conn = get_conn()
    try:
        tables_df = get_all_tables(conn)
        if tables_df.empty:
            return None, "Nenhuma tabela encontrada no banco."

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for _, row in tables_df.iterrows():
                schema = row["table_schema"]
                table = row["table_name"]

                df = fetch_table_data(conn, schema, table)
                df = sanitize_for_excel(df, target_tz=target_tz)

                sheet_name = f"{schema}_{table}"[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)
        return output, None
    finally:
        conn.close()


st.title("Exportar banco para Excel")

target_tz = st.selectbox(
    "Fuso horário para datas no Excel",
    options=["America/Sao_Paulo", "UTC"],
    index=0
)

# Evita clique duplo / reruns durante export
if st.button("Gerar Excel", type="primary", disabled=st.session_state.get("exporting", False)):
    st.session_state["exporting"] = True

    status = st.status("Gerando Excel... (isso pode demorar)", expanded=True)
    try:
        status.write("Conectando ao banco e lendo tabelas...")
        progress = st.progress(0)

        # Faz a exportação (sem criar st.write a cada tabela)
        excel_bytes, err = export_db_to_excel(target_tz=target_tz)

        progress.progress(100)

        if err:
            status.update(label="Falhou", state="error")
            st.error(err)
        else:
            status.update(label="Concluído", state="complete")
            st.download_button(
                "Baixar Excel",
                data=excel_bytes,
                file_name="banco_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        status.update(label="Falhou", state="error")
        st.error(f"Erro ao gerar Excel: {e}")

    finally:
        st.session_state["exporting"] = False
