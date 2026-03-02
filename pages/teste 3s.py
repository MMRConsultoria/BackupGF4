import streamlit as st
import pandas as pd
import psycopg2
import uuid
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="3S - Dump de Tabelas (RAW em Excel)", layout="wide")

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
    q = f'SELECT * FROM "{schema}"."{table_name}" LIMIT {int(limit)}'
    return pd.read_sql(q, conn)

def _make_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte df para um formato 100% compatível com Excel.
    Mantém os valores "raw" o máximo possível, mas faz cast para string
    quando Excel não suporta o tipo.
    """
    df_safe = df.copy()

    for col in df_safe.columns:
        s = df_safe[col]

        # 1) datetime com timezone (Excel não aceita)
        if pd.api.types.is_datetime64_any_dtype(s):
            try:
                # tz-aware -> naive
                df_safe[col] = s.dt.tz_localize(None)
            except Exception:
                try:
                    df_safe[col] = s.dt.tz_convert(None)
                except Exception:
                    # fallback: string
                    df_safe[col] = s.astype(str)

        # 2) timedelta -> string
        if pd.api.types.is_timedelta64_dtype(df_safe[col]):
            df_safe[col] = df_safe[col].astype(str)

        # 3) object: dict/list/bytes/uuid/etc.
        if df_safe[col].dtype == object:
            def _safe_obj(x):
                if x is None:
                    return ""
                if isinstance(x, (dict, list, bytes)):
                    return str(x)
                if isinstance(x, uuid.UUID):
                    return str(x)
                return x

            try:
                df_safe[col] = df_safe[col].apply(_safe_obj)
            except Exception:
                # Se tiver objetos não iteráveis/estranhos, garante string
                df_safe[col] = df_safe[col].astype(str)

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

st.title("3S (Postgres) — Exportador RAW (Somente Excel)")

with st.sidebar:
    st.header("Config")
    schema = st.text_input("Schema", value="public")
    limit = st.number_input("LIMIT por tabela", min_value=1, max_value=500000, value=1000, step=1000)
    show_columns = st.checkbox("Mostrar colunas e tipos", value=True)
    show_preview = st.checkbox("Mostrar preview (dataframe)", value=True)
    st.caption("Exportação RAW em Excel. Alguns tipos são convertidos para string por limitação do Excel.")

# Conecta
try:
    conn = get_db_conn()
except Exception as e:
    st.error(f"Falha ao conectar no banco: {e}")
    st.stop()

try:
    tables = list_tables(conn, schema=schema)
except Exception as e:
    st.error(f"Falha listando tabelas: {e}")
    conn.close()
    st.stop()

st.success(f"Conectado. {len(tables)} tabelas encontradas em `{schema}`.")

tab_mode = st.radio("Modo", ["Selecionar tabela", "Dump em lote"], horizontal=True)

if tab_mode == "Selecionar tabela":
    tbl = st.selectbox("Tabela", tables)

    if show_columns:
        st.subheader("Colunas")
        try:
            df_cols = list_columns(conn, tbl, schema=schema)
            st.dataframe(df_cols, use_container_width=True)
        except Exception as e:
            st.warning(f"Erro ao listar colunas: {e}")

    if st.button("Carregar dados", type="primary"):
        with st.spinner(f"Carregando `{tbl}`..."):
            try:
                df = fetch_table_limit(conn, tbl, limit=limit, schema=schema)
            except Exception as e:
                st.error(f"Erro ao carregar tabela: {e}")
                conn.close()
                st.stop()

        st.write(f"**Linhas:** {len(df)} | **Colunas:** {df.shape[1]}")

        if show_preview:
            st.dataframe(df, use_container_width=True, height=520)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        try:
            xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)
            st.download_button(
                "📥 Baixar Excel",
                data=xlsx_bytes,
                file_name=f"{tbl}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx_{tbl}_{ts}",
            )
        except Exception as e:
            st.error(f"Erro ao gerar Excel: {e}")

else:
    st.subheader("Dump em lote")
    selected = st.multiselect("Selecione as tabelas", tables, default=tables[:3])

    if st.button("Carregar selecionadas", type="primary"):
        for tbl in selected:
            st.markdown(f"### {tbl}")
            with st.spinner(f"Carregando `{tbl}`..."):
                try:
                    df = fetch_table_limit(conn, tbl, limit=limit, schema=schema)
                except Exception as e:
                    st.error(f"Erro em `{tbl}`: {e}")
                    continue

            st.write(f"**Linhas:** {len(df)} | **Colunas:** {df.shape[1]}")

            if show_columns:
                try:
                    df_cols = list_columns(conn, tbl, schema=schema)
                    st.dataframe(df_cols, use_container_width=True)
                except Exception as e:
                    st.warning(f"Erro ao listar colunas de `{tbl}`: {e}")

            if show_preview:
                st.dataframe(df, use_container_width=True, height=360)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")

            try:
                xlsx_bytes = df_to_excel_bytes(df, sheet_name=tbl)
                st.download_button(
                    "📥 Baixar Excel",
                    data=xlsx_bytes,
                    file_name=f"{tbl}_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"xlsx_{tbl}_{ts}",
                )
            except Exception as e:
                st.error(f"Erro ao gerar Excel ({tbl}): {e}")

conn.close()
