import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO
from datetime import datetime
import json
import ast

CERT_PATH = "aws-us-east-2-bundle.pem"

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


def _to_dict(x):
    """Converte JSON/str em dict Python."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return {}
    if isinstance(x, dict):
        return x
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return {}
        try:
            return json.loads(s)
        except Exception:
            pass
        try:
            return ast.literal_eval(s)
        except Exception:
            return {}
    return {}


def explode_custom_properties(df: pd.DataFrame, col: str = "custom_properties") -> pd.DataFrame:
    """
    Transforma cada chave do JSON em custom_properties numa coluna separada.
    Mantém o conteúdo original (mesmo se for JSON aninhado).
    """
    df = df.copy()

    if col not in df.columns:
        return df

    # Converte a coluna em dicionários
    parsed = df[col].apply(_to_dict)

    # Normaliza (explode) em colunas
    custom_df = pd.json_normalize(parsed)

    # Junta com o DataFrame original
    df = pd.concat([df, custom_df], axis=1)

    return df


def export_order_picture_to_excel(target_tz: str = "America/Sao_Paulo"):
    conn = get_conn()
    try:
        df = fetch_table_data(conn, "public", "order_picture")

        # Explode da coluna custom_properties
        df = explode_custom_properties(df, col="custom_properties")

        df = sanitize_for_excel(df, target_tz=target_tz)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="public_order_picture", index=False)

        output.seek(0)
        return output, None
    finally:
        conn.close()


# -------------------------
# UI
# -------------------------
st.title("Exportar public.order_picture (custom_properties expandido)")

target_tz = st.selectbox(
    "Fuso horário para datas no Excel",
    options=["America/Sao_Paulo", "UTC"],
    index=0
)

if st.button("Gerar Excel", type="primary", disabled=st.session_state.get("exporting", False)):
    st.session_state["exporting"] = True
    status = st.status("Gerando Excel...", expanded=True)

    try:
        excel_bytes, err = export_order_picture_to_excel(target_tz=target_tz)

        if err:
            status.update(label="Falhou", state="error")
            st.error(err)
        else:
            status.update(label="Concluído", state="complete")
            st.download_button(
                "Baixar Excel",
                data=excel_bytes,
                file_name="public_order_picture_expandido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        status.update(label="Falhou", state="error")
        st.error(f"Erro ao gerar Excel: {e}")
    finally:
        st.session_state["exporting"] = False
