import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO
from datetime import datetime
import json
import ast

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo só uma vez por sessão
if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w", encoding="utf-8") as f:
        f.write(st.secrets["certs"]["aws_rds_us_east_2"])
    st.session_state["cert_written"] = True

# Inicializa o estado de exportação
if "exporting" not in st.session_state:
    st.session_state["exporting"] = False


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


def sanitize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Remove timezones para compatibilidade com Excel."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)
        elif df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (pd.Timestamp, datetime)) and getattr(x, "tzinfo", None) is not None:
                    return pd.Timestamp(x).tz_localize(None).to_pydatetime()
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
    """Transforma cada chave do JSON em custom_properties numa coluna separada."""
    df = df.copy()
    if col not in df.columns:
        return df

    parsed = df[col].apply(_to_dict)
    custom_df = pd.json_normalize(parsed)
    
    # Garante que os índices batam antes de concatenar
    custom_df.index = df.index
    df = pd.concat([df, custom_df], axis=1)
    return df


def export_order_picture_to_excel():
    conn = get_conn()
    try:
        df = fetch_table_data(conn, "public", "order_picture")

        # Explode da coluna custom_properties
        df = explode_custom_properties(df, col="custom_properties")

        # Limpa datas para o Excel
        df = sanitize_for_excel(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="order_picture", index=False)

        output.seek(0)
        return output, None
    except Exception as e:
        return None, str(e)
    finally:
        conn.close()


# -------------------------
# UI
# -------------------------
st.title("Exportar public.order_picture")
st.subheader("Expansão de custom_properties")

# Botão de reset (caso fique travado)
if st.button("🔄 Resetar Página", type="secondary"):
    st.session_state["exporting"] = False
    st.rerun()

st.write("Clique no botão abaixo para ler o banco e gerar o arquivo Excel.")

if st.button("Gerar Excel", type="primary", disabled=st.session_state["exporting"]):
    st.session_state["exporting"] = True
    status = st.status("Processando dados...", expanded=True)

    try:
        status.write("Conectando ao banco e lendo tabela...")
        excel_bytes, err = export_order_picture_to_excel()

        if err:
            status.update(label="Falhou", state="error")
            st.error(f"Erro no banco: {err}")
        else:
            status.update(label="Concluído ✅", state="complete")
            st.download_button(
                "📥 Baixar Excel",
                data=excel_bytes,
                file_name=f"order_picture_expandido_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        status.update(label="Falhou", state="error")
        st.error(f"Erro inesperado: {e}")
    finally:
        st.session_state["exporting"] = False
