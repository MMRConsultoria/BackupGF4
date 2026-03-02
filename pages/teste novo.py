import streamlit as st
import pandas as pd
import psycopg2
from datetime import date, timedelta
from io import BytesIO
import uuid

st.set_page_config(layout="wide", page_title="Diagnóstico Meio de Pagamento 3S")

CERT_PATH = "aws-us-east-2-bundle.pem"

def ensure_cert_written():
    if "cert_written_diag" not in st.session_state:
        with open(CERT_PATH, "w", encoding="utf-8") as f:
            f.write(st.secrets["certs"]["aws_rds_us_east_2"])
        st.session_state["cert_written_diag"] = True

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

def _make_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    df_safe = df.copy()
    for col in df_safe.columns:
        if pd.api.types.is_datetime64_any_dtype(df_safe[col]):
            try:
                df_safe[col] = df_safe[col].dt.tz_localize(None)
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

# ========================
ensure_cert_written()

st.title("Diagnóstico: Meio de Pagamento vazio por loja (3S)")

col1, col2, col3 = st.columns(3)
with col1:
    loja = st.text_input("Store code (ex: 0087 ou 87)", value="")
with col2:
    data_inicio = st.date_input("De:", value=date.today() - timedelta(days=30))
with col3:
    data_fim = st.date_input("Até:", value=date.today())

if st.button("🔍 Consultar", type="primary"):
    if not loja.strip():
        st.error("Informe o store_code.")
        st.stop()

    loja_norm = loja.strip().lstrip("0")

    conn = get_db_conn()
    try:
        with st.spinner("Buscando order_picture..."):
            # Todas as colunas de order_picture
            q_op = """
                SELECT *
                FROM public.order_picture
                WHERE business_dt >= %s
                  AND business_dt <= %s
                  AND state_id = 5
            """
            df_op = pd.read_sql(q_op, conn, params=(data_inicio, data_fim))
            df_op["store_code_norm"] = df_op["store_code"].astype(str).str.lstrip("0").str.strip()
            df_op = df_op[df_op["store_code_norm"] == loja_norm].drop(columns=["store_code_norm"])

        st.write("### 📋 order_picture — todas as colunas")
        st.write(f"Linhas: `{len(df_op)}` | Colunas: `{df_op.shape[1]}`")
        st.dataframe(df_op, use_container_width=True, height=400)

        ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "📥 Baixar order_picture (Excel)",
            data=df_to_excel_bytes(df_op, sheet_name="order_picture"),
            file_name=f"order_picture_loja{loja}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_op"
        )

        if df_op.empty:
            st.warning("Nenhum order_picture para essa loja no período.")
            st.stop()

        ids = df_op["order_picture_id"].dropna().astype(int).tolist()

        with st.spinner("Buscando order_picture_tender..."):
            # Todas as colunas de order_picture_tender
            q_t = """
                SELECT *
                FROM public.order_picture_tender
                WHERE order_picture_id = ANY(%s)
            """
            df_t = pd.read_sql(q_t, conn, params=(ids,))

        st.write("### 💳 order_picture_tender — todas as colunas")
        st.write(f"Linhas: `{len(df_t)}` | Colunas: `{df_t.shape[1]}`")
        st.dataframe(df_t, use_container_width=True, height=400)

        st.download_button(
            "📥 Baixar order_picture_tender (Excel)",
            data=df_to_excel_bytes(df_t, sheet_name="order_picture_tender"),
            file_name=f"order_picture_tender_loja{loja}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tender"
        )

        # JOIN completo das duas tabelas
        with st.spinner("Gerando JOIN completo..."):
            df_join = df_t.merge(
                df_op,
                on="order_picture_id",
                how="left",
                suffixes=("_tender", "_op")
            )

        st.write("### 🔗 JOIN completo (order_picture + order_picture_tender)")
        st.write(f"Linhas: `{len(df_join)}` | Colunas: `{df_join.shape[1]}`")
        st.dataframe(df_join, use_container_width=True, height=400)

        st.download_button(
            "📥 Baixar JOIN completo (Excel)",
            data=df_to_excel_bytes(df_join, sheet_name="join_completo"),
            file_name=f"join_completo_loja{loja}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_join"
        )

        # Diagnóstico rápido do campo details
        st.write("### 🔍 Diagnóstico do campo `details`")
        if "details" in df_t.columns:
            nulos = int(df_t["details"].isna().sum())
            vazios = int((df_t["details"].astype(str).str.strip() == "").sum())
            st.write(f"- `details` nulo: **{nulos}**")
            st.write(f"- `details` vazio (string): **{vazios}**")
            st.write(f"- `details` preenchido: **{len(df_t) - nulos - vazios}**")

            st.write("#### Amostras de `details` (RAW)")
            amostra = df_t["details"].dropna().astype(str).head(10).tolist()
            for i, d in enumerate(amostra, start=1):
                st.code(d, language="json")
        else:
            st.warning("Coluna `details` não encontrada em order_picture_tender.")

    except Exception as e:
        st.error(f"Erro: {e}")
    finally:
        conn.close()
