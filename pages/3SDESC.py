# streamlit_relatorio_desconto_summary.py
import os
import re
import json
from io import BytesIO
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd
import psycopg2
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ----------------- Helpers -----------------
def _parse_money_to_float(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^\d,\-\.]", "", s)
    if s == "":
        return 0.0
    # normaliza separadores
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",", ) == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return 0.0

def _format_brl(v):
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def _get_db_params():
    try:
        db = st.secrets["db"]
        return {
            "host": db["host"],
            "port": int(db.get("port", 5432)),
            "dbname": db["database"],
            "user": db["user"],
            "password": db["password"]
        }
    except Exception:
        return {
            "host": os.environ.get("PGHOST", "localhost"),
            "port": int(os.environ.get("PGPORT", 5432)),
            "dbname": os.environ.get("PGDATABASE", ""),
            "user": os.environ.get("PGUSER", ""),
            "password": os.environ.get("PGPASSWORD", "")
        }

def create_db_conn(params):
    return psycopg2.connect(
        host=params["host"],
        port=params["port"],
        dbname=params["dbname"],
        user=params["user"],
        password=params["password"]
    )

def create_gspread_client():
    if "GOOGLE_SERVICE_ACCOUNT" not in st.secrets:
        raise RuntimeError(
            "Google credentials not found in st.secrets['GOOGLE_SERVICE_ACCOUNT']."
        )
    creds_json = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    creds_dict = json.loads(creds_json) if isinstance(creds_json, str) else creds_json
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(credentials)

# ----------------- Fetch / Cache -----------------
@st.cache_data(ttl=300)
def fetch_tabela_empresa():
    gc = create_gspread_client()
    try:
        sh = gc.open("Vendas diarias")
    except Exception as e:
        raise RuntimeError(f"Erro ao abrir a planilha 'Vendas diarias': {e}")
    try:
        ws = sh.worksheet("Tabela Empresa")
    except Exception as e:
        raise RuntimeError(f"Erro ao abrir a aba 'Tabela Empresa': {e}")

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    max_cols = max(len(r) for r in values)
    rows = [r + [""] * (max_cols - len(r)) for r in values]
    cols = [chr(ord("A") + i) for i in range(max_cols)]
    data_rows = rows[1:] if len(rows) > 1 else []
    df = pd.DataFrame(data_rows, columns=cols)
    df = df.loc[~(df[cols].apply(lambda r: all(str(x).strip() == "" for x in r), axis=1))]
    return df

@st.cache_data(ttl=300)
def fetch_order_picture(data_de, data_ate, excluir_stores=("0000", "0001", "9999"), estado_filtrar=5):
    params = _get_db_params()
    if not params["dbname"] or not params["user"] or not params["password"]:
        raise RuntimeError("Credenciais do banco nao encontradas. Configure st.secrets['db'] ou variaveis de ambiente PG*.")

    conn = create_db_conn(params)
    try:
        base_sql = """
            SELECT store_code, business_dt, order_discount_amount
            FROM public.order_picture
            WHERE business_dt >= %s
              AND business_dt <= %s
              AND store_code NOT IN %s
              AND state_id = %s
        """
        try_cols = [
            ("VOID_TYPE", "AND (VOID_TYPE IS NULL OR VOID_TYPE = '' OR LOWER(VOID_TYPE) NOT LIKE %s)"),
            ("pod_type", "AND (pod_type IS NULL OR pod_type = '' OR LOWER(pod_type) NOT LIKE %s)")
        ]
        like_void = "%void%"

        last_exc = None
        for col_name, cond_sql in try_cols:
            sql = f"{base_sql} {cond_sql} ORDER BY business_dt, store_code"
            try:
                df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar, like_void))
                return df
            except Exception as e:
                last_exc = e
                msg = str(e).lower()
                if "does not exist" in msg or (("column" in msg) and (col_name.lower() in msg)):
                    continue
                else:
                    raise

        # fallback: search without void/pod filter
        sql = f"{base_sql} ORDER BY business_dt, store_code"
        df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar))
        return df
    finally:
        conn.close()

# ----------------- Processing / Summary -----------------
def process_and_build_report_summary(df_orders: pd.DataFrame, df_empresa: pd.DataFrame) -> pd.DataFrame:
    if df_orders is None or df_orders.empty:
        return pd.DataFrame(columns=[
            "3S Checkout", "Business Month", "Loja", "Grupo",
            "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
        ])

    df = df_orders.copy()
    df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
    df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")
    df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)

    if df_empresa is None or df_empresa.empty:
        mapa_codigo_para_nome = {}
        mapa_codigo_para_colB = {}
        mapa_codigo_para_grupo = {}
    else:
        for col in ["A", "B", "C", "D"]:
            if col not in df_empresa.columns:
                df_empresa[col] = ""
        codigo_col = df_empresa["C"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
        mapa_codigo_para_nome = dict(zip(codigo_col, df_empresa["A"].astype(str)))
        mapa_codigo_para_colB = dict(zip(codigo_col, df_empresa["B"].astype(str)))
        mapa_codigo_para_grupo = dict(zip(codigo_col, df_empresa["D"].astype(str)))

    df["Loja Nome (lookup)"] = df["store_code"].map(mapa_codigo_para_nome)
    df["ColB (lookup)"] = df["store_code"].map(mapa_codigo_para_colB)
    df["Grupo (lookup)"] = df["store_code"].map(mapa_codigo_para_grupo)

    grouped = df.groupby(["business_month", "store_code"], as_index=False).agg({
        "order_discount_amount_val": "sum",
        "Loja Nome (lookup)": "first",
        "ColB (lookup)": "first",
        "Grupo (lookup)": "first"
    })

    grouped["order_discount_amount_fmt"] = grouped["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

    df_final = pd.DataFrame({
        "3S Checkout": ["3S Checkout"] * len(grouped),
        "Business Month": grouped["business_month"],
        "Loja": grouped["Loja Nome (lookup)"],
        "Grupo": grouped["ColB (lookup)"],
        "Loja Nome": grouped["Loja Nome (lookup)"],
        "Order Discount Amount (BRL)": grouped["order_discount_amount_fmt"],
        "Store Code": grouped["store_code"],
        "Código do Grupo": grouped["Grupo (lookup)"]
    })

    col_order = [
        "3S Checkout", "Business Month", "Loja", "Grupo",
        "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
    ]
    df_final = df_final[col_order]
    return df_final

def to_excel_bytes(df: pd.DataFrame, sheet_name="Desconto"):
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max() if not df.empty else 0, len(col)) + 2
                worksheet.set_column(i, i, max_len)
    except Exception:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ----------------- Streamlit UI -----------------
st.title("Relatório Desconto — Resumo por Business Month e Store Code")

st.markdown(
    "Gera um relatório Excel resumido (agregado) por Business Month e Store Code. "
    "Agrupa somando os descontos e mantém nome da loja / grupo conforme Tabela Empresa."
)

dias_default = st.number_input("Últimos quantos dias", min_value=1, max_value=365, value=30)
data_ate = st.date_input("Data até", value=(datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date())
data_de = st.date_input("Data de", value=(data_ate - timedelta(days=dias_default - 1)))

nome_arquivo = st.text_input("Nome do arquivo para download", value="relatorio_desconto_resumido.xlsx")

if st.button("🔁 Gerar relatório resumido"):
    try:
        with st.spinner("Buscando Tabela Empresa (Google Sheets)..."):
            df_empresa = fetch_tabela_empresa()
        with st.spinner("Buscando dados do banco..."):
            df_orders = fetch_order_picture(data_de, data_ate)
        with st.spinner("Processando resumo..."):
            df_final = process_and_build_report_summary(df_orders, df_empresa)

        if df_final.empty:
            st.warning("Nenhum dado encontrado para o período selecionado.")
        else:
            st.dataframe(df_final.head(200))
            excel_bytes = to_excel_bytes(df_final, sheet_name="Desconto")
            st.download_button(
                label="⬇️ Baixar relatório Excel (resumido)",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        st.exception(e)
