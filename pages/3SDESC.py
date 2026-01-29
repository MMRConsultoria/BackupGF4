import os
import re
import json
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd
import numpy as np
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
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return 0.0

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
            "Credenciais do Google não encontradas em st.secrets['GOOGLE_SERVICE_ACCOUNT']."
        )
    creds_json = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    creds_dict = json.loads(creds_json) if isinstance(creds_json, str) else creds_json
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(credentials)

@st.cache_data(ttl=300)
def fetch_tabela_empresa():
    gc = create_gspread_client()
    sh = gc.open("Vendas diarias")
    ws = sh.worksheet("Tabela Empresa")
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
        for col_name, cond_sql in try_cols:
            sql = f"{base_sql} {cond_sql} ORDER BY business_dt, store_code"
            try:
                df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar, like_void))
                return df
            except Exception as e:
                msg = str(e).lower()
                if "does not exist" in msg or (("column" in msg) and (col_name.lower() in msg)):
                    continue
                else:
                    raise
        sql = f"{base_sql} ORDER BY business_dt, store_code"
        df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar))
        return df
    finally:
        conn.close()

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
    df_final = pd.DataFrame({
        "3S Checkout": ["3S Checkout"] * len(grouped),
        "Business Month": grouped["business_month"],
        "Loja": grouped["Loja Nome (lookup)"],
        "Grupo": grouped["ColB (lookup)"],
        "Loja Nome": grouped["Loja Nome (lookup)"],
        "Order Discount Amount (BRL)": grouped["order_discount_amount_val"],
        "Store Code": grouped["store_code"],
        "Código do Grupo": grouped["Grupo (lookup)"]
    })
    col_order = [
        "3S Checkout", "Business Month", "Loja", "Grupo",
        "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
    ]
    df_final = df_final[col_order]
    return df_final

def upload_df_to_gsheet_replace_months(df: pd.DataFrame,
                                       spreadsheet_name="Vendas diarias",
                                       worksheet_name="Desconto"):
    if df is None or df.empty:
        raise ValueError("DataFrame vazio. Nada a importar.")
    gc = create_gspread_client()
    sh = gc.open(spreadsheet_name)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows="1000", cols="20")
    existing = ws.get_all_values()
    header = existing[0] if existing else df.columns.tolist()
    existing_rows = existing[1:] if len(existing) > 1 else []
    meses_importar = set(df["Business Month"].astype(str).unique())
    def keep_row(row):
        a = row[0].strip() if len(row) > 0 else ""
        b = row[1].strip() if len(row) > 1 else ""
        if a == "3S Checkout" and b in meses_importar:
            return False
        return True
    filtered_existing = [r for r in existing_rows if keep_row(r)]
    df_clean = df.copy()
    df_clean = df_clean.where(pd.notnull(df_clean), None)
    numeric_cols = df_clean.select_dtypes(include=[np.number]).columns.tolist()
    df_rows = []
    for _, row in df_clean.iterrows():
        converted = []
        for col in df_clean.columns:
            val = row[col]
            if val is None:
                converted.append("")
            elif col in numeric_cols:
                if isinstance(val, (np.integer,)):
                    converted.append(int(val))
                elif isinstance(val, (np.floating,)):
                    converted.append(float(val))
                else:
                    converted.append(val)
            else:
                converted.append(str(val))
        df_rows.append(converted)
    final_values = [header] + filtered_existing + df_rows
    ws.clear()
    ws.update("A1", final_values)
    return {"kept_rows": len(filtered_existing), "inserted_rows": len(df_rows), "header": header}

# ----------------- Streamlit UI -----------------
st.title("Atualizar Desconto 3S no Google Sheets")

st.markdown(
    "Clique no botão para buscar os dados, processar e atualizar diretamente a aba 'Desconto' na planilha 'Vendas diarias'."
)

dias_default = st.number_input("Últimos quantos dias", min_value=1, max_value=365, value=30)
data_ate = st.date_input("Data até", value=(datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date())
data_de = st.date_input("Data de", value=(data_ate - timedelta(days=dias_default - 1)))

if st.button("🔄 Atualizar Desconto 3S"):
    try:
        with st.spinner("Buscando Tabela Empresa..."):
            df_empresa = fetch_tabela_empresa()
        with st.spinner("Buscando dados do banco..."):
            df_orders = fetch_order_picture(data_de, data_ate)
        with st.spinner("Processando dados..."):
            df_final = process_and_build_report_summary(df_orders, df_empresa)
        if df_final.empty:
            st.warning("Nenhum dado encontrado para o período selecionado.")
        else:
            with st.spinner("Atualizando Google Sheets..."):
                result = upload_df_to_gsheet_replace_months(df_final,
                                                           spreadsheet_name="Vendas diarias",
                                                           worksheet_name="Desconto")
            st.success(f"Atualização concluída! Linhas mantidas: {result['kept_rows']}; Linhas inseridas: {result['inserted_rows']}.")
    except Exception as e:
        st.error(f"Erro durante a atualização: {e}")
        st.exception(e)
