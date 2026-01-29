# streamlit_desconto_gsheet_upload.py
import os
import re
import json
from io import BytesIO
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
    # normaliza separadores
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
            "Credenciais do Google não encontradas em st.secrets['GOOGLE_SERVICE_ACCOUNT']."
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

    # Mantemos o valor numérico para envio ao Sheets
    df_final = pd.DataFrame({
        "3S Checkout": ["3S Checkout"]  len(grouped),
        "Business Month": grouped["business_month"],
        "Loja": grouped["Loja Nome (lookup)"],
        "Grupo": grouped["ColB (lookup)"],
        "Loja Nome": grouped["Loja Nome (lookup)"],
        "Order Discount Amount (BRL)": grouped["order_discount_amount_val"],  # numeric
        "Store Code": grouped["store_code"],
        "Código do Grupo": grouped["Grupo (lookup)"]
    })

    # opcional: se quiser também uma coluna formatada para visualização local/Excel,
    # você pode descomentar a linha abaixo:
    # df_final["Order Discount Amount (Formatted)"] = df_final["Order Discount Amount (BRL)"].apply(lambda x: _format_brl(x))

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
            df_to_write = df.copy()
            # Para Excel, se preferir colunas formatadas, podemos formatar aqui.
            df_to_write.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df_to_write.columns):
                max_len = max(df_to_write[col].astype(str).map(len).max() if not df_to_write.empty else 0, len(col)) + 2
                worksheet.set_column(i, i, max_len)
    except Exception:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ----------------- Google Sheets import (remove existing months + append) -----------------
def upload_df_to_gsheet_replace_months(df: pd.DataFrame,
                                       spreadsheet_name="Vendas diarias",
                                       worksheet_name="Desconto"):
    """
    Remove linhas existentes onde:
      - coluna A == "3S Checkout" e
      - coluna B está em qualquer Business Month presente em df
    Depois cola (append) as linhas de df.
    Garante que colunas numéricas sejam enviadas como números.
    """
    if df is None or df.empty:
        raise ValueError("DataFrame vazio. Nada a importar.")

    gc = create_gspread_client()
    sh = gc.open(spreadsheet_name)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows="1000", cols="20")

    # lê todos os valores existentes
    existing = ws.get_all_values()
    header = existing[0] if existing else df.columns.tolist()

    existing_rows = existing[1:] if len(existing) > 1 else []

    # meses que vamos substituir (strings)
    meses_importar = set(df["Business Month"].astype(str).unique())

    # filtra linhas: mantém as que NÃO correspondem ao padrão (3S Checkout, mês em meses_importar)
    def keep_row(row):
        a = row[0].strip() if len(row) > 0 else ""
        b = row[1].strip() if len(row) > 1 else ""
        if a == "3S Checkout" and b in meses_importar:
            return False
        return True

    filtered_existing = [r for r in existing_rows if keep_row(r)]

    # PREPARA df_rows preservando tipos numéricos
    df_clean = df.copy()
    # substitui NaN por None (vamos lidar depois)
    df_clean = df_clean.where(pd.notnull(df_clean), None)

    # detecta colunas numéricas
    numeric_cols = df_clean.select_dtypes(include=[np.number]).columns.tolist()

    df_rows = []
    for _, row in df_clean.iterrows():
        converted = []
        for col in df_clean.columns:
            val = row[col]
            if val is None:
                # usa string vazia para células vazias
                converted.append("")
            elif col in numeric_cols:
                # converte numpy numerics para python native
                if isinstance(val, (np.integer,)):
                    converted.append(int(val))
                elif isinstance(val, (np.floating,)):
                    converted.append(float(val))
                else:
                    # caso já seja int/float nativo
                    converted.append(val)
            else:
                # para texto, garante string
                converted.append(str(val))
        df_rows.append(converted)

    # constrói o conteúdo final: header + filtered_existing + df_rows
    final_values = [header] + filtered_existing + df_rows

    # atualiza a planilha: limpa e escreve o novo conteúdo
    ws.clear()
    ws.update("A1", final_values)
    return {"kept_rows": len(filtered_existing), "inserted_rows": len(df_rows), "header": header}

# ----------------- Streamlit UI -----------------
st.title("Relatório Desconto — Resumo por Business Month e Store Code")

st.markdown(
    "Gera um relatório Excel resumido e permite importar para Google Sheets substituindo os meses importados."
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
            st.session_state["last_df_final"] = df_final
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        st.exception(e)

if st.button("⬆️ Importar para Google Sheets (substituir meses importados)"):
    df_to_upload = st.session_state.get("last_df_final")
    if df_to_upload is None:
        st.error("Não há relatório gerado. Gere o relatório resumido antes de importar.")
    else:
        try:
            with st.spinner("Importando para Google Sheets..."):
                result = upload_df_to_gsheet_replace_months(df_to_upload,
                                                           spreadsheet_name="Vendas diarias",
                                                           worksheet_name="Desconto")
            st.success(f"Importação concluída. Linhas mantidas: {result['kept_rows']}; Linhas inseridas: {result['inserted_rows']}.")
            st.write("Header usado:", result["header"])
        except Exception as e:
            st.error(f"Erro ao importar para Google Sheets: {e}")
            st.exception(e)
