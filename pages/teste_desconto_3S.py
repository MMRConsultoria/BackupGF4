# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import psycopg2
import json
import re
from datetime import datetime, timedelta
import unicodedata
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------- Helpers ----------
def _parse_money_to_float(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    # remove símbolos comuns e espaços no meio
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^\d,\-\.]", "", s)
    if s == "":
        return None
    # normaliza separador decimal (tenta lidar com milhares)
    # Se existir '.' e ',' assume que '.' são milhares -> remove '.' e troca ',' por '.'
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
            return None

def _format_brl(v):
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"{v:,.2f}"              # 1,234.56
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")  # 1.234,56
    return f"R$ {s}"

def _strip_accents(text: str) -> str:
    if text is None:
        return ""
    text = str(text)
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

# ---------- Conexões ----------
def create_db_conn_from_secrets():
    try:
        db = st.secrets["db"]
        conn = psycopg2.connect(
            host=db["host"],
            port=db.get("port", 5432),
            dbname=db["database"],
            user=db["user"],
            password=db["password"]
        )
        return conn
    except Exception as e:
        raise RuntimeError(f"Erro criando conexão com DB: {e}")

def create_gspread_client_from_secrets():
    try:
        creds_json = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
        # creds_json pode ser já um dict ou uma string JSON
        if isinstance(creds_json, str):
            creds_dict = json.loads(creds_json)
        else:
            creds_dict = creds_json
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        gc = gspread.authorize(credentials)
        return gc
    except Exception as e:
        raise RuntimeError(f"Erro criando cliente Google Sheets: {e}")

# ---------- Função principal ----------
def atualizar_desconto_3s_checkout(
    data_de=None,
    data_ate=None,
    planilha_nome="Vendas diarias",
    aba_destino_nome="Desconto",
    excluir_stores=("0000", "0001", "9999"),
    estado_filtrar=5
):
    conn = None
    try:
        # 1) abrir conexões
        conn = create_db_conn_from_secrets()
        gc = create_gspread_client_from_secrets()

        # 2) definir período padrão (últimos 30 dias até ontem)
        hoje_utc = datetime.utcnow()
        ontem = (hoje_utc - timedelta(hours=3) - timedelta(days=1)).date()
        if data_ate is None:
            data_ate = ontem
        if data_de is None:
            data_de = (data_ate - timedelta(days=29))
        # aceita string 'YYYY-MM-DD'
        if isinstance(data_de, str):
            data_de = pd.to_datetime(data_de).date()
        if isinstance(data_ate, str):
            data_ate = pd.to_datetime(data_ate).date()

        # 3) query
        sql = """
            SELECT store_code, business_dt, order_discount_amount
            FROM public.order_picture
            WHERE business_dt >= %s
              AND business_dt <= %s
              AND store_code NOT IN %s
              AND state_id = %s
        """
        params = (data_de, data_ate, tuple(excluir_stores), estado_filtrar)
        df = pd.read_sql(sql, conn, params=params)
        if df is None or df.empty:
            return True, "Nenhum registro encontrado no período.", 0

        # 4) processamentos
        # limpar store_code e remover zeros à esquerda
        df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")

        # business_dt -> mm/aaaa
        df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
        df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")

        # order_discount_amount -> numeric -> format BRL
        df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)
        df["order_discount_amount_fmt"] = df["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

        # 5) preparar df de saída com colunas desejadas (B, D, E, F, G, H não aplicam aqui, adaptamos)
        # Como solicitado antes: armazenamos store_code, business_month e valor formatado
        df_out = pd.DataFrame({
            "Store Code": df["store_code"].astype(str),
            "Business Month": df["business_month"],
            "Order Discount Amount": df["order_discount_amount_fmt"]
        })

        # 6) abrir planilha e aba destino (cria se não existir)
        try:
            sh = gc.open(planilha_nome)
        except Exception as e:
            return False, f"Erro ao abrir planilha '{planilha_nome}': {e}", 0

        try:
            ws = sh.worksheet(aba_destino_nome)
        except Exception:
            # cria nova aba com linhas e colunas suficientes
            ws = sh.add_worksheet(title=aba_destino_nome, rows=max(1000, len(df_out)+10), cols=max(3, df_out.shape[1]))

        # 7) tenta mapear cabeçalho existente; se não encontrar, usa padrão
        try:
            header_row = ws.row_values(1)
            header_norm = [ _strip_accents(str(h)).strip().lower() for h in header_row ]
            # tenta encontrar colunas equivalentes
            col_store = None
            col_month = None
            col_discount = None
            for i, hn in enumerate(header_norm):
                if hn and ("loja" in hn or "store" in hn or "codigo" in hn or "store code" in hn):
                    col_store = header_row[i]
                if hn and ("mes" in hn or "mês" in hn or "business" in hn or "data" in hn or "month" in hn):
                    col_month = header_row[i]
                if hn and ("descont" in hn or "discount" in hn or "order_discount" in hn):
                    col_discount = header_row[i]
            if col_store and col_month and col_discount:
                headers_to_write = [col_store, col_month, col_discount]
                df_write = pd.DataFrame({
                    col_store: df_out["Store Code"],
                    col_month: df_out["Business Month"],
                    col_discount: df_out["Order Discount Amount"]
                })
            else:
                headers_to_write = ["Store Code", "Business Month", "Order Discount Amount"]
                df_write = df_out[headers_to_write]
        except Exception:
            headers_to_write = ["Store Code", "Business Month", "Order Discount Amount"]
            df_write = df_out[headers_to_write]

        # 8) escrever (substitui todo o conteúdo da aba)
        send_vals = [headers_to_write] + df_write.fillna("").values.tolist()
        ws.clear()
        ws.update("A1", send_vals, value_input_option="USER_ENTERED")
        linhas = len(df_write)
        return True, f"Aba '{aba_destino_nome}' atualizada com {linhas} linhas.", linhas

    except Exception as e:
        return False, f"Erro interno na atualização: {e}", 0
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass

# ---------- UI Streamlit ----------
st.title("Atualizar aba Desconto (autônomo)")

st.markdown(
    "Este botão executa uma extração de `public.order_picture` e grava o resultado na aba "
    "`Desconto` da planilha especificada (usa credenciais em `st.secrets`)."
)

planilha_input = st.text_input("Nome da planilha (Google Sheets)", value="Vendas diarias")
aba_input = st.text_input("Nome da aba destino", value="Desconto")
dias_default = st.number_input("Últimos quantos dias (padrão 30)", min_value=1, max_value=365, value=30)

if st.button("🔁 Atualizar Desconto (autônomo)"):
    with st.spinner("Executando consulta e atualizando aba..."):
        data_ate = (datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date()
        data_de = data_ate - timedelta(days=int(dias_default) - 1)
        ok, msg, n = atualizar_desconto_3s_checkout(
            data_de=data_de,
            data_ate=data_ate,
            planilha_nome=planilha_input,
            aba_destino_nome=aba_input
        )
        if ok:
            st.success(msg)
        else:
            st.error(msg)
