#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import json
from datetime import datetime, timedelta
import pandas as pd
import psycopg2
from psycopg2 import sql

# Optional: try to import streamlit to read st.secrets if running in Streamlit
try:
    import streamlit as st  # type: ignore
    _HAS_ST = True
except Exception:
    _HAS_ST = False

# ----------------- Helpers -----------------
def _parse_money_to_float(x):
    """Tenta converter textos como 'R$ 1.234,56' para float 1234.56"""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s == "":
        return None
    # remove non-numeric except comma, dot and minus
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^\d\-,\.]", "", s)
    # Se houver '.' e ',' assume que '.' são milhares -> remove '.' e troca ',' por '.'
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    # else assume '.' is decimal or no decimal
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return None

def _format_brl(v):
    """Formata número em padrão BRL como string 'R$ 1.234,56'."""
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"{v:,.2f}"               # 1,234.56
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")  # 1.234,56
    return f"R$ {s}"

def _get_db_params_from_env_or_secrets():
    """Tenta obter credenciais do ambiente ou de st.secrets (se disponível)."""
    # Primeiro tenta st.secrets["db"] quando disponível
    if _HAS_ST:
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
            pass
    # Fallback para variáveis de ambiente
    return {
        "host": os.environ.get("PGHOST", "localhost"),
        "port": int(os.environ.get("PGPORT", 5432)),
        "dbname": os.environ.get("PGDATABASE", ""),
        "user": os.environ.get("PGUSER", ""),
        "password": os.environ.get("PGPASSWORD", "")
    }

def create_db_conn(params):
    conn = psycopg2.connect(
        host=params["host"],
        port=params["port"],
        dbname=params["dbname"],
        user=params["user"],
        password=params["password"]
    )
    return conn

# ----------------- Main report generator -----------------
def gerar_relatorio_excel(
    filename: str = "relatorio_desconto.xlsx",
    data_de: str | None = None,   # 'YYYY-MM-DD' ou None
    data_ate: str | None = None,  # 'YYYY-MM-DD' ou None
    dias_default: int = 30,
    excluir_stores: tuple = ("0000", "0001", "9999"),
    estado_filtrar: int = 5
):
    """
    Gera um Excel local com colunas:
      - store_code (sem zeros à esquerda)
      - business_dt (datetime original)
      - business_month (mm/YYYY)
      - order_discount_amount (valor numérico)
      - order_discount_amount_fmt (string 'R$ ...')
    """
    # definir período
    hoje_utc = datetime.utcnow()
    ontem = (hoje_utc - timedelta(hours=3) - timedelta(days=1)).date()
    if data_ate is None:
        data_ate = ontem
    if data_de is None:
        data_de = (pd.to_datetime(data_ate).date() - timedelta(days=dias_default - 1))
    # aceitar strings
    if isinstance(data_de, str):
        data_de = pd.to_datetime(data_de).date()
    if isinstance(data_ate, str):
        data_ate = pd.to_datetime(data_ate).date()

    # query
    sql_query = """
        SELECT store_code, business_dt, order_discount_amount
        FROM public.order_picture
        WHERE business_dt >= %s
          AND business_dt <= %s
          AND store_code NOT IN %s
          AND state_id = %s
        ORDER BY business_dt, store_code
    """

    params = (data_de, data_ate, tuple(excluir_stores), estado_filtrar)

    # conectar e ler
    db_params = _get_db_params_from_env_or_secrets()
    if not db_params["dbname"] or not db_params["user"]:
        raise RuntimeError(
            "Credenciais do banco não encontradas. Defina variáveis de ambiente PGDATABASE/PGUSER/PGPASSWORD/etc "
            "ou forneça st.secrets['db'] quando rodando em Streamlit."
        )

    conn = create_db_conn(db_params)
    try:
        df = pd.read_sql(sql_query, conn, params=params)
    finally:
        conn.close()

    if df is None or df.empty:
        print("Nenhum registro encontrado no período solicitado.")
        # cria excel vazio com cabeçalho
        df_empty = pd.DataFrame(columns=[
            "store_code", "business_dt", "business_month",
            "order_discount_amount", "order_discount_amount_fmt"
        ])
        df_empty.to_excel(filename, index=False)
        print(f"Arquivo gerado: {filename}")
        return filename

    # processamentos
    df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
    df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")

    df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)
    df["order_discount_amount_fmt"] = df["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

    # organizar colunas finais
    df_out = df[[
        "store_code", "business_dt", "business_month",
        "order_discount_amount_val", "order_discount_amount_fmt"
    ]].rename(columns={
        "store_code": "Store Code",
        "business_dt": "Business Date",
        "business_month": "Business Month",
        "order_discount_amount_val": "Order Discount Amount (num)",
        "order_discount_amount_fmt": "Order Discount Amount (BRL)"
    })

    # gravar excel com formatação simples
    try:
        # tenta xlsxwriter para formatar coluna numérica
        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df_out.to_excel(writer, sheet_name="Desconto", index=False, startrow=0)
            workbook = writer.book
            worksheet = writer.sheets["Desconto"]
            # ajustar larguras
            for i, col in enumerate(df_out.columns):
                max_len = max(
                    df_out[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(i, i, max_len)
            # formatar coluna numérica (índice da coluna)
            num_col_idx = df_out.columns.get_loc("Order Discount Amount (num)")
            money_fmt = workbook.add_format({"num_format": "#,##0.00"})
            worksheet.set_column(num_col_idx, num_col_idx, 16, money_fmt)
    except Exception:
        # fallback sem xlsxwriter
        df_out.to_excel(filename, sheet_name="Desconto", index=False)

    print(f"Relatório gerado em: {filename}  (linhas: {len(df_out)})")
    return filename

# ----------------- Execução direta -----------------
if __name__ == "__main__":
    # Exemplo de uso:
    # - Para usar variáveis de ambiente, exporte PGHOST/PGPORT/PGDATABASE/PGUSER/PGPASSWORD antes de rodar
    # - Ou rode dentro do Streamlit com st.secrets['db'] configurado
    try:
        arquivo = gerar_relatorio_excel(filename="relatorio_desconto.xlsx", dias_default=30)
    except Exception as e:
        print("Erro ao gerar relatório:", e)
        raise
