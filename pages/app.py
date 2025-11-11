import os
import re
import json
from typing import List, Optional

import pyodbc
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

# Segurança
API_KEY = os.getenv("API_KEY", "mude-esta-chave")
# Conexão ODBC (ajuste o DRIVER conforme seu SO)
ODBC_DRIVER = os.getenv("ODBC_DRIVER", "{MariaDB ODBC 3.2 Driver}")  # Windows
# Para Linux, pode ser algo como: "/usr/lib/x86_64-linux-gnu/odbc/libmaodbc.so"
HOST = os.getenv("DB_HOST", "170.231.15.56")
PORT = int(os.getenv("DB_PORT", "43307"))
DB   = os.getenv("DB_NAME", "C2020147")
USER = os.getenv("DB_USER", "")
PASS = os.getenv("DB_PASS", "")
MAX_ROWS = int(os.getenv("MAX_ROWS", "50000"))
ALLOWED_TABLES_RE = os.getenv("ALLOWED_TABLES_RE", "")  # ex: r"^(vendas|lojas|clientes)$"

def conn_str():
    # Sem DSN, direto por string:
    return f"DRIVER={ODBC_DRIVER};SERVER={HOST};PORT={PORT};DATABASE={DB};UID={USER};PWD={PASS};CHARSET=UTF8"

def get_conn():
    return pyodbc.connect(conn_str(), autocommit=True, timeout=30)

app = FastAPI()

class QueryIn(BaseModel):
    sql: str
    params: Optional[List] = []

def check_auth(x_api_key: Optional[str]):
    if not x_api_key or x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

def is_select_only(sql: str) -> bool:
    s = sql.strip().lower()
    if not s.startswith("select"):
        return False
    forbidden = [" insert ", " update ", " delete ", " drop ", " truncate ", " alter ", " create "]
    return not any(f in s for f in forbidden)

def enforce_allowed_tables(sql: str):
    if not ALLOWED_TABLES_RE:
        return
    pat = re.compile(ALLOWED_TABLES_RE, re.IGNORECASE)
    tables = re.findall(r"(?:from|join)\s+`?([a-zA-Z0-9_\.]+)`?", sql, re.IGNORECASE)
    for t in tables:
        simple = t.split(".")[-1]
        if not pat.match(simple):
            raise HTTPException(status_code=403, detail=f"Tabela não permitida: {t}")

@app.post("/db/query")
def db_query(body: QueryIn, x_api_key: Optional[str] = Header(None)):
    check_auth(x_api_key)
    sql = body.sql

    if not is_select_only(sql):
        raise HTTPException(status_code=400, detail="Apenas SELECT é permitido.")

    enforce_allowed_tables(sql)

    if re.search(r"\blimit\b", sql, re.IGNORECASE) is None:
        sql = f"{sql.strip()} LIMIT {MAX_ROWS}"

    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(sql, body.params or [])
            columns = [c[0] for c in cur.description] if cur.description else []
            rows = cur.fetchall()
            # Normaliza tipos para JSON
            out = []
            for r in rows:
                out.append([None if v is None else str(v) for v in r])
        return {"columns": columns, "rows": out}
    except pyodbc.Error as e:
        raise HTTPException(status_code=502, detail=f"ODBC error: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro: {e}")
