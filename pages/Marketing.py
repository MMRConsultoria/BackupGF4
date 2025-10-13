import streamlit as st
import pandas as pd
import numpy as np
import re, unicodedata, json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Importar Materiais por Loja", layout="wide")
st.title("📥 Importar Materiais por Loja (com Tabela Empresa)")

# ---------------- Helpers ----------------
def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def normalizar_loja(txt: str) -> str:
    s = str(txt or "").strip()
    s = re.sub(r"^\s*\d+\s*-\s*", "", s)  # remove "123 - "
    return s.strip()

def pick_name(cols, targets):
    m = {_ns(c): c for c in cols}
    for t in targets:
        if _ns(t) in m:
            return m[_ns(t)]
    return None

def _parse_brl(x) -> float:
    s = str(x or "").strip()
    if s == "":
        return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s.count(",") >= 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") >= 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        v = float(s)
        return -v if neg else v
    except:
        return np.nan

def _fmt_brl(v) -> str:
    try:
        v = float(v)
    except:
        return "R$ 0,00"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s

# --------- Google Sheets: Tabela Empresa ----------
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    key = "GOOGLE_SERVICE_ACCOUNT" if "GOOGLE_SERVICE_ACCOUNT" in st.secrets else (
        "gcp_service_account" if "gcp_service_account" in st.secrets else None
    )
    if key is None:
        raise RuntimeError("Configure st.secrets['GOOGLE_SERVICE_ACCOUNT'] (ou 'gcp_service_account').")

    creds_any = st.secrets[key]
    creds_dict = json.loads(creds_any) if isinstance(creds_any, str) else creds_any
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(credentials)

    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])

    cols = df.columns.tolist()
    col_loja  = pick_name(cols, ["Loja"])
    col_grupo = pick_name(cols, ["Grupo","Operação"])
    col_cod   = pick_name(cols, ["Código Everest","Codigo Everest","Cod Everest"])
    col_codg  = pick_name(cols, ["Código Grupo Everest","Codigo Grupo Everest","Cod Grupo Empresas","Código Grupo Empresas"])

    out = pd.DataFrame()
    out["Loja"] = df[col_loja].astype(str).str.strip() if col_loja else ""
    out["Loja_norm"] = out["Loja"].map(lambda x: x.strip().lower())
    out["Grupo"] = df[col_grupo].astype(str).str.strip() if col_grupo else ""
    out["Código Everest"] = pd.to_numeric(df[col_cod], errors="coerce") if col_cod else pd.NA
    out["Código Grupo Everest"] = pd.to_numeric(df[col_codg], errors="coerce") if col_codg else pd.NA
    return out

# --------- Parser do Excel de Upload ----------
def ler_relatorio(uploaded_file) -> pd.DataFrame:
    df0 = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
    if df0.shape[0] < 6:
        return pd.DataFrame()

    ROW_LOJA = 3   # linha 4 (0-based)
    ROW_HDR  = 4   # linha 5 (0-based)
    COL_B, COL_C, COL_D = 1, 2, 3  # Grupo, Código, Material

    r5 = df0.iloc[ROW_HDR].astype(str).fillna("")
    r5n = r5.map(_ns)

    lojas_row = df0.iloc[ROW_LOJA].astype(str)
    lojas_row = lojas_row.replace(["", "nan", "None", "NaN"], pd.NA)
    lojas_row_ff = lojas_row.ffill()

    def eh_qtde(tok: str) -> bool:
        return _ns(tok) == "qtde"

    VAL_TOKS = {"valor(r$)", "valor r$", "valor (r$)", "valor(r$ )", "valor r$)", "valor"}
    def eh_valor(tok: str) -> bool:
        return _ns(tok) in VAL_TOKS

    pairs = []  # (col_qt, col_vl, loja_name)
    j = 0
    ncols = df0.shape[1]
    while j < ncols:
        if eh_qtde(r5n.iloc[j]):
            vcol = None
            limite = min(ncols, j + 4)
            k = j + 1
            while k < limite and not eh_qtde(r5n.iloc[k]):
                if eh_valor(r5n.iloc[k]):
                    vcol = k
                    break
                k += 1
            if vcol is not None:
                loja_bruta = str(lojas_row_ff.iloc[j] if j < len(lojas_row_ff) else "").strip()
                loja_bruta = normalizar_loja(loja_bruta)
                if loja_bruta and "total" not in _ns(loja_bruta):
                    pairs.append((j, vcol, loja_bruta))
                j = vcol + 1
                continue
        j += 1

    base = df0.iloc[ROW_HDR+1:].copy()
    base = base.rename(columns={COL_B: "GrupoColB", COL_C: "Codigo", COL_D: "Material"})

    base["GrupoProduto"] = (
        base["GrupoColB"]
        .where(base["GrupoColB"].notna() & (base["GrupoColB"].astype(str).str.strip() != ""), np.nan)
        .ffill().astype(str).str.strip()
    )
    base["Material"] = base["Material"].astype(str).str.strip()
    base["Codigo"] = base["Codigo"].where(base["Codigo"].astype(str).str.strip() != "", np.nan).ffill()
    base["Codigo"] = base["Codigo"].astype(str).str.strip()

    if base.empty or not pairs:
        return pd.DataFrame(columns=["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"])

    registros = []
    for c_q, c_v, loja_nome in pairs:
        sub = base[["GrupoProduto","Codigo","Material", c_q, c_v]].copy()
        sub = sub.rename(columns={c_q: "Qtde", c_v: "Valor"})

        mask_total = sub["Codigo"].astype(str).str.contains(r"\btotal\b", case=False, na=False) | \
                     sub["Material"].astype(str).str.contains(r"\btotal\b", case=False, na=False)
        sub = sub[~mask_total]

        sub["Qtde"] = pd.to_numeric(sub["Qtde"], errors="coerce")
        sub = sub[sub["Qtde"].notna()]

        sub["Valor"] = sub["Valor"].apply(_parse_brl)
        sub["Valor"] = pd.to_numeric(sub["Valor"], errors="coerce").fillna(0.0)
        sub = sub[sub["Valor"] > 0]

        sub["Loja"] = loja_nome
        if not sub.empty:
            registros.append(sub)

    if not registros:
        return pd.DataFrame(columns=["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"])

    out = pd.concat(registros, ignore_index=True)
    out = out[["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"]].copy()
    return out

# --------------- UI ----------------
c1, c2 = st.columns(2)
with c1:
    nome_planilha = st.text_input("Planilha no Google Sheets", value="Vendas diarias")
with c2:
    aba_empresa = st.text_input("Aba Tabela Empresa", value="Tabela Empresa")

up = st.file_uploader("Envie o Excel (linha 4 = lojas, linha 5 = cabeçalhos Qtde/Valor)", type=["xlsx","xls"])

if not up:
    st.info("Envie o arquivo para começar.")
    st.stop()

# Ler Excel
df_items = ler_relatorio(up)
if df_items.empty:
    st.warning("Nenhum item elegível foi encontrado (verifique valores/qtde e colunas Qtde/Valor).")
    st.stop()

# Tabela Empresa
try:
    df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
except Exception as e:
    st.error(f"❌ Erro ao carregar Tabela Empresa: {e}")
    st.stop()

# -------- Join por loja normalizada (padrão Tabela Empresa) --------
df_emp["Loja"] = df_emp["Loja"].astype(str).str.strip()
df_emp["Loja_norm"] = df_emp["Loja"].str.lower()

df_items["Loja"] = df_items["Loja"].astype(str).str.strip()
df_items["Loja"] = df_items["Loja"].str.replace(r"^\s*\d+\s*-\s*", "", regex=True)
df_items["Loja_norm"] = df_items["Loja"].str.lower()

merged = df_items.merge(
    df_emp[["Loja_norm","Loja","Grupo","Código Everest","Código Grupo Everest"]],
    on="Loja_norm", how="left"
)

if "Loja_x" in merged.columns and "Loja_y" in merged.columns:
    merged["Loja"] = merged["Loja_y"].where(merged["Loja_y"].astype(str).str.strip() != "", merged["Loja_x"])
    merged.drop(columns=["Loja_x", "Loja_y"], inplace=True)
elif "Loja_y" in merged.columns:
    merged["Loja"] = merged["Loja_y"]; merged.drop(columns=["Loja_y"], inplace=True)
elif "Loja_x" in merged.columns:
    merged["Loja"] = merged["Loja_x"]; merged.drop(columns=["Loja_x"], inplace=True)

merged = merged.rename(columns={
    "Grupo": "Operação",
    "GrupoProduto": "Grupo Material",
    "Codigo": "Codigo Material",
})

final_cols = [
    "Operação","Loja","Código Everest","Código Grupo Everest",
    "Grupo Material","Codigo Material","Material","Qtde","Valor"
]
for c in final_cols:
    if c not in merged.columns:
        merged[c] = ""

df_final = merged[final_cols].copy()

# ----- TOTAL numérico -----
linha_total = {
    "Operação": "TOTAL",
    "Loja": "",
    "Código Everest": "",
    "Código Grupo Everest": "",
    "Grupo Material": "",
    "Codigo Material": "",
    "Material": "",
    "Qtde": pd.to_numeric(df_final["Qtde"], errors="coerce").sum(skipna=True),
    "Valor": pd.to_numeric(df_final["Valor"], errors="coerce").sum(skipna=True),
}
df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)

# --- Prévia (formata só na tela) ---
df_view = df_final.copy()
df_view["Valor"] = df_view["Valor"].apply(_fmt_brl)
st.subheader("Prévia")
st.dataframe(df_view.head(120), use_container_width=True, hide_index=True)
st.caption(f"Linhas (com TOTAL): {len(df_view):,}".replace(",", "."))

# -------- Download: mantém número e aplica formato no Excel --------
def to_excel(df_num: pd.DataFrame):
    # garante numéricos
    df_exp = df_num.copy()
    df_exp["Qtde"] = pd.to_numeric(df_exp["Qtde"], errors="coerce")
    df_exp["Valor"] = pd.to_numeric(df_exp["Valor"], errors="coerce")

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df_exp.to_excel(w, index=False, sheet_name="MateriaisPorLoja")
        wb = w.book
        ws = w.sheets["MateriaisPorLoja"]

        # largura + formato BR
        fmt_money = wb.add_format({"num_format": "#.##0,00"})
        fmt_int   = wb.add_format({"num_format": "0"})
        # descobre índices das colunas
        headers = list(df_exp.columns)
        if "Valor" in headers:
            col_v = headers.index("Valor")
            ws.set_column(col_v, col_v, 14, fmt_money)
        if "Qtde" in headers:
            col_q = headers.index("Qtde")
            ws.set_column(col_q, col_q, 10, fmt_int)
        # ajusta um pouco as demais
        for i, h in enumerate(headers):
            if h not in ("Valor", "Qtde"):
                ws.set_column(i, i, 18)
    buf.seek(0)
    return buf

st.download_button(
    "⬇️ Baixar Excel",
    data=to_excel(df_final),
    file_name="materiais_por_loja.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
