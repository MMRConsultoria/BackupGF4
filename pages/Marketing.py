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
    # remove "123 - " do começo
    s = re.sub(r"^\s*\d+\s*-\s*", "", s)
    return s.strip()

def loja_join_key(txt: str) -> str:
    return normalizar_loja(txt).lower()

def pick_name(cols, targets):
    m = {_ns(c): c for c in cols}
    for t in targets:
        if _ns(t) in m:
            return m[_ns(t)]
    return None

def _parse_brl(x) -> float:
    """Converte '1.234,56', '1234,56', '0,46', '(1.234,56)' -> float (1234.56)"""
    s = str(x or "").strip()
    if s == "":
        return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s.count(",") >= 1 and s.count(".") >= 1:
        # BR: ponto milhar, vírgula decimal
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
    return s  # ex.: 1.234,56

# --------- Google Sheets: Tabela Empresa ----------
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    # aceita secrets como dict ou string JSON
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
# topo do arquivo (ou logo acima de ler_relatorio)
IGNORAR_TOTAIS = False  # <- por enquanto deixamos False (não remove Sub.Total/Total Geral)
# --------- Parser do Excel de Upload ----------
# topo do arquivo (ou logo acima de ler_relatorio)
IGNORAR_TOTAIS = False  # <- por enquanto deixamos False (não remove Sub.Total/Total Geral)
e substitua a função ler_relatorio por esta versão:

python
Copiar código
def ler_relatorio(uploaded_file) -> pd.DataFrame:
    """
    Linha 4: lojas (mescladas).
    Linha 5: cabeçalhos; pares 'Qtde' e 'Valor(R$)' por loja.
    Col B: Grupo do produto (ffill). Col C: Código (ffill). Col D: Material.
    **Sem** remover Sub.Total/Total Geral quando IGNORAR_TOTAIS=False.
    Descartar somente: qtde vazia e valor <= 0.
    """
    df0 = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
    if df0.shape[0] < 6:
        return pd.DataFrame()

    ROW_LOJA = 3   # linha 4
    ROW_HDR  = 4   # linha 5
    COL_B, COL_C, COL_D = 1, 2, 3

    # Detecta pares (Qtde, Valor)
    r5 = df0.iloc[ROW_HDR].astype(str).fillna("")
    r5n = r5.map(_ns)
    pairs = []
    j = 0
    while j < df0.shape[1] - 1:
        is_q = r5n.iloc[j] == "qtde"
        is_v = r5n.iloc[j+1] in ("valor(r$)", "valor r$", "valor (r$)", "valor(r$ )", "valor r$)")
        if is_q and is_v:
            # loja (linha 4), voltando pela mescla
            k = j
            loja = ""
            while k >= 0:
                val = str(df0.iat[ROW_LOJA, k] if k < df0.shape[1] else "")
                if val and str(val).strip().lower() not in ("nan",):
                    loja = str(val).strip()
                    break
                k -= 1
            # mesmo que a loja seja "Total", por enquanto não filtramos
            pairs.append((j, j+1, normalizar_loja(loja)))
            j += 2
        else:
            j += 1

    # Base a partir da linha 6
    base = df0.iloc[ROW_HDR+1:].copy()
    base = base.rename(columns={COL_B: "GrupoColB", COL_C: "Codigo", COL_D: "Material"})

    # Material como texto limpo; grupo/código preservam NaN para ffill correto
    base["Material"] = base["Material"].astype(str).fillna("").str.strip()

    # (opcional) marcar textos de total/subtotal — só usado se a flag estiver True
    def _is_totalish_text_local(x):
        return _is_totalish_text(x)

    if IGNORAR_TOTAIS:
        # zera textos de total/subtotal antes do ffill, se desejar ignorar
        grp_txt = base["GrupoColB"].astype(str)
        base.loc[grp_txt.map(_is_totalish_text_local), "GrupoColB"] = pd.NA

        cod_txt = base["Codigo"].astype(str)
        base.loc[cod_txt.map(_is_totalish_text_local), "Codigo"] = pd.NA

    # Grupo (ffill)
    gp = base["GrupoColB"]
    gp = gp.where(~gp.isna() & (gp.astype(str).str.strip() != ""), pd.NA).ffill()
    base["GrupoProduto"] = gp

    # Código (ffill) e padronização sem '.0'
    cd = base["Codigo"]
    cd = cd.where(~cd.isna() & (cd.astype(str).str.strip() != ""), pd.NA).ffill()

    def _to_intish_text(v):
        if pd.isna(v):
            return pd.NA
        s = str(v).strip()
        if s == "":
            return pd.NA
        try:
            return str(int(float(s)))
        except:
            return s

    base["Codigo"] = cd.map(_to_intish_text)

    if base.empty or not pairs:
        return pd.DataFrame(columns=["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"])

    registros = []
    for c_q, c_v, loja_nome in pairs:
        sub = base[["GrupoProduto","Codigo","Material", c_q, c_v]].copy()
        sub = sub.rename(columns={c_q: "Qtde", c_v: "Valor"})

        # filtros mínimos: qtde e valor
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
    st.warning("Nenhum item elegível foi encontrado (verifique Sub.Total / Total Geral / valores e qtde).")
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

# Se houver colisão Loja_x/Loja_y, escolhe o nome da Tabela Empresa
if "Loja_x" in merged.columns and "Loja_y" in merged.columns:
    merged["Loja"] = merged["Loja_y"].where(
        merged["Loja_y"].astype(str).str.strip() != "", 
        merged["Loja_x"]
    )
    merged.drop(columns=["Loja_x", "Loja_y"], inplace=True)
elif "Loja_y" in merged.columns:
    merged["Loja"] = merged["Loja_y"]
    merged.drop(columns=["Loja_y"], inplace=True)
elif "Loja_x" in merged.columns:
    merged["Loja"] = merged["Loja_x"]
    merged.drop(columns=["Loja_x"], inplace=True)

# “Operação” = Grupo (tabela empresa) ; “Grupo Material” = GrupoProduto
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

# ----- TOTAL (numérico) -----
linha_total = {
    "Operação": "TOTAL",
    "Loja": "",
    "Código Everest": "",
    "Código Grupo Everest": "",
    "Grupo Material": "",
    "Codigo Material": "",
    "Material": "",
    "Qtde": df_final["Qtde"].sum(skipna=True),
    "Valor": df_final["Valor"].sum(skipna=True),
}
df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)

# --- Visão formatada para tela/Excel ---
df_view = df_final.copy()
df_view["Valor"] = df_view["Valor"].apply(_fmt_brl)

st.subheader("Prévia")
st.dataframe(df_view.head(120), use_container_width=True, hide_index=True)
st.caption(f"Linhas (com TOTAL): {len(df_view):,}".replace(",", "."))

# Download
def to_excel(df_num: pd.DataFrame):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df_exp = df_num.copy()
        df_exp["Valor"] = df_exp["Valor"].apply(_fmt_brl)  # exporta já em 1.000,00
        df_exp.to_excel(w, index=False, sheet_name="MateriaisPorLoja")
    buf.seek(0)
    return buf

st.download_button(
    "⬇️ Baixar Excel",
    data=to_excel(df_final),  # passa df numérico; a função formata a coluna Valor
    file_name="materiais_por_loja.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
