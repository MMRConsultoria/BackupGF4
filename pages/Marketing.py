import streamlit as st
import pandas as pd
import numpy as np
import re, unicodedata, json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Importar Materiais por Loja", layout="wide")
st.title("📥 Importar Materiais por Loja (com Tabela Empresa)")

# ----------------- helpers -----------------
def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def _pick(colnames, targets):
    m = {_ns(c): c for c in colnames}
    for t in targets:
        k = _ns(t)
        if k in m:
            return m[k]
    return None

def normalizar_loja_para_join(txt: str) -> str:
    s = str(txt or "").strip()
    s = re.sub(r"^\d+\s*-\s*", "", s)  # remove "123 - "
    return s.lower()

# ----------------- tabela empresa (gsheets) -----------------
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    # aceita dict OU string JSON em st.secrets
    key = "GOOGLE_SERVICE_ACCOUNT" if "GOOGLE_SERVICE_ACCOUNT" in st.secrets else "gcp_service_account"
    creds_any = st.secrets.get(key)
    if creds_any is None:
        st.error("Configure st.secrets['GOOGLE_SERVICE_ACCOUNT'] (ou 'gcp_service_account').")
        st.stop()

    if isinstance(creds_any, str):
        creds_dict = json.loads(creds_any)
    else:
        creds_dict = creds_any

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(credentials)

    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])

    cols = df.columns.tolist()
    col_loja  = _pick(cols, ["Loja"])
    col_grupo = _pick(cols, ["Grupo","Operação"])
    col_cod   = _pick(cols, ["Código Everest","Codigo Everest","Cod Everest"])
    col_codg  = _pick(cols, ["Código Grupo Everest","Codigo Grupo Everest","Cod Grupo Empresas","Código Grupo Empresas"])

    out = pd.DataFrame()
    out["Loja"] = df[col_loja].astype(str).str.strip() if col_loja else ""
    out["Loja_norm"] = out["Loja"].map(normalizar_loja_para_join)
    out["Grupo"] = df[col_grupo].astype(str).str.strip() if col_grupo else ""
    out["Código Everest"] = pd.to_numeric(df[col_cod], errors="coerce") if col_cod else pd.NA
    out["Código Grupo Everest"] = pd.to_numeric(df[col_codg], errors="coerce") if col_codg else pd.NA
    return out

# ----------------- parser do Excel de upload -----------------
def ler_relatorio(uploaded_file) -> pd.DataFrame:
    """
    Regras:
      - usar a linha 5 (índice 4) para achar 'Qtde' e 'Valor(R$)'
      - a loja está na linha 4 (índice 3), em célula mesclada; buscar à esquerda se a célula estiver vazia
      - descartar lojas cujo nome contenha 'total'
      - Grupo de produto: coluna B (índice 1) — forward-fill; descartar linhas com 'Sub.Total' na coluna C
      - Código do material: coluna C (índice 2) — forward-fill quando vazio
      - Material: coluna D (índice 3)
      - descartar registros com Valor <= 0
    """
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
    if df.shape[0] < 6:
        return pd.DataFrame()

    row_lojas = 3  # linha 4
    row_rotulos = 4  # linha 5

    # 1) detectar pares (Qtde, Valor(R$))
    r5 = df.iloc[row_rotulos].astype(str).fillna("")
    r5_norm = r5.map(_ns)

    pairs = []
    j = 0
    while j < df.shape[1]-1:
        is_q = r5_norm.iloc[j] == "qtde"
        is_v = r5_norm.iloc[j+1] in ("valor(r$)", "valor r$","valor(r$ )","valor (r$)")
        if is_q and is_v:
            # achar a loja na linha 4 (mesclada): procurar para esquerda até encontrar valor
            k = j
            nome_loja = ""
            while k >= 0:
                val = str(df.iat[row_lojas, k] if k < df.shape[1] else "")
                if val and str(val).strip().lower() not in ("nan",):
                    nome_loja = str(val).strip()
                    break
                k -= 1
            if nome_loja and "total" not in _ns(nome_loja):
                pairs.append((j, j+1, nome_loja))
            j += 2
        else:
            j += 1

    # 2) base de itens a partir da linha 6 (para pular cabeçalhos abaixo)
    base = df.iloc[row_rotulos+1:].copy()

    # colunas fixas do layout
    colB, colC, colD = 1, 2, 3
    base = base.rename(columns={colB: "GrupoColB", colC: "CodigoMaterial", colD: "Material"})

    # marcar subtotais (na COLUNA C)
    def is_subtotal_c(x):
        s = _ns(x)
        return ("sub" in s and "total" in s) or s == "subtotal" or "sub total" in s

    base["_is_subtotal"] = base[colC].apply(is_subtotal_c)

    # forward-fill do grupo (col B) e do código (col C)
    base["GrupoProduto"] = base["GrupoColB"].where(base["GrupoColB"].notna() & (base["GrupoColB"].astype(str).str.strip() != ""), np.nan).ffill()
    base["CodigoMaterial"] = base["CodigoMaterial"].where(base["CodigoMaterial"].notna() & (base["CodigoMaterial"].astype(str).str.strip() != ""), np.nan).ffill()

    # material (col D) válido
    base["Material"] = base["Material"].astype(str).str.strip()

    # excluir linhas 'Sub.Total' e linhas sem material
    base = base[(~base["_is_subtotal"]) & (base["Material"] != "")].copy()

    # 3) montar long para cada loja (Qtde/Valor)
    registros = []
    for qt_col, vl_col, loja in pairs:
        sub = base[["GrupoProduto","CodigoMaterial","Material", qt_col, vl_col]].copy()
        sub = sub.rename(columns={qt_col: "Qtde", vl_col: "Valor"})

        # numéricos
        sub["Qtde"] = pd.to_numeric(sub["Qtde"], errors="coerce").fillna(0.0)

        val = sub["Valor"].astype(str)
        val = val.str.replace("R$", "", regex=False)
        val = val.str.replace("\u00A0", "", regex=False)  # NBSP
        val = val.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        sub["Valor"] = pd.to_numeric(val, errors="coerce").fillna(0.0)

        sub["LojaArquivo"] = loja
        # descartar sem valor nessa loja
        sub = sub[sub["Valor"] > 0]

        if not sub.empty:
            registros.append(sub)

    if not registros:
        return pd.DataFrame(columns=["GrupoProduto","CodigoMaterial","Material","LojaArquivo","Qtde","Valor"])

    df_long = pd.concat(registros, ignore_index=True)
    df_long["CodigoMaterial"] = df_long["CodigoMaterial"].astype(str).str.strip()
    df_long["GrupoProduto"] = df_long["GrupoProduto"].astype(str).str.strip()

    return df_long[["GrupoProduto","CodigoMaterial","Material","LojaArquivo","Qtde","Valor"]]

# ----------------- UI -----------------
c1, c2 = st.columns(2)
with c1:
    nome_planilha = st.text_input("Nome da planilha no Google Sheets", value="Vendas diarias")
with c2:
    aba_empresa = st.text_input("Aba da Tabela Empresa", value="Tabela Empresa")

up = st.file_uploader("Envie o Excel (linhas 4/5 = Loja / Qtde-Valor)", type=["xlsx","xls"])

if up is None:
    st.info("Envie o arquivo para começar.")
    st.stop()

# parse do excel
try:
    df_items = ler_relatorio(up)
except Exception as e:
    st.error(f"Erro ao ler o arquivo: {e}")
    st.stop()

st.subheader("Prévia do que foi lido do Excel")
st.dataframe(df_items.head(50), use_container_width=True, hide_index=True)
st.caption(f"Linhas elegíveis (Valor>0): {len(df_items):,}".replace(",", "."))

# tabela empresa
try:
    df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
except Exception as e:
    st.error(f"❌ Erro ao carregar Tabela Empresa: {e}")
    st.stop()

# join por loja normalizada
df_items["Loja_norm"] = df_items["LojaArquivo"].map(normalizar_loja_para_join)
merged = df_items.merge(
    df_emp[["Loja_norm","Loja","Grupo","Código Everest","Código Grupo Everest"]],
    on="Loja_norm", how="left"
)

# renomear Grupo -> Operação e Valor
merged = merged.rename(columns={"Grupo": "Operação", "Valor": "Valor (R$)"})

# colunas finais
final_cols = [
    "Operação","Loja","Código Everest","Código Grupo Everest",
    "GrupoProduto","CodigoMaterial","Material","Qtde","Valor (R$)"
]
for c in final_cols:
    if c not in merged.columns:
        merged[c] = ""

df_final = merged[final_cols].copy()

st.subheader("Resultado com Operação/Loja/Códigos (Tabela Empresa)")
st.dataframe(df_final.head(100), use_container_width=True, hide_index=True)
st.info(f"Total de linhas após join: {len(df_final):,}".replace(",", "."))

# download
def to_excel(df: pd.DataFrame):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="MateriaisPorLoja")
    buf.seek(0)
    return buf

st.download_button(
    "⬇️ Baixar Excel",
    data=to_excel(df_final),
    file_name="materiais_por_loja_com_empresa.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
