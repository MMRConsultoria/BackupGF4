# pages/Importar_Lojas_TabelaEmpresa.py
import streamlit as st
import pandas as pd
import re, unicodedata, json
from io import BytesIO

import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Lojas ↔ Tabela Empresa", layout="wide")

# ---------- helpers ----------
def _strip_invis(s: str) -> str:
    return re.sub(r"[\u200B-\u200D\uFEFF]", "", str(s or ""))

def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)

def _norm_loja(s: str) -> str:
    s = str(s or "").strip()
    s = re.sub(r"^\s*\d+\s*[-–]?\s*", "", s)  # remove prefixo "123 - "
    return s.strip().lower()

def localizar_linha_cabecalho_qtde_valor(df: pd.DataFrame) -> int | None:
    """linha que contém Qtde e Valor (R$)"""
    lim = min(60, len(df))
    for r in range(lim):
        vals = [_ns(df.iat[r, c]) for c in range(df.shape[1])]
        if "qtde" in vals and any("valor" in v for v in vals):
            return r
    return None

def _titulo_acima(df, r_sub, c):
    """texto de título logo acima dos pares (Qtde/Valor)"""
    for up in (1, 2, 3):
        r = r_sub - up
        if r < 0: break
        raw = str(df.iat[r, c]).strip()
        if raw and _ns(raw) not in ("qtde", "valor", "valor (r$)"):
            return raw
    return ""

def mapear_lojas(df: pd.DataFrame, r_sub: int):
    """varre pares (Qtde, Valor) e pega o título acima como nome da loja;
       ignora cabeçalhos 'Total/Subtotal' que ficam no início"""
    header = [_ns(df.iat[r_sub, c]) for c in range(df.shape[1])]
    lojas, c = [], 0
    while c < len(header):
        eh_qtde = header[c] == "qtde"
        eh_val  = c+1 < len(header) and ("valor" in header[c+1])
        if eh_qtde and eh_val:
            nome = _titulo_acima(df, r_sub, c) or _titulo_acima(df, r_sub, c+1)
            nome_norm = _ns(nome)
            if nome and not re.search(r"\b(total|subtotal)\b", nome_norm):
                lojas.append(_norm_loja(nome))
            c += 2
        else:
            c += 1
    # dedup mantendo a ordem
    seen, out = set(), []
    for l in lojas:
        if l and l not in seen:
            seen.add(l); out.append(l)
    return out

# ---------- Google Auth ----------
def _get_service_account_dict() -> dict:
    for k in ("GOOGLE_SERVICE_ACCOUNT", "gcp_service_account"):
        if k in st.secrets:
            raw = st.secrets[k]
            return json.loads(raw) if isinstance(raw, str) else raw
    raise RuntimeError("Defina a credencial em st.secrets como GOOGLE_SERVICE_ACCOUNT (ou gcp_service_account).")

def _get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(_get_service_account_dict(), scope)
    return gspread.authorize(creds)

def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    gc = _get_gspread_client()
    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    df.columns = [_strip_invis(c).strip() for c in df.columns]

    ren = {}
    for c in df.columns:
        cn = _ns(c)
        if cn == "loja": ren[c] = "Loja"
        elif cn == "grupo": ren[c] = "Grupo"
        elif ("codigo" in cn and "everest" in cn and "grupo" not in cn): ren[c] = "Código Everest"
        elif ("codigo" in cn and "grupo" in cn and "everest" in cn):     ren[c] = "Código Grupo Everest"
    df = df.rename(columns=ren)

    if "Loja" in df.columns:
        df["Loja_norm"] = df["Loja"].astype(str).map(_norm_loja)
    else:
        df["Loja_norm"] = ""

    cols = ["Loja","Grupo","Código Everest","Código Grupo Everest","Loja_norm"]
    return df[[c for c in cols if c in df.columns]]

# ---------- UI ----------
st.title("🧭 Lojas ↔ Tabela Empresa")
st.caption("Extrai as lojas (ignorando a linha Total/Subtotal do início) e cruza com **Vendas diarias → Tabela Empresa**.")

c1, c2 = st.columns(2)
with c1:
    nome_planilha = st.text_input("Planilha Google Sheets", value="Vendas diarias")
with c2:
    aba_empresa = st.text_input("Aba da Tabela Empresa", value="Tabela Empresa")

uploaded = st.file_uploader("Envie o Excel (venda-de-materiais-por-grupo-e-loja.xlsx)", type=["xlsx","xls"])

if uploaded:
    try:
        df_raw = pd.read_excel(uploaded, sheet_name=0, header=None, dtype=object)
        for c in range(df_raw.shape[1]):
            df_raw.iloc[:, c] = df_raw.iloc[:, c].map(_strip_invis)
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
        st.stop()

    r_sub = localizar_linha_cabecalho_qtde_valor(df_raw)
    if r_sub is None:
        st.error("Não encontrei a linha com 'Qtde' e 'Valor(R$)'.")
        st.stop()

    lojas_norm = mapear_lojas(df_raw, r_sub)
    if not lojas_norm:
        st.warning("Nenhuma loja identificada no arquivo.")
        st.stop()

    try:
        df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"Erro ao carregar Tabela Empresa do Google Sheets: {e}")
        st.stop()

    base = pd.DataFrame({"Loja_norm": lojas_norm})
    out = base.merge(df_emp, on="Loja_norm", how="left")

    # apenas o que você pediu (sem 'Loja (arquivo)')
    resultado = out[["Loja","Grupo","Código Everest","Código Grupo Everest"]].copy()

    st.subheader("Lojas encontradas")
    st.dataframe(resultado, use_container_width=True, hide_index=True)
    st.info(f"Total de lojas: {len(resultado)}")

    faltando = resultado[resultado["Código Everest"].isna() | (resultado["Código Everest"].astype(str).str.strip() == "")]
    if not faltando.empty:
        st.warning(
            "Lojas sem **Código Everest** na Tabela Empresa: " +
            ", ".join(sorted(faltando["Loja"].dropna().astype(str).unique()))
        )

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
        resultado.to_excel(wr, index=False, sheet_name="Lojas")
    buf.seek(0)
    st.download_button(
        "⬇️ Baixar Excel (Loja / Grupo / Código Everest / Código Grupo Everest)",
        data=buf,
        file_name="lojas_tabela_empresa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Envie o arquivo Excel para começar.")
