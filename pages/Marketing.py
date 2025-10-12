# pages/Importar_Lojas_TabelaEmpresa.py
import streamlit as st
import pandas as pd
import re, unicodedata, json
from io import BytesIO

import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Lojas ↔ Tabela Empresa", layout="wide")

# ----------------- helpers -----------------
def _strip_invis(s: str) -> str:
    return re.sub(r"[\u200B-\u200D\uFEFF]", "", str(s or ""))

def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)

def _norm_loja(s: str) -> str:
    s = str(s or "").strip()
    s = re.sub(r"^\s*\d+\s*[-–]?\s*", "", s)   # remove "123 - "
    return s.strip().lower()

def localizar_linha_qtde_valor(df: pd.DataFrame) -> int | None:
    lim = min(60, len(df))
    for r in range(lim):
        vals = [_ns(df.iat[r, c]) for c in range(df.shape[1])]
        if "qtde" in vals and any("valor" in v for v in vals):
            return r
    return None

def mapear_lojas(df: pd.DataFrame, r_sub: int):
    header = [_ns(df.iat[r_sub, c]) for c in range(df.shape[1])]

    def titulo_acima(c):
        for up in (1, 2, 3):
            r = r_sub - up
            if r < 0: break
            raw = str(df.iat[r, c]).strip()
            if raw and _ns(raw) not in ("qtde", "valor", "valor (r$)"):
                return raw
        return ""

    lojas = []
    c = 0
    while c < len(header):
        eh_qtde = header[c] == "qtde"
        eh_val  = c+1 < len(header) and ("valor" in header[c+1])
        if eh_qtde and eh_val:
            loja_raw = titulo_acima(c) or titulo_acima(c+1)
            if loja_raw:
                lojas.append({
                    "Loja (original)": loja_raw,
                    "Loja_norm": _norm_loja(loja_raw),
                })
            c += 2
        else:
            c += 1
    return lojas

# ----------------- Auth Google robusto -----------------
def _get_service_account_dict() -> dict:
    """
    Lê as credenciais do Streamlit Secrets.
    Aceita:
      - st.secrets["GOOGLE_SERVICE_ACCOUNT"] (str JSON ou dict)
      - st.secrets["gcp_service_account"]    (str JSON ou dict)
    """
    cand_keys = ["GOOGLE_SERVICE_ACCOUNT", "gcp_service_account"]
    raw = None
    for k in cand_keys:
        if k in st.secrets:
            raw = st.secrets[k]
            break
    if raw is None:
        raise RuntimeError(
            "Credenciais não encontradas em st.secrets. "
            "Defina 'GOOGLE_SERVICE_ACCOUNT' (ou 'gcp_service_account')."
        )
    if isinstance(raw, str):
        return json.loads(raw)  # era string -> vira dict
    if isinstance(raw, dict):
        return raw
    raise TypeError("Credenciais em formato inesperado (use string JSON ou dict).")

def _get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    keydict = _get_service_account_dict()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(keydict, scope)
    return gspread.authorize(creds)

# ----------------- Google Sheets: Tabela Empresa -----------------
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    gc = _get_gspread_client()
    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())

    # limpar cabeçalho e padronizar nomes
    df.columns = [_strip_invis(c).strip() for c in df.columns]
    ren = {}
    for c in df.columns:
        cn = _ns(c)
        if cn == "loja": ren[c] = "Loja"
        elif cn == "grupo": ren[c] = "Grupo"
        elif ("codigo" in cn and "everest" in cn and "grupo" not in cn): ren[c] = "Código Everest"
        elif ("codigo" in cn and "grupo"  in cn and "everest" in cn):    ren[c] = "Código Grupo Everest"
    df = df.rename(columns=ren)

    if "Loja" in df.columns:
        df["Loja_norm"] = df["Loja"].astype(str).map(_norm_loja)
    else:
        df["Loja_norm"] = ""
    return df[["Loja","Grupo","Código Everest","Código Grupo Everest","Loja_norm"]]

# ----------------- UI -----------------
st.title("🧭 Lojas do Excel + Tabela Empresa")
st.caption("Extrai as lojas do relatório e cruza com **Vendas diarias → Tabela Empresa** para gerar Loja, Grupo, Código Everest e Código Grupo Everest.")

c1, c2 = st.columns(2)
with c1:
    nome_planilha = st.text_input("Planilha no Google Sheets", value="Vendas diarias")
with c2:
    aba_empresa = st.text_input("Aba com Tabela Empresa", value="Tabela Empresa")

uploaded = st.file_uploader("Envie o Excel (venda-de-materiais-por-grupo-e-loja.xlsx)", type=["xlsx","xls"])

if uploaded:
    try:
        df_raw = pd.read_excel(uploaded, sheet_name=0, header=None, dtype=object)
        for c in range(df_raw.shape[1]):
            df_raw.iloc[:, c] = df_raw.iloc[:, c].map(_strip_invis)
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
        st.stop()

    # 1) localizar subcabeçalho Qtde/Valor e listar lojas
    r_sub = localizar_linha_qtde_valor(df_raw)
    if r_sub is None:
        st.error("Não encontrei a linha com 'Qtde' e 'Valor(R$)' para identificar as lojas.")
        st.stop()
    lojas = pd.DataFrame(mapear_lojas(df_raw, r_sub)).drop_duplicates(subset=["Loja_norm"])

    if lojas.empty:
        st.warning("Nenhuma loja identificada no arquivo.")
        st.stop()

    # 2) carregar Tabela Empresa e cruzar
    try:
        df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"Erro ao carregar a Tabela Empresa do Google Sheets: {e}")
        st.stop()

    base = lojas.merge(df_emp, on="Loja_norm", how="left")

    # 3) organizar colunas exigidas
    out = base[["Loja","Grupo","Código Everest","Código Grupo Everest"]].copy()
    out.insert(0, "Loja (arquivo)", base["Loja (original)"])

    st.subheader("Lojas encontradas + Tabela Empresa")
    st.dataframe(out, use_container_width=True, hide_index=True)
    st.info(f"Total de lojas no arquivo: {len(out)}")

    # aviso de não mapeadas
    nao_mapeadas = out[out["Código Everest"].isna() | (out["Código Everest"].astype(str).str.strip() == "")]
    if not nao_mapeadas.empty:
        st.warning(
            "Algumas lojas não possuem **Código Everest** cadastradas na Tabela Empresa: "
            + ", ".join(sorted(nao_mapeadas["Loja (arquivo)"].astype(str).unique()))
        )

    # 4) download com exatamente as 4 colunas requisitadas
    to_save = out[["Loja","Grupo","Código Everest","Código Grupo Everest"]].copy()
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
        to_save.to_excel(wr, index=False, sheet_name="Lojas")
    buf.seek(0)
    st.download_button(
        "⬇️ Baixar Excel (Loja / Grupo / Código Everest / Código Grupo Everest)",
        data=buf,
        file_name="lojas_tabela_empresa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Envie o arquivo Excel para começar.")
