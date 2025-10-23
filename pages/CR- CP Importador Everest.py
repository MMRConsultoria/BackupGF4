# pages/CR-CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import numpy as np
import re
import json
import unicodedata
from io import StringIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest • Contas a Receber", layout="wide")

# ====== VISUAL BÁSICO (igual ao seu padrão) ======
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
    .stSpinner { visibility: visible !important; }
    </style>
""", unsafe_allow_html=True)

# 🔒 Bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ====== HELPERS ======
def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII", "ignore").decode("ASCII")

def _norm(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    """
    Converte bloco colado (Excel/Sheets) em DataFrame (TSV > ; > ,).
    """
    text = (text or "").strip("\n\r ")
    if not text:
        return pd.DataFrame()

    if "\t" in text.splitlines()[0]:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")

    df = df.dropna(how="all")
    df.columns = [str(c).strip() if str(c).strip() != "" else f"col_{i}" for i, c in enumerate(df.columns)]
    return df

# ====== GOOGLE SHEETS ROBUSTO ======
@st.cache_data(show_spinner=False)
def gs_client():
    """
    Aceita GOOGLE_SERVICE_ACCOUNT como string JSON ou dict.
    """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        st.error("❌ st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")
        st.stop()

    # pode vir string JSON ou dict
    if isinstance(secret, str):
        credentials_dict = json.loads(secret)
    else:
        credentials_dict = dict(secret)

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(credentials)

@st.cache_data(show_spinner=False)
def _open_planilha(gc, nome_titulo="Vendas diarias"):
    """
    Tenta abrir por título; se falhar, tenta por ID via st.secrets['VENDAS_DIARIAS_SHEET_ID'].
    """
    try:
        return gc.open(nome_titulo)
    except Exception as e1:
        sheet_id = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sheet_id:
            try:
                return gc.open_by_key(sheet_id)
            except Exception as e2:
                raise RuntimeError(f"Falha abrindo planilha por título e por ID. Título erro: {e1} | ID erro: {e2}")
        raise RuntimeError(f"Falha abrindo planilha por título '{nome_titulo}': {e1}")

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(planilha_nome="Vendas diarias", aba_nome="Tabela Empresa") -> pd.DataFrame:
    gc = gs_client()
    planilha = _open_planilha(gc, planilha_nome)
    df_emp = pd.DataFrame(planilha.worksheet(aba_nome).get_all_records())

    # normaliza cabeçalhos e garante campos
    df_emp.columns = [str(c).strip() for c in df_emp.columns]
    ren = {
        "Codigo Everest": "Código Everest",
        "Codigo Grupo Everest": "Código Grupo Everest",
        "Cod Grupo Empresas": "Código Grupo Everest",
        "Loja Nome": "Loja",
        "Empresa": "Loja",
        "Grupo Nome": "Grupo",
        "Grupo_Empresa": "Grupo",
        "Tipo Loja": "Tipo",
    }
    df_emp = df_emp.rename(columns={k: v for k, v in ren.items() if k in df_emp.columns})
    for col in ["Código Grupo Everest", "Grupo", "Loja", "Código Everest", "Tipo"]:
        if col not in df_emp.columns:
            df_emp[col] = ""

    for c in df_emp.columns:
        df_emp[c] = df_emp[c].astype(str).str.strip()

    # remove .0 em códigos
    for c in ["Código Grupo Everest", "Código Everest"]:
        df_emp[c] = df_emp[c].str.replace(r"\.0$", "", regex=True)

    # apenas linhas com Grupo válido
    df_emp = df_emp[df_emp["Grupo"].astype(str).str.strip().ne("")].copy()
    return df_emp

# ====== TÍTULO ======
st.markdown("""
<div style='display:flex;align-items:center;gap:10px;'>
  <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
  <h1 style='margin:0;font-size:2.0rem;'>CR-CP Importador Everest — Contas a Receber</h1>
</div>
""", unsafe_allow_html=True)

with st.spinner("⏳ Carregando referência de empresas..."):
    df_emp = carregar_tabela_empresa()

# ==============================
# 1) SELEÇÃO: GRUPO (NOME) ➜ EMPRESA (NOME)
# ==============================
st.subheader("1) Parâmetros")

# Grupos por NOME (únicos, ordenados)
grupos = (
    df_emp["Grupo"].astype(str).str.strip()
         .dropna().drop_duplicates().sort_values().tolist()
)

col_g, col_e = st.columns([0.45, 0.55])

with col_g:
    grupo_nome = st.selectbox("Grupo (nome)", ["— selecione —"] + grupos, index=0)

with col_e:
    empresa_nome = "— selecione —"
    if grupo_nome != "— selecione —":
        # ✅ CORREÇÃO: aplicar _norm na Series com .apply
        mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grupo_nome)
        lojas = (
            df_emp.loc[mask_grupo, "Loja"]
                  .astype(str).str.strip().drop_duplicates().sort_values().tolist()
        )
        empresa_nome = st.selectbox("Empresa (nome)", ["— selecione —"] + lojas, index=0)

st.markdown("---")

# ==============================
# 2) COLAGEM / UPLOAD ABAIXO
# ==============================
st.subheader("2) Colar ou Enviar Arquivo")

c1, c2 = st.columns([0.55, 0.45])
with c1:
    pasted = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                          placeholder="Cole aqui a grade copiada do Excel/Google Sheets…")
    df_paste = _try_parse_paste(pasted) if pasted.strip() else pd.DataFrame()

with c2:
    up = st.file_uploader("📎 Ou enviar arquivo (.xlsx / .xlsm / .xls / .csv)",
                          type=["xlsx", "xlsm", "xls", "csv"])
    df_file = pd.DataFrame()
    if up is not None:
        try:
            if up.name.lower().endswith(".csv"):
                try:
                    df_file = pd.read_csv(up, sep=";", dtype=str, engine="python")
                except Exception:
                    up.seek(0)
                    df_file = pd.read_csv(up, sep=",", dtype=str, engine="python")
            else:
                df_file = pd.read_excel(up, dtype=str)
            df_file = df_file.dropna(how="all")
            df_file.columns = [str(c).strip() if str(c).strip() != "" else f"col_{i}"
                               for i, c in enumerate(df_file.columns)]
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

df_raw = df_paste if not df_paste.empty else df_file

st.markdown("#### Pré-visualização")
if df_raw.empty:
    st.info("Cole ou envie um arquivo para visualizar aqui.")
else:
    st.dataframe(df_raw, use_container_width=True, height=320)

st.markdown("---")

# ==============================
# 3) SALVAR EM SESSÃO
# ==============================
col_a, col_b = st.columns([0.6, 0.4])
with col_a:
    salvar = st.button("✅ Salvar seleção e dados", use_container_width=True, type="primary")
with col_b:
    limpar = st.button("↩️ Limpar", use_container_width=True)

if limpar:
    for k in ["cr_df_raw", "cr_grupo_nome", "cr_empresa_nome", "cr_empresa_row"]:
        st.session_state.pop(k, None)
    st.experimental_rerun()

if salvar:
    if grupo_nome == "— selecione —":
        st.error("Selecione primeiro o **Grupo (nome)**.")
    elif empresa_nome == "— selecione —":
        st.error("Selecione o **Nome da Empresa**.")
    elif df_raw.empty:
        st.error("Cole ou envie o arquivo antes de salvar.")
    else:
        # guarda seleção
        st.session_state["cr_grupo_nome"]   = grupo_nome
        st.session_state["cr_empresa_nome"] = empresa_nome

        # linha (ou linhas) da empresa selecionada — útil para obter Código Everest depois
        mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grupo_nome)
        mask_loja  = df_emp["Loja"].astype(str).apply(_norm) == _norm(empresa_nome)
        df_empresa_row = df_emp[mask_grupo & mask_loja].copy()
        st.session_state["cr_empresa_row"] = df_empresa_row.reset_index(drop=True)

        # dados colados/enviados
        st.session_state["cr_df_raw"] = df_raw

        st.success("✅ Seleção e dados salvos. Pronto para o mapeamento do layout do Importador (próxima etapa).")

# ==============================
# Dica do próximo passo
# ==============================
st.caption(
    "Próximo passo: mapear as colunas do arquivo colado (ex.: Data, Complemento, Valor, Sinal) "
    "para o layout do Importador Everest (Contas a Receber) e gerar a saída/integração."
)
