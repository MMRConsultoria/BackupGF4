# pages/CR_CP_Importador_Everest.py
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

# 🔥 CSS igual ao seu padrão
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
    </style>
""", unsafe_allow_html=True)

# 🔒 Bloqueio de acesso (mesmo padrão)
if not st.session_state.get("acesso_liberado"):
    st.stop()

# Esconde toolbar
st.markdown("""
    <style>
        [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
        .stSpinner { visibility: visible !important; }
    </style>
""", unsafe_allow_html=True)

# ======================
# Helpers (reuso do seu estilo)
# ======================
def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII", "ignore").decode("ASCII")

def _norm(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    """
    Converte bloco colado (Excel/Sheets) em DataFrame.
    Detecta TSV na 1ª linha; senão tenta ';' e depois ','.
    """
    text = (text or "").strip("\n\r ")
    if not text:
        return pd.DataFrame()

    # TSV se houver \t na primeira linha
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

# ======================
# Conexão Google Sheets (mesmo secrets da sua página)
# ======================
@st.cache_data(show_spinner=False)
def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(credentials)

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(planilha_nome="Vendas diarias", aba_nome="Tabela Empresa") -> pd.DataFrame:
    gc = gs_client()
    planilha = gc.open(planilha_nome)
    df_emp = pd.DataFrame(planilha.worksheet(aba_nome).get_all_records())

    # normaliza cabeçalhos comuns
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

    # garante colunas essenciais
    for col in ["Código Grupo Everest", "Grupo", "Loja", "Código Everest", "Tipo"]:
        if col not in df_emp.columns:
            df_emp[col] = ""

    # limpeza básica
    for c in df_emp.columns:
        df_emp[c] = df_emp[c].astype(str).str.strip()

    # remove .0 em códigos numéricos
    for c in ["Código Grupo Everest", "Código Everest"]:
        df_emp[c] = df_emp[c].str.replace(r"\.0$", "", regex=True)

    # filtra grupos válidos
    df_emp = df_emp[df_emp["Código Grupo Everest"].astype(str).str.len() > 0].copy()
    return df_emp

# ======================
# UI — Contas a Receber
# ======================
st.markdown("""
<div style='display:flex;align-items:center;gap:10px;'>
  <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
  <h1 style='margin:0;font-size:2.0rem;'>CR-CP Importador Everest — Contas a Receber</h1>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📥 Upload / Colagem", "⚙️ Parâmetros (Grupo → Empresas)"])

with st.spinner("⏳ Carregando Tabela Empresa..."):
    df_empresa = carregar_tabela_empresa()

# --------- 📥 Upload/Colagem ----------
with tab1:
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
    if st.button("Salvar etapa (dados colados/enviados) ➜", use_container_width=True, type="primary"):
        if df_raw.empty:
            st.error("Cole ou envie o arquivo antes de salvar.")
        else:
            st.session_state["cr_df_raw"] = df_raw
            st.success("✅ Dados salvos em sessão (`cr_df_raw`). Vá para a aba **Parâmetros**.")

# --------- ⚙️ Parâmetros ----------
with tab2:
    st.caption("Selecione o **Código Grupo Everest** e depois as **empresas** desse grupo.")

    # lista de grupos
    cod_grupos = (
        df_empresa["Código Grupo Everest"]
        .astype(str).str.strip()
        .dropna().drop_duplicates().sort_values().tolist()
    )

    col_g, col_e = st.columns([0.38, 0.62])
    with col_g:
        grp = st.selectbox("Código Grupo Everest", ["— selecione —"] + cod_grupos, index=0)

    with col_e:
        empresas_sel = []
        if grp != "— selecione —":
            base = df_empresa[df_empresa["Código Grupo Everest"].astype(str).str.strip() == str(grp).strip()].copy()

            # label "Loja (Código)"
            def _label(row):
                loja = str(row.get("Loja", "")).strip()
                cod  = str(row.get("Código Everest", "")).strip()
                return f"{loja} ({cod})" if cod else loja or "—"

            base["__label__"] = base.apply(_label, axis=1)
            opts = ["Todas"] + base["__label__"].drop_duplicates().sort_values().tolist()
            empresas_sel = st.multiselect("Empresa(s) do grupo", options=opts, default=["Todas"])

    # resolve empresas escolhidas
    if grp != "— selecione —":
        if (not empresas_sel) or ("Todas" in empresas_sel):
            df_escolhidas = df_empresa[df_empresa["Código Grupo Everest"].astype(str).str.strip() == str(grp).strip()].copy()
        else:
            base = df_empresa[df_empresa["Código Grupo Everest"].astype(str).str.strip() == str(grp).strip()].copy()
            def _label(row):
                loja = str(row.get("Loja","")).strip()
                cod  = str(row.get("Código Everest","")).strip()
                return f"{loja} ({cod})" if cod else loja or "—"
            base["__label__"] = base.apply(_label, axis=1)
            df_escolhidas = base[base["__label__"].isin(set(empresas_sel))].copy()

        st.markdown("#### Seleção atual")
        cL, cR = st.columns([0.45, 0.55])
        with cL:
            st.metric("Código Grupo Everest", grp)
            st.write(f"**Empresas selecionadas:** {len(df_escolhidas)}")
        with cR:
            st.dataframe(
                df_escolhidas[["Grupo", "Loja", "Código Everest", "Tipo", "Código Grupo Everest"]]
                .sort_values(["Grupo", "Loja"]),
                use_container_width=True, height=220
            )

    st.markdown("---")
    col_a, col_b = st.columns([0.6, 0.4])
    with col_a:
        continuar = st.button("✅ Continuar com essas seleções", use_container_width=True)
    with col_b:
        limpar = st.button("↩️ Limpar tudo", use_container_width=True)

    if limpar:
        st.session_state.pop("cr_df_raw", None)
        st.session_state.pop("cr_grupo", None)
        st.session_state.pop("cr_empresas", None)
        st.experimental_rerun()

    if continuar:
        if "cr_df_raw" not in st.session_state or st.session_state["cr_df_raw"].empty:
            st.error("Primeiro salve os **dados** na aba *Upload / Colagem*.")
        elif grp == "— selecione —":
            st.error("Selecione um **Código Grupo Everest**.")
        else:
            st.session_state["cr_grupo"]    = grp
            st.session_state["cr_empresas"] = df_escolhidas.reset_index(drop=True)
            st.success("✅ Parâmetros salvos. Pode seguir para a etapa de mapeamento/geração do layout do Importador.")

# ======================
# Dica do próximo passo
# ======================
st.markdown("""
<hr/>
<b>Próximo passo</b>: mapeamento das colunas do arquivo colado/enviado para o layout do <i>Importador Everest (Contas a Receber)</i>
(Data, Complemento, Valor, Sinal etc.).  
Quando quiser, me diga os nomes mínimos de colunas que devemos exigir e eu integro aqui sem mexer no seu layout.
""", unsafe_allow_html=True)
