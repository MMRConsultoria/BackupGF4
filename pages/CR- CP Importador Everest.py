# pages/CR_CP_Importador_Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import StringIO

import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# ========= utils =========
def _norm(s):
    if s is None: return ""
    s = str(s).replace("\u00A0"," ")
    s = re.sub(r"\s+"," ", s).strip()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    text = (text or "").strip("\n\r ")
    if not text:
        return pd.DataFrame()
    # tenta TSV; senão ; e depois ,
    if "\t" in text.splitlines()[0]:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")
    df = df.dropna(how="all")
    df.columns = [_norm(c) if c else f"col_{i}" for i,c in enumerate(df.columns)]
    return df

@st.cache_data(show_spinner=False)
def _gspread():
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    # aceite de nomes comuns no st.secrets
    for key in ("gcp_service_account","gsheets_service_account"):
        if key in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets[key]), scope)
            return gspread.authorize(creds)
    st.error("Credenciais do Google não encontradas em st.secrets."); st.stop()

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    sh = _gspread().open(planilha)
    ws = sh.worksheet(aba)
    df = pd.DataFrame(ws.get_all_records(numeric_value_strategy="RAW"))
    df.columns = [_norm(c) for c in df.columns]
    ren = {
        "codigo everest": "Código Everest",
        "codigo grupo everest": "Código Grupo Everest",
        "cod grupo empresas": "Código Grupo Everest",
        "loja nome": "Loja",
        "empresa": "Loja",
        "grupo nome": "Grupo",
        "grupo_empresa": "Grupo",
        "tipo loja": "Tipo",
    }
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for col in ["Código Grupo Everest","Grupo","Loja","Código Everest","Tipo"]:
        if col not in df.columns: df[col] = np.nan
    for c in df.columns: df[c] = df[c].apply(_norm)
    for c in ["Código Grupo Everest","Código Everest"]:
        df[c] = df[c].str.replace(r"\.0$","", regex=True)
    df = df[df["Código Grupo Everest"].astype(str).str.len()>0].copy()
    return df

def bloco_upload_ou_colagem(key_prefix: str):
    c1,c2 = st.columns([0.55,0.45])
    with c1:
        txt = st.text_area("Colar (Ctrl+V)", height=200, key=f"{key_prefix}_paste",
                           placeholder="Cole aqui os dados copiados do Excel/Sheets…")
        df_paste = _try_parse_paste(txt) if txt.strip() else pd.DataFrame()
    with c2:
        up = st.file_uploader("Ou enviar arquivo (.xlsx/.csv)", type=["xlsx","csv"], key=f"{key_prefix}_file")
        df_file = pd.DataFrame()
        if up is not None:
            try:
                if up.name.lower().endswith(".xlsx"):
                    df_file = pd.read_excel(up, dtype=str)
                else:
                    try:
                        df_file = pd.read_csv(up, sep=";", dtype=str, engine="python")
                    except Exception:
                        up.seek(0)
                        df_file = pd.read_csv(up, sep=",", dtype=str, engine="python")
                df_file.columns = [_norm(c) if c else f"col_{i}" for i,c in enumerate(df_file.columns)]
                df_file = df_file.dropna(how="all")
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")
    df_raw = df_paste if not df_paste.empty else df_file
    if not df_raw.empty:
        st.dataframe(df_raw, use_container_width=True, height=300)
    else:
        st.info("Cole ou envie um arquivo para pré-visualizar.")
    return df_raw

def bloco_grupo_empresas(df_emp: pd.DataFrame, key_prefix: str):
    grupos = (df_emp["Código Grupo Everest"].dropna().astype(str).map(_norm)
              .drop_duplicates().sort_values().tolist())
    col_g, col_e = st.columns([0.38,0.62])
    with col_g:
        grp = st.selectbox("Código Grupo Everest", ["— selecione —"]+grupos, key=f"{key_prefix}_grp")
    with col_e:
        empresas_sel = []
        if grp != "— selecione —":
            df_grp = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
            def _label(r):
                loja = r.get("Loja","").strip(); cod = r.get("Código Everest","").strip()
                return f"{loja} ({cod})" if cod else (loja or cod or "—")
            df_grp["__label__"] = df_grp.apply(_label, axis=1)
            opts = ["Todas"] + df_grp["__label__"].drop_duplicates().sort_values().tolist()
            empresas_sel = st.multiselect("Empresa(s) do grupo", options=opts, default=["Todas"], key=f"{key_prefix}_emp")

    # resolve seleção
    if grp == "— selecione —":
        df_escolhidas = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"])
    else:
        base = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
        if (not empresas_sel) or ("Todas" in empresas_sel):
            df_escolhidas = base
        else:
            def _label(r):
                loja = r.get("Loja","").strip(); cod = r.get("Código Everest","").strip()
                return f"{loja} ({cod})" if cod else (loja or cod or "—")
            base["__label__"] = base.apply(_label, axis=1)
            df_escolhidas = base[base["__label__"].isin(set(empresas_sel))].copy()
    return grp, df_escolhidas

# ========= UI (Receber apenas, sem layout adicional) =========
st.header("CR-CP • Contas a Receber (Importador Everest)")
st.caption("Cole/enviar o arquivo; selecione o Código Grupo Everest e as empresas. Sem imposição de layout — apenas captura dos dados.")

with st.spinner("Carregando Tabela Empresa…"):
    df_emp = carregar_tabela_empresa()

df_raw = bloco_upload_ou_colagem("cr")
st.markdown("---")
grp, df_emp_sel = bloco_grupo_empresas(df_emp, "cr")

# resumo
if grp != "— selecione —":
    st.write(f"**Grupo:** {grp}  |  **Empresas selecionadas:** {len(df_emp_sel)}")
    if not df_emp_sel.empty:
        st.dataframe(df_emp_sel[["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"]]
                     .sort_values(["Grupo","Loja"]),
                     use_container_width=True, height=220)

st.markdown("---")
col_a, col_b = st.columns([0.6,0.4])
with col_a:
    ok = st.button("Continuar (Receber)", use_container_width=True)
with col_b:
    reset = st.button("Limpar", use_container_width=True)
if reset: st.experimental_rerun()

if ok:
    if df_raw.empty:
        st.error("Cole ou envie os dados antes de continuar.")
    elif grp == "— selecione —":
        st.error("Selecione o **Código Grupo Everest**.")
    else:
        st.session_state["cr_df_raw"] = df_raw
        st.session_state["cr_grupo"]  = grp
        st.session_state["cr_empresas"] = df_emp_sel.reset_index(drop=True)
        st.success("Dados e seleções salvos. Pronto para a etapa de mapeamento/integração.")
