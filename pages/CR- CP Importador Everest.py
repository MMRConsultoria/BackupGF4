# pages/CR_CP_Importador_Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import numpy as np
from io import StringIO, BytesIO
import re

# ===== Conexão Google Sheets (gspread) =====
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# ======================
# Aparência opcional
# ======================
st.markdown("""
<style>
[data-testid="stToolbar"]{visibility:hidden;height:0;position:fixed}
</style>
""", unsafe_allow_html=True)

# ======================
# Utilidades
# ======================
def _norm(s):
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u00A0", " ")  # NBSP
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    """
    Converte um bloco colado (Excel/Sheets) em DataFrame.
    Tenta por:
    1) TSV (tabs)
    2) CSV ; (pt-BR)
    3) CSV ,
    Remove linhas 100% vazias.
    """
    text = text.strip("\n\r ")
    if not text:
        return pd.DataFrame()

    # 1) tenta TSV (se tiver \t na primeira linha, é muito provável)
    if "\t" in text.splitlines()[0]:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        # 2) tenta ; (pt-BR)
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            # 3) tenta ,
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")

    if df.empty:
        return df

    # limpa linhas totalmente vazias
    df = df.dropna(how="all")
    df.columns = [ _norm(c) if c else f"col_{i}" for i,c in enumerate(df.columns) ]
    return df

@st.cache_data(show_spinner=False)
def get_gspread_client():
    """
    Tenta criar o client gspread usando st.secrets.
    Aceita chaves comuns: 'gcp_service_account' OU 'gsheets_service_account'.
    """
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    secrets_candidates = ["gcp_service_account", "gsheets_service_account"]
    creds_json = None
    for key in secrets_candidates:
        if key in st.secrets:
            creds_json = dict(st.secrets[key])
            break
    if creds_json is None:
        st.stop()  # Sem credencial, paramos para evitar erros silenciosos

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
    return gspread.authorize(credentials)

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    gc = get_gspread_client()
    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    data = ws.get_all_records(numeric_value_strategy="RAW")
    df = pd.DataFrame(data)

    # Normalizações de cabeçalho frequentes
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
    df.columns = [ _norm(c) for c in df.columns ]
    df = df.rename(columns={_norm(k): v for k,v in ren.items() if _norm(k) in df.columns})

    # Garante colunas essenciais
    for col in ["Código Grupo Everest", "Grupo", "Loja", "Código Everest", "Tipo"]:
        if col not in df.columns:
            df[col] = np.nan

    # Limpezas
    for c in df.columns:
        df[c] = df[c].apply(_norm)

    # padroniza códigos como string (sem .0)
    for c in ["Código Grupo Everest", "Código Everest"]:
        df[c] = df[c].str.replace(r"\.0$", "", regex=True)

    # remove linhas vazias de código de grupo
    df = df[ df["Código Grupo Everest"].astype(str).str.len() > 0 ].copy()

    return df

# ======================
# UI
# ======================
st.title("CR-CP Importador Everest")

with st.expander("Instruções rápidas", expanded=False):
    st.markdown("""
- **Cole** diretamente do Excel/Sheets no campo abaixo **ou** envie um arquivo.
- Depois escolha o **Grupo (Código Grupo Everest)** e a(s) **empresa(s)** desse grupo.
- Esta tela é flexível: não impomos layout fixo agora — vamos apenas **capturar** os dados que você colar/enviar.
""")

# ---- Entrada de dados (colar OU arquivo)
col_paste, col_file = st.columns([0.55, 0.45])

with col_paste:
    txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                       placeholder="Cole aqui os dados copiados do Excel/Google Sheets…")
    df_paste = _try_parse_paste(txt) if txt.strip() else pd.DataFrame()

with col_file:
    up = st.file_uploader("📎 Ou enviar arquivo (.xlsx ou .csv)", type=["xlsx","csv"])
    df_file = pd.DataFrame()
    if up is not None:
        try:
            if up.name.lower().endswith(".xlsx"):
                df_file = pd.read_excel(up, dtype=str)
            else:
                # tenta ; depois ,
                try:
                    df_file = pd.read_csv(up, sep=";", dtype=str, engine="python")
                except Exception:
                    up.seek(0)
                    df_file = pd.read_csv(up, sep=",", dtype=str, engine="python")
            df_file.columns = [ _norm(c) if c else f"col_{i}" for i,c in enumerate(df_file.columns) ]
            df_file = df_file.dropna(how="all")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")

# Escolhe a melhor fonte (colar prioriza, senão arquivo)
df_raw = df_paste if not df_paste.empty else df_file

st.markdown("### Pré-visualização dos dados colados/enviados")
if df_raw.empty:
    st.info("Nenhum dado ainda. Cole ou envie um arquivo para ver a pré-visualização.")
else:
    st.dataframe(df_raw, use_container_width=True, height=340)

st.divider()

# ======================
# Perguntas: Grupo e Empresas
# ======================
with st.spinner("Carregando Tabela Empresa..."):
    df_emp = carregar_tabela_empresa()  # 'Vendas diarias' / 'Tabela Empresa'

# Lista de grupos (Código Grupo Everest)
cod_grupos = (
    df_emp["Código Grupo Everest"]
    .dropna()
    .astype(str)
    .map(_norm)
    .drop_duplicates()
    .sort_values()
    .tolist()
)

st.subheader("Parametros de importação")
col_g, col_e = st.columns([0.38, 0.62])

with col_g:
    grp = st.selectbox(
        "1) Selecione o **Grupo (Código Grupo Everest)**",
        options=["— selecione —"] + cod_grupos,
        index=0
    )

with col_e:
    empresas_sel = []
    if grp != "— selecione —":
        df_grp = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()

        # monta label bonitinha: "Loja (Código Everest)"
        def _label(row):
            loja = row.get("Loja","").strip()
            cod  = row.get("Código Everest","").strip()
            if cod:
                return f"{loja} ({cod})"
            return loja or cod or "—"
        df_grp["__label__"] = df_grp.apply(_label, axis=1)

        opts = ["Todas"] + df_grp["__label__"].drop_duplicates().sort_values().tolist()
        empresas_sel = st.multiselect("2) Empresa(s) do grupo", options=opts, default=["Todas"])

# Resumo da escolha
if grp != "— selecione —":
    if "Todas" in empresas_sel or not empresas_sel:
        df_escolhidas = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
    else:
        labels = set(empresas_sel)
        df_tmp = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
        def _label(row):
            loja = row.get("Loja","").strip()
            cod  = row.get("Código Everest","").strip()
            return f"{loja} ({cod})" if cod else (loja or cod or "—")
        df_tmp["__label__"] = df_tmp.apply(_label, axis=1)
        df_escolhidas = df_tmp[df_tmp["__label__"].isin(labels)].copy()

    st.markdown("#### Seleção atual")
    c1, c2 = st.columns([0.45, 0.55])
    with c1:
        st.metric("Código Grupo Everest", grp)
        st.write(f"**Empresas selecionadas:** {len(df_escolhidas)}")
    with c2:
        st.dataframe(
            df_escolhidas[["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"]]
            .sort_values(["Grupo","Loja"]),
            use_container_width=True, height=240
        )

# ======================
# Próximo passo (placeholder)
# ======================
st.divider()
col_left, col_right = st.columns([0.6, 0.4])
with col_left:
    avancar = st.button("✅ Continuar com essas seleções", use_container_width=True)
with col_right:
    cancelar = st.button("↩️ Limpar/começar de novo", use_container_width=True)

if cancelar:
    st.experimental_rerun()

if avancar:
    if df_raw.empty:
        st.error("Cole ou envie os dados antes de continuar.")
    elif grp == "— selecione —":
        st.error("Selecione um **Código Grupo Everest**.")
    else:
        # Aqui você pode: validar colunas, mapear para o layout do Importador,
        # salvar em sessão, ou abrir próxima etapa (ex.: CR/CP, mapeios, etc.)
        st.success("Ok! Dados carregados e parâmetros definidos. Pronto para a próxima etapa do Importador. 👍")
        st.session_state["crcp_df_raw"] = df_raw
        st.session_state["crcp_grupo"]  = grp
        st.session_state["crcp_empresas"] = (
            df_escolhidas[["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"]].reset_index(drop=True)
        )
        st.caption("As seleções foram guardadas em `st.session_state` para uso na etapa seguinte.")

