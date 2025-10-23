# pages/CR_CP_Importador_Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import StringIO

# ===== Conexão Google Sheets (gspread) =====
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# ======================
# Aparência semelhante ao padrão do seu app
# ======================
st.markdown("""
<style>
/* esconde toolbar padrão */
[data-testid="stToolbar"]{visibility:hidden;height:0;position:fixed}
/* título e subtítulo */
.h-title{font-size:34px;font-weight:700;margin:6px 0 0 0}
.h-sub  {color:#6b7280;margin:0 0 14px 0}
/* “pílulas” do topo (mimetiza seus botões) */
.pillbar{display:flex;gap:12px;margin:6px 0 20px 0;flex-wrap:wrap}
.pill{background:#eef2ff;border:1px solid #e5e7eb;border-radius:12px;
      padding:8px 12px;font-weight:600;color:#3b82f6}
.pill.muted{background:#f3f4f6;color:#4b5563}
</style>
""", unsafe_allow_html=True)

# ======================
# Utils
# ======================
def _norm(s):
    if s is None: return ""
    s = str(s).replace("\u00A0"," ")
    s = re.sub(r"\s+"," ", s).strip()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    """Converte bloco colado (Excel/Sheets) em DataFrame (TSV/CSV)."""
    text = text.strip("\n\r ")
    if not text:
        return pd.DataFrame()

    # 1) TSV se houver \t na primeira linha
    if "\t" in text.splitlines()[0]:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        # 2) CSV ; (pt-BR) depois ,
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")

    df = df.dropna(how="all")
    df.columns = [_norm(c) if c else f"col_{i}" for i,c in enumerate(df.columns)]
    return df

@st.cache_data(show_spinner=False)
def get_gspread_client():
    """Cria client gspread a partir de st.secrets."""
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    # aceita chaves com nomes comuns
    secrets_candidates = ["gcp_service_account", "gsheets_service_account"]
    creds_json = None
    for key in secrets_candidates:
        if key in st.secrets:
            creds_json = dict(st.secrets[key]); break
    if creds_json is None:
        st.error("Credenciais do Google não encontradas em st.secrets.")
        st.stop()

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
    return gspread.authorize(credentials)

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    """Lê Tabela Empresa e padroniza colunas essenciais."""
    gc = get_gspread_client()
    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    data = ws.get_all_records(numeric_value_strategy="RAW")
    df = pd.DataFrame(data)

    # normaliza cabeçalhos comuns
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

    for c in df.columns:
        df[c] = df[c].apply(_norm)

    for c in ["Código Grupo Everest","Código Everest"]:
        df[c] = df[c].str.replace(r"\.0$","", regex=True)

    df = df[df["Código Grupo Everest"].astype(str).str.len()>0].copy()
    return df

def ui_upload_paste(key_prefix: str):
    """Bloco de upload + colagem, reaproveitável nas duas abas."""
    col_paste, col_file = st.columns([0.55, 0.45])
    with col_paste:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                           placeholder="Cole aqui os dados copiados do Excel/Google Sheets…",
                           key=f"{key_prefix}_paste")
        df_paste = _try_parse_paste(txt) if txt.strip() else pd.DataFrame()
    with col_file:
        up = st.file_uploader("📎 Ou enviar arquivo (.xlsx/.csv)", type=["xlsx","csv"], key=f"{key_prefix}_file")
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
                st.error(f"Erro ao ler o arquivo: {e}")

    df_raw = df_paste if not df_paste.empty else df_file
    st.markdown("##### Pré-visualização")
    if df_raw.empty:
        st.info("Nenhum dado ainda. Cole ou envie um arquivo para ver a pré-visualização.")
    else:
        st.dataframe(df_raw, use_container_width=True, height=320)
    return df_raw

def ui_grupo_empresas(df_emp: pd.DataFrame, key_prefix: str):
    """Perguntas: Grupo (Código Grupo Everest) → Empresas do grupo."""
    cod_grupos = (
        df_emp["Código Grupo Everest"].dropna().astype(str).map(_norm)
        .drop_duplicates().sort_values().tolist()
    )
    col_g, col_e = st.columns([0.38, 0.62])
    with col_g:
        grp = st.selectbox("1) **Código Grupo Everest**", ["— selecione —"] + cod_grupos, key=f"{key_prefix}_grp")
    with col_e:
        empresas_sel = []
        if grp != "— selecione —":
            df_grp = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
            def _label(row):
                loja = row.get("Loja","").strip()
                cod  = row.get("Código Everest","").strip()
                return f"{loja} ({cod})" if cod else (loja or cod or "—")
            df_grp["__label__"] = df_grp.apply(_label, axis=1)
            opts = ["Todas"] + df_grp["__label__"].drop_duplicates().sort_values().tolist()
            empresas_sel = st.multiselect("2) Empresa(s) do grupo", options=opts, default=["Todas"], key=f"{key_prefix}_emp")

    # Resolve seleção
    if grp != "— selecione —":
        if not empresas_sel or "Todas" in empresas_sel:
            df_escolhidas = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
        else:
            df_grp = df_emp[df_emp["Código Grupo Everest"].astype(str).map(_norm) == _norm(grp)].copy()
            def _label(row):
                loja = row.get("Loja","").strip()
                cod  = row.get("Código Everest","").strip()
                return f"{loja} ({cod})" if cod else (loja or cod or "—")
            df_grp["__label__"] = df_grp.apply(_label, axis=1)
            df_escolhidas = df_grp[df_grp["__label__"].isin(set(empresas_sel))].copy()
    else:
        df_escolhidas = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"])

    return grp, df_escolhidas

# ======================
# Cabeçalho no padrão do seu app
# ======================
st.markdown('<div class="h-title">Relatório Vendas Diarias</div>', unsafe_allow_html=True)
st.markdown('<div class="h-sub">Módulo: CR-CP Importador Everest</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="pillbar">'
    '<span class="pill">📥 Upload e Processamento</span>'
    '<span class="pill muted">📤 Atualizar Google Sheets</span>'
    '<span class="pill muted">📊 Auditar integração Everest</span>'
    '<span class="pill muted">📈 Auditar Faturamento X Meio Pagamento</span>'
    '</div>', unsafe_allow_html=True
)

# ======================
# Lê Tabela Empresa (uma vez)
# ======================
with st.spinner("Carregando Tabela Empresa..."):
    df_emp = carregar_tabela_empresa()  # 'Vendas diarias' / 'Tabela Empresa'

# ======================
# Duas abas: Receber / Pagar
# ======================
tab_cr, tab_cp = st.tabs(["📥 Contas a Receber", "📤 Contas a Pagar"])

# --------- 📥 CONTAS A RECEBER ---------
with tab_cr:
    st.subheader("📥 Contas a Receber — Importador Everest")
    st.caption("Cole ou envie o extrato/planilha de recebíveis; depois selecione Grupo e Empresa(s).")

    df_raw_cr = ui_upload_paste(key_prefix="cr")
    st.divider()

    grp_cr, df_emp_sel_cr = ui_grupo_empresas(df_emp, key_prefix="cr")
    st.markdown("#### Seleção atual (Receber)")
    left, right = st.columns([0.45, 0.55])
    with left:
        st.metric("Código Grupo Everest", grp_cr if grp_cr!="— selecione —" else "—")
        st.write(f"**Empresas selecionadas:** {len(df_emp_sel_cr)}")
    with right:
        if not df_emp_sel_cr.empty:
            st.dataframe(df_emp_sel_cr[["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"]]
                         .sort_values(["Grupo","Loja"]),
                         use_container_width=True, height=220)

    st.divider()
    col_a, col_b = st.columns([0.6, 0.4])
    with col_a:
        ok_cr = st.button("✅ Continuar (Receber)", use_container_width=True, key="cr_ok")
    with col_b:
        reset_cr = st.button("↩️ Limpar", use_container_width=True, key="cr_reset")
    if reset_cr: st.experimental_rerun()

    if ok_cr:
        if df_raw_cr.empty:
            st.error("Cole ou envie os dados de **Contas a Receber** antes de continuar.")
        elif grp_cr == "— selecione —":
            st.error("Selecione um **Código Grupo Everest**.")
        else:
            st.session_state["cr_df_raw"] = df_raw_cr
            st.session_state["cr_grupo"]  = grp_cr
            st.session_state["cr_empresas"] = df_emp_sel_cr.reset_index(drop=True)
            st.success("Receber: dados e seleções salvos. Próxima etapa pronta (mapeamento/layout).")

# --------- 📤 CONTAS A PAGAR ---------
with tab_cp:
    st.subheader("📤 Contas a Pagar — Importador Everest")
    st.caption("Fluxo idêntico ao de Receber (colagem/arquivo + Grupo → Empresa).")

    df_raw_cp = ui_upload_paste(key_prefix="cp")
    st.divider()

    grp_cp, df_emp_sel_cp = ui_grupo_empresas(df_emp, key_prefix="cp")
    st.markdown("#### Seleção atual (Pagar)")
    left, right = st.columns([0.45, 0.55])
    with left:
        st.metric("Código Grupo Everest", grp_cp if grp_cp!="— selecione —" else "—")
        st.write(f"**Empresas selecionadas:** {len(df_emp_sel_cp)}")
    with right:
        if not df_emp_sel_cp.empty:
            st.dataframe(df_emp_sel_cp[["Grupo","Loja","Código Everest","Tipo","Código Grupo Everest"]]
                         .sort_values(["Grupo","Loja"]),
                         use_container_width=True, height=220)

    st.divider()
    col_a, col_b = st.columns([0.6, 0.4])
    with col_a:
        ok_cp = st.button("✅ Continuar (Pagar)", use_container_width=True, key="cp_ok")
    with col_b:
        reset_cp = st.button("↩️ Limpar", use_container_width=True, key="cp_reset")
    if reset_cp: st.experimental_rerun()

    if ok_cp:
        if df_raw_cp.empty:
            st.error("Cole ou envie os dados de **Contas a Pagar** antes de continuar.")
        elif grp_cp == "— selecione —":
            st.error("Selecione um **Código Grupo Everest**.")
        else:
            st.session_state["cp_df_raw"] = df_raw_cp
            st.session_state["cp_grupo"]  = grp_cp
            st.session_state["cp_empresas"] = df_emp_sel_cp.reset_index(drop=True)
            st.success("Pagar: dados e seleções salvos. Próxima etapa pronta (mapeamento/layout).")
