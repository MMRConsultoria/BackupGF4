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
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# ====== VISUAL BÁSICO ======
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

# ======================
# Helpers
# ======================
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

# ======================
# Google Sheets (robusto)
# ======================
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
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(credentials)

def _open_planilha(planilha_nome="Vendas diarias"):
    """
    NÃO cacheia e NÃO recebe objetos não-hashable.
    Tenta abrir por título; se falhar, tenta por ID via st.secrets['VENDAS_DIARIAS_SHEET_ID'].
    """
    gc = gs_client()
    try:
        return gc.open(planilha_nome)
    except Exception as e1:
        sheet_id = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sheet_id:
            try:
                return gc.open_by_key(sheet_id)
            except Exception as e2:
                raise RuntimeError(f"Falha abrindo planilha por título e por ID. Título erro: {e1} | ID erro: {e2}")
        raise RuntimeError(f"Falha abrindo planilha por título '{planilha_nome}': {e1}")

@st.cache_data(show_spinner=False)
def carregar_tabela_empresa(planilha_nome="Vendas diarias", aba_nome="Tabela Empresa") -> pd.DataFrame:
    planilha = _open_planilha(planilha_nome)
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

    for c in ["Código Grupo Everest", "Código Everest"]:
        df_emp[c] = df_emp[c].str.replace(r"\.0$", "", regex=True)

    df_emp = df_emp[df_emp["Grupo"].astype(str).str.strip().ne("")].copy()
    return df_emp

# ======================
# UI Comum
# ======================
def ui_sel_grupo_empresa(df_emp: pd.DataFrame, key_prefix: str):
    grupos = (
        df_emp["Grupo"].astype(str).str.strip()
             .dropna().drop_duplicates().sort_values().tolist()
    )
    col_g, col_e = st.columns([0.45, 0.55])
    with col_g:
        grupo_nome = st.selectbox("Grupo (nome)", ["— selecione —"] + grupos, index=0, key=f"{key_prefix}_grp")
    with col_e:
        empresa_nome = "— selecione —"
        if grupo_nome != "— selecione —":
            mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grupo_nome)
            lojas = (
                df_emp.loc[mask_grupo, "Loja"]
                      .astype(str).str.strip().drop_duplicates().sort_values().tolist()
            )
            empresa_nome = st.selectbox("Empresa (nome)", ["— selecione —"] + lojas, index=0, key=f"{key_prefix}_emp")
    return grupo_nome, empresa_nome

def ui_paste_upload(key_prefix: str):
    c1, c2 = st.columns([0.55, 0.45])
    with c1:
        pasted = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                              placeholder="Cole aqui a grade copiada do Excel/Google Sheets…",
                              key=f"{key_prefix}_paste")
        df_paste = _try_parse_paste(pasted) if pasted.strip() else pd.DataFrame()
    with c2:
        up = st.file_uploader("📎 Ou enviar arquivo (.xlsx / .xlsm / .xls / .csv)",
                              type=["xlsx", "xlsm", "xls", "csv"], key=f"{key_prefix}_file")
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
    return df_raw

# ======================
# Cabeçalho
# ======================
st.markdown("""
<div style='display:flex;align-items:center;gap:10px;'>
  <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
  <h1 style='margin:0;font-size:2.0rem;'>CR-CP Importador Everest</h1>
</div>
""", unsafe_allow_html=True)

with st.spinner("⏳ Carregando referência de empresas..."):
    df_emp = carregar_tabela_empresa()

# ======================
# 3 ABAS
# ======================
tab_cr, tab_cp, tab_cad = st.tabs(["💰 Contas a Receber", "💸 Contas a Pagar", "🧾 Cadastro Cliente/Fornecedor"])

# --------- 💰 CONTAS A RECEBER ---------
with tab_cr:
    st.subheader("Contas a Receber")
    grp, emp = ui_sel_grupo_empresa(df_emp, key_prefix="cr")
    st.markdown("---")
    df_raw = ui_paste_upload(key_prefix="cr")

    col_a, col_b = st.columns([0.6, 0.4])
    with col_a:
        salvar = st.button("✅ Salvar seleção e dados (Receber)", use_container_width=True, type="primary", key="cr_save")
    with col_b:
        limpar = st.button("↩️ Limpar", use_container_width=True, key="cr_clear")

    if limpar:
        for k in ["cr_df_raw", "cr_grupo_nome", "cr_empresa_nome", "cr_empresa_row"]:
            st.session_state.pop(k, None)
        st.experimental_rerun()

    if salvar:
        if grp == "— selecione —":
            st.error("Selecione o **Grupo (nome)**.")
        elif emp == "— selecione —":
            st.error("Selecione o **Nome da Empresa**.")
        elif df_raw.empty:
            st.error("Cole ou envie o arquivo antes de salvar.")
        else:
            st.session_state["cr_grupo_nome"]   = grp
            st.session_state["cr_empresa_nome"] = emp
            mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grp)
            mask_loja  = df_emp["Loja"].astype(str).apply(_norm) == _norm(emp)
            st.session_state["cr_empresa_row"] = df_emp[mask_grupo & mask_loja].reset_index(drop=True)
            st.session_state["cr_df_raw"] = df_raw
            st.success("✅ Receber: seleção e dados salvos. Pronto para a próxima etapa.")

# --------- 💸 CONTAS A PAGAR ---------
with tab_cp:
    st.subheader("Contas a Pagar")
    grp, emp = ui_sel_grupo_empresa(df_emp, key_prefix="cp")
    st.markdown("---")
    df_raw = ui_paste_upload(key_prefix="cp")

    col_a, col_b = st.columns([0.6, 0.4])
    with col_a:
        salvar = st.button("✅ Salvar seleção e dados (Pagar)", use_container_width=True, type="primary", key="cp_save")
    with col_b:
        limpar = st.button("↩️ Limpar", use_container_width=True, key="cp_clear")

    if limpar:
        for k in ["cp_df_raw", "cp_grupo_nome", "cp_empresa_nome", "cp_empresa_row"]:
            st.session_state.pop(k, None)
        st.experimental_rerun()

    if salvar:
        if grp == "— selecione —":
            st.error("Selecione o **Grupo (nome)**.")
        elif emp == "— selecione —":
            st.error("Selecione o **Nome da Empresa**.")
        elif df_raw.empty:
            st.error("Cole ou envie o arquivo antes de salvar.")
        else:
            st.session_state["cp_grupo_nome"]   = grp
            st.session_state["cp_empresa_nome"] = emp
            mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grp)
            mask_loja  = df_emp["Loja"].astype(str).apply(_norm) == _norm(emp)
            st.session_state["cp_empresa_row"] = df_emp[mask_grupo & mask_loja].reset_index(drop=True)
            st.session_state["cp_df_raw"] = df_raw
            st.success("✅ Pagar: seleção e dados salvos. Pronto para a próxima etapa.")

# --------- 🧾 CADASTRO CLIENTE/FORNECEDOR ---------
with tab_cad:
    st.subheader("Cadastro de Cliente / Fornecedor")

    # Configuração (ajuste os nomes se quiser salvar em abas separadas)
    PLANILHA_DESTINO = "Vendas diarias"
    ABA_CLIENTE      = "Cadastro Clientes"
    ABA_FORNECEDOR   = "Cadastro Fornecedores"

    tipo = st.radio("Tipo de cadastro", ["Cliente", "Fornecedor"], horizontal=True)
    col1, col2 = st.columns(2)
    with col1:
        nome = st.text_input("Nome/Razão Social")
        cpf_cnpj = st.text_input("CPF/CNPJ")
        email = st.text_input("E-mail")
    with col2:
        telefone = st.text_input("Telefone")
        cidade = st.text_input("Cidade")
        uf = st.text_input("UF", max_chars=2)

    col3, col4 = st.columns(2)
    with col3:
        grupo_nome = st.selectbox("Grupo (nome)", ["— selecione —"] + sorted(df_emp["Grupo"].astype(str).unique().tolist()), index=0, key="cad_grp")
    with col4:
        empresa_nome = "— selecione —"
        if grupo_nome != "— selecione —":
            mask_grupo = df_emp["Grupo"].astype(str).apply(_norm) == _norm(grupo_nome)
            lojas = df_emp.loc[mask_grupo, "Loja"].astype(str).drop_duplicates().sort_values().tolist()
            empresa_nome = st.selectbox("Empresa (nome)", ["— selecione —"] + lojas, index=0, key="cad_emp")

    obs = st.text_area("Observações", height=100, placeholder="Opcional…")

    colA, colB = st.columns([0.6, 0.4])
    with colA:
        salvar_local = st.button("💾 Salvar somente na sessão", use_container_width=True)
    with colB:
        salvar_sheet = st.button("🗂️ Salvar no Google Sheets", use_container_width=True, type="primary")

    cadastro = {
        "Tipo": tipo,
        "Nome/Razão Social": nome.strip(),
        "CPF/CNPJ": cpf_cnpj.strip(),
        "E-mail": email.strip(),
        "Telefone": telefone.strip(),
        "Cidade": cidade.strip(),
        "UF": uf.strip().upper(),
        "Grupo": "" if grupo_nome == "— selecione —" else grupo_nome,
        "Empresa": "" if empresa_nome == "— selecione —" else empresa_nome,
        "Observações": obs.strip(),
    }

    if salvar_local:
        st.session_state.setdefault("cadastros", []).append(cadastro)
        st.success("✅ Cadastro salvo na sessão.")

    if salvar_sheet:
        # validações mínimas
        faltando = [k for k, v in cadastro.items() if k in ["Nome/Razão Social"] and v == ""]
        if faltando:
            st.error("Preencha os campos obrigatórios: Nome/Razão Social.")
        else:
            try:
                plan = _open_planilha(PLANILHA_DESTINO)
                aba_nome = ABA_CLIENTE if tipo == "Cliente" else ABA_FORNECEDOR
                try:
                    ws = plan.worksheet(aba_nome)
                except WorksheetNotFound:
                    # cria aba com cabeçalho
                    ws = plan.add_worksheet(title=aba_nome, rows=1000, cols=20)
                    ws.append_row(list(cadastro.keys()))
                # garante cabeçalho
                valores = ws.get_all_values()
                if not valores:
                    ws.append_row(list(cadastro.keys()))
                elif valores and valores[0] != list(cadastro.keys()):
                    # ajusta colunas: reordena/insere ausentes
                    headers = valores[0]
                    for h in cadastro.keys():
                        if h not in headers:
                            headers.append(h)
                    ws.delete_rows(1)
                    ws.insert_row(headers, 1)
                # reordena conforme cabeçalho atual
                headers = ws.row_values(1)
                row = [cadastro.get(h, "") for h in headers]
                ws.append_row(row)
                st.success(f"✅ Cadastro salvo na planilha: {PLANILHA_DESTINO} › {aba_nome}")
            except Exception as e:
                st.error(f"❌ Erro ao salvar no Google Sheets: {e}")

    # pré-visualização local
    if "cadastros" in st.session_state and st.session_state["cadastros"]:
        st.markdown("#### Cadastros na sessão (não enviados):")
        st.dataframe(pd.DataFrame(st.session_state["cadastros"]), use_container_width=True, height=220)
