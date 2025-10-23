# pages/CR-CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re
import json
import unicodedata
from io import StringIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# 🔒 Bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS (layout igual ao seu PainelResultados.py) =====
st.markdown("""
    <style>
        [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
        .stSpinner { visibility: visible !important; }
        .stApp { background-color: #f9f9f9; }
        div[data-baseweb="tab-list"] { margin-top: 20px; }
        button[data-baseweb="tab"] {
            background-color: #f0f2f6; border-radius: 10px;
            padding: 10px 20px; margin-right: 10px;
            transition: all 0.3s ease; font-size: 16px; font-weight: 600;
        }
        button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
        button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }

        /* multiselect sem tags coloridas (mantido do seu modelo) */
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] { background-color: transparent !important; border: none !important; color: black !important; }
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] * { color: black !important; fill: black !important; }
        div[data-testid="stMultiSelect"] > div { background-color: transparent !important; }
    </style>
""", unsafe_allow_html=True)
st.markdown("""
    <style>
    /* separador mais fino e com pouco espaço */
    hr.compact { height:1px; background:#e6e9f0; border:none; margin:8px 0 10px; }
    
    /* encurta o espaço vertical entre controles dentro da área 'compact' */
    .compact [data-testid="stSelectbox"] { margin-bottom:6px !important; }
    .compact [data-testid="stFileUploader"] { margin-top:8px !important; }
    .compact [data-testid="stTextArea"] { margin-top:8px !important; }
    
    /* reduz espaço padrão entre blocos verticais nessa seção */
    .compact [data-testid="stVerticalBlock"] > div { margin-bottom:8px; }
    </style>
""", unsafe_allow_html=True)
# ===== Cabeçalho =====
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 12px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.0rem;'>CR-CP Importador Everest</h1>
    </div>
""", unsafe_allow_html=True)

# ======================
# Helpers (mantidos)
# ======================
def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII","ignore").decode("ASCII")

def _norm(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+"," ", s).strip().lower()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    text = (text or "").strip("\n\r ")
    if not text: return pd.DataFrame()
    if "\t" in text.splitlines()[0]:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")
    df = df.dropna(how="all")
    df.columns = [str(c).strip() if str(c).strip() else f"col_{i}" for i,c in enumerate(df.columns)]
    return df

# ======================
# Google Sheets (NÃO cachear cliente; funções resilientes)
# ======================
def gs_client():
    """Cria o client do Google (não cachear!)."""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    """
    Tenta abrir por título e, se falhar, tenta por ID.
    Nunca levanta exceção: retorna None e mostra aviso.
    """
    try:
        gc = gs_client()
    except Exception as e:
        st.warning(f"⚠️ Falha ao criar o cliente do Google. Verifique as credenciais. Detalhes: {e}")
        return None

    # 1) Título
    try:
        return gc.open(title)
    except Exception as e_title:
        # 2) ID (se existir em secrets)
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sid:
            try:
                return gc.open_by_key(sid)
            except Exception as e_id:
                st.warning(
                    "⚠️ Não consegui abrir a planilha ‘Vendas diarias’. "
                    "Confira o compartilhamento com o service account e a variável VENDAS_DIARIAS_SHEET_ID.\n\n"
                    f"Erro por título: {e_title}\nErro por ID: {e_id}"
                )
                return None
        st.warning(f"⚠️ Não consegui abrir a planilha por título. Detalhes: {e_title}")
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    """
    Retorna (df_empresas, grupos, lojas_map) — só objetos pickláveis.
    Se não conseguir conectar/lêr, devolve estruturas vazias e mantém o layout.
    """
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        df_vazio = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Código Grupo Everest"])
        return df_vazio, [], {}

    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning(f"⚠️ Não consegui ler a aba 'Tabela Empresa'. Detalhes: {e}")
        df = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Código Grupo Everest"])

    # normalizações
    ren = {
        "Codigo Everest":"Código Everest","Codigo Grupo Everest":"Código Grupo Everest",
        "Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo",
    }
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype(str).str.strip()
    df = df[df["Grupo"]!=""].copy()

    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    lojas_map = (
        df.groupby("Grupo")["Loja"]
          .apply(lambda s: sorted(pd.Series(s.dropna().unique()).astype(str).tolist()))
          .to_dict()
    )
    return df, grupos, lojas_map

df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
def LOJAS_DO(grupo_nome: str): return LOJAS_MAP.get(grupo_nome, [])

# ======================
# Componentes de UI (layout)
# ======================
def filtros_grupo_empresa(prefix: str):
    col1, col2,  = st.columns([1, 1 ])
    with col1:
        gsel = st.selectbox("Grupo:", ["— selecione —"]+GRUPOS, key=f"{prefix}_grupo")
    with col2:
        lojas = LOJAS_DO(gsel) if gsel!="— selecione —" else []
        esel = st.selectbox("Empresa:", ["— selecione —"]+lojas, key=f"{prefix}_empresa")
    
    return gsel, esel

def bloco_colagem(prefix: str):
    c1,c2 = st.columns([0.55,0.45])
    with c1:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                           placeholder="Cole aqui os dados copiados do Excel/Sheets…",
                           key=f"{prefix}_paste")
        df_paste = _try_parse_paste(txt) if (txt and txt.strip()) else pd.DataFrame()
    with c2:
        up = st.file_uploader("Ou enviar arquivo (.xlsx/.xlsm/.xls/.csv)", 
                              type=["xlsx","xlsm","xls","csv"], key=f"{prefix}_file")
        df_file = pd.DataFrame()
        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    try:
                        df_file = pd.read_csv(up, sep=";", dtype=str, engine="python")
                    except Exception:
                        up.seek(0); df_file = pd.read_csv(up, sep=",", dtype=str, engine="python")
                else:
                    df_file = pd.read_excel(up, dtype=str)
                df_file = df_file.dropna(how="all")
                df_file.columns = [str(c).strip() if str(c).strip() else f"col_{i}" for i,c in enumerate(df_file.columns)]
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")
    df_raw = df_paste if not df_paste.empty else df_file

    st.markdown("#### Pré-visualização")
    if df_raw.empty: st.info("Cole ou envie um arquivo para visualizar.")
    else: st.dataframe(df_raw, use_container_width=True, height=320)
    return df_raw

# ======================
# ABAS (layout apenas; nada específico de CR/CP alterado)
# ======================
aba_cr, aba_cp, aba_cad = st.tabs(["Contas a Receber", " Contas a Pagar", "Cadastro Cliente/Fornecedor"])

# --------- 💰 CONTAS A RECEBER ---------
with aba_cr:
  

    # ↓↓↓ abre uma seção "compact" para reduzir os espaços verticais
    st.markdown('<div class="compact">', unsafe_allow_html=True)

    gsel, esel = filtros_grupo_empresa("cr")

    # em vez de st.divider():
    st.markdown('<hr class="compact">', unsafe_allow_html=True)

    df_raw = bloco_colagem("cr")

    # fecha a seção "compact"
    st.markdown('</div>', unsafe_allow_html=True)

    colA, colB = st.columns([0.6, 0.4])
    with colA:
        salvar = st.button("✅ Salvar seleção e dados (Receber)", use_container_width=True, type="primary", key="cr_save_btn")
    with colB:
        limpar = st.button("↩️ Limpar", use_container_width=True, key="cr_clear_btn")


    if limpar:
        for k in ["cr_df_raw", "cr_grupo_nome", "cr_empresa_nome", "cr_empresa_row"]:
            st.session_state.pop(k, None)
        st.experimental_rerun()

    if salvar:
        if gsel=="— selecione —": st.error("Selecione o **Grupo**.")
        elif esel=="— selecione —": st.error("Selecione a **Empresa**.")
        elif df_raw.empty: st.error("Cole ou envie o arquivo.")
        else:
            st.session_state["cr_grupo_nome"]=gsel
            st.session_state["cr_empresa_nome"]=esel
            # linha da empresa (se precisar depois)
            mask_g = df_emp["Grupo"].astype(str).apply(_norm)==_norm(gsel)
            mask_e = df_emp["Loja"].astype(str).apply(_norm)==_norm(esel)
            st.session_state["cr_empresa_row"]=df_emp[mask_g & mask_e].reset_index(drop=True)
            st.session_state["cr_df_raw"]=df_raw
            st.success("Receber salvo em sessão.")

# --------- 💸 CONTAS A PAGAR ---------
with aba_cp:
    #st.subheader("Contas a Pagar")
    gsel, esel = filtros_grupo_empresa("cp")
    st.divider()
    df_raw = bloco_colagem("cp")

    colA, colB = st.columns([0.6, 0.4])
    with colA:
        salvar = st.button("✅ Salvar seleção e dados (Pagar)", use_container_width=True, type="primary", key="cp_save_btn")
    with colB:
        limpar = st.button("↩️ Limpar", use_container_width=True, key="cp_clear_btn")

    if limpar:
        for k in ["cp_df_raw", "cp_grupo_nome", "cp_empresa_nome", "cp_empresa_row"]:
            st.session_state.pop(k, None)
        st.experimental_rerun()

    if salvar:
        if gsel=="— selecione —": st.error("Selecione o **Grupo**.")
        elif esel=="— selecione —": st.error("Selecione a **Empresa**.")
        elif df_raw.empty: st.error("Cole ou envie o arquivo.")
        else:
            st.session_state["cp_grupo_nome"]=gsel
            st.session_state["cp_empresa_nome"]=esel
            mask_g = df_emp["Grupo"].astype(str).apply(_norm)==_norm(gsel)
            mask_e = df_emp["Loja"].astype(str).apply(_norm)==_norm(esel)
            st.session_state["cp_empresa_row"]=df_emp[mask_g & mask_e].reset_index(drop=True)
            st.session_state["cp_df_raw"]=df_raw
            st.success("Pagar salvo em sessão.")

# --------- 🧾 CADASTRO Cliente/Fornecedor ---------
with aba_cad:
    st.subheader("Cadastro de Cliente / Fornecedor")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        gsel = st.selectbox("Grupo:", ["— selecione —"]+GRUPOS, key="cad_grupo")
    with col_g2:
        lojas = LOJAS_DO(gsel) if gsel!="— selecione —" else []
        esel = st.selectbox("Empresa:", ["— selecione —"]+lojas, key="cad_empresa")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        tipo = st.radio("Tipo", ["Cliente","Fornecedor"], horizontal=True)
        nome = st.text_input("Nome/Razão Social")
        doc  = st.text_input("CPF/CNPJ")
    with col2:
        email = st.text_input("E-mail")
        fone  = st.text_input("Telefone")
        obs   = st.text_area("Observações", height=80)

    colA, colB = st.columns([0.6,0.4])
    with colA:
        if st.button("💾 Salvar na sessão", use_container_width=True):
            st.session_state.setdefault("cadastros", []).append(
                {"Tipo":tipo,"Grupo":gsel,"Empresa":esel,"Nome":nome,"CPF/CNPJ":doc,"E-mail":email,"Telefone":fone,"Obs":obs}
            )
            st.success("Cadastro salvo localmente.")
    with colB:
        if st.button("🗂️ Enviar ao Google Sheets", use_container_width=True, type="primary"):
            try:
                sh = _open_planilha("Vendas diarias")
                if sh is None:
                    raise RuntimeError("Planilha indisponível")
                aba = "Cadastro Clientes" if tipo=="Cliente" else "Cadastro Fornecedores"
                try:
                    ws = sh.worksheet(aba)
                except WorksheetNotFound:
                    ws = sh.add_worksheet(aba, rows=1000, cols=20)
                    ws.append_row(["Tipo","Grupo","Empresa","Nome","CPF/CNPJ","E-mail","Telefone","Obs"])
                ws.append_row([tipo,gsel,esel,nome,doc,email,fone,obs])
                st.success(f"Salvo em {aba}.")
            except Exception as e:
                st.error(f"Erro ao salvar no Sheets: {e}")

    if st.session_state.get("cadastros"):
        st.markdown("#### Cadastros na sessão (não enviados)")
        st.dataframe(pd.DataFrame(st.session_state["cadastros"]), use_container_width=True, height=220)
