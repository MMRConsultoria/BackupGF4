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

# ======================
# Estilo (replica o visual do seu modelo)
# ======================
st.markdown("""
<style>
.stApp { background:#f9fafb; }
[data-testid="stToolbar"] { visibility:hidden;height:0;position:fixed; }
/* Cabeçalho */
.hwrap{display:flex;align-items:center;gap:12px;margin:4px 0 10px}
.hwrap h1{margin:0;font-size:38px;font-weight:800;letter-spacing:.2px}
/* Pill bar */
.pillbar{display:flex;gap:10px;margin:14px 0 16px}
.pill{
  border:1px solid #e5e7eb;background:#eef2ff;color:#374151;
  border-radius:12px;padding:10px 14px;font-weight:700;cursor:pointer;
}
.pill.active{background:#0b5bd3;color:#fff;box-shadow:0 1px 0 #0b5bd3}
.pill.muted{background:#f3f4f6!important;color:#6b7280}
.pill:hover{filter:brightness(0.96)}
/* Linha de filtros (labels pequenas + selects grandes) */
.frow{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:16px;margin:6px 0 8px}
.flabel{font-size:13px;color:#6b7280;margin-bottom:6px}
.fslot{background:#f3f6fb;border:1px solid #e5e7f0;border-radius:10px;padding:8px 10px}
hr{border:none;height:1px;background:#e5e7eb;margin:12px 0}
</style>
""", unsafe_allow_html=True)

# 🔒 login
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ======================
# Helpers
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
# Google Sheets (robusto)
# ======================
@st.cache_data(show_spinner=False)
def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    credentials_dict = json.loads(secret) if isinstance(secret,str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    gc = gs_client()
    try:
        return gc.open(title)
    except Exception:
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sid: return gc.open_by_key(sid)
        raise

@st.cache_data(show_spinner=False)
def carregar_empresas():
    sh = _open_planilha("Vendas diarias")
    df = pd.DataFrame(sh.worksheet("Tabela Empresa").get_all_records())
    # normalizações simples
    ren = {
        "Codigo Everest":"Código Everest","Codigo Grupo Everest":"Código Grupo Everest",
        "Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo",
    }
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest"]:
        if c not in df.columns: df[c]=""
        df[c]=df[c].astype(str).str.strip()
    df = df[df["Grupo"]!=""]
    # listas prontas
    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    def lojas_do_grupo(g):
        m = df["Grupo"].astype(str).apply(_norm)==_norm(g)
        return sorted(df.loc[m,"Loja"].astype(str).dropna().unique().tolist())
    return df, grupos, lojas_do_grupo

df_emp, GRUPOS, LOJAS_DO = carregar_empresas()

# ======================
# Header
# ======================
st.markdown("""
<div class="hwrap">
  <img src="https://img.icons8.com/color/48/graph.png" width="40"/>
  <h1>Relatório CR-CP Everest</h1>
</div>
""", unsafe_allow_html=True)

# ======================
# Pill "abas" (um layout, sem st.tabs)
# ======================
if "view" not in st.session_state:
    st.session_state.view = "CR"  # CR | CP | CAD

colA, colB, colC = st.columns([0.22,0.22,0.56])
with colA:
    if st.button("💰 Analise Receber", use_container_width=True,
                 key="pill_cr"):
        st.session_state.view = "CR"
        st.rerun()
with colB:
    if st.button("💸 Analise Pagar", use_container_width=True,
                 key="pill_cp"):
        st.session_state.view = "CP"
        st.rerun()
with colC:
    if st.button("🧾 Cadastro", use_container_width=True,
                 key="pill_cad"):
        st.session_state.view = "CAD"
        st.rerun()

# pintar como ativo
st.markdown(f"""
<script>
const pills = Array.from(parent.document.querySelectorAll('button[kind="secondary"]'));
if (pills && pills.length>=3){{
  const v = "{st.session_state.view}";
  const map={{"CR":0,"CP":1,"CAD":2}};
  pills.forEach((b,i)=>{{b.classList.add('pill'); b.classList.remove('active');}});
  pills[map[v]].classList.add('active');
}}
</script>
""", unsafe_allow_html=True)

st.markdown("<hr/>", unsafe_allow_html=True)

# ======================
# Filtros linha (layout como seu print)
# ======================
def filtros_grupo_empresa(prefix: str):
    st.markdown('<div class="frow">', unsafe_allow_html=True)
    # Grupo
    st.markdown('<div class="fslot">', unsafe_allow_html=True)
    st.markdown('<div class="flabel">Grupo:</div>', unsafe_allow_html=True)
    gsel = st.selectbox("", ["— selecione —"]+GRUPOS, key=f"{prefix}_grupo", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Empresa
    st.markdown('<div class="fslot">', unsafe_allow_html=True)
    st.markdown('<div class="flabel">Empresa:</div>', unsafe_allow_html=True)
    lojas = LOJAS_DO(gsel) if gsel!="— selecione —" else []
    esel = st.selectbox("", ["— selecione —"]+lojas, key=f"{prefix}_empresa", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Placeholder 1 (Visão)
    st.markdown('<div class="fslot">', unsafe_allow_html=True)
    st.markdown('<div class="flabel">Visão:</div>', unsafe_allow_html=True)
    vis = st.selectbox("", ["Por Empresa"], key=f"{prefix}_visao", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Placeholder 2 (Tipo)
    st.markdown('<div class="fslot">', unsafe_allow_html=True)
    st.markdown('<div class="flabel">Tipo:</div>', unsafe_allow_html=True)
    tip = st.selectbox("", ["TODOS"], key=f"{prefix}_tipo", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)  # fecha frow
    return gsel, esel

# ======================
# Views
# ======================
def bloco_colagem(prefix: str):
    c1,c2 = st.columns([0.55,0.45])
    with c1:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                           placeholder="Cole aqui os dados copiados do Excel/Sheets…",
                           key=f"{prefix}_paste")
        df_paste = _try_parse_paste(txt) if txt.strip() else pd.DataFrame()
    with c2:
        up = st.file_uploader("📎 Ou enviar arquivo (.xlsx/.xlsm/.xls/.csv)", 
                              type=["xlsx","xlsm","xls","csv"], key=f"{prefix}_file")
        df_file = pd.DataFrame()
        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    import pandas as pd
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

# --- Receber ---
if st.session_state.view == "CR":
    st.subheader("💰 Contas a Receber")
    gsel, esel = filtros_grupo_empresa("cr")
    st.markdown("<hr/>", unsafe_allow_html=True)
    df_raw = bloco_colagem("cr")
    btn_save = st.button("✅ Salvar seleção e dados (Receber)", use_container_width=True, type="primary")
    if btn_save:
        if gsel=="— selecione —": st.error("Selecione o **Grupo**.")
        elif esel=="— selecione —": st.error("Selecione a **Empresa**.")
        elif df_raw.empty: st.error("Cole ou envie o arquivo.")
        else:
            st.session_state["cr_grupo_nome"]=gsel
            st.session_state["cr_empresa_nome"]=esel
            st.session_state["cr_df_raw"]=df_raw
            st.success("Receber salvo em sessão. Pronto para o mapeamento/integração.")

# --- Pagar ---
elif st.session_state.view == "CP":
    st.subheader("💸 Contas a Pagar")
    gsel, esel = filtros_grupo_empresa("cp")
    st.markdown("<hr/>", unsafe_allow_html=True)
    df_raw = bloco_colagem("cp")
    btn_save = st.button("✅ Salvar seleção e dados (Pagar)", use_container_width=True, type="primary")
    if btn_save:
        if gsel=="— selecione —": st.error("Selecione o **Grupo**.")
        elif esel=="— selecione —": st.error("Selecione a **Empresa**.")
        elif df_raw.empty: st.error("Cole ou envie o arquivo.")
        else:
            st.session_state["cp_grupo_nome"]=gsel
            st.session_state["cp_empresa_nome"]=esel
            st.session_state["cp_df_raw"]=df_raw
            st.success("Pagar salvo em sessão. Pronto para o mapeamento/integração.")

# --- Cadastro ---
else:
    st.subheader("🧾 Cadastro Cliente/Fornecedor")
    gsel, esel = filtros_grupo_empresa("cad")
    st.markdown("<hr/>", unsafe_allow_html=True)
    col1,col2 = st.columns(2)
    with col1:
        tipo = st.radio("Tipo", ["Cliente","Fornecedor"], horizontal=True)
        nome = st.text_input("Nome/Razão Social")
        doc  = st.text_input("CPF/CNPJ")
    with col2:
        email = st.text_input("E-mail")
        fone  = st.text_input("Telefone")
        obs   = st.text_area("Observações", height=80)
    cA,cB = st.columns([0.6,0.4])
    with cA:
        if st.button("💾 Salvar na sessão", use_container_width=True):
            st.session_state.setdefault("cadastros", []).append(
                {"Tipo":tipo,"Grupo":gsel,"Empresa":esel,"Nome":nome,"CPF/CNPJ":doc,"E-mail":email,"Telefone":fone,"Obs":obs}
            )
            st.success("Cadastro salvo localmente.")
    with cB:
        if st.button("🗂️ Enviar ao Google Sheets", use_container_width=True, type="primary"):
            try:
                sh = _open_planilha("Vendas diarias")
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
