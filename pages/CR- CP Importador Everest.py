# pages/CR- CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re, json, unicodedata
from io import StringIO, BytesIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

# ======= FUSÍVEL GLOBAL: remove qualquer ajuda ou docstring visível =======
try:
    import builtins
    def _noop_help(*args, **kwargs): 
        return None
    builtins.help = _noop_help
    st.help = lambda *a, **k: None
except Exception:
    pass

# ======= CONFIGURAÇÃO DA PÁGINA =======
st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")
st.set_option("client.showErrorDetails", False)

if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS =====
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
  hr.compact { height:1px; background:#e6e9f0; border:none; margin:8px 0 10px; }
  .compact [data-testid="stSelectbox"] { margin-bottom:6px !important; }
  .compact [data-testid="stTextArea"] { margin-top:8px !important; }
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
# Funções auxiliares
# ======================
def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII","ignore").decode("ASCII")

def _norm_basic(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+"," ", s).strip().lower()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    text = (text or "").strip("\n\r ")
    if not text: return pd.DataFrame()
    first = text.splitlines()[0] if text else ""
    if "\t" in first:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")
    df = df.dropna(how="all")
    df.columns = [str(c).strip() if str(c).strip() else f"col_{i}" for i,c in enumerate(df.columns)]
    return df

def _to_float_br(x):
    s = str(x or "").strip()
    s = s.replace("R$","").replace(" ","").replace(".","").replace(",",".")
    try: return float(s)
    except: return None

def _tokenize(txt: str):
    return [w for w in re.findall(r"[0-9a-zA-Z]+", _norm_basic(txt)) if w]

# ======================
# Google Sheets
# ======================
def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    gc = gs_client()
    try:
        return gc.open(title)
    except:
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        return gc.open_by_key(sid) if sid else None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    sh = _open_planilha("Vendas diarias")
    ws = sh.worksheet("Tabela Empresa")
    df = pd.DataFrame(ws.get_all_records())
    ren = {"Codigo Everest":"Código Everest","Codigo Grupo Everest":"Código Grupo Everest",
           "Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo"}
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest","CNPJ"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype(str).str.strip()
    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    lojas_map = df.groupby("Grupo")["Loja"].apply(lambda s: sorted(s.dropna().unique().tolist())).to_dict()
    return df, grupos, lojas_map

@st.cache_data(show_spinner=False)
def carregar_portadores():
    sh = _open_planilha("Vendas diarias")
    ws = sh.worksheet("Portador")
    rows = ws.get_all_values()
    if not rows: return [], {}
    header = [str(h).strip() for h in rows[0]]
    def idx_of(names):
        for i, h in enumerate(header):
            if _norm_basic(h) in names: return i
        return None
    i_banco = idx_of({"banco","banco/portador"})
    i_porta = idx_of({"portador"})
    bancos, mapa = set(), {}
    for r in rows[1:]:
        b = str(r[i_banco]).strip() if i_banco is not None else ""
        p = str(r[i_porta]).strip() if i_porta is not None else ""
        if b:
            bancos.add(b)
            if p: mapa[b] = p
    return sorted(bancos), mapa

@st.cache_data(show_spinner=False)
def carregar_tabela_meio_pagto():
    COL_PADRAO, COL_COD, COL_CNPJ = "Padrão Cod Gerencial","Cod Gerencial Everest","CNPJ Bandeira"
    sh = _open_planilha("Vendas diarias")
    ws = sh.worksheet("Tabela Meio Pagamento")
    df = pd.DataFrame(ws.get_all_records()).astype(str)
    for c in [COL_PADRAO,COL_COD,COL_CNPJ]: df[c] = df[c].astype(str).str.strip()
    rules=[]
    for _,r in df.iterrows():
        padrao=r[COL_PADRAO]; codigo=r[COL_COD]; cnpj=r[COL_CNPJ]
        if padrao and codigo:
            tokens=sorted(set(_tokenize(padrao)))
            if tokens: rules.append({"tokens":tokens,"codigo_gerencial":codigo,"cnpj_bandeira":cnpj})
    return df, rules

def _match_bandeira_to_gerencial(ref_text: str):
    if not ref_text or not MEIO_RULES: return "","",""
    ref_tokens=set(_tokenize(ref_text))
    best,best_hits,best_len=None,0,0
    for rule in MEIO_RULES:
        tokens=rule["tokens"]
        hits=sum(1 for t in tokens if t in ref_tokens)
        if hits>best_hits or (hits==best_hits and len(tokens)>best_len):
            best,best_hits,best_len=rule,hits,len(tokens)
    if best: return best["codigo_gerencial"],best.get("cnpj_bandeira",""),""
    return "","",""

# ===== Dados base =====
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
DF_MEIO, MEIO_RULES = carregar_tabela_meio_pagto()

def LOJAS_DO(g): return LOJAS_MAP.get(g,[])

IMPORTADOR_ORDER = [
    "CNPJ Empresa","Série Título","Nº Título","Nº Parcela","Nº Documento",
    "CNPJ/Cliente","Portador","Data Documento","Data Vencimento","Data",
    "Valor Desconto","Valor Multa","Valor Juros Dia","Valor Original",
    "Observações do Título","Cód Conta Gerencial","Cód Centro de Custo"
]

# ======================
# Componentes de UI
# ======================
def filtros_grupo_empresa(prefix, with_portador=False, with_tipo_imp=False):
    c1,c2,c3,c4 = st.columns([1,1,1,1])
    with c1:
        gsel = st.selectbox("Grupo:", ["— selecione —"]+GRUPOS, key=f"{prefix}_grupo")
    with c2:
        esel = st.selectbox("Empresa:", ["— selecione —"]+LOJAS_DO(gsel), key=f"{prefix}_empresa")
    with c3:
        if with_portador: st.selectbox("Banco:", ["Todos"]+PORTADORES, index=0, key=f"{prefix}_portador")
        else: st.empty()
    with c4:
        if with_tipo_imp: st.selectbox("Tipo de Importação:", ["Todos","Adquirente","Cliente","Outros"], index=0, key=f"{prefix}_tipo_imp")
        else: st.empty()
    return gsel, esel

def _on_paste_change(prefix):
    if not str(st.session_state.get(f"{prefix}_paste","")).strip():
        st.session_state.pop(f"{prefix}_df_imp",None)
        st.session_state.pop(f"{prefix}_edited_once",None)

def bloco_colagem(prefix):
    c1,c2 = st.columns([0.65,0.35])
    with c1:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=180,
            placeholder="Cole aqui os dados copiados do Excel/Sheets…",
            key=f"{prefix}_paste", on_change=_on_paste_change, args=(prefix,))
        df_paste = _try_parse_paste(txt) if txt else pd.DataFrame()
    with c2:
        show_prev = st.checkbox("Mostrar pré-visualização", value=False, key=f"{prefix}_show_prev")
        if show_prev and not df_paste.empty:
            st.dataframe(df_paste, use_container_width=True, height=120)
    return df_paste

def _column_mapping_ui(prefix, df_raw):
    cols = ["— selecione —"]+list(df_raw.columns)
    c1,c2,c3 = st.columns(3)
    with c1: st.selectbox("Coluna de **Data**", cols, key=f"{prefix}_col_data")
    with c2: st.selectbox("Coluna de **Valor**", cols, key=f"{prefix}_col_valor")
    with c3: st.selectbox("Coluna de **Referência**", cols, key=f"{prefix}_col_bandeira")

def _build_importador_df(df_raw, prefix, grupo, loja, banco):
    cd,cv,cb = st.session_state.get(f"{prefix}_col_data"),st.session_state.get(f"{prefix}_col_valor"),st.session_state.get(f"{prefix}_col_bandeira")
    if not cd or not cv or not cb or "— selecione —" in (cd,cv,cb): return pd.DataFrame()
    cnpj_loja=""
    row=df_emp[(df_emp["Loja"]==loja)&(df_emp["Grupo"]==grupo)]
    if not row.empty: cnpj_loja=str(row.iloc[0].get("CNPJ",""))
    portador_nome=MAPA_BANCO_PARA_PORTADOR.get(banco,banco)
    data=df_raw[cd].astype(str); valor=pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    ref=df_raw[cb].astype(str).str.strip()
    cods,cnpjs=[],[]
    for b in ref:
        cod,cnpj,_=_match_bandeira_to_gerencial(b)
        cods.append(cod); cnpjs.append(cnpj)
    out=pd.DataFrame({
        "CNPJ Empresa":cnpj_loja,"Série Título":"DRE","Nº Título":"",
        "Nº Parcela":1,"Nº Documento":"DRE","CNPJ/Cliente":cnpjs,"Portador":portador_nome,
        "Data Documento":data,"Data Vencimento":data,"Data":data,
        "Valor Desconto":0.00,"Valor Multa":0.00,"Valor Juros Dia":0.00,"Valor Original":valor,
        "Observações do Título":ref.tolist(),"Cód Conta Gerencial":cods,"Cód Centro de Custo":3
    })
    out=out[(out["Data"].astype(str)!="")&(out["Valor Original"].notna())]
    out.insert(0,"🔴 Falta CNPJ?",out["CNPJ/Cliente"].astype(str).str.strip().eq(""))
    return out

def _download_excel(df, filename, label, disabled=False):
    if df.empty: st.button(label, disabled=True, use_container_width=True); return
    bio=BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w: df.to_excel(w,index=False,sheet_name="Importador")
    bio.seek(0)
    st.download_button(label, data=bio, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, disabled=disabled)

# ======================
# ABAS
# ======================
aba_cr, aba_cp = st.tabs(["💰 Contas a Receber","💸 Contas a Pagar"])

# -------- CONTAS A RECEBER --------
with aba_cr:
    gsel,esel=filtros_grupo_empresa("cr",with_portador=True,with_tipo_imp=True)
    df_raw=bloco_colagem("cr")
    if st.session_state.get("cr_tipo_imp")=="Adquirente" and not df_raw.empty:
        _column_mapping_ui("cr",df_raw)
    ready=(
        st.session_state.get("cr_tipo_imp")=="Adquirente" and not df_raw.empty and
        all(st.session_state.get(k) and st.session_state.get(k)!="— selecione —"
            for k in["cr_col_data","cr_col_valor","cr_col_bandeira"])
        and gsel not in (None,"","— selecione —") and esel not in (None,"","— selecione —")
    )
    if ready:
        df_imp=_build_importador_df(df_raw,"cr",gsel,esel,st.session_state.get("cr_portador",""))
        st.session_state["cr_df_imp"]=df_imp
    df_imp=st.session_state.get("cr_df_imp")
    if isinstance(df_imp,pd.DataFrame) and not df_imp.empty:
        st.data_editor(df_imp,use_container_width=True,height=420,disabled=True)
        _download_excel(df_imp,"Importador_Receber.xlsx","📥 Baixar Importador (Receber)")

# -------- CONTAS A PAGAR --------
with aba_cp:
    gsel,esel=filtros_grupo_empresa("cp",with_portador=True,with_tipo_imp=True)
    df_raw=bloco_colagem("cp")
    if st.session_state.get("cp_tipo_imp")=="Adquirente" and not df_raw.empty:
        _column_mapping_ui("cp",df_raw)
    ready=(
        st.session_state.get("cp_tipo_imp")=="Adquirente" and not df_raw.empty and
        all(st.session_state.get(k) and st.session_state.get(k)!="— selecione —"
            for k in["cp_col_data","cp_col_valor","cp_col_bandeira"])
        and gsel not in (None,"","— selecione —") and esel not in (None,"","— selecione —")
    )
    if ready:
        df_imp=_build_importador_df(df_raw,"cp",gsel,esel,st.session_state.get("cp_portador",""))
        st.session_state["cp_df_imp"]=df_imp
    df_imp=st.session_state.get("cp_df_imp")
    if isinstance(df_imp,pd.DataFrame) and not df_imp.empty:
        st.data_editor(df_imp,use_container_width=True,height=420,disabled=True)
        _download_excel(df_imp,"Importador_Pagar.xlsx","📥 Baixar Importador (Pagar)")
