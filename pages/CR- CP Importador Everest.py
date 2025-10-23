# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re, json, unicodedata
from io import StringIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# 🔒 bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ======= CSS =======
st.markdown("""
<style>
[data-testid="stToolbar"]{visibility:hidden;height:0;position:fixed;}
.stApp{background-color:#f9f9f9;}
button[data-baseweb="tab"]{background:#f0f2f6;border-radius:10px;padding:10px 20px;margin-right:8px;font-weight:600;}
button[data-baseweb="tab"]:hover{background:#dce0ea;color:black;}
button[data-baseweb="tab"][aria-selected="true"]{background:#0366d6;color:white;}
hr.compact{height:1px;background:#e6e9f0;border:none;margin:8px 0 10px;}
.compact [data-testid="stSelectbox"]{margin-bottom:6px!important;}
.compact [data-testid="stFileUploader"]{margin-top:8px!important;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='display:flex;align-items:center;gap:10px;margin-bottom:12px;'>
<img src='https://img.icons8.com/color/48/graph.png' width='40'/>
<h1 style='margin:0;'>CR-CP Importador Everest</h1>
</div>
""", unsafe_allow_html=True)

# ======= Helpers =======
def _strip_accents_keep_case(s): 
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII","ignore").decode("ASCII")

def _norm(s): 
    return re.sub(r"\s+"," ",_strip_accents_keep_case(s)).strip().lower()

def _try_parse_paste(txt):
    txt=(txt or "").strip()
    if not txt: return pd.DataFrame()
    if "\t" in txt.splitlines()[0]:
        df=pd.read_csv(StringIO(txt),sep="\t",dtype=str,engine="python")
    else:
        for sep in ["; ", ";", ","]:
            try: df=pd.read_csv(StringIO(txt),sep=sep.strip(),dtype=str,engine="python");break
            except: continue
    df=df.dropna(how="all")
    df.columns=[str(c).strip() for c in df.columns]
    return df

# ======= Google Sheets =======
def gs_client():
    scope=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    secret=st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    credentials_dict=json.loads(secret) if isinstance(secret,str) else dict(secret)
    creds=ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict,scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    try:
        gc=gs_client()
        return gc.open(title)
    except Exception as e:
        sid=st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        return gc.open_by_key(sid) if sid else None

@st.cache_data
def carregar_empresas():
    sh=_open_planilha("Vendas diarias")
    ws=sh.worksheet("Tabela Empresa")
    df=pd.DataFrame(ws.get_all_records())
    ren={"Codigo Everest":"Código Everest","Codigo Grupo Everest":"Código Grupo Everest"}
    df=df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest"]:
        if c not in df: df[c]=""
    grupos=sorted(df["Grupo"].dropna().unique())
    lojas_map=df.groupby("Grupo")["Loja"].apply(lambda s:sorted(s.dropna().unique())).to_dict()
    return df,grupos,lojas_map

@st.cache_data
def carregar_portadores():
    sh=_open_planilha("Vendas diarias")
    ws=sh.worksheet("Portador")
    rows=ws.get_all_values()
    header=[h.strip() for h in rows[0]]
    col_idx=next((i for i,h in enumerate(header) if _norm(h) in {"banco","portador"}),None)
    if col_idx is None: return []
    bancos={r[col_idx].strip() for r in rows[1:] if len(r)>col_idx and r[col_idx].strip()!=""}
    return sorted(bancos)

df_emp,GRUPOS,LOJAS_MAP=carregar_empresas()
PORTADORES=carregar_portadores()
def LOJAS_DO(g): return LOJAS_MAP.get(g,[])

# ======= Filtros linha única =======
def filtros_grupo_empresa(prefix,with_portador=False,with_tipo_imp=False):
    if with_portador and with_tipo_imp:
        c1,c2,c3,c4=st.columns([1,1,1,1])
    elif with_portador:
        c1,c2,c3=st.columns([1,1,1]);c4=None
    elif with_tipo_imp:
        c1,c2,c4=st.columns([1,1,1]);c3=None
    else:
        c1,c2=st.columns([1,1]);c3=c4=None
    with c1:
        gsel=st.selectbox("Grupo:",["— selecione —"]+GRUPOS,key=f"{prefix}_grupo")
    with c2:
        lojas=LOJAS_DO(gsel) if gsel!="— selecione —" else []
        esel=st.selectbox("Empresa:",["— selecione —"]+lojas,key=f"{prefix}_empresa")
    if with_portador and c3:
        st.selectbox("Portador (Banco):",["Todos"]+PORTADORES,index=0,key=f"{prefix}_portador")
    if with_tipo_imp and c4:
        st.selectbox("Tipo de Importação:",["Todos","Adquirente","Cliente","Outros"],index=0,key=f"{prefix}_tipo_imp")
    return gsel,esel

# ======= bloco colagem/upload =======
def bloco_colagem(prefix):
    c1,c2=st.columns([0.55,0.45])
    with c1:
        txt=st.text_area("📋 Colar tabela (Ctrl+V)",height=220,key=f"{prefix}_paste")
        df_paste=_try_parse_paste(txt)
    with c2:
        up=st.file_uploader("📎 Ou enviar arquivo (.xlsx/.xls/.csv)",type=["xlsx","xls","csv"],key=f"{prefix}_file")
        df_file=pd.DataFrame()
        if up:
            if up.name.lower().endswith(".csv"):
                try: df_file=pd.read_csv(up,sep=";",dtype=str)
                except: up.seek(0);df_file=pd.read_csv(up,sep=",",dtype=str)
            else: df_file=pd.read_excel(up,dtype=str)
            df_file=df_file.dropna(how="all")
            df_file.columns=[str(c).strip() for c in df_file.columns]
    df=df_paste if not df_paste.empty else df_file
    st.markdown("#### Pré-visualização")
    st.dataframe(df if not df.empty else pd.DataFrame(),use_container_width=True,height=300)
    return df

# ======= mapeamento adquirente =======
def _guess_defaults_cols(df):
    cols=[str(c) for c in df.columns]
    def f(pat):
        for c in cols:
            if any(re.search(p,str(c).lower()) for p in pat): return c
        return None
    return f([r"data",r"date"]), f([r"valor",r"total",r"vlr"]), f([r"bandeira",r"visa",r"master"])

def mapping_minimo_adquirente(prefix,df):
    st.markdown("##### Mapear colunas mínimas (Adquirente)")
    cols=[str(c) for c in df.columns]
    d1,d2,d3=_guess_defaults_cols(df)
    c0,c1,c2=st.columns(3)
    with c0:
        st.selectbox("Coluna de Data",["— selecione —"]+cols,
                     index=(cols.index(d1)+1) if d1 in cols else 0,key=f"{prefix}_col_data")
    with c1:
        st.selectbox("Coluna de Valor",["— selecione —"]+cols,
                     index=(cols.index(d2)+1) if d2 in cols else 0,key=f"{prefix}_col_valor")
    with c2:
        st.selectbox("Coluna de Bandeira",["— selecione —"]+cols,
                     index=(cols.index(d3)+1) if d3 in cols else 0,key=f"{prefix}_col_bandeira")
    cd,cv,cb=[st.session_state.get(f"{prefix}_{x}") for x in ["col_data","col_valor","col_bandeira"]]
    if all([cd,cv,cb]) and "— selecione —" not in [cd,cv,cb]:
        prev=pd.DataFrame({
            "Data":pd.to_datetime(df[cd],dayfirst=True,errors="coerce"),
            "Valor (R$)":pd.to_numeric(df[cv].astype(str)
                .str.replace("R$","").str.replace(".","").str.replace(",","."),errors="coerce"),
            "Bandeira":df[cb].astype(str).str.strip()
        })
        st.dataframe(prev.head(15),use_container_width=True,height=220)

# ======= ABAS =======
aba_cr,aba_cp,aba_cad=st.tabs(["💰 Contas a Receber","💸 Contas a Pagar","🧾 Cadastro Cliente/Fornecedor"])

# ---------- RECEBER ----------
with aba_cr:
    st.subheader("Contas a Receber")
    st.markdown('<div class="compact">',unsafe_allow_html=True)
    gsel,esel=filtros_grupo_empresa("cr",with_portador=True,with_tipo_imp=True)
    tipo_imp=st.session_state.get("cr_tipo_imp","Todos")
    st.markdown('<hr class="compact">',unsafe_allow_html=True)
    df_raw=bloco_colagem("cr")
    if tipo_imp=="Adquirente" and not df_raw.empty:
        mapping_minimo_adquirente("cr",df_raw)
    st.markdown('</div>',unsafe_allow_html=True)

    cA,cB=st.columns([0.6,0.4])
    if cB.button("↩️ Limpar",use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith("cr_"): st.session_state.pop(k,None)
        st.experimental_rerun()
    if cA.button("✅ Salvar seleção e dados (Receber)",use_container_width=True,type="primary"):
        if gsel=="— selecione —" or esel=="— selecione —":
            st.error("Selecione Grupo e Empresa.")
        elif df_raw.empty:
            st.error("Cole ou envie o arquivo.")
        elif tipo_imp=="Adquirente":
            cd,cv,cb=[st.session_state.get(f"cr_{x}") for x in ["col_data","col_valor","col_bandeira"]]
            if not cd or "— selecione —" in [cd,cv,cb]:
                st.error("Defina Data, Valor e Bandeira para Adquirente.")
            else:
                st.session_state["cr_df_raw"]=df_raw;st.success("Receber salvo.")
        else:
            st.session_state["cr_df_raw"]=df_raw;st.success("Receber salvo.")

# ---------- PAGAR ----------
with aba_cp:
    st.subheader("Contas a Pagar")
    st.markdown('<div class="compact">',unsafe_allow_html=True)
    gsel,esel=filtros_grupo_empresa("cp",with_portador=True,with_tipo_imp=True)
    tipo_imp=st.session_state.get("cp_tipo_imp","Todos")
    st.markdown('<hr class="compact">',unsafe_allow_html=True)
    df_raw=bloco_colagem("cp")
    if tipo_imp=="Adquirente" and not df_raw.empty:
        mapping_minimo_adquirente("cp",df_raw)
    st.markdown('</div>',unsafe_allow_html=True)

    cA,cB=st.columns([0.6,0.4])
    if cB.button("↩️ Limpar ",use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith("cp_"): st.session_state.pop(k,None)
        st.experimental_rerun()
    if cA.button("✅ Salvar seleção e dados (Pagar)",use_container_width=True,type="primary"):
        if gsel=="— selecione —" or esel=="— selecione —":
            st.error("Selecione Grupo e Empresa.")
        elif df_raw.empty:
            st.error("Cole ou envie o arquivo.")
        elif tipo_imp=="Adquirente":
            cd,cv,cb=[st.session_state.get(f"cp_{x}") for x in ["col_data","col_valor","col_bandeira"]]
            if not cd or "— selecione —" in [cd,cv,cb]:
                st.error("Defina Data, Valor e Bandeira para Adquirente.")
            else:
                st.session_state["cp_df_raw"]=df_raw;st.success("Pagar salvo.")
        else:
            st.session_state["cp_df_raw"]=df_raw;st.success("Pagar salvo.")

# ---------- CADASTRO ----------
with aba_cad:
    st.subheader("Cadastro Cliente/Fornecedor")
    g1,g2=st.columns(2)
    with g1: gsel=st.selectbox("Grupo:",["— selecione —"]+GRUPOS,key="cad_grupo")
    with g2: esel=st.selectbox("Empresa:",["— selecione —"]+LOJAS_DO(gsel),key="cad_empresa")
    st.divider()
    c1,c2=st.columns(2)
    with c1:
        tipo=st.radio("Tipo",["Cliente","Fornecedor"],horizontal=True)
        nome=st.text_input("Nome/Razão Social")
        doc=st.text_input("CPF/CNPJ")
    with c2:
        email=st.text_input("E-mail");fone=st.text_input("Telefone");obs=st.text_area("Observações",height=80)
    cA,cB=st.columns([0.6,0.4])
    if cA.button("💾 Salvar na sessão",use_container_width=True):
        st.session_state.setdefault("cadastros",[]).append(
            {"Tipo":tipo,"Grupo":gsel,"Empresa":esel,"Nome":nome,"CPF/CNPJ":doc,"E-mail":email,"Telefone":fone,"Obs":obs})
        st.success("Cadastro salvo localmente.")
    if cB.button("🗂️ Enviar ao Google Sheets",use_container_width=True,type="primary"):
        try:
            sh=_open_planilha("Vendas diarias")
            aba="Cadastro Clientes" if tipo=="Cliente" else "Cadastro Fornecedores"
            try: ws=sh.worksheet(aba)
            except WorksheetNotFound:
                ws=sh.add_worksheet(aba,rows=1000,cols=20)
                ws.append_row(["Tipo","Grupo","Empresa","Nome","CPF/CNPJ","E-mail","Telefone","Obs"])
            ws.append_row([tipo,gsel,esel,nome,doc,email,fone,obs])
            st.success(f"Salvo em {aba}.")
        except Exception as e: st.error(f"Erro: {e}")
    if st.session_state.get("cadastros"):
        st.dataframe(pd.DataFrame(st.session_state["cadastros"]),use_container_width=True,height=220)
