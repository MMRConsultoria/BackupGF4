# pages/CR- CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re, json, unicodedata
from io import StringIO, BytesIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# 🔒 Bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS (layout + seção compacta; sem “cards” brancos) =====
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
    .compact [data-testid="stFileUploader"] { margin-top:8px !important; }
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

# ======================
# Google Sheets
# ======================
def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    try:
        gc = gs_client()
    except Exception as e:
        st.warning(f"⚠️ Falha ao criar cliente Google: {e}")
        return None
    try:
        return gc.open(title)
    except Exception as e_title:
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sid:
            try:
                return gc.open_by_key(sid)
            except Exception as e_id:
                st.warning(f"⚠️ Não consegui abrir a planilha. Erros: {e_title} | {e_id}")
                return None
        st.warning(f"⚠️ Não consegui abrir por título. Detalhes: {e_title}")
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        df_vazio = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Código Grupo Everest","CNPJ"])
        return df_vazio, [], {}
    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning(f"⚠️ Erro lendo 'Tabela Empresa': {e}")
        df = pd.DataFrame(columns=["Grupo","Loja","Código Everest","Código Grupo Everest","CNPJ"])

    ren = {"Codigo Everest":"Código Everest","Codigo Grupo Everest":"Código Grupo Everest",
           "Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo"}
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest","CNPJ"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    lojas_map = (
        df.groupby("Grupo")["Loja"]
          .apply(lambda s: sorted(pd.Series(s.dropna().unique()).astype(str).tolist()))
          .to_dict()
    )
    return df, grupos, lojas_map

@st.cache_data(show_spinner=False)
def carregar_portadores():
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return [], {}
    try:
        ws = sh.worksheet("Portador")
    except Exception:
        return [], {}
    rows = ws.get_all_values()
    if not rows:
        return [], {}

    header = [str(h).strip() for h in rows[0]]

    def idx_of(names):
        for i, h in enumerate(header):
            if _norm(h) in names:
                return i
        return None

    i_banco = idx_of({"banco","banco/portador","nome banco"})
    i_porta = idx_of({"portador","nome portador"})

    bancos = set()
    mapa = {}
    for r in rows[1:]:
        b = str(r[i_banco]).strip() if (i_banco is not None and i_banco < len(r)) else ""
        p = str(r[i_porta]).strip()  if (i_porta is not None  and i_porta  < len(r)) else ""
        if b:
            bancos.add(b)
            if p: mapa[b] = p
    return sorted(bancos), mapa

@st.cache_data(show_spinner=False)
def carregar_tabela_meio_pagto():
    COL_PADRAO = "Padrão Cod Gerencial"
    COL_COD    = "Cod Gerencial Everest"
    COL_CNPJ   = "CNPJ Bandeira"

    sh = _open_planilha("Vendas diarias")
    if not sh:
        return pd.DataFrame(), []

    try:
        ws = sh.worksheet("Tabela Meio Pagamento")
    except WorksheetNotFound:
        st.warning("⚠️ Aba 'Tabela Meio Pagamento' não encontrada.")
        return pd.DataFrame(), []

    df = pd.DataFrame(ws.get_all_records())

    ren = {}
    for c in df.columns:
        n = _norm(c)
        if n in {"padrao cod gerencial","padrão cod gerencial","padrao","padrao gerencial"}:
            ren[c] = COL_PADRAO
        elif n in {"cod gerencial everest","codigo gerencial everest","cod_gerencial_everest"}:
            ren[c] = COL_COD
        elif n in {"cnpj bandeira","cnpj da bandeira","cnpj_bandeira"}:
            ren[c] = COL_CNPJ
        elif n == "meio de pagamento":
            ren[c] = "Meio de Pagamento"
    if ren: df = df.rename(columns=ren)

    for c in ["Meio de Pagamento", COL_PADRAO, COL_COD, COL_CNPJ]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    rules = []
    for _, row in df.iterrows():
        padrao = row[COL_PADRAO]
        codigo = row[COL_COD]
        cnpj   = row[COL_CNPJ]
        if not padrao or not codigo:
            continue
        tokens = [_norm(t) for t in re.split(r"[;,/|]", padrao) if str(t).strip()]
        if not tokens:
            continue
        rules.append({
            "tokens": tokens,
            "codigo_gerencial": codigo,
            "cnpj_bandeira": cnpj,
            "padrao_str": padrao
        })
    return df, rules

def _match_bandeira_to_gerencial(band_value: str):
    if not band_value or not MEIO_RULES:
        return "", "", ""
    txt = _norm(band_value)
    for rule in MEIO_RULES:
        for tok in rule["tokens"]:
            if tok and tok in txt:
                return rule["codigo_gerencial"], rule.get("cnpj_bandeira",""), rule.get("padrao_str","")
    return "", "", ""

# ===== Carregamentos base =====
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
DF_MEIO, MEIO_RULES = carregar_tabela_meio_pagto()

def LOJAS_DO(grupo_nome: str):
    return LOJAS_MAP.get(grupo_nome, [])

# ======================
# UI Components
# ======================
def filtros_grupo_empresa(prefix, with_portador=False, with_tipo_imp=False):
    """Grupo | Empresa | Banco | Tipo de Importação (lado a lado)."""
    c1, c2, c3, c4 = st.columns([1,1,1,1])

    with c1:
        gsel = st.selectbox("Grupo:", ["— selecione —"] + GRUPOS, key=f"{prefix}_grupo")
    with c2:
        lojas = LOJAS_DO(gsel) if gsel!="— selecione —" else []
        esel = st.selectbox("Empresa:", ["— selecione —"] + lojas, key=f"{prefix}_empresa")
    with c3:
        if with_portador:
            st.selectbox("Banco:", ["Todos"] + PORTADORES, index=0, key=f"{prefix}_portador")
        else:
            st.empty()
    with c4:
        if with_tipo_imp:
            st.selectbox("Tipo de Importação:", ["Todos","Adquirente","Cliente","Outros"], index=0, key=f"{prefix}_tipo_imp")
        else:
            st.empty()

    return gsel, esel

def bloco_colagem(prefix: str):
    # sem “cards” — limpo
    c1,c2 = st.columns([0.55,0.45])
    with c1:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=180,
                           placeholder="Cole aqui os dados copiados do Excel/Sheets…",
                           key=f"{prefix}_paste")
        df_paste = _try_parse_paste(txt) if (txt and txt.strip()) else pd.DataFrame()

    with c2:
        up = st.file_uploader("📎 Ou enviar arquivo (.xlsx/.xlsm/.xls/.csv)",
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
    if df_raw.empty:
        st.info("Cole ou envie um arquivo para visualizar.")
    else:
        # Altura ~4 linhas
        st.dataframe(df_raw, use_container_width=True, height=120)
    return df_raw

def _column_mapping_ui(prefix: str, df_raw: pd.DataFrame):
    st.markdown("##### Mapear colunas para **Adquirente**")
    cols = ["— selecione —"] + list(df_raw.columns)
    c1,c2,c3 = st.columns(3)
    with c1:
        st.selectbox("Coluna de **Data**", cols, key=f"{prefix}_col_data")
    with c2:
        st.selectbox("Coluna de **Valor**", cols, key=f"{prefix}_col_valor")
    with c3:
        st.selectbox("Coluna de **Bandeira**", cols, key=f"{prefix}_col_bandeira")

# ===== monta o Importador (mantendo ordem original + Falta CNPJ primeiro) ====
IMPORTADOR_ORDER = [
    "CNPJ/Cliente", "CNPJ Empresa", "Série Título", "Nº Título", "Nº Parcela",
    "Nº Documento", "Portador", "Data Documento", "Data Vencimento", "Data",
    "Valor Desconto", "Valor Multa", "Valor Juros Dia", "Valor Original",
    "Observações do Título", "Cód Conta Gerencial", "Cód Centro de Custo"
]

def _build_importador_df(df_raw: pd.DataFrame, prefix: str, grupo: str, loja: str,
                         banco_escolhido: str, tipo_imp: str):
    cd = st.session_state.get(f"{prefix}_col_data")
    cv = st.session_state.get(f"{prefix}_col_valor")
    cb = st.session_state.get(f"{prefix}_col_bandeira")

    if not cd or not cv or not cb or "— selecione —" in (cd, cv, cb):
        st.error("Defina **Data**, **Valor** e **Bandeira** para gerar o importador.")
        return pd.DataFrame()

    # CNPJ da loja
    cnpj_loja = ""
    if not df_emp.empty and loja:
        row = df_emp[
            (df_emp["Loja"].astype(str).str.strip() == loja) &
            (df_emp["Grupo"].astype(str).str.strip() == grupo)
        ]
        if not row.empty:
            cnpj_loja = str(row.iloc[0].get("CNPJ", "") or "")

    # Portador (nome) a partir do Banco selecionado
    portador_nome = MAPA_BANCO_PARA_PORTADOR.get(banco_escolhido, banco_escolhido or "")

    # dados do usuário (mantém a data exatamente como veio)
    data_original  = df_raw[cd].astype(str)
    valor_original = pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    bandeira_txt   = df_raw[cb].astype(str).str.strip()

    # mapeamento por bandeira usando tokens do Padrão
    cod_conta_list, cnpj_cli_list = [], []
    for b in bandeira_txt:
        cod, cnpj_band, _ = _match_bandeira_to_gerencial(b)
        cod_conta_list.append(cod)            # Cod Gerencial Everest
        cnpj_cli_list.append(cnpj_band)       # CNPJ da Bandeira

    # campos fixos
    serie_titulo   = "DRE"
    num_titulo     = ""
    num_parcela    = 1
    num_documento  = "DRE"
    centro_custo   = 3

    # Observação = texto da coluna Bandeira
    obs_list = bandeira_txt.tolist()

    out = pd.DataFrame({
       
        "CNPJ Empresa":          cnpj_loja,
        "Série Título":          serie_titulo,
        "Nº Título":             num_titulo,
        "Nº Parcela":            num_parcela,
        "Nº Documento":          num_documento,
        "CNPJ/Cliente":          cnpj_cli_list,
        "Portador":              portador_nome,
        "Data Documento":        data_original,
        "Data Vencimento":       data_original,
        "Data":                  data_original,
        "Valor Desconto":        0.00,
        "Valor Multa":           0.00,
        "Valor Juros Dia":       0.00,
        "Valor Original":        valor_original,
        "Observações do Título": obs_list,
        "Cód Conta Gerencial":   cod_conta_list,
        "Cód Centro de Custo":   centro_custo
    })

    # filtra linhas válidas
    out = out[(out["Data"].astype(str).str.strip() != "") & (out["Valor Original"].notna())]

    # reordena exatamente como o importador, depois insere a flag na frente
    out = out.reindex(columns=[c for c in IMPORTADOR_ORDER if c in out.columns])
    out.insert(0, "🔴 Falta CNPJ?", out["CNPJ/Cliente"].astype(str).str.strip().eq(""))

    return out

def _download_excel(df: pd.DataFrame, filename: str, label_btn: str):
    if df.empty: return
    to_save = df.copy()
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        to_save.to_excel(writer, index=False, sheet_name="Importador")
    bio.seek(0)
    st.download_button(label_btn, data=bio,
                       file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ======================
# ABAS
# ======================
aba_cr, aba_cp, aba_cad = st.tabs(["Contas a Receber", "Contas a Pagar", "🧾 Cadastro Cliente/Fornecedor"])

# --------- 💰 CONTAS A RECEBER ---------
with aba_cr:
    st.subheader("Contas a Receber")
    st.markdown('<div class="compact">', unsafe_allow_html=True)

    gsel, esel = filtros_grupo_empresa("cr", with_portador=True, with_tipo_imp=True)
    st.markdown('<hr class="compact">', unsafe_allow_html=True)

    df_raw = bloco_colagem("cr")

    if st.session_state.get("cr_tipo_imp") == "Adquirente" and not df_raw.empty:
        _column_mapping_ui("cr", df_raw)

    st.markdown('</div>', unsafe_allow_html=True)

    # --- Gerar / manter editor aberto usando session_state ---
    if st.button("🧾 Gerar Excel Importador (Receber)", use_container_width=True, key="cr_gen_btn"):
        if st.session_state.get("cr_tipo_imp") == "Adquirente":
            df_imp = _build_importador_df(
                df_raw, "cr",
                gsel if gsel!="— selecione —" else "",
                esel if esel!="— selecione —" else "",
                st.session_state.get("cr_portador",""),
                st.session_state.get("cr_tipo_imp","")
            )
            if not df_imp.empty:
                st.session_state["cr_df_imp"] = df_imp
        else:
            st.info("Selecione **Tipo de Importação = Adquirente** e mapeie as colunas.")

    if "cr_df_imp" in st.session_state and isinstance(st.session_state["cr_df_imp"], pd.DataFrame):
        df_imp = st.session_state["cr_df_imp"]

        show_only_missing = st.checkbox("Mostrar apenas linhas com 🔴 Falta CNPJ", value=False, key="cr_only_missing")
        df_view = df_imp[df_imp["🔴 Falta CNPJ?"]] if show_only_missing else df_imp

        # editáveis na visualização
        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]

        edited_cr = st.data_editor(
            df_view,
            disabled=disabled_cols,
            use_container_width=True,
            height=420,
            key="cr_editor"
        )

        # Atualiza DF completo e reavalia a flag; mantém ordem
        edited_full = df_imp.copy()
        edited_full.update(edited_cr)
        edited_full["🔴 Falta CNPJ?"] = edited_full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        # reordenar de novo para garantir consistência (flag primeiro + ordem do importador)
        cols = ["🔴 Falta CNPJ?"] + IMPORTADOR_ORDER
        edited_full = edited_full.reindex(columns=[c for c in cols if c in edited_full.columns])

        st.session_state["cr_df_imp"] = edited_full

        faltam = int(edited_full["🔴 Falta CNPJ?"].sum())
        total  = int(len(edited_full))
        if faltam:
            st.warning(f"⚠️ {faltam} de {total} linha(s) sem CNPJ/Cliente.")
        else:
            st.success("✅ Todos os CNPJs foram preenchidos.")

        _download_excel(edited_full, "Importador_Receber.xlsx", "📥 Baixar Importador (Receber)")

# --------- 💸 CONTAS A PAGAR ---------
with aba_cp:
    st.subheader("Contas a Pagar")
    st.markdown('<div class="compact">', unsafe_allow_html=True)

    gsel, esel = filtros_grupo_empresa("cp", with_portador=True, with_tipo_imp=True)
    st.markdown('<hr class="compact">', unsafe_allow_html=True)

    df_raw = bloco_colagem("cp")

    if st.session_state.get("cp_tipo_imp") == "Adquirente" and not df_raw.empty:
        _column_mapping_ui("cp", df_raw)

    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🧾 Gerar Excel Importador (Pagar)", use_container_width=True, key="cp_gen_btn"):
        if st.session_state.get("cp_tipo_imp") == "Adquirente":
            df_imp = _build_importador_df(
                df_raw, "cp",
                gsel if gsel!="— selecione —" else "",
                esel if esel!="— selecione —" else "",
                st.session_state.get("cp_portador",""),
                st.session_state.get("cp_tipo_imp","")
            )
            if not df_imp.empty:
                st.session_state["cp_df_imp"] = df_imp
        else:
            st.info("Selecione **Tipo de Importação = Adquirente** e mapeie as colunas.")

    if "cp_df_imp" in st.session_state and isinstance(st.session_state["cp_df_imp"], pd.DataFrame):
        df_imp = st.session_state["cp_df_imp"]

        show_only_missing = st.checkbox("Mostrar apenas linhas com 🔴 Falta CNPJ", value=False, key="cp_only_missing")
        df_view = df_imp[df_imp["🔴 Falta CNPJ?"]] if show_only_missing else df_imp

        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]

        edited_cp = st.data_editor(
            df_view,
            disabled=disabled_cols,
            use_container_width=True,
            height=420,
            key="cp_editor"
        )

        edited_full = df_imp.copy()
        edited_full.update(edited_cp)
        edited_full["🔴 Falta CNPJ?"] = edited_full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        cols = ["🔴 Falta CNPJ?"] + IMPORTADOR_ORDER
        edited_full = edited_full.reindex(columns=[c for c in cols if c in edited_full.columns])
        st.session_state["cp_df_imp"] = edited_full

        faltam = int(edited_full["🔴 Falta CNPJ?"].sum())
        total  = int(len(edited_full))
        if faltam:
            st.warning(f"⚠️ {faltam} de {total} linha(s) sem CNPJ/Cliente.")
        else:
            st.success("✅ Todos os CNPJs foram preenchidos.")

        _download_excel(edited_full, "Importador_Pagar.xlsx", "📥 Baixar Importador (Pagar)")

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
                if sh is None: raise RuntimeError("Planilha indisponível")
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
