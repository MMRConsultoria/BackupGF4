# pages/CR-CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re, json, unicodedata, math
from io import StringIO, BytesIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")

# 🔒 Bloqueio de acesso (mesmo padrão do seu app)
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS (visual + compactação) =====
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

        /* multiselect sem tags coloridas */
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] { background-color: transparent !important; border: none !important; color: black !important; }
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] * { color: black !important; fill: black !important; }
        div[data-testid="stMultiSelect"] > div { background-color: transparent !important; }

        /* compactar blocos de filtros/colagem */
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

# =============================================================================
# Helpers
# =============================================================================
def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII","ignore").decode("ASCII")

def _norm(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+"," ", s).strip().lower()
    return s

def _try_parse_paste(text: str) -> pd.DataFrame:
    text = (text or "").strip("\n\r ")
    if not text: 
        return pd.DataFrame()
    first = text.splitlines()[0] if "\n" in text else text
    # tenta identificar separador
    for sep in ["\t",";","; ",",","|"]:
        if sep.strip() and sep in first:
            try:
                df = pd.read_csv(StringIO(text), sep=sep.strip(), dtype=str, engine="python")
                return df.dropna(how="all")
            except Exception:
                pass
    # fallback: tab
    try:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
        return df.dropna(how="all")
    except Exception:
        return pd.DataFrame()

def _to_float_br(x):
    s = str(x or "").strip()
    s = s.replace("R$","").replace(" ","").replace(".","").replace(",",".")
    try:
        return float(s)
    except:
        return math.nan

# =============================================================================
# Google Sheets
# =============================================================================
def gs_client():
    """Cria o client do Google (Service Account)"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    gc = gs_client()
    try:
        return gc.open(title)
    except Exception:
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sid:
            try:
                return gc.open_by_key(sid)
            except Exception:
                return None
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    """
    Lê 'Tabela Empresa' e retorna:
      df_empresa (com colunas: Grupo, Loja, Código Everest, Código Grupo Everest, CNPJ),
      lista de grupos,
      dict {grupo: [lojas]}
    """
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        st.warning("⚠️ Não consegui abrir a planilha 'Vendas diarias'.")
        return pd.DataFrame(), [], {}

    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning(f"⚠️ Falha ao ler 'Tabela Empresa': {e}")
        return pd.DataFrame(), [], {}

    # normalizações
    ren = {
        "Codigo Everest":"Código Everest",
        "Codigo Grupo Everest":"Código Grupo Everest",
        "Cnpj":"CNPJ", "CNPJ ":"CNPJ", "CNPJ  ":"CNPJ", "CNPJ Empresa":"CNPJ",
        "Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo",
    }
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Código Everest","Código Grupo Everest","CNPJ"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    grupos = sorted([g for g in df["Grupo"].dropna().unique() if str(g).strip()])
    lojas_map = (
        df.groupby("Grupo")["Loja"]
          .apply(lambda s: sorted(set([str(x).strip() for x in s if str(x).strip()])))
          .to_dict()
    )
    return df, grupos, lojas_map

@st.cache_data(show_spinner=False)
def carregar_portadores():
    """
    Lê a aba 'Portador' e retorna os nomes únicos da coluna Banco/Portador.
    """
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return []
    try:
        ws = sh.worksheet("Portador")
    except Exception:
        return []
    rows = ws.get_all_values()
    if not rows:
        return []
    header = [str(h).strip() for h in rows[0]]
    col_idx = None
    for i,h in enumerate(header):
        if _norm(h) in {"banco","portador","nome banco","banco/portador"}:
            col_idx = i
            break
    if col_idx is None:
        st.warning("⚠️ Na aba 'Portador' não encontrei a coluna Banco/Portador.")
        return []
    bancos = {
        str(r[col_idx]).strip()
        for r in rows[1:]
        if len(r) > col_idx and str(r[col_idx]).strip() != ""
    }
    return sorted(bancos)

@st.cache_data(show_spinner=False)
def carregar_tabela_meio_pagto():
    """
    Lê 'Tabela Meio Pagamento' e cria regras:
      - 'Padrão Cod Gerencial' -> tokens (palavras-chave para procurar na Bandeira)
      - 'Cod Gerencial Everest' -> valor a gravar em 'Cód Conta Gerencial'
    """
    COL_PADRAO = "Padrão Cod Gerencial"
    COL_COD    = "Cod Gerencial Everest"

    sh = _open_planilha("Vendas diarias")
    if not sh:
        st.warning("⚠️ Sem planilha para Tabela Meio Pagamento.")
        return pd.DataFrame(), []

    try:
        ws = sh.worksheet("Tabela Meio Pagamento")
    except WorksheetNotFound:
        st.warning("⚠️ Aba 'Tabela Meio Pagamento' não encontrada.")
        return pd.DataFrame(), []

    df = pd.DataFrame(ws.get_all_records())
    # garante colunas
    if "Meio de Pagamento" not in df.columns:
        df["Meio de Pagamento"] = ""
    if COL_PADRAO not in df.columns:
        # tenta variantes
        alt = [c for c in df.columns if _norm(c) in {"padrao cod gerencial","padrão cod gerencial","padrao","padrao gerencial"}]
        df[COL_PADRAO] = df[alt[0]] if alt else ""
    if COL_COD not in df.columns:
        alt2 = [c for c in df.columns if _norm(c) in {"cod gerencial everest","codigo gerencial everest","cod_gerencial_everest"}]
        df[COL_COD] = df[alt2[0]] if alt2 else ""

    df["Meio de Pagamento"] = df["Meio de Pagamento"].astype(str).str.strip()
    df[COL_PADRAO] = df[COL_PADRAO].astype(str).str.strip()
    df[COL_COD]    = df[COL_COD].astype(str).str.strip()

    rules = []
    for _, row in df.iterrows():
        padrao = row[COL_PADRAO]
        codigo = row[COL_COD]
        meio   = row["Meio de Pagamento"]
        if not padrao or not codigo:
            continue
        tokens = [ _norm(t) for t in re.split(r"[;,/|]", padrao) if str(t).strip() ]
        if not tokens:
            continue
        rules.append({
            "tokens": tokens,                # palavras-chave a buscar
            "codigo_gerencial": codigo,      # valor final para 'Cód Conta Gerencial'
            "meio": meio                     # opcional (auditoria)
        })
    return df, rules

# ======================
# Carregamentos base
# ======================
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES = carregar_portadores()
DF_MEIO, MEIO_RULES = carregar_tabela_meio_pagto()

def LOJAS_DO(grupo_nome: str):
    return LOJAS_MAP.get(grupo_nome, [])

# =============================================================================
# UI: filtros lado a lado (Grupo | Empresa | Portador | Tipo de Importação)
# =============================================================================
def filtros_grupo_empresa(prefix, with_portador=False, with_tipo_imp=False):
    n = 2 + int(with_portador) + int(with_tipo_imp)
    cols = st.columns([1]*n)
    i = 0
    with cols[i]:
        gsel = st.selectbox("Grupo:", ["— selecione —"] + GRUPOS, key=f"{prefix}_grupo")
    i += 1
    with cols[i]:
        lojas = LOJAS_DO(gsel) if gsel != "— selecione —" else []
        esel  = st.selectbox("Empresa:", ["— selecione —"] + lojas, key=f"{prefix}_empresa")
    i += 1
    if with_portador:
        with cols[i]:
            st.selectbox("Portador (Banco):", ["Todos"] + PORTADORES, index=0, key=f"{prefix}_portador")
        i += 1
    if with_tipo_imp:
        with cols[i]:
            st.selectbox("Tipo de Importação:", ["Todos","Adquirente","Cliente","Outros"], index=0, key=f"{prefix}_tipo_imp")
    return gsel, esel

# =============================================================================
# Colagem / Upload
# =============================================================================
def bloco_colagem(prefix: str):
    c1, c2 = st.columns([0.55, 0.45])
    with c1:
        txt = st.text_area("📋 Colar tabela (Ctrl+V)", height=220,
                           placeholder="Cole aqui os dados copiados do Excel/Sheets…",
                           key=f"{prefix}_paste")
        df_paste = _try_parse_paste(txt)
    with c2:
        up = st.file_uploader("📎 Ou enviar arquivo (.xlsx/.xls/.csv)",
                              type=["xlsx","xls","csv"], key=f"{prefix}_file")
        df_file = pd.DataFrame()
        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    try:
                        df_file = pd.read_csv(up, sep=";", dtype=str)
                    except Exception:
                        up.seek(0); df_file = pd.read_csv(up, sep=",", dtype=str)
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
        st.dataframe(df_raw, use_container_width=True, height=300)
    return df_raw

# =============================================================================
# Mapeamento mínimo (Adquirente): Data / Valor / Bandeira
# =============================================================================
def _guess_defaults_cols(df: pd.DataFrame):
    cols = [str(c) for c in df.columns]
    def find(pats):
        for c in cols:
            n = _norm(c)
            if any(re.search(p, n) for p in pats):
                return c
        return None
    col_data = find([r"\bdata\b", r"\bdate\b", r"\bcompet", r"\bdt"])
    col_val  = find([r"\bvalor\b", r"\bvlr\b", r"\btotal\b", r"\bmont", r"\bamount\b", r"\bbruto\b", r"\bliq"])
    col_band = find([r"\bbandeira\b", r"\bband\b", r"\bbrand\b", r"\bcard\b", r"\bvisa\b", r"\bmaster", r"\belo\b", r"\bamex\b"])
    return col_data, col_val, col_band

def mapping_minimo_adquirente(prefix: str, df_raw: pd.DataFrame):
    st.markdown("##### Mapear colunas mínimas (Adquirente)")
    cols = [str(c) for c in df_raw.columns]
    d_data, d_valor, d_band = _guess_defaults_cols(df_raw)

    c0, c1, c2 = st.columns(3)
    with c0:
        st.selectbox("Coluna de **Data**", ["— selecione —"] + cols,
                     index=(cols.index(d_data)+1) if d_data in cols else 0,
                     key=f"{prefix}_col_data")
    with c1:
        st.selectbox("Coluna de **Valor**", ["— selecione —"] + cols,
                     index=(cols.index(d_valor)+1) if d_valor in cols else 0,
                     key=f"{prefix}_col_valor")
    with c2:
        st.selectbox("Coluna de **Bandeira**", ["— selecione —"] + cols,
                     index=(cols.index(d_band)+1) if d_band in cols else 0,
                     key=f"{prefix}_col_bandeira")

# =============================================================================
# Matching por palavras-chave → Cód Conta Gerencial
# =============================================================================
def _match_bandeira_to_gerencial(band_value: str):
    """
    Procura cada token (normalizado) de MEIO_RULES dentro do texto da bandeira.
    Retorna (codigo_gerencial, meio_opcional)
    """
    if not band_value or not MEIO_RULES:
        return "", ""
    txt = _norm(band_value)
    for rule in MEIO_RULES:
        for tok in rule["tokens"]:
            if tok and tok in txt:
                return rule["codigo_gerencial"], rule.get("meio","")
    return "", ""

# =============================================================================
# Montagem do DataFrame final (Importador)
# =============================================================================
def _build_importador_df(df_raw: pd.DataFrame, prefix: str, grupo: str, loja: str, portador: str, tipo_imp: str):
    cd = st.session_state.get(f"{prefix}_col_data")
    cv = st.session_state.get(f"{prefix}_col_valor")
    cb = st.session_state.get(f"{prefix}_col_bandeira")
    if not cd or not cv or not cb or "— selecione —" in (cd, cv, cb):
        st.error("Defina **Data**, **Valor** e **Bandeira** para gerar o importador.")
        return pd.DataFrame()

    # Busca CNPJ / Códigos da empresa selecionada
    cnpj = ""; cod_ev = ""; cod_grp = ""
    if not df_emp.empty and loja:
        row = df_emp[(df_emp["Loja"].astype(str).str.strip()==loja) & (df_emp["Grupo"].astype(str).str.strip()==grupo)]
        if not row.empty:
            cnpj    = str(row.iloc[0].get("CNPJ","") or "")
            cod_ev  = str(row.iloc[0].get("Código Everest","") or "")
            cod_grp = str(row.iloc[0].get("Código Grupo Everest","") or "")

    # Normaliza valor (mantém as datas exatamente como o arquivo trouxe)
    col_valor = pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    col_data_original = df_raw[cd]  # mantém como veio (string)

    col_band = df_raw[cb].astype(str).str.strip()
    cod_conta = []
    meio_ref  = []
    for btxt in col_band:
        cod, meio = _match_bandeira_to_gerencial(btxt)
        cod_conta.append(cod)
        meio_ref.append(meio)

    # DataFrame no layout exato do Importador
    out = pd.DataFrame({
        "CNPJ Empresa":              cnpj,
        "Série Título":              "",
        "Nº Título":                 "",
        "Nº Parcela":                "",
        "Nº Documento":              "",
        "CNPJ/Cliente":              "",
        "Portador":                  portador or "",
        "Data Documento":            col_data_original,   # mantém string original
        "Data Vencimento":           "",                  # pode mapear depois
        "Data":                      col_data_original,   # mantém igual ao arquivo
        "Valor Desconto":            0.00,
        "Valor Multa":               0.00,
        "Valor Juros Dia":           0.00,
        "Valor Original":            col_valor,
        "Observações do Título":     col_band,            # guardo bandeira original
        "Cód Conta Gerencial":       cod_conta,           # vindo de Cod Gerencial Everest
        "Cód Centro de Custo":       "",                  # regra futura se desejar
    })

    # Metadados úteis no preview
    out["_Grupo"]                = grupo or ""
    out["_Empresa"]              = loja or ""
    out["_Código Everest"]       = cod_ev
    out["_Código Grupo Everest"] = cod_grp
    out["_Tipo Importação"]      = tipo_imp or ""
    out["_Meio Mapeado"]         = meio_ref

    # limpa linhas sem data ou sem valor
    out = out[(out["Data Documento"].astype(str).str.strip()!="") & (out["Valor Original"].notna())]

    # Garante ordem final + mantém metadados no fim
    col_order = [
        "CNPJ Empresa","Série Título","Nº Título","Nº Parcela","Nº Documento","CNPJ/Cliente","Portador",
        "Data Documento","Data Vencimento","Data",
        "Valor Desconto","Valor Multa","Valor Juros Dia","Valor Original",
        "Observações do Título","Cód Conta Gerencial","Cód Centro de Custo"
    ]
    out = out[col_order + [c for c in out.columns if c.startswith("_")]]
    return out

# =============================================================================
# Download do Excel (xlsxwriter)
# =============================================================================
def _download_importador_button(df: pd.DataFrame, prefix: str):
    if df.empty:
        return
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
        df.to_excel(wr, index=False, sheet_name="Importador")
        ws = wr.sheets["Importador"]
        money_fmt = wr.book.add_format({'num_format':'"R$" #,##0.00'})
        # auto larguras básicas e formatação de moeda
        for col_idx, col_name in enumerate(df.columns):
            maxw = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].head(300).fillna("").astype(str)])
            ws.set_column(col_idx, col_idx, min(maxw + 2, 60))
            if col_name == "Valor Original":
                ws.set_column(col_idx, col_idx, None, money_fmt)
    buf.seek(0)
    fname = f"Importador_{prefix.upper()}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button("📥 Baixar Excel — Importador", data=buf.getvalue(),
                       file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"dl_import_{prefix}")

# =============================================================================
# ABAS
# =============================================================================
aba_cr, aba_cp, aba_cad = st.tabs(["💰 Contas a Receber", "💸 Contas a Pagar", "🧾 Cadastro Cliente/Fornecedor"])

# --------- 💰 CONTAS A RECEBER ---------
with aba_cr:
    st.subheader("Contas a Receber")
    st.markdown('<div class="compact">', unsafe_allow_html=True)

    gsel, esel = filtros_grupo_empresa("cr", with_portador=True, with_tipo_imp=True)
    portador = st.session_state.get("cr_portador", "Todos")
    tipo_imp = st.session_state.get("cr_tipo_imp", "Todos")

    st.markdown('<hr class="compact">', unsafe_allow_html=True)
    df_raw = bloco_colagem("cr")

    if tipo_imp == "Adquirente" and not df_raw.empty:
        mapping_minimo_adquirente("cr", df_raw)

        cd = st.session_state.get("cr_col_data")
        cv = st.session_state.get("cr_col_valor")
        cb = st.session_state.get("cr_col_bandeira")
        if cd and cv and cb and "— selecione —" not in (cd, cv, cb):
            if st.button("⚙️ Gerar arquivo Importador (Receber)", type="primary"):
                df_imp = _build_importador_df(
                    df_raw, "cr",
                    gsel if gsel!="— selecione —" else "",
                    esel if esel!="— selecione —" else "",
                    portador, tipo_imp
                )
                if not df_imp.empty:
                    st.success(f"Gerado {len(df_imp)} linha(s).")
                    st.dataframe(df_imp.head(300), use_container_width=True, height=350)
                    _download_importador_button(df_imp, "cr")

    st.markdown('</div>', unsafe_allow_html=True)

# --------- 💸 CONTAS A PAGAR ---------
with aba_cp:
    st.subheader("Contas a Pagar")
    st.markdown('<div class="compact">', unsafe_allow_html=True)

    gsel, esel = filtros_grupo_empresa("cp", with_portador=True, with_tipo_imp=True)
    portador = st.session_state.get("cp_portador", "Todos")
    tipo_imp = st.session_state.get("cp_tipo_imp", "Todos")

    st.markdown('<hr class="compact">', unsafe_allow_html=True)
    df_raw = bloco_colagem("cp")

    if tipo_imp == "Adquirente" and not df_raw.empty:
        mapping_minimo_adquirente("cp", df_raw)
        cd = st.session_state.get("cp_col_data")
        cv = st.session_state.get("cp_col_valor")
        cb = st.session_state.get("cp_col_bandeira")
        if cd and cv and cb and "— selecione —" not in (cd, cv, cb):
            if st.button("⚙️ Gerar arquivo Importador (Pagar)", type="primary"):
                df_imp = _build_importador_df(
                    df_raw, "cp",
                    gsel if gsel!="— selecione —" else "",
                    esel if esel!="— selecione —" else "",
                    portador, tipo_imp
                )
                if not df_imp.empty:
                    st.success(f"Gerado {len(df_imp)} linha(s).")
                    st.dataframe(df_imp.head(300), use_container_width=True, height=350)
                    _download_importador_button(df_imp, "cp")

    st.markdown('</div>', unsafe_allow_html=True)

# --------- 🧾 CADASTRO Cliente/Fornecedor (simples, mantém layout) ---------
with aba_cad:
    st.subheader("Cadastro de Cliente / Fornecedor")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        gsel = st.selectbox("Grupo:", ["— selecione —"] + GRUPOS, key="cad_grupo")
    with col_g2:
        lojas = LOJAS_DO(gsel) if gsel != "— selecione —" else []
        esel = st.selectbox("Empresa:", ["— selecione —"] + lojas, key="cad_empresa")

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
        from gspread.exceptions import WorksheetNotFound
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
