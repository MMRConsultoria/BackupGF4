# pages/CR- CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re, json, unicodedata
from io import StringIO, BytesIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials
# --- fusível anti-help: evita que qualquer help() imprima no app ---
try:
    import builtins
    def _noop_help(*args, **kwargs):
        return None
    builtins.help = _noop_help
except Exception:
    pass

st.set_page_config(page_title="CR-CP Importador Everest", layout="wide")
st.set_option("client.showErrorDetails", False)
# 🔒 Bloqueio de acesso
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
# Helpers
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
    # normaliza e separa por palavras/nums
    return [w for w in re.findall(r"[0-9a-zA-Z]+", _norm_basic(txt)) if w]

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
            if _norm_basic(h) in names:
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

# ====== CARREGAMENTO DAS REGRAS (para o matching) ======
@st.cache_data(show_spinner=False)
def carregar_tabela_meio_pagto():
    """
    Lê as colunas necessárias:
      - 'Padrão Cod Gerencial'       (regras normais)
      - 'Cod Gerencial Everest'
      - 'CNPJ Bandeira'
      - 'PIX Padrão Cod Gerencial'   (regras de fallback para PIX)
    Retorna: DF_MEIO, MEIO_RULES, PIX_RULES
    """
    COL_PADRAO = "Padrão Cod Gerencial"
    COL_COD    = "Cod Gerencial Everest"
    COL_CNPJ   = "CNPJ Bandeira"
    COL_PIXPAD = "PIX Padrão Cod Gerencial"

    sh = _open_planilha("Vendas diarias")
    if not sh:
        return pd.DataFrame(), [], []

    try:
        ws = sh.worksheet("Tabela Meio Pagamento")
    except WorksheetNotFound:
        st.warning("⚠️ Aba 'Tabela Meio Pagamento' não encontrada.")
        return pd.DataFrame(), [], []

    df = pd.DataFrame(ws.get_all_records()).astype(str)

    # garante colunas
    for c in [COL_PADRAO, COL_COD, COL_CNPJ, COL_PIXPAD]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    # Regras normais (Padrão Cod Gerencial)
    meio_rules = []
    for _, row in df.iterrows():
        padrao = row[COL_PADRAO]
        codigo = row[COL_COD]
        cnpj   = row[COL_CNPJ]
        if not padrao or not codigo:
            continue
        tokens = sorted(set(_tokenize(padrao)))
        if not tokens:
            continue
        meio_rules.append({"tokens": tokens, "codigo_gerencial": codigo, "cnpj_bandeira": cnpj})

    # Regras PIX (fallback) — usa palavras da coluna "PIX Padrão Cod Gerencial"
    pix_rules = []
    for _, row in df.iterrows():
        pix_text = row[COL_PIXPAD]
        codigo   = row[COL_COD]
        cnpj     = row[COL_CNPJ]
        if not pix_text or not codigo:
            continue
        tokens = sorted(set(_tokenize(pix_text)))
        if not tokens:
            continue
        pix_rules.append({"tokens": tokens, "codigo_gerencial": codigo, "cnpj_bandeira": cnpj})

    return df, meio_rules, pix_rules


def _best_rule_for_tokens(ref_tokens: set):
    best = None
    best_hits = 0
    best_tokens_len = 0
    best_matched = set()
    for rule in MEIO_RULES:
        tokens = set(rule["tokens"])
        matched = tokens & ref_tokens
        hits = len(matched)
        if hits == 0:
            continue
        if (hits > best_hits) or (hits == best_hits and len(tokens) > best_tokens_len):
            best = rule
            best_hits = hits
            best_tokens_len = len(tokens)
            best_matched = matched
    return best, best_hits, best_matched

def _match_bandeira_to_gerencial(ref_text: str):
    if not ref_text or not MEIO_RULES:
        return "", "", ""
    ref_tokens = set(_tokenize(ref_text))
    if not ref_tokens:
        return "", "", ""
    best, _, _ = _best_rule_for_tokens(ref_tokens)
    if best:
        return best["codigo_gerencial"], best.get("cnpj_bandeira",""), ""
    return "", "", ""
def _best_rule_for_tokens_from(rules_list, ref_tokens: set):
    best = None
    best_hits = 0
    best_tokens_len = 0
    for rule in rules_list:
        tokens = set(rule["tokens"])
        hits = len(tokens & ref_tokens)
        if hits == 0:
            continue
        if (hits > best_hits) or (hits == best_hits and len(tokens) > best_tokens_len):
            best = rule
            best_hits = hits
            best_tokens_len = len(tokens)
    return best

def _apply_pix_fallback_on_errors(df_importador: pd.DataFrame) -> pd.DataFrame:
    """
    Para linhas com 'Cód Conta Gerencial' vazio:
      - procura palavras das PIX_RULES (coluna 'PIX Padrão Cod Gerencial')
      - se casar, preenche 'Cód Conta Gerencial' e 'CNPJ/Cliente' com
        'Cod Gerencial Everest' e 'CNPJ Bandeira' da mesma linha da tabela.
    Não mexe no que já tem código.
    """
    if df_importador.empty:
        return df_importador
    if 'PIX_RULES' not in globals() or not PIX_RULES:
        return df_importador

    df = df_importador.copy()
    col_ref  = "Observações do Título"
    col_cod  = "Cód Conta Gerencial"
    col_cnpj = "CNPJ/Cliente"

    if col_ref not in df.columns:
        return df
    if col_cod not in df.columns:
        df[col_cod] = ""
    if col_cnpj not in df.columns:
        df[col_cnpj] = ""

    mask_err = df[col_cod].astype(str).str.strip().eq("")
    for i in df.index[mask_err]:
        ref_text = str(df.at[i, col_ref] or "")
        ref_tokens = set(_tokenize(ref_text))
        if not ref_tokens:
            continue
        best = _best_rule_for_tokens_from(PIX_RULES, ref_tokens)
        if best:
            df.at[i, col_cod] = best["codigo_gerencial"]
            if str(df.at[i, col_cnpj]).strip() == "":
                df.at[i, col_cnpj] = best.get("cnpj_bandeira", "")
    return df


def _issues_summary(df: pd.DataFrame):
    miss_cnpj = int(df["CNPJ/Cliente"].astype(str).str.strip().eq("").sum()) if "CNPJ/Cliente" in df else 0
    miss_cod  = int(df["Cód Conta Gerencial"].astype(str).str.strip().eq("").sum()) if "Cód Conta Gerencial" in df else 0
    total     = len(df)
    return miss_cnpj, miss_cod, total

# ===== Dados base (carrega ANTES de montar a UI) =====
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
DF_MEIO, MEIO_RULES, PIX_RULES = carregar_tabela_meio_pagto()

# fallbacks na sessão (evita NameError em re-runs)
st.session_state["_grupos"] = GRUPOS
st.session_state["_lojas_map"] = LOJAS_MAP
st.session_state["_portadores"] = PORTADORES

def LOJAS_DO(grupo_nome: str):
    lojas_map = globals().get("LOJAS_MAP") or st.session_state.get("_lojas_map", {})
    return lojas_map.get(grupo_nome, [])

# ======= BOTÕES DISCRETOS (ESQ) + EDITORES: MEIO DE PAGAMENTO e PORTADOR =======

def _load_sheet_raw_full(sheet_name: str):
    """Lê a aba informada exatamente como está (todas as colunas/ordem)."""
    sh = _open_planilha("Vendas diarias")
    if not sh:
        raise RuntimeError("Planilha 'Vendas diarias' indisponível.")
    try:
        ws = sh.worksheet(sheet_name)
    except WorksheetNotFound:
        raise RuntimeError(f"Aba '{sheet_name}' não encontrada.")
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(), ws
    header = values[0]
    rows = values[1:]
    max_cols = len(header)
    norm_rows = [r + [""] * (max_cols - len(r)) for r in rows]
    df = pd.DataFrame(norm_rows, columns=header)
    return df, ws

def _save_sheet_full(df_edit: pd.DataFrame, ws):
    """Salva de volta o conteúdo exatamente como está no grid (inclui cabeçalhos)."""
    ws.clear()
    if df_edit.empty:
        return
    header = list(df_edit.columns)
    data = [header] + df_edit.astype(str).values.tolist()
    ws.update(data)

# --- barra discreta à esquerda com os dois botões ---
left, _ = st.columns([0.22, 0.78])
with left:
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("TB MeioPag", use_container_width=True, help="Abrir/editar aba Tabela Meio Pagamento"):
            st.session_state["editor_on_meio"] = True
    with c2:
        if st.button("TB Portador", use_container_width=True, help="Abrir/editar aba Portador"):
            st.session_state["editor_on_portador"] = True

# --- EDITOR: Tabela Meio Pagamento ---
if st.session_state.get("editor_on_meio"):
    st.markdown("Meio de Pagamento")
    try:
        df_rules_raw, ws_rules = _load_sheet_raw_full("Tabela Meio Pagamento")
    except Exception as e:
        st.error(f"Não foi possível abrir a tabela: {e}")
        st.session_state["editor_on_meio"] = False
    else:
        backup = BytesIO()
        with pd.ExcelWriter(backup, engine="openpyxl") as w:
            df_rules_raw.to_excel(w, index=False, sheet_name="Tabela Meio Pagamento")
        backup.seek(0)
        #st.download_button("Backup (.xlsx)", backup,
        #                   file_name="Tabela_Meio_Pagamento_backup.xlsx",
        #                   use_container_width=True)

        #st.info("Edite livremente; ao **Salvar e Fechar**, a aba será sobrescrita e as regras serão recarregadas.")
        edited = st.data_editor(
            df_rules_raw,
            num_rows="dynamic",
            use_container_width=True,
            height=520,
        )

        col_actions = st.columns([0.25, 0.25, 0.5])
        with col_actions[0]:
            if st.button("Salvar e Fechar", type="primary", use_container_width=True, key="meio_save"):
                try:
                    _save_sheet_full(edited, ws_rules)
                    # recarrega regras do app
                    st.cache_data.clear()
                    DF_MEIO, MEIO_RULES, PIX_RULES = carregar_tabela_meio_pagto()
                    st.session_state["editor_on_meio"] = False
                    st.success("Alterações salvas, regras atualizadas e editor fechado.")
                except Exception as e:
                    st.error(f"Falha ao salvar: {e}")
        #with col_actions[1]:
        #    if st.button("Fechar sem salvar", use_container_width=True, key="meio_close"):
        #        st.session_state["editor_on_meio"] = False

# --- EDITOR: Portador ---
if st.session_state.get("editor_on_portador"):
    st.markdown("Portador")
    try:
        df_port_raw, ws_port = _load_sheet_raw_full("Portador")
    except Exception as e:
        st.error(f"Não foi possível abrir a aba Portador: {e}")
        st.session_state["editor_on_portador"] = False
    else:
        backup2 = BytesIO()
        with pd.ExcelWriter(backup2, engine="openpyxl") as w:
            df_port_raw.to_excel(w, index=False, sheet_name="Portador")
        backup2.seek(0)
        #st.download_button("Backup Portador (.xlsx)", backup2,
        #                   file_name="Portador_backup.xlsx",
        #                   use_container_width=True)

       #st.info("Edite livremente; ao **Salvar e Fechar**, a aba será sobrescrita e o mapa de portadores será recarregado.")
        edited_port = st.data_editor(
            df_port_raw,
            num_rows="dynamic",
            use_container_width=True,
            height=520,
        )

        col_actions2 = st.columns([0.25, 0.25, 0.5])
        with col_actions2[0]:
            if st.button("Salvar e Fechar", type="primary", use_container_width=True, key="port_save"):
                try:
                    _save_sheet_full(edited_port, ws_port)
                    # recarrega portadores do app
                    st.cache_data.clear()
                    PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
                    # atualiza fallbacks em sessão
                    st.session_state["_portadores"] = PORTADORES
                    st.session_state["editor_on_portador"] = False
                    st.success("Alterações salvas, portadores atualizados e editor fechado.")
                except Exception as e:
                    st.error(f"Falha ao salvar: {e}")
        #with col_actions2[1]:
        #    if st.button("Fechar sem salvar", use_container_width=True, key="port_close"):
        #        st.session_state["editor_on_portador"] = False



# ===== Ordem de saída (sem a flag; a flag entra na frente) =====
IMPORTADOR_ORDER = [
    "CNPJ Empresa",
    "Série Título",
    "Nº Título",
    "Nº Parcela",
    "Nº Documento",
    "CNPJ/Cliente",
    "Portador",
    "Data Documento",
    "Data Vencimento",
    "Data",
    "Valor Desconto",
    "Valor Multa",
    "Valor Juros Dia",
    "Valor Original",
    "Observações do Título",
    "Cód Conta Gerencial",
    "Cód Centro de Custo",
]

# ======================
# UI Components
# ======================
def filtros_grupo_empresa(prefix, with_portador=False, with_tipo_imp=False):
    """Grupo | Empresa | Banco | Tipo de Importação (lado a lado) com fallback seguro."""
    c1, c2, c3, c4 = st.columns([1,1,1,1])

    grupos = globals().get("GRUPOS") or st.session_state.get("_grupos", [])
    try:
        grupos = list(grupos)
    except Exception:
        grupos = []

    with c1:
        gsel = st.selectbox("Grupo:", ["— selecione —"] + grupos, key=f"{prefix}_grupo")

    with c2:
        lojas = LOJAS_DO(gsel) if gsel and gsel != "— selecione —" else []
        esel = st.selectbox("Empresa:", ["— selecione —"] + lojas, key=f"{prefix}_empresa")

    with c3:
        if with_portador:
            portadores = globals().get("PORTADORES") or st.session_state.get("_portadores", [])
            st.selectbox("Banco:", ["Todos"] + list(portadores), index=0, key=f"{prefix}_portador")
        else:
            st.empty()

    with c4:
        if with_tipo_imp:
            st.selectbox("Tipo de Importação:", ["Todos","Adquirente","Cliente","Outros"], index=0, key=f"{prefix}_tipo_imp")
        else:
            st.empty()

    return gsel, esel

# limpa DF gerado quando o usuário apaga a colagem
def _on_paste_change(prefix: str):
    txt = st.session_state.get(f"{prefix}_paste", "")
    if not str(txt).strip():
        st.session_state.pop(f"{prefix}_df_imp", None)
        st.session_state.pop(f"{prefix}_edited_once", None)

def bloco_colagem(prefix: str):
    """Apenas colagem + pré-visualização opcional."""
    c1,c2 = st.columns([0.65,0.35])
    with c1:
        txt = st.text_area(
            "📋 Colar tabela (Ctrl+V)",
            height=180,
            placeholder="Cole aqui os dados copiados do Excel/Sheets… (ex.: a coluna 'Complemento')",
            key=f"{prefix}_paste",
            on_change=_on_paste_change,
            args=(prefix,)
        )
        df_paste = _try_parse_paste(txt) if (txt and str(txt).strip()) else pd.DataFrame()

    with c2:
        show_prev = st.checkbox("Mostrar pré-visualização da colagem", value=False, key=f"{prefix}_show_prev")
        if show_prev and not df_paste.empty:
            st.dataframe(df_paste, use_container_width=True, height=120)
        elif df_paste.empty:
            st.info("Cole dados para prosseguir.")

    return df_paste

def _column_mapping_ui(prefix: str, df_raw: pd.DataFrame):
    st.markdown("##### Mapear colunas para **Adquirente**")
    cols = ["— selecione —"] + list(df_raw.columns)
    c1,c2,c3 = st.columns(3)
    with c1:
        st.selectbox("Coluna de **Data**", cols, key=f"{prefix}_col_data")
    with c2:
        st.selectbox("Coluna de **Valor**", cols, key=f"{prefix}_col_valor")
    with c3:
        st.selectbox("Coluna de **Referência (texto do extrato)**", cols, key=f"{prefix}_col_bandeira")

def _build_importador_df(df_raw: pd.DataFrame, prefix: str, grupo: str, loja: str,
                         banco_escolhido: str):
    cd = st.session_state.get(f"{prefix}_col_data")
    cv = st.session_state.get(f"{prefix}_col_valor")
    cb = st.session_state.get(f"{prefix}_col_bandeira")

    if not cd or not cv or not cb or "— selecione —" in (cd, cv, cb):
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
    banco_escolhido = banco_escolhido or ""
    portador_nome = MAPA_BANCO_PARA_PORTADOR.get(banco_escolhido, banco_escolhido)

    # dados do usuário (mantém a data exatamente como veio)
    data_original  = df_raw[cd].astype(str)
    valor_original = pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    ref_txt        = df_raw[cb].astype(str).str.strip()

    # mapeamento por tokens do Padrão Cod Gerencial
    cod_conta_list, cnpj_cli_list = [], []
    for b in ref_txt:
        cod, cnpj_band, _ = _match_bandeira_to_gerencial(b)
        cod_conta_list.append(cod)
        cnpj_cli_list.append(cnpj_band)

    out = pd.DataFrame({
        "CNPJ Empresa":          cnpj_loja,
        "Série Título":          "DRE",
        "Nº Título":             "",
        "Nº Parcela":            1,
        "Nº Documento":          "DRE",
        "CNPJ/Cliente":          cnpj_cli_list,
        "Portador":              portador_nome,
        "Data Documento":        data_original,
        "Data Vencimento":       data_original,
        "Data":                  data_original,
        "Valor Desconto":        0.00,
        "Valor Multa":           0.00,
        "Valor Juros Dia":       0.00,
        "Valor Original":        valor_original,
        "Observações do Título": ref_txt.tolist(),
        "Cód Conta Gerencial":   cod_conta_list,
        "Cód Centro de Custo":   3
    })

    # filtra linhas válidas
    out = out[(out["Data"].astype(str).str.strip() != "") & (out["Valor Original"].notna())]

    # reordena conforme importador e coloca flag no início
    out = out.reindex(columns=[c for c in IMPORTADOR_ORDER if c in out.columns])
    out.insert(0, "🔴 Falta CNPJ?", out["CNPJ/Cliente"].astype(str).str.strip().eq(""))

    final_cols = ["🔴 Falta CNPJ?"] + [c for c in IMPORTADOR_ORDER if c in out.columns]
    out = out[final_cols]
    return out

def _download_excel(df: pd.DataFrame, filename: str, label_btn: str, disabled=False):
    if df.empty:
        st.button(label_btn, disabled=True, use_container_width=True)
        return
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Importador")
    bio.seek(0)
    st.download_button(label_btn, data=bio,
                       file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True,
                       disabled=disabled)

# ======================
# ABAS
# ======================
aba_cr, aba_cp, aba_cad = st.tabs(["💰 Contas a Receber", "💸 Contas a Pagar", "🧾 Cadastro Cliente/Fornecedor"])

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

    cr_ready = (
        st.session_state.get("cr_tipo_imp") == "Adquirente"
        and not df_raw.empty
        and all(st.session_state.get(k) and st.session_state.get(k) != "— selecione —"
                for k in ["cr_col_data","cr_col_valor","cr_col_bandeira"])
        and gsel not in (None, "", "— selecione —")
        and esel not in (None, "", "— selecione —")
    )

    if cr_ready:
        df_imp = _build_importador_df(
            df_raw, "cr",
            gsel, esel,
            st.session_state.get("cr_portador","")
        )
        df_imp = _apply_pix_fallback_on_errors(df_imp)
        # 👉 Recalcula a flag SÓ AGORA (depois de PIX e demais regras)
        df_imp["🔴 Falta CNPJ?"] = df_imp["CNPJ/Cliente"].astype(str).str.strip().eq("")
        df_imp = df_imp[ ["🔴 Falta CNPJ?"] + [c for c in df_imp.columns if c != "🔴 Falta CNPJ?"] ]
        
        # 👉 Alerta de pendências (não bloqueia o download)
        m_cnpj, m_cod, tot = _issues_summary(df_imp)
        if m_cnpj or m_cod:
            st.warning(f"⚠️ Atenção: {m_cnpj} linha(s) sem CNPJ e {m_cod} sem Cód Conta Gerencial. "
                       "Revise antes de gerar o Excel (você ainda pode baixar).")

        st.session_state["cr_edited_once"] = False
        st.session_state["cr_df_imp"] = df_imp.copy()

    df_imp_state = st.session_state.get("cr_df_imp")
    if isinstance(df_imp_state, pd.DataFrame) and not df_imp_state.empty:
        df_imp = df_imp_state
        show_only_missing = st.checkbox("Mostrar apenas linhas com 🔴 Falta CNPJ", value=st.session_state.get("cr_only_missing", False), key="cr_only_missing")
        df_view = df_imp[df_imp["🔴 Falta CNPJ?"]] if show_only_missing else df_imp

        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]

        editor_key = f"cr_editor_{gsel}_{esel}_{st.session_state.get('cr_col_data')}_{st.session_state.get('cr_col_valor')}_{st.session_state.get('cr_col_bandeira')}"
        edited_cr = st.data_editor(df_view, disabled=disabled_cols, use_container_width=True, height=420, key=editor_key)

        if not edited_cr.equals(df_view):
            st.session_state["cr_edited_once"] = True

        edited_full = df_imp.copy()
        edited_full.update(edited_cr)
        edited_full["🔴 Falta CNPJ?"] = edited_full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        cols_final = ["🔴 Falta CNPJ?"] + [c for c in edited_full.columns if c != "🔴 Falta CNPJ?"]
        edited_full = edited_full.reindex(columns=cols_final)
        # 👉 Alerta de pendências com base no que está na tela
        m_cnpj, m_cod, tot = _issues_summary(edited_full)
        if m_cnpj or m_cod:
            st.warning(f"⚠️ Existem {m_cnpj} linha(s) sem CNPJ e {m_cod} sem Cód Conta Gerencial. "
                       "Você pode baixar mesmo assim, mas recomenda-se revisar.")
        st.session_state["cr_df_imp"] = edited_full

        #faltam = int(edited_full["🔴 Falta CNPJ?"].sum())
        #total  = int(len(edited_full))
        
        def _download_excel(df: pd.DataFrame, filename: str, label_btn: str, disabled=False):
            if df.empty:
                st.button(label_btn, disabled=True, use_container_width=True)
                return
        
            df_export = df.copy()
        
            # 1) Remove a flag da exportação
            if "🔴 Falta CNPJ?" in df_export.columns:
                df_export = df_export.drop(columns=["🔴 Falta CNPJ?"], errors="ignore")
        
            # 2) Regras de tipos
            DEC_COLS = {"Valor Desconto", "Valor Multa", "Valor Juros Dia", "Valor Original"}
            INT_PREF = {"Portador", "Cód Conta Gerencial", "Cód Centro de Custo", "Nº Parcela"}
        
            for col in df_export.columns:
                # ===== CNPJ/Cliente =====
                if col == "CNPJ/Cliente":
                    s_raw = df_export[col].astype(str).str.strip()
                    s_digits = s_raw.str.replace(r"\D", "", regex=True)
                    mask_cnpj = s_digits.str.len() == 14
                    mask_only_digits = s_raw.str.match(r"^\d+$")
        
                    # CNPJ (14 dígitos) => manter TEXTO exatamente como veio
                    if mask_cnpj.any():
                        df_export.loc[mask_cnpj, col] = s_raw[mask_cnpj]
        
                    # Não CNPJ mas só dígitos => NÚMERO
                    to_num_mask = (~mask_cnpj) & mask_only_digits
                    if to_num_mask.any():
                        df_export.loc[to_num_mask, col] = pd.to_numeric(
                            s_raw[to_num_mask], errors="coerce", downcast="integer"
                        )
        
                    # Demais casos (tem letra/símbolo e não é CNPJ): mantém texto como está
                    continue
        
                # ===== Decimais (R$ etc.) =====
                if col in DEC_COLS:
                    s = df_export[col].astype(str).str.strip()
                    s_norm = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                    mask_num = s_norm.str.match(r"^\d+(\.\d+)?$")
                    if mask_num.any():
                        df_export.loc[mask_num, col] = pd.to_numeric(s_norm[mask_num], errors="coerce")
                    continue
        
                # ===== Inteiros preferenciais =====
                if col in INT_PREF:
                    s = df_export[col].astype(str).str.strip()
                    mask_int = s.str.match(r"^\d+$")
                    if mask_int.any():
                        df_export.loc[mask_int, col] = pd.to_numeric(
                            s[mask_int], errors="coerce", downcast="integer"
                        )
                    # se não for só dígitos, mantém texto (não força)
                    continue
        
                # Demais colunas: não forçar tipo
        
            # 3) Gerar Excel
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Importador")
            bio.seek(0)
        
            st.download_button(
                label_btn,
                data=bio,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=disabled
            )

        _download_excel(edited_full, "Importador_Receber.xlsx", "📥 Baixar Importador (Receber)", disabled=False)

    else:
        if st.session_state.get("cr_tipo_imp") == "Adquirente" and not df_raw.empty:
            st.info("Mapeie as colunas (Data, Valor, Referência) e selecione Grupo/Empresa para gerar.")

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

    cp_ready = (
        st.session_state.get("cp_tipo_imp") == "Adquirente"
        and not df_raw.empty
        and all(st.session_state.get(k) and st.session_state.get(k) != "— selecione —"
                for k in ["cp_col_data","cp_col_valor","cp_col_bandeira"])
        and gsel not in (None, "", "— selecione —")
        and esel not in (None, "", "— selecione —")
    )

    if cp_ready:
        df_imp = _build_importador_df(
            df_raw, "cp",
            gsel, esel,
            st.session_state.get("cp_portador","")
        )
        df_imp = _apply_pix_fallback_on_errors(df_imp)
        st.session_state["cp_edited_once"] = False
        st.session_state["cp_df_imp"] = df_imp.copy()

    df_imp_state = st.session_state.get("cp_df_imp")
    if isinstance(df_imp_state, pd.DataFrame) and not df_imp_state.empty:
        df_imp = df_imp_state

        show_only_missing = st.checkbox("Mostrar apenas linhas com 🔴 Falta CNPJ", value=st.session_state.get("cp_only_missing", False), key="cp_only_missing")
        df_view = df_imp[df_imp["🔴 Falta CNPJ?"]] if show_only_missing else df_imp

        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]

        editor_key = f"cp_editor_{gsel}_{esel}_{st.session_state.get('cp_col_data')}_{st.session_state.get('cp_col_valor')}_{st.session_state.get('cp_col_bandeira')}"
        edited_cp = st.data_editor(df_view, disabled=disabled_cols, use_container_width=True, height=420, key=editor_key)

        if not edited_cp.equals(df_view):
            st.session_state["cp_edited_once"] = True

        edited_full = df_imp.copy()
        edited_full.update(edited_cp)
        edited_full["🔴 Falta CNPJ?"] = edited_full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        cols_final = ["🔴 Falta CNPJ?"] + [c for c in edited_full.columns if c != "🔴 Falta CNPJ?"]
        edited_full = edited_full.reindex(columns=cols_final)

        st.session_state["cp_df_imp"] = edited_full

        faltam = int(edited_full["🔴 Falta CNPJ?"].sum())
        total  = int(len(edited_full))
        #st.warning(f"⚠️ {faltam} de {total} linha(s) sem CNPJ/Cliente.") if faltam else st.success("✅ Todos os CNPJs foram preenchidos.")

        #_download_excel(edited_full, "Importador_Pagar.xlsx", "📥 Baixar Importador (Pagar)", disabled=not st.session_state.get("cp_edited_once", False))
        def _download_excel(df: pd.DataFrame, filename: str, label_btn: str, disabled=False):
            if df.empty:
                st.button(label_btn, disabled=True, use_container_width=True)
                return
        
            # --- faz cópia para não alterar o DataFrame original
            df_export = df.copy()
        
            # 1️⃣ remove a coluna de flag "Falta CNPJ" da exportação
            if "🔴 Falta CNPJ?" in df_export.columns:
                df_export = df_export.drop(columns=["🔴 Falta CNPJ?"], errors="ignore")
        
            # 2️⃣ tenta converter colunas numéricas que vieram como texto
            for col in ["CNPJ/Cliente", "Portador", "Cód Conta Gerencial"]:
                if col in df_export.columns:
                    df_export[col] = (
                        pd.to_numeric(df_export[col].astype(str).str.replace(r"[^0-9]", "", regex=True), errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )
        
            # 3️⃣ gera o Excel
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Importador")
            bio.seek(0)
        
            # 4️⃣ botão de download
            st.download_button(
                label_btn,
                data=bio,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=disabled
            )

    else:
        if st.session_state.get("cp_tipo_imp") == "Adquirente" and not df_raw.empty:
            st.info("Mapeie as colunas (Data, Valor, Referência) e selecione Grupo/Empresa para gerar.")

# --------- 🧾 CADASTRO Cliente/Fornecedor ---------
with aba_cad:
    st.subheader("Cadastro de Cliente / Fornecedor")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        gsel = st.selectbox("Grupo:", ["— selecione —"]+ (globals().get("GRUPOS") or st.session_state.get("_grupos", [])), key="cad_grupo")
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
