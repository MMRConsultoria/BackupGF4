# pages/CR- CP Importador Everest.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re, json, unicodedata
from io import StringIO, BytesIO
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

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
    Lê colunas necessárias:
      - 'Padrão Cod Gerencial'
      - 'Cod Gerencial Everest'
      - 'CNPJ Bandeira'
      - 'PIX Padrão Cod Gerencial'   (NOVO)
    """
    COL_PADRAO = "Padrão Cod Gerencial"
    COL_COD    = "Cod Gerencial Everest"
    COL_CNPJ   = "CNPJ Bandeira"
    COL_PIXPAD = "PIX Padrão Cod Gerencial"

    sh = _open_planilha("Vendas diarias")
    if not sh:
        return pd.DataFrame(), [], None

    try:
        ws = sh.worksheet("Tabela Meio Pagamento")
    except WorksheetNotFound:
        st.warning("⚠️ Aba 'Tabela Meio Pagamento' não encontrada.")
        return pd.DataFrame(), [], None

    df = pd.DataFrame(ws.get_all_records()).astype(str)

    # Garantir colunas
    for c in [COL_PADRAO, COL_COD, COL_CNPJ, COL_PIXPAD]:
        if c not in df.columns:
            df[c] = ""

    # Normaliza
    for c in [COL_PADRAO, COL_COD, COL_CNPJ, COL_PIXPAD]:
        df[c] = df[c].astype(str).str.strip()

    # Regras de token para bandeira/gerencial
    rules = []
    for _, row in df.iterrows():
        padrao = row[COL_PADRAO]
        codigo = row[COL_COD]
        cnpj   = row[COL_CNPJ]
        if not padrao or not codigo:
            continue
        tokens = sorted(set(_tokenize(padrao)))
        if not tokens:
            continue
        rules.append({"tokens": tokens, "codigo_gerencial": codigo, "cnpj_bandeira": cnpj})

    # PIX padrão (se houver múltiplos diferentes, ficamos com o primeiro e logamos em tela)
    pix_vals = [v for v in df[COL_PIXPAD].tolist() if str(v).strip()]
    pix_default = pix_vals[0].strip() if pix_vals else ""
    if pix_vals and len(set(pix_vals)) > 1:
        st.warning("⚠️ Existem múltiplos valores diferentes em 'PIX Padrão Cod Gerencial'. Usando o primeiro encontrado.")

    pix_default = pix_default if pix_default else None
    return df, rules, pix_default

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

# ===== Dados base (carrega ANTES de montar a UI) =====
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
DF_MEIO, MEIO_RULES, PIX_DEFAULT_CODE = carregar_tabela_meio_pagto()

# fallbacks na sessão
st.session_state["_grupos"] = GRUPOS
st.session_state["_lojas_map"] = LOJAS_MAP
st.session_state["_portadores"] = PORTADORES

def LOJAS_DO(grupo_nome: str):
    lojas_map = globals().get("LOJAS_MAP") or st.session_state.get("_lojas_map", {})
    return lojas_map.get(grupo_nome, [])

# ======================
# LOGS DE CLASSIFICAÇÃO PIX (sessão)
# ======================
def _init_logs():
    st.session_state.setdefault("pix_log_ok", [])
    st.session_state.setdefault("pix_log_err", [])

def _add_log_ok(**kwargs):
    _init_logs()
    st.session_state["pix_log_ok"].append(kwargs)

def _add_log_err(**kwargs):
    _init_logs()
    st.session_state["pix_log_err"].append(kwargs)

def _logs_to_df():
    ok = pd.DataFrame(st.session_state.get("pix_log_ok", []))
    err = pd.DataFrame(st.session_state.get("pix_log_err", []))
    # Ordena colunas amigáveis
    cols = ["Quando","Módulo","Grupo","Empresa","Banco/Portador","Data","Valor","Referência",
            "CNPJ Empresa","CNPJ/Cliente","Cod Antes","Cod Depois","Motivo/Regra"]
    ok = ok.reindex(columns=cols, fill_value="")
    err = err.reindex(columns=cols, fill_value="")
    return ok, err

def _now():
    # timestamp amigável
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _save_logs_to_sheet():
    """Grava (append) os logs na aba 'Log Classificação PIX'."""
    ok_df, err_df = _logs_to_df()
    if ok_df.empty and err_df.empty:
        st.info("Não há logs a gravar.")
        return

    sh = _open_planilha("Vendas diarias")
    if not sh:
        st.error("Planilha 'Vendas diarias' indisponível.")
        return
    ws = None
    try:
        ws = sh.worksheet("Log Classificação PIX")
    except WorksheetNotFound:
        ws = sh.add_worksheet("Log Classificação PIX", rows=1000, cols=30)
        ws.append_row(["Quando","Módulo","Grupo","Empresa","Banco/Portador","Data","Valor","Referência",
                       "CNPJ Empresa","CNPJ/Cliente","Cod Antes","Cod Depois","Motivo/Regra","Tipo Log"])

    def _append(df, tipo):
        if df.empty: return
        values = df.values.tolist()
        for r in values:
            ws.append_row(r + [tipo])

    _append(ok_df, "OK")
    _append(err_df, "ERRO")
    st.success("Logs gravados na aba 'Log Classificação PIX'.")

# ======================
# Detector de PIX e Classificador
# ======================
PIX_PATTERNS = [
    r"\bpix\b",                         # palavra pix
    r"\bqr\b",                          # muitas adquirentes colocam 'qr'
    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",  # GUID comum em QR
    r"[A-Z0-9]{25,}",                   # chaves/txid longos
]

def _is_pix_reference(ref_text: str) -> bool:
    if not ref_text: return False
    t = _norm_basic(ref_text)
    for pat in PIX_PATTERNS:
        if re.search(pat, t):
            return True
    # pistas adicionais
    if any(w in t for w in ["chave","txid","qr code","qrcode","pagamento instantaneo","instantaneo"]):
        return True
    return False

def _classificar_pix_em_df(df_importador: pd.DataFrame, modulo: str, grupo: str, empresa: str, banco_nome: str):
    """
    - NÃO sobrescreve 'Cód Conta Gerencial' (só preenche se vier vazio).
    - Usa PIX_DEFAULT_CODE (da tabela) para linhas cuja referência é PIX.
    - Gera logs (ok/erro).
    """
    if df_importador.empty:
        return df_importador

    _init_logs()

    if not PIX_DEFAULT_CODE:
        st.info("ℹ️ Nenhum 'PIX Padrão Cod Gerencial' definido na Tabela Meio Pagamento. Sem preenchimento automático de PIX.")
        # Mesmo sem código padrão, ainda registramos erros quando detectar PIX sem código
    # Trabalha em cópia para segurança
    df = df_importador.copy()

    # Colunas base seguras
    col_ref = "Observações do Título"
    col_cod = "Cód Conta Gerencial"

    if col_ref not in df.columns:
        return df
    if col_cod not in df.columns:
        df[col_cod] = ""

    # Loop para log granular
    for idx, row in df.iterrows():
        ref = str(row.get(col_ref, "") or "")
        cod_antes = str(row.get(col_cod, "") or "").strip()
        is_pix = _is_pix_reference(ref)

        if not is_pix:
            # não é pix -> nada a fazer
            continue

        if cod_antes:
            # já classificado (regra antiga/bandeira). Não mexer.
            _add_log_ok(
                Quando=_now(), Módulo=modulo, Grupo=grupo, Empresa=empresa, Banco/Portador=banco_nome,
                Data=str(row.get("Data","")), Valor=row.get("Valor Original",""),
                Referência=ref, **{"CNPJ Empresa":str(row.get("CNPJ Empresa",""))},
                **{"CNPJ/Cliente":str(row.get("CNPJ/Cliente",""))},
                **{"Cod Antes":cod_antes}, **{"Cod Depois":cod_antes},
                **{"Motivo/Regra":"PIX detectado, mas já havia classificação anterior — mantido"}
            )
            continue

        if not PIX_DEFAULT_CODE:
            # é pix mas não temos padrão -> erro
            _add_log_err(
                Quando=_now(), Módulo=modulo, Grupo=grupo, Empresa=empresa, Banco/Portador=banco_nome,
                Data=str(row.get("Data","")), Valor=row.get("Valor Original",""),
                Referência=ref, **{"CNPJ Empresa":str(row.get("CNPJ Empresa",""))},
                **{"CNPJ/Cliente":str(row.get("CNPJ/Cliente",""))},
                **{"Cod Antes":cod_antes}, **{"Cod Depois":""},
                **{"Motivo/Regra":"PIX detectado, porém 'PIX Padrão Cod Gerencial' está vazio"}
            )
            continue

        # aplicar padrão
        df.at[idx, col_cod] = PIX_DEFAULT_CODE
        _add_log_ok(
            Quando=_now(), Módulo=modulo, Grupo=grupo, Empresa=empresa, Banco/Portador=banco_nome,
            Data=str(row.get("Data","")), Valor=row.get("Valor Original",""),
            Referência=ref, **{"CNPJ Empresa":str(row.get("CNPJ Empresa",""))},
            **{"CNPJ/Cliente":str(row.get("CNPJ/Cliente",""))},
            **{"Cod Antes":cod_antes}, **{"Cod Depois":PIX_DEFAULT_CODE},
            **{"Motivo/Regra":"Classificado via 'PIX Padrão Cod Gerencial'"}
        )

    # Atualiza flag de Falta CNPJ
    if "CNPJ/Cliente" in df.columns:
        df["🔴 Falta CNPJ?"] = df["CNPJ/Cliente"].astype(str).str.strip().eq("")
        # reordenar com a flag primeiro (mantém seu padrão)
        flag = df.pop("🔴 Falta CNPJ?")
        df.insert(0, "🔴 Falta CNPJ?", flag)

    return df

# ======= BOTÕES DISCRETOS (ESQ) + EDITORES: MEIO DE PAGAMENTO e PORTADOR =======

def _load_sheet_raw_full(sheet_name: str):
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
    ws.clear()
    if df_edit.empty:
        return
    header = list(df_edit.columns)
    data = [header] + df_edit.astype(str).values.tolist()
    ws.update(data)

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
        st.download_button("Backup (.xlsx)", backup,
                           file_name="Tabela_Meio_Pagamento_backup.xlsx",
                           use_container_width=True)

        st.info("Edite livremente; ao **Salvar e Fechar**, a aba será sobrescrita e as regras serão recarregadas.")
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
                    st.cache_data.clear()
                    # recarrega globais
                    global DF_MEIO, MEIO_RULES, PIX_DEFAULT_CODE
                    DF_MEIO, MEIO_RULES, PIX_DEFAULT_CODE = carregar_tabela_meio_pagto()
                    st.session_state["editor_on_meio"] = False
                    st.success("Alterações salvas, regras/PIX atualizados e editor fechado.")
                except Exception as e:
                    st.error(f"Falha ao salvar: {e}")
        with col_actions[1]:
            if st.button("Fechar sem salvar", use_container_width=True, key="meio_close"):
                st.session_state["editor_on_meio"] = False

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
        st.download_button("Backup Portador (.xlsx)", backup2,
                           file_name="Portador_backup.xlsx",
                           use_container_width=True)

        st.info("Edite livremente; ao **Salvar e Fechar**, a aba será sobrescrita e o mapa de portadores será recarregado.")
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
                    st.cache_data.clear()
                    global PORTADORES, MAPA_BANCO_PARA_PORTADOR
                    PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
                    st.session_state["_portadores"] = PORTADORES
                    st.session_state["editor_on_portador"] = False
                    st.success("Alterações salvas, portadores atualizados e editor fechado.")
                except Exception as e:
                    st.error(f"Falha ao salvar: {e}")
        with col_actions2[1]:
            if st.button("Fechar sem salvar", use_container_width=True, key="port_close"):
                st.session_state["editor_on_portador"] = False

# ===== Ordem de saída =====
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

def _on_paste_change(prefix: str):
    txt = st.session_state.get(f"{prefix}_paste", "")
    if not str(txt).strip():
        st.session_state.pop(f"{prefix}_df_imp", None)
        st.session_state.pop(f"{prefix}_edited_once", None)

def bloco_colagem(prefix: str):
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

    banco_escolhido = banco_escolhido or ""
    portador_nome = MAPA_BANCO_PARA_PORTADOR.get(banco_escolhido, banco_escolhido)

    data_original  = df_raw[cd].astype(str)
    valor_original = pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    ref_txt        = df_raw[cb].astype(str).str.strip()

    # mapeamento por tokens
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

    out = out[(out["Data"].astype(str).str.strip() != "") & (out["Valor Original"].notna())]
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

    # ⚙️ Opções de PIX/LOG para CR
    with st.expander("⚙️ Opções de classificação PIX (CR)"):
        do_pix_cr = st.checkbox("Aplicar classificação automática de PIX (CR)", value=True, key="do_pix_cr")
        colx1, colx2, colx3 = st.columns([0.38,0.38,0.24])
        with colx1:
            if st.button("📥 Baixar Log (OK/Erro) – CR", use_container_width=True):
                ok_df, err_df = _logs_to_df()
                with pd.ExcelWriter(BytesIO(), engine="openpyxl") as w:
                    pass
            # Baixar dois arquivos (OK e ERRO)
            ok_df, err_df = _logs_to_df()
            if not ok_df.empty:
                bio_ok = BytesIO()
                with pd.ExcelWriter(bio_ok, engine="openpyxl") as w:
                    ok_df.to_excel(w, index=False, sheet_name="Log_OK")
                bio_ok.seek(0)
                st.download_button("⬇️ Baixar Log OK (CR/CP)", bio_ok, file_name="Log_PIX_OK.xlsx", use_container_width=True)
            if not err_df.empty:
                bio_er = BytesIO()
                with pd.ExcelWriter(bio_er, engine="openpyxl") as w:
                    err_df.to_excel(w, index=False, sheet_name="Log_ERRO")
                bio_er.seek(0)
                st.download_button("⬇️ Baixar Log ERRO (CR/CP)", bio_er, file_name="Log_PIX_ERRO.xlsx", use_container_width=True)
        with colx2:
            if st.button("📝 Gravar logs no Google Sheets", use_container_width=True):
                _save_logs_to_sheet()
        with colx3:
            if st.button("🧹 Limpar logs (sessão)", use_container_width=True):
                st.session_state["pix_log_ok"] = []
                st.session_state["pix_log_err"] = []
                st.success("Logs da sessão limpos.")

    if cr_ready:
        df_imp = _build_importador_df(
            df_raw, "cr",
            gsel, esel,
            st.session_state.get("cr_portador","")
        )
        # aplica PIX (sem mexer no que já tem código)
        if do_pix_cr:
            df_imp = _classificar_pix_em_df(
                df_imp, modulo="CR", grupo=gsel, empresa=esel,
                banco_nome=st.session_state.get("cr_portador","")
            )

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

        st.session_state["cr_df_imp"] = edited_full

        _download_excel(edited_full, "Importador_Receber.xlsx", "📥 Baixar Importador (Receber)", disabled=not st.session_state.get("cr_edited_once", False))
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

    # ⚙️ Opções de PIX/LOG para CP
    with st.expander("⚙️ Opções de classificação PIX (CP)"):
        do_pix_cp = st.checkbox("Aplicar classificação automática de PIX (CP)", value=True, key="do_pix_cp")
        coly1, coly2, coly3 = st.columns([0.5,0.25,0.25])
        with coly1:
            ok_df, err_df = _logs_to_df()
            if not ok_df.empty:
                bio_ok2 = BytesIO()
                with pd.ExcelWriter(bio_ok2, engine="openpyxl") as w:
                    ok_df.to_excel(w, index=False, sheet_name="Log_OK")
                bio_ok2.seek(0)
                st.download_button("⬇️ Baixar Log OK (CR/CP)", bio_ok2, file_name="Log_PIX_OK.xlsx", use_container_width=True)
            if not err_df.empty:
                bio_er2 = BytesIO()
                with pd.ExcelWriter(bio_er2, engine="openpyxl") as w:
                    err_df.to_excel(w, index=False, sheet_name="Log_ERRO")
                bio_er2.seek(0)
                st.download_button("⬇️ Baixar Log ERRO (CR/CP)", bio_er2, file_name="Log_PIX_ERRO.xlsx", use_container_width=True)
        with coly2:
            if st.button("📝 Gravar logs no Google Sheets (CP)", use_container_width=True):
                _save_logs_to_sheet()
        with coly3:
            if st.button("🧹 Limpar logs (sessão) (CP)", use_container_width=True):
                st.session_state["pix_log_ok"] = []
                st.session_state["pix_log_err"] = []
                st.success("Logs da sessão limpos.")

    if cp_ready:
        df_imp = _build_importador_df(
            df_raw, "cp",
            gsel, esel,
            st.session_state.get("cp_portador","")
        )
        if do_pix_cp:
            df_imp = _classificar_pix_em_df(
                df_imp, modulo="CP", grupo=gsel, empresa=esel,
                banco_nome=st.session_state.get("cp_portador","")
            )

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

        _download_excel(edited_full, "Importador_Pagar.xlsx", "📥 Baixar Importador (Pagar)", disabled=not st.session_state.get("cp_edited_once", False))
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
