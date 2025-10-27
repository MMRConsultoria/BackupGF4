# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re
import json
import unicodedata
from io import StringIO, BytesIO
from datetime import datetime

import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

# ---------- CONFIG ----------
st.set_page_config(page_title="CR-CP Importador Everest (Safe)", layout="wide")
st.set_option("client.showErrorDetails", False)
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ---------- HELPERS ----------
def _strip_accents_keep_case(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII","ignore").decode("ASCII")

def _norm_basic(s):
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+"," ", s).strip().lower()
    return s

def _try_parse_paste(text):
    text = (text or "").strip("\n\r ")
    if not text:
        return pd.DataFrame()
    first = text.splitlines()[0] if text else ""
    if "\t" in first:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str, engine="python")
    else:
        try:
            df = pd.read_csv(StringIO(text), sep=";", dtype=str, engine="python")
        except Exception:
            df = pd.read_csv(StringIO(text), sep=",", dtype=str, engine="python")
    df = df.dropna(how="all")
    df.columns = [str(c).strip() if str(c).strip() else "col_%d"%i for i,c in enumerate(df.columns)]
    return df

def _to_float_br(x):
    s = str(x or "").strip()
    s = s.replace("R$","").replace(" ","").replace(".","").replace(",",".")
    try:
        return float(s)
    except:
        return None

def _tokenize(txt):
    return [w for w in re.findall(r"[0-9a-zA-Z]+", _norm_basic(txt)) if w]

# ---------- SHEETS ----------
def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] nao encontrado.")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    try:
        gc = gs_client()
    except Exception as e:
        st.warning("Falha ao criar cliente Google: %s" % e)
        return None
    try:
        return gc.open(title)
    except Exception as e_title:
        sid = st.secrets.get("VENDAS_DIARIAS_SHEET_ID")
        if sid:
            try:
                return gc.open_by_key(sid)
            except Exception as e_id:
                st.warning("Nao consegui abrir a planilha. Erros: %s | %s" % (e_title, e_id))
                return None
        st.warning("Nao consegui abrir por titulo. Detalhes: %s" % e_title)
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        df_vazio = pd.DataFrame(columns=["Grupo","Loja","Codigo Everest","Codigo Grupo Everest","CNPJ"])
        return df_vazio, [], {}
    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning("Erro lendo 'Tabela Empresa': %s" % e)
        df = pd.DataFrame(columns=["Grupo","Loja","Codigo Everest","Codigo Grupo Everest","CNPJ"])

    ren = {"Código Everest":"Codigo Everest","Código Grupo Everest":"Codigo Grupo Everest","Loja Nome":"Loja","Empresa":"Loja","Grupo Nome":"Grupo"}
    df = df.rename(columns={k:v for k,v in ren.items() if k in df.columns})
    for c in ["Grupo","Loja","Codigo Everest","Codigo Grupo Everest","CNPJ"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    lojas_map = (
        df.groupby("Grupo")["Loja"].apply(lambda s: sorted(pd.Series(s.dropna().unique()).astype(str).tolist())).to_dict()
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

@st.cache_data(show_spinner=False)
def carregar_tabela_meio_pagto():
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
        st.warning("Aba 'Tabela Meio Pagamento' nao encontrada.")
        return pd.DataFrame(), [], None

    df = pd.DataFrame(ws.get_all_records()).astype(str)

    for c in [COL_PADRAO, COL_COD, COL_CNPJ, COL_PIXPAD]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip()

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

    pix_vals = [v for v in df[COL_PIXPAD].tolist() if str(v).strip()]
    pix_default = pix_vals[0].strip() if pix_vals else ""
    if pix_vals and len(set(pix_vals)) > 1:
        st.warning("Existem valores diferentes em 'PIX Padrão Cod Gerencial'. Usando o primeiro.")
    pix_default = pix_default if pix_default else None
    return df, rules, pix_default

def _best_rule_for_tokens(ref_tokens):
    best = None
    best_hits = 0
    best_tokens_len = 0
    for rule in MEIO_RULES:
        tokens = set(rule["tokens"])
        hits = len(tokens & ref_tokens)
        if hits == 0:
            continue
        if (hits > best_hits) or (hits == best_hits and len(tokens) > best_tokens_len):
            best = rule
            best_hits = hits
            best_tokens_len = len(tokens)
    return best

def _match_bandeira_to_gerencial(ref_text):
    if not ref_text or not MEIO_RULES:
        return "", "", ""
    ref_tokens = set(_tokenize(ref_text))
    if not ref_tokens:
        return "", "", ""
    best = _best_rule_for_tokens(ref_tokens)
    if best:
        return best["codigo_gerencial"], best.get("cnpj_bandeira",""), ""
    return "", "", ""

# -------- base --------
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
PORTADORES, MAPA_BANCO_PARA_PORTADOR = carregar_portadores()
DF_MEIO, MEIO_RULES, PIX_DEFAULT_CODE = carregar_tabela_meio_pagto()

# ---------- LOGS ----------
def _init_logs():
    st.session_state.setdefault("pix_log_ok", [])
    st.session_state.setdefault("pix_log_err", [])

def _add_log_ok(payload):
    _init_logs()
    st.session_state["pix_log_ok"].append(payload)

def _add_log_err(payload):
    _init_logs()
    st.session_state["pix_log_err"].append(payload)

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _log_pack(modulo, grupo, empresa, banco_nome, row, cod_antes, cod_depois, ref, motivo):
    return {
        "Quando": _now(),
        "Módulo": modulo,
        "Grupo": grupo,
        "Empresa": empresa,
        "Banco/Portador": banco_nome,
        "Data": str(row.get("Data","")),
        "Valor": row.get("Valor Original",""),
        "Referência": ref,
        "CNPJ Empresa": str(row.get("CNPJ Empresa","")),
        "CNPJ/Cliente": str(row.get("CNPJ/Cliente","")),
        "Cod Antes": cod_antes,
        "Cod Depois": cod_depois,
        "Motivo/Regra": motivo,
    }

def _logs_to_df():
    ok = pd.DataFrame(st.session_state.get("pix_log_ok", []))
    err = pd.DataFrame(st.session_state.get("pix_log_err", []))
    cols = ["Quando","Módulo","Grupo","Empresa","Banco/Portador","Data","Valor","Referência",
            "CNPJ Empresa","CNPJ/Cliente","Cod Antes","Cod Depois","Motivo/Regra"]
    ok = ok.reindex(columns=cols, fill_value="")
    err = err.reindex(columns=cols, fill_value="")
    return ok, err

def _save_logs_to_sheet():
    ok_df, err_df = _logs_to_df()
    if ok_df.empty and err_df.empty:
        st.info("Nao ha logs a gravar.")
        return
    sh = _open_planilha("Vendas diarias")
    if not sh:
        st.error("Planilha 'Vendas diarias' indisponivel.")
        return
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
    st.success("Logs gravados.")

# ---------- PIX ----------
PIX_PATTERNS = [
    r"\bpix\b",
    r"\bqr\b",
    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
    r"[A-Z0-9]{25,}",
]
def _is_pix_reference(ref_text):
    if not ref_text: return False
    t = _norm_basic(ref_text)
    for pat in PIX_PATTERNS:
        if re.search(pat, t):
            return True
    if any(w in t for w in ["chave","txid","qr code","qrcode","pagamento instantaneo","instantaneo"]):
        return True
    return False

def _classificar_pix_em_df(df_importador, modulo, grupo, empresa, banco_nome):
    if df_importador.empty:
        return df_importador
    _init_logs()
    df = df_importador.copy()
    col_ref = "Observações do Título"
    col_cod = "Cód Conta Gerencial"
    if col_ref not in df.columns:
        return df
    if col_cod not in df.columns:
        df[col_cod] = ""
    for idx, row in df.iterrows():
        ref = str(row.get(col_ref, "") or "")
        cod_antes = str(row.get(col_cod, "") or "").strip()
        is_pix = _is_pix_reference(ref)
        if not is_pix:
            continue
        if cod_antes:
            _add_log_ok(_log_pack(modulo, grupo, empresa, banco_nome, row, cod_antes, cod_antes, ref,
                                  "PIX detectado, mas ja havia classificacao anterior — mantido"))
            continue
        if not PIX_DEFAULT_CODE:
            _add_log_err(_log_pack(modulo, grupo, empresa, banco_nome, row, cod_antes, "", ref,
                                   "PIX detectado, mas 'PIX Padrão Cod Gerencial' esta vazio"))
            continue
        df.at[idx, col_cod] = PIX_DEFAULT_CODE
        _add_log_ok(_log_pack(modulo, grupo, empresa, banco_nome, row, cod_antes, PIX_DEFAULT_CODE, ref,
                              "Classificado via 'PIX Padrão Cod Gerencial'"))
    if "CNPJ/Cliente" in df.columns:
        flag = df["CNPJ/Cliente"].astype(str).str.strip().eq("")
        df.insert(0, "Falta CNPJ?", flag)
    return df

# ---------- UI helpers ----------
def LOJAS_DO(grupo_nome):
    return (LOJAS_MAP or {}).get(grupo_nome, [])

IMPORTADOR_ORDER = [
    "CNPJ Empresa","Série Título","Nº Título","Nº Parcela","Nº Documento","CNPJ/Cliente","Portador",
    "Data Documento","Data Vencimento","Data","Valor Desconto","Valor Multa","Valor Juros Dia",
    "Valor Original","Observações do Título","Cód Conta Gerencial","Cód Centro de Custo"
]

def _on_paste_change(prefix):
    txt = st.session_state.get(f"{prefix}_paste", "")
    if not str(txt).strip():
        st.session_state.pop(f"{prefix}_df_imp", None)
        st.session_state.pop(f"{prefix}_edited_once", None)

def bloco_colagem(prefix):
    c1,c2 = st.columns([0.65,0.35])
    with c1:
        txt = st.text_area("Colar tabela (Ctrl+V)", height=160, key=f"{prefix}_paste",
                           on_change=_on_paste_change, args=(prefix,))
        df_paste = _try_parse_paste(txt) if (txt and str(txt).strip()) else pd.DataFrame()
    with c2:
        show_prev = st.checkbox("Mostrar pré-visualização", value=False, key=f"{prefix}_show_prev")
        if show_prev and not df_paste.empty:
            st.dataframe(df_paste, use_container_width=True, height=120)
        elif df_paste.empty:
            st.info("Cole dados para prosseguir.")
    return df_paste

def _column_mapping_ui(prefix, df_raw):
    cols = ["— selecione —"] + list(df_raw.columns)
    c1,c2,c3 = st.columns(3)
    with c1: st.selectbox("Coluna de Data", cols, key=f"{prefix}_col_data")
    with c2: st.selectbox("Coluna de Valor", cols, key=f"{prefix}_col_valor")
    with c3: st.selectbox("Coluna de Referência", cols, key=f"{prefix}_col_bandeira")

def _build_importador_df(df_raw, prefix, grupo, loja, banco_escolhido):
    cd = st.session_state.get(f"{prefix}_col_data")
    cv = st.session_state.get(f"{prefix}_col_valor")
    cb = st.session_state.get(f"{prefix}_col_bandeira")
    if not cd or not cv or not cb or "— selecione —" in (cd, cv, cb):
        return pd.DataFrame()

    cnpj_loja = ""
    if not df_emp.empty and loja:
        row = df_emp[(df_emp["Loja"].astype(str).str.strip()==loja) & (df_emp["Grupo"].astype(str).str.strip()==grupo)]
        if not row.empty: cnpj_loja = str(row.iloc[0].get("CNPJ","") or "")

    portador_nome = (MAPA_BANCO_PARA_PORTADOR or {}).get(banco_escolhido or "", banco_escolhido or "")

    data_original  = df_raw[cd].astype(str)
    valor_original = pd.to_numeric(df_raw[cv].apply(_to_float_br), errors="coerce").round(2)
    ref_txt        = df_raw[cb].astype(str).str.strip()

    cod_conta_list, cnpj_cli_list = [], []
    for b in ref_txt:
        cod, cnpj_band, _ = _match_bandeira_to_gerencial(b)
        cod_conta_list.append(cod)
        cnpj_cli_list.append(cnpj_band)

    out = pd.DataFrame({
        "CNPJ Empresa": cnpj_loja,
        "Série Título": "DRE",
        "Nº Título": "",
        "Nº Parcela": 1,
        "Nº Documento": "DRE",
        "CNPJ/Cliente": cnpj_cli_list,
        "Portador": portador_nome,
        "Data Documento": data_original,
        "Data Vencimento": data_original,
        "Data": data_original,
        "Valor Desconto": 0.00,
        "Valor Multa": 0.00,
        "Valor Juros Dia": 0.00,
        "Valor Original": valor_original,
        "Observações do Título": ref_txt.tolist(),
        "Cód Conta Gerencial": cod_conta_list,
        "Cód Centro de Custo": 3,
    })
    out = out[(out["Data"].astype(str).str.strip()!="") & (out["Valor Original"].notna())]
    out = out.reindex(columns=[c for c in IMPORTADOR_ORDER if c in out.columns])
    out.insert(0, "Falta CNPJ?", out["CNPJ/Cliente"].astype(str).str.strip().eq(""))
    final_cols = ["Falta CNPJ?"] + [c for c in IMPORTADOR_ORDER if c in out.columns]
    out = out[final_cols]
    return out

def _download_excel(df, filename, label_btn, disabled=False):
    if df.empty:
        st.button(label_btn, disabled=True, use_container_width=True)
        return
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Importador")
    bio.seek(0)
    st.download_button(label_btn, data=bio, file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, disabled=disabled)

def filtros_grupo_empresa(prefix, with_portador=False, with_tipo_imp=False):
    c1, c2, c3, c4 = st.columns(4)
    grupos = list(GRUPOS or [])
    with c1:
        gsel = st.selectbox("Grupo", ["— selecione —"] + grupos, key=f"{prefix}_grupo")
    with c2:
        lojas = LOJAS_DO(gsel) if gsel and gsel!="— selecione —" else []
        esel = st.selectbox("Empresa", ["— selecione —"] + lojas, key=f"{prefix}_empresa")
    with c3:
        if with_portador:
            st.selectbox("Banco", ["Todos"] + list(PORTADORES or []), index=0, key=f"{prefix}_portador")
        else:
            st.empty()
    with c4:
        if with_tipo_imp:
            st.selectbox("Tipo de Importacao", ["Todos","Adquirente","Cliente","Outros"], index=0, key=f"{prefix}_tipo_imp")
        else:
            st.empty()
    return gsel, esel

# ---------- ABAS ----------
aba_cr, aba_cp = st.tabs(["Contas a Receber", "Contas a Pagar"])

with aba_cr:
    gsel, esel = filtros_grupo_empresa("cr", with_portador=True, with_tipo_imp=True)
    df_raw = bloco_colagem("cr")
    if st.session_state.get("cr_tipo_imp") == "Adquirente" and not df_raw.empty:
        _column_mapping_ui("cr", df_raw)
    ready = (
        st.session_state.get("cr_tipo_imp")=="Adquirente" and not df_raw.empty and
        all(st.session_state.get(k) and st.session_state.get(k)!="— selecione —" for k in ["cr_col_data","cr_col_valor","cr_col_bandeira"]) and
        gsel not in (None,"","— selecione —") and esel not in (None,"","— selecione —")
    )
    if ready:
        df_imp = _build_importador_df(df_raw, "cr", gsel, esel, st.session_state.get("cr_portador",""))
        df_imp = _classificar_pix_em_df(df_imp, "CR", gsel, esel, st.session_state.get("cr_portador",""))
        st.session_state["cr_df_imp"] = df_imp.copy()

    df_imp = st.session_state.get("cr_df_imp")
    if isinstance(df_imp, pd.DataFrame) and not df_imp.empty:
        show_only = st.checkbox("Mostrar apenas Falta CNPJ", value=False, key="cr_only_missing")
        df_view = df_imp[df_imp["Falta CNPJ?"]] if show_only else df_imp
        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]
        edited = st.data_editor(df_view, disabled=disabled_cols, use_container_width=True, height=420, key="cr_editor")
        full = df_imp.copy(); full.update(edited)
        full["Falta CNPJ?"] = full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        cols_final = ["Falta CNPJ?"] + [c for c in full.columns if c!="Falta CNPJ?"]
        full = full.reindex(columns=cols_final)
        st.session_state["cr_df_imp"] = full
        _download_excel(full, "Importador_Receber.xlsx", "Baixar Importador (Receber)")

with aba_cp:
    gsel, esel = filtros_grupo_empresa("cp", with_portador=True, with_tipo_imp=True)
    df_raw = bloco_colagem("cp")
    if st.session_state.get("cp_tipo_imp") == "Adquirente" and not df_raw.empty:
        _column_mapping_ui("cp", df_raw)
    ready = (
        st.session_state.get("cp_tipo_imp")=="Adquirente" and not df_raw.empty and
        all(st.session_state.get(k) and st.session_state.get(k)!="— selecione —" for k in ["cp_col_data","cp_col_valor","cp_col_bandeira"]) and
        gsel not in (None,"","— selecione —") and esel not in (None,"","— selecione —")
    )
    if ready:
        df_imp = _build_importador_df(df_raw, "cp", gsel, esel, st.session_state.get("cp_portador",""))
        df_imp = _classificar_pix_em_df(df_imp, "CP", gsel, esel, st.session_state.get("cp_portador",""))
        st.session_state["cp_df_imp"] = df_imp.copy()

    df_imp = st.session_state.get("cp_df_imp")
    if isinstance(df_imp, pd.DataFrame) and not df_imp.empty:
        show_only = st.checkbox("Mostrar apenas Falta CNPJ", value=False, key="cp_only_missing")
        df_view = df_imp[df_imp["Falta CNPJ?"]] if show_only else df_imp
        editable = {"CNPJ/Cliente","Cód Conta Gerencial","Cód Centro de Custo"}
        disabled_cols = [c for c in df_view.columns if c not in editable]
        edited = st.data_editor(df_view, disabled=disabled_cols, use_container_width=True, height=420, key="cp_editor")
        full = df_imp.copy(); full.update(edited)
        full["Falta CNPJ?"] = full["CNPJ/Cliente"].astype(str).str.strip().eq("")
        cols_final = ["Falta CNPJ?"] + [c for c in full.columns if c!="Falta CNPJ?"]
        full = full.reindex(columns=cols_final)
        st.session_state["cp_df_imp"] = full
        _download_excel(full, "Importador_Pagar.xlsx", "Baixar Importador (Pagar)")

# --------- LOGS UI ----------
st.divider()
st.caption("Logs de classificação de PIX (sessão)")
ok_df, err_df = _logs_to_df()
c1, c2, c3 = st.columns([0.5,0.25,0.25])
with c1:
    if not ok_df.empty:
        bio_ok = BytesIO()
        with pd.ExcelWriter(bio_ok, engine="openpyxl") as w:
            ok_df.to_excel(w, index=False, sheet_name="Log_OK")
        bio_ok.seek(0)
        st.download_button("Baixar Log OK", bio_ok, file_name="Log_PIX_OK.xlsx", use_container_width=True)
    if not err_df.empty:
        bio_er = BytesIO()
        with pd.ExcelWriter(bio_er, engine="openpyxl") as w:
            err_df.to_excel(w, index=False, sheet_name="Log_ERRO")
        bio_er.seek(0)
        st.download_button("Baixar Log ERRO", bio_er, file_name="Log_PIX_ERRO.xlsx", use_container_width=True)
with c2:
    if st.button("Gravar logs no Google Sheets", use_container_width=True):
        _save_logs_to_sheet()
with c3:
    if st.button("Limpar logs (sessao)", use_container_width=True):
        st.session_state["pix_log_ok"] = []
        st.session_state["pix_log_err"] = []
        st.success("Logs limpos.")
