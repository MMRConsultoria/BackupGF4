# pages/Conciliação Bancária.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re
import json
import io
from datetime import datetime
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

# Para leitura de PDF (adicionar pdfplumber no requirements.txt)
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

st.set_page_config(page_title="Conciliação Bancária - Extratos", layout="wide")

# 🔒 Bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS =====
st.markdown("""
<style>
  [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
  .stApp { background-color: #f9f9f9; }
</style>
""", unsafe_allow_html=True)

# ===== Cabeçalho =====
st.markdown("""
  <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
      <img src='https://img.icons8.com/color/48/bank.png' width='40'/>
      <h1 style='display: inline; margin: 0; font-size: 2.0rem;'>Conciliação Bancária - Padronização de Extratos</h1>
  </div>
""", unsafe_allow_html=True)

# ======================
# Funções auxiliares
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
        return gc.open(title)
    except Exception as e:
        st.error(f"⚠️ Erro ao abrir planilha '{title}': {e}")
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    """Lê Tabela Empresa para montar Grupo x Loja (igual seu outro módulo)."""
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return pd.DataFrame(), [], {}
    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning(f"⚠️ Erro lendo 'Tabela Empresa': {e}")
        return pd.DataFrame(), [], {}

    ren = {
        "Codigo Everest": "Código Everest",
        "Codigo Grupo Everest": "Código Grupo Everest",
        "Loja Nome": "Loja",
        "Empresa": "Loja",
        "Grupo Nome": "Grupo"
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    for c in ["Grupo", "Loja", "Código Everest", "Código Grupo Everest", "CNPJ"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    grupos = sorted(df["Grupo"].dropna().unique().tolist())
    lojas_map = (
        df.groupby("Grupo")["Loja"]
          .apply(lambda s: sorted(pd.Series(s.dropna().unique()).astype(str).tolist()))
          .to_dict()
    )
    return df, grupos, lojas_map

@st.cache_data(show_spinner=False)
def carregar_fluxo_caixa():
    """
    Lê a aba 'Fluxo de Caixa' e mapeia:
    - Grupo  (col F)
    - Empresa (col B)
    - Banco  (col G)
    - Agência (col M)
    - Conta Corrente (col N)
    - Extrato Nome Empresa (coluna com esse cabeçalho na planilha)

    Cria um DF padronizado:
    Grupo, Loja, Banco, Agencia, ContaCorrente, ExtratoNomeEmpresa
    """
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return pd.DataFrame()

    try:
        ws = sh.worksheet("Fluxo de Caixa")
    except WorksheetNotFound:
        st.warning("⚠️ Aba 'Fluxo de Caixa' não encontrada.")
        return pd.DataFrame()

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    header = values[0]
    rows = values[1:]
    df_raw = pd.DataFrame(rows, columns=header)

    df = pd.DataFrame()
    try:
        # Mantendo mapeamento por posição como você já usava:
        df["Grupo"] = df_raw.iloc[:, 5]       # F
        df["Loja"] = df_raw.iloc[:, 1]        # B (Empresa)
        df["Banco"] = df_raw.iloc[:, 6]       # G
        df["Agencia"] = df_raw.iloc[:, 12]    # M
        df["ContaCorrente"] = df_raw.iloc[:, 13]  # N
    except Exception as e:
        st.error(f"Erro ao mapear colunas da aba 'Fluxo de Caixa': {e}")
        return pd.DataFrame()

    # ➕ Novo: coluna "Extrato Nome Empresa" (pelo cabeçalho)
    if "Extrato Nome Empresa" in df_raw.columns:
        df["ExtratoNomeEmpresa"] = df_raw["Extrato Nome Empresa"].astype(str).str.strip()
    else:
        df["ExtratoNomeEmpresa"] = ""

    # Limpa espaços
    for c in ["Grupo", "Loja", "Banco", "Agencia", "ContaCorrente", "ExtratoNomeEmpresa"]:
        df[c] = df[c].astype(str).str.strip()

    # Remove linhas totalmente vazias (nas colunas principais)
    df = df[~(df[["Grupo", "Loja", "Banco", "Agencia", "ContaCorrente"]].eq("").all(axis=1))]
    return df

def gerar_nome_padronizado(grupo, loja, banco, agencia, conta, data_inicio, data_fim):
    grupo_limpo = re.sub(r"[^\w\s-]", "", str(grupo)).strip()
    loja_limpa = re.sub(r"[^\w\s-]", "", str(loja)).strip()
    banco_limpo = re.sub(r"[^\w\s-]", "", str(banco)).strip()

    try:
        if isinstance(data_inicio, str):
            dt_ini = datetime.fromisoformat(data_inicio).strftime("%d-%m-%Y")
        else:
            dt_ini = data_inicio.strftime("%d-%m-%Y")
        if isinstance(data_fim, str):
            dt_fim = datetime.fromisoformat(data_fim).strftime("%d-%m-%Y")
        else:
            dt_fim = data_fim.strftime("%d-%m-%Y")
    except Exception:
        dt_ini = str(data_inicio)
        dt_fim = str(data_fim)

    return f"{grupo_limpo} - {loja_limpa} - {banco_limpo} - Ag {agencia} - CC {conta} - {dt_ini} a {dt_fim}.pdf"

def salvar_registro_extrato(grupo, loja, banco, agencia, conta, data_inicio, data_fim, nome_arquivo):
    """Registra o extrato em uma aba de controle no Google Sheets."""
    try:
        sh = _open_planilha("Vendas diarias")
        if not sh:
            return False, "Planilha 'Vendas diarias' não encontrada."

        nome_aba = "Controle Extratos Bancários"
        try:
            ws = sh.worksheet(nome_aba)
        except WorksheetNotFound:
            ws = sh.add_worksheet(nome_aba, rows=1000, cols=20)
            ws.append_row([
                "Data Registro", "Grupo", "Loja", "Banco",
                "Agência", "Conta Corrente", "Período Início",
                "Período Fim", "Nome Arquivo"
            ])

        ws.append_row([
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            grupo,
            loja,
            banco,
            agencia,
            conta,
            str(data_inicio),
            str(data_fim),
            nome_arquivo
        ])
        return True, f"Registro salvo em '{nome_aba}'."
    except Exception as e:
        return False, f"Erro ao salvar registro: {e}"

# ======================
# Funções de reconhecimento automático
# ======================

def extrair_texto_arquivo(file_bytes: bytes, file_name: str) -> str:
    """Transforma o arquivo de extrato em um grande texto (para busca de conta, agência, nome empresa e datas)."""
    nome = file_name.lower()
    ext = nome.split(".")[-1] if "." in nome else ""

    try:
        if ext in ("csv", "txt"):
            try:
                texto = file_bytes.decode("utf-8", errors="ignore")
            except Exception:
                texto = file_bytes.decode("latin-1", errors="ignore")
            return texto

        if ext in ("xlsx", "xls"):
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
            except Exception:
                df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, engine="openpyxl")
            df = df.fillna("")
            return " ".join(df.astype(str).values.ravel().tolist())

        if ext == "pdf" and pdfplumber is not None:
            texto = []
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text() or ""
                    texto.append(t)
            return "\n".join(texto)

    except Exception as e:
        st.warning(f"⚠️ Não foi possível extrair texto do arquivo para reconhecimento automático: {e}")

    return ""

def extrair_datas_do_texto(texto: str):
    """Procura datas no formato dd/mm/aaaa ou dd/mm/aa e retorna (data_min, data_max) ou (None, None)."""
    padrao = r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b"
    matchs = re.findall(padrao, texto)
    datas = []

    for m in matchs:
        for fmt in ("%d/%m/%Y", "%d/%m/%y"):
            try:
                d = datetime.strptime(m, fmt).date()
                datas.append(d)
                break
            except ValueError:
                continue

    if not datas:
        return None, None

    return min(datas), max(datas)

def reconhecer_conta_no_texto(texto: str, df_fluxo: pd.DataFrame):
    """
    Tenta encontrar uma combinação única de Agência + Conta + ExtratoNomeEmpresa no texto,
    comparando com a aba Fluxo de Caixa.
    - Usa:
      - match de Agência (1 ponto)
      - match de Conta (2 pontos)
      - match de "Extrato Nome Empresa" (3 pontos)
    """
    if df_fluxo.empty or not texto:
        return None, "Fluxo de Caixa vazio ou texto não disponível."

    texto_digitos = re.sub(r"\D", "", texto)
    texto_lower = texto.lower()

    melhor_linha = None
    melhor_score = 0
    candidatos = []

    for _, row in df_fluxo.iterrows():
        ag = re.sub(r"\D", "", str(row["Agencia"]))
        cc = re.sub(r"\D", "", str(row["ContaCorrente"]))
        nome_extrato = str(row.get("ExtratoNomeEmpresa", "")).strip()
        nome_extrato_lower = nome_extrato.lower()

        score = 0

        # Agência
        if ag and ag in texto_digitos:
            score += 1

        # Conta
        if cc and cc in texto_digitos:
            score += 2  # peso maior para conta

        # Nome da empresa que aparece no extrato
        if nome_extrato_lower and nome_extrato_lower in texto_lower:
            score += 3  # peso forte para o nome da empresa no extrato

        if score > 0:
            candidatos.append((score, row))

        if score > melhor_score:
            melhor_score = score
            melhor_linha = row

    if melhor_score == 0 or melhor_linha is None:
        return None, "Nenhuma conta/agência/nome de empresa do Fluxo de Caixa foi encontrada no arquivo."

    # Verifica se há ambiguidade (mais de um com mesmo score máximo)
    qtd_max = sum(1 for s, _ in candidatos if s == melhor_score)
    if qtd_max > 1:
        return None, "Mais de uma conta possível encontrada no arquivo (ambiguidade)."

    return melhor_linha, None

def aplicar_reconhecimento_automatico(uploaded_file, df_fluxo, grupos, lojas_map):
    """
    Lê o arquivo enviado, tenta reconhecer:
    - Grupo
    - Loja
    - Banco / Agência / Conta
    - Período (Data Inicial / Data Final)

    Usando:
    - Dígitos de agência e conta
    - Nome da empresa no extrato (coluna 'Extrato Nome Empresa')
    """
    if uploaded_file is None:
        return

    file_id = getattr(uploaded_file, "name", "") or ""
    if st.session_state.get("last_file_id") == file_id and st.session_state.get("auto_aplicado"):
        return

    st.session_state["last_file_id"] = file_id
    st.session_state["auto_aplicado"] = False
    st.session_state["auto_info"] = {}

    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    texto = extrair_texto_arquivo(file_bytes, uploaded_file.name)
    if not texto.strip():
        st.session_state["auto_info"]["mensagem"] = "Não foi possível extrair texto do arquivo para reconhecimento automático."
        return

    linha_conta, erro_conta = reconhecer_conta_no_texto(texto, df_fluxo)
    data_ini_auto, data_fim_auto = extrair_datas_do_texto(texto)

    reconhecido = {}

    if linha_conta is not None:
        grupo = str(linha_conta["Grupo"]).strip()
        loja = str(linha_conta["Loja"]).strip()
        banco = str(linha_conta["Banco"]).strip()
        ag = str(linha_conta["Agencia"]).strip()
        cc = str(linha_conta["ContaCorrente"]).strip()

        reconhecido.update({
            "grupo": grupo,
            "loja": loja,
            "banco": banco,
            "agencia": ag,
            "conta": cc,
            "extrato_nome_empresa": str(linha_conta.get("ExtratoNomeEmpresa", "")).strip()
        })

        if grupo in grupos:
            st.session_state["cb_grupo"] = grupo

        if loja and grupo in lojas_map and loja in lojas_map[grupo]:
            st.session_state["cb_loja"] = loja

        label_conta = f"{banco} - Ag {ag} - CC {cc}"
        st.session_state["cb_conta"] = label_conta

    if data_ini_auto is not None and data_fim_auto is not None:
        reconhecido["data_inicio"] = data_ini_auto
        reconhecido["data_fim"] = data_fim_auto
        st.session_state["cb_dt_ini"] = data_ini_auto
        st.session_state["cb_dt_fim"] = data_fim_auto

    msg_lista = []
    if linha_conta is not None:
        msg_lista.append("✅ Conta/loja reconhecida automaticamente a partir do arquivo (agência, conta e/ou nome da empresa do extrato).")
    elif erro_conta:
        msg_lista.append(f"⚠️ {erro_conta}")

    if data_ini_auto and data_fim_auto:
        msg_lista.append(f"✅ Período provável reconhecido: {data_ini_auto.strftime('%d/%m/%Y')} a {data_fim_auto.strftime('%d/%m/%Y')}")
    else:
        msg_lista.append("⚠️ Não foi possível reconhecer o período completo do extrato.")

    st.session_state["auto_info"]["reconhecido"] = reconhecido
    st.session_state["auto_info"]["mensagem"] = "\n".join(msg_lista)

    st.session_state["auto_aplicado"] = True

# ======================
# Carregar bases
# ======================
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
df_fluxo = carregar_fluxo_caixa()

for key in ["cb_grupo", "cb_loja", "cb_conta", "cb_dt_ini", "cb_dt_fim"]:
    st.session_state.setdefault(key, None)
st.session_state.setdefault("auto_info", {})
st.session_state.setdefault("auto_aplicado", False)
st.session_state.setdefault("last_file_id", "")

# ======================
# UI Principal
# ======================
st.markdown("### 📤 Upload de Extrato Bancário")

uploaded_file = st.file_uploader(
    "Selecione o arquivo do extrato (PDF, Excel, CSV, TXT)",
    type=["pdf", "xlsx", "xls", "csv", "txt"],
    help="Arquivo do extrato bancário"
)

# 🔍 Reconhecimento Automático (após upload)
if uploaded_file is not None:
    with st.spinner("🔍 Lendo o arquivo e tentando reconhecer as informações automaticamente..."):
        aplicar_reconhecimento_automatico(uploaded_file, df_fluxo, GRUPOS, LOJAS_MAP)

    auto_info = st.session_state.get("auto_info", {})
    if auto_info:
        with st.expander("🤖 Informações reconhecidas automaticamente (clique para ver)", expanded=True):
            msg = auto_info.get("mensagem", "")
            if msg:
                for linha in msg.split("\n"):
                    st.markdown(f"- {linha}")
            rec = auto_info.get("reconhecido", {})
            if rec:
                st.markdown("**Resumo dos dados sugeridos:**")
                st.json(rec)
            st.markdown(
                "👉 *Confira as informações abaixo. Você só precisa alterar algo se o reconhecimento estiver incorreto.*"
            )

# Seleção de Grupo e Loja
col_g, col_l = st.columns(2)
with col_g:
    grupo_sel = st.selectbox(
        "Grupo:",
        ["— selecione —"] + GRUPOS,
        index=(
            (["— selecione —"] + GRUPOS).index(st.session_state["cb_grupo"])
            if st.session_state.get("cb_grupo") in GRUPOS
            else 0
        ),
        key="cb_grupo"
    )

with col_l:
    lojas = LOJAS_MAP.get(grupo_sel, []) if grupo_sel and grupo_sel != "— selecione —" else []
    loja_options = ["— selecione —"] + lojas
    if st.session_state.get("cb_loja") in lojas:
        idx_loja = loja_options.index(st.session_state["cb_loja"])
    else:
        idx_loja = 0

    loja_sel = st.selectbox(
        "Loja / Empresa:",
        loja_options,
        index=idx_loja,
        key="cb_loja"
    )

# Contas da aba Fluxo de Caixa
contas_filtradas = pd.DataFrame()
if grupo_sel not in (None, "", "— selecione —") and loja_sel not in (None, "", "— selecione —"):
    contas_filtradas = df_fluxo[
        (df_fluxo["Grupo"] == grupo_sel) &
        (df_fluxo["Loja"] == loja_sel)
    ]

st.markdown("### 🏦 Seleção de Conta (Fluxo de Caixa)")

banco_sel = agencia_sel = conta_sel = ""

if contas_filtradas.empty:
    st.info("Nenhuma conta encontrada na aba **Fluxo de Caixa** para este Grupo/Loja.")
else:
    contas_filtradas = contas_filtradas.reset_index(drop=True)
    contas_filtradas["label"] = contas_filtradas.apply(
        lambda r: f"{r['Banco']} - Ag {r['Agencia']} - CC {r['ContaCorrente']}",
        axis=1
    )
    conta_labels = contas_filtradas["label"].tolist()

    conta_default = st.session_state.get("cb_conta")
    if conta_default in conta_labels:
        idx_conta = conta_labels.index(conta_default) + 1
    else:
        idx_conta = 0

    conta_escolhida = st.selectbox(
        "Selecione a conta (Banco / Agência / Conta) conforme cadastro no Fluxo de Caixa:",
        ["— selecione —"] + conta_labels,
        index=idx_conta,
        key="cb_conta"
    )

    if conta_escolhida != "— selecione —":
        linha_sel = contas_filtradas[contas_filtradas["label"] == conta_escolhida].iloc[0]
        banco_sel = linha_sel["Banco"]
        agencia_sel = linha_sel["Agencia"]
        conta_sel = linha_sel["ContaCorrente"]

# Período do extrato
st.markdown("### 📅 Período do Extrato")
col_d1, col_d2 = st.columns(2)

with col_d1:
    data_inicio = st.date_input(
        "Data Inicial:",
        value=st.session_state.get("cb_dt_ini") or datetime.today().date(),
        key="cb_dt_ini"
    )
with col_d2:
    data_fim = st.date_input(
        "Data Final:",
        value=st.session_state.get("cb_dt_fim") or datetime.today().date(),
        key="cb_dt_fim"
    )

# Nome padronizado
st.markdown("### 📄 Nome Padronizado do Arquivo")

dados_ok = (
    uploaded_file is not None and
    grupo_sel not in (None, "", "— selecione —") and
    loja_sel not in (None, "", "— selecione —") and
    banco_sel not in (None, "", "— selecione —") and
    agencia_sel not in (None, "") and
    conta_sel not in (None, "") and
    data_inicio is not None and
    data_fim is not None
)

if dados_ok:
    nome_padrao = gerar_nome_padronizado(
        grupo_sel, loja_sel, banco_sel, agencia_sel, conta_sel,
        data_inicio, data_fim
    )
    st.code(nome_padrao, language="text")
else:
    st.warning("Preencha/valide Grupo, Loja, Conta (Fluxo de Caixa), Período e faça o upload do arquivo para gerar o nome padronizado.")
    nome_padrao = None

st.markdown("### ✅ Ações")

col_a1, col_a2 = st.columns(2)
with col_a1:
    if dados_ok and nome_padrao and uploaded_file is not None:
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        st.download_button(
            "📥 Baixar arquivo com nome padronizado",
            data=file_bytes,
            file_name=nome_padrao,
            mime="application/pdf" if uploaded_file.name.lower().endswith(".pdf") else "application/octet-stream",
            use_container_width=True,
            type="primary"
        )
    else:
        st.button("📥 Baixar arquivo com nome padronizado", disabled=True, use_container_width=True)

with col_a2:
    if dados_ok and nome_padrao:
        if st.button("📊 Registrar extrato no Google Sheets", use_container_width=True):
            with st.spinner("Registrando extrato..."):
                sucesso, msg = salvar_registro_extrato(
                    grupo_sel, loja_sel, banco_sel, agencia_sel, conta_sel,
                    data_inicio, data_fim, nome_padrao
                )
                if sucesso:
                    st.success(msg)
                else:
                    st.error(msg)
    else:
        st.button("📊 Registrar extrato no Google Sheets", disabled=True, use_container_width=True)

# Ajuda
with st.expander("ℹ️ Como funciona a amarração com a aba 'Fluxo de Caixa'?"):
    st.markdown("""
    - Este módulo lê a aba **Fluxo de Caixa** da planilha *Vendas diarias*.
    - Usa as colunas:
      - **Grupo** → Coluna **F**
      - **Empresa (Loja)** → Coluna **B**
      - **Banco** → Coluna **G**
      - **Agência** → Coluna **M**
      - **Conta Corrente** → Coluna **N**
      - **Extrato Nome Empresa** → nome exato que aparece no extrato do banco
    - Quando você faz o **upload do extrato**, o sistema tenta:
      - Ler o arquivo (PDF/Excel/CSV/TXT),
      - Encontrar **Agência**, **Conta** e o **nome da empresa do extrato**
        de acordo com a coluna **Extrato Nome Empresa**,
      - Cruzar com as contas cadastradas na aba Fluxo de Caixa,
      - Sugerir automaticamente **Grupo, Loja, Banco, Agência e Conta**,
      - Identificar as datas presentes no extrato e sugerir o período.
    - Depois disso, os campos já vêm preenchidos para você **apenas confirmar**.
    - O botão **Registrar extrato no Google Sheets** grava um log na aba **Controle Extratos Bancários**.
    """)
