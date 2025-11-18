# pages/Conciliação Bancária.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re
import json
from datetime import datetime
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

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
    Cria um DF padronizado: Grupo, Loja, Banco, Agencia, ContaCorrente
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

    # Vamos referenciar por letra de coluna, como você informou:
    # F (6) -> Grupo
    # B (2) -> Empresa
    # G (7) -> Banco
    # M (13) -> Agência
    # N (14) -> Conta Corrente
    # Como gspread trouxe por nome, usamos iloc para garantir se faltar cabeçalho.
    df = pd.DataFrame()
    try:
        df["Grupo"] = df_raw.iloc[:, 5]   # F
        df["Loja"] = df_raw.iloc[:, 1]    # B
        df["Banco"] = df_raw.iloc[:, 6]   # G
        df["Agencia"] = df_raw.iloc[:, 12]  # M
        df["ContaCorrente"] = df_raw.iloc[:, 13]  # N
    except Exception as e:
        st.error(f"Erro ao mapear colunas da aba 'Fluxo de Caixa': {e}")
        return pd.DataFrame()

    # Limpa espaços
    for c in ["Grupo", "Loja", "Banco", "Agencia", "ContaCorrente"]:
        df[c] = df[c].astype(str).str.strip()

    # Remove linhas totalmente vazias
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
# Carregar bases
# ======================
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
df_fluxo = carregar_fluxo_caixa()

# ======================
# UI Principal
# ======================
st.markdown("### 📤 Upload de Extrato Bancário")

uploaded_file = st.file_uploader(
    "Selecione o arquivo do extrato (PDF, Excel, CSV, TXT)",
    type=["pdf", "xlsx", "xls", "csv", "txt"],
    help="Arquivo do extrato bancário"
)

# Seleção de Grupo e Loja
col_g, col_l = st.columns(2)
with col_g:
    grupo_sel = st.selectbox("Grupo:", ["— selecione —"] + GRUPOS, key="cb_grupo")
with col_l:
    lojas = LOJAS_MAP.get(grupo_sel, []) if grupo_sel and grupo_sel != "— selecione —" else []
    loja_sel = st.selectbox("Loja / Empresa:", ["— selecione —"] + lojas, key="cb_loja")

# Filtra as contas da aba Fluxo de Caixa
contas_filtradas = pd.DataFrame()
if grupo_sel not in (None, "", "— selecione —") and loja_sel not in (None, "", "— selecione —"):
    contas_filtradas = df_fluxo[
        (df_fluxo["Grupo"] == grupo_sel) &
        (df_fluxo["Loja"] == loja_sel)
    ]

st.markdown("### 🏦 Seleção de Conta (Fluxo de Caixa)")

if contas_filtradas.empty:
    st.info("Nenhuma conta encontrada na aba **Fluxo de Caixa** para este Grupo/Loja.")
    conta_opcoes = []
else:
    # Cria label amigável para escolha
    contas_filtradas = contas_filtradas.reset_index(drop=True)
    contas_filtradas["label"] = contas_filtradas.apply(
        lambda r: f"{r['Banco']} - Ag {r['Agencia']} - CC {r['ContaCorrente']}",
        axis=1
    )
    conta_labels = contas_filtradas["label"].tolist()
    conta_escolhida = st.selectbox(
        "Selecione a conta (Banco / Agência / Conta) conforme cadastro no Fluxo de Caixa:",
        ["— selecione —"] + conta_labels,
        key="cb_conta"
    )

    if conta_escolhida != "— selecione —":
        linha_sel = contas_filtradas[contas_filtradas["label"] == conta_escolhida].iloc[0]
        banco_sel = linha_sel["Banco"]
        agencia_sel = linha_sel["Agencia"]
        conta_sel = linha_sel["ContaCorrente"]
    else:
        banco_sel = agencia_sel = conta_sel = ""

# Período do extrato
st.markdown("### 📅 Período do Extrato")
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_inicio = st.date_input("Data Inicial:", key="cb_dt_ini")
with col_d2:
    data_fim = st.date_input("Data Final:", key="cb_dt_fim")

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
    st.warning("Preencha Grupo, Loja, Conta (Fluxo de Caixa), Período e faça o upload do arquivo para gerar o nome padronizado.")
    nome_padrao = None

st.markdown("### ✅ Ações")

col_a1, col_a2 = st.columns(2)
with col_a1:
    if dados_ok and nome_padrao:
        # leitura bruta para download
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
    - Quando você escolhe **Grupo** e **Loja**, são listadas apenas as contas dessa combinação.
    - Ao escolher a conta, o sistema preenche automaticamente **Banco**, **Agência** e **Conta** e monta o nome:
      `Grupo - Loja - Banco - Ag XXXX - CC YYYYY - dd-mm-aaaa a dd-mm-aaaa.pdf`
    - O botão **Registrar extrato no Google Sheets** grava um log na aba **Controle Extratos Bancários**.
    """)
