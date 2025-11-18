# pages/Conciliação Bancária.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import re
import json
from io import BytesIO
from datetime import datetime
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

# Configuração da página
st.set_page_config(page_title="Padronizador de Extratos Bancários", layout="wide")

# 🔒 Bloqueio de acesso
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ===== CSS =====
st.markdown("""
<style>
  [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
  .stApp { background-color: #f9f9f9; }
  .upload-box {
      border: 2px dashed #0366d6;
      border-radius: 10px;
      padding: 20px;
      text-align: center;
      background-color: #f0f8ff;
  }
  .info-card {
      background-color: white;
      border-radius: 8px;
      padding: 15px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      margin: 10px 0;
  }
  .success-badge {
      background-color: #28a745;
      color: white;
      padding: 5px 10px;
      border-radius: 5px;
      font-weight: bold;
  }
  .error-badge {
      background-color: #dc3545;
      color: white;
      padding: 5px 10px;
      border-radius: 5px;
      font-weight: bold;
  }
</style>
""", unsafe_allow_html=True)

# ===== Cabeçalho =====
st.markdown("""
  <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
      <img src='https://img.icons8.com/color/48/bank.png' width='40'/>
      <h1 style='display: inline; margin: 0; font-size: 2.0rem;'>Padronizador de Extratos Bancários</h1>
  </div>
""", unsafe_allow_html=True)

# ======================
# Funções auxiliares
# ======================

def gs_client():
    """Cria cliente Google Sheets"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
    if secret is None:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")
    credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(creds)

def _open_planilha(title="Vendas diarias"):
    """Abre planilha do Google Sheets"""
    try:
        gc = gs_client()
        return gc.open(title)
    except Exception as e:
        st.error(f"⚠️ Erro ao abrir planilha: {e}")
        return None

@st.cache_data(show_spinner=False)
def carregar_empresas():
    """Carrega dados das empresas da planilha"""
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return pd.DataFrame(), [], {}
    
    try:
        ws = sh.worksheet("Tabela Empresa")
        df = pd.DataFrame(ws.get_all_records())
    except Exception as e:
        st.warning(f"⚠️ Erro lendo 'Tabela Empresa': {e}")
        return pd.DataFrame(), [], {}
    
    # Normalizar nomes de colunas
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
def carregar_bancos():
    """Carrega lista de bancos da aba Portador"""
    sh = _open_planilha("Vendas diarias")
    if sh is None:
        return []
    
    try:
        ws = sh.worksheet("Portador")
        df = pd.DataFrame(ws.get_all_records())
        
        # Procura coluna de banco
        col_banco = None
        for col in df.columns:
            if "banco" in str(col).lower():
                col_banco = col
                break
        
        if col_banco:
            bancos = df[col_banco].astype(str).str.strip().unique().tolist()
            return sorted([b for b in bancos if b and b != ""])
        
        return []
    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar bancos: {e}")
        return []

def extrair_periodo_do_texto(texto):
    """Tenta extrair período (datas) do texto"""
    try:
        # Procura padrões de data
        padroes = [
            r'(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})',
            r'(\d{2}/\d{2}/\d{4})\s*até\s*(\d{2}/\d{2}/\d{4})',
            r'Período:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})',
            r'De\s*(\d{2}/\d{2}/\d{4})\s*até\s*(\d{2}/\d{2}/\d{4})',
        ]
        
        for padrao in padroes:
            match = re.search(padrao, texto)
            if match:
                return match.group(1), match.group(2)
        
        return None, None
    except Exception:
        return None, None

def processar_arquivo_upload(uploaded_file):
    """Processa arquivo e tenta extrair informações"""
    try:
        # Tenta ler como texto
        if uploaded_file.name.endswith(('.txt', '.csv')):
            texto = uploaded_file.read().decode('utf-8', errors='ignore')
            uploaded_file.seek(0)
            return texto
        
        # Para Excel
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
            uploaded_file.seek(0)
            # Converte para texto para buscar datas
            texto = df.to_string()
            return texto
        
        # Para PDF (sem biblioteca externa, apenas avisa)
        elif uploaded_file.name.endswith('.pdf'):
            return "[PDF] - Informe o período manualmente"
        
        return ""
    except Exception as e:
        st.warning(f"⚠️ Não foi possível processar o arquivo: {e}")
        return ""

def validar_dados(grupo, loja, banco, agencia, conta, data_inicio, data_fim):
    """Valida se todos os dados necessários foram preenchidos"""
    erros = []
    
    if not grupo or grupo == "— selecione —":
        erros.append("Grupo não selecionado")
    if not loja or loja == "— selecione —":
        erros.append("Loja não selecionada")
    if not banco or banco == "— selecione —":
        erros.append("Banco não selecionado")
    if not agencia:
        erros.append("Agência não informada")
    if not conta:
        erros.append("Conta não informada")
    if not data_inicio:
        erros.append("Data inicial não informada")
    if not data_fim:
        erros.append("Data final não informada")
    
    return erros

def gerar_nome_padronizado(grupo, loja, banco, agencia, conta, data_inicio, data_fim):
    """Gera nome padronizado do arquivo"""
    # Remove caracteres especiais
    grupo_limpo = re.sub(r'[^\w\s-]', '', grupo).strip()
    loja_limpa = re.sub(r'[^\w\s-]', '', loja).strip()
    banco_limpo = re.sub(r'[^\w\s-]', '', banco).strip()
    
    # Formata datas
    try:
        if isinstance(data_inicio, str):
            if "-" in data_inicio:  # formato YYYY-MM-DD
                dt_ini = datetime.strptime(data_inicio, "%Y-%m-%d").strftime("%d-%m-%Y")
            else:  # formato DD/MM/YYYY
                dt_ini = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%d-%m-%Y")
        else:
            dt_ini = data_inicio.strftime("%d-%m-%Y")
            
        if isinstance(data_fim, str):
            if "-" in data_fim:
                dt_fim = datetime.strptime(data_fim, "%Y-%m-%d").strftime("%d-%m-%Y")
            else:
                dt_fim = datetime.strptime(data_fim, "%d/%m/%Y").strftime("%d-%m-%Y")
        else:
            dt_fim = data_fim.strftime("%d-%m-%Y")
    except:
        dt_ini = str(data_inicio)
        dt_fim = str(data_fim)
    
    nome = f"{grupo_limpo} - {loja_limpa} - {banco_limpo} - Ag {agencia} - CC {conta} - {dt_ini} a {dt_fim}.pdf"
    return nome

def salvar_no_sheets(grupo, loja, banco, agencia, conta, data_inicio, data_fim, nome_arquivo):
    """Salva registro do extrato na planilha Google Sheets"""
    try:
        sh = _open_planilha("Vendas diarias")
        if not sh:
            return False, "Planilha não encontrada"
        
        # Nome da planilha de controle
        nome_planilha_controle = "Controle Extratos Bancários"
        
        # Tenta abrir ou criar a planilha de controle
        try:
            ws = sh.worksheet(nome_planilha_controle)
        except WorksheetNotFound:
            ws = sh.add_worksheet(nome_planilha_controle, rows=1000, cols=20)
            # Adiciona cabeçalho
            ws.append_row([
                "Data Processamento", "Grupo", "Loja", "Banco", 
                "Agência", "Conta", "Período Início", "Período Fim", 
                "Nome Arquivo", "Status"
            ])
        
        # Adiciona registro
        ws.append_row([
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            grupo,
            loja,
            banco,
            agencia,
            conta,
            str(data_inicio),
            str(data_fim),
            nome_arquivo,
            "Processado"
        ])
        
        return True, f"Registro salvo em '{nome_planilha_controle}'"
    
    except Exception as e:
        return False, f"Erro ao salvar: {str(e)}"

# ======================
# Carregamento de dados
# ======================
df_emp, GRUPOS, LOJAS_MAP = carregar_empresas()
BANCOS = carregar_bancos()

# ======================
# Interface principal
# ======================

st.markdown("### 📤 Upload de Extrato Bancário")

# Upload de arquivo
uploaded_file = st.file_uploader(
    "Selecione o arquivo do extrato (PDF, Excel, CSV, TXT)",
    type=['pdf', 'xlsx', 'xls', 'csv', 'txt'],
    help="Faça upload do extrato bancário para padronização"
)

if uploaded_file:
    st.success(f"✅ Arquivo carregado: **{uploaded_file.name}**")
    
    # Processa arquivo e tenta extrair período
    texto_arquivo = processar_arquivo_upload(uploaded_file)
    data_ini_auto, data_fim_auto = extrair_periodo_do_texto(texto_arquivo)
    
    st.markdown("---")
    st.markdown("### 🔍 Validação e Padronização")
    
    # Formulário de validação
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 🏢 Dados da Empresa")
        grupo_sel = st.selectbox("Grupo:", ["— selecione —"] + GRUPOS, key="grupo_extrato")
        
        lojas = LOJAS_MAP.get(grupo_sel, []) if grupo_sel != "— selecione —" else []
        loja_sel = st.selectbox("Loja:", ["— selecione —"] + lojas, key="loja_extrato")
        
        st.markdown("#### 🏦 Dados Bancários")
        banco_sel = st.selectbox("Banco:", ["— selecione —"] + BANCOS, key="banco_extrato")
        
        col_ag, col_cc = st.columns(2)
        with col_ag:
            agencia = st.text_input("Agência:", key="agencia_extrato", placeholder="Ex: 1234")
        with col_cc:
            conta = st.text_input("Conta:", key="conta_extrato", placeholder="Ex: 56789-0")
    
    with col2:
        st.markdown("#### 📅 Período do Extrato")
        
        if data_ini_auto and data_fim_auto:
            st.info(f"📌 Período detectado: {data_ini_auto} a {data_fim_auto}")
            try:
                data_ini_default = datetime.strptime(data_ini_auto, "%d/%m/%Y")
                data_fim_default = datetime.strptime(data_fim_auto, "%d/%m/%Y")
            except:
                data_ini_default = None
                data_fim_default = None
        else:
            data_ini_default = None
            data_fim_default = None
        
        data_inicio = st.date_input(
            "Data Inicial:",
            value=data_ini_default,
            key="data_ini_extrato"
        )
        
        data_fim = st.date_input(
            "Data Final:",
            value=data_fim_default,
            key="data_fim_extrato"
        )
        
        st.markdown("#### 📄 Nome do Arquivo Padronizado")
        
        if all([grupo_sel != "— selecione —", loja_sel != "— selecione —", 
                banco_sel != "— selecione —", agencia, conta, data_inicio, data_fim]):
            nome_padrao = gerar_nome_padronizado(
                grupo_sel, loja_sel, banco_sel, agencia, conta,
                str(data_inicio), str(data_fim)
            )
            st.code(nome_padrao, language=None)
        else:
            st.warning("⚠️ Preencha todos os campos para visualizar o nome padronizado")
    
    st.markdown("---")
    
    # Validação
    erros = validar_dados(
        grupo_sel, loja_sel, banco_sel, agencia, conta,
        str(data_inicio) if data_inicio else None,
        str(data_fim) if data_fim else None
    )
    
    if erros:
        st.error("❌ **Erros de validação:**")
        for erro in erros:
            st.markdown(f"- {erro}")
    else:
        st.success("✅ **Todos os dados foram validados com sucesso!**")
        
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            # Gera nome padronizado
            nome_padrao = gerar_nome_padronizado(
                grupo_sel, loja_sel, banco_sel, agencia, conta,
                str(data_inicio), str(data_fim)
            )
            
            # Lê o arquivo original
            file_bytes = uploaded_file.read()
            uploaded_file.seek(0)
            
            st.download_button(
                label="📥 Baixar PDF Padronizado",
                data=file_bytes,
                file_name=nome_padrao,
                mime="application/pdf" if uploaded_file.name.endswith('.pdf') else "application/octet-stream",
                use_container_width=True,
                type="primary"
            )
        
        with col_btn2:
            # Botão para salvar registro no Google Sheets
            if st.button("📊 Registrar no Google Sheets", use_container_width=True):
                with st.spinner("Salvando registro..."):
                    sucesso, mensagem = salvar_no_sheets(
                        grupo_sel, loja_sel, banco_sel, agencia, conta,
                        str(data_inicio), str(data_fim), nome_padrao
                    )
                    
                    if sucesso:
                        st.success(f"✅ {mensagem}")
                        # Adiciona ao histórico
                        if "historico_extratos" not in st.session_state:
                            st.session_state.historico_extratos = []
                        
                        st.session_state.historico_extratos.append({
                            "arquivo": uploaded_file.name,
                            "arquivo_padrao": nome_padrao,
                            "grupo": grupo_sel,
                            "loja": loja_sel,
                            "banco": banco_sel,
                            "agencia": agencia,
                            "conta": conta,
                            "periodo": f"{data_inicio} a {data_fim}",
                            "data_processamento": datetime.now().strftime("%d/%m/%Y %H:%M")
                        })
                    else:
                        st.error(f"❌ {mensagem}")

else:
    st.info("👆 Faça upload de um arquivo de extrato bancário para começar")

# ======================
# Histórico de uploads
# ======================
if st.session_state.get("historico_extratos"):
    st.markdown("---")
    st.markdown("### 📜 Histórico de Processamentos (Sessão Atual)")
    df_hist = pd.DataFrame(st.session_state.historico_extratos)
    st.dataframe(df_hist, use_container_width=True, height=250)
    
    # Botão para limpar histórico
    if st.button("🗑️ Limpar Histórico"):
        st.session_state.historico_extratos = []
        st.rerun()

# ======================
# Seção de ajuda
# ======================
with st.expander("ℹ️ Como usar este módulo"):
    st.markdown("""
    ### 📖 Instruções de uso:
    
    1. **Upload do arquivo**: Faça upload do extrato bancário (PDF, Excel, CSV ou TXT)
    
    2. **Detecção automática**: O sistema tentará detectar automaticamente o período do extrato
    
    3. **Preenchimento dos dados**:
       - Selecione o **Grupo** e a **Loja**
       - Selecione o **Banco**
       - Informe **Agência** e **Conta**
       - Confirme ou ajuste o **período** (datas)
    
    4. **Padronização**: O sistema gerará automaticamente o nome padronizado:
       ```
       Grupo - Loja - Banco - Ag XXXX - CC YYYY - DD-MM-AAAA a DD-MM-AAAA.pdf
       ```
    
    5. **Salvamento**:
       - **Baixar PDF**: Download do arquivo com nome padronizado
       - **Registrar no Sheets**: Salva o registro na planilha de controle
    
    ### 📁 Estrutura planejada no Drive:
    ```
    📁 Conciliação Bancária
      └─ 📁 [Nome do Grupo]
          └─ 📁 [Nome da Loja]
              ├─ 📊 Extratos Bancários.xlsx
              └─ 📄 [Extratos em PDF padronizados]
    ```
    
    ### ✅ Benefícios:
    - ✔️ Nomenclatura padronizada
    - ✔️ Fácil localização de extratos
    - ✔️ Controle centralizado no Google Sheets
    - ✔️ Histórico de processamentos
    """)
