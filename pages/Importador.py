# -*- coding: utf-8 -*-

import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# ================= FUSÍVEL ANTI-HELP =================
try:
    import builtins
    def _noop_help(*args, **kwargs):
        return None
    builtins.help = _noop_help
except Exception:
    pass

# ================= CONFIG STREAMLIT =================
st.set_page_config(
    page_title="Resumo Contrato",
    layout="wide"
)

st.set_option("client.showErrorDetails", False)

# ================= BLOQUEIO DE ACESSO =================
if not st.session_state.get("acesso_liberado"):
    st.stop()


# ================= CSS PADRÃO =================
st.markdown("""
<style>
  /* Oculta toolbar e menu Streamlit */
  [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
  header { visibility: hidden; }

  /* Aparência geral */
  .stApp { background-color: #f9f9f9; }

  /* Tabs (se houver) */
  div[data-baseweb="tab-list"] { margin-top: 20px; }

  button[data-baseweb="tab"] {
      background-color: #f0f2f6;
      border-radius: 10px;
      padding: 10px 20px;
      margin-right: 10px;
      transition: all 0.3s ease;
      font-size: 16px;
      font-weight: 600;
  }

  button[data-baseweb="tab"]:hover {
      background-color: #dce0ea;
      color: black;
  }

  button[data-baseweb="tab"][aria-selected="true"] {
      background-color: #0366d6;
      color: white;
  }

  hr.compact {
      height: 1px;
      background: #e6e9f0;
      border: none;
      margin: 8px 0 10px;
  }

  .compact [data-testid="stSelectbox"],
  .compact [data-testid="stTextArea"],
  .compact [data-testid="stVerticalBlock"] > div {
      margin-bottom: 8px !important;
  }
</style>
""", unsafe_allow_html=True)

# ================= CABEÇALHO PADRÃO =================
st.markdown("""
  <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 12px;'>
      <img src='https://img.icons8.com/color/48/document.png' width='40'/>
      <h1 style='display: inline; margin: 0; font-size: 2.0rem;'>
          Resumo Contrato
      </h1>
  </div>
""", unsafe_allow_html=True)



# ---------------- utilitários ----------------
_money_re = re.compile(r'^\d{1,3}(?:\.\d{3})*,\d{2}$')
_token_hours_part = re.compile(r'\d+:\d+')

def is_money(tok: str) -> bool:
    t = str(tok or "").strip()
    if not t:
        return False
    if re.match(r'^\d+,\d{2}$', t):
        return True
    return bool(_money_re.match(t))

def _to_float_br(x):
    t = str(x or "").strip()
    if not t or t.upper() in ("NAN", "NONE"):
        return None
    t = t.replace(" ", "")
    has_c = "," in t
    has_p = "." in t
    # Caso formato brasileiro com pontos como milhares e vírgula decimal
    if has_c and has_p:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif has_c:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except:
        return None

_MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

def extrair_mes_ano(periodo_str):
    match = re.search(r"(\d{2})/(\d{2})/(\d{4})", periodo_str)
    if match:
        mes_num = int(match.group(2))
        ano = match.group(3)
        mes_nome = _MONTHS_PT.get(mes_num, "")
        return mes_nome, ano
    return "", ""

# limpa o nome da empresa removendo datas, horas, página, CNPJ etc.
def clean_company_name(raw_name: str) -> str:
    if not raw_name:
        return ""
    s = raw_name.strip()
    # remover intervalos de data "dd/mm/yyyy a dd/mm/yyyy" e datas soltas
    s = re.sub(r'\d{2}/\d{2}/\d{4}\s*(?:a|-)\s*\d{2}/\d{2}/\d{4}', '', s)
    s = re.sub(r'\d{2}/\d{2}/\d{4}', '', s)
    # remover hora hh:mm
    s = re.sub(r'\b\d{1,2}:\d{2}\b', '', s)
    # remover "Pág" ou "Pág." e número de página
    s = re.sub(r'\bPág(?:\.|:)?\s*\d+\b', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\bPage(?:\.|:)?\s*\d+\b', '', s, flags=re.IGNORECASE)
    # remover CNPJ formais
    s = re.sub(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', '', s)
    # remover sequências de números com barras ou traços (por segurança)
    s = re.sub(r'\b\d{1,6}[\/-]\d{1,6}[\/-]?\d*\b', '', s)
    # remover múltiplos espaços
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()

def extract_company_code_and_name(texto: str):
    """
    Extrai código e nome da empresa da linha 'Empresa:'.
    Remove tudo que vier após o nome (datas, hora, Pág, etc).
    Retorna (codigo_str, nome_limpo)
    """
    if not texto:
        return "", ""
    # procurar trecho começando por "Empresa" e capturar código e resto
    m = re.search(r"Empresa[:\s]*\s*(\d+)\s*[-\u2013\u2014]?\s*(.+)", texto, re.IGNORECASE)
    if m:
        codigo = m.group(1).strip()
        resto = m.group(2).strip()
        # cortar tudo após a primeira ocorrência de data, hora ou "Pág" ou "Page"
        corte = re.search(r"(\d{2}/\d{2}/\d{4}|\b\d{1,2}:\d{2}\b|\bPág\b|\bPág\.?\b|\bPage\b)", resto, re.IGNORECASE)
        if corte:
            nome_raw = resto[:corte.start()].strip()
        else:
            corte2 = re.search(r"(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\bInscrição\b|\bPeríodo\b)", resto, re.IGNORECASE)
            if corte2:
                nome_raw = resto[:corte2.start()].strip()
            else:
                nome_raw = resto
    else:
        # fallback simples: pega tudo após "Empresa:" se existir
        m2 = re.search(r"Empresa[:\s]*\s*(.+)", texto, re.IGNORECASE)
        if m2:
            codigo = ""
            nome_raw = m2.group(1).strip()
        else:
            codigo = ""
            nome_raw = ""
    nome_limpo = clean_company_name(nome_raw)
    return codigo, nome_limpo

# ---------------- parsing de linhas da tabela ----------------
def split_line_into_blocks(line: str):
    tokens = [t for t in line.strip().split() if t != ""]
    if not tokens:
        return []

    # achar índices de tokens que parecem dinheiro
    money_idxs = [i for i, t in enumerate(tokens) if is_money(t)]
    if not money_idxs:
        return [tokens]

    # consolidar sequências adjacentes de índices (por segurança)
    filtered_money_idxs = []
    i = 0
    while i < len(money_idxs):
        j = i
        while j + 1 < len(money_idxs) and money_idxs[j + 1] == money_idxs[j] + 1:
            j += 1
        filtered_money_idxs.append(money_idxs[j])
        i = j + 1

    blocks = []
    start = 0
    for mi in filtered_money_idxs:
        block = tokens[start:mi + 1]
        if block:
            blocks.append(block)
        start = mi + 1

    # se sobrar tokens finais, anexar ao último bloco (provavelmente descrição estendida)
    if start < len(tokens):
        if blocks:
            blocks[-1].extend(tokens[start:])
        else:
            blocks.append(tokens[start:])

    return blocks

def normalize_block_tokens(block_tokens):
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", ""]

    # encontrar índice do valor (último token que seja dinheiro)
    value_idx = None
    for i in range(len(toks) - 1, -1, -1):
        if is_money(toks[i]):
            value_idx = i
            break
    if value_idx is None:
        value_idx = len(toks) - 1

    value = toks[value_idx]

    # identificar possível token de horas entre colunas
    hour_idx = None
    for i in range(2, value_idx):
        t = toks[i].lower()
        if _token_hours_part.search(t) or t == "hs" or t == "0,00":
            hour_idx = i
            break

    col1 = toks[0] if len(toks) > 0 and not is_money(toks[0]) else ""
    col2 = toks[1] if len(toks) > 1 and not is_money(toks[1]) else ""

    start_desc = 2
    stop_desc = hour_idx if hour_idx is not None else value_idx
    if stop_desc < start_desc:
        stop_desc = start_desc

    desc_tokens = []
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            token = toks[i]
            lower = token.lower()
            if lower in ("hs", "h"):
                continue
            if _token_hours_part.search(token):
                continue
            if lower == "0,00":
                continue
            if is_money(token):
                continue
            desc_tokens.append(token)

    description = " ".join(desc_tokens).strip()
    return [col1 or "", col2 or "", description or "", value or ""]

# ---------------- extrair dados principais ----------------
def extrair_dados(texto):
    # extrair codigo e nome da empresa
    codigo_empresa, nome_empresa = extract_company_code_and_name(texto)

    cnpj_match = re.search(r"Inscrição Federal[:\s]*\s*([\d./-]+)", texto, re.IGNORECASE)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    periodo_match = re.search(r"Período[:\s]*\s*([0-3]?\d/[0-1]?\d/\d{4})\s*(?:a|-)\s*([0-3]?\d/[0-1]?\d/\d{4})", texto, re.IGNORECASE)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    # extrair trecho entre "Resumo Contrato" e "Totais"
    tabela_match = re.search(r"Resumo Contrato(.*?)(?:\nTotais\b|\nTotais\s*$)", texto, re.DOTALL | re.IGNORECASE)
    if not tabela_match:
        tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL | re.IGNORECASE)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else texto

    linhas = [ln.strip() for ln in tabela_texto.split("\n") if ln.strip()]

    output_rows = []
    debug_blocks = []
    for linha in linhas:
        tokens = [t for t in linha.split() if t]
        blocks = split_line_into_blocks(linha)
        normalized_for_line = []
        if not blocks:
            # linha sem blocos detectados: tentar normalizar a própria linha tokens
            normalized = normalize_block_tokens(tokens)
            normalized_for_line.append(normalized)
            output_rows.append(normalized)
        else:
            for b in blocks:
                normalized = normalize_block_tokens(b)
                normalized_for_line.append(normalized)
                output_rows.append(normalized)
        debug_blocks.append({
            "linha": linha,
            "tokens": tokens,
            "blocks": blocks,
            "normalized": normalized_for_line
        })

    df = pd.DataFrame(output_rows, columns=["Col1", "Col2", "Descrição", "Valor"])
    df = df.replace("", pd.NA).dropna(how="all").fillna("")

    tipo_map = {
        "1": "Proventos",
        "2": "Vantagens",
        "3": "Descontos",
        "4": "Informativo",
        "5": "Informativo"
    }
    df["Tipo"] = df["Col2"].map(tipo_map).fillna("")

    mes, ano = extrair_mes_ano(periodo)

    # adicionar colunas fixas incluindo Codigo Empresa (primeira coluna solicitada)
    df["Codigo Empresa"] = codigo_empresa
    df["Empresa"] = nome_empresa
    df["CNPJ"] = cnpj
    df["Período"] = periodo
    df["Mês"] = mes
    df["Ano"] = ano

    # renomear Col1 para Codigo da Descrição e reorganizar colunas:
    df = df.rename(columns={"Col1": "Codigo da Descrição"})
    df = df[["Codigo Empresa", "Empresa", "CNPJ", "Período", "Mês", "Ano", "Tipo", "Codigo da Descrição", "Descrição", "Valor"]]

    df["Valor_num"] = df["Valor"].apply(_to_float_br)

    valores_match = re.search(
        r"Proventos[:\s]*([\d\.,]+)\s*Vantagens[:\s]*([\d\.,]+)\s*Descontos[:\s]*([\d\.,]+)\s*Líquido[:\s]*([\d\.,]+)",
        texto, re.IGNORECASE
    )
    proventos = vantagens = descontos = liquido = ""
    if valores_match:
        proventos = valores_match.group(1)
        vantagens = valores_match.group(2)
        descontos = valores_match.group(3)
        liquido = valores_match.group(4)

    return {
        "codigo_empresa": codigo_empresa,
        "nome_empresa": nome_empresa,
        "cnpj": cnpj,
        "periodo": periodo,
        "tabela": df,
        "debug_blocks": debug_blocks,
        "proventos": proventos,
        "vantagens": vantagens,
        "descontos": descontos,
        "liquido": liquido
    }

def extrair_dados_csv(file):
    df_raw = pd.read_csv(
        file,
        sep=";",
        header=None,
        dtype=str,
        encoding="latin1"
    ).fillna("")

    texto_completo = "\n".join(
        df_raw.astype(str).agg(" ".join, axis=1).tolist()
    )

    

    # Sistema = última coluna da direita (linha 3)
    sistema = df_raw.iloc[2, -1].strip()
   # ================= EXTRAÇÃO ROBUSTA (CSV) =================

    codigo_empresa = ""
    nome_empresa = ""
    cnpj = ""
    periodo = ""
    
    for i in range(len(df_raw)):
        linha = " ".join(df_raw.iloc[i].astype(str).values)
    
        # Empresa (B3 visual)
        if "Empresa" in linha and "-" in linha and not codigo_empresa:
            partes = linha.split("-", 1)
            codigo_empresa = (
                partes[0]
                .replace("Empresa", "")
                .replace(":", "")
                .strip()
            )
            nome_empresa = clean_company_name(partes[1])
    
        # CNPJ (B4 visual)
        if "CNPJ" in linha or "Inscrição" in linha:
            m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", linha)
            if m:
                cnpj = m.group(1)
    
        # Período (A8 visual)
        if "Período" in linha:
            m = re.search(
                r"(\d{2}/\d{2}/\d{4}\s*a\s*\d{2}/\d{2}/\d{4})",
                linha
            )
            if m:
                periodo = m.group(1)
    
    
       
    

    mes, ano = extrair_mes_ano(periodo)

    # Encontrar Resumo Contrato → Totais
    start = None
    end = None

    
    for i in range(len(df_raw)):
        linha_txt = " ".join(df_raw.iloc[i].astype(str).values)
    
        if "Resumo Contrato" in linha_txt:
            start = i + 1
    
        if start and "Totais" in linha_txt:
            end = i
            break

    linhas = df_raw.iloc[start:end].values.tolist()

    rows = []

    num_colunas = len(df_raw.columns)

    for ln in linhas:
        for start in range(0, num_colunas, 5):
            try:
                codigo = str(ln[start]).strip()
                tipo = str(ln[start + 1]).strip()
                descricao = str(ln[start + 2]).strip()
                valor = str(ln[start + 4]).strip()
            except IndexError:
                continue
    
            # 🔒 FILTRO QUE DEFINE LINHA REAL
            if tipo not in {"1", "2", "3", "4", "5"}:
                continue
    
            if not codigo or not descricao or not valor:
                continue
    
            rows.append([
                codigo,
                tipo,
                descricao,
                valor
            ])



    df = pd.DataFrame(
        rows,
        columns=["Codigo da Descrição", "Tipo_raw", "Descrição", "Valor"]
    )

    tipo_map = {
        "1": "Proventos",
        "2": "Vantagens",
        "3": "Descontos",
        "4": "Informativo",
        "5": "Informativo"
    }

    df["Tipo"] = df["Tipo_raw"].map(tipo_map).fillna("")
    df["Valor_num"] = df["Valor"].apply(_to_float_br)

    df["Codigo Empresa"] = codigo_empresa
    df["Empresa"] = nome_empresa
    df["CNPJ"] = cnpj
    df["Período"] = periodo
    df["Mês"] = mes
    df["Ano"] = ano
    df["Sistema"] = sistema

    return df[
        ["Codigo Empresa", "Empresa", "CNPJ", "Período", "Mês", "Ano",
         "Tipo", "Codigo da Descrição", "Descrição", "Valor", "Valor_num", "Sistema"]
    ]

    



# ---------------- Streamlit UI ----------------

#st.title("📄 Extrator – Resumo Contrato")

uploaded_files = st.file_uploader(
    "Faça upload de PDFs ou CSVs",
    accept_multiple_files=True
)
#show_debug = st.checkbox("Mostrar debug (tokens & blocks)")
show_debug = False
if uploaded_files:
    all_dfs = []
    all_proventos = []
    all_vantagens = []
    all_descontos = []
    all_liquido = []


    for uploaded_file in uploaded_files:
        try:
            nome = uploaded_file.name.lower()
    
            # ================= PDF =================
            if nome.endswith(".pdf"):
                with pdfplumber.open(uploaded_file) as pdf:
                    texto = ""
                    for p in pdf.pages:
                        texto += (p.extract_text() or "") + "\n"
    
                dados = extrair_dados(texto)
                df = dados["tabela"].copy()
                df["Codigo Empresa"] = df["Codigo Empresa"].astype(str)
                all_dfs.append(df)
    
                all_proventos.append(dados["proventos"])
                all_vantagens.append(dados["vantagens"])
                all_descontos.append(dados["descontos"])
                all_liquido.append(dados["liquido"])
    
                if show_debug:
                    st.subheader(f"Debug do arquivo: {uploaded_file.name}")
                    st.markdown(
                        f"- Raw Empresa extraída: `{dados['codigo_empresa']} - {dados['nome_empresa']}`"
                    )
                    for i, dbg in enumerate(dados["debug_blocks"], start=1):
                        st.markdown(f"**Linha {i}:** {dbg['linha']}")
                        st.write("Tokens:", dbg["tokens"])
                        st.write("Blocks (tokens por bloco):", dbg["blocks"])
                        st.write("Normalized rows from this line:", dbg["normalized"])
                        st.markdown("---")
    
            # ================= CSV =================
            elif nome.endswith(".csv"):
                df_csv = extrair_dados_csv(uploaded_file)
                all_dfs.append(df_csv)
    
            else:
                st.warning(f"Arquivo ignorado: {uploaded_file.name}")
    
        except Exception as e:
            st.error(f"Erro ao processar o arquivo {uploaded_file.name}: {e}")
    
    # ================= CONCAT FINAL =================
    if all_dfs:
        df_all = pd.concat(all_dfs, ignore_index=True)
        

        # ---------------- Resumo por Codigo Empresa e Mês (fixo) ----------------
        st.subheader("Resumo Operação")

        df_resumo = df_all.copy()
        df_resumo['Mês'] = df_resumo['Mês'].astype(str)
        df_resumo['Tipo'] = df_resumo['Tipo'].astype(str)
        df_resumo['Codigo Empresa'] = df_resumo['Codigo Empresa'].astype(str)
        df_resumo['Valor_num'] = pd.to_numeric(df_resumo['Valor_num'], errors='coerce').fillna(0.0)

        # Agrupa por Codigo Empresa, Mês e Tipo
        resumo_agrupado = df_resumo.groupby(['Codigo Empresa', 'Mês', 'Tipo'])['Valor_num'].sum().reset_index()

        # Pivot para ter cada Tipo como coluna
        resumo_pivot = resumo_agrupado.pivot_table(index=['Codigo Empresa', 'Mês'], columns='Tipo', values='Valor_num', fill_value=0).reset_index()

        # Garantir colunas fixas na ordem desejada
        colunas_esperadas = ['Proventos', 'Vantagens', 'Descontos', 'Informativo']
        for col in colunas_esperadas:
            if col not in resumo_pivot.columns:
                resumo_pivot[col] = 0.0

        # Calcular Líquido = Proventos + Vantagens - Descontos
        resumo_pivot['Líquido'] = resumo_pivot['Proventos'] + resumo_pivot['Vantagens'] - resumo_pivot['Descontos']

        # Remover linhas onde Proventos, Vantagens, Descontos e Informativo são todos zero
        resumo_pivot = resumo_pivot[
            (resumo_pivot['Proventos'] != 0) |
            (resumo_pivot['Vantagens'] != 0) |
            (resumo_pivot['Descontos'] != 0) |
            (resumo_pivot['Informativo'] != 0)
        ].copy()

        # Ordenar por Codigo Empresa e por mês usando ordem dos meses PT
        month_order = {name: idx for idx, name in enumerate(_MONTHS_PT.values(), start=1)}
        resumo_pivot['_mes_ord'] = resumo_pivot['Mês'].map(lambda m: month_order.get(m, 999))
        resumo_pivot = resumo_pivot.sort_values(by=['Codigo Empresa', '_mes_ord', 'Mês']).drop(columns=['_mes_ord'])

        # Reordenar colunas para exibição e export
        colunas_final = ['Codigo Empresa', 'Mês'] + colunas_esperadas + ['Líquido']
        resumo_pivot = resumo_pivot[colunas_final]

        # Preparar versão formatada para exibição
        formatted = resumo_pivot.copy()
        for col in colunas_esperadas + ['Líquido']:
            formatted[col] = formatted[col].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.dataframe(
            formatted,
            use_container_width=True,
            height=min(360, 45 + 35 * len(formatted))
        )
        # Botão para baixar o resumo em Excel (com valores numéricos)
        out_summary = BytesIO()
        with pd.ExcelWriter(out_summary, engine="xlsxwriter") as writer:
            resumo_pivot.to_excel(writer, index=False, sheet_name="Resumo")
            ws = writer.sheets["Resumo"]
            # formatar colunas numéricas
            book = writer.book
            money_fmt = book.add_format({'num_format': '#,##0.00'})
            # localizar índices das colunas numéricas e aplicar formato
            for col_name in colunas_esperadas + ['Líquido']:
                try:
                    idx = resumo_pivot.columns.get_loc(col_name)
                    ws.set_column(idx, idx, 15, money_fmt)
                except Exception:
                    pass
            # ajustar larguras
            for i, col in enumerate(resumo_pivot.columns):
                max_len = max(resumo_pivot[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(i, i, max_len)
        out_summary.seek(0)

        st.download_button(
            label="📥 Excel",
            data=out_summary,
            file_name="resumo_por_empresa_mes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # ---------------- Tabela combinada (exibição e export) ----------------
        df_show = df_all.copy()
        df_show = df_show[df_show["Valor_num"].fillna(0) != 0]
        df_show["Valor"] = df_show["Valor_num"].apply(
            lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(v) else ""
        )
        

        st.subheader("Folha de Pagamento Detalhado")
        st.dataframe(
            df_show[["Codigo Empresa", "Empresa", "CNPJ", "Período", "Mês", "Ano", "Tipo", "Codigo da Descrição", "Descrição", "Valor"]],
            use_container_width=True,
            height=480
        )
        
        # Exportar para Excel com Valor numérico
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            export_df = df_all.copy()
            export_df = export_df.drop(columns=["Valor"]).rename(columns={"Valor_num": "Valor"})
            export_df.to_excel(writer, index=False, sheet_name="Resumo_Contrato")
            ws = writer.sheets["Resumo_Contrato"]
            last_col_idx = export_df.columns.get_loc("Valor")
            money_fmt = writer.book.add_format({'num_format': '#,##0.00'})
            ws.set_column(last_col_idx, last_col_idx, 15, money_fmt)
            for i, col in enumerate(export_df.columns):
                max_len = max(export_df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(i, i, max_len)
        output.seek(0)

        st.download_button(
            label="📥 Excel",
            data=output,
            file_name="resumo_contrato_combinado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # ---------------- Totais combinados (somatório dos campos extraídos dos arquivos) ----------------
        def parse_valor_str(v):
            try:
                return float(v.replace(".", "").replace(",", "."))
            except:
                return 0.0

        total_proventos = sum(parse_valor_str(v) for v in all_proventos if v)
        total_vantagens = sum(parse_valor_str(v) for v in all_vantagens if v)
        total_descontos = sum(parse_valor_str(v) for v in all_descontos if v)
        total_liquido = sum(parse_valor_str(v) for v in all_liquido if v)

       # st.subheader("Totais combinados dos PDFs (se extraídos das seções de totais)")
       # st.markdown(f"- **Proventos:** R$ {total_proventos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
       # st.markdown(f"- **Vantagens:** R$ {total_vantagens:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
       # st.markdown(f"- **Descontos:** R$ {total_descontos:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
       # st.markdown(f"- **Líquido:** R$ {total_liquido:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

else:
    st.info("Faça upload de um ou mais arquivos PDF para extrair as tabelas.")
