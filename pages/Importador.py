import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# ======================================================
# UTILITÁRIOS
# ======================================================
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
    try:
        return float(str(x).replace(".", "").replace(",", "."))
    except:
        return None

_MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

def extrair_mes_ano(periodo_str):
    m = re.search(r"\d{2}/(\d{2})/(\d{4})", periodo_str or "")
    if m:
        return _MONTHS_PT.get(int(m.group(1)), ""), m.group(2)
    return "", ""

# ======================================================
# LIMPEZA EMPRESA (PDF + QUESTOR)
# ======================================================
def clean_company_name(raw):
    if not raw:
        return ""
    s = str(raw)

    # remover separadores do CSV Questor
    s = s.replace(";", " ")

    # remover datas
    s = re.sub(r'\d{2}/\d{2}/\d{4}', '', s)

    # remover paginação
    s = re.sub(r'\bPág.*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\bPágina.*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\bPage.*', '', s, flags=re.IGNORECASE)

    # normalizar espaços
    s = re.sub(r'\s{2,}', ' ', s)

    return s.strip()

def extract_company_code_and_name(texto):
    m = re.search(r"Empresa[:\s]*\s*(\d+)\s*[-–—]?\s*(.+)", texto, re.IGNORECASE)
    if not m:
        return "", ""
    return m.group(1), clean_company_name(m.group(2))

# ======================================================
# PARSER DE LINHAS
# ======================================================
def split_line_into_blocks(line):
    tokens = [t for t in line.split() if t]
    money_idxs = [i for i, t in enumerate(tokens) if is_money(t)]
    if not money_idxs:
        return [tokens]

    blocks, start = [], 0
    for idx in money_idxs:
        blocks.append(tokens[start:idx+1])
        start = idx + 1

    if start < len(tokens):
        blocks[-1].extend(tokens[start:])

    return blocks

def normalize_block_tokens(toks):
    value = next((t for t in reversed(toks) if is_money(t)), "")

    col1 = toks[0] if len(toks) > 0 else ""
    col2 = toks[1] if len(toks) > 1 else ""

    texto_full = " ".join(toks).lower()

    # ================== IDENTIFICAÇÃO DE TIPO (QUESTOR) ==================
    tipo = ""

    if any(p in texto_full for p in ["provento", "salário", "salario", "hora extra", "adicional"]):
        tipo = "Proventos"
    elif any(p in texto_full for p in ["vantagem", "benefício", "beneficio"]):
        tipo = "Vantagens"
    elif any(p in texto_full for p in ["desconto", "inss", "fgts", "irrf", "vale", "plano"]):
        tipo = "Descontos"
    elif any(p in texto_full for p in ["informativo", "base", "referência", "referencia"]):
        tipo = "Informativo"

    desc = " ".join(toks[2:-1]).strip()

    return [col1, col2 if not tipo else tipo, desc, value]

# ======================================================
# EXTRAÇÃO PRINCIPAL (COMUM)
# ======================================================
def extrair_dados(texto):
    codigo, empresa = extract_company_code_and_name(texto)

    periodo_match = re.search(
        r"Período[:\s]*(\d{2}/\d{2}/\d{4}.*?\d{4})",
        texto, re.IGNORECASE
    )
    periodo = periodo_match.group(1) if periodo_match else ""

    tabela_match = re.search(
        r"Resumo Contrato(.*?)(?:Proventos|Totais)",
        texto, re.DOTALL | re.IGNORECASE
    )

    linhas = tabela_match.group(1).splitlines() if tabela_match else []
    rows = []

    for ln in linhas:
        if not ln.strip():
            continue
        for b in split_line_into_blocks(ln):
            rows.append(normalize_block_tokens(b))

    df = pd.DataFrame(rows, columns=["Col1", "Col2", "Descrição", "Valor"])
    df["Valor_num"] = df["Valor"].apply(_to_float_br)

    # ================== TIPO PDF (fallback) ==================
    tipo_map = {
        "1": "Proventos",
        "2": "Vantagens",
        "3": "Descontos",
        "4": "Informativo",
        "5": "Informativo"
    }

    df["Tipo"] = df["Col2"].replace(tipo_map)
    df["Tipo"] = df["Tipo"].fillna(df["Col2"])

    mes, ano = extrair_mes_ano(periodo)

    df["Codigo Empresa"] = codigo
    df["Empresa"] = empresa
    df["Período"] = periodo
    df["Mês"] = mes
    df["Ano"] = ano

    df = df.rename(columns={"Col1": "Codigo da Descrição"})

    return df[
        ["Codigo Empresa", "Empresa", "Período", "Mês", "Ano",
         "Tipo", "Codigo da Descrição", "Descrição", "Valor", "Valor_num"]
    ]

# ======================================================
# LEITOR QUESTOR (CSV / XLSX COMO TEXTO)
# ======================================================
def ler_questor(uploaded):
    if uploaded.name.lower().endswith(".csv"):
        content = uploaded.read().decode("latin1")
    else:
        df = pd.read_excel(uploaded, header=None)
        content = "\n".join(df.astype(str).fillna("").agg(" ".join, axis=1))

    linhas = []
    for ln in content.splitlines():
        if re.search(r"Pág|Página|Page", ln, re.IGNORECASE):
            continue
        linhas.append(ln)

    return "\n".join(linhas)

# ======================================================
# STREAMLIT
# ======================================================
st.set_page_config("Extrator PDF + Questor", layout="wide")
st.title("📄 Extrator Resumo Contrato – PDF + Questor")

files = st.file_uploader(
    "Envie PDFs (Sistema Antigo) ou CSV/XLSX (Questor)",
    type=["pdf", "csv", "xlsx"],
    accept_multiple_files=True
)

if files:
    dfs = []

    for f in files:
        if f.name.lower().endswith(".pdf"):
            with pdfplumber.open(f) as pdf:
                texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
            df = extrair_dados(texto)
            df["Sistema"] = "Antigo"

        else:
            texto = ler_questor(f)
            df = extrair_dados(texto)
            df["Sistema"] = "Questor"

        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)

    # ================= VISUALIZAÇÃO =================
    st.subheader("Tabela combinada")
    df_show = df_all.copy()
    df_show["Valor"] = df_show["Valor_num"].apply(
        lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if pd.notna(x) else ""
    )

    st.dataframe(
        df_show[
            ["Sistema", "Codigo Empresa", "Empresa", "Período", "Mês", "Ano",
             "Tipo", "Codigo da Descrição", "Descrição", "Valor"]
        ],
        use_container_width=True,
        height=500
    )

    # ================= EXCEL =================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df = df_all.drop(columns=["Valor"]).rename(columns={"Valor_num": "Valor"})
        export_df.to_excel(writer, index=False, sheet_name="Resumo")
        ws = writer.sheets["Resumo"]
        money_fmt = writer.book.add_format({'num_format': '#,##0.00'})
        idx = export_df.columns.get_loc("Valor")
        ws.set_column(idx, idx, 15, money_fmt)

        for i, col in enumerate(export_df.columns):
            ws.set_column(i, i, max(len(col) + 2, 14))

    output.seek(0)

    st.download_button(
        "📥 Baixar Excel",
        data=output,
        file_name="resumo_contrato_unificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie os arquivos para processar.")
