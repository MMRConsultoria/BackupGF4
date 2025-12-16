import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# -------- regex helpers --------
_money_re = re.compile(r'^\d{1,3}(?:\.\d{3})*,\d{2}$')  # ex: 101.662,53 ou 0,00

def is_money(tok: str) -> bool:
    t = str(tok or "").strip()
    if not t:
        return False
    if re.match(r'^\d+,\d{2}$', t):
        return True
    return bool(_money_re.match(t))

def _to_float_br(x):
    """Converte string BR '101.662,53' -> float 101662.53"""
    t = str(x or "").strip()
    if not t:
        return None
    t = t.replace(" ", "")
    has_c = "," in t
    has_p = "." in t
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

# -------- parsing helpers --------
def split_line_into_blocks(line: str):
    """
    Tenta dividir a linha em blocos. Primeiro tenta split por >=2 espaços.
    Se não, divide por espaço e fecha bloco ao encontrar money token.
    Retorna: lista de blocos (cada bloco = lista de tokens).
    """
    parts = re.split(r'\s{2,}', line.strip())
    tokens = [p for p in parts if p.strip() != ""]
    if len(tokens) >= 5:
        blocks = []
        i = 0
        while i < len(tokens):
            blocks.append(tokens[i:i+5])
            i += 5
        return blocks

    tokens = line.strip().split()
    blocks = []
    current = []
    for tok in tokens:
        current.append(tok)
        if is_money(tok):
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    return blocks

def normalize_block_tokens(block_tokens):
    """
    Retorna [Col1, Col2, Descrição, Valor]
    Regras:
    - Valor = último token money.
    - Se token antes do valor for '0,00', ignorar ele na descrição.
    - Descrição = tokens entre Col2 e valor (ou token antes do valor se for '0,00').
    """
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", ""]

    # Encontrar último token money -> Valor
    value = ""
    value_idx = None
    for i in range(len(toks)-1, -1, -1):
        if is_money(toks[i]):
            value = toks[i]
            value_idx = i
            break
    if value_idx is None:
        value_idx = len(toks) - 1
        value = toks[value_idx]

    prev_idx = value_idx - 1

    drop_prev_zero = False
    if prev_idx >= 0 and toks[prev_idx] == "0,00":
        drop_prev_zero = True

    start_desc = 2
    stop_desc = prev_idx if drop_prev_zero else value_idx
    if stop_desc < start_desc:
        stop_desc = start_desc

    desc_tokens = []
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            desc_tokens.append(toks[i])
    description = " ".join(desc_tokens).strip()

    col1 = toks[0] if len(toks) > 0 else ""
    col2 = toks[1] if len(toks) > 1 else ""

    return [col1 or "", col2 or "", description or "", value or ""]

# -------- extrair dados do texto --------
def extrair_dados(texto):
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    periodo_match = re.search(r"Período:\s*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    tabela_match = re.search(r"Resumo Contrato(.*?)(?:\nTotais\b|\nTotais\s*$)", texto, re.DOTALL | re.IGNORECASE)
    if not tabela_match:
        tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL | re.IGNORECASE)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else ""

    linhas = [l for l in [ln.strip() for ln in tabela_texto.split('\n')] if l]
    output_rows = []
    for linha in linhas:
        blocks = split_line_into_blocks(linha)
        for b in blocks:
            normalized = normalize_block_tokens(b)
            output_rows.append(normalized)

    df = pd.DataFrame(output_rows, columns=["Col1", "Col2", "Descrição", "Valor"])

    # converter Valor para float
    df["Valor_num"] = df["Valor"].apply(_to_float_br)

    # valores finais
    valores_match = re.search(
        r"Proventos:\s*([\d\.,]+)\s*Vantagens:\s*([\d\.,]+)\s*Descontos:\s*([\d\.,]+)\s*Líquido:\s*([\d\.,]+)",
        texto, re.IGNORECASE
    )
    proventos = vantagens = descontos = liquido = ""
    if valores_match:
        proventos = valores_match.group(1)
        vantagens = valores_match.group(2)
        descontos = valores_match.group(3)
        liquido = valores_match.group(4)

    return {
        "nome_empresa": nome_empresa,
        "cnpj": cnpj,
        "periodo": periodo,
        "tabela": df,
        "proventos": proventos,
        "vantagens": vantagens,
        "descontos": descontos,
        "liquido": liquido
    }

# -------- Streamlit UI --------
st.set_page_config(page_title="Extrair Resumo Contrato", layout="wide")
st.title("📄 Extrator - Resumo Contrato (4 colunas)")

uploaded_file = st.file_uploader("Faça upload do PDF (Relação de Cálculo)", type="pdf")

if uploaded_file:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                txt = page.extract_text() or ""
                texto_completo += txt + "\n"

        dados = extrair_dados(texto_completo)

        st.subheader("Informações extraídas")
        st.markdown(f"**Nome da Empresa:** {dados['nome_empresa']}")
        st.markdown(f"**CNPJ:** {dados['cnpj']}")
        st.markdown(f"**Período:** {dados['periodo']}")

        st.subheader("Tabela - Resumo Contrato (formatada)")
        df_final = dados["tabela"].copy()

        # mostrar com coluna Valor_num formatada no padrão BR
        df_show = df_final[["Col1", "Col2", "Descrição", "Valor_num"]].rename(columns={"Valor_num":"Valor"})
        df_show["Valor"] = df_show["Valor"].apply(lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(v) else "")

        st.dataframe(df_show, use_container_width=True)

        # Exportar para Excel: manter Col1,Col2,Descrição,Valor (numérico)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            export_df = df_final[["Col1", "Col2", "Descrição", "Valor_num"]].rename(columns={"Valor_num":"Valor"})
            export_df.to_excel(writer, index=False, sheet_name='Resumo_Contrato')
            ws = writer.sheets['Resumo_Contrato']
            money_fmt = writer.book.add_format({'num_format': '#,##0.00'})
            last_col_idx = export_df.columns.get_loc("Valor")
            ws.set_column(last_col_idx, last_col_idx, 15, money_fmt)
            for i, col in enumerate(export_df.columns):
                max_len = max(export_df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(i, i, max_len)
        output.seek(0)

        st.download_button(
            label="📥 Baixar tabela (Excel) com coluna Valor numérica",
            data=output,
            file_name="resumo_contrato_formatado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.subheader("Totais")
        st.markdown(f"- **Proventos:** {dados['proventos']}")
        st.markdown(f"- **Vantagens:** {dados['vantagens']}")
        st.markdown(f"- **Descontos:** {dados['descontos']}")
        st.markdown(f"- **Líquido:** {dados['liquido']}")

    except Exception as e:
        st.error(f"Erro ao processar o PDF: {e}")
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                debug_text = ""
                for i, page in enumerate(pdf.pages[:3]):
                    debug_text += f"--- Página {i+1} ---\n"
                    debug_text += (page.extract_text() or "") + "\n\n"
            st.text_area("Texto (preview) para debug", debug_text, height=300)
        except Exception:
            pass
else:
    st.info("Faça upload do PDF para extrair a tabela.")
