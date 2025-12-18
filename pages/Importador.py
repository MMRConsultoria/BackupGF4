import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# ---------- regex / helpers ----------
_money_re = re.compile(r'^\d{1,3}(?:\.\d{3})*,\d{2}$')  # ex: 101.662,53 ou 0,00
_token_hours_part = re.compile(r'\d+:\d+')              # achar hh:mm em qualquer parte do token

def is_money(tok: str) -> bool:
    t = str(tok or "").strip()
    if not t:
        return False
    # aceitar formatos sem milhar também (ex: 598,02)
    if re.match(r'^\d+,\d{2}$', t):
        return True
    return bool(_money_re.match(t))

def _to_float_br(x):
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

# ---------- split line into blocks (robusto) ----------
def split_line_into_blocks(line: str):
    """
    Quebra a linha em blocos usando cada ocorrência de token monetário como final de bloco.
    Retorna lista de blocos (cada bloco = lista de tokens).
    """
    tokens = [t for t in line.strip().split() if t != ""]
    if not tokens:
        return []

    money_idxs = [i for i, t in enumerate(tokens) if is_money(t)]
    if not money_idxs:
        # sem money: retorna tudo como um bloco (será tratado depois)
        return [tokens]

    blocks = []
    start = 0
    for mi in money_idxs:
        # bloco do start até mi (inclusive)
        block = tokens[start:mi+1]
        if block:
            blocks.append(block)
        start = mi+1

    # se sobrou tokens após último money, anexar ao último bloco
    if start < len(tokens):
        if blocks:
            blocks[-1].extend(tokens[start:])
        else:
            blocks.append(tokens[start:])

    return blocks

# ---------- normalização de bloco -> 4 colunas ----------
def normalize_block_tokens(block_tokens):
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", ""]

    # Encontrar o índice do último token que é valor monetário (Valor)
    value_idx = None
    for i in range(len(toks)-1, -1, -1):
        if is_money(toks[i]):
            value_idx = i
            break
    if value_idx is None:
        value_idx = len(toks) - 1

    value = toks[value_idx]

    # Procurar índice do token de horas (ex: 11459:20, hs) ou do token '0,00' que aparece no lugar da hora
    hour_idx = None
    for i in range(2, value_idx):
        t = toks[i].lower()
        if _token_hours_part.search(t) or t == "hs" or t == "0,00":
            hour_idx = i
            break

    # Col1 e Col2 (somente se não forem valores monetários)
    col1 = toks[0] if len(toks) > 0 and not is_money(toks[0]) else ""
    col2 = toks[1] if len(toks) > 1 and not is_money(toks[1]) else ""

    # Descrição: tokens entre índice 2 e hour_idx (se hour_idx existir), senão até value_idx
    start_desc = 2
    stop_desc = hour_idx if hour_idx is not None else value_idx
    if stop_desc < start_desc:
        stop_desc = start_desc

    desc_tokens = []
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            desc_tokens.append(toks[i])
    description = " ".join(desc_tokens).strip()

    return [col1 or "", col2 or "", description or "", value or ""]

# ---------- extrair dados ----------
def extrair_dados(texto):
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    periodo_match = re.search(r"Período:\s*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    # bloco entre "Resumo Contrato" e "Totais"
    tabela_match = re.search(r"Resumo Contrato(.*?)(?:\nTotais\b|\nTotais\s*$)", texto, re.DOTALL | re.IGNORECASE)
    if not tabela_match:
        tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL | re.IGNORECASE)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else texto  # se não achar, processa todo texto (debug)

    linhas = [ln.strip() for ln in tabela_texto.split("\n") if ln.strip()]

    output_rows = []
    debug_blocks = []  # para debug opcional: (linha, tokens, blocks, normalized_rows)
    for linha in linhas:
        tokens = [t for t in linha.split() if t]
        blocks = split_line_into_blocks(linha)
        normalized_for_line = []
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
    # remover linhas vazias (todas colunas vazias)
    df = df.replace("", pd.NA).dropna(how="all").fillna("")

    df["Valor_num"] = df["Valor"].apply(_to_float_br)

    # Totais (Proventos/Vantagens/Descontos/Líquido)
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
        "debug_blocks": debug_blocks,
        "proventos": proventos,
        "vantagens": vantagens,
        "descontos": descontos,
        "liquido": liquido
    }

# ---------- Streamlit UI ----------
st.set_page_config(page_title="Extrair Resumo Contrato", layout="wide")
st.title("📄 Extrator - Resumo Contrato (4 colunas)")

uploaded_file = st.file_uploader("Faça upload do PDF (Relação de Cálculo)", type="pdf")
show_debug = st.checkbox("Mostrar debug (tokens & blocks)")

if uploaded_file:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            texto = ""
            for p in pdf.pages:
                texto += (p.extract_text() or "") + "\n"

        dados = extrair_dados(texto)

        st.subheader("Informações extraídas")
        st.markdown(f"**Nome da Empresa:** {dados['nome_empresa']}")
        st.markdown(f"**CNPJ:** {dados['cnpj']}")
        st.markdown(f"**Período:** {dados['periodo']}")

        df = dados["tabela"].copy()

        # Mostrar tabela com Valor formatado para visual
        df_show = df[["Col1", "Col2", "Descrição", "Valor_num"]].rename(columns={"Valor_num":"Valor"})
        df_show["Valor"] = df_show["Valor"].apply(lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(v) else "")

        st.subheader("Tabela - Resumo Contrato (formatada)")
        st.dataframe(df_show, use_container_width=True, height=420)

        # Botão download Excel (Valor numérico)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            export_df = df[["Col1", "Col2", "Descrição", "Valor_num"]].rename(columns={"Valor_num":"Valor"})
            export_df.to_excel(writer, index=False, sheet_name="Resumo_Contrato")
            ws = writer.sheets["Resumo_Contrato"]
            money_fmt = writer.book.add_format({'num_format': '#,##0.00'})
            last_col_idx = export_df.columns.get_loc("Valor")
            ws.set_column(last_col_idx, last_col_idx, 15, money_fmt)
            for i, col in enumerate(export_df.columns):
                max_len = max(export_df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(i, i, max_len)
        output.seek(0)

        st.download_button(
            label="📥 Baixar tabela (Excel) com Valor numérico",
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

        if show_debug:
            st.subheader("Debug por linha (tokens, blocos, normalizados)")
            for i, dbg in enumerate(dados["debug_blocks"], start=1):
                st.markdown(f"**Linha {i}:** {dbg['linha']}")
                st.write("Tokens:", dbg["tokens"])
                st.write("Blocks (tokens por bloco):", dbg["blocks"])
                st.write("Normalized rows from this line:", dbg["normalized"])
                st.markdown("---")

    except Exception as e:
        st.error(f"Erro ao processar o PDF: {e}")
        # mostrar preview do texto para ajudar debug
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                preview = ""
                for i, p in enumerate(pdf.pages[:4]):
                    preview += f"--- Página {i+1} ---\n"
                    preview += (p.extract_text() or "") + "\n\n"
            st.text_area("Preview texto extraído (debug)", preview, height=300)
        except Exception:
            pass
else:
    st.info("Faça upload do PDF para extrair a tabela.")
