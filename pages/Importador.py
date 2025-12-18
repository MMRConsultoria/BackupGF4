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
        # se houver pontos de milhar e vírgula decimal
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

# Mapeamento mês em português (evita problemas de locale)
_MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

def extrair_mes_ano(periodo_str):
    """
    Recebe período no formato 'dd/mm/aaaa a dd/mm/aaaa' e retorna (MêsNome, Ano)
    Usa o mês da primeira data do período.
    """
    match = re.search(r"(\d{2})/(\d{2})/(\d{4})", periodo_str)
    if match:
        mes_num = int(match.group(2))
        ano = match.group(3)
        mes_nome = _MONTHS_PT.get(mes_num, "")
        return mes_nome, ano
    return "", ""

# ---------- split line into blocks ----------
def split_line_into_blocks(line: str):
    """
    Quebra a linha em blocos usando cada ocorrência de token monetário (is_money) como final de bloco.
    Se houver sequência de money tokens consecutivos (ex: '0,00 1.847,81'), usa o ÚLTIMO money da sequência
    como final do bloco (evita criar um bloco separado com 0,00 como 'valor').
    """
    tokens = [t for t in line.strip().split() if t != ""]
    if not tokens:
        return []

    money_idxs = [i for i, t in enumerate(tokens) if is_money(t)]
    if not money_idxs:
        return [tokens]

    # agrupa runs de índices consecutivos e pega o último índice de cada run
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

    # se restaram tokens após último money, anexar ao último bloco
    if start < len(tokens):
        if blocks:
            blocks[-1].extend(tokens[start:])
        else:
            blocks.append(tokens[start:])

    return blocks

# ---------- normalize block tokens ----------
def normalize_block_tokens(block_tokens):
    """
    Converte um bloco de tokens em [Col1, Col2, Descrição, Valor] aplicando as regras:
    - Valor = último token money do bloco
    - Ignorar tokens de horas (hh:mm, 'hs') e também '0,00' quando aparecer no lugar da hora
    - Descrição = tokens entre Col2 e início das horas (ou até o valor se não houver horas)
    - Col1 e Col2 só são preenchidos se não forem money (proteção contra deslocamentos)
    """
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", ""]

    # encontrar último token money no bloco
    value_idx = None
    for i in range(len(toks) - 1, -1, -1):
        if is_money(toks[i]):
            value_idx = i
            break
    if value_idx is None:
        value_idx = len(toks) - 1

    value = toks[value_idx]

    # detectar token de horas ou placeholder 0,00 entre índice 2 e value_idx-1
    hour_idx = None
    for i in range(2, value_idx):
        t = toks[i].lower()
        if _token_hours_part.search(t) or t == "hs" or t == "0,00":
            hour_idx = i
            break

    # Col1 = código da descrição; Col2 = código numérico (para mapear Tipo)
    col1 = toks[0] if len(toks) > 0 and not is_money(toks[0]) else ""
    col2 = toks[1] if len(toks) > 1 and not is_money(toks[1]) else ""

    # descrição: entre índice 2 e hour_idx (se existir) ou até value_idx
    start_desc = 2
    stop_desc = hour_idx if hour_idx is not None else value_idx
    if stop_desc < start_desc:
        stop_desc = start_desc

    desc_tokens = []
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            token = toks[i]
            lower = token.lower()
            # ignorar 'hs', formatos hh:mm e '0,00'
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

# ---------- extrair dados do texto ----------
def extrair_dados(texto):
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    periodo_match = re.search(r"Período:\s*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    # captura o bloco entre "Resumo Contrato" e "Totais"
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

    # Mapear Col2 (código numérico) para nome do tipo conforme mapeamento
    tipo_map = {
        "1": "Proventos",
        "2": "Vantagens",
        "3": "Descontos",
        "4": "Informativo",
        "5": "Informativo"
    }
    df["Tipo"] = df["Col2"].map(tipo_map).fillna("")

    # extrair mês e ano do período
    mes, ano = extrair_mes_ano(periodo)

    # adicionar colunas fixas
    df["Empresa"] = nome_empresa
    df["CNPJ"] = cnpj
    df["Período"] = periodo
    df["Mês"] = mes
    df["Ano"] = ano

    # renomear Col1 para Codigo da Descrição e reorganizar colunas:
    df = df.rename(columns={"Col1": "Codigo da Descrição"})
    df = df[["Empresa", "CNPJ", "Período", "Mês", "Ano", "Tipo", "Codigo da Descrição", "Descrição", "Valor"]]

    # converter Valor para numérico em nova coluna
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
st.title("📄 Extrator - Resumo Contrato (formato final)")

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

        # Preparar exibição: formatar Valor_num para exibir como BR
        df_show = df.copy()
        df_show["Valor"] = df_show["Valor_num"].apply(
            lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(v) else ""
        )

        # Exibir com a ordem solicitada e coluna "Codigo da Descrição" ao lado da Descrição
        st.subheader("Tabela - Resumo Contrato (formatada)")
        st.dataframe(
            df_show[["Empresa", "CNPJ", "Período", "Mês", "Ano", "Tipo", "Codigo da Descrição", "Descrição", "Valor"]],
            use_container_width=True,
            height=480
        )

        # Exportar para Excel com Valor numérico
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            export_df = df.copy()
            # renomear Valor_num -> Valor (numérico) para o Excel e remover string Valor
            export_df = export_df.drop(columns=["Valor"]).rename(columns={"Valor_num": "Valor"})
            export_df.to_excel(writer, index=False, sheet_name="Resumo_Contrato")
            ws = writer.sheets["Resumo_Contrato"]
            # formatar coluna Valor
            last_col_idx = export_df.columns.get_loc("Valor")
            money_fmt = writer.book.add_format({'num_format': '#,##0.00'})
            ws.set_column(last_col_idx, last_col_idx, 15, money_fmt)
            # ajustar largura das colunas
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
        # mostrar preview do texto extraído para ajudar o debug
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
