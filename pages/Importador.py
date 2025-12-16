import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

# --- helpers de parsing ---
_money_re = re.compile(r'^\d{1,3}(?:\.\d{3})*,\d{2}$')          # exemplo: 101.662,53 ou 0,00
_hours_re = re.compile(r'\d+:\d+')                             # exemplo: 11459:20, 277:35

def is_money(tok: str) -> bool:
    tok = tok.strip()
    return bool(_money_re.match(tok))

def find_hours_token_index(tokens):
    # procura token com padrão horas (contém ':') ou token 'hs' / 'h'
    for i, t in enumerate(tokens):
        if _hours_re.search(t):
            return i
    for i, t in enumerate(tokens):
        if t.lower().endswith('hs') or t.lower() == 'hs' or t.lower().endswith('h'):
            return i
    return None

def normalize_block_tokens(block_tokens):
    """
    Recebe uma lista de tokens que correspondem a um bloco (até o valor monetário)
    e retorna uma lista com exatamente 5 campos: [Col1, Col2, Col3(desc), Col4(hours), Col5(value)]
    Preenche com "" quando faltar informação.
    """
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", "", ""]

    # valor monetário provavelmente é o último token que casa com money_re
    value = ""
    for i in range(len(toks)-1, -1, -1):
        if is_money(toks[i]):
            value = toks[i]
            end_idx = i
            break
    else:
        # se não encontrou valor, pega último token como value (fallback)
        value = toks[-1]
        end_idx = len(toks)-1

    # procurar índice do token de horas antes de end_idx
    hours_idx = None
    for i in range(end_idx-1, -1, -1):
        if _hours_re.search(toks[i]) or toks[i].lower().endswith('hs') or toks[i].lower() == 'hs':
            hours_idx = i
            break

    # col1 e col2: assumimos primeiros dois tokens se existirem
    col1 = toks[0] if len(toks) > 0 else ""
    col2 = toks[1] if len(toks) > 1 else ""

    # col4 = hours string (se hour token for separado em '11459:20' e 'hs', junta ambos)
    col4 = ""
    if hours_idx is not None:
        # junte tokens hours_idx até end_idx (exclusive) ou incluindo 'hs' se estiver depois
        h_parts = []
        i = hours_idx
        while i < end_idx and i < len(toks):
            # pare se chegamos ao token de value
            if is_money(toks[i]):
                break
            h_parts.append(toks[i])
            i += 1
        col4 = " ".join(h_parts).strip()
    else:
        # fallback: se não achou horas, mas existe token imediatamente antes do valor, pode ser horas
        if end_idx >= 1:
            maybe = toks[end_idx-1]
            if _hours_re.search(maybe) or maybe.lower().endswith('hs'):
                col4 = maybe

    # col3 = descrição: tokens entre col2 (índice 1) e hours_idx (ou até end_idx-1) juntados
    desc_tokens = []
    start_desc = 2  # depois de col1 e col2
    stop_desc = hours_idx if hours_idx is not None else end_idx
    if stop_desc < start_desc:
        stop_desc = start_desc
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            desc_tokens.append(toks[i])
    col3 = " ".join(desc_tokens).strip()

    # Garantir que todos existam
    result = [col1 or "", col2 or "", col3 or "", col4 or "", value or ""]
    return result

def split_line_into_blocks(line):
    """
    Tenta dividir uma linha em blocos. Primeiro tenta dividir por 2+ espaços.
    Se não houver separadores duplos, usa heurística por valores monetários para separar blocos.
    Retorna lista de blocos (cada bloco é lista de tokens).
    """
    # 1) split por >=2 espaços (bom quando pdf mantém colunas)
    parts = re.split(r'\s{2,}', line.strip())
    # se encontramos separadores suficientes (pelo menos 5 tokens para um bloco)
    if len(parts) >= 5:
        # parts pode já conter blocos inteiros (cada part pode ser uma célula)
        # reconstruir tokens: tratar 'parts' como tokens já
        tokens = [p for p in parts if p.strip() != ""]
        # construir blocos de 5 tokens sequenciais
        blocks = []
        i = 0
        while i < len(tokens):
            block = tokens[i:i+5]
            blocks.append(block)
            i += 5
        return blocks

    # 2) fallback: tokenizar por espaço simples e separar por ocorrência de valores monetários
    tokens = line.strip().split()
    blocks = []
    current = []
    for tok in tokens:
        current.append(tok)
        if is_money(tok):
            # finalize bloco no primeiro money token encontrado
            blocks.append(current)
            current = []
    # se sobrar algo, acrescentar (pode ser bloco incompleto)
    if current:
        blocks.append(current)

    return blocks

# --- função principal de extração do texto bruto para dados estruturados ---
def extrair_dados(texto):
    # nome empresa
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    # cnpj
    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    # período (capturar "dd/dd/dddd a dd/dd/dddd")
    periodo_match = re.search(r"Período:\s*([0-3]?\d/[0-1]?\d/\d{4})\s*a\s*([0-3]?\d/[0-1]?\d/\d{4})", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    # bloco entre "Resumo Contrato" e a próxima ocorrência de "Totais" (captura mais próxima)
    tabela_match = re.search(r"Resumo Contrato(.*?)(?:\nTotais\b|\nTotais\s*$)", texto, re.DOTALL | re.IGNORECASE)
    if not tabela_match:
        # tentativa alternativa: pegar entre "Resumo Contrato" e "Totais" sem \n
        tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL | re.IGNORECASE)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else ""

    # processar cada linha da tabela_texto
    linhas = [l for l in [ln.strip() for ln in tabela_texto.split('\n')] if l]
    output_rows = []
    for linha in linhas:
        blocks = split_line_into_blocks(linha)
        # cada block é lista de tokens; normalizar cada block para 5 colunas
        for b in blocks:
            normalized = normalize_block_tokens(b)
            output_rows.append(normalized)

    # criar DataFrame final com colunas fixas
    df = pd.DataFrame(output_rows, columns=["Col1", "Col2", "Col3", "Col4", "Col5"])

    # valores finais Proventos/Vantagens/Descontos/Líquido
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

# --- Streamlit UI ---
st.set_page_config(page_title="Extrair Resumo Contrato", layout="wide")
st.title("📄 Extrator - Resumo Contrato -> tabela desdobrada")

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

        st.subheader("Tabela - Resumo Contrato (desdobrada)")
        df_final = dados["tabela"]

        if df_final.empty:
            st.info("Nenhuma linha encontrada dentro do bloco 'Resumo Contrato'. Verifique o PDF ou o padrão do texto.")
        else:
            # Mostrar tabela com largura adequada
            st.dataframe(df_final, use_container_width=True)

            # permitir edição mínima (opcional)
            # edited = st.data_editor(df_final, use_container_width=True, height=300)

            # Exportar para Excel exatamente com as colunas Col1..Col5
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name='Resumo_Contrato')
                # Ajustar largura das colunas no excel (opcional)
                ws = writer.sheets['Resumo_Contrato']
                for i, col in enumerate(df_final.columns):
                    # largura baseada no maior conteúdo da coluna
                    max_len = max(df_final[col].astype(str).map(len).max(), len(col)) + 2
                    ws.set_column(i, i, max_len)
            output.seek(0)

            st.download_button(
                label="📥 Baixar tabela desdobrada em Excel",
                data=output,
                file_name="resumo_contrato_desdobrado.xlsx",
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
        # opcional: mostrar trecho do texto extraído para debugging
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
