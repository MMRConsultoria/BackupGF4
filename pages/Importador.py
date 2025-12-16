import re
import pandas as pd

def extrair_dados(texto):
    # Extrair nome da empresa
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    # Extrair CNPJ
    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    # Extrair período
    periodo_match = re.search(r"Período:\s*([\d/]+)\s*a\s*([\d/]+)", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    # Extrair bloco da tabela entre "Resumo Contrato" e "Totais"
    tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else ""

    # Processar tabela em linhas e colunas
    linhas = [l.strip() for l in tabela_texto.split('\n') if l.strip()]
    dados = []
    for linha in linhas:
        # Dividir por múltiplos espaços (ajuste conforme o layout)
        cols = re.split(r'\s{2,}', linha)
        dados.append(cols)

    # Criar DataFrame da tabela
    df_tabela = pd.DataFrame(dados)

    # Extrair valores finais (Proventos, Vantagens, Descontos, Líquido)
    valores_match = re.search(
        r"Proventos:\s*([\d.,]+)\s*Vantagens:\s*([\d.,]+)\s*Descontos:\s*([\d.,]+)\s*Líquido:\s*([\d.,]+)",
        texto
    )
    proventos, vantagens, descontos, liquido = ("", "", "", "")
    if valores_match:
        proventos = valores_match.group(1)
        vantagens = valores_match.group(2)
        descontos = valores_match.group(3)
        liquido = valores_match.group(4)

    return {
        "nome_empresa": nome_empresa,
        "cnpj": cnpj,
        "periodo": periodo,
        "tabela": df_tabela,
        "proventos": proventos,
        "vantagens": vantagens,
        "descontos": descontos,
        "liquido": liquido
    }

# Exemplo de uso:
# texto_extraido = ... (seu texto extraído do PDF)
# dados = extrair_dados(texto_extraido)
# print(dados["nome_empresa"], dados["cnpj"], dados["periodo"])
# print(dados["tabela"])
# print(dados["proventos"], dados["vantagens"], dados["descontos"], dados["liquido"])
