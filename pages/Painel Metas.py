# pages/Painel Metas.py

import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
import io
from io import BytesIO
import openpyxl
from st_aggrid import AgGrid, GridOptionsBuilder

if not st.session_state.get("acesso_liberado"):
    st.stop()

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# ========================
# 🔐 Autenticação Google Sheets
# ========================
#from google.oauth2.service_account import Credentials
#import gspread
#import streamlit as st

#scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

#credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]  # já vem como dict
#credentials = Credentials.from_service_account_info(credentials_dict, scopes=scope)
#gc = gspread.authorize(credentials)

# Sempre use open_by_key (mais seguro que open pelo nome)
planilha_empresa = gc.open_by_key("1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU")



df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# Padronizar Tabela Empresa
for col in ["Loja", "Grupo", "Tipo"]:
    df_empresa[col] = df_empresa[col].astype(str).str.strip().str.upper()

# ================================
# 2. Estilo e layout
# ================================
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
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
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: black; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Metas Mensais</h1>
    </div>
""", unsafe_allow_html=True)

# ================================
# Função auxiliar para converter valores
# ================================
def parse_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except:
        return 0.0

# Função auxiliar para tratar datas misturadas

def tratar_data(val):
    try:
        val_str = str(val).strip()
        if val_str.replace('.', '', 1).isdigit():
            return pd.Timestamp('1899-12-30') + pd.to_timedelta(float(val), unit='D')
        else:
            return pd.to_datetime(val_str, dayfirst=True, errors='coerce')
    except:
        return pd.NaT

def garantir_escalar(x):
    if isinstance(x, list):
        if len(x) == 1:
            return x[0]
        return str(x)
    return x

def formatar_moeda_br(val):
    try:
        if val in [None, np.nan, "None", "nan", "NaN", ""]:
            return ""
        val_float = float(val)
        return f"R$ {val_float:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    except:
        return ""



# ================================
# Abas
# ================================
aba1, aba2, aba3 = st.tabs(["📈 Analise Metas", "📊 Auditoria Metas","📥Importador"])



#===========================================
with aba1:
#===========================================
    # --- Metas ---
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())
    df_metas["Meta"] = df_metas["Meta"].apply(parse_valor)
    df_metas["Loja"] = df_metas["Loja Vendas"].astype(str).str.strip().str.upper()
    df_metas["Grupo"] = df_metas["Grupo"].astype(str).str.strip().str.upper()
    df_metas = df_metas[df_metas["Loja"] != ""]

    # --- Realizado ---
    df_anos = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())
    df_anos.columns = df_anos.columns.str.strip()
    df_anos["Loja"] = df_anos["Loja"].astype(str).str.strip().str.upper()
    df_anos["Grupo"] = df_anos["Grupo"].astype(str).str.strip().str.upper()
    df_anos["Data"] = df_anos["Data"].apply(tratar_data)
    df_anos = df_anos.dropna(subset=["Data"])

  

    meses_map = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
    df_anos["Mês"] = df_anos["Data"].dt.month.map(meses_map)
    df_anos["Ano"] = df_anos["Data"].dt.year
    df_anos["Fat.Total"] = df_anos["Fat.Total"].apply(parse_valor)
    
    df_anos = df_anos.merge(df_empresa[["Loja", "Grupo", "Tipo"]], on=["Loja", "Grupo"], how="left")
    ano_atual = datetime.now().year
    from datetime import datetime

    # Mapeamento de meses (sistema → abreviações)
    mapa_meses = {
        "janeiro": "Jan", "fevereiro": "Fev", "março": "Mar", "abril": "Abr",
        "maio": "Mai", "junho": "Jun", "julho": "Jul", "agosto": "Ago",
        "setembro": "Set", "outubro": "Out", "novembro": "Nov", "dezembro": "Dez"
    }
    
    # Lista de meses abreviados para o filtro
    ordem_meses = list(mapa_meses.values())
    
    # Mês atual do sistema (em português) e mapeado
    mes_extenso = datetime.now().strftime("%B").lower()  # ex: 'agosto'
    mes_atual = mapa_meses.get(mes_extenso, "Jan")
    
    # Evita erro caso algo dê errado
    index_mes = ordem_meses.index(mes_atual) if mes_atual in ordem_meses else 0

    anos_disponiveis = sorted(df_anos["Ano"].unique())

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        ano_selecionado = st.selectbox("Ano:", anos_disponiveis, index=anos_disponiveis.index(ano_atual))
    with col2:
        mes_selecionado = st.selectbox("Mês:", ordem_meses, index=index_mes)
    with col3:
        modo_visao = st.selectbox("Visão:", ["Por Loja", "Por Grupo"])
    with col4:
        tipos_disponiveis = sorted(df_anos["Tipo"].dropna().unique())
        tipo_selecionado = st.selectbox("Tipo:", ["TODOS"] + tipos_disponiveis)
        df_anos_filtrado = df_anos[(df_anos["Ano"] == ano_selecionado) & (df_anos["Mês"] == mes_selecionado)].copy()

    if tipo_selecionado != "TODOS":
        df_anos_filtrado = df_anos_filtrado[df_anos_filtrado["Tipo"] == tipo_selecionado]
    
    df_metas_filtrado = df_metas[(df_metas["Ano"] == ano_selecionado) & (df_metas["Mês"] == mes_selecionado)].copy()

    import calendar
    ultima_data_realizado_dt = df_anos_filtrado["Data"].max()
    
    if pd.isna(ultima_data_realizado_dt):
        ultima_data_realizado = "N/A"
        percentual_meta_desejavel = 0
    else:
        ultima_data_realizado = ultima_data_realizado_dt.strftime("%d/%m/%Y")
        dias_do_mes = calendar.monthrange(ano_selecionado, ordem_meses.index(mes_selecionado) + 1)[1]
        dias_corridos = ultima_data_realizado_dt.day
        percentual_meta_desejavel = dias_corridos / dias_do_mes




    

    for col in ["Ano", "Mês", "Loja", "Grupo"]:
        df_metas_filtrado[col] = df_metas_filtrado[col].apply(garantir_escalar)
        df_anos_filtrado[col] = df_anos_filtrado[col].apply(garantir_escalar)

    metas_grouped = df_metas_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo"])["Meta"].sum().reset_index().rename(columns={"Meta": "Meta"})
    realizado_grouped = df_anos_filtrado.groupby(["Ano", "Mês", "Loja", "Grupo", "Tipo"])["Fat.Total"].sum().reset_index().rename(columns={"Fat.Total": "Realizado"})

    comparativo = pd.merge(metas_grouped, realizado_grouped, on=["Ano", "Mês", "Loja", "Grupo"], how="outer").fillna(0)
    comparativo["% Atingido"] = np.where(comparativo["Meta"] == 0, 0, comparativo["Realizado"] / comparativo["Meta"])
    comparativo["Diferença"] = comparativo["Realizado"] - comparativo["Meta"]
    comparativo["% Falta Atingir"] = np.maximum(0, 1 - comparativo["% Atingido"])
    comparativo["Mês"] = pd.Categorical(comparativo["Mês"], categories=ordem_meses, ordered=True)

    # 🔥 AQUI: garante que o filtro por tipo seja respeitado
    if tipo_selecionado != "TODOS":
        comparativo = comparativo[comparativo["Tipo"] == tipo_selecionado]
    

    tipo_subtotais = []
    for tipo, dados_tipo in comparativo.groupby("Tipo"):
        soma_meta_tipo = dados_tipo["Meta"].sum()
        soma_realizado_tipo = dados_tipo["Realizado"].sum()
        soma_diferenca_tipo = dados_tipo["Diferença"].sum()
        perc_atingido_tipo = soma_realizado_tipo / soma_meta_tipo if soma_meta_tipo != 0 else 0
        perc_falta_tipo = max(0, 1 - perc_atingido_tipo)
        qtde_lojas_tipo = dados_tipo["Loja"].nunique()

        linha_tipo = pd.DataFrame({
            "Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"{tipo} - Lojas: {qtde_lojas_tipo:02}"],
            "Meta": [soma_meta_tipo], "Realizado": [soma_realizado_tipo], "% Atingido": [perc_atingido_tipo], "% Falta Atingir": [perc_falta_tipo], "Diferença": [soma_diferenca_tipo], "Tipo": [tipo]
        })
        tipo_subtotais.append(linha_tipo)

    # (até aqui seu código é igual, vamos direto no ponto de alteração)
    
    # ✅ Aqui começa o bloco do resultado_final com a ordenação por Tipo
    resultado_final = []
    total_lojas_geral = comparativo["Loja"].nunique()
    
    # Primeiro, criamos uma lista auxiliar com os subtotais incluindo o tipo já capturado
    subtotais_aux = []
    for grupo, dados_grupo in comparativo.groupby("Grupo"):
        soma_meta_grupo = dados_grupo["Meta"].sum()
        soma_realizado_grupo = dados_grupo["Realizado"].sum()
        soma_diferenca_grupo = dados_grupo["Diferença"].sum()
        perc_atingido_grupo = soma_realizado_grupo / soma_meta_grupo if soma_meta_grupo != 0 else 0
        perc_falta_grupo = max(0, 1 - perc_atingido_grupo)
        qtde_lojas_grupo = dados_grupo["Loja"].nunique()
        
        tipo_grupo = df_empresa[df_empresa["Grupo"] == grupo]["Tipo"].dropna().unique()
        tipo_str = tipo_grupo[0] if len(tipo_grupo) == 1 else "OUTROS"
    
        subtotais_aux.append({
            "grupo": grupo,
            "tipo": tipo_str,
            "qtde_lojas": qtde_lojas_grupo,
            "meta": soma_meta_grupo,
            "realizado": soma_realizado_grupo,
            "diferenca": soma_diferenca_grupo,
            "perc_atingido": perc_atingido_grupo,
            "perc_falta": perc_falta_grupo
        })
    
    # Definimos a ordem de prioridade dos tipos
    ordem_tipo = {"AIRPORTS": 1, "AIRPORTS - KOPP": 2, "ON-PREMISE": 3, "OUTROS": 4}


    # Normaliza o tipo (garantir comparação segura)
    for item in subtotais_aux:
        item["tipo"] = str(item["tipo"]).strip().upper()

    
    # Ordenamos os grupos com base no tipo
    subtotais_aux = sorted(subtotais_aux, key=lambda x: (ordem_tipo.get(x["tipo"], 99), x["grupo"]))
    
    # Agora, com os grupos já ordenados, montamos o resultado final
    for subtotal in subtotais_aux:
        grupo = subtotal["grupo"]
        tipo_str = subtotal["tipo"]
        qtde_lojas_grupo = subtotal["qtde_lojas"]
    
        dados_grupo = comparativo[comparativo["Grupo"] == grupo]

        # pega o Ano e Mês do grupo para inserir no subtotal
        ano_grupo = dados_grupo["Ano"].iloc[0] if not dados_grupo.empty else ""
        mes_grupo = dados_grupo["Mês"].iloc[0] if not dados_grupo.empty else ""
        
        
        resultado_final.append(dados_grupo)
    
        linha_subtotal = pd.DataFrame({
            "Ano": [ano_grupo], "Mês": [mes_grupo], "Grupo": [grupo],
            "Loja": [f"{grupo} - Lojas: {qtde_lojas_grupo:02}"],
            "Meta": [subtotal["meta"]], "Realizado": [subtotal["realizado"]],
            "% Atingido": [subtotal["perc_atingido"]], "% Falta Atingir": [subtotal["perc_falta"]],
            "Diferença": [subtotal["diferenca"]], "Tipo": [tipo_str]
        })
        resultado_final.append(linha_subtotal)
    
    # ✅ Total Geral continua exatamente igual ao seu
    total_meta = comparativo["Meta"].sum()
    total_realizado = comparativo["Realizado"].sum()
    total_diferenca = comparativo["Diferença"].sum()
    percentual_total = total_realizado / total_meta if total_meta != 0 else 0
    percentual_falta_total = max(0, 1 - percentual_total)
    
    linha_total = pd.DataFrame({
        "Ano": [""], "Mês": [""], "Grupo": [""], "Loja": [f"TOTAL GERAL - Lojas: {total_lojas_geral:02}"],
        "Meta": [total_meta], "Realizado": [total_realizado],
        "% Atingido": [percentual_total], "% Falta Atingir": [percentual_falta_total],
        "Diferença": [total_diferenca], "Tipo": [""]
    })

    # Cria a linha única da FATURAMENTO DESEJÁVEL
    linha_meta_desejavel = pd.DataFrame({
        "Ano": [""], 
        "Mês": [""], 
        "Grupo": [""],
        "Loja": [f"FATURAMENTO DESEJÁVEL ATÉ {ultima_data_realizado}"],
        "Meta": [np.nan],
        "Realizado": [np.nan],
        "% Atingido": [percentual_meta_desejavel],
        "% Falta Atingir": [np.nan],
        "Diferença": [np.nan],
        "Tipo": [""]
    })











    
    # ✅ Monta o comparativo final preservando o seu restante
    comparativo_final = pd.concat([linha_meta_desejavel] + tipo_subtotais + [linha_total] + resultado_final, ignore_index=True)
    # ✅ Ajusta o nome da coluna "Realizado"
    comparativo_final.rename(columns={"Realizado": f"Realizado até {ultima_data_realizado}"}, inplace=True)

    # Ajusta os tipos para garantir float e evitar None no Styler
    for col in ["Meta", f"Realizado até {ultima_data_realizado}", "Diferença", "% Atingido", "% Falta Atingir"]:
        comparativo_final[col] = pd.to_numeric(comparativo_final[col], errors='coerce')



    # Alternância de cor por grupo (bloco)
    grupos_unicos = comparativo_final["Grupo"].dropna().unique()
    cores_blocos = ["#fff9f0", "#f9fbf0"]
    mapa_cor_por_grupo = {
        grupo: cores_blocos[i % len(cores_blocos)]
        for i, grupo in enumerate(sorted(grupos_unicos))
    }

    
    
    
    def formatar_linha(row):
        estilo = []
        loja_valor = str(row["Loja"])
        loja_valor_upper = loja_valor.upper()
        atingido = row["% Atingido"]
    
        if "FATURAMENTO DESEJÁVEL" in loja_valor_upper:
            fundo = "#FF6666"  # vermelho clarinho
            meta_desejavel_linha = True
        elif "TOTAL GERAL" in loja_valor_upper:
            fundo = "#99c7f3"
            meta_desejavel_linha = False
        elif "- LOJAS:" in loja_valor_upper and not row["Grupo"]:
            fundo = "#f9f9f9"
            meta_desejavel_linha = False
        elif "LOJAS:" in loja_valor_upper:
            fundo = "#f2f2f2"
            meta_desejavel_linha = False
        else:
            fundo = "white"
            meta_desejavel_linha = False
    
        for coluna in row.index:
            if coluna == "% Atingido" and not pd.isna(atingido):
                if meta_desejavel_linha:
                    # para linha FATURAMENTO DESEJÁVEL sempre fonte preta negrito
                    estilo.append(f"background-color: {fundo}; color: black; font-weight: bold; font-size: 1.1em;")
                elif atingido >= percentual_meta_desejavel:
                    estilo.append(f"background-color: {fundo}; color: green; font-weight: bold; font-size: 1.1em;")
                else:
                    estilo.append(f"background-color: {fundo}; color: red; font-weight: bold; font-size: 1.1em;")
            else:
                if meta_desejavel_linha:
                    estilo.append(f"background-color: {fundo}; color: black; font-weight: bold;")
                else:
                    estilo.append(f"background-color: {fundo}; color: black;")
        return estilo


    # ✅ Exibe a data de realizado antes da tabela
    st.markdown(f"**Última data realizada:** {ultima_data_realizado}")
    
    
    
    if modo_visao == "Por Grupo":
        dados_exibir = comparativo_final[
            comparativo_final["Loja"].astype(str).str.contains("Lojas:") |
            comparativo_final["Loja"].astype(str).str.contains("FATURAMENTO DESEJÁVEL")
        ]
    else:
        dados_exibir = comparativo_final.copy()

    st.dataframe(
        dados_exibir.style
            .format({
                "Meta": formatar_moeda_br, 
                f"Realizado até {ultima_data_realizado}": formatar_moeda_br, 
                "Diferença": formatar_moeda_br, 
                "% Atingido": "{:.2%}", 
                "% Falta Atingir": "{:.2%}"
            }, na_rep="")
            .set_table_styles([
                {
                    'selector': 'thead th',
                    'props': [('background-color', '#dbeeff'),  # azul pastel claro
                              ('color', 'black'),
                              ('font-weight', 'bold')]
                }
            ])
            .apply(formatar_linha, axis=1),
        use_container_width=True
    )
 
    # 🔍 Remove colunas indesejadas apenas do Excel
    colunas_para_remover = ["Tipo", "% Falta Atingir"]
    dados_exportar_excel = dados_exibir.drop(columns=[col for col in colunas_para_remover if col in dados_exibir.columns])
    
    output = io.BytesIO()

    # Alternância de cores entre grupos
    cor_grupo1 = "#fff6eb"
    cor_grupo2 = "#f8faec"
    grupo_cores = {}
    cor_atual = cor_grupo1
    ultimo_grupo = None
    
    # Remove colunas indesejadas
    dados_excel = dados_exibir.drop(columns=["Tipo", "% Falta Atingir", "eh_tipo"], errors="ignore")
    
    # Remove duplicatas da linha FATURAMENTO DESEJÁVEL""
    dados_excel = dados_excel.drop_duplicates(subset=["Loja"], keep="first")
    
    # Substitui NaN por string vazia
    dados_excel = dados_excel.fillna("")
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dados_excel.to_excel(writer, index=False, sheet_name="Metas")
        workbook = writer.book
        worksheet = writer.sheets["Metas"]
    
        # Formatos
        header_format = workbook.add_format({'bold': True, 'bg_color': '#0366d6', 'font_color': 'black', 'border': 1})
        moeda_format_dict = {'num_format': 'R$ #,##0.00', 'border': 1}
        percentual_format_dict = {'num_format': '0.00%', 'border': 1}
        normal_format = workbook.add_format({'border': 1})
    
        estilos_especiais = {
            "Faturamento Desejável": {'bg_color': '#FF6666', 'font_color': 'black', 'border': 1},
            "TOTAL GERAL": {'bg_color': '#0366d6', 'font_color': 'black', 'border': 1},
            "Tipo:": {'bg_color': '#FFE699', 'border': 1},
            "Lojas:": {'bg_color': '#d0e6f7', 'border': 1},
        }
    
       
        
        # Ajusta altura do cabeçalho
        worksheet.set_row(0, 27)
        
        # Ajusta largura das colunas
        worksheet.set_column('A:C', 7)
        worksheet.set_column('D:D', 29)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 22)
        worksheet.set_column('G:G', 10)
        worksheet.set_column('H:H', 15)
        
        # Centraliza e formata o cabeçalho
        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bold': True,
            'bg_color': '#0366d6',
            'border': 1
        })
        for col_num, value in enumerate(dados_excel.columns):
            worksheet.write(0, col_num, value, center_format)

           
        linha_excel = 1
        for _, row in dados_excel.iterrows():
                
            loja_valor = str(row["Loja"])
            grupo_valor = str(row["Grupo"])
            atingido = row.get("% Atingido", "")
    
            # Alternância de cor por grupo
            if grupo_valor != ultimo_grupo and "Lojas:" not in loja_valor and not loja_valor.startswith("Tipo:"):
                cor_atual = cor_grupo2 if cor_atual == cor_grupo1 else cor_grupo1
                ultimo_grupo = grupo_valor
            grupo_cores[grupo_valor] = cor_atual
    
            # Estilo da linha
            # Decide estilo conforme mesmo critério do Styler
            if "FATURAMENTO DESEJÁVEL" in loja_valor.upper():
                estilo_linha = {'bg_color': '#FFE5E5', 'font_color': 'black', 'border': 1}
            elif "TOTAL GERAL" in loja_valor.upper():
                estilo_linha = {'bg_color': '#cce7fc', 'font_color': 'black', 'border': 1}
            elif "- LOJAS:" in loja_valor.upper() and not row["Grupo"]:
                estilo_linha = {'bg_color': '#f9f9f9', 'font_color': 'black', 'border': 1}  # subtotal tipo
            elif "LOJAS:" in loja_valor.upper():
                estilo_linha = {'bg_color': '#e0e0e0', 'font_color': 'black', 'border': 1}  # subtotal grupo
            else:
                estilo_linha = {'bg_color': '#ffffff', 'font_color': 'black', 'border': 1}  # loja normal

    
            for col_num, col_name in enumerate(dados_excel.columns):
                val = row[col_name]
            
                if col_name in ["Meta", f"Realizado até {ultima_data_realizado}", "Diferença"]:
                    fmt = workbook.add_format({**estilo_linha, **moeda_format_dict})
            
                elif col_name == "% Atingido":
                    if "FATURAMENTO DESEJÁVEL" in loja_valor.upper():
                        fmt = workbook.add_format({
                            'bg_color': estilo_linha['bg_color'], 'font_color': 'black',
                            'bold': True, 'num_format': '0.00%', 'border': 1
                        })
                    elif atingido >= percentual_meta_desejavel:
                        fmt = workbook.add_format({
                            'bg_color': estilo_linha['bg_color'], 'font_color': 'green',
                            'bold': True, 'num_format': '0.00%', 'border': 1
                        })
                    else:
                        fmt = workbook.add_format({
                            'bg_color': estilo_linha['bg_color'], 'font_color': 'red',
                            'bold': True, 'num_format': '0.00%', 'border': 1
                        })
                else:
                    # as demais colunas seguem o estilo da linha normalmente
                    fmt = workbook.add_format(estilo_linha)
            
                # Escreve a célula
                if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
                    worksheet.write_number(linha_excel, col_num, val, fmt)
                else:
                    worksheet.write(linha_excel, col_num, str(val), fmt)

                   
    
                   
                # Escreve a célula
                if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
                    worksheet.write_number(linha_excel, col_num, val, fmt)
                else:
                    worksheet.write(linha_excel, col_num, str(val), fmt)
    
            linha_excel += 1
    
    # Finaliza o arquivo e gera o botão de download
    output.seek(0)
    
    st.download_button(
        label="📥 Baixar Excel",
        data=output,
        file_name=f"Relatorio_Metas_{ano_selecionado}_{mes_selecionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_excel_{ano_selecionado}_{mes_selecionado}"
    )




# ================================
with aba2:
# ================================    
    st.info("🔧 Esta aba está em desenvolvimento.")
    st.markdown("## Painel Gráfico Semanal")

    # --- Preparação dos dados (exemplo)
    # Aqui você vai ajustar conforme sua base real
    # Vou simular dados para o esqueleto

    semanas = [22, 23, 24, 25, 26]
    faturamento = [0.7, 5.1, 5.2, 5.4, 0.9]
    ideal = [3.3, 3.4, 3.4, 3.5, 3.4]

    cores_colunas = ["red" if fat < ideal[i] else "green" for i, fat in enumerate(faturamento)]

    import plotly.graph_objects as go

    fig = go.Figure()
    fig.add_trace(go.Bar(x=semanas, y=faturamento, marker_color=cores_colunas, name="Faturamento"))
    fig.add_trace(go.Scatter(x=semanas, y=ideal, mode="lines+markers", name="Meta Ideal", line=dict(color="black")))

    fig.update_layout(
        title="Venda por Semana",
        xaxis_title="Semana Ano",
        yaxis_title="Faturamento 2025 (Mi)",
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Velocímetro (gauge)
    ritmo_real = 71.4  # Exemplo (você vai calcular real depois)

    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=ritmo_real,
        number={'suffix': "%"},
        gauge={
            'axis': {'range': [0, 100]},
            'bar': {'color': "red"},
            'steps': [
                {'range': [0, 70], 'color': "#ffcccc"},
                {'range': [70, 100], 'color': "#ccffcc"}
            ],
        }
    ))
    fig_gauge.update_layout(height=300, margin=dict(t=30, b=30, l=30, r=30), title="Ritmo Ideal")
    st.plotly_chart(fig_gauge, use_container_width=False)

    # --- Tabelinha de operações
    st.markdown("### Operação")
    operacao_df = pd.DataFrame({
        "Operação": ["GF4 - BSB", "GF4 - CFN", "GF4 - CGH", "GF4 - CWB"],
        "2024": [4, 4, 2, 2],
        "2025": [4, 4, 7, 7]
    })
    st.dataframe(operacao_df, use_container_width=True)

    # --- Índice de Crescimento
    st.markdown("### Índice de Crescimento")
    crescimento_df = pd.DataFrame({
        "Operação": ["GF4 - GYN", "GF4 - CWB", "GF4 - VCP", "GF4 - GRU", "GF4 - VIX", "GF4 - REC"],
        "Crescimento": [395.8, 140.5, 6.6, -2.1, -9.5, -10.4]
    })

    fig_cresc = go.Figure(go.Bar(
        x=crescimento_df["Crescimento"],
        y=crescimento_df["Operação"],
        orientation="h",
        marker_color=["green" if x > 0 else "red" for x in crescimento_df["Crescimento"]]
    ))
    fig_cresc.update_layout(height=300)
    st.plotly_chart(fig_cresc, use_container_width=True)

    # --- Participação Faturamento
    st.markdown("### Participação Faturamento")
    participacao_df = pd.DataFrame({
        "Operação": ["GF4 - GRU", "GF4 - VCP", "GF4 - BSB", "GF4 - VIX", "GF4 - CWB", "GF4 - CFN", "GF4 - GYN", "GF4 - FLN"],
        "% Participação": [55.5, 8.7, 6.9, 5.4, 4.7, 4.5, 3.1, 2.5]
    })

    fig_part = go.Figure(go.Bar(
        x=participacao_df["% Participação"],
        y=participacao_df["Operação"],
        orientation="h",
        marker_color="green"
    ))
    fig_part.update_layout(height=300)
    st.plotly_chart(fig_part, use_container_width=True)



# ===========================================
with aba3:
# ===========================================

    uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

    def formatar_excel_contabil(df, nome_aba="Metas"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=nome_aba)
            workbook = writer.book
            worksheet = writer.sheets[nome_aba]

            for idx, cell in enumerate(worksheet[1], 1):
                if cell.value == "Meta":
                    col_meta_idx = idx
                    for row in worksheet.iter_rows(min_row=2, min_col=col_meta_idx, max_col=col_meta_idx):
                        for cell in row:
                            cell.number_format = '#,##0.00'
                    break
        output.seek(0)
        return output

    # 🔁 Estado para armazenar resultado e arquivo anterior
    if "df_resultado" not in st.session_state:
        st.session_state.df_resultado = pd.DataFrame()
    if "nome_arquivo_carregado" not in st.session_state:
        st.session_state.nome_arquivo_carregado = None

    if uploaded_file:
        if st.session_state.nome_arquivo_carregado != uploaded_file.name:
            st.session_state.df_resultado = pd.DataFrame()
        st.session_state.nome_arquivo_carregado = uploaded_file.name

        xls = pd.ExcelFile(uploaded_file)
        todas_abas = xls.sheet_names

        abas_escolhidas = st.multiselect(
            "Selecione as abas a processar:",
            options=todas_abas,
            default=[]
        )

        if not abas_escolhidas:
            st.session_state.df_resultado = pd.DataFrame()

        if abas_escolhidas:
            aba_referencia = abas_escolhidas[0]
            df_preview = pd.read_excel(xls, sheet_name=aba_referencia, header=None).copy()
            df_preview.iloc[1, :] = df_preview.iloc[1, :].ffill()

            linha_lojas = df_preview.iloc[1, :].astype(str).str.strip()
            linha_colunas = df_preview.iloc[2, :].fillna("").astype(str).str.strip()

            colunas_filtradas = []
            for col in range(df_preview.shape[1]):
                loja = linha_lojas[col]
                nome_coluna = linha_colunas[col]
                
                # Limpeza e padronização do texto
                loja = loja.strip() if isinstance(loja, str) else ""
                nome_coluna = nome_coluna.strip().lower() if isinstance(nome_coluna, str) else ""
            
                if not loja or "consolidado" in loja.lower():
                    continue
                if not nome_coluna:
                    continue
                if re.fullmatch(r"\d{4}", nome_coluna):  # ignora colunas como "2025"
                    continue
                if any(substr in nome_coluna for substr in ["%", "variação", "diferença", "dif.", "delta"]):
                    continue
            
                colunas_filtradas.append(linha_colunas[col])  # mantém o nome original (sem .lower())

            colunas_unicas = sorted(set(colunas_filtradas))

            colunas_escolhidas_nomes = st.multiselect(
                "📜 Selecione o(s) nome(s) das colunas abaixo das lojas a serem importadas:",
                options=colunas_unicas,
                default=[]
            )

            if not abas_escolhidas:
                st.session_state.df_resultado = pd.DataFrame()
                st.warning("⚠️ Nenhuma aba selecionada.")
                st.stop()

            if not colunas_escolhidas_nomes:
                st.session_state.df_resultado = pd.DataFrame()
                st.warning("⚠️ Nenhuma coluna selecionada.")
                st.stop()

            import re

            mapa_meses = {
                "janeiro": "Jan", "fevereiro": "Fev", "março": "Mar", "abril": "Abr",
                "maio": "Mai", "junho": "Jun", "julho": "Jul", "agosto": "Ago",
                "setembro": "Set", "outubro": "Out", "novembro": "Nov", "dezembro": "Dez"
            }
            ordem_meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

            df_final = pd.DataFrame(columns=["Mês", "Ano", "Grupo", "Loja", "Meta"])

            for aba in abas_escolhidas:
                df_raw_original = pd.read_excel(xls, sheet_name=aba, header=None)
                df_raw_ffill = df_raw_original.copy()
                
                # Aplicar ffill apenas nas linhas de cabeçalho (0, 1 e 2)
                for i in [0, 1, 2]:
                    df_raw_ffill.iloc[i, :] = df_raw_ffill.iloc[i, :].ffill()
                
                grupo = df_raw_ffill.iloc[0, 0]
                df_raw_ffill.iloc[1, :] = df_raw_ffill.iloc[1, :].ffill()

                linha_lojas = df_raw_ffill.iloc[1, :].astype(str).str.strip()
                linha_colunas = df_raw_ffill.iloc[2, :].fillna("").astype(str).str.strip()

                colunas_validas = {}
                for col in range(df_raw_ffill.shape[1]):
                    # Verifica se o cabeçalho original (linha 2) está vazio → ignora a coluna
                    valor_cabecalho_original = str(df_raw_original.iloc[2, col]).strip().lower()
                    if not valor_cabecalho_original or valor_cabecalho_original in ["", "nan", "-", None]:
                        continue
                
                    nome_coluna = linha_colunas[col].strip()
                    loja = linha_lojas[col].strip()
                
                    if not loja or "consolidado" in loja.lower():
                        continue
                    if re.fullmatch(r"\d{4}", nome_coluna):
                        continue
                    if any(substr in nome_coluna.lower() for substr in ["%", "variação", "diferença", "dif.", "delta"]):
                        continue
                    if nome_coluna not in colunas_escolhidas_nomes:
                        continue
                
                    colunas_validas[col] = (loja, nome_coluna)

                linha_dados_inicio = 4
                for idx in range(linha_dados_inicio, len(df_raw_ffill)):
                    celula_mes = df_raw_original.iloc[idx, 1]
                    if pd.isna(celula_mes):
                        continue
                    mes_original = str(celula_mes).strip().lower().replace("marco", "março")
                    if mes_original in ["", "-", "nan", "total"]:
                        continue
                    if mes_original not in mapa_meses:
                        continue
                    mes = mapa_meses[mes_original]

                    for col, (loja, nome_coluna) in colunas_validas.items():
                        valor = df_raw_ffill.iloc[idx, col]
                        if pd.isna(valor) or str(valor).strip() in ["", "-", "nan"]:
                            continue
                        if isinstance(valor, str):
                            valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
                            try:
                                valor = float(valor)
                            except:
                                continue

                        match_ano = re.search(r"(20\d{2})", nome_coluna)
                        ano = int(match_ano.group(1)) if match_ano else None

                        linha = {"Mês": mes, "Ano": ano, "Grupo": grupo, "Loja": loja, "Meta": valor}
                        df_final = pd.concat([df_final, pd.DataFrame([linha])], ignore_index=True)

            df_final = df_final.drop_duplicates()

            if not df_final.empty and colunas_escolhidas_nomes:
                df_final["Meta"] = df_final["Meta"].fillna(0)
                df_final["Mês"] = pd.Categorical(df_final["Mês"], categories=ordem_meses, ordered=True)
                df_final = df_final.sort_values(["Ano", "Mês", "Loja"])

                total_meta = df_final["Meta"].sum()
                linha_total = pd.DataFrame([{ "Mês": "TOTAL", "Ano": "", "Grupo": "", "Loja": "", "Meta": total_meta }])
                df_final = pd.concat([linha_total, df_final], ignore_index=True)

                df_final_fmt = df_final.copy()
                df_final_fmt["Meta"] = df_final_fmt["Meta"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                st.session_state.df_resultado = df_final_fmt
            else:
                st.session_state.df_resultado = pd.DataFrame()

            with st.container():
                if st.session_state.df_resultado.empty:
                    st.write("")
                else:
                    st.success("✅ Dados consolidados")
                    st.dataframe(st.session_state.df_resultado)
                    excel_file = formatar_excel_contabil(df_final)
                    st.download_button(
                        label="📅 Baixar Excel (.xlsx)",
                        data=excel_file,
                        file_name="metas_consolidado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            if not abas_escolhidas:
                st.session_state.df_resultado = pd.DataFrame()
                st.warning("⚠️ Nenhuma aba selecionada.")
            elif not uploaded_file:
                st.session_state.df_resultado = pd.DataFrame()
                st.info("💡 Faça o upload de um arquivo Excel para começar.")
        

    
