import streamlit as st

st.set_page_config(page_title="Relatórios", layout="wide", initial_sidebar_state="expanded")

# Definição das páginas (apontando para arquivos dentro de pages/)
acesso  = st.Page("pages/Login.py",             title="Acesso")
sangria = st.Page("pages/Operacional Meio Pagamento.py",           title="Meio Pagamento")
caixa   = st.Page("pages/Operacional Vendas Diárias.py",    title="Vendas Diarias")
caixa1   = st.Page("pages/Controle Caixa e Sangria.py",    title="Caixa e Sangria")
evx     = st.Page("pages/Painel Metas.py", title="Metas")
evx1    = st.Page("pages/Rateio.py", title="Rateio")
painel  = st.Page("pages/Relatório.py",  title="Relatórios")

# Gate de acesso: antes do login, mostra só a página "Acesso"
if not st.session_state.get("acesso_liberado"):
    nav = st.navigation({"": [acesso]})
else:
    nav = st.navigation({
        "Relatórios Caixa e Sangria": [acesso, sangria, caixa, caixa1, evx, evx1],
        "Outros": [painel],
    })

nav.run()  # executa apenas a página selecionada

