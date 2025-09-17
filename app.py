import streamlit as st

st.set_page_config(page_title="Relatórios", layout="wide", initial_sidebar_state="expanded")

# a pasta é "páginas", não "pages"
acesso  = st.Page("pages/Login.py",                         title="Acesso")
meio    = st.Page("pages/Operacional Meio Pagamento.py",    title="Meio Pagamento")
vendas  = st.Page("pages/Operacional Vendas Diárias.py",    title="Vendas Diárias")
caixa   = st.Page("pages/Controle Caixa e Sangria.py",      title="Caixa e Sangria")
metas   = st.Page("pages/Painel Metas.py",                  title="Metas")
rateio  = st.Page("pages/Rateio.py",                        title="Rateio")
relats  = st.Page("pages/Relatorios.py",                    title="Relatórios")  # <- sem acento e plural

# Antes do login: só "Acesso"
if not st.session_state.get("acesso_liberado"):
    nav = st.navigation({"": [acesso]})
else:
    # Depois do login: não inclui "Acesso" (se quiser esconder)
    nav = st.navigation({
        "Relatórios Caixa e Sangria": [vendas, meio, caixa],
        "Outros": [metas, rateio, relats],
    })

nav.run()
