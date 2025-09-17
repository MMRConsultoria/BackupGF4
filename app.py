import streamlit as st
st.set_page_config(page_title="Relatórios", layout="wide", initial_sidebar_state="expanded")
st.caption(f"Streamlit {st.__version__}")  # opcional: só pra confirmar versão em runtime

# Páginas (apontando para os arquivos na raiz / pasta pages)
acesso = st.Page("pages/Login.py",                      title="Acesso",            icon=":material/lock:")
home   = st.Page("Home.py",                             title="Início",            icon="🏠")
meio   = st.Page("pages/Operacional Meio Pagamento.py", title="Meio Pagamento",    icon="💳")
vendas = st.Page("pages/Operacional Vendas Diárias.py", title="Vendas Diárias",    icon="📈")
caixa  = st.Page("pages/Controle Caixa e Sangria.py",   title="Caixa e Sangria",   icon="💸")
metas  = st.Page("pages/Painel Metas.py",               title="Metas",             icon="🎯")
rateio = st.Page("pages/Rateio.py",                     title="Rateio",            icon="🧮")
relats = st.Page("pages/Relatorios.py",                 title="Relatórios",        icon=":material/description:")

# Antes do login: só "Acesso"
if not st.session_state.get("acesso_liberado"):
    nav = st.navigation({"": [acesso]})
else:
    # Depois do login: ordem e grupos do jeito que você quer
    nav = st.navigation({
        "Início": [home],
        "Relatórios Caixa e Sangria": [vendas, meio, caixa],
        "Outros": [metas, rateio, relats],
    })

nav.run()
