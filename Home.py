# Home.py
import streamlit as st

# ✅ Configuração da página
st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# ✅ Verificação de login ANTES de exibir o conteúdo
if not st.session_state.get("acesso_liberado"):
    st.switch_page("pages/Login.py")
    st.stop()




# ✅ Código da empresa logada
codigo_empresa = st.session_state.get("empresa")

# ✅ Dicionário com logos por código
LOGOS_CLIENTES = {
    "1825": "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_grupofit.png",
    "3377": "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/rossi_ferramentas_logo.png",
    "0041": "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_empresa3.png"
}

# ✅ Busca o logo do cliente
logo_cliente = LOGOS_CLIENTES.get(codigo_empresa)

# ✅ Logo do cliente na sidebar
if logo_cliente:
    st.sidebar.markdown(
        f"""
        <div style="text-align: center; padding: 10px 0 30px 0;">
            <img src="{logo_cliente}" width="100">
        </div>
        """,
        unsafe_allow_html=True
    )

# ✅ Logo principal na área central (cliente ou MMR)
st.image(logo_cliente or "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo-mmr.png", width=150)

# ✅ Mensagem de boas-vindas
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {codigo_empresa}!")
