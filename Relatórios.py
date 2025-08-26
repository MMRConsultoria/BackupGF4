import streamlit as st

st.set_page_config(page_title="MMR Consultoria – Relatórios", layout="centered")

if not st.session_state.get("acesso_liberado"):
    st.stop()


st.title("📊 Relatórios – MMR Consultoria")
st.markdown("Selecione um relatório no menu lateral à esquerda.")
st.image("logo-mmr.png", width=200)
