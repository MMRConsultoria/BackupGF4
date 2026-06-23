st.markdown("---")
st.markdown("### Testar faturamento")

loja_id = st.selectbox(
    "Escolha a loja",
    [
        "5fa1ff35-6145-4130-8707-8fd62e8822a5",
        "0e377e7a-9920-4674-916f-2ab196f8b691"
    ]
)

dtinicio = st.date_input("Data início")
dtfim = st.date_input("Data fim")

if st.button("Buscar faturamento"):
    resp = requests.get(
        "https://api.zigcore.com.br/integration/erp/faturamento",
        headers={"Authorization": token},
        params={
            "dtinicio": dtinicio.strftime("%Y-%m-%d"),
            "dtfim": dtfim.strftime("%Y-%m-%d"),
            "loja": loja_id
        },
        timeout=60
    )

    st.write("Status:", resp.status_code)
    st.write("URL:", resp.url)

    try:
        dados = resp.json()
        st.json(dados)

        if isinstance(dados, list):
            df = pd.DataFrame(dados)

            if "value" in df.columns:
                df["valor_reais"] = df["value"] / 100

            st.dataframe(df, use_container_width=True)

            if "valor_reais" in df.columns:
                st.write("Total:", df["valor_reais"].sum())

    except Exception:
        st.text(resp.text)
