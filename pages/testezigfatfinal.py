    # ==========================================
    # TABELA FINAL - FATURAMENTO MEIO PAGAMENTO
    # Crédito/Débito = Máquina Integrada
    # Demais meios = Faturamento Geral
    # ==========================================

    df_fat["Eh Cartao"] = df_fat["Meio de Pagamento"].apply(eh_credito_debito)
    df_maq["Eh Cartao"] = df_maq["Meio de Pagamento"].apply(eh_credito_debito)

    df_cartao = df_maq[df_maq["Eh Cartao"] == True].copy()
    df_outros = df_fat[df_fat["Eh Cartao"] == False].copy()

    df_cartao["Tipo Origem"] = "Máquina Integrada"
    df_outros["Tipo Origem"] = "Faturamento Geral"

    df_final = pd.concat(
        [df_cartao, df_outros],
        ignore_index=True
    )

    df_final["Valor"] = pd.to_numeric(
        df_final["Valor"],
        errors="coerce"
    ).fillna(0)

    df_final["Bandeira"] = df_final["Bandeira"].fillna("").astype(str).str.upper()

    df_final["Data"] = pd.to_datetime(
        df_final["Data"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    df_final["Dia da Semana"] = df_final["Data"].dt.day_name()
    df_final["Mês"] = df_final["Data"].dt.month
    df_final["Ano"] = df_final["Data"].dt.year

    df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")

    tabela_meio_pagamento = (
        df_final
        .groupby(
            [
                "Data",
                "Dia da Semana",
                "Mês",
                "Ano",
                "Loja",
                "Loja ID",
                "Meio de Pagamento",
                "Bandeira",
                "Tipo Origem"
            ],
            as_index=False
        )
        .agg({"Valor": "sum"})
    )

    tabela_meio_pagamento["Valor"] = tabela_meio_pagamento["Valor"].round(2)

    tabela_meio_pagamento = tabela_meio_pagamento.rename(columns={
        "Loja ID": "Cód Loja",
        "Tipo Origem": "Origem"
    })

    st.subheader("Faturamento Meio de Pagamento")
    st.dataframe(tabela_meio_pagamento, use_container_width=True, hide_index=True)

    total_mp = tabela_meio_pagamento["Valor"].sum()

    st.metric("Total Meio de Pagamento", brl(total_mp))

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        tabela_meio_pagamento.to_excel(
            writer,
            index=False,
            sheet_name="Faturamento Meio Pagamento"
        )

        df_cartao.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Credito Debito Maquina"
        )

        df_outros.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Outros Faturamento"
        )

    output.seek(0)

    st.download_button(
        label="📥 Baixar Faturamento Meio Pagamento",
        data=output,
        file_name=f"zig_faturamento_meio_pagamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
