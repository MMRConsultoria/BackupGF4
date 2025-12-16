def desdobrar_tabela(df):
    n = 5  # número fixo de colunas por bloco

    bloco1 = df.iloc[:, :n].copy()
    bloco2 = df.iloc[:, n:].copy()

    # Ajustar número de colunas para serem iguais
    diff = bloco1.shape[1] - bloco2.shape[1]
    if diff > 0:
        for i in range(diff):
            bloco2[f"extra_{i}"] = ""
    elif diff < 0:
        for i in range(-diff):
            bloco1[f"extra_{i}"] = ""

    bloco2.columns = bloco1.columns

    df_desdobrado = pd.concat([bloco1, bloco2], ignore_index=True)

    # Garantir que o DataFrame tenha exatamente 5 colunas
    if df_desdobrado.shape[1] > 5:
        df_desdobrado = df_desdobrado.iloc[:, :5]
    elif df_desdobrado.shape[1] < 5:
        for i in range(5 - df_desdobrado.shape[1]):
            df_desdobrado[f"extra_fill_{i}"] = ""

    df_desdobrado.columns = ["Col1", "Col2", "Col3", "Col4", "Col5"]

    df_desdobrado = df_desdobrado.dropna(how='all').reset_index(drop=True)

    return df_desdobrado
