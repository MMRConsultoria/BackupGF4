import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

def atualizar_planilhas(
    gc,
    id_planilha_origem: str,
    nome_aba_origem: str,
    ids_pastas_destino: list,
    grupo_filtro: str = None,
    filtro_extra: str = None,
    data_minima: datetime = None,
):
    """
    Atualiza planilhas destino nas pastas indicadas com dados filtrados da planilha origem.

    Args:
        gc: objeto gspread autorizado
        id_planilha_origem: ID da planilha origem (cache)
        nome_aba_origem: nome da aba na planilha origem para ler dados
        ids_pastas_destino: lista de IDs das pastas que contêm as planilhas destino
        grupo_filtro: filtro para coluna 'Grupo' (str maiúscula)
        filtro_extra: filtro extra para coluna extra (str maiúscula)
        data_minima: filtra datas >= data_minima (datetime)
    Returns:
        dict com resumo: total_atualizados, lista de falhas
    """
    planilha_origem = gc.open_by_key(id_planilha_origem)
    aba_origem = planilha_origem.worksheet(nome_aba_origem)
    dados = aba_origem.get_all_values()
    if not dados or len(dados) < 2:
        raise ValueError(f"Aba '{nome_aba_origem}' está vazia ou não tem dados suficientes.")

    df = pd.DataFrame(dados[1:], columns=dados[0])
    df.columns = [c.strip() for c in df.columns]

    # Normaliza colunas importantes
    if "Grupo" not in df.columns or "Data" not in df.columns:
        raise ValueError("Colunas 'Grupo' e/ou 'Data' não encontradas na origem.")

    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.upper()
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    idx_col_extra = 2  # Coluna C (0-based)

    total_atualizados = 0
    falhas = []

    for id_pasta in ids_pastas_destino:
        try:
            pasta = gc.drive_service.files().get(fileId=id_pasta).execute()
        except Exception as e:
            falhas.append(f"Erro ao acessar pasta {id_pasta}: {e}")
            continue

        arquivos = gc.list_spreadsheet_files_in_folder(id_pasta)  # Você pode implementar essa função via Drive API

        for arquivo in arquivos:
            try:
                planilha_destino = gc.open_by_key(arquivo['id'])
                abas = planilha_destino.worksheets()
                aba_filtro = next((aba for aba in abas if "rel comp" in aba.title.lower()), None)
                if not aba_filtro:
                    falhas.append(f"{arquivo['name']} - Aba 'rel comp' não encontrada")
                    continue

                grupo_aba = aba_filtro.acell("B4").value
                if not grupo_aba:
                    falhas.append(f"{arquivo['name']} - Grupo em B4 vazio")
                    continue
                grupo_aba = grupo_aba.strip().upper()

                filtro_extra_aba = aba_filtro.acell("B6").value
                filtro_extra_aba = filtro_extra_aba.strip().upper() if filtro_extra_aba else None

                # Filtra dados
                def linha_valida(linha):
                    grupo = str(linha["Grupo"]).strip().upper()
                    data = linha["Data"]
                    extra = str(linha.iloc[idx_col_extra]).strip().upper() if len(linha) > idx_col_extra else None
                    if pd.isna(data):
                        return False
                    if data_minima and data < data_minima:
                        return False
                    grupo_ok = (grupo == grupo_aba)
                    extra_ok = (not filtro_extra_aba) or (extra == filtro_extra_aba)
                    return grupo_ok and extra_ok

                df_filtrado = df[df.apply(linha_valida, axis=1)]

                if df_filtrado.empty:
                    falhas.append(f"{arquivo['name']} - Nenhum dado para grupo '{grupo_aba}'")
                    continue

                # Atualiza aba destino
                try:
                    aba_destino = planilha_destino.worksheet("Importado_Fat")
                except gspread.exceptions.WorksheetNotFound:
                    aba_destino = planilha_destino.add_worksheet(title="Importado_Fat", rows="1000", cols=str(len(df_filtrado.columns)))

                aba_destino.clear()
                valores = [df_filtrado.columns.tolist()] + df_filtrado.values.tolist()
                aba_destino.update(valores)

                total_atualizados += 1

            except Exception as e:
                falhas.append(f"{arquivo['name']} - Erro: {e}")

    return {"total_atualizados": total_atualizados, "falhas": falhas}
