import streamlit as st
import pandas as pd
import json
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title="Atualizar Planilhas Google", layout="wide")

st.title("Atualização Automática de Planilhas Google")

# Autenticação Google Sheets e Drive
@st.cache_resource
def autenticar_gspread():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    drive_service = build('drive', 'v3', credentials=credentials)
    return gc, drive_service

gc, drive_service = autenticar_gspread()

# Parâmetros fixos (pode adaptar para inputs do usuário)
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"  # Exemplo
NOME_ABA_ORIGEM = "Fat Sistema Externo"
IDS_PASTAS_DESTINO = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
    "1GdGvFRvikkFR1S-R9lGRfiPW0T1XD9FG"
]
DATA_MINIMA = datetime(2025, 1, 1)

def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    page_token = None
    query = f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    while True:
        try:
            response = drive_service.files().list(
                q=query,
                spaces='drive',
                fields='nextPageToken, files(id, name)',
                pageToken=page_token
            ).execute()
            arquivos.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break
        except HttpError as e:
            st.error(f"Erro ao listar arquivos na pasta {pasta_id}: {e}")
            break
    return arquivos

def atualizar_planilhas_varias_pastas(
    gc,
    drive_service,
    id_planilha_origem,
    nome_aba_origem,
    ids_pastas_destino,
    data_minima=None,
):
    planilha_origem = gc.open_by_key(id_planilha_origem)
    aba_origem = planilha_origem.worksheet(nome_aba_origem)
    dados = aba_origem.get_all_values()
    if not dados or len(dados) < 2:
        st.error(f"Aba '{nome_aba_origem}' está vazia ou não tem dados suficientes.")
        return 0, []

    df = pd.DataFrame(dados[1:], columns=dados[0])
    df.columns = [c.strip() for c in df.columns]

    if "Grupo" not in df.columns or "Data" not in df.columns:
        st.error("Colunas 'Grupo' e/ou 'Data' não encontradas na origem.")
        return 0, []

    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.upper()
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")

    total_atualizados = 0
    falhas = []

    for id_pasta in ids_pastas_destino:
        arquivos = listar_arquivos_pasta(drive_service, id_pasta)
        if not arquivos:
            falhas.append(f"Pasta {id_pasta} está vazia ou inacessível.")
            continue

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

                def linha_valida(linha):
                    grupo = str(linha["Grupo"]).strip().upper()
                    data = linha["Data"]
                    if pd.isna(data):
                        return False
                    if data_minima and data < data_minima:
                        return False
                    return grupo == grupo_aba

                df_filtrado = df[df.apply(linha_valida, axis=1)]

                if df_filtrado.empty:
                    falhas.append(f"{arquivo['name']} - Nenhum dado para grupo '{grupo_aba}'")
                    continue

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

    return total_atualizados, falhas

if st.button("Atualizar Planilhas"):
    with st.spinner("Atualizando planilhas..."):
        total, falhas = atualizar_planilhas_varias_pastas(
            gc,
            drive_service,
            ID_PLANILHA_ORIGEM,
            NOME_ABA_ORIGEM,
            IDS_PASTAS_DESTINO,
            data_minima=DATA_MINIMA
        )
    st.success(f"Total de planilhas atualizadas: {total}")
    if falhas:
        st.warning("Falhas encontradas:")
        for f in falhas:
            st.write(f"- {f}")
