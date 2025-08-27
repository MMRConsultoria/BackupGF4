# utils/sessoes.py
import pytz, uuid
from datetime import datetime

def _open_aba_sessoes(gc, planilha_key, sheet_name):
    planilha = gc.open_by_key(planilha_key)
    try:
        aba = planilha.worksheet(sheet_name)
    except:
        aba = planilha.add_worksheet(title=sheet_name, rows=200, cols=6)
        aba.update("A1:E1", [["email","token","data","hora","ultimo_acesso"]])
    return aba

def _liberar_sessao(gc, planilha_key, sheet_name, email):
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    todas = aba.get_all_values()
    if not todas: return
    novas = [todas[0]] + [row for row in todas[1:] if row and row[0] != email]
    aba.clear()
    aba.update("A1", novas)

def registrar_sessao_assumindo(gc, planilha_key, sheet_name, email):
    """Login sempre assume: remove sessão antiga e cria nova."""
    _liberar_sessao(gc, planilha_key, sheet_name, email)
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    fuso = pytz.timezone("America/Sao_Paulo"); agora = datetime.now(fuso)
    token = str(uuid.uuid4())
    aba.append_row([email, token, agora.strftime("%d/%m/%Y"),
                    agora.strftime("%H:%M:%S"), agora.isoformat()])
    return token

def validar_sessao(gc, planilha_key, sheet_name, email, token):
    """Confere se o token na planilha ainda é o mesmo desta máquina."""
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    for r in aba.get_all_records():
        if r.get("email") == email:
            return r.get("token") == token
    return False

def atualizar_sessao(gc, planilha_key, sheet_name, email):
    """Renova carimbo de tempo da sessão atual (se existir)."""
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    todas = aba.get_all_values()
    if not todas: return
    fuso = pytz.timezone("America/Sao_Paulo"); agora = datetime.now(fuso)
    for i in range(1, len(todas)):
        row = todas[i]
        if row and row[0] == email:
            if len(row) < 5: row += [""] * (5 - len(row))
            row[2] = agora.strftime("%d/%m/%Y")
            row[3] = agora.strftime("%H:%M:%S")
            row[4] = agora.isoformat()
            todas[i] = row
    aba.clear()
    aba.update("A1", todas)

def encerrar_sessao(gc, planilha_key, sheet_name, email):
    """Logout explícito (opcional)."""
    _liberar_sessao(gc, planilha_key, sheet_name, email)
