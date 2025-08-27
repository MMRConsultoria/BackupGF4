# utils/sessoes.py
import pytz
import uuid
from datetime import datetime

# Espera receber um cliente gspread autorizado (gc) e a PLANILHA_KEY
SHEET_SESSOES = "SessõesAtivas"

def _open_aba_sessoes(gc, planilha_key):
    planilha = gc.open_by_key(planilha_key)
    try:
        aba = planilha.worksheet(SHEET_SESSOES)
    except:
        aba = planilha.add_worksheet(title=SHEET_SESSOES, rows=200, cols=6)
        aba.update("A1:E1", [["email", "token", "data", "hora", "ultimo_acesso"]])
    return aba

def _liberar_sessao(gc, planilha_key, email: str):
    """Remove TODAS as linhas desse e-mail (derruba sessão anterior)."""
    aba = _open_aba_sessoes(gc, planilha_key)
    todas = aba.get_all_values()
    if not todas:
        return
    novas = [todas[0]] + [row for row in todas[1:] if row and row[0] != email]
    aba.clear()
    aba.update("A1", novas)

def registrar_sessao_assumindo(gc, planilha_key, email: str):
    """
    Login preemptivo: apaga sessão anterior desse e-mail e cria uma nova.
    Retorna o token novo.
    """
    _liberar_sessao(gc, planilha_key, email)

    aba = _open_aba_sessoes(gc, planilha_key)
    fuso = pytz.timezone("America/Sao_Paulo")
    agora = datetime.now(fuso)
    token = str(uuid.uuid4())
    nova = [email, token, agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), agora.isoformat()]
    aba.append_row(nova)
    return token

def atualizar_sessao(gc, planilha_key, email: str):
    """Renova carimbo de tempo da sessão atual (se existir)."""
    try:
        aba = _open_aba_sessoes(gc, planilha_key)
        todas = aba.get_all_values()
        if not todas:
            return
        fuso = pytz.timezone("America/Sao_Paulo")
        agora = datetime.now(fuso)
        for i in range(1, len(todas)):
            row = todas[i]
            if row and row[0] == email:
                if len(row) < 5:
                    row += [""] * (5 - len(row))
                row[2] = agora.strftime("%d/%m/%Y")
                row[3] = agora.strftime("%H:%M:%S")
                row[4] = agora.isoformat()
                todas[i] = row
        aba.clear()
        aba.update("A1", todas)
    except Exception:
        pass

def validar_sessao(gc, planilha_key, email: str, token: str) -> bool:
    """Confere se o token na planilha ainda é o mesmo desta máquina."""
    try:
        aba = _open_aba_sessoes(gc, planilha_key)
        registros = aba.get_all_records()
        for r in registros:
            if r.get("email") == email:
                return r.get("token") == token
        return False
    except Exception:
        return False

def encerrar_sessao(gc, planilha_key, email: str):
    try:
        _liberar_sessao(gc, planilha_key, email)
    except Exception:
        pass

