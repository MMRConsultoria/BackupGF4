# utils/sessoes.py
import pytz, uuid, time
from datetime import datetime

def _open_aba_sessoes(gc, planilha_key, sheet_name):
    planilha = gc.open_by_key(planilha_key)
    try:
        aba = planilha.worksheet(sheet_name)
    except:
        aba = planilha.add_worksheet(title=sheet_name, rows=200, cols=6)
        aba.update("A1:E1", [["email","token","data","hora","ultimo_acesso"]])
    return aba

def _remover_linhas_email(aba, email):
    """Remove todas as linhas (exceto cabeçalho) cujo email == email."""
    vals = aba.get_all_values()
    if not vals or len(vals) < 2:
        return
    header, rows = vals[0], vals[1:]
    # Reescreve a planilha mantendo apenas linhas de outros e-mails
    novas = [header] + [r for r in rows if (len(r) >= 1 and r[0] != email)]
    aba.clear()
    aba.update("A1", novas)

def registrar_sessao_assumindo(gc, planilha_key, sheet_name, email):
    """
    Remove qualquer sessão anterior do e-mail e cria uma nova (token único).
    """
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    _remover_linhas_email(aba, email)

    fuso = pytz.timezone("America/Sao_Paulo")
    agora = datetime.now(fuso)
    token = str(uuid.uuid4())
    nova = [email, token, agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), agora.isoformat()]
    aba.append_row(nova)
    # pequena espera para consistência do Sheets (evita corrida)
    time.sleep(0.3)
    return token

def validar_sessao(gc, planilha_key, sheet_name, email, token):
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    for r in aba.get_all_records():
        if r.get("email") == email:
            return r.get("token") == token
    return False

def atualizar_sessao(gc, planilha_key, sheet_name, email):
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    vals = aba.get_all_values()
    if not vals or len(vals) < 2:
        return
    header, rows = vals[0], vals[1:]
    fuso = pytz.timezone("America/Sao_Paulo"); agora = datetime.now(fuso)
    for i, r in enumerate(rows, start=2):  # linhas 2..N
        if len(r) >= 1 and r[0] == email:
            if len(r) < 5:
                r += [""] * (5 - len(r))
            r[2] = agora.strftime("%d/%m/%Y")
            r[3] = agora.strftime("%H:%M:%S")
            r[4] = agora.isoformat()
            aba.update(f"A{i}:E{i}", [r[:5]])
            break

def encerrar_sessao(gc, planilha_key, sheet_name, email):
    aba = _open_aba_sessoes(gc, planilha_key, sheet_name)
    _remover_linhas_email(aba, email)
