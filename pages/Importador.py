def normalize_block_tokens(block_tokens):
    toks = [t.strip() for t in block_tokens if t is not None and str(t).strip() != ""]
    if not toks:
        return ["", "", "", "", ""]

    # Identificar o último token que é valor monetário (Valor)
    value = ""
    end_idx = len(toks) - 1
    for i in range(len(toks)-1, -1, -1):
        if is_money(toks[i]):
            value = toks[i]
            end_idx = i
            break

    col4 = ""
    hours_idx = None

    # Se o token anterior ao valor for também valor monetário, considerar como horas
    if end_idx >= 1 and is_money(toks[end_idx - 1]):
        col4 = toks[end_idx - 1]
        hours_idx = end_idx - 1
    else:
        # Procurar token que contenha hh:mm ou token 'hs'
        for i in range(end_idx-1, -1, -1):
            if _hours_re.search(toks[i]) or toks[i].lower().endswith('hs') or toks[i].lower() == 'hs':
                # Juntar tokens consecutivos que fazem parte da hora
                hours_tokens = [toks[i]]
                j = i + 1
                while j < end_idx and (toks[j].lower() == 'hs' or _hours_re.search(toks[j])):
                    hours_tokens.append(toks[j])
                    j += 1
                col4 = " ".join(hours_tokens).strip()
                hours_idx = i
                break

    col1 = toks[0] if len(toks) > 0 else ""
    col2 = toks[1] if len(toks) > 1 else ""

    start_desc = 2
    stop_desc = hours_idx if hours_idx is not None else end_idx
    if stop_desc < start_desc:
        stop_desc = start_desc

    desc_tokens = []
    for i in range(start_desc, stop_desc):
        if i < len(toks):
            desc_tokens.append(toks[i])
    col3 = " ".join(desc_tokens).strip()

    return [col1 or "", col2 or "", col3 or "", col4 or "", value or ""]
