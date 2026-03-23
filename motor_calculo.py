from datetime import datetime, timedelta

def format_timedelta(td):
    total_seconds = int(td.total_seconds())
    sinal = "-" if total_seconds < 0 else ""
    hours, remainder = divmod(abs(total_seconds), 3600)
    minutes, _ = divmod(remainder, 60)
    return f"{sinal}{hours:02d}:{minutes:02d}"

def process_timecards(timecards, mapa_usuarios, regras_contratos, aprovacoes_he, feriados, ferias):
    print("Processando pontos, calculando tolerâncias e faltas...")
    
    agrupamento = {}
    fmt_api = "%Y-%m-%dT%H:%M:%S.%fZ"
    fmt_curto = "%Y-%m-%dT%H:%M:%SZ"

    # Fase 1: Agrupar Batidas
    for tc in timecards:
        uid = tc["userId"]
        reg = tc.get("originalEntry", tc)
        
        # === A TRAVA DE SEGURANÇA QUE FALTAVA ===
        # Se a pessoa ainda não bateu a saída, pula esse registro para não quebrar o código
        if not reg.get("clockOutEvent"):
            print(f"⚠️ Ponto aberto (sem saída) detectado e ignorado na leitura.")
            continue
        try:
            in_dt_utc = datetime.strptime(reg["clockInEvent"]["dateTime"], fmt_api)
            out_dt_utc = datetime.strptime(reg["clockOutEvent"]["dateTime"], fmt_api)
        except ValueError:
            in_dt_utc = datetime.strptime(reg["clockInEvent"]["dateTime"], fmt_curto)
            out_dt_utc = datetime.strptime(reg["clockOutEvent"]["dateTime"], fmt_curto)

        in_dt = in_dt_utc - timedelta(hours=3)
        out_dt = out_dt_utc - timedelta(hours=3)
        data_str = in_dt.strftime("%Y-%m-%d")
        chave = (uid, data_str)
        
        breaks = reg.get("breaks", [])
        if chave not in agrupamento:
            agrupamento[chave] = {"in_dt": in_dt, "out_dt": out_dt, "breaks": breaks}
        else:
            agrupamento[chave]["in_dt"] = min(agrupamento[chave]["in_dt"], in_dt)
            agrupamento[chave]["out_dt"] = max(agrupamento[chave]["out_dt"], out_dt)
            agrupamento[chave]["breaks"].extend(breaks)

    registros = []
    batidas_existentes = set(agrupamento.keys())

    # Fase 2: Processar as Batidas Reais
    for chave, dados in agrupamento.items():
        uid, data_str = chave
        user_info = mapa_usuarios.get(uid, {"nome": "Desconhecido", "email": "", "setor": "Sem Setor"})
        email_func = user_info["email"]
        
        contrato = regras_contratos.get(email_func, {"carga": 8.0, "almoco": 75})
        minutos_multa_almoco = contrato["almoco"]
        
        in_dt = dados["in_dt"]
        out_dt = dados["out_dt"]
        
        desconto_almoco = timedelta(0)
        str_desconto_almoco = f"00:{minutos_multa_almoco} (Padrão)" if minutos_multa_almoco < 60 else "01:15 (Padrão)"
        inicio_intervalo = "-"
        fim_intervalo = "-"
        
        if len(dados["breaks"]) > 0:
            intervalos_proc = []
            for brk in dados["breaks"]:
                try:
                    b_in = datetime.strptime(brk["start"]["dateTime"], fmt_api) - timedelta(hours=3)
                    b_out = datetime.strptime(brk["end"]["dateTime"], fmt_api) - timedelta(hours=3)
                except:
                    b_in = datetime.strptime(brk["start"]["dateTime"], fmt_curto) - timedelta(hours=3)
                    b_out = datetime.strptime(brk["end"]["dateTime"], fmt_curto) - timedelta(hours=3)
                intervalos_proc.append((b_in, b_out))
                desconto_almoco += (b_out - b_in)
            
            intervalos_proc.sort(key=lambda x: x[0])
            inicio_intervalo = intervalos_proc[0][0].strftime("%H:%M")
            fim_intervalo = intervalos_proc[-1][1].strftime("%H:%M")
            str_desconto_almoco = format_timedelta(desconto_almoco)
        else:
            desconto_almoco = timedelta(minutes=minutos_multa_almoco)

        tempo_total = out_dt - in_dt
        horas_reais = tempo_total - desconto_almoco
        horas_brutas = tempo_total - timedelta(minutes=minutos_multa_almoco)

        # CÁLCULO DE HORA EXTRA
        hora_extra = timedelta(0)
        status_he = "-"
        
        limite_18h_padrao = in_dt.replace(hour=18, minute=0, second=0, microsecond=0)
        base_calculo_he = max(in_dt, limite_18h_padrao)
        
        if out_dt > base_calculo_he:
            tempo_excedente = out_dt - base_calculo_he
            if base_calculo_he == limite_18h_padrao and tempo_excedente.total_seconds() <= 600:
                hora_extra = timedelta(0)
            else:
                hora_extra = tempo_excedente
                chave_busca_email = f"{email_func}_{data_str}"
                chave_busca_nome = f"{user_info['nome'].lower()}_{data_str}"
                status_he = aprovacoes_he.get(chave_busca_email) or aprovacoes_he.get(chave_busca_nome) or "NÃO AUTORIZADA"

        h_prev, m_prev = divmod(int(contrato["carga"] * 60), 60)
        
        registros.append({
            "Setor": user_info["setor"],
            "Nome do Colaborador": user_info["nome"],
            "Data": in_dt.strftime("%d/%m/%Y"),
            "Carga Prevista": f"{h_prev:02d}:{m_prev:02d}",
            "Primeira Entrada": in_dt.strftime("%H:%M"),
            "Início do Intervalo": inicio_intervalo,
            "Fim do Intervalo": fim_intervalo,
            "Última Saída": out_dt.strftime("%H:%M"),
            "Desconto Almoço": str_desconto_almoco,
            "Horas Reais": format_timedelta(horas_reais),
            "Horas Brutas": format_timedelta(horas_brutas),
            "Hora Extra (Qtd)": format_timedelta(hora_extra) if hora_extra.total_seconds() > 0 else "00:00",
            "Status HE": status_he
        })

    # Fase 3: Detecção de Faltas
    if agrupamento:
        datas = [d["in_dt"].date() for d in agrupamento.values()]
        data_inicial, data_final = min(datas), max(datas)
        
        for uid, user_info in mapa_usuarios.items():
            email_func = user_info["email"]
            contrato = regras_contratos.get(email_func, {"carga": 8.0, "almoco": 75})
            h_prev, m_prev = divmod(int(contrato["carga"] * 60), 60)
            
            data_atual = data_inicial
            while data_atual <= data_final:
                if data_atual.weekday() <= 4:
                    data_str = data_atual.strftime("%Y-%m-%d")
                    chave_falta = (uid, data_str)
                    
                    if chave_falta not in batidas_existentes:
                        # O ROBÔ DECIDE: É férias, feriado ou falta?
                        status_entrada = "FALTA" # Começa assumindo que é falta

                        # 1º Checa Férias (É a prioridade máxima para a pessoa)
                        esta_de_ferias = False
                        periodos_ferias = ferias.get(email_func, [])
                        for inicio_f, fim_f in periodos_ferias:
                            if inicio_f <= data_str <= fim_f:
                                esta_de_ferias = True
                                break
                                
                        if esta_de_ferias:
                            status_entrada = "FÉRIAS"
                        # 2º Checa Feriado/Recesso da empresa
                        elif data_str in feriados:
                            status_entrada = "RECESSO"
                        
                        registros.append({
                            "Setor": user_info["setor"],
                            "Nome do Colaborador": user_info["nome"],
                            "Data": data_atual.strftime("%d/%m/%Y"),
                            "Carga Prevista": f"{h_prev:02d}:{m_prev:02d}",
                            "Primeira Entrada": status_entrada,
                            "Início do Intervalo": "-",
                            "Fim do Intervalo": "-",
                            "Última Saída": "-",
                            "Desconto Almoço": "-",
                            "Horas Reais": "00:00",
                            "Horas Brutas": "00:00",
                            "Hora Extra (Qtd)": "00:00",
                            "Status HE": "-"
                        })
                data_atual += timedelta(days=1)

    # Fase 4: Ordenação Oficial do RH
    registros.sort(key=lambda x: (x["Setor"], x["Nome do Colaborador"], datetime.strptime(x["Data"], "%d/%m/%Y")))
    
    colunas_finais = ["Data", "Nome do Colaborador", "Setor", "Carga Prevista", "Primeira Entrada", 
                      "Início do Intervalo", "Fim do Intervalo", "Última Saída", "Desconto Almoço", 
                      "Horas Reais", "Horas Brutas", "Hora Extra (Qtd)", "Status HE"]
    
    return [{col: linha[col] for col in colunas_finais} for linha in registros]
