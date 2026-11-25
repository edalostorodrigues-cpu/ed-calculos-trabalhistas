# ... existing code ...
        months.append(
            {
                "ano": cur_year,
                "mes": cur_month,
                "competencia": f"{cur_month:02d}/{cur_year}",
                "competencia_iso": f"{cur_year}-{cur_month:02d}",
                "periodo_mensal_pje": periodo_mensal_pje,
                "inicio": m_start.strftime("%d/%m/%Y"),
                "fim": m_end.strftime("%d/%m/%Y"),
                "dias_corridos": cal_days,
                "dias_uteis_calc": work_days,
                "domingos_feriados": max(1, sundays),
                "avos_15": 1 if cal_days >= 15 else 0,
            }
        )
# ... existing code ...
    # G. REFLEXOS LEGAIS EXATOS NO PADRÃO PJE-CALC CIDADÃO
    # (RSR mês a mês, 13º por ano civil, Férias + 1/3 por período aquisitivo,
    #  Aviso Prévio pela média dos últimos 12 meses e FGTS 8% + 40% mês a mês)
    principais_por_comp: Dict[str, Decimal] = {}
    for o in ocorrencias:
        comp_k = str(o.get("competencia", ""))
        principais_por_comp[comp_k] = principais_por_comp.get(comp_k, Decimal("0")) + D(o.get("valor_devido", 0))

    base_reflexos_total = total_adicional + total_he50 + total_he100 + total_intrajornada + total_noturno
    if aplicar_reflexos and base_reflexos_total > 0:
        ref_dsr = Decimal("0")
        dsr_por_comp: Dict[str, Decimal] = {}
        if reflexos_em_dsr:
            bruto_dsr = Decimal("0")
            rows_dsr: List[Dict[str, Any]] = []
            for m in months:
                comp_k = m["competencia"]
                # Base do RSR: parcelas variáveis por hora/dia (exclui adicional mensal fixo)
                base_m_dsr = sum(
                    D(o["valor_devido"])
                    for o in ocorrencias
                    if o["competencia"] == comp_k and o["verba"] != adicional_nome
                )
                if base_m_dsr > 0:
                    du = Decimal(str(max(1, m["dias_efetivos"])))
                    dom = Decimal(str(max(1, m["domingos_feriados"])))
                    v_dsr_m = money_round((base_m_dsr / du) * dom)
                    bruto_dsr += v_dsr_m
                    dsr_por_comp[comp_k] = v_dsr_m
                    rows_dsr.append(
                        {
                            "periodo_mensal": m["periodo_mensal_pje"],
                            "base": float(base_m_dsr),
                            "divisor": float(du),
                            "multiplicador": 1.0,
                            "quantidade": float(dom),
                            "dobra": "Não",
                            "devido": float(v_dsr_m),
                            "pago": 0.0,
                            "diferenca": float(v_dsr_m),
                            "indice_correcao": 1.0,
                            "valor_corrigido": float(v_dsr_m),
                        }
                    )
            if bruto_dsr > 0:
                ref_dsr = add_verba(
                    "071",
                    "REPOUSO SEMANAL REMUNERADO E FERIADO SOBRE PARCELAS",
                    "Contribuição Social / IRPF",
                    "((((PARCELAS HABITUAIS) / DIAS ÚTEIS) X 1,00000000) X REPOUSOS E FERIADOS)",
                    "Apuração mensal oficial PJe-Calc: (Base Mensal / Dias Úteis) x Domingos e Feriados",
                    bruto_dsr,
                    Decimal("0"),
                    rows_dsr,
                )

        # 13º Salário sobre Parcelas Deferidas: apuração oficial por ANO CIVIL
        anos_contrato = sorted({m["ano"] for m in months})
        bruto_13 = Decimal("0")
        rows_13: List[Dict[str, Any]] = []
        for ano_c in anos_contrato:
            meses_ano = [m for m in months if m["ano"] == ano_c]
            avos_ano_c = sum(m["avos_15"] for m in meses_ano)
            if avos_ano_c <= 0:
                continue
            soma_ano_c = sum(
                principais_por_comp.get(m["competencia"], Decimal("0")) + dsr_por_comp.get(m["competencia"], Decimal("0"))
                for m in meses_ano
            )
            if soma_ano_c <= 0:
                continue
            qtd_meses_ano = Decimal(str(len(meses_ano)))
            media_ano_c = (soma_ano_c / qtd_meses_ano).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)
            val_13_ano = money_round((media_ano_c / Decimal("12")) * Decimal(str(avos_ano_c)))
            bruto_13 += val_13_ano
            m_last_ano = meses_ano[-1]
            rows_13.append(
                {
                    "periodo_mensal": m_last_ano["periodo_mensal_pje"],
                    "base": float(money_round(media_ano_c)),
                    "divisor": 12.0,
                    "multiplicador": 1.0,
                    "quantidade": float(avos_ano_c),
                    "dobra": "Não",
                    "devido": float(val_13_ano),
                    "pago": 0.0,
                    "diferenca": float(val_13_ano),
                    "indice_correcao": 1.0,
                    "valor_corrigido": float(val_13_ano),
                }
            )
        ref_13 = Decimal("0")
        if bruto_13 > 0:
            ref_13 = add_verba(
                "072",
                "13º SALÁRIO SOBRE PARCELAS DEFERIDAS",
                "Contribuição Social / IRPF",
                "((((MÉDIA DAS PARCELAS NO ANO) / 12,0000) X 1,00000000) X AVOS NO ANO)",
                "Apuração oficial PJe-Calc por ano civil: (Média Mensal do Ano / 12) x Avos (fração >= 15 dias)",
                bruto_13,
                Decimal("0"),
                rows_13,
            )

        # Férias + 1/3 sobre Parcelas Deferidas: apuração oficial por PERÍODO AQUISITIVO (blocos de 12 meses a partir da admissão)
        bruto_ferias = Decimal("0")
        rows_ferias: List[Dict[str, Any]] = []
        for idx_start in range(0, len(months), 12):
            bloco_aq = months[idx_start : idx_start + 12]
            avos_aq = min(12, sum(m["avos_15"] for m in bloco_aq))
            if avos_aq <= 0:
                continue
            soma_aq = sum(
                principais_por_comp.get(m["competencia"], Decimal("0")) + dsr_por_comp.get(m["competencia"], Decimal("0"))
                for m in bloco_aq
            )
            if soma_aq <= 0:
                continue
            media_aq = (soma_aq / Decimal(str(len(bloco_aq)))).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)
            val_fer_aq = money_round((media_aq / Decimal("12")) * (Decimal("4") / Decimal("3")) * Decimal(str(avos_aq)))
            bruto_ferias += val_fer_aq
            m_last_aq = bloco_aq[-1]
            rows_ferias.append(
                {
                    "periodo_mensal": m_last_aq["periodo_mensal_pje"],
                    "base": float(money_round(media_aq)),
                    "divisor": 12.0,
                    "multiplicador": 1.33333333,
                    "quantidade": float(avos_aq),
                    "dobra": "Não",
                    "devido": float(val_fer_aq),
                    "pago": 0.0,
                    "diferenca": float(val_fer_aq),
                    "indice_correcao": 1.0,
                    "valor_corrigido": float(val_fer_aq),
                }
            )
        ref_ferias = Decimal("0")
        if bruto_ferias > 0:
            ref_ferias = add_verba(
                "073",
                "FÉRIAS + 1/3 SOBRE PARCELAS DEFERIDAS",
                "Não há.",
                "((((MÉDIA PERÍODO AQUISITIVO) / 12,0000) X 1,33333333) X AVOS)",
                "Apuração oficial PJe-Calc por período aquisitivo: (Média Mensal / 12) x 1,33333333 x Avos",
                bruto_ferias,
                Decimal("0"),
                rows_ferias,
            )

        # Aviso Prévio sobre Parcelas Deferidas (média dos últimos 12 meses x dias Lei 12.506/2011)
        ref_aviso = Decimal("0")
        if is_dispensa_sem_justa:
            ultimos_12 = months[-12:]
            soma_12 = sum(
                principais_por_comp.get(m["competencia"], Decimal("0")) + dsr_por_comp.get(m["competencia"], Decimal("0"))
                for m in ultimos_12
            )
            media_12 = (soma_12 / Decimal(str(max(1, len(ultimos_12))))).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)
            val_ref_aviso = money_round((media_12 / Decimal("30")) * Decimal(str(dias_aviso_previo)))
            if val_ref_aviso > 0:
                row_aviso = [
                    {
                        "periodo_mensal": f"{rescisao.day:02d} a {rescisao.day:02d}/{rescisao.month:02d}/{rescisao.year}",
                        "base": float(money_round(media_12)),
                        "divisor": 30.0,
                        "multiplicador": 1.0,
                        "quantidade": float(dias_aviso_previo),
                        "dobra": "Não",
                        "devido": float(val_ref_aviso),
                        "pago": 0.0,
                        "diferenca": float(val_ref_aviso),
                        "indice_correcao": 1.0,
                        "valor_corrigido": float(val_ref_aviso),
                    }
                ]
                ref_aviso = add_verba(
                    "074",
                    "AVISO PRÉVIO SOBRE PARCELAS DEFERIDAS",
                    "Não há.",
                    "((((MÉDIA ÚLTIMOS 12 MESES) / 30,0000) X 1,00000000) X APURADA)",
                    f"Média últimos 12 meses {fmt_brl(media_12)} / 30 x {dias_aviso_previo} dias (Lei 12.506/2011)",
                    val_ref_aviso,
                    Decimal("0"),
                    row_aviso,
                )

        # FGTS 8% (+ 40% se dispensa sem justa causa) apurado mês a mês no padrão PJe-Calc
        if reflexos_em_fgts:
            rows_fgts: List[Dict[str, Any]] = []
            bruto_fgts_8 = Decimal("0")
            for m in months:
                comp_k = m["competencia"]
                base_fgts_m = principais_por_comp.get(comp_k, Decimal("0")) + dsr_por_comp.get(comp_k, Decimal("0"))
                # Soma 13º do ano civil na competência final do ano
                for r13 in rows_13:
                    if r13["periodo_mensal"] == m["periodo_mensal_pje"]:
                        base_fgts_m += D(r13["devido"])
                if m == months[-1]:
                    base_fgts_m += ref_aviso
                if base_fgts_m > 0:
                    v_fgts_m = money_round(base_fgts_m * Decimal("0.08"))
                    bruto_fgts_8 += v_fgts_m
                    rows_fgts.append(
                        {
                            "periodo_mensal": m["periodo_mensal_pje"],
                            "base": float(base_fgts_m),
                            "divisor": 1.0,
                            "multiplicador": 0.08,
                            "quantidade": 1.0,
                            "dobra": "Não",
                            "devido": float(v_fgts_m),
                            "pago": 0.0,
                            "diferenca": float(v_fgts_m),
                            "indice_correcao": 1.0,
                            "valor_corrigido": float(v_fgts_m),
                        }
                    )
            if bruto_fgts_8 > 0:
                add_verba(
                    "075",
                    "FGTS 8%",
                    "Fundiária",
                    "((((BASE INCIDENTE FGTS) / 1,0000) X 0,08000000) X 1,0000)",
                    "Apuração mensal oficial PJe-Calc: 8% sobre parcelas remuneratórias incidentes",
                    bruto_fgts_8,
                    Decimal("0"),
                    rows_fgts,
                    comentario="PAGAR AO RECLAMANTE" if is_dispensa_sem_justa else "DEPOSITAR / PAGAR CONFORME SENTENÇA",
                )
                if is_dispensa_sem_justa:
                    multa_40_fgts = money_round(bruto_fgts_8 * Decimal("0.40"))
                    row_m40 = [
                        {
                            "periodo_mensal": f"{rescisao.day:02d} a {rescisao.day:02d}/{rescisao.month:02d}/{rescisao.year}",
                            "base": float(bruto_fgts_8),
                            "divisor": 1.0,
                            "multiplicador": 0.40,
                            "quantidade": 1.0,
                            "dobra": "Não",
                            "devido": float(multa_40_fgts),
                            "pago": 0.0,
                            "diferenca": float(multa_40_fgts),
                            "indice_correcao": 1.0,
                            "valor_corrigido": float(multa_40_fgts),
                        }
                    ]
                    add_verba(
                        "076",
                        "MULTA DO FGTS 40%",
                        "Fundiária",
                        "((((DEPÓSITOS FGTS APURADOS) / 1,0000) X 0,40000000) X 1,0000)",
                        f"40% sobre FGTS apurado ({fmt_brl(bruto_fgts_8)})",
                        multa_40_fgts,
                        Decimal("0"),
                        row_m40,
                        comentario="PAGAR AO RECLAMANTE",
                    )
# ... existing code ...
    # Separação Verbas x FGTS para o quadro "Descrição de Créditos e Descontos do Reclamante" do PJe-Calc
    total_fgts_bruto = money_round(sum(D(v.get("total_verba", v["valor"])) for v in verbas if v["codigo"] in ("075", "076", "086")))
    total_verbas_bruto = money_round(subtotal_corrigido - total_fgts_bruto)
# ... existing code ...
