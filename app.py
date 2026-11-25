# ... existing code ...
import os
import io
import json
import hashlib
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import List, Optional, Dict, Any, Tuple
# ... existing code ...
def _serialize_calc_for_cloud(calc: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(calc)
    for k in ("bruto_atualizado", "abatimentos", "liquido_devido", "base_remuneracao"):
        if k in out:
            out[k] = str(d_money(out[k]))
    for k in ("admission_date", "dismissal_date"):
        if isinstance(out.get(k), (date, datetime)):
            out[k] = out[k].isoformat()[:10]
    verbas_clean = []
    for v in (out.get("verbas") or []):
        vc = dict(v)
        for mk in ("Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"):
            if mk in vc:
                vc[mk] = str(d_money(vc[mk]))
        verbas_clean.append(vc)
    out["verbas"] = verbas_clean
    occ_clean = []
    for oc in (out.get("ocorrencias_mensais") or []):
        occ_item = dict(oc)
        for mk in ("Base de Cálculo (R$)", "Devido (R$)", "Pago (R$)", "Diferença (R$)", "Correção/Juros (R$)", "Total Corrigido (R$)"):
            if mk in occ_item:
                occ_item[mk] = str(d_money(occ_item[mk]))
        occ_clean.append(occ_item)
    out["ocorrencias_mensais"] = occ_clean
    return out


def _deserialize_calc_from_cloud(payload: Dict[str, Any]) -> Dict[str, Any]:
    calc = dict(payload)
    for k in ("bruto_atualizado", "abatimentos", "liquido_devido", "base_remuneracao"):
        if k in calc:
            calc[k] = d_money(calc[k])
    verbas_restored = []
    for v in (calc.get("verbas") or []):
        vr = dict(v)
        for mk in ("Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"):
            if mk in vr:
                vr[mk] = d_money(vr[mk])
        verbas_restored.append(vr)
    calc["verbas"] = verbas_restored
    return calc
# ... existing code ...
# ==============================================================================
# 4. MOTOR DE CÁLCULO TRABALHISTA DECIMAL EXATO (PADRÃO PJE-CALC)
# ==============================================================================
def count_working_days_in_range(start_dt: date, end_dt: date, work_days_per_week: int = 5) -> int:
    """Conta dias úteis reais no calendário entre duas datas conforme escala semanal (5=Seg-Sex, 6=Seg-Sáb)."""
    if end_dt < start_dt:
        return 0
    max_wd = 4 if work_days_per_week <= 5 else 5
    cur = start_dt
    count = 0
    while cur <= end_dt:
        if cur.weekday() <= max_wd:
            count += 1
        cur += timedelta(days=1)
    return count


def iter_contract_competencias(adm: date, dism: date, work_days_per_week: int = 5) -> List[Dict[str, Any]]:
    """Gera lista mês a mês (competências YYYY-MM) com dias úteis exatos proporcionais ao vínculo."""
    if dism < adm:
        return []
    comps: List[Dict[str, Any]] = []
    cur_y, cur_m = adm.year, adm.month
    while (cur_y, cur_m) <= (dism.year, dism.month):
        first_day = date(cur_y, cur_m, 1)
        if cur_m == 12:
            last_day = date(cur_y, 12, 31)
        else:
            last_day = date(cur_y, cur_m + 1, 1) - timedelta(days=1)
        eff_start = max(adm, first_day)
        eff_end = min(dism, last_day)
        wdays = count_working_days_in_range(eff_start, eff_end, work_days_per_week)
        comps.append({
            "competencia": f"{cur_m:02d}/{cur_y}",
            "iso_ym": f"{cur_y}-{cur_m:02d}",
            "inicio": eff_start.strftime("%d/%m/%Y"),
            "fim": eff_end.strftime("%d/%m/%Y"),
            "dias_uteis": wdays,
        })
        if cur_m == 12:
            cur_y += 1
            cur_m = 1
        else:
            cur_m += 1
    return comps


def run_labor_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
    adm: date = params["admission_date"]
    dism: date = params["dismissal_date"]
    term_type: str = params.get("termination_type", "sem_justa_causa")
    divisor = Decimal(str(max(1, f_int(params.get("monthly_divisor"), 220))))
    work_days_week = max(5, min(6, f_int(params.get("work_days_per_week"), 5)))
    explicit_worked_days = max(0, f_int(params.get("explicit_worked_days"), 0))

    sal_base = d_money(params.get("base_salary"))
    comissoes = d_money(params.get("commissions"))
    insal_mensal = d_money(params.get("insalubrity"))
    plus_acumulo_pct = Decimal(str(f_float(params.get("plus_acumulo_pct"), 0.0))) / Decimal("100")
    dif_equip = d_money(params.get("dif_equiparacao_mensal"))

    plus_acumulo_val = (sal_base * plus_acumulo_pct).quantize(CENT, rounding=ROUND_HALF_UP)
    remuneracao_base = sal_base + comissoes + insal_mensal + plus_acumulo_val + dif_equip
    valor_hora = (remuneracao_base / divisor).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    anos_completos = max(0, (dism - adm).days // 365)
    dias_aviso = min(90, 30 + anos_completos * 3)
    competencias_list = iter_contract_competencias(adm, dism, work_days_week)
    meses_contrato = max(1, len(competencias_list))
    meses_presc = min(60, meses_contrato)
    competencias_presc = competencias_list[-meses_presc:] if competencias_list else []

    dias_uteis_calendario = sum(c["dias_uteis"] for c in competencias_presc)
    dias_uteis_efetivos = explicit_worked_days if explicit_worked_days > 0 else dias_uteis_calendario

    periodo_str = f"{adm.strftime('%d/%m/%Y')} a {dism.strftime('%d/%m/%Y')}"
    verbas: List[Dict[str, Any]] = []
    ocorrencias_mensais: List[Dict[str, Any]] = []

    def add_verba(
        cod: str,
        nome: str,
        base: Decimal,
        formula: str,
        qtd: str,
        pct: str,
        orig: Decimal,
        refl: Decimal = Decimal("0.00"),
        inc_fgts: bool = True,
        distribuir_mensal: bool = False,
    ) -> None:
        orig_q = orig.quantize(CENT, rounding=ROUND_HALF_UP)
        refl_q = refl.quantize(CENT, rounding=ROUND_HALF_UP)
        tot_q = (orig_q + refl_q).quantize(CENT, rounding=ROUND_HALF_UP)
        if tot_q <= Decimal("0.00"):
            return
        verbas.append({
            "Código": cod,
            "Verba": nome,
            "Período": periodo_str,
            "Base de Cálculo": base,
            "Fórmula": formula,
            "Quantidade": qtd,
            "Percentual": pct,
            "Valor Original": orig_q,
            "Reflexos": refl_q,
            "Total Apurado": tot_q,
            "Incide FGTS": inc_fgts,
        })
        if distribuir_mensal and competencias_presc:
            tot_dias_comp = max(1, sum(c["dias_uteis"] for c in competencias_presc))
            acum_dist = Decimal("0.00")
            for idx_c, comp in enumerate(competencias_presc):
                if idx_c == len(competencias_presc) - 1:
                    parcela = (tot_q - acum_dist).quantize(CENT, rounding=ROUND_HALF_UP)
                else:
                    peso = Decimal(str(comp["dias_uteis"])) / Decimal(str(tot_dias_comp))
                    parcela = (tot_q * peso).quantize(CENT, rounding=ROUND_HALF_UP)
                    acum_dist += parcela
                ocorrencias_mensais.append({
                    "Competência": comp["competencia"],
                    "Código": cod,
                    "Verba": nome,
                    "Dias Úteis": comp["dias_uteis"],
                    "Base de Cálculo (R$)": base,
                    "Devido (R$)": parcela,
                    "Pago (R$)": Decimal("0.00"),
                    "Diferença (R$)": parcela,
                })
        else:
            ocorrencias_mensais.append({
                "Competência": dism.strftime("%m/%Y"),
                "Código": cod,
                "Verba": nome,
                "Dias Úteis": dias_uteis_efetivos,
                "Base de Cálculo (R$)": base,
                "Devido (R$)": tot_q,
                "Pago (R$)": Decimal("0.00"),
                "Diferença (R$)": tot_q,
            })

    # 1. Saldo de Salário
    if params.get("inc_saldo_salario"):
        dias_s = max(0, min(31, f_int(params.get("saldo_dias"), dism.day)))
        val_saldo = ((remuneracao_base / Decimal("30")) * Decimal(str(dias_s))).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-01", "Saldo de Salário", remuneracao_base, "(Remuneração / 30) × Dias", f"{dias_s} dias", "100%", val_saldo)

    # 2. Aviso-Prévio Indenizado (Lei 12.506/2011)
    if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "acordo"):
        fator_ap = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
        val_ap = ((remuneracao_base / Decimal("30")) * Decimal(str(dias_aviso)) * fator_ap).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-02", "Aviso-Prévio Indenizado (Lei 12.506/2011)", remuneracao_base, "(Remuneração / 30) × Dias Aviso", f"{dias_aviso} dias", f"{int(fator_ap * 100)}%", val_ap)

    # 3. 13º Salário Proporcional + Projeção
    if params.get("inc_13") and term_type != "justa_causa":
        meses_13 = min(12, max(1, dism.month + (1 if dism.day >= 15 else 0) + (1 if params.get("inc_aviso_previo") else 0)))
        val_13 = ((remuneracao_base / Decimal("12")) * Decimal(str(meses_13))).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-03", "13º Salário Proporcional (c/ Projeção)", remuneracao_base, "(Remuneração / 12) × Avos", f"{meses_13}/12 avos", "100%", val_13)

    # 4. Férias Proporcionais + Vencidas + 1/3 Constitucional
    if params.get("inc_ferias") and term_type != "justa_causa":
        avos_fer = min(12, max(1, (meses_contrato % 12) or 12))
        val_fer_prop = ((remuneracao_base / Decimal("12")) * Decimal(str(avos_fer))).quantize(CENT, rounding=ROUND_HALF_UP)
        terco_prop = (val_fer_prop / Decimal("3")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-04A", "Férias Proporcionais + 1/3 Constitucional", remuneracao_base, "((Remuneração / 12) × Avos) + 1/3", f"{avos_fer}/12 avos", "133,33%", val_fer_prop, terco_prop, inc_fgts=False)

    fer_venc = f_int(params.get("periodos_ferias_vencidas"), 0)
    if fer_venc > 0:
        val_fv = (remuneracao_base * Decimal(str(fer_venc))).quantize(CENT, rounding=ROUND_HALF_UP)
        terco_fv = (val_fv / Decimal("3")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-04B", "Férias Vencidas (em dobro) + 1/3", remuneracao_base, "Remuneração × Períodos + 1/3", f"{fer_venc} per.", "133,33%", val_fv, terco_fv, inc_fgts=False)

    # 5. Adicional de Insalubridade / Periculosidade Retroativo
    if params.get("inc_adicional_retroativo") and insal_mensal > Decimal("0.00"):
        val_adic_ret = (insal_mensal * Decimal(str(meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_adic = (val_adic_ret * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP) if params.get("inc_reflexos_completos") else Decimal("0.00")
        add_verba("VER-05", f"{params.get('adicional_label', 'Adicional')} Retroativo", insal_mensal, "Adicional Mensal × Meses + Reflexos", f"{meses_presc} meses", str(params.get("adicional_pct_label", "")), val_adic_ret, refl_adic, distribuir_mensal=True)

    # 6. Plus Acúmulo de Função / Equiparação Salarial Retroativa
    if plus_acumulo_val > Decimal("0.00"):
        val_ac = (plus_acumulo_val * Decimal(str(meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_ac = (val_ac * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05B", "Plus Salarial por Acúmulo/Desvio de Função", plus_acumulo_val, "Plus Mensal × Meses + Reflexos", f"{meses_presc} meses", f"{params.get('plus_acumulo_pct', 0)}%", val_ac, refl_ac, distribuir_mensal=True)

    if dif_equip > Decimal("0.00"):
        val_eq = (dif_equip * Decimal(str(meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_eq = (val_eq * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05C", "Diferenças por Equiparação Salarial (Art. 461 CLT)", dif_equip, "Diferença Mensal × Meses + Reflexos", f"{meses_presc} meses", "100%", val_eq, refl_eq, distribuir_mensal=True)

    # 7. Horas Extras 50%, 100%, Intrajornada e Adicional Noturno
    fator_refl_jornada = Decimal("0.3935") if params.get("inc_reflexos_completos") else Decimal("0.1667")

    if params.get("inc_he50") and f_float(params.get("he50_hours"), 0.0) > 0:
        h50 = Decimal(str(f_float(params.get("he50_hours"), 0.0)))
        val_he50 = (valor_hora * Decimal("1.50") * h50).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he50 = (val_he50 * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06A", "Horas Extras 50% + DSR e Reflexos", remuneracao_base, "(Valor Hora × 1,50) × Horas + DSR/Reflexos", f"{h50}h", "+50%", val_he50, refl_he50, distribuir_mensal=True)

    if params.get("inc_he100") and f_float(params.get("he100_hours"), 0.0) > 0:
        h100 = Decimal(str(f_float(params.get("he100_hours"), 0.0)))
        val_he100 = (valor_hora * Decimal("2.00") * h100).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he100 = (val_he100 * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06B", "Horas Extras 100% (Domingos/Feriados) + Reflexos", remuneracao_base, "(Valor Hora × 2,00) × Horas + DSR/Reflexos", f"{h100}h", "+100%", val_he100, refl_he100, distribuir_mensal=True)

    if params.get("inc_intrajornada") and f_float(params.get("intrajornada_hours"), 0.0) > 0:
        hintra = Decimal(str(f_float(params.get("intrajornada_hours"), 0.0)))
        val_intra = (valor_hora * Decimal("1.50") * hintra).quantize(CENT, rounding=ROUND_HALF_UP)
        teto_intra = d_money(params.get("teto_intrajornada"))
        formula_intra = f"(Valor Hora × 1,50) × {hintra}h ({dias_uteis_efetivos} dias úteis, escala {work_days_week}d/sem)"
        if teto_intra > Decimal("0.00") and val_intra > teto_intra:
            val_intra = teto_intra
            formula_intra = f"Travado no Teto da Sentença (Art. 879, § 1º CLT / IAC TST — {dias_uteis_efetivos} dias úteis)"
        refl_intra = (val_intra * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP) if params.get("inc_reflexos_intrajornada") else Decimal("0.00")
        add_verba("VER-06C", "Intervalo Intrajornada Suprimido (Art. 71, §4º CLT)", remuneracao_base, formula_intra, f"{hintra}h ({dias_uteis_efetivos}d)", "+50%", val_intra, refl_intra, inc_fgts=bool(params.get("inc_reflexos_intrajornada")), distribuir_mensal=True)

    if params.get("inc_noturno") and f_float(params.get("night_hours"), 0.0) > 0:
        hnot = Decimal(str(f_float(params.get("night_hours"), 0.0)))
        val_not = (valor_hora * Decimal("0.20") * hnot).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_not = (val_not * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06D", "Adicional Noturno (20% + Hora Ficta) + Reflexos", remuneracao_base, "(Valor Hora × 20%) × Horas + Reflexos", f"{hnot}h", "+20%", val_not, refl_not, distribuir_mensal=True)

    # 8. Estabilidade Provisória Indenizada
    meses_estab = f_int(params.get("meses_estabilidade"), 0)
    if meses_estab > 0:
        val_estab = (remuneracao_base * Decimal(str(meses_estab))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_estab = (val_estab * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-07", "Indenização Estabilidade Provisória (Salários + 13º/Férias)", remuneracao_base, "Remuneração × Meses Estabilidade", f"{meses_estab} meses", "100%", val_estab, refl_estab)

    # 9. FGTS (8%) + Multa Rescisória
    if params.get("inc_fgts"):
        base_fgts_verbas = sum((v["Total Apurado"] for v in verbas if v.get("Incide FGTS")), Decimal("0.00"))
        fgts_sobre_verbas = (base_fgts_verbas * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP)
        fgts_contrato_base = (remuneracao_base * Decimal(str(meses_contrato)) * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP)
        fgts_faltante = d_money(params.get("fgts_nao_depositado"))
        multa_pct = Decimal(str(f_float(params.get("fgts_multa_pct"), 40.0))) / Decimal("100")
        base_multa = fgts_contrato_base + fgts_sobre_verbas + fgts_faltante
        val_multa_fgts = (base_multa * multa_pct).quantize(CENT, rounding=ROUND_HALF_UP)
        tot_fgts_item = fgts_sobre_verbas + fgts_faltante
        add_verba(
            "VER-08",
            f"FGTS (8%) + Multa Rescisória ({params.get('fgts_multa_pct', 40):g}%)",
            base_multa,
            "FGTS s/ Verbas + Faltante + Multa Rescisória",
            f"{meses_contrato} meses",
            f"{params.get('fgts_multa_pct', 40):g}%",
            tot_fgts_item,
            val_multa_fgts,
            inc_fgts=False,
        )

    # 10. Multas Arts. 477 e 467 CLT
    if params.get("inc_multa_477"):
        add_verba("VER-09", "Multa do Art. 477, § 8º da CLT (Atraso Verbas Rescisórias)", remuneracao_base, "1 Remuneração Mensal", "1 mês", "100%", remuneracao_base, inc_fgts=False)

    if params.get("inc_multa_467"):
        base_467 = sum((v["Total Apurado"] for v in verbas if v["Código"] in ("VER-01", "VER-02", "VER-03", "VER-04A")), Decimal("0.00"))
        val_467 = (base_467 * Decimal("0.50")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-10", "Multa do Art. 467 da CLT (50% sobre Verbas Incontroversas)", base_467, "50% × Verbas Rescisórias Incontroversas", "1", "50%", val_467, inc_fgts=False)

    # Restituição de Descontos Indevidos (antes da atualização e honorários)
    restit_val = d_money(params.get("restituicao_descontos_val"))
    if restit_val > Decimal("0.00"):
        add_verba("VER-11", "Restituição de Descontos Indevidos (Art. 462 CLT)", restit_val, "Devolução integral conforme título/comprovantes", "1", "100%", restit_val, inc_fgts=False)

    bruto_total = sum((v["Total Apurado"] for v in verbas), Decimal("0.00")).quantize(CENT, rounding=ROUND_HALF_UP)
    teto_global = d_money(params.get("teto_global_condenacao"))
    if teto_global > Decimal("0.00") and bruto_total > teto_global:
        bruto_total = teto_global

    abatimentos = d_money(params.get("offsets"))
    liquido_historico = max(Decimal("0.00"), bruto_total - abatimentos).quantize(CENT, rounding=ROUND_HALF_UP)

    atualizacao_pct = Decimal(str(f_float(params.get("atualizacao_juros_pct"), 0.0))) / Decimal("100")
    val_atualizacao = (liquido_historico * atualizacao_pct).quantize(CENT, rounding=ROUND_HALF_UP)
    if val_atualizacao > Decimal("0.00"):
        add_verba("VER-12", f"Atualização Monetária + Juros Legais ({params.get('atualizacao_juros_pct', 0):g}%)", liquido_historico, "Principal Líquido × Índice Acumulado (ADC 58 STF)", "1", f"{params.get('atualizacao_juros_pct', 0):g}%", val_atualizacao, inc_fgts=False)

    liquido_corrigido = (liquido_historico + val_atualizacao).quantize(CENT, rounding=ROUND_HALF_UP)

    hon_pct = Decimal(str(f_float(params.get("honorarios_sucumbenciais_pct"), 0.0))) / Decimal("100")
    val_honorarios = (liquido_corrigido * hon_pct).quantize(CENT, rounding=ROUND_HALF_UP)
    if val_honorarios > Decimal("0.00"):
        add_verba("VER-13", f"Honorários Advocatícios Sucumbenciais ({params.get('honorarios_sucumbenciais_pct', 0):g}%)", liquido_corrigido, "Total Líquido Atualizado × % Sucumbência", "1", f"{params.get('honorarios_sucumbenciais_pct', 0):g}%", val_honorarios, inc_fgts=False)

    liquido_devido = (liquido_corrigido + val_honorarios).quantize(CENT, rounding=ROUND_HALF_UP)

    # Atualiza colunas de correção nas ocorrências mensais proporcionalmente
    fator_corr_occ = (Decimal("1.00") + atualizacao_pct)
    for oc in ocorrencias_mensais:
        dif_oc = d_money(oc["Diferença (R$)"])
        corr_oc = (dif_oc * atualizacao_pct).quantize(CENT, rounding=ROUND_HALF_UP)
        oc["Correção/Juros (R$)"] = corr_oc
        oc["Total Corrigido (R$)"] = (dif_oc + corr_oc).quantize(CENT, rounding=ROUND_HALF_UP)

    raw_hash = hashlib.sha256(f"{params.get('worker_name')}|{params.get('employer_name')}|{liquido_devido}|{datetime.utcnow().isoformat()}".encode()).hexdigest()[:8].upper()
    audit_id = f"ED-{raw_hash}"

    return {
        "audit_id": audit_id,
        "process_number": params.get("process_number", ""),
        "office_header": params.get("office_header", "ED Cálculos Trabalhistas"),
        "worker_name": params.get("worker_name", ""),
        "employer_name": params.get("employer_name", ""),
        "role_title": params.get("role_title", ""),
        "admission_date": adm.isoformat(),
        "dismissal_date": dism.isoformat(),
        "termination_type": term_type,
        "monthly_divisor": int(divisor),
        "work_days_per_week": work_days_week,
        "dias_uteis_efetivos": dias_uteis_efetivos,
        "base_remuneracao": remuneracao_base,
        "verbas": verbas,
        "ocorrencias_mensais": ocorrencias_mensais,
        "bruto_atualizado": bruto_total,
        "abatimentos": abatimentos,
        "liquido_historico": liquido_historico,
        "valor_atualizacao": val_atualizacao,
        "valor_honorarios": val_honorarios,
        "liquido_devido": liquido_devido,
        "ai_opinion": params.get("ai_opinion") or {},
        "lawyer_narrative": params.get("lawyer_narrative") or "",
        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }


# ==============================================================================
# 5. EXPORTAÇÃO PADRÃO PJE-CALC (.XLSX MULTI-ABAS, .PDF DEMONSTRATIVO E .PJC XML)
# ==============================================================================
def build_complete_excel_bytes(calc: Dict[str, Any]) -> bytes:
    buf = io.BytesIO()
    resumo_rows = [
        {"Campo": "Padrão Metodológico", "Valor": "Demonstrativo de Liquidação — Padrão PJe-Calc / TRT"},
        {"Campo": "Código Auditoria", "Valor": calc.get("audit_id", "")},
        {"Campo": "Processo / Referência", "Valor": calc.get("process_number", "")},
        {"Campo": "Reclamante / Exequente", "Valor": calc.get("worker_name", "")},
        {"Campo": "Reclamada / Executada", "Valor": calc.get("employer_name", "")},
        {"Campo": "Período Contratual", "Valor": f"{calc.get('admission_date')} a {calc.get('dismissal_date')}"},
        {"Campo": "Escala / Dias Úteis Apurados", "Valor": f"{calc.get('work_days_per_week', 5)} dias/semana • {calc.get('dias_uteis_efetivos', 0)} dias úteis"},
        {"Campo": "Divisor / Remuneração Base", "Valor": f"Divisor {calc.get('monthly_divisor', 220)} • {fmt_brl(calc.get('base_remuneracao'))}"},
        {"Campo": "Total Principal Histórico", "Valor": fmt_brl(calc.get("bruto_atualizado"))},
        {"Campo": "(-) Abatimentos / Pagos", "Valor": fmt_brl(calc.get("abatimentos"))},
        {"Campo": "Atualização Monetária + Juros (ADC 58 STF)", "Valor": fmt_brl(calc.get("valor_atualizacao"))},
        {"Campo": "Honorários Advocatícios Sucumbenciais", "Valor": fmt_brl(calc.get("valor_honorarios"))},
        {"Campo": "TOTAL LÍQUIDO GERAL DA EXECUÇÃO", "Valor": fmt_brl(calc.get("liquido_devido"))},
    ]
    verbas_rows = []
    for v in calc.get("verbas", []):
        verbas_rows.append({
            "Código": v["Código"],
            "Verba": v["Verba"],
            "Período": v["Período"],
            "Quantidade": v["Quantidade"],
            "Percentual": v["Percentual"],
            "Base de Cálculo (R$)": float(d_money(v["Base de Cálculo"])),
            "Fórmula": v["Fórmula"],
            "Valor Original (R$)": float(d_money(v["Valor Original"])),
            "Reflexos (R$)": float(d_money(v["Reflexos"])),
            "Total Apurado (R$)": float(d_money(v["Total Apurado"])),
        })

    occ_rows = []
    for oc in calc.get("ocorrencias_mensais", []):
        occ_rows.append({
            "Competência": oc.get("Competência", ""),
            "Código": oc.get("Código", ""),
            "Verba": oc.get("Verba", ""),
            "Dias Úteis": oc.get("Dias Úteis", 0),
            "Base de Cálculo (R$)": float(d_money(oc.get("Base de Cálculo (R$)"))),
            "Devido (R$)": float(d_money(oc.get("Devido (R$)"))),
            "Pago (R$)": float(d_money(oc.get("Pago (R$)"))),
            "Diferença (R$)": float(d_money(oc.get("Diferença (R$)"))),
            "Correção/Juros (R$)": float(d_money(oc.get("Correção/Juros (R$)"))),
            "Total Corrigido (R$)": float(d_money(oc.get("Total Corrigido (R$)"))),
        })

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(resumo_rows).to_excel(writer, sheet_name="Resumo_PJeCalc", index=False)
        pd.DataFrame(verbas_rows).to_excel(writer, sheet_name="Verbas_Consolidado", index=False)
        if occ_rows:
            pd.DataFrame(occ_rows).to_excel(writer, sheet_name="Ocorrencias_Mes_a_Mes", index=False)
    return buf.getvalue()


def build_pjecalc_xml_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera arquivo estruturado .PJC (XML compatível com estrutura de importação PJe-Calc Cidadão)."""
    root = ET.Element("CalculoPJeCalc", {
        "versaoSistema": TABLES_VERSION,
        "codigoAuditoria": str(calc.get("audit_id", "")),
        "dataLiquidacao": datetime.now().strftime("%Y-%m-%d"),
    })
    dados_proc = ET.SubElement(root, "DadosProcesso")
    ET.SubElement(dados_proc, "NumeroProcesso").text = str(calc.get("process_number", ""))
    ET.SubElement(dados_proc, "Reclamante").text = str(calc.get("worker_name", ""))
    ET.SubElement(dados_proc, "Reclamada").text = str(calc.get("employer_name", ""))
    ET.SubElement(dados_proc, "Admissao").text = str(calc.get("admission_date", ""))
    ET.SubElement(dados_proc, "Demissao").text = str(calc.get("dismissal_date", ""))
    ET.SubElement(dados_proc, "Divisor").text = str(calc.get("monthly_divisor", 220))
    ET.SubElement(dados_proc, "RemuneracaoBase").text = str(d_money(calc.get("base_remuneracao")))

    verbas_el = ET.SubElement(root, "Verbas")
    for v in calc.get("verbas", []):
        v_el = ET.SubElement(verbas_el, "Verba", {"codigo": str(v["Código"])})
        ET.SubElement(v_el, "Descricao").text = str(v["Verba"])
        ET.SubElement(v_el, "Periodo").text = str(v["Período"])
        ET.SubElement(v_el, "Quantidade").text = str(v["Quantidade"])
        ET.SubElement(v_el, "ValorOriginal").text = str(d_money(v["Valor Original"]))
        ET.SubElement(v_el, "Reflexos").text = str(d_money(v["Reflexos"]))
        ET.SubElement(v_el, "TotalApurado").text = str(d_money(v["Total Apurado"]))

    occ_el = ET.SubElement(root, "OcorrenciasMensais")
    for oc in calc.get("ocorrencias_mensais", []):
        o_el = ET.SubElement(occ_el, "Ocorrencia", {
            "competencia": str(oc.get("Competência", "")),
            "codigoVerba": str(oc.get("Código", "")),
        })
        ET.SubElement(o_el, "DiasUteis").text = str(oc.get("Dias Úteis", 0))
        ET.SubElement(o_el, "Devido").text = str(d_money(oc.get("Devido (R$)")))
        ET.SubElement(o_el, "Diferenca").text = str(d_money(oc.get("Diferença (R$)")))
        ET.SubElement(o_el, "TotalCorrigido").text = str(d_money(oc.get("Total Corrigido (R$)")))

    resumo_el = ET.SubElement(root, "ResumoLiquidacao")
    ET.SubElement(resumo_el, "BrutoHistorico").text = str(d_money(calc.get("bruto_atualizado")))
    ET.SubElement(resumo_el, "Abatimentos").text = str(d_money(calc.get("abatimentos")))
    ET.SubElement(resumo_el, "CorrecaoEJuros").text = str(d_money(calc.get("valor_atualizacao")))
    ET.SubElement(resumo_el, "HonorariosSucumbenciais").text = str(d_money(calc.get("valor_honorarios")))
    ET.SubElement(resumo_el, "LiquidoDevido").text = str(d_money(calc.get("liquido_devido")))

    xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return xml_str


def _pdf_safe(txt: Any) -> str:
    return str(txt or "").encode("latin-1", errors="replace").decode("latin-1")


def build_printable_dossier_pdf_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera PDF no layout clássico de Demonstrativo de Cálculo Padrão PJe-Calc."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    # Cabeçalho institucional estilo PJe-Calc
    pdf.set_fill_color(15, 37, 55)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, _pdf_safe("DEMONSTRATIVO DE LIQUIDAÇÃO TRABALHISTA — PADRÃO PJE-CALC / CLT"), 0, 1, "C", True)

    pdf.set_text_color(15, 37, 55)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, _pdf_safe(calc.get("office_header") or "ED Cálculos Trabalhistas"), ln=True)
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(0, 5, _pdf_safe(f"Código Auditoria: {calc.get('audit_id')} | Data Liquidação: {calc.get('created_at')} | Tabelas: {TABLES_VERSION}"), ln=True)
    pdf.ln(2)

    # Quadro de Identificação do Processo e Parâmetros
    pdf.set_fill_color(240, 244, 248)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 6, _pdf_safe("1. DADOS DO PROCESSO E PARÂMETROS DO CÁLCULO"), 1, 1, "L", True)
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(95, 5.5, _pdf_safe(f"Processo: {calc.get('process_number')}"), 1)
    pdf.cell(95, 5.5, _pdf_safe(f"Período: {calc.get('admission_date')} a {calc.get('dismissal_date')}"), 1, 1)
    pdf.cell(95, 5.5, _pdf_safe(f"Reclamante: {calc.get('worker_name')}"), 1)
    pdf.cell(95, 5.5, _pdf_safe(f"Reclamada: {calc.get('employer_name')}"), 1, 1)
    pdf.cell(95, 5.5, _pdf_safe(f"Divisor: {calc.get('monthly_divisor', 220)} | Escala: {calc.get('work_days_per_week', 5)}d/sem ({calc.get('dias_uteis_efetivos', 0)} dias úteis)"), 1)
    pdf.cell(95, 5.5, _pdf_safe(f"Remuneração Base: {fmt_brl(calc.get('base_remuneracao'))}"), 1, 1)
    pdf.ln(3)

    # Quadro 2: Planilha Resumo das Verbas
    pdf.set_fill_color(15, 37, 55)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(16, 6.5, "Cod.", 1, 0, "L", True)
    pdf.cell(84, 6.5, "Descrição da Verba / Critério", 1, 0, "L", True)
    pdf.cell(30, 6.5, "Valor Original", 1, 0, "R", True)
    pdf.cell(30, 6.5, "Reflexos", 1, 0, "R", True)
    pdf.cell(30, 6.5, "Total Apurado", 1, 1, "R", True)

    pdf.set_text_color(15, 37, 55)
    pdf.set_font("Helvetica", "", 7.5)
    for v in calc.get("verbas", []):
        pdf.cell(16, 5.5, _pdf_safe(v["Código"]), 1)
        pdf.cell(84, 5.5, _pdf_safe(str(v["Verba"])[:48]), 1)
        pdf.cell(30, 5.5, _pdf_safe(fmt_brl(v["Valor Original"])), 1, 0, "R")
        pdf.cell(30, 5.5, _pdf_safe(fmt_brl(v["Reflexos"])), 1, 0, "R")
        pdf.cell(30, 5.5, _pdf_safe(fmt_brl(v["Total Apurado"])), 1, 1, "R")

    pdf.ln(3)
    # Quadro 3: Fechamento Resumo do Cálculo
    pdf.set_fill_color(240, 244, 248)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(0, 6, _pdf_safe("2. RESUMO GERAL DO CRÉDITO EXEQUENDO"), 1, 1, "L", True)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.cell(130, 5.5, _pdf_safe("Total Principal Histórico das Verbas:"), 1)
    pdf.cell(60, 5.5, _pdf_safe(fmt_brl(calc.get("bruto_atualizado"))), 1, 1, "R")
    pdf.cell(130, 5.5, _pdf_safe("(-) Deduções / Valores Comprovadamente Pagos:"), 1)
    pdf.cell(60, 5.5, _pdf_safe(f"-{fmt_brl(calc.get('abatimentos'))}"), 1, 1, "R")
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(130, 6.5, _pdf_safe("TOTAL LÍQUIDO GERAL DEVIDO (C/ ATUALIZAÇÃO E HONORÁRIOS):"), 1)
    pdf.cell(60, 6.5, _pdf_safe(fmt_brl(calc.get("liquido_devido"))), 1, 1, "R")

    # Página 2: Demonstrativo Mês a Mês por Competência (PJe-Calc)
    occs = calc.get("ocorrencias_mensais") or []
    if occs:
        pdf.add_page()
        pdf.set_fill_color(15, 37, 55)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 7, _pdf_safe("3. DEMONSTRATIVO DE OCORRÊNCIAS POR COMPETÊNCIA (MÊS A MÊS — PADRÃO PJE-CALC)"), 0, 1, "L", True)
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(22, 6, "Compet.", 1, 0, "C", True)
        pdf.cell(74, 6, "Verba", 1, 0, "L", True)
        pdf.cell(18, 6, "Dias", 1, 0, "C", True)
        pdf.cell(25, 6, "Devido", 1, 0, "R", True)
        pdf.cell(25, 6, "Corr./Juros", 1, 0, "R", True)
        pdf.cell(26, 6, "Total Corr.", 1, 1, "R", True)

        pdf.set_text_color(15, 37, 55)
        pdf.set_font("Helvetica", "", 7)
        for oc in occs[:120]:
            pdf.cell(22, 5, _pdf_safe(oc.get("Competência", "")), 1, 0, "C")
            pdf.cell(74, 5, _pdf_safe(str(oc.get("Verba", ""))[:42]), 1)
            pdf.cell(18, 5, _pdf_safe(str(oc.get("Dias Úteis", ""))), 1, 0, "C")
            pdf.cell(25, 5, _pdf_safe(fmt_brl(oc.get("Devido (R$)"))), 1, 0, "R")
            pdf.cell(25, 5, _pdf_safe(fmt_brl(oc.get("Correção/Juros (R$)"))), 1, 0, "R")
            pdf.cell(26, 5, _pdf_safe(fmt_brl(oc.get("Total Corrigido (R$)"))), 1, 1, "R")

    raw_out = pdf.output(dest="S")
    if isinstance(raw_out, str):
        return raw_out.encode("latin-1")
    return bytes(raw_out)
# ... existing code ...
def analyze_sentence_with_gemini(notes: str, files: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = _get_gemini_client()
    prompt = (
        "Você é um Perito Calculista Trabalhista especialista em Liquidação de Sentença no PJe-Calc (Art. 879, § 1º CLT e IAC-1000156-59.2020.5.09.0000 TST). "
        "Analise a cópia do processo/sentença transitada em julgado e documentos anexados. "
        "REGRAS OBRIGATÓRIAS DE EXTRAÇÃO:\n"
        "1) DATAS DO VÍNCULO: Extraia estritamente a data de ADMISSÃO e a data de RESCISÃO/TÉRMINO DO CONTRATO DE TRABALHO (CTPS/TRCT/fundamentação). Nunca confunda com data de ajuizamento, despacho ou sentença!\n"
        "2) ESCALA E DIAS TRABALHADOS: Identifique se a jornada reconhecida era Segunda a Sexta (work_days_per_week = 5) ou Segunda a Sábado (work_days_per_week = 6), e se há número exato de dias trabalhados mencionados nos autos (worked_days_exact, ou 0 para calcular pelo calendário).\n"
        "3) COISA JULGADA / TETOS DA INICIAL: Verifique se a sentença limitou a condenação aos valores indicados na petição inicial (tetos históricos). Se houver teto expresso (ex.: intervalo intrajornada limitado a R$ 1.337,60, descontos limitados a R$ 2.000,00), preencha os campos de teto!\n"
        "4) DISPOSITIVO: Separe rigorosamente o que foi DEFERIDO (descartando os pedidos improcedentes).\n"
        "Retorne EXCLUSIVAMENTE um JSON válido com as chaves: "
        "process_number (str), worker_name (str), employer_name (str), role_title (str), "
        "admission_date (YYYY-MM-DD ou null), dismissal_date (YYYY-MM-DD ou null), "
        "work_days_per_week (int: 5 para Seg-Sex ou 6 para Seg-Sab), worked_days_exact (int: dias exatos trabalhados se constar nos autos, senão 0), "
        "termination_type (sem_justa_causa|rescisao_indireta|pedido_demissao|acordo|justa_causa), monthly_divisor (int), "
        "base_salary (float), commissions (float), adicional_sugerido (Nenhum|Insalubridade|Periculosidade), "
        "adicional_percentual_sugerido (float), trct_paid_total (float), "
        "teto_intrajornada (float: valor teto fixado na sentença para intervalo intrajornada, ou 0 se ilimitado), "
        "incluir_reflexos_intrajornada (bool), "
        "teto_global_condenacao (float: teto global se a sentença limitou o total, ou 0), "
        "honorarios_sucumbenciais_pct (float: ex 15.0 se deferidos honorários ao patrono do autor, ou 0), "
        "indice_correcao_juros_estimado_pct (float: percentual estimado de correção + juros até hoje, ex 12.0), "
        "dispositivo_resumo (str), parametros_liquidacao (str), divergencias_calculo_adverso (str ou null), "
        "verbas_deferidas (lista de strings), verbas_indeferidas (lista de strings), "
        "weekly_he50_hours (float), monthly_sundays_worked (int), suppressed_break_minutes (int), weekly_night_hours (float), "
        "sugerir_saldo_salario (bool), sugerir_aviso_previo (bool), sugerir_13_e_ferias (bool), "
        "sugerir_fgts_e_multa (bool), multa_fgts_percentual_sugerido (float), fgts_nao_depositado_contrato (float), "
        "meses_estabilidade_restantes (int), sugerir_multa_477 (bool), sugerir_multa_467 (bool), "
        "restituicao_descontos (float), minuta_peticao_liquidacao (str pronta para protocolar no PJe).\n\n"
        f"OBSERVAÇÕES DO ADVOGADO:\n{notes}"
    )
    contents: List[Any] = [prompt]
    for f in files:
        contents.append(types.Part.from_bytes(data=f["bytes"], mime_type=f["mime"]))

    resp = client.models.generate_content(
        model=GEMINI_MODEL_NAME,
        contents=contents,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1),
    )
    return json.loads(resp.text)
# ... existing code ...
        sd1, sd2, sd3, sd4 = st.columns(4)
        s_adm = sd1.date_input("Admissão Reconhecida *", value=_parse_dt_sent(s_ext.get("admission_date"), date(2024, 2, 1)), key="s_adm_dt")
        s_dism = sd2.date_input("Rescisão Reconhecida *", value=_parse_dt_sent(s_ext.get("dismissal_date"), date(2026, 10, 31)), key="s_dism_dt")
        s_motivo = sd3.selectbox("Modalidade da Rescisão", s_motivos, index=s_motivos.index(s_mot_ext))
        s_divisor = sd4.number_input("Divisor Fixado", min_value=1, max_value=300, value=max(1, min(300, f_int(s_ext.get("monthly_divisor"), 220))))

        sw1, sw2 = st.columns(2)
        s_escala_opts = [5, 6]
        s_escala_default = 5 if f_int(s_ext.get("work_days_per_week"), 5) <= 5 else 6
        s_work_days_week = sw1.selectbox(
            "Escala Semanal Reconhecida (PJe-Calc)",
            s_escala_opts,
            index=s_escala_opts.index(s_escala_default),
            format_func=lambda x: "5 dias/semana (Segunda a Sexta-feira)" if x == 5 else "6 dias/semana (Segunda a Sábado)",
        )
        s_worked_days_exact = sw2.number_input(
            "Total Exato de Dias Trabalhados no Período (0 = apurar automático no calendário)",
            min_value=0,
            max_value=2000,
            value=max(0, f_int(s_ext.get("worked_days_exact"), 0)),
        )

        sr1, sr2, sr3, sr4, sr5 = st.columns([2, 2, 2.2, 1.5, 2])
# ... existing code ...
            if s_tc is not None:
                tot_s_he50 = f_float(s_tc.get("he50_hours"), 0.0)
                tot_s_he100 = f_float(s_tc.get("he100_hours"), 0.0)
                tot_s_intra = f_float(s_tc.get("intrajornada_hours"), 0.0)
                tot_s_not = f_float(s_tc.get("night_ficta_hours"), 0.0)
            else:
                dias_uteis_liq = s_worked_days_exact if s_worked_days_exact > 0 else count_working_days_in_range(s_adm, s_dism, s_work_days_week)
                tot_s_he50 = float((Decimal(str(s_he50_sem)) * s_semanas).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_he100 = float((Decimal(str(s_dom_mes * 8)) * Decimal(str(s_meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_intra = float(((Decimal(str(s_intra_min)) / Decimal("60")) * Decimal(str(dias_uteis_liq))).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_not = float((Decimal(str(s_not_sem)) * s_semanas * Decimal("1.142857")).quantize(CENT, rounding=ROUND_HALF_UP))

            s_sm_vigente = get_historical_min_wage(s_dism)
# ... existing code ...
            res_sent = run_labor_calculation({
                "process_number": f"LIQ-{s_proc.strip()}",
                "office_header": s_timbre,
                "worker_name": s_reclamante,
                "employer_name": s_reclamada,
                "role_title": s_cargo,
                "admission_date": s_adm,
                "dismissal_date": s_dism,
                "termination_type": s_motivo,
                "monthly_divisor": s_divisor,
                "work_days_per_week": s_work_days_week,
                "explicit_worked_days": s_worked_days_exact,
                "base_salary": s_sal,
                "commissions": s_comis,
                "insalubrity": float(s_insal_mensal),
                "adicional_label": s_adic,
                "adicional_pct_label": s_adic_lbl,
                "inc_adicional_retroativo": s_insal_mensal > Decimal("0.00"),
                "offsets": s_deducao,
                "inc_saldo_salario": s_ped_saldo,
                "saldo_dias": s_dias_saldo,
                "inc_aviso_previo": s_ped_aviso,
                "inc_13": s_ped_13_fer,
                "inc_ferias": s_ped_13_fer,
                "periodos_ferias_vencidas": 0,
                "meses_sem_registro": 0,
                "plus_acumulo_pct": 0.0,
                "dif_equiparacao_mensal": 0.0,
                "inc_fgts": s_ped_fgts,
                "fgts_multa_pct": s_multa_fgts_pct,
                "fgts_nao_depositado": s_fgts_dif,
                "meses_estabilidade": s_estab_m,
                "inc_multa_477": s_ped_477,
                "inc_multa_467": s_ped_467,
                "inc_reflexos_completos": s_ped_reflexos,
                "inc_he50": tot_s_he50 > 0,
                "he50_hours": tot_s_he50,
                "he50_pct": 50,
                "inc_he100": tot_s_he100 > 0,
                "he100_hours": tot_s_he100,
                "inc_intrajornada": tot_s_intra > 0 or s_teto_intra > 0,
                "intrajornada_hours": tot_s_intra if tot_s_intra > 0 else (1.0 if s_teto_intra > 0 else 0.0),
                "teto_intrajornada": s_teto_intra,
                "inc_reflexos_intrajornada": s_refl_intra,
                "teto_global_condenacao": s_teto_global,
                "atualizacao_juros_pct": s_atu_juros_pct,
                "honorarios_sucumbenciais_pct": s_hon_suc_pct,
                "inc_noturno": tot_s_not > 0,
                "night_hours": tot_s_not,
                "night_pct": 20,
                "restituicao_descontos_val": s_restit,
                "ai_opinion": s_ext,
                "lawyer_narrative": sent_notes,
            })

            sent_client_key = f"{s_reclamante.strip().lower()}|{s_reclamada.strip().lower()}"
# ... existing code ...
            df_sent = pd.DataFrame(res_sent["verbas"])
            if not df_sent.empty:
                for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"]:
                    df_sent[col] = df_sent[col].apply(fmt_brl)
                st.dataframe(
                    df_sent[["Código", "Verba", "Quantidade", "Percentual", "Base de Cálculo", "Fórmula", "Valor Original", "Reflexos", "Total Apurado"]],
                    use_container_width=True,
                    hide_index=True,
                )

            if res_sent.get("ocorrencias_mensais"):
                with st.expander("Demonstrativo Mês a Mês por Competência (Padrão PJe-Calc)", expanded=False):
                    df_occ = pd.DataFrame(res_sent["ocorrencias_mensais"])
                    for mcol in ["Base de Cálculo (R$)", "Devido (R$)", "Pago (R$)", "Diferença (R$)", "Correção/Juros (R$)", "Total Corrigido (R$)"]:
                        if mcol in df_occ.columns:
                            df_occ[mcol] = df_occ[mcol].apply(fmt_brl)
                    st.dataframe(df_occ, use_container_width=True, hide_index=True)

            if s_ext.get("minuta_peticao_liquidacao"):
                with st.expander("Petição de Apresentação de Cálculos de Liquidação (Art. 879 CLT)", expanded=True):
                    st.text_area("Minuta pronta para protocolar no PJe:", value=str(s_ext["minuta_peticao_liquidacao"]), height=240)

            sdl1, sdl2, sdl3 = st.columns(3)
            sdl1.download_button(
                "Baixar Planilha PJe-Calc (.XLSX)",
                data=build_complete_excel_bytes(res_sent),
                file_name=f"Liquidacao_PJeCalc_{res_sent['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            sdl2.download_button(
                "Baixar Demonstrativo Oficial (.PDF)",
                data=build_printable_dossier_pdf_bytes(res_sent),
                file_name=f"Demonstrativo_PJeCalc_{res_sent['audit_id']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            sdl3.download_button(
                "Baixar Arquivo Estruturado (.PJC)",
                data=build_pjecalc_xml_bytes(res_sent),
                file_name=f"Calculo_{res_sent['audit_id']}.pjc",
                mime="application/xml",
                use_container_width=True,
            )
# ... existing code ...
                b1, b2, b3 = st.columns(3)
                b1.download_button(
                    "Exportar Excel PJe-Calc (.XLSX)",
                    data=build_complete_excel_bytes(calc),
                    file_name=f"Dossie_PJeCalc_{calc['audit_id']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_xlsx_{calc['audit_id']}_{idx}",
                    use_container_width=True,
                )
                b2.download_button(
                    "Exportar Demonstrativo (.PDF)",
                    data=build_printable_dossier_pdf_bytes(calc),
                    file_name=f"Demonstrativo_PJeCalc_{calc['audit_id']}.pdf",
                    mime="application/pdf",
                    key=f"dl_pdf_{calc['audit_id']}_{idx}",
                    use_container_width=True,
                )
                b3.download_button(
                    "Exportar Arquivo (.PJC)",
                    data=build_pjecalc_xml_bytes(calc),
                    file_name=f"Calculo_{calc['audit_id']}.pjc",
                    mime="application/xml",
                    key=f"dl_pjc_{calc['audit_id']}_{idx}",
                    use_container_width=True,
                )
