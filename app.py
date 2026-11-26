# ... existing code ...
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, getcontext
import io
import json
import os
import re
import sqlite3
from typing import Any, Dict, List, Optional, Tuple
import urllib.request
import uuid
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
# ... existing code ...
# ==============================================================================
# 0. PRECISÃO DECIMAL E VERSIONAMENTO DE REGRAS LEGAIS
# ==============================================================================
getcontext().prec = 28
getcontext().rounding = ROUND_HALF_UP

ULTIMA_COMPETENCIA_TABELA_OFICIAL = "2026-12"
HISTORY_DB_PATH = "juriscalc_pro_historico.db"

RULES_REGISTRY: List[Dict[str, str]] = [
# ... existing code ...
def make_trace_field(
    valor: Any,
    doc_name: str = "",
    page: int = 1,
    trecho: str = "",
    confianca: str = "BAIXA",
) -> Dict[str, Any]:
    return {
        "valor": valor,
        "doc_name": doc_name,
        "page": page,
        "trecho": trecho,
        "confianca": confianca,
    }


def extract_contracts_from_pages(pages_index: List[Dict[str, Any]], reclamado_hint: str = "") -> List[Dict[str, Any]]:
# ... existing code ...
def _default_empty_params() -> Dict[str, Any]:
    return {
        "processo": "",
        "numero_calculo": "1",
        "estado": "RS",
        "municipio": "GRAVATAI",
        "cidade_assinatura": "Gravataí",
        "vara": "VARA DO TRABALHO",
        "reclamante": "",
        "reclamado": "",
        "parte_representada": "Reclamante",
        "advogado_assinante": "",
        "oab_advogado": "",
        "admissao": None,
        "demissao": None,
        "ajuizamento": None,
        "liquidacao": date.today(),
        "salario_base": 0.0,
        "maior_remuneracao": 0.0,
        "ultima_remuneracao": 0.0,
        "regime_trabalho": "Tempo Integral",
        "modalidade_rescisao": "Dispensa sem justa causa",
        "carga_horaria": 220,
        "divisor": 220.0,
        "projetar_aviso_previo": False,
        "aplicar_prescricao_quinquenal": False,
        "sabado_dia_util": True,
        "criterio_correcao": "ADC 58 STF (IPCA-E pré-judicial + SELIC) + Lei 14.905/2024",
        "inss_patronal_pct": 20.0,
        "rat_sat_pct": 2.0,
        "terceiros_pct": 5.8,
        "tem_he50": False,
        "qtd_he50_mensal": 0.0,
        "reflexos_he50": False,
        "tem_he100": False,
        "qtd_he100_mensal": 0.0,
        "reflexos_he100": False,
        "tem_intrajornada": False,
        "qtd_intrajornada_mensal": 0.0,
        "teto_intrajornada": 0.0,
        "teto_intrajornada_escopo": "VERBA_PRINCIPAL",
        "reflexos_intrajornada": False,
        "tem_descontos_indevidos": False,
        "valor_descontos_indevidos": 0.0,
        "teto_descontos_indevidos": 0.0,
        "reflexos_em_fgts": False,
        "multa_477": False,
        "multa_467": False,
        "verbas_personalizadas": [],
        "honorarios_sucumbencia_pct": 0.0,
        "honorarios_juiz_pct": 0.0,
        "honorarios_reciprocos_pct": 0.0,
        "base_pedidos_improcedentes": 0.0,
        "custas_processuais": 0.0,
        "custas_ja_recolhidas": False,
        "valor_pago_abatimento": 0.0,
        "_contracts": [],
        "_selected_contract_id": "",
        "_traceability": {},
        "_decision_timeline": [],
        "_verbas_indeferidas": [],
        "_representation_conflict": False,
        "_representation_details": "",
        "_human_confirmed_ambiguities": False,
        "_ingestion": {},
    }
# ... existing code ...
    # 3. Contratos na CTPS e cruzamento autônomo de datas em todas as páginas
    contracts = extract_contracts_from_pages(pages_index, reclamado_hint=params.get("reclamado", ""))
    params["_contracts"] = contracts

    # Busca consolidada de datas em decisões judiciais / inicial / TRCT
    judicial_pages = [p for p in pages_index if p["doc_type"] in ("SENTENCA", "EMBARGOS_DECLARACAO", "ACORDAO", "PETICAO_INICIAL", "TRCT")]
    search_pool = judicial_pages if judicial_pages else pages_index

    adm_found: Optional[date] = None
    dem_found: Optional[date] = None
    aju_found: Optional[date] = None
    sal_found: float = 0.0

    # Seleção autônoma pela inteligência documental:
    # 1º: Contrato na CTPS cujo empregador coincide com o Reclamado da ação;
    # 2º: Se houver vários, o contrato mais recente compatível com a data de ajuizamento (prescrição bienal).
    matching_ctrs = [c for c in contracts if c.get("matches_reclamado")]
    candidate_ctrs = matching_ctrs if matching_ctrs else contracts
    if candidate_ctrs:
        def _ctr_sort_key(c_item: Dict[str, Any]) -> date:
            return parse_br_date_optional(c_item.get("admissao")) or date(1900, 1, 1)

        c_pref = sorted(candidate_ctrs, key=_ctr_sort_key)[-1]
        params["_selected_contract_id"] = c_pref["contrato_id"]
        if c_pref.get("matches_reclamado") or len(contracts) == 1:
            adm_found = parse_br_date_optional(c_pref["admissao"])
            dem_found = parse_br_date_optional(c_pref["demissao"])
            if c_pref["salario"]:
                sal_found = float(c_pref["salario"])
            if adm_found:
                trace["admissao"] = make_trace_field(fmt_date_br(adm_found), c_pref["origem_doc"], c_pref["pagina"], c_pref["trecho"], "ALTA")
            if dem_found:
                trace["demissao"] = make_trace_field(fmt_date_br(dem_found), c_pref["origem_doc"], c_pref["pagina"], c_pref["trecho"], "ALTA")
            if sal_found > 0:
                trace["salario_base"] = make_trace_field(sal_found, c_pref["origem_doc"], c_pref["pagina"], c_pref["trecho"], "ALTA")
# ... existing code ...
    params["admissao"] = adm_found
    params["demissao"] = dem_found
    params["ajuizamento"] = aju_found
    params["salario_base"] = sal_found
    params["maior_remuneracao"] = sal_found
    params["ultima_remuneracao"] = sal_found
# ... existing code ...
    # Honorários e custas determinados pelo Juiz na sentença/acórdão
    m_hon = re.search(
        r"honor[áa]rios\s+(?:advocat[íi]cios|sucumbenciais)[^\n\.]{0,110}?(\d{1,2}(?:[\.,]\d+)?)\s*%",
        decision_text,
        flags=re.IGNORECASE,
    )
    if m_hon:
        hon_det = float(D(m_hon.group(1)))
        params["honorarios_juiz_pct"] = hon_det
        params["honorarios_sucumbencia_pct"] = hon_det

    m_custas = re.search(
        r"custas\s+(?:processuais\s+)?(?:de|no\s+valor\s+de|arbitradas\s+em)\s*R\$\s*([\d\.]+,\d{2})",
        decision_text,
        flags=re.IGNORECASE,
    )
    if m_custas:
        params["custas_processuais"] = float(D(m_custas.group(1)))

    # Extração automática de município/comarca da vara trabalhista
    m_mun = re.search(r"VARA\s+DO\s+TRABALHO\s+DE\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]{3,40})", full_text, flags=re.IGNORECASE)
    if m_mun:
        mun_clean = re.sub(r"\s+", " ", m_mun.group(1)).strip().rstrip(".,;-")
        params["municipio"] = mun_clean.upper()
        params["cidade_assinatura"] = mun_clean.title()

    params["_traceability"] = trace
    params["_decision_timeline"] = timeline
    params["_verbas_indeferidas"] = indeferidas
    return params
# ... existing code ...
        else:
            # Padrão ADC 58 STF + Lei 14.905/2024
            if comp_dt < ajuizamento_dt:
                fator_corr *= (Decimal("1") + ipca_r)
            elif comp_dt < marco_lei_14905:
                # Fase judicial pré-Lei 14.905/2024: SELIC engloba correção e juros
                fator_corr *= (Decimal("1") + selic_r)
            else:
                # A partir de 30/08/2024 (Lei 14.905/2024): IPCA + Taxa Legal (SELIC - IPCA, piso 0)
                fator_corr *= (Decimal("1") + ipca_r)
                taxa_legal = max(Decimal("0"), selic_r - ipca_r)
                fator_juros += taxa_legal

    return fator_corr.quantize(Decimal("0.000000001"), rounding=ROUND_HALF_UP), fator_juros.quantize(
        Decimal("0.000001"), rounding=ROUND_HALF_UP
    )
# ... existing code ...
    def _build_monthly_occurrences(
        base_val: Decimal,
        div_val: Decimal,
        mult_val: Decimal,
        qtd_fn,
    ) -> List[Dict[str, Any]]:
        ocs: List[Dict[str, Any]] = []
        for y, m, s_dt, e_dt in months:
            qtd = D(qtd_fn(y, m, s_dt, e_dt))
            if qtd <= 0 or base_val <= 0:
                continue
            devido = money_round((base_val / div_val) * mult_val * qtd)
            pago = Decimal("0.00")
            diferenca = money_round(devido - pago)
            f_corr, f_jur = _get_correction_and_interest_factors(e_dt, aju_dt, liq_dt, criterio)
            v_corr = money_round(diferenca * f_corr)
            v_jur = money_round(v_corr * f_jur)
            ocs.append(
                {
                    "competencia": f"{m:02d}/{y:04d}",
                    "periodo_mensal": f"{s_dt.day:02d} a {e_dt.day:02d}/{m:02d}/{y:04d}",
                    "ano": y,
                    "data_vencimento": e_dt.strftime("%d/%m/%Y"),
                    "base": base_val,
                    "divisor": div_val,
                    "multiplicador": mult_val,
                    "quantidade": qtd,
                    "dobra": "Não",
                    "devido": devido,
                    "pago": pago,
                    "diferenca": diferenca,
                    "indice_correcao": f_corr,
                    "valor_corrigido": v_corr,
                    "taxa_juros": f_jur,
                    "juros": v_jur,
                }
            )
        return ocs
# ... existing code ...
            ocs_13.append(
                {
                    "competencia": f"13/{y:04d}",
                    "periodo_mensal": f"20 a 20/12/{y:04d}",
                    "ano": y,
                    "data_vencimento": venc_13.strftime("%d/%m/%Y"),
                    "base": media_anual,
                    "divisor": Decimal("12"),
                    "multiplicador": Decimal("1"),
                    "quantidade": avos,
                    "dobra": "Não",
                    "devido": devido_13,
                    "pago": Decimal("0.00"),
                    "diferenca": devido_13,
                    "indice_correcao": f_corr,
                    "valor_corrigido": v_corr,
                    "taxa_juros": f_jur,
                    "juros": v_jur,
                }
            )
        _add_verba(
            nome=f"13º SALÁRIO SOBRE {origem_label}",
            verba_origem=origem_label,
            incidencias="Contribuição Social / IRPF / FGTS",
            formula=f"(((({origem_label}) / 12,0000) X 1,00000000) X AVOS)",
            ocorrencias=ocs_13,
            salarial=True,
        )

        # 2) Reflexo em Férias + 1/3 Constitucional
        total_src_dev = sum(D(o["devido"]) for o in ocs_src)
        media_geral = money_round(total_src_dev / Decimal(str(len(ocs_src))))
        avos_totais = Decimal(str(min(12, len(ocs_src))))
        devido_fer = money_round((media_geral / Decimal("12")) * avos_totais * Decimal("1.33333333"))
        venc_fer = periodo_fim or liq_dt
        f_corr_f, f_jur_f = _get_correction_and_interest_factors(venc_fer, aju_dt, liq_dt, criterio)
        v_corr_f = money_round(devido_fer * f_corr_f)
        v_jur_f = money_round(v_corr_f * f_jur_f)
        _add_verba(
            nome=f"FÉRIAS + 1/3 SOBRE {origem_label}",
            verba_origem=origem_label,
            incidencias="FGTS",
            formula=f"(((({origem_label}) / 12,0000) X 1,33333333) X AVOS)",
            ocorrencias=[
                {
                    "competencia": venc_fer.strftime("%m/%Y"),
                    "periodo_mensal": f"{venc_fer.day:02d} a {venc_fer.day:02d}/{venc_fer.month:02d}/{venc_fer.year:04d}",
                    "ano": venc_fer.year,
                    "data_vencimento": venc_fer.strftime("%d/%m/%Y"),
                    "base": media_geral,
                    "divisor": Decimal("12"),
                    "multiplicador": Decimal("1.33333333"),
                    "quantidade": avos_totais,
                    "dobra": "Não",
                    "devido": devido_fer,
                    "pago": Decimal("0.00"),
                    "diferenca": devido_fer,
                    "indice_correcao": f_corr_f,
                    "valor_corrigido": v_corr_f,
                    "taxa_juros": f_jur_f,
                    "juros": v_jur_f,
                }
            ],
            salarial=False,
        )

    # A. Horas Extras 50% (somente se deferidas)
    if params.get("tem_he50") and D(params.get("qtd_he50_mensal", 0)) > 0:
        qtd_he50 = D(params.get("qtd_he50_mensal", 0))
        ocs_he50 = _build_monthly_occurrences(salario, divisor, Decimal("1.50"), lambda y, m, s, e: qtd_he50)
        v_he50 = _add_verba(
            nome="HORAS EXTRAS 50%",
            verba_origem="HORAS EXTRAS 50%",
            incidencias="FGTS / Contribuição Social / IRPF",
            formula="(((SALÁRIO BASE / CARGA HORÁRIA) X 1,50000000) X QUANTIDADE)",
            ocorrencias=ocs_he50,
            salarial=True,
        )
        if v_he50 and params.get("reflexos_he50"):
            _build_segregated_reflexes(v_he50, "HORAS EXTRAS 50%")

    # B. Horas Extras 100%
    if params.get("tem_he100") and D(params.get("qtd_he100_mensal", 0)) > 0:
        qtd_he100 = D(params.get("qtd_he100_mensal", 0))
        ocs_he100 = _build_monthly_occurrences(salario, divisor, Decimal("2.00"), lambda y, m, s, e: qtd_he100)
        v_he100 = _add_verba(
            nome="HORAS EXTRAS 100%",
            verba_origem="HORAS EXTRAS 100%",
            incidencias="FGTS / Contribuição Social / IRPF",
            formula="(((SALÁRIO BASE / CARGA HORÁRIA) X 2,00000000) X QUANTIDADE)",
            ocorrencias=ocs_he100,
            salarial=True,
        )
        if v_he100 and params.get("reflexos_he100"):
            _build_segregated_reflexes(v_he100, "HORAS EXTRAS 100%")

    # C. Intervalo Intrajornada (Art. 71, §4º CLT — natureza indenizatória pós-Reforma Trabalhista)
    if params.get("tem_intrajornada"):
        hrs_dia = D(params.get("qtd_intrajornada_mensal", 1)) or Decimal("1")
        ocs_intra = _build_monthly_occurrences(
            salario,
            divisor,
            Decimal("1.50"),
            lambda y, m, s, e: Decimal(str(_count_business_days(s, e, sabado_util))) * hrs_dia,
        )
        v_intra = _add_verba(
            nome="INTERVALO INTRAJORNADA (ART. 71 §4º CLT)",
            verba_origem="INTERVALO INTRAJORNADA",
            incidencias="Não há.",
            formula="(((SALÁRIO BASE / CARGA HORÁRIA) X 1,50000000) X DIAS ÚTEIS)",
            ocorrencias=ocs_intra,
            salarial=False,
            teto_valor=D(params.get("teto_intrajornada", 0)),
            escopo_teto=str(params.get("teto_intrajornada_escopo", "VERBA_PRINCIPAL")),
        )
        if v_intra and params.get("reflexos_intrajornada"):
            _build_segregated_reflexes(v_intra, "INTERVALO INTRAJORNADA")

    # D. Devolução de Descontos Indevidos
    if params.get("tem_descontos_indevidos") and D(params.get("valor_descontos_indevidos", 0)) > 0:
        val_desc = money_round(D(params.get("valor_descontos_indevidos", 0)))
        venc_desc = periodo_fim or aju_dt
        f_corr_d, f_jur_d = _get_correction_and_interest_factors(venc_desc, aju_dt, liq_dt, criterio)
        v_corr_d = money_round(val_desc * f_corr_d)
        v_jur_d = money_round(v_corr_d * f_jur_d)
        _add_verba(
            nome="DEVOLUÇÃO DE DESCONTOS INDEVIDOS",
            verba_origem="DEVOLUÇÃO DE DESCONTOS",
            incidencias="Não há.",
            formula="(VALOR ARBITRADO / COMPROVADO)",
            ocorrencias=[
                {
                    "competencia": venc_desc.strftime("%m/%Y"),
                    "periodo_mensal": f"{venc_desc.day:02d} a {venc_desc.day:02d}/{venc_desc.month:02d}/{venc_desc.year:04d}",
                    "ano": venc_desc.year,
                    "data_vencimento": venc_desc.strftime("%d/%m/%Y"),
                    "base": val_desc,
                    "divisor": Decimal("1"),
                    "multiplicador": Decimal("1"),
                    "quantidade": Decimal("1"),
                    "dobra": "Não",
                    "devido": val_desc,
                    "pago": Decimal("0.00"),
                    "diferenca": val_desc,
                    "indice_correcao": f_corr_d,
                    "valor_corrigido": v_corr_d,
                    "taxa_juros": f_jur_d,
                    "juros": v_jur_d,
                }
            ],
            salarial=False,
            teto_valor=D(params.get("teto_descontos_indevidos", 0)),
            escopo_teto="VERBA_PRINCIPAL",
        )

    # E. Multas dos Arts. 477 e 467 da CLT
    if params.get("multa_477") and salario > 0:
        venc_477 = periodo_fim or aju_dt
        f_c, f_j = _get_correction_and_interest_factors(venc_477, aju_dt, liq_dt, criterio)
        vc = money_round(salario * f_c)
        vj = money_round(vc * f_j)
        _add_verba(
            nome="MULTA DO ART. 477, § 8º, DA CLT",
            verba_origem="MULTA ART. 477 CLT",
            incidencias="Não há.",
            formula="(1 SALÁRIO BASE)",
            ocorrencias=[
                {
                    "competencia": venc_477.strftime("%m/%Y"),
                    "periodo_mensal": f"{venc_477.day:02d} a {venc_477.day:02d}/{venc_477.month:02d}/{venc_477.year:04d}",
                    "ano": venc_477.year,
                    "data_vencimento": venc_477.strftime("%d/%m/%Y"),
                    "base": salario,
                    "divisor": Decimal("1"),
                    "multiplicador": Decimal("1"),
                    "quantidade": Decimal("1"),
                    "dobra": "Não",
                    "devido": salario,
                    "pago": Decimal("0.00"),
                    "diferenca": salario,
                    "indice_correcao": f_c,
                    "valor_corrigido": vc,
                    "taxa_juros": f_j,
                    "juros": vj,
                }
            ],
            salarial=False,
        )

    if params.get("multa_467"):
        base_467 = money_round(sum(D(v["valor"]) for v in verbas) * Decimal("0.50"))
        if base_467 > 0:
            venc_467 = aju_dt
            f_c, f_j = _get_correction_and_interest_factors(venc_467, aju_dt, liq_dt, criterio)
            vc = money_round(base_467 * f_c)
            vj = money_round(vc * f_j)
            _add_verba(
                nome="MULTA DO ART. 467 DA CLT (50%)",
                verba_origem="MULTA ART. 467 CLT",
                incidencias="Não há.",
                formula="(50% SOBRE VERBAS RESCISÓRIAS INCONTROVERSAS)",
                ocorrencias=[
                    {
                        "competencia": venc_467.strftime("%m/%Y"),
                        "periodo_mensal": f"{venc_467.day:02d} a {venc_467.day:02d}/{venc_467.month:02d}/{venc_467.year:04d}",
                        "ano": venc_467.year,
                        "data_vencimento": venc_467.strftime("%d/%m/%Y"),
                        "base": base_467,
                        "divisor": Decimal("1"),
                        "multiplicador": Decimal("1"),
                        "quantidade": Decimal("1"),
                        "dobra": "Não",
                        "devido": base_467,
                        "pago": Decimal("0.00"),
                        "diferenca": base_467,
                        "indice_correcao": f_c,
                        "valor_corrigido": vc,
                        "taxa_juros": f_j,
                        "juros": vj,
                    }
                ],
                salarial=False,
            )

    # F. Verbas Personalizadas da Sentença
    for vp in params.get("verbas_personalizadas", []):
        vp_nome = str(vp.get("nome", "")).strip()
        vp_val = money_round(D(vp.get("valor_principal", 0)))
        vp_sal = bool(vp.get("salarial", False))
        if vp_nome and vp_val > 0:
            venc_vp = periodo_fim or aju_dt
            f_c, f_j = _get_correction_and_interest_factors(venc_vp, aju_dt, liq_dt, criterio)
            vc = money_round(vp_val * f_c)
            vj = money_round(vc * f_j)
            _add_verba(
                nome=vp_nome.upper(),
                verba_origem=vp_nome.upper(),
                incidencias="FGTS / Contribuição Social / IRPF" if vp_sal else "Não há.",
                formula="(VALOR FIXADO NA SENTENÇA)",
                ocorrencias=[
                    {
                        "competencia": venc_vp.strftime("%m/%Y"),
                        "periodo_mensal": f"{venc_vp.day:02d} a {venc_vp.day:02d}/{venc_vp.month:02d}/{venc_vp.year:04d}",
                        "ano": venc_vp.year,
                        "data_vencimento": venc_vp.strftime("%d/%m/%Y"),
                        "base": vp_val,
                        "divisor": Decimal("1"),
                        "multiplicador": Decimal("1"),
                        "quantidade": Decimal("1"),
                        "dobra": "Não",
                        "devido": vp_val,
                        "pago": Decimal("0.00"),
                        "diferenca": vp_val,
                        "indice_correcao": f_c,
                        "valor_corrigido": vc,
                        "taxa_juros": f_j,
                        "juros": vj,
                    }
                ],
                salarial=vp_sal,
            )

    # G. FGTS (8%) + Multa de 40% sobre incidências elegíveis (detalhado mês a mês no padrão PJe-Calc)
    total_fgts_deposito = Decimal("0.00")
    total_fgts_deposito_corr = Decimal("0.00")
    total_fgts_juros = Decimal("0.00")
    total_fgts_multa40 = Decimal("0.00")
    fgts_memoria: List[Dict[str, Any]] = []
    fgts_por_comp: Dict[str, Dict[str, Any]] = {}
    if params.get("reflexos_em_fgts"):
        for v in verbas:
            if "FGTS" in v.get("incidencias", ""):
                for oc in v.get("ocorrencias", []):
                    comp = oc.get("competencia", "")
                    if not comp:
                        continue
                    rec_f = fgts_por_comp.setdefault(
                        comp,
                        {
                            "competencia": comp,
                            "data_vencimento": oc.get("data_vencimento", ""),
                            "base_fgts_historico": Decimal("0.00"),
                            "deposito_historico": Decimal("0.00"),
                            "indice_correcao": D(oc.get("indice_correcao", 1)),
                            "base_fgts_corrigida": Decimal("0.00"),
                            "deposito_corrigido": Decimal("0.00"),
                            "juros_fgts": Decimal("0.00"),
                        },
                    )
                    dev_oc = D(oc["devido"])
                    corr_oc = D(oc["valor_corrigido"])
                    jur_oc = D(oc["juros"])
                    rec_f["base_fgts_historico"] += dev_oc
                    rec_f["deposito_historico"] += money_round(dev_oc * Decimal("0.08"))
                    rec_f["base_fgts_corrigida"] += corr_oc
                    rec_f["deposito_corrigido"] += money_round(corr_oc * Decimal("0.08"))
                    rec_f["juros_fgts"] += money_round(jur_oc * Decimal("0.08"))
        for comp_k in sorted(fgts_por_comp.keys(), key=lambda s: (s.split("/")[1], s.split("/")[0]) if "/" in s else ("0", s)):
            item_f = fgts_por_comp[comp_k]
            item_f["total_fgts_comp"] = money_round(item_f["deposito_corrigido"] + item_f["juros_fgts"])
            total_fgts_deposito_corr += item_f["deposito_corrigido"]
            total_fgts_juros += item_f["juros_fgts"]
            total_fgts_deposito += item_f["total_fgts_comp"]
            fgts_memoria.append(item_f)
        total_fgts_deposito = money_round(total_fgts_deposito)
        total_fgts_deposito_corr = money_round(total_fgts_deposito_corr)
        total_fgts_juros = money_round(total_fgts_juros)
        if "sem justa causa" in str(params.get("modalidade_rescisao", "")).lower():
            total_fgts_multa40 = money_round(total_fgts_deposito * Decimal("0.40"))
    total_fgts_bruto = money_round(total_fgts_deposito + total_fgts_multa40)

    # Histórico salarial mês a mês do período do cálculo
    historico_salarial_memoria: List[Dict[str, Any]] = []
    for y, m, s_dt, e_dt in months:
        dias_no_mes = (e_dt - s_dt).days + 1
        val_hora = money_round(salario / divisor) if divisor > 0 else Decimal("0.00")
        historico_salarial_memoria.append(
            {
                "competencia": f"{m:02d}/{y:04d}",
                "dias_trabalhados": dias_no_mes,
                "salario_base": salario,
                "divisor": divisor,
                "valor_hora": val_hora,
            }
        )

    # H. Contribuição Previdenciária (Segurado progressivo por competência + Patronal)
    aliq_patronal_total = (
        D(params.get("inss_patronal_pct", 20))
        + D(params.get("rat_sat_pct", 2))
        + D(params.get("terceiros_pct", 5.8))
    ) / Decimal("100")

    inss_empregado = Decimal("0.00")
    inss_empregador = Decimal("0.00")
    inss_memoria: List[Dict[str, Any]] = []
    for comp, base_c in sorted(inss_por_comp.items(), key=lambda kv: (kv[0].split("/")[1], kv[0].split("/")[0]) if "/" in kv[0] else ("0", kv[0])):
        ano_c = int(comp.split("/")[1]) if "/" in comp else liq_dt.year
        seg_c = _calc_inss_segurado_competencia(base_c, ano_c)
        pat_c = money_round(base_c * aliq_patronal_total)
        aliq_efetiva = ((seg_c / base_c) * Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if base_c > 0 else Decimal("0.00")
        inss_empregado += seg_c
        inss_empregador += pat_c
        inss_memoria.append(
            {
                "competencia": comp,
                "base_inss": base_c,
                "aliquota_efetiva_segurado": aliq_efetiva,
                "inss_segurado": seg_c,
                "inss_patronal": pat_c,
                "total_comp": money_round(seg_c + pat_c),
            }
        )
    inss_empregado = money_round(inss_empregado)
    inss_empregador = money_round(inss_empregador)

    base_tributavel_total = money_round(sum(inss_por_comp.values()))
    total_contribuicao_social = money_round(inss_empregado + inss_empregador)

    # I. Demonstrativo de Juros sobre Verbas (Regra PJe-Calc item 7: apurados APÓS dedução da contribuição social do segurado)
    inss_seg_map = {x["competencia"]: D(x["inss_segurado"]) for x in inss_memoria}
    juros_por_comp_map: Dict[str, Dict[str, Any]] = {}
    for v in verbas:
        for oc in v.get("ocorrencias", []):
            comp = oc.get("competencia", "")
            if not comp:
                continue
            item_j = juros_por_comp_map.setdefault(
                comp,
                {
                    "competencia": comp,
                    "data_inicial": aju_dt.strftime("%d/%m/%Y"),
                    "total_verbas_corrigido": Decimal("0.00"),
                    "taxa_juros_pct": D(oc.get("taxa_juros", 0)) * Decimal("100"),
                    "juros_bruto": Decimal("0.00"),
                },
            )
            item_j["total_verbas_corrigido"] += D(oc["valor_corrigido"])
            item_j["juros_bruto"] += D(oc["juros"])

    juros_memoria: List[Dict[str, Any]] = []
    for comp_k in sorted(juros_por_comp_map.keys(), key=lambda s: (s.split("/")[1], s.split("/")[0]) if "/" in s else ("0", s)):
        ij = juros_por_comp_map[comp_k]
        inss_ded = inss_seg_map.get(comp_k, Decimal("0.00"))
        tot_vc = money_round(ij["total_verbas_corrigido"])
        capital = max(Decimal("0.00"), money_round(tot_vc - inss_ded))
        taxa_dec = D(ij["taxa_juros_pct"]) / Decimal("100")
        juros_liq = money_round(capital * taxa_dec) if taxa_dec > 0 else Decimal("0.00")
        juros_memoria.append(
            {
                "competencia": comp_k,
                "data_inicial": ij["data_inicial"],
                "total_verbas_corrigido": tot_vc,
                "contribuicao_social_segurado": inss_ded,
                "previdencia_privada": Decimal("0.00"),
                "capital": capital,
                "taxa_juros_pct": ij["taxa_juros_pct"],
                "juros": juros_liq,
            }
        )

    # J. Totalizações das Verbas
    verbas_principal = money_round(sum(D(v["valor"]) for v in verbas))
    verbas_corrigido = money_round(sum(D(v["valor_corrigido"]) for v in verbas))
    verbas_juros = money_round(sum(D(v["juros"]) for v in verbas))
    total_verbas_com_juros = money_round(verbas_corrigido + verbas_juros)

    # K. IRRF RRA (Art. 12-A Lei 7.713/88)
    meses_rra = max(1, len(inss_por_comp)) if base_tributavel_total > 0 else 1
    base_irrf_liquida = max(Decimal("0.00"), base_tributavel_total - inss_empregado)
    irpf_reclamante = _calc_irrf_rra(base_irrf_liquida, meses_rra)
    irrf_memoria = {
        "base_bruta_tributavel": base_tributavel_total,
        "deducao_inss_segurado": inss_empregado,
        "base_liquida_rra": base_irrf_liquida,
        "meses_acumulaveis_nm": meses_rra,
        "rendimento_medio_mensal": money_round(base_irrf_liquida / Decimal(str(meses_rra))),
        "irpf_devido": irpf_reclamante,
    }

    # L. Bruto, Descontos, Líquido, Honorários, Custas e Total Devido pelo Reclamado
    bruto_devido_reclamante = money_round(total_verbas_com_juros + total_fgts_bruto)
    abatimentos = money_round(D(params.get("valor_pago_abatimento", 0)))
    total_descontos = money_round(inss_empregado + irpf_reclamante + abatimentos)
    liquido_devido_reclamante = money_round(max(Decimal("0.00"), bruto_devido_reclamante - total_descontos))

    hon_pct = D(params.get("honorarios_sucumbencia_pct", 0))
    honorarios_valor = money_round(bruto_devido_reclamante * (hon_pct / Decimal("100")))

    hon_rec_pct = D(params.get("honorarios_reciprocos_pct", 0))
    base_improc = D(params.get("base_pedidos_improcedentes", 0))
    honorarios_reciprocos_valor = money_round(base_improc * (hon_rec_pct / Decimal("100")))

    custas_arbitradas = money_round(D(params.get("custas_processuais", 0)))
    custas_devidas_execucao = Decimal("0.00") if params.get("custas_ja_recolhidas") else custas_arbitradas

    total_devido_reclamado = money_round(
        liquido_devido_reclamante
        + total_contribuicao_social
        + irpf_reclamante
        + honorarios_valor
        + custas_devidas_execucao
    )

    pct_remuneratorio = (
        ((base_tributavel_total / verbas_corrigido) * Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if verbas_corrigido > 0
        else Decimal("0.00")
    )

    return {
        "meta": {
            "processo": params.get("processo") or "INFORMAÇÃO NÃO LOCALIZADA",
            "numero_calculo": params.get("numero_calculo", "1"),
            "estado": params.get("estado", "RS"),
            "municipio": params.get("municipio", "GRAVATAI"),
            "cidade_assinatura": params.get("cidade_assinatura", "Gravataí"),
            "vara": params.get("vara", "VARA DO TRABALHO"),
            "reclamante": params.get("reclamante") or "INFORMAÇÃO NÃO LOCALIZADA",
            "reclamado": params.get("reclamado") or "INFORMAÇÃO NÃO LOCALIZADA",
            "admissao": fmt_date_br(adm_dt),
            "demissao": fmt_date_br(dem_dt),
            "ajuizamento": fmt_date_br(aju_dt),
            "liquidacao": fmt_date_br(liq_dt),
            "periodo_calculo": f"{fmt_date_br(periodo_inicio)} a {fmt_date_br(periodo_fim)}",
            "salario_base": salario,
            "maior_remuneracao": D(params.get("maior_remuneracao", salario)),
            "ultima_remuneracao": D(params.get("ultima_remuneracao", salario)),
            "regime_trabalho": params.get("regime_trabalho", "Tempo Integral"),
            "aplicar_prescricao_quinquenal": "Sim" if params.get("aplicar_prescricao_quinquenal") else "Não",
            "projetar_aviso_previo": "Sim" if params.get("projetar_aviso_previo") else "Não",
            "sabado_dia_util": "Sim" if params.get("sabado_dia_util", True) else "Não",
            "divisor": divisor,
            "criterio_correcao": criterio,
        },
        "verbas": verbas,
        "historico_salarial_memoria": historico_salarial_memoria,
        "fgts_memoria": fgts_memoria,
        "inss_memoria": inss_memoria,
        "irrf_memoria": irrf_memoria,
        "juros_memoria": juros_memoria,
        "limites_aplicados": limites_aplicados,
        "resumo": {
            "verbas_principal": verbas_principal,
            "verbas_corrigido": verbas_corrigido,
            "verbas_juros": verbas_juros,
            "total_verbas_com_juros": total_verbas_com_juros,
            "total_fgts_deposito": total_fgts_deposito,
            "total_fgts_deposito_corr": total_fgts_deposito_corr,
            "total_fgts_juros": total_fgts_juros,
            "total_fgts_multa40": total_fgts_multa40,
            "total_fgts_bruto": total_fgts_bruto,
            "bruto_devido_reclamante": bruto_devido_reclamante,
            "inss_empregado": inss_empregado,
            "inss_empregador": inss_empregador,
            "total_contribuicao_social": total_contribuicao_social,
            "irpf_reclamante": irpf_reclamante,
            "abatimentos": abatimentos,
            "total_descontos": total_descontos,
            "liquido_devido_reclamante": liquido_devido_reclamante,
            "honorarios_pct": hon_pct,
            "honorarios_valor": honorarios_valor,
            "honorarios_reciprocos_valor": honorarios_reciprocos_valor,
            "custas_devidas_execucao": custas_devidas_execucao,
            "total_devido_reclamado": total_devido_reclamado,
            "pct_remuneratorio": pct_remuneratorio,
        },
    }
# ... existing code ...
    ing = params.get("_ingestion", {})
    if ing.get("unreadable_pages"):
        doc_issues.append(f"Existem {len(ing['unreadable_pages'])} página(s) sem camada textual (OCR pendente).")
    if not parse_br_date_optional(params.get("admissao")):
        doc_issues.append("Data de admissão não localizada ou pendente de confirmação.")
    if not parse_br_date_optional(params.get("demissao")):
        doc_issues.append("Data de demissão não localizada ou pendente de confirmação.")
    if D(params.get("salario_base", 0)) <= 0:
        doc_issues.append("Salário base contratual não localizado ou zerado.")

    if params.get("_representation_conflict"):
# ... existing code ...
def generate_judicial_decision_summary(params: Dict[str, Any]) -> str:
    if params.get("_ai_decision_summary"):
        return str(params["_ai_decision_summary"])
    ativas: List[str] = []
    if params.get("tem_he50"):
        ativas.append(f"Horas Extras 50% ({params.get('qtd_he50_mensal', 0)}h/mês)")
    if params.get("tem_he100"):
        ativas.append(f"Horas Extras 100% ({params.get('qtd_he100_mensal', 0)}h/mês)")
    if params.get("tem_intrajornada"):
        teto_i = D(params.get("teto_intrajornada", 0))
        ativas.append(
            f"Intervalo Intrajornada{f' (limitado ao teto judicial de R$ {fmt_brl(teto_i)})' if teto_i > 0 else ''}"
        )
    if params.get("tem_descontos_indevidos"):
        ativas.append(f"Devolução de Descontos Indevidos (R$ {fmt_brl(params.get('valor_descontos_indevidos', 0))})")
    if params.get("reflexos_em_fgts"):
        ativas.append("FGTS 8% (+40% se dispensa imotivada)")
    if params.get("multa_477"):
        ativas.append("Multa do Art. 477 da CLT")
    if params.get("multa_467"):
        ativas.append("Multa do Art. 467 da CLT")
    for vp in params.get("verbas_personalizadas", []):
        if vp.get("nome") and D(vp.get("valor_principal", 0)) > 0:
            ativas.append(f"{vp['nome']} (R$ {fmt_brl(vp['valor_principal'])})")

    excluidas = [v["verba"] for v in params.get("_verbas_indeferidas", [])]
    txt = (
        f"Processo nº {params.get('processo') or 'Não identificado'} | "
        f"Vínculo selecionado automaticamente: {fmt_date_br(params.get('admissao'))} a {fmt_date_br(params.get('demissao'))} | "
        f"Verbas deferidas consideradas: {', '.join(ativas) if ativas else 'Nenhuma verba ativa'}."
    )
    if excluidas:
        txt += f" Verbas excluídas/indeferidas: {', '.join(excluidas)}."
    return txt


# ==============================================================================
# 5. GERADOR DE RELATÓRIO PDF (LAYOUT IDÊNTICO AO ANEXO PJE-CALC CIDADÃO)
# ==============================================================================
def _pdf_safe(txt: Any) -> str:
    s = str(txt or "")
    return s.encode("latin-1", errors="replace").decode("latin-1")


class PJeCalcExactPDF(FPDF if FPDF is not None else object):
    def __init__(self, meta: Dict[str, Any]):
        if FPDF is not None:
            super().__init__(orientation="P", unit="mm", format="A4")
        self.meta = meta
        self.generated_ts = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")

    def header(self) -> None:
        self.set_font("Helvetica", "B", 11)
        self.cell(190, 5, _pdf_safe("JurisCalc Pro — Sistema de Cálculos Trabalhistas"), ln=1, align="C")
        self.set_font("Helvetica", "", 8)
        self.cell(190, 4, _pdf_safe("Memória de Cálculo de Liquidação Trabalhista"), ln=1, align="C")
        self.ln(2)
        self.set_font("Helvetica", "B", 8)
        self.cell(135, 4.5, _pdf_safe(f"Reclamante: {self.meta['reclamante']}"))
        self.cell(55, 4.5, _pdf_safe(f"Processo: {self.meta['processo']}"), ln=1, align="R")
        self.cell(135, 4.5, _pdf_safe(f"Reclamado: {self.meta['reclamado']}"))
        self.cell(55, 4.5, _pdf_safe(f"Cálculo: {self.meta['numero_calculo']}"), ln=1, align="R")
        self.cell(90, 4.5, _pdf_safe(f"Período do Cálculo: {self.meta['periodo_calculo']}"))
        self.cell(50, 4.5, _pdf_safe(f"Data Ajuizamento: {self.meta['ajuizamento']}"))
        self.cell(50, 4.5, _pdf_safe(f"Data Liquidação: {self.meta['liquidacao']}"), ln=1, align="R")
        self.set_font("Helvetica", "B", 9.5)
        self.cell(190, 5.5, _pdf_safe("PLANILHA DE CÁLCULO"), ln=1, align="C")
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(2)

    def footer(self) -> None:
        self.set_y(-13)
        self.set_font("Helvetica", "I", 7)
        self.cell(
            145,
            4,
            _pdf_safe(f"Cálculo liquidado por JurisCalc Pro v2.14.0 em {self.generated_ts}."),
            align="L",
        )
        self.cell(45, 4, _pdf_safe(f"Pág. {self.page_no()} de {{nb}}"), align="R")


def build_pjecalc_pdf(calc_result: Dict[str, Any], audit_report: Dict[str, Any]) -> bytes:
    meta = calc_result["meta"]
    res = calc_result["resumo"]
    verbas = calc_result["verbas"]

    if FPDF is None:
        return f"PLANILHA DE CALCULO - {meta['processo']} - TOTAL: {fmt_brl(res['total_devido_reclamado'])}".encode("utf-8")

    pdf = PJeCalcExactPDF(meta)
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=16)

    # --------------------------------------------------------------------------
    # PÁGINA 1: RESUMO DO CÁLCULO (EXATAMENTE COMO NO ANEXO PÁGINA 2)
    # --------------------------------------------------------------------------
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(190, 6, _pdf_safe("Resumo do Cálculo"), ln=1, align="C")
    pdf.ln(1)

    # Quadro 1: Descrição do Bruto Devido ao Reclamante
    pdf.set_fill_color(230, 234, 240)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(112, 5.5, _pdf_safe("Descrição do Bruto Devido ao Reclamante"), border=1, fill=True)
    pdf.cell(26, 5.5, _pdf_safe("Valor Corrigido"), border=1, align="R", fill=True)
    pdf.cell(26, 5.5, _pdf_safe("Juros"), border=1, align="R", fill=True)
    pdf.cell(26, 5.5, _pdf_safe("Total"), border=1, align="R", ln=1, fill=True)

    pdf.set_font("Helvetica", "", 7.8)
    for v in verbas:
        pdf.cell(112, 5, _pdf_safe(v["nome"]), border=1)
        pdf.cell(26, 5, fmt_brl(v["valor_corrigido"]), border=1, align="R")
        pdf.cell(26, 5, fmt_brl(v["juros"]), border=1, align="R")
        pdf.cell(26, 5, fmt_brl(v["total_verba"]), border=1, align="R", ln=1)

    if res["total_fgts_deposito"] > 0:
        pdf.cell(112, 5, _pdf_safe("FGTS 8%"), border=1)
        pdf.cell(26, 5, fmt_brl(res["total_fgts_deposito_corr"]), border=1, align="R")
        pdf.cell(26, 5, fmt_brl(res["total_fgts_juros"]), border=1, align="R")
        pdf.cell(26, 5, fmt_brl(res["total_fgts_deposito"]), border=1, align="R", ln=1)

    if res["total_fgts_multa40"] > 0:
        pdf.cell(112, 5, _pdf_safe("MULTA DO FGTS 40%"), border=1)
        pdf.cell(26, 5, fmt_brl(res["total_fgts_multa40"]), border=1, align="R")
        pdf.cell(26, 5, "0,00", border=1, align="R")
        pdf.cell(26, 5, fmt_brl(res["total_fgts_multa40"]), border=1, align="R", ln=1)

    total_quadro1_corr = money_round(res["verbas_corrigido"] + res["total_fgts_deposito_corr"] + res["total_fgts_multa40"])
    total_quadro1_juros = money_round(res["verbas_juros"] + res["total_fgts_juros"])
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(112, 5.5, _pdf_safe("Total"), border=1, fill=True)
    pdf.cell(26, 5.5, fmt_brl(total_quadro1_corr), border=1, align="R", fill=True)
    pdf.cell(26, 5.5, fmt_brl(total_quadro1_juros), border=1, align="R", fill=True)
    pdf.cell(26, 5.5, fmt_brl(res["bruto_devido_reclamante"]), border=1, align="R", ln=1, fill=True)

    pdf.ln(3)

    # Quadro 2: Descrição de Créditos e Descontos do Reclamante
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(155, 5.5, _pdf_safe("Descrição de Créditos e Descontos do Reclamante"), border=1, fill=True)
    pdf.cell(35, 5.5, _pdf_safe("Valor"), border=1, align="R", ln=1, fill=True)

    rows_q2 = [
        ("VERBAS", fmt_brl(res["total_verbas_com_juros"]), False),
        ("FGTS", fmt_brl(res["total_fgts_bruto"]), False),
        ("Bruto Devido ao Reclamante", fmt_brl(res["bruto_devido_reclamante"]), True),
        ("DEDUÇÃO DE CONTRIBUIÇÃO SOCIAL", f"({fmt_brl(res['inss_empregado'])})", False),
        ("IRPF DEVIDO PELO RECLAMANTE", f"({fmt_brl(res['irpf_reclamante'])})" if res["irpf_reclamante"] > 0 else "0,00", False),
    ]
    if res["abatimentos"] > 0:
        rows_q2.append(("VALORES PAGOS / ABATIMENTOS", f"({fmt_brl(res['abatimentos'])})", False))
    rows_q2.append(("Total de Descontos", f"({fmt_brl(res['total_descontos'])})", True))
    rows_q2.append(("Líquido Devido ao Reclamante", fmt_brl(res["liquido_devido_reclamante"]), True))

    for lbl_q2, val_q2, is_bold in rows_q2:
        pdf.set_font("Helvetica", "B" if is_bold else "", 7.8)
        pdf.cell(155, 5, _pdf_safe(lbl_q2), border=1, fill=is_bold)
        pdf.cell(35, 5, _pdf_safe(val_q2), border=1, align="R", ln=1, fill=is_bold)

    pdf.ln(1.5)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(190, 5, _pdf_safe(f"Percentual de Parcelas Remuneratórias e Tributáveis: {fmt_num(res['pct_remuneratorio'], 2)}%"), ln=1, align="R")
    pdf.ln(1.5)

    # Quadro 3: Descrição de Débitos do Reclamado por Credor
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(155, 5.5, _pdf_safe("Descrição de Débitos do Reclamado por Credor"), border=1, fill=True)
    pdf.cell(35, 5.5, _pdf_safe("Valor"), border=1, align="R", ln=1, fill=True)

    rows_q3 = [
        ("LÍQUIDO DEVIDO AO RECLAMANTE", fmt_brl(res["liquido_devido_reclamante"]), False),
        ("CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS", fmt_brl(res["total_contribuicao_social"]), False),
        ("IRPF DEVIDO PELO RECLAMANTE", fmt_brl(res["irpf_reclamante"]), False),
    ]
    if res["honorarios_valor"] > 0:
        rows_q3.append((f"HONORÁRIOS ADVOCATÍCIOS ({fmt_num(res['honorarios_pct'], 2)}%)", fmt_brl(res["honorarios_valor"]), False))
    if res["custas_devidas_execucao"] > 0:
        rows_q3.append(("CUSTAS PROCESSUAIS", fmt_brl(res["custas_devidas_execucao"]), False))
    rows_q3.append(("Total Devido pelo Reclamado", fmt_brl(res["total_devido_reclamado"]), True))

    for lbl_q3, val_q3, is_bold in rows_q3:
        pdf.set_font("Helvetica", "B" if is_bold else "", 7.8)
        pdf.cell(155, 5, _pdf_safe(lbl_q3), border=1, fill=is_bold)
        pdf.cell(35, 5, _pdf_safe(val_q3), border=1, align="R", ln=1, fill=is_bold)

    # --------------------------------------------------------------------------
    # PÁGINA 2: CRITÉRIO DE CÁLCULO E FUNDAMENTAÇÃO LEGAL (ANEXO PÁGINA 3)
    # --------------------------------------------------------------------------
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(190, 6, _pdf_safe("Critério de Cálculo e Fundamentação Legal"), ln=1, align="C")
    pdf.ln(2)
    criterios_oficiais = [
        "1. Prazo do aviso prévio apurado segundo a Lei nº 12.506/2011.",
        "2. Avos de férias e/ou 13º salário apurados considerando a projeção do prazo do aviso prévio, quando deferido.",
        f"3. Valores atualizados segundo o critério '{meta['criterio_correcao']}', acumulados a partir do mês subsequente ao vencimento, conforme Súmula nº 381 do TST.",
        "4. Contribuições sociais sobre salários devidos apuradas mês a mês pelas alíquotas progressivas da EC 103/2019 e encargos patronais incidentes.",
        "5. Imposto de renda apurado através da 'tabela progressiva acumulada' vigente no mês da liquidação (Art. 12-A da Lei nº 7.713/1988).",
        "6. Juros de mora sobre verbas apurados após a dedução da contribuição social devida pelo reclamante (item 7 do padrão PJe-Calc).",
    ]
    pdf.set_font("Helvetica", "", 8.2)
    for crit_line in criterios_oficiais:
        pdf.multi_cell(190, 5, _pdf_safe(crit_line))
        pdf.ln(1)

    # --------------------------------------------------------------------------
    # PÁGINA 3: DADOS DO CÁLCULO E HISTÓRICO SALARIAL (ANEXO PÁGINAS 4 A 7)
    # --------------------------------------------------------------------------
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(190, 6, _pdf_safe("Dados do Cálculo"), ln=1, align="C")
    pdf.set_font("Helvetica", "", 7.8)
    dados_grade = [
        (f"Estado: {meta['estado']}    Município: {meta['municipio']}", f"Admissão: {meta['admissao']}    Demissão: {meta['demissao']}"),
        (f"Regime de Trabalho: {meta['regime_trabalho']}", f"Aplicar Prescrição Quinquenal: {meta['aplicar_prescricao_quinquenal']}"),
        (f"Maior Remuneração: {fmt_brl(meta['maior_remuneracao'])}", f"Última Remuneração: {fmt_brl(meta['ultima_remuneracao'])}"),
        (f"Carga Horária (Padrão): {fmt_num(meta['divisor'], 2)}", f"Projetar Aviso Prévio Indenizado: {meta['projetar_aviso_previo']}"),
        (f"Sábado como Dia Útil: {meta['sabado_dia_util']}", f"Data da Liquidação: {meta['liquidacao']}"),
    ]
    for c_esq, c_dir in dados_grade:
        pdf.cell(95, 5.2, _pdf_safe(c_esq), border=1)
        pdf.cell(95, 5.2, _pdf_safe(c_dir), border=1, ln=1)

    hist_sal = calc_result.get("historico_salarial_memoria", [])
    if hist_sal:
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(190, 6, _pdf_safe("Histórico Salarial — Ocorrências do Histórico Salarial"), ln=1, align="C")
        pdf.set_fill_color(230, 234, 240)
        pdf.set_font("Helvetica", "B", 7.5)
        h_cols = ["MÊS/ANO", "DIAS NO PERÍODO", "SALÁRIO BASE", "CARGA HORÁRIA", "SALÁRIO-HORA"]
        h_w = [34, 34, 44, 34, 44]
        for ht, hw in zip(h_cols, h_w):
            pdf.cell(hw, 5, _pdf_safe(ht), border=1, align="C" if ht == "MÊS/ANO" else "R", fill=True)
        pdf.ln()
        pdf.set_font("Helvetica", "", 7.5)
        for hs in hist_sal:
            vals_hs = [
                hs["competencia"],
                fmt_num(hs["dias_trabalhados"], 2),
                fmt_brl(hs["salario_base"]),
                fmt_num(hs["divisor"], 2),
                fmt_brl(hs["valor_hora"]),
            ]
            for vh, hw in zip(vals_hs, h_w):
                pdf.cell(hw, 4.5, _pdf_safe(vh), border=1, align="C" if vh == vals_hs[0] else "R")
            pdf.ln()

    # --------------------------------------------------------------------------
    # DEMONSTRATIVO DE VERBAS (EXATAMENTE AS 11 COLUNAS DO ANEXO PÁGS. 7 A 19)
    # --------------------------------------------------------------------------
    if verbas:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(190, 6, _pdf_safe("Demonstrativo de Verbas"), ln=1, align="C")

        for v in verbas:
            pdf.ln(2.5)
            pdf.set_fill_color(230, 234, 240)
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(130, 5, _pdf_safe(f"Nome: {v['nome']}"), border=1, fill=True)
            pdf.cell(60, 5, _pdf_safe(f"Período: {meta['periodo_calculo']}"), border=1, align="R", ln=1, fill=True)
            pdf.set_font("Helvetica", "", 7.5)
            pdf.cell(190, 4.8, _pdf_safe(f"Incidência(s): {v['incidencias']}"), border=1, ln=1)
            pdf.cell(190, 4.8, _pdf_safe(v["formula_pje"]), border=1, ln=1)

            headers_v = ["Período Mensal", "Base", "Divisor", "Multiplicador", "Quantidade", "Dobra", "Devido", "Pago", "Diferença", "Índice Correção", "Valor Corrigido"]
            widths_v = [21, 17, 13, 16, 15, 10, 17, 15, 17, 25, 24]
            pdf.set_font("Helvetica", "B", 6.4)
            for h_txt, w_col in zip(headers_v, widths_v):
                pdf.cell(w_col, 5, _pdf_safe(h_txt), border=1, align="C" if h_txt in ("Período Mensal", "Dobra") else "R", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 6.4)
            for oc in v.get("ocorrencias", []):
                vals_row = [
                    oc.get("periodo_mensal", oc.get("competencia", "")),
                    fmt_brl(oc["base"]),
                    fmt_num(oc["divisor"], 4),
                    fmt_num(oc["multiplicador"], 8),
                    fmt_num(oc["quantidade"], 4),
                    str(oc.get("dobra", "Não")),
                    fmt_brl(oc["devido"]),
                    fmt_brl(oc.get("pago", 0)),
                    fmt_brl(oc.get("diferenca", oc["devido"])),
                    fmt_num(oc["indice_correcao"], 9),
                    fmt_brl(oc["valor_corrigido"]),
                ]
                for idx_c, (val_c, w_col) in enumerate(zip(vals_row, widths_v)):
                    pdf.cell(w_col, 4.4, _pdf_safe(val_c), border=1, align="C" if idx_c in (0, 5) else "R")
                pdf.ln()

            pdf.set_font("Helvetica", "B", 7.2)
            pdf.cell(166, 5, _pdf_safe("Total"), border=1, align="R", fill=True)
            pdf.cell(24, 5, fmt_brl(v["valor_corrigido"]), border=1, align="R", ln=1, fill=True)

    # --------------------------------------------------------------------------
    # DEMONSTRATIVO DE JUROS SOBRE VERBAS (ANEXO PÁGS. 19 A 21)
    # --------------------------------------------------------------------------
    juros_mem = calc_result.get("juros_memoria", [])
    if juros_mem:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(190, 6, _pdf_safe("Demonstrativo de Juros sobre Verbas"), ln=1, align="C")
        pdf.set_fill_color(230, 234, 240)
        pdf.set_font("Helvetica", "B", 7)
        j_headers = ["Ocorrência", "Data Inicial", "Total de Verbas", "Contribuição Social", "Previdência Privada", "Capital", "Taxa", "Juros"]
        j_widths = [20, 22, 26, 26, 24, 26, 18, 28]
        for jh, jw in zip(j_headers, j_widths):
            pdf.cell(jw, 5, _pdf_safe(jh), border=1, align="C" if jh in ("Ocorrência", "Data Inicial") else "R", fill=True)
        pdf.ln()
        pdf.set_font("Helvetica", "", 7)
        for jm in juros_mem:
            j_vals = [
                jm["competencia"],
                jm["data_inicial"],
                fmt_brl(jm["total_verbas_corrigido"]),
                fmt_brl(jm["contribuicao_social_segurado"]),
                fmt_brl(jm.get("previdencia_privada", 0)),
                fmt_brl(jm["capital"]),
                f"{fmt_num(jm['taxa_juros_pct'], 2)}%",
                fmt_brl(jm["juros"]),
            ]
            for idx_j, (jv, jw) in enumerate(zip(j_vals, j_widths)):
                pdf.cell(jw, 4.4, _pdf_safe(jv), border=1, align="C" if idx_j in (0, 1) else "R")
            pdf.ln()
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(162, 5, _pdf_safe("Total"), border=1, align="R", fill=True)
        pdf.cell(28, 5, fmt_brl(sum(D(x["juros"]) for x in juros_mem)), border=1, align="R", ln=1, fill=True)

    # --------------------------------------------------------------------------
    # DEMONSTRATIVO DE FGTS (ANEXO PÁGS. 21 A 24)
    # --------------------------------------------------------------------------
    fgts_mem = calc_result.get("fgts_memoria", [])
    if fgts_mem:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(190, 6, _pdf_safe("Demonstrativo de FGTS"), ln=1, align="C")
        pdf.set_fill_color(230, 234, 240)
        pdf.set_font("Helvetica", "B", 7)
        f_headers = ["Ocorrência", "Base", "Alíquota", "Devido", "Recolhido", "Diferença", "Índice Correção", "Valor Corrigido", "Juros", "Total"]
        f_widths = [18, 20, 14, 18, 16, 18, 24, 22, 18, 22]
        for fh, fw in zip(f_headers, f_widths):
            pdf.cell(fw, 5, _pdf_safe(fh), border=1, align="C" if fh in ("Ocorrência", "Alíquota") else "R", fill=True)
        pdf.ln()
        pdf.set_font("Helvetica", "", 6.8)
        for fm in fgts_mem:
            f_vals = [
                fm["competencia"],
                fmt_brl(fm["base_fgts_historico"]),
                "8%",
                fmt_brl(fm["deposito_historico"]),
                "0,00",
                fmt_brl(fm["deposito_historico"]),
                fmt_num(fm.get("indice_correcao", 1), 9),
                fmt_brl(fm["deposito_corrigido"]),
                fmt_brl(fm["juros_fgts"]),
                fmt_brl(fm["total_fgts_comp"]),
            ]
            for idx_f, (fv, fw) in enumerate(zip(f_vals, f_widths)):
                pdf.cell(fw, 4.4, _pdf_safe(fv), border=1, align="C" if idx_f in (0, 2) else "R")
            pdf.ln()
        pdf.set_font("Helvetica", "B", 7.2)
        pdf.cell(128, 5, _pdf_safe("Total"), border=1, align="R", fill=True)
        pdf.cell(22, 5, fmt_brl(res["total_fgts_deposito_corr"]), border=1, align="R", fill=True)
        pdf.cell(18, 5, fmt_brl(res["total_fgts_juros"]), border=1, align="R", fill=True)
        pdf.cell(22, 5, fmt_brl(res["total_fgts_deposito"]), border=1, align="R", ln=1, fill=True)

    # --------------------------------------------------------------------------
    # DEMONSTRATIVO DE CONTRIBUIÇÃO SOCIAL (ANEXO PÁGS. 24 A 27)
    # --------------------------------------------------------------------------
    inss_mem = calc_result.get("inss_memoria", [])
    if inss_mem:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(190, 6, _pdf_safe("Demonstrativo de Contribuição Social"), ln=1, align="C")
        pdf.set_fill_color(230, 234, 240)
        pdf.set_font("Helvetica", "B", 7.2)
        i_headers = ["Ocorrência", "Salário Devido", "Alíquota Efetiva", "Devido Segurado", "Encargos Patronais", "Total"]
        i_widths = [26, 34, 28, 34, 34, 34]
        for ih, iw in zip(i_headers, i_widths):
            pdf.cell(iw, 5, _pdf_safe(ih), border=1, align="C" if ih == "Ocorrência" else "R", fill=True)
        pdf.ln()
        pdf.set_font("Helvetica", "", 7.2)
        for im in inss_mem:
            i_vals = [
                im["competencia"],
                fmt_brl(im["base_inss"]),
                f"{fmt_num(im.get('aliquota_efetiva_segurado', 0), 2)}%",
                fmt_brl(im["inss_segurado"]),
                fmt_brl(im.get("inss_patronal", 0)),
                fmt_brl(im.get("total_comp", im["inss_segurado"])),
            ]
            for idx_i, (iv, iw) in enumerate(zip(i_vals, i_widths)):
                pdf.cell(iw, 4.5, _pdf_safe(iv), border=1, align="C" if idx_i == 0 else "R")
            pdf.ln()
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(88, 5, _pdf_safe("Total"), border=1, align="R", fill=True)
        pdf.cell(34, 5, fmt_brl(res["inss_empregado"]), border=1, align="R", fill=True)
        pdf.cell(34, 5, fmt_brl(res["inss_empregador"]), border=1, align="R", fill=True)
        pdf.cell(34, 5, fmt_brl(res["total_contribuicao_social"]), border=1, align="R", ln=1, fill=True)

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)


def build_pjc_xml(calc_result: Dict[str, Any]) -> bytes:
    meta = calc_result["meta"]
    res = calc_result["resumo"]
    root = ET.Element("Calculo", attrib={"sistema": "JurisCalcPro-PJeCalc", "versao": "2.14.0"})
# ... existing code ...
def build_minuta_peticao(
    params: Dict[str, Any],
    calc_result: Dict[str, Any],
    audit_report: Dict[str, Any],
    lawyer_notes: str = "",
) -> str:
    meta = calc_result["meta"]
    res = calc_result["resumo"]
    parte_rep = params.get("parte_representada", "Reclamante")
    nome_parte_rep = meta["reclamante"] if parte_rep == "Reclamante" else meta["reclamado"]
    adv_nome = (params.get("advogado_assinante") or "").strip() or "[NOME DO(A) ADVOGADO(A)]"
    adv_oab = (params.get("oab_advogado") or "").strip() or "[OAB/UF 00.000]"
    cidade_ass = (params.get("cidade_assinatura") or "").strip() or meta["municipio"].title()
# ... existing code ...
Termos em que,
Pede deferimento.

{cidade_ass}/{meta['estado']}, {meta['liquidacao']}.

____________________________________________________
{adv_nome}
{adv_oab}
"""
    return peticao
# ... existing code ...
    prev_p = st.session_state.get("params", {})
    if prev_p.get("advogado_assinante"):
        parsed["advogado_assinante"] = prev_p["advogado_assinante"]
    if prev_p.get("oab_advogado"):
        parsed["oab_advogado"] = prev_p["oab_advogado"]
    if prev_p.get("parte_representada"):
        parsed["parte_representada"] = prev_p["parte_representada"]
    if prev_p.get("cidade_assinatura"):
        parsed["cidade_assinatura"] = prev_p["cidade_assinatura"]
# ... existing code ...
        if ai_data.get("honorarios_sucumbencia_pct") and parsed.get("honorarios_sucumbencia_pct", 0) == 0:
            hon_ai = float(D(ai_data["honorarios_sucumbencia_pct"]))
            parsed["honorarios_juiz_pct"] = hon_ai
            parsed["honorarios_sucumbencia_pct"] = hon_ai
        if ai_data.get("custas_processuais") and parsed.get("custas_processuais", 0) == 0:
            parsed["custas_processuais"] = float(D(ai_data["custas_processuais"]))
        if ai_data.get("resumo_decisao_judicial"):
            parsed["_ai_decision_summary"] = str(ai_data["resumo_decisao_judicial"])

    st.session_state["current_calc_uid"] = uuid.uuid4().hex[:12]
    st.session_state["has_analyzed"] = True
    _recompute_from_current_params(parsed, save_to_history=True)


st.set_page_config(
    page_title="JurisCalc Pro • Liquidação Trabalhista Inteligente",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .jc-banner {
        padding: 22px 28px;
        border-radius: 16px;
        background: linear-gradient(125deg, #0B132B 0%, #1C2541 60%, #1E3A5F 100%);
        color: #F8FAFC;
        border-left: 6px solid #D97706;
        margin-bottom: 20px;
        box-shadow: 0 12px 28px -8px rgba(11, 19, 43, 0.35);
    }
    .jc-tag {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 6px;
        background: rgba(217, 119, 6, 0.18);
        border: 1px solid rgba(251, 191, 36, 0.45);
        color: #FBBF24;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin-bottom: 8px;
    }
    .jc-step {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 6px;
        background: #B45309;
        color: #FFFFFF;
        font-weight: 700;
        font-size: 0.75rem;
        margin-right: 8px;
    }
    </style>
    <div class="jc-banner">
        <div class="jc-tag">JurisCalc Pro • Motor de Liquidação Trabalhista Padrão PJe-Calc</div>
        <h2 style="margin: 0 0 6px 0; font-size: 1.62rem; font-weight: 800;">⚖️ Liquidação Trabalhista Automatizada (.PJC + Planilha PJe-Calc + Minuta)</h2>
        <p style="margin: 0; font-size: 0.92rem; color: #E2E8F0;">
            Leitura integral dos autos com identificação autônoma do vínculo • Cálculo fiel às regras do PJe-Calc Cidadão • Planilha PDF idêntica ao padrão oficial
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### ⚖️ JurisCalc Pro")
    st.caption("Liquidação Trabalhista & Padrão PJe-Calc")
    nav_page = st.radio(
        "Navegação",
        [
            "⚖️ Novo Cálculo (.PJC + Planilha)",
            "🗂️ Histórico & Versões Salvas",
            "🛡️ Auditoria & Testes de Regressão",
        ],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown("**Dados do(a) Advogado(a) e Minuta**")
    adv_nome_sb = st.text_input(
        "Nome Completo",
        value=st.session_state["params"].get("advogado_assinante", ""),
        placeholder="Dr(a). Nome do Advogado",
        key="sb_adv_nome",
    )
    adv_oab_sb = st.text_input(
        "Inscrição OAB",
        value=st.session_state["params"].get("oab_advogado", ""),
        placeholder="OAB/RS 00.000",
        key="sb_adv_oab",
    )
    cidade_ass_sb = st.text_input(
        "Cidade da Assinatura (Minuta)",
        value=st.session_state["params"].get("cidade_assinatura", "Gravataí"),
        placeholder="Ex.: Gravataí, Porto Alegre...",
        key="sb_cidade_ass",
    )
    parte_rep_sb = st.selectbox(
        "Parte Representada",
        ["Reclamante", "Reclamado"],
        index=0 if st.session_state["params"].get("parte_representada", "Reclamante") == "Reclamante" else 1,
        key="sb_parte_rep",
    )
    st.session_state["params"]["advogado_assinante"] = adv_nome_sb
    st.session_state["params"]["oab_advogado"] = adv_oab_sb
    st.session_state["params"]["cidade_assinatura"] = cidade_ass_sb
    st.session_state["params"]["parte_representada"] = parte_rep_sb

    st.divider()
    st.markdown("**Motor & Regras Vigentes**")
    st.caption(
        f"• Precisão: `Decimal` (28 dígitos)\n"
        f"• Tabelas oficiais até: `{ULTIMA_COMPETENCIA_TABELA_OFICIAL}`\n"
        f"• Atualização: ADC 58 STF / Lei 14.905/24"
    )

if nav_page == "⚖️ Novo Cálculo (.PJC + Planilha)":
    with st.container(border=True):
        st.markdown(
            '<div><span class="jc-step">PASSO 01</span><strong>Cole a Sentença / Decisões ou Anexe os Autos do Processo</strong></div>',
            unsafe_allow_html=True,
        )
        st.write("")
        sentence_text_input = st.text_area(
            "Texto da Sentença, Embargos de Declaração, Acórdão ou Dispositivo Judicial",
            value=st.session_state["sentence_pasted_text"],
            height=180,
            placeholder="Cole aqui o texto da sentença, embargos e acórdãos (ou anexe os PDFs completos do processo abaixo)...",
        )

        col_up, col_dt = st.columns([3, 1])
        with col_up:
            uploaded_files = st.file_uploader(
                "Anexar PDFs do Processo (Autos, CTPS, Contracheques, Sentença) ou importar .JSON salvo",
                type=["pdf", "txt", "json", "csv", "xlsx", "xls"],
                accept_multiple_files=True,
            )
        with col_dt:
            data_base_input = st.date_input(
                "Data-base da atualização",
                value=st.session_state["data_base_liq"],
                format="DD/MM/YYYY",
            )
            default_api_key = ""
            try:
                default_api_key = st.secrets.get("GEMINI_API_KEY", "") or st.secrets.get("OPENAI_API_KEY", "")
            except Exception:
                default_api_key = os.environ.get("GEMINI_API_KEY", "") or os.environ.get("OPENAI_API_KEY", "")

        lawyer_notes_input = st.text_input(
            "Orientações específicas do advogado para a liquidação (opcional)",
            value=st.session_state["lawyer_notes"],
            placeholder="Ex.: Observar limite de R$ 2.500,00 fixado na sentença; excluir horas extras afastadas em embargos.",
        )

        if st.button("⚖️ PASSO 02 — Analisar Documentos e Gerar Planilha PJe-Calc + .PJC", type="primary", use_container_width=True):
# ... existing code ...
        with st.container(border=True):
            st.markdown(
                '<div><span class="jc-step">PASSO 03</span><strong>Arquivos Gerados (Planilha PJe-Calc PDF + .PJC + JSON + Minuta)</strong></div>',
                unsafe_allow_html=True,
            )
            st.write("")
            pdf_report_bytes = build_pjecalc_pdf(res, aud)
# ... existing code ...
        with tab_visual_edit:
            if p.get("honorarios_juiz_pct", 0) > 0:
                st.info(
                    f"⚖️ **Honorários determinados pelo Juiz na decisão:** {fmt_num(p['honorarios_juiz_pct'], 2)}% "
                    f"(você pode manter ou ajustar o percentual abaixo, bem como a cidade da assinatura para a minuta)."
                )

            with st.form("form_conferencia_calculo"):
                fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1.5])
                edit_proc = fc1.text_input("Processo CNJ", value=p.get("processo", ""))
                edit_rec = fc2.text_input("Reclamante", value=p.get("reclamante", ""))
                edit_recd = fc3.text_input("Reclamado", value=p.get("reclamado", ""))
                edit_cidade_ass = fc4.text_input("Cidade da Minuta", value=p.get("cidade_assinatura", "Gravataí"))

                fd1, fd2, fd3, fd4 = st.columns(4)
# ... existing code ...
                hc1, hc2, hc3, hc4 = st.columns(4)
                edit_hon = hc1.number_input(
                    f"Honorários (%) [Juiz: {fmt_num(p.get('honorarios_juiz_pct', 0), 2)}%]",
                    min_value=0.0,
                    max_value=30.0,
                    value=float(p.get("honorarios_sucumbencia_pct", 0.0)),
                    step=1.0,
                )
                edit_custas = hc2.number_input("Custas Processuais (R$)", min_value=0.0, value=float(p.get("custas_processuais", 0.0)), step=10.0, format="%.2f")
                edit_custas_rec = hc3.checkbox("Custas já recolhidas", value=bool(p.get("custas_ja_recolhidas", False)))
                edit_abat = hc4.number_input("Abatimentos (R$)", min_value=0.0, value=float(p.get("valor_pago_abatimento", 0.0)), step=50.0, format="%.2f")

                edit_confirm = st.checkbox(
                    "Confirmo que conferi os parâmetros e verbas acima para emissão definitiva.",
                    value=bool(p.get("_human_confirmed_ambiguities", False)),
                )

                if st.form_submit_button("🔄 Salvar Ajustes e Regenerar Planilha + .PJC", type="primary", use_container_width=True):
                    p["processo"] = edit_proc.strip()
                    p["reclamante"] = edit_rec.strip()
                    p["reclamado"] = edit_recd.strip()
                    p["cidade_assinatura"] = edit_cidade_ass.strip() or p.get("cidade_assinatura", "Gravataí")
                    p["admissao"] = edit_adm
# ... existing code ...
