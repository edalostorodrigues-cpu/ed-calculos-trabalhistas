# ==============================================================================
# ED CÁLCULOS TRABALHISTAS & PARECER JURÍDICO (PADRÃO PJE-CALC COMPLETO)
# Versão: Full-Stack Auditável (Aritmética Decimal Exata + Ocorrências Mensais
#         + Travas de Coisa Julgada Art. 879 §1º CLT + Exportação PDF/PJC/XLSX
#         + Parecer IA Gemini 3.8 Flash + Persistência Nuvem Supabase)
# ==============================================================================

from __future__ import annotations

import calendar
import datetime as dt
import io
import json
import os
import re
import uuid
import xml.etree.ElementTree as ET
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from fpdf import FPDF
from google import genai
from google.genai import types

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E CHAVE DE LOGIN (DESABILITADA TEMPORARIAMENTE)
# ==============================================================================
st.set_page_config(
    page_title="ED Cálculos Trabalhistas | Padrão PJe-Calc & IA",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Deixe False enquanto ajustamos as abas; mude para True para ativar login/licença
REQUIRE_LOGIN = False

GEMINI_MODEL_PRIMARY = "gemini-3.8-flash"
GEMINI_MODEL_FALLBACK = "gemini-2.5-flash"

TWO_PLACES = Decimal("0.01")
FOUR_PLACES = Decimal("0.0001")


# ==============================================================================
# 2. UTILITÁRIOS DE ARITMÉTICA DECIMAL, FORMATAÇÃO E SECRETS
# ==============================================================================
def D(val: Any) -> Decimal:
    """Converte com segurança qualquer número/string para Decimal exato."""
    if val is None or val == "":
        return Decimal("0")
    if isinstance(val, Decimal):
        return val
    if isinstance(val, (int, float)):
        return Decimal(str(val))
    s = str(val).strip()
    s = re.sub(r"[R$\s%]", "", s)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def money_round(val: Decimal | float | int) -> Decimal:
    return D(val).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)


def fmt_brl(val: Decimal | float | int) -> str:
    v = money_round(val)
    neg = v < 0
    v = abs(v)
    inteiro, dec = f"{v:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
    return f"{'-' if neg else ''}R$ {inteiro_fmt},{dec}"


def fmt_num(val: Decimal | float | int, places: int = 2) -> str:
    q = Decimal("10") ** (-places) if places > 0 else Decimal("1")
    v = D(val).quantize(q, rounding=ROUND_HALF_UP)
    s = f"{v:.{places}f}"
    if "." in s:
        inteiro, dec = s.split(".")
        inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
        return f"{inteiro_fmt},{dec}"
    return f"{int(s):,}".replace(",", ".")
def safe_latin1(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    replacements = {
        "—": "-",
        "–": "-",
        "•": "*",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "…": "...",
        "×": "x",
        "≤": "<=",
        "≥": ">=",
        "§": "Sec.",
        "º": "o",
        "ª": "a",
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    return s.encode("latin-1", errors="replace").decode("latin-1")


def get_secret(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    return os.environ.get(key, default).strip()


def parse_date_safe(val: Any, fallback: dt.date) -> dt.date:
    if isinstance(val, dt.date):
        return val
    if not val:
        return fallback
    s = str(val).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return fallback


# ==============================================================================
# 3. CONEXÃO COM SUPABASE E PERSISTÊNCIA MULTI-CLIENTE NA NUVEM
# ==============================================================================
def get_supabase_headers() -> Optional[Tuple[str, Dict[str, str]]]:
    url = get_secret("SUPABASE_URL").rstrip("/")
    key = get_secret("SUPABASE_KEY")
    if not url or not key:
        return None
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=representation",
    }
    return url, headers


def load_cloud_cases(office_email: str) -> List[Dict[str, Any]]:
    cfg = get_supabase_headers()
    if not cfg:
        return []
    url, headers = cfg
    try:
        endpoint = f"{url}/rest/v1/ed_calculos_clientes?select=*&order=updated_at.desc"
        resp = requests.get(endpoint, headers=headers, timeout=15)
        if resp.status_code >= 400:
            return []
        rows = resp.json() or []
        seen_keys = set()
        cases: List[Dict[str, Any]] = []
        for row in rows:
            row_office = str(row.get("office_email") or "").strip().lower()
            if office_email and row_office and row_office != office_email.lower():
                continue
            payload = row.get("payload") or {}
            if isinstance(payload, str):
                try:
                    payload = json.loads(payload)
                except Exception:
                    payload = {}
            case_id = str(row.get("id") or payload.get("case_id") or uuid.uuid4().hex[:8]).upper()
            proc_num = str(payload.get("process_number") or "").strip()
            client_name = str(row.get("client_name") or payload.get("client_name") or "Cliente Sem Nome").strip()
            dedup_key = f"{proc_num.lower()}|{client_name.lower()}" if proc_num else client_name.lower()
            if dedup_key in seen_keys:
                continue
            seen_keys.add(dedup_key)
            record = {
                "case_id": case_id,
                "client_name": client_name,
                "process_number": proc_num,
                "mode": payload.get("mode", "preliminar"),
                "updated_at": str(row.get("updated_at") or dt.datetime.now().isoformat()),
                "calc_result": payload.get("calc_result"),
                "ai_report": payload.get("ai_report", ""),
                "chat_history": payload.get("chat_history", []),
                "form_snapshot": payload.get("form_snapshot", {}),
            }
            cases.append(record)
        return cases
    except Exception as exc:
        st.warning(f"Aviso ao sincronizar nuvem Supabase: {exc}")
        return []


def upsert_cloud_case(office_email: str, record: Dict[str, Any]) -> Tuple[bool, str]:
    cfg = get_supabase_headers()
    if not cfg:
        return False, "Supabase não configurado nos Secrets."
    url, headers = cfg
    try:
        case_id = str(record.get("case_id") or f"ED-{uuid.uuid4().hex[:8].upper()}")
        payload = {
            "case_id": case_id,
            "process_number": record.get("process_number", ""),
            "mode": record.get("mode", "preliminar"),
            "calc_result": record.get("calc_result"),
            "ai_report": record.get("ai_report", ""),
            "chat_history": record.get("chat_history", []),
            "form_snapshot": record.get("form_snapshot", {}),
        }
        data = {
            "id": case_id,
            "office_email": office_email or "escritorio@edcalculos.com.br",
            "client_name": record.get("client_name", "Cliente"),
            "payload": payload,
            "updated_at": dt.datetime.now(dt.timezone.utc).isoformat(),
        }
        endpoint = f"{url}/rest/v1/ed_calculos_clientes"
        resp = requests.post(endpoint, headers=headers, json=data, timeout=15)
        if resp.status_code >= 400:
            return False, f"Erro HTTP {resp.status_code}: {resp.text}"
        return True, "Salvo na nuvem Supabase."
    except Exception as exc:
        return False, f"Falha ao gravar no Supabase: {exc}"


# ==============================================================================
# 4. CALENDÁRIO TRABALHISTA REAL (COMPETÊNCIAS MÊS A MÊS E DIAS ÚTEIS POR ESCALA)
# ==============================================================================
def build_monthly_competences(
    start_date: dt.date,
    end_date: dt.date,
    work_days_per_week: int = 5,
    exact_total_days: int = 0,
) -> List[Dict[str, Any]]:
    """
    Gera a grade mensal (competência YYYY-MM) no padrão PJe-Calc, contando
    os dias úteis reais do calendário conforme a escala (5 = Seg-Sex, 6 = Seg-Sáb)
    e ajustando proporcionalmente caso seja informado o total exato de dias trabalhados.
    """
    if end_date < start_date:
        end_date = start_date

    allowed_weekdays = {0, 1, 2, 3, 4} if work_days_per_week <= 5 else {0, 1, 2, 3, 4, 5}
    months: List[Dict[str, Any]] = []

    cur_year, cur_month = start_date.year, start_date.month
    while (cur_year, cur_month) <= (end_date.year, end_date.month):
        last_day = calendar.monthrange(cur_year, cur_month)[1]
        m_start = max(start_date, dt.date(cur_year, cur_month, 1))
        m_end = min(end_date, dt.date(cur_year, cur_month, last_day))

        cal_days = (m_end - m_start).days + 1
        work_days = 0
        sundays = 0
        d_iter = m_start
        while d_iter <= m_end:
            if d_iter.weekday() in allowed_weekdays:
                work_days += 1
            if d_iter.weekday() == 6:
                sundays += 1
            d_iter += dt.timedelta(days=1)

        months.append(
            {
                "competencia": f"{cur_month:02d}/{cur_year}",
                "competencia_iso": f"{cur_year}-{cur_month:02d}",
                "inicio": m_start.strftime("%d/%m/%Y"),
                "fim": m_end.strftime("%d/%m/%Y"),
                "dias_corridos": cal_days,
                "dias_uteis_calc": work_days,
                "domingos_feriados": max(1, sundays),
                "avos_15": 1 if cal_days >= 15 else 0,
            }
        )

        if cur_month == 12:
            cur_year += 1
            cur_month = 1
        else:
            cur_month += 1

    raw_total_days = sum(m["dias_uteis_calc"] for m in months)
    if exact_total_days and exact_total_days > 0 and raw_total_days > 0:
        # Distribui proporcionalmente o total exato de dias trabalhados nas competências
        remaining = int(exact_total_days)
        for idx, m in enumerate(months):
            if idx == len(months) - 1:
                m["dias_efetivos"] = max(0, remaining)
            else:
                share = round((m["dias_uteis_calc"] / raw_total_days) * exact_total_days)
                share = min(share, remaining)
                m["dias_efetivos"] = int(share)
                remaining -= int(share)
    else:
        for m in months:
            m["dias_efetivos"] = m["dias_uteis_calc"]

    return months


# ==============================================================================
# 5. MOTOR DE CÁLCULO TRABALHISTA DECIMAL EXATO (PADRÃO PJE-CALC + COISA JULGADA)
# ==============================================================================
def run_pjecalc_engine(params: Dict[str, Any]) -> Dict[str, Any]:
    """
    Motor de liquidação trabalhista com aritmética Decimal exata, apuração
    mês a mês por competência (padrão PJe-Calc) e travas de teto da coisa julgada
    (Art. 879, § 1º CLT / IAC-1000156-59.2020.5.09.0000).
    """
    admissao = parse_date_safe(params.get("admissao"), dt.date(2024, 8, 13))
    rescisao = parse_date_safe(params.get("rescisao"), dt.date(2025, 7, 7))
    if rescisao < admissao:
        rescisao = admissao

    salario_base = D(params.get("salario_base", 0))
    salario_por_fora = D(params.get("salario_por_fora", 0))
    salario_minimo = D(params.get("salario_minimo", "1518.00")) or Decimal("1518.00")
    divisor = D(params.get("divisor", 220)) or Decimal("220")
    work_days_per_week = int(params.get("work_days_per_week", 5) or 5)
    exact_total_days = int(params.get("exact_total_days", 0) or 0)

    tipo_rescisao = str(params.get("tipo_rescisao", "Sem Justa Causa (Iniciativa Empregador)"))
    aviso_indenizado = bool(params.get("aviso_indenizado", False))
    ferias_vencidas_qtd = int(params.get("ferias_vencidas_qtd", 0) or 0)
    multa_477 = bool(params.get("multa_477", False))
    multa_467 = bool(params.get("multa_467", False))

    # Adicionais ambientais
    insalubridade_pct = D(params.get("insalubridade_pct", 0))
    periculosidade_pct = D(params.get("periculosidade_pct", 0))

    # Jornada / Horas Extras / Intervalo / Noturno
    he50_horas_mes = D(params.get("he50_horas_mes", 0))
    he100_horas_mes = D(params.get("he100_horas_mes", 0))
    intrajornada_min_dia = D(params.get("intrajornada_min_dia", 0))
    noturno_horas_mes = D(params.get("noturno_horas_mes", 0))

    # Outras parcelas deferidas
    descontos_devolver = D(params.get("descontos_devolver", 0))
    outras_verbas_valor = D(params.get("outras_verbas_valor", 0))
    outras_verbas_desc = str(params.get("outras_verbas_desc", "Outras Parcelas Deferidas")).strip()

    # Parâmetros de Coisa Julgada / Sentença (Tetos Art. 879 § 1º CLT, Reflexos, Atualização, Honorários)
    aplicar_reflexos = bool(params.get("aplicar_reflexos", True))
    reflexos_em_fgts = bool(params.get("reflexos_em_fgts", True))
    incluir_verbas_rescisorias = bool(params.get("incluir_verbas_rescisorias", False))

    teto_he50 = D(params.get("teto_he50", 0))
    teto_intrajornada = D(params.get("teto_intrajornada", 0))
    teto_adicional = D(params.get("teto_adicional", 0))
    teto_descontos = D(params.get("teto_descontos", 0))
    teto_global = D(params.get("teto_global", 0))

    atualizacao_juros_pct = D(params.get("atualizacao_juros_pct", 0))
    criterio_atualizacao = str(
        params.get("criterio_atualizacao", "ADC 58 STF (IPCA-E pré-judicial + SELIC / IPCA + Art. 406 CC)")
    )
    honorarios_sucumbencia_pct = D(params.get("honorarios_sucumbencia_pct", 0))

    months = build_monthly_competences(
        start_date=admissao,
        end_date=rescisao,
        work_days_per_week=work_days_per_week,
        exact_total_days=exact_total_days,
    )
    total_meses = Decimal(str(len(months)))
    total_dias_efetivos = sum(m["dias_efetivos"] for m in months)
    total_avos = sum(m["avos_15"] for m in months)

    # Anos completos para Aviso Prévio Proporcional (Lei 12.506/2011)
    anos_completos = max(0, (rescisao - admissao).days // 365)
    dias_aviso_previo = min(90, 30 + (3 * anos_completos))

    # Base de cálculo mensal (Súmula 264/TST)
    val_insalub = money_round(salario_minimo * (insalubridade_pct / Decimal("100")))
    val_peric = money_round((salario_base + salario_por_fora) * (periculosidade_pct / Decimal("100")))
    # Não cumula insalubridade e periculosidade (prevalece o mais benéfico - art. 193 §2º CLT)
    if val_peric >= val_insalub:
        adicional_mensal = val_peric
        adicional_nome = f"Adicional de Periculosidade ({fmt_num(periculosidade_pct, 0)}%)"
    else:
        adicional_mensal = val_insalub
        adicional_nome = f"Adicional de Insalubridade ({fmt_num(insalubridade_pct, 0)}% s/ Sal. Mínimo)"

    remuneracao_base_hora = salario_base + salario_por_fora + adicional_mensal
    valor_hora_normal = (remuneracao_base_hora / divisor).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)

    verbas: List[Dict[str, Any]] = []
    ocorrencias: List[Dict[str, Any]] = []

    def add_verba(
        codigo: str,
        descricao: str,
        incidencia: str,
        formula: str,
        valor_bruto_calc: Decimal,
        teto: Decimal = Decimal("0"),
    ) -> Decimal:
        val_calc = money_round(valor_bruto_calc)
        if val_calc <= 0 and teto <= 0:
            return Decimal("0")
        if teto > 0 and val_calc > teto:
            final_val = money_round(teto)
            formula_final = f"{formula} -> TRAVADO NO TETO SENTENCIADO ({fmt_brl(teto)} - Art. 879 §1o CLT)"
        elif teto > 0 and val_calc == 0:
            final_val = money_round(teto)
            formula_final = f"Valor fixado/teto na sentença ({fmt_brl(teto)})"
        else:
            final_val = val_calc
            formula_final = formula

        verbas.append(
            {
                "codigo": codigo,
                "verba": descricao,
                "incidencia": incidencia,
                "formula": formula_final,
                "valor": float(final_val),
                "valor_fmt": fmt_brl(final_val),
            }
        )
        return final_val

    # --------------------------------------------------------------------------
    # A. ADICIONAL AMBIENTAL (INSALUBRIDADE / PERICULOSIDADE)
    # --------------------------------------------------------------------------
    total_adicional = Decimal("0")
    if adicional_mensal > 0:
        bruto_adic = Decimal("0")
        for m in months:
            fator = Decimal(str(m["dias_corridos"])) / Decimal("30")
            if fator > 1:
                fator = Decimal("1")
            v_mes = money_round(adicional_mensal * fator)
            bruto_adic += v_mes
            ocorrencias.append(
                {
                    "competencia": m["competencia"],
                    "verba": adicional_nome,
                    "base_calculo": float(adicional_mensal),
                    "quantidade": float(fator.quantize(TWO_PLACES)),
                    "divisor": 1,
                    "multiplicador": 1.0,
                    "valor_devido": float(v_mes),
                }
            )
        total_adicional = add_verba(
            "010",
            adicional_nome,
            "Salarial / Reflexos",
            f"{fmt_brl(adicional_mensal)}/mês x {len(months)} competências",
            bruto_adic,
            teto_adicional,
        )

    # --------------------------------------------------------------------------
    # B. HORAS EXTRAS 50% E 100%
    # --------------------------------------------------------------------------
    total_he50 = Decimal("0")
    if he50_horas_mes > 0:
        bruto_he50 = Decimal("0")
        for m in months:
            v_mes = money_round(valor_hora_normal * Decimal("1.5") * he50_horas_mes)
            bruto_he50 += v_mes
            ocorrencias.append(
                {
                    "competencia": m["competencia"],
                    "verba": "Horas Extras 50%",
                    "base_calculo": float(remuneracao_base_hora),
                    "quantidade": float(he50_horas_mes),
                    "divisor": float(divisor),
                    "multiplicador": 1.5,
                    "valor_devido": float(v_mes),
                }
            )
        total_he50 = add_verba(
            "020",
            "Horas Extras 50% (Art. 59 CLT)",
            "Salarial",
            f"({fmt_brl(remuneracao_base_hora)} / {int(divisor)}) x 1,50 x {fmt_num(he50_horas_mes)}h/mês x {len(months)} meses",
            bruto_he50,
            teto_he50,
        )

    total_he100 = Decimal("0")
    if he100_horas_mes > 0:
        bruto_he100 = Decimal("0")
        for m in months:
            v_mes = money_round(valor_hora_normal * Decimal("2.0") * he100_horas_mes)
            bruto_he100 += v_mes
            ocorrencias.append(
                {
                    "competencia": m["competencia"],
                    "verba": "Horas Extras 100% (Domingos/Feriados)",
                    "base_calculo": float(remuneracao_base_hora),
                    "quantidade": float(he100_horas_mes),
                    "divisor": float(divisor),
                    "multiplicador": 2.0,
                    "valor_devido": float(v_mes),
                }
            )
        total_he100 = add_verba(
            "025",
            "Horas Extras 100% (Domingos e Feriados)",
            "Salarial",
            f"({fmt_brl(remuneracao_base_hora)} / {int(divisor)}) x 2,00 x {fmt_num(he100_horas_mes)}h/mês x {len(months)} meses",
            bruto_he100,
        )

    # --------------------------------------------------------------------------
    # C. INTERVALO INTRAJORNADA (ART. 71, § 4º CLT) - MÊS A MÊS POR DIAS REAIS
    # --------------------------------------------------------------------------
    total_intrajornada = Decimal("0")
    if intrajornada_min_dia > 0 or teto_intrajornada > 0:
        horas_dia_intra = (intrajornada_min_dia / Decimal("60")).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)
        bruto_intra = Decimal("0")
        for m in months:
            qtd_horas_mes = (horas_dia_intra * Decimal(str(m["dias_efetivos"]))).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
            v_mes = money_round(valor_hora_normal * Decimal("1.5") * qtd_horas_mes)
            bruto_intra += v_mes
            ocorrencias.append(
                {
                    "competencia": m["competencia"],
                    "verba": "Intervalo Intrajornada (Art. 71 §4º CLT)",
                    "base_calculo": float(remuneracao_base_hora),
                    "quantidade": float(qtd_horas_mes),
                    "divisor": float(divisor),
                    "multiplicador": 1.5,
                    "valor_devido": float(v_mes),
                }
            )
        total_intrajornada = add_verba(
            "030",
            "Intervalo Intrajornada Suprimido (+50% - Art. 71 § 4º CLT)",
            "Indenizatória/Salarial",
            f"({fmt_brl(remuneracao_base_hora)} / {int(divisor)}) x 1,50 x {fmt_num(horas_dia_intra)}h/dia x {total_dias_efetivos} dias úteis ({work_days_per_week}d/sem)",
            bruto_intra,
            teto_intrajornada,
        )

    # --------------------------------------------------------------------------
    # D. ADICIONAL NOTURNO (ART. 73 CLT - HORA FICTA 52m30s)
    # --------------------------------------------------------------------------
    total_noturno = Decimal("0")
    if noturno_horas_mes > 0:
        fator_ficto = Decimal("60") / Decimal("52.5")
        bruto_not = Decimal("0")
        for m in months:
            horas_not_fictas = (noturno_horas_mes * fator_ficto).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
            v_mes = money_round(valor_hora_normal * Decimal("0.20") * horas_not_fictas)
            bruto_not += v_mes
            ocorrencias.append(
                {
                    "competencia": m["competencia"],
                    "verba": "Adicional Noturno 20% (Hora 52m30s)",
                    "base_calculo": float(remuneracao_base_hora),
                    "quantidade": float(horas_not_fictas),
                    "divisor": float(divisor),
                    "multiplicador": 0.2,
                    "valor_devido": float(v_mes),
                }
            )
        total_noturno = add_verba(
            "040",
            "Adicional Noturno 20% + Hora Ficta (Art. 73 CLT)",
            "Salarial",
            f"({fmt_brl(remuneracao_base_hora)} / {int(divisor)}) x 20% x (60/52,5) x {fmt_num(noturno_horas_mes)}h/mês x {len(months)} meses",
            bruto_not,
        )

    # --------------------------------------------------------------------------
    # E. DEVOLUÇÃO DE DESCONTOS INDEVIDOS (ART. 462 CLT)
    # --------------------------------------------------------------------------
    total_descontos = Decimal("0")
    if descontos_devolver > 0 or teto_descontos > 0:
        base_desc = descontos_devolver if descontos_devolver > 0 else teto_descontos
        total_descontos = add_verba(
            "050",
            "Devolução de Descontos Indevidos (Art. 462 CLT)",
            "Indenizatória",
            f"Restituição integral de descontos indevidos ({fmt_brl(base_desc)})",
            base_desc,
            teto_descontos,
        )

    # --------------------------------------------------------------------------
    # F. OUTRAS PARCELAS DEFERIDAS
    # --------------------------------------------------------------------------
    total_outras = Decimal("0")
    if outras_verbas_valor > 0:
        total_outras = add_verba(
            "060",
            outras_verbas_desc or "Outras Parcelas Deferidas",
            "Conforme Título",
            f"Parcela apurada/fixada ({fmt_brl(outras_verbas_valor)})",
            outras_verbas_valor,
        )

    # --------------------------------------------------------------------------
    # G. REFLEXOS LEGAIS DAS PARCELAS HABITUAIS (DSR, 13º, FÉRIAS+1/3, AVISO, FGTS)
    # --------------------------------------------------------------------------
    base_reflexos_total = total_adicional + total_he50 + total_he100 + total_intrajornada + total_noturno
    total_reflexos = Decimal("0")
    if aplicar_reflexos and base_reflexos_total > 0:
        media_mensal_reflexo = (base_reflexos_total / total_meses).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)

        # 1. DSR (Súmula 172/TST) sobre horas extras / noturno / intrajornada (exclui adicional mensal fixo)
        base_dsr = total_he50 + total_he100 + total_intrajornada + total_noturno
        ref_dsr = Decimal("0")
        if base_dsr > 0:
            ref_dsr = add_verba(
                "071",
                "Reflexos em Repouso Semanal Remunerado - DSR (Súmula 172/TST)",
                "Salarial",
                f"{fmt_brl(base_dsr)} x (1/6 proporcional aos domingos/feriados)",
                base_dsr / Decimal("6"),
            )

        base_com_dsr_mensal = ((base_reflexos_total + ref_dsr) / total_meses).quantize(FOUR_PLACES, rounding=ROUND_HALF_UP)
        avos_dec = Decimal(str(min(12, max(1, total_avos))))

        ref_13 = add_verba(
            "072",
            "Reflexos em 13º Salário",
            "Salarial",
            f"Média mensal {fmt_brl(base_com_dsr_mensal)} x ({int(avos_dec)}/12 avos)",
            base_com_dsr_mensal * (avos_dec / Decimal("12")),
        )
        ref_ferias = add_verba(
            "073",
            "Reflexos em Férias + 1/3 Constitucional",
            "Indenizatória",
            f"Média mensal {fmt_brl(base_com_dsr_mensal)} x ({int(avos_dec)}/12 avos) x 1,3333",
            base_com_dsr_mensal * (avos_dec / Decimal("12")) * (Decimal("4") / Decimal("3")),
        )
        ref_aviso = Decimal("0")
        if "Sem Justa Causa" in tipo_rescisao or "Indireta" in tipo_rescisao:
            ref_aviso = add_verba(
                "074",
                f"Reflexos em Aviso Prévio Proporcional ({dias_aviso_previo} dias)",
                "Indenizatória",
                f"Média mensal {fmt_brl(base_com_dsr_mensal)} x ({dias_aviso_previo}/30 dias)",
                base_com_dsr_mensal * (Decimal(str(dias_aviso_previo)) / Decimal("30")),
            )

        ref_fgts = Decimal("0")
        if reflexos_em_fgts:
            base_fgts_ref = base_reflexos_total + ref_dsr + ref_13 + ref_aviso
            mult_fgts = Decimal("0.112") if ("Sem Justa Causa" in tipo_rescisao or "Indireta" in tipo_rescisao) else Decimal("0.08")
            pct_label = "11,2% (8% + 40%)" if mult_fgts == Decimal("0.112") else "8%"
            ref_fgts = add_verba(
                "075",
                f"Reflexos em FGTS ({pct_label})",
                "Fundiária",
                f"Base incidente {fmt_brl(base_fgts_ref)} x {pct_label}",
                base_fgts_ref * mult_fgts,
            )
        total_reflexos = ref_dsr + ref_13 + ref_ferias + ref_aviso + ref_fgts

    # --------------------------------------------------------------------------
    # H. VERBAS RESCISÓRIAS (AVALIAÇÃO PRELIMINAR / RESCISÃO INADIMPLIDA)
    # --------------------------------------------------------------------------
    if incluir_verbas_rescisorias and remuneracao_base_hora > 0:
        dias_saldo = rescisao.day
        add_verba(
            "081",
            f"Saldo de Salário ({dias_saldo} dias no mês da rescisão)",
            "Salarial",
            f"({fmt_brl(remuneracao_base_hora)} / 30) x {dias_saldo} dias",
            (remuneracao_base_hora / Decimal("30")) * Decimal(str(dias_saldo)),
        )
        if aviso_indenizado and ("Sem Justa Causa" in tipo_rescisao or "Indireta" in tipo_rescisao):
            add_verba(
                "082",
                f"Aviso Prévio Indenizado ({dias_aviso_previo} dias - Lei 12.506/2011)",
                "Indenizatória",
                f"({fmt_brl(remuneracao_base_hora)} / 30) x {dias_aviso_previo} dias",
                (remuneracao_base_hora / Decimal("30")) * Decimal(str(dias_aviso_previo)),
            )
        avos_ano = rescisao.month if rescisao.day >= 15 else max(0, rescisao.month - 1)
        if "Justa Causa" != tipo_rescisao and avos_ano > 0:
            add_verba(
                "083",
                f"13º Salário Proporcional ({avos_ano}/12 avos)",
                "Salarial",
                f"({fmt_brl(remuneracao_base_hora)} / 12) x {avos_ano} avos",
                (remuneracao_base_hora / Decimal("12")) * Decimal(str(avos_ano)),
            )
            add_verba(
                "084",
                f"Férias Proporcionais + 1/3 ({avos_ano}/12 avos)",
                "Indenizatória",
                f"({fmt_brl(remuneracao_base_hora)} / 12) x {avos_ano} x 1,3333",
                (remuneracao_base_hora / Decimal("12")) * Decimal(str(avos_ano)) * (Decimal("4") / Decimal("3")),
            )
        if ferias_vencidas_qtd > 0:
            add_verba(
                "085",
                f"Férias Vencidas + 1/3 ({ferias_vencidas_qtd} período(s))",
                "Indenizatória",
                f"{fmt_brl(remuneracao_base_hora)} x {ferias_vencidas_qtd} x 1,3333",
                remuneracao_base_hora * Decimal(str(ferias_vencidas_qtd)) * (Decimal("4") / Decimal("3")),
            )
        if "Sem Justa Causa" in tipo_rescisao or "Indireta" in tipo_rescisao:
            fgts_contrato = remuneracao_base_hora * Decimal("0.08") * total_meses
            add_verba(
                "086",
                "Multa Rescisória de 40% sobre o FGTS",
                "Fundiária",
                f"FGTS estimado do contrato ({fmt_brl(fgts_contrato)}) x 40%",
                fgts_contrato * Decimal("0.40"),
            )
        if multa_477:
            add_verba(
                "087",
                "Multa do Art. 477, § 8º da CLT (Atraso Rescisório)",
                "Indenizatória",
                f"1 remuneração mensal ({fmt_brl(remuneracao_base_hora)})",
                remuneracao_base_hora,
            )
        if multa_467:
            subtotal_resc = sum(D(v["valor"]) for v in verbas if v["codigo"].startswith("08"))
            add_verba(
                "088",
                "Multa do Art. 467 da CLT (50% sobre Verbas Rescisórias Incontroversas)",
                "Indenizatória",
                f"50% x {fmt_brl(subtotal_resc)}",
                subtotal_resc * Decimal("0.50"),
            )

    # --------------------------------------------------------------------------
    # I. FECHAMENTO PRINCIPAL + TETO GLOBAL + ATUALIZAÇÃO/JUROS + HONORÁRIOS
    # --------------------------------------------------------------------------
    principal_historico = money_round(sum(D(v["valor"]) for v in verbas))
    if teto_global > 0 and principal_historico > teto_global:
        principal_historico = money_round(teto_global)
        verbas.append(
            {
                "codigo": "090",
                "verba": "Redutor de Teto Global da Condenação (Art. 879 § 1º CLT)",
                "incidencia": "Limite Coisa Julgada",
                "formula": f"Ajuste para respeitar o teto global fixado na sentença ({fmt_brl(teto_global)})",
                "valor": float(principal_historico - money_round(sum(D(v["valor"]) for v in verbas))),
                "valor_fmt": fmt_brl(principal_historico - money_round(sum(D(v["valor"]) for v in verbas))),
            }
        )

    valor_atualizacao_juros = money_round(principal_historico * (atualizacao_juros_pct / Decimal("100")))
    subtotal_corrigido = money_round(principal_historico + valor_atualizacao_juros)
    valor_honorarios = money_round(subtotal_corrigido * (honorarios_sucumbencia_pct / Decimal("100")))
    total_geral = money_round(subtotal_corrigido + valor_honorarios)

    return {
        "admissao": admissao.strftime("%d/%m/%Y"),
        "rescisao": rescisao.strftime("%d/%m/%Y"),
        "admissao_iso": admissao.isoformat(),
        "rescisao_iso": rescisao.isoformat(),
        "total_meses": len(months),
        "total_dias_efetivos": total_dias_efetivos,
        "work_days_per_week": work_days_per_week,
        "salario_base": float(salario_base),
        "salario_por_fora": float(salario_por_fora),
        "remuneracao_base_hora": float(remuneracao_base_hora),
        "divisor": int(divisor),
        "valor_hora_normal": float(valor_hora_normal),
        "tipo_rescisao": tipo_rescisao,
        "criterio_atualizacao": criterio_atualizacao,
        "atualizacao_juros_pct": float(atualizacao_juros_pct),
        "honorarios_sucumbencia_pct": float(honorarios_sucumbencia_pct),
        "principal_historico": float(principal_historico),
        "principal_historico_fmt": fmt_brl(principal_historico),
        "valor_atualizacao_juros": float(valor_atualizacao_juros),
        "valor_atualizacao_juros_fmt": fmt_brl(valor_atualizacao_juros),
        "subtotal_corrigido": float(subtotal_corrigido),
        "subtotal_corrigido_fmt": fmt_brl(subtotal_corrigido),
        "valor_honorarios": float(valor_honorarios),
        "valor_honorarios_fmt": fmt_brl(valor_honorarios),
        "total_geral": float(total_geral),
        "total_geral_fmt": fmt_brl(total_geral),
        "verbas": verbas,
        "ocorrencias": ocorrencias,
        "competencias": months,
    }


# ==============================================================================
# 6. EXPORTAÇÃO NO PADRÃO PJE-CALC: PDF DEMONSTRATIVO, ARQUIVO .PJC E EXCEL
# ==============================================================================
def build_pjecalc_pdf_bytes(calc: Dict[str, Any], office_header: str = "ED Cálculos Trabalhistas") -> bytes:
    """Gera o Demonstrativo de Liquidação de Sentença em PDF no layout padrão PJe-Calc."""
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    # Cabeçalho Institucional Padrão PJe-Calc
    pdf.set_fill_color(28, 43, 66)
    pdf.rect(10, 10, 190, 22, style="F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_xy(14, 13)
    pdf.cell(0, 6, safe_latin1("DEMONSTRATIVO DE CÁLCULO TRABALHISTA (PADRÃO PJE-CALC)"), ln=1)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_x(14)
    pdf.cell(0, 5, safe_latin1(f"{office_header}  |  Emissão: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=1)

    pdf.set_text_color(20, 20, 20)
    pdf.ln(6)

    # Bloco 1: Identificação do Processo e Parâmetros do Contrato
    pdf.set_fill_color(238, 242, 247)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(0, 7, safe_latin1(" 1. IDENTIFICAÇÃO DO PROCESSO E PARÂMETROS DO VÍNCULO"), border=1, ln=1, fill=True)

    pdf.set_font("Helvetica", "", 8.5)
    proc_num = calc.get("process_number") or "Pré-Processual / Art. 840 §1º CLT"
    reclamante = calc.get("client_name") or "Reclamante"
    reclamada = calc.get("defendant_name") or "Reclamada"
    vara = calc.get("court_name") or "Justiça do Trabalho"

    rows_id = [
        (f"Processo: {proc_num}", f"Órgão/Vara: {vara}"),
        (f"Reclamante: {reclamante}", f"Reclamada: {reclamada}"),
        (f"Admissão: {calc.get('admissao')}  |  Saída: {calc.get('rescisao')}", f"Competências: {calc.get('total_meses')} meses ({calc.get('total_dias_efetivos')} dias úteis - {calc.get('work_days_per_week')}d/sem)"),
        (f"Salário-Base: {fmt_brl(calc.get('salario_base', 0))}  |  Divisor: {calc.get('divisor', 220)}", f"Hora Normal (Súm. 264): {fmt_brl(calc.get('valor_hora_normal', 0))}"),
        (f"Critério Atualização: {calc.get('criterio_atualizacao', '')}", ""),
    ]
    for left_txt, right_txt in rows_id:
        if right_txt:
            pdf.cell(95, 6, safe_latin1(f" {left_txt}"), border=1)
            pdf.cell(95, 6, safe_latin1(f" {right_txt}"), border=1, ln=1)
        else:
            pdf.cell(190, 6, safe_latin1(f" {left_txt}"), border=1, ln=1)

    pdf.ln(4)

    # Bloco 2: Planilha Resumo das Verbas Liquidadas (Padrão PJe-Calc)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(0, 7, safe_latin1(" 2. RESUMO DAS VERBAS LIQUIDADAS E MEMÓRIA DE CÁLCULO"), border=1, ln=1, fill=True)

    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(14, 6, "Cód.", border=1, align="C")
    pdf.cell(68, 6, "Descrição da Verba", border=1)
    pdf.cell(76, 6, "Fórmula / Parâmetro Legal", border=1)
    pdf.cell(32, 6, "Valor Apurado", border=1, align="R", ln=1)

    pdf.set_font("Helvetica", "", 7.5)
    for v in calc.get("verbas", []):
        cod = safe_latin1(v.get("codigo", ""))
        nome = safe_latin1(v.get("verba", ""))[:42]
        form = safe_latin1(v.get("formula", ""))[:50]
        val_s = safe_latin1(v.get("valor_fmt", fmt_brl(v.get("valor", 0))))
        pdf.cell(14, 6, cod, border=1, align="C")
        pdf.cell(68, 6, f" {nome}", border=1)
        pdf.cell(76, 6, f" {form}", border=1)
        pdf.cell(32, 6, f"{val_s} ", border=1, align="R", ln=1)

    pdf.ln(4)

    # Bloco 3: Quadro de Fechamento da Liquidação (Principal + Juros + Honorários)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(0, 7, safe_latin1(" 3. FECHAMENTO GERAL DA LIQUIDAÇÃO"), border=1, ln=1, fill=True)

    summary_rows = [
        ("Principal Histórico Apurado (Verbas Deferidas)", calc.get("principal_historico_fmt", fmt_brl(0)), False),
        (f"Atualização Monetária + Juros Legais ({fmt_num(calc.get('atualizacao_juros_pct', 0))}%)", calc.get("valor_atualizacao_juros_fmt", fmt_brl(0)), False),
        ("Subtotal Corrigido ao Reclamante", calc.get("subtotal_corrigido_fmt", fmt_brl(0)), True),
        (f"Honorários Advocatícios Sucumbenciais ({fmt_num(calc.get('honorarios_sucumbencia_pct', 0))}%)", calc.get("valor_honorarios_fmt", fmt_brl(0)), False),
        ("TOTAL GERAL DA EXECUÇÃO / LIQUIDAÇÃO", calc.get("total_geral_fmt", fmt_brl(0)), True),
    ]
    for label, val_str, is_bold in summary_rows:
        pdf.set_font("Helvetica", "B" if is_bold else "", 8.5)
        pdf.cell(148, 6.5, safe_latin1(f" {label}"), border=1, fill=is_bold)
        pdf.cell(42, 6.5, safe_latin1(f"{val_str} "), border=1, align="R", fill=is_bold, ln=1)

    # Bloco 4: Demonstrativo de Ocorrências Mês a Mês (Competências)
    ocorr = calc.get("ocorrencias", [])
    if ocorr:
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(0, 7, safe_latin1(" 4. DEMONSTRATIVO DE OCORRÊNCIAS POR COMPETÊNCIA (PADRÃO PJE-CALC)"), border=1, ln=1, fill=True)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(22, 5.5, "Comp.", border=1, align="C")
        pdf.cell(74, 5.5, "Verba", border=1)
        pdf.cell(28, 5.5, "Base (R$)", border=1, align="R")
        pdf.cell(20, 5.5, "Qtd/Div", border=1, align="C")
        pdf.cell(16, 5.5, "Mult.", border=1, align="C")
        pdf.cell(30, 5.5, "Devido (R$)", border=1, align="R", ln=1)

        pdf.set_font("Helvetica", "", 7.2)
        for item in ocorr[:120]:
            pdf.cell(22, 5, safe_latin1(item.get("competencia", "")), border=1, align="C")
            pdf.cell(74, 5, safe_latin1(f" {str(item.get('verba', ''))[:44]}"), border=1)
            pdf.cell(28, 5, safe_latin1(f"{fmt_num(item.get('base_calculo', 0))} "), border=1, align="R")
            pdf.cell(20, 5, safe_latin1(f"{fmt_num(item.get('quantidade', 0))}/{int(item.get('divisor', 1))}"), border=1, align="C")
            pdf.cell(16, 5, safe_latin1(f"{fmt_num(item.get('multiplicador', 1))}"), border=1, align="C")
            pdf.cell(30, 5, safe_latin1(f"{fmt_num(item.get('valor_devido', 0))} "), border=1, align="R", ln=1)

    raw_out = pdf.output(dest="S")
    if isinstance(raw_out, str):
        return raw_out.encode("latin-1")
    return bytes(raw_out)


def build_pjecalc_pjc_bytes(calc: Dict[str, Any]) -> bytes:
    """
    Gera o arquivo estruturado .PJC (XML compatível com o padrão PJe-Calc Cidadão)
    contendo dados do processo, contrato, verbas liquidadas e ocorrências mensais.
    """
    root = ET.Element("CalculoPJeCalc", versao="2.11.2", sistema="ED-Calculos-Trabalhistas")

    proc_el = ET.SubElement(root, "Processo")
    ET.SubElement(proc_el, "Numero").text = str(calc.get("process_number") or "")
    ET.SubElement(proc_el, "Vara").text = str(calc.get("court_name") or "")
    ET.SubElement(proc_el, "Reclamante").text = str(calc.get("client_name") or "")
    ET.SubElement(proc_el, "Reclamada").text = str(calc.get("defendant_name") or "")

    param_el = ET.SubElement(root, "ParametrosCalculo")
    ET.SubElement(param_el, "DataAdmissao").text = str(calc.get("admissao_iso") or "")
    ET.SubElement(param_el, "DataDemissao").text = str(calc.get("rescisao_iso") or "")
    ET.SubElement(param_el, "UltimaRemuneracao").text = f"{D(calc.get('salario_base', 0)):.2f}"
    ET.SubElement(param_el, "CargaHorariaMensal").text = str(calc.get("divisor", 220))
    ET.SubElement(param_el, "DiasUteisSemana").text = str(calc.get("work_days_per_week", 5))
    ET.SubElement(param_el, "TotalDiasEfetivos").text = str(calc.get("total_dias_efetivos", 0))
    ET.SubElement(param_el, "CriterioCorrecaoJuros").text = str(calc.get("criterio_atualizacao") or "")
    ET.SubElement(param_el, "PercentualAtualizacaoJuros").text = f"{D(calc.get('atualizacao_juros_pct', 0)):.2f}"
    ET.SubElement(param_el, "PercentualHonorariosSucumbencia").text = f"{D(calc.get('honorarios_sucumbencia_pct', 0)):.2f}"

    verbas_el = ET.SubElement(root, "VerbasLiquidadas")
    for v in calc.get("verbas", []):
        v_node = ET.SubElement(verbas_el, "Verba", codigo=str(v.get("codigo", "")))
        ET.SubElement(v_node, "Descricao").text = str(v.get("verba", ""))
        ET.SubElement(v_node, "Incidencia").text = str(v.get("incidencia", ""))
        ET.SubElement(v_node, "Formula").text = str(v.get("formula", ""))
        ET.SubElement(v_node, "ValorApurado").text = f"{D(v.get('valor', 0)):.2f}"

    ocorr_el = ET.SubElement(root, "OcorrenciasMensais")
    for o in calc.get("ocorrencias", []):
        o_node = ET.SubElement(ocorr_el, "Ocorrencia", competencia=str(o.get("competencia", "")))
        ET.SubElement(o_node, "Verba").text = str(o.get("verba", ""))
        ET.SubElement(o_node, "BaseCalculo").text = f"{D(o.get('base_calculo', 0)):.2f}"
        ET.SubElement(o_node, "Quantidade").text = f"{D(o.get('quantidade', 0)):.2f}"
        ET.SubElement(o_node, "Divisor").text = str(o.get("divisor", 1))
        ET.SubElement(o_node, "Multiplicador").text = f"{D(o.get('multiplicador', 1)):.2f}"
        ET.SubElement(o_node, "ValorDevido").text = f"{D(o.get('valor_devido', 0)):.2f}"

    resumo_el = ET.SubElement(root, "ResumoLiquidacao")
    ET.SubElement(resumo_el, "PrincipalHistorico").text = f"{D(calc.get('principal_historico', 0)):.2f}"
    ET.SubElement(resumo_el, "AtualizacaoEJuros").text = f"{D(calc.get('valor_atualizacao_juros', 0)):.2f}"
    ET.SubElement(resumo_el, "SubtotalCorrigido").text = f"{D(calc.get('subtotal_corrigido', 0)):.2f}"
    ET.SubElement(resumo_el, "HonorariosSucumbencia").text = f"{D(calc.get('valor_honorarios', 0)):.2f}"
    ET.SubElement(resumo_el, "TotalGeral").text = f"{D(calc.get('total_geral', 0)):.2f}"

    xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return xml_str


def build_pjecalc_excel_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera pasta de trabalho Excel multi-abas no padrão PJe-Calc."""
    buf = io.BytesIO()
    df_resumo = pd.DataFrame(
        [
            {"Item": "Processo", "Valor": calc.get("process_number", "")},
            {"Item": "Reclamante", "Valor": calc.get("client_name", "")},
            {"Item": "Reclamada", "Valor": calc.get("defendant_name", "")},
            {"Item": "Período do Vínculo", "Valor": f"{calc.get('admissao')} a {calc.get('rescisao')}"},
            {"Item": "Competências / Dias Úteis", "Valor": f"{calc.get('total_meses')} meses / {calc.get('total_dias_efetivos')} dias ({calc.get('work_days_per_week')}d/sem)"},
            {"Item": "Salário-Base / Divisor", "Valor": f"{fmt_brl(calc.get('salario_base', 0))} / {calc.get('divisor', 220)}"},
            {"Item": "Principal Histórico", "Valor": calc.get("principal_historico_fmt", "")},
            {"Item": "Atualização + Juros", "Valor": calc.get("valor_atualizacao_juros_fmt", "")},
            {"Item": "Subtotal Corrigido", "Valor": calc.get("subtotal_corrigido_fmt", "")},
            {"Item": "Honorários Sucumbenciais", "Valor": calc.get("valor_honorarios_fmt", "")},
            {"Item": "TOTAL GERAL", "Valor": calc.get("total_geral_fmt", "")},
        ]
    )
    df_verbas = pd.DataFrame(calc.get("verbas", []))
    df_ocorr = pd.DataFrame(calc.get("ocorrencias", []))

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo_PJeCalc", index=False)
        df_verbas.to_excel(writer, sheet_name="Verbas_Consolidado", index=False)
        if not df_ocorr.empty:
            df_ocorr.to_excel(writer, sheet_name="Ocorrencias_Mes_a_Mes", index=False)

    return buf.getvalue()


# ==============================================================================
# 7. INTEGRAÇÃO GEMINI 3.8 FLASH (EXTRAÇÃO DE SENTENÇA, TRIAGEM E CHAT)
# ==============================================================================
def get_gemini_client() -> Optional[genai.Client]:
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


def build_uploaded_parts(uploaded_files: List[Any]) -> Tuple[List[Any], str]:
    parts: List[Any] = []
    tabular_notes: List[str] = []
    for f in uploaded_files or []:
        if f is None:
            continue
        raw = f.getvalue()
        name = f.name.lower()
        if name.endswith(".pdf"):
            parts.append(types.Part.from_bytes(data=raw, mime_type="application/pdf"))
        elif name.endswith(".png"):
            parts.append(types.Part.from_bytes(data=raw, mime_type="image/png"))
        elif name.endswith((".jpg", ".jpeg")):
            parts.append(types.Part.from_bytes(data=raw, mime_type="image/jpeg"))
        elif name.endswith(".csv"):
            try:
                txt = raw.decode("utf-8", errors="replace")[:25000]
                tabular_notes.append(f"--- Arquivo CSV ({f.name}) ---\n{txt}")
            except Exception:
                pass
        elif name.endswith((".xlsx", ".xls")):
            try:
                df_tmp = pd.read_excel(io.BytesIO(raw))
                tabular_notes.append(f"--- Planilha ({f.name}) ---\n{df_tmp.head(120).to_csv(index=False)}")
            except Exception:
                pass
        elif name.endswith(".txt"):
            tabular_notes.append(f"--- Texto ({f.name}) ---\n{raw.decode('utf-8', errors='replace')[:25000]}")
    return parts, "\n\n".join(tabular_notes)


def call_gemini_with_fallback(client: genai.Client, contents: List[Any], config: Optional[types.GenerateContentConfig] = None) -> str:
    last_err = None
    for model_name in [GEMINI_MODEL_PRIMARY, GEMINI_MODEL_FALLBACK]:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=config,
            )
            return resp.text or ""
        except Exception as exc:
            last_err = exc
    raise RuntimeError(f"Falha nos modelos Gemini ({GEMINI_MODEL_PRIMARY}/{GEMINI_MODEL_FALLBACK}): {last_err}")


def analyze_sentence_with_gemini(
    obs_advogado: str,
    uploaded_files: List[Any],
) -> Dict[str, Any]:
    """
    Lê a sentença / cópia do processo + observações do advogado e extrai JSON
    estruturado com datas reais do vínculo, escala semanal, dias efetivos,
    verbas deferidas/improcedentes, tetos da coisa julgada (Art. 879 §1º CLT),
    critérios de juros/correção e honorários sucumbenciais.
    """
    client = get_gemini_client()
    if not client:
        raise RuntimeError("Chave GEMINI_API_KEY não configurada nos Secrets do Streamlit.")

    parts, tab_text = build_uploaded_parts(uploaded_files)

    prompt = f"""
Você é um Perito Contador Trabalhista Sênior especialista em PJe-Calc e Liquidação de Sentença (Art. 879 da CLT).
Analise rigorosamente os documentos anexados (Sentença, Cópia do Processo, TRCT, Cartões de Ponto, Despachos) e as observações do advogado abaixo.

OBSERVAÇÕES / DECLARAÇÃO DO ADVOGADO:
{obs_advogado or 'Nenhuma observação adicional.'}

{tab_text}

REGRAS CRÍTICAS DE EXTRAÇÃO (SIGA À RISCA):
1. DATAS DO VÍNCULO: Extraia exclusivamente a data de ADMISSÃO e a data de RESCISÃO/SAÍDA do contrato de trabalho (CTPS/TRCT/fundamentação). NUNCA confunda com data de ajuizamento, despacho ou publicação da sentença! Formato YYYY-MM-DD.
2. ESCALA SEMANAL E DIAS TRABALHADOS: Verifique se a jornada reconhecida era Segunda a Sexta (work_days_per_week = 5) ou Segunda a Sábado (work_days_per_week = 6). Se a sentença ou os autos indicarem o número exato de dias úteis trabalhados (ex.: 225 dias), preencha exact_total_days; caso contrário, coloque 0.
3. COISA JULGADA E TETOS DA INICIAL (Art. 879 §1º CLT / IAC TST): Se o dispositivo da sentença limitar a condenação aos valores indicados/estimados na petição inicial (ex.: "no limite do valor estimado de R$ 1.337,60" ou "no valor estimado de R$ 2.000,00"), preencha obrigatoriamente os campos de teto correspondentes!
4. VERBAS IMPROCEDENTES: Tudo que foi julgado improcedente ou que o advogado informou que NÃO TEM DIREITO deve ir com valor 0 e listado em verbas_improcedentes.
5. HONORÁRIOS E ATUALIZAÇÃO: Extraia o percentual de honorários sucumbenciais (ex.: 15) e o critério de correção monetária/juros (ADC 58 STF / IPCA-E / SELIC / IPCA + Art. 406 CC).

Retorne APENAS um JSON válido com esta estrutura exata:
{{
  "process_number": "0000000-00.2025.5.04.0000",
  "court_name": "Vara do Trabalho",
  "client_name": "Nome do Reclamante",
  "defendant_name": "Nome da Reclamada",
  "admissao": "YYYY-MM-DD",
  "rescisao": "YYYY-MM-DD",
  "salario_base": 0.0,
  "salario_por_fora": 0.0,
  "divisor": 220,
  "work_days_per_week": 5,
  "exact_total_days": 0,
  "tipo_rescisao": "Sem Justa Causa (Iniciativa Empregador)",
  "insalubridade_pct": 0.0,
  "periculosidade_pct": 0.0,
  "he50_horas_mes": 0.0,
  "he100_horas_mes": 0.0,
  "intrajornada_min_dia": 0.0,
  "noturno_horas_mes": 0.0,
  "descontos_devolver": 0.0,
  "outras_verbas_desc": "",
  "outras_verbas_valor": 0.0,
  "aplicar_reflexos": true,
  "reflexos_em_fgts": true,
  "incluir_verbas_rescisorias": false,
  "teto_he50": 0.0,
  "teto_intrajornada": 0.0,
  "teto_adicional": 0.0,
  "teto_descontos": 0.0,
  "teto_global": 0.0,
  "criterio_atualizacao": "ADC 58 STF (IPCA-E pré-judicial + SELIC / IPCA + Art. 406 CC)",
  "atualizacao_juros_pct": 0.0,
  "honorarios_sucumbencia_pct": 15.0,
  "verbas_deferidas": ["Lista detalhada de cada parcela deferida e seus parâmetros"],
  "verbas_improcedentes": ["Lista detalhada dos pedidos julgados improcedentes"],
  "parecer_sentenca": "Resumo técnico claro do comando condenatório, limites da coisa julgada, reflexos e cuidados contra excesso de execução."
}}
"""
    contents = parts + [prompt]
    raw_text = call_gemini_with_fallback(
        client,
        contents,
        config=types.GenerateContentConfig(
            temperature=0.1,
            response_mime_type="application/json",
        ),
    )
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return json.loads(cleaned)


def analyze_preliminary_case_with_gemini(
    relato_caso: str,
    uploaded_files: List[Any],
) -> Dict[str, Any]:
    """Triagem preliminar de entrevista + fotos/documentos para Petição Inicial (Art. 840 §1º CLT)."""
    client = get_gemini_client()
    if not client:
        raise RuntimeError("Chave GEMINI_API_KEY não configurada nos Secrets do Streamlit.")

    parts, tab_text = build_uploaded_parts(uploaded_files)

    prompt = f"""
Você é um Advogado Trabalhista Sênior e Perito Calculista (NRs 15/16, CLT, Súmulas do TST e PJe-Calc).
Analise o relato da entrevista inicial e os documentos/fotos anexados para levantar todas as verbas devidas e estimar os parâmetros para liquidação da Petição Inicial (Art. 840, § 1º da CLT).

RELATO DO CASO / ENTREVISTA:
{relato_caso or 'Não informado.'}

{tab_text}

Retorne APENAS um JSON válido com esta estrutura exata:
{{
  "client_name": "Nome do Trabalhador",
  "defendant_name": "Nome da Empresa",
  "admissao": "YYYY-MM-DD",
  "rescisao": "YYYY-MM-DD",
  "salario_base": 0.0,
  "salario_por_fora": 0.0,
  "divisor": 220,
  "work_days_per_week": 5,
  "tipo_rescisao": "Sem Justa Causa (Iniciativa Empregador)",
  "aviso_indenizado": true,
  "ferias_vencidas_qtd": 0,
  "multa_477": false,
  "multa_467": false,
  "insalubridade_pct": 0.0,
  "periculosidade_pct": 0.0,
  "he50_horas_mes": 0.0,
  "he100_horas_mes": 0.0,
  "intrajornada_min_dia": 0.0,
  "noturno_horas_mes": 0.0,
  "descontos_devolver": 0.0,
  "outras_verbas_desc": "",
  "outras_verbas_valor": 0.0,
  "incluir_verbas_rescisorias": true,
  "parecer_juridico": "Parecer técnico completo com enquadramento legal (CLT, NRs 15/16, Súmulas TST), análise de provas/fotos e estratégia processual."
}}
"""
    contents = parts + [prompt]
    raw_text = call_gemini_with_fallback(
        client,
        contents,
        config=types.GenerateContentConfig(
            temperature=0.15,
            response_mime_type="application/json",
        ),
    )
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return json.loads(cleaned)


def run_legal_chat_turn(
    user_question: str,
    case_context: Dict[str, Any],
    chat_history: List[Dict[str, str]],
    mode_label: str,
) -> str:
    client = get_gemini_client()
    if not client:
        return "⚠️ Configure a chave GEMINI_API_KEY nos Secrets para utilizar o Chat Jurídico."

    calc_summary = json.dumps(case_context.get("calc_result") or {}, ensure_ascii=False, indent=2)[:18000]
    ai_report = str(case_context.get("ai_report") or "")[:8000]

    history_txt = ""
    for msg in chat_history[-10:]:
        role = "Advogado" if msg.get("role") == "user" else "Consultor Perito"
        history_txt += f"{role}: {msg.get('content', '')}\n\n"

    prompt = f"""
Você é um Consultor Jurídico Trabalhista Sênior e Perito em PJe-Calc ({mode_label}).
Responda com rigor técnico, citando artigos da CLT, CPC, Súmulas do TST, ADC 58 STF e regras de liquidação.
Se o advogado apontar algum limite da sentença (coisa julgada / Art. 879 §1º CLT), teto de pedido ou divergência em planilha adversa, demonstre a conta exata passo a passo e redija a fundamentação pronta para peticionar se solicitado.

DADOS DO CASO / CÁLCULO ATUAL:
Cliente/Reclamante: {case_context.get('client_name', '')}
Processo: {case_context.get('process_number', '')}
Parecer / Extração da Sentença:
{ai_report}

Memória de Cálculo Atual (JSON):
{calc_summary}

HISTÓRICO DA CONVERSA:
{history_txt}

NOVA PERGUNTA DO ADVOGADO:
{user_question}
"""
    return call_gemini_with_fallback(
        client,
        [prompt],
        config=types.GenerateContentConfig(temperature=0.25),
    )


# ==============================================================================
# 8. INICIALIZAÇÃO DO ESTADO DA SESSÃO
# ==============================================================================
if "office_email" not in st.session_state:
    st.session_state["office_email"] = get_secret("ADMIN_EMAIL", "escritorio@edcalculos.com.br")
if "office_header" not in st.session_state:
    st.session_state["office_header"] = "ED Cálculos Trabalhistas & Parecer Jurídico"
if "cloud_cases" not in st.session_state:
    st.session_state["cloud_cases"] = load_cloud_cases(st.session_state["office_email"])
if "prelim_extracted" not in st.session_state:
    st.session_state["prelim_extracted"] = {}
if "prelim_calc_result" not in st.session_state:
    st.session_state["prelim_calc_result"] = None
if "prelim_chat_current" not in st.session_state:
    st.session_state["prelim_chat_current"] = []
if "sentence_extracted" not in st.session_state:
    st.session_state["sentence_extracted"] = {}
if "sentence_calc_result" not in st.session_state:
    st.session_state["sentence_calc_result"] = None
if "sentence_chat_current" not in st.session_state:
    st.session_state["sentence_chat_current"] = []


def upsert_local_and_cloud_case(
    client_name: str,
    process_number: str,
    mode: str,
    calc_result: Dict[str, Any],
    ai_report: str,
    chat_history: List[Dict[str, str]],
) -> Dict[str, Any]:
    """Atualiza ou cria 1 único registro por processo/cliente (sem duplicar no seletor)."""
    cases: List[Dict[str, Any]] = st.session_state["cloud_cases"]
    clean_client = (client_name or "Cliente Sem Nome").strip()
    clean_proc = (process_number or "").strip()

    existing = None
    for c in cases:
        c_proc = str(c.get("process_number") or "").strip()
        c_name = str(c.get("client_name") or "").strip()
        if clean_proc and c_proc and clean_proc.lower() == c_proc.lower():
            existing = c
            break
        if not clean_proc and clean_client.lower() == c_name.lower():
            existing = c
            break

    if existing:
        existing["client_name"] = clean_client
        existing["process_number"] = clean_proc
        existing["mode"] = mode
        existing["calc_result"] = calc_result
        if ai_report:
            existing["ai_report"] = ai_report
        if chat_history:
            existing["chat_history"] = chat_history
        existing["updated_at"] = dt.datetime.now().isoformat()
        target = existing
    else:
        target = {
            "case_id": f"ED-{uuid.uuid4().hex[:8].upper()}",
            "client_name": clean_client,
            "process_number": clean_proc,
            "mode": mode,
            "updated_at": dt.datetime.now().isoformat(),
            "calc_result": calc_result,
            "ai_report": ai_report,
            "chat_history": chat_history or [],
        }
        cases.insert(0, target)

    upsert_cloud_case(st.session_state["office_email"], target)
    return target


# ==============================================================================
# 9. BARRA LATERAL (NAVEGAÇÃO, TIMBRE DO ESCRITÓRIO E STATUS DA NUVEM)
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas")
    st.caption("Motor Decimal Padrão PJe-Calc + IA Jurídica")

    menu = st.radio(
        "Selecione o Módulo:",
        [
            "1. Avaliação Preliminar & Ajuizamento",
            "2. Liquidação de Sentença (PJe-Calc)",
            "3. Carteira de Clientes & Relatórios",
        ],
        index=1,
    )

        if st.button("🔍 Analisar Caso com IA e Calcular Petição Inicial", type="primary", key="btn_prelim_ai"):
            with st.spinner("Analisando fatos, documentos, NRs 15/16 e calculando no padrão PJe-Calc..."):
                try:
                    ext = analyze_preliminary_case_with_gemini(relato_txt, prelim_files)
                    st.session_state["prelim_extracted"] = ext
                    calc_auto = run_pjecalc_engine(
                        {
                            "admissao": parse_date_safe(ext.get("admissao"), dt.date(2024, 2, 1)),
                            "rescisao": parse_date_safe(ext.get("rescisao"), dt.date.today()),
                            "salario_base": float(ext.get("salario_base") or 0.0),
                            "salario_por_fora": float(ext.get("salario_por_fora") or 0.0),
                            "divisor": int(ext.get("divisor") or 220),
                            "work_days_per_week": int(ext.get("work_days_per_week") or 5),
                            "insalubridade_pct": float(ext.get("insalubridade_pct") or 0.0),
                            "periculosidade_pct": float(ext.get("periculosidade_pct") or 0.0),
                            "he50_horas_mes": float(ext.get("he50_horas_mes") or 0.0),
                            "he100_horas_mes": float(ext.get("he100_horas_mes") or 0.0),
                            "intrajornada_min_dia": float(ext.get("intrajornada_min_dia") or 0.0),
                            "noturno_horas_mes": float(ext.get("noturno_horas_mes") or 0.0),
                            "descontos_devolver": float(ext.get("descontos_devolver") or 0.0),
                            "outras_verbas_desc": str(ext.get("outras_verbas_desc") or ""),
                            "outras_verbas_valor": float(ext.get("outras_verbas_valor") or 0.0),
                            "incluir_verbas_rescisorias": bool(ext.get("incluir_verbas_rescisorias", True)),
                            "aviso_indenizado": bool(ext.get("aviso_indenizado", True)),
                            "ferias_vencidas_qtd": int(ext.get("ferias_vencidas_qtd") or 0),
                            "multa_477": bool(ext.get("multa_477", False)),
                            "multa_467": bool(ext.get("multa_467", False)),
                            "aplicar_reflexos": True,
                            "reflexos_em_fgts": True,
                        }
                    )
                    calc_auto["client_name"] = str(ext.get("client_name") or "Cliente Inicial")
                    calc_auto["defendant_name"] = str(ext.get("defendant_name") or "Reclamada")
                    calc_auto["process_number"] = "Pré-Processual (Art. 840 §1º CLT)"
                    calc_auto["court_name"] = "Petição Inicial Trabalhista"
                    st.session_state["prelim_calc_result"] = calc_auto
                    upsert_local_and_cloud_case(
                        client_name=calc_auto["client_name"],
                        process_number="",
                        mode="preliminar",
                        calc_result=calc_auto,
                        ai_report=ext.get("parecer_juridico", ""),
                        chat_history=st.session_state.get("prelim_chat_current", []),
                    )
                    st.success(f"Análise e cálculos concluídos automaticamente! Total estimado: {calc_auto['total_geral_fmt']}")
                except Exception as exc:
                    st.error(f"Erro na análise preliminar: {exc}")

    p_ext = st.session_state.get("prelim_extracted") or {}
    if p_ext.get("parecer_juridico"):
        with st.expander("📋 Parecer Jurídico Técnico da IA (Clique para expandir/recolher)", expanded=True):
            st.markdown(p_ext["parecer_juridico"])

    with st.container(border=True):
        st.subheader("2. Parâmetros Contratuais e Verbas para Liquidação da Inicial")
        c1, c2, c3, c4 = st.columns([2.2, 2.2, 1.5, 1.5])
        p_client = c1.text_input("Nome do Cliente (Reclamante) *", value=str(p_ext.get("client_name") or ""))
        p_def = c2.text_input("Empresa Reclamada", value=str(p_ext.get("defendant_name") or ""))
        p_adm = c3.date_input("Data de Admissão", value=parse_date_safe(p_ext.get("admissao"), dt.date(2024, 2, 1)), key="p_adm")
        p_res = c4.date_input("Data de Saída / Projeção", value=parse_date_safe(p_ext.get("rescisao"), dt.date.today()), key="p_res")

        c5, c6, c7, c8 = st.columns(4)
        p_sal = c5.number_input("Salário-Base (R$)", min_value=0.0, value=float(p_ext.get("salario_base") or 2200.0), step=50.0, key="p_sal")
        p_fora = c6.number_input("Salário 'Por Fora' (R$)", min_value=0.0, value=float(p_ext.get("salario_por_fora") or 0.0), step=50.0, key="p_fora")
        p_div = c7.number_input("Divisor Mensal", min_value=100, max_value=240, value=int(p_ext.get("divisor") or 220), step=10, key="p_div")
        p_esc = c8.selectbox("Escala Semanal", [5, 6], index=0 if int(p_ext.get("work_days_per_week") or 5) == 5 else 1, format_func=lambda x: f"{x} dias/semana ({'Seg-Sex' if x==5 else 'Seg-Sáb'})", key="p_esc")

        c9, c10, c11, c12 = st.columns(4)
        p_insal = c9.selectbox("Insalubridade (%)", [0.0, 10.0, 20.0, 40.0], index=[0.0, 10.0, 20.0, 40.0].index(float(p_ext.get("insalubridade_pct") or 0.0)) if float(p_ext.get("insalubridade_pct") or 0.0) in [0.0, 10.0, 20.0, 40.0] else 0, key="p_insal")
        p_peric = c10.selectbox("Periculosidade (%)", [0.0, 30.0], index=1 if float(p_ext.get("periculosidade_pct") or 0.0) >= 30.0 else 0, key="p_peric")
        p_he50 = c11.number_input("Horas Extras 50% / mês", min_value=0.0, value=float(p_ext.get("he50_horas_mes") or 0.0), step=1.0, key="p_he50")
        p_intra = c12.number_input("Intervalo Suprimido (min/dia)", min_value=0.0, max_value=120.0, value=float(p_ext.get("intrajornada_min_dia") or 0.0), step=5.0, key="p_intra")

        c13, c14, c15, c16 = st.columns(4)
        p_not = c13.number_input("Horas Noturnas / mês", min_value=0.0, value=float(p_ext.get("noturno_horas_mes") or 0.0), step=1.0, key="p_not")
        p_desc = c14.number_input("Descontos Indevidos (R$)", min_value=0.0, value=float(p_ext.get("descontos_devolver") or 0.0), step=50.0, key="p_desc")
        p_resc_chk = c15.checkbox("Incluir Verbas Rescisórias", value=bool(p_ext.get("incluir_verbas_rescisorias", True)), key="p_resc_chk")
        p_aviso_chk = c16.checkbox("Aviso Prévio Indenizado", value=bool(p_ext.get("aviso_indenizado", True)), key="p_aviso_chk")

        if st.button("🧮 Calcular Liquidação da Inicial e Salvar na Nuvem", type="primary", key="btn_calc_prelim"):
            calc_res = run_pjecalc_engine(
                {
                    "admissao": p_adm,
                    "rescisao": p_res,
                    "salario_base": p_sal,
                    "salario_por_fora": p_fora,
                    "divisor": p_div,
                    "work_days_per_week": p_esc,
                    "insalubridade_pct": p_insal,
                    "periculosidade_pct": p_peric,
                    "he50_horas_mes": p_he50,
                    "intrajornada_min_dia": p_intra,
                    "noturno_horas_mes": p_not,
                    "descontos_devolver": p_desc,
                    "incluir_verbas_rescisorias": p_resc_chk,
                    "aviso_indenizado": p_aviso_chk,
                    "aplicar_reflexos": True,
                    "reflexos_em_fgts": True,
                }
            )
            calc_res["client_name"] = p_client or "Cliente Inicial"
            calc_res["defendant_name"] = p_def or "Reclamada"
            calc_res["process_number"] = "Pré-Processual (Art. 840 §1º CLT)"
            calc_res["court_name"] = "Petição Inicial Trabalhista"
            st.session_state["prelim_calc_result"] = calc_res
            upsert_local_and_cloud_case(
                client_name=calc_res["client_name"],
                process_number="",
                mode="preliminar",
                calc_result=calc_res,
                ai_report=p_ext.get("parecer_juridico", ""),
                chat_history=st.session_state.get("prelim_chat_current", []),
            )
            st.success(f"Liquidação calculada e salva! Valor da Causa estimado: {calc_res['total_geral_fmt']}")

    if st.session_state.get("prelim_calc_result"):
        res_p = st.session_state["prelim_calc_result"]
        st.markdown("### Quadro de Pedidos Liquidados (Art. 840, § 1º CLT)")
        if res_p.get("verbas"):
            st.dataframe(pd.DataFrame(res_p["verbas"])[["codigo", "verba", "formula", "valor_fmt"]], use_container_width=True, hide_index=True)
        else:
            st.warning("Nenhuma verba foi calculada porque todos os parâmetros de quantidade/valor estão zerados.")
        st.metric("Valor Total Liquidado da Causa", res_p["total_geral_fmt"])

        d1, d2, d3 = st.columns(3)
        safe_cli = re.sub(r"[^0-9A-Za-z_-]+", "_", res_p.get("client_name", "cliente"))
        d1.download_button("📄 Baixar Demonstrativo PJe-Calc (.PDF)", build_pjecalc_pdf_bytes(res_p, st.session_state["office_header"]), f"Inicial_PJeCalc_{safe_cli}.pdf", "application/pdf", use_container_width=True)
        d2.download_button("⚖️ Baixar Arquivo PJe-Calc (.PJC)", build_pjecalc_pjc_bytes(res_p), f"Inicial_PJeCalc_{safe_cli}.pjc", "application/xml", use_container_width=True)
        if st.button("⚖️ Ler Sentença / Processo e Calcular Liquidação Pronta (PJe-Calc)", type="primary", key="btn_read_sentence"):
            with st.spinner("Lendo sentença, extraindo comando condenatório e calculando liquidação pronta no padrão PJe-Calc..."):
                try:
                    ext_s = analyze_sentence_with_gemini(obs_sentenca, sent_files)
                    st.session_state["sentence_extracted"] = ext_s
                    res_auto_s = run_pjecalc_engine(
                        {
                            "admissao": parse_date_safe(ext_s.get("admissao"), dt.date(2024, 8, 13)),
                            "rescisao": parse_date_safe(ext_s.get("rescisao"), dt.date(2025, 7, 7)),
                            "salario_base": float(ext_s.get("salario_base") or 0.0),
                            "salario_por_fora": float(ext_s.get("salario_por_fora") or 0.0),
                            "divisor": int(ext_s.get("divisor") or 220),
                            "work_days_per_week": int(ext_s.get("work_days_per_week") or 5),
                            "exact_total_days": int(ext_s.get("exact_total_days") or 0),
                            "insalubridade_pct": float(ext_s.get("insalubridade_pct") or 0.0),
                            "periculosidade_pct": float(ext_s.get("periculosidade_pct") or 0.0),
                            "he50_horas_mes": float(ext_s.get("he50_horas_mes") or 0.0),
                            "he100_horas_mes": float(ext_s.get("he100_horas_mes") or 0.0),
                            "intrajornada_min_dia": float(ext_s.get("intrajornada_min_dia") or 0.0),
                            "noturno_horas_mes": float(ext_s.get("noturno_horas_mes") or 0.0),
                            "descontos_devolver": float(ext_s.get("descontos_devolver") or 0.0),
                            "outras_verbas_desc": str(ext_s.get("outras_verbas_desc") or "Outras Verbas Deferidas"),
                            "outras_verbas_valor": float(ext_s.get("outras_verbas_valor") or 0.0),
                            "aplicar_reflexos": bool(ext_s.get("aplicar_reflexos", True)),
                            "reflexos_em_fgts": bool(ext_s.get("reflexos_em_fgts", True)),
                            "incluir_verbas_rescisorias": False,
                            "teto_intrajornada": float(ext_s.get("teto_intrajornada") or 0.0),
                            "teto_descontos": float(ext_s.get("teto_descontos") or 0.0),
                            "teto_he50": float(ext_s.get("teto_he50") or 0.0),
                            "teto_adicional": float(ext_s.get("teto_adicional") or 0.0),
                            "teto_global": float(ext_s.get("teto_global") or 0.0),
                            "criterio_atualizacao": str(
                                ext_s.get("criterio_atualizacao")
                                or "ADC 58 STF (IPCA-E pré-judicial + SELIC / IPCA + Art. 406 CC)"
                            ),
                            "atualizacao_juros_pct": float(ext_s.get("atualizacao_juros_pct") or 0.0),
                            "honorarios_sucumbencia_pct": float(ext_s.get("honorarios_sucumbencia_pct") or 15.0),
                        }
                    )
                    res_auto_s["process_number"] = str(ext_s.get("process_number") or "").strip()
                    res_auto_s["client_name"] = str(ext_s.get("client_name") or "Reclamante").strip()
                    res_auto_s["defendant_name"] = str(ext_s.get("defendant_name") or "Reclamada").strip()
                    res_auto_s["court_name"] = str(ext_s.get("court_name") or "Vara do Trabalho").strip()
                    st.session_state["sentence_calc_result"] = res_auto_s
                    upsert_local_and_cloud_case(
                        client_name=res_auto_s["client_name"],
                        process_number=res_auto_s["process_number"],
                        mode="liquidacao_sentenca",
                        calc_result=res_auto_s,
                        ai_report=ext_s.get("parecer_sentenca", ""),
                        chat_history=st.session_state.get("sentence_chat_current", []),
                    )
                    st.success(f"Sentença lida e liquidação pronta no padrão PJe-Calc! Total Geral: {res_auto_s['total_geral_fmt']}")
                except Exception as exc:
                    st.error(f"Erro ao analisar o processo via Gemini: {exc}")

    s_ext = st.session_state.get("sentence_extracted") or {}

    if s_ext:
        col_def, col_imp = st.columns(2)
        with col_def:
            with st.container(border=True):
                st.markdown("#### ✅ Verbas DEFERIDAS no Dispositivo")
                v_defs = s_ext.get("verbas_deferidas") or []
                if v_defs:
                    for item in v_defs:
                        st.markdown(f"- {item}")
                else:
                    st.info("Nenhuma verba deferida listada automaticamente.")
        with col_imp:
            with st.container(border=True):
                st.markdown("#### ❌ Verbas IMPROCEDENTES / Excluídas")
                v_imps = s_ext.get("verbas_improcedentes") or []
                if v_imps:
                    for item in v_imps:
                        st.markdown(f"- {item}")
                else:
                    st.caption("Nenhum pedido improcedente destacado.")

        if s_ext.get("parecer_sentenca"):
            with st.expander("📋 Parecer Técnico do Comando Condenatório e Limites da Coisa Julgada", expanded=True):
                st.markdown(s_ext["parecer_sentenca"])

    with st.container(border=True):
        st.subheader("2. Conferência dos Parâmetros da Liquidação (Motor PJe-Calc + Coisa Julgada)")

        sp1, sp2, sp3, sp4 = st.columns([1.8, 2.2, 2.2, 1.8])
        s_proc = sp1.text_input("Número do Processo *", value=str(s_ext.get("process_number") or "0000000-00.2025.5.04.0000"))
        s_client = sp2.text_input("Reclamante *", value=str(s_ext.get("client_name") or ""))
        s_def = sp3.text_input("Reclamada *", value=str(s_ext.get("defendant_name") or ""))
        s_court = sp4.text_input("Vara do Trabalho", value=str(s_ext.get("court_name") or "Vara do Trabalho"))

        sd1, sd2, sd3, sd4 = st.columns(4)
        s_adm = sd1.date_input("Data de Admissão (Vínculo)", value=parse_date_safe(s_ext.get("admissao"), dt.date(2024, 8, 13)), key="s_adm")
        s_res = sd2.date_input("Data de Rescisão (Vínculo)", value=parse_date_safe(s_ext.get("rescisao"), dt.date(2025, 7, 7)), key="s_res")
        s_esc = sd3.selectbox(
            "Escala Semanal Reconhecida",
            [5, 6],
            index=0 if int(s_ext.get("work_days_per_week") or 5) == 5 else 1,
            format_func=lambda x: f"{x} dias/sem ({'Seg-Sex' if x==5 else 'Seg-Sáb'})",
            key="s_esc",
        )
        s_exact_days = sd4.number_input(
            "Dias Trabalhados Exatos (0 = auto)",
            min_value=0,
            max_value=3650,
            value=int(s_ext.get("exact_total_days") or 0),
            step=1,
            help="Informe o número exato de dias úteis trabalhados (ex.: 225) ou deixe 0 para o calendário calcular automaticamente.",
            key="s_exact_days",
        )

        sv1, sv2, sv3, sv4 = st.columns(4)
        s_sal = sv1.number_input("Salário-Base (R$)", min_value=0.0, value=float(s_ext.get("salario_base") or 2228.60), step=10.0, key="s_sal")
        s_div = sv2.number_input("Divisor Mensal", min_value=100, max_value=240, value=int(s_ext.get("divisor") or 220), step=10, key="s_div")
        s_intra = sv3.number_input("Intervalo Intrajornada (min/dia)", min_value=0.0, max_value=120.0, value=float(s_ext.get("intrajornada_min_dia") or 0.0), step=5.0, key="s_intra")
        s_he50 = sv4.number_input("Horas Extras 50% / mês", min_value=0.0, value=float(s_ext.get("he50_horas_mes") or 0.0), step=1.0, key="s_he50")

        sv5, sv6, sv7, sv8 = st.columns(4)
        s_desc = sv5.number_input("Restituição de Descontos (R$)", min_value=0.0, value=float(s_ext.get("descontos_devolver") or 0.0), step=50.0, key="s_desc")
        s_insal = sv6.selectbox("Insalubridade (%)", [0.0, 10.0, 20.0, 40.0], index=[0.0, 10.0, 20.0, 40.0].index(float(s_ext.get("insalubridade_pct") or 0.0)) if float(s_ext.get("insalubridade_pct") or 0.0) in [0.0, 10.0, 20.0, 40.0] else 0, key="s_insal")
        s_peric = sv7.selectbox("Periculosidade (%)", [0.0, 30.0], index=1 if float(s_ext.get("periculosidade_pct") or 0.0) >= 30.0 else 0, key="s_peric")
        s_outras_val = sv8.number_input("Outras Verbas Deferidas (R$)", min_value=0.0, value=float(s_ext.get("outras_verbas_valor") or 0.0), step=50.0, key="s_outras_val")

        st.markdown("##### 🔒 Travas de Coisa Julgada (Tetos da Inicial - Art. 879, § 1º CLT), Reflexos e Honorários")
        st1, st2, st3, st4 = st.columns(4)
        s_teto_intra = st1.number_input("Teto Intervalo Intrajornada (R$)", min_value=0.0, value=float(s_ext.get("teto_intrajornada") or 0.0), step=10.0, help="Se a sentença limitou ao valor da inicial (ex.: R$ 1.337,60), informe aqui.", key="s_teto_intra")
        s_teto_desc = st2.number_input("Teto Descontos Indevidos (R$)", min_value=0.0, value=float(s_ext.get("teto_descontos") or 0.0), step=10.0, key="s_teto_desc")
        s_teto_he50 = st3.number_input("Teto Horas Extras 50% (R$)", min_value=0.0, value=float(s_ext.get("teto_he50") or 0.0), step=10.0, key="s_teto_he50")
        s_teto_global = st4.number_input("Teto Global Condenação (R$)", min_value=0.0, value=float(s_ext.get("teto_global") or 0.0), step=50.0, key="s_teto_global")

        sh1, sh2, sh3, sh4 = st.columns(4)
        s_ref_chk = sh1.checkbox("Calcular Reflexos Deferidos", value=bool(s_ext.get("aplicar_reflexos", True)), key="s_ref_chk")
        s_fgts_chk = sh2.checkbox("Incluir Reflexos em FGTS", value=bool(s_ext.get("reflexos_em_fgts", True)), key="s_fgts_chk")
        s_juros_pct = sh3.number_input("Atualização + Juros Acumulados (%)", min_value=0.0, max_value=500.0, value=float(s_ext.get("atualizacao_juros_pct") or 0.0), step=0.5, key="s_juros_pct")
        s_hon_pct = sh4.number_input("Honorários Sucumbência (%)", min_value=0.0, max_value=30.0, value=float(s_ext.get("honorarios_sucumbencia_pct") or 15.0), step=1.0, key="s_hon_pct")

        s_crit_atu = st.text_input(
            "Critério de Correção Monetária e Juros (Sentença / ADC 58 STF):",
            value=str(s_ext.get("criterio_atualizacao") or "ADC 58 STF (IPCA-E pré-judicial + SELIC / IPCA + Art. 406 CC)"),
            key="s_crit_atu",
        )

        conf_liq = st.checkbox(
            "Confirmo que conferi o dispositivo da sentença e os tetos da coisa julgada acima.",
            value=True,
            key="conf_liq_chk",
        )

        if st.button("🧮 Calcular Liquidação de Sentença (Padrão PJe-Calc)", type="primary", disabled=not conf_liq, key="btn_calc_sentence"):
            res_s = run_pjecalc_engine(
                {
                    "admissao": s_adm,
                    "rescisao": s_res,
                    "salario_base": s_sal,
                    "divisor": s_div,
                    "work_days_per_week": s_esc,
                    "exact_total_days": s_exact_days,
                    "insalubridade_pct": s_insal,
                    "periculosidade_pct": s_peric,
                    "he50_horas_mes": s_he50,
                    "intrajornada_min_dia": s_intra,
                    "descontos_devolver": s_desc,
                    "outras_verbas_desc": str(s_ext.get("outras_verbas_desc") or "Outras Verbas Deferidas"),
                    "outras_verbas_valor": s_outras_val,
                    "aplicar_reflexos": s_ref_chk,
                    "reflexos_em_fgts": s_fgts_chk,
                    "incluir_verbas_rescisorias": False,
                    "teto_intrajornada": s_teto_intra,
                    "teto_descontos": s_teto_desc,
                    "teto_he50": s_teto_he50,
                    "teto_global": s_teto_global,
                    "criterio_atualizacao": s_crit_atu,
                    "atualizacao_juros_pct": s_juros_pct,
                    "honorarios_sucumbencia_pct": s_hon_pct,
                }
            )
            res_s["process_number"] = s_proc.strip()
            res_s["client_name"] = s_client.strip() or "Reclamante"
            res_s["defendant_name"] = s_def.strip() or "Reclamada"
            res_s["court_name"] = s_court.strip() or "Vara do Trabalho"

            st.session_state["sentence_calc_result"] = res_s
            upsert_local_and_cloud_case(
                client_name=res_s["client_name"],
                process_number=res_s["process_number"],
                mode="liquidacao_sentenca",
                calc_result=res_s,
                ai_report=s_ext.get("parecer_sentenca", ""),
                chat_history=st.session_state.get("sentence_chat_current", []),
            )
            st.success(f"Liquidação calculada no padrão PJe-Calc e salva! Total Geral: {res_s['total_geral_fmt']}")

    # Exibição do Demonstrativo PJe-Calc e Botões de Exportação (PDF + .PJC + Excel)
    if st.session_state.get("sentence_calc_result"):
        res_liq = st.session_state["sentence_calc_result"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Principal Histórico", res_liq["principal_historico_fmt"])
        m2.metric("Atualização + Juros", res_liq["valor_atualizacao_juros_fmt"])
        m3.metric("Honorários Sucumbenciais", res_liq["valor_honorarios_fmt"])
        m4.metric("TOTAL GERAL DA EXECUÇÃO", res_liq["total_geral_fmt"])

        tab_res_v, tab_ocorr_m = st.tabs(["📑 Resumo das Verbas Liquidadas", "📅 Ocorrências Mês a Mês (PJe-Calc)"])
        with tab_res_v:
            if res_liq.get("verbas"):
                st.dataframe(
                    pd.DataFrame(res_liq["verbas"])[["codigo", "verba", "incidencia", "formula", "valor_fmt"]],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.warning("Nenhuma verba foi apurada porque os parâmetros acima estão zerados (ex.: preencha o Intervalo Intrajornada, Restituição de Descontos ou clique em 'Ler Sentença' primeiro).")
        with tab_ocorr_m:
            if res_liq.get("ocorrencias"):
                st.dataframe(pd.DataFrame(res_liq["ocorrencias"]), use_container_width=True, hide_index=True)
            else:
                st.info("As verbas liquidadas nesta sentença são de parcela única / teto fixo.")

        st.markdown("### 3. Arquivos Prontos para Protocolar (Padrão PJe-Calc: PDF + .PJC + Excel)")
        dl1, dl2, dl3 = st.columns(3)

        pdf_pjecalc_bytes = build_pjecalc_pdf_bytes(res_liq, st.session_state["office_header"])
        pjc_xml_bytes = build_pjecalc_pjc_bytes(res_liq)
        xlsx_pjecalc_bytes = build_pjecalc_excel_bytes(res_liq)

        safe_proc = re.sub(r"[^0-9A-Za-z_-]+", "_", str(res_liq.get("process_number") or "processo"))

        dl1.download_button(
            label="📄 Baixar Demonstrativo PJe-Calc (.PDF)",
            data=pdf_pjecalc_bytes,
            file_name=f"Demonstrativo_PJeCalc_{safe_proc}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_liq_pdf_pjecalc",
        )
        dl2.download_button(
            label="⚖️ Baixar Arquivo PJe-Calc (.PJC)",
            data=pjc_xml_bytes,
            file_name=f"Calculo_PJeCalc_{safe_proc}.pjc",
            mime="application/xml",
            use_container_width=True,
            key="dl_liq_pjc_xml",
        )
        dl3.download_button(
            label="📊 Baixar Planilha Detalhada (.XLSX)",
            data=xlsx_pjecalc_bytes,
            file_name=f"Liquidacao_PJeCalc_{safe_proc}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_liq_xlsx_pjecalc",
        )

    # 4. Chat Jurídico de Liquidação & Impugnação (1 entrada única por processo)
    st.divider()
    st.subheader("4. Chat Jurídico de Liquidação & Impugnação com a IA")
    st.caption("Converse com a IA sobre os parâmetros da sentença, tetos da coisa julgada, honorários sucumbenciais ou peça a minuta de cumprimento/impugnação.")

    options_map: Dict[str, Optional[Dict[str, Any]]] = {"Sentença / Processo Atual em Análise": None}
    for c in st.session_state.get("cloud_cases", []):
        c_res = c.get("calc_result") or {}
        tot_s = c_res.get("total_geral_fmt", "R$ 0,00")
        proc_s = c.get("process_number") or c.get("case_id")
        label = f"{c.get('client_name')} ({proc_s} • {tot_s})"
        if label not in options_map:
            options_map[label] = c

    col_sel, col_clr = st.columns([4.5, 1])
    selected_label = col_sel.selectbox("Processo / Liquidação selecionada para conversar:", list(options_map.keys()), key="liq_chat_selector")
    selected_case = options_map[selected_label]

    if col_clr.button("Limpar Conversa", use_container_width=True, key="btn_clear_liq_chat"):
        if selected_case:
            selected_case["chat_history"] = []
            upsert_cloud_case(st.session_state["office_email"], selected_case)
        else:
            st.session_state["sentence_chat_current"] = []
        st.rerun()

    active_history = selected_case.get("chat_history", []) if selected_case else st.session_state.get("sentence_chat_current", [])

    for msg in active_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    q_liq = st.chat_input("Ex.: Confira os tetos da sentença, os 15% de honorários ou redija a petição de liquidação...", key="chat_input_liq")
    if q_liq:
        active_history.append({"role": "user", "content": q_liq})
        if selected_case:
            ctx_liq = selected_case
        else:
            ctx_liq = {
                "client_name": (st.session_state.get("sentence_calc_result") or {}).get("client_name", s_ext.get("client_name", "Reclamante")),
                "process_number": (st.session_state.get("sentence_calc_result") or {}).get("process_number", s_ext.get("process_number", "")),
                "calc_result": st.session_state.get("sentence_calc_result"),
                "ai_report": s_ext.get("parecer_sentenca", ""),
            }
        reply_liq = run_legal_chat_turn(q_liq, ctx_liq, active_history, "Liquidação de Sentença & Impugnação PJe-Calc")
        active_history.append({"role": "assistant", "content": reply_liq})

        if selected_case:
            selected_case["chat_history"] = active_history
            upsert_cloud_case(st.session_state["office_email"], selected_case)
        else:
            st.session_state["sentence_chat_current"] = active_history
        st.rerun()


# ==============================================================================
# 12. ABA 3: CARTEIRA DE CLIENTES & RELATÓRIOS SALVOS NA NUVEM
# ==============================================================================
else:
    st.title("Carteira de Clientes & Cálculos Salvos na Nuvem")
    cases_list = st.session_state.get("cloud_cases", [])
    if not cases_list:
        st.info("Nenhum cálculo salvo na nuvem ainda. Calcule um caso na Aba 1 ou Aba 2 para vê-lo aqui.")
    else:
        rows_summary = []
        for c in cases_list:
            cr = c.get("calc_result") or {}
            rows_summary.append(
                {
                    "ID": c.get("case_id"),
                    "Cliente / Reclamante": c.get("client_name"),
                    "Processo": c.get("process_number") or "Pré-Processual",
                    "Modalidade": "Liquidação de Sentença" if c.get("mode") == "liquidacao_sentenca" else "Avaliação Preliminar",
                    "Principal": cr.get("principal_historico_fmt", "-"),
                    "Total Geral": cr.get("total_geral_fmt", "-"),
                }
            )
        st.dataframe(pd.DataFrame(rows_summary), use_container_width=True, hide_index=True)

        chosen_idx = st.selectbox(
            "Selecione um caso para baixar PDF PJe-Calc, .PJC ou Excel:",
            range(len(cases_list)),
            format_func=lambda i: f"{cases_list[i].get('client_name')} — {cases_list[i].get('process_number') or cases_list[i].get('case_id')}",
        )
        chosen_case = cases_list[chosen_idx]
        chosen_calc = chosen_case.get("calc_result")
        if chosen_calc:
            b1, b2, b3 = st.columns(3)
            safe_id = re.sub(r"[^0-9A-Za-z_-]+", "_", str(chosen_case.get("process_number") or chosen_case.get("case_id")))
            b1.download_button("📄 Baixar PDF PJe-Calc", build_pjecalc_pdf_bytes(chosen_calc, st.session_state["office_header"]), f"PJeCalc_{safe_id}.pdf", "application/pdf", use_container_width=True, key="arch_pdf")
            b2.download_button("⚖️ Baixar Arquivo .PJC", build_pjecalc_pjc_bytes(chosen_calc), f"PJeCalc_{safe_id}.pjc", "application/xml", use_container_width=True, key="arch_pjc")
            b3.download_button("📊 Baixar Planilha .XLSX", build_pjecalc_excel_bytes(chosen_calc), f"PJeCalc_{safe_id}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="arch_xlsx")
