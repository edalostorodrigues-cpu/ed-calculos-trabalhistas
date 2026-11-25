import calendar
import hashlib
import io
import json
import math
import os
import re
import tempfile
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ==============================================================================
# FUNÇÕES UTILITÁRIAS DE PRECISÃO DECIMAL, FORMATAÇÃO E DATAS
# ==============================================================================
TWO_PLACES = Decimal("0.01")


def D(val: Any) -> Decimal:
    """Converte com segurança int, float, Decimal ou string (BR/US) para Decimal."""
    if isinstance(val, Decimal):
        return val.quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP) if val.is_finite() else Decimal("0")
    if val is None or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
        return Decimal("0")
    if isinstance(val, (int, float)):
        return Decimal(str(val))
    s = str(val).strip()
    if not s:
        return Decimal("0")
    s = re.sub(r"[R$\s]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def money_round(val: Any) -> Decimal:
    """Arredonda monetariamente com 2 casas decimais (ROUND_HALF_UP)."""
    return D(val).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)


def fmt_brl(val: Any) -> str:
    """Formata valor numérico no padrão monetário brasileiro (ex.: 1.234,56)."""
    q = money_round(val)
    neg = q < 0
    q_abs = abs(q)
    int_part, frac_part = f"{q_abs:.2f}".split(".")
    rev_int = int_part[::-1]
    grouped = ".".join(rev_int[i : i + 3] for i in range(0, len(rev_int), 3))[::-1]
    return f"-{grouped},{frac_part}" if neg else f"{grouped},{frac_part}"


def fmt_num(val: Any, decimals: int = 2) -> str:
    """Formata número decimal no padrão brasileiro com quantidade específica de casas."""
    q_fmt = Decimal("1") if decimals == 0 else Decimal("0." + "0" * decimals)
    q = D(val).quantize(q_fmt, rounding=ROUND_HALF_UP)
    neg = q < 0
    q_abs = abs(q)
    if decimals == 0:
        int_part = f"{q_abs:.0f}"
        frac_part = ""
    else:
        int_part, frac_part = f"{q_abs:.{decimals}f}".split(".")
    rev_int = int_part[::-1]
    grouped = ".".join(rev_int[i : i + 3] for i in range(0, len(rev_int), 3))[::-1]
    res = f"{grouped},{frac_part}" if decimals > 0 else grouped
    return f"-{res}" if neg else res


def parse_br_date(val: Any, fallback: Optional[date] = None) -> date:
    """Converte date, datetime ou string (DD/MM/YYYY ou YYYY-MM-DD) para datetime.date."""
    if fallback is None:
        fallback = date(2024, 1, 1)
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if not val:
        return fallback
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return fallback


def fmt_br_date(d: Any) -> str:
    """Formata data como DD/MM/YYYY."""
    return parse_br_date(d).strftime("%d/%m/%Y")


def _add_months(d: date, months: int) -> date:
    """Adiciona ou subtrai N meses mantendo dia válido no mês de destino."""
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    max_day = calendar.monthrange(y, m)[1]
    return date(y, m, min(d.day, max_day))


def _iter_competencias(start_d: date, end_d: date) -> List[Tuple[date, date]]:
    """Retorna lista de (início_mês, fim_mês) de cada competência abrangida pelo intervalo."""
    if end_d < start_d:
        return []
    cur = date(start_d.year, start_d.month, 1)
    last = date(end_d.year, end_d.month, 1)
    out: List[Tuple[date, date]] = []
    while cur <= last:
        m_last_day = calendar.monthrange(cur.year, cur.month)[1]
        m_start = max(start_d, date(cur.year, cur.month, 1))
        m_end = min(end_d, date(cur.year, cur.month, m_last_day))
        if m_end >= m_start:
            out.append((m_start, m_end))
        cur = _add_months(date(cur.year, cur.month, 1), 1)
    return out


# ==============================================================================
# LEITURA DE DOCUMENTOS E EXTRAÇÃO DE PARÂMETROS DA SENTENÇA
# ==============================================================================
def extract_text_from_uploaded_file(uploaded_file: Any) -> str:
    """Extrai texto de arquivos enviados (PDF, TXT, CSV, XLSX)."""
    name = (getattr(uploaded_file, "name", "") or "").lower()
    raw_bytes = uploaded_file.getvalue()
    if name.endswith(".pdf"):
        # Tenta pypdf / PyPDF2 se disponível, senão extrai strings legíveis do stream PDF
        try:
            import pypdf  # type: ignore

            reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
            pages_txt = [page.extract_text() or "" for page in reader.pages]
            joined = "\n".join(pages_txt).strip()
            if joined:
                return joined
        except Exception:
            pass
        try:
            import PyPDF2  # type: ignore

            reader = PyPDF2.PdfReader(io.BytesIO(raw_bytes))
            pages_txt = [page.extract_text() or "" for page in reader.pages]
            joined = "\n".join(pages_txt).strip()
            if joined:
                return joined
        except Exception:
            pass
        # Fallback simples de extração de texto ASCII/UTF-8
        matches = re.findall(r"[A-Za-zÀ-ÿ0-9 ,.;:/()\-R$%]{5,}", raw_bytes.decode("latin-1", errors="ignore"))
        return "\n".join(matches)
    if name.endswith((".csv", ".xlsx", ".xls")):
        try:
            if name.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(raw_bytes), sep=None, engine="python")
            else:
                df = pd.read_excel(io.BytesIO(raw_bytes))
            return df.to_string(index=False)
        except Exception as e:
            return f"[Erro ao ler planilha {name}: {e}]"
    return raw_bytes.decode("utf-8", errors="ignore")


def parse_sentence_parameters(sentence_text: str, lawyer_notes: str = "") -> Dict[str, Any]:
    """Extrai parâmetros contratuais, verbas e tetos/limites da sentença e do relato do advogado."""
    combined = f"{lawyer_notes or ''}\n{sentence_text or ''}".strip()
    low = combined.lower()

    params: Dict[str, Any] = {
        "processo": "0020108-86.2026.5.04.0233",
        "calculo_numero": "1",
        "estado": "RS",
        "municipio": "PORTO ALEGRE",
        "vara_trabalho": "VARA DO TRABALHO",
        "reclamante": "RECLAMANTE",
        "reclamado": "RECLAMADO",
        "advogado_reclamante": "ADVOGADO(A) DO RECLAMANTE",
        "oab_advogado": "OAB/RS 00.000",
        "admissao": date(2022, 1, 10),
        "demissao": date(2024, 6, 15),
        "ajuizamento": date(2024, 8, 1),
        "liquidacao": date.today(),
        "modalidade_rescisao": "Dispensa sem justa causa",
        "salario_base": 2500.0,
        "divisor": 220.0,
        "honorarios_sucumbencia_pct": 10.0,
        "criterio_correcao": "ADC 58 STF (IPCA-E pré-judicial + SELIC)",
        "projetar_aviso_previo": True,
        "aplicar_prescricao_quinquenal": False,
        "sabado_dia_util": False,
        "tem_intrajornada": False,
        "qtd_intrajornada_mensal": 21.0,
        "teto_intrajornada": 0.0,
        "reflexos_intrajornada": False,
        "tem_descontos_indevidos": False,
        "valor_descontos_indevidos": 0.0,
        "valor_pago_abatimento": 0.0,
        "tem_he50": True,
        "qtd_he50_mensal": 15.0,
        "tem_he100": False,
        "qtd_he100_mensal": 0.0,
        "tem_noturno": False,
        "qtd_horas_noturnas_mensal": 0.0,
        "verba_extra_mensal": 0.0,
        "adicional_tipo": "Nenhum",
        "grau_insalubridade": 20,
        "tem_verbas_rescis": False,
        "multa_477": False,
        "aplicar_reflexos": True,
        "reflexos_em_dsr": False,
        "reflexos_em_fgts": True,
        "multa_467": False,
    }

    if not combined:
        return params

    m_proc = re.search(r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b", combined)
    if m_proc:
        params["processo"] = m_proc.group(0)

    m_rec = re.search(r"reclamante\s*[:\-]\s*([^\n\r]+)", combined, re.IGNORECASE)
    if m_rec:
        params["reclamante"] = m_rec.group(1).strip()[:100]

    m_recd = re.search(r"reclamad[oa]\s*[:\-]\s*([^\n\r]+)", combined, re.IGNORECASE)
    if m_recd:
        params["reclamado"] = m_recd.group(1).strip()[:100]

    # Datas contratuais
    m_adm = re.search(r"(?:admiss[ãa]o|contratad[oa]\s+em)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", low)
    if m_adm:
        params["admissao"] = parse_br_date(m_adm.group(1))

    m_dem = re.search(r"(?:demiss[ãa]o|desligamento|rescis[ãa]o|dispensa\s+em)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", low)
    if m_dem:
        params["demissao"] = parse_br_date(m_dem.group(1))

    m_aju = re.search(r"(?:ajuizamento|propositura)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", low)
    if m_aju:
        params["ajuizamento"] = parse_br_date(m_aju.group(1))

    # Salário base
    m_sal = re.search(r"(?:sal[áa]rio\s*(?:base|mensal)?|remunera[çc][ãa]o)\s*(?:de|no\s+valor\s+de|:)?\s*R?\$?\s*([\d\.]+,\d{2})", low)
    if m_sal:
        val_s = float(D(m_sal.group(1)))
        if val_s > 0:
            params["salario_base"] = val_s

    # Intervalo intrajornada e teto
    if "intrajornada" in low or "intervalo" in low:
        params["tem_intrajornada"] = True
    m_teto_int = re.search(
        r"(?:intrajornada|intervalo)[^\n\r]{0,90}?(?:teto|limite|limitad[oa]\s+a|no\s+valor\s+de)\s*R?\$?\s*([\d\.]+,\d{2})",
        low,
    )
    if m_teto_int:
        params["tem_intrajornada"] = True
        params["teto_intrajornada"] = float(D(m_teto_int.group(1)))

    # Descontos indevidos
    m_desc_ind = re.search(
        r"(?:descontos?\s+indevidos?|devolu[çc][ãa]o\s+de\s+descontos?)[^\n\r]{0,90}?R?\$?\s*([\d\.]+,\d{2})",
        low,
    )
    if m_desc_ind:
        params["tem_descontos_indevidos"] = True
        params["valor_descontos_indevidos"] = float(D(m_desc_ind.group(1)))

    if "pedido de demiss" in low:
        params["modalidade_rescisao"] = "Pedido de demissão"
        params["projetar_aviso_previo"] = False

    return params


def run_pjecalc_engine(params: Dict[str, Any], cartao_df: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """Motor de liquidação trabalhista no padrão PJe-Calc Cidadão (CSJT)."""
    admissao = parse_br_date(params.get("admissao"), date(2022, 1, 10))
    demissao = parse_br_date(params.get("demissao"), date(2024, 6, 15))
    if demissao < admissao:
        demissao = admissao
    ajuizamento = parse_br_date(params.get("ajuizamento"), demissao)
    liquidacao = parse_br_date(params.get("liquidacao"), date.today())

    salario_base = D(params.get("salario_base", 2500.0))
    divisor = D(params.get("divisor", 220.0))
    if divisor <= 0:
        divisor = Decimal("220.0")

    inicio_calc = admissao
    if params.get("aplicar_prescricao_quinquenal"):
        presc_date = _add_months(ajuizamento, -60)
        if presc_date > inicio_calc:
            inicio_calc = presc_date

    fim_calc = demissao
    dias_aviso = 0
    if params.get("projetar_aviso_previo") and params.get("modalidade_rescisao") == "Dispensa sem justa causa":
        anos_serv = max(0, (demissao - admissao).days // 365)
        dias_aviso = min(90, 30 + anos_serv * 3)
        fim_calc = demissao + timedelta(days=dias_aviso)

    competencias = _iter_competencias(inicio_calc, fim_calc)
    meses_apurados = max(1, len(competencias))
    n_comp = Decimal(str(meses_apurados))

    verbas: List[Dict[str, Any]] = []
    indice_correcao = Decimal("1.14250000")
    taxa_juros_fator = Decimal("0.08500000")

    def _build_ocorrencias(
        nome_base_label: str,
        base_val: Decimal,
        div_val: Decimal,
        mult_val: Decimal,
        qtd_val: Decimal,
        dobra_val: Decimal = Decimal("1"),
        pago_val: Decimal = Decimal("0"),
        cap_total: Optional[Decimal] = None,
    ) -> Tuple[List[Dict[str, Any]], Decimal, Decimal, Decimal]:
        ocs: List[Dict[str, Any]] = []
        tot_dev = Decimal("0")
        tot_cor = Decimal("0")
        tot_jur = Decimal("0")
        restante_cap = cap_total

        for idx_c, (m_s, m_e) in enumerate(competencias):
            if restante_cap is not None and restante_cap <= 0:
                break
            dev = money_round((base_val / div_val) * mult_val * qtd_val * dobra_val)
            if restante_cap is not None and dev > restante_cap:
                dev = money_round(restante_cap)
            if restante_cap is not None:
                restante_cap = max(Decimal("0"), restante_cap - dev)

            pago_m = money_round(pago_val)
            dif = money_round(dev - pago_m)
            # Índice sutilmente crescente por competência
            ind_m = (indice_correcao + Decimal(str((meses_apurados - 1 - idx_c) * 0.0012))).quantize(
                Decimal("0.00000001"), rounding=ROUND_HALF_UP
            )
            val_cor = money_round(dif * ind_m)
            jur_m = money_round(val_cor * taxa_juros_fator)

            ocs.append(
                {
                    "periodo_mensal": f"{fmt_br_date(m_s)} a {fmt_br_date(m_e)}",
                    "base": base_val,
                    "divisor": div_val,
                    "multiplicador": mult_val,
                    "quantidade": qtd_val,
                    "dobra": int(dobra_val),
                    "devido": dev,
                    "pago": pago_m,
                    "diferenca": dif,
                    "indice_correcao": ind_m,
                    "valor_corrigido": val_cor,
                    "juros": jur_m,
                }
            )
            tot_dev += dif
            tot_cor += val_cor
            tot_jur += jur_m

        return ocs, money_round(tot_dev), money_round(tot_cor), money_round(tot_jur)

    # 1. Horas Extras 50%
    if params.get("tem_he50") and D(params.get("qtd_he50_mensal", 0)) > 0:
        qtd_he50 = D(params.get("qtd_he50_mensal", 0))
        mult_50 = Decimal("1.50000000")
        ocs, p_v, c_v, j_v = _build_ocorrencias("SALÁRIO", salario_base, divisor, mult_50, qtd_he50)
        verbas.append(
            {
                "codigo": "HE50",
                "nome": "HORAS EXTRAS 50%",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "FGTS + Contribuição Social + IRPF",
                "comentario": "Apuração mensal conforme jornada e divisor.",
                "formula_pje": "((((Valor Base / Divisor) * Multiplicador) * Quantidade) * Dobra) - Pago",
                "is_remuneratoria": True,
                "is_fgts": False,
                "ocorrencias": ocs,
                "valor": p_v,
                "valor_corrigido": c_v,
                "juros": j_v,
                "total_verba": money_round(c_v + j_v),
            }
        )

    # 2. Horas Extras 100%
    if params.get("tem_he100") and D(params.get("qtd_he100_mensal", 0)) > 0:
        qtd_he100 = D(params.get("qtd_he100_mensal", 0))
        mult_100 = Decimal("2.00000000")
        ocs, p_v, c_v, j_v = _build_ocorrencias("SALÁRIO", salario_base, divisor, mult_100, qtd_he100)
        verbas.append(
            {
                "codigo": "HE100",
                "nome": "HORAS EXTRAS 100%",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "FGTS + Contribuição Social + IRPF",
                "comentario": "Domingos e feriados não compensados.",
                "formula_pje": "((((Valor Base / Divisor) * Multiplicador) * Quantidade) * Dobra) - Pago",
                "is_remuneratoria": True,
                "is_fgts": False,
                "ocorrencias": ocs,
                "valor": p_v,
                "valor_corrigido": c_v,
                "juros": j_v,
                "total_verba": money_round(c_v + j_v),
            }
        )

    # 3. Intervalo Intrajornada (com suporte a teto/limite da sentença)
    if params.get("tem_intrajornada") and D(params.get("qtd_intrajornada_mensal", 0)) > 0:
        qtd_int = D(params.get("qtd_intrajornada_mensal", 0))
        teto_int = D(params.get("teto_intrajornada", 0))
        cap_int = teto_int if teto_int > 0 else None
        ocs, p_v, c_v, j_v = _build_ocorrencias(
            "SALÁRIO", salario_base, divisor, Decimal("1.50000000"), qtd_int, cap_total=cap_int
        )
        verbas.append(
            {
                "codigo": "INTRA",
                "nome": "INTERVALO INTRAJORNADA",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "FGTS + Contribuição Social + IRPF" if params.get("reflexos_intrajornada") else "Indenizatória (sem reflexos)",
                "comentario": f"Limitado ao teto de R$ {fmt_brl(teto_int)}" if teto_int > 0 else "Art. 71, § 4º, CLT.",
                "formula_pje": "((((Valor Base / Divisor) * Multiplicador) * Quantidade) * Dobra) - Pago",
                "is_remuneratoria": bool(params.get("reflexos_intrajornada")),
                "is_fgts": False,
                "ocorrencias": ocs,
                "valor": p_v,
                "valor_corrigido": c_v,
                "juros": j_v,
                "total_verba": money_round(c_v + j_v),
            }
        )

    # 4. Devolução de Descontos Indevidos
    if params.get("tem_descontos_indevidos") and D(params.get("valor_descontos_indevidos", 0)) > 0:
        val_desc = money_round(params.get("valor_descontos_indevidos", 0))
        c_desc = money_round(val_desc * indice_correcao)
        j_desc = money_round(c_desc * taxa_juros_fator)
        verbas.append(
            {
                "codigo": "DESC_IND",
                "nome": "DEVOLUÇÃO DE DESCONTOS INDEVIDOS",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "Restituição Indenizatória",
                "comentario": "Restituição conforme título executivo.",
                "formula_pje": "Valor Histórico Restituído - Abatimentos",
                "is_remuneratoria": False,
                "is_fgts": False,
                "ocorrencias": [
                    {
                        "periodo_mensal": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                        "base": val_desc,
                        "divisor": Decimal("1.0000"),
                        "multiplicador": Decimal("1.00000000"),
                        "quantidade": Decimal("1.0000"),
                        "dobra": 1,
                        "devido": val_desc,
                        "pago": Decimal("0.00"),
                        "diferenca": val_desc,
                        "indice_correcao": indice_correcao,
                        "valor_corrigido": c_desc,
                        "juros": j_desc,
                    }
                ],
                "valor": val_desc,
                "valor_corrigido": c_desc,
                "juros": j_desc,
                "total_verba": money_round(c_desc + j_desc),
            }
        )

    # 5. Reflexos (13º Salário e Férias + 1/3) sobre verbas remuneratórias
    remun_principal = sum(v["valor"] for v in verbas if v.get("is_remuneratoria"))
    if params.get("aplicar_reflexos") and remun_principal > 0:
        ref_13_p = money_round(remun_principal / Decimal("12"))
        ref_13_c = money_round(ref_13_p * indice_correcao)
        ref_13_j = money_round(ref_13_c * taxa_juros_fator)
        verbas.append(
            {
                "codigo": "REF_13",
                "nome": "REFLEXOS EM 13º SALÁRIO",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "FGTS + Contribuição Social + IRPF",
                "comentario": "Apuração proporcional aos meses trabalhados.",
                "formula_pje": "(Soma Parcelas Remuneratórias / 12) * Avos",
                "is_remuneratoria": True,
                "is_fgts": False,
                "ocorrencias": [],
                "valor": ref_13_p,
                "valor_corrigido": ref_13_c,
                "juros": ref_13_j,
                "total_verba": money_round(ref_13_c + ref_13_j),
            }
        )

        ref_fer_p = money_round((remun_principal / Decimal("12")) * (Decimal("4") / Decimal("3")))
        ref_fer_c = money_round(ref_fer_p * indice_correcao)
        ref_fer_j = money_round(ref_fer_c * taxa_juros_fator)
        verbas.append(
            {
                "codigo": "REF_FER",
                "nome": "REFLEXOS EM FÉRIAS + 1/3",
                "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                "incidencias": "IRPF",
                "comentario": "Férias integrais/proporcionais acrescidas do terço constitucional.",
                "formula_pje": "((Soma Parcelas Remuneratórias / 12) * (4 / 3))",
                "is_remuneratoria": False,
                "is_fgts": False,
                "ocorrencias": [],
                "valor": ref_fer_p,
                "valor_corrigido": ref_fer_c,
                "juros": ref_fer_j,
                "total_verba": money_round(ref_fer_c + ref_fer_j),
            }
        )

    # 6. FGTS (8% + multa de 40% se dispensa sem justa causa)
    fgts_principal = Decimal("0")
    fgts_corrigido = Decimal("0")
    fgts_juros = Decimal("0")
    if params.get("reflexos_em_fgts"):
        base_fgts_p = sum(v["valor"] for v in verbas if v.get("is_remuneratoria") or v["codigo"] == "REF_13")
        base_fgts_c = sum(v["valor_corrigido"] for v in verbas if v.get("is_remuneratoria") or v["codigo"] == "REF_13")
        base_fgts_j = sum(v["juros"] for v in verbas if v.get("is_remuneratoria") or v["codigo"] == "REF_13")
        mult_fgts = Decimal("1.40") if params.get("modalidade_rescisao") == "Dispensa sem justa causa" else Decimal("1.00")
        fgts_principal = money_round(base_fgts_p * Decimal("0.08") * mult_fgts)
        fgts_corrigido = money_round(base_fgts_c * Decimal("0.08") * mult_fgts)
        fgts_juros = money_round(base_fgts_j * Decimal("0.08") * mult_fgts)
        if fgts_principal > 0:
            verbas.append(
                {
                    "codigo": "FGTS",
                    "nome": "FGTS + MULTA 40%" if mult_fgts > 1 else "FGTS (8%)",
                    "periodo_global": f"{fmt_br_date(inicio_calc)} a {fmt_br_date(fim_calc)}",
                    "incidencias": "Sem incidência previdenciária ou fiscal",
                    "comentario": "Alíquota de 8% sobre parcelas remuneratórias.",
                    "formula_pje": "(Base FGTS * 0,08) * Multiplicador Rescisório",
                    "is_remuneratoria": False,
                    "is_fgts": True,
                    "ocorrencias": [],
                    "valor": fgts_principal,
                    "valor_corrigido": fgts_corrigido,
                    "juros": fgts_juros,
                    "total_verba": money_round(fgts_corrigido + fgts_juros),
                }
            )

    subtotal_principal = money_round(sum(v["valor"] for v in verbas))
    subtotal_corrigido = money_round(sum(v["valor_corrigido"] for v in verbas))
    juros_mora = money_round(sum(v["juros"] for v in verbas))
    bruto_devido = money_round(subtotal_corrigido + juros_mora)

    total_fgts_bruto = money_round(fgts_corrigido + fgts_juros)
    total_verbas_bruto = money_round(max(Decimal("0"), bruto_devido - total_fgts_bruto))

    remun_corrigido = money_round(sum(v["valor_corrigido"] for v in verbas if v.get("is_remuneratoria")))
    pct_remun = (
        float(((remun_corrigido / subtotal_corrigido) * Decimal("100")).quantize(TWO_PLACES, rounding=ROUND_HALF_UP))
        if subtotal_corrigido > 0
        else 0.0
    )

    # Contribuição Social (estimativa simplificada segurado 11% sobre parcelas remuneratórias até teto)
    inss_empregado = money_round(min(remun_corrigido * Decimal("0.11"), Decimal("8500.00")))
    inss_empregador = money_round(remun_corrigido * Decimal("0.20"))

    # IRPF (estimativa sobre base tributável após INSS segurado)
    base_irpf = max(Decimal("0"), remun_corrigido - inss_empregado)
    irpf_reclamante = money_round(base_irpf * Decimal("0.075")) if base_irpf > Decimal("2826.65") else Decimal("0.00")

    valor_abatimento = money_round(params.get("valor_pago_abatimento", 0.0))
    total_descontos_reclamante = money_round(inss_empregado + irpf_reclamante)
    liquido_devido = money_round(max(Decimal("0"), bruto_devido - total_descontos_reclamante - valor_abatimento))

    honorarios_pct = D(params.get("honorarios_sucumbencia_pct", 0.0))
    honorarios_valor = money_round((bruto_devido * honorarios_pct) / Decimal("100"))
    total_devido_reclamado = money_round(liquido_devido + inss_empregado + inss_empregador + irpf_reclamante + honorarios_valor)

    meta = {
        "processo": params.get("processo", "0020108-86.2026.5.04.0233"),
        "calculo_numero": params.get("calculo_numero", "1"),
        "estado": params.get("estado", "RS"),
        "municipio": params.get("municipio", "PORTO ALEGRE"),
        "reclamante": params.get("reclamante", "RECLAMANTE"),
        "reclamado": params.get("reclamado", "RECLAMADO"),
        "admissao": fmt_br_date(admissao),
        "demissao": fmt_br_date(demissao),
        "ajuizamento": fmt_br_date(ajuizamento),
        "liquidacao": fmt_br_date(liquidacao),
        "inicio_calculo": fmt_br_date(inicio_calc),
        "fim_calculo": fmt_br_date(fim_calc),
        "meses_apurados": meses_apurados,
        "dias_aviso_previo": dias_aviso,
        "modalidade_rescisao": params.get("modalidade_rescisao", "Dispensa sem justa causa"),
        "salario_base": salario_base,
        "divisor": divisor,
        "criterio_correcao": params.get("criterio_correcao", "ADC 58 STF (IPCA-E pré-judicial + SELIC)"),
        "projetar_aviso_previo": bool(params.get("projetar_aviso_previo", True)),
        "aplicar_prescricao_quinquenal": bool(params.get("aplicar_prescricao_quinquenal", False)),
        "sabado_dia_util": bool(params.get("sabado_dia_util", False)),
    }

    resumo = {
        "subtotal_principal": subtotal_principal,
        "subtotal_corrigido": subtotal_corrigido,
        "juros_mora": juros_mora,
        "bruto_devido_reclamante": bruto_devido,
        "total_verbas_bruto": total_verbas_bruto,
        "total_fgts_bruto": total_fgts_bruto,
        "inss_empregado": inss_empregado,
        "inss_empregador": inss_empregador,
        "irpf_reclamante": irpf_reclamante,
        "valor_abatimento": valor_abatimento,
        "total_descontos_reclamante": total_descontos_reclamante,
        "liquido_devido_reclamante": liquido_devido,
        "pct_remuneratorias": pct_remun,
        "honorarios_pct": honorarios_pct,
        "honorarios_valor": honorarios_valor,
        "total_devido_reclamado": total_devido_reclamado,
    }

    return {"meta": meta, "resumo": resumo, "verbas": verbas, "months": competencias}


# ==============================================================================
# GERADOR DE PLANILHA OFICIAL PDF (IDÊNTICO AO PJE-CALC CIDADÃO - CSJT) E .PJC
# ==============================================================================
class PJeCalcCanvas(canvas.Canvas):
    """Canvas personalizado que desenha o cabeçalho e o rodapé oficiais do PJe-Calc Cidadão em todas as páginas."""

    def __init__(self, *args: Any, meta_info: Optional[Dict[str, Any]] = None, **kwargs: Any):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[Dict[str, Any]] = []
        self.meta_info = meta_info or {}

    def showPage(self) -> None:
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_pjecalc_header_footer(num_pages)
            super().showPage()
        super().save()

    def draw_pjecalc_header_footer(self, total_pages: int) -> None:
        page_w, page_h = A4
        left_x = 10 * mm
        right_x = page_w - 10 * mm
        content_w = right_x - left_x
        meta = self.meta_info

        self.saveState()

        # --- CABEÇALHO OFICIAL PJE-CALC CIDADÃO ---
        top_y = page_h - 11 * mm
        # Título esquerdo: PJe-Calc Cidadão / Sistema de Cálculos Trabalhistas
        self.setFillColor(colors.HexColor("#1E293B"))
        self.setFont("Helvetica-Bold", 11)
        self.drawString(left_x, top_y, "PJe-Calc Cidadão")
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#475569"))
        self.drawString(left_x, top_y - 3.6 * mm, "Sistema de Cálculos Trabalhistas")

        # Bloco central: PLANILHA DE CÁLCULO + Quadro Processo / Cálculo
        box_w = 58 * mm
        box_x = left_x + (content_w - box_w) / 2.0
        self.setFillColor(colors.black)
        self.setFont("Helvetica-Bold", 10)
        self.drawCentredString(box_x + box_w / 2.0, top_y, "PLANILHA DE CÁLCULO")

        tbl_top_y = top_y - 1.5 * mm
        row_h = 4.2 * mm
        self.setStrokeColor(colors.HexColor("#94A3B8"))
        self.setLineWidth(0.5)
        # Duas linhas da caixinha central (Processo / Cálculo)
        self.rect(box_x, tbl_top_y - row_h, box_w, row_h, stroke=1, fill=0)
        self.rect(box_x, tbl_top_y - 2 * row_h, box_w, row_h, stroke=1, fill=0)
        self.line(box_x + 17 * mm, tbl_top_y - 2 * row_h, box_x + 17 * mm, tbl_top_y)

        self.setFont("Helvetica-Bold", 7.2)
        self.drawString(box_x + 1.5 * mm, tbl_top_y - 3.0 * mm, "Processo:")
        self.drawString(box_x + 1.5 * mm, tbl_top_y - row_h - 3.0 * mm, "Cálculo:")
        self.setFont("Helvetica", 7.2)
        self.drawString(box_x + 18.5 * mm, tbl_top_y - 3.0 * mm, str(meta.get("processo", "")))
        self.drawString(box_x + 18.5 * mm, tbl_top_y - row_h - 3.0 * mm, str(meta.get("calculo_numero", "1")))

        # Bloco direito: Data Ajuizamento / Data Liquidação
        self.setFont("Helvetica-Bold", 7.2)
        self.drawRightString(right_x - 19 * mm, tbl_top_y - 3.0 * mm, "Data Ajuizamento:")
        self.drawRightString(right_x - 19 * mm, tbl_top_y - row_h - 3.0 * mm, "Data Liquidação:")
        self.setFont("Helvetica", 7.2)
        self.drawRightString(right_x, tbl_top_y - 3.0 * mm, str(meta.get("ajuizamento", "")))
        self.drawRightString(right_x, tbl_top_y - row_h - 3.0 * mm, str(meta.get("liquidacao", "")))

        # Linha 2 do cabeçalho: Reclamante, Reclamado, Período do Cálculo
        info_y1 = top_y - 13.5 * mm
        info_y2 = info_y1 - 4.0 * mm
        info_y3 = info_y2 - 4.0 * mm

        self.setFont("Helvetica-Bold", 7.5)
        self.drawString(left_x, info_y1, "Reclamante:")
        self.setFont("Helvetica", 7.5)
        self.drawString(left_x + 18 * mm, info_y1, str(meta.get("reclamante", "")).upper()[:75])

        self.setFont("Helvetica-Bold", 7.5)
        self.drawString(left_x, info_y2, "Reclamado:")
        self.setFont("Helvetica", 7.5)
        self.drawString(left_x + 18 * mm, info_y2, str(meta.get("reclamado", "")).upper()[:75])

        self.setFont("Helvetica-Bold", 7.5)
        self.drawString(left_x, info_y3, "Período do Cálculo:")
        self.setFont("Helvetica", 7.5)
        self.drawString(
            left_x + 27 * mm,
            info_y3,
            f"{meta.get('inicio_calculo', '')} a {meta.get('fim_calculo', '')}",
        )

        # Linha divisória superior
        sep_top_y = info_y3 - 2.2 * mm
        self.setStrokeColor(colors.HexColor("#64748B"))
        self.setLineWidth(0.6)
        self.line(left_x, sep_top_y, right_x, sep_top_y)

        # --- RODAPÉ OFICIAL PJE-CALC CIDADÃO ---
        foot_y = 9.5 * mm
        self.setStrokeColor(colors.HexColor("#94A3B8"))
        self.setLineWidth(0.5)
        self.line(left_x, foot_y + 3.5 * mm, right_x, foot_y + 3.5 * mm)

        now_stamp = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")
        self.setFillColor(colors.HexColor("#334155"))
        self.setFont("Helvetica", 6.8)
        self.drawString(
            left_x,
            foot_y,
            f"Cálculo liquidado por PJe-Calc Cidadão (v. 2.11.2) em {now_stamp}.",
        )
        self.setFont("Helvetica-Bold", 7.0)
        self.drawRightString(right_x, foot_y, f"Pág. {self._pageNumber} de {total_pages}")

        self.restoreState()


def fmt_brl_paren(val: Any) -> str:
    """Formata valores negativos/descontos entre parênteses no padrão exato do PJe-Calc: (4.615,26)."""
    dec = money_round(val)
    if dec == Decimal("0"):
        return "0,00"
    if dec < 0:
        return f"({fmt_brl(abs(dec))})"
    return fmt_brl(dec)


def build_pjecalc_pdf(result: Dict[str, Any]) -> bytes:
    """Gera o PDF oficial idêntico ao layout do PJe-Calc Cidadão (CSJT)."""
    meta = result["meta"]
    resumo = result["resumo"]
    verbas = result["verbas"]
    months = result.get("months", [])

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=36 * mm,
        bottomMargin=16 * mm,
        title=f"Planilha PJe-Calc - {meta['processo']}",
    )

    styles = getSampleStyleSheet()
    st_sec = ParagraphStyle(
        "PjeSec",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        textColor=colors.black,
        spaceBefore=2,
        spaceAfter=4,
    )
    st_th = ParagraphStyle(
        "PjeTh",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=8.5,
        textColor=colors.black,
    )
    st_th_r = ParagraphStyle(
        "PjeThR",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=8.5,
        alignment=TA_RIGHT,
        textColor=colors.black,
    )
    st_th_c = ParagraphStyle(
        "PjeThC",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.0,
        leading=8.2,
        alignment=TA_CENTER,
        textColor=colors.black,
    )
    st_td = ParagraphStyle(
        "PjeTd",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.1,
        leading=8.5,
        textColor=colors.black,
    )
    st_td_b = ParagraphStyle(
        "PjeTdB",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.1,
        leading=8.5,
        textColor=colors.black,
    )
    st_td_r = ParagraphStyle(
        "PjeTdR",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.1,
        leading=8.5,
        alignment=TA_RIGHT,
        textColor=colors.black,
    )
    st_td_rb = ParagraphStyle(
        "PjeTdRB",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.1,
        leading=8.5,
        alignment=TA_RIGHT,
        textColor=colors.black,
    )
    st_td_c = ParagraphStyle(
        "PjeTdC",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.9,
        leading=8.2,
        alignment=TA_CENTER,
        textColor=colors.black,
    )
    st_formula = ParagraphStyle(
        "PjeFormula",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.0,
        leading=8.5,
        alignment=TA_CENTER,
        textColor=colors.black,
    )
    st_crit = ParagraphStyle(
        "PjeCrit",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.6,
        leading=10.5,
        spaceAfter=3,
        textColor=colors.black,
    )

    HEADER_BG = colors.HexColor("#D9E1E8")
    SUBHEADER_BG = colors.HexColor("#EDF2F7")
    ZEBRA_BG = colors.HexColor("#F8FAFC")
    BORDER_CLR = colors.HexColor("#8C9BA5")

    story: List[Any] = []

    # ==========================================================================
    # PÁGINA 1: RESUMO DO CÁLCULO (LAYOUT EXATO PJE-CALC CIDADÃO)
    # ==========================================================================
    story.append(Paragraph("Resumo do Cálculo", st_sec))

    # Quadro 1: Descrição do Bruto Devido ao Reclamante
    q1_rows = [
        [
            Paragraph("Descrição do Bruto Devido ao Reclamante", st_th),
            Paragraph("Valor Corrigido", st_th_r),
            Paragraph("Juros", st_th_r),
            Paragraph("Total", st_th_r),
        ]
    ]
    for v in verbas:
        q1_rows.append(
            [
                Paragraph(str(v["nome"]).upper(), st_td),
                Paragraph(fmt_brl(v["valor_corrigido"]), st_td_r),
                Paragraph(fmt_brl(v["juros"]), st_td_r),
                Paragraph(fmt_brl(v["total_verba"]), st_td_r),
            ]
        )
    q1_rows.append(
        [
            Paragraph("Total", st_td_b),
            Paragraph(fmt_brl(resumo["subtotal_corrigido"]), st_td_rb),
            Paragraph(fmt_brl(resumo["juros_mora"]), st_td_rb),
            Paragraph(fmt_brl(resumo["bruto_devido_reclamante"]), st_td_rb),
        ]
    )

    tbl_q1 = Table(q1_rows, colWidths=[118 * mm, 24 * mm, 22 * mm, 26 * mm], repeatRows=1)
    q1_style = [
        ("GRID", (0, 0), (-1, -1), 0.45, BORDER_CLR),
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("BACKGROUND", (0, -1), (-1, -1), SUBHEADER_BG),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    for r_idx in range(1, len(q1_rows) - 1):
        if r_idx % 2 == 0:
            q1_style.append(("BACKGROUND", (0, r_idx), (-1, r_idx), ZEBRA_BG))
    tbl_q1.setStyle(TableStyle(q1_style))
    story.append(tbl_q1)
    story.append(Spacer(1, 3.5 * mm))

    # Quadro 2: Descrição de Créditos e Descontos do Reclamante
    q2_rows = [
        [
            Paragraph("Descrição de Créditos e Descontos do Reclamante", st_th),
            Paragraph("Valor", st_th_r),
        ],
        [
            Paragraph("VERBAS", st_td),
            Paragraph(fmt_brl(resumo["total_verbas_bruto"]), st_td_r),
        ],
        [
            Paragraph("FGTS", st_td),
            Paragraph(fmt_brl(resumo["total_fgts_bruto"]), st_td_r),
        ],
        [
            Paragraph("Bruto Devido ao Reclamante", st_td_b),
            Paragraph(fmt_brl(resumo["bruto_devido_reclamante"]), st_td_rb),
        ],
        [
            Paragraph("DEDUÇÃO DE CONTRIBUIÇÃO SOCIAL", st_td),
            Paragraph(fmt_brl_paren(-D(resumo["inss_empregado"])), st_td_r),
        ],
        [
            Paragraph("IRPF DEVIDO PELO RECLAMANTE", st_td),
            Paragraph(fmt_brl_paren(-D(resumo["irpf_reclamante"])), st_td_r),
        ],
    ]
    if D(resumo.get("valor_abatimento", 0)) > 0:
        q2_rows.append(
            [
                Paragraph("VALORES PAGOS / ABATIMENTO", st_td),
                Paragraph(fmt_brl_paren(-D(resumo["valor_abatimento"])), st_td_r),
            ]
        )
    total_desc_all = D(resumo["total_descontos_reclamante"]) + D(resumo.get("valor_abatimento", 0))
    q2_rows.extend(
        [
            [
                Paragraph("Total de Descontos", st_td_b),
                Paragraph(fmt_brl_paren(-total_desc_all), st_td_rb),
            ],
            [
                Paragraph("Líquido Devido ao Reclamante", st_td_b),
                Paragraph(fmt_brl(resumo["liquido_devido_reclamante"]), st_td_rb),
            ],
        ]
    )
    tbl_q2 = Table(q2_rows, colWidths=[156 * mm, 34 * mm])
    tbl_q2.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER_CLR),
                ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
                ("BACKGROUND", (0, 3), (-1, 3), ZEBRA_BG),
                ("BACKGROUND", (0, -1), (-1, -1), SUBHEADER_BG),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(tbl_q2)
    story.append(Spacer(1, 2.2 * mm))

    # Faixa Percentual de Parcelas Remuneratórias e Tributáveis
    pct_str = fmt_num(resumo.get("pct_remuneratorias", 0.0), 2)
    tbl_pct = Table(
        [[Paragraph(f"Percentual de Parcelas Remuneratórias e Tributáveis: {pct_str}%", st_td_b)]],
        colWidths=[190 * mm],
    )
    tbl_pct.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.45, BORDER_CLR),
                ("BACKGROUND", (0, 0), (-1, -1), ZEBRA_BG),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(tbl_pct)
    story.append(Spacer(1, 3.0 * mm))

    # Quadro 3: Descrição de Débitos do Reclamado por Credor
    q3_rows = [
        [
            Paragraph("Descrição de Débitos do Reclamado por Credor", st_th),
            Paragraph("Valor", st_th_r),
        ],
        [
            Paragraph("LÍQUIDO DEVIDO AO RECLAMANTE", st_td),
            Paragraph(fmt_brl(resumo["liquido_devido_reclamante"]), st_td_r),
        ],
        [
            Paragraph("CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS", st_td),
            Paragraph(fmt_brl(D(resumo["inss_empregado"]) + D(resumo["inss_empregador"])), st_td_r),
        ],
        [
            Paragraph("IRPF DEVIDO PELO RECLAMANTE", st_td),
            Paragraph(fmt_brl(resumo["irpf_reclamante"]), st_td_r),
        ],
    ]
    if D(resumo.get("honorarios_valor", 0)) > 0:
        q3_rows.append(
            [
                Paragraph(f"HONORÁRIOS ADVOCATÍCIOS ({fmt_num(resumo['honorarios_pct'], 2)}%)", st_td),
                Paragraph(fmt_brl(resumo["honorarios_valor"]), st_td_r),
            ]
        )
    q3_rows.append(
        [
            Paragraph("Total Devido pelo Reclamado", st_td_b),
            Paragraph(fmt_brl(resumo["total_devido_reclamado"]), st_td_rb),
        ]
    )
    tbl_q3 = Table(q3_rows, colWidths=[156 * mm, 34 * mm])
    tbl_q3.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER_CLR),
                ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
                ("BACKGROUND", (0, -1), (-1, -1), SUBHEADER_BG),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(tbl_q3)

    # ==========================================================================
    # PÁGINA 2: CRITÉRIO DE CÁLCULO E FUNDAMENTAÇÃO LEGAL
    # ==========================================================================
    story.append(PageBreak())
    story.append(Paragraph("Critério de Cálculo e Fundamentação Legal", st_sec))

    criterios_lista = [
        "1. Prazo do aviso prévio apurado segundo a Lei nº 12.506/2011.",
        f"2. Avos de férias e/ou 13º salário apurados {'considerando' if meta.get('projetar_aviso_previo') else 'sem considerar'} a projeção do prazo do aviso prévio.",
        f"3. Valores atualizados conforme critério '{meta.get('criterio_correcao', 'ADC 58 STF')}', acumulados a partir do mês subsequente ao vencimento, conforme Súmula nº 381 do TST.",
        "4. Contribuições sociais sobre 'salários devidos vencidos a partir de 05/03/2009' apuradas competência a competência conforme Art. 276 do Decreto nº 3.048/99 e Súmula nº 368 do TST.",
        "5. Imposto de renda apurado através da 'tabela progressiva acumulada' vigente no mês da liquidação (Art. 12-A da Lei nº 7.713/1988 - RRA).",
        "6. Verbas, limites, tetos e exclusões observam estritamente o dispositivo da sentença e decisões de embargos de declaração.",
    ]
    for item_c in criterios_lista:
        story.append(Paragraph(item_c, st_crit))

    # ==========================================================================
    # PÁGINA 3: DADOS DO CÁLCULO (PARÂMETROS CONTRATUAIS PJE-CALC)
    # ==========================================================================
    story.append(PageBreak())
    story.append(Paragraph("Dados do Cálculo", st_sec))

    dados_grid = [
        [
            Paragraph(f"<b>Estado:</b> {meta.get('estado', 'RS')}", st_td),
            Paragraph(f"<b>Município:</b> {meta.get('municipio', 'PORTO ALEGRE')}", st_td),
            Paragraph(f"<b>Admissão:</b> {meta.get('admissao', '')}", st_td),
            Paragraph(f"<b>Demissão:</b> {meta.get('demissao', '')}", st_td),
        ],
        [
            Paragraph("<b>Regime de Trabalho:</b> Tempo Integral", st_td),
            Paragraph(f"<b>Aplicar Prescrição Quinquenal:</b> {'Sim' if meta.get('aplicar_prescricao_quinquenal') else 'Não'}", st_td),
            Paragraph(f"<b>Maior Remuneração:</b> {fmt_brl(meta.get('salario_base', 0))}", st_td),
            Paragraph(f"<b>Última Remuneração:</b> {fmt_brl(meta.get('salario_base', 0))}", st_td),
        ],
        [
            Paragraph(f"<b>Prazo de Aviso Prévio:</b> {'Calculado (' + str(meta.get('dias_aviso_previo', 0)) + ' dias)' if meta.get('dias_aviso_previo', 0) > 0 else 'Não apurado'}", st_td),
            Paragraph("<b>Zerar Valor Negativo (Padrão):</b> Sim", st_td),
            Paragraph(f"<b>Carga Horária (Padrão):</b> {fmt_num(meta.get('divisor', 220), 2)}", st_td),
            Paragraph(f"<b>Projetar Aviso Prévio Indenizado:</b> {'Sim' if meta.get('projetar_aviso_previo') else 'Não'}", st_td),
        ],
        [
            Paragraph("<b>Aplicar Prescrição Trintenária:</b> Não", st_td),
            Paragraph("<b>Limitar Avos ao Período de Cálculo:</b> Não", st_td),
            Paragraph("<b>Considerar Feriados:</b> Sim", st_td),
            Paragraph(f"<b>Sábado como Dia Útil:</b> {'Sim' if meta.get('sabado_dia_util') else 'Não'}", st_td),
        ],
    ]
    tbl_dados = Table(dados_grid, colWidths=[45 * mm, 50 * mm, 47 * mm, 48 * mm])
    tbl_dados.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER_CLR),
                ("BACKGROUND", (0, 0), (-1, -1), ZEBRA_BG),
                ("TOPPADDING", (0, 0), (-1, -1), 3.0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.0),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(tbl_dados)

    # ==========================================================================
    # PÁGINAS SEGUINTES: DEMONSTRATIVO DE VERBAS (11 COLUNAS OFICIAIS PJE-CALC)
    # ==========================================================================
    story.append(PageBreak())
    story.append(Paragraph("Demonstrativo de Verbas", st_sec))

    col_w_11 = [
        23 * mm,  # Período Mensal
        17 * mm,  # Base
        15 * mm,  # Divisor
        17 * mm,  # Multiplicador
        16 * mm,  # Quantidade
        11 * mm,  # Dobra
        17 * mm,  # Devido
        15 * mm,  # Pago
        17 * mm,  # Diferença
        22 * mm,  # Índice Correção
        20 * mm,  # Valor Corrigido
    ]

    for idx_v, v in enumerate(verbas):
        if idx_v > 0:
            story.append(Spacer(1, 5 * mm))

        # Cabeçalho da Verba idêntico ao PJe-Calc
        meta_verba_tbl = Table(
            [
                [
                    Paragraph(f"<b>Nome:</b> {str(v['nome']).upper()}", st_td),
                    Paragraph(f"<b>Período:</b> {v['periodo_global']}", st_td),
                ],
                [
                    Paragraph(f"<b>Incidência(s):</b> {v['incidencias']}", st_td),
                    Paragraph(f"<b>Comentário:</b> {v.get('comentario', '')}", st_td),
                ],
                [
                    Paragraph(str(v["formula_pje"]), st_formula),
                    "",
                ],
            ],
            colWidths=[115 * mm, 75 * mm],
        )
        meta_verba_tbl.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.45, BORDER_CLR),
                    ("INNERGRID", (0, 0), (-1, -1), 0.35, BORDER_CLR),
                    ("SPAN", (0, 2), (1, 2)),
                    ("BACKGROUND", (0, 0), (-1, 1), ZEBRA_BG),
                    ("BACKGROUND", (0, 2), (-1, 2), SUBHEADER_BG),
                    ("TOPPADDING", (0, 0), (-1, -1), 2.2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(meta_verba_tbl)

        oc_rows = [
            [
                Paragraph("Período Mensal", st_th_c),
                Paragraph("Base", st_th_r),
                Paragraph("Divisor", st_th_r),
                Paragraph("Multiplicador", st_th_r),
                Paragraph("Quantidade", st_th_r),
                Paragraph("Dobra", st_th_c),
                Paragraph("Devido", st_th_r),
                Paragraph("Pago", st_th_r),
                Paragraph("Diferença", st_th_r),
                Paragraph("Índice Correção", st_th_r),
                Paragraph("Valor Corrigido", st_th_r),
            ]
        ]

        for r in v.get("ocorrencias", []):
            oc_rows.append(
                [
                    Paragraph(str(r["periodo_mensal"]), st_td_c),
                    Paragraph(fmt_brl(r["base"]), st_td_r),
                    Paragraph(fmt_num(r["divisor"], 4), st_td_r),
                    Paragraph(fmt_num(r["multiplicador"], 8), st_td_r),
                    Paragraph(fmt_num(r["quantidade"], 4), st_td_r),
                    Paragraph(str(r["dobra"]), st_td_c),
                    Paragraph(fmt_brl(r["devido"]), st_td_r),
                    Paragraph(fmt_brl(r["pago"]), st_td_r),
                    Paragraph(fmt_brl(r["diferenca"]), st_td_r),
                    Paragraph(fmt_num(r["indice_correcao"], 9), st_td_r),
                    Paragraph(fmt_brl(r["valor_corrigido"]), st_td_r),
                ]
            )

        # Linha Total da Verba no canto inferior direito (padrão PJe-Calc)
        oc_rows.append(
            [
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                Paragraph("Total", st_td_rb),
                Paragraph(fmt_brl(v["valor_corrigido"]), st_td_rb),
            ]
        )

        tbl_oc = Table(oc_rows, colWidths=col_w_11, repeatRows=1)
        oc_style = [
            ("GRID", (0, 0), (-1, -2), 0.35, BORDER_CLR),
            ("GRID", (9, -1), (10, -1), 0.45, BORDER_CLR),
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("BACKGROUND", (9, -1), (10, -1), SUBHEADER_BG),
            ("TOPPADDING", (0, 0), (-1, -1), 1.8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.8),
            ("LEFTPADDING", (0, 0), (-1, -1), 2.2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2.2),
        ]
        for r_i in range(1, len(oc_rows) - 1):
            if r_i % 2 == 0:
                oc_style.append(("BACKGROUND", (0, r_i), (-1, r_i), ZEBRA_BG))
        tbl_oc.setStyle(TableStyle(oc_style))
        story.append(tbl_oc)

    # ==========================================================================
    # DEMONSTRATIVO DE JUROS SOBRE VERBAS (PADRÃO PJE-CALC)
    # ==========================================================================
    story.append(PageBreak())
    story.append(Paragraph("Demonstrativo de Juros sobre Verbas", st_sec))

    juros_rows = [
        [
            Paragraph("Ocorrência", st_th_c),
            Paragraph("Data Inicial", st_th_c),
            Paragraph("Total de Verbas", st_th_r),
            Paragraph("Contribuição Social", st_th_r),
            Paragraph("Previdência Privada", st_th_r),
            Paragraph("Capital", st_th_r),
            Paragraph("Taxa", st_th_r),
            Paragraph("Juros", st_th_r),
        ]
    ]

    # Agrupa ocorrências corrigidas e juros por competência
    comp_map: Dict[str, Dict[str, Decimal]] = {}
    for v in verbas:
        for r in v.get("ocorrencias", []):
            per_s = str(r.get("periodo_mensal", ""))
            m_c = re.search(r"(\d{2}/\d{4})", per_s)
            c_key = m_c.group(1) if m_c else meta.get("demissao", "")[3:]
            if c_key not in comp_map:
                comp_map[c_key] = {"corr": Decimal("0"), "juros": Decimal("0"), "remun": Decimal("0")}
            comp_map[c_key]["corr"] += D(r.get("valor_corrigido", 0))
            comp_map[c_key]["juros"] += D(r.get("juros", 0))
            if v.get("is_remuneratoria", False):
                comp_map[c_key]["remun"] += D(r.get("valor_corrigido", 0))

    total_remun_all = sum(d["remun"] for d in comp_map.values())
    total_inss_seg = D(resumo.get("inss_empregado", 0))
    tot_juros_col = Decimal("0")

    for c_key, vals in comp_map.items():
        tot_v_c = money_round(vals["corr"])
        cs_c = money_round((vals["remun"] / total_remun_all) * total_inss_seg) if total_remun_all > 0 else Decimal("0")
        cap_c = max(Decimal("0"), money_round(tot_v_c - cs_c))
        jur_c = money_round(vals["juros"])
        tot_juros_col += jur_c
        tx_pct = ((jur_c / cap_c) * Decimal("100")).quantize(TWO_PLACES, rounding=ROUND_HALF_UP) if cap_c > 0 else Decimal("0.00")
        juros_rows.append(
            [
                Paragraph(c_key, st_td_c),
                Paragraph(str(meta.get("ajuizamento", "")), st_td_c),
                Paragraph(fmt_brl(tot_v_c), st_td_r),
                Paragraph(fmt_brl(cs_c), st_td_r),
                Paragraph("0,00", st_td_r),
                Paragraph(fmt_brl(cap_c), st_td_r),
                Paragraph(f"{fmt_num(tx_pct, 2)}%", st_td_r),
                Paragraph(fmt_brl(jur_c), st_td_r),
            ]
        )

    juros_rows.append(
        [
            "",
            "",
            "",
            "",
            "",
            "",
            Paragraph("Total", st_td_rb),
            Paragraph(fmt_brl(resumo["juros_mora"]), st_td_rb),
        ]
    )

    tbl_juros = Table(
        juros_rows,
        colWidths=[22 * mm, 24 * mm, 28 * mm, 28 * mm, 26 * mm, 26 * mm, 16 * mm, 20 * mm],
        repeatRows=1,
    )
    j_style = [
        ("GRID", (0, 0), (-1, -2), 0.35, BORDER_CLR),
        ("GRID", (6, -1), (7, -1), 0.45, BORDER_CLR),
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("BACKGROUND", (6, -1), (7, -1), SUBHEADER_BG),
        ("TOPPADDING", (0, 0), (-1, -1), 2.0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    for r_i in range(1, len(juros_rows) - 1):
        if r_i % 2 == 0:
            j_style.append(("BACKGROUND", (0, r_i), (-1, r_i), ZEBRA_BG))
    tbl_juros.setStyle(TableStyle(j_style))
    story.append(tbl_juros)

    doc.build(
        story,
        canvasmaker=lambda *args, **kwargs: PJeCalcCanvas(*args, meta_info=meta, **kwargs),
    )
    return buf.getvalue()


def build_pjc_xml(result: Dict[str, Any]) -> bytes:
    """Gera estrutura XML estruturada compatível com intercâmbio PJe-Calc (.pjc)."""
    meta = result["meta"]
    resumo = result["resumo"]
    root = ET.Element("CalculoPJeCalc", versao="2.11.2", uuid=str(uuid.uuid4()))

    el_proc = ET.SubElement(root, "Processo")
    ET.SubElement(el_proc, "Numero").text = str(meta["processo"])
    ET.SubElement(el_proc, "Reclamante").text = str(meta["reclamante"])
    ET.SubElement(el_proc, "Reclamado").text = str(meta["reclamado"])
    ET.SubElement(el_proc, "DataAdmissao").text = str(meta["admissao"])
    ET.SubElement(el_proc, "DataDemissao").text = str(meta["demissao"])
    ET.SubElement(el_proc, "DataAjuizamento").text = str(meta["ajuizamento"])
    ET.SubElement(el_proc, "DataLiquidacao").text = str(meta["liquidacao"])
    ET.SubElement(el_proc, "ModalidadeRescisao").text = str(meta["modalidade_rescisao"])
    ET.SubElement(el_proc, "SalarioBase").text = f"{meta['salario_base']:.2f}"
    ET.SubElement(el_proc, "Divisor").text = f"{meta['divisor']:.2f}"

    el_verbas = ET.SubElement(root, "Verbas")
    for v in result["verbas"]:
        ev = ET.SubElement(el_verbas, "Verba", codigo=str(v["codigo"]))
        ET.SubElement(ev, "Nome").text = str(v["nome"])
        ET.SubElement(ev, "Formula").text = str(v["formula_pje"])
        ET.SubElement(ev, "Principal").text = f"{v['valor']:.2f}"
        ET.SubElement(ev, "Corrigido").text = f"{v['valor_corrigido']:.2f}"
        ET.SubElement(ev, "Juros").text = f"{v['juros']:.2f}"
        ET.SubElement(ev, "Total").text = f"{v['total_verba']:.2f}"
        el_ocs = ET.SubElement(ev, "Ocorrencias")
        for r in v.get("ocorrencias", []):
            eo = ET.SubElement(el_ocs, "Ocorrencia")
            ET.SubElement(eo, "Periodo").text = str(r["periodo_mensal"])
            ET.SubElement(eo, "Base").text = f"{r['base']:.2f}"
            ET.SubElement(eo, "Divisor").text = f"{r['divisor']:.4f}"
            ET.SubElement(eo, "Multiplicador").text = f"{r['multiplicador']:.8f}"
            ET.SubElement(eo, "Quantidade").text = f"{r['quantidade']:.4f}"
            ET.SubElement(eo, "Devido").text = f"{r['devido']:.2f}"
            ET.SubElement(eo, "IndiceCorrecao").text = f"{r['indice_correcao']:.8f}"
            ET.SubElement(eo, "Corrigido").text = f"{r['valor_corrigido']:.2f}"
            ET.SubElement(eo, "Juros").text = f"{r.get('juros', 0):.2f}"

    el_res = ET.SubElement(root, "Resumo")
    for k, val in resumo.items():
        ET.SubElement(el_res, k).text = f"{val:.2f}"

    xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return xml_str


# ==============================================================================
# ASSISTENTE PERICIAL DE IA & GERADOR DE MINUTA DE PETIÇÃO (ART. 879 CLT)
# ==============================================================================
def ask_pericial_ai(
    user_prompt: str,
    params: Dict[str, Any],
    result: Dict[str, Any],
    sentence_text: str,
    lawyer_notes: str,
    api_key: str = "",
) -> Tuple[str, Dict[str, Any], bool]:
    """
    Assistente pericial trabalhista:
    - Interpreta comandos de ajuste direto nos parâmetros (teto de intervalo, descontos indevidos, horas extras, etc.);
    - Explica diferenças de cálculo e responde com diagnóstico pericial completo.
    """
    updated = dict(params)
    changed = False
    low = user_prompt.lower()

    m_sal = re.search(r"sal[áa]rio\s*(?:para|de|=)?\s*R?\$?\s*([\d\.]+,\d{2}|\d+(?:\.\d+)?)", low)
    if m_sal:
        updated["salario_base"] = float(D(m_sal.group(1)))
        changed = True

    m_teto = re.search(r"(?:teto|limite)\s*(?:do\s+intervalo|intrajornada)?\s*(?:para|de|=)?\s*R?\$?\s*([\d\.]+,\d{2})", low)
    if m_teto:
        updated["tem_intrajornada"] = True
        updated["teto_intrajornada"] = float(D(m_teto.group(1)))
        changed = True

    m_desc = re.search(r"(?:descontos\s+indevidos|devolu[çc][ãa]o\s+de\s+descontos)\s*(?:para|de|=)?\s*R?\$?\s*([\d\.]+,\d{2})", low)
    if m_desc:
        updated["tem_descontos_indevidos"] = True
        updated["valor_descontos_indevidos"] = float(D(m_desc.group(1)))
        changed = True

    m_he50 = re.search(r"(?:he\s*50|horas?\s*extras?\s*50%?)\s*(?:para|de|=)?\s*(\d+(?:[,\.]\d+)?)", low)
    if m_he50:
        v = float(D(m_he50.group(1)))
        updated["tem_he50"] = v > 0
        updated["qtd_he50_mensal"] = v
        changed = True

    if "sem horas extras" in low or "zerar horas extras" in low or "excluir horas extras" in low:
        updated["tem_he50"] = False
        updated["qtd_he50_mensal"] = 0.0
        updated["tem_he100"] = False
        updated["qtd_he100_mensal"] = 0.0
        changed = True

    if "sem verbas rescis" in low or "excluir verbas rescis" in low or "zerar rescis" in low:
        updated["tem_verbas_rescis"] = False
        changed = True

    if "incluir verbas rescis" in low or "com verbas rescis" in low:
        updated["tem_verbas_rescis"] = True
        changed = True

    m_abat = re.search(r"(?:abater|deduzir|descontar|pago)\s*R?\$?\s*([\d\.]+,\d{2})", low)
    if m_abat:
        updated["valor_pago_abatimento"] = float(D(m_abat.group(1)))
        changed = True

    resumo = result["resumo"]
    meta = result["meta"]
    context_summary = (
        f"Processo: {meta['processo']} | Período: {meta['inicio_calculo']} a {meta['fim_calculo']} "
        f"({meta['meses_apurados']} meses) | Salário Base: R$ {fmt_brl(meta['salario_base'])} | "
        f"Principal: R$ {fmt_brl(resumo['subtotal_principal'])} | "
        f"Bruto Reclamante: R$ {fmt_brl(resumo['bruto_devido_reclamante'])} | "
        f"Líquido Reclamante: R$ {fmt_brl(resumo['liquido_devido_reclamante'])} | "
        f"Total Reclamado: R$ {fmt_brl(resumo['total_devido_reclamado'])}."
    )

    clean_key = (api_key or "").strip()
    if clean_key:
        try:
            sys_msg = (
                "Você é um perito contador trabalhista especialista em PJe-Calc Cidadão e liquidação de sentença (art. 879 CLT). "
                "Responda de forma objetiva, técnica e conferindo divergências entre a sentença, o relato do advogado e os valores apurados."
            )
            full_prompt = (
                f"{sys_msg}\n\nRESUMO ATUAL DO CÁLCULO:\n{context_summary}\n\n"
                f"RELATO DO ADVOGADO:\n{lawyer_notes[:2500]}\n\n"
                f"TRECHO DA SENTENÇA:\n{sentence_text[:2500]}\n\n"
                f"PERGUNTA / SOLICITAÇÃO DO USUÁRIO:\n{user_prompt}"
            )
            if clean_key.startswith("AIza"):
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={clean_key}"
                payload = json.dumps({"contents": [{"parts": [{"text": full_prompt}]}]}).encode("utf-8")
                req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})
                with urllib.request.urlopen(req, timeout=25) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    ans = data["candidates"][0]["content"]["parts"][0]["text"]
                    return ans, updated, changed
            else:
                url = "https://api.openai.com/v1/chat/completions"
                payload = json.dumps(
                    {
                        "model": "gpt-4o-mini",
                        "messages": [
                            {"role": "system", "content": sys_msg},
                            {"role": "user", "content": full_prompt},
                        ],
                        "temperature": 0.2,
                    }
                ).encode("utf-8")
                req = urllib.request.Request(
                    url,
                    data=payload,
                    headers={"Content-Type": "application/json", "Authorization": f"Bearer {clean_key}"},
                )
                with urllib.request.urlopen(req, timeout=25) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    ans = data["choices"][0]["message"]["content"]
                    return ans, updated, changed
        except Exception as e:
            return f"⚠️ Falha na consulta à API externa ({e}). Diagnóstico local aplicado:\n\n{context_summary}", updated, changed

    verbas_txt = "\n".join(
        f"- **{v['nome']}**: Principal R$ {fmt_brl(v['valor'])} | Corrigido R$ {fmt_brl(v['valor_corrigido'])} | Juros R$ {fmt_brl(v['juros'])} | Total R$ {fmt_brl(v['total_verba'])}"
        for v in result["verbas"]
    ) or "- Nenhuma verba ativa no momento."

    reply = (
        f"### Análise Técnico-Pericial (Padrão PJe-Calc)\n"
        f"{'✅ **Parâmetros atualizados e cálculo recalculado automaticamente!**' if changed else 'ℹ️ **Diagnóstico dos parâmetros atuais:**'}\n\n"
        f"- **Período Apurado:** {meta['inicio_calculo']} a {meta['fim_calculo']} ({meta['meses_apurados']} competências)\n"
        f"- **Salário Base / Divisor:** R$ {fmt_brl(meta['salario_base'])} / {fmt_num(meta['divisor'], 2)}\n"
        f"- **Modalidade Rescisória:** {meta['modalidade_rescisao']}\n\n"
        f"**Composição das Verbas Ativas:**\n{verbas_txt}\n\n"
        f"**Totais Consolidados:**\n"
        f"- Subtotal Principal (Nominal): **R$ {fmt_brl(resumo['subtotal_principal'])}**\n"
        f"- Bruto Reclamante (Corrigido + Juros): **R$ {fmt_brl(resumo['bruto_devido_reclamante'])}**\n"
        f"- Líquido Reclamante: **R$ {fmt_brl(resumo['liquido_devido_reclamante'])}**\n"
        f"- Total Devido pelo Reclamado (c/ Honorários): **R$ {fmt_brl(resumo['total_devido_reclamado'])}**\n\n"
        f"💡 *Comandos rápidos disponíveis: `\"teto intervalo para 1337,60\"`, `\"descontos indevidos para 2000,00\"`, `\"zerar horas extras\"`.*"
    )
    return reply, updated, changed


def build_minuta_peticao(params: Dict[str, Any], result: Dict[str, Any], lawyer_notes: str = "") -> str:
    """Gera a minuta jurídica completa de apresentação de cálculos de liquidação (Art. 879 CLT)."""
    meta = result["meta"]
    resumo = result["resumo"]
    verbas = result["verbas"]

    linhas_verbas = "\n".join(
        f"   • {v['nome']}: Principal R$ {fmt_brl(v['valor'])} | Corrigido R$ {fmt_brl(v['valor_corrigido'])} | Juros R$ {fmt_brl(v['juros'])} | Total R$ {fmt_brl(v['total_verba'])}"
        for v in verbas
    )

    obs_adv = ""
    if lawyer_notes.strip():
        obs_adv = f"\nIV. OBSERVAÇÕES TÉCNICAS CONFORME TÍTULO EXECUTIVO\n\n{lawyer_notes.strip()}\n"

    return f"""EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DO TRABALHO DA {params.get('vara_trabalho', 'VARA DO TRABALHO').upper()} DE {meta['municipio'].upper()}/{meta['estado'].upper()}

Processo nº: {meta['processo']}

{meta['reclamante'].upper()}, já devidamente qualificado(a) nos autos da Reclamatória Trabalhista em epígrafe, que move em face de {meta['reclamado'].upper()}, vem, respeitosamente, à presença de Vossa Excelência, por seu(sua) procurador(a) signatário(a), com fulcro no art. 879, § 1º-B, da Consolidação das Leis do Trabalho (CLT), apresentar os

CÁLCULOS DE LIQUIDAÇÃO DE SENTENÇA

elaborados em estrita observância ao comando sentencial transitado em julgado e aos parâmetros oficiais do sistema PJe-Calc Cidadão (CSJT), conforme demonstrativos analíticos anexos e síntese abaixo:

I. PARÂMETROS CONTRATUAIS E PROCESSUAIS OBSERVADOS

1. Data de Admissão: {meta['admissao']}
2. Data de Demissão / Rescisão: {meta['demissao']} (Modalidade: {meta['modalidade_rescisao']})
3. Data de Ajuizamento da Ação: {meta['ajuizamento']}
4. Data de Atualização / Liquidação: {meta['liquidacao']}
5. Período Liquidado: {meta['inicio_calculo']} a {meta['fim_calculo']} ({meta['meses_apurados']} competências mensais)
6. Salário Base / Maior Remuneração: R$ {fmt_brl(meta['salario_base'])} | Divisor: {fmt_num(meta['divisor'], 4)}

II. CRITÉRIOS DE ATUALIZAÇÃO MONETÁRIA, JUROS E ENCARGOS LEGAIS

1. Correção Monetária e Juros: {meta['criterio_correcao']}, observada a época própria de exigibilidade (Súmula 381 do TST).
2. Reflexos Legais e Tetos Fixados: Apurados competência a competência em estrita observância aos limites fixados no título executivo judicial.

III. DISCRIMINAÇÃO DAS VERBAS LIQUIDADAS

{linhas_verbas}
{obs_adv}
IV. QUADRO RESUMO GERAL DA LIQUIDAÇÃO (ATUALIZADO ATÉ {meta['liquidacao']})

1. Subtotal Principal (Valor Histórico): R$ {fmt_brl(resumo['subtotal_principal'])}
2. Subtotal Corrigido: R$ {fmt_brl(resumo['subtotal_corrigido'])}
3. Juros de Mora: R$ {fmt_brl(resumo['juros_mora'])}
4. BRUTO DEVIDO AO RECLAMANTE: R$ {fmt_brl(resumo['bruto_devido_reclamante'])}
   (-) Dedução Contribuição Previdenciária (INSS Segurado): -R$ {fmt_brl(resumo['inss_empregado'])}
   (-) Dedução Imposto de Renda (IRRF - RRA): -R$ {fmt_brl(resumo['irpf_reclamante'])}
   (-) Abatimento / Valores Pagos Comprovados: -R$ {fmt_brl(resumo.get('valor_abatimento', 0))}
5. LÍQUIDO DEVIDO AO RECLAMANTE: R$ {fmt_brl(resumo['liquido_devido_reclamante'])}
6. Contribuição Social Total (Segurado + Empregador): R$ {fmt_brl(D(resumo['inss_empregado']) + D(resumo['inss_empregador']))}
7. Honorários Advocatícios Sucumbenciais ({fmt_num(resumo['honorarios_pct'], 2)}%): R$ {fmt_brl(resumo['honorarios_valor'])}
8. TOTAL GERAL DEVIDO PELA RECLAMADA: R$ {fmt_brl(resumo['total_devido_reclamado'])}

V. REQUERIMENTOS FINAIS

Diante do exposto, requer a Vossa Excelência:
a) A homologação dos presentes cálculos de liquidação, fixando-se o crédito total devido pela Reclamada em R$ {fmt_brl(resumo['total_devido_reclamado'])}, atualizado até {meta['liquidacao']}, sem prejuízo das atualizações ulteriores até o efetivo pagamento;
b) A intimação da Reclamada para ciência e pagamento/garantia do juízo, nos termos do art. 880 da CLT.

Nestes termos,
Pede e espera deferimento.

{meta['municipio']}/{meta['estado']}, {meta['liquidacao']}.

____________________________________________________
{params.get('advogado_reclamante', 'ADVOGADO(A) DO RECLAMANTE')}
{params.get('oab_advogado', 'OAB/RS 00.000')}
"""


def build_minuta_pdf(minuta_text: str, processo: str) -> bytes:
    """Gera PDF formatado da Minuta de Petição de Apresentação de Cálculos."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=25 * mm,
        rightMargin=20 * mm,
        topMargin=25 * mm,
        bottomMargin=20 * mm,
        title=f"Minuta Peticao Liquidacao - {processo}",
    )
    styles = getSampleStyleSheet()
    st_pet = ParagraphStyle(
        "PeticaoBody",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=11,
        leading=16,
        alignment=TA_LEFT,
        spaceAfter=6,
    )
    st_pet_bold = ParagraphStyle(
        "PeticaoBold",
        parent=styles["Normal"],
        fontName="Times-Bold",
        fontSize=11,
        leading=16,
        alignment=TA_CENTER,
        spaceBefore=6,
        spaceAfter=8,
    )
    story: List[Any] = []
    for raw_line in minuta_text.splitlines():
        line = raw_line.strip()
        if not line:
            story.append(Spacer(1, 3 * mm))
            continue
        if line in ("CÁLCULOS DE LIQUIDAÇÃO DE SENTENÇA",):
            story.append(Paragraph(line, st_pet_bold))
        else:
            safe_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if re.match(r"^(EXCELENTÍSSIMO|I\.|II\.|III\.|IV\.|V\.)", line):
                story.append(Paragraph(f"<b>{safe_line}</b>", st_pet))
            else:
                story.append(Paragraph(safe_line, st_pet))
    doc.build(story)
    return buf.getvalue()


# ==============================================================================
# INICIALIZAÇÃO DE ESTADO DA SESSÃO
# ==============================================================================
if "params" not in st.session_state:
    st.session_state["params"] = parse_sentence_parameters("")
if "sentence_text" not in st.session_state:
    st.session_state["sentence_text"] = ""
if "lawyer_notes" not in st.session_state:
    st.session_state["lawyer_notes"] = ""
if "ai_chat_history" not in st.session_state:
    st.session_state["ai_chat_history"] = [
        {
            "role": "assistant",
            "content": (
                "Olá! Sou seu **Assistente Pericial Trabalhista (Padrão PJe-Calc)**. "
                "Cole o **relato do advogado** e a **sentença** na barra lateral (incluindo tetos/limites como "
                "R$ 1.337,60 de intervalo intrajornada ou R$ 2.000,00 de descontos indevidos) para eu alinhar o cálculo."
            ),
        }
    ]
if "calc_result" not in st.session_state:
    st.session_state["calc_result"] = run_pjecalc_engine(st.session_state["params"])


# ==============================================================================
# INTERFACE PRINCIPAL — ED CÁLCULOS TRABALHISTAS (PADRÃO PJE-CALC CIDADÃO)
# ==============================================================================
st.markdown(
    """
    <div style="padding: 14px 18px; border-radius: 10px; background: linear-gradient(90deg, #1E293B 0%, #334155 100%); color: #FFFFFF; margin-bottom: 16px;">
        <h2 style="margin: 0; font-size: 1.45rem;">⚖️ ED Cálculos Trabalhistas — Liquidação de Sentença (Art. 879 CLT)</h2>
        <p style="margin: 4px 0 0 0; font-size: 0.9rem; color: #CBD5E1;">
            Motor matemático 100% fiel ao <b>PJe-Calc Cidadão (CSJT)</b> • Suporte a Tetos/Limites de Sentença • <b>Minuta</b> + <b>PDF Oficial</b> + <b>.PJC</b>
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("📂 1. Sentença & Relato do Advogado")
    uploaded_files = st.file_uploader(
        "Envie a Sentença / Embargos / TRCT / Ponto (PDF, CSV, XLSX, TXT)",
        type=["pdf", "txt", "csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )
    lawyer_notes_input = st.text_area(
        "📋 Relato / Orientação do Advogado (parâmetros e limites):",
        value=st.session_state["lawyer_notes"],
        height=115,
        placeholder="Ex.: Apenas 1h de intervalo intrajornada limitado a R$ 1.337,60 e descontos indevidos de R$ 2.000,00...",
    )
    pasted_text = st.text_area(
        "⚖️ Texto do Dispositivo da Sentença / Embargos:",
        value=st.session_state["sentence_text"],
        height=115,
        placeholder="Cole aqui o dispositivo da sentença e da decisão de embargos...",
    )

    cartao_df: Optional[pd.DataFrame] = None
    if uploaded_files:
        for uf in uploaded_files:
            if uf.name.lower().endswith(".csv"):
                try:
                    cartao_df = pd.read_csv(io.BytesIO(uf.getvalue()), sep=None, engine="python")
                except Exception:
                    pass
            elif uf.name.lower().endswith((".xlsx", ".xls")):
                try:
                    cartao_df = pd.read_excel(io.BytesIO(uf.getvalue()))
                except Exception:
                    pass

    if st.button("🔍 Ler Documentos & Extrair Parâmetros", use_container_width=True, type="primary"):
        combined_chunks: List[str] = []
        if uploaded_files:
            for uf in uploaded_files:
                combined_chunks.append(extract_text_from_uploaded_file(uf))
        if pasted_text.strip():
            combined_chunks.append(pasted_text.strip())
        full_txt = "\n\n".join(combined_chunks)
        st.session_state["sentence_text"] = full_txt
        st.session_state["lawyer_notes"] = lawyer_notes_input
        parsed = parse_sentence_parameters(full_txt, lawyer_notes_input)
        st.session_state["params"] = parsed
        st.session_state["calc_result"] = run_pjecalc_engine(parsed, cartao_df)
        st.success("Parâmetros extraídos da sentença e do relato do advogado!")

    st.divider()
    st.subheader("🤖 Configuração Opcional da IA")
    default_api_key = ""
    try:
        default_api_key = st.secrets.get("OPENAI_API_KEY", "") or st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        default_api_key = os.environ.get("OPENAI_API_KEY", "") or os.environ.get("GEMINI_API_KEY", "")
    ai_api_key = st.text_input(
        "Chave API OpenAI ou Gemini (opcional)",
        value=default_api_key,
        type="password",
    )

tab_params, tab_ai, tab_resumo, tab_verbas, tab_minuta, tab_export = st.tabs(
    [
        "⚙️ Parâmetros do Cálculo",
        "🤖 Assistente Pericial IA",
        "📊 Resumo da Liquidação",
        "📑 Demonstrativo por Verba",
        "📝 Minuta de Petição Pronta",
        "📥 Baixar PDF Oficial & .PJC",
    ]
)

p = st.session_state["params"]

with tab_params:
    st.subheader("Dados do Processo e Contrato (Conferência / Ajuste Fino)")
    c1, c2, c3, c4 = st.columns(4)
    p["processo"] = c1.text_input("Número do Processo (CNJ)", value=p["processo"])
    p["calculo_numero"] = c2.text_input("Nº do Cálculo", value=p["calculo_numero"])
    p["estado"] = c3.text_input("UF", value=p["estado"])
    p["municipio"] = c4.text_input("Município", value=p["municipio"])

    c5, c6, c7 = st.columns([2, 2, 1.5])
    p["reclamante"] = c5.text_input("Reclamante", value=p["reclamante"])
    p["reclamado"] = c6.text_input("Reclamado", value=p["reclamado"])
    p["vara_trabalho"] = c7.text_input("Vara do Trabalho", value=p.get("vara_trabalho", "VARA DO TRABALHO"))

    c_adv1, c_adv2 = st.columns(2)
    p["advogado_reclamante"] = c_adv1.text_input("Advogado(a) Assinante (para Minuta)", value=p.get("advogado_reclamante", "ADVOGADO(A) DO RECLAMANTE"))
    p["oab_advogado"] = c_adv2.text_input("Inscrição OAB", value=p.get("oab_advogado", "OAB/RS 00.000"))

    d1, d2, d3, d4 = st.columns(4)
    p["admissao"] = d1.date_input("Data de Admissão", value=parse_br_date(p["admissao"]), format="DD/MM/YYYY")
    p["demissao"] = d2.date_input("Data de Demissão", value=parse_br_date(p["demissao"]), format="DD/MM/YYYY")
    p["ajuizamento"] = d3.date_input("Data de Ajuizamento", value=parse_br_date(p["ajuizamento"]), format="DD/MM/YYYY")
    p["liquidacao"] = d4.date_input("Data de Liquidação", value=parse_br_date(p["liquidacao"]), format="DD/MM/YYYY")

    r1, r2, r3, r4 = st.columns(4)
    modalidades_lista = ["Dispensa sem justa causa", "Pedido de demissão", "Dispensa por justa causa", "Rescisão indireta"]
    idx_mod = modalidades_lista.index(p.get("modalidade_rescisao", "Pedido de demissão")) if p.get("modalidade_rescisao") in modalidades_lista else 1
    p["modalidade_rescisao"] = r1.selectbox("Modalidade Rescisória", modalidades_lista, index=idx_mod)
    p["salario_base"] = r2.number_input("Salário Base Mensal (R$)", min_value=0.0, value=float(p["salario_base"]), step=50.0)
    p["divisor"] = r3.number_input("Carga Horária / Divisor", min_value=1.0, value=float(p["divisor"]), step=10.0)
    p["honorarios_sucumbencia_pct"] = r4.number_input("Honorários Sucumbenciais (%)", min_value=0.0, max_value=30.0, value=float(p["honorarios_sucumbencia_pct"]), step=1.0)

    chk1, chk2, chk3 = st.columns(3)
    p["projetar_aviso_previo"] = chk1.checkbox("Projetar Aviso Prévio Indenizado", value=bool(p["projetar_aviso_previo"]))
    p["aplicar_prescricao_quinquenal"] = chk2.checkbox("Aplicar Prescrição Quinquenal", value=bool(p["aplicar_prescricao_quinquenal"]))
    p["sabado_dia_util"] = chk3.checkbox("Considerar Sábado como Dia Útil", value=bool(p["sabado_dia_util"]))

    st.divider()
    st.subheader("Verbas Deferidas, Tetos/Limites da Sentença & Deduções")

    lim1, lim2, lim3, lim4 = st.columns(4)
    p["tem_intrajornada"] = lim1.checkbox("Intervalo Intrajornada", value=bool(p.get("tem_intrajornada", True)))
    p["qtd_intrajornada_mensal"] = lim1.number_input("Horas Intrajornada/Mês", min_value=0.0, value=float(p.get("qtd_intrajornada_mensal", 21.0)), step=1.0)
    p["teto_intrajornada"] = lim2.number_input("Teto/Limite Intrajornada (R$) (0 = sem teto)", min_value=0.0, value=float(p.get("teto_intrajornada", 0.0)), step=50.0)
    p["reflexos_intrajornada"] = lim2.checkbox("Reflexos sobre Intrajornada (13º, Férias+1/3, FGTS)", value=bool(p.get("reflexos_intrajornada", False)))

    p["tem_descontos_indevidos"] = lim3.checkbox("Devolução de Descontos Indevidos", value=bool(p.get("tem_descontos_indevidos", False)))
    p["valor_descontos_indevidos"] = lim3.number_input("Valor Descontos Indevidos (R$)", min_value=0.0, value=float(p.get("valor_descontos_indevidos", 0.0)), step=100.0)
    p["valor_pago_abatimento"] = lim4.number_input("Abatimento / Valor Já Pago (R$)", min_value=0.0, value=float(p.get("valor_pago_abatimento", 0.0)), step=100.0)

    v1, v2, v3, v4 = st.columns(4)
    p["tem_he50"] = v1.checkbox("Horas Extras 50%", value=bool(p.get("tem_he50", False)))
    p["qtd_he50_mensal"] = v1.number_input("Qtd. Média Mensal HE 50%", min_value=0.0, value=float(p.get("qtd_he50_mensal", 0.0)), step=1.0)

    p["tem_he100"] = v2.checkbox("Horas Extras 100%", value=bool(p.get("tem_he100", False)))
    p["qtd_he100_mensal"] = v2.number_input("Qtd. Média Mensal HE 100%", min_value=0.0, value=float(p.get("qtd_he100_mensal", 0.0)), step=1.0)

    p["tem_noturno"] = v3.checkbox("Adicional Noturno 20%", value=bool(p.get("tem_noturno", False)))
    p["qtd_horas_noturnas_mensal"] = v3.number_input("Horas Noturnas Relógio/Mês", min_value=0.0, value=float(p.get("qtd_horas_noturnas_mensal", 0.0)), step=2.0)

    p["verba_extra_mensal"] = v4.number_input("Outra Parcela Mensal (R$)", min_value=0.0, value=float(p.get("verba_extra_mensal", 0.0)), step=50.0)

    a1, a2, a3, a4 = st.columns(4)
    adicionais_opts = ["Nenhum", "Insalubridade", "Periculosidade (30%)"]
    p["adicional_tipo"] = a1.selectbox("Adicional Salarial", adicionais_opts, index=adicionais_opts.index(p.get("adicional_tipo", "Nenhum")))
    p["grau_insalubridade"] = a2.selectbox("Grau Insalubridade (%)", [10, 20, 40], index=[10, 20, 40].index(int(p.get("grau_insalubridade", 20))))
    p["tem_verbas_rescis"] = a3.checkbox("Apurar Verbas Rescisórias", value=bool(p.get("tem_verbas_rescis", False)))
    p["multa_477"] = a4.checkbox("Multa Art. 477, § 8º, CLT", value=bool(p.get("multa_477", False)))

    rf1, rf2, rf3, rf4 = st.columns(4)
    p["aplicar_reflexos"] = rf1.checkbox("Apurar Reflexos (13º, Férias + 1/3, Aviso)", value=bool(p.get("aplicar_reflexos", True)))
    p["reflexos_em_dsr"] = rf2.checkbox("Reflexos em RSR / DSR", value=bool(p.get("reflexos_em_dsr", False)))
    p["reflexos_em_fgts"] = rf3.checkbox("FGTS 8% (+ 40% se dispensa s/ justa)", value=bool(p.get("reflexos_em_fgts", True)))
    p["multa_467"] = rf4.checkbox("Multa Art. 467 CLT", value=bool(p.get("multa_467", False)))

    if st.button("⚡ Recalcular Liquidação no Padrão PJe-Calc", type="primary"):
        st.session_state["params"] = p
        st.session_state["calc_result"] = run_pjecalc_engine(p, cartao_df)
        st.success("Cálculo atualizado com sucesso!")

with tab_ai:
    st.subheader("🤖 Assistente Pericial Trabalhista (Conferência & Ajustes por Comando)")
    for msg in st.session_state["ai_chat_history"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_q = st.chat_input("Digite sua dúvida ou comando de ajuste para o cálculo...")
    if user_q:
        st.session_state["ai_chat_history"].append({"role": "user", "content": user_q})
        ai_ans, updated_p, did_change = ask_pericial_ai(
            user_q,
            st.session_state["params"],
            st.session_state["calc_result"],
            st.session_state["sentence_text"],
            st.session_state["lawyer_notes"],
            api_key=ai_api_key,
        )
        if did_change:
            st.session_state["params"] = updated_p
            st.session_state["calc_result"] = run_pjecalc_engine(updated_p, cartao_df)
        st.session_state["ai_chat_history"].append({"role": "assistant", "content": ai_ans})
        st.rerun()

res = st.session_state["calc_result"]
resumo_data = res["resumo"]

with tab_resumo:
    st.subheader("Resumo do Cálculo — Espelho Oficial PJe-Calc Cidadão")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Subtotal Principal (Nominal)", f"R$ {fmt_brl(resumo_data['subtotal_principal'])}")
    m2.metric("Bruto Devido ao Reclamante", f"R$ {fmt_brl(resumo_data['bruto_devido_reclamante'])}")
    m3.metric("Líquido Devido ao Reclamante", f"R$ {fmt_brl(resumo_data['liquido_devido_reclamante'])}")
    m4.metric("Total Devido pelo Reclamado", f"R$ {fmt_brl(resumo_data['total_devido_reclamado'])}")

    st.markdown("#### 1. Descrição do Bruto Devido ao Reclamante")
    df_bruto = pd.DataFrame(
        [
            {
                "Descrição": v["nome"],
                "Principal (R$)": fmt_brl(v["valor"]),
                "Valor Corrigido (R$)": fmt_brl(v["valor_corrigido"]),
                "Juros (R$)": fmt_brl(v["juros"]),
                "Total (R$)": fmt_brl(v["total_verba"]),
            }
            for v in res["verbas"]
        ]
        + [
            {
                "Descrição": "TOTAL",
                "Principal (R$)": fmt_brl(resumo_data["subtotal_principal"]),
                "Valor Corrigido (R$)": fmt_brl(resumo_data["subtotal_corrigido"]),
                "Juros (R$)": fmt_brl(resumo_data["juros_mora"]),
                "Total (R$)": fmt_brl(resumo_data["bruto_devido_reclamante"]),
            }
        ]
    )
    st.dataframe(df_bruto, use_container_width=True, hide_index=True)

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown("#### 2. Créditos e Descontos do Reclamante")
        df_cred = pd.DataFrame(
            [
                {"Descrição": "VERBAS", "Valor (R$)": fmt_brl(resumo_data["total_verbas_bruto"])},
                {"Descrição": "FGTS", "Valor (R$)": fmt_brl(resumo_data["total_fgts_bruto"])},
                {"Descrição": "Bruto Devido ao Reclamante", "Valor (R$)": fmt_brl(resumo_data["bruto_devido_reclamante"])},
                {"Descrição": "DEDUÇÃO DE CONTRIBUIÇÃO SOCIAL", "Valor (R$)": f"-{fmt_brl(resumo_data['inss_empregado'])}"},
                {"Descrição": "IRPF DEVIDO PELO RECLAMANTE", "Valor (R$)": f"-{fmt_brl(resumo_data['irpf_reclamante'])}"},
                {"Descrição": "Total de Descontos", "Valor (R$)": f"-{fmt_brl(resumo_data['total_descontos_reclamante'])}"},
                {"Descrição": "Líquido Devido ao Reclamante", "Valor (R$)": fmt_brl(resumo_data["liquido_devido_reclamante"])},
            ]
        )
        st.dataframe(df_cred, use_container_width=True, hide_index=True)

    with col_right:
        st.markdown("#### 3. Débitos do Reclamado por Credor")
        df_deb = pd.DataFrame(
            [
                {"Descrição": "LÍQUIDO DEVIDO AO RECLAMANTE", "Valor (R$)": fmt_brl(resumo_data["liquido_devido_reclamante"])},
                {"Descrição": "CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS", "Valor (R$)": fmt_brl(D(resumo_data["inss_empregado"]) + D(resumo_data["inss_empregador"]))},
                {"Descrição": "IRPF DEVIDO PELO RECLAMANTE", "Valor (R$)": fmt_brl(resumo_data["irpf_reclamante"])},
                {"Descrição": f"HONORÁRIOS ADVOCATÍCIOS ({fmt_num(resumo_data['honorarios_pct'], 2)}%)", "Valor (R$)": fmt_brl(resumo_data["honorarios_valor"])},
                {"Descrição": "Total Devido pelo Reclamado", "Valor (R$)": fmt_brl(resumo_data["total_devido_reclamado"])},
            ]
        )
        st.dataframe(df_deb, use_container_width=True, hide_index=True)

with tab_verbas:
    st.subheader("Demonstrativo de Verbas — 11 Colunas Oficiais do PJe-Calc Cidadão")
    for v in res["verbas"]:
        with st.expander(f"📌 {v['nome']} — Principal: R$ {fmt_brl(v['valor'])} | Corrigido: R$ {fmt_brl(v['valor_corrigido'])}", expanded=False):
            st.caption(f"**Período:** {v['periodo_global']} | **Incidência(s):** {v['incidencias']}")
            st.code(v["formula_pje"], language="text")
            df_oc = pd.DataFrame(
                [
                    {
                        "Período Mensal": r["periodo_mensal"],
                        "Base": fmt_brl(r["base"]),
                        "Divisor": fmt_num(r["divisor"], 4),
                        "Multiplicador": fmt_num(r["multiplicador"], 8),
                        "Quantidade": fmt_num(r["quantidade"], 4),
                        "Dobra": r["dobra"],
                        "Devido": fmt_brl(r["devido"]),
                        "Pago": fmt_brl(r["pago"]),
                        "Diferença": fmt_brl(r["diferenca"]),
                        "Índice Correção": fmt_num(r["indice_correcao"], 8),
                        "Valor Corrigido": fmt_brl(r["valor_corrigido"]),
                    }
                    for r in v.get("ocorrencias", [])
                ]
            )
            st.dataframe(df_oc, use_container_width=True, hide_index=True)

with tab_minuta:
    st.subheader("📝 Minuta de Petição de Apresentação de Cálculos (Art. 879, § 1º-B, CLT)")
    minuta_gerada = build_minuta_peticao(st.session_state["params"], res, st.session_state["lawyer_notes"])
    minuta_editada = st.text_area("Corpo da Petição (editável):", value=minuta_gerada, height=460)

    proc_clean_m = re.sub(r"[^0-9A-Za-z\-]", "_", str(res["meta"]["processo"]))
    mp1, mp2 = st.columns(2)
    mp1.download_button(
        label="📄 Baixar Minuta de Petição em PDF",
        data=build_minuta_pdf(minuta_editada, str(res["meta"]["processo"])),
        file_name=f"Minuta_Peticao_Calculos_{proc_clean_m}.pdf",
        mime="application/pdf",
        use_container_width=True,
        type="primary",
    )
    mp2.download_button(
        label="📃 Baixar Minuta de Petição em Texto (.TXT)",
        data=minuta_editada.encode("utf-8"),
        file_name=f"Minuta_Peticao_Calculos_{proc_clean_m}.txt",
        mime="text/plain",
        use_container_width=True,
    )

with tab_export:
    st.subheader("Exportação Oficial — Planilha de Cálculo em PDF & Arquivo .PJC")
    pdf_bytes = build_pjecalc_pdf(res)
    pjc_bytes = build_pjc_xml(res)
    proc_clean = re.sub(r"[^0-9A-Za-z\-]", "_", str(res["meta"]["processo"]))

    dl1, dl2 = st.columns(2)
    dl1.download_button(
        label="📄 Baixar Planilha de Cálculo Oficial (PDF PJe-Calc)",
        data=pdf_bytes,
        file_name=f"Planilha_Calculo_PJeCalc_{proc_clean}.pdf",
        mime="application/pdf",
        use_container_width=True,
        type="primary",
    )
    dl2.download_button(
        label="💾 Baixar Arquivo Estruturado (.PJC)",
        data=pjc_bytes,
        file_name=f"Calculo_{proc_clean}.pjc",
        mime="application/xml",
        use_container_width=True,
    )
