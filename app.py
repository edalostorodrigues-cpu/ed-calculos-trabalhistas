# app.py — ED Cálculos Trabalhistas (Versão Python + Streamlit)
# Requisitos: streamlit, pandas, openpyxl, pydantic, google-genai, fpdf2
# Execução: streamlit run app.py

import os
import io
import json
import hashlib
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Optional, Dict, Any

import pandas as pd
import streamlit as st
from pydantic import BaseModel
from fpdf import FPDF

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL
# ==============================================================================
st.set_page_config(
    page_title="ED Cálculos Trabalhistas — Parecer & Auditoria Pré-Processual",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp { background-color: #F4F6F9; }
    section[data-testid="stSidebar"] { background-color: #0B1B2B; color: #FFFFFF; }
    section[data-testid="stSidebar"] * { color: #F8FAFC !important; }
    .stMetric {
        background-color: #FFFFFF;
        padding: 14px 18px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 2. TABELAS HISTÓRICAS E CONSTANTES DECIMAIS
# ==============================================================================
TABLES_VERSION = "2026.11-completa"
CENT = Decimal("0.01")

# Tabela histórica oficial do Salário Mínimo Nacional (para base de insalubridade)
MIN_WAGE_BY_YEAR = {
    2020: Decimal("1045.00"),
    2021: Decimal("1100.00"),
    2022: Decimal("1212.00"),
    2023: Decimal("1320.00"),
    2024: Decimal("1412.00"),
    2025: Decimal("1518.00"),
    2026: Decimal("1621.00"),
}


def get_historical_min_wage(ref_date: date) -> Decimal:
    if ref_date.year in MIN_WAGE_BY_YEAR:
        return MIN_WAGE_BY_YEAR[ref_date.year]
    return MIN_WAGE_BY_YEAR[max(MIN_WAGE_BY_YEAR.keys())]


def d_money(val: Any) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    return Decimal(str(val)).quantize(CENT, rounding=ROUND_HALF_UP)


def fmt_brl(val: Decimal) -> str:
    q = d_money(val)
    sign = "-" if q < 0 else ""
    abs_q = abs(q)
    inteiro, dec = f"{abs_q:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
    return f"{sign}R$ {inteiro_fmt},{dec}"


# ==============================================================================
# 3. ESTADO DA SESSÃO
# ==============================================================================
if "calculations_history" not in st.session_state:
    st.session_state.calculations_history = []
if "timecard_result" not in st.session_state:
    st.session_state.timecard_result = None
if "extracted_ai_data" not in st.session_state:
    st.session_state.extracted_ai_data = None
if "extraction_confirmed" not in st.session_state:
    st.session_state.extraction_confirmed = False


# ==============================================================================
# 4. APURADOR DETERMINÍSTICO DE CARTÃO DE PONTO (EM MINUTOS INTEIROS, SEM IA)
# ==============================================================================
def parse_hhmm(val: Any) -> Optional[int]:
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or ":" not in s:
        return None
    parts = s.split(":")
    try:
        h, m = int(parts[0]), int(parts[1])
        if 0 <= h <= 29 and 0 <= m <= 59:
            return h * 60 + m
    except ValueError:
        return None
    return None


def calc_night_minutes(start_min: int, end_min: int) -> int:
    if end_min <= start_min:
        return 0
    night = 0
    for n_start, n_end in [(0, 300), (1320, 1740)]:
        ov_s = max(start_min, n_start)
        ov_e = min(end_min, n_end)
        if ov_e > ov_s:
            night += ov_e - ov_s
    return night


def process_timecard_dataframe(df: pd.DataFrame, daily_limit_min: int = 480) -> Dict[str, Any]:
    rows_out = []
    tot_he50_min = 0
    tot_he100_min = 0
    tot_intra_min = 0
    tot_night_ficta_min = 0
    incomplete_days = 0

    for _, row in df.iterrows():
        dt_str = str(row.get("Data", "")).strip()
        dia_sem = str(row.get("Dia", row.get("DiaSemana", ""))).strip().upper()
        feriado = str(row.get("Feriado", "N")).strip().upper() in ("S", "SIM", "1", "TRUE") or dia_sem == "DOM"

        e1 = parse_hhmm(row.get("Entrada1"))
        s1 = parse_hhmm(row.get("Saida1"))
        e2 = parse_hhmm(row.get("Entrada2"))
        s2 = parse_hhmm(row.get("Saida2"))

        if e1 is None or s1 is None:
            incomplete_days += 1
            rows_out.append({
                "Data": dt_str, "Dia": dia_sem, "HE 50% (h)": 0.0,
                "HE 100% (h)": 0.0, "Intrajornada (h)": 0.0, "Noturna Ficta (h)": 0.0,
                "Status": "Batida incompleta (revisar)",
            })
            continue

        p1_end = s1 + 1440 if s1 < e1 else s1
        w1 = max(0, p1_end - e1)
        n1 = calc_night_minutes(e1, p1_end)

        w2, n2, break_min = 0, 0, 0
        if e2 is not None and s2 is not None:
            p2_start = e2 + 1440 if e2 < p1_end else e2
            p2_end = s2 + 1440 if s2 < p2_start else s2
            break_min = max(0, p2_start - p1_end)
            w2 = max(0, p2_end - p2_start)
            n2 = calc_night_minutes(p2_start, p2_end)

        worked = w1 + w2
        night_clock = n1 + n2
        night_ficta = round(night_clock * (60.0 / 52.5))

        intra_sup = 0
        if worked > 360 and break_min < 60:
            intra_sup = 60 - break_min
        elif 240 < worked <= 360 and break_min < 15:
            intra_sup = 15 - break_min

        he50 = 0
        he100 = 0
        if feriado:
            he100 = worked
        else:
            diff = worked - daily_limit_min
            if diff > 10:
                he50 = diff

        tot_he50_min += he50
        tot_he100_min += he100
        tot_intra_min += intra_sup
        tot_night_ficta_min += night_ficta

        rows_out.append({
            "Data": dt_str,
            "Dia": dia_sem,
            "HE 50% (h)": round(he50 / 60.0, 2),
            "HE 100% (h)": round(he100 / 60.0, 2),
            "Intrajornada (h)": round(intra_sup / 60.0, 2),
            "Noturna Ficta (h)": round(night_ficta / 60.0, 2),
            "Status": "Intrajornada suprimida" if intra_sup > 0 else "Apurado",
        })

    return {
        "table": pd.DataFrame(rows_out),
        "he50_hours": round(tot_he50_min / 60.0, 2),
        "he100_hours": round(tot_he100_min / 60.0, 2),
        "intrajornada_hours": round(tot_intra_min / 60.0, 2),
        "night_ficta_hours": round(tot_night_ficta_min / 60.0, 2),
        "incomplete_days": incomplete_days,
    }


# ==============================================================================
# 5. ANÁLISE CRUZADA POR GEMINI (DOCUMENTOS + RELATO LIVRE DO ADVOGADO)
# ==============================================================================
class PaidItem(BaseModel):
    verba: str
    valor: float


class IrregularityFinding(BaseModel):
    categoria: str
    origem_evidencia: str
    descricao: str
    fundamento_juridico: str
    tese_ou_pedido_sugerido: str


class ExtractedCaseData(BaseModel):
    worker_name: Optional[str] = None
    employer_name: Optional[str] = None
    role_title: Optional[str] = None
    admission_date: Optional[str] = None
    dismissal_date: Optional[str] = None
    termination_type: Optional[str] = None
    base_salary: Optional[float] = None
    commissions: Optional[float] = None
    monthly_divisor: Optional[int] = None
    weekly_he50_hours: Optional[float] = None
    monthly_sundays_worked: Optional[int] = None
    suppressed_break_minutes: Optional[int] = None
    weekly_night_hours: Optional[float] = None
    adicional_sugerido: Optional[str] = None
    nr_technical_assessment: Optional[str] = None
    restituicao_atestados_ou_descontos: Optional[float] = None
    fgts_nao_depositado_contrato: Optional[float] = None
    meses_estabilidade_restantes: Optional[int] = None
    sugerir_saldo_salario: bool = True
    sugerir_aviso_previo: bool = True
    sugerir_13_e_ferias: bool = True
    sugerir_fgts_e_multa: bool = True
    sugerir_multa_477: bool = False
    sugerir_multa_467: bool = False
    trct_paid_items: List[PaidItem] = []
    irregularities_found: List[IrregularityFinding] = []
    medical_leave_summary: Optional[str] = None
    stability_alert: Optional[str] = None
    viability_conclusion: Optional[str] = None
    recommended_claims_checklist: List[str] = []
    ambiguities: List[str] = []


def get_gemini_api_key() -> Optional[str]:
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY")


def analyze_case_with_gemini(lawyer_narrative: str, files_payload: List[Dict[str, Any]]) -> Dict[str, Any]:
    api_key = get_gemini_api_key()
    if not api_key or not GENAI_AVAILABLE:
        raise RuntimeError("GEMINI_API_KEY não configurada em Secrets ou google-genai ausente.")

    client = genai.Client(api_key=api_key)
    contents: List[Any] = [
        (
            "Você é um advogado trabalhista sênior, perito técnico em NRs e auditor de cálculos judiciais. "
            "Analise em conjunto o RELATO DO ADVOGADO (entrevista com o cliente) e TODOS OS DOCUMENTOS E FOTOS ANEXADOS "
            "(TRCT, contracheques, cartões de ponto, extratos FGTS, atestados médicos, laudos INSS, ASO, PPP, LTCAT, PGR, CCT e "
            "FOTOS DO AMBIENTE DE TRABALHO / EPIs / EQUIPAMENTOS). "
            "Sua missão:\n"
            "1) Extrair datas, salário-base, comissões/'por fora', jornada habitual, valores já pagos no TRCT, descontos indevidos "
            "(ex.: atestados médicos descontados), saldo de FGTS não depositado no contrato (fgts_nao_depositado_contrato) e "
            "meses de estabilidade provisória violados (meses_estabilidade_restantes).\n"
            "2) AVALIAÇÃO TÉCNICA DE INSALUBRIDADE (NR-15) E PERICULOSIDADE (NR-16): examine fotos do local de trabalho, produtos químicos, "
            "câmara fria, ruído, agentes biológicos, inflamáveis, alta tensão, explosivos ou uso de motocicleta (art. 193 §4º CLT), "
            "bem como PPP/LTCAT/contracheques (se já recebia ou não). Preencha 'adicional_sugerido' com uma das opções exatas: "
            "'Nenhum', 'Insalubridade Grau Mínimo (10%)', 'Insalubridade Grau Médio (20%)', 'Insalubridade Grau Máximo (40%)' ou "
            "'Periculosidade (30% sobre salário)', e detalhe em 'nr_technical_assessment' o anexo da NR-15/NR-16 aplicável, "
            "o que foi observado nas fotos/relato e a necessidade de perícia judicial (Art. 195 CLT).\n"
            "3) Identificar todas as irregularidades, nulidades e brechas jurídicas (estabilidade acidentária B91, gestante, CIPA, "
            "atestados médicos descontados indevidamente, demissão com ASO inapto, rescisão indireta por falta de FGTS ou atraso salarial, "
            "atraso no pagamento das verbas rescisórias > 10 dias [multa art. 477 §8º CLT], verbas rescisórias incontroversas não pagas [multa art. 467 CLT], "
            "horas extras não pagas, intervalo intrajornada suprimido, vínculo sem registro).\n"
            "4) Emitir 'viability_conclusion' (parecer objetivo e fundamentado se há viabilidade para ajuizar ação e quais os pontos fortes) e "
            "'recommended_claims_checklist' (lista objetiva dos pedidos prontos para a Petição Inicial).\n"
            "5) Marcar os booleans sugerir_saldo_salario, sugerir_aviso_previo, sugerir_13_e_ferias, sugerir_fgts_e_multa, "
            "sugerir_multa_477 e sugerir_multa_467 conforme os fatos e documentos.\n"
            "6) Nunca inventar valores ou datas ausentes; sinalize dúvidas em 'ambiguities'."
        )
    ]

    if lawyer_narrative.strip():
        contents.append(f"=== RELATO DO ADVOGADO / FATOS NARRADOS PELO CLIENTE ===\n{lawyer_narrative.strip()}")

    for item in files_payload:
        contents.append(f"=== DOCUMENTO ANEXADO: {item['name']} ===")
        contents.append(types.Part.from_bytes(data=item["bytes"], mime_type=item["mime"]))

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ExtractedCaseData,
            temperature=0.0,
        ),
    )
    return json.loads(response.text)


# ==============================================================================
# 6. MOTOR DE CÁLCULO DECIMAL EXATO (COMPLETO COM REFLEXOS, MULTAS E ESTABILIDADE)
# ==============================================================================
def run_labor_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
    base_salary = d_money(params["base_salary"])
    commissions = d_money(params.get("commissions", 0))
    insalubrity = d_money(params.get("insalubrity", 0))
    remuneration = base_salary + commissions + insalubrity

    divisor = Decimal(str(params["monthly_divisor"]))
    hourly_rate = (remuneration / divisor).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    adm_date: date = params["admission_date"]
    dism_date: date = params["dismissal_date"]
    term_type: str = params["termination_type"]
    years_worked = max(0, (dism_date - adm_date).days // 365)
    meses_prescricao = max(1, min(60, ((dism_date.year - adm_date.year) * 12 + (dism_date.month - adm_date.month))))

    verbas: List[Dict[str, Any]] = []

    def add_verba(
        code: str,
        name: str,
        period: str,
        base: Decimal,
        formula: str,
        qty: str,
        pct_str: str,
        original: Decimal,
        reflexos: Decimal = Decimal("0.00"),
        incide_fgts: bool = True,
    ):
        orig_q = original.quantize(CENT, rounding=ROUND_HALF_UP)
        refl_q = reflexos.quantize(CENT, rounding=ROUND_HALF_UP)
        total_item = orig_q + refl_q
        verbas.append({
            "Código": code,
            "Verba": name,
            "Período": period,
            "Base de Cálculo": base.quantize(CENT),
            "Fórmula": formula,
            "Quantidade": qty,
            "Percentual": pct_str,
            "Valor Original": orig_q,
            "Reflexos": refl_q,
            "Total Apurado": total_item,
            "Incide FGTS": incide_fgts,
        })

    periodo_geral = f"{adm_date.strftime('%d/%m/%Y')} a {dism_date.strftime('%d/%m/%Y')}"

    # A. Verbas Rescisórias Básicas
    if params.get("inc_saldo_salario"):
        dias = Decimal(str(params.get("saldo_dias", dism_date.day)))
        val = (remuneration / Decimal("30")) * dias
        add_verba(
            "VER-01", "Saldo de Salário", dism_date.strftime("%m/%Y"),
            remuneration, "(Remuneração / 30) × Dias Trabalhados",
            f"{dias} dias", "100%", val, incide_fgts=True
        )

    aviso_dias = min(90, 30 + years_worked * 3)
    if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "acordo"):
        fator = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
        val = ((remuneration / Decimal("30")) * Decimal(str(aviso_dias))) * fator
        add_verba(
            "VER-02", "Aviso-Prévio Indenizado (Lei 12.506/2011)", periodo_geral,
            remuneration, "(Remuneração / 30) × Dias Aviso × Fator",
            f"{aviso_dias} dias", f"{int(fator * 100)}%", val, incide_fgts=True
        )

    if params.get("inc_13") and term_type != "justa_causa":
        avos_13 = Decimal(str(params.get("avos_13", dism_date.month)))
    # B. Adicional de Insalubridade ou Periculosidade (todo o período imprescrito)
    if insalubrity > Decimal("0.00") and params.get("inc_adicional_retroativo", True):
        orig_adic = insalubrity * Decimal(str(meses_prescricao))
        anos_fator_ad = Decimal(str(meses_prescricao)) / Decimal("12")
        refl_13_ad = insalubrity * anos_fator_ad
        refl_fer_ad = (insalubrity * anos_fator_ad) * (Decimal("4") / Decimal("3"))
        refl_av_ad = Decimal("0.00")
        if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "acordo"):
            fator_av_ad = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
            refl_av_ad = ((insalubrity / Decimal("30")) * Decimal(str(aviso_dias))) * fator_av_ad
        refl_tot_ad = refl_13_ad + refl_fer_ad + refl_av_ad
        add_verba(
            "VER-04B", f"Adicional de {params.get('adicional_label', 'Insalubridade/Periculosidade')} + Reflexos", periodo_geral,
            insalubrity, "Adicional Mensal × Meses + Reflexos (13º, Férias + 1/3 e Aviso)",
            f"{meses_prescricao} meses", params.get("adicional_pct_label", "NR-15/16"), orig_adic, reflexos=refl_tot_ad, incide_fgts=True
        )

    # C. Jornada de Trabalho e Adicionais
    dsr_factor = Decimal("1") / Decimal("6")
    parcelas_habituais_total = Decimal("0.00")

    if params.get("inc_he50") and d_money(params.get("he50_hours", 0)) > 0:
        h50 = Decimal(str(params["he50_hours"]))
        pct_he = Decimal(str(params.get("he50_pct", 50))) / Decimal("100")
        unit_he = hourly_rate * (Decimal("1") + pct_he)
        orig_he = unit_he * h50
        refl_dsr = orig_he * dsr_factor
        parcelas_habituais_total += orig_he
        add_verba(
            "VER-05", f"Horas Extras ({params.get('he50_pct', 50)}%) + Reflexo DSR", periodo_geral,
            hourly_rate.quantize(CENT), f"(Hora Normal × (1 + {params.get('he50_pct', 50)}%)) × Horas + DSR (1/6)",
            f"{h50}h", f"+{params.get('he50_pct', 50)}%", orig_he, reflexos=refl_dsr, incide_fgts=True
        )

    if params.get("inc_he100") and d_money(params.get("he100_hours", 0)) > 0:
        h100 = Decimal(str(params["he100_hours"]))
        unit_100 = hourly_rate * Decimal("2")
        orig_100 = unit_100 * h100
        refl_dsr_100 = orig_100 * dsr_factor
        parcelas_habituais_total += orig_100
        add_verba(
            "VER-05B", "Horas Extras Domingos/Feriados (100%) + Reflexo DSR", periodo_geral,
            hourly_rate.quantize(CENT), "(Hora Normal × 2,00) × Horas 100% + DSR (1/6)",
            f"{h100}h", "+100%", orig_100, reflexos=refl_dsr_100, incide_fgts=True
        )

    if params.get("inc_intrajornada") and d_money(params.get("intrajornada_hours", 0)) > 0:
        h_intra = Decimal(str(params["intrajornada_hours"]))
        unit_intra = hourly_rate * Decimal("1.5")
        orig_intra = unit_intra * h_intra
        add_verba(
            "VER-05C", "Intervalo Intrajornada Suprimido (+50% — Art. 71 §4º CLT)", periodo_geral,
            hourly_rate.quantize(CENT), "(Hora Normal × 1,50) × Horas Suprimidas",
            f"{h_intra}h", "+50%", orig_intra, incide_fgts=False
        )

    if params.get("inc_noturno") and d_money(params.get("night_hours", 0)) > 0:
        h_not = Decimal(str(params["night_hours"]))
        pct_not = Decimal(str(params.get("night_pct", 20))) / Decimal("100")
        orig_not = (hourly_rate * pct_not) * h_not
        refl_dsr_not = orig_not * dsr_factor
        parcelas_habituais_total += orig_not
        add_verba(
            "VER-06", "Adicional Noturno (Redução Ficta 52m30s) + Reflexo DSR", periodo_geral,
            hourly_rate.quantize(CENT), f"(Hora Normal × {params.get('night_pct', 20)}%) × Horas Fictas + DSR",
            f"{h_not}h", f"+{params.get('night_pct', 20)}%", orig_not, reflexos=refl_dsr_not, incide_fgts=True
        )

    # C. Reflexos das Parcelas Habituais (HE + Adic. Noturno) em 13º, Férias + 1/3 e Aviso-Prévio
    if params.get("inc_reflexos_completos") and parcelas_habituais_total > Decimal("0.00"):
        media_mensal_habitual = (parcelas_habituais_total / Decimal(str(meses_prescricao))).quantize(CENT, rounding=ROUND_HALF_UP)
        anos_fator = Decimal(str(meses_prescricao)) / Decimal("12")
        refl_13 = media_mensal_habitual * anos_fator
        refl_ferias = (media_mensal_habitual * anos_fator) * (Decimal("4") / Decimal("3"))
        refl_aviso = Decimal("0.00")
        if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "acordo"):
            fator_av = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
            refl_aviso = ((media_mensal_habitual / Decimal("30")) * Decimal(str(aviso_dias))) * fator_av
        total_refl_cascata = refl_13 + refl_ferias + refl_aviso
        add_verba(
            "VER-06B", "Reflexos Habituais (HE/Noturno) em 13º, Férias + 1/3 e Aviso-Prévio", periodo_geral,
            media_mensal_habitual, "Média Mensal Habitual × (13º + Férias 4/3 + Aviso Prévio)",
            f"{meses_prescricao} meses", "Súm. 347 TST", total_refl_cascata, incide_fgts=True
        )

    # D. Indenização de Estabilidade Provisória (B91 / Gestante / CIPA)
    meses_estab = int(params.get("meses_estabilidade", 0))
    if meses_estab > 0:
        ind_estab = remuneration * Decimal(str(meses_estab))
        add_verba(
            "VER-07", "Indenização Substitutiva do Período de Estabilidade Provisória", "Pós-demissão",
            remuneration, "Remuneração Mensal × Meses de Estabilidade Restantes",
            f"{meses_estab} meses", "100%", ind_estab, incide_fgts=True
        )

    # E. FGTS (8%) sobre Verbas + FGTS Não Depositado no Contrato + Multa Rescisória
    multa_pct = Decimal("0.40") if term_type in ("sem_justa_causa", "rescisao_indireta") else (
        Decimal("0.20") if term_type == "acordo" else Decimal("0.00")
    )
    fgts_nao_dep = d_money(params.get("fgts_nao_depositado", 0))
    if params.get("inc_fgts") or fgts_nao_dep > Decimal("0.00"):
        base_fgts_verbas = sum((v["Total Apurado"] for v in verbas if v["Incide FGTS"]), Decimal("0.00"))
        fgts_8_verbas = (base_fgts_verbas * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP) if params.get("inc_fgts") else Decimal("0.00")
        base_multa = fgts_8_verbas + fgts_nao_dep
        multa_val = (base_multa * multa_pct).quantize(CENT, rounding=ROUND_HALF_UP)
        total_fgts_item = fgts_8_verbas + fgts_nao_dep + multa_val
        if total_fgts_item > Decimal("0.00"):
            add_verba(
                "VER-08", "FGTS (8% Verbas + Depósitos Faltantes) + Multa Rescisória", periodo_geral,
                base_fgts_verbas + fgts_nao_dep, "(Verbas Incidentes × 8% + FGTS Faltante) + Multa",
                "Contrato + Rescisão", f"8% + {int(multa_pct * 100)}%", total_fgts_item, incide_fgts=False
            )

    # F. Multas dos Arts. 477 § 8º e 467 da CLT
    if params.get("inc_multa_477"):
        add_verba(
            "VER-09", "Multa por Atraso na Quitação Rescisória (Art. 477, § 8º CLT)", dism_date.strftime("%m/%Y"),
            remuneration, "1 Remuneração Mensal do Empregado",
            "1 salário", "100%", remuneration, incide_fgts=False
        )

    if params.get("inc_multa_467"):
        # 50% sobre as verbas rescisórias estritas incontroversas (Saldo + Aviso + 13º + Férias)
        base_467 = sum(
            (v["Total Apurado"] for v in verbas if v["Código"] in ("VER-01", "VER-02", "VER-03", "VER-04")),
            Decimal("0.00"),
        )
        abat_prev = d_money(params.get("offsets", 0))
        saldo_incontroverso = max(Decimal("0.00"), base_467 - abat_prev)
        multa_467_val = (saldo_incontroverso * Decimal("0.50")).quantize(CENT, rounding=ROUND_HALF_UP)
        if multa_467_val > Decimal("0.00"):
            add_verba(
                "VER-10", "Multa sobre Verbas Rescisórias Incontroversas (Art. 467 CLT)", dism_date.strftime("%m/%Y"),
                saldo_incontroverso, "Verbas Rescisórias Líquidas × 50%",
                "1 parcela", "50%", multa_467_val, incide_fgts=False
            )

    bruto_atualizado = sum((v["Total Apurado"] for v in verbas), Decimal("0.00"))
    abatimentos = d_money(params.get("offsets", 0))
    liquido_devido = max(Decimal("0.00"), bruto_atualizado - abatimentos)

    audit_payload = json.dumps(
        {"process": params["process_number"], "bruto": str(bruto_atualizado), "liquido": str(liquido_devido), "tables": TABLES_VERSION},
        sort_keys=True,
    )
    audit_hash = hashlib.sha256(audit_payload.encode()).hexdigest()[:16].upper()

    return {
        "audit_id": f"ED-{audit_hash}",
        "worker_name": params.get("worker_name", ""),
        "employer_name": params.get("employer_name", ""),
        "role_title": params.get("role_title", ""),
        "admission_date": adm_date.strftime("%d/%m/%Y"),
        "dismissal_date": dism_date.strftime("%d/%m/%Y"),
        "termination_type": term_type,
        "origin": "Motor próprio decimal exato (Python Decimal) — Cálculo pré-processual (Art. 840, § 1º CLT)",
        "tables_version": TABLES_VERSION,
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "verbas": verbas,
        "bruto_atualizado": bruto_atualizado,
        "abatimentos": abatimentos,
        "liquido_devido": liquido_devido,
        "ai_opinion": params.get("ai_opinion") or {},
        "lawyer_narrative": params.get("lawyer_narrative", ""),
    }


def build_printable_dossier_pdf_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera arquivo PDF binário real (via fpdf2) com Parecer, Irregularidades, Memória e Pedidos Liquidados."""
    def latin(txt: Any) -> str:
        return str(txt or "").encode("latin-1", errors="replace").decode("latin-1")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_fill_color(11, 27, 43)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 11, latin("ED Calculos Trabalhistas - Parecer e Avaliacao Preliminar"), new_x="LMARGIN", new_y="NEXT", fill=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(
        0, 7,
        latin(f"Auditoria: {calc['audit_id']} | Gerado em: {calc['generated_at']} | Tabelas: {calc['tables_version']}"),
        new_x="LMARGIN", new_y="NEXT", fill=True,
    )
    pdf.ln(4)

    pdf.set_text_color(15, 37, 55)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, latin("1. Identificacao das Partes e Contrato"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, latin(f"Reclamante (Cliente): {calc.get('worker_name', '')}"), new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, latin(f"Reclamada (Empresa): {calc.get('employer_name', '')}"), new_x="LMARGIN", new_y="NEXT")
    pdf.cell(
        0, 6,
        latin(f"Cargo: {calc.get('role_title', '')} | Periodo: {calc.get('admission_date', '')} a {calc.get('dismissal_date', '')}"),
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.ln(3)

    ai_op = calc.get("ai_opinion") or {}
    if ai_op.get("viability_conclusion") or ai_op.get("stability_alert") or ai_op.get("medical_leave_summary") or ai_op.get("nr_technical_assessment"):
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, latin("2. Parecer Preliminar de Viabilidade, NRs e Brechas"), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9.5)
        if ai_op.get("stability_alert"):
            pdf.multi_cell(0, 5.5, latin(f"ALERTA DE ESTABILIDADE: {ai_op['stability_alert']}"))
        if ai_op.get("nr_technical_assessment"):
            pdf.multi_cell(0, 5.5, latin(f"INSALUBRIDADE / PERICULOSIDADE (NR-15 / NR-16): {ai_op['nr_technical_assessment']}"))
        if ai_op.get("medical_leave_summary"):
            pdf.multi_cell(0, 5.5, latin(f"ATESTADOS / AFASTAMENTOS / ASO: {ai_op['medical_leave_summary']}"))
        if ai_op.get("viability_conclusion"):
            pdf.multi_cell(0, 5.5, latin(f"CONCLUSAO: {ai_op['viability_conclusion']}"))
        pdf.ln(3)

    irregs = ai_op.get("irregularities_found") or []
    if irregs:
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.cell(0, 7, latin("Irregularidades Identificadas nos Documentos / Relato:"), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        for idx_i, irr in enumerate(irregs, 1):
            linha_irr = f"{idx_i}. [{irr.get('categoria', '')}] {irr.get('descricao', '')} ({irr.get('fundamento_juridico', '')})"
            pdf.multi_cell(0, 5, latin(linha_irr))
        pdf.ln(3)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, latin("3. Memoria de Calculo Auditavel - Pedidos Liquidados"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_fill_color(241, 245, 249)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(18, 7, "Cod.", border=1, fill=True)
    pdf.cell(82, 7, "Verba / Parcela", border=1, fill=True)
    pdf.cell(28, 7, "Qtd / Perc", border=1, fill=True)
    pdf.cell(31, 7, "Base", border=1, fill=True, align="R")
    pdf.cell(31, 7, "Total", border=1, fill=True, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "", 8.5)
    for v in calc["verbas"]:
        pdf.cell(18, 6.5, latin(v["Código"]), border=1)
        pdf.cell(82, 6.5, latin(v["Verba"][:44]), border=1)
        pdf.cell(28, 6.5, latin(f"{v['Quantidade']} ({v['Percentual']})"), border=1)
        pdf.cell(31, 6.5, latin(fmt_brl(v["Base de Cálculo"])), border=1, align="R")
        pdf.cell(31, 6.5, latin(fmt_brl(v["Total Apurado"])), border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, latin(f"Total Apurado das Verbas: {fmt_brl(calc['bruto_atualizado'])}"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, latin(f"(-) Valores Ja Pagos no TRCT: -{fmt_brl(calc['abatimentos'])}"), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(21, 128, 61)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, latin(f"VALOR DA CAUSA / RESTITUICAO A PEDIR: {fmt_brl(calc['liquido_devido'])}"), align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(4)
    pdf.set_text_color(100, 116, 139)
    pdf.set_font("Helvetica", "I", 8)
    pdf.multi_cell(0, 4.5, latin(f"Origem: {calc['origin']}"))

    return bytes(pdf.output())


def build_complete_excel_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera planilha Excel multi-abas com Resumo, Verbas Liquidadas e Parecer/Irregularidades."""
    buf = io.BytesIO()
    df_verbas = pd.DataFrame(calc["verbas"])
    for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"]:
        if col in df_verbas.columns:
            df_verbas[col] = df_verbas[col].apply(fmt_brl)

    df_resumo = pd.DataFrame([
        {"Campo": "Código de Auditoria", "Valor": calc["audit_id"]},
        {"Campo": "Data/Hora Geração", "Valor": calc["generated_at"]},
        {"Campo": "Reclamante (Cliente)", "Valor": calc.get("worker_name", "")},
        {"Campo": "Reclamada (Empresa)", "Valor": calc.get("employer_name", "")},
        {"Campo": "Período Contratual", "Valor": f"{calc.get('admission_date','')} a {calc.get('dismissal_date','')}"},
        {"Campo": "Total Apurado das Verbas", "Valor": fmt_brl(calc["bruto_atualizado"])},
        {"Campo": "(-) Total Já Pago no TRCT", "Valor": fmt_brl(calc["abatimentos"])},
        {"Campo": "Valor da Causa / Restituição", "Valor": fmt_brl(calc["liquido_devido"])},
    ])

    ai_op = calc.get("ai_opinion") or {}
    irregs = ai_op.get("irregularities_found") or []
    df_irregs = pd.DataFrame(irregs) if irregs else pd.DataFrame([{"Aviso": "Nenhuma irregularidade estruturada registrada."}])

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, index=False, sheet_name="Resumo_Dossie")
        df_verbas.to_excel(writer, index=False, sheet_name="Pedidos_Liquidados")
        df_irregs.to_excel(writer, index=False, sheet_name="Parecer_Irregularidades")
    return buf.getvalue()


# ==============================================================================
# 7. MENU LATERAL
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas")
    st.caption("Parecer Pré-Processual & Auditoria")
    st.divider()
    menu = st.radio(
        "Navegação",
        [
            "Avaliação Preliminar & Auditoria",
            "Meus Cálculos e Relatórios",
        ],
    )
    st.divider()
    st.caption(f"Tabelas ativas: `{TABLES_VERSION}`")


# ==============================================================================
# 8. TELAS DO SISTEMA
# ==============================================================================
if menu == "Avaliação Preliminar & Auditoria":
    st.title("Avaliação Preliminar, Parecer IA & Cálculo para Ajuizamento")
    st.markdown(
        "Relate os fatos narrados pelo cliente e/ou anexe os documentos disponíveis (TRCT, contracheques, ponto, atestados, laudos INSS, FGTS, CCT). "
        "A IA cruza tudo para emitir o **parecer de viabilidade e brechas**, enquanto o **motor decimal** entrega os pedidos liquidados prontos para a Petição Inicial."
    )

    with st.container(border=True):
        st.subheader("1. Relato do Caso (Entrevista) + Fotos do Ambiente + Upload Universal")
        lawyer_narrative = st.text_area(
            "Relato livre do advogado / condições do ambiente de trabalho e fatos narrados pelo cliente:",
            placeholder=(
                "Ex.: Cliente admitido em 01/02/2024 como operador/motoboy com salário registrado de R$ 2.400,00 + R$ 600,00 por fora. "
                "Trabalhava em contato diário com produtos químicos/graxa sem EPI adequado ou em câmara fria/motocicleta (ver fotos anexas), "
                "sem receber insalubridade ou periculosidade. Fazia 6h extras semanais, tinha só 30 min de almoço e teve 2 atestados descontados..."
            ),
            height=135,
            key="lawyer_narrative_box",
        )

        uploaded_all = st.file_uploader(
            "Anexe documentos E FOTOS do ambiente de trabalho / EPIs / produtos (PDF, Fotos JPG/PNG, Excel .xlsx, CSV ou TXT)",
            type=["pdf", "png", "jpg", "jpeg", "csv", "xlsx", "txt"],
            accept_multiple_files=True,
            key="universal_dossier_uploader",
        )

        if uploaded_all:
            st.caption(f"**{len(uploaded_all)} arquivo(s) anexado(s):** " + ", ".join(f"`{f.name}`" for f in uploaded_all))

        if st.button("Analisar Relato + Documentos e Preparar Ação (IA + Motor Decimal)", type="primary"):
            if not lawyer_narrative.strip() and not uploaded_all:
                st.warning("Escreva um breve relato do caso ou anexe pelo menos um documento para analisar.")
            else:
                gemini_files = []
                timecard_dfs = []
                for f in (uploaded_all or []):
                    fname_low = f.name.lower()
                    if fname_low.endswith(".csv"):
                        raw_txt = f.getvalue().decode("utf-8", errors="ignore")
                        sep = ";" if ";" in raw_txt.splitlines()[0] else ","
                        df_tmp = pd.read_csv(io.StringIO(raw_txt), sep=sep)
                        if {"Entrada1", "Saida1"}.issubset(set(df_tmp.columns)):
                            timecard_dfs.append(df_tmp)
                        else:
                            gemini_files.append({"name": f.name, "bytes": raw_txt.encode("utf-8"), "mime": "text/plain"})
                    elif fname_low.endswith(".xlsx"):
                        df_tmp = pd.read_excel(f)
                        if {"Entrada1", "Saida1"}.issubset(set(df_tmp.columns)):
                            timecard_dfs.append(df_tmp)
                        else:
                            csv_conv = df_tmp.to_csv(index=False).encode("utf-8")
                            gemini_files.append({"name": f.name, "bytes": csv_conv, "mime": "text/plain"})
                    else:
                        mime_map = {
                            ".pdf": "application/pdf",
                            ".png": "image/png",
                            ".jpg": "image/jpeg",
                            ".jpeg": "image/jpeg",
                            ".txt": "text/plain",
                        }
                        ext = "." + fname_low.rsplit(".", 1)[-1]
                        gemini_files.append({
                            "name": f.name,
                            "bytes": f.getvalue(),
                            "mime": mime_map.get(ext, "application/pdf"),
                        })

                if timecard_dfs:
                    df_concat = pd.concat(timecard_dfs, ignore_index=True)
                    st.session_state.timecard_result = process_timecard_dataframe(df_concat)
                    st.success("Cartão(ões) de ponto apurado(s) matematicamente em minutos!")

                if lawyer_narrative.strip() or gemini_files:
                    try:
                        with st.spinner("Cruzando relato jurídico e documentos para gerar parecer e teses..."):
                            st.session_state.extracted_ai_data = analyze_case_with_gemini(lawyer_narrative, gemini_files)
                            st.session_state.extraction_confirmed = False
                        st.success("Parecer e extração concluídos! Revise as teses e parâmetros abaixo.")
                    except Exception as exc:
                        st.warning(f"Análise via Gemini indisponível ({exc}). Verifique `GEMINI_API_KEY` em Secrets.")

    tc_pre = st.session_state.timecard_result
    ext_pre = st.session_state.extracted_ai_data or {}

    if ext_pre:
        with st.container(border=True):
            st.subheader("Parecer Jurídico Preliminar & Estratégia para Ajuizamento")
            if ext_pre.get("stability_alert"):
                st.error(f"**ALERTA DE ESTABILIDADE / REINTEGRAÇÃO OU INDENIZAÇÃO:** {ext_pre['stability_alert']}")
            if ext_pre.get("nr_technical_assessment"):
                st.warning(f"**Enquadramento Técnico NR-15 / NR-16 (Fotos & Relato):** {ext_pre['nr_technical_assessment']}")
            if ext_pre.get("medical_leave_summary"):
                st.warning(f"**Atestados Médicos / Afastamentos / ASO:** {ext_pre['medical_leave_summary']}")
            if ext_pre.get("viability_conclusion"):
                st.info(f"**Conclusão de Viabilidade da Ação:** {ext_pre['viability_conclusion']}")

            c_left, c_right = st.columns([3, 2])
            with c_left:
                irregs = ext_pre.get("irregularities_found") or []
                if irregs:
                    st.markdown("#### Irregularidades e Brechas Identificadas")
                    st.dataframe(pd.DataFrame(irregs), use_container_width=True, hide_index=True)
            with c_right:
                checklist = ext_pre.get("recommended_claims_checklist") or []
                if checklist:
                    st.markdown("#### Pedidos Recomendados para a Inicial")
                    for item in checklist:
                        st.markdown(f"- ✅ {item}")

            if ext_pre.get("ambiguities"):
                st.caption("**Pontos pendentes de confirmação pelo advogado:** " + " | ".join(ext_pre["ambiguities"]))

            st.session_state.extraction_confirmed = st.checkbox(
                "Confirmo que revisei o parecer e os parâmetros extraídos antes de fechar os valores",
                value=st.session_state.extraction_confirmed,
            )

    if tc_pre is not None:
        with st.expander("Ver Espelho de Ponto Apurado Dia a Dia (Tolerância CLT + Redução Ficta)"):
            st.dataframe(tc_pre["table"], use_container_width=True, hide_index=True)

    trct_paid_list = ext_pre.get("trct_paid_items") or []
    trct_auto_sum = float(round(sum(float(i.get("valor", 0.0)) for i in trct_paid_list), 2))

    def parse_iso_date(val: Optional[str], fallback: date) -> date:
        if not val:
            return fallback
        try:
            return datetime.strptime(val[:10], "%Y-%m-%d").date()
        except Exception:
            return fallback

    with st.form("preliminary_eval_form"):
        st.subheader("2. Conferência das Condições do Contrato e Pedidos")
        if ext_pre:
            st.success("Campos abaixo preenchidos automaticamente com base no seu relato e nos documentos anexados. Ajuste apenas se desejar alterar algo.")
        else:
            st.caption("Preenchimento inicial de exemplo — clique em 'Analisar Relato + Documentos' acima para preencher automaticamente ou edite diretamente abaixo.")

        p1, p2, p3 = st.columns(3)
        cliente_nome = p1.text_input("Nome do Cliente (Reclamante) *", value=ext_pre.get("worker_name") or "João da Silva")
        empresa_nome = p2.text_input("Empresa (Reclamada) *", value=ext_pre.get("employer_name") or "Comércio e Serviços Alfa Ltda.")
        cargo_funcao = p3.text_input("Cargo / Função", value=ext_pre.get("role_title") or "Auxiliar Operacional / Vendedor")

        motivos_opts = ["sem_justa_causa", "rescisao_indireta", "pedido_demissao", "acordo", "justa_causa"]
        motivo_ext = ext_pre.get("termination_type") if ext_pre.get("termination_type") in motivos_opts else "sem_justa_causa"

        d1, d2, d3, d4 = st.columns(4)
        adm_pre = d1.date_input("Data de Admissão *", value=parse_iso_date(ext_pre.get("admission_date"), date(2024, 2, 1)), key="adm_pre")
        dism_pre = d2.date_input("Data de Saída (ou Hoje) *", value=parse_iso_date(ext_pre.get("dismissal_date"), date(2026, 11, 20)), key="dism_pre")
        motivo_saida = d3.selectbox(
            "Modalidade / Tese da Saída *",
            motivos_opts,
            index=motivos_opts.index(motivo_ext),
            format_func=lambda x: {
                "sem_justa_causa": "Dispensa sem justa causa",
                "rescisao_indireta": "Rescisão Indireta (falta grave do empregador)",
                "pedido_demissao": "Pedido de demissão (verbas pendentes)",
                "acordo": "Rescisão por acordo (art. 484-A CLT)",
                "justa_causa": "Reversão de Justa Causa",
            }[x],
        )
        divisor_pre = d4.number_input("Divisor Mensal *", min_value=1, max_value=300, value=int(ext_pre.get("monthly_divisor") or 220))

        adicionais_opts = [
            "Nenhum",
            "Insalubridade Grau Mínimo (10%)",
            "Insalubridade Grau Médio (20%)",
            "Insalubridade Grau Máximo (40%)",
            "Periculosidade (30% sobre salário)",
        ]
        adic_ext = ext_pre.get("adicional_sugerido") if ext_pre.get("adicional_sugerido") in adicionais_opts else "Nenhum"

        r1, r2, r3, r4 = st.columns(4)
        sal_registrado = r1.number_input("Salário-Base Registrado (R$) *", min_value=0.0, value=float(ext_pre.get("base_salary") or 2400.00), step=50.0)
        sal_por_fora = r2.number_input("Comissões / Salário 'Por Fora' (R$/mês)", min_value=0.0, value=float(ext_pre.get("commissions") or 0.0), step=50.0)
        adicional_cond = r3.selectbox("Adicional Devido", adicionais_opts, index=adicionais_opts.index(adic_ext))
        valor_ja_pago_trct = r4.number_input("Total Já Pago no TRCT (R$)", min_value=0.0, value=float(trct_auto_sum), step=100.0)

        meses_contrato = max(1, ((dism_pre.year - adm_pre.year) * 12 + (dism_pre.month - adm_pre.month)))
        j1, j2, j3, j4 = st.columns(4)
        he_semanal_relatada = j1.number_input("Horas Extras / Semana (50%)", min_value=0.0, value=float(ext_pre.get("weekly_he50_hours") or 0.0), step=0.5)
        domingos_mes = j2.number_input("Domingos/Feriados sem Folga / Mês", min_value=0, max_value=10, value=int(ext_pre.get("monthly_sundays_worked") or 0))
        min_intervalo_suprimido = j3.number_input("Intervalo Suprimido (min/dia)", min_value=0, max_value=60, value=int(ext_pre.get("suppressed_break_minutes") or 0))
        horas_noturnas_semana = j4.number_input("Horas Noturnas / Semana", min_value=0.0, value=float(ext_pre.get("weekly_night_hours") or 0.0), step=0.5)

        st.subheader("3. Seleção de Pedidos, Reflexos e Multas da Inicial")
        v1, v2, v3 = st.columns(3)
        ped_saldo = v1.checkbox("Saldo de Salário", value=bool(ext_pre.get("sugerir_saldo_salario", True)))
        dias_saldo_pre = v1.number_input("Dias trabalhados no último mês", min_value=0, max_value=31, value=dism_pre.day)
        ped_reflexos_completos = v1.checkbox("Reflexos Habituais em 13º, Férias + 1/3 e Aviso", value=True)

        ped_aviso = v2.checkbox("Aviso-Prévio Indenizado + Projeção", value=bool(ext_pre.get("sugerir_aviso_previo", True)))
        ped_13_ferias = v2.checkbox("13º Proporcional e Férias + 1/3", value=bool(ext_pre.get("sugerir_13_e_ferias", True)))
        ped_multa_477 = v2.checkbox("Multa Art. 477, § 8º CLT (Atraso TRCT)", value=bool(ext_pre.get("sugerir_multa_477", False)))
        ped_multa_467 = v2.checkbox("Multa Art. 467 CLT (50% Verbas Incontroversas)", value=bool(ext_pre.get("sugerir_multa_467", False)))

        ped_fgts = v3.checkbox("FGTS (8%) + Multa Rescisória", value=bool(ext_pre.get("sugerir_fgts_e_multa", True)))
        fgts_faltante_val = v3.number_input(
            "FGTS Não Depositado no Contrato (R$)",
            min_value=0.0,
            value=float(ext_pre.get("fgts_nao_depositado_contrato") or 0.0),
            step=100.0,
        )
        meses_estab_val = v3.number_input(
            "Meses de Estabilidade Violada (B91/Gestante/CIPA)",
            min_value=0,
            max_value=24,
            value=int(ext_pre.get("meses_estabilidade_restantes") or 0),
        )
        rest_auto = float(ext_pre.get("restituicao_atestados_ou_descontos") or 0.0)
        restituicao_atestados = v3.number_input(
            "Restituição de Atestados / Descontos Indevidos (R$)",
            min_value=0.0,
            value=rest_auto,
            step=50.0,
            help="Preenchido automaticamente se constar desconto indevido no TRCT/contracheque/atestado, ou informe manualmente o total em R$ a restituir.",
        )

        avaliar_btn = st.form_submit_button("Gerar Cálculos e Quadro de Pedidos para Ajuizamento", type="primary")

    if avaliar_btn:
        if ext_pre and not st.session_state.extraction_confirmed:
            st.error("Marque a caixa de confirmação da revisão antes de fechar os valores.")
        elif sal_registrado <= 0 or not cliente_nome.strip():
            st.error("Preencha o nome do cliente e o salário-base registrado.")
        elif dism_pre < adm_pre:
            st.error("A data de saída não pode ser anterior à admissão.")
        else:
            meses_prescricao = min(60, meses_contrato)
            semanas_periodo = Decimal(str(meses_prescricao)) * Decimal("4.3333")

            if tc_pre is not None:
                total_he50 = float(tc_pre["he50_hours"])
                total_he100 = float(tc_pre["he100_hours"])
                total_intra = float(tc_pre["intrajornada_hours"])
                total_not = float(tc_pre["night_ficta_hours"])
            else:
                total_he50 = float((Decimal(str(he_semanal_relatada)) * semanas_periodo).quantize(Decimal("0.01")))
                total_he100 = float((Decimal(str(domingos_mes * 8)) * Decimal(str(meses_prescricao))).quantize(Decimal("0.01")))
                dias_uteis_periodo = Decimal(str(meses_prescricao)) * Decimal("22")
                total_intra = float(((Decimal(str(min_intervalo_suprimido)) / Decimal("60")) * dias_uteis_periodo).quantize(Decimal("0.01")))
                total_not = float((Decimal(str(horas_noturnas_semana)) * semanas_periodo * Decimal("1.142857")).quantize(Decimal("0.01")))

            sm_vigente = get_historical_min_wage(dism_pre)
            insal_mensal = Decimal("0.00")
            adic_pct_lbl = "0%"
            if "10%" in adicional_cond:
                insal_mensal = (sm_vigente * Decimal("0.10")).quantize(CENT)
                adic_pct_lbl = "10% SM (NR-15)"
            elif "20%" in adicional_cond:
                insal_mensal = (sm_vigente * Decimal("0.20")).quantize(CENT)
                adic_pct_lbl = "20% SM (NR-15)"
            elif "40%" in adicional_cond:
                insal_mensal = (sm_vigente * Decimal("0.40")).quantize(CENT)
                adic_pct_lbl = "40% SM (NR-15)"
            elif "30%" in adicional_cond:
                insal_mensal = (d_money(sal_registrado) * Decimal("0.30")).quantize(CENT)
                adic_pct_lbl = "30% Sal. (NR-16)"

            res_pre = run_labor_calculation({
                "process_number": f"INICIAL-{cliente_nome.upper().replace(' ', '_')[:18]}",
                "worker_name": cliente_nome,
                "employer_name": empresa_nome,
                "role_title": cargo_funcao,
                "admission_date": adm_pre,
                "dismissal_date": dism_pre,
                "termination_type": motivo_saida,
                "monthly_divisor": divisor_pre,
                "base_salary": sal_registrado,
                "commissions": sal_por_fora,
                "insalubrity": float(insal_mensal),
                "adicional_label": adicional_cond,
                "adicional_pct_label": adic_pct_lbl,
                "inc_adicional_retroativo": insal_mensal > Decimal("0.00"),
                "offsets": valor_ja_pago_trct,
                "inc_saldo_salario": ped_saldo,
                "saldo_dias": dias_saldo_pre,
                "inc_aviso_previo": ped_aviso,
                "inc_13": ped_13_ferias,
                "inc_ferias": ped_13_ferias,
                "inc_fgts": ped_fgts,
                "fgts_nao_depositado": fgts_faltante_val,
                "meses_estabilidade": meses_estab_val,
                "inc_multa_477": ped_multa_477,
                "inc_multa_467": ped_multa_467,
                "inc_reflexos_completos": ped_reflexos_completos,
                "inc_he50": total_he50 > 0,
                "he50_hours": total_he50,
                "he50_pct": 50,
                "inc_he100": total_he100 > 0,
                "he100_hours": total_he100,
                "inc_intrajornada": total_intra > 0,
                "intrajornada_hours": total_intra,
                "inc_noturno": total_not > 0,
                "night_hours": total_not,
                "night_pct": 20,
                "ai_opinion": ext_pre,
                "lawyer_narrative": lawyer_narrative,
            })

            if restituicao_atestados > 0:
                rest_dec = d_money(restituicao_atestados)
                res_pre["verbas"].append({
                    "Código": "VER-11",
                    "Verba": "Restituição de Descontos Indevidos / Atestados Médicos",
                    "Período": f"{adm_pre.strftime('%d/%m/%Y')} a {dism_pre.strftime('%d/%m/%Y')}",
                    "Base de Cálculo": rest_dec,
                    "Fórmula": "Devolução integral de descontos indevidos (Art. 462 CLT)",
                    "Quantidade": "1",
                    "Percentual": "100%",
                    "Valor Original": rest_dec,
                    "Reflexos": Decimal("0.00"),
                    "Total Apurado": rest_dec,
                    "Incide FGTS": False,
                })
                res_pre["bruto_atualizado"] += rest_dec
                res_pre["liquido_devido"] += rest_dec

            st.session_state.calculations_history.insert(0, res_pre)
            st.success(f"Cálculo preliminar pronto para ajuizamento! Código: **{res_pre['audit_id']}**")

            k1, k2, k3 = st.columns(3)
            k1.metric("Total Apurado das Verbas", fmt_brl(res_pre["bruto_atualizado"]))
            k2.metric("(-) Já Pago no TRCT", fmt_brl(res_pre["abatimentos"]))
            k3.metric("Valor da Causa / Restituição a Pedir", fmt_brl(res_pre["liquido_devido"]))

            df_pre = pd.DataFrame(res_pre["verbas"])
            for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"]:
                df_pre[col] = df_pre[col].apply(fmt_brl)
            st.dataframe(
                df_pre[["Código", "Verba", "Quantidade", "Percentual", "Base de Cálculo", "Fórmula", "Valor Original", "Reflexos", "Total Apurado"]],
                use_container_width=True,
                hide_index=True,
            )

            linhas_pedidos = [
                f"DOS PEDIDOS LIQUIDADOS (Art. 840, § 1º da CLT) — {cliente_nome} x {empresa_nome}:",
                "",
            ]
            for v in res_pre["verbas"]:
                linhas_pedidos.append(f"• {v['Verba']}: {fmt_brl(v['Total Apurado'])}")
            if res_pre["abatimentos"] > Decimal("0.00"):
                linhas_pedidos.append(f"• (-) Abatimento dos valores comprovadamente pagos no TRCT: -{fmt_brl(res_pre['abatimentos'])}")
            linhas_pedidos.append(f"\nVALOR TOTAL ATRIBUÍDO À CAUSA: {fmt_brl(res_pre['liquido_devido'])}")

            with st.expander("Copiar Texto dos Pedidos Liquidados para a Petição Inicial (Art. 840, § 1º CLT)", expanded=True):
                st.text_area("Minuta pronta para colar na inicial:", value="\n".join(linhas_pedidos), height=220)

            dl1, dl2 = st.columns(2)
            dl1.download_button(
                "Baixar Dossiê Completo em Excel (.XLSX — Resumo + Verbas + Parecer)",
                data=build_complete_excel_bytes(res_pre),
                file_name=f"Dossie_Trabalhista_{res_pre['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            dl2.download_button(
                "Baixar Relatório Oficial em PDF (.PDF)",
                data=build_printable_dossier_pdf_bytes(res_pre),
                file_name=f"Relatorio_Dossie_{res_pre['audit_id']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

elif menu == "Meus Cálculos e Relatórios":
    st.title("Meus Cálculos e Relatórios Salvos")
    st.caption("Baixe o dossiê completo em Excel (.XLSX) ou diretamente em PDF (.PDF) para arquivar no processo.")
    if not st.session_state.calculations_history:
        st.info("Nenhuma avaliação realizada nesta sessão ainda.")
    else:
        for idx, calc in enumerate(st.session_state.calculations_history):
            titulo_exp = (
                f"{calc['audit_id']} • {calc.get('worker_name', 'Cliente')} x {calc.get('employer_name', 'Empresa')} "
                f"• Valor da Causa: {fmt_brl(calc['liquido_devido'])}"
            )
            with st.expander(titulo_exp, expanded=(idx == 0)):
                df_rep = pd.DataFrame(calc["verbas"])
                for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"]:
                    df_rep[col] = df_rep[col].apply(fmt_brl)
                st.dataframe(df_rep, use_container_width=True, hide_index=True)

                b1, b2 = st.columns(2)
                b1.download_button(
                    "Exportar Dossiê Excel (.XLSX)",
                    data=build_complete_excel_bytes(calc),
                    file_name=f"Dossie_Trabalhista_{calc['audit_id']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_xlsx_{calc['audit_id']}_{idx}",
                    use_container_width=True,
                )
                b2.download_button(
                    "Exportar Relatório em PDF (.PDF)",
                    data=build_printable_dossier_pdf_bytes(calc),
                    file_name=f"Relatorio_Dossie_{calc['audit_id']}.pdf",
                    mime="application/pdf",
                    key=f"dl_pdf_{calc['audit_id']}_{idx}",
                    use_container_width=True,
                )
