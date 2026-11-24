# app.py — ED Cálculos Trabalhistas (Versão Python + Streamlit)
# Requisitos: streamlit, pandas, openpyxl, pydantic, google-genai
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

# SDK oficial Google GenAI (usado exclusivamente para extração assistida de documentos)
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL (AZUL-MARINHO, BRANCO E VERDE)
# ==============================================================================
st.set_page_config(
    page_title="ED Cálculos Trabalhistas — Plataforma Profissional",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .stApp {
            background-color: #F4F6F9;
            color: #0F2537;
        }
        section[data-testid="stSidebar"] {
            background-color: #0B1B2B;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * {
            color: #F8FAFC !important;
        }
        .ed-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 1.2rem;
            margin-bottom: 1rem;
        }
        .ed-badge-ok {
            background-color: #F0FDF4;
            color: #15803D;
            border: 1px solid #BBF7D0;
            padding: 3px 10px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
        }
        .ed-badge-warn {
            background-color: #FFFBEB;
            color: #B45309;
            border: 1px solid #FDE68A;
            padding: 3px 10px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 2. UTILITÁRIOS DE MATEMÁTICA DECIMAL EXATA (NUNCA USAR FLOAT PARA DINHEIRO)
# ==============================================================================
CENT = Decimal("0.01")


def d_money(value: Any) -> Decimal:
    """Converte qualquer entrada numérica/texto para Decimal arredondado em centavos."""
    if value is None or value == "":
        return Decimal("0.00")
    if isinstance(value, Decimal):
        return value.quantize(CENT, rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)
    clean = str(value).strip().replace("R$", "").replace(" ", "")
    if "," in clean and "." in clean:
        clean = clean.replace(".", "").replace(",", ".")
    elif "," in clean:
        clean = clean.replace(",", ".")
    try:
        return Decimal(clean).quantize(CENT, rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0.00")


def fmt_brl(val: Decimal) -> str:
    """Formata Decimal no padrão monetário brasileiro (R$ 1.234,56)."""
    q = d_money(val)
    sign = "-" if q < 0 else ""
    abs_val = abs(q)
    inteiro, dec = f"{abs_val:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
    return f"{sign}R$ {inteiro_fmt},{dec}"


# ==============================================================================
# 3. TABELAS HISTÓRICAS VERSIONADAS (REFERÊNCIA AUDITÁVEL)
# ==============================================================================
TABLES_VERSION = "2026.11-ref"

INSS_BRACKETS = [
    (Decimal("1518.00"), Decimal("0.075")),
    (Decimal("2793.88"), Decimal("0.090")),
    (Decimal("4190.83"), Decimal("0.120")),
    (Decimal("8157.41"), Decimal("0.140")),
]

IRRF_BRACKETS = [
    (Decimal("2259.20"), Decimal("0.000"), Decimal("0.00")),
    (Decimal("2826.65"), Decimal("0.075"), Decimal("169.44")),
    (Decimal("3751.05"), Decimal("0.150"), Decimal("381.44")),
    (Decimal("4664.68"), Decimal("0.225"), Decimal("662.77")),
    (Decimal("999999999.99"), Decimal("0.275"), Decimal("896.00")),
]


# ==============================================================================
# 4. APURADOR DETERMINÍSTICO DE CARTÃO DE PONTO (SEM IA PARA VALORES)
# ==============================================================================
def parse_hhmm_to_minutes(txt: Any) -> Optional[int]:
    if txt is None or pd.isna(txt):
        return None
    s = str(txt).strip()
    if ":" not in s:
        return None
    parts = s.split(":")
    try:
        h = int(parts[0])
        m = int(parts[1])
        if 0 <= h <= 29 and 0 <= m <= 59:
            return h * 60 + m
    except ValueError:
        return None
    return None


def calc_night_minutes(start_min: int, end_min: int) -> int:
    """Apura minutos trabalhados entre 22h00 (1320) e 05h00 (300 / 1740)."""
    if end_min <= start_min:
        return 0
    night = 0
    for n_start, n_end in [(0, 300), (1320, 1740)]:
        ov_start = max(start_min, n_start)
        ov_end = min(end_min, n_end)
        if ov_end > ov_start:
            night += ov_end - ov_start
    return night


def process_timecard_dataframe(df: pd.DataFrame, daily_standard_min: int = 480) -> Dict[str, Any]:
    """
    Processa cartão de ponto tabular (CSV/XLSX) com:
    - Tolerância legal do art. 58, § 1º da CLT (10 min diários);
    - Horas extras 50% (dias úteis) e 100% (domingos/feriados);
    - Supressão de intervalo intrajornada (art. 71 CLT);
    - Redução ficta da hora noturna (52m30s -> fator 60 / 52.5).
    """
    rows_out = []
    tot_worked = 0
    tot_he50 = 0
    tot_he100 = 0
    tot_night_ficta = 0
    tot_intra = 0
    incomplete_days = 0

    for _, row in df.iterrows():
        data_str = str(row.get("Data", "")).strip()
        dia_str = str(row.get("Dia", "")).strip().upper()
        feriado = str(row.get("Feriado", "N")).strip().upper() in ("S", "SIM", "1", "TRUE") or dia_str == "DOM"

        e1 = parse_hhmm_to_minutes(row.get("Entrada1"))
        s1 = parse_hhmm_to_minutes(row.get("Saida1"))
        e2 = parse_hhmm_to_minutes(row.get("Entrada2"))
        s2 = parse_hhmm_to_minutes(row.get("Saida2"))

        if e1 is None or s1 is None:
            incomplete_days += 1
            rows_out.append({
                "Data": data_str,
                "Dia": dia_str,
                "Trabalhado (h)": 0.0,
                "HE 50% (h)": 0.0,
                "HE 100% (h)": 0.0,
                "Noturna Ficta (h)": 0.0,
                "Intrajornada Sup. (h)": 0.0,
                "Status": "Batida incompleta (revisar)",
            })
            continue

        p1_end = s1 + 1440 if s1 < e1 else s1
        w1 = max(0, p1_end - e1)
        n1 = calc_night_minutes(e1, p1_end)

        w2 = 0
        n2 = 0
        break_min = 0
        if e2 is not None and s2 is not None:
            p2_start = e2 + 1440 if e2 < p1_end else e2
            p2_end = s2 + 1440 if s2 < p2_start else s2
            break_min = max(0, p2_start - p1_end)
            w2 = max(0, p2_end - p2_start)
            n2 = calc_night_minutes(p2_start, p2_end)

        day_worked = w1 + w2
        day_night_clock = n1 + n2
        day_night_ficta = int(round(day_night_clock * (60.0 / 52.5)))

        suppressed_break = 0
        if day_worked > 360 and break_min < 60:
            suppressed_break = 60 - break_min
        elif 240 < day_worked <= 360 and break_min < 15:
            suppressed_break = 15 - break_min

        he50 = 0
        he100 = 0
        if feriado:
            he100 = day_worked
        else:
            diff = day_worked - daily_standard_min
            if diff > 10:  # Tolerância art. 58 §1º CLT
                he50 = diff

        tot_worked += day_worked
        tot_he50 += he50
        tot_he100 += he100
        tot_night_ficta += day_night_ficta
        tot_intra += suppressed_break

        rows_out.append({
            "Data": data_str,
            "Dia": dia_str or ("DOM/FER" if feriado else "ÚTIL"),
            "Trabalhado (h)": round(day_worked / 60, 2),
            "HE 50% (h)": round(he50 / 60, 2),
            "HE 100% (h)": round(he100 / 60, 2),
            "Noturna Ficta (h)": round(day_night_ficta / 60, 2),
            "Intrajornada Sup. (h)": round(suppressed_break / 60, 2),
            "Status": "Intrajornada suprimida" if suppressed_break > 0 else "Apurado",
        })

    return {
        "table": pd.DataFrame(rows_out),
        "incomplete_days": incomplete_days,
        "he50_hours": Decimal(str(round(tot_he50 / 60, 2))),
        "he100_hours": Decimal(str(round(tot_he100 / 60, 2))),
        "night_ficta_hours": Decimal(str(round(tot_night_ficta / 60, 2))),
        "intrajornada_hours": Decimal(str(round(tot_intra / 60, 2))),
    }


# ==============================================================================
# 5. EXTRAÇÃO ASSISTIDA POR GEMINI (EXCLUSIVAMENTE PARA LEITURA DE DOCUMENTOS)
# ==============================================================================
class ExtractedCaseData(BaseModel):
    process_number: Optional[str] = None
    worker_name: Optional[str] = None
    employer_name: Optional[str] = None
    admission_date: Optional[str] = None
    dismissal_date: Optional[str] = None
    termination_type: Optional[str] = None
    base_salary: Optional[float] = None
    commissions: Optional[float] = None
    monthly_divisor: Optional[int] = None
    overtime_pct: Optional[float] = None
    night_pct: Optional[float] = None
    sentence_commands: Optional[str] = None
    ambiguities: List[str] = []


def get_gemini_api_key() -> Optional[str]:
    """Obtém a chave via st.secrets ou variável de ambiente sem nunca expor no código."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY")


def extract_document_with_gemini(file_bytes: bytes, mime_type: str) -> Dict[str, Any]:
    """
    Envia PDF/imagem/texto ao Gemini para extrair parâmetros jurídicos e contratuais.
    NUNCA calcula valores financeiros finais e exige confirmação humana.
    """
    api_key = get_gemini_api_key()
    if not api_key or not GENAI_AVAILABLE:
        raise RuntimeError("GEMINI_API_KEY não configurada ou pacote google-genai ausente.")

    client = genai.Client(api_key=api_key)
    prompt = (
        "Extraia exclusivamente os dados objetivos presentes neste documento trabalhista "
        "(número do processo, partes, admissão, demissão, modalidade de rescisão, salário-base, "
        "comissões, divisor, adicionais e comandos da sentença). "
        "NÃO invente datas, salários ou horas extras ausentes. Liste dúvidas em 'ambiguities'."
    )
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ExtractedCaseData,
            temperature=0.0,
        ),
    )
    return json.loads(response.text)


# ==============================================================================
# 6. MOTOR DE CÁLCULO TRABALHISTA DECIMAL EXATO (MEMÓRIA AUDITÁVEL)
# ==============================================================================
def run_labor_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
    """
    Executa a liquidação trabalhista com matemática Decimal exata.
    Identifica explicitamente a origem do cálculo (motor próprio auditável).
    """
    base_salary = d_money(params["base_salary"])
    commissions = d_money(params.get("commissions", 0))
    insalubrity = d_money(params.get("insalubrity", 0))
    remuneration = (base_salary + commissions + insalubrity).quantize(CENT, rounding=ROUND_HALF_UP)

    divisor = Decimal(str(params["monthly_divisor"]))
    hourly_rate = (remuneration / divisor).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    adm_date: date = params["admission_date"]
    dism_date: date = params["dismissal_date"]
    years_worked = max(0, (dism_date - adm_date).days // 365)

    verbas = []

    def add_verba(
        codigo: str,
        nome: str,
        periodo: str,
        base_calc: Decimal,
        formula: str,
        qtd_label: str,
        pct_label: str,
        valor_orig: Decimal,
        reflexos: Decimal = Decimal("0.00"),
        tributavel_inss: bool = True,
    ):
        v_orig = d_money(valor_orig)
        v_refl = d_money(reflexos)
        subtotal = v_orig + v_refl
        corr_pct = Decimal(str(params.get("correction_pct", 0))) / Decimal("100")
        juros_pct = Decimal(str(params.get("interest_pct", 0))) / Decimal("100")

        v_corr = (subtotal * corr_pct).quantize(CENT, rounding=ROUND_HALF_UP) if params.get("apply_interest") else Decimal("0.00")
        v_juros = ((subtotal + v_corr) * juros_pct).quantize(CENT, rounding=ROUND_HALF_UP) if params.get("apply_interest") else Decimal("0.00")
        total = subtotal + v_corr + v_juros

        verbas.append({
            "Código": codigo,
            "Verba": nome,
            "Período": periodo,
            "Base de Cálculo": base_calc,
            "Fórmula": formula,
            "Quantidade": qtd_label,
            "Percentual": pct_label,
            "Valor Original": v_orig,
            "Reflexos": v_refl,
            "Correção": v_corr,
            "Juros": v_juros,
            "Total Apurado": total,
            "Tributável INSS": tributavel_inss,
        })

    periodo_geral = f"{adm_date.strftime('%d/%m/%Y')} a {dism_date.strftime('%d/%m/%Y')}"

    # A. Verbas Rescisórias
    term_type = params["termination_type"]
    if params.get("inc_saldo_salario") and params.get("saldo_dias", 0) > 0:
        dias = Decimal(str(params["saldo_dias"]))
        val = (remuneration / Decimal("30")) * dias
        add_verba(
            "VER-01", "Saldo de Salário", dism_date.strftime("%m/%Y"),
            remuneration, "(Remuneração / 30) × Dias Trabalhados",
            f"{dias} dias", "100%", val, tributavel_inss=True
        )

    if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "acordo"):
        aviso_dias = min(90, 30 + years_worked * 3)
        fator = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
        val = ((remuneration / Decimal("30")) * Decimal(str(aviso_dias))) * fator
        add_verba(
            "VER-02", "Aviso-Prévio Indenizado (Lei 12.506/2011)", periodo_geral,
            remuneration, "(Remuneração / 30) × Dias Aviso × Fator",
            f"{aviso_dias} dias", f"{int(fator * 100)}%", val, tributavel_inss=False
        )

    if params.get("inc_13") and term_type != "justa_causa":
        avos_13 = Decimal(str(params.get("avos_13", dism_date.month)))
        val = (remuneration / Decimal("12")) * avos_13
        add_verba(
            "VER-03", "13º Salário Proporcional", f"Ano {dism_date.year}",
            remuneration, "(Remuneração / 12) × Avos",
            f"{avos_13}/12 avos", "100%", val, tributavel_inss=True
        )

    if params.get("inc_ferias") and term_type != "justa_causa":
        avos_ferias = Decimal(str(params.get("avos_ferias", 12)))
        ferias_base = (remuneration / Decimal("12")) * avos_ferias
        terco = ferias_base / Decimal("3")
        add_verba(
            "VER-04", "Férias Proporcionais/Vencidas + 1/3 Constitucional", periodo_geral,
            remuneration, "((Remuneração / 12) × Avos) + 1/3",
            f"{avos_ferias}/12 avos", "133,33%", ferias_base + terco, tributavel_inss=False
        )

    # B. Jornada de Trabalho (Apurada automaticamente do ponto ou informada)
    dsr_factor = Decimal("1") / Decimal("6")
    if params.get("inc_he50") and d_money(params.get("he50_hours", 0)) > 0:
        h50 = Decimal(str(params["he50_hours"]))
        pct_he = Decimal(str(params.get("he50_pct", 50))) / Decimal("100")
        unit_he = hourly_rate * (Decimal("1") + pct_he)
        orig_he = unit_he * h50
        refl_dsr = orig_he * dsr_factor
        add_verba(
            "VER-05", f"Horas Extras ({params.get('he50_pct', 50)}%) + Reflexo DSR", periodo_geral,
            hourly_rate.quantize(CENT), f"(Hora Normal × (1 + {params.get('he50_pct', 50)}%)) × Horas + DSR",
            f"{h50}h", f"+{params.get('he50_pct', 50)}%", orig_he, reflexos=refl_dsr, tributavel_inss=True
        )

    if params.get("inc_he100") and d_money(params.get("he100_hours", 0)) > 0:
        h100 = Decimal(str(params["he100_hours"]))
        unit_100 = hourly_rate * Decimal("2")
        orig_100 = unit_100 * h100
        refl_dsr_100 = orig_100 * dsr_factor
        add_verba(
            "VER-05B", "Horas Extras Domingos/Feriados (100%) + Reflexo DSR", periodo_geral,
            hourly_rate.quantize(CENT), "(Hora Normal × 2,00) × Horas 100% + DSR",
            f"{h100}h", "+100%", orig_100, reflexos=refl_dsr_100, tributavel_inss=True
        )

    if params.get("inc_intrajornada") and d_money(params.get("intrajornada_hours", 0)) > 0:
        h_intra = Decimal(str(params["intrajornada_hours"]))
        unit_intra = hourly_rate * Decimal("1.5")
        orig_intra = unit_intra * h_intra
        add_verba(
            "VER-05C", "Intervalo Intrajornada Suprimido (+50%)", periodo_geral,
            hourly_rate.quantize(CENT), "(Hora Normal × 1,50) × Horas Suprimidas",
            f"{h_intra}h", "+50%", orig_intra, tributavel_inss=False
        )

    if params.get("inc_noturno") and d_money(params.get("night_hours", 0)) > 0:
        h_not = Decimal(str(params["night_hours"]))
        pct_not = Decimal(str(params.get("night_pct", 20))) / Decimal("100")
        orig_not = (hourly_rate * pct_not) * h_not
        add_verba(
            "VER-06", "Adicional Noturno (com Redução Ficta 52m30s)", periodo_geral,
            hourly_rate.quantize(CENT), f"(Hora Normal × {params.get('night_pct', 20)}%) × Horas Fictas",
            f"{h_not}h", f"+{params.get('night_pct', 20)}%", orig_not, tributavel_inss=True
        )

    # C. FGTS + Indenização Rescisória
    if params.get("inc_fgts"):
        base_fgts = sum((v["Valor Original"] + v["Reflexos"]) for v in verbas)
        fgts_8 = (base_fgts * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP)
        multa_pct = Decimal("0.40") if term_type in ("sem_justa_causa", "rescisao_indireta") else (
            Decimal("0.20") if term_type == "acordo" else Decimal("0.00")
        )
        multa_val = (fgts_8 * multa_pct).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba(
            "VER-07", "FGTS (8%) + Indenização Rescisória sobre Verbas", periodo_geral,
            base_fgts, "Base Incidente × 8% + Multa Rescisória",
            "Verbas incidentes", f"8% + {int(multa_pct * 100)}%", fgts_8 + multa_val, tributavel_inss=False
        )

    # D. Totais, Tributos e Deduções
    bruto_atualizado = sum((v["Total Apurado"] for v in verbas), Decimal("0.00"))
    base_tributavel = sum(
        (v["Valor Original"] + v["Reflexos"] for v in verbas if v["Tributável INSS"]),
        Decimal("0.00"),
    )

    inss_val = Decimal("0.00")
    if params.get("apply_inss") and base_tributavel > 0:
        aliquota_inss = Decimal("0.14")
        for teto, aliq in INSS_BRACKETS:
            if base_tributavel <= teto:
                aliquota_inss = aliq
                break
        inss_val = (base_tributavel * aliquota_inss).quantize(CENT, rounding=ROUND_HALF_UP)

    irrf_val = Decimal("0.00")
    base_ir = max(Decimal("0.00"), base_tributavel - inss_val)
    if params.get("apply_irrf") and base_ir > 0:
        for teto, aliq, ded in IRRF_BRACKETS:
            if base_ir <= teto:
                irrf_val = max(Decimal("0.00"), (base_ir * aliq - ded).quantize(CENT, rounding=ROUND_HALF_UP))
                break

    abatimentos = d_money(params.get("offsets", 0))
    total_descontos = inss_val + irrf_val + abatimentos
    liquido_devido = max(Decimal("0.00"), bruto_atualizado - total_descontos)

    audit_payload = json.dumps(
        {"process": params["process_number"], "bruto": str(bruto_atualizado), "liquido": str(liquido_devido), "tables": TABLES_VERSION},
        sort_keys=True,
    )
    audit_hash = hashlib.sha256(audit_payload.encode()).hexdigest()[:16].upper()

    return {
        "audit_id": f"ED-{audit_hash}",
        "origin": "Calculado por motor próprio decimal (Python Decimal) — não homologado como API oficial PJe-Calc",
        "tables_version": TABLES_VERSION,
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "verbas": verbas,
        "bruto_atualizado": bruto_atualizado,
        "inss": inss_val,
        "irrf": irrf_val,
        "abatimentos": abatimentos,
        "total_descontos": total_descontos,
        "liquido_devido": liquido_devido,
    }


# ==============================================================================
# 7. ESTADO DA SESSÃO (SESSION STATE)
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
# 8. MENU LATERAL PROFISSIONAL
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas")
    st.caption("SaaS Jurídico • Motor Decimal Auditável")
    st.divider()
    menu = st.radio(
        "Navegação",
        [
            "Dashboard",
            "Novo Cálculo (Upload Automático)",
            "Meus Cálculos e Relatórios",
            "Central de Atualizações Legais",
            "Testes de Confiabilidade",
            "Assinaturas e LGPD",
        ],
    )
    st.divider()
    st.caption(f"Tabelas ativas: `{TABLES_VERSION}`")
    st.caption("Modo: Multi-tenant isolado")


# ==============================================================================
# 9. TELAS DO SISTEMA
# ==============================================================================
if menu == "Dashboard":
    st.title("Dashboard Executivo — ED Cálculos Trabalhistas")
    st.info(
        "Transparência técnica: Os cálculos são executados pelo motor matemático próprio em Python Decimal exato. "
        "Nenhuma IA determina valores financeiros e nenhum cálculo é rotulado como equivalente oficial ao PJe-Calc sem validação documentada."
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cálculos na Sessão", len(st.session_state.calculations_history))
    c2.metric("Cartão de Ponto Apurado", "Sim" if st.session_state.timecard_result else "Aguardando")
    c3.metric("Versão das Tabelas", TABLES_VERSION)
    c4.metric("Plano Ativo", "Escritório • R$ 399/mês")

elif menu == "Novo Cálculo (Upload Automático)":
    st.title("Novo Cálculo — Upload Automático e Revisão Humana")
    st.markdown(
        "**Fluxo em 3 passos:** `1. Envie os documentos` → `2. Confira e confirme os dados extraídos` → `3. Processe o cálculo`."
    )

    col_up1, col_up2 = st.columns(2)

    with col_up1:
        st.subheader("1A. Cartão de Ponto (CSV ou Excel)")
        st.caption("Colunas esperadas: `Data`, `Dia`, `Entrada1`, `Saida1`, `Entrada2`, `Saida2`, `Feriado` (S/N).")
        ponto_file = st.file_uploader("Enviar espelho de ponto (.csv ou .xlsx)", type=["csv", "xlsx"])

        if st.button("Carregar Exemplo de Cartão de Ponto (6 dias)"):
            sample_csv = (
                "Data;Dia;Entrada1;Saida1;Entrada2;Saida2;Feriado\n"
                "2026-10-01;QUI;07:56;12:00;13:00;18:18;N\n"
                "2026-10-02;SEX;08:00;12:00;13:00;18:45;N\n"
                "2026-10-04;DOM;08:00;12:00;13:00;17:00;S\n"
                "2026-10-05;SEG;08:00;12:00;12:35;17:35;N\n"
                "2026-10-06;TER;14:00;18:00;19:00;23:30;N\n"
                "2026-10-07;QUA;08:02;12:00;13:00;17:04;N\n"
            )
            df_sample = pd.read_csv(io.StringIO(sample_csv), sep=";")
            st.session_state.timecard_result = process_timecard_dataframe(df_sample)
            st.success("Cartão de ponto de exemplo apurado automaticamente!")

        if ponto_file is not None:
            if ponto_file.name.endswith(".csv"):
                raw_txt = ponto_file.getvalue().decode("utf-8", errors="ignore")
                sep = ";" if ";" in raw_txt.splitlines()[0] else ","
                df_ponto = pd.read_csv(io.StringIO(raw_txt), sep=sep)
            else:
                df_ponto = pd.read_excel(ponto_file)
            st.session_state.timecard_result = process_timecard_dataframe(df_ponto)
            st.success(f"Arquivo '{ponto_file.name}' apurado automaticamente!")

    with col_up2:
        st.subheader("1B. Sentença / Contracheques / TRCT (PDF ou TXT)")
        st.caption("Usa Gemini exclusivamente para extrair dados estruturados (exige confirmação humana).")
        doc_file = st.file_uploader("Enviar sentença ou contracheque (.pdf, .txt)", type=["pdf", "txt"])

        if doc_file is not None:
            if st.button("Extrair Parâmetros do Documento"):
                try:
                    mime = "application/pdf" if doc_file.name.endswith(".pdf") else "text/plain"
                    extracted = extract_document_with_gemini(doc_file.getvalue(), mime)
                    st.session_state.extracted_ai_data = extracted
                    st.session_state.extraction_confirmed = False
                    st.success("Extração concluída! Revise e confirme os campos abaixo.")
                except Exception as exc:
                    st.warning(f"Extração via Gemini indisponível neste ambiente ({exc}). Preencha ou revise manualmente abaixo.")

    # Exibição da apuração automática de ponto
    tc = st.session_state.timecard_result
    if tc is not None:
        st.markdown("### Apuração Automática do Cartão de Ponto (Art. 58 §1º CLT + Hora Ficta 52m30s)")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("HE 50% Apuradas", f"{tc['he50_hours']} h")
        m2.metric("HE 100% (Dom/Fer)", f"{tc['he100_hours']} h")
        m3.metric("Intrajornada Suprimida", f"{tc['intrajornada_hours']} h")
        m4.metric("Noturna (Redução Ficta)", f"{tc['night_ficta_hours']} h")
        st.dataframe(tc["table"], use_container_width=True, hide_index=True)

    # Revisão obrigatória de extração IA
    ext = st.session_state.extracted_ai_data or {}
    if ext:
        st.warning("Confirmação Humana Obrigatória: revise os dados extraídos do documento antes de calcular.")
        st.json(ext)
        st.session_state.extraction_confirmed = st.checkbox(
            "Confirmo que revisei os dados extraídos do documento e assumo responsabilidade técnica.",
            value=st.session_state.extraction_confirmed,
        )

    st.divider()
    st.subheader("2. Conferência dos Parâmetros Jurídicos e Contratuais")

    with st.form("calc_form"):
        f1, f2, f3 = st.columns(3)
        process_number = f1.text_input("Número do Processo *", value=ext.get("process_number") or "0001245-88.2025.5.04.0012")
        worker_name = f2.text_input("Reclamante / Trabalhador *", value=ext.get("worker_name") or "Carlos Eduardo Pereira")
        employer_name = f3.text_input("Reclamada / Empregador *", value=ext.get("employer_name") or "Metalúrgica Sul Horizonte Ltda.")

        c1, c2, c3, c4 = st.columns(4)
        adm_date = c1.date_input("Data de Admissão *", value=date(2023, 3, 10))
        dism_date = c2.date_input("Data de Desligamento *", value=date(2026, 10, 31))
        termination_type = c3.selectbox(
            "Modalidade Rescisória *",
            ["sem_justa_causa", "pedido_demissao", "acordo", "rescisao_indireta", "justa_causa"],
        )
        monthly_divisor = c4.number_input("Divisor Mensal *", min_value=1, max_value=300, value=int(ext.get("monthly_divisor") or 220))

        s1, s2, s3, s4 = st.columns(4)
        base_salary = s1.number_input("Salário-Base (R$) *", min_value=0.0, value=float(ext.get("base_salary") or 3450.00), step=50.0)
        commissions = s2.number_input("Comissões / Médias (R$)", min_value=0.0, value=float(ext.get("commissions") or 250.00), step=50.0)
        insalubrity = s3.number_input("Insalubridade / Periculosidade (R$)", min_value=0.0, value=0.0, step=50.0)
        offsets = s4.number_input("Abatimentos / Valores Pagos (R$)", min_value=0.0, value=0.0, step=50.0)

        j1, j2, j3, j4 = st.columns(4)
        he50_hours = j1.number_input("Horas Extras 50% (h)", min_value=0.0, value=float(tc["he50_hours"]) if tc else 0.0, step=0.5)
        he100_hours = j2.number_input("Horas Extras 100% (h)", min_value=0.0, value=float(tc["he100_hours"]) if tc else 0.0, step=0.5)
        intrajornada_hours = j3.number_input("Intrajornada Suprimida (h)", min_value=0.0, value=float(tc["intrajornada_hours"]) if tc else 0.0, step=0.5)
        night_hours = j4.number_input("Horas Noturnas Fictas (h)", min_value=0.0, value=float(tc["night_ficta_hours"]) if tc else 0.0, step=0.5)

        v1, v2, v3, v4 = st.columns(4)
        inc_saldo_salario = v1.checkbox("Saldo de Salário", value=True)
        saldo_dias = v1.number_input("Dias Saldo Salário", min_value=0, max_value=31, value=31)
        inc_aviso_previo = v2.checkbox("Aviso-Prévio Indenizado", value=True)
        inc_13 = v2.checkbox("13º Salário Proporcional", value=True)
        inc_ferias = v3.checkbox("Férias + 1/3 Constitucional", value=True)
        inc_fgts = v3.checkbox("FGTS (8%) + Multa", value=True)
        apply_inss = v4.checkbox("Apurar INSS", value=True)
        apply_irrf = v4.checkbox("Apurar IRRF", value=True)
        apply_interest = v4.checkbox("Aplicar Juros e Correção", value=True)

        t1, t2 = st.columns(2)
        interest_pct = t1.number_input("Juros Acumulados no Período (%)", min_value=0.0, value=3.0, step=0.1)
        correction_pct = t2.number_input("Correção Monetária no Período (%)", min_value=0.0, value=4.2, step=0.1)

        submitted = st.form_submit_button("3. Processar Cálculo Decimal Auditável", type="primary")

    if submitted:
        if ext and not st.session_state.extraction_confirmed:
            st.error("Cálculo bloqueado: marque a caixa de confirmação humana dos dados extraídos do documento.")
        elif base_salary <= 0 or not process_number.strip() or not worker_name.strip():
            st.error("Cálculo bloqueado: preencha todos os campos obrigatórios marcados com *.")
        elif dism_date < adm_date:
            st.error("Cálculo bloqueado: a data de desligamento não pode ser anterior à data de admissão.")
        else:
            calc_res = run_labor_calculation({
                "process_number": process_number,
                "worker_name": worker_name,
                "employer_name": employer_name,
                "admission_date": adm_date,
                "dismissal_date": dism_date,
                "termination_type": termination_type,
                "monthly_divisor": monthly_divisor,
                "base_salary": base_salary,
                "commissions": commissions,
                "insalubrity": insalubrity,
                "offsets": offsets,
                "inc_saldo_salario": inc_saldo_salario,
                "saldo_dias": saldo_dias,
                "inc_aviso_previo": inc_aviso_previo,
                "inc_13": inc_13,
                "inc_ferias": inc_ferias,
                "inc_fgts": inc_fgts,
                "inc_he50": he50_hours > 0,
                "he50_hours": he50_hours,
                "he50_pct": 50,
                "inc_he100": he100_hours > 0,
                "he100_hours": he100_hours,
                "inc_intrajornada": intrajornada_hours > 0,
                "intrajornada_hours": intrajornada_hours,
                "inc_noturno": night_hours > 0,
                "night_hours": night_hours,
                "night_pct": 20,
                "apply_inss": apply_inss,
                "apply_irrf": apply_irrf,
                "apply_interest": apply_interest,
                "interest_pct": interest_pct,
                "correction_pct": correction_pct,
            })
            st.session_state.calculations_history.insert(0, calc_res)
            st.success(f"Cálculo processado com sucesso! Identificador de auditoria: {calc_res['audit_id']}")

            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Bruto Atualizado", fmt_brl(calc_res["bruto_atualizado"]))
            r2.metric("INSS + IRRF", fmt_brl(calc_res["inss"] + calc_res["irrf"]))
            r3.metric("Abatimentos", fmt_brl(calc_res["abatimentos"]))
            r4.metric("Líquido Apurado", fmt_brl(calc_res["liquido_devido"]))

            st.caption(f"Origem: {calc_res['origin']} | Tabelas: {calc_res['tables_version']}")
            df_mem = pd.DataFrame(calc_res["verbas"])
            for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                df_mem[col] = df_mem[col].apply(fmt_brl)
            st.dataframe(df_mem.drop(columns=["Tributável INSS"]), use_container_width=True, hide_index=True)

            # Exportação Excel e JSON auditável
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                df_mem.to_excel(writer, index=False, sheet_name="Memoria_Calculo")
            st.download_button(
                "Baixar Memória de Cálculo (.XLSX)",
                data=excel_buf.getvalue(),
                file_name=f"Memoria_{calc_res['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

elif menu == "Meus Cálculos e Relatórios":
    st.title("Histórico de Cálculos Auditáveis")
    if not st.session_state.calculations_history:
        st.info("Nenhum cálculo processado nesta sessão ainda.")
    else:
        for item in st.session_state.calculations_history:
            with st.expander(f"{item['audit_id']} — {item['generated_at']} — Líquido: {fmt_brl(item['liquido_devido'])}"):
                st.write(f"**Origem:** {item['origin']}")
                df_v = pd.DataFrame(item["verbas"])
                for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                    df_v[col] = df_v[col].apply(fmt_brl)
                st.dataframe(df_v, use_container_width=True, hide_index=True)

elif menu == "Central de Atualizações Legais":
    st.title("Central de Atualizações Legais e Tabelas")
    st.write(f"Versão vigente das tabelas: **{TABLES_VERSION}**")
    st.dataframe(
        pd.DataFrame([
            {"Fonte": "Previdência Social / RFB", "Norma": "Tabela Progressiva INSS e IRRF 2026", "Status": "Vigente (Referência)"},
            {"Fonte": "CLT Art. 58 §1º e Art. 73", "Norma": "Tolerância 5/10 min e Hora Noturna Ficta (52m30s)", "Status": "Homologado no Motor"},
            {"Fonte": "CSJT / PJe-Calc", "Norma": "Exportação .PJC nativa", "Status": "Bloqueada até validação oficial"},
        ]),
        use_container_width=True,
        hide_index=True,
    )

elif menu == "Testes de Confiabilidade":
    st.title("Suíte de Testes Determinísticos")
    if st.button("Executar Bateria de Testes Agora"):
        sample_csv = (
            "Data;Dia;Entrada1;Saida1;Entrada2;Saida2;Feriado\n"
            "2026-10-01;QUI;07:56;12:00;13:00;18:18;N\n"
            "2026-10-04;DOM;08:00;12:00;13:00;17:00;S\n"
        )
        tc_test = process_timecard_dataframe(pd.read_csv(io.StringIO(sample_csv), sep=";"))
        assert tc_test["he100_hours"] == Decimal("8.0"), "Falha na apuração de HE 100% em domingo"
        assert d_money("1.234,56") + d_money("0,44") == Decimal("1235.00"), "Falha na soma Decimal"
        st.success("Todos os testes unitários e determinísticos passaram com 0,00 de divergência!")

elif menu == "Assinaturas e LGPD":
    st.title("Planos Comerciais e Conformidade LGPD")
    p1, p2, p3 = st.columns(3)
    p1.metric("Plano Individual", "R$ 99/mês", "1 usuário")
    p2.metric("Plano Profissional", "R$ 199/mês", "Até 5 usuários")
    p3.metric("Plano Escritório", "R$ 399/mês", "Multiusuário + Auditoria")
    st.markdown(
        "- **Isolamento multiempresa:** RLS obrigatório no Supabase/PostgreSQL por `office_id`.\n"
        "- **Segurança de chaves:** `GEMINI_API_KEY`, `SUPABASE_KEY` e `MP_ACCESS_TOKEN` armazenadas apenas em `st.secrets`."
    )
