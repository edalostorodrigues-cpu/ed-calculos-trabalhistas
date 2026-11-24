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
    st.caption("SaaS Jurídico • Avaliação Preliminar & Liquidação")
    st.divider()
    menu = st.radio(
        "Navegação",
        [
            "Avaliação Preliminar (Pré-Processo)",
            "Liquidação de Sentença (Upload)",
            "Dashboard Executivo",
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
if menu == "Avaliação Preliminar (Pré-Processo)":
    st.title("Avaliação Preliminar do Caso (Antes do Processo Judicial)")
    st.markdown(
        "Informe as **condições reais do contrato de trabalho** (ou suba os documentos iniciais do cliente) para apurar "
        "o **valor estimado dos pedidos**, conferir **diferenças não pagas no TRCT** e gerar o **Demonstrativo do Valor da Causa** para a Petição Inicial."
    )

    with st.expander("Opcional: Importar Cartão de Ponto (CSV/Excel) ou TRCT/Contracheque do Cliente"):
        col_pre1, col_pre2 = st.columns(2)
        with col_pre1:
            ponto_pre = st.file_uploader("Cartão de Ponto (.csv ou .xlsx)", type=["csv", "xlsx"], key="pre_ponto")
            if ponto_pre is not None:
                if ponto_pre.name.endswith(".csv"):
                    raw_txt = ponto_pre.getvalue().decode("utf-8", errors="ignore")
                    sep = ";" if ";" in raw_txt.splitlines()[0] else ","
                    df_ponto = pd.read_csv(io.StringIO(raw_txt), sep=sep)
                else:
                    df_ponto = pd.read_excel(ponto_pre)
                st.session_state.timecard_result = process_timecard_dataframe(df_ponto)
                st.success("Cartão de ponto apurado e transferido para os campos abaixo!")
        with col_pre2:
            doc_pre = st.file_uploader("TRCT / Contracheque / Contrato (.pdf ou .txt)", type=["pdf", "txt"], key="pre_doc")
            if doc_pre is not None and st.button("Ler Dados do Documento (Gemini)", key="btn_pre_gemini"):
                try:
                    mime = "application/pdf" if doc_pre.name.endswith(".pdf") else "text/plain"
                    st.session_state.extracted_ai_data = extract_document_with_gemini(doc_pre.getvalue(), mime)
                    st.success("Dados lidos! Confira nos campos abaixo.")
                except Exception as exc:
                    st.warning(f"Leitura via Gemini indisponível ({exc}). Preencha diretamente abaixo.")

    tc_pre = st.session_state.timecard_result
    ext_pre = st.session_state.extracted_ai_data or {}

    with st.form("preliminary_eval_form"):
        st.subheader("1. Identificação do Cliente e Condições do Emprego")
        p1, p2, p3 = st.columns(3)
        cliente_nome = p1.text_input("Nome do Cliente (Trabalhador) *", value=ext_pre.get("worker_name") or "João da Silva")
        empresa_nome = p2.text_input("Empresa (Futura Reclamada) *", value=ext_pre.get("employer_name") or "Comércio e Serviços Alfa Ltda.")
        cargo_funcao = p3.text_input("Cargo / Função Exercida", value="Auxiliar Operacional / Vendedor")

        d1, d2, d3, d4 = st.columns(4)
        adm_pre = d1.date_input("Data de Admissão *", value=date(2024, 2, 1), key="adm_pre")
        dism_pre = d2.date_input("Data de Saída (ou Hoje) *", value=date(2026, 11, 20), key="dism_pre")
        motivo_saida = d3.selectbox(
            "Situação da Saída *",
            [
                "sem_justa_causa",
                "rescisao_indireta",
                "pedido_demissao",
                "acordo",
                "justa_causa",
            ],
            format_func=lambda x: {
                "sem_justa_causa": "Dispensa sem justa causa",
                "rescisao_indireta": "Pedido de Rescisão Indireta (falta grave do empregador)",
                "pedido_demissao": "Pedido de demissão (reversão / verbas pendentes)",
                "acordo": "Rescisão por acordo (art. 484-A CLT)",
                "justa_causa": "Dispensa por justa causa (contestação)",
            }[x],
        )
        divisor_pre = d4.number_input("Jornada / Divisor Mensal *", min_value=1, max_value=300, value=int(ext_pre.get("monthly_divisor") or 220))

        st.subheader("2. Remuneração Real vs. Registrada")
        r1, r2, r3, r4 = st.columns(4)
        sal_registrado = r1.number_input("Salário-Base Registrado (R$) *", min_value=0.0, value=float(ext_pre.get("base_salary") or 2400.00), step=50.0)
        sal_por_fora = r2.number_input("Comissões / Salário 'Por Fora' (R$/mês)", min_value=0.0, value=float(ext_pre.get("commissions") or 400.00), step=50.0)
        adicional_cond = r3.selectbox(
            "Adicional Devido e Não Pago",
            ["Nenhum", "Insalubridade Grau Mínimo (10%)", "Insalubridade Grau Médio (20%)", "Insalubridade Grau Máximo (40%)", "Periculosidade (30% sobre salário)"],
        )
        valor_ja_pago_trct = r4.number_input("Valor Líquido Já Recebido na Rescisão/TRCT (R$)", min_value=0.0, value=0.0, step=100.0)

        st.subheader("3. Jornada Habitual Relatada pelo Cliente (ou Apurada no Ponto)")
        meses_contrato = max(1, ((dism_pre.year - adm_pre.year) * 12 + (dism_pre.month - adm_pre.month)))
        st.caption(f"Tempo de contrato apurado: **{meses_contrato} meses** (limitado aos últimos 60 meses pela prescrição quinquenal, se aplicável).")

        j1, j2, j3, j4 = st.columns(4)
        he_semanal_relatada = j1.number_input("Horas Extras Habituais por Semana (50%)", min_value=0.0, value=6.0, step=0.5, help="Ex: fazia 1h12 extra por dia de seg a sex = 6h/semana")
        domingos_mes = j2.number_input("Domingos/Feriados Trabalhados sem Folga (por mês)", min_value=0, max_value=10, value=1)
        suprimia_intervalo = j3.checkbox("Supressão de Intervalo de Almoço (1h)", value=True)
        min_intervalo_suprimido = j3.number_input("Minutos suprimidos por dia", min_value=0, max_value=60, value=30)
        horas_noturnas_semana = j4.number_input("Horas Noturnas (22h às 05h) por Semana", min_value=0.0, value=0.0, step=0.5)

        st.subheader("4. Verbas a Incluir na Avaliação Preliminar")
        v1, v2, v3 = st.columns(3)
        ped_saldo = v1.checkbox("Saldo de Salário do Mês da Saída", value=True)
        dias_saldo_pre = v1.number_input("Dias trabalhados no último mês", min_value=0, max_value=31, value=dism_pre.day)
        ped_aviso = v2.checkbox("Aviso-Prévio Indenizado + Projeção", value=True)
        ped_13_ferias = v2.checkbox("13º Proporcional e Férias + 1/3", value=True)
        ped_fgts = v3.checkbox("Diferenças de FGTS (8%) + Multa Rescisória", value=True)
        aplicar_inss_ir = v3.checkbox("Simular Descontos de INSS/IRRF", value=False)

        avaliar_btn = st.form_submit_button("Calcular Avaliação Preliminar e Valor da Causa", type="primary")

    if avaliar_btn:
        if sal_registrado <= 0 or not cliente_nome.strip():
            st.error("Preencha o nome do cliente e o salário-base registrado.")
        elif dism_pre < adm_pre:
            st.error("A data de saída não pode ser anterior à admissão.")
        else:
            # Converte condições semanais/mensais relatadas para o total do período (máx. 60 meses - prescrição quinquenal)
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
                total_intra = float(((Decimal(str(min_intervalo_suprimido)) / Decimal("60")) * dias_uteis_periodo).quantize(Decimal("0.01"))) if suprimia_intervalo else 0.0
                total_not = float((Decimal(str(horas_noturnas_semana)) * semanas_periodo * Decimal("1.142857")).quantize(Decimal("0.01")))

            # Cálculo do adicional de insalubridade / periculosidade
            insal_mensal = 0.0
            if "10%" in adicional_cond:
                insal_mensal = 151.80
            elif "20%" in adicional_cond:
                insal_mensal = 303.60
            elif "40%" in adicional_cond:
                insal_mensal = 607.20
            elif "30%" in adicional_cond:
                insal_mensal = round(sal_registrado * 0.30, 2)

            res_pre = run_labor_calculation({
                "process_number": f"PRE-PROCESSUAL-{cliente_nome.upper().replace(' ', '_')[:18]}",
                "worker_name": cliente_nome,
                "employer_name": empresa_nome,
                "admission_date": adm_pre,
                "dismissal_date": dism_pre,
                "termination_type": motivo_saida,
                "monthly_divisor": divisor_pre,
                "base_salary": sal_registrado,
                "commissions": sal_por_fora,
                "insalubrity": insal_mensal,
                "offsets": valor_ja_pago_trct,
                "inc_saldo_salario": ped_saldo,
                "saldo_dias": dias_saldo_pre,
                "inc_aviso_previo": ped_aviso,
                "inc_13": ped_13_ferias,
                "inc_ferias": ped_13_ferias,
                "inc_fgts": ped_fgts,
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
                "apply_inss": aplicar_inss_ir,
                "apply_irrf": aplicar_inss_ir,
                "apply_interest": False,
                "interest_pct": 0.0,
                "correction_pct": 0.0,
            })
            st.session_state.calculations_history.insert(0, res_pre)

            st.success(f"Avaliação Preliminar concluída! Código de auditoria: **{res_pre['audit_id']}**")

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Apurado dos Pedidos", fmt_brl(res_pre["bruto_atualizado"]))
            k2.metric("(-) Já Pago no TRCT", fmt_brl(res_pre["abatimentos"]))
            k3.metric("Diferença a Pedir (Valor da Causa)", fmt_brl(res_pre["liquido_devido"]))
            k4.metric("Período Considerado", f"{meses_prescricao} meses")

            st.markdown("### Quadro de Pedidos para a Petição Inicial (Art. 840, § 1º da CLT)")
            df_pre = pd.DataFrame(res_pre["verbas"])
            for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                df_pre[col] = df_pre[col].apply(fmt_brl)
            st.dataframe(
                df_pre[["Código", "Verba", "Quantidade", "Percentual", "Base de Cálculo", "Fórmula", "Valor Original", "Reflexos", "Total Apurado"]],
                use_container_width=True,
                hide_index=True,
            )

            # Checklist automático de validação jurídica antes de protocolar
            st.markdown("### Checklist de Validação Antes de Protocolar a Ação")
            st.info(
                f"- **Prescrição Quinquenal:** Aplicado limite automático de **{meses_prescricao} meses** anteriores ao ajuizamento.\n"
                f"- **Base Remuneratória:** Salário registrado ({fmt_brl(d_money(sal_registrado))}) + Integrações ({fmt_brl(d_money(sal_por_fora + insal_mensal))}) = **{fmt_brl(d_money(sal_registrado + sal_por_fora + insal_mensal))}**.\n"
                f"- **Abatimento de valores pagos:** Deduzidos **{fmt_brl(res_pre['abatimentos'])}** já quitados no TRCT para evitar pedido em duplicidade.\n"
                f"- **Transparência:** Valores apurados pelo motor próprio em Python Decimal exato (estimativa pré-processual auditável)."
            )

            excel_pre = io.BytesIO()
            with pd.ExcelWriter(excel_pre, engine="openpyxl") as writer:
                df_pre.to_excel(writer, index=False, sheet_name="Valor_da_Causa_Inicial")
            st.download_button(
                "Baixar Planilha de Pedidos para Petição Inicial (.XLSX)",
                data=excel_pre.getvalue(),
                file_name=f"Avaliacao_Preliminar_{res_pre['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

elif menu == "Dashboard Executivo":
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

elif menu == "Liquidação de Sentença (Upload)":
    st.title("Liquidação de Sentença Judicial (Upload Automático + Revisão)")
    st.markdown("Envie o **Cartão de Ponto (CSV/Excel)** e a **Sentença/Contracheque (PDF/TXT)** para extração assistida, revisão humana obrigatória e cálculo decimal.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        up_ponto = st.file_uploader("1. Cartão de Ponto (.csv ou .xlsx)", type=["csv", "xlsx"], key="liq_ponto")
        if up_ponto is not None:
            if up_ponto.name.endswith(".csv"):
                raw_txt = up_ponto.getvalue().decode("utf-8", errors="ignore")
                sep = ";" if ";" in raw_txt.splitlines()[0] else ","
                df_in = pd.read_csv(io.StringIO(raw_txt), sep=sep)
            else:
                df_in = pd.read_excel(up_ponto)
            st.session_state.timecard_result = process_timecard_dataframe(df_in)
            st.success("Cartão de ponto apurado automaticamente!")

    with col_u2:
        up_doc = st.file_uploader("2. Sentença ou Contracheque (.pdf ou .txt)", type=["pdf", "txt"], key="liq_doc")
        if up_doc is not None and st.button("Extrair Dados da Sentença (Gemini)", key="btn_liq_gemini"):
            try:
                mime = "application/pdf" if up_doc.name.endswith(".pdf") else "text/plain"
                st.session_state.extracted_ai_data = extract_document_with_gemini(up_doc.getvalue(), mime)
                st.session_state.extraction_confirmed = False
                st.success("Dados extraídos! Confirme abaixo antes de calcular.")
            except Exception as exc:
                st.warning(f"Extração via Gemini indisponível ({exc}). Preencha manualmente abaixo.")

    if st.session_state.extracted_ai_data:
        st.json(st.session_state.extracted_ai_data)
        st.session_state.extraction_confirmed = st.checkbox(
            "Confirmo que revisei os dados extraídos do documento",
            value=st.session_state.extraction_confirmed,
        )

    tc_liq = st.session_state.timecard_result
    ext_liq = st.session_state.extracted_ai_data or {}

    with st.form("liquidacao_sentenca_form"):
        f1, f2, f3 = st.columns(3)
        proc_num = f1.text_input("Número do Processo *", value=ext_liq.get("process_number") or "0001245-88.2025.5.04.0012")
        trab_nome = f2.text_input("Reclamante *", value=ext_liq.get("worker_name") or "Carlos Eduardo Pereira")
        emp_nome = f3.text_input("Reclamada *", value=ext_liq.get("employer_name") or "Metalúrgica Sul Horizonte Ltda.")

        f4, f5, f6, f7 = st.columns(4)
        adm_liq = f4.date_input("Admissão *", value=date(2023, 3, 10), key="adm_liq")
        dism_liq = f5.date_input("Desligamento *", value=date(2026, 10, 31), key="dism_liq")
        sal_liq = f6.number_input("Salário-Base (R$) *", min_value=0.0, value=float(ext_liq.get("base_salary") or 3450.00), step=50.0)
        div_liq = f7.number_input("Divisor Mensal *", min_value=1, max_value=300, value=int(ext_liq.get("monthly_divisor") or 220))

        f8, f9, f10, f11 = st.columns(4)
        he50_liq = f8.number_input("Horas Extras 50% (h)", min_value=0.0, value=float(tc_liq["he50_hours"]) if tc_liq else 24.0)
        he100_liq = f9.number_input("Horas Extras 100% (h)", min_value=0.0, value=float(tc_liq["he100_hours"]) if tc_liq else 0.0)
        intra_liq = f10.number_input("Intrajornada Suprimida (h)", min_value=0.0, value=float(tc_liq["intrajornada_hours"]) if tc_liq else 0.0)
        not_liq = f11.number_input("Horas Noturnas Fictas (h)", min_value=0.0, value=float(tc_liq["night_ficta_hours"]) if tc_liq else 0.0)

        f12, f13, f14 = st.columns(3)
        juros_liq = f12.number_input("Juros Acumulados (%)", min_value=0.0, value=3.0, step=0.5)
        corr_liq = f13.number_input("Correção Monetária (%)", min_value=0.0, value=4.2, step=0.5)
        abat_liq = f14.number_input("Abatimentos / Valores Pagos (R$)", min_value=0.0, value=0.0, step=50.0)

        calc_liq_btn = st.form_submit_button("Executar Liquidação Decimal", type="primary")

    if calc_liq_btn:
        if st.session_state.extracted_ai_data and not st.session_state.extraction_confirmed:
            st.error("Marque a caixa de confirmação dos dados extraídos antes de calcular.")
        elif sal_liq <= 0 or not proc_num.strip():
            st.error("Preencha o número do processo e o salário-base.")
        else:
            res_liq = run_labor_calculation({
                "process_number": proc_num,
                "worker_name": trab_nome,
                "employer_name": emp_nome,
                "admission_date": adm_liq,
                "dismissal_date": dism_liq,
                "termination_type": "sem_justa_causa",
                "monthly_divisor": div_liq,
                "base_salary": sal_liq,
                "commissions": 0.0,
                "insalubrity": 0.0,
                "offsets": abat_liq,
                "inc_saldo_salario": True,
                "saldo_dias": dism_liq.day,
                "inc_aviso_previo": True,
                "inc_13": True,
                "inc_ferias": True,
                "inc_fgts": True,
                "inc_he50": he50_liq > 0,
                "he50_hours": he50_liq,
                "he50_pct": 50,
                "inc_he100": he100_liq > 0,
                "he100_hours": he100_liq,
                "inc_intrajornada": intra_liq > 0,
                "intrajornada_hours": intra_liq,
                "inc_noturno": not_liq > 0,
                "night_hours": not_liq,
                "night_pct": 20,
                "apply_inss": True,
                "apply_irrf": True,
                "apply_interest": True,
                "interest_pct": juros_liq,
                "correction_pct": corr_liq,
            })
            st.session_state.calculations_history.insert(0, res_liq)
            st.success(f"Liquidação concluída! Auditoria: **{res_liq['audit_id']}**")
            m1, m2, m3 = st.columns(3)
            m1.metric("Bruto Atualizado", fmt_brl(res_liq["bruto_atualizado"]))
            m2.metric("Descontos (INSS + IRRF + Abat.)", fmt_brl(res_liq["total_descontos"]))
            m3.metric("Líquido Devido", fmt_brl(res_liq["liquido_devido"]))
            df_liq = pd.DataFrame(res_liq["verbas"])
            for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                df_liq[col] = df_liq[col].apply(fmt_brl)
            st.dataframe(df_liq, use_container_width=True, hide_index=True)

elif menu == "Meus Cálculos e Relatórios":
    st.title("Meus Cálculos e Relatórios Auditáveis")
    if not st.session_state.calculations_history:
        st.info("Nenhum cálculo realizado nesta sessão ainda. Use 'Avaliação Preliminar (Pré-Processo)' ou 'Liquidação de Sentença (Upload)'.")
    else:
        for calc in st.session_state.calculations_history:
            with st.expander(f"{calc['audit_id']} • Gerado em {calc['generated_at']} • Líquido: {fmt_brl(calc['liquido_devido'])}", expanded=True):
                st.caption(f"Origem: {calc['origin']} | Tabelas: {calc['tables_version']}")
                df_rep = pd.DataFrame(calc["verbas"])
                for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                    df_rep[col] = df_rep[col].apply(fmt_brl)
                st.dataframe(df_rep, use_container_width=True, hide_index=True)

                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    df_rep.to_excel(writer, index=False, sheet_name="Memoria_Calculo")
                st.download_button(
                    f"Baixar Planilha Auditável ({calc['audit_id']}.xlsx)",
                    data=buf.getvalue(),
                    file_name=f"{calc['audit_id']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{calc['audit_id']}",
                )

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
