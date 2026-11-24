# app.py — ED Cálculos Trabalhistas (Versão Python + Streamlit)
# Requisitos: streamlit, pandas, openpyxl, pydantic, google-genai, fpdf2
# Execução: streamlit run app.py

import os
import io
import json
import hashlib
import urllib.request
import urllib.parse
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import List, Optional, Dict, Any

import pandas as pd
import streamlit as st
from pydantic import BaseModel, Field
from fpdf import FPDF

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E ESTILO EXECUTIVO
# ==============================================================================
st.set_page_config(
    page_title="ED Cálculos Trabalhistas",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp { background-color: #F4F6F9; color: #0F2537; }
    section[data-testid="stSidebar"] { background-color: #0B1B2B; color: #FFFFFF; }
    section[data-testid="stSidebar"] * { color: #F8FAFC; }
    section[data-testid="stSidebar"] button {
        background-color: #16324A !important;
        color: #FFFFFF !important;
        border: 1px solid #2A4F70 !important;
    }
    section[data-testid="stSidebar"] button:hover {
        background-color: #1E4363 !important;
        border-color: #60A5FA !important;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 1px 2px rgba(15, 37, 55, 0.04);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 2. TABELAS HISTÓRICAS VERSIONADAS E UTILITÁRIOS DECIMAIS EXATOS
# ==============================================================================
TABLES_VERSION = "2026.11-estavel"
CENT = Decimal("0.01")

HISTORICAL_MIN_WAGE: List[tuple[date, Decimal]] = [
    (date(2026, 1, 1), Decimal("1621.00")),
    (date(2025, 1, 1), Decimal("1518.00")),
    (date(2024, 1, 1), Decimal("1412.00")),
    (date(2023, 5, 1), Decimal("1320.00")),
    (date(2023, 1, 1), Decimal("1302.00")),
    (date(2022, 1, 1), Decimal("1212.00")),
    (date(2021, 1, 1), Decimal("1100.00")),
]


def d_money(val: Any) -> Decimal:
    """Converte qualquer entrada numérica/texto para Decimal exato com 2 casas."""
    if val is None:
        return Decimal("0.00")
    if isinstance(val, Decimal):
        return val.quantize(CENT, rounding=ROUND_HALF_UP)
    try:
        s = str(val).strip().replace("R$", "").replace(" ", "")
        if not s:
            return Decimal("0.00")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return Decimal(s).quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def f_float(val: Any, default: float = 0.0) -> float:
    """Converte valor opcional (incluindo None vindo da IA) para float seguro."""
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def f_int(val: Any, default: int = 0) -> int:
    """Converte valor opcional (incluindo None vindo da IA) para int seguro."""
    if val is None:
        return default
    try:
        return int(round(float(val)))
    except (TypeError, ValueError):
        return default


def fmt_brl(val: Any) -> str:
    """Formata Decimal ou float no padrão monetário brasileiro."""
    dec = d_money(val)
    sign = "-" if dec < Decimal("0.00") else ""
    abs_val = abs(dec)
    inteiro, frac = f"{abs_val:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
    return f"{sign}R$ {inteiro_fmt},{frac}"


def get_historical_min_wage(ref_date: date) -> Decimal:
    """Retorna o salário mínimo nacional vigente na data de referência."""
    for start_dt, mw in HISTORICAL_MIN_WAGE:
        if ref_date >= start_dt:
            return mw
    return HISTORICAL_MIN_WAGE[-1][1]


def parse_hhmm(raw: Any) -> Optional[int]:
    """Converte string HH:MM para minutos inteiros desde 00:00."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or ":" not in s:
        return None
    parts = s.split(":")
    if len(parts) < 2:
        return None
    try:
        h, m = int(parts[0]), int(parts[1])
        if 0 <= h <= 29 and 0 <= m <= 59:
            return h * 60 + m
    except ValueError:
        return None
    return None


# ==============================================================================
# 3. CONEXÃO NA NUVEM (SUPABASE REST NATIVO) E ESTADO DA SESSÃO
# ==============================================================================
def get_supabase_credentials() -> tuple[Optional[str], Optional[str]]:
    """Obtém SUPABASE_URL e SUPABASE_KEY via st.secrets ou variáveis de ambiente."""
    url_val: Optional[str] = None
    key_val: Optional[str] = None
    try:
        if "SUPABASE_URL" in st.secrets:
            url_val = str(st.secrets["SUPABASE_URL"]).strip().rstrip("/")
        if "SUPABASE_KEY" in st.secrets:
            key_val = str(st.secrets["SUPABASE_KEY"]).strip()
    except Exception:
        pass
    if not url_val:
        url_val = os.environ.get("SUPABASE_URL", "").strip().rstrip("/") or None
    if not key_val:
        key_val = os.environ.get("SUPABASE_KEY", "").strip() or None
    return url_val, key_val


def _serialize_calc_for_cloud(calc: Dict[str, Any]) -> Dict[str, Any]:
    """Converte Decimals e datas do cálculo para JSON seguro no Supabase."""
    def _conv(obj: Any) -> Any:
        if isinstance(obj, Decimal):
            return str(obj)
        if isinstance(obj, (date, datetime)):
            return obj.isoformat()
        if isinstance(obj, dict):
            return {k: _conv(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [_conv(i) for i in obj]
        return obj

    return _conv(calc)


def _deserialize_calc_from_cloud(raw_calc: Dict[str, Any]) -> Dict[str, Any]:
    """Restaura Decimals monetários de um cálculo carregado do Supabase."""
    calc = dict(raw_calc)
    for m_key in ("bruto_atualizado", "abatimentos", "liquido_devido", "base_remuneracao"):
        if m_key in calc:
            calc[m_key] = d_money(calc[m_key])
    verbas_restored: List[Dict[str, Any]] = []
    for v in (calc.get("verbas") or []):
        v_copy = dict(v)
        for col in ("Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"):
            if col in v_copy:
                v_copy[col] = d_money(v_copy[col])
        verbas_restored.append(v_copy)
    calc["verbas"] = verbas_restored
    return calc


def supabase_load_all_cases() -> tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, str]]], bool]:
    """Carrega todos os clientes e históricos de chat salvos na nuvem Supabase."""
    sb_url, sb_key = get_supabase_credentials()
    if not sb_url or not sb_key:
        return [], {}, False
    try:
        endpoint = f"{sb_url}/rest/v1/ed_calculos_clientes?select=*&order=updated_at.desc"
        req = urllib.request.Request(
            endpoint,
            headers={
                "apikey": sb_key,
                "Authorization": f"Bearer {sb_key}",
                "Accept": "application/json",
            },
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            rows = json.loads(resp.read().decode("utf-8"))
        calcs_out: List[Dict[str, Any]] = []
        chats_out: Dict[str, List[Dict[str, str]]] = {}
        for row in (rows or []):
            payload = row.get("payload_json")
            if isinstance(payload, str):
                payload = json.loads(payload)
            if isinstance(payload, dict) and payload.get("audit_id"):
                restored = _deserialize_calc_from_cloud(payload)
                calcs_out.append(restored)
                cid = str(restored["audit_id"])
                ch = row.get("chat_history_json") or []
                if isinstance(ch, str):
                    ch = json.loads(ch)
                chats_out[cid] = ch if isinstance(ch, list) else []
        return calcs_out, chats_out, True
    except Exception:
        return [], {}, False


def supabase_upsert_case(calc: Dict[str, Any], chat_history: List[Dict[str, str]]) -> bool:
    """Salva ou atualiza (upsert) um cliente e seu histórico de chat no Supabase."""
    sb_url, sb_key = get_supabase_credentials()
    if not sb_url or not sb_key:
        return False
    try:
        client_key = f"{str(calc.get('worker_name', '')).strip().lower()}|{str(calc.get('employer_name', '')).strip().lower()}"
        body = {
            "client_key": client_key,
            "audit_id": str(calc["audit_id"]),
            "worker_name": str(calc.get("worker_name") or "Cliente"),
            "employer_name": str(calc.get("employer_name") or "Empresa"),
            "liquido_devido": float(d_money(calc.get("liquido_devido", Decimal("0.00")))),
            "payload_json": _serialize_calc_for_cloud(calc),
            "chat_history_json": chat_history,
            "updated_at": datetime.utcnow().isoformat() + "Z",
        }
        endpoint = f"{sb_url}/rest/v1/ed_calculos_clientes?on_conflict=client_key"
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
            headers={
                "apikey": sb_key,
                "Authorization": f"Bearer {sb_key}",
                "Content-Type": "application/json",
                "Prefer": "resolution=merge-duplicates,return=minimal",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            return resp.status in (200, 201, 204)
    except Exception:
        return False


if "cloud_loaded_once" not in st.session_state:
    cloud_calcs, cloud_chats, cloud_ok = supabase_load_all_cases()
    st.session_state.cloud_connected = cloud_ok
    st.session_state.calculations_history = cloud_calcs
    st.session_state.case_chat_threads = {"caso_atual": [], **cloud_chats}
    st.session_state.cloud_loaded_once = True

if "calculations_history" not in st.session_state:
    st.session_state.calculations_history = []
if "timecard_result" not in st.session_state:
    st.session_state.timecard_result = None
if "extracted_ai_data" not in st.session_state:
    st.session_state.extracted_ai_data = None
if "extraction_confirmed" not in st.session_state:
    st.session_state.extraction_confirmed = False
if "case_chat_threads" not in st.session_state:
    st.session_state.case_chat_threads = {"caso_atual": []}
if "selected_chat_case_id" not in st.session_state:
    st.session_state.selected_chat_case_id = "caso_atual"
if "cloud_connected" not in st.session_state:
    st.session_state.cloud_connected = False


# ==============================================================================
# 4. APURADOR DETERMINÍSTICO DE CARTÃO DE PONTO (EM MINUTOS INTEIROS)
# ==============================================================================
def calc_night_minutes(start_min: int, end_min: int) -> int:
    """Calcula minutos trabalhados no período noturno urbano (22h às 05h)."""
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
    """Apura horas extras 50%/100%, intrajornada suprimida e redução ficta noturna."""
    rows_out: List[Dict[str, Any]] = []
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
                "Data": dt_str,
                "Dia": dia_sem,
                "HE 50% (h)": 0.0,
                "HE 100% (h)": 0.0,
                "Intrajornada (h)": 0.0,
                "Noturna Ficta (h)": 0.0,
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
# 5. ANÁLISE CRUZADA POR GEMINI (DOCUMENTOS + FOTOS + RELATO LIVRE)
# ==============================================================================
class PaidItem(BaseModel):
    verba: str
    valor: float = 0.0


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
    adicional_percentual_sugerido: Optional[float] = None
    nr_technical_assessment: Optional[str] = None
    multa_fgts_percentual_sugerido: Optional[float] = None
    meses_sem_registro_ctps: Optional[int] = None
    plus_salarial_acumulo_pct: Optional[float] = None
    diferenca_equiparacao_mensal: Optional[float] = None
    periodos_ferias_vencidas: Optional[int] = None
    restituicao_atestados_ou_descontos: Optional[float] = None
    fgts_nao_depositado_contrato: Optional[float] = None
    meses_estabilidade_restantes: Optional[int] = None
    sugerir_saldo_salario: bool = True
    sugerir_aviso_previo: bool = True
    sugerir_13_e_ferias: bool = True
    sugerir_fgts_e_multa: bool = True
    sugerir_multa_477: bool = False
    sugerir_multa_467: bool = False
    trct_paid_items: List[PaidItem] = Field(default_factory=list)
    irregularities_found: List[IrregularityFinding] = Field(default_factory=list)
    medical_leave_summary: Optional[str] = None
    stability_alert: Optional[str] = None
    viability_conclusion: Optional[str] = None
    petition_draft_topics: Optional[str] = None
    recommended_claims_checklist: List[str] = Field(default_factory=list)
    ambiguities: List[str] = Field(default_factory=list)


def get_gemini_api_key() -> Optional[str]:
    """Obtém a GEMINI_API_KEY via st.secrets ou variáveis de ambiente."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            val = str(st.secrets["GEMINI_API_KEY"]).strip()
            if val:
                return val
    except Exception:
        pass
    env_val = os.environ.get("GEMINI_API_KEY", "").strip()
    return env_val or None


def generate_gemini_with_fallback(client: Any, contents: List[Any], config: Any) -> Any:
    """Chama o modelo Gemini ativo com descoberta dinâmica via API e fallback automático."""
    models_to_try = ["gemini-3-flash-preview", "gemini-2.5-pro", "gemini-2.5-flash"]
    try:
        for m in client.models.list():
            m_name = str(getattr(m, "name", "") or "").replace("models/", "")
            if "gemini" in m_name and "embedding" not in m_name and "imagen" not in m_name and m_name not in models_to_try:
                models_to_try.append(m_name)
    except Exception:
        pass

    last_err: Optional[Exception] = None
    for model_name in models_to_try:
        try:
            return client.models.generate_content(
                model=model_name,
                contents=contents,
                config=config,
            )
        except Exception as exc:
            last_err = exc
            err_s = str(exc)
            if "404" in err_s or "NOT_FOUND" in err_s or "no longer available" in err_s:
                continue
            raise
    if last_err:
        raise last_err
    raise RuntimeError("Nenhum modelo Gemini disponível.")


def analyze_case_with_gemini(lawyer_narrative: str, files_payload: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Executa análise jurídica, técnica (NR-15/NR-16) e documental com Gemini."""
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
            "1) Extrair datas (incluindo início real se houve período sem registro), salário-base, comissões/'por fora', jornada habitual, "
            "valores já pagos no TRCT, descontos indevidos (ex.: atestados médicos descontados), saldo de FGTS não depositado no contrato "
            "(fgts_nao_depositado_contrato), meses sem registro na CTPS (meses_sem_registro_ctps), percentual de plus salarial por acúmulo/desvio "
            "de função (plus_salarial_acumulo_pct), diferença mensal por equiparação salarial (diferenca_equiparacao_mensal), "
            "períodos de férias vencidas não gozadas (periodos_ferias_vencidas) e meses de estabilidade provisória violados (meses_estabilidade_restantes).\n"
            "2) AVALIAÇÃO TÉCNICA DE INSALUBRIDADE (NR-15) E PERICULOSIDADE (NR-16): examine fotos do local de trabalho, produtos químicos, "
            "câmara fria, ruído, agentes biológicos, inflamáveis, alta tensão, explosivos ou uso de motocicleta (art. 193 §4º CLT), "
            "bem como PPP/LTCAT/contracheques. Preencha 'adicional_sugerido' com uma das opções exatas: "
            "'Nenhum', 'Insalubridade (sobre Salário Mínimo)' ou 'Periculosidade (sobre Salário-Base)', "
            "preencha 'adicional_percentual_sugerido' com o número do percentual (ex.: 10, 20, 40 ou 30; 0 se Nenhum), "
            "e detalhe obrigatoriamente em 'nr_technical_assessment' o diagnóstico explícito (ex.: 'IDENTIFICADA INSALUBRIDADE 20% - NR-15 Anexo ...' "
            "ou 'NÃO IDENTIFICADO AGENTE INSALUBRE/PERICULOSO'), o que foi observado nas fotos/relato e a necessidade de perícia judicial (Art. 195 CLT).\n"
            "3) Identificar irregularidades, nulidades e brechas jurídicas (estabilidade B91, gestante, CIPA, atestados descontados, "
            "ASO inapto, rescisão indireta, multa art. 477 §8º CLT, multa art. 467 CLT, horas extras, intrajornada, vínculo sem registro, "
            "acúmulo/desvio de função, equiparação salarial), além de 'multa_fgts_percentual_sugerido' (ex.: 40.0 na dispensa sem justa causa, "
            "20.0 no acordo, ou 10.0 se o relato/TRCT mostrar que já pagaram 30% e falta a diferença).\n"
            "sugerir_multa_477 e sugerir_multa_467.\n"
            "6) Nunca inventar valores ou datas ausentes; sinalize dúvidas em 'ambiguities'."
        )
    ]

    if lawyer_narrative.strip():
        contents.append(f"=== RELATO DO ADVOGADO / FATOS NARRADOS PELO CLIENTE ===\n{lawyer_narrative.strip()}")

    for item in files_payload:
        contents.append(f"=== DOCUMENTO ANEXADO: {item['name']} ===")
        contents.append(types.Part.from_bytes(data=item["bytes"], mime_type=item["mime"]))

    response = generate_gemini_with_fallback(
        client=client,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ExtractedCaseData,
            temperature=0.0,
        ),
    )

    parsed: Dict[str, Any] = {}
    if getattr(response, "parsed", None) is not None:
        if hasattr(response.parsed, "model_dump"):
            parsed = response.parsed.model_dump()
        elif isinstance(response.parsed, dict):
            parsed = response.parsed
    if not parsed and getattr(response, "text", None):
        parsed = json.loads(response.text)
    if not isinstance(parsed, dict):
        raise RuntimeError("Resposta estruturada do Gemini vazia ou inválida.")
    return parsed


def chat_with_case_strategist(
    user_message: str,
    chat_history: List[Dict[str, str]],
    lawyer_narrative: str,
    extracted_data: Dict[str, Any],
    latest_calc: Optional[Dict[str, Any]] = None,
    all_saved_cases: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """Chat jurídico multi-turno real com o Gemini (memória contínua + acesso a todos os clientes da sessão)."""
    api_key = get_gemini_api_key()
    if not api_key or not GENAI_AVAILABLE:
        raise RuntimeError("GEMINI_API_KEY não configurada em Secrets ou google-genai ausente.")

    client = genai.Client(api_key=api_key)

    system_blocks = [
        "Você é um Advogado Trabalhista Sênior, Parecerista Jurídico e Perito em Cálculos Trabalhistas no Brasil (CLT, Reforma Trabalhista, Constituição Federal, Súmulas e OJs do TST, NRs 15 e 16).",
        "Você está conversando em tempo real com o advogado titular dentro do sistema 'ED Cálculos Trabalhistas'.",
        "Responda sempre no mais alto nível técnico-jurídico, de forma natural, interativa, clara e resolutiva — exatamente como um sócio sênior debatendo estratégia processual.",
        "IMPORTANTE: O advogado pode ter analisado vários clientes nesta sessão (ex.: Pedro, João, Maria). Você tem acesso completo ao CASO SELECIONADO NO CHAT e também ao DOSSIÊ DE TODOS OS OUTROS CLIENTES ANALISADOS NA SESSÃO, podendo responder perguntas sobre qualquer um deles a qualquer momento.",
        "Você pode:\n"
        "- Discutir teses, ônus da prova (Art. 818 CLT), riscos de sucumbência e jurisprudência atualizada do TST/TRTs;\n"
        "- Incorporar novos fatos que o advogado trouxer na conversa e explicar o impacto jurídico e financeiro;\n"
        "- Redigir tópicos completos da Petição Inicial (Dos Fatos, Do Direito e Dos Pedidos), quesitos periciais (insalubridade/periculosidade) ou perguntas para audiência de instrução;\n"
        "- Responder dúvidas sobre o cliente selecionado ou sobre qualquer outro cliente já analisado na sessão.",
        f"\n=== CASO EM FOCO SELECIONADO NO CHAT ===\nRelato: {lawyer_narrative.strip() or '(Nenhum relato preenchido ainda)'}",
    ]

    if extracted_data:
        safe_ext = {k: v for k, v in extracted_data.items() if v not in (None, "", [], {})}
        system_blocks.append(f"\n=== PARECER E DADOS EXTRAÍDOS DO CASO EM FOCO ===\n{json.dumps(safe_ext, ensure_ascii=False, default=str)}")

    if latest_calc:
        verbas_resumo = [
            f"- {v.get('Verba')}: {fmt_brl(v.get('Total Apurado', Decimal('0.00')))}"
            for v in (latest_calc.get("verbas") or [])
        ]
        system_blocks.append(
            "\n=== LIQUIDAÇÃO DO CASO EM FOCO ===\n"
            f"Cliente: {latest_calc.get('worker_name')} x {latest_calc.get('employer_name')} (ID: {latest_calc.get('audit_id')})\n"
            f"Cargo: {latest_calc.get('role_title', '')} | Período: {latest_calc.get('admission_date', '')} a {latest_calc.get('dismissal_date', '')}\n"
            f"Total Bruto: {fmt_brl(latest_calc.get('bruto_atualizado', Decimal('0.00')))} | "
            f"Abatimentos TRCT: {fmt_brl(latest_calc.get('abatimentos', Decimal('0.00')))} | "
            f"Valor da Causa: {fmt_brl(latest_calc.get('liquido_devido', Decimal('0.00')))}\n"
            + "\n".join(verbas_resumo)
        )

    if all_saved_cases:
        outros_resumos: List[str] = []
        for c_item in all_saved_cases:
            v_list = ", ".join(
                f"{vv.get('Verba')} ({fmt_brl(vv.get('Total Apurado', Decimal('0.00')))})"
                for vv in (c_item.get("verbas") or [])
            )
            op_item = c_item.get("ai_opinion") or {}
            outros_resumos.append(
                f"* Cliente: {c_item.get('worker_name')} x {c_item.get('employer_name')} [{c_item.get('audit_id')}] | "
                f"Período: {c_item.get('admission_date')} a {c_item.get('dismissal_date')} | "
                f"Valor da Causa: {fmt_brl(c_item.get('liquido_devido', Decimal('0.00')))}\n"
                f"  Relato: {str(c_item.get('lawyer_narrative') or '')[:600]}\n"
                f"  Diagnóstico NR/Parecer: {op_item.get('nr_technical_assessment') or ''} | {op_item.get('viability_conclusion') or ''}\n"
                f"  Verbas: {v_list}"
            )
        system_blocks.append("\n=== TODOS OS CLIENTES ANALISADOS NESTA SESSÃO (MEMÓRIA MULTI-CLIENTE) ===\n" + "\n".join(outros_resumos))

    sdk_history: List[Any] = []
    for msg in chat_history[-16:]:
        role_sdk = "user" if msg.get("role") == "user" else "model"
        txt_sdk = str(msg.get("content") or "").strip()
        if txt_sdk:
            sdk_history.append(
                types.Content(
                    role=role_sdk,
                    parts=[types.Part.from_text(text=txt_sdk)],
                )
            )

    models_to_try = ["gemini-3-flash-preview", "gemini-2.5-pro", "gemini-2.5-flash"]
    try:
        for m in client.models.list():
            m_name = str(getattr(m, "name", "") or "").replace("models/", "")
            if "gemini" in m_name and "embedding" not in m_name and "imagen" not in m_name and m_name not in models_to_try:
                models_to_try.append(m_name)
    except Exception:
        pass

    last_err: Optional[Exception] = None
    for model_name in models_to_try:
        try:
            chat_session = client.chats.create(
                model=model_name,
                history=sdk_history,
                config=types.GenerateContentConfig(
                    system_instruction="\n".join(system_blocks),
                    temperature=0.35,
                ),
            )
            response = chat_session.send_message(user_message)
            reply_txt = (getattr(response, "text", None) or "").strip()
            if reply_txt:
                return reply_txt
        except Exception as exc:
            last_err = exc
            err_s = str(exc)
            if "404" in err_s or "NOT_FOUND" in err_s or "no longer available" in err_s:
                continue
            raise

    if last_err:
        raise last_err
    return "Não foi possível obter resposta do Gemini no momento."


# ==============================================================================
# 6. MOTOR DE CÁLCULO DECIMAL EXATO E EXPORTAÇÃO XLSX / PDF
# ==============================================================================
def run_labor_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
    """Executa a liquidação trabalhista determinística em Decimal exato."""
    adm_dt: date = params["admission_date"]
    dism_dt: date = params["dismissal_date"]
    term_type: str = str(params.get("termination_type") or "sem_justa_causa")
    divisor = max(1, f_int(params.get("monthly_divisor"), 220))

    base_sal = d_money(params.get("base_salary"))
    commissions = d_money(params.get("commissions"))
    insal_val = d_money(params.get("insalubrity"))
    plus_pct = Decimal(str(f_float(params.get("plus_acumulo_pct"), 0.0))) / Decimal("100")
    dif_equip = d_money(params.get("dif_equiparacao_mensal"))

    plus_mensal = (base_sal * plus_pct).quantize(CENT, rounding=ROUND_HALF_UP)
    rem_base = base_sal + commissions + insal_val + plus_mensal + dif_equip
    hora_normal = (rem_base / Decimal(str(divisor))).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    verbas: List[Dict[str, Any]] = []
    period_str = f"{adm_dt.strftime('%d/%m/%Y')} a {dism_dt.strftime('%d/%m/%Y')}"

    def add_verba(
        cod: str,
        nome: str,
        base_calc: Decimal,
        formula: str,
        qtd: str,
        pct: str,
        valor_orig: Decimal,
        reflexos: Decimal = Decimal("0.00"),
        inc_fgts: bool = True,
    ) -> None:
        orig_q = valor_orig.quantize(CENT, rounding=ROUND_HALF_UP)
        refl_q = reflexos.quantize(CENT, rounding=ROUND_HALF_UP)
        tot_q = (orig_q + refl_q).quantize(CENT, rounding=ROUND_HALF_UP)
        if tot_q <= Decimal("0.00"):
            return
        verbas.append({
            "Código": cod,
            "Verba": nome,
            "Período": period_str,
            "Base de Cálculo": base_calc.quantize(CENT, rounding=ROUND_HALF_UP),
            "Fórmula": formula,
            "Quantidade": qtd,
            "Percentual": pct,
            "Valor Original": orig_q,
            "Reflexos": refl_q,
            "Total Apurado": tot_q,
            "Incide FGTS": inc_fgts,
        })

    # 1. Saldo de Salário
    if params.get("inc_saldo_salario"):
        dias_s = max(0, min(31, f_int(params.get("saldo_dias"), dism_dt.day)))
        val_saldo = (rem_base / Decimal("30") * Decimal(str(dias_s))).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-01", "Saldo de Salário", rem_base, "(Remuneração / 30) × Dias", f"{dias_s} dias", "100%", val_saldo, inc_fgts=True)

    # 2. Aviso-Prévio Indenizado (Lei 12.506/2011)
    anos_completos = max(0, (dism_dt - adm_dt).days // 365)
    dias_aviso = min(90, 30 + anos_completos * 3)
    if params.get("inc_aviso_previo") and term_type in ("sem_justa_causa", "rescisao_indireta", "justa_causa", "acordo"):
        fator_aviso = Decimal("0.5") if term_type == "acordo" else Decimal("1.0")
        val_aviso = (rem_base / Decimal("30") * Decimal(str(dias_aviso)) * fator_aviso).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-02", f"Aviso-Prévio Indenizado ({dias_aviso} dias)", rem_base, "(Remuneração / 30) × Dias Aviso", f"{dias_aviso} dias", "50%" if term_type == "acordo" else "100%", val_aviso, inc_fgts=True)

    # 3. 13º Salário Proporcional e Férias Proporcionais + 1/3 + Vencidas
    meses_ano = dism_dt.month if dism_dt.day >= 15 else max(0, dism_dt.month - 1)
    if params.get("inc_13") and meses_ano > 0:
        val_13 = (rem_base / Decimal("12") * Decimal(str(meses_ano))).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-03A", "13º Salário Proporcional (c/ Projeção)", rem_base, "(Remuneração / 12) × Avos", f"{meses_ano}/12 avos", "100%", val_13, inc_fgts=True)

    if params.get("inc_ferias"):
        avos_ferias = max(1, ((dism_dt.year - adm_dt.year) * 12 + (dism_dt.month - adm_dt.month)) % 12 or 12)
        val_ferias_prop = (rem_base / Decimal("12") * Decimal(str(avos_ferias))).quantize(CENT, rounding=ROUND_HALF_UP)
        terco_prop = (val_ferias_prop / Decimal("3")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-03B", "Férias Proporcionais + 1/3 Constitucional (c/ Projeção)", rem_base, "((Remuneração / 12) × Avos) + 1/3", f"{avos_ferias}/12 avos", "133,33%", val_ferias_prop, terco_prop, inc_fgts=False)

    ferias_venc = max(0, f_int(params.get("periodos_ferias_vencidas"), 0))
    if ferias_venc > 0:
        val_fv = (rem_base * Decimal(str(ferias_venc)) * Decimal("2")).quantize(CENT, rounding=ROUND_HALF_UP)
        terco_fv = (val_fv / Decimal("3")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-03C", "Férias Vencidas em Dobro + 1/3 (Art. 137 CLT)", rem_base, "(Remuneração × 2) × Períodos + 1/3", f"{ferias_venc} per.", "200% + 1/3", val_fv, terco_fv, inc_fgts=False)

    # 4. Adicional de Insalubridade / Periculosidade Retroativo
    meses_contrato = max(1, min(60, (dism_dt.year - adm_dt.year) * 12 + (dism_dt.month - adm_dt.month) + 1))
    fator_refl_hab = Decimal("0.225") if params.get("inc_reflexos_completos") else Decimal("0.00")

    if params.get("inc_adicional_retroativo") and insal_val > Decimal("0.00"):
        orig_adic = (insal_val * Decimal(str(meses_contrato))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_adic = (orig_adic * fator_refl_hab).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba(
            "VER-04B",
            f"Adicional Retroativo — {params.get('adicional_label', 'Insalubridade/Periculosidade')}",
            insal_val,
            "Adicional Mensal × Meses Imprescritos + Reflexos",
            f"{meses_contrato} meses",
            str(params.get("adicional_pct_label", "")),
            orig_adic,
            refl_adic,
            inc_fgts=True,
        )

    # 5. Horas Extras 50%, 100%, Intrajornada e Adicional Noturno
    if params.get("inc_he50") and f_float(params.get("he50_hours"), 0.0) > 0:
        h50 = Decimal(str(f_float(params.get("he50_hours"), 0.0)))
        orig_he50 = (hora_normal * Decimal("1.50") * h50).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he50 = (orig_he50 * (Decimal("0.1667") + fator_refl_hab)).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05A", "Horas Extras 50% + DSR e Reflexos", hora_normal, "(Hora Normal × 1,50) × Horas + DSR/Reflexos", f"{h50:.2f}h", "+50%", orig_he50, refl_he50, inc_fgts=True)

    if params.get("inc_he100") and f_float(params.get("he100_hours"), 0.0) > 0:
        h100 = Decimal(str(f_float(params.get("he100_hours"), 0.0)))
        orig_he100 = (hora_normal * Decimal("2.00") * h100).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he100 = (orig_he100 * (Decimal("0.1667") + fator_refl_hab)).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05B", "Horas Extras 100% (Domingos/Feriados) + Reflexos", hora_normal, "(Hora Normal × 2,00) × Horas + Reflexos", f"{h100:.2f}h", "+100%", orig_he100, refl_he100, inc_fgts=True)

    if params.get("inc_intrajornada") and f_float(params.get("intrajornada_hours"), 0.0) > 0:
        hintra = Decimal(str(f_float(params.get("intrajornada_hours"), 0.0)))
        orig_intra = (hora_normal * Decimal("1.50") * hintra).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05C", "Intervalo Intrajornada Suprimido (+50% — Art. 71, § 4º CLT)", hora_normal, "(Hora Normal × 1,50) × Horas Suprimidas", f"{hintra:.2f}h", "+50%", orig_intra, inc_fgts=False)

    if params.get("inc_noturno") and f_float(params.get("night_hours"), 0.0) > 0:
        hnot = Decimal(str(f_float(params.get("night_hours"), 0.0)))
        orig_not = (hora_normal * Decimal("0.20") * hnot).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_not = (orig_not * (Decimal("0.1667") + fator_refl_hab)).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06", "Adicional Noturno (20%) + Hora Ficta e Reflexos", hora_normal, "(Hora Normal × 0,20) × Horas Fictas + Reflexos", f"{hnot:.2f}h", "+20%", orig_not, refl_not, inc_fgts=True)

    # 6. Período sem registro CTPS, Acúmulo de Função, Equiparação e Estabilidade
    meses_sem_reg = max(0, f_int(params.get("meses_sem_registro"), 0))
    if meses_sem_reg > 0:
        verbas_sem_reg = (rem_base * Decimal("0.272") * Decimal(str(meses_sem_reg))).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-07A", "Reconhecimento de Vínculo / Período Sem Registro CTPS (13º, Férias+1/3 e FGTS)", rem_base, "(Remuneração × 27,2%) × Meses Sem Registro", f"{meses_sem_reg} meses", "27,2%", verbas_sem_reg, inc_fgts=False)

    if plus_mensal > Decimal("0.00"):
        orig_plus = (plus_mensal * Decimal(str(meses_contrato))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_plus = (orig_plus * fator_refl_hab).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-07B", "Plus Salarial por Acúmulo / Desvio de Função + Reflexos", base_sal, "(Salário × %) × Meses + Reflexos", f"{meses_contrato} meses", f"{f_float(params.get('plus_acumulo_pct'), 0.0):g}%", orig_plus, refl_plus, inc_fgts=True)

    if dif_equip > Decimal("0.00"):
        orig_eq = (dif_equip * Decimal(str(meses_contrato))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_eq = (orig_eq * fator_refl_hab).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-07C", "Diferenças Salariais por Equiparação (Art. 461 CLT) + Reflexos", dif_equip, "Diferença Mensal × Meses + Reflexos", f"{meses_contrato} meses", "100%", orig_eq, refl_eq, inc_fgts=True)

    meses_estab = max(0, f_int(params.get("meses_estabilidade"), 0))
    if meses_estab > 0:
        ind_estab = (rem_base * Decimal(str(meses_estab)) * Decimal("1.272")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-08", "Indenização Estabilidade Provisória Violada (Salários + 13º + Férias + FGTS)", rem_base, "(Remuneração × 1,272) × Meses Restantes", f"{meses_estab} meses", "100%", ind_estab, inc_fgts=False)

    # 7. FGTS (8%) + Multa Rescisória configurável (%)
    if params.get("inc_fgts"):
        base_fgts_verbas = sum((v["Total Apurado"] for v in verbas if v.get("Incide FGTS")), Decimal("0.00"))
        fgts_sobre_verbas = (base_fgts_verbas * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP)
        fgts_faltante = d_money(params.get("fgts_nao_depositado"))
        fgts_contrato_estimado = (rem_base * Decimal(str(meses_contrato)) * Decimal("0.08")).quantize(CENT, rounding=ROUND_HALF_UP)
        base_para_multa = fgts_contrato_estimado + fgts_sobre_verbas + fgts_faltante
        multa_pct = Decimal(str(f_float(params.get("fgts_multa_pct"), 40.0))) / Decimal("100")
        multa_fgts_val = (base_para_multa * multa_pct).quantize(CENT, rounding=ROUND_HALF_UP)
        total_fgts_item = fgts_sobre_verbas + fgts_faltante
        if total_fgts_item > Decimal("0.00") or multa_fgts_val > Decimal("0.00"):
            add_verba(
                "VER-09",
                f"FGTS (8% Verbas + Depósitos Faltantes) + Multa Rescisória ({f_float(params.get('fgts_multa_pct'), 40.0):g}%)",
                base_para_multa,
                f"FGTS 8% + Multa {f_float(params.get('fgts_multa_pct'), 40.0):g}% sobre Saldo Total FGTS",
                f"{meses_contrato} meses",
                f"{f_float(params.get('fgts_multa_pct'), 40.0):g}%",
                total_fgts_item,
                multa_fgts_val,
                inc_fgts=False,
            )

    # 8. Multas Arts. 477 e 467 da CLT
    if params.get("inc_multa_477"):
        add_verba("VER-10A", "Multa por Atraso na Quitação Rescisória (Art. 477, § 8º CLT)", rem_base, "1 Remuneração Mensal", "1 mês", "100%", rem_base, inc_fgts=False)

    if params.get("inc_multa_467"):
        base_467 = sum((v["Total Apurado"] for v in verbas if v["Código"] in ("VER-01", "VER-02", "VER-03A", "VER-03B")), Decimal("0.00"))
        multa_467_val = (base_467 * Decimal("0.50")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-10B", "Multa de 50% sobre Verbas Rescisórias Incontroversas (Art. 467 CLT)", base_467, "50% × Verbas Rescisórias", "1", "50%", multa_467_val, inc_fgts=False)

    bruto_total = sum((v["Total Apurado"] for v in verbas), Decimal("0.00")).quantize(CENT, rounding=ROUND_HALF_UP)
    abatimentos = d_money(params.get("offsets"))
    liquido = max(Decimal("0.00"), bruto_total - abatimentos).quantize(CENT, rounding=ROUND_HALF_UP)

    raw_sig = f"{params.get('worker_name')}|{params.get('employer_name')}|{adm_dt}|{dism_dt}|{liquido}"
    audit_hash = "ED-" + hashlib.sha256(raw_sig.encode("utf-8")).hexdigest()[:8].upper()

    return {
        "audit_id": audit_hash,
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "office_header": str(params.get("office_header") or "Escritório de Advocacia Trabalhista"),
        "worker_name": str(params.get("worker_name") or "Cliente"),
        "employer_name": str(params.get("employer_name") or "Empresa"),
        "role_title": str(params.get("role_title") or ""),
        "admission_date": adm_dt.strftime("%d/%m/%Y"),
        "dismissal_date": dism_dt.strftime("%d/%m/%Y"),
        "termination_type": term_type,
        "base_remuneracao": rem_base,
        "verbas": verbas,
        "bruto_atualizado": bruto_total,
        "abatimentos": abatimentos,
        "liquido_devido": liquido,
        "ai_opinion": params.get("ai_opinion") or {},
        "lawyer_narrative": str(params.get("lawyer_narrative") or ""),
    }


def build_complete_excel_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera planilha Excel (.XLSX) com Resumo, Verbas Liquidadas e Parecer Técnico."""
    output = io.BytesIO()
    resumo_rows = [
        {"Campo": "Escritório / Timbre", "Valor": calc.get("office_header", "")},
        {"Campo": "Código de Auditoria", "Valor": calc.get("audit_id", "")},
        {"Campo": "Data/Hora Emissão", "Valor": calc.get("generated_at", "")},
        {"Campo": "Reclamante (Cliente)", "Valor": calc.get("worker_name", "")},
        {"Campo": "Reclamada (Empresa)", "Valor": calc.get("employer_name", "")},
        {"Campo": "Cargo / Função", "Valor": calc.get("role_title", "")},
        {"Campo": "Período Contratual", "Valor": f"{calc.get('admission_date')} a {calc.get('dismissal_date')}"},
        {"Campo": "Remuneração Base Apurada", "Valor": fmt_brl(calc.get("base_remuneracao", Decimal("0.00")))},
        {"Campo": "Total Bruto das Verbas", "Valor": fmt_brl(calc.get("bruto_atualizado", Decimal("0.00")))},
        {"Campo": "(-) Abatimentos TRCT", "Valor": fmt_brl(calc.get("abatimentos", Decimal("0.00")))},
        {"Campo": "VALOR TOTAL DA CAUSA", "Valor": fmt_brl(calc.get("liquido_devido", Decimal("0.00")))},
    ]
    df_resumo = pd.DataFrame(resumo_rows)

    verbas_export = []
    for v in (calc.get("verbas") or []):
        verbas_export.append({
            "Código": v.get("Código"),
            "Verba": v.get("Verba"),
            "Período": v.get("Período"),
            "Quantidade": v.get("Quantidade"),
            "Percentual": v.get("Percentual"),
            "Base de Cálculo (R$)": float(d_money(v.get("Base de Cálculo"))),
            "Fórmula": v.get("Fórmula"),
            "Valor Original (R$)": float(d_money(v.get("Valor Original"))),
            "Reflexos (R$)": float(d_money(v.get("Reflexos"))),
            "Total Apurado (R$)": float(d_money(v.get("Total Apurado"))),
        })
    df_verbas = pd.DataFrame(verbas_export)

    op = calc.get("ai_opinion") or {}
    parecer_rows = [
        {"Tópico": "Diagnóstico NR-15 / NR-16", "Conteúdo": str(op.get("nr_technical_assessment") or "Não aplicável")},
        {"Tópico": "Alerta de Estabilidade", "Conteúdo": str(op.get("stability_alert") or "Nenhum")},
        {"Tópico": "Atestados / ASO / INSS", "Conteúdo": str(op.get("medical_leave_summary") or "Nenhum")},
        {"Tópico": "Conclusão de Viabilidade", "Conteúdo": str(op.get("viability_conclusion") or "")},
        {"Tópico": "Relato do Caso", "Conteúdo": str(calc.get("lawyer_narrative") or "")},
    ]
    df_parecer = pd.DataFrame(parecer_rows)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo_Causa", index=False)
        df_verbas.to_excel(writer, sheet_name="Memoria_Verbas", index=False)
        df_parecer.to_excel(writer, sheet_name="Parecer_Tecnico", index=False)

    return output.getvalue()


def build_printable_dossier_pdf_bytes(calc: Dict[str, Any]) -> bytes:
    """Gera relatório oficial em PDF (fpdf2) com timbre, memória de cálculo e parecer."""
    def _latin(txt: Any) -> str:
        return str(txt or "").encode("latin-1", errors="replace").decode("latin-1")

    pdf = FPDF()
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    usable_w = pdf.epw

    def _safe_multi(text: str, line_h: float = 5.0) -> None:
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(usable_w, line_h, _latin(text))

    pdf.set_font("Helvetica", "B", 13)
    _safe_multi(calc.get("office_header", "ED Cálculos Trabalhistas"), 7)
    pdf.set_font("Helvetica", "", 9)
    _safe_multi(f"Dossiê Pré-Processual e Liquidação (Art. 840, § 1º CLT) | ID: {calc.get('audit_id')} | Emissão: {calc.get('generated_at')}", 5)
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    _safe_multi(f"Reclamante: {calc.get('worker_name')}  |  Reclamada: {calc.get('employer_name')}", 6)
    pdf.set_font("Helvetica", "", 9)
    _safe_multi(f"Cargo: {calc.get('role_title')}  |  Contrato: {calc.get('admission_date')} a {calc.get('dismissal_date')}  |  Remuneração Base: {fmt_brl(calc.get('base_remuneracao'))}", 5)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 10)
    _safe_multi("QUADRO DE PEDIDOS LIQUIDADOS", 6)
    pdf.set_font("Helvetica", "", 8.5)
    for v in (calc.get("verbas") or []):
        linha = f"[{v.get('Código')}] {v.get('Verba')} ({v.get('Quantidade')}) — Orig: {fmt_brl(v.get('Valor Original'))} | Refl: {fmt_brl(v.get('Reflexos'))} | Total: {fmt_brl(v.get('Total Apurado'))}"
        _safe_multi(linha, 5)

    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 10)
    _safe_multi(f"Total Bruto: {fmt_brl(calc.get('bruto_atualizado'))}   |   (-) Abatimentos TRCT: {fmt_brl(calc.get('abatimentos'))}", 6)
    pdf.set_font("Helvetica", "B", 11)
    _safe_multi(f"VALOR TOTAL ATRIBUÍDO À CAUSA: {fmt_brl(calc.get('liquido_devido'))}", 7)

    op = calc.get("ai_opinion") or {}
    if op.get("nr_technical_assessment") or op.get("viability_conclusion"):
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 10)
        _safe_multi("PARECER TÉCNICO E ESTRATÉGICO", 6)
        pdf.set_font("Helvetica", "", 8.5)
        if op.get("nr_technical_assessment"):
            _safe_multi(f"NR-15 / NR-16: {op.get('nr_technical_assessment')}", 5)
        if op.get("viability_conclusion"):
            _safe_multi(f"Viabilidade: {op.get('viability_conclusion')}", 5)

    raw_pdf = pdf.output()
    if isinstance(raw_pdf, (bytes, bytearray)):
        return bytes(raw_pdf)
    return str(raw_pdf).encode("latin-1", errors="replace")


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
    if st.session_state.get("cloud_connected"):
        st.caption("Nuvem Supabase: `Conectada ☁️`")
    else:
        st.caption("Nuvem Supabase: `Modo Sessão Local`")
    if st.session_state.calculations_history:
        st.divider()
        # Agrupa por cliente/empresa para nunca duplicar o mesmo cliente na lista lateral
        unique_clients_sb: Dict[str, Dict[str, Any]] = {}
        for c_sb in st.session_state.calculations_history:
            c_key = f"{str(c_sb.get('worker_name', '')).strip().lower()}|{str(c_sb.get('employer_name', '')).strip().lower()}"
            if c_key not in unique_clients_sb:
                unique_clients_sb[c_key] = c_sb

        st.caption(f"Clientes analisados ({len(unique_clients_sb)}) — clique para abrir o chat:")
        for idx_sb, c_sb in enumerate(unique_clients_sb.values()):
            cid_sb = str(c_sb["audit_id"])
            msgs_count = len(st.session_state.case_chat_threads.get(cid_sb, []))
            badge_chat = f" 💬{msgs_count}" if msgs_count > 0 else ""
            is_selected_sb = st.session_state.selected_chat_case_id == cid_sb
            btn_prefix = "▶ " if is_selected_sb else ""
            if st.button(
                f"{btn_prefix}{c_sb.get('worker_name', 'Cliente')} ({fmt_brl(c_sb['liquido_devido'])}){badge_chat}",
                key=f"sb_select_case_{cid_sb}_{idx_sb}",
                use_container_width=True,
            ):
                st.session_state.selected_chat_case_id = cid_sb
                st.session_state["chat_case_selector_widget"] = cid_sb
                st.rerun()


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
                gemini_files: List[Dict[str, Any]] = []
                timecard_dfs: List[pd.DataFrame] = []
                for f in (uploaded_all or []):
                    fname_low = f.name.lower()
                    if fname_low.endswith(".csv"):
                        raw_txt = f.getvalue().decode("utf-8", errors="ignore")
                        first_line = raw_txt.splitlines()[0] if raw_txt.splitlines() else ""
                        sep = ";" if ";" in first_line else ","
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
                        ext = "." + fname_low.rsplit(".", 1)[-1] if "." in fname_low else ".txt"
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
                            st.session_state.selected_chat_case_id = "caso_atual"
                            st.session_state.case_chat_threads["caso_atual"] = []
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
            adic_ia_badge = str(ext_pre.get("adicional_sugerido") or "Nenhum")
            pct_ia_badge = f_float(ext_pre.get("adicional_percentual_sugerido"), 0.0)
            nr_txt = str(ext_pre.get("nr_technical_assessment") or "Sem agentes insalubres ou periculosos relatados/anexados.")
            if adic_ia_badge != "Nenhum" and pct_ia_badge > 0:
                st.warning(f"**DIAGNÓSTICO IA — {adic_ia_badge.upper()} ({pct_ia_badge:g}%):** {nr_txt}")
            else:
                st.info(f"**DIAGNÓSTICO IA — INSALUBRIDADE / PERICULOSIDADE (NR-15 / NR-16):** {nr_txt}")
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
    trct_auto_sum = float(round(sum(f_float(i.get("valor"), 0.0) for i in trct_paid_list if isinstance(i, dict)), 2))

    def parse_iso_date(val: Optional[str], fallback: date) -> date:
        if not val:
            return fallback
        try:
            return datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
        except Exception:
            return fallback

    with st.form("preliminary_eval_form"):
        st.subheader("2. Conferência das Condições do Contrato e Pedidos")
        if ext_pre:
            st.success("Campos abaixo preenchidos automaticamente com base no seu relato e nos documentos anexados. Ajuste apenas se desejar alterar algo.")
        else:
            st.caption("Preenchimento inicial de exemplo — clique em 'Analisar Relato + Documentos' acima para preencher automaticamente ou edite diretamente abaixo.")

        office_timbre = st.text_input(
            "Timbre do Escritório / Advogado e OAB (aparece no topo do PDF)",
            value="Escritório de Advocacia Trabalhista — Parecer Pré-Processual",
        )

        p1, p2, p3 = st.columns(3)
        cliente_nome = p1.text_input("Nome do Cliente (Reclamante) *", value=str(ext_pre.get("worker_name") or "João da Silva"))
        empresa_nome = p2.text_input("Empresa (Reclamada) *", value=str(ext_pre.get("employer_name") or "Comércio e Serviços Alfa Ltda."))
        cargo_funcao = p3.text_input("Cargo / Função", value=str(ext_pre.get("role_title") or "Auxiliar Operacional / Vendedor"))

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
        divisor_pre = d4.number_input("Divisor Mensal *", min_value=1, max_value=300, value=max(1, min(300, f_int(ext_pre.get("monthly_divisor"), 220))))

        adicionais_opts = [
            "Nenhum",
            "Insalubridade (sobre Salário Mínimo)",
            "Periculosidade (sobre Salário-Base)",
        ]
        raw_adic_ext = str(ext_pre.get("adicional_sugerido") or "Nenhum")
        if "Insalubridade" in raw_adic_ext:
            adic_ext = "Insalubridade (sobre Salário Mínimo)"
        elif "Periculosidade" in raw_adic_ext:
            adic_ext = "Periculosidade (sobre Salário-Base)"
        else:
            adic_ext = "Nenhum"

        pct_adic_default = f_float(ext_pre.get("adicional_percentual_sugerido"), 0.0)
        if pct_adic_default == 0.0:
            if "10%" in raw_adic_ext:
                pct_adic_default = 10.0
            elif "20%" in raw_adic_ext:
                pct_adic_default = 20.0
            elif "40%" in raw_adic_ext:
                pct_adic_default = 40.0
            elif "30%" in raw_adic_ext:
                pct_adic_default = 30.0

        r1, r2, r3, r4, r5 = st.columns([2, 2, 2.2, 1.5, 2])
        sal_registrado = r1.number_input("Salário-Base Registrado (R$) *", min_value=0.0, value=max(0.0, f_float(ext_pre.get("base_salary"), 2400.00)), step=50.0)
        sal_por_fora = r2.number_input("Comissões / Salário 'Por Fora' (R$/mês)", min_value=0.0, value=max(0.0, f_float(ext_pre.get("commissions"), 0.0)), step=50.0)
        adicional_cond = r3.selectbox("Tipo Adicional (NR-15 / NR-16)", adicionais_opts, index=adicionais_opts.index(adic_ext))
        adicional_pct_custom = r4.number_input("Percentual Adicional (%)", min_value=0.0, max_value=100.0, value=max(0.0, min(100.0, pct_adic_default)), step=5.0, help="Ex.: 10%, 20% ou 40% (Insalubridade) ou 30% (Periculosidade) — ou qualquer % que desejar.")
        valor_ja_pago_trct = r5.number_input("Total Já Pago no TRCT (R$)", min_value=0.0, value=max(0.0, trct_auto_sum), step=100.0)

        if ext_pre and ext_pre.get("nr_technical_assessment"):
            st.caption(f"Parecer técnico da IA sobre Insalubridade/Periculosidade: {ext_pre.get('nr_technical_assessment')}")

        if tc_pre is not None:
            st.info(
                f"Cartão de ponto anexado apurado: **{tc_pre['he50_hours']}h** extras 50% | "
                f"**{tc_pre['he100_hours']}h** extras 100% | **{tc_pre['intrajornada_hours']}h** intrajornada | "
                f"**{tc_pre['night_ficta_hours']}h** noturnas fictas."
            )
            he_semanal_relatada = 0.0
            domingos_mes = 0
            min_intervalo_suprimido = 0
            horas_noturnas_semana = 0.0
        else:
            j1, j2, j3, j4 = st.columns(4)
            he_semanal_relatada = j1.number_input("Horas Extras / Semana (50%)", min_value=0.0, value=max(0.0, f_float(ext_pre.get("weekly_he50_hours"), 0.0)), step=0.5)
            domingos_mes = j2.number_input("Domingos/Feriados sem Folga / Mês", min_value=0, max_value=10, value=max(0, min(10, f_int(ext_pre.get("monthly_sundays_worked"), 0))))
            min_intervalo_suprimido = j3.number_input("Intervalo Suprimido (min/dia)", min_value=0, max_value=120, value=max(0, min(120, f_int(ext_pre.get("suppressed_break_minutes"), 0))), step=5)
            horas_noturnas_semana = j4.number_input("Horas Noturnas / Semana", min_value=0.0, value=max(0.0, f_float(ext_pre.get("weekly_night_hours"), 0.0)), step=0.5)

        t1, t2, t3, t4 = st.columns(4)
        meses_sem_ctps_val = t1.number_input("Meses Sem Registro CTPS", min_value=0, max_value=120, value=max(0, min(120, f_int(ext_pre.get("meses_sem_registro_ctps"), 0))))
        ferias_vencidas_val = t2.number_input("Períodos Férias Vencidas", min_value=0, max_value=5, value=max(0, min(5, f_int(ext_pre.get("periodos_ferias_vencidas"), 0))))
        plus_acumulo_val = t3.number_input("Acúmulo/Desvio Função (%)", min_value=0.0, max_value=100.0, value=max(0.0, min(100.0, f_float(ext_pre.get("plus_salarial_acumulo_pct"), 0.0))), step=5.0)
        dif_equip_val = t4.number_input("Equiparação Salarial (R$/mês)", min_value=0.0, value=max(0.0, f_float(ext_pre.get("diferenca_equiparacao_mensal"), 0.0)), step=50.0)

        st.subheader("3. Seleção de Pedidos, Reflexos e Multas da Inicial")
        meses_contrato = max(1, (dism_pre.year - adm_pre.year) * 12 + (dism_pre.month - adm_pre.month) + 1)

        v1, v2, v3 = st.columns(3)
        ped_saldo = v1.checkbox("Saldo de Salário", value=bool(ext_pre.get("sugerir_saldo_salario", True)))
        dias_saldo_pre = v1.number_input("Dias trabalhados no último mês", min_value=0, max_value=31, value=max(0, min(31, int(dism_pre.day))))
        ped_reflexos_completos = v1.checkbox("Reflexos HE/Noturno em 13º, Férias + 1/3 e Aviso", value=True)

        ped_aviso = v2.checkbox("Aviso-Prévio Indenizado + Projeção", value=bool(ext_pre.get("sugerir_aviso_previo", True)))
        ped_13_ferias = v2.checkbox("13º Proporcional e Férias + 1/3", value=bool(ext_pre.get("sugerir_13_e_ferias", True)))
        ped_multa_477 = v2.checkbox("Multa Art. 477, § 8º CLT (Atraso TRCT)", value=bool(ext_pre.get("sugerir_multa_477", False)))
        ped_multa_467 = v2.checkbox("Multa Art. 467 CLT (50% Incontroversas)", value=bool(ext_pre.get("sugerir_multa_467", False)))

        default_multa_fgts = f_float(
            ext_pre.get("multa_fgts_percentual_sugerido"),
            40.0 if motivo_saida in ("sem_justa_causa", "rescisao_indireta") else (20.0 if motivo_saida == "acordo" else 0.0),
        )
        ped_fgts = v3.checkbox("FGTS (8%) + Multa Rescisória", value=bool(ext_pre.get("sugerir_fgts_e_multa", True)))
        multa_fgts_pct_val = v3.number_input(
            "Percentual da Multa FGTS a Cobrar (%)",
            min_value=0.0,
            max_value=100.0,
            value=max(0.0, min(100.0, default_multa_fgts)),
            step=5.0,
            help="Padrão: 40% (dispensa sem justa causa), 20% (acordo), ou digite qualquer % (ex.: 10% se já pagaram 30%).",
        )
        fgts_faltante_val = v3.number_input(
            "FGTS Não Depositado no Contrato (R$)",
            min_value=0.0,
            value=max(0.0, f_float(ext_pre.get("fgts_nao_depositado_contrato"), 0.0)),
            step=100.0,
        )
        meses_estab_val = v3.number_input(
            "Meses de Estabilidade Violada (B91/Gestante/CIPA)",
            min_value=0,
            max_value=24,
            value=max(0, min(24, f_int(ext_pre.get("meses_estabilidade_restantes"), 0))),
        )
        rest_auto = max(0.0, f_float(ext_pre.get("restituicao_atestados_ou_descontos"), 0.0))
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
                total_he50 = f_float(tc_pre.get("he50_hours"), 0.0)
                total_he100 = f_float(tc_pre.get("he100_hours"), 0.0)
                total_intra = f_float(tc_pre.get("intrajornada_hours"), 0.0)
                total_not = f_float(tc_pre.get("night_ficta_hours"), 0.0)
            else:
                total_he50 = float((Decimal(str(he_semanal_relatada)) * semanas_periodo).quantize(CENT, rounding=ROUND_HALF_UP))
                total_he100 = float((Decimal(str(domingos_mes * 8)) * Decimal(str(meses_prescricao))).quantize(CENT, rounding=ROUND_HALF_UP))
                dias_uteis_periodo = Decimal(str(meses_prescricao)) * Decimal("22")
                total_intra = float(((Decimal(str(min_intervalo_suprimido)) / Decimal("60")) * dias_uteis_periodo).quantize(CENT, rounding=ROUND_HALF_UP))
                total_not = float((Decimal(str(horas_noturnas_semana)) * semanas_periodo * Decimal("1.142857")).quantize(CENT, rounding=ROUND_HALF_UP))

            sm_vigente = get_historical_min_wage(dism_pre)
            insal_mensal = Decimal("0.00")
            adic_pct_lbl = "0%"
            pct_adic_dec = Decimal(str(adicional_pct_custom)) / Decimal("100")
            if adicional_cond == "Insalubridade (sobre Salário Mínimo)" and adicional_pct_custom > 0:
                insal_mensal = (sm_vigente * pct_adic_dec).quantize(CENT, rounding=ROUND_HALF_UP)
                adic_pct_lbl = f"{adicional_pct_custom:g}% SM (NR-15)"
            elif adicional_cond == "Periculosidade (sobre Salário-Base)" and adicional_pct_custom > 0:
                insal_mensal = (d_money(sal_registrado) * pct_adic_dec).quantize(CENT, rounding=ROUND_HALF_UP)
                adic_pct_lbl = f"{adicional_pct_custom:g}% Sal. (NR-16)"

            res_pre = run_labor_calculation({
                "process_number": f"INICIAL-{cliente_nome.upper().replace(' ', '_')[:18]}",
                "office_header": office_timbre,
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
                "periodos_ferias_vencidas": ferias_vencidas_val,
                "meses_sem_registro": meses_sem_ctps_val,
                "plus_acumulo_pct": plus_acumulo_val,
                "dif_equiparacao_mensal": dif_equip_val,
                "inc_fgts": ped_fgts,
                "fgts_multa_pct": multa_fgts_pct_val,
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

            # Substitui cálculo anterior do mesmo cliente (preservando todo o histórico de chat já conversado!)
            new_client_key = f"{cliente_nome.strip().lower()}|{empresa_nome.strip().lower()}"
            existing_chat_backup: List[Dict[str, str]] = []
            filtered_history: List[Dict[str, Any]] = []
            for old_c in st.session_state.calculations_history:
                old_key = f"{str(old_c.get('worker_name', '')).strip().lower()}|{str(old_c.get('employer_name', '')).strip().lower()}"
                if old_key == new_client_key:
                    old_id = str(old_c["audit_id"])
                    if st.session_state.case_chat_threads.get(old_id):
                        existing_chat_backup = list(st.session_state.case_chat_threads[old_id])
                else:
                    filtered_history.append(old_c)

            st.session_state.calculations_history = [res_pre] + filtered_history
            draft_msgs = list(st.session_state.case_chat_threads.get("caso_atual", []))
            final_case_msgs = existing_chat_backup or draft_msgs
            st.session_state.case_chat_threads[res_pre["audit_id"]] = final_case_msgs
            st.session_state.selected_chat_case_id = res_pre["audit_id"]
            st.session_state["chat_case_selector_widget"] = res_pre["audit_id"]
            if supabase_upsert_case(res_pre, final_case_msgs):
                st.session_state.cloud_connected = True
            st.success(f"Cálculo preliminar pronto para ajuizamento! Código: **{res_pre['audit_id']}**")

            k1, k2, k3 = st.columns(3)
            k1.metric("Total Apurado das Verbas", fmt_brl(res_pre["bruto_atualizado"]))
            k2.metric("(-) Já Pago no TRCT", fmt_brl(res_pre["abatimentos"]))
            k3.metric("Valor da Causa / Restituição a Pedir", fmt_brl(res_pre["liquido_devido"]))

            df_pre = pd.DataFrame(res_pre["verbas"])
            if not df_pre.empty:
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

            if ext_pre.get("petition_draft_topics"):
                with st.expander("Esboço dos Tópicos Jurídicos (Dos Fatos e Do Direito) para a Petição Inicial", expanded=False):
                    st.text_area("Fundamentação preliminar sugerida pela IA:", value=str(ext_pre["petition_draft_topics"]), height=260)

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

    # ==========================================================================
    # SEÇÃO 4: CHAT JURÍDICO INTERATIVO COM MEMÓRIA MULTI-CLIENTE (PEDRO, JOÃO...)
    # ==========================================================================
    st.divider()
    with st.container(border=True):
        ch_top_l, ch_top_r = st.columns([3, 1])
        with ch_top_l:
            st.subheader("4. Chat Jurídico Estratégico com a IA (Memória por Cliente)")
            st.caption(
                "Converse livremente com a IA em tela ampla para discutir teses, pedir trechos da Petição Inicial, "
                "quesitos de perícia ou alternar entre qualquer cliente já analisado na sessão (ex.: Pedro, João...)."
            )

        case_options: Dict[str, str] = {"caso_atual": "Caso Atual em Edição na Tela"}
        for c_item in st.session_state.calculations_history:
            cid = str(c_item["audit_id"])
            case_options[cid] = (
                f"{c_item.get('worker_name', 'Cliente')} × {c_item.get('employer_name', 'Empresa')} "
                f"({cid} • {fmt_brl(c_item['liquido_devido'])})"
            )

        if st.session_state.selected_chat_case_id not in case_options:
            st.session_state.selected_chat_case_id = "caso_atual"

        sel_col, clr_col = st.columns([4, 1])
        active_case_id = sel_col.selectbox(
            "Cliente / Caso selecionado para conversar:",
            options=list(case_options.keys()),
            index=list(case_options.keys()).index(st.session_state.selected_chat_case_id),
            format_func=lambda k: case_options[k],
            key="chat_case_selector_widget",
        )
        st.session_state.selected_chat_case_id = active_case_id

        if active_case_id not in st.session_state.case_chat_threads:
            st.session_state.case_chat_threads[active_case_id] = []

        if clr_col.button("Limpar Conversa", key="main_clear_chat_btn", use_container_width=True):
            st.session_state.case_chat_threads[active_case_id] = []
            if active_case_id != "caso_atual":
                calc_to_sync = next((c for c in st.session_state.calculations_history if c["audit_id"] == active_case_id), None)
                if calc_to_sync:
                    supabase_upsert_case(calc_to_sync, [])
            st.rerun()

        current_thread = st.session_state.case_chat_threads[active_case_id]

        chat_viewport = st.container(height=420, border=True)
        with chat_viewport:
            if not current_thread:
                st.info(
                    "Digite sua pergunta abaixo para iniciar a conversa com a IA. "
                    "Você pode perguntar sobre o cliente selecionado acima ou citar qualquer outro cliente já analisado (ex.: 'E no caso do Pedro, como fica a insalubridade?')."
                )
            for msg_item in current_thread:
                with st.chat_message(msg_item["role"]):
                    st.markdown(msg_item["content"])

        with st.form("main_wide_chat_form", clear_on_submit=True):
            main_user_msg = st.text_area(
                "Mensagem para o Advogado Trabalhista Sênior (IA):",
                placeholder="Ex.: Analise os riscos desse pedido, redija o tópico 'Do Direito' da inicial ou compare com o caso do Pedro...",
                height=95,
                label_visibility="collapsed",
            )
            send_col1, send_col2 = st.columns([4, 1.2])
            send_col1.caption("Dica: Você pode alternar de cliente no seletor acima a qualquer momento — o histórico de cada cliente fica salvo separadamente.")
            main_send_btn = send_col2.form_submit_button("Enviar para IA", type="primary", use_container_width=True)

        if main_send_btn and main_user_msg.strip():
            user_txt_clean = main_user_msg.strip()
            st.session_state.case_chat_threads[active_case_id].append({"role": "user", "content": user_txt_clean})

            if active_case_id == "caso_atual":
                ctx_calc = st.session_state.calculations_history[0] if st.session_state.calculations_history else None
                ctx_narrative = str(st.session_state.get("lawyer_narrative_box") or "")
                ctx_ext = st.session_state.extracted_ai_data or {}
            else:
                ctx_calc = next(
                    (c for c in st.session_state.calculations_history if c["audit_id"] == active_case_id),
                    None,
                )
                ctx_narrative = str((ctx_calc or {}).get("lawyer_narrative") or "")
                ctx_ext = (ctx_calc or {}).get("ai_opinion") or {}

            try:
                with st.spinner("Consultando IA trabalhista..."):
                    ai_reply_wide = chat_with_case_strategist(
                        user_message=user_txt_clean,
                        chat_history=st.session_state.case_chat_threads[active_case_id][:-1],
                        lawyer_narrative=ctx_narrative,
                        extracted_data=ctx_ext,
                        latest_calc=ctx_calc,
                        all_saved_cases=st.session_state.calculations_history,
                    )
                st.session_state.case_chat_threads[active_case_id].append({"role": "assistant", "content": ai_reply_wide})
                if active_case_id != "caso_atual" and ctx_calc is not None:
                    supabase_upsert_case(ctx_calc, st.session_state.case_chat_threads[active_case_id])
                st.rerun()
            except Exception as wide_chat_exc:
                st.error(f"Erro no chat: {wide_chat_exc}")

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
                if not df_rep.empty:
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
