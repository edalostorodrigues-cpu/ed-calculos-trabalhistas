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


def get_admin_email() -> str:
    """Retorna o e-mail administrador configurado em st.secrets ou ambiente."""
    try:
        if "ADMIN_EMAIL" in st.secrets:
            return str(st.secrets["ADMIN_EMAIL"]).strip().lower()
    except Exception:
        pass
    return os.environ.get("ADMIN_EMAIL", "").strip().lower()


def supabase_check_authorization(email: str) -> tuple[bool, str, Dict[str, Any]]:
    """Verifica se o e-mail possui licença ativa e válida na tabela ed_licencas_usuarios."""
    email_clean = (email or "").strip().lower()
    if not email_clean:
        return False, "E-mail não informado para verificação de licença.", {}

    admin_mail = get_admin_email()
    if admin_mail and email_clean == admin_mail:
        return True, "Acesso administrador autorizado.", {"status": "ativo", "plano": "admin"}

    sb_url, sb_key = get_supabase_credentials()
    if not sb_url or not sb_key:
        # Mantém compatibilidade caso os secrets ainda não tenham sido configurados
        return True, "Sem bloqueio de licença (Secrets Supabase não configurados).", {"status": "ativo"}

    try:
        email_q = urllib.parse.quote(email_clean, safe="")
        endpoint = f"{sb_url}/rest/v1/ed_licencas_usuarios?email=eq.{email_q}&select=*"
        req = urllib.request.Request(
            endpoint,
            headers=_build_supabase_headers(sb_key, {"Accept": "application/json"}),
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            rows = json.loads(resp.read().decode("utf-8"))

        if not rows:
            return (
                False,
                "Seu e-mail ainda não possui licença ativa. Solicite a liberação ao administrador do sistema.",
                {"status": "nao_autorizado"},
            )

        lic = rows[0] or {}
        status = str(lic.get("status") or "").strip().lower()
        validade_raw = lic.get("validade_ate")

        if status != "ativo":
            return (
                False,
                f"Sua licença encontra-se com status '{status or 'inativo'}'. Entre em contato para regularizar o acesso.",
                lic,
            )

        if validade_raw:
            try:
                val_date = datetime.strptime(str(validade_raw)[:10], "%Y-%m-%d").date()
                if date.today() > val_date:
                    return (
                        False,
                        f"Sua licença expirou em {val_date.strftime('%d/%m/%Y')}. Renove para continuar utilizando.",
                        lic,
                    )
            except ValueError:
                pass

        return True, "Licença ativa e verificada.", lic
    except urllib.error.HTTPError as http_err:
        # Se a tabela ainda não foi criada no banco, não trava o usuário indevidamente
        if http_err.code in (400, 404):
            return True, "Tabela de licenças ainda não criada; acesso liberado por padrão.", {}
        return False, f"Falha ao validar licença na nuvem (HTTP {http_err.code}).", {}
    except Exception as exc:
        return False, f"Erro ao consultar autorização: {exc}", {}


def _serialize_calc_for_cloud(calc: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "audit_id": calc.get("audit_id"),
        "process_number": calc.get("process_number"),
        "office_header": calc.get("office_header"),
        "worker_name": calc.get("worker_name"),
        "employer_name": calc.get("employer_name"),
        "role_title": calc.get("role_title"),
        "admission_date": str(calc.get("admission_date") or ""),
        "dismissal_date": str(calc.get("dismissal_date") or ""),
        "termination_type": calc.get("termination_type"),
        "monthly_divisor": calc.get("monthly_divisor"),
        "adicional_label": calc.get("adicional_label"),
        "adicional_pct_label": calc.get("adicional_pct_label"),
        "bruto_atualizado": str(d_money(calc.get("bruto_atualizado"))),
        "abatimentos": str(d_money(calc.get("abatimentos"))),
        "liquido_devido": str(d_money(calc.get("liquido_devido"))),
        "lawyer_narrative": calc.get("lawyer_narrative") or "",
        "ai_opinion": calc.get("ai_opinion") or {},
        "verbas": [
            {
                **{k: (str(v_item) if isinstance(v_item, Decimal) else v_item) for k, v_item in v.items()}
            }
            for v in (calc.get("verbas") or [])
        ],
    }


def _ensure_session_state() -> None:
    """Garante a presença de todas as chaves necessárias no st.session_state."""
    if "auth_user" not in st.session_state:
        st.session_state.auth_user = None
    if "auth_blocked_info" not in st.session_state:
        st.session_state.auth_blocked_info = None
    if "calculations_history" not in st.session_state:
        st.session_state.calculations_history = []
    if "case_chat_threads" not in st.session_state:
        st.session_state.case_chat_threads = {"caso_atual": []}
    if "selected_chat_case_id" not in st.session_state:
        st.session_state.selected_chat_case_id = "caso_atual"
    if "extracted_ai_data" not in st.session_state:
        st.session_state.extracted_ai_data = None
    if "extraction_confirmed" not in st.session_state:
        st.session_state.extraction_confirmed = False
    if "timecard_result" not in st.session_state:
        st.session_state.timecard_result = None
    if "cloud_connected" not in st.session_state:
        st.session_state.cloud_connected = False
    if "cloud_status_detail" not in st.session_state:
        st.session_state.cloud_status_detail = ""


_ensure_session_state()
            with tab_login:
                with st.form("login_auth_form"):
                    login_email = st.text_input("E-mail profissional *", placeholder="advogado@escritorio.com.br")
                    login_pass = st.text_input("Senha *", type="password", placeholder="••••••••")
                    login_submit = st.form_submit_button("Entrar no Sistema", type="primary", use_container_width=True)

                if login_submit:
                    if not login_email.strip() or not login_pass:
                        st.error("Informe seu e-mail e sua senha.")
                    else:
                        with st.spinner("Autenticando no Supabase..."):
                            ok_auth, user_prof, auth_msg = supabase_auth_sign_in(login_email, login_pass)
                        if ok_auth:
                            allowed, reason, lic_data = supabase_check_authorization(user_prof["email"])
                            if not allowed:
                                st.session_state.auth_blocked_info = {
                                    "email": user_prof["email"],
                                    "reason": reason,
                                    "profile": user_prof,
                                    **lic_data,
                                }
                                st.rerun()
                            st.session_state.auth_user = {**user_prof, "license": lic_data}
                            c_calcs, c_chats, c_ok, c_msg = supabase_load_all_cases(user_prof["email"])
                            st.session_state.cloud_connected = c_ok
                            st.session_state.cloud_status_detail = c_msg
                            st.session_state.calculations_history = c_calcs
                            st.session_state.case_chat_threads = {"caso_atual": [], **c_chats}
                            st.rerun()
                        else:
                            st.error(auth_msg)

            with tab_signup:
                with st.form("signup_auth_form"):
                    reg_name = st.text_input("Nome do Advogado(a) / Responsável *", placeholder="Dr. Eduardo Dalosto")
                    reg_office = st.text_input("Nome do Escritório / OAB (Timbre dos PDFs) *", placeholder="Dalosto Advocacia Trabalhista — OAB/RS 00.000")
                    reg_email = st.text_input("E-mail para acesso *", placeholder="contato@escritorio.com.br")
                    reg_pass = st.text_input("Crie uma senha (mín. 6 caracteres) *", type="password")
                    reg_pass_conf = st.text_input("Confirme a senha *", type="password")
                    signup_submit = st.form_submit_button("Criar Conta e Acessar", type="primary", use_container_width=True)

                if signup_submit:
                    if not reg_name.strip() or not reg_email.strip() or len(reg_pass) < 6:
                        st.error("Preencha nome, e-mail válido e senha com no mínimo 6 caracteres.")
                    elif reg_pass != reg_pass_conf:
                        st.error("As senhas digitadas não conferem.")
                    else:
                        with st.spinner("Criando conta do escritório..."):
                            ok_reg, new_prof, reg_msg = supabase_auth_sign_up(reg_email, reg_pass, reg_name, reg_office)
                        if ok_reg and new_prof:
                            allowed, reason, lic_data = supabase_check_authorization(new_prof["email"])
                            if not allowed:
                                st.session_state.auth_blocked_info = {
                                    "email": new_prof["email"],
                                    "reason": reason,
                                    "profile": new_prof,
                                    **lic_data,
                                }
                                st.rerun()
                            st.session_state.auth_user = {**new_prof, "license": lic_data}
                            c_calcs, c_chats, c_ok, c_msg = supabase_load_all_cases(new_prof["email"])
                            st.session_state.cloud_connected = c_ok
                            st.session_state.cloud_status_detail = c_msg
                            st.session_state.calculations_history = c_calcs
                            st.session_state.case_chat_threads = {"caso_atual": [], **c_chats}
                            st.rerun()
                        elif ok_reg:
                            st.success(reg_msg)
                        else:
                            st.error(reg_msg)
    st.stop()

current_user: Dict[str, Any] = st.session_state.auth_user
current_office_email: str = str(current_user.get("email") or "").strip().lower()

# Revalida licença periodicamente durante a sessão ativa
still_allowed, still_reason, still_lic = supabase_check_authorization(current_office_email)
if not still_allowed:
    st.session_state.auth_user = None
    st.session_state.auth_blocked_info = {
        "email": current_office_email,
        "reason": still_reason,
        **still_lic,
    }
    st.rerun()

with st.sidebar:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas")
    st.caption(f"👤 {current_user.get('full_name', 'Advogado')} ({current_office_email})")
    lic_meta = current_user.get("license") or {}
    if lic_meta.get("validade_ate"):
        st.caption(f"🔑 Licença válida até: {str(lic_meta['validade_ate'])[:10]}")
    st.divider()
    menu = st.radio(
        "Navegação",
        [
            "Avaliação Preliminar & Auditoria",
            "Meus Cálculos e Relatórios",
        ],
    )
    st.divider()
    st.markdown(f"**Tabelas ativas:** <span style='background:#16324A;color:#93C5FD;padding:2px 8px;border-radius:6px;font-family:monospace;font-size:12px;'>{TABLES_VERSION}</span>", unsafe_allow_html=True)
    if st.session_state.get("cloud_connected"):
        st.markdown("**Nuvem Supabase:** <span style='background:#14532D;color:#86EFAC;padding:2px 8px;border-radius:6px;font-size:12px;'>Conectada ☁️</span>", unsafe_allow_html=True)
    else:
        st.markdown("**Nuvem Supabase:** <span style='background:#7F1D1D;color:#FCA5A5;padding:2px 8px;border-radius:6px;font-size:12px;'>Desconectada</span>", unsafe_allow_html=True)
        if st.session_state.get("cloud_status_detail"):
            st.caption(f"Detalhe: {st.session_state['cloud_status_detail']}")
        if st.button("Reconectar Nuvem", key="sb_reconnect_cloud_btn", use_container_width=True):
            c_calcs, c_chats, c_ok, c_msg = supabase_load_all_cases(current_office_email)
            st.session_state.cloud_connected = c_ok
            st.session_state.cloud_status_detail = c_msg
            if c_ok:
                st.session_state.calculations_history = c_calcs
                st.session_state.case_chat_threads = {"caso_atual": [], **c_chats}
            st.rerun()

    st.divider()
    if st.button("Sair da Conta (Logout)", key="sb_logout_btn", use_container_width=True):
        st.session_state.auth_user = None
        st.session_state.calculations_history = []
        st.session_state.case_chat_threads = {"caso_atual": []}
        st.session_state.extracted_ai_data = None
        st.session_state.timecard_result = None
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
            value=str(current_user.get("office_header") or "Escritório de Advocacia Trabalhista — Parecer Pré-Processual"),
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
            up_ok, up_msg = supabase_upsert_case(res_pre, final_case_msgs, current_office_email)
            if up_ok:
                st.session_state.cloud_connected = True
                st.session_state.cloud_status_detail = "Conectada"
            else:
                st.session_state.cloud_status_detail = up_msg
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
                    supabase_upsert_case(calc_to_sync, [], current_office_email)
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
                    supabase_upsert_case(ctx_calc, st.session_state.case_chat_threads[active_case_id], current_office_email)
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
