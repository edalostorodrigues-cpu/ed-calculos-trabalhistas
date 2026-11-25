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
from typing import List, Optional, Dict, Any, Tuple

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

st.set_page_config(
    page_title="ED Cálculos Trabalhistas",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# 1. CONSTANTES, TABELAS HISTÓRICAS E UTILITÁRIOS DECIMAIS EXATOS
# ==============================================================================
TABLES_VERSION = "2026.11-estavel"
CENT = Decimal("0.01")


def d_money(val: Any) -> Decimal:
    """Converte qualquer entrada monetária para Decimal exato com 2 casas."""
    if isinstance(val, Decimal):
        return val.quantize(CENT, rounding=ROUND_HALF_UP)
    if val is None:
        return Decimal("0.00")
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
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def f_int(val: Any, default: int = 0) -> int:
    if val is None:
        return default
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return default


def fmt_brl(val: Any) -> str:
    dec = d_money(val)
    neg = dec < Decimal("0.00")
    dec = abs(dec)
    inteiro = int(dec)
    centavos = int((dec - Decimal(inteiro)) * 100)
    int_fmt = f"{inteiro:,}".replace(",", ".")
    return f"{'-' if neg else ''}R$ {int_fmt},{centavos:02d}"


def get_historical_min_wage(ref_date: date) -> Decimal:
    """Retorna o salário mínimo nacional de referência conforme o ano."""
    year = ref_date.year
    table = {
        2021: Decimal("1100.00"),
        2022: Decimal("1212.00"),
        2023: Decimal("1320.00"),
        2024: Decimal("1412.00"),
        2025: Decimal("1518.00"),
        2026: Decimal("1621.00"),
    }
    return table.get(year, Decimal("1621.00"))


def get_secret_str(key_name: str, default: str = "") -> str:
    try:
        if key_name in st.secrets:
            return str(st.secrets[key_name]).strip()
    except Exception:
        pass
    return os.environ.get(key_name, default).strip()


def get_admin_email() -> str:
    return get_secret_str("ADMIN_EMAIL", "e.dalostorodrigues@gmail.com").lower()


# ==============================================================================
# 2. INTEGRAÇÃO SUPABASE (BANCO NA NUVEM + AUTH + LICENÇAS)
# ==============================================================================
def _get_supabase_creds() -> Tuple[str, str]:
    url = get_secret_str("SUPABASE_URL", "").rstrip("/")
    key = get_secret_str("SUPABASE_KEY", "")
    return url, key


def _build_supabase_headers(sb_key: str, extra: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    headers: Dict[str, str] = {
        "apikey": sb_key,
        "Content-Type": "application/json",
    }
    if sb_key and not sb_key.startswith("sb_publishable_"):
        headers["Authorization"] = f"Bearer {sb_key}"
    if extra:
        headers.update(extra)
    return headers


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


def supabase_load_all_cases(office_email: str) -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, str]]], bool, str]:
    sb_url, sb_key = _get_supabase_creds()
    if not sb_url or not sb_key:
        return [], {}, False, "Credenciais SUPABASE_URL / SUPABASE_KEY não configuradas."
    try:
        endpoint = f"{sb_url}/rest/v1/ed_calculos_clientes?select=*&order=updated_at.desc"
        req = urllib.request.Request(endpoint, headers=_build_supabase_headers(sb_key), method="GET")
        with urllib.request.urlopen(req, timeout=10) as resp:
            rows = json.loads(resp.read().decode("utf-8"))
        calcs: List[Dict[str, Any]] = []
        chats: Dict[str, List[Dict[str, str]]] = {}
        target_email = (office_email or "").strip().lower()
        for r in rows:
            row_office = str(r.get("office_email") or "").strip().lower()
            if row_office and target_email and row_office != target_email and target_email != get_admin_email():
                continue
            payload = r.get("payload_json") or {}
            if isinstance(payload, str):
                payload = json.loads(payload)
            if payload:
                c_obj = _deserialize_calc_from_cloud(payload)
                calcs.append(c_obj)
                aid = str(c_obj.get("audit_id") or r.get("audit_id") or "")
                ch = r.get("chat_history_json") or []
                if isinstance(ch, str):
                    ch = json.loads(ch)
                if aid:
                    chats[aid] = ch
        return calcs, chats, True, "Conectada"
    except Exception as exc:
        return [], {}, False, str(exc)


def supabase_upsert_case(calc: Dict[str, Any], chat_msgs: List[Dict[str, str]], office_email: str) -> Tuple[bool, str]:
    sb_url, sb_key = _get_supabase_creds()
    if not sb_url or not sb_key:
        return False, "Credenciais ausentes."
    try:
        w_name = str(calc.get("worker_name") or "Cliente").strip()
        e_name = str(calc.get("employer_name") or "Empresa").strip()
        client_key = f"{w_name.lower()}|{e_name.lower()}"
        body = {
            "client_key": client_key,
            "audit_id": str(calc.get("audit_id") or ""),
            "worker_name": w_name,
            "employer_name": e_name,
            "office_email": (office_email or "escritorio@edcalculos.com.br").strip().lower(),
            "liquido_devido": float(d_money(calc.get("liquido_devido"))),
            "payload_json": _serialize_calc_for_cloud(calc),
            "chat_history_json": chat_msgs or [],
            "updated_at": datetime.utcnow().isoformat() + "Z",
        }
        endpoint = f"{sb_url}/rest/v1/ed_calculos_clientes?on_conflict=client_key"
        headers = _build_supabase_headers(sb_key, {"Prefer": "resolution=merge-duplicates,return=minimal"})
        req = urllib.request.Request(endpoint, data=json.dumps(body).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            _ = resp.read()
        return True, "Salvo na nuvem"
    except Exception as exc:
        return False, str(exc)


def supabase_auth_sign_in(email: str, password: str) -> Tuple[bool, Optional[Dict[str, Any]], str]:
    sb_url, sb_key = _get_supabase_creds()
    if not sb_url or not sb_key:
        return False, None, "Supabase não configurado."
    try:
        endpoint = f"{sb_url}/auth/v1/token?grant_type=password"
        req = urllib.request.Request(
            endpoint,
            data=json.dumps({"email": email.strip(), "password": password}).encode("utf-8"),
            headers=_build_supabase_headers(sb_key),
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        user = data.get("user") or {}
        meta = user.get("user_metadata") or {}
        prof = {
            "email": str(user.get("email") or email).strip().lower(),
            "full_name": meta.get("full_name") or email.split("@")[0],
            "office_header": meta.get("office_header") or "Escritório de Advocacia Trabalhista",
        }
        return True, prof, "Autenticado"
    except Exception as exc:
        return False, None, f"Falha no login: {exc}"


def supabase_auth_sign_up(email: str, password: str, full_name: str, office_header: str) -> Tuple[bool, Optional[Dict[str, Any]], str]:
    sb_url, sb_key = _get_supabase_creds()
    if not sb_url or not sb_key:
        return False, None, "Supabase não configurado."
    try:
        endpoint = f"{sb_url}/auth/v1/signup"
        payload = {
            "email": email.strip(),
            "password": password,
            "data": {"full_name": full_name.strip(), "office_header": office_header.strip()},
        }
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers=_build_supabase_headers(sb_key),
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        user = data.get("user") or {}
        if data.get("access_token"):
            prof = {
                "email": str(user.get("email") or email).strip().lower(),
                "full_name": full_name.strip(),
                "office_header": office_header.strip(),
            }
            return True, prof, "Conta criada com sucesso!"
        return True, None, "Cadastro realizado! Verifique seu e-mail para confirmar ou faça login."
    except Exception as exc:
        return False, None, f"Erro no cadastro: {exc}"


def supabase_check_authorization(email: str) -> Tuple[bool, str, Dict[str, Any]]:
    clean_email = (email or "").strip().lower()
    if clean_email and clean_email == get_admin_email():
        return True, "Admin autorizado", {"status": "ativo", "plano": "admin", "validade_ate": "2099-12-31"}
    sb_url, sb_key = _get_supabase_creds()
    if not sb_url or not sb_key:
        return True, "Modo local", {"status": "ativo"}
    try:
        q_email = urllib.parse.quote(clean_email)
        endpoint = f"{sb_url}/rest/v1/ed_licencas_usuarios?email=eq.{q_email}&select=*"
        req = urllib.request.Request(endpoint, headers=_build_supabase_headers(sb_key), method="GET")
        with urllib.request.urlopen(req, timeout=8) as resp:
            rows = json.loads(resp.read().decode("utf-8"))
        if not rows:
            return False, "E-mail ainda não autorizado. Solicite a liberação da sua licença ao administrador.", {}
        lic = rows[0]
        if str(lic.get("status") or "").lower() != "ativo":
            return False, f"Status da licença: {lic.get('status', 'pendente')}.", lic
        val_str = str(lic.get("validade_ate") or "")[:10]
        if val_str:
            val_dt = datetime.strptime(val_str, "%Y-%m-%d").date()
            if val_dt < date.today():
                return False, f"Licença expirada em {val_dt.strftime('%d/%m/%Y')}.", lic
        return True, "Autorizado", lic
    except Exception:
        return True, "Verificação ignorada", {"status": "ativo"}


# ==============================================================================
# 3. APURADOR DETERMINÍSTICO DE CARTÃO DE PONTO
# ==============================================================================
def _hhmm_to_min(val: Any) -> Optional[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if ":" not in s:
        return None
    parts = s.split(":")
    try:
        h, m = int(parts[0]), int(parts[1])
        if 0 <= h <= 29 and 0 <= m <= 59:
            return h * 60 + m
    except ValueError:
        return None
    return None


def _calc_night_min(start_m: int, end_m: int) -> int:
    if end_m <= start_m:
        return 0
    total = 0
    for n_start, n_end in [(0, 300), (1320, 1740)]:
        ov_s = max(start_m, n_start)
        ov_e = min(end_m, n_end)
        if ov_e > ov_s:
            total += ov_e - ov_s
    return total


def process_timecard_dataframe(df: pd.DataFrame, daily_std_min: int = 480) -> Dict[str, Any]:
    rows_out = []
    tot_he50 = 0
    tot_he100 = 0
    tot_intra = 0
    tot_night_ficta = 0

    for _, r in df.iterrows():
        dt_str = str(r.get("Data", ""))
        dia_str = str(r.get("Dia", r.get("DiaSemana", ""))).upper()
        fer_str = str(r.get("Feriado", "N")).upper()
        is_hol = fer_str == "S" or "DOM" in dia_str

        e1 = _hhmm_to_min(r.get("Entrada1"))
        s1 = _hhmm_to_min(r.get("Saida1"))
        e2 = _hhmm_to_min(r.get("Entrada2"))
        s2 = _hhmm_to_min(r.get("Saida2"))

        if e1 is None or s1 is None:
            rows_out.append({
                "Data": dt_str,
                "Dia": dia_str,
                "Trabalhado (min)": 0,
                "HE 50% (min)": 0,
                "HE 100% (min)": 0,
                "Intrajornada Sup. (min)": 0,
                "Noturna Ficta (min)": 0,
                "Status": "Batida incompleta",
            })
            continue

        p1_end = s1 + 1440 if s1 < e1 else s1
        w1 = max(0, p1_end - e1)
        n1 = _calc_night_min(e1, p1_end)

        w2, n2, brk = 0, 0, 0
        if e2 is not None and s2 is not None:
            p2_s = e2 + 1440 if e2 < p1_end else e2
            p2_e = s2 + 1440 if s2 < p2_s else s2
            brk = max(0, p2_s - p1_end)
            w2 = max(0, p2_e - p2_s)
            n2 = _calc_night_min(p2_s, p2_e)

        worked = w1 + w2
        night_ficta = int(round((n1 + n2) * (60.0 / 52.5)))

        intra = 0
        if worked > 360 and brk < 60:
            intra = 60 - brk
        elif 240 < worked <= 360 and brk < 15:
            intra = 15 - brk

        he50, he100 = 0, 0
        if is_hol:
            he100 = worked
        else:
            diff = worked - daily_std_min
            if diff > 10:
                he50 = diff

        tot_he50 += he50
        tot_he100 += he100
        tot_intra += intra
        tot_night_ficta += night_ficta

        rows_out.append({
            "Data": dt_str,
            "Dia": dia_str,
            "Trabalhado (min)": worked,
            "HE 50% (min)": he50,
            "HE 100% (min)": he100,
            "Intrajornada Sup. (min)": intra,
            "Noturna Ficta (min)": night_ficta,
            "Status": "Intrajornada suprimida" if intra > 0 else "Apurado",
        })

    return {
        "table": pd.DataFrame(rows_out),
        "he50_hours": round(tot_he50 / 60.0, 2),
        "he100_hours": round(tot_he100 / 60.0, 2),
        "intrajornada_hours": round(tot_intra / 60.0, 2),
        "night_ficta_hours": round(tot_night_ficta / 60.0, 2),
    }


# ==============================================================================
# 4. MOTOR DE CÁLCULO TRABALHISTA DECIMAL EXATO
# ==============================================================================
def run_labor_calculation(params: Dict[str, Any]) -> Dict[str, Any]:
    adm: date = params["admission_date"]
    dism: date = params["dismissal_date"]
    term_type: str = params.get("termination_type", "sem_justa_causa")
    divisor = Decimal(str(max(1, f_int(params.get("monthly_divisor"), 220))))

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
    meses_contrato = max(1, (dism.year - adm.year) * 12 + (dism.month - adm.month) + 1)
    meses_presc = min(60, meses_contrato)

    periodo_str = f"{adm.strftime('%d/%m/%Y')} a {dism.strftime('%d/%m/%Y')}"
    verbas: List[Dict[str, Any]] = []

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
        add_verba("VER-05", f"{params.get('adicional_label', 'Adicional')} Retroativo", insal_mensal, "Adicional Mensal × Meses + Reflexos", f"{meses_presc} meses", str(params.get("adicional_pct_label", "")), val_adic_ret, refl_adic)

    # 6. Plus Acúmulo de Função / Equiparação Salarial Retroativa
    if plus_acumulo_val > Decimal("0.00"):
        val_ac = (plus_acumulo_val * Decimal(str(meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_ac = (val_ac * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05B", "Plus Salarial por Acúmulo/Desvio de Função", plus_acumulo_val, "Plus Mensal × Meses + Reflexos", f"{meses_presc} meses", f"{params.get('plus_acumulo_pct', 0)}%", val_ac, refl_ac)

    if dif_equip > Decimal("0.00"):
        val_eq = (dif_equip * Decimal(str(meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_eq = (val_eq * Decimal("0.1944")).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-05C", "Diferenças por Equiparação Salarial (Art. 461 CLT)", dif_equip, "Diferença Mensal × Meses + Reflexos", f"{meses_presc} meses", "100%", val_eq, refl_eq)

    # 7. Horas Extras 50%, 100%, Intrajornada e Adicional Noturno
    fator_refl_jornada = Decimal("0.3935") if params.get("inc_reflexos_completos") else Decimal("0.1667")

    if params.get("inc_he50") and f_float(params.get("he50_hours"), 0.0) > 0:
        h50 = Decimal(str(f_float(params.get("he50_hours"), 0.0)))
        val_he50 = (valor_hora * Decimal("1.50") * h50).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he50 = (val_he50 * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06A", "Horas Extras 50% + DSR e Reflexos", remuneracao_base, "(Valor Hora × 1,50) × Horas + DSR/Reflexos", f"{h50}h", "+50%", val_he50, refl_he50)

    if params.get("inc_he100") and f_float(params.get("he100_hours"), 0.0) > 0:
        h100 = Decimal(str(f_float(params.get("he100_hours"), 0.0)))
        val_he100 = (valor_hora * Decimal("2.00") * h100).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_he100 = (val_he100 * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06B", "Horas Extras 100% (Domingos/Feriados) + Reflexos", remuneracao_base, "(Valor Hora × 2,00) × Horas + DSR/Reflexos", f"{h100}h", "+100%", val_he100, refl_he100)

    if params.get("inc_intrajornada") and f_float(params.get("intrajornada_hours"), 0.0) > 0:
        hintra = Decimal(str(f_float(params.get("intrajornada_hours"), 0.0)))
        val_intra = (valor_hora * Decimal("1.50") * hintra).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06C", "Intervalo Intrajornada Suprimido (Art. 71, §4º CLT)", remuneracao_base, "(Valor Hora × 1,50) × Horas Suprimidas", f"{hintra}h", "+50%", val_intra, inc_fgts=False)

    if params.get("inc_noturno") and f_float(params.get("night_hours"), 0.0) > 0:
        hnot = Decimal(str(f_float(params.get("night_hours"), 0.0)))
        val_not = (valor_hora * Decimal("0.20") * hnot).quantize(CENT, rounding=ROUND_HALF_UP)
        refl_not = (val_not * fator_refl_jornada).quantize(CENT, rounding=ROUND_HALF_UP)
        add_verba("VER-06D", "Adicional Noturno (20% + Hora Ficta) + Reflexos", remuneracao_base, "(Valor Hora × 20%) × Horas + Reflexos", f"{hnot}h", "+20%", val_not, refl_not)

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

    bruto_total = sum((v["Total Apurado"] for v in verbas), Decimal("0.00")).quantize(CENT, rounding=ROUND_HALF_UP)
    abatimentos = d_money(params.get("offsets"))
    liquido_devido = max(Decimal("0.00"), bruto_total - abatimentos).quantize(CENT, rounding=ROUND_HALF_UP)

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
        "base_remuneracao": remuneracao_base,
        "verbas": verbas,
        "bruto_atualizado": bruto_total,
        "abatimentos": abatimentos,
        "liquido_devido": liquido_devido,
        "ai_opinion": params.get("ai_opinion") or {},
        "lawyer_narrative": params.get("lawyer_narrative") or "",
        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }


# ==============================================================================
# 5. EXPORTAÇÃO EXCEL (.XLSX) E PDF OFICIAL (.PDF)
# ==============================================================================
def build_complete_excel_bytes(calc: Dict[str, Any]) -> bytes:
    buf = io.BytesIO()
    resumo_rows = [
        {"Campo": "Código Auditoria", "Valor": calc.get("audit_id", "")},
        {"Campo": "Processo / Referência", "Valor": calc.get("process_number", "")},
        {"Campo": "Reclamante / Exequente", "Valor": calc.get("worker_name", "")},
        {"Campo": "Reclamada / Executada", "Valor": calc.get("employer_name", "")},
        {"Campo": "Período Contratual", "Valor": f"{calc.get('admission_date')} a {calc.get('dismissal_date')}"},
        {"Campo": "Remuneração Base Apurada", "Valor": fmt_brl(calc.get("base_remuneracao"))},
        {"Campo": "Total Bruto Apurado", "Valor": fmt_brl(calc.get("bruto_atualizado"))},
        {"Campo": "(-) Abatimentos / Pagos", "Valor": fmt_brl(calc.get("abatimentos"))},
        {"Campo": "Total Líquido Devido", "Valor": fmt_brl(calc.get("liquido_devido"))},
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

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(resumo_rows).to_excel(writer, sheet_name="Resumo_Executivo", index=False)
        pd.DataFrame(verbas_rows).to_excel(writer, sheet_name="Memoria_Verbas", index=False)
    return buf.getvalue()


def _pdf_safe(txt: Any) -> str:
    return str(txt or "").encode("latin-1", errors="replace").decode("latin-1")


def build_printable_dossier_pdf_bytes(calc: Dict[str, Any]) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, _pdf_safe(calc.get("office_header") or "ED Cálculos Trabalhistas"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, _pdf_safe(f"Código: {calc.get('audit_id')} | Emitido em: {calc.get('created_at')} | Tabelas: {TABLES_VERSION}"), ln=True)
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, _pdf_safe(f"Reclamante: {calc.get('worker_name')}  |  Reclamada: {calc.get('employer_name')}"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, _pdf_safe(f"Processo/Ref: {calc.get('process_number')} | Período: {calc.get('admission_date')} a {calc.get('dismissal_date')}"), ln=True)
    pdf.ln(4)

    pdf.set_fill_color(15, 37, 55)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(18, 7, "Cod.", 1, 0, "L", True)
    pdf.cell(82, 7, "Verba / Rubrica", 1, 0, "L", True)
    pdf.cell(30, 7, "Original", 1, 0, "R", True)
    pdf.cell(30, 7, "Reflexos", 1, 0, "R", True)
    pdf.cell(30, 7, "Total", 1, 1, "R", True)

    pdf.set_text_color(15, 37, 55)
    pdf.set_font("Helvetica", "", 8)
    for v in calc.get("verbas", []):
        pdf.cell(18, 6, _pdf_safe(v["Código"]), 1)
        pdf.cell(82, 6, _pdf_safe(str(v["Verba"])[:46]), 1)
        pdf.cell(30, 6, _pdf_safe(fmt_brl(v["Valor Original"])), 1, 0, "R")
        pdf.cell(30, 6, _pdf_safe(fmt_brl(v["Reflexos"])), 1, 0, "R")
        pdf.cell(30, 6, _pdf_safe(fmt_brl(v["Total Apurado"])), 1, 1, "R")

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, _pdf_safe(f"Total Bruto Apurado: {fmt_brl(calc.get('bruto_atualizado'))}"), ln=True)
    pdf.cell(0, 6, _pdf_safe(f"(-) Abatimentos / Deduções: {fmt_brl(calc.get('abatimentos'))}"), ln=True)
    pdf.cell(0, 7, _pdf_safe(f"TOTAL LÍQUIDO DEVIDO: {fmt_brl(calc.get('liquido_devido'))}"), ln=True)

    raw_out = pdf.output(dest="S")
    if isinstance(raw_out, str):
        return raw_out.encode("latin-1")
    return bytes(raw_out)


# ==============================================================================
# 6. INTEGRAÇÃO GEMINI (PARECER INICIAL, LIQUIDAÇÃO DE SENTENÇA E CHAT)
# ==============================================================================
def _get_gemini_client():
    api_key = get_secret_str("GEMINI_API_KEY", "")
    if not GENAI_AVAILABLE:
        raise RuntimeError("Biblioteca google-genai não instalada.")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY não configurada nos Secrets.")
    return genai.Client(api_key=api_key)


def analyze_case_with_gemini(narrative: str, files: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = _get_gemini_client()
    prompt = (
        "Você é um Advogado Trabalhista Sênior e Perito Calculista (CLT, Súmulas TST, NRs 15 e 16). "
        "Analise o relato e todos os documentos/fotos anexados. Retorne EXCLUSIVAMENTE um JSON válido com as chaves: "
        "worker_name, employer_name, role_title, admission_date (YYYY-MM-DD ou null), dismissal_date (YYYY-MM-DD ou null), "
        "termination_type (sem_justa_causa|rescisao_indireta|pedido_demissao|acordo|justa_causa), monthly_divisor (int), "
        "base_salary (float), commissions (float), adicional_sugerido (Nenhum|Insalubridade|Periculosidade), "
        "adicional_percentual_sugerido (float: 0, 10, 20, 30 ou 40), nr_technical_assessment (str), "
        "stability_alert (str ou null), medical_leave_summary (str ou null), viability_conclusion (str), "
        "irregularities_found (lista de objetos {Irregularidade, Fundamento_Legal, Impacto}), "
        "recommended_claims_checklist (lista de strings), ambiguities (lista de strings), "
        "weekly_he50_hours (float), monthly_sundays_worked (int), suppressed_break_minutes (int), weekly_night_hours (float), "
        "meses_sem_registro_ctps (int), periodos_ferias_vencidas (int), plus_salarial_acumulo_pct (float), "
        "diferenca_equiparacao_mensal (float), sugerir_saldo_salario (bool), sugerir_aviso_previo (bool), "
        "sugerir_13_e_ferias (bool), sugerir_fgts_e_multa (bool), multa_fgts_percentual_sugerido (float), "
        "fgts_nao_depositado_contrato (float), meses_estabilidade_restantes (int), sugerir_multa_477 (bool), "
        "sugerir_multa_467 (bool), restituicao_atestados_ou_descontos (float), trct_paid_items (lista de {verba, valor}), "
        "petition_draft_topics (str com esboço Dos Fatos e Do Direito).\n\n"
        f"RELATO DO ADVOGADO:\n{narrative}"
    )
    contents: List[Any] = [prompt]
    for f in files:
        contents.append(types.Part.from_bytes(data=f["bytes"], mime_type=f["mime"]))

    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=contents,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1),
    )
    return json.loads(resp.text)


def analyze_sentence_with_gemini(notes: str, files: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = _get_gemini_client()
    prompt = (
        "Você é um Perito Calculista Trabalhista especialista em Liquidação de Sentença (Art. 879 CLT). "
        "Analise a cópia do processo/sentença transitada em julgado e documentos anexados. "
        "Separe rigorosamente o que foi DEFERIDO no dispositivo (descartando os pedidos improcedentes). "
        "Retorne EXCLUSIVAMENTE um JSON válido com as chaves: "
        "process_number (str), worker_name (str), employer_name (str), role_title (str), "
        "admission_date (YYYY-MM-DD ou null), dismissal_date (YYYY-MM-DD ou null), "
        "termination_type (sem_justa_causa|rescisao_indireta|pedido_demissao|acordo|justa_causa), monthly_divisor (int), "
        "base_salary (float), commissions (float), adicional_sugerido (Nenhum|Insalubridade|Periculosidade), "
        "adicional_percentual_sugerido (float), trct_paid_total (float), "
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
        model="gemini-2.5-flash",
        contents=contents,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1),
    )
    return json.loads(resp.text)


def chat_with_case_strategist(
    user_message: str,
    chat_history: List[Dict[str, str]],
    lawyer_narrative: str,
    extracted_data: Dict[str, Any],
    latest_calc: Optional[Dict[str, Any]],
    all_saved_cases: List[Dict[str, Any]],
) -> str:
    client = _get_gemini_client()
    cases_summary = [
        {
            "audit_id": c.get("audit_id"),
            "cliente": c.get("worker_name"),
            "empresa": c.get("employer_name"),
            "liquido": fmt_brl(c.get("liquido_devido")),
        }
        for c in all_saved_cases[:15]
    ]
    active_calc_summary = None
    if latest_calc:
        active_calc_summary = {
            "audit_id": latest_calc.get("audit_id"),
            "cliente": latest_calc.get("worker_name"),
            "empresa": latest_calc.get("employer_name"),
            "bruto": fmt_brl(latest_calc.get("bruto_atualizado")),
            "abatimentos": fmt_brl(latest_calc.get("abatimentos")),
            "liquido": fmt_brl(latest_calc.get("liquido_devido")),
            "verbas": [
                {"verba": v["Verba"], "total": fmt_brl(v["Total Apurado"])}
                for v in latest_calc.get("verbas", [])
            ],
        }

    system_ctx = (
        "Você é um Advogado Trabalhista Sênior e Perito Calculista auxiliando outro advogado em seu escritório. "
        "Responda em português jurídico claro, objetivo e estratégico, citando artigos da CLT, Súmulas/OJs do TST e NRs quando pertinente.\n"
        f"CLIENTES SALVOS NO ESCRITÓRIO: {json.dumps(cases_summary, ensure_ascii=False)}\n"
        f"CASO ATIVO SELECIONADO: {json.dumps(active_calc_summary, ensure_ascii=False)}\n"
        f"RELATO DO CASO ATIVO: {lawyer_narrative}\n"
        f"PARECER EXTRAÍDO: {json.dumps(extracted_data, ensure_ascii=False)[:3500]}"
    )

    conv_lines = [system_ctx, "\n--- HISTÓRICO DA CONVERSA ---"]
    for m in chat_history[-12:]:
        role_lbl = "ADVOGADO" if m["role"] == "user" else "CONSULTOR IA"
        conv_lines.append(f"{role_lbl}: {m['content']}")
    conv_lines.append(f"ADVOGADO: {user_message}")

    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents="\n".join(conv_lines),
        config=types.GenerateContentConfig(temperature=0.3),
    )
    return resp.text or "Sem resposta gerada."


# ==============================================================================
# 7. INICIALIZAÇÃO DO ESTADO DA SESSÃO
# ==============================================================================
defaults_state: Dict[str, Any] = {
    "auth_user": None,
    "auth_blocked_info": None,
    "cloud_connected": False,
    "cloud_status_detail": "",
    "calculations_history": [],
    "case_chat_threads": {"caso_atual": []},
    "selected_chat_case_id": "caso_atual",
    "extracted_ai_data": None,
    "extraction_confirmed": False,
    "timecard_result": None,
    "sentence_ai_data": None,
    "sentence_confirmed": False,
    "sentence_timecard_result": None,
}
for k_state, v_state in defaults_state.items():
    if k_state not in st.session_state:
        st.session_state[k_state] = v_state


# ==============================================================================
# 8. PORTAL DE LOGIN E CADASTRO (DESABILITADO TEMPORARIAMENTE PARA DESENVOLVIMENTO)
# ==============================================================================
REQUIRE_LOGIN = False  # Altere para True quando quiser reativar o bloqueio de login/licença

if st.session_state.auth_user is None and not REQUIRE_LOGIN:
    st.session_state.auth_user = {
        "email": get_admin_email() or "e.dalostorodrigues@gmail.com",
        "full_name": "Dr. Eduardo Dalosto (Modo Desenvolvimento)",
        "office_header": "Dalosto Advocacia Trabalhista — Parecer Pré-Processual & Liquidação",
        "license": {"status": "ativo", "plano": "admin_dev", "validade_ate": "2099-12-31"},
    }
    c_calcs, c_chats, c_ok, c_msg = supabase_load_all_cases(st.session_state.auth_user["email"])
    st.session_state.cloud_connected = c_ok
    st.session_state.cloud_status_detail = c_msg
    st.session_state.calculations_history = c_calcs
    st.session_state.case_chat_threads = {"caso_atual": [], **c_chats}

if st.session_state.auth_user is None and REQUIRE_LOGIN:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas — Acesso ao Sistema")
    if st.session_state.auth_blocked_info:
        blk = st.session_state.auth_blocked_info
        st.error(f"**Acesso Restrito ({blk.get('email', '')}):** {blk.get('reason', 'Licença inativa.')}")
        if st.button("Tentar Novamente / Voltar ao Login"):
            st.session_state.auth_blocked_info = None
            st.rerun()
        st.stop()

    _, center_col, _ = st.columns([1, 1.4, 1])
    with center_col:
        with st.container(border=True):
            tab_login, tab_signup = st.tabs(["Entrar (Login)", "Cadastrar Escritório"])

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
                        if ok_auth and user_prof:
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

if REQUIRE_LOGIN:
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
            "Liquidação de Sentença & Impugnação",
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

    if REQUIRE_LOGIN:
        st.divider()
        if st.button("Sair da Conta (Logout)", key="sb_logout_btn", use_container_width=True):
            st.session_state.auth_user = None
            st.session_state.calculations_history = []
            st.session_state.case_chat_threads = {"caso_atual": []}
            st.session_state.extracted_ai_data = None
            st.session_state.timecard_result = None
            st.rerun()


# ==============================================================================
# 9. TELAS DO SISTEMA
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
        adicional_pct_custom = r4.number_input("Percentual Adicional (%)", min_value=0.0, max_value=100.0, value=max(0.0, min(100.0, pct_adic_default)), step=5.0)
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
    # SEÇÃO 4: CHAT JURÍDICO INTERATIVO COM MEMÓRIA MULTI-CLIENTE
    # ==========================================================================
    st.divider()
    with st.container(border=True):
        ch_top_l, _ = st.columns([3, 1])
        with ch_top_l:
            st.subheader("4. Chat Jurídico Estratégico com a IA (Memória por Cliente)")
            st.caption(
                "Converse livremente com a IA em tela ampla para discutir teses, pedir trechos da Petição Inicial, "
                "quesitos de perícia ou alternar entre qualquer cliente já analisado na sessão."
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
                    "Você pode perguntar sobre o cliente selecionado acima ou citar qualquer outro cliente já analisado."
                )
            for msg_item in current_thread:
                with st.chat_message(msg_item["role"]):
                    st.markdown(msg_item["content"])

        with st.form("main_wide_chat_form", clear_on_submit=True):
            main_user_msg = st.text_area(
                "Mensagem para o Advogado Trabalhista Sênior (IA):",
                placeholder="Ex.: Analise os riscos desse pedido, redija o tópico 'Do Direito' da inicial...",
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

elif menu == "Liquidação de Sentença & Impugnação":
    st.title("Liquidação de Sentença Transitada em Julgado (Art. 879 CLT)")
    st.markdown(
        "Anexe a **cópia integral do processo / sentença sem recurso** (e, se houver, cartões de ponto, contracheques ou planilha da parte contrária). "
        "A IA identifica **exclusivamente o que foi deferido no dispositivo**, descarta os pedidos improcedentes e preenche o motor decimal exato para gerar a planilha de liquidação."
    )

    with st.container(border=True):
        st.subheader("1. Upload da Cópia do Processo / Sentença + Documentos dos Autos")
        sent_notes = st.text_area(
            "Observações do advogado sobre a sentença ou pontos para conferir no cálculo adverso (opcional):",
            placeholder="Ex.: Processo transitado em julgado sem recurso. Conferir horas extras deferidas acima da 8ª diária, reflexos em DSR/13º/férias/FGTS + 40% e dedução do TRCT.",
            height=105,
            key="sentence_notes_box",
        )
        sent_files = st.file_uploader(
            "Anexe a Sentença / Cópia do Processo (PDF) e anexos (Cartão de Ponto CSV/XLSX, Contracheques, Planilha Adversa)",
            type=["pdf", "png", "jpg", "jpeg", "csv", "xlsx", "txt"],
            accept_multiple_files=True,
            key="sentence_dossier_uploader",
        )
        if sent_files:
            st.caption(f"**{len(sent_files)} arquivo(s) anexado(s):** " + ", ".join(f"`{f.name}`" for f in sent_files))

        if st.button("Ler Sentença / Processo e Extrair Comando Condenatório (IA)", type="primary"):
            if not sent_notes.strip() and not sent_files:
                st.warning("Anexe a cópia do processo/sentença em PDF ou descreva o dispositivo da sentença.")
            else:
                g_files: List[Dict[str, Any]] = []
                tc_dfs: List[pd.DataFrame] = []
                for f in (sent_files or []):
                    fn_low = f.name.lower()
                    if fn_low.endswith(".csv"):
                        raw_t = f.getvalue().decode("utf-8", errors="ignore")
                        first_l = raw_t.splitlines()[0] if raw_t.splitlines() else ""
                        sep = ";" if ";" in first_l else ","
                        df_t = pd.read_csv(io.StringIO(raw_t), sep=sep)
                        if {"Entrada1", "Saida1"}.issubset(set(df_t.columns)):
                            tc_dfs.append(df_t)
                        else:
                            g_files.append({"name": f.name, "bytes": raw_t.encode("utf-8"), "mime": "text/plain"})
                    elif fn_low.endswith(".xlsx"):
                        df_t = pd.read_excel(f)
                        if {"Entrada1", "Saida1"}.issubset(set(df_t.columns)):
                            tc_dfs.append(df_t)
                        else:
                            g_files.append({"name": f.name, "bytes": df_t.to_csv(index=False).encode("utf-8"), "mime": "text/plain"})
                    else:
                        m_map = {
                            ".pdf": "application/pdf",
                            ".png": "image/png",
                            ".jpg": "image/jpeg",
                            ".jpeg": "image/jpeg",
                            ".txt": "text/plain",
                        }
                        ext = "." + fn_low.rsplit(".", 1)[-1] if "." in fn_low else ".txt"
                        g_files.append({"name": f.name, "bytes": f.getvalue(), "mime": m_map.get(ext, "application/pdf")})

                if tc_dfs:
                    st.session_state.sentence_timecard_result = process_timecard_dataframe(pd.concat(tc_dfs, ignore_index=True))
                    st.success("Cartão de ponto dos autos apurado matematicamente!")

                try:
                    with st.spinner("Lendo os autos e extraindo dispositivo condenatório da sentença..."):
                        st.session_state.sentence_ai_data = analyze_sentence_with_gemini(sent_notes, g_files)
                        st.session_state.sentence_confirmed = False
                    st.success("Leitura da sentença concluída! Confira abaixo o que foi deferido e indeferido.")
                except Exception as s_exc:
                    st.error(f"Erro ao analisar o processo via Gemini: {s_exc}")

    s_ext = st.session_state.sentence_ai_data or {}
    s_tc = st.session_state.sentence_timecard_result

    if s_ext:
        with st.container(border=True):
            st.subheader("Diagnóstico do Dispositivo da Sentença (Trânsito em Julgado)")
            if s_ext.get("dispositivo_resumo"):
                st.info(f"**Resumo do Comando Condenatório:** {s_ext['dispositivo_resumo']}")
            if s_ext.get("parametros_liquidacao"):
                st.warning(f"**Critérios de Liquidação Fixados pelo Juízo:** {s_ext['parametros_liquidacao']}")
            if s_ext.get("divergencias_calculo_adverso"):
                st.error(f"**Divergências / Erros Identificados no Cálculo Adverso:** {s_ext['divergencias_calculo_adverso']}")

            col_def, col_indef = st.columns(2)
            with col_def:
                st.markdown("#### ✅ Verbas DEFERIDAS (Entram no Cálculo)")
                for item in (s_ext.get("verbas_deferidas") or ["Nenhuma verba listada"]):
                    st.markdown(f"- ✅ {item}")
            with col_indef:
                st.markdown("#### ❌ Verbas IMPROCEDENTES (Excluídas)")
                for item in (s_ext.get("verbas_indeferidas") or ["Nenhum pedido improcedente informado"]):
                    st.markdown(f"- ❌ {item}")

            st.session_state.sentence_confirmed = st.checkbox(
                "Confirmo que conferi o dispositivo da sentença e os parâmetros extraídos antes de liquidar",
                value=st.session_state.sentence_confirmed,
                key="chk_confirm_sentence_liq",
            )

    def _parse_dt_sent(val: Optional[str], fallback: date) -> date:
        if not val:
            return fallback
        try:
            return datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
        except Exception:
            return fallback

    with st.form("sentence_liquidation_form"):
        st.subheader("2. Parâmetros da Liquidação de Sentença (Art. 879 CLT)")
        s_timbre = st.text_input(
            "Timbre do Escritório / Advogado (PDF)",
            value=str(current_user.get("office_header") or "Escritório de Advocacia Trabalhista — Liquidação de Sentença"),
        )
        sp1, sp2, sp3, sp4 = st.columns(4)
        s_proc = sp1.text_input("Número do Processo (CNJ) *", value=str(s_ext.get("process_number") or "0000000-00.2025.5.04.0000"))
        s_reclamante = sp2.text_input("Exequente / Reclamante *", value=str(s_ext.get("worker_name") or ""))
        s_reclamada = sp3.text_input("Executada / Reclamada *", value=str(s_ext.get("employer_name") or ""))
        s_cargo = sp4.text_input("Cargo / Função", value=str(s_ext.get("role_title") or ""))

        s_motivos = ["sem_justa_causa", "rescisao_indireta", "pedido_demissao", "acordo", "justa_causa"]
        s_mot_ext = s_ext.get("termination_type") if s_ext.get("termination_type") in s_motivos else "sem_justa_causa"

        sd1, sd2, sd3, sd4 = st.columns(4)
        s_adm = sd1.date_input("Admissão Reconhecida *", value=_parse_dt_sent(s_ext.get("admission_date"), date(2024, 2, 1)), key="s_adm_dt")
        s_dism = sd2.date_input("Rescisão Reconhecida *", value=_parse_dt_sent(s_ext.get("dismissal_date"), date(2026, 10, 31)), key="s_dism_dt")
        s_motivo = sd3.selectbox("Modalidade da Rescisão", s_motivos, index=s_motivos.index(s_mot_ext))
        s_divisor = sd4.number_input("Divisor Fixado", min_value=1, max_value=300, value=max(1, min(300, f_int(s_ext.get("monthly_divisor"), 220))))

        sr1, sr2, sr3, sr4, sr5 = st.columns([2, 2, 2.2, 1.5, 2])
        s_sal = sr1.number_input("Salário-Base (R$) *", min_value=0.0, value=max(0.0, f_float(s_ext.get("base_salary"), 0.0)), step=50.0)
        s_comis = sr2.number_input("Comissões / Integrações (R$/mês)", min_value=0.0, value=max(0.0, f_float(s_ext.get("commissions"), 0.0)), step=50.0)
        s_adic_opts = ["Nenhum", "Insalubridade (sobre Salário Mínimo)", "Periculosidade (sobre Salário-Base)"]
        s_raw_adic = str(s_ext.get("adicional_sugerido") or "Nenhum")
        s_adic_idx = 1 if "Insalubridade" in s_raw_adic else (2 if "Periculosidade" in s_raw_adic else 0)
        s_adic = sr3.selectbox("Adicional Deferido", s_adic_opts, index=s_adic_idx)
        s_adic_pct = sr4.number_input("Percentual (%)", min_value=0.0, max_value=100.0, value=max(0.0, min(100.0, f_float(s_ext.get("adicional_percentual_sugerido"), 0.0))), step=5.0)
        s_deducao = sr5.number_input("Dedução Valores Pagos (R$)", min_value=0.0, value=max(0.0, f_float(s_ext.get("trct_paid_total"), 0.0)), step=100.0)

        if s_tc is not None:
            st.info(
                f"Cartão de ponto dos autos apurado: **{s_tc['he50_hours']}h** HE 50% | "
                f"**{s_tc['he100_hours']}h** HE 100% | **{s_tc['intrajornada_hours']}h** intrajornada | "
                f"**{s_tc['night_ficta_hours']}h** noturnas fictas."
            )
            s_he50_sem = 0.0
            s_dom_mes = 0
            s_intra_min = 0
            s_not_sem = 0.0
        else:
            sj1, sj2, sj3, sj4 = st.columns(4)
            s_he50_sem = sj1.number_input("HE 50% / Semana Deferidas", min_value=0.0, value=max(0.0, f_float(s_ext.get("weekly_he50_hours"), 0.0)), step=0.5)
            s_dom_mes = sj2.number_input("Domingos/Fer. 100% / Mês", min_value=0, max_value=10, value=max(0, min(10, f_int(s_ext.get("monthly_sundays_worked"), 0))))
            s_intra_min = sj3.number_input("Intrajornada Suprimida (min/dia)", min_value=0, max_value=120, value=max(0, min(120, f_int(s_ext.get("suppressed_break_minutes"), 0))), step=5)
            s_not_sem = sj4.number_input("Horas Noturnas / Semana", min_value=0.0, value=max(0.0, f_float(s_ext.get("weekly_night_hours"), 0.0)), step=0.5)

        st.subheader("3. Verbas Deferidas no Dispositivo da Sentença")
        sv1, sv2, sv3 = st.columns(3)
        s_ped_saldo = sv1.checkbox("Saldo de Salário Deferido", value=bool(s_ext.get("sugerir_saldo_salario", False)))
        s_dias_saldo = sv1.number_input("Dias Saldo de Salário", min_value=0, max_value=31, value=int(s_dism.day))
        s_ped_reflexos = sv1.checkbox("Reflexos em 13º, Férias + 1/3 e Aviso", value=True)

        s_ped_aviso = sv2.checkbox("Aviso-Prévio Indenizado Deferido", value=bool(s_ext.get("sugerir_aviso_previo", False)))
        s_ped_13_fer = sv2.checkbox("13º e/ou Férias + 1/3 Deferidos", value=bool(s_ext.get("sugerir_13_e_ferias", False)))
        s_ped_477 = sv2.checkbox("Multa Art. 477 CLT Deferida", value=bool(s_ext.get("sugerir_multa_477", False)))
        s_ped_467 = sv2.checkbox("Multa Art. 467 CLT Deferida", value=bool(s_ext.get("sugerir_multa_467", False)))

        s_ped_fgts = sv3.checkbox("FGTS + Multa Rescisória Deferidos", value=bool(s_ext.get("sugerir_fgts_e_multa", False)))
        s_multa_fgts_pct = sv3.number_input("Percentual Multa FGTS (%)", min_value=0.0, max_value=100.0, value=max(0.0, min(100.0, f_float(s_ext.get("multa_fgts_percentual_sugerido"), 40.0))), step=5.0)
        s_fgts_dif = sv3.number_input("Diferenças FGTS Contrato (R$)", min_value=0.0, value=max(0.0, f_float(s_ext.get("fgts_nao_depositado_contrato"), 0.0)), step=100.0)
        s_estab_m = sv3.number_input("Meses Estabilidade Deferidos", min_value=0, max_value=24, value=max(0, min(24, f_int(s_ext.get("meses_estabilidade_restantes"), 0))))
        s_restit = sv3.number_input("Devolução Descontos Deferida (R$)", min_value=0.0, value=max(0.0, f_float(s_ext.get("restituicao_descontos"), 0.0)), step=50.0)

        liquidar_btn = st.form_submit_button("Calcular Liquidação de Sentença e Gerar Planilha", type="primary")

    if liquidar_btn:
        if s_ext and not st.session_state.sentence_confirmed:
            st.error("Marque a caixa de confirmação de revisão da sentença antes de calcular a liquidação.")
        elif s_sal <= 0 or not s_reclamante.strip():
            st.error("Informe o nome do reclamante e o salário-base reconhecido na sentença.")
        elif s_dism < s_adm:
            st.error("A data de rescisão não pode ser anterior à admissão.")
        else:
            s_meses_contrato = max(1, (s_dism.year - s_adm.year) * 12 + (s_dism.month - s_adm.month) + 1)
            s_meses_presc = min(60, s_meses_contrato)
            s_semanas = Decimal(str(s_meses_presc)) * Decimal("4.3333")

            if s_tc is not None:
                tot_s_he50 = f_float(s_tc.get("he50_hours"), 0.0)
                tot_s_he100 = f_float(s_tc.get("he100_hours"), 0.0)
                tot_s_intra = f_float(s_tc.get("intrajornada_hours"), 0.0)
                tot_s_not = f_float(s_tc.get("night_ficta_hours"), 0.0)
            else:
                tot_s_he50 = float((Decimal(str(s_he50_sem)) * s_semanas).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_he100 = float((Decimal(str(s_dom_mes * 8)) * Decimal(str(s_meses_presc))).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_intra = float(((Decimal(str(s_intra_min)) / Decimal("60")) * (Decimal(str(s_meses_presc)) * Decimal("22"))).quantize(CENT, rounding=ROUND_HALF_UP))
                tot_s_not = float((Decimal(str(s_not_sem)) * s_semanas * Decimal("1.142857")).quantize(CENT, rounding=ROUND_HALF_UP))

            s_sm_vigente = get_historical_min_wage(s_dism)
            s_insal_mensal = Decimal("0.00")
            s_adic_lbl = "0%"
            s_pct_dec = Decimal(str(s_adic_pct)) / Decimal("100")
            if s_adic == "Insalubridade (sobre Salário Mínimo)" and s_adic_pct > 0:
                s_insal_mensal = (s_sm_vigente * s_pct_dec).quantize(CENT, rounding=ROUND_HALF_UP)
                s_adic_lbl = f"{s_adic_pct:g}% SM (NR-15)"
            elif s_adic == "Periculosidade (sobre Salário-Base)" and s_adic_pct > 0:
                s_insal_mensal = (d_money(s_sal) * s_pct_dec).quantize(CENT, rounding=ROUND_HALF_UP)
                s_adic_lbl = f"{s_adic_pct:g}% Sal. (NR-16)"

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
                "inc_intrajornada": tot_s_intra > 0,
                "intrajornada_hours": tot_s_intra,
                "inc_noturno": tot_s_not > 0,
                "night_hours": tot_s_not,
                "night_pct": 20,
                "ai_opinion": s_ext,
                "lawyer_narrative": sent_notes,
            })

            if s_restit > 0:
                r_dec = d_money(s_restit)
                res_sent["verbas"].append({
                    "Código": "VER-11",
                    "Verba": "Restituição de Descontos Deferida em Sentença",
                    "Período": f"{s_adm.strftime('%d/%m/%Y')} a {s_dism.strftime('%d/%m/%Y')}",
                    "Base de Cálculo": r_dec,
                    "Fórmula": "Devolução integral conforme comando sentencial",
                    "Quantidade": "1",
                    "Percentual": "100%",
                    "Valor Original": r_dec,
                    "Reflexos": Decimal("0.00"),
                    "Total Apurado": r_dec,
                    "Incide FGTS": False,
                })
                res_sent["bruto_atualizado"] += r_dec
                res_sent["liquido_devido"] += r_dec

            st.session_state.calculations_history = [res_sent] + st.session_state.calculations_history
            st.session_state.case_chat_threads[res_sent["audit_id"]] = []
            st.session_state.selected_chat_case_id = res_sent["audit_id"]
            supabase_upsert_case(res_sent, [], current_office_email)

            st.success(f"Liquidação de Sentença concluída e salva na nuvem! Código: **{res_sent['audit_id']}**")
            sk1, sk2, sk3 = st.columns(3)
            sk1.metric("Total Bruto Liquidado", fmt_brl(res_sent["bruto_atualizado"]))
            sk2.metric("(-) Deduções Autorizadas", fmt_brl(res_sent["abatimentos"]))
            sk3.metric("Total Líquido da Execução", fmt_brl(res_sent["liquido_devido"]))

            df_sent = pd.DataFrame(res_sent["verbas"])
            if not df_sent.empty:
                for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Total Apurado"]:
                    df_sent[col] = df_sent[col].apply(fmt_brl)
                st.dataframe(
                    df_sent[["Código", "Verba", "Quantidade", "Percentual", "Base de Cálculo", "Fórmula", "Valor Original", "Reflexos", "Total Apurado"]],
                    use_container_width=True,
                    hide_index=True,
                )

            if s_ext.get("minuta_peticao_liquidacao"):
                with st.expander("Petição de Apresentação de Cálculos de Liquidação (Art. 879 CLT)", expanded=True):
                    st.text_area("Minuta pronta para protocolar no PJe:", value=str(s_ext["minuta_peticao_liquidacao"]), height=240)

            sdl1, sdl2 = st.columns(2)
            sdl1.download_button(
                "Baixar Planilha de Liquidação (.XLSX)",
                data=build_complete_excel_bytes(res_sent),
                file_name=f"Liquidacao_Sentenca_{res_sent['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            sdl2.download_button(
                "Baixar Relatório de Liquidação (.PDF)",
                data=build_printable_dossier_pdf_bytes(res_sent),
                file_name=f"Liquidacao_Sentenca_{res_sent['audit_id']}.pdf",
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
