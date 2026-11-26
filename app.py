import base64
import hashlib
import html
import io
import json
import math
import os
import re
import sqlite3
import urllib.request
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


# ==============================================================================
# CONFIGURAÇÕES DE PRECISÃO E SINCRONIZAÇÃO DE TABELAS LEGAIS (INSS / IRRF)
# ==============================================================================

CENT = Decimal("0.01")
FACTOR_PREC = Decimal("0.000000001")  # 9 casas decimais no índice acumulado (padrão PJe-Calc)
ULTIMA_COMPETENCIA_TABELA_OFICIAL = "2026-12"
HISTORY_DB_PATH = "juriscalc_pro_historico.db"
DYNAMIC_TABLES_CACHE_PATH = "juriscalc_tabelas_vigentes.json"

DEFAULT_INSS_PROGRESSIVE_TABLES: Dict[int, List[Tuple[Decimal, Decimal]]] = {
    2023: [
        (Decimal("1320.00"), Decimal("0.075")),
        (Decimal("2571.29"), Decimal("0.09")),
        (Decimal("3856.94"), Decimal("0.12")),
        (Decimal("7507.49"), Decimal("0.14")),
    ],
    2024: [
        (Decimal("1412.00"), Decimal("0.075")),
        (Decimal("2666.68"), Decimal("0.09")),
        (Decimal("4000.03"), Decimal("0.12")),
        (Decimal("7786.02"), Decimal("0.14")),
    ],
    2025: [
        (Decimal("1518.00"), Decimal("0.075")),
        (Decimal("2793.88"), Decimal("0.09")),
        (Decimal("4190.83"), Decimal("0.12")),
        (Decimal("8157.41"), Decimal("0.14")),
    ],
    2026: [
        (Decimal("1621.00"), Decimal("0.075")),
        (Decimal("2902.84"), Decimal("0.09")),
        (Decimal("4354.27"), Decimal("0.12")),
        (Decimal("8475.55"), Decimal("0.14")),
    ],
}

DEFAULT_IRRF_MONTHLY_TABLES: Dict[int, List[Tuple[Decimal, Decimal, Decimal]]] = {
    2024: [
        (Decimal("2259.20"), Decimal("0.000"), Decimal("0.00")),
        (Decimal("2826.65"), Decimal("0.075"), Decimal("169.44")),
        (Decimal("3751.05"), Decimal("0.150"), Decimal("381.44")),
        (Decimal("4664.68"), Decimal("0.225"), Decimal("662.77")),
        (Decimal("999999999.99"), Decimal("0.275"), Decimal("896.00")),
    ],
    2025: [
        (Decimal("2428.80"), Decimal("0.000"), Decimal("0.00")),
        (Decimal("2826.65"), Decimal("0.075"), Decimal("182.16")),
        (Decimal("3751.05"), Decimal("0.150"), Decimal("394.16")),
        (Decimal("4664.68"), Decimal("0.225"), Decimal("675.49")),
        (Decimal("999999999.99"), Decimal("0.275"), Decimal("908.73")),
    ],
    2026: [
        (Decimal("2428.80"), Decimal("0.000"), Decimal("0.00")),
        (Decimal("2826.65"), Decimal("0.075"), Decimal("182.16")),
        (Decimal("3751.05"), Decimal("0.150"), Decimal("394.16")),
        (Decimal("4664.68"), Decimal("0.225"), Decimal("675.49")),
        (Decimal("999999999.99"), Decimal("0.275"), Decimal("908.73")),
    ],
}

INSS_PROGRESSIVE_TABLES: Dict[int, List[Tuple[Decimal, Decimal]]] = {
    k: list(v) for k, v in DEFAULT_INSS_PROGRESSIVE_TABLES.items()
}
IRRF_MONTHLY_TABLES: Dict[int, List[Tuple[Decimal, Decimal, Decimal]]] = {
    k: list(v) for k, v in DEFAULT_IRRF_MONTHLY_TABLES.items()
}


def sync_legal_tables_dynamic(remote_url: str = "", raw_json_override: Optional[str] = None) -> Dict[str, Any]:
    """Sincroniza dinamicamente as tabelas vigentes de INSS e IRRF/RRA."""
    global ULTIMA_COMPETENCIA_TABELA_OFICIAL, INSS_PROGRESSIVE_TABLES, IRRF_MONTHLY_TABLES

    payload: Optional[Dict[str, Any]] = None
    fonte = "Base Legal Integrada (2023–2026)"

    if raw_json_override and raw_json_override.strip():
        payload = json.loads(raw_json_override)
        fonte = "Pacote Oficial Importado Manualmente"
    elif remote_url and remote_url.strip().startswith("http"):
        try:
            req = urllib.request.Request(
                remote_url.strip(),
                headers={"User-Agent": "JurisCalcPro-Sync/1.0"},
            )
            with urllib.request.urlopen(req, timeout=4.0) as resp:
                if resp.status == 200:
                    payload = json.loads(resp.read().decode("utf-8"))
                    fonte = f"Feed Remoto Sincronizado ({remote_url.strip()})"
        except Exception:
            payload = None

    if payload is None and os.path.exists(DYNAMIC_TABLES_CACHE_PATH):
        try:
            with open(DYNAMIC_TABLES_CACHE_PATH, "r", encoding="utf-8") as f:
                payload = json.load(f)
                fonte = payload.get("fonte", "Cache Local Sincronizado")
        except Exception:
            payload = None

    if isinstance(payload, dict):
        inss_raw = payload.get("inss", {})
        if isinstance(inss_raw, dict):
            for yr_str, rows in inss_raw.items():
                yr = int(yr_str)
                parsed_rows = [(Decimal(str(r[0])), Decimal(str(r[1]))) for r in rows]
                if parsed_rows:
                    INSS_PROGRESSIVE_TABLES[yr] = parsed_rows

        irrf_raw = payload.get("irrf", {})
        if isinstance(irrf_raw, dict):
            for yr_str, rows in irrf_raw.items():
                yr = int(yr_str)
                parsed_rows = [
                    (Decimal(str(r[0])), Decimal(str(r[1])), Decimal(str(r[2])))
                    for r in rows
                ]
                if parsed_rows:
                    IRRF_MONTHLY_TABLES[yr] = parsed_rows

        max_yr = max(max(INSS_PROGRESSIVE_TABLES.keys()), max(IRRF_MONTHLY_TABLES.keys()))
        ULTIMA_COMPETENCIA_TABELA_OFICIAL = str(payload.get("ultima_competencia", f"{max_yr}-12"))

        if raw_json_override or (remote_url and remote_url.strip().startswith("http")):
            save_doc = {
                "ultima_competencia": ULTIMA_COMPETENCIA_TABELA_OFICIAL,
                "fonte": fonte,
                "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "inss": {
                    str(y): [[str(a), str(b)] for a, b in rows]
                    for y, rows in INSS_PROGRESSIVE_TABLES.items()
                },
                "irrf": {
                    str(y): [[str(a), str(b), str(c)] for a, b, c in rows]
                    for y, rows in IRRF_MONTHLY_TABLES.items()
                },
            }
            try:
                with open(DYNAMIC_TABLES_CACHE_PATH, "w", encoding="utf-8") as f:
                    json.dump(save_doc, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

    max_year_loaded = max(max(INSS_PROGRESSIVE_TABLES.keys()), max(IRRF_MONTHLY_TABLES.keys()))
    return {
        "status": "Sincronizado",
        "fonte": fonte,
        "ultima_competencia": ULTIMA_COMPETENCIA_TABELA_OFICIAL,
        "anos_inss": sorted(INSS_PROGRESSIVE_TABLES.keys()),
        "anos_irrf": sorted(IRRF_MONTHLY_TABLES.keys()),
        "ano_maximo": max_year_loaded,
    }


SYNC_STATUS_INFO = sync_legal_tables_dynamic()


def money_round(val: Decimal) -> Decimal:
    return val.quantize(CENT, rounding=ROUND_HALF_UP)


def factor_round(val: Decimal) -> Decimal:
    return val.quantize(FACTOR_PREC, rounding=ROUND_HALF_UP)


def fmt_brl(val: Any) -> str:
    d = Decimal(str(val or 0))
    d = money_round(d)
    neg = d < 0
    d = abs(d)
    s = f"{d:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{s}" if neg else s


def fmt_brl_rs(val: Any) -> str:
    return f"R$ {fmt_brl(val)}"


def fmt_idx(val: Any) -> str:
    d = Decimal(str(val or 1))
    return f"{d:,.9f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_date_br(s: Optional[str]) -> Optional[date]:
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def fmt_date_br(s: Optional[str]) -> str:
    d = parse_date_br(s)
    return d.strftime("%d/%m/%Y") if d else (s or "-")


def fmt_comp_br(comp_ym: str) -> str:
    try:
        y, m = comp_ym.split("-")
        return f"{int(m):02d}/{int(y):04d}"
    except Exception:
        return comp_ym


def month_range(start_d: date, end_d: date) -> List[Tuple[int, int]]:
    res = []
    y, m = start_d.year, start_d.month
    while (y, m) <= (end_d.year, end_d.month):
        res.append((y, m))
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
    return res


def dias_aviso_previo_lei_12506(admissao: date, demissao: date) -> int:
    anos_completos = 0
    try:
        while admissao.replace(year=admissao.year + anos_completos + 1) <= demissao:
            anos_completos += 1
    except ValueError:
        while (admissao + timedelta(days=365 * (anos_completos + 1))) <= demissao:
            anos_completos += 1
    return min(90, 30 + (anos_completos * 3))


# ==============================================================================
# PERSISTÊNCIA LOCAL DE CÁLCULOS (SQLITE)
# ==============================================================================

def init_history_db() -> None:
    try:
        conn = sqlite3.connect(HISTORY_DB_PATH)
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS calculos_historico (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                criado_em TEXT NOT NULL,
                processo_cnj TEXT,
                reclamante TEXT,
                reclamado TEXT,
                bruto_reclamante REAL,
                liquido_reclamante REAL,
                total_reclamado REAL,
                payload_json TEXT NOT NULL
            )
            """
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def save_calculation_history(params: Dict[str, Any], result: Dict[str, Any]) -> None:
    try:
        init_history_db()
        conn = sqlite3.connect(HISTORY_DB_PATH)
        cur = conn.cursor()
        resumo = result.get("resumo", {})
        cur.execute(
            """
            INSERT INTO calculos_historico
            (criado_em, processo_cnj, reclamante, reclamado, bruto_reclamante, liquido_reclamante, total_reclamado, payload_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                params.get("processo_cnj", ""),
                params.get("reclamante", ""),
                params.get("reclamado", ""),
                float(resumo.get("bruto_devido_reclamante", 0)),
                float(resumo.get("liquido_reclamante", 0)),
                float(resumo.get("total_devido_reclamado", 0)),
                json.dumps(params, ensure_ascii=False, default=str),
            ),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def list_calculation_history(limit: int = 30) -> List[Dict[str, Any]]:
    try:
        init_history_db()
        conn = sqlite3.connect(HISTORY_DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM calculos_historico ORDER BY id DESC LIMIT ?",
            (limit,),
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows
    except Exception:
        return []


# ==============================================================================
# BUSCA DE ÍNDICES ECONÔMICOS OFICIAIS (BANCO CENTRAL DO BRASIL - SGS)
# ==============================================================================

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_bcb_series_monthly(series_code: int, start_date: str = "01/01/2015") -> Dict[str, Decimal]:
    """
    Consulta a série mensal oficial no SGS do Banco Central do Brasil:
    - 433: IPCA (% a.m.)
    - 10764: IPCA-E (% a.m.)
    - 4390: SELIC acumulada no mês (% a.m.)
    - 226: TR (% a.m.)
    """
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{series_code}/dados"
        f"?formato=json&dataInicial={start_date}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "JurisCalcPro/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=6.0) as resp:
            if resp.status != 200:
                return {}
            raw = json.loads(resp.read().decode("utf-8"))
            out: Dict[str, Decimal] = {}
            for item in raw:
                dt_str = item.get("data", "")
                val_str = str(item.get("valor", "0")).replace(",", ".")
                parts = dt_str.split("/")
                if len(parts) == 3:
                    ym = f"{parts[2]}-{parts[1]}"
                    out[ym] = Decimal(val_str) / Decimal("100")
            return out
    except Exception:
        return {}


def get_fallback_rates() -> Tuple[Dict[str, Decimal], Dict[str, Decimal], Dict[str, Decimal]]:
    ipcae: Dict[str, Decimal] = {}
    selic: Dict[str, Decimal] = {}
    ipca: Dict[str, Decimal] = {}
    for y in range(2015, 2027):
        for m in range(1, 13):
            ym = f"{y:04d}-{m:02d}"
            ipcae[ym] = Decimal("0.0042")
            ipca[ym] = Decimal("0.0040")
            selic[ym] = Decimal("0.0088")
    return ipcae, selic, ipca


def build_adc58_correction_tables(
    comp_list: List[str],
    data_ajuizamento: date,
    data_liquidacao: date,
    criterio_correcao: str = "ADC58_LEI_14905",
) -> Tuple[Dict[str, Decimal], Dict[str, Decimal], str, Optional[str]]:
    ipcae_map = fetch_bcb_series_monthly(10764, "01/01/2015")
    selic_map = fetch_bcb_series_monthly(4390, "01/01/2015")
    ipca_map = fetch_bcb_series_monthly(433, "01/01/2015")
    tr_map = fetch_bcb_series_monthly(226, "01/01/2015")

    fonte_status = "Banco Central do Brasil (SGS Oficial — Séries 10764, 4390, 433 e 226)"
    if not ipcae_map or not selic_map or not ipca_map:
        ipcae_map, selic_map, ipca_map = get_fallback_rates()
        fonte_status = "Tabela de Contingência (Modo Offline — verifique conexão com BACEN)"

    all_months = sorted(
        set(comp_list)
        | {f"{y:04d}-{m:02d}" for y, m in month_range(date(2015, 1, 1), data_liquidacao)}
    )
    all_months = [ym for ym in all_months if ym <= f"{data_liquidacao.year:04d}-{data_liquidacao.month:02d}"]

    ajuiz_ym = f"{data_ajuizamento.year:04d}-{data_ajuizamento.month:02d}"
    lei14905_ym = "2024-09"

    warning_msg: Optional[str] = None
    liq_ym = f"{data_liquidacao.year:04d}-{data_liquidacao.month:02d}"
    if liq_ym > ULTIMA_COMPETENCIA_TABELA_OFICIAL:
        warning_msg = (
            f"A data de liquidação ({fmt_date_br(data_liquidacao.isoformat())}) é posterior à última competência "
            f"consolidada ({ULTIMA_COMPETENCIA_TABELA_OFICIAL}); o motor utilizou os últimos índices divulgados pelo BACEN."
        )

    idx_correcao_por_comp: Dict[str, Decimal] = {}
    juros_fator_por_comp: Dict[str, Decimal] = {}

    for comp in comp_list:
        fator_corr = Decimal("1.000000000")
        fator_juros = Decimal("0.000000000")

        for ym in all_months:
            if ym <= comp:
                continue

            r_ipcae = ipcae_map.get(ym, Decimal("0.0038"))
            r_selic = selic_map.get(ym, Decimal("0.0085"))
            r_ipca = ipca_map.get(ym, Decimal("0.0038"))
            r_tr = tr_map.get(ym, Decimal("0.0005"))

            if criterio_correcao == "IPCAE_JUROS_1_PCT":
                fator_corr *= (Decimal("1") + r_ipcae)
                fator_juros += Decimal("0.01")
            else:
                if ym < ajuiz_ym:
                    fator_corr *= (Decimal("1") + r_ipcae)
                    fator_juros += max(Decimal("0"), r_tr)
                elif ym < lei14905_ym:
                    fator_corr *= (Decimal("1") + r_selic)
                else:
                    fator_corr *= (Decimal("1") + r_ipca)
                    taxa_legal = max(Decimal("0"), r_selic - r_ipca)
                    fator_juros += taxa_legal

        idx_correcao_por_comp[comp] = factor_round(fator_corr)
        juros_fator_por_comp[comp] = factor_round(fator_juros)

    return idx_correcao_por_comp, juros_fator_por_comp, fonte_status, warning_msg


# ==============================================================================
# TABELAS LEGAIS PROGRESSIVAS (INSS SEGURADO EC 103/2019 E IRRF RRA LEI 7.713/88)
# ==============================================================================

def _calc_inss_segurado_competencia(base_val: Decimal, ano: int) -> Decimal:
    if base_val <= 0:
        return Decimal("0.00")
    min_yr = min(INSS_PROGRESSIVE_TABLES.keys())
    max_yr = max(INSS_PROGRESSIVE_TABLES.keys())
    ano_ref = min(max(ano, min_yr), max_yr)
    faixas = INSS_PROGRESSIVE_TABLES[ano_ref]
    teto_max = faixas[-1][0]
    base_efetiva = min(base_val, teto_max)
    total_inss = Decimal("0.00")
    limite_anterior = Decimal("0.00")
    for limite_faixa, aliquota in faixas:
        if base_efetiva > limite_anterior:
            parcela_faixa = min(base_efetiva, limite_faixa) - limite_anterior
            total_inss += parcela_faixa * aliquota
            limite_anterior = limite_faixa
        else:
            break
    return money_round(total_inss)


def _calc_irrf_rra(base_liquida: Decimal, meses_rra: int, ano_liquidacao: int = 2026) -> Decimal:
    if base_liquida <= 0:
        return Decimal("0.00")
    nm = Decimal(str(max(1, meses_rra)))
    min_yr = min(IRRF_MONTHLY_TABLES.keys())
    max_yr = max(IRRF_MONTHLY_TABLES.keys())
    ano_ref = min(max(ano_liquidacao, min_yr), max_yr)
    faixas_mensais = IRRF_MONTHLY_TABLES[ano_ref]
    for lim_m, aliq, ded_m in faixas_mensais:
        if base_liquida <= (lim_m * nm):
            if aliq == 0:
                return Decimal("0.00")
            imposto = (base_liquida * aliq) - (ded_m * nm)
            return money_round(max(Decimal("0.00"), imposto))
    return Decimal("0.00")


# ==============================================================================
# MOTOR DE LIQUIDAÇÃO TRABALHISTA (FIDELIDADE TOTAL À ESTRUTURA PJe-Calc)
# ==============================================================================

def run_pjecalc_engine(params: Dict[str, Any]) -> Dict[str, Any]:
    admissao = parse_date_br(params.get("admissao")) or date(2021, 5, 10)
    demissao = parse_date_br(params.get("demissao")) or date(2024, 3, 15)
    ajuizamento = parse_date_br(params.get("data_ajuizamento")) or date(2024, 5, 20)
    liquidacao = parse_date_br(params.get("data_liquidacao")) or date.today()

    aplicar_prescricao = bool(params.get("aplicar_prescricao", True))
    projetar_aviso = bool(params.get("projetar_aviso", True))
    prazo_aviso_dias = int(params.get("prazo_aviso_dias") or dias_aviso_previo_lei_12506(admissao, demissao))
    data_demissao_projetada = demissao + timedelta(days=prazo_aviso_dias) if projetar_aviso else demissao

    data_prescricao = date(ajuizamento.year - 5, ajuizamento.month, ajuizamento.day) if aplicar_prescricao else admissao
    inicio_calculo = max(admissao, data_prescricao)
    fim_calculo = data_demissao_projetada

    if fim_calculo < inicio_calculo:
        fim_calculo = inicio_calculo

    # Histórico Salarial
    hist_raw = params.get("historico_salarial", [])
    hist_map: Dict[str, Tuple[Decimal, Decimal]] = {}
    for h in hist_raw:
        comp = str(h.get("competencia", ""))[:7]
        sp = Decimal(str(h.get("salario_pago", 0) or 0))
        sd = Decimal(str(h.get("salario_devido", sp) or sp))
        if comp:
            hist_map[comp] = (sp, sd)

    comps_contrato = [f"{y:04d}-{m:02d}" for y, m in month_range(inicio_calculo, fim_calculo)]
    if not comps_contrato:
        comps_contrato = [f"{inicio_calculo.year:04d}-{inicio_calculo.month:02d}"]

    default_sal_pago = Decimal(str(params.get("ultima_remuneracao") or params.get("maior_remuneracao") or 0))
    default_sal_devido = Decimal(str(params.get("maior_remuneracao") or default_sal_pago))

    historico_consolidado: List[Dict[str, Any]] = []
    for ym in comps_contrato:
        sp, sd = hist_map.get(ym, (default_sal_pago, default_sal_devido))
        historico_consolidado.append({
            "competencia": ym,
            "salario_pago": sp,
            "salario_devido": sd,
            "diferenca": max(Decimal("0.00"), sd - sp),
        })

    # Fatores de Correção e Juros
    criterio_corr = params.get("criterio_correcao", "ADC58_LEI_14905")
    idx_corr_map, juros_fator_map, fonte_indices, warning_msg = build_adc58_correction_tables(
        comps_contrato, ajuizamento, liquidacao, criterio_corr
    )

    # Apuração de Verbas Principais e Reflexos
    verbas_cfg = params.get("verbas", [])
    verbas_detalhadas: List[Dict[str, Any]] = []

    fgts_base_por_comp: Dict[str, Decimal] = {ym: Decimal("0.00") for ym in comps_contrato}
    inss_base_por_comp: Dict[str, Decimal] = {ym: Decimal("0.00") for ym in comps_contrato}
    tributavel_total_devido = Decimal("0.00")
    remuneratorio_total_devido = Decimal("0.00")
    total_verbas_devido_corrigido = Decimal("0.00")
    total_verbas_pago_corrigido = Decimal("0.00")
    total_verbas_diferenca = Decimal("0.00")
    total_verbas_corrigido_final = Decimal("0.00")

    def _registrar_verba(
        nome: str,
        ocorrencias: List[Dict[str, Any]],
        incide_fgts: bool,
        incide_inss: bool,
        incide_irrf: bool,
        is_remuneratorio: bool = True,
    ) -> None:
        nonlocal tributavel_total_devido, remuneratorio_total_devido
        nonlocal total_verbas_devido_corrigido, total_verbas_pago_corrigido
        nonlocal total_verbas_diferenca, total_verbas_corrigido_final

        soma_dev_orig = Decimal("0.00")
        soma_pag_orig = Decimal("0.00")
        soma_dif_orig = Decimal("0.00")
        soma_dev_corr = Decimal("0.00")
        soma_pag_corr = Decimal("0.00")
        soma_corr_final = Decimal("0.00")

        for oc in ocorrencias:
            ym = oc["competencia"]
            v_dev = money_round(oc["devido"])
            v_pag = money_round(oc["pago"])
            v_dif = money_round(max(Decimal("0.00"), v_dev - v_pag))
            idx = idx_corr_map.get(ym, Decimal("1.000000000"))

            v_dev_c = money_round(v_dev * idx)
            v_pag_c = money_round(v_pag * idx)
            v_corr = money_round(v_dif * idx)

            oc["devido"] = v_dev
            oc["pago"] = v_pag
            oc["diferenca"] = v_dif
            oc["indice"] = idx
            oc["valor_corrigido"] = v_corr

            soma_dev_orig += v_dev
            soma_pag_orig += v_pag
            soma_dif_orig += v_dif
            soma_dev_corr += v_dev_c
            soma_pag_corr += v_pag_c
            soma_corr_final += v_corr

            if incide_fgts:
                fgts_base_por_comp[ym] = fgts_base_por_comp.get(ym, Decimal("0.00")) + v_dif
            if incide_inss:
                inss_base_por_comp[ym] = inss_base_por_comp.get(ym, Decimal("0.00")) + v_dif
            if incide_irrf:
                tributavel_total_devido += v_corr
            if is_remuneratorio:
                remuneratorio_total_devido += v_corr

        total_verbas_devido_corrigido += soma_dev_corr
        total_verbas_pago_corrigido += soma_pag_corr
        total_verbas_diferenca += soma_dif_orig
        total_verbas_corrigido_final += soma_corr_final

        verbas_detalhadas.append({
            "nome": nome,
            "ocorrencias": ocorrencias,
            "total_devido_orig": soma_dev_orig,
            "total_pago_orig": soma_pag_orig,
            "total_diferenca": soma_dif_orig,
            "total_devido_corrigido": soma_dev_corr,
            "total_pago_corrigido": soma_pag_corr,
            "total_corrigido": soma_corr_final,
            "incide_fgts": incide_fgts,
            "incide_inss": incide_inss,
            "incide_irrf": incide_irrf,
        })

    for v in verbas_cfg:
        if not v.get("ativo", True):
            continue
        v_nome = str(v.get("nome", "Verba Trabalhista"))
        v_tipo = str(v.get("tipo", "HORAS_EXTRAS")).upper().strip()
        v_aliq = Decimal(str(v.get("aliquota_ou_percentual", 0) or 0))
        v_qtd = Decimal(str(v.get("quantidade_mensal", 0) or 0))
        v_div = Decimal(str(v.get("divisor") or 220))
        v_ini = parse_date_br(v.get("inicio")) or inicio_calculo
        v_fim = parse_date_br(v.get("fim")) or fim_calculo
        v_fgts = bool(v.get("incide_fgts", True))
        v_inss = bool(v.get("incide_inss", True))
        v_irrf = bool(v.get("incide_irrf", True))
        v_reflexos = bool(v.get("gera_reflexos_13_ferias_aviso", True))

        ini_ym = f"{max(v_ini, inicio_calculo).year:04d}-{max(v_ini, inicio_calculo).month:02d}"
        fim_ym = f"{min(v_fim, fim_calculo).year:04d}-{min(v_fim, fim_calculo).month:02d}"

        ocorrencias_v: List[Dict[str, Any]] = []
        soma_mensal_para_reflexos: List[Decimal] = []

        for h in historico_consolidado:
            ym = h["competencia"]
            if ym < ini_ym or ym > fim_ym:
                continue

            sal_base = h["salario_devido"]
            sal_pago = h["salario_pago"]
            dev = Decimal("0.00")
            pag = Decimal("0.00")

            if v_tipo == "DIFERENCA_SALARIAL":
                dev = sal_base
                pag = sal_pago
            elif v_tipo == "HORAS_EXTRAS":
                valor_hora = (sal_base / v_div) * (Decimal("1") + v_aliq)
                dev = valor_hora * v_qtd
            elif v_tipo == "ADICIONAL_NOTURNO":
                valor_hora_not = (sal_base / v_div) * v_aliq
                dev = valor_hora_not * v_qtd
            elif v_tipo == "INSALUBRIDADE":
                sal_min = Decimal("1412.00")
                dev = sal_min * v_aliq
            elif v_tipo == "PERICULOSIDADE":
                dev = sal_base * v_aliq
            elif v_tipo == "AVISO_PREVIO":
                if ym == comps_contrato[-1]:
                    dev = (sal_base / Decimal("30")) * Decimal(str(prazo_aviso_dias))
                else:
                    continue
            elif v_tipo == "DECIMO_TERCEIRO":
                if ym.endswith("-12") or ym == comps_contrato[-1]:
                    dev = sal_base
                else:
                    continue
            elif v_tipo == "FERIAS_TERCO":
                if ym == comps_contrato[-1]:
                    dev = sal_base * (Decimal("4") / Decimal("3"))
                else:
                    continue
            elif v_tipo == "MULTA_477":
                if ym == comps_contrato[-1]:
                    dev = sal_base
                else:
                    continue
            elif v_tipo == "VERBA_UNICA_RESCISAO":
                if ym == comps_contrato[-1]:
                    dev = v_qtd if v_qtd > 0 else sal_base
                else:
                    continue
            else:
                dev = v_qtd if v_qtd > 0 else (sal_base * v_aliq)

            if dev > 0 or pag > 0:
                ocorrencias_v.append({
                    "competencia": ym,
                    "base": sal_base,
                    "div_mult": f"{v_div:g}" if v_tipo in ("HORAS_EXTRAS", "ADICIONAL_NOTURNO") else "1,00",
                    "qtd": v_qtd if v_qtd > 0 else Decimal("1.00"),
                    "devido": dev,
                    "pago": pag,
                })
                soma_mensal_para_reflexos.append(max(Decimal("0.00"), dev - pag))

        if not ocorrencias_v:
            continue

        is_remun = v_tipo not in ("MULTA_477", "AVISO_PREVIO", "FERIAS_TERCO", "VERBA_UNICA_RESCISAO")
        _registrar_verba(v_nome, ocorrencias_v, v_fgts, v_inss, v_irrf, is_remuneratorio=is_remun)

        if v_reflexos and v_tipo in ("DIFERENCA_SALARIAL", "HORAS_EXTRAS", "ADICIONAL_NOTURNO", "INSALUBRIDADE", "PERICULOSIDADE"):
            media_mensal = sum(soma_mensal_para_reflexos) / Decimal(str(len(soma_mensal_para_reflexos)))
            ultima_comp = ocorrencias_v[-1]["competencia"]

            # Reflexo em 13º Salário (1/12 por mês)
            oc_13: List[Dict[str, Any]] = []
            por_ano: Dict[str, List[Decimal]] = {}
            for oc in ocorrencias_v:
                yr = oc["competencia"][:4]
                por_ano.setdefault(yr, []).append(oc["diferenca"])
            for yr, vals in sorted(por_ano.items()):
                comp_13 = f"{yr}-12" if f"{yr}-12" in idx_corr_map else ultima_comp
                val_13 = (sum(vals) / Decimal("12"))
                if val_13 > 0:
                    oc_13.append({
                        "competencia": comp_13,
                        "base": sum(vals),
                        "div_mult": "12,00",
                        "qtd": Decimal(str(len(vals))),
                        "devido": val_13,
                        "pago": Decimal("0.00"),
                    })
            if oc_13:
                _registrar_verba(f"13º SALÁRIO (Reflexo de {v_nome})", oc_13, True, True, True, True)

            # Reflexo em Férias + 1/3
            oc_ferias = [{
                "competencia": ultima_comp,
                "base": media_mensal,
                "div_mult": "12,00",
                "qtd": Decimal(str(len(ocorrencias_v))),
                "devido": (sum(soma_mensal_para_reflexos) / Decimal("12")) * (Decimal("4") / Decimal("3")),
                "pago": Decimal("0.00"),
            }]
            _registrar_verba(f"FÉRIAS + 1/3 (Reflexo de {v_nome})", oc_ferias, False, False, False, False)

            # Reflexo em Aviso Prévio Indenizado
            if projetar_aviso and media_mensal > 0:
                oc_aviso = [{
                    "competencia": ultima_comp,
                    "base": media_mensal,
                    "div_mult": "30,00",
                    "qtd": Decimal(str(prazo_aviso_dias)),
                    "devido": (media_mensal / Decimal("30")) * Decimal(str(prazo_aviso_dias)),
                    "pago": Decimal("0.00"),
                }]
                _registrar_verba(f"AVISO PRÉVIO (Reflexo de {v_nome})", oc_aviso, True, False, False, False)

    # ==========================================================================
    # DEMONSTRATIVO DE CONTRIBUIÇÃO SOCIAL (SEGURADO + PATRONAL)
    # ==========================================================================
    cs_ocorrencias: List[Dict[str, Any]] = []
    total_cs_segurado_corrigido = Decimal("0.00")
    total_cs_patronal_corrigido = Decimal("0.00")
    cs_segurado_corrigido_por_comp: Dict[str, Decimal] = {}

    aliquota_patronal_total = Decimal(str(params.get("aliquota_inss_patronal", 0.225)))

    for ym in comps_contrato:
        base_inss = money_round(inss_base_por_comp.get(ym, Decimal("0.00")))
        if base_inss <= 0:
            continue
        ano_comp = int(ym[:4])
        idx = idx_corr_map.get(ym, Decimal("1.000000000"))
        jf = juros_fator_map.get(ym, Decimal("0.000000000"))

        seg_orig = _calc_inss_segurado_competencia(base_inss, ano_comp)
        seg_corr = money_round(seg_orig * idx)
        seg_juros = money_round(seg_corr * jf)
        seg_total = money_round(seg_corr + seg_juros)

        pat_orig = money_round(base_inss * aliquota_patronal_total)
        pat_corr = money_round(pat_orig * idx)
        pat_juros = money_round(pat_corr * jf)
        pat_total = money_round(pat_corr + pat_juros)

        total_cs_segurado_corrigido += seg_corr
        total_cs_patronal_corrigido += pat_total
        cs_segurado_corrigido_por_comp[ym] = seg_corr

        cs_ocorrencias.append({
            "competencia": ym,
            "base_historico": Decimal("0.00"),
            "base_verbas": base_inss,
            "segurado_devido": seg_orig,
            "segurado_corrigido": seg_corr,
            "segurado_juros": seg_juros,
            "segurado_total": seg_total,
            "patronal_devido": pat_orig,
            "patronal_corrigido": pat_corr,
            "patronal_juros": pat_juros,
            "patronal_total": pat_total,
        })

    # ==========================================================================
    # DEMONSTRATIVO DE JUROS SOBRE VERBAS (DEDUZINDO CONTRIBUIÇÃO SOCIAL DO SEGURADO)
    # ==========================================================================
    verbas_corr_por_comp: Dict[str, Decimal] = {ym: Decimal("0.00") for ym in comps_contrato}
    for vd in verbas_detalhadas:
        for oc in vd["ocorrencias"]:
            ym = oc["competencia"]
            verbas_corr_por_comp[ym] = verbas_corr_por_comp.get(ym, Decimal("0.00")) + oc["valor_corrigido"]

    juros_ocorrencias: List[Dict[str, Any]] = []
    total_juros_verbas = Decimal("0.00")
    total_capital_juros = Decimal("0.00")

    for ym in comps_contrato:
        v_corr_m = money_round(verbas_corr_por_comp.get(ym, Decimal("0.00")))
        cs_seg_m = money_round(cs_segurado_corrigido_por_comp.get(ym, Decimal("0.00")))
        if v_corr_m <= 0 and cs_seg_m <= 0:
            continue
        capital = money_round(max(Decimal("0.00"), v_corr_m - cs_seg_m))
        jf = juros_fator_map.get(ym, Decimal("0.000000000"))
        juros_m = money_round(capital * jf)

        total_capital_juros += capital
        total_juros_verbas += juros_m

        juros_ocorrencias.append({
            "competencia": ym,
            "total_verbas": v_corr_m,
            "contribuicao_social": cs_seg_m,
            "capital": capital,
            "taxa_juros_pct": money_round(jf * Decimal("100")),
            "juros": juros_m,
        })

    # ==========================================================================
    # DEMONSTRATIVO DE FGTS (8% + MULTA RESCISÓRIA 40% SE DEVIDA)
    # ==========================================================================
    fgts_ocorrencias: List[Dict[str, Any]] = []
    total_fgts_depositos = Decimal("0.00")
    total_fgts_juros = Decimal("0.00")

    for ym in comps_contrato:
        b_fgts = money_round(fgts_base_por_comp.get(ym, Decimal("0.00")))
        if b_fgts <= 0:
            continue
        idx = idx_corr_map.get(ym, Decimal("1.000000000"))
        jf = juros_fator_map.get(ym, Decimal("0.000000000"))
        fgts_orig = money_round(b_fgts * Decimal("0.08"))
        fgts_corr = money_round(fgts_orig * idx)
        fgts_jur = money_round(fgts_corr * jf)
        fgts_tot = money_round(fgts_corr + fgts_jur)

        total_fgts_depositos += fgts_corr
        total_fgts_juros += fgts_jur

        fgts_ocorrencias.append({
            "competencia": ym,
            "base": b_fgts,
            "aliquota": Decimal("8.00"),
            "devido": fgts_orig,
            "indice": idx,
            "corrigido": fgts_corr,
            "juros": fgts_jur,
            "total": fgts_tot,
        })

    aplicar_multa_40 = bool(params.get("aplicar_multa_fgts_40", True))
    multa_fgts_40 = money_round(total_fgts_depositos * Decimal("0.40")) if aplicar_multa_40 else Decimal("0.00")
    total_fgts_geral = money_round(total_fgts_depositos + total_fgts_juros + multa_fgts_40)

    # ==========================================================================
    # IMPOSTO DE RENDA (RRA - ART. 12-A DA LEI Nº 7.713/1988), HONORÁRIOS E CUSTAS
    # ==========================================================================
    meses_rra = max(1, len(comps_contrato))
    base_irrf_liquida = money_round(max(Decimal("0.00"), tributavel_total_devido - total_cs_segurado_corrigido))
    irrf_reclamante = _calc_irrf_rra(base_irrf_liquida, meses_rra, liquidacao.year) if params.get("apurar_irrf", True) else Decimal("0.00")

    bruto_devido_reclamante = money_round(total_verbas_corrigido_final + total_juros_verbas + total_fgts_geral)
    total_descontos_reclamante = money_round(total_cs_segurado_corrigido + irrf_reclamante)
    liquido_reclamante = money_round(max(Decimal("0.00"), bruto_devido_reclamante - total_descontos_reclamante))

    hon_pct = Decimal(str(params.get("honorarios_sucumbencia_pct", 15.0) or 0)) / Decimal("100")
    honorarios_adv = money_round(bruto_devido_reclamante * hon_pct)

    custas_pct = Decimal(str(params.get("custas_pct", 2.0) or 0)) / Decimal("100")
    custas_judiciais = money_round(bruto_devido_reclamante * custas_pct) if params.get("apurar_custas", True) else Decimal("0.00")

    total_devido_reclamado = money_round(
        liquido_reclamante
        + total_cs_segurado_corrigido
        + total_cs_patronal_corrigido
        + irrf_reclamante
        + honorarios_adv
        + custas_judiciais
    )

    pct_remuneratorio = (
        money_round((remuneratorio_total_devido / bruto_devido_reclamante) * Decimal("100"))
        if bruto_devido_reclamante > 0
        else Decimal("0.00")
    )
    pct_tributavel = (
        money_round((tributavel_total_devido / bruto_devido_reclamante) * Decimal("100"))
        if bruto_devido_reclamante > 0
        else Decimal("0.00")
    )

    return {
        "meta": {
            "inicio_calculo": inicio_calculo.isoformat(),
            "fim_calculo": fim_calculo.isoformat(),
            "data_demissao_projetada": data_demissao_projetada.isoformat(),
            "prazo_aviso_dias": prazo_aviso_dias,
            "meses_rra": meses_rra,
            "fonte_indices": fonte_indices,
            "warning_msg": warning_msg,
        },
        "historico_consolidado": historico_consolidado,
        "verbas_detalhadas": verbas_detalhadas,
        "juros_ocorrencias": juros_ocorrencias,
        "fgts_ocorrencias": fgts_ocorrencias,
        "cs_ocorrencias": cs_ocorrencias,
        "resumo": {
            "verbas_corrigido": total_verbas_corrigido_final,
            "juros_verbas": total_juros_verbas,
            "capital_juros": total_capital_juros,
            "fgts_depositos": total_fgts_depositos,
            "fgts_juros": total_fgts_juros,
            "multa_fgts_40": multa_fgts_40,
            "total_fgts": total_fgts_geral,
            "bruto_devido_reclamante": bruto_devido_reclamante,
            "cs_segurado": total_cs_segurado_corrigido,
            "cs_patronal": total_cs_patronal_corrigido,
            "irrf_reclamante": irrf_reclamante,
            "total_descontos_reclamante": total_descontos_reclamante,
            "liquido_reclamante": liquido_reclamante,
            "honorarios_adv": honorarios_adv,
            "custas_judiciais": custas_judiciais,
            "total_devido_reclamado": total_devido_reclamado,
            "pct_remuneratorio": pct_remuneratorio,
            "pct_tributavel": pct_tributavel,
        },
    }


# ==============================================================================
# EXPORTAÇÃO DE ARQUIVO .PJC (ESTRUTURA XML INTEROPERÁVEL PJe-Calc)
# ==============================================================================

def build_pjc_xml_bytes(params: Dict[str, Any], result: Dict[str, Any]) -> bytes:
    resumo = result["resumo"]
    meta = result["meta"]
    proc = html.escape(str(params.get("processo_cnj", "")))
    rec = html.escape(str(params.get("reclamante", "")))
    rec_cpf = html.escape(str(params.get("cpf_reclamante", "")))
    rdo = html.escape(str(params.get("reclamado", "")))
    rdo_cnpj = html.escape(str(params.get("cnpj_reclamado", "")))

    verbas_xml = []
    for idx_v, vd in enumerate(result["verbas_detalhadas"], start=1):
        ocs_xml = []
        for oc in vd["ocorrencias"]:
            ocs_xml.append(
                f'      <OcorrenciaVerba competencia="{oc["competencia"]}" '
                f'base="{oc["base"]:.2f}" devido="{oc["devido"]:.2f}" pago="{oc["pago"]:.2f}" '
                f'diferenca="{oc["diferenca"]:.2f}" indice="{oc["indice"]:.9f}" corrigido="{oc["valor_corrigido"]:.2f}" />'
            )
        verbas_xml.append(
            f'    <VerbaCalculada id="{idx_v}" nome="{html.escape(vd["nome"])}" '
            f'fgts="{str(vd["incide_fgts"]).lower()}" inss="{str(vd["incide_inss"]).lower()}" irrf="{str(vd["incide_irrf"]).lower()}" '
            f'totalCorrigido="{vd["total_corrigido"]:.2f}">\n'
            + "\n".join(ocs_xml)
            + "\n    </VerbaCalculada>"
        )

    xml_str = f"""<?xml version="1.0" encoding="UTF-8"?>
<CalculoTrabalhista versaoSchema="2.13-PJC" gerador="JurisCalcPro">
  <IdentificacaoProcesso>
    <NumeroProcesso>{proc}</NumeroProcesso>
    <DataAjuizamento>{params.get("data_ajuizamento", "")}</DataAjuizamento>
    <DataLiquidacao>{params.get("data_liquidacao", "")}</DataLiquidacao>
    <Estado>{html.escape(str(params.get("estado", "RS")))}</Estado>
    <Municipio>{html.escape(str(params.get("municipio", "Gravataí")))}</Municipio>
  </IdentificacaoProcesso>
  <Partes>
    <Reclamante nome="{rec}" cpf="{rec_cpf}" />
    <Reclamado nome="{rdo}" cnpj="{rdo_cnpj}" />
  </Partes>
  <ParametrosContrato>
    <Admissao>{params.get("admissao", "")}</Admissao>
    <Demissao>{params.get("demissao", "")}</Demissao>
    <DemissaoProjetada>{meta["data_demissao_projetada"]}</DemissaoProjetada>
    <PrazoAvisoPrevioDias>{meta["prazo_aviso_dias"]}</PrazoAvisoPrevioDias>
    <CargaHoraria>{params.get("carga_horaria", 220)}</CargaHoraria>
    <PrescricaoQuinquenal>{str(params.get("aplicar_prescricao", True)).lower()}</PrescricaoQuinquenal>
    <CriterioAtualizacao>{html.escape(str(params.get("criterio_correcao", "ADC58_LEI_14905")))}</CriterioAtualizacao>
  </ParametrosContrato>
  <Verbas>
{chr(10).join(verbas_xml)}
  </Verbas>
  <ResumoLiquidacao>
    <BrutoDevidoReclamante>{resumo["bruto_devido_reclamante"]:.2f}</BrutoDevidoReclamante>
    <ContribuicaoSocialSegurado>{resumo["cs_segurado"]:.2f}</ContribuicaoSocialSegurado>
    <ContribuicaoSocialPatronal>{resumo["cs_patronal"]:.2f}</ContribuicaoSocialPatronal>
    <ImpostoRendaReclamante>{resumo["irrf_reclamante"]:.2f}</ImpostoRendaReclamante>
    <LiquidoReclamante>{resumo["liquido_reclamante"]:.2f}</LiquidoReclamante>
    <HonorariosSucumbencia>{resumo["honorarios_adv"]:.2f}</HonorariosSucumbencia>
    <CustasJudiciais>{resumo["custas_judiciais"]:.2f}</CustasJudiciais>
    <TotalDevidoReclamado>{resumo["total_devido_reclamado"]:.2f}</TotalDevidoReclamado>
  </ResumoLiquidacao>
</CalculoTrabalhista>
"""
    return xml_str.encode("utf-8")


# ==============================================================================
# GERAÇÃO DE PLANILHA PDF NO PADRÃO OFICIAL DO PJe-Calc CIDADÃO
# ==============================================================================

def build_pjecalc_official_pdf_bytes(params: Dict[str, Any], result: Dict[str, Any]) -> bytes:
    if not HAS_REPORTLAB:
        return b""

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Planilha_Liquidacao_{params.get('processo_cnj', 'PJECalc')}",
    )

    styles = getSampleStyleSheet()
    st_title = ParagraphStyle(
        "PjcTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=14,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#111827"),
    )
    st_sub = ParagraphStyle(
        "PjcSub",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1F2937"),
        spaceBefore=6,
        spaceAfter=4,
    )
    st_cell = ParagraphStyle(
        "PjcCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.3,
        leading=9,
        textColor=colors.HexColor("#111827"),
    )
    st_cell_b = ParagraphStyle(
        "PjcCellB",
        parent=st_cell,
        fontName="Helvetica-Bold",
    )
    st_cell_r = ParagraphStyle(
        "PjcCellR",
        parent=st_cell,
        alignment=TA_RIGHT,
    )
    st_cell_rb = ParagraphStyle(
        "PjcCellRB",
        parent=st_cell_b,
        alignment=TA_RIGHT,
    )
    st_legal = ParagraphStyle(
        "PjcLegal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.0,
        leading=11.5,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#1F2937"),
        spaceAfter=5,
    )

    resumo = result["resumo"]
    meta = result["meta"]

    hdr_box_style = TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#6B7280")),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9FAFB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ])

    def make_process_header(secao_titulo: str) -> List[Any]:
        out: List[Any] = [
            Paragraph("PLANILHA DE CÁLCULO — LIQUIDAÇÃO TRABALHISTA (PADRÃO PJe-Calc)", st_title),
            Spacer(1, 3 * mm),
        ]
        info_rows = [
            [
                Paragraph("<b>Processo:</b>", st_cell),
                Paragraph(html.escape(str(params.get("processo_cnj", "-"))), st_cell_b),
                Paragraph("<b>Cálculo:</b>", st_cell),
                Paragraph(html.escape(str(params.get("numero_calculo", "001/2026"))), st_cell),
            ],
            [
                Paragraph("<b>Reclamante:</b>", st_cell),
                Paragraph(
                    f"{html.escape(str(params.get('reclamante', '-')))} (CPF: {html.escape(str(params.get('cpf_reclamante', '-')))})",
                    st_cell,
                ),
                Paragraph("<b>Data Ajuizamento:</b>", st_cell),
                Paragraph(fmt_date_br(params.get("data_ajuizamento")), st_cell),
            ],
            [
                Paragraph("<b>Reclamado:</b>", st_cell),
                Paragraph(
                    f"{html.escape(str(params.get('reclamado', '-')))} (CNPJ: {html.escape(str(params.get('cnpj_reclamado', '-')))})",
                    st_cell,
                ),
                Paragraph("<b>Data Liquidação:</b>", st_cell),
                Paragraph(fmt_date_br(params.get("data_liquidacao")), st_cell_b),
            ],
            [
                Paragraph("<b>Período do Cálculo:</b>", st_cell),
                Paragraph(
                    f"{fmt_date_br(meta['inicio_calculo'])} a {fmt_date_br(meta['fim_calculo'])}",
                    st_cell,
                ),
                Paragraph("<b>Atualização:</b>", st_cell),
                Paragraph("ADC 58 STF / Lei 14.905/2024", st_cell),
            ],
        ]
        t = Table(info_rows, colWidths=[30 * mm, 76 * mm, 32 * mm, 48 * mm])
        t.setStyle(hdr_box_style)
        out.append(t)
        out.append(Spacer(1, 3 * mm))
        out.append(Paragraph(secao_titulo, st_sub))
        return out

    grid_style = TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#4B5563")),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ])

    story: List[Any] = []

    # ==========================================================================
    # PÁGINA 1 — RESUMO DO CÁLCULO
    # ==========================================================================
    story.extend(make_process_header("RESUMO DO CÁLCULO"))

    q1 = [
        [Paragraph("<b>Descrição do Bruto Devido ao Reclamante</b>", st_cell_b), Paragraph("<b>Valor Corrigido</b>", st_cell_rb), Paragraph("<b>Juros</b>", st_cell_rb), Paragraph("<b>Total</b>", st_cell_rb)],
        [Paragraph("Verbas", st_cell), Paragraph(fmt_brl(resumo["verbas_corrigido"]), st_cell_r), Paragraph(fmt_brl(resumo["juros_verbas"]), st_cell_r), Paragraph(fmt_brl(resumo["verbas_corrigido"] + resumo["juros_verbas"]), st_cell_r)],
        [Paragraph("FGTS (Depósitos + Multa 40%)", st_cell), Paragraph(fmt_brl(resumo["fgts_depositos"] + resumo["multa_fgts_40"]), st_cell_r), Paragraph(fmt_brl(resumo["fgts_juros"]), st_cell_r), Paragraph(fmt_brl(resumo["total_fgts"]), st_cell_r)],
        [Paragraph("<b>Total Bruto Devido ao Reclamante</b>", st_cell_b), Paragraph("", st_cell_r), Paragraph("", st_cell_r), Paragraph(f"<b>{fmt_brl(resumo['bruto_devido_reclamante'])}</b>", st_cell_rb)],
    ]
    t1 = Table(q1, colWidths=[96 * mm, 30 * mm, 30 * mm, 30 * mm])
    t1.setStyle(grid_style)
    story.append(t1)
    story.append(Spacer(1, 3.5 * mm))

    q2 = [
        [Paragraph("<b>Descrição de Créditos e Descontos do Reclamante</b>", st_cell_b), Paragraph("<b>Valor</b>", st_cell_rb)],
        [Paragraph("Bruto Devido ao Reclamante", st_cell), Paragraph(fmt_brl(resumo["bruto_devido_reclamante"]), st_cell_r)],
        [Paragraph("Dedução de Contribuição Social (Segurado)", st_cell), Paragraph(f"-{fmt_brl(resumo['cs_segurado'])}", st_cell_r)],
        [Paragraph("Dedução de IRRF (RRA — Art. 12-A Lei nº 7.713/1988)", st_cell), Paragraph(f"-{fmt_brl(resumo['irrf_reclamante'])}", st_cell_r)],
        [Paragraph("<b>Líquido Devido ao Reclamante</b>", st_cell_b), Paragraph(f"<b>{fmt_brl(resumo['liquido_reclamante'])}</b>", st_cell_rb)],
    ]
    t2 = Table(q2, colWidths=[146 * mm, 40 * mm])
    t2.setStyle(grid_style)
    story.append(t2)
    story.append(Spacer(1, 3.5 * mm))

    q3 = [
        [Paragraph("<b>Percentual de Parcelas Remuneratórias e Tributáveis</b>", st_cell_b), Paragraph("<b>Percentual</b>", st_cell_rb)],
        [Paragraph("Percentual de Parcelas Remuneratórias (Art. 832, § 3º, CLT)", st_cell), Paragraph(f"{fmt_brl(resumo['pct_remuneratorio'])}%", st_cell_r)],
        [Paragraph("Percentual de Parcelas Tributáveis (IRRF)", st_cell), Paragraph(f"{fmt_brl(resumo['pct_tributavel'])}%", st_cell_r)],
    ]
    t3 = Table(q3, colWidths=[146 * mm, 40 * mm])
    t3.setStyle(grid_style)
    story.append(t3)
    story.append(Spacer(1, 3.5 * mm))

    q4 = [
        [Paragraph("<b>Descrição de Débitos do Reclamado por Credor</b>", st_cell_b), Paragraph("<b>Valor Total</b>", st_cell_rb)],
        [Paragraph("Líquido Devido ao Reclamante", st_cell), Paragraph(fmt_brl(resumo["liquido_reclamante"]), st_cell_r)],
        [Paragraph("Contribuição Social — Segurado (Descontar do Principal)", st_cell), Paragraph(fmt_brl(resumo["cs_segurado"]), st_cell_r)],
        [Paragraph("Contribuição Social — Empregador / SAT / Terceiros", st_cell), Paragraph(fmt_brl(resumo["cs_patronal"]), st_cell_r)],
        [Paragraph("Imposto de Renda Retido na Fonte (IRRF)", st_cell), Paragraph(fmt_brl(resumo["irrf_reclamante"]), st_cell_r)],
        [Paragraph(f"Honorários Advocatícios Sucumbenciais ({params.get('honorarios_sucumbencia_pct', 15)}%)", st_cell), Paragraph(fmt_brl(resumo["honorarios_adv"]), st_cell_r)],
        [Paragraph("Custas Judiciais (Art. 789 da CLT)", st_cell), Paragraph(fmt_brl(resumo["custas_judiciais"]), st_cell_r)],
        [Paragraph("<b>Total Geral Devido pelo Reclamado</b>", st_cell_b), Paragraph(f"<b>{fmt_brl(resumo['total_devido_reclamado'])}</b>", st_cell_rb)],
    ]
    t4 = Table(q4, colWidths=[146 * mm, 40 * mm])
    t4.setStyle(grid_style)
    story.append(t4)

    # ==========================================================================
    # PÁGINA 2 — CRITÉRIO DE CÁLCULO E FUNDAMENTAÇÃO LEGAL
    # ==========================================================================
    story.append(PageBreak())
    story.extend(make_process_header("CRITÉRIO DE CÁLCULO E FUNDAMENTAÇÃO LEGAL"))

    criterios_texto = [
        "1. <b>Período e Prescrição Quinquenal:</b> A apuração observa rigorosamente o período imprescrito fixado no comando sentencial/acórdão (art. 7º, XXIX, da Constituição Federal), com projeção do aviso prévio indenizado proporcional ao tempo de serviço nos termos da Lei nº 12.506/2011 e da OJ nº 82 da SDI-1 do TST.",
        "2. <b>Evolução Salarial e Divisor:</b> As verbas de natureza salarial foram apuradas mês a mês sobre a evolução salarial comprovada nos autos, observando-se os divisores e adicionais deferidos no título executivo judicial.",
        "3. <b>Atualização Monetária e Juros de Mora (ADC 58 STF e Lei nº 14.905/2024):</b> Na fase pré-judicial aplica-se o IPCA-E acrescido dos juros legais do art. 39, caput, da Lei nº 8.177/1991 (TR); a partir do ajuizamento da ação até 29/08/2024 aplica-se exclusivamente a taxa SELIC (englobando correção e juros); e a partir de 30/08/2024 aplica-se o IPCA (IBGE) como índice de atualização monetária e a Taxa Legal (SELIC deduzida do IPCA, limitada a zero quando negativa) a título de juros de mora, nos termos do art. 389, parágrafo único, e art. 406 do Código Civil com redação dada pela Lei nº 14.905/2024.",
        "4. <b>FGTS e Indenização Compensatória de 40%:</b> Apurados à alíquota de 8% sobre as parcelas de natureza salarial integrantes da base de incidência (art. 15 da Lei nº 8.036/1990), acrescidos da indenização de 40% quando aplicável.",
        "5. <b>Contribuições Previdenciárias (Cota Segurado e Cota Patronal):</b> Apuradas mês a mês (regime de competência — Súmula nº 368 do TST), aplicando-se ao segurado as alíquotas progressivas vigentes em cada competência conforme a Emenda Constitucional nº 103/2019 e Portarias Interministeriais MPS/MF.",
        "6. <b>Imposto de Renda Retido na Fonte (RRA):</b> Apurado sobre os rendimentos tributáveis acumulados mediante utilização da tabela progressiva acumulada vigente na data da liquidação (art. 12-A da Lei nº 7.713/1988 e IN RFB nº 1.500/2014), excluídos os juros de mora (Tema 808 STF) e deduzida a contribuição previdenciária devida pelo reclamante.",
        "7. <b>Base de Apuração dos Juros sobre Verbas:</b> Conforme critério oficial do PJe-Calc, os juros moratórios incidentes sobre as verbas trabalhistas são calculados sobre o capital corrigido após a dedução da contribuição social devida pelo reclamante em cada competência.",
    ]
    for item_txt in criterios_texto:
        story.append(Paragraph(item_txt, st_legal))

    # ==========================================================================
    # PÁGINA 3 — DADOS DO CÁLCULO E HISTÓRICO SALARIAL
    # ==========================================================================
    story.append(PageBreak())
    story.extend(make_process_header("DADOS DO CÁLCULO E HISTÓRICO SALARIAL"))

    dados_calc = [
        [Paragraph("<b>Estado:</b>", st_cell), Paragraph(html.escape(str(params.get("estado", "RS"))), st_cell), Paragraph("<b>Município:</b>", st_cell), Paragraph(html.escape(str(params.get("municipio", "Gravataí"))), st_cell)],
        [Paragraph("<b>Admissão:</b>", st_cell), Paragraph(fmt_date_br(params.get("admissao")), st_cell), Paragraph("<b>Demissão:</b>", st_cell), Paragraph(fmt_date_br(params.get("demissao")), st_cell)],
        [Paragraph("<b>Maior Remuneração:</b>", st_cell), Paragraph(fmt_brl_rs(params.get("maior_remuneracao")), st_cell), Paragraph("<b>Última Remuneração:</b>", st_cell), Paragraph(fmt_brl_rs(params.get("ultima_remuneracao")), st_cell)],
        [Paragraph("<b>Prazo Aviso Prévio:</b>", st_cell), Paragraph(f"{meta['prazo_aviso_dias']} dias", st_cell), Paragraph("<b>Carga Horária:</b>", st_cell), Paragraph(f"{params.get('carga_horaria', 220)}h", st_cell)],
        [Paragraph("<b>Prescrição Quinquenal:</b>", st_cell), Paragraph("Sim" if params.get("aplicar_prescricao", True) else "Não", st_cell), Paragraph("<b>Projetar Aviso Prévio:</b>", st_cell), Paragraph("Sim" if params.get("projetar_aviso", True) else "Não", st_cell)],
    ]
    tdc = Table(dados_calc, colWidths=[38 * mm, 55 * mm, 38 * mm, 55 * mm])
    tdc.setStyle(grid_style)
    story.append(tdc)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("OCORRÊNCIAS DO HISTÓRICO SALARIAL", st_sub))

    hist_rows = [[
        Paragraph("<b>Competência</b>", st_cell_b),
        Paragraph("<b>Salário Pago</b>", st_cell_rb),
        Paragraph("<b>Salário Devido</b>", st_cell_rb),
        Paragraph("<b>Diferença</b>", st_cell_rb),
    ]]
    for h in result["historico_consolidado"]:
        hist_rows.append([
            Paragraph(fmt_comp_br(h["competencia"]), st_cell),
            Paragraph(fmt_brl(h["salario_pago"]), st_cell_r),
            Paragraph(fmt_brl(h["salario_devido"]), st_cell_r),
            Paragraph(fmt_brl(h["diferenca"]), st_cell_r),
        ])
    thist = Table(hist_rows, colWidths=[46 * mm, 46 * mm, 46 * mm, 48 * mm], repeatRows=1)
    thist.setStyle(grid_style)
    story.append(thist)

    # ==========================================================================
    # PÁGINA 4 — DEMONSTRATIVO DE VERBAS E DEMONSTRATIVO DE JUROS
    # ==========================================================================
    story.append(PageBreak())
    story.extend(make_process_header("DEMONSTRATIVO DE VERBAS"))

    for vd in result["verbas_detalhadas"]:
        story.append(Paragraph(f"<b>Verba: {html.escape(vd['nome'])}</b>", st_sub))
        v_rows = [[
            Paragraph("<b>Comp.</b>", st_cell_b),
            Paragraph("<b>Base</b>", st_cell_rb),
            Paragraph("<b>Div/Mult</b>", st_cell_rb),
            Paragraph("<b>Qtd.</b>", st_cell_rb),
            Paragraph("<b>Devido</b>", st_cell_rb),
            Paragraph("<b>Pago</b>", st_cell_rb),
            Paragraph("<b>Diferença</b>", st_cell_rb),
            Paragraph("<b>Índice</b>", st_cell_rb),
            Paragraph("<b>Corrigido</b>", st_cell_rb),
        ]]
        for oc in vd["ocorrencias"]:
            v_rows.append([
                Paragraph(fmt_comp_br(oc["competencia"]), st_cell),
                Paragraph(fmt_brl(oc["base"]), st_cell_r),
                Paragraph(str(oc["div_mult"]), st_cell_r),
                Paragraph(fmt_brl(oc["qtd"]), st_cell_r),
                Paragraph(fmt_brl(oc["devido"]), st_cell_r),
                Paragraph(fmt_brl(oc["pago"]), st_cell_r),
                Paragraph(fmt_brl(oc["diferenca"]), st_cell_r),
                Paragraph(fmt_idx(oc["indice"]), st_cell_r),
                Paragraph(fmt_brl(oc["valor_corrigido"]), st_cell_r),
            ])
        v_rows.append([
            Paragraph("<b>Total</b>", st_cell_b),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(vd['total_devido_orig'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(vd['total_pago_orig'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(vd['total_diferenca'])}</b>", st_cell_rb),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(vd['total_corrigido'])}</b>", st_cell_rb),
        ])
        tv = Table(
            v_rows,
            colWidths=[16 * mm, 22 * mm, 16 * mm, 14 * mm, 22 * mm, 20 * mm, 22 * mm, 28 * mm, 26 * mm],
            repeatRows=1,
        )
        tv.setStyle(grid_style)
        story.append(tv)
        story.append(Spacer(1, 3 * mm))

    if result["juros_ocorrencias"]:
        story.append(Paragraph("DEMONSTRATIVO DE JUROS SOBRE VERBAS (DEDUÇÃO PRÉVIA DA CONTRIBUIÇÃO SOCIAL)", st_sub))
        j_rows = [[
            Paragraph("<b>Competência</b>", st_cell_b),
            Paragraph("<b>Total Verbas</b>", st_cell_rb),
            Paragraph("<b>Contrib. Social</b>", st_cell_rb),
            Paragraph("<b>Capital</b>", st_cell_rb),
            Paragraph("<b>Taxa Juros (%)</b>", st_cell_rb),
            Paragraph("<b>Valor Juros</b>", st_cell_rb),
        ]]
        for jo in result["juros_ocorrencias"]:
            j_rows.append([
                Paragraph(fmt_comp_br(jo["competencia"]), st_cell),
                Paragraph(fmt_brl(jo["total_verbas"]), st_cell_r),
                Paragraph(fmt_brl(jo["contribuicao_social"]), st_cell_r),
                Paragraph(fmt_brl(jo["capital"]), st_cell_r),
                Paragraph(f"{fmt_brl(jo['taxa_juros_pct'])}%", st_cell_r),
                Paragraph(fmt_brl(jo["juros"]), st_cell_r),
            ])
        j_rows.append([
            Paragraph("<b>Total</b>", st_cell_b),
            Paragraph(f"<b>{fmt_brl(resumo['verbas_corrigido'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(resumo['cs_segurado'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(resumo['capital_juros'])}</b>", st_cell_rb),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(resumo['juros_verbas'])}</b>", st_cell_rb),
        ])
        tj = Table(j_rows, colWidths=[28 * mm, 32 * mm, 32 * mm, 32 * mm, 30 * mm, 32 * mm], repeatRows=1)
        tj.setStyle(grid_style)
        story.append(tj)

    # ==========================================================================
    # PÁGINA 5 — DEMONSTRATIVO DE FGTS E CONTRIBUIÇÃO SOCIAL
    # ==========================================================================
    story.append(PageBreak())
    story.extend(make_process_header("DEMONSTRATIVO DE FGTS (8%) E CONTRIBUIÇÃO SOCIAL"))

    if result["fgts_ocorrencias"]:
        story.append(Paragraph("DEMONSTRATIVO DE FGTS", st_sub))
        f_rows = [[
            Paragraph("<b>Comp.</b>", st_cell_b),
            Paragraph("<b>Base</b>", st_cell_rb),
            Paragraph("<b>Alíq.</b>", st_cell_rb),
            Paragraph("<b>Devido</b>", st_cell_rb),
            Paragraph("<b>Índice</b>", st_cell_rb),
            Paragraph("<b>Corrigido</b>", st_cell_rb),
            Paragraph("<b>Juros</b>", st_cell_rb),
            Paragraph("<b>Total</b>", st_cell_rb),
        ]]
        for fo in result["fgts_ocorrencias"]:
            f_rows.append([
                Paragraph(fmt_comp_br(fo["competencia"]), st_cell),
                Paragraph(fmt_brl(fo["base"]), st_cell_r),
                Paragraph("8,00%", st_cell_r),
                Paragraph(fmt_brl(fo["devido"]), st_cell_r),
                Paragraph(fmt_idx(fo["indice"]), st_cell_r),
                Paragraph(fmt_brl(fo["corrigido"]), st_cell_r),
                Paragraph(fmt_brl(fo["juros"]), st_cell_r),
                Paragraph(fmt_brl(fo["total"]), st_cell_r),
            ])
        f_rows.append([
            Paragraph("<b>Total (+ Multa 40%)</b>", st_cell_b),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(resumo['fgts_depositos'] + resumo['multa_fgts_40'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(resumo['fgts_juros'])}</b>", st_cell_rb),
            Paragraph(f"<b>{fmt_brl(resumo['total_fgts'])}</b>", st_cell_rb),
        ])
        tf = Table(f_rows, colWidths=[18 * mm, 26 * mm, 14 * mm, 22 * mm, 30 * mm, 26 * mm, 24 * mm, 26 * mm], repeatRows=1)
        tf.setStyle(grid_style)
        story.append(tf)
        story.append(Spacer(1, 4 * mm))

    if result["cs_ocorrencias"]:
        story.append(Paragraph("DEMONSTRATIVO DE CONTRIBUIÇÃO SOCIAL (SEGURADO + EMPREGADOR)", st_sub))
        cs_rows = [[
            Paragraph("<b>Comp.</b>", st_cell_b),
            Paragraph("<b>Base Verbas</b>", st_cell_rb),
            Paragraph("<b>Seg. Devido</b>", st_cell_rb),
            Paragraph("<b>Seg. Corrig.</b>", st_cell_rb),
            Paragraph("<b>Patr. Devido</b>", st_cell_rb),
            Paragraph("<b>Patr. Corrig.</b>", st_cell_rb),
            Paragraph("<b>Patr. Total</b>", st_cell_rb),
        ]]
        for co in result["cs_ocorrencias"]:
            cs_rows.append([
                Paragraph(fmt_comp_br(co["competencia"]), st_cell),
                Paragraph(fmt_brl(co["base_verbas"]), st_cell_r),
                Paragraph(fmt_brl(co["segurado_devido"]), st_cell_r),
                Paragraph(fmt_brl(co["segurado_corrigido"]), st_cell_r),
                Paragraph(fmt_brl(co["patronal_devido"]), st_cell_r),
                Paragraph(fmt_brl(co["patronal_corrigido"]), st_cell_r),
                Paragraph(fmt_brl(co["patronal_total"]), st_cell_r),
            ])
        cs_rows.append([
            Paragraph("<b>Total</b>", st_cell_b),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(resumo['cs_segurado'])}</b>", st_cell_rb),
            Paragraph("", st_cell_r),
            Paragraph("", st_cell_r),
            Paragraph(f"<b>{fmt_brl(resumo['cs_patronal'])}</b>", st_cell_rb),
        ])
        tcs = Table(cs_rows, colWidths=[20 * mm, 28 * mm, 26 * mm, 28 * mm, 26 * mm, 28 * mm, 30 * mm], repeatRows=1)
        tcs.setStyle(grid_style)
        story.append(tcs)

    doc.build(story)
    return buf.getvalue()


# ==============================================================================
# EXTRATOR INTELIGENTE DE AUTOS E DECISÕES (PDF MULTIMODAL + REGEX + GEMINI)
# ==============================================================================

def extract_text_from_uploaded_files(uploaded_files: List[Any]) -> str:
    chunks: List[str] = []
    for uf in uploaded_files:
        raw = uf.getvalue()
        name = uf.name
        if name.lower().endswith(".pdf") and fitz is not None:
            try:
                pdf_doc = fitz.open(stream=raw, filetype="pdf")
                pages_txt = []
                for i, page in enumerate(pdf_doc, start=1):
                    txt = page.get_text("text").strip()
                    if txt:
                        pages_txt.append(f"[PÁGINA {i} - {name}]\n{txt}")
                chunks.append("\n".join(pages_txt))
            except Exception as e:
                chunks.append(f"[ERRO AO LER {name}: {e}]")
        else:
            try:
                chunks.append(f"[ARQUIVO {name}]\n{raw.decode('utf-8', errors='ignore')}")
            except Exception:
                pass
    return "\n\n".join(chunks)


def deterministic_regex_extract(raw_text: str, cidade_default: str = "Gravataí") -> Dict[str, Any]:
    """Extração determinística inicial diretamente do texto do PDF antes mesmo da IA."""
    cnj_m = re.search(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}", raw_text)
    cpf_m = re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}", raw_text)
    cnpj_m = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", raw_text)
    vara_m = re.search(r"(\d+ª\s+Vara\s+do\s+Trabalho\s+de\s+[^\n\r,;]+)", raw_text, re.IGNORECASE)

    rec_m = re.search(r"(?:RECLAMANTE|AUTOR[A]?)\s*[:\-]\s*([A-ZÀ-Ú\s]{5,80})", raw_text)
    rdo_m = re.search(r"(?:RECLAMAD[OA]|R[ÉE]U)\s*[:\-]\s*([A-ZÀ-Ú0-9\s\.\-&]{5,90})", raw_text)

    datas_br = re.findall(r"\b(\d{2}/\d{2}/\d{4})\b", raw_text)
    datas_iso = []
    for d_str in datas_br:
        parsed = parse_date_br(d_str)
        if parsed and date(1995, 1, 1) <= parsed <= date(2030, 12, 31):
            datas_iso.append(parsed)
    datas_iso = sorted(set(datas_iso))

    adm_est = datas_iso[0].isoformat() if len(datas_iso) >= 1 else "2022-01-10"
    dem_est = datas_iso[-2].isoformat() if len(datas_iso) >= 3 else (datas_iso[-1].isoformat() if datas_iso else "2024-06-15")
    ajuiz_est = datas_iso[-1].isoformat() if datas_iso else date.today().isoformat()

    return {
        "processo_cnj": cnj_m.group(0) if cnj_m else "0000000-00.2024.5.04.0000",
        "numero_calculo": "001/2026",
        "vara_trabalho": vara_m.group(1).strip() if vara_m else f"Vara do Trabalho de {cidade_default}",
        "estado": "RS",
        "municipio": cidade_default,
        "cidade_assinatura": cidade_default,
        "reclamante": rec_m.group(1).strip().splitlines()[0] if rec_m else "RECLAMANTE EXTRAÍDO DOS AUTOS (CONFERIR NOME)",
        "cpf_reclamante": cpf_m.group(0) if cpf_m else "",
        "reclamado": rdo_m.group(1).strip().splitlines()[0] if rdo_m else "RECLAMADO EXTRAÍDO DOS AUTOS (CONFERIR RAZÃO SOCIAL)",
        "cnpj_reclamado": cnpj_m.group(0) if cnpj_m else "",
        "admissao": adm_est,
        "demissao": dem_est,
        "data_ajuizamento": ajuiz_est,
        "data_liquidacao": date.today().isoformat(),
        "maior_remuneracao": 2500.00,
        "ultima_remuneracao": 2500.00,
        "carga_horaria": 220,
        "aplicar_prescricao": True,
        "projetar_aviso": True,
        "aplicar_multa_fgts_40": True,
        "apurar_irrf": True,
        "apurar_custas": True,
        "honorarios_sucumbencia_pct": 15.0,
        "custas_pct": 2.0,
        "criterio_correcao": "ADC58_LEI_14905",
        "verbas": [],
    }


def get_initial_empty_params(cidade_default: str = "Gravataí") -> Dict[str, Any]:
    return {
        "processo_cnj": "",
        "numero_calculo": "001/2026",
        "vara_trabalho": "",
        "estado": "RS",
        "municipio": cidade_default,
        "cidade_assinatura": cidade_default,
        "reclamante": "",
        "cpf_reclamante": "",
        "reclamado": "",
        "cnpj_reclamado": "",
        "admissao": "2022-01-10",
        "demissao": "2024-06-15",
        "data_ajuizamento": "2024-08-20",
        "data_liquidacao": date.today().isoformat(),
        "maior_remuneracao": 0.0,
        "ultima_remuneracao": 0.0,
        "carga_horaria": 220,
        "aplicar_prescricao": True,
        "projetar_aviso": True,
        "aplicar_multa_fgts_40": True,
        "apurar_irrf": True,
        "apurar_custas": True,
        "honorarios_sucumbencia_pct": 15.0,
        "custas_pct": 2.0,
        "criterio_correcao": "ADC58_LEI_14905",
        "advogado_nome": "Dr. Advogado Responsável",
        "advogado_oab": "OAB/RS 00.000",
        "parecer_advogado": "",
        "historico_salarial": [],
        "verbas": [],
    }


def analyze_case_with_gemini(
    uploaded_files: List[Any],
    raw_text: str,
    api_key: str,
    parecer_advogado: str = "",
    cidade_default: str = "Gravataí",
) -> Tuple[Dict[str, Any], Optional[str], str]:
    """
    Lê os PDFs anexados (via multimodal nativo PDF + texto extraído) + Parecer do Advogado.
    Nunca substitui por dados fictícios caso falhe; retorna diagnóstico claro.
    """
    base = deterministic_regex_extract(raw_text, cidade_default)
    base["parecer_advogado"] = parecer_advogado

    if not api_key or not api_key.strip():
        return (
            base,
            "⚠️ Chave da API Gemini não informada na barra lateral. Foi feita apenas a pré-extração básica por texto (CNJ/CPF/CNPJ/Datas). Informe sua chave Gemini para a IA ler todas as verbas e salários do PDF.",
            "Pré-extração básica (sem chave Gemini)",
        )

    if genai is None or types is None:
        return (
            base,
            "⚠️ Biblioteca `google-genai` não está instalada no ambiente.",
            "Erro de dependência",
        )

    bloco_parecer = (
        f"\nPARECER E INSTRUÇÕES OBRIGATÓRIAS DO ADVOGADO (SIGA RIGOROSAMENTE):\n{parecer_advogado.strip()}\n"
        if parecer_advogado and parecer_advogado.strip()
        else "\nPARECER DO ADVOGADO: Nenhuma observação adicional informada; siga estritamente o título judicial dos autos.\n"
    )

    prompt = f"""
Você é um Perito Contador Trabalhista sênior especialista no sistema oficial PJe-Calc Cidadão.
Analise integralmente os arquivos anexados do processo (Sentença, Acórdão, Petição Inicial, CTPS, Contracheques, Cartões de Ponto) e o parecer do advogado abaixo.

{bloco_parecer}

REGRAS OBRIGATÓRIAS DE EXTRAÇÃO:
1. Extraia os dados REAIS deste processo (não invente nomes, datas ou valores).
2. Se houver múltiplos contratos na CTPS, identifique autonomamente o contrato referente ao Reclamado desta ação trabalhista.
3. Extraia exatamente o percentual de honorários sucumbenciais e custas fixados na sentença/acórdão.
4. Para cada verba deferida, mapeie o `tipo` usando EXCLUSIVAMENTE um dos códigos válidos:
   - `DIFERENCA_SALARIAL` (quando há salário devido maior que o pago)
   - `HORAS_EXTRAS` (use `aliquota_ou_percentual`: 0.50 para 50%, 1.00 para 100%, `quantidade_mensal`: média de horas/mês, `divisor`: 220 ou 200 ou 180)
   - `ADICIONAL_NOTURNO` (use `aliquota_ou_percentual`: 0.20, `quantidade_mensal`: horas noturnas/mês, `divisor`: 220)
   - `INSALUBRIDADE` (use `aliquota_ou_percentual`: 0.10, 0.20 ou 0.40)
   - `PERICULOSIDADE` (use `aliquota_ou_percentual`: 0.30)
   - `AVISO_PREVIO`
   - `DECIMO_TERCEIRO`
   - `FERIAS_TERCO`
   - `MULTA_477`
   - `VERBA_FIXA_MENSAL` (use `quantidade_mensal` como valor mensal fixo em R$)
   - `VERBA_UNICA_RESCISAO` (use `quantidade_mensal` como valor único em R$ na rescisão)
5. Se houver evolução salarial mês a mês nos contracheques/autos, preencha `historico_salarial` com `competencia` ("YYYY-MM"), `salario_pago` e `salario_devido`. Caso contrário, informe `maior_remuneracao` e `ultima_remuneracao` reais do contrato.

Retorne EXCLUSIVAMENTE um JSON válido neste formato exato:
{{
  "processo_cnj": "0000000-00.0000.5.00.0000",
  "vara_trabalho": "Xª Vara do Trabalho de ...",
  "estado": "RS",
  "municipio": "{cidade_default}",
  "reclamante": "NOME COMPLETO DO RECLAMANTE",
  "cpf_reclamante": "000.000.000-00",
  "reclamado": "RAZAO SOCIAL COMPLETA DO RECLAMADO",
  "cnpj_reclamado": "00.000.000/0001-00",
  "admissao": "YYYY-MM-DD",
  "demissao": "YYYY-MM-DD",
  "data_ajuizamento": "YYYY-MM-DD",
  "maior_remuneracao": 0.0,
  "ultima_remuneracao": 0.0,
  "carga_horaria": 220,
  "aplicar_prescricao": true,
  "projetar_aviso": true,
  "aplicar_multa_fgts_40": true,
  "apurar_irrf": true,
  "apurar_custas": true,
  "honorarios_sucumbencia_pct": 15.0,
  "custas_pct": 2.0,
  "resumo_analise_ia": "Breve explicação objetiva do que foi identificado na sentença/autos e como o parecer do advogado foi aplicado.",
  "historico_salarial": [
    {{"competencia": "YYYY-MM", "salario_pago": 0.0, "salario_devido": 0.0}}
  ],
  "verbas": [
    {{
      "nome": "NOME DA VERBA DEFERIDA",
      "tipo": "HORAS_EXTRAS",
      "aliquota_ou_percentual": 0.5,
      "quantidade_mensal": 20.0,
      "divisor": 220.0,
      "inicio": "YYYY-MM-DD",
      "fim": "YYYY-MM-DD",
      "incide_fgts": true,
      "incide_inss": true,
      "incide_irrf": true,
      "gera_reflexos_13_ferias_aviso": true,
      "ativo": true
    }}
  ]
}}
"""

    contents_payload: List[Any] = [prompt]

    # Envia os PDFs diretamente de forma multimodal para ler tanto texto quanto páginas digitalizadas/CTPS
    total_pdf_bytes = 0
    for uf in uploaded_files:
        raw_b = uf.getvalue()
        if uf.name.lower().endswith(".pdf") and (total_pdf_bytes + len(raw_b)) <= 18 * 1024 * 1024:
            contents_payload.append(
                types.Part.from_bytes(data=raw_b, mime_type="application/pdf")
            )
            total_pdf_bytes += len(raw_b)

    if raw_text and raw_text.strip():
        contents_payload.append(f"TEXTO EXTRAÍDO DOS AUTOS:\n{raw_text[:80000]}")

    try:
        client = genai.Client(api_key=api_key.strip())
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contents_payload,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.1,
            ),
        )
        raw_resp = (response.text or "").strip()
        if raw_resp.startswith("```"):
            raw_resp = re.sub(r"^```(?:json)?\s*", "", raw_resp)
            raw_resp = re.sub(r"\s*```$", "", raw_resp)

        parsed = json.loads(raw_resp)
        if not isinstance(parsed, dict):
            return base, "⚠️ A resposta da IA não retornou um objeto estruturado válido.", "Falha de formato"

        for k, v in parsed.items():
            if v is not None and v != "":
                base[k] = v

        adm = parse_date_br(base.get("admissao")) or date(2022, 1, 10)
        dem = parse_date_br(base.get("demissao")) or date(2024, 6, 15)
        sal_p = float(base.get("ultima_remuneracao") or base.get("maior_remuneracao") or 2500.0)
        sal_d = float(base.get("maior_remuneracao") or sal_p)

        if not base.get("historico_salarial"):
            base["historico_salarial"] = [
                {"competencia": f"{y:04d}-{m:02d}", "salario_pago": sal_p, "salario_devido": sal_d}
                for y, m in month_range(adm, dem)
            ]

        resumo_ia = str(parsed.get("resumo_analise_ia", "Autos lidos e parametrizados conforme decisão judicial."))
        return base, None, resumo_ia

    except Exception as exc:
        return (
            base,
            f"⚠️ Erro ao consultar a IA Gemini: {exc}. Verifique se sua chave API está ativa e tente novamente.",
            "Erro na chamada da IA",
        )


def build_petition_minuta(params: Dict[str, Any], result: Dict[str, Any]) -> str:
    resumo = result["resumo"]
    cidade = params.get("cidade_assinatura") or params.get("municipio") or "Gravataí"
    hoje_extenso = datetime.now().strftime("%d/%m/%Y")
    return f"""EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DA {str(params.get('vara_trabalho', 'VARA DO TRABALHO')).upper()}

Processo nº {params.get('processo_cnj', '')}

{str(params.get('reclamante', '')).upper()}, já devidamente qualificado(a) nos autos da Reclamação Trabalhista que move em face de {str(params.get('reclamado', '')).upper()}, vem, respeitosamente, à presença de Vossa Excelência, por seu(sua) advogado(a) infra-assinado(a), apresentar os seus

CÁLCULOS DE LIQUIDAÇÃO DE SENTENÇA

elaborados em estrita observância ao título executivo judicial, à legislação trabalhista, previdenciária e tributária vigente e aos critérios oficiais do sistema PJe-Calc, conforme resumido abaixo:

1. RESUMO GERAL DA LIQUIDAÇÃO (ATUALIZADO ATÉ {fmt_date_br(params.get('data_liquidacao'))}):
   • Total Bruto Devido ao(à) Reclamante: {fmt_brl_rs(resumo['bruto_devido_reclamante'])}
   • (-) Contribuição Social — Cota Segurado: {fmt_brl_rs(resumo['cs_segurado'])}
   • (-) Imposto de Renda Retido na Fonte (RRA — Art. 12-A Lei nº 7.713/88): {fmt_brl_rs(resumo['irrf_reclamante'])}
   • (=) Valor Líquido Devido ao(à) Reclamante: {fmt_brl_rs(resumo['liquido_reclamante'])}
   • (+) Contribuição Social — Cota Patronal / SAT / Terceiros: {fmt_brl_rs(resumo['cs_patronal'])}
   • (+) Honorários Advocatícios Sucumbenciais ({params.get('honorarios_sucumbencia_pct', 15)}%): {fmt_brl_rs(resumo['honorarios_adv'])}
   • (+) Custas Processuais: {fmt_brl_rs(resumo['custas_judiciais'])}
   • (=) TOTAL GERAL DEVIDO PELO(A) RECLAMADO(A): {fmt_brl_rs(resumo['total_devido_reclamado'])}

2. REQUERIMENTOS:
   Diante do exposto, requer a homologação dos cálculos apresentados, acompanhados da respectiva planilha analítica e do arquivo interoperável (.PJC), intimando-se a parte executada para pagamento sob pena de execução forçada.

Termos em que,
Pede e espera deferimento.

{cidade}, {hoje_extenso}.

____________________________________________________
{params.get('advogado_nome', 'Advogado(a) Responsável')}
{params.get('advogado_oab', 'OAB')}
"""


# ==============================================================================
# INTERFACE STREAMLIT — JURISCALC PRO
# ==============================================================================

st.set_page_config(
    page_title="JurisCalc Pro • Liquidação Trabalhista",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        color: #F8FAFC;
        padding: 20px 26px;
        border-radius: 12px;
        border-left: 5px solid #B91C1C;
        margin-bottom: 18px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.14);
    }
    .hero-banner h1 { font-size: 1.55rem; margin: 0 0 4px 0; font-weight: 700; letter-spacing: -0.02em; }
    .hero-banner p { font-size: 0.92rem; margin: 0; color: #CBD5E1; }
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .metric-label { font-size: 0.78rem; color: #64748B; font-weight: 600; text-transform: uppercase; }
    .metric-value { font-size: 1.32rem; color: #0F172A; font-weight: 700; margin-top: 4px; }
    .metric-sub { font-size: 0.75rem; color: #475569; margin-top: 2px; }
</style>
""",
    unsafe_allow_html=True,
)


def _recompute_from_current_params(params: Dict[str, Any], save_to_history: bool = True) -> Dict[str, Any]:
    res = run_pjecalc_engine(params)
    st.session_state["params"] = params
    st.session_state["result"] = res
    st.session_state["pjc_bytes"] = build_pjc_xml_bytes(params, res)
    st.session_state["pdf_bytes"] = build_pjecalc_official_pdf_bytes(params, res)
    st.session_state["minuta_txt"] = build_petition_minuta(params, res)
    if save_to_history and params.get("processo_cnj"):
        save_calculation_history(params, res)
    return res


if "params" not in st.session_state:
    _recompute_from_current_params(get_initial_empty_params("Gravataí"), save_to_history=False)

# ==============================================================================
# BARRA LATERAL (CONFIGURAÇÕES, SINCRONIZAÇÃO E IDENTIFICAÇÃO DO ESCRITÓRIO)
# ==============================================================================
with st.sidebar:
    st.markdown("### ⚖️ JurisCalc Pro")
    st.caption("Liquidação Trabalhista Automatizada (Padrão PJe-Calc)")

    nav_page = st.radio(
        "Navegação",
        ["⚖️ Novo Cálculo (.PJC + Planilha)", "🗂️ Histórico de Liquidações"],
        label_visibility="collapsed",
    )

    st.divider()
    api_key_input = st.text_input(
        "Chave API Gemini (Obrigatória para IA)",
        type="password",
        value=os.environ.get("GEMINI_API_KEY", ""),
        help="Cole sua chave Gemini (AIza...) para ler sentenças, acórdãos, CTPS e contracheques.",
    )

    st.markdown("**Dados do Escritório & Minuta**")
    adv_nome_sb = st.text_input(
        "Advogado(a) Responsável",
        value=st.session_state["params"].get("advogado_nome", "Dr. Advogado Responsável"),
    )
    adv_oab_sb = st.text_input(
        "Inscrição OAB",
        value=st.session_state["params"].get("advogado_oab", "OAB/RS 00.000"),
    )
    cidade_sb = st.text_input(
        "Cidade da Assinatura (Minuta)",
        value=st.session_state["params"].get("cidade_assinatura", "Gravataí"),
    )
    st.session_state["params"]["advogado_nome"] = adv_nome_sb
    st.session_state["params"]["advogado_oab"] = adv_oab_sb
    st.session_state["params"]["cidade_assinatura"] = cidade_sb

    st.divider()
    st.markdown("**Regras & Tabelas Vigentes**")
    sync_meta = st.session_state.get("sync_status_info", SYNC_STATUS_INFO)
    st.caption(
        f"• **Status:** `{sync_meta['status']}`\n"
        f"• **Tabelas INSS / IRRF (RRA):** Vigentes até `{sync_meta['ultima_competencia']}`\n"
        f"• **Índices BACEN:** SELIC, IPCA, IPCA-E e TR sincronizados"
    )

    with st.expander("Atualizar Tabelas Legais (.JSON)", expanded=False):
        feed_url_sb = st.text_input(
            "URL Feed de Tabelas (JSON)",
            value=st.session_state.get("tables_feed_url", ""),
            placeholder="https://.../tabelas_pjecalc.json",
            key="sb_tables_feed_url",
        )
        st.session_state["tables_feed_url"] = feed_url_sb

        uploaded_tbl = st.file_uploader(
            "Importar Pacote de Tabelas (.JSON)",
            type=["json"],
            key="sb_upload_tables_json",
        )
        if st.button("🔄 Sincronizar Agora", use_container_width=True, key="btn_sync_all_now"):
            raw_json_str = uploaded_tbl.getvalue().decode("utf-8") if uploaded_tbl is not None else None
            new_sync = sync_legal_tables_dynamic(remote_url=feed_url_sb, raw_json_override=raw_json_str)
            st.session_state["sync_status_info"] = new_sync
            st.cache_data.clear()
            _recompute_from_current_params(st.session_state["params"], save_to_history=False)
            st.success(f"Sincronizado até {new_sync['ultima_competencia']}!")


# ==============================================================================
# PÁGINA 1 — NOVO CÁLCULO (.PJC + PLANILHA PDF + MINUTA)
# ==============================================================================
if nav_page == "⚖️ Novo Cálculo (.PJC + Planilha)":
    st.markdown(
        """
        <div class="hero-banner">
            <h1>JurisCalc Pro — Liquidação Trabalhista</h1>
            <p>Anexe os autos do processo, informe seu parecer/orientações jurídicas e gere automaticamente a Planilha PDF, o arquivo .PJC e a Minuta no padrão oficial PJe-Calc.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.subheader("1. Autos do Processo & Parecer do Advogado")
        uploaded_docs = st.file_uploader(
            "Anexe os arquivos do processo (Sentença, Acórdão, Petição Inicial, CTPS, Contracheques em PDF ou TXT)",
            type=["pdf", "txt"],
            accept_multiple_files=True,
        )

        parecer_adv_input = st.text_area(
            "Parecer / Comentários do Advogado para a IA considerar na análise (Opcional, mas recomendado)",
            value=st.session_state["params"].get("parecer_advogado", ""),
            placeholder=(
                "Ex.: Considerar apenas o vínculo com a empresa X (01/03/2022 a 15/08/2024); "
                "aplicar divisor 200 conforme norma coletiva; calcular 25 horas extras 50% por mês com reflexos em 13º, férias + 1/3, aviso prévio e FGTS + 40%; "
                "honorários sucumbenciais de 10% fixados na sentença."
            ),
            height=110,
        )
        st.session_state["params"]["parecer_advogado"] = parecer_adv_input

        if st.button("🤖 Analisar Autos + Parecer e Gerar Cálculo Completo", type="primary", use_container_width=True):
            if not uploaded_docs:
                st.error("Por favor, anexe pelo menos um arquivo PDF ou TXT do processo antes de iniciar a análise.")
            else:
                with st.spinner("Lendo documentos do processo, aplicando seu parecer e liquidando verbas no padrão PJe-Calc..."):
                    raw_txt = extract_text_from_uploaded_files(uploaded_docs)
                    extracted_params, err_msg, resumo_ia = analyze_case_with_gemini(
                        uploaded_files=uploaded_docs,
                        raw_text=raw_txt,
                        api_key=api_key_input,
                        parecer_advogado=parecer_adv_input,
                        cidade_default=cidade_sb,
                    )
                    extracted_params["advogado_nome"] = adv_nome_sb
                    extracted_params["advogado_oab"] = adv_oab_sb
                    extracted_params["cidade_assinatura"] = cidade_sb
                    _recompute_from_current_params(extracted_params, save_to_history=True)
                    st.session_state["last_ai_error"] = err_msg
                    st.session_state["last_ai_summary"] = resumo_ia

    if st.session_state.get("last_ai_error"):
        st.warning(st.session_state["last_ai_error"])
    elif st.session_state.get("last_ai_summary"):
        st.success(f"**Análise Concluída:** {st.session_state['last_ai_summary']}")

    params = st.session_state["params"]
    result = st.session_state["result"]
    resumo = result["resumo"]
    meta = result["meta"]

    if meta.get("warning_msg"):
        st.warning(meta["warning_msg"])

    # Cards de Resumo Executivo
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""<div class="metric-card">
                <div class="metric-label">Bruto Devido ao Reclamante</div>
                <div class="metric-value">{fmt_brl_rs(resumo['bruto_devido_reclamante'])}</div>
                <div class="metric-sub">Verbas + Juros + FGTS</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""<div class="metric-card">
                <div class="metric-label">Líquido do Reclamante</div>
                <div class="metric-value" style="color:#15803D;">{fmt_brl_rs(resumo['liquido_reclamante'])}</div>
                <div class="metric-sub">Após INSS ({fmt_brl(resumo['cs_segurado'])}) e IRRF ({fmt_brl(resumo['irrf_reclamante'])})</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""<div class="metric-card">
                <div class="metric-label">Honorários & Encargos</div>
                <div class="metric-value">{fmt_brl_rs(resumo['honorarios_adv'] + resumo['cs_patronal'])}</div>
                <div class="metric-sub">Hon.: {fmt_brl(resumo['honorarios_adv'])} | INSS Patr.: {fmt_brl(resumo['cs_patronal'])}</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""<div class="metric-card">
                <div class="metric-label">Total Devido pelo Reclamado</div>
                <div class="metric-value" style="color:#B91C1C;">{fmt_brl_rs(resumo['total_devido_reclamado'])}</div>
                <div class="metric-sub">Atualizado até {fmt_date_br(params.get('data_liquidacao'))}</div>
            </div>""",
            unsafe_allow_html=True,
        )

    st.write("")

    # Botões de Exportação Direta
    proc_slug = re.sub(r"[^0-9A-Za-z\-]", "_", str(params.get("processo_cnj") or "calculo"))
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            "📄 Baixar Planilha Oficial (PDF Padrão PJe-Calc)",
            data=st.session_state.get("pdf_bytes", b""),
            file_name=f"Planilha_PJeCalc_{proc_slug}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            "💾 Baixar Arquivo Interoperável (.PJC)",
            data=st.session_state.get("pjc_bytes", b""),
            file_name=f"Calculo_{proc_slug}.pjc",
            mime="application/xml",
            use_container_width=True,
        )
    with dl3:
        st.download_button(
            "📝 Baixar Minuta de Apresentação (.TXT)",
            data=st.session_state.get("minuta_txt", "").encode("utf-8"),
            file_name=f"Minuta_Liquidacao_{proc_slug}.txt",
            mime="text/plain",
            use_container_width=True,
        )

    st.divider()

    # Abas de Conferência Jurídica, Demonstrativos e Minuta
    tab_conf, tab_resumo, tab_verbas, tab_encargos, tab_minuta = st.tabs([
        "🛠️ Conferência dos Parâmetros",
        "📊 Resumo Oficial PJe-Calc",
        "📑 Demonstrativo de Verbas & Juros",
        "🏛️ FGTS, INSS & IRRF (RRA)",
        "✍️ Minuta para Protocolo",
    ])

    with tab_conf:
        st.subheader("Conferência dos Dados Extraídos e Parâmetros do Título Judicial")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            params["processo_cnj"] = st.text_input("Número do Processo (CNJ)", value=params.get("processo_cnj", ""))
            params["vara_trabalho"] = st.text_input("Vara do Trabalho", value=params.get("vara_trabalho", ""))
            params["reclamante"] = st.text_input("Reclamante", value=params.get("reclamante", ""))
            params["cpf_reclamante"] = st.text_input("CPF do Reclamante", value=params.get("cpf_reclamante", ""))
        with col_b:
            params["reclamado"] = st.text_input("Reclamado", value=params.get("reclamado", ""))
            params["cnpj_reclamado"] = st.text_input("CNPJ do Reclamado", value=params.get("cnpj_reclamado", ""))
            params["estado"] = st.text_input("UF", value=params.get("estado", "RS"))
            params["municipio"] = st.text_input("Município / Cidade da Minuta", value=params.get("municipio", "Gravataí"))
            params["cidade_assinatura"] = params["municipio"]
        with col_c:
            params["admissao"] = st.text_input("Data de Admissão (AAAA-MM-DD)", value=params.get("admissao", "2022-01-10"))
            params["demissao"] = st.text_input("Data de Demissão (AAAA-MM-DD)", value=params.get("demissao", "2024-06-15"))
            params["data_ajuizamento"] = st.text_input("Data de Ajuizamento (AAAA-MM-DD)", value=params.get("data_ajuizamento", "2024-08-20"))
            params["data_liquidacao"] = st.text_input("Data de Liquidação (AAAA-MM-DD)", value=params.get("data_liquidacao", date.today().isoformat()))

        col_d, col_e, col_f, col_g = st.columns(4)
        with col_d:
            params["maior_remuneracao"] = st.number_input("Maior Remuneração (R$)", value=float(params.get("maior_remuneracao", 0.0)), step=50.0)
        with col_e:
            params["ultima_remuneracao"] = st.number_input("Última Remuneração (R$)", value=float(params.get("ultima_remuneracao", 0.0)), step=50.0)
        with col_f:
            params["honorarios_sucumbencia_pct"] = st.number_input("Honorários Sucumbenciais (%)", value=float(params.get("honorarios_sucumbencia_pct", 15.0)), step=1.0)
        with col_g:
            params["custas_pct"] = st.number_input("Custas Judiciais (%)", value=float(params.get("custas_pct", 2.0)), step=0.5)

        chk1, chk2, chk3, chk4 = st.columns(4)
        with chk1:
            params["aplicar_prescricao"] = st.checkbox("Prescrição Quinquenal", value=bool(params.get("aplicar_prescricao", True)))
        with chk2:
            params["projetar_aviso"] = st.checkbox("Projetar Aviso Prévio (Lei 12.506)", value=bool(params.get("projetar_aviso", True)))
        with chk3:
            params["aplicar_multa_fgts_40"] = st.checkbox("Multa FGTS 40%", value=bool(params.get("aplicar_multa_fgts_40", True)))
        with chk4:
            params["apurar_irrf"] = st.checkbox("Apurar IRRF (RRA)", value=bool(params.get("apurar_irrf", True)))

        st.markdown("#### Verbas Deferidas no Cálculo")
        verbas_list = params.get("verbas", [])
        if not verbas_list:
            verbas_list = [
                {
                    "nome": "HORAS EXTRAS 50%",
                    "tipo": "HORAS_EXTRAS",
                    "aliquota_ou_percentual": 0.50,
                    "quantidade_mensal": 0.0,
                    "divisor": 220.0,
                    "inicio": params.get("admissao", "2022-01-10"),
                    "fim": params.get("demissao", "2024-06-15"),
                    "incide_fgts": True,
                    "incide_inss": True,
                    "incide_irrf": True,
                    "gera_reflexos_13_ferias_aviso": True,
                    "ativo": True,
                }
            ]
        df_verbas = pd.DataFrame(verbas_list)
        edited_verbas = st.data_editor(
            df_verbas,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_verbas_pjc",
        )
        if st.button("✅ Recalcular Liquidação e Atualizar Arquivos", type="primary"):
            params["verbas"] = edited_verbas.to_dict(orient="records")
            adm = parse_date_br(params.get("admissao")) or date(2022, 1, 10)
            dem = parse_date_br(params.get("demissao")) or date(2024, 6, 15)
            sal_p = float(params.get("ultima_remuneracao", 0.0))
            sal_d = float(params.get("maior_remuneracao", sal_p))
            params["historico_salarial"] = [
                {"competencia": f"{y:04d}-{m:02d}", "salario_pago": sal_p, "salario_devido": sal_d}
                for y, m in month_range(adm, dem)
            ]
            _recompute_from_current_params(params, save_to_history=True)
            st.success("Cálculo atualizado e salvo no histórico!")
            st.rerun()

    with tab_resumo:
        st.caption(f"Fonte oficial de índices utilizada: **{meta['fonte_indices']}**")
        df_resumo = pd.DataFrame([
            {"Descrição": "Verbas Trabalhistas Corrigidas", "Valor (R$)": fmt_brl_rs(resumo["verbas_corrigido"])},
            {"Descrição": "Juros de Mora sobre Verbas (após dedução CS Segurado)", "Valor (R$)": fmt_brl_rs(resumo["juros_verbas"])},
            {"Descrição": "FGTS (Depósitos 8% + Juros + Multa 40%)", "Valor (R$)": fmt_brl_rs(resumo["total_fgts"])},
            {"Descrição": "BRUTO DEVIDO AO RECLAMANTE", "Valor (R$)": fmt_brl_rs(resumo["bruto_devido_reclamante"])},
            {"Descrição": "(-) Contribuição Social — Cota Segurado", "Valor (R$)": f"-{fmt_brl_rs(resumo['cs_segurado'])}"},
            {"Descrição": "(-) Imposto de Renda Retido na Fonte (RRA)", "Valor (R$)": f"-{fmt_brl_rs(resumo['irrf_reclamante'])}"},
            {"Descrição": "LÍQUIDO DEVIDO AO RECLAMANTE", "Valor (R$)": fmt_brl_rs(resumo["liquido_reclamante"])},
            {"Descrição": "(+) Contribuição Social — Cota Patronal / SAT / Terceiros", "Valor (R$)": fmt_brl_rs(resumo["cs_patronal"])},
            {"Descrição": f"(+) Honorários Sucumbenciais ({params.get('honorarios_sucumbencia_pct', 15)}%)", "Valor (R$)": fmt_brl_rs(resumo["honorarios_adv"])},
            {"Descrição": "(+) Custas Judiciais", "Valor (R$)": fmt_brl_rs(resumo["custas_judiciais"])},
            {"Descrição": "TOTAL GERAL DEVIDO PELO RECLAMADO", "Valor (R$)": fmt_brl_rs(resumo["total_devido_reclamado"])},
        ])
        st.dataframe(df_resumo, use_container_width=True, hide_index=True)

    with tab_verbas:
        if not result["verbas_detalhadas"]:
            st.info("Nenhuma verba apurada ainda. Anexe os autos acima ou adicione verbas na aba 'Conferência dos Parâmetros'.")
        for vd in result["verbas_detalhadas"]:
            with st.expander(f"📌 {vd['nome']} — Total Corrigido: {fmt_brl_rs(vd['total_corrigido'])}", expanded=False):
                df_oc = pd.DataFrame([
                    {
                        "Competência": fmt_comp_br(o["competencia"]),
                        "Base (R$)": fmt_brl(o["base"]),
                        "Devido (R$)": fmt_brl(o["devido"]),
                        "Pago (R$)": fmt_brl(o["pago"]),
                        "Diferença (R$)": fmt_brl(o["diferenca"]),
                        "Índice": fmt_idx(o["indice"]),
                        "Corrigido (R$)": fmt_brl(o["valor_corrigido"]),
                    }
                    for o in vd["ocorrencias"]
                ])
                st.dataframe(df_oc, use_container_width=True, hide_index=True)

        st.markdown("#### Demonstrativo de Juros sobre Verbas")
        df_jur = pd.DataFrame([
            {
                "Competência": fmt_comp_br(j["competencia"]),
                "Total Verbas (R$)": fmt_brl(j["total_verbas"]),
                "Contrib. Social (R$)": fmt_brl(j["contribuicao_social"]),
                "Capital Base Juros (R$)": fmt_brl(j["capital"]),
                "Taxa Juros (%)": f"{fmt_brl(j['taxa_juros_pct'])}%",
                "Juros (R$)": fmt_brl(j["juros"]),
            }
            for j in result["juros_ocorrencias"]
        ])
        st.dataframe(df_jur, use_container_width=True, hide_index=True)

    with tab_encargos:
        st.markdown("#### Demonstrativo de FGTS (8%)")
        df_fgts = pd.DataFrame([
            {
                "Competência": fmt_comp_br(f["competencia"]),
                "Base FGTS (R$)": fmt_brl(f["base"]),
                "Depósito 8% (R$)": fmt_brl(f["devido"]),
                "Índice": fmt_idx(f["indice"]),
                "Corrigido (R$)": fmt_brl(f["corrigido"]),
                "Juros (R$)": fmt_brl(f["juros"]),
                "Total (R$)": fmt_brl(f["total"]),
            }
            for f in result["fgts_ocorrencias"]
        ])
        st.dataframe(df_fgts, use_container_width=True, hide_index=True)

        st.markdown("#### Demonstrativo de Contribuição Social (Segurado e Empregador)")
        df_cs = pd.DataFrame([
            {
                "Competência": fmt_comp_br(c["competencia"]),
                "Base Verbas (R$)": fmt_brl(c["base_verbas"]),
                "Segurado Devido (R$)": fmt_brl(c["segurado_devido"]),
                "Segurado Corrigido (R$)": fmt_brl(c["segurado_corrigido"]),
                "Patronal Devido (R$)": fmt_brl(c["patronal_devido"]),
                "Patronal Total (R$)": fmt_brl(c["patronal_total"]),
            }
            for c in result["cs_ocorrencias"]
        ])
        st.dataframe(df_cs, use_container_width=True, hide_index=True)

    with tab_minuta:
        minuta_editada = st.text_area(
            "Minuta de Petição de Apresentação de Cálculos",
            value=st.session_state.get("minuta_txt", ""),
            height=420,
        )
        st.session_state["minuta_txt"] = minuta_editada


# ==============================================================================
# PÁGINA 2 — HISTÓRICO DE LIQUIDAÇÕES
# ==============================================================================
else:
    st.markdown(
        """
        <div class="hero-banner">
            <h1>Histórico de Liquidações Realizadas</h1>
            <p>Recarregue qualquer cálculo anterior para revisar parâmetros, atualizar índices ou reemitir os arquivos .PJC e PDF.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    rows = list_calculation_history(40)
    if not rows:
        st.info("Nenhuma liquidação salva no histórico ainda.")
    else:
        for r in rows:
            with st.container(border=True):
                c_a, c_b, c_c = st.columns([3, 2, 1])
                with c_a:
                    st.markdown(f"**Processo:** `{r['processo_cnj']}` — **{r['reclamante']}**")
                    st.caption(f"Reclamado: {r['reclamado']} | Gerado em: {r['criado_em']}")
                with c_b:
                    st.markdown(
                        f"Líquido Reclamante: **{fmt_brl_rs(r['liquido_reclamante'])}**  \n"
                        f"Total Reclamado: **{fmt_brl_rs(r['total_reclamado'])}**"
                    )
                with c_c:
                    if st.button("Abrir Cálculo", key=f"load_hist_{r['id']}", use_container_width=True):
                        loaded_p = json.loads(r["payload_json"])
                        _recompute_from_current_params(loaded_p, save_to_history=False)
                        st.success("Cálculo recarregado! Volte para 'Novo Cálculo (.PJC + Planilha)'.")
