# ... existing code ...
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, getcontext
import hashlib
import io
import json
import os
import re
import sqlite3
from typing import Any, Dict, List, Optional, Tuple
import uuid
import xml.etree.ElementTree as ET

    # Inicialização do estado da sessão
    if "params" not in st.session_state:
        st.session_state["params"] = _default_empty_params()
    if "calc_result" not in st.session_state:
        st.session_state["calc_result"] = None
    if "audit_report" not in st.session_state:
        st.session_state["audit_report"] = None
    if "has_analyzed" not in st.session_state:
        st.session_state["has_analyzed"] = False
    if "sentence_pasted_text" not in st.session_state:
        st.session_state["sentence_pasted_text"] = ""
    if "lawyer_notes" not in st.session_state:
        st.session_state["lawyer_notes"] = ""
    if "data_base_liq" not in st.session_state:
        st.session_state["data_base_liq"] = date.today()
    if "calc_json_text" not in st.session_state:
        st.session_state["calc_json_text"] = ""
    if "current_calc_uid" not in st.session_state:
        st.session_state["current_calc_uid"] = ""
# ... existing code ...
def build_pjc_xml(calc_result: Dict[str, Any]) -> bytes:
    """
    Módulo 11 — Exportação Estruturada (.PJC XML) compatível com estrutura PJe-Calc.
    NOTA DE TRANSPARÊNCIA: Gerado por motor independente; não contém assinatura digital oficial do CSJT.
    """
    meta = calc_result["meta"]
    resumo = calc_result["resumo"]
    root = ET.Element("Calculo", attrib={"versaoSistema": "2.13.2", "gerador": "CalcMachine-MotorDeterministico"})

    ident = ET.SubElement(root, "Identificacao")
    ET.SubElement(ident, "NumeroProcesso").text = str(meta.get("processo", ""))
    ET.SubElement(ident, "NumeroCalculo").text = str(meta.get("numero_calculo", "1"))
    ET.SubElement(ident, "Reclamante").text = str(meta.get("reclamante", ""))
    ET.SubElement(ident, "Reclamado").text = str(meta.get("reclamado", ""))
    ET.SubElement(ident, "Estado").text = str(meta.get("estado", "RS"))
    ET.SubElement(ident, "Municipio").text = str(meta.get("municipio", "PORTO ALEGRE"))

    dados = ET.SubElement(root, "ParametrosCalculo")
    ET.SubElement(dados, "DataAdmissao").text = fmt_br_date(meta.get("admissao"))
    ET.SubElement(dados, "DataDemissao").text = fmt_br_date(meta.get("demissao"))
    ET.SubElement(dados, "DataAjuizamento").text = fmt_br_date(meta.get("ajuizamento"))
    ET.SubElement(dados, "DataLiquidacao").text = fmt_br_date(meta.get("liquidacao"))
    ET.SubElement(dados, "InicioCalculo").text = fmt_br_date(meta.get("inicio_calc"))
    ET.SubElement(dados, "TerminoCalculo").text = fmt_br_date(meta.get("fim_calc"))
    ET.SubElement(dados, "CargaHoraria").text = str(meta.get("carga_horaria", 220))
    ET.SubElement(dados, "Divisor").text = str(meta.get("divisor", 220))
    ET.SubElement(dados, "SalarioBase").text = f"{D(meta.get('salario_base', 0)):.2f}"
    ET.SubElement(dados, "CriterioCorrecao").text = str(meta.get("criterio_correcao", ""))

    hist_el = ET.SubElement(root, "HistoricoSalarial")
    base_h = ET.SubElement(hist_el, "BaseHistorico", attrib={"nome": "SALARIO CONTRATUAL"})
    ET.SubElement(base_h, "DataInicio").text = fmt_br_date(meta.get("admissao"))
    ET.SubElement(base_h, "DataFim").text = fmt_br_date(meta.get("demissao"))
    ET.SubElement(base_h, "Valor").text = f"{D(meta.get('salario_base', 0)):.2f}"

    verbas_el = ET.SubElement(root, "Verbas")
    for idx_v, v in enumerate(calc_result.get("verbas", []), start=1):
        v_el = ET.SubElement(
            verbas_el,
            "Verba",
            attrib={
                "ordem": str(idx_v),
                "codigo": str(v.get("codigo", "")),
                "ativo": "true",
                "caracteristica": "COMUM",
                "ocorrenciaPagamento": "MENSAL",
            },
        )
        ET.SubElement(v_el, "Nome").text = str(v["nome"])
        ET.SubElement(v_el, "VerbaOrigem").text = str(v.get("verba_origem", ""))
        ET.SubElement(v_el, "Periodo").text = str(v["periodo_global"])
        ET.SubElement(v_el, "Formula").text = str(v["formula_pje"])
        ET.SubElement(v_el, "IncideINSS").text = "true" if v.get("incide_inss") else "false"
        ET.SubElement(v_el, "IncideIRPF").text = "true" if v.get("incide_irrf") else "false"
        ET.SubElement(v_el, "IncideFGTS").text = "true" if v.get("incide_fgts") else "false"
        ET.SubElement(v_el, "ValorDevido").text = f"{D(v['valor']):.2f}"
        ET.SubElement(v_el, "ValorCorrigido").text = f"{D(v['valor_corrigido']):.2f}"
        ET.SubElement(v_el, "Juros").text = f"{D(v['juros']):.2f}"
        ET.SubElement(v_el, "Total").text = f"{D(v['total_verba']):.2f}"

        ocs_el = ET.SubElement(v_el, "Ocorrencias")
        for oc in v.get("ocorrencias", []):
            oc_el = ET.SubElement(ocs_el, "Ocorrencia")
            ET.SubElement(oc_el, "Competencia").text = str(oc.get("competencia", ""))
            ET.SubElement(oc_el, "Periodo").text = str(oc.get("periodo_mensal", ""))
            ET.SubElement(oc_el, "Base").text = f"{D(oc.get('base', 0)):.2f}"
            ET.SubElement(oc_el, "Divisor").text = f"{D(oc.get('divisor', 1)):.4f}"
            ET.SubElement(oc_el, "Multiplicador").text = f"{D(oc.get('multiplicador', 1)):.8f}"
            ET.SubElement(oc_el, "Quantidade").text = f"{D(oc.get('quantidade', 1)):.4f}"
            ET.SubElement(oc_el, "Devido").text = f"{D(oc.get('devido', 0)):.2f}"
            ET.SubElement(oc_el, "Pago").text = f"{D(oc.get('pago', 0)):.2f}"
            ET.SubElement(oc_el, "Diferenca").text = f"{D(oc.get('diferenca', 0)):.2f}"
            ET.SubElement(oc_el, "IndiceCorrecao").text = f"{D(oc.get('indice_correcao', 1)):.9f}"
            ET.SubElement(oc_el, "ValorCorrigido").text = f"{D(oc.get('valor_corrigido', 0)):.2f}"
            ET.SubElement(oc_el, "Juros").text = f"{D(oc.get('juros', 0)):.2f}"

    res_el = ET.SubElement(root, "ResumoLiquidacao")
    for k, val in resumo.items():
        if isinstance(val, Decimal):
            ET.SubElement(res_el, k).text = f"{val:.2f}"

    xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return xml_str


# ==============================================================================
# PERSISTÊNCIA E HISTÓRICO DE CÁLCULOS + SERIALIZAÇÃO JSON EDITÁVEL (CALCMACHINE)
# ==============================================================================
HISTORY_DB_PATH = "calcmachine_historico.db"


def _init_history_db() -> None:
    """Inicializa o banco SQLite local para histórico completo e versionamento de cálculos."""
    try:
        conn = sqlite3.connect(HISTORY_DB_PATH)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS calc_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                calc_uid TEXT NOT NULL,
                versao INTEGER NOT NULL,
                processo TEXT,
                reclamante TEXT,
                reclamado TEXT,
                total_execucao TEXT,
                liquido_reclamante TEXT,
                status_auditoria TEXT,
                payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def serialize_params_to_editable_json(params: Dict[str, Any], calc_result: Optional[Dict[str, Any]] = None) -> str:
    """Converte os parâmetros do cálculo em JSON limpo, estruturado e editável pelo usuário."""
    def _dt_str(v: Any) -> Optional[str]:
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y")
        if isinstance(v, date):
            return v.strftime("%d/%m/%Y")
        if isinstance(v, str) and v.strip():
            return v.strip()
        return None

    clean_obj: Dict[str, Any] = {
        "identificacao": {
            "processo": params.get("processo", ""),
            "numero_calculo": params.get("numero_calculo", "1"),
            "estado": params.get("estado", "RS"),
            "municipio": params.get("municipio", "PORTO ALEGRE"),
            "vara": params.get("vara", "VARA DO TRABALHO"),
            "reclamante": params.get("reclamante", ""),
            "reclamado": params.get("reclamado", ""),
            "parte_representada": params.get("parte_representada", "Reclamante"),
            "advogado_assinante": params.get("advogado_assinante", ""),
            "oab_advogado": params.get("oab_advogado", ""),
        },
        "contrato_e_datas": {
            "admissao": _dt_str(params.get("admissao")),
            "demissao": _dt_str(params.get("demissao")),
            "ajuizamento": _dt_str(params.get("ajuizamento")),
            "liquidacao": _dt_str(params.get("liquidacao")),
            "salario_base": float(params.get("salario_base", 0.0)),
            "modalidade_rescisao": params.get("modalidade_rescisao", "Dispensa sem justa causa"),
            "carga_horaria": int(params.get("carga_horaria", 220)),
            "divisor": float(params.get("divisor", 220.0)),
            "projetar_aviso_previo": bool(params.get("projetar_aviso_previo", False)),
            "aplicar_prescricao_quinquenal": bool(params.get("aplicar_prescricao_quinquenal", False)),
        },
        "criterios_correcao": {
            "criterio": params.get("criterio_correcao", "ADC 58 STF (IPCA-E pré-judicial + SELIC) + Lei 14.905/2024"),
            "inss_patronal_pct": float(params.get("inss_patronal_pct", 20.0)),
            "rat_sat_pct": float(params.get("rat_sat_pct", 2.0)),
            "terceiros_pct": float(params.get("terceiros_pct", 5.8)),
        },
        "verbas_deferidas": {
            "horas_extras_50": {
                "ativo": bool(params.get("tem_he50", False)),
                "quantidade_mensal": float(params.get("qtd_he50_mensal", 0.0)),
                "reflexos_13_ferias": bool(params.get("reflexos_he50", False)),
            },
            "horas_extras_100": {
                "ativo": bool(params.get("tem_he100", False)),
                "quantidade_mensal": float(params.get("qtd_he100_mensal", 0.0)),
                "reflexos_13_ferias": bool(params.get("reflexos_he100", False)),
            },
            "intervalo_intrajornada": {
                "ativo": bool(params.get("tem_intrajornada", False)),
                "quantidade_diaria_horas": float(params.get("qtd_intrajornada_mensal", 1.0)),
                "teto_judicial_reais": float(params.get("teto_intrajornada", 0.0)),
                "escopo_teto": params.get("teto_intrajornada_escopo", "VERBA_PRINCIPAL"),
                "reflexos_13_ferias": bool(params.get("reflexos_intrajornada", False)),
            },
            "devolucao_descontos": {
                "ativo": bool(params.get("tem_descontos_indevidos", False)),
                "valor_principal": float(params.get("valor_descontos_indevidos", 0.0)),
                "teto_judicial_reais": float(params.get("teto_descontos_indevidos", 0.0)),
            },
            "fgts_e_multa_40": bool(params.get("reflexos_em_fgts", False)),
            "multa_art_477": bool(params.get("multa_477", False)),
            "multa_art_467": bool(params.get("multa_467", False)),
            "verbas_personalizadas": params.get("verbas_personalizadas", []),
        },
        "honorarios_custas_abatimentos": {
            "honorarios_sucumbencia_pct": float(params.get("honorarios_sucumbencia_pct", 0.0)),
            "honorarios_reciprocos_pct": float(params.get("honorarios_reciprocos_pct", 0.0)),
            "base_improcedentes_reciprocos": float(params.get("base_pedidos_improcedentes", 0.0)),
            "custas_processuais": float(params.get("custas_processuais", 0.0)),
            "custas_ja_recolhidas": bool(params.get("custas_ja_recolhidas", False)),
            "valor_pago_abatimento": float(params.get("valor_pago_abatimento", 0.0)),
        },
        "validacao_humana_confirmada": bool(params.get("_human_confirmed_ambiguities", False)),
    }
    if calc_result is not None:
        resumo = calc_result["resumo"]
        clean_obj["resumo_calculado_referencia"] = {
            "bruto_reclamante": float(resumo["bruto_devido_reclamante"]),
            "inss_segurado": float(resumo["inss_empregado"]),
            "irpf_rra": float(resumo["irpf_reclamante"]),
            "liquido_reclamante": float(resumo["liquido_devido_reclamante"]),
            "fgts": float(resumo["total_fgts_bruto"]),
            "honorarios": float(resumo["honorarios_valor"]),
            "custas": float(resumo["custas_devidas_execucao"]),
            "total_devido_reclamado": float(resumo["total_devido_reclamado"]),
        }
    return json.dumps(clean_obj, indent=2, ensure_ascii=False)


def apply_editable_json_to_params(json_str: str, base_params: Dict[str, Any]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Valida e aplica edições feitas pelo usuário no JSON estruturado de volta para os parâmetros do motor."""
    try:
        data = json.loads(json_str)
    except Exception as exc:
        return None, f"JSON inválido: {exc}"

    updated = dict(base_params)
    ident = data.get("identificacao", {})
    ctr = data.get("contrato_e_datas", {})
    crit = data.get("criterios_correcao", {})
    verb = data.get("verbas_deferidas", {})
    hon = data.get("honorarios_custas_abatimentos", {})

    for k in ("processo", "numero_calculo", "estado", "municipio", "vara", "reclamante", "reclamado", "parte_representada", "advogado_assinante", "oab_advogado"):
        if k in ident and ident[k] is not None:
            updated[k] = str(ident[k]).strip()

    for dt_k in ("admissao", "demissao", "ajuizamento", "liquidacao"):
        if ctr.get(dt_k):
            dt_parsed = parse_br_date_optional(ctr[dt_k])
            if not dt_parsed:
                return None, f"Data inválida no JSON em '{dt_k}': {ctr[dt_k]} (use DD/MM/AAAA)."
            updated[dt_k] = dt_parsed

    if "salario_base" in ctr:
        updated["salario_base"] = float(D(ctr["salario_base"]))
    if "modalidade_rescisao" in ctr:
        updated["modalidade_rescisao"] = str(ctr["modalidade_rescisao"])
    if "carga_horaria" in ctr:
        updated["carga_horaria"] = int(ctr["carga_horaria"])
    if "divisor" in ctr:
        updated["divisor"] = float(D(ctr["divisor"]))
    if "projetar_aviso_previo" in ctr:
        updated["projetar_aviso_previo"] = bool(ctr["projetar_aviso_previo"])
    if "aplicar_prescricao_quinquenal" in ctr:
        updated["aplicar_prescricao_quinquenal"] = bool(ctr["aplicar_prescricao_quinquenal"])

    if "criterio" in crit:
        updated["criterio_correcao"] = str(crit["criterio"])
    if "inss_patronal_pct" in crit:
        updated["inss_patronal_pct"] = float(D(crit["inss_patronal_pct"]))
    if "rat_sat_pct" in crit:
        updated["rat_sat_pct"] = float(D(crit["rat_sat_pct"]))
    if "terceiros_pct" in crit:
        updated["terceiros_pct"] = float(D(crit["terceiros_pct"]))

    he50 = verb.get("horas_extras_50", {})
    updated["tem_he50"] = bool(he50.get("ativo", updated.get("tem_he50", False)))
    updated["qtd_he50_mensal"] = float(D(he50.get("quantidade_mensal", updated.get("qtd_he50_mensal", 0.0))))
    updated["reflexos_he50"] = bool(he50.get("reflexos_13_ferias", updated.get("reflexos_he50", False)))

    he100 = verb.get("horas_extras_100", {})
    updated["tem_he100"] = bool(he100.get("ativo", updated.get("tem_he100", False)))
    updated["qtd_he100_mensal"] = float(D(he100.get("quantidade_mensal", updated.get("qtd_he100_mensal", 0.0))))
    updated["reflexos_he100"] = bool(he100.get("reflexos_13_ferias", updated.get("reflexos_he100", False)))

    intra = verb.get("intervalo_intrajornada", {})
    updated["tem_intrajornada"] = bool(intra.get("ativo", updated.get("tem_intrajornada", False)))
    updated["qtd_intrajornada_mensal"] = float(D(intra.get("quantidade_diaria_horas", updated.get("qtd_intrajornada_mensal", 1.0))))
    updated["teto_intrajornada"] = float(D(intra.get("teto_judicial_reais", updated.get("teto_intrajornada", 0.0))))
    updated["teto_intrajornada_escopo"] = str(intra.get("escopo_teto", updated.get("teto_intrajornada_escopo", "VERBA_PRINCIPAL")))
    updated["reflexos_intrajornada"] = bool(intra.get("reflexos_13_ferias", updated.get("reflexos_intrajornada", False)))

    desc = verb.get("devolucao_descontos", {})
    updated["tem_descontos_indevidos"] = bool(desc.get("ativo", updated.get("tem_descontos_indevidos", False)))
    updated["valor_descontos_indevidos"] = float(D(desc.get("valor_principal", updated.get("valor_descontos_indevidos", 0.0))))
    updated["teto_descontos_indevidos"] = float(D(desc.get("teto_judicial_reais", updated.get("teto_descontos_indevidos", 0.0))))

    if "fgts_e_multa_40" in verb:
        updated["reflexos_em_fgts"] = bool(verb["fgts_e_multa_40"])
    if "multa_art_477" in verb:
        updated["multa_477"] = bool(verb["multa_art_477"])
    if "multa_art_467" in verb:
        updated["multa_467"] = bool(verb["multa_art_467"])
    if "verbas_personalizadas" in verb and isinstance(verb["verbas_personalizadas"], list):
        updated["verbas_personalizadas"] = verb["verbas_personalizadas"]

    if "honorarios_sucumbencia_pct" in hon:
        updated["honorarios_sucumbencia_pct"] = float(D(hon["honorarios_sucumbencia_pct"]))
    if "honorarios_reciprocos_pct" in hon:
        updated["honorarios_reciprocos_pct"] = float(D(hon["honorarios_reciprocos_pct"]))
    if "base_improcedentes_reciprocos" in hon:
        updated["base_pedidos_improcedentes"] = float(D(hon["base_improcedentes_reciprocos"]))
    if "custas_processuais" in hon:
        updated["custas_processuais"] = float(D(hon["custas_processuais"]))
    if "custas_ja_recolhidas" in hon:
        updated["custas_ja_recolhidas"] = bool(hon["custas_ja_recolhidas"])
    if "valor_pago_abatimento" in hon:
        updated["valor_pago_abatimento"] = float(D(hon["valor_pago_abatimento"]))

    if "validacao_humana_confirmada" in data:
        updated["_human_confirmed_ambiguities"] = bool(data["validacao_humana_confirmada"])

    return updated, None


def save_calculation_version(calc_uid: str, params: Dict[str, Any], calc_result: Dict[str, Any], audit_report: Dict[str, Any], json_str: str) -> None:
    """Grava versão imutável do cálculo no histórico local."""
    _init_history_db()
    try:
        conn = sqlite3.connect(HISTORY_DB_PATH)
        cur = conn.execute("SELECT COALESCE(MAX(versao), 0) FROM calc_history WHERE calc_uid = ?", (calc_uid,))
        next_ver = int(cur.fetchone()[0]) + 1
        resumo = calc_result["resumo"]
        conn.execute(
            """
            INSERT INTO calc_history (
                calc_uid, versao, processo, reclamante, reclamado,
                total_execucao, liquido_reclamante, status_auditoria, payload_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                calc_uid,
                next_ver,
                str(params.get("processo", "")),
                str(params.get("reclamante", "")),
                str(params.get("reclamado", "")),
                fmt_brl(resumo["total_devido_reclamado"]),
                fmt_brl(resumo["liquido_devido_reclamante"]),
                str(audit_report.get("status", "")),
                json_str,
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            ),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def load_calculation_history(limit: int = 30) -> List[Dict[str, Any]]:
    """Lista os cálculos e versões armazenados no histórico."""
    _init_history_db()
    rows_out: List[Dict[str, Any]] = []
    try:
        conn = sqlite3.connect(HISTORY_DB_PATH)
        cur = conn.execute(
            """
            SELECT id, calc_uid, versao, processo, reclamante, reclamado,
                   total_execucao, liquido_reclamante, status_auditoria, payload_json, created_at
            FROM calc_history
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        )
        for r in cur.fetchall():
            rows_out.append(
                {
                    "id": r[0],
                    "calc_uid": r[1],
                    "versao": r[2],
                    "processo": r[3] or "Sem número",
                    "reclamante": r[4] or "Não informado",
                    "reclamado": r[5] or "Não informado",
                    "total_execucao": r[6],
                    "liquido_reclamante": r[7],
                    "status_auditoria": r[8],
                    "payload_json": r[9],
                    "created_at": r[10],
                }
            )
        conn.close()
    except Exception:
        pass
    return rows_out
# ... existing code ...
def _recompute_from_current_params(updated_params: Dict[str, Any], save_to_history: bool = True) -> None:
    """Recalcula imediatamente o motor matemático, atualiza o JSON estruturado e grava no histórico."""
    res_calc = run_pjecalc_engine(updated_params)
    aud_rep = run_triple_audit(updated_params, res_calc)
    pdf_b = build_pjecalc_pdf(res_calc, aud_rep)
    minuta_t = build_minuta_peticao(updated_params, res_calc, aud_rep, st.session_state.get("lawyer_notes", ""))
    pjc_b = build_pjc_xml(res_calc)
    exp_aud = run_export_consistency_audit(res_calc, pdf_b, minuta_t, pjc_b)
    aud_rep["auditoria_exportacao"] = exp_aud
    if not exp_aud["ok"]:
        aud_rep["status"] = "REPROVADO"
        aud_rep["can_export_definitive"] = False
    json_str = serialize_params_to_editable_json(updated_params, res_calc)
    if not st.session_state.get("current_calc_uid"):
        st.session_state["current_calc_uid"] = uuid.uuid4().hex[:12]
    st.session_state["params"] = updated_params
    st.session_state["calc_result"] = res_calc
    st.session_state["audit_report"] = aud_rep
    st.session_state["calc_json_text"] = json_str
    if save_to_history:
        save_calculation_version(st.session_state["current_calc_uid"], updated_params, res_calc, aud_rep, json_str)


def _execute_full_pipeline(
    uploaded_files: Optional[List[Any]],
    sentence_pasted_text: str,
    lawyer_notes: str,
    data_base_liq: date,
    api_key: str,
    selected_contract_idx: Optional[int] = None,
    confirm_human: bool = False,
) -> None:
    """Executa o fluxo CalcMachine: Sentença/Autos -> IA & Extração -> JSON + .PJC -> Histórico Versionado."""
    combined_text = "\n\n".join(x for x in [sentence_pasted_text.strip(), lawyer_notes.strip()] if x)
    ingestion = ingest_documents_with_page_tracking(uploaded_files, sentence_pasted_text, lawyer_notes)
    parsed = parse_sentence_parameters(sentence_pasted_text, lawyer_notes, ingestion_report=ingestion)
    parsed["liquidacao"] = data_base_liq

    # Preserva nome do advogado e OAB já digitados na sessão, se houver
    prev_p = st.session_state.get("params", {})
    if prev_p.get("advogado_assinante"):
        parsed["advogado_assinante"] = prev_p["advogado_assinante"]
    if prev_p.get("oab_advogado"):
        parsed["oab_advogado"] = prev_p["oab_advogado"]

    # Enriquecimento semântico via Gemini API (sem substituir o motor matemático)
    ai_data = analyze_documents_with_gemini_if_configured(ingestion["pages_index"], combined_text, api_key)
    if isinstance(ai_data, dict):
        for str_k in ("processo", "reclamante", "reclamado", "advogado_assinante", "oab_advogado", "modalidade_rescisao"):
            if ai_data.get(str_k) and not parsed.get(str_k):
                parsed[str_k] = str(ai_data[str_k]).strip()
        if ai_data.get("parte_representada") in ("Reclamante", "Reclamado"):
            parsed["parte_representada"] = ai_data["parte_representada"]
        for dt_k in ("admissao", "demissao", "ajuizamento"):
            dt_parsed = parse_br_date_optional(ai_data.get(dt_k))
            if dt_parsed and not parsed.get(dt_k):
                parsed[dt_k] = dt_parsed
        if ai_data.get("salario_base") and D(parsed.get("salario_base", 0)) <= 0:
            parsed["salario_base"] = float(D(ai_data["salario_base"]))
        if ai_data.get("tem_he50") is True:
            parsed["tem_he50"] = True
            if ai_data.get("qtd_he50_mensal"):
                parsed["qtd_he50_mensal"] = float(D(ai_data["qtd_he50_mensal"]))
            parsed["reflexos_he50"] = bool(ai_data.get("reflexos_he50", True))
        if ai_data.get("tem_intrajornada") is True:
            parsed["tem_intrajornada"] = True
            if ai_data.get("teto_intrajornada"):
                parsed["teto_intrajornada"] = float(D(ai_data["teto_intrajornada"]))
            parsed["reflexos_intrajornada"] = bool(ai_data.get("reflexos_intrajornada", False))
        if ai_data.get("tem_descontos_indevidos") is True:
            parsed["tem_descontos_indevidos"] = True
            if ai_data.get("valor_descontos_indevidos"):
                parsed["valor_descontos_indevidos"] = float(D(ai_data["valor_descontos_indevidos"]))
            if ai_data.get("teto_descontos_indevidos"):
                parsed["teto_descontos_indevidos"] = float(D(ai_data["teto_descontos_indevidos"]))
        if ai_data.get("honorarios_sucumbencia_pct") and parsed.get("honorarios_sucumbencia_pct", 0) == 0:
            parsed["honorarios_sucumbencia_pct"] = float(D(ai_data["honorarios_sucumbencia_pct"]))
        if ai_data.get("custas_processuais") and parsed.get("custas_processuais", 0) == 0:
            parsed["custas_processuais"] = float(D(ai_data["custas_processuais"]))
        if ai_data.get("resumo_decisao_judicial"):
            parsed["_ai_decision_summary"] = str(ai_data["resumo_decisao_judicial"])

    contracts = parsed.get("_contracts", [])
    if selected_contract_idx is not None and 0 <= selected_contract_idx < len(contracts):
        chosen = contracts[selected_contract_idx]
        parsed["_selected_contract_id"] = chosen["contrato_id"]
        if parse_br_date_optional(chosen["admissao"]):
            parsed["admissao"] = parse_br_date_optional(chosen["admissao"])
        if parse_br_date_optional(chosen["demissao"]):
            parsed["demissao"] = parse_br_date_optional(chosen["demissao"])
        if chosen["salario"]:
            parsed["salario_base"] = float(chosen["salario"])

    if confirm_human:
        parsed["_human_confirmed_ambiguities"] = True

    st.session_state["current_calc_uid"] = uuid.uuid4().hex[:12]
    st.session_state["has_analyzed"] = True
    _recompute_from_current_params(parsed, save_to_history=True)


# ==============================================================================
# INTERFACE ESTILO SUITE CALCMACHINE (CONFIABILIDADE DETERMINÍSTICA + .PJC + JSON)
# ==============================================================================
st.set_page_config(
    page_title="CalcMachine Suite • Cálculos Trabalhistas (.PJC + JSON + PJe-Calc)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inicialização segura do estado da sessão no nível raiz (sem indentação indevida)
if "params" not in st.session_state:
    st.session_state["params"] = _default_empty_params()
if "calc_result" not in st.session_state:
    st.session_state["calc_result"] = None
if "audit_report" not in st.session_state:
    st.session_state["audit_report"] = None
if "has_analyzed" not in st.session_state:
    st.session_state["has_analyzed"] = False
if "sentence_pasted_text" not in st.session_state:
    st.session_state["sentence_pasted_text"] = ""
if "lawyer_notes" not in st.session_state:
    st.session_state["lawyer_notes"] = ""
if "data_base_liq" not in st.session_state:
    st.session_state["data_base_liq"] = date.today()
if "calc_json_text" not in st.session_state:
    st.session_state["calc_json_text"] = ""
if "current_calc_uid" not in st.session_state:
    st.session_state["current_calc_uid"] = ""

st.markdown(
    """
    <style>
    .cm-hero {
        padding: 20px 26px;
        border-radius: 14px;
        background: linear-gradient(135deg, #090D16 0%, #111827 55%, #0F2942 100%);
        color: #F8FAFC;
        border: 1px solid #1E293B;
        margin-bottom: 18px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.25);
    }
    .cm-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        background: rgba(56, 189, 248, 0.14);
        border: 1px solid rgba(56, 189, 248, 0.4);
        color: #38BDF8;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        margin-bottom: 8px;
    }
    .cm-step-pill {
        display: inline-block;
        padding: 3px 9px;
        border-radius: 6px;
        background: #0284C7;
        color: #FFFFFF;
        font-weight: 700;
        font-size: 0.75rem;
        margin-right: 8px;
    }
    </style>
    <div class="cm-hero">
        <div class="cm-badge">CalcMachine Suite • Motor Determinístico &amp; Auditoria Tripla</div>
        <h2 style="margin: 0 0 6px 0; font-size: 1.6rem; font-weight: 800;">⚡ Sentenças e Autos em Arquivos Prontos (.PJC + JSON + PJe-Calc PDF)</h2>
        <p style="margin: 0; font-size: 0.92rem; color: #CBD5E1;">
            Extração integral com rastreabilidade documental • Cálculo matemático determinístico (Decimal) • Saída dupla editável em tempo real • Histórico versionado
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Sidebar estilo Suíte de Aplicações
with st.sidebar:
    st.markdown("### ⚡ CalcMachine Suite")
    st.caption("Liquidação Trabalhista & Interoperabilidade PJe-Calc")
    nav_page = st.radio(
        "Navegação",
        [
            "⚡ Novo Cálculo (.PJC + JSON)",
            "🗂️ Histórico & Versões Salvas",
            "🛡️ Auditoria & Testes de Regressão",
        ],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown("**Advogado(a) Assinante (Minuta)**")
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
    parte_rep_sb = st.selectbox(
        "Parte Representada",
        ["Reclamante", "Reclamado"],
        index=0 if st.session_state["params"].get("parte_representada", "Reclamante") == "Reclamante" else 1,
        key="sb_parte_rep",
    )
    st.session_state["params"]["advogado_assinante"] = adv_nome_sb
    st.session_state["params"]["oab_advogado"] = adv_oab_sb
    st.session_state["params"]["parte_representada"] = parte_rep_sb

    st.divider()
    st.markdown("**Motor & Regras Vigentes**")
    st.caption(f"• Precisão: `Decimal` (28 dígitos)\n• Tabelas oficiais até: `{ULTIMA_COMPETENCIA_TABELA_OFICIAL}`\n• Atualização: ADC 58 STF / Lei 14.905/24")

if nav_page == "⚡ Novo Cálculo (.PJC + JSON)":
    # PASSO 01: COLE A SENTENÇA OU ANEXE OS AUTOS
    with st.container(border=True):
        st.markdown('<div><span class="cm-step-pill">PASSO 01</span><strong>Cole a Sentença / Decisões ou Anexe os Autos do Processo</strong></div>', unsafe_allow_html=True)
        st.write("")
        sentence_text_input = st.text_area(
            "Texto da Sentença, Embargos de Declaração, Acórdão ou Dispositivo Judicial",
            value=st.session_state["sentence_pasted_text"],
            height=185,
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

        if st.button("⚡ PASSO 02 — Analisar Documentos e Gerar Arquivos (.PJC + JSON + Relatório)", type="primary", use_container_width=True):
            json_uploaded = [f for f in (uploaded_files or []) if f.name.lower().endswith(".json")]
            if json_uploaded:
                raw_j = json_uploaded[0].read().decode("utf-8", errors="ignore")
                new_p, err_j = apply_editable_json_to_params(raw_j, _default_empty_params())
                if err_j:
                    st.error(err_j)
                else:
                    if adv_nome_sb:
                        new_p["advogado_assinante"] = adv_nome_sb
                    if adv_oab_sb:
                        new_p["oab_advogado"] = adv_oab_sb
                    st.session_state["current_calc_uid"] = uuid.uuid4().hex[:12]
                    st.session_state["has_analyzed"] = True
                    _recompute_from_current_params(new_p, save_to_history=True)
            elif not uploaded_files and not (sentence_text_input or "").strip() and not (lawyer_notes_input or "").strip():
                st.error("Cole o texto da sentença ou anexe ao menos um documento do processo.")
            else:
                st.session_state["sentence_pasted_text"] = sentence_text_input
                st.session_state["lawyer_notes"] = lawyer_notes_input
                st.session_state["data_base_liq"] = data_base_input
                with st.spinner("Lendo todas as páginas, cruzando decisões judiciais e executando o motor determinístico..."):
                    _execute_full_pipeline(
                        uploaded_files=uploaded_files,
                        sentence_pasted_text=sentence_text_input,
                        lawyer_notes=lawyer_notes_input,
                        data_base_liq=data_base_input,
                        api_key=default_api_key,
                    )

    # PASSO 03: SAÍDA DUPLA (.PJC + JSON EDITÁVEL) E RESULTADOS DETALHADOS
    if st.session_state["has_analyzed"] and st.session_state["calc_result"] is not None:
        p = st.session_state["params"]
        res = st.session_state["calc_result"]
        aud = st.session_state["audit_report"]
        resumo = res["resumo"]
        meta = res["meta"]
        ing = p.get("_ingestion", {})

        st.write("")
        if aud["status"] == "APROVADO TECNICAMENTE":
            st.success(
                f"✅ **CÁLCULO AUDITADO E APROVADO** ({ing.get('total_pages_processed', 0)}/{ing.get('total_pages_detected', 0)} páginas analisadas • Auditorias Documental, Jurídica, Matemática e de Exportação aprovadas)."
            )
        elif aud["status"] == "REPROVADO":
            st.error("🚫 **BLOQUEIO DE SEGURANÇA**: Existe inconsistência crítica nos parâmetros ou no cálculo. Revise na aba de Conferência ou no JSON.")
        else:
            st.warning(
                f"⚠️ **AGUARDANDO CONFERÊNCIA HUMANA** ({ing.get('total_pages_processed', 0)}/{ing.get('total_pages_detected', 0)} páginas analisadas): Revise os campos extraídos abaixo e confirme para liberar a emissão definitiva."
            )

        # Métricas principais
        r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
        r1c1.metric("Bruto do Reclamante", f"R$ {fmt_brl(resumo['bruto_devido_reclamante'])}")
        r1c2.metric("INSS Segurado", f"R$ {fmt_brl(resumo['inss_empregado'])}")
        r1c3.metric("IRRF (RRA)", f"R$ {fmt_brl(resumo['irpf_reclamante'])}")
        r1c4.metric("Líquido do Reclamante", f"R$ {fmt_brl(resumo['liquido_devido_reclamante'])}")
        r1c5.metric("FGTS", f"R$ {fmt_brl(resumo['total_fgts_bruto'])}")

        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        r2c1.metric("Honorários Advocatícios", f"R$ {fmt_brl(resumo['honorarios_valor'])}")
        r2c2.metric("INSS Patronal", f"R$ {fmt_brl(resumo['inss_empregador'])}")
        r2c3.metric("Custas Processuais", f"R$ {fmt_brl(resumo['custas_devidas_execucao'])}")
        r2c4.metric("Total Devido pelo Reclamado", f"R$ {fmt_brl(resumo['total_devido_reclamado'])}")

        with st.container(border=True):
            st.markdown('<div><span class="cm-step-pill">PASSO 03</span><strong>Saída Dupla Instantânea (.PJC + JSON) e Documentos da Liquidação</strong></div>', unsafe_allow_html=True)
            st.write("")
            pdf_report_bytes = build_pjecalc_pdf(res, aud)
            pjc_file_bytes = build_pjc_xml(res)
            json_file_bytes = st.session_state["calc_json_text"].encode("utf-8")
            minuta_txt = build_minuta_peticao(p, res, aud, st.session_state["lawyer_notes"])
            minuta_pdf_bytes = build_minuta_pdf(minuta_txt, str(meta["processo"]))
            minuta_docx_bytes = build_minuta_docx(minuta_txt, str(meta["processo"]))
            proc_slug = re.sub(r"[^0-9A-Za-z\-]", "_", str(meta["processo"]))

            dl1, dl2, dl3, dl4, dl5 = st.columns(5)
            dl1.download_button(
                label="💾 Baixar Arquivo .PJC",
                data=pjc_file_bytes,
                file_name=f"Calculo_{proc_slug}.pjc",
                mime="application/xml",
                use_container_width=True,
                type="primary",
            )
            dl2.download_button(
                label="🧩 Baixar Dados .JSON",
                data=json_file_bytes,
                file_name=f"Calculo_{proc_slug}.json",
                mime="application/json",
                use_container_width=True,
                type="primary",
            )
            dl3.download_button(
                label="📄 Relatório PJe-Calc (PDF)",
                data=pdf_report_bytes,
                file_name=f"Relatorio_Liquidacao_{proc_slug}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            dl4.download_button(
                label="📝 Minuta Petição (Word)",
                data=minuta_docx_bytes,
                file_name=f"Minuta_Peticao_{proc_slug}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            dl5.download_button(
                label="📑 Minuta Petição (PDF)",
                data=minuta_pdf_bytes,
                file_name=f"Minuta_Peticao_{proc_slug}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.caption(
                "Nota Técnica: O arquivo .pjc segue estrutura XML de intercâmbio gerada por motor independente (sem chancela digital oficial do CSJT); confira os valores após importação no PJe-Calc Cidadão."
            )

        tab_visual_edit, tab_json_edit, tab_memoria, tab_minuta = st.tabs(
            [
                "🛠️ Conferência & Ajuste Visual",
                "🧩 Editor JSON em Tempo Real (.PJC)",
                "🔍 Decisão & Memória Mês a Mês",
                "📜 Minuta da Petição (Copiar)",
            ]
        )

        with tab_visual_edit:
            contracts_found = p.get("_contracts", [])
            if len(contracts_found) > 1:
                ctr_labels = [
                    f"{c['contrato_id']} — Admissão: {c['admissao']} | Demissão: {c['demissao']} | Salário: {c['salario'] if c['salario'] else 'Não localizado'} ({c['origem_doc']}, pág. {c['pagina']})"
                    for c in contracts_found
                ]
                chosen_ctr_idx = st.selectbox(
                    "Múltiplos contratos encontrados na CTPS/autos — selecione o vínculo objeto da ação:",
                    range(len(ctr_labels)),
                    format_func=lambda i: ctr_labels[i],
                    key="sel_contrato_conferencia",
                )
                if st.button("📥 Aplicar vínculo selecionado"):
                    chosen_c = contracts_found[chosen_ctr_idx]
                    p["_selected_contract_id"] = chosen_c["contrato_id"]
                    if parse_br_date_optional(chosen_c["admissao"]):
                        p["admissao"] = parse_br_date_optional(chosen_c["admissao"])
                    if parse_br_date_optional(chosen_c["demissao"]):
                        p["demissao"] = parse_br_date_optional(chosen_c["demissao"])
                    if chosen_c["salario"]:
                        p["salario_base"] = float(chosen_c["salario"])
                    _recompute_from_current_params(p, save_to_history=True)
                    st.rerun()

            with st.form("form_conferencia_calculo"):
                fc1, fc2, fc3 = st.columns(3)
                edit_proc = fc1.text_input("Processo CNJ", value=p.get("processo", ""))
                edit_rec = fc2.text_input("Reclamante", value=p.get("reclamante", ""))
                edit_recd = fc3.text_input("Reclamado", value=p.get("reclamado", ""))

                fd1, fd2, fd3, fd4 = st.columns(4)
                edit_adm = fd1.date_input("Data de Admissão", value=parse_br_date(p.get("admissao"), date(2024, 1, 1)), format="DD/MM/YYYY")
                edit_dem = fd2.date_input("Data de Demissão", value=parse_br_date(p.get("demissao"), date(2025, 1, 1)), format="DD/MM/YYYY")
                edit_aju = fd3.date_input("Data de Ajuizamento", value=parse_br_date(p.get("ajuizamento"), date.today()), format="DD/MM/YYYY")
                edit_sal = fd4.number_input("Salário Base (R$)", min_value=0.0, value=float(p.get("salario_base", 0.0)), step=50.0, format="%.2f")

                v1c1, v1c2, v1c3 = st.columns([2, 1, 1])
                edit_he50 = v1c1.checkbox("Horas Extras 50% (Deferida)", value=bool(p.get("tem_he50", False)))
                edit_qtd_he50 = v1c2.number_input("Qtd. Mensal HE 50% (h)", min_value=0.0, value=float(p.get("qtd_he50_mensal", 0.0)), step=1.0)
                edit_ref_he50 = v1c3.checkbox("Reflexos HE 50%", value=bool(p.get("reflexos_he50", False)))

                v2c1, v2c2, v2c3 = st.columns([2, 1, 1])
                edit_intra = v2c1.checkbox("Intervalo Intrajornada (1h/dia útil)", value=bool(p.get("tem_intrajornada", False)))
                edit_teto_intra = v2c2.number_input("Teto Intrajornada (R$)", min_value=0.0, value=float(p.get("teto_intrajornada", 0.0)), step=50.0, format="%.2f")
                edit_ref_intra = v2c3.checkbox("Reflexos Intrajornada", value=bool(p.get("reflexos_intrajornada", False)))

                v3c1, v3c2, v3c3 = st.columns([2, 1, 1])
                edit_desc = v3c1.checkbox("Devolução de Descontos Indevidos", value=bool(p.get("tem_descontos_indevidos", False)))
                edit_val_desc = v3c2.number_input("Valor Descontos (R$)", min_value=0.0, value=float(p.get("valor_descontos_indevidos", 0.0)), step=50.0, format="%.2f")
                edit_teto_desc = v3c3.number_input("Teto Descontos (R$)", min_value=0.0, value=float(p.get("teto_descontos_indevidos", 0.0)), step=50.0, format="%.2f")

                v4c1, v4c2, v4c3 = st.columns(3)
                edit_fgts = v4c1.checkbox("FGTS (+40% se sem justa causa)", value=bool(p.get("reflexos_em_fgts", False)))
                edit_m477 = v4c2.checkbox("Multa Art. 477 CLT", value=bool(p.get("multa_477", False)))
                edit_m467 = v4c3.checkbox("Multa Art. 467 CLT", value=bool(p.get("multa_467", False)))

                vp_atual = p.get("verbas_personalizadas", [])
                vp0 = vp_atual[0] if vp_atual else {"nome": "", "valor_principal": 0.0, "salarial": False}
                vpc1, vpc2, vpc3 = st.columns([2, 1, 1])
                edit_vp_nome = vpc1.text_input("Verba Adicional da Sentença", value=vp0.get("nome", ""))
                edit_vp_val = vpc2.number_input("Valor Principal (R$)", min_value=0.0, value=float(vp0.get("valor_principal", 0.0)), step=50.0, format="%.2f")
                edit_vp_sal = vpc3.checkbox("Natureza Salarial", value=bool(vp0.get("salarial", False)))

                hc1, hc2, hc3, hc4 = st.columns(4)
                edit_hon = hc1.number_input("Honorários (%)", min_value=0.0, max_value=30.0, value=float(p.get("honorarios_sucumbencia_pct", 0.0)), step=1.0)
                edit_custas = hc2.number_input("Custas Processuais (R$)", min_value=0.0, value=float(p.get("custas_processuais", 0.0)), step=10.0, format="%.2f")
                edit_custas_rec = hc3.checkbox("Custas já recolhidas", value=bool(p.get("custas_ja_recolhidas", False)))
                edit_abat = hc4.number_input("Abatimentos (R$)", min_value=0.0, value=float(p.get("valor_pago_abatimento", 0.0)), step=50.0, format="%.2f")

                edit_confirm = st.checkbox(
                    "Confirmo que conferi os parâmetros e verbas acima para emissão definitiva.",
                    value=bool(p.get("_human_confirmed_ambiguities", False)),
                )

                if st.form_submit_button("🔄 Salvar Ajustes e Regenerar .PJC + JSON", type="primary", use_container_width=True):
                    p["processo"] = edit_proc.strip()
                    p["reclamante"] = edit_rec.strip()
                    p["reclamado"] = edit_recd.strip()
                    p["admissao"] = edit_adm
                    p["demissao"] = edit_dem
                    p["ajuizamento"] = edit_aju
                    p["salario_base"] = edit_sal
                    p["tem_he50"] = edit_he50
                    p["qtd_he50_mensal"] = edit_qtd_he50
                    p["reflexos_he50"] = edit_ref_he50
                    p["tem_intrajornada"] = edit_intra
                    p["qtd_intrajornada_mensal"] = 1.0 if edit_intra else 0.0
                    p["teto_intrajornada"] = edit_teto_intra
                    p["reflexos_intrajornada"] = edit_ref_intra
                    p["tem_descontos_indevidos"] = edit_desc
                    p["valor_descontos_indevidos"] = edit_val_desc
                    p["teto_descontos_indevidos"] = edit_teto_desc
                    p["reflexos_em_fgts"] = edit_fgts
                    p["multa_477"] = edit_m477
                    p["multa_467"] = edit_m467
                    p["verbas_personalizadas"] = (
                        [{"nome": edit_vp_nome.strip(), "valor_principal": edit_vp_val, "salarial": edit_vp_sal}]
                        if edit_vp_nome.strip() and edit_vp_val > 0
                        else []
                    )
                    p["honorarios_sucumbencia_pct"] = edit_hon
                    p["custas_processuais"] = edit_custas
                    p["custas_ja_recolhidas"] = edit_custas_rec
                    p["valor_pago_abatimento"] = edit_abat
                    p["_human_confirmed_ambiguities"] = edit_confirm
                    _recompute_from_current_params(p, save_to_history=True)
                    st.rerun()

        with tab_json_edit:
            st.caption(
                "Edite diretamente qualquer parâmetro, data, salário, verba ou teto judicial no JSON abaixo e clique em **Aplicar JSON e Regenerar .PJC**."
            )
            edited_json_str = st.text_area(
                "JSON Estruturado do Cálculo",
                value=st.session_state["calc_json_text"],
                height=420,
                key="json_live_editor",
            )
            if st.button("⚡ Aplicar JSON e Regenerar .PJC + PDF + Minuta", type="primary", use_container_width=True):
                updated_p, err_msg = apply_editable_json_to_params(edited_json_str, p)
                if err_msg:
                    st.error(err_msg)
                else:
                    _recompute_from_current_params(updated_p, save_to_history=True)
                    st.success("JSON aplicado e nova versão gravada no histórico!")
                    st.rerun()

        with tab_memoria:
            st.info(generate_judicial_decision_summary(p))
            if res["verbas"]:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "Verba": v["nome"],
                                "Origem": v.get("verba_origem", ""),
                                "Principal (R$)": fmt_brl(v["valor"]),
                                "Corrigido (R$)": fmt_brl(v["valor_corrigido"]),
                                "Juros (R$)": fmt_brl(v["juros"]),
                                "Total (R$)": fmt_brl(v["total_verba"]),
                            }
                            for v in res["verbas"]
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
                for v in res["verbas"]:
                    ocs = v.get("ocorrencias", [])
                    if ocs:
                        st.markdown(f"**{v['nome']}** (`{v['formula_pje']}`)")
                        st.dataframe(
                            pd.DataFrame(
                                [
                                    {
                                        "Competência": oc.get("competencia", ""),
                                        "Base (R$)": fmt_brl(oc["base"]),
                                        "Divisor": fmt_num(oc["divisor"], 2),
                                        "Mult.": fmt_num(oc["multiplicador"], 4),
                                        "Qtd.": fmt_num(oc["quantidade"], 4),
                                        "Principal (R$)": fmt_brl(oc["devido"]),
                                        "Índice": fmt_num(oc["indice_correcao"], 6),
                                        "Corrigido (R$)": fmt_brl(oc["valor_corrigido"]),
                                        "Juros (R$)": fmt_brl(oc["juros"]),
                                    }
                                    for oc in ocs
                                ]
                            ),
                            use_container_width=True,
                            hide_index=True,
                        )

            todas_pendencias = (
                aud["auditoria_documental"]["achados"]
                + aud["auditoria_juridica"]["achados"]
                + aud["auditoria_matematica"]["achados"]
                + aud.get("auditoria_exportacao", {}).get("achados", [])
            )
            if todas_pendencias:
                st.markdown("**Achados da Auditoria:**")
                for pend in todas_pendencias:
                    st.warning(f"• {pend}")

        with tab_minuta:
            st.text_area(
                "Minuta da Petição de Apresentação de Cálculos (pronta para copiar)",
                value=minuta_txt,
                height=380,
                key="minuta_copia_box",
            )

elif nav_page == "🗂️ Histórico & Versões Salvas":
    st.markdown("### 🗂️ Histórico Completo de Cálculos e Versões")
    st.caption("Cada processamento ou ajuste no JSON gera uma versão rastreável. Selecione qualquer registro para reabrir, editar ou baixar novamente.")
    hist_items = load_calculation_history(limit=50)
    if not hist_items:
        st.info("Nenhum cálculo salvo ainda. Processe uma sentença na tela 'Novo Cálculo' para iniciar o histórico.")
    else:
        df_hist = pd.DataFrame(
            [
                {
                    "ID": h["id"],
                    "Cálculo": h["calc_uid"],
                    "Versão": f"v{h['versao']}",
                    "Data/Hora": h["created_at"],
                    "Processo": h["processo"],
                    "Reclamante": h["reclamante"],
                    "Líquido Rec. (R$)": h["liquido_reclamante"],
                    "Total Execução (R$)": h["total_execucao"],
                    "Auditoria": h["status_auditoria"],
                }
                for h in hist_items
            ]
        )
        st.dataframe(df_hist, use_container_width=True, hide_index=True)

        sel_hist_idx = st.selectbox(
            "Selecione uma versão para carregar:",
            range(len(hist_items)),
            format_func=lambda i: f"#{hist_items[i]['id']} • {hist_items[i]['processo']} (v{hist_items[i]['versao']}) — {hist_items[i]['reclamante']} — {hist_items[i]['created_at']}",
        )
        if st.button("📂 Carregar Versão no Editor e Regenerar Arquivos", type="primary"):
            chosen_h = hist_items[sel_hist_idx]
            restored_p, err_r = apply_editable_json_to_params(chosen_h["payload_json"], _default_empty_params())
            if err_r:
                st.error(err_r)
            else:
                st.session_state["current_calc_uid"] = chosen_h["calc_uid"]
                st.session_state["has_analyzed"] = True
                _recompute_from_current_params(restored_p, save_to_history=False)
                st.success(f"Versão v{chosen_h['versao']} do processo {chosen_h['processo']} carregada! Volte para 'Novo Cálculo (.PJC + JSON)' no menu lateral.")

elif nav_page == "🛡️ Auditoria & Testes de Regressão":
    st.markdown("### 🛡️ Suíte Permanente de Testes Automatizados e Auditoria")
    st.caption("Valida precisão decimal, detecção de múltiplos contratos na CTPS, embargos modificativos, teto judicial, segregação de reflexos, bloqueio de representação invertida, custas e proteção contra prompt injection.")
    if st.button("▶️ Executar Suíte Completa de Testes Agora", type="primary"):
        suite_res = run_automated_regression_suite()
        if suite_res["passed"] == suite_res["total"]:
            st.success(f"✅ Todos os {suite_res['passed']}/{suite_res['total']} testes críticos passaram!")
        else:
            st.error(f"❌ Falha em testes: {suite_res['passed']}/{suite_res['total']} aprovados.")
        st.dataframe(pd.DataFrame(suite_res["details"]), use_container_width=True, hide_index=True)
