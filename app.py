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
# 5. EXTRAÇÃO ASSISTIDA POR GEMINI (ANÁLISE COMPLETA DO DOSSIÊ TRABALHISTA)
# ==============================================================================
class PaidItem(BaseModel):
    verba: str
    valor: float


class IrregularityFinding(BaseModel):
    categoria: str
    documento_origem: str
    descricao: str
    fundamento_juridico: str
    impacto_financeiro: str


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
    trct_paid_items: List[PaidItem] = []
    irregularities_found: List[IrregularityFinding] = []
    medical_leave_summary: Optional[str] = None
    stability_alert: Optional[str] = None
    general_legal_opinion: Optional[str] = None
    ambiguities: List[str] = []


def get_gemini_api_key() -> Optional[str]:
    """Obtém a chave via st.secrets ou variável de ambiente sem nunca expor no código."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY")


def analyze_full_dossier_with_gemini(files_payload: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Envia TODOS os arquivos enviados pelo advogado (TRCT, contracheques, atestados médicos,
    laudos de afastamento/INSS, ASO, ponto, FGTS, CCT) em uma única análise cruzada do Gemini.
    Identifica brechas, estabilidade, descontos indevidos e extrai dados objetivos.
    NUNCA calcula valores financeiros finais sem o motor Decimal e exige revisão humana.
    """
    api_key = get_gemini_api_key()
    if not api_key or not GENAI_AVAILABLE:
        raise RuntimeError("GEMINI_API_KEY não configurada ou pacote google-genai ausente.")

    client = genai.Client(api_key=api_key)
    contents: List[Any] = [
        (
            "Você é um auditor jurídico trabalhista sênior. Analise TODOS os documentos anexados em conjunto "
            "(TRCT, contracheques, cartões de ponto, extratos FGTS, atestados médicos, laudos de afastamento/INSS, "
            "ASO admissional/demissional, CAT, advertências ou CCT). "
            "1) Extraia os dados objetivos do contrato (partes, datas, modalidade de rescisão, salário-base, comissões, "
            "divisor e valores já pagos no TRCT). "
            "2) Identifique TODAS as irregularidades, nulidades e brechas para pedido de restituição ou indenização "
            "(ex.: dispensa durante estabilidade acidentária B91 ou gestante, atestado médico válido descontado como falta "
            "no contracheque/TRCT, demissão com ASO inapto, diferenças de verbas rescisórias, ausência de FGTS, horas extras/intervalo). "
            "3) Não invente datas ou valores ausentes. Aponte qualquer dúvida em 'ambiguities'."
        )
    ]

    for item in files_payload:
        contents.append(f"--- ARQUIVO ANEXADO: {item['name']} ---")
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


def merge_extracted_documents(results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Consolida dados lidos de múltiplos documentos (TRCT + Contracheques + Ponto + CTPS + FGTS + CCT)."""
    merged: Dict[str, Any] = {
        "paid_items_breakdown": [],
        "ambiguities": [],
        "sources_count": len(results),
    }
    scalar_fields = [
        "process_number", "worker_name", "employer_name", "admission_date",
        "dismissal_date", "termination_type", "base_salary", "commissions",
        "insalubrity_or_danger_pct", "monthly_divisor", "weekly_overtime_hours_detected",
        "suppressed_break_minutes_daily", "weekly_night_hours_detected",
        "total_paid_trct", "fgts_balance_detected", "cct_rules_summary",
    ]
    for item in results:
        for f in scalar_fields:
            val = item.get(f)
            if val is not None and val != "" and val != 0:
                merged[f] = val
        if item.get("paid_items_breakdown"):
            merged["paid_items_breakdown"].extend(item["paid_items_breakdown"])
        if item.get("ambiguities"):
            merged["ambiguities"].extend(item["ambiguities"])

    if not merged.get("total_paid_trct") and merged["paid_items_breakdown"]:
        merged["total_paid_trct"] = round(
            sum(float(x.get("valor_pago") or 0) for x in merged["paid_items_breakdown"]), 2
        )
    return merged


def parse_safe_date(date_str: Optional[str], fallback: date) -> date:
    if not date_str:
        return fallback
    try:
        return datetime.strptime(str(date_str).strip()[:10], "%Y-%m-%d").date()
    except Exception:
        return fallback
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
# 8. MENU LATERAL PROFISSIONAL
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚖️ ED Cálculos Trabalhistas")
    st.caption("SaaS Jurídico • Auditoria Rescisória & Avaliação Preliminar")
    st.divider()
    menu = st.radio(
        "Navegação",
        [
            "Avaliação Preliminar & Auditoria",
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
if menu == "Avaliação Preliminar & Auditoria":
    st.title("Avaliação Preliminar & Auditoria Completa do Dossiê")
    st.markdown(
        "Envie **quaisquer documentos do caso** (TRCT, contracheques, cartões de ponto, extrato FGTS, **atestados médicos, "
        "laudos de afastamento/INSS, ASO**, contrato ou CCT). A IA faz a **leitura cruzada para encontrar brechas e irregularidades** "
        "e o **motor decimal exato** apura cada centavo a restituir."
    )

    with st.container(border=True):
        st.subheader("1. Central Universal de Upload (Envie Todos os Arquivos do Cliente)")
        st.caption(
            "Aceita múltiplos arquivos simultâneos em PDF, Imagem (JPG/PNG), Excel (.xlsx) e CSV: "
            "TRCT, Contracheques, Cartão de Ponto, Atestados Médicos, Laudos/Benefícios INSS, ASO, Extrato FGTS e CCT."
        )

        uploaded_all = st.file_uploader(
            "Arraste e solte aqui todos os arquivos do cliente (sem limite de categoria)",
            type=["pdf", "png", "jpg", "jpeg", "csv", "xlsx", "txt"],
            accept_multiple_files=True,
            key="universal_dossier_uploader",
        )

        if uploaded_all:
            st.write(f"**{len(uploaded_all)} arquivo(s) pronto(s) para análise:** " + ", ".join(f"`{f.name}`" for f in uploaded_all))
            if st.button("Ler Todos os Arquivos e Auditar Irregularidades (IA + Motor Decimal)", type="primary"):
                gemini_files = []
                timecard_dfs = []
                for f in uploaded_all:
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

                if gemini_files:
                    try:
                        with st.spinner("Auditando documentos, atestados, laudos, contracheques e TRCT..."):
                            st.session_state.extracted_ai_data = analyze_full_dossier_with_gemini(gemini_files)
                            st.session_state.extraction_confirmed = False
                        st.success("Leitura completa concluída! Confira o diagnóstico de irregularidades e revise os dados abaixo.")
                    except Exception as exc:
                        st.warning(f"Análise via Gemini indisponível ({exc}). Verifique `GEMINI_API_KEY` em Secrets.")

    tc_pre = st.session_state.timecard_result
    ext_pre = st.session_state.extracted_ai_data or {}

    # Painel de Diagnóstico Jurídico e Irregularidades Encontradas pela IA
    if ext_pre:
        with st.container(border=True):
            st.subheader("Diagnóstico da Auditoria Documental (Brechas e Irregularidades Detectadas)")
            if ext_pre.get("stability_alert"):
                st.error(f"**ALERTA DE ESTABILIDADE / NULIDADE DA DISPENSA:** {ext_pre['stability_alert']}")
            if ext_pre.get("medical_leave_summary"):
                st.warning(f"**Atestados Médicos / Afastamentos / ASO:** {ext_pre['medical_leave_summary']}")
            if ext_pre.get("general_legal_opinion"):
                st.info(f"**Síntese Jurídica do Dossiê:** {ext_pre['general_legal_opinion']}")

            irregs = ext_pre.get("irregularities_found") or []
            if irregs:
                st.markdown("#### Irregularidades e Pontos de Restituição Identificados")
                st.dataframe(pd.DataFrame(irregs), use_container_width=True, hide_index=True)

            if ext_pre.get("ambiguities"):
                st.caption("**Pontos ambíguos ou ilegíveis que exigem atenção do advogado:** " + " | ".join(ext_pre["ambiguities"]))

            st.session_state.extraction_confirmed = st.checkbox(
                "Confirmo que revisei os dados extraídos e o diagnóstico antes de gerar o cálculo final",
                value=st.session_state.extraction_confirmed,
            )

    if tc_pre is not None:
        with st.expander("Ver Espelho de Ponto Apurado Dia a Dia (Tolerância CLT + Redução Ficta)"):
            st.dataframe(tc_pre["table"], use_container_width=True, hide_index=True)

    # Soma automática de itens pagos no TRCT extraídos pela IA
    trct_paid_list = ext_pre.get("trct_paid_items") or []
    trct_auto_sum = round(sum(float(i.get("valor", 0)) for i in trct_paid_list), 2)

    def parse_iso_date(val: Optional[str], fallback: date) -> date:
        if not val:
            return fallback
        try:
            return datetime.strptime(val[:10], "%Y-%m-%d").date()
        except Exception:
            return fallback

    with st.form("preliminary_eval_form"):
        st.subheader("2. Conferência das Condições do Contrato (Preenchidas pelos Arquivos ou Manualmente)")
        p1, p2, p3 = st.columns(3)
        cliente_nome = p1.text_input("Nome do Cliente (Trabalhador) *", value=ext_pre.get("worker_name") or "João da Silva")
        empresa_nome = p2.text_input("Empresa (Futura Reclamada) *", value=ext_pre.get("employer_name") or "Comércio e Serviços Alfa Ltda.")
        cargo_funcao = p3.text_input("Cargo / Função Exercida", value=ext_pre.get("role_title") or "Auxiliar Operacional / Vendedor")

        motivos_opts = ["sem_justa_causa", "rescisao_indireta", "pedido_demissao", "acordo", "justa_causa"]
        motivo_ext = ext_pre.get("termination_type") if ext_pre.get("termination_type") in motivos_opts else "sem_justa_causa"

        d1, d2, d3, d4 = st.columns(4)
        adm_pre = d1.date_input("Data de Admissão *", value=parse_iso_date(ext_pre.get("admission_date"), date(2024, 2, 1)), key="adm_pre")
        dism_pre = d2.date_input("Data de Saída (ou Hoje) *", value=parse_iso_date(ext_pre.get("dismissal_date"), date(2026, 11, 20)), key="dism_pre")
        motivo_saida = d3.selectbox(
            "Situação da Saída *",
            motivos_opts,
            index=motivos_opts.index(motivo_ext),
            format_func=lambda x: {
                "sem_justa_causa": "Dispensa sem justa causa",
                "rescisao_indireta": "Pedido de Rescisão Indireta (falta grave do empregador)",
                "pedido_demissao": "Pedido de demissão (reversão / verbas pendentes)",
                "acordo": "Rescisão por acordo (art. 484-A CLT)",
                "justa_causa": "Dispensa por justa causa (contestação)",
            }[x],
        )
        divisor_pre = d4.number_input("Jornada / Divisor Mensal *", min_value=1, max_value=300, value=int(ext_pre.get("monthly_divisor") or 220))

        st.subheader("3. Remuneração Real vs. O que Foi Pago no TRCT")
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
        adicional_cond = r3.selectbox("Adicional Devido e Não Pago", adicionais_opts, index=adicionais_opts.index(adic_ext))
        valor_ja_pago_trct = r4.number_input(
            "Total Já Pago na Rescisão/TRCT (R$)",
            min_value=0.0,
            value=trct_auto_sum,
            step=100.0,
            help="Preenchido automaticamente se o TRCT foi enviado acima.",
        )

        st.subheader("4. Jornada Habitual e Descontos Indevidos (Atestados / Faltas)")
        meses_contrato = max(1, ((dism_pre.year - adm_pre.year) * 12 + (dism_pre.month - adm_pre.month)))
        st.caption(f"Tempo de contrato apurado: **{meses_contrato} meses** (limitado aos últimos 60 meses pela prescrição quinquenal).")

        j1, j2, j3, j4 = st.columns(4)
        he_semanal_relatada = j1.number_input("Horas Extras Habituais / Semana (50%)", min_value=0.0, value=float(ext_pre.get("weekly_he50_hours") or 0.0), step=0.5)
        domingos_mes = j2.number_input("Domingos/Feriados sem Folga / Mês", min_value=0, max_value=10, value=int(ext_pre.get("monthly_sundays_worked") or 0))
        min_intervalo_suprimido = j3.number_input("Intervalo Almoço Suprimido (min/dia)", min_value=0, max_value=60, value=int(ext_pre.get("suppressed_break_minutes") or 0))
        horas_noturnas_semana = j4.number_input("Horas Noturnas (22h-05h) / Semana", min_value=0.0, value=float(ext_pre.get("weekly_night_hours") or 0.0), step=0.5)

        st.subheader("5. Verbas e Restituições a Incluir no Confronto")
        v1, v2, v3 = st.columns(3)
        ped_saldo = v1.checkbox("Saldo de Salário do Mês da Saída", value=True)
        dias_saldo_pre = v1.number_input("Dias trabalhados no último mês", min_value=0, max_value=31, value=dism_pre.day)
        ped_aviso = v2.checkbox("Aviso-Prévio Indenizado + Projeção", value=True)
        ped_13_ferias = v2.checkbox("13º Proporcional e Férias + 1/3", value=True)
        ped_fgts = v3.checkbox("Diferenças de FGTS (8%) + Multa Rescisória", value=True)
        restituicao_atestados = v3.number_input("Restituição de Descontos Indevidos / Atestados (R$)", min_value=0.0, value=0.0, step=50.0)

        avaliar_btn = st.form_submit_button("Conferir Divergências e Calcular Restituição", type="primary")

    if avaliar_btn:
        if ext_pre and not st.session_state.extraction_confirmed:
            st.error("Marque a caixa de confirmação da revisão documental antes de gerar o cálculo final.")
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
                "process_number": f"AUDITORIA-{cliente_nome.upper().replace(' ', '_')[:18]}",
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
                "apply_inss": False,
                "apply_irrf": False,
                "apply_interest": False,
                "interest_pct": 0.0,
                "correction_pct": 0.0,
            })

            # Soma restituição de descontos indevidos / atestados médicos
            if restituicao_atestados > 0:
                rest_dec = d_money(restituicao_atestados)
                res_pre["verbas"].append({
                    "Código": "VER-08",
                    "Verba": "Restituição de Descontos Indevidos (Atestados Médicos / Faltas)",
                    "Período": f"{adm_pre.strftime('%d/%m/%Y')} a {dism_pre.strftime('%d/%m/%Y')}",
                    "Base de Cálculo": rest_dec,
                    "Fórmula": "Devolução integral de descontos indevidos (Art. 462 CLT)",
                    "Quantidade": "1",
                    "Percentual": "100%",
                    "Valor Original": rest_dec,
                    "Reflexos": Decimal("0.00"),
                    "Correção": Decimal("0.00"),
                    "Juros": Decimal("0.00"),
                    "Total Apurado": rest_dec,
                    "Tributável INSS": False,
                })
                res_pre["bruto_atualizado"] += rest_dec
                res_pre["liquido_devido"] += rest_dec

            st.session_state.calculations_history.insert(0, res_pre)
            st.success(f"Auditoria concluída! Código: **{res_pre['audit_id']}**")

            k1, k2, k3 = st.columns(3)
            k1.metric("Total Realmente Devido", fmt_brl(res_pre["bruto_atualizado"]))
            k2.metric("(-) Já Pago no TRCT", fmt_brl(res_pre["abatimentos"]))
            k3.metric("Diferença / Restituição a Pedir", fmt_brl(res_pre["liquido_devido"]))

            if res_pre["liquido_devido"] > Decimal("0.00"):
                st.warning(
                    f"**DIVERGÊNCIA ENCONTRADA:** Há uma diferença a menor de **{fmt_brl(res_pre['liquido_devido'])}** "
                    "passível de cobrança / restituição na Justiça do Trabalho."
                )
            else:
                st.success("**CONFERÊNCIA OK:** Os valores pagos no TRCT cobrem o total apurado para os parâmetros informados.")

            df_pre = pd.DataFrame(res_pre["verbas"])
            for col in ["Base de Cálculo", "Valor Original", "Reflexos", "Correção", "Juros", "Total Apurado"]:
                df_pre[col] = df_pre[col].apply(fmt_brl)
            st.dataframe(
                df_pre[["Código", "Verba", "Quantidade", "Percentual", "Base de Cálculo", "Fórmula", "Valor Original", "Reflexos", "Total Apurado"]],
                use_container_width=True,
                hide_index=True,
            )

            excel_pre = io.BytesIO()
            with pd.ExcelWriter(excel_pre, engine="openpyxl") as writer:
                df_pre.to_excel(writer, index=False, sheet_name="Auditoria_Rescisoria")
            st.download_button(
                "Baixar Relatório de Divergências e Valor da Causa (.XLSX)",
                data=excel_pre.getvalue(),
                file_name=f"Auditoria_Rescisoria_{res_pre['audit_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

elif menu == "Meus Cálculos e Relatórios":
    st.title("Meus Cálculos e Relatórios Auditáveis")
    if not st.session_state.calculations_history:
        st.info("Nenhuma auditoria realizada nesta sessão ainda. Use 'Avaliação Preliminar & Auditoria'.")
    else:
        for calc in st.session_state.calculations_history:
            with st.expander(f"{calc['audit_id']} • Gerado em {calc['generated_at']} • Diferença a Pedir: {fmt_brl(calc['liquido_devido'])}", expanded=True):
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
