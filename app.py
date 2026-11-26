# ... existing code ...
ULTIMA_COMPETENCIA_TABELA_OFICIAL = "2026-12"
HISTORY_DB_PATH = "juriscalc_pro_historico.db"
DYNAMIC_TABLES_CACHE_PATH = "juriscalc_tabelas_vigentes.json"

# Tabelas Progressivas Vigentes do INSS Segurado (Portarias Interministeriais MPS/MF)
# Estrutura: ano -> lista de (teto_faixa, aliquota_decimal)
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

# Tabelas Progressivas Mensais do IRRF Vigentes (para RRA Art. 12-A Lei 7.713/88)
# Estrutura: ano -> lista de (limite_faixa, aliquota_decimal, parcela_deduzir_mensal)
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
    """
    Sincroniza automaticamente as tabelas vigentes de INSS, IRRF/RRA e Salário Mínimo:
    - Carrega cache persistido em disco (juriscalc_tabelas_vigentes.json) se existir;
    - Se houver URL remota configurada ou JSON importado, valida, atualiza em memória e salva;
    - Atualiza dinamicamente ULTIMA_COMPETENCIA_TABELA_OFICIAL para que cálculos futuros
      usem imediatamente as novas faixas legais sem alterar o código-fonte.
    """
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
# ... existing code ...
def _calc_inss_segurado_competencia(base_val: Decimal, ano: int) -> Decimal:
    """Calcula a contribuição previdenciária progressiva vigente na competência (EC 103/2019)."""
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
    """Calcula o IRRF sobre Rendimentos Recebidos Acumuladamente (Art. 12-A Lei 7.713/88) com tabela vigente na liquidação."""
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
# ... existing code ...
    st.divider()
    st.markdown("**Agente PJe-Calc & Sincronização Automática**")
    sync_meta = st.session_state.get("sync_status_info", SYNC_STATUS_INFO)
    st.caption(
        f"• **Status:** `{sync_meta['status']}` ({sync_meta['fonte']})\n"
        f"• **Tabelas INSS / IRRF (RRA):** Vigentes até `{sync_meta['ultima_competencia']}`\n"
        f"• **Índices BACEN (SGS):** Atualização automática em tempo real (SELIC, IPCA, IPCA-E, TR)"
    )
    with st.expander("Sincronização Automática & PJe-Calc", expanded=False):
        pjc_ep_sb = st.text_input(
            "URL PJe-Calc Cidadão Local",
            value=st.session_state.get("pjecalc_endpoint", "http://localhost:9257/pjecalc"),
            help="Quando o PJe-Calc Cidadão estiver aberto na máquina, o agente submete o cálculo diretamente para ele.",
            key="sb_pjc_endpoint",
        )
        st.session_state["pjecalc_endpoint"] = pjc_ep_sb

        feed_url_sb = st.text_input(
            "URL Feed de Tabelas Legais (JSON)",
            value=st.session_state.get("tables_feed_url", ""),
            placeholder="https://.../tabelas_pjecalc.json",
            help="Opcional: URL JSON remoto para atualizar automaticamente futuras Portarias de INSS/IRRF (2027+).",
            key="sb_tables_feed_url",
        )
        st.session_state["tables_feed_url"] = feed_url_sb

        uploaded_tbl = st.file_uploader(
            "Atualizar Tabelas Vigentes (.JSON)",
            type=["json"],
            key="sb_upload_tables_json",
            help="Importe um pacote JSON com novas faixas de INSS/IRRF para atualizar o motor sem mexer no código.",
        )
        if st.button("🔄 Sincronizar Agora (BACEN + Tabelas)", use_container_width=True, key="btn_sync_all_now"):
            raw_json_str = uploaded_tbl.getvalue().decode("utf-8") if uploaded_tbl is not None else None
            new_sync = sync_legal_tables_dynamic(remote_url=feed_url_sb, raw_json_override=raw_json_str)
            st.session_state["sync_status_info"] = new_sync
            st.cache_data.clear()
            if "params" in st.session_state:
                _recompute_from_current_params(st.session_state["params"], save_to_history=False)
            st.success(f"Sincronizado até {new_sync['ultima_competencia']}!")

if nav_page == "⚖️ Novo Cálculo (.PJC + Planilha)":
# ... existing code ...
