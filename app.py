# ... existing code ...
        sp1, sp2, sp3, sp4 = st.columns([1.8, 2.2, 2.2, 1.8])
        s_proc = sp1.text_input("Número do Processo *", value=str(s_ext.get("process_number") or "0000000-00.2025.5.04.0000"))
# ... existing code ...
        # Botões de Exportação no Padrão PJe-Calc (PDF Oficial PJe-Calc, Arquivo .PJC e Excel Multi-Abas)
        if st.session_state.get("sentence_calc_result"):
            res_liq = st.session_state["sentence_calc_result"]
            st.markdown("### 3. Arquivos Prontos para Protocolar (Padrão PJe-Calc: PDF + .PJC + Excel)")
            dl1, dl2, dl3 = st.columns(3)

            pdf_pjecalc_bytes = build_pjecalc_pdf_bytes(res_liq)
            pjc_xml_bytes = build_pjecalc_pjc_bytes(res_liq)
            xlsx_pjecalc_bytes = build_pjecalc_excel_bytes(res_liq)

            safe_proc = re.sub(r"[^0-9A-Za-z_-]+", "_", str(res_liq.get("process_number") or "processo"))

            dl1.download_button(
                label="📄 Baixar Demonstrativo PJe-Calc (.PDF)",
                data=pdf_pjecalc_bytes,
                file_name=f"Demonstrativo_PJeCalc_{safe_proc}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_liq_pdf_pjecalc",
            )
            dl2.download_button(
                label="⚖️ Baixar Arquivo PJe-Calc (.PJC)",
                data=pjc_xml_bytes,
                file_name=f"Calculo_PJeCalc_{safe_proc}.pjc",
                mime="application/xml",
                use_container_width=True,
                key="dl_liq_pjc_xml",
            )
            dl3.download_button(
                label="📊 Baixar Planilha Detalhada (.XLSX)",
                data=xlsx_pjecalc_bytes,
                file_name=f"Liquidacao_PJeCalc_{safe_proc}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_liq_xlsx_pjecalc",
            )
# ... existing code ...
