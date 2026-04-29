
        # Remove somente se a linha for EXATAMENTE "TOTAL"
        if l.strip().upper() == "TOTAL":
            continue

        out.append(l)
    return out


def _parse_secao(texto_completo: str, nome_arquivo: str, tipo: str) -> list[dict]:
    bloco = _recorta_bloco(texto_completo, tipo)
    if not bloco:
        return []

    data_rdo = extrair_data_rdo(texto_completo)
    linhas = _limpa_linhas(bloco)

    dados = []
    i = 0
    while i < len(linhas):
        # detecta começo do bloco numérico
        if re.fullmatch(r"\d+", linhas[i]):
            nums = []
            j = i
            while j < len(linhas) and re.fullmatch(r"\d+", linhas[j]):
                nums.append(int(linhas[j]))
                j += 1

            if len(nums) >= 6:
                classificacao = ""
                frente = ""
                funcao_linhas = []
                achou = False

                # backtracking para achar classificação e frente
                for k in range(i - 1, -1, -1):
                    classificacao_encontrada = _label_classificacao(linhas[k], tipo)
                    if classificacao_encontrada:
                        classificacao = classificacao_encontrada

                        # frente tende a estar logo acima; pula números “perdidos”
                        idx_frente = k - 1
                        while idx_frente >= 0 and re.fullmatch(r"\d+", linhas[idx_frente]):
                            idx_frente -= 1
                        if idx_frente >= 0:
                            frente = linhas[idx_frente].strip()

                        funcao_linhas = [x.strip() for x in linhas[k + 1:i] if x.strip()]
                        achou = True
                        break

                if not achou:
                    frente = "FRENTE DE OBRA ÚNICA"
                    funcao_linhas = [x.strip() for x in linhas[max(0, i - 3):i] if x.strip()]

                funcao = " ".join(funcao_linhas).strip() if funcao_linhas else ""

                # completa para 7 números
                while len(nums) < 7:
                    nums.append(0)

                # mapeamento
                if tipo == "Mão de Obra":
                    contratado, eom, fm, eot, ft, eon, fn = nums[0:7]
                else:  # Equipamento (ordem específica)
                    contratado = nums[0]
                    eom, fm, eot, ft, eon, fn = nums[5], nums[6], nums[3], nums[4], nums[1], nums[2]

                dados.append({
                    "Nome do Arquivo": nome_arquivo,
                    "Data da RDO": data_rdo,
                    "Tipo": tipo,
                    "Função/Equipamento": funcao,
                    "Frente de Obra": frente,
                    "Classificação": classificacao,
                    "Contratado Geral": contratado,
                    "Em operação (manhã)": eom,
                    "Fiscalizado (manhã)": fm,
                    "Em operação (tarde)": eot,
                    "Fiscalizado (tarde)": ft,
                    "Em operação (noite)": eon,
                    "Fiscalizado (noite)": fn,
                })

                i = j
            else:
                i += 1
        else:
            i += 1

    return dados


def processar_arquivos(files):
    linhas, inconsistencias = [], []

    for f in files:
        try:
            raw = f.read()
            texto = _texto_pdf(raw)

            dados_mo = _parse_secao(texto, f.name, "Mão de Obra")
            dados_eq = _parse_secao(texto, f.name, "Equipamento")

            if not dados_mo and not dados_eq:
                inconsistencias.append({
                    "Nome do Arquivo": f.name,
                    "Linha": "[BLOCOS NÃO ENCONTRADOS OU SEM PADRÃO]"
                })
            else:
                linhas.extend(dados_mo)
                linhas.extend(dados_eq)

        except Exception as e:
            inconsistencias.append({"Nome do Arquivo": f.name, "Linha": f"[ERRO] {e}"})

    cols_ordem = [
        "Nome do Arquivo", "Data da RDO", "Tipo", "Função/Equipamento",
        "Frente de Obra", "Classificação",
        "Contratado Geral",
        "Em operação (manhã)", "Fiscalizado (manhã)",
        "Em operação (tarde)", "Fiscalizado (tarde)",
        "Em operação (noite)", "Fiscalizado (noite)",
    ]

    df = pd.DataFrame(linhas)
    if df.empty:
        df = pd.DataFrame(columns=cols_ordem)
    else:
        df = df[cols_ordem]

    df_incons = pd.DataFrame(inconsistencias)
    return df, df_incons


# -------- UI --------
col1, col2 = st.columns([1, 2])
with col1:
    executar = st.button(
        "🚀 Extrair",
        type="primary",
        use_container_width=True,
        disabled=not arquivos,
    )
with col2:
    if arquivos:
        st.info(f"{len(arquivos)} arquivo(s) selecionado(s).")

if executar:
    with st.spinner("Processando PDFs..."):
        df, df_incons = processar_arquivos(arquivos)

    st.success("Extração concluída!")
    st.subheader("Prévia dos dados (Mão de Obra + Equipamentos)")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df_incons.empty:
        with st.expander("Inconsistências / linhas não parseadas"):
            st.dataframe(df_incons, use_container_width=True, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Consolidado", index=False)
        if not df_incons.empty:
            df_incons.to_excel(writer, sheet_name="Inconsistencias", index=False)

    st.download_button(
        "💾 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"{(nome_excel or 'RDO_CONSOLIDADO').strip()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption("Se algum PDF específico não vier, envie 1 exemplo (sem dados sensíveis) e eu ajusto as âncoras/filtros.")
