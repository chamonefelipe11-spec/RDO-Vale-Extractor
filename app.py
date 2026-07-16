# app.py
# -*- coding: utf-8 -*-
# Extrai RDO (Mão de Obra + Equipamentos) e consolida em um único Excel.
# Compatível com Streamlit Cloud (sem tkinter).
# Inclui correção: NÃO remover linhas que contenham "TOTAL" (ex.: "ESTAÇÃO TOTAL");
# remove apenas se a linha for EXATAMENTE "TOTAL".
#
# Correções aplicadas:
# - Para a seção "Equipamento", reconhece também "DIRETO/INDIRETO"
#   (além de "MECÂNICO/ELÉTRICO"), evitando juntar "Indireto" e "Frente de Obra"
#   dentro de "Função/Equipamento".
# - Classifica "INDIRETO" corretamente. Antes, a checagem por substring fazia
#   "INDIRETO" cair como "DIRETO", porque "DIRETO" está contido em "INDIRETO".

import io
import re
import unicodedata
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Hub de Parsers RDO VALE",
    page_icon="🧰",
    layout="wide",
)

st.title("🧰 Hub de Parsers de RDO VALE (PDF → Excel)")

with st.sidebar:
    st.header("Parser")
    parser_selecionado = st.selectbox(
        "Selecione o tipo de extração",
        [
            "Mão de Obra + Equipamentos",
            "Atividades + Comentários",
        ],
    )

    st.header("Entrada")
    arquivos = st.file_uploader(
        "Selecione 1 ou mais PDFs",
        type=["pdf"],
        accept_multiple_files=True,
    )
    nome_excel = st.text_input(
        "Nome do arquivo Excel (sem extensão)",
        value="RDO_CONSOLIDADO",
    )
    st.markdown("---")
    if parser_selecionado == "Mão de Obra + Equipamentos":
        st.caption("Linhas fora do padrão vão para a aba **Inconsistencias**.")
    else:
        st.caption(
            "O Excel terá apenas as abas **Atividades** e **Comentarios**. "
            "Avisos de leitura aparecem somente na tela."
        )

if parser_selecionado == "Mão de Obra + Equipamentos":
    st.caption(
        "Consolida Mão de Obra + Equipamentos no mesmo Excel. "
        "Mantém linhas como 'ESTAÇÃO TOTAL' e classifica corretamente Direto/Indireto."
    )
else:
    st.caption(
        "Extrai exclusivamente Atividades e Comentários dos RDOs, incluindo "
        "a Data da RDO e o nome do arquivo de origem."
    )

# -------- Utils --------
def _texto_pdf(file_like: bytes) -> str:
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper().strip()


def extrair_data_rdo(texto_completo: str) -> str:
    """Usa a linha 11 do arquivo (index 10) como data; fallback para dd/mm/aaaa no topo."""
    linhas = texto_completo.splitlines()
    try:
        data = linhas[10].strip()
        return data if data else "Data não encontrada"
    except Exception:
        topo = "\n".join(linhas[:30]) if linhas else texto_completo[:1000]
        m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", topo)
        return m.group(1) if m else "Data não encontrada"


def _recorta_bloco(texto: str, tipo: str) -> str | None:
    """
    Recorta trecho entre:
      - Mão de Obra: 'RECURSOS EM OPERAÇÃO MÃO DE OBRA' → 'RECURSOS EM OPERAÇÃO EQUIPAMENTO'
      - Equipamento: 'RECURSOS EM OPERAÇÃO EQUIPAMENTO' → 'ASSINATURAS' (ou fim se não achar)
    Robusto a variações e acentos.
    """
    tnorm = _norm(texto)

    if tipo == "Mão de Obra":
        starts = [
            "RECURSOS EM OPERACAO MAO DE OBRA",
            "RECURSOS EM OPERACAO - MAO DE OBRA",
            "RECURSOS DE OPERACAO MAO DE OBRA",
        ]
        ends = [
            "RECURSOS EM OPERACAO EQUIPAMENTO",
            "RECURSOS EM OPERACAO - EQUIPAMENTO",
            "RECURSOS DE OPERACAO EQUIPAMENTO",
        ]
    else:  # Equipamento
        starts = [
            "RECURSOS EM OPERACAO EQUIPAMENTO",
            "RECURSOS EM OPERACAO - EQUIPAMENTO",
            "RECURSOS DE OPERACAO EQUIPAMENTO",
        ]
        ends = [
            "ASSINATURAS",
            "ASSINATURA",
            "RESPONSAVEL",
            "RESPONSÁVEL",
            "OBSERVACOES",
            "OBSERVAÇÕES",
        ]

    s = next((tnorm.find(x) for x in starts if tnorm.find(x) != -1), -1)
    if s == -1:
        return None

    e = next((tnorm.find(x, s + 1) for x in ends if tnorm.find(x, s + 1) != -1), -1)
    if e == -1 or e <= s:
        e = len(tnorm)

    # mapeia de volta ao texto original (aproximação por proporção)
    ratio = len(texto) / max(len(tnorm), 1)
    return texto[int(s * ratio): int(e * ratio)]


# -------- Parser --------
HEADERS_TO_IGNORE = {
    "Frente de Obra", "Classificação", "Função",
    "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado"
}

CLASSIFICACOES_MAO_DE_OBRA = {
    "DIRETO": "Direto",
    "INDIRETO": "Indireto",
}

CLASSIFICACOES_EQUIPAMENTO = {
    **CLASSIFICACOES_MAO_DE_OBRA,
    "MECANICO": "Mecânico",
    "ELETRICO": "Elétrico",
}


def _label_classificacao(linha: str, tipo: str) -> str:
    """Retorna a classificação só quando a linha inteira é a classificação."""
    classificacoes = (
        CLASSIFICACOES_EQUIPAMENTO
        if tipo == "Equipamento"
        else CLASSIFICACOES_MAO_DE_OBRA
    )
    return classificacoes.get(_norm(linha), "")


def _limpa_linhas(bloco: str) -> list[str]:
    """Remove vazios, cabeçalhos e TOTAL (apenas quando for exatamente 'TOTAL')."""
    linhas_brutas = [l.strip() for l in bloco.splitlines()]
    out = []
    for l in linhas_brutas:
        if not l:
            continue
        if l in HEADERS_TO_IGNORE:
            continue

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


# -------- Novo parser: Atividades + Comentários --------
COLUNAS_ATIVIDADES = [
    "Nome do Arquivo",
    "Data da RDO",
    "Frente de Obra",
    "Área",
    "Sub-Área",
    "Descrição",
]

COLUNAS_COMENTARIOS = [
    "Nome do Arquivo",
    "Data da RDO",
    "Data do Comentário",
    "Área Responsável",
    "Usuário",
    "Comentário",
]

_DIAS_SEMANA = {
    "SEGUNDA-FEIRA", "TERCA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA",
    "SEXTA-FEIRA", "SABADO", "DOMINGO",
}


def _limpar_celula_tabela(valor) -> str:
    """Converte a célula extraída do PDF em texto limpo, sem alterar o conteúdo."""
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def _linhas_tabela_limpas(tabela) -> list[list[str]]:
    linhas = tabela.extract() or []
    return [[_limpar_celula_tabela(c) for c in linha] for linha in linhas]


def _extrair_data_rdo_atividades_comentarios(doc) -> str:
    """
    Obtém a Data da RDO no quadro de identificação da primeira página.
    Usa a geometria das tabelas como método principal e texto como fallback.
    """
    if len(doc) == 0:
        return "Data não encontrada"

    pagina = doc[0]

    # Método principal: célula do quadro de identificação contendo "Data RDO".
    try:
        tabelas = pagina.find_tables().tables
        for tabela in tabelas:
            for linha in _linhas_tabela_limpas(tabela):
                for celula in linha:
                    celula_norm = _norm(celula)
                    if "DATA RDO" not in celula_norm:
                        continue
                    m = re.search(r"DATA\s*RDO\s*:?[ ]*(\d{2}/\d{2}/\d{4})", celula_norm)
                    if m:
                        return m.group(1)
    except Exception:
        pass

    texto = pagina.get_text("text")
    linhas = [l.strip() for l in texto.splitlines() if l.strip()]

    # Fallback 1: data imediatamente anterior ao dia da semana.
    for i, linha in enumerate(linhas):
        if _norm(linha) in _DIAS_SEMANA and i > 0:
            m = re.fullmatch(r"\d{2}/\d{2}/\d{4}", linhas[i - 1])
            if m:
                return m.group(0)

    # Fallback 2: formato em que o rótulo vem antes da data.
    texto_norm = _norm(texto)
    m = re.search(r"DATA\s*RDO\s*:?[ ]*(\d{2}/\d{2}/\d{4})", texto_norm)
    if m:
        return m.group(1)

    # Último fallback: mantém a lógica já existente no aplicativo.
    return extrair_data_rdo(texto)


def _eh_cabecalho_atividades(linha: list[str]) -> bool:
    valores = [_norm(c) for c in linha[:4]]
    return (
        len(valores) >= 4
        and valores[0] == "FRENTE DE OBRA"
        and valores[1] == "AREA"
        and valores[2] in {"SUB-AREA", "SUB AREA"}
        and valores[3] == "DESCRICAO"
    )


def _tabela_contem_titulo(linhas: list[list[str]], titulo: str) -> bool:
    titulo_norm = _norm(titulo)
    return any(
        titulo_norm in _norm(celula)
        for linha in linhas
        for celula in linha
        if celula
    )


def _parece_continuacao_atividades(linhas: list[list[str]]) -> bool:
    """Reconhece a continuação da tabela de atividades nas páginas seguintes."""
    for linha in linhas:
        if len(linha) < 4:
            continue
        frente, area, subarea, descricao = (linha + [""] * 4)[:4]
        if _norm(frente).startswith("FRENTE DE OBRA") and any([area, subarea, descricao]):
            return True
    return False


def _extrair_atividades_tabela(
    linhas: list[list[str]],
    nome_arquivo: str,
    data_rdo: str,
) -> list[dict]:
    registros = []

    for linha in linhas:
        valores = (linha + [""] * 4)[:4]
        frente, area, subarea, descricao = valores

        texto_linha_norm = _norm(" ".join(valores))
        if not texto_linha_norm:
            continue
        if texto_linha_norm == "ATIVIDADES":
            continue
        if _eh_cabecalho_atividades(valores):
            continue
        if "COMENTARIOS RDO" in texto_linha_norm:
            continue

        # Somente linhas que efetivamente pertencem ao quadro de atividades.
        if not any([frente, area, subarea, descricao]):
            continue

        registros.append({
            "Nome do Arquivo": nome_arquivo,
            "Data da RDO": data_rdo,
            "Frente de Obra": frente,
            "Área": area,
            "Sub-Área": subarea,
            "Descrição": descricao,
        })

    return registros


def _extrair_comentarios_tabela(
    linhas: list[list[str]],
    nome_arquivo: str,
    data_rdo: str,
) -> list[dict]:
    """Extrai todas as linhas dos quadros 'COMENTÁRIOS RDO'."""
    registros = []
    atual = None

    def salvar_atual():
        nonlocal atual
        if atual is not None:
            registros.append(atual)
            atual = None

    for linha in linhas:
        valores = (linha + [""] * 4)[:4]
        data_comentario, area_responsavel, usuario, comentario = valores
        texto_linha_norm = _norm(" ".join(valores))

        if not texto_linha_norm:
            continue
        if "COMENTARIOS RDO" in texto_linha_norm:
            continue
        if (
            _norm(data_comentario) == "DATA"
            and _norm(area_responsavel) == "AREA RESPONSAVEL"
            and _norm(usuario) == "USUARIO"
            and _norm(comentario) == "COMENTARIO"
        ):
            continue

        inicio_novo = bool(re.match(r"^\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?$", data_comentario))

        if inicio_novo:
            salvar_atual()
            atual = {
                "Nome do Arquivo": nome_arquivo,
                "Data da RDO": data_rdo,
                "Data do Comentário": data_comentario,
                "Área Responsável": area_responsavel,
                "Usuário": usuario,
                "Comentário": comentario,
            }
        elif atual is not None:
            # Caso uma linha da tabela seja quebrada pelo gerador do PDF,
            # concatena somente nas respectivas colunas.
            if data_comentario:
                atual["Data do Comentário"] = _limpar_celula_tabela(
                    f'{atual["Data do Comentário"]} {data_comentario}'
                )
            if area_responsavel:
                atual["Área Responsável"] = _limpar_celula_tabela(
                    f'{atual["Área Responsável"]} {area_responsavel}'
                )
            if usuario:
                atual["Usuário"] = _limpar_celula_tabela(
                    f'{atual["Usuário"]} {usuario}'
                )
            if comentario:
                atual["Comentário"] = _limpar_celula_tabela(
                    f'{atual["Comentário"]} {comentario}'
                )

    salvar_atual()
    return registros


def _parse_atividades_comentarios_pdf(raw: bytes, nome_arquivo: str):
    atividades = []
    comentarios = []
    avisos = []

    with fitz.open(stream=raw, filetype="pdf") as doc:
        data_rdo = _extrair_data_rdo_atividades_comentarios(doc)
        em_atividades = False
        encontrou_tabela = False

        for numero_pagina, pagina in enumerate(doc, start=1):
            try:
                tabelas = sorted(
                    pagina.find_tables().tables,
                    key=lambda t: (round(t.bbox[1], 2), round(t.bbox[0], 2)),
                )
            except Exception as e:
                avisos.append({
                    "Nome do Arquivo": nome_arquivo,
                    "Página": numero_pagina,
                    "Aviso": f"Não foi possível detectar tabelas: {e}",
                })
                continue

            if tabelas:
                encontrou_tabela = True

            for tabela in tabelas:
                linhas = _linhas_tabela_limpas(tabela)
                if not linhas:
                    continue

                eh_comentarios = _tabela_contem_titulo(linhas, "COMENTÁRIOS RDO")
                eh_atividades = (
                    _tabela_contem_titulo(linhas, "ATIVIDADES")
                    or any(_eh_cabecalho_atividades(linha) for linha in linhas)
                )

                if eh_comentarios:
                    em_atividades = False
                    comentarios.extend(
                        _extrair_comentarios_tabela(linhas, nome_arquivo, data_rdo)
                    )
                    continue

                if eh_atividades:
                    em_atividades = True
                    atividades.extend(
                        _extrair_atividades_tabela(linhas, nome_arquivo, data_rdo)
                    )
                    continue

                if em_atividades and _parece_continuacao_atividades(linhas):
                    atividades.extend(
                        _extrair_atividades_tabela(linhas, nome_arquivo, data_rdo)
                    )

        if not encontrou_tabela:
            avisos.append({
                "Nome do Arquivo": nome_arquivo,
                "Página": "-",
                "Aviso": (
                    "Nenhuma tabela vetorial foi detectada. O arquivo pode estar "
                    "digitalizado como imagem e exigir OCR."
                ),
            })

        if data_rdo == "Data não encontrada":
            avisos.append({
                "Nome do Arquivo": nome_arquivo,
                "Página": 1,
                "Aviso": "Data da RDO não encontrada.",
            })

    return atividades, comentarios, avisos


def processar_arquivos_atividades_comentarios(files):
    """
    Processa lotes de PDFs e retorna exclusivamente os dados solicitados:
    atividades, comentários, Data da RDO e nome do arquivo.
    """
    atividades = []
    comentarios = []
    avisos = []

    for f in files:
        try:
            raw = f.getvalue() if hasattr(f, "getvalue") else f.read()
            dados_atividades, dados_comentarios, avisos_arquivo = (
                _parse_atividades_comentarios_pdf(raw, f.name)
            )
            atividades.extend(dados_atividades)
            comentarios.extend(dados_comentarios)
            avisos.extend(avisos_arquivo)

            if not dados_atividades:
                avisos.append({
                    "Nome do Arquivo": f.name,
                    "Página": "-",
                    "Aviso": "Nenhuma atividade encontrada.",
                })
            if not dados_comentarios:
                avisos.append({
                    "Nome do Arquivo": f.name,
                    "Página": "-",
                    "Aviso": "Nenhum comentário encontrado.",
                })

        except Exception as e:
            avisos.append({
                "Nome do Arquivo": getattr(f, "name", "Arquivo sem nome"),
                "Página": "-",
                "Aviso": f"[ERRO] {e}",
            })

    df_atividades = pd.DataFrame(atividades, columns=COLUNAS_ATIVIDADES)
    df_comentarios = pd.DataFrame(comentarios, columns=COLUNAS_COMENTARIOS)
    df_avisos = pd.DataFrame(avisos, columns=["Nome do Arquivo", "Página", "Aviso"])

    return df_atividades, df_comentarios, df_avisos


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
    if parser_selecionado == "Mão de Obra + Equipamentos":
        # Parser original preservado, sem alteração em suas funções.
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

    else:
        with st.spinner("Extraindo atividades e comentários dos PDFs..."):
            df_atividades, df_comentarios, df_avisos = (
                processar_arquivos_atividades_comentarios(arquivos)
            )

        st.success(
            f"Extração concluída: {len(df_atividades)} atividade(s) e "
            f"{len(df_comentarios)} comentário(s)."
        )

        aba_atividades, aba_comentarios = st.tabs(["Atividades", "Comentários"])
        with aba_atividades:
            st.dataframe(
                df_atividades,
                use_container_width=True,
                hide_index=True,
            )
        with aba_comentarios:
            st.dataframe(
                df_comentarios,
                use_container_width=True,
                hide_index=True,
            )

        if not df_avisos.empty:
            with st.expander("Avisos de leitura"):
                st.dataframe(df_avisos, use_container_width=True, hide_index=True)

        # O Excel contém estritamente os dados solicitados.
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_atividades.to_excel(writer, sheet_name="Atividades", index=False)
            df_comentarios.to_excel(writer, sheet_name="Comentarios", index=False)

        st.download_button(
            "💾 Baixar Excel",
            data=buffer.getvalue(),
            file_name=f"{(nome_excel or 'RDO_ATIVIDADES_COMENTARIOS').strip()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.markdown("---")
st.caption(
    "O parser de Atividades + Comentários utiliza as tabelas vetoriais do PDF. "
    "Arquivos digitalizados apenas como imagem podem exigir OCR."
)
