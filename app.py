# app.py
# -*- coding: utf-8 -*-
# Extrai RDO (Mão de Obra + Equipamentos) no estilo do seu script Colab,
# consolidando tudo em uma única planilha.

import io
import re
import unicodedata
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator RDO (Mão de Obra + Equipamentos)", page_icon="🧰", layout="wide")
st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption("Replica a lógica do seu script: números em linhas separadas, backtracking de classificação/frente e mapeamento específico para Equipamentos.")

with st.sidebar:
    st.header("Entrada")
    arquivos = st.file_uploader("Selecione 1 ou mais PDFs", type=["pdf"], accept_multiple_files=True)
    nome_excel = st.text_input("Nome do arquivo Excel (sem extensão)", value="RDO_CONSOLIDADO")
    st.markdown("---")
    st.caption("Linhas fora do padrão vão para a aba **Inconsistencias**.")

# -------- Utils --------
def _texto_pdf(file_like: bytes) -> str:
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper()

def extrair_data_rdo(texto_completo: str) -> str:
    """Copia a sua lógica: usa a linha 11 do arquivo (index 10) como data, com fallback simples."""
    try:
        linhas = texto_completo.splitlines()
        data = linhas[10].strip()
        return data if data else "Data não encontrada"
    except IndexError:
        # fallback: tenta dd/mm/aaaa em qualquer lugar do topo
        m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", "\n".join(linhas[:30]) if 'linhas' in locals() else texto_completo[:1000])
        return m.group(1) if m else "Data não encontrada"

def _recorta_bloco(texto: str, tipo: str) -> str | None:
    """
    Recorta trecho entre:
      - Mão de Obra: 'RECURSOS EM OPERAÇÃO MÃO DE OBRA' → 'RECURSOS EM OPERAÇÃO EQUIPAMENTO'
      - Equipamento: 'RECURSOS EM OPERAÇÃO EQUIPAMENTO' → 'ASSINATURAS' (ou fim do doc se não achar)
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
        # se não achar o fim em Equipamentos, recorta até o fim do texto
        e = len(tnorm)

    # volta para o texto original por proporção
    ratio = len(texto) / max(len(tnorm), 1)
    return texto[int(s * ratio): int(e * ratio)]

# -------- Parser (copiando a "pegada" do seu Colab) --------
HEADERS_TO_IGNORE = {
    "Frente de Obra", "Classificação", "Função",
    "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado"
}

def _parse_secao(texto_completo: str, nome_arquivo: str, tipo: str) -> list[dict] | None:
    bloco = _recorta_bloco(texto_completo, tipo)
    if not bloco:
        return []

    data_rdo = extrair_data_rdo(texto_completo)
    linhas = [l.strip() for l in bloco.splitlines()]
    linhas = [l for l in linhas if l and l not in HEADERS_TO_IGNORE and "TOTAL" not in l.upper()]

    dados = []
    i = 0
    while i < len(linhas):
        if re.fullmatch(r"\d+", linhas[i]):  # começo de bloco numérico
            nums = []
            j = i
            while j < len(linhas) and re.fullmatch(r"\d+", linhas[j]):
                nums.append(int(linhas[j]))
                j += 1

            if len(nums) >= 6:
                # backtracking para Classificação / Frente / Função
                classificacao = ""
                frente = ""
                funcao_linhas = []
                achou = False

                # conjunto de palavras que denotam classificação por tipo
                if tipo == "Mão de Obra":
                    class_words = {"Direto", "Indireto", "DIRETO", "INDIRETO"}
                else:  # Equipamento
                    class_words = {"Mecânico", "Elétrico", "MECANICO", "ELETRICO", "MECÂNICO", "ELÉTRICO"}

                for k in range(i - 1, -1, -1):
                    lk = linhas[k].strip()
                    if lk in class_words:
                        # normaliza Direto/Indireto/Mecânico/Elétrico
                        up = _norm(lk)
                        if "DIRETO" in up:
                            classificacao = "Direto"
                        elif "INDIRETO" in up:
                            classificacao = "Indireto"
                        elif "MECANICO" in up:
                            classificacao = "Mecânico"
                        elif "ELETRICO" in up:
                            classificacao = "Elétrico"
                        # frente = linha anterior se não for outra classificação
                        if k > 0:
                            ant = linhas[k - 1].strip()
                            if ant not in class_words:
                                frente = ant
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

                # mapeamento de colunas
                if tipo == "Mão de Obra":
                    contratado, eom, fm, eot, ft, eon, fn = nums[0:7]
                else:  # Equipamento (ordem específica do seu script)
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

                i = j  # salta o bloco numérico
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

            # Mão de Obra
            dados_mo = _parse_secao(texto, f.name, "Mão de Obra")
            # Equipamentos
            dados_eq = _parse_secao(texto, f.name, "Equipamento")

            if not dados_mo and not dados_eq:
                inconsistencias.append({"Nome do Arquivo": f.name, "Linha": "[BLOCOS NÃO ENCONTRADOS OU SEM PADRÃO]"})
            else:
                for row in (dados_mo or []):
                    linhas.append(row)
                for row in (dados_eq or []):
                    linhas.append(row)

        except Exception as e:
            inconsistencias.append({"Nome do Arquivo": f.name, "Linha": f"[ERRO] {e}"})

    df = pd.DataFrame(linhas)
    cols_ordem = [
        "Nome do Arquivo", "Data da RDO", "Tipo", "Função/Equipamento", "Frente de Obra", "Classificação",
        "Contratado Geral", "Em operação (manhã)", "Fiscalizado (manhã)",
        "Em operação (tarde)", "Fiscalizado (tarde)",
        "Em operação (noite)", "Fiscalizado (noite)"
    ]
    df = df[cols_ordem] if not df.empty else pd.DataFrame(columns=cols_ordem)
    df_incons = pd.DataFrame(inconsistencias)
    return df, df_incons

# -------- UI --------
col1, col2 = st.columns([1, 2])
with col1:
    executar = st.button("🚀 Extrair", type="primary", use_container_width=True, disabled=not arquivos)
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

    # exporta
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
st.caption("Se algum arquivo ainda não vier, me envie 1 PDF exemplo (sem dados sensíveis) que ajusto as âncoras ou filtros.")
