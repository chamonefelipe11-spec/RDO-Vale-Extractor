# app.py
# -*- coding: utf-8 -*-
# Extrai RDO (Mão de Obra + Equipamentos) e consolida em um único Excel.
# Compatível com Streamlit Cloud (sem tkinter).
# Inclui correção: NÃO remover linhas que contenham "TOTAL" (ex.: "ESTAÇÃO TOTAL");
# remove apenas se a linha for EXATAMENTE "TOTAL".
#
# ✅ Correção solicitada (aplicada):
# Para a seção "Equipamento", o parser agora reconhece também "DIRETO/INDIRETO"
# (além de "MECÂNICO/ELÉTRICO"), evitando juntar "Indireto" e "Frente de Obra"
# dentro de "Função/Equipamento".

import io
import re
import unicodedata
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Extrator RDO (Mão de Obra + Equipamentos)",
    page_icon="🧰",
    layout="wide",
)

st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption(
    "Consolida Mão de Obra + Equipamentos no mesmo Excel. "
    "Correção aplicada: só remove a linha exatamente 'TOTAL' (mantém 'ESTAÇÃO TOTAL')."
)

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
    """Usa a linha 11 do arquivo (index 10) como data; fallback para dd/mm/aaaa no topo."""
    linhas = texto_completo.splitlines()
    try:
        data = linhas[10].strip()
        return data if data else "Data não encontrada"
    except Exception:
        topo = "\n".join(linhas[:30]) if linhas else texto_completo[:1000]
        m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", topo)
        return m.group(1) if m else "Data não encontrada"

def extrair_pluviometria_e_numero(texto_completo: str) -> dict:
    """
    Extrai Número RDO + Data + Pluviometria manhã/tarde
    usando a MESMA lógica do código 2 (Colab).
    """
    dados = {
        "Número RDO": "Não encontrado",
        "Data RDO": "Não encontrada",
        "Pluviometria Manhã": "Não encontrada",
        "Pluviometria Tarde": "Não encontrada",
    }

    linhas = texto_completo.splitlines()

    # reaproveita sua função de data
    dados["Data RDO"] = extrair_data_rdo(texto_completo)

    try:
        idx_status = linhas.index("Status :")
        dados["Número RDO"] = linhas[idx_status - 2].strip()
    except (ValueError, IndexError):
        pass

    try:
        idx_num_rdo = linhas.index("Número RDO :")
        dados["Pluviometria Manhã"] = linhas[idx_num_rdo - 3].strip()
        dados["Pluviometria Tarde"] = linhas[idx_num_rdo - 2].strip()
    except (ValueError, IndexError):
        pass

    return dados

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


def _limpa_linhas(bloco: str) -> list[str]:
    """Remove vazios, cabeçalhos e TOTAL (apenas quando for exatamente 'TOTAL')."""
    linhas_brutas = [l.strip() for l in bloco.splitlines()]
    out = []
    for l in linhas_brutas:
        if not l:
            continue
        if l in HEADERS_TO_IGNORE:
            continue

        # ✅ Correção: remove somente se a linha for EXATAMENTE "TOTAL"
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

                if tipo == "Mão de Obra":
                    class_words = {"DIRETO", "INDIRETO"}
                else:  # Equipamento
                    # ✅ CORREÇÃO AQUI:
                    # Equipamento também pode vir como Direto/Indireto (como no seu caso),
                    # além de alguns modelos que usam Mecânico/Elétrico.
                    class_words = {"DIRETO", "INDIRETO", "MECANICO", "ELETRICO", "MECÂNICO", "ELÉTRICO"}

                # backtracking para achar classificação e frente
                for k in range(i - 1, -1, -1):
                    lk = linhas[k].strip()
                    upk = _norm(lk)

                    if upk in class_words:
                        if "DIRETO" in upk:
                            classificacao = "Direto"
                        elif "INDIRETO" in upk:
                            classificacao = "Indireto"
                        elif "MECANICO" in upk:
                            classificacao = "Mecânico"
                        elif "ELETRICO" in upk:
                            classificacao = "Elétrico"

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
    pluv_rows = []

    for f in files:
        try:
            raw = f.read()
            texto = _texto_pdf(raw)

            # 🔹 SEU CÓDIGO ORIGINAL (intacto)
            dados_mo = _parse_secao(texto, f.name, "Mão de Obra")
            dados_eq = _parse_secao(texto, f.name, "Equipamento")

            # 🔹 NOVO: pluviometria
            pluv = extrair_pluviometria_e_numero(texto)
            pluv_rows.append({
                "Nome do Arquivo": f.name,
                "Data RDO": pluv["Data RDO"],
                "Número RDO": pluv["Número RDO"],
                "Pluviometria Manhã": pluv["Pluviometria Manhã"],
                "Pluviometria Tarde": pluv["Pluviometria Tarde"],
            })

            if not dados_mo and not dados_eq:
                inconsistencias.append({
                    "Nome do Arquivo": f.name,
                    "Linha": "[BLOCOS NÃO ENCONTRADOS OU SEM PADRÃO]"
                })
            else:
                linhas.extend(dados_mo)
                linhas.extend(dados_eq)

        except Exception as e:
            inconsistencias.append({
                "Nome do Arquivo": f.name,
                "Linha": f"[ERRO] {e}"
            })

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

    df_pluv = pd.DataFrame(pluv_rows)

    return df, df_incons, df_pluv



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
        df, df_incons, df_pluv = processar_arquivos(arquivos)

    st.success("Extração concluída!")
    st.subheader("Prévia dos dados (Mão de Obra + Equipamentos)")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.subheader("🌧️ Prévia da Pluviometria")
    st.dataframe(df_pluv, use_container_width=True, hide_index=True)


    if not df_incons.empty:
        with st.expander("Inconsistências / linhas não parseadas"):
            st.dataframe(df_incons, use_container_width=True, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Consolidado", index=False)
    df_pluv.to_excel(writer, sheet_name="Pluviometria", index=False)
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

