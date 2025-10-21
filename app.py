# app.py
import io
import re
import unicodedata
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator RDO", page_icon="🧰", layout="wide")
st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption("Extrai o bloco de MÃO DE OBRA exatamente como no script original (blocos de números em linhas separadas).")

with st.sidebar:
    st.header("Entrada")
    arquivos = st.file_uploader(
        "Selecione 1 ou mais PDFs",
        type=["pdf"],
        accept_multiple_files=True,
    )
    nome_excel = st.text_input("Nome do arquivo Excel (sem extensão)", value="rdo_consolidado")
    st.markdown("---")
    st.caption("Linhas que não aderirem ao padrão irão para a aba **Inconsistencias**.")

# ---------- Utils ----------
def _texto_pdf(file_like: bytes) -> str:
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper()

def _recorta_bloco(texto: str) -> str | None:
    """Recorta trecho entre 'RECURSOS EM OPERAÇÃO MÃO DE OBRA' e 'RECURSOS EM OPERAÇÃO EQUIPAMENTO' (robusto a variações)."""
    tnorm = _norm(texto)
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
    s = next((tnorm.find(x) for x in starts if tnorm.find(x) != -1), -1)
    if s == -1:
        return None
    e = next((tnorm.find(x, s + 1) for x in ends if tnorm.find(x, s + 1) != -1), -1)
    if e == -1 or e <= s:
        return None
    # volta para o texto original por proporção
    ratio = len(texto) / max(len(tnorm), 1)
    return texto[int(s * ratio): int(e * ratio)]

# ---------- Parser (igual à lógica do original) ----------
def _parse_bloco_mao_de_obra(texto: str) -> list[dict]:
    bloco = _recorta_bloco(texto)
    if not bloco:
        return []

    linhas = [l.strip() for l in bloco.splitlines()]

    # remove cabeçalhos/linhas inúteis (igual ao script original)
    ignorar = {
        "Frente de Obra", "Classificação", "Função",
        "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado"
    }
    linhas = [l for l in linhas if l and l not in ignorar and "TOTAL" not in l.upper()]

    dados = []
    i = 0
    while i < len(linhas) - 6:
        # junta sequência de linhas estritamente numéricas
        bloco_numeros = []
        j = i
        while j < len(linhas) and re.fullmatch(r"\d+", linhas[j]):
            bloco_numeros.append(int(linhas[j]))
            j += 1

        if len(bloco_numeros) >= 6:
            # retrocede para achar a Classificação (Direto/Indireto),
            # pega Frente na linha anterior e Função nas linhas entre a Classificação e os números
            classificacao = ""
            frente = ""
            funcao_linhas = []
            achou = False
            for k in range(i - 1, -1, -1):
                lk = linhas[k].strip()
                if lk in ("Direto", "Indireto", "DIRETO", "INDIRETO"):
                    classificacao = "Direto" if "DIRETO" in lk.upper() else "Indireto"
                    frente = linhas[k - 1].strip() if k - 1 >= 0 else ""
                    funcao_linhas = [x.strip() for x in linhas[k + 1:i] if x.strip()]
                    achou = True
                    break
            # fallback se não achar
            if not achou:
                classificacao = ""
                frente = "FRENTE DE OBRA ÚNICA"
                funcao_linhas = [x.strip() for x in linhas[max(0, i - 3):i] if x.strip()]

            funcao = " ".join(funcao_linhas).strip() or (funcao_linhas[0] if funcao_linhas else "")

            # pad para 7 números (contratado + 6 turnos)
            while len(bloco_numeros) < 7:
                bloco_numeros.append(0)

            dados.append({
                "Função": funcao,
                "Frente de Obra": frente,
                "Classificação": classificacao,
                "Contratado Geral": bloco_numeros[0],
                "Em operação (manhã)": bloco_numeros[1],
                "Fiscalizado (manhã)": bloco_numeros[2],
                "Em operação (tarde)": bloco_numeros[3],
                "Fiscalizado (tarde)": bloco_numeros[4],
                "Em operação (noite)": bloco_numeros[5],
                "Fiscalizado (noite)": bloco_numeros[6],
            })

            i = j  # pula para depois do bloco numérico
        else:
            i += 1

    return dados

# ---------- Pipeline ----------
def processar_arquivos(files):
    linhas, inconsistencias = [], []
    for f in files:
        try:
            texto = _texto_pdf(f.read())
            dados = _parse_bloco_mao_de_obra(texto)
            if not dados:
                inconsistencias.append({"Nome do Arquivo": f.name, "Linha": "[BLOCO NÃO ENCONTRADO OU SEM PADRÃO]"})
            for row in dados:
                row["Nome do Arquivo"] = f.name
                linhas.append(row)
        except Exception as e:
            inconsistencias.append({"Nome do Arquivo": f.name, "Linha": f"[ERRO] {e}"})

    df = pd.DataFrame(linhas)
    cols_ordem = [
        "Nome do Arquivo",
        "Função",
        "Frente de Obra",
        "Classificação",
        "Contratado Geral",
        "Em operação (manhã)",
        "Fiscalizado (manhã)",
        "Em operação (tarde)",
        "Fiscalizado (tarde)",
        "Em operação (noite)",
        "Fiscalizado (noite)",
    ]
    df = df[cols_ordem] if not df.empty else pd.DataFrame(columns=cols_ordem)
    df_incons = pd.DataFrame(inconsistencias)
    return df, df_incons

# ---------- UI ----------
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
    st.subheader("Prévia dos dados")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df_incons.empty:
        with st.expander("Inconsistências / linhas não parseadas"):
            st.dataframe(df_incons, use_container_width=True, hide_index=True)

    # exporta
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mao_de_Obra", index=False)
        if not df_incons.empty:
            df_incons.to_excel(writer, sheet_name="Inconsistencias", index=False)

    st.download_button(
        "💾 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"{(nome_excel or 'rdo_consolidado').strip()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption("Parser replica a lógica do app desktop (sequências de números em linhas separadas, com backtracking para Classificação/Frente/Função).")
