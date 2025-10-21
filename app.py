# app.py
import io
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator RDO", page_icon="🧰", layout="wide")

st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption("Versão Streamlit — envia seus PDFs de RDO e gere uma planilha consolidada")

with st.sidebar:
    st.header("Entrada")
    arquivos = st.file_uploader(
        "Selecione 1 ou mais PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Arraste e solte ou clique para escolher"
    )
    nome_excel = st.text_input("Nome do arquivo Excel (sem extensão)", value="rdo_consolidado")
    st.markdown("---")
    st.caption("O app extrai o bloco entre **RECURSOS EM OPERAÇÃO MÃO DE OBRA** e **RECURSOS EM OPERAÇÃO EQUIPAMENTO**.")
    st.caption("Linhas fora do padrão ficam como *raw_line* para conferência.")

def _texto_pdf(file_like: bytes) -> str:
    """Lê o PDF (bytes) e concatena o texto de todas as páginas."""
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

def _parse_bloco_mao_de_obra(texto: str) -> list[dict]:
    """
    Extrai o bloco 'RECURSOS EM OPERAÇÃO MÃO DE OBRA' → 'RECURSOS EM OPERAÇÃO EQUIPAMENTO'
    e tenta estruturar em colunas. Linhas não aderentes viram 'raw_line'.
    """
    start = texto.find("RECURSOS EM OPERAÇÃO MÃO DE OBRA")
    end = texto.find("RECURSOS EM OPERAÇÃO EQUIPAMENTO")
    if start == -1 or end == -1 or end <= start:
        return []

    bloco = texto[start:end]
    linhas = [l.strip() for l in bloco.splitlines()]

    # Remove cabeçalhos e "TOTAL"
    ignorar = {
        "Frente de Obra", "Classificação", "Função",
        "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado"
    }
    linhas = [l for l in linhas if l and l not in ignorar and "TOTAL" not in l.upper()]

    # Heurística de parsing por colunas separadas por múltiplos espaços
    registros = []
    padrao = re.compile(
        r"^(?P<frente>.+?)\s{2,}(?P<classificacao>.+?)\s{2,}(?P<funcao>.+?)\s{2,}"
        r"(?P<manha>\d+)\s+(?P<tarde>\d+)\s+(?P<noite>\d+)\s+"
        r"(?P<em_operacao>\d+)\s+(?P<fiscalizado>\d+)\s+(?P<geral>\d+)\s+(?P<contratado>\d+)$"
    )

    for l in linhas:
        m = padrao.match(l)
        if m:
            d = m.groupdict()
            for k in ["manha","tarde","noite","em_operacao","fiscalizado","geral","contratado"]:
                d[k] = int(d[k])
            registros.append(d)
        else:
            # Se não couber no padrão, guarda para conferência
            registros.append({"raw_line": l})

    return registros

def processar_arquivos(files) -> pd.DataFrame:
    linhas = []
    for i, f in enumerate(files, start=1):
        try:
            texto = _texto_pdf(f.read())
            dados = _parse_bloco_mao_de_obra(texto)
            for row in dados:
                row["_arquivo"] = f.name
                linhas.append(row)
        except Exception as e:
            linhas.append({"_arquivo": f.name, "erro": str(e)})

    # Normaliza para DataFrame (colunas ausentes viram NaN)
    df = pd.DataFrame(linhas)
    # Ordena colunas se as estruturadas existirem
    cols_ordem = [
        "_arquivo", "frente", "classificacao", "funcao",
        "manha", "tarde", "noite", "em_operacao", "fiscalizado", "geral", "contratado",
        "raw_line", "erro"
    ]
    return df[[c for c in cols_ordem if c in df.columns]]

col1, col2 = st.columns([1, 2])
with col1:
    executar = st.button("🚀 Extrair", use_container_width=True, type="primary", disabled=not arquivos)

with col2:
    if arquivos:
        st.info(f"{len(arquivos)} arquivo(s) selecionado(s).")

if executar:
    with st.spinner("Processando PDFs..."):
        df = processar_arquivos(arquivos)

    st.success("Extração concluída!")
    st.subheader("Prévia dos dados")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Excel para download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mao_de_Obra", index=False)

    st.download_button(
        label="💾 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"{(nome_excel or 'rdo_consolidado').strip()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption("Dica: se alguma linha cair em **raw_line**, é porque a diagramação do PDF fugiu do padrão. Dá para ajustar o regex no código.")
