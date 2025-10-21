# app.py
import io
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator RDO", page_icon="🧰", layout="wide")

st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption("Versão Streamlit — selecione seus PDFs e gere a planilha exatamente no layout desejado.")

with st.sidebar:
    st.header("Entrada")
    arquivos = st.file_uploader(
        "Selecione 1 ou mais PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Arraste e solte ou clique para escolher",
    )
    nome_excel = st.text_input("Nome do arquivo Excel (sem extensão)", value="rdo_consolidado")
    st.markdown("---")
    st.caption("O app extrai o bloco entre **RECURSOS EM OPERAÇÃO MÃO DE OBRA** e **RECURSOS EM OPERAÇÃO EQUIPAMENTO**.")
    st.caption("Linhas fora do padrão vão para a aba 'Inconsistências'.")

# ---------- Utilidades ----------

def _texto_pdf(file_like: bytes) -> str:
    """Lê o PDF (bytes) e concatena o texto de todas as páginas."""
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

# ---------- Parser (corrigido) ----------

def _parse_bloco_mao_de_obra(texto: str) -> list[dict]:
    """
    Extrai o bloco MÃO DE OBRA e estrutura linhas no formato desejado:

    Função | Frente de Obra | Classificação | Contratado Geral |
    Em operação (manhã) | Fiscalizado (manhã) |
    Em operação (tarde) | Fiscalizado (tarde) |
    Em operação (noite) | Fiscalizado (noite)
    """
    start = texto.find("RECURSOS EM OPERAÇÃO MÃO DE OBRA")
    end = texto.find("RECURSOS EM OPERAÇÃO EQUIPAMENTO")
    if start == -1 or end == -1 or end <= start:
        return []

    bloco = texto[start:end]
    linhas = [l.strip() for l in bloco.splitlines()]

    # Remove cabeçalhos e totais comuns
    cabecalhos = {
        "Frente de Obra", "Frente de obra", "Frente", "Classificação", "Função",
        "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado",
        "Em operação", "Fiscalizado (manhã)", "Fiscalizado (tarde)", "Fiscalizado (noite)"
    }
    linhas = [l for l in linhas if l and "TOTAL" not in l.upper() and l not in cabecalhos]

    registros = []

    # Espera 3 campos textuais + 7 números.
    # Ordem típica (mas pode variar nos 3 primeiros):
    # [Função]  [Frente de Obra]  [Classificação]  [Contratado] [EOM] [FM] [EOT] [FT] [EON] [FN]
    padrao = re.compile(
        r"^(?P<campo1>.+?)\s{2,}(?P<campo2>.+?)\s{2,}(?P<campo3>.+?)\s{2,}"
        r"(?P<n1>\d+)\s+(?P<n2>\d+)\s+(?P<n3>\d+)\s+(?P<n4>\d+)\s+(?P<n5>\d+)\s+(?P<n6>\d+)\s+(?P<n7>\d+)$"
    )

    def _classif_guess(t: str) -> str | None:
        t_up = t.upper()
        if "DIRETO" in t_up:
            return "Direto"
        if "INDIRETO" in t_up:
            return "Indireto"
        return None

    for l in linhas:
        m = padrao.match(l)
        if not m:
            registros.append({"raw_line": l})
            continue

        g = m.groupdict()
        c1, c2, c3 = g["campo1"], g["campo2"], g["campo3"]
        possiveis = [c1, c2, c3]

        classificacao = next((x for x in possiveis if _classif_guess(x)), None)
        frente = next((x for x in possiveis if "FRENTE" in x.upper()), None)
        funcao = next((x for x in possiveis if x not in {classificacao, frente}), None)

        # Defaults seguros
        if not classificacao:
            classificacao = "Direto" if "DIRETO" in l.upper() else ("Indireto" if "INDIRETO" in l.upper() else "")
        if not frente:
            frente = "FRENTE DE OBRA ÚNICA"
        if not funcao:
            funcao = c1  # fallback

        reg = {
            "Função": funcao.strip(),
            "Frente de Obra": frente.strip(),
            "Classificação": classificacao.strip(),
            "Contratado Geral": int(g["n1"]),
            "Em operação (manhã)": int(g["n2"]),
            "Fiscalizado (manhã)": int(g["n3"]),
            "Em operação (tarde)": int(g["n4"]),
            "Fiscalizado (tarde)": int(g["n5"]),
            "Em operação (noite)": int(g["n6"]),
            "Fiscalizado (noite)": int(g["n7"]),
        }
        registros.append(reg)

    return registros

# ---------- Pipeline ----------

def processar_arquivos(files) -> pd.DataFrame:
    linhas = []
    inconsistencias = []
    for f in files:
        try:
            texto = _texto_pdf(f.read())
            dados = _parse_bloco_mao_de_obra(texto)
            for row in dados:
                if "raw_line" in row:
                    inconsistencias.append({"Nome do Arquivo": f.name, "Linha": row["raw_line"]})
                else:
                    row["Nome do Arquivo"] = f.name
                    linhas.append(row)
        except Exception as e:
            inconsistencias.append({"Nome do Arquivo": f.name, "Linha": f"[ERRO] {e}"})

    df = pd.DataFrame(linhas)

    # Ordenação final de colunas exatamente como no print
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
    # Mantém extras (se surgirem) ao fim
    cols = [c for c in cols_ordem if c in df.columns] + [c for c in df.columns if c not in cols_ordem]
    df = df[cols] if not df.empty else pd.DataFrame(columns=cols_ordem)

    df_incons = pd.DataFrame(inconsistencias)
    return df, df_incons

# ---------- UI ----------

col1, col2 = st.columns([1, 2])
with col1:
    executar = st.button("🚀 Extrair", use_container_width=True, type="primary", disabled=not arquivos)
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
        with st.expander("Inconsistências (linhas que não aderiram ao padrão)", expanded=False):
            st.dataframe(df_incons, use_container_width=True, hide_index=True)

    # Excel para download (aba principal + inconsistências, se houver)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mao_de_Obra", index=False)
        if not df_incons.empty:
            df_incons.to_excel(writer, sheet_name="Inconsistencias", index=False)

    st.download_button(
        label="💾 Baixar Excel",
        data=buffer.getvalue(),
        file_name=f"{(nome_excel or 'rdo_consolidado').strip()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption("Se alguma linha cair em 'Inconsistências', envie um PDF de exemplo que eu ajusto o regex para cobrir o caso.")
