# app.py
import io
import re
import unicodedata
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator RDO", page_icon="🧰", layout="wide")
st.title("🧰 Extrator de RDO (PDF → Excel)")
st.caption("Extrai MÃO DE OBRA entre os blocos do RDO e gera planilha exatamente no layout desejado.")

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
    st.caption("Linhas fora do padrão vão para a aba **Inconsistencias**.")

# -------- Utils --------
def _texto_pdf(file_like: bytes) -> str:
    with fitz.open(stream=file_like, filetype="pdf") as doc:
        # texto simples por página
        return "\n".join(page.get_text() for page in doc)

def _norm(s: str) -> str:
    """Uppercase sem acento para buscas robustas."""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper()

def _recorta_bloco(texto: str) -> str | None:
    """Recorta o trecho entre MÃO DE OBRA e EQUIPAMENTO (tolerante a variações)."""
    tnorm = _norm(texto)
    anchors = [
        "RECURSOS EM OPERACAO MAO DE OBRA",
        "RECURSOS EM OPERACAO - MAO DE OBRA",
        "RECURSOS DE OPERACAO MAO DE OBRA",
    ]
    enders = [
        "RECURSOS EM OPERACAO EQUIPAMENTO",
        "RECURSOS EM OPERACAO - EQUIPAMENTO",
        "RECURSOS DE OPERACAO EQUIPAMENTO",
    ]
    start = -1
    for a in anchors:
        start = tnorm.find(a)
        if start != -1:
            break
    if start == -1:
        return None

    end = -1
    for b in enders:
        end = tnorm.find(b, start + 1)
        if end != -1:
            break
    if end == -1 or end <= start:
        return None

    # recorta usando os índices do texto normalizado
    # para manter os caracteres originais, fazemos proporção aproximada
    ratio = len(texto) / max(len(tnorm), 1)
    s0 = int(start * ratio)
    e0 = int(end * ratio)
    return texto[s0:e0]

# -------- Parser robusto --------
def _parse_bloco_mao_de_obra(texto: str) -> list[dict]:
    bloco = _recorta_bloco(texto)
    if not bloco:
        return []

    linhas = [l.strip() for l in bloco.splitlines() if l.strip()]

    # filtros de cabeçalho/total
    filtros = [
        r"^FUN(C|Ç)AO$",
        r"^FRENTE( DE OBR(A|A) .*|)$",
        r"^CLASSIFICA(C|Ç)AO$",
        r"^EM OPERACAO$",
        r"^FISCALIZADO$",
        r"^MANHA|^TARDE|^NOITE$",
        r"TOTAL",
    ]
    filtros_re = [re.compile(pat, re.IGNORECASE) for pat in filtros]

    def _eh_cabecalho(l: str) -> bool:
        ln = _norm(l)
        return any(r.search(ln) for r in filtros_re)

    registros = []
    for l in linhas:
        if _eh_cabecalho(l):
            continue

        # pega TODAS as ocorrências de números na linha
        nums_iter = list(re.finditer(r"\d+", l))
        if len(nums_iter) < 7:
            registros.append({"raw_line": l})
            continue

        # usa os 7 últimos números => contratado, EOM, FM, EOT, FT, EON, FN
        last7 = nums_iter[-7:]
        # início do primeiro desses 7 números => separa texto/nums
        cut = last7[0].start()
        texto_esq = l[:cut].rstrip()

        # números na ordem desejada
        n = [int(m.group()) for m in last7]
        contratado, eom, fm, eot, ft, eon, fn = n

        # quebra os 3 campos textuais por blocos de 2+ espaços/tabs
        partes = [p.strip() for p in re.split(r"[ \t]{2,}", texto_esq) if p.strip()]

        classificacao = ""
        frente = ""
        funcao = ""

        # heurísticas
        for p in partes:
            up = _norm(p)
            if not classificacao and "DIRETO" in up:
                classificacao = "Direto"
                continue
            if not classificacao and "INDIRETO" in up:
                classificacao = "Indireto"
                continue
            if not frente and "FRENTE" in up:
                frente = p
                continue

        # função = o restante “mais descritivo”
        restantes = [p for p in partes if p not in {classificacao, frente} and p]
        if restantes:
            # pega o mais longo como função
            funcao = max(restantes, key=len)
        else:
            # fallback: primeira parte da linha
            funcao = partes[0] if partes else texto_esq

        if not frente:
            frente = "FRENTE DE OBRA ÚNICA"

        reg = {
            "Função": funcao,
            "Frente de Obra": frente,
            "Classificação": classificacao,
            "Contratado Geral": contratado,
            "Em operação (manhã)": eom,
            "Fiscalizado (manhã)": fm,
            "Em operação (tarde)": eot,
            "Fiscalizado (tarde)": ft,
            "Em operação (noite)": eon,
            "Fiscalizado (noite)": fn,
        }
        registros.append(reg)

    return registros

# -------- Pipeline --------
def processar_arquivos(files):
    linhas, inconsistencias = [], []
    for f in files:
        try:
            texto = _texto_pdf(f.read())
            dados = _parse_bloco_mao_de_obra(texto)
            if not dados:
                inconsistencias.append({"Nome do Arquivo": f.name, "Linha": "[BLOCO NÃO ENCONTRADO]"})
                continue
            for row in dados:
                if "raw_line" in row:
                    inconsistencias.append({"Nome do Arquivo": f.name, "Linha": row["raw_line"]})
                else:
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
    st.subheader("Prévia dos dados")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df_incons.empty:
        with st.expander("Inconsistências / linhas não parseadas", expanded=False):
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
st.caption("Se ainda vier vazio, me manda 1 PDF de exemplo (sem dados sensíveis) que eu ajusto o parser exatamente para o seu layout.")
