import io
import math
import re
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st
from fpdf import FPDF

st.set_page_config(page_title="THOTH PRO FINAL", layout="wide")

# =========================
# ESTILO / LAYOUT
# =========================
st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}
h1 {
    font-weight: 800 !important;
    letter-spacing: -0.5px;
}
.stButton > button {
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    font-weight: 600;
}
.result-card {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 16px 18px;
    margin-bottom: 12px;
}
.small-muted {
    color: #6b7280;
    font-size: 0.92rem;
}
</style>
""", unsafe_allow_html=True)

st.title("🚀 THOTH PRO FINAL (PDF + EXCEL)")
st.write("Upload dos pedidos (PDF ou Excel)")

files = st.file_uploader(
    "Envie os arquivos",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True
)

# =========================
# BASE DE CONVERSÃO
# =========================
BASE_PRODUTOS = {
    "ABACAXI": {"modo": "un", "por_caixa": 1, "grupo": "FRUTAS"},
    "MELANCIA": {"modo": "un", "por_caixa": 1, "grupo": "FRUTAS"},
    "MELAO": {"modo": "un", "por_caixa": 10, "grupo": "FRUTAS"},
    "MAMAO PAPAYA": {"modo": "un", "por_caixa": 18, "grupo": "FRUTAS"},
    "LARANJA": {"modo": "kg", "por_caixa": 20, "grupo": "FRUTAS"},
    "LIMAO": {"modo": "kg", "por_caixa": 20, "grupo": "FRUTAS"},
    "MANGA": {"modo": "kg", "por_caixa": 12, "grupo": "FRUTAS"},
    "UVA 500G": {"modo": "bdj", "por_caixa": 10, "grupo": "FRUTAS"},
    "MORANGO": {"modo": "bdj", "por_caixa": 4, "grupo": "FRUTAS"},
    "MIRTILO": {"modo": "bdj", "por_caixa": 12, "grupo": "FRUTAS"},
    "FRAMBOESA": {"modo": "bdj", "por_caixa": 10, "grupo": "FRUTAS"},
    "TOMATE GRAPE 180G": {"modo": "bdj", "por_caixa": 24, "grupo": "LEGUMES"},
    "BATATA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "CEBOLA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "REPOLHO ROXO": {"modo": "kg", "por_caixa": 10, "grupo": "LEGUMES"},
    "MILHO": {"modo": "un", "por_caixa": 30, "grupo": "LEGUMES"},
    "MAXIXE": {"modo": "kg", "por_caixa": 12, "grupo": "LEGUMES"},
    "CARAMBOLA": {"modo": "bdj", "por_caixa": 4, "grupo": "FRUTAS"},
    "BLUEBERRY": {"modo": "bdj", "por_caixa": 12, "grupo": "FRUTAS"},
    "FIGO ROXO": {"modo": "un", "por_caixa": 1, "grupo": "FRUTAS"},
    "PHYSALIS IMPORTADO": {"modo": "bdj", "por_caixa": 8, "grupo": "FRUTAS"},
}

# =========================
# FUNÇÕES
# =========================
def normalizar_nome(texto: str) -> str:
    texto = (texto or "").upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace("Ç", "C").replace("Ã", "A").replace("Á", "A").replace("À", "A")
    texto = texto.replace("É", "E").replace("Ê", "E").replace("Í", "I")
    texto = texto.replace("Ó", "O").replace("Õ", "O").replace("Ô", "O").replace("Ú", "U")
    return texto

def detectar_cliente(nome_arquivo: str) -> str:
    n = normalizar_nome(nome_arquivo)
    if "KROSS" in n:
        return "KROSS"
    if "CD" in n:
        return "BRASAO CD"
    return "BRASAO"

def extrair_texto_pdf(uploaded_file) -> str:
    texto = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip():
                texto.append(t)
    return "\n".join(texto)

def extrair_linhas_relevantes(texto: str):
    linhas = []
    for linha in texto.splitlines():
        limpa = linha.strip()
        if limpa:
            linhas.append(limpa)
    return linhas

def parse_linha_produto(linha: str):
    """
    Parser Flexível com suporte nativo ao layout do ERP Flex (Brasão/Kross).
    Lê a linha normal de pedido e também a tabela de PENDÊNCIAS.
    """
    l = normalizar_nome(linha)

    # 1. PADRÃO ERP FLEX (PEDIDO NORMAL)
    m_flex = re.search(r"^\s*\d+\s+\d+\s+(.*?)\s+(\d+[.,]\d+)\s+(\d+)\s+\d+[.,]\d+\s+\d+[.,]\d+\s*$", l)
    if m_flex:
        descricao = m_flex.group(1)
        qtd = float(m_flex.group(3))
        
        m_un = re.search(r"\b(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS|CX|CXS|CAIXA|CAIXAS|VOL|VOLUME|VOLUMES)\b", descricao)
        un_raw = m_un.group(1) if m_un else "CX"

        produto = re.sub(r"\b(BRASAO FRUTA|DE MARCHI|SHELF \d+)\b", "", descricao).strip()
        produto = normalizar_nome(produto)
        
        if un_raw in ["KG", "KGS", "QUILO", "QUILOS"]: unidade = "kg"
        elif un_raw in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES"]: unidade = "un"
        elif un_raw in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"]: unidade = "cx"
        else: unidade = "bdj"
            
        return produto, qtd, unidade

    # 2. PADRÃO ERP FLEX (TABELA DE PENDÊNCIAS)
    # Ex: "003 003 TC 130915 MELANCIA INTEIRA KG BRASAO FRU0000000032162 97,035 184,37"
    m_pend = re.search(r"^\s*\d{3}\s+\d{3}\s+[A-Z]{2}\s+\d+\s+(.*?)\s+(\d+[.,]\d+)\s+\d+[.,]\d+\s*$", l)
    if m_pend:
        descricao = m_pend.group(1)
        qtd = float(m_pend.group(2).replace(",", "."))
        
        m_un = re.search(r"\b(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS|CX|CXS|CAIXA|CAIXAS|VOL|VOLUME|VOLUMES)\b", descricao)
        un_raw = m_un.group(1) if m_un else "CX"
        
        # Limpa o código de barras colado na marca (ex: Fru0000000032162)
        produto = re.sub(r"\d{10,}", "", descricao)
        produto = re.sub(r"\b(BRASAO FRU\w*|BRASAO FRUTA|DE MARCHI|SHELF \d+)\b", "", produto).strip()
        produto = normalizar_nome(produto)
        
        if un_raw in ["KG", "KGS", "QUILO", "QUILOS"]: unidade = "kg"
        elif un_raw in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES"]: unidade = "un"
        elif un_raw in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"]: unidade = "cx"
        else: unidade = "bdj"
            
        return produto, qtd, unidade

    # 3. PADRÃO GENÉRICO (Fallback)
    padrao = r"(\d+[.,]?\d*)\s*(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS|CX|CXS|CAIXA|CAIXAS|VOL|VOLUME|VOLUMES)\b"
    m = re.search(padrao, l)

    if m:
        qtd = float(m.group(1).replace(",", "."))
        un_raw = m.group(2)
        
        texto_restante = l.replace(m.group(0), " ")
        texto_restante = re.sub(r"^\s*\d+\s*[-:]?\s*", "", texto_restante)
        texto_restante = re.sub(r"\s+\d+[.,]\d+\s*.*$", "", texto_restante)
        
        produto = normalizar_nome(texto_restante)
        
        if not produto:
            return None

        if un_raw in ["KG", "KGS", "QUILO", "QUILOS"]: unidade = "kg"
        elif un_raw in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES"]: unidade = "un"
        elif un_raw in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"]: unidade = "cx"
        else: unidade = "bdj"

        return produto, qtd, unidade

    return None

        if un_raw in ["KG", "KGS", "QUILO", "QUILOS"]:
            unidade = "kg"
        elif un_raw in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES"]:
            unidade = "un"
        elif un_raw in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"]:
            unidade = "cx"
        else:
            unidade = "bdj"

        return produto, qtd, unidade

    return None

def localizar_base(produto: str):
    p = normalizar_nome(produto)
    if p in BASE_PRODUTOS:
        return p, BASE_PRODUTOS[p]
    for chave in BASE_PRODUTOS:
        if p == chave or p in chave or chave in p:
            return chave, BASE_PRODUTOS[chave]
    return p, None

def converter_para_caixa(produto: str, quantidade: float, unidade_encontrada: str):
    nome_base, info = localizar_base(produto)

    if not info:
        return {
            "produto_final": nome_base,
            "grupo": "NAO_IDENTIFICADO",
            "qtd_original": quantidade,
            "unidade_original": unidade_encontrada,
            "qtd_caixa": math.ceil(quantidade) if unidade_encontrada == "cx" else quantidade,
            "observacao": "SEM_BASE"
        }

    por_caixa = float(info["por_caixa"])
    grupo = info["grupo"]

    if por_caixa <= 0:
        return {
            "produto_final": nome_base,
            "grupo": grupo,
            "qtd_original": quantidade,
            "unidade_original": unidade_encontrada,
            "qtd_caixa": math.ceil(quantidade) if unidade_encontrada == "cx" else quantidade,
            "observacao": "BASE_INVALIDA"
        }

    if unidade_encontrada == "cx":
        qtd_caixa = math.ceil(quantidade)
    else:
        qtd_caixa = math.ceil(quantidade / por_caixa)

    return {
        "produto_final": nome_base,
        "grupo": grupo,
        "qtd_original": quantidade,
        "unidade_original": unidade_encontrada,
        "qtd_caixa": qtd_caixa,
        "observacao": ""
    }

def processar_arquivo(uploaded_file):
    nome = uploaded_file.name
    cliente = detectar_cliente(nome)

    if nome.lower().endswith(".pdf"):
        texto = extrair_texto_pdf(uploaded_file)
        linhas = extrair_linhas_relevantes(texto)
    else:
        df_excel = pd.read_excel(uploaded_file)
        linhas = [" ".join([str(x) for x in row if pd.notna(x)]) for _, row in df_excel.iterrows()]

    itens = []
    ignoradas = []

    for linha in linhas:
        item = parse_linha_produto(linha)
        if item:
            produto, qtd, unidade = item
            conv = converter_para_caixa(produto, qtd, unidade)
            conv["cliente"] = cliente
            conv["arquivo"] = nome
            itens.append(conv)
        else:
            ignoradas.append(linha)

    return itens, ignoradas

def gerar_excel_em_memoria(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for cliente in sorted(df["cliente"].unique()):
            df_cliente = df[df["cliente"] == cliente].copy()

            df_frutas = df_cliente[df_cliente["grupo"] == "FRUTAS"][
                ["produto_final", "qtd_caixa", "qtd_original", "unidade_original", "observacao"]
            ].rename(columns={
                "produto_final": "Produto",
                "qtd_caixa": "Qtd CX (Final)",
                "qtd_original": "Qtd Original",
                "unidade_original": "Unidade",
                "observacao": "Obs"
            })

            df_legumes = df_cliente[df_cliente["grupo"] == "LEGUMES"][
                ["produto_final", "qtd_caixa", "qtd_original", "unidade_original", "observacao"]
            ].rename(columns={
                "produto_final": "Produto",
                "qtd_caixa": "Qtd CX (Final)",
                "qtd_original": "Qtd Original",
                "unidade_original": "Unidade",
                "observacao": "Obs"
            })

            df_outros = df_cliente[~df_cliente["grupo"].isin(["FRUTAS", "LEGUMES"])][
                ["produto_final", "grupo", "qtd_caixa", "qtd_original", "unidade_original", "observacao"]
            ].rename(columns={
                "produto_final": "Produto",
                "grupo": "Grupo",
                "qtd_caixa": "Qtd CX (Final)",
                "qtd_original": "Qtd Original",
                "unidade_original": "Unidade",
                "observacao": "Obs"
            })

            if not df_frutas.empty:
                df_frutas.to_excel(writer, sheet_name=f"{cliente[:20]} FRUTAS", index=False)
            if not df_legumes.empty:
                df_legumes.to_excel(writer, sheet_name=f"{cliente[:20]} LEGUMES", index=False)
            if not df_outros.empty:
                df_outros.to_excel(writer, sheet_name=f"{cliente[:20]} OUTROS", index=False)

        resumo = df.groupby(["cliente", "grupo"], dropna=False)["qtd_caixa"].sum().reset_index()
        resumo.to_excel(writer, sheet_name="RESUMO", index=False)

    output.seek(0)
    return output

def gerar_pdf_em_memoria(df):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)

    for cliente in sorted(df["cliente"].unique()):
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 10, f"Relatorio - {cliente}", ln=True)

        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 8, "THOTH PRO FINAL", ln=True)
        pdf.ln(4)

        for grupo in ["FRUTAS", "LEGUMES", "NAO_IDENTIFICADO"]:
            bloco = df[df["cliente"] == cliente]
            bloco = bloco[bloco["grupo"] == grupo]

            if bloco.empty:
                continue

            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 8, grupo, ln=True)

            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(95, 7, "Produto", border=1)
            pdf.cell(25, 7, "Qtd CX", border=1, align="C")
            pdf.cell(30, 7, "Original", border=1, align="C")
            pdf.cell(25, 7, "Unid", border=1, align="C")
            pdf.cell(15, 7, "Obs", border=1, align="C")
            pdf.ln()

            pdf.set_font("Helvetica", "", 8)
            for _, row in bloco.iterrows():
                produto = str(row["produto_final"])[:42]
                pdf.cell(95, 7, produto, border=1)
                pdf.cell(25, 7, str(row["qtd_caixa"]), border=1, align="C")
                pdf.cell(30, 7, str(row["qtd_original"]), border=1, align="C")
                pdf.cell(25, 7, str(row["unidade_original"]), border=1, align="C")
                pdf.cell(15, 7, str(row["observacao"])[:8], border=1, align="C")
                pdf.ln()

            pdf.ln(4)

    return bytes(pdf.output(dest="S"))

# =========================
# BOTÃO PROCESSAR
# =========================
if st.button("🔥 PROCESSAR PEDIDOS", use_container_width=False):
    if not files:
        st.warning("Envie pelo menos um arquivo.")
        st.stop()

    todos_itens = []
    todas_ignoradas = []

    with st.spinner("Lendo PDFs e processando pedidos..."):
        for f in files:
            try:
                itens, ignoradas = processar_arquivo(f)
                todos_itens.extend(itens)
                todas_ignoradas.extend([(f.name, x) for x in ignoradas[:20]])
            except Exception as e:
                st.error(f"Erro ao processar {f.name}: {e}")

    if not todos_itens:
        st.error("Nenhum item foi reconhecido. Isso significa que o formato do PDF não bateu com o parser atual.")
        st.info("Nesse caso, eu recomendo ajustar o parser exatamente no padrão dos seus PDFs.")
        st.stop()

    df = pd.DataFrame(todos_itens)
    df = df.sort_values(["cliente", "grupo", "produto_final"]).reset_index(drop=True)

    st.success("Pedidos processados com sucesso.")

    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="result-card"><b>Arquivos</b><br>{len(files)}</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="result-card"><b>Itens reconhecidos</b><br>{len(df)}</div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="result-card"><b>Sem base</b><br>{int((df["observacao"] == "SEM_BASE").sum())}</div>', unsafe_allow_html=True)

    st.subheader("Prévia do resultado")
    st.dataframe(df, use_container_width=True)

    excel_bytes = gerar_excel_em_memoria(df)
    pdf_bytes = gerar_pdf_em_memoria(df)

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Baixar Excel",
            data=excel_bytes,
            file_name="thoth_pro_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        st.download_button(
            "📄 Baixar PDF",
            data=pdf_bytes,
            file_name="thoth_pro_final.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    if todas_ignoradas:
        with st.expander("Linhas ignoradas para conferência"):
            for arq, linha in todas_ignoradas[:50]:
                st.write(f"**{arq}** → {linha}")

else:
    st.markdown('<div class="small-muted">Layout atualizado e alinhamento corrigido. O motor agora procura a quantidade em qualquer lugar da linha.</div>', unsafe_allow_html=True)
