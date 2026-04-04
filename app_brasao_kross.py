import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
from io import BytesIO

st.set_page_config(layout="wide")

# ================= CONFIG =================
FORCE_UND = ["ABACAXI", "MELANCIA"]

# ================= FUNÇÕES =================

def extrair_texto(pdf):
    texto = ""
    with pdfplumber.open(pdf) as p:
        for page in p.pages:
            texto += (page.extract_text() or "") + "\n"
    return texto


def limpar_nome(nome):
    nome = nome.upper()
    nome = re.sub(r"\d.*", "", nome)
    nome = re.sub(r"(KG|UND|BDJ|BANDEJA|DE MARCHI|DEMARCHI|SHELF)", "", nome)
    return nome.strip()


def extrair_itens(texto):
    itens = []
    for linha in texto.split("\n"):
        if "KG" in linha or "UND" in linha or "BDJ" in linha:
            nums = re.findall(r"\d+[\.,]?\d*", linha)
            if not nums:
                continue
            qtd = float(nums[0].replace(",", "."))
            produto = limpar_nome(linha)
            itens.append((produto, qtd))
    return itens


def detectar_tipo(produto, tipo_base):
    for p in FORCE_UND:
        if p in produto:
            return "UND"
    return tipo_base


def converter(qtd, fator):
    if fator == 0 or pd.isna(fator):
        return None
    return math.ceil(qtd / fator)


# ================= APP =================

st.title("THOTH PRO FINAL (CORRIGIDO)")

base_file = st.file_uploader("Base (Excel)", type=["xlsx"])
pdfs = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR"):

    # 🔒 VALIDAÇÕES
    if base_file is None:
        st.error("Envie a base Excel antes de processar.")
        st.stop()

    if not pdfs:
        st.error("Envie pelo menos um PDF.")
        st.stop()

    # 🔥 CARREGAR BASE
    base = pd.read_excel(base_file)
    base.columns = [c.upper() for c in base.columns]

    if "PRODUTO" not in base.columns:
        st.error("A base precisa ter coluna PRODUTO")
        st.stop()

    base["PRODUTO"] = base["PRODUTO"].astype(str).str.upper()

    resultado = {}
    log = []

    # ================= PROCESSAMENTO =================
    for pdf in pdfs:
        texto = extrair_texto(pdf)
        itens = extrair_itens(texto)

        for prod, qtd in itens:

            match = base[base["PRODUTO"].apply(lambda x: x in prod)]

            if match.empty:
                log.append([prod, qtd, "SEM BASE"])
                continue

            row = match.iloc[0]

            tipo = detectar_tipo(prod, row.get("TIPO", "KG"))
            fator = row.get("MEDIDA", 0)

            caixas = converter(qtd, fator)

            if caixas is None:
                log.append([prod, qtd, "ERRO CONVERSÃO"])
                continue

            resultado[prod] = resultado.get(prod, 0) + caixas
            log.append([prod, qtd, caixas])

    # ================= RESULTADO =================

    df = pd.DataFrame(
        list(resultado.items()), columns=["PRODUTO", "CAIXAS"]
    ).sort_values("PRODUTO")

    log_df = pd.DataFrame(log, columns=["PRODUTO", "QTD_ORIGINAL", "RESULTADO"])

    st.subheader("Resultado")
    st.dataframe(df)

    st.subheader("Log")
    st.dataframe(log_df)

    # ================= EXPORT =================

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RESULTADO", index=False)
        log_df.to_excel(writer, sheet_name="LOG", index=False)

    st.download_button(
        "BAIXAR EXCEL FINAL",
        data=output.getvalue(),
        file_name="THOTH_RESULTADO.xlsx"
    )
