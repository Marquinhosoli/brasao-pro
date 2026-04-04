# THOTH PRO FINAL - CONVERSÃO REAL

import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
from io import BytesIO

st.set_page_config(layout="wide")

FORCE_UND = ["ABACAXI", "MELANCIA"]

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
    if fator == 0:
        return None
    return math.ceil(qtd / fator)

st.title("THOTH PRO FINAL")

base_file = st.file_uploader("Base", type=["xlsx"])
pdfs = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR"):

    base = pd.read_excel(base_file)
    base["PRODUTO"] = base["PRODUTO"].str.upper()

    resultado = {}
    log = []

    for pdf in pdfs:
        texto = extrair_texto(pdf)
        itens = extrair_itens(texto)

        for prod, qtd in itens:
            match = base[base["PRODUTO"].apply(lambda x: x in prod)]

            if match.empty:
                log.append([prod, qtd, "SEM BASE"])
                continue

            row = match.iloc[0]
            tipo = detectar_tipo(prod, row["TIPO"])
            fator = row["MEDIDA"]

            caixas = converter(qtd, fator)

            if caixas is None:
                continue

            resultado[prod] = resultado.get(prod, 0) + caixas
            log.append([prod, qtd, caixas])

    df = pd.DataFrame(list(resultado.items()), columns=["PRODUTO","CAIXAS"]).sort_values("PRODUTO")
    log_df = pd.DataFrame(log, columns=["PRODUTO","QTD","RESULTADO"])

    st.dataframe(df)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RESULTADO", index=False)
        log_df.to_excel(writer, sheet_name="LOG", index=False)

    st.download_button("BAIXAR", output.getvalue(), "resultado.xlsx")
