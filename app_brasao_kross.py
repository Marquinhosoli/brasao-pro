import streamlit as st
import pandas as pd
from fpdf import FPDF

st.set_page_config(layout="wide")

st.title("🚀 THOTH PRO FINAL (PDF + EXCEL)")

st.markdown("Upload dos pedidos (PDF ou Excel)")

files = st.file_uploader(
    "Envie os arquivos",
    accept_multiple_files=True
)

# ------------------------------
# BASE SIMPLES (EXEMPLO)
# ------------------------------
BASE = {
    "LARANJA": {"tipo": "kg", "cx": 20},
    "LIMAO": {"tipo": "kg", "cx": 20},
    "ABACAXI": {"tipo": "un", "cx": 1},
    "MELANCIA": {"tipo": "un", "cx": 1},
    "MELAO": {"tipo": "un", "cx": 10},
}

# ------------------------------
# CONVERSÃO
# ------------------------------
def converter(produto, qtd):
    p = produto.upper()

    if p in BASE:
        info = BASE[p]

        if info["tipo"] == "kg":
            return round(qtd / info["cx"], 2)

        elif info["tipo"] == "un":
            return round(qtd / info["cx"], 2)

    return qtd

# ------------------------------
# GERAR PDF
# ------------------------------
def gerar_pdf(df, nome):
    pdf = FPDF()
    pdf.add_page()

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Relatório - {nome}", ln=True)

    for i, row in df.iterrows():
        linha = f"{row['Produto']} - {row['Qtd CX']} cx"
        pdf.cell(200, 8, txt=linha, ln=True)

    caminho = f"{nome}.pdf"
    pdf.output(caminho)

    return caminho

# ------------------------------
# PROCESSAR
# ------------------------------
if st.button("🔥 PROCESSAR PEDIDOS"):

    dados = []

    # ⚠️ SIMULAÇÃO (depois ligamos no parser real)
    exemplo = [
        ("Laranja", 400),
        ("Limão", 1000),
        ("Abacaxi", 200),
    ]

    for prod, qtd in exemplo:
        cx = converter(prod, qtd)

        dados.append({
            "Produto": prod,
            "Qtd CX": cx
        })

    df = pd.DataFrame(dados)

    st.success("Processado com sucesso!")

    st.dataframe(df)

    # Excel
    excel_file = "resultado.xlsx"
    df.to_excel(excel_file, index=False)

    # PDF
    pdf_file = gerar_pdf(df, "resultado")

    col1, col2 = st.columns(2)

    with col1:
        with open(excel_file, "rb") as f:
            st.download_button("📥 Baixar Excel", f, file_name=excel_file)

    with col2:
        with open(pdf_file, "rb") as f:
            st.download_button("📄 Baixar PDF", f, file_name=pdf_file)
