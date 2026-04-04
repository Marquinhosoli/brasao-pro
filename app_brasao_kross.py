import io
import math
import re

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="THOTH PRO FINAL", layout="wide")

# =========================
# ESTILO / LAYOUT
# =========================
st.markdown("""
<style>
.block-container { padding-top: 2rem; padding-bottom: 2rem; }
h1 { font-weight: 800 !important; letter-spacing: -0.5px; }
.stButton > button { border-radius: 10px; padding: 0.6rem 1.2rem; font-weight: 600; }
.result-card { background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 16px 18px; margin-bottom: 12px; }
</style>
""", unsafe_allow_html=True)

st.title("🚀 THOTH PRO FINAL (PDF + EXCEL)")
st.write("Layout Rigoroso Thoth e Conversão Automática de Caixas")

files = st.file_uploader(
    "Envie os PDFs de pedidos",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True
)

# =========================
# BASE DE CONVERSÃO EXATA (Extraída da sua planilha)
# =========================
BASE_PRODUTOS = {
    "ABACATE KG": {"por_caixa": 20, "grupo": "FRUTAS"},
    "ABACAXI PEROLA UND": {"por_caixa": 1, "grupo": "FRUTAS"},
    "ABOBORA PESCOCO KG": {"por_caixa": 1, "grupo": "LEGUMES"},
    "ALECRIM MACO": {"por_caixa": 4, "grupo": "LEGUMES"},
    "ALHO PORO UND": {"por_caixa": 12, "grupo": "LEGUMES"},
    "AMEIXA NACIONAL DEMARCHI BDJ 500G SHELF 30": {"por_caixa": 30, "grupo": "FRUTAS"},
    "BATATA DOCE BRANCA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA DOCE ROXA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA SALSA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "BERINJELA KG": {"por_caixa": 10, "grupo": "LEGUMES"},
    "BETERRABA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "CAQUI RAMA FORTE KG": {"por_caixa": 6, "grupo": "FRUTAS"},
    "CARAMBOLA DE MARCHI 400G": {"por_caixa": 4, "grupo": "FRUTAS"},
    "CEBOLA ARGENTINA BRANCA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "CEBOLA CONSERVA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "CENOURA KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "CHUCHU KG": {"por_caixa": 20, "grupo": "LEGUMES"},
    "COCO SECO FRUTA KG": {"por_caixa": 20, "grupo": "FRUTAS"},
    "COENTRO MACO": {"por_caixa": 10, "grupo": "LEGUMES"},
    "FIGO ROXO DE MARCHI 300G": {"por_caixa": 1, "grupo": "FRUTAS"},
    "FRAMBOESA FRUTA 120G SHELF 15": {"por_caixa": 15, "grupo": "FRUTAS"},
    "GOIABA NACIONAL VERMELHA KG": {"por_caixa": 20, "grupo": "FRUTAS"},
    "HORTELA MACO": {"por_caixa": 10, "grupo": "LEGUMES"},
    "JATOBA FRUTA KG": {"por_caixa": 1, "grupo": "FRUTAS"},
    "KINKAN BANDEJA FRUTAMINA 500G": {"por_caixa": 10, "grupo": "FRUTAS"},
    "KIWI IMPORTADO GRECIA KG": {"por_caixa": 10, "grupo": "FRUTAS"},
    "KIWI NACIONAL DE MARCHI BANDEJA 600G SHELF 15": {"por_caixa": 15, "grupo": "FRUTAS"},
    "LARANJA MAQUINA DE SUCO": {"por_caixa": 20, "grupo": "FRUTAS"},
    "LIMAO SICILIANO KG": {"por_caixa": 15, "grupo": "FRUTAS"},
    "LIMAO TAHITI KG": {"por_caixa": 20, "grupo": "FRUTAS"},
    "LOURO MACO": {"por_caixa": 10, "grupo": "LEGUMES"},
    "MACA FUJI CAT 1 KG": {"por_caixa": 18, "grupo": "FRUTAS"},
    "MAMAO FORMOSA KG": {"por_caixa": 15, "grupo": "FRUTAS"},
    "MAMAOZINHO PAPAIA UNIDADE": {"por_caixa": 18, "grupo": "FRUTAS"},
    "MANGA PALMER KG": {"por_caixa": 12, "grupo": "FRUTAS"},
    "MANJERICAO MACO": {"por_caixa": 10, "grupo": "LEGUMES"},
    "MANJERONA MACO": {"por_caixa": 10, "grupo": "LEGUMES"},
    "MAXIXE BDJ DE MARCHI 300G": {"por_caixa": 12, "grupo": "LEGUMES"},
    "MELAO CANTALOUPE UNIDADE": {"por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO CHARANTEAIS KG": {"por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO DINO KG": {"por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO ESPANHOL AMARELO KG": {"por_caixa": 13, "grupo": "FRUTAS"},
    "MELAO GALIA UNIDADE": {"por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO ORANGE UNIDADE": {"por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO REI DOCE REDINHA KG": {"por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO SAPO KG": {"por_caixa": 13, "grupo": "FRUTAS"},
    "MELANCIA INTEIRA KG": {"por_caixa": 1, "grupo": "FRUTAS"},
    "MILHO VERDE ESPIGA DE MARCHI BDJ 700G SHELF 10": {"por_caixa": 10, "grupo": "LEGUMES"},
    "MIRTILO BLUEBERRY IMP. DEMARCHI 125G": {"por_caixa": 12, "grupo": "FRUTAS"},
    "NABO UNIDADE": {"por_caixa": 6, "grupo": "LEGUMES"},
    "PEPINO JAPONES KG": {"por_caixa": 18, "grupo": "LEGUMES"},
    "PERA WILLIANS ARGENTINA KG": {"por_caixa": 18, "grupo": "FRUTAS"},
    "PESSEGO IMP ARGENTINA POLPA AMARELA KG": {"por_caixa": 10, "grupo": "FRUTAS"},
    "PHYSALIS IMPORTADO COLOMBIA 100G": {"por_caixa": 8, "grupo": "FRUTAS"},
    "PIMENTA BIQUINHO KG": {"por_caixa": 1, "grupo": "LEGUMES"},
    "PIMENTA CAMBUCI KG": {"por_caixa": 1, "grupo": "LEGUMES"},
    "PIMENTA JALAPENO KG": {"por_caixa": 1, "grupo": "LEGUMES"},
    "SALSAO AIPO UNIDADE": {"por_caixa": 1, "grupo": "LEGUMES"},
    "SALVIA UNIDADE": {"por_caixa": 1, "grupo": "LEGUMES"},
    "TOMATE GRAPE DEMARCHI 180G SHELF 10": {"por_caixa": 10, "grupo": "LEGUMES"},
    "TOMILHO MACO": {"por_caixa": 1, "grupo": "LEGUMES"},
    "UVA THOMPSON S/SEMENTE DEMARCHI BDJ 500G": {"por_caixa": 10, "grupo": "FRUTAS"},
}

# =========================
# CABEÇALHOS DO THOTH EXATOS
# =========================
configs = [
    ("BRASAO", "FRUTAS", "BRASAO - FRUTAS PRE PEDIDO BRANCO.xlsx", ["1", "2", "3", "4"],
     ["BRASAO FRUTAS", "LOJA 1 CE", "LOJA 2 JÁ", "LOJA 3 XX", "LOJA 4 AV", "", "DATA ENTREGA"],
     ["PRODUTO", "1", "2", "3", "4", "TOTAL", ""]),
     
    ("BRASAO", "LEGUMES", "BRASAO - LEGUMES PRE PEDIDO BRANCO.xlsx", ["1", "2", "3", "4"],
     ["BRASAO LEGUMES", "LOJA 1 CE", "LOJA 2 JÁ", "LOJA 3 XX", "LOJA 4 AV", "", "DATA ENTREGA"],
     ["PRODUTO", "1", "2", "3", "4", "TOTAL", ""]),
     
    ("KROSS", "FRUTAS", "KROSS - FRUTAS PRE PEDIDO BRANCO.xlsx", ["1", "2"],
     ["KROSS", "KROSS ATACADISTA", "KROSS XAXIM", ""],
     ["PRODUTO", "1", "2", "TOTAL"]),
     
    ("KROSS", "LEGUMES", "KROSS - LEGUMES PRE PEDIDO BRANCO.xlsx", ["1", "2"],
     ["KROSS - LEGUMES", "KROSS ATACADISTA", "KROSS XAXIM", ""],
     ["PRODUTO", "1", "2", "TOTAL"]),
     
    ("BRASAO CD", "FRUTAS", "BRASAO CD - FRUTAS PRE PEDIDO BRANCO.xlsx", ["1"],
     ["BRASAO CD - FRUTAS", "BRASAO CD", ""],
     ["PRODUTO", "1", "TOTAL"]),
     
    ("BRASAO CD", "LEGUMES", "BRASAO CD - LEGUMES PRE PEDIDO BRANCO.xlsx", ["1"],
     ["BRASAO CD - LEGUMES", "BRASAO CD", ""],
     ["PRODUTO", "1", "TOTAL"]),
]

# =========================
# FUNÇÕES CORE
# =========================
def identificar_loja(nome_arquivo: str):
    n = nome_arquivo.upper()
    if "KROSS" in n and "XAXIM" in n: return "KROSS", "2"
    if "KROSS" in n: return "KROSS", "1"
    if "CD" in n: return "BRASAO CD", "1"
    if "FERNANDO" in n: return "BRASAO", "1"
    if "JARDIM" in n: return "BRASAO", "2"
    if "XAXIM" in n: return "BRASAO", "3"
    if "AVENIDA" in n: return "BRASAO", "4"
    return "OUTROS", "0"

def parse_linha_produto(linha: str):
    l = linha.strip()
    
    if any(x in l.upper() for x in ["TOTAL", "PESO", "FRETE", "VALOR"]):
        return None

    m_flex = re.search(r"^(.*?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$", l)
    if m_flex:
        descricao_bruta = m_flex.group(1).strip()
        
        qtd_str = m_flex.group(2).replace(".", "").replace(",", ".")
        try:
            qtd = float(qtd_str)
        except ValueError:
            return None
            
        m_desc = re.search(r"[A-Za-z].*$", descricao_bruta)
        if not m_desc: return None
            
        produto = m_desc.group(0).upper()
        
        # Recuperando a unidade (para saber se o PDF já pediu direto em Caixa)
        m_un = re.search(r"\b(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS|CX|CXS|CAIXA|CAIXAS|VOL|VOLUME|VOLUMES|MACO|MACOS)\b", produto)
        unidade_encontrada = "cx" if m_un and m_un.group(1) in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"] else "outros"
        
        for m in ["BRASAO FRUTA", "DE MARCHI", "FRUTAMINA"]:
            if produto.endswith(m):
                produto = produto[:-len(m)].strip()
        
        return produto, qtd, unidade_encontrada

    return None

def localizar_base(produto: str):
    p = produto.upper()
    if p in BASE_PRODUTOS:
        return p, BASE_PRODUTOS[p]
    for chave in BASE_PRODUTOS:
        if chave in p or p in chave:
            return chave, BASE_PRODUTOS[chave]
    return p, None

def converter_para_final(produto: str, quantidade_original: float, unidade_encontrada: str):
    nome_base, info = localizar_base(produto)

    if not info:
        return {
            "produto_final": produto,
            "grupo": "NAO_IDENTIFICADO",
            "qtd_original": quantidade_original,
            "qtd_final": math.ceil(quantidade_original),
            "observacao": "PRODUTO NOVO"
        }

    por_caixa = info.get("por_caixa")
    grupo = info["grupo"]

    # ======= CÁLCULO MATEMÁTICO REAL =======
    if unidade_encontrada == "cx":
        # Se no PDF já estiver escrito "CX", mantém.
        qtd_final = math.ceil(quantidade_original)
        obs = "Já em CX"
    elif por_caixa and float(por_caixa) > 0:  
        # Divide pelo peso da caixa e arredonda pra cima (Ex: 360kg / 18 = 20 Caixas)
        qtd_final = math.ceil(quantidade_original / float(por_caixa))
        obs = f"Divisão por {por_caixa}"
    else:           
        # Fator de segurança
        qtd_final = math.ceil(quantidade_original)
        obs = "Base = 1"

    return {
        "produto_final": nome_base,
        "grupo": grupo,
        "qtd_original": quantidade_original,
        "qtd_final": qtd_final,
        "observacao": obs
    }

def processar_arquivo(uploaded_file):
    nome = uploaded_file.name
    cliente, loja_num = identificar_loja(nome)

    texto = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip(): texto.append(t)
            
    linhas_texto = "\n".join(texto).splitlines()

    itens = []
    for l in linhas_texto:
        limpa = l.strip()
        if not limpa: continue
        
        if "PENDENCIAS DE MERCADORIAS" in limpa.upper():
            break 
            
        item = parse_linha_produto(limpa)
        if item:
            produto, qtd, unid = item
            conv = converter_para_final(produto, qtd, unid)
            conv["cliente"] = cliente
            conv["loja_cod"] = loja_num
            conv["arquivo"] = nome
            itens.append(conv)

    return itens

# =========================
# GERAÇÃO THOTH EXCEL
# =========================
def gerar_planilha_thoth(df_itens, cliente, grupo, colunas_numericas, sub_head):
    df_filtro = df_itens[(df_itens["cliente"] == cliente) & (df_itens["grupo"] == grupo)]
    if df_filtro.empty: return pd.DataFrame() 

    pivot = pd.pivot_table(
        df_filtro, values='qtd_final', index='produto_final',
        columns='loja_cod', aggfunc='sum', fill_value=0
    ).reset_index()

    pivot.rename(columns={'produto_final': 'PRODUTO'}, inplace=True)

    for col in colunas_numericas:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot["TOTAL"] = pivot[colunas_numericas].sum(axis=1)
    
    current_cols = ["PRODUTO"] + colunas_numericas + ["TOTAL"]
    for i in range(len(sub_head) - len(current_cols)):
        col_nome = f"EXTRA_{i}"
        pivot[col_nome] = ""
        current_cols.append(col_nome)
        
    pivot = pivot[current_cols].sort_values(by="PRODUTO").reset_index(drop=True)
    pivot = pivot[pivot["TOTAL"] > 0] 

    return pivot

def gerar_arquivos_excel(df):
    arquivos = {}

    for cliente, grupo, nome_arquivo, colunas, sup_head, sub_head in configs:
        df_gerado = gerar_planilha_thoth(df, cliente, grupo, colunas, sub_head)
        
        if not df_gerado.empty:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("Plan1")
                
                for col_i, val in enumerate(sup_head): worksheet.write(0, col_i, val)
                for col_i, val in enumerate(sub_head): worksheet.write(1, col_i, val)
                    
                for row_i, row_data in enumerate(df_gerado.values):
                    for col_i, val in enumerate(row_data):
                        worksheet.write(row_i + 2, col_i, val)

            arquivos[nome_arquivo] = output.getvalue()

    df_erros = df[df["grupo"] == "NAO_IDENTIFICADO"]
    if not df_erros.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_erros[["produto_final", "qtd_original", "loja_cod", "arquivo"]].rename(columns={
                "produto_final": "NOME DO PRODUTO (Cadastre na Base)",
                "qtd_original": "Quantidade Recebida",
                "loja_cod": "Loja Número",
                "arquivo": "PDF"
            }).to_excel(writer, index=False, sheet_name="ITENS REJEITADOS")
        arquivos["ITENS_NAO_RECONHECIDOS.xlsx"] = output.getvalue()

    return arquivos

# =========================
# BOTÃO PROCESSAR E BAIXAR
# =========================
if st.button("🔥 PROCESSAR PEDIDOS E GERAR MATRIZ THOTH", use_container_width=False):
    if not files:
        st.warning("Envie pelo menos um arquivo de pedido.")
        st.stop()

    todos_itens = []
    with st.spinner("Processando pedidos, filtrando lixo e construindo planilhas..."):
        for f in files:
            try:
                todos_itens.extend(processar_arquivo(f))
            except Exception as e:
                st.error(f"Erro ao processar {f.name}: {e}")

    if not todos_itens:
        st.error("Nenhum item válido encontrado nos arquivos.")
        st.stop()

    df = pd.DataFrame(todos_itens)
    st.success("Tudo pronto! Arquivos formatados no modelo exato do ERP.")

    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="result-card"><b>Arquivos Lidos</b><br>{len(files)}</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="result-card"><b>Itens Convertidos</b><br>{len(df[df["grupo"] != "NAO_IDENTIFICADO"])}</div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="result-card"><b>Não Cadastrados</b><br>{len(df[df["grupo"] == "NAO_IDENTIFICADO"])}</div>', unsafe_allow_html=True)

    arquivos_gerados = gerar_arquivos_excel(df)
    
    cols = st.columns(2)
    for index, (nome_arquivo, dados_bytes) in enumerate(arquivos_gerados.items()):
        with cols[index % 2]:
            st.download_button(
                label=f"Baixar {nome_arquivo}",
                data=dados_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
