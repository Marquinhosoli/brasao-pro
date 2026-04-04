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
st.write("Processamento e Conversão para Importação Thoth (Arquivos Separados)")

files = st.file_uploader(
    "Envie os PDFs de pedidos",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True
)

# =========================
# BASE DE CONVERSÃO (Atualizada com seus itens)
# =========================
BASE_PRODUTOS = {
    "ABACAXI PEROLA": {"modo": "un", "por_caixa": 1, "grupo": "FRUTAS"},
    "ABACATE": {"modo": "kg", "por_caixa": 20, "grupo": "FRUTAS"},
    "ABOBORA PESCOCO": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "ALECRIM": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "ALHO PORO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "AMEIXA NACIONAL": {"modo": "bdj", "por_caixa": 30, "grupo": "FRUTAS"},
    "BATATA ASTERIX": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA BOLINHA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA DOCE BRANCA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA DOCE ROXA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "BATATA SALSA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "BERINJELA": {"modo": "kg", "por_caixa": 10, "grupo": "LEGUMES"},
    "BETERRABA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "CAQUI RAMA FORTE": {"modo": "kg", "por_caixa": 6, "grupo": "FRUTAS"},
    "CARAMBOLA": {"modo": "bdj", "por_caixa": 4, "grupo": "FRUTAS"},
    "CEBOLA ARGENTINA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "CEBOLA CONSERVA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "CENOURA": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "CHUCHU": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "COCO SECO": {"modo": "kg", "por_caixa": 20, "grupo": "FRUTAS"},
    "COENTRO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "FIGO ROXO": {"modo": "bdj", "por_caixa": 1, "grupo": "FRUTAS"},
    "FRAMBOESA": {"modo": "bdj", "por_caixa": 15, "grupo": "FRUTAS"},
    "GOIABA NACIONAL": {"modo": "kg", "por_caixa": 20, "grupo": "FRUTAS"},
    "HORTELA": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "JATOBA": {"modo": "kg", "por_caixa": 1, "grupo": "FRUTAS"},
    "KINKAN": {"modo": "bdj", "por_caixa": 10, "grupo": "FRUTAS"},
    "KIWI IMPORTADO": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "KIWI NACIONAL": {"modo": "bdj", "por_caixa": 15, "grupo": "FRUTAS"},
    "LOURO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "MACA FUJI": {"modo": "kg", "por_caixa": 18, "grupo": "FRUTAS"},
    "MAMAO FORMOSA": {"modo": "kg", "por_caixa": 15, "grupo": "FRUTAS"},
    "MAMAOZINHO PAPAIA": {"modo": "un", "por_caixa": 18, "grupo": "FRUTAS"},
    "MANGA PALMER": {"modo": "kg", "por_caixa": 12, "grupo": "FRUTAS"},
    "MANJERICAO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "MANJERONA": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "MAXIXE": {"modo": "bdj", "por_caixa": 12, "grupo": "LEGUMES"},
    "MELAO CANTALOUPE": {"modo": "un", "por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO CHARANTEAIS": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO DINO": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO ESPANHOL": {"modo": "kg", "por_caixa": 13, "grupo": "FRUTAS"},
    "MELAO GALIA": {"modo": "un", "por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO ORANGE": {"modo": "un", "por_caixa": 6, "grupo": "FRUTAS"},
    "MELAO REI DOCE": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "MELAO SAPO": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "MELANCIA": {"modo": "kg", "por_caixa": 1, "grupo": "FRUTAS"},
    "MILHO VERDE": {"modo": "bdj", "por_caixa": 10, "grupo": "LEGUMES"},
    "MIRTILO BLUEBERRY": {"modo": "bdj", "por_caixa": 12, "grupo": "FRUTAS"},
    "NABO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "PEPINO JAPONES": {"modo": "kg", "por_caixa": 20, "grupo": "LEGUMES"},
    "PERA WILLIANS": {"modo": "kg", "por_caixa": 19, "grupo": "FRUTAS"},
    "PESSEGO IMP": {"modo": "kg", "por_caixa": 10, "grupo": "FRUTAS"},
    "PHYSALIS": {"modo": "bdj", "por_caixa": 8, "grupo": "FRUTAS"},
    "PIMENTA BIQUINHO": {"modo": "kg", "por_caixa": 1, "grupo": "LEGUMES"},
    "PIMENTA CAMBUCI": {"modo": "kg", "por_caixa": 1, "grupo": "LEGUMES"},
    "PIMENTA JALAPENO": {"modo": "kg", "por_caixa": 1, "grupo": "LEGUMES"},
    "SALSAO AIPO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "SALVIA": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "TOMATE GRAPE": {"modo": "bdj", "por_caixa": 10, "grupo": "LEGUMES"},
    "TOMILHO": {"modo": "un", "por_caixa": 1, "grupo": "LEGUMES"},
    "UVA THOMPSON": {"modo": "bdj", "por_caixa": 10, "grupo": "FRUTAS"}
}

# =========================
# FUNÇÕES DE EXTRAÇÃO E LIMPEZA
# =========================
def normalizar_nome(texto: str) -> str:
    texto = (texto or "").upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace("Ç", "C").replace("Ã", "A").replace("Á", "A").replace("À", "A")
    texto = texto.replace("É", "E").replace("Ê", "E").replace("Í", "I")
    texto = texto.replace("Ó", "O").replace("Õ", "O").replace("Ô", "O").replace("Ú", "U")
    return texto

def identificar_loja(nome_arquivo: str):
    """
    Identifica o cliente e mapeia para a Coluna numérica exata que a planilha Thoth exige.
    """
    n = normalizar_nome(nome_arquivo)
    
    # Regras KROSS (1 = Chapeco/Atacado, 2 = Xaxim)
    if "KROSS" in n and "XAXIM" in n: return "KROSS", "2"
    if "KROSS" in n: return "KROSS", "1"
    
    # Regras CD
    if "CD" in n: return "BRASAO CD", "1"
    
    # Regras BRASAO (1 = Fernando, 2 = Jardim, 3 = Xaxim, 4 = Avenida)
    if "FERNANDO" in n: return "BRASAO", "1"
    if "JARDIM" in n: return "BRASAO", "2"
    if "XAXIM" in n: return "BRASAO", "3"
    if "AVENIDA" in n: return "BRASAO", "4"
    
    return "OUTROS", "0"

def extrair_texto_pdf(uploaded_file) -> str:
    texto = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip(): texto.append(t)
    return "\n".join(texto)

def parse_linha_produto(linha: str):
    l = normalizar_nome(linha)
    
    # Ignora cabeçalhos que possam gerar falsos positivos
    if "TOTAL" in l or "PESO" in l or "FRETE" in l:
        return None

    # NOVO PARSER BLINDADO: Aceita 0, 1 ou 2 códigos no início e isola a matemática no final.
    # Ex: [131393] [ABACATE KG BRASAO FRUTA] [200,000] [200] [3,3000] [660,00]
    m_flex = re.search(r"^(?:[\d\s-]*\s+)?([A-Z].*?)\s+(\d+[.,]\d+)\s+(\d+)\s+\d+[.,]\d+\s+\d+[.,]\d+\s*$", l)
    
    if m_flex:
        descricao = m_flex.group(1)
        qtd = float(m_flex.group(2).replace(",", ".")) # Usamos a Quantidade exata do peso/unidade
        
        # Procura a unidade dentro do nome (KG, BDJ, UN, etc)
        m_un = re.search(r"\b(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS|CX|CXS|CAIXA|CAIXAS|VOL|VOLUME|VOLUMES|MACO|MACOS)\b", descricao)
        un_raw = m_un.group(1) if m_un else "CX"

        # Limpa marcas e descrições inúteis do nome final
        produto = re.sub(r"\b(BRASAO FRU\w*|BRASAO FRUTA|DE MARCHI|SHELF \d+|FRUTAMINA)\b", "", descricao)
        produto = re.sub(r"\b(KG|KGS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|CX|CXS|MACO|MACOS)\b", "", produto).strip()
        produto = normalizar_nome(produto)
        
        if un_raw in ["KG", "KGS", "QUILO", "QUILOS"]: unidade = "kg"
        elif un_raw in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES", "MACO", "MACOS"]: unidade = "un"
        elif un_raw in ["CX", "CXS", "CAIXA", "CAIXAS", "VOL", "VOLUME", "VOLUMES"]: unidade = "cx"
        else: unidade = "bdj"
            
        return produto, qtd, unidade

    return None

def localizar_base(produto: str):
    p = normalizar_nome(produto)
    if p in BASE_PRODUTOS:
        return p, BASE_PRODUTOS[p]
    for chave in BASE_PRODUTOS:
        if chave in p:
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
    cliente, loja_num = identificar_loja(nome)

    if nome.lower().endswith(".pdf"):
        texto = extrair_texto_pdf(uploaded_file)
        linhas = []
        for l in texto.splitlines():
            limpa = l.strip()
            if not limpa: continue
            
            # BLOQUEIO ANTI-LIXO: Ignora Pendências do CD
            if "PENDENCIAS DE MERCADORIAS" in normalizar_nome(limpa):
                break # Para de ler o arquivo daqui pra baixo
                
            linhas.append(limpa)
    else:
        df_excel = pd.read_excel(uploaded_file)
        linhas = [" ".join([str(x) for x in row if pd.notna(x)]) for _, row in df_excel.iterrows()]

    itens = []
    
    for linha in linhas:
        item = parse_linha_produto(linha)
        if item:
            produto, qtd, unidade = item
            conv = converter_para_caixa(produto, qtd, unidade)
            conv["cliente"] = cliente
            conv["loja_cod"] = loja_num
            conv["arquivo"] = nome
            itens.append(conv)

    return itens

# =========================
# GERAÇÃO DE ARQUIVOS SEPARADOS (Padrão Thoth CSV)
# =========================
def gerar_planilha_thoth(df_itens, cliente, grupo, colunas_numericas):
    """
    Cria a Tabela Dinâmica com colunas '1', '2', '3', '4' e 'TOTAL' para bater exato com o ERP.
    """
    df_filtro = df_itens[(df_itens["cliente"] == cliente) & (df_itens["grupo"] == grupo)]
    
    if df_filtro.empty:
        return pd.DataFrame() 

    # Agrupa por Produto e Loja (Pivot Table / Matriz)
    pivot = pd.pivot_table(
        df_filtro,
        values='qtd_caixa',
        index='produto_final',
        columns='loja_cod',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    pivot.rename(columns={'produto_final': 'PRODUTO'}, inplace=True)

    # Garante que as colunas (1, 2, 3...) existam
    for col in colunas_numericas:
        if col not in pivot.columns:
            pivot[col] = 0

    # Adiciona a coluna TOTAL
    pivot["TOTAL"] = pivot[colunas_numericas].sum(axis=1)

    # Organiza a ordem exata do CSV
    colunas_finais = ["PRODUTO"] + colunas_numericas + ["TOTAL"]
    pivot = pivot[colunas_finais].sort_values(by="PRODUTO").reset_index(drop=True)

    # Limpa quem não teve pedido em nenhuma loja
    pivot = pivot[pivot["TOTAL"] > 0]

    return pivot

def gerar_arquivos_excel(df):
    """
    Gera um dicionário contendo os arquivos Excel separados em memória.
    """
    arquivos = {}

    configs = [
        ("BRASAO", "FRUTAS", "BRASAO_FRUTAS_IMPORTACAO.xlsx", ["1", "2", "3", "4"]),
        ("BRASAO", "LEGUMES", "BRASAO_LEGUMES_IMPORTACAO.xlsx", ["1", "2", "3", "4"]),
        ("KROSS", "FRUTAS", "KROSS_FRUTAS_IMPORTACAO.xlsx", ["1", "2"]),
        ("KROSS", "LEGUMES", "KROSS_LEGUMES_IMPORTACAO.xlsx", ["1", "2"]),
        ("BRASAO CD", "FRUTAS", "BRASAO_CD_FRUTAS_IMPORTACAO.xlsx", ["1"]),
        ("BRASAO CD", "LEGUMES", "BRASAO_CD_LEGUMES_IMPORTACAO.xlsx", ["1"]),
    ]

    for cliente, grupo, nome_arquivo, colunas in configs:
        df_gerado = gerar_planilha_thoth(df, cliente, grupo, colunas)
        
        if not df_gerado.empty:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                # O Thoth geralmente lê a aba "Plan1" padrão do Excel
                df_gerado.to_excel(writer, index=False, sheet_name="Plan1")
            arquivos[nome_arquivo] = output.getvalue()

    # Planilha de erros para auditoria
    df_sem_base = df[df["grupo"] == "NAO_IDENTIFICADO"][
        ["produto_final", "qtd_original", "unidade_original", "loja_cod", "arquivo"]
    ].rename(columns={
        "produto_final": "NOME EXTRAÍDO (Adicione à Base)",
        "qtd_original": "Qtd Original",
        "unidade_original": "Unidade",
        "loja_cod": "Filial",
        "arquivo": "PDF de Origem"
    })
    
    if not df_sem_base.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_sem_base.to_excel(writer, index=False, sheet_name="ITENS REJEITADOS")
        arquivos["ITENS_NAO_RECONHECIDOS.xlsx"] = output.getvalue()

    return arquivos

# =========================
# BOTÃO PROCESSAR E BAIXAR
# =========================
if st.button("🔥 PROCESSAR PEDIDOS (Gerar Arquivos Individuais)", use_container_width=False):
    if not files:
        st.warning("Envie pelo menos um arquivo de pedido (PDF/Excel).")
        st.stop()

    todos_itens = []

    with st.spinner("Processando pedidos, filtrando lixo e construindo planilhas..."):
        for f in files:
            try:
                itens = processar_arquivo(f)
                todos_itens.extend(itens)
            except Exception as e:
                st.error(f"Erro ao processar {f.name}: {e}")

    if not todos_itens:
        st.error("Nenhum item foi reconhecido. Verifique se o formato bate com o padrão Flex.")
        st.stop()

    df = pd.DataFrame(todos_itens)

    st.success("Tudo pronto! Matrizes geradas rigorosamente no padrão CSV do Thoth.")

    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="result-card"><b>Arquivos Lidos</b><br>{len(files)}</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="result-card"><b>Itens Convertidos</b><br>{len(df[df["grupo"] != "NAO_IDENTIFICADO"])}</div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="result-card"><b>Não Cadastrados</b><br>{len(df[df["grupo"] == "NAO_IDENTIFICADO"])}</div>', unsafe_allow_html=True)

    st.subheader("📥 Arquivos Prontos para o Thoth")
    st.write("Cada arquivo baixado terá exatamente a estrutura: PRODUTO | 1 | 2 | 3 | 4 | TOTAL.")
    
    arquivos_gerados = gerar_arquivos_excel(df)
    
    cols = st.columns(2)
    for index, (nome_arquivo, dados_bytes) in enumerate(arquivos_gerados.items()):
        col = cols[index % 2]
        with col:
            st.download_button(
                label=f"Baixar {nome_arquivo}",
                data=dados_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    st.markdown('<div class="small-muted">O leitor foi corrigido para ignorar "Pendências", não capturar números como nome, e exportar as planilhas exatas com colunas 1, 2, 3, 4 e TOTAL.</div>', unsafe_allow_html=True)
