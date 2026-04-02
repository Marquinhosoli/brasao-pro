import streamlit as st
import pandas as pd
import unicodedata
import math
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="THOTH PRO Multi-Cliente", page_icon="📦", layout="wide")

# --- FUNÇÕES ÚTEIS ---
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto

@st.cache_data
def processar_base(df_base):
    dicionario = {}
    for index, row in df_base.iterrows():
        # Busca o nome do produto nas colunas possíveis
        produto_base = normalizar_texto(row.get('produto_base', row.get('produto', row.get('PRODUTO', ''))))
        if not produto_base:
            continue

        item = {
            'produto_base': produto_base,
            'categoria': normalizar_texto(row.get('categoria', row.get('CATEGORIA', 'INDEFINIDO'))),
            'tipo': normalizar_texto(row.get('modo_conversao', row.get('tipo', 'CAIXA'))),
            'peso_caixa': float(row.get('peso_caixa', 0)) if pd.notna(row.get('peso_caixa')) else 0,
            'bandejas_por_caixa': float(row.get('bandejas_por_caixa', 0)) if pd.notna(row.get('bandejas_por_caixa')) else 0,
            'unidades_caixa': float(row.get('itens_por_caixa', row.get('unidades_caixa', 0))) if pd.notna(row.get('itens_por_caixa', row.get('unidades_caixa'))) else 0
        }

        dicionario[produto_base] = item

        # Processar sinônimos (variações de escrita do Brasão/Krill)
        sinonimos_str = str(row.get('sinonimos', row.get('SINONIMOS', '')))
        if sinonimos_str and sinonimos_str.lower() != 'nan':
            sinonimos = [normalizar_texto(s.strip()) for s in sinonimos_str.split(';')]
            for sin in sinonimos:
                if sin:
                    dicionario[sin] = item
                    
    return dicionario

def encontrar_colunas(df):
    idx_prod, idx_qtd, idx_loja = None, None, None
    colunas = [normalizar_texto(c) for c in df.columns]

    for i, col in enumerate(colunas):
        if any(x in col for x in ["PRODUTO", "DESCRICAO", "MERCADORIA", "ITEM"]):
            idx_prod = df.columns[i]
        elif any(x in col for x in ["QTD", "QUANT", "PEDIDO", "VOL", "CX"]):
            idx_qtd = df.columns[i]
        elif any(x in col for x in ["LOJA", "CLIENTE", "DESTINO"]):
            idx_loja = df.columns[i]

    # Fallback caso a planilha venha sem cabeçalho claro
    if not idx_prod: idx_prod = df.columns[0]
    if not idx_qtd:
        if len(df.columns) > 1:
            idx_qtd = df.columns[1]
        else:
            idx_qtd = df.columns[0]

    return idx_prod, idx_qtd, idx_loja

# --- INTERFACE LATERAL (SIDEBAR) ---
st.sidebar.title("📦 THOTH PRO")
st.sidebar.markdown("**Motor Dinâmico de Conversão**")
st.sidebar.markdown("---")

st.sidebar.markdown("### 🟡 1. CARREGAR BASE PRODUTOS")
base_file = st.sidebar.file_uploader("Selecione sua base (Excel ou CSV)", type=["xlsx", "xls", "csv"], key="base")

st.sidebar.markdown("---")

st.sidebar.markdown("### 🔵 2. SELECIONAR PEDIDOS")
uploaded_files = st.sidebar.file_uploader("Selecione os arquivos de Pedido das Lojas", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pedidos")

# --- TRAVAS DE SEGURANÇA ---
if base_file is None:
    st.info("👈 Por favor, carregue a sua BASE DE PRODUTOS no menu lateral para iniciar o sistema.")
    st.stop()

if not uploaded_files:
    st.info("👈 Agora selecione os arquivos de pedido dos clientes (Brasão, Kross, CD, etc).")
    st.stop()

# --- PROCESSAMENTO PRINCIPAL ---
st.title("📊 Painel de Processamento")

try:
    # 1. Lê a Base
    if base_file.name.endswith('.csv'):
        df_base = pd.read_csv(base_file)
    else:
        df_base = pd.read_excel(base_file)

    dicionario_produtos = processar_base(df_base)
    st.sidebar.success(f"✅ Base carregada! {len(df_base)} itens na memória.")

    # Listas para separar as abas
    processados_frutas = []
    processados_legumes = []
    processados_sem_base = []
    consolidado = []

    cont_total = 0
    cont_sucesso = 0
    cont_falha = 0

    # 2. Varre todos os arquivos de pedido selecionados
    for file in uploaded_files:
        nome_arquivo_norm = normalizar_texto(file.name)
        loja_arquivo = "INDEFINIDA"
        if "KRILL" in nome_arquivo_norm: loja_arquivo = "KRILL"
        elif "BRASAO" in nome_arquivo_norm: loja_arquivo = "BRASAO"
        elif "KROSS" in nome_arquivo_norm: loja_arquivo = "KROSS"
        elif "CD" in nome_arquivo_norm: loja_arquivo = "CENTRO DISTRIBUICAO"

        if file.name.endswith('.csv'):
            df_temp = pd.read_csv(file)
        else:
            df_temp = pd.read_excel(file)

        if df_temp.empty:
            continue

        col_prod, col_qtd, col_loja = encontrar_colunas(df_temp)

        for index, row in df_temp.iterrows():
            produto_raw = str(row[col_prod])
            qtd_raw = row[col_qtd]

            # Ignora linhas em branco
            if pd.isna(produto_raw) or str(produto_raw).strip() == "" or pd.isna(qtd_raw):
                continue

            nome_normalizado = normalizar_texto(produto_raw)

            # Ignora se for sujeira de cabeçalho
            if nome_normalizado in ["PRODUTO", "DESCRICAO", "MERCADORIA", "ITEM", "TOTAL"]:
                continue

            try:
                qtd_num = float(str(qtd_raw).replace(',', '.'))
            except ValueError:
                continue

            loja_row = str(row[col_loja]) if col_loja and pd.notna(row[col_loja]) else loja_arquivo
            cont_total += 1
            
            base_ref = dicionario_produtos.get(nome_normalizado)

            registro = {
                "LOJA": loja_row,
                "PRODUTO_ORIGINAL": produto_raw,
                "QTD_PEDIDA": qtd_num,
                "PRODUTO_BASE": base_ref['produto_base'] if base_ref else "NÃO ENCONTRADO",
                "CAIXAS_CONVERTIDAS": 0,
                "TIPO": base_ref['tipo'] if base_ref else "-",
                "CATEGORIA": base_ref['categoria'] if base_ref else "-"
            }

            if base_ref:
                caixas = 0
                tipo = base_ref['tipo']
                # REGRA CRÍTICA DE DIVISÃO (Sempre arredondando para cima com math.ceil)
                if tipo in ["PESO", "KG", "PESO_CAIXA"] and base_ref['peso_caixa'] > 0:
                    caixas = math.ceil(qtd_num / base_ref['peso_caixa'])
                elif tipo in ["BDJ", "BANDEJA"] and base_ref['bandejas_por_caixa'] > 0:
                    caixas = math.ceil(qtd_num / base_ref['bandejas_por_caixa'])
                elif tipo in ["UN", "UNIDADE"] and base_ref['unidades_caixa'] > 0:
                    caixas = math.ceil(qtd_num / base_ref['unidades_caixa'])
                else:
                    caixas = math.ceil(qtd_num) # Cai aqui se for CAIXA fechada

                registro["CAIXAS_CONVERTIDAS"] = caixas
                cont_sucesso += 1

                consolidado.append(registro)
                if base_ref['categoria'] == "FRUTAS":
                    processados_frutas.append(registro)
                elif base_ref['categoria'] == "LEGUMES":
                    processados_legumes.append(registro)
                else:
                    processados_frutas.append(registro) 

            else:
                registro["MOTIVO"] = "Não mapeado ou falta fator na Base"
                processados_sem_base.append(registro)
                consolidado.append(registro)
                cont_falha += 1

    # --- EXIBIR DASHBOARD ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Itens Lidos", cont_total)
    col2.metric("Itens Convertidos", cont_sucesso)
    col3.metric("Sem Base / Pendentes", cont_falha)

    # --- GERAR EXCEL DE SAÍDA ---
    if consolidado:
        df_frutas = pd.DataFrame(processados_frutas)
        df_legumes = pd.DataFrame(processados_legumes)
        df_sem_base = pd.DataFrame(processados_sem_base)
        df_consolidado = pd.DataFrame(consolidado)

        # Aba de Códigos e Preços unificada e em ordem alfabética
        df_codigos = pd.concat([df_frutas, df_legumes], ignore_index=True)
        if not df_codigos.empty:
            df_codigos = df_codigos.sort_values(by="PRODUTO_BASE")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            if not df_frutas.empty: df_frutas.to_excel(writer, sheet_name="FRUTAS", index=False)
            if not df_legumes.empty: df_legumes.to_excel(writer, sheet_name="LEGUMES", index=False)
            if not df_sem_base.empty: df_sem_base.to_excel(writer, sheet_name="SEM_BASE", index=False)
            if not df_codigos.empty: df_codigos.to_excel(writer, sheet_name="CODIGOS_PRECOS", index=False)
            if not df_consolidado.empty: df_consolidado.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
            df_base.to_excel(writer, sheet_name="BASE_MESTRE", index=False)

        processed_data = output.getvalue()

        st.success("✅ Processamento concluído em segundos!")
        st.download_button(
            label="⬇️ BAIXAR EXCEL FINAL PADRONIZADO",
            data=processed_data,
            file_name="THOTH_PRO_RESULTADO_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        if cont_falha > 0:
            st.warning(f"Atenção: {cont_falha} itens não cruzaram com a base. Edite sua planilha de produtos e rode novamente.")
            st.dataframe(df_sem_base)

except Exception as e:
    st.error(f"Ocorreu um erro no processamento: {e}")
