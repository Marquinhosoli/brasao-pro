import io
import math
import unicodedata
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# =========================================================
# THOTH PRO Multi-Cliente (Produção Final)
# Streamlit app
# =========================================================

st.set_page_config(
    page_title="THOTH PRO Multi-Cliente (Produção Final)",
    page_icon="📦",
    layout="wide",
)

ERP_COLUMNS = [
    "PRODUTO",
    "BRASAO FERNANDO",
    "BRASAO JARDIM",
    "BRASAO XAXIM",
    "BRASAO AVENIDA",
    "KROSS ATACADO",
    "KROSS XAXIM",
]

DEMO_BASE = [
    {
        "produto": "ABACAXI",
        "categoria": "FRUTAS",
        "tipo": "UN",
        "unidades_caixa": 10,
        "sinonimos": "ABACAXI HAWAI;ABACAXI PEROLA",
    },
    {
        "produto": "BLUEBERRY 125G",
        "categoria": "FRUTAS",
        "tipo": "BDJ",
        "bandejas_por_caixa": 12,
        "sinonimos": "MIRTILO BLUEBERRY;MIRTILO;BLUEBERRY",
    },
    {
        "produto": "CEBOLA ALBINA",
        "categoria": "LEGUMES",
        "tipo": "KG",
        "peso_caixa": 20,
        "sinonimos": "CEBOLA ARGENTINA",
    },
    {
        "produto": "LARANJA PERA",
        "categoria": "FRUTAS",
        "tipo": "KG",
        "peso_caixa": 20,
        "sinonimos": "LARANJA",
    },
    {
        "produto": "LIMAO TAHITI",
        "categoria": "FRUTAS",
        "tipo": "KG",
        "peso_caixa": 20,
        "sinonimos": "LIMAO;LIMAO TAHITI",
    },
    {
        "produto": "MAMAO PAPAYA",
        "categoria": "FRUTAS",
        "tipo": "UN",
        "unidades_caixa": 18,
        "sinonimos": "MAMAO PAPAYA;PAPAYA",
    },
    {
        "produto": "MELAO",
        "categoria": "FRUTAS",
        "tipo": "UN",
        "unidades_caixa": 10,
        "sinonimos": "MELAO",
    },
    {
        "produto": "MORANGO",
        "categoria": "FRUTAS",
        "tipo": "BDJ",
        "bandejas_por_caixa": 4,
        "sinonimos": "MORANGO 250G",
    },
    {
        "produto": "PHYSALIS IMPORTADO 100G",
        "categoria": "FRUTAS",
        "tipo": "UN",
        "unidades_caixa": 8,
        "sinonimos": "PHYSALIS IMPORTADO;PHYSALIS",
    },
    {
        "produto": "TOMATE GRAP",
        "categoria": "LEGUMES",
        "tipo": "BDJ",
        "bandejas_por_caixa": 24,
        "sinonimos": "TOMATE GRAPE DEMARCHI;TOMATE GRAPE;TOMATE GRAPE 180G",
    },
    {
        "produto": "UVA THOMPSON 500G",
        "categoria": "FRUTAS",
        "tipo": "BDJ",
        "bandejas_por_caixa": 10,
        "sinonimos": "UVA THOMPSON S/SEMENTE;UVA THOMPSON S SEMENTE;UVA THOMPSON",
    },
    {
        "produto": "CARAMBOLA 400G",
        "categoria": "FRUTAS",
        "tipo": "BDJ",
        "bandejas_por_caixa": 4,
        "sinonimos": "CARAMBOLA",
    },
]


def normalizar_texto(texto) -> str:
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    return " ".join(texto.split())


def normalizar_produto(texto: str) -> str:
    s = normalizar_texto(texto)

    remocoes = [
        " DEMARCHI ",
        " FRUTA ",
        " LEGUME ",
        " UND ",
        " UNID ",
        " UNIDADE ",
        " BDJ ",
        " BANDEJA ",
    ]

    s = f" {s} "
    for token in remocoes:
        s = s.replace(token, " ")

    import re

    s = re.sub(r"\bSHELF\s*\d+\b", " ", s)
    s = re.sub(r"\bC/\d+G\b", " ", s)
    s = re.sub(r"\b\d+G\b", " ", s)
    s = re.sub(r"\bKG\b", " ", s)
    s = " ".join(s.split()).strip()

    mapa_exato = {
        "TOMATE GRAPE DEMARCHI": "TOMATE GRAP",
        "TOMATE GRAPE": "TOMATE GRAP",
        "UVA THOMPSON S/SEMENTE": "UVA THOMPSON 500G",
        "UVA THOMPSON S SEMENTE": "UVA THOMPSON 500G",
        "MIRTILO BLUEBERRY": "BLUEBERRY 125G",
        "BLUEBERRY": "BLUEBERRY 125G",
        "CEBOLA ARGENTINA": "CEBOLA ALBINA",
        "PHYSALIS IMPORTADO": "PHYSALIS IMPORTADO 100G",
        "PHYSALIS": "PHYSALIS IMPORTADO 100G",
        "CARAMBOLA": "CARAMBOLA 400G",
    }

    if s in mapa_exato:
        return mapa_exato[s]

    if "TOMATE" in s and "GRAPE" in s:
        return "TOMATE GRAP"
    if "UVA" in s and "THOMPSON" in s:
        return "UVA THOMPSON 500G"
    if "MIRTILO" in s or "BLUEBERRY" in s:
        return "BLUEBERRY 125G"
    if "CEBOLA" in s and "ARGENTINA" in s:
        return "CEBOLA ALBINA"
    if "PHYSALIS" in s:
        return "PHYSALIS IMPORTADO 100G"
    if "CARAMBOLA" in s:
        return "CARAMBOLA 400G"

    return s


def parse_numero(valor) -> float:
    if pd.isna(valor):
        return float("nan")
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    if not s:
        return float("nan")
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return float("nan")


def padronizar_loja(nome_loja: str, nome_arquivo: str = "") -> str:
    nome = normalizar_texto(f"{nome_loja or ''} {nome_arquivo or ''}")

    if "FERNANDO" in nome:
        return "BRASAO FERNANDO"
    if "JARDIM" in nome:
        return "BRASAO JARDIM"
    if "BRASAO" in nome and "XAXIM" in nome:
        return "BRASAO XAXIM"
    if "AVENIDA" in nome:
        return "BRASAO AVENIDA"
    if "KROSS" in nome and "ATACADO" in nome:
        return "KROSS ATACADO"
    if "KROSS" in nome and "XAXIM" in nome:
        return "KROSS XAXIM"

    if "BRASAO" in nome and "CE" in nome:
        return "BRASAO FERNANDO"
    if "BRASAO" in nome and "JA" in nome:
        return "BRASAO JARDIM"
    if "BRASAO" in nome and "XX" in nome:
        return "BRASAO XAXIM"
    if "BRASAO" in nome and "AV" in nome:
        return "BRASAO AVENIDA"
    if "KROSS" in nome and "XX" in nome:
        return "KROSS XAXIM"

    return ""


def detectar_tipo_no_texto(texto: str) -> str:
    s = normalizar_texto(texto)
    if "KG" in s:
        return "KG"
    if "BDJ" in s or "BANDEJA" in s:
        return "BDJ"
    if "UND" in s or "UNID" in s or "UNIDADE" in s or " UN " in f" {s} ":
        return "UN"
    return ""


def encontrar_colunas(df: pd.DataFrame) -> Tuple[str, str, str, str]:
    headers = list(df.columns)
    norm = [normalizar_texto(c) for c in headers]

    col_prod = headers[0] if headers else None
    col_qtd = headers[1] if len(headers) > 1 else (headers[0] if headers else None)
    col_loja = None
    col_tipo = None

    for i, h in enumerate(norm):
        if any(k in h for k in ["PRODUTO", "DESCRICAO", "DESCRIÇÃO", "ITEM", "MERCADORIA"]):
            col_prod = headers[i]
        if any(k in h for k in ["QTD", "QUANT", "QUANTIDADE", "PEDIDO", "QTDE", "VOLUME"]):
            col_qtd = headers[i]
        if any(k in h for k in ["LOJA", "CLIENTE", "DESTINO", "FILIAL"]):
            col_loJA = headers[i]
            col_loja = col_loJA
        if any(k in h for k in ["TIPO", "UNIDADE", "UND", "MEDIDA"]):
            col_tipo = headers[i]

    return col_prod, col_qtd, col_loja, col_tipo


@st.cache_data(show_spinner=False)
def ler_arquivo_para_df(file_bytes: bytes, nome: str) -> pd.DataFrame:
    buffer = io.BytesIO(file_bytes)
    if nome.lower().endswith(".csv"):
        return pd.read_csv(buffer)
    return pd.read_excel(buffer)


def carregar_base_demo() -> pd.DataFrame:
    return pd.DataFrame(DEMO_BASE)


@st.cache_data(show_spinner=False)
def processar_base(df_base: pd.DataFrame) -> Dict[str, dict]:
    dicionario = {}

    for _, row in df_base.iterrows():
        produto = row.get("produto", row.get("PRODUTO", row.get("produto_base", row.get("PRODUTO_BASE", ""))))
        categoria = normalizar_texto(row.get("categoria", row.get("CATEGORIA", "")))
        tipo = normalizar_texto(row.get("tipo", row.get("TIPO", row.get("modo_conversao", row.get("MODO_CONVERSAO", "")))))
        peso_caixa = parse_numero(row.get("peso_caixa", row.get("PESO_CAIXA", 0)))
        bandejas_por_caixa = parse_numero(row.get("bandejas_por_caixa", row.get("BANDEJAS_POR_CAIXA", 0)))
        unidades_caixa = parse_numero(row.get("unidades_caixa", row.get("UNIDADES_CAIXA", row.get("itens_por_caixa", row.get("ITENS_POR_CAIXA", 0)))))
        sinonimos_str = str(row.get("sinonimos", row.get("SINONIMOS", "")) or "")

        produto_norm = normalizar_produto(produto)
        if not produto_norm:
            continue

        item = {
            "produto": produto_norm,
            "categoria": categoria,
            "tipo": tipo,
            "peso_caixa": peso_caixa if not math.isnan(peso_caixa) else 0,
            "bandejas_por_caixa": bandejas_por_caixa if not math.isnan(bandejas_por_caixa) else 0,
            "unidades_caixa": unidades_caixa if not math.isnan(unidades_caixa) else 0,
        }

        dicionario[produto_norm] = item

        if sinonimos_str and sinonimos_str.lower() != "nan":
            sinonimos = [normalizar_produto(s.strip()) for s in sinonimos_str.split(";") if s.strip()]
            for sin in sinonimos:
                dicionario[sin] = item

    return dicionario


def base_incompleta(item_base: dict) -> bool:
    tipo = item_base.get("tipo", "")
    if tipo == "KG":
        return not (item_base.get("peso_caixa", 0) > 0)
    if tipo == "BDJ":
        return not (item_base.get("bandejas_por_caixa", 0) > 0)
    if tipo == "UN":
        return not (item_base.get("unidades_caixa", 0) > 0)
    return True


def converter_para_caixas(qtd: float, item_base: dict) -> int:
    tipo = item_base["tipo"]

    if tipo == "KG":
        return math.ceil(qtd / item_base["peso_caixa"])
    if tipo == "BDJ":
        return math.ceil(qtd / item_base["bandejas_por_caixa"])
    if tipo == "UN":
        return math.ceil(qtd / item_base["unidades_caixa"])

    raise ValueError("Tipo de conversão inválido")


def gerar_saida_erp(lista: List[dict], categoria: str) -> pd.DataFrame:
    dados = [x for x in lista if x["status"] == "OK" and x["categoria"] == categoria]
    if not dados:
        return pd.DataFrame(columns=ERP_COLUMNS)

    df = pd.DataFrame(dados)
    df_group = (
        df.groupby(["produto_base", "loja"], as_index=False)["caixas"]
        .sum()
    )
    df_pivot = (
        df_group.pivot(index="produto_base", columns="loja", values="caixas")
        .fillna(0)
        .reset_index()
    )

    df_pivot.rename(columns={"produto_base": "PRODUTO"}, inplace=True)

    for col in ERP_COLUMNS:
        if col not in df_pivot.columns:
            df_pivot[col] = 0 if col != "PRODUTO" else ""

    df_pivot = df_pivot[ERP_COLUMNS]
    df_pivot = df_pivot.sort_values("PRODUTO").reset_index(drop=True)

    for col in ERP_COLUMNS[1:]:
        df_pivot[col] = df_pivot[col].astype(int)

    return df_pivot


def gerar_excel_final(df_frutas: pd.DataFrame, df_legumes: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        if df_frutas.empty:
            pd.DataFrame(columns=ERP_COLUMNS).to_excel(writer, sheet_name="FRUTAS", index=False)
        else:
            df_frutas.to_excel(writer, sheet_name="FRUTAS", index=False)

        if df_legumes.empty:
            pd.DataFrame(columns=ERP_COLUMNS).to_excel(writer, sheet_name="LEGUMES", index=False)
        else:
            df_legumes.to_excel(writer, sheet_name="LEGUMES", index=False)

    output.seek(0)
    return output.getvalue()


# =========================================================
# UI
# =========================================================

st.title("📦 THOTH PRO Multi-Cliente (Produção Final)")
st.caption("Motor profissional de processamento hortifruti para gerar planilhas no padrão exato do ERP Thoth.")

with st.sidebar:
    st.header("Configuração")
    usar_base_demo = st.checkbox("Usar base modelo embutida", value=False)
    base_file = st.file_uploader("Base mestre", type=["xlsx", "xls", "csv"])
    pedidos_files = st.file_uploader(
        "Pedidos",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )

if usar_base_demo:
    df_base = carregar_base_demo()
    st.sidebar.success(f"Base modelo carregada: {len(df_base)} item(ns).")
elif base_file is not None:
    df_base = ler_arquivo_para_df(base_file.getvalue(), base_file.name)
    st.sidebar.success(f"Base carregada: {len(df_base)} linha(s).")
else:
    df_base = None

if df_base is None:
    st.info("Carregue uma base mestre ou marque a base modelo embutida.")
    st.stop()

dicionario_base = processar_base(df_base)

if not pedidos_files:
    st.info("Selecione pelo menos um pedido para processar.")
    st.stop()

if st.button("PROCESSAR", type="primary", use_container_width=True):
    all_rows = []
    erros = []
    total_lidos = 0
    itens_convertidos = 0
    sem_base = 0
    base_incompleta_count = 0

    try:
        for pedido in pedidos_files:
            df_pedido = ler_arquivo_para_df(pedido.getvalue(), pedido.name)

            if df_pedido.empty:
                continue

            col_prod, col_qtd, col_loja, col_tipo = encontrar_colunas(df_pedido)

            for _, row in df_pedido.iterrows():
                produto_original = str(row[col_prod]).strip() if col_prod in df_pedido.columns else ""
                qtd = parse_numero(row[col_qtd]) if col_qtd in df_pedido.columns else float("nan")

                if not produto_original or math.isnan(qtd):
                    continue

                produto_normalizado = normalizar_produto(produto_original)
                loja = padronizar_loja(row[col_loja] if col_loja else "", pedido.name)
                tipo_detectado = normalizar_texto(row[col_tipo]) if col_tipo else detectar_tipo_no_texto(produto_original)

                if not produto_normalizado or not loja:
                    continue

                total_lidos += 1
                item_base = dicionario_base.get(produto_normalizado)

                if not item_base:
                    sem_base += 1
                    erros.append(
                        {
                            "STATUS": "SEM BASE",
                            "PRODUTO_ORIGINAL": produto_original,
                            "PRODUTO_NORMALIZADO": produto_normalizado,
                            "LOJA": loja,
                            "QTD": qtd,
                            "TIPO_DETECTADO": tipo_detectado,
                            "MOTIVO": "Produto não encontrado na base mestre.",
                        }
                    )
                    continue

                if base_incompleta(item_base):
                    base_incompleta_count += 1
                    erros.append(
                        {
                            "STATUS": "BASE INCOMPLETA",
                            "PRODUTO_ORIGINAL": produto_original,
                            "PRODUTO_NORMALIZADO": produto_normalizado,
                            "LOJA": loja,
                            "QTD": qtd,
                            "TIPO_DETECTADO": item_base.get("tipo", "") or tipo_detectado,
                            "MOTIVO": "Fator obrigatório ausente para conversão.",
                        }
                    )
                    continue

                caixas = converter_para_caixas(qtd, item_base)
                itens_convertidos += 1

                all_rows.append(
                    {
                        "status": "OK",
                        "produto_original": produto_original,
                        "produto_normalizado": produto_normalizado,
                        "produto_base": item_base["produto"],
                        "categoria": item_base["categoria"],
                        "tipo": item_base["tipo"],
                        "qtd": qtd,
                        "caixas": caixas,
                        "loja": loja,
                    }
                )

        df_frutas = gerar_saida_erp(all_rows, "FRUTAS")
        df_legumes = gerar_saida_erp(all_rows, "LEGUMES")
        excel_bytes = gerar_excel_final(df_frutas, df_legumes)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de itens lidos", total_lidos)
        c2.metric("Itens convertidos", itens_convertidos)
        c3.metric("SEM BASE", sem_base)
        c4.metric("BASE INCOMPLETA", base_incompleta_count)

        st.success("Processamento concluído.")

        aba1, aba2, aba3 = st.tabs(["FRUTAS", "LEGUMES", "ERROS"])

        with aba1:
            st.dataframe(df_frutas, use_container_width=True)

        with aba2:
            st.dataframe(df_legumes, use_container_width=True)

        with aba3:
            if erros:
                st.dataframe(pd.DataFrame(erros), use_container_width=True)
            else:
                st.info("Nenhum item com erro.")

        st.download_button(
            label="BAIXAR EXCEL",
            data=excel_bytes,
            file_name="THOTH_PRO_MULTI_CLIENTE_RESULTADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
