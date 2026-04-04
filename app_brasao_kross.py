import math
import re
import unicodedata
import zipfile
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="THOTH PRO Final com Layout", page_icon="📦", layout="wide")

# =========================================================
# CONFIG
# =========================================================

FORCE_UND_KEYWORDS = ["ABACAXI", "MELANCIA"]

STORE_RULES = [
    {"id": "BRASAO_F", "grupo": "BRASAO", "coluna": "BRASAO F", "nome": "Brasão Fernando",
     "file_signals": ["FERNANDO", "CE"], "text_signals": ["FERNANDO MACHADO", "CENTRO, 226"]},
    {"id": "BRASAO_J", "grupo": "BRASAO", "coluna": "BRASAO J", "nome": "Brasão Jardim",
     "file_signals": ["JARDIM", "JA"], "text_signals": ["SAO PEDRO", "JARDIM AMERICA", "2199"]},
    {"id": "BRASAO_X", "grupo": "BRASAO", "coluna": "BRASAO X", "nome": "Brasão Xaxim",
     "file_signals": ["XAXIM", "XX"], "text_signals": ["LUIZ LUNARDI", "XAXIM", "810"]},
    {"id": "BRASAO_A", "grupo": "BRASAO", "coluna": "BRASAO A", "nome": "Brasão Avenida",
     "file_signals": ["AVENIDA", "AV"], "text_signals": ["RIO DE JANEIRO", "CENTRO, 108"]},
    {"id": "KROSS_AT", "grupo": "KROSS", "coluna": "KROSS AT", "nome": "Kross Atacadista",
     "file_signals": ["KROSS", "ATACADO", "CHAPECO"], "text_signals": ["JOHN KENNEDY", "PASSO DOS FORTES", "550"]},
    {"id": "KROSS_X", "grupo": "KROSS", "coluna": "KROSS XAXIM", "nome": "Kross Xaxim",
     "file_signals": ["KROSS", "XAXIM", "XX"], "text_signals": ["AMELIO PANIZZI", "XAXIM"]},
    {"id": "BRASAO_CD", "grupo": "CD", "coluna": "CD", "nome": "Brasão CD",
     "file_signals": ["CD"], "text_signals": ["RUA GASPAR", "ELDORADO", "153"]},
]

IGNORE_LINE_TOKENS = [
    "NUMERO DO PEDIDO", "PEDIDO DE COMPRA", "TRANSACAO", "USUARIO", "FORNECEDOR",
    "EMPRESA:", "CNPJ:", "INSCR EST", "DT. PEDIDO", "FRETE:", "CODIGO COD FORN",
    "AGENDAR A ENTREGA", "TROCAS:", "PENDENCIAS DE MERCADORIAS", "TOTAL DO FORNECEDOR",
    "CONTATOS DO FORNECEDOR", "COMPRADOR", "TOTAIS", "VALOR TOTAL", "PESO TOTAL", "PG:"
]

# =========================================================
# HELPERS
# =========================================================

def norm(text) -> str:
    text = str(text or "").strip().upper()
    text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", text).strip()


def parse_br_number(text):
    s = str(text).strip()
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def load_base(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    required = ["Produto", "Produto Demarchi", "Tipo", "Medida", "Unidade"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"A base precisa ter estas colunas: {', '.join(required)}. Faltando: {', '.join(missing)}")

    df = df.copy()
    df["Produto"] = df["Produto"].astype(str)
    df["Produto Demarchi"] = df["Produto Demarchi"].fillna("").astype(str)
    df["Tipo"] = df["Tipo"].astype(str)
    df["Medida"] = pd.to_numeric(df["Medida"], errors="coerce")
    df["Unidade"] = df["Unidade"].astype(str)

    df["PRODUTO_BASE"] = df["Produto"].apply(norm)
    df["PRODUTO_DEMARCHI_NORM"] = df["Produto Demarchi"].apply(norm)
    df["CATEGORIA"] = df["Tipo"].apply(lambda x: "FRUTAS" if norm(x) == "FRUTAS" else "LEGUMES")
    df["UNIDADE_NORM"] = df["Unidade"].apply(norm)
    return df


def identify_store(filename: str, pdf_text: str) -> dict:
    fn = norm(filename)
    txt = norm(pdf_text[:4000])
    best = STORE_RULES[0]
    best_score = -1

    for rule in STORE_RULES:
        score = 0
        for sig in rule["file_signals"]:
            if norm(sig) in fn:
                score += 5
        for sig in rule["text_signals"]:
            if norm(sig) in txt:
                score += 3
        if score > best_score:
            best_score = score
            best = rule
    return best


def extract_pdf_text(pdf_file) -> str:
    parts = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def is_ignored_line(line: str) -> bool:
    t = norm(line)
    if len(t) < 5:
        return True
    return any(tok in t for tok in IGNORE_LINE_TOKENS)


def clean_product_piece(product_piece: str) -> str:
    t = norm(product_piece)
    t = re.sub(r"^(?:\d+[,\.\d]*\s+)+", "", t)  # remove códigos iniciais
    t = re.sub(r"\bBRASAO FRUTA\b", " ", t)
    t = re.sub(r"\bDE MARCHI\b", " ", t)
    t = re.sub(r"\bDEMARCHI\b", " ", t)
    t = re.sub(r"\bSHELF\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def infer_unit_from_line(line: str) -> str:
    l = f" {norm(line)} "
    if " BANDEJA " in l or " BDJ " in l:
        return "BANDEJA"
    if " UND " in l or " UNIDADE " in l:
        return "UNIDADE"
    if " KG " in l:
        return "KG"
    if " MACO " in l or " MAÇO " in l:
        return "UNIDADE"
    return ""


def extract_items_from_pdf(text: str) -> pd.DataFrame:
    rows = []

    for raw in text.splitlines():
        line = norm(raw)
        if not line or is_ignored_line(line):
            continue

        unit_match = re.search(r"\b(KG|UND|UNIDADE|BDJ|BANDEJA|MACO)\b", line)
        if not unit_match:
            continue

        unit_token = unit_match.group(1)
        before = line[:unit_match.start()].strip()
        after = line[unit_match.end():].strip()

        # produto = trecho antes da unidade, sem códigos
        produto_pdf = clean_product_piece(before)
        if not produto_pdf:
            continue

        # quantidade = primeiro número depois da unidade
        qty_match = re.search(r"(\d{1,4},\d{3}|\d+)", after)
        if not qty_match:
            continue

        qtd = parse_br_number(qty_match.group(1))
        if qtd is None or qtd <= 0:
            continue

        rows.append({
            "produto_pdf": produto_pdf,
            "qtd_original": qtd,
            "unidade_pdf": "BANDEJA" if unit_token == "BDJ" else ("UNIDADE" if unit_token in {"UND", "MACO"} else unit_token),
            "linha_pdf": line,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.drop_duplicates(subset=["linha_pdf"]).reset_index(drop=True)


def canonical_lookup_name(produto_pdf: str) -> str:
    p = norm(produto_pdf)
    if "ABACAXI" in p:
        return "ABACAXI PEROLA UND"
    if "MELANCIA" in p:
        return "MELANCIA INTEIRA KG"
    return p


def match_base(produto_pdf: str, base_df: pd.DataFrame):
    candidato = canonical_lookup_name(produto_pdf)

    # 1) match no Produto Demarchi
    exact_dem = base_df[base_df["PRODUTO_DEMARCHI_NORM"] == candidato]
    if not exact_dem.empty:
        return exact_dem.iloc[0], "Produto Demarchi exato"

    # 2) contains no Produto Demarchi
    cont_dem = base_df[base_df["PRODUTO_DEMARCHI_NORM"].apply(lambda x: bool(x) and (x in candidato or candidato in x))]
    if not cont_dem.empty:
        return cont_dem.iloc[0], "Produto Demarchi parcial"

    # 3) exact no Produto
    exact_prod = base_df[base_df["PRODUTO_BASE"] == candidato]
    if not exact_prod.empty:
        return exact_prod.iloc[0], "Produto exato"

    # 4) contains no Produto
    cont_prod = base_df[base_df["PRODUTO_BASE"].apply(lambda x: x in candidato or candidato in x)]
    if not cont_prod.empty:
        return cont_prod.iloc[0], "Produto parcial"

    # 5) heurística com palavras relevantes
    words = [w for w in candidato.split() if len(w) >= 4]
    if words:
        scored = []
        for _, row in base_df.iterrows():
            target = f"{row['PRODUTO_BASE']} {row['PRODUTO_DEMARCHI_NORM']}".strip()
            score = sum(1 for w in words if w in target)
            if score:
                scored.append((score, row))
        if scored:
            scored.sort(key=lambda x: x[0], reverse=True)
            return scored[0][1], "Heurística por palavras"

    return None, ""


def force_unit_if_needed(produto_pdf: str, unidade_base: str) -> str:
    p = norm(produto_pdf)
    if any(k in p for k in FORCE_UND_KEYWORDS):
        return "UNIDADE"
    return unidade_base


def convert_to_boxes(qtd: float, unidade_base: str, medida_caixa: float):
    if pd.isna(medida_caixa) or medida_caixa == 0:
        return None

    unidade_base = norm(unidade_base)
    if unidade_base in {"KG", "BANDEJA", "UNIDADE"}:
        return math.ceil(qtd / medida_caixa)
    return None


def build_layout_matrix(df: pd.DataFrame, columns_layout: list) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Produto"] + columns_layout)

    piv = (
        df.groupby(["Produto", "Coluna"], as_index=False)["Caixas"]
        .sum()
        .pivot(index="Produto", columns="Coluna", values="Caixas")
        .fillna(0)
        .reset_index()
        .sort_values("Produto")
        .reset_index(drop=True)
    )

    for col in columns_layout:
        if col not in piv.columns:
            piv[col] = 0

    piv = piv[["Produto"] + columns_layout]
    for col in columns_layout:
        piv[col] = piv[col].astype(int)
    return piv


def write_workbook(sheets: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            data = df if not df.empty else pd.DataFrame(columns=list(df.columns) if hasattr(df, "columns") else ["Produto"])
            data.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.column_dimensions["A"].width = 42
            for col in ["B", "C", "D", "E", "F", "G", "H"]:
                ws.column_dimensions[col].width = 14
    out.seek(0)
    return out.getvalue()


def zip_files(file_map: dict) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, content in file_map.items():
            zf.writestr(filename, content)
    out.seek(0)
    return out.getvalue()


# =========================================================
# UI
# =========================================================

st.title("📦 THOTH PRO Final com Layout")
st.caption("Suba a base table.xlsx e os PDFs. O sistema gera Brasão Frutas, Brasão Legumes, Kross Frutas, Kross Legumes e Brasão CD no layout final.")

base_file = st.file_uploader("Base (Excel)", type=["xlsx", "xls", "csv"])
pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR", type="primary", use_container_width=True):
    if base_file is None:
        st.error("Envie a base Excel antes de processar.")
        st.stop()

    if not pdf_files:
        st.error("Envie pelo menos um PDF.")
        st.stop()

    try:
        base_df = load_base(base_file)

        logs_match = []
        logs_error = []
        logs_files = []
        converted = []

        for pdf in pdf_files:
            pdf_text = extract_pdf_text(pdf)
            store = identify_store(pdf.name, pdf_text)
            items_df = extract_items_from_pdf(pdf_text)

            logs_files.append({
                "Arquivo": pdf.name,
                "Loja": store["nome"],
                "Grupo": store["grupo"],
                "Itens extraídos": len(items_df),
            })

            for _, item in items_df.iterrows():
                base_row, criterio = match_base(item["produto_pdf"], base_df)

                if base_row is None:
                    logs_error.append({
                        "Loja": store["nome"],
                        "Produto PDF": item["produto_pdf"],
                        "Qtd PDF": item["qtd_original"],
                        "Erro": "SEM BASE",
                    })
                    continue

                unidade_final = force_unit_if_needed(item["produto_pdf"], base_row["UNIDADE_NORM"])
                caixas = convert_to_boxes(item["qtd_original"], unidade_final, base_row["Medida"])

                if caixas is None:
                    logs_error.append({
                        "Loja": store["nome"],
                        "Produto PDF": item["produto_pdf"],
                        "Qtd PDF": item["qtd_original"],
                        "Erro": "MEDIDA/UNIDADE INVÁLIDA",
                    })
                    continue

                converted.append({
                    "Grupo": store["grupo"],
                    "Coluna": store["coluna"],
                    "Produto": base_row["PRODUTO_BASE"],
                    "Categoria": base_row["CATEGORIA"],
                    "Caixas": caixas,
                })

                logs_match.append({
                    "Loja": store["nome"],
                    "Produto PDF": item["produto_pdf"],
                    "Produto Base": base_row["PRODUTO_BASE"],
                    "Critério": criterio,
                    "Unidade Base": unidade_final,
                    "Medida Caixa": base_row["Medida"],
                    "Qtd PDF": item["qtd_original"],
                    "Caixas": caixas,
                })

        conv_df = pd.DataFrame(converted)
        match_df = pd.DataFrame(logs_match) if logs_match else pd.DataFrame(columns=["Loja", "Produto PDF", "Produto Base", "Critério", "Unidade Base", "Medida Caixa", "Qtd PDF", "Caixas"])
        error_df = pd.DataFrame(logs_error) if logs_error else pd.DataFrame(columns=["Loja", "Produto PDF", "Qtd PDF", "Erro"])
        files_df = pd.DataFrame(logs_files)

        if conv_df.empty:
            st.error("Nenhum item foi convertido. Confira a base e os PDFs.")
            st.dataframe(error_df, use_container_width=True)
            st.stop()

        brasao = conv_df[conv_df["Grupo"] == "BRASAO"]
        kross = conv_df[conv_df["Grupo"] == "KROSS"]
        cd = conv_df[conv_df["Grupo"] == "CD"]

        brasao_frutas = build_layout_matrix(brasao[brasao["Categoria"] == "FRUTAS"], ["BRASAO F", "BRASAO J", "BRASAO X", "BRASAO A"])
        brasao_legumes = build_layout_matrix(brasao[brasao["Categoria"] == "LEGUMES"], ["BRASAO F", "BRASAO J", "BRASAO X", "BRASAO A"])
        kross_frutas = build_layout_matrix(kross[kross["Categoria"] == "FRUTAS"], ["KROSS AT", "KROSS XAXIM"])
        kross_legumes = build_layout_matrix(kross[kross["Categoria"] == "LEGUMES"], ["KROSS AT", "KROSS XAXIM"])
        cd_frutas = build_layout_matrix(cd[cd["Categoria"] == "FRUTAS"], ["CD"])
        cd_legumes = build_layout_matrix(cd[cd["Categoria"] == "LEGUMES"], ["CD"])

        zip_content = zip_files({
            "BRASAO_FRUTAS_Thoth.xlsx": write_workbook({"BRASAO_FRUTAS": brasao_frutas}),
            "BRASAO_LEGUMES_Thoth.xlsx": write_workbook({"BRASAO_LEGUMES": brasao_legumes}),
            "KROSS_FRUTAS_Thoth.xlsx": write_workbook({"KROSS_FRUTAS": kross_frutas}),
            "KROSS_LEGUMES_Thoth.xlsx": write_workbook({"KROSS_LEGUMES": kross_legumes}),
            "BRASAO_CD.xlsx": write_workbook({"FRUTAS": cd_frutas, "LEGUMES": cd_legumes}),
            "LOG_PROCESSAMENTO.xlsx": write_workbook({"ARQUIVOS": files_df, "MATCHES": match_df, "ERROS": error_df}),
        })

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("PDFs", len(pdf_files))
        c2.metric("Itens convertidos", len(conv_df))
        c3.metric("Erros", len(error_df))
        c4.metric("Matches", len(match_df))

        tab1, tab2, tab3, tab4 = st.tabs(["Prévia", "Matches", "Erros", "Arquivos"])

        with tab1:
            st.write("Brasão Frutas")
            st.dataframe(brasao_frutas, use_container_width=True)
            st.write("Brasão Legumes")
            st.dataframe(brasao_legumes, use_container_width=True)
            st.write("Kross Frutas")
            st.dataframe(kross_frutas, use_container_width=True)
            st.write("Kross Legumes")
            st.dataframe(kross_legumes, use_container_width=True)
            st.write("Brasão CD - Frutas")
            st.dataframe(cd_frutas, use_container_width=True)
            st.write("Brasão CD - Legumes")
            st.dataframe(cd_legumes, use_container_width=True)

        with tab2:
            st.dataframe(match_df, use_container_width=True)

        with tab3:
            st.dataframe(error_df, use_container_width=True)

        with tab4:
            st.dataframe(files_df, use_container_width=True)

        st.download_button(
            "BAIXAR ZIP FINAL",
            data=zip_content,
            file_name="THOTH_FINAL_LAYOUT.zip",
            mime="application/zip",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
