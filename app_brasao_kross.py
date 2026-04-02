import io
import math
import re
import unicodedata
import zipfile
from io import BytesIO
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="BRASÃO / KROSS → THOTH (PRO)",
    page_icon="📦",
    layout="wide",
)

STORE_RULES = [
    {"store_id": "BRASAO_FERNANDO", "group": "BRASAO", "col_key": "1", "display": "Brasão Fernando", "filename_signals": ["FERNANDO", "CE"], "signals": ["FERNANDO"]},
    {"store_id": "BRASAO_JARDIM", "group": "BRASAO", "col_key": "2", "display": "Brasão Jardim", "filename_signals": ["JARDIM", "JA"], "signals": ["JARDIM"]},
    {"store_id": "BRASAO_XAXIM", "group": "BRASAO", "col_key": "3", "display": "Brasão Xaxim", "filename_signals": ["XAXIM", "XX"], "signals": ["XAXIM"]},
    {"store_id": "BRASAO_AVENIDA", "group": "BRASAO", "col_key": "4", "display": "Brasão Avenida", "filename_signals": ["AVENIDA", "AV"], "signals": ["AVENIDA"]},
    {"store_id": "BRASAO_CD", "group": "BRASAO_CD", "col_key": "CD", "display": "Brasão CD", "filename_signals": ["CD"], "signals": ["RUA GASPAR", "ELDORADO", "153", "CENTRO DE DISTRIBUICAO"]},
    {"store_id": "KROSS_CHAPECO", "group": "KROSS", "col_key": "1", "display": "Kross Atacadista", "filename_signals": ["KROSS", "CHAPECO", "ATACADO"], "signals": ["JOHN KENNEDY", "PASSO DOS FORTES", "550", "ATACADO"]},
    {"store_id": "KROSS_XAXIM", "group": "KROSS", "col_key": "2", "display": "Kross Xaxim", "filename_signals": ["KROSS", "XAXIM", "XX"], "signals": ["AMELIO PANIZZI", "XAXIM"]},
]

ORDER_STOP_MARKERS = [
    "AGENDAR A ENTREGA", "PENDENCIAS DE MERCADORIAS", "TOTAL DO FORNECEDOR", "CONTATOS DO FORNECEDOR",
    "COMPRADOR", "TOTAIS", "VALOR TOTAL", "PESO TOTAL", "ORIG DEST TP CODIGO",
]

UNIT_PATTERNS = [
    (r"\bBDJ\b", "BDJ"), (r"\bBANDEJA\b", "BDJ"), (r"\bMACO\b", "UND"),
    (r"\bUNIDADE\b", "UND"), (r"\bUND\b", "UND"), (r"\bUN\b", "UND"), (r"\bKG\b", "KG"),
]

DEMO_BASE = [
    {"produto": "ABACAXI", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 10, "sinonimos": "ABACAXI HAWAI;ABACAXI PEROLA"},
    {"produto": "BLUEBERRY 125G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 12, "sinonimos": "MIRTILO BLUEBERRY;MIRTILO;BLUEBERRY"},
    {"produto": "CEBOLA ALBINA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20, "sinonimos": "CEBOLA ARGENTINA"},
    {"produto": "LARANJA PERA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20, "sinonimos": "LARANJA"},
    {"produto": "LIMAO TAHITI", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20, "sinonimos": "LIMAO;LIMÃO"},
    {"produto": "MAMAO PAPAYA", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 18, "sinonimos": "MAMÃO PAPAYA;PAPAYA"},
    {"produto": "MELAO", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 10, "sinonimos": "MELÃO"},
    {"produto": "MORANGO", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 4, "sinonimos": "MORANGO 250G"},
    {"produto": "PHYSALIS IMPORTADO 100G", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 8, "sinonimos": "PHYSALIS IMPORTADO;PHYSALIS"},
    {"produto": "TOMATE GRAP", "categoria": "LEGUMES", "tipo": "BDJ", "bandejas_por_caixa": 24, "sinonimos": "TOMATE GRAPE DEMARCHI;TOMATE GRAPE;TOMATE GRAPE 180G"},
    {"produto": "UVA THOMPSON 500G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 10, "sinonimos": "UVA THOMPSON S/SEMENTE;UVA THOMPSON S SEMENTE;UVA THOMPSON"},
    {"produto": "CARAMBOLA 400G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 4, "sinonimos": "CARAMBOLA"},
]


def norm_text(v) -> str:
    if v is None:
        return ""
    t = str(v).strip()
    if t.lower() in {"nan", "none"}:
        return ""
    return " ".join(t.split())


def norm_key(v) -> str:
    txt = norm_text(v).upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", txt).strip()


def parse_number(v):
    if v is None:
        return None
    txt = norm_text(v)
    if not txt:
        return None
    txt = txt.replace(".", "").replace(",", ".")
    txt = re.sub(r"[^0-9.-]", "", txt)
    if txt in {"", ".", "-", "-."}:
        return None
    try:
        val = float(txt)
        if math.isnan(val):
            return None
        return val
    except Exception:
        return None


def infer_unit(text: str) -> str:
    t = norm_key(text)
    for pattern, unit in UNIT_PATTERNS:
        if re.search(pattern, t):
            return unit
    return ""


def normalize_product_name(raw: str) -> str:
    t = norm_key(raw)
    for pat in [r"\bDEMARCHI\b", r"\bSHELF\s*\d+\b", r"\bKG\b", r"\bBDJ\b", r"\bBANDEJA\b", r"\bUND\b", r"\bUNIDADE\b", r"\bUN\b", r"\b\d+G\b", r"\bC/\d+G\b"]:
        t = re.sub(pat, " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    exact_map = {
        "TOMATE GRAPE DEMARCHI": "TOMATE GRAP", "TOMATE GRAPE": "TOMATE GRAP",
        "UVA THOMPSON S/SEMENTE": "UVA THOMPSON 500G", "UVA THOMPSON S SEMENTE": "UVA THOMPSON 500G",
        "MIRTILO BLUEBERRY": "BLUEBERRY 125G", "BLUEBERRY": "BLUEBERRY 125G",
        "CEBOLA ARGENTINA": "CEBOLA ALBINA", "PHYSALIS IMPORTADO": "PHYSALIS IMPORTADO 100G",
        "PHYSALIS": "PHYSALIS IMPORTADO 100G", "CARAMBOLA": "CARAMBOLA 400G",
    }
    if t in exact_map:
        return exact_map[t]
    if "TOMATE" in t and "GRAPE" in t:
        return "TOMATE GRAP"
    if "UVA" in t and "THOMPSON" in t:
        return "UVA THOMPSON 500G"
    if "MIRTILO" in t or "BLUEBERRY" in t:
        return "BLUEBERRY 125G"
    if "CEBOLA" in t and "ARGENTINA" in t:
        return "CEBOLA ALBINA"
    if "PHYSALIS" in t:
        return "PHYSALIS IMPORTADO 100G"
    if "CARAMBOLA" in t:
        return "CARAMBOLA 400G"
    return t


def detect_category_from_name(name: str) -> str:
    k = norm_key(name)
    if "FRUTA" in k:
        return "FRUTAS"
    if "LEGUME" in k or "VERDURA" in k:
        return "LEGUMES"
    return ""


def load_base(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(DEMO_BASE)
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def build_base_map(base_df: pd.DataFrame) -> Dict[str, dict]:
    result = {}
    for _, row in base_df.iterrows():
        produto = row.get("produto", row.get("PRODUTO", row.get("produto_base", row.get("PRODUTO_BASE", ""))))
        produto_norm = normalize_product_name(produto)
        if not produto_norm:
            continue
        item = {
            "produto": produto_norm,
            "categoria": norm_key(row.get("categoria", row.get("CATEGORIA", ""))),
            "tipo": norm_key(row.get("tipo", row.get("TIPO", row.get("modo_conversao", row.get("MODO_CONVERSAO", ""))))),
            "peso_caixa": parse_number(row.get("peso_caixa", row.get("PESO_CAIXA", 0))) or 0,
            "bandejas_por_caixa": parse_number(row.get("bandejas_por_caixa", row.get("BANDEJAS_POR_CAIXA", 0))) or 0,
            "unidades_caixa": parse_number(row.get("unidades_caixa", row.get("UNIDADES_CAIXA", row.get("itens_por_caixa", row.get("ITENS_POR_CAIXA", 0))))) or 0,
        }
        result[produto_norm] = item
        syn = str(row.get("sinonimos", row.get("SINONIMOS", "")) or "")
        for s in [x.strip() for x in syn.split(";") if x.strip()]:
            result[normalize_product_name(s)] = item
    return result


def identify_store(text: str, filename: str) -> dict:
    hay = f"{norm_key(filename)}\n{norm_key(text)}"
    best_rule = STORE_RULES[0]
    best_score = -1
    for rule in STORE_RULES:
        score = 0
        for signal in rule["filename_signals"]:
            if norm_key(signal) in norm_key(filename):
                score += 5
        for signal in rule["signals"]:
            if norm_key(signal) in hay:
                score += 2
        if rule["group"] == "BRASAO_CD" and "CD" in norm_key(filename):
            score += 10
        if score > best_score:
            best_score = score
            best_rule = rule
    return best_rule


def pdf_to_text(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def line_irrelevant(line: str) -> bool:
    t = norm_key(line)
    if len(t) < 3:
        return True
    if any(marker in t for marker in ORDER_STOP_MARKERS):
        return True
    noise = ["PAGINA", "PAG.", "EMISSAO", "DATA", "HORA", "CNPJ", "IE", "ENDERECO", "CLIENTE", "VENDEDOR", "CONDICAO", "TRANSPORTE", "OBSERVACAO", "TOTAL", "SUBTOTAL", "DESCONTO", "CODIGO", "VALOR", "PRECO"]
    return any(n in t for n in noise)


def parse_order_items(text: str) -> pd.DataFrame:
    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line_irrelevant(line):
            continue
        upper = norm_key(line)
        patterns = [
            r"^(?P<produto>.+?)\s+(?P<qtd>\d[\d\.,]*)\s*(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN)\b",
            r"^(?P<produto>.+?)\s*[-–]\s*(?P<qtd>\d[\d\.,]*)\s*(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN)\b",
            r"^(?P<produto>.+?)\s+(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN)\s*(?P<qtd>\d[\d\.,]*)\b",
        ]
        match = None
        for p in patterns:
            match = re.search(p, upper)
            if match:
                break
        produto, qtd, unit = None, None, ""
        if match:
            produto = match.group("produto").strip(" -")
            qtd = parse_number(match.group("qtd"))
            unit = norm_key(match.group("unit"))
        else:
            nums = re.findall(r"\d[\d\.,]*", upper)
            if nums:
                qtd = parse_number(nums[-1])
                produto = upper.replace(nums[-1], " ").strip(" -")
                unit = infer_unit(upper)
        if not produto or qtd is None or qtd <= 0:
            continue
        if unit == "BANDEJA":
            unit = "BDJ"
        if unit in {"UNIDADE", "UN"}:
            unit = "UND"
        rows.append({"produto_original": produto, "qtd_original": qtd, "tipo_detectado": unit or infer_unit(produto)})
    return pd.DataFrame(rows)


def convert_item(produto_original: str, qtd: float, tipo_detectado: str, base_map: Dict[str, dict]) -> Tuple[dict, str]:
    produto_norm = normalize_product_name(produto_original)
    base = base_map.get(produto_norm)
    if not base:
        return {"produto_base": produto_norm, "caixas": 0, "categoria": "", "status": "SEM BASE", "tipo_usado": tipo_detectado}, "Produto não encontrado na base"
    categoria = base["categoria"]
    tipo = base["tipo"] or tipo_detectado
    if tipo == "KG":
        fator = base["peso_caixa"]
    elif tipo == "BDJ":
        fator = base["bandejas_por_caixa"]
    elif tipo in {"UND", "UN"}:
        tipo = "UND"
        fator = base["unidades_caixa"]
    else:
        fator = 0
    if fator <= 0:
        return {"produto_base": base["produto"], "caixas": 0, "categoria": categoria, "status": "BASE INCOMPLETA", "tipo_usado": tipo}, "Fator obrigatório ausente para conversão"
    caixas = math.ceil(qtd / fator)
    return {"produto_base": base["produto"], "caixas": int(caixas), "categoria": categoria, "status": "OK", "tipo_usado": tipo}, ""


def transform_items(order_df: pd.DataFrame, store_rule: dict, base_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    base_map = build_base_map(base_df)
    good_rows, error_rows = [], []
    for _, row in order_df.iterrows():
        produto, qtd, tipo = row["produto_original"], row["qtd_original"], row["tipo_detectado"]
        conv, err = convert_item(produto, qtd, tipo, base_map)
        categoria = conv["categoria"] or detect_category_from_name(produto)
        if conv["status"] == "OK":
            good_rows.append({"Grupo": store_rule["group"], "LojaID": store_rule["store_id"], "Loja": store_rule["display"], "Coluna": store_rule["col_key"], "Produto": conv["produto_base"], "Categoria": categoria, "QtdOriginal": qtd, "Caixas": conv["caixas"], "Tipo": conv["tipo_usado"]})
        else:
            error_rows.append({"loja": store_rule["display"], "produto": produto, "erro": err})
    return pd.DataFrame(good_rows), pd.DataFrame(error_rows)


def group_to_matrix(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    pivot = df.groupby(["Produto", "Coluna"], as_index=False)["Caixas"].sum().pivot(index="Produto", columns="Coluna", values="Caixas").fillna(0)
    pivot.columns = [str(c) for c in pivot.columns]
    pivot = pivot.reset_index().sort_values("Produto").reset_index(drop=True)
    for c in pivot.columns:
        if c != "Produto":
            pivot[c] = pivot[c].astype(int)
    return pivot


def write_output(matrix_df: pd.DataFrame, label: str) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (matrix_df if not matrix_df.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name=label, index=False)
        ws = writer.sheets[label]
        ws.column_dimensions["A"].width = 40
        for col_letter in ["B", "C", "D", "E", "F", "G", "H"]:
            ws.column_dimensions[col_letter].width = 12
    out.seek(0)
    return out.getvalue()


def build_cd_workbook(frutas_matrix: pd.DataFrame, legumes_matrix: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (frutas_matrix if not frutas_matrix.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name="FRUTAS", index=False)
        (legumes_matrix if not legumes_matrix.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name="LEGUMES", index=False)
        for sh in ["FRUTAS", "LEGUMES"]:
            ws = writer.sheets[sh]
            ws.column_dimensions["A"].width = 40
            for col_letter in ["B", "C", "D", "E", "F", "G", "H"]:
                ws.column_dimensions[col_letter].width = 12
    out.seek(0)
    return out.getvalue()


def build_prices_sheet(df: pd.DataFrame) -> bytes:
    rows = []
    if not df.empty:
        base = df[["Produto", "Categoria"]].drop_duplicates().sort_values(["Categoria", "Produto"])
        for _, r in base.iterrows():
            rows.append({"CATEGORIA": r["Categoria"], "PRODUTO": r["Produto"], "COD_FORN": "", "PRECO": ""})
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="PRECOS", index=False)
        ws = writer.sheets["PRECOS"]
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 14
    out.seek(0)
    return out.getvalue()


def build_zip(files_dict: dict) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files_dict.items():
            zf.writestr(name, content)
    out.seek(0)
    return out.getvalue()


st.title("📦 BRASÃO / KROSS → THOTH (PRO)")
st.caption("Suba apenas os PDFs. Clique em PROCESSAR. Baixe o ZIP com Brasão Frutas, Brasão Legumes, Kross Frutas, Kross Legumes e Brasão CD.")

with st.sidebar:
    st.subheader("Como usar")
    st.write("1. Envie os PDFs do dia.")
    st.write("2. Clique em PROCESSAR.")
    st.write("3. Baixe o ZIP final.")
    st.info("Fluxos suportados: Brasão lojas, Brasão CD e Kross.")
    usar_base_modelo = st.checkbox("Usar base modelo embutida", value=True)
    base_file = st.file_uploader("Base mestre (opcional)", type=["xlsx", "xls", "csv"])
    pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

base_df = pd.DataFrame(DEMO_BASE) if usar_base_modelo and base_file is None else load_base(base_file) if base_file is not None else pd.DataFrame(DEMO_BASE)

if st.button("PROCESSAR", use_container_width=True, type="primary"):
    if not pdf_files:
        st.error("Envie ao menos um PDF.")
    else:
        try:
            transformed_parts, all_errors, identified, seen_store_ids = [], [], [], set()
            for pdf in pdf_files:
                text = pdf_to_text(pdf.getvalue())
                store_rule = identify_store(text, pdf.name)
                if store_rule["store_id"] in seen_store_ids:
                    all_errors.append(pd.DataFrame([{"loja": store_rule["display"], "produto": pdf.name, "erro": "PDF duplicado da mesma unidade ignorado"}]))
                    continue
                seen_store_ids.add(store_rule["store_id"])
                order_df = parse_order_items(text)
                transformed_df, errors_df = transform_items(order_df, store_rule, base_df)
                if not transformed_df.empty:
                    transformed_parts.append(transformed_df)
                if not errors_df.empty:
                    all_errors.append(errors_df)
                identified.append({"arquivo": pdf.name, "loja": store_rule["display"], "grupo": store_rule["group"], "itens_extraidos": len(order_df), "itens_convertidos": len(transformed_df)})
            if not transformed_parts:
                raise ValueError("Nenhum item foi convertido. Verifique a base e os PDFs.")
            all_data = pd.concat(transformed_parts, ignore_index=True)
            errors_data = pd.concat(all_errors, ignore_index=True) if all_errors else pd.DataFrame(columns=["loja", "produto", "erro"])
            identified_df = pd.DataFrame(identified)
            brasao_df = all_data[all_data["Grupo"] == "BRASAO"].copy()
            kross_df = all_data[all_data["Grupo"] == "KROSS"].copy()
            cd_df = all_data[all_data["Grupo"] == "BRASAO_CD"].copy()
            brasao_frutas_matrix = group_to_matrix(brasao_df[brasao_df["Categoria"] == "FRUTAS"])
            brasao_legumes_matrix = group_to_matrix(brasao_df[brasao_df["Categoria"] == "LEGUMES"])
            kross_frutas_matrix = group_to_matrix(kross_df[kross_df["Categoria"] == "FRUTAS"])
            kross_legumes_matrix = group_to_matrix(kross_df[kross_df["Categoria"] == "LEGUMES"])
            cd_frutas_matrix = group_to_matrix(cd_df[cd_df["Categoria"] == "FRUTAS"])
            cd_legumes_matrix = group_to_matrix(cd_df[cd_df["Categoria"] == "LEGUMES"])
            files_to_zip = {
                "BRASAO_FRUTAS_Thoth.xlsx": write_output(brasao_frutas_matrix, "BRASAO_FRUTAS"),
                "BRASAO_LEGUMES_Thoth.xlsx": write_output(brasao_legumes_matrix, "BRASAO_LEGUMES"),
                "KROSS_FRUTAS_Thoth.xlsx": write_output(kross_frutas_matrix, "KROSS_FRUTAS"),
                "KROSS_LEGUMES_Thoth.xlsx": write_output(kross_legumes_matrix, "KROSS_LEGUMES"),
                "BRASAO_CD.xlsx": build_cd_workbook(cd_frutas_matrix, cd_legumes_matrix),
                "BRASAO_PRECOS.xlsx": build_prices_sheet(brasao_df),
                "KROSS_PRECOS.xlsx": build_prices_sheet(kross_df),
                "BRASAO_CD_PRECOS.xlsx": build_prices_sheet(cd_df),
            }
            expected_ids = {"BRASAO_FERNANDO", "BRASAO_JARDIM", "BRASAO_XAXIM", "BRASAO_AVENIDA", "BRASAO_CD", "KROSS_CHAPECO", "KROSS_XAXIM"}
            missing_units = []
            for rule in STORE_RULES:
                if rule["store_id"] in expected_ids and rule["store_id"] not in seen_store_ids:
                    missing_units.append({"loja": rule["display"], "produto": "", "erro": "PDF da unidade não enviado"})
            if missing_units:
                errors_data = pd.concat([errors_data, pd.DataFrame(missing_units)], ignore_index=True)
            log_out = BytesIO()
            with pd.ExcelWriter(log_out, engine="openpyxl") as writer:
                identified_df.to_excel(writer, sheet_name="ARQUIVOS", index=False)
                errors_data.to_excel(writer, sheet_name="ERROS", index=False)
            log_out.seek(0)
            files_to_zip["LOG_PROCESSAMENTO.xlsx"] = log_out.getvalue()
            zip_bytes = build_zip(files_to_zip)
            st.success("Processamento concluído com sucesso.")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("PDFs enviados", len(pdf_files))
            c2.metric("Unidades identificadas", len(identified_df))
            c3.metric("Itens convertidos", len(all_data))
            c4.metric("Ocorrências no log", len(errors_data))
            st.subheader("Arquivos processados")
            st.dataframe(identified_df, use_container_width=True)
            if not errors_data.empty:
                st.subheader("Log de ocorrências")
                st.dataframe(errors_data, use_container_width=True)
            st.download_button("Baixar ZIP final", zip_bytes, file_name="THOTH_BRASAO_KROSS_PRO.zip", mime="application/zip", use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
