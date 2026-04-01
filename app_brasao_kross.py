from io import BytesIO
from pathlib import Path
from copy import copy
import re
import traceback
import zipfile

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


st.set_page_config(page_title="BRASÃO / KROSS → THOTH (PRO)", page_icon="📦", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
IGNORE_NAMES = {"", "TOTAL", "TOTAIS", "SUBTOTAL", "SUB-TOTAL", "PRODUTO", "PRODUTOS"}

FIXED_FILE_CANDIDATES = {
    "MODEL_BRASAO_FRUTAS": [
        "BRASAO - FRUTAS PRE PEDIDO BRANCO.xlsx",
        "BRASAO FRUTAS BRANCO.xlsx",
    ],
    "MODEL_BRASAO_LEGUMES": [
        "BRASAO - LEGUMES PRE PEDIDO BRANCO.xlsx",
        "BRASAO LEGUMES BRANCO.xlsx",
    ],
    "MODEL_KROSS_FRUTAS": [
        "KROSS - FRUTAS PRE PEDIDO BRANCO.xlsx",
        "KROSS - PRE PEDIDO FRUTAS BRANCO.xlsx",
    ],
    "MODEL_KROSS_LEGUMES": [
        "KROSS - LEGUMES PRE PEDIDO BRANCO.xlsx",
    ],
    "BASE_PRODUCTS_FILE": [
        "base_thoth_app_normalizada_brasao_kross.xlsx",
    ],
}

STORE_RULES = [
    {
        "store_id": "BRASAO_FERNANDO",
        "group": "BRASAO",
        "col_key": "1",
        "display": "Brasão Fernando",
        "signals": ["FERNANDO MACHADO", "CENTRO", "226"],
        "file_signals": ["FERNANDO"],
    },
    {
        "store_id": "BRASAO_JARDIM",
        "group": "BRASAO",
        "col_key": "2",
        "display": "Brasão Jardim",
        "signals": ["SAO PEDRO", "JARDIM AMERICA", "2199"],
        "file_signals": ["JARDIM"],
    },
    {
        "store_id": "BRASAO_XAXIM",
        "group": "BRASAO",
        "col_key": "3",
        "display": "Brasão Xaxim",
        "signals": ["LUIZ LUNARDI", "XAXIM", "810"],
        "file_signals": ["XAXIM"],
    },
    {
        "store_id": "BRASAO_AVENIDA",
        "group": "BRASAO",
        "col_key": "4",
        "display": "Brasão Avenida",
        "signals": ["RIO DE JANEIRO", "CENTRO", "108", "CHAPECO"],
        "file_signals": ["AVENIDA"],
    },
    {
        "store_id": "BRASAO_CD",
        "group": "BRASAO_CD",
        "col_key": "CD",
        "display": "Brasão CD",
        "signals": ["RUA GASPAR", "ELDORADO", "153"],
        "file_signals": ["CD"],
    },
    {
        "store_id": "KROSS_CHAPECO",
        "group": "KROSS",
        "col_key": "1",
        "display": "Kross Atacadista",
        "signals": ["JOHN KENNEDY", "PASSO DOS FORTES", "550"],
        "file_signals": ["ATACADO", "CHAPECO"],
    },
    {
        "store_id": "KROSS_XAXIM",
        "group": "KROSS",
        "col_key": "2",
        "display": "Kross Xaxim",
        "signals": ["AMELIO PANIZZI", "XAXIM"],
        "file_signals": ["XAXIM"],
    },
]

ORDER_STOP_MARKERS = [
    "AGENDAR A ENTREGA",
    "PENDENCIAS DE MERCADORIAS",
    "TOTAL DO FORNECEDOR",
    "CONTATOS DO FORNECEDOR",
    "COMPRADOR",
    "TOTAIS",
    "VALOR TOTAL",
    "PESO TOTAL",
    "ORIG DEST TP CODIGO",
]

UNIT_PATTERNS = [
    (r"\bBDJ\b", "BDJ"),
    (r"\bBANDEJA\b", "BDJ"),
    (r"\bMACO\b", "UND"),
    (r"\bUNIDADE\b", "UND"),
    (r"\bUND\b", "UND"),
    (r"\bKG\b", "KG"),
]


NOISE_TOKENS = {
    "BRASAO", "FRUTA", "FRUTAS", "LEGUME", "LEGUMES", "KROSS", "DE", "DO", "DA",
    "DEMARCHI", "MARCHI", "SHELF", "IMPORTADO", "IMPORTADA", "NACIONAL", "KG", "UND", "BDJ",
    "BANDEJA", "UNIDADE", "UNIDADES", "MACO", "MACOS", "BRS", "CX", "CAIXA", "CAIXAS"
}

CATEGORY_HINTS = {
    "FRUTAS": {"ABACAXI","ABACATE","AMEIXA","AMORA","BANANA","CAQUI","COCO","FIGO","FRAMBOESA","GOIABA","JACA","KIWI","LARANJA","LIMAO","MACA","MAMAO","MANGA","MELANCIA","MELAO","MEXERICA","MORANGO","NECTARINA","PERA","PESSEGO","PITAYA","TANGERINA","UVA","MARACUJA"},
    "LEGUMES": {"ABOBORA","ABOBRINHA","ALHO","BATATA","BATATA DOCE","BATATINHA","BERINJELA","BETERRABA","BROCOLIS","BROCOLIS NINJA","BROCOLIS RAMO","CEBOLA","CENOURA","CHUCHU","COUVE","COUVE FLOR","GENGIBRE","INHAME","MANDIOCA","MILHO","PEPINO","PIMENTA","PIMENTAO","QUIABO","RABANETE","REPOLHO","TOMATE","VAGEM"}
}


def norm_text(v):
    if v is None:
        return ""
    t = str(v).strip()
    if t.lower() in {"nan", "none"}:
        return ""
    return " ".join(t.split())


def norm_key(v):
    txt = norm_text(v).upper()
    txt = txt.replace("Á", "A").replace("À", "A").replace("Ã", "A").replace("Â", "A")
    txt = txt.replace("É", "E").replace("Ê", "E")
    txt = txt.replace("Í", "I")
    txt = txt.replace("Ó", "O").replace("Õ", "O").replace("Ô", "O")
    txt = txt.replace("Ú", "U")
    txt = txt.replace("Ç", "C")
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
        return float(txt)
    except Exception:
        return None


def is_missing_number(v):
    if v is None:
        return True
    try:
        return pd.isna(v)
    except Exception:
        return False


def ceil_div(qtd, base):
    if is_missing_number(qtd) or is_missing_number(base):
        return 0
    try:
        qtd = float(qtd)
        base = float(base)
    except Exception:
        return 0
    if base <= 0 or qtd <= 0:
        return 0
    return int(-(-qtd // base))


def resolve_first_existing(candidates):
    for filename in candidates:
        path = BASE_DIR / filename
        if path.exists():
            return path
    return None


@st.cache_data(show_spinner=False)
def load_base_from_disk(path_str: str) -> pd.DataFrame:
    file_path = Path(path_str)
    if not file_path.exists():
        raise FileNotFoundError(f"Base de produtos não encontrada: {file_path.name}")

    df = pd.read_excel(file_path, sheet_name=0)
    cols_map = {c: norm_key(c) for c in df.columns}

    def pick(col_options, required=False):
        for original, key in cols_map.items():
            if key in col_options:
                return original
        if required:
            raise ValueError(f"Coluna obrigatória não encontrada na base: {col_options}")
        return None

    col_categoria = pick({"CATEGORIA"}, required=True)
    col_produto = pick({"PRODUTO_BASE", "PRODUTO", "ITEM", "DESCRICAO", "DESCRICAO BASE"}, required=True)
    col_sinonimos = pick({"SINONIMOS", "SINONIMO", "APELIDOS"})
    col_codigo = pick({"CODIGO", "COD", "COD ITEM"})
    col_cod_forn = pick({"COD FORN", "COD_FORN", "CODFORN", "CODIGO FORNECEDOR"})
    col_modo = pick({"MODO_CONVERSAO", "MODO CONVERSAO", "MODO"})
    col_peso = pick({"PESO_CAIXA", "PESO CAIXA", "KG POR CAIXA"})
    col_itens = pick({"ITENS_POR_CAIXA", "ITENS POR CAIXA", "UN POR CAIXA", "UND POR CAIXA"})
    col_bdj = pick({"BANDEJAS_POR_CAIXA", "BANDEJAS POR CAIXA", "BDJ POR CAIXA"})

    rows = []
    for _, r in df.iterrows():
        produto = norm_text(r.get(col_produto))
        if not produto:
            continue

        sinonimos_raw = norm_text(r.get(col_sinonimos)) if col_sinonimos else ""
        sinonimos = []
        if sinonimos_raw:
            for part in re.split(r"[|;]", sinonimos_raw):
                p = norm_text(part)
                if p:
                    sinonimos.append(p)

        rows.append({
            "categoria": norm_key(r.get(col_categoria)),
            "produto_base": produto,
            "produto_key": norm_key(produto),
            "sinonimos": sinonimos,
            "sinonimos_key": [norm_key(x) for x in sinonimos],
            "codigo": norm_text(r.get(col_codigo)) if col_codigo else "",
            "cod_forn": norm_text(r.get(col_cod_forn)) if col_cod_forn else "",
            "modo": norm_key(r.get(col_modo)) if col_modo else "",
            "peso_caixa": parse_number(r.get(col_peso)) if col_peso else None,
            "itens_por_caixa": parse_number(r.get(col_itens)) if col_itens else None,
            "bandejas_por_caixa": parse_number(r.get(col_bdj)) if col_bdj else None,
        })

    base_df = pd.DataFrame(rows)
    if base_df.empty:
        raise ValueError("A base de produtos está vazia.")
    return base_df


@st.cache_data(show_spinner=False)
def pdf_to_text(file_bytes: bytes) -> str:
    if pdfplumber is None:
        raise RuntimeError("Falta instalar pdfplumber no ambiente. Adicione em requirements.txt")

    all_text = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            all_text.append(txt)
    return "\n".join(all_text)


def identify_store(pdf_text: str, file_name: str = ""):
    text = norm_key(pdf_text)
    name = norm_key(Path(file_name).stem)

    # prioridade pelo nome do arquivo
    if "KROSS" in name:
        if "XAXIM" in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "KROSS_XAXIM")
        if any(token in name for token in ["ATACADO", "ATACADISTA", "CHAPECO"]):
            return next(rule for rule in STORE_RULES if rule["store_id"] == "KROSS_CHAPECO")

    if "BRASAO" in name or "BRASAO" not in name and any(token in name for token in ["FERNANDO", "JARDIM", "AVENIDA", "CD"]):
        if "CD" in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "BRASAO_CD")
        if "FERNANDO" in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "BRASAO_FERNANDO")
        if "JARDIM" in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "BRASAO_JARDIM")
        if "AVENIDA" in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "BRASAO_AVENIDA")
        if "XAXIM" in name and "KROSS" not in name:
            return next(rule for rule in STORE_RULES if rule["store_id"] == "BRASAO_XAXIM")

    # fallback pelo conteúdo do PDF
    for rule in STORE_RULES:
        if all(signal in text for signal in rule["signals"]):
            return rule

    raise ValueError(f"Não consegui identificar a loja/unidade pelo cabeçalho do PDF: {file_name}")


def detect_unit(desc: str) -> str:
    key = norm_key(desc)
    for pattern, unit in UNIT_PATTERNS:
        if re.search(pattern, key):
            return unit
    return "CX"


def clean_desc(desc: str) -> str:
    return re.sub(r"\s+", " ", norm_text(desc)).strip()


def is_stop_line(line: str) -> bool:
    key = norm_key(line)
    return any(marker in key for marker in ORDER_STOP_MARKERS)


def parse_order_items(pdf_text: str) -> pd.DataFrame:
    lines = [norm_text(x) for x in pdf_text.splitlines() if norm_text(x)]
    item_lines = []
    start_collecting = False

    row_pattern = re.compile(
        r"^(?P<codigo>\d[\d.,]*)\s+"
        r"(?P<cod_forn>[\d,./-]+)\s+"
        r"(?P<descricao>.+?)\s+"
        r"(?P<quant>\d+[\d.,]*)\s+"
        r"(?P<qtde_emb>\d+[\d.,]*)\s+"
        r"(?P<pr_unit>\d+[\d.,]*)\s+"
        r"(?P<vl_total>\d+[\d.,]*)$"
    )

    for line in lines:
        key = norm_key(line)
        if "CODIGO COD FORN DESCRICAO" in key or "CODIGO COD FORN DESCRI" in key:
            start_collecting = True
            continue

        if not start_collecting:
            continue

        if is_stop_line(line):
            break

        match = row_pattern.match(line)
        if not match:
            continue

        descricao = clean_desc(match.group("descricao"))
        item_lines.append({
            "CodigoPedido": norm_text(match.group("codigo")),
            "CodFornPedido": norm_text(match.group("cod_forn")),
            "DescricaoOriginal": descricao,
            "Qtde": parse_number(match.group("quant")) or 0,
            "QtdeEmb": parse_number(match.group("qtde_emb")) or 0,
            "PrecoPedido": parse_number(match.group("pr_unit")),
            "ValorTotal": parse_number(match.group("vl_total")),
            "UnidadeDetectada": detect_unit(descricao),
        })

    df = pd.DataFrame(item_lines)
    if df.empty:
        raise ValueError("Nenhum item válido foi extraído do PDF. Verifique se o layout do pedido segue o padrão atual.")
    return df


def tokenize_desc(desc: str):
    key = norm_key(desc)
    key = re.sub(r"\b\d+[.,]?\d*G\b", " ", key)
    key = re.sub(r"\b\d+[.,]?\d*KG\b", " ", key)
    key = re.sub(r"\b\d+[.,]?\d*ML\b", " ", key)
    key = re.sub(r"\b\d+[.,]?\d*L\b", " ", key)
    key = re.sub(r"[^A-Z0-9 ]", " ", key)
    tokens = [t for t in key.split() if t and t not in NOISE_TOKENS and not t.isdigit()]
    return tokens


def cleaned_product_name(desc: str) -> str:
    key = norm_key(desc)
    key = re.sub(r"\bSHELF\s*\d+\b", " ", key)
    key = re.sub(r"\b\d+[.,]?\d*G\b", " ", key)
    key = re.sub(r"\b\d+[.,]?\d*ML\b", " ", key)
    key = re.sub(r"\b(IMPORTADO|IMPORTADA|NACIONAL|DEMARCHI|MARCHI|BRASAO|KROSS|FRUTA|FRUTAS|LEGUME|LEGUMES)\b", " ", key)
    key = re.sub(r"\s+", " ", key).strip()
    return key.title() if key else clean_desc(desc)


def infer_category_from_desc(desc: str, base_df=None) -> str:
    key = norm_key(desc)
    best_cat = None
    best_score = 0
    tokens = set(tokenize_desc(desc))
    if base_df is not None and tokens:
        for _, r in base_df.iterrows():
            rtokens = set(tokenize_desc(r["produto_base"]))
            if not rtokens:
                continue
            score = len(tokens & rtokens)
            if score > best_score:
                best_score = score
                best_cat = norm_key(r["categoria"])
        if best_score >= 1 and best_cat:
            return best_cat

    for cat, words in CATEGORY_HINTS.items():
        for word in words:
            if word in key:
                return cat
    return "FRUTAS"


def match_base_item(desc: str, base_df: pd.DataFrame):
    key = norm_key(desc)

    exact = base_df[base_df["produto_key"] == key]
    if not exact.empty:
        return exact.iloc[0]

    for _, row in base_df.iterrows():
        if key in row["sinonimos_key"]:
            return row

    desc_tokens = set(tokenize_desc(desc))
    if not desc_tokens:
        return None

    # contains mais seguro: exige pelo menos 2 tokens relevantes em comum
    contains_candidates = []
    for _, row in base_df.iterrows():
        row_key = row["produto_key"]
        row_tokens = set(tokenize_desc(row["produto_base"]))
        overlap = len(desc_tokens & row_tokens)
        if overlap >= 2 and (row_key in key or key in row_key):
            contains_candidates.append((overlap, len(row_key), row))
        for syn in row.get("sinonimos", []) or []:
            syn_key = norm_key(syn)
            syn_tokens = set(tokenize_desc(syn))
            syn_overlap = len(desc_tokens & syn_tokens)
            if syn_overlap >= 2 and (syn_key in key or key in syn_key):
                contains_candidates.append((syn_overlap, len(syn_key), row))

    if contains_candidates:
        contains_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return contains_candidates[0][2]

    # similaridade por tokens: bem mais conservadora para não converter produto errado
    best_row = None
    best_overlap = 0
    best_ratio = 0.0

    for _, row in base_df.iterrows():
        options = [row["produto_base"]] + list(row.get("sinonimos", []) or [])
        for option in options:
            row_tokens = set(tokenize_desc(option))
            if not row_tokens:
                continue
            overlap = len(desc_tokens & row_tokens)
            ratio = overlap / max(len(desc_tokens), len(row_tokens))
            if overlap > best_overlap or (overlap == best_overlap and ratio > best_ratio):
                best_overlap = overlap
                best_ratio = ratio
                best_row = row

    if best_row is not None:
        # 2+ tokens e cobertura mínima razoável
        if best_overlap >= 2 and best_ratio >= 0.5:
            return best_row
        # caso de nome curto e muito específico
        if len(desc_tokens) == 1 and best_overlap == 1 and best_ratio == 1:
            return best_row

    return None
def infer_boxes_without_base(row):
    unit = norm_key(row.get("UnidadeDetectada"))
    qtd = row.get("Qtde")
    emb = row.get("QtdeEmb")
    if is_missing_number(qtd):
        return 0, "quantidade ausente"
    try:
        qtd = float(qtd)
    except Exception:
        return 0, "quantidade invalida"
    if qtd <= 0:
        return 0, "quantidade zerada"

    # Sem base, só faz conversão automática quando existir fator confiável no PDF.
    if not is_missing_number(emb):
        try:
            emb = float(emb)
        except Exception:
            emb = None
        if emb and emb > 0:
            return ceil_div(qtd, emb), f"conversao automatica via QtdeEmb ({unit or 'SEM_UNIDADE'})"

    # Tratar como caixa direta só quando o PDF já vier claramente em caixas.
    if unit == "CX":
        return int(qtd), "quantidade ja em caixas"

    return 0, f"sem base e sem fator de conversao ({unit or 'SEM_UNIDADE'})"
def convert_to_boxes(qtd: float, unit: str, base_row) -> int:
    modo = norm_key(base_row.get("modo", ""))
    unit = norm_key(unit)

    if is_missing_number(qtd):
        return 0

    try:
        qtd = float(qtd)
    except Exception:
        return 0

    if qtd <= 0:
        return 0

    peso_caixa = base_row.get("peso_caixa")
    itens_por_caixa = base_row.get("itens_por_caixa")
    bandejas_por_caixa = base_row.get("bandejas_por_caixa")

    # Prioridade total para o modo da base. Isso evita erro tipo Laranja Kinkan virar 400 caixas.
    if modo == "CAIXA":
        return int(qtd)
    if modo == "PESO":
        return ceil_div(qtd, peso_caixa)
    if modo in {"UNIDADE", "UND"}:
        return ceil_div(qtd, itens_por_caixa)
    if modo in {"BANDEJA", "BDJ"}:
        return ceil_div(qtd, bandejas_por_caixa)

    # Sem modo definido na base, cai para a unidade detectada no PDF.
    if unit == "CX":
        return int(qtd)
    if unit == "KG":
        return ceil_div(qtd, peso_caixa)
    if unit == "UND":
        return ceil_div(qtd, itens_por_caixa)
    if unit == "BDJ":
        return ceil_div(qtd, bandejas_por_caixa)
    return 0
def transform_items(order_df: pd.DataFrame, store_rule: dict, base_df: pd.DataFrame):
    out_rows = []
    errors = []

    for _, row in order_df.iterrows():
        base_row = match_base_item(row["DescricaoOriginal"], base_df)
        conversion_note = ""

        if base_row is not None:
            caixas = convert_to_boxes(row["Qtde"], row["UnidadeDetectada"], base_row)
            produto_modelo = base_row["produto_base"]
            produto_key = base_row["produto_key"]
            categoria = norm_key(base_row["categoria"])

            if caixas <= 0:
                # Se a base existe mas o fator está faltando, tenta QtdeEmb sem trocar o produto.
                caixas, conversion_note = infer_boxes_without_base(row)
                if caixas <= 0:
                    errors.append({
                        "loja": store_rule["display"],
                        "produto": row["DescricaoOriginal"],
                        "erro": f"Falha na conversão para caixa ({row['UnidadeDetectada']})",
                    })
                    continue
                conversion_note = f"fallback do PDF para produto cadastrado: {conversion_note}"
        else:
            caixas, conversion_note = infer_boxes_without_base(row)
            if caixas <= 0:
                errors.append({
                    "loja": store_rule["display"],
                    "produto": row["DescricaoOriginal"],
                    "erro": "Produto não encontrado na base e sem conversão automática",
                })
                continue
            produto_modelo = cleaned_product_name(row["DescricaoOriginal"])
            produto_key = norm_key(produto_modelo)
            categoria = infer_category_from_desc(row["DescricaoOriginal"], base_df)

        out_rows.append({
            "Grupo": store_rule["group"],
            "LojaKey": store_rule["col_key"],
            "LojaNome": store_rule["display"],
            "Categoria": categoria,
            "ProdutoModelo": produto_modelo,
            "ProdutoKey": produto_key,
            "Caixas": caixas,
            "CodigoPedido": row["CodigoPedido"],
            "CodFornPedido": row["CodFornPedido"],
            "PrecoPedido": row["PrecoPedido"],
            "ObservacaoConversao": conversion_note,
            "DescricaoOriginal": row["DescricaoOriginal"],
            "QtdeOriginal": row["Qtde"],
            "QtdeEmb": row["QtdeEmb"],
            "UnidadeDetectada": row["UnidadeDetectada"],
            "AutoAdicionado": "SIM" if base_row is None else "NAO",
        })

    out_df = pd.DataFrame(out_rows)
    errors_df = pd.DataFrame(errors)
    return out_df, errors_df
def group_to_matrix(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    p = df.pivot_table(index="ProdutoModelo", columns="LojaKey", values="Caixas", aggfunc="sum", fill_value=0)
    p.columns = [str(c) for c in p.columns]
    return p


def product_rows(ws):
    rows = {}
    for row in range(3, ws.max_row + 1):
        prod = norm_text(ws.cell(row, 1).value)
        if prod and norm_key(prod) not in IGNORE_NAMES:
            rows[norm_key(prod)] = row
    return rows


def copy_row_style(ws, src_row, dst_row):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def extract_store_number_from_model(v1, v2):
    c1 = norm_key(v1)
    c2 = norm_key(v2)
    combo = f"{c1} {c2}"
    if "LOJA 1" in combo or re.fullmatch(r"1", c2):
        return "1"
    if "LOJA 2" in combo or re.fullmatch(r"2", c2):
        return "2"
    if "LOJA 3" in combo or re.fullmatch(r"3", c2):
        return "3"
    if "LOJA 4" in combo or re.fullmatch(r"4", c2):
        return "4"
    return None


def model_map(ws):
    store_to_col = {}
    total_col = None
    for col in range(2, ws.max_column + 1):
        line1 = ws.cell(1, col).value
        line2 = ws.cell(2, col).value
        top = norm_key(line1)
        second = norm_key(line2)

        if "TOTAL" in top or "TOTAL" in second:
            total_col = col
            continue

        key = extract_store_number_from_model(line1, line2)
        if key:
            store_to_col[key] = col

    return store_to_col, total_col


def model_map_kross(ws):
    store_to_col = {}
    total_col = None
    for col in range(2, ws.max_column + 1):
        line1 = norm_key(ws.cell(1, col).value)
        line2 = norm_key(ws.cell(2, col).value)
        combo = f"{line1} {line2}"

        if "TOTAL" in combo:
            total_col = col
            continue
        if "ATACADISTA" in combo or re.fullmatch(r"1", line2):
            store_to_col["1"] = col
        elif "XAXIM" in combo or re.fullmatch(r"2", line2):
            store_to_col["2"] = col

    return store_to_col, total_col


def write_output(model_path: Path, data: pd.DataFrame, model_type: str) -> bytes:
    wb = load_workbook(str(model_path))
    ws = wb.active

    if model_type == "KROSS":
        stores, total_col = model_map_kross(ws)
    else:
        stores, total_col = model_map(ws)

    prod_map = product_rows(ws)
    cols_to_clear = list(stores.values())
    if total_col:
        cols_to_clear.append(total_col)

    for row in range(3, ws.max_row + 1):
        if norm_text(ws.cell(row, 1).value):
            for col in cols_to_clear:
                ws.cell(row, col).value = None

    used = set()

    for prod in data.index.tolist():
        key = norm_key(prod)
        if key not in prod_map:
            continue

        row = prod_map[key]
        used.add(key)
        row_total = 0

        for loja in data.columns:
            if loja in stores:
                val = float(data.loc[prod, loja])
                if val:
                    ws.cell(row, stores[loja]).value = val
                    row_total += val

        if total_col:
            ws.cell(row, total_col).value = row_total if row_total else None

    missing = [prod for prod in data.index.tolist() if norm_key(prod) not in used]
    if missing:
        last_filled = max(prod_map.values()) if prod_map else 3
        style_row = last_filled
        current_row = last_filled + 1

        for prod in missing:
            copy_row_style(ws, style_row, current_row)
            ws.cell(current_row, 1).value = prod

            row_total = 0
            for loja in data.columns:
                if loja in stores:
                    val = float(data.loc[prod, loja])
                    if val:
                        ws.cell(current_row, stores[loja]).value = val
                        row_total += val

            if total_col:
                ws.cell(current_row, total_col).value = row_total if row_total else None
            current_row += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def build_prices_sheet(df: pd.DataFrame) -> bytes:
    if df.empty:
        out = BytesIO()
        pd.DataFrame(columns=["CODIGO", "COD_FORN", "PRODUTO", "PRECO"]).to_excel(out, index=False)
        out.seek(0)
        return out.getvalue()

    base = df[["ProdutoModelo", "CodigoPedido", "CodFornPedido", "PrecoPedido"]].drop_duplicates(subset=["ProdutoModelo"])
    base = base.sort_values("ProdutoModelo")

    rows = []
    for _, r in base.iterrows():
        codigo = norm_text(r["CodigoPedido"])
        cod_num = "".join(filter(str.isdigit, codigo)) if codigo else ""
        rows.append({
            "CODIGO": f"'{cod_num}" if cod_num else "",
            "COD_FORN": norm_text(r["CodFornPedido"]),
            "PRODUTO": norm_text(r["ProdutoModelo"]),
            "PRECO": r["PrecoPedido"],
        })

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="PRECOS", index=False)
        ws = writer.sheets["PRECOS"]
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 45
        ws.column_dimensions["D"].width = 12
    out.seek(0)
    return out.getvalue()


def build_auto_added_sheet(df: pd.DataFrame) -> bytes:
    cols = [
        "Grupo", "LojaNome", "Categoria", "ProdutoModelo", "DescricaoOriginal", "UnidadeDetectada",
        "QtdeOriginal", "QtdeEmb", "Caixas", "CodigoPedido", "CodFornPedido", "PrecoPedido",
        "ObservacaoConversao", "AutoAdicionado"
    ]
    if df.empty or "AutoAdicionado" not in df.columns:
        base = pd.DataFrame(columns=cols)
    else:
        base = df[df["AutoAdicionado"] == "SIM"].copy()
        for c in cols:
            if c not in base.columns:
                base[c] = ""
        base = base[cols].sort_values(["Grupo", "LojaNome", "ProdutoModelo"])

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        base.to_excel(writer, sheet_name="AUTO_ADICIONADOS", index=False)
        ws = writer.sheets["AUTO_ADICIONADOS"]
        for letter, width in {"A":14,"B":22,"C":12,"D":30,"E":42,"F":14,"G":12,"H":12,"I":10,"J":12,"K":12,"L":12,"M":32,"N":14}.items():
            ws.column_dimensions[letter].width = width
    out.seek(0)
    return out.getvalue()


def build_cd_workbook(frutas_matrix: pd.DataFrame, legumes_matrix: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (frutas_matrix if not frutas_matrix.empty else pd.DataFrame()).to_excel(writer, sheet_name="FRUTAS")
        (legumes_matrix if not legumes_matrix.empty else pd.DataFrame()).to_excel(writer, sheet_name="LEGUMES")
    out.seek(0)
    return out.getvalue()


def ensure_fixed_files():
    resolved = {}
    missing = []

    for key, candidates in FIXED_FILE_CANDIDATES.items():
        path = resolve_first_existing(candidates)
        if path is None:
            missing.append(candidates[0])
        else:
            resolved[key] = path

    if missing:
        raise FileNotFoundError("Arquivos fixos não encontrados no repositório: " + ", ".join(missing))

    return resolved


def build_zip(files_dict: dict) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files_dict.items():
            zf.writestr(name, content)
    out.seek(0)
    return out.getvalue()


st.title("📦 BRASÃO / KROSS → THOTH (PRO)")
st.caption("Suba apenas os PDFs. A base e os modelos já ficam fixos dentro do sistema.")

with st.sidebar:
    st.subheader("Como usar")
    st.write("1. Envie todos os PDFs do dia.")
    st.write("2. Clique em PROCESSAR.")
    st.write("3. Baixe o ZIP com os arquivos finais.")
    st.info("Fluxos suportados: Brasão lojas, Brasão CD e Kross.")
    st.info("Brasão Fernando = Loja 1 | Jardim = Loja 2 | Xaxim = Loja 3 | Avenida = Loja 4.")
    st.info("Kross Atacadista = Loja 1 | Kross Xaxim = Loja 2.")
    st.info("Base fixa: base_thoth_app_normalizada_brasao_kross.xlsx")

pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR", use_container_width=True, type="primary"):
    if not pdf_files:
        st.error("Envie ao menos um PDF.")
    else:
        try:
            fixed_files = ensure_fixed_files()
            base_df = load_base_from_disk(str(fixed_files["BASE_PRODUCTS_FILE"]))

            transformed_parts = []
            all_errors = []
            identified = []
            seen_store_ids = set()

            for pdf in pdf_files:
                text = pdf_to_text(pdf.getvalue())
                store_rule = identify_store(text, pdf.name)

                if store_rule["store_id"] in seen_store_ids:
                    all_errors.append(pd.DataFrame([{
                        "loja": store_rule["display"],
                        "produto": pdf.name,
                        "erro": "PDF duplicado da mesma unidade ignorado",
                    }]))
                    continue

                seen_store_ids.add(store_rule["store_id"])
                order_df = parse_order_items(text)
                transformed_df, errors_df = transform_items(order_df, store_rule, base_df)

                if not transformed_df.empty:
                    transformed_parts.append(transformed_df)
                if not errors_df.empty:
                    all_errors.append(errors_df)

                identified.append({
                    "arquivo": pdf.name,
                    "loja": store_rule["display"],
                    "grupo": store_rule["group"],
                    "itens_extraidos": len(order_df),
                    "itens_convertidos": len(transformed_df),
                })

            if not transformed_parts:
                raise ValueError("Nenhum item foi convertido. Verifique a base e os PDFs.")

            all_data = pd.concat(transformed_parts, ignore_index=True)
            errors_data = pd.concat(all_errors, ignore_index=True) if all_errors else pd.DataFrame(columns=["loja", "produto", "erro"])
            identified_df = pd.DataFrame(identified)

            brasao_df = all_data[all_data["Grupo"] == "BRASAO"].copy()
            kross_df = all_data[all_data["Grupo"] == "KROSS"].copy()
            cd_df = all_data[all_data["Grupo"] == "BRASAO_CD"].copy()

            brasao_frutas = brasao_df[brasao_df["Categoria"] == "FRUTAS"]
            brasao_legumes = brasao_df[brasao_df["Categoria"] == "LEGUMES"]
            kross_frutas = kross_df[kross_df["Categoria"] == "FRUTAS"]
            kross_legumes = kross_df[kross_df["Categoria"] == "LEGUMES"]
            cd_frutas = cd_df[cd_df["Categoria"] == "FRUTAS"]
            cd_legumes = cd_df[cd_df["Categoria"] == "LEGUMES"]

            brasao_frutas_matrix = group_to_matrix(brasao_frutas)
            brasao_legumes_matrix = group_to_matrix(brasao_legumes)
            kross_frutas_matrix = group_to_matrix(kross_frutas)
            kross_legumes_matrix = group_to_matrix(kross_legumes)
            cd_frutas_matrix = group_to_matrix(cd_frutas)
            cd_legumes_matrix = group_to_matrix(cd_legumes)

            files_to_zip = {
                "BRASAO_FRUTAS_Thoth.xlsx": write_output(fixed_files["MODEL_BRASAO_FRUTAS"], brasao_frutas_matrix, "BRASAO"),
                "BRASAO_LEGUMES_Thoth.xlsx": write_output(fixed_files["MODEL_BRASAO_LEGUMES"], brasao_legumes_matrix, "BRASAO"),
                "KROSS_FRUTAS_Thoth.xlsx": write_output(fixed_files["MODEL_KROSS_FRUTAS"], kross_frutas_matrix, "KROSS"),
                "KROSS_LEGUMES_Thoth.xlsx": write_output(fixed_files["MODEL_KROSS_LEGUMES"], kross_legumes_matrix, "KROSS"),
                "BRASAO_CD.xlsx": build_cd_workbook(cd_frutas_matrix, cd_legumes_matrix),
                "BRASAO_PRECOS.xlsx": build_prices_sheet(brasao_df),
                "KROSS_PRECOS.xlsx": build_prices_sheet(kross_df),
                "BRASAO_CD_PRECOS.xlsx": build_prices_sheet(cd_df),
            }

            expected_ids = {
                "BRASAO_FERNANDO", "BRASAO_JARDIM", "BRASAO_XAXIM", "BRASAO_AVENIDA",
                "BRASAO_CD", "KROSS_CHAPECO", "KROSS_XAXIM"
            }
            missing_units = []
            for rule in STORE_RULES:
                if rule["store_id"] in expected_ids and rule["store_id"] not in seen_store_ids:
                    missing_units.append({"loja": rule["display"], "produto": "", "erro": "PDF da unidade não enviado"})

            if missing_units:
                missing_df = pd.DataFrame(missing_units)
                errors_data = pd.concat([errors_data, missing_df], ignore_index=True)

            err_out = BytesIO()
            with pd.ExcelWriter(err_out, engine="openpyxl") as writer:
                identified_df.to_excel(writer, sheet_name="ARQUIVOS", index=False)
                errors_data.to_excel(writer, sheet_name="ERROS", index=False)
            err_out.seek(0)
            files_to_zip["LOG_PROCESSAMENTO.xlsx"] = err_out.getvalue()

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

            st.download_button(
                "Baixar ZIP final",
                zip_bytes,
                file_name="THOTH_BRASAO_KROSS_PRO.zip",
                mime="application/zip",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
            with st.expander("Ver detalhes do erro"):
                st.code(traceback.format_exc())
