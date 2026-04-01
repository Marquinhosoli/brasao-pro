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

MODEL_CANDIDATES = {
    "BRASAO_FRUTAS": [
        "BRASAO - FRUTAS PRE PEDIDO BRANCO.xlsx",
        "BRASAO FRUTAS BRANCO.xlsx",
    ],
    "BRASAO_LEGUMES": [
        "BRASAO - LEGUMES PRE PEDIDO BRANCO.xlsx",
        "BRASAO LEGUMES BRANCO.xlsx",
    ],
    "KROSS_FRUTAS": [
        "KROSS - FRUTAS PRE PEDIDO BRANCO.xlsx",
        "KROSS - PRE PEDIDO FRUTAS BRANCO.xlsx",
    ],
    "KROSS_LEGUMES": [
        "KROSS - LEGUMES PRE PEDIDO BRANCO.xlsx",
    ],
}
BASE_PRODUCTS_FILE = BASE_DIR / "base_thoth_app_normalizada_brasao_kross.xlsx"

STORE_RULES = [
    {
        "store_id": "BRASAO_FERNANDO",
        "group": "BRASAO",
        "col_key": "1",
        "display": "Brasão Fernando",
        "signals": ["FERNANDO MACHADO", "CENTRO", "226"],
    },
    {
        "store_id": "BRASAO_JARDIM",
        "group": "BRASAO",
        "col_key": "2",
        "display": "Brasão Jardim",
        "signals": ["SAO PEDRO", "JARDIM AMERICA", "2199"],
    },
    {
        "store_id": "BRASAO_XAXIM",
        "group": "BRASAO",
        "col_key": "3",
        "display": "Brasão Xaxim",
        "signals": ["LUIZ LUNARDI", "XAXIM", "810"],
    },
    {
        "store_id": "BRASAO_AVENIDA",
        "group": "BRASAO",
        "col_key": "4",
        "display": "Brasão Avenida",
        "signals": ["RIO DE JANEIRO", "CENTRO", "108", "CHAPECO"],
    },
    {
        "store_id": "BRASAO_CD",
        "group": "BRASAO_CD",
        "col_key": "CD",
        "display": "Brasão CD",
        "signals": ["RUA GASPAR", "ELDORADO", "153"],
    },
    {
        "store_id": "KROSS_CHAPECO",
        "group": "KROSS",
        "col_key": "1",
        "display": "Kross Atacadista",
        "signals": ["JOHN KENNEDY", "PASSO DOS FORTES", "550"],
    },
    {
        "store_id": "KROSS_XAXIM",
        "group": "KROSS",
        "col_key": "2",
        "display": "Kross Xaxim",
        "signals": ["AMELIO PANIZZI", "XAXIM"],
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

RESOLVED_MODEL_PATHS = {}


def resolve_model_path(model_key: str) -> Path:
    if model_key in RESOLVED_MODEL_PATHS:
        return RESOLVED_MODEL_PATHS[model_key]

    for file_name in MODEL_CANDIDATES[model_key]:
        path = BASE_DIR / file_name
        if path.exists():
            RESOLVED_MODEL_PATHS[model_key] = path
            return path

    raise FileNotFoundError(
        f"Modelo não encontrado para {model_key}. Esperado um destes arquivos: {', '.join(MODEL_CANDIDATES[model_key])}"
    )


def get_model_paths():
    return {
        "BRASAO_FRUTAS": resolve_model_path("BRASAO_FRUTAS"),
        "BRASAO_LEGUMES": resolve_model_path("BRASAO_LEGUMES"),
        "KROSS_FRUTAS": resolve_model_path("KROSS_FRUTAS"),
        "KROSS_LEGUMES": resolve_model_path("KROSS_LEGUMES"),
    }


def tokens_from_text(v: str):
    key = norm_key(v)
    return [t for t in re.split(r"[^A-Z0-9]+", key) if len(t) >= 2]


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
    return v is None or pd.isna(v)


def safe_number(v):
    if is_missing_number(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def ceil_div(qtd, base):
    qtd = safe_number(qtd)
    base = safe_number(base)
    if qtd is None or base is None or base <= 0:
        return 0
    return int(-(-qtd // base))


def choose_real_quantity(qtd, qtd_emb, unit="", modo=""):
    qtd = safe_number(qtd) or 0
    qtd_emb = safe_number(qtd_emb) or 0
    unit = norm_key(unit)
    modo = norm_key(modo)

    if qtd_emb > qtd:
        chosen = qtd_emb
    else:
        chosen = qtd

    # Quando o PDF traz 1 no campo Qtde e a carga real no campo embalagem/peso
    if chosen <= 1 and qtd_emb > 1:
        chosen = qtd_emb

    # Em pedidos de KG/UND/BDJ, normalmente o maior campo é o valor real a converter
    if (unit in {"KG", "UND", "BDJ"} or modo in {"PESO", "UNIDADE", "UND", "BANDEJA", "BDJ"}) and qtd_emb > qtd:
        chosen = qtd_emb

    return chosen


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

    # prioridade por nome do arquivo
    if " CD" in f" {name}" or name.endswith("CD") or "_CD" in name or " BRASAO CD" in name:
        for rule in STORE_RULES:
            if rule["store_id"] == "BRASAO_CD":
                return rule

    name_hints = [
        ("FERNANDO", "BRASAO_FERNANDO"),
        ("JARDIM", "BRASAO_JARDIM"),
        ("AVENIDA", "BRASAO_AVENIDA"),
        ("XAXIM", "BRASAO_XAXIM"),
        ("ATACADO", "KROSS_CHAPECO"),
        ("CHAPECO", "KROSS_CHAPECO"),
        ("KROSS XAXIM", "KROSS_XAXIM"),
    ]
    for hint, store_id in name_hints:
        if hint in name:
            for rule in STORE_RULES:
                if rule["store_id"] == store_id:
                    return rule

    # depois tenta pelo conteúdo
    for rule in STORE_RULES:
        if all(signal in text for signal in rule["signals"]):
            return rule

    # fallback por sinais parciais
    scored = []
    for rule in STORE_RULES:
        score = sum(1 for signal in rule["signals"] if signal in text)
        if score:
            scored.append((score, rule))
    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        if scored[0][0] >= 2:
            return scored[0][1]

    raise ValueError(f"Não consegui identificar a loja/unidade pelo PDF: {file_name}")


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


def match_base_item(desc: str, base_df: pd.DataFrame):
    key = norm_key(desc)
    desc_tokens = set(tokens_from_text(desc))
    if not key:
        return None

    exact = base_df[base_df["produto_key"] == key]
    if not exact.empty:
        return exact.iloc[0]

    # sinônimos exatos/parciais
    best_row = None
    best_score = 0

    for _, row in base_df.iterrows():
        row_key = row["produto_key"]

        # match por sinônimos
        for syn in row["sinonimos_key"]:
            if syn == key:
                return row
            if syn and (syn in key or key in syn):
                score = 90 + min(len(syn), len(key)) / 100
                if score > best_score:
                    best_score = score
                    best_row = row

        # contains direto
        if row_key == key:
            return row
        if row_key and (row_key in key or key in row_key):
            score = 80 + min(len(row_key), len(key)) / 100
            if score > best_score:
                best_score = score
                best_row = row

        # score por tokens
        row_tokens = set(tokens_from_text(row_key))
        common = desc_tokens & row_tokens
        if common:
            score = len(common) * 10
            if row.get("categoria") and norm_key(row.get("categoria")) in {"FRUTAS", "LEGUMES"}:
                score += 1
            if score > best_score and len(common) >= max(1, min(2, len(row_tokens))):
                best_score = score
                best_row = row

    return best_row


def convert_to_boxes(qtd: float, unit: str, base_row, qtd_emb: float = None) -> int:
    modo = norm_key(base_row.get("modo", ""))
    unit = norm_key(unit)

    qtd_real = choose_real_quantity(qtd, qtd_emb, unit=unit, modo=modo)

    if modo == "CAIXA":
        return int(qtd_real)

    if modo == "PESO":
        peso = safe_number(base_row.get("peso_caixa"))
        if peso and peso > 0:
            return ceil_div(qtd_real, peso)

    if modo in {"UNIDADE", "UND"}:
        itens = safe_number(base_row.get("itens_por_caixa"))
        if itens and itens > 0:
            return ceil_div(qtd_real, itens)

    if modo in {"BANDEJA", "BDJ"}:
        bdj = safe_number(base_row.get("bandejas_por_caixa"))
        if bdj and bdj > 0:
            return ceil_div(qtd_real, bdj)

    # fallback baseado na unidade detectada, caso o modo esteja vazio
    if unit == "KG":
        peso = safe_number(base_row.get("peso_caixa"))
        if peso and peso > 0:
            return ceil_div(qtd_real, peso)

    if unit == "UND":
        itens = safe_number(base_row.get("itens_por_caixa"))
        if itens and itens > 0:
            return ceil_div(qtd_real, itens)

    if unit == "BDJ":
        bdj = safe_number(base_row.get("bandejas_por_caixa"))
        if bdj and bdj > 0:
            return ceil_div(qtd_real, bdj)

    if unit == "CX":
        return int(qtd_real)

    return 0


def transform_items(order_df: pd.DataFrame, store_rule: dict, base_df: pd.DataFrame):
    out_rows = []
    errors = []

    for _, row in order_df.iterrows():
        base_row = match_base_item(row["DescricaoOriginal"], base_df)
        if base_row is None:
            errors.append({
                "loja": store_rule["display"],
                "produto": row["DescricaoOriginal"],
                "erro": "Produto não encontrado na base",
            })
            continue

        caixas = convert_to_boxes(row["Qtde"], row["UnidadeDetectada"], base_row, row.get("QtdeEmb"))
        if caixas <= 0:
            errors.append({
                "loja": store_rule["display"],
                "produto": row["DescricaoOriginal"],
                "erro": f"Falha na conversão para caixa ({row['UnidadeDetectada']}) | qtd={row.get('Qtde')} emb={row.get('QtdeEmb')}",
            })
            continue

        out_rows.append({
            "Grupo": store_rule["group"],
            "LojaKey": store_rule["col_key"],
            "LojaNome": store_rule["display"],
            "Categoria": norm_key(base_row["categoria"]),
            "ProdutoModelo": base_row["produto_base"],
            "ProdutoKey": base_row["produto_key"],
            "Caixas": caixas,
            "CodigoPedido": row["CodigoPedido"],
            "CodFornPedido": row["CodFornPedido"],
            "PrecoPedido": row["PrecoPedido"],
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


def build_cd_workbook(frutas_matrix: pd.DataFrame, legumes_matrix: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (frutas_matrix if not frutas_matrix.empty else pd.DataFrame()).to_excel(writer, sheet_name="FRUTAS")
        (legumes_matrix if not legumes_matrix.empty else pd.DataFrame()).to_excel(writer, sheet_name="LEGUMES")
    out.seek(0)
    return out.getvalue()


def ensure_fixed_files():
    missing = []
    if not BASE_PRODUCTS_FILE.exists():
        missing.append(BASE_PRODUCTS_FILE.name)

    for model_key, candidates in MODEL_CANDIDATES.items():
        found = False
        for file_name in candidates:
            if (BASE_DIR / file_name).exists():
                found = True
                break
        if not found:
            missing.append(" ou ".join(candidates))

    if missing:
        raise FileNotFoundError("Arquivos fixos não encontrados no repositório: " + ", ".join(missing))


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
    st.info(f"Base fixa: {BASE_PRODUCTS_FILE.name}")

pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR", use_container_width=True, type="primary"):
    if not pdf_files:
        st.error("Envie ao menos um PDF.")
    else:
        try:
            ensure_fixed_files()
            base_df = load_base_from_disk(str(BASE_PRODUCTS_FILE))

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

            model_paths = get_model_paths()

            files_to_zip = {
                "BRASAO_FRUTAS_Thoth.xlsx": write_output(model_paths["BRASAO_FRUTAS"], brasao_frutas_matrix, "BRASAO"),
                "BRASAO_LEGUMES_Thoth.xlsx": write_output(model_paths["BRASAO_LEGUMES"], brasao_legumes_matrix, "BRASAO"),
                "KROSS_FRUTAS_Thoth.xlsx": write_output(model_paths["KROSS_FRUTAS"], kross_frutas_matrix, "KROSS"),
                "KROSS_LEGUMES_Thoth.xlsx": write_output(model_paths["KROSS_LEGUMES"], kross_legumes_matrix, "KROSS"),
                "BRASAO_CD.xlsx": build_cd_workbook(cd_frutas_matrix, cd_legumes_matrix),
                "BRASAO_PRECOS.xlsx": build_prices_sheet(brasao_df),
                "KROSS_PRECOS.xlsx": build_prices_sheet(kross_df),
                "BRASAO_CD_PRECOS.xlsx": build_prices_sheet(cd_df),
            }

            missing_units = []
            expected_ids = {
                "BRASAO_FERNANDO", "BRASAO_JARDIM", "BRASAO_XAXIM", "BRASAO_AVENIDA",
                "BRASAO_CD", "KROSS_CHAPECO", "KROSS_XAXIM"
            }
            for rule in STORE_RULES:
                if rule["store_id"] in expected_ids and rule["store_id"] not in seen_store_ids:
                    missing_units.append({"loja": rule["display"], "produto": "", "erro": "PDF da unidade não enviado"})

            if missing_units:
                missing_df = pd.DataFrame(missing_units)
                errors_data = pd.concat([errors_data, missing_df], ignore_index=True)

            auditoria_df = all_data.copy()
            if not auditoria_df.empty:
                auditoria_df["QtdRealUsada"] = auditoria_df.apply(lambda r: choose_real_quantity(r.get("CodigoPedido") if False else None, None), axis=1) if False else None

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
