import io
import math
import re
import unicodedata
import zipfile
from io import BytesIO
from difflib import SequenceMatcher
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="BRASÃO / KROSS → THOTH (Modo Inteligente)",
    page_icon="📦",
    layout="wide",
)

# =========================================================
# CONFIG
# =========================================================

STORE_RULES = [
    {"store_id": "BRASAO_FERNANDO", "group": "BRASAO", "col_key": "1", "display": "Brasão Fernando", "filename_signals": ["FERNANDO", "CE"], "text_signals": ["FERNANDO MACHADO", "CENTRO, 226"]},
    {"store_id": "BRASAO_JARDIM", "group": "BRASAO", "col_key": "2", "display": "Brasão Jardim", "filename_signals": ["JARDIM", "JA"], "text_signals": ["JARDIM AMERICA", "SAO PEDRO", "2199"]},
    {"store_id": "BRASAO_XAXIM", "group": "BRASAO", "col_key": "3", "display": "Brasão Xaxim", "filename_signals": ["XAXIM", "XX"], "text_signals": ["LUIZ LUNARDI", "XAXIM", "810"]},
    {"store_id": "BRASAO_AVENIDA", "group": "BRASAO", "col_key": "4", "display": "Brasão Avenida", "filename_signals": ["AVENIDA", "AV"], "text_signals": ["RIO DE JANEIRO", "CENTRO, 108"]},
    {"store_id": "BRASAO_CD", "group": "BRASAO_CD", "col_key": "CD", "display": "Brasão CD", "filename_signals": ["CD"], "text_signals": ["RUA GASPAR", "ELDORADO", "153"]},
    {"store_id": "KROSS_CHAPECO", "group": "KROSS", "col_key": "1", "display": "Kross Atacadista", "filename_signals": ["KROSS", "CHAPECO", "ATACADO"], "text_signals": ["JOHN KENNEDY", "PASSO DOS FORTES", "550"]},
    {"store_id": "KROSS_XAXIM", "group": "KROSS", "col_key": "2", "display": "Kross Xaxim", "filename_signals": ["KROSS", "XAXIM", "XX"], "text_signals": ["AMELIO PANIZZI", "XAXIM"]},
]

NOISE_PATTERNS = [
    r"\bDE MARCHI\b",
    r"\bDEMARCHI\b",
    r"\bBRASAO\b",
    r"\bFRUTA\b",
    r"\bFRUTA\b",
    r"\bFRU\b",
    r"\bBANDEJA\b",
    r"\bBDJ\b",
    r"\bSHELF\b",
    r"\bFRUTAMINA\b",
    r"\bNACIONAL\b",
    r"\bIMPORTADO\b",
    r"\bARGENTINA\b",
    r"\bGRECIA\b",
    r"\bPOLPA AMARELA\b",
    r"\bINTEIRA\b",
    r"\bESPIGA\b",
    r"\bUNIDADE\b",
    r"\bUND\b",
    r"\bUN\b",
    r"\bMACO\b",
    r"\bKG\b",
    r"\bCAT\b",
    r"\bCX\b",
    r"\bCAIXA\b",
    r"\bMARCHI\b",
]

STOP_LINES = [
    "NUMERO DO PEDIDO",
    "TRANSACAO",
    "EMPRESA:",
    "CEP:",
    "TROCAS:",
    "JOAB FATURAMENTO",
    "KELLY",
    "LEANDRO GERENTE",
    "E-MAIL:",
    "PG:",
    "PEDIDO DE COMPRA",
    "DCTO:",
    "USUARIO:",
    "TOTAL DO FORNECEDOR",
    "CONTATOS DO FORNECEDOR",
    "PAGINA",
]

UNIT_HINTS = [
    ("BDJ", "BDJ"),
    ("BANDEJA", "BDJ"),
    ("MACO", "UND"),
    ("UNIDADE", "UND"),
    ("UND", "UND"),
    (" UN ", "UND"),
    ("KG", "KG"),
]

MANUAL_MAP = {
    "ABACAXI PEROLA": "ABACAXI",
    "ABACAXI": "ABACAXI",
    "AMEIXA NACIONAL": "AMEIXA",
    "AMEIXA": "AMEIXA",
    "BATATA DOCE BRANCA": "BATATA DOCE BRANCA",
    "BATATA DOCE ROXA": "BATATA DOCE ROXA",
    "BATATA SALSA": "BATATA SALSA",
    "BERINJELA": "BERINJELA",
    "BETERRABA": "BETERRABA",
    "CAQUI RAMA FORTE": "CAQUI RAMA FORTE",
    "CEBOLA CONSERVA": "CEBOLA CONSERVA",
    "CEBOLA ARGENTINA": "CEBOLA ALBINA",
    "CENOURA": "CENOURA",
    "CHUCHU": "CHUCHU",
    "COCO SECO": "COCO SECO",
    "FIGO ROXO": "FIGO ROXO",
    "FRAMBOESA": "FRAMBOESA",
    "GOIABA NACIONAL VERMELHA": "GOIABA",
    "GOIABA": "GOIABA",
    "JATOBA": "JATOBA",
    "KIWI IMPORTADO GRECIA": "KIWI IMPORTADO",
    "KIWI IMPORTADO": "KIWI IMPORTADO",
    "KIWI NACIONAL": "KIWI NACIONAL 600G",
    "LARANJA MAQUINA DE SUCO": "LARANJA MAQUINA DE SUCO",
    "LIMAO SICILIANO": "LIMAO SICILIANO",
    "LIMAO TAHITI": "LIMAO TAHITI",
    "MACA FUJI": "MACA FUJI",
    "MAMAO FORMOSA": "MAMAO FORMOSA",
    "MAMAOZINHO PAPAIA": "MAMAO PAPAYA",
    "MAMAO PAPAYA": "MAMAO PAPAYA",
    "MANGA PALMER": "MANGA PALMER",
    "MAXIXE": "MAXIXE",
    "MELAO CANTALOUPE": "MELAO CANTALOUPE",
    "MELAO DINO": "MELAO DINO",
    "MELAO ESPANHOL AMARELO": "MELAO ESPANHOL AMARELO",
    "MELAO GALIA": "MELAO GALIA",
    "MELAO ORANGE": "MELAO ORANGE",
    "MELAO REI DOCE REDINHA": "MELAO REI DOCE REDINHA",
    "MELAO SAPO": "MELAO SAPO",
    "MELANCIA INTEIRA": "MELANCIA",
    "MILHO VERDE": "MILHO VERDE",
    "NABO": "NABO",
    "PEPINO JAPONES": "PEPINO JAPONES",
    "PERA WILLIANS": "PERA WILLIANS",
    "PESSEGO IMP": "PESSEGO IMPORTADO",
    "PESSEGO IMPORTADO": "PESSEGO IMPORTADO",
    "PIMENTA BIQUINHO": "PIMENTA BIQUINHO",
    "PIMENTA JALAPENO": "PIMENTA JALAPENO",
    "SALSAO AIPO": "SALSAO AIPO",
    "TOMATE GRAPE": "TOMATE GRAP",
    "BLUEBERRY": "BLUEBERRY 125G",
    "MIRTILO BLUEBERRY": "BLUEBERRY 125G",
    "UVA THOMPSON": "UVA THOMPSON 500G",
    "PHYSALIS": "PHYSALIS IMPORTADO 100G",
    "CARAMBOLA": "CARAMBOLA 400G",
}

DEMO_BASE = [
    {"produto": "ABACAXI", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 10, "sinonimos": "ABACAXI PEROLA"},
    {"produto": "ABACATE", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 18},
    {"produto": "ALECRIM", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "ALHO PORO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "AMEIXA", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 20, "sinonimos": "AMEIXA NACIONAL"},
    {"produto": "BATATA DOCE BRANCA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "BATATA DOCE ROXA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "BATATA SALSA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "BERINJELA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "BETERRABA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "BLUEBERRY 125G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 12, "sinonimos": "MIRTILO BLUEBERRY;BLUEBERRY"},
    {"produto": "CAQUI RAMA FORTE", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "CARAMBOLA 400G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 4, "sinonimos": "CARAMBOLA"},
    {"produto": "CEBOLA ALBINA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20, "sinonimos": "CEBOLA ARGENTINA"},
    {"produto": "CEBOLA CONSERVA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "CENOURA", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "CHUCHU", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "COCO SECO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20},
    {"produto": "COENTRO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "FIGO ROXO", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 12},
    {"produto": "FRAMBOESA", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 10},
    {"produto": "GOIABA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "HORTELA", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "JATOBA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "KIWI IMPORTADO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10, "sinonimos": "KIWI IMPORTADO GRECIA"},
    {"produto": "KIWI NACIONAL 600G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 15, "sinonimos": "KIWI NACIONAL"},
    {"produto": "LARANJA MAQUINA DE SUCO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20},
    {"produto": "LIMAO SICILIANO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20},
    {"produto": "LIMAO TAHITI", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20, "sinonimos": "LIMAO"},
    {"produto": "LOURO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "MACA FUJI", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 18},
    {"produto": "MAMAO FORMOSA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 18},
    {"produto": "MAMAO PAPAYA", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 18, "sinonimos": "MAMAOZINHO PAPAIA;PAPAYA"},
    {"produto": "MANGA PALMER", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 12},
    {"produto": "MANJERICAO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "MANJERONA", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "MAXIXE", "categoria": "LEGUMES", "tipo": "BDJ", "bandejas_por_caixa": 12},
    {"produto": "MELANCIA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 20},
    {"produto": "MELAO CANTALOUPE", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 6},
    {"produto": "MELAO DINO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "MELAO ESPANHOL AMARELO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "MELAO GALIA", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 6},
    {"produto": "MELAO ORANGE", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 6},
    {"produto": "MELAO REI DOCE REDINHA", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "MELAO SAPO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "MILHO VERDE", "categoria": "LEGUMES", "tipo": "BDJ", "bandejas_por_caixa": 10},
    {"produto": "NABO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 10},
    {"produto": "PEPINO JAPONES", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 20},
    {"produto": "PERA WILLIANS", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 12},
    {"produto": "PESSEGO IMPORTADO", "categoria": "FRUTAS", "tipo": "KG", "peso_caixa": 10},
    {"produto": "PIMENTA BIQUINHO", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 6},
    {"produto": "PIMENTA JALAPENO", "categoria": "LEGUMES", "tipo": "KG", "peso_caixa": 6},
    {"produto": "SALSAO AIPO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "SALVIA", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "TOMATE GRAP", "categoria": "LEGUMES", "tipo": "BDJ", "bandejas_por_caixa": 24, "sinonimos": "TOMATE GRAPE"},
    {"produto": "TOMILHO", "categoria": "LEGUMES", "tipo": "UND", "unidades_caixa": 12},
    {"produto": "UVA THOMPSON 500G", "categoria": "FRUTAS", "tipo": "BDJ", "bandejas_por_caixa": 10, "sinonimos": "UVA THOMPSON"},
    {"produto": "PHYSALIS IMPORTADO 100G", "categoria": "FRUTAS", "tipo": "UND", "unidades_caixa": 8, "sinonimos": "PHYSALIS"},
]

# =========================================================
# HELPERS
# =========================================================

def clean_text(v) -> str:
    if v is None:
        return ""
    txt = str(v).strip()
    if txt.lower() in {"nan", "none"}:
        return ""
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", txt).strip().upper()

def parse_number(v):
    if v is None:
        return None
    s = clean_text(v)
    if not s:
        return None
    m = re.search(r"\d[\d\.,]*", s)
    if not m:
        return None
    num = m.group(0)
    if "," in num and "." in num:
        num = num.replace(".", "").replace(",", ".")
    elif "," in num:
        num = num.replace(",", ".")
    try:
        return float(num)
    except:
        return None

def infer_unit(text: str) -> str:
    s = f" {clean_text(text)} "
    for token, unit in UNIT_HINTS:
        if f" {token} " in s:
            return unit
    return ""

def strip_noise_product(s: str) -> str:
    t = f" {clean_text(s)} "
    t = re.sub(r"\b\d{5,}\b", " ", t)
    t = re.sub(r"\b\d+G\b", " ", t)
    t = re.sub(r"\b\d+,\d+\b", " ", t)
    t = re.sub(r"\b\d+\b", " ", t)
    for pat in NOISE_PATTERNS:
        t = re.sub(pat, " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def canonical_product(s: str) -> str:
    raw = strip_noise_product(s)
    if raw in MANUAL_MAP:
        return MANUAL_MAP[raw]
    # heurísticas
    if "TOMATE" in raw and "GRAPE" in raw:
        return "TOMATE GRAP"
    if "UVA" in raw and "THOMPSON" in raw:
        return "UVA THOMPSON 500G"
    if "MIRTILO" in raw or "BLUEBERRY" in raw:
        return "BLUEBERRY 125G"
    if "PHYSALIS" in raw:
        return "PHYSALIS IMPORTADO 100G"
    if "MAMAOZINHO" in raw or "PAPAIA" in raw:
        return "MAMAO PAPAYA"
    if "LARANJA" in raw and "SUCO" in raw:
        return "LARANJA MAQUINA DE SUCO"
    return raw

def score_similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()

def identify_store(filename: str, pdf_text: str) -> dict:
    fn = clean_text(filename)
    txt = clean_text(pdf_text[:4000])
    best, best_score = STORE_RULES[0], -1
    for rule in STORE_RULES:
        score = 0
        for sig in rule["filename_signals"]:
            if clean_text(sig) in fn:
                score += 5
        for sig in rule["text_signals"]:
            if clean_text(sig) in txt:
                score += 3
        if score > best_score:
            best_score = score
            best = rule
    return best

def build_base_df(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame(DEMO_BASE)
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

def build_base_map(df: pd.DataFrame) -> Dict[str, dict]:
    base_map = {}
    for _, row in df.iterrows():
        produto = row.get("produto", row.get("PRODUTO", row.get("produto_base", row.get("PRODUTO_BASE", ""))))
        produto = canonical_product(produto)
        if not produto:
            continue
        item = {
            "produto": produto,
            "categoria": clean_text(row.get("categoria", row.get("CATEGORIA", ""))),
            "tipo": clean_text(row.get("tipo", row.get("TIPO", row.get("modo_conversao", row.get("MODO_CONVERSAO", ""))))),
            "peso_caixa": float(row.get("peso_caixa", row.get("PESO_CAIXA", 0)) or 0),
            "bandejas_por_caixa": float(row.get("bandejas_por_caixa", row.get("BANDEJAS_POR_CAIXA", 0)) or 0),
            "unidades_caixa": float(row.get("unidades_caixa", row.get("UNIDADES_CAIXA", row.get("itens_por_caixa", row.get("ITENS_POR_CAIXA", 0)))) or 0),
        }
        base_map[produto] = item
        syn = str(row.get("sinonimos", row.get("SINONIMOS", "")) or "")
        for s in [x.strip() for x in syn.split(";") if x.strip()]:
            base_map[canonical_product(s)] = item
    return base_map

def best_base_match(produto: str, base_map: Dict[str, dict]):
    if produto in base_map:
        return base_map[produto], 1.0, produto
    best_key, best_item, best_score = None, None, 0.0
    for k, item in base_map.items():
        score = score_similarity(produto, k)
        if produto and k and (produto in k or k in produto):
            score = max(score, 0.92)
        if score > best_score:
            best_key, best_item, best_score = k, item, score
    if best_score >= 0.78:
        return best_item, best_score, best_key
    return None, best_score, best_key

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)

def is_stop_line(line: str) -> bool:
    s = clean_text(line)
    return any(token in s for token in STOP_LINES)

def parse_order_pdf(pdf_text: str) -> pd.DataFrame:
    rows = []
    for raw in pdf_text.splitlines():
        line = clean_text(raw)
        if not line or is_stop_line(line):
            continue

        # tenta capturar linhas do pedido no formato típico
        patterns = [
            r"^\d+\s+[\d,]+\s+(?P<produto>.+?)\s+(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN|MACO)\s+[\d,]+\s+[\d,]+\s+[\d,]+$",
            r"^\d+\s+(?P<produto>.+?)\s+(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN|MACO)\s+[\d,]+\s+[\d,]+\s+[\d,]+$",
            r"^(?P<produto>.+?)\s+(?P<unit>KG|BDJ|BANDEJA|UND|UNIDADE|UN|MACO)\s+(?P<qtd>\d[\d,\.]*)\s+[\d,\.]+$",
        ]

        item = None
        for p in patterns:
            m = re.search(p, line)
            if m:
                item = m
                break

        produto_txt = ""
        qtd = None
        unit = ""

        if item:
            produto_txt = item.group("produto")
            unit = clean_text(item.group("unit"))
            nums = re.findall(r"\d[\d,\.]*", line)
            if len(nums) >= 2:
                qtd = parse_number(nums[-2])  # penúltimo costuma ser quantidade
        else:
            # fallback: ignora linhas sem unidade reconhecível
            unit = infer_unit(line)
            if not unit:
                continue
            nums = re.findall(r"\d[\d,\.]*", line)
            if len(nums) < 2:
                continue
            qtd = parse_number(nums[-2])
            produto_txt = line

        if qtd is None or qtd <= 0:
            continue

        produto = canonical_product(produto_txt)
        if not produto:
            continue

        if unit == "BANDEJA":
            unit = "BDJ"
        if unit in {"UNIDADE", "UN"}:
            unit = "UND"

        rows.append({
            "produto_original": produto_txt,
            "produto_normalizado": produto,
            "qtd": qtd,
            "tipo_detectado": unit,
            "linha_original": line,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.drop_duplicates(subset=["linha_original"]).reset_index(drop=True)
    return df

def convert_item(produto: str, qtd: float, tipo_detectado: str, base_map: Dict[str, dict]):
    base_item, score, matched_key = best_base_match(produto, base_map)
    if not base_item:
        return None, f"SEM BASE ({produto})"

    tipo = base_item["tipo"] or tipo_detectado
    fator = 0
    if tipo == "KG":
        fator = base_item["peso_caixa"]
    elif tipo == "BDJ":
        fator = base_item["bandejas_por_caixa"]
    elif tipo in {"UND", "UN"}:
        tipo = "UND"
        fator = base_item["unidades_caixa"]

    if not fator or fator <= 0:
        return None, f"BASE INCOMPLETA ({base_item['produto']})"

    caixas = math.ceil(qtd / fator)
    return {
        "produto_base": base_item["produto"],
        "categoria": base_item["categoria"],
        "tipo_usado": tipo,
        "fator": fator,
        "caixas": int(caixas),
        "match_score": round(score, 3),
        "match_key": matched_key,
    }, ""

def group_matrix(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Produto"])
    out = (
        df.groupby(["Produto", "Coluna"], as_index=False)["Caixas"]
        .sum()
        .pivot(index="Produto", columns="Coluna", values="Caixas")
        .fillna(0)
        .reset_index()
        .sort_values("Produto")
        .reset_index(drop=True)
    )
    for c in out.columns:
        if c != "Produto":
            out[c] = out[c].astype(int)
    return out

def make_xlsx(df: pd.DataFrame, sheet_name: str) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (df if not df.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        ws.column_dimensions["A"].width = 42
        for col in ["B", "C", "D", "E", "F", "G", "H"]:
            ws.column_dimensions[col].width = 12
    out.seek(0)
    return out.getvalue()

def make_cd_xlsx(frutas: pd.DataFrame, legumes: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        (frutas if not frutas.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name="FRUTAS", index=False)
        (legumes if not legumes.empty else pd.DataFrame(columns=["Produto"])).to_excel(writer, sheet_name="LEGUMES", index=False)
        for sh in ["FRUTAS", "LEGUMES"]:
            ws = writer.sheets[sh]
            ws.column_dimensions["A"].width = 42
            for col in ["B", "C", "D", "E", "F", "G", "H"]:
                ws.column_dimensions[col].width = 12
    out.seek(0)
    return out.getvalue()

def make_log_xlsx(files_df: pd.DataFrame, errors_df: pd.DataFrame, matches_df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        files_df.to_excel(writer, sheet_name="ARQUIVOS", index=False)
        errors_df.to_excel(writer, sheet_name="ERROS", index=False)
        matches_df.to_excel(writer, sheet_name="MATCHES", index=False)
    out.seek(0)
    return out.getvalue()

def make_zip(files: Dict[str, bytes]) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)
    out.seek(0)
    return out.getvalue()

# =========================================================
# UI
# =========================================================

st.title("📦 BRASÃO / KROSS → THOTH (Modo Inteligente)")
st.caption("Suba os PDFs. O sistema identifica loja, limpa os produtos, faz match inteligente com a base e gera as planilhas separadas de Brasão, Kross e CD.")

with st.sidebar:
    st.subheader("Fluxo")
    st.write("1. Suba os PDFs.")
    st.write("2. Clique em PROCESSAR.")
    st.write("3. Baixe o ZIP final.")
    use_demo = st.checkbox("Usar base modelo embutida", value=True)
    base_file = st.file_uploader("Base mestre (opcional)", type=["xlsx", "xls", "csv"])
    pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

base_df = pd.DataFrame(DEMO_BASE) if use_demo and base_file is None else build_base_df(base_file) if base_file is not None else pd.DataFrame(DEMO_BASE)
base_map = build_base_map(base_df)

if st.button("PROCESSAR", type="primary", use_container_width=True):
    if not pdf_files:
        st.error("Envie ao menos um PDF.")
    else:
        try:
            files_info = []
            converted_rows = []
            error_rows = []
            match_rows = []
            seen_units = set()

            for pdf in pdf_files:
                pdf_text = extract_text_from_pdf(pdf.getvalue())
                store = identify_store(pdf.name, pdf_text)

                if store["store_id"] in seen_units:
                    error_rows.append({"Loja": store["display"], "Produto": pdf.name, "Erro": "PDF duplicado da mesma unidade"})
                    continue
                seen_units.add(store["store_id"])

                parsed = parse_order_pdf(pdf_text)
                files_info.append({
                    "Arquivo": pdf.name,
                    "Loja": store["display"],
                    "Grupo": store["group"],
                    "Itens extraídos": len(parsed),
                })

                for _, row in parsed.iterrows():
                    conv, err = convert_item(row["produto_normalizado"], row["qtd"], row["tipo_detectado"], base_map)
                    if conv is None:
                        error_rows.append({"Loja": store["display"], "Produto": row["produto_original"], "Erro": err})
                        continue

                    converted_rows.append({
                        "Grupo": store["group"],
                        "Loja": store["display"],
                        "Coluna": store["col_key"],
                        "Produto": conv["produto_base"],
                        "Categoria": conv["categoria"],
                        "Caixas": conv["caixas"],
                    })

                    match_rows.append({
                        "Loja": store["display"],
                        "Produto PDF": row["produto_original"],
                        "Produto normalizado": row["produto_normalizado"],
                        "Produto base": conv["produto_base"],
                        "Score": conv["match_score"],
                        "Tipo": conv["tipo_usado"],
                        "Qtd PDF": row["qtd"],
                        "Fator": conv["fator"],
                        "Caixas": conv["caixas"],
                    })

            converted_df = pd.DataFrame(converted_rows)
            errors_df = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(columns=["Loja", "Produto", "Erro"])
            files_df = pd.DataFrame(files_info)
            matches_df = pd.DataFrame(match_rows) if match_rows else pd.DataFrame(columns=["Loja","Produto PDF","Produto normalizado","Produto base","Score","Tipo","Qtd PDF","Fator","Caixas"])

            if converted_df.empty:
                st.error("Nenhum item foi convertido. A base precisa de mais produtos ou os PDFs vieram em formato difícil de leitura.")
                if not errors_df.empty:
                    st.dataframe(errors_df, use_container_width=True)
                st.stop()

            brasao = converted_df[converted_df["Grupo"] == "BRASAO"]
            kross = converted_df[converted_df["Grupo"] == "KROSS"]
            cd = converted_df[converted_df["Grupo"] == "BRASAO_CD"]

            brasao_frutas = group_matrix(brasao[brasao["Categoria"] == "FRUTAS"])
            brasao_legumes = group_matrix(brasao[brasao["Categoria"] == "LEGUMES"])
            kross_frutas = group_matrix(kross[kross["Categoria"] == "FRUTAS"])
            kross_legumes = group_matrix(kross[kross["Categoria"] == "LEGUMES"])
            cd_frutas = group_matrix(cd[cd["Categoria"] == "FRUTAS"])
            cd_legumes = group_matrix(cd[cd["Categoria"] == "LEGUMES"])

            zip_files = {
                "BRASAO_FRUTAS_Thoth.xlsx": make_xlsx(brasao_frutas, "BRASAO_FRUTAS"),
                "BRASAO_LEGUMES_Thoth.xlsx": make_xlsx(brasao_legumes, "BRASAO_LEGUMES"),
                "KROSS_FRUTAS_Thoth.xlsx": make_xlsx(kross_frutas, "KROSS_FRUTAS"),
                "KROSS_LEGUMES_Thoth.xlsx": make_xlsx(kross_legumes, "KROSS_LEGUMES"),
                "BRASAO_CD.xlsx": make_cd_xlsx(cd_frutas, cd_legumes),
                "LOG_PROCESSAMENTO.xlsx": make_log_xlsx(files_df, errors_df, matches_df),
            }
            zip_bytes = make_zip(zip_files)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("PDFs", len(pdf_files))
            c2.metric("Itens convertidos", len(converted_df))
            c3.metric("Ocorrências", len(errors_df))
            c4.metric("Match médio", round(matches_df["Score"].mean(), 2) if not matches_df.empty else 0)

            st.success("Modo inteligente concluído.")

            tab1, tab2, tab3, tab4 = st.tabs(["Arquivos", "Matches", "Erros", "Prévia"])
            with tab1:
                st.dataframe(files_df, use_container_width=True)
            with tab2:
                st.dataframe(matches_df, use_container_width=True)
            with tab3:
                st.dataframe(errors_df, use_container_width=True)
            with tab4:
                st.write("Brasão Frutas")
                st.dataframe(brasao_frutas, use_container_width=True)
                st.write("Brasão Legumes")
                st.dataframe(brasao_legumes, use_container_width=True)
                st.write("Kross Frutas")
                st.dataframe(kross_frutas, use_container_width=True)
                st.write("Kross Legumes")
                st.dataframe(kross_legumes, use_container_width=True)

            st.download_button(
                "Baixar ZIP final",
                data=zip_bytes,
                file_name="THOTH_BRASAO_KROSS_MODO_INTELIGENTE.zip",
                mime="application/zip",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
