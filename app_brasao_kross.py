import io
import math
import re
import unicodedata
import zipfile
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st
from difflib import SequenceMatcher

st.set_page_config(page_title="THOTH Brasão/Kross Final", page_icon="📦", layout="wide")

EMBEDDED_BASE = [{'produto': 'Abacate Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 20.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'ABACATE BREDA'}, {'produto': 'Abacaxi Perola Und', 'categoria': 'FRUTAS', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'ABACAXI PEROLA T. 8'}, {'produto': 'Alecrim Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'ALECRIM'}, {'produto': 'Alho Poro Und', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 12.0, 'sinonimos': 'ALHO PORO'}, {'produto': 'Ameixa Nacional Demarchi Bdj 500g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 20.0, 'unidades_caixa': 0, 'sinonimos': 'AMEIXA NACIONAL BDJ'}, {'produto': 'Batata Doce Branca Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'BATATA DOCE BRANCA'}, {'produto': 'Batata Doce Roxa Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'BATATA DOCE ROSADA A'}, {'produto': 'Batata Salsa Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 15.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MANDIOQUINHA MÉDIA'}, {'produto': 'Berinjela Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'BERINJELA AAA'}, {'produto': 'Beterraba Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'BETERRABA AA'}, {'produto': 'Caqui Rama Forte Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 6.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'CAQUI RAMA FORTE'}, {'produto': 'Carambola De Marchi 400g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 4.0, 'unidades_caixa': 0, 'sinonimos': 'CARAMBOLA BDJ 500GR'}, {'produto': 'Cebola Argentina Branca Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 20.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'CEBOLA IMPORTADA CX 03'}, {'produto': 'Cebola Conserva Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 20.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'CEBOLA CONSERVA'}, {'produto': 'Cenoura Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'CENOURA A'}, {'produto': 'Chuchu Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'CHUCHU'}, {'produto': 'Coco Seco Fruta Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 20.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'COCO SECO'}, {'produto': 'Coentro Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'COENTRO'}, {'produto': 'Figo Roxo De Marchi 300g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 3.0, 'unidades_caixa': 0, 'sinonimos': 'FIGO ROXO BAND'}, {'produto': 'Framboesa Fruta 120g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 10.0, 'unidades_caixa': 0, 'sinonimos': 'FRAMBOESA BAND'}, {'produto': 'Goiaba Nacional Vermelha Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 6.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'GOIABA VERMELHA'}, {'produto': 'Hortelã Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'HORTELA'}, {'produto': 'Jatobá Fruta Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 4.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'JATOBÁ'}, {'produto': 'Kinkan Bandeja Frutamina 500g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 2.0, 'unidades_caixa': 0, 'sinonimos': 'KINKAN BAND'}, {'produto': 'Kiwi Importado Grecia Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 9.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'KIWI IMPORTADO CAL 23'}, {'produto': 'Kiwi Nacional De Marchi Bandeja 600g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 1.0, 'unidades_caixa': 0, 'sinonimos': 'KIWI BANDEJA'}, {'produto': 'Louro Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'LOURO MAÇO'}, {'produto': 'Maçã Fuji Cat 1 Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MACA FUJI CAL 100'}, {'produto': 'Mamão Formosa Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MAMÃO FORMOSA'}, {'produto': 'Mamãozinho Papaya Unidade', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MAMÃO PAPAYA T. 15'}, {'produto': 'Manga Palmer Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 9.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MANGA PALMER'}, {'produto': 'Manjericão Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'MANJERICÃO'}, {'produto': 'Manjerona Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'MANJERONA'}, {'produto': 'Maxixe Bdj De Marchi 300g', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 13.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MAXIXI'}, {'produto': 'Melão Cantaloupe Unidade', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 12.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO CANTALOPE'}, {'produto': 'Melão Charanteais Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 9.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELAO CHARANTEAIS'}, {'produto': 'Melão Dino Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO DINO'}, {'produto': 'Melão Espanhol Amarelo Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 13.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO AMARELO'}, {'produto': 'Melão Galia Unidade', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 6.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO GALIA'}, {'produto': 'Melão Orange Unidade', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 6.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO ORANGE'}, {'produto': 'Melão Rei Doce Redinha Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO (REDE)'}, {'produto': 'Melão Sapo Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 10.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'MELÃO (REI)'}, {'produto': 'Milho Verde Espiga De Marchi Bdj 700g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 10.0, 'unidades_caixa': 0, 'sinonimos': 'MILHO BAND'}, {'produto': 'Milho Verde Espiga De Marchi Bdj 700g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 34.0, 'unidades_caixa': 0, 'sinonimos': 'SWEET MILHO 450GR'}, {'produto': 'Milho Verde Espiga De Marchi Bdj 700g', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'MILHO VERDE DOCE'}, {'produto': 'Mirtilo Blueberry Imp. Demarchi 125g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 12.0, 'unidades_caixa': 0, 'sinonimos': 'MIRTILLO'}, {'produto': 'Nabo Unidade', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'NABO'}, {'produto': 'Pepino Japonês Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PEPINO JAPONÊS'}, {'produto': 'Pêra Williams Argentina Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 18.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PÊRA WILLIANS (CAL 120/135)'}, {'produto': 'Pêssego Imp Argentina Polpa Amarela Kg', 'categoria': 'FRUTAS', 'tipo': 'KG', 'peso_caixa': 9.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PESSEGO AMARELO IMPORTADO'}, {'produto': 'Physalis Importado Colombia 100g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 8.0, 'unidades_caixa': 0, 'sinonimos': 'PHYSALES'}, {'produto': 'Pimenta Biquinho Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 1.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PIMENTA BIQUINHO'}, {'produto': 'Pimenta Cambuci Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 8.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PIMENTA CAMBUCI'}, {'produto': 'Pimenta Jalapeño Kg', 'categoria': 'LEGUMES', 'tipo': 'KG', 'peso_caixa': 1.0, 'bandejas_por_caixa': 0, 'unidades_caixa': 0, 'sinonimos': 'PIMENTA JALAPENO'}, {'produto': 'Salsão Aipo Unidade', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 12.0, 'sinonimos': 'SALSAO UND'}, {'produto': 'Salsão Aipo Unidade', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'SALSÃO BDJ 37419'}, {'produto': 'Salsão Aipo Unidade', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'SALSÃO VERDE 38234'}, {'produto': 'Sálvia Unidade', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'SALVIA'}, {'produto': 'Tomate Grape Demarchi 180g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 24.0, 'unidades_caixa': 0, 'sinonimos': 'TOMATE SWEET GRAPE'}, {'produto': 'Tomate Grape Demarchi 180g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 24.0, 'unidades_caixa': 0, 'sinonimos': 'TOMATE YELOW GRAPE'}, {'produto': 'Tomate Grape Demarchi 180g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 1.0, 'unidades_caixa': 0, 'sinonimos': 'TOMATE SWEET GRAPE'}, {'produto': 'Tomate Grape Demarchi 180g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 1.0, 'unidades_caixa': 0, 'sinonimos': 'TOMATE RED GRAPE'}, {'produto': 'Tomate Grape Demarchi 180g', 'categoria': 'LEGUMES', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 1.0, 'unidades_caixa': 0, 'sinonimos': 'TOMATE GRAPE 180G'}, {'produto': 'Tomilho Maço', 'categoria': 'LEGUMES', 'tipo': 'UND', 'peso_caixa': 0, 'bandejas_por_caixa': 0, 'unidades_caixa': 1.0, 'sinonimos': 'TOMILHO'}, {'produto': 'Uva Thompson S/Semente Demarchi Bdj 500g', 'categoria': 'FRUTAS', 'tipo': 'BDJ', 'peso_caixa': 0, 'bandejas_por_caixa': 10.0, 'unidades_caixa': 0, 'sinonimos': 'UVA THOMPSON BAND 500g CAT1'}]

STORE_RULES = [
    {"store_id":"BRASAO_F","group":"BRASAO","col":"BRASAO F","label":"Brasão Fernando","file_signals":["FERNANDO","CE"],"text_signals":["FERNANDO MACHADO","CENTRO, 226"]},
    {"store_id":"BRASAO_J","group":"BRASAO","col":"BRASAO J","label":"Brasão Jardim","file_signals":["JARDIM","JA"],"text_signals":["JARDIM AMERICA","SAO PEDRO","2199"]},
    {"store_id":"BRASAO_X","group":"BRASAO","col":"BRASAO X","label":"Brasão Xaxim","file_signals":["XAXIM","XX"],"text_signals":["LUIZ LUNARDI","XAXIM","810"]},
    {"store_id":"BRASAO_A","group":"BRASAO","col":"BRASAO A","label":"Brasão Avenida","file_signals":["AVENIDA","AV"],"text_signals":["RIO DE JANEIRO","CENTRO, 108"]},
    {"store_id":"KROSS_AT","group":"KROSS","col":"KROSS AT","label":"Kross Atacadista","file_signals":["KROSS","ATACADO","CHAPECO"],"text_signals":["JOHN KENNEDY","PASSO DOS FORTES","550"]},
    {"store_id":"KROSS_X","group":"KROSS","col":"KROSS XAXIM","label":"Kross Xaxim","file_signals":["KROSS","XAXIM","XX"],"text_signals":["AMELIO PANIZZI","XAXIM"]},
    {"store_id":"BRASAO_CD","group":"CD","col":"CD","label":"Brasão CD","file_signals":["CD"],"text_signals":["RUA GASPAR","ELDORADO","153"]},
]

STOP_WORDS = [
    "NUMERO DO PEDIDO","TRANSACAO","EMPRESA:","CEP:","TROCAS:","JOAB FATURAMENTO","KELLY",
    "LEANDRO GERENTE","E-MAIL:","PG:","PEDIDO DE COMPRA","DCTO:","USUARIO:","TOTAL DO FORNECEDOR",
    "CONTATOS DO FORNECEDOR","PAGINA","ORIG DEST","RUA ","AV. ","AV ","RUA SAO PEDRO"
]

REMOVE_TOKENS = [
    "BRASAO","FRUTA","FRU","DEMARCHI","DE MARCHI","SHELF","BANDEJA","BDJ","KG","UND","UNIDADE","UN",
    "MACO","CAT","CX","CAIXA","IMPORTADO","NACIONAL","ARGENTINA","GRECIA","DE MARCHI","FRUTAMINA"
]

def norm(v):
    t = str(v or "").strip().upper()
    t = unicodedata.normalize("NFKD", t).encode("ASCII","ignore").decode("utf-8")
    return re.sub(r"\s+", " ", t).strip()

def parse_br_number(s):
    s = str(s).strip()
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def load_base(uploaded):
    if uploaded is None:
        return pd.DataFrame(EMBEDDED_BASE)
    if uploaded.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded)
    return pd.read_excel(uploaded)

def canonical_name(txt):
    t = norm(txt)
    t = re.sub(r"\b\d{5,}\b", " ", t)
    t = re.sub(r"\b\d+G\b", " ", t)
    t = re.sub(r"\b\d+,\d+\b", " ", t)
    t = re.sub(r"\b\d+\b", " ", t)
    for tok in REMOVE_TOKENS:
        t = re.sub(rf"\b{re.escape(tok)}\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip()

    manual = {
        "ABACAXI PEROLA":"ABACAXI PEROLA UND",
        "ABACATE":"ABACATE KG",
        "ALECRIM":"ALECRIM MACO",
        "ALHO PORO":"ALHO PORO UND",
        "AMEIXA":"AMEIXA NACIONAL DEMARCHI BDJ 500G",
        "BATATA DOCE BRANCA":"BATATA DOCE BRANCA KG",
        "BATATA DOCE ROXA":"BATATA DOCE ROXA KG",
        "BATATA SALSA":"BATATA SALSA KG",
        "BERINJELA":"BERINJELA KG",
        "BETERRABA":"BETERRABA KG",
        "CAQUI RAMA FORTE":"CAQUI RAMA FORTE KG",
        "CARAMBOLA":"CARAMBOLA DE MARCHI 400G",
        "CEBOLA ARGENTINA BRANCA":"CEBOLA ARGENTINA BRANCA KG",
        "CEBOLA CONSERVA":"CEBOLA CONSERVA KG",
        "CENOURA":"CENOURA KG",
        "CHUCHU":"CHUCHU KG",
        "COCO SECO":"COCO SECO FRUTA KG",
        "COENTRO":"COENTRO MACO",
        "FIGO ROXO":"FIGO ROXO DE MARCHI 300G",
        "FRAMBOESA":"FRAMBOESA FRUTA 120G",
        "GOIABA":"GOIABA NACIONAL VERMELHA KG",
        "HORTELA":"HORTELA MACO",
        "JATOBA":"JATOBA FRUTA KG",
        "KIWI IMPORTADO":"KIWI IMPORTADO GRECIA KG",
        "KIWI NACIONAL":"KIWI NACIONAL DE MARCHI 600G",
        "LARANJA MAQUINA DE SUCO":"LARANJA MAQUINA DE SUCO KG",
        "LIMAO SICILIANO":"LIMAO SICILIANO KG",
        "LIMAO TAHITI":"LIMAO TAHITI KG",
        "LOURO":"LOURO MACO",
        "MACA FUJI":"MACA FUJI KG",
        "MAMAO FORMOSA":"MAMAO FORMOSA KG",
        "MAMAOZINHO PAPAIA":"MAMAOZINHO PAPAIA UND",
        "MANGA PALMER":"MANGA PALMER KG",
        "MANJERICAO":"MANJERICAO MACO",
        "MANJERONA":"MANJERONA MACO",
        "MAXIXE":"MAXIXE DE MARCHI 300G",
        "MELAO CANTALOUPE":"MELAO CANTALOUPE UND",
        "MELAO DINO":"MELAO DINO KG",
        "MELAO ESPANHOL AMARELO":"MELAO ESPANHOL AMARELO KG",
        "MELAO GALIA":"MELAO GALIA UND",
        "MELAO ORANGE":"MELAO ORANGE UND",
        "MELAO REI DOCE REDINHA":"MELAO REI DOCE REDINHA KG",
        "MELAO SAPO":"MELAO SAPO KG",
        "MELANCIA INTEIRA":"MELANCIA INTEIRA KG",
        "MILHO VERDE":"MILHO VERDE ESPIGA 700G",
        "NABO":"NABO UND",
        "PEPINO JAPONES":"PEPINO JAPONES KG",
        "PERA WILLIANS":"PERA WILLIANS ARGENTINA KG",
        "PESSEGO IMP":"PESSEGO IMP ARGENTINA KG",
        "PIMENTA BIQUINHO":"PIMENTA BIQUINHO KG",
        "PIMENTA JALAPENO":"PIMENTA JALAPENO KG",
        "SALSAO AIPO":"SALSAO AIPO UND",
        "SALVIA":"SALVIA UND",
        "TOMILHO":"TOMILHO MACO",
        "TOMATE GRAPE":"TOMATE GRAPE DEMARCHI 180G",
        "UVA THOMPSON":"UVA THOMPSON S/SEMENTE DEMARCHI BDJ 500G",
    }
    return manual.get(t, t)

def build_base_map(df):
    out = {}
    for _, r in df.iterrows():
        produto = canonical_name(r.get("produto", r.get("Produto", "")))
        item = {
            "produto": produto,
            "categoria": norm(r.get("categoria", r.get("Categoria", ""))),
            "tipo": norm(r.get("tipo", r.get("Tipo", ""))),
            "peso_caixa": float(r.get("peso_caixa", r.get("Medida", 0)) or 0),
            "bandejas_por_caixa": float(r.get("bandejas_por_caixa", r.get("Medida", 0)) or 0),
            "unidades_caixa": float(r.get("unidades_caixa", r.get("Medida", 0)) or 0),
        }
        # se veio da base embedded por unidade
        if "Unidade" in df.columns:
            unidade = norm(r.get("Unidade", ""))
            if unidade == "KG":
                item["tipo"] = "KG"; item["peso_caixa"] = float(r.get("Medida", 0) or 0); item["bandejas_por_caixa"] = 0; item["unidades_caixa"] = 0
            elif unidade == "BANDEJA":
                item["tipo"] = "BDJ"; item["bandejas_por_caixa"] = float(r.get("Medida", 0) or 0); item["peso_caixa"] = 0; item["unidades_caixa"] = 0
            else:
                item["tipo"] = "UND"; item["unidades_caixa"] = float(r.get("Medida", 0) or 0); item["peso_caixa"] = 0; item["bandejas_por_caixa"] = 0
            item["categoria"] = "FRUTAS" if norm(r.get("Tipo","")) == "FRUTAS" else "LEGUMES"

        out[produto] = item
        syn = str(r.get("sinonimos", r.get("Produto Demarchi", "")) or "")
        for s in [x.strip() for x in syn.split(";") if x.strip()]:
            out[canonical_name(s)] = item
    return out

def best_match(name, base_map):
    if name in base_map:
        return base_map[name], 1.0
    best_item, best_score = None, 0
    for k, item in base_map.items():
        score = SequenceMatcher(None, name, k).ratio()
        if name in k or k in name:
            score = max(score, 0.92)
        if score > best_score:
            best_item, best_score = item, score
    if best_score >= 0.78:
        return best_item, best_score
    return None, best_score

def identify_store(filename, text):
    fn = norm(filename)
    tx = norm(text[:4000])
    best = STORE_RULES[0]
    best_score = -1
    for rule in STORE_RULES:
        score = 0
        for sig in rule["file_signals"]:
            if norm(sig) in fn:
                score += 5
        for sig in rule["text_signals"]:
            if norm(sig) in tx:
                score += 3
        if score > best_score:
            best = rule
            best_score = score
    return best

def pdf_text(file):
    parts = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            parts.append(p.extract_text() or "")
    return "\n".join(parts)

def is_noise(line):
    s = norm(line)
    if len(s) < 3:
        return True
    return any(tok in s for tok in STOP_WORDS)

def extract_items(text):
    rows = []
    for raw in text.splitlines():
        line = norm(raw)
        if not line or is_noise(line):
            continue
        # reconhece unidade
        unit_match = re.search(r"\b(KG|BDJ|BANDEJA|UND|UNIDADE|UN|MACO)\b", line)
        if not unit_match:
            continue
        unit = unit_match.group(1)
        before = line[:unit_match.start()].strip()
        after = line[unit_match.end():].strip()
        # produto: remover códigos iniciais
        produto = re.sub(r"^(?:\d+[\s,\.]*)+", "", before).strip()
        # quantidade: primeiro número decimal/br depois da unidade
        qty_match = re.search(r"(\d{1,4},\d{3}|\d+)", after)
        if not qty_match:
            continue
        qtd = parse_br_number(qty_match.group(1))
        if qtd is None or qtd <= 0:
            continue
        rows.append({
            "produto_pdf": produto,
            "produto_norm": canonical_name(produto),
            "tipo_detectado": "BDJ" if unit == "BANDEJA" else ("UND" if unit in {"UNIDADE","UN","MACO"} else unit),
            "qtd": qtd,
            "linha": line,
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.drop_duplicates(subset=["linha"]).reset_index(drop=True)

def convert_row(produto, qtd, tipo_detectado, base_map):
    item, score = best_match(produto, base_map)
    if item is None:
        return None, f"SEM BASE: {produto}"
    tipo = item["tipo"] or tipo_detectado
    if tipo == "KG":
        fator = item["peso_caixa"]
    elif tipo == "BDJ":
        fator = item["bandejas_por_caixa"]
    else:
        tipo = "UND"
        fator = item["unidades_caixa"]
    if not fator or fator <= 0:
        return None, f"BASE INCOMPLETA: {item['produto']}"
    caixas = math.ceil(qtd / fator)
    return {
        "produto_base": item["produto"],
        "categoria": item["categoria"],
        "tipo": tipo,
        "fator": fator,
        "caixas": int(caixas),
        "score": round(score,3)
    }, ""

def matrix_for(df, columns):
    if df.empty:
        out = pd.DataFrame(columns=["Produto"] + columns)
        return out
    piv = (
        df.groupby(["Produto","Coluna"], as_index=False)["Caixas"]
        .sum()
        .pivot(index="Produto", columns="Coluna", values="Caixas")
        .fillna(0)
        .reset_index()
        .sort_values("Produto")
        .reset_index(drop=True)
    )
    for col in columns:
        if col not in piv.columns:
            piv[col] = 0
    piv = piv[["Produto"] + columns]
    for col in columns:
        piv[col] = piv[col].astype(int)
    return piv

def write_xlsx(sheets):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            (df if not df.empty else pd.DataFrame(columns=df.columns if hasattr(df, "columns") else ["Produto"])).to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.column_dimensions["A"].width = 42
            for col in ["B","C","D","E","F","G","H"]:
                ws.column_dimensions[col].width = 14
    out.seek(0)
    return out.getvalue()

def build_zip(file_map):
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in file_map.items():
            zf.writestr(name, content)
    out.seek(0)
    return out.getvalue()

st.title("📦 THOTH Brasão / Kross Final")
st.caption("Suba os PDFs e gere automaticamente Brasão Frutas, Brasão Legumes, Kross Frutas, Kross Legumes e Brasão CD.")

with st.sidebar:
    st.subheader("Entrada")
    st.write("1. Suba os PDFs")
    st.write("2. Clique em PROCESSAR")
    st.write("3. Baixe o ZIP final")
    base_file = st.file_uploader("Base mestre (opcional)", type=["xlsx","xls","csv"])
    pdf_files = st.file_uploader("Pedidos PDF", type=["pdf"], accept_multiple_files=True)

if st.button("PROCESSAR", type="primary", use_container_width=True):
    if not pdf_files:
        st.error("Envie pelo menos um PDF.")
        st.stop()

    try:
        base_df = load_base(base_file)
        base_map = build_base_map(base_df)

        conv_rows = []
        err_rows = []
        match_rows = []
        file_rows = []

        for pdf in pdf_files:
            text = pdf_text(pdf)
            store = identify_store(pdf.name, text)
            parsed = extract_items(text)

            file_rows.append({
                "Arquivo": pdf.name,
                "Loja": store["label"],
                "Grupo": store["group"],
                "Itens extraídos": len(parsed)
            })

            for _, r in parsed.iterrows():
                conv, err = convert_row(r["produto_norm"], r["qtd"], r["tipo_detectado"], base_map)
                if conv is None:
                    err_rows.append({"Loja": store["label"], "Produto": r["produto_pdf"], "Erro": err})
                    continue
                conv_rows.append({
                    "Grupo": store["group"],
                    "Coluna": store["col"],
                    "Produto": conv["produto_base"],
                    "Categoria": conv["categoria"],
                    "Caixas": conv["caixas"],
                })
                match_rows.append({
                    "Loja": store["label"],
                    "Produto PDF": r["produto_pdf"],
                    "Produto base": conv["produto_base"],
                    "Qtd PDF": r["qtd"],
                    "Tipo": conv["tipo"],
                    "Fator": conv["fator"],
                    "Caixas": conv["caixas"],
                    "Score": conv["score"],
                })

        conv_df = pd.DataFrame(conv_rows)
        errors_df = pd.DataFrame(err_rows) if err_rows else pd.DataFrame(columns=["Loja","Produto","Erro"])
        files_df = pd.DataFrame(file_rows)
        matches_df = pd.DataFrame(match_rows) if match_rows else pd.DataFrame(columns=["Loja","Produto PDF","Produto base","Qtd PDF","Tipo","Fator","Caixas","Score"])

        if conv_df.empty:
            st.error("Nenhum item convertido. Verifique os PDFs e a base.")
            if not errors_df.empty:
                st.dataframe(errors_df, use_container_width=True)
            st.stop()

        brasao = conv_df[conv_df["Grupo"] == "BRASAO"]
        kross = conv_df[conv_df["Grupo"] == "KROSS"]
        cd = conv_df[conv_df["Grupo"] == "CD"]

        brasao_frutas = matrix_for(brasao[brasao["Categoria"] == "FRUTAS"], ["BRASAO F","BRASAO J","BRASAO X","BRASAO A"])
        brasao_legumes = matrix_for(brasao[brasao["Categoria"] == "LEGUMES"], ["BRASAO F","BRASAO J","BRASAO X","BRASAO A"])
        kross_frutas = matrix_for(kross[kross["Categoria"] == "FRUTAS"], ["KROSS AT","KROSS XAXIM"])
        kross_legumes = matrix_for(kross[kross["Categoria"] == "LEGUMES"], ["KROSS AT","KROSS XAXIM"])
        cd_frutas = matrix_for(cd[cd["Categoria"] == "FRUTAS"], ["CD"])
        cd_legumes = matrix_for(cd[cd["Categoria"] == "LEGUMES"], ["CD"])

        zip_bytes = build_zip({
            "BRASAO_FRUTAS_Thoth.xlsx": write_xlsx({"BRASAO_FRUTAS": brasao_frutas}),
            "BRASAO_LEGUMES_Thoth.xlsx": write_xlsx({"BRASAO_LEGUMES": brasao_legumes}),
            "KROSS_FRUTAS_Thoth.xlsx": write_xlsx({"KROSS_FRUTAS": kross_frutas}),
            "KROSS_LEGUMES_Thoth.xlsx": write_xlsx({"KROSS_LEGUMES": kross_legumes}),
            "BRASAO_CD.xlsx": write_xlsx({"FRUTAS": cd_frutas, "LEGUMES": cd_legumes}),
            "LOG_PROCESSAMENTO.xlsx": write_xlsx({"ARQUIVOS": files_df, "MATCHES": matches_df, "ERROS": errors_df}),
        })

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("PDFs", len(pdf_files))
        c2.metric("Itens convertidos", len(conv_df))
        c3.metric("Ocorrências", len(errors_df))
        c4.metric("Match médio", round(matches_df["Score"].mean(), 2) if not matches_df.empty else 0)

        tab1, tab2, tab3, tab4 = st.tabs(["Prévia", "Matches", "Erros", "Arquivos"])
        with tab1:
            st.write("Brasão Legumes")
            st.dataframe(brasao_legumes, use_container_width=True)
            st.write("Brasão Frutas")
            st.dataframe(brasao_frutas, use_container_width=True)
            st.write("Kross Legumes")
            st.dataframe(kross_legumes, use_container_width=True)
            st.write("Kross Frutas")
            st.dataframe(kross_frutas, use_container_width=True)
        with tab2:
            st.dataframe(matches_df, use_container_width=True)
        with tab3:
            st.dataframe(errors_df, use_container_width=True)
        with tab4:
            st.dataframe(files_df, use_container_width=True)

        st.download_button("Baixar ZIP final", data=zip_bytes, file_name="THOTH_FINAL_BRASAO_KROSS.zip", mime="application/zip", use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
