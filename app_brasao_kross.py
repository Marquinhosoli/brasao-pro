def extrair_texto_pdf(uploaded_file) -> str:
    partes = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            if txt.strip():
                partes.append(txt)
    return "\n".join(partes)


def extrair_linhas_relevantes(texto: str):
    linhas = []
    for linha in texto.splitlines():
        limpa = re.sub(r"\s+", " ", linha).strip()
        if limpa:
            linhas.append(limpa)
    return linhas


def limpar_produto(txt: str) -> str:
    txt = normalizar_nome(txt)
    txt = re.sub(r"\bLOJA\s*\d+\b", "", txt)
    txt = re.sub(r"\bCE\b|\bJA\b|\bXX\b|\bAV\b|\bCD\b", "", txt)
    txt = re.sub(r"[-–—:]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def parse_linha_produto(linha: str):
    l = normalizar_nome(linha)
    l = re.sub(r"\s+", " ", l).strip()

    # ignora cabeçalhos comuns
    ignorar = [
        "PEDIDO", "CLIENTE", "TOTAL", "OBS", "OBSERVACAO", "PAGINA",
        "HORTIFRUTI", "DATA", "EMISSAO", "ENDERECO", "VENDEDOR"
    ]
    if any(x in l for x in ignorar):
        return None

    # formato: PRODUTO 400 KG
    m = re.search(r"^(.*?)[ ]+(\d+[.,]?\d*)[ ]*(KG|KGS|QUILO|QUILOS|UN|UND|UNID|UNIDADE|UNIDADES|BDJ|BANDEJA|BANDEJAS)\s*$", l)
    if m:
        produto = limpar_produto(m.group(1))
        qtd = float(m.group(2).replace(",", "."))
        unidade = m.group(3)

        if unidade in ["KG", "KGS", "QUILO", "QUILOS"]:
            unidade = "kg"
        elif unidade in ["UN", "UND", "UNID", "UNIDADE", "UNIDADES"]:
            unidade = "un"
        else:
            unidade = "bdj"

        if produto and qtd > 0:
            return produto, qtd, unidade

    # formato: PRODUTO 400
    m = re.search(r"^(.*?)[ ]+(\d+[.,]?\d*)\s*$", l)
    if m:
        produto = limpar_produto(m.group(1))
        qtd = float(m.group(2).replace(",", "."))

        if not produto or qtd <= 0:
            return None

        # tenta descobrir a unidade pela base
        nome_base, info = localizar_base(produto)
        if info:
            unidade = info["modo"]
        else:
            unidade = "kg"  # padrão provisório para não perder item

        return produto, qtd, unidade

    return None


def processar_arquivo(uploaded_file):
    nome = uploaded_file.name
    cliente = detectar_cliente(nome)

    if nome.lower().endswith(".pdf"):
        texto = extrair_texto_pdf(uploaded_file)
        linhas = extrair_linhas_relevantes(texto)
    else:
        df_excel = pd.read_excel(uploaded_file, header=None)
        linhas = []
        for _, row in df_excel.iterrows():
            partes = [str(x).strip() for x in row if pd.notna(x) and str(x).strip()]
            if partes:
                linhas.append(" ".join(partes))

    itens = []
    ignoradas = []

    for linha in linhas:
        item = parse_linha_produto(linha)
        if item:
            produto, qtd, unidade = item
            conv = converter_para_caixa(produto, qtd, unidade)
            conv["cliente"] = cliente
            conv["arquivo"] = nome
            conv["linha_original"] = linha
            itens.append(conv)
        else:
            ignoradas.append(linha)

    return itens, ignoradas
