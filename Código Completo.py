# ==========================================================
# Coleta de produtos do portal LMMoto + armazenamento em SQL
# + Importação automática de vários arquivos Excel com detecção do cabeçalho
# + Indicação da origem dos produtos (site ou arquivo .xlsx)
# + Correção automática de tabela antiga (coluna origem)
# ==========================================================

import requests
from lxml import html
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import sqlite3
import sys
import pandas as pd
import glob
import os
import unicodedata

# ----------------------------------------------------------
# CONFIGURAÇÕES
# ----------------------------------------------------------
login_url = "https://portal.lmmoto.com.br/glstorefront/glmotos/pt/BRL/login"
USERNAME = "contato@siromotos.com.br"
PASSWORD = "Whats_123"

# XPaths do formulário
xpath_user = "/html/body/main/div[3]/div/div[1]/div/div/div/form/div[1]/input"
xpath_pass = "/html/body/main/div[3]/div/div[1]/div/div/div/form/div[2]/input"
xpath_form = "/html/body/main/div[3]/div/div[1]/div/div/div/form"

headers = {
    "User-Agent": "Mozilla/5.0 (compatible; Python requests)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ----------------------------------------------------------
# FUNÇÃO: LOGIN
# ----------------------------------------------------------
def login():
    session = requests.Session()
    r = session.get(login_url, headers=headers, timeout=15)
    tree = html.fromstring(r.text)

    form_els = tree.xpath(xpath_form)
    if not form_els:
        sys.exit("❌ Formulário de login não encontrado.")
    form = form_els[0]

    action_url = urljoin(login_url, form.get("action") or login_url)
    user_input = tree.xpath(xpath_user)
    pass_input = tree.xpath(xpath_pass)
    if not user_input or not pass_input:
        sys.exit("❌ Campos de usuário ou senha não encontrados.")

    user_name = user_input[0].get("name")
    pass_name = pass_input[0].get("name")
    if not user_name or not pass_name:
        sys.exit("❌ Campos sem atributo 'name' — login automatizado não possível.")

    payload = {}
    hidden_inputs = form.xpath(".//input[@type='hidden']")
    for hid in hidden_inputs:
        n = hid.get("name")
        v = hid.get("value", "")
        if n:
            payload[n] = v

    payload[user_name] = USERNAME
    payload[pass_name] = PASSWORD

    post_headers = headers.copy()
    post_headers.update({
        "Referer": login_url,
        "Content-Type": "application/x-www-form-urlencoded"
    })

    resp = session.post(action_url, data=payload, headers=post_headers, allow_redirects=True, timeout=15)

    if "logout" in resp.text.lower() or "sair" in resp.text.lower() or resp.url != login_url:
        print("✅ Login bem-sucedido!")
        return session
    else:
        sys.exit("❌ Falha no login. Verifique as credenciais ou o site.")


# ----------------------------------------------------------
# FUNÇÃO: COLETAR PRODUTOS (scraping)
# ----------------------------------------------------------
def coletar_produtos(session, total_paginas=10):
    produtos = []

    for pagina in range(total_paginas):
        url_pagina = f"https://portal.lmmoto.com.br/glstorefront/glmotos/pt/BRL/Motos/Pe%C3%A7as/c/glmotos_pecas?q=%3Aname-asc&page={pagina}&sort=name-asc&layoutMode=DETAILED"
        dados = session.get(url_pagina, headers=headers, timeout=15)

        if dados.status_code != 200:
            print(f"⚠️ Erro ao acessar página {pagina}: Status {dados.status_code}")
            continue

        soup = BeautifulSoup(dados.content, 'html.parser')
        nomes = soup.find_all('div', class_='product__list--name')
        precos = soup.find_all('div', class_='precoPor')

        if not nomes or not precos:
            print(f"⚠️ Nenhum produto encontrado na página {pagina}.")
            continue

        for i in range(min(len(nomes), len(precos))):
            nome = " ".join(nomes[i].text.split())
            preco = " ".join(precos[i].text.replace("Por", "").split())
            produtos.append((nome, preco, "Portal LMMoto"))

        print(f"📦 Coletados {len(produtos)} produtos até agora...")

    return produtos


# ----------------------------------------------------------
# FUNÇÃO: NORMALIZAR STRINGS
# ----------------------------------------------------------
def normalize_str(s):
    if not isinstance(s, str):
        return ''
    s = s.strip().lower()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


# ----------------------------------------------------------
# FUNÇÃO: ENCONTRAR LINHA DE CABEÇALHO
# ----------------------------------------------------------
def encontrar_linha_cabecalho(df, chaves_desc=['descricao', 'descrição'], chaves_preco=['preco', 'preço']):
    for i in range(min(10, len(df))):
        linha = df.iloc[i].astype(str).apply(normalize_str).tolist()
        if any(chave in linha for chave in chaves_desc) and any(chave in linha for chave in chaves_preco):
            return i
    return None


# ----------------------------------------------------------
# FUNÇÃO: IMPORTAR VÁRIOS ARQUIVOS EXCEL
# ----------------------------------------------------------
def importar_varios_excel(pasta):
    arquivos_excel = glob.glob(os.path.join(pasta, "*.xls*"))
    todos_produtos = []

    if not arquivos_excel:
        print(f"⚠️ Nenhum arquivo Excel encontrado na pasta {pasta}.")
        return todos_produtos

    for arquivo in arquivos_excel:
        nome_arquivo = os.path.basename(arquivo)
        print(f"📄 Lendo arquivo: {nome_arquivo}")
        df_raw = pd.read_excel(arquivo, header=None)

        linha_cabecalho = encontrar_linha_cabecalho(df_raw)
        if linha_cabecalho is None:
            print(f"⚠️ Cabeçalho não encontrado no arquivo {arquivo}. Ignorando.")
            continue

        print(f"➡️ Cabeçalho encontrado na linha {linha_cabecalho + 1} (1-based)")
        df = pd.read_excel(arquivo, header=linha_cabecalho)

        desc_col = None
        preco_col = None
        for c in df.columns:
            c_norm = normalize_str(c)
            if c_norm in ['descricao', 'descrição', 'descriçao'] and desc_col is None:
                desc_col = c
            if c_norm in ['preco', 'preço'] and preco_col is None:
                preco_col = c

        if desc_col is None or preco_col is None:
            print(f"⚠️ Colunas 'Descrição' e 'Preço' não encontradas no arquivo {arquivo}. Ignorando.")
            continue

        for _, linha in df.iterrows():
            nome = str(linha[desc_col])
            preco = str(linha[preco_col])
            todos_produtos.append((nome, preco, nome_arquivo))

    print(f"\n✅ Total combinado: {len(todos_produtos)} produtos importados de {len(arquivos_excel)} arquivos.")
    return todos_produtos


# ----------------------------------------------------------
# FUNÇÃO: SALVAR NO BANCO (SQLite) — AGORA INTELIGENTE
# ----------------------------------------------------------
def salvar_no_banco(produtos):
    conn = sqlite3.connect("produtos.db")
    cur = conn.cursor()

    # Cria tabela se não existir
    cur.execute("""
    CREATE TABLE IF NOT EXISTS produtos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        preco TEXT NOT NULL
    )
    """)

    # Garante que a coluna 'origem' exista (caso seja banco antigo)
    try:
        cur.execute("ALTER TABLE produtos ADD COLUMN origem TEXT DEFAULT 'Desconhecida'")
        print("🧱 Coluna 'origem' adicionada à tabela existente.")
    except sqlite3.OperationalError:
        pass  # já existe

    # Insere os dados normalmente
    cur.executemany("INSERT INTO produtos (nome, preco, origem) VALUES (?, ?, ?)", produtos)
    conn.commit()
    conn.close()
    print(f"💾 {len(produtos)} produtos salvos em produtos.db!")


# ----------------------------------------------------------
# FUNÇÃO: BUSCAR NO BANCO (COM OPÇÃO DE ORDENAR)
# ----------------------------------------------------------
def buscar_produtos():
    conn = sqlite3.connect("produtos.db")
    cur = conn.cursor()

    # Pergunta ao usuário se deseja ordenar crescente ou decrescente
    while True:
        ordem = input("\n🔎 Como deseja ordenar os preços? (crescente/decrescente ou 'sair' para encerrar): ").strip().lower()
        if ordem == "sair":
            print("👋 Encerrando busca.")
            break

        if ordem not in ["crescente", "decrescente"]:
            print("⚠️ Opção inválida! Escolha 'crescente' ou 'decrescente'.")
            continue

        termo = input("Digite um termo para buscar: ").strip()

        cur.execute("SELECT nome, preco, origem FROM produtos WHERE nome LIKE ?", (f"%{termo}%",))
        resultados = cur.fetchall()

        # Função para extrair e converter preços
        def extrair_preco(p):
            if not isinstance(p, str):
                return 0
            p = p.replace("R$", "").replace(".", "").replace(",", ".")
            try:
                return float(p)
            except:
                return 0

        # Ordena os resultados com base na escolha do usuário
        resultados = sorted(resultados, key=lambda item: extrair_preco(item[1]), reverse=(ordem == "decrescente"))

        if resultados:
            print(f"\n{len(resultados)} resultados encontrados para '{termo}' (ordenados por preço {ordem}):\n")
            for nome, preco, origem in resultados:
                print(f"🛠️ {nome}\n💰 {preco}\n📦 Origem: {origem}")
                print("-" * 60)
        else:
            print("Nenhum resultado encontrado.")

    conn.close()


# ----------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# ----------------------------------------------------------
if __name__ == "__main__":
    session = login()
    produtos_online = coletar_produtos(session, total_paginas=2)

    pasta_excel = input("👉 Informe o caminho da pasta com os arquivos Excel: ").strip()
    produtos_excel = importar_varios_excel(pasta_excel)

    produtos = produtos_online + produtos_excel
    salvar_no_banco(produtos)
    buscar_produtos()