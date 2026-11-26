#!/usr/bin/env python3
"""
Monitor de promoções de BEBIDAS ALCOÓLICAS - região de Joinville/SC

Fontes: Mercado Livre, Amazon, Fort Atacadista, Angeloni e Tiendeo (configuráveis).

O que faz:
  1. Busca cada termo (cerveja, vinho, whisky...) nas lojas configuradas
  2. Extrai produto, preço, preço original e desconto
  3. Filtra para manter somente bebidas alcoólicas (e descarta versões 0,0%)
  4. Calcula o preço por litro
  5. Salva em Excel e mantém um histórico
  6. Dispara alertas (console e, opcionalmente, Telegram) conforme suas regras

Instalação:
  pip install playwright beautifulsoup4 pandas openpyxl requests
  (o robô usa o Edge/Chrome já instalado; se não houver, rode: playwright install chromium)

Uso:
  python monitor_bebidas.py
  python monitor_bebidas.py --sem-navegador   # usa só requests (sites simples)
  python monitor_bebidas.py --html pagina.html --tipo mercadolivre   # testa um HTML salvo

Agendamento diário:
  Linux/Mac (cron):   0 8 * * *  /usr/bin/python3 /caminho/monitor_bebidas.py
  Windows:            Agendador de Tarefas -> executar python monitor_bebidas.py
"""

import argparse
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------
# CONFIGURAÇÃO (edite aqui)
# ----------------------------------------------------------------------

# Pasta onde ficam o config.json e a saída (ao lado do .exe ou do script)
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
ARQ_CONFIG = BASE_DIR / "config.json"

VERSAO_CONFIG = 2

CONFIG_PADRAO = {
    "versao_config": VERSAO_CONFIG,
    "somente_alcoolicas": True,
    # Mostrar a janela do navegador ajuda quando Amazon/Mercado Livre pedem captcha
    "mostrar_navegador": False,
    # Cada termo é pesquisado nas fontes que têm "url_modelo"
    "termos_busca": ["cerveja", "vinho", "espumante", "whisky", "vodka", "gin", "cachaça", "tequila"],
    # tipo: mercadolivre | amazon | vtex | jsonld
    "fontes": [
        {"nome": "Mercado Livre", "tipo": "mercadolivre",
         "url_modelo": "https://lista.mercadolivre.com.br/{termo}"},
        {"nome": "Amazon", "tipo": "amazon",
         "url_modelo": "https://www.amazon.com.br/s?k={termo}"},
        {"nome": "Fort Atacadista (loja online)", "tipo": "vtex",
         "url_modelo": "https://www.fortatacadista.com.br/api/catalog_system/pub/products/search?ft={termo}&_from=0&_to=49"},
        {"nome": "Fort Atacadista (encartes Tiendeo)", "tipo": "jsonld",
         "url": "https://www.tiendeo.com.br/joinville/fort-atacadista"},
        {"nome": "Angeloni - Black Friday Bebidas", "tipo": "jsonld",
         "url": "https://super.angeloni.com.br/black-friday-em-bebidas"},
        {"nome": "Tiendeo - Cerveja Joinville", "tipo": "jsonld",
         "url": "https://www.tiendeo.com.br/joinville/ofertas/cerveja"},
        {"nome": "Tiendeo - Bebidas alcoólicas Joinville", "tipo": "jsonld",
         "url": "https://www.tiendeo.com.br/joinville/ofertas/bebidas-alcoolicas"},
    ],
    # Alerta: o nome contém a palavra E (preço <= max_preco OU preço/litro <= max_por_litro
    # OU desconto >= min_desconto_pct). Use null para ignorar um critério.
    "regras": [
        {"palavra": "heineken", "max_preco": 5.00,  "max_por_litro": None,  "min_desconto_pct": None},
        {"palavra": "stella",   "max_preco": None,  "max_por_litro": 16.00, "min_desconto_pct": None},
        {"palavra": "skol",     "max_preco": None,  "max_por_litro": 10.00, "min_desconto_pct": None},
        {"palavra": "brahma",   "max_preco": None,  "max_por_litro": 10.00, "min_desconto_pct": None},
        {"palavra": "vinho",    "max_preco": 30.00, "max_por_litro": None,  "min_desconto_pct": 30},
        {"palavra": "espumante","max_preco": 40.00, "max_por_litro": None,  "min_desconto_pct": 30},
        {"palavra": "whisky",   "max_preco": None,  "max_por_litro": None,  "min_desconto_pct": 25},
        {"palavra": "vodka",    "max_preco": None,  "max_por_litro": None,  "min_desconto_pct": 25},
        {"palavra": "gin",      "max_preco": None,  "max_por_litro": None,  "min_desconto_pct": 25},
    ],
    "pausa_entre_paginas_s": 4,
    "telegram_token": "",
    "telegram_chat_id": "",
}


def carregar_config() -> dict:
    """Lê config.json; se não existir (ou for de versão antiga), cria um novo."""
    if ARQ_CONFIG.exists():
        try:
            atual = json.loads(ARQ_CONFIG.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            atual = {}
        if atual.get("versao_config") == VERSAO_CONFIG:
            cfg = dict(CONFIG_PADRAO)
            cfg.update(atual)
            return cfg
        # versão antiga: guarda uma cópia e recria
        ARQ_CONFIG.rename(ARQ_CONFIG.with_name("config.antigo.json"))
        print("config.json de versão antiga salvo como config.antigo.json.")
    ARQ_CONFIG.write_text(json.dumps(CONFIG_PADRAO, indent=2, ensure_ascii=False),
                          encoding="utf-8")
    print(f"Criei {ARQ_CONFIG} com valores de exemplo. Edite-o para personalizar.\n")
    return dict(CONFIG_PADRAO)


CFG = carregar_config()
FONTES = CFG["fontes"]
REGRAS = CFG["regras"]
PASTA_SAIDA = BASE_DIR / "saida_bebidas"
PAUSA_ENTRE_PAGINAS_S = CFG["pausa_entre_paginas_s"]
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# Telegram (opcional): preencha no config.json ou use variáveis de ambiente
TELEGRAM_TOKEN = CFG.get("telegram_token") or os.getenv("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = CFG.get("telegram_chat_id") or os.getenv("TELEGRAM_CHAT_ID")


# ----------------------------------------------------------------------
# COLETA DE HTML
# ----------------------------------------------------------------------

class Navegador:
    """Abre UM navegador (Edge/Chrome) e reutiliza em todas as páginas."""

    def __init__(self, usar_navegador=True, visivel=False):
        self.usar = usar_navegador
        self.visivel = visivel
        self._pw = None
        self._browser = None
        self._ctx = None

    def __enter__(self):
        if not self.usar:
            return self
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        for canal in ("msedge", "chrome", None):   # Edge já vem no Windows
            try:
                kw = {"headless": not self.visivel}
                if canal:
                    kw["channel"] = canal
                self._browser = self._pw.chromium.launch(**kw)
                break
            except Exception:
                continue
        if self._browser is None:
            raise RuntimeError("Nenhum navegador encontrado. Instale o Microsoft Edge ou o Google Chrome.")
        self._ctx = self._browser.new_context(
            user_agent=USER_AGENT, locale="pt-BR", viewport={"width": 1366, "height": 900})
        return self

    def __exit__(self, *exc):
        try:
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.stop()
        except Exception:
            pass

    def html(self, url: str) -> str:
        if not self.usar:
            r = requests.get(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "pt-BR"}, timeout=30)
            r.raise_for_status()
            return r.text
        pagina = self._ctx.new_page()
        try:
            pagina.goto(url, wait_until="domcontentloaded", timeout=60000)
            try:
                pagina.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            for _ in range(4):                      # rola para carregar itens (lazy load)
                pagina.mouse.wheel(0, 2500)
                time.sleep(0.7)
            return pagina.content()
        finally:
            pagina.close()


def pagina_bloqueada(html: str) -> bool:
    t = html.lower()
    return any(x in t for x in ("captcha", "digite os caracteres", "robot check",
                                "sorry, we just need to make sure", "access denied"))


# ----------------------------------------------------------------------
# EXTRAÇÃO
# ----------------------------------------------------------------------

def _para_float(valor):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in s and "." in s:          # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                     # 4,99
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _percorrer(no):
    """Gera todos os dicionários dentro de uma estrutura JSON."""
    if isinstance(no, dict):
        yield no
        for v in no.values():
            yield from _percorrer(v)
    elif isinstance(no, list):
        for item in no:
            yield from _percorrer(item)


def extrair_jsonld(html: str):
    """Lê blocos <script type="application/ld+json"> e acha produto + oferta."""
    soup = BeautifulSoup(html, "html.parser")
    itens = []
    for tag in soup.find_all("script", type="application/ld+json"):
        texto = (tag.string or tag.get_text() or "").strip()
        if not texto:
            continue
        try:
            dados = json.loads(texto)
        except json.JSONDecodeError:
            continue
        for d in _percorrer(dados):
            if "offers" not in d or "name" not in d:
                continue
            ofertas = d["offers"]
            if isinstance(ofertas, dict) and isinstance(ofertas.get("offers"), list):
                ofertas = ofertas["offers"]          # AggregateOffer
            if isinstance(ofertas, dict):
                ofertas = [ofertas]
            for of in ofertas or []:
                if not isinstance(of, dict):
                    continue
                preco = _para_float(of.get("price") or of.get("lowPrice"))
                if preco is None:
                    continue
                itens.append({
                    "produto": str(d["name"]).strip(),
                    "preco": preco,
                    "preco_original": None,
                    "validade": of.get("validThrough") or of.get("priceValidUntil") or "",
                    "link": d.get("url") or of.get("url") or "",
                })
    return itens


def extrair_texto_fallback(html: str):
    """Plano B: procura padrões 'Nome ... R$ 9,99' no texto visível."""
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    linhas = [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]
    itens = []
    for i, linha in enumerate(linhas):
        m = re.search(r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
        if not m:
            continue
        nome = linhas[i - 1] if i > 0 else ""          # o nome costuma vir na linha anterior
        if len(nome) < 5 or "R$" in nome:
            continue
        itens.append({"produto": nome, "preco": _para_float(m.group(1)),
                      "preco_original": None, "validade": "", "link": ""})
    return itens


def parse_generico(html: str):
    return extrair_jsonld(html) or extrair_texto_fallback(html)


# ---- Mercado Livre ---------------------------------------------------

def _preco_ml(no):
    """Converte um bloco .andes-money-amount em float."""
    if no is None:
        return None
    fracao = no.select_one(".andes-money-amount__fraction")
    if fracao is None:
        return None
    centavos = no.select_one(".andes-money-amount__cents")
    texto = fracao.get_text(strip=True).replace(".", "")
    c = centavos.get_text(strip=True) if centavos else "00"
    try:
        return float(f"{texto}.{c}")
    except ValueError:
        return None


def parse_mercadolivre(html: str):
    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select("li.ui-search-layout__item, div.poly-card, div.ui-search-result__wrapper")
    itens = []
    vistos = set()
    for c in cards:
        titulo = c.select_one("a.poly-component__title, .poly-component__title, "
                              "h2.ui-search-item__title, .ui-search-item__title, h2, h3")
        if titulo is None:
            continue
        nome = titulo.get_text(" ", strip=True)
        precos = c.select(".andes-money-amount")
        atuais = [m for m in precos if "andes-money-amount--previous" not in (m.get("class") or [])]
        antigos = c.select(".andes-money-amount--previous")
        preco = _preco_ml(atuais[0]) if atuais else None
        if not nome or preco is None or (nome, preco) in vistos:
            continue
        vistos.add((nome, preco))
        a_tag = titulo if titulo.name == "a" else (titulo.find("a") or c.select_one("a[href]"))
        itens.append({
            "produto": nome,
            "preco": preco,
            "preco_original": _preco_ml(antigos[0]) if antigos else None,
            "validade": "",
            "link": a_tag.get("href", "") if a_tag else "",
        })
    return itens or parse_generico(html)


# ---- Amazon ----------------------------------------------------------

def parse_amazon(html: str):
    soup = BeautifulSoup(html, "html.parser")
    itens = []
    for c in soup.select('div[data-component-type="s-search-result"]'):
        h2 = c.select_one("h2")
        nome = ""
        if h2 is not None:
            nome = (h2.get("aria-label") or h2.get_text(" ", strip=True) or "").strip()
        atual = c.select_one("span.a-price:not(.a-text-price) span.a-offscreen")
        antigo = c.select_one("span.a-price.a-text-price span.a-offscreen")
        preco = _para_float(atual.get_text(strip=True)) if atual else None
        if not nome or preco is None:
            continue
        asin = c.get("data-asin")
        itens.append({
            "produto": nome,
            "preco": preco,
            "preco_original": _para_float(antigo.get_text(strip=True)) if antigo else None,
            "validade": "",
            "link": f"https://www.amazon.com.br/dp/{asin}" if asin else "",
        })
    return itens or parse_generico(html)


# ---- VTEX (API pública usada por muitas lojas; experimental para o Fort) ----

def coletar_vtex(url: str):
    r = requests.get(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"}, timeout=30)
    try:
        dados = r.json()
    except ValueError:
        raise RuntimeError("resposta não é JSON: o site pode não usar a API VTEX "
                           "(troque por uma fonte 'jsonld' ou use os encartes)")
    itens = []
    for p in dados if isinstance(dados, list) else []:
        nome_base = p.get("productName") or ""
        for it in p.get("items", []):
            nome = it.get("nameComplete") or it.get("name") or nome_base
            for sel in it.get("sellers", []):
                o = sel.get("commertialOffer", {})
                preco = o.get("Price")
                if not preco or (o.get("AvailableQuantity") is not None and o["AvailableQuantity"] <= 0):
                    continue
                lista = o.get("ListPrice")
                itens.append({
                    "produto": nome,
                    "preco": float(preco),
                    "preco_original": float(lista) if lista and lista > preco else None,
                    "validade": "",
                    "link": p.get("link", ""),
                })
                break
    return itens


PARSERS = {"mercadolivre": parse_mercadolivre, "amazon": parse_amazon, "jsonld": parse_generico}


# ----------------------------------------------------------------------
# FILTRO: SOMENTE BEBIDAS ALCOÓLICAS
# ----------------------------------------------------------------------

def sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(texto))
                   if unicodedata.category(c) != "Mn").lower()


_ALCOOL = [
    "cerveja", "cervejas", "chope", "chopp", "vinho", "vinhos", "espumante", "espumantes",
    "champagne", "champanhe", "prosecco", "vodka", "vodca", "whisky", "whiskey", "uisque",
    "gin", "rum", "cachaca", "pinga", "tequila", "licor", "conhaque", "brandy", "sake",
    "saque", "aperol", "campari", "vermute", "absinto", "sidra", "hidromel", "fernet",
    "corote", "batida", "caipirinha", "cabernet", "merlot", "malbec", "chardonnay",
    "sauvignon", "carmenere", "tannat", "pinot",
    # marcas inequívocas
    "heineken", "budweiser", "brahma", "skol", "stella artois", "eisenbahn", "becks",
    "amstel", "corona extra", "spaten", "devassa", "itaipava", "bohemia", "jack daniel", "johnnie walker",
    "absolut", "smirnoff", "tanqueray", "bombay", "jose cuervo", "ypioca", "velho barreiro",
    "chandon", "salton", "miolo", "casillero", "concha y toro",
]
_EXCLUIR = [
    r"sem alcool", r"zero alcool", r"nao alcoolic", r"s/\s*alcool", r"\balcool\s*0",
    r"(?<![\d.,])0[.,]0", r"\bzero\b", r"malzbier", r"guarana", r"vinagre", r"\bmolho\b",
    r"chocolate", r"bombom", r"gelatina", r"sorvete", r"geleia", r"tempero", r"\bpao\b",
    r"\bcopos?\b", r"\btacas?\b", r"abridor", r"saca[\s-]?rolhas?", r"adega climatizada",
    r"caixa termica", r"\bbalde\b", r"decanter", r"\bcaneca", r"camiseta", r"chaveiro",
    r"\bsuco\b", r"refrigerante", r"energetico", r"shampoo", r"sabonete", r"racao",
    r"\bessencia\b", r"\baroma\b", r"luminaria", r"\bquadro\b", r"\bkit de\b.*\bcopos?\b",
]
RE_ALCOOL = re.compile(r"\b(?:" + "|".join(re.escape(p) for p in _ALCOOL) + r")\b")
RE_EXCLUIR = re.compile("|".join(_EXCLUIR))


def eh_alcoolica(nome: str) -> bool:
    n = sem_acento(nome)
    if RE_EXCLUIR.search(n):
        return False
    return bool(RE_ALCOOL.search(n))


# ----------------------------------------------------------------------
# VOLUME / COLETA
# ----------------------------------------------------------------------

def volume_em_litros(nome: str):
    """Descobre o volume total (em litros) a partir do nome do produto."""
    n = nome.lower().replace(",", ".")
    m = re.search(r"(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(ml|l|lt|litros?)\b", n)    # 12 x 350ml
    if m:
        qtd, vol, un = int(m.group(1)), float(m.group(2)), m.group(3)
        return qtd * (vol / 1000 if un == "ml" else vol)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(ml|l|lt|litros?)\b", n)                # 350ml, 1.5l
    if m:
        vol, un = float(m.group(1)), m.group(2)
        litros = vol / 1000 if un == "ml" else vol
        q = re.search(r"(?:fardo|pack|caixa|c/|com|kit)\s*(\d+)\s*(?:un|unid|latas?|garrafas?)?", n)
        if q and int(q.group(1)) > 1 and litros < 1:
            litros *= int(q.group(1))
        return litros
    return None


def montar_tarefas(termos):
    """Expande as fontes em uma lista de (fonte, termo, url)."""
    tarefas = []
    for f in FONTES:
        tipo = f.get("tipo", "jsonld")
        if "url_modelo" in f:
            for termo in termos:
                if tipo == "mercadolivre":
                    t = sem_acento(termo).strip().replace(" ", "-")
                else:
                    t = quote_plus(termo)
                tarefas.append((f, termo, f["url_modelo"].format(termo=t)))
        elif "url" in f:
            tarefas.append((f, "", f["url"]))
    return tarefas


def coletar(fonte: dict, termo: str, url: str, nav: Navegador):
    rotulo = f"{fonte['nome']}" + (f" [{termo}]" if termo else "")
    print(f"-> {rotulo}")
    tipo = fonte.get("tipo", "jsonld")
    try:
        if tipo == "vtex":
            itens = coletar_vtex(url)
        else:
            html = nav.html(url)
            if pagina_bloqueada(html) and not PARSERS[tipo](html):
                print("   BLOQUEADO (captcha/robô). Tente \"mostrar_navegador\": true no config.json.")
                return []
            itens = PARSERS.get(tipo, parse_generico)(html)
    except Exception as e:
        print(f"   ERRO: {e}")
        return []
    print(f"   {len(itens)} itens")

    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    for it in itens:
        it["fonte"] = fonte["nome"]
        it["termo"] = termo
        it["coletado_em"] = agora
        if it.get("link", "").startswith("/"):
            base = re.match(r"https?://[^/]+", url)
            it["link"] = (base.group(0) if base else "") + it["link"]
        litros = volume_em_litros(it["produto"])
        it["litros"] = litros
        it["preco_por_litro"] = round(it["preco"] / litros, 2) if litros else None
        po = it.get("preco_original")
        it["desconto_pct"] = round((1 - it["preco"] / po) * 100, 1) if po and po > it["preco"] else None
    return itens


# ----------------------------------------------------------------------
# ALERTAS
# ----------------------------------------------------------------------

def avaliar_alertas(df: pd.DataFrame) -> pd.DataFrame:
    achados = []
    nomes = df["produto"].map(sem_acento)
    for regra in REGRAS:
        sub = df[nomes.str.contains(re.escape(sem_acento(regra["palavra"])), na=False)]
        for _, l in sub.iterrows():
            ok_preco = regra.get("max_preco") is not None and l["preco"] <= regra["max_preco"]
            ok_litro = (regra.get("max_por_litro") is not None and pd.notna(l["preco_por_litro"])
                        and l["preco_por_litro"] <= regra["max_por_litro"])
            ok_desc = (regra.get("min_desconto_pct") is not None and pd.notna(l["desconto_pct"])
                       and l["desconto_pct"] >= regra["min_desconto_pct"])
            if ok_preco or ok_litro or ok_desc:
                achados.append(l)
    if not achados:
        return pd.DataFrame()
    return pd.DataFrame(achados).drop_duplicates(subset=["produto", "preco", "fonte"])


def enviar_telegram(texto: str):
    if not (TELEGRAM_TOKEN and TELEGRAM_CHAT_ID):
        return
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={"chat_id": TELEGRAM_CHAT_ID, "text": texto[:4000]},
            timeout=20,
        )
    except Exception as e:
        print(f"Falha ao enviar Telegram: {e}")


# ----------------------------------------------------------------------
# PRINCIPAL
# ----------------------------------------------------------------------

COLUNAS = ["fonte", "termo", "produto", "preco", "preco_original", "desconto_pct",
           "litros", "preco_por_litro", "validade", "link", "coletado_em"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-navegador", action="store_true",
                    help="usa apenas requests (mais rápido, mas falha em sites com JavaScript)")
    ap.add_argument("--sem-pausa", action="store_true",
                    help="não espera Enter no final (use no agendador de tarefas)")
    ap.add_argument("--html", help="testa a extração em um arquivo HTML salvo")
    ap.add_argument("--tipo", default="jsonld", help="tipo do parser para --html")
    args = ap.parse_args()

    if args.html:
        html = Path(args.html).read_text(encoding="utf-8", errors="ignore")
        for it in PARSERS[args.tipo](html)[:30]:
            print(f"{it['preco']:>9.2f}  {'OK ' if eh_alcoolica(it['produto']) else '-- '}{it['produto']}")
        return

    PASTA_SAIDA.mkdir(exist_ok=True)
    termos = CFG["termos_busca"]
    todos = []
    with Navegador(not args.sem_navegador, CFG["mostrar_navegador"]) as nav:
        for fonte, termo, url in montar_tarefas(termos):
            todos.extend(coletar(fonte, termo, url, nav))
            time.sleep(PAUSA_ENTRE_PAGINAS_S)

    if not todos:
        print("Nenhum item coletado. Verifique conexão, URLs ou bloqueios.")
        sys.exit(1)

    df = pd.DataFrame(todos).drop_duplicates(subset=["produto", "preco", "fonte"])
    total = len(df)
    if CFG["somente_alcoolicas"]:
        df = df[df["produto"].map(eh_alcoolica)]
    df = df.sort_values(["preco_por_litro", "preco"], na_position="last")
    df = df[[c for c in COLUNAS if c in df.columns]]

    print(f"\n{total} itens coletados, {len(df)} após o filtro de bebidas alcoólicas.")
    print(df.groupby("fonte").size().to_string())

    if df.empty:
        print("Nada restou após o filtro.")
        return

    alertas = avaliar_alertas(df)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M")
    xlsx = PASTA_SAIDA / f"promocoes_{carimbo}.xlsx"
    with pd.ExcelWriter(xlsx) as w:
        df.to_excel(w, sheet_name="Promocoes", index=False)
        if not alertas.empty:
            alertas.to_excel(w, sheet_name="Alertas", index=False)

    hist = PASTA_SAIDA / "historico.csv"
    df.to_csv(hist, mode="a", header=not hist.exists(), index=False, encoding="utf-8-sig")
    print(f"Planilha salva em {xlsx}")

    if alertas.empty:
        print("Nenhum alerta pelas suas regras hoje.")
        return

    linhas = ["ALERTAS DE PROMOÇÃO - bebidas alcoólicas"]
    for _, a in alertas.iterrows():
        pl = f" (R$ {a['preco_por_litro']:.2f}/L)" if pd.notna(a["preco_por_litro"]) else ""
        dc = f" -{a['desconto_pct']:.0f}%" if pd.notna(a["desconto_pct"]) else ""
        linhas.append(f"• {a['produto']}: R$ {a['preco']:.2f}{dc}{pl} [{a['fonte']}]")
    mensagem = "\n".join(linhas)
    print("\n" + mensagem)
    enviar_telegram(mensagem)


if __name__ == "__main__":
    try:
        main()
    finally:
        # com duplo clique no .exe, mantém a janela aberta para você ler o resultado
        if getattr(sys, "frozen", False) and sys.stdin and sys.stdin.isatty() and "--sem-pausa" not in sys.argv:
            input("\nPressione Enter para fechar...")
