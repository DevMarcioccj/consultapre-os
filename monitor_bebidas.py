#!/usr/bin/env python3
"""
Monitor de promoções de BEBIDAS ALCOÓLICAS - região de Joinville/SC

Fontes ativas: Amazon, Angeloni e Tiendeo (Joinville). Novas fontes podem ser adicionadas no config.json.

O que faz:
  1. Busca cada termo (cerveja, vinho, whisky...) nas lojas configuradas
  2. (Mercado Livre e Fort estão desativados por enquanto)
  3. Extrai produto, preço, preço original e desconto
  4. Mantém somente bebidas alcoólicas e classifica o TIPO (cerveja, vinho, whisky...)
  5. Calcula o preço por litro
  6. Salva em Excel (com filtros ativados) e mantém um histórico
  7. Dispara alertas (console e, opcionalmente, Telegram) conforme suas regras

Instalação:
  pip install playwright beautifulsoup4 pandas openpyxl requests
  (pdfplumber só é necessário se reativar a leitura dos PDFs do Fort)
  (o robô usa o Edge/Chrome já instalado; se não houver, rode: playwright install chromium)

Uso:
  python monitor_bebidas.py
  python monitor_bebidas.py --categorias cerveja          # só cervejas (aceita várias: cerveja,vinho)
  python monitor_bebidas.py --sem-navegador               # usa só requests
  python monitor_bebidas.py --html pagina.html --parser mercadolivre   # testa um HTML salvo
  python monitor_bebidas.py --pdf encarte.pdf             # testa a leitura de um PDF do Fort
  python monitor_bebidas.py --login                       # entra uma vez na sua conta do Mercado Livre

Agendamento diário:
  Linux/Mac (cron):   0 8 * * *  /usr/bin/python3 /caminho/monitor_bebidas.py --sem-pausa
  Windows:            Agendador de Tarefas -> executar monitor_bebidas.exe --sem-pausa
"""

import argparse
import io
import json
import os
import re
import sys
import time
import unicodedata
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------
# CONFIGURAÇÃO (edite o config.json que o programa cria ao lado dele)
# ----------------------------------------------------------------------

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
ARQ_CONFIG = BASE_DIR / "config.json"
VERSAO_CONFIG = 4

CONFIG_PADRAO = {
    "versao_config": VERSAO_CONFIG,
    "somente_alcoolicas": True,
    # Categorias a manter na planilha. Vazio = todas. Ex.: ["cerveja"] ou ["cerveja", "vinho"]
    # Categorias possíveis: cerveja, vinho, espumante, whisky, vodka, gin, cachaça, tequila,
    # rum, licor, aperitivo, conhaque, saquê, drink pronto, outros
    "tipos_desejados": [],
    # false = navegador escondido. Use true para ver a janela (e resolver um captcha, se
    # a Amazon pedir).
    "mostrar_navegador": False,
    # Se a página vier vazia/bloqueada, espera até N segundos (com a janela visível)
    "aguardar_verificacao_s": 90,
    # Salva o HTML de páginas que retornaram 0 itens em saida_bebidas/debug/
    "salvar_html_debug": True,
    "termos_busca": ["cerveja", "vinho", "espumante", "whisky", "vodka", "gin", "cachaça", "tequila"],
    # tipo: amazon | jsonld  (os tipos mercadolivre e fort_pdf ainda existem no código,
    # mas estão desativados por padrão; veja o histórico da conversa se quiser reativar)
    "fontes": [
        {"nome": "Amazon", "tipo": "amazon",
         "url_modelo": "https://www.amazon.com.br/s?k={termo}"},
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
        {"palavra": "espumante", "max_preco": 40.00, "max_por_litro": None, "min_desconto_pct": 30},
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
        ARQ_CONFIG.replace(ARQ_CONFIG.with_name("config.antigo.json"))
        print("config.json de versão antiga salvo como config.antigo.json.")
    ARQ_CONFIG.write_text(json.dumps(CONFIG_PADRAO, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Criei {ARQ_CONFIG} com valores de exemplo. Edite-o para personalizar.\n")
    return dict(CONFIG_PADRAO)


CFG = carregar_config()
FONTES = CFG["fontes"]
REGRAS = CFG["regras"]
PASTA_SAIDA = BASE_DIR / "saida_bebidas"
PERFIL_NAVEGADOR = BASE_DIR / "perfil_navegador"
PAUSA_ENTRE_PAGINAS_S = CFG["pausa_entre_paginas_s"]
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
TELEGRAM_TOKEN = CFG.get("telegram_token") or os.getenv("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = CFG.get("telegram_chat_id") or os.getenv("TELEGRAM_CHAT_ID")


# ----------------------------------------------------------------------
# TEXTO / CLASSIFICAÇÃO
# ----------------------------------------------------------------------

def sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(texto))
                   if unicodedata.category(c) != "Mn").lower()


_ALCOOL = [
    "cerveja", "cervejas", "chope", "chopp", "vinho", "vinhos", "espumante", "espumantes",
    "champagne", "champanhe", "prosecco", "vodka", "vodca", "whisky", "whiskey", "uisque",
    "bourbon", "gin", "rum", "cachaca", "pinga", "tequila", "licor", "conhaque", "brandy",
    "sake", "saque", "aperol", "campari", "vermute", "absinto", "sidra", "hidromel", "fernet",
    "aperitivo", "bitter", "corote", "batida", "caipirinha", "cabernet", "merlot", "malbec",
    "chardonnay", "sauvignon", "carmenere", "tannat", "pinot",
    # marcas inequívocas
    "heineken", "budweiser", "brahma", "skol", "stella artois", "eisenbahn", "becks",
    "amstel", "corona extra", "spaten", "devassa", "itaipava", "bohemia", "jack daniel",
    "johnnie walker", "absolut", "smirnoff", "tanqueray", "bombay", "beefeater", "jose cuervo",
    "jimador", "ypioca", "velho barreiro", "chandon", "salton", "miolo", "casillero",
    "concha y toro", "orloff", "skyy", "ballantine", "chivas", "buchanan", "white horse",
    "jim beam", "bacardi", "amarula", "baileys", "jagermeister", "flying fish", "skol beats",
]
_EXCLUIR = [
    r"sem alcool", r"zero alcool", r"nao alcoolic", r"s/\s*alcool", r"\balcool\s*0",
    r"(?<![\d.,])0[.,]0", r"\bzero\b", r"malzbier", r"guarana", r"vinagre", r"\bmolho\b",
    r"chocolate", r"bombom", r"gelatina", r"sorvete", r"geleia", r"tempero", r"\bpao\b",
    r"\bcopos?\b", r"\btacas?\b", r"abridor", r"saca[\s-]?rolhas?", r"adega climatizada",
    r"caixa termica", r"\bbalde\b", r"decanter", r"\bcaneca", r"camiseta", r"chaveiro",
    r"\bsuco\b", r"refrigerante", r"energetico", r"shampoo", r"sabonete", r"racao",
    r"\bessencia\b", r"\baroma\b", r"luminaria", r"\bquadro\b",
]
RE_ALCOOL = re.compile(r"\b(?:" + "|".join(re.escape(p) for p in _ALCOOL) + r")\b")
RE_EXCLUIR = re.compile("|".join(_EXCLUIR))


def eh_alcoolica(nome: str) -> bool:
    n = sem_acento(nome)
    return not RE_EXCLUIR.search(n) and bool(RE_ALCOOL.search(n))


# Ordem importa: o primeiro tipo que casar vence.
_TIPOS = [
    ("drink pronto", ["corote", "skol beats", "smirnoff ice", "flying fish", "ice red", "batida",
                      "caipirinha", "drink"]),
    ("espumante", ["espumante", "champagne", "champanhe", "prosecco", "chandon"]),
    ("vinho", ["vinho", "cabernet", "merlot", "malbec", "chardonnay", "sauvignon", "carmenere",
               "tannat", "pinot", "casillero", "concha y toro", "miolo", "quinta do"]),
    ("whisky", ["whisky", "whiskey", "uisque", "bourbon", "jack daniel", "johnnie walker",
                "ballantine", "chivas", "buchanan", "passport", "white horse", "jim beam",
                "black & white"]),
    ("vodka", ["vodka", "vodca", "absolut", "orloff", "skyy", "smirnoff", "volk"]),
    ("gin", ["gin", "tanqueray", "bombay", "beefeater", "gordon"]),
    ("cachaça", ["cachaca", "pinga", "ypioca", "velho barreiro"]),
    ("tequila", ["tequila", "jose cuervo", "jimador"]),
    ("rum", ["rum", "bacardi"]),
    ("licor", ["licor", "amarula", "baileys", "jagermeister"]),
    ("aperitivo", ["aperol", "campari", "aperitivo", "bitter", "vermute", "fernet", "absinto"]),
    ("conhaque", ["conhaque", "brandy"]),
    ("saquê", ["saque", "sake"]),
    ("cerveja", ["cerveja", "cervejas", "chope", "chopp", "heineken", "budweiser", "brahma",
                 "skol", "stella artois", "eisenbahn", "becks", "amstel", "corona", "spaten",
                 "devassa", "itaipava", "bohemia", "ipa", "pilsen", "lager", "puro malte",
                 "long neck"]),
]
_TIPOS_RE = [(t, re.compile(r"\b(?:" + "|".join(re.escape(p) for p in ps) + r")\b"))
             for t, ps in _TIPOS]
NOMES_TIPOS = [t for t, _ in _TIPOS] + ["outros"]


def classificar(nome: str, termo: str = "") -> str:
    """Devolve o tipo da bebida (cerveja, vinho, whisky...)."""
    n = sem_acento(nome)
    for tipo, rx in _TIPOS_RE:
        if rx.search(n):
            return tipo
    t = sem_acento(termo)
    for tipo, _ in _TIPOS:                      # usa o termo pesquisado como último recurso
        if sem_acento(tipo) == t:
            return tipo
    return "outros"


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


# ----------------------------------------------------------------------
# NAVEGADOR
# ----------------------------------------------------------------------

class Navegador:
    """Abre UM navegador (Edge/Chrome) com perfil persistente e reutiliza nas páginas."""

    def __init__(self, usar_navegador=True, visivel=True, espera_s=90):
        self.usar = usar_navegador
        self.visivel = visivel
        self.espera_s = espera_s
        self._pw = None
        self._ctx = None
        self.ultima_url = ""

    def __enter__(self):
        if not self.usar:
            return self
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        PERFIL_NAVEGADOR.mkdir(exist_ok=True)
        for canal in ("msedge", "chrome", None):          # Edge já vem no Windows
            try:
                kw = dict(user_data_dir=str(PERFIL_NAVEGADOR), headless=not self.visivel,
                          locale="pt-BR", viewport={"width": 1366, "height": 900})
                if canal:
                    kw["channel"] = canal
                if not self.visivel:
                    kw["user_agent"] = USER_AGENT         # evita "HeadlessChrome" no UA
                self._ctx = self._pw.chromium.launch_persistent_context(**kw)
                break
            except Exception:
                continue
        if self._ctx is None:
            raise RuntimeError("Nenhum navegador encontrado. Instale o Microsoft Edge ou o Google Chrome.")
        return self

    def __exit__(self, *exc):
        try:
            if self._ctx:
                self._ctx.close()
            if self._pw:
                self._pw.stop()
        except Exception:
            pass

    def html(self, url: str, validar=None) -> str:
        """Baixa a página. Se `validar(html)` falhar e a janela estiver visível,
        espera você resolver captcha/login (até espera_s segundos)."""
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
            html = pagina.content()
            self.ultima_url = pagina.url
            if validar and not validar(html) and self.visivel and self.espera_s > 0:
                print(f"   sem resultados; aguardando até {self.espera_s}s "
                      f"(resolva captcha/login na janela, se aparecer)...")
                fim = time.time() + self.espera_s
                while time.time() < fim:
                    time.sleep(3)
                    try:
                        html = pagina.content()
                        self.ultima_url = pagina.url
                    except Exception:
                        continue
                    if validar(html):
                        break
            return html
        finally:
            pagina.close()

    def baixar(self, url: str) -> bytes:
        """Baixa um arquivo (PDF/JSON). Tenta requests e, se falhar, o contexto do navegador."""
        try:
            r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=60)
            r.raise_for_status()
            return r.content
        except Exception:
            if not self.usar:
                raise
            resp = self._ctx.request.get(url, timeout=60000)
            if not resp.ok:
                raise RuntimeError(f"HTTP {resp.status} em {url}")
            return resp.body()


def pagina_bloqueada(html: str) -> bool:
    t = html.lower()
    return any(x in t for x in ("captcha", "digite os caracteres", "robot check",
                                "account-verification", "access denied"))


# ----------------------------------------------------------------------
# EXTRAÇÃO DE PÁGINAS HTML
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
                    "produto": str(d["name"]).strip(), "preco": preco, "preco_original": None,
                    "validade": of.get("validThrough") or of.get("priceValidUntil") or "",
                    "link": d.get("url") or of.get("url") or "",
                })
    return itens


_LIXO_NOME = re.compile(r"^(r\$|frete|mais vendido|oferta|ultima|desconto|\d|ver |comprar|\W)", re.I)


def extrair_texto_fallback(html: str):
    """Plano B: procura 'R$ 9,99' no texto e usa a linha de nome mais próxima acima."""
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    linhas = [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]
    itens = []
    for i, linha in enumerate(linhas):
        m = re.search(r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
        if not m:
            continue
        nome = ""
        for j in range(i - 1, max(i - 5, -1), -1):       # olha até 4 linhas acima
            cand = linhas[j]
            if len(cand) >= 15 and "R$" not in cand and not _LIXO_NOME.match(cand):
                nome = cand
                break
        if nome:
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


def _ml_por_cards(soup):
    itens, vistos = [], set()
    cards = soup.select("li.ui-search-layout__item, div.poly-card, div.ui-search-result__wrapper")
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
        itens.append({"produto": nome, "preco": preco,
                      "preco_original": _preco_ml(antigos[0]) if antigos else None,
                      "validade": "", "link": a_tag.get("href", "") if a_tag else ""})
    return itens


def _ml_por_links(soup):
    """Independe de classes CSS: acha links de anúncio (MLB...) e o preço no mesmo bloco."""
    itens, vistos = [], set()
    for a in soup.select("a[href]"):
        href = a.get("href", "")
        if not re.search(r"MLB-?\d{6,}", href):
            continue
        nome = (a.get("title") or a.get_text(" ", strip=True) or "").strip()
        if len(nome) < 12:
            continue
        bloco = a
        for _ in range(8):                      # sobe até achar um ancestral com preço
            bloco = bloco.parent
            if bloco is None or "R$" in bloco.get_text(" ", strip=True)[:600]:
                break
        if bloco is None:
            continue
        atuais = [m for m in bloco.select(".andes-money-amount")
                  if "andes-money-amount--previous" not in (m.get("class") or [])]
        preco = _preco_ml(atuais[0]) if atuais else None
        if preco is None:
            m = re.search(r"R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)", bloco.get_text(" ", strip=True))
            preco = _para_float(m.group(1)) if m else None
        if preco is None or (nome, preco) in vistos:
            continue
        vistos.add((nome, preco))
        antigos = bloco.select(".andes-money-amount--previous")
        itens.append({"produto": nome, "preco": preco,
                      "preco_original": _preco_ml(antigos[0]) if antigos else None,
                      "validade": "", "link": href})
    return itens


def parse_mercadolivre(html: str):
    soup = BeautifulSoup(html, "html.parser")
    return _ml_por_cards(soup) or _ml_por_links(soup) or parse_generico(html)


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
        itens.append({"produto": nome, "preco": preco,
                      "preco_original": _para_float(antigo.get_text(strip=True)) if antigo else None,
                      "validade": "",
                      "link": f"https://www.amazon.com.br/dp/{asin}" if asin else ""})
    return itens or parse_generico(html)


PARSERS = {"mercadolivre": parse_mercadolivre, "amazon": parse_amazon, "jsonld": parse_generico}


# ----------------------------------------------------------------------
# FORT ATACADISTA: ENCARTES EM PDF
# ----------------------------------------------------------------------
# O Fort não tem loja online com busca: as ofertas saem em PDFs semanais por estado
# (ex.: SC_Fort_Especial-Bebidas_12-a-14_JUN_26.pdf). Aqui listamos, baixamos e lemos esses PDFs.

_RE_PRECO_ANCORA = re.compile(r"^R\$\^?(\d{0,3})$")
_RE_NUM = re.compile(r"^\^?(\d{1,3})$")
_RE_CENTAVOS = re.compile(r"^[,.]?\^?(\d{2})$")


def _segmentos(palavras, ignorar_ids):
    """Agrupa palavras em segmentos de texto (mesma linha e sem grandes vãos horizontais)."""
    uteis = [w for i, w in enumerate(palavras) if i not in ignorar_ids]
    uteis.sort(key=lambda w: (round(w["top"] / 3), w["x0"]))
    linhas, atual = [], []
    for w in uteis:
        if atual and abs(w["top"] - atual[0]["top"]) > 3:
            linhas.append(atual)
            atual = []
        atual.append(w)
    if atual:
        linhas.append(atual)
    segs = []
    for ln in linhas:
        ln.sort(key=lambda w: w["x0"])
        grupo = [ln[0]]
        for w in ln[1:]:
            if w["x0"] - grupo[-1]["x1"] > 14:            # vão grande = outra coluna
                segs.append(grupo)
                grupo = []
            grupo.append(w)
        segs.append(grupo)
    saida = []
    for g in segs:
        texto = " ".join(w["text"] for w in g).strip()
        saida.append({"x0": g[0]["x0"], "x1": g[-1]["x1"], "top": min(w["top"] for w in g),
                      "bottom": max(w["bottom"] for w in g), "texto": texto})
    return saida


_RE_SEG_LIXO = re.compile(r"^(?:[\d,.^\s]+|un\.?|u n\.?|R\$.*|de:?.*R\$.*|\W+)$", re.I)


def _nome_proximo(ancora, segs, abaixo=True):
    cx = (ancora["x0"] + ancora["x1"]) / 2
    if abaixo:
        cand = [s for s in segs if s["top"] >= ancora["bottom"] - 3 and s["top"] <= ancora["bottom"] + 95]
        cand.sort(key=lambda s: s["top"])
    else:
        cand = [s for s in segs if s["bottom"] <= ancora["top"] + 3 and s["bottom"] >= ancora["top"] - 95]
        cand.sort(key=lambda s: -s["bottom"])
    cand = [s for s in cand if abs((s["x0"] + s["x1"]) / 2 - cx) <= 75]
    linhas, ultimo = [], None
    for s in cand:
        if "R$" in s["texto"]:
            break                                           # chegou em outro preço
        if _RE_SEG_LIXO.match(s["texto"]):
            continue
        if ultimo is not None:
            gap = (s["top"] - ultimo["bottom"]) if abaixo else (ultimo["top"] - s["bottom"])
            if gap > 8:
                break
        linhas.append(s["texto"])
        ultimo = s
        if len(linhas) == 4:
            break
    if not abaixo:
        linhas.reverse()
    return " ".join(linhas).strip()


def _itens_da_pagina(pagina):
    palavras = pagina.extract_words(extra_attrs=["size"], x_tolerance=2, y_tolerance=2)
    usadas, ancoras = set(), []
    for i, w in enumerate(palavras):
        m = _RE_PRECO_ANCORA.match(w["text"])
        if not m:
            continue
        inteiro, cents, tam = None, None, w["size"]
        usadas.add(i)
        resto = []
        if m.group(1):
            inteiro = int(m.group(1))
        # palavras à direita, na mesma linha (sobreposição vertical), em até 60pt
        for j, v in enumerate(palavras):
            if j == i or v["x0"] < w["x1"] - 2 or v["x0"] - w["x1"] > 60:
                continue
            if v["top"] > w["bottom"] + 4 or v["bottom"] < w["top"] - 4:
                continue
            resto.append((j, v))
        resto.sort(key=lambda t: t[1]["x0"])
        topo, fundo = w["top"], w["bottom"]
        viu_virgula = False
        for j, v in resto:
            if inteiro is None and _RE_NUM.match(v["text"]):
                inteiro, tam = int(_RE_NUM.match(v["text"]).group(1)), v["size"]
                usadas.add(j)
                topo, fundo = min(topo, v["top"]), max(fundo, v["bottom"])
            elif inteiro is not None and cents is None:
                t = v["text"].replace("^", "")
                if t in (",", "."):
                    usadas.add(j)
                    viu_virgula = True
                    continue
                mc = _RE_CENTAVOS.match(v["text"])
                if mc and (viu_virgula or "," in v["text"] or "^" in v["text"] or v["size"] < tam * 0.85):
                    cents = int(mc.group(1))
                    usadas.add(j)
                    topo, fundo = min(topo, v["top"]), max(fundo, v["bottom"])
        if inteiro is None:
            continue
        ancora = {"x0": w["x0"], "x1": max([w["x1"]] + [v["x1"] for _, v in resto]),
                  "top": topo, "bottom": fundo, "inteiro": inteiro, "cents": cents,
                  "tam": tam, "i": i}
        ancoras.append(ancora)

    # centavos que ficam abaixo/ao lado do inteiro (layout de tabela)
    for a in ancoras:
        if a["cents"] is not None:
            continue
        melhor, dmin = None, 1e9
        for j, v in enumerate(palavras):
            if j in usadas:
                continue
            mc = _RE_CENTAVOS.match(v["text"])
            if not mc:
                continue
            if not (a["x0"] - 5 <= v["x0"] <= a["x1"] + 50 and a["top"] - 10 <= v["top"] <= a["bottom"] + 45):
                continue
            if not ("," in v["text"] or "^" in v["text"] or v["size"] < a["tam"] * 0.85):
                continue
            d = abs(v["top"] - a["top"]) + abs(v["x0"] - a["x1"])
            if d < dmin:
                melhor, dmin = (j, int(mc.group(1))), d
        if melhor:
            a["cents"] = melhor[1]
            usadas.add(melhor[0])

    segs = _segmentos(palavras, usadas)
    itens = []
    antigos = [a for a in ancoras if any(
        s["texto"].lower().startswith("de") and abs(s["top"] - a["top"]) < 6 and 0 < a["x0"] - s["x1"] < 25
        for s in segs)]
    principais = [a for a in ancoras if a not in antigos]
    for a in principais:
        preco = a["inteiro"] + (a["cents"] or 0) / 100
        nome = _nome_proximo(a, segs, abaixo=True)
        if not eh_alcoolica(nome):
            nome_acima = _nome_proximo(a, segs, abaixo=False)
            if eh_alcoolica(nome_acima):
                nome = nome_acima
        if not nome or not eh_alcoolica(nome):
            continue
        original = None
        for o in antigos:
            if abs(o["x0"] - a["x0"]) < 60 and abs(o["top"] - a["top"]) < 50:
                po = o["inteiro"] + (o["cents"] or 0) / 100
                if po > preco:
                    original = po
        itens.append({"produto": nome, "preco": preco, "preco_original": original,
                      "validade": "", "link": "",
                      "obs": "" if a["cents"] is not None else "centavos não identificados"})
    return itens


def _data_fim_validade(texto: str):
    t = texto.lower()
    m = re.search(r"validade\s+de\s+(\d{2})/(\d{2})/(\d{4})\s+at[eé]\s+(\d{2})/(\d{2})/(\d{4})", t)
    try:
        if m:
            return date(int(m.group(6)), int(m.group(5)), int(m.group(4)))
        m = re.search(r"v[aá]lid[ao]s?\s+(?:de\s+)?(\d{1,2})(?:/(\d{1,2}))?(?:/(\d{2,4}))?\s*"
                      r"(?:a|at[eé])\s*(\d{1,2})/(\d{1,2})/(\d{2,4})", t)
        if m:
            ano = int(m.group(6))
            ano += 2000 if ano < 100 else 0
            return date(ano, int(m.group(5)), int(m.group(4)))
    except ValueError:
        pass
    return None


def parse_pdf_fort(conteudo: bytes):
    """Lê um encarte em PDF. Devolve (itens, data_fim_da_validade)."""
    import pdfplumber
    itens, textos = [], []
    with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
        for pagina in pdf.pages:
            textos.append(pagina.extract_text() or "")
            itens.extend(_itens_da_pagina(pagina))
    return itens, _data_fim_validade("\n".join(textos))


def listar_pdfs_fort(fonte: dict, nav: Navegador):
    base = fonte.get("base_url", "https://www.fortatacadista.com.br").rstrip("/")
    uf = fonte.get("uf", "SC").upper()
    achados = []
    try:                                             # 1) API de mídia do WordPress
        api = (f"{base}/wp-json/wp/v2/media?mime_type=application/pdf&per_page=100"
               f"&orderby=date&order=desc&_fields=source_url,date")
        for d in json.loads(nav.baixar(api).decode("utf-8")):
            achados.append(d.get("source_url", ""))
    except Exception as e:
        print(f"   API de mídia indisponível ({str(e)[:60]}); procurando links nas páginas do site")
    if not any(os.path.basename(u).upper().startswith(uf + "_") for u in achados):
        for caminho in ("/", "/ofertas/", "/encartes/", "/tabloides/", "/ofertas-da-semana/"):
            try:
                html = nav.html(base + caminho)
                achados += re.findall(r"https?://[^\"'\s<>]+\.pdf", html)
            except Exception:
                continue
    vistos, urls = set(), []
    for u in achados + list(fonte.get("pdfs_extras", [])):
        if u and u not in vistos and (os.path.basename(u).upper().startswith(uf + "_")
                                      or u in fonte.get("pdfs_extras", [])):
            vistos.add(u)
            urls.append(u)
    return urls[: int(fonte.get("max_pdfs", 8))]


def coletar_fort_pdf(fonte: dict, nav: Navegador):
    urls = listar_pdfs_fort(fonte, nav)
    if not urls:
        print("   nenhum PDF encontrado. Informe links manualmente em \"pdfs_extras\" no config.json.")
        return []
    hoje, todos = date.today(), []
    for u in urls:
        nome = os.path.basename(u)
        try:
            itens, fim = parse_pdf_fort(nav.baixar(u))
        except Exception as e:
            print(f"   {nome}: erro ({str(e)[:80]})")
            continue
        if fim and fim < hoje:
            print(f"   {nome}: vencido em {fim:%d/%m/%Y} (ignorado)")
            continue
        print(f"   {nome}: {len(itens)} bebidas alcoólicas"
              + (f", válido até {fim:%d/%m/%Y}" if fim else ""))
        for it in itens:
            it["validade"] = f"até {fim:%d/%m/%Y}" if fim else ""
            it["link"] = u
        todos.extend(itens)
    return todos


# ----------------------------------------------------------------------
# COLETA
# ----------------------------------------------------------------------

def montar_tarefas(termos):
    """Expande as fontes em uma lista de (fonte, termo, url)."""
    tarefas = []
    for f in FONTES:
        if f.get("tipo") == "fort_pdf":
            tarefas.append((f, "", f.get("base_url", "")))
        elif "url_modelo" in f:
            for termo in termos:
                if f.get("tipo") == "mercadolivre":
                    t = sem_acento(termo).strip().replace(" ", "-")
                else:
                    t = quote_plus(termo)
                tarefas.append((f, termo, f["url_modelo"].format(termo=t)))
        elif "url" in f:
            tarefas.append((f, "", f["url"]))
    return tarefas


def _salvar_debug(fonte, termo, html):
    pasta = PASTA_SAIDA / "debug"
    pasta.mkdir(parents=True, exist_ok=True)
    slug = re.sub(r"[^a-z0-9]+", "_", sem_acento(f"{fonte['nome']}_{termo}")).strip("_")
    arq = pasta / f"{slug}.html"
    arq.write_text(html, encoding="utf-8")
    titulo = BeautifulSoup(html, "html.parser").title
    print(f"   0 itens. Título da página: {titulo.get_text(strip=True) if titulo else '(sem título)'}")
    print(f"   HTML salvo em {arq} (envie este arquivo para ajustarmos o leitor)")


def coletar(fonte: dict, termo: str, url: str, nav: Navegador):
    rotulo = fonte["nome"] + (f" [{termo}]" if termo else "")
    print(f"-> {rotulo}")
    tipo = fonte.get("tipo", "jsonld")
    html = ""
    try:
        if tipo == "fort_pdf":
            itens = coletar_fort_pdf(fonte, nav)
        else:
            parser = PARSERS.get(tipo, parse_generico)
            html = nav.html(url, validar=lambda h: bool(parser(h)))
            itens = parser(html)
            if not itens and re.search(r"login|registration|verification", getattr(nav, "ultima_url", "")):
                print("   O site pediu login. Rode uma vez  monitor_bebidas.exe --login  "
                      "(ou o login_mercadolivre.bat), entre na SUA conta e confirme; a sessão fica salva.")
            elif not itens and pagina_bloqueada(html):
                print("   BLOQUEADO (captcha/verificação). Mantenha \"mostrar_navegador\": true e resolva na janela.")
    except Exception as e:
        print(f"   ERRO: {e}")
        return []
    print(f"   {len(itens)} itens")
    if not itens and html and CFG.get("salvar_html_debug"):
        _salvar_debug(fonte, termo, html)

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
        it["categoria"] = classificar(it["produto"], termo)
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
        requests.post(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
                      json={"chat_id": TELEGRAM_CHAT_ID, "text": texto[:4000]}, timeout=20)
    except Exception as e:
        print(f"Falha ao enviar Telegram: {e}")


# ----------------------------------------------------------------------
# PRINCIPAL
# ----------------------------------------------------------------------

COLUNAS = ["fonte", "termo", "categoria", "produto", "preco", "preco_original", "desconto_pct",
           "litros", "preco_por_litro", "validade", "obs", "link", "coletado_em"]


def salvar_excel(caminho, df, alertas):
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Promocoes", index=False)
        if not alertas.empty:
            alertas.to_excel(w, sheet_name="Alertas", index=False)
        for ws in w.book.worksheets:
            ws.auto_filter.ref = ws.dimensions           # filtros nos cabeçalhos
            ws.freeze_panes = "A2"
            for col in ws.columns:
                tam = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, tam + 2), 60)


SITES_LOGIN = {
    "mercadolivre": "https://www.mercadolivre.com.br/",
    "amazon": "https://www.amazon.com.br/",
}


def login_manual(site: str):
    """Abre o navegador com o perfil do programa para você entrar na SUA conta.
    A sessão fica salva em perfil_navegador/ e é reutilizada nas próximas execuções."""
    print(f"Abrindo {site}. Faça login com a sua conta na janela que vai aparecer.")
    with Navegador(True, visivel=True, espera_s=0) as nav:
        pagina = nav._ctx.new_page()
        pagina.goto(SITES_LOGIN[site], timeout=60000)
        input("\nDepois de entrar na sua conta, volte aqui e pressione Enter para salvar a sessão...")
    print("Sessão salva. Agora rode o programa normalmente.")


def diagnostico(args):
    if args.html:
        html = Path(args.html).read_text(encoding="utf-8", errors="ignore")
        for it in PARSERS[args.parser](html)[:40]:
            print(f"{it['preco']:>9.2f}  {'OK ' if eh_alcoolica(it['produto']) else '-- '}"
                  f"[{classificar(it['produto'])}] {it['produto']}")
    if args.pdf:
        itens, fim = parse_pdf_fort(Path(args.pdf).read_bytes())
        print(f"Validade até: {fim}")
        for it in itens:
            print(f"{it['preco']:>9.2f}  [{classificar(it['produto'])}] {it['produto']}"
                  + (f"   (de {it['preco_original']:.2f})" if it.get("preco_original") else "")
                  + (f"   <{it['obs']}>" if it.get("obs") else ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-navegador", action="store_true",
                    help="usa apenas requests (mais rápido, mas falha em sites com JavaScript)")
    ap.add_argument("--sem-pausa", action="store_true",
                    help="não espera Enter no final (use no agendador de tarefas)")
    ap.add_argument("--tipos", "--categorias", dest="tipos",
                    help="filtra categorias nesta rodada, ex.: cerveja  ou  cerveja,vinho")
    ap.add_argument("--html", help="testa a extração em um arquivo HTML salvo")
    ap.add_argument("--parser", default="jsonld", choices=list(PARSERS), help="parser para --html")
    ap.add_argument("--pdf", help="testa a leitura de um PDF de encarte do Fort")
    ap.add_argument("--login", nargs="?", const="mercadolivre", choices=list(SITES_LOGIN),
                    help="abre o navegador para você entrar na sua conta (padrão: mercadolivre)")
    args = ap.parse_args()

    if args.login:
        return login_manual(args.login)

    if args.html or args.pdf:
        return diagnostico(args)

    tipos = [t.strip() for t in args.tipos.split(",")] if args.tipos else CFG["tipos_desejados"]
    tipos_norm = {sem_acento(t) for t in tipos if t.strip()}

    PASTA_SAIDA.mkdir(exist_ok=True)
    todos = []
    with Navegador(not args.sem_navegador, CFG["mostrar_navegador"], CFG["aguardar_verificacao_s"]) as nav:
        for fonte, termo, url in montar_tarefas(CFG["termos_busca"]):
            todos.extend(coletar(fonte, termo, url, nav))
            time.sleep(PAUSA_ENTRE_PAGINAS_S)

    if not todos:
        print("Nenhum item coletado. Verifique conexão, URLs ou bloqueios.")
        sys.exit(1)

    df = pd.DataFrame(todos).drop_duplicates(subset=["produto", "preco", "fonte"])
    total = len(df)
    if CFG["somente_alcoolicas"]:
        df = df[df["produto"].map(eh_alcoolica)]
    if tipos_norm:
        df = df[df["categoria"].map(sem_acento).isin(tipos_norm)]
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = None
    df = df.sort_values(["categoria", "preco_por_litro", "preco"], na_position="last")[COLUNAS]

    print(f"\n{total} itens coletados, {len(df)} após os filtros"
          + (f" (tipos: {', '.join(sorted(tipos_norm))})" if tipos_norm else "") + ".")
    if df.empty:
        print("Nada restou após o filtro.")
        return
    print(df.groupby(["fonte", "categoria"]).size().to_string())

    alertas = avaliar_alertas(df)
    xlsx = PASTA_SAIDA / f"promocoes_{datetime.now():%Y%m%d_%H%M}.xlsx"
    salvar_excel(xlsx, df, alertas)
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
        linhas.append(f"• [{a['categoria']}] {a['produto']}: R$ {a['preco']:.2f}{dc}{pl} [{a['fonte']}]")
    mensagem = "\n".join(linhas)
    print("\n" + mensagem)
    enviar_telegram(mensagem)


if __name__ == "__main__":
    try:
        main()
    finally:
        if getattr(sys, "frozen", False) and sys.stdin and sys.stdin.isatty() and "--sem-pausa" not in sys.argv:
            input("\nPressione Enter para fechar...")
