#!/usr/bin/env python3
"""
Monitor de promoções de bebidas - região de Joinville/SC

O que faz:
  1. Abre as páginas de bebidas do Angeloni e do Tiendeo (Joinville)
  2. Extrai produtos e preços (via dados estruturados JSON-LD da página)
  3. Calcula o preço por litro
  4. Salva em Excel/CSV e mantém um histórico
  5. Dispara alertas (console e, opcionalmente, Telegram) conforme suas regras

Instalação:
  pip install playwright beautifulsoup4 pandas openpyxl requests
  (o robô usa o Edge/Chrome já instalado; se não houver, rode: playwright install chromium)

Uso:
  python monitor_bebidas.py
  python monitor_bebidas.py --sem-navegador   # usa só requests (sites simples)

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
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------
# CONFIGURAÇÃO (edite aqui)
# ----------------------------------------------------------------------

# Pasta onde ficam o config.json e a saída (ao lado do .exe ou do script)
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
ARQ_CONFIG = BASE_DIR / "config.json"

CONFIG_PADRAO = {
    "fontes": [
        {"nome": "Angeloni - Black Friday Bebidas",
         "url": "https://super.angeloni.com.br/black-friday-em-bebidas"},
        {"nome": "Tiendeo - Cerveja Joinville",
         "url": "https://www.tiendeo.com.br/joinville/ofertas/cerveja"},
        {"nome": "Tiendeo - Bebidas alcoólicas Joinville",
         "url": "https://www.tiendeo.com.br/joinville/ofertas/bebidas-alcoolicas"},
    ],
    # Alerta: nome contém a palavra E (preço <= max_preco OU preço/litro <= max_por_litro).
    # Use null para ignorar um limite.
    "regras": [
        {"palavra": "heineken",   "max_preco": 5.00, "max_por_litro": None},
        {"palavra": "stella",     "max_preco": None, "max_por_litro": 16.00},
        {"palavra": "coca-cola",  "max_preco": None, "max_por_litro": 7.00},
        {"palavra": "skol",       "max_preco": None, "max_por_litro": 10.00},
        {"palavra": "antarctica", "max_preco": None, "max_por_litro": 10.00},
    ],
    "pausa_entre_paginas_s": 3,
    "telegram_token": "",
    "telegram_chat_id": "",
}


def carregar_config() -> dict:
    """Lê config.json; se não existir, cria um com os valores padrão."""
    if not ARQ_CONFIG.exists():
        ARQ_CONFIG.write_text(json.dumps(CONFIG_PADRAO, indent=2, ensure_ascii=False),
                              encoding="utf-8")
        print(f"Criei {ARQ_CONFIG} com valores de exemplo. Edite-o e rode de novo para personalizar.\n")
        return dict(CONFIG_PADRAO)
    cfg = dict(CONFIG_PADRAO)
    cfg.update(json.loads(ARQ_CONFIG.read_text(encoding="utf-8")))
    return cfg


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

def baixar_html_requests(url: str) -> str:
    r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
    r.raise_for_status()
    return r.text


def baixar_html_playwright(url: str) -> str:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        navegador = None
        for canal in ("msedge", "chrome", None):   # Edge já vem no Windows
            try:
                navegador = (p.chromium.launch(channel=canal, headless=True)
                             if canal else p.chromium.launch(headless=True))
                break
            except Exception:
                continue
        if navegador is None:
            raise RuntimeError("Nenhum navegador encontrado. Instale o Microsoft Edge ou o Google Chrome.")
        contexto = navegador.new_context(user_agent=USER_AGENT, locale="pt-BR")
        pagina = contexto.new_page()
        pagina.goto(url, wait_until="domcontentloaded", timeout=60000)
        try:
            pagina.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass  # segue com o que já carregou
        # rola a página para disparar carregamento preguiçoso (lazy load)
        for _ in range(5):
            pagina.mouse.wheel(0, 2500)
            time.sleep(0.8)
        html = pagina.content()
        navegador.close()
    return html


# ----------------------------------------------------------------------
# EXTRAÇÃO
# ----------------------------------------------------------------------

def _para_float(valor):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "")
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
                    "validade": of.get("validThrough") or of.get("priceValidUntil") or "",
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
        # o nome costuma estar na linha anterior
        nome = linhas[i - 1] if i > 0 else ""
        if len(nome) < 5 or "R$" in nome:
            continue
        itens.append({"produto": nome, "preco": _para_float(m.group(1)), "validade": ""})
    return itens


def volume_em_litros(nome: str):
    """Descobre o volume total (em litros) a partir do nome do produto."""
    n = nome.lower().replace(",", ".")
    # padrão "12 x 350ml" ou "12x350 ml"
    m = re.search(r"(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(ml|l|lt|litros?)\b", n)
    if m:
        qtd, vol, un = int(m.group(1)), float(m.group(2)), m.group(3)
        return qtd * (vol / 1000 if un == "ml" else vol)
    # padrão simples "350ml", "1.5l", "2 litros"
    m = re.search(r"(\d+(?:\.\d+)?)\s*(ml|l|lt|litros?)\b", n)
    if m:
        vol, un = float(m.group(1)), m.group(2)
        litros = vol / 1000 if un == "ml" else vol
        # "fardo com 12 un 350ml" -> multiplica pela quantidade, se informada
        q = re.search(r"(?:fardo|pack|caixa|c/|com)\s*(\d+)\s*(?:un|unid|latas?|garrafas?)?", n)
        if q and int(q.group(1)) > 1 and litros < 1:
            litros *= int(q.group(1))
        return litros
    return None


def coletar(fonte: dict, usar_navegador: bool):
    print(f"-> {fonte['nome']}")
    try:
        html = (baixar_html_playwright if usar_navegador else baixar_html_requests)(fonte["url"])
    except Exception as e:
        print(f"   ERRO ao baixar: {e}")
        return []

    itens = extrair_jsonld(html)
    metodo = "JSON-LD"
    if not itens:
        itens = extrair_texto_fallback(html)
        metodo = "texto (fallback)"
    print(f"   {len(itens)} itens via {metodo}")

    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    for it in itens:
        it["fonte"] = fonte["nome"]
        it["url"] = fonte["url"]
        it["coletado_em"] = agora
        litros = volume_em_litros(it["produto"])
        it["litros"] = litros
        it["preco_por_litro"] = round(it["preco"] / litros, 2) if litros else None
    return itens


# ----------------------------------------------------------------------
# ALERTAS
# ----------------------------------------------------------------------

def avaliar_alertas(df: pd.DataFrame) -> pd.DataFrame:
    achados = []
    for regra in REGRAS:
        filtro = df["produto"].str.lower().str.contains(regra["palavra"], na=False)
        sub = df[filtro]
        for _, linha in sub.iterrows():
            ok_preco = regra["max_preco"] is not None and linha["preco"] <= regra["max_preco"]
            ok_litro = (regra["max_por_litro"] is not None
                        and pd.notna(linha["preco_por_litro"])
                        and linha["preco_por_litro"] <= regra["max_por_litro"])
            if ok_preco or ok_litro:
                achados.append(linha)
    return pd.DataFrame(achados).drop_duplicates(subset=["produto", "preco", "fonte"]) \
        if achados else pd.DataFrame()


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

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-navegador", action="store_true",
                    help="usa apenas requests (mais rápido, mas falha em sites com JavaScript)")
    ap.add_argument("--sem-pausa", action="store_true",
                    help="não espera Enter no final (use no agendador de tarefas)")
    args = ap.parse_args()

    PASTA_SAIDA.mkdir(exist_ok=True)
    todos = []
    for fonte in FONTES:
        todos.extend(coletar(fonte, usar_navegador=not args.sem_navegador))
        time.sleep(PAUSA_ENTRE_PAGINAS_S)

    if not todos:
        print("Nenhum item coletado. Verifique conexão, URLs ou seletores.")
        sys.exit(1)

    df = pd.DataFrame(todos).drop_duplicates(subset=["produto", "preco", "fonte"])
    df = df.sort_values(["preco_por_litro", "preco"], na_position="last")

    carimbo = datetime.now().strftime("%Y%m%d_%H%M")
    xlsx = PASTA_SAIDA / f"promocoes_{carimbo}.xlsx"
    df.to_excel(xlsx, index=False)

    # histórico acumulado (para ver a evolução dos preços)
    hist = PASTA_SAIDA / "historico.csv"
    df.to_csv(hist, mode="a", header=not hist.exists(), index=False, encoding="utf-8-sig")

    print(f"\n{len(df)} produtos salvos em {xlsx}")

    alertas = avaliar_alertas(df)
    if alertas.empty:
        print("Nenhum alerta pelas suas regras hoje.")
        return

    linhas = ["ALERTAS DE PROMOÇÃO - Joinville"]
    for _, a in alertas.iterrows():
        pl = f" (R$ {a['preco_por_litro']:.2f}/L)" if pd.notna(a["preco_por_litro"]) else ""
        linhas.append(f"• {a['produto']}: R$ {a['preco']:.2f}{pl} [{a['fonte']}]")
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
