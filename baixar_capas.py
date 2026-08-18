#!/usr/bin/env python3
"""
Baixa as capas dos livros e gera uma versao 100% offline do site
(sem precisar de internet depois de pronto, e sem usar localStorage).

Como usar:
    python3 baixar_capas.py index.html
    python3 baixar_capas.py index.html index.html   # sobrescreve o proprio arquivo

Isso cria (por padrao) um novo arquivo "index_offline.html" com todas
as capas encontradas embutidas diretamente no HTML (em base64). Só usa
bibliotecas padrão do Python, não precisa instalar nada.

Voce nao precisa rodar isso na sua maquina: o workflow em
.github/workflows/atualizar-capas.yml roda esse script automaticamente
nos servidores do GitHub (Actions), que tem internet livre. Basta subir
os arquivos no repositorio e clicar em "Run workflow".

Baixa varios livros ao mesmo tempo, com novas tentativas automaticas em
caso de falha/limite de requisicoes. Para ~340 livros costuma levar de
3 a 10 minutos. O arquivo final fica bem maior (pode passar de 10-15 MB),
porque as imagens vão embutidas no próprio HTML.
"""
import sys
import re
import json
import time
import base64
import threading
import urllib.request
import urllib.parse
import urllib.error
from concurrent.futures import ThreadPoolExecutor, as_completed

OPENLIBRARY = "https://covers.openlibrary.org/b/isbn/{isbn}-M.jpg?default=false"
GOOGLE_BOOKS = "https://www.googleapis.com/books/v1/volumes?maxResults=1&q={q}"
TIMEOUT = 8
WORKERS = 8
RETRIES = 2

# A Open Library aceita varias chamadas ao mesmo tempo sem problema.
# O Google Books, vindo dos servidores do GitHub, costuma limitar bastante
# (muita gente usa o mesmo IP compartilhado). Por isso as chamadas pra ele
# passam por um freio global: no maximo 1 a cada ~1.1s, para todo mundo.
_google_lock = threading.Lock()
_google_next_time = [0.0]
GOOGLE_MIN_INTERVAL = 1.1


def google_books_throttle():
    with _google_lock:
        now = time.time()
        wait = _google_next_time[0] - now
        if wait > 0:
            time.sleep(wait)
            now = time.time()
        _google_next_time[0] = now + GOOGLE_MIN_INTERVAL


def with_retries(fn, *args):
    """Roda fn(*args). Erros definitivos (404 = 'não existe') não são
    tentados de novo, pra não perder tempo à toa. Só insiste em erros
    temporários (limite de requisições, timeout, erro 5xx do servidor).
    Devolve (resultado, None) em caso de sucesso, ou (None, motivo) se falhar."""
    last_error = None
    for attempt in range(RETRIES):
        try:
            return fn(*args), None
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None, "HTTP 404 (não encontrado)"  # definitivo, não insiste
            last_error = f"HTTP {e.code}"
            if e.code == 429:
                time.sleep(2 + attempt * 2)
            else:
                time.sleep(0.5)
        except Exception as e:
            last_error = str(e)[:80]
            time.sleep(0.5)
    return None, last_error


def fetch_bytes(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (biblioteca-pessoal)"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        ctype = resp.headers.get("Content-Type", "image/jpeg")
        data = resp.read()
        return data, ctype


def fetch_json_google(url):
    """Igual ao fetch_json, mas passa pelo freio global antes de chamar."""
    google_books_throttle()
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (biblioteca-pessoal)"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def cover_for_book(book):
    """Tenta Open Library pelo ISBN (rapido, sem limite forte), depois UMA
    busca no Google Books (a melhor opção disponível: ISBN, ou título+autor
    se não tiver ISBN). Devolve (data, ctype, motivo_da_falha)."""
    isbn13 = book.get("i13")
    isbn10 = book.get("i10")
    last_reason = "não encontrado em nenhuma fonte"

    # 1) Open Library por ISBN
    for isbn in filter(None, [isbn13, isbn10]):
        result, err = with_retries(fetch_bytes, OPENLIBRARY.format(isbn=isbn))
        if result:
            data, ctype = result
            if data and len(data) > 300:
                return data, ctype, None
        elif err:
            last_reason = f"Open Library: {err}"

    # 2) Uma única busca no Google Books (throttled) — prioriza ISBN, senão título+autor
    if isbn13:
        q = "isbn:" + isbn13
    elif isbn10:
        q = "isbn:" + isbn10
    else:
        q = "intitle:{} inauthor:{}".format(book.get("t", ""), book.get("a", ""))

    url = GOOGLE_BOOKS.format(q=urllib.parse.quote(q))
    result, err = with_retries(fetch_json_google, url)
    if err:
        return None, None, f"Google Books: {err}"

    items = (result or {}).get("items") or []
    if items:
        links = items[0].get("volumeInfo", {}).get("imageLinks", {})
        thumb = links.get("thumbnail") or links.get("smallThumbnail")
        if thumb:
            thumb = thumb.replace("http://", "https://")
            img_result, err = with_retries(fetch_bytes, thumb)
            if img_result:
                data, ctype = img_result
                if data and len(data) > 300:
                    return data, ctype, None
            elif err:
                return None, None, f"download da capa: {err}"

    return None, None, last_reason


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 baixar_capas.py index.html [arquivo_de_saida.html]")
        sys.exit(1)

    src_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "index_offline.html"
    with open(src_path, encoding="utf-8") as f:
        html = f.read()

    m = re.search(
        r'<script id="books-data" type="application/json">(.*?)</script>',
        html, re.S,
    )
    if not m:
        print("Não achei os dados dos livros nesse HTML. Use o index.html gerado pelo Claude.")
        sys.exit(1)
    books = json.loads(m.group(1))

    cover_map = {}
    fail_reasons = {}
    total = len(books)
    done = 0

    def worker(book):
        data, ctype, reason = cover_for_book(book)
        return book, data, ctype, reason

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = [pool.submit(worker, book) for book in books]
        for future in as_completed(futures):
            book, data, ctype, reason = future.result()
            done += 1
            if data:
                b64 = base64.b64encode(data).decode("ascii")
                cover_map[book["id"]] = f"data:{ctype};base64,{b64}"
                status = "OK"
            else:
                status = f"falhou ({reason})"
                fail_reasons[reason] = fail_reasons.get(reason, 0) + 1
            print(f"[{done}/{total}] {book['t'][:50]}... {status}", flush=True)

    cover_json = json.dumps(cover_map, ensure_ascii=False)
    new_html = re.sub(
        r'<script id="cover-data" type="application/json">.*?</script>',
        f'<script id="cover-data" type="application/json">{cover_json}</script>',
        html, flags=re.S,
    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    found = len(cover_map)
    print(f"\nPronto! {found}/{total} capas embutidas em {out_path}")
    if fail_reasons:
        print("\nResumo dos motivos de falha (mais comuns primeiro):")
        for reason, count in sorted(fail_reasons.items(), key=lambda x: -x[1])[:10]:
            print(f"  {count}x — {reason}")
    print("\nEsse arquivo funciona sem internet (menos os gráficos de gênero, que continuam usando o Google Books).")


if __name__ == "__main__":
    main()
