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
import urllib.request
import urllib.parse
import urllib.error
from concurrent.futures import ThreadPoolExecutor, as_completed

OPENLIBRARY = "https://covers.openlibrary.org/b/isbn/{isbn}-M.jpg?default=false"
GOOGLE_BOOKS = "https://www.googleapis.com/books/v1/volumes?maxResults=1&q={q}"
TIMEOUT = 10
WORKERS = 5
RETRIES = 3


def with_retries(fn, *args):
    """Roda fn(*args), tentando de novo em caso de erro de rede/limite de
    requisicoes, com uma pausa maior a cada tentativa. Devolve (resultado, None)
    em caso de sucesso, ou (None, motivo_do_erro) se todas as tentativas falharem."""
    last_error = None
    for attempt in range(RETRIES):
        try:
            return fn(*args), None
        except urllib.error.HTTPError as e:
            last_error = f"HTTP {e.code}"
            if e.code == 429:
                time.sleep(2 + attempt * 2)
            else:
                time.sleep(0.5 + attempt)
        except Exception as e:
            last_error = str(e)[:80]
            time.sleep(0.5 + attempt)
    return None, last_error


def fetch_bytes(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (biblioteca-pessoal)"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        ctype = resp.headers.get("Content-Type", "image/jpeg")
        data = resp.read()
        return data, ctype


def fetch_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (biblioteca-pessoal)"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def cover_for_book(book):
    """Tenta Open Library pelo ISBN, depois Google Books (ISBN e depois
    titulo+autor). Devolve (data, ctype, motivo_da_falha)."""
    isbn13 = book.get("i13")
    isbn10 = book.get("i10")
    last_reason = "sem ISBN e sem resultado no Google Books"

    # 1) Open Library por ISBN
    for isbn in filter(None, [isbn13, isbn10]):
        result, err = with_retries(fetch_bytes, OPENLIBRARY.format(isbn=isbn))
        if result:
            data, ctype = result
            if data and len(data) > 300:
                return data, ctype, None
        elif err:
            last_reason = f"Open Library: {err}"

    # 2) Google Books: por ISBN, depois por titulo+autor
    queries = []
    if isbn13:
        queries.append("isbn:" + isbn13)
    if isbn10:
        queries.append("isbn:" + isbn10)
    queries.append("intitle:{} inauthor:{}".format(book.get("t", ""), book.get("a", "")))

    for q in queries:
        url = GOOGLE_BOOKS.format(q=urllib.parse.quote(q))
        result, err = with_retries(fetch_json, url)
        if err:
            last_reason = f"Google Books: {err}"
            continue
        items = (result or {}).get("items") or []
        if not items:
            continue
        links = items[0].get("volumeInfo", {}).get("imageLinks", {})
        thumb = links.get("thumbnail") or links.get("smallThumbnail")
        if not thumb:
            continue
        thumb = thumb.replace("http://", "https://")
        img_result, err = with_retries(fetch_bytes, thumb)
        if img_result:
            data, ctype = img_result
            if data and len(data) > 300:
                return data, ctype, None
        elif err:
            last_reason = f"download da capa: {err}"

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
