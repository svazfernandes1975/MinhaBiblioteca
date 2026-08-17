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

Baixa varios livros ao mesmo tempo (ate 10 em paralelo), entao para
~340 livros costuma levar de 2 a 6 minutos. O arquivo final fica bem
maior (pode passar de 10-15 MB), porque as imagens vão embutidas
no próprio HTML.
"""
import sys
import re
import json
import base64
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

OPENLIBRARY = "https://covers.openlibrary.org/b/isbn/{isbn}-M.jpg?default=false"
GOOGLE_BOOKS = "https://www.googleapis.com/books/v1/volumes?maxResults=1&q={q}"
TIMEOUT = 6
WORKERS = 10


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
    """Tenta Open Library pelo ISBN, depois Google Books (ISBN e depois titulo+autor)."""
    isbn13 = book.get("i13")
    isbn10 = book.get("i10")

    # 1) Open Library por ISBN
    for isbn in filter(None, [isbn13, isbn10]):
        try:
            data, ctype = fetch_bytes(OPENLIBRARY.format(isbn=isbn))
            if data and len(data) > 300:  # descarta respostas vazias/lixo
                return data, ctype
        except Exception:
            pass

    # 2) Google Books: por ISBN, depois por titulo+autor
    queries = []
    if isbn13:
        queries.append("isbn:" + isbn13)
    if isbn10:
        queries.append("isbn:" + isbn10)
    queries.append("intitle:{} inauthor:{}".format(book.get("t", ""), book.get("a", "")))

    for q in queries:
        try:
            url = GOOGLE_BOOKS.format(q=urllib.parse.quote(q))
            result = fetch_json(url)
            items = result.get("items") or []
            if not items:
                continue
            links = items[0].get("volumeInfo", {}).get("imageLinks", {})
            thumb = links.get("thumbnail") or links.get("smallThumbnail")
            if not thumb:
                continue
            thumb = thumb.replace("http://", "https://")
            data, ctype = fetch_bytes(thumb)
            if data and len(data) > 300:
                return data, ctype
        except Exception:
            pass

    return None, None


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
    total = len(books)
    done = 0

    def worker(book):
        return book, *cover_for_book(book)

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = [pool.submit(worker, book) for book in books]
        for future in as_completed(futures):
            book, data, ctype = future.result()
            done += 1
            if data:
                b64 = base64.b64encode(data).decode("ascii")
                cover_map[book["id"]] = f"data:{ctype};base64,{b64}"
                status = "OK"
            else:
                status = "sem capa encontrada"
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
    print("Esse arquivo funciona sem internet (menos os gráficos de gênero, que continuam usando o Google Books).")


if __name__ == "__main__":
    main()
