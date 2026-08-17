# Minha Biblioteca

Site pessoal de controle de leitura, gerado a partir do export do Goodreads.

## Publicar no GitHub Pages (sem instalar nada)

1. Crie um repositório novo no GitHub e suba estes arquivos:
   - `index.html`
   - `baixar_capas.py`
   - `.github/workflows/atualizar-capas.yml`

2. Vá em **Settings → Pages** e configure:
   - Source: `Deploy from a branch`
   - Branch: `main` (ou `master`), pasta `/ (root)`
   - Salve. Em alguns minutos o site fica no ar em `https://SEU_USUARIO.github.io/SEU_REPO/`

3. Para embutir as capas dos livros diretamente no HTML (sem precisar de
   internet depois de carregado, e sem usar localStorage):
   - Vá na aba **Actions** do repositório
   - Clique no workflow **"Atualizar capas dos livros"**
   - Clique em **"Run workflow"**
   - Aguarde alguns minutos — ele baixa as capas (Open Library + Google
     Books) e faz um commit automático do `index.html` já com tudo embutido
   - O GitHub Pages atualiza sozinho depois desse commit

Repita esse passo 3 sempre que quiser recarregar/atualizar as capas
(por exemplo, depois de eu te mandar uma nova versão do site com livros
novos).

## Atualizar a lista de livros

Quando tiver um export novo do Goodreads, me manda o CSV aqui no chat que
eu gero um `index.html` novo pra você subir no lugar do antigo (repita o
passo 3 depois, se quiser as capas dos livros novos embutidas também).
