# Atualizar a versão web (repo guipasquetti/pasquetti-cotacoes)

A versão web está atrasada. Esta pasta `deploy_web/` tem **exatamente** os arquivos que
devem ir para o GitHub. Os segredos (chaves) NÃO estão aqui — eles ficam só nos *Secrets*
da Streamlit Cloud.

## O que mudou em relação ao que está no repo hoje

- `dados_supabase.py` — **faltava no repo inteiro** (por isso a web não aprende, não aplica ST nem salva clientes).
- `catalogo_produtos.json` — agora com **694 produtos** (a tabela **SMU**, 130 itens, não estava na web).
- `app_cotacao.py` — versão nova (layout, busca em tempo real, condição de pagamento, botão "salvar e ensinar", banco de clientes/SEFAZ).
- `requirements.txt` — agora inclui `streamlit-searchbox`.
- `logo_horizontal.png`, `logo_vertical.png`, `.streamlit/config.toml` — faltavam/desatualizados.

## Como subir (pelo navegador, sem terminal)

1. Abra https://github.com/guipasquetti/pasquetti-cotacoes
2. Botão **Add file → Upload files**.
3. Arraste **todos os arquivos desta pasta** (inclusive a pasta `.streamlit`).
4. Em "Commit changes", escreva algo como `Atualiza app, catálogo SMU e dados_supabase` → **Commit**.
5. A Streamlit Cloud detecta e redeploya sozinha em 1–2 minutos.

> Confira nos *Secrets* da Streamlit Cloud (App → ⋮ → Settings → Secrets) que existem
> `APP_PASSWORD`, `ANTHROPIC_API_KEY`, `SUPABASE_URL` e `SUPABASE_KEY`. Sem `SUPABASE_*`,
> a web continua sem aprender/ST mesmo com o `dados_supabase.py` no lugar.
