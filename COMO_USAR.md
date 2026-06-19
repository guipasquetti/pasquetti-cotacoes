# 🔧 Automação de Cotações — Pasquetti

## O que faz

Este sistema recebe uma solicitação de cotação em qualquer formato e gera uma planilha Excel formatada no padrão do orçamento Pasquetti, com preços buscados automaticamente nas tabelas de preço.

---

## Arquivos do sistema

| Arquivo | Descrição |
|---|---|
| `app_cotacao.py` | App principal (interface web). Abra com **2_INICIAR COTAÇÕES.command** |
| `matcher_ia.py` | Camada de IA: interpreta pedidos vagos e lê foto/PDF |
| `extrair_catalogo.py` | Atualiza o catálogo. Rode pelo **3_ATUALIZAR PREÇOS.command** |
| `catalogo_produtos.json` | Base de dados de produtos e preços (gerada automaticamente) |
| `gerar_cotacao.py` | Versão antiga por terminal (CLI) — opcional |

---

## 🤖 Interpretação com IA (recomendado)

O sistema tem duas camadas de busca de produto:

1. **Motor automático** (sempre ligado, grátis, offline) — entende abreviações e
   casa o produto por similaridade. Resolve a maioria dos pedidos bem escritos.
2. **Camada de IA** (opcional) — entende pedidos **vagos ou que exigem
   interpretação** ("curva de 200 pra ligar dois tubos", "o anel que veda a bolsa
   do SMU 100") e também **lê foto e PDF** automaticamente, sem instalar nada.

### Como ligar a IA
1. Gere uma chave de API em **console.anthropic.com → Settings → API keys → Create Key**.
2. Abra o sistema (**2_INICIAR COTAÇÕES.command**).
3. Na barra lateral, em **🤖 Interpretação com IA**, cole a chave (`sk-ant-...`) e
   clique em **Salvar e ativar IA**. Recarregue a página.

> A chave fica salva no arquivo `config_ia.json` na pasta. **Custo:** cerca de
> **2 a 4 centavos de dólar por cotação** (modelo Claude Haiku 4.5). Precisa de internet.

---

## Como usar (passo a passo)

### 1. Preparar a solicitação

Crie um arquivo de texto (`.txt`) com os itens pedidos, um por linha. Exemplos de formatos aceitos:

```
90  joelho fofo 100mm
72  JF 75mm 90 graus
3   joelho HL 150 90gr
6   anel 150mm
190 anel borracha 100
150 anel 75mm
10  luva fofo 100mm
5   tee 100x75
```

Ou envie uma planilha `.xlsx` com colunas de Produto e Quantidade.

> **Dica:** O sistema entende abreviações como `JF`, `JFF`, `anel borr`, `TCI`, etc.

---

### 2. Abrir o Terminal

- Pressione `Cmd + Espaço`, digite **Terminal** e pressione Enter

---

### 3. Navegar até a pasta da automação

```bash
cd ~/Library/Mobile\ Documents/com\~apple\~CloudDocs/Pasquetti\ Com.\ Mat.\ Hidra/Automacao\ Cotacoes
```

---

### 4. Gerar a cotação

**Para tabela de CONSUMO (padrão):**
```bash
python3 gerar_cotacao.py meu_pedido.txt --cliente "Nome do Cliente" --num 1322
```

**Para tabela de REVENDA:**
```bash
python3 gerar_cotacao.py meu_pedido.txt --tabela revenda --cliente "Nome Revenda" --num 1323
```

**Para tabela de PRESSÃO (JGS):**
```bash
python3 gerar_cotacao.py meu_pedido.txt --tabela pressao --cliente "Construtora ABC" --num 1324
```

**Com dados completos do cliente:**
```bash
python3 gerar_cotacao.py pedido.txt \
    --tabela consumo \
    --cliente "JUQUEI INCORPORADORA LTDA" \
    --cnpj "28.450.705/0001-20" \
    --endereco "AV REPUBLICA DO LIBANO, 1921 - IBIRAPUERA - SP" \
    --telefone "(11) 5056-8300" \
    --email "contato@cliente.com.br" \
    --num 1325 \
    --output cotacao_juquei.xlsx
```

---

### 5. Verificar o resultado

O arquivo gerado `cotacao_<num>.xlsx` aparece na mesma pasta. Ele tem duas abas:

- **Orçamento** — cotação formatada pronta para enviar
- **Revisão** — lista completa com nível de confiança de cada item

#### Níveis de confiança:
- ✅ **ALTA / MÉDIA** — produto identificado com segurança
- ⚠️ **BAIXA** (laranja) — item incluído mas merece confirmação visual
- ❌ **NÃO ENCONTRADO** (vermelho) — precisa ser adicionado manualmente

---

## Abreviações reconhecidas

| Você escreve | O sistema entende |
|---|---|
| `JF` ou `JFF` | Joelho Fofo HL |
| `anel borr` ou `anel` | Anel Borracha HL |
| `tee fofo` ou `te` | Tee Fofo HL |
| `luva HL` ou `LJF` | Luva Fofo HL |
| `90 gr`, `90°`, `90 graus` | 90 graus |
| `45 gr`, `45°` | 45 graus |
| `b/b`, `BB` | Bolsa/Bolsa |
| `JR`, `JRSMU` | Junta Rapid JRSMU |
| `TCI` | Tee c/ Inspeção |
| `valv gav` | Válvula Gaveta |
| `hidrante col` | Hidrante Coluna |

---

## Atualizar os preços

Quando receber novas tabelas de preço, substitua os arquivos Excel e rode:

```bash
python3 extrair_catalogo.py
```

O `catalogo_produtos.json` será regenerado automaticamente.

---

## Foto e PDF de pedido

Com a **IA ligada** (ver seção acima), basta enviar a foto (`.jpg`/`.png`) ou o
`.pdf` na aba **Carregar Arquivo** — a IA lê a lista e as quantidades sozinha.
Não precisa instalar Tesseract.

*(Sem a IA, fotos ainda funcionam por OCR se você instalar `brew install tesseract tesseract-lang`, mas a leitura por IA é bem mais precisa e lê PDF também.)*

---

*Criado por Pasquetti Automação — Maio/2026*
