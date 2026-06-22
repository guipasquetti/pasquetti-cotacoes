"""
matcher_ia.py — Camada de interpretação com IA (API Claude) para a Pasquetti.

Resolve os dois casos que o matching por texto não cobre:
  1. Pedidos imprecisos / que exigem interpretação ("curva de 200 pra ligar
     dois tubos", "o anel que veda a bolsa do SMU 100").
  2. Leitura de fontes bagunçadas: foto (WhatsApp), PDF e texto livre.

A chave de API é lida de:
  • variável de ambiente ANTHROPIC_API_KEY, ou
  • arquivo  config_ia.json  na mesma pasta:  {"api_key": "sk-ant-...","modelo":"..."}

Se não houver chave / a lib 'anthropic' não estiver instalada, o app continua
funcionando só com o motor determinístico (a IA é um reforço opcional).
"""
import os
import json
import re
import base64
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config_ia.json"

# Modelo padrão: rápido e barato (~centavos por cotação). Pode trocar no config.
MODELO_PADRAO = "claude-haiku-4-5-20251001"


# ════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ════════════════════════════════════════════════════════════════════════════

def carregar_config():
    cfg = {"api_key": "", "modelo": MODELO_PADRAO}
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            cfg["api_key"] = (data.get("api_key") or "").strip()
            cfg["modelo"]  = (data.get("modelo") or MODELO_PADRAO).strip()
        except Exception:
            pass
    if not cfg["api_key"]:
        cfg["api_key"] = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    return cfg

def salvar_config(api_key, modelo=MODELO_PADRAO):
    CONFIG_PATH.write_text(
        json.dumps({"api_key": api_key.strip(), "modelo": modelo.strip()},
                   ensure_ascii=False, indent=2),
        encoding="utf-8")

def ia_disponivel():
    """True se a lib 'anthropic' está instalada E há uma chave configurada."""
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return False
    return bool(carregar_config()["api_key"])

def _cliente():
    import anthropic
    cfg = carregar_config()
    if not cfg["api_key"]:
        raise RuntimeError("Chave de API não configurada.")
    return anthropic.Anthropic(api_key=cfg["api_key"]), cfg["modelo"]


# ════════════════════════════════════════════════════════════════════════════
# GLOSSÁRIO DE DOMÍNIO  (ensina as abreviações da casa para a IA)
# ════════════════════════════════════════════════════════════════════════════

GLOSSARIO = """
Glossário Pasquetti (materiais hidráulicos de ferro fundido):
- JF / JFF = joelho fofo (ferro fundido); "fofo" = ferro fundido.
- TÊ / TEE = conexão T (três saídas). "TCI" = tê com inspeção.
- Junção = conexão em Y.
- Luva = união reta entre dois tubos. "LUVA CR" / "LUVA JM" (junta mecânica) / "Luva Gr Tol".
- Anel borr. = anel de borracha de vedação. "Anel trav." = anel travante interno.
- Bucha de redução / Redução (RED.) = passa de um diâmetro para outro (ex.: 100x75).
- Tubo SMU = ponta/ponta (P/P); Tubo SME = ponta/bolsa (P/B).
- JGS = linha de pressão (água/esgoto) em ferro fundido dúctil.
- TUBO K7 / K9 = classe de espessura do tubo de pressão (K9 mais reforçado).
- TUBINT = tubo internão (esgoto, ESG).
- Conexões: BB = bolsa/bolsa; PP = ponta/ponta; PB = ponta/bolsa; C/Fl ou CFL = com flange.
- Curva BB 90º/45º/22'30/11º15' = curvas de pressão por ângulo.
- Caps = tampão de extremidade. Extrem. = extremidade (P/FL, B/FL, com aba de vedação).
- Válv. Gav. = válvula de gaveta; Válv. Ret. = válvula de retenção.
- Ultralink / Ultraquik = juntas de reparo/acoplamento rápido.
- Junta Gibault / Junta de Desmontagem = acessórios de montagem em flange.
- Colar de tomada = derivação para tubo. Ventosa = válvula de ar.
- Diâmetros em mm (50, 75, 80, 100, 125, 150, 200, 250, 300, 350, 400...).
  Reduções e tês usam dois números: 100x75 (entrada x saída).
"""


# ════════════════════════════════════════════════════════════════════════════
# 1) INTERPRETAÇÃO + MATCHING
# ════════════════════════════════════════════════════════════════════════════

def _catalogo_texto(catalogo):
    """Lista compacta numerada do catálogo para enviar à IA."""
    linhas = []
    for i, p in enumerate(catalogo):
        linhas.append(f"{i}\t{p['descricao']}\t(R$ {p['preco']:.2f})")
    return "\n".join(linhas)

def _exemplos_correcoes(correcoes, itens_brutos, limite=60):
    """Monta um bloco de exemplos 'texto do cliente => produto' a partir das
    correções já ensinadas, priorizando as mais parecidas com o pedido atual,
    para a IA aprender o PADRÃO e generalizar (não só decorar o texto exato)."""
    if not correcoes:
        return ""
    def toks(s):
        return set(re.sub(r"[^\w\s]", " ", str(s).lower()).split())
    alvo = set()
    for it in itens_brutos:
        alvo |= toks(it.get("descricao", ""))
    pares = list(correcoes.items())  # (texto_norm, produto_descricao)
    pares.sort(key=lambda p: len((toks(p[0]) | toks(p[1])) & alvo), reverse=True)
    pares = pares[:limite]
    if not pares:
        return ""
    linhas = "\n".join(f'- "{k}" => "{prod}"' for k, prod in pares)
    return (
        "\nEXEMPLOS DE CORREÇÕES JÁ ENSINADAS PELOS VENDEDORES "
        "(o texto do cliente aparece normalizado: minúsculas, sem acento). "
        "Aprenda o PADRÃO por trás delas — abreviações, jeitos de escrever, "
        "equivalências — e generalize para pedidos parecidos, mesmo que o texto "
        "não seja idêntico:\n"
        f"{linhas}\n"
    )

def interpretar_itens(itens_brutos, catalogo, hints=None, correcoes=None):
    """
    Recebe itens brutos [{descricao, quantidade}] e o catálogo (lista de dicts
    com 'descricao' e 'preco'). Retorna lista alinhada por item:
      {descricao, quantidade, indice (int|None), confianca (0-100),
       status (CONFIRMADO|SUGESTAO|NAO_ENCONTRADO), justificativa}

    'hints' (opcional): para cada item, uma string com os melhores candidatos
    do motor determinístico, para ancorar a IA.
    'correcoes' (opcional): dict {texto_norm: produto} já ensinado pelos
    vendedores; entra como exemplos no prompt para a IA generalizar padrões.
    """
    if not itens_brutos:
        return []
    client, modelo = _cliente()
    cat_txt = _catalogo_texto(catalogo)

    pedidos = []
    for i, it in enumerate(itens_brutos):
        linha = f"{i}\tqtd={it.get('quantidade',1)}\t{it['descricao']}"
        if hints and i < len(hints) and hints[i]:
            linha += f"\t[candidatos prováveis: {hints[i]}]"
        pedidos.append(linha)
    pedidos_txt = "\n".join(pedidos)

    exemplos = _exemplos_correcoes(correcoes, itens_brutos)

    system = (
        "Você é um especialista em cotações da Pasquetti, distribuidora de tubos e "
        "conexões de ferro fundido. Sua tarefa: para cada item solicitado por um "
        "cliente (texto frequentemente abreviado, informal ou impreciso), encontrar "
        "o ÚNICO produto correspondente no catálogo fornecido.\n"
        + GLOSSARIO + exemplos +
        "\nRegras:\n"
        "- O DIÂMETRO precisa bater. Se o cliente diz 100mm, não escolha 150mm. "
        "Para reduções/tês, os dois números devem bater (ex.: 100x75).\n"
        "- O TIPO de produto precisa bater (não troque joelho por luva, nem curva "
        "BB por curva com flange).\n"
        "- Se o pedido for ambíguo entre tabelas/variações, escolha o mais provável "
        "e baixe a confiança.\n"
        "- Se nenhum produto servir, use indice = null.\n"
        "- confianca: 0-100. Use >=80 só quando tiver certeza; 60-79 quando provável; "
        "<60 quando duvidoso.\n"
        "Responda SOMENTE com JSON válido, sem texto fora do JSON."
    )

    user = (
        "CATÁLOGO (índice<TAB>descrição<TAB>preço):\n"
        f"{cat_txt}\n\n"
        "ITENS SOLICITADOS (índice<TAB>qtd<TAB>texto do cliente):\n"
        f"{pedidos_txt}\n\n"
        "Devolva um objeto JSON no formato:\n"
        '{"itens":[{"indice_item":0,"indice_catalogo":12,"confianca":92,'
        '"justificativa":"..."}, ...]}\n'
        "Inclua TODOS os itens solicitados, na ordem. indice_catalogo = null se nada servir."
    )

    msg = client.messages.create(
        model=modelo,
        max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    texto = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    dados = _extrair_json(texto)

    # Mapear resposta de volta, validando índices
    por_item = {}
    for r in (dados.get("itens") or []):
        try:
            ii = int(r["indice_item"])
        except (KeyError, TypeError, ValueError):
            continue
        por_item[ii] = r

    resultado = []
    for i, it in enumerate(itens_brutos):
        r = por_item.get(i, {})
        idx = r.get("indice_catalogo", None)
        conf = r.get("confianca", 0)
        try:
            conf = float(conf)
        except (TypeError, ValueError):
            conf = 0.0
        valido = isinstance(idx, int) and 0 <= idx < len(catalogo)
        if not valido:
            idx = None
        if idx is not None and conf >= 80:
            status = "CONFIRMADO"
        elif idx is not None and conf >= 60:
            status = "SUGESTÃO"
        else:
            status = "NÃO ENCONTRADO"
        resultado.append({
            "descricao":     it["descricao"],
            "quantidade":    it.get("quantidade", 1),
            "indice":        idx,
            "confianca":     conf,
            "status":        status,
            "justificativa": r.get("justificativa", ""),
        })
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# 2) LEITURA DE FONTES (foto / PDF) VIA VISÃO DO CLAUDE
# ════════════════════════════════════════════════════════════════════════════

_EXTRACAO_SYS = (
    "Você extrai listas de pedido de materiais hidráulicos a partir de imagens ou "
    "PDFs (fotos de WhatsApp, listas impressas, planilhas fotografadas). "
    "Identifique cada item e sua quantidade. Mantenha o texto do produto EXATAMENTE "
    "como escrito pelo cliente (não corrija nem interprete aqui). "
    "Responda SOMENTE com JSON: "
    '{"itens":[{"quantidade":10,"descricao":"texto do item"}, ...]}. '
    "Se não houver quantidade explícita, use 1."
)

def _extrair_lista(client, modelo, content_blocks):
    msg = client.messages.create(
        model=modelo,
        max_tokens=2000,
        system=_EXTRACAO_SYS,
        messages=[{"role": "user", "content": content_blocks}],
    )
    texto = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    dados = _extrair_json(texto)
    itens = []
    for r in (dados.get("itens") or []):
        desc = str(r.get("descricao", "")).strip()
        if not desc:
            continue
        try:
            qtd = float(r.get("quantidade", 1) or 1)
        except (TypeError, ValueError):
            qtd = 1
        itens.append({"descricao": desc, "quantidade": qtd})
    return itens

def extrair_itens_de_imagem(image_bytes, media_type="image/jpeg"):
    client, modelo = _cliente()
    b64 = base64.standard_b64encode(image_bytes).decode()
    blocks = [
        {"type": "image",
         "source": {"type": "base64", "media_type": media_type, "data": b64}},
        {"type": "text", "text": "Extraia a lista de itens e quantidades desta imagem."},
    ]
    return _extrair_lista(client, modelo, blocks)

def extrair_itens_de_pdf(pdf_bytes):
    client, modelo = _cliente()
    b64 = base64.standard_b64encode(pdf_bytes).decode()
    blocks = [
        {"type": "document",
         "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
        {"type": "text", "text": "Extraia a lista de itens e quantidades deste PDF."},
    ]
    return _extrair_lista(client, modelo, blocks)


# ════════════════════════════════════════════════════════════════════════════
# UTIL
# ════════════════════════════════════════════════════════════════════════════

def _extrair_json(texto):
    """Extrai o primeiro objeto JSON do texto (tolerante a cercas de código)."""
    if not texto:
        return {}
    t = texto.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        pass
    m = re.search(r"\{.*\}", t, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            return {}
    return {}
