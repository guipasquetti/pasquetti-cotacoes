"""
dados_supabase.py — Persistência de aprendizado no Supabase (projeto pasquetti-cotacoes).

Guarda, de forma permanente e compartilhada na web:
  • aprendizado_produtos  → correções de matching (texto do cliente → produto certo)
  • st_regras             → alíquota de ST por NCM + UF de destino

Credenciais (URL + chave publishable) vêm de, nesta ordem:
  1. st.secrets  (SUPABASE_URL / SUPABASE_KEY)  — usado na web
  2. variáveis de ambiente SUPABASE_URL / SUPABASE_KEY
  3. arquivo  config_supabase.json  na pasta do projeto  — usado no Mac

Tudo via REST (PostgREST), sem dependências extras além da biblioteca padrão.
Se não houver credenciais ou a internet falhar, as funções degradam para vazio
(o app continua funcionando, só sem aprendizado persistente).
"""
import os
import json
import unicodedata
import re
import urllib.request
import urllib.parse
import ssl
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config_supabase.json"
_CTX = ssl.create_default_context()


# ── Configuração ─────────────────────────────────────────────────────────────

def _from_secrets(nome):
    try:
        import streamlit as st
        return st.secrets.get(nome)
    except Exception:
        return None

def carregar_config():
    url = _from_secrets("SUPABASE_URL") or os.environ.get("SUPABASE_URL") or ""
    key = _from_secrets("SUPABASE_KEY") or os.environ.get("SUPABASE_KEY") or ""
    if (not url or not key) and CONFIG_PATH.exists():
        try:
            d = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            url = url or d.get("url", "")
            key = key or d.get("key", "")
        except Exception:
            pass
    return url.rstrip("/"), key

def disponivel():
    url, key = carregar_config()
    return bool(url and key)


# ── HTTP helpers ──────────────────────────────────────────────────────────────

def _req(method, path, params=None, body=None, extra_headers=None, timeout=8):
    url, key = carregar_config()
    if not (url and key):
        raise RuntimeError("Supabase não configurado.")
    full = f"{url}/rest/v1/{path}"
    if params:
        full += "?" + urllib.parse.urlencode(params)
    data = json.dumps(body).encode() if body is not None else None
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if extra_headers:
        headers.update(extra_headers)
    req = urllib.request.Request(full, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=timeout, context=_CTX) as r:
        txt = r.read().decode()
        return json.loads(txt) if txt.strip() else None


def normalizar(texto):
    """Normalização para casar texto de cliente (acentos/maiúsculas/espaços)."""
    if not texto:
        return ""
    t = "".join(c for c in unicodedata.normalize("NFD", str(texto))
                if unicodedata.category(c) != "Mn").lower()
    t = re.sub(r"[^\w\s./x-]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


# ── Aprendizado de produtos (correções) ───────────────────────────────────────

def listar_correcoes(tabela=None):
    """Retorna dict {texto_norm: produto_descricao} das correções salvas."""
    try:
        params = {"select": "texto_norm,produto_descricao,tabela"}
        rows = _req("GET", "aprendizado_produtos", params=params) or []
        out = {}
        for r in rows:
            if tabela and r.get("tabela") and r["tabela"] != tabela:
                continue
            out[r["texto_norm"]] = r["produto_descricao"]
        return out
    except Exception:
        return {}

def salvar_correcao(texto_cliente, produto_descricao, tabela="", vendedor=""):
    _req("POST", "aprendizado_produtos",
         body={"texto_cliente": texto_cliente,
               "texto_norm": normalizar(texto_cliente),
               "produto_descricao": produto_descricao,
               "tabela": tabela, "vendedor": vendedor},
         extra_headers={"Prefer": "return=minimal"})
    return True


# ── Produtos manuais (cadastrados fora das tabelas de preço) ──────────────────

def listar_produtos_manuais():
    """Lista os produtos cadastrados manualmente. Entram no catálogo em qualquer
    tabela e sobrevivem à regeneração do catalogo_produtos.json."""
    try:
        rows = _req("GET", "produtos_manuais",
                    params={"select": "descricao,preco,ncm,categoria,tabela,unidade"}) or []
        out = []
        for r in rows:
            try:
                preco = float(r.get("preco") or 0)
            except (TypeError, ValueError):
                preco = 0.0
            out.append({
                "descricao": r.get("descricao", "") or "",
                "preco": preco,
                "ncm": r.get("ncm", "") or "",
                "categoria": r.get("categoria", "Manual") or "Manual",
                "tabela": r.get("tabela", "MANUAL") or "MANUAL",
                "tamanho": "",
                "unidade": r.get("unidade", "UN") or "UN",
                "manual": True,
            })
        return out
    except Exception:
        return []

def salvar_produto_manual(descricao, preco, ncm="", categoria="Manual",
                          unidade="UN", vendedor=""):
    """Cadastra (ou atualiza) um produto manual. Idempotente por descricao_norm."""
    norm = normalizar(descricao)
    if not norm:
        return False
    try:
        _req("POST", "produtos_manuais",
             params={"on_conflict": "descricao_norm"},
             body={"descricao": descricao, "descricao_norm": norm,
                   "preco": float(preco or 0), "ncm": ncm or "",
                   "categoria": categoria or "Manual", "tabela": "MANUAL",
                   "unidade": unidade or "UN", "vendedor": vendedor or ""},
             extra_headers={"Prefer": "resolution=merge-duplicates,return=minimal"})
        return True
    except Exception:
        return False

def remover_produto_manual(descricao):
    """Remove um produto manual pela descrição."""
    norm = normalizar(descricao)
    if not norm:
        return False
    try:
        _req("DELETE", "produtos_manuais",
             params={"descricao_norm": f"eq.{norm}"},
             extra_headers={"Prefer": "return=minimal"})
        return True
    except Exception:
        return False


# ── Itens que não trabalhamos (não fornecemos) ────────────────────────────────

def listar_nao_trabalhados():
    """Retorna um set com os texto_norm dos itens marcados como 'não trabalhamos'."""
    try:
        rows = _req("GET", "itens_nao_trabalhados", params={"select": "texto_norm"}) or []
        return {r["texto_norm"] for r in rows if r.get("texto_norm")}
    except Exception:
        return set()

def salvar_nao_trabalhado(texto_cliente, vendedor=""):
    """Marca (idempotente) um item como 'não trabalhamos' para próximas cotações."""
    norm = normalizar(texto_cliente)
    if not norm:
        return False
    _req("POST", "itens_nao_trabalhados",
         params={"on_conflict": "texto_norm"},
         body={"texto_cliente": texto_cliente, "texto_norm": norm, "vendedor": vendedor},
         extra_headers={"Prefer": "resolution=merge-duplicates,return=minimal"})
    return True

def remover_nao_trabalhado(texto_cliente):
    """Desfaz a marcação 'não trabalhamos' de um item."""
    norm = normalizar(texto_cliente)
    if not norm:
        return False
    _req("DELETE", "itens_nao_trabalhados",
         params={"texto_norm": f"eq.{norm}"},
         extra_headers={"Prefer": "return=minimal"})
    return True


# ── Regras de ST (NCM + UF → alíquota %) ──────────────────────────────────────

# UFs com alíquota interestadual de 12% saindo de SP (demais = 7%; SP = interna)
UF_INTER_12 = {"MG", "PR", "RJ", "RS", "SC"}

# Alíquota INTERNA do destino por UF — transcrita dos prints de DIFAL do ERP
# (22/06/2026). É constante por estado (não varia por NCM). Confirmar c/ contador.
ALIQUOTA_INTERNA_PADRAO = {
    "AC": 22.5, "AL": 19.0, "AM": 20.0, "AP": 18.0, "BA": 20.5, "CE": 20.0,
    "DF": 20.0, "ES": 17.0, "GO": 19.0, "MA": 23.0, "MT": 17.0, "MS": 17.0,
    "MG": 18.0, "PA": 19.0, "PB": 20.0, "PR": 19.5, "PE": 20.5, "PI": 22.5,
    "RJ": 20.0, "RN": 20.0, "RO": 19.5, "RR": 20.0, "RS": 17.0, "SC": 17.0,
    "SP": 18.0, "SE": 20.0, "TO": 20.0,
}

# FCP (Fundo de Combate à Pobreza) somado ao DIFAL, por UF (dos prints).
FCP_DIFAL = {"AL": 1.0, "RJ": 2.0}

# Alíquota interestadual usada no DIFAL (valores do ERP): 12% p/ S/SE;
# MA = 8%, MS = 10% (exceções do ERP); demais 7%. SP = None (interna).
# Obs.: nas válvulas 8481.* o ES aparece com 4% (bem importado) — não tratado aqui.
DIFAL_INTER = {
    "MG": 12.0, "PR": 12.0, "RJ": 12.0, "RS": 12.0, "SC": 12.0,
    "MA": 8.0, "MS": 10.0,
}


def aliquota_interestadual(uf):
    """% de ICMS interestadual saindo de SP para a UF (ST). SP -> None."""
    uf = (uf or "").strip().upper()
    if uf == "SP":
        return None
    return 12.0 if uf in UF_INTER_12 else 7.0


def difal_interestadual(uf):
    """% interestadual usada no DIFAL (inclui exceções MA=8, MS=10). SP -> None."""
    uf = (uf or "").strip().upper()
    if uf == "SP":
        return None
    return DIFAL_INTER.get(uf, 7.0)


def listar_regras_st():
    """Retorna dict {(ncm, uf): {cst, mva, icms_interno, aliquota}}."""
    try:
        rows = _req("GET", "st_regras",
                    params={"select": "ncm,uf,cst,mva,icms_interno,aliquota"}) or []
        out = {}
        for r in rows:
            k = (str(r["ncm"]).strip(), str(r["uf"]).strip().upper())
            out[k] = {
                "cst": int(r.get("cst") or 6),
                "mva": float(r.get("mva") or 0),
                "icms_interno": float(r.get("icms_interno") or r.get("aliquota") or 0),
                "aliquota": float(r.get("aliquota") or 0),
            }
        return out
    except Exception:
        return {}


def calcular_st_difal(preco_unit, ncm, uf, regras=None, consumidor_final=False):
    """ST/DIFAL por unidade. Revenda(contribuinte)->ST; consumidor final->DIFAL.
    base ST = preco*(1+MVA); ST = base*interno - preco*interestadual.
    DIFAL = preco*(interno - interestadual)."""
    uf = (uf or "").strip().upper()
    if uf == "SP":                          # venda dentro de SP
        return 0.0
    if regras is None:
        regras = listar_regras_st()
    reg = regras.get((str(ncm).strip(), uf))
    if consumidor_final:
        # DIFAL = preco * (interno + FCP - interestadual) / 100
        interno = ALIQUOTA_INTERNA_PADRAO.get(uf, 0) or (reg or {}).get("icms_interno", 0)
        if interno <= 0:
            return 0.0
        fcp = FCP_DIFAL.get(uf, 0.0)
        inter_dif = difal_interestadual(uf) or 0.0
        return round(max(preco_unit * (interno + fcp - inter_dif) / 100.0, 0.0), 4)
    inter = aliquota_interestadual(uf)
    # ST (revenda / contribuinte) — só quando há regra com ST (CST 4)
    if not reg or reg.get("cst") != 4 or reg.get("icms_interno", 0) <= 0:
        return 0.0
    mva = reg["mva"]; interno = reg["icms_interno"]
    base = preco_unit * (1 + mva / 100.0)
    st = base * (interno / 100.0) - preco_unit * (inter / 100.0)
    return round(max(st, 0.0), 4)


def salvar_regra_st(ncm, uf, aliquota, observacao="", cst=4, mva=0.0, icms_interno=None):
    """Insere/atualiza (upsert) a regra de ST para um NCM + UF."""
    if icms_interno is None:
        icms_interno = aliquota
    _req("POST", "st_regras",
         params={"on_conflict": "ncm,uf"},
         body={"ncm": str(ncm).strip(), "uf": str(uf).strip().upper(),
               "aliquota": float(aliquota), "cst": int(cst),
               "mva": float(mva), "icms_interno": float(icms_interno),
               "observacao": observacao},
         extra_headers={"Prefer": "resolution=merge-duplicates,return=minimal"})
    return True


# ── Teste rápido de conexão ───────────────────────────────────────────────────

def testar_conexao():
    try:
        _req("GET", "aprendizado_produtos", params={"select": "id", "limit": "1"})
        return True, "Conexão OK"
    except Exception as e:
        return False, str(e)


# ── Condições de pagamento ─────────────────────────────────────────────────

def listar_condicoes_pagamento():
    """Lista os textos das condições de pagamento cadastradas."""
    try:
        rows = _req("GET", "condicoes_pagamento",
                    params={"select": "texto", "order": "texto.asc"}) or []
        return [r["texto"] for r in rows if r.get("texto")]
    except Exception:
        return []

def salvar_condicao_pagamento(texto):
    """Cadastra uma nova condição de pagamento (idempotente por texto_norm)."""
    _req("POST", "condicoes_pagamento",
         params={"on_conflict": "texto_norm"},
         body={"texto": texto.strip(), "texto_norm": normalizar(texto)},
         extra_headers={"Prefer": "resolution=merge-duplicates,return=minimal"})
    return True


# ── Clientes (cadastro para próximas cotações) ─────────────────────────────

def _so_digitos(t):
    return "".join(ch for ch in str(t or "") if ch.isdigit())

def buscar_clientes(termo, limite=8):
    """Busca clientes por prefixo de CNPJ (dígitos) ou por nome (ilike)."""
    termo = (termo or "").strip()
    if not termo:
        return []
    dig = _so_digitos(termo)
    try:
        if dig:
            params = {"select": "*", "cnpj": f"like.{dig}*", "limit": str(limite)}
        else:
            params = {"select": "*", "nome": f"ilike.*{termo}*", "limit": str(limite)}
        return _req("GET", "clientes", params=params) or []
    except Exception:
        return []

def salvar_cliente(cnpj, nome, fantasia="", endereco="", telefone="", email="", uf=""):
    """Insere/atualiza (upsert por CNPJ) um cliente. CNPJ guardado só com dígitos."""
    dig = _so_digitos(cnpj)
    if not dig or not (nome or "").strip():
        return False
    _req("POST", "clientes", params={"on_conflict": "cnpj"},
         body={"cnpj": dig, "nome": nome, "fantasia": fantasia, "endereco": endereco,
               "telefone": telefone, "email": email, "uf": uf},
         extra_headers={"Prefer": "resolution=merge-duplicates,return=minimal"})
    return True
