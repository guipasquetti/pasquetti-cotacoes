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

def listar_regras_st():
    """Retorna dict {(ncm, uf): aliquota_float}."""
    try:
        rows = _req("GET", "st_regras", params={"select": "ncm,uf,aliquota"}) or []
        return {(str(r["ncm"]).strip(), str(r["uf"]).strip().upper()): float(r["aliquota"])
                for r in rows}
    except Exception:
        return {}

def salvar_regra_st(ncm, uf, aliquota, observacao=""):
    """Insere/atualiza (upsert) a alíquota de ST para um NCM + UF."""
    _req("POST", "st_regras",
         params={"on_conflict": "ncm,uf"},
         body={"ncm": str(ncm).strip(), "uf": str(uf).strip().upper(),
               "aliquota": float(aliquota), "observacao": observacao},
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
