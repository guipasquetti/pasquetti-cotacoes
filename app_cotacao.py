"""
app_cotacao.py — Pasquetti Gerador de Cotações
Iniciar: streamlit run app_cotacao.py
"""

import streamlit as st
import json, re, difflib, unicodedata, datetime, io, base64, os
from pathlib import Path
import openpyxl

# Camada de IA (opcional). Se a lib/chave não existirem, o app segue só no
# motor determinístico.
try:
    import matcher_ia
except Exception:
    matcher_ia = None
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuração da página ─────────────────────────────────────────────────
st.set_page_config(
    page_title="Pasquetti — Cotações",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

SCRIPT_DIR   = Path(__file__).parent
CATALOG_JSON = SCRIPT_DIR / "catalogo_produtos.json"
LOGO_PATH    = SCRIPT_DIR / "logo_pasquetti.png"

# ── Web: ponte de segredos + tela de senha ──────────────────────────────────
# Em produção (Streamlit Cloud) a chave de IA e a senha vêm de st.secrets.
# Localmente, sem secrets, nada disso atrapalha (o app abre direto).
def _secret(nome):
    try:
        return st.secrets.get(nome)
    except Exception:
        return None

# Disponibiliza a chave de IA do servidor para o matcher_ia (via variável de ambiente)
_chave_servidor = _secret("ANTHROPIC_API_KEY")
if _chave_servidor and not os.environ.get("ANTHROPIC_API_KEY"):
    os.environ["ANTHROPIC_API_KEY"] = _chave_servidor

def _checar_senha():
    """Se houver APP_PASSWORD nos segredos, exige senha. Sem senha definida, libera."""
    senha = _secret("APP_PASSWORD")
    if not senha:
        return  # uso local / sem proteção configurada
    if st.session_state.get("_auth_ok"):
        return
    st.markdown("### 🔒 Pasquetti — Sistema de Cotações")
    st.caption("Acesso restrito. Informe a senha para continuar.")
    pw = st.text_input("Senha", type="password", label_visibility="collapsed",
                       placeholder="Senha de acesso")
    if st.button("Entrar", type="primary"):
        if pw == senha:
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

_checar_senha()

# ── Cores da marca ──────────────────────────────────────────────────────────
NAVY   = "#1B3065"   # azul-marinho Pasquetti
COPPER = "#C47A3A"   # cobre/laranja Pasquetti
NAVY_L = "#2a4a8a"   # azul claro hover
COPPER_L = "#e09050" # cobre claro

# ── Logo como SVG inline (ícone "P" estilizado com cano) ──────────────────
LOGO_SVG = """
<svg width="52" height="52" viewBox="0 0 52 52" fill="none" xmlns="http://www.w3.org/2000/svg">
  <!-- Fundo do ícone -->
  <rect width="52" height="52" rx="8" fill="#1B3065"/>
  <!-- Corpo do P (cano vertical) -->
  <rect x="12" y="8" width="9" height="36" rx="3" fill="white"/>
  <!-- Arco superior (curva do cano) -->
  <path d="M21 8 Q38 8 38 20 Q38 32 21 32" stroke="white" stroke-width="9" fill="none" stroke-linecap="round"/>
  <!-- Detalhe cobre - conector -->
  <rect x="9" y="34" width="15" height="8" rx="2" fill="#C47A3A"/>
  <!-- Detalhe cobre - flange superior -->
  <rect x="9" y="6" width="15" height="6" rx="2" fill="#C47A3A"/>
</svg>
"""

# Carregar logo real se existir
def get_logo_b64():
    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

# ════════════════════════════════════════════════════════════════════════════
# CSS — Identidade Visual Pasquetti
# ════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  html, body, .stApp {{
    font-family: 'Inter', sans-serif !important;
    background: #f0f3f9 !important;
  }}

  /* ── Sidebar ── */
  [data-testid="stSidebar"] {{
    background: {NAVY} !important;
    border-right: 3px solid {COPPER} !important;
  }}
  [data-testid="stSidebar"] * {{ color: #d8e4ff !important; }}
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] strong {{ color: #ffffff !important; }}
  [data-testid="stSidebar"] .stTextInput input,
  [data-testid="stSidebar"] .stSelectbox select {{
    background: rgba(255,255,255,0.1) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    border-radius: 6px !important;
  }}
  [data-testid="stSidebar"] .stTextInput input::placeholder {{ color: rgba(255,255,255,0.45) !important; }}
  [data-testid="stSidebar"] label {{ color: #b8cfff !important; font-size:12px !important; font-weight:500 !important; }}
  [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.15) !important; }}

  /* ── Header ── */
  .pq-header {{
    background: linear-gradient(135deg, {NAVY} 0%, #243d7a 100%);
    border-radius: 12px;
    padding: 20px 28px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 16px;
    border-bottom: 4px solid {COPPER};
  }}
  .pq-title {{ color: #fff; font-size: 22px; font-weight: 700; margin: 0; letter-spacing: -0.3px; }}
  .pq-sub   {{ color: #b8cfff; font-size: 12px; margin: 3px 0 0 0; }}
  .pq-tag   {{
    background: {COPPER}; color: white;
    font-size: 10px; font-weight: 600;
    padding: 3px 8px; border-radius: 20px;
    text-transform: uppercase; letter-spacing: .5px;
    margin-left: auto;
  }}

  /* ── Tabs ── */
  .stTabs [data-baseweb="tab-list"] {{
    background: white !important;
    border-radius: 8px !important;
    padding: 4px !important;
    border: 1px solid #dce3f0 !important;
    gap: 2px !important;
  }}
  .stTabs [data-baseweb="tab"] {{
    border-radius: 6px !important;
    padding: 8px 16px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #5a6a8a !important;
  }}
  .stTabs [aria-selected="true"] {{
    background: {NAVY} !important;
    color: white !important;
  }}

  /* ── Botão principal ── */
  div.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, {COPPER}, #e09050) !important;
    color: white !important;
    border: none !important;
    padding: 13px 28px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    width: 100% !important;
    letter-spacing: 0.3px !important;
  }}
  div.stButton > button[kind="primary"]:hover {{ opacity: 0.9 !important; }}

  /* ── Botão secundário (CNPJ) ── */
  div.stButton > button:not([kind="primary"]) {{
    background: {NAVY} !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    padding: 6px 14px !important;
  }}

  /* ── Cards de resultado ── */
  .pq-metric {{
    background: white;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    border: 1px solid #dce3f0;
    border-top: 4px solid {NAVY};
  }}
  .pq-metric.copper {{ border-top-color: {COPPER}; }}
  .pq-metric.red    {{ border-top-color: #c0392b; }}
  .pq-metric.green  {{ border-top-color: #27ae60; }}
  .pq-metric-val  {{ font-size: 28px; font-weight: 700; color: {NAVY}; }}
  .pq-metric-val.copper {{ color: {COPPER}; }}
  .pq-metric-val.red    {{ color: #c0392b; }}
  .pq-metric-val.green  {{ color: #27ae60; }}
  .pq-metric-lbl  {{ font-size: 11px; color: #7a8aaa; margin-top: 4px; font-weight: 500; }}

  /* ── Tabela de itens ── */
  .pq-table {{ width:100%; border-collapse:collapse; font-size:12px; }}
  .pq-table th {{
    background: {NAVY}; color: white;
    padding: 9px 12px; text-align: left;
    font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing:.4px;
  }}
  .pq-table td {{ padding: 9px 12px; border-bottom: 1px solid #eef1f8; vertical-align:middle; }}
  .pq-table tr:nth-child(even) td {{ background: #f7f9ff; }}
  .pq-table tr:hover td {{ background: #eef3ff; }}
  .badge {{ border-radius:4px; padding:2px 8px; font-size:10px; font-weight:700; }}
  .badge-a  {{ background:#e8f5e9; color:#2e7d32; }}
  .badge-m  {{ background:#fff8e1; color:#c77a00; }}
  .badge-s  {{ background:#fce4ec; color:#b71c1c; }}

  /* ── Seção de sugestões ── */
  .pq-suggest {{
    background: #fffbf3;
    border: 1px solid #f0d090;
    border-left: 4px solid {COPPER};
    border-radius: 8px;
    padding: 14px 18px;
    margin-top: 12px;
    font-size: 12px;
  }}
  .pq-suggest-title {{ font-weight: 700; color: {COPPER}; margin-bottom: 8px; font-size:13px; }}

  /* ── Seção não encontrados ── */
  .pq-notfound {{
    background: #fff5f5;
    border: 1px solid #f0c0c0;
    border-left: 4px solid #c0392b;
    border-radius: 8px;
    padding: 14px 18px;
    margin-top: 12px;
    font-size: 12px;
  }}
  .pq-notfound-title {{ font-weight: 700; color: #c0392b; margin-bottom: 8px; font-size:13px; }}

  /* ── Upload area ── */
  [data-testid="stFileUploader"] {{
    background: white !important;
    border: 2px dashed #b0bdd8 !important;
    border-radius: 10px !important;
    padding: 20px !important;
  }}

  /* ── Seção títulos ── */
  .pq-section {{
    font-size: 14px; font-weight: 700;
    color: {NAVY};
    border-bottom: 2px solid {COPPER};
    padding-bottom: 6px;
    margin: 16px 0 12px 0;
    display: inline-block;
  }}

  /* ── CNPJ strip ── */
  .cnpj-strip {{
    background: white;
    border: 1px solid #dce3f0;
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 10px;
    font-size: 11px;
  }}
  .cnpj-label {{ color: #7a8aaa; font-size: 10px; text-transform:uppercase; letter-spacing:.4px; }}

  #MainMenu, footer, [data-testid="stToolbar"] {{ visibility: hidden !important; }}

  /* ── Campos de entrada: contraste no fundo claro (regra global) ── */
  .stApp input, .stApp textarea {{
    color: {NAVY} !important;
    background: #ffffff !important;
  }}
  .stApp input::placeholder,
  .stApp textarea::placeholder {{
    color: #8a97b5 !important;   /* cinza-azulado legível sobre branco */
    opacity: 1 !important;
  }}

  /* ── Reafirma os campos da BARRA LATERAL (fundo escuro) ── */
  [data-testid="stSidebar"] input,
  [data-testid="stSidebar"] textarea {{
    color: #ffffff !important;
    background: rgba(255,255,255,0.10) !important;
  }}
  [data-testid="stSidebar"] input::placeholder,
  [data-testid="stSidebar"] textarea::placeholder {{
    color: rgba(255,255,255,0.55) !important;
    opacity: 1 !important;
  }}

  /* ── Texto dos expanders ("Ver itens lidos") legível no fundo claro ── */
  [data-testid="stExpander"] summary,
  [data-testid="stExpander"] p,
  [data-testid="stExpander"] li,
  [data-testid="stExpander"] span,
  [data-testid="stExpander"] div {{
    color: {NAVY} !important;
  }}
  /* mantém os expanders da barra lateral (fundo escuro) com texto claro */
  [data-testid="stSidebar"] [data-testid="stExpander"] summary,
  [data-testid="stSidebar"] [data-testid="stExpander"] p,
  [data-testid="stSidebar"] [data-testid="stExpander"] li,
  [data-testid="stSidebar"] [data-testid="stExpander"] span,
  [data-testid="stSidebar"] [data-testid="stExpander"] div {{
    color: #d8e4ff !important;
  }}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# MOTOR DE CORRESPONDÊNCIA
# ════════════════════════════════════════════════════════════════════════════

SINONIMOS = {
    r'\bjf\b':'joelho fofo', r'\bjff\b':'joelho fofo',
    r'\bjoelho\s*fofo\b':'joelho fofo', r'\bjoelho\b':'joelho fofo',
    r'\bcurva\s*fofo\b':'joelho fofo', r'\bcurva\b':'joelho fofo',
    r'\btee\s*fofo\b':'tee fofo', r'\btê\b':'tee fofo',
    r'\btee\b':'tee fofo', r'\bte\b':'tee fofo',
    r'\bluva\s*fofo\b':'luva fofo', r'\bluva\s*hl\b':'luva fofo',
    r'\bljf\b':'luva fofo',
    r'\bbucha\s*red\b':'bucha reducao', r'\bbucha\b':'bucha reducao',
    r'\bred\s*conc\b':'reducao', r'\bred\b':'reducao',
    r'\bpl\s*cega\b':'placa cega fofo', r'\bplaca\s*cega\b':'placa cega fofo',
    r'\bjuncao\b':'juncao fofo', r'\bjunc\b':'juncao fofo',
    r'\banel\s*borr\b':'anel borr', r'\banel\s*borracha\b':'anel borr',
    r'\banel\b':'anel borr',
    r'\bjunta\s*rap\b':'junta rapid', r'\bjr\b':'junta rapid jrsmu',
    r'\bjrsmu\b':'junta rapid jrsmu',
    r'\btampao\b':'tampao borr',
    r'\bgrelha\b':'grelha hemisf fofo',
    r'\bvalvula\b':'valv', r'\bvalv\s*gav\b':'valv gav',
    r'\bvalv\s*ret\b':'valv ret', r'\bvalv\b':'valv',
    r'\bhidrante\b':'hidrante col',
    r'\bventosa\b':'ventosa', r'\bvsr\b':'ventosa vsr',
    r'\bvsf\b':'ventosa vsf', r'\bvtf\b':'ventosa vtf',
    r'\bcolar\b':'colar de tomada',
    r'\banel\s*trav\b':'anel trav interno',
    r'\bjtd\b':'junta desmontagem', r'\bjg\b':'junta gibault',
    r'\bul\b':'ultralink', r'\buq\b':'ultraquik',
    r'\bflange\s*av\b':'flange avulsa', r'\bflange\s*cego\b':'flange cego',
    r'\btoco\b':'toco c/fl', r'\bluva\s*cr\b':'luva cr jgs',
    r'\bluva\s*jm\b':'luva junta mecanica', r'\bluva\s*gt\b':'luva gr tol',
    r'\bluva\s*trip\b':'luva tripartida', r'\bcruz\b':'cruzeta',
    r'\bbb\b':'b/b', r'\bpp\b':'p/p', r'\bpb\b':'p/b',
    r'\bbfl\b':'b/fl', r'\bcfl\b':'c/fl', r'\bpfl\b':'p/fl',
    r'\b90\s*gr(?:aus?)?\b':'90', r'\b45\s*gr(?:aus?)?\b':'45',
    r'\b90°':'90', r'\b45°':'45',
    r'\b(\d+)\s*mm\b':r'\1mm', r'\b(\d+)\s*x\s*(\d+)\b':r'\1x\2',
    r'\bhl\b':'hl', r'\bjgs\b':'jgs', r'\bsmu\b':'smu', r'\bsme\b':'sme',
}

# Palavras-chave de tipo de produto para bônus de categoria
TIPO_KEYWORDS = {
    'joelho fofo': ['joelho','jf','jff','curva'],
    'tee fofo':    ['tee','te'],
    'luva fofo':   ['luva'],
    'anel borr':   ['anel','borracha'],
    'tubo':        ['tubo','tub'],
    'junta rapid': ['junta','jr','jrsmu'],
    'bucha':       ['bucha','reducao','red'],
    'placa cega':  ['placa','cega'],
    'valv gav':    ['valvula','valv','gav'],
    'valv ret':    ['valv','ret'],
    'hidrante':    ['hidrante'],
    'ventosa':     ['ventosa','vsr','vsf','vtf'],
    'flange':      ['flange'],
}

ANGULOS = {"90", "45", "22", "11", "15"}

def remover_acentos(t):
    return ''.join(c for c in unicodedata.normalize('NFD', t)
                   if unicodedata.category(c) != 'Mn')

def normalizar(texto):
    if not texto: return ""
    t = remover_acentos(str(texto)).lower()
    t = re.sub(r'[^\w\s./x-]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def expandir(texto):
    t = normalizar(texto)
    for p, s in SINONIMOS.items():
        t = re.sub(p, s, t, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', t).strip()

def extrair_diametros(texto):
    t = normalizar(texto)
    result = set()
    for m in re.finditer(r'(\d+)x(\d+)', t): result.add(m.group(0))
    for m in re.finditer(r'(\d+)\s*mm', t):
        if m.group(1) not in ANGULOS: result.add(m.group(1))
    for m in re.finditer(r'\b(\d+)\b', t):
        n = m.group(1)
        if n not in ANGULOS and 40 <= int(n) <= 600: result.add(n)
    return result

def bonus_tipo(query_exp, cat_desc):
    """Bônus se query e produto são do mesmo tipo."""
    for tipo, keywords in TIPO_KEYWORDS.items():
        tipo_norm = normalizar(tipo)
        if any(k in query_exp for k in keywords):
            if tipo_norm in cat_desc:
                return 12
    return 0

def score_ts(a, b):
    ta = " ".join(sorted(a.split()))
    tb = " ".join(sorted(b.split()))
    return difflib.SequenceMatcher(None, ta, tb).ratio() * 100

def score_set(a, b):
    sa, sb = set(a.split()), set(b.split())
    inter = sa & sb
    t0 = " ".join(sorted(inter))
    t1 = " ".join(sorted(inter | (sa - inter)))
    t2 = " ".join(sorted(inter | (sb - inter)))
    return max(difflib.SequenceMatcher(None, t0, t1).ratio(),
               difflib.SequenceMatcher(None, t0, t2).ratio(),
               difflib.SequenceMatcher(None, t1, t2).ratio()) * 100

def buscar(descricao, catalogo):
    """
    Retorna (produto, score, confiança).
    CONFIRMADO: score ≥ 80%  → entra na cotação automaticamente
    SUGESTÃO:   score 60-79% → listado para revisão manual, NÃO entra na cotação
    NÃO ENCONTRADO: < 60%
    """
    qo = normalizar(descricao)
    qe = expandir(descricao)
    best_prod, best_score = None, 0.0

    for p in catalogo:
        d = p["_norm"]
        # Múltiplas estratégias
        s = max(
            score_ts(qe, d),
            score_set(qe, d),
            score_ts(qo, d) * 0.9,
            score_set(qo, d) * 0.9,
        )
        # Bônus por tipo de produto
        s += bonus_tipo(qe, d)
        # Bônus/penalidade por diâmetro
        dq, dp = extrair_diametros(qe), extrair_diametros(d)
        if dq and dp:
            s += 18 if dq & dp else -28
        s = min(100, max(0, s))

        if s > best_score:
            best_score, best_prod = s, p

    if best_score >= 80:
        return best_prod, best_score, "CONFIRMADO"
    elif best_score >= 60:
        return best_prod, best_score, "SUGESTÃO"
    return best_prod, best_score, "NÃO ENCONTRADO"


def processar(itens_brutos, catalogo):
    confirmados, sugestoes, nao_encontrados = [], [], []
    for item in itens_brutos:
        p, score, conf = buscar(item["descricao"], catalogo)
        qtd = item["quantidade"]
        base = {**item, "score": score, "conf": conf}
        if p:
            base.update({"produto": p["descricao"], "ncm": p.get("ncm",""),
                          "preco": p["preco"], "total": qtd * p["preco"]})
        if conf == "CONFIRMADO":
            confirmados.append(base)
        elif conf == "SUGESTÃO":
            sugestoes.append(base)
        else:
            nao_encontrados.append(base)
    return confirmados, sugestoes, nao_encontrados


def _hint_deterministico(descricao, catalogo, n=3):
    """Top-N candidatos do motor determinístico, como dica para a IA."""
    qe = expandir(descricao)
    scored = []
    for p in catalogo:
        s = max(score_ts(qe, p["_norm"]), score_set(qe, p["_norm"]))
        scored.append((s, p["descricao"]))
    scored.sort(reverse=True)
    return " | ".join(d for _, d in scored[:n])


def processar_hibrido(itens_brutos, catalogo, usar_ia):
    """
    Motor determinístico + (opcional) camada de IA.
    Quando a IA está ligada, ela é a fonte primária; o determinístico serve de
    dica e de rede de segurança (se a IA falhar ou perder um match exato fácil).
    Retorna (confirmados, sugestoes, nao_encontrados) no mesmo formato de processar().
    """
    # 1) Determinístico para todos (sempre)
    det = []
    for item in itens_brutos:
        p, score, conf = buscar(item["descricao"], catalogo)
        det.append({"item": item, "prod": p, "score": score, "conf": conf})

    ia_res = None
    if usar_ia and matcher_ia is not None:
        try:
            hints = [_hint_deterministico(it["descricao"], catalogo) for it in itens_brutos]
            ia_res = matcher_ia.interpretar_itens(itens_brutos, catalogo, hints=hints)
        except Exception as e:
            st.warning(f"IA indisponível nesta cotação ({e}). Usando apenas o motor determinístico.")
            ia_res = None

    confirmados, sugestoes, nao = [], [], []
    for i, item in enumerate(itens_brutos):
        qtd = item["quantidade"]
        d = det[i]

        # Resultado determinístico como padrão
        prod, score, conf, fonte, just = d["prod"], d["score"], d["conf"], "auto", ""

        if ia_res is not None:
            r = ia_res[i]
            ia_prod = catalogo[r["indice"]] if r["indice"] is not None else None
            # IA é primária quando achou algo
            if ia_prod is not None:
                prod, score, conf = ia_prod, r["confianca"], r["status"]
                fonte, just = "IA", r.get("justificativa", "")
            # Rede de segurança: IA não achou, mas determinístico tem match exato forte
            elif d["conf"] == "CONFIRMADO":
                prod, score, conf, fonte = d["prod"], d["score"], "CONFIRMADO", "auto"
            else:
                prod, score, conf = None, r["confianca"], "NÃO ENCONTRADO"
                just = r.get("justificativa", "")

        base = {"descricao": item["descricao"], "quantidade": qtd,
                "score": score, "conf": conf, "fonte": fonte, "justificativa": just}
        if prod:
            base.update({"produto": prod["descricao"], "ncm": prod.get("ncm", ""),
                         "preco": prod["preco"], "total": qtd * prod["preco"]})
        if conf == "CONFIRMADO":
            confirmados.append(base)
        elif conf == "SUGESTÃO":
            sugestoes.append(base)
        else:
            nao.append(base)
    return confirmados, sugestoes, nao


# ════════════════════════════════════════════════════════════════════════════
# BUSCA CNPJ (BrasilAPI — Receita Federal)
# ════════════════════════════════════════════════════════════════════════════

def buscar_cnpj(cnpj_raw: str):
    """Consulta dados da empresa pelo CNPJ via BrasilAPI (Receita Federal)."""
    try:
        import urllib.request, ssl
        cnpj = re.sub(r'\D', '', cnpj_raw)
        if len(cnpj) != 14:
            return None, "CNPJ deve ter 14 dígitos."
        url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"User-Agent": "PasquettiCotacoes/1.0"})
        with urllib.request.urlopen(req, timeout=8, context=ctx) as r:
            data = json.loads(r.read().decode())
        tel = data.get("ddd_telefone_1","").strip()
        if tel and len(tel) >= 8:
            tel = f"({tel[:2]}) {tel[2:]}" if len(tel) >= 10 else tel
        end_parts = [
            data.get("logradouro",""), data.get("numero",""),
            data.get("complemento",""), data.get("bairro",""),
            data.get("municipio",""), data.get("uf",""),
        ]
        endereco = " ".join(p for p in end_parts if p and p.strip())
        cep = data.get("cep","")
        if cep: endereco += f" — CEP: {cep}"
        return {
            "nome":         data.get("razao_social",""),
            "fantasia":     data.get("nome_fantasia",""),
            "cnpj_fmt":     f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}",
            "endereco":     endereco,
            "telefone":     tel,
            "email":        data.get("email",""),
            "situacao":     data.get("descricao_situacao_cadastral",""),
            "atividade":    data.get("cnae_fiscal_descricao",""),
        }, None
    except Exception as e:
        msg = str(e)
        if "404" in msg:
            return None, "CNPJ não encontrado na Receita Federal."
        if "timeout" in msg.lower() or "urlopen" in msg.lower():
            return None, "Sem conexão com a internet. Preencha os dados manualmente."
        return None, f"Erro ao consultar: {msg}"


# ════════════════════════════════════════════════════════════════════════════
# PARSING DE ENTRADA
# ════════════════════════════════════════════════════════════════════════════

def extrair_qtd_desc(texto):
    texto = texto.strip()
    m = re.match(r'^(\d+(?:[.,]\d+)?)\s*(?:x|un|pç|pc|pcs|und|unid)?\s+(.+)', texto, re.I)
    if m:
        try: qtd = float(m.group(1).replace(",","."))
        except: qtd = 1
        return qtd, m.group(2).strip()
    m = re.search(r'\s+(\d+(?:[.,]\d+)?)\s*(?:x|un|pç|pc|pcs|und|unid)?$', texto, re.I)
    if m:
        try: qtd = float(m.group(1).replace(",","."))
        except: qtd = 1
        return qtd, texto[:m.start()].strip()
    return 1, texto

def parse_texto(texto):
    itens = []
    for linha in texto.splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#"): continue
        qtd, desc = extrair_qtd_desc(linha)
        if desc and desc.lower() not in ("none","nan"): itens.append({"descricao": desc, "quantidade": qtd})
    return itens

def parse_xlsx_bytes(data):
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    header_row = col_desc = col_qtd = None
    for i, row in enumerate(ws.iter_rows(max_row=10, values_only=True), 1):
        row_s = [str(c).lower().strip() if c else "" for c in row]
        for j, cell in enumerate(row_s):
            if any(k in cell for k in ["produto","descri","item","material","especif"]):
                col_desc = j; header_row = i
            if any(k in cell for k in ["qtd","quant","qt ","pcs","unid"]):
                col_qtd = j
        if header_row: break
    if not header_row: col_desc, col_qtd, header_row = 0, 1, 1
    itens = []
    for row in ws.iter_rows(min_row=header_row+1, values_only=True):
        desc = str(row[col_desc]).strip() if col_desc is not None and col_desc < len(row) and row[col_desc] else ""
        try: qtd = float(row[col_qtd]) if col_qtd is not None and col_qtd < len(row) and row[col_qtd] else 1
        except: qtd = 1
        if desc and desc.lower() not in ("none","nan",""):
            itens.append({"descricao": desc, "quantidade": qtd})
    return itens

def parse_imagem_bytes(data):
    try:
        import pytesseract
        from PIL import Image
        img = Image.open(io.BytesIO(data))
        texto = pytesseract.image_to_string(img, lang="por")
        return parse_texto(texto)
    except ImportError:
        st.error("OCR não disponível. Instale: `pip install pytesseract pillow` e o Tesseract OCR.")
        return []


# ════════════════════════════════════════════════════════════════════════════
# CATÁLOGO
# ════════════════════════════════════════════════════════════════════════════

@st.cache_data
def carregar_catalogo(tabela):
    if not CATALOG_JSON.exists(): return []
    with open(CATALOG_JSON, encoding="utf-8") as f: todos = json.load(f)
    mapa = {"consumo":["HL_CONSUMO"],"revenda":["HL_REVENDA"],
            "pressao":["PRESSAO_JGS"],"todos":["HL_CONSUMO","HL_REVENDA","PRESSAO_JGS"]}
    catalogo = [p for p in todos if p["tabela"] in mapa.get(tabela,["HL_CONSUMO"])]
    for p in catalogo: p["_norm"] = expandir(p["descricao"])
    return catalogo


# ════════════════════════════════════════════════════════════════════════════
# GERADOR DE XLSX
# ════════════════════════════════════════════════════════════════════════════

def bd(e="thin"):
    s = Side(style=e)
    return Border(left=s,right=s,top=s,bottom=s)
def bb():
    return Border(bottom=Side(style="thin"))

def gerar_xlsx_bytes(confirmados, sugestoes_manuais, nao_enc, num, cliente, tabela_nome):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Orçamento"
    COR_H, COR_L = "1B3065", "E8EEF8"
    COR_COPPER = "C47A3A"

    for i, w in enumerate([5,44,14,10,13,10,8,13,15],1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ROW = 2
    ws.row_dimensions[ROW].height = 32
    ws.merge_cells(f"C{ROW}:I{ROW}")
    c = ws[f"C{ROW}"]; c.value = "PASQUETTI COMERCIO DE MATERIAIS HIDRÁULICOS LTDA"
    c.font = Font(name="Calibri",bold=True,size=15,color=COR_H)
    c.alignment = Alignment(horizontal="right",vertical="center")

    ROW+=1; ws.row_dimensions[ROW].height = 13
    ws.merge_cells(f"C{ROW}:I{ROW}"); c = ws[f"C{ROW}"]
    c.value = "www.pasquetti.com.br  |  tubos e conexões de ferro fundido"
    c.font = Font(name="Calibri",size=9,color="555555")
    c.alignment = Alignment(horizontal="right",vertical="center")

    for txt in ["CNPJ: 48.509.178/0001-99  |  Inscrição Estadual: 109.769.567.118",
                "RUA ARAPIRANGA, 55 - 93  |  VILA FORMOSA  |  São Paulo - SP  |  CEP: 03363-070",
                "Telefone: (11) 2784-4188  |  contato@pasquetti.com.br"]:
        ROW+=1; ws.row_dimensions[ROW].height = 13
        ws.merge_cells(f"C{ROW}:I{ROW}"); c = ws[f"C{ROW}"]
        c.value = txt; c.font = Font(name="Calibri",size=9,color="444444")
        c.alignment = Alignment(horizontal="right",vertical="center")

    # Linha cobre separadora
    ROW+=1; ws.row_dimensions[ROW].height = 4
    ws.merge_cells(f"A{ROW}:I{ROW}"); ws[f"A{ROW}"].fill = PatternFill("solid",fgColor=COR_COPPER)

    ROW+=1; ws.row_dimensions[ROW].height = 30
    ws.merge_cells(f"A{ROW}:E{ROW}"); c = ws[f"A{ROW}"]
    c.value = f"ORÇAMENTO Nº {num}"
    c.font = Font(name="Calibri",bold=True,size=17,color=COR_H)
    c.alignment = Alignment(horizontal="left",vertical="center")
    ws.merge_cells(f"F{ROW}:I{ROW}"); c = ws[f"F{ROW}"]
    c.value = datetime.date.today().strftime("%d/%m/%Y")
    c.font = Font(name="Calibri",size=11,color="666666")
    c.alignment = Alignment(horizontal="right",vertical="center")

    ROW+=1; ws.row_dimensions[ROW].height = 6

    # Cliente
    ROW+=1; ws.row_dimensions[ROW].height = 22
    ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
    c.value = "Informações do Cliente"; c.border = bb()
    c.font = Font(name="Calibri",bold=True,size=12,color=COR_H)
    c.alignment = Alignment(horizontal="left",vertical="center")

    ROW+=1; ws.row_dimensions[ROW].height = 20
    ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
    c.value = cliente.get("nome",""); c.font = Font(name="Calibri",bold=True,size=13)
    c.alignment = Alignment(horizontal="left",vertical="center")

    ROW+=1; ws.row_dimensions[ROW].height = 14
    ws.merge_cells(f"A{ROW}:D{ROW}"); ws[f"A{ROW}"].value = f"CNPJ: {cliente.get('cnpj','')}"
    ws[f"A{ROW}"].font = Font(name="Calibri",size=10)
    ws.merge_cells(f"E{ROW}:I{ROW}"); ws[f"E{ROW}"].value = cliente.get("endereco","")
    ws[f"E{ROW}"].font = Font(name="Calibri",size=10)

    ROW+=1; ws.row_dimensions[ROW].height = 14
    ws.merge_cells(f"A{ROW}:D{ROW}"); ws[f"A{ROW}"].value = f"Telefone: {cliente.get('telefone','')}"
    ws[f"A{ROW}"].font = Font(name="Calibri",bold=True,size=10)
    ws.merge_cells(f"E{ROW}:I{ROW}"); ws[f"E{ROW}"].value = f"Email: {cliente.get('email','')}"
    ws[f"E{ROW}"].font = Font(name="Calibri",bold=True,size=10)
    ROW+=1; ws.row_dimensions[ROW].height = 8

    # Tabela
    ROW+=1; ws.row_dimensions[ROW].height = 22
    ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
    c.value = "Itens do ORÇAMENTO"; c.border = bb()
    c.font = Font(name="Calibri",bold=True,size=12,color=COR_H)
    c.alignment = Alignment(horizontal="left",vertical="center")

    ROW+=1; ws.row_dimensions[ROW].height = 22
    for j,(h,a) in enumerate(zip(
        ["Produto","NCM","Quant.","Unit. (R$)","ICMS ST","IPI","Unit. c/ Imp.","Valor Total (R$)"],
        ["left","center","center","right","right","right","right","right"]),2):
        c = ws.cell(row=ROW,column=j,value=h)
        c.font = Font(name="Calibri",bold=True,size=10,color="FFFFFF")
        c.fill = PatternFill("solid",fgColor=COR_H)
        c.alignment = Alignment(horizontal=a,vertical="center"); c.border = bd()

    subtotal = 0
    for k,item in enumerate(confirmados):
        ROW+=1; ws.row_dimensions[ROW].height = 20
        bg = COR_L if k%2==0 else "FFFFFF"
        for j,(v,f,a) in enumerate(zip(
            [item["produto"],item["ncm"],item["quantidade"],item["preco"],0.0,0.0,item["preco"],item["total"]],
            [None,None,'#,##0.00','#,##0.0000','#,##0.00','#,##0.00','#,##0.0000','#,##0.00'],
            ["left","center","center","right","right","right","right","right"]),2):
            c = ws.cell(row=ROW,column=j,value=v)
            c.font = Font(name="Calibri",size=9)
            c.fill = PatternFill("solid",fgColor=bg)
            c.alignment = Alignment(horizontal=a,vertical="center",wrap_text=(j==2))
            c.border = bd()
            if f: c.number_format = f
        subtotal += item["total"]

    # Totais
    ROW+=1; ws.row_dimensions[ROW].height = 6
    for lbl,val,bold in [("Subtotal:",subtotal,False),("IPI:",0,False),
                          ("ICMS ST:",0,False),("Total:",subtotal,True)]:
        ROW+=1; ws.row_dimensions[ROW].height = 18
        ws.merge_cells(f"B{ROW}:H{ROW}"); c = ws[f"B{ROW}"]
        c.value = lbl; c.font = Font(name="Calibri",bold=bold,size=10,color=COR_H if bold else "444444")
        c.alignment = Alignment(horizontal="right",vertical="center")
        c2 = ws.cell(row=ROW,column=9,value=val)
        c2.font = Font(name="Calibri",bold=bold,size=10,color=COR_H if bold else "444444")
        c2.alignment = Alignment(horizontal="right",vertical="center")
        c2.number_format = '#,##0.00'
        if bold: c2.border = bd()

    # Condições
    ROW+=2; ws.row_dimensions[ROW].height = 22
    ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
    c.value = "Vencimentos e Condições"; c.border = bb()
    c.font = Font(name="Calibri",bold=True,size=12,color=COR_H)
    c.alignment = Alignment(horizontal="left",vertical="center")
    for lbl,val in [("Tabela utilizada:",tabela_nome),("Pagamento:","A Vista"),
                    ("Validade:","30 dias"),
                    ("Obs.:","Preços sujeitos a alteração sem aviso prévio | Entrega sujeita a confirmação de estoque")]:
        ROW+=1; ws.row_dimensions[ROW].height = 16
        ws.cell(row=ROW,column=2,value=lbl).font = Font(name="Calibri",bold=True,size=10)
        ws.merge_cells(f"C{ROW}:I{ROW}")
        ws.cell(row=ROW,column=3,value=val).font = Font(name="Calibri",size=10)

    # Itens para revisão manual
    if sugestoes_manuais:
        ROW+=2; ws.row_dimensions[ROW].height = 20
        ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
        c.value = "⚠ Itens com correspondência parcial (60–79%) — confirmar manualmente antes de incluir"
        c.font = Font(name="Calibri",bold=True,size=10,color="7a4f00")
        c.fill = PatternFill("solid",fgColor="FFF3CD")
        c.alignment = Alignment(horizontal="left",vertical="center")
        for s in sugestoes_manuais:
            ROW+=1; ws.row_dimensions[ROW].height = 16
            ws.cell(row=ROW,column=2,value=s["descricao"]).font = Font(name="Calibri",size=10,italic=True)
            ws.cell(row=ROW,column=4,value=f"Sugestão: {s['produto']}  ({s['score']:.0f}%)").font = Font(name="Calibri",size=10,color=COR_COPPER)
            ws.merge_cells(f"D{ROW}:I{ROW}")

    # Não encontrados
    if nao_enc:
        ROW+=2; ws.row_dimensions[ROW].height = 20
        ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
        c.value = "✗ Itens NÃO ENCONTRADOS no catálogo — adicionar manualmente"
        c.font = Font(name="Calibri",bold=True,size=10,color="7a0000")
        c.fill = PatternFill("solid",fgColor="FFE0E0")
        c.alignment = Alignment(horizontal="left",vertical="center")
        for na in nao_enc:
            ROW+=1; ws.row_dimensions[ROW].height = 16
            ws.cell(row=ROW,column=2,value=f"{na['descricao']}  (qtd: {na['quantidade']:.0f})").font = Font(name="Calibri",size=10)

    # Rodapé
    ROW+=2; ws.row_dimensions[ROW].height = 14
    ws.merge_cells(f"A{ROW}:I{ROW}"); c = ws[f"A{ROW}"]
    c.value = f"Gerado em {datetime.datetime.now().strftime('%d/%m/%Y às %H:%M:%S')} | Pasquetti — Sistema de Cotações"
    c.font = Font(name="Calibri",size=8,color="999999",italic=True)
    c.alignment = Alignment(horizontal="right",vertical="center")

    # Aba de revisão
    ws2 = wb.create_sheet("Revisão Completa")
    for w,col in zip([40,42,10,12,12,10],range(1,7)):
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.append(["Descrição Solicitada","Produto Correspondido","Qtd","Preço Unit.","Total","Status"])
    for it in confirmados: ws2.append([it["descricao"],it["produto"],it["quantidade"],it["preco"],it["total"],"CONFIRMADO"])
    for it in sugestoes_manuais: ws2.append([it["descricao"],it.get("produto",""),it["quantidade"],"","",f"SUGESTÃO {it['score']:.0f}%"])
    for it in nao_enc: ws2.append([it["descricao"],"NÃO ENCONTRADO",it["quantidade"],"","","NÃO ENCONTRADO"])

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ════════════════════════════════════════════════════════════════════════════

# ── Cabeçalho ──────────────────────────────────────────────────────────────
logo_b64 = get_logo_b64()
if logo_b64:
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:52px;object-fit:contain;" />'
else:
    logo_html = LOGO_SVG

st.markdown(f"""
<div class="pq-header">
  {logo_html}
  <div>
    <p class="pq-title">Gerador de Cotações</p>
    <p class="pq-sub">Preencha os dados do cliente, insira os itens e clique em Gerar Cotação</p>
  </div>
  <span class="pq-tag">Pasquetti</span>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    if logo_b64:
        st.markdown(f'<div style="text-align:center;padding:8px 0 16px;"><img src="data:image/png;base64,{logo_b64}" style="max-width:90%;max-height:60px;object-fit:contain;" /></div>', unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div style="text-align:center;padding:16px 0;border-bottom:1px solid rgba(255,255,255,0.15);margin-bottom:12px;">
          <div style="display:inline-flex;align-items:center;gap:10px;">
            {LOGO_SVG}
            <div style="text-align:left;">
              <div style="color:white;font-size:16px;font-weight:700;letter-spacing:1px;">PASQUETTI</div>
              <div style="color:{COPPER};font-size:9px;font-weight:500;letter-spacing:.5px;">TUBOS E CONEXÕES</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### 👤 Dados do Cliente")

    # CNPJ com busca automática
    cnpj_input = st.text_input("CNPJ", placeholder="00.000.000/0001-00",
                                help="Digite o CNPJ e clique em Buscar para preencher automaticamente")
    buscar_btn = st.button("🔍 Buscar dados na Receita Federal", use_container_width=True)

    # Inicializar as chaves dos campos (uma única fonte da verdade = a chave do widget)
    for k in ["input_nome","input_fantasia","input_end","input_tel","input_email"]:
        if k not in st.session_state: st.session_state[k] = ""

    if buscar_btn and cnpj_input.strip():
        with st.spinner("Consultando Receita Federal..."):
            dados, erro = buscar_cnpj(cnpj_input)
        if dados:
            # Escreve DIRETO nas chaves dos widgets (antes deles serem criados neste run)
            st.session_state["input_nome"]     = dados["nome"]
            st.session_state["input_fantasia"] = dados.get("fantasia","")
            st.session_state["input_end"]      = dados["endereco"]
            st.session_state["input_tel"]      = dados["telefone"]
            st.session_state["input_email"]    = dados["email"]
            sit  = dados.get("situacao","")
            ativ = dados.get("atividade","")
            st.success(f"✅ {dados['nome']}")
            if sit:  st.caption(f"Situação: {sit}")
            if ativ: st.caption(f"Atividade: {ativ}")
        else:
            st.error(f"❌ {erro}")

    # Campos: só 'key' (sem value=), para a busca de CNPJ poder preenchê-los
    nome     = st.text_input("Nome / Razão Social", placeholder="Construtora ABC Ltda", key="input_nome")
    fantasia = st.text_input("Nome Fantasia", placeholder="(opcional)", key="input_fantasia")
    endereco = st.text_input("Endereço", placeholder="Rua X, 123 - Bairro - Cidade", key="input_end")
    telefone = st.text_input("Telefone", placeholder="(11) 9999-9999", key="input_tel")
    email    = st.text_input("E-mail", placeholder="contato@cliente.com.br", key="input_email")

    st.markdown("---")
    st.markdown("### ⚙️ Configurações")
    tabela = st.selectbox("Tabela de preços", options=["consumo","revenda","pressao","todos"],
        format_func=lambda x: {"consumo":"🏠 Consumo — HL Mar/2026",
                                "revenda":"🏪 Revenda — HL Mar/2026",
                                "pressao":"💧 Pressão — JGS Abr/2026",
                                "todos":  "📋 Todas as tabelas"}[x])
    num_orcamento = st.text_input("Nº do Orçamento", placeholder="Gerado automaticamente")
    if not num_orcamento:
        import random; num_orcamento = str(random.randint(1000,9999))

    st.markdown("---")
    if not CATALOG_JSON.exists():
        st.error("⚠️ Catálogo não encontrado!\nRode extrair_catalogo.py primeiro.")
    else:
        with open(CATALOG_JSON) as f: cat_data = json.load(f)
        st.markdown(f"""<div style="background:rgba(255,255,255,0.08);border-radius:6px;padding:8px 10px;font-size:11px;color:#90c090;">
        ✅ <strong style="color:white">{len(cat_data)} produtos</strong> no catálogo<br>
        <span style="color:#aaa">Precisão mínima: 80%</span>
        </div>""", unsafe_allow_html=True)

    # ── Interpretação com IA ──────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🤖 Interpretação com IA")
    usar_ia = False
    if matcher_ia is None:
        st.caption("Biblioteca 'anthropic' não instalada. Rode o instalador novamente.")
    else:
        cfg_ia = matcher_ia.carregar_config()
        tem_chave = bool(cfg_ia["api_key"])
        if tem_chave:
            usar_ia = st.toggle("Usar IA para interpretar itens", value=True,
                                help="Entende pedidos vagos/abreviados e lê foto/PDF. Custa ~centavos por cotação.")
            mascara = cfg_ia["api_key"][:7] + "…" + cfg_ia["api_key"][-4:]
            st.caption(f"🔑 Chave configurada ({mascara}) · modelo: {cfg_ia['modelo']}")
            with st.expander("Trocar chave / modelo"):
                nova = st.text_input("Chave de API (sk-ant-...)", type="password", key="ia_key_new")
                modelo = st.text_input("Modelo", value=cfg_ia["modelo"], key="ia_model_new")
                if st.button("Salvar configuração de IA"):
                    matcher_ia.salvar_config(nova or cfg_ia["api_key"], modelo)
                    st.success("Configuração salva. Recarregue a página.")
        else:
            st.caption("Cole sua chave de API da Anthropic para ligar a interpretação por IA.")
            nova = st.text_input("Chave de API (sk-ant-...)", type="password", key="ia_key")
            modelo = st.text_input("Modelo", value=matcher_ia.MODELO_PADRAO, key="ia_model")
            if st.button("Salvar e ativar IA") and nova.strip():
                matcher_ia.salvar_config(nova, modelo)
                st.success("Chave salva. Recarregue a página para ativar.")


# ── Área principal ──────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📤  Carregar Arquivo", "✏️  Digitar / Colar", "📊  Tabela Interativa"])
itens_brutos = []

# Tab 1: Upload
with tab1:
    st.markdown('<span class="pq-section">Envie o arquivo da solicitação</span>', unsafe_allow_html=True)
    st.caption("Formatos aceitos: `.txt` · `.xlsx` · `.pdf` · `.jpg` · `.png`  (foto e PDF usam a IA)")
    arquivo = st.file_uploader("Arquivo", type=["txt","xlsx","xls","pdf","jpg","jpeg","png","bmp"],
                                label_visibility="collapsed")
    if arquivo:
        ext = Path(arquivo.name).suffix.lower()
        data = arquivo.read()
        ia_ok = matcher_ia is not None and matcher_ia.ia_disponivel()
        if ext == ".txt":
            itens_brutos = parse_texto(data.decode("utf-8", errors="ignore"))
            st.success(f"✅ **{len(itens_brutos)} itens** lidos do arquivo de texto")
        elif ext in (".xlsx",".xls"):
            itens_brutos = parse_xlsx_bytes(data)
            st.success(f"✅ **{len(itens_brutos)} itens** lidos da planilha")
        elif ext == ".pdf":
            if ia_ok:
                with st.spinner("Lendo o PDF com a IA..."):
                    try:
                        itens_brutos = matcher_ia.extrair_itens_de_pdf(data)
                    except Exception as e:
                        st.error(f"Erro ao ler o PDF com a IA: {e}"); itens_brutos = []
                if itens_brutos: st.success(f"✅ **{len(itens_brutos)} itens** lidos do PDF")
                else: st.warning("Não consegui extrair itens do PDF.")
            else:
                st.warning("Leitura de PDF requer a IA ligada (configure a chave na barra lateral).")
        elif ext in (".jpg",".jpeg",".png",".bmp"):
            mt = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png","bmp":"image/bmp"}[ext.strip(".")]
            if ia_ok:
                with st.spinner("Lendo a imagem com a IA..."):
                    try:
                        itens_brutos = matcher_ia.extrair_itens_de_imagem(data, media_type=mt)
                    except Exception as e:
                        st.error(f"Erro ao ler a imagem com a IA: {e}"); itens_brutos = []
                if itens_brutos: st.success(f"✅ **{len(itens_brutos)} itens** lidos da imagem")
                else: st.warning("Não consegui extrair itens da imagem.")
            else:
                with st.spinner("Reconhecendo texto na imagem (OCR)..."):
                    itens_brutos = parse_imagem_bytes(data)
                if itens_brutos: st.success(f"✅ **{len(itens_brutos)} itens** reconhecidos (OCR)")
                else: st.warning("Não foi possível reconhecer itens. Ligue a IA ou use o modo de digitação.")
        if itens_brutos:
            with st.expander("👁  Ver itens lidos"):
                for it in itens_brutos:
                    st.write(f"• **{it['quantidade']:.0f}x** {it['descricao']}")

# Tab 2: Texto livre
with tab2:
    st.markdown('<span class="pq-section">Cole ou digite os itens</span>', unsafe_allow_html=True)
    st.caption("Um item por linha — coloque a **quantidade no início**:")
    texto_livre = st.text_area("Itens", height=270, label_visibility="collapsed",
        placeholder="90  joelho fofo 100mm\n72  JF 75mm 90 graus\n6   anel 150mm\n190 anel borracha 100mm")
    if texto_livre.strip():
        itens_brutos = parse_texto(texto_livre)
        if itens_brutos: st.success(f"✅ **{len(itens_brutos)} itens** prontos")

# Tab 3: Tabela editável
with tab3:
    import pandas as pd
    st.markdown('<span class="pq-section">Preencha a tabela de itens</span>', unsafe_allow_html=True)
    if "tabela_itens" not in st.session_state:
        st.session_state.tabela_itens = pd.DataFrame({"Produto / Descrição":[""] * 8, "Quantidade":[1]*8})

    c1, c2, _ = st.columns([1,1,4])
    with c1:
        if st.button("➕ Adicionar linhas"):
            novas = pd.DataFrame({"Produto / Descrição":[""]*5,"Quantidade":[1]*5})
            st.session_state.tabela_itens = pd.concat([st.session_state.tabela_itens,novas],ignore_index=True)
    with c2:
        if st.button("🗑 Limpar"):
            st.session_state.tabela_itens = pd.DataFrame({"Produto / Descrição":[""] * 8,"Quantidade":[1]*8})

    tabela_edit = st.data_editor(
        st.session_state.tabela_itens, use_container_width=True, num_rows="dynamic",
        column_config={
            "Produto / Descrição": st.column_config.TextColumn("Produto / Descrição",
                help="Ex: joelho fofo 90 150mm | JF 100mm | anel borracha 75mm", width="large"),
            "Quantidade": st.column_config.NumberColumn("Qtd", min_value=0, step=1, width="small"),
        }, hide_index=True, key="editor_tabela")
    st.session_state.tabela_itens = tabela_edit

    itens_tab = []
    for _, row in tabela_edit.iterrows():
        desc = str(row.get("Produto / Descrição","")).strip()
        try: qtd = float(row.get("Quantidade",1) or 1)
        except: qtd = 1
        if desc and desc.lower() not in ("none","nan",""): itens_tab.append({"descricao":desc,"quantidade":qtd})
    if itens_tab:
        itens_brutos = itens_tab
        st.success(f"✅ **{len(itens_tab)} itens** preenchidos")


# ── Botão Gerar ────────────────────────────────────────────────────────────
st.markdown("---")
col_btn, col_hint = st.columns([2,3])
with col_btn:
    gerar = st.button("🧾 Gerar Cotação", type="primary",
                      disabled=(not itens_brutos or not CATALOG_JSON.exists()))
with col_hint:
    if not itens_brutos:
        st.info("Insira os itens em uma das abas acima.")
    elif not nome:
        st.warning("💡 Preencha o nome do cliente no menu lateral para identificar a cotação.")


# ── Resultado ──────────────────────────────────────────────────────────────
if gerar and itens_brutos:
    spin_msg = "Interpretando itens com IA e buscando no catálogo..." if usar_ia \
               else "Buscando correspondências no catálogo (precisão ≥ 80%)..."
    with st.spinner(spin_msg):
        catalogo  = carregar_catalogo(tabela)
        conf, sug, nao = processar_hibrido(itens_brutos, catalogo, usar_ia)

    subtotal = sum(i["total"] for i in conf)

    # Métricas
    st.markdown("---")
    st.markdown('<span class="pq-section">Resultado da Cotação</span>', unsafe_allow_html=True)

    m1,m2,m3,m4 = st.columns(4)
    with m1: st.markdown(f'<div class="pq-metric green"><div class="pq-metric-val green">{len(conf)}</div><div class="pq-metric-lbl">Confirmados (≥80%)</div></div>', unsafe_allow_html=True)
    with m2: st.markdown(f'<div class="pq-metric copper"><div class="pq-metric-val copper">{len(sug)}</div><div class="pq-metric-lbl">Sugestões (60–79%)</div></div>', unsafe_allow_html=True)
    with m3: st.markdown(f'<div class="pq-metric red"><div class="pq-metric-val red">{len(nao)}</div><div class="pq-metric-lbl">Não encontrados</div></div>', unsafe_allow_html=True)
    with m4: st.markdown(f'<div class="pq-metric"><div class="pq-metric-val">R$ {subtotal:,.2f}</div><div class="pq-metric-lbl">Subtotal (confirmados)</div></div>', unsafe_allow_html=True)

    # Tabela de confirmados
    if conf:
        st.markdown('<span class="pq-section" style="color:#27ae60">✅ Itens Confirmados — entram na cotação</span>', unsafe_allow_html=True)
        linhas = ""
        for it in conf:
            s = it["score"]
            badge = f'<span class="badge badge-a">{s:.0f}%</span>' if s>=90 else f'<span class="badge badge-m">{s:.0f}%</span>'
            linhas += f"""<tr>
              <td style="color:#555;font-style:italic;font-size:11px">{it['descricao']}</td>
              <td><strong>{it['produto']}</strong></td>
              <td style="text-align:center">{it['quantidade']:.0f}</td>
              <td style="text-align:right">R$ {it['preco']:,.4f}</td>
              <td style="text-align:right"><strong>R$ {it['total']:,.2f}</strong></td>
              <td style="text-align:center">{badge}</td>
            </tr>"""
        st.markdown(f"""<table class="pq-table">
          <thead><tr>
            <th>Solicitado</th><th>Produto no Catálogo</th>
            <th style="text-align:center">Qtd</th>
            <th style="text-align:right">Preço Unit.</th>
            <th style="text-align:right">Total</th>
            <th style="text-align:center">Score</th>
          </tr></thead>
          <tbody>{linhas}</tbody>
        </table>
        <p style="text-align:right;font-weight:700;font-size:15px;color:{NAVY};margin-top:10px;">
          Subtotal: R$ {subtotal:,.2f}
        </p>""", unsafe_allow_html=True)

    # Sugestões (não entram automaticamente)
    if sug:
        sug_rows = ""
        for s in sug:
            just = s.get("justificativa", "")
            just_html = f'<br><span style="color:#999;font-size:10px">{just}</span>' if just else ""
            sug_rows += f"""<tr>
              <td style="color:#555;font-style:italic">{s['descricao']}</td>
              <td>{s.get('produto','—')}{just_html}</td>
              <td style="text-align:center">{s['quantidade']:.0f}</td>
              <td style="text-align:center"><span class="badge badge-s">{s['score']:.0f}%</span></td>
            </tr>"""
        st.markdown(f"""<div class="pq-suggest">
          <div class="pq-suggest-title">⚠️ Sugestões com correspondência parcial (60–79%) — NÃO incluídas na cotação</div>
          <p style="color:#7a5a20;margin-bottom:10px;font-size:11px;">
            Revise cada item abaixo. Se a sugestão estiver correta, confirme e inclua manualmente na cotação.
          </p>
          <table class="pq-table" style="font-size:11px;">
            <thead><tr>
              <th>Solicitado</th><th>Melhor Sugestão</th>
              <th style="text-align:center">Qtd</th>
              <th style="text-align:center">Precisão</th>
            </tr></thead>
            <tbody>{sug_rows}</tbody>
          </table>
        </div>""", unsafe_allow_html=True)

    # Não encontrados
    if nao:
        nao_rows = "".join(f"<li><strong>{n['descricao']}</strong> (qtd: {n['quantidade']:.0f})</li>" for n in nao)
        st.markdown(f"""<div class="pq-notfound">
          <div class="pq-notfound-title">❌ Itens não encontrados no catálogo</div>
          <ul style="margin:6px 0;padding-left:18px;font-size:12px;">{nao_rows}</ul>
        </div>""", unsafe_allow_html=True)

    # Download
    st.markdown("---")
    tabelas_nome = {"consumo":"TABELA HL CONSUMO — Março/2026","revenda":"TABELA HL REVENDA — Março/2026",
                    "pressao":"TABELA PRESSÃO JGS — Abril/2026","todos":"MÚLTIPLAS TABELAS"}
    cliente_dict = {"nome":nome,"cnpj":cnpj_input,"endereco":endereco,"telefone":telefone,"email":email}
    xlsx_bytes = gerar_xlsx_bytes(conf, sug, nao, num_orcamento, cliente_dict, tabelas_nome[tabela])
    nome_arq = f"cotacao_{num_orcamento}_{nome.replace(' ','_')[:18]}.xlsx" if nome else f"cotacao_{num_orcamento}.xlsx"

    dcol, icol = st.columns([1,2])
    with dcol:
        st.download_button("⬇️  Baixar Cotação (.xlsx)", data=xlsx_bytes, file_name=nome_arq,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with icol:
        st.info(f"📄 Orçamento **Nº {num_orcamento}** · {len(conf)} itens confirmados · R$ {subtotal:,.2f}")


# ── Rodapé ─────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(f"""<p style='text-align:center;color:#aaa;font-size:11px;'>
  <span style='color:{COPPER};font-weight:600'>PASQUETTI</span>
  Comercio de Materiais Hidráulicos Ltda
  · tubos e conexões de ferro fundido ·
  Sistema de Cotações v2.0
</p>""", unsafe_allow_html=True)
