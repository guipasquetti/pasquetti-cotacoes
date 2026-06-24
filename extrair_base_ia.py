#!/usr/bin/env python3
"""
Conversor da 'tabela base IA' (catálogo SAP/JGS completo da linha de pressão)
para o formato do catalogo_produtos.json (tabela PRESSAO_JGS).

Cada linha da planilha já é uma variação completa (tipo, conexão, DN, comprimento,
classe), com NCM, código SAP e preço de venda Pasquetti. Aqui só lemos as colunas
e montamos uma descrição legível + estruturada (para o matching) preservando os
campos extras (cod_sap, figura, tipo, unidade).

Uso:  python3 extrair_base_ia.py "tabela base IA  24 06 2026.xlsx"
Gera: pressao_base_ia.json  (entradas tabela=PRESSAO_JGS)
"""
import sys, os, re, json
import openpyxl

HDR = 30  # linha do cabeçalho de colunas na aba TabPreço

# Expansão das siglas de conexão/variação que aparecem na DESCRIÇÃO SAP
CONEX = {
    "FL": "flange", "FLS": "flange flange", "BOL": "bolsa", "PTA": "ponta",
    "PTAS": "ponta ponta", "EXC": "excentrica", "CONC": "concentrica",
    "AV": "avulso", "GAVETA": "gaveta", "BORB": "borboleta", "CEGO": "cego",
    "EXTRE": "extremidade", "CR": "de correr", "PÉ": "pe", "CILIND": "cilindrico",
    "VEDAÇÃO": "vedacao", "ABA": "aba",
}
# Tipo canônico a partir do 1º termo da DESCRIÇÃO SAP
TIPO_HEAD = {
    "TUBO": "tubo", "TÊ": "te", "TE": "te", "CURVA": "curva",
    "REDUÇÃO": "reducao", "REDUCAO": "reducao", "CRUZETA": "cruzeta",
    "JUNÇÃO": "juncao", "JUNCAO": "juncao", "LUVA": "luva", "TOCO": "toco",
    "CARRETEL": "carretel", "VÁL": "valvula", "VÁLV": "valvula",
    "FLANGE": "flange", "EXTRE": "extremidade", "EXTREMIDADE": "extremidade",
    "ADAPTADOR": "adaptador", "ARRUELA": "arruela", "COLAR": "colar",
    "ANEL": "anel", "TAMPÃO": "tampao",
}
DOIS_DN = {"te", "reducao", "cruzeta", "juncao"}  # figura traz DN1 DN2


def fmt_ncm(n):
    s = re.sub(r"\D", "", str(n or ""))
    if s.startswith("00"):
        s = s[2:]
    return f"{s[0:4]}.{s[4:6]}.{s[6:8]}" if len(s) == 8 else s


def parse(desc_sap, figura):
    parts = re.split(r"[-/ ]+", str(desc_sap).strip())
    head = parts[0].upper()
    tipo = TIPO_HEAD.get(head, parts[0].lower())
    conex = " ".join(CONEX.get(p.upper(), "") for p in parts[1:]).strip()
    F = str(figura or "").upper()
    nums = [int(n) for n in re.findall(r"\d+", F)]
    dns = [n for n in nums if 40 <= n <= 2000]
    comp = [n for n in nums if n >= 900 and n not in dns]  # comprimento (mm)
    cls = []
    mk = re.search(r"K(7|9)", F)
    if mk:
        cls.append("k" + mk.group(1))
    mp = re.search(r"\.(\d{2}(?:/\d{2})*)", F)
    if mp:
        cls += ["pn" + n for n in mp.group(1).split("/") if n in ("10", "16", "25")]
    ma = re.search(r"\bC(90|45|22|11)", F)
    ang = ma.group(1) if ma else ""
    if tipo in DOIS_DN and len(dns) >= 2:
        dn_txt = f"{dns[0]}mm {dns[0]}x{dns[1]} {dns[1]}mm"
    elif dns:
        dn_txt = f"{dns[0]}mm"
    else:
        dn_txt = ""
    comp_txt = str(comp[0]) if comp else ""   # número solto: não polui o DN
    desc = re.sub(r"\s+", " ",
                  f"{tipo} {conex} {ang} {dn_txt} {comp_txt} {' '.join(cls)}").strip()
    return desc, tipo, (dns[0] if dns else "")


def converter(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb["TabPreço"]
    entries = []
    for row in ws.iter_rows(min_row=HDR + 1, values_only=True):
        d = row[2]                      # DESCRIÇÃO SAP
        if not d or not str(d).strip():
            continue
        try:
            pv = float(row[13])         # preço venda pasquetti
        except (TypeError, ValueError):
            pv = 0.0
        if pv <= 0:
            continue
        desc, tipo, dn = parse(d, row[3])
        entries.append({
            "categoria": str(d).strip(),
            "tamanho": str(dn),
            "descricao": desc,
            "preco": round(pv, 2),
            "tabela": "PRESSAO_JGS",
            "ncm": fmt_ncm(row[6]),
            "unidade": "M" if str(row[7] or "").strip().upper() == "M" else "UN",
            "tipo": tipo,
            "cod_sap": str(row[4] or "").strip(),
            "figura": str(row[3] or "").strip(),
        })
    return entries


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "tabela base IA  24 06 2026.xlsx"
    ent = converter(src)
    out = os.path.join(os.path.dirname(os.path.abspath(src)) or ".", "pressao_base_ia.json")
    json.dump(ent, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"{len(ent)} itens PRESSAO_JGS gerados -> {out}")
