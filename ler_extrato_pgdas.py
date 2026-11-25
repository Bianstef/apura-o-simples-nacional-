#!/usr/bin/env python3
"""Lê o extrato do PGDAS-D (PDF da Receita Federal) e devolve, em JSON, o faturamento mês a mês.

Uso:  python3 ler_extrato_pgdas.py extrato.pdf [outro.pdf ...]

Para cada extrato:
- cnpj_basico, nome, pa (período de apuração "MM/AAAA");
- rpa: receita bruta do próprio PA (mercado interno + externo);
- rbt12_extrato: RBT12 informada no extrato (para conferência);
- meses: {"MM/AAAA": receita} das "Receitas Brutas Anteriores" (interno + externo) + o próprio PA;
- folha: {"MM/AAAA": valor} das "Folhas de Salários Anteriores", quando houver;
- confere: true se a soma dos meses da RBT12 bate com a RBT12 do extrato
  (até 2026: 12 meses anteriores ao PA; a partir de 2027: 12 meses antes do mês anterior ao PA).
"""
import json, re, subprocess, sys

VAL = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
PAR = re.compile(r"(\d{2}/\d{4})\s+" + VAL)

def num(txt):
    return float(txt.replace(".", "").replace(",", "."))

def texto(pdf):
    return subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True, check=True).stdout

def secao(t, inicio, fins):
    i = re.search(inicio, t, re.I)
    if not i: return ""
    resto = t[i.end():]
    corte = len(resto)
    for f in fins:
        j = re.search(f, resto, re.I)
        if j: corte = min(corte, j.start())
    return resto[:corte]

def pares(bloco):
    out = {}
    for m, v in PAR.findall(bloco):
        out[m] = out.get(m, 0.0) + num(v)
    return out

def mes_anterior(pa, n):
    m, a = int(pa[:2]), int(pa[3:])
    t = a*12 + m - 1 - n
    return f"{t % 12 + 1:02d}/{t // 12}"

def ler(pdf):
    t = texto(pdf)
    r = {"arquivo": pdf}
    # extrato: "Período de Apuração (PA): 10/2026"; declaração: "Período de Apuração: 01/10/2026 a 31/10/2026"
    m = re.search(r"Per[ií]odo de Apura[cç][aã]o(?: \(PA\))?:?\s*(?:\d{2}/)?(\d{2}/\d{4})", t, re.I)
    r["pa"] = m.group(1) if m else None
    m = re.search(r"CNPJ Matriz:?\s*([\d\./-]{18})", t)
    r["cnpj"] = m.group(1) if m else None
    m = re.search(r"CNPJ Estabelecimento:?\s*([\d\./-]{18})", t)
    if m and not r["cnpj"]: r["cnpj"] = m.group(1)
    m = re.search(r"UF:\s*([A-Z]{2})", t); r["uf"] = m.group(1) if m else None
    m = re.search(r"CNPJ B[aá]sico:?\s*([\d\.]{8,11})", t, re.I)
    r["cnpj_basico"] = re.sub(r"\D", "", m.group(1)) if m else (re.sub(r"\D", "", r["cnpj"])[:8] if r["cnpj"] else None)
    m = re.search(r"Nome [Ee]mpresarial:?\s*(.+)", t)
    r["nome"] = m.group(1).strip() if m else None
    # Receita do PA: primeira linha "Receita Bruta do PA (RPA)" -> pega o último valor da linha (Total)
    m = re.search(r"Receita Bruta do PA \(RPA\)[^\n]*", t, re.I)
    if m:
        vals = re.findall(VAL, m.group(0)); r["rpa"] = num(vals[-1]) if vals else None
    # a linha da RBT12 quebra no PDF: os valores ficam na linha "...doze meses anteriores (ao PA)"
    m = re.search(r"doze meses anteriores(?: ao PA)?\s{2,}[^\n]*", t)
    if m:
        vals = re.findall(VAL, m.group(0)); r["rbt12_extrato"] = num(vals[-1]) if vals else None
    ant = secao(t, r"Receitas? Brutas? Anteriores", [r"Folhas? de Sal[aá]rios? Anteriores", r"2\.4\)", r"Fator r", r"\n\s*3\)"])
    meses = pares(ant)
    folha_bloco = secao(t, r"Folhas? de Sal[aá]rios? Anteriores", [r"2\.4\)", r"Fator r", r"\n\s*3\)", r"Valor Original"])
    r["folha"] = pares(folha_bloco)
    if r.get("pa") and r.get("rpa") is not None:
        meses[r["pa"]] = r["rpa"]
    r["meses"] = dict(sorted(meses.items(), key=lambda kv: (kv[0][3:], kv[0][:2])))
    if r.get("pa") and r.get("rbt12_extrato") is not None:
        ini = 2 if int(r["pa"][3:]) >= 2027 else 1
        soma = sum(meses.get(mes_anterior(r["pa"], k), 0.0) for k in range(ini, ini + 12))
        r["soma_12_anteriores"] = round(soma, 2)
        r["confere"] = abs(soma - r["rbt12_extrato"]) < 0.05
    return r

if __name__ == "__main__":
    print(json.dumps([ler(p) for p in sys.argv[1:]], ensure_ascii=False, indent=1))
