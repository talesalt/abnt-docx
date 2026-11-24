#!/usr/bin/env python3
"""revisar_md — encontra e corrige no Markdown o que não combina com texto acadêmico.

  py -3 revisar_md.py analisar arquivo.md            → relatório numerado de achados
  py -3 revisar_md.py aplicar  arquivo.md --itens bullets,regua,...
                                                     → grava arquivo_academico.md (o original fica intacto)

Cada achado tem uma CHAVE, a contagem, um exemplo e o tipo:
  auto   = o script corrige sozinho
  manual = precisa de reescrita (o Claude edita o *_academico.md à mão, se a pessoa aceitar)
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

MARCADOR = re.compile(r"^([-*+]|\d{1,3}[.)])(\s+)(.*)$")
FENCE = re.compile(r"^\s*(```|~~~)")
EMOJI = re.compile("[\U0001F300-\U0001FAFF\U00002600-\U000027BF\U0001F000-\U0001F2FF\uFE0F\u200D]")
LINK_LOCAL = re.compile(r"\[([^\]]+)\]\((?:file:|[A-Za-z]:[\\/])[^)]*\)")
REGUA = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
# nota editorial em itálico: *(...)* ou trecho em itálico terminado em "!"
NOTA_ED = re.compile(r"\s*(?:[—–-]\s*)?(?<!\*)\*(?!\*)\(?[^*\n]{3,}?!\)?\*(?!\*)|\s*(?<!\*)\*(?!\*)\([^*\n]{3,}?\)\*(?!\*)")
# só rótulos genéricos de fichamento; termos definidos (**Hospitalismo:** ...) ficam
ROTULO = re.compile(r"^(\s*(?:[-*+]|\d{1,3}[.)])?\s*)\*\*((?:Conceitos?|Exemplos?(?: práticos?| marcantes?| empíricos?)?"
                    r"|Objetivo|Obs\.?|Observação|Nota|Resumo|Definição|Ideia central|Destaque|Síntese"
                    r"|Comentário|Importante|Atenção|Dica)(?: [^*]{0,40})?):\*\*\s*", re.I)
NEGRITO = re.compile(r"\*\*([^*\n]+)\*\*")
CODIGO = re.compile(r"`([^`\n]+)`")
TIL = re.compile(r"~(?=\d)")
SETA = re.compile(r"\s(→|⇒|->)\s")
EXCL = re.compile(r"(?<=\w)!(?=\s|$|\)|\*)")
CAIXA_ALTA = re.compile(r"\b[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{4,}(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,})+\b")


def blocos_texto(linhas):
    """Itera (i, linha) fora de front matter e blocos de código."""
    fence = False
    inicio = 0
    if linhas and linhas[0].strip() == "---":
        for j in range(1, len(linhas)):
            if linhas[j].strip() in ("---", "..."):
                inicio = j + 1
                break
    for i in range(inicio, len(linhas)):
        l = linhas[i]
        if FENCE.match(l):
            fence = not fence
            continue
        if not fence:
            yield i, l


def eh_titulo(l):
    return l.lstrip().startswith("#")


def eh_citacao(l):
    return l.lstrip().startswith(">")


def cortar(s, n=90):
    s = s.strip()
    return s if len(s) <= n else s[: n - 1] + "…"


# ------------------------------------------------------------------ análise
def analisar(texto: str):
    L = texto.splitlines()
    achados = []

    def add(chave, tipo, descricao, ocorr, recomendado=True):
        if ocorr:
            achados.append(dict(chave=chave, tipo=tipo, descricao=descricao, n=len(ocorr),
                                exemplo=cortar(ocorr[0]), recomendado=recomendado))

    corpo = list(blocos_texto(L))
    prosa = [l for _, l in corpo if l.strip()]
    topo = [l for _, l in corpo if MARCADOR.match(l)]
    if topo:
        prop = len(topo) / max(1, len([l for l in prosa if not eh_titulo(l)]))
        add("bullets", "auto",
            f"Texto em tópicos: {len(topo)} itens no 1º nível de lista ({prop:.0%} das linhas de texto). "
            "O 1º nível vira parágrafos corridos; subníveis sobem um nível", topo)
    add("regua", "auto", "Linhas horizontais (---) separando seções", [l for _, l in corpo if REGUA.match(l)])
    add("emoji", "auto", "Emojis e símbolos decorativos", [l for _, l in corpo if EMOJI.search(l)])
    add("links-locais", "auto", "Links para arquivos locais (file:/// ou C:\\...) viram texto simples",
        [l for _, l in corpo if LINK_LOCAL.search(l)])
    add("notas-editoriais", "auto",
        "Comentários informais em itálico, como *(gancho para X!)*: são removidos",
        [m.group(0) for _, l in corpo if not eh_citacao(l) for m in NOTA_ED.finditer(l)])
    add("rotulos", "auto",
        "Rótulos de fichamento no início do parágrafo (**Conceito:**, **Exemplo:**): o rótulo sai, o texto fica",
        [l for _, l in corpo if ROTULO.match(l) and ROTULO.sub("", l, 1).strip()], recomendado=False)
    neg = [m.group(0) for _, l in corpo if not eh_titulo(l) for m in NEGRITO.finditer(ROTULO.sub(r"\1", l, 1))]
    add("negrito", "auto",
        f"Negrito usado como ênfase no meio do texto ({len(neg)} trechos): vira texto normal",
        neg, recomendado=len(neg) > 15)
    add("codigo", "auto", "Trechos em `código` (nomes de arquivo etc.) viram texto normal",
        [l for _, l in corpo if CODIGO.search(l)])
    add("til", "auto", "Til como 'aproximadamente' (~3º mês) vira 'aprox.'",
        [l for _, l in corpo if TIL.search(l)])
    add("exclamacao", "auto", "Pontos de exclamação fora das citações viram ponto final",
        [l for _, l in corpo if not eh_citacao(l) and EXCL.search(NOTA_ED.sub("", l))])
    # ---- manuais
    add("setas", "manual", "Setas (→) e símbolos no lugar de palavras: reescrever com conectivos",
        [l for _, l in corpo if not eh_citacao(l) and SETA.search(l)])
    add("caixa-alta", "manual", "Expressões em CAIXA ALTA no meio do texto (fora de títulos e referências)",
        [m.group(0) for _, l in corpo if not eh_titulo(l) and not eh_citacao(l)
         for m in CAIXA_ALTA.finditer(re.sub(r"\*\*|\*", "", l))], recomendado=False)
    cit_rot = [l for _, l in corpo if re.match(r"^\s*(?:[-*+]\s+)?(?:\*\*)?Cita[çc][ãa]o[^*\n]*\([^)]*\d{4}[^)]*\):?(?:\*\*)?\s*$", l)]
    add("chamadas-citacao", "manual",
        "Citações introduzidas por rótulo (**Citação longa (Autor, ano, p.):**): reescrever como frase "
        "de apresentação, ex.: 'Conforme Fulgencio (2022, p. 159):'", cit_rot)
    meta = [l for _, l in corpo[:15] if re.match(r"^\s*\*\*(Objetivo|Nota|Obs\.?|Observação)\b", l)]
    add("meta", "manual", "Parágrafo de apresentação do arquivo (**Objetivo:** ...) em vez de introdução acadêmica", meta,
        recomendado=False)
    return achados


def imprimir_relatorio(nome, achados):
    if not achados:
        print(f"OK {nome}: nada fora do padrão acadêmico encontrado")
        return
    print(f"RELATORIO {nome}: {len(achados)} achados")
    for k, a in enumerate(achados, 1):
        rec = " [recomendado]" if a["recomendado"] else ""
        print(f"{k}. [{a['chave']}] ({a['tipo']}, {a['n']}x){rec} {a['descricao']}")
        print(f"   ex.: {a['exemplo']}")


# ------------------------------------------------------------------ correções
def tirar_primeiro_nivel(L):
    out = []
    i, n = 0, len(L)
    fence = False
    larg = None          # largura do marcador do item de 1º nível atual
    while i < n:
        l = L[i]
        if FENCE.match(l):
            fence = not fence
        if fence or FENCE.match(l):
            out.append(l); i += 1; continue
        m = MARCADOR.match(l)
        if m:
            larg = len(m.group(1)) + len(m.group(2))
            if out and out[-1].strip():
                out.append("")
            out.append(m.group(3))
            i += 1
            # linhas do mesmo item
            while i < n:
                l2 = L[i]
                if MARCADOR.match(l2) or FENCE.match(l2):
                    break
                if not l2.strip():
                    # fim do item se a próxima linha não-vazia não estiver indentada
                    j = i
                    while j < n and not L[j].strip():
                        j += 1
                    if j >= n or (not L[j].startswith(" " * min(larg, 2)) and not MARCADOR.match(L[j])):
                        break
                    out.append(""); i += 1; continue
                ind = len(l2) - len(l2.lstrip(" "))
                if ind >= min(larg, 2):
                    tira = min(ind, larg)
                    novo = l2[tira:]
                    if out and out[-1].strip() and not (MARCADOR.match(out[-1]) or eh_citacao(out[-1])) \
                            and (MARCADOR.match(novo) or eh_citacao(novo)):
                        out.append("")
                    if out and eh_citacao(out[-1]) and MARCADOR.match(novo):
                        out.append("")
                    out.append(novo)
                else:
                    out.append(l2)    # continuação preguiçosa do parágrafo
                i += 1
            out.append("")
            continue
        out.append(l); i += 1
    # colapsa linhas em branco repetidas
    res = []
    for l in out:
        if not l.strip() and res and not res[-1].strip():
            continue
        res.append(l)
    return res


def aplicar(texto: str, itens: set[str]):
    L = texto.splitlines()
    cont = {}

    def sub_linhas(chave, padrao, repl, filtro=lambda l: True):
        total = 0
        for i, l in list(blocos_texto(L)):
            if filtro(l):
                novo, k = padrao.subn(repl, l)
                if k:
                    L[i] = novo
                    total += k
        cont[chave] = total

    if "notas-editoriais" in itens:
        sub_linhas("notas-editoriais", NOTA_ED, "", lambda l: not eh_citacao(l))
    if "regua" in itens:
        idx = [i for i, l in blocos_texto(L) if REGUA.match(l)]
        for i in reversed(idx):
            L[i] = ""
        cont["regua"] = len(idx)
    if "emoji" in itens:
        sub_linhas("emoji", EMOJI, "")
    if "links-locais" in itens:
        sub_linhas("links-locais", LINK_LOCAL, r"\1")
    if "rotulos" in itens:
        sub_linhas("rotulos", ROTULO, r"\1", lambda l: bool(ROTULO.sub("", l, 1).strip()))
    if "negrito" in itens:
        total = 0
        for i, l in list(blocos_texto(L)):
            if eh_titulo(l):
                continue
            m = ROTULO.match(l)
            pre, resto = (l[: m.end()], l[m.end():]) if m else ("", l)
            novo, k = NEGRITO.subn(r"\1", resto)
            if k:
                L[i] = pre + novo; total += k
        cont["negrito"] = total
    if "codigo" in itens:
        sub_linhas("codigo", CODIGO, r"\1")
    if "til" in itens:
        sub_linhas("til", TIL, "aprox. ")
    if "exclamacao" in itens:
        sub_linhas("exclamacao", EXCL, ".", lambda l: not eh_citacao(l))
    if "bullets" in itens:
        antes = sum(1 for _, l in blocos_texto(L) if MARCADOR.match(l))
        L = tirar_primeiro_nivel(L)
        cont["bullets"] = antes
    # limpa espaços sobrando
    L = [re.sub(r"[ \t]+$", "", l) for l in L]
    return "\n".join(L).rstrip() + "\n", cont


def main():
    ap = argparse.ArgumentParser(description="Revisão acadêmica de Markdown")
    ap.add_argument("acao", choices=["analisar", "aplicar"])
    ap.add_argument("arquivo")
    ap.add_argument("--itens", default="", help="chaves separadas por vírgula (ou 'recomendados')")
    ap.add_argument("-o", "--saida")
    a = ap.parse_args()
    p = Path(a.arquivo)
    texto = p.read_text(encoding="utf-8-sig")
    achados = analisar(texto)
    if a.acao == "analisar":
        imprimir_relatorio(p.name, achados)
        return
    if a.itens.strip() == "recomendados":
        itens = {x["chave"] for x in achados if x["recomendado"] and x["tipo"] == "auto"}
    else:
        itens = {x.strip() for x in a.itens.split(",") if x.strip()}
    novo, cont = aplicar(texto, itens)
    out = Path(a.saida) if a.saida else p.with_name(p.stem + "_academico.md")
    out.write_text(novo, encoding="utf-8")
    feitos = ", ".join(f"{k}={v}" for k, v in cont.items())
    print(f"OK {out} ({feitos or 'nenhuma alteração automática'})")
    restantes = [x["chave"] for x in analisar(novo)]
    if restantes:
        print("AINDA PRESENTES (rode analisar de novo se quiser outra passada): " + ", ".join(restantes))


if __name__ == "__main__":
    sys.exit(main())
