#!/usr/bin/env python3
"""md2abnt — converte Markdown em .docx formatado segundo a ABNT.

Regras aplicadas (NBR 14724:2024, 6024:2012, 10520:2023, 6023:2018):
  A4; margens sup/esq 3 cm, inf/dir 2 cm; Times New Roman ou Arial 12;
  entrelinha 1,5; recuo de 1a linha 1,25 cm; texto justificado;
  citação longa (blockquote) com recuo de 4 cm, fonte 10, espaço simples, sem aspas;
  notas, tabelas e legendas em fonte 10 e espaço simples;
  títulos de seção conforme NBR 6024 (primária NEGRITO CAIXA ALTA, secundária CAIXA ALTA,
  terciária negrito, quaternária itálico, quinária normal);
  títulos sem indicativo numérico (Sumário, Referências...) centralizados;
  referências em espaço simples, alinhadas à esquerda, separadas por uma linha em branco;
  paginação no canto superior direito, fonte 10; capa e folha de rosto opcionais
  (capa não é contada; a folha de rosto é contada e não numerada).

Uso:  py -3 md2abnt.py arquivo.md [mais.md ...] [-o saida.docx|pasta] [--fonte arial]
          [--capa] [--sumario] [--numerar] [--quebrar-secoes] [--manter-aspas]
          [--instituicao ... --autor ... --cidade ... --ano ... --natureza ... --orientador ...]
Os dados da capa também podem vir do front matter YAML do .md (ver SKILL.md).
"""
from __future__ import annotations

import argparse
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import docx
    from docx.enum.section import WD_SECTION
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    sys.exit("ERRO python-docx ausente: py -3 -m pip install --user python-docx")

AQUI = Path(__file__).resolve().parent
LUA = AQUI / "abnt.lua"
FONTES = {"times": "Times New Roman", "arial": "Arial"}
PRETO = RGBColor(0, 0, 0)


# ----------------------------------------------------------------- utilidades
def pandoc_bin() -> str:
    p = shutil.which("pandoc")
    if not p:
        sys.exit("ERRO pandoc não encontrado (instale: winget install JohnMacFarlane.Pandoc)")
    return p


def set_fonte_rpr(rpr_parent, nome: str):
    """Força a fonte em todos os scripts, removendo atributos de tema."""
    rpr = rpr_parent.get_or_add_rPr() if hasattr(rpr_parent, "get_or_add_rPr") else rpr_parent
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in list(rf.attrib):
        if a.endswith("Theme"):
            del rf.attrib[a]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), nome)


def estilo_par(st, fonte, tam=12, negrito=False, italico=False, caixa_alta=False,
               alinh=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinha=1.5, recuo1=0.0,
               recuo_esq=0.0, antes=0, depois=0, manter_prox=False):
    f = st.font
    f.name, f.size, f.bold, f.italic = fonte, Pt(tam), negrito, italico
    f.all_caps = caixa_alta
    f.color.rgb = PRETO
    set_fonte_rpr(st.element, fonte)
    pf = st.paragraph_format
    pf.alignment = alinh
    if entrelinha == 1:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing = entrelinha
    pf.first_line_indent = Cm(recuo1)
    pf.left_indent = Cm(recuo_esq)
    pf.right_indent = Cm(0)
    pf.space_before, pf.space_after = Pt(antes), Pt(depois)
    pf.keep_with_next = manter_prox
    pf.widow_control = True


def campo(par, instr: str):
    """Insere um campo do Word (PAGE, TOC...) num parágrafo."""
    r = par.add_run()
    for tipo, texto in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        if tipo:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), tipo)
            r._r.append(e)
        elif texto == instr:
            e = OxmlElement("w:instrText")
            e.set(qn("xml:space"), "preserve")
            e.text = f" {instr} "
            r._r.append(e)
        else:
            e = OxmlElement("w:t")
            e.text = texto
            r._r.append(e)
    return r


# --------------------------------------------------------- modelo de estilos
def criar_referencia(destino: Path, fonte: str, tam_cit: int):
    subprocess.run([pandoc_bin(), "-o", str(destino), "--print-default-data-file", "reference.docx"],
                   check=True)
    d = docx.Document(str(destino))
    S = {st.name: st for st in d.styles}  # busca pelo nome exato do XML
    S["_obj"] = d.styles

    # padrões do documento
    rpr_def = d.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    set_fonte_rpr(rpr_def, fonte)

    J, E, C = WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER
    corpo = dict(alinh=J, entrelinha=1.5, recuo1=1.25)
    for nome in ("Normal", "Body Text", "First Paragraph"):
        estilo_par(S[nome], fonte, **corpo)
    estilo_par(S["Compact"], fonte, alinh=J, entrelinha=1.5)
    # Título do documento e metadados
    estilo_par(S["Title"], fonte, negrito=True, caixa_alta=True, alinh=C, entrelinha=1.5, depois=12)
    estilo_par(S["Subtitle"], fonte, alinh=C, entrelinha=1.5, depois=12)
    estilo_par(S["Author"], fonte, alinh=C, entrelinha=1.5)
    estilo_par(S["Date"], fonte, alinh=C, entrelinha=1.5, depois=12)
    estilo_par(S["Abstract Title"], fonte, negrito=True, alinh=C, entrelinha=1.5, depois=12)
    estilo_par(S["Abstract"], fonte, alinh=J, entrelinha=1.5)
    # Seções — NBR 6024 (espaço de uma linha de 1,5 antes/depois ≈ 18 pt)
    niveis = [
        dict(negrito=True, caixa_alta=True),   # primária
        dict(caixa_alta=True),                 # secundária
        dict(negrito=True),                    # terciária
        dict(italico=True),                    # quaternária
        dict(),                                # quinária
    ]
    for i in range(1, 10):
        kw = niveis[min(i, 5) - 1]
        estilo_par(S[f"Heading {i}"], fonte, alinh=E, entrelinha=1.5, antes=18, depois=18,
                   manter_prox=True, **kw)
        ch = S[f"Heading {i} Char"]
        ch.font.color.rgb = PRETO
        ch.font.size = Pt(12)
        set_fonte_rpr(ch.element, fonte)
    # Citação longa — NBR 10520:2023
    estilo_par(S["Block Text"], fonte, tam=tam_cit, alinh=J, entrelinha=1, recuo_esq=4, antes=12, depois=12)
    bt = S["Block Text"].element.find(qn("w:pPr"))
    if bt is not None:  # remove bordas/sombreado do estilo padrão
        for tag in ("w:pBdr", "w:shd"):
            for e in bt.findall(qn(tag)):
                bt.remove(e)
    # Notas, legendas, tabelas
    estilo_par(S["Footnote Text"], fonte, tam=10, alinh=J, entrelinha=1)
    estilo_par(S["Footnote Block Text"], fonte, tam=10, alinh=J, entrelinha=1, recuo_esq=1)
    for nome in ("Caption", "Table Caption", "Image Caption"):
        estilo_par(S[nome], fonte, tam=10, alinh=E, entrelinha=1, antes=6, depois=6)
    estilo_par(S["Figure"], fonte, alinh=C, entrelinha=1)
    estilo_par(S["Captioned Figure"], fonte, alinh=C, entrelinha=1)
    estilo_par(S["Bibliography"], fonte, alinh=E, entrelinha=1, depois=12)
    estilo_par(S["Definition Term"], fonte, negrito=True, alinh=E, entrelinha=1.5)
    estilo_par(S["Definition"], fonte, alinh=J, entrelinha=1.5, recuo_esq=1.25)
    estilo_par(S["TOC Heading"], fonte, negrito=True, caixa_alta=True, alinh=C, entrelinha=1.5, depois=18)
    S["TOC Heading"].font.color.rgb = PRETO
    # Referências — NBR 6023
    ref = S["_obj"].add_style("Referencia", WD_STYLE_TYPE.PARAGRAPH)
    ref.base_style = S["Normal"]
    estilo_par(ref, fonte, alinh=E, entrelinha=1, depois=12)
    # Hiperlink em preto
    hl = S["Hyperlink"].font
    hl.color.rgb = PRETO
    hl.underline = False
    for nome in ("Verbatim Char",):
        S[nome].font.size = Pt(10)

    # Página A4 + margens + número no canto superior direito
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.left_margin = Cm(3), Cm(3)
    sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)
    sec.header_distance, sec.footer_distance = Cm(2), Cm(1.25)
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    r = campo(p, "PAGE")
    r.font.size = Pt(10)
    r.font.name = fonte
    d.save(str(destino))


# ----------------------------------------------------------- front matter
def ler_front_matter(md: str) -> dict:
    m = re.match(r"\A---\s*\n(.*?)\n(?:---|\.\.\.)\s*\n", md, re.S)
    if not m:
        return {}
    dados = {}
    for linha in m.group(1).splitlines():
        mm = re.match(r"^([A-Za-zÀ-ú_-]+)\s*:\s*(.*)$", linha)
        if mm:
            v = mm.group(2).strip().strip('"').strip("'")
            dados[mm.group(1).lower()] = v
    return dados


def contar_h1(md: str) -> int:
    n, fence = 0, False
    for linha in md.splitlines():
        if re.match(r"^\s*(```|~~~)", linha):
            fence = not fence
        elif not fence and re.match(r"^# \S", linha):
            n += 1
    return n


# ------------------------------------------------------------- pós-processo
def sectpr_basico(modelo):
    """Cópia de um sectPr (tamanho, margens) sem header nem quebra contínua."""
    s = copy.deepcopy(modelo)
    for e in list(s):
        if e.tag in (qn("w:headerReference"), qn("w:footerReference"), qn("w:pgNumType"),
                     qn("w:titlePg"), qn("w:vAlign"), qn("w:type")):
            s.remove(e)
    return s


ORDEM_SECTPR = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
                "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
                "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
                "printerSettings", "sectPrChange"]


def add_child_ordered(sectpr, el):
    """Insere respeitando a ordem do schema OOXML (o Word recusa fora de ordem)."""
    nome = el.tag.split("}")[1]
    idx = ORDEM_SECTPR.index(nome)
    for i, filho in enumerate(list(sectpr)):
        fn = filho.tag.split("}")[1]
        if fn in ORDEM_SECTPR and ORDEM_SECTPR.index(fn) > idx:
            filho.addprevious(el)
            return
    sectpr.append(el)


def novo_par(body, ref_el, texto_linhas, fonte, negrito=False, caixa=False, alinh="center",
             tam=12, recuo_esq_cm=0.0, entrelinha_simples=True, antes=False):
    """Cria um <w:p> antes de ref_el. texto_linhas: lista de linhas (quebra de linha entre elas)."""
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), "Normal"); ppr.append(ps)
    sp = OxmlElement("w:spacing"); sp.set(qn("w:before"), "0"); sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "240" if entrelinha_simples else "360"); sp.set(qn("w:lineRule"), "auto")
    ppr.append(sp)
    ind = OxmlElement("w:ind"); ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), str(int(recuo_esq_cm * 567)))
    ppr.append(ind)
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), alinh); ppr.append(jc)
    p.append(ppr)
    for i, linha in enumerate(texto_linhas):
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        set_fonte_rpr(rpr, fonte)
        if negrito:
            rpr.append(OxmlElement("w:b"))
        if caixa:
            rpr.append(OxmlElement("w:caps"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(tam * 2)); rpr.append(sz)
        r.append(rpr)
        if i:
            r.append(OxmlElement("w:br"))
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = linha
        r.append(t)
        p.append(r)
    ref_el.addprevious(p)
    return p


def fecha_secao(par_el, modelo_sectpr, vjust=False, pg_inicio=None):
    """Transforma o último parágrafo de uma página em fim de seção (próxima página)."""
    ppr = par_el.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr"); par_el.insert(0, ppr)
    s = sectpr_basico(modelo_sectpr)
    if pg_inicio is not None:
        pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), str(pg_inicio)); add_child_ordered(s, pg)
    if vjust:
        va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "both"); add_child_ordered(s, va)
    ppr.append(s)


def pos_processar(caminho: Path, fonte: str, meta: dict, capa: bool, sumario: bool,
                  quebrar: bool, tam_cit: int):
    d = docx.Document(str(caminho))
    body = d.element.body
    sect_final = body.find(qn("w:sectPr"))

    # 1. títulos sem indicativo numérico → centralizados
    SEM = re.compile(r"^(refer[eê]ncias|sum[aá]rio|resumo|abstract|agradecimentos|gloss[aá]rio|errata|"
                     r"bibliografia|lista de .*|[ií]ndice|ap[eê]ndice.*|anexo.*)$", re.I)
    primeira_secao = True
    for p in d.paragraphs:
        nome = p.style.name if p.style is not None else ""
        if nome.startswith("Heading") and SEM.match(p.text.strip()):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if quebrar and nome == "Heading 1":
            if not primeira_secao:
                p.paragraph_format.page_break_before = True
            primeira_secao = False

    # 1b. NBR 6024: indicativo numérico separado do título por um espaço (pandoc usa tabulação)
    for p in d.paragraphs:
        if p.style is not None and p.style.name.startswith("Heading"):
            for r in p._p.findall(qn("w:r")):
                tab = r.find(qn("w:tab"))
                if tab is not None and len(r) <= 2:
                    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = " "
                    tab.addprevious(t); r.remove(tab)
                    break
    # notas de rodapé com numeração contínua no documento todo
    for sp in body.iter(qn("w:sectPr")):
        for fp in sp.findall(qn("w:footnotePr")):
            sp.remove(fp)

    # 2. tabelas: bordas simples, texto 10 pt, espaço simples, cabeçalho em negrito
    for t in d.tables:
        tblpr = t._tbl.tblPr
        for e in tblpr.findall(qn("w:tblBorders")):
            tblpr.remove(e)
        b = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            x = OxmlElement(f"w:{lado}")
            x.set(qn("w:val"), "single"); x.set(qn("w:sz"), "4"); x.set(qn("w:color"), "000000")
            b.append(x)
        depois_de = [qn(f"w:{n}") for n in ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption")]
        alvo = next((e for e in tblpr if e.tag in depois_de), None)
        if alvo is not None:
            alvo.addprevious(b)
        else:
            tblpr.append(b)
        for i, row in enumerate(t.rows):
            for cell in row.cells:
                for cp in cell.paragraphs:
                    pf = cp.paragraph_format
                    pf.first_line_indent = Cm(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    pf.space_before = pf.space_after = Pt(2)
                    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for r in cp.runs:
                        r.font.size = Pt(10)
                        if i == 0:
                            r.font.bold = True

    # 3. parágrafos de lista: sem recuo de primeira linha;
    #    citação longa dentro de lista: volta ao recuo padrão de 4 cm da margem
    for p in d.paragraphs:
        ppr = p._p.pPr
        if ppr is None:
            continue
        if p.style is not None and p.style.name == "Block Text":
            for tag in ("w:numPr", "w:ind"):
                for e in ppr.findall(qn(tag)):
                    ppr.remove(e)
        elif ppr.numPr is not None:
            p.paragraph_format.first_line_indent = Cm(0)

    # 4. pré-textuais: capa, folha de rosto, sumário (sem número de página)
    primeiro = body[0]
    titulo = meta.get("title") or meta.get("titulo") or ""
    if capa:
        # remove título/autor/data gerados pelo pandoc (vão para a capa)
        for p in list(d.paragraphs)[:6]:
            if p.style is not None and p.style.name in ("Title", "Subtitle", "Author", "Date"):
                if not titulo and p.style.name == "Title":
                    titulo = p.text
                p._p.getparent().remove(p._p)
        primeiro = body[0]
        inst = [x for x in (meta.get("instituicao"), meta.get("unidade"), meta.get("curso")) if x]
        autor = meta.get("autor") or meta.get("author") or ""
        sub = meta.get("subtitulo") or meta.get("subtitle") or ""
        cidade, ano = meta.get("cidade", ""), str(meta.get("ano", ""))
        tit_linhas = [titulo + (":" if sub else "")] + ([sub] if sub else [])
        # CAPA
        novo_par(body, primeiro, inst or [" "], fonte, negrito=True, caixa=True)
        novo_par(body, primeiro, [autor or " "], fonte, caixa=True)
        novo_par(body, primeiro, [tit_linhas[0]], fonte, negrito=True, caixa=True)
        if sub:
            novo_par(body, primeiro, [sub], fonte)
        ult = novo_par(body, primeiro, [cidade or " ", ano or " "], fonte)
        fecha_secao(ult, sect_final, vjust=True, pg_inicio=0)
        # FOLHA DE ROSTO
        novo_par(body, primeiro, [autor or " "], fonte, caixa=True)
        novo_par(body, primeiro, [tit_linhas[0]], fonte, negrito=True, caixa=True)
        if sub:
            novo_par(body, primeiro, [sub], fonte)
        nat = [x for x in (meta.get("natureza"),) if x]
        if meta.get("orientador"):
            nat += ["", "Orientador(a): " + meta["orientador"]]
        if nat:
            novo_par(body, primeiro, nat, fonte, alinh="both", tam=12, recuo_esq_cm=8)
        ult = novo_par(body, primeiro, [cidade or " ", ano or " "], fonte)
        fecha_secao(ult, sect_final, vjust=True)

    if sumario:
        sdt = body.find(qn("w:sdt"))
        if sdt is not None:
            body.remove(sdt)
            primeiro_txt = body[0]
            if capa:
                # localiza o 1º parágrafo após a folha de rosto
                for el in body:
                    ppr = el.find(qn("w:pPr")) if el.tag == qn("w:p") else None
                    if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                        primeiro_txt = el.getnext()
            primeiro_txt.addprevious(sdt)
            p = OxmlElement("w:p")
            sdt.addnext(p)
            fecha_secao(p, sect_final, pg_inicio=None if capa else 0)
        # pede ao Word para atualizar campos (sumário) ao abrir
        st = d.settings.element
        if st.find(qn("w:updateFields")) is None:
            u = OxmlElement("w:updateFields"); u.set(qn("w:val"), "true"); st.append(u)

    d.save(str(caminho))


# ------------------------------------------------------------------- main
def converter(md_path: Path, saida: Path, a, ref: Path) -> str:
    md = md_path.read_text(encoding="utf-8-sig")
    meta = ler_front_matter(md)
    for k in ("instituicao", "curso", "autor", "cidade", "ano", "natureza", "orientador", "subtitulo", "titulo"):
        v = getattr(a, k, None)
        if v:
            meta[k] = v
    capa = a.capa or str(meta.get("capa", "")).lower() in ("true", "sim", "yes")
    cmd = [pandoc_bin(), str(md_path), "-f", "commonmark_x+yaml_metadata_block", "-t", "docx",
           "--reference-doc", str(ref), "--lua-filter", str(LUA), "-o", str(saida),
           "--resource-path", str(md_path.parent)]
    if contar_h1(md) == 1:
        cmd.append("--shift-heading-level-by=-1")   # o único H1 vira título do documento
    if a.sumario:
        cmd += ["--toc", "--toc-depth", str(a.profundidade), "-M", "toc-title=Sumário"]
    if a.numerar:
        cmd.append("--number-sections")
    if a.manter_aspas:
        cmd += ["-M", "abnt-manter-aspas=true"]
    if meta.get("titulo") and not meta.get("title"):
        cmd += ["-M", f"title={meta['titulo']}"]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        return f"ERRO {md_path.name}: {r.stderr.strip()[:500]}"
    if not meta.get("title") and not meta.get("titulo") and contar_h1(md) == 1:
        m = re.search(r"^# (.+)$", md, re.M)
        meta["title"] = m.group(1).strip() if m else ""
    pos_processar(saida, FONTES[a.fonte], meta, capa, a.sumario, a.quebrar_secoes, a.tam_citacao)
    avisos = f" pandoc: {r.stderr.strip()[:200]}" if r.stderr.strip() else ""
    return f"OK {saida}{avisos}"


def main():
    ap = argparse.ArgumentParser(description="Markdown → DOCX no padrão ABNT")
    ap.add_argument("entradas", nargs="+", help="arquivos .md ou pastas")
    ap.add_argument("-o", "--saida", help="arquivo .docx (1 entrada) ou pasta")
    ap.add_argument("--fonte", choices=FONTES, default="times")
    ap.add_argument("--tam-citacao", type=int, default=10, help="fonte da citação longa (padrão 10)")
    ap.add_argument("--capa", action="store_true", help="gera capa e folha de rosto")
    ap.add_argument("--sumario", action="store_true", help="insere sumário automático")
    ap.add_argument("--profundidade", type=int, default=3, help="níveis no sumário")
    ap.add_argument("--numerar", action="store_true", help="numera as seções (1, 1.1, 1.1.1)")
    ap.add_argument("--quebrar-secoes", action="store_true", help="seção primária em nova página")
    ap.add_argument("--manter-aspas", action="store_true", help="não remove aspas das citações longas")
    for k in ("instituicao", "curso", "autor", "titulo", "subtitulo", "cidade", "ano", "natureza", "orientador"):
        ap.add_argument(f"--{k}")
    a = ap.parse_args()

    arquivos = []
    for e in a.entradas:
        p = Path(e)
        arquivos += sorted(p.glob("*.md")) if p.is_dir() else [p]
    if not arquivos:
        sys.exit("ERRO nenhum .md encontrado")

    with tempfile.TemporaryDirectory() as tmp:
        ref = Path(tmp) / "abnt-ref.docx"
        criar_referencia(ref, FONTES[a.fonte], a.tam_citacao)
        ok = 0
        for md in arquivos:
            if a.saida and len(arquivos) == 1 and a.saida.lower().endswith(".docx"):
                out = Path(a.saida)
            elif a.saida:
                Path(a.saida).mkdir(parents=True, exist_ok=True)
                out = Path(a.saida) / (md.stem + ".docx")
            else:
                out = md.with_suffix(".docx")
            msg = converter(md, out, a, ref)
            ok += msg.startswith("OK")
            print(msg)
        print(f"RESUMO {ok}/{len(arquivos)} convertidos")
        sys.exit(0 if ok == len(arquivos) else 1)


if __name__ == "__main__":
    main()
