---
name: abnt-docx
description: Converte arquivos Markdown (.md) em documentos Word (.docx) formatados nas normas da ABNT (NBR 14724:2024, 6024, 10520:2023, 6023), em Times New Roman ou Arial 12, margens 3/2 cm, entrelinha 1,5, citação longa com recuo de 4 cm, títulos de seção, paginação, capa, folha de rosto e sumário opcionais. Use quando o usuário pedir para transformar .md/resumo/fichamento/resenha/trabalho em .docx ou Word "no padrão ABNT", "formatado ABNT", "normas da ABNT", trabalho acadêmico, TCC ou artigo em Word. Tudo por script, gastando poucos tokens.
---

# abnt-docx

Converte `.md` → `.docx` ABNT com **pandoc + python-docx**. Não leia o .docx gerado: confie na linha `OK`/`ERRO`.

`S` = pasta `scripts/` desta skill (ex.: `~/.claude/skills/abnt-docx/scripts`). Rode com `py -3` (Windows) ou `python3`.

## Uso

```bash
py -3 "$S/md2abnt.py" arquivo.md [outro.md | pasta/] [-o saida.docx | -o pasta/]
      [--fonte times|arial] [--capa] [--sumario] [--numerar] [--quebrar-secoes]
      [--manter-aspas] [--tam-citacao 10] [--profundidade 3]
      [--instituicao ... --curso ... --autor ... --titulo ... --subtitulo ...
       --cidade ... --ano ... --natureza ... --orientador ...]
```

- Sem `-o`, grava `<nome>.docx` ao lado do `.md`. Saída: `OK <arquivo>` por arquivo + `RESUMO n/m`.
- `--fonte`: `times` (padrão, Times New Roman) ou `arial`. Pergunte ao usuário se ele não disser.
- `--capa`: gera **capa + folha de rosto** (a capa não é contada; a folha de rosto conta, sem número impresso).
- `--sumario`: sumário automático (campo do Word; o Word pede para atualizar campos ao abrir — responder **Sim**).
- `--numerar`: numera seções `1`, `1.1`, `1.1.1` (NBR 6024). Não use se o .md já traz números nos títulos.
- `--quebrar-secoes`: cada seção primária (`#`) começa em nova página (exigido em TCC/dissertação).
- `--manter-aspas`: por padrão as aspas externas das citações longas (`> "..."`) são removidas (NBR 10520).

Os dados da capa podem vir do **front matter YAML** do .md (os argumentos da linha de comando têm prioridade):

```yaml
---
titulo: Título do trabalho
subtitulo: subtítulo
autor: Nome Completo
instituicao: Universidade de São Paulo
curso: Instituto de Psicologia
cidade: São Paulo
ano: 2026
natureza: Trabalho apresentado à disciplina X como requisito parcial de avaliação.
orientador: Prof. Dr. Fulano
capa: true
---
```

## Como o Markdown vira ABNT

| Markdown | Resultado |
|---|---|
| único `# Título` no topo | título do documento (centralizado, negrito, CAIXA ALTA); `##` passa a ser seção primária |
| `#`/`##`/`###`/`####`/`#####` | seção primária **NEGRITO CAIXA ALTA**, secundária CAIXA ALTA, terciária **negrito**, quaternária *itálico*, quinária normal |
| `# Referências`, `Sumário`, `Resumo`, `Apêndice`, `Anexo`... | título centralizado, sem número |
| parágrafos sob `Referências` (ou lista) | espaço simples, alinhados à esquerda, separados por uma linha |
| `> citação` (inclusive dentro de listas) | recuo de 4 cm, fonte 10, espaço simples, sem aspas |
| `[^1]` nota | nota de rodapé fonte 10, espaço simples, numeração contínua |
| tabela | bordas finas, fonte 10, espaço simples, cabeçalho em negrito |
| `---` (linha horizontal) | removida; links `file:///` viram texto simples |

Regras completas e fontes: `references/regras-abnt.md`. Leia só se o usuário perguntar sobre uma regra específica.

## Requisitos

- `pandoc` ≥ 3 no PATH (`winget install JohnMacFarlane.Pandoc`)
- `py -3 -m pip install --user python-docx`

## Limites (diga ao usuário quando relevante)

- Não reescreve o conteúdo: chamadas de citação, referências (NBR 6023) e títulos de ilustrações seguem como estão no .md.
- Elementos como ficha catalográfica, folha de aprovação, resumo em língua estrangeira e listas de ilustrações não são gerados automaticamente; escreva-os no .md como seções.
- Normas institucionais (ex.: Diretrizes da USP) podem diferir em detalhes; ajuste com as opções acima.
