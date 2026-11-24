---
name: abnt-docx
description: Converte arquivos Markdown (.md) em documentos Word (.docx) formatados nas normas da ABNT (NBR 14724:2024, 6024, 10520:2023, 6023), em Times New Roman ou Arial 12, margens 3/2 cm, entrelinha 1,5, citação longa com recuo de 4 cm, títulos de seção, paginação, capa, folha de rosto e sumário opcionais. Antes de converter, oferece revisar o texto para o padrão acadêmico (tira o 1º nível de bullets, comentários informais, emojis, réguas etc.) com relatório numerado e escolha item a item. Use quando o usuário pedir para transformar .md/resumo/fichamento/resenha/trabalho em .docx ou Word "no padrão ABNT", "formatado ABNT", "normas da ABNT", trabalho acadêmico, TCC ou artigo em Word. Tudo por script, gastando poucos tokens.
---

# abnt-docx

Converte `.md` → `.docx` ABNT com **pandoc + python-docx**. Não leia o .docx gerado: confie na linha `OK`/`ERRO`.

`S` = pasta `scripts/` desta skill (ex.: `~/.claude/skills/abnt-docx/scripts`). Rode com `py -3` (Windows) ou `python3`.

## Protocolo (siga sempre, nesta ordem)

1. **Pergunte antes de converter**, com a ferramenta `AskUserQuestion`: "Quer que eu melhore o texto para o padrão acadêmico antes de gerar o .docx?". Opções: `Sim, revisar o texto (Recomendado)` e `Não, só formatar`. Pergunte também a fonte (Times New Roman / Arial) se a pessoa não disse. Se responder **Não**, pule para o passo 5.
2. **Analise** (não leia o .md inteiro; confie no relatório):
   ```bash
   py -3 "$S/revisar_md.py" analisar arquivo.md
   ```
   Sai um **relatório numerado**: `N. [chave] (auto|manual, Kx) [recomendado] descrição` + um exemplo. Detecta: texto em tópicos (1º nível de bullets), linhas horizontais, emojis, links locais, comentários informais em itálico (*(gancho!)*), rótulos de fichamento (**Conceito:**), excesso de negrito, `código`, `~` como "aproximadamente", exclamações, setas (→), CAIXA ALTA no texto, citações introduzidas por rótulo e parágrafo "Objetivo:" do arquivo.
3. **Mostre o relatório e deixe a pessoa escolher** com `AskUserQuestion` em `multiSelect: true`: uma opção por achado (label curto com o número e a chave, ex.: `1. Tirar bullets do 1º nível`; descrição com a contagem e o exemplo). Coloque os `[recomendado]` primeiro, com "(Recomendado)" no label. São no máximo 4 opções por pergunta e 4 perguntas por chamada: agrupe em perguntas como "Estrutura", "Limpeza", "Destaques" e "Reescrita manual". Se houver mais de 16 achados, faça uma segunda chamada.
4. **Aplique só o que foi escolhido.** O original nunca é alterado; o resultado vai para `arquivo_academico.md`:
   ```bash
   py -3 "$S/revisar_md.py" aplicar arquivo.md --itens bullets,regua,notas-editoriais
   ```
   (`--itens recomendados` aplica todos os automáticos recomendados.) Para os itens **manuais** escolhidos (`setas`, `caixa-alta`, `chamadas-citacao`, `meta`), edite o `arquivo_academico.md` com a ferramenta Edit, trecho por trecho (use `grep -n` para achar as linhas). Ex.: `**Citação longa (Fulgencio, 2022, p. 159):**` → `Conforme Fulgencio (2022, p. 159):`; `A → B` → `A, o que leva a B`. Nunca altere o texto dentro de citações (`>`).
5. **Converta** o `arquivo_academico.md` (ou o original, se a pessoa disse Não) com `md2abnt.py` (abaixo).
6. **Relate** em poucas linhas o que mudou, por item: a linha `OK ... (chave=quantidade, ...)` do script e as edições manuais. Diga o caminho do .md revisado e do .docx.

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
