# abnt-docx

Skill para o [Claude Code](https://claude.com/claude-code) que converte arquivos **Markdown** em documentos **Word (.docx) formatados nas normas da ABNT**: Times New Roman ou Arial 12, A4, margens 3/2 cm, entrelinha 1,5, citações longas recuadas, títulos de seção segundo a NBR 6024, paginação no canto superior direito e, se você quiser, capa, folha de rosto e sumário.

Funciona com resumos, fichamentos, resenhas, trabalhos de disciplina, TCCs e artigos escritos em Markdown (inclusive os gerados por IA).

## Instalação

```bash
git clone https://github.com/talesalt/abnt-docx ~/.claude/skills/abnt-docx
py -3 -m pip install --user python-docx      # ou: python3 -m pip install python-docx
```

Também é preciso ter o [pandoc](https://pandoc.org) 3 ou mais recente (`winget install JohnMacFarlane.Pandoc` / `brew install pandoc`).

## Uso

No Claude Code, basta pedir: *"converte esse resumo.md em docx ABNT com Arial"*. Pelo terminal:

```bash
py -3 ~/.claude/skills/abnt-docx/scripts/md2abnt.py resumo.md
py -3 ~/.claude/skills/abnt-docx/scripts/md2abnt.py resenha.md --fonte arial --capa --sumario \
   --autor "Nome" --instituicao "Universidade de São Paulo" --cidade "São Paulo" --ano 2026
py -3 ~/.claude/skills/abnt-docx/scripts/md2abnt.py pasta_com_mds/ -o saida/
```

| Opção | Efeito |
|---|---|
| `--fonte times\|arial` | fonte do documento (padrão Times New Roman) |
| `--capa` | capa e folha de rosto (dados por argumentos ou front matter YAML) |
| `--sumario` | sumário automático (o Word pergunta se deve atualizar os campos: responda Sim) |
| `--numerar` | numeração progressiva das seções (1, 1.1, 1.1.1) |
| `--quebrar-secoes` | cada seção primária começa em nova página |
| `--manter-aspas` | não remove as aspas das citações longas |
| `--tam-citacao N` | tamanho da fonte da citação longa (padrão 10) |

Veja [SKILL.md](SKILL.md) para o mapeamento Markdown → ABNT e [references/regras-abnt.md](references/regras-abnt.md) para as regras, com a norma de origem de cada uma.

## Revisão acadêmica do texto (opcional)

Antes de converter, a skill pergunta se você quer melhorar o texto. Se disser sim, ela analisa o .md e mostra um **relatório numerado** do que não combina com texto acadêmico. Exemplos: texto todo em tópicos, linhas horizontais, emojis, comentários informais em itálico, rótulos de fichamento (**Conceito:**), excesso de negrito, setas e citações introduzidas por rótulo. Você marca quais correções quer, e ela grava um `arquivo_academico.md` (o original fica intacto) antes de gerar o .docx.

```bash
py -3 ~/.claude/skills/abnt-docx/scripts/revisar_md.py analisar resumo.md
py -3 ~/.claude/skills/abnt-docx/scripts/revisar_md.py aplicar resumo.md --itens bullets,regua,notas-editoriais
```

A correção `bullets` transforma o 1º nível de tópicos em parágrafos corridos, e os subníveis sobem um nível. Itens marcados como `manual` (setas, chamadas de citação etc.) são reescritos pelo Claude, trecho por trecho.

## Como funciona

1. Gera um modelo `reference.docx` com os estilos ABNT (a partir do modelo padrão do pandoc).
2. O pandoc converte o Markdown (CommonMark com tabelas, notas e front matter) com um filtro Lua (`scripts/abnt.lua`) que tira as aspas das citações longas, marca a seção de referências e os títulos sem numeração e limpa links locais.
3. O `python-docx` ajusta o resultado: tabelas, citações dentro de listas, espaço entre número e título, capa, folha de rosto, sumário e paginação por seções.

## Sobre as normas

As normas ABNT são protegidas por direito autoral e vendidas pela ABNT, então **o texto delas não está neste repositório**. As regras de formatação foram conferidas em guias públicos de bibliotecas universitárias (UFV 2025, UFRGS, IFCE, Unifor), listados em `references/regras-abnt.md`. Para consultar as normas completas, use a **ABNT Coleção** pela sua universidade.

## Outras skills que inspiraram esta

- [anthropics/skills — docx](https://github.com/anthropics/skills): manipulação geral de .docx.
- `latex-academic-br` e `abnt2-review` (thejoaov): trabalhos ABNT em LaTeX/abnTeX2.
- `markdown-to-docx` (duc01226/easyplatform): Markdown → Word sem regras ABNT.

## Licença

MIT
