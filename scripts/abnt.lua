-- abnt.lua — filtro pandoc para a skill abnt-docx
-- * citação longa (blockquote): remove aspas externas (NBR 10520:2023)
-- * seção "Referências": parágrafos/itens viram estilo "Referencia" (NBR 6023)
-- * links file:// viram texto simples; linhas horizontais (---) são removidas
-- * títulos sem indicativo numérico (Sumário, Referências...) ficam sem número

local manter_aspas = false
local manter_linhas = false

local SEM_NUMERO = {
  ["referências"] = true, ["referencias"] = true, ["sumário"] = true,
  ["sumario"] = true, ["resumo"] = true, ["abstract"] = true,
  ["agradecimentos"] = true, ["glossário"] = true, ["glossario"] = true,
  ["errata"] = true, ["lista de ilustrações"] = true, ["lista de tabelas"] = true,
  ["lista de abreviaturas e siglas"] = true, ["lista de siglas"] = true,
  ["lista de símbolos"] = true, ["índice"] = true, ["bibliografia"] = true,
}

local function norm(s)
  s = pandoc.text.lower(s)
  s = s:gsub("^%s+", ""):gsub("%s+$", "")
  return s
end

local function eh_sem_numero(txt)
  local t = norm(txt)
  if SEM_NUMERO[t] then return true end
  if t:match("^apêndice") or t:match("^anexo") then return true end
  return false
end

local function eh_referencias(txt)
  local t = norm(txt)
  return t == "referências" or t == "referencias" or t == "bibliografia"
end

-- aspas de abertura/fechamento aceitas
local ABRE = { '"', '“', '«', '„' }
local FECHA = { '"', '”', '»' }

local function tira_prefixo(s, lista)
  for _, q in ipairs(lista) do
    if s:sub(1, #q) == q then return s:sub(#q + 1), true end
  end
  return s, false
end

local function tira_sufixo(s, lista)
  for _, q in ipairs(lista) do
    if #s >= #q and s:sub(-#q) == q then return s:sub(1, -#q - 1), true end
  end
  return s, false
end

-- primeiro/último Str de uma lista de inlines (desce em Emph/Strong)
local function borda_str(inlines, ultimo)
  local i0, i1, passo = 1, #inlines, 1
  if ultimo then i0, i1, passo = #inlines, 1, -1 end
  for i = i0, i1, passo do
    local el = inlines[i]
    if el.t == "Str" then return el end
    if el.t == "Quoted" then return el end
    if el.content and type(el.content) == "table" and el.t ~= "Note" then
      local r = borda_str(el.content, ultimo)
      if r then return r end
    end
    if el.t ~= "Space" and el.t ~= "SoftBreak" and el.t ~= "LineBreak" then
      return nil
    end
  end
  return nil
end

local function tira_aspas_bloco(bq)
  local blocos = bq.content
  local primeiro, ultimo
  for _, b in ipairs(blocos) do
    if b.t == "Para" or b.t == "Plain" then primeiro = primeiro or b; ultimo = b end
  end
  if not primeiro then return bq end
  local a = borda_str(primeiro.content, false)
  local z = borda_str(ultimo.content, true)
  if not a or not z then return bq end
  -- caso Quoted (smart quotes já interpretadas): desembrulha
  if a.t == "Quoted" and a == z and #primeiro.content == 1 then
    primeiro.content = a.content
    return bq
  end
  if a.t ~= "Str" or z.t ~= "Str" then return bq end
  local na, ok1 = tira_prefixo(a.text, ABRE)
  if not ok1 then return bq end
  -- sufixo pode vir seguido de pontuação: "texto". ou "texto".
  local zt = (a == z) and na or z.text
  local pont = ""
  local zt2, ok2 = tira_sufixo(zt, FECHA)
  if not ok2 then
    local corpo, p = zt:match("^(.-)([%.;,]+)$")
    if corpo then
      zt2, ok2 = tira_sufixo(corpo, FECHA)
      pont = p
    end
  end
  if not ok2 then return bq end
  if a == z then
    a.text = zt2 .. pont
  else
    a.text = na
    z.text = zt2 .. pont
  end
  return bq
end

local function inlines_texto(inl)
  return pandoc.utils.stringify(inl)
end

function Meta(m)
  if m["abnt-manter-aspas"] then manter_aspas = true end
  if m["abnt-manter-linhas"] then manter_linhas = true end
end

function Link(el)
  if el.target:match("^file:") then return el.content end
end

function Header(h)
  if eh_sem_numero(inlines_texto(h.content)) then
    h.classes:insert("unnumbered")
  end
  return h
end

local function para_referencia(b)
  return pandoc.Div({ pandoc.Para(b.content) }, { ["custom-style"] = "Referencia" })
end

function Pandoc(doc)
  local out = pandoc.List()
  local em_ref, nivel_ref = false, 0
  for _, b in ipairs(doc.blocks) do
    if b.t == "Header" then
      if em_ref and b.level <= nivel_ref then em_ref = false end
      if eh_referencias(inlines_texto(b.content)) then em_ref, nivel_ref = true, b.level end
      out:insert(b)
    elseif b.t == "HorizontalRule" and not manter_linhas then
      -- descarta
    elseif em_ref and (b.t == "Para" or b.t == "Plain") then
      out:insert(para_referencia(b))
    elseif em_ref and (b.t == "BulletList" or b.t == "OrderedList") then
      for _, item in ipairs(b.content) do
        for _, ib in ipairs(item) do
          if ib.t == "Para" or ib.t == "Plain" then out:insert(para_referencia(ib)) else out:insert(ib) end
        end
      end
    else
      out:insert(b)
    end
  end
  doc.blocks = out
  -- citações longas: varre todos os níveis (inclusive dentro de listas)
  if not manter_aspas then
    doc = doc:walk({ BlockQuote = tira_aspas_bloco })
  end
  return doc
end

return {
  { Meta = Meta },
  { Link = Link, Header = Header },
  { Pandoc = Pandoc },
}
