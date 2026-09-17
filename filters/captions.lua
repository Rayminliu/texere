-- 在 AST 层把表题/图注标记成语义样式，post.py 之后按样式名处理，
-- 不再用正则反推"这段是不是题注"。
--
-- 判定优先级（结构优先，文本兜底）：
--   1. 段落后面紧跟一个 Table          -> TableCaption
--   2. 段落只含一张图片                -> FigureCaption
--   3. 文本以 表/图/Table/Figure 开头  -> 对应样式
--
-- 必须配合 ref.docx 里已定义的同名样式使用（见 make_ref.py）。

local TABLE_STYLE = "TableCaption"
local FIGURE_STYLE = "FigureCaption"

local function text_head(inlines)
  for _, inline in ipairs(inlines) do
    if inline.t == "Str" then
      return inline.text
    end
  end
  return nil
end

-- 题注关键字默认值。可用 config 的 caption_words 覆盖：
-- render.py 会以 -M dk-table-words=... -M dk-figure-words=... 传进来，
-- post.py 那边同步重建正则，两边始终一致。
local DEFAULT_TABLE_WORDS = { "表", "表格", "圖片", "Table" }
local DEFAULT_FIGURE_WORDS = { "图", "圖", "图片", "圖片", "Figure", "Fig" }

local function meta_words(meta, key, default)
  local v = meta[key]
  if v == nil then
    return default
  end
  local s = pandoc.utils.stringify(v)
  local out = {}
  for w in s:gmatch("[^,%s]+") do
    out[#out + 1] = w
  end
  if #out == 0 then
    return default
  end
  return out
end

local function starts_any(s, words)
  if not s then
    return false
  end
  for _, w in ipairs(words) do
    if s:sub(1, #w) == w then
      return true
    end
  end
  return false
end

local function single_image(para)
  local n, img = 0, nil
  for _, inline in ipairs(para.content) do
    if inline.t == "Image" then
      n = n + 1
      img = inline
    elseif inline.t == "Str" and inline.text:match("^%s*$") then
      -- 纯空白不计
    elseif inline.t == "Space" or inline.t == "SoftBreak" then
      -- 忽略
    else
      return nil
    end
  end
  if n == 1 then
    return img
  end
  return nil
end

function Pandoc(doc)
  local tw = meta_words(doc.meta, "dk-table-words", DEFAULT_TABLE_WORDS)
  local fw = meta_words(doc.meta, "dk-figure-words", DEFAULT_FIGURE_WORDS)
  local blocks = doc.blocks
  for i, block in ipairs(blocks) do
    if block.t == "Para" then
      local kind = nil
      local nxt = blocks[i + 1]
      if nxt ~= nil and nxt.t == "Table" then
        kind = TABLE_STYLE
      elseif single_image(block) then
        kind = FIGURE_STYLE
      else
        local head = text_head(block.content)
        if starts_any(head, tw) then
          kind = TABLE_STYLE
        elseif starts_any(head, fw) then
          kind = FIGURE_STYLE
        end
      end
      if kind ~= nil then
        blocks[i] = pandoc.Div({ block },
          pandoc.Attr("", {}, { ["custom-style"] = kind }))
      end
    end
  end
  return doc
end
