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

local function kind_by_text(s)
  if not s then
    return nil
  end
  if s:match("^表") or s:match("^表格") or s:match("^圖片") then
    return TABLE_STYLE
  end
  if s:match("^图") or s:match("^圖") or s:match("^图片") then
    return FIGURE_STYLE
  end
  local lower = s:lower()
  if lower:match("^table") then
    return TABLE_STYLE
  end
  if lower:match("^figure") or lower:match("^fig%.") then
    return FIGURE_STYLE
  end
  return nil
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
        kind = kind_by_text(text_head(block.content))
      end
      if kind ~= nil then
        blocks[i] = pandoc.Div({ block },
          pandoc.Attr("", {}, { ["custom-style"] = kind }))
      end
    end
  end
  return doc
end
