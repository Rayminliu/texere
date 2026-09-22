# CONFIG.md

> 从 `README.md` 下沉而来的配置参考。中英两份由 `tests/test_docs_sync.py` 守着同构。

## Configuration

### config.json fields

| Field | Purpose |
|---|---|
| `cover` | Cover lines `[[style_name, text], …]`; style names: CoverTop / CoverTitle / CoverSub / CoverInfo / CoverDate |
| `header` | Header text for the body section |
| `title` / `author` / `subject` / `comments` | Document properties |
| `reference_doc` | Use a specific template docx (for a client-mandated format) |
| `toc_heading` | TOC title, default 「目　　录」 |
| `toc` | `false` → no TOC page (short notices / announcements); headings keep their styles, and with no cover the document stays a single section |
| `style` | Fine-grained layout, see below |
| `caption_words` | Custom caption keywords (default 表/图/Table/Figure), see below |
| `resource_paths` | Extra directories to search for images |
| `content_fixes` | Editorial replacement table `[[old, new], …]`, applied after merging and before conversion |
| `content_fixes_file` | Replacement table file (`.json`, or the `CONTENT_FIXES` literal inside a `.py`; relative paths resolve against the config file's directory) |

> `content_fixes` is for dropping explanatory parentheticals or unifying wording. When it points at an
> existing `.py`, the values are read with `ast` — **parsed, never executed** — so you don't have to keep a
> second copy of the rules.

`caption_words` (only needed for non-Chinese or unusual caption keywords; `post.py` and the Lua filter are
kept in sync):

```json
"caption_words": {"table": ["表", "表格"], "figure": ["图", "图片"]}
```

### The `style` section

**This table is the single source for `style` keys and their defaults** — other sections link here instead of
restating them. Everything is optional; the defaults are the Chinese formal-document conventions:

| Group | Keys | Default |
|---|---|---|
| Footer | `page_number` / `page_number_size` | `— {n} —` / `9` |
| Header | `header_size` / `header_gray` / `header_rule_color` / `header_rule_size` | `9` / `595959` / `BFBFBF` / `4` |
| TOC | `toc_depth` / `toc_title_size` / `toc_title_color` / `toc_placeholder` / `toc_placeholder_size` | `1-2` / `16` / `000000` / see source / `12` |
| Captions | `caption_gray` / `caption_size` / `caption_space_before` / `caption_space_after` | `404040` / `10.5` / `6` / `4` |
| Captions | `caption_keep_with_next` | `true` (keeps a table caption with its table; turning it off saves a page in testing, but the caption may strand at a page foot) |
| Tables | `header_rows` | auto — **per table**, by reading the `w:tblHeader` pandoc emits; a number forces it; `0` = this table has no header (forms / appendix tables whose first row is a field name), which also drops the repeating-header flag. See [Visual control](TABLES.md#visual-control) |
| Tables | `table_border` / `table_shade` / `table_size` | `full` / `EDEDED` / `10.5` |
| Tables | `table_header_color` / `table_zebra` / `table_zebra_fill` | unset / `false` / `F7F7F7` |
| Tables | `cell_margin_v` / `cell_margin_h` / `table_para_space` | `40` / `80` / `1` |
| Tables | `border_size` / `border_color` / `three_line_size` | `6` / `808080` / `12` |
| Fonts | `east_font` / `latin_font` | `宋体` / `Times New Roman` (**tables and captions only**) |

> Colours are 6-digit hex (`404040`); border weights are in 1/8 pt; spacing is in pt; cell margins are in twips.

