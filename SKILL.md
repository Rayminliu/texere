---
name: texere
description: Renders Markdown into properly typeset Chinese formal documents (docx + PDF) — tenders and bids, project plans, grant applications, final reports, white papers — and verifies the result by opening it in real Microsoft Word, refreshing the TOC field, exporting a PDF, and checking for blank pages and layout drift. This skill should be used when a task involves generating a formal Chinese document from Markdown, adding a cover page, table of contents, header or per-section page numbers, applying Chinese formal typesetting (SimSun body, SimHei headings, 2-character first-line indent, centred captions), repeating table header rows across pages, or running pre-delivery checks for blank pages and layout drift. It also makes targeted edits to an existing docx — replacing text, inserting or deleting paragraphs, setting table cells, rewriting headers and footers — without re-typesetting the rest. It should not be used for tracked changes or comments, for theses, or for English documents.
license: MIT
compatibility: Requires Windows with a local Microsoft Word (COM), pandoc 3.1 or newer, Python 3.10 or newer, python-docx and lxml. On Linux and macOS only the docx half works — the Word acceptance and PDF export chain is unavailable.
---

# texere

Renders a Markdown chapter directory into a production-grade Chinese formal document (docx + PDF).

Two ideas carry the whole design: **template-driven** (every visual rule lives in `assets/ref.docx`; the
Markdown source carries semantics only) and **real-machine acceptance** (open the result in real Word, export
a PDF, then look at the rendered pages — never trust "it read back fine").

## Quick start

```bash
python scripts/render.py --doctor     # environment self-check: pandoc / deps / Word engine identity
python scripts/render.py --version
python scripts/render.py --sample     # smoke test -> sample_out.docx/.pdf

# Real render (--check implies --pdf)
python scripts/render.py --src chapters/ --out bid.docx --config cfg.json --pdf --check

python -m pytest -q                          # 62 assertions, no Word needed
python scripts/snapshot.py bid.pdf           # layout regression; exit 1 on drift
python scripts/snapshot.py bid.pdf --update  # re-record baseline after an intended change
python scripts/make_ref.py --body-font 楷体   # rebuild the typesetting template

# Edit an existing docx (see "Editing an existing docx")
python scripts/edit.py bid.docx --list
python scripts/edit.py bid.docx --replace "示例科技=某某科技" --verify
```

Run from the repository root, or pass absolute paths — `--out` is resolved against the current working
directory.

Dependencies: `pandoc >= 3.1` (required), `python-docx` + `lxml` (required), `pywin32` (`--pdf` only),
`PyMuPDF` (`--check` only), a local Microsoft Word (`--pdf` only). This is a folder of scripts, not a pip
package — install with `pip install -r requirements.txt`.

## When to use / when not to

Use it for:

- Generating a Chinese formal document from Markdown (bid, grant application, final report, white paper)
- Cover page, TOC field, per-section page numbers, header, table/figure caption conventions, repeating
  table header rows across pages
- One-command pre-delivery acceptance (does Word open it, are there blank pages, did the layout drift)
- **Targeted edits to an existing docx** — replacing text, inserting or deleting paragraphs, setting table
  cells, rewriting headers and footers, without re-typesetting the rest

Do **not** use it for:

- **Comments, tracked changes, redaction, watermarks, content controls** on an existing docx. Not
  supported, and out of scope by design — use a general-purpose docx skill for those (for example
  Codex's Documents plugin, which ships 36 OOXML scripts).
- **Theses or book manuscripts** — bibliography (citeproc), equation numbering and odd/even page headers
  are not supported.
- **English-language documents** — the defaults (A4, SimSun/SimHei, full-width punctuation, Chinese caption
  keywords) are Chinese-document conventions. `caption_words` changes the keywords, not the typography.

## Input requirements

- `--src` points at a directory; its `*.md` files merge **in filename order** (`01_`, `02_` prefixes control
  the order). `#` is a chapter, `##` a section.
- A table caption is its own line: `表 1-1 Caption`. A figure caption goes in the image alt text:
  `![图 1-1 Caption](a.png){width=13cm}`.
- Image search path: the `src` directory, all its subdirectories, and **the parent of `src` plus that
  parent's subdirectories** (images often live in a sibling directory); add more with `resource_paths`.
- **Do not put content before the first `#`** — it is neither treated as a cover nor kept in place (it lands
  after the TOC). `post.py` prints `[warn]` but will not delete it.
- config.json and source `.md` may carry a UTF-8 BOM.
- **Form-style / appendix documents** with no `#` headings are supported: no TOC, no section split, cover plus
  table/caption typesetting only. Set `"style": {"header_rows": 0}` so the first row is not shaded as a header.

## Examples

### Example 1 — minimal tender document

`config.json`:

```json
{
  "cover": [["CoverTitle", "投 标 文 件"], ["CoverInfo", "投标人：示例科技有限公司"]],
  "header": "示例项目 · 投标文件"
}
```

`01_bid.md`:

```markdown
# 第一章 投标函

## 1.1 致招标人

**示例科技有限公司** 谨按招标文件要求提交本投标文件。

表 1-1 商务条款响应表

| 条款 | 招标要求 | 投标响应 |
|:---|:---|:---|
| 工期 | 180 日历天 | 完全响应 |
```

Result: cover page, TOC field, chapter/section headings from the template, greyed centred caption above the
table, shaded repeating header row, `— n —` page numbers, configured header on every body page.

### Example 2 — grid table with a multi-level header

Use a grid table (not a pipe table) for complex headers, merged cells, in-cell line breaks and column widths.
`+===+` separates header from body; omitting an inner `|` spans columns:

```
+------------------+------------------+
| 商务部分         | 技术部分         |
+--------+---------+--------+---------+
| 条款   | 响应    | 模块   | 说明   |
+========+=========+========+=========+
| 工期   | 完全响应| 接入层 | 设备   |
+--------+---------+--------+---------+
```

Column widths come from the **narrowest column separator row**: the row `+--------+---------+` above yields
8:9:8:9. To get a narrow "No." column, write it narrow.

### Example 3 — figure with a caption

```markdown
![图 2-1 总体架构图](arch.png){width=13cm}
```

The caption text lives in the alt text; `post.py` centres it, greys it and keeps it with the image.

### Example 4 — form-style document (no headings)

No `#` heading at all; the first row is a field name, not a column header:

```json
{"cover": [["CoverTitle", "项目申报书"]], "style": {"header_rows": 0}}
```

```markdown
| 项目名称 | 示例智慧园区平台 |
| 申报单位 | 示例科技有限公司 |
```

Result: no TOC, no section split, first row left unshaded and without the repeating-header flag.

## Editing an existing docx

`scripts/edit.py` is a **separate chain** from rendering. The two have opposite contracts:

| | Rendering chain | Editing chain |
|---|---|---|
| Input | Markdown | an existing docx |
| Contract | layout only, never content | **change only what is asked; leave every other byte alone** |
| Never | rewrite text | inject a cover, TOC, page numbers, or re-typeset styles |

```bash
python scripts/edit.py 标书.docx --list                        # structure: sections / tables / paragraphs
python scripts/edit.py 标书.docx --replace "旧=新" [--replace "旧2=新2"]
python scripts/edit.py 标书.docx --replace "A=B" --scope body,tables,header,footer
python scripts/edit.py 标书.docx --after  "锚点文字" --text "新段落"       # \n = another paragraph
python scripts/edit.py 标书.docx --before "锚点文字" --text "新段落"
python scripts/edit.py 标书.docx --delete "段落所含文字"
python scripts/edit.py 标书.docx --cell 0 2 1 "1,060,000"        # table row column value
python scripts/edit.py 标书.docx --add-row 0 "接入层" "设备" "320,000"
python scripts/edit.py 标书.docx --del-row 0 2
python scripts/edit.py 标书.docx --header "新版页眉"              # default: body section only
python scripts/edit.py 标书.docx --footer "— X —" --section all
python scripts/edit.py 标书.docx --replace "A=B" --verify        # open in Word afterwards
```

Backups go to `<name>.bak.docx` by default (`--no-backup` disables, `--out` writes elsewhere).

Editing in place is safe because `python-docx` round-trips a docx without losing any part — opening and
saving a 4-page document loses and adds zero package parts (verified).

Two rules that matter:

1. **Word splits text into runs unpredictably.** `自开标之日起 90 日历天` can be three runs with the number
   on its own. `edit.py` matches against the concatenated paragraph text and writes the replacement into
   the run where the match starts, so run count and per-run formatting survive.
2. **An anchor matching more than one paragraph is refused.** After Word refreshes a TOC field, a heading
   such as `1.2 资质与业绩` exists both in the TOC and in the body; inserting after it twice would put
   content into the table of contents. Use a more precise anchor, or pass `--all-anchors`.

Not supported, deliberately: tracked changes, comments, watermarks, content controls, and `.docm` files
with macros (`python-docx` drops `vbaProject.bin` on save).

### Converting an external docx instead of editing it

To **re-typeset** someone else's content into this project's format, convert it and render it back:

```bash
pandoc 甲方文档.docx -t markdown --wrap=none --extract-media=media -o 01_内容.md
python scripts/render.py --src . --out 新版.docx --config cfg.json --pdf --check
```

The TOC field, headers, footers, page numbers, comments and tracked changes are dropped, and `post.py`
injects a fresh cover and TOC — so **clean the intermediate Markdown by hand** (delete the old cover lines
and TOC block, strip `**` from table headers), or the result gets a duplicated cover and TOC and leaks
`[1](#...)` link syntax into the body. Image references from `--extract-media` are relative to the working
directory, so run from the repository root or put `media/` inside `--src`.

## Guidelines

1. **Hard contract: layout only, never content.** `post.py` does not touch **any** text in the body or the
   captions. Every character in the output comes from the source Markdown. Do not add features that rewrite
   text: an `auto_number` feature once existed, silently corrupted body text (`**表层…**` became
   `表 3-5 层…`) and added numbers to captions that had none; it was removed entirely.
   `tests/test_postprocess.py::test_never_touches_text` guards this.
   If automatic numbering is ever needed, use Word's native `SEQ` / `REF` fields — not caption rewriting.
2. **Cover lines carry no leading or trailing blanks.** `post.py` appends one blank `CoverInfo` paragraph
   itself; extra blanks overflow the cover onto a second page.
3. **Figure and table numbers are hand-written.** After inserting or deleting a figure, renumber manually —
   the tool will not do it (see the contract above).
4. **Body and heading fonts live in the template**, not in config. Change them with
   `python scripts/make_ref.py --body-font 楷体 --body-size 14`; the `style` fonts cover tables and
   captions only.
5. **Grid tables are whitespace-sensitive** — the pipe characters must line up exactly.
6. **`--check` also flags legitimately sparse pages** (a signature block at the end of a form, a heading
   alone before a large table). Those are false positives; relax with `--max-empty N`.
7. **Snapshot baselines are machine-specific** (local Word version and fonts). After switching machines,
   re-record with `snapshot.py --update`.
8. **When the client mandates a template, theirs wins** — point `reference_doc` at their docx. Cover sheets,
   sealing, signature pages and page-number rules still need manual compliance review.
9. **After editing an existing docx, run `--verify`** (or at least open the result in Word). Editing writes
   OOXML directly; the only reliable proof the structure survived is that Word still opens the file.

## Acceptance gate: all four must pass

After any real render, check each of these:

1. The log shows `images: n/m ok` with **n == m** (m = occurrences of `![` in the source)
2. `near-empty pages: 0`
3. `OK` (Word opened the file and exported a PDF)
4. **The first time a new document is rendered, look through the rendered page images.** A machine can count
   pages, images and blank pages; it cannot tell whether a diagram is correct or a table header got clipped.

## Layout

| Path | Role |
|---|---|
| `scripts/render.py` | Entry point: merge md → pandoc → post.py → (optional) finalize / check |
| `scripts/post.py` | Cover, TOC field, per-section page numbers, headers/footers, table rules, caption styling |
| `scripts/finalize.py` | Word COM: open for acceptance, refresh TOC field, export PDF |
| `scripts/check_pdf.py` | PyMuPDF: blank-page detection + rendered page PNGs |
| `scripts/snapshot.py` | PDF layout snapshot regression against `baselines/` |
| `scripts/make_ref.py` | Rebuild `assets/ref.docx` (fonts / sizes / spacing) |
| `scripts/edit.py` | Edit an existing docx: replace / insert / delete / table cells / headers and footers |
| `scripts/filters/captions.lua` | pandoc Lua filter: marks captions at the AST level, so `post.py` never guesses with regexes |
| `assets/ref.docx` | Chinese typesetting template (SimSun body, SimHei headings, cover and callout styles) |
| `assets/sample.md`, `assets/sample_config.json` | Smoke-test sample |
| `examples/` | Runnable examples: form-style document, table styling |

Every `config.json` field and `style` key is listed in `README.md`; this file intentionally does not
duplicate those tables.
