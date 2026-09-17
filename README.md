English | [简体中文](README.zh-CN.md)

# texere — turn Markdown into properly typeset Chinese documents

*From Latin* texere*, "to weave" — the shared root of* text *and* textile*. Typesetting is the same act:
body text, tables, captions and page numbers woven into an ordered page.*

Renders a Markdown chapter directory into a **production-grade Chinese formal document** (docx + PDF):
tenders and bids, project plans, grant applications, final reports, white papers.

Three things define it:

1. **Template-driven** — every visual rule lives in `ref.docx`. The Markdown source carries semantics only.
   You can restyle the whole document without touching a single word of content.
2. **Real Word acceptance** — it opens the result in your local Word, refreshes the TOC field, exports a PDF,
   and renders pages for eyeballing. It does not trust "it read back fine".
3. **Layout only, never content** — every character in the output comes from your source file.
   This is enforced by a test.

## What it does / doesn't do

| | |
|---|---|
| **Does** | Markdown → docx / PDF; cover page, TOC field, per-section page numbers, headers; table typesetting (borders / header rows / repeating headers / zebra stripes); caption conventions for tables and figures; one-command pre-delivery acceptance (does Word open it, are there blank pages, did the layout drift) |
| **Doesn't** | **Edit existing docx files** (comments, tracked changes, redaction, accessibility, watermarks) — use a general-purpose docx skill for that; automatic figure numbering (numbers are hand-written, see *Contract*); thesis features such as bibliography, equation numbering, odd/even page headers |

> **Scope note.** This tool is opinionated about *Chinese* formal documents: A4 paper, 宋体 (SimSun) body text,
> 黑体 (SimHei) headings, 2-character first-line indent, full-width punctuation. It is not a general
> Markdown → docx converter — that is what pandoc is. Enabling the `caption_words` option lets you
> work with non-Chinese caption keywords, but the typographic defaults stay Chinese.

## Quick start

External dependencies: `pandoc` is required; `--pdf` needs a local Microsoft Word.

| Dependency | Purpose | Install |
|---|---|---|
| pandoc >= 3.1 | **Required**, md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx, lxml | **Required**, docx I/O and OOXML handling | see below |
| pywin32 | `--pdf` only (Word acceptance + PDF export) | see below |
| PyMuPDF | `--check` only (PDF visual acceptance) | see below |
| Microsoft Word | `--pdf` only | system-level; neither pip nor uv can install it |

> **This repository is a collection of scripts, not a pip-installable package** — `pip install .` fails
> (there is no build backend, and the tool locates `ref.docx` / `filters/` / `sample.md` by relative path).
> Put the folder anywhere and run the scripts in place.

```powershell
winget install --id JohnMacFarlane.Pandoc

# Dependencies — pick one:
uv sync --all-extras                 # with uv (recommended; pinned in uv.lock)
pip install -r requirements.txt      # plain pip (exact versions exported by uv)

python render.py --doctor            # environment self-check
python render.py --sample            # smoke test -> sample_out.docx/.pdf
```

`--doctor` **identifies the Word engine by actually launching it via COM**. It does not read the registry:
a stale `CurVer` key left over from an uninstalled Office can make that report lie.

## Result

Output of `python render.py --sample` (a 4-page tender-style sample):

![sample cover](baselines/p001.png)
![sample body page](baselines/p003.png)

**Measured throughput** (Word acceptance + PDF export + blank-page check included):

| Size | Time |
|---|---|
| 69 pages / 63 tables / 28 images | **19.0 s** |
| 272 pages / 252 tables / 112 images | **69.7 / 66.9 s** (two consecutive runs) |

The large run produced a 2.9 MB docx; both runs matched item by item (page count, word count, tables, images,
and sampled page pixels), with no orphaned Word processes.

## Usage

```bash
python render.py --doctor                                    # environment self-check
python render.py --version                                   # version
python render.py --sample                                    # smoke test

# Real render (--check implies --pdf)
python render.py --src chapters/ --out bid.docx --config cfg.json --pdf --check

python -m pytest -q                        # 36 assertions, no Word needed, ~9 s
python snapshot.py bid.pdf --update        # record layout baseline (after confirming the layout)
python snapshot.py bid.pdf                 # regression compare; exits 1 on drift
python make_ref.py --body-font 楷体        # rebuild the typesetting template
```

### Input requirements

- Put `01_xxx.md … 0N_xxx.md` in the source directory; they are merged **in filename order**.
  `#` is a chapter, `##` a section.
- Table captions are their own line: `表 1-1 Caption` (centred and greyed automatically). Figures:
  `![图 1-1 Caption](a.jpg){width=13cm}`.
- **Image search path**: the `src` directory, all of its subdirectories, and **the parent of `src` plus that
  parent's subdirectories** are added automatically; add more with `resource_paths`. The render prints
  `images: n/m ok`, and reports `[ERROR]` if the source references images that were not embedded.
- Cover and header are configured in the config file; without a `cover` only the TOC is injected.
  **Note**: after injecting cover lines, `post.py` appends one blank `CoverInfo` paragraph itself — so do
  not put leading/trailing blank lines in the config, or the cover overflows onto a second (blank) page.
- Emphasis is `**bold**`; a grey callout box is `::: {custom-style="Lead"} … :::`.
- **Do not put content before the first `#`** (a title-like phrase, say): it is neither treated as a cover
  nor kept in place — it ends up after the TOC. `post.py` prints a `[warn]` listing it, but will not delete
  it for you (the contract forbids changing content).
- config.json and source `.md` files may carry a UTF-8 BOM (Windows Notepad writes one by default).
- **Form-style / appendix documents** (no `#` headings) are supported: no TOC, no section split, just cover
  and table/caption typesetting. Their tables usually start with field names rather than column headers, so
  set `"style": {"header_rows": 0}` to keep the first row from being shaded as a header.

### Contract: layout only, never content

`post.py` does not touch **any** of the text in the body or the captions. Figure and table numbers are
hand-written in the source.

Version 0.2.0 briefly shipped an "automatic figure numbering + `@tab:` cross-references" feature. It was
**removed entirely** because it rewrote caption text, added numbers to captions that had none, and left
hand-written in-text references out of sync. If numbering is ever revisited, the right implementation is
Word's native `SEQ` / `REF` fields (updatable, no text rewriting) — not rewriting caption text.

### Acceptance gate

After a real render, check all four. All four must pass:

1. The log shows `images: n/m ok` with **n == m** (m = number of `![` in the source).
2. `near-empty pages: 0`
3. `OK` (Word opened the file and exported a PDF)
4. **The first time you render a new document, look through the rendered page images yourself.**
   A machine can count pages, images and blank pages; it cannot tell whether a diagram is correct or a
   table header got clipped.

## Configuration

### config.json fields

| Field | Purpose |
|---|---|
| `cover` | Cover lines `[[style_name, text], …]`; style names: CoverTop / CoverTitle / CoverSub / CoverInfo / CoverDate |
| `header` | Header text for the body section |
| `title` / `author` / `subject` / `comments` | Document properties |
| `reference_doc` | Use a specific template docx (for a client-mandated format) |
| `toc_heading` | TOC title, default 「目　　录」 |
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

Everything is optional; the defaults are the Chinese formal-document conventions:

| Group | Keys | Default |
|---|---|---|
| Footer | `page_number` / `page_number_size` | `— {n} —` / `9` |
| Header | `header_size` / `header_gray` / `header_rule_color` / `header_rule_size` | `9` / `595959` / `BFBFBF` / `4` |
| TOC | `toc_depth` / `toc_title_size` / `toc_title_color` / `toc_placeholder` / `toc_placeholder_size` | `1-2` / `16` / `000000` / see source / `12` |
| Captions | `caption_gray` / `caption_size` / `caption_space_before` / `caption_space_after` | `404040` / `10.5` / `6` / `4` |
| Captions | `caption_keep_with_next` | `true` (keeps a table caption with its table; turning it off saves a page in testing, but the caption may strand at a page foot) |
| Tables | `header_rows` / `table_border` / `table_shade` / `table_size` | auto / `full` / `EDEDED` / `10.5` |
| Tables | `cell_margin_v` / `cell_margin_h` / `table_para_space` | `40` / `80` / `1` |
| Tables | `border_size` / `border_color` / `three_line_size` | `6` / `808080` / `12` |
| Tables | `table_header_color` / `table_zebra` / `table_zebra_fill` | unset / `false` / `F7F7F7` |
| Fonts | `east_font` / `latin_font` | `宋体` / `Times New Roman` (**tables and captions only**) |

> Colours are 6-digit hex (`404040`); border weights are in 1/8 pt; spacing is in pt; cell margins are in twips.

## Tables

### Two syntaxes, pick per table

| | pipe table | grid table |
|---|---|---|
| Syntax | `\| a \| b \|` | `+---+---+` |
| Multi-level headers | ❌ | ✅ (everything above `+===+` is header) |
| Merged cells | ❌ | ✅ (column and row spans) |
| Line breaks inside a cell | ❌ | ✅ |
| **Column widths** | ❌ all equal | ✅ see below |
| Column alignment | ✅ `:--` `:-:` `--:` | ❌ |

Use pipe tables for simple lists and price tables; **use grid tables for complex headers, merged cells, and
when you need to control column widths**.

### Grid tables: the three tricks

**1. Multi-level headers + merged cells** — `+===+` separates header from body; omitting an inner `|` on a
row means the cell spans columns:

```
+------------------+------------------+
| 商务部分         | 技术部分         |
+--------+---------+--------+---------+
| 条款   | 响应    | 模块   | 说明   |
+========+=========+========+=========+
| 工期   | 完全响应| 接入层 | 设备   |
+--------+---------+--------+---------+
```

**2. Column widths** — pandoc derives widths from the **narrowest column separator**. The narrowest row above
is `+--------+---------+` (8/9/8/9), and the four columns come out 8:9:8:9 (measured 990/1100/990/1100 twips).
Want a narrow "No." column? Write it narrow.

**3. Line breaks inside a cell** — just write multiple lines in the cell (pipe tables cannot).

### Visual control

| Key | Default | Meaning |
|---|---|---|
| `header_rows` | auto | First N rows get the **visual** header treatment (shaded, bold, centred). By default this is **detected per table** (reading the `w:tblHeader` pandoc emits), so tables in one document can differ; set a number to force it; **`0` means the table has no header** (forms/appendix tables whose first row is a field name) and also drops the repeating-header flag |
| `table_border` | `full` | `full` all rules / `three` three-line (booktabs) / `none` |
| `table_shade` | `EDEDED` | Header fill |
| `table_header_color` | unset | Header text colour; **pair with `FFFFFF` on a dark fill** |
| `table_zebra` / `table_zebra_fill` | `false` / `F7F7F7` | Alternate row shading; the first data row stays white |
| `table_size` | `10.5` | Table font size, pt |

> Repeating headers across pages need no configuration: pandoc already sets `w:tblHeader` on header rows
> (everything above `+===+` in a grid table counts as header). The line in `post.py` is just idempotent hardening.

## Tender/bid notes

1. **When the client mandates a template, theirs wins**: point `reference_doc` at their docx and body styles
   are inherited from it. Cover sheets, sealing, signature pages and page-number rules still need manual
   checking against the tender document — this tool does not replace compliance review.
2. The usual bid structure (bid letter / commercial / technical / pricing / credentials) maps directly onto
   `#` chapters. Keep price tables as pipe tables so they are easy to swap for the client's own table later;
   use grid tables for complex headers.
3. Before delivery, run `--pdf --check` and look through the page images: Word opens it, no blank pages,
   header rows shaded, captions centred. Only then package it.

## Design principles (follow these when editing the template or hand-writing content)

1. Template-driven: visual rules live only in `ref.docx`; the source text carries semantics.
2. The Chinese quartet: 宋体 body + 黑体 headings + `eastAsia` font attributes + 2-character first-line indent.
3. The table triplet: 100 % width + shaded bold centred header + repeating header row.
4. Captions: centred, one size smaller, grey, **not italic**.
5. CJK/Latin mixing: 宋体 for Chinese, Times New Roman for digits and Latin.
6. Before delivery: open in real Word and look at the rendered pages. Never trust "it read back fine".

## Known limitations

- The TOC is a Word field; if it is not refreshed on first open, select all and press F9 (the document sets
  `updateFields`, so it usually refreshes itself).
- Figure and table numbers are **hand-written**: after inserting or deleting a figure you must renumber
  manually (see *Contract*).
- **Body and heading font/size live in the template**; change them with
  `python make_ref.py --body-font 楷体 --body-size 14`. The fonts in the `style` section cover tables and
  captions only.
- **Grid tables are whitespace-sensitive**: the pipe characters must line up exactly, or parsing goes wrong
  (we have hit a stray trailing `|`).
- **Windows + Word bound**: `--pdf` needs a local Word (COM). On Linux/macOS only the docx half works
  (the parsing and typesetting logic does not depend on Word, but half the acceptance chain is missing).
- `--check` also flags pages that are legitimately sparse: the signature block at the end of a form, a
  heading alone before a large table. Those are false positives — relax with `--max-empty N`.
- Snapshot baselines depend on the local Word version and fonts. **After switching machines, re-record with
  `snapshot.py --update`**, or you will see drift everywhere. The threshold is 0.1 % (measured: exporting the
  same document twice gives 0.00 %, changing one header line gives 0.16 %).
- Quarto is not used: its 1.10.x docx output drops table bodies when captions are auto-numbered.

## Repository map

| File | Role |
|---|---|
| `render.py` | Entry point: merge md → pandoc → post.py → (optional) finalize / check |
| `post.py` | Post-processing: cover injection, TOC field, per-section page numbers, headers/footers, table rules, caption styling; hand-written OOXML inserted in ECMA-376 order |
| `filters/captions.lua` | pandoc Lua filter: marks table/figure captions as `TableCaption` / `FigureCaption` **at the AST level**, so `post.py` never has to guess with regexes |
| `ref.docx` | The Chinese typesetting template: 宋体 body, 黑体 heading ladder, table borders, caption styles, cover styles (CoverTop/…), callout styles (Lead/SmallNote) |
| `make_ref.py` | Rebuild `ref.docx` (use when changing fonts / sizes / spacing) |
| `finalize.py` | Word COM: open for acceptance (failure to open = structural error), refresh TOC field, export PDF, save back |
| `check_pdf.py` | PyMuPDF: blank-page detection (exits 1 over threshold) + renders page PNGs for review |
| `snapshot.py` | PDF layout snapshot regression: pixel comparison against `baselines/`, exits 1 on drift |
| `tests/` | 36 pytest assertions: layout rules, caption recognition, table features, exit codes, snapshot logic, the layout-only contract |
| `sample.md` / `sample_config.json` | Smoke-test sample (tender-document style) |
| `baselines/` | Snapshot baselines (4 PNG pages of the sample) |
| `SKILL.md` | Skill description for other agents (when to use, acceptance gate, hard contract) |
| `CHANGELOG.md` | Version history and the reasoning behind each fix |

## License

MIT — see [LICENSE](LICENSE).
