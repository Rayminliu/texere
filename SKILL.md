---
name: docx-kit
description: Render Markdown into properly typeset Chinese formal documents (docx + PDF) — tenders and bids, project plans, grant applications, final reports, white papers — and verify the result by opening it in real Word, refreshing the TOC field, exporting a PDF, and checking for blank pages and layout drift. Use when a task involves generating a formal document from Markdown, adding a cover page / table of contents / header / per-section page numbers, Chinese formal typesetting (SimSun body, SimHei headings, 2-character first-line indent, centred captions), repeating table header rows across pages, or pre-delivery checks for blank pages and layout changes.
---

# docx-kit

A local toolkit that renders a Markdown chapter directory into a Chinese formal docx / PDF. Two ideas carry
the whole design: **template-driven** (every visual rule lives in `ref.docx`; the source text carries
semantics only) and **real-machine acceptance** (open in Word, export a PDF, then look at the rendered pages —
never trust "it read back fine").

## Hard contract: layout only, never content

`post.py` does not touch **any** of the text in the body or the captions. Every character in the output comes
from the source Markdown.

Do not add features to this pipeline that rewrite text. An `auto_number` (automatic figure/table numbering)
feature once existed and silently corrupted body text (`**表层…**` became `表 3-5 层…`) and added numbers to
captions that had none; it was removed entirely. `tests/test_postprocess.py::test_never_touches_text` guards
this contract.

If automatic numbering and cross-references are ever needed, the correct implementation is Word's native
`SEQ` / `REF` fields (updatable, no text rewriting) — not rewriting caption text.

## When to use / when not to

Use it for:

- Generating a Chinese formal document from Markdown (bid, grant application, final report, white paper, proposal)
- Cover page, TOC field, per-section page numbers, header, table/figure caption conventions, repeating table
  header rows
- One-command pre-delivery acceptance (does Word open it, are there blank pages, did the layout drift)

Do **not** use it for:

- **Editing an existing docx**: comments, tracked changes, redaction, accessibility fixes, watermarks, form
  controls. This tool has no editing capability and should not gain one — use a general-purpose docx skill
  instead (for example Codex's Documents plugin, which renders via LibreOffice and ships 36 OOXML scripts).
- Theses or book manuscripts: bibliography (citeproc), equation numbering and odd/even page headers are not
  supported.
- English-language documents: the defaults (A4, SimSun/SimHei, full-width punctuation, Chinese caption
  keywords) are Chinese-document conventions. `caption_words` lets you change the keywords, not the typography.

## Commands

```bash
python render.py --doctor                    # environment self-check: pandoc / deps / Word engine identity
python render.py --version                   # version
python render.py --sample                    # smoke test (produces sample_out.docx/.pdf)

# Real render (--check implies --pdf)
python render.py --src <md dir> --out bid.docx --config cfg.json --pdf --check

python -m pytest -q                          # 36 assertions, no Word needed, ~9 s
python snapshot.py bid.pdf                   # layout snapshot regression; exits 1 on drift
python snapshot.py bid.pdf --update          # re-record the baseline after an intended layout change
python make_ref.py --body-font 楷体          # rebuild the typesetting template (fonts / sizes / spacing)
```

## Input requirements

- `--src` points at a directory; its `*.md` files are merged **in filename order** (use `01_`, `02_` prefixes
  to control the order)
- `#` is a chapter, `##` a section; a table caption is its own line (`表 1-1 Caption`); a figure caption goes
  in the image alt text: `![图 1-1 Caption](a.png){width=13cm}`
- Image search path: the `src` directory, all of its subdirectories, and **the parent of `src` plus that
  parent's subdirectories** (images often live in a sibling directory); add more with `resource_paths`
- **Do not put content before the first `#`** (it is neither treated as a cover nor kept in place — it lands
  after the TOC). `post.py` prints a `[warn]` but will not delete it for you
- config.json and source `.md` files may carry a UTF-8 BOM

## Acceptance gate: all four must pass

After any real render, check each of these:

1. The log shows `images: n/m ok` with **n == m** (m = occurrences of `![` in the source)
2. `near-empty pages: 0`
3. `OK` (Word opened the file and exported a PDF)
4. **The first time you render a new document, look through the rendered page images yourself.** A machine
   can count pages, images and blank pages; it cannot tell whether a diagram is correct or a table header
   got clipped.

`--check` also reports FAIL for pages that are legitimately sparse (the signature block at the end of a form,
a heading alone before a large table). Those are false positives — relax with `--max-empty N`.

## Configuration

Every field is listed in `README.md` (see the config field tables there). The ones used most:

| Key | Meaning |
|---|---|
| `cover` | Cover lines `[[style_name, text], …]`; style names CoverTop/CoverTitle/CoverSub/CoverInfo/CoverDate. **Do not add leading/trailing blank lines** — `post.py` appends one blank `CoverInfo` paragraph itself, and extras overflow the cover onto a second page |
| `header` | Header text for the body section |
| `reference_doc` | Use a client-supplied docx as the template |
| `content_fixes_file` | Editorial replacement table; when it points at a `.py`, values are read with `ast` and never executed |
| `resource_paths` | Extra directories to search for images |
| `style` | Fine-grained layout (page-number template, TOC depth, caption colour/size, table borders, `header_rows: 0` for a table with no header, …) |

## Known limitations

- **Windows + Word bound**: `--pdf` needs a local Word (COM). On Linux/macOS only the docx half works
- **Snapshot baselines are machine-specific**: after switching machines or Word versions, re-record with
  `snapshot.py --update`
- Figure and table numbers are **hand-written**: after inserting or deleting a figure you renumber manually
  (the tool will not do it — see the hard contract)
- Chinese full-width punctuation and Chinese-style full-grid tables (or `table_border: three` for a
  three-line table) differ from English-document conventions (Letter paper, Latin fonts). This tool targets
  Chinese formal documents.

## Repository map

| File | Role |
|---|---|
| `render.py` | Entry point: merge md → pandoc → post.py → (optional) finalize / check |
| `post.py` | Post-processing: cover, TOC field, per-section page numbers, headers/footers, table rules, caption styling |
| `filters/captions.lua` | pandoc Lua filter: marks table/figure captions at the **AST level**, so `post.py` never guesses with regexes |
| `make_ref.py` | Rebuild `ref.docx` (fonts / sizes / spacing) |
| `finalize.py` | Word COM: open for acceptance + refresh TOC field + export PDF |
| `check_pdf.py` | PyMuPDF: blank-page detection (with exit code) + render page PNGs |
| `snapshot.py` | PDF layout snapshot regression: pixel comparison against `baselines/` |
| `tests/` | 36 pytest assertions |
| `README.md` | English README (full field tables, typesetting rules, tender notes) |
| `README.zh-CN.md` | 中文版 README |
| `CHANGELOG.md` | Version history and the reasoning behind each fix |

## Performance

Measured (Word acceptance + PDF export + blank-page check included):

| Size | Time |
|---|---|
| 69 pages / 63 tables / 28 images | **19.0 s** |
| 272 pages / 252 tables / 112 images | **69.7 / 66.9 s** (two consecutive runs; page count, word count, tables, images and sampled page pixels all matched, no orphaned Word processes) |
