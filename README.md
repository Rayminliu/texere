English | [简体中文](README.zh-CN.md)

# texere — verified Chinese Word documents, from Markdown

Markdown + a Word template → DOCX → a renderer (Word / WPS / LibreOffice) → PDF → visual regression → evidence.

```text
Markdown  +  ref.docx / profile.json        the design contract
        │
        ▼
     texere                                 pandoc → deterministic OOXML post-processing
        │
        ▼
     DOCX ──► renderer (Word / WPS / LibreOffice) ──► PDF
        │
        ▼
   9 checks · per-page visual regression · evidence package
```

Generate Chinese tenders, official reports, grant applications and other formal documents **without trusting
DOCX structure alone** — the result is opened in the chosen renderer (real Word by default), exported to PDF, and compared page by page
against a layout baseline.

### See the result

**You write this** (`examples/tender/01_bid.md`):

```markdown
表 1-1 商务条款响应表

| 条款 | 招标要求 | 投标响应 |
|:---|:---|:---|
| 工期 | 90 日历天 | 完全响应 |
| 质量要求 | 符合国家验收标准 | 完全响应 |
| 付款方式 | 按招标文件 | 完全响应 |
```

**The client opens this** — same words, with the template's borders, header-row shading, a centred
caption, section page numbers and a TOC field applied:

![tender](assets/previews/tender.png)

The same pipeline on other document types — every directory in [`examples/`](examples/README.md)
ships its Markdown + config and runs with one command; regenerate these with
`python scripts/make_previews.py`.

| Official document (`gongwen/`) | Application form (`form/`) |
|---|---|
| ![gongwen](assets/previews/gongwen.png) | ![form](assets/previews/form.png) |
| Meeting minutes (`minutes/`) — cover carries the meeting metadata | Business analysis report (`report/`) — multi-file merge, Lead callout, grid header |
| ![minutes](assets/previews/minutes.png) | ![report](assets/previews/report.png) |
| Contract (`contract/`) — clause sections, in-cell line breaks, signature block | Table styling (`tables/`) — multi-level headers, merged cells, column widths |
| ![contract](assets/previews/contract.png) | ![tables](assets/previews/tables.png) |

> **Platform**: DOCX generation is cross-platform. PDF / renderer acceptance needs a renderer —
> **Word** (Windows + Microsoft Word), **LibreOffice** (soffice, any OS), or **WPS** (Windows + WPS Office).
> Pick one with `--renderer word|libreoffice|wps` (default `word`). On a machine with no renderer installed,
> only the docx half runs (see [Known limitations](#known-limitations)).

**Contents** · [Quick start](#quick-start) · [Examples](#see-the-result) ·
[Workflows](#usage) · [Validation](#validation-and-evidence-package) ·
[Editing](#editing-an-existing-docx) · [Docs](#documentation-map)

## What it does and doesn't do

**Does**

- Markdown → docx / PDF, with cover page, TOC field, per-section page numbers and headers
- Table typesetting — borders, header rows, repeating headers, zebra stripes
- Caption conventions for tables and figures
- One-command pre-delivery acceptance: does Word open it, are there blank pages, did the layout drift
- **Targeted edits to an existing docx** — text, paragraphs, table cells, headers/footers — without
  re-typesetting the rest

**Doesn't**

- Comments, tracked changes, redaction, accessibility work, watermarks on an existing docx —
  use a general-purpose docx skill for those
- Automatic figure numbering — numbers are hand-written (see [Contract](#contract-layout-only-never-content))
- Thesis features: bibliography, equation numbering, odd/even page headers

Finer-grained limits — OOXML features that don't survive round-tripping, `.docm`, Word COM specifics —
are in [Known limitations](#known-limitations).

> **Scope note.** This tool is opinionated about *Chinese* formal documents: A4 paper, 宋体 (SimSun) body text,
> 黑体 (SimHei) headings, 2-character first-line indent, full-width punctuation. It is not a general
> Markdown → docx converter — that is what pandoc is. Enabling the `caption_words` option lets you
> work with non-Chinese caption keywords, but the typographic defaults stay Chinese.

## Quick start

```powershell
python scripts/render.py --doctor     # is my environment ready?
python scripts/render.py --sample     # smoke test -> sample_out.docx/.pdf
python scripts/render.py --src examples/tender --out bid.docx `
    --config examples/tender/config.json --pdf --check
```

That third command renders a real tender document, accepts it in the chosen renderer (Word by default), exports the PDF and checks the
result — nothing to configure first. Read on only if `--doctor` complains.

**If `--doctor` complains** — `pandoc` is required; `--pdf` and the validator need a renderer (default Word; pick LibreOffice/WPS with `--renderer`):

| Dependency | Purpose | Install |
|---|---|---|
| pandoc >= 3.1 | **Required**, md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx, lxml | **Required**, docx I/O and OOXML handling | see below |
| pywin32 | `--pdf` only (Word / WPS acceptance + PDF export; not needed for LibreOffice) | see below |
| PyMuPDF | `--check` only (PDF visual acceptance) | see below |
| Microsoft Word / WPS / LibreOffice | `--pdf` only (pick one via `--renderer`) | system-level; neither pip nor uv can install Word/WPS; LibreOffice via package manager |

```powershell
winget install --id JohnMacFarlane.Pandoc
pip install -r requirements.txt      # plain pip; versions are exported from uv.lock
                                     # (uv users: `uv sync --all-extras`)
```

> **This repository is a collection of scripts, not a pip-installable package** — `pip install .` fails
> (there is no build backend, and the tool resolves `assets/ref.docx` / `scripts/` / `assets/sample.md` by
> relative path from each script's own location).
> Put the folder anywhere and run the scripts in place.

`--doctor` **probes the renderers you can use**: it launches the default Word via COM and checks for `soffice` (LibreOffice) / `KWPS` (WPS) on PATH. It does not read the registry:
a stale `CurVer` key left over from an uninstalled Office can make that report lie.
Contributor tooling (ruff, pre-commit, the test suite) is in [Development](#development).

## Measured throughput

On Windows with a Microsoft Word renderer, with Word acceptance and PDF export included:

69 pages / 63 tables / 28 images → **19.0 s** end-to-end

272 pages / 252 tables / 112 images → **66.9–69.7 s** end-to-end

Both runs produced a 2.9 MB docx and matched item by item (page count, word count, tables, images,
sampled page pixels), with no orphaned Word processes. Reproducible examples and benchmark details
→ [`examples/`](examples/README.md).

## Why not pandoc alone?

Pandoc generates DOCX. texere additionally:

1. treats a Word template as the presentation contract (`ref.docx` / `profile.json`);
2. applies deterministic OOXML post-processing — cover, TOC field, per-section page numbers, table styling;
3. opens the result in **real Microsoft Word** and exports PDF, because parseable is not the same as acceptable;
4. checks the rendered artifact — page numbering, blank pages, per-page visual regression;
5. produces evidence for the result: `report.json` + screenshots + a checksum manifest.

## Have an existing Word template? Keep it.

Point `reference_doc` at the client-supplied `.docx` and their design stays in charge:

| The template controls | texere supplies |
|---|---|
| page setup and margins | content from Markdown |
| body / heading / table / caption styles | deterministic post-processing |
| header (conditionally), table borders | Word acceptance + PDF export |

Covers and section structure are **not** inherited — texere builds those itself. Full inheritance matrix:
[Reusing an existing template](#reusing-an-existing-template).

## Usage

The five flows you'll actually run; every flag for every script lives in
[`docs/SCRIPT_HELP.md`](docs/SCRIPT_HELP.md).

```bash
# 0. Environment
python scripts/render.py --doctor                                # self-check: pandoc / deps / Word engine
python scripts/render.py --sample                                # smoke test -> sample_out.docx/.pdf

# 1. Render Markdown → DOCX + PDF + visual check (--check implies --pdf)
python scripts/render.py --src chapters/ --out bid.docx --config cfg.json --pdf --check

# 2. Accept before delivery (9 checks → report.json + screenshots + manifest)
python scripts/validate.py bid.docx --out evidence/
python scripts/validate.py bid.docx --profile profiles/formal-cn-v1.json   # + visual drift vs baseline

# 3. Edit an existing docx (targeted; the editing chain, opposite contract)
python scripts/edit.py bid.docx --replace "工期=进度" --verify
python scripts/edit.py bid.docx --fill data.json                 # batch table filling

# 4. Declarative Patch (Agent-friendly: dry-run, hash preconditions)
python scripts/patch.py bid.docx patch.json --dry-run
python scripts/patch.py bid.docx patch.json --apply --validate

# Supporting cast
python scripts/distill.py 甲方模板.docx --out cfg.json           # template → suggested config
python scripts/make_ref.py --body-font 楷体 --body-size 14       # rebuild the typesetting template
python scripts/snapshot.py bid.pdf --update                      # record baseline (after confirming layout)
python scripts/snapshot.py bid.pdf                               # regression compare; exit 1 on drift
python -m pytest -q                                              # 259 assertions, ~7 min (measured 7m6s, needs a renderer — defaults to Word)
```

Runnable examples — each directory ships its Markdown + config and runs with one command
(see [examples/README.md](examples/README.md)):

```bash
python scripts/render.py --src examples/tables --out examples/tables/tables.docx \
       --config examples/tables/config.json --pdf --check
```

### Input requirements

- Put `01_xxx.md … 0N_xxx.md` in the source directory; they are merged **in filename order**.
  `#` is a chapter, `##` a section. A single `.md` file works too (`--src notice.md`) — no need to
  create a directory for a one-pager.
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

`post.py` never touches body or caption text — figure and table numbers are hand-written in the source.
That guarantee is enforced by `tests/test_postprocess.py::test_never_touches_text`. The render pipeline
*may* transform text, but only via an explicit `content_fixes` table you opt into; with it empty, every
output character comes from the source. Full contract, the `content_fixes` mechanism and the removed
auto-numbering feature → [`SKILL.md`](SKILL.md).

## Validation and evidence package

```bash
python scripts/validate.py bid.docx --out evidence/
```

Produces `report.json`, sampled page screenshots and a checksum manifest. Nine checks run against a
**single** shared Word export. Every check reports `PASS` / `FAIL` / `SKIP` / `ERROR`: `SKIP` means a
precondition was missing so the check never ran, and it is **not** counted as passed — only `FAIL` and
`ERROR` set exit code 1. Check-by-check table, the manual gate and caveats:
[`docs/VALIDATION.md`](docs/VALIDATION.md).

## Editing an existing docx

A **separate chain** from rendering, with the opposite contract: **change only what is asked; leave every
other byte alone** — never inject a cover, TOC, page numbers or re-typeset styles.

```bash
python scripts/edit.py bid.docx --replace "old=new" --scope body,tables
python scripts/edit.py bid.docx --after "anchor text" --text "new paragraph"
python scripts/edit.py bid.docx --cell 0 2 1 "1,060,000"
python scripts/patch.py bid.docx patch.json --dry-run --apply   # declarative Patch, agent-friendly
```

Operations, the Patch schema and re-typesetting: [`docs/EDITING.md`](docs/EDITING.md).

## Reusing an existing template

Point `reference_doc` at any docx — a client-mandated bid template, or one of your own — and **more is
inherited than the previous README claimed**. Measured with a template carrying 3 cm margins, 楷体 14pt
body text and a custom header:

| In the template | Inherited? | Condition |
|---|---|---|
| Page setup (paper size, orientation, margins) | ✅ | automatic |
| Body / heading / table / caption styles (font, size, colour, spacing) | ✅ | automatic |
| **Header** | ✅ | **as long as you do not set `header` in config** |
| Footer / page numbers | ⚠️ | always rewritten — see below |
| Cover | ❌ | `post.py` injects its own |
| Section structure | ❌ | rebuilt as "cover + TOC" / "body" |

```json
{
  "reference_doc": "甲方模板.docx",
  "style": { "page_number": null }
}
```

**Footers**: by default the page number is *appended* to whatever the template already has
(`文件编号 XYZ-2026— 1 —`). To keep the template footer exactly as it is, set `"page_number": null`.

**Inspect a template before using it** — distill it:

```bash
python scripts/distill.py 甲方模板.docx                  # report + suggested config
python scripts/distill.py 甲方模板.docx --out cfg.json    # also write the config
```

The report lists page setup, body/heading fonts, per-section headers and footers, and prints the matching
`make_ref.py` command if you would rather rebuild your own template than reference theirs. Cover structure,
page-number format and table styling cannot be inferred — they are listed as items to confirm by hand.

> **Distillation, not conversion.** The script does not try to "understand" a template. It splits the
> template into config fields and lets you decide — the same division of labour as the rest of the tool.

## Configuration and tables

Reference material lives in the docs, so this manual can stay a product page. Start with the defaults —
they are what `render.py --sample` and every example uses — and open the reference when you need a key:

| Need | Where |
|---|---|
| Every `config.json` field, every `style` key | [`docs/CONFIG.md`](docs/CONFIG.md) |
| Table syntax, grid tables, visual control | [`docs/TABLES.md`](docs/TABLES.md) |

## Tender/bid notes

1. **When the client mandates a template, theirs wins** — see
   [Reusing an existing template](#reusing-an-existing-template) for what gets inherited. What that section
   can't cover for you: cover sheets, sealing, signature pages and page-number rules still need a manual pass
   against the tender document. This tool does not replace compliance review.
2. The usual bid structure (bid letter / commercial / technical / pricing / credentials) maps directly onto
   `#` chapters. Keep price tables as pipe tables so they are easy to swap for the client's own table later;
   use grid tables for complex headers.
3. Before delivery, run `--pdf --check` and look through the page images: Word opens it, no blank pages,
   header rows shaded, captions centred. Only then package it.

## Three pillars

1. **Compiler, not converter** — Markdown → Semantic IR → Layout Spec → DOCX. Every visual rule is explicit
   in a design contract (`ref.docx` + `profile.json`), not magic.
2. **Contract, not guesswork** — The profile declares fonts, spacing, borders, headers, footers, captions.
   Agents can read it; humans can audit it. No hidden assumptions.
3. **Evidence, not hope** — One command validates everything: package integrity, image embedding, TOC fields,
   page numbering, blank pages, Word acceptance, visual drift. Output: `report.json` + screenshots + a checksum
   manifest (see [Evidence](#validation-and-evidence-package)).

## Design principles

Follow these when editing the template or hand-writing content:

1. Template-driven: visual rules live only in `assets/ref.docx`; the source text carries semantics.
2. The Chinese quartet: 宋体 body + 黑体 headings + `eastAsia` font attributes + 2-character first-line indent.
3. The table triplet: 100 % width + shaded bold centred header + repeating header row.
4. Captions: centred, one size smaller, grey, **not italic**.
5. CJK/Latin mixing: 宋体 for Chinese, Times New Roman for digits and Latin.
6. Before delivery: open in real Word and look at the rendered pages. Never trust "it read back fine".

## Known limitations

- The TOC is a Word field; if it is not refreshed on first open, select all and press F9 (the document sets
  `updateFields`, so it usually refreshes itself).
- Figure and table numbers are **hand-written**: after inserting or deleting a figure you must renumber
  manually (see [Contract](#contract-layout-only-never-content)).
- **Body and heading font/size live in the template**; change them with
  `python scripts/make_ref.py --body-font 楷体 --body-size 14`. The fonts in the `style` section cover tables and
  captions only.
- **Grid tables are whitespace-sensitive**: the pipe characters must line up exactly, or parsing goes wrong
  (we have hit a stray trailing `|`). Note that pandoc aligns columns by **display width** — a CJK
  character counts as two columns, so "looks aligned in a monospace editor" can still produce a broken
  (single-column) table. Verify by rendering, or align programmatically.
- **Renderer-bound, not Word-bound**: see the Platform note at the top. Parsing and typesetting don't need
  any Office; the PDF export and half of the acceptance chain run through whichever renderer you pick
  (`--renderer word|libreoffice|wps`). `LibreOffice` is the portable option; `Word`/`WPS` are the
  Chinese-Office fidelity options.
- `--check` also flags pages that are legitimately sparse: the signature block at the end of a form, a
  heading alone before a large table. Those are false positives — relax with `--max-empty N`.
- Snapshot baselines depend on the local Word version and fonts. **After switching machines, re-record with
  `snapshot.py --update`**, or you will see drift everywhere. The threshold is 0.1 % (measured: exporting the
  same document twice gives 0.00 %, changing one header line gives 0.16 %).
- Quarto is not used: its 1.10.x docx output drops table bodies when captions are auto-numbered.

## Development

```bash
pip install ruff pre-commit && pre-commit install      # ruff replaces flake8 + black + isort
pre-commit run --all-files                             # lint + format + a Word-free test subset
python -m pytest -q                                    # full suite: 259 assertions, ~7 min (measured 7m6s)
```

> **There is no hosted CI.** The acceptance tests drive a real Microsoft Word over COM, which no hosted
> runner provides — pre-commit covers the Word-free subset locally; run the full suite before pushing.

## Documentation map

Every piece of information lives in exactly **one** place; the other files link instead of duplicating
(the parts that can be machine-checked are enforced by `tests/test_docs_sync.py`):

| File | Owns |
|---|---|
| `README.md` (+ `.zh-CN` mirror) | The product page: what it is, quickstart, workflows, limitations — not a reference |
| `SKILL.md` | Agent entry: when to use / not, hard contract, acceptance gate, Patch schema, pitfalls |
| `BEST_PRACTICES.md` | Scenario experience: profile choice, styling recipes, debugging cases, FAQ |
| `docs/SCRIPT_HELP.md` | Per-script CLI reference (**the** single source for every option; guarded against the code by `tests/test_docs_sync.py`) |
| `docs/CONFIG.md` (+ `.zh-CN` mirror) | Every `config.json` field and every `style` key |
| `docs/TABLES.md` (+ `.zh-CN` mirror) | Table syntax, grid tables, visual control |
| `docs/VALIDATION.md` (+ `.zh-CN` mirror) | The 9 checks one by one, the manual gate, caveats |
| `docs/EDITING.md` (+ `.zh-CN` mirror) | Editing operations, Patch schema, re-typesetting |

So: **this manual never lists a CLI flag table** — it shows the commands you'll actually run and points at
SCRIPT_HELP for the rest. `test_docs_sync.py` fails the commit if a script grows an option that isn't
documented there, or if SCRIPT_HELP invents one that doesn't exist.

**Repository layout**

| File | Role |
|---|---|
| `scripts/render.py` | Entry point: merge md → pandoc → post.py → (optional) finalize / check |
| `scripts/post.py` | Post-processing: cover injection, TOC field, per-section page numbers, headers/footers, table rules, caption styling; hand-written OOXML inserted in ECMA-376 order |
| `scripts/filters/captions.lua` | pandoc Lua filter: marks table/figure captions as `TableCaption` / `FigureCaption` **at the AST level**, so `post.py` never has to guess with regexes |
| `scripts/finalize.py` | Word COM: open for acceptance (failure to open = structural error), refresh TOC field, export PDF, save back |
| `scripts/check_pdf.py` | PyMuPDF: blank-page detection (exits 1 over threshold) + renders page PNGs for review |
| `scripts/validate.py` | Unified validation entry — 9 automated checks, structured report, evidence package |
| `scripts/snapshot.py` | PDF layout snapshot regression: pixel comparison against `baselines/`, exits 1 on drift |
| `scripts/make_previews.py` | Regenerate the examples gallery images (renders each example, picks a representative page) |
| `scripts/patch.py` | Agent-friendly Patch API — declarative operations, dry-run, hash precondition, assessment |
| `scripts/edit.py` | Edit an existing docx: replace / insert / delete / table cells / headers and footers |
| `scripts/distill.py` | Distill a template docx into a suggested `config.json` (page setup, fonts, headers, footers) |
| `scripts/make_ref.py` | Rebuild `assets/ref.docx` (use when changing fonts / sizes / spacing) |
| `profiles/*.json` | Design contracts: `formal-cn-v1` (general formal), `gongwen-v1`, `tender-v1`, `application-v1` |
| `assets/ref.docx` | The Chinese typesetting template: 宋体 body, 黑体 heading ladder, table borders, caption styles, cover styles (CoverTop/…), callout styles (Lead/SmallNote) |
| `assets/sample.md` / `assets/sample_config.json` | Pinned regression fixture, not a showcase: `--sample`, four test files and `baselines/` all render this one document. Showcases live in `examples/` |
| `examples/` | Runnable examples: tender, official document (gongwen), application form, meeting minutes, business analysis report, contract, table styling (see `examples/README.md`) |
| `docs/` | `SCRIPT_HELP.md` (CLI), `CONFIG.md` (config fields), `TABLES.md` (tables), `VALIDATION.md` (9 checks), `EDITING.md` (editing + Patch) — each with a `.zh-CN` mirror |
| `baselines/` | Snapshot baselines (4 PNG pages of the sample) |
| `tests/` | 259 pytest assertions: layout rules, caption recognition, table features, exit codes, snapshot logic, the layout-only contract, cross-run editing (`test_edit.py`), template reuse and distillation (`test_distill.py`), the 9-check validator (`test_validate.py`), the Patch API (`test_patch.py`), version consistency and page-number/baseline pure functions (`test_version.py` / `test_validate_units.py`), Renderer abstraction guard (`test_renderer.py`: adapter contract + PDF-derived checks are renderer-agnostic; `test_renderer_contract.py`: 抽象契约 + 超时/并发集成), docs-vs-code sync guard (`test_docs_sync.py`: CLI options ↔ SCRIPT_HELP both ways, single-source key tables, EN/ZH mirror structure, internal anchors, SKILL.md front-matter YAML) |
| `CHANGELOG.md` | Version history and the reasoning behind each fix |

## License

MIT — see [LICENSE](LICENSE).
