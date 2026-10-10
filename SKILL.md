---
name: texere
description: "The Chinese formal document compiler — Markdown → typeset docx/PDF with evidence. Renders Chinese formal documents (tenders, bids, grant applications, final reports, official notices) from Markdown with explicit design contracts, then validates with 9 automated checks (package integrity, source content, image embedding, TOC fields, page numbering, blank pages, renderer acceptance, visual drift) producing report.json + screenshots + signature. Also does targeted edits to an existing docx without re-typesetting. Not for tracked changes, comments, watermarks, or theses. English documents: a neutral layout baseline exists (profiles/neutral-en-v1.json); full English conventions are out of scope."
license: MIT
compatibility: pandoc 3.1+, Python 3.10+, python-docx and lxml; PDF export and renderer acceptance need a local renderer (Word / WPS on Windows, LibreOffice elsewhere — select with `--renderer`). DOCX generation is cross-platform.
---

# texere

**Compiler, not converter** — every visual rule is explicit in `ref.docx` + config. **Contract, not
guesswork** — the pipeline never rewrites body/caption text (unless you configure `content_fixes`,
an explicit replacement table you wrote). **Evidence, not hope** — one command produces a 9-check
report where SKIP is never counted as PASS.

> Where details live (do not restate them here — `tests/test_docs_sync.py` fails the commit):
> every config key → `docs/CONFIG.md`; CLI options → `docs/SCRIPT_HELP.md`; table syntax →
> `docs/TABLES.md`; check-by-check validation → `docs/VALIDATION.md`; scenario recipes →
> `BEST_PRACTICES.md`.

## Quick start

```bash
python scripts/render.py --doctor     # environment self-check: pandoc / deps / renderer
python scripts/render.py --sample     # smoke test -> sample_out.docx/.pdf

python scripts/render.py --src chapters/ --out bid.docx --config cfg.json --pdf --check
#    --src takes a directory (merged in filename order) or a single .md

python scripts/validate.py bid.docx --out evidence/          # 9-check gate + evidence package
python scripts/patch.py bid.docx patch.json --dry-run        # then --apply --validate
python scripts/edit.py bid.docx --replace "示例科技=某某科技" --verify   # targeted edit
```

Run from the repository root. Dependencies: `pip install -r requirements.txt`; pandoc and a
renderer are system-level (`--doctor` verifies both).

## When to use / when not to

Use for: Chinese formal documents from Markdown (tenders, bids, grants, reports, notices); one-command
pre-delivery acceptance with evidence; declarative patch editing with dry-run and hash preconditions;
targeted edits to an existing docx without re-typesetting.

Do **not** use for: comments / tracked changes / redaction / watermarks / content controls on an
existing docx (out of scope — use a general docx skill); bulk content filling (the edit chain is for
point edits; `--fill` covers coordinate-based batch filling only); theses or book manuscripts (no
bibliography, equation numbering, or odd/even headers); full English business conventions (the
`neutral-en-v1` profile is a layout baseline, not convention coverage).

## Hard contract and pitfalls

1. **The pipeline never rewrites content** — layout only. `content_fixes` / `content_fixes_file` are
   the sanctioned exception (explicit replacement table, applied at merge time). An `auto_number`
   feature once silently corrupted body text and was removed — do not reintroduce caption rewriting;
   use Word `SEQ`/`REF` fields. Guard: `test_postprocess.py::test_never_touches_text`.
2. **`--out` must not be a source path** — render refuses (it would overwrite the .md with binary).
3. **Raw HTML passes through unfiltered** — the pandoc reader enables `raw_html`; review Markdown
   from untrusted sources before rendering.
4. **Cover lines take no leading/trailing blanks** — `post.py` appends one blank `CoverInfo`
   paragraph itself; extra blanks overflow the cover onto a second page.
5. **No body text before the first `#`** — a YAML `title`/`author` block is kept and centered (no
   cover); plain paragraphs land after the TOC with a `[warn]` (never deleted — contract #1).
6. **Fonts/sizes/margins**: rebuild the global template with `make_ref.py`, or override per document
   via config `style` (`title_size`, `body_size`, `h1_size`–`h3_size`, `margin_*` in cm);
   `east_font`/`latin_font` cover tables and captions only.
7. **Grid tables are parsed by display width** (CJK = 2 columns) — render aligns them automatically
   at merge time (source untouched); `align_tables.py --fix` tidies the source itself.
8. **`--check` sparse pages**: a short final page is only a WARN (natural ending); other sparse pages
   FAIL — tune with `--max-empty` / `--near-empty-chars` / `--strict-last`.
9. **Snapshot baselines are machine-specific** — re-record with `snapshot.py --update` after
   switching machines.
10. **A client-mandated template wins** — point `reference_doc` at their docx; sealing, signature
    pages and page-number rules still need human compliance review.
11. **After editing an existing docx, always `--verify`** (or open in Word) — the only proof the
    OOXML survived is that Word still opens it.

## Minimal config shape

```json
{
  "cover": [["CoverTitle", "投 标 文 件"], ["CoverInfo", "投标人：示例科技有限公司"]],
  "header": "示例项目 · 投标文件",
  "toc": true,
  "style": {"table_zebra": true}
}
```

Three starting points: formal documents write nothing; short documents (notices, homework) use
`{"mode": "simple-report"}` (no TOC, no page break before H1, title/author kept centered); forms
(first table row is field names) use `{"style": {"header_rows": 0}}`.

CJK text: ASCII double quotes are paired into full-width `“”` at merge time and pandoc `smart` is
disabled (its heuristic assumes space-separated Western text and mispairs every quote in Chinese).
Every field and `style` key: `docs/CONFIG.md`.

## Two chains, opposite contracts

| | Rendering chain | Editing chain |
|---|---|---|
| Input | Markdown | an existing docx |
| Contract | layout only, never content | change only what is asked; every other byte untouched |

1. **Word splits text into runs unpredictably** — `edit.py` matches concatenated paragraph text and
   writes into the run where the match starts; never hand-edit `run.text`.
2. **An anchor matching more than one paragraph is refused** — after a TOC refresh a heading exists
   twice; use a precise anchor or `--all-anchors`.

Re-typesetting someone else's docx goes through Markdown; see README §Reusing an existing template.

## Acceptance gate

Scale the gate to the document: a short one (notice, homework) needs only the default render plus
one `--pdf --check` before delivery; reserve `validate.py` for formal multi-page deliverables.

```bash
python scripts/validate.py bid.docx --out evidence/ --source-md chapters/01_bid.md
```

Nine checks on one shared export: package integrity, source content (`--source-md` body comparison),
image embedding (per-image sha256), section count, TOC field, page numbering, blank pages, renderer
acceptance, visual drift. Exit 1 on FAIL **or ERROR**. Output: `report.json` + screenshots +
`signature` (a checksum manifest proving *which artifact* the evidence describes — not tamper-proof).

Read the statuses: `SKIP` means "could not run" and is never counted as passed —
`Passed: 7/9 (skipped: 2)` is weaker evidence, not a pass. Details: `docs/VALIDATION.md`.

## Patch schema (Agent API)

```json
{
  "id": "patch-001",
  "description": "工期 90 → 120",
  "preconditions": { "hash": "<sha256>", "must_contain": ["90 日历天"] },
  "operations": [
    { "op": "replace_text", "target": { "paragraph": 42 },
      "expected_old_text": "90 日历天", "new_text": "120 日历天" }
  ]
}
```

Operations: `replace_text` / `insert_after` / `insert_before` / `delete_paragraph` / `set_cell` /
`add_row` / `del_row`. Flow: `--dry-run` → `--apply --validate`. Full schema: `docs/SCRIPT_HELP.md`
§patch.py.
