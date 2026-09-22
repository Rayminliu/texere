---
name: texere
description: "The Chinese formal document compiler — Reference/Spec → Deterministic Document → Evidence. Renders Markdown into properly typeset Chinese formal documents (docx + PDF) with unified validation: 9 automated checks (package integrity, image embedding, TOC fields, page numbering, blank pages, Word acceptance, visual drift), structured report (report.json), and evidence package (screenshots + signature). Use when generating tenders, bids, grant applications, final reports, white papers from Markdown with explicit design contracts (profile.json or ref.docx). Also supports targeted edits to existing docx without re-typesetting. Should not be used for tracked changes, comments, watermarks, theses, or English documents."
license: MIT
compatibility: Requires Windows with a local Microsoft Word (COM), pandoc 3.1 or newer, Python 3.10 or newer, python-docx and lxml. On Linux and macOS only the docx half works — the Word acceptance and PDF export chain is unavailable.
---

# texere

**Reference/Spec → Deterministic Document → Evidence**

Three pillars: **compiler, not converter** (every visual rule is explicit in `ref.docx` + config);
**contract, not guesswork** (post-processing is layout-only — it never rewrites body or caption text;
the pipeline stays faithful *unless* `content_fixes` is configured, and that is an explicit
replacement table you wrote, not a silent rewrite);
**evidence, not hope** (one command produces a 9-check report + screenshots + signature, where every
check reports PASS / FAIL / SKIP / ERROR and SKIP is never counted as PASS).

> **Documentation map**: this file is the agent entry point — decisions, contracts, pitfalls, commands.
> Field-level detail (every config key, `style` table, table syntax, known limitations) lives **only**
> in `README.md` / `README.zh-CN.md`; per-script CLI options live **only** in `docs/SCRIPT_HELP.md`.
> Do not restate those tables here — `tests/test_docs_sync.py` fails the commit if you do.

## Quick start

```bash
python scripts/render.py --doctor     # environment self-check: pandoc / deps / Word engine identity
python scripts/render.py --sample     # smoke test -> sample_out.docx/.pdf

# Render Markdown → DOCX + PDF（--src 可以是目录，也可以是单个 .md 文件）
python scripts/render.py --src chapters/ --out bid.docx --config cfg.json --pdf --check

# Validate (Compiler + Contract + Evidence)
python scripts/validate.py bid.docx --out evidence/

# Edit with declarative Patch (Agent-friendly)
python scripts/patch.py bid.docx patch.json --dry-run
python scripts/patch.py bid.docx patch.json --apply --validate

# Edit an existing docx (targeted, opposite contract — see "Two chains")
python scripts/edit.py bid.docx --replace "示例科技=某某科技" --verify

python -m pytest -q                   # 213 assertions (~6-7 min; validator tests need local Word; no hosted CI)
python scripts/snapshot.py bid.pdf    # layout regression; exit 1 on drift
```

Run from the repository root. Dependencies: `pip install -r requirements.txt`; pandoc and Microsoft
Word are system-level (`--doctor` verifies both).

## When to use / when not to

Use it for:

- **Chinese formal documents from Markdown**: tenders, bids, grant applications, final reports,
  official documents — with an explicit design contract (config + `ref.docx`, or `reference_doc`).
- **One-command pre-delivery acceptance**: 9 checks, structured report, evidence package.
- **Agent-friendly Patch editing**: declarative operations with dry-run and hash preconditions.
- **Targeted edits to an existing docx**: text, paragraphs, table cells, headers/footers — without
  re-typesetting anything else.

Do **not** use it for:

- **Comments, tracked changes, redaction, watermarks, content controls** on an existing docx —
  out of scope by design; use a general-purpose docx skill.
- **Bulk content filling** — the editing chain is for targeted point edits. Do the bulk fill
  (`--fill` covers coordinate-based batch filling; anything smarter is on the caller), then use
  this skill for typesetting and acceptance.
- **Theses or book manuscripts** — no bibliography (citeproc), equation numbering, or odd/even headers.
- **English-language documents** — defaults (A4, SimSun/SimHei, full-width punctuation, Chinese
  caption keywords) are Chinese-document conventions.

## Hard contract and pitfalls

1. **Layout only, never content — scoped to `post.py`.** `post.py` does not touch any text in body or
   captions. Figure and table numbers are **hand-written** in the source; renumber manually after
   insert/delete. An `auto_number` feature once existed, silently corrupted body text, and was
   removed entirely — do not reintroduce caption rewriting. If automatic numbering is ever needed:
   Word `SEQ`/`REF` fields. Guarded by `tests/test_postprocess.py::test_never_touches_text`.
   The scope matters: `content_fixes` / `content_fixes_file` run *earlier*, in the render pipeline
   (after merging Markdown, before conversion), and they do rewrite text — on purpose, from a table
   the user supplied. Never tell an agent "output text is byte-identical to the source" without
   checking whether that config key is populated.
2. **Cover lines carry no leading/trailing blanks** — `post.py` appends one blank `CoverInfo`
   paragraph itself; extra blanks overflow the cover onto a second page.
3. **Do not put content before the first `#`** — it lands after the TOC (or at the top when
   `toc:false`); `post.py` warns but never deletes it (contract #1).
4. **Body/heading fonts live in the template** — change with
   `python scripts/make_ref.py --body-font 楷体 --body-size 14`; config `style` fonts cover tables
   and captions only.
5. **Grid tables align by display width** — a CJK character counts as two columns. "Looks aligned in
   a monospace editor" can still parse into a single-column broken table; verify by rendering.
6. **`--check` flags legitimately sparse pages** (signature block, heading alone before a table) —
   false positives; relax with `--max-empty N`.
7. **Snapshot baselines are machine-specific** (Word version + fonts) — re-record with
   `snapshot.py --update` after switching machines.
8. **Client-mandated template wins** — point `reference_doc` at their docx; cover sheets, sealing,
   signature pages and page-number rules still need manual compliance review. Inheritance details:
   README §Reusing an existing template.
9. **After editing an existing docx, always `--verify`** (or open in Word) — the only reliable proof
   the OOXML survived is that Word still opens the file.

## Minimal config shape

```json
{
  "cover": [["CoverTitle", "投 标 文 件"], ["CoverInfo", "投标人：示例科技有限公司"]],
  "header": "示例项目 · 投标文件",
  "toc": true,
  "style": {"table_zebra": true}
}
```

Short documents (notices/announcements) that need headings but no table of contents: `"toc": false` —
heading styles stay intact, no TOC page. Form-style documents (no `#` headings) need
`"style": {"header_rows": 0}` so the field-name first row isn't shaded as a header.

Every field and `style` key: `docs/CONFIG.md`. Table syntax (pipe vs grid, multi-level headers,
merged cells, column widths): `docs/TABLES.md`.

## Two chains, opposite contracts

| | Rendering chain | Editing chain |
|---|---|---|
| Input | Markdown | an existing docx |
| Contract | layout only, never content | **change only what is asked; leave every other byte alone** |

Full command reference for both: README §Usage, `docs/EDITING.md` and `docs/SCRIPT_HELP.md`.
Two rules that matter for agents:

1. **Word splits text into runs unpredictably.** `edit.py` matches against the concatenated
   paragraph text and writes the replacement into the run where the match starts, so per-run
   formatting survives — never hand-edit `run.text`.
2. **An anchor matching more than one paragraph is refused.** After a TOC refresh, a heading exists
   both in the TOC and in the body; use a precise anchor or `--all-anchors`.

Re-typesetting someone else's docx goes through Markdown (`pandoc` → clean → render); the cleaning
steps are in README §Reusing an existing template.

## Acceptance gate

After any real render, the delivery gate is one command:

```bash
python scripts/validate.py bid.docx --out evidence/
```

Nine checks run against a **single** shared Word export: package integrity, source content
(`--source-md` body comparison, else `--expected-hash` artifact hash), image embedding, section
count, TOC field, page numbering (footer region only, continuity), blank pages (threshold), Word
acceptance, visual drift (baseline, **all pages** by default). Exit code 1 on FAIL or ERROR.
Output: `report.json` + sampled page screenshots + `signature` — a **checksum manifest** (docx + report
hashes, plus how many checks were skipped). It is not a cryptographic signature: no key, so it proves
*which artifact this evidence describes*, not that the evidence was not tampered with.

Read the statuses before trusting the gate: a check that could not run reports `SKIP` and is not
counted as passed. `Passed: 7/9 (skipped: 2)` means two things were never verified — treat that as
weaker evidence, not as a pass. Check-by-check table: `docs/VALIDATION.md`.

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
`add_row` / `del_row`. Flow: `--dry-run` → `--apply --validate`. Full schema and CLI:
`docs/SCRIPT_HELP.md` §patch.py.
