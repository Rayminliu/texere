# VALIDATION.md

> 从 `README.md` 下沉而来的明细。中英两份由 `tests/test_docs_sync.py` 守着同构。

## Validation and evidence package

After rendering, one command turns "I think it's fine" into an auditable artifact:

```bash
python scripts/validate.py bid.docx --out evidence/
```

This produces:

```
evidence/
├── report.json          # Structured validation report (9 checks)
├── page-001.png         # Sample screenshots (first, middle, last pages)
├── page-069.png
├── page-272.png
└── signature            # checksum manifest: docx + report + all screenshot hashes, check summary
```

The file is named `signature` for historical reasons, but it is a **checksum manifest, not a cryptographic
signature** — there is no key, so anyone can recompute these hashes. It proves *which artifact this evidence
describes*, not *that the evidence was not tampered with*.

The report contains 9 automated checks:

| Check | What it verifies |
|---|---|
| ✅ Package integrity | DOCX is a valid ZIP with required parts |
| ✅ Source content | `--source-md`: every Markdown segment is present in the docx body. `--expected-hash`: file-level SHA256 of the docx (byte equality only). Neither given → SKIP |
| ✅ Image embedding | With `--source-md`: per-image SHA256 identity + document order (catches swapped, wrong, or duplicated images; pandoc embeds bytes verbatim). Without it: embedded count ≥ referenced count, a lower bound only |
| ✅ Section count | Reasonable number of sections (1–100) |
| ✅ TOC field | A real `TOC` field exists in the OOXML (document has no TOC → SKIP) |
| ✅ Page numbering | Footer numbers form a gap-free sequence (no numbers detected → SKIP) |
| ✅ Blank pages | Below threshold (default: 0 allowed) |
| ✅ renderer acceptance | The selected renderer opens the file and exports PDF successfully |
| ✅ Visual baseline drift | Per-page pixel comparison against the baseline — all pages by default |

**Four statuses, and SKIP is not PASS.** Every check reports `PASS` / `FAIL` / `SKIP` / `ERROR`.
`SKIP` means a precondition was missing — no baseline, no `--source-md`, no PyMuPDF — so the check
never actually ran; it is excluded from the passed count and does not by itself fail the gate. Only
`FAIL` and `ERROR` set exit code 1. `Passed: 7/9 (skipped: 2)` is a materially weaker statement than
`Passed: 9/9`, and the report says so out loud.

Example output:

```
Document Validation
────────────────────────────
✅ [PASS] package_integrity: OK
⏭️ [SKIP] source_content: 跳过 (未提供 --source-md 或 --expected-hash)
✅ [PASS] image_embedding: 图片逐图比对：28/28 张身份与顺序一致
✅ [PASS] section_count: 分节数：3 (合理)
✅ [PASS] toc_field: 目录域：1 个 TOC 域
✅ [PASS] page_numbering: 页码：69 页 (连续，检测到页码 1-69)
✅ [PASS] blank_pages: 空白页：0/69 (阈值：0)
✅ [PASS] renderer_acceptance: 渲染器验收：OK
⏭️ [SKIP] visual_drift: 跳过 (未提供基线目录；如需版面漂移防护：先 python scripts/snapshot.py <pdf> --update 录基线，再用 --baseline <目录> 或 profile 的 baseline_dir 指定)

Summary
────────────────────────────
Passed: 7/9 (skipped: 2)

✅ All checks passed

Evidence package saved to: evidence/
  - report.json (structured validation report)
  - page-XXX.png (sample screenshots)
  - signature (checksum manifest: docx + report + all screenshot hashes)
```

If any check fails, exit code is 1 and you get a detailed error message. Flags
(`--profile`, `--max-empty`, `--quiet`, …): `docs/SCRIPT_HELP.md` §validate.py.

### Profile contract boundary (profile = acceptance contract, config = rendering input)

A profile and a render config are **two different objects**, deliberately not connected:

- `profiles/*.json` is an **acceptance contract**: it describes what the finished document must satisfy.
  It is consumed only by `validate.py` (`--profile` attaches it to the evidence report; `--enforce-profile`
  compiles its declared fields into `profile.<field>` gate assertions).
- `config.json`'s `style.*` is the **rendering input**: it is what `post.py` actually consumes to typeset the
  document.
- Injecting profile into rendering is a **non-goal**: two sources of truth for the same layout knobs
  (contract vs input) would drift apart and make gate verdicts unexplainable.

Which profile fields the gate actually enforces (details: `docs/SCRIPT_HELP.md` §Profile):

| Category | Enforced when `--enforce-profile` | Declared-only (not asserted) |
|---|---|---|
| Page | `page.width/height` (±0.1cm), `page.margin_*` (±0.2cm) | — |
| Body font | `styles.body.font_eastAsia` / `font_latin` / `size` (±0.5pt) | `line_spacing`, `first_line_indent`, `space_*` |
| Headings | `styles.h1/h2/h3` eastAsia / latin / size / bold | `page_break_before`, `space_*` |
| Tables | `table.border` ≠ none → at least one visible border | `border_size`, `border_color`, `header_shade`, `zebra`, … |
| TOC | `toc` section declared → a TOC field exists | `toc.depth`, `toc.title`, `placeholder`, … |
| Other | — | `caption.*`, `header.*`, `footer.*`, `*_specific.*` |

Optional keys (valid in any profile; none of the five built-in profiles uses them): `baseline_dir`
(visual-drift fallback when `--baseline` is not passed) and `acceptance.renderers`
(multi-renderer acceptance matrix, e.g. `["word", "wps"]`; consumed by `policy.py`, which is not wired into the validate exit code).

**Profiles are deliberately self-contained — no inheritance** (a design decision, not an oversight):
each of the five built-in profiles states its whole contract, there is no `base` / `extends`. The reason:
a profile is the acceptance contract that gets audited on its own, so "which contract did this run sign"
must stay readable at a glance; inheritance turns it into "whatever the merge produced" and would change
the profile embedded in the evidence and the `report_hash`. The price is that the same layout facts are
written out in several profiles, so drift is not left to human review:
`tests/test_profiles.py::test_duplicated_segments_match_formal_cn_byte_for_byte` requires every segment
shared with `formal-cn-v1` to be byte-identical (the segment list lives in `SHARED_SEGMENTS_WITH_FORMAL`).
If a fork is intentional, document it in this section and drop the segment from the guard — never change it silently.

### Manual gate

The 9 checks cover structure; four things stay human, on the first pass over any new document:

1. `images: n/m ok` in the render log with **n == m** (m = count of `![` in the source)
2. `near-empty pages: 0`
3. `OK` — Word opened it and exported the PDF
4. **Look at the rendered pages.** Machines count pages, images and blanks; they can't tell you the figure is
   wrong or the header row got clipped

## Capability modules and test layers

The pipeline's pure logic lives in importable sibling modules under `scripts/`, while the user-facing
scripts stay thin CLI shells. The split exists so the logic can be exercised in-process without paying for
a Word launch:

| Capability module | Owns | Shell that re-exports it |
|---|---|---|
| `_verify.py` | check primitives + `CheckResult` / status constants | `validate.py` |
| `_visual_diff.py` | `diff_ratio`, page sampling, `compute_visual_diff` | `validate.py`, `snapshot.py` |
| `_evidence.py` | report skeleton, tally, `report.json` / signature writers | `validate.py` |
| `_compile.py` | pandoc command builder + code-fence / prompt helpers | `render.py` |
| `_mutate.py` | declarative Patch engine + `MutationResult` op primitives | `patch.py`, `edit.py` |

Each shell imports the same names it used to define, so `stdout` wording, `argparse` flags and exit codes
are unchanged — the re-export is itself the equivalence proof: the existing subprocess and in-process
tests pass without a single edit. Dependency direction is a one-way DAG
(`_version → _shared → {capability modules} → shells`); a capability module never imports its shell.

Tests are layered so the slow boundary is crossed only where it must be:

- **L1 / L2 — in-process unit + logic** (`test_validate_units.py`, `test_patch.py::TestPatchUnits`): call
  the re-exported functions directly against in-memory docx / PyMuPDF fixtures. No Word, sub-second.
- **L3 — CLI contract** (subprocess): assert the frozen `stdout` wording, `--help`, `report.json` shape and
  exit codes; marked `@pytest.mark.word` when a real renderer must boot.
- **L4 — end-to-end** acceptance against a golden document.

CI runs the `word`-marked cases separately (`pytest -m "not word"` then `-m word`), so pure-logic
regressions stay fast while layout-affecting changes still gate on the 0.00% pixel snapshot.

