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
└── signature            # checksum manifest: docx + report hashes, check summary
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
| ✅ Word acceptance | Real Word opens and exports PDF successfully |
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
✅ [PASS] renderer_acceptance: Word 验收：OK
⏭️ [SKIP] visual_drift: 跳过 (未提供基线目录)

Summary
────────────────────────────
Passed: 7/9 (skipped: 2)

✅ All checks passed

Evidence package saved to: evidence/
  - report.json (structured validation report)
  - page-XXX.png (sample screenshots)
  - signature (checksum manifest: docx + report hashes)
```

If any check fails, exit code is 1 and you get a detailed error message. Flags
(`--profile`, `--max-empty`, `--quiet`, …): `docs/SCRIPT_HELP.md` §validate.py.

### Manual gate

The 9 checks cover structure; four things stay human, on the first pass over any new document:

1. `images: n/m ok` in the render log with **n == m** (m = count of `![` in the source)
2. `near-empty pages: 0`
3. `OK` — Word opened it and exported the PDF
4. **Look at the rendered pages.** Machines count pages, images and blanks; they can't tell you the figure is
   wrong or the header row got clipped

