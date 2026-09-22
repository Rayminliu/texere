# EDITING.md

> 从 `README.md` 下沉而来的明细。中英两份由 `tests/test_docs_sync.py` 守着同构。

## Editing an existing docx

`scripts/edit.py` is a **separate chain** from rendering, with the opposite contract:

| | Rendering chain | Editing chain |
|---|---|---|
| Input | Markdown | an existing docx |
| Contract | layout only, never content | **change only what is asked; leave every other byte alone** |
| Never | rewrite text | inject a cover, TOC, page numbers, or re-typeset styles |

The shapes of the operations (full option list: `docs/SCRIPT_HELP.md` §edit.py):

```bash
python scripts/edit.py 标书.docx --list                      # structure: sections / tables / paragraphs
python scripts/edit.py 标书.docx --replace "旧=新" --scope body,tables,header,footer
python scripts/edit.py 标书.docx --after "锚点文字" --text "新段落"   # also --before / --delete
python scripts/edit.py 标书.docx --cell 0 2 1 "1,060,000"            # table / row / col / value
python scripts/edit.py 标书.docx --add-rows 0 3 --template-row 2     # also --add-row / --del-row
python scripts/edit.py 标书.docx --fill data.json                    # batch fill, JSON or CSV
python scripts/edit.py 标书.docx --header "新版页眉" --footer "— X —" --section all
python scripts/edit.py 标书.docx --replace "A=B" --verify            # let Word open the result
```

Backups go to `<name>.bak.docx` by default (`--no-backup` disables, `--out` writes elsewhere).

**Why editing in place is safe**: `python-docx` round-trips a docx without losing any part — opening and
saving a 4-page document loses and adds zero package parts (measured; the file only gets smaller because
it is re-zipped).

**Two traps worth knowing**

1. *Word splits text into runs unpredictably.* `自开标之日起 90 日历天` can be three runs with the number
   on its own, so a naive `run.text` search misses it. `edit.py` matches against the whole paragraph text
   and writes the replacement into the run where the match starts, so the run count and each run's
   formatting survive.
2. *An anchor matching several paragraphs is refused.* Once Word refreshes a TOC field, a heading like
   `1.2 资质与业绩` exists both in the TOC and in the body; inserting after it would put content into the
   table of contents. Use a more precise anchor, or pass `--all-anchors`.

**Not supported, deliberately**: tracked changes, comments, watermarks, content controls, and `.docm`
files with macros (`python-docx` drops `vbaProject.bin` on save).

### Re-typesetting instead of editing

To bring someone else's docx into this project's format, go through Markdown — but clean it first:

```bash
pandoc 甲方文档.docx -t markdown --wrap=none --extract-media=media -o 01_内容.md
# 手工清理：删掉原来的封面行与目录块、去掉表头残留的 ** 加粗
python scripts/render.py --src . --out 新版.docx --config cfg.json --pdf --check
```

Measured on a 4-page document: without cleaning, the result has a **duplicated cover and TOC**, leaks
`[1](#...)` link syntax into the body, and grows from 4 to 5 pages. Image references from
`--extract-media` are relative to the working directory, so run from the repository root or put `media/`
inside `--src`.

