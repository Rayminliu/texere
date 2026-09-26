# TABLES.md

> 从 `README.md` 下沉而来的表格参考。中英两份由 `tests/test_docs_sync.py` 守着同构。

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

> Misaligned pipes (CJK wide characters are the usual culprit) don't need hand-fixing:
> `python scripts/align_tables.py doc.md --fix` re-aligns by display width with cell bytes preserved.
> See [SCRIPT_HELP](SCRIPT_HELP.md).

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

Defaults live in [The `style` section](CONFIG.md#the-style-section); this is the *when do I reach for it* half.

- **`header_rows`** — the only key you're likely to set per document. Leave it alone and each table's header is
  detected from the source (everything above `+===+` in a grid table). Set `0` for form-style tables whose first
  row is a field name, not a column title — that also stops the header from repeating across pages.
- **`table_border: "three"`** — three-line (booktabs) look for analytical tables; `none` for layout tables that
  shouldn't read as data. Rule weight comes from `border_size` / `three_line_size`.
- **`table_shade` + `table_header_color`** — on a dark fill, pair them (`table_header_color: "FFFFFF"`), or the
  header text disappears against its own background.
- **`table_zebra`** — alternate row shading for wide tables; the first data row stays white.
- **Repeating headers need no configuration**: pandoc already sets `w:tblHeader` on header rows. The line in
  `post.py` is just idempotent hardening.

