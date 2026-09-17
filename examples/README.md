# examples

可运行的最小示例。每个目录自带 Markdown 与 config，直接渲染即可：

```bash
cd <仓库根>

python scripts/render.py --src examples/form   --out examples/form/form.docx     --config examples/form/config.json   --pdf --check
python scripts/render.py --src examples/tables --out examples/tables/tables.docx --config examples/tables/config.json --pdf --check
```

| 目录 | 演示什么 | 关键配置 |
|---|---|---|
| `form/` | 表单式/申报书：无 `#` 标题，表格首行是字段名而非表头 | `style.header_rows: 0` |
| `tables/` | grid table 多级表头与合并单元格、列宽控制、斑马纹 | `style.table_zebra` |

两个示例都会生成 `*.docx` / `*.pdf`（已被 `.gitignore` 忽略，不会进版本库）。

> 示例正文是中文的，这是有意的——本工具的排版默认值就是中文正式文档惯例
> （A4 / 宋体正文 / 黑体标题 / 全角标点）。参见 `README.md` 的 Scope note。
