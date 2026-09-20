# examples

可运行的最小示例。每个目录自带 Markdown 与 config，直接渲染即可：

```bash
cd <仓库根>

python scripts/render.py --src examples/form     --out examples/form/form.docx       --config examples/form/config.json     --pdf --check
python scripts/render.py --src examples/tables   --out examples/tables/tables.docx   --config examples/tables/config.json   --pdf --check
python scripts/render.py --src examples/tender   --out examples/tender/tender.docx   --config examples/tender/config.json   --pdf --check
python scripts/render.py --src examples/gongwen  --out examples/gongwen/gongwen.docx --config examples/gongwen/config.json  --pdf --check
python scripts/render.py --src examples/minutes  --out examples/minutes/minutes.docx --config examples/minutes/config.json  --pdf --check
python scripts/render.py --src examples/report   --out examples/report/report.docx   --config examples/report/config.json   --pdf --check
python scripts/render.py --src examples/contract --out examples/contract/contract.docx --config examples/contract/config.json --pdf --check
```

| 目录 | 场景 | 演示什么 | 关键配置 |
|---|---|---|---|
| `form/` | 项目申报书 | 表单式文档：无 `#` 标题、跳过目录，表格首行是字段名 | `style.header_rows: 0` |
| `tables/` | 报价与配置明细 | grid table 多级表头与合并单元格、列宽控制、斑马纹 | `style.table_zebra` |
| `tender/` | 投标文件 | 封面 + 目录 + 章节页码 + 商务/技术响应表 | `cover`、`caption_words` |
| `gongwen/` | 公文请示 | 中文编号层级（一、（一）、1.）、仿宋、页码装饰线 | `style.east_font`、`page_number` |
| `minutes/` | 会议纪要 | 短文档：封面信息行承载时间地点、决议清单表 | `cover` 多行 `CoverInfo` |
| `report/` | 经营分析报告 | **多文件合并**（按文件名序）、Lead 灰底提示框、grid 二级表头、图注 | `::: {custom-style="Lead"}`、`table_shade` |
| `contract/` | 技术服务合同 | 条款章节 + 价款表 + grid 单元格内换行 + 签署栏 | grid 多行单元格 |

所有示例的 `*.docx` / `*.pdf` 渲染产物已被 `.gitignore` 忽略，不会进版本库；
示例里的架构图/趋势图是纯色占位 PNG，换成真实图片即可。

> 示例正文是中文的，这是有意的——本工具的排版默认值就是中文正式文档惯例
> （A4 / 宋体正文 / 黑体标题 / 全角标点）。参见 `README.md` 的 Scope note。
