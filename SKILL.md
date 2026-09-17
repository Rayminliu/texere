---
name: docx-kit
description: 把 Markdown 渲染成排版合格的中文正式 docx 与 PDF（标书、申报书、结题报告、白皮书），并用真 Word 做打开验收、导出 PDF、检查空白页与版式漂移。当任务涉及"从 Markdown 生成正式文档""给文档加封面/目录/页眉/分节页码""中文正式排版"（宋体正文、黑体标题、首行缩进两字符、题注居中）"表格跨页重复表头""交付前检查空白页或版式变化"时使用本技能。
---

# docx-kit

把 Markdown 章节目录渲染成中文正式 docx / PDF 的本地工具包。设计核心是两条：
**样式表驱动**（全部视觉规则写在 `ref.docx`，源文本只写语义）与**真机验收**
（用 Word 打开并导 PDF，再肉眼过渲染图，不信任何"读回正常"）。

## 硬契约：只改版式，不改内容

`post.py` 不触碰正文与题注的**任何文字**。输出里的每一个字都来自源 Markdown。

不要给这条管线添加"改写文字"的功能。历史上曾有过 `auto_number`（图表自动编号），
它静默篡改了正文（`**表层…**` → `表 3-5 层…`）、并给原本无编号的题注补号，
因此被整体删除。`tests/test_postprocess.py::test_never_touches_text` 守着这条契约。

需要图表自动编号与交叉引用时，正确做法是插入 Word 原生 `SEQ` / `REF` 域
（域可更新、不改文本），而不是重写题注文字。

## 何时使用 / 何时不使用

使用：

- 从 Markdown 生成中文正式文档（标书、申报书、结题报告、白皮书、方案建议书）
- 需要封面、目录域、分节页码、页眉、表题与图注规范、表格跨页重复表头
- 交付前需要一键验收（Word 能否打开、有无空白页、版式有没有变）

不使用：

- **编辑已有 docx**：批注、修订痕迹、脱敏、无障碍修复、水印、表单控件。
  本工具没有编辑能力，也不该有——那类任务请用通用 docx 技能（例如 Codex 的
  Documents 插件，它走 LibreOffice 渲染 + 36 个 OOXML 脚本）。
- 学位论文、书稿：需要参考文献（citeproc）、公式编号、奇偶页不同页眉，本工具不支持。

## 命令

```bash
python render.py --doctor                    # 环境自检：pandoc / 依赖 / Word 引擎身份
python render.py --version                   # 版本
python render.py --sample                    # 冒烟测试（产出 sample_out.docx/.pdf）

# 正式渲染（--check 会自动补 --pdf）
python render.py --src <md目录> --out 标书.docx --config cfg.json --pdf --check

python -m pytest -q                          # 33 项断言，不需要 Word，约 8 秒
python snapshot.py 标书.pdf                  # 版式快照回归，漂移即 exit 1
python snapshot.py 标书.pdf --update         # 确认版式变更后重录基线
python make_ref.py --body-font 楷体          # 重建排版模板（改字体/字号用）
```

## 输入要求

- `--src` 指向一个目录，里面的 `*.md` **按文件名排序**后合并（用 `01_`、`02_` 前缀控制顺序）
- `#` 为章、`##` 为节；表题写成独立一行 `表 1-1 标题`；图注写进图片 alt：`![图 1-1 标题](a.png){width=13cm}`
- 图片搜索路径：`src` 目录、其全部子目录、**`src` 的父目录及其子目录**（图片常在兄弟目录），
  仍找不到就用 config 的 `resource_paths` 补充
- 第一个 `#` **之前不要放内容**（不会被当成封面，且会排在目录之后）；真放了会打印 `[warn]`，但不会替你删
- config.json 与源 md 带 UTF-8 BOM 也能读

## 验证门槛：四项全过才算完成

任何一次正式渲染之后，逐项确认：

1. 日志出现 `images: n/m ok` 且 **n == m**（m 是源 md 中 `![` 的出现次数）
2. `near-empty pages: 0`
3. `OK`（Word 成功打开并导出 PDF）
4. **首次用于某份新文档时，必须人工看一遍 PDF 渲染图**——机器能查页数、图片数、
   空白页，查不了"这张图画得对不对、表头有没有被截断"

`--check` 对「内容稀疏但合法」的页面也会报 FAIL（表单尾页的签字盖章区、
大表格前的单独标题页）。这类是误报，用 `--max-empty N` 放宽。

## 配置

config.json 全部字段见 `README.md` 的字段表。最常用的：

| 键 | 说明 |
|---|---|
| `cover` | 封面行 `[[样式名, 文本], …]`；样式名用 CoverTop/CoverTitle/CoverSub/CoverInfo/CoverDate。**不要再写行首/行尾空行**（`post.py` 会自动补一个空 CoverInfo 段，多写会导致封面溢出成两页） |
| `header` | 正文节页眉 |
| `reference_doc` | 用招标方/甲方给的 docx 当模板 |
| `content_fixes_file` | 编辑性替换表；指向 `.py` 时用 `ast` 只读取值、不执行代码 |
| `resource_paths` | 额外的图片搜索目录 |
| `style` | 版式微调（页码模板、目录深度、题注颜色字号、表格边框、`header_rows: 0` 表示该表无表头…） |

## 已知边界

- **Windows + Word 绑定**：`--pdf` 需要本机 Word（COM）。Linux/macOS 只能出 docx
- **快照基线绑本机**：换机器或换 Word 版本后先 `snapshot.py --update` 重录
- 图表编号为**手写**，插删图表后需人工对号（本工具不代改编号，见硬契约）
- 中文标点是全角、表格用中式全框线（或 `table_border: three` 三线表），
  与英文文档惯例（Letter 纸、西文字体）不同——本工具面向中文正式文档

## 文件地图

| 文件 | 职责 |
|---|---|
| `render.py` | 入口：合并 md → pandoc → post.py →（可选）finalize/check |
| `post.py` | 后处理：封面、目录域、分节页码、页眉页脚、表格规则、题注样式 |
| `filters/captions.lua` | pandoc Lua filter：在 **AST 层**把表题/图注标成语义样式，post.py 不必用正则猜 |
| `make_ref.py` | 重建 `ref.docx`（改字体/字号/间距时） |
| `finalize.py` | Word COM：打开验收 + 刷目录域 + 导 PDF |
| `check_pdf.py` | PyMuPDF：空白页检测（带退出码）+ 渲染 PNG |
| `snapshot.py` | PDF 版式快照回归：逐像素比对 `baselines/` |
| `tests/` | 33 项 pytest 断言 |
| `README.md` | 完整字段表、排版军规、标书场景注意事项 |
| `CHANGELOG.md` | 版本历史与每条修复的理由 |

## 性能参考

实测（含 Word 验收 + 导 PDF + 空白页检查）：

| 规模 | 耗时 |
|---|---|
| 69 页 / 63 表 / 28 图 | **19.0 秒** |
| 272 页 / 252 表 / 112 图 | **69.7 / 66.9 秒**（连跑两次，页数、字数、表格、图片及抽查页像素全部一致，Word 进程无残留） |
