# docx-kit —— 中文正式文档渲染工具包

把 Markdown 变成**排版合格的中文正式 docx**（标书 / 计划书 / 申报书 / 结题报告 / 白皮书）。
从"农域多视图"参赛计划书管线抽离，2026-09-13 独立成包；总体积约 40KB，零新依赖。

## 文件

| 文件 | 职责 |
|---|---|
| `ref.docx` | 中文排版模板：宋体小四正文、黑体标题阶梯、表格边框、题注样式、封面自定义样式（CoverTop/CoverTitle/CoverSub/CoverInfo/CoverDate）、提示框样式（Lead/SmallNote） |
| `make_ref.py` | 重新生成 ref.docx（改字体/字号/间距时用它） |
| `render.py` | 入口：合并 md → pandoc → post.py →（可选）finalize/check |
| `post.py` | 后处理：封面注入、目录域、分节页码、页眉页脚、表格规则、表题图注居中；OOXML 按 schema 顺序插入 |
| `finalize.py` | Word COM：打开验收（打不开=结构错）、刷目录域、导 PDF、存回 |
| `check_pdf.py` | PyMuPDF：空白页检测 + 渲染页面 PNG 供目视；空白页超阈值 exit 1 |
| `snapshot.py` | PDF 版式快照回归：逐像素比对 `baselines/`，漂移即 exit 1 |
| `filters/captions.lua` | pandoc Lua filter：在 **AST 层**把表题/图注标记成 `TableCaption`/`FigureCaption`，`post.py` 不再靠正则猜 |
| `tests/` | pytest 断言：排版军规、自动编号与交叉引用、退出码、快照逻辑 |
| `sample.md` / `sample_config.json` | 冒烟测试样例（投标文件风格） |

## 依赖

依赖声明以 `pyproject.toml` 为唯一事实来源，`uv.lock` 锁定精确版本，
`requirements.txt` 由 `uv export` 生成、供纯 pip 环境使用。

| 依赖 | 用途 | 安装 |
|---|---|---|
| pandoc >= 3.1 | **必需**，md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx、lxml | **必需**，docx 读写与 OOXML 处理 | `pip install .` |
| pywin32 | 仅 `--pdf`（Word 验收 + 导 PDF） | `pip install ".[pdf]"` |
| PyMuPDF | 仅 `--check`（PDF 目视验收） | `pip install ".[check]"` |
| Microsoft Word | 仅 `--pdf` | 系统级，pip/uv 都装不了 |

**不需要 Quarto/LibreOffice/商业库。** 换机器后先跑 `python render.py --doctor`，
它会报告 pandoc 版本、缺失的依赖，并**实测 Word 引擎身份**（走 COM 直接问，
不读注册表——注册表里的 `CurVer` 可能是旧 Office 卸载后的残留值，会误报）。

## 安装

```powershell
winget install --id JohnMacFarlane.Pandoc   # 必需，md -> docx
pip install .                                # 只出 docx 的最小集
pip install ".[pdf,check]"                   # 需要 Word 验收 + PDF 目视时
python render.py --doctor                    # 确认装齐
```

## 用法

```
python render.py --doctor                          # 环境自检（pandoc / 依赖 / Word 版本）
python render.py --sample                          # 冒烟测试，产出 sample_out.docx/.pdf
python render.py --src 章节目录 --out 标书.docx --config config.json --pdf --check

python -m pytest -q                            # 排版断言（不需要 Word）
python snapshot.py 标书.pdf --update           # 录版式基线（确认版式无误后执行）
python snapshot.py 标书.pdf                    # 回归比对，漂移即 exit 1
```

- `章节目录` 内放 `01_xxx.md … 0N_xxx.md`，按文件名排序合并；`#` 为章、`##` 为节；
- 表格用 pipe table，表题写成独立一行 `表 1-1 标题`（会自动居中灰字）；
- 图用 `![图 1-1 标题](path.jpg){width=13cm}`（自动居中灰字题注）；
- 封面与页眉在 config.json 里配（见 sample_config.json）；不配 cover 则只注入目录；
- 强调用 `**加粗**`，灰底提示框用 `::: {custom-style="Lead"} … :::`。

**图表自动编号与交叉引用（可选）**：config 里加 `"auto_number": true` 后——

| 场景 | 写法 | 结果 |
|---|---|---|
| 表题 | `表 商务条款响应表 @tab:clause` | `表 1-1 商务条款响应表` |
| 图注 | `![图 架构示意](a.png) @fig:arch` | `图 1-2 架构示意` |
| 引用 | `详见 @tab:clause` | `详见 表 1-1` |

编号按「章-序」自动生成；**已手写的编号会被重排**（比如写成 `表 9-9` 也会被纠正），
所以插入或删除图表后不用再人工对号。不开这个开关则完全不动原文。

### config.json 字段

| 字段 | 作用 |
|---|---|
| `cover` | 封面行 `[[样式名, 文本], …]` |
| `header` | 正文节页眉 |
| `title` / `author` / `subject` / `comments` | 文档属性 |
| `reference_doc` | 指定模板 docx（招标方给强制格式时用它） |
| `toc_heading` | 目录标题，默认「目　　录」 |
| `auto_number` | `true` 开启图表自动编号与交叉引用 |
| `style` | 版式微调，见下表 |

`style` 段（全部可省，缺省即中文正式文档惯例）：

| 键 | 默认 | 说明 |
|---|---|---|
| `page_number` | `— {n} —` | 页码模板，`{n}` 处插入页码域 |
| `toc_depth` | `1-2` | 目录收录层级 |
| `caption_gray` | `404040` | 题注颜色（6 位十六进制） |
| `caption_size` | `10.5` | 题注字号 pt |
| `table_shade` | `EDEDED` | 表头底纹 |
| `table_size` | `10.5` | 表格字号 pt |
| `east_font` / `latin_font` | `宋体` / `Times New Roman` | **只作用于表格与题注** |

## 表格

### 两种写法，按需选

| | pipe table | grid table |
|---|---|---|
| 写法 | `\| a \| b \|` | `+---+---+` |
| 多级表头 | ❌ | ✅（`+===+` 以上全是表头行） |
| 合并单元格 | ❌ | ✅（跨列、跨行） |
| 单元格内换行 | ❌ | ✅ |
| **控制列宽** | ❌ 一律等宽 | ✅ 见下 |
| 列对齐 | ✅ `:--` `:-:` `--:` | ❌ |

模块清单、报价表这类简单表用 pipe table；**复杂表头、合并单元格、要控制列宽的，一律用 grid table**。

### grid table 三板斧

**1. 多级表头 + 合并单元格**——`+===+` 分隔表头与表体；某一行不写中间竖线即为跨列：

```
+------------------+------------------+
| 商务部分         | 技术部分         |
+--------+---------+--------+---------+
| 条款   | 响应    | 模块   | 说明   |
+========+=========+========+=========+
| 工期   | 完全响应| 接入层 | 设备   |
+--------+---------+--------+---------+
```

**2. 控制列宽**——pandoc 按**最细的列分隔**分配列宽。上表最细一行是 `+--------+---------+`（8/9/8/9），
四列实际宽度就是 8:9:8:9（实测 990/1100/990/1100 twips）。想让「序号」列窄，把它写窄就行。

**3. 单元格内换行**——同一个单元格里写多行即可（pipe table 做不到）。

### 视觉控制（config 的 `style` 段）

| 键 | 默认 | 说明 |
|---|---|---|
| `header_rows` | `1` | 前 N 行做**视觉**表头（灰底加粗居中）；多级表头要设 `2` |
| `table_border` | `full` | `full` 全框线 / `three` 三线表 / `none` 无框线 |
| `table_shade` | `EDEDED` | 表头底纹 |
| `table_size` | `10.5` | 表格字号 pt |

> 跨页重复表头**不用配**：pandoc 原生就给表头行设了 `w:tblHeader`
> （grid table 的 `+===+` 以上全部算表头），`post.py` 里那行只是幂等加固。

## 标书场景三条注意

1. **招标方给了强制格式模板时，以对方为准**：把它的 docx 路径写进 config 的 `"reference_doc"`，正文样式即继承对方模板；封面/密封/签字页/页码规则仍按招标文件手工核对，本工具不替代合规审查。
2. 标书常见结构（投标函/商务标/技术标/报价/资质业绩）直接对应 `#` 章即可；报价表建议保留 pipe table，便于后期整体替换为招标方表格。
3. 交付前跑 `--pdf --check` 并肉眼过一遍渲染图：Word 能打开、无空白页、表头灰底、题注居中，四条都过再封包。

## 排版六条军规（改模板或手写内容时遵守）

1. 样式表驱动：视觉规则只写在 ref.docx，源文本只写语义；
2. 中文四件套：宋体正文 + 黑体标题 + eastAsia 字体属性 + 首行缩进 2 字符；
3. 表格三件套：100% 宽 + 灰底加粗居中表头 + 跨页重复表头；
4. 题注居中、小一号、灰色、**去斜体**；
5. 中西混排：中文宋体、数字西文 Times New Roman；
6. 交付前真 Word 打开 + 渲染目视，不信任何"读回正常"。

## 已知边界

- 目录为 Word 域，首次打开若未刷新请全选按 F9（已设 updateFields，通常自动）；
- 图表编号：默认手写；开启 `auto_number` 后可自动按章编号并交叉引用（见「用法」）。
  未开启时插删图表仍需人工对号；
- 不使用 Quarto：其 1.10.x 的 docx 对带自动编号题注的表格会丢失表体；
- 快照基线依赖本机 Word 版本与字体，**换机器后先 `snapshot.py --update` 重录**，
  否则会满屏漂移；阈值默认 0.1%（实测同文档重复导出为 0.00%，改一处页眉为 0.16%）；
- **正文与标题的字体/字号在模板层**，改它是 `python make_ref.py --body-font 楷体 --body-size 14`
  （`style` 段的字体只管表格与题注，管不到正文）；
- **grid table 对空格敏感**：每行的竖线必须严格对齐，否则会解析错乱（实测过末尾多出一个 `|` 的情况）。
  写完用 `--pdf --check` 看一眼渲染图。
