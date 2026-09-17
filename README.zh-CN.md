[English](README.md) | 简体中文

# docx-kit —— 把 Markdown 渲染成中文正式文档

把 Markdown 章节目录变成**排版合格的中文正式 docx / PDF**：标书、计划书、申报书、结题报告、白皮书。

三条硬特色：

1. **样式表驱动** —— 全部视觉规则写在 `ref.docx`，源 Markdown 只写语义。改版式不用碰一个字的内容。
2. **真 Word 验收** —— 用本机 Word 打开、刷新目录域、导出 PDF，再肉眼过渲染图。不信任何"读回正常"。
3. **只改版式，不改内容** —— 输出里每一个字都来自源文件。这条有测试守着。

## 能做什么，不做什么

| | |
|---|---|
| **做** | Markdown 目录 → docx / PDF；封面、目录域、分节页码、页眉；表格排版（框线 / 表头 / 跨页重复 / 斑马纹）；表题与图注规范；交付前的一键验收（Word 能否打开、有无空白页、版式有无漂移） |
| **不做** | **编辑已有 docx**（批注、修订、脱敏、无障碍、水印）——那类任务请用通用 docx 技能；图表自动编号（编号手写，见「契约」）；学位论文所需的参考文献、公式编号、奇偶页页眉 |

## 快速开始

外部依赖：`pandoc` 必需；`--pdf` 需要本机 Word。

| 依赖 | 用途 | 安装 |
|---|---|---|
| pandoc >= 3.1 | **必需**，md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx、lxml | **必需**，docx 读写与 OOXML 处理 | 见下 |
| pywin32 | 仅 `--pdf`（Word 验收 + 导 PDF） | 见下 |
| PyMuPDF | 仅 `--check`（PDF 目视验收） | 见下 |
| Microsoft Word | 仅 `--pdf` | 系统级，pip / uv 都装不了 |

> **本仓库是「脚本集合」，不是可 pip 安装的包**——`pip install .` 会失败（没有构建后端，
> 而且工具要靠相对路径找到 `ref.docx` / `filters/` / `sample.md`）。
> 把文件夹放好，**直接跑脚本**即可。

```powershell
winget install --id JohnMacFarlane.Pandoc

# 依赖二选一：
uv sync --all-extras                 # 用 uv（推荐，版本锁在 uv.lock）
pip install -r requirements.txt      # 纯 pip（uv export 生成的精确版本）

python render.py --doctor            # 环境自检
python render.py --sample            # 冒烟测试，产出 sample_out.docx/.pdf
```

`--doctor` 会**实测 Word 引擎身份**（走 COM 直接问）。不读注册表——注册表里的 `CurVer`
可能是旧 Office 卸载后的残留值，会误报。

## 效果

`python render.py --sample` 的产物（4 页投标文件风格样例）：

![样例封面](baselines/p001.png)
![样例正文](baselines/p003.png)

**规模实测**（含 Word 验收 + 导 PDF + 空白页检查）：

| 规模 | 耗时 |
|---|---|
| 69 页 / 63 表 / 28 图 | **19.0 秒** |
| 272 页 / 252 表 / 112 图 | **69.7 / 66.9 秒**（连跑两次） |

大文档那次产物 docx 2.9 MB，两次运行结果逐项一致（页数、字数、表格、图片及抽查页像素全部相同），
Word 进程无残留。

## 用法

```bash
python render.py --doctor                                    # 环境自检
python render.py --version                                   # 版本
python render.py --sample                                    # 冒烟测试

# 正式渲染（--check 会自动补 --pdf）
python render.py --src 章节目录 --out 标书.docx --config cfg.json --pdf --check

python -m pytest -q                        # 36 项断言，不需要 Word，约 9 秒
python snapshot.py 标书.pdf --update       # 录版式基线（确认版式无误后执行）
python snapshot.py 标书.pdf                # 回归比对，漂移即 exit 1
python make_ref.py --body-font 楷体        # 重建排版模板（改字体/字号时）
```

### 输入要求

- `章节目录` 内放 `01_xxx.md … 0N_xxx.md`，**按文件名排序**合并；`#` 为章、`##` 为节；
- 表题写成独立一行 `表 1-1 标题`（自动居中灰字）；图用 `![图 1-1 标题](a.jpg){width=13cm}`；
- **图片搜索路径**：`src` 目录、其全部子目录、**`src` 的父目录及其子目录**都会自动加入；
  还找不到就用 config 的 `resource_paths` 补充。渲染结束会打印 `images: n/m ok`，
  源里有图却没嵌进去时直接报 `[ERROR]`；
- 封面与页眉在 config 里配；不配 `cover` 则只注入目录。
  **注意**：注入封面后 `post.py` 会自动补一个空 `CoverInfo` 段，所以 config 里不要再写
  行首 / 行尾空行，否则封面容易溢出成两页（多出一张空白页）；
- 强调用 `**加粗**`，灰底提示框用 `::: {custom-style="Lead"} … :::`；
- **第一个 `#` 之前别放内容**（比如标题性的短语）：它既不会被当成封面，还会排在目录之后。
  真放了 `post.py` 会打印 `[warn]` 列出来，但不会替你删（契约是不改内容）；
- config.json / 源 md 带 UTF-8 BOM 也能正常读（Windows 记事本默认写 BOM）；
- **表单 / 附件类文档**（没有 `#` 标题）也能处理：不插目录、不分节，只做封面与表格 / 题注排版。
  这类表格首行通常是字段名而非列标题，配 `"style": {"header_rows": 0}` 以免被灰底加粗。

### 契约：只改版式，不改内容

`post.py` 不触碰正文与题注的**任何文字**。图表编号由源文件手写。

0.2.0 期间曾有过「图表自动编号 + `@tab:` 交叉引用」功能，因会改写题注文字、给无编号题注
补号、且不同步手写引用，已**整体移除**。若将来要重做编号，正确做法是插入 Word 原生
`SEQ` / `REF` 域（域可更新、不改文本），而不是重写题注文字。

### 验证门槛

正式渲染之后逐项确认，四项全过才算完成：

1. 日志出现 `images: n/m ok` 且 **n == m**（m 是源 md 中 `![` 的次数）
2. `near-empty pages: 0`
3. `OK`（Word 成功打开并导出 PDF）
4. **首次用于某份新文档时，人工过一遍 PDF 渲染图**——机器能查页数、图片数、空白页，
   查不了"这张图画得对不对、表头有没有被截断"

## 配置

### config.json 字段

| 字段 | 作用 |
|---|---|
| `cover` | 封面行 `[[样式名, 文本], …]`，样式名用 CoverTop / CoverTitle / CoverSub / CoverInfo / CoverDate |
| `header` | 正文节页眉 |
| `title` / `author` / `subject` / `comments` | 文档属性 |
| `reference_doc` | 指定模板 docx（招标方给强制格式时用它） |
| `toc_heading` | 目录标题，默认「目　　录」 |
| `style` | 版式微调，见下表 |
| `caption_words` | 自定义题注关键字（默认 表/图/Table/Figure），见下 |
| `resource_paths` | 额外的图片搜索目录列表（默认已含 src、其子目录与父目录） |
| `content_fixes` | 编辑性替换表 `[[旧文本, 新文本], …]`，合并 md 后、转换前套用 |
| `content_fixes_file` | 替换表文件（`.json`，或 `.py` 里的 `CONTENT_FIXES` 字面量；相对路径按 config 所在目录解析） |

> `content_fixes` 用来删掉注释性括号、统一措辞。指向上游已有的 `.py` 时，是用 `ast`
> **只读取值、不执行代码**，因此不必把规则复制成第二份。

`caption_words`（需要非中文或不惯用叫法时才配，`post.py` 与 lua filter 会同步）：

```json
"caption_words": {"table": ["表", "表格"], "figure": ["图", "图片"]}
```

### `style` 段

全部可省，缺省即中文正式文档惯例：

| 分组 | 键 | 默认 |
|---|---|---|
| 页脚 | `page_number` / `page_number_size` | `— {n} —` / `9` |
| 页眉 | `header_size` / `header_gray` / `header_rule_color` / `header_rule_size` | `9` / `595959` / `BFBFBF` / `4` |
| 目录 | `toc_depth` / `toc_title_size` / `toc_title_color` / `toc_placeholder` / `toc_placeholder_size` | `1-2` / `16` / `000000` / 见提示语 / `12` |
| 题注 | `caption_gray` / `caption_size` / `caption_space_before` / `caption_space_after` | `404040` / `10.5` / `6` / `4` |
| 题注 | `caption_keep_with_next` | `true`（表题不与表格分家；实测关掉可省 1 页，但表题可能落在页尾） |
| 表格 | `header_rows` / `table_border` / `table_shade` / `table_size` | 自动 / `full` / `EDEDED` / `10.5` |
| 表格 | `cell_margin_v` / `cell_margin_h` / `table_para_space` | `40` / `80` / `1` |
| 表格 | `border_size` / `border_color` / `three_line_size` | `6` / `808080` / `12` |
| 表格 | `table_header_color` / `table_zebra` / `table_zebra_fill` | 不指定 / `false` / `F7F7F7` |
| 字体 | `east_font` / `latin_font` | `宋体` / `Times New Roman`（**只作用于表格与题注**） |

> 颜色写 6 位十六进制（`404040`），粗细单位 1/8 pt，间距单位 pt，单元格边距单位 twips。

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

**2. 控制列宽**——pandoc 按**最细的列分隔**分配列宽。上表最细一行是 `+--------+---------+`
（8/9/8/9），四列实际宽度就是 8:9:8:9（实测 990/1100/990/1100 twips）。想让「序号」列窄，把它写窄就行。

**3. 单元格内换行**——同一个单元格里写多行即可（pipe table 做不到）。

### 表格视觉控制

| 键 | 默认 | 说明 |
|---|---|---|
| `header_rows` | 自动 | 前 N 行做**视觉**表头（灰底加粗居中）。默认**按表自动识别**（读 pandoc 打的 `w:tblHeader`）；想强制才填数字；**填 `0` = 该表没有表头**（表格首行不是列标题，如表单 / 附件类），此时连「跨页重复表头」也一并去掉 |
| `table_border` | `full` | `full` 全框线 / `three` 三线表 / `none` 无框线 |
| `table_shade` | `EDEDED` | 表头底纹 |
| `table_header_color` | 不指定 | 表头文字颜色；**深色底时配 `FFFFFF` 白字** |
| `table_zebra` / `table_zebra_fill` | `false` / `F7F7F7` | 表体隔行浅底（斑马纹），首条数据行保持白底 |
| `table_size` | `10.5` | 表格字号 pt |

> 跨页重复表头**不用配**：pandoc 原生就给表头行设了 `w:tblHeader`（grid table 的 `+===+`
> 以上全部算表头），`post.py` 里那行只是幂等加固。

## 标书场景三条注意

1. **招标方给了强制格式模板时，以对方为准**：把它的 docx 路径写进 config 的 `reference_doc`，
   正文样式即继承对方模板；封面 / 密封 / 签字页 / 页码规则仍按招标文件手工核对，
   本工具不替代合规审查。
2. 标书常见结构（投标函 / 商务标 / 技术标 / 报价 / 资质业绩）直接对应 `#` 章即可；
   报价表建议保留 pipe table，便于后期整体替换为招标方表格；复杂表头用 grid table。
3. 交付前跑 `--pdf --check` 并肉眼过一遍渲染图：Word 能打开、无空白页、表头灰底、题注居中，
   四条都过再封包。

## 设计原则（改模板或手写内容时遵守）

1. 样式表驱动：视觉规则只写在 `ref.docx`，源文本只写语义；
2. 中文四件套：宋体正文 + 黑体标题 + eastAsia 字体属性 + 首行缩进 2 字符；
3. 表格三件套：100% 宽 + 灰底加粗居中表头 + 跨页重复表头；
4. 题注居中、小一号、灰色、**去斜体**；
5. 中西混排：中文宋体、数字西文 Times New Roman；
6. 交付前真 Word 打开 + 渲染目视，不信任何"读回正常"。

## 已知边界

- 目录为 Word 域，首次打开若未刷新请全选按 F9（已设 `updateFields`，通常自动）；
- 图表编号为**手写**：插删图表后需要人工对号（见「契约」）；
- **正文与标题的字体 / 字号在模板层**，改它是 `python make_ref.py --body-font 楷体 --body-size 14`
  （`style` 段的字体只管表格与题注，管不到正文）；
- **grid table 对空格敏感**：每行竖线必须严格对齐，否则会解析错乱（实测过末尾多出一个 `|`）；
- **Windows + Word 绑定**：`--pdf` 需要本机 Word（COM）。Linux / macOS 只能出 docx
  （核心解析与排版逻辑不依赖 Word，但验收链条缺一半）；
- `--check` 会把「内容稀疏但合法的页面」也算作近空白页：表单尾页的签字盖章区、
  大表格前的单独标题页等。这类是误报，用 `--max-empty N` 放宽；
- 快照基线依赖本机 Word 版本与字体，**换机器后先 `snapshot.py --update` 重录**，
  否则会满屏漂移；阈值默认 0.1%（实测同文档重复导出为 0.00%，改一处页眉为 0.16%）；
- 不使用 Quarto：其 1.10.x 的 docx 对带自动编号题注的表格会丢失表体。

## 文件地图

| 文件 | 职责 |
|---|---|
| `render.py` | 入口：合并 md → pandoc → post.py →（可选）finalize / check |
| `post.py` | 后处理：封面注入、目录域、分节页码、页眉页脚、表格规则、题注样式；手写 OOXML 按 ECMA-376 顺序插入 |
| `filters/captions.lua` | pandoc Lua filter：在 **AST 层**把表题 / 图注标记成 `TableCaption` / `FigureCaption`，`post.py` 不必用正则猜 |
| `ref.docx` | 中文排版模板：宋体小四正文、黑体标题阶梯、表格边框、题注样式、封面样式（CoverTop/…）、提示框样式（Lead/SmallNote） |
| `make_ref.py` | 重新生成 `ref.docx`（改字体 / 字号 / 间距时用它） |
| `finalize.py` | Word COM：打开验收（打不开 = 结构错）、刷目录域、导 PDF、存回 |
| `check_pdf.py` | PyMuPDF：空白页检测（超阈值 exit 1）+ 渲染页面 PNG 供目视 |
| `snapshot.py` | PDF 版式快照回归：逐像素比对 `baselines/`，漂移即 exit 1 |
| `tests/` | 36 项 pytest 断言：排版规则、题注识别、表格特性、退出码、快照逻辑、只改版式契约 |
| `sample.md` / `sample_config.json` | 冒烟测试样例（投标文件风格） |
| `baselines/` | 快照基线（样例 4 页 PNG） |
| `SKILL.md` | 给其他 agent 用的技能说明（何时用、验证门槛、硬契约） |
| `CHANGELOG.md` | 版本历史与每条修复的理由 |

## 许可

MIT，见 [LICENSE](LICENSE)。
