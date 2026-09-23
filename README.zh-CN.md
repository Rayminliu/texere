[English](README.md) | 简体中文

# texere —— 从 Markdown 生成可验收的中文 Word 文档

Markdown + Word 模板 → DOCX → 渲染器(Word/WPS/LibreOffice) → PDF → 版式回归 → 证据。

```text
Markdown  +  ref.docx / profile.json        设计契约
        │
        ▼
     texere                                 pandoc → 确定性 OOXML 后处理
        │
        ▼
     DOCX ──► 渲染器 (Word / WPS / LibreOffice) ──► PDF
        │
        ▼
   9 项检查 · 逐页版式回归 · 证据包
```

生成中文标书、正式报告、申报书等正式文档，**不只相信 DOCX 结构本身**——产物要在所选渲染器
（真 Word / WPS / LibreOffice）里打开、导出 PDF，并与版式基线逐页比对。

### 看看产物

**你写的是这个**（`examples/tender/01_bid.md`）：

```markdown
表 1-1 商务条款响应表

| 条款 | 招标要求 | 投标响应 |
|:---|:---|:---|
| 工期 | 90 日历天 | 完全响应 |
| 质量要求 | 符合国家验收标准 | 完全响应 |
| 付款方式 | 按招标文件 | 完全响应 |
```

**甲方打开的是这个**——字一个没变，加上了模板的框线、表头底纹、居中表题、分节页码与目录域：

![tender](assets/previews/tender.png)

同一条管线用在别的文档类型上——[`examples/`](examples/README.md) 每个目录自带 Markdown + config，
一条命令即可跑通；重跑 `python scripts/make_previews.py` 可刷新这些图。

| 公文请示（`gongwen/`） | 项目申报书（`form/`） |
|---|---|
| ![gongwen](assets/previews/gongwen.png) | ![form](assets/previews/form.png) |
| 会议纪要（`minutes/`）——封面承载会议信息 | 经营分析报告（`report/`）——多文件合并、Lead 提示框、grid 表头 |
| ![minutes](assets/previews/minutes.png) | ![report](assets/previews/report.png) |
| 技术服务合同（`contract/`）——条款章节、单元格内换行、签署栏 | 表格排版（`tables/`）——多级表头、合并单元格、列宽控制 |
| ![contract](assets/previews/contract.png) | ![tables](assets/previews/tables.png) |

> **运行平台**：DOCX 生成本身跨平台。PDF / 渲染器验收需要一个渲染器——
> **Word**（Windows + 本机 Microsoft Word）、**LibreOffice**（soffice，任意系统）、**WPS**（Windows + WPS Office）。
> 用 `--renderer word|libreoffice|wps` 选（默认 word）。本机没装任何渲染器时，只有 docx 那一半能跑（见[已知边界](#已知边界)）。

**目录** · [快速开始](#快速开始) · [看看产物](#看看产物) · [工作流](#用法) ·
[验证与证据包](#验证与证据包) · [编辑已有 docx](#编辑已有-docx) · [文档](#文档地图)

## 能力边界

**做**

- Markdown → docx / PDF，含封面、目录域、分节页码、页眉
- 表格排版——框线、表头行、跨页重复表头、斑马纹
- 表题与图注规范
- 交付前一键验收：Word 能否打开、有无空白页、版式有无漂移
- **对已有 docx 做定点编辑**——改文字、增删段落、改单元格、改页眉页脚——且不重排其余部分

**不做**

- 已有 docx 上的批注、修订、脱敏、无障碍处理、水印——那类任务请用通用 docx 技能
- 图表自动编号——编号手写，见[契约](#契约只改版式不改内容)
- 学位论文所需的参考文献、公式编号、奇偶页页眉

更细的边界——无法往返的 OOXML 特性、`.docm`、Word COM 的具体限制——见[已知边界](#已知边界)。

> **范围说明。** 本工具对*中文*正式文档是有立场的：A4、宋体正文、黑体标题、首行缩进 2 字符、全角标点。
> 它不是通用 Markdown → docx 转换器（那是 pandoc 的活）。配 `caption_words` 可以适配非中文题注关键字，
> 但排版默认值始终是中文惯例。

## 快速开始

```powershell
python scripts/render.py --doctor     # 我的环境准备好了吗？
python scripts/render.py --sample     # 冒烟测试，产出 sample_out.docx/.pdf
python scripts/render.py --src examples/tender --out 标书.docx `
    --config examples/tender/config.json --pdf --check
```

第三条命令渲染一份真实投标文件、在所选渲染器（默认 Word）里验收、导出 PDF 并检查产物——**不用先配任何东西**。
只有 `--doctor` 报错时才需要往下看。

**如果 `--doctor` 报错**——`pandoc` 必需；`--pdf` 与验证器需要本机渲染器（默认 Word；可 `--renderer` 选 LibreOffice / WPS）：

| 依赖 | 用途 | 安装 |
|---|---|---|
| pandoc >= 3.1 | **必需**，md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx、lxml | **必需**，docx 读写与 OOXML 处理 | 见下 |
| pywin32 | 仅 `--pdf` 的 Word / WPS 渲染器（COM）；选 LibreOffice 则不需要 | 见下 |
| PyMuPDF | 仅 `--check`（PDF 目视验收） | 见下 |
| Microsoft Word / WPS / LibreOffice | 仅 `--pdf`（三选一，用 `--renderer` 指定） | 系统级，pip / uv 都装不了（LO 走系统包管理器） |

```powershell
winget install --id JohnMacFarlane.Pandoc
pip install -r requirements.txt      # 纯 pip；版本由 uv.lock 导出锁死
                                     # 用 uv 的话：uv sync --all-extras
```

> **本仓库是「脚本集合」，不是可 pip 安装的包**——`pip install .` 会失败（没有构建后端，
> 而且工具要靠相对路径（从各脚本自身位置出发）找到 `assets/ref.docx` / `scripts/` / `assets/sample.md`）。
> 把文件夹放好，**直接跑脚本**即可。

`--doctor` 会**实测三种渲染器的可用性**（默认 Word 走 COM；LibreOffice 查 soffice；WPS 查 KWPS）。不读注册表——注册表里的 `CurVer`
可能是旧 Office 卸载后的残留值，会误报。
贡献者工具（ruff、pre-commit、测试套件）见[开发](#开发)。

## 规模实测

在 Windows + Microsoft Word 渲染器上实测，含 Word 验收与导 PDF：

69 页 / 63 表 / 28 图 → **19.0 秒** 端到端

272 页 / 252 表 / 112 图 → **66.9–69.7 秒** 端到端

两次均产出 2.9 MB 的 docx，且逐项一致（页数、字数、表格、图片、抽查页像素），Word 进程无残留。
可复现示例与基准细节 → [`examples/`](examples/README.md)。

## 为什么不只是 Pandoc？

Pandoc 生成 DOCX，texere 在此之外还做五件事：

1. 把 Word 模板当作版式契约（`ref.docx` / `profile.json`）；
2. 做确定性的 OOXML 后处理——封面、目录域、分节页码、表格样式；
3. 在**所选渲染器（真 Word / WPS / LibreOffice）**里打开产物并导出 PDF——「能被解析」不等于「能被接受」；
4. 检查渲染后的产物——页码连续性、空白页、逐页版式回归；
5. 为结果产出证据：`report.json` + 截图 + 校验清单。

## 已有 Word 模板？继续用它。

把 `reference_doc` 指向甲方给的 `.docx`，版式仍然由它说了算：

| 模板负责 | texere 负责 |
|---|---|
| 页面设置与页边距 | 来自 Markdown 的内容 |
| 正文 / 标题 / 表格 / 题注样式 | 确定性后处理 |
| 页眉（有条件）、表格框线 | 渲染器验收 + PDF 导出 |

封面与分节结构**不继承**——那两部分由 texere 自己生成。完整继承矩阵见
[复用已有模板](#复用已有模板)。

## 用法

下面五是真要跑的流程；每个脚本的完整参数在 [`docs/SCRIPT_HELP.md`](docs/SCRIPT_HELP.md)。

```bash
# 0. 环境
python scripts/render.py --doctor                                # 自检：pandoc / 依赖 / Word 引擎
python scripts/render.py --sample                                # 冒烟测试 -> sample_out.docx/.pdf

# 1. 正式渲染（--check 会自动补 --pdf）
python scripts/render.py --src 章节目录 --out 标书.docx --config cfg.json --pdf --check

# 2. 交付前验收（9 项检查 → report.json + 截图 + 签名）
python scripts/validate.py 标书.docx --out evidence/
python scripts/validate.py 标书.docx --profile profiles/formal-cn-v1.json   # 加版式基线比对

# 3. 编辑已有 docx（定点改动，属编辑链路，契约相反）
python scripts/edit.py 标书.docx --replace "工期=进度" --verify
python scripts/edit.py 标书.docx --fill data.json                # 批量填表

# 4. 声明式 Patch（对 agent 友好：dry-run、哈希前置条件）
python scripts/patch.py 标书.docx patch.json --dry-run
python scripts/patch.py 标书.docx patch.json --apply --validate

# 配套工具
python scripts/distill.py 甲方模板.docx --out cfg.json           # 模板 → 建议 config
python scripts/make_ref.py --body-font 楷体 --body-size 14       # 重建排版模板
python scripts/snapshot.py 标书.pdf --update                     # 录版式基线（确认版式无误后）
python scripts/snapshot.py 标书.pdf                              # 回归比对，漂移即 exit 1
python -m pytest -q                                              # 259 项断言，约 7 分钟（实测 7 分 6 秒，需本机渲染器，默认 Word）
```

可运行示例——每个目录自带 Markdown + config，一条命令跑通（见 [examples/README.md](examples/README.md)）：

```bash
python scripts/render.py --src examples/tables --out examples/tables/tables.docx \
       --config examples/tables/config.json --pdf --check
```

### 输入要求

- `章节目录` 内放 `01_xxx.md … 0N_xxx.md`，**按文件名排序**合并；`#` 为章、`##` 为节；
  一页的通知也可直接给单个文件（`--src notice.md`），不必先建目录；
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

`post.py` 绝不触碰正文与题注的**任何文字**——图表编号由源文件手写。这条保证由
`tests/test_postprocess.py::test_never_touches_text` 守住。渲染管线*可以*改内容，但只有你显式
开 `content_fixes` 替换表时才会；表为空时，输出文字逐字来自源文件。完整契约、`content_fixes`
机制与已移除的自动编号功能 → [`SKILL.md`](SKILL.md)。

## 验证与证据包

```bash
python scripts/validate.py 标书.docx --out evidence/
```

产出 `report.json`、抽样页面截图与校验清单。9 项检查共用**一次** Word 导出。每项检查报
`PASS` / `FAIL` / `SKIP` / `ERROR`：`SKIP` 表示前置条件缺失、这项根本没查，**不计入通过数**——
只有 `FAIL` 与 `ERROR` 会让退出码变成 1。逐项明细、人工门槛与注意事项见
[`docs/VALIDATION.zh-CN.md`](docs/VALIDATION.zh-CN.md)。

## 编辑已有 docx

**独立于渲染链路**的第二条链，契约正好相反：**只改你指定的地方，其余字节原样保留**——
绝不注入封面 / 目录 / 页码，也不重排样式。

```bash
python scripts/edit.py 标书.docx --replace "旧=新" --scope body,tables
python scripts/edit.py 标书.docx --after "锚点文字" --text "新段落"
python scripts/edit.py 标书.docx --cell 0 2 1 "1,060,000"
python scripts/patch.py 标书.docx patch.json --dry-run --apply   # 声明式 Patch，对 agent 友好
```

操作清单、Patch schema 与「重排而非编辑」见 [`docs/EDITING.zh-CN.md`](docs/EDITING.zh-CN.md)。

## 复用已有模板

把 `reference_doc` 指向任意 docx（甲方强制模板，或你自己攒的模板），**能继承的比之前文档里写的多**。
用一份「3cm 页边距 + 楷体 14pt 正文 + 自定义页眉」的模板实测：

| 模板里的东西 | 能继承吗 | 条件 |
|---|---|---|
| 页面设置（纸型 / 方向 / 页边距） | ✅ | 自动 |
| 正文 / 标题 / 表格 / 题注样式（字体 / 字号 / 颜色 / 间距） | ✅ | 自动 |
| **页眉** | ✅ | **config 里不要写 `header`** |
| 页脚 / 页码 | ⚠️ | 总被重写，见下 |
| 封面 | ❌ | `post.py` 自己注入 |
| 分节结构 | ❌ | 重建为「封面+目录 / 正文」两节 |

```json
{
  "reference_doc": "甲方模板.docx",
  "style": { "page_number": null }
}
```

**页脚**：默认行为是把页码**追加**到模板原有内容后面（`文件编号 XYZ-2026— 1 —`）。
想让模板页脚一个字都不变，就写 `"page_number": null`。

**用之前先蒸馏一遍**：

```bash
python scripts/distill.py 甲方模板.docx                  # 报告 + 建议 config
python scripts/distill.py 甲方模板.docx --out cfg.json    # 同时写出 config
```

报告会给出页面设置、正文与各标题的字体字号、各节页眉页脚，以及配套的 `make_ref.py` 命令
（如果你想照这套字体重建自己的模板，而不是引用甲方的文件）。封面结构、页码格式、表格样式判断不了，
报告里会列成人工确认项。

> **是蒸馏，不是转换。** 脚本不去"理解"模板，它把模板拆成 config 字段，决定权在你——
> 这跟整个工具的分工方式一致。

## 配置与表格

参考类内容下沉到 docs，本手册才能保持「产品首页」的定位。默认值已经够跑通
`render.py --sample` 和所有示例，需要查某个键时再翻参考：

| 要查什么 | 在哪 |
|---|---|
| `config.json` 全部字段、`style` 全部键 | [`docs/CONFIG.zh-CN.md`](docs/CONFIG.zh-CN.md) |
| 表格写法、grid table、视觉控制 | [`docs/TABLES.zh-CN.md`](docs/TABLES.zh-CN.md) |

## 标书场景注意

1. **招标方给了强制格式模板时，以对方为准**——能继承哪些见[复用已有模板](#复用已有模板)。
   那一节替不了你的部分：封面、密封、签字页、页码规则仍要对着招标文件人工核对，
   本工具不替代合规审查。
2. 标书常见结构（投标函 / 商务标 / 技术标 / 报价 / 资质业绩）直接对应 `#` 章即可；
   报价表建议保留 pipe table，便于后期整体替换为招标方表格；复杂表头用 grid table。
3. 交付前跑 `--pdf --check` 并肉眼过一遍渲染图：Word 能打开、无空白页、表头灰底、题注居中，
   四条都过再封包。

## 三条支柱

1. **是编译器，不是转换器** —— Markdown → 语义 IR → 版式规格 → DOCX。每条视觉规则都明写在设计契约里
   （`ref.docx` + `profile.json`），没有魔法。
2. **是契约，不是猜测** —— profile 声明字体、间距、框线、页眉、页脚、题注。agent 读得懂，人查得清，
   没有隐藏假设。
3. **是证据，不是希望** —— 一条命令验完：包完整性、图片嵌入、目录域、页码、空白页、Word 验收、版式漂移。
   产出 `report.json` + 页面截图 + 校验清单（见[证据包](#验证与证据包)）。

## 设计原则

改模板或手写内容时遵守：

1. 样式表驱动：视觉规则只写在 `assets/ref.docx`，源文本只写语义；
2. 中文四件套：宋体正文 + 黑体标题 + eastAsia 字体属性 + 首行缩进 2 字符；
3. 表格三件套：100% 宽 + 灰底加粗居中表头 + 跨页重复表头；
4. 题注居中、小一号、灰色、**去斜体**；
5. 中西混排：中文宋体、数字西文 Times New Roman；
6. 交付前真 Word 打开 + 渲染目视，不信任何"读回正常"。

## 已知边界

- 目录为 Word 域，首次打开若未刷新请全选按 F9（已设 `updateFields`，通常自动）；
- 图表编号为**手写**：插删图表后需要人工对号（见[契约](#契约只改版式不改内容)）；
- **正文与标题的字体 / 字号在模板层**，改它是 `python scripts/make_ref.py --body-font 楷体 --body-size 14`
  （`style` 段的字体只管表格与题注，管不到正文）；
- **grid table 对空格敏感**：每行竖线必须严格对齐，否则会解析错乱（实测过末尾多出一个 `|`）。
  注意 pandoc 按**显示宽度**对齐——中文占 2 列，「等宽编辑器里看着齐」仍可能解析成单列坏表；
  改完务必渲染验证，或用脚本按显示宽度对齐；
- **渲染器绑定，而非 Word 绑定**：见顶部「运行平台」。解析与排版逻辑不依赖任何 Office，PDF 导出和一半验收链条走你选的渲染器（`--renderer word|libreoffice|wps`）。`LibreOffice` 是便携选项，`Word`/`WPS` 是中文 Office 保真选项；
- `--check` 会把「内容稀疏但合法的页面」也算作近空白页：表单尾页的签字盖章区、
  大表格前的单独标题页等。这类是误报，用 `--max-empty N` 放宽；
- 快照基线依赖本机 Word 版本与字体，**换机器后先 `snapshot.py --update` 重录**，
  否则会满屏漂移；阈值默认 0.1%（实测同文档重复导出为 0.00%，改一处页眉为 0.16%）；
- 不使用 Quarto：其 1.10.x 的 docx 对带自动编号题注的表格会丢失表体。

## 开发

```bash
pip install ruff pre-commit && pre-commit install      # ruff 一个工具顶 flake8 + black + isort
pre-commit run --all-files                             # lint + 格式 + 不启 Word 的测试子集
python -m pytest -q                                    # 全量：259 项断言，约 7 分钟（实测 7 分 6 秒）
```

> **没有托管 CI。** 验收类用例要通过 COM 驱动一台真 Microsoft Word，托管 runner 给不了——
> 本地靠 pre-commit 覆盖无 Word 的那部分，全量套件推送前手动跑。

## 文档地图

每条信息只住在**一个**地方，其余文档只指路不复述（能机器查的那部分已由 `tests/test_docs_sync.py` 把守）：

| 文件 | 负责什么 |
|---|---|
| `README.zh-CN.md` / `README.md` | 产品首页：是什么、快速开始、工作流、边界——不做参考手册 |
| `SKILL.md` | Agent 入口：何时用/不用、硬契约、验收门槛、Patch schema、陷阱清单 |
| `BEST_PRACTICES.md` | 场景经验：profile 选择、表格配方、调试案例、FAQ |
| `docs/SCRIPT_HELP.md` | 各脚本 CLI 参考（参数的**唯一**来源，与代码双向对拍把守） |
| `docs/CONFIG.zh-CN.md` / `docs/CONFIG.md` | `config.json` 全部字段与 `style` 全部键 |
| `docs/TABLES.zh-CN.md` / `docs/TABLES.md` | 表格写法、grid table、视觉控制 |
| `docs/VALIDATION.zh-CN.md` / `docs/VALIDATION.md` | 9 项检查逐项说明、人工门槛、注意事项 |
| `docs/EDITING.zh-CN.md` / `docs/EDITING.md` | 编辑操作、Patch schema、重排流程 |

所以：**本手册不列 CLI 参数表**——只给你真要跑的那几条命令，其余指向 SCRIPT_HELP。
脚本新增了参数却没写进 SCRIPT_HELP，或者 SCRIPT_HELP 编了个不存在的参数，`test_docs_sync.py`
都会在 commit 时拦下。

**仓库结构**

| 路径 | 职责 |
|---|---|
| `scripts/render.py` | 入口：合并 md → pandoc → post.py →（可选）finalize / check |
| `scripts/post.py` | 后处理：封面注入、目录域、分节页码、页眉页脚、表格规则、题注样式；手写 OOXML 按 ECMA-376 顺序插入 |
| `scripts/filters/captions.lua` | pandoc Lua filter：在 **AST 层**把表题 / 图注标记成 `TableCaption` / `FigureCaption`，`post.py` 不必用正则猜 |
| `scripts/finalize.py` | Word COM：打开验收（打不开 = 结构错）、刷目录域、导 PDF、存回 |
| `scripts/check_pdf.py` | PyMuPDF：空白页检测（超阈值 exit 1）+ 渲染页面 PNG 供目视 |
| `scripts/validate.py` | 统一验证入口：9 项自动检查、结构化报告、证据包 |
| `scripts/snapshot.py` | PDF 版式快照回归：逐像素比对 `baselines/`，漂移即 exit 1 |
| `scripts/make_previews.py` | 重生示例画廊图（逐个渲染示例，挑选代表页） |
| `scripts/patch.py` | 对 agent 友好的 Patch API：声明式操作、dry-run、哈希前置条件、评估 |
| `scripts/edit.py` | 编辑已有 docx：改文字 / 增删段落 / 改单元格 / 改页眉页脚 |
| `scripts/distill.py` | 蒸馏模板 docx，输出建议 config（页面设置 / 字体 / 页眉页脚） |
| `scripts/make_ref.py` | 重新生成 `assets/ref.docx`（改字体 / 字号 / 间距时用它） |
| `profiles/*.json` | 设计契约：`formal-cn-v1`（通用正式）、`gongwen-v1`、`tender-v1`、`application-v1` |
| `assets/ref.docx` | 中文排版模板：宋体小四正文、黑体标题阶梯、表格边框、题注样式、封面样式（CoverTop/…）、提示框样式（Lead/SmallNote） |
| `assets/sample.md` / `assets/sample_config.json` | 钉死的回归夹具，不是展示样例：`--sample`、四个测试文件与 `baselines/` 都只渲染这一份文档。展示请看 `examples/` |
| `examples/` | 可运行示例：投标文件、公文请示、项目申报书、会议纪要、经营分析报告、技术服务合同、表格排版（见 `examples/README.md`） |
| `docs/` | `SCRIPT_HELP.md`（CLI）、`CONFIG.md`（config 字段）、`TABLES.md`（表格）、`VALIDATION.md`（9 项检查）、`EDITING.md`（编辑与 Patch）——各有 `.zh-CN` 镜像 |
| `baselines/` | 快照基线（样例 4 页 PNG） |
| `tests/` | 259 项 pytest 断言：排版规则、题注识别、表格特性、退出码、快照逻辑、只改版式契约、跨 run 编辑（`test_edit.py`）、模板复用与蒸馏（`test_distill.py`）、9 项验收器（`test_validate.py`）、Patch API（`test_patch.py`）、版本一致性与页码/基线纯函数（`test_version.py` / `test_validate_units.py`）、Renderer 抽象护栏（`test_renderer.py`：适配器契约 + PDF 派生检查 renderer-agnostic；`test_renderer_contract.py`：抽象契约 + 超时/并发集成）、文档与代码同步守卫（`test_docs_sync.py`：CLI 参数 ↔ SCRIPT_HELP 双向对拍、单一来源、中英镜像结构与脚本覆盖、锚点有效性、SKILL.md front matter 合法性、断言数 ↔ 实际收集数） |
| `CHANGELOG.md` | 版本历史与每条修复的理由 |

## 许可

MIT，见 [LICENSE](LICENSE)。
