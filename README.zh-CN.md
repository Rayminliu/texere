[English](README.md) | 简体中文

# texere —— 中文正式文档编译器

> 名字取自拉丁语 **texere**——"编织"，*text*（文本）与 *textile*（织物）的共同词源。
> 排版做的正是这件事：把正文、表格、题注、页码编织成一页有序的版面。

**Reference/Spec → Deterministic Document → Evidence**

texere 是面向**中文正式文档**（标书、申报书、结题报告、白皮书）的高保真文档编译器：给它 Markdown 源文件
和一份参考模板（`ref.docx` 或 `profile.json`），它产出一个**经真 Word 验收过**、导出 PDF、并做过版式漂移
比对的 Word 文档——附带一个证据包，证明它达到了可交付标准。

> **运行平台**：Windows + 本机 Microsoft Word。docx 之后的每一步——Word 验收、PDF 导出、9 项验证器——
> 都走 Word COM。Linux / macOS 只有 docx 那一半能跑（见[已知边界](#已知边界)）。

**目录** · [三条支柱](#三条支柱) · [能力边界](#能力边界) · [文档地图](#文档地图) ·
[快速开始](#快速开始) · [效果](#效果) · [用法](#用法) · [验证与证据包](#验证与证据包) ·
[编辑已有 docx](#编辑已有-docx) · [复用已有模板](#复用已有模板) · [配置](#配置) · [表格](#表格) ·
[标书场景注意](#标书场景注意) · [设计原则](#设计原则) · [已知边界](#已知边界) ·
[开发](#开发) · [文件地图](#文件地图)

## 三条支柱

1. **是编译器，不是转换器** —— Markdown → 语义 IR → 版式规格 → DOCX。每条视觉规则都明写在设计契约里
   （`ref.docx` + `profile.json`），没有魔法。
2. **是契约，不是猜测** —— profile 声明字体、间距、框线、页眉、页脚、题注。agent 读得懂，人查得清，
   没有隐藏假设。
3. **是证据，不是希望** —— 一条命令验完：包完整性、图片嵌入、目录域、页码、空白页、Word 验收、版式漂移。
   产出 `report.json` + 页面截图 + 签名。

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

## 文档地图

每条信息只住在**一个**地方，其余文档只指路不复述（信息漂移让我们吃过亏，
能机器查的那部分已由 `tests/test_docs_sync.py` 把守）：

| 文件 | 负责什么 |
|---|---|
| `README.zh-CN.md` / `README.md` | 面向人的手册：config 字段、`style` 键、表格语法、输入规则、已知限制 |
| `SKILL.md` | Agent 入口：何时用/不用、硬契约、验收门槛、Patch schema、陷阱清单 |
| `BEST_PRACTICES.md` | 场景经验：profile 选择、表格配方、调试案例、FAQ |
| `docs/SCRIPT_HELP.md` | 各脚本 CLI 参考（参数的**唯一**来源，与代码双向对拍把守） |

所以：**本手册不列 CLI 参数表**——只给你真要跑的那几条命令，其余指向 SCRIPT_HELP。
脚本新增了参数却没写进 SCRIPT_HELP，或者 SCRIPT_HELP 编了个不存在的参数，`test_docs_sync.py`
都会在 commit 时拦下。

> **范围说明。** 本工具对*中文*正式文档是有立场的：A4、宋体正文、黑体标题、首行缩进 2 字符、全角标点。
> 它不是通用 Markdown → docx 转换器（那是 pandoc 的活）。配 `caption_words` 可以适配非中文题注关键字，
> 但排版默认值始终是中文惯例。

## 快速开始

外部依赖：`pandoc` 必需；`--pdf` 与验证器需要本机 Word。

| 依赖 | 用途 | 安装 |
|---|---|---|
| pandoc >= 3.1 | **必需**，md → docx | `winget install --id JohnMacFarlane.Pandoc` |
| python-docx、lxml | **必需**，docx 读写与 OOXML 处理 | 见下 |
| pywin32 | 仅 `--pdf`（Word 验收 + 导 PDF） | 见下 |
| PyMuPDF | 仅 `--check`（PDF 目视验收） | 见下 |
| Microsoft Word | 仅 `--pdf` | 系统级，pip / uv 都装不了 |

> **本仓库是「脚本集合」，不是可 pip 安装的包**——`pip install .` 会失败（没有构建后端，
> 而且工具要靠相对路径（从各脚本自身位置出发）找到 `assets/ref.docx` / `scripts/` / `assets/sample.md`）。
> 把文件夹放好，**直接跑脚本**即可。

```powershell
winget install --id JohnMacFarlane.Pandoc

# 依赖（一条命令装齐；版本由 uv.lock 导出锁死）
pip install -r requirements.txt      # 纯 pip
                                     # 用 uv 的话：uv sync --all-extras

python scripts/render.py --doctor            # 环境自检
python scripts/render.py --sample            # 冒烟测试，产出 sample_out.docx/.pdf
```

`--doctor` 会**实测 Word 引擎身份**（走 COM 直接问）。不读注册表——注册表里的 `CurVer`
可能是旧 Office 卸载后的残留值，会误报。
贡献者工具（ruff、pre-commit、测试套件）见[开发](#开发)。

## 效果

`python scripts/render.py --sample` 的产物（4 页投标文件风格样例）：

![样例封面](baselines/p001.png)
![样例正文](baselines/p003.png)

### 示例画廊

以下均为 [`examples/`](examples/README.md) 的真实渲染结果——每个目录自带 Markdown + config，
一条命令即可跑通；重跑 `python scripts/make_previews.py` 可刷新这些图。

| 投标文件（`tender/`） | 公文请示（`gongwen/`） |
|---|---|
| ![tender](assets/previews/tender.png) | ![gongwen](assets/previews/gongwen.png) |
| 项目申报书（`form/`）——无标题无目录，首行是字段名 | 会议纪要（`minutes/`）——封面承载会议信息 |
| ![form](assets/previews/form.png) | ![minutes](assets/previews/minutes.png) |
| 经营分析报告（`report/`）——多文件合并、Lead 提示框、grid 表头 | 技术服务合同（`contract/`）——条款章节、单元格内换行、签署栏 |
| ![report](assets/previews/report.png) | ![contract](assets/previews/contract.png) |
| 表格排版（`tables/`）——多级表头、合并单元格、列宽控制 | |
| ![tables](assets/previews/tables.png) | |

**规模实测**（含 Word 验收 + 导 PDF + 空白页检查）：

| 规模 | 耗时 |
|---|---|
| 69 页 / 63 表 / 28 图 | **19.0 秒** |
| 272 页 / 252 表 / 112 图 | **69.7 / 66.9 秒**（连跑两次） |

大文档那次产物 docx 2.9 MB，两次运行结果逐项一致（页数、字数、表格、图片及抽查页像素全部相同），
Word 进程无残留。

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
python -m pytest -q                                              # 162 项断言，约 6-7 分钟（需本机 Word）
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

`post.py` 不触碰正文与题注的**任何文字**。图表编号由源文件手写。

0.2.0 期间曾有过「图表自动编号 + `@tab:` 交叉引用」功能，因会改写题注文字、给无编号题注
补号、且不同步手写引用，已**整体移除**。若将来要重做编号，正确做法是插入 Word 原生
`SEQ` / `REF` 域（域可更新、不改文本），而不是重写题注文字。

## 验证与证据包

渲染之后，一条命令把"我看着没问题"变成可审计的产物：

```bash
python scripts/validate.py 标书.docx --out evidence/
```

产出：

```
evidence/
├── report.json          # 结构化验证报告（9 项）
├── page-001.png         # 抽样页面截图（首 / 中 / 尾）
├── page-069.png
├── page-272.png
└── signature            # SHA256 哈希 + 检查摘要
```

报告含 9 项自动检查：

| 检查项 | 验的是什么 |
|---|---|
| ✅ 包完整性 | docx 是含必需部件的合法 ZIP |
| ✅ 源内容完整性 | 可选：与预期哈希比对 |
| ✅ 图片嵌入 | 源里引用的图片全部嵌入（n/m ok） |
| ✅ 分节数 | 分节数量合理（1–100） |
| ✅ 目录域 | 目录存在且可更新 |
| ✅ 页码 | 页码连续无缺口（只看页脚区域） |
| ✅ 空白页 | 不超阈值（默认允许 0 个） |
| ✅ Word 验收 | 真 Word 能打开并成功导出 PDF |
| ✅ 版式基线漂移 | 与基线逐像素比对（需先提供基线） |

示例输出：

```
Document Validation
────────────────────────────
✅ [PASS] package_integrity: OK
✅ [PASS] source_content: Skip (未提供 expected_hash)
✅ [PASS] image_embedding: 图片嵌入：28/28 ok
✅ [PASS] section_count: 分节数：3 (合理)
✅ [PASS] toc_field: 目录域：存在 (Table of Contents 1)
✅ [PASS] page_numbering: 页码：69 页 (连续)
✅ [PASS] blank_pages: 空白页：0/69 (阈值：0)
✅ [PASS] word_acceptance: Word 验收：OK
✅ [PASS] visual_drift: 视觉基线：一致

Summary
────────────────────────────
Passed: 9/9

✅ All checks passed

Evidence package saved to: evidence/
  - report.json (structured validation report)
  - page-XXX.png (sample screenshots)
  - signature (SHA256 signature)
```

任何一项失败即 exit code 1，并给出详细错误。参数（`--profile` / `--max-empty` / `--quiet` 等）见
`docs/SCRIPT_HELP.md` §validate.py。

### 人工门槛

9 项检查管的是结构，头一遍过某份新文档时，这四件事仍然得靠人：

1. 日志出现 `images: n/m ok` 且 **n == m**（m 是源 md 中 `![` 的次数）
2. `near-empty pages: 0`
3. `OK`（Word 成功打开并导出 PDF）
4. **人工过一遍 PDF 渲染图**——机器能查页数、图片数、空白页，
   查不了"这张图画得对不对、表头有没有被截断"

## 编辑已有 docx

`scripts/edit.py` 是**独立于渲染链路**的第二条链，契约正好相反：

| | 渲染链路 | 编辑链路 |
|---|---|---|
| 输入 | Markdown | 已有 docx |
| 契约 | 只改版式，不改内容 | **只改你指定的地方，其余字节原样保留** |
| 禁止 | 改写文字 | 注入封面 / 目录 / 页码 / 重排样式 |

操作长这样（完整参数：`docs/SCRIPT_HELP.md` §edit.py）：

```bash
python scripts/edit.py 标书.docx --list                      # 看结构：分节 / 表格 / 段落
python scripts/edit.py 标书.docx --replace "旧=新" --scope body,tables,header,footer
python scripts/edit.py 标书.docx --after "锚点文字" --text "新段落"   # 另有 --before / --delete
python scripts/edit.py 标书.docx --cell 0 2 1 "1,060,000"            # 表号 / 行 / 列 / 值
python scripts/edit.py 标书.docx --add-rows 0 3 --template-row 2     # 另有 --add-row / --del-row
python scripts/edit.py 标书.docx --fill data.json                    # 批量填表，JSON 或 CSV
python scripts/edit.py 标书.docx --header "新版页眉" --footer "— X —" --section all
python scripts/edit.py 标书.docx --replace "A=B" --verify            # 改完让 Word 打开一次验收
```

默认先把原文件备份成 `<name>.bak.docx`（`--no-backup` 关闭，`--out` 另存）。

**为什么可以放心原地改**：实测 `python-docx` 打开再保存，部件**零丢失、零新增**（4 页文档实测；
文件变小只是重新压缩）。

**两个必须知道的坑**

1. **Word 会把文字切成多个 run。** `自开标之日起 90 日历天` 可能是三个 run，数字单独一个，
   直接遍历 `run.text` 根本找不到。`edit.py` 的做法是：拼整段文本定位，把替换内容写进
   「命中起点所在的那个 run」，其余 run 只删掉被覆盖的字符——run 数和各 run 的格式都保住。
2. **锚点命中多处时拒绝执行。** 目录被 Word 刷成静态文本后，`1.2 资质与业绩` 这类标题在目录和
   正文里各有一份，照着插两遍就会插进目录。此时会列出候选并退出，换更精确的锚点，
   或显式加 `--all-anchors`。

**明确不做**：批注、修订、水印、内容控件，以及带宏的 `.docm`（`python-docx` 保存会丢
`vbaProject.bin`）。

### 要「重排」而不是「编辑」时

想把别人的 docx 换成我们这套格式，走 Markdown 中转，但**中间产物必须手工清理**：

```bash
pandoc 甲方文档.docx -t markdown --wrap=none --extract-media=media -o 01_内容.md
# 手工清理：删掉原来的封面行与目录块、去掉表头残留的 ** 加粗
python scripts/render.py --src . --out 新版.docx --config cfg.json --pdf --check
```

4 页文档实测：不清理会得到**双封面 + 双目录**，`[1](#...)` 链接语法漏进正文，页数从 4 变 5。
`--extract-media` 抽出的图片路径是相对**执行目录**的，所以在仓库根执行，或把 `media/` 放进 `--src`。

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

## 配置

### config.json 字段

| 字段 | 作用 |
|---|---|
| `cover` | 封面行 `[[样式名, 文本], …]`，样式名用 CoverTop / CoverTitle / CoverSub / CoverInfo / CoverDate |
| `header` | 正文节页眉 |
| `title` / `author` / `subject` / `comments` | 文档属性 |
| `reference_doc` | 指定模板 docx（招标方给强制格式时用它） |
| `toc_heading` | 目录标题，默认「目　　录」 |
| `toc` | `false` → 不插目录页（通知/公示类短文档）；标题样式照常保留，无封面时不分节 |
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

**这张表是 `style` 键与默认值的唯一来源**，别处只指路不复述。全部可省，缺省即中文正式文档惯例：

| 分组 | 键 | 默认 |
|---|---|---|
| 页脚 | `page_number` / `page_number_size` | `— {n} —` / `9` |
| 页眉 | `header_size` / `header_gray` / `header_rule_color` / `header_rule_size` | `9` / `595959` / `BFBFBF` / `4` |
| 目录 | `toc_depth` / `toc_title_size` / `toc_title_color` / `toc_placeholder` / `toc_placeholder_size` | `1-2` / `16` / `000000` / 见提示语 / `12` |
| 题注 | `caption_gray` / `caption_size` / `caption_space_before` / `caption_space_after` | `404040` / `10.5` / `6` / `4` |
| 题注 | `caption_keep_with_next` | `true`（表题不与表格分家；实测关掉可省 1 页，但表题可能落在页尾） |
| 表格 | `header_rows` | 自动——**逐表**识别（读 pandoc 打的 `w:tblHeader`）；填数字则强制；`0` = 该表没有表头（表单 / 附件类首行是字段名），并一并去掉「跨页重复表头」。何时该动它见[视觉控制](#视觉控制) |
| 表格 | `table_border` / `table_shade` / `table_size` | `full` / `EDEDED` / `10.5` |
| 表格 | `table_header_color` / `table_zebra` / `table_zebra_fill` | 不指定 / `false` / `F7F7F7` |
| 表格 | `cell_margin_v` / `cell_margin_h` / `table_para_space` | `40` / `80` / `1` |
| 表格 | `border_size` / `border_color` / `three_line_size` | `6` / `808080` / `12` |
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

### 视觉控制

默认值在 [`style` 段](#style-段)，这里只讲*什么时候该动它*。

- **`header_rows`**——唯一一份文档大概率要改的键。不管它，每张表的表头按源文件自动识别
  （grid table 里 `+===+` 以上全算表头）；表单 / 附件类首行是字段名而不是列标题时填 `0`，
  这会同时去掉跨页重复表头。
- **`table_border: "three"`**——分析型表格用三线表；纯排版用的表想让它别看起来像数据，填 `none`。
  框线粗细由 `border_size` / `three_line_size` 管。
- **`table_shade` + `table_header_color`**——深色表头底一定要配 `table_header_color: "FFFFFF"`，
  否则表头文字淹死在自己的底色里。
- **`table_zebra`**——宽表隔行浅底，首条数据行保持白底。
- **跨页重复表头不用配**：pandoc 原生就给表头行设了 `w:tblHeader`，`post.py` 里那行只是幂等加固。

## 标书场景注意

1. **招标方给了强制格式模板时，以对方为准**——能继承哪些见[复用已有模板](#复用已有模板)。
   那一节替不了你的部分：封面、密封、签字页、页码规则仍要对着招标文件人工核对，
   本工具不替代合规审查。
2. 标书常见结构（投标函 / 商务标 / 技术标 / 报价 / 资质业绩）直接对应 `#` 章即可；
   报价表建议保留 pipe table，便于后期整体替换为招标方表格；复杂表头用 grid table。
3. 交付前跑 `--pdf --check` 并肉眼过一遍渲染图：Word 能打开、无空白页、表头灰底、题注居中，
   四条都过再封包。

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
- **Windows + Word 绑定**：见顶部「运行平台」。解析与排版逻辑不依赖 Word，PDF 导出和一半验收链条依赖；
- `--check` 会把「内容稀疏但合法的页面」也算作近空白页：表单尾页的签字盖章区、
  大表格前的单独标题页等。这类是误报，用 `--max-empty N` 放宽；
- 快照基线依赖本机 Word 版本与字体，**换机器后先 `snapshot.py --update` 重录**，
  否则会满屏漂移；阈值默认 0.1%（实测同文档重复导出为 0.00%，改一处页眉为 0.16%）；
- 不使用 Quarto：其 1.10.x 的 docx 对带自动编号题注的表格会丢失表体。

## 开发

```bash
pip install ruff pre-commit      # ruff 一个工具顶 flake8 + black + isort
pre-commit install               # 一次性

ruff check scripts/ tests/              # 静态检查
ruff format --check scripts/ tests/     # 格式
pre-commit run --all-files              # 上面这些一次跑完
python -m pytest -q                     # 全量：162 项断言，约 6-7 分钟（需本机 Word）
```

pre-commit 钩子在 lint / format 之外，还跑一组**不启动 Word 的快速测试子集**
（`test_version` / `test_docs_sync` / `test_validate_units` / `test_edit` / `test_snapshot`），
所以提交照样快，而「文档与代码互相把守」那条不会形同虚设。全量套件推送前手动跑。

> **没有托管 CI。** 验收类用例要通过 COM 驱动一台真 Microsoft Word，GitHub 托管 runner 给不了。
> 所以本仓库靠 pre-commit 在本地检查，而不是靠流水线。真要加 CI，也只能覆盖上面那个无 Word 的子集。

## 文件地图

| 路径 | 职责 |
|---|---|
| `SKILL.md` | 给其他 agent 用的技能说明（何时用、验证门槛、硬契约） |
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
| `assets/sample.md` / `assets/sample_config.json` | 冒烟测试样例（投标文件风格） |
| `examples/` | 可运行示例：投标文件、公文请示、项目申报书、会议纪要、经营分析报告、技术服务合同、表格排版（见 `examples/README.md`） |
| `docs/` | `SCRIPT_HELP.md` —— 各脚本 CLI 参考（单一来源，见顶部「文档地图」） |
| `baselines/` | 快照基线（样例 4 页 PNG） |
| `tests/` | 162 项 pytest 断言：排版规则、题注识别、表格特性、退出码、快照逻辑、只改版式契约、跨 run 编辑（`test_edit.py`）、模板复用与蒸馏（`test_distill.py`）、9 项验收器（`test_validate.py`）、Patch API（`test_patch.py`）、版本一致性与页码/基线纯函数（`test_version.py` / `test_validate_units.py`）、文档与代码同步守卫（`test_docs_sync.py`：CLI 参数 ↔ SCRIPT_HELP 双向对拍、单一来源、中英镜像结构与脚本覆盖、锚点有效性、SKILL.md front matter 合法性、断言数 ↔ 实际收集数） |
| `CHANGELOG.md` | 版本历史与每条修复的理由 |

## 许可

MIT，见 [LICENSE](LICENSE)。
