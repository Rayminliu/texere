# 更新记录

遵循语义化版本：主版本 = 不兼容变更，次版本 = 新增功能，修订 = 修复。
`ref.docx` 模板一旦改动会体现在次版本号上，因为输出版式可能随之变化
（可用 `snapshot.py` 回归）。

## 0.3.0 — 2026-09-18

按 Agent Skills 规范做结构重组，并新增「编辑已有 docx」的第二条链路。
输出版式零变化：62 项断言全过，`--sample` 仍是 4 页、快照比对 0.00%。

### 变更（破坏性：路径全变）
- **脚本移入 `scripts/`，模板与样例移入 `assets/`**，符合规范的
  `SKILL.md` + `scripts/` + `assets/` 结构
- `render.py` / `make_ref.py` 的 `KIT` 由「脚本所在目录」改为「其父目录」（即仓库根）。
  `tests/` 里本来就是这么算的，所以测试侧只需改引用路径，逻辑不动
- **命令行随之变化**：`python render.py` → `python scripts/render.py`（双语 README 已同步）

### 修复
- **`--doctor` / `preflight` 给出的安装命令必然失败**：代码里 4 处仍写着
  `pip install "."` / `".[pdf]"` / `".[check]"` / `".[pdf,check]"`，而本仓库没有构建后端。
  0.2.1 只改了 README、没改代码，恰好是那类"会让人直接失败"的错误。
  现统一改为 `pip install -r requirements.txt`（缺哪个可选依赖就单装哪个）

### 新增
- **`scripts/edit.py`：编辑已有 docx 的第二条链路**。与渲染链路契约相反：
  渲染是「只改版式、不改内容」，编辑是「**只改你指定的地方，其余字节原样保留**」——
  不注入封面 / 目录 / 页码，不重排样式。支持：替换文字、锚点前后插入段落、删除段落、
  改单元格、增删行、改页眉页脚
- 跨 run 替换（不写出来必踩的两个坑）：
  - **Word 会把文字切成多个 run**：`自开标之日起 90 日历天` 实测是三个 run，数字单独一个，
    逐 run 搜索会漏。做法是拼接整段 `w:t` 定位，替换内容只写进「命中起点所在的 run」，
    其余 run 仅删掉被覆盖的字符——run 数与各 run 的格式（加粗/颜色）都保住
  - **`xml:space="preserve"`**：`w:t` 首尾有空格时必须设，否则 Word 会吞掉空格
  - 写法借鉴 Codex Documents 插件的 `redact_docx.py`（拼接 → 定位 → 按原切分写回），只取思路
- **锚点命中多处时拒绝执行**：实测目录被 Word 刷成静态文本后，`1.2 资质与业绩` 在目录与正文
  各有一份，照着插两遍会把内容插进目录。现在列出候选并退出，可换更精确的锚点或加 `--all-anchors`
- 默认备份 `<name>.bak.docx`；`--verify` 复用 `finalize.py` 让 Word 真机打开一次，
  编辑链路与渲染链路共用同一个验收环节
- **明确不做**（写进文档，避免以后被当成遗漏）：批注、修订、水印、内容控件；
  带宏的 `.docm`（`python-docx` 保存会丢 `vbaProject.bin`）
- `tests/test_edit.py`：26 项断言，守跨 run 替换、run 结构与格式保留、含图段落跳过、
  歧义锚点拒绝、页眉只动正文节、备份与 `--list` 只读
- `examples/`：可运行的最小示例——`form/`（表单式文档，`header_rows: 0`）与
  `tables/`（grid 多级表头 + 合并单元格 + 斑马纹）。自带 md 与 config，渲染产物已忽略

### SKILL.md 按规范重写
- frontmatter 补 **`license: MIT`** 与 **`compatibility`**（Windows + 本机 Word + pandoc ≥3.1；
  Linux/macOS 只有 docx 那半条链）。后者让 agent 在**加载前**就能判断环境是否达标，
  不必等跑完一轮才发现没有 Word
- `description` 改为规范要求的**第三人称**（`This skill should be used when...`），
  并补上负向约束（不该用于编辑既有 docx / 学位论文 / 英文文档）。实测 811 字符（上限 1024）
- 正文新增规范要求的 **`## Examples`**（4 段 input→output：最小标书、grid 多级表头、
  图注、无标题表单）与 **`## Guidelines`**（原 hard contract 与边界规则合流）
- **删去与 README 逐字重复的四节**（Configuration 全表 / Known limitations / Repository map /
  Performance），改为一句指向 `README.md`——规范明确要求同一信息只存一处

## 0.2.1 — 2026-09-17

面向公开使用（全球读者）的文档与元数据整理。功能无变化，`ref.docx` 未动。

### 新增
- `SKILL.md`：给其他 agent 用的技能说明（何时用/不用、验证门槛、硬契约）。frontmatter 的
  `description` 用英文写，便于全球 agent 命中
- `README.zh-CN.md`：原中文 README 改名保留；`README.md` 改为**英文主文档**
  （GitHub 只把 `README.md` 渲染成仓库首页），两文件首行互为语言切换入口
- `LICENSE`：**MIT**；`pyproject.toml` 加 `license` 字段

### 变更
- **改名 `docx-kit` → `texere`**（拉丁语"编织"，即 `text` / `textile` 的共同词源；排版本就是
  把正文、表格、题注、页码编织成一页）。改名原因：原名实测已被多处占用——npm 包名 `docx-kit`
  被占、GitHub 另有 `ntnyq/docx-kit`（TS 库，有独立文档站）与 `LLYN077/docx-kit`。
  涉及 10 处手改 + 2 处自动重生成。**功能与输出版式零变化**（快照 0.00%，36 项测试全过）。
  两个容易漏的点：`SKILL.md` 的 `name` 必须与技能目录名一致；`render.py` 的临时目录前缀与
  `tests/` 里检查"临时目录泄漏"的 glob **互相耦合，必须同步改**
- 候选名核验方法与实测结论（供将来命名参考）：
  - 用 GitHub `in:name` 计数，**看占用质量而非数量**（是否活跃、多少 star）
  - **子串陷阱**：`pagina` 命中 **229,988**（pagination）、`collatio` 命中 **444**（collation）、
    `maat` 命中 **2,268**（code-maat）——名字藏在常见词里就永远搜不到自己
  - **拼音同音陷阱**：`kaogong`（考工）被 **191** 个"考公"备考仓库淹没
  - 神话/古籍类名字基本被占：`maat` 2,268、`scriptorium` 727、`scriba` 404、`seshat` 402、
    `erya` 329、`colophon` 161、`nisaba` 71、`shuowen` 41
  - 最终存活的是**拼写独特且非英语日常词**的：`texere` 13 个（全是 0–5 star 空壳）
- README 重写为使用者视角：新增「能做什么/不做什么」「效果」（引用 `baselines/` 的两张渲染图）
  「验证门槛」；「文件地图」移到末尾
- 修正 README 两处**会让人直接失败**的错误：
  - `pip install .` —— 实测失败（本仓库无构建后端，且工具靠相对路径找 `ref.docx`/`filters/`），
    改为 `uv sync --all-extras` / `pip install -r requirements.txt`（均已验证）
  - `tests/` 的描述仍写着已删除的「自动编号与交叉引用」

### 仓库
- 清除文档中残留的项目名与开发机绝对路径
- 提交身份与全局配置改为 GitHub 匿名邮箱（`<id>+<user>@users.noreply.github.com`），
  并 `reflog expire` + `gc --prune=now` 使旧提交不可取回；**因此本日之前的 commit hash 全部变更**
- 删除从未验证过的 GitHub Actions 配置（本地自用，且 Linux runner 无 Word，只能覆盖半条验收链）

## 0.2.0 — 2026-09-16

### 新增
- 表/图按章自动编号与 `@tab:` / `@fig:` 交叉引用（config `"auto_number": true` 启用）
- `filters/captions.lua`：在 pandoc **AST 层**把题注标记成 `TableCaption` / `FigureCaption`，
  `post.py` 不再靠正则反推
- `snapshot.py`：PDF 版式快照回归，逐像素比对 `baselines/`，漂移即 `exit 1`
- `post.py` 支持 config 的 `style` 段：**全部版式数值可配**——页码模板与字号、页眉字号/灰色/下边框、
  目录标题字号与占位提示、题注颜色/字号/段前后、单元格边距、表格边框粗细颜色、三线表顶底线
- `caption_words`：题注关键字可自定义（默认 表/图/Table/Figure），`post.py` 与 lua filter 同步

### 已知限制（`auto_number` 开启时）
- 会改写题注段落文字：题注内的局部加粗/斜体被合并掉
- 会给原本无编号的题注补号（如 `附件 8-1 …` → `图 8-1 附件 8-1 …`）
- 不同步正文里手写的引用（`（表2-2）` 这类），只有 `@tab:` / `@fig:` 标签会更新
- **编号本已正确的文档建议不开**：真实项目实测，关闭后页数一致且原文一字不改
- **表格**：支持 grid table（多级表头 / 合并单元格 / 单元格内换行 / 列宽控制，均为 pandoc 原生能力，已文档化）
- `style.header_rows`：多级表头的视觉表头行数；`style.table_border`：`full` / `three`（三线表）/ `none`
- `make_ref.py` 支持 `--body-font / --latin-font / --heading-font / --body-size`
- `render.py --version`

### 移除
- **GitHub Actions CI 配置**：本地自用 + 只有熟人使用，且 CI 的 Linux runner 没有 Word，
  只能覆盖 33 项断言那半条链，而 Word 真机验收才是这个项目最值钱的部分。
  留着一个从未跑过的配置反而会造成"有保障"的错觉。需要时再写，五分钟的事。
- **图表自动编号与 `@tab:` / `@fig:` 交叉引用**（0.2.0 内测期间的功能）整体删除。
  理由：它会改写题注文字（题注内局部加粗被合并）、给原本无编号的题注补号
  （`附件 8-1 …` → `图 8-1 附件 8-1 …`）、且不同步正文里手写的引用。
  真实项目实测：关闭后页数一致、原文一字不改，说明它在"编号已正确"的文档上零收益、纯风险。
  现在 `post.py` 的契约是**只改版式、不改内容**，并有测试守着。

### 新增（借鉴通用 docx 技能的表格规范）
- `style.table_header_color`：表头文字颜色（深色底配白字）
- `style.table_zebra` / `table_zebra_fill`：表体隔行浅底，首条数据行保持白底
- 默认值不变（浅灰表头 + 黑字 + 无斑马纹，中文正式文档惯例），
  有测试同时守住「可配」与「默认不变」两侧

### 新增（真实项目实测暴露）
- **`content_fixes` / `content_fixes_file`**：编辑性替换表（删注释性括号、统一措辞），
  在合并 md 之后、转换之前套用。指向上游已有的 `.py` 时用 `ast` 只读取值、不执行代码。
  真实项目里这一步影响 122 处、约 1581 字，不做的话正文与定稿版本不一致

### 修复（第三份真实素材：表单式申报书）
- **无一级标题的文档被直接拒绝**：附件/申报书这类表单文档天然没有章节标题，
  以前 `post.py` 会 `sys.exit`。现改为：跳过目录注入（无标题可索引）、
  封面仍插到最前面、表格与题注排版照做
- **空分节符产生完全空白的首页**：没有封面也不插目录时，第 1 节里只剩一个
  空的分节段 → 整页空白。现在只有真往第 1 节放东西时才分节
- **表单表格首行被误当表头**：外部 docx 转 md 后 pandoc 会给首行加 `w:tblHeader`，
  「项目名称」那行就被灰底加粗、还跨页重复。新增 `style.header_rows: 0`
  显式声明"这张表没有表头"（同时清掉跨页重复标记）

### 修复（第二份真实素材实测暴露）
- **封面插错位置**：封面/目录原本插在「第一个一级标题之前」，源文件在第一个 `#`
  之前写了内容时，那些段落会排到封面之前单独占一页。现改为插入 body 最前面；
  同时**提示而不删除**这些前置内容（契约要求不改内容）
- **`keep_with_next` 语义修正**：原本无差别加在所有题注上，但图注在图片**之后**，
  粘住它会把后面内容整块推走——47 图的真实文档实测多出 **4 页**。
  现只对「位于表格之前的表题」与「载有图片的段落」生效：**60 页 → 56 页**
- **UTF-8 BOM 导致读配置失败**：`render.py` / `post.py` 用 `utf-8` 读 config，
  而 Windows 记事本、PowerShell 写出的 json 默认带 BOM。全部改用 `utf-8-sig`
  （源 md 同样处理），并加了回归测试

### 修复（真实项目实测暴露）
- **自动编号篡改正文**：`**表层（边缘轻算力）：**` 被当成表题，改写成
  「表 3-5 层（边缘轻算力）：」并毁掉加粗。现在文档里一旦存在 AST 层打的题注样式，
  就**只信样式**，不再用文本正则兜底
- `filters/captions.lua` 的文字判定收紧为「关键字 + 编号」（或短句无句号），
  避免把「表层…」这类正文标记成题注
- **自动编号会抹掉图片**：pandoc 把图片放在 `Captioned Figure` 样式的段落里，
  按文本重写该段落时清空了带 `w:drawing` 的 run——真实项目里 28 张图全部丢失。
  现跳过含图段落，并加了回归测试（去掉防护即测试失败，已验证）
- **图片搜索范围不含 src 的兄弟目录**：真实项目 md 在 `build/src/`、图在 `build/media_plan/`，
  导致 28 张图全部找不到。现自动加入父目录及其子目录，并支持 `resource_paths`
- **pandoc 的 WARNING 被静默吞掉**：缺图只给警告不报错，会让人交付一份没图的文档。
  现在 `run()` 始终上报诊断信息，渲染后再自检 `images: n/m ok`

### 修复（其他）
- ~~跨页重复表头（`w:tblHeader`）此前只写在 README 军规里，**代码未实现**~~
  **更正**：pandoc 原生就给表头行设了 `tblHeader`（grid table 的 `+===+` 以上全部算表头），
  本轮新增的 `set_repeat_header()` 只是**幂等加固**，并非该功能的实现者。
  之前的判断是只 grep 了代码、没验证 pandoc 输出得出的，在此更正。
- `make_ref.py` 硬编码开发机上的绝对路径，换机器即失效且模板会悄悄漂移
- `post.py` 缺一级标题时抛 `StopIteration`，现给出可诊断提示
- Windows GBK 控制台下 `render.py` 打印子进程输出会 `UnicodeEncodeError` 崩溃
- `--check` 单独使用时被静默忽略，现自动补 `--pdf`
- 表题未设 `keep_with_next`，可能与表格分家（表题留页尾、表格跑下页）
- `--doctor` 依据注册表 `CurVer` 误报 WPS，改为 **COM 实测**引擎身份
- `check_pdf.py` 无论检出多少空白页都 `exit 0`，现带阈值退出码

## 0.1.0 — 2026-09-16
- 从一份中文正式文档渲染管线抽离并独立成包
- 依赖迁移到 `pyproject.toml` + `uv.lock`，可选依赖按用途分离
- 新增 `render.py --doctor` 环境自检
- 新增覆盖「排版六条军规」的 pytest 断言
