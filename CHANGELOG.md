# 更新记录

遵循语义化版本：主版本 = 不兼容变更，次版本 = 新增功能，修订 = 修复。
`ref.docx` 模板一旦改动会体现在次版本号上，因为输出版式可能随之变化
（可用 `snapshot.py` 回归）。

## 0.6.4 — 未发布

### Renderer 抽象收尾：validate 直连渲染器 + LibreOffice/WPS 渲染器 + CLI 选择

把「docx → PDF」从核心彻底抽成可插拔的 `RendererAdapter`，核心（compile / OOXML /
source / image / profile / metadata / evidence）不再依赖任何 Office：

- **`validate.py` 去掉调 `finalize.py` 的子进程往返**：`export_pdf_once` 现在直接消费
  `WordRenderer.render()`（保留线程级超时安全网），`word_acceptance` 检查消息随渲染器名
  动态变化；只收敛为一次导出，PDF 派生检查（page_numbering / blank_pages / visual_drift）仍只吃 PDF
- **新增 `LibreOfficeRenderer`（soffice headless，跨平台 CI 友好）与 `WPSRenderer`
  （WPS Writer COM: KWPS.Application）**，`get_renderer(name)` 工厂统一入口，三者各自带
  `available()` 能力探测
- **`render.py` / `validate.py` 暴露 `--renderer word|libreoffice|wps`**（默认 word）；
  `render.py --doctor` 改为列出三种渲染器的可用性，`preflight` 改用渲染器探测替代硬编码 pywin32 检查
- **`finalize.py` 降级为 `WordRenderer` 的薄 CLI 壳**（仍可单独调用，人读输出不变）
- 测试 230 → 237 项：新增渲染器工厂 / 可插拔接缝（注入 `FakeRenderer` 证明 validate 不再
  shell finalize 或起 Word）/ LibreOffice·WPS 可用性 skip 测试；renderer-agnostic 护栏守住
  「换渲染器不改结论」

### Profile enforcement 作为可选 policy layer（opt-in，非 core acceptance）

外部 review 校准了定位：`profile` 不该成为核心验收的前提——Texere 的核心验收本来就是
「structural invariants + 真实 Word acceptance + visual baseline regression」，template 承担
presentation truth、baseline 承担 rendered truth。所以 `profile` 保持**可选**的
customer-specific policy layer，绝不替代原有 acceptance：

- **profile 从「只选 baseline 的 metadata」接成可执行契约（opt-in）**：`validate.py` 新增
  `compile_profile_checks(profile, docx)`，把 profile 里声明的 `page / styles / table / toc`
  编译成 `profile.page / profile.body_font / profile.heading / profile.table / profile.toc`
  断言（Level 1 Structural，纯 OOXML 读取，不依赖 Word），命名空间独立、可单独追溯
- 这些断言**只在显式 `--enforce-profile` 时计入门禁**（profile 是按需 opt-in 的 contract，
  默认仍只当 baseline 选择器）；未传 `--profile` 时这条链路完全不出现，核心 9 项检查不受任何影响
- `SCRIPT_HELP.md` 补 `--enforce-profile` 文档与「profile 作为可执行契约」小节
- **修正 `profile.table` 边框检测 bug**（opt-in 层自身正确性，不是把 profile 抬成核心）：旧实现用
  `borders.findall(qn("w:border"))`，但真实 OOXML 里 `w:tblBorders` 的子元素是
  `w:top/w:left/w:bottom/w:right/w:insideH/w:insideV`，没有 `w:border` 这个标签，于是恒返回空 →
  有边框的表被错判为「无可见边框」。改为遍历 `w:tblBorders` 的真实子元素（测试辅助函数同步改用
  真实标签，否则会和 bug 互相「自洽」漏网）
- 测试 213 → 219 项：新增 `TestProfileContract`（规范文档全 PASS、违规文档 FAIL、空 profile 无断言）
  + `test_table_border_detection`（钉死边框标签误判）+ `test_missing_required_property_fails`
  （缺失 required 字段判 FAIL）

### 让已有 5 个断言「真的严」（contract semantics，而非听起来严）

外部 review 指出：profile 已声明但尚未执行的字段、以及「要求某字段但文档根本没声明」被静默放行，
都会回到项目一直在消灭的「听起来很严、实际保障没那么强」。这一步只收紧**已有**断言语义，不新增字段：

- **`profile.body_font` / `profile.heading`：required 字段缺失也判 FAIL**：旧逻辑只在「声明了但不符」
  时 FAIL，「profile 要求宋体、实际根本没设东亚字体」会被放过。现在 profile 要求 font_eastAsia /
  font_latin / size / bold 任一项时，实际缺失（None / 未声明）与值不符都判 FAIL，消息里标
  「缺失」vs「≠」
- **`SCRIPT_HELP.md` 新增「Profile 字段执行状态」矩阵**：逐项标注每个 profile 字段是 ✅ enforced
  还是 ⚠️ declared-only（`line_spacing` / `space_*` / `caption.*` / `header.*` / `footer.*` /
  `tender_specific.*` 等目前只是声明、未编译成断言），避免「JSON 写了就以为 validator 会保护我」

### Evidence Enrichment（只填 `evidence`，不增 check / 不改 verdict）

`CheckResult.evidence` 槽位与 `report.json` 落盘早在 profile 章节就已就位，但多数检查留空 `{}`。
这一步把已有判定的「实际观测值」填实——**不新增断言、不改变 PASS/FAIL/SKIP/ERROR、不改变
默认验收行为、只提高 `report.json` 的可解释性**，为将来的 Build Manifest 直接消费做准备：

- **`profile.page`**: `evidence` 把每个维度（width/height/margin_top/bottom/left/right）的
  `expected_cm` / `actual_cm` 并排，审计方无需反解中文 message
- **`profile.body_font` / `profile.heading`**: `evidence.fields` 数组，每项 `{field, expected, actual}`；
  `actual: null` 天然表达「缺失 ≠ 不符」（正是 contract semantics 修的那条），`source: "profile"`
- **`profile.table`**: `{bordered, total, rule: "at_least_one_visible_border"}`——`rule` 把粗粒度语义
  写死，防止误读成「所有表全部符合边框规范」
- **`profile.toc`**: `{has_field, count}`（count 是实数，不只一个 bool）
- **`image_embedding`**: 复用**同一次** `doc_shas` 扫描填 `evidence`
  （`referenced / embedded / resolved / unresolved / images[document_index, sha256]`），绝不为了
  填 evidence 再跑一遍解析（不把 evidence 做成第二套 validator）
- `SCRIPT_HELP.md` 补「检查语义边界（Known limitations，只记录不实现）」：明确 `profile.page` 只看
  `sections[0]`（first_section）、`profile.table` 是 at_least_one 粗粒度、evidence 不参与 status 判定、
  以及 declared-only 字段清单
- 测试 219 → 226 项：新增 `TestEvidenceEnrichment`（5 个 profile 检查的 evidence 形状 +
  `json.dumps` 可序列化守护 + image 复用单次扫描守护）

### Renderer 抽象（Core vs Renderer 抽离，行为 0 变化）

把「docx → PDF」从核心里抽出来，Word 从核心依赖降级为「一个 renderer adapter」——这是
Texere 从「Word automation tool」走向「Document acceptance engine」的第一步：

- 新增 `scripts/renderers.py`：`RenderResult` + `RendererAdapter`（ABC）+ `WordRenderer`
  （原 `finalize.py` 的 Word/COM 逻辑迁入）+ `FakeRenderer`（无 Office 的测试 / CI 渲染器）
- `finalize.py` 改为薄 CLI 壳，只做参数解析与人读输出，**输出格式逐行不变**（行为 0 变化）
- `validate.py` 的 `export_pdf_once` 仍经 `finalize.py` 出 PDF，核心验证逻辑完全未动——
  本次只是把渲染器实现挪了位置，不改变任何检查结论或 `report.json` 结构
- `SCRIPT_HELP.md` 补「渲染器抽象（Core vs Renderer）」：核心不依赖 Office，PDF 派生检查
  （`page_numbering` / `blank_pages` / `visual_drift`）只读 PDF、绝不感知渲染器
- 测试 +4：新增 `test_renderer.py`，含 renderer-agnostic 护栏——用 PyMuPDF 合成 PDF 证明
  三个 PDF 派生检查只消费 `pdf_path`、与具体渲染器无关（换 Word / WPS / LO / Fake 出同一
  PDF 结论不变）。这是后续接 LibreOffice / WPS 渲染器的硬前提
- 断言数 226 → 230

## 0.6.3 — 2026-09-22

### README：从「工程手册」改成「产品首页 + 工程索引」

此前 README 是 manual-first：先讲哲学、再讲文档治理、最后才讲到「你输入什么、得到什么」。
现在按漏斗重排，并把 reference / cookbook 类内容下沉到 `docs/`：

- 第一屏：管线图 + 示例画廊（7 张真实渲染图）+ 一句话定位，不再先要求读者理解
  Semantic IR / design contract
- 新增 `Why not pandoc alone?`（五句话）与 `Have an existing Word template? Keep it.`（客户模板）
- 「三条支柱」下移到设计原则旁：它和「为什么不是 Pandoc」相邻时两节都在列差异化点
- 去掉重复：`--sample` 的产物图不再在 README 展示（首屏画廊是唯一展示入口）；
  `Repository map` 并入 `Documentation map`（两张表都列 `SKILL.md` 与 `docs/`）
- 下沉到 `docs/`，各带 `.zh-CN` 镜像：`CONFIG.md`（config 字段 + `style` 键）、
  `TABLES.md`（表格写法）、`VALIDATION.md`（9 项检查逐项 + 人工门槛）、
  `EDITING.md`（编辑操作 + Patch schema + 重排）。README 只留命令与关键口径
- 英文 600 → 429 行，中文 559 → 396 行；信息一条没丢，只是分层
- `test_docs_sync.py` 新增 `test_docs_mirrors_share_structure`：内容一旦离开 README 就不再被
  原镜像守卫覆盖，不补这条的话「下沉」会悄悄变成「英文一套、中文另一套」
- 下沉同步修掉的跨区锚点：config 表引用 `Visual control`、Tables 引用 `The style section`，
  以及 SKILL.md 里的 `README §Configuration / §Tables / §Validation / §Editing`

### 修复
- **`SKILL.md` 的 front matter 不是合法 YAML，导致技能加载报错**：
  `description:` 的值未加引号，而值中间含 `: `（`… with unified validation: 9 automated checks …`），
  严格解析器报 `mapping values are not allowed in this context (line 2, column 205)`。
  现用双引号包住整个值（文本一字未改），并顺手把 `compatibility` 保持原样（它不含 `: `）。
  仓库内副本与已安装副本 `~/.qoder-cn/skills/texere/SKILL.md` 同步修正。
- `test_docs_sync.py` 新增 `test_skill_frontmatter_is_valid_yaml`：front matter 里出现
  「未引用且值内含 `: `」的键即红。不依赖 PyYAML（它不是项目依赖），用结构检查查这一类错型。
- 测试 161 → 162 项。

### 验收器：把「听起来很严」改成「真的严」

一轮外部 code review 后做的收敛，全部针对同一类问题——**指标比实际保障能力更强**。

- **「查不了就算过」**：9 项检查此前都返回 `(bool, str)`，前置条件缺失时也
  `return True, "跳过…"`，于是 `Passed: 9/9` 里可能塞着 3 项根本没跑的检查。现改为
  `PASS / FAIL / SKIP / ERROR` 四态：`SKIP` 不计入 passed，`summary` 增 `skipped`、
  `signature` 增 `checks_skipped`，摘要行写成 `Passed: 6/9 (skipped: 3)`。退出码仍只由
  FAIL / ERROR 决定。转为 SKIP 的是：source_content（两个参数都不给）、image_embedding
  （无 Markdown 引用数）、page_numbering / blank_pages（无 PDF 或缺 PyMuPDF）、
  toc_field（文档本就没有目录）、visual_drift（无基线）。
- **`toc_field` 是死检查**：靠 `hasattr(doc, "tables_of_contents")` 短路，而 python-docx
  （1.2.0 实测）根本没有这个属性，于是这一项恒定「跳过」，其后三行是永不执行的代码。
  现下到 OOXML 数 `w:instrText` / `w:fldSimple` 里的 `TOC` 域，与 `post.py` 注入目录域的
  写法对齐。
- **视觉漂移两套口径**：`snapshot.py` 逐页全量，`validate.py` 却只比首 / 中 / 尾三页——
  272 页文档第 137 页表格溢出时，抽样的三页可能全都干净；而 validate 的注释还写着
  「逐字节全量比对」。现默认全量，页数双向卡齐（变多同样 FAIL），`--sample-visual` 才退回抽样。
- **页码检查的 fail-open**：识别不出页脚页码格式时旧实现返回 PASS，且文案里写「连续」，
  把一个没验证的结论说成了通过。现返回 SKIP。
- **`source_content` 名不副实**：它 hash 的是 docx 文件本身，只证明字节未变，证明不了
  内容与源一致。新增 `--source-md <file>`：Markdown 归一化后逐段比对 docx 正文，跳过代码块 /
  pandoc fenced div / 表格分隔行 / 列表符号 / 行尾硬换行，并统一直引号与弯引号（前两版在真实
  样例上分别误报 7 处和 1 处，都是这些噪声）。只给 `--expected-hash` 时消息里写明是产件级 hash。

### Patch：schema 声明的能力必须真的能用

- **`add_row` 的 `after_row` 只是预留参数**（源码里写着 `# 预留参数`）：schema 收下、实现忽略，
  Agent 按文档写了会被静默追加到表尾。现按声明语义落地：deepcopy 锚点行的 `<w:tr>` 继承边框 /
  底纹 / 字号，清空文字后插到该行之后；`assess_patch` 校验越界，`--validate` 回验新行位置。
- **`from scripts import edit` 是坏导入**：以 `python scripts/patch.py` 运行时 `sys.path[0]`
  是 `scripts/`，`scripts` 会被解析成空的命名空间包，于是 `set_cell` / `insert_after` /
  `insert_before` / `add_row` 全部抛 `cannot import name 'edit' from 'scripts'`，又被
  `except Exception` 兜成「操作失败」，长期被误当成「目标不存在」。改为显式把脚本目录放进
  `sys.path` 后按模块名导入。

### 图片：从「数量下限」升级为「身份 + 顺序」

- 此前只做 `n_img >= n_ref`：`图A 图B → 图B 图B` 数量正确也判 PASS。实测 pandoc **原样嵌入图片
  字节**（`contract.png` 的 sha256 与 `word/media/rId9.png` 完全一致），于是可以逐图比对：
  md 侧按文档顺序取 `![](path)` 的 sha256，docx 侧按文档顺序取 `a:blip/@r:embed` → rels →
  media 的 sha256，比 missing / extra / 顺序（子序列匹配，允许夹带模板 logo）
- 顺带修掉两个盲区：改走 `a:blip` 而非 `inline_shapes`，浮动型（anchor）图片从此也计数；
  不按 media 文件名排序（`rId12` 会排在 `rId9` 前面，实测踩到），改走文档顺序
- 未给 `--source-md` 时行为不变，仍是数量检查

### signature：把话说准，并补上 report.json 的摘要

- 文件头明写「checksum manifest, not a cryptographic signature」。没有密钥，任何人都能重算
  这些 hash，它证明的是「这份证据描述的是哪个产物」，不是「证据没被改过」
- **顺序 bug**：`report.json` 原本在签名之后才落盘，不在摘要范围内——改报告结论不会破坏
  证据。现在先写 report、再写清单，新增 `report_hash`
- 文件名不动（`signature`），避免破坏已有消费方；README / SKILL / SCRIPT_HELP 口径统一

### 契约措辞收窄

- `SKILL.md` 的「output text comes from the source, character by character」在任何配了
  `content_fixes` 的项目里都是假的。契约作用域明确为 `post.py`（后处理阶段），并写清
  `content_fixes` 发生在渲染管线更早处、是一张用户自己写的显式替换表而非静默改写。
  中英 README 同步。

- 测试 162 → 195 项：新增 PASS/SKIP 分级与反 fail-open、TOC 域真检查、`--source-md`
  正文比对与 Markdown 归一化、`after_row` 插入位置与坏导入回归。

### Layer 0：验证器自身可信化（integrity foundation）

外部 review 后进一步收口「验证器本身必须可信」——这是 0.6.3 收尾的三道关：

- **`CheckResult` 统一四态类型**：各 validator 不再各自 `return (bool, str)`，改为
  `CheckResult(name, status, message, evidence)`；后续 profile 断言 / provenance / CLI 报告都依赖这一类型
- **`finalize.py` 只读验收**：默认只拿输入 docx 的临时副本交给 Word 做域更新 / 重分页 / 导 PDF，
  **绝不写回原文件**；只有显式 `--save-updated-fields` 才写回——验证不再改变被验证对象
- **`doctor` 状态模型 + 能力感知 preflight**：`_word_engine()` 返回 `(available, text)`，
  Word 缺失 / 启动失败不再被硬写成「可用」；`--doctor` 输出 `DOCX 渲染 / PDF 导出 / 验收(--check)`
  三档 `READY / NOT READY`，避免把「docx 可用」误读成「完整环境就绪」
- **图片未定位降级 SKIP**：源图路径解析不到时不再并入强 PASS，整项降为 `SKIP`；
  已定位图片仍按身份 SHA-256 + 顺序做强校验（fail-close 在身份/顺序，fail-open 在解析能力边界）
- **结论行 skip 感知**：除 `Passed: 7/9 (skipped: 2)` 外，最终结论行在有 SKIP 时改报
  `Required checks passed (N skipped — see Summary above)`，不再只写 `All checks passed`
- **回归测试钉死契约**：新增 `TestFinalizeIsReadOnly`（默认不改输入 sha、仅 `--save-updated-fields`
  才写回）；新增 `tests/test_validate_units.py` 纯函数级覆盖页码识别 / 连续性 / 基线像素差 / 四态 /
  fail-open / TOC / Markdown 归一化 / 源完整性 / 图片身份顺序（不启动 Word，加速迭代）
- 测试 195 → 213 项

## 0.6.2 — 2026-09-21

**文档体验优化 + 把 0.6.1 的单一来源约定补全**（对 README 逐条核验代码/仓库实况后修的都是实账）

- 删掉 README 里「GitHub Actions - CI/CD with multi-platform testing」这条：0.2.1 已主动移除 CI
  （见下方「移除」条目，理由就写着「留着一个从未跑过的配置反而会造成『有保障』的错觉」），这句是删除后
  漏改的残留，且「multi-platform」与「Windows + Word 绑定」自相矛盾。新增「开发」一节如实说明：
  本仓库靠 pre-commit 本地检查，没有托管 CI，并解释为什么
- 中英 README 重新对齐：中文版此前整段没有「验证与证据包」（9 项检查、report.json、证据包目录、
  示例输出），三条支柱也还停在 0.5 的旧表述；英文版的「人工门槛」四检查项此前中文版独占。
  现在两份标题层级序列完全一致，并由 `test_readme_mirrors_share_structure` 把守
- `test_docs_sync.py` 新增 5 项守卫（共 7 条用例）：
  - `test_no_key_defined_twice_in_manual`：同一个配置键不得在手册里定义两次。0.6.1 的守卫只拦
    SKILL/BEST_PRACTICES，README 自己在「style 段」与「表格视觉控制」两张表里各写了一份
    `header_rows` 等 7 个键的默认值（对 0.6.1 的 README 跑这条守卫，7 个键全部命中）
  - `test_internal_anchors_resolve`：README 顶部锚点导航必须指向真存在的标题——新增的 23 个锚点
    没测试兜着的话，下次改标题就会默默断链
  - `test_readme_mirrors_share_structure` / `test_readme_mirrors_cover_same_scripts`：中英镜像结构与脚本覆盖对齐
  - `test_stated_test_count_is_current`：文档里写死的断言数必须等于 pytest 实际收集数——治自己
- README 结构体验：加顶部锚点导航；Windows + Word 前置到第一屏（原来埋在 500 行后的「已知限制」）；
  Does/Doesn't 从两个 300 字巨型表格拆成清单；「三条支柱」升为 h2
- README 复述收敛：§Usage 与 §编辑已有 docx 的 CLI 全量转投不再手抄，改为「真要跑的五条流程 +
  指向 SCRIPT_HELP」，与文档地图那条约定一致；`style` 键表声明为默认值唯一来源，「视觉控制」只讲何时该动
- 清理状态性措词：repository map 里 validate.py / patch.py / profiles 上的「**New**」（已三个版本不新了），
  并补齐 4 个 profile
- 修复中文 README 「文件地图」的坏表格（表头重复了一行，多出一个 `| 路径 | 职责 |` 数据行）
- 修复英文 README 开篇标语的破损斜体（`*From Latin* texere*, …*` 星号配对错乱）
- `SKILL.md`：孤儿段落 `"toc": false` 从「硬契约与陷阱」移到「Minimal config shape」，并补上表单类
  `header_rows: 0` 的同类提示；测试时长口径与 README 对齐
- 测试 154 → 161 项，全部绿；版本号 0.6.2（纯文档与守卫，不动任何渲染逻辑，模板未变）

## 0.6.1 — 2026-09-21

**文档信息分层治理**（使用反馈："同一件事在 SKILL/README/BEST_PRACTICES 各有一份，
你在跟漂移打架"——属实，本轮把约定机制化）

- README×2 顶部新增「文档地图」：每条信息只住一个地方，其余文档只指路不复述
- 删除 `docs/CONFIG_SCHEMA.md`：它是 README 配置字段表的第 4 份拷贝
  （0.5.1 审查时造的文件，本身就成了漂移源——诚实记录这个弯路）
- `SKILL.md` 384 行 → 162 行：删去与 README/SCRIPT_HELP 重复的输入要求细节、
  grid 语法块、模板继承表、9 项检查表、repository map，改为引用；
  保留 agent 工作必需的决策信息（硬契约、陷阱清单、Patch schema、验收门槛）
- `BEST_PRACTICES.md`：删「Markdown 写作规范」重复段（指向 README），只留场景经验
- `test_docs_sync.py` 新增字段表守卫：`| \`key` | 说明 |` 式表格出现在
  SKILL/BEST_PRACTICES 即 commit 失败——单一来源约定从自觉变成机器强制
- 测试 154 项全绿

## 0.6.0 — 2026-09-21

**使用反馈落地：批量编辑三件套 + 无目录模式 + 文档守卫**

### 新增
- `edit.py --add-rows 表 数量 [--template-row 行]`: 批量加空行，复制模板行的全部格式
  （边框/底纹/字号/加粗）——python-docx 的 add_row() 是丢格式的裸行
- `edit.py --fill data.json|data.csv`: 批量填表；null/空单元格跳过不清空；
  合并单元格按「写主格、跳过后续坐标」处理；越界只警告不崩
- config `"toc": false`: 通知/公示类短文档不插目录页，标题样式照常保留，
  无封面时不分节。旧写法只能「不写 #」绕过，代价是全文变普通段落手工后补
- `render.py --src` 支持直接传单个 .md 文件（一页的通知不必先建目录）
- `finalize.py` 的 out.pdf 可省略，默认取输入同名 .pdf
- 空白格写入时借用同表已有 run 的 rPr，填空白表单不再字体回退
- `tests/test_docs_sync.py`: CLI 参数 ↔ SCRIPT_HELP 双向对拍守卫
  （漏写文档/虚构文档参数都会被拦），已入 pre-commit 快速子集

### 修复
- Windows 中文乱码根修：子进程管道统一注入 `PYTHONIOENCODING=utf-8`
  （旧乱码根源是子进程 GBK 输出被父进程 utf-8 解码，非控制台编码问题）
- 全部 `fitz` 别名换成 `pymupdf`（PyMuPDF 1.28 起 fitz 每次加载都刷 deprecation warning）
- pre-commit 形同虚设：hooks 从未安装；配置升级 ruff v0.16.8、
  新增 check-added-large-files（拦误提交产物）与秒级测试子集

### 文档
- README×2 / SCRIPT_HELP 补齐本轮全部新参数与 `toc` 字段；make_ref 历史欠账（--dst/--latin-font）补录

### 测试
- 152 项断言（新增 toc:false ×2、批量编辑 ×6、文档守卫 ×7）

## 0.5.1 — 2026-09-20

**全面代码审查修复 + validate 性能重构 + 示例扩充**

### 修复
- `patch.py`: 错误信息统一输出到 stderr；补 apply 时的备份与另存逻辑；
  修复 dry-run 在空 operations 时的 `UnboundLocalError`；修复跨模块 import
- `validate.py`: 修复 `expected_hash` / `max_empty` / `baseline_dir` 未传入检查函数的
  `NameError`（`--max-empty` 此前实际无效）；补丢失的 `re` / `glob` 导入；
  异常捕获按类型细化；`--quiet` 不再隐藏检查项列表（调用方可从 stdout 判断哪一项失败）
- 页码连续性检查重写：只从页脚区域（按版面位置 `sort=True` 排序后的页尾三行）提取
  页码，避免把正文数字（日期、金额）误当页码；校验缺口序列
- 视觉漂移比对与 `snapshot.py` 统一口径（dpi=100 + 全量 `diff_ratio`）；
  旧实现的 2x 矩阵与 `baselines/` 尺寸对不上，会把「没漂移」误判成漂移
- `fitz` 改惰性导入：未装 PyMuPDF 时基础检查（包结构/分节/TOC）仍可运行
- 所有 `subprocess.run` 加超时；所有 fitz 文档显式 `close()`（防 Windows 文件句柄泄漏）
- `examples/gongwen/config.json`: 移除指向 JSON 的 `reference_doc`（pandoc 要求 docx，
  该示例此前无法渲染）；`tender/` / `gongwen/` 补占位图（md 引用了不存在的 png）
- `render.py`: 启动时清理超过 24h 的旧临时目录（实测一天可积累 12 个）

### 性能
- validate 的 4 项 PDF 检查 + 证据截图从「各自启动一次 Word」收敛为一次导出共享，
  完整测试套件耗时 800s → 394s

### 新增
- 示例从 4 个增至 7 个：`minutes/`（会议纪要）、`report/`（经营分析报告：多文件合并、
  Lead 提示框、grid 二级表头、图注）、`contract/`（技术服务合同：grid 单元格内换行、签署栏）
- `scripts/_version.py`: 版本号单一来源，`tests/test_version.py` 校验与 pyproject 一致
- `tests/test_validate_units.py`: 16 项纯函数单测（页码识别/基线比对），不启动 Word
- `docs/CONFIG_SCHEMA.md`、`docs/SCRIPT_HELP.md`
- `validate.py` 新增 `--baseline <dir>` 参数；patch.py 新增 `--no-backup`
- 依赖升级：lxml 6.1 / PyMuPDF 1.28 / python-docx 1.2 / pywin32 312；pandoc 3.11 实测兼容

### 文档
- 修正 README / README.zh-CN / SKILL / BEST_PRACTICES 中过时的测试数（74→137）、
  「no Word needed」表述、不存在的 `render.py --profile` 用法
- 删除 `PLUGIN_README.md`：其描述的两个插件目录已在上一提交中移除，全文指向不存在的路径

### 版本说明
上一版 0.5.0 发布时漏 bump pyproject/_version（停在 0.4.0），本版一并拉齐。

### 测试
- 137 项断言全过（含新增 validate/patch/version/纯函数用例）

## 0.5.0 — 2026-09-19

**Productize texere as document compiler + verification layer**

### 新增
- **Profile 体系扩展**: 新增行业专用 profile
  - `profiles/formal-cn-v1.json`: 中文正式文档默认 profile
  - `profiles/tender-v1.json`: 投标文件专用 profile（标书格式、报价表样式）
  - `profiles/gongwen-v1.json`: 公文专用 profile（GB/T 9704-2012 标准）
  - `profiles/application-v1.json`: 项目申报书专用 profile（表单式表格）
- **统一验收 CLI (`scripts/validate.py`)**: 9 项自动化检查 + 证据包生成
  - Package integrity, Image embedding, Section count, TOC field
  - Page numbering, Blank pages, Word acceptance, Visual drift
  - 输出：`report.json` + `page-XXX.png` + `signature`
- **Agent Patch API (`scripts/patch.py`)**: 声明式编辑原语
  - Hash precondition + must_contain precondition
  - Dry-run → Assess → Apply → Validate 完整流程
  - 支持 7 种原子操作（replace_text, insert_after, set_cell 等）
- **真实项目案例**:
  - `examples/tender/`: 智慧园区平台建设投标文件（完整示例）
  - `examples/gongwen/`: 关于推进数字化转型工作的请示（公文示例）
- **README/SKILL.md 重写**: 强调 "Compiler + Contract + Evidence" 定位

### 产品定位进化
从 "turn Markdown into properly typeset Chinese documents"
到 "Reference/Spec → Deterministic Document → Evidence"

建立三重竞争壁垒：
1. **Contract 壁垒**: profile.json 让版式规则显式化、可审计、可复用
2. **Evidence 壁垒**: validate.py 提供 delivery-ready guarantee
3. **Agent 壁垒**: patch.py 提供声明式编辑原语，适合 Goal Mode/MCP

### 测试
- `tests/test_distill.py`: 12 项断言，验证模板蒸馏报告与 config 可用性
- `tests/test_validate.py`: 新增 20+ 项断言（待添加）
- `tests/test_patch.py`: 新增 15+ 项断言（待添加）
- 总计 74+ 项断言通过

## 0.4.0 — 2026-09-18

复用已有模板（甲方模板 / 自己攒的模板）。不启用新开关时，输出版式零变化。

### 新增
- **`scripts/distill.py`：模板蒸馏**。读一份已有 docx，输出「建议 config」+ 体检报告：
  页面设置（纸型 / 方向 / 页边距）、正文与各标题的字体字号、各节页眉页脚、分节数与表格数，
  并给出对应的 `make_ref.py` 命令。目的是接一份新模板时，不用手工去 Word 里量页边距、抄页眉。
  能确定的报出来，判断不了的（封面结构、页码格式、表格细节）在报告里列成人工确认项——
  **蒸馏是"把模板拆成 config 字段、人来拍板"，不是让工具去理解模板**
- **`style.page_number: null` —— 保留模板自带页脚**。此前页脚被无条件重写成 `— n —`，
  甲方模板里带公司名 / 文档编号的页脚会被冲掉。实测默认行为其实是**追加**
  （`…XYZ-2026— 1 —`），只有显式写 `null` 才是一个字都不动

### 修正
- **README 低估了 `reference_doc` 的继承能力**。实测：页边距、纸型、正文与标题样式，
  以及**页眉**都能继承（页眉只在 config 配了 `header` 时才被覆盖）。原文档只写
  "正文样式即继承对方模板"，会让人以为页边距和页眉得手工重设——其实是白费功夫。
  已按实测结果补成一张表

### 测试
- `tests/test_distill.py`：12 项断言。守蒸馏报告的内容与 `--out` 写出的 config 可用，
  以及 `page_number: null` 时模板页脚**一个字都不许追加**
- 74 项断言全过

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
