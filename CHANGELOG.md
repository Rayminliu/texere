# 更新记录

遵循语义化版本：主版本 = 不兼容变更，次版本 = 新增功能，修订 = 修复。
`ref.docx` 模板一旦改动会体现在次版本号上，因为输出版式可能随之变化
（可用 `snapshot.py` 回归）。

> 本文件只保留 0.9 起的条目；0.8.1 及更早见 [docs/CHANGELOG-archive.md](docs/CHANGELOG-archive.md)。

## 状态：维护模式（0.6.4 起）

0.6.4 是最后一轮大规模架构更新。核心模型已收敛为：

    Template = 设计源  ·  Baseline = 观测真相  ·  Validator = 产物不变量
    Profile  = 显式策略  ·  Renderer  = 可插拔后端  ·  Evidence  = 事实记录

此后进入 maintenance mode，更新从 roadmap 驱动改为**问题驱动**：

- **功能开发：停止**（不再排 0.7 / 0.8 … 的功能路线图）
- **Bug 修复：开启**（DOCX / PDF / OOXML 边界情况）
- **兼容性维护：开启**（新版本 Word / WPS / LibreOffice / Python / python-docx 导致的行为变化）
- **依赖 / 安全：开启**（依赖库出现安全问题或不兼容）
- **新功能：仅在确有未覆盖、且值得进入核心的真实需求时**

目标不是「永不更新」，而是「即使半年、一年不更新，项目依然完整、可信、可用」。

> **状态更正（0.9.3）**：上面"功能开发：停止"的声明与实际行为脱节——0.6.4 之后
> 实际发布了 0.7.0 / 0.8.0 / 0.8.1 / 0.9.0 / 0.9.1 / 0.9.2，其中 0.8.0 与 0.9.0
> 是功能版本，近半年保持每周量级的提交。实际执行的模式是**问题驱动的小步演进**
> （新功能仅在确有未覆盖的真实需求时进入核心），本表如实更正，不再自称维护模式。

## 未发布

### 二轮清理：caption 幽灵键、tender-v1 并入 formal-cn、CHANGELOG 归档

过度设计审计的延续（用户裁决"都可以做"）：

- **policy 模式预设里的 `caption` 删除**：`PROPERTY_CHECKS` 里 caption 映射为空
  （caption_recognition 检查从未实现），预设里的 warn/enforce 严格度永远查不到
  ——上一轮注释说"原幽灵值已清除"其实只清了一半，本轮连空映射行一起删净。
- **`profiles/tender-v1.json` 删除**：`*_specific` 段清掉后它与 formal-cn-v1 除
  name/description 外逐字相同（七段等值守卫钉住的事实），一份 JSON + 等值守卫 +
  测试参数化只为"投标场景有个名字"存在。投标文档直接用 formal-cn-v1；README
  双语仓库地图、BEST_PRACTICES 场景表与实战案例同步改指向；SHARED_SEGMENTS
  守卫只剩 application-v1 一员；test_validate 参数化从四个 profile 减为三个。
- **CHANGELOG 归档**：0.8.1 及更早条目（全文约八成，91KB）移至
  docs/CHANGELOG-archive.md，主文件 113KB→22KB；test_version 的
  "首个版本标题 = 当前版本"守卫不受影响（0.9.2 仍在主文件顶部）。
- 顺手修正：README 双语仓库地图里 test_policy 的描述仍写着上轮已删的"覆盖分层"。

### 删除 policy 分层叠加与 profile `*_specific` 声明段（死代码清理）

过度设计审计（用户裁决"清理删除"）：两处"文档承诺了、代码从不消费"的残留。

- `policy.resolve_policy` 删除：三层叠加（template spec → approved → task）从未被
  validate.py 调用（主链路只用单层 `mode_policy` 预设），唯一调用方是它自己的单测
  ——经典的 speculative generality。模式预设与严重度裁决语义不受影响。
- 三个 profile 的 `*_specific` 段删除（tender_specific / official_specific /
  application_specific，共 20 个字段）：零代码消费（无 validator、无渲染器读取），
  SCRIPT_HELP 一直如实标注 declared-only，但 BEST_PRACTICES 却把"报价表样式、
  骑缝章要求"当作 tender-v1 的关键特性宣传——对以"验证"为卖点的项目，摆设字段
  是自我打脸。骑缝章、签章页这类客户语义回到它们真正的家：README / SKILL 的
  "仍需人工检查"清单本就写着。
- 文档同步：SCRIPT_HELP 执行状态表删 `tender_specific.*` 行与清单项；VALIDATION
  双语表格删 `*_specific.*` 通配；test_profiles 的 `_specific` 通配守卫分支随之
  删除（未来再写 `_specific` 键会直接红"文档未覆盖的顶层键"，保护不变）。
- BEST_PRACTICES 场景匹配表改讲真话：formal-cn 的"三线表"实为全框线
  （`table.border: full`）；tender-v1 与 formal-cn 七段逐字等值（守卫钉住的），
  如实写"完全同版式"；application 的真实差异是页脚「第 {n} 页」。

### 删除 --sample CLI 入口（examples/ 接管全部演示职责）

用户裁决："sample 不需要，要测试有 example"。上一轮已把 `--sample` 降位出 Quick start，
本轮彻底删除——开箱叙事只剩一条线：examples/tender 直跑就是第一份产物：

- `render.py`：`--sample` 分支与 argparse 选项删除；缺 `--src/--out` 的报错改为
  直接给出示例命令；docstring/epilog 同步。
- **`assets/sample.md` 保留**：它仍是钉死的回归夹具（四个测试文件与 `baselines/`
  渲染它），只是不再有 CLI 入口——夹具身份在 README 双语仓库地图改写。
- 文档清引用：README 双语（"没装渲染器"段、配置节、仓库地图）、SKILL、
  SCRIPT_HELP（示例/参数表/FAQ）。
- 根目录 `sample_out.docx` 残留清除；测试套件确认零依赖 `--sample`（全部走
  `_sample_render` 的显式 `--src` 渲染）；守卫全绿，开箱命令实测 exit 0。

### 测试提速：全量套件 354s → 164s（-54%），validate e2e 165s → 85s

用户反馈 pre-push 的 6 分钟太慢。诊断（pytest --durations）：test_validate 的 20 个
word 用例共 165s，其中 15 个用例**各自 spawn 一次完整 validate**——同一份样例文档、
每用例一次 Word COM 导出（~8s/次），而绝大多数是只读断言（report 结构 / 检查状态 /
signature 内容 / stdout 措辞），对同一份文档重复问同一个问题。修复沿用渲染侧已有的
收敛模式（tests/_sample_render.py 进程级渲染缓存）：

- `_sample_render.shared_evidence()`：进程内**唯一一次** validate 运行，共享
  report.json / stdout / evidence 目录给全部只读断言——Word 往返从 15 次降到 1 次；
- 改写只读用例 11 个（valid_package / images / sections / toc / page_numbering /
  blank_pages / word_acceptance / drift_skip / evidence_artifacts / exit_code /
  report_structure）为读共享产物，断言语义逐字保留；
- 四个内置 profile 用例参数化合并（TestProfileValidation），断言不变；
- **不动的**：不同输入的用例（坏包、--max-empty、drift_with_profile、--config/
  --reference provenance、finalize 写回、patch 链）保持独立 spawn——它们验证的
  是「不同参数下的行为」，共享会改变被测语义；真机 Word 验收一项不砍。

实测：word 套件 165.7s → 85.4s；全量 354s → 163.9s（465 passed, 11 skipped 零回归）。

### distill 报告新增字体本机可用性检测（回答"字体怎么保证"）

字体是引用不是嵌入：docx 只写名字，渲染时系统按名解析，**本机没有就静默回退**——
版式契约（reference_doc / profile 断言）无声失效，用户到最后才知道字不对。蒸馏报告
现在逐字体标注解析结果：

- 实现 `distill._font_available`：GDI+ flat API（ctypes，零新依赖），
  `GdipCreateFontFamilyFromName` 按名解析（与 Word 同源，中英文别名内部处理），
  缺失返回 `FontFamilyNotFound(17)`——状态码即真相，不靠回落值猜；
- 展示规范英文名（`GdipGetFamilyName` 传 LANGID 0x0409）：**宋体→SimSun、
  楷体→KaiTi**——跨机器稳定不变的注册名，可作对拍锚点（实测：传 0 取进程
  locale 名，zh-CN 机器上请求 SimSun 也返回"宋体"，locale 名做不了锚点）；
- 缺失标 `[字体缺失]` 并指路（安装该字体或换本机等价字体）；非 Windows /
  探测失败直通不拦路（探测是提示不是门禁，宁缺毋假）。
- 测试 +2（别名解析展示 / 缺失必须点名不许静默）；SCRIPT_HELP distill 节同步。

### --sample 回归纯 docx 冒烟（用户实测抓到的"隐性 PDF 默认"）

用户反馈"看起来会默认产出一个 PDF"。排查结论：`render.py` 本体确实不带默认 PDF
（`--pdf` 是 store_true），但 **`--sample` 内部把 `want_pdf=True` 和 Word 渲染器
硬编码**（0.6.x 时代写死）——而它恰恰是新手跑的第一条命令，被 CLI 帮助宣传为
"零配置冒烟测试"。后果：没装 Word 的机器第一条命令必炸；与本仓 0.9.2 刚立档的
「按文档阶段选流程、迭代不带 PDF」自相矛盾；CI 作者也早知道（ci.yml 注释原话
"--sample 内部强制走 PDF，在 runner 上必炸"，故绕行显式 --src）。修复：

- `render.py`：`--sample` 改为纯 docx 路径（want_pdf=False, want_check=False），
  preflight 不再触碰渲染器探测——没装 Office 的机器零依赖跑通；`--help` 措辞同步；
  结尾摘要自然落到 docx-only 分支（"需要 PDF/真机验收时追加 --pdf --check"）。
- 根目录 `sample_out.pdf` 残留清除（gitignored 本地产物）；README 双语 ×2、SKILL、
  SCRIPT_HELP 共 5 处 `sample_out.docx/.pdf` 注释改为 `sample_out.docx`（注明纯 docx
  不需渲染器）；ci.yml 注释更正并说明 CI 保持显式路径的理由。
- 实测：`--sample` 0.6s 出 14KB docx，无 PDF；想看带 PDF 与截图的完整演示，
  Quick start 第三条命令（显式 `--pdf --check`）本来就是。

### 路由修复：distill.py 从文档深处提到入口（用户实测三轮返工的教训）

用户拿着甲方 docx 模板，第一反应是手写脚本解析 XML 猜样式，三轮才撞到"封面继承不了"
——而这条早就写在 README 复用模板一节（第 346 行的 `distill.py` 提示）。三个事实：

- **SKILL.md 里 distill 零命中**：守卫防"说的每句是谎"，不防"该说的没说"——这是
  路由缺口，不是诚实性缺口。修复全部是文档级轻改动，零新机制：

- **SKILL.md Quick start**：加"Client supplied a template docx? Distill it first"一段
  （一句命令 + 封面继承不了 + 先看一眼封面）。
- **README 双语快速开始**：`distill.py --out cfg.json` 与"按文档阶段选流程"并列；
  明确"先蒸馏再写 config""搭封面之前先看一眼模板封面"。
- **validate 摘要加范围声明**（用户反馈"SKIP 让全绿像质量良好"）：每次报告末尾打印
  `Note: these checks prove structure & content, not typography — judge fonts/cover/
  layout by eye on the rendered pages.`——9 项没有一项覆盖字体字号封面，全绿证明不了
  "好看"；把工具能证明什么一句话说破。quiet 模式也不消失（--quiet 省的是装饰不是
  诚实）。测试 +2（test_validate_units.py::TestReportScopeStatement）。

- 用户自查的另一条（"14 页作业该不该用"的判断题）不加：SKILL「When to use / when
  not to」与快速开始「按文档阶段选流程」已覆盖，重复即冗余。

### policy.py 接入 validate 退出码（0.9.2 遗留的"未接线"挂起件落地）

`scripts/policy.py`（四级严重度模型）此前一直是 README 自述"尚未接入 validate 退出码"
的挂起模型。本轮按 0.6.4 时代写下的 Phase 2 计划接进门禁，**opt-in，缺省行为逐字不变**：

- **新增 `--acceptance-mode free|advisory|strict`**：策略只裁决 **profile 契约层**
  （`--enforce-profile` 编译出的 `profile.*` 断言）——`free` 全 observe 不拦、
  `advisory` 把契约 FAIL 降级为 WARN（亮黄灯，退出码 0）、`strict` 保持硬失败。
- **核心 9 项 + `profile.load` 是物理不变量**，任何模式 FAIL/ERROR 都退 1：策略无权
  把「包坏了」调成绿。实现在 `policy.evaluate` 的 `enforced_check_results`（在这些键里
  severity 不作数），名单由 `policy.core_invariants()` 给出。
- **ERROR 一律 fail-close**：FAIL 是「属性不符」（策略可按严格度降级），ERROR 是
  「无法判定」（检查崩了）——任何 severity 都无权把「不知道」降级成「合格」，只有
  `ignore`（明确声明不观测）例外。这是本仓「查不了当 PASS 是最大失真」铁律在策略
  模型里的对应约束。
- **降级不销毁证据**：被 observe/WARN 掩盖的失败原样留在 report 的 checks 里，同时记进
  `metadata.acceptance.observed_failures`；结论行 `Acceptance verdict: PASS/WARN/FAIL
  (mode: …)` 打到 stdout；`metadata.acceptance` 记录 `mode` / `verdict` /
  `renderers_required` / `observed_failures` / `warnings`。
- **renderer 矩阵不做成检查项**（诚实性取舍）：`acceptance.renderers` 如实记进
  `metadata.acceptance.renderers_required`——单次运行只跑一个渲染器，「多渲染器都通过」
  只能由编排层跨运行聚合，单次运行里做这个判定就是撒谎。
- 配套：`_run_profile_asserts` 的兜底 ERROR 检查名从 `profile_assert` 统一为
  `profile.load`（与 `_verify.compile_profile_checks` 打开失败时的命名一致）；
  `policy.core_invariants()` / `policy.mode_policy()` 新增；`AcceptanceSummary`
  新增 `observed_failures`。
- 测试 +12：`test_policy.py` 补 enforced_map 覆盖 / ERROR fail-close / 契约层不越界 /
  模式解析；`test_validate_units.py::TestAcceptancePolicyWiring` 十项（缺省不扰动 /
  advisory 降级 / strict 硬失败 / 核心不变量压策略 / ERROR fail-close / verdict↔退出码
  自洽 / renderers_required 落盘）。端到端实测：同一份契约层 4 FAIL 的文档，free→exit 0、
  advisory→exit 0（verdict WARN）、strict→exit 1、缺省→exit 1。
- 文档：`SCRIPT_HELP.md`（参数 + 示例 + §验收策略模式表）、`VALIDATION.md` /
  `VALIDATION.zh-CN.md`（§Acceptance modes）、README 双语仓库地图（policy.py 行改为
  已接线）、SKILL.md 验收节一句提及。  

## 0.9.2 — 2026-10-09 — 校验器自洽（grid 假 FAIL + 图片口径 + grid 自动对齐）

同一审计者的 skill 层审查：两条源码 bug（其一"咬了两次"）+ 流程分级建议，全部核实
成立。核心矛盾是**校验器与文档自相矛盾**：文档推荐 grid table，校验器却不认。

- **表格分隔行不再误判为正文（源码 bug，最严重）**：`_source_md_segments` 旧正则
  `[-=:\s]+` 漏了 `|` 和 `+`——宽管道表（≥5 列）分隔行归一化成 `:-:-:-:-:-`（≥8 字）
  假 FAIL；grid 表分隔行任何列数都假 FAIL，而 grid 恰是文档自己推荐的复杂表格方案
  （列宽/合并单元格/多级表头）。修在**原始行**上：`[|:+\-=\s]+` 一并剔除，含两种
  分隔行的端到端回归钉死。实测修复前 `validate --source-md` 对 grid 文档必炸。
- **图片解析口径与 render 统一（源码 bug）**：`_md_image_paths` 只查 md 目录和 cwd，
  render 能嵌进去的图（子目录/兄弟目录/父目录扩展）validate 定位不到，image_embedding
  被迫假降级 SKIP。现在：候选目录唯一实现 `_shared.iter_image_search_dirs`，render 的
  `--resource-path` 与 validate 的逐图定位共用；validate 新增 `--resource-paths` 显式
  补充。实测同一份"图在兄弟目录"的 md：render `images: 1/1 ok` + validate
  `逐图比对 1/1 身份与顺序一致`（修复前 SKIP）。docstring 里"不该重复实现一套"的
  承诺这次真兑现了。
- **grid table 渲染前自动对齐（设计 #5/#6 的合并解）**：pandoc 按显示宽度解析 grid，
  中文手写几乎必错位。`render.py` 合并阶段跑 `align_tables.process_text`（壳→壳复用，
  不搬代码）：错位表格自动重排后再交给 pandoc，源文件不写回，单元格内容逐字自校验，
  对不了的 `[warn]` 指路 `align_tables --fix`。pipe table 不受影响。至此"手写必错、
  靠外部工具兜底"的半残体验闭环：校验器认 grid（上一条）+ 渲染自动对齐（本条）。
- **流程分级写进文档（设计 #3，不加 `--light` 旗标）**：render 默认本就是秒级 docx-only，
  `--pdf`/`--check`/`validate` 全是显式 opt-in——缺的不是开关是引导。README 双语快速
  开始加"按文档阶段选流程"，SKILL 验收门槛节加"短文档的整个门禁 = 默认渲染 + 交付前
  一次 `--pdf --check`"。
- **`simple-report` 提到 config 文档顶部（设计 #4）**：CONFIG 双语开头加"三种常见起点"
  （正式 / 短文档 / 表单），按文档类型选配置不再埋在字段表中间。
- **报错可操作（#7）**：`源内容缺失` 的样例若形似分隔行残片，附"0.9.2 已修；若仍出现
  请附源 md 报 issue"提示。
- **测试 +7**：宽管道/grid 分隔行单测（含端到端 source_content PASS）、兄弟目录图片
  口径（`_md_image_paths` 与 check 级）、grid 自动对齐 e2e（源不动 + 单元格正确）、
  validate 兄弟目录逐图身份 PASS；端到端实证：含宽管道表 + 错位 grid + 兄弟目录图的
  一份 md，render 与 validate 结论首次完全一致。

## 0.9.1 — 2026-10-09 — 防呆三连 + pre-push Word 门禁（用户边界实测五项）

同一审计者实测大量边界场景 + 读核心代码，五项发现全部核实成立（另确认六项为好，
含无注入面、并发安全）。两条"用户踩了也不自知"的隐蔽问题优先修。

- **`--out` 撞源文件拒绝渲染（高，毁数据路径）**：`--src x.md --out x.md` 会把源
  Markdown 永久覆盖成 docx 二进制且无任何提示（实测确凿）。现在 merge 后、pandoc 前
  fail-fast：与 src 目录本身**或目录里任何一个 .md** 同路径都拒绝（目录合并场景），
  比较走 `normcase(realpath)`——Windows 大小写不敏感，`REPORT.MD` 同样拦下。
- **`style` 未知键 [warn]（高，静默失效）**：`{"style": {"body_ize": 16}}` 此前完全
  静默——顶层未知键有 warn（0.7.1），style 子键是盲区。现在 `apply_style_cfg` 对
  白名单之外的键发 `[warn]` 并指路键表；白名单 `KNOWN_STYLE_KEYS` 由测试钉死等于
  schema 的 style 键集合——将来加键漏进白名单会被未知键 warn 误伤，测试当场红。
- **pre-push Word 门禁**：新增 `pytest-word` 钩子（`stages: [pre-push]`，跑全量
  `python -m pytest -q`）。CI 没有 Word，word 标记用例只有本地能跑——这道钩子把
  "推送前手动跑全量"的约定变成强制。安装必须 `pre-commit install --hook-type
  pre-push`（README 开发节双语同步，并写明无 Word 机器自动 SKIP 标记用例）。
- **`raw_html` 文档化（中）**：pandoc reader 开着 `raw_html` 此前零文档；README 双语
  「输入要求」与 SKILL 陷阱清单补上——源 md 里的 HTML 原样进 docx（契约不改内容，
  不过滤），不可信来源先过目。默认保持开启（关掉会破坏插 HTML 结构的既有用法）。
- **LibreOffice 域写回进交付行（低）**：LO 是唯一不能把刷新后的域写回 docx 的渲染器
  （Word/WPS 走 COM 不受限），此前只缩在步骤日志的 `[warn]` 里。现在 `✓ 完成` 行在
  LO 渲染时追加"交付 docx 前请用 Word/WPS 打开刷新"。
- **`--out` 覆盖语义写进帮助（低，行为维持不变）**：同名覆盖是渲染器的幂等常态
  （迭代/CI 重跑每次命中），不做运行时警告——那会在每次正常重渲染时制造噪音；
  `--help` 与 SCRIPT_HELP 写明"同名文件会被覆盖"。顺带修正 README 双语里 0.8.0 后
  已过时的"第一个 # 之前"描述（YAML 标题块保留居中是新行为）。
- **测试 +6**：撞源守卫三边界（同路径 / 目录内 md / Windows 大小写，后者 nt 专属）、
  style 未知键 warn + 白名单 schema 守卫、LO 交付行提示双向断言。

## 0.9.0 — 2026-10-09 — 终端 UX 转向「用户进度」（外部 UX 审查六项）

用户拿真实文档跑通后给出六项 UX 审查（基于实际跑通 + 读代码，未改文件），全部核实
成立。主题：把工具内部的验收/清理/管线流程从「终端提问」和「日志前缀」里收回来，
最后一行回答用户真正想问的——成功了吗、文件在哪、几页。

- **渲染后不再弹「验收确认」提问（高）**：TTY 下此前必问
  「验收确认无需返修吗？[y/回车]」——设计哲学（人工验收可返修）以提问形式泄漏到
  每次运行的必经路上。现在人在终端：**静默保留 + 一行提示**（位置 / 24h 自动回收 /
  `--discard-work` 立即删），返修能力不变（24h 窗口 + `cleanup_old_temp` 兜底磁盘，
  失败路径依旧一律回收）；非交互维持跑完即回收。`_ask()` 整个删除。
- **最后一行改为用户视角的主进度（高）**：新增 `✓ 完成：x.docx + x.pdf（N 页，
  可打开 PDF 预览）`——页数读自导出的 PDF 本身；只出 docx 时报文件大小并提示如何
  加 `--pdf --check`（原 `[hint]` 行并入）。`[check]` 截图去向两行并进主进度
  （`--check --keep-pages` 时附截图路径），截图「默认移走」的行为不变（R1 反馈立的
  规矩）。`[1/3]…[3/3]` 步骤行保留——三行可 grep，调试价值大于噪音。
- **`--help` 从裸选项补齐为可自助（中）**：`--src/--out/--config/--pdf/--check`
  此前一行帮助都没有；现在每个选项一句话（单文件还是目录、可省、需本机渲染器、
  自动启用 `--pdf`），epilog 附三条最短可跑命令（`RawDescriptionHelpFormatter`）。
- **`--pdf` 隐性依赖：补说明、不做静默回落（中）**：渲染器身份是验收契约的一部分
  （evidence 记录渲染器、跨渲染器版式有差异、snapshot 基线绑定机器+渲染器），
  缺 Word 时维持显式报错 + 替代方案指路，不自动换渲染器。
- **中间产物概念退到一行（中）+ 截图位置说明（低）**：由前两条吸收；`[work]` 静默
  保留行收敛为一行，详细返修指引只在显式 `--keep-work` 时打印。
- **测试**：提问流测试改写为「不再提问」断言（注入 `input` 即炸），新增主进度行 /
  `--help` 内容 / `_page_count` 单测；`--keep-work` 可复用性与非交互回收两条既有
  契约测试原样通过（行为未变）。SCRIPT_HELP「人工验收确认」章节改写为「中间产物
  的去留」并附输出示例。

