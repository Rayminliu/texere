# VALIDATION.zh-CN.md

> 从 `README.zh-CN.md` 下沉而来的明细。中英两份由 `tests/test_docs_sync.py` 守着同构。

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
└── signature            # 校验清单：docx + report + 全部截图的哈希、检查摘要
```

这个文件叫 `signature` 是历史原因，但它是**校验清单（checksum manifest），不是密码学签名**——
没有密钥，任何人都能重算这些哈希。它证明的是「这份证据描述的是哪个产物」，
不是「这份证据没被改过」。

报告含 9 项自动检查：

| 检查项 | 验的是什么 |
|---|---|
| ✅ 包完整性 | docx 是含必需部件的合法 ZIP |
| ✅ 源内容 | `--source-md`：Markdown 各段是否都出现在 docx 正文里；`--expected-hash`：docx 文件级 SHA256（只能证明字节未变）；两者都不给 → SKIP |
| ✅ 图片嵌入 | 给了 `--source-md`：逐图 SHA256 身份 + 文档顺序校验（抓串位 / 错图 / 同一张图重复占位；pandoc 原样嵌入字节）。没给：仍是嵌入数 ≥ 引用数的下限计数 |
| ✅ 分节数 | 分节数量合理（1–100） |
| ✅ 目录域 | OOXML 里存在真实的 `TOC` 域（文档本就没有目录 → SKIP） |
| ✅ 页码 | 页脚页码构成无缺口序列（识别不出页码格式 → SKIP） |
| ✅ 空白页 | 不超阈值（默认允许 0 个） |
| ✅ 渲染器验收 | 所选渲染器能打开并成功导出 PDF |
| ✅ 版式基线漂移 | 与基线逐页逐像素比对，默认全量 |

**四种状态，SKIP 不等于 PASS。** 每项检查报 `PASS` / `FAIL` / `SKIP` / `ERROR` 之一。
`SKIP` 表示前置条件缺失（没给基线、没给 `--source-md`、没装 PyMuPDF），这项**根本没查**；
它不计入通过数，但也不单独让门禁失败。只有 `FAIL` 与 `ERROR` 会让退出码变成 1。
`Passed: 7/9 (skipped: 2)` 与 `Passed: 9/9` 是分量完全不同的两句话，报告里会分开写出来。

示例输出：

```
Document Validation
────────────────────────────
✅ [PASS] package_integrity: OK
⏭️ [SKIP] source_content: 跳过 (未提供 --source-md 或 --expected-hash)
✅ [PASS] image_embedding: 图片逐图比对：28/28 张身份与顺序一致
✅ [PASS] section_count: 分节数：3 (合理)
✅ [PASS] toc_field: 目录域：1 个 TOC 域
✅ [PASS] page_numbering: 页码：69 页 (连续，检测到页码 1-69)
✅ [PASS] blank_pages: 空白页：0/69 (阈值：0)
✅ [PASS] renderer_acceptance: 渲染器验收：OK
⏭️ [SKIP] visual_drift: 跳过 (未提供基线目录；如需版面漂移防护：先 python scripts/snapshot.py <pdf> --update 录基线，再用 --baseline <目录> 或 profile 的 baseline_dir 指定)

Summary
────────────────────────────
Passed: 7/9 (skipped: 2)

✅ All checks passed

Evidence package saved to: evidence/
  - report.json (structured validation report)
  - page-XXX.png (sample screenshots)
  - signature (校验清单：docx + report + 全部截图的哈希)
```

任何一项失败即 exit code 1，并给出详细错误。参数（`--profile` / `--max-empty` / `--quiet` 等）见
`docs/SCRIPT_HELP.md` §validate.py。

### Profile 契约边界（profile=验收契约 / config=渲染输入）

profile 与渲染 config 是**两个不同的对象**，刻意不打通：

- `profiles/*.json` 是**验收契约**：描述成品文档必须满足什么。只被 `validate.py` 消费
  （`--profile` 把它附进证据报告；`--enforce-profile` 把其声明字段编译成 `profile.<field>` 门禁断言）。
- `config.json` 的 `style.*` 是**渲染输入**：`post.py` 真正消费它来排版。
- 把 profile 注入渲染是**明确不做的**（non-goal）：同一套排版旋钮出现契约与输入两个事实源
  会各自漂移，且让门禁结论变得不可解释。

门禁实际强制哪些字段（细节见 `docs/SCRIPT_HELP.md` §Profile）：

| 类别 | `--enforce-profile` 时强制 | 仅声明（不断言） |
|---|---|---|
| 页面 | `page.width/height`（±0.1cm）、`page.margin_*`（±0.2cm） | — |
| 正文字体 | `styles.body.font_eastAsia` / `font_latin` / `size`（±0.5pt） | `line_spacing`、`first_line_indent`、`space_*` |
| 标题字体 | `styles.h1/h2/h3` 中/西文字体、字号、加粗 | `page_break_before`、`space_*` |
| 表边框 | `table.border` ≠ none → 至少一个可见边框 | `border_size`、`border_color`、`header_shade`、`zebra` 等 |
| 目录 | 声明了 `toc` 段 → TOC 域存在 | `toc.depth`、`toc.title`、`placeholder` 等 |
| 其余 | — | `caption.*`、`header.*`、`footer.*`、`*_specific.*` |

可选键（任何 profile 都可写；5 个内置 profile 均未使用）：`baseline_dir`（未传 `--baseline` 时
视觉漂移检查回退读它）与 `acceptance.renderers`（多渲染器验收矩阵，如 `["word", "wps"]`；
由 `policy.py` 消费，未接入 validate 退出码）。

### 人工门槛

9 项检查管的是结构，头一遍过某份新文档时，这四件事仍然得靠人：

1. 日志出现 `images: n/m ok` 且 **n == m**（m 是源 md 中 `![` 的次数）
2. `near-empty pages: 0`
3. `OK`（Word 成功打开并导出 PDF）
4. **人工过一遍 PDF 渲染图**——机器能查页数、图片数、空白页，
   查不了"这张图画得对不对、表头有没有被截断"

## 能力模块与测试分层

流水线的纯逻辑住在 `scripts/` 下可导入的平铺模块里，面向用户的脚本保持为 thin CLI 壳。这样拆，
是为了让逻辑能在进程内被直接测到，而不必为每次调用付出启动 Word 的代价：

| 能力模块 | 管什么 | re-export 它的壳 |
|---|---|---|
| `_verify.py` | 检查断言内核 + `CheckResult` / 状态常量 | `validate.py` |
| `_visual_diff.py` | `diff_ratio`、抽样页号、`compute_visual_diff` | `validate.py`、`snapshot.py` |
| `_evidence.py` | 报告骨架、计数、`report.json` / signature 写盘 | `validate.py` |
| `_compile.py` | pandoc 命令构造 + 代码围栏 / 交互提示助手 | `render.py` |
| `_mutate.py` | 声明式 Patch 引擎 + `MutationResult` 操作原语 | `patch.py`、`edit.py` |

每个壳导入的正是它原本自己定义的那些同名符号，所以 `stdout` 措辞、`argparse` 选项与退出码一字未变
——re-export 本身就是等价性证明：既有的 subprocess 与进程内用例不改一行即全绿。依赖是单向 DAG
（`_version → _shared → {能力模块} → 壳`）；能力模块绝不反向 import 它的壳。

测试按分层组织，只在必须的地方跨过那道慢边界：

- **L1 / L2 — 进程内单元 + 逻辑**（`test_validate_units.py`、`test_patch.py::TestPatchUnits`）：直接对内存
  docx / PyMuPDF fixture 调用 re-export 的函数。不启 Word，亚秒级。
- **L3 — CLI 契约**（subprocess）：断言被冻结的 `stdout` 措辞、`--help`、`report.json` 结构与退出码；需要
  真渲染器启动的打 `@pytest.mark.word`。
- **L4 — 端到端**：对一份 golden 文档做验收。

CI 把打了 `word` 标记的用例单独拆步跑（`pytest -m "not word"` 再 `-m word`），于是纯逻辑回归保持轻量，
而碰版式的改动仍以 0.00% 像素快照作为硬门禁。

