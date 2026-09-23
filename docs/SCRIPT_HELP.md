# Texere 脚本帮助文档

## render.py - Markdown 渲染为正式文档

### 用法
```bash
python scripts/render.py [OPTIONS]
```

### 示例
```bash
# 环境自检
python scripts/render.py --doctor

# 冒烟测试（使用样本文档）
python scripts/render.py --sample

# 渲染真实文档
python scripts/render.py \
  --src chapters/ \
  --out bid.docx \
  --config config.json \
  --pdf \
  --check

# 指定版本
python scripts/render.py --version
```

### 参数说明
- `--src <dir|file.md>`: Markdown 源（目录按文件名序合并，也可直接给单个 .md 文件）（必需）
- `--out <file>`: 输出 DOCX 文件路径（必需）
- `--config <file>`: 配置文件路径（可选）
- `--pdf`: 生成 PDF 格式（默认用本机 Word 渲染器，可用 `--renderer` 切换为 libreoffice / wps）
- `--check`: 检查 PDF 视觉质量（需要 PyMuPDF）
- `--renderer <word|libreoffice|wps>`: PDF 导出渲染器（默认 word）
- `--doctor`: 环境自检
- `--sample`: 运行样本文档测试
- `--version`: 显示版本号

---

## post.py - 文档后处理

### 用法
```bash
python scripts/post.py <body.docx> <out.docx> [config.json]
```

### 功能
- 注入封面和目录
- 设置页眉页脚
- 表格样式化
- 题注居中格式化

---

## validate.py - 文档验证器

### 用法
```bash
python scripts/validate.py <document.docx> [OPTIONS]
```

### 示例
```bash
# 完整验证
python scripts/validate.py bid.docx

# 生成证据包
python scripts/validate.py bid.docx --out evidence/

# 自定义空白页阈值
python scripts/validate.py bid.docx --max-empty 2

# 使用 profile 进行视觉基线比对
python scripts/validate.py bid.docx --profile profiles/formal-cn-v1.json

# 把 profile 当成可执行契约强制执行（profile.<field> 计入门禁）
python scripts/validate.py bid.docx --profile profiles/formal-cn-v1.json --enforce-profile

# 静默模式（仅输出摘要）
python scripts/validate.py bid.docx --quiet
```

### 参数说明
- `<document.docx>`: 待验证的文档（必需）
- `--out <dir>`: 证据包输出目录（默认：evidence）
- `--profile <file>`: Profile JSON，附进证据包报告（可声明 baseline_dir）
- `--baseline <dir>`: 视觉基线目录（页图 p001.png…，与 snapshot.py 同口径 dpi=100 比对）
- `--expected-hash <hash>`: 期望的 docx 文件 SHA256（产件级，只证明字节未变）
- `--source-md <file>`: 源 Markdown，开启正文等价性比对（比 hash 强得多，见下）
- `--sample-visual`: 视觉比对只比首 / 中 / 尾三页（默认逐页全量）
- `--max-empty <n>`: 允许的最大空白页数（默认：0）
- `--enforce-profile`: 把 profile 里声明的页面/字体/标题/表格/目录编译成硬断言
  （profile 即文档规范；默认只把 profile 当 baseline 选择器，不强制）
- `--renderer <word|libreoffice|wps>`: PDF 导出渲染器（默认 word；PDF 相关验收随之切换）
- `--quiet`: 静默模式

### 验证项（9 项）

每项返回 `PASS` / `FAIL` / `SKIP` / `ERROR` 之一：

- `SKIP` = 前置条件缺失，这项**没查**，不计入通过数，也不单独让门禁失败。
- 只有 `FAIL` / `ERROR` 会让退出码变成 1。
- 摘要行形如 `Passed: 7/9 (skipped: 2)`，别把 skipped 当通过。

1. ✅ Package integrity - DOCX 包结构完整性
2. ✅ Source content - 给了 `--source-md` 就做 Markdown↔docx 正文比对；只给
   `--expected-hash` 就退化为产件文件级 hash；两者都无 → SKIP
3. ✅ Image embedding - 给了 `--source-md`：逐图 SHA256 身份 + 文档顺序校验（pandoc 原样
   嵌入字节，实测 sha 一致；能抓串位 / 错图 / 重复占位，且含浮动图）。没给：嵌入数 ≥
   引用数的下限计数。无 Markdown 引用 → SKIP
4. ✅ Section count - 分节数合理性
5. ✅ TOC field - OOXML 里是否存在真实 `TOC` 域；文档本就没有目录 → SKIP
6. ✅ Page numbering - 页码连续性；识别不出页脚页码格式 → SKIP
7. ✅ Blank pages - 空白页数量
8. ✅ Renderer acceptance - 渲染器真机验收（随 --renderer：Word / WPS / LibreOffice）
9. ✅ Visual drift - 与基线逐页比对（默认全量；`--sample-visual` 才抽样）；

### Profile 作为可执行契约（`--enforce-profile`）

`--profile` 默认只把 profile 当 baseline 选择器；加 `--enforce-profile` 时，profile 里声明
的字段会被编译成额外断言，命名 `profile.<field>`，独立计入报告与门禁（任一 FAIL → 退出码 1）：

- `profile.page` - 页面尺寸（A4）+ 上下左右边距，与 profile.page 对齐（容差 0.1–0.2cm）
- `profile.body_font` - 正文 Normal 样式的中/西文字体 + 字号（容差 0.5pt）
- `profile.heading` - 标题 1/2/3 的中/西文字体 + 字号 + 加粗
- `profile.table` - 表格是否含可见边框（profile.table.border ≠ none 时）
- `profile.toc` - 文档是否存在 TOC 域（profile.toc 声明时）

这些断言走 OOXML 读取，**不依赖 Word**；只覆盖 profile 真正声明的字段，未声明的不凭空编造。
未传 `--profile` 时这条链路完全不出现，核心 9 项检查不受任何影响。

#### 检查语义边界（Known limitations，只记录不实现）

- **`profile.page` 只看 `doc.sections[0]`（first_section）**：多节文档（如 cover=A4、body=A3
  横版）不会逐节比对，只校验主 section。若 profile 合同需要多节，需显式扩展 scope 语义。
- **`profile.table` 是「至少一个表含可见边框」的粗粒度断言**：不逐边（top/left/right/
  bottom/insideH/insideV）校验颜色或粗细；evidence 里 `rule` 固定为 `at_least_one_visible_border`，
  防止误读成「所有表全部符合边框规范」。
- **`profile.<field>` 的 `evidence` 只是观测记录**：report.json 里每个 profile 检查都带
  `evidence`（机器可读的 `expected` / `actual` / `field` / `source`），但它**不参与 status 判定**，
  缺失 evidence 也绝不改变 verdict——仅用于可解释审计与将来的 Build Manifest 直接消费。
- 下列 profile 字段**目前仅声明、未强制验证**（declared-only，非 enforced）：`line_spacing`、
  `first_line_indent`、`space_*`、`caption.*`、`header.*`、`footer.*`、`tender_specific.*`。

#### Profile 字段执行状态（声明 ≠ 一定检查）

为避免「JSON 写了就以为 validator 会保护我」，逐项标注每个字段的实际执行状态。**缺失即 FAIL**
指 profile 要求该字段时，文档实际没声明也会判 FAIL（不只是「声明了但不符」才 FAIL）。

| Profile 字段 | 状态 | 说明 |
| --- | --- | --- |
| `page.width` / `page.height` | ✅ enforced | 与 A4（21×29.7cm）等比对，容差 0.1cm |
| `page.margin_*` | ✅ enforced | 与 profile 比对，容差 0.2cm；**缺失或不符都 FAIL** |
| `styles.body.font_eastAsia` / `font_latin` / `size` | ✅ enforced | 缺失或不符都 FAIL（容差 0.5pt） |
| `styles.h1/h2/h3.font_eastAsia` / `font_latin` / `size` / `bold` | ✅ enforced | 缺失或不符都 FAIL |
| `table.border` | ✅ enforced | ≠ `none` 时检查可见边框（真实 OOXML `w:top/w:left/...`） |
| `toc` | ✅ enforced | 声明即检查 TOC 域存在 |
| `styles.body.line_spacing` / `first_line_indent` / `space_*` | ⚠️ declared-only | profile 可声明，目前**未**编译成断言 |
| `styles.h1/h2/h3.page_break_before` / `space_*` | ⚠️ declared-only | profile 可声明，目前**未**编译成断言 |
| `caption.*` / `header.*` / `footer.*` | ⚠️ declared-only | 渲染指令，未编译成断言 |
| `table` 其余键（border_size / border_color / header_shade / zebra / ...） | ⚠️ declared-only | 渲染指令，未编译成断言 |
| `toc.depth` / `toc.title` / ... | ⚠️ declared-only | 只检查 TOC 域存在，不校验深度 / 标题 |
| `tender_specific.*`（signature_page / sealing_requirement / price_table_style / ...） | ⚠️ declared-only | 客户语义信息，目前无对应 validator |
   无基线目录 → SKIP

---

## patch.py - 声明式文档编辑 API

### 用法
```bash
python scripts/patch.py <document.docx> <patch.json> [OPTIONS]
```

### 示例
```bash
# Dry run 模拟执行
python scripts/patch.py doc.docx patch.json --dry-run

# 应用 Patch（写回原文件，自动备份 .bak.docx）
python scripts/patch.py doc.docx patch.json --apply

# 应用并验证结果
python scripts/patch.py doc.docx patch.json --apply --validate

# 应用并另存 + 生成证据包
python scripts/patch.py doc.docx patch.json --apply --out result.docx
python scripts/patch.py doc.docx patch.json --apply --out evidence/
```

### 参数说明
- `--dry-run`: 模拟执行不写盘（深拷贝文档跑一遍全部操作）
- `--apply`: 实际执行
- `--validate`: 应用后逐项验证结果
- `--out <路径>`: 另存为文件或证据包输出目录（缺省写回原文件）
- `--no-backup`: 写回原文件时不创建 `.bak.docx` 备份

### Patch Schema
```json
{
  "id": "patch-001",
  "description": "修改工期从 90 天到 120 天",
  "preconditions": {
    "must_contain": ["90 日历天"]
  },
  "operations": [
    {
      "op": "replace_text",
      "target": {"paragraph": 42},
      "expected_old_text": "90 日历天",
      "new_text": "120 日历天"
    }
  ]
}
```

### 支持的 Operation
- `replace_text` - 替换文本
- `insert_after` - 在锚点后插入
- `insert_before` - 在锚点前插入
- `delete_paragraph` - 删除段落
- `set_cell` - 设置单元格值
- `add_row` - 添加表格行；`target.after_row`（0 基）指定插在哪一行之后，
  省略或 `-1` 表示追加到表尾。新行克隆该行的边框/底纹/字号，文字清空后按 `values` 写入
- `del_row` - 删除表格行

---

## edit.py - 定点文档编辑

### 用法
```bash
python scripts/edit.py <document.docx> [OPTIONS]
```

### 示例
```bash
# 查看文档结构
python scripts/edit.py 标书.docx --list

# 批量替换文本
python scripts/edit.py 标书.docx \
  --replace "90 日历天=120 日历天" \
  --replace "A 公司=B 公司"

# 在锚点后插入段落
python scripts/edit.py 标书.docx \
  --after "资质要求" \
  --text "新增条款 1\n新增条款 2"

# 修改表格单元格
python scripts/edit.py 标书.docx \
  --cell 0 2 1 "1,060,000"

# 添加表格行
python scripts/edit.py 标书.docx \
  --add-row 0 "接入层" "设备" "320,000"

# 批量加空行（复制模板行的边框/底纹/字号格式，文字清空）
python scripts/edit.py 标书.docx \
  --add-rows 0 3 --template-row 2

# 批量填表（JSON：null 跳过不清空；CSV：空单元格跳过）
python scripts/edit.py 标书.docx --fill data.json
python scripts/edit.py 标书.docx --fill rows.csv

# 修改页眉
python scripts/edit.py 标书.docx \
  --header "新版页眉" \
  --section body

# 验证修改后的文档
python scripts/edit.py 标书.docx \
  --replace "旧=新" \
  --verify
```

### 参数说明
- `--list`: 列出文档结构（章节、表格、段落）
- `--replace "旧=新"`: 文本替换（可多次指定）
- `--scope <范围>`: 替换作用范围（body,tables,header,footer）
- `--after <锚点>`: 在锚点后插入
- `--before <锚点>`: 在锚点前插入
- `--text <内容>`: 配合 insert 操作
- `--style <样式名>`: 新段落样式
- `--delete <文字>`: 删除包含该文字的段落
- `--all-anchors`: 对多个锚点全部生效
- `--cell <表> <行> <列> <值>`: 设置单元格
- `--add-row <表> <值1> <值2> ...`: 添加行并写入值
- `--add-rows <表> <数量> [--template-row <行>]`: 批量加空行，复制模板行的全部格式（默认复制最后一行）
- `--fill <data.json|data.csv>`: 批量填表，可多次指定。JSON 格式：`[{"table":0,"start_row":1,"start_col":0,"values":[[...]]}]`；合并单元格按「写主格、跳过后续坐标」处理；越界只警告
- `--del-row <表> <行>`: 删除行
- `--header <文字>`: 修改页眉
- `--footer <文字>`: 修改页脚
- `--section <范围>`: 页眉页脚作用范围（body/all/序号）
- `--out <文件>`: 另存为（默认写回原文件）
- `--no-backup`: 不创建备份
- `--verify`: 修改后用 Word 打开验证

---

## distill.py - 模板蒸馏工具

### 用法
```bash
python scripts/distill.py <template.docx> [OPTIONS]
```

### 示例
```bash
# 打印报告和建议配置
python scripts/distill.py 甲方模板.docx

# 保存建议配置到文件
python scripts/distill.py 甲方模板.docx --out cfg.json
```

### 输出内容
- 页面设置（纸型、页边距、分节数）
- 样式信息（字体、字号）
- 页眉页脚内容
- 建议的 config.json
- 需人工确认项提示

---

## finalize.py - Word 验收和 PDF 导出

### 用法
```bash
python scripts/finalize.py <in.docx> [out.pdf]
```

out.pdf 可省略，默认取输入同名 `.pdf`（长中文名不用再手打）。

### 功能
- 调用 `scripts/renderers.py` 里的 `WordRenderer`（实现统一 `RendererAdapter` 契约）
  完成 Word COM 验收 + 导出 PDF；`finalize.py` 现在只是薄 CLI 壳，**行为不变**
- 验收信号：Word 打不开 = OOXML 结构有问题（立即失败）
- 默认**只读验收**：在临时副本上刷新目录域 / 重排 / 导 PDF，绝不写回输入 docx；
  只有 `--save-updated-fields` 才把刷新后的域写回原文件（render --pdf 交付物需要）
- 输出：页数 / 字数 / 表数 / 图数 / 节数

### 渲染器抽象（Core vs Renderer）
Texere 核心（compile / OOXML / source / image / profile / metadata / evidence）
**不依赖任何 Office**。只有「真机验收 + 出 PDF」需要具体渲染器，这一层就是
`RendererAdapter`：`WordRenderer`（当前默认）/ `LibreOfficeRenderer`（soffice headless，
跨平台 CI 友好）/ `WPSRenderer`（WPS Writer COM）——三者只是不同的「事实渲染器」。
`page_numbering` / `blank_pages` / `visual_drift` 这些 PDF 派生检查只消费 PDF、绝不感知
渲染器——换渲染器不会改变它们的结论，这正是 renderer-agnostic 的护栏
（见 `tests/test_renderer.py`）。

统一入口是 `get_renderer(name)`（name ∈ `word` / `libreoffice` / `wps`），每个渲染器
自带 `available()` 能力探测。`render.py` 与 `validate.py` 都暴露 `--renderer` 选项；
不指定时默认 `word`。`finalize.py` 现在只是 `WordRenderer` 的薄 CLI 壳，仍可单独调用：

```bash
python scripts/finalize.py <doc.docx> <out.pdf> [--save-updated-fields]
```

#### Renderer 已知边界（只记录不实现）

- **COM 渲染器（Word / WPS）的超时是 best-effort**：`WordRenderer` / `WPSRenderer`
  在进程内跑 COM，线程级超时（`validate.export_pdf_once` 的 `ThreadPoolExecutor`）
  能返回结构化错误，但**无法硬杀卡死的 COM 进程**——彻底消 zombie 需要进程级隔离
  （把渲染放进可被 `terminate` 的子进程）。当前 `validate` 走的是 WordRenderer 默认路径，
  单次 `validate` 通常足够；高并发 / 长文档场景若遇僵尸，优先用 LibreOffice 后端
  （子进程，超时会被强杀，见 `renderers._kill_process_tree`）。
- **`LibreOfficeRenderer` 不回写刷新后的域**：`save_updated_fields=True` 仅 Word / WPS
  支持；LO 出 PDF 后会带 warning，交付 docx 请仍走 Word / WPS。
- **跨渲染器视觉一致性未实测**：同一份 docx 经 Word / LO / WPS 出的 PDF 在字体、分页、
  页眉页脚上可能有差异，尚未建立 compatibility corpus（见 roadmap ⑥）。
- **Word / WPS COM 退出期回溯（pythoncom atexit）**：本进程内只要跑过任何 Word / WPS
  渲染，解释器退出期都可能打印一条 COM teardown 访问违规回溯（stderr 噪音）；
  **pytest 退出码仍为 0，不影响判定**，是进程级隔离要根除的目标。并发 render 会更响，
  故 `test_word_concurrent_renders_isolated` 默认跳过，需 `TEXERE_COM_CONCURRENCY=1`
  显式开启才能复现 / 探测。

---

## snapshot.py - 视觉基线比对

### 用法
```bash
python scripts/snapshot.py <document.pdf> [OPTIONS]
```

### 示例
```bash
# 首次记录基线（写入 <PDF 同目录>/baselines/）
python scripts/snapshot.py output.pdf --update

# 回归比对（漂移 exit 1）
python scripts/snapshot.py output.pdf

# 更高精度比对（更慢更敏感）
python scripts/snapshot.py output.pdf --dpi 150 --max-diff 0.0005
```

### 参数说明
- `<document.pdf>`: PDF 文件（必需）；基线目录固定为其同级的 `baselines/`，无命令行参数可改
- `--update`: 录制/更新基线图片与 meta.json
- `--dpi <n>`: 渲染精度（默认 100；与基线 meta 不一致时拒比对）
- `--max-diff <r>`: 差异比例阈值（默认 0.001 = 0.1%；实测同文档重复导出 0.00%，改一处页眉 0.16%）

---

## make_ref.py - 重建参考模板

### 用法
```bash
python scripts/make_ref.py [OPTIONS]
```

### 示例
```bash
# 使用默认设置重建
python scripts/make_ref.py

# 指定正文字体和字号
python scripts/make_ref.py --body-font 楷体 --body-size 14

# 指定标题字体
python scripts/make_ref.py --heading-font 黑体

# 指定西文字体 / 输入输出模板
python scripts/make_ref.py --latin-font Georgia --src base.docx --dst out_ref.docx
```

### 参数说明
- `--body-font <字体>`: 正文中文字体（默认 宋体）
- `--latin-font <字体>`: 西文字体（默认 Times New Roman）
- `--heading-font <字体>`: 标题中文字体（默认 黑体）
- `--body-size <pt>`: 正文字号（默认 12，小四）
- `--src <docx>`: 基准模板（默认 assets/ref_default.docx，缺失时由 pandoc 生成）
- `--dst <docx>`: 输出模板（默认覆盖 assets/ref.docx）

---

## check_pdf.py - PDF 质量检查

### 用法
```bash
python scripts/check_pdf.py <document.pdf> [OPTIONS]
```

### 功能
- 检测空白页
- 渲染页面为 PNG 截图
- 返回非零退出码如果有问题

---

## 常见问题

### Q: 如何添加新的配置字段？
A: 在 `post.py` 的 `S` 字典中添加默认值，在 `apply_style_cfg()` 中解析。

### Q: 如何调试渲染问题？
A: 
1. 先用 `--doctor` 检查环境
2. 用 `--sample` 测试基本功能
3. 检查 `scripts/` 子进程输出中的警告信息

### Q: 如何扩展验证项？
A: 在 `validate.py` 中添加新的 `check_xxx()` 函数，并在 `generate_evidence_package()` 中注册。

### Q: Patch API 如何设计新的操作？
A:
1. 实现 `xxx_op()` 函数
2. 在 `OPERATIONS` 字典中注册
3. 在 `assess_patch()` 中添加静态检查逻辑
