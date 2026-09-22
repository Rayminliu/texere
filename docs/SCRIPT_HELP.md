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
- `--pdf`: 生成 PDF 格式（需要 Microsoft Word）
- `--check`: 检查 PDF 视觉质量（需要 PyMuPDF）
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
8. ✅ Word acceptance - Word 真机验收
9. ✅ Visual drift - 与基线逐页比对（默认全量；`--sample-visual` 才抽样）；
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
- Word COM 打开文档（验证结构完整性）
- 刷新目录域
- 重新分页
- 导出为 PDF
- 保存并关闭文档

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
