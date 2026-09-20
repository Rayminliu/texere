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
- `--src <dir>`: Markdown 源文件目录（必需）
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
- `--profile <file>`: Profile JSON 用于视觉基线比对
- `--expected-hash <hash>`: 期望的文件 SHA256 hash
- `--max-empty <n>`: 允许的最大空白页数（默认：0）
- `--quiet`: 静默模式

### 验证项（9 项）
1. ✅ Package integrity - DOCX 包结构完整性
2. ✅ Source content - 内容完整性（可选 hash 校验）
3. ✅ Image embedding - 图片嵌入检查
4. ✅ Section count - 分节数合理性
5. ✅ TOC field - 目录域存在性
6. ✅ Page numbering - 页码连续性
7. ✅ Blank pages - 空白页数量
8. ✅ Word acceptance - Word 真机验收
9. ✅ Visual drift - 视觉基线比对

---

## patch.py - 声明式文档编辑 API

### 用法
```bash
python scripts/patch.py <document.docx> <patch.json> [OPTIONS]
```

### 示例
```bash
# Dry run 模拟执行
python scripts/patch.py doc.pdf patch.json --dry-run

# 应用 Patch
python scripts/patch.py doc.pdf patch.json --apply

# 应用并验证结果
python scripts/patch.py doc.pdf patch.json --apply --validate

# 应用并生成证据包
python scripts/patch.py doc.pdf patch.json --apply --out evidence/
```

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
- `add_row` - 添加表格行
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
- `--add-row <表> <值1> <值2> ...`: 添加行
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
python scripts/finalize.py <in.docx> <out.pdf>
```

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
# 首次记录基线
python scripts/snapshot.py output.pdf --update

# 回归比对
python scripts/snapshot.py output.pdf

# 静默模式
python scripts/snapshot.py output.pdf --quiet
```

### 参数说明
- `<document.pdf>`: PDF 文件（必需）
- `--update`: 更新基线图片
- `--baseline <dir>`: 基线目录（默认：baselines/）
- `--quiet`: 静默模式
- `--threshold <n>`: 漂移阈值（默认：0.001 = 0.1%）

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
```

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
