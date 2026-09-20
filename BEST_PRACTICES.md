# texere 最佳实践指南

本文档提供 texere 在实际项目中的使用经验和最佳实践。

## 📚 目录

1. [Profile 选择指南](#profile-选择指南)
2. [Markdown 写作规范](#markdown-写作规范)
3. [表格样式最佳实践](#表格样式最佳实践)
4. [图片处理技巧](#图片处理技巧)
5. [验证与调试](#验证与调试)
6. [常见问题解决](#常见问题解决)

---

## Profile 选择指南

### 场景匹配表

| 文档类型 | 推荐 Profile | 关键特性 |
|---------|-------------|---------|
| **通用正式文档** | `formal-cn-v1.json` | 宋体正文、黑体标题、三线表 |
| **投标文件** | `tender-v1.json` | 报价表样式、骑缝章要求 |
| **政府公文** | `gongwen-v1.json` | GB/T 9704-2012、仿宋字体 |
| **项目申报书** | `application-v1.json` | 表单式表格、预算表 |

### 自定义 Profile

如果默认 profile 不满足需求，可以：

```bash
# 1. 蒸馏现有模板，得到 config.json
python scripts/distill.py 甲方模板.docx --out my-config.json

# 2. 微调配置
# 编辑 my-config.json，在 style 段调整字体/间距等参数

# 3. 使用自定义配置渲染（渲染链路只认 --config）
python scripts/render.py --src . --out result.docx \
  --config my-config.json --pdf --check

# 4. 验收时可用 --profile 把设计契约附进证据包
python scripts/validate.py result.docx --profile profiles/formal-cn-v1.json
```

> **注意**：`profiles/*.json` 是设计契约文档（供人审阅和 validate 报告引用），
> 渲染参数请直接写进 `--config` 的 `style` 段；`render.py` 不支持 `--profile` 参数。

---

## Markdown 写作规范

### 章节结构

```markdown
# 第一章 标题（一级标题）

## 1.1 小节标题（二级标题）

### 1.1.1 子小节（三级标题）

正文内容...
```

### 表格写法

**简单表格** - 使用 pipe table:

```markdown
| 列 1 | 列 2 | 列 3 |
|:----|:----:|-----:|
| 左对齐 | 居中 | 右对齐 |
```

**复杂表格** - 使用 grid table:

```markdown
+--------+---------+--------+---------+
| 商务部分   | 技术部分   |
+--------+---------+--------+---------+
| 条款   | 响应    | 模块   | 说明   |
+========+=========+========+=========+
| 工期   | 完全响应 | 接入层 | 设备   |
+--------+---------+--------+---------+
```

### 图片引用

```markdown
![图 1-1 图片标题](images/photo.png){width=13cm}
```

**注意**:
- 图片路径相对 Markdown 文件
- 宽度单位用 cm（厘米）
- 图片文件名避免中文和空格

---

## 表格样式最佳实践

### 1. 报价表（推荐使用三线表）

```json
{
  "style": {
    "table_border": "three",
    "table_shade": null,
    "table_header_color": "000000"
  }
}
```

```markdown
+----+----------+----------+
| 项目 | 金额 (万) | 占比 |
+====+==========+==========+
| A 项  | 100      | 50%     |
+----+----------+----------+
```

### 2. 数据对比表（使用全框线 + 斑马纹）

```json
{
  "style": {
    "table_border": "full",
    "table_zebra": true,
    "table_zebra_fill": "F7F7F7"
  }
}
```

### 3. 表单式表格（无表头）

```json
{
  "style": {
    "header_rows": 0
  }
}
```

```markdown
| 项目名称 | 示例智慧园区平台 |
| 申报单位 | 示例科技有限公司 |
| 项目负责人 | 张三 |
```

---

## 图片处理技巧

### 1. 图片组织

```
project/
├── md/
│   └── 01_chapter.md
├── images/
│   ├── arch.png
│   ├── system.png
│   └── team/
│       └── org-chart.png
```

### 2. 配置资源路径

```json
{
  "resource_paths": ["images", "docs/media"]
}
```

### 3. 图片尺寸建议

- **架构图**: 13cm × 9cm（宽屏）
- **流程图**: 15cm × 10cm
- **截图**: 12cm × 8cm
- **组织架构图**: 10cm × 12cm（竖屏）

---

## 验证与调试

### 完整工作流

```bash
# 1. 环境检查
python scripts/render.py --doctor

# 2. 渲染文档
python scripts/render.py --src . --out output.docx \
  --config config.json --pdf --check

# 3. 全面验证
python scripts/validate.py output.docx --out evidence/

# 3b. 带视觉基线的验证（比对 baselines/ 页图，dpi=100 同 snapshot.py）
python scripts/validate.py output.docx --baseline baselines/ --out evidence/

# 4. 查看报告
cat evidence/report.json

# 5. 基线比对（首次运行）
python scripts/snapshot.py output.pdf --update

# 6. 回归测试（后续运行）
python scripts/snapshot.py output.pdf
```

### 调试技巧

#### 问题 1: 图片未嵌入

**症状**: `images: 0/5 ok`

**解决**:
```json
{
  "resource_paths": ["./images", "../media"]
}
```

#### 问题 2: 表格错位

**症状**: 表格跨页或表头丢失

**解决**:
- 检查是否设置了 `w:tblHeader`
- 确保网格表格格式正确（`+===+` 分隔表头）
- 增加单元格边距：`"cell_margin_v": 60`

#### 问题 3: 空白页过多

**症状**: `blank_pages: 3/20`

**解决**:
- 减少 `page_break_before` 使用
- 调整 caption 的 `keep_with_next`
- 放宽阈值：`--max-empty 2`

---

## 常见问题解决

### Q1: Word 打不开生成的文档

**原因**: OOXML 结构错误

**解决**:
1. 检查 `post.py` 是否正常运行
2. 使用 `scripts/validate.py` 检查 package integrity
3. 尝试简化文档结构（减少嵌套表格）

### Q2: 目录刷新失败

**原因**: Word 未自动更新域

**解决**:
```bash
# 在 Word 中按 Ctrl+A → F9 刷新所有域
# 或使用 finalize.py（已包含自动刷新）
python scripts/finalize.py input.docx output.pdf
```

### Q3: 字体显示异常

**原因**: 缺少对应字体

**解决**:
1. 确认系统已安装所需字体（宋体、黑体、仿宋等）
2. 使用 `distill.py` 检测模板字体
3. 在 config 中指定字体：
   ```json
   {
     "style": {
       "east_font": "宋体",
       "latin_font": "Times New Roman"
     }
   }
   ```

### Q4: Patch 操作失败

**原因**: 锚点不存在或文本不匹配

**解决**:
```json
{
  "preconditions": {
    "must_contain": ["预期文本"]
  },
  "operations": [
    {
      "op": "replace_text",
      "target": {"paragraph": 10},
      "expected_old_text": "旧文本",
      "new_text": "新文本"
    }
  ]
}
```

先用 `--dry-run` 测试，再 `--apply`。

---

## 🎯 实战案例

### 案例 1: 投标文件编制

**步骤**:
1. 用 distill 或手写生成投标 config.json（可参考 `profiles/tender-v1.json` 的设计约定）
2. 编写 Markdown 章节（投标函、技术方案、报价等）
3. 渲染并验证
4. 生成证据包供审计

**命令**:
```bash
python scripts/render.py --src tender-md/ --out bid.docx \
  --config tender-config.json --pdf --check

python scripts/validate.py bid.docx --out bid-evidence/
```

### 案例 2: 公文起草

**步骤**:
1. 用 distill 或手写生成公文 config.json（可参考 `profiles/gongwen-v1.json` 的 GB/T 9704-2012 约定）
2. 按 GB/T 9704-2012 标准写作
3. 验证版式合规性

**命令**:
```bash
python scripts/render.py --src gongwen-md/ --out document.docx \
  --config gongwen-config.json --pdf --check
```

---

## 📞 获取帮助

- **GitHub Issues**: [https://github.com/Rayminliu/texere/issues](https://github.com/Rayminliu/texere/issues)
- **文档**: [README.md](README.md)
- **示例**: [examples/](examples/)

---

*最后更新：2026-09-20*
