# Texere 配置 Schema

本文档定义了 texere 项目的统一配置格式。

## 配置结构

```json
{
  "version": "1.0",                    // 配置版本
  "name": "配置文件名称",              // 人类可读的名称
  "description": "配置描述",           // 详细说明
  
  // ========== 封面和元数据 =========
  "cover": [                           // 封面行 [[样式名，文本], ...]
    ["CoverTop", "顶部文字"],
    ["CoverTitle", "标题"],
    ["CoverSub", "副标题"],
    ["CoverInfo", "信息行"],
    ["CoverDate", "日期"]
  ],
  "header": "页眉文字",                // 正文节页眉
  "title": "文档标题",                 // 文档属性
  "author": "作者",                    // 文档属性
  "subject": "主题",                   // 文档属性
  "comments": "备注",                  // 文档属性
  
  // ========== 页面设置 =========
  "page": {
    "width": 21,                       // 页面宽度 (cm)
    "height": 29.7,                    // 页面高度 (cm)
    "margin_top": 2.5,                 // 上边距 (cm)
    "margin_bottom": 2.5,              // 下边距 (cm)
    "margin_left": 3.17,               // 左边距 (cm)
    "margin_right": 3.17               // 右边距 (cm)
  },
  
  // ========== 目录设置 =========
  "toc": {
    "depth": "1-2",                    // TOC 收录层级
    "title": "目　录",                  // TOC 标题
    "title_size": 16,                  // TOC 标题字号 (pt)
    "title_color": "#000000",          // TOC 标题颜色 (HEX)
    "placeholder": "【自动生成】",      // TOC 占位符文本
    "placeholder_size": 12             // TOC 占位符字号 (pt)
  },
  
  // ========== 页眉页脚 =========
  "footer": {
    "page_number_template": "— {n} —", // 页码模板，{n}为页码占位符
    "page_number_size": 9,             // 页码字号 (pt)
    "alignment": "center"              // 对齐方式：left/center/right
  },
  "header": {
    "size": 9,                         // 页眉字号 (pt)
    "color": "#595959",                // 页眉颜色 (HEX)
    "rule_color": "#BFBFBF",           // 页眉下边框颜色 (HEX)
    "rule_size": 4                     // 页眉下边框粗细 (1/8 pt)
  },
  
  // ========== 字体样式 =========
  "styles": {
    "body": {
      "font_eastAsia": "宋体",         // 中文字体
      "font_latin": "Times New Roman", // 西文字体
      "size": 12,                      // 字号 (pt)
      "line_spacing": 1.5,             // 行距倍数
      "first_line_indent": 24,         // 首行缩进 (twips, 2 字符=24)
      "space_before": 0,               // 段前间距 (pt)
      "space_after": 0                 // 段后间距 (pt)
    },
    "h1": {                            // 一级标题
      "font_eastAsia": "黑体",
      "font_latin": "Arial",
      "size": 16,
      "bold": true,
      "page_break_before": true,
      "space_before": 12,
      "space_after": 6
    },
    "h2": {                            // 二级标题
      "font_eastAsia": "黑体",
      "font_latin": "Arial",
      "size": 14,
      "bold": true,
      "page_break_before": false,
      "space_before": 10,
      "space_after": 5
    },
    "h3": {                            // 三级标题
      "font_eastAsia": "黑体",
      "font_latin": "Arial",
      "size": 12,
      "bold": true,
      "page_break_before": false,
      "space_before": 8,
      "space_after": 4
    }
  },
  
  // ========== 题注设置 =========
  "caption": {
    "alignment": "center",             // 对齐方式
    "keep_with_next": true,            // 与下一段同页
    "color": "#404040",                // 颜色 (HEX)
    "size": 10.5,                      // 字号 (pt)
    "space_before": 6,                 // 段前间距 (pt)
    "space_after": 4,                  // 段后间距 (pt)
    "italic": false                    // 是否斜体
  },
  
  // ========== 表格样式 =========
  "table": {
    "border": "full",                  // 边框类型：full/three/none
    "border_size": 6,                  // 全框线粗细 (1/8 pt)
    "border_color": "#808080",         // 边框颜色 (HEX)
    "header_shade": "#EDEDED",         // 表头底纹颜色 (HEX)
    "header_bold": true,               // 表头加粗
    "header_center": true,             // 表头居中
    "repeat_header": true,             // 跨页重复表头
    "cell_margin_v": 40,               // 单元格上下边距 (twips)
    "cell_margin_h": 80,               // 单元格左右边距 (twips)
    "size": 10.5,                      // 表格字号 (pt)
    "zebra": false,                    // 斑马纹开关
    "zebra_fill": "#F7F7F7"            // 斑马纹填充色 (HEX)
  },
  
  // ========== 内容修复 =========
  "content_fixes": [                   // 编辑性替换表 [[旧文本，新文本], ...]
    ["旧值 1", "新值 1"],
    ["旧值 2", "新值 2"]
  ],
  
  // ========== 资源路径 =========
  "resource_paths": [                  // 图片搜索路径列表
    "./images",
    "../media"
  ],
  
  // ========== 题注关键字 =========
  "caption_words": {                   // 自定义题注关键字
    "table": ["表", "表格"],
    "figure": ["图", "图片"]
  },
  
  // ========== 参考文档 =========
  "reference_doc": "template.docx"     // 引用外部 docx 作为样式模板
}
```

## 兼容性说明

### 向后兼容
- 现有 `sample_config.json` 格式完全兼容，仅包含 `cover`, `header`, `title`, `author`, `subject`, `comments`
- `post.py` 会自动忽略未定义的字段，使用默认值

### Profile 迁移
旧的 `profiles/*.json` 文件可以转换为标准格式：

```bash
python scripts/distill.py template.docx --out config.json
```

生成的 `config.json` 将包含完整的配置结构。

## 示例配置

### 最小化配置（仅封面）
```json
{
  "cover": [
    ["CoverTitle", "投标文件"]
  ]
}
```

### 完整配置
参见 `profiles/formal-cn-v1.json`

## 注意事项

1. **颜色格式**: 所有颜色使用 6 位 HEX 代码（如 `#404040`）
2. **单位说明**:
   - `pt`: 磅 (1pt = 1/72 inch)
   - `cm`: 厘米
   - `twips`: 1/20 point (Word 内部单位)
3. **必填字段**: 无（所有字段都是可选的）
4. **验证**: 建议使用 JSON Schema 验证器检查配置合法性
