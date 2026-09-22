# CONFIG.zh-CN.md

> 从 `README.zh-CN.md` 下沉而来的配置参考。中英两份由 `tests/test_docs_sync.py` 守着同构。

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
| 表格 | `header_rows` | 自动——**逐表**识别（读 pandoc 打的 `w:tblHeader`）；填数字则强制；`0` = 该表没有表头（表单 / 附件类首行是字段名），并一并去掉「跨页重复表头」。何时该动它见[视觉控制](TABLES.zh-CN.md#视觉控制) |
| 表格 | `table_border` / `table_shade` / `table_size` | `full` / `EDEDED` / `10.5` |
| 表格 | `table_header_color` / `table_zebra` / `table_zebra_fill` | 不指定 / `false` / `F7F7F7` |
| 表格 | `cell_margin_v` / `cell_margin_h` / `table_para_space` | `40` / `80` / `1` |
| 表格 | `border_size` / `border_color` / `three_line_size` | `6` / `808080` / `12` |
| 字体 | `east_font` / `latin_font` | `宋体` / `Times New Roman`（**只作用于表格与题注**） |

> 颜色写 6 位十六进制（`404040`），粗细单位 1/8 pt，间距单位 pt，单元格边距单位 twips。

