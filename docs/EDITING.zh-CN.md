# EDITING.zh-CN.md

> 从 `README.zh-CN.md` 下沉而来的明细。中英两份由 `tests/test_docs_sync.py` 守着同构。

## 编辑已有 docx

`scripts/edit.py` 是**独立于渲染链路**的第二条链，契约正好相反：

| | 渲染链路 | 编辑链路 |
|---|---|---|
| 输入 | Markdown | 已有 docx |
| 契约 | 只改版式，不改内容 | **只改你指定的地方，其余字节原样保留** |
| 禁止 | 改写文字 | 注入封面 / 目录 / 页码 / 重排样式 |

操作长这样（完整参数：`docs/SCRIPT_HELP.md` §edit.py）：

```bash
python scripts/edit.py 标书.docx --list                      # 看结构：分节 / 表格 / 段落
python scripts/edit.py 标书.docx --replace "旧=新" --scope body,tables,header,footer
python scripts/edit.py 标书.docx --after "锚点文字" --text "新段落"   # 另有 --before / --delete
python scripts/edit.py 标书.docx --cell 0 2 1 "1,060,000"            # 表号 / 行 / 列 / 值
python scripts/edit.py 标书.docx --add-rows 0 3 --template-row 2     # 另有 --add-row / --del-row
python scripts/edit.py 标书.docx --fill data.json                    # 批量填表，JSON 或 CSV
python scripts/edit.py 标书.docx --header "新版页眉" --footer "— X —" --section all
python scripts/edit.py 标书.docx --replace "A=B" --verify            # 改完让 Word 打开一次验收
```

默认先把原文件备份成 `<name>.bak.docx`（`--no-backup` 关闭，`--out` 另存）。

**为什么可以放心原地改**：实测 `python-docx` 打开再保存，部件**零丢失、零新增**（4 页文档实测；
文件变小只是重新压缩）。

**两个必须知道的坑**

1. **Word 会把文字切成多个 run。** `自开标之日起 90 日历天` 可能是三个 run，数字单独一个，
   直接遍历 `run.text` 根本找不到。`edit.py` 的做法是：拼整段文本定位，把替换内容写进
   「命中起点所在的那个 run」，其余 run 只删掉被覆盖的字符——run 数和各 run 的格式都保住。
2. **锚点命中多处时拒绝执行。** 目录被 Word 刷成静态文本后，`1.2 资质与业绩` 这类标题在目录和
   正文里各有一份，照着插两遍就会插进目录。此时会列出候选并退出，换更精确的锚点，
   或显式加 `--all-anchors`。

**明确不做**：批注、修订、水印、内容控件，以及带宏的 `.docm`（`python-docx` 保存会丢
`vbaProject.bin`）。

### 要「重排」而不是「编辑」时

想把别人的 docx 换成我们这套格式，走 Markdown 中转，但**中间产物必须手工清理**：

```bash
pandoc 甲方文档.docx -t markdown --wrap=none --extract-media=media -o 01_内容.md
# 手工清理：删掉原来的封面行与目录块、去掉表头残留的 ** 加粗
python scripts/render.py --src . --out 新版.docx --config cfg.json --pdf --check
```

4 页文档实测：不清理会得到**双封面 + 双目录**，`[1](#...)` 链接语法漏进正文，页数从 4 变 5。
`--extract-media` 抽出的图片路径是相对**执行目录**的，所以在仓库根执行，或把 `media/` 放进 `--src`。

