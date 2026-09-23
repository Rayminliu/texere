# Renderer Compatibility Corpus（渲染器兼容性语料）

用一组**真实中文正式文档**，在三种 renderer 上各跑一遍，产出第一张
**Chinese Office Renderer Compatibility Matrix**——目标不是"证明支持三个 renderer"，
而是**测量并回归不同 renderer 的结果到底差多少**。

> 架构层（0.6.4）已脱离 Word：`Core` 不依赖任何 Office，`RendererAdapter` 把 docx→PDF
> 抽成 Word / LibreOffice / WPS 三个可插拔实现，`page_numbering` / `blank_pages` /
> `visual_drift` 只消费 PDF、与 renderer 无关（契约测试用 `FakeRenderer` 守住）。
> 但**产品层尚未证明 WPS / LO 的中文 fidelity**——本 corpus 就是用来补这块证据的。

## 三个 renderer 的定位（不是抢"谁是真 renderer"）

| Renderer | 角色 | 适用 |
|---|---|---|
| `word` | Microsoft Office 保真 renderer | Windows + Microsoft Word |
| `wps` | 中文 Office renderer | Windows + WPS Office |
| `libreoffice` | 便携 renderer | 任意系统（soffice headless，CI 友好） |

## 案例清单（`cases/`，合并渲染为一份 docx 做第一遍兼容性扫描）

| # | 案例 | 重点考察 |
|---|---|---|
| 01 | 普通公文 | 基础排版、样式继承、提示框 |
| 02 | 长表格跨页 | 跨页表格、续页表头重复、框线 |
| 03 | 多级表头 | 多级/合并表头、列宽 |
| 04 | 图片 + 图注 | 图片嵌入、题注居中灰字 |
| 05 | 目录 (TOC) | 目录域、分节页码 |
| 06 | 页码 | 分节页码连续性 |
| 07 | 页眉页脚 | 页眉条件继承、页脚 |
| 08 | 客户模板 | `reference_doc` 版式接管 |
| 09 | 封面 | 封面页、CoverInfo 段 |
| 10 | 签章页 | 签章块、末页稀疏容忍 |

> 第一遍把 10 个案例合并成一份 docx，先看"整体渲染差异"；后续可拆成单案例各自建基线，
> 做更细的逐案例矩阵。

## 怎么跑（在装了对应 renderer 的机器上，从仓库根目录）

```bash
# Word（默认 renderer）
python scripts/render.py --src corpus/cases --out corpus/out/word/样例.docx \
    --config corpus/config.json --pdf --check
python scripts/validate.py corpus/out/word/样例.docx --out corpus/evidence/word/

# LibreOffice
python scripts/render.py --src corpus/cases --out corpus/out/libreoffice/样例.docx \
    --config corpus/config.json --renderer libreoffice --pdf --check
python scripts/validate.py corpus/out/libreoffice/样例.docx --out corpus/evidence/libreoffice/ --renderer libreoffice

# WPS
python scripts/render.py --src corpus/cases --out corpus/out/wps/样例.docx \
    --config corpus/config.json --renderer wps --pdf --check
python scripts/validate.py corpus/out/wps/样例.docx --out corpus/evidence/wps/ --renderer wps
```

## baseline 必须 renderer-scoped（关键）

baseline 是 **rendered truth**，不是 abstract document truth。同一份 docx 经 Word / WPS / LO
出的 PDF 在字体、分页、TOC 上可能不同，**绝不共用一个基线**：

```
corpus/baselines/
  word/        page-001.png ...
  wps/         page-001.png ...
  libreoffice/ page-001.png ...
```

录制：`python scripts/snapshot.py corpus/out/<renderer>/样例.pdf --update --out corpus/baselines/<renderer>/`
回归：`python scripts/snapshot.py corpus/out/<renderer>/样例.pdf --baseline corpus/baselines/<renderer>/`

> 备注：`validate.py` 的 `--renderer` 当前只决定"用哪个 renderer 出 PDF / 验收"；
> baseline 目录仍由调用方按 renderer 分目录管理。下一步会把 baseline 自动按 renderer 分桶（代码层）。

## 兼容性矩阵模板（真机填到这里）

| Case | Word | WPS | LibreOffice |
|---|---|---|---|
| 01 普通公文 | PASS | ? | ? |
| 02 长表格跨页 | PASS | ? | ? |
| 03 多级表头 | PASS | ? | ? |
| 04 图片+图注 | PASS | ? | ? |
| 05 目录 | PASS | ? | ? |
| 06 页码 | PASS | ? | ? |
| 07 页眉页脚 | PASS | ? | ? |
| 08 客户模板 | PASS | ? | ? |
| 09 封面 | PASS | ? | ? |
| 10 签章页 | PASS | ? | ? |

注：`Word` 列标 PASS 是"已知在 Word 上 OK"的参照；WPS / LO 列需在真机填。
WPS 与 Word 的"刷新逻辑"不完全等价（WPS 未显式 `TablesOfContents.Update()` /
`Repaginate()`），**TOC + 分节页码是最值得测的两项**。
