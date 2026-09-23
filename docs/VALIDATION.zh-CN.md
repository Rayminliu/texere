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
└── signature            # 校验清单：docx + report 的哈希、检查摘要
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
| ✅ Word 验收 | 真 Word 能打开并成功导出 PDF |
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
✅ [PASS] renderer_acceptance: Word 验收：OK
⏭️ [SKIP] visual_drift: 跳过 (未提供基线目录)

Summary
────────────────────────────
Passed: 7/9 (skipped: 2)

✅ All checks passed

Evidence package saved to: evidence/
  - report.json (structured validation report)
  - page-XXX.png (sample screenshots)
  - signature (校验清单：docx + report 的哈希)
```

任何一项失败即 exit code 1，并给出详细错误。参数（`--profile` / `--max-empty` / `--quiet` 等）见
`docs/SCRIPT_HELP.md` §validate.py。

### 人工门槛

9 项检查管的是结构，头一遍过某份新文档时，这四件事仍然得靠人：

1. 日志出现 `images: n/m ok` 且 **n == m**（m 是源 md 中 `![` 的次数）
2. `near-empty pages: 0`
3. `OK`（Word 成功打开并导出 PDF）
4. **人工过一遍 PDF 渲染图**——机器能查页数、图片数、空白页，
   查不了"这张图画得对不对、表头有没有被截断"

