# 更新记录

遵循语义化版本：主版本 = 不兼容变更，次版本 = 新增功能，修订 = 修复。
`ref.docx` 模板一旦改动会体现在次版本号上，因为输出版式可能随之变化
（可用 `snapshot.py` 回归）。

## 0.2.0 — 2026-09-16

### 新增
- 表/图按章自动编号与 `@tab:` / `@fig:` 交叉引用（config `"auto_number": true` 启用）
- `filters/captions.lua`：在 pandoc **AST 层**把题注标记成 `TableCaption` / `FigureCaption`，
  `post.py` 不再靠正则反推
- `snapshot.py`：PDF 版式快照回归，逐像素比对 `baselines/`，漂移即 `exit 1`
- `post.py` 支持 config 的 `style` 段：页码模板、目录深度、题注颜色/字号、表头底纹、中西文字体
- `make_ref.py` 支持 `--body-font / --latin-font / --heading-font / --body-size`
- `render.py --version`；CI（GitHub Actions）跑不依赖 Word 的 22 项断言与 docx 冒烟

### 修复
- 跨页重复表头（`w:tblHeader`）此前只写在 README 军规里，**代码未实现**
- `make_ref.py` 硬编码 `d:\desktop\国创\build\` 路径，换机器即失效且模板会悄悄漂移
- `post.py` 缺一级标题时抛 `StopIteration`，现给出可诊断提示
- Windows GBK 控制台下 `render.py` 打印子进程输出会 `UnicodeEncodeError` 崩溃
- `--check` 单独使用时被静默忽略，现自动补 `--pdf`
- `--doctor` 依据注册表 `CurVer` 误报 WPS，改为 **COM 实测**引擎身份
- `check_pdf.py` 无论检出多少空白页都 `exit 0`，现带阈值退出码

## 0.1.0 — 2026-09-16
- 从"农域多视图"参赛计划书管线抽离并独立成包
- 依赖迁移到 `pyproject.toml` + `uv.lock`，可选依赖按用途分离
- 新增 `render.py --doctor` 环境自检
- 新增覆盖「排版六条军规」的 pytest 断言
