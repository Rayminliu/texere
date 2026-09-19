# texere Plugin Packages

texere 提供多个插件包，适配不同的 AI 编程助手平台。

## 📦 Available Plugins

### 1. Claude Code Plugin
**Location**: `./texere-plugin/`

适配 Claude Code 的插件系统。

**安装方式**:
```bash
/plugin marketplace add Rayminliu/texere
/plugin install texere@latest
```

**目录结构**:
```
texere-plugin/
├── .claude-plugin/
│   └── plugin.json      # Claude Code 插件配置
├── skills/
│   └── texere/
│       └── SKILL.md     # Skill 定义（通用）
└── README.md
```

---

### 2. ChatGPT Work & Codex Plugin
**Location**: `./texere-codex-plugin/`

适配 ChatGPT Work 和 OpenAI Codex 的插件系统。

**安装方式**:
- **Desktop App**: 打开 ChatGPT Desktop → Codex/Work → Plugins 目录 → 搜索"texere"
- **Web**: Work 模式 → Plugins 目录 → Skills 标签 → 安装 texere
- **CLI**: `/plugins` → 选择并安装 texere

**目录结构**:
```
texere-codex-plugin/
├── .codex-plugin/
│   └── plugin.json      # OpenAI Codex 插件配置
├── skills/
│   └── texere/
│       └── SKILL.md     # Skill 定义（通用）
└── README.md
```

---

## 🎯 Common Structure

两个插件包共享相同的 **Skill** 部分：

```
skills/texere/
└── SKILL.md             ← 通用 Skill 文件（两个插件包相同）
```

**SKILL.md** 在所有平台上都通用，这是 Agent Plugins 1.0 标准的核心优势！

---

## 🔧 Platform Differences

| 组件 | Claude Code | ChatGPT Work/Codex |
|------|-------------|-------------------|
| **Skill (SKILL.md)** | ✅ 通用 | ✅ 通用 |
| **Plugin Config** | `.claude-plugin/plugin.json` | `.codex-plugin/plugin.json` |
| **Installation** | `/plugin install` | Desktop/Web UI or `/plugins` |
| **Marketplace** | Agents Skills Registry | OpenAI Plugin Directory |

---

## 📋 Usage Examples

### In Claude Code:
```bash
/texere render chapters/ --out bid.docx --pdf
/texere validate bid.docx
```

### In ChatGPT Work:
```python
@texere render chapters/ --out bid.docx --pdf
```

### In Codex CLI:
```bash
/plugins
# Select and install texere
```

---

## 🌐 Cross-Platform Compatibility

All texere plugins share the same core Skill definition, ensuring consistent behavior across platforms:

- ✅ **Same logic**: All platforms use identical document compilation logic
- ✅ **Same validation**: 9-item validation works everywhere
- ✅ **Same profiles**: `profile.json` files work on all platforms
- ✅ **Same examples**: Tutorial examples are platform-agnostic

---

## 🚀 Contributing

To add support for a new platform:

1. Keep `skills/texere/SKILL.md` unchanged (it's universal)
2. Create platform-specific `plugin.json` in appropriate directory
3. Follow that platform's plugin specification
4. Update this README with installation instructions

---

## 📄 License

MIT — see [LICENSE](../LICENSE).
