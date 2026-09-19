# texere for ChatGPT Work & Codex

## Installation

### Method 1: ChatGPT Desktop App (Recommended)

1. Open ChatGPT Desktop App
2. Go to Codex or switch to ChatGPT Work
3. Open **Plugins** directory
4. Search for "texere"
5. Click **+** to install

### Method 2: ChatGPT Web (Work Mode)

1. Switch to **Work** mode
2. Open **Plugins** directory from sidebar
3. Navigate to **Skills** tab
4. Find and install "texere"

### Method 3: Codex CLI

```bash
# Open plugin browser
/plugins

# Install texere
# Follow the interactive prompts
```

## Usage

Once installed, you can use texere in ChatGPT Work or Codex:

```python
# In ChatGPT Work interface
@texere render chapters/ --out bid.docx --pdf

# Or use the skill directly
/texere validate bid.docx
```

## Features

- **Compiler**: Markdown → DOCX + PDF with Chinese formal document formatting
- **Contract**: Design contracts via `profile.json` or `ref.docx`
- **Evidence**: 9-item validation with structured reports and screenshots

## Requirements

- Windows with Microsoft Word (for PDF export)
- pandoc >= 3.1
- Python 3.10+
- python-docx, lxml

On Linux/macOS, only DOCX generation is available (no PDF export).

## License

MIT

## Compatibility

This plugin works with:
- ✅ ChatGPT Work (Desktop & Web)
- ✅ ChatGPT Codex (Desktop & CLI)
- ✅ Claude Code (via separate plugin)

One Skill, multiple platforms!
