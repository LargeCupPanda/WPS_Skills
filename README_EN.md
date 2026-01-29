# WPS Office AI Assistant

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Skills-Framework-purple?style=flat-square" alt="Skills Framework">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square" alt="Windows | macOS">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

<p align="center">
  English | <a href="./README.md">中文</a>
</p>

---

## 🚀 One-Click Install

**Just tell Claude Code:**

```
Install WPS Skills for me, here's the guide: https://github.com/lc2panda/wps-mcp/blob/main/INSTALL.md
```

Claude Code will read the installation guide and complete all steps automatically!

> ⚠️ Prerequisite: Please install [WPS Office](https://www.wps.com/) first

---

## Overview

WPS Office AI Assistant is a natural language office automation tool powered by Claude AI. Built on **Anthropic's official MCP + Skills dual-layer architecture**, you can control WPS Office using natural language - no more memorizing formulas or navigating complex menus.

### Key Features

- 🗣️ **Natural Language Control** - Describe what you need, AI executes
- 📊 **Full Office Suite** - Excel, Word, PPT all supported
- 🔢 **Smart Formula Generation** - Describe calculations, get formulas
- 🎨 **One-Click Beautification** - Professional PPT styling
- 🧠 **Skills Guidance** - 4 professional Skills teach AI how to complete tasks
- 🔧 **196 MCP Tools** - Complete underlying tool capabilities

### Usage Examples

```bash
# Excel Operations
User: Write a formula to lookup product prices
User: Create a pivot table for sales data
User: Highlight cells in column B greater than 100

# Word Operations
User: Generate a table of contents
User: Change all text to Arial 12pt
User: Insert a 3x4 table

# PPT Operations
User: Apply business style to this slide
User: Draw a project flowchart
User: Create a set of KPI data cards
```

---

## Requirements

| Item | Windows | macOS |
|------|---------|-------|
| OS | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 or later | Mac version latest |
| Node.js | 18.0.0 or later | 18.0.0 or later |
| Claude Code | Latest version | Latest version |
| **Feature Support** | ⚠️ Basic features (~25 methods) | ✅ Full features (196 methods) |

> ⚠️ **Windows Compatibility Note**: Windows version currently uses PowerShell COM bridge, supporting basic Excel/Word/PPT operations. Advanced features (pivot tables, conditional formatting, flowcharts, 3D effects, etc.) are being adapted. macOS version has full functionality.

---

<details>
<summary><b>📦 Manual Installation (Click to expand)</b></summary>

### Option 1: One-Click Script

```bash
git clone https://github.com/lc2panda/wps-mcp.git
cd wps-mcp

# Windows
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1

# macOS
./scripts/auto-install-mac.sh
```

### Option 2: Manual Steps

1. **Install dependencies and build**
   ```bash
   cd wps-office-mcp
   npm install
   npm run build
   ```

2. **Configure MCP Server**
   ```bash
   claude mcp add wps-office node /path/to/wps-mcp/wps-office-mcp/dist/index.js
   ```

3. **Register Skills (create symlinks to global directory)**
   ```bash
   mkdir -p ~/.claude/skills
   ln -sf /path/to/wps-mcp/skills/wps-excel ~/.claude/skills/wps-excel
   ln -sf /path/to/wps-mcp/skills/wps-word ~/.claude/skills/wps-word
   ln -sf /path/to/wps-mcp/skills/wps-ppt ~/.claude/skills/wps-ppt
   ln -sf /path/to/wps-mcp/skills/wps-office ~/.claude/skills/wps-office
   ```

4. **Install WPS Add-in** - See INSTALL.md

5. **Restart Claude Code and WPS Office**

</details>

---

## Architecture

Built on **Anthropic's official MCP + Skills dual-layer architecture**:

```
┌─────────────────────────────────────────────────────────────┐
│                    User Natural Language Request            │
│                "Write a VLOOKUP formula to find prices"     │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    Skills Layer (Instruction Packages)      │
│  skills/wps-excel/SKILL.md  - Teaches Claude Excel tasks    │
│  skills/wps-word/SKILL.md   - Teaches Claude Word tasks     │
│  skills/wps-ppt/SKILL.md    - Teaches Claude PPT tasks      │
│  skills/wps-office/SKILL.md - Teaches cross-app coordination│
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    MCP Layer (Tool Capabilities)            │
│  wps-office-mcp/            - 196 MCP Tools                 │
│  wps_get_active_workbook    - Get current workbook          │
│  wps_execute_method         - Execute operations            │
│  ...                                                        │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    WPS Add-in Layer (Executor)              │
│  Windows: PowerShell COM → WPS Office                       │
│  macOS: HTTP Polling → WPS Add-in (JS API) → WPS Office     │
└─────────────────────────────────────────────────────────────┘
```

### MCP vs Skills

| Layer | Purpose | Content |
|-------|---------|---------|
| **Skills** | Teaches Claude "how to do it" | 4 SKILL.md files with workflows and best practices |
| **MCP** | Tells Claude "what can be done" | 196 tools providing underlying capabilities |

---

## Project Structure

```
wps-mcp/
├── wps-office-mcp/          # MCP Server (Core)
│   ├── src/                 # TypeScript source
│   ├── dist/                # Compiled output
│   └── package.json
├── wps-claude-assistant/    # WPS Add-in (macOS)
│   ├── main.js              # HTTP Polling + All Handlers
│   ├── manifest.xml         # Add-in manifest
│   └── ribbon.xml           # Ribbon config
├── wps-claude-addon/        # WPS Add-in (Windows)
│   ├── ribbon.xml           # Ribbon config
│   └── js/main.js           # Add-in logic
├── skills/                  # Claude Skills Definitions
│   ├── wps-excel/SKILL.md   # Excel skill (60+ methods)
│   ├── wps-word/SKILL.md    # Word skill (25+ methods)
│   ├── wps-ppt/SKILL.md     # PPT skill (85+ methods)
│   └── wps-office/SKILL.md  # Cross-app skill
├── scripts/
│   ├── auto-install.ps1     # Windows one-click install
│   └── auto-install-mac.sh  # macOS one-click install
├── INSTALL.md               # Claude Code installation guide
└── README.md
```

---

## Features

### Excel Features (86)

| Category | Count | Features |
|----------|-------|----------|
| Workbook/Sheet Operations | 12 | Open/Create/Switch/Rename/Copy/Move |
| Cell Read/Write | 7 | Read/Write cells/ranges/formulas/info |
| Formatting | 15 | Style/Border/Number format/Merge/AutoFit |
| Row/Column Operations | 8 | Insert/Delete/Hide/Show rows/columns |
| Conditional Formatting | 3 | Add/Remove/Get conditional formats |
| Data Validation | 3 | Add/Remove/Get data validation |
| Data Processing | 10 | Sort/Filter/Dedupe/Clean/Copy/Transpose |
| Charts/Pivot Tables | 4 | Create/Update charts and pivot tables |
| Formula Functions | 5 | Set formula/Array formula/Diagnose/Calculate |
| Other | 19 | Comments/Protection/Named ranges/Find replace |

### Word Features (25)

| Category | Features |
|----------|----------|
| Document Management | Get info/Open/Switch/Get full text |
| Text Operations | Insert text/Find replace |
| Formatting | Font/Style/Paragraph |
| Document Structure | TOC/Page break/Header/Footer |
| Insert Content | Table/Image/Hyperlink/Bookmark |
| Other | Comments/Document stats |

### PPT Features (85)

| Category | Count | Features |
|----------|-------|----------|
| Presentation Management | 5 | Create/Open/Close/Switch |
| Slide Operations | 10 | Add/Delete/Duplicate/Move/Notes |
| TextBox/Shapes | 21 | Add/Delete/Style/Shadow/Gradient/Border |
| Smart Layout | 10 | Align/Distribute/Group/Connectors/Arrows |
| Image/Table/Chart | 12 | Insert/Set style |
| Data Visualization | 6 | KPI cards/Progress bars/Gauges/Donut charts |
| Flowcharts/Org Charts | 3 | Flowcharts/Org charts/Timelines |
| Animations/Transitions | 9 | Animations/Emphasis/Transitions |
| Master/3D Effects | 7 | Master operations/3D rotation/depth/material |
| Other | 2 | Slideshow |

---

## FAQ

### Q: Claude Assistant tab not showing?

**A:** Check:
1. Add-in folder name ends with `_`
2. `publish.xml` is configured correctly
3. Restart WPS Office

### Q: Skills not loaded?

**A:** Check if symlinks exist:
```bash
ls ~/.claude/skills/
```

If empty, create manually:
```bash
mkdir -p ~/.claude/skills
ln -sf /path/to/wps-mcp/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf /path/to/wps-mcp/skills/wps-word ~/.claude/skills/wps-word
ln -sf /path/to/wps-mcp/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf /path/to/wps-mcp/skills/wps-office ~/.claude/skills/wps-office
```

Then restart Claude Code.

### Q: MCP Server connection failed?

**A:** Troubleshoot:
1. Confirm `npm run build` was executed
2. Run `claude mcp list` to check configuration
3. Restart Claude Code

---

## TODO

### Near-term (v1.1) ✅ Completed

- [x] macOS Support
- [x] Excel formula diagnosis, pivot tables, conditional formatting
- [x] Word TOC generation, insert images
- [x] PPT animations, advanced beautification, 6 major capabilities
- [x] Cross-app data transfer
- [x] **Skills Framework** - Anthropic official standard

### Mid-term (v1.2)

- [ ] Cross-app format conversion
- [ ] Word to PPT
- [ ] Batch format conversion
- [ ] Mail merge

### Long-term (v2.0)

- [ ] PDF Support
- [ ] AI content generation
- [ ] Automation workflows
- [ ] Enterprise deployment

---

## License

MIT License

## Developer

**lc2panda** - [GitHub](https://github.com/lc2panda)

---

<p align="center">Made with ❤️ for WPS Office users</p>
