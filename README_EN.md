# WPS Office AI Assistant

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
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
Install WPS Skills for me, here's the guide: https://github.com/LargeCupPanda/WPS_Skills/blob/main/INSTALL.md
```

Claude Code will read the installation guide and complete all steps automatically!

> ⚠️ Prerequisite: Please install [WPS Office](https://www.wps.com/) first

---

## Overview

WPS Office AI Assistant is a natural language office automation tool powered by Claude AI. Through MCP (Model Context Protocol), you can control WPS Office using natural language - no more memorizing formulas or navigating complex menus.

### Key Features

- **Natural Language Control** - Describe what you need, AI executes
- **Full Office Suite** - Excel, Word, PPT all supported
- **Smart Formula Generation** - Describe calculations, get formulas
- **One-Click Beautification** - Professional PPT styling
- **Stable COM Bridge** - Reliable PowerShell COM interface

### Usage Examples

```bash
# Excel Operations
User: Read data from A1 to C5
User: Change B3 cell value to 4.8
User: Create a bar chart for sales data
User: Sort by column B descending
User: Calculate average, max, min for column B

# Word Operations
User: Insert text at the end of document
User: Replace all "OldCompany" with "NewCompany"
User: Insert a 3x4 table
User: Change all text to Arial 12pt

# PPT Operations
User: Add a new slide with title "Project Summary"
User: Unify all fonts to Arial
User: Apply business style to current slide
User: Add a text box on slide 1
```

---

## Requirements

| Item | Windows | macOS |
|------|---------|-------|
| OS | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 or later | Mac version latest |
| Node.js | 18.0.0 or later | 18.0.0 or later |
| Claude Code | Latest version | Latest version |

---

<details>
<summary><b>📦 Manual Installation (Click to expand)</b></summary>

### Option 1: One-Click Script

```powershell
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

### Option 2: Manual Steps

1. **Install dependencies and build**
   ```bash
   cd wps-office-mcp
   npm install
   npm run build
   ```

2. **Configure Claude Code** - Edit `~/.claude/settings.json`:
   ```json
   {
     "mcpServers": {
       "wps-office": {
         "command": "node",
         "args": ["C:\\path\\to\\WPS_Skills\\wps-office-mcp\\dist\\index.js"]
       }
     }
   }
   ```

3. **Install WPS Add-in**
   - Copy `wps-claude-addon` to `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\`
   - Add to `publish.xml`:
     ```xml
     <jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_/"/>
     ```

4. **Restart Claude Code and WPS Office**

</details>

---

## Features

### Excel Features (86 implemented)

| Category | Count | Features | Windows | macOS |
|----------|-------|----------|---------|-------|
| Workbook Operations | 5 | Open/Create/Switch/Close workbooks | ✅ | ✅ |
| Worksheet Operations | 7 | Create/Delete/Rename/Copy/Move sheets | ✅ | ✅ |
| Cell Read/Write | 7 | Read/Write cells/ranges/formulas/info | ✅ | ✅ |
| Formatting | 15 | Style/Border/Number format/Merge/AutoFit | ✅ | ✅ |
| Row/Column Operations | 8 | Insert/Delete/Hide/Show rows/columns | ✅ | ✅ |
| Conditional Formatting | 3 | Add/Remove/Get conditional formats | ✅ | ✅ |
| Data Validation | 3 | Add/Remove/Get data validation | ✅ | ✅ |
| Find & Replace | 2 | Find/Replace in worksheet | ✅ | ✅ |
| Data Processing | 10 | Sort/Filter/Dedupe/Clean/Copy/Transpose | ✅ | ✅ |
| Named Ranges | 3 | Create/Delete/Get named ranges | ✅ | ✅ |
| Comments | 3 | Add/Delete/Get comments | ✅ | ✅ |
| Protection | 3 | Protect sheet/workbook | ✅ | ✅ |
| Formula Functions | 5 | Set formula/Array formula/Diagnose/Calculate | ✅ | ✅ |
| Charts | 2 | Create/Update charts | ✅ | ✅ |
| Pivot Tables | 2 | Create/Update pivot tables | ✅ | ✅ |
| Financial Features | 5 | Cross-workbook refs/Hyperlinks/Images/Wrap | ✅ | ✅ |
| Extended Features | 5 | Print area/Grouping/Lock cells | ✅ | ✅ |

> 📌 Excel features cover 95%+ daily scenarios, including financial/business use cases

### Word Features (22 implemented)

| Feature | Description | Windows | macOS |
|---------|-------------|---------|-------|
| Get Document Info | Name, paragraphs, word count | ✅ | ✅ |
| Get Document Stats | Pages, characters, lines | ✅ | ✅ |
| Read Text | Get document content | ✅ | ✅ |
| Insert Text | At start/end/cursor | ✅ | ✅ |
| Set Font | Font, size, bold, etc. | ✅ | ✅ |
| Find & Replace | Batch replace text | ✅ | ✅ |
| Insert Table | Create and fill tables | ✅ | ✅ |
| Apply Style | Apply Word styles | ✅ | ✅ |
| Paragraph Format | Alignment, indent, spacing | ✅ | ✅ |
| Page Setup | Margins, paper, orientation | ✅ | ✅ |
| Generate TOC | Auto generate TOC | ✅ | ✅ |
| Insert Image | Insert and resize images | ✅ | ✅ |
| Insert Header | Add header text | ✅ | ✅ |
| Insert Footer | Add footer and page number | ✅ | ✅ |
| Insert Page Break | Page/section breaks | ✅ | ✅ |
| Insert Hyperlink | Add hyperlinks | ✅ | ✅ |
| Insert Bookmark | Add bookmarks | ✅ | ✅ |
| Get Bookmarks | Get bookmark list | ✅ | ✅ |
| Add Comment | Add document comments | ✅ | ✅ |
| Get Comments | Get comment list | ✅ | ✅ |
| Save Document | Save current document | ✅ | ✅ |
| Export PDF | Convert to PDF format | ✅ | ✅ |

### PPT Features

| Feature | Description | Windows | macOS |
|---------|-------------|---------|-------|
| Get Presentation Info | Name, slide count, shapes | ✅ | 🔧 |
| Add Slide | Multiple layouts | ✅ | 🔧 |
| Set Title | Modify slide title | ✅ | 🔧 |
| Add Text Box | Custom position and style | ✅ | 🔧 |
| Unify Font | Consistent fonts | ✅ | 🔧 |
| Beautify Slide | Business/Tech/Creative/Minimal | ✅ | 🔧 |
| Add Shape | Insert shapes | 🚧 | 🚧 |
| Add Animation | Enter/exit animations | 🚧 | 🚧 |
| Set Theme | Apply PPT themes | 🚧 | 🚧 |

> Legend: ✅ Tested | 🔧 Pending test | 🚧 In development | ⚠️ Known issue

### Common Features

| Feature | Description | Windows | macOS |
|---------|-------------|---------|-------|
| Save File | Save current document | ✅ | ✅ |
| Export PDF | Convert to PDF | ✅ | ✅ |
| Format Conversion | Word/Excel/PPT conversion | 🚧 | 🚧 |

---

## Architecture

```
Windows:
Claude Code → MCP Server (Node.js) → PowerShell COM → WPS Office

macOS:
Claude Code → MCP Server (Node.js) → HTTP → WPS Add-in (JS API) → WPS Office
```

- **MCP Server**: 29 tools handling AI requests
- **Windows COM Bridge**: PowerShell calls WPS COM interfaces (Ket/Kwps/Kwpp)
- **macOS HTTP Bridge**: HTTP calls to WPS Add-in built-in service (port 58891)
- **WPS Add-in**: Shows connection status, provides HTTP API on Mac

---

## Project Structure

```
WPS_Skills/
├── wps-office-mcp/          # MCP Server (Core)
│   ├── src/                 # TypeScript source
│   ├── dist/                # Compiled output
│   ├── scripts/             # PowerShell COM bridge (Windows)
│   │   └── wps-com.ps1      # COM operations script
│   └── package.json
├── wps-claude-addon/        # WPS Add-in (Windows)
│   ├── ribbon.xml           # Ribbon config
│   └── js/main.js           # Add-in logic
├── wps-claude-assistant/    # WPS Add-in (macOS)
│   ├── main.js              # HTTP Server + All Handlers
│   ├── manifest.xml         # Add-in manifest
│   └── ribbon.xml           # Ribbon config
├── scripts/
│   ├── auto-install.ps1     # Windows one-click install
│   └── auto-install-mac.sh  # macOS one-click install
├── skills/                  # Claude Skills definitions
├── docs/                    # Design docs (private)
└── README.md
```

---

## FAQ

### Q: Claude Assistant tab not showing?

**A:** Check:
1. Add-in folder name ends with `_`
2. `publish.xml` is configured correctly
3. Restart WPS Office

### Q: MCP Server connection failed?

**A:** Troubleshoot:
1. Verify `settings.json` path is correct
2. Confirm `npm run build` was executed
3. Restart Claude Code

### Q: WPS operation shows connection error?

**A:** Ensure:
1. WPS Office is running with a document open
2. The correct app is open (Excel for Excel operations, etc.)

---

## TODO

### Near-term (v1.1)

- [x] **macOS Support** - Cross-platform compatibility ✅ Completed
- [x] **Excel Formula Diagnosis** - Analyze errors, suggest fixes ✅ Completed
- [x] **Excel Pivot Tables** - Create and manipulate pivot tables ✅ Completed
- [x] **Excel Conditional Formatting** - Set format rules ✅ Completed
- [x] **Word TOC Generation** - Auto generate table of contents ✅ Completed
- [x] **Word Insert Image** - Insert and position images ✅ Completed
- [ ] **PPT Animations** - Enter, exit, emphasis animations
- [ ] **PPT Themes** - Apply built-in themes

### Mid-term (v1.2)

- [ ] **Cross-app Conversion** - Word/Excel/PPT conversion
- [ ] **Word to PPT** - Generate PPT from Word outline
- [ ] **Batch Conversion** - Bulk file format conversion
- [ ] **Batch Watermark** - Add text/image watermarks
- [ ] **Mail Merge** - Word mail merge functionality
- [x] **Array Formulas** - Advanced formula support ✅ Completed

### Long-term (v2.0)

- [ ] **PDF Support** - WPS PDF operations
- [ ] **Mind Map Support** - WPS Mind Map
- [ ] **AI Content Generation** - Smart document generation
- [ ] **Multi-document Operations** - Cross-document actions
- [ ] **Automation Workflows** - Define and execute workflows
- [ ] **Cloud Collaboration** - WPS Cloud document support
- [ ] **Enterprise Deployment** - Enterprise deployment solutions

---

## License

MIT License

## Developer

**LargeCupPanda** - [GitHub](https://github.com/LargeCupPanda)

---

<p align="center">Made with love for WPS Office users</p>
