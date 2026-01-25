# WPS Office AI Assistant

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

<p align="center">
  English | <a href="./README.md">中文</a>
</p>

---

## 🚀 One-Click Install

**Just tell Claude Code:**

```
Install WPS Skills for me: https://github.com/LargeCupPanda/WPS_Skills
```

Claude Code will automatically complete all installation steps!

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

| Item | Requirement |
|------|-------------|
| OS | Windows 10/11 |
| WPS Office | 2019 or later |
| Node.js | 18.0.0 or later |
| Claude Code | Latest version |

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

### Excel Features

| Feature | Description | Status |
|---------|-------------|--------|
| Get Workbook Info | Name, path, sheet list | ✅ |
| Get Context | Headers, selection, used range | ✅ |
| Read Cells | Single or range | ✅ |
| Write Cells | Single or range | ✅ |
| Set Formula | Write Excel formulas | ✅ |
| Sort | Sort by column | ✅ |
| Filter | Auto filter | ✅ |
| Remove Duplicates | Delete duplicate rows | ✅ |
| Create Chart | Bar, line, pie, etc. | ✅ |
| Formula Diagnosis | Analyze formula errors | 🚧 |
| Pivot Table | Create pivot tables | 🚧 |
| Conditional Formatting | Set format rules | 🚧 |

### Word Features

| Feature | Description | Status |
|---------|-------------|--------|
| Get Document Info | Name, paragraphs, word count | ✅ |
| Read Text | Get document content | ✅ |
| Insert Text | At start/end/cursor | ✅ |
| Set Font | Font, size, bold, etc. | ✅ |
| Find & Replace | Batch replace text | ✅ |
| Insert Table | Create and fill tables | ✅ |
| Apply Style | Apply Word styles | ✅ |
| Generate TOC | Auto generate TOC | 🚧 |
| Insert Image | Insert and resize images | 🚧 |

### PPT Features

| Feature | Description | Status |
|---------|-------------|--------|
| Get Presentation Info | Name, slide count, shapes | ✅ |
| Add Slide | Multiple layouts | ✅ |
| Set Title | Modify slide title | ✅ |
| Add Text Box | Custom position and style | ✅ |
| Unify Font | Consistent fonts | ✅ |
| Beautify Slide | Business/Tech/Creative/Minimal | ✅ |
| Add Shape | Insert shapes | 🚧 |
| Add Animation | Enter/exit animations | 🚧 |
| Set Theme | Apply PPT themes | 🚧 |

### Common Features

| Feature | Description | Status |
|---------|-------------|--------|
| Save File | Save current document | ✅ |
| Format Conversion | Word/Excel/PPT conversion | 🚧 |

> ✅ Completed | 🚧 In Development

---

## Architecture

```
Claude Code → MCP Server (Node.js) → PowerShell COM → WPS Office
```

- **MCP Server**: 29 tools handling AI requests
- **COM Bridge**: PowerShell calls WPS COM interfaces (Ket/Kwps/Kwpp)
- **WPS Add-in**: Shows connection status

---

## Project Structure

```
WPS_Skills/
├── wps-office-mcp/          # MCP Server (Core)
│   ├── src/                 # TypeScript source
│   ├── dist/                # Compiled output
│   ├── scripts/             # PowerShell COM bridge
│   │   └── wps-com.ps1      # COM operations script
│   └── package.json
├── wps-claude-addon/        # WPS Add-in
│   ├── ribbon.xml           # Ribbon config
│   └── js/main.js           # Add-in logic
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

- [ ] **macOS Support** - Cross-platform compatibility
- [ ] **Excel Formula Diagnosis** - Analyze errors, suggest fixes
- [ ] **Excel Pivot Tables** - Create and manipulate pivot tables
- [ ] **Excel Conditional Formatting** - Set format rules
- [ ] **Word TOC Generation** - Auto generate table of contents
- [ ] **Word Insert Image** - Insert and position images
- [ ] **PPT Animations** - Enter, exit, emphasis animations
- [ ] **PPT Themes** - Apply built-in themes

### Mid-term (v1.2)

- [ ] **Cross-app Conversion** - Word/Excel/PPT conversion
- [ ] **Word to PPT** - Generate PPT from Word outline
- [ ] **Batch Conversion** - Bulk file format conversion
- [ ] **Batch Watermark** - Add text/image watermarks
- [ ] **Mail Merge** - Word mail merge functionality
- [ ] **Array Formulas** - Advanced formula support

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
