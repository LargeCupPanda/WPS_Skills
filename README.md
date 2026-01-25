# WPS Office 智能助手 | WPS Office AI Assistant

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square" alt="Windows">
</p>

> 🇨🇳 [中文](#中文文档) | 🇺🇸 [English](#english-documentation)

---

# 中文文档

## 📖 项目简介

WPS Office 智能助手是一个基于 Claude AI 的自然语言办公自动化工具。通过 MCP (Model Context Protocol) 协议，让您可以用自然语言直接操控 WPS Office，告别繁琐的菜单操作和公式记忆。

### ✨ 核心特性

- **自然语言操作** - 用中文描述需求，AI 自动执行
- **全套办公支持** - Excel、Word、PPT 三大组件全覆盖
- **公式智能生成** - 描述计算需求，自动生成公式
- **稳定 COM 桥接** - 通过 PowerShell COM 接口，稳定可靠

### 🎯 使用示例

```
# Excel 操作
用户: 帮我读取当前Excel的A1到C5的数据
用户: 把B3单元格的值改成4.8
用户: 创建一个柱状图展示销售数据
用户: 按B列降序排序

# Word 操作
用户: 在文档末尾插入一段文字
用户: 把所有的"旧公司"替换成"新公司"
用户: 插入一个3行4列的表格

# PPT 操作
用户: 新增一页幻灯片，标题是"项目总结"
用户: 统一全文字体为微软雅黑
用户: 用商务风格美化当前页面
```

## 📋 系统要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows 10/11 |
| WPS Office | 2019 或更高版本 |
| Node.js | 18.0.0 或更高版本 |
| Claude Code | 最新版本 |

## 🚀 安装步骤

### 第一步：克隆项目

```bash
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
```

### 第二步：安装 MCP Server 依赖

```bash
cd wps-office-mcp
npm install
npm run build
```

### 第三步：配置 Claude Code

找到 Claude Code 配置文件：
- 路径：`C:\Users\<用户名>\.claude\settings.json`

添加 MCP Server 配置：

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

> ⚠️ 注意：请将路径替换为您的实际项目路径，Windows 路径使用双反斜杠 `\\`

### 第四步：安装 WPS 加载项

1. 找到 WPS 加载项目录：
   ```
   C:\Users\<用户名>\AppData\Roaming\kingsoft\wps\jsaddons\
   ```

2. 复制 `wps-claude-addon` 文件夹到该目录，并重命名为 `wps-claude-addon_`（注意末尾下划线）

3. 编辑 `publish.xml` 文件，添加加载项注册：
   ```xml
   <jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_/"/>
   ```

### 第五步：重启并验证

1. **重启 Claude Code** - 加载新的 MCP Server 配置
2. **重启 WPS Office** - 加载新的加载项
3. **验证安装**：
   - 在 WPS 中查看是否有 "Claude助手" 选项卡
   - 点击 "连接状态" 按钮查看状态

## 📖 使用方法

### 基本操作

在 Claude Code 中直接用自然语言描述需求：

```
# 读取数据
帮我读取当前Excel的A1到D10的数据

# 修改单元格
把C2单元格的值改成"测试数据"

# 获取工作簿信息
当前打开的是什么文件？有几个工作表？
```

### 支持的功能

| 应用 | 功能类别 | 支持操作 |
|------|----------|----------|
| **Excel** | 数据读写 | 单元格值、范围数据、工作簿信息、上下文获取 |
| **Excel** | 数据处理 | 公式设置、排序、筛选、去重、创建图表 |
| **Word** | 文档操作 | 获取文档信息、读取文本、插入文本 |
| **Word** | 格式编辑 | 字体设置、查找替换、插入表格、应用样式 |
| **PPT** | 幻灯片 | 获取演示文稿信息、新增幻灯片、设置标题 |
| **PPT** | 美化功能 | 添加文本框、统一字体、配色美化 |
| **通用** | 文件操作 | 保存文件 |

## ❓ 常见问题

### Q: Claude助手选项卡没有出现？

**A:** 检查以下几点：
1. 确认加载项文件夹名称以 `_` 结尾
2. 确认 `publish.xml` 已正确配置
3. 重启 WPS Office

### Q: MCP Server 连接失败？

**A:** 排查步骤：
1. 确认 `settings.json` 路径配置正确
2. 确认已执行 `npm run build`
3. 重启 Claude Code

### Q: 操作 WPS 时提示连接错误？

**A:** 确保：
1. WPS Office 已启动并打开了文档
2. 加载项已正确加载（查看Claude助手选项卡）

## 📁 项目结构

```
WPS_Skills/
├── wps-office-mcp/          # MCP Server (核心服务)
│   ├── src/                 # TypeScript 源码
│   ├── dist/                # 编译输出
│   ├── scripts/             # PowerShell COM 桥接脚本
│   └── package.json
├── wps-claude-addon/        # WPS 加载项
│   ├── ribbon.xml           # 功能区配置
│   └── js/main.js           # 加载项逻辑
├── skills/                  # Claude Skills 定义
└── README.md
```

## 🔧 技术架构

```
Claude Code → MCP Server (Node.js) → PowerShell COM → WPS Office
```

- **MCP Server**: 29 个工具，处理 AI 请求
- **COM 桥接**: 通过 PowerShell 调用 WPS COM 接口
- **WPS 加载项**: 显示连接状态

## 📄 许可证

MIT License

## 👨‍💻 开发者

**熊猫大侠** - [GitHub](https://github.com/LargeCupPanda)

---

# English Documentation

## 📖 Introduction

WPS Office AI Assistant is a natural language office automation tool powered by Claude AI. Through the MCP (Model Context Protocol), you can control WPS Office using natural language, eliminating the need for complex menu navigation and formula memorization.

### ✨ Key Features

- **Natural Language Control** - Describe your needs in plain language, AI executes automatically
- **Full Office Suite Support** - Excel, Word, and PPT all covered
- **Smart Formula Generation** - Describe calculations, get formulas automatically
- **Stable COM Bridge** - Reliable PowerShell COM interface

### 🎯 Usage Examples

```
# Excel Operations
User: Read data from A1 to C5 in the current Excel
User: Change the value of cell B3 to 4.8
User: Create a bar chart for the sales data
User: Sort by column B in descending order

# Word Operations
User: Insert text at the end of the document
User: Replace all "old company" with "new company"
User: Insert a 3x4 table

# PPT Operations
User: Add a new slide with title "Project Summary"
User: Unify all fonts to Microsoft YaHei
User: Beautify current slide with business style
```

## 📋 System Requirements

| Item | Requirement |
|------|-------------|
| OS | Windows 10/11 |
| WPS Office | 2019 or later |
| Node.js | 18.0.0 or later |
| Claude Code | Latest version |

## 🚀 Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
```

### Step 2: Install MCP Server Dependencies

```bash
cd wps-office-mcp
npm install
npm run build
```

### Step 3: Configure Claude Code

Locate the Claude Code configuration file:
- Path: `C:\Users\<username>\.claude\settings.json`

Add MCP Server configuration:

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

> ⚠️ Note: Replace the path with your actual project path. Use double backslashes `\\` for Windows paths.

### Step 4: Install WPS Add-in

1. Locate the WPS add-ins directory:
   ```
   C:\Users\<username>\AppData\Roaming\kingsoft\wps\jsaddons\
   ```

2. Copy the `wps-claude-addon` folder to this directory and rename it to `wps-claude-addon_` (note the trailing underscore)

3. Edit the `publish.xml` file to register the add-in:
   ```xml
   <jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_/"/>
   ```

### Step 5: Restart and Verify

1. **Restart Claude Code** - Load the new MCP Server configuration
2. **Restart WPS Office** - Load the new add-in
3. **Verify Installation**:
   - Check for the "Claude助手" tab in WPS
   - Click "连接状态" button to view status

## 📖 Usage

### Basic Operations

Use natural language in Claude Code:

```
# Read data
Read data from A1 to D10 in the current Excel

# Modify cells
Change the value of C2 to "Test Data"

# Get workbook info
What file is currently open? How many sheets?
```

### Supported Features

| App | Category | Operations |
|-----|----------|------------|
| **Excel** | Data R/W | Cell values, range data, workbook info, context |
| **Excel** | Processing | Formulas, sort, filter, remove duplicates, charts |
| **Word** | Document | Get document info, read text, insert text |
| **Word** | Formatting | Font settings, find/replace, insert table, styles |
| **PPT** | Slides | Get presentation info, add slides, set title |
| **PPT** | Beautify | Add textbox, unify fonts, color schemes |
| **Common** | File | Save file |

## ❓ FAQ

### Q: The Claude Assistant tab doesn't appear?

**A:** Check the following:
1. Confirm the add-in folder name ends with `_`
2. Confirm `publish.xml` is correctly configured
3. Restart WPS Office

### Q: MCP Server connection failed?

**A:** Troubleshooting steps:
1. Verify the path in `settings.json` is correct
2. Confirm `npm run build` was executed
3. Restart Claude Code

### Q: Connection error when operating WPS?

**A:** Ensure:
1. WPS Office is running with a document open
2. The add-in is properly loaded (check for Claude Assistant tab)

## 📁 Project Structure

```
WPS_Skills/
├── wps-office-mcp/          # MCP Server (Core Service)
│   ├── src/                 # TypeScript source
│   ├── dist/                # Build output
│   ├── scripts/             # PowerShell COM bridge scripts
│   └── package.json
├── wps-claude-addon/        # WPS Add-in
│   ├── ribbon.xml           # Ribbon configuration
│   └── js/main.js           # Add-in logic
├── skills/                  # Claude Skills definitions
└── README.md
```

## 🔧 Architecture

```
Claude Code → MCP Server (Node.js) → PowerShell COM → WPS Office
```

- **MCP Server**: 29 tools for handling AI requests
- **COM Bridge**: PowerShell calls to WPS COM interface
- **WPS Add-in**: Connection status display

## 📄 License

MIT License

## 👨‍💻 Developer

**熊猫大侠 (LargeCupPanda)** - [GitHub](https://github.com/LargeCupPanda)

---

<p align="center">Made with ❤️ for WPS Office users</p>
