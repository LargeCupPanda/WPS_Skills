# WPS Office 智能助手

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square" alt="Windows | macOS">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

<p align="center">
  <a href="./README_EN.md">English</a> | 中文
</p>

---

## 🚀 一键安装

**只需告诉 Claude Code：**

```
帮我安装 WPS Skills，安装指南在这里：https://github.com/LargeCupPanda/WPS_Skills/blob/main/INSTALL.md
```

Claude Code 会自动读取安装指南并完成所有步骤！

> ⚠️ 前提：请先安装 [WPS Office](https://www.wps.cn/)

---

## 📖 项目简介

WPS Office 智能助手是一个基于 Claude AI 的自然语言办公自动化工具。通过 MCP (Model Context Protocol) 协议，让您可以用自然语言直接操控 WPS Office，告别繁琐的菜单操作和公式记忆。

### ✨ 核心特性

- 🗣️ **自然语言操作** - 用中文描述需求，AI 自动执行
- 📊 **全套办公支持** - Excel、Word、PPT 三大组件全覆盖
- 🔢 **公式智能生成** - 描述计算需求，自动生成公式
- 🎨 **一键美化** - PPT配色、字体统一，专业设计
- 🔗 **稳定 COM 桥接** - 通过 PowerShell COM 接口，稳定可靠

### 🎯 使用示例

```bash
# Excel 操作
用户: 帮我读取当前Excel的A1到C5的数据
用户: 把B3单元格的值改成4.8
用户: 创建一个柱状图展示销售数据
用户: 按B列降序排序
用户: 计算B列的平均值、最大值、最小值

# Word 操作
用户: 在文档末尾插入一段文字
用户: 把所有的"旧公司"替换成"新公司"
用户: 插入一个3行4列的表格
用户: 把全文字体改成宋体12号

# PPT 操作
用户: 新增一页幻灯片，标题是"项目总结"
用户: 统一全文字体为微软雅黑
用户: 用商务风格美化当前页面
用户: 在第一页添加一个文本框
```

---

## 📋 系统要求

| 项目 | Windows | macOS |
|------|---------|-------|
| 操作系统 | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 或更高版本 | Mac 版最新版 |
| Node.js | 18.0.0 或更高版本 | 18.0.0 或更高版本 |
| Claude Code | 最新版本 | 最新版本 |

---

<details>
<summary><b>📦 手动安装（点击展开）</b></summary>

### 方式一：一键脚本

```powershell
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

### 方式二：手动步骤

1. **安装依赖并编译**
   ```bash
   cd wps-office-mcp
   npm install
   npm run build
   ```

2. **配置 Claude Code** - 编辑 `~/.claude/settings.json`：
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

3. **安装 WPS 加载项**
   - 复制 `wps-claude-addon` 到 `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\`
   - 在 `publish.xml` 中添加：
     ```xml
     <jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_/"/>
     ```

4. **重启 Claude Code 和 WPS Office**

</details>

---

## 📖 功能列表

### Excel 功能

| 功能 | 说明 | 状态 |
|------|------|------|
| 获取工作簿信息 | 名称、路径、工作表列表 | ✅ |
| 获取上下文 | 表头、选中单元格、使用范围 | ✅ |
| 读取单元格 | 单个或范围读取 | ✅ |
| 写入单元格 | 单个或范围写入 | ✅ |
| 设置公式 | 写入Excel公式 | ✅ |
| 排序 | 按指定列排序 | ✅ |
| 筛选 | 自动筛选 | ✅ |
| 去重 | 删除重复行 | ✅ |
| 创建图表 | 柱状图、折线图、饼图等 | ✅ |
| 公式诊断 | 分析公式错误原因 | 🚧 |
| 数据透视表 | 创建透视表 | 🚧 |
| 条件格式 | 设置条件格式 | 🚧 |

### Word 功能

| 功能 | 说明 | 状态 |
|------|------|------|
| 获取文档信息 | 名称、段落数、字数 | ✅ |
| 读取文本 | 获取文档内容 | ✅ |
| 插入文本 | 开头/末尾/光标处插入 | ✅ |
| 设置字体 | 字体、字号、粗体等 | ✅ |
| 查找替换 | 批量替换文本 | ✅ |
| 插入表格 | 创建表格并填充数据 | ✅ |
| 应用样式 | 应用Word样式 | ✅ |
| 生成目录 | 自动生成目录 | 🚧 |
| 插入图片 | 插入并调整图片 | 🚧 |

### PPT 功能

| 功能 | 说明 | 状态 |
|------|------|------|
| 获取演示文稿信息 | 名称、页数、形状列表 | ✅ |
| 新增幻灯片 | 多种布局可选 | ✅ |
| 设置标题 | 修改幻灯片标题 | ✅ |
| 添加文本框 | 自定义位置和样式 | ✅ |
| 统一字体 | 全文字体统一 | ✅ |
| 美化幻灯片 | 商务/科技/创意/简约风格 | ✅ |
| 添加形状 | 插入各种形状 | 🚧 |
| 添加动画 | 进入/退出动画 | 🚧 |
| 设置主题 | 应用PPT主题 | 🚧 |

### 通用功能

| 功能 | 说明 | 状态 |
|------|------|------|
| 保存文件 | 保存当前文档 | ✅ |
| 格式转换 | Word/Excel/PPT互转 | 🚧 |

> ✅ 已完成 | 🚧 开发中

---

## 🔧 技术架构

```
Windows:
Claude Code → MCP Server (Node.js) → PowerShell COM → WPS Office

macOS:
Claude Code → MCP Server (Node.js) → HTTP → WPS 加载项 (JS API) → WPS Office
```

- **MCP Server**: 29 个工具，处理 AI 请求
- **Windows COM 桥接**: 通过 PowerShell 调用 WPS COM 接口（Ket/Kwps/Kwpp）
- **macOS HTTP 桥接**: 通过 HTTP 调用 WPS 加载项内置服务（端口 58891）
- **WPS 加载项**: 显示连接状态，Mac 上提供 HTTP API

---

## 📁 项目结构

```
WPS_Skills/
├── wps-office-mcp/          # MCP Server (核心服务)
│   ├── src/                 # TypeScript 源码
│   ├── dist/                # 编译输出
│   ├── scripts/             # PowerShell COM 桥接脚本 (Windows)
│   │   └── wps-com.ps1      # COM操作脚本
│   └── package.json
├── wps-claude-addon/        # WPS 加载项 (Windows)
│   ├── ribbon.xml           # 功能区配置
│   └── js/main.js           # 加载项逻辑
├── wps-claude-assistant/    # WPS 加载项 (macOS)
│   ├── main.js              # HTTP 轮询 + 所有 Handler
│   ├── manifest.xml         # 加载项清单
│   └── ribbon.xml           # 功能区配置
├── scripts/
│   ├── auto-install.ps1     # Windows 一键安装
│   └── auto-install-mac.sh  # macOS 一键安装
├── skills/                  # Claude Skills 定义
├── docs/                    # 设计文档（私有）
└── README.md
```

---

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
2. 对应的应用已打开（操作Excel需打开Excel，操作Word需打开Word）

---

## 📋 TODO

### 近期计划 (v1.1)

- [x] **macOS 兼容** - 支持 macOS 平台 ✅ 已完成
- [ ] **Excel 公式诊断** - 分析公式错误，提供修复建议
- [ ] **Excel 数据透视表** - 创建和操作透视表
- [ ] **Excel 条件格式** - 设置条件格式规则
- [ ] **Word 生成目录** - 自动生成文档目录
- [ ] **Word 插入图片** - 插入并调整图片位置
- [ ] **PPT 添加动画** - 进入、退出、强调动画
- [ ] **PPT 设置主题** - 应用内置主题

### 中期计划 (v1.2)

- [ ] **跨应用格式转换** - Word/Excel/PPT 互转
- [ ] **Word 转 PPT** - 根据Word大纲生成PPT
- [ ] **批量格式转换** - 批量转换文件格式
- [ ] **批量添加水印** - 批量添加文字/图片水印
- [ ] **邮件合并** - Word邮件合并功能
- [ ] **高级公式** - 数组公式支持

### 长期计划 (v2.0)

- [ ] **PDF 支持** - 支持 WPS PDF 操作
- [ ] **脑图支持** - 支持 WPS 脑图
- [ ] **AI 内容生成** - 智能生成文档内容
- [ ] **多文档协同** - 跨文档操作
- [ ] **自动化工作流** - 定义和执行工作流
- [ ] **云端协作** - 支持 WPS 云文档
- [ ] **企业级部署** - 企业部署方案

---

## 📄 许可证

MIT License

## 👨‍💻 开发者

**熊猫大侠** - [GitHub](https://github.com/LargeCupPanda)

---

<p align="center">Made with ❤️ for WPS Office users</p>
