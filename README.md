# WPS Office 智能助手

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Skills-Framework-purple?style=flat-square" alt="Skills Framework">
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
帮我安装 WPS Skills，安装指南在这里：https://github.com/lc2panda/wps-mcp/blob/main/INSTALL.md
```

Claude Code 会自动读取安装指南并完成所有步骤！

> ⚠️ 前提：请先安装 [WPS Office](https://www.wps.cn/)

---

## 📖 项目简介

WPS Office 智能助手是一个基于 Claude AI 的自然语言办公自动化工具。采用 **Anthropic 官方标准的 MCP + Skills 双层架构**，让您可以用自然语言直接操控 WPS Office，告别繁琐的菜单操作和公式记忆。

### ✨ 核心特性

- 🗣️ **自然语言操作** - 用中文描述需求，AI 自动执行
- 📊 **全套办公支持** - Excel、Word、PPT 三大组件全覆盖
- 🔢 **公式智能生成** - 描述计算需求，自动生成公式
- 🎨 **一键美化** - PPT配色、字体统一，专业设计
- 🧠 **Skills 智能指导** - 4个专业Skills教会AI如何完成任务
- 🔧 **196个MCP工具** - 完整的底层工具能力

### 🎯 使用示例

```bash
# Excel 操作
用户: 帮我写个公式查产品价格
用户: 创建一个销售数据透视表
用户: 把B列大于100的单元格标红

# Word 操作
用户: 帮我生成文档目录
用户: 把全文字体改成宋体12号
用户: 插入一个3行4列的表格

# PPT 操作
用户: 用商务风格美化这页PPT
用户: 帮我画个项目流程图
用户: 创建一组KPI数据卡片
```

---

## 📋 系统要求

| 项目 | Windows | macOS |
|------|---------|-------|
| 操作系统 | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 或更高版本 | Mac 版最新版 |
| Node.js | 18.0.0 或更高版本 | 18.0.0 或更高版本 |
| Claude Code | 最新版本 | 最新版本 |
| **功能支持** | ⚠️ 基础功能（~25个方法） | ✅ 完整功能（196个方法） |

> ⚠️ **Windows 适配说明**：Windows 版本目前通过 PowerShell COM 桥接实现，已支持 Excel/Word/PPT 的基础操作。高级功能（透视表、条件格式、流程图、3D效果等）正在适配中。macOS 版本功能完整。

---

<details>
<summary><b>📦 手动安装（点击展开）</b></summary>

### 方式一：一键脚本

```bash
git clone https://github.com/lc2panda/wps-mcp.git
cd wps-mcp

# Windows
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1

# macOS
./scripts/auto-install-mac.sh
```

### 方式二：手动步骤

1. **安装依赖并编译**
   ```bash
   cd wps-office-mcp
   npm install
   npm run build
   ```

2. **配置 MCP Server**
   ```bash
   claude mcp add wps-office node /path/to/wps-mcp/wps-office-mcp/dist/index.js
   ```

3. **注册 Skills（创建软链接到全局目录）**
   ```bash
   mkdir -p ~/.claude/skills
   ln -sf /path/to/wps-mcp/skills/wps-excel ~/.claude/skills/wps-excel
   ln -sf /path/to/wps-mcp/skills/wps-word ~/.claude/skills/wps-word
   ln -sf /path/to/wps-mcp/skills/wps-ppt ~/.claude/skills/wps-ppt
   ln -sf /path/to/wps-mcp/skills/wps-office ~/.claude/skills/wps-office
   ```

4. **安装 WPS 加载项** - 参考 INSTALL.md

5. **重启 Claude Code 和 WPS Office**

</details>

---

## 🔧 技术架构

采用 **Anthropic 官方标准的 MCP + Skills 双层架构**：

```
┌─────────────────────────────────────────────────────────────┐
│                    用户自然语言请求                           │
│                "帮我写个VLOOKUP公式查价格"                     │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    Skills 层（指令包）                        │
│  skills/wps-excel/SKILL.md  - 教Claude怎么处理Excel任务       │
│  skills/wps-word/SKILL.md   - 教Claude怎么处理Word任务        │
│  skills/wps-ppt/SKILL.md    - 教Claude怎么处理PPT任务         │
│  skills/wps-office/SKILL.md - 教Claude怎么协调跨应用任务       │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    MCP 层（工具能力）                         │
│  wps-office-mcp/            - 196个MCP工具                   │
│  wps_get_active_workbook    - 获取当前工作簿                  │
│  wps_execute_method         - 执行具体操作                    │
│  ...                                                        │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    WPS 加载项层（执行器）                      │
│  Windows: PowerShell COM → WPS Office                       │
│  macOS: HTTP轮询 → WPS 加载项 (JS API) → WPS Office          │
└─────────────────────────────────────────────────────────────┘
```

### MCP vs Skills

| 层级 | 作用 | 内容 |
|------|------|------|
| **Skills** | 教Claude"怎么做" | 4个SKILL.md文件，包含工作流程和最佳实践 |
| **MCP** | 告诉Claude"能做什么" | 196个工具，提供底层操作能力 |

---

## 📁 项目结构

```
wps-mcp/
├── wps-office-mcp/          # MCP Server (核心服务)
│   ├── src/                 # TypeScript 源码
│   ├── dist/                # 编译输出
│   └── package.json
├── wps-claude-assistant/    # WPS 加载项 (macOS)
│   ├── main.js              # HTTP 轮询 + 所有 Handler
│   ├── manifest.xml         # 加载项清单
│   └── ribbon.xml           # 功能区配置
├── wps-claude-addon/        # WPS 加载项 (Windows)
│   ├── ribbon.xml           # 功能区配置
│   └── js/main.js           # 加载项逻辑
├── skills/                  # Claude Skills 定义
│   ├── wps-excel/SKILL.md   # Excel 技能（60+方法）
│   ├── wps-word/SKILL.md    # Word 技能（25+方法）
│   ├── wps-ppt/SKILL.md     # PPT 技能（85+方法）
│   └── wps-office/SKILL.md  # 跨应用技能
├── scripts/
│   ├── auto-install.ps1     # Windows 一键安装
│   └── auto-install-mac.sh  # macOS 一键安装
├── INSTALL.md               # Claude Code 安装指南
└── README.md
```

---

## 📖 功能列表

### Excel 功能 (86个)

| 分类 | 数量 | 功能 |
|------|------|------|
| 工作簿/工作表操作 | 12 | 打开/创建/切换/重命名/复制/移动 |
| 单元格读写 | 7 | 读写单元格/范围/公式/完整信息 |
| 格式美化 | 15 | 样式/边框/数字格式/合并/自动调整 |
| 行列操作 | 8 | 插入/删除/隐藏/显示行列 |
| 条件格式 | 3 | 添加/删除/获取条件格式 |
| 数据验证 | 3 | 添加/删除/获取数据验证 |
| 数据处理 | 10 | 排序/筛选/去重/清洗/复制/转置 |
| 图表/透视表 | 4 | 创建/更新图表和透视表 |
| 公式功能 | 5 | 设置公式/数组公式/诊断/重算 |
| 其他 | 19 | 批注/保护/命名区域/查找替换等 |

### Word 功能 (25个)

| 分类 | 功能 |
|------|------|
| 文档管理 | 获取信息/打开/切换/获取全文 |
| 文本操作 | 插入文本/查找替换 |
| 格式设置 | 字体/样式/段落 |
| 文档结构 | 目录/分页符/页眉/页脚 |
| 插入内容 | 表格/图片/超链接/书签 |
| 其他 | 批注/文档统计 |

### PPT 功能 (85个)

| 分类 | 数量 | 功能 |
|------|------|------|
| 演示文稿管理 | 5 | 创建/打开/关闭/切换 |
| 幻灯片操作 | 10 | 添加/删除/复制/移动/备注 |
| 文本框/形状 | 21 | 添加/删除/样式/阴影/渐变/边框 |
| 智能布局 | 10 | 对齐/分布/组合/连接线/箭头 |
| 图片/表格/图表 | 12 | 插入/设置样式 |
| 数据可视化 | 6 | KPI卡片/进度条/仪表盘/环形图 |
| 流程图/架构图 | 3 | 流程图/组织架构图/时间轴 |
| 动画/切换 | 9 | 动画/强调/切换效果 |
| 母版/3D效果 | 7 | 母版操作/3D旋转/深度/材质 |
| 其他 | 2 | 演示放映 |

---

## ❓ 常见问题

### Q: Claude助手选项卡没有出现？

**A:** 检查以下几点：
1. 确认加载项文件夹名称以 `_` 结尾
2. 确认 `publish.xml` 已正确配置
3. 重启 WPS Office

### Q: Skills 没有加载？

**A:** 检查软链接是否存在：
```bash
ls ~/.claude/skills/
```

如果为空，手动创建：
```bash
mkdir -p ~/.claude/skills
ln -sf /path/to/wps-mcp/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf /path/to/wps-mcp/skills/wps-word ~/.claude/skills/wps-word
ln -sf /path/to/wps-mcp/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf /path/to/wps-mcp/skills/wps-office ~/.claude/skills/wps-office
```

然后重启 Claude Code。

### Q: MCP Server 连接失败？

**A:** 排查步骤：
1. 确认已执行 `npm run build`
2. 运行 `claude mcp list` 检查配置
3. 重启 Claude Code

---

## 📋 TODO

### 近期计划 (v1.1) ✅ 已完成

- [x] macOS 兼容
- [x] Excel 公式诊断、透视表、条件格式
- [x] Word 目录生成、插入图片
- [x] PPT 动画、高端美化、6大类高级能力
- [x] 跨应用数据传递
- [x] **Skills 框架** - Anthropic 官方标准

### 中期计划 (v1.2)

- [ ] 跨应用格式转换
- [ ] Word 转 PPT
- [ ] 批量格式转换
- [ ] 邮件合并

### 长期计划 (v2.0)

- [ ] PDF 支持
- [ ] AI 内容生成
- [ ] 自动化工作流
- [ ] 企业级部署

---

## 📄 许可证

MIT License

## 👨‍💻 开发者

**lc2panda** - [GitHub](https://github.com/lc2panda)

---

<p align="center">Made with ❤️ for WPS Office users</p>
