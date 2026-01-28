# WPS Skills 安装指南 (For Claude Code)

当用户要求安装 WPS Skills 时，请按以下步骤操作：

---

## 🖥️ Windows 安装

### 第一步：前置条件检查

1. **检测操作系统** - Windows 10/11
2. **检测 WPS Office** - 检查目录是否存在：`%APPDATA%\kingsoft\wps\jsaddons\`
   - 如果不存在，提示用户先安装 WPS Office: https://www.wps.cn/
3. **检测 Node.js** - 运行 `node -v`，需要 >= 18.0.0
   - 如果未安装，提示用户安装: https://nodejs.org/

### 第二步：克隆项目

```powershell
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
```

### 第三步：运行一键安装脚本

```powershell
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

脚本会自动完成：
- 安装 npm 依赖
- 编译 TypeScript
- 配置 Claude Code 的 settings.json
- 拷贝 WPS 加载项到正确目录
- 更新 publish.xml 注册加载项

### 第四步：提示用户

安装完成后，提示用户：
1. 重启 Claude Code
2. 重启 WPS Office
3. 在 WPS 中查看 "Claude助手" 选项卡

### 验证安装

```powershell
# 检查 MCP Server 是否编译成功
Test-Path "wps-office-mcp\dist\index.js"

# 检查 WPS 加载项是否安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_\ribbon.xml"
```

---

## 🍎 macOS 安装

### 技术架构说明

Mac版采用**反向轮询架构**：
```
Claude Code → MCP Server (HTTP服务端:58891) ← 轮询 ← WPS加载项 (HTTP客户端)
```

- MCP Server 启动 HTTP 服务器，监听端口 58891
- WPS 加载项每 500ms 轮询一次获取命令
- 根据命令类型自动切换 WPS 应用 (Excel/Word/PPT)

### 第一步：前置条件检查

1. **检测操作系统** - macOS 12+
2. **检测 WPS Office** - 检查是否存在：`/Applications/wpsoffice.app`
   - 如果不存在，提示用户先安装 WPS Office: https://www.wps.cn/product/wpsmac
3. **检测 Node.js** - 运行 `node -v`，需要 >= 18.0.0
   - 如果未安装，提示用户安装: https://nodejs.org/

### 第二步：克隆项目

```bash
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
```

### 第三步：运行一键安装脚本

```bash
./scripts/auto-install-mac.sh
```

脚本会自动完成：
- 检查所有前置条件
- 安装 npm 依赖
- 编译 TypeScript
- 拷贝 WPS 加载项到 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/`
- 更新 publish.xml 注册加载项
- **使用 `claude mcp add` 命令注册 MCP Server**

> ⚠️ **踩坑提醒**：直接编辑 `~/.claude/settings.json` 是无效的！必须使用 `claude mcp add` 命令注册 MCP Server。

### 手动配置 MCP（如果自动配置失败）

```bash
claude mcp add wps-office node /path/to/WPS_Skills/wps-office-mcp/dist/index.js
```

验证 MCP 配置：
```bash
# 查看已注册的 MCP Server
claude mcp list

# 应该看到类似输出：
# wps-office: node /path/to/wps-office-mcp/dist/index.js
```

### 第四步：提示用户

安装完成后，提示用户：
1. 重启 Claude Code
2. 重启 WPS Office
3. 打开任意 Excel/Word/PPT 文档
4. 查看 "Claude助手" 选项卡，确认状态为 "轮询中"

### 验证安装

```bash
# 检查 MCP Server 是否编译成功
ls wps-office-mcp/dist/index.js

# 检查 WPS 加载项是否安装
ls ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/

# 检查 Claude Code 配置
cat ~/.claude/settings.json | grep wps-office
```

---

## ⚠️ 常见问题

### Mac: Claude助手显示"轮询中"但命令无响应

确保：
1. MCP Server 已正确配置在 Claude Code 中
2. 重启 Claude Code 使配置生效
3. WPS 加载项已正确安装（目录名必须以 `_` 结尾）

### Mac: 加载项未加载

尝试：
1. 完全退出 WPS Office（包括后台进程）：`pkill -f wpsoffice`
2. 重新启动 WPS Office
3. 检查 publish.xml 配置是否正确

### Windows: PowerShell 执行策略错误

运行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
