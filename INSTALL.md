# WPS Skills 安装指南 (For Claude Code)

当用户要求安装 WPS Skills 时，请按以下步骤操作：

## 前置条件检查

1. **检测操作系统** - 目前仅支持 Windows
2. **检测 WPS Office** - 检查目录是否存在：`%APPDATA%\kingsoft\wps\jsaddons\`
   - 如果不存在，提示用户先安装 WPS Office: https://www.wps.cn/
3. **检测 Node.js** - 运行 `node -v`，需要 >= 18.0.0
   - 如果未安装或版本过低，提示用户安装: https://nodejs.org/

## 自动安装步骤

如果前置条件满足，执行以下命令：

```powershell
# 进入项目目录
cd <项目路径>

# 运行一键安装脚本
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

或者手动执行：

```powershell
# 1. 安装依赖并编译
cd wps-office-mcp
npm install
npm run build

# 2. 配置 Claude Code (修改 ~/.claude/settings.json)
# 添加 mcpServers.wps-office 配置

# 3. 拷贝 WPS 加载项
# 源: wps-claude-addon/
# 目标: %APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\

# 4. 更新 publish.xml 注册加载项
```

## 安装完成后

提示用户：
1. 重启 Claude Code
2. 重启 WPS Office
3. 在 WPS 中查看 "Claude助手" 选项卡

## 验证安装

```powershell
# 检查 MCP Server 是否编译成功
Test-Path "<项目路径>\wps-office-mcp\dist\index.js"

# 检查 WPS 加载项是否安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_\ribbon.xml"

# 检查 Claude Code 配置
Get-Content "$env:USERPROFILE\.claude\settings.json"
```
