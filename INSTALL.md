# WPS Skills 安装指南 (For Claude Code)

当用户要求安装 WPS Skills 时，请按以下步骤操作：

## 第一步：前置条件检查

1. **检测操作系统** - 目前仅支持 Windows
2. **检测 WPS Office** - 检查目录是否存在：`%APPDATA%\kingsoft\wps\jsaddons\`
   - 如果不存在，提示用户先安装 WPS Office: https://www.wps.cn/
3. **检测 Node.js** - 运行 `node -v`，需要 >= 18.0.0
   - 如果未安装，提示用户安装: https://nodejs.org/

## 第二步：克隆项目

```powershell
git clone https://github.com/LargeCupPanda/WPS_Skills.git
cd WPS_Skills
```

## 第三步：运行一键安装脚本

```powershell
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

脚本会自动完成：
- 安装 npm 依赖
- 编译 TypeScript
- 配置 Claude Code 的 settings.json
- 拷贝 WPS 加载项到正确目录
- 更新 publish.xml 注册加载项

## 第四步：提示用户

安装完成后，提示用户：
1. 重启 Claude Code
2. 重启 WPS Office
3. 在 WPS 中查看 "Claude助手" 选项卡

## 验证安装

```powershell
# 检查 MCP Server 是否编译成功
Test-Path "wps-office-mcp\dist\index.js"

# 检查 WPS 加载项是否安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_\ribbon.xml"
```
