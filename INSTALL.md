# WPS Skills 自动安装指南 (For Claude Code)

> 🤖 **本指南专为 Claude Code 设计**，Claude Code 读取后可自动执行安装。

当用户要求安装 WPS Skills 时，请**按顺序执行**以下步骤：

---

## 🍎 macOS 自动安装

### Step 1: 检测前置条件

执行以下命令检测环境：

```bash
# 检测 macOS
uname -s

# 检测 WPS Office
ls /Applications/wpsoffice.app

# 检测 Node.js 版本
node -v
```

如果 WPS Office 不存在，提示用户安装：https://www.wps.cn/product/wpsmac
如果 Node.js 未安装或版本低于 18，提示用户安装：https://nodejs.org/

### Step 2: 克隆项目

```bash
cd ~/Downloads
git clone https://github.com/lc2panda/wps-mcp.git
cd wps-mcp
```

### Step 3: 运行一键安装脚本

```bash
chmod +x scripts/auto-install-mac.sh
./scripts/auto-install-mac.sh
```

脚本会自动完成：
- ✅ 安装 npm 依赖
- ✅ 编译 TypeScript
- ✅ 安装 WPS 加载项
- ✅ 注册 MCP Server (`claude mcp add`)
- ✅ 注册 Skills 到 `~/.claude/skills/`

### Step 4: 验证安装

```bash
# 检查 MCP 是否注册
claude mcp list

# 检查 Skills 是否注册
ls ~/.claude/skills/
```

### Step 5: 提示用户

安装完成后，告诉用户：
1. **重启 Claude Code**（必须！Skills 需要重启后生效）
2. 重启 WPS Office
3. 打开 Excel/Word/PPT 文档
4. 查看 "Claude助手" 选项卡

---

## 🖥️ Windows 自动安装

### Step 1: 检测前置条件

执行以下命令检测环境：

```powershell
# 检测 WPS Office
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons"

# 检测 Node.js 版本
node -v
```

如果 WPS Office 不存在，提示用户安装：https://www.wps.cn/
如果 Node.js 未安装或版本低于 18，提示用户安装：https://nodejs.org/

### Step 2: 克隆项目

```powershell
cd $env:USERPROFILE\Downloads
git clone https://github.com/lc2panda/wps-mcp.git
cd wps-mcp
```

### Step 3: 运行一键安装脚本

```powershell
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

### Step 4: 验证安装

```powershell
# 检查 MCP 是否注册
claude mcp list

# 检查 Skills 是否注册
Get-ChildItem "$env:USERPROFILE\.claude\skills"
```

### Step 5: 提示用户

安装完成后，告诉用户：
1. **重启 Claude Code**（必须！）
2. 重启 WPS Office
3. 查看 "Claude助手" 选项卡

---

## ⚠️ 常见问题处理

### Skills 没有加载

重启 Claude Code 后检查：
```bash
ls ~/.claude/skills/
```

如果目录为空，手动创建软链接：
```bash
PROJECT_DIR=~/Downloads/wps-mcp
mkdir -p ~/.claude/skills
ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
```

### MCP Server 未注册

手动注册：
```bash
claude mcp add wps-office node ~/Downloads/wps-mcp/wps-office-mcp/dist/index.js
```

### WPS 加载项未显示

```bash
# 强制退出 WPS
pkill -f wpsoffice

# 重新启动 WPS Office
open /Applications/wpsoffice.app
```
