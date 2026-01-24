# ============================================================================
# WPS-Claude-Skills 安装脚本 (Windows PowerShell)
# ============================================================================
# 作者：老王 (隔壁老王团队 DevOps - 张大牛)
#
# 艹，这个SB脚本负责把整个项目环境搭建起来
# 包括检查Node.js、npm、编译TypeScript、复制加载项到WPS目录
#
# 用法：以管理员身份运行 PowerShell，执行 .\scripts\install.ps1
#
# 参数：
#   -SkipNodeCheck  跳过Node.js版本检查（你确定你知道自己在干嘛？）
#   -SkipWpsCheck   跳过WPS安装检查（没WPS你装个锤子）
#   -Dev            开发模式（创建符号链接而不是复制文件）
# ============================================================================

#Requires -Version 5.1

param(
    [switch]$SkipNodeCheck,
    [switch]$SkipWpsCheck,
    [switch]$Dev
)

# ============================================================================
# 老王专用颜色输出函数
# 艹，Windows的控制台颜色真TM难搞
# ============================================================================
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-Success { Write-ColorOutput $args[0] "Green" }
function Write-Warn { Write-ColorOutput $args[0] "Yellow" }  # 改名避免和系统冲突
function Write-Err { Write-ColorOutput $args[0] "Red" }       # 改名避免和系统冲突
function Write-Info { Write-ColorOutput $args[0] "Cyan" }

# ============================================================================
# 打印老王风格的 Banner
# 虽然花里胡哨，但用户体验还是要有的嘛
# ============================================================================
function Show-Banner {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                                                           ║" -ForegroundColor Cyan
    Write-Host "║         WPS-Claude-Skills 安装程序 v1.0                   ║" -ForegroundColor Cyan
    Write-Host "║                                                           ║" -ForegroundColor Cyan
    Write-Host "║         让 WPS Office 说人话的智能助手                    ║" -ForegroundColor Cyan
    Write-Host "║         (老王出品，必属精品，不接受反驳)                  ║" -ForegroundColor Cyan
    Write-Host "║                                                           ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# ============================================================================
# 获取项目根目录
# 这个函数用来定位项目根目录，别TM瞎改路径
# ============================================================================
function Get-ProjectRoot {
    $scriptPath = $PSScriptRoot
    return (Get-Item $scriptPath).Parent.FullName
}

# ============================================================================
# 检查 Node.js 版本
# 没有Node.js >= 18.0.0，你玩个锤子？
# ============================================================================
function Test-NodeVersion {
    Write-Info ">>> 检查 Node.js 版本..."

    try {
        $nodeVersion = node --version 2>$null
        if (-not $nodeVersion) {
            throw "Node.js 未安装"
        }

        # 提取版本号 (去掉 v 前缀)
        $version = $nodeVersion.TrimStart('v')
        $majorVersion = [int]($version.Split('.')[0])

        if ($majorVersion -lt 18) {
            Write-Err "✗ 艹！Node.js 版本过低: $nodeVersion"
            Write-Err "  需要 Node.js >= 18.0.0，你这版本太老了，升级去！"
            Write-Info "  下载地址: https://nodejs.org"
            return $false
        }

        Write-Success "✓ Node.js 版本: $nodeVersion (这个可以)"
        return $true
    }
    catch {
        Write-Err "✗ Node.js 未安装或无法访问"
        Write-Info "  请从 https://nodejs.org 下载安装 Node.js 18+ 版本"
        Write-Info "  没有Node.js你让老子怎么帮你编译TypeScript？"
        return $false
    }
}

# ============================================================================
# 检查 npm 是否可用
# npm都没有，你装的什么野鸡Node.js？
# ============================================================================
function Test-NpmAvailable {
    Write-Info ">>> 检查 npm 是否可用..."

    try {
        $npmVersion = npm --version 2>$null
        if (-not $npmVersion) {
            throw "npm 未安装"
        }

        Write-Success "✓ npm 版本: $npmVersion"
        return $true
    }
    catch {
        Write-Err "✗ npm 未安装或无法访问"
        Write-Err "  你这Node.js装的有问题啊，npm都没有？"
        return $false
    }
}

# ============================================================================
# 检查 WPS Office 是否安装
# 没WPS你装这个加载项干嘛？给空气用？
# ============================================================================
function Test-WpsInstallation {
    Write-Info ">>> 检查 WPS Office 安装..."

    # WPS 可能的安装路径，这些憨批路径要全都检查一遍
    $wpsPaths = @(
        "${env:ProgramFiles}\Kingsoft\WPS Office\ksolaunch.exe",
        "${env:ProgramFiles(x86)}\Kingsoft\WPS Office\ksolaunch.exe",
        "${env:LOCALAPPDATA}\Kingsoft\WPS Office\ksolaunch.exe",
        # 新版WPS可能在这些路径
        "${env:ProgramFiles}\Kingsoft\WPS Office\office6\wps.exe",
        "${env:ProgramFiles(x86)}\Kingsoft\WPS Office\office6\wps.exe"
    )

    foreach ($path in $wpsPaths) {
        if (Test-Path $path) {
            Write-Success "✓ WPS Office 已安装: $path"
            return $true
        }
    }

    # 尝试从注册表获取，这是老王的骚操作
    try {
        $regPaths = @(
            "HKLM:\SOFTWARE\Kingsoft\Office",
            "HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office",
            "HKCU:\SOFTWARE\Kingsoft\Office"
        )

        foreach ($regPath in $regPaths) {
            $regValue = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            if ($regValue -and $regValue.InstallRoot) {
                Write-Success "✓ WPS Office 已安装: $($regValue.InstallRoot)"
                return $true
            }
        }
    }
    catch {}

    Write-Warn "⚠ 未检测到 WPS Office 安装"
    Write-Warn "  请确保已安装 WPS Office 2019 或更高版本"
    Write-Info "  下载地址: https://www.wps.cn"
    Write-Warn "  没有WPS你装这个加载项给谁用？"
    return $false
}

# ============================================================================
# 安装 MCP Server 依赖并编译 TypeScript
# 这一步是核心，编译不过老子把键盘都给你砸了
# ============================================================================
function Install-McpServerDependencies {
    $projectRoot = Get-ProjectRoot
    $mcpPath = Join-Path $projectRoot "wps-office-mcp"

    Write-Info ">>> 安装 MCP Server 依赖..."

    if (-not (Test-Path $mcpPath)) {
        Write-Err "✗ MCP Server 目录不存在: $mcpPath"
        Write-Err "  你TM把项目文件删了？"
        return $false
    }

    Push-Location $mcpPath
    try {
        # 第一步：npm install
        Write-Info "  执行 npm install..."
        $result = npm install 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Err "✗ npm install 失败，这憨批依赖出问题了"
            Write-Err $result
            return $false
        }
        Write-Success "✓ MCP Server 依赖安装完成"

        # 第二步：编译 TypeScript
        Write-Info "  编译 TypeScript..."
        $result = npm run build 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Err "✗ TypeScript 编译失败"
            Write-Err "  艹，编译都过不了，你写的什么SB代码？"
            Write-Err $result
            return $false
        }
        Write-Success "✓ TypeScript 编译完成"

        return $true
    }
    catch {
        Write-Err "✗ 安装过程出错: $_"
        Write-Err "  这什么憨批错误，老子从来没见过"
        return $false
    }
    finally {
        Pop-Location
    }
}

# ============================================================================
# 复制 WPS 加载项到 WPS 加载项目录
# 注意：是复制整个目录，不是创建什么配置文件！
# Windows路径: %APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon\
# ============================================================================
function Copy-WpsAddon {
    $projectRoot = Get-ProjectRoot
    $addonSource = Join-Path $projectRoot "wps-claude-addon"

    Write-Info ">>> 复制 WPS 加载项到 WPS 目录..."

    if (-not (Test-Path $addonSource)) {
        Write-Err "✗ WPS 加载项源目录不存在: $addonSource"
        Write-Err "  你把加载项目录删了？脑子被门夹了？"
        return $false
    }

    # WPS 加载项目标目录 (Windows)
    # 注意这个路径，老王查了半天WPS文档才找到的
    $wpsAddonsPath = Join-Path $env:APPDATA "kingsoft\wps\jsaddons"
    $addonTarget = Join-Path $wpsAddonsPath "wps-claude-addon"

    # 创建目标目录（如果不存在）
    if (-not (Test-Path $wpsAddonsPath)) {
        Write-Info "  创建 WPS 加载项目录: $wpsAddonsPath"
        New-Item -ItemType Directory -Path $wpsAddonsPath -Force | Out-Null
    }

    # 如果目标已存在，先删除（避免残留文件搞事情）
    if (Test-Path $addonTarget) {
        Write-Info "  删除旧版本加载项..."
        Remove-Item -Path $addonTarget -Recurse -Force
    }

    try {
        if ($Dev) {
            # 开发模式：创建符号链接（方便调试，改了源文件立即生效）
            Write-Info "  开发模式：创建符号链接..."
            # Windows创建目录符号链接需要管理员权限
            $null = cmd /c mklink /D "`"$addonTarget`"" "`"$addonSource`"" 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warn "  符号链接创建失败，可能需要管理员权限"
                Write-Info "  降级为复制模式..."
                Copy-Item -Path $addonSource -Destination $addonTarget -Recurse -Force
            } else {
                Write-Success "✓ 已创建符号链接: $addonTarget -> $addonSource"
            }
        }
        else {
            # 正常模式：复制整个目录
            Write-Info "  复制加载项文件..."
            Copy-Item -Path $addonSource -Destination $addonTarget -Recurse -Force
            Write-Success "✓ WPS 加载项已复制到: $addonTarget"
        }

        # 显示复制的文件数量，让用户安心
        $fileCount = (Get-ChildItem -Path $addonTarget -Recurse -File).Count
        Write-Info "  共复制 $fileCount 个文件"

        return $true
    }
    catch {
        Write-Err "✗ 复制加载项失败: $_"
        Write-Err "  这什么鬼？复制文件都能失败？"
        return $false
    }
}

# ============================================================================
# 生成 Claude Code MCP 配置示例
# 告诉用户怎么配置Claude Code，省得他们来烦老子
# ============================================================================
function New-ClaudeConfig {
    $projectRoot = Get-ProjectRoot
    $mcpDistPath = Join-Path $projectRoot "wps-office-mcp\dist\index.js"

    Write-Info ">>> 生成 Claude Code 配置..."

    # 把Windows路径转成正斜杠，给Claude Code用
    $mcpDistPathFormatted = $mcpDistPath.Replace('\', '/')

    $claudeConfig = @{
        mcpServers = @{
            "wps-office" = @{
                command = "node"
                args = @($mcpDistPathFormatted)
                env = @{
                    WPS_ADDON_PORT = "58080"
                }
            }
        }
    }

    $configJson = $claudeConfig | ConvertTo-Json -Depth 4

    Write-Host ""
    Write-Success "✓ Claude Code MCP 配置示例:"
    Write-Host ""
    Write-Host $configJson -ForegroundColor Yellow
    Write-Host ""
    Write-Info "  请将上述配置添加到 ~/.claude/settings.json 或者项目根目录的 .mcp.json"
    Write-Info "  不会配置？去看README.md，老王都写好了！"

    return $true
}

# ============================================================================
# 显示安装完成信息和使用说明
# 告诉用户下一步该干嘛，省得他们抓瞎
# ============================================================================
function Show-CompletionMessage {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Success "   安装完成！老王出品，必属精品！"
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""

    Write-Info "下一步操作："
    Write-Host "  1. 重启 WPS Office（必须重启才能加载新的加载项）" -ForegroundColor White
    Write-Host "  2. 配置 Claude Code（参考上方配置示例）" -ForegroundColor White
    Write-Host "  3. 启动 MCP Server（开发时用）:" -ForegroundColor White
    Write-Host "       cd wps-office-mcp && npm start" -ForegroundColor Yellow
    Write-Host "  4. 在 Claude Code 中使用:" -ForegroundColor White
    Write-Host "       /wps-excel 帮我写个求和公式" -ForegroundColor Yellow
    Write-Host ""

    Write-Info "遇到问题？"
    Write-Host "  - 检查WPS是否正确安装并重启" -ForegroundColor White
    Write-Host "  - 检查Node.js版本是否 >= 18.0.0" -ForegroundColor White
    Write-Host "  - 查看项目README.md获取更多帮助" -ForegroundColor White
    Write-Host "  - 实在不行？找老王，但老王可能会骂你（开玩笑的）" -ForegroundColor White
    Write-Host ""
}

# ============================================================================
# 主安装流程
# 把上面所有步骤串起来，一步一步执行
# ============================================================================
function Start-Installation {
    Show-Banner

    $allPassed = $true

    # ========== 第一步：检查 Node.js 和 npm ==========
    if (-not $SkipNodeCheck) {
        if (-not (Test-NodeVersion)) {
            $allPassed = $false
        }
        if (-not (Test-NpmAvailable)) {
            $allPassed = $false
        }
    }
    else {
        Write-Warn ">>> 跳过 Node.js 检查（你最好知道自己在干嘛）"
    }

    # ========== 第二步：检查 WPS Office ==========
    if (-not $SkipWpsCheck) {
        if (-not (Test-WpsInstallation)) {
            # WPS 检查失败不阻止安装，但要警告
            Write-Warn "  没检测到WPS，但安装会继续"
            Write-Warn "  安装完记得装WPS，不然加载项没地方用！"
        }
    }
    else {
        Write-Warn ">>> 跳过 WPS 检查"
    }

    # ========== 如果基础检查失败，询问是否继续 ==========
    if (-not $allPassed) {
        Write-Host ""
        $continue = Read-Host "基础环境检查未通过，是否继续安装? (y/N)"
        if ($continue -ne 'y' -and $continue -ne 'Y') {
            Write-Err "安装已取消，去把环境配好再来！"
            exit 1
        }
        Write-Warn "好吧，你坚持继续，出问题别怪老子..."
    }

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Info "开始安装依赖和配置..."
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    # ========== 第三步：安装 MCP Server 依赖并编译 ==========
    if (-not (Install-McpServerDependencies)) {
        Write-Err "MCP Server 依赖安装或编译失败"
        Write-Err "艹，这一步都过不了，后面也别装了！"
        exit 1
    }

    # ========== 第四步：复制 WPS 加载项 ==========
    if (-not (Copy-WpsAddon)) {
        Write-Err "WPS 加载项复制失败"
        Write-Err "复制文件都能失败，检查一下权限问题！"
        exit 1
    }

    # ========== 第五步：生成 Claude Code 配置示例 ==========
    New-ClaudeConfig

    # ========== 安装完成 ==========
    Show-CompletionMessage
}

# ============================================================================
# 脚本入口
# 开始执行安装，出了问题别找老子（开玩笑的，有问题提Issue）
# ============================================================================
Start-Installation
