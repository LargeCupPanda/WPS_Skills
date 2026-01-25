# WPS Skills 一键安装脚本
# 用法: powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
# 作者: 老王

param(
    [switch]$SkipNodeCheck,
    [switch]$SkipWpsAddon
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WPS Skills 一键安装脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ==================== 1. 检测 Node.js ====================
if (-not $SkipNodeCheck) {
    Write-Host "[1/5] 检测 Node.js..." -ForegroundColor Yellow
    try {
        $nodeVersion = node -v 2>$null
        if ($nodeVersion -match "v(\d+)\.") {
            $majorVersion = [int]$Matches[1]
            if ($majorVersion -ge 18) {
                Write-Host "  Node.js $nodeVersion" -ForegroundColor Green
            } else {
                Write-Host "  Node.js 版本过低: $nodeVersion (需要 >= 18)" -ForegroundColor Red
                Write-Host "  请访问 https://nodejs.org/ 下载最新版本" -ForegroundColor Yellow
                exit 1
            }
        }
    } catch {
        Write-Host "  未检测到 Node.js" -ForegroundColor Red
        Write-Host "  请访问 https://nodejs.org/ 下载安装" -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "[1/5] 跳过 Node.js 检测" -ForegroundColor Gray
}

# ==================== 2. 安装依赖 ====================
Write-Host "[2/5] 安装依赖..." -ForegroundColor Yellow
Set-Location "$ProjectRoot\wps-office-mcp"
npm install --silent 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "  npm install 失败" -ForegroundColor Red
    exit 1
}
Write-Host "  依赖安装完成" -ForegroundColor Green

# ==================== 3. 编译 TypeScript ====================
Write-Host "[3/5] 编译 TypeScript..." -ForegroundColor Yellow
npm run build --silent 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "  编译失败" -ForegroundColor Red
    exit 1
}
Write-Host "  编译完成" -ForegroundColor Green

# ==================== 4. 配置 Claude Code ====================
Write-Host "[4/5] 配置 Claude Code..." -ForegroundColor Yellow
$claudeSettingsPath = "$env:USERPROFILE\.claude\settings.json"
$mcpServerPath = "$ProjectRoot\wps-office-mcp\dist\index.js" -replace "\\", "\\"

# 确保 .claude 目录存在
if (-not (Test-Path "$env:USERPROFILE\.claude")) {
    New-Item -ItemType Directory -Path "$env:USERPROFILE\.claude" -Force | Out-Null
}

# 读取或创建 settings.json
if (Test-Path $claudeSettingsPath) {
    $settings = Get-Content $claudeSettingsPath -Raw | ConvertFrom-Json
} else {
    $settings = @{}
}

# 确保 mcpServers 存在
if (-not $settings.mcpServers) {
    $settings | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue @{} -Force
}

# 添加 wps-office 配置
$settings.mcpServers | Add-Member -NotePropertyName "wps-office" -NotePropertyValue @{
    command = "node"
    args = @($mcpServerPath)
} -Force

# 保存
$settings | ConvertTo-Json -Depth 10 | Set-Content $claudeSettingsPath -Encoding UTF8
Write-Host "  Claude Code 配置完成: $claudeSettingsPath" -ForegroundColor Green

# ==================== 5. 安装 WPS 加载项 ====================
if (-not $SkipWpsAddon) {
    Write-Host "[5/5] 安装 WPS 加载项..." -ForegroundColor Yellow
    $wpsAddonDir = "$env:APPDATA\kingsoft\wps\jsaddons"
    $targetDir = "$wpsAddonDir\wps-claude-addon_"
    $sourceDir = "$ProjectRoot\wps-claude-addon"
    $publishXml = "$wpsAddonDir\publish.xml"

    # 检查 WPS 目录是否存在
    if (-not (Test-Path $wpsAddonDir)) {
        Write-Host "  WPS 加载项目录不存在，请先安装 WPS Office" -ForegroundColor Red
        Write-Host "  下载地址: https://www.wps.cn/" -ForegroundColor Yellow
        exit 1
    }

    # 拷贝加载项文件
    if (Test-Path $targetDir) {
        Remove-Item $targetDir -Recurse -Force
    }
    Copy-Item $sourceDir $targetDir -Recurse
    Write-Host "  加载项文件已拷贝" -ForegroundColor Green

    # 更新 publish.xml
    $addonEntry = '<jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_/"/>'
    if (Test-Path $publishXml) {
        $content = Get-Content $publishXml -Raw
        if ($content -notmatch "wps-claude-addon") {
            # 在 </jsplugins> 前插入
            $content = $content -replace "</jsplugins>", "    $addonEntry`n</jsplugins>"
            Set-Content $publishXml $content -Encoding UTF8
            Write-Host "  publish.xml 已更新" -ForegroundColor Green
        } else {
            Write-Host "  publish.xml 已包含配置" -ForegroundColor Green
        }
    } else {
        # 创建新的 publish.xml
        $newPublish = @"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
    $addonEntry
</jsplugins>
"@
        Set-Content $publishXml $newPublish -Encoding UTF8
        Write-Host "  publish.xml 已创建" -ForegroundColor Green
    }
} else {
    Write-Host "[5/5] 跳过 WPS 加载项安装" -ForegroundColor Gray
}

# ==================== 完成 ====================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  安装完成!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "下一步:" -ForegroundColor Yellow
Write-Host "  1. 重启 Claude Code" -ForegroundColor White
Write-Host "  2. 重启 WPS Office" -ForegroundColor White
Write-Host "  3. 在 WPS 中查看 'Claude助手' 选项卡" -ForegroundColor White
Write-Host ""
