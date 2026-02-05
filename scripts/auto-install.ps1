# Input: 用户环境与安装参数
# Output: 安装与配置结果
# Pos: Windows 一键安装脚本。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
# WPS Skills One-Click Installer (Windows)
# Usage: powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
# Author: Laowang

[CmdletBinding()]
param(
    [switch]$SkipNodeCheck,
    [switch]$SkipWpsAddon,
    [switch]$SkipMcpCheck
)

$ErrorActionPreference = "Continue"

$ScriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ProjectRoot = $PSScriptRoot
} else {
    $ProjectRoot = Split-Path -Parent (Split-Path -Parent $ScriptPath)
}

$ClaudeSettingsPath = "$env:USERPROFILE\.claude\settings.json"
$ClaudeSkillsDir = "$env:USERPROFILE\.claude\skills"
$WpsAddonDir = "$env:APPDATA\kingsoft\wps\jsaddons"
$TotalSteps = 5

function Write-Step {
    param([string]$Message, [int]$Step)
    Write-Host "[$Step/$TotalSteps] $Message..." -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-Error-Msg {
    param([string]$Message)
    Write-Host "  [ERR] $Message" -ForegroundColor Red
}

function Write-Warning-Msg {
    param([string]$Message)
    Write-Host "  [WARN] $Message" -ForegroundColor DarkYellow
}

function Write-Separator {
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

Write-Separator
Write-Host "  WPS Skills One-Click Installer (Windows)" -ForegroundColor Cyan
Write-Separator
Write-Host ""

if (-not $SkipNodeCheck) {
    Write-Step "Checking Node.js" 1
    try {
        $nodeVersionOutput = node -v 2>&1
        if ($LASTEXITCODE -eq 0 -and $nodeVersionOutput -match "v(\d+)\.") {
            $majorVersion = [int]$Matches[1]
            if ($majorVersion -ge 18) {
                Write-Success "Node.js $nodeVersionOutput"
            } else {
                Write-Error-Msg "Node.js version too low: $nodeVersionOutput (need >= 18)"
                Write-Host "  Download: https://nodejs.org/" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Tip: Use -SkipNodeCheck to skip this check" -ForegroundColor Gray
                exit 1
            }
        } else {
            throw "Node.js not installed"
        }
    } catch {
        Write-Error-Msg "Node.js not detected"
        Write-Host "  Download: https://nodejs.org/" -ForegroundColor Yellow
        Write-Host ""
        exit 1
    }
} else {
    Write-Step "Skipping Node.js check" 1
    Write-Warning-Msg "Skipped Node.js version check"
}

Write-Step "Installing npm dependencies" 2

$McpDir = "$ProjectRoot\wps-office-mcp"
if (-not (Test-Path $McpDir)) {
    Write-Error-Msg "MCP directory not found: $McpDir"
    exit 1
}

Set-Location $McpDir

Write-Host "  Running npm install..." -ForegroundColor Gray
npm install 2>&1 | Out-Null

if ($LASTEXITCODE -eq 0) {
    Write-Success "Dependencies installed"
} else {
    Write-Error-Msg "npm install failed"
    Write-Host "  Try manually: cd wps-office-mcp && npm install" -ForegroundColor Gray
    exit 1
}

Write-Step "Compiling TypeScript" 3

Write-Host "  Running npm run build..." -ForegroundColor Gray
npm run build 2>&1 | Out-Null

if ($LASTEXITCODE -eq 0) {
    Write-Success "Compilation complete"
} else {
    Write-Error-Msg "Compilation failed"
    Write-Host "  Try manually: cd wps-office-mcp && npm run build" -ForegroundColor Gray
    exit 1
}

Write-Step "Configuring Claude Code" 4

if (-not (Test-Path "$env:USERPROFILE\.claude")) {
    New-Item -ItemType Directory -Path "$env:USERPROFILE\.claude" -Force | Out-Null
    Write-Success "Created .claude directory"
}

$McpServerPath = "$ProjectRoot\wps-office-mcp\dist\index.js" -replace "\\", "\\"

if (Test-Path $ClaudeSettingsPath) {
    try {
        $settings = Get-Content $ClaudeSettingsPath -Raw | ConvertFrom-Json
    } catch {
        Write-Warning-Msg "Failed to read settings.json, creating new file"
        $settings = @{}
    }
} else {
    $settings = @{}
}

if (-not $settings.mcpServers) {
    $settings | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue @{} -Force
}

$McpConfig = @{
    command = "node"
    args    = @($McpServerPath)
}
$settings.mcpServers | Add-Member -NotePropertyName "wps-office" -NotePropertyValue $McpConfig -Force

try {
    $settings | ConvertTo-Json -Depth 10 | Set-Content $ClaudeSettingsPath -Encoding UTF8
    Write-Success "MCP Server configured"
} catch {
    Write-Error-Msg "Failed to save settings.json: $_"
}

if (-not (Test-Path $ClaudeSkillsDir)) {
    New-Item -ItemType Directory -Path $ClaudeSkillsDir -Force | Out-Null
}

$SkillsSourceDir = "$ProjectRoot\skills"
$SkillsList = @("wps-excel", "wps-word", "wps-ppt", "wps-office")

foreach ($skill in $SkillsList) {
    $SourcePath = "$SkillsSourceDir\$skill"
    $DestPath = "$ClaudeSkillsDir\$skill"

    if (Test-Path $DestPath) {
        Remove-Item $DestPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $SourcePath) {
        Copy-Item -Path $SourcePath -Destination $DestPath -Recurse -Force
    }
}

Write-Success "Skills installed ($($SkillsList.Count) skills)"

if (-not $SkipWpsAddon) {
    Write-Step "Installing WPS Add-on" 5

    if (-not (Test-Path $WpsAddonDir)) {
        Write-Error-Msg "WPS add-on directory not found: $WpsAddonDir"
        Write-Host "  Please install WPS Office first: https://www.wps.cn/" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Tip: Use -SkipWpsAddon to skip this step" -ForegroundColor Gray
        exit 1
    }

    $SourceAddonDir = "$ProjectRoot\wps-claude-addon"
    $TargetAddonDir = "$WpsAddonDir\wps-claude-addon_"

    if (-not (Test-Path $SourceAddonDir)) {
        Write-Error-Msg "Add-on source directory not found: $SourceAddonDir"
        exit 1
    }

    if (Test-Path $TargetAddonDir) {
        Remove-Item $TargetAddonDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    Copy-Item -Path $SourceAddonDir -Destination $TargetAddonDir -Recurse -Force
    Write-Success "Add-on files copied"

    $PublishXmlPath = "$WpsAddonDir\publish.xml"
    $AddonEntry = '<jsplugin type="wps,et,wpp" enable="enable_dev" name="wps-claude-addon" url="wps-claude-addon_"/>'

    if (Test-Path $PublishXmlPath) {
        $content = Get-Content $PublishXmlPath -Raw -Encoding UTF8

        if ($content -notmatch "wps-claude-addon") {
            $content = $content -replace "</jsplugins>", "    $AddonEntry`n</jsplugins>"
            Set-Content $PublishXmlPath $content -Encoding UTF8 -NoNewline
            Write-Success "publish.xml updated"
        } else {
            Write-Success "publish.xml already contains configuration (skipped)"
        }
    } else {
        $NewPublishXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
    $AddonEntry
</jsplugins>
"@
        Set-Content $PublishXmlPath $NewPublishXml -Encoding UTF8 -NoNewline
        Write-Success "publish.xml created"
    }
} else {
    Write-Step "Skipping WPS Add-on installation" 5
    Write-Warning-Msg "Skipped WPS add-on installation"
}

Write-Host ""
Write-Separator
Write-Host "  Installation Verification" -ForegroundColor Cyan
Write-Separator

$VerifyPassed = $true

Write-Host "`n[MCP Server]" -ForegroundColor Cyan
if (Test-Path $ClaudeSettingsPath) {
    try {
        $settings = Get-Content $ClaudeSettingsPath -Raw | ConvertFrom-Json
        if ($settings.mcpServers.'wps-office') {
            Write-Success "settings.json configured with wps-office"
            Write-Host "    Path: $($settings.mcpServers.'wps-office'.args[0])" -ForegroundColor Gray
        } else {
            Write-Error-Msg "wps-office not found in settings.json"
            $VerifyPassed = $false
        }
    } catch {
        Write-Error-Msg "Cannot read settings.json"
        $VerifyPassed = $false
    }
} else {
    Write-Error-Msg "settings.json not found"
    $VerifyPassed = $false
}

Write-Host "`n[Skills]" -ForegroundColor Cyan
$InstalledSkills = @()
foreach ($skill in $SkillsList) {
    if (Test-Path "$ClaudeSkillsDir\$skill") {
        $InstalledSkills += $skill
    }
}

if ($InstalledSkills.Count -eq $SkillsList.Count) {
    Write-Success "All skills installed ($($InstalledSkills.Count) skills)"
    Write-Host "    $($InstalledSkills -join ', ')" -ForegroundColor Gray
} else {
    Write-Warning-Msg "Some skills missing: $($InstalledSkills.Count)/$($SkillsList.Count)"
    $VerifyPassed = $false
}

if (-not $SkipWpsAddon) {
    Write-Host "`n[WPS Add-on]" -ForegroundColor Cyan
    if (Test-Path "$WpsAddonDir\wps-claude-addon_") {
        Write-Success "Add-on files in place"
    } else {
        Write-Error-Msg "Add-on files not found"
        $VerifyPassed = $false
    }

    if (Test-Path "$WpsAddonDir\publish.xml") {
        $content = Get-Content "$WpsAddonDir\publish.xml" -Raw -Encoding UTF8
        if ($content -match "wps-claude-addon") {
            Write-Success "publish.xml configured"
        } else {
            Write-Error-Msg "wps-claude-addon not found in publish.xml"
            $VerifyPassed = $false
        }
    }
}

Write-Host ""
Write-Separator

if ($VerifyPassed) {
    Write-Host "  Installation Complete!" -ForegroundColor Green
    Write-Separator
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart Claude Code (Required!)" -ForegroundColor White
    Write-Host "  2. Restart WPS Office" -ForegroundColor White
    Write-Host "  3. Look for 'Claude Assistant' tab in WPS" -ForegroundColor White
    Write-Host ""
    Write-Host "Available Skills:" -ForegroundColor Yellow
    Write-Host "  /wps-excel  - Excel Assistant" -ForegroundColor White
    Write-Host "  /wps-word   - Word Assistant" -ForegroundColor White
    Write-Host "  /wps-ppt    - PowerPoint Assistant" -ForegroundColor White
    Write-Host "  /wps-office - General Assistant" -ForegroundColor White
    Write-Host ""
} else {
    Write-Host "  Installation Complete with Warnings!" -ForegroundColor Yellow
    Write-Separator
    Write-Host ""
    Write-Host "Please check the errors above and fix manually." -ForegroundColor Red
    Write-Host ""
}
