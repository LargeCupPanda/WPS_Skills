# WPS-Claude-Skills 开发环境启动脚本 (Windows PowerShell)
# 作者：老王
# 说明：这个脚本会同时启动 MCP Server 和 WPS 加载项开发服务
# 用法：.\scripts\dev.ps1

#Requires -Version 5.1

param(
    [switch]$McpOnly,      # 仅启动 MCP Server
    [switch]$AddonOnly,    # 仅启动 WPS 加载项
    [switch]$Watch,        # 启用文件监听模式
    [int]$McpPort = 0,     # MCP Server 端口 (0 = 使用 stdio)
    [int]$AddonPort = 58080 # WPS 加载项 HTTP 端口
)

# 颜色输出
function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Write-Success { Write-ColorOutput $args[0] "Green" }
function Write-Warning { Write-ColorOutput $args[0] "Yellow" }
function Write-Error { Write-ColorOutput $args[0] "Red" }
function Write-Info { Write-ColorOutput $args[0] "Cyan" }
function Write-Mcp { Write-ColorOutput "[MCP] $($args[0])" "Magenta" }
function Write-Addon { Write-ColorOutput "[ADDON] $($args[0])" "Blue" }

# 打印 Banner
function Show-Banner {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         WPS-Claude-Skills 开发环境                        ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# 获取项目根目录
function Get-ProjectRoot {
    $scriptPath = $PSScriptRoot
    return (Get-Item $scriptPath).Parent.FullName
}

$ProjectRoot = Get-ProjectRoot

# 检查进程占用端口
function Test-PortInUse {
    param([int]$Port)

    $connection = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
    return $null -ne $connection
}

# 停止已有进程
function Stop-ExistingProcesses {
    Write-Info "检查已有进程..."

    # 检查 MCP Server
    $nodeProcesses = Get-Process -Name "node" -ErrorAction SilentlyContinue
    foreach ($proc in $nodeProcesses) {
        try {
            $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.Id)").CommandLine
            if ($cmdLine -match "wps-office-mcp") {
                Write-Warning "发现已运行的 MCP Server (PID: $($proc.Id))，正在停止..."
                Stop-Process -Id $proc.Id -Force
            }
        }
        catch {}
    }

    # 检查端口占用
    if (Test-PortInUse -Port $AddonPort) {
        Write-Warning "端口 $AddonPort 已被占用"
        Write-Info "尝试释放端口..."

        $connection = Get-NetTCPConnection -LocalPort $AddonPort -ErrorAction SilentlyContinue
        if ($connection) {
            $proc = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
            if ($proc) {
                Write-Warning "终止占用端口的进程: $($proc.Name) (PID: $($proc.Id))"
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
            }
        }
    }
}

# 启动 MCP Server
function Start-McpServer {
    $mcpPath = Join-Path $ProjectRoot "wps-office-mcp"

    if (-not (Test-Path $mcpPath)) {
        Write-Error "MCP Server 目录不存在: $mcpPath"
        return $null
    }

    $packageJson = Join-Path $mcpPath "package.json"
    if (-not (Test-Path $packageJson)) {
        Write-Error "MCP Server package.json 不存在"
        return $null
    }

    Write-Mcp "启动 MCP Server..."
    Write-Mcp "目录: $mcpPath"

    # 设置环境变量
    $env:WPS_ADDON_PORT = $AddonPort
    $env:NODE_ENV = "development"

    # 启动命令
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "cmd.exe"
    $startInfo.WorkingDirectory = $mcpPath

    if ($Watch) {
        $startInfo.Arguments = "/c npm run dev"
        Write-Mcp "模式: 监听模式 (npm run dev)"
    }
    else {
        $startInfo.Arguments = "/c npm start"
        Write-Mcp "模式: 普通模式 (npm start)"
    }

    $startInfo.UseShellExecute = $true
    $startInfo.CreateNoWindow = $false

    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
        Write-Success "MCP Server 已启动 (PID: $($process.Id))"
        return $process
    }
    catch {
        Write-Error "MCP Server 启动失败: $_"
        return $null
    }
}

# 启动 WPS 加载项开发服务器
function Start-WpsAddon {
    $addonPath = Join-Path $ProjectRoot "wps-claude-addon"

    if (-not (Test-Path $addonPath)) {
        Write-Warning "WPS 加载项目录不存在: $addonPath"
        Write-Warning "跳过加载项启动，请确保 WPS 中已加载加载项"
        return $null
    }

    Write-Addon "WPS 加载项配置..."
    Write-Addon "目录: $addonPath"
    Write-Addon "HTTP 端口: $AddonPort"

    # 检查是否有开发服务器脚本
    $packageJson = Join-Path $addonPath "package.json"
    if (Test-Path $packageJson) {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "cmd.exe"
        $startInfo.WorkingDirectory = $addonPath
        $startInfo.Arguments = "/c npm run dev"
        $startInfo.UseShellExecute = $true
        $startInfo.CreateNoWindow = $false

        try {
            $process = [System.Diagnostics.Process]::Start($startInfo)
            Write-Success "WPS 加载项开发服务器已启动 (PID: $($process.Id))"
            return $process
        }
        catch {
            Write-Warning "WPS 加载项开发服务器启动失败: $_"
        }
    }
    else {
        Write-Info "WPS 加载项无独立开发服务器"
        Write-Info "请在 WPS 中手动加载加载项"
    }

    return $null
}

# 显示帮助信息
function Show-Help {
    Write-Host ""
    Write-Info "使用帮助:"
    Write-Host ""
    Write-Host "  常用命令:" -ForegroundColor White
    Write-Host "    .\scripts\dev.ps1              # 启动完整开发环境"
    Write-Host "    .\scripts\dev.ps1 -McpOnly     # 仅启动 MCP Server"
    Write-Host "    .\scripts\dev.ps1 -Watch       # 启用文件监听模式"
    Write-Host ""
    Write-Host "  测试连接:" -ForegroundColor White
    Write-Host "    curl http://localhost:$AddonPort/ping"
    Write-Host ""
    Write-Host "  查看日志:" -ForegroundColor White
    Write-Host "    # MCP Server 日志在控制台输出"
    Write-Host "    # WPS 加载项日志: %APPDATA%\kingsoft\wps\jsaddins\logs\"
    Write-Host ""
    Write-Host "  停止服务:" -ForegroundColor White
    Write-Host "    Ctrl+C 或关闭控制台窗口"
    Write-Host ""
}

# 主函数
function Start-DevEnvironment {
    Show-Banner

    # 停止已有进程
    Stop-ExistingProcesses

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Info "启动开发环境..."
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    $processes = @()

    # 启动 MCP Server
    if (-not $AddonOnly) {
        $mcpProcess = Start-McpServer
        if ($mcpProcess) {
            $processes += $mcpProcess
        }
    }

    # 启动 WPS 加载项
    if (-not $McpOnly) {
        Start-Sleep -Seconds 1  # 等待 MCP Server 启动
        $addonProcess = Start-WpsAddon
        if ($addonProcess) {
            $processes += $addonProcess
        }
    }

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Success "开发环境已启动!"
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan

    Show-Help

    Write-Host ""
    Write-Warning "按 Ctrl+C 停止所有服务..."
    Write-Host ""

    # 保持脚本运行，监控子进程
    try {
        while ($true) {
            Start-Sleep -Seconds 5

            # 检查进程状态
            foreach ($proc in $processes) {
                if ($proc -and $proc.HasExited) {
                    Write-Warning "进程 $($proc.Id) 已退出 (退出代码: $($proc.ExitCode))"
                }
            }
        }
    }
    catch {
        Write-Info "收到停止信号..."
    }
    finally {
        Write-Info "正在停止所有服务..."
        foreach ($proc in $processes) {
            if ($proc -and -not $proc.HasExited) {
                Write-Info "停止进程 $($proc.Id)..."
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
        Write-Success "开发环境已停止"
    }
}

# 执行
Start-DevEnvironment
