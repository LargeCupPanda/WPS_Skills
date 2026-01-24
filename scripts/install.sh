#!/bin/bash
# ============================================================================
# WPS-Claude-Skills 安装脚本 (macOS/Linux)
# ============================================================================
# 作者：老王 (隔壁老王团队 DevOps - 张大牛)
#
# 艹，这个SB脚本负责把整个项目环境搭建起来
# 包括检查Node.js、npm、编译TypeScript、复制加载项到WPS目录
#
# 用法：chmod +x scripts/install.sh && ./scripts/install.sh
#
# 参数：
#   --skip-node-check  跳过Node.js版本检查（你确定你知道自己在干嘛？）
#   --skip-wps-check   跳过WPS安装检查（没WPS你装个锤子）
#   --dev              开发模式（创建符号链接而不是复制文件）
# ============================================================================

set -e  # 遇到错误立即退出，老王不喜欢半吊子的安装

# ============================================================================
# 颜色定义
# 让控制台输出好看点，虽然老王觉得花里胡哨，但用户体验还是要有的
# ============================================================================
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
WHITE='\033[1;37m'
NC='\033[0m' # No Color - 重置颜色

# ============================================================================
# 输出函数
# 统一的输出格式，看起来专业一点
# ============================================================================
success() { echo -e "${GREEN}✓ $1${NC}"; }
warn() { echo -e "${YELLOW}⚠ $1${NC}"; }
error() { echo -e "${RED}✗ $1${NC}"; }
info() { echo -e "${CYAN}>>> $1${NC}"; }

# ============================================================================
# 参数解析
# 解析命令行参数，不会用的去看帮助
# ============================================================================
SKIP_NODE_CHECK=false
SKIP_WPS_CHECK=false
DEV_MODE=false

while [[ $# -gt 0 ]]; do
    case $1 in
        --skip-node-check)
            SKIP_NODE_CHECK=true
            shift
            ;;
        --skip-wps-check)
            SKIP_WPS_CHECK=true
            shift
            ;;
        --dev)
            DEV_MODE=true
            shift
            ;;
        -h|--help)
            echo "用法: $0 [选项]"
            echo ""
            echo "选项:"
            echo "  --skip-node-check  跳过Node.js版本检查"
            echo "  --skip-wps-check   跳过WPS安装检查"
            echo "  --dev              开发模式（创建符号链接）"
            echo "  -h, --help         显示帮助信息"
            exit 0
            ;;
        *)
            warn "未知参数: $1"
            shift
            ;;
    esac
done

# ============================================================================
# 打印老王风格的 Banner
# 虽然花里胡哨，但用户体验还是要有的嘛
# ============================================================================
show_banner() {
    echo ""
    echo -e "${CYAN}╔═══════════════════════════════════════════════════════════╗${NC}"
    echo -e "${CYAN}║                                                           ║${NC}"
    echo -e "${CYAN}║         WPS-Claude-Skills 安装程序 v1.0                   ║${NC}"
    echo -e "${CYAN}║                                                           ║${NC}"
    echo -e "${CYAN}║         让 WPS Office 说人话的智能助手                    ║${NC}"
    echo -e "${CYAN}║         (老王出品，必属精品，不接受反驳)                  ║${NC}"
    echo -e "${CYAN}║                                                           ║${NC}"
    echo -e "${CYAN}╚═══════════════════════════════════════════════════════════╝${NC}"
    echo ""
}

# ============================================================================
# 获取项目根目录
# 这个函数用来定位项目根目录，别TM瞎改路径
# ============================================================================
get_project_root() {
    local script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    echo "$(dirname "$script_dir")"
}

PROJECT_ROOT=$(get_project_root)

# ============================================================================
# 检测操作系统
# 不同系统的WPS路径不一样，这是个坑
# ============================================================================
detect_os() {
    case "$(uname -s)" in
        Darwin*)    echo "macos";;
        Linux*)     echo "linux";;
        CYGWIN*|MINGW*|MSYS*) echo "windows";;  # 如果你在WSL里跑，那就是你自己的问题
        *)          echo "unknown";;
    esac
}

OS=$(detect_os)
info "检测到操作系统: $OS"

# ============================================================================
# 检查 Node.js 版本
# 没有Node.js >= 18.0.0，你玩个锤子？
# ============================================================================
check_node_version() {
    info "检查 Node.js 版本..."

    if ! command -v node &> /dev/null; then
        error "Node.js 未安装"
        echo "  请安装 Node.js 18+ 版本:"
        case $OS in
            macos)
                echo "    brew install node@18"
                echo "    或者去 https://nodejs.org 下载安装"
                ;;
            linux)
                echo "    # Ubuntu/Debian:"
                echo "    curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -"
                echo "    sudo apt-get install -y nodejs"
                echo ""
                echo "    # 或者用 nvm:"
                echo "    curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.0/install.sh | bash"
                echo "    nvm install 18"
                ;;
            *)
                echo "    访问 https://nodejs.org 下载安装"
                ;;
        esac
        echo "  没有Node.js你让老子怎么帮你编译TypeScript？"
        return 1
    fi

    local node_version=$(node --version | tr -d 'v')
    local major_version=$(echo "$node_version" | cut -d. -f1)

    if [ "$major_version" -lt 18 ]; then
        error "艹！Node.js 版本过低: v$node_version"
        echo "  需要 Node.js >= 18.0.0，你这版本太老了，升级去！"
        return 1
    fi

    success "Node.js 版本: v$node_version (这个可以)"
    return 0
}

# ============================================================================
# 检查 npm 是否可用
# npm都没有，你装的什么野鸡Node.js？
# ============================================================================
check_npm_available() {
    info "检查 npm 是否可用..."

    if ! command -v npm &> /dev/null; then
        error "npm 未安装"
        echo "  你这Node.js装的有问题啊，npm都没有？"
        return 1
    fi

    local npm_version=$(npm --version)
    success "npm 版本: $npm_version"
    return 0
}

# ============================================================================
# 检查 WPS Office 安装 (macOS)
# macOS的WPS路径比较统一，还算好找
# ============================================================================
check_wps_macos() {
    info "检查 WPS Office 安装..."

    local wps_paths=(
        "/Applications/wpsoffice.app"
        "/Applications/WPS Office.app"
        "$HOME/Applications/wpsoffice.app"
        "$HOME/Applications/WPS Office.app"
    )

    for path in "${wps_paths[@]}"; do
        if [ -d "$path" ]; then
            success "WPS Office 已安装: $path"
            return 0
        fi
    done

    warn "未检测到 WPS Office 安装"
    echo "  请从 https://mac.wps.cn 下载安装 WPS Office"
    echo "  没有WPS你装这个加载项给谁用？"
    return 1
}

# ============================================================================
# 检查 WPS Office 安装 (Linux)
# Linux的WPS路径比较乱，得多检查几个地方
# ============================================================================
check_wps_linux() {
    info "检查 WPS Office 安装..."

    # 先检查命令是否存在
    if command -v wps &> /dev/null; then
        success "WPS Office 已安装 (通过命令行检测)"
        return 0
    fi

    # 检查常见安装路径
    local wps_paths=(
        "/opt/kingsoft/wps-office"
        "/usr/share/wps-office"
        "/usr/lib/office6"
    )

    for path in "${wps_paths[@]}"; do
        if [ -d "$path" ]; then
            success "WPS Office 已安装: $path"
            return 0
        fi
    done

    warn "未检测到 WPS Office 安装"
    echo "  请从 https://linux.wps.cn 下载安装 WPS Office"
    echo "  没有WPS你装这个加载项给谁用？"
    return 1
}

# ============================================================================
# 检查 WPS 安装（统一入口）
# 根据操作系统调用对应的检查函数
# ============================================================================
check_wps_installation() {
    case $OS in
        macos)
            check_wps_macos
            ;;
        linux)
            check_wps_linux
            ;;
        *)
            warn "当前系统不支持 WPS 检查"
            return 1
            ;;
    esac
}

# ============================================================================
# 安装 MCP Server 依赖并编译 TypeScript
# 这一步是核心，编译不过老子把键盘都给你砸了
# ============================================================================
install_mcp_server_dependencies() {
    local mcp_path="$PROJECT_ROOT/wps-office-mcp"

    info "安装 MCP Server 依赖..."

    if [ ! -d "$mcp_path" ]; then
        error "MCP Server 目录不存在: $mcp_path"
        echo "  你TM把项目文件删了？"
        return 1
    fi

    cd "$mcp_path"

    # 第一步：npm install
    echo "  执行 npm install..."
    if ! npm install; then
        error "npm install 失败，这憨批依赖出问题了"
        return 1
    fi
    success "MCP Server 依赖安装完成"

    # 第二步：编译 TypeScript
    echo "  编译 TypeScript..."
    if ! npm run build; then
        error "TypeScript 编译失败"
        echo "  艹，编译都过不了，你写的什么SB代码？"
        return 1
    fi
    success "TypeScript 编译完成"

    cd "$PROJECT_ROOT"
    return 0
}

# ============================================================================
# 复制 WPS 加载项到 WPS 目录 (macOS)
# macOS路径: ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/Documents/wps/jsaddons/
# ============================================================================
copy_wps_addon_macos() {
    local addon_source="$PROJECT_ROOT/wps-claude-addon"

    info "复制 WPS 加载项到 WPS 目录 (macOS)..."

    if [ ! -d "$addon_source" ]; then
        error "WPS 加载项源目录不存在: $addon_source"
        echo "  你把加载项目录删了？脑子被门夹了？"
        return 1
    fi

    # WPS 加载项目标目录 (macOS)
    # 这个路径老王查了半天WPS文档才找到的
    local wps_addons_path="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/Documents/wps/jsaddons"
    local addon_target="$wps_addons_path/wps-claude-addon"

    # 创建目标目录（如果不存在）
    if [ ! -d "$wps_addons_path" ]; then
        echo "  创建 WPS 加载项目录: $wps_addons_path"
        mkdir -p "$wps_addons_path"
    fi

    # 如果目标已存在，先删除（避免残留文件搞事情）
    if [ -e "$addon_target" ]; then
        echo "  删除旧版本加载项..."
        rm -rf "$addon_target"
    fi

    if [ "$DEV_MODE" = true ]; then
        # 开发模式：创建符号链接（方便调试，改了源文件立即生效）
        echo "  开发模式：创建符号链接..."
        ln -s "$addon_source" "$addon_target"
        success "已创建符号链接: $addon_target -> $addon_source"
    else
        # 正常模式：复制整个目录
        echo "  复制加载项文件..."
        cp -R "$addon_source" "$addon_target"
        success "WPS 加载项已复制到: $addon_target"
    fi

    # 显示复制的文件数量，让用户安心
    local file_count=$(find "$addon_target" -type f | wc -l | tr -d ' ')
    echo "  共处理 $file_count 个文件"

    return 0
}

# ============================================================================
# 复制 WPS 加载项到 WPS 目录 (Linux)
# Linux路径: ~/.local/share/Kingsoft/wps/jsaddons/
# ============================================================================
copy_wps_addon_linux() {
    local addon_source="$PROJECT_ROOT/wps-claude-addon"

    info "复制 WPS 加载项到 WPS 目录 (Linux)..."

    if [ ! -d "$addon_source" ]; then
        error "WPS 加载项源目录不存在: $addon_source"
        echo "  你把加载项目录删了？脑子被门夹了？"
        return 1
    fi

    # WPS 加载项目标目录 (Linux)
    local wps_addons_path="$HOME/.local/share/Kingsoft/wps/jsaddons"
    local addon_target="$wps_addons_path/wps-claude-addon"

    # 创建目标目录（如果不存在）
    if [ ! -d "$wps_addons_path" ]; then
        echo "  创建 WPS 加载项目录: $wps_addons_path"
        mkdir -p "$wps_addons_path"
    fi

    # 如果目标已存在，先删除（避免残留文件搞事情）
    if [ -e "$addon_target" ]; then
        echo "  删除旧版本加载项..."
        rm -rf "$addon_target"
    fi

    if [ "$DEV_MODE" = true ]; then
        # 开发模式：创建符号链接（方便调试，改了源文件立即生效）
        echo "  开发模式：创建符号链接..."
        ln -s "$addon_source" "$addon_target"
        success "已创建符号链接: $addon_target -> $addon_source"
    else
        # 正常模式：复制整个目录
        echo "  复制加载项文件..."
        cp -R "$addon_source" "$addon_target"
        success "WPS 加载项已复制到: $addon_target"
    fi

    # 显示复制的文件数量，让用户安心
    local file_count=$(find "$addon_target" -type f | wc -l | tr -d ' ')
    echo "  共处理 $file_count 个文件"

    return 0
}

# ============================================================================
# 复制 WPS 加载项（统一入口）
# 根据操作系统调用对应的复制函数
# ============================================================================
copy_wps_addon() {
    case $OS in
        macos)
            copy_wps_addon_macos
            ;;
        linux)
            copy_wps_addon_linux
            ;;
        *)
            warn "当前系统不支持 WPS 加载项自动安装"
            echo "  请手动将 wps-claude-addon 目录复制到 WPS 加载项目录"
            return 1
            ;;
    esac
}

# ============================================================================
# 生成 Claude Code MCP 配置示例
# 告诉用户怎么配置Claude Code，省得他们来烦老子
# ============================================================================
generate_claude_config() {
    local mcp_dist_path="$PROJECT_ROOT/wps-office-mcp/dist/index.js"

    info "生成 Claude Code 配置..."

    success "Claude Code MCP 配置示例:"
    echo ""
    echo -e "${YELLOW}{"
    echo '  "mcpServers": {'
    echo '    "wps-office": {'
    echo '      "command": "node",'
    echo "      \"args\": [\"$mcp_dist_path\"],"
    echo '      "env": {'
    echo '        "WPS_ADDON_PORT": "58080"'
    echo '      }'
    echo '    }'
    echo '  }'
    echo -e "}${NC}"
    echo ""
    echo "  请将上述配置添加到 ~/.claude/settings.json 或者项目根目录的 .mcp.json"
    echo "  不会配置？去看README.md，老王都写好了！"
}

# ============================================================================
# 显示安装完成信息和使用说明
# 告诉用户下一步该干嘛，省得他们抓瞎
# ============================================================================
show_completion_message() {
    echo ""
    echo -e "${GREEN}═══════════════════════════════════════════════════════════${NC}"
    success "   安装完成！老王出品，必属精品！"
    echo -e "${GREEN}═══════════════════════════════════════════════════════════${NC}"
    echo ""

    info "下一步操作："
    echo -e "  1. 重启 WPS Office（必须重启才能加载新的加载项）"
    echo -e "  2. 配置 Claude Code（参考上方配置示例）"
    echo -e "  3. 启动 MCP Server（开发时用）:"
    echo -e "       ${YELLOW}cd wps-office-mcp && npm start${NC}"
    echo -e "  4. 在 Claude Code 中使用:"
    echo -e "       ${YELLOW}/wps-excel 帮我写个求和公式${NC}"
    echo ""

    info "遇到问题？"
    echo "  - 检查WPS是否正确安装并重启"
    echo "  - 检查Node.js版本是否 >= 18.0.0"
    echo "  - 查看项目README.md获取更多帮助"
    echo "  - 实在不行？找老王，但老王可能会骂你（开玩笑的）"
    echo ""
}

# ============================================================================
# 主安装流程
# 把上面所有步骤串起来，一步一步执行
# ============================================================================
main() {
    show_banner

    local all_passed=true

    # ========== 第一步：检查 Node.js 和 npm ==========
    if [ "$SKIP_NODE_CHECK" = false ]; then
        if ! check_node_version; then
            all_passed=false
        fi
        if ! check_npm_available; then
            all_passed=false
        fi
    else
        warn ">>> 跳过 Node.js 检查（你最好知道自己在干嘛）"
    fi

    # ========== 第二步：检查 WPS Office ==========
    if [ "$SKIP_WPS_CHECK" = false ]; then
        if ! check_wps_installation; then
            # WPS 检查失败不阻止安装，但要警告
            warn "  没检测到WPS，但安装会继续"
            warn "  安装完记得装WPS，不然加载项没地方用！"
        fi
    else
        warn ">>> 跳过 WPS 检查"
    fi

    # ========== 如果基础检查失败，询问是否继续 ==========
    if [ "$all_passed" = false ]; then
        echo ""
        read -p "基础环境检查未通过，是否继续安装? (y/N) " -n 1 -r
        echo
        if [[ ! $REPLY =~ ^[Yy]$ ]]; then
            error "安装已取消，去把环境配好再来！"
            exit 1
        fi
        warn "好吧，你坚持继续，出问题别怪老子..."
    fi

    echo ""
    echo -e "${CYAN}═══════════════════════════════════════════════════════════${NC}"
    info "开始安装依赖和配置..."
    echo -e "${CYAN}═══════════════════════════════════════════════════════════${NC}"
    echo ""

    # ========== 第三步：安装 MCP Server 依赖并编译 ==========
    if ! install_mcp_server_dependencies; then
        error "MCP Server 依赖安装或编译失败"
        error "艹，这一步都过不了，后面也别装了！"
        exit 1
    fi

    # ========== 第四步：复制 WPS 加载项 ==========
    if ! copy_wps_addon; then
        error "WPS 加载项复制失败"
        error "复制文件都能失败，检查一下权限问题！"
        exit 1
    fi

    # ========== 第五步：生成 Claude Code 配置示例 ==========
    generate_claude_config

    # ========== 安装完成 ==========
    show_completion_message
}

# ============================================================================
# 脚本入口
# 开始执行安装，出了问题别找老子（开玩笑的，有问题提Issue）
# ============================================================================
main "$@"
