#!/bin/bash
# WPS Skills Mac 一键安装脚本 (反向轮询架构版)
# @author 老王
# @date 2026-01-27
#
# 架构说明：
# - MCP Server 启动 HTTP 服务 (端口 58891)
# - WPS 加载项作为客户端轮询获取命令
# - 根据命令类型自动切换 WPS 应用

set -e

echo "================================================"
echo "   WPS Skills Mac 一键安装脚本 (老王出品)"
echo "   版本: 2.0.0 (反向轮询架构)"
echo "================================================"
echo ""

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_DIR="$( cd "$SCRIPT_DIR/.." && pwd )"

# 加载项源目录 (wps-claude-assistant)
ADDON_SOURCE="$PROJECT_DIR/wps-claude-assistant"

# WPS 加载项目录 (注意：目录名必须以_结尾)
WPS_ADDON_DIR="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons"
ADDON_TARGET="$WPS_ADDON_DIR/claude-assistant_"

echo "项目目录: $PROJECT_DIR"
echo ""

# ========== 前置条件检查 ==========
check_prerequisites() {
    echo "========== [1/5] 前置条件检查 =========="

    # 检查操作系统
    if [[ "$(uname)" != "Darwin" ]]; then
        echo -e "${RED}❌ 错误: 此脚本仅支持 macOS${NC}"
        exit 1
    fi
    echo -e "${GREEN}✓ 操作系统: macOS $(sw_vers -productVersion)${NC}"

    # 检查 WPS Office
    if [ ! -d "/Applications/wpsoffice.app" ]; then
        echo -e "${RED}❌ 错误: 未检测到 WPS Office${NC}"
        echo "请先安装 WPS Office for Mac: https://www.wps.cn/product/wpsmac"
        exit 1
    fi
    echo -e "${GREEN}✓ WPS Office 已安装${NC}"

    # 检查 Node.js
    if ! command -v node &> /dev/null; then
        echo -e "${RED}❌ 错误: 未找到 Node.js${NC}"
        echo "请先安装 Node.js: https://nodejs.org/"
        exit 1
    fi
    NODE_VERSION=$(node -v | sed 's/v//' | cut -d. -f1)
    if [[ $NODE_VERSION -lt 18 ]]; then
        echo -e "${RED}❌ 错误: Node.js 版本过低 (需要 >= 18.0.0)${NC}"
        exit 1
    fi
    echo -e "${GREEN}✓ Node.js $(node -v)${NC}"

    echo ""
}

# ========== 检查加载项目录 ==========
check_addon_dir() {
    echo "========== [2/5] 检查加载项目录 =========="

    if [ -d "$WPS_ADDON_DIR" ]; then
        echo -e "${GREEN}✓ 加载项目录存在${NC}"
    else
        echo "创建加载项目录..."
        mkdir -p "$WPS_ADDON_DIR"
        echo -e "${GREEN}✓ 加载项目录已创建${NC}"
    fi
    echo ""
}

# ========== 安装加载项 ==========
install_addon() {
    echo "========== [3/5] 安装 WPS Claude助手 加载项 =========="

    # 检查源目录
    if [ ! -d "$ADDON_SOURCE" ]; then
        echo -e "${RED}❌ 找不到加载项源文件: $ADDON_SOURCE${NC}"
        exit 1
    fi

    # 删除旧版本
    if [ -d "$ADDON_TARGET" ]; then
        echo "删除旧版本..."
        rm -rf "$ADDON_TARGET"
    fi

    # 复制新版本
    cp -r "$ADDON_SOURCE" "$ADDON_TARGET"

    # 确保脚本可执行
    chmod +x "$ADDON_TARGET/wps-auto.sh" 2>/dev/null || true

    echo -e "${GREEN}✓ 加载项已安装到: $ADDON_TARGET${NC}"
    echo ""
}

# ========== 更新 publish.xml ==========
update_publish_xml() {
    echo "========== [4/5] 更新加载项配置 =========="

    PUBLISH_XML="$WPS_ADDON_DIR/publish.xml"

    # 创建或更新 publish.xml
    cat > "$PUBLISH_XML" << 'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<jsplugins>
  <jsplugin name="claude-assistant" type="wps,et,wpp" url="claude-assistant_/" enable="enable_dev"/>
</jsplugins>
EOF

    echo -e "${GREEN}✓ publish.xml 配置完成${NC}"
    echo ""
}

# ========== 构建 MCP Server ==========
build_mcp_server() {
    echo "========== [5/5] 构建 MCP Server =========="

    MCP_DIR="$PROJECT_DIR/wps-office-mcp"

    if [ -d "$MCP_DIR" ]; then
        cd "$MCP_DIR"

        # 检查是否有 node_modules
        if [ ! -d "node_modules" ]; then
            echo "安装依赖..."
            npm install
        fi

        echo "构建项目..."
        npm run build

        if [ -f "dist/index.js" ]; then
            echo -e "${GREEN}✓ MCP Server 构建完成${NC}"
        else
            echo -e "${RED}❌ MCP Server 构建失败${NC}"
            exit 1
        fi
    else
        echo -e "${RED}❌ 未找到 MCP Server 目录: $MCP_DIR${NC}"
        exit 1
    fi
    echo ""
}

# ========== 配置 Claude Code ==========
configure_claude_code() {
    echo "========== [+] 配置 Claude Code MCP 设置 =========="

    MCP_INDEX_PATH="$PROJECT_DIR/wps-office-mcp/dist/index.js"

    # 检查 claude 命令是否存在
    if ! command -v claude &> /dev/null; then
        echo -e "${YELLOW}⚠ 未找到 claude 命令，跳过自动配置${NC}"
        echo ""
        echo "请手动配置 MCP Server："
        echo "  claude mcp add wps-office node $MCP_INDEX_PATH"
        echo ""
        return
    fi

    # 检查是否已配置 wps-office MCP
    if claude mcp list 2>/dev/null | grep -q "wps-office"; then
        echo -e "${YELLOW}⚠ wps-office MCP 已配置，跳过${NC}"
    else
        echo "注册 WPS Office MCP Server..."
        # 使用 claude mcp add 命令注册（正确方式！）
        # 注意：直接编辑 settings.json 是无效的，必须用这个命令
        claude mcp add wps-office node "$MCP_INDEX_PATH" --scope user 2>/dev/null || \
        claude mcp add wps-office node "$MCP_INDEX_PATH" 2>/dev/null

        if [ $? -eq 0 ]; then
            echo -e "${GREEN}✓ Claude Code MCP 配置完成${NC}"
        else
            echo -e "${YELLOW}⚠ 自动配置失败，请手动运行：${NC}"
            echo "  claude mcp add wps-office node $MCP_INDEX_PATH"
        fi
    fi
    echo ""
}

# ========== 注册 Skills (全局) ==========
install_skills() {
    echo "========== [+] 注册 WPS Skills (全局) =========="

    SKILLS_DIR="$HOME/.claude/skills"

    # 创建全局skills目录
    mkdir -p "$SKILLS_DIR"

    # 创建软链接（-sf 强制覆盖已存在的链接）
    ln -sf "$PROJECT_DIR/skills/wps-excel" "$SKILLS_DIR/wps-excel"
    ln -sf "$PROJECT_DIR/skills/wps-word" "$SKILLS_DIR/wps-word"
    ln -sf "$PROJECT_DIR/skills/wps-ppt" "$SKILLS_DIR/wps-ppt"
    ln -sf "$PROJECT_DIR/skills/wps-office" "$SKILLS_DIR/wps-office"

    echo -e "${GREEN}✓ Skills 已注册到: $SKILLS_DIR${NC}"
    echo "  - wps-excel"
    echo "  - wps-word"
    echo "  - wps-ppt"
    echo "  - wps-office"
    echo ""
}

# ========== 显示完成信息 ==========
show_complete() {
    echo "================================================"
    echo -e "${GREEN}   ✅ 安装完成！${NC}"
    echo "================================================"
    echo ""
    echo "已自动完成："
    echo "  ✓ WPS 加载项安装"
    echo "  ✓ MCP Server 构建和注册"
    echo "  ✓ Skills 全局注册"
    echo ""
    echo "下一步操作:"
    echo "  1. 重启 Claude Code"
    echo "  2. 重启 WPS Office"
    echo "  3. 打开 Excel/Word/PPT 文档"
    echo "  4. 查看 'Claude助手' 选项卡，确认状态为 '轮询中'"
    echo ""
    echo "技术架构 (MCP + Skills 双层架构):"
    echo "  ┌─────────────┐      ┌─────────────┐      ┌─────────────┐"
    echo "  │   Skills    │ ──── │ MCP Server  │ ←─── │ WPS 加载项  │"
    echo "  │ (指令层)    │      │ (HTTP:58891)│      │ (轮询客户端)│"
    echo "  └─────────────┘      └─────────────┘      └─────────────┘"
    echo ""
    echo "验证安装:"
    echo "  ls $ADDON_TARGET"
    echo "  claude mcp list"
    echo "  ls ~/.claude/skills/"
    echo ""
    echo "================================================"
}

# ========== 主流程 ==========
main() {
    check_prerequisites
    check_addon_dir
    install_addon
    update_publish_xml
    build_mcp_server
    configure_claude_code
    install_skills
    show_complete
}

main "$@"
