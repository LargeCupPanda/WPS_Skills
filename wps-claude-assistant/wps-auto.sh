#!/bin/bash
# WPS自动化启动脚本 - Mac版
# 用于自动启动指定的WPS应用并等待加载项连接
# @author 老王

SERVER_URL="http://127.0.0.1:58891"
POLL_TIMEOUT=30
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 创建空白文件
create_blank_file() {
    local file_type=$1
    local file_path="/tmp/wps_auto_blank"

    case $file_type in
        "xlsx")
            python3 -c "
import zipfile
with zipfile.ZipFile('${file_path}.xlsx', 'w') as zf:
    zf.writestr('[Content_Types].xml', '<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/></Types>')
    zf.writestr('_rels/.rels', '<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>')
    zf.writestr('xl/workbook.xml', '<?xml version=\"1.0\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/></sheets></workbook>')
    zf.writestr('xl/_rels/workbook.xml.rels', '<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>')
    zf.writestr('xl/worksheets/sheet1.xml', '<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>')
" 2>/dev/null
            echo "${file_path}.xlsx"
            ;;
        "docx")
            python3 -c "
import zipfile
with zipfile.ZipFile('${file_path}.docx', 'w') as zf:
    zf.writestr('[Content_Types].xml', '<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>')
    zf.writestr('_rels/.rels', '<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>')
    zf.writestr('word/document.xml', '<?xml version=\"1.0\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t></w:t></w:r></w:p></w:body></w:document>')
" 2>/dev/null
            echo "${file_path}.docx"
            ;;
        "pptx")
            # 使用python-pptx创建正确格式的pptx
            python3 -c "
from pptx import Presentation
p = Presentation()
p.slides.add_slide(p.slide_layouts[6])  # 空白布局
p.save('${file_path}.pptx')
" 2>/dev/null
            echo "${file_path}.pptx"
            ;;
    esac
}

# 启动指定的WPS应用
start_app() {
    local app_type=$1
    local file_path=""

    case $app_type in
        "et"|"excel")
            echo "[WPS-Auto] 启动 Excel..."
            file_path=$(create_blank_file "xlsx")
            ;;
        "wps"|"word")
            echo "[WPS-Auto] 启动 Word..."
            file_path=$(create_blank_file "docx")
            ;;
        "wpp"|"ppt")
            echo "[WPS-Auto] 启动 PPT..."
            file_path=$(create_blank_file "pptx")
            ;;
        *)
            echo "[WPS-Auto] 未知应用类型: $app_type"
            return 1
            ;;
    esac

    open "$file_path" 2>/dev/null
    return 0
}

# 关闭所有WPS应用
close_all() {
    echo "[WPS-Auto] 关闭所有WPS应用..."
    pkill -f "wpsoffice" 2>/dev/null
    sleep 2
}

# 等待加载项连接
wait_connection() {
    local expected=$1
    local timeout=${2:-$POLL_TIMEOUT}
    local elapsed=0

    echo "[WPS-Auto] 等待 $expected 连接..."

    while [ $elapsed -lt $timeout ]; do
        result=$(curl -s -X POST "$SERVER_URL/send" -H "Content-Type: application/json" -d '{"action":"getAppInfo"}' 2>/dev/null)

        if [ $? -eq 0 ]; then
            sleep 2
            app_name=$(cat /tmp/server.log 2>/dev/null | grep "收到结果" | tail -1 | grep -o '"appName":"[^"]*"' | cut -d'"' -f4)

            case $expected in
                "et"|"excel")
                    if echo "$app_name" | grep -qi "Excel"; then
                        echo "[WPS-Auto] ✅ Excel 已连接!"
                        return 0
                    fi
                    ;;
                "wps"|"word")
                    if echo "$app_name" | grep -qi "Word"; then
                        echo "[WPS-Auto] ✅ Word 已连接!"
                        return 0
                    fi
                    ;;
                "wpp"|"ppt")
                    if echo "$app_name" | grep -qi "PowerPoint"; then
                        echo "[WPS-Auto] ✅ PPT 已连接!"
                        return 0
                    fi
                    ;;
            esac
        fi

        sleep 1
        elapsed=$((elapsed + 3))
    done

    echo "[WPS-Auto] ❌ 连接超时"
    return 1
}

# 切换应用（核心功能）
switch_to() {
    local target=$1

    close_all
    sleep 2
    start_app "$target"
    sleep 3
    wait_connection "$target"
}

# 获取当前连接的应用
get_current() {
    curl -s -X POST "$SERVER_URL/send" -H "Content-Type: application/json" -d '{"action":"getAppInfo"}' >/dev/null 2>&1
    sleep 2
    cat /tmp/server.log 2>/dev/null | grep "收到结果" | tail -1 | grep -o '"appName":"[^"]*"' | cut -d'"' -f4
}

# 主入口
case $1 in
    "start")
        start_app "$2"
        ;;
    "stop"|"close")
        close_all
        ;;
    "switch")
        switch_to "$2"
        ;;
    "wait")
        wait_connection "$2" "$3"
        ;;
    "current")
        echo "当前连接: $(get_current)"
        ;;
    *)
        echo "WPS自动化脚本 - Mac版 (老王出品)"
        echo ""
        echo "用法:"
        echo "  $0 start <app>    启动WPS应用"
        echo "  $0 stop           关闭所有WPS"
        echo "  $0 switch <app>   切换到指定应用"
        echo "  $0 wait <app>     等待应用连接"
        echo "  $0 current        查看当前连接"
        echo ""
        echo "应用: et/excel, wps/word, wpp/ppt"
        ;;
esac
