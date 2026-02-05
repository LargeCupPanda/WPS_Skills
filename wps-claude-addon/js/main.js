/**
 * Input: WPS 加载项回调与用户交互请求
 * Output: 状态展示与日志输出
 * Pos: Windows WPS 加载项脚本入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS Claude 智能助手 - COM桥接版
 * MCP Server通过PowerShell COM接口直接操作WPS
 * @author 老王
 */

// ==================== 全局状态 ====================
var addonStatus = {
    initialized: false
};

// ==================== ribbon.xml回调函数 ====================

function OnAddinLoad(ribbonUI) {
    console.log('[Claude助手] 加载项已加载 - COM桥接模式');
    addonStatus.initialized = true;
}

function ribbonCheckStatus(control) {
    var statusText = '=== Claude助手状态 ===\n\n';
    statusText += '初始化状态: ' + (addonStatus.initialized ? '已完成' : '未完成') + '\n';
    statusText += '通信模式: PowerShell COM桥接\n';
    statusText += '说明: MCP Server直接通过COM接口操作WPS\n';

    try {
        if (typeof Application !== 'undefined') {
            statusText += '\n=== WPS信息 ===\n';
            statusText += '应用: ' + (Application.Name || 'WPS Office') + '\n';
            if (Application.ActiveWorkbook) {
                statusText += '当前文档: ' + Application.ActiveWorkbook.Name + ' (Excel)\n';
                statusText += '工作表数: ' + Application.ActiveWorkbook.Sheets.Count + '\n';
            } else if (Application.ActiveDocument) {
                statusText += '当前文档: ' + Application.ActiveDocument.Name + ' (Word)\n';
            } else if (Application.ActivePresentation) {
                statusText += '当前文档: ' + Application.ActivePresentation.Name + ' (PPT)\n';
            }
        }
    } catch (e) {
        statusText += '\nWPS信息获取失败: ' + e.message;
    }
    alert(statusText);
}

// ==================== WPS生命周期函数 ====================

function OnAddinNotify(params) {
    console.log('[Claude助手] 收到通知:', JSON.stringify(params));
}

// ==================== 初始化 ====================
console.log('[Claude助手] main.js 已加载 - COM桥接模式');
