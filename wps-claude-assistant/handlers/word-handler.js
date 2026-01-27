/**
 * Word操作处理器 - Mac版WPS加载项
 * 使用WPS JavaScript API实现Word文档操作
 * @author 老李（参考老王的PowerShell实现）
 */

/**
 * 获取当前活动文档信息
 */
function getActiveDocument(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        return {
            success: true,
            data: {
                name: doc.Name,
                path: doc.FullName,
                paragraphCount: doc.Paragraphs.Count,
                wordCount: doc.Words.Count,
                characterCount: doc.Characters.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 读取文档文本内容
 */
function getDocumentText(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        var text = doc.Content.Text;
        var length = text.length;
        // 限制返回长度，防止内存爆炸
        if (length > 10000) {
            text = text.substring(0, 10000) + '...(truncated)';
        }
        return {
            success: true,
            data: { text: text, length: length }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 插入文本
 * @param {Object} params - { text: string, position?: 'start'|'end'|'cursor' }
 */
function insertText(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        var text = params.text || '';
        var position = params.position || 'cursor';

        switch (position) {
            case 'start':
                var range = doc.Range(0, 0);
                range.InsertBefore(text);
                break;
            case 'end':
                var endPos = doc.Content.End - 1;
                var range = doc.Range(endPos, endPos);
                range.InsertAfter(text);
                break;
            default: // cursor
                app.Selection.TypeText(text);
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 设置字体
 * @param {Object} params - { range?: 'all'|'selection', fontName?, fontSize?, bold?, italic? }
 */
function setFont(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        var range = (params.range === 'all') ? doc.Content : app.Selection.Range;

        if (params.fontName) range.Font.Name = params.fontName;
        if (params.fontSize) range.Font.Size = params.fontSize;
        if (params.bold !== undefined) range.Font.Bold = params.bold;
        if (params.italic !== undefined) range.Font.Italic = params.italic;

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 查找替换
 * @param {Object} params - { findText: string, replaceText: string, replaceAll?: boolean }
 */
function findReplace(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        var find = doc.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = params.findText;
        find.Replacement.Text = params.replaceText || '';
        // replaceType: 1=单个, 2=全部
        var replaceType = params.replaceAll ? 2 : 1;
        var result = find.Execute(
            params.findText,  // FindText
            false,            // MatchCase
            false,            // MatchWholeWord
            false,            // MatchWildcards
            false,            // MatchSoundsLike
            false,            // MatchAllWordForms
            true,             // Forward
            1,                // Wrap (wdFindContinue)
            false,            // Format
            params.replaceText || '',  // ReplaceWith
            replaceType       // Replace
        );
        return { success: true, data: { replaced: result } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 插入表格
 * @param {Object} params - { rows: number, cols: number, data?: array[][] }
 */
function insertTable(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: 'No active document' };
        }
        var rows = params.rows || 3;
        var cols = params.cols || 3;
        var range = app.Selection.Range;
        var table = doc.Tables.Add(range, rows, cols);

        // 填充数据
        if (params.data && Array.isArray(params.data)) {
            var maxRows = Math.min(params.data.length, rows);
            for (var r = 0; r < maxRows; r++) {
                var rowData = params.data[r];
                if (Array.isArray(rowData)) {
                    var maxCols = Math.min(rowData.length, cols);
                    for (var c = 0; c < maxCols; c++) {
                        table.Cell(r + 1, c + 1).Range.Text = String(rowData[c]);
                    }
                }
            }
        }
        // 启用边框
        table.Borders.Enable = true;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 应用样式
 * @param {Object} params - { styleName: string }
 */
function applyStyle(params) {
    try {
        var app = Application;
        var range = app.Selection.Range;
        range.Style = params.styleName;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 导出模块
module.exports = {
    getActiveDocument: getActiveDocument,
    getDocumentText: getDocumentText,
    insertText: insertText,
    setFont: setFont,
    findReplace: findReplace,
    insertTable: insertTable,
    applyStyle: applyStyle
};
