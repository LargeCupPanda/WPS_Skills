/**
 * Claude助手 - Mac版（轮询模式）
 * 加载项作为HTTP客户端，轮询MCP Server获取命令
 *
 * 架构：MCP Server (HTTP服务端:58891) ← 轮询 ← WPS加载项 (HTTP客户端)
 *
 * @author 老王
 */

var CONFIG = {
    SERVER_URL: 'http://127.0.0.1:58891',
    POLL_INTERVAL: 500  // 轮询间隔ms
};

var _ribbonUI = null;
var _pollTimer = null;
var _isPolling = false;

// ==================== 加载项生命周期 ====================

function OnAddinLoad(ribbonUI) {
    _ribbonUI = ribbonUI;
    console.log('=== Claude助手 (轮询模式) 加载中 ===');
    startPolling();
    return true;
}

function OnStatusClick() {
    var status = _isPolling ? '轮询中 (间隔: ' + CONFIG.POLL_INTERVAL + 'ms)' : '已停止';
    alert('Claude助手 状态: ' + status + '\n服务器: ' + CONFIG.SERVER_URL);
    return true;
}

// ==================== 轮询逻辑 ====================

function startPolling() {
    if (_pollTimer) return;
    _isPolling = true;
    console.log('开始轮询 MCP Server: ' + CONFIG.SERVER_URL);
    poll();
}

function stopPolling() {
    if (_pollTimer) {
        clearTimeout(_pollTimer);
        _pollTimer = null;
    }
    _isPolling = false;
    console.log('停止轮询');
}

function poll() {
    if (!_isPolling) return;

    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', CONFIG.SERVER_URL + '/poll', true);
        xhr.timeout = 5000;

        xhr.onload = function() {
            if (xhr.status === 200) {
                try {
                    var response = JSON.parse(xhr.responseText);
                    if (response.command) {
                        handleCommand(response.command);
                    }
                } catch (e) {
                    console.error('解析响应失败:', e);
                }
            }
            scheduleNextPoll();
        };

        xhr.onerror = function() {
            console.log('轮询失败，服务器可能未启动');
            scheduleNextPoll();
        };

        xhr.ontimeout = function() {
            scheduleNextPoll();
        };

        xhr.send();
    } catch (e) {
        console.error('轮询异常:', e);
        scheduleNextPoll();
    }
}

function scheduleNextPoll() {
    _pollTimer = setTimeout(poll, CONFIG.POLL_INTERVAL);
}

function sendResult(requestId, result) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', CONFIG.SERVER_URL + '/result', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send(JSON.stringify({
            requestId: requestId,
            result: result
        }));
    } catch (e) {
        console.error('发送结果失败:', e);
    }
}

// ==================== 工具函数 ====================

function getAppType() {
    try {
        if (typeof Application !== 'undefined') {
            var appName = '';
            try { appName = Application.Name || ''; } catch (e) {}

            // Mac版WPS可能返回不同的名称，增加更多匹配
            if (appName.indexOf('表格') !== -1 || appName.indexOf('Excel') !== -1 || appName.indexOf('ET') !== -1 || appName.indexOf('Spreadsheet') !== -1) {
                return 'et';
            }
            if (appName.indexOf('演示') !== -1 || appName.indexOf('Presentation') !== -1 || appName.indexOf('WPP') !== -1 || appName.indexOf('Slide') !== -1) {
                return 'wpp';
            }
            if (appName.indexOf('文字') !== -1 || appName.indexOf('Writer') !== -1 || appName.indexOf('Word') !== -1 || appName.indexOf('WPS') !== -1) {
                return 'wps';
            }

            // 兜底：通过检测活动对象类型判断
            try {
                if (Application.ActiveDocument) return 'wps';
            } catch (e) {}
            try {
                if (Application.ActiveWorkbook) return 'et';
            } catch (e) {}
            try {
                if (Application.ActivePresentation) return 'wpp';
            } catch (e) {}
        }
    } catch (e) {}
    return 'unknown';
}

// ==================== 命令处理 ====================

function handleCommand(cmd) {
    console.log('收到命令:', cmd.action);
    var result;

    try {
        switch (cmd.action) {
            // 通用
            case 'ping':
                result = { success: true, message: 'pong', timestamp: Date.now() };
                break;
            case 'wireCheck':
                result = { success: true, message: 'WPS MCP Bridge 已连接' };
                break;
            case 'getAppInfo':
                result = handleGetAppInfo();
                break;
            case 'getSelectedText':
                result = handleGetSelectedText();
                break;
            case 'setSelectedText':
                result = handleSetSelectedText(cmd.params);
                break;

            // Excel
            case 'getActiveWorkbook':
                result = handleGetActiveWorkbook();
                break;
            case 'getCellValue':
                result = handleGetCellValue(cmd.params);
                break;
            case 'setCellValue':
                result = handleSetCellValue(cmd.params);
                break;
            case 'getRangeData':
                result = handleGetRangeData(cmd.params);
                break;
            case 'setRangeData':
                result = handleSetRangeData(cmd.params);
                break;
            case 'setFormula':
                result = handleSetFormula(cmd.params);
                break;

            // Word
            case 'getActiveDocument':
                result = handleGetActiveDocument();
                break;
            case 'getDocumentText':
                result = handleGetDocumentText();
                break;
            case 'insertText':
                result = handleInsertText(cmd.params);
                break;

            // PPT
            case 'getActivePresentation':
                result = handleGetActivePresentation();
                break;
            case 'addSlide':
                result = handleAddSlide(cmd.params);
                break;
            case 'unifyFont':
                result = handleUnifyFont(cmd.params);
                break;
            case 'beautifySlide':
                result = handleBeautifySlide(cmd.params);
                break;

            // Word 高级功能
            case 'findReplace':
                result = handleFindReplace(cmd.params);
                break;
            case 'setFont':
                result = handleSetFont(cmd.params);
                break;
            case 'applyStyle':
                result = handleApplyStyle(cmd.params);
                break;
            case 'insertTable':
                result = handleInsertTable(cmd.params);
                break;
            case 'generateTOC':
                result = handleGenerateTOC(cmd.params);
                break;
            case 'setParagraph':
                result = handleSetParagraph(cmd.params);
                break;
            case 'insertPageBreak':
                result = handleInsertPageBreak(cmd.params);
                break;
            case 'setPageSetup':
                result = handleSetPageSetup(cmd.params);
                break;
            case 'insertHeader':
                result = handleInsertHeader(cmd.params);
                break;
            case 'insertFooter':
                result = handleInsertFooter(cmd.params);
                break;
            case 'insertHyperlink':
                result = handleInsertHyperlink(cmd.params);
                break;
            case 'insertBookmark':
                result = handleInsertBookmark(cmd.params);
                break;
            case 'getBookmarks':
                result = handleGetBookmarks(cmd.params);
                break;
            case 'addComment':
                result = handleAddComment(cmd.params);
                break;
            case 'getComments':
                result = handleGetComments(cmd.params);
                break;
            case 'getDocumentStats':
                result = handleGetDocumentStats(cmd.params);
                break;
            case 'insertImage':
                result = handleInsertImage(cmd.params);
                break;

            // Excel 高级功能
            case 'sortRange':
                result = handleSortRange(cmd.params);
                break;
            case 'autoFilter':
                result = handleAutoFilter(cmd.params);
                break;
            case 'createChart':
                result = handleCreateChart(cmd.params);
                break;
            case 'updateChart':
                result = handleUpdateChart(cmd.params);
                break;
            case 'createPivotTable':
                result = handleCreatePivotTable(cmd.params);
                break;
            case 'updatePivotTable':
                result = handleUpdatePivotTable(cmd.params);
                break;
            case 'removeDuplicates':
                result = handleRemoveDuplicates(cmd.params);
                break;
            case 'cleanData':
                result = handleCleanData(cmd.params);
                break;
            case 'getContext':
                result = handleGetContext(cmd.params);
                break;
            case 'diagnoseFormula':
                result = handleDiagnoseFormula(cmd.params);
                break;

            // Excel 工作表操作
            case 'createSheet':
                result = handleCreateSheet(cmd.params);
                break;
            case 'deleteSheet':
                result = handleDeleteSheet(cmd.params);
                break;
            case 'renameSheet':
                result = handleRenameSheet(cmd.params);
                break;
            case 'copySheet':
                result = handleCopySheet(cmd.params);
                break;
            case 'getSheetList':
                result = handleGetSheetList(cmd.params);
                break;
            case 'switchSheet':
                result = handleSwitchSheet(cmd.params);
                break;
            case 'moveSheet':
                result = handleMoveSheet(cmd.params);
                break;

            // Excel 单元格格式
            case 'setCellFormat':
                result = handleSetCellFormat(cmd.params);
                break;
            case 'setCellStyle':
                result = handleSetCellStyle(cmd.params);
                break;
            case 'mergeCells':
                result = handleMergeCells(cmd.params);
                break;
            case 'unmergeCells':
                result = handleUnmergeCells(cmd.params);
                break;
            case 'setColumnWidth':
                result = handleSetColumnWidth(cmd.params);
                break;
            case 'setRowHeight':
                result = handleSetRowHeight(cmd.params);
                break;
            case 'autoFitColumn':
                result = handleAutoFitColumn(cmd.params);
                break;
            case 'autoFitRow':
                result = handleAutoFitRow(cmd.params);
                break;
            case 'freezePanes':
                result = handleFreezePanes(cmd.params);
                break;
            case 'unfreezePanes':
                result = handleUnfreezePanes(cmd.params);
                break;

            // Excel 美化增强
            case 'autoFitAll':
                result = handleAutoFitAll(cmd.params);
                break;
            case 'copyFormat':
                result = handleCopyFormat(cmd.params);
                break;
            case 'clearFormats':
                result = handleClearFormats(cmd.params);
                break;
            case 'setBorder':
                result = handleSetBorder(cmd.params);
                break;
            case 'setNumberFormat':
                result = handleSetNumberFormat(cmd.params);
                break;

            // Excel 行列操作
            case 'insertRows':
                result = handleInsertRows(cmd.params);
                break;
            case 'insertColumns':
                result = handleInsertColumns(cmd.params);
                break;
            case 'deleteRows':
                result = handleDeleteRows(cmd.params);
                break;
            case 'deleteColumns':
                result = handleDeleteColumns(cmd.params);
                break;
            case 'hideRows':
                result = handleHideRows(cmd.params);
                break;
            case 'hideColumns':
                result = handleHideColumns(cmd.params);
                break;
            case 'showRows':
                result = handleShowRows(cmd.params);
                break;
            case 'showColumns':
                result = handleShowColumns(cmd.params);
                break;

            // Excel 条件格式
            case 'addConditionalFormat':
                result = handleAddConditionalFormat(cmd.params);
                break;
            case 'removeConditionalFormat':
                result = handleRemoveConditionalFormat(cmd.params);
                break;
            case 'getConditionalFormats':
                result = handleGetConditionalFormats(cmd.params);
                break;

            // Excel 数据验证
            case 'addDataValidation':
                result = handleAddDataValidation(cmd.params);
                break;
            case 'removeDataValidation':
                result = handleRemoveDataValidation(cmd.params);
                break;
            case 'getDataValidations':
                result = handleGetDataValidations(cmd.params);
                break;

            // Excel 查找替换
            case 'findInSheet':
                result = handleFindInSheet(cmd.params);
                break;
            case 'replaceInSheet':
                result = handleReplaceInSheet(cmd.params);
                break;

            // Excel 高级数据处理
            case 'copyRange':
                result = handleCopyRange(cmd.params);
                break;
            case 'pasteRange':
                result = handlePasteRange(cmd.params);
                break;
            case 'fillSeries':
                result = handleFillSeries(cmd.params);
                break;
            case 'transpose':
                result = handleTranspose(cmd.params);
                break;
            case 'textToColumns':
                result = handleTextToColumns(cmd.params);
                break;
            case 'subtotal':
                result = handleSubtotal(cmd.params);
                break;

            // Excel 命名区域
            case 'createNamedRange':
                result = handleCreateNamedRange(cmd.params);
                break;
            case 'deleteNamedRange':
                result = handleDeleteNamedRange(cmd.params);
                break;
            case 'getNamedRanges':
                result = handleGetNamedRanges(cmd.params);
                break;

            // Excel 批注功能
            case 'addCellComment':
                result = handleAddCellComment(cmd.params);
                break;
            case 'deleteCellComment':
                result = handleDeleteCellComment(cmd.params);
                break;
            case 'getCellComments':
                result = handleGetCellComments(cmd.params);
                break;

            // Excel 保护功能
            case 'protectSheet':
                result = handleProtectSheet(cmd.params);
                break;
            case 'unprotectSheet':
                result = handleUnprotectSheet(cmd.params);
                break;
            case 'protectWorkbook':
                result = handleProtectWorkbook(cmd.params);
                break;

            // P0 - 财务/金融核心功能
            case 'openWorkbook':
                result = handleOpenWorkbook(cmd.params);
                break;
            case 'getOpenWorkbooks':
                result = handleGetOpenWorkbooks(cmd.params);
                break;
            case 'switchWorkbook':
                result = handleSwitchWorkbook(cmd.params);
                break;
            case 'closeWorkbook':
                result = handleCloseWorkbook(cmd.params);
                break;
            case 'createWorkbook':
                result = handleCreateWorkbook(cmd.params);
                break;
            case 'getFormula':
                result = handleGetFormula(cmd.params);
                break;
            case 'getCellInfo':
                result = handleGetCellInfo(cmd.params);
                break;
            case 'clearRange':
                result = handleClearRange(cmd.params);
                break;

            // P1 - 财务/金融重要补充
            case 'refreshLinks':
                result = handleRefreshLinks(cmd.params);
                break;
            case 'consolidate':
                result = handleConsolidate(cmd.params);
                break;
            case 'setArrayFormula':
                result = handleSetArrayFormula(cmd.params);
                break;
            case 'calculateSheet':
                result = handleCalculateSheet(cmd.params);
                break;
            case 'insertExcelImage':
                result = handleInsertExcelImage(cmd.params);
                break;
            case 'setHyperlink':
                result = handleSetHyperlink(cmd.params);
                break;
            case 'wrapText':
                result = handleWrapText(cmd.params);
                break;

            // P2 - 扩展功能
            case 'setPrintArea':
                result = handleSetPrintArea(cmd.params);
                break;
            case 'getSelection':
                result = handleGetSelection(cmd.params);
                break;
            case 'groupRows':
                result = handleGroupRows(cmd.params);
                break;
            case 'groupColumns':
                result = handleGroupColumns(cmd.params);
                break;
            case 'lockCells':
                result = handleLockCells(cmd.params);
                break;

            // 通用高级功能
            case 'convertToPDF':
                result = handleConvertToPDF(cmd.params);
                break;
            case 'save':
                result = handleSave(cmd.params);
                break;
            case 'saveAs':
                result = handleSaveAs(cmd.params);
                break;

            default:
                result = { success: false, error: '未知命令: ' + cmd.action };
        }
    } catch (e) {
        result = { success: false, error: e.message || String(e) };
    }

    sendResult(cmd.requestId, result);
}

// ==================== 通用 Handlers ====================

function handleGetAppInfo() {
    try {
        var appType = getAppType();
        var appName = '';
        try { appName = Application.Name || ''; } catch (e) {}

        return {
            success: true,
            data: {
                appType: appType,
                appName: appName,
                hasSelection: !!(Application && Application.Selection)
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetSelectedText() {
    try {
        if (typeof Application !== 'undefined' && Application.Selection) {
            var text = Application.Selection.Text || '';
            return { success: true, data: { text: text.trim() } };
        }
        return { success: false, error: 'Application.Selection 不可用' };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetSelectedText(params) {
    try {
        if (typeof Application !== 'undefined' && Application.Selection) {
            Application.Selection.Text = params.text || '';
            return { success: true };
        }
        return { success: false, error: 'Application.Selection 不可用' };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel Handlers ====================

function handleGetActiveWorkbook() {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            sheets.push(wb.Sheets.Item(i).Name);
        }

        return {
            success: true,
            data: {
                name: wb.Name,
                path: wb.FullName,
                sheets: sheets,
                activeSheet: wb.ActiveSheet.Name
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 辅助函数：数字列号转字母
function colToLetter(col) {
    var letter = '';
    while (col > 0) {
        var mod = (col - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        col = Math.floor((col - 1) / 26);
    }
    return letter;
}

function handleGetCellValue(params) {
    try {
        var sheet = params.sheet || Application.ActiveSheet;
        if (typeof sheet === 'string') {
            sheet = Application.ActiveWorkbook.Sheets.Item(sheet);
        }
        // 支持两种方式：cell地址（如"A1"）或 row/col数字
        var cellAddr;
        if (params.cell) {
            cellAddr = params.cell;
        } else if (params.row && params.col) {
            // Mac版不支持sheet.Cells()，转成A1格式
            cellAddr = colToLetter(params.col) + params.row;
        } else {
            return { success: false, error: '请指定单元格位置(cell或row/col)' };
        }
        var cell = sheet.Range(cellAddr);
        return { success: true, data: { value: cell.Value, formula: cell.Formula } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetCellValue(params) {
    try {
        var sheet = params.sheet || Application.ActiveSheet;
        if (typeof sheet === 'string') {
            sheet = Application.ActiveWorkbook.Sheets.Item(sheet);
        }
        // 支持两种方式：cell地址（如"A1"）或 row/col数字
        var cellAddr;
        if (params.cell) {
            cellAddr = params.cell;
        } else if (params.row && params.col) {
            // Mac版不支持sheet.Cells()，转成A1格式
            cellAddr = colToLetter(params.col) + params.row;
        } else {
            return { success: false, error: '请指定单元格位置(cell或row/col)' };
        }
        var cell = sheet.Range(cellAddr);
        cell.Value2 = params.value;
        return { success: true, data: { cell: cellAddr } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetRangeData(params) {
    try {
        var sheet = params.sheet || Application.ActiveSheet;
        if (typeof sheet === 'string') {
            sheet = Application.ActiveWorkbook.Sheets.Item(sheet);
        }
        var range = sheet.Range(params.range);

        // Mac版WPS：用Range("A1")格式 + Value2
        var data = [];
        var rowCount = range.Rows.Count;
        var colCount = range.Columns.Count;
        var startRow = range.Row;
        var startCol = range.Column;

        for (var r = 0; r < rowCount; r++) {
            var rowData = [];
            for (var c = 0; c < colCount; c++) {
                var cellAddr = colToLetter(startCol + c) + (startRow + r);
                var cellVal = sheet.Range(cellAddr).Value2;
                rowData.push(cellVal !== undefined ? cellVal : null);
            }
            data.push(rowData);
        }

        return { success: true, data: { data: data, rows: rowCount, cols: colCount } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetRangeData(params) {
    try {
        var sheet = params.sheet || Application.ActiveSheet;
        if (typeof sheet === 'string') {
            sheet = Application.ActiveWorkbook.Sheets.Item(sheet);
        }
        // Mac版WPS不支持批量赋值range.Value = data，需要逐个单元格写入
        var data = params.data;
        if (!data || !Array.isArray(data)) {
            return { success: false, error: '数据格式错误，需要二维数组' };
        }

        // 解析起始位置
        var startRange = sheet.Range(params.range);
        var startRow = startRange.Row;
        var startCol = startRange.Column;

        // 逐个单元格写入
        for (var r = 0; r < data.length; r++) {
            var rowData = data[r];
            if (Array.isArray(rowData)) {
                for (var c = 0; c < rowData.length; c++) {
                    var cell = sheet.Cells(startRow + r, startCol + c);
                    cell.Value = rowData[c];
                }
            }
        }
        return { success: true, data: { rows: data.length, cols: data[0] ? data[0].length : 0 } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetFormula(params) {
    try {
        var sheet = params.sheet || Application.ActiveSheet;
        if (typeof sheet === 'string') {
            sheet = Application.ActiveWorkbook.Sheets.Item(sheet);
        }
        // 支持两种方式：cell地址（如"C10"）或 row/col数字
        var cell;
        if (params.cell) {
            cell = sheet.Range(params.cell);
        } else if (params.row && params.col) {
            cell = sheet.Cells(params.row, params.col);
        } else if (params.range) {
            cell = sheet.Range(params.range);
        } else {
            return { success: false, error: '请指定单元格位置(cell或row/col)' };
        }
        cell.Formula = params.formula;
        return { success: true, data: { cell: cell.Address, calculatedValue: cell.Value } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Word Handlers ====================

function handleGetActiveDocument() {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        return {
            success: true,
            data: {
                name: doc.Name,
                path: doc.FullName,
                paragraphs: doc.Paragraphs.Count,
                words: doc.Words.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetDocumentText() {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };
        // Mac版用Selection.WholeStory选中全文再获取
        var sel = Application.Selection;
        sel.WholeStory();
        var text = sel.Text || '';
        // 取消选择，恢复光标
        sel.Collapse(1);
        return { success: true, data: { text: text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertText(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var pos = params.position || 'end';
        var range;

        if (pos === 'start') {
            range = doc.Range(0, 0);
        } else if (pos === 'end') {
            range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
        } else {
            range = Application.Selection.Range;
        }

        range.InsertAfter(params.text);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT Handlers ====================

function handleGetActivePresentation() {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        return {
            success: true,
            data: {
                name: ppt.Name,
                path: ppt.FullName,
                slideCount: ppt.Slides.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleAddSlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        var index = params.index || ppt.Slides.Count + 1;
        var layout = params.layout || 1; // ppLayoutTitle
        var slide = ppt.Slides.Add(index, layout);

        if (params.title && slide.Shapes.HasTitle) {
            slide.Shapes.Title.TextFrame.TextRange.Text = params.title;
        }

        return { success: true, data: { slideIndex: slide.SlideIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 高级 Handlers ====================

var COLOR_SCHEMES = {
    business: { title: 0x2F5496, body: 0x333333 },
    tech: { title: 0x00B0F0, body: 0x404040 },
    creative: { title: 0xFF6B6B, body: 0x4A4A4A },
    minimal: { title: 0x000000, body: 0x666666 }
};

function handleUnifyFont(params) {
    try {
        var pres = Application.ActivePresentation;
        if (!pres) return { success: false, error: '没有打开的演示文稿' };

        var fontName = params.fontName || '微软雅黑';
        var count = 0;

        for (var i = 1; i <= pres.Slides.Count; i++) {
            var slide = pres.Slides.Item(i);
            for (var j = 1; j <= slide.Shapes.Count; j++) {
                var shape = slide.Shapes.Item(j);
                try {
                    if (shape.HasTextFrame && shape.TextFrame.HasText) {
                        shape.TextFrame.TextRange.Font.Name = fontName;
                        count++;
                    }
                } catch (e) {}
            }
        }

        return { success: true, data: { fontName: fontName, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleBeautifySlide(params) {
    try {
        var pres = Application.ActivePresentation;
        if (!pres) return { success: false, error: '没有打开的演示文稿' };

        var slideIndex = params.slideIndex || Application.ActiveWindow.Selection.SlideRange.SlideIndex;
        var slide = pres.Slides.Item(slideIndex);
        var scheme = COLOR_SCHEMES[params.style] || COLOR_SCHEMES.business;
        var count = 0;

        for (var j = 1; j <= slide.Shapes.Count; j++) {
            var shape = slide.Shapes.Item(j);
            try {
                if (shape.HasTextFrame && shape.TextFrame.HasText) {
                    var textRange = shape.TextFrame.TextRange;
                    if (textRange.Font.Size >= 24) {
                        textRange.Font.Color.RGB = scheme.title;
                    } else {
                        textRange.Font.Color.RGB = scheme.body;
                    }
                    count++;
                }
            } catch (e) {}
        }

        return { success: true, data: { style: params.style || 'business', count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Word 高级 Handlers ====================

function handleFindReplace(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var find = doc.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = params.findText;
        find.Replacement.Text = params.replaceText || '';
        var replaceType = params.replaceAll ? 2 : 1;
        var result = find.Execute(
            params.findText, false, false, false, false, false,
            true, 1, false, params.replaceText || '', replaceType
        );
        return { success: true, data: { replaced: result } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetFont(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var range = (params.range === 'all') ? doc.Content : Application.Selection.Range;
        if (params.fontName) range.Font.Name = params.fontName;
        if (params.fontSize) range.Font.Size = params.fontSize;
        if (params.bold !== undefined) range.Font.Bold = params.bold;
        if (params.italic !== undefined) range.Font.Italic = params.italic;

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleApplyStyle(params) {
    try {
        var range = Application.Selection.Range;
        range.Style = params.styleName;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertTable(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var rows = params.rows || 3;
        var cols = params.cols || 3;
        var range = Application.Selection.Range;
        var table = doc.Tables.Add(range, rows, cols);

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
        table.Borders.Enable = true;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGenerateTOC(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var position = params.position || 'start';
        var levels = params.levels || 3;
        var range;

        if (position === 'start') {
            range = doc.Range(0, 0);
            range.InsertBreak(7); // wdPageBreak
            range = doc.Range(0, 0);
        } else {
            range = Application.Selection.Range;
        }

        doc.TablesOfContents.Add(range, true, 1, levels, false, null, true, true, null, true);
        return { success: true, data: { levels: levels } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Word 高级功能扩展 ====================

function handleSetParagraph(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var range = params.range === 'all' ? doc.Content : Application.Selection.Range;
        var para = range.ParagraphFormat;

        // 对齐方式: left=0, center=1, right=2, justify=3
        if (params.alignment !== undefined) {
            var alignMap = { 'left': 0, 'center': 1, 'right': 2, 'justify': 3 };
            para.Alignment = alignMap[params.alignment] || 0;
        }

        // 行间距
        if (params.lineSpacing) {
            para.LineSpacingRule = 4; // wdLineSpaceMultiple
            para.LineSpacing = params.lineSpacing * 12; // 倍数转磅值
        }

        // 段前段后间距
        if (params.spaceBefore !== undefined) para.SpaceBefore = params.spaceBefore;
        if (params.spaceAfter !== undefined) para.SpaceAfter = params.spaceAfter;

        // 首行缩进
        if (params.firstLineIndent !== undefined) para.FirstLineIndent = params.firstLineIndent * 28.35; // 厘米转磅

        // 左右缩进
        if (params.leftIndent !== undefined) para.LeftIndent = params.leftIndent * 28.35;
        if (params.rightIndent !== undefined) para.RightIndent = params.rightIndent * 28.35;

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertPageBreak(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var breakType = params.type || 'page';
        var breakTypeMap = {
            'page': 7,      // wdPageBreak
            'column': 8,    // wdColumnBreak
            'section': 2,   // wdSectionBreakNextPage
            'sectionContinuous': 3  // wdSectionBreakContinuous
        };

        Application.Selection.InsertBreak(breakTypeMap[breakType] || 7);
        return { success: true, data: { type: breakType } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSetPageSetup(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var ps = doc.PageSetup;

        // 页边距 (厘米转磅)
        if (params.topMargin !== undefined) ps.TopMargin = params.topMargin * 28.35;
        if (params.bottomMargin !== undefined) ps.BottomMargin = params.bottomMargin * 28.35;
        if (params.leftMargin !== undefined) ps.LeftMargin = params.leftMargin * 28.35;
        if (params.rightMargin !== undefined) ps.RightMargin = params.rightMargin * 28.35;

        // 纸张方向: portrait=0, landscape=1
        if (params.orientation !== undefined) {
            ps.Orientation = params.orientation === 'landscape' ? 1 : 0;
        }

        // 纸张大小
        if (params.paperSize !== undefined) {
            var sizeMap = { 'A4': 7, 'A3': 6, 'Letter': 1, 'Legal': 5 };
            ps.PaperSize = sizeMap[params.paperSize] || 7;
        }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertHeader(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        // Mac版WPS API兼容处理
        var section = doc.Sections.Item(1) || doc.Sections(1);
        var header = section.Headers.Item(1) || section.Headers(1); // wdHeaderFooterPrimary
        header.Range.Text = params.text || '';

        if (params.alignment) {
            var alignMap = { 'left': 0, 'center': 1, 'right': 2 };
            header.Range.ParagraphFormat.Alignment = alignMap[params.alignment] || 1;
        }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertFooter(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        // Mac版WPS API兼容处理
        var section = doc.Sections.Item(1) || doc.Sections(1);
        var footer = section.Footers.Item(1) || section.Footers(1); // wdHeaderFooterPrimary
        footer.Range.Text = params.text || '';

        if (params.alignment) {
            var alignMap = { 'left': 0, 'center': 1, 'right': 2 };
            footer.Range.ParagraphFormat.Alignment = alignMap[params.alignment] || 1;
        }

        // 插入页码
        if (params.includePageNumber) {
            footer.Range.InsertAfter(' - 第 ');
            footer.Range.Fields.Add(footer.Range, -1, 'PAGE', false);
            footer.Range.InsertAfter(' 页 ');
        }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertHyperlink(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var range = Application.Selection.Range;
        var url = params.url || params.address;
        var text = params.text || params.displayText || url;

        if (range.Text && range.Text.trim() !== '') {
            // 选中了文本，把它变成超链接
            doc.Hyperlinks.Add(range, url);
        } else {
            // 没选中文本，插入新的超链接
            range.Text = text;
            doc.Hyperlinks.Add(doc.Range(range.Start, range.Start + text.length), url);
        }

        return { success: true, data: { url: url, text: text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertBookmark(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var name = params.name;
        if (!name) return { success: false, error: '书签名称不能为空' };

        var range = Application.Selection.Range;
        doc.Bookmarks.Add(name, range);

        return { success: true, data: { name: name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetBookmarks(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var bookmarks = [];
        var count = doc.Bookmarks.Count;
        for (var i = 1; i <= count; i++) {
            var bm = doc.Bookmarks.Item(i) || doc.Bookmarks(i);
            bookmarks.push({
                name: bm.Name,
                start: bm.Start,
                end: bm.End
            });
        }

        return { success: true, data: { bookmarks: bookmarks, count: bookmarks.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleAddComment(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var text = params.text || params.comment;
        if (!text) return { success: false, error: '批注内容不能为空' };

        var range = Application.Selection.Range;
        doc.Comments.Add(range, text);

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetComments(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var comments = [];
        for (var i = 1; i <= doc.Comments.Count; i++) {
            var c = doc.Comments(i);
            comments.push({
                index: i,
                text: c.Range.Text,
                author: c.Author || '',
                date: c.Date ? c.Date.toString() : ''
            });
        }

        return { success: true, data: { comments: comments, count: comments.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetDocumentStats(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var stats = {
            name: doc.Name,
            path: doc.FullName,
            pages: doc.ComputeStatistics(2), // wdStatisticPages
            words: doc.ComputeStatistics(0), // wdStatisticWords
            characters: doc.ComputeStatistics(3), // wdStatisticCharacters
            paragraphs: doc.ComputeStatistics(4), // wdStatisticParagraphs
            lines: doc.ComputeStatistics(1) // wdStatisticLines
        };

        return { success: true, data: stats };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleInsertImage(params) {
    try {
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var path = params.path || params.filePath;
        if (!path) return { success: false, error: '图片路径不能为空' };

        var range = Application.Selection.Range;
        var shape = doc.InlineShapes.AddPicture(path, false, true, range);

        // 调整大小
        if (params.width) shape.Width = params.width;
        if (params.height) shape.Height = params.height;

        // 保持比例缩放
        if (params.scale) {
            shape.ScaleWidth = params.scale;
            shape.ScaleHeight = params.scale;
        }

        return { success: true, data: { width: shape.Width, height: shape.Height } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 高级 Handlers ====================

function handleSortRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var keyCol = sheet.Range(params.keyColumn);
        var order = params.order === 'desc' ? 2 : 1;
        range.Sort(keyCol, order);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleAutoFilter(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        if (params.criteria) {
            range.AutoFilter(params.field, params.criteria);
        } else {
            range.AutoFilter();
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleCreateChart(params) {
    try {
        var sheet = Application.ActiveSheet;
        // 兼容两种参数名：dataRange 和 data_range
        var dataRange = params.dataRange || params.data_range;
        if (!dataRange) {
            return { success: false, error: '请指定数据范围(dataRange)' };
        }
        var range = sheet.Range(dataRange);
        // 兼容更多图表类型名
        var chartTypes = {
            column: 51, column_clustered: 51, bar: 57, bar_clustered: 57,
            line: 4, line_markers: 65, pie: 5, doughnut: -4120,
            area: 1, scatter: -4169, radar: -4151
        };
        var chartType = params.chartType || params.chart_type || 'column';
        var chartTypeNum = chartTypes[chartType] || (typeof chartType === 'number' ? chartType : 51);

        // 兼容position对象或直接left/top
        var pos = params.position || {};
        var left = pos.left || params.left || (range.Left + range.Width + 20);
        var top = pos.top || params.top || range.Top;
        var width = pos.width || params.width || 400;
        var height = pos.height || params.height || 300;

        var chartObj = sheet.ChartObjects().Add(left, top, width, height);
        chartObj.Chart.SetSourceData(range);
        chartObj.Chart.ChartType = chartTypeNum;

        if (params.title) {
            chartObj.Chart.HasTitle = true;
            chartObj.Chart.ChartTitle.Text = params.title;
        }
        return {
            success: true,
            data: {
                chartName: chartObj.Name,
                chartIndex: chartObj.Index || 1,
                dataRange: dataRange,
                chartType: chartType,
                position: { left: left, top: top, width: width, height: height }
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 更新图表
function handleUpdateChart(params) {
    try {
        var sheet = Application.ActiveSheet;
        var chartObj;

        // 通过索引或名称找到图表
        if (params.chartIndex) {
            chartObj = sheet.ChartObjects().Item(params.chartIndex);
        } else if (params.chartName) {
            chartObj = sheet.ChartObjects(params.chartName);
        } else {
            return { success: false, error: '请指定chartIndex或chartName' };
        }

        var chart = chartObj.Chart;
        var updated = [];

        // 更新标题
        if (params.title !== undefined) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = params.title;
            updated.push('title');
        }

        // 更新图表类型
        if (params.chartType !== undefined) {
            chart.ChartType = params.chartType;
            updated.push('chartType');
        }

        // 更新数据标签
        if (params.showDataLabels !== undefined) {
            for (var i = 1; i <= chart.SeriesCollection().Count; i++) {
                chart.SeriesCollection(i).HasDataLabels = params.showDataLabels;
            }
            updated.push('showDataLabels');
        }

        // 更新图例
        if (params.showLegend !== undefined) {
            chart.HasLegend = params.showLegend;
            updated.push('showLegend');
        }

        return {
            success: true,
            data: {
                chartName: chartObj.Name,
                updatedProperties: updated
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建透视表
function handleCreatePivotTable(params) {
    try {
        var wb = Application.ActiveWorkbook;
        var sourceSheet = Application.ActiveSheet;
        var sourceRange = sourceSheet.Range(params.sourceRange);

        // 目标工作表
        var destSheet = params.destinationSheet ?
            wb.Sheets.Item(params.destinationSheet) : sourceSheet;
        var destCell = destSheet.Range(params.destinationCell);

        // 创建透视表缓存
        var cache = wb.PivotCaches().Create(1, sourceRange); // xlDatabase = 1

        // 创建透视表
        var tableName = params.tableName || ('PivotTable' + (new Date()).getTime());
        var pivotTable = cache.CreatePivotTable(destCell, tableName);

        // 添加行字段
        if (params.rowFields && params.rowFields.length > 0) {
            for (var i = 0; i < params.rowFields.length; i++) {
                var field = pivotTable.PivotFields(params.rowFields[i]);
                field.Orientation = 1; // xlRowField
            }
        }

        // 添加列字段
        if (params.columnFields && params.columnFields.length > 0) {
            for (var i = 0; i < params.columnFields.length; i++) {
                var field = pivotTable.PivotFields(params.columnFields[i]);
                field.Orientation = 2; // xlColumnField
            }
        }

        // 添加值字段
        if (params.valueFields && params.valueFields.length > 0) {
            for (var i = 0; i < params.valueFields.length; i++) {
                var vf = params.valueFields[i];
                var field = pivotTable.PivotFields(vf.field);
                field.Orientation = 4; // xlDataField
                // 设置聚合函数
                var funcMap = { 'SUM': -4157, 'COUNT': -4112, 'AVERAGE': -4106, 'MAX': -4136, 'MIN': -4139 };
                if (vf.aggregation && funcMap[vf.aggregation]) {
                    field.Function = funcMap[vf.aggregation];
                }
            }
        }

        return {
            success: true,
            data: {
                pivotTableName: tableName,
                location: destCell.Address,
                rowCount: pivotTable.TableRange1.Rows.Count,
                columnCount: pivotTable.TableRange1.Columns.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 更新透视表
function handleUpdatePivotTable(params) {
    try {
        var sheet = Application.ActiveSheet;
        var pivotTable;

        // 找到透视表
        if (params.pivotTableName) {
            pivotTable = sheet.PivotTables(params.pivotTableName);
        } else if (params.pivotTableCell) {
            var cell = sheet.Range(params.pivotTableCell);
            pivotTable = cell.PivotTable;
        } else {
            return { success: false, error: '请指定pivotTableName或pivotTableCell' };
        }

        var operations = [];

        // 刷新数据
        if (params.refresh) {
            pivotTable.RefreshTable();
            operations.push({ operation: 'refresh', success: true, message: '刷新成功' });
        }

        // 添加行字段
        if (params.addRowFields) {
            for (var i = 0; i < params.addRowFields.length; i++) {
                try {
                    var field = pivotTable.PivotFields(params.addRowFields[i]);
                    field.Orientation = 1;
                    operations.push({ operation: 'addRowField', success: true, message: params.addRowFields[i] });
                } catch (e) {
                    operations.push({ operation: 'addRowField', success: false, message: e.message });
                }
            }
        }

        return {
            success: true,
            data: {
                pivotTableName: pivotTable.Name,
                operations: operations
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleRemoveDuplicates(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var hasHeader = params.has_header !== false ? 1 : 2; // xlYes=1, xlNo=2
        // columns参数：数组形式的列索引
        var cols = params.columns;
        if (!cols || cols.length === 0) {
            // 默认根据所有列判断重复
            var colCount = range.Columns.Count;
            cols = [];
            for (var i = 1; i <= colCount; i++) {
                cols.push(i);
            }
        }
        range.RemoveDuplicates(cols, hasHeader);
        return { success: true, data: { message: '删除重复行成功' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleCleanData(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var operations = params.operations || ['trim'];
        var startRow = range.Row;
        var startCol = range.Column;
        var opResults = [];

        // 处理每个操作
        for (var opIdx = 0; opIdx < operations.length; opIdx++) {
            var op = operations[opIdx];
            var count = 0;
            try {
                for (var r = 0; r < range.Rows.Count; r++) {
                    for (var c = 0; c < range.Columns.Count; c++) {
                        var cellAddr = colToLetter(startCol + c) + (startRow + r);
                        var cell = sheet.Range(cellAddr);
                        var val = cell.Value2;
                        if (val && typeof val === 'string') {
                            var newVal = val;
                            if (op === 'trim') {
                                newVal = newVal.replace(/^\s+|\s+$/g, '');
                            }
                            if (newVal !== val) {
                                cell.Value2 = newVal;
                                count++;
                            }
                        }
                    }
                }
                opResults.push({ operation: op, success: true, message: '处理了' + count + '个单元格' });
            } catch (opErr) {
                opResults.push({ operation: op, success: false, message: opErr.message });
            }
        }
        return { success: true, data: { range: params.range, operations: opResults } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleGetContext(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };

        var sheet = Application.ActiveSheet;
        var usedRange = sheet.UsedRange;
        var headers = [];
        var startCol = usedRange.Column;
        var startRow = usedRange.Row;
        var colCount = Math.min(usedRange.Columns.Count, 20);

        // 构建headers数组，格式: [{column: 'A', value: '姓名'}, ...]
        for (var c = 0; c < colCount; c++) {
            var colLetter = colToLetter(startCol + c);
            var cellAddr = colLetter + startRow;
            var val = sheet.Range(cellAddr).Value2;
            if (val) {
                headers.push({ column: colLetter, value: String(val) });
            }
        }

        // 获取所有工作表名称
        var allSheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            allSheets.push(wb.Sheets.Item(i).Name);
        }

        // 获取选中区域
        var selectedCell = 'A1';
        try {
            var sel = Application.Selection;
            if (sel && sel.Address) {
                // Mac版Address是方法，需要调用
                var addr = typeof sel.Address === 'function' ? sel.Address() : sel.Address;
                selectedCell = String(addr).replace(/\$/g, '');
            }
        } catch (e) {}

        // Mac版Address是方法，需要调用
        var usedAddr = typeof usedRange.Address === 'function' ? usedRange.Address() : usedRange.Address;

        return {
            success: true,
            data: {
                workbookName: wb.Name,
                currentSheet: sheet.Name,
                allSheets: allSheets,
                selectedCell: selectedCell,
                usedRangeAddress: String(usedAddr).replace(/\$/g, ''),
                headers: headers,
                rowCount: usedRange.Rows.Count,
                colCount: usedRange.Columns.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 公式诊断
function handleDiagnoseFormula(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cell = sheet.Range(params.cell);
        var formula = cell.Formula || '';
        var currentValue = cell.Value2;

        // 错误类型映射
        var errorTypes = {
            '#DIV/0!': { type: '#DIV/0!', diagnosis: '除数为零', suggestion: '检查公式中的除法运算，确保除数不为零' },
            '#VALUE!': { type: '#VALUE!', diagnosis: '参数类型错误', suggestion: '检查函数参数是否为正确的数据类型' },
            '#REF!': { type: '#REF!', diagnosis: '引用了不存在的单元格', suggestion: '检查公式中引用的单元格是否被删除' },
            '#NAME?': { type: '#NAME?', diagnosis: '函数名称错误或未定义的名称', suggestion: '检查函数名拼写是否正确' },
            '#N/A': { type: '#N/A', diagnosis: '查找函数未找到匹配值', suggestion: '检查查找值是否存在于数据源中' },
            '#NUM!': { type: '#NUM!', diagnosis: '数值问题', suggestion: '检查数值是否超出范围或参数是否有效' },
            '#NULL!': { type: '#NULL!', diagnosis: '交集为空', suggestion: '检查范围引用是否正确使用了冒号或逗号' }
        };

        var errorType = null;
        var diagnosis = '公式正常';
        var suggestion = '无需修复';

        // 检查是否有错误
        var valStr = String(currentValue);
        for (var errKey in errorTypes) {
            if (valStr.indexOf(errKey) !== -1) {
                errorType = errorTypes[errKey].type;
                diagnosis = errorTypes[errKey].diagnosis;
                suggestion = errorTypes[errKey].suggestion;
                break;
            }
        }

        // 获取引用的单元格
        var precedents = [];
        if (formula) {
            var matches = formula.match(/[A-Z]+[0-9]+/g);
            if (matches) {
                precedents = matches;
            }
        }

        return {
            success: true,
            data: {
                cell: params.cell,
                formula: formula,
                currentValue: currentValue,
                errorType: errorType,
                diagnosis: diagnosis,
                suggestion: suggestion,
                precedents: precedents
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 工作表操作 Handlers ====================

// 新建工作表
function handleCreateSheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var newSheet = wb.Sheets.Add();
        if (params.name) {
            newSheet.Name = params.name;
        }
        return { success: true, data: { sheetName: newSheet.Name, sheetIndex: newSheet.Index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除工作表
function handleDeleteSheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = params.sheet ? wb.Sheets.Item(params.sheet) : Application.ActiveSheet;
        var name = sheet.Name;
        Application.DisplayAlerts = false;
        sheet.Delete();
        Application.DisplayAlerts = true;
        return { success: true, data: { deletedSheet: name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 重命名工作表
function handleRenameSheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = params.sheet ? wb.Sheets.Item(params.sheet) : Application.ActiveSheet;
        var oldName = sheet.Name;
        sheet.Name = params.newName;
        return { success: true, data: { oldName: oldName, newName: params.newName } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 复制工作表
function handleCopySheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = params.sheet ? wb.Sheets.Item(params.sheet) : Application.ActiveSheet;
        if (params.before) {
            sheet.Copy(wb.Sheets.Item(params.before));
        } else if (params.after) {
            sheet.Copy(null, wb.Sheets.Item(params.after));
        } else {
            sheet.Copy(null, wb.Sheets.Item(wb.Sheets.Count));
        }
        return { success: true, data: { copiedFrom: sheet.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取工作表列表
function handleGetSheetList(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            var s = wb.Sheets.Item(i);
            sheets.push({ name: s.Name, index: i, visible: s.Visible });
        }
        return { success: true, data: { sheets: sheets, count: sheets.length, activeSheet: Application.ActiveSheet.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 切换活动工作表
function handleSwitchSheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = wb.Sheets.Item(params.sheet);
        sheet.Activate();
        return { success: true, data: { activeSheet: sheet.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移动工作表
function handleMoveSheet(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = params.sheet ? wb.Sheets.Item(params.sheet) : Application.ActiveSheet;
        if (params.before) {
            sheet.Move(wb.Sheets.Item(params.before));
        } else if (params.after) {
            sheet.Move(null, wb.Sheets.Item(params.after));
        }
        return { success: true, data: { movedSheet: sheet.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 单元格格式 Handlers ====================

// 设置单元格数字格式
function handleSetCellFormat(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        if (params.numberFormat) {
            range.NumberFormat = params.numberFormat;
        }
        return { success: true, data: { range: params.range, format: params.numberFormat } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置单元格样式（背景色、边框、对齐、字体）
function handleSetCellStyle(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);

        // 字号
        if (params.fontSize) {
            range.Font.Size = params.fontSize;
        }

        // 加粗
        if (params.bold !== undefined) {
            range.Font.Bold = params.bold;
        }

        // 斜体
        if (params.italic !== undefined) {
            range.Font.Italic = params.italic;
        }

        // 字体名称
        if (params.fontName) {
            range.Font.Name = params.fontName;
        }

        // 背景色
        if (params.backgroundColor) {
            var color = params.backgroundColor.replace('#', '');
            var r = parseInt(color.substr(0, 2), 16);
            var g = parseInt(color.substr(2, 2), 16);
            var b = parseInt(color.substr(4, 2), 16);
            range.Interior.Color = r + g * 256 + b * 65536;
        }

        // 字体颜色
        if (params.fontColor) {
            var fc = params.fontColor.replace('#', '');
            var fr = parseInt(fc.substr(0, 2), 16);
            var fg = parseInt(fc.substr(2, 2), 16);
            var fb = parseInt(fc.substr(4, 2), 16);
            range.Font.Color = fr + fg * 256 + fb * 65536;
        }

        // 对齐
        if (params.horizontalAlignment) {
            var hAlignMap = { left: -4131, center: -4108, right: -4152 };
            range.HorizontalAlignment = hAlignMap[params.horizontalAlignment] || -4108;
        }
        if (params.verticalAlignment) {
            var vAlignMap = { top: -4160, center: -4108, bottom: -4107 };
            range.VerticalAlignment = vAlignMap[params.verticalAlignment] || -4108;
        }

        // 边框
        if (params.border) {
            range.Borders.LineStyle = 1; // xlContinuous
            if (params.borderColor) {
                var bc = params.borderColor.replace('#', '');
                var br = parseInt(bc.substr(0, 2), 16);
                var bg = parseInt(bc.substr(2, 2), 16);
                var bb = parseInt(bc.substr(4, 2), 16);
                range.Borders.Color = br + bg * 256 + bb * 65536;
            }
        }

        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 合并单元格
function handleMergeCells(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.Merge(params.across || false);
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 取消合并单元格
function handleUnmergeCells(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.UnMerge();
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置列宽
function handleSetColumnWidth(params) {
    try {
        var sheet = Application.ActiveSheet;
        var col = params.column;
        if (typeof col === 'number') {
            col = colToLetter(col);
        }
        // 使用Range("A:A")格式选择整列
        sheet.Range(col + ':' + col).ColumnWidth = params.width;
        return { success: true, data: { column: col, width: params.width } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置行高
function handleSetRowHeight(params) {
    try {
        var sheet = Application.ActiveSheet;
        // 使用Range("1:1")格式选择整行
        sheet.Range(params.row + ':' + params.row).RowHeight = params.height;
        return { success: true, data: { row: params.row, height: params.height } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 自动调整列宽 - 使用Range格式
function handleAutoFitColumn(params) {
    try {
        var sheet = Application.ActiveSheet;
        if (params.range) {
            sheet.Range(params.range).Columns.AutoFit();
        } else if (params.column) {
            var col = params.column;
            if (typeof col === 'number') col = colToLetter(col);
            sheet.Range(col + ':' + col).AutoFit();
        } else {
            sheet.UsedRange.Columns.AutoFit();
        }
        return { success: true, data: { message: '列宽已自动调整' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 自动调整行高 - 使用Range格式
function handleAutoFitRow(params) {
    try {
        var sheet = Application.ActiveSheet;
        if (params.range) {
            sheet.Range(params.range).Rows.AutoFit();
        } else if (params.row) {
            sheet.Range(params.row + ':' + params.row).AutoFit();
        } else {
            sheet.UsedRange.Rows.AutoFit();
        }
        return { success: true, data: { message: '行高已自动调整' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 美化增强 Handlers ====================

// 一键自动调整所有列宽行高
function handleAutoFitAll(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = params.range ? sheet.Range(params.range) : sheet.UsedRange;
        range.Columns.AutoFit();
        range.Rows.AutoFit();
        return { success: true, data: { message: '列宽行高已自动调整' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 复制格式（格式刷）
function handleCopyFormat(params) {
    try {
        var sheet = Application.ActiveSheet;
        var sourceRange = sheet.Range(params.source);
        var targetRange = sheet.Range(params.target);
        sourceRange.Copy();
        targetRange.PasteSpecial(-4122); // xlPasteFormats
        Application.CutCopyMode = false;
        return { success: true, data: { source: params.source, target: params.target } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 清除格式保留内容
function handleClearFormats(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.ClearFormats();
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置边框样式
function handleSetBorder(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var styleMap = { 'thin': 1, 'medium': 2, 'thick': 4, 'double': 6, 'none': 0 };
        var style = styleMap[params.style] || 1;

        // 边框位置: all, left, right, top, bottom, inside, outside
        var position = params.position || 'all';
        var borders = [];

        if (position === 'all' || position === 'outside') {
            borders.push(7, 8, 9, 10); // xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        }
        if (position === 'all' || position === 'inside') {
            borders.push(11, 12); // xlInsideVertical, xlInsideHorizontal
        }
        if (position === 'left') borders.push(7);
        if (position === 'top') borders.push(8);
        if (position === 'bottom') borders.push(9);
        if (position === 'right') borders.push(10);

        for (var i = 0; i < borders.length; i++) {
            var border = range.Borders.Item(borders[i]);
            if (style === 0) {
                border.LineStyle = -4142; // xlNone
            } else {
                border.LineStyle = 1; // xlContinuous
                border.Weight = style;
            }
            if (params.color) {
                var c = params.color.replace('#', '');
                border.Color = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
            }
        }
        return { success: true, data: { range: params.range, style: params.style, position: position } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置数字格式
function handleSetNumberFormat(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var formatMap = {
            'number': '0.00',
            'integer': '0',
            'percent': '0.00%',
            'currency': '¥#,##0.00',
            'currencyUSD': '$#,##0.00',
            'date': 'yyyy-mm-dd',
            'datetime': 'yyyy-mm-dd hh:mm:ss',
            'time': 'hh:mm:ss',
            'text': '@',
            'scientific': '0.00E+00',
            'fraction': '# ?/?',
            'accounting': '_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ '
        };
        var format = formatMap[params.format] || params.format || '0.00';
        range.NumberFormat = format;
        return { success: true, data: { range: params.range, format: format } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 冻结窗格
function handleFreezePanes(params) {
    try {
        var sheet = Application.ActiveSheet;
        if (params.cell) {
            sheet.Range(params.cell).Select();
        } else if (params.row && params.column) {
            var cellAddr = colToLetter(params.column) + params.row;
            sheet.Range(cellAddr).Select();
        }
        Application.ActiveWindow.FreezePanes = true;
        return { success: true, data: { message: '窗格已冻结' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 取消冻结窗格
function handleUnfreezePanes(params) {
    try {
        Application.ActiveWindow.FreezePanes = false;
        return { success: true, data: { message: '窗格冻结已取消' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 行列操作 Handlers ====================

// 插入行
function handleInsertRows(params) {
    try {
        var sheet = Application.ActiveSheet;
        var startRow = params.row || params.startRow;
        var count = params.count || 1;
        for (var i = 0; i < count; i++) {
            // 使用Range("1:1")格式
            sheet.Range(startRow + ':' + startRow).Insert();
        }
        return { success: true, data: { insertedAt: startRow, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 插入列
function handleInsertColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var col = params.column || params.startColumn;
        if (typeof col === 'number') col = colToLetter(col);
        var count = params.count || 1;
        for (var i = 0; i < count; i++) {
            // 使用Range("A:A")格式
            sheet.Range(col + ':' + col).Insert();
        }
        return { success: true, data: { insertedAt: col, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除行
function handleDeleteRows(params) {
    try {
        var sheet = Application.ActiveSheet;
        var startRow = params.row || params.startRow;
        var count = params.count || 1;
        var endRow = startRow + count - 1;
        // 使用Range("1:3")格式
        sheet.Range(startRow + ':' + endRow).Delete();
        return { success: true, data: { deletedFrom: startRow, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除列
function handleDeleteColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var col = params.column || params.startColumn;
        if (typeof col === 'number') col = colToLetter(col);
        var count = params.count || 1;
        for (var i = 0; i < count; i++) {
            // 使用Range("A:A")格式
            sheet.Range(col + ':' + col).Delete();
        }
        return { success: true, data: { deletedFrom: col, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 隐藏行 - 使用Range("1:1")格式
function handleHideRows(params) {
    try {
        var sheet = Application.ActiveSheet;
        var rows = params.rows || [params.row];
        for (var i = 0; i < rows.length; i++) {
            sheet.Range(rows[i] + ':' + rows[i]).Hidden = true;
        }
        return { success: true, data: { hiddenRows: rows } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 隐藏列 - 使用Range("A:A")格式
function handleHideColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cols = params.columns || [params.column];
        for (var i = 0; i < cols.length; i++) {
            var c = cols[i];
            if (typeof c === 'number') c = colToLetter(c);
            sheet.Range(c + ':' + c).Hidden = true;
        }
        return { success: true, data: { hiddenColumns: cols } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 显示行 - 使用Range("1:1")格式
function handleShowRows(params) {
    try {
        var sheet = Application.ActiveSheet;
        var rows = params.rows || [params.row];
        for (var i = 0; i < rows.length; i++) {
            sheet.Range(rows[i] + ':' + rows[i]).Hidden = false;
        }
        return { success: true, data: { shownRows: rows } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 显示列 - 使用Range("A:A")格式
function handleShowColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cols = params.columns || [params.column];
        for (var i = 0; i < cols.length; i++) {
            var c = cols[i];
            if (typeof c === 'number') c = colToLetter(c);
            sheet.Range(c + ':' + c).Hidden = false;
        }
        return { success: true, data: { shownColumns: cols } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 条件格式 Handlers ====================

// 添加条件格式
function handleAddConditionalFormat(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var formatType = params.type || 'cellValue';

        if (formatType === 'cellValue') {
            // 单元格值条件格式
            var operatorMap = {
                'greater': 5, 'greaterThan': 5, 'less': 6, 'lessThan': 6,
                'equal': 3, 'notEqual': 4, 'greaterEqual': 7, 'greaterThanOrEqual': 7,
                'lessEqual': 8, 'lessThanOrEqual': 8, 'between': 1
            };
            var op = operatorMap[params.operator] || 3;
            var val1 = params.value1 || params.value;
            var val2 = params.value2;
            var cf = range.FormatConditions.Add(1, op, val1, val2);

            if (cf && params.backgroundColor) {
                var color = params.backgroundColor.replace('#', '');
                var r = parseInt(color.substr(0, 2), 16);
                var g = parseInt(color.substr(2, 2), 16);
                var b = parseInt(color.substr(4, 2), 16);
                cf.Interior.Color = r + g * 256 + b * 65536;
            }
            if (cf && params.fontColor) {
                var fc = params.fontColor.replace('#', '');
                cf.Font.Color = parseInt(fc.substr(0, 2), 16) + parseInt(fc.substr(2, 2), 16) * 256 + parseInt(fc.substr(4, 2), 16) * 65536;
            }
        } else if (formatType === 'colorScale') {
            // 色阶
            range.FormatConditions.AddColorScale(params.colorScaleType || 3);
        } else if (formatType === 'dataBar') {
            // 数据条
            range.FormatConditions.AddDatabar();
        }

        return { success: true, data: { range: params.range, type: formatType } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移除条件格式
function handleRemoveConditionalFormat(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        if (params.index) {
            range.FormatConditions.Item(params.index).Delete();
        } else {
            range.FormatConditions.Delete();
        }
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取条件格式列表
function handleGetConditionalFormats(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var formats = [];
        var count = range.FormatConditions.Count;
        for (var i = 1; i <= count; i++) {
            var cf = range.FormatConditions.Item(i);
            formats.push({ index: i, type: cf.Type });
        }
        return { success: true, data: { range: params.range, formats: formats, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 数据验证 Handlers ====================

// 添加数据验证
function handleAddDataValidation(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);

        // 验证类型映射
        var typeMap = {
            'list': 3, 'whole': 1, 'decimal': 2, 'date': 4, 'time': 5, 'textLength': 6, 'custom': 7
        };
        var validationType = typeMap[params.validationType] || 3;

        range.Validation.Delete(); // 先删除已有验证

        if (params.validationType === 'list') {
            range.Validation.Add(validationType, 1, 1, params.formula1 || params.list.join(','));
            if (params.showDropdown !== false) {
                range.Validation.InCellDropdown = true;
            }
        } else {
            var operatorMap = { 'between': 1, 'notBetween': 2, 'equal': 3, 'notEqual': 4, 'greater': 5, 'less': 6, 'greaterEqual': 7, 'lessEqual': 8 };
            var op = operatorMap[params.operator] || 1;
            range.Validation.Add(validationType, 1, op, params.formula1, params.formula2);
        }

        if (params.inputTitle || params.inputMessage) {
            range.Validation.InputTitle = params.inputTitle || '';
            range.Validation.InputMessage = params.inputMessage || '';
        }
        if (params.errorTitle || params.errorMessage) {
            range.Validation.ErrorTitle = params.errorTitle || '';
            range.Validation.ErrorMessage = params.errorMessage || '';
        }

        return { success: true, data: { range: params.range, type: params.validationType } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移除数据验证
function handleRemoveDataValidation(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.Validation.Delete();
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取数据验证
function handleGetDataValidations(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var validation = range.Validation;
        return {
            success: true,
            data: {
                range: params.range,
                type: validation.Type,
                formula1: validation.Formula1,
                formula2: validation.Formula2,
                inputTitle: validation.InputTitle,
                inputMessage: validation.InputMessage
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 查找替换 Handlers ====================

// 在工作表中查找
function handleFindInSheet(params) {
    try {
        var sheet = Application.ActiveSheet;
        var searchRange = params.range ? sheet.Range(params.range) : sheet.UsedRange;
        var results = [];
        var found = searchRange.Find(params.searchText, null, -4163, params.matchCase ? 1 : 2);

        if (found) {
            var firstAddr = typeof found.Address === 'function' ? found.Address() : found.Address;
            do {
                var addr = typeof found.Address === 'function' ? found.Address() : found.Address;
                results.push({ address: addr.replace(/\$/g, ''), value: found.Value2 });
                found = searchRange.FindNext(found);
                var currAddr = typeof found.Address === 'function' ? found.Address() : found.Address;
            } while (found && currAddr !== firstAddr);
        }

        return { success: true, data: { searchText: params.searchText, results: results, count: results.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 在工作表中替换
function handleReplaceInSheet(params) {
    try {
        var sheet = Application.ActiveSheet;
        var searchRange = params.range ? sheet.Range(params.range) : sheet.UsedRange;
        var replaced = searchRange.Replace(params.searchText, params.replaceText, params.matchCase ? 1 : 2);
        return { success: true, data: { searchText: params.searchText, replaceText: params.replaceText, success: replaced } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 高级数据处理 Handlers ====================

// 复制范围
function handleCopyRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.Copy();
        return { success: true, data: { range: params.range, message: '已复制到剪贴板' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 粘贴范围
function handlePasteRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var destRange = sheet.Range(params.destination);

        if (params.pasteType === 'values') {
            destRange.PasteSpecial(-4163); // xlPasteValues
        } else if (params.pasteType === 'formats') {
            destRange.PasteSpecial(-4122); // xlPasteFormats
        } else if (params.pasteType === 'formulas') {
            destRange.PasteSpecial(-4123); // xlPasteFormulas
        } else {
            sheet.Paste(destRange);
        }

        return { success: true, data: { destination: params.destination } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 填充序列
function handleFillSeries(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);

        // 设置起始值
        var startCell = range.Cells.Item(1, 1);
        startCell.Value2 = params.startValue || 1;

        // 填充类型
        var typeMap = { 'linear': 0, 'growth': 1, 'date': 2, 'autoFill': 3 };
        var fillType = typeMap[params.type] || 0;

        var step = params.step || 1;
        range.DataSeries(null, -4132, fillType, step);

        return { success: true, data: { range: params.range, type: params.type } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 转置数据
function handleTranspose(params) {
    try {
        var sheet = Application.ActiveSheet;
        var sourceRange = sheet.Range(params.sourceRange);
        var destCell = params.destinationCell || params.targetCell;
        var destRange = sheet.Range(destCell);

        sourceRange.Copy();
        destRange.PasteSpecial(-4163, -4142, false, true); // xlPasteValues, transpose=true
        Application.CutCopyMode = false;

        return { success: true, data: { source: params.sourceRange, destination: destCell } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 分列
function handleTextToColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var delimiter = params.delimiter || ',';

        var delimiterMap = { ',': 1, '\t': 2, ';': 3, ' ': 4 };
        var delimType = delimiterMap[delimiter] || 1;

        range.TextToColumns(null, 1, 1, false, true);
        return { success: true, data: { range: params.range, delimiter: delimiter } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 分类汇总
function handleSubtotal(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);

        var funcMap = { 'sum': 9, 'count': 2, 'average': 1, 'max': 4, 'min': 5 };
        var func = funcMap[params.function] || 9;

        // totalColumn需要是数组格式
        var totalCol = params.totalColumn;
        if (!Array.isArray(totalCol)) {
            totalCol = [totalCol];
        }

        range.Subtotal(params.groupBy, func, totalCol, params.replace !== false, false, true);
        return { success: true, data: { range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 命名区域 Handlers ====================

// 创建命名区域
function handleCreateNamedRange(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var addr = typeof range.Address === 'function' ? range.Address() : range.Address;
        wb.Names.Add(params.name, '=' + sheet.Name + '!' + addr);
        return { success: true, data: { name: params.name, range: params.range } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除命名区域
function handleDeleteNamedRange(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        wb.Names.Item(params.name).Delete();
        return { success: true, data: { deletedName: params.name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取命名区域列表
function handleGetNamedRanges(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var names = [];
        for (var i = 1; i <= wb.Names.Count; i++) {
            var n = wb.Names.Item(i);
            names.push({ name: n.Name, refersTo: n.RefersTo, visible: n.Visible });
        }
        return { success: true, data: { names: names, count: names.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 批注功能 Handlers ====================

// 添加单元格批注
function handleAddCellComment(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cell = sheet.Range(params.cell);
        if (cell.Comment) {
            cell.Comment.Delete();
        }
        cell.AddComment(params.text);
        if (params.visible) {
            cell.Comment.Visible = true;
        }
        return { success: true, data: { cell: params.cell, text: params.text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除单元格批注
function handleDeleteCellComment(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cell = sheet.Range(params.cell);
        if (cell.Comment) {
            cell.Comment.Delete();
        }
        return { success: true, data: { cell: params.cell } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取批注列表
function handleGetCellComments(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = params.range ? sheet.Range(params.range) : sheet.UsedRange;
        var comments = [];

        for (var i = 1; i <= sheet.Comments.Count; i++) {
            var c = sheet.Comments.Item(i);
            var addr = typeof c.Parent.Address === 'function' ? c.Parent.Address() : c.Parent.Address;
            comments.push({
                cell: addr.replace(/\$/g, ''),
                text: c.Text(),
                author: c.Author || ''
            });
        }

        return { success: true, data: { comments: comments, count: comments.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== Excel 保护功能 Handlers ====================

// 保护工作表
function handleProtectSheet(params) {
    try {
        var sheet = params.sheet ? Application.ActiveWorkbook.Sheets.Item(params.sheet) : Application.ActiveSheet;
        var password = params.password || '';
        sheet.Protect(password, params.drawingObjects, params.contents, params.scenarios);
        return { success: true, data: { sheet: sheet.Name, protected: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 取消保护工作表
function handleUnprotectSheet(params) {
    try {
        var sheet = params.sheet ? Application.ActiveWorkbook.Sheets.Item(params.sheet) : Application.ActiveSheet;
        var password = params.password || '';
        sheet.Unprotect(password);
        return { success: true, data: { sheet: sheet.Name, protected: false } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 保护工作簿
function handleProtectWorkbook(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var password = params.password || '';
        wb.Protect(password, params.structure !== false, params.windows);
        return { success: true, data: { workbook: wb.Name, protected: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== P0 财务/金融核心功能 Handlers ====================

// 打开工作簿
function handleOpenWorkbook(params) {
    try {
        var wb = Application.Workbooks.Open(params.path, params.updateLinks, params.readOnly);
        return { success: true, data: { name: wb.Name, path: wb.FullName, sheets: wb.Sheets.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取所有打开的工作簿
function handleGetOpenWorkbooks(params) {
    try {
        var workbooks = [];
        for (var i = 1; i <= Application.Workbooks.Count; i++) {
            var wb = Application.Workbooks.Item(i);
            workbooks.push({ name: wb.Name, path: wb.FullName, sheets: wb.Sheets.Count, active: wb.Name === Application.ActiveWorkbook.Name });
        }
        return { success: true, data: { workbooks: workbooks, count: workbooks.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 切换活动工作簿
function handleSwitchWorkbook(params) {
    try {
        var wb = Application.Workbooks.Item(params.name || params.index);
        wb.Activate();
        return { success: true, data: { name: wb.Name, path: wb.FullName } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 关闭工作簿
function handleCloseWorkbook(params) {
    try {
        var wb = params.name ? Application.Workbooks.Item(params.name) : Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有找到工作簿' };
        var name = wb.Name;
        wb.Close(params.saveChanges !== false);
        return { success: true, data: { closed: name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建新工作簿
function handleCreateWorkbook(params) {
    try {
        var wb = Application.Workbooks.Add();
        if (params.name) {
            // 如果指定了名称，另存为
            wb.SaveAs(params.name);
        }
        return { success: true, data: { name: wb.Name, path: wb.FullName, sheets: wb.Sheets.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取单元格公式
function handleGetFormula(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cell = sheet.Range(params.cell);
        var formula = cell.Formula || '';
        var formulaLocal = cell.FormulaLocal || '';
        return { success: true, data: { cell: params.cell, formula: formula, formulaLocal: formulaLocal, hasFormula: formula.indexOf('=') === 0 } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取单元格完整信息
function handleGetCellInfo(params) {
    try {
        var sheet = Application.ActiveSheet;
        var cell = sheet.Range(params.cell);
        var value = cell.Value2;
        var formula = cell.Formula || '';
        var numberFormat = cell.NumberFormat || '';
        var fontName = cell.Font.Name || '';
        var fontSize = cell.Font.Size || 0;
        var bold = cell.Font.Bold || false;
        var bgColor = cell.Interior.Color || 0;
        return { success: true, data: { cell: params.cell, value: value, formula: formula, numberFormat: numberFormat, font: { name: fontName, size: fontSize, bold: bold }, backgroundColor: bgColor } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 清除范围
function handleClearRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var clearType = params.type || 'all';
        if (clearType === 'contents') {
            range.ClearContents();
        } else if (clearType === 'formats') {
            range.ClearFormats();
        } else if (clearType === 'comments') {
            range.ClearComments();
        } else {
            range.Clear();
        }
        return { success: true, data: { range: params.range, clearType: clearType } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== P1 财务/金融重要补充 Handlers ====================

// 刷新外部链接
function handleRefreshLinks(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: '没有打开的工作簿' };
        var links = wb.LinkSources(1); // xlExcelLinks
        if (links) {
            for (var i = 1; i <= links.length; i++) {
                wb.UpdateLink(links[i - 1], 1);
            }
            return { success: true, data: { refreshed: links.length } };
        }
        return { success: true, data: { refreshed: 0, message: '没有外部链接' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 合并计算
function handleConsolidate(params) {
    try {
        var sheet = Application.ActiveSheet;
        var destRange = sheet.Range(params.destination);
        var funcMap = { 'sum': 9, 'count': 2, 'average': 1, 'max': 4, 'min': 5 };
        var func = funcMap[params.function] || 9;
        destRange.Consolidate(params.sources, func, params.topRow, params.leftColumn, params.createLinks);
        return { success: true, data: { destination: params.destination, sources: params.sources } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置数组公式
function handleSetArrayFormula(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.FormulaArray = params.formula;
        return { success: true, data: { range: params.range, formula: params.formula } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 强制重算工作表
function handleCalculateSheet(params) {
    try {
        if (params.all) {
            Application.Calculate();
            return { success: true, data: { calculated: 'all' } };
        } else {
            var sheet = Application.ActiveSheet;
            sheet.Calculate();
            return { success: true, data: { calculated: sheet.Name } };
        }
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 插入图片(Excel)
function handleInsertExcelImage(params) {
    try {
        var sheet = Application.ActiveSheet;
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || -1;
        var height = params.height || -1;
        var pic = sheet.Shapes.AddPicture(params.path, false, true, left, top, width, height);
        return { success: true, data: { name: pic.Name, path: params.path } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置超链接
function handleSetHyperlink(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.cell);
        sheet.Hyperlinks.Add(range, params.address || '', params.subAddress || '', params.screenTip || '', params.textToDisplay || '');
        return { success: true, data: { cell: params.cell, address: params.address } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 自动换行
function handleWrapText(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.WrapText = params.wrap !== false;
        return { success: true, data: { range: params.range, wrapText: params.wrap !== false } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== P2 扩展功能 Handlers ====================

// 设置打印区域
function handleSetPrintArea(params) {
    try {
        var sheet = Application.ActiveSheet;
        if (params.range) {
            sheet.PageSetup.PrintArea = params.range;
        } else {
            sheet.PageSetup.PrintArea = '';
        }
        return { success: true, data: { printArea: params.range || 'cleared' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取当前选中区域
function handleGetSelection(params) {
    try {
        var sel = Application.Selection;
        if (!sel) return { success: false, error: '没有选中区域' };
        var addr = typeof sel.Address === 'function' ? sel.Address() : sel.Address;
        var rowCount = sel.Rows.Count;
        var colCount = sel.Columns.Count;
        return { success: true, data: { address: addr, rows: rowCount, columns: colCount } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 行分组
function handleGroupRows(params) {
    try {
        var sheet = Application.ActiveSheet;
        var startRow = params.startRow;
        var endRow = params.endRow;
        sheet.Range(startRow + ':' + endRow).Group();
        return { success: true, data: { grouped: startRow + ':' + endRow } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 列分组
function handleGroupColumns(params) {
    try {
        var sheet = Application.ActiveSheet;
        var startCol = params.startColumn;
        var endCol = params.endColumn;
        sheet.Range(startCol + ':' + endCol).Group();
        return { success: true, data: { grouped: startCol + ':' + endCol } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 锁定单元格
function handleLockCells(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        range.Locked = params.locked !== false;
        return { success: true, data: { range: params.range, locked: params.locked !== false } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== 通用高级 Handlers ====================

function handleConvertToPDF(params) {
    try {
        var appType = getAppType();
        var outputPath = params.outputPath;

        if (appType === 'wps') {
            var doc = Application.ActiveDocument;
            if (!doc) return { success: false, error: '没有打开的文档' };
            if (!outputPath) outputPath = doc.FullName.replace(/\.\w+$/, '.pdf');
            doc.SaveAs2(outputPath, 17); // wdFormatPDF
        } else if (appType === 'et') {
            var wb = Application.ActiveWorkbook;
            if (!wb) return { success: false, error: '没有打开的工作簿' };
            if (!outputPath) outputPath = wb.FullName.replace(/\.\w+$/, '.pdf');
            wb.ExportAsFixedFormat(0, outputPath); // xlTypePDF
        } else if (appType === 'wpp') {
            var pres = Application.ActivePresentation;
            if (!pres) return { success: false, error: '没有打开的演示文稿' };
            if (!outputPath) outputPath = pres.FullName.replace(/\.\w+$/, '.pdf');
            pres.SaveAs(outputPath, 32); // ppSaveAsPDF
        } else {
            return { success: false, error: '无法识别当前应用类型' };
        }

        return { success: true, data: { outputPath: outputPath } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSave(params) {
    try {
        var appType = getAppType();
        var filePath = '';

        if (appType === 'wps') {
            var doc = Application.ActiveDocument;
            if (!doc) return { success: false, error: '没有打开的文档' };
            doc.Save();
            filePath = doc.FullName;
        } else if (appType === 'et') {
            var wb = Application.ActiveWorkbook;
            if (!wb) return { success: false, error: '没有打开的工作簿' };
            wb.Save();
            filePath = wb.FullName;
        } else if (appType === 'wpp') {
            var pres = Application.ActivePresentation;
            if (!pres) return { success: false, error: '没有打开的演示文稿' };
            pres.Save();
            filePath = pres.FullName;
        } else {
            return { success: false, error: '无法识别当前应用类型' };
        }

        return { success: true, data: { filePath: filePath } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

function handleSaveAs(params) {
    try {
        var appType = getAppType();
        var outputPath = params.path || params.outputPath;
        if (!outputPath) return { success: false, error: '请指定保存路径' };

        if (appType === 'wps') {
            var doc = Application.ActiveDocument;
            if (!doc) return { success: false, error: '没有打开的文档' };
            doc.SaveAs2(outputPath);
        } else if (appType === 'et') {
            var wb = Application.ActiveWorkbook;
            if (!wb) return { success: false, error: '没有打开的工作簿' };
            wb.SaveAs(outputPath);
        } else if (appType === 'wpp') {
            var pres = Application.ActivePresentation;
            if (!pres) return { success: false, error: '没有打开的演示文稿' };
            pres.SaveAs(outputPath);
        } else {
            return { success: false, error: '无法识别当前应用类型' };
        }

        return { success: true, data: { filePath: outputPath } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}
