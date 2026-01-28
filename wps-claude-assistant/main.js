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
            case 'getOpenDocuments':
                result = handleGetOpenDocuments(cmd.params);
                break;
            case 'switchDocument':
                result = handleSwitchDocument(cmd.params);
                break;
            case 'openDocument':
                result = handleOpenDocument(cmd.params);
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

            // PPT 演示文稿操作
            case 'createPresentation':
                result = handleCreatePresentation(cmd.params);
                break;
            case 'openPresentation':
                result = handleOpenPresentation(cmd.params);
                break;
            case 'closePresentation':
                result = handleClosePresentation(cmd.params);
                break;
            case 'getOpenPresentations':
                result = handleGetOpenPresentations(cmd.params);
                break;
            case 'switchPresentation':
                result = handleSwitchPresentation(cmd.params);
                break;

            // PPT 幻灯片操作
            case 'deleteSlide':
                result = handleDeleteSlide(cmd.params);
                break;
            case 'duplicateSlide':
                result = handleDuplicateSlide(cmd.params);
                break;
            case 'moveSlide':
                result = handleMoveSlide(cmd.params);
                break;
            case 'getSlideCount':
                result = handleGetSlideCount(cmd.params);
                break;
            case 'getSlideInfo':
                result = handleGetSlideInfo(cmd.params);
                break;
            case 'switchSlide':
                result = handleSwitchSlide(cmd.params);
                break;
            case 'setSlideLayout':
                result = handleSetSlideLayout(cmd.params);
                break;
            case 'getSlideNotes':
                result = handleGetSlideNotes(cmd.params);
                break;
            case 'setSlideNotes':
                result = handleSetSlideNotes(cmd.params);
                break;

            // PPT 文本框操作
            case 'addTextBox':
                result = handleAddTextBox(cmd.params);
                break;
            case 'deleteTextBox':
                result = handleDeleteTextBox(cmd.params);
                break;
            case 'getTextBoxes':
                result = handleGetTextBoxes(cmd.params);
                break;
            case 'setTextBoxText':
                result = handleSetTextBoxText(cmd.params);
                break;
            case 'setTextBoxStyle':
                result = handleSetTextBoxStyle(cmd.params);
                break;

            // PPT 标题操作
            case 'setSlideTitle':
                result = handleSetSlideTitle(cmd.params);
                break;
            case 'getSlideTitle':
                result = handleGetSlideTitle(cmd.params);
                break;
            case 'setSlideSubtitle':
                result = handleSetSlideSubtitle(cmd.params);
                break;
            case 'setSlideContent':
                result = handleSetSlideContent(cmd.params);
                break;

            // PPT 形状操作
            case 'addShape':
                result = handleAddShape(cmd.params);
                break;
            case 'deleteShape':
                result = handleDeleteShape(cmd.params);
                break;
            case 'getShapes':
                result = handleGetShapes(cmd.params);
                break;
            case 'setShapeStyle':
                result = handleSetShapeStyle(cmd.params);
                break;
            case 'setShapeText':
                result = handleSetShapeText(cmd.params);
                break;
            case 'setShapePosition':
                result = handleSetShapePosition(cmd.params);
                break;

            // PPT 图片操作
            case 'insertPptImage':
                result = handleInsertPptImage(cmd.params);
                break;
            case 'deletePptImage':
                result = handleDeletePptImage(cmd.params);
                break;
            case 'setImageStyle':
                result = handleSetImageStyle(cmd.params);
                break;

            // PPT 表格操作
            case 'insertPptTable':
                result = handleInsertPptTable(cmd.params);
                break;
            case 'setPptTableCell':
                result = handleSetPptTableCell(cmd.params);
                break;
            case 'getPptTableCell':
                result = handleGetPptTableCell(cmd.params);
                break;
            case 'setPptTableStyle':
                result = handleSetPptTableStyle(cmd.params);
                break;
            case 'setPptTableCellStyle':
                result = handleSetPptTableCellStyle(cmd.params);
                break;
            case 'setPptTableRowStyle':
                result = handleSetPptTableRowStyle(cmd.params);
                break;
            case 'setShapeShadow':
                result = handleSetShapeShadow(cmd.params);
                break;
            case 'setBackgroundGradient':
                result = handleSetBackgroundGradient(cmd.params);
                break;
            case 'setShapeGradient':
                result = handleSetShapeGradient(cmd.params);
                break;
            case 'setShapeBorder':
                result = handleSetShapeBorder(cmd.params);
                break;
            case 'setShapeTransparency':
                result = handleSetShapeTransparency(cmd.params);
                break;
            case 'setShapeRoundness':
                result = handleSetShapeRoundness(cmd.params);
                break;
            case 'setShapeFullStyle':
                result = handleSetShapeFullStyle(cmd.params);
                break;
            case 'alignShapes':
                result = handleAlignShapes(cmd.params);
                break;
            case 'distributeShapes':
                result = handleDistributeShapes(cmd.params);
                break;
            case 'groupShapes':
                result = handleGroupShapes(cmd.params);
                break;
            case 'duplicateShape':
                result = handleDuplicateShape(cmd.params);
                break;
            case 'setShapeZOrder':
                result = handleSetShapeZOrder(cmd.params);
                break;
            case 'addConnector':
                result = handleAddConnector(cmd.params);
                break;
            case 'addArrow':
                result = handleAddArrow(cmd.params);
                break;

            // PPT 专业美化工具
            case 'applyColorScheme':
                result = handleApplyColorScheme(cmd.params);
                break;
            case 'autoBeautifySlide':
                result = handleAutoBeautifySlide(cmd.params);
                break;
            case 'createKpiCards':
                result = handleCreateKpiCards(cmd.params);
                break;
            case 'createStyledTable':
                result = handleCreateStyledTable(cmd.params);
                break;
            case 'addTitleDecoration':
                result = handleAddTitleDecoration(cmd.params);
                break;
            case 'addPageIndicator':
                result = handleAddPageIndicator(cmd.params);
                break;
            case 'beautifyAllSlides':
                result = handleBeautifyAllSlides(cmd.params);
                break;

            // PPT 高端能力 - 数据可视化组件
            case 'createProgressBar':
                result = handleCreateProgressBar(cmd.params);
                break;
            case 'createGauge':
                result = handleCreateGauge(cmd.params);
                break;
            case 'createMiniCharts':
                result = handleCreateMiniCharts(cmd.params);
                break;
            case 'createDonutChart':
                result = handleCreateDonutChart(cmd.params);
                break;

            // PPT 高端能力 - 智能布局工具
            case 'autoLayout':
                result = handleAutoLayout(cmd.params);
                break;
            case 'smartDistribute':
                result = handleSmartDistribute(cmd.params);
                break;
            case 'createGrid':
                result = handleCreateGrid(cmd.params);
                break;

            // PPT 高端能力 - 高级动画组合
            case 'addAnimationPreset':
                result = handleAddAnimationPreset(cmd.params);
                break;
            case 'addEmphasisAnimation':
                result = handleAddEmphasisAnimation(cmd.params);
                break;

            // PPT 高端能力 - 流程图/组织架构图
            case 'createFlowChart':
                result = handleCreateFlowChart(cmd.params);
                break;
            case 'createOrgChart':
                result = handleCreateOrgChart(cmd.params);
                break;
            case 'createTimeline':
                result = handleCreateTimeline(cmd.params);
                break;

            // PPT 高端能力 - 母版操作
            case 'getSlideMaster':
                result = handleGetSlideMaster(cmd.params);
                break;
            case 'setMasterBackground':
                result = handleSetMasterBackground(cmd.params);
                break;
            case 'addMasterElement':
                result = handleAddMasterElement(cmd.params);
                break;

            // PPT 高端能力 - 3D效果
            case 'set3DRotation':
                result = handleSet3DRotation(cmd.params);
                break;
            case 'set3DDepth':
                result = handleSet3DDepth(cmd.params);
                break;
            case 'set3DMaterial':
                result = handleSet3DMaterial(cmd.params);
                break;
            case 'create3DText':
                result = handleCreate3DText(cmd.params);
                break;

            // PPT 图表操作
            case 'insertPptChart':
                result = handleInsertPptChart(cmd.params);
                break;
            case 'setPptChartData':
                result = handleSetPptChartData(cmd.params);
                break;
            case 'setPptChartStyle':
                result = handleSetPptChartStyle(cmd.params);
                break;

            // PPT 动画效果
            case 'addAnimation':
                result = handleAddAnimation(cmd.params);
                break;
            case 'removeAnimation':
                result = handleRemoveAnimation(cmd.params);
                break;
            case 'getAnimations':
                result = handleGetAnimations(cmd.params);
                break;
            case 'setAnimationOrder':
                result = handleSetAnimationOrder(cmd.params);
                break;

            // PPT 切换效果
            case 'setSlideTransition':
                result = handleSetSlideTransition(cmd.params);
                break;
            case 'removeSlideTransition':
                result = handleRemoveSlideTransition(cmd.params);
                break;
            case 'applyTransitionToAll':
                result = handleApplyTransitionToAll(cmd.params);
                break;

            // PPT 主题/背景
            case 'setSlideBackground':
                result = handleSetSlideBackground(cmd.params);
                break;
            case 'setBackgroundColor':
                result = handleSetBackgroundColor(cmd.params);
                break;
            case 'setBackgroundImage':
                result = handleSetBackgroundImage(cmd.params);
                break;

            // PPT 超链接
            case 'addPptHyperlink':
                result = handleAddPptHyperlink(cmd.params);
                break;
            case 'removePptHyperlink':
                result = handleRemovePptHyperlink(cmd.params);
                break;

            // PPT 页眉页脚
            case 'setSlideNumber':
                result = handleSetSlideNumber(cmd.params);
                break;
            case 'setPptFooter':
                result = handleSetPptFooter(cmd.params);
                break;
            case 'setPptDateTime':
                result = handleSetPptDateTime(cmd.params);
                break;

            // PPT 查找替换
            case 'findPptText':
                result = handleFindPptText(cmd.params);
                break;
            case 'replacePptText':
                result = handleReplacePptText(cmd.params);
                break;

            // PPT 放映
            case 'startSlideShow':
                result = handleStartSlideShow(cmd.params);
                break;
            case 'endSlideShow':
                result = handleEndSlideShow(cmd.params);
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

// 获取所有打开的Word文档
function handleGetOpenDocuments(params) {
    try {
        var documents = [];
        for (var i = 1; i <= Application.Documents.Count; i++) {
            var doc = Application.Documents.Item(i);
            documents.push({
                name: doc.Name,
                path: doc.FullName,
                paragraphs: doc.Paragraphs.Count,
                active: doc.Name === Application.ActiveDocument.Name
            });
        }
        return { success: true, data: { documents: documents, count: documents.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 切换活动Word文档
function handleSwitchDocument(params) {
    try {
        var doc = Application.Documents.Item(params.name || params.index);
        doc.Activate();
        return { success: true, data: { name: doc.Name, path: doc.FullName } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 打开Word文档
function handleOpenDocument(params) {
    try {
        if (!params.path) return { success: false, error: '请提供文档路径' };
        var doc = Application.Documents.Open(params.path);
        return { success: true, data: { name: doc.Name, path: doc.FullName, paragraphs: doc.Paragraphs.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

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

        // WPS PPT布局常量映射
        var layoutMap = {
            'title': 1,           // ppLayoutTitle
            'title_content': 2,   // ppLayoutText
            'blank': 12,          // ppLayoutBlank
            'two_column': 3,      // ppLayoutTwoColumnText
            'comparison': 4       // ppLayoutComparison
        };

        var index = params.index || params.position || ppt.Slides.Count + 1;
        var layoutInput = params.layout || 'title_content';
        var layout = typeof layoutInput === 'string' ? (layoutMap[layoutInput] || 2) : layoutInput;
        var slide = ppt.Slides.Add(index, layout);

        // 设置标题（带防护）
        if (params.title && slide && slide.Shapes) {
            try {
                if (slide.Shapes.HasTitle && slide.Shapes.Title) {
                    slide.Shapes.Title.TextFrame.TextRange.Text = params.title;
                }
            } catch (titleErr) {
                // 标题设置失败不影响整体
            }
        }

        // 设置内容（带防护）
        if (params.content && slide && slide.Shapes) {
            try {
                // 遍历找到内容占位符
                for (var i = 1; i <= slide.Shapes.Count; i++) {
                    var shape = slide.Shapes.Item(i);
                    if (shape.Type === 14 && shape.PlaceholderFormat) { // msoPlaceholder
                        var phType = shape.PlaceholderFormat.Type;
                        if (phType === 2 || phType === 15) { // ppPlaceholderBody or ppPlaceholderObject
                            shape.TextFrame.TextRange.Text = params.content;
                            break;
                        }
                    }
                }
            } catch (contentErr) {
                // 内容设置失败不影响整体
            }
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

// ==================== PPT 演示文稿操作 Handlers ====================

// 创建新演示文稿
function handleCreatePresentation(params) {
    try {
        var ppt = Application.Presentations.Add();
        return { success: true, data: { name: ppt.Name, slideCount: ppt.Slides.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 打开演示文稿
function handleOpenPresentation(params) {
    try {
        var ppt = Application.Presentations.Open(params.path);
        return { success: true, data: { name: ppt.Name, path: ppt.FullName, slideCount: ppt.Slides.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 关闭演示文稿
function handleClosePresentation(params) {
    try {
        var ppt = params.name ? Application.Presentations.Item(params.name) : Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有找到演示文稿' };
        var name = ppt.Name;
        ppt.Close();
        return { success: true, data: { closed: name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取所有打开的演示文稿
function handleGetOpenPresentations(params) {
    try {
        var presentations = [];
        for (var i = 1; i <= Application.Presentations.Count; i++) {
            var ppt = Application.Presentations.Item(i);
            presentations.push({ name: ppt.Name, path: ppt.FullName, slideCount: ppt.Slides.Count });
        }
        return { success: true, data: { presentations: presentations, count: presentations.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 切换演示文稿
function handleSwitchPresentation(params) {
    try {
        var ppt = Application.Presentations.Item(params.name || params.index);
        ppt.Windows.Item(1).Activate();
        return { success: true, data: { name: ppt.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 幻灯片操作 Handlers ====================

// 删除幻灯片
function handleDeleteSlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex;
        ppt.Slides.Item(index).Delete();
        return { success: true, data: { deleted: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 复制幻灯片
function handleDuplicateSlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var newSlide = slide.Duplicate();
        return { success: true, data: { sourceIndex: index, newIndex: newSlide.Item(1).SlideIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移动幻灯片
function handleMoveSlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var fromIndex = params.from || params.fromIndex;
        var toIndex = params.to || params.toIndex;
        ppt.Slides.Item(fromIndex).MoveTo(toIndex);
        return { success: true, data: { from: fromIndex, to: toIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取幻灯片数量
function handleGetSlideCount(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        return { success: true, data: { count: ppt.Slides.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取幻灯片信息
function handleGetSlideInfo(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shapes = [];
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            shapes.push({ name: shape.Name, type: shape.Type, hasText: shape.HasTextFrame ? true : false });
        }
        return { success: true, data: { index: index, shapeCount: slide.Shapes.Count, shapes: shapes, layout: slide.Layout } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 切换到指定幻灯片
function handleSwitchSlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex;
        Application.ActiveWindow.View.GotoSlide(index);
        return { success: true, data: { currentSlide: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置幻灯片布局
function handleSetSlideLayout(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex || 1;
        var layoutMap = { 'title': 1, 'titleContent': 2, 'blank': 12, 'twoColumn': 4, 'comparison': 5, 'titleOnly': 11 };
        var layout = layoutMap[params.layout] || params.layout || 2;
        ppt.Slides.Item(index).Layout = layout;
        return { success: true, data: { slideIndex: index, layout: layout } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取幻灯片备注
function handleGetSlideNotes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var notes = '';
        try {
            notes = slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text;
        } catch (e) { notes = ''; }
        return { success: true, data: { slideIndex: index, notes: notes } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置幻灯片备注
function handleSetSlideNotes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.index || params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = params.notes || '';
        return { success: true, data: { slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 文本框操作 Handlers ====================

// 添加文本框
function handleAddTextBox(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || 200;
        var height = params.height || 50;
        var textBox = slide.Shapes.AddTextbox(1, left, top, width, height);
        if (params.text) {
            textBox.TextFrame.TextRange.Text = params.text;
        }
        if (params.fontSize) {
            textBox.TextFrame.TextRange.Font.Size = params.fontSize;
        }
        if (params.fontName) {
            textBox.TextFrame.TextRange.Font.Name = params.fontName;
        }
        return { success: true, data: { name: textBox.Name, slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除文本框
function handleDeleteTextBox(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        shape.Delete();
        return { success: true, data: { deleted: params.name || params.shapeIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取所有文本框
function handleGetTextBoxes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var textBoxes = [];
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            if (shape.HasTextFrame) {
                var text = '';
                try { text = shape.TextFrame.TextRange.Text; } catch (e) {}
                textBoxes.push({ name: shape.Name, index: i, text: text, left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height });
            }
        }
        return { success: true, data: { slideIndex: index, textBoxes: textBoxes, count: textBoxes.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置文本框内容
function handleSetTextBoxText(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        shape.TextFrame.TextRange.Text = params.text || '';
        return { success: true, data: { name: shape.Name, text: params.text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置文本框样式
function handleSetTextBoxStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        var tr = shape.TextFrame.TextRange;
        if (params.fontSize) tr.Font.Size = params.fontSize;
        if (params.fontName) tr.Font.Name = params.fontName;
        if (params.bold !== undefined) tr.Font.Bold = params.bold;
        if (params.italic !== undefined) tr.Font.Italic = params.italic;
        if (params.color) {
            var c = params.color.replace('#', '');
            tr.Font.Color.RGB = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }
        if (params.alignment) {
            var alignMap = { 'left': 1, 'center': 2, 'right': 3 };
            tr.ParagraphFormat.Alignment = alignMap[params.alignment] || 1;
        }
        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 标题操作 Handlers ====================

// 设置幻灯片标题
function handleSetSlideTitle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        if (slide.Shapes.HasTitle) {
            slide.Shapes.Title.TextFrame.TextRange.Text = params.title || '';
        }
        return { success: true, data: { slideIndex: index, title: params.title } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取幻灯片标题
function handleGetSlideTitle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var title = '';
        if (slide.Shapes.HasTitle) {
            title = slide.Shapes.Title.TextFrame.TextRange.Text;
        }
        return { success: true, data: { slideIndex: index, title: title } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置幻灯片副标题
function handleSetSlideSubtitle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            if (shape.PlaceholderFormat && shape.PlaceholderFormat.Type === 2) {
                shape.TextFrame.TextRange.Text = params.subtitle || '';
                return { success: true, data: { slideIndex: index, subtitle: params.subtitle } };
            }
        }
        return { success: false, error: '未找到副标题占位符' };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置幻灯片内容
function handleSetSlideContent(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            if (shape.PlaceholderFormat && shape.PlaceholderFormat.Type === 7) {
                shape.TextFrame.TextRange.Text = params.content || '';
                return { success: true, data: { slideIndex: index } };
            }
        }
        return { success: false, error: '未找到内容占位符' };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 形状操作 Handlers ====================

// 添加形状
function handleAddShape(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shapeTypeMap = { 'rectangle': 1, 'oval': 9, 'triangle': 7, 'diamond': 4, 'pentagon': 51, 'hexagon': 52, 'arrow': 13, 'star': 12, 'heart': 21, 'cloud': 179 };
        var shapeType = shapeTypeMap[params.type] || params.type || 1;
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || 100;
        var height = params.height || 100;
        var shape = slide.Shapes.AddShape(shapeType, left, top, width, height);
        if (params.text) {
            shape.TextFrame.TextRange.Text = params.text;
        }
        if (params.fillColor) {
            var c = params.fillColor.replace('#', '');
            shape.Fill.ForeColor.RGB = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }
        return { success: true, data: { name: shape.Name, slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除形状
function handleDeleteShape(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        shape.Delete();
        return { success: true, data: { deleted: params.name || params.shapeIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取所有形状
function handleGetShapes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shapes = [];
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            shapes.push({ name: shape.Name, index: i, type: shape.Type, left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height });
        }
        return { success: true, data: { slideIndex: index, shapes: shapes, count: shapes.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状样式
function handleSetShapeStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        if (params.fillColor) {
            var c = params.fillColor.replace('#', '');
            shape.Fill.ForeColor.RGB = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }
        if (params.lineColor) {
            var lc = params.lineColor.replace('#', '');
            shape.Line.ForeColor.RGB = parseInt(lc.substr(0, 2), 16) + parseInt(lc.substr(2, 2), 16) * 256 + parseInt(lc.substr(4, 2), 16) * 65536;
        }
        if (params.lineWidth) {
            shape.Line.Weight = params.lineWidth;
        }
        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状文本
function handleSetShapeText(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        shape.TextFrame.TextRange.Text = params.text || '';
        return { success: true, data: { name: shape.Name, text: params.text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状位置
function handleSetShapePosition(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        if (params.left !== undefined) shape.Left = params.left;
        if (params.top !== undefined) shape.Top = params.top;
        if (params.width !== undefined) shape.Width = params.width;
        if (params.height !== undefined) shape.Height = params.height;
        return { success: true, data: { name: shape.Name, left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 图片操作 Handlers ====================

// 插入图片
function handleInsertPptImage(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || -1;
        var height = params.height || -1;
        var pic = slide.Shapes.AddPicture(params.path, false, true, left, top, width, height);
        return { success: true, data: { name: pic.Name, path: params.path } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除图片
function handleDeletePptImage(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        shape.Delete();
        return { success: true, data: { deleted: params.name || params.shapeIndex } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置图片样式
function handleSetImageStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.name || params.shapeIndex);
        if (params.left !== undefined) shape.Left = params.left;
        if (params.top !== undefined) shape.Top = params.top;
        if (params.width !== undefined) shape.Width = params.width;
        if (params.height !== undefined) shape.Height = params.height;
        if (params.rotation !== undefined) shape.Rotation = params.rotation;
        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 表格操作 Handlers ====================

// 插入表格
function handleInsertPptTable(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var rows = params.rows || 3;
        var cols = params.cols || 3;
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || 400;
        var height = params.height || 200;
        var table = slide.Shapes.AddTable(rows, cols, left, top, width, height);
        return { success: true, data: { name: table.Name, rows: rows, cols: cols } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置表格单元格
function handleSetPptTableCell(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);

        // 修复：遍历所有shapes找到表格，而不是直接用tableIndex
        var shape = null;
        var tableCount = 0;
        var targetTableIndex = params.tableIndex || 1;

        if (params.tableName) {
            // 用名称查找
            shape = slide.Shapes.Item(params.tableName);
        } else {
            // 遍历找第N个表格
            for (var i = 1; i <= slide.Shapes.Count; i++) {
                var s = slide.Shapes.Item(i);
                if (s.HasTable) {
                    tableCount++;
                    if (tableCount === targetTableIndex) {
                        shape = s;
                        break;
                    }
                }
            }
        }

        if (!shape || !shape.HasTable) {
            return { success: false, error: '找不到表格，slideIndex=' + index + ', tableIndex=' + targetTableIndex + ', 共找到' + tableCount + '个表格' };
        }

        var cell = shape.Table.Cell(params.row, params.col);
        // 兼容 text 和 value 两种参数名
        var textValue = params.text || params.value || '';
        cell.Shape.TextFrame.TextRange.Text = textValue;
        return { success: true, data: { row: params.row, col: params.col, text: textValue } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取表格单元格
function handleGetPptTableCell(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.tableName || params.tableIndex);
        var cell = shape.Table.Cell(params.row, params.col);
        var value = cell.Shape.TextFrame.TextRange.Text;
        return { success: true, data: { row: params.row, col: params.col, value: value } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置表格样式
function handleSetPptTableStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.tableName || params.tableIndex);
        if (params.left !== undefined) shape.Left = params.left;
        if (params.top !== undefined) shape.Top = params.top;
        if (params.width !== undefined) shape.Width = params.width;
        if (params.height !== undefined) shape.Height = params.height;
        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置表格单元格样式（P0核心功能）
function handleSetPptTableCellStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);

        // 遍历找表格
        var shape = null;
        var tableCount = 0;
        var targetTableIndex = params.tableIndex || 1;

        if (params.tableName) {
            shape = slide.Shapes.Item(params.tableName);
        } else {
            for (var i = 1; i <= slide.Shapes.Count; i++) {
                var s = slide.Shapes.Item(i);
                if (s.HasTable) {
                    tableCount++;
                    if (tableCount === targetTableIndex) {
                        shape = s;
                        break;
                    }
                }
            }
        }

        if (!shape || !shape.HasTable) {
            return { success: false, error: '找不到表格' };
        }

        var cell = shape.Table.Cell(params.row, params.col);
        var cellShape = cell.Shape;

        // 背景色
        if (params.backgroundColor) {
            var bgColor = params.backgroundColor.replace('#', '');
            var r = parseInt(bgColor.substr(0, 2), 16);
            var g = parseInt(bgColor.substr(2, 2), 16);
            var b = parseInt(bgColor.substr(4, 2), 16);
            cellShape.Fill.Visible = true;
            cellShape.Fill.Solid();
            cellShape.Fill.ForeColor.RGB = r + g * 256 + b * 65536;
        }

        // 字体颜色
        if (params.fontColor) {
            var fc = params.fontColor.replace('#', '');
            var fr = parseInt(fc.substr(0, 2), 16);
            var fg = parseInt(fc.substr(2, 2), 16);
            var fb = parseInt(fc.substr(4, 2), 16);
            cellShape.TextFrame.TextRange.Font.Color.RGB = fr + fg * 256 + fb * 65536;
        }

        // 字号
        if (params.fontSize) {
            cellShape.TextFrame.TextRange.Font.Size = params.fontSize;
        }

        // 加粗
        if (params.bold !== undefined) {
            cellShape.TextFrame.TextRange.Font.Bold = params.bold;
        }

        // 对齐
        if (params.alignment) {
            var alignMap = { 'left': 1, 'center': 2, 'right': 3 };
            cellShape.TextFrame.TextRange.ParagraphFormat.Alignment = alignMap[params.alignment] || 2;
        }

        return { success: true, data: { row: params.row, col: params.col } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 批量设置表格行样式（P0核心功能 - 美化表头）
function handleSetPptTableRowStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);

        // 遍历找表格
        var shape = null;
        var tableCount = 0;
        var targetTableIndex = params.tableIndex || 1;

        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var s = slide.Shapes.Item(i);
            if (s.HasTable) {
                tableCount++;
                if (tableCount === targetTableIndex) {
                    shape = s;
                    break;
                }
            }
        }

        if (!shape || !shape.HasTable) {
            return { success: false, error: '找不到表格' };
        }

        var table = shape.Table;
        var row = params.row || 1;
        var colCount = table.Columns.Count;

        // 解析颜色
        var bgColor = null;
        if (params.backgroundColor) {
            var bg = params.backgroundColor.replace('#', '');
            var r = parseInt(bg.substr(0, 2), 16);
            var g = parseInt(bg.substr(2, 2), 16);
            var b = parseInt(bg.substr(4, 2), 16);
            bgColor = r + g * 256 + b * 65536;
        }

        var fontColor = null;
        if (params.fontColor) {
            var fc = params.fontColor.replace('#', '');
            var fr = parseInt(fc.substr(0, 2), 16);
            var fg = parseInt(fc.substr(2, 2), 16);
            var fb = parseInt(fc.substr(4, 2), 16);
            fontColor = fr + fg * 256 + fb * 65536;
        }

        // 遍历行中的所有单元格
        for (var col = 1; col <= colCount; col++) {
            var cell = table.Cell(row, col);
            var cellShape = cell.Shape;

            if (bgColor !== null) {
                cellShape.Fill.Visible = true;
                cellShape.Fill.Solid();
                cellShape.Fill.ForeColor.RGB = bgColor;
            }

            if (fontColor !== null) {
                cellShape.TextFrame.TextRange.Font.Color.RGB = fontColor;
            }

            if (params.fontSize) {
                cellShape.TextFrame.TextRange.Font.Size = params.fontSize;
            }

            if (params.bold !== undefined) {
                cellShape.TextFrame.TextRange.Font.Bold = params.bold;
            }

            if (params.alignment) {
                var alignMap = { 'left': 1, 'center': 2, 'right': 3 };
                cellShape.TextFrame.TextRange.ParagraphFormat.Alignment = alignMap[params.alignment] || 2;
            }
        }

        return { success: true, data: { row: row, cols: colCount } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 高级美化 Handlers ====================

// 设置形状阴影效果
function handleSetShapeShadow(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        var shadow = shape.Shadow;
        shadow.Visible = true;
        shadow.Type = params.type || 2; // msoShadow21 (外阴影)
        shadow.Blur = params.blur || 10;
        shadow.OffsetX = params.offsetX || 3;
        shadow.OffsetY = params.offsetY || 3;
        shadow.Transparency = params.transparency || 0.5;

        if (params.color) {
            var c = params.color.replace('#', '');
            var r = parseInt(c.substr(0, 2), 16);
            var g = parseInt(c.substr(2, 2), 16);
            var b = parseInt(c.substr(4, 2), 16);
            shadow.ForeColor.RGB = r + g * 256 + b * 65536;
        }

        return { success: true, data: { name: shape.Name, shadow: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置渐变背景
function handleSetBackgroundGradient(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var fill = slide.Background.Fill;
        fill.TwoColorGradient(params.direction || 1, 1); // 1=msoGradientHorizontal

        // 起始颜色
        if (params.color1) {
            var c1 = params.color1.replace('#', '');
            var r1 = parseInt(c1.substr(0, 2), 16);
            var g1 = parseInt(c1.substr(2, 2), 16);
            var b1 = parseInt(c1.substr(4, 2), 16);
            fill.ForeColor.RGB = r1 + g1 * 256 + b1 * 65536;
        }

        // 结束颜色
        if (params.color2) {
            var c2 = params.color2.replace('#', '');
            var r2 = parseInt(c2.substr(0, 2), 16);
            var g2 = parseInt(c2.substr(2, 2), 16);
            var b2 = parseInt(c2.substr(4, 2), 16);
            fill.BackColor.RGB = r2 + g2 * 256 + b2 * 65536;
        }

        return { success: true, data: { slideIndex: params.slideIndex || 1, gradient: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状渐变填充
function handleSetShapeGradient(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        var fill = shape.Fill;
        fill.TwoColorGradient(params.direction || 1, 1);

        if (params.color1) {
            var c1 = params.color1.replace('#', '');
            var r1 = parseInt(c1.substr(0, 2), 16);
            var g1 = parseInt(c1.substr(2, 2), 16);
            var b1 = parseInt(c1.substr(4, 2), 16);
            fill.ForeColor.RGB = r1 + g1 * 256 + b1 * 65536;
        }

        if (params.color2) {
            var c2 = params.color2.replace('#', '');
            var r2 = parseInt(c2.substr(0, 2), 16);
            var g2 = parseInt(c2.substr(2, 2), 16);
            var b2 = parseInt(c2.substr(4, 2), 16);
            fill.BackColor.RGB = r2 + g2 * 256 + b2 * 65536;
        }

        return { success: true, data: { name: shape.Name, gradient: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状边框
function handleSetShapeBorder(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        var line = shape.Line;

        if (params.visible === false) {
            line.Visible = false;
        } else {
            line.Visible = true;
            if (params.weight) line.Weight = params.weight;
            if (params.color) {
                var c = params.color.replace('#', '');
                var r = parseInt(c.substr(0, 2), 16);
                var g = parseInt(c.substr(2, 2), 16);
                var b = parseInt(c.substr(4, 2), 16);
                line.ForeColor.RGB = r + g * 256 + b * 65536;
            }
            if (params.dashStyle) {
                var dashMap = { 'solid': 1, 'dash': 4, 'dot': 2, 'dashDot': 5 };
                line.DashStyle = dashMap[params.dashStyle] || 1;
            }
        }

        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状透明度
function handleSetShapeTransparency(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        shape.Fill.Transparency = params.transparency || 0;

        return { success: true, data: { name: shape.Name, transparency: params.transparency } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状圆角
function handleSetShapeRoundness(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        // 圆角调整 (0-1之间) - 修复赋值语法
        if (shape.Adjustments && shape.Adjustments.Count > 0) {
            var adj = shape.Adjustments;
            adj(1) = params.roundness || 0.2;
        }

        return { success: true, data: { name: shape.Name, roundness: params.roundness } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 批量设置形状样式（一次性设置多个属性）
function handleSetShapeFullStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        // 填充色
        if (params.fillColor) {
            var fc = params.fillColor.replace('#', '');
            var fr = parseInt(fc.substr(0, 2), 16);
            var fg = parseInt(fc.substr(2, 2), 16);
            var fb = parseInt(fc.substr(4, 2), 16);
            shape.Fill.Visible = true;
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = fr + fg * 256 + fb * 65536;
        }

        // 透明度
        if (params.transparency !== undefined) {
            shape.Fill.Transparency = params.transparency;
        }

        // 边框
        if (params.borderColor) {
            var bc = params.borderColor.replace('#', '');
            var br = parseInt(bc.substr(0, 2), 16);
            var bg = parseInt(bc.substr(2, 2), 16);
            var bb = parseInt(bc.substr(4, 2), 16);
            shape.Line.Visible = true;
            shape.Line.ForeColor.RGB = br + bg * 256 + bb * 65536;
        }
        if (params.borderWeight) {
            shape.Line.Weight = params.borderWeight;
        }
        if (params.noBorder) {
            shape.Line.Visible = false;
        }

        // 阴影
        if (params.shadow) {
            shape.Shadow.Visible = true;
            shape.Shadow.Blur = params.shadowBlur || 8;
            shape.Shadow.OffsetX = params.shadowX || 3;
            shape.Shadow.OffsetY = params.shadowY || 3;
            shape.Shadow.Transparency = params.shadowTransparency || 0.6;
        }

        // 文字样式
        if (shape.HasTextFrame && shape.TextFrame.HasText) {
            var textRange = shape.TextFrame.TextRange;
            if (params.fontColor) {
                var tc = params.fontColor.replace('#', '');
                var tr = parseInt(tc.substr(0, 2), 16);
                var tg = parseInt(tc.substr(2, 2), 16);
                var tb = parseInt(tc.substr(4, 2), 16);
                textRange.Font.Color.RGB = tr + tg * 256 + tb * 65536;
            }
            if (params.fontSize) textRange.Font.Size = params.fontSize;
            if (params.bold !== undefined) textRange.Font.Bold = params.bold;
            if (params.fontName) textRange.Font.Name = params.fontName;
        }

        return { success: true, data: { name: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 对齐多个形状
function handleAlignShapes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var alignMap = {
            'left': 0, 'center': 1, 'right': 2,
            'top': 3, 'middle': 4, 'bottom': 5
        };
        var alignType = alignMap[params.alignment] || 1;

        // 选择指定的形状
        var shapeIndices = params.shapeIndices || [];
        if (shapeIndices.length >= 2) {
            var range = slide.Shapes.Range(shapeIndices);
            range.Align(alignType, false);
        }

        return { success: true, data: { aligned: shapeIndices.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 分布形状（等间距）
function handleDistributeShapes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var distributeType = params.direction === 'vertical' ? 1 : 0;

        var shapeIndices = params.shapeIndices || [];
        if (shapeIndices.length >= 3) {
            var range = slide.Shapes.Range(shapeIndices);
            range.Distribute(distributeType, false);
        }

        return { success: true, data: { distributed: shapeIndices.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 组合形状
function handleGroupShapes(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var shapeIndices = params.shapeIndices || [];
        if (shapeIndices.length >= 2) {
            var range = slide.Shapes.Range(shapeIndices);
            var group = range.Group();
            return { success: true, data: { name: group.Name } };
        }

        return { success: false, error: '至少需要2个形状' };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 复制形状
function handleDuplicateShape(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        var duplicate = shape.Duplicate();
        duplicate.Left = params.left || (shape.Left + 20);
        duplicate.Top = params.top || (shape.Top + 20);

        return { success: true, data: { name: duplicate.Name, original: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置形状层级
function handleSetShapeZOrder(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeIndex || params.shapeName);

        var orderMap = {
            'front': 0, 'back': 1, 'forward': 2, 'backward': 3
        };
        var order = orderMap[params.order] || 0;
        shape.ZOrder(order);

        return { success: true, data: { name: shape.Name, order: params.order } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加连接线
function handleAddConnector(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var connectorType = params.type === 'curved' ? 2 : 1; // 1=直线, 2=曲线
        var connector = slide.Shapes.AddConnector(connectorType,
            params.startX || 0, params.startY || 0,
            params.endX || 100, params.endY || 100);

        if (params.color) {
            var c = params.color.replace('#', '');
            var r = parseInt(c.substr(0, 2), 16);
            var g = parseInt(c.substr(2, 2), 16);
            var b = parseInt(c.substr(4, 2), 16);
            connector.Line.ForeColor.RGB = r + g * 256 + b * 65536;
        }
        if (params.weight) connector.Line.Weight = params.weight;

        return { success: true, data: { name: connector.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加箭头
function handleAddArrow(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var line = slide.Shapes.AddLine(
            params.startX || 0, params.startY || 0,
            params.endX || 100, params.endY || 100);

        line.Line.EndArrowheadStyle = 2; // 箭头
        line.Line.EndArrowheadLength = 2;
        line.Line.EndArrowheadWidth = 2;

        if (params.color) {
            var c = params.color.replace('#', '');
            var r = parseInt(c.substr(0, 2), 16);
            var g = parseInt(c.substr(2, 2), 16);
            var b = parseInt(c.substr(4, 2), 16);
            line.Line.ForeColor.RGB = r + g * 256 + b * 65536;
        }
        if (params.weight) line.Line.Weight = params.weight || 2;

        return { success: true, data: { name: line.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 图表操作 Handlers ====================

// 插入图表
function handleInsertPptChart(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var chartTypeMap = { 'column': 51, 'bar': 57, 'line': 4, 'pie': 5, 'area': 1, 'scatter': -4169, 'doughnut': -4120 };
        var chartType = chartTypeMap[params.type] || params.type || 51;
        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || 400;
        var height = params.height || 300;
        var chart = slide.Shapes.AddChart(chartType, left, top, width, height);
        return { success: true, data: { name: chart.Name, type: params.type } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置图表数据
function handleSetPptChartData(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.chartName || params.chartIndex);
        var chart = shape.Chart;
        var dataSheet = chart.ChartData.Workbook.Worksheets.Item(1);
        if (params.data) {
            for (var r = 0; r < params.data.length; r++) {
                for (var c = 0; c < params.data[r].length; c++) {
                    dataSheet.Cells.Item(r + 1, c + 1).Value2 = params.data[r][c];
                }
            }
        }
        return { success: true, data: { chartName: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置图表样式
function handleSetPptChartStyle(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.chartName || params.chartIndex);
        var chart = shape.Chart;
        if (params.title) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = params.title;
        }
        if (params.hasLegend !== undefined) {
            chart.HasLegend = params.hasLegend;
        }
        return { success: true, data: { chartName: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 专业美化 Handlers ====================

// 应用专业配色方案到幻灯片所有形状
function handleApplyColorScheme(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var schemes = {
            'business': { primary: '#1a365d', secondary: '#2d5a87', accent: '#ffc107', text: '#ffffff' },
            'tech': { primary: '#0d47a1', secondary: '#1976d2', accent: '#00bcd4', text: '#ffffff' },
            'nature': { primary: '#1b5e20', secondary: '#388e3c', accent: '#8bc34a', text: '#ffffff' },
            'elegant': { primary: '#37474f', secondary: '#546e7a', accent: '#ffab00', text: '#ffffff' },
            'vibrant': { primary: '#6a1b9a', secondary: '#9c27b0', accent: '#e91e63', text: '#ffffff' },
            'corporate': { primary: '#263238', secondary: '#455a64', accent: '#ff5722', text: '#ffffff' },
            'finance': { primary: '#0d47a1', secondary: '#1565c0', accent: '#4caf50', text: '#ffffff' }
        };
        var scheme = schemes[params.scheme] || schemes['business'];

        function hexToRgb(hex) {
            var c = hex.replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var count = 0;
        // 遍历所有形状应用配色
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            try {
                // 根据形状类型应用不同颜色
                if (shape.HasTextFrame && shape.TextFrame.HasText) {
                    shape.TextFrame.TextRange.Font.Color.RGB = hexToRgb(scheme.text);
                }
                if (shape.Type === 1 || shape.Type === 6 || shape.Type === 13) { // AutoShape/FreeForm/Callout
                    shape.Fill.Solid();
                    shape.Fill.ForeColor.RGB = hexToRgb(count % 2 === 0 ? scheme.primary : scheme.secondary);
                    count++;
                }
            } catch (e) { /* 跳过不支持的形状 */ }
        }

        // 设置背景
        try {
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = hexToRgb(scheme.primary);
        } catch (e) {}

        return { success: true, data: { scheme: params.scheme, shapesProcessed: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 一键美化幻灯片
function handleAutoBeautifySlide(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var style = params.style || 'business';
        var schemes = {
            'business': { bg1: '#1a365d', bg2: '#2d5a87', accent: '#ffc107' },
            'tech': { bg1: '#0d47a1', bg2: '#1976d2', accent: '#00bcd4' },
            'elegant': { bg1: '#37474f', bg2: '#546e7a', accent: '#ffab00' }
        };
        var scheme = schemes[style] || schemes['business'];

        // 设置渐变背景
        var fill = slide.Background.Fill;
        fill.TwoColorGradient(1, 1);
        var c1 = scheme.bg1.replace('#', '');
        var r1 = parseInt(c1.substr(0, 2), 16);
        var g1 = parseInt(c1.substr(2, 2), 16);
        var b1 = parseInt(c1.substr(4, 2), 16);
        fill.ForeColor.RGB = r1 + g1 * 256 + b1 * 65536;

        var c2 = scheme.bg2.replace('#', '');
        var r2 = parseInt(c2.substr(0, 2), 16);
        var g2 = parseInt(c2.substr(2, 2), 16);
        var b2 = parseInt(c2.substr(4, 2), 16);
        fill.BackColor.RGB = r2 + g2 * 256 + b2 * 65536;

        // 添加装饰条
        var accentBar = slide.Shapes.AddShape(1, 50, slide.Master.Height - 80, slide.Master.Width - 100, 6);
        var ca = scheme.accent.replace('#', '');
        var ra = parseInt(ca.substr(0, 2), 16);
        var ga = parseInt(ca.substr(2, 2), 16);
        var ba = parseInt(ca.substr(4, 2), 16);
        accentBar.Fill.Solid();
        accentBar.Fill.ForeColor.RGB = ra + ga * 256 + ba * 65536;
        accentBar.Line.Visible = false;

        return { success: true, data: { style: style, beautified: true } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建KPI卡片组
function handleCreateKpiCards(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var cards = params.cards || [];
        var startX = params.startX || 50;
        var startY = params.startY || 350;
        var cardWidth = params.cardWidth || 150;
        var cardHeight = params.cardHeight || 80;
        var gap = params.gap || 20;

        var colors = ['#28a745', '#17a2b8', '#ffc107', '#dc3545', '#6f42c1', '#fd7e14'];
        var createdCards = [];

        for (var i = 0; i < cards.length; i++) {
            var card = cards[i];
            var x = startX + i * (cardWidth + gap);
            var colorIdx = i % colors.length;
            var color = card.color || colors[colorIdx];

            // 创建卡片形状
            var shape = slide.Shapes.AddShape(5, x, startY, cardWidth, cardHeight); // 5 = 圆角矩形

            // 设置渐变填充
            var c = color.replace('#', '');
            var r = parseInt(c.substr(0, 2), 16);
            var g = parseInt(c.substr(2, 2), 16);
            var b = parseInt(c.substr(4, 2), 16);
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = r + g * 256 + b * 65536;

            // 去除边框
            shape.Line.Visible = false;

            // 添加阴影
            shape.Shadow.Visible = true;
            shape.Shadow.Blur = 10;
            shape.Shadow.OffsetX = 3;
            shape.Shadow.OffsetY = 3;
            shape.Shadow.Transparency = 0.6;

            // 设置文本
            shape.TextFrame.TextRange.Text = (card.title || '') + '\n' + (card.value || '');
            shape.TextFrame.TextRange.Font.Color.RGB = 16777215; // 白色
            shape.TextFrame.TextRange.Font.Size = 14;
            shape.TextFrame.TextRange.Font.Bold = true;
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2; // 居中

            createdCards.push(shape.Name);
        }

        return { success: true, data: { count: createdCards.length, cards: createdCards } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建专业表格（带样式）
function handleCreateStyledTable(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var rows = params.rows || 5;
        var cols = params.cols || 4;
        var left = params.left || 50;
        var top = params.top || 150;
        var width = params.width || 620;
        var height = params.height || 200;

        // 创建表格
        var table = slide.Shapes.AddTable(rows, cols, left, top, width, height);
        var tbl = table.Table;

        // 样式配置
        var style = params.style || 'business';
        var styles = {
            'business': { header: '#1a365d', headerText: '#ffffff', oddRow: '#f8f9fa', evenRow: '#ffffff' },
            'tech': { header: '#0d47a1', headerText: '#ffffff', oddRow: '#e3f2fd', evenRow: '#ffffff' },
            'elegant': { header: '#37474f', headerText: '#ffffff', oddRow: '#eceff1', evenRow: '#ffffff' }
        };
        var s = styles[style] || styles['business'];

        // 设置表头样式
        for (var c = 1; c <= cols; c++) {
            var headerCell = tbl.Cell(1, c);
            var hc = s.header.replace('#', '');
            var hr = parseInt(hc.substr(0, 2), 16);
            var hg = parseInt(hc.substr(2, 2), 16);
            var hb = parseInt(hc.substr(4, 2), 16);
            headerCell.Shape.Fill.Solid();
            headerCell.Shape.Fill.ForeColor.RGB = hr + hg * 256 + hb * 65536;
            headerCell.Shape.TextFrame.TextRange.Font.Color.RGB = 16777215;
            headerCell.Shape.TextFrame.TextRange.Font.Bold = true;
            headerCell.Shape.TextFrame.TextRange.Font.Size = 12;
        }

        // 设置数据行斑马纹
        for (var r = 2; r <= rows; r++) {
            var rowColor = (r % 2 === 0) ? s.evenRow : s.oddRow;
            var rc = rowColor.replace('#', '');
            var rr = parseInt(rc.substr(0, 2), 16);
            var rg = parseInt(rc.substr(2, 2), 16);
            var rb = parseInt(rc.substr(4, 2), 16);
            var rgb = rr + rg * 256 + rb * 65536;

            for (var c = 1; c <= cols; c++) {
                var cell = tbl.Cell(r, c);
                cell.Shape.Fill.Solid();
                cell.Shape.Fill.ForeColor.RGB = rgb;
                cell.Shape.TextFrame.TextRange.Font.Size = 11;
            }
        }

        // 填充数据
        if (params.data) {
            for (var r = 0; r < params.data.length && r < rows; r++) {
                for (var c = 0; c < params.data[r].length && c < cols; c++) {
                    tbl.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = String(params.data[r][c]);
                }
            }
        }

        return { success: true, data: { name: table.Name, rows: rows, cols: cols } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加标题装饰条
function handleAddTitleDecoration(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var left = params.left || 50;
        var top = params.top || 130;
        var width = params.width || 200;
        var height = params.height || 5;
        var color = params.color || '#1a365d';

        var bar = slide.Shapes.AddShape(1, left, top, width, height);
        var c = color.replace('#', '');
        var r = parseInt(c.substr(0, 2), 16);
        var g = parseInt(c.substr(2, 2), 16);
        var b = parseInt(c.substr(4, 2), 16);
        bar.Fill.Solid();
        bar.Fill.ForeColor.RGB = r + g * 256 + b * 65536;
        bar.Line.Visible = false;

        return { success: true, data: { name: bar.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加页码指示器
function handleAddPageIndicator(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var total = ppt.Slides.Count;
        var current = params.slideIndex || 1;
        var text = current + ' / ' + total;

        var left = params.left || 880;
        var top = params.top || 500;
        var box = slide.Shapes.AddTextbox(1, left, top, 80, 30);
        box.TextFrame.TextRange.Text = text;
        box.TextFrame.TextRange.Font.Size = 12;
        box.TextFrame.TextRange.Font.Color.RGB = params.dark ? 0 : 16777215;
        box.TextFrame.TextRange.ParagraphFormat.Alignment = 3; // 右对齐

        return { success: true, data: { name: box.Name, page: text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 批量美化所有幻灯片
function handleBeautifyAllSlides(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        var style = params.style || 'business';
        var schemes = {
            'business': { bg1: '#1a365d', bg2: '#2d5a87', accent: '#ffc107', light: '#f8f9fa' },
            'tech': { bg1: '#0d47a1', bg2: '#1976d2', accent: '#00bcd4', light: '#e3f2fd' },
            'elegant': { bg1: '#37474f', bg2: '#546e7a', accent: '#ffab00', light: '#eceff1' }
        };
        var scheme = schemes[style] || schemes['business'];

        var count = ppt.Slides.Count;
        for (var i = 1; i <= count; i++) {
            var slide = ppt.Slides.Item(i);
            var fill = slide.Background.Fill;

            // 首页和末页用深色渐变，中间页用浅色
            if (i === 1 || i === count) {
                fill.TwoColorGradient(1, 1);
                var c1 = scheme.bg1.replace('#', '');
                fill.ForeColor.RGB = parseInt(c1.substr(0, 2), 16) + parseInt(c1.substr(2, 2), 16) * 256 + parseInt(c1.substr(4, 2), 16) * 65536;
                var c2 = scheme.bg2.replace('#', '');
                fill.BackColor.RGB = parseInt(c2.substr(0, 2), 16) + parseInt(c2.substr(2, 2), 16) * 256 + parseInt(c2.substr(4, 2), 16) * 65536;
            } else {
                fill.Solid();
                var cl = scheme.light.replace('#', '');
                fill.ForeColor.RGB = parseInt(cl.substr(0, 2), 16) + parseInt(cl.substr(2, 2), 16) * 256 + parseInt(cl.substr(4, 2), 16) * 65536;
            }
        }

        return { success: true, data: { style: style, slidesBeautified: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 动画效果 Handlers ====================

// 添加动画
function handleAddAnimation(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex);
        var effectTypeMap = { 'appear': 1, 'fade': 10, 'fly': 2, 'float': 14, 'split': 13, 'wipe': 22, 'zoom': 23, 'swivel': 15, 'bounce': 26, 'spin': 49 };
        var effectType = effectTypeMap[params.effect] || params.effect || 10;
        var effect = slide.TimeLine.MainSequence.AddEffect(shape, effectType, 0, 1);
        if (params.duration) effect.Timing.Duration = params.duration;
        if (params.delay) effect.Timing.TriggerDelayTime = params.delay;
        return { success: true, data: { shapeName: shape.Name, effect: params.effect } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 删除动画
function handleRemoveAnimation(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var seq = slide.TimeLine.MainSequence;
        if (params.index) {
            seq.Item(params.index).Delete();
        } else {
            while (seq.Count > 0) {
                seq.Item(1).Delete();
            }
        }
        return { success: true, data: { slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取动画列表
function handleGetAnimations(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var animations = [];
        var seq = slide.TimeLine.MainSequence;
        for (var i = 1; i <= seq.Count; i++) {
            var effect = seq.Item(i);
            animations.push({ index: i, shapeName: effect.Shape.Name, effectType: effect.EffectType, duration: effect.Timing.Duration });
        }
        return { success: true, data: { slideIndex: index, animations: animations, count: animations.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置动画顺序
function handleSetAnimationOrder(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var seq = slide.TimeLine.MainSequence;
        var effect = seq.Item(params.from);
        effect.MoveTo(params.to);
        return { success: true, data: { from: params.from, to: params.to } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 切换效果 Handlers ====================

// 设置幻灯片切换效果
function handleSetSlideTransition(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var transitionMap = { 'none': 0, 'fade': 1, 'push': 2, 'wipe': 3, 'split': 4, 'reveal': 5, 'random': 6, 'cover': 7, 'uncover': 8, 'flash': 9 };
        var transition = transitionMap[params.transition] || params.transition || 1;
        slide.SlideShowTransition.EntryEffect = transition;
        if (params.duration) slide.SlideShowTransition.Duration = params.duration;
        if (params.advanceOnClick !== undefined) slide.SlideShowTransition.AdvanceOnClick = params.advanceOnClick;
        if (params.advanceTime) {
            slide.SlideShowTransition.AdvanceOnTime = true;
            slide.SlideShowTransition.AdvanceTime = params.advanceTime;
        }
        return { success: true, data: { slideIndex: index, transition: params.transition } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移除幻灯片切换效果
function handleRemoveSlideTransition(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        slide.SlideShowTransition.EntryEffect = 0;
        return { success: true, data: { slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 应用切换效果到所有幻灯片
function handleApplyTransitionToAll(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var transitionMap = { 'none': 0, 'fade': 1, 'push': 2, 'wipe': 3, 'split': 4, 'reveal': 5, 'random': 6 };
        var transition = transitionMap[params.transition] || params.transition || 1;
        for (var i = 1; i <= ppt.Slides.Count; i++) {
            ppt.Slides.Item(i).SlideShowTransition.EntryEffect = transition;
            if (params.duration) ppt.Slides.Item(i).SlideShowTransition.Duration = params.duration;
        }
        return { success: true, data: { transition: params.transition, appliedTo: ppt.Slides.Count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 主题/背景 Handlers ====================

// 设置幻灯片背景
function handleSetSlideBackground(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        if (params.color) {
            var c = params.color.replace('#', '');
            slide.FollowMasterBackground = false;
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }
        return { success: true, data: { slideIndex: index } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置背景颜色
function handleSetBackgroundColor(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var c = params.color.replace('#', '');
        var rgb = parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        if (params.applyToAll) {
            for (var i = 1; i <= ppt.Slides.Count; i++) {
                var slide = ppt.Slides.Item(i);
                slide.FollowMasterBackground = false;
                slide.Background.Fill.Solid();
                slide.Background.Fill.ForeColor.RGB = rgb;
            }
            return { success: true, data: { color: params.color, appliedTo: ppt.Slides.Count } };
        } else {
            var index = params.slideIndex || 1;
            var slide = ppt.Slides.Item(index);
            slide.FollowMasterBackground = false;
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = rgb;
            return { success: true, data: { slideIndex: index, color: params.color } };
        }
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置背景图片
function handleSetBackgroundImage(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        slide.FollowMasterBackground = false;
        slide.Background.Fill.UserPicture(params.path);
        return { success: true, data: { slideIndex: index, path: params.path } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 超链接 Handlers ====================

// 添加超链接
function handleAddPptHyperlink(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex);
        var actionSettings = shape.ActionSettings.Item(1);
        actionSettings.Hyperlink.Address = params.address || '';
        if (params.subAddress) actionSettings.Hyperlink.SubAddress = params.subAddress;
        return { success: true, data: { shapeName: shape.Name, address: params.address } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 移除超链接
function handleRemovePptHyperlink(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var index = params.slideIndex || 1;
        var slide = ppt.Slides.Item(index);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex);
        shape.ActionSettings.Item(1).Hyperlink.Address = '';
        return { success: true, data: { shapeName: shape.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 页眉页脚 Handlers ====================

// 设置幻灯片编号
function handleSetSlideNumber(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        if (params.visible !== undefined) {
            ppt.SlideMaster.HeadersFooters.SlideNumber.Visible = params.visible;
        }
        return { success: true, data: { visible: params.visible } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置页脚
function handleSetPptFooter(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        ppt.SlideMaster.HeadersFooters.Footer.Visible = true;
        ppt.SlideMaster.HeadersFooters.Footer.Text = params.text || '';
        return { success: true, data: { text: params.text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置日期时间
function handleSetPptDateTime(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        ppt.SlideMaster.HeadersFooters.DateAndTime.Visible = params.visible !== false;
        if (params.useFixed) {
            ppt.SlideMaster.HeadersFooters.DateAndTime.UseFormat = false;
            ppt.SlideMaster.HeadersFooters.DateAndTime.Text = params.text || '';
        }
        return { success: true, data: { visible: params.visible !== false } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 查找替换 Handlers ====================

// 查找文本
function handleFindPptText(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var results = [];
        for (var i = 1; i <= ppt.Slides.Count; i++) {
            var slide = ppt.Slides.Item(i);
            for (var j = 1; j <= slide.Shapes.Count; j++) {
                var shape = slide.Shapes.Item(j);
                try {
                    if (shape.HasTextFrame && shape.TextFrame.HasText) {
                        var text = shape.TextFrame.TextRange.Text;
                        if (text.indexOf(params.text) !== -1) {
                            results.push({ slideIndex: i, shapeName: shape.Name, text: text });
                        }
                    }
                } catch (e) {}
            }
        }
        return { success: true, data: { searchText: params.text, results: results, count: results.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 替换文本
function handleReplacePptText(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var count = 0;
        for (var i = 1; i <= ppt.Slides.Count; i++) {
            var slide = ppt.Slides.Item(i);
            for (var j = 1; j <= slide.Shapes.Count; j++) {
                var shape = slide.Shapes.Item(j);
                try {
                    if (shape.HasTextFrame && shape.TextFrame.HasText) {
                        var text = shape.TextFrame.TextRange.Text;
                        if (text.indexOf(params.findText) !== -1) {
                            shape.TextFrame.TextRange.Text = text.split(params.findText).join(params.replaceText || '');
                            count++;
                        }
                    }
                } catch (e) {}
            }
        }
        return { success: true, data: { findText: params.findText, replaceText: params.replaceText, count: count } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ==================== PPT 放映 Handlers ====================

// 开始放映
function handleStartSlideShow(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var startSlide = params.startSlide || 1;
        ppt.SlideShowSettings.StartingSlide = startSlide;
        ppt.SlideShowSettings.Run();
        return { success: true, data: { startSlide: startSlide } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 结束放映
function handleEndSlideShow(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        ppt.SlideShowWindow.View.Exit();
        return { success: true, data: { ended: true } };
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

// ==================== PPT 高端能力 - 6大类 ====================

// ========== 1. 数据可视化组件 ==========

// 创建进度条
function handleCreateProgressBar(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var left = params.left || 100;
        var top = params.top || 300;
        var width = params.width || 400;
        var height = params.height || 30;
        var progress = Math.min(Math.max(params.progress || 0.5, 0), 1);
        var label = params.label || '';

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        // 背景条
        var bgBar = slide.Shapes.AddShape(5, left, top, width, height);
        bgBar.Fill.Solid();
        bgBar.Fill.ForeColor.RGB = hexToRgb(params.bgColor || '#e0e0e0');
        bgBar.Line.Visible = false;

        // 进度条
        var progressWidth = width * progress;
        var progressBar = slide.Shapes.AddShape(5, left, top, progressWidth, height);
        progressBar.Fill.Solid();
        progressBar.Fill.ForeColor.RGB = hexToRgb(params.color || '#28a745');
        progressBar.Line.Visible = false;

        // 百分比文本
        var percentText = slide.Shapes.AddTextbox(1, left + width + 10, top, 60, height);
        percentText.TextFrame.TextRange.Text = Math.round(progress * 100) + '%';
        percentText.TextFrame.TextRange.Font.Size = 14;
        percentText.TextFrame.TextRange.Font.Bold = true;

        // 标签
        if (label) {
            var labelBox = slide.Shapes.AddTextbox(1, left, top - 25, width, 20);
            labelBox.TextFrame.TextRange.Text = label;
            labelBox.TextFrame.TextRange.Font.Size = 12;
        }

        return { success: true, data: { bgBar: bgBar.Name, progressBar: progressBar.Name, progress: progress } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建仪表盘
function handleCreateGauge(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var centerX = params.centerX || 360;
        var centerY = params.centerY || 300;
        var radius = params.radius || 100;
        var value = Math.min(Math.max(params.value || 0.5, 0), 1);
        var title = params.title || '';

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        // 外圆环背景
        var outerCircle = slide.Shapes.AddShape(9, centerX - radius, centerY - radius, radius * 2, radius * 2);
        outerCircle.Fill.Solid();
        outerCircle.Fill.ForeColor.RGB = hexToRgb('#e0e0e0');
        outerCircle.Line.Visible = false;

        // 内圆（白色遮挡）
        var innerRadius = radius * 0.7;
        var innerCircle = slide.Shapes.AddShape(9, centerX - innerRadius, centerY - innerRadius, innerRadius * 2, innerRadius * 2);
        innerCircle.Fill.Solid();
        innerCircle.Fill.ForeColor.RGB = hexToRgb('#ffffff');
        innerCircle.Line.Visible = false;

        // 数值显示
        var valueText = slide.Shapes.AddTextbox(1, centerX - 40, centerY - 20, 80, 40);
        valueText.TextFrame.TextRange.Text = Math.round(value * 100) + '%';
        valueText.TextFrame.TextRange.Font.Size = 24;
        valueText.TextFrame.TextRange.Font.Bold = true;
        valueText.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

        // 标题
        if (title) {
            var titleBox = slide.Shapes.AddTextbox(1, centerX - 60, centerY + radius + 10, 120, 25);
            titleBox.TextFrame.TextRange.Text = title;
            titleBox.TextFrame.TextRange.Font.Size = 14;
            titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
        }

        return { success: true, data: { gauge: outerCircle.Name, value: value } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建迷你图组（多个小指标）
function handleCreateMiniCharts(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var items = params.items || [];
        var startX = params.startX || 50;
        var startY = params.startY || 400;
        var itemWidth = params.itemWidth || 100;
        var gap = params.gap || 30;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var colors = ['#28a745', '#17a2b8', '#ffc107', '#dc3545', '#6f42c1'];
        var created = [];

        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var x = startX + i * (itemWidth + gap);
            var color = item.color || colors[i % colors.length];

            // 数值
            var valueBox = slide.Shapes.AddTextbox(1, x, startY, itemWidth, 30);
            valueBox.TextFrame.TextRange.Text = item.value || '0';
            valueBox.TextFrame.TextRange.Font.Size = 20;
            valueBox.TextFrame.TextRange.Font.Bold = true;
            valueBox.TextFrame.TextRange.Font.Color.RGB = hexToRgb(color);

            // 标签
            var labelBox = slide.Shapes.AddTextbox(1, x, startY + 30, itemWidth, 20);
            labelBox.TextFrame.TextRange.Text = item.label || '';
            labelBox.TextFrame.TextRange.Font.Size = 11;
            labelBox.TextFrame.TextRange.Font.Color.RGB = hexToRgb('#666666');

            // 趋势箭头
            if (item.trend) {
                var trendBox = slide.Shapes.AddTextbox(1, x + itemWidth - 30, startY, 30, 20);
                trendBox.TextFrame.TextRange.Text = item.trend === 'up' ? '↑' : '↓';
                trendBox.TextFrame.TextRange.Font.Size = 14;
                trendBox.TextFrame.TextRange.Font.Color.RGB = hexToRgb(item.trend === 'up' ? '#28a745' : '#dc3545');
            }

            created.push({ label: item.label, value: item.value });
        }

        return { success: true, data: { count: created.length, items: created } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建环形图（模拟）
function handleCreateDonutChart(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var centerX = params.centerX || 360;
        var centerY = params.centerY || 280;
        var radius = params.radius || 80;
        var value = Math.min(Math.max(params.value || 0.75, 0), 1);
        var title = params.title || '';
        var centerText = params.centerText || Math.round(value * 100) + '%';

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        // 外环
        var outerCircle = slide.Shapes.AddShape(9, centerX - radius, centerY - radius, radius * 2, radius * 2);
        outerCircle.Fill.Solid();
        outerCircle.Fill.ForeColor.RGB = hexToRgb(params.color || '#0d47a1');
        outerCircle.Line.Visible = false;

        // 内环（创建空心效果）
        var innerRadius = radius * 0.6;
        var innerCircle = slide.Shapes.AddShape(9, centerX - innerRadius, centerY - innerRadius, innerRadius * 2, innerRadius * 2);
        innerCircle.Fill.Solid();
        innerCircle.Fill.ForeColor.RGB = hexToRgb('#ffffff');
        innerCircle.Line.Visible = false;

        // 中心文本
        var textBox = slide.Shapes.AddTextbox(1, centerX - 40, centerY - 15, 80, 30);
        textBox.TextFrame.TextRange.Text = centerText;
        textBox.TextFrame.TextRange.Font.Size = 18;
        textBox.TextFrame.TextRange.Font.Bold = true;
        textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

        // 标题
        if (title) {
            var titleBox = slide.Shapes.AddTextbox(1, centerX - 60, centerY + radius + 15, 120, 25);
            titleBox.TextFrame.TextRange.Text = title;
            titleBox.TextFrame.TextRange.Font.Size = 12;
            titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
        }

        return { success: true, data: { donut: outerCircle.Name, value: value } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========== 2. 智能布局工具 ==========

// 自动排版（将形状整齐排列）
function handleAutoLayout(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var layout = params.layout || 'grid'; // grid, horizontal, vertical
        var margin = params.margin || 50;
        var gap = params.gap || 20;
        var slideWidth = 720;
        var slideHeight = 540;

        var shapes = [];
        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var s = slide.Shapes.Item(i);
            if (s.Type === 1 || s.Type === 6 || s.Type === 17) { // AutoShape, FreeForm, TextBox
                shapes.push(s);
            }
        }

        if (shapes.length === 0) {
            return { success: true, data: { message: '没有可排列的形状', count: 0 } };
        }

        if (layout === 'horizontal') {
            var totalWidth = 0;
            for (var j = 0; j < shapes.length; j++) {
                totalWidth += shapes[j].Width;
            }
            var availWidth = slideWidth - 2 * margin;
            var spacing = (availWidth - totalWidth) / (shapes.length + 1);
            var currentX = margin + spacing;
            for (var k = 0; k < shapes.length; k++) {
                shapes[k].Left = currentX;
                shapes[k].Top = (slideHeight - shapes[k].Height) / 2;
                currentX += shapes[k].Width + spacing;
            }
        } else if (layout === 'vertical') {
            var currentY = margin;
            for (var m = 0; m < shapes.length; m++) {
                shapes[m].Left = (slideWidth - shapes[m].Width) / 2;
                shapes[m].Top = currentY;
                currentY += shapes[m].Height + gap;
            }
        } else { // grid
            var cols = params.cols || Math.ceil(Math.sqrt(shapes.length));
            var rows = Math.ceil(shapes.length / cols);
            var cellWidth = (slideWidth - 2 * margin) / cols;
            var cellHeight = (slideHeight - 2 * margin) / rows;

            for (var n = 0; n < shapes.length; n++) {
                var col = n % cols;
                var row = Math.floor(n / cols);
                shapes[n].Left = margin + col * cellWidth + (cellWidth - shapes[n].Width) / 2;
                shapes[n].Top = margin + row * cellHeight + (cellHeight - shapes[n].Height) / 2;
            }
        }

        return { success: true, data: { layout: layout, count: shapes.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 智能等距分布
function handleSmartDistribute(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var direction = params.direction || 'horizontal';
        var shapeNames = params.shapes || [];

        var shapes = [];
        if (shapeNames.length > 0) {
            for (var i = 0; i < shapeNames.length; i++) {
                try {
                    shapes.push(slide.Shapes.Item(shapeNames[i]));
                } catch (e) {}
            }
        } else {
            for (var j = 1; j <= slide.Shapes.Count; j++) {
                var s = slide.Shapes.Item(j);
                if (s.Type === 1 || s.Type === 6 || s.Type === 17) {
                    shapes.push(s);
                }
            }
        }

        if (shapes.length < 2) {
            return { success: true, data: { message: '需要至少2个形状', count: shapes.length } };
        }

        // 按位置排序
        shapes.sort(function(a, b) {
            return direction === 'horizontal' ? a.Left - b.Left : a.Top - b.Top;
        });

        if (direction === 'horizontal') {
            var minX = shapes[0].Left;
            var maxX = shapes[shapes.length - 1].Left + shapes[shapes.length - 1].Width;
            var totalShapeWidth = 0;
            for (var k = 0; k < shapes.length; k++) {
                totalShapeWidth += shapes[k].Width;
            }
            var spacing = (maxX - minX - totalShapeWidth) / (shapes.length - 1);
            var currentX = minX;
            for (var m = 0; m < shapes.length; m++) {
                shapes[m].Left = currentX;
                currentX += shapes[m].Width + spacing;
            }
        } else {
            var minY = shapes[0].Top;
            var maxY = shapes[shapes.length - 1].Top + shapes[shapes.length - 1].Height;
            var totalShapeHeight = 0;
            for (var n = 0; n < shapes.length; n++) {
                totalShapeHeight += shapes[n].Height;
            }
            var spacingV = (maxY - minY - totalShapeHeight) / (shapes.length - 1);
            var currentY = minY;
            for (var p = 0; p < shapes.length; p++) {
                shapes[p].Top = currentY;
                currentY += shapes[p].Height + spacingV;
            }
        }

        return { success: true, data: { direction: direction, count: shapes.length } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建网格布局
function handleCreateGrid(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var rows = params.rows || 2;
        var cols = params.cols || 3;
        var margin = params.margin || 50;
        var gap = params.gap || 15;
        var slideWidth = 720;
        var slideHeight = 540;

        function hexToRgb(hex) {
            var c = (hex || '#f0f0f0').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var cellWidth = (slideWidth - 2 * margin - (cols - 1) * gap) / cols;
        var cellHeight = (slideHeight - 2 * margin - (rows - 1) * gap) / rows;

        var cells = [];
        for (var r = 0; r < rows; r++) {
            for (var c = 0; c < cols; c++) {
                var left = margin + c * (cellWidth + gap);
                var top = margin + r * (cellHeight + gap);
                var cell = slide.Shapes.AddShape(5, left, top, cellWidth, cellHeight);
                cell.Fill.Solid();
                cell.Fill.ForeColor.RGB = hexToRgb(params.cellColor || '#f8f9fa');
                cell.Line.ForeColor.RGB = hexToRgb(params.borderColor || '#dee2e6');
                cell.Line.Weight = 1;
                cells.push(cell.Name);
            }
        }

        return { success: true, data: { rows: rows, cols: cols, cells: cells } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========== 3. 高级动画组合 ==========

// 添加预设动画组（一键添加专业动画效果）
function handleAddAnimationPreset(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var preset = params.preset || 'fadeIn'; // fadeIn, flyIn, zoomIn, wipeIn, bounceIn
        var timeline = slide.TimeLine;
        var delay = 0;
        var delayIncrement = params.delayIncrement || 0.3;

        var presets = {
            'fadeIn': { effect: 10, duration: 0.5 },      // msoAnimEffectFade
            'flyIn': { effect: 2, duration: 0.5 },        // msoAnimEffectFly
            'zoomIn': { effect: 53, duration: 0.4 },      // msoAnimEffectGrowShrink
            'wipeIn': { effect: 22, duration: 0.5 },      // msoAnimEffectWipe
            'appear': { effect: 1, duration: 0 }          // msoAnimEffectAppear
        };

        var config = presets[preset] || presets['fadeIn'];
        var animatedCount = 0;

        for (var i = 1; i <= slide.Shapes.Count; i++) {
            var shape = slide.Shapes.Item(i);
            try {
                var effect = timeline.MainSequence.AddEffect(shape, config.effect, 0, 1);
                effect.Timing.Duration = config.duration;
                effect.Timing.TriggerDelayTime = delay;
                delay += delayIncrement;
                animatedCount++;
            } catch (e) { /* 跳过不支持动画的形状 */ }
        }

        return { success: true, data: { preset: preset, animatedShapes: animatedCount } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加强调动画
function handleAddEmphasisAnimation(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex || 1);

        var effectType = params.effect || 'pulse'; // pulse, spin, grow, teeter
        var effectMap = {
            'pulse': 63,    // msoAnimEffectFlashBulb
            'spin': 15,     // msoAnimEffectSpin
            'grow': 53,     // msoAnimEffectGrowShrink
            'teeter': 28    // msoAnimEffectTeeter
        };

        var timeline = slide.TimeLine;
        var effect = timeline.MainSequence.AddEffect(shape, effectMap[effectType] || 63, 0, 2); // 2 = 点击时
        effect.Timing.Duration = params.duration || 0.5;

        return { success: true, data: { shape: shape.Name, effect: effectType } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========== 4. 流程图/组织架构图 ==========

// 创建流程图
function handleCreateFlowChart(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var steps = params.steps || ['开始', '步骤1', '步骤2', '结束'];
        var direction = params.direction || 'horizontal';
        var startX = params.startX || 80;
        var startY = params.startY || 250;
        var boxWidth = params.boxWidth || 100;
        var boxHeight = params.boxHeight || 50;
        var gap = params.gap || 60;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var colors = {
            start: '#28a745',
            process: '#0d47a1',
            end: '#dc3545'
        };

        var shapes = [];
        for (var i = 0; i < steps.length; i++) {
            var x, y;
            if (direction === 'horizontal') {
                x = startX + i * (boxWidth + gap);
                y = startY;
            } else {
                x = startX;
                y = startY + i * (boxHeight + gap);
            }

            // 确定形状类型和颜色
            var shapeType = 5; // 圆角矩形
            var color = colors.process;
            if (i === 0) {
                shapeType = 9; // 椭圆
                color = colors.start;
            } else if (i === steps.length - 1) {
                shapeType = 9;
                color = colors.end;
            }

            var shape = slide.Shapes.AddShape(shapeType, x, y, boxWidth, boxHeight);
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = hexToRgb(color);
            shape.Line.Visible = false;
            shape.TextFrame.TextRange.Text = steps[i];
            shape.TextFrame.TextRange.Font.Color.RGB = 16777215;
            shape.TextFrame.TextRange.Font.Size = 12;
            shape.TextFrame.TextRange.Font.Bold = true;
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

            shapes.push(shape.Name);

            // 添加箭头连接（除了最后一个）
            if (i < steps.length - 1) {
                var arrowX1, arrowY1, arrowX2, arrowY2;
                if (direction === 'horizontal') {
                    arrowX1 = x + boxWidth;
                    arrowY1 = y + boxHeight / 2;
                    arrowX2 = x + boxWidth + gap;
                    arrowY2 = arrowY1;
                } else {
                    arrowX1 = x + boxWidth / 2;
                    arrowY1 = y + boxHeight;
                    arrowX2 = arrowX1;
                    arrowY2 = y + boxHeight + gap;
                }
                var arrow = slide.Shapes.AddLine(arrowX1, arrowY1, arrowX2, arrowY2);
                arrow.Line.EndArrowheadStyle = 2;
                arrow.Line.Weight = 2;
                arrow.Line.ForeColor.RGB = hexToRgb('#666666');
            }
        }

        return { success: true, data: { steps: steps.length, shapes: shapes } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建组织架构图
function handleCreateOrgChart(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var nodes = params.nodes || [
            { name: 'CEO', level: 0 },
            { name: '技术总监', level: 1 },
            { name: '市场总监', level: 1 },
            { name: '财务总监', level: 1 }
        ];

        var centerX = 360;
        var startY = params.startY || 80;
        var boxWidth = params.boxWidth || 100;
        var boxHeight = params.boxHeight || 40;
        var levelGap = params.levelGap || 80;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var levelColors = ['#1a365d', '#2d5a87', '#4a7ab0', '#6b9bd1'];

        // 按层级分组
        var levels = {};
        for (var i = 0; i < nodes.length; i++) {
            var level = nodes[i].level || 0;
            if (!levels[level]) levels[level] = [];
            levels[level].push(nodes[i]);
        }

        var shapes = [];
        var levelKeys = Object.keys(levels).sort();

        for (var l = 0; l < levelKeys.length; l++) {
            var levelNodes = levels[levelKeys[l]];
            var levelY = startY + l * levelGap;
            var totalWidth = levelNodes.length * boxWidth + (levelNodes.length - 1) * 30;
            var startX = centerX - totalWidth / 2;

            for (var n = 0; n < levelNodes.length; n++) {
                var x = startX + n * (boxWidth + 30);
                var shape = slide.Shapes.AddShape(5, x, levelY, boxWidth, boxHeight);
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = hexToRgb(levelColors[l % levelColors.length]);
                shape.Line.Visible = false;
                shape.TextFrame.TextRange.Text = levelNodes[n].name;
                shape.TextFrame.TextRange.Font.Color.RGB = 16777215;
                shape.TextFrame.TextRange.Font.Size = 11;
                shape.TextFrame.TextRange.Font.Bold = true;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

                // 添加阴影
                shape.Shadow.Visible = true;
                shape.Shadow.Blur = 5;
                shape.Shadow.OffsetX = 2;
                shape.Shadow.OffsetY = 2;
                shape.Shadow.Transparency = 0.7;

                shapes.push(shape.Name);
            }
        }

        return { success: true, data: { levels: levelKeys.length, totalNodes: nodes.length, shapes: shapes } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建时间轴
function handleCreateTimeline(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var events = params.events || [
            { date: '2020', title: '成立' },
            { date: '2021', title: '融资' },
            { date: '2022', title: '扩张' },
            { date: '2023', title: '上市' }
        ];

        var startX = params.startX || 80;
        var startY = params.startY || 280;
        var lineLength = params.lineLength || 560;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        // 主轴线
        var mainLine = slide.Shapes.AddLine(startX, startY, startX + lineLength, startY);
        mainLine.Line.Weight = 3;
        mainLine.Line.ForeColor.RGB = hexToRgb(params.lineColor || '#1a365d');

        var gap = lineLength / (events.length - 1 || 1);
        var shapes = [];

        for (var i = 0; i < events.length; i++) {
            var x = startX + i * gap;
            var above = i % 2 === 0;

            // 节点圆点
            var dot = slide.Shapes.AddShape(9, x - 8, startY - 8, 16, 16);
            dot.Fill.Solid();
            dot.Fill.ForeColor.RGB = hexToRgb(params.dotColor || '#0d47a1');
            dot.Line.Visible = false;

            // 连接线
            var lineY = above ? startY - 50 : startY + 50;
            var connector = slide.Shapes.AddLine(x, startY, x, lineY);
            connector.Line.Weight = 2;
            connector.Line.ForeColor.RGB = hexToRgb('#cccccc');

            // 日期标签
            var dateBox = slide.Shapes.AddTextbox(1, x - 30, above ? lineY - 45 : lineY + 5, 60, 20);
            dateBox.TextFrame.TextRange.Text = events[i].date;
            dateBox.TextFrame.TextRange.Font.Size = 12;
            dateBox.TextFrame.TextRange.Font.Bold = true;
            dateBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

            // 标题
            var titleBox = slide.Shapes.AddTextbox(1, x - 40, above ? lineY - 25 : lineY + 25, 80, 20);
            titleBox.TextFrame.TextRange.Text = events[i].title;
            titleBox.TextFrame.TextRange.Font.Size = 10;
            titleBox.TextFrame.TextRange.Font.Color.RGB = hexToRgb('#666666');
            titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

            shapes.push(dot.Name);
        }

        return { success: true, data: { events: events.length, shapes: shapes } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========== 5. 母版操作 ==========

// 获取母版信息
function handleGetSlideMaster(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        var masters = [];
        for (var i = 1; i <= ppt.SlideMaster.Shapes.Count; i++) {
            var shape = ppt.SlideMaster.Shapes.Item(i);
            masters.push({
                name: shape.Name,
                type: shape.Type,
                left: shape.Left,
                top: shape.Top,
                width: shape.Width,
                height: shape.Height
            });
        }

        return { success: true, data: { shapeCount: masters.length, shapes: masters } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置母版背景
function handleSetMasterBackground(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        function hexToRgb(hex) {
            var c = (hex || '#ffffff').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var master = ppt.SlideMaster;

        if (params.gradient) {
            master.Background.Fill.TwoColorGradient(1, 1);
            master.Background.Fill.GradientStops.Item(1).Color.RGB = hexToRgb(params.color1 || '#1a365d');
            master.Background.Fill.GradientStops.Item(2).Color.RGB = hexToRgb(params.color2 || '#2d5a87');
        } else {
            master.Background.Fill.Solid();
            master.Background.Fill.ForeColor.RGB = hexToRgb(params.color || '#ffffff');
        }

        return { success: true, data: { message: '母版背景已更新' } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 添加母版元素（如Logo占位）
function handleAddMasterElement(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };

        var master = ppt.SlideMaster;
        var type = params.type || 'textbox'; // textbox, shape, logo
        var left = params.left || 600;
        var top = params.top || 20;
        var width = params.width || 100;
        var height = params.height || 40;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        var shape;
        if (type === 'textbox') {
            shape = master.Shapes.AddTextbox(1, left, top, width, height);
            shape.TextFrame.TextRange.Text = params.text || 'Logo';
            shape.TextFrame.TextRange.Font.Size = params.fontSize || 14;
            shape.TextFrame.TextRange.Font.Bold = true;
        } else if (type === 'shape') {
            shape = master.Shapes.AddShape(params.shapeType || 5, left, top, width, height);
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = hexToRgb(params.color || '#1a365d');
            shape.Line.Visible = false;
        }

        return { success: true, data: { name: shape.Name, type: type } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========== 6. 3D效果 ==========

// 设置3D旋转效果
function handleSet3DRotation(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex || 1);

        var rotX = params.rotationX || 0;
        var rotY = params.rotationY || 0;
        var rotZ = params.rotationZ || 0;

        shape.ThreeD.RotationX = rotX;
        shape.ThreeD.RotationY = rotY;
        shape.ThreeD.RotationZ = rotZ;

        // 预设效果
        if (params.preset) {
            var presets = {
                'isometric': { x: 45, y: 45, z: 0 },
                'perspective': { x: 30, y: 30, z: 0 },
                'oblique': { x: 20, y: 60, z: 0 },
                'tiltLeft': { x: 0, y: -30, z: 0 },
                'tiltRight': { x: 0, y: 30, z: 0 }
            };
            var p = presets[params.preset];
            if (p) {
                shape.ThreeD.RotationX = p.x;
                shape.ThreeD.RotationY = p.y;
                shape.ThreeD.RotationZ = p.z;
            }
        }

        return { success: true, data: { shape: shape.Name, rotationX: shape.ThreeD.RotationX, rotationY: shape.ThreeD.RotationY } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置3D深度效果
function handleSet3DDepth(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex || 1);

        var depth = params.depth || 20;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        shape.ThreeD.Depth = depth;
        if (params.depthColor) {
            shape.ThreeD.ExtrusionColor.RGB = hexToRgb(params.depthColor);
        }

        // 设置光照
        if (params.lighting) {
            var lightingMap = {
                'bright': 1,
                'normal': 2,
                'dim': 3,
                'flat': 4
            };
            shape.ThreeD.PresetLighting = lightingMap[params.lighting] || 2;
        }

        return { success: true, data: { shape: shape.Name, depth: depth } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置3D材质效果
function handleSet3DMaterial(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);
        var shape = slide.Shapes.Item(params.shapeName || params.shapeIndex || 1);

        var material = params.material || 'plastic'; // plastic, metal, glass, matte

        var materialMap = {
            'matte': 1,
            'plastic': 2,
            'metal': 3,
            'wireFrame': 4,
            'glass': 5
        };

        shape.ThreeD.PresetMaterial = materialMap[material] || 2;

        return { success: true, data: { shape: shape.Name, material: material } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 创建3D文字效果
function handleCreate3DText(params) {
    try {
        var ppt = Application.ActivePresentation;
        if (!ppt) return { success: false, error: '没有打开的演示文稿' };
        var slide = ppt.Slides.Item(params.slideIndex || 1);

        var text = params.text || '3D文字';
        var left = params.left || 200;
        var top = params.top || 200;
        var width = params.width || 300;
        var height = params.height || 100;

        function hexToRgb(hex) {
            var c = (hex || '#1a365d').replace('#', '');
            return parseInt(c.substr(0, 2), 16) + parseInt(c.substr(2, 2), 16) * 256 + parseInt(c.substr(4, 2), 16) * 65536;
        }

        // 创建文本框
        var textBox = slide.Shapes.AddTextbox(1, left, top, width, height);
        textBox.TextFrame.TextRange.Text = text;
        textBox.TextFrame.TextRange.Font.Size = params.fontSize || 48;
        textBox.TextFrame.TextRange.Font.Bold = true;
        textBox.TextFrame.TextRange.Font.Color.RGB = hexToRgb(params.color || '#1a365d');
        textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;

        // 应用3D效果
        textBox.ThreeD.RotationX = params.rotationX || 15;
        textBox.ThreeD.RotationY = params.rotationY || 30;
        textBox.ThreeD.Depth = params.depth || 10;

        // 添加阴影
        textBox.Shadow.Visible = true;
        textBox.Shadow.Blur = 8;
        textBox.Shadow.OffsetX = 5;
        textBox.Shadow.OffsetY = 5;
        textBox.Shadow.Transparency = 0.5;

        return { success: true, data: { name: textBox.Name, text: text } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}
