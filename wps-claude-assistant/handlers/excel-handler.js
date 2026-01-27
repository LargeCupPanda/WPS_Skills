/**
 * WPS Excel Handler - Mac版 JavaScript API实现
 * 对标Windows PowerShell COM接口，提供一致的API
 * @author 老王
 */

// 获取工作表，支持索引或名称
function getSheet(wb, sheet) {
    if (!sheet) return wb.ActiveSheet;
    return wb.Sheets.Item(sheet);
}

// 获取当前工作簿信息
function getActiveWorkbook(params) {
    try {
        var app = Application;
        var wb = app.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            sheets.push(wb.Sheets.Item(i).Name);
        }
        return {
            success: true,
            data: { name: wb.Name, path: wb.FullName, sheetCount: wb.Sheets.Count, sheets: sheets }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 读取单元格值
function getCellValue(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = getSheet(wb, params.sheet);
        var cell = sheet.Cells.Item(params.row, params.col);
        return {
            success: true,
            data: { value: cell.Value2, text: cell.Text, formula: cell.Formula }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 写入单元格值
function setCellValue(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = getSheet(wb, params.sheet);
        sheet.Cells.Item(params.row, params.col).Value2 = params.value;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 读取范围数据
function getRangeData(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = getSheet(wb, params.sheet);
        var range = sheet.Range(params.range);
        var data = [];
        for (var r = 1; r <= range.Rows.Count; r++) {
            var row = [];
            for (var c = 1; c <= range.Columns.Count; c++) {
                row.push(range.Cells.Item(r, c).Value2);
            }
            data.push(row);
        }
        return { success: true, data: { data: data } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 写入范围数据
function setRangeData(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = getSheet(wb, params.sheet);
        var range = sheet.Range(params.range);
        for (var r = 0; r < params.data.length; r++) {
            for (var c = 0; c < params.data[r].length; c++) {
                range.Cells.Item(r + 1, c + 1).Value2 = params.data[r][c];
            }
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 设置公式
function setFormula(params) {
    try {
        var wb = Application.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = getSheet(wb, params.sheet);
        sheet.Cells.Item(params.row, params.col).Formula = params.formula;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 获取Excel上下文（表头、选中区域等）
function getExcelContext(params) {
    try {
        var app = Application;
        var wb = app.ActiveWorkbook;
        if (!wb) return { success: false, error: "No active workbook" };

        var sheet = app.ActiveSheet;
        var usedRange = sheet.UsedRange;
        var headers = [];

        if (usedRange.Rows.Count > 0) {
            var headerRow = usedRange.Rows.Item(1);
            var colCount = Math.min(headerRow.Columns.Count, 26);
            for (var i = 1; i <= colCount; i++) {
                headers.push({
                    column: String.fromCharCode(64 + i),
                    value: headerRow.Cells.Item(1, i).Value2
                });
            }
        }

        var sheets = [];
        for (var i = 1; i <= wb.Sheets.Count; i++) {
            sheets.push(wb.Sheets.Item(i).Name);
        }

        return {
            success: true,
            data: {
                workbookName: wb.Name,
                currentSheet: sheet.Name,
                allSheets: sheets,
                selectedCell: app.Selection.Address(),
                headers: headers,
                usedRange: usedRange.Address()
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 排序
function sortRange(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var keyCol = sheet.Range(params.keyColumn);
        var order = params.order === "desc" ? 2 : 1;
        range.Sort(keyCol, order);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 筛选
function autoFilter(params) {
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

// 创建图表
function createChart(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.dataRange);
        var chartTypes = { column: 51, bar: 57, line: 4, pie: 5, area: 1, scatter: -4169 };
        var chartType = chartTypes[params.chartType] || 51;

        var left = params.left || (range.Left + range.Width + 20);
        var top = params.top || range.Top;

        var chartObj = sheet.ChartObjects().Add(left, top, 400, 300);
        chartObj.Chart.SetSourceData(range);
        chartObj.Chart.ChartType = chartType;

        if (params.title) {
            chartObj.Chart.HasTitle = true;
            chartObj.Chart.ChartTitle.Text = params.title;
        }
        return { success: true, data: { chartName: chartObj.Name } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 去重
function removeDuplicates(params) {
    try {
        var sheet = Application.ActiveSheet;
        var range = sheet.Range(params.range);
        var cols = params.columns || [1];
        range.RemoveDuplicates(cols, 1);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getActiveWorkbook: getActiveWorkbook,
        getCellValue: getCellValue,
        setCellValue: setCellValue,
        getRangeData: getRangeData,
        setRangeData: setRangeData,
        setFormula: setFormula,
        getExcelContext: getExcelContext,
        sortRange: sortRange,
        autoFilter: autoFilter,
        createChart: createChart,
        removeDuplicates: removeDuplicates
    };
}
