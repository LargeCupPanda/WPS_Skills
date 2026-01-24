/**
 * WPS Claude 智能助手 - Excel 处理器
 *
 * 艹，终于轮到老王来写这个核心模块了
 * 这个SB模块负责处理所有Excel/表格相关的操作
 * 包括获取上下文、设置公式、诊断公式、数据清洗、读写范围等
 *
 * 注意：所有API都必须严格按照WPS SDK文档来写，别TM瞎编！
 *
 * @author 老王 (李四)
 * @date 2026-01-24
 */

// 创建日志记录器
var excelLogger = new Logger('ExcelHandler');

/**
 * Excel 处理器
 * 这个憨批对象封装了所有Excel操作的核心方法
 */
var ExcelHandler = {

    /**
     * 获取工作簿上下文
     * 返回当前工作簿的所有关键信息，让AI知道现在表格是什么状态
     *
     * @returns {object} 标准响应对象，包含上下文信息
     */
    getContext: function() {
        var startTime = Date.now();
        excelLogger.info('开始获取工作簿上下文');

        try {
            // 先检查有没有活动工作簿，没有就直接报错滚蛋
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                excelLogger.error('没有找到活动工作簿，你TM打开个Excel再说');
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: '请先打开一个Excel工作簿'
                });
            }

            // 获取活动工作表
            var activeSheet = workbook.ActiveSheet;
            if (!activeSheet) {
                excelLogger.error('没有活动工作表，这什么憨批情况');
                return Response.error('SHEET_NOT_FOUND', null, startTime, {
                    message: '没有活动的工作表'
                });
            }

            // 获取所有工作表名称列表
            var sheetsList = [];
            var sheetsCount = workbook.Sheets.Count;
            for (var i = 1; i <= sheetsCount; i++) {
                var sheet = workbook.Sheets.Item(i);
                sheetsList.push({
                    index: i,
                    name: sheet.Name,
                    visible: sheet.Visible
                });
            }

            // 获取当前选中的单元格/范围
            var selection = null;
            try {
                var sel = Application.Selection;
                if (sel) {
                    selection = {
                        address: sel.Address(),
                        rowCount: sel.Rows.Count,
                        columnCount: sel.Columns.Count
                    };
                    // 如果只选中一个单元格，获取它的值
                    if (sel.Rows.Count === 1 && sel.Columns.Count === 1) {
                        selection.value = sel.Value2;
                        selection.formula = sel.Formula;
                        selection.text = sel.Text;
                    }
                }
            } catch (selErr) {
                // 选区获取失败不影响整体，记个日志继续
                excelLogger.warn('获取选区信息失败', { error: selErr.message });
            }

            // 获取使用范围信息
            var usedRange = null;
            try {
                var used = activeSheet.UsedRange;
                if (used) {
                    usedRange = {
                        address: used.Address(),
                        rowCount: used.Rows.Count,
                        columnCount: used.Columns.Count,
                        firstRow: used.Row,
                        firstColumn: used.Column
                    };
                }
            } catch (usedErr) {
                excelLogger.warn('获取使用范围失败', { error: usedErr.message });
            }

            // 获取表头信息（假设第一行是表头）
            var headers = [];
            try {
                if (usedRange && usedRange.columnCount > 0) {
                    var headerRow = activeSheet.Range(
                        activeSheet.Cells.Item(1, 1),
                        activeSheet.Cells.Item(1, Math.min(usedRange.columnCount, 100)) // 最多读100列表头
                    );
                    var headerValues = headerRow.Value2;

                    // Value2返回的可能是单值或二维数组
                    if (headerValues) {
                        if (Array.isArray(headerValues)) {
                            // 二维数组，取第一行
                            var row = headerValues[0] || headerValues;
                            for (var h = 0; h < row.length; h++) {
                                if (row[h] !== null && row[h] !== undefined && row[h] !== '') {
                                    headers.push({
                                        column: h + 1,
                                        columnLetter: this._columnToLetter(h + 1),
                                        value: row[h]
                                    });
                                }
                            }
                        } else {
                            // 单值
                            headers.push({
                                column: 1,
                                columnLetter: 'A',
                                value: headerValues
                            });
                        }
                    }
                }
            } catch (headerErr) {
                excelLogger.warn('获取表头信息失败', { error: headerErr.message });
            }

            // 构建上下文对象
            var context = {
                workbook: {
                    name: workbook.Name,
                    path: workbook.Path || '',
                    fullName: workbook.FullName || workbook.Name,
                    saved: workbook.Saved,
                    readOnly: workbook.ReadOnly
                },
                activeSheet: {
                    name: activeSheet.Name,
                    index: activeSheet.Index
                },
                sheets: sheetsList,
                selection: selection,
                usedRange: usedRange,
                headers: headers
            };

            excelLogger.info('工作簿上下文获取成功', { sheetCount: sheetsList.length });
            return Response.success(context, null, startTime);

        } catch (err) {
            excelLogger.error('获取工作簿上下文失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 设置公式
     * 往指定单元格范围写入公式，并返回计算结果
     *
     * @param {object} params - 参数对象
     * @param {string} params.range - 单元格范围，如 "A1" 或 "B2:C5"
     * @param {string} params.formula - 公式字符串，必须以=开头
     * @returns {object} 标准响应对象
     */
    setFormula: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始设置公式', params);

        try {
            // 参数校验，别TM传空的进来
            var validation = Validator.checkRequired(params, ['range', 'formula']);
            if (!validation.valid) {
                excelLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 校验公式格式
            if (!Validator.isValidFormula(params.formula)) {
                excelLogger.error('公式格式不对，必须以=开头，憨批', { formula: params.formula });
                return Response.error('PARAM_FORMULA_INVALID', null, startTime, {
                    formula: params.formula,
                    message: '公式必须以=开头'
                });
            }

            // 校验范围格式
            if (!Validator.isValidRange(params.range)) {
                excelLogger.error('范围格式不对', { range: params.range });
                return Response.error('PARAM_RANGE_INVALID', null, startTime, {
                    range: params.range
                });
            }

            // 获取活动工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 检查是否只读
            if (workbook.ReadOnly) {
                excelLogger.error('工作簿是只读的，改不了');
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            // 获取目标范围
            var targetRange = sheet.Range(params.range);
            if (!targetRange) {
                excelLogger.error('找不到指定的单元格范围', { range: params.range });
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    range: params.range
                });
            }

            // 设置公式
            targetRange.Formula = params.formula;

            // 获取计算结果
            var resultValue = targetRange.Value2;
            var resultText = targetRange.Text;

            // 检测公式错误
            var hasError = false;
            var errorInfo = null;

            // 检查常见的公式错误值
            if (typeof resultText === 'string') {
                var errorPatterns = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#ERROR!'];
                for (var i = 0; i < errorPatterns.length; i++) {
                    if (resultText.indexOf(errorPatterns[i]) !== -1) {
                        hasError = true;
                        errorInfo = {
                            type: errorPatterns[i],
                            message: this._getFormulaErrorMessage(errorPatterns[i])
                        };
                        break;
                    }
                }
            }

            var result = {
                range: params.range,
                formula: params.formula,
                address: targetRange.Address(),
                calculatedValue: resultValue,
                displayText: resultText,
                hasError: hasError,
                errorInfo: errorInfo
            };

            if (hasError) {
                excelLogger.warn('公式设置成功但有计算错误', result);
            } else {
                excelLogger.info('公式设置成功', result);
            }

            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('设置公式失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 诊断公式错误
     * 分析指定单元格的公式错误，给出诊断和修复建议
     *
     * @param {object} params - 参数对象
     * @param {string} params.cell - 单元格地址，如 "A1"
     * @returns {object} 标准响应对象
     */
    diagnoseFormula: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始诊断公式错误', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['cell']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 获取目标单元格
            var cell = sheet.Range(params.cell);
            if (!cell) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    cell: params.cell
                });
            }

            // 获取单元格信息
            var formula = cell.Formula;
            var value = cell.Value2;
            var text = cell.Text;
            var hasFormula = cell.HasFormula;

            // 如果没有公式，直接返回
            if (!hasFormula) {
                return Response.success({
                    cell: params.cell,
                    hasFormula: false,
                    message: '该单元格没有公式',
                    value: value
                }, null, startTime);
            }

            // 分析公式错误
            var diagnosis = {
                cell: params.cell,
                address: cell.Address(),
                hasFormula: true,
                formula: formula,
                value: value,
                displayText: text,
                hasError: false,
                errorType: null,
                errorDescription: null,
                possibleCauses: [],
                suggestions: []
            };

            // 检测错误类型
            var errorPatterns = {
                '#DIV/0!': {
                    description: '除数为零错误',
                    causes: ['公式中的除数为0', '除数引用的单元格为空', '除数引用的单元格包含文本'],
                    suggestions: [
                        '使用IFERROR函数包装公式：=IFERROR(原公式, 0)',
                        '检查除数单元格是否有值',
                        '使用IF函数判断除数是否为0'
                    ]
                },
                '#VALUE!': {
                    description: '值类型错误',
                    causes: ['参数类型不正确', '文本被用于需要数字的地方', '数组公式使用不当'],
                    suggestions: [
                        '检查公式中的参数类型是否正确',
                        '使用VALUE()函数将文本转换为数字',
                        '使用TEXT()函数将数字转换为文本'
                    ]
                },
                '#REF!': {
                    description: '引用无效错误',
                    causes: ['引用的单元格已被删除', '工作表被删除', '复制公式导致引用超出范围'],
                    suggestions: [
                        '检查公式中的单元格引用是否有效',
                        '使用INDIRECT函数进行间接引用',
                        '重新输入正确的单元格引用'
                    ]
                },
                '#NAME?': {
                    description: '名称无法识别错误',
                    causes: ['函数名拼写错误', '使用了未定义的名称', '文本字符串缺少引号'],
                    suggestions: [
                        '检查函数名是否拼写正确',
                        '确保文本字符串用双引号括起来',
                        '检查是否使用了已定义的名称'
                    ]
                },
                '#NUM!': {
                    description: '数值错误',
                    causes: ['数值超出范围', '迭代计算无法收敛', '函数参数无效'],
                    suggestions: [
                        '检查数值是否在有效范围内',
                        '检查函数参数是否合理',
                        '使用ISNUMBER函数验证数值'
                    ]
                },
                '#N/A': {
                    description: '值不可用错误',
                    causes: ['VLOOKUP/HLOOKUP找不到匹配值', 'MATCH函数找不到匹配项', '数组维度不匹配'],
                    suggestions: [
                        '使用IFERROR或IFNA处理找不到的情况',
                        '检查查找值是否存在于查找范围中',
                        '确保数据格式一致（文本vs数字）'
                    ]
                },
                '#NULL!': {
                    description: '空交集错误',
                    causes: ['两个不相交区域使用了交集运算符', '范围引用格式错误'],
                    suggestions: [
                        '检查范围引用是否正确使用了冒号(:)或逗号(,)',
                        '确保引用的区域有交集'
                    ]
                }
            };

            // 检查是否有错误
            if (typeof text === 'string') {
                for (var errType in errorPatterns) {
                    if (text.indexOf(errType) !== -1) {
                        diagnosis.hasError = true;
                        diagnosis.errorType = errType;
                        diagnosis.errorDescription = errorPatterns[errType].description;
                        diagnosis.possibleCauses = errorPatterns[errType].causes;
                        diagnosis.suggestions = errorPatterns[errType].suggestions;
                        break;
                    }
                }
            }

            // 如果没有检测到错误，给出一般性建议
            if (!diagnosis.hasError) {
                diagnosis.message = '公式执行正常，没有检测到错误';
            }

            // 分析公式结构
            diagnosis.formulaAnalysis = this._analyzeFormulaStructure(formula);

            excelLogger.info('公式诊断完成', diagnosis);
            return Response.success(diagnosis, null, startTime);

        } catch (err) {
            excelLogger.error('公式诊断失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 数据清洗
     * 对指定范围的数据进行清洗操作
     *
     * @param {object} params - 参数对象
     * @param {string} params.range - 数据范围
     * @param {Array} params.operations - 清洗操作列表
     *   支持的操作：trim（去空格）、remove_duplicates（去重）、
     *              remove_empty_rows（删空行）、unify_date（统一日期格式）
     * @returns {object} 标准响应对象
     */
    cleanData: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始数据清洗', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['range', 'operations']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            if (!Array.isArray(params.operations) || params.operations.length === 0) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: 'operations必须是非空数组'
                });
            }

            // 获取活动工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 获取目标范围
            var targetRange = sheet.Range(params.range);
            if (!targetRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    range: params.range
                });
            }

            // 记录清洗结果
            var results = {
                range: params.range,
                address: targetRange.Address(),
                operations: [],
                totalCellsProcessed: targetRange.Rows.Count * targetRange.Columns.Count
            };

            // 执行各项清洗操作
            for (var i = 0; i < params.operations.length; i++) {
                var op = params.operations[i];
                var opResult = { operation: op, success: false, details: null };

                try {
                    switch (op) {
                        case 'trim':
                            opResult = this._cleanTrim(sheet, targetRange);
                            break;
                        case 'remove_duplicates':
                            opResult = this._cleanRemoveDuplicates(sheet, targetRange);
                            break;
                        case 'remove_empty_rows':
                            opResult = this._cleanRemoveEmptyRows(sheet, targetRange);
                            break;
                        case 'unify_date':
                            opResult = this._cleanUnifyDate(sheet, targetRange, params.dateFormat);
                            break;
                        default:
                            opResult = {
                                operation: op,
                                success: false,
                                details: '不支持的操作类型: ' + op
                            };
                    }
                } catch (opErr) {
                    opResult = {
                        operation: op,
                        success: false,
                        details: '操作执行失败: ' + opErr.message
                    };
                }

                results.operations.push(opResult);
            }

            excelLogger.info('数据清洗完成', results);
            return Response.success(results, null, startTime);

        } catch (err) {
            excelLogger.error('数据清洗失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 读取范围数据
     * 读取指定范围的数据，返回二维数组
     *
     * @param {object} params - 参数对象
     * @param {string} params.range - 数据范围，如 "A1:C10"
     * @returns {object} 标准响应对象
     */
    readRange: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始读取范围数据', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['range']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 获取目标范围
            var targetRange = sheet.Range(params.range);
            if (!targetRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    range: params.range
                });
            }

            // 读取数据
            var rawData = targetRange.Value2;
            var rowCount = targetRange.Rows.Count;
            var colCount = targetRange.Columns.Count;

            // 处理数据格式
            // Value2返回的可能是单值（1x1）、一维数组（1xN或Nx1）或二维数组
            var data = [];

            if (rowCount === 1 && colCount === 1) {
                // 单个单元格
                data = [[rawData]];
            } else if (Array.isArray(rawData)) {
                // 多个单元格
                if (Array.isArray(rawData[0])) {
                    // 已经是二维数组
                    data = rawData;
                } else {
                    // 一维数组，转换为二维
                    if (rowCount === 1) {
                        data = [rawData];
                    } else {
                        // 每个元素是一行
                        for (var r = 0; r < rawData.length; r++) {
                            data.push([rawData[r]]);
                        }
                    }
                }
            } else {
                data = [[rawData]];
            }

            var result = {
                range: params.range,
                address: targetRange.Address(),
                rowCount: rowCount,
                columnCount: colCount,
                data: data
            };

            excelLogger.info('范围数据读取成功', {
                range: params.range,
                rows: rowCount,
                cols: colCount
            });
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('读取范围数据失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 写入范围数据
     * 将二维数组数据写入指定范围
     *
     * @param {object} params - 参数对象
     * @param {string} params.range - 起始单元格或范围，如 "A1"
     * @param {Array} params.data - 二维数组数据
     * @returns {object} 标准响应对象
     */
    writeRange: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始写入范围数据', { range: params.range });

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['range', 'data']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            if (!Array.isArray(params.data)) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: 'data必须是数组'
                });
            }

            // 获取活动工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 确定数据维度
            var data = params.data;
            var rowCount = data.length;
            if (rowCount === 0) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '数据不能为空'
                });
            }

            // 确保是二维数组
            if (!Array.isArray(data[0])) {
                // 如果是一维数组，转换为二维
                data = [data];
                rowCount = 1;
            }

            var colCount = data[0].length;
            for (var i = 1; i < data.length; i++) {
                if (data[i].length > colCount) {
                    colCount = data[i].length;
                }
            }

            // 获取起始单元格
            var startCell = sheet.Range(params.range);
            if (!startCell) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    range: params.range
                });
            }

            // 计算目标范围
            var startRow = startCell.Row;
            var startCol = startCell.Column;
            var endRow = startRow + rowCount - 1;
            var endCol = startCol + colCount - 1;

            var targetRange = sheet.Range(
                sheet.Cells.Item(startRow, startCol),
                sheet.Cells.Item(endRow, endCol)
            );

            // 写入数据
            // 需要补齐每行的列数，确保是规整的二维数组
            var normalizedData = [];
            for (var r = 0; r < rowCount; r++) {
                var row = [];
                for (var c = 0; c < colCount; c++) {
                    if (data[r] && c < data[r].length) {
                        row.push(data[r][c]);
                    } else {
                        row.push(null);
                    }
                }
                normalizedData.push(row);
            }

            targetRange.Value2 = normalizedData;

            var result = {
                range: params.range,
                targetAddress: targetRange.Address(),
                rowCount: rowCount,
                columnCount: colCount,
                cellsWritten: rowCount * colCount
            };

            excelLogger.info('范围数据写入成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('写入范围数据失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    // ==================== 私有辅助方法 ====================

    /**
     * 列号转字母
     * 把数字列号转成Excel的字母表示，比如1->A, 27->AA
     *
     * @private
     * @param {number} col - 列号（从1开始）
     * @returns {string} 列字母
     */
    _columnToLetter: function(col) {
        var letter = '';
        while (col > 0) {
            var temp = (col - 1) % 26;
            letter = String.fromCharCode(65 + temp) + letter;
            col = Math.floor((col - temp - 1) / 26);
        }
        return letter;
    },

    /**
     * 获取公式错误的中文描述
     *
     * @private
     * @param {string} errorType - 错误类型
     * @returns {string} 错误描述
     */
    _getFormulaErrorMessage: function(errorType) {
        var messages = {
            '#DIV/0!': '除数为零',
            '#VALUE!': '值类型错误',
            '#REF!': '引用无效',
            '#NAME?': '名称无法识别',
            '#NUM!': '数值错误',
            '#N/A': '值不可用',
            '#NULL!': '空交集错误',
            '#ERROR!': '计算错误'
        };
        return messages[errorType] || '未知错误';
    },

    /**
     * 分析公式结构
     *
     * @private
     * @param {string} formula - 公式字符串
     * @returns {object} 分析结果
     */
    _analyzeFormulaStructure: function(formula) {
        if (!formula || formula.charAt(0) !== '=') {
            return { isFormula: false };
        }

        var analysis = {
            isFormula: true,
            formula: formula,
            functions: [],
            references: [],
            hasArrayFormula: formula.indexOf('{') === 0
        };

        // 提取使用的函数
        var funcPattern = /([A-Z_][A-Z0-9_]*)\s*\(/gi;
        var match;
        while ((match = funcPattern.exec(formula)) !== null) {
            if (analysis.functions.indexOf(match[1].toUpperCase()) === -1) {
                analysis.functions.push(match[1].toUpperCase());
            }
        }

        // 提取单元格引用
        var refPattern = /(\$?[A-Z]+\$?[0-9]+)/gi;
        while ((match = refPattern.exec(formula)) !== null) {
            if (analysis.references.indexOf(match[1]) === -1) {
                analysis.references.push(match[1]);
            }
        }

        return analysis;
    },

    /**
     * 清洗操作：去除空格
     *
     * @private
     * @param {object} sheet - 工作表对象
     * @param {object} range - 范围对象
     * @returns {object} 操作结果
     */
    _cleanTrim: function(sheet, range) {
        var trimmedCount = 0;
        var rowCount = range.Rows.Count;
        var colCount = range.Columns.Count;
        var startRow = range.Row;
        var startCol = range.Column;

        for (var r = 0; r < rowCount; r++) {
            for (var c = 0; c < colCount; c++) {
                var cell = sheet.Cells.Item(startRow + r, startCol + c);
                var value = cell.Value2;

                if (typeof value === 'string') {
                    var trimmed = value.trim();
                    // 同时去除中间多余的空格
                    trimmed = trimmed.replace(/\s+/g, ' ');

                    if (trimmed !== value) {
                        cell.Value2 = trimmed;
                        trimmedCount++;
                    }
                }
            }
        }

        return {
            operation: 'trim',
            success: true,
            details: {
                cellsProcessed: rowCount * colCount,
                cellsTrimmed: trimmedCount
            }
        };
    },

    /**
     * 清洗操作：删除重复行
     *
     * @private
     * @param {object} sheet - 工作表对象
     * @param {object} range - 范围对象
     * @returns {object} 操作结果
     */
    _cleanRemoveDuplicates: function(sheet, range) {
        var originalRowCount = range.Rows.Count;

        try {
            // 使用WPS内置的RemoveDuplicates方法
            // 参数是列索引数组，从1开始
            var columns = [];
            for (var i = 1; i <= range.Columns.Count; i++) {
                columns.push(i);
            }

            range.RemoveDuplicates(columns);

            // 重新获取范围的行数来计算删除了多少行
            // 注意：RemoveDuplicates会修改原范围
            var newRowCount = sheet.UsedRange.Rows.Count;
            var removedCount = originalRowCount - newRowCount;

            return {
                operation: 'remove_duplicates',
                success: true,
                details: {
                    originalRows: originalRowCount,
                    duplicatesRemoved: removedCount > 0 ? removedCount : 0
                }
            };
        } catch (err) {
            // 如果内置方法不可用，使用手动方式
            return this._cleanRemoveDuplicatesManual(sheet, range);
        }
    },

    /**
     * 手动删除重复行（备用方案）
     *
     * @private
     */
    _cleanRemoveDuplicatesManual: function(sheet, range) {
        var rowCount = range.Rows.Count;
        var colCount = range.Columns.Count;
        var startRow = range.Row;
        var startCol = range.Column;

        // 读取所有数据
        var rows = [];
        for (var r = 0; r < rowCount; r++) {
            var rowData = [];
            for (var c = 0; c < colCount; c++) {
                var cell = sheet.Cells.Item(startRow + r, startCol + c);
                rowData.push(cell.Value2);
            }
            rows.push({
                index: r,
                data: rowData,
                key: JSON.stringify(rowData)
            });
        }

        // 找出重复行（保留第一次出现的）
        var seen = {};
        var duplicateIndices = [];
        for (var i = 0; i < rows.length; i++) {
            if (seen[rows[i].key]) {
                duplicateIndices.push(startRow + rows[i].index);
            } else {
                seen[rows[i].key] = true;
            }
        }

        // 从后往前删除行，避免索引变化
        duplicateIndices.sort(function(a, b) { return b - a; });
        for (var d = 0; d < duplicateIndices.length; d++) {
            sheet.Rows.Item(duplicateIndices[d]).Delete();
        }

        return {
            operation: 'remove_duplicates',
            success: true,
            details: {
                originalRows: rowCount,
                duplicatesRemoved: duplicateIndices.length
            }
        };
    },

    /**
     * 清洗操作：删除空行
     *
     * @private
     * @param {object} sheet - 工作表对象
     * @param {object} range - 范围对象
     * @returns {object} 操作结果
     */
    _cleanRemoveEmptyRows: function(sheet, range) {
        var rowCount = range.Rows.Count;
        var colCount = range.Columns.Count;
        var startRow = range.Row;
        var startCol = range.Column;

        // 找出空行
        var emptyRowIndices = [];
        for (var r = 0; r < rowCount; r++) {
            var isEmpty = true;
            for (var c = 0; c < colCount; c++) {
                var cell = sheet.Cells.Item(startRow + r, startCol + c);
                var value = cell.Value2;
                if (value !== null && value !== undefined && value !== '') {
                    isEmpty = false;
                    break;
                }
            }
            if (isEmpty) {
                emptyRowIndices.push(startRow + r);
            }
        }

        // 从后往前删除行
        emptyRowIndices.sort(function(a, b) { return b - a; });
        for (var d = 0; d < emptyRowIndices.length; d++) {
            sheet.Rows.Item(emptyRowIndices[d]).Delete();
        }

        return {
            operation: 'remove_empty_rows',
            success: true,
            details: {
                originalRows: rowCount,
                emptyRowsRemoved: emptyRowIndices.length
            }
        };
    },

    /**
     * 清洗操作：统一日期格式
     *
     * @private
     * @param {object} sheet - 工作表对象
     * @param {object} range - 范围对象
     * @param {string} dateFormat - 目标日期格式，默认 "yyyy-mm-dd"
     * @returns {object} 操作结果
     */
    _cleanUnifyDate: function(sheet, range, dateFormat) {
        var format = dateFormat || 'yyyy-mm-dd';
        var rowCount = range.Rows.Count;
        var colCount = range.Columns.Count;
        var startRow = range.Row;
        var startCol = range.Column;
        var convertedCount = 0;

        for (var r = 0; r < rowCount; r++) {
            for (var c = 0; c < colCount; c++) {
                var cell = sheet.Cells.Item(startRow + r, startCol + c);
                var value = cell.Value2;

                // 检查是否是日期类型或可能是日期的字符串
                if (value !== null && value !== undefined) {
                    var isDateValue = false;

                    // Excel中日期存储为数字
                    if (typeof value === 'number' && value > 0 && value < 2958466) {
                        // 可能是日期序号（1900年1月1日到9999年12月31日）
                        isDateValue = true;
                    } else if (typeof value === 'string') {
                        // 尝试解析日期字符串
                        var datePatterns = [
                            /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/,  // yyyy-mm-dd 或 yyyy/mm/dd
                            /^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/,  // mm-dd-yyyy 或 dd-mm-yyyy
                            /^\d{4}年\d{1,2}月\d{1,2}日$/       // yyyy年mm月dd日
                        ];

                        for (var p = 0; p < datePatterns.length; p++) {
                            if (datePatterns[p].test(value)) {
                                isDateValue = true;
                                break;
                            }
                        }
                    }

                    if (isDateValue) {
                        try {
                            cell.NumberFormat = format;
                            convertedCount++;
                        } catch (fmtErr) {
                            // 格式设置失败，跳过
                        }
                    }
                }
            }
        }

        return {
            operation: 'unify_date',
            success: true,
            details: {
                targetFormat: format,
                cellsProcessed: rowCount * colCount,
                datesFormatted: convertedCount
            }
        };
    },

    // ==================== 透视表相关方法（马铁锤出品） ====================

    /**
     * 创建透视表
     * 老王说：透视表是数据分析的核心，这个方法必须写得漂漂亮亮的
     *
     * @param {object} params - 参数对象
     * @param {string} params.sourceRange - 数据源范围，如 "A1:E100"
     * @param {string} params.destinationCell - 透视表放置位置，如 "G1"
     * @param {string} params.destinationSheet - 目标工作表名称（可选）
     * @param {Array} params.rowFields - 行字段列名列表
     * @param {Array} params.columnFields - 列字段列名列表（可选）
     * @param {Array} params.valueFields - 值字段配置列表 [{field, aggregation}]
     * @param {Array} params.filterFields - 筛选字段列名列表（可选）
     * @param {string} params.tableName - 透视表名称（可选）
     * @returns {object} 标准响应对象
     */
    createPivotTable: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始创建透视表', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['sourceRange', 'destinationCell', 'rowFields', 'valueFields']);
            if (!validation.valid) {
                excelLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 校验行字段
            if (!Array.isArray(params.rowFields) || params.rowFields.length === 0) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: 'rowFields必须是非空数组'
                });
            }

            // 校验值字段
            if (!Array.isArray(params.valueFields) || params.valueFields.length === 0) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: 'valueFields必须是非空数组'
                });
            }

            // 校验值字段格式
            var validAggregations = ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'];
            for (var v = 0; v < params.valueFields.length; v++) {
                var vf = params.valueFields[v];
                if (!vf.field || !vf.aggregation) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: 'valueFields中每项必须包含field和aggregation'
                    });
                }
                if (validAggregations.indexOf(vf.aggregation.toUpperCase()) === -1) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: 'aggregation必须是SUM/COUNT/AVERAGE/MAX/MIN之一，你给的"' + vf.aggregation + '"是个啥？'
                    });
                }
            }

            // 获取活动工作簿
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var sourceSheet = workbook.ActiveSheet;
            if (!sourceSheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 获取数据源范围
            var sourceRange = sourceSheet.Range(params.sourceRange);
            if (!sourceRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    range: params.sourceRange
                });
            }

            // 确定目标工作表
            var destSheet = sourceSheet;
            if (params.destinationSheet) {
                try {
                    destSheet = workbook.Sheets.Item(params.destinationSheet);
                } catch (sheetErr) {
                    // 如果目标工作表不存在，创建一个新的
                    destSheet = workbook.Sheets.Add();
                    destSheet.Name = params.destinationSheet;
                    excelLogger.info('创建了新工作表', { name: params.destinationSheet });
                }
            }

            // 获取目标单元格
            var destCell = destSheet.Range(params.destinationCell);
            if (!destCell) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    cell: params.destinationCell
                });
            }

            // 构建表头映射（列名 -> 列索引）
            var headerMap = this._buildHeaderMap(sourceSheet, sourceRange);
            if (!headerMap || Object.keys(headerMap).length === 0) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '无法从数据源获取表头信息，确保第一行是表头'
                });
            }

            // 验证所有字段是否存在于表头中
            var allFields = params.rowFields.concat(
                params.columnFields || [],
                params.valueFields.map(function(vf) { return vf.field; }),
                params.filterFields || []
            );

            for (var f = 0; f < allFields.length; f++) {
                var fieldName = allFields[f];
                if (headerMap[fieldName] === undefined) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: '字段"' + fieldName + '"在表头中不存在，你确定名字写对了？'
                    });
                }
            }

            // 生成透视表名称
            var pivotTableName = params.tableName || ('PivotTable_' + Date.now());

            // 创建透视表缓存
            var pivotCache = workbook.PivotCaches().Create(
                1,  // xlDatabase
                sourceRange
            );

            // 创建透视表
            var pivotTable = pivotCache.CreatePivotTable(
                destCell,
                pivotTableName
            );

            // 添加行字段
            for (var r = 0; r < params.rowFields.length; r++) {
                var rowField = params.rowFields[r];
                try {
                    var rowPivotField = pivotTable.PivotFields(rowField);
                    rowPivotField.Orientation = 1;  // xlRowField
                    rowPivotField.Position = r + 1;
                } catch (rowErr) {
                    excelLogger.warn('添加行字段失败', { field: rowField, error: rowErr.message });
                }
            }

            // 添加列字段
            if (params.columnFields && params.columnFields.length > 0) {
                for (var c = 0; c < params.columnFields.length; c++) {
                    var colField = params.columnFields[c];
                    try {
                        var colPivotField = pivotTable.PivotFields(colField);
                        colPivotField.Orientation = 2;  // xlColumnField
                        colPivotField.Position = c + 1;
                    } catch (colErr) {
                        excelLogger.warn('添加列字段失败', { field: colField, error: colErr.message });
                    }
                }
            }

            // 添加值字段
            var aggregationMap = {
                'SUM': -4157,      // xlSum
                'COUNT': -4112,   // xlCount
                'AVERAGE': -4106, // xlAverage
                'MAX': -4136,     // xlMax
                'MIN': -4139      // xlMin
            };

            for (var vIdx = 0; vIdx < params.valueFields.length; vIdx++) {
                var valueField = params.valueFields[vIdx];
                try {
                    var dataField = pivotTable.AddDataField(
                        pivotTable.PivotFields(valueField.field),
                        valueField.field + '_' + valueField.aggregation.toUpperCase(),
                        aggregationMap[valueField.aggregation.toUpperCase()]
                    );
                } catch (valErr) {
                    excelLogger.warn('添加值字段失败', { field: valueField.field, error: valErr.message });
                }
            }

            // 添加筛选字段
            if (params.filterFields && params.filterFields.length > 0) {
                for (var ff = 0; ff < params.filterFields.length; ff++) {
                    var filterField = params.filterFields[ff];
                    try {
                        var filterPivotField = pivotTable.PivotFields(filterField);
                        filterPivotField.Orientation = 3;  // xlPageField
                        filterPivotField.Position = ff + 1;
                    } catch (filterErr) {
                        excelLogger.warn('添加筛选字段失败', { field: filterField, error: filterErr.message });
                    }
                }
            }

            // 获取透视表范围信息
            var pivotRange = pivotTable.TableRange2;
            var result = {
                success: true,
                pivotTableName: pivotTableName,
                location: destCell.Address() + ' (' + destSheet.Name + ')',
                rowCount: pivotRange ? pivotRange.Rows.Count : 0,
                columnCount: pivotRange ? pivotRange.Columns.Count : 0
            };

            excelLogger.info('透视表创建成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('创建透视表失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 更新透视表配置
     * 透视表创建之后经常需要调整，这个方法负责更新配置
     *
     * @param {object} params - 参数对象
     * @param {string} params.pivotTableName - 透视表名称
     * @param {string} params.pivotTableCell - 透视表所在的任意单元格地址（可选）
     * @param {Array} params.addRowFields - 要添加的行字段
     * @param {Array} params.removeRowFields - 要移除的行字段
     * @param {Array} params.addColumnFields - 要添加的列字段
     * @param {Array} params.removeColumnFields - 要移除的列字段
     * @param {Array} params.addValueFields - 要添加的值字段
     * @param {Array} params.removeValueFields - 要移除的值字段
     * @param {Array} params.updateValueFields - 要更新的值字段
     * @param {Array} params.addFilterFields - 要添加的筛选字段
     * @param {Array} params.removeFilterFields - 要移除的筛选字段
     * @param {boolean} params.refresh - 是否刷新数据
     * @returns {object} 标准响应对象
     */
    updatePivotTable: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始更新透视表', params);

        try {
            // 至少需要透视表名称或位置
            if (!params.pivotTableName && !params.pivotTableCell) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    message: '必须提供pivotTableName或pivotTableCell'
                });
            }

            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var sheet = workbook.ActiveSheet;
            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 查找透视表
            var pivotTable = null;

            if (params.pivotTableName) {
                // 通过名称查找
                try {
                    pivotTable = sheet.PivotTables(params.pivotTableName);
                } catch (nameErr) {
                    // 在所有工作表中查找
                    var sheetsCount = workbook.Sheets.Count;
                    for (var s = 1; s <= sheetsCount; s++) {
                        try {
                            var tempSheet = workbook.Sheets.Item(s);
                            pivotTable = tempSheet.PivotTables(params.pivotTableName);
                            if (pivotTable) {
                                sheet = tempSheet;
                                break;
                            }
                        } catch (e) {
                            // 继续查找
                        }
                    }
                }
            }

            if (!pivotTable && params.pivotTableCell) {
                // 通过单元格位置查找
                try {
                    var cell = sheet.Range(params.pivotTableCell);
                    if (cell && cell.PivotTable) {
                        pivotTable = cell.PivotTable;
                    }
                } catch (cellErr) {
                    excelLogger.warn('通过单元格查找透视表失败', { error: cellErr.message });
                }
            }

            if (!pivotTable) {
                return Response.error('PIVOT_TABLE_NOT_FOUND', null, startTime, {
                    message: '找不到指定的透视表，你确定名字写对了？'
                });
            }

            var operations = [];

            // 聚合方式映射
            var aggregationMap = {
                'SUM': -4157,
                'COUNT': -4112,
                'AVERAGE': -4106,
                'MAX': -4136,
                'MIN': -4139
            };

            // 移除行字段
            if (params.removeRowFields && params.removeRowFields.length > 0) {
                for (var rr = 0; rr < params.removeRowFields.length; rr++) {
                    var removeRowField = params.removeRowFields[rr];
                    try {
                        var rrField = pivotTable.PivotFields(removeRowField);
                        rrField.Orientation = 0;  // xlHidden
                        operations.push({
                            operation: '移除行字段: ' + removeRowField,
                            success: true,
                            message: '已移除'
                        });
                    } catch (rrErr) {
                        operations.push({
                            operation: '移除行字段: ' + removeRowField,
                            success: false,
                            message: rrErr.message
                        });
                    }
                }
            }

            // 添加行字段
            if (params.addRowFields && params.addRowFields.length > 0) {
                for (var ar = 0; ar < params.addRowFields.length; ar++) {
                    var addRowField = params.addRowFields[ar];
                    try {
                        var arField = pivotTable.PivotFields(addRowField);
                        arField.Orientation = 1;  // xlRowField
                        operations.push({
                            operation: '添加行字段: ' + addRowField,
                            success: true,
                            message: '已添加'
                        });
                    } catch (arErr) {
                        operations.push({
                            operation: '添加行字段: ' + addRowField,
                            success: false,
                            message: arErr.message
                        });
                    }
                }
            }

            // 移除列字段
            if (params.removeColumnFields && params.removeColumnFields.length > 0) {
                for (var rc = 0; rc < params.removeColumnFields.length; rc++) {
                    var removeColField = params.removeColumnFields[rc];
                    try {
                        var rcField = pivotTable.PivotFields(removeColField);
                        rcField.Orientation = 0;  // xlHidden
                        operations.push({
                            operation: '移除列字段: ' + removeColField,
                            success: true,
                            message: '已移除'
                        });
                    } catch (rcErr) {
                        operations.push({
                            operation: '移除列字段: ' + removeColField,
                            success: false,
                            message: rcErr.message
                        });
                    }
                }
            }

            // 添加列字段
            if (params.addColumnFields && params.addColumnFields.length > 0) {
                for (var ac = 0; ac < params.addColumnFields.length; ac++) {
                    var addColField = params.addColumnFields[ac];
                    try {
                        var acField = pivotTable.PivotFields(addColField);
                        acField.Orientation = 2;  // xlColumnField
                        operations.push({
                            operation: '添加列字段: ' + addColField,
                            success: true,
                            message: '已添加'
                        });
                    } catch (acErr) {
                        operations.push({
                            operation: '添加列字段: ' + addColField,
                            success: false,
                            message: acErr.message
                        });
                    }
                }
            }

            // 移除值字段
            if (params.removeValueFields && params.removeValueFields.length > 0) {
                for (var rv = 0; rv < params.removeValueFields.length; rv++) {
                    var removeValField = params.removeValueFields[rv];
                    try {
                        // 值字段在DataFields中
                        var dataFieldCount = pivotTable.DataFields.Count;
                        for (var df = dataFieldCount; df >= 1; df--) {
                            var dataField = pivotTable.DataFields.Item(df);
                            if (dataField.SourceName === removeValField ||
                                dataField.Name.indexOf(removeValField) !== -1) {
                                dataField.Orientation = 0;  // xlHidden
                                break;
                            }
                        }
                        operations.push({
                            operation: '移除值字段: ' + removeValField,
                            success: true,
                            message: '已移除'
                        });
                    } catch (rvErr) {
                        operations.push({
                            operation: '移除值字段: ' + removeValField,
                            success: false,
                            message: rvErr.message
                        });
                    }
                }
            }

            // 添加值字段
            if (params.addValueFields && params.addValueFields.length > 0) {
                for (var av = 0; av < params.addValueFields.length; av++) {
                    var addValField = params.addValueFields[av];
                    try {
                        pivotTable.AddDataField(
                            pivotTable.PivotFields(addValField.field),
                            addValField.field + '_' + addValField.aggregation.toUpperCase(),
                            aggregationMap[addValField.aggregation.toUpperCase()]
                        );
                        operations.push({
                            operation: '添加值字段: ' + addValField.field + '(' + addValField.aggregation + ')',
                            success: true,
                            message: '已添加'
                        });
                    } catch (avErr) {
                        operations.push({
                            operation: '添加值字段: ' + addValField.field,
                            success: false,
                            message: avErr.message
                        });
                    }
                }
            }

            // 更新值字段聚合方式
            if (params.updateValueFields && params.updateValueFields.length > 0) {
                for (var uv = 0; uv < params.updateValueFields.length; uv++) {
                    var updateValField = params.updateValueFields[uv];
                    try {
                        var dataFieldsCount = pivotTable.DataFields.Count;
                        for (var dfi = 1; dfi <= dataFieldsCount; dfi++) {
                            var dfItem = pivotTable.DataFields.Item(dfi);
                            if (dfItem.SourceName === updateValField.field ||
                                dfItem.Name.indexOf(updateValField.field) !== -1) {
                                dfItem.Function = aggregationMap[updateValField.aggregation.toUpperCase()];
                                dfItem.Name = updateValField.field + '_' + updateValField.aggregation.toUpperCase();
                                break;
                            }
                        }
                        operations.push({
                            operation: '更新值字段: ' + updateValField.field + ' -> ' + updateValField.aggregation,
                            success: true,
                            message: '已更新'
                        });
                    } catch (uvErr) {
                        operations.push({
                            operation: '更新值字段: ' + updateValField.field,
                            success: false,
                            message: uvErr.message
                        });
                    }
                }
            }

            // 移除筛选字段
            if (params.removeFilterFields && params.removeFilterFields.length > 0) {
                for (var rf = 0; rf < params.removeFilterFields.length; rf++) {
                    var removeFilterField = params.removeFilterFields[rf];
                    try {
                        var rfField = pivotTable.PivotFields(removeFilterField);
                        rfField.Orientation = 0;  // xlHidden
                        operations.push({
                            operation: '移除筛选字段: ' + removeFilterField,
                            success: true,
                            message: '已移除'
                        });
                    } catch (rfErr) {
                        operations.push({
                            operation: '移除筛选字段: ' + removeFilterField,
                            success: false,
                            message: rfErr.message
                        });
                    }
                }
            }

            // 添加筛选字段
            if (params.addFilterFields && params.addFilterFields.length > 0) {
                for (var af = 0; af < params.addFilterFields.length; af++) {
                    var addFilterField = params.addFilterFields[af];
                    try {
                        var afField = pivotTable.PivotFields(addFilterField);
                        afField.Orientation = 3;  // xlPageField
                        operations.push({
                            operation: '添加筛选字段: ' + addFilterField,
                            success: true,
                            message: '已添加'
                        });
                    } catch (afErr) {
                        operations.push({
                            operation: '添加筛选字段: ' + addFilterField,
                            success: false,
                            message: afErr.message
                        });
                    }
                }
            }

            // 刷新透视表
            if (params.refresh) {
                try {
                    pivotTable.RefreshTable();
                    operations.push({
                        operation: '刷新透视表',
                        success: true,
                        message: '已刷新'
                    });
                } catch (refreshErr) {
                    operations.push({
                        operation: '刷新透视表',
                        success: false,
                        message: refreshErr.message
                    });
                }
            }

            var result = {
                success: true,
                pivotTableName: pivotTable.Name,
                operations: operations
            };

            excelLogger.info('透视表更新完成', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('更新透视表失败', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 构建表头映射
     * 从数据源的第一行获取列名与列索引的映射
     *
     * @private
     * @param {object} sheet - 工作表对象
     * @param {object} range - 数据范围对象
     * @returns {object} 表头映射 {列名: 列索引}
     */
    _buildHeaderMap: function(sheet, range) {
        var headerMap = {};

        try {
            var startRow = range.Row;
            var startCol = range.Column;
            var colCount = range.Columns.Count;

            for (var c = 0; c < colCount; c++) {
                var cell = sheet.Cells.Item(startRow, startCol + c);
                var value = cell.Value2;

                if (value !== null && value !== undefined && value !== '') {
                    headerMap[String(value)] = c + 1;  // 列索引从1开始
                }
            }
        } catch (err) {
            excelLogger.warn('构建表头映射失败', { error: err.message });
        }

        return headerMap;
    },

    // ==================== 图表相关方法（刘大炮出品） ====================

    /**
     * 创建图表
     * 根据数据范围创建指定类型的图表
     *
     * @param {object} params - 参数对象
     * @param {string} params.dataRange - 数据范围，如 "A1:C10"
     * @param {number} params.chartType - WPS图表类型常量
     * @param {string} params.chartTypeName - 图表类型名称，用于日志
     * @param {string} params.title - 图表标题
     * @param {object} params.position - 图表位置 {left, top, width, height}
     * @param {string} params.sheet - 工作表名称
     * @param {boolean} params.hasHeader - 数据是否包含表头
     * @param {boolean} params.showLegend - 是否显示图例
     * @param {boolean} params.showDataLabels - 是否显示数据标签
     * @returns {object} 标准响应对象
     */
    createChart: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始创建图表，刘大炮要搞事情了', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['dataRange', 'chartType']);
            if (!validation.valid) {
                excelLogger.error('创建图表参数缺失，你TM传啥玩意儿', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取工作簿和工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                excelLogger.error('工作簿只读，图表创建不了');
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            // 获取目标工作表
            var sheet;
            if (params.sheet) {
                try {
                    sheet = workbook.Sheets.Item(params.sheet);
                } catch (sheetErr) {
                    excelLogger.error('找不到指定的工作表', { sheet: params.sheet });
                    return Response.error('SHEET_NOT_FOUND', null, startTime, {
                        sheet: params.sheet
                    });
                }
            } else {
                sheet = workbook.ActiveSheet;
            }

            if (!sheet) {
                return Response.error('SHEET_NOT_FOUND', null, startTime);
            }

            // 获取数据范围
            var dataRange;
            try {
                dataRange = sheet.Range(params.dataRange);
            } catch (rangeErr) {
                excelLogger.error('数据范围格式不对', { range: params.dataRange });
                return Response.error('PARAM_RANGE_INVALID', null, startTime, {
                    range: params.dataRange
                });
            }

            // 计算图表位置
            // 如果没指定位置，就放在数据范围的右侧
            var chartLeft = params.position && params.position.left !== undefined
                ? params.position.left
                : dataRange.Left + dataRange.Width + 20;
            var chartTop = params.position && params.position.top !== undefined
                ? params.position.top
                : dataRange.Top;
            var chartWidth = (params.position && params.position.width) || 480;
            var chartHeight = (params.position && params.position.height) || 300;

            // 创建图表对象
            // WPS/Excel的ChartObjects.Add方法
            var chartObjects = sheet.ChartObjects();
            var chartObject = chartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight);
            var chart = chartObject.Chart;

            // 设置图表数据源
            chart.SetSourceData(dataRange);

            // 设置图表类型
            chart.ChartType = params.chartType;

            // 设置图表标题
            if (params.title) {
                chart.HasTitle = true;
                chart.ChartTitle.Text = params.title;
            }

            // 设置图例
            if (params.showLegend === false) {
                chart.HasLegend = false;
            } else {
                chart.HasLegend = true;
            }

            // 设置数据标签
            if (params.showDataLabels) {
                try {
                    // 遍历所有系列，设置数据标签
                    var seriesCount = chart.SeriesCollection().Count;
                    for (var s = 1; s <= seriesCount; s++) {
                        var series = chart.SeriesCollection(s);
                        series.HasDataLabels = true;
                    }
                } catch (labelErr) {
                    excelLogger.warn('设置数据标签失败，但不影响图表创建', { error: labelErr.message });
                }
            }

            // 获取图表信息
            var chartName = chartObject.Name;
            var chartIndex = chartObject.Index;

            var result = {
                chartName: chartName,
                chartIndex: chartIndex,
                dataRange: params.dataRange,
                chartType: params.chartTypeName || params.chartType,
                position: {
                    left: chartLeft,
                    top: chartTop,
                    width: chartWidth,
                    height: chartHeight
                }
            };

            excelLogger.info('图表创建成功，刘大炮干得漂亮', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('创建图表失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 更新图表属性
     * 修改已存在图表的各种属性
     *
     * @param {object} params - 参数对象
     * @param {number} params.chartIndex - 图表索引
     * @param {string} params.chartName - 图表名称
     * @param {string} params.title - 新标题
     * @param {number} params.chartType - 新图表类型
     * @param {boolean} params.showLegend - 是否显示图例
     * @param {string} params.legendPosition - 图例位置
     * @param {boolean} params.showDataLabels - 是否显示数据标签
     * @param {string} params.dataRange - 新数据范围
     * @param {Array} params.colors - 系列颜色数组
     * @returns {object} 标准响应对象
     */
    updateChart: function(params) {
        var startTime = Date.now();
        excelLogger.info('开始更新图表属性', params);

        try {
            // 必须指定图表索引或名称
            if (params.chartIndex === undefined && !params.chartName) {
                excelLogger.error('没指定要更新哪个图表，憨批');
                return Response.error('PARAM_MISSING', null, startTime, {
                    message: '必须指定chartIndex或chartName'
                });
            }

            // 获取工作簿和工作表
            var workbook = Application.ActiveWorkbook;
            if (!workbook) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (workbook.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            // 获取目标工作表
            var sheet;
            if (params.sheet) {
                try {
                    sheet = workbook.Sheets.Item(params.sheet);
                } catch (sheetErr) {
                    return Response.error('SHEET_NOT_FOUND', null, startTime, {
                        sheet: params.sheet
                    });
                }
            } else {
                sheet = workbook.ActiveSheet;
            }

            // 查找图表对象
            var chartObject = null;
            var chartObjects = sheet.ChartObjects();

            if (params.chartIndex !== undefined) {
                try {
                    chartObject = chartObjects.Item(params.chartIndex);
                } catch (indexErr) {
                    excelLogger.error('找不到指定索引的图表', { index: params.chartIndex });
                    return Response.error('CHART_NOT_FOUND', null, startTime, {
                        chartIndex: params.chartIndex
                    });
                }
            } else if (params.chartName) {
                try {
                    chartObject = chartObjects.Item(params.chartName);
                } catch (nameErr) {
                    excelLogger.error('找不到指定名称的图表', { name: params.chartName });
                    return Response.error('CHART_NOT_FOUND', null, startTime, {
                        chartName: params.chartName
                    });
                }
            }

            if (!chartObject) {
                return Response.error('CHART_NOT_FOUND', null, startTime);
            }

            var chart = chartObject.Chart;
            var updatedProperties = [];

            // 更新标题
            if (params.title !== undefined) {
                if (params.title) {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = params.title;
                } else {
                    chart.HasTitle = false;
                }
                updatedProperties.push('title');
            }

            // 更新图表类型
            if (params.chartType !== undefined) {
                chart.ChartType = params.chartType;
                updatedProperties.push('chartType');
            }

            // 更新图例显示
            if (params.showLegend !== undefined) {
                chart.HasLegend = params.showLegend;
                updatedProperties.push('showLegend');
            }

            // 更新图例位置
            if (params.legendPosition !== undefined && chart.HasLegend) {
                var legendPositionMap = {
                    'bottom': -4107,   // xlLegendPositionBottom
                    'top': -4160,      // xlLegendPositionTop
                    'left': -4131,     // xlLegendPositionLeft
                    'right': -4152     // xlLegendPositionRight
                };
                var posValue = legendPositionMap[params.legendPosition];
                if (posValue !== undefined) {
                    chart.Legend.Position = posValue;
                    updatedProperties.push('legendPosition');
                }
            }

            // 更新数据标签
            if (params.showDataLabels !== undefined) {
                try {
                    var seriesCount = chart.SeriesCollection().Count;
                    for (var s = 1; s <= seriesCount; s++) {
                        var series = chart.SeriesCollection(s);
                        series.HasDataLabels = params.showDataLabels;
                    }
                    updatedProperties.push('showDataLabels');
                } catch (labelErr) {
                    excelLogger.warn('更新数据标签失败', { error: labelErr.message });
                }
            }

            // 更新数据源
            if (params.dataRange !== undefined) {
                try {
                    var newRange = sheet.Range(params.dataRange);
                    chart.SetSourceData(newRange);
                    updatedProperties.push('dataRange');
                } catch (rangeErr) {
                    excelLogger.warn('更新数据源失败', { error: rangeErr.message });
                }
            }

            // 更新系列颜色
            if (params.colors && Array.isArray(params.colors) && params.colors.length > 0) {
                try {
                    var seriesCount = chart.SeriesCollection().Count;
                    for (var i = 0; i < Math.min(params.colors.length, seriesCount); i++) {
                        var series = chart.SeriesCollection(i + 1);
                        var colorHex = params.colors[i];
                        // 将十六进制颜色转换为RGB值
                        var rgb = this._hexToRgb(colorHex);
                        if (rgb) {
                            // WPS使用RGB函数创建颜色值
                            var colorValue = rgb.r + rgb.g * 256 + rgb.b * 65536;
                            series.Format.Fill.ForeColor.RGB = colorValue;
                        }
                    }
                    updatedProperties.push('colors');
                } catch (colorErr) {
                    excelLogger.warn('更新系列颜色失败', { error: colorErr.message });
                }
            }

            var result = {
                chartName: chartObject.Name,
                updatedProperties: updatedProperties
            };

            excelLogger.info('图表更新成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            excelLogger.error('更新图表失败', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 十六进制颜色转RGB
     *
     * @private
     * @param {string} hex - 十六进制颜色值，如 "#FF0000"
     * @returns {object|null} RGB对象 {r, g, b} 或 null
     */
    _hexToRgb: function(hex) {
        // 去掉#号
        hex = hex.replace(/^#/, '');

        // 处理简写形式 #F00 -> #FF0000
        if (hex.length === 3) {
            hex = hex.split('').map(function(c) { return c + c; }).join('');
        }

        if (hex.length !== 6) {
            return null;
        }

        var r = parseInt(hex.substring(0, 2), 16);
        var g = parseInt(hex.substring(2, 4), 16);
        var b = parseInt(hex.substring(4, 6), 16);

        if (isNaN(r) || isNaN(g) || isNaN(b)) {
            return null;
        }

        return { r: r, g: g, b: b };
    }
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        ExcelHandler: ExcelHandler
    };
}
