/**
 * WPS Claude 智能助手 - Common 处理器
 *
 * 艹，二狗来写通用处理器了
 * 这个SB模块负责处理所有跨应用的通用操作
 * 主要是文档格式转换，Word/Excel/PPT都能用
 *
 * 注意：转换功能需要调用WPS的ExportAsFixedFormat或SaveAs方法
 * 不同应用的方法参数不太一样，这里做了统一封装
 *
 * @author 孙二狗
 * @date 2026-01-24
 */

// 创建日志记录器
var commonLogger = new Logger('CommonHandler');

/**
 * Common 处理器
 * 这个对象封装了所有跨应用通用操作的核心方法
 */
var CommonHandler = {

    /**
     * 转换为PDF
     * 自动检测当前打开的文档类型（Word/Excel/PPT），然后导出为PDF
     *
     * @param {object} params - 参数对象
     * @param {string} params.outputPath - PDF输出路径，可选，不填则用原文件路径改扩展名
     * @param {boolean} params.openAfterExport - 导出后是否打开PDF，默认false
     * @returns {object} 标准响应对象
     */
    convertToPdf: function(params) {
        var startTime = Date.now();
        commonLogger.info('开始转换为PDF', params);

        try {
            params = params || {};

            // 检测当前活动的应用类型
            var appInfo = this._detectActiveApp();
            if (!appInfo.success) {
                commonLogger.error('无法检测到活动文档', appInfo);
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: appInfo.message || '请先打开一个Word/Excel/PPT文档'
                });
            }

            var appType = appInfo.appType;
            var activeDoc = appInfo.document;
            var sourcePath = appInfo.path;
            var sourceFileName = appInfo.fileName;

            // 检查是否已保存
            if (!sourcePath && !params.outputPath) {
                commonLogger.error('文档未保存且未指定输出路径');
                return Response.error('PARAM_MISSING', null, startTime, {
                    message: '文档未保存，请先保存或指定输出路径'
                });
            }

            // 确定输出路径
            var outputPath = params.outputPath;
            if (!outputPath) {
                // 用原文件路径，把扩展名改成.pdf
                outputPath = sourcePath.replace(/\.[^.]+$/, '.pdf');
            }

            // 根据应用类型调用不同的导出方法
            var result = null;
            switch (appType) {
                case 'word':
                    result = this._exportWordToPdf(activeDoc, outputPath, params);
                    break;
                case 'excel':
                    result = this._exportExcelToPdf(activeDoc, outputPath, params);
                    break;
                case 'ppt':
                    result = this._exportPptToPdf(activeDoc, outputPath, params);
                    break;
                default:
                    return Response.error('INTERNAL_ERROR', null, startTime, {
                        message: '不支持的应用类型: ' + appType
                    });
            }

            if (!result.success) {
                return Response.error('INTERNAL_ERROR', null, startTime, {
                    message: result.message
                });
            }

            var responseData = {
                sourcePath: sourcePath || sourceFileName,
                outputPath: outputPath,
                appType: appType,
                pageCount: result.pageCount
            };

            commonLogger.info('PDF导出成功', responseData);
            return Response.success(responseData, null, startTime);

        } catch (err) {
            commonLogger.error('PDF导出失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 格式互转
     * 将当前文档转换为其他格式
     *
     * @param {object} params - 参数对象
     * @param {string} params.targetFormat - 目标格式扩展名，如doc, xlsx, ppt等
     * @param {string} params.outputPath - 输出路径，可选
     * @returns {object} 标准响应对象
     */
    convertFormat: function(params) {
        var startTime = Date.now();
        commonLogger.info('开始格式转换', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['targetFormat']);
            if (!validation.valid) {
                commonLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            var targetFormat = params.targetFormat.toLowerCase().replace(/^\./, '');

            // 检测当前活动的应用类型
            var appInfo = this._detectActiveApp();
            if (!appInfo.success) {
                commonLogger.error('无法检测到活动文档', appInfo);
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: appInfo.message || '请先打开一个Word/Excel/PPT文档'
                });
            }

            var appType = appInfo.appType;
            var activeDoc = appInfo.document;
            var sourcePath = appInfo.path;
            var sourceFileName = appInfo.fileName;

            // 获取源文件格式
            var sourceFormat = '';
            if (sourcePath) {
                var parts = sourcePath.split('.');
                sourceFormat = parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
            }

            // 检查目标格式是否支持
            var formatCode = this._getFormatCode(targetFormat, appType);
            if (formatCode === -1) {
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '不支持的目标格式: ' + targetFormat,
                    suggestion: '请检查目标格式是否正确，Word支持doc/docx/rtf/txt/html等，Excel支持xls/xlsx/csv/html等，PPT支持ppt/pptx/html等'
                });
            }

            // 确定输出路径
            var outputPath = params.outputPath;
            if (!outputPath) {
                if (!sourcePath) {
                    // 文档未保存，需要指定输出路径
                    return Response.error('PARAM_MISSING', null, startTime, {
                        message: '文档未保存，请先保存或指定输出路径'
                    });
                }
                // 用原文件路径，改扩展名
                outputPath = sourcePath.replace(/\.[^.]+$/, '.' + targetFormat);
            }

            // 根据应用类型调用不同的保存方法
            var result = null;
            switch (appType) {
                case 'word':
                    result = this._saveWordAs(activeDoc, outputPath, targetFormat, formatCode);
                    break;
                case 'excel':
                    result = this._saveExcelAs(activeDoc, outputPath, targetFormat, formatCode);
                    break;
                case 'ppt':
                    result = this._savePptAs(activeDoc, outputPath, targetFormat, formatCode);
                    break;
                default:
                    return Response.error('INTERNAL_ERROR', null, startTime, {
                        message: '不支持的应用类型: ' + appType
                    });
            }

            if (!result.success) {
                return Response.error('INTERNAL_ERROR', null, startTime, {
                    message: result.message
                });
            }

            var responseData = {
                sourcePath: sourcePath || sourceFileName,
                sourceFormat: sourceFormat,
                targetFormat: targetFormat,
                outputPath: outputPath,
                appType: appType
            };

            commonLogger.info('格式转换成功', responseData);
            return Response.success(responseData, null, startTime);

        } catch (err) {
            commonLogger.error('格式转换失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    // ==================== 私有辅助方法 ====================

    /**
     * 检测当前活动的应用类型
     * 依次检查Word、Excel、PPT是否有活动文档
     *
     * @private
     * @returns {object} 检测结果
     */
    _detectActiveApp: function() {
        // 尝试获取Word活动文档
        try {
            if (typeof Application !== 'undefined' && Application.ActiveDocument) {
                var doc = Application.ActiveDocument;
                return {
                    success: true,
                    appType: 'word',
                    document: doc,
                    path: doc.FullName || doc.Path + '\\' + doc.Name || '',
                    fileName: doc.Name
                };
            }
        } catch (wordErr) {
            // Word没有活动文档，继续检查
        }

        // 尝试获取Excel活动工作簿
        try {
            if (typeof Application !== 'undefined' && Application.ActiveWorkbook) {
                var wb = Application.ActiveWorkbook;
                return {
                    success: true,
                    appType: 'excel',
                    document: wb,
                    path: wb.FullName || wb.Path + '\\' + wb.Name || '',
                    fileName: wb.Name
                };
            }
        } catch (excelErr) {
            // Excel没有活动工作簿，继续检查
        }

        // 尝试获取PPT活动演示文稿
        try {
            if (typeof Application !== 'undefined' && Application.ActivePresentation) {
                var pres = Application.ActivePresentation;
                return {
                    success: true,
                    appType: 'ppt',
                    document: pres,
                    path: pres.FullName || pres.Path + '\\' + pres.Name || '',
                    fileName: pres.Name
                };
            }
        } catch (pptErr) {
            // PPT没有活动演示文稿
        }

        return {
            success: false,
            message: '没有找到活动文档，请先打开一个Word/Excel/PPT文档'
        };
    },

    /**
     * 导出Word文档为PDF
     *
     * @private
     */
    _exportWordToPdf: function(doc, outputPath, params) {
        try {
            // Word的ExportAsFixedFormat方法
            // OutputFileName, ExportFormat (wdExportFormatPDF=17, wdExportFormatXPS=18)
            // OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, etc.
            doc.ExportAsFixedFormat(
                outputPath,         // OutputFileName
                17,                 // ExportFormat: wdExportFormatPDF = 17
                params.openAfterExport || false,  // OpenAfterExport
                0,                  // OptimizeFor: wdExportOptimizeForPrint = 0
                0,                  // Range: wdExportAllDocument = 0
                1,                  // From (starting page, 1-based)
                0,                  // To (ending page, 0 = all)
                0                   // Item: wdExportDocumentContent = 0
            );

            // 获取页数
            var pageCount = 0;
            try {
                pageCount = doc.ComputeStatistics(2); // wdStatisticPages = 2
            } catch (pcErr) {
                // 获取页数失败不影响整体
            }

            return {
                success: true,
                pageCount: pageCount
            };
        } catch (err) {
            commonLogger.error('Word导出PDF失败', { error: err.message });
            return {
                success: false,
                message: 'Word导出PDF失败: ' + err.message
            };
        }
    },

    /**
     * 导出Excel工作簿为PDF
     *
     * @private
     */
    _exportExcelToPdf: function(wb, outputPath, params) {
        try {
            // Excel的ExportAsFixedFormat方法
            // Type (xlTypePDF=0, xlTypeXPS=1), Filename, Quality, IncludeDocProps,
            // IgnorePrintAreas, From, To, OpenAfterPublish
            wb.ExportAsFixedFormat(
                0,                  // Type: xlTypePDF = 0
                outputPath,         // Filename
                0,                  // Quality: xlQualityStandard = 0
                true,               // IncludeDocProperties
                false,              // IgnorePrintAreas
                null,               // From (starting page)
                null,               // To (ending page)
                params.openAfterExport || false  // OpenAfterPublish
            );

            return {
                success: true
            };
        } catch (err) {
            commonLogger.error('Excel导出PDF失败', { error: err.message });
            return {
                success: false,
                message: 'Excel导出PDF失败: ' + err.message
            };
        }
    },

    /**
     * 导出PPT演示文稿为PDF
     *
     * @private
     */
    _exportPptToPdf: function(pres, outputPath, params) {
        try {
            // PPT的SaveAs方法，使用PDF格式
            // ppSaveAsPDF = 32
            pres.SaveAs(
                outputPath,         // FileName
                32                  // FileFormat: ppSaveAsPDF = 32
            );

            // 如果需要打开，用Shell打开
            if (params.openAfterExport) {
                try {
                    var shell = new ActiveXObject('WScript.Shell');
                    shell.Run('"' + outputPath + '"');
                } catch (shellErr) {
                    commonLogger.warn('打开PDF失败', { error: shellErr.message });
                }
            }

            var pageCount = 0;
            try {
                pageCount = pres.Slides.Count;
            } catch (scErr) {
                // 获取幻灯片数失败不影响整体
            }

            return {
                success: true,
                pageCount: pageCount
            };
        } catch (err) {
            commonLogger.error('PPT导出PDF失败', { error: err.message });
            return {
                success: false,
                message: 'PPT导出PDF失败: ' + err.message
            };
        }
    },

    /**
     * 获取格式代码
     *
     * @private
     */
    _getFormatCode: function(format, appType) {
        var formatLower = format.toLowerCase();

        switch (appType) {
            case 'word':
                // Word文档格式代码
                var wordFormats = {
                    'doc': 0,       // wdFormatDocument
                    'docx': 16,     // wdFormatDocumentDefault
                    'pdf': 17,      // wdFormatPDF
                    'rtf': 6,       // wdFormatRTF
                    'xps': 18,      // wdFormatXPS
                    'html': 8,      // wdFormatHTML
                    'htm': 8,
                    'txt': 2,       // wdFormatText
                    'xml': 11,      // wdFormatXML
                    'odt': 23       // wdFormatOpenDocumentText
                };
                return wordFormats.hasOwnProperty(formatLower) ? wordFormats[formatLower] : -1;

            case 'excel':
                // Excel工作簿格式代码
                var excelFormats = {
                    'xls': 56,      // xlExcel8 (.xls)
                    'xlsx': 51,     // xlOpenXMLWorkbook
                    'xlsm': 52,     // xlOpenXMLWorkbookMacroEnabled
                    'xlsb': 50,     // xlExcel12
                    'csv': 6,       // xlCSV
                    'html': 44,     // xlHtml
                    'htm': 44,
                    'txt': -4158,   // xlCurrentPlatformText
                    'xml': 46,      // xlXMLSpreadsheet
                    'ods': 60       // xlOpenDocumentSpreadsheet
                };
                return excelFormats.hasOwnProperty(formatLower) ? excelFormats[formatLower] : -1;

            case 'ppt':
                // PPT演示格式代码
                var pptFormats = {
                    'ppt': 1,       // ppSaveAsPresentation
                    'pptx': 24,     // ppSaveAsOpenXMLPresentation
                    'pptm': 25,     // ppSaveAsOpenXMLPresentationMacroEnabled
                    'pdf': 32,      // ppSaveAsPDF
                    'xps': 33,      // ppSaveAsXPS
                    'html': 12,     // ppSaveAsHTML
                    'htm': 12,
                    'png': 18,      // ppSaveAsPNG
                    'jpg': 17,      // ppSaveAsJPG
                    'jpeg': 17,
                    'gif': 16,      // ppSaveAsGIF
                    'bmp': 19,      // ppSaveAsBMP
                    'odp': 35       // ppSaveAsOpenDocumentPresentation
                };
                return pptFormats.hasOwnProperty(formatLower) ? pptFormats[formatLower] : -1;

            default:
                return -1;
        }
    },

    /**
     * Word另存为指定格式
     *
     * @private
     */
    _saveWordAs: function(doc, outputPath, format, formatCode) {
        try {
            // PDF用ExportAsFixedFormat，其他格式用SaveAs2
            if (format === 'pdf') {
                return this._exportWordToPdf(doc, outputPath, {});
            }

            // SaveAs2方法
            // FileName, FileFormat, LockComments, Password, AddToRecentFiles,
            // WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, etc.
            doc.SaveAs2(
                outputPath,         // FileName
                formatCode          // FileFormat
            );

            return { success: true };
        } catch (err) {
            commonLogger.error('Word另存为失败', { error: err.message, format: format });
            return {
                success: false,
                message: 'Word另存为' + format + '失败: ' + err.message
            };
        }
    },

    /**
     * Excel另存为指定格式
     *
     * @private
     */
    _saveExcelAs: function(wb, outputPath, format, formatCode) {
        try {
            // PDF用ExportAsFixedFormat
            if (format === 'pdf') {
                return this._exportExcelToPdf(wb, outputPath, {});
            }

            // SaveAs方法
            wb.SaveAs(
                outputPath,         // Filename
                formatCode          // FileFormat
            );

            return { success: true };
        } catch (err) {
            commonLogger.error('Excel另存为失败', { error: err.message, format: format });
            return {
                success: false,
                message: 'Excel另存为' + format + '失败: ' + err.message
            };
        }
    },

    /**
     * PPT另存为指定格式
     *
     * @private
     */
    _savePptAs: function(pres, outputPath, format, formatCode) {
        try {
            // SaveAs方法
            pres.SaveAs(
                outputPath,         // FileName
                formatCode          // FileFormat
            );

            return { success: true };
        } catch (err) {
            commonLogger.error('PPT另存为失败', { error: err.message, format: format });
            return {
                success: false,
                message: 'PPT另存为' + format + '失败: ' + err.message
            };
        }
    }
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        CommonHandler: CommonHandler
    };
}
