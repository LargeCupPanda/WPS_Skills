/**
 * WPS Claude 智能助手 - Word 处理器
 *
 * 艹，这个SB模块是老王的替补写的，负责处理所有Word/文字相关的操作
 * 包括获取上下文、插入文本、应用样式、生成目录、设置字体、查找替换等
 *
 * 再强调一遍：所有API都必须严格按照WPS SDK文档来写，别TM瞎编！
 * WPS的Word API和微软的VBA基本兼容，Application.ActiveDocument就是根对象
 *
 * @author 王五的替补
 * @date 2026-01-24
 */

// 创建日志记录器
var wordLogger = new Logger('WordHandler');

/**
 * Word 处理器
 * 这个憨批对象封装了所有Word操作的核心方法
 */
var WordHandler = {

    /**
     * 获取文档上下文
     * 返回当前文档的所有关键信息，让AI知道现在文档是什么状态
     * 包括文档名称、段落数量、字数统计、文档结构、当前选中内容
     *
     * @returns {object} 标准响应对象，包含上下文信息
     */
    getContext: function() {
        var startTime = Date.now();
        wordLogger.info('开始获取文档上下文');

        try {
            // 先检查有没有活动文档，没有就直接报错滚蛋
            var doc = Application.ActiveDocument;
            if (!doc) {
                wordLogger.error('没有找到活动文档，你TM打开个Word再说');
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: '请先打开一个Word文档'
                });
            }

            // 获取基本文档信息
            var docInfo = {
                name: doc.Name,
                path: doc.Path || '',
                fullName: doc.FullName || doc.Name,
                saved: doc.Saved,
                readOnly: doc.ReadOnly
            };

            // 获取段落数量
            var paragraphCount = 0;
            try {
                paragraphCount = doc.Paragraphs.Count;
            } catch (paraErr) {
                wordLogger.warn('获取段落数量失败', { error: paraErr.message });
            }

            // 获取字数统计
            var wordCount = null;
            try {
                // WPS Word的统计属性
                // wdStatisticWords = 0 字数
                // wdStatisticCharacters = 3 字符数（不计空格）
                // wdStatisticCharactersWithSpaces = 5 字符数（计空格）
                // wdStatisticParagraphs = 4 段落数
                // wdStatisticPages = 2 页数
                wordCount = {
                    words: doc.ComputeStatistics(0),           // wdStatisticWords
                    characters: doc.ComputeStatistics(3),       // wdStatisticCharacters
                    charactersWithSpaces: doc.ComputeStatistics(5), // wdStatisticCharactersWithSpaces
                    paragraphs: doc.ComputeStatistics(4),       // wdStatisticParagraphs
                    pages: doc.ComputeStatistics(2)             // wdStatisticPages
                };
            } catch (statErr) {
                wordLogger.warn('获取字数统计失败', { error: statErr.message });
                // 备用方案：使用Range的字符数
                try {
                    var range = doc.Content;
                    wordCount = {
                        characters: range.Characters.Count,
                        words: range.Words.Count
                    };
                } catch (rangeErr) {
                    wordLogger.warn('备用统计方案也失败了，憨批', { error: rangeErr.message });
                }
            }

            // 获取文档结构（标题层级）
            var structure = [];
            try {
                // 遍历段落，找出所有标题
                var maxHeadings = 50; // 最多获取50个标题，防止文档太大卡死
                var headingCount = 0;

                for (var i = 1; i <= paragraphCount && headingCount < maxHeadings; i++) {
                    var para = doc.Paragraphs.Item(i);
                    var style = para.Style;

                    if (style) {
                        var styleName = '';
                        try {
                            styleName = style.NameLocal || style.Name || '';
                        } catch (styleErr) {
                            styleName = '';
                        }

                        // 判断是否是标题样式（中文或英文）
                        var headingLevel = this._getHeadingLevel(styleName);
                        if (headingLevel > 0) {
                            var headingText = '';
                            try {
                                headingText = para.Range.Text;
                                // 去掉段落标记
                                headingText = headingText.replace(/[\r\n]+/g, '').trim();
                            } catch (textErr) {
                                headingText = '';
                            }

                            structure.push({
                                level: headingLevel,
                                text: headingText.substring(0, 100), // 最多100字符
                                paragraphIndex: i,
                                style: styleName
                            });
                            headingCount++;
                        }
                    }
                }
            } catch (structErr) {
                wordLogger.warn('获取文档结构失败', { error: structErr.message });
            }

            // 获取当前选中内容
            var selection = null;
            try {
                var sel = Application.Selection;
                if (sel) {
                    var selType = sel.Type;
                    // wdSelectionNormal = 1, wdSelectionIP = 0 (光标), wdNoSelection = -1
                    selection = {
                        type: selType,
                        typeName: this._getSelectionTypeName(selType),
                        start: sel.Start,
                        end: sel.End
                    };

                    // 如果有选中文本，获取文本内容
                    if (selType === 1 && sel.Text) { // wdSelectionNormal
                        selection.text = sel.Text.substring(0, 500); // 最多500字符
                        selection.length = sel.Text.length;
                    }

                    // 获取选区的样式信息
                    try {
                        var selStyle = sel.Style;
                        if (selStyle) {
                            selection.style = selStyle.NameLocal || selStyle.Name || '';
                        }
                    } catch (selStyleErr) {
                        // 样式获取失败不影响整体
                    }

                    // 获取选区的字体信息
                    try {
                        var font = sel.Font;
                        if (font) {
                            selection.font = {
                                name: font.Name,
                                size: font.Size,
                                bold: font.Bold,
                                italic: font.Italic,
                                color: font.Color
                            };
                        }
                    } catch (fontErr) {
                        // 字体获取失败不影响整体
                    }
                }
            } catch (selErr) {
                wordLogger.warn('获取选区信息失败', { error: selErr.message });
            }

            // 获取页面设置信息
            var pageSetup = null;
            try {
                var ps = doc.PageSetup;
                if (ps) {
                    pageSetup = {
                        pageWidth: ps.PageWidth,
                        pageHeight: ps.PageHeight,
                        topMargin: ps.TopMargin,
                        bottomMargin: ps.BottomMargin,
                        leftMargin: ps.LeftMargin,
                        rightMargin: ps.RightMargin,
                        orientation: ps.Orientation // 0=纵向, 1=横向
                    };
                }
            } catch (psErr) {
                wordLogger.warn('获取页面设置失败', { error: psErr.message });
            }

            // 构建上下文对象
            var context = {
                document: docInfo,
                paragraphCount: paragraphCount,
                wordCount: wordCount,
                structure: structure,
                selection: selection,
                pageSetup: pageSetup
            };

            wordLogger.info('文档上下文获取成功', { paragraphCount: paragraphCount });
            return Response.success(context, null, startTime);

        } catch (err) {
            wordLogger.error('获取文档上下文失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 插入文本
     * 在指定位置插入文本内容
     *
     * @param {object} params - 参数对象
     * @param {string} params.text - 要插入的文本内容
     * @param {string} params.position - 插入位置：start（文档开头）、end（文档结尾）、cursor（光标处）
     * @param {object} params.style - 可选，文本样式
     * @returns {object} 标准响应对象
     */
    insertText: function(params) {
        var startTime = Date.now();
        wordLogger.info('开始插入文本', { position: params ? params.position : 'unknown' });

        try {
            // 参数校验，别TM传空的进来
            var validation = Validator.checkRequired(params, ['text']);
            if (!validation.valid) {
                wordLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动文档
            var doc = Application.ActiveDocument;
            if (!doc) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            // 检查是否只读
            if (doc.ReadOnly) {
                wordLogger.error('文档是只读的，改不了');
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var position = params.position || 'cursor';
            var text = params.text;
            var targetRange = null;

            // 根据position确定插入位置
            switch (position.toLowerCase()) {
                case 'start':
                    // 文档开头
                    targetRange = doc.Range(0, 0);
                    break;

                case 'end':
                    // 文档结尾
                    var endPos = doc.Content.End;
                    targetRange = doc.Range(endPos, endPos);
                    break;

                case 'cursor':
                default:
                    // 光标当前位置（使用Selection）
                    targetRange = Application.Selection.Range;
                    break;
            }

            if (!targetRange) {
                wordLogger.error('找不到插入位置，这什么憨批情况');
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    message: '无法确定插入位置'
                });
            }

            // 插入文本
            targetRange.InsertAfter(text);

            // 如果指定了样式，应用样式
            if (params.style) {
                try {
                    // 获取刚插入的文本范围
                    var insertedStart = targetRange.Start;
                    var insertedEnd = insertedStart + text.length;
                    var insertedRange = doc.Range(insertedStart, insertedEnd);

                    // 应用样式
                    if (params.style.styleName) {
                        insertedRange.Style = params.style.styleName;
                    }
                    if (params.style.fontName) {
                        insertedRange.Font.Name = params.style.fontName;
                    }
                    if (params.style.fontSize) {
                        insertedRange.Font.Size = params.style.fontSize;
                    }
                    if (params.style.bold !== undefined) {
                        insertedRange.Font.Bold = params.style.bold ? -1 : 0;
                    }
                    if (params.style.italic !== undefined) {
                        insertedRange.Font.Italic = params.style.italic ? -1 : 0;
                    }
                } catch (styleErr) {
                    wordLogger.warn('应用样式失败，但文本已插入', { error: styleErr.message });
                }
            }

            var result = {
                text: text.substring(0, 100) + (text.length > 100 ? '...' : ''),
                textLength: text.length,
                position: position,
                insertedAt: targetRange.Start
            };

            wordLogger.info('文本插入成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            wordLogger.error('插入文本失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 应用样式
     * 对指定范围应用Word内置或自定义样式
     *
     * @param {object} params - 参数对象
     * @param {string} params.styleName - 样式名称，如"标题 1"、"正文"、"Heading 1"等
     * @param {string} params.range - 可选，范围类型：selection（当前选区）、paragraph（当前段落）、all（全文）
     * @returns {object} 标准响应对象
     */
    applyStyle: function(params) {
        var startTime = Date.now();
        wordLogger.info('开始应用样式', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['styleName']);
            if (!validation.valid) {
                wordLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动文档
            var doc = Application.ActiveDocument;
            if (!doc) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (doc.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var styleName = params.styleName;
            var rangeType = params.range || 'selection';
            var targetRange = null;

            // 先检查样式是否存在
            var styleExists = false;
            try {
                var testStyle = doc.Styles.Item(styleName);
                if (testStyle) {
                    styleExists = true;
                }
            } catch (styleCheckErr) {
                wordLogger.error('样式不存在', { styleName: styleName });
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '样式 "' + styleName + '" 不存在',
                    suggestion: '请使用有效的样式名称，如"标题 1"、"正文"、"Heading 1"等'
                });
            }

            // 根据rangeType确定目标范围
            switch (rangeType.toLowerCase()) {
                case 'selection':
                    targetRange = Application.Selection.Range;
                    break;

                case 'paragraph':
                    // 当前段落
                    var para = Application.Selection.Paragraphs.Item(1);
                    if (para) {
                        targetRange = para.Range;
                    }
                    break;

                case 'all':
                    targetRange = doc.Content;
                    break;

                default:
                    targetRange = Application.Selection.Range;
            }

            if (!targetRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    message: '无法确定目标范围'
                });
            }

            // 应用样式
            targetRange.Style = styleName;

            var result = {
                styleName: styleName,
                rangeType: rangeType,
                rangeStart: targetRange.Start,
                rangeEnd: targetRange.End,
                textLength: targetRange.End - targetRange.Start
            };

            wordLogger.info('样式应用成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            wordLogger.error('应用样式失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 生成目录
     * 自动根据文档标题生成目录
     *
     * @param {object} params - 参数对象
     * @param {string} params.position - 目录插入位置：start（文档开头）、cursor（光标处）
     * @param {number} params.levels - 目录层级，默认3（显示1-3级标题）
     * @returns {object} 标准响应对象
     */
    generateTOC: function(params) {
        var startTime = Date.now();
        wordLogger.info('开始生成目录', params);

        try {
            params = params || {};

            // 获取活动文档
            var doc = Application.ActiveDocument;
            if (!doc) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (doc.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var position = params.position || 'start';
            var levels = params.levels || 3;

            // 限制层级范围
            if (levels < 1) levels = 1;
            if (levels > 9) levels = 9;

            var targetRange = null;

            // 确定插入位置
            switch (position.toLowerCase()) {
                case 'start':
                    targetRange = doc.Range(0, 0);
                    break;

                case 'cursor':
                default:
                    targetRange = Application.Selection.Range;
                    break;
            }

            if (!targetRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    message: '无法确定目录插入位置'
                });
            }

            // 在插入目录之前，先插入一个段落分隔
            targetRange.InsertParagraphAfter();

            // 生成目录
            // TableOfContents.Add方法参数：
            // Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel,
            // UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers,
            // AddedStyles, UseHyperlinks, HidePageNumbersInWeb, UseOutlineLevels
            var toc = null;
            try {
                toc = doc.TablesOfContents.Add(
                    targetRange,     // 范围
                    true,           // 使用标题样式
                    1,              // 上层级（最高级别）
                    levels,         // 下层级
                    false,          // 不使用域
                    '',             // TableID
                    true,           // 右对齐页码
                    true,           // 包含页码
                    '',             // 添加的样式
                    true            // 使用超链接
                );
            } catch (tocErr) {
                // 如果上面的方法不行，尝试简化版本
                wordLogger.warn('完整目录生成失败，尝试简化方案', { error: tocErr.message });
                try {
                    toc = doc.TablesOfContents.Add(targetRange, true, 1, levels);
                } catch (tocErr2) {
                    wordLogger.error('目录生成失败', { error: tocErr2.message });
                    return Response.error('INTERNAL_ERROR', null, startTime, {
                        message: '目录生成失败: ' + tocErr2.message,
                        suggestion: '请确保文档中有使用标题样式的段落'
                    });
                }
            }

            var result = {
                position: position,
                levels: levels,
                tocCreated: toc !== null
            };

            // 获取目录信息
            if (toc) {
                try {
                    result.tocRange = {
                        start: toc.Range.Start,
                        end: toc.Range.End
                    };
                } catch (tocInfoErr) {
                    // 获取目录信息失败不影响整体
                }
            }

            wordLogger.info('目录生成成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            wordLogger.error('生成目录失败，艹', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 设置字体
     * 设置指定范围的字体格式
     *
     * @param {object} params - 参数对象
     * @param {string} params.fontName - 字体名称，如"宋体"、"微软雅黑"、"Arial"
     * @param {number} params.fontSize - 字号（磅）
     * @param {boolean} params.bold - 是否加粗
     * @param {boolean} params.italic - 是否斜体
     * @param {string|number} params.color - 字体颜色，可以是颜色名或RGB值
     * @param {string} params.range - 范围：selection（选区）、paragraph（段落）、all（全文）
     * @returns {object} 标准响应对象
     */
    setFont: function(params) {
        var startTime = Date.now();
        wordLogger.info('开始设置字体', params);

        try {
            params = params || {};

            // 获取活动文档
            var doc = Application.ActiveDocument;
            if (!doc) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (doc.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var rangeType = params.range || 'selection';
            var targetRange = null;

            // 确定目标范围
            switch (rangeType.toLowerCase()) {
                case 'selection':
                    targetRange = Application.Selection.Range;
                    break;

                case 'paragraph':
                    var para = Application.Selection.Paragraphs.Item(1);
                    if (para) {
                        targetRange = para.Range;
                    }
                    break;

                case 'all':
                    targetRange = doc.Content;
                    break;

                default:
                    targetRange = Application.Selection.Range;
            }

            if (!targetRange) {
                return Response.error('CELL_NOT_FOUND', null, startTime, {
                    message: '无法确定目标范围'
                });
            }

            // 获取Font对象
            var font = targetRange.Font;
            var changes = [];

            // 应用字体设置
            if (params.fontName !== undefined && params.fontName !== null) {
                font.Name = params.fontName;
                changes.push('fontName: ' + params.fontName);
            }

            if (params.fontSize !== undefined && params.fontSize !== null) {
                font.Size = params.fontSize;
                changes.push('fontSize: ' + params.fontSize);
            }

            if (params.bold !== undefined) {
                // WPS中，-1表示true，0表示false
                font.Bold = params.bold ? -1 : 0;
                changes.push('bold: ' + params.bold);
            }

            if (params.italic !== undefined) {
                font.Italic = params.italic ? -1 : 0;
                changes.push('italic: ' + params.italic);
            }

            if (params.underline !== undefined) {
                // wdUnderlineSingle = 1, wdUnderlineNone = 0
                font.Underline = params.underline ? 1 : 0;
                changes.push('underline: ' + params.underline);
            }

            if (params.color !== undefined && params.color !== null) {
                var colorValue = this._parseColor(params.color);
                if (colorValue !== null) {
                    font.Color = colorValue;
                    changes.push('color: ' + params.color);
                }
            }

            if (params.strikethrough !== undefined) {
                font.StrikeThrough = params.strikethrough ? -1 : 0;
                changes.push('strikethrough: ' + params.strikethrough);
            }

            var result = {
                rangeType: rangeType,
                rangeStart: targetRange.Start,
                rangeEnd: targetRange.End,
                changesApplied: changes
            };

            wordLogger.info('字体设置成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            wordLogger.error('设置字体失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 查找替换
     * 在文档中查找并替换文本
     *
     * @param {object} params - 参数对象
     * @param {string} params.findText - 要查找的文本
     * @param {string} params.replaceText - 替换后的文本
     * @param {boolean} params.replaceAll - 是否替换所有匹配项，默认true
     * @param {boolean} params.matchCase - 是否区分大小写，默认false
     * @param {boolean} params.matchWholeWord - 是否全字匹配，默认false
     * @returns {object} 标准响应对象
     */
    findReplace: function(params) {
        var startTime = Date.now();
        wordLogger.info('开始查找替换', { findText: params ? params.findText : 'unknown' });

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['findText']);
            if (!validation.valid) {
                wordLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动文档
            var doc = Application.ActiveDocument;
            if (!doc) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            // 如果要替换，检查是否只读
            if (params.replaceText !== undefined && doc.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var findText = params.findText;
            var replaceText = params.replaceText !== undefined ? params.replaceText : '';
            var replaceAll = params.replaceAll !== false; // 默认true
            var matchCase = params.matchCase === true;
            var matchWholeWord = params.matchWholeWord === true;

            // 使用Find对象进行查找替换
            var range = doc.Content;
            var find = range.Find;

            // 清除之前的查找设置
            find.ClearFormatting();
            if (find.Replacement) {
                find.Replacement.ClearFormatting();
            }

            // 设置查找参数
            find.Text = findText;
            find.MatchCase = matchCase;
            find.MatchWholeWord = matchWholeWord;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;
            find.Forward = true;
            find.Wrap = 1; // wdFindContinue = 1

            var replaceCount = 0;
            var foundPositions = [];

            if (params.replaceText !== undefined) {
                // 执行替换
                find.Replacement.Text = replaceText;

                // wdReplaceNone = 0, wdReplaceOne = 1, wdReplaceAll = 2
                var replaceType = replaceAll ? 2 : 1;

                // 执行替换操作
                var result = find.Execute(
                    findText,       // FindText
                    matchCase,      // MatchCase
                    matchWholeWord, // MatchWholeWord
                    false,          // MatchWildcards
                    false,          // MatchSoundsLike
                    false,          // MatchAllWordForms
                    true,           // Forward
                    1,              // Wrap (wdFindContinue)
                    false,          // Format
                    replaceText,    // ReplaceWith
                    replaceType     // Replace
                );

                // 计算替换次数（通过再次查找统计）
                // 注意：WPS的Execute返回的是boolean，不是替换次数
                // 我们需要用另一种方式统计
                if (result) {
                    // 简单估算：如果替换成功，至少替换了1次
                    // 如果是replaceAll，需要重新计数
                    if (replaceAll) {
                        // 重新搜索替换后的文本来估算替换次数
                        // 这不是精确的，但可以给出大概的数字
                        try {
                            var searchRange = doc.Content;
                            var searchFind = searchRange.Find;
                            searchFind.ClearFormatting();
                            searchFind.Text = replaceText;
                            searchFind.Forward = true;
                            searchFind.Wrap = 0; // wdFindStop

                            while (searchFind.Execute()) {
                                replaceCount++;
                                if (replaceCount > 1000) break; // 防止无限循环
                            }
                        } catch (countErr) {
                            replaceCount = -1; // 表示无法统计
                        }
                    } else {
                        replaceCount = 1;
                    }
                }
            } else {
                // 只查找不替换，收集所有匹配位置
                find.Wrap = 0; // wdFindStop = 0，不循环

                var maxFinds = 100; // 最多记录100个位置
                while (find.Execute() && foundPositions.length < maxFinds) {
                    foundPositions.push({
                        start: range.Start,
                        end: range.End,
                        text: range.Text
                    });
                    // 移动到下一个位置继续查找
                    range.Start = range.End;
                }
            }

            var resultObj = {
                findText: findText,
                replaceText: params.replaceText !== undefined ? replaceText : null,
                matchCase: matchCase,
                matchWholeWord: matchWholeWord,
                replaceAll: replaceAll
            };

            if (params.replaceText !== undefined) {
                resultObj.replaced = true;
                resultObj.replaceCount = replaceCount;
            } else {
                resultObj.replaced = false;
                resultObj.foundCount = foundPositions.length;
                resultObj.positions = foundPositions;
            }

            wordLogger.info('查找替换完成', resultObj);
            return Response.success(resultObj, null, startTime);

        } catch (err) {
            wordLogger.error('查找替换失败', { error: err.message });
            return Response.fromException(err, null, startTime);
        }
    },

    // ==================== 私有辅助方法 ====================

    /**
     * 获取标题级别
     * 根据样式名称判断是几级标题
     *
     * @private
     * @param {string} styleName - 样式名称
     * @returns {number} 标题级别（1-9），0表示不是标题
     */
    _getHeadingLevel: function(styleName) {
        if (!styleName) return 0;

        var name = styleName.toLowerCase();

        // 中文样式名
        var chineseHeadings = {
            '标题 1': 1, '标题 2': 2, '标题 3': 3, '标题 4': 4,
            '标题 5': 5, '标题 6': 6, '标题 7': 7, '标题 8': 8, '标题 9': 9,
            '标题1': 1, '标题2': 2, '标题3': 3, '标题4': 4,
            '标题5': 5, '标题6': 6, '标题7': 7, '标题8': 8, '标题9': 9
        };

        // 检查中文样式
        for (var key in chineseHeadings) {
            if (styleName === key) {
                return chineseHeadings[key];
            }
        }

        // 英文样式名 (Heading 1, Heading 2, ...)
        var headingMatch = name.match(/^heading\s*(\d)$/i);
        if (headingMatch) {
            return parseInt(headingMatch[1], 10);
        }

        // TOC样式
        var tocMatch = name.match(/^toc\s*(\d)$/i);
        if (tocMatch) {
            return parseInt(tocMatch[1], 10);
        }

        return 0;
    },

    /**
     * 获取选区类型名称
     *
     * @private
     * @param {number} selType - 选区类型数值
     * @returns {string} 选区类型名称
     */
    _getSelectionTypeName: function(selType) {
        var typeNames = {
            '-1': 'wdNoSelection',      // 没有选区
            '0': 'wdSelectionIP',       // 光标（插入点）
            '1': 'wdSelectionNormal',   // 正常选区
            '2': 'wdSelectionFrame',    // 框架选区
            '3': 'wdSelectionColumn',   // 列选区
            '4': 'wdSelectionRow',      // 行选区
            '5': 'wdSelectionBlock',    // 块选区
            '6': 'wdSelectionInlineShape', // 内联形状
            '7': 'wdSelectionShape'     // 形状
        };
        return typeNames[String(selType)] || 'Unknown';
    },

    /**
     * 解析颜色值
     * 支持颜色名称、十六进制、RGB数值
     *
     * @private
     * @param {string|number} color - 颜色值
     * @returns {number|null} WPS颜色值（BGR格式）
     */
    _parseColor: function(color) {
        if (color === null || color === undefined) return null;

        // 如果已经是数字，直接返回
        if (typeof color === 'number') {
            return color;
        }

        var colorStr = String(color).toLowerCase().trim();

        // 预定义颜色名称映射（WPS使用BGR格式）
        var colorNames = {
            'black': 0x000000,
            'white': 0xFFFFFF,
            'red': 0x0000FF,       // BGR格式的红色
            'green': 0x00FF00,
            'blue': 0xFF0000,      // BGR格式的蓝色
            'yellow': 0x00FFFF,
            'cyan': 0xFFFF00,
            'magenta': 0xFF00FF,
            'gray': 0x808080,
            'grey': 0x808080,
            'orange': 0x00A5FF,
            'purple': 0x800080,
            'pink': 0xCBC0FF
        };

        if (colorNames[colorStr] !== undefined) {
            return colorNames[colorStr];
        }

        // 解析十六进制颜色（#RRGGBB格式，需要转换为BGR）
        var hexMatch = colorStr.match(/^#?([0-9a-f]{6})$/i);
        if (hexMatch) {
            var hex = hexMatch[1];
            var r = parseInt(hex.substr(0, 2), 16);
            var g = parseInt(hex.substr(2, 2), 16);
            var b = parseInt(hex.substr(4, 2), 16);
            // 转换为BGR格式
            return (b << 16) | (g << 8) | r;
        }

        // 解析RGB格式 rgb(r, g, b)
        var rgbMatch = colorStr.match(/^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
        if (rgbMatch) {
            var r = parseInt(rgbMatch[1], 10);
            var g = parseInt(rgbMatch[2], 10);
            var b = parseInt(rgbMatch[3], 10);
            // 转换为BGR格式
            return (b << 16) | (g << 8) | r;
        }

        return null;
    }
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        WordHandler: WordHandler
    };
}
