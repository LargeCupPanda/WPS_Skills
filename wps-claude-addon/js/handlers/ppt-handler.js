/**
 * WPS Claude 智能助手 - PPT 处理器
 *
 * 艹，赵六那货跑了，留下这个烂摊子让老王来收拾
 * 这个SB模块负责处理所有PPT/演示文稿相关的操作
 * 包括获取上下文、添加幻灯片、添加文本框、美化幻灯片等
 *
 * 注意：所有API都必须严格按照WPS SDK文档来写，别TM瞎编！
 * WPS PPT的对象模型：Application -> ActivePresentation -> Slides -> Shapes
 *
 * @author 老王 (接替赵六的锅)
 * @date 2026-01-24
 */

// 创建日志记录器
var pptLogger = new Logger('PPTHandler');

/**
 * 预定义的配色方案
 * 这些配色都是老王熬夜调出来的，别TM说丑
 */
var ColorSchemes = {
    // 商务风格 - 稳重大气
    business: {
        name: '商务风格',
        primary: '#1F4E79',      // 深蓝色
        secondary: '#2E75B6',    // 中蓝色
        accent: '#BDD7EE',       // 浅蓝色
        background: '#FFFFFF',   // 白色背景
        text: '#333333',         // 深灰文字
        textLight: '#666666'     // 浅灰文字
    },
    // 科技风格 - 现代感
    tech: {
        name: '科技风格',
        primary: '#00B4D8',      // 青色
        secondary: '#0077B6',    // 深青色
        accent: '#90E0EF',       // 浅青色
        background: '#FFFFFF',
        text: '#212529',
        textLight: '#6C757D'
    },
    // 创意风格 - 活泼多彩
    creative: {
        name: '创意风格',
        primary: '#FF6B6B',      // 珊瑚红
        secondary: '#4ECDC4',    // 青绿色
        accent: '#FFE66D',       // 明黄色
        background: '#FFFFFF',
        text: '#2C3E50',
        textLight: '#7F8C8D'
    },
    // 极简风格 - 简洁清爽
    minimal: {
        name: '极简风格',
        primary: '#2D3436',      // 深灰黑
        secondary: '#636E72',    // 中灰色
        accent: '#B2BEC3',       // 浅灰色
        background: '#FFFFFF',
        text: '#2D3436',
        textLight: '#636E72'
    }
};

/**
 * PPT 处理器
 * 这个憨批对象封装了所有PPT操作的核心方法
 */
var PPTHandler = {

    /**
     * 获取演示文稿上下文
     * 返回当前PPT的所有关键信息，让AI知道现在演示文稿是什么状态
     *
     * @returns {object} 标准响应对象，包含上下文信息
     */
    getContext: function() {
        var startTime = Date.now();
        pptLogger.info('开始获取演示文稿上下文');

        try {
            // 先检查有没有活动演示文稿，没有就直接报错滚蛋
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                pptLogger.error('没有找到活动演示文稿，你TM打开个PPT再说');
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: '请先打开一个PPT演示文稿'
                });
            }

            // 获取幻灯片数量
            var slidesCount = presentation.Slides.Count;

            // 获取当前选中的幻灯片索引
            var currentSlideIndex = 1;
            try {
                // 通过ActiveWindow获取当前视图的幻灯片
                var activeWindow = Application.ActiveWindow;
                if (activeWindow && activeWindow.View) {
                    var slide = activeWindow.View.Slide;
                    if (slide) {
                        currentSlideIndex = slide.SlideIndex;
                    }
                }
            } catch (viewErr) {
                pptLogger.warn('获取当前幻灯片索引失败，使用默认值1', { error: viewErr.message });
            }

            // 获取每页幻灯片的信息
            var slidesInfo = [];
            for (var i = 1; i <= slidesCount; i++) {
                var slide = presentation.Slides.Item(i);
                var slideInfo = {
                    index: i,
                    slideId: slide.SlideID,
                    layout: this._getLayoutName(slide.Layout),
                    shapes: []
                };

                // 获取该幻灯片上的所有形状信息
                var shapesCount = slide.Shapes.Count;
                for (var j = 1; j <= shapesCount; j++) {
                    var shape = slide.Shapes.Item(j);
                    var shapeInfo = this._getShapeInfo(shape);
                    slideInfo.shapes.push(shapeInfo);
                }

                slidesInfo.push(slideInfo);
            }

            // 获取演示文稿的设计信息
            var designInfo = null;
            try {
                if (presentation.SlideMaster) {
                    designInfo = {
                        hasMaster: true,
                        colorScheme: this._getCurrentColorScheme(presentation)
                    };
                }
            } catch (designErr) {
                pptLogger.warn('获取设计信息失败', { error: designErr.message });
            }

            // 构建上下文对象
            var context = {
                presentation: {
                    name: presentation.Name,
                    path: presentation.Path || '',
                    fullName: presentation.FullName || presentation.Name,
                    saved: presentation.Saved,
                    readOnly: presentation.ReadOnly
                },
                slides: {
                    count: slidesCount,
                    currentIndex: currentSlideIndex,
                    list: slidesInfo
                },
                design: designInfo
            };

            pptLogger.info('演示文稿上下文获取成功', { slideCount: slidesCount });
            return Response.success(context, null, startTime);

        } catch (err) {
            pptLogger.error('获取演示文稿上下文失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 添加幻灯片
     * 在指定位置添加新的幻灯片
     *
     * @param {object} params - 参数对象
     * @param {number} params.position - 插入位置（可选，默认在最后）
     * @param {number} params.layout - 布局类型（可选，默认空白）
     * @param {string} params.title - 标题文本（可选）
     * @param {string} params.content - 内容文本（可选）
     * @returns {object} 标准响应对象
     */
    addSlide: function(params) {
        var startTime = Date.now();
        pptLogger.info('开始添加幻灯片', params);

        try {
            // 获取活动演示文稿
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime, {
                    message: '请先打开一个PPT演示文稿'
                });
            }

            // 检查是否只读
            if (presentation.ReadOnly) {
                pptLogger.error('演示文稿是只读的，加不了幻灯片');
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            params = params || {};

            // 确定插入位置
            var position = params.position;
            if (!position || position < 1) {
                position = presentation.Slides.Count + 1;
            }
            if (position > presentation.Slides.Count + 1) {
                position = presentation.Slides.Count + 1;
            }

            // 确定布局类型
            // WPS PPT布局常量：
            // 1 - ppLayoutTitle (标题幻灯片)
            // 2 - ppLayoutText (标题和文本)
            // 3 - ppLayoutTwoColumnText (两栏文本)
            // 7 - ppLayoutBlank (空白)
            // 12 - ppLayoutTitleOnly (仅标题)
            var layout = params.layout || 7; // 默认空白布局

            // 添加幻灯片
            var newSlide = presentation.Slides.Add(position, layout);

            if (!newSlide) {
                pptLogger.error('添加幻灯片失败，返回了空对象');
                return Response.error('INTERNAL_ERROR', null, startTime, {
                    message: '添加幻灯片失败'
                });
            }

            // 如果提供了标题，设置标题
            if (params.title) {
                try {
                    // 查找标题占位符
                    var titleShape = this._findPlaceholder(newSlide, 1); // 1 = ppPlaceholderTitle
                    if (titleShape && titleShape.TextFrame && titleShape.TextFrame.TextRange) {
                        titleShape.TextFrame.TextRange.Text = params.title;
                    } else {
                        // 没有标题占位符，添加一个文本框作为标题
                        var titleBox = newSlide.Shapes.AddTextbox(
                            1, // msoTextOrientationHorizontal
                            50, // Left
                            30, // Top
                            620, // Width
                            60  // Height
                        );
                        titleBox.TextFrame.TextRange.Text = params.title;
                        titleBox.TextFrame.TextRange.Font.Size = 32;
                        titleBox.TextFrame.TextRange.Font.Bold = true;
                    }
                } catch (titleErr) {
                    pptLogger.warn('设置标题失败', { error: titleErr.message });
                }
            }

            // 如果提供了内容，设置内容
            if (params.content) {
                try {
                    // 查找内容占位符
                    var contentShape = this._findPlaceholder(newSlide, 2); // 2 = ppPlaceholderBody
                    if (contentShape && contentShape.TextFrame && contentShape.TextFrame.TextRange) {
                        contentShape.TextFrame.TextRange.Text = params.content;
                    } else {
                        // 没有内容占位符，添加一个文本框
                        var contentBox = newSlide.Shapes.AddTextbox(
                            1, // msoTextOrientationHorizontal
                            50,  // Left
                            100, // Top
                            620, // Width
                            350  // Height
                        );
                        contentBox.TextFrame.TextRange.Text = params.content;
                        contentBox.TextFrame.TextRange.Font.Size = 18;
                    }
                } catch (contentErr) {
                    pptLogger.warn('设置内容失败', { error: contentErr.message });
                }
            }

            var result = {
                slideIndex: newSlide.SlideIndex,
                slideId: newSlide.SlideID,
                layout: this._getLayoutName(layout),
                position: position,
                hasTitle: !!params.title,
                hasContent: !!params.content
            };

            pptLogger.info('幻灯片添加成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            pptLogger.error('添加幻灯片失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 添加文本框
     * 在指定幻灯片上添加文本框
     *
     * @param {object} params - 参数对象
     * @param {number} params.slideIndex - 幻灯片索引（从1开始）
     * @param {string} params.text - 文本内容
     * @param {object} params.position - 位置信息 {left, top, width, height}
     * @param {object} params.style - 样式信息 {fontSize, fontName, bold, italic, color}
     * @returns {object} 标准响应对象
     */
    addTextBox: function(params) {
        var startTime = Date.now();
        pptLogger.info('开始添加文本框', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['slideIndex', 'text']);
            if (!validation.valid) {
                pptLogger.error('参数校验失败，缺少必要参数', { missing: validation.missing });
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动演示文稿
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (presentation.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            // 校验幻灯片索引
            if (params.slideIndex < 1 || params.slideIndex > presentation.Slides.Count) {
                pptLogger.error('幻灯片索引超出范围', {
                    index: params.slideIndex,
                    max: presentation.Slides.Count
                });
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '幻灯片索引超出范围，有效范围: 1-' + presentation.Slides.Count
                });
            }

            // 获取目标幻灯片
            var slide = presentation.Slides.Item(params.slideIndex);

            // 确定位置和大小，使用默认值如果没有提供
            var pos = params.position || {};
            var left = pos.left !== undefined ? pos.left : 100;
            var top = pos.top !== undefined ? pos.top : 100;
            var width = pos.width !== undefined ? pos.width : 400;
            var height = pos.height !== undefined ? pos.height : 100;

            // 添加文本框
            // AddTextbox(Orientation, Left, Top, Width, Height)
            // Orientation: 1 = msoTextOrientationHorizontal
            var textBox = slide.Shapes.AddTextbox(1, left, top, width, height);

            if (!textBox) {
                return Response.error('INTERNAL_ERROR', null, startTime, {
                    message: '添加文本框失败'
                });
            }

            // 设置文本内容
            textBox.TextFrame.TextRange.Text = params.text;

            // 应用样式
            var style = params.style || {};
            var textRange = textBox.TextFrame.TextRange;

            if (style.fontSize) {
                textRange.Font.Size = style.fontSize;
            }
            if (style.fontName) {
                textRange.Font.Name = style.fontName;
            }
            if (style.bold !== undefined) {
                textRange.Font.Bold = style.bold ? -1 : 0; // -1 = msoTrue, 0 = msoFalse
            }
            if (style.italic !== undefined) {
                textRange.Font.Italic = style.italic ? -1 : 0;
            }
            if (style.color) {
                // 颜色可以是RGB值或十六进制字符串
                var colorValue = this._parseColor(style.color);
                if (colorValue !== null) {
                    textRange.Font.Color.RGB = colorValue;
                }
            }

            // 设置文本框的一些默认属性，让它看起来不那么丑
            try {
                textBox.TextFrame.WordWrap = -1; // 自动换行
                textBox.TextFrame.AutoSize = 0;  // 不自动调整大小
            } catch (frameErr) {
                // 某些属性可能不支持，忽略
            }

            var result = {
                slideIndex: params.slideIndex,
                shapeId: textBox.Id,
                shapeName: textBox.Name,
                position: {
                    left: left,
                    top: top,
                    width: width,
                    height: height
                },
                text: params.text,
                style: style
            };

            pptLogger.info('文本框添加成功', result);
            return Response.success(result, null, startTime);

        } catch (err) {
            pptLogger.error('添加文本框失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 美化幻灯片（核心功能）
     * 统一字体、应用配色、对齐元素、优化间距
     *
     * @param {object} params - 参数对象
     * @param {number} params.slideIndex - 幻灯片索引（可选，不填则处理全部）
     * @param {string} params.style - 样式风格（business/tech/creative/minimal）
     * @returns {object} 标准响应对象
     */
    beautifySlide: function(params) {
        var startTime = Date.now();
        pptLogger.info('开始美化幻灯片，让这个憨批PPT变得好看点', params);

        try {
            // 获取活动演示文稿
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (presentation.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            params = params || {};
            var styleName = params.style || 'business';
            var colorScheme = ColorSchemes[styleName];

            if (!colorScheme) {
                pptLogger.warn('未知的样式风格，使用默认商务风格', { style: styleName });
                colorScheme = ColorSchemes.business;
                styleName = 'business';
            }

            // 确定要处理的幻灯片范围
            var startIndex = 1;
            var endIndex = presentation.Slides.Count;

            if (params.slideIndex) {
                if (params.slideIndex < 1 || params.slideIndex > presentation.Slides.Count) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: '幻灯片索引超出范围'
                    });
                }
                startIndex = params.slideIndex;
                endIndex = params.slideIndex;
            }

            var results = {
                style: styleName,
                colorScheme: colorScheme.name,
                slidesProcessed: 0,
                operations: []
            };

            // 遍历处理每张幻灯片
            for (var i = startIndex; i <= endIndex; i++) {
                var slide = presentation.Slides.Item(i);
                var slideResult = {
                    slideIndex: i,
                    shapesProcessed: 0,
                    fontUnified: 0,
                    colorApplied: 0,
                    aligned: 0
                };

                // 处理该幻灯片上的所有形状
                var shapesCount = slide.Shapes.Count;
                for (var j = 1; j <= shapesCount; j++) {
                    var shape = slide.Shapes.Item(j);

                    // 1. 统一字体
                    if (this._hasTextFrame(shape)) {
                        try {
                            var textRange = shape.TextFrame.TextRange;
                            // 统一字体为微软雅黑或默认字体
                            textRange.Font.Name = '微软雅黑';
                            // 根据形状类型设置颜色
                            var isTitle = this._isLikelyTitle(shape);
                            var textColor = isTitle ?
                                this._parseColor(colorScheme.primary) :
                                this._parseColor(colorScheme.text);
                            if (textColor !== null) {
                                textRange.Font.Color.RGB = textColor;
                            }
                            slideResult.fontUnified++;
                        } catch (fontErr) {
                            // 字体设置失败，继续处理下一个
                        }
                    }

                    // 2. 对齐形状（简单的左对齐逻辑）
                    try {
                        // 检查形状是否太靠边缘，如果是就调整
                        if (shape.Left < 20) {
                            shape.Left = 50;
                            slideResult.aligned++;
                        }
                        if (shape.Top < 20) {
                            shape.Top = 30;
                            slideResult.aligned++;
                        }
                    } catch (alignErr) {
                        // 对齐失败，继续
                    }

                    slideResult.shapesProcessed++;
                }

                // 3. 尝试设置幻灯片背景色
                try {
                    slide.FollowMasterBackground = 0; // 不跟随母版
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = this._parseColor(colorScheme.background);
                    slideResult.colorApplied++;
                } catch (bgErr) {
                    pptLogger.warn('设置背景色失败', { slide: i, error: bgErr.message });
                }

                results.operations.push(slideResult);
                results.slidesProcessed++;
            }

            pptLogger.info('幻灯片美化完成，虽然还是有点丑但比之前好多了', results);
            return Response.success(results, null, startTime);

        } catch (err) {
            pptLogger.error('美化幻灯片失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 统一字体
     * 将指定幻灯片（或全部）的所有文本统一为指定字体
     *
     * @param {object} params - 参数对象
     * @param {string} params.fontName - 字体名称
     * @param {number} params.slideIndex - 幻灯片索引（可选，不填则处理全部）
     * @returns {object} 标准响应对象
     */
    unifyFont: function(params) {
        var startTime = Date.now();
        pptLogger.info('开始统一字体', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['fontName']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取活动演示文稿
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (presentation.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            var fontName = params.fontName;

            // 确定要处理的幻灯片范围
            var startIndex = 1;
            var endIndex = presentation.Slides.Count;

            if (params.slideIndex) {
                if (params.slideIndex < 1 || params.slideIndex > presentation.Slides.Count) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: '幻灯片索引超出范围'
                    });
                }
                startIndex = params.slideIndex;
                endIndex = params.slideIndex;
            }

            var results = {
                fontName: fontName,
                slidesProcessed: 0,
                textRangesModified: 0,
                details: []
            };

            // 遍历处理每张幻灯片
            for (var i = startIndex; i <= endIndex; i++) {
                var slide = presentation.Slides.Item(i);
                var slideModified = 0;

                var shapesCount = slide.Shapes.Count;
                for (var j = 1; j <= shapesCount; j++) {
                    var shape = slide.Shapes.Item(j);

                    // 检查是否有文本框架
                    if (this._hasTextFrame(shape)) {
                        try {
                            shape.TextFrame.TextRange.Font.Name = fontName;
                            slideModified++;
                            results.textRangesModified++;
                        } catch (fontErr) {
                            // 字体设置失败，继续处理下一个
                            pptLogger.warn('设置字体失败', {
                                slide: i,
                                shape: j,
                                error: fontErr.message
                            });
                        }
                    }
                }

                results.details.push({
                    slideIndex: i,
                    textRangesModified: slideModified
                });
                results.slidesProcessed++;
            }

            pptLogger.info('字体统一完成', results);
            return Response.success(results, null, startTime);

        } catch (err) {
            pptLogger.error('统一字体失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    /**
     * 应用配色方案
     * 将预设的配色方案应用到幻灯片
     *
     * @param {object} params - 参数对象
     * @param {string} params.scheme - 配色方案名称（business/tech/creative/minimal）
     * @param {number} params.slideIndex - 幻灯片索引（可选，不填则处理全部）
     * @returns {object} 标准响应对象
     */
    applyColorScheme: function(params) {
        var startTime = Date.now();
        pptLogger.info('开始应用配色方案', params);

        try {
            // 参数校验
            var validation = Validator.checkRequired(params, ['scheme']);
            if (!validation.valid) {
                return Response.error('PARAM_MISSING', null, startTime, {
                    missing: validation.missing
                });
            }

            // 获取配色方案
            var schemeName = params.scheme;
            var colorScheme = ColorSchemes[schemeName];

            if (!colorScheme) {
                pptLogger.error('未知的配色方案', { scheme: schemeName });
                return Response.error('PARAM_INVALID', null, startTime, {
                    message: '未知的配色方案: ' + schemeName,
                    availableSchemes: Object.keys(ColorSchemes)
                });
            }

            // 获取活动演示文稿
            var presentation = Application.ActivePresentation;
            if (!presentation) {
                return Response.error('DOCUMENT_NOT_FOUND', null, startTime);
            }

            if (presentation.ReadOnly) {
                return Response.error('DOCUMENT_READONLY', null, startTime);
            }

            // 确定要处理的幻灯片范围
            var startIndex = 1;
            var endIndex = presentation.Slides.Count;

            if (params.slideIndex) {
                if (params.slideIndex < 1 || params.slideIndex > presentation.Slides.Count) {
                    return Response.error('PARAM_INVALID', null, startTime, {
                        message: '幻灯片索引超出范围'
                    });
                }
                startIndex = params.slideIndex;
                endIndex = params.slideIndex;
            }

            var results = {
                scheme: schemeName,
                schemeName: colorScheme.name,
                colors: colorScheme,
                slidesProcessed: 0,
                details: []
            };

            // 遍历处理每张幻灯片
            for (var i = startIndex; i <= endIndex; i++) {
                var slide = presentation.Slides.Item(i);
                var slideResult = {
                    slideIndex: i,
                    backgroundSet: false,
                    titlesColored: 0,
                    textsColored: 0
                };

                // 1. 设置背景色
                try {
                    slide.FollowMasterBackground = 0;
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = this._parseColor(colorScheme.background);
                    slideResult.backgroundSet = true;
                } catch (bgErr) {
                    pptLogger.warn('设置背景色失败', { slide: i, error: bgErr.message });
                }

                // 2. 设置形状颜色
                var shapesCount = slide.Shapes.Count;
                for (var j = 1; j <= shapesCount; j++) {
                    var shape = slide.Shapes.Item(j);

                    if (this._hasTextFrame(shape)) {
                        try {
                            var textRange = shape.TextFrame.TextRange;
                            var isTitle = this._isLikelyTitle(shape);

                            if (isTitle) {
                                // 标题用主色
                                textRange.Font.Color.RGB = this._parseColor(colorScheme.primary);
                                slideResult.titlesColored++;
                            } else {
                                // 正文用文字色
                                textRange.Font.Color.RGB = this._parseColor(colorScheme.text);
                                slideResult.textsColored++;
                            }
                        } catch (colorErr) {
                            // 颜色设置失败，继续
                        }
                    }

                    // 如果是形状（非文本框），尝试设置填充色
                    try {
                        // 检查形状类型，避免处理图片等
                        if (shape.Type === 1 || shape.Type === 5) { // msoAutoShape or msoFreeform
                            if (shape.Fill) {
                                shape.Fill.ForeColor.RGB = this._parseColor(colorScheme.accent);
                            }
                        }
                    } catch (fillErr) {
                        // 填充设置失败，继续
                    }
                }

                results.details.push(slideResult);
                results.slidesProcessed++;
            }

            pptLogger.info('配色方案应用完成', results);
            return Response.success(results, null, startTime);

        } catch (err) {
            pptLogger.error('应用配色方案失败，艹', { error: err.message, stack: err.stack });
            return Response.fromException(err, null, startTime);
        }
    },

    // ==================== 私有辅助方法 ====================

    /**
     * 获取形状信息
     * 把形状的关键信息提取出来
     *
     * @private
     * @param {object} shape - 形状对象
     * @returns {object} 形状信息
     */
    _getShapeInfo: function(shape) {
        var info = {
            id: shape.Id,
            name: shape.Name,
            type: this._getShapeTypeName(shape.Type),
            position: {
                left: shape.Left,
                top: shape.Top,
                width: shape.Width,
                height: shape.Height
            }
        };

        // 如果有文本，获取文本内容
        if (this._hasTextFrame(shape)) {
            try {
                var text = shape.TextFrame.TextRange.Text;
                info.text = text ? text.substring(0, 200) : ''; // 最多取200字符
                info.hasText = true;
            } catch (textErr) {
                info.hasText = false;
            }
        } else {
            info.hasText = false;
        }

        return info;
    },

    /**
     * 检查形状是否有文本框架
     *
     * @private
     * @param {object} shape - 形状对象
     * @returns {boolean} 是否有文本框架
     */
    _hasTextFrame: function(shape) {
        try {
            // HasTextFrame 属性在某些形状上可能不存在
            if (shape.HasTextFrame !== undefined) {
                return shape.HasTextFrame === -1 || shape.HasTextFrame === true;
            }
            // 尝试访问 TextFrame 来判断
            return shape.TextFrame && shape.TextFrame.TextRange;
        } catch (err) {
            return false;
        }
    },

    /**
     * 判断形状是否可能是标题
     * 根据位置、大小和字号来推测
     *
     * @private
     * @param {object} shape - 形状对象
     * @returns {boolean} 是否可能是标题
     */
    _isLikelyTitle: function(shape) {
        try {
            // 在页面顶部
            if (shape.Top < 150) {
                // 检查字号
                if (this._hasTextFrame(shape)) {
                    var fontSize = shape.TextFrame.TextRange.Font.Size;
                    if (fontSize >= 24) {
                        return true;
                    }
                }
            }
            // 检查是否是占位符类型的标题
            if (shape.PlaceholderFormat) {
                var phType = shape.PlaceholderFormat.Type;
                // 1 = ppPlaceholderTitle, 3 = ppPlaceholderCenterTitle
                if (phType === 1 || phType === 3) {
                    return true;
                }
            }
        } catch (err) {
            // 判断失败，默认不是标题
        }
        return false;
    },

    /**
     * 查找幻灯片中的占位符
     *
     * @private
     * @param {object} slide - 幻灯片对象
     * @param {number} placeholderType - 占位符类型
     * @returns {object|null} 找到的形状或null
     */
    _findPlaceholder: function(slide, placeholderType) {
        try {
            var shapesCount = slide.Shapes.Count;
            for (var i = 1; i <= shapesCount; i++) {
                var shape = slide.Shapes.Item(i);
                try {
                    if (shape.PlaceholderFormat) {
                        if (shape.PlaceholderFormat.Type === placeholderType) {
                            return shape;
                        }
                    }
                } catch (phErr) {
                    // 这个形状不是占位符，继续
                }
            }
        } catch (err) {
            // 查找失败
        }
        return null;
    },

    /**
     * 获取布局类型名称
     *
     * @private
     * @param {number} layoutType - 布局类型常量
     * @returns {string} 布局名称
     */
    _getLayoutName: function(layoutType) {
        var layouts = {
            1: '标题幻灯片',
            2: '标题和文本',
            3: '两栏文本',
            4: '两列文本',
            5: '标题和图表',
            6: '图表和文本',
            7: '空白',
            8: '带标题的内容',
            9: '两列内容',
            10: '比较',
            11: '仅标题',
            12: '空白'
        };
        return layouts[layoutType] || '未知布局(' + layoutType + ')';
    },

    /**
     * 获取形状类型名称
     *
     * @private
     * @param {number} shapeType - 形状类型常量
     * @returns {string} 形状类型名称
     */
    _getShapeTypeName: function(shapeType) {
        var types = {
            1: '自选图形',
            2: '标注',
            3: '图表',
            4: '注释',
            5: '自由曲线',
            6: '组合',
            7: '嵌入OLE对象',
            8: '表单控件',
            9: '线条',
            10: '链接OLE对象',
            11: '链接图片',
            12: '媒体',
            13: '图片',
            14: '占位符',
            15: '脚本锚点',
            16: '表格',
            17: '文本框',
            18: '文本效果',
            19: 'Canvas',
            20: '图示',
            21: 'SmartArt',
            22: 'Slicer',
            23: 'WebVideo'
        };
        return types[shapeType] || '未知类型(' + shapeType + ')';
    },

    /**
     * 解析颜色值
     * 支持十六进制颜色字符串和RGB对象
     *
     * @private
     * @param {string|object|number} color - 颜色值
     * @returns {number|null} RGB颜色值（WPS格式：BGR）
     */
    _parseColor: function(color) {
        if (color === null || color === undefined) {
            return null;
        }

        // 如果已经是数字，直接返回
        if (typeof color === 'number') {
            return color;
        }

        // 如果是字符串，解析十六进制
        if (typeof color === 'string') {
            // 移除 # 前缀
            var hex = color.replace('#', '');

            // 确保是6位
            if (hex.length === 3) {
                hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
            }

            if (hex.length !== 6) {
                return null;
            }

            // 解析RGB
            var r = parseInt(hex.substring(0, 2), 16);
            var g = parseInt(hex.substring(2, 4), 16);
            var b = parseInt(hex.substring(4, 6), 16);

            // WPS使用BGR格式
            return b * 65536 + g * 256 + r;
        }

        // 如果是对象 {r, g, b}
        if (typeof color === 'object' && color.r !== undefined) {
            return color.b * 65536 + color.g * 256 + color.r;
        }

        return null;
    },

    /**
     * 获取当前配色方案信息
     *
     * @private
     * @param {object} presentation - 演示文稿对象
     * @returns {object|null} 配色信息
     */
    _getCurrentColorScheme: function(presentation) {
        try {
            // WPS的配色方案API可能和微软的略有不同
            // 这里返回一个简单的描述
            return {
                description: '当前使用的配色方案'
            };
        } catch (err) {
            return null;
        }
    }
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        PPTHandler: PPTHandler,
        ColorSchemes: ColorSchemes
    };
}
