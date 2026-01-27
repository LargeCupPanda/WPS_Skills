/**
 * PPT操作处理器 - Mac版
 * 使用WPS JavaScript API实现PPT操作
 * @author 老王的手下
 */

// 配色方案，别TM乱改这些颜色值
var COLOR_SCHEMES = {
    business: { title: 0x2F5496, body: 0x333333 },
    tech: { title: 0x00B0F0, body: 0x404040 },
    creative: { title: 0xFF6B6B, body: 0x4A4A4A },
    minimal: { title: 0x000000, body: 0x666666 }
};

/**
 * 获取当前演示文稿信息
 */
function getActivePresentation(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

        var slides = [];
        for (var i = 1; i <= pres.Slides.Count; i++) {
            var slide = pres.Slides.Item(i);
            var shapes = [];
            for (var j = 1; j <= slide.Shapes.Count; j++) {
                var shape = slide.Shapes.Item(j);
                var text = '';
                try {
                    if (shape.HasTextFrame && shape.TextFrame.HasText) {
                        var fullText = shape.TextFrame.TextRange.Text;
                        text = fullText.substring(0, Math.min(50, fullText.length));
                    }
                } catch (e) {}
                shapes.push({ name: shape.Name, type: shape.Type, text: text });
            }
            slides.push({ index: i, shapeCount: slide.Shapes.Count, shapes: shapes });
        }

        return {
            success: true,
            data: {
                name: pres.Name,
                path: pres.FullName,
                slideCount: pres.Slides.Count,
                slides: slides
            }
        };
    } catch (e) {
        return { success: false, error: 'getActivePresentation失败: ' + e.message };
    }
}

/**
 * 添加幻灯片
 * @param {Object} params - { layout: 'title'|'title_content'|'blank'|'two_column', position: number, title: string }
 */
function addSlide(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

        var layouts = { title: 1, title_content: 2, blank: 12, two_column: 3 };
        var layoutType = layouts[params.layout] || 2;
        var position = params.position || (pres.Slides.Count + 1);

        var slide = pres.Slides.Add(position, layoutType);

        if (params.title && slide.Shapes.HasTitle) {
            slide.Shapes.Title.TextFrame.TextRange.Text = params.title;
        }

        return { success: true, data: { slideIndex: position } };
    } catch (e) {
        return { success: false, error: 'addSlide失败: ' + e.message };
    }
}

/**
 * 添加文本框
 * @param {Object} params - { slideIndex, text, left, top, width, height, fontSize, fontName }
 */
function addTextBox(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

        var slideIndex = params.slideIndex || app.ActiveWindow.Selection.SlideRange.SlideIndex;
        var slide = pres.Slides.Item(slideIndex);

        var left = params.left || 100;
        var top = params.top || 100;
        var width = params.width || 400;
        var height = params.height || 50;

        // 1 = msoTextOrientationHorizontal
        var shape = slide.Shapes.AddTextbox(1, left, top, width, height);
        shape.TextFrame.TextRange.Text = params.text || '';

        if (params.fontSize) {
            shape.TextFrame.TextRange.Font.Size = params.fontSize;
        }
        if (params.fontName) {
            shape.TextFrame.TextRange.Font.Name = params.fontName;
        }

        return { success: true, data: { shapeName: shape.Name } };
    } catch (e) {
        return { success: false, error: 'addTextBox失败: ' + e.message };
    }
}

/**
 * 设置幻灯片标题
 * @param {Object} params - { slideIndex, title }
 */
function setSlideTitle(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

        var slide = pres.Slides.Item(params.slideIndex);
        if (slide.Shapes.HasTitle) {
            slide.Shapes.Title.TextFrame.TextRange.Text = params.title;
        }

        return { success: true, data: {} };
    } catch (e) {
        return { success: false, error: 'setSlideTitle失败: ' + e.message };
    }
}

/**
 * 统一字体
 * @param {Object} params - { fontName }
 */
function unifyFont(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

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
        return { success: false, error: 'unifyFont失败: ' + e.message };
    }
}

/**
 * 美化幻灯片
 * @param {Object} params - { slideIndex, style: 'business'|'tech'|'creative'|'minimal' }
 */
function beautifySlide(params) {
    try {
        var app = Application;
        var pres = app.ActivePresentation;
        if (!pres) {
            return { success: false, error: '没有打开的演示文稿' };
        }

        var slideIndex = params.slideIndex || app.ActiveWindow.Selection.SlideRange.SlideIndex;
        var slide = pres.Slides.Item(slideIndex);

        var scheme = COLOR_SCHEMES[params.style] || COLOR_SCHEMES.business;
        var count = 0;

        for (var j = 1; j <= slide.Shapes.Count; j++) {
            var shape = slide.Shapes.Item(j);
            try {
                if (shape.HasTextFrame && shape.TextFrame.HasText) {
                    var textRange = shape.TextFrame.TextRange;
                    // 字号>=24视为标题，否则为正文
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
        return { success: false, error: 'beautifySlide失败: ' + e.message };
    }
}

/**
 * PPT操作路由
 */
function handlePPT(action, params) {
    switch (action) {
        case 'getActivePresentation':
            return getActivePresentation(params);
        case 'addSlide':
            return addSlide(params);
        case 'addTextBox':
            return addTextBox(params);
        case 'setSlideTitle':
            return setSlideTitle(params);
        case 'unifyFont':
            return unifyFont(params);
        case 'beautifySlide':
            return beautifySlide(params);
        default:
            return { success: false, error: '未知的PPT操作: ' + action };
    }
}

// 导出
module.exports = {
    handlePPT: handlePPT,
    getActivePresentation: getActivePresentation,
    addSlide: addSlide,
    addTextBox: addTextBox,
    setSlideTitle: setSlideTitle,
    unifyFont: unifyFont,
    beautifySlide: beautifySlide
};
