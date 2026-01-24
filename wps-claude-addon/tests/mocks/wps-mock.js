/**
 * WPS API Mock对象 - 陈十三出品
 *
 * 艹，这个SB文件是用来模拟WPS Office的JavaScript API的
 * 因为测试环境没有真正的WPS，所以需要Mock这些API
 *
 * @author 陈十三
 * @date 2026-01-24
 */

(function(global) {
    'use strict';

    /**
     * 模拟的工作表对象
     */
    function MockSheet(name, index) {
        this.Name = name;
        this.Index = index;
        this.Visible = true;
        this._cells = {};
        this._usedRange = null;

        // 模拟UsedRange
        this.UsedRange = {
            Address: function() { return '$A$1:$C$10'; },
            Rows: { Count: 10 },
            Columns: { Count: 3 },
            Row: 1,
            Column: 1
        };

        // 模拟Cells
        this.Cells = {
            Item: function(row, col) {
                return new MockCell(row, col);
            }
        };

        // 模拟Range
        this.Range = function(start, end) {
            return new MockRange(start, end);
        };
    }

    /**
     * 模拟的单元格对象
     */
    function MockCell(row, col) {
        this.Row = row;
        this.Column = col;
        this.Value2 = null;
        this.Formula = '';
        this.Text = '';
    }

    /**
     * 模拟的范围对象
     */
    function MockRange(start, end) {
        this._start = start;
        this._end = end;
        this.Value2 = [['A', 'B', 'C']];

        this.Address = function() {
            if (this._end) {
                return this._start.Address() + ':' + this._end.Address();
            }
            return '$A$1:$C$1';
        };

        this.Rows = { Count: 1 };
        this.Columns = { Count: 3 };
    }

    /**
     * 模拟的工作簿对象
     */
    function MockWorkbook(name) {
        this.Name = name || 'TestWorkbook.xlsx';
        this.Path = 'C:\\Users\\Test';
        this.FullName = this.Path + '\\' + this.Name;
        this.Saved = true;
        this.ReadOnly = false;
        this._sheets = [
            new MockSheet('Sheet1', 1),
            new MockSheet('Sheet2', 2),
            new MockSheet('Sheet3', 3)
        ];

        this.ActiveSheet = this._sheets[0];

        // 模拟Sheets集合
        this.Sheets = {
            Count: this._sheets.length,
            Item: function(index) {
                return this._parent._sheets[index - 1];
            }
        };
        this.Sheets._parent = this;
    }

    /**
     * 模拟的文档对象 (Word)
     */
    function MockDocument(name) {
        this.Name = name || 'TestDocument.docx';
        this.Path = 'C:\\Users\\Test';
        this.FullName = this.Path + '\\' + this.Name;
        this.Saved = true;
        this.ReadOnly = false;

        // 模拟段落
        this.Paragraphs = {
            Count: 10,
            Item: function(index) {
                return {
                    Range: { Text: '这是第' + index + '段落的测试文本。' },
                    Style: { NameLocal: index <= 2 ? '标题 1' : '正文' }
                };
            }
        };

        // 模拟内容范围
        this.Content = {
            Text: '这是测试文档的全部内容。',
            Characters: { Count: 500 },
            Words: { Count: 100 }
        };

        // 模拟统计功能
        this.ComputeStatistics = function(type) {
            var stats = {
                0: 100,    // wdStatisticWords
                2: 5,      // wdStatisticPages
                3: 500,    // wdStatisticCharacters
                4: 10,     // wdStatisticParagraphs
                5: 600     // wdStatisticCharactersWithSpaces
            };
            return stats[type] || 0;
        };
    }

    /**
     * 模拟的演示文稿对象 (PPT)
     */
    function MockPresentation(name) {
        this.Name = name || 'TestPresentation.pptx';
        this.Path = 'C:\\Users\\Test';
        this.FullName = this.Path + '\\' + this.Name;
        this.Saved = true;
        this.ReadOnly = false;

        // 模拟幻灯片集合
        this.Slides = {
            Count: 5,
            Item: function(index) {
                return new MockSlide(index);
            },
            Add: function(index, layout) {
                return new MockSlide(index);
            }
        };
    }

    /**
     * 模拟的幻灯片对象
     */
    function MockSlide(index) {
        this.SlideIndex = index;
        this.SlideNumber = index;
        this.Layout = 1; // ppLayoutTitle
        this.Hidden = false;

        // 模拟形状集合
        this.Shapes = {
            Count: 2,
            Item: function(index) {
                return new MockShape(index);
            },
            AddTextbox: function(orientation, left, top, width, height) {
                return new MockShape(this.Count + 1);
            },
            Title: {
                TextFrame: {
                    TextRange: { Text: '幻灯片标题' }
                }
            }
        };
    }

    /**
     * 模拟的形状对象
     */
    function MockShape(index) {
        this.Id = index;
        this.Name = 'Shape' + index;
        this.Left = 0;
        this.Top = 0;
        this.Width = 100;
        this.Height = 50;
        this.Visible = true;

        this.TextFrame = {
            TextRange: {
                Text: '形状中的文本',
                Font: {
                    Name: '微软雅黑',
                    Size: 12,
                    Bold: false,
                    Italic: false,
                    Color: { RGB: 0 }
                }
            },
            HasText: true
        };

        this.Fill = {
            ForeColor: { RGB: 16777215 },
            Solid: function() {}
        };

        this.Line = {
            ForeColor: { RGB: 0 },
            Weight: 1
        };
    }

    /**
     * 模拟的Selection对象
     */
    function MockSelection() {
        this.Value2 = 'Selected Value';
        this.Formula = '';
        this.Text = 'Selected Value';

        this.Address = function() {
            return '$A$1';
        };

        this.Rows = { Count: 1 };
        this.Columns = { Count: 1 };
    }

    /**
     * 模拟的Window对象
     */
    function MockWindow() {
        this.View = {
            Slide: new MockSlide(1)
        };
    }

    /**
     * 模拟的Application对象 - WPS全局对象
     */
    var MockApplication = {
        // Excel相关
        ActiveWorkbook: new MockWorkbook(),
        Selection: new MockSelection(),

        // Word相关
        ActiveDocument: new MockDocument(),

        // PPT相关
        ActivePresentation: new MockPresentation(),
        ActiveWindow: new MockWindow(),

        // 通用方法
        Workbooks: {
            Add: function() { return new MockWorkbook('NewWorkbook.xlsx'); },
            Open: function(path) { return new MockWorkbook(path.split('\\').pop()); }
        },

        Documents: {
            Add: function() { return new MockDocument('NewDocument.docx'); },
            Open: function(path) { return new MockDocument(path.split('\\').pop()); }
        },

        Presentations: {
            Add: function() { return new MockPresentation('NewPresentation.pptx'); },
            Open: function(path) { return new MockPresentation(path.split('\\').pop()); }
        }
    };

    /**
     * 模拟的Logger类
     */
    function MockLogger(name) {
        this.name = name;
        this.logs = [];

        this.info = function(message, data) {
            this.logs.push({ level: 'info', message: message, data: data });
        };

        this.debug = function(message, data) {
            this.logs.push({ level: 'debug', message: message, data: data });
        };

        this.warn = function(message, data) {
            this.logs.push({ level: 'warn', message: message, data: data });
        };

        this.error = function(message, data) {
            this.logs.push({ level: 'error', message: message, data: data });
        };

        this.clear = function() {
            this.logs = [];
        };
    }

    /**
     * 模拟的Response工具类
     */
    var MockResponse = {
        success: function(data, requestId, startTime) {
            return {
                success: true,
                data: data,
                requestId: requestId,
                duration: startTime ? Date.now() - startTime : 0
            };
        },

        error: function(code, requestId, startTime, details) {
            return {
                success: false,
                error: {
                    code: code,
                    details: details
                },
                requestId: requestId,
                duration: startTime ? Date.now() - startTime : 0
            };
        },

        fromException: function(err, requestId, startTime) {
            return {
                success: false,
                error: {
                    code: 'EXCEPTION',
                    message: err.message,
                    stack: err.stack
                },
                requestId: requestId,
                duration: startTime ? Date.now() - startTime : 0
            };
        }
    };

    /**
     * 模拟的Validator工具类
     */
    var MockValidator = {
        checkRequired: function(params, required) {
            var missing = [];
            for (var i = 0; i < required.length; i++) {
                var key = required[i];
                if (params[key] === undefined || params[key] === null) {
                    missing.push(key);
                }
            }
            return {
                valid: missing.length === 0,
                missing: missing
            };
        },

        isValidFormula: function(formula) {
            return typeof formula === 'string' && formula.charAt(0) === '=';
        },

        isValidRange: function(range) {
            // 简单的范围验证：A1, B2:C3等格式
            var rangePattern = /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i;
            return rangePattern.test(range);
        }
    };

    // 导出到全局
    global.Application = MockApplication;
    global.Logger = MockLogger;
    global.Response = MockResponse;
    global.Validator = MockValidator;

    // 导出Mock类供测试使用
    global.WpsMock = {
        Application: MockApplication,
        Workbook: MockWorkbook,
        Sheet: MockSheet,
        Cell: MockCell,
        Range: MockRange,
        Document: MockDocument,
        Presentation: MockPresentation,
        Slide: MockSlide,
        Shape: MockShape,
        Selection: MockSelection,
        Window: MockWindow,
        Logger: MockLogger,
        Response: MockResponse,
        Validator: MockValidator,

        // 重置所有Mock状态
        reset: function() {
            MockApplication.ActiveWorkbook = new MockWorkbook();
            MockApplication.ActiveDocument = new MockDocument();
            MockApplication.ActivePresentation = new MockPresentation();
            MockApplication.Selection = new MockSelection();
            MockApplication.ActiveWindow = new MockWindow();
        },

        // 设置没有活动文档的场景
        setNoActiveDocument: function() {
            MockApplication.ActiveWorkbook = null;
            MockApplication.ActiveDocument = null;
            MockApplication.ActivePresentation = null;
        }
    };

})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this));
