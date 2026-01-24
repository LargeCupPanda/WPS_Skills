/**
 * Excel处理器单元测试 - 陈十三出品
 *
 * 艹，这个测试模块是用来测试excel-handler.js的
 * 测试getContext、setFormula等核心功能
 *
 * 使用方式：在浏览器中打开test-runner.html运行测试
 *
 * @author 陈十三
 * @date 2026-01-24
 */

(function(global) {
    'use strict';

    // 测试套件
    var ExcelHandlerTests = {

        // 测试名称
        name: 'ExcelHandler Tests',

        // 测试前准备
        beforeEach: function() {
            // 重置Mock状态
            WpsMock.reset();
        },

        // 测试后清理
        afterEach: function() {
            // 清理测试数据
        },

        // ==================== getContext 测试 ====================

        /**
         * 测试：正常获取工作簿上下文
         */
        testGetContextSuccess: function() {
            var result = ExcelHandler.getContext();

            TestRunner.assert(result.success === true, 'getContext应该返回成功');
            TestRunner.assert(result.data !== null, '应该返回上下文数据');
            TestRunner.assert(result.data.workbook !== undefined, '应该包含workbook信息');
            TestRunner.assert(result.data.workbook.name === 'TestWorkbook.xlsx', '工作簿名称应该正确');
            TestRunner.assert(result.data.activeSheet !== undefined, '应该包含activeSheet信息');
            TestRunner.assert(result.data.sheets !== undefined, '应该包含sheets列表');
            TestRunner.assert(Array.isArray(result.data.sheets), 'sheets应该是数组');
        },

        /**
         * 测试：没有活动工作簿时应该返回错误
         */
        testGetContextNoWorkbook: function() {
            // 设置没有活动文档
            WpsMock.setNoActiveDocument();

            var result = ExcelHandler.getContext();

            TestRunner.assert(result.success === false, '没有工作簿时应该返回失败');
            TestRunner.assert(result.error !== undefined, '应该包含错误信息');
            TestRunner.assert(result.error.code === 'DOCUMENT_NOT_FOUND', '错误码应该是DOCUMENT_NOT_FOUND');
        },

        /**
         * 测试：上下文应该包含正确的工作表数量
         */
        testGetContextSheetCount: function() {
            var result = ExcelHandler.getContext();

            TestRunner.assert(result.success === true, '应该成功获取上下文');
            TestRunner.assert(result.data.sheets.length === 3, '应该有3个工作表');
        },

        /**
         * 测试：上下文应该包含选区信息
         */
        testGetContextSelection: function() {
            var result = ExcelHandler.getContext();

            TestRunner.assert(result.success === true, '应该成功获取上下文');
            TestRunner.assert(result.data.selection !== undefined, '应该包含选区信息');
            TestRunner.assert(result.data.selection.address !== undefined, '选区应该有地址');
        },

        /**
         * 测试：上下文应该包含使用范围信息
         */
        testGetContextUsedRange: function() {
            var result = ExcelHandler.getContext();

            TestRunner.assert(result.success === true, '应该成功获取上下文');
            TestRunner.assert(result.data.usedRange !== undefined, '应该包含使用范围信息');
            TestRunner.assert(result.data.usedRange.rowCount > 0, '使用范围应该有行数');
            TestRunner.assert(result.data.usedRange.columnCount > 0, '使用范围应该有列数');
        },

        // ==================== setFormula 测试 ====================

        /**
         * 测试：正常设置公式
         */
        testSetFormulaSuccess: function() {
            var params = {
                range: 'A1',
                formula: '=SUM(B1:B10)'
            };

            var result = ExcelHandler.setFormula(params);

            TestRunner.assert(result.success === true, 'setFormula应该返回成功');
        },

        /**
         * 测试：缺少必填参数应该返回错误
         */
        testSetFormulaMissingParams: function() {
            var params = {
                range: 'A1'
                // 缺少formula参数
            };

            var result = ExcelHandler.setFormula(params);

            TestRunner.assert(result.success === false, '缺少参数时应该返回失败');
            TestRunner.assert(result.error.code === 'PARAM_MISSING', '错误码应该是PARAM_MISSING');
        },

        /**
         * 测试：公式格式不正确应该返回错误
         */
        testSetFormulaInvalidFormat: function() {
            var params = {
                range: 'A1',
                formula: 'SUM(B1:B10)' // 缺少=号
            };

            var result = ExcelHandler.setFormula(params);

            TestRunner.assert(result.success === false, '公式格式不正确时应该返回失败');
            TestRunner.assert(result.error.code === 'PARAM_FORMULA_INVALID', '错误码应该是PARAM_FORMULA_INVALID');
        },

        /**
         * 测试：没有活动工作簿时设置公式应该失败
         */
        testSetFormulaNoWorkbook: function() {
            WpsMock.setNoActiveDocument();

            var params = {
                range: 'A1',
                formula: '=SUM(B1:B10)'
            };

            var result = ExcelHandler.setFormula(params);

            TestRunner.assert(result.success === false, '没有工作簿时应该返回失败');
        },

        // ==================== 辅助方法测试 ====================

        /**
         * 测试：列号转字母
         */
        testColumnToLetter: function() {
            // 假设ExcelHandler有_columnToLetter方法
            if (typeof ExcelHandler._columnToLetter === 'function') {
                TestRunner.assert(ExcelHandler._columnToLetter(1) === 'A', '列1应该是A');
                TestRunner.assert(ExcelHandler._columnToLetter(26) === 'Z', '列26应该是Z');
                TestRunner.assert(ExcelHandler._columnToLetter(27) === 'AA', '列27应该是AA');
                TestRunner.assert(ExcelHandler._columnToLetter(52) === 'AZ', '列52应该是AZ');
            } else {
                TestRunner.skip('_columnToLetter方法不可访问');
            }
        },

        // ==================== 性能测试 ====================

        /**
         * 测试：getContext性能
         */
        testGetContextPerformance: function() {
            var startTime = Date.now();
            var iterations = 100;

            for (var i = 0; i < iterations; i++) {
                ExcelHandler.getContext();
            }

            var duration = Date.now() - startTime;
            var avgTime = duration / iterations;

            TestRunner.assert(avgTime < 50, 'getContext平均执行时间应该小于50ms，实际: ' + avgTime.toFixed(2) + 'ms');
        }
    };

    // 注册测试套件
    if (typeof TestRunner !== 'undefined') {
        TestRunner.registerSuite(ExcelHandlerTests);
    }

    // 导出供直接调用
    global.ExcelHandlerTests = ExcelHandlerTests;

})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this));
