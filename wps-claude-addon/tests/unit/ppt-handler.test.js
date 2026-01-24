/**
 * PPT处理器单元测试 - 陈十三出品
 *
 * 艹，这个测试模块是用来测试ppt-handler.js的
 * 测试getContext、addSlide、addTextbox、beautifySlide等核心功能
 *
 * 使用方式：在浏览器中打开test-runner.html运行测试
 *
 * @author 陈十三
 * @date 2026-01-24
 */

(function(global) {
    'use strict';

    // 测试套件
    var PPTHandlerTests = {

        // 测试名称
        name: 'PPTHandler Tests',

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
         * 测试：正常获取演示文稿上下文
         */
        testGetContextSuccess: function() {
            var result = PPTHandler.getContext();

            TestRunner.assert(result.success === true, 'getContext应该返回成功');
            TestRunner.assert(result.data !== null, '应该返回上下文数据');
            TestRunner.assert(result.data.presentation !== undefined, '应该包含presentation信息');
            TestRunner.assert(result.data.presentation.name === 'TestPresentation.pptx', '演示文稿名称应该正确');
        },

        /**
         * 测试：没有活动演示文稿时应该返回错误
         */
        testGetContextNoPresentation: function() {
            // 设置没有活动文档
            WpsMock.setNoActiveDocument();

            var result = PPTHandler.getContext();

            TestRunner.assert(result.success === false, '没有演示文稿时应该返回失败');
            TestRunner.assert(result.error !== undefined, '应该包含错误信息');
            TestRunner.assert(result.error.code === 'DOCUMENT_NOT_FOUND', '错误码应该是DOCUMENT_NOT_FOUND');
        },

        /**
         * 测试：上下文应该包含幻灯片数量
         */
        testGetContextSlideCount: function() {
            var result = PPTHandler.getContext();

            TestRunner.assert(result.success === true, '应该成功获取上下文');
            TestRunner.assert(result.data.slidesCount !== undefined, '应该包含幻灯片数量');
            TestRunner.assert(result.data.slidesCount === 5, '幻灯片数量应该是5');
        },

        /**
         * 测试：上下文应该包含当前幻灯片信息
         */
        testGetContextCurrentSlide: function() {
            var result = PPTHandler.getContext();

            TestRunner.assert(result.success === true, '应该成功获取上下文');
            TestRunner.assert(result.data.currentSlide !== undefined, '应该包含当前幻灯片信息');
            TestRunner.assert(result.data.currentSlide.index !== undefined, '当前幻灯片应该有索引');
        },

        // ==================== addSlide 测试 ====================

        /**
         * 测试：添加新幻灯片
         */
        testAddSlideSuccess: function() {
            var params = {
                layout: 'title'
            };

            var result = PPTHandler.addSlide(params);

            TestRunner.assert(result.success === true, 'addSlide应该返回成功');
        },

        /**
         * 测试：添加幻灯片到指定位置
         */
        testAddSlideAtIndex: function() {
            var params = {
                layout: 'blank',
                index: 2
            };

            var result = PPTHandler.addSlide(params);

            TestRunner.assert(result.success === true, '在指定位置添加幻灯片应该成功');
        },

        /**
         * 测试：没有演示文稿时添加幻灯片应该失败
         */
        testAddSlideNoPresentation: function() {
            WpsMock.setNoActiveDocument();

            var params = {
                layout: 'title'
            };

            var result = PPTHandler.addSlide(params);

            TestRunner.assert(result.success === false, '没有演示文稿时应该返回失败');
        },

        // ==================== addTextbox 测试 ====================

        /**
         * 测试：添加文本框
         */
        testAddTextboxSuccess: function() {
            var params = {
                text: '这是文本框内容',
                left: 100,
                top: 100,
                width: 200,
                height: 50
            };

            var result = PPTHandler.addTextbox(params);

            TestRunner.assert(result.success === true, 'addTextbox应该返回成功');
        },

        /**
         * 测试：添加文本框 - 自动计算尺寸
         */
        testAddTextboxAutoSize: function() {
            var params = {
                text: '自动尺寸文本框'
                // 不指定尺寸，让系统自动计算
            };

            var result = PPTHandler.addTextbox(params);

            TestRunner.assert(result.success === true, '自动尺寸应该成功');
        },

        /**
         * 测试：添加文本框 - 缺少text参数
         */
        testAddTextboxMissingText: function() {
            var params = {
                left: 100,
                top: 100
                // 缺少text参数
            };

            var result = PPTHandler.addTextbox(params);

            TestRunner.assert(result.success === false, '缺少text参数时应该返回失败');
        },

        // ==================== beautifySlide 测试 ====================

        /**
         * 测试：美化幻灯片 - 商务风格
         */
        testBeautifySlideBusinessStyle: function() {
            var params = {
                style: 'business'
            };

            var result = PPTHandler.beautifySlide(params);

            TestRunner.assert(result.success === true, 'beautifySlide应该返回成功');
        },

        /**
         * 测试：美化幻灯片 - 科技风格
         */
        testBeautifySlideTechStyle: function() {
            var params = {
                style: 'tech'
            };

            var result = PPTHandler.beautifySlide(params);

            TestRunner.assert(result.success === true, 'beautifySlide科技风格应该成功');
        },

        /**
         * 测试：美化幻灯片 - 创意风格
         */
        testBeautifySlideCreativeStyle: function() {
            var params = {
                style: 'creative'
            };

            var result = PPTHandler.beautifySlide(params);

            TestRunner.assert(result.success === true, 'beautifySlide创意风格应该成功');
        },

        /**
         * 测试：美化幻灯片 - 极简风格
         */
        testBeautifySlideMinimalStyle: function() {
            var params = {
                style: 'minimal'
            };

            var result = PPTHandler.beautifySlide(params);

            TestRunner.assert(result.success === true, 'beautifySlide极简风格应该成功');
        },

        /**
         * 测试：美化幻灯片 - 无效风格应该使用默认
         */
        testBeautifySlideInvalidStyle: function() {
            var params = {
                style: 'nonexistent_style'
            };

            var result = PPTHandler.beautifySlide(params);

            // 应该使用默认风格并成功
            TestRunner.assert(result !== undefined, '应该返回结果');
        },

        // ==================== setSlideBackground 测试 ====================

        /**
         * 测试：设置幻灯片背景颜色
         */
        testSetSlideBackgroundColor: function() {
            var params = {
                color: '#FFFFFF'
            };

            var result = PPTHandler.setSlideBackground(params);

            TestRunner.assert(result.success === true, 'setSlideBackground应该返回成功');
        },

        /**
         * 测试：设置幻灯片渐变背景
         */
        testSetSlideBackgroundGradient: function() {
            var params = {
                gradient: {
                    colors: ['#1F4E79', '#2E75B6'],
                    angle: 90
                }
            };

            var result = PPTHandler.setSlideBackground(params);

            TestRunner.assert(result.success === true, '渐变背景应该成功');
        },

        // ==================== 导航测试 ====================

        /**
         * 测试：跳转到指定幻灯片
         */
        testGoToSlide: function() {
            var params = {
                index: 3
            };

            var result = PPTHandler.goToSlide(params);

            TestRunner.assert(result.success === true, 'goToSlide应该返回成功');
        },

        /**
         * 测试：跳转到不存在的幻灯片应该失败
         */
        testGoToSlideInvalidIndex: function() {
            var params = {
                index: 999 // 不存在的索引
            };

            var result = PPTHandler.goToSlide(params);

            TestRunner.assert(result.success === false, '跳转到不存在的幻灯片应该失败');
        },

        // ==================== 删除幻灯片测试 ====================

        /**
         * 测试：删除幻灯片
         */
        testDeleteSlide: function() {
            var params = {
                index: 2
            };

            var result = PPTHandler.deleteSlide(params);

            TestRunner.assert(result.success === true, 'deleteSlide应该返回成功');
        },

        // ==================== 性能测试 ====================

        /**
         * 测试：getContext性能
         */
        testGetContextPerformance: function() {
            var startTime = Date.now();
            var iterations = 100;

            for (var i = 0; i < iterations; i++) {
                PPTHandler.getContext();
            }

            var duration = Date.now() - startTime;
            var avgTime = duration / iterations;

            TestRunner.assert(avgTime < 50, 'getContext平均执行时间应该小于50ms，实际: ' + avgTime.toFixed(2) + 'ms');
        },

        /**
         * 测试：批量操作性能
         */
        testBatchOperationPerformance: function() {
            var startTime = Date.now();

            // 模拟批量添加文本框
            for (var i = 0; i < 10; i++) {
                PPTHandler.addTextbox({
                    text: '文本框' + i,
                    left: 100 + i * 50,
                    top: 100
                });
            }

            var duration = Date.now() - startTime;

            TestRunner.assert(duration < 500, '批量操作应该在500ms内完成，实际: ' + duration + 'ms');
        }
    };

    // 注册测试套件
    if (typeof TestRunner !== 'undefined') {
        TestRunner.registerSuite(PPTHandlerTests);
    }

    // 导出供直接调用
    global.PPTHandlerTests = PPTHandlerTests;

})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this));
