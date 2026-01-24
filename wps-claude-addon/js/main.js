/**
 * WPS Claude 智能助手 - 主入口文件
 *
 * 艹，这是整个加载项的灵魂所在
 * 负责初始化所有模块、启动HTTP服务、注册路由
 * 没有这个文件，其他模块就是一堆废铁
 *
 * 主要职责：
 * 1. 加载所有依赖模块（Logger、Response、HttpServer、Router、Handlers）
 * 2. 初始化HTTP服务器，监听23333端口
 * 3. 注册所有API路由（Excel、Word、PPT操作）
 * 4. 实现ribbon.xml中定义的回调函数
 * 5. 实现WPS加载项标准生命周期函数
 *
 * @author 老王 (隔壁老王，五金店兼职程序员)
 * @date 2026-01-24
 * @version 1.0.0
 */

// ==================== 全局变量定义 ====================
// 这些变量是全局的，整个加载项都能访问，别他妈乱改

/**
 * 主日志记录器
 * 整个加载项的日志都从这里输出
 */
var mainLogger = null;

/**
 * HTTP服务器实例
 * 负责监听来自MCP Server的请求
 */
var httpServer = null;

/**
 * 路由器实例
 * 负责把请求分发到对应的处理器
 */
var router = null;

/**
 * 服务器配置
 * 从manifest.xml读取，或使用默认值
 */
var serverConfig = {
    port: 23333,
    logLevel: 'info'
};

/**
 * 加载项状态
 * 用于ribbon按钮的状态显示
 */
var addonStatus = {
    initialized: false,
    serverRunning: false,
    lastError: null
};

// ==================== 模块加载 ====================
// WPS加载项不支持ES6的import，也不支持CommonJS的require
// 所有依赖模块都通过script标签加载，变量直接挂在全局作用域
// 这里检查一下关键模块是否存在

/**
 * 检查依赖模块是否加载
 * 如果关键模块没加载，这加载项就是个废物
 *
 * @returns {boolean} 是否所有依赖都加载成功
 */
function checkDependencies() {
    var dependencies = [
        { name: 'Logger', obj: typeof Logger },
        { name: 'LogLevel', obj: typeof LogLevel },
        { name: 'Response', obj: typeof Response },
        { name: 'ErrorCodes', obj: typeof ErrorCodes },
        { name: 'Validator', obj: typeof Validator },
        { name: 'HttpServer', obj: typeof HttpServer },
        { name: 'Router', obj: typeof Router },
        { name: 'ExcelHandler', obj: typeof ExcelHandler },
        { name: 'WordHandler', obj: typeof WordHandler },
        { name: 'PPTHandler', obj: typeof PPTHandler },
        { name: 'CommonHandler', obj: typeof CommonHandler }
    ];

    var allLoaded = true;
    var missing = [];

    for (var i = 0; i < dependencies.length; i++) {
        var dep = dependencies[i];
        if (dep.obj === 'undefined') {
            allLoaded = false;
            missing.push(dep.name);
        }
    }

    if (!allLoaded) {
        console.error('[Main] 依赖模块加载失败，缺少: ' + missing.join(', '));
    }

    return allLoaded;
}

// ==================== 初始化函数 ====================

/**
 * 初始化加载项
 * 这是整个加载项启动的入口，OnAddinLoad会调用这个
 *
 * @returns {boolean} 是否初始化成功
 */
function initAddon() {
    try {
        console.log('[Main] 开始初始化加载项...');

        // 1. 先检查依赖
        if (!checkDependencies()) {
            console.error('[Main] 艹，依赖检查失败，加载项无法启动');
            return false;
        }

        // 2. 初始化主日志记录器
        mainLogger = new Logger('Main');
        mainLogger.setLevel(LogLevel.INFO);
        mainLogger.info('===========================================');
        mainLogger.info('WPS Claude 智能助手 v1.0.0 正在启动...');
        mainLogger.info('作者: 老王出品，必属精品');
        mainLogger.info('===========================================');

        // 3. 读取配置（如果有的话）
        loadConfig();

        // 4. 初始化路由器
        initRouter();

        // 5. 初始化HTTP服务器
        initHttpServer();

        // 6. 标记初始化完成
        addonStatus.initialized = true;
        mainLogger.info('加载项初始化完成，等待启动HTTP服务...');

        // 7. 自动启动服务器（可选）
        // 默认不自动启动，让用户手动点击启动按钮
        // startServer();

        return true;

    } catch (error) {
        console.error('[Main] 初始化失败: ' + error.message);
        addonStatus.lastError = error.message;
        return false;
    }
}

/**
 * 加载配置
 * 从manifest.xml的settings中读取配置
 */
function loadConfig() {
    try {
        // WPS加载项可能有自己的配置读取API
        // 这里先用默认值，后续可以扩展
        if (typeof Application !== 'undefined' && Application.GetAddinSetting) {
            var portSetting = Application.GetAddinSetting('httpPort');
            var logLevelSetting = Application.GetAddinSetting('logLevel');

            if (portSetting) {
                serverConfig.port = parseInt(portSetting, 10) || 23333;
            }
            if (logLevelSetting) {
                serverConfig.logLevel = logLevelSetting;
            }
        }

        mainLogger.info('配置加载完成', serverConfig);

        // 根据配置设置日志级别
        switch (serverConfig.logLevel.toLowerCase()) {
            case 'debug':
                mainLogger.setLevel(LogLevel.DEBUG);
                break;
            case 'info':
                mainLogger.setLevel(LogLevel.INFO);
                break;
            case 'warn':
                mainLogger.setLevel(LogLevel.WARN);
                break;
            case 'error':
                mainLogger.setLevel(LogLevel.ERROR);
                break;
        }

    } catch (error) {
        mainLogger.warn('配置加载失败，使用默认配置', { error: error.message });
    }
}

/**
 * 初始化路由器
 * 注册所有API端点
 */
function initRouter() {
    mainLogger.info('正在初始化路由器...');

    router = new Router(mainLogger.createChild('Router'));

    // 注册Excel相关路由
    registerExcelRoutes();

    // 注册Word相关路由（暂时用占位符）
    registerWordRoutes();

    // 注册PPT相关路由（暂时用占位符）
    registerPptRoutes();

    // 注册通用路由（格式转换等）
    registerCommonRoutes();

    mainLogger.info('路由器初始化完成，已注册 ' + router.getActions().length + ' 个路由');
}

/**
 * 注册Excel相关路由
 * 把ExcelHandler的方法都挂到路由上
 */
function registerExcelRoutes() {
    mainLogger.debug('注册Excel路由...');

    // excel.getContext - 获取工作簿上下文
    router.register('excel.getContext', function(params, requestId, startTime) {
        return ExcelHandler.getContext();
    });

    // excel.setFormula - 设置公式
    router.register('excel.setFormula', function(params, requestId, startTime) {
        return ExcelHandler.setFormula(params);
    });

    // excel.diagnoseFormula - 诊断公式错误
    router.register('excel.diagnoseFormula', function(params, requestId, startTime) {
        return ExcelHandler.diagnoseFormula(params);
    });

    // excel.cleanData - 数据清洗
    router.register('excel.cleanData', function(params, requestId, startTime) {
        return ExcelHandler.cleanData(params);
    });

    // excel.readRange - 读取范围数据
    router.register('excel.readRange', function(params, requestId, startTime) {
        return ExcelHandler.readRange(params);
    });

    // excel.writeRange - 写入范围数据
    router.register('excel.writeRange', function(params, requestId, startTime) {
        return ExcelHandler.writeRange(params);
    });

    // excel.createChart - 创建图表（刘大炮出品）
    router.register('excel.createChart', function(params, requestId, startTime) {
        return ExcelHandler.createChart(params);
    });

    // excel.updateChart - 更新图表（刘大炮出品）
    router.register('excel.updateChart', function(params, requestId, startTime) {
        return ExcelHandler.updateChart(params);
    });

    // excel.createPivotTable - 创建透视表（马铁锤出品）
    router.register('excel.createPivotTable', function(params, requestId, startTime) {
        return ExcelHandler.createPivotTable(params);
    });

    // excel.updatePivotTable - 更新透视表（马铁锤出品）
    router.register('excel.updatePivotTable', function(params, requestId, startTime) {
        return ExcelHandler.updatePivotTable(params);
    });

    mainLogger.debug('Excel路由注册完成');
}

/**
 * 注册Word相关路由
 * 老王终于把WordHandler写好了，把这些SB占位符干掉
 */
function registerWordRoutes() {
    mainLogger.debug('注册Word路由...');

    // word.getContext - 获取文档上下文
    router.register('word.getContext', function(params, requestId, startTime) {
        return WordHandler.getContext();
    });

    // word.insertText - 插入文本
    router.register('word.insertText', function(params, requestId, startTime) {
        return WordHandler.insertText(params);
    });

    // word.formatText - 格式化文本（实际调用setFont）
    router.register('word.formatText', function(params, requestId, startTime) {
        return WordHandler.setFont(params);
    });

    mainLogger.debug('Word路由注册完成');
}

/**
 * 注册PPT相关路由
 * 老王接手了赵六的烂摊子，把PPT处理器搞定了
 */
function registerPptRoutes() {
    mainLogger.debug('注册PPT路由...');

    // ppt.getContext - 获取演示文稿上下文
    router.register('ppt.getContext', function(params, requestId, startTime) {
        return PPTHandler.getContext();
    });

    // ppt.addSlide - 添加幻灯片
    router.register('ppt.addSlide', function(params, requestId, startTime) {
        return PPTHandler.addSlide(params);
    });

    // ppt.addTextBox - 添加文本框
    router.register('ppt.addTextBox', function(params, requestId, startTime) {
        return PPTHandler.addTextBox(params);
    });

    // ppt.beautifySlide - 美化幻灯片
    router.register('ppt.beautifySlide', function(params, requestId, startTime) {
        return PPTHandler.beautifySlide(params);
    });

    // ppt.unifyFont - 统一字体
    router.register('ppt.unifyFont', function(params, requestId, startTime) {
        return PPTHandler.unifyFont(params);
    });

    // ppt.applyColorScheme - 应用配色方案
    router.register('ppt.applyColorScheme', function(params, requestId, startTime) {
        return PPTHandler.applyColorScheme(params);
    });

    mainLogger.debug('PPT路由注册完成，赵六的锅终于填上了');
}

/**
 * 注册通用路由
 * 跨应用的通用功能，比如格式转换、PDF导出
 * 孙二狗出品
 */
function registerCommonRoutes() {
    mainLogger.debug('注册通用路由...');

    // common.convertToPdf - 转换为PDF
    router.register('common.convertToPdf', function(params, requestId, startTime) {
        return CommonHandler.convertToPdf(params);
    });

    // common.convertFormat - 格式互转
    router.register('common.convertFormat', function(params, requestId, startTime) {
        return CommonHandler.convertFormat(params);
    });

    mainLogger.debug('通用路由注册完成，二狗干得不错');
}

/**
 * 初始化HTTP服务器
 * 创建服务器实例并设置路由器
 */
function initHttpServer() {
    mainLogger.info('正在初始化HTTP服务器...');

    var serverLogger = mainLogger.createChild('HttpServer');
    httpServer = new HttpServer(serverConfig.port, serverLogger);
    httpServer.setRouter(router);

    mainLogger.info('HTTP服务器初始化完成，端口: ' + serverConfig.port);
}

/**
 * 启动HTTP服务器
 *
 * @returns {boolean} 是否启动成功
 */
function startServer() {
    if (!httpServer) {
        mainLogger.error('HTTP服务器未初始化，无法启动');
        return false;
    }

    try {
        var success = httpServer.start();
        if (success) {
            addonStatus.serverRunning = true;
            mainLogger.info('HTTP服务器启动成功！监听端口: ' + serverConfig.port);

            // 通知用户
            showNotification('Claude助手已启动', '服务已在端口 ' + serverConfig.port + ' 上运行');
        } else {
            addonStatus.serverRunning = false;
            addonStatus.lastError = '服务器启动失败';
            mainLogger.error('HTTP服务器启动失败');
        }
        return success;

    } catch (error) {
        addonStatus.serverRunning = false;
        addonStatus.lastError = error.message;
        mainLogger.error('启动服务器时发生异常', { error: error.message });
        return false;
    }
}

/**
 * 停止HTTP服务器
 *
 * @returns {boolean} 是否停止成功
 */
function stopServer() {
    if (!httpServer) {
        mainLogger.warn('HTTP服务器未初始化');
        return true;
    }

    try {
        var success = httpServer.stop();
        if (success) {
            addonStatus.serverRunning = false;
            mainLogger.info('HTTP服务器已停止');
            showNotification('Claude助手已停止', '服务已关闭');
        }
        return success;

    } catch (error) {
        mainLogger.error('停止服务器时发生异常', { error: error.message });
        return false;
    }
}

/**
 * 显示通知消息
 * WPS可能有自己的通知API，这里做个封装
 *
 * @param {string} title - 通知标题
 * @param {string} message - 通知内容
 */
function showNotification(title, message) {
    try {
        if (typeof Application !== 'undefined' && Application.Alert) {
            // WPS的弹窗API
            Application.Alert(title + '\n\n' + message);
        } else if (typeof alert !== 'undefined') {
            alert(title + '\n' + message);
        }
    } catch (e) {
        mainLogger.debug('无法显示通知', { title: title, message: message });
    }
}

// ==================== Ribbon回调函数 ====================
// 这些函数在ribbon.xml中定义，用于响应用户点击功能区按钮

/**
 * 检查连接状态回调
 * 对应ribbon.xml中的btnStatus按钮
 *
 * @param {object} control - 控件对象
 */
function ribbonCheckStatus(control) {
    mainLogger.info('用户点击了检查状态按钮');

    try {
        var status = getServerStatus();
        var statusText = '=== Claude助手状态 ===\n\n';

        statusText += '初始化状态: ' + (addonStatus.initialized ? '已完成' : '未完成') + '\n';
        statusText += '服务状态: ' + (status.isRunning ? '运行中' : '已停止') + '\n';

        if (status.isRunning) {
            statusText += '监听端口: ' + status.port + '\n';
            statusText += '运行时长: ' + status.runTime + '\n';
            statusText += '请求总数: ' + status.requestCount + '\n';
        }

        if (addonStatus.lastError) {
            statusText += '\n最近错误: ' + addonStatus.lastError;
        }

        // 获取WPS信息
        statusText += '\n\n=== WPS信息 ===\n';
        try {
            if (typeof Application !== 'undefined') {
                statusText += '应用名称: ' + (Application.Name || 'WPS Office') + '\n';
                statusText += '版本: ' + (Application.Version || '未知') + '\n';

                // 尝试获取当前文档信息
                if (Application.ActiveWorkbook) {
                    statusText += '当前文档: ' + Application.ActiveWorkbook.Name + ' (Excel)\n';
                } else if (Application.ActiveDocument) {
                    statusText += '当前文档: ' + Application.ActiveDocument.Name + ' (Word)\n';
                } else if (Application.ActivePresentation) {
                    statusText += '当前文档: ' + Application.ActivePresentation.Name + ' (PPT)\n';
                } else {
                    statusText += '当前文档: 无\n';
                }
            }
        } catch (e) {
            statusText += 'WPS信息获取失败\n';
        }

        showNotification('Claude助手', statusText);

    } catch (error) {
        mainLogger.error('检查状态失败', { error: error.message });
        showNotification('错误', '检查状态失败: ' + error.message);
    }
}

/**
 * 启动服务器回调
 * 对应ribbon.xml中的btnStartServer按钮
 *
 * @param {object} control - 控件对象
 */
function ribbonStartServer(control) {
    mainLogger.info('用户点击了启动服务按钮');

    if (!addonStatus.initialized) {
        showNotification('错误', '加载项尚未初始化，请稍候...');
        return;
    }

    if (addonStatus.serverRunning) {
        showNotification('提示', '服务已经在运行中，不需要重复启动');
        return;
    }

    var success = startServer();
    if (!success) {
        showNotification('错误', '服务启动失败，请查看日志获取详情');
    }
}

/**
 * 停止服务器回调
 * 对应ribbon.xml中的btnStopServer按钮
 *
 * @param {object} control - 控件对象
 */
function ribbonStopServer(control) {
    mainLogger.info('用户点击了停止服务按钮');

    if (!addonStatus.serverRunning) {
        showNotification('提示', '服务当前未运行');
        return;
    }

    var success = stopServer();
    if (!success) {
        showNotification('错误', '服务停止失败');
    }
}

/**
 * 显示日志回调
 * 对应ribbon.xml中的btnShowLog按钮
 *
 * @param {object} control - 控件对象
 */
function ribbonShowLog(control) {
    mainLogger.info('用户点击了查看日志按钮');

    try {
        var logText = '=== 最近日志记录 ===\n\n';

        if (mainLogger && mainLogger.getHistory) {
            var logs = mainLogger.getHistory(50); // 获取最近50条日志

            if (logs.length === 0) {
                logText += '暂无日志记录\n';
            } else {
                for (var i = logs.length - 1; i >= 0; i--) {
                    var log = logs[i];
                    logText += '[' + log.timestamp + '] ';
                    logText += '[' + log.level + '] ';
                    logText += '[' + log.module + '] ';
                    logText += log.message + '\n';
                }
            }
        } else {
            logText += '日志记录器未初始化\n';
        }

        // 日志太长可能显示不全，考虑用其他方式展示
        // 这里暂时用弹窗
        showNotification('运行日志', logText);

    } catch (error) {
        showNotification('错误', '获取日志失败: ' + error.message);
    }
}

/**
 * 获取状态图标回调
 * 对应ribbon.xml中btnStatus的getImage属性
 * 根据服务状态返回不同的图标
 *
 * @param {object} control - 控件对象
 * @returns {string} 图标名称
 */
function getStatusImage(control) {
    // 返回WPS内置图标名称
    // 绿色表示运行中，红色表示停止
    if (addonStatus.serverRunning) {
        return 'HappyFace'; // 运行中 - 笑脸图标
    } else if (addonStatus.initialized) {
        return 'SadFace'; // 已初始化但未运行 - 悲伤图标
    } else {
        return 'Help'; // 未初始化 - 问号图标
    }
}

/**
 * 获取服务器状态
 *
 * @returns {object} 状态信息
 */
function getServerStatus() {
    if (httpServer) {
        return httpServer.getStatus();
    }

    return {
        isRunning: false,
        port: serverConfig.port,
        startTime: null,
        runTime: '0s',
        requestCount: 0
    };
}

// ==================== WPS加载项生命周期函数 ====================
// 这些是WPS加载项框架定义的标准生命周期钩子
// 必须全局可访问，WPS会在适当的时机调用

/**
 * 加载项加载回调
 * 当WPS加载这个加载项时调用
 * 这是加载项的入口点，在这里进行初始化
 *
 * @returns {boolean} 是否加载成功
 */
function OnAddinLoad() {
    console.log('[Main] OnAddinLoad 被调用，开始初始化加载项...');

    try {
        var success = initAddon();

        if (success) {
            console.log('[Main] 加载项加载成功！');

            // 可选：自动启动服务器
            // 根据需求决定是否在加载时自动启动
            // startServer();
        } else {
            console.error('[Main] 加载项加载失败！');
        }

        return success;

    } catch (error) {
        console.error('[Main] OnAddinLoad异常: ' + error.message);
        return false;
    }
}

/**
 * 加载项通知回调
 * 当有消息需要通知加载项时调用
 *
 * @param {object} params - 通知参数
 */
function OnAddinNotify(params) {
    if (mainLogger) {
        mainLogger.debug('收到通知', params);
    } else {
        console.log('[Main] OnAddinNotify: ', params);
    }

    // 处理不同类型的通知
    if (params && params.eventType) {
        switch (params.eventType) {
            case 'documentOpened':
                // 文档打开事件
                if (mainLogger) {
                    mainLogger.info('文档已打开', { name: params.documentName });
                }
                break;

            case 'documentClosed':
                // 文档关闭事件
                if (mainLogger) {
                    mainLogger.info('文档已关闭', { name: params.documentName });
                }
                break;

            case 'selectionChanged':
                // 选区变化事件（可能会很频繁，只记录debug级别）
                if (mainLogger) {
                    mainLogger.debug('选区已变化');
                }
                break;

            default:
                if (mainLogger) {
                    mainLogger.debug('未处理的通知类型', { type: params.eventType });
                }
        }
    }
}

/**
 * 加载项启用回调
 * 当用户在WPS中启用这个加载项时调用
 */
function OnAddinEnable() {
    if (mainLogger) {
        mainLogger.info('加载项已启用');
    } else {
        console.log('[Main] OnAddinEnable: 加载项已启用');
    }

    // 如果之前已初始化，可以考虑自动启动服务
    if (addonStatus.initialized && !addonStatus.serverRunning) {
        // 用户启用加载项时自动启动服务
        startServer();
    }
}

/**
 * 加载项禁用回调
 * 当用户在WPS中禁用这个加载项时调用
 */
function OnAddinDisable() {
    if (mainLogger) {
        mainLogger.info('加载项已禁用');
    } else {
        console.log('[Main] OnAddinDisable: 加载项已禁用');
    }

    // 禁用时停止服务器
    if (addonStatus.serverRunning) {
        stopServer();
    }
}

/**
 * 加载项卸载回调（可选）
 * 当加载项被完全卸载时调用
 */
function OnAddinUnload() {
    if (mainLogger) {
        mainLogger.info('加载项正在卸载...');
    }

    // 清理资源
    if (httpServer && addonStatus.serverRunning) {
        stopServer();
    }

    mainLogger = null;
    httpServer = null;
    router = null;

    console.log('[Main] 加载项已卸载');
}

// ==================== 调试和测试辅助函数 ====================
// 这些函数用于开发调试，生产环境可以删掉

/**
 * 手动初始化（用于调试）
 * 如果OnAddinLoad没有被自动调用，可以手动调用这个
 */
function debugInit() {
    console.log('[Debug] 手动初始化加载项...');
    return initAddon();
}

/**
 * 手动启动服务器（用于调试）
 */
function debugStartServer() {
    console.log('[Debug] 手动启动服务器...');
    return startServer();
}

/**
 * 手动发送测试请求（用于调试）
 * 模拟一个来自MCP Server的请求
 *
 * @param {string} action - 操作名称
 * @param {object} params - 参数
 */
function debugTestAction(action, params) {
    console.log('[Debug] 测试action: ' + action);

    if (!router) {
        console.error('[Debug] 路由器未初始化');
        return null;
    }

    var result = router.dispatch(action, params || {}, 'debug-' + Date.now(), Date.now());
    console.log('[Debug] 结果: ', JSON.stringify(result, null, 2));
    return result;
}

/**
 * 打印当前状态（用于调试）
 */
function debugPrintStatus() {
    console.log('============ 调试信息 ============');
    console.log('addonStatus:', JSON.stringify(addonStatus, null, 2));
    console.log('serverConfig:', JSON.stringify(serverConfig, null, 2));

    if (httpServer) {
        console.log('httpServer.getStatus():', JSON.stringify(httpServer.getStatus(), null, 2));
    }

    if (router) {
        console.log('已注册的路由:', router.getActions().join(', '));
    }

    console.log('==================================');
}

// ==================== 模块导出（兼容性处理） ====================
// 虽然WPS加载项不用CommonJS，但保留导出以便单元测试

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        // 核心函数
        initAddon: initAddon,
        startServer: startServer,
        stopServer: stopServer,
        getServerStatus: getServerStatus,

        // Ribbon回调
        ribbonCheckStatus: ribbonCheckStatus,
        ribbonStartServer: ribbonStartServer,
        ribbonStopServer: ribbonStopServer,
        ribbonShowLog: ribbonShowLog,
        getStatusImage: getStatusImage,

        // 生命周期
        OnAddinLoad: OnAddinLoad,
        OnAddinNotify: OnAddinNotify,
        OnAddinEnable: OnAddinEnable,
        OnAddinDisable: OnAddinDisable,
        OnAddinUnload: OnAddinUnload,

        // 调试函数
        debugInit: debugInit,
        debugStartServer: debugStartServer,
        debugTestAction: debugTestAction,
        debugPrintStatus: debugPrintStatus
    };
}

// ==================== 立即执行代码 ====================
// 这段代码在脚本加载时立即执行

console.log('[Main] main.js 已加载');
console.log('[Main] 等待WPS调用 OnAddinLoad...');

// 如果是在测试环境，可以自动初始化
// 检测是否在WPS环境中
if (typeof Application === 'undefined') {
    console.log('[Main] 未检测到WPS环境，可能是测试模式');
    console.log('[Main] 使用 debugInit() 手动初始化');
}
