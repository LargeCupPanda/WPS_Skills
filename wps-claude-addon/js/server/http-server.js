/**
 * WPS Claude 智能助手 - HTTP服务器实现
 *
 * 这个SB模块是加载项的核心，负责监听HTTP请求
 * 没有它，MCP Server就没法和WPS通信
 *
 * 重要说明：
 * WPS加载项使用内置的HTTP服务能力 (Application.CreateHttpServer)
 * 这个API是WPS提供的，别他妈去找Node.js的http模块
 *
 * @author 老王
 * @date 2026-01-24
 */

// 引入依赖模块（WPS加载项环境下可能需要调整路径）
// 注意：WPS加载项不支持ES6的import，得用传统方式

/**
 * HttpServer HTTP服务器类
 *
 * @param {number} port - 监听端口，默认23333
 * @param {Logger} logger - 日志记录器实例
 */
function HttpServer(port, logger) {
    this.port = port || 23333;
    this.logger = logger || new Logger('HttpServer');
    this.server = null;
    this.isRunning = false;
    this.router = null; // 路由器引用
    this.startTime = null;
    this.requestCount = 0;
}

/**
 * 设置路由器
 *
 * @param {Router} router - 路由器实例
 */
HttpServer.prototype.setRouter = function(router) {
    this.router = router;
    this.logger.info('路由器已设置');
};

/**
 * 启动HTTP服务器
 * 这是核心方法，调用WPS的CreateHttpServer API
 *
 * @returns {boolean} 是否启动成功
 */
HttpServer.prototype.start = function() {
    var self = this;

    if (this.isRunning) {
        this.logger.warn('服务器已经在运行了，你他妈是不是眼瞎');
        return true;
    }

    try {
        this.logger.info('正在启动HTTP服务器...', { port: this.port });

        // 使用WPS提供的API创建HTTP服务器
        // 注意：这个API在不同版本的WPS可能略有不同
        if (typeof Application !== 'undefined' && Application.CreateHttpServer) {
            this.server = Application.CreateHttpServer(this.port);
        } else {
            // 如果在非WPS环境下（比如测试），使用模拟对象
            this.logger.warn('未检测到WPS环境，使用模拟HTTP服务器');
            this.server = this._createMockServer();
        }

        // 设置请求处理回调
        var server = this.server;
        server.OnRequest = function(request, response) {
            self._handleRequest(request, response);
        };

        // 标记运行状态
        this.isRunning = true;
        this.startTime = new Date();

        this.logger.info('HTTP服务器启动成功', {
            port: this.port,
            startTime: this.startTime.toISOString()
        });

        return true;

    } catch (error) {
        this.logger.error('HTTP服务器启动失败，艹！', {
            error: error.message,
            stack: error.stack
        });
        return false;
    }
};

/**
 * 停止HTTP服务器
 *
 * @returns {boolean} 是否停止成功
 */
HttpServer.prototype.stop = function() {
    if (!this.isRunning) {
        this.logger.warn('服务器本来就没运行，停个屁');
        return true;
    }

    try {
        if (this.server && this.server.Close) {
            this.server.Close();
        }

        this.isRunning = false;
        this.logger.info('HTTP服务器已停止', {
            totalRequests: this.requestCount,
            runTime: this._getRunTime()
        });

        return true;

    } catch (error) {
        this.logger.error('停止服务器失败', { error: error.message });
        return false;
    }
};

/**
 * 处理HTTP请求
 * 这是请求的入口，负责解析请求并分发到路由器
 *
 * @param {object} request - HTTP请求对象
 * @param {object} response - HTTP响应对象
 */
HttpServer.prototype._handleRequest = function(request, response) {
    var startTime = Date.now();
    var requestId = null;
    var body = null;

    this.requestCount++;

    try {
        // 设置CORS头，允许跨域访问
        response.SetHeader('Access-Control-Allow-Origin', '*');
        response.SetHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        response.SetHeader('Access-Control-Allow-Headers', 'Content-Type');
        response.SetHeader('Content-Type', 'application/json; charset=utf-8');

        // 处理OPTIONS预检请求
        if (request.Method === 'OPTIONS') {
            response.SetStatus(204);
            response.End();
            return;
        }

        // 解析请求体
        try {
            body = JSON.parse(request.Body || '{}');
        } catch (e) {
            this.logger.warn('请求体解析失败', { body: request.Body });
            body = {};
        }

        requestId = body.requestId || this._generateRequestId();
        var action = body.action;
        var params = body.params || {};

        this.logger.info('收到请求', {
            requestId: requestId,
            action: action,
            method: request.Method
        });

        // 检查是否有路由器
        if (!this.router) {
            throw new Error('路由器未设置');
        }

        // 分发请求到路由器
        var result = this.router.dispatch(action, params, requestId, startTime);

        // 发送响应
        this._sendResponse(response, result);

    } catch (error) {
        this.logger.error('请求处理异常', {
            requestId: requestId,
            error: error.message,
            stack: error.stack
        });

        // 发送错误响应
        var errorResponse = Response.fromException(error, requestId, startTime);
        this._sendResponse(response, errorResponse, 500);
    }
};

/**
 * 发送HTTP响应
 *
 * @param {object} response - HTTP响应对象
 * @param {object} data - 响应数据
 * @param {number} statusCode - HTTP状态码，默认200
 */
HttpServer.prototype._sendResponse = function(response, data, statusCode) {
    try {
        var status = statusCode || (data.success ? 200 : 500);
        response.SetStatus(status);

        var jsonStr = JSON.stringify(data);
        response.Write(jsonStr);
        response.End();

        this.logger.debug('响应已发送', {
            requestId: data.requestId,
            success: data.success,
            duration: data.duration
        });

    } catch (error) {
        this.logger.error('发送响应失败', { error: error.message });
    }
};

/**
 * 生成请求ID
 *
 * @returns {string} 请求ID
 */
HttpServer.prototype._generateRequestId = function() {
    return 'req-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
};

/**
 * 获取运行时长
 *
 * @returns {string} 运行时长字符串
 */
HttpServer.prototype._getRunTime = function() {
    if (!this.startTime) {
        return '0s';
    }

    var diff = Date.now() - this.startTime.getTime();
    var seconds = Math.floor(diff / 1000);
    var minutes = Math.floor(seconds / 60);
    var hours = Math.floor(minutes / 60);

    if (hours > 0) {
        return hours + 'h ' + (minutes % 60) + 'm';
    } else if (minutes > 0) {
        return minutes + 'm ' + (seconds % 60) + 's';
    } else {
        return seconds + 's';
    }
};

/**
 * 获取服务器状态
 *
 * @returns {object} 状态信息
 */
HttpServer.prototype.getStatus = function() {
    return {
        isRunning: this.isRunning,
        port: this.port,
        startTime: this.startTime ? this.startTime.toISOString() : null,
        runTime: this._getRunTime(),
        requestCount: this.requestCount
    };
};

/**
 * 创建模拟服务器（用于非WPS环境测试）
 *
 * @returns {object} 模拟服务器对象
 */
HttpServer.prototype._createMockServer = function() {
    return {
        OnRequest: null,
        Close: function() {
            console.log('[MockServer] 模拟服务器已关闭');
        }
    };
};

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        HttpServer: HttpServer
    };
}
