/**
 * WPS Claude 智能助手 - 请求路由分发器
 *
 * 这个SB模块负责把请求分发到对应的处理器
 * 没有路由，你的请求就是无头苍蝇，不知道往哪飞
 *
 * 路由规则：
 * - action 格式: "模块.方法"，如 "excel.getContext", "word.insertText"
 * - 系统级 action: "system.ping", "system.status"
 *
 * @author 老王
 * @date 2026-01-24
 */

/**
 * Router 路由器类
 *
 * @param {Logger} logger - 日志记录器实例
 */
function Router(logger) {
    this.logger = logger || new Logger('Router');
    this.handlers = {}; // 处理器注册表
    this.middleware = []; // 中间件列表

    // 注册系统级路由
    this._registerSystemRoutes();
}

/**
 * 注册系统级路由
 * 这些是加载项自身的管理接口，别他妈删了
 */
Router.prototype._registerSystemRoutes = function() {
    var self = this;

    // ping - 心跳检测
    this.register('system.ping', function(params, requestId, startTime) {
        return Response.success({
            message: 'pong',
            timestamp: new Date().toISOString(),
            version: '1.0.0'
        }, requestId, startTime);
    });

    // status - 获取加载项状态
    this.register('system.status', function(params, requestId, startTime) {
        // 这里后续会从HttpServer获取状态
        return Response.success({
            addon: {
                name: 'WPS Claude 智能助手',
                version: '1.0.0',
                status: 'running'
            },
            wps: self._getWpsInfo(),
            registeredActions: Object.keys(self.handlers)
        }, requestId, startTime);
    });

    // listActions - 列出所有可用的action
    this.register('system.listActions', function(params, requestId, startTime) {
        var actions = [];
        for (var action in self.handlers) {
            if (self.handlers.hasOwnProperty(action)) {
                actions.push(action);
            }
        }
        return Response.success({
            count: actions.length,
            actions: actions.sort()
        }, requestId, startTime);
    });

    this.logger.info('系统路由注册完成');
};

/**
 * 获取WPS信息
 *
 * @returns {object} WPS信息对象
 */
Router.prototype._getWpsInfo = function() {
    try {
        if (typeof Application !== 'undefined') {
            return {
                name: Application.Name || 'WPS Office',
                version: Application.Version || 'Unknown',
                build: Application.Build || 'Unknown',
                activeDocument: this._getActiveDocumentInfo()
            };
        }
    } catch (e) {
        // 忽略错误
    }

    return {
        name: 'WPS Office',
        version: 'Unknown',
        activeDocument: null
    };
};

/**
 * 获取当前活动文档信息
 *
 * @returns {object|null} 文档信息
 */
Router.prototype._getActiveDocumentInfo = function() {
    try {
        // 尝试获取不同类型的活动文档
        if (typeof Application !== 'undefined') {
            // Excel
            if (Application.ActiveWorkbook) {
                return {
                    type: 'excel',
                    name: Application.ActiveWorkbook.Name
                };
            }
            // Word
            if (Application.ActiveDocument) {
                return {
                    type: 'word',
                    name: Application.ActiveDocument.Name
                };
            }
            // PPT
            if (Application.ActivePresentation) {
                return {
                    type: 'ppt',
                    name: Application.ActivePresentation.Name
                };
            }
        }
    } catch (e) {
        // 忽略错误
    }

    return null;
};

/**
 * 注册路由处理器
 *
 * @param {string} action - 操作名称，如 "excel.getContext"
 * @param {Function} handler - 处理函数，签名: (params, requestId, startTime) => Response
 */
Router.prototype.register = function(action, handler) {
    if (typeof handler !== 'function') {
        this.logger.error('注册失败：handler必须是函数', { action: action });
        return;
    }

    this.handlers[action] = handler;
    this.logger.debug('路由注册成功', { action: action });
};

/**
 * 批量注册路由
 *
 * @param {object} routes - 路由映射对象 { action: handler }
 */
Router.prototype.registerAll = function(routes) {
    for (var action in routes) {
        if (routes.hasOwnProperty(action)) {
            this.register(action, routes[action]);
        }
    }
};

/**
 * 注销路由
 *
 * @param {string} action - 操作名称
 */
Router.prototype.unregister = function(action) {
    if (this.handlers[action]) {
        delete this.handlers[action];
        this.logger.debug('路由注销成功', { action: action });
    }
};

/**
 * 添加中间件
 * 中间件会在处理器执行前依次执行
 *
 * @param {Function} middleware - 中间件函数，签名: (action, params, next) => void
 */
Router.prototype.use = function(middleware) {
    if (typeof middleware === 'function') {
        this.middleware.push(middleware);
        this.logger.debug('中间件添加成功');
    }
};

/**
 * 分发请求到对应的处理器
 * 这是路由器的核心方法
 *
 * @param {string} action - 操作名称
 * @param {object} params - 请求参数
 * @param {string} requestId - 请求ID
 * @param {number} startTime - 请求开始时间
 * @returns {object} 响应对象
 */
Router.prototype.dispatch = function(action, params, requestId, startTime) {
    this.logger.debug('分发请求', { action: action, requestId: requestId });

    // 检查action是否存在
    if (!action) {
        return Response.error('PARAM_MISSING', requestId, startTime, {
            missing: ['action']
        }, '请在请求体中指定 action 参数');
    }

    // 查找处理器
    var handler = this.handlers[action];

    if (!handler) {
        this.logger.warn('未找到处理器', { action: action });

        // 尝试给出建议
        var suggestion = this._findSimilarAction(action);

        return Response.error('ACTION_NOT_FOUND', requestId, startTime, {
            action: action,
            availableActions: Object.keys(this.handlers).slice(0, 10)
        }, suggestion);
    }

    try {
        // 执行处理器
        var result = handler(params, requestId, startTime);

        this.logger.debug('请求处理完成', {
            action: action,
            requestId: requestId,
            success: result.success
        });

        return result;

    } catch (error) {
        this.logger.error('处理器执行异常', {
            action: action,
            requestId: requestId,
            error: error.message,
            stack: error.stack
        });

        return Response.fromException(error, requestId, startTime);
    }
};

/**
 * 查找相似的action名称（用于错误提示）
 *
 * @param {string} action - 用户输入的action
 * @returns {string} 建议信息
 */
Router.prototype._findSimilarAction = function(action) {
    var parts = action.split('.');
    var module = parts[0];
    var method = parts[1];

    // 查找同模块的其他方法
    var sameModuleActions = [];
    for (var registeredAction in this.handlers) {
        if (registeredAction.startsWith(module + '.')) {
            sameModuleActions.push(registeredAction);
        }
    }

    if (sameModuleActions.length > 0) {
        return '模块 "' + module + '" 下可用的操作: ' + sameModuleActions.join(', ');
    }

    // 列出所有可用模块
    var modules = {};
    for (var act in this.handlers) {
        var mod = act.split('.')[0];
        modules[mod] = true;
    }

    return '可用的模块: ' + Object.keys(modules).join(', ');
};

/**
 * 检查action是否已注册
 *
 * @param {string} action - 操作名称
 * @returns {boolean} 是否已注册
 */
Router.prototype.has = function(action) {
    return this.handlers.hasOwnProperty(action);
};

/**
 * 获取所有已注册的action列表
 *
 * @returns {Array} action列表
 */
Router.prototype.getActions = function() {
    return Object.keys(this.handlers).sort();
};

/**
 * 获取指定模块的所有action
 *
 * @param {string} module - 模块名称
 * @returns {Array} action列表
 */
Router.prototype.getModuleActions = function(module) {
    var actions = [];
    var prefix = module + '.';

    for (var action in this.handlers) {
        if (action.startsWith(prefix)) {
            actions.push(action);
        }
    }

    return actions.sort();
};

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        Router: Router
    };
}
