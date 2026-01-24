/**
 * WPS Claude 智能助手 - 统一响应封装
 *
 * 这个SB模块负责统一所有HTTP响应的格式
 * 不管成功还是失败，都用同一个格式返回
 * 这样MCP Server那边处理起来才方便，不然你他妈的每个接口格式都不一样，谁受得了
 *
 * @author 老王
 * @date 2026-01-24
 */

/**
 * 错误码定义
 * 按照详细设计文档的规范来，别他妈自己瞎编
 */
var ErrorCodes = {
    // 连接错误 (1000-1099)
    CONNECTION_REFUSED: { code: '1001', message: 'WPS 未启动或加载项未安装' },
    CONNECTION_TIMEOUT: { code: '1002', message: 'WPS 响应超时' },
    CONNECTION_LOST: { code: '1003', message: 'WPS 连接断开' },

    // 参数错误 (2000-2099)
    PARAM_MISSING: { code: '2001', message: '缺少必要参数' },
    PARAM_INVALID: { code: '2002', message: '参数格式无效' },
    PARAM_RANGE_INVALID: { code: '2003', message: '单元格范围格式无效' },
    PARAM_FORMULA_INVALID: { code: '2004', message: '公式格式无效' },

    // 执行错误 (3000-3099)
    DOCUMENT_NOT_FOUND: { code: '3001', message: '未找到活动文档' },
    DOCUMENT_READONLY: { code: '3002', message: '文档为只读状态' },
    CELL_NOT_FOUND: { code: '3003', message: '未找到指定单元格' },
    FORMULA_ERROR: { code: '3004', message: '公式执行出错' },
    SHEET_NOT_FOUND: { code: '3005', message: '未找到指定工作表' },
    SHAPE_NOT_FOUND: { code: '3006', message: '未找到指定形状' },
    ACTION_NOT_FOUND: { code: '3007', message: '未知的操作类型' },

    // 权限错误 (4000-4099)
    PERMISSION_DENIED: { code: '4001', message: '没有操作权限' },

    // 系统错误 (5000-5099)
    UNKNOWN_ERROR: { code: '5001', message: '未知错误' },
    INTERNAL_ERROR: { code: '5002', message: '内部错误' },
    SERVER_NOT_RUNNING: { code: '5003', message: 'HTTP服务未运行' }
};

/**
 * Response 响应构建器
 * 用于构建统一格式的响应对象
 */
var Response = {

    /**
     * 构建成功响应
     *
     * @param {any} data - 返回的数据
     * @param {string} requestId - 请求ID
     * @param {number} startTime - 请求开始时间戳（用于计算耗时）
     * @returns {object} 标准响应对象
     */
    success: function(data, requestId, startTime) {
        var duration = startTime ? (Date.now() - startTime) : 0;

        return {
            success: true,
            data: data,
            error: null,
            requestId: requestId || null,
            duration: duration,
            timestamp: new Date().toISOString()
        };
    },

    /**
     * 构建失败响应
     *
     * @param {string} errorKey - 错误码键名（对应ErrorCodes中的key）
     * @param {string} requestId - 请求ID
     * @param {number} startTime - 请求开始时间戳
     * @param {object} details - 错误详情
     * @param {string} suggestion - 修复建议
     * @returns {object} 标准响应对象
     */
    error: function(errorKey, requestId, startTime, details, suggestion) {
        var duration = startTime ? (Date.now() - startTime) : 0;
        var errorDef = ErrorCodes[errorKey] || ErrorCodes.UNKNOWN_ERROR;

        return {
            success: false,
            data: null,
            error: {
                code: errorDef.code,
                type: errorKey,
                message: errorDef.message,
                details: details || null,
                suggestion: suggestion || null
            },
            requestId: requestId || null,
            duration: duration,
            timestamp: new Date().toISOString()
        };
    },

    /**
     * 构建自定义错误响应
     * 用于那些不在ErrorCodes里的憨批错误
     *
     * @param {string} code - 错误码
     * @param {string} message - 错误消息
     * @param {string} requestId - 请求ID
     * @param {number} startTime - 请求开始时间戳
     * @param {object} details - 错误详情
     * @returns {object} 标准响应对象
     */
    customError: function(code, message, requestId, startTime, details) {
        var duration = startTime ? (Date.now() - startTime) : 0;

        return {
            success: false,
            data: null,
            error: {
                code: code,
                type: 'CUSTOM_ERROR',
                message: message,
                details: details || null,
                suggestion: null
            },
            requestId: requestId || null,
            duration: duration,
            timestamp: new Date().toISOString()
        };
    },

    /**
     * 从异常对象构建错误响应
     * 捕获异常的时候用这个，省得每次都手动构建
     *
     * @param {Error} err - 异常对象
     * @param {string} requestId - 请求ID
     * @param {number} startTime - 请求开始时间戳
     * @returns {object} 标准响应对象
     */
    fromException: function(err, requestId, startTime) {
        var duration = startTime ? (Date.now() - startTime) : 0;

        return {
            success: false,
            data: null,
            error: {
                code: '5001',
                type: 'EXCEPTION',
                message: err.message || '未知异常',
                details: {
                    name: err.name,
                    stack: err.stack
                },
                suggestion: '请检查操作参数是否正确，或联系技术支持'
            },
            requestId: requestId || null,
            duration: duration,
            timestamp: new Date().toISOString()
        };
    },

    /**
     * 将响应对象转换为JSON字符串
     *
     * @param {object} responseObj - 响应对象
     * @returns {string} JSON字符串
     */
    toJSON: function(responseObj) {
        try {
            return JSON.stringify(responseObj);
        } catch (e) {
            // 如果序列化失败，返回一个简单的错误响应
            return JSON.stringify({
                success: false,
                error: {
                    code: '5002',
                    message: '响应序列化失败'
                }
            });
        }
    }
};

/**
 * 参数校验工具
 * 检查必填参数是否存在
 */
var Validator = {

    /**
     * 检查必填参数
     *
     * @param {object} params - 参数对象
     * @param {Array} required - 必填参数列表
     * @returns {object} { valid: boolean, missing: string[] }
     */
    checkRequired: function(params, required) {
        var missing = [];

        if (!params) {
            return { valid: false, missing: required };
        }

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

    /**
     * 校验单元格范围格式
     * 合法格式：A1, A1:B10, Sheet1!A1:B10
     *
     * @param {string} range - 范围字符串
     * @returns {boolean} 是否合法
     */
    isValidRange: function(range) {
        if (!range || typeof range !== 'string') {
            return false;
        }

        // 简单的正则校验，支持 A1, A1:B10, Sheet1!A1:B10 格式
        var rangePattern = /^([^!]+!)?[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i;
        return rangePattern.test(range.trim());
    },

    /**
     * 校验公式格式
     * 公式必须以=开头
     *
     * @param {string} formula - 公式字符串
     * @returns {boolean} 是否合法
     */
    isValidFormula: function(formula) {
        if (!formula || typeof formula !== 'string') {
            return false;
        }

        return formula.trim().startsWith('=');
    }
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        Response: Response,
        ErrorCodes: ErrorCodes,
        Validator: Validator
    };
}
