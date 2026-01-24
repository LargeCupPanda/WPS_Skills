/**
 * WPS Claude 智能助手 - 日志工具
 *
 * 艹，没有日志记录，出了问题你就抓瞎了
 * 这个SB模块提供统一的日志记录功能
 *
 * @author 老王
 * @date 2026-01-24
 */

// 日志级别定义，别TM乱改顺序
var LogLevel = {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3
};

// 日志级别名称映射
var LogLevelName = {
    0: 'DEBUG',
    1: 'INFO',
    2: 'WARN',
    3: 'ERROR'
};

/**
 * Logger 日志记录器
 * 这个憨批类负责所有日志的格式化和输出
 *
 * @param {string} moduleName - 模块名称，用于区分日志来源
 */
function Logger(moduleName) {
    this.moduleName = moduleName || 'Unknown';
    this.level = LogLevel.INFO; // 默认日志级别
    this.logHistory = []; // 日志历史记录
    this.maxHistorySize = 500; // 最多保留500条日志
}

/**
 * 设置日志级别
 * 低于这个级别的日志不会输出，省点性能
 *
 * @param {number} level - 日志级别
 */
Logger.prototype.setLevel = function(level) {
    if (level >= LogLevel.DEBUG && level <= LogLevel.ERROR) {
        this.level = level;
    }
};

/**
 * 格式化日志消息
 * 把日志格式化成人能看懂的样子
 *
 * @param {string} level - 日志级别
 * @param {string} message - 日志消息
 * @param {object} data - 附加数据
 * @returns {string} 格式化后的日志字符串
 */
Logger.prototype.formatMessage = function(level, message, data) {
    var timestamp = new Date().toISOString();
    var formattedMsg = '[' + timestamp + '] [' + level + '] [' + this.moduleName + '] ' + message;

    if (data !== undefined && data !== null) {
        try {
            formattedMsg += ' | Data: ' + JSON.stringify(data);
        } catch (e) {
            formattedMsg += ' | Data: [无法序列化]';
        }
    }

    return formattedMsg;
};

/**
 * 保存日志到历史记录
 *
 * @param {object} logEntry - 日志条目
 */
Logger.prototype.saveToHistory = function(logEntry) {
    this.logHistory.push(logEntry);

    // 超过最大数量就删掉最早的，别让内存爆了
    if (this.logHistory.length > this.maxHistorySize) {
        this.logHistory.shift();
    }
};

/**
 * 输出日志
 *
 * @param {number} level - 日志级别
 * @param {string} message - 日志消息
 * @param {object} data - 附加数据
 */
Logger.prototype.log = function(level, message, data) {
    // 低于当前级别的日志不输出
    if (level < this.level) {
        return;
    }

    var levelName = LogLevelName[level] || 'UNKNOWN';
    var formattedMsg = this.formatMessage(levelName, message, data);

    // 创建日志条目
    var logEntry = {
        timestamp: new Date().toISOString(),
        level: levelName,
        module: this.moduleName,
        message: message,
        data: data
    };

    // 保存到历史记录
    this.saveToHistory(logEntry);

    // 根据级别选择输出方式
    switch (level) {
        case LogLevel.DEBUG:
            console.log(formattedMsg);
            break;
        case LogLevel.INFO:
            console.info(formattedMsg);
            break;
        case LogLevel.WARN:
            console.warn(formattedMsg);
            break;
        case LogLevel.ERROR:
            console.error(formattedMsg);
            break;
        default:
            console.log(formattedMsg);
    }
};

/**
 * DEBUG 级别日志
 * 调试信息，生产环境可以关掉
 */
Logger.prototype.debug = function(message, data) {
    this.log(LogLevel.DEBUG, message, data);
};

/**
 * INFO 级别日志
 * 常规信息，让人知道程序在干嘛
 */
Logger.prototype.info = function(message, data) {
    this.log(LogLevel.INFO, message, data);
};

/**
 * WARN 级别日志
 * 警告信息，出了点问题但还能跑
 */
Logger.prototype.warn = function(message, data) {
    this.log(LogLevel.WARN, message, data);
};

/**
 * ERROR 级别日志
 * 错误信息，他妈的出大问题了
 */
Logger.prototype.error = function(message, data) {
    this.log(LogLevel.ERROR, message, data);
};

/**
 * 获取日志历史记录
 *
 * @param {number} count - 获取的数量，不传则返回全部
 * @returns {Array} 日志历史数组
 */
Logger.prototype.getHistory = function(count) {
    if (count && count > 0) {
        return this.logHistory.slice(-count);
    }
    return this.logHistory.slice();
};

/**
 * 清空日志历史
 */
Logger.prototype.clearHistory = function() {
    this.logHistory = [];
};

/**
 * 创建子日志记录器
 * 继承父级设置，但使用不同的模块名
 *
 * @param {string} subModuleName - 子模块名称
 * @returns {Logger} 新的日志记录器
 */
Logger.prototype.createChild = function(subModuleName) {
    var childLogger = new Logger(this.moduleName + '.' + subModuleName);
    childLogger.level = this.level;
    return childLogger;
};

// 导出模块（兼容WPS加载项环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        Logger: Logger,
        LogLevel: LogLevel
    };
}
