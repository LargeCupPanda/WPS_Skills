/**
 * 响应封装工具
 * 统一的响应格式，别TM乱改格式
 */

function success(data) {
    return {
        success: true,
        data: data || {},
        error: null
    };
}

function error(message) {
    return {
        success: false,
        data: null,
        error: message || '未知错误'
    };
}

// 导出
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { success, error };
}
