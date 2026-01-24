/**
 * 错误处理工具 - 老王的错误处理系统
 * 报错见了这个模块都得绕道走，艹
 */

import { log } from './logger';

/**
 * 错误码枚举 - 各种SB错误都有对应的码
 */
export enum ErrorCode {
  // 通用错误 1xxx
  UNKNOWN = 1000,
  INVALID_PARAMS = 1001,
  TIMEOUT = 1002,
  INTERNAL_ERROR = 1003,

  // WPS连接错误 2xxx
  WPS_CONNECTION_FAILED = 2001,
  WPS_NOT_RUNNING = 2002,
  WPS_API_ERROR = 2003,
  WPS_TIMEOUT = 2004,

  // Tool错误 3xxx
  TOOL_NOT_FOUND = 3001,
  TOOL_EXECUTION_FAILED = 3002,
  TOOL_INVALID_ARGS = 3003,
  TOOL_ALREADY_REGISTERED = 3004,

  // MCP协议错误 4xxx
  MCP_INVALID_REQUEST = 4001,
  MCP_METHOD_NOT_FOUND = 4002,
  MCP_PARSE_ERROR = 4003,
}

/**
 * 自定义错误基类 - 所有错误的祖宗
 */
export class McpError extends Error {
  public readonly code: ErrorCode;
  public readonly details?: Record<string, unknown>;
  public readonly timestamp: Date;

  constructor(
    message: string,
    code: ErrorCode = ErrorCode.UNKNOWN,
    details?: Record<string, unknown>
  ) {
    super(message);
    this.name = 'McpError';
    this.code = code;
    this.details = details;
    this.timestamp = new Date();

    // 保持正确的原型链
    Object.setPrototypeOf(this, McpError.prototype);
  }

  /**
   * 转换为JSON格式 - 方便传输
   */
  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      message: this.message,
      code: this.code,
      details: this.details,
      timestamp: this.timestamp.toISOString(),
      stack: this.stack,
    };
  }
}

/**
 * WPS连接错误 - 连不上WPS就用这个骂
 */
export class WpsConnectionError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.WPS_CONNECTION_FAILED, details);
    this.name = 'WpsConnectionError';
    Object.setPrototypeOf(this, WpsConnectionError.prototype);
  }
}

/**
 * WPS API错误 - WPS API返回的SB错误
 */
export class WpsApiError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.WPS_API_ERROR, details);
    this.name = 'WpsApiError';
    Object.setPrototypeOf(this, WpsApiError.prototype);
  }
}

/**
 * Tool不存在错误 - 找不到Tool就骂这个
 */
export class ToolNotFoundError extends McpError {
  constructor(toolName: string) {
    super(`Tool not found: ${toolName}`, ErrorCode.TOOL_NOT_FOUND, { toolName });
    this.name = 'ToolNotFoundError';
    Object.setPrototypeOf(this, ToolNotFoundError.prototype);
  }
}

/**
 * Tool执行错误 - Tool跑出问题了
 */
export class ToolExecutionError extends McpError {
  constructor(toolName: string, originalError: Error, details?: Record<string, unknown>) {
    super(
      `Tool execution failed: ${toolName} - ${originalError.message}`,
      ErrorCode.TOOL_EXECUTION_FAILED,
      { toolName, originalError: originalError.message, ...details }
    );
    this.name = 'ToolExecutionError';
    Object.setPrototypeOf(this, ToolExecutionError.prototype);
  }
}

/**
 * 参数验证错误 - 参数传错了就骂这个
 */
export class InvalidParamsError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.INVALID_PARAMS, details);
    this.name = 'InvalidParamsError';
    Object.setPrototypeOf(this, InvalidParamsError.prototype);
  }
}

/**
 * 超时错误 - 等太久了，不等了
 */
export class TimeoutError extends McpError {
  constructor(operation: string, timeoutMs: number) {
    super(
      `Operation timed out: ${operation} (${timeoutMs}ms)`,
      ErrorCode.TIMEOUT,
      { operation, timeoutMs }
    );
    this.name = 'TimeoutError';
    Object.setPrototypeOf(this, TimeoutError.prototype);
  }
}

/**
 * 错误处理工具函数
 */
export const errorUtils = {
  /**
   * 包装错误 - 把各种SB错误统一包装成McpError
   */
  wrap(error: unknown, defaultMessage = 'An error occurred'): McpError {
    if (error instanceof McpError) {
      return error;
    }

    if (error instanceof Error) {
      return new McpError(error.message, ErrorCode.INTERNAL_ERROR, {
        originalName: error.name,
        stack: error.stack,
      });
    }

    return new McpError(
      typeof error === 'string' ? error : defaultMessage,
      ErrorCode.UNKNOWN,
      { originalError: error }
    );
  },

  /**
   * 记录并重新抛出错误
   */
  logAndThrow(error: unknown, context?: string): never {
    const mcpError = errorUtils.wrap(error);
    log.error(context ? `${context}: ${mcpError.message}` : mcpError.message, mcpError);
    throw mcpError;
  },

  /**
   * 安全执行函数 - 出错了返回默认值，不会炸
   */
  async safeExecute<T>(
    fn: () => Promise<T>,
    defaultValue: T,
    context?: string
  ): Promise<T> {
    try {
      return await fn();
    } catch (error) {
      const mcpError = errorUtils.wrap(error);
      log.error(
        context ? `${context}: ${mcpError.message}` : mcpError.message,
        mcpError
      );
      return defaultValue;
    }
  },

  /**
   * 判断是否是特定类型的错误
   */
  isErrorCode(error: unknown, code: ErrorCode): boolean {
    return error instanceof McpError && error.code === code;
  },
};

/**
 * 格式化错误信息给用户看
 */
export const formatErrorForUser = (error: McpError): string => {
  switch (error.code) {
    case ErrorCode.WPS_CONNECTION_FAILED:
    case ErrorCode.WPS_NOT_RUNNING:
      return '无法连接到WPS Office，请确保WPS已启动并且加载项已安装';
    case ErrorCode.WPS_TIMEOUT:
      return 'WPS操作超时，请稍后重试';
    case ErrorCode.TOOL_NOT_FOUND:
      return `找不到指定的工具: ${error.details?.toolName}`;
    case ErrorCode.INVALID_PARAMS:
      return `参数错误: ${error.message}`;
    default:
      return `操作失败: ${error.message}`;
  }
};

export default McpError;
