/**
 * Tool注册管理 - 老王的Tool管理系统
 * 所有Tool都得在这儿注册，不注册的Tool就是野鸡Tool
 * 这个设计遵循OCP原则：扩展开放，修改关闭
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  RegisteredTool,
  ToolCallRequest,
  ToolCallResult,
  ToolCategory,
  ListToolsResponse,
} from '../types/tools';
import { createChildLogger } from '../utils/logger';
import {
  ToolNotFoundError,
  ToolExecutionError,
  InvalidParamsError,
  McpError,
  ErrorCode,
} from '../utils/error';

const logger = createChildLogger('ToolRegistry');

/**
 * Tool注册表 - 单例模式，全局唯一
 */
export class ToolRegistry {
  private static instance: ToolRegistry;
  private readonly tools: Map<string, RegisteredTool>;
  private readonly categories: Map<ToolCategory, Set<string>>;

  private constructor() {
    this.tools = new Map();
    this.categories = new Map();

    // 初始化分类
    Object.values(ToolCategory).forEach((category) => {
      this.categories.set(category, new Set());
    });

    logger.info('ToolRegistry initialized');
  }

  /**
   * 获取单例实例
   */
  static getInstance(): ToolRegistry {
    if (!ToolRegistry.instance) {
      ToolRegistry.instance = new ToolRegistry();
    }
    return ToolRegistry.instance;
  }

  /**
   * 注册Tool - 把Tool加到注册表里
   * 遵循SRP：只负责注册，不负责其他
   */
  register(definition: ToolDefinition, handler: ToolHandler): void {
    const { name, category } = definition;

    // 检查是否已注册
    if (this.tools.has(name)) {
      throw new McpError(
        `Tool already registered: ${name}`,
        ErrorCode.TOOL_ALREADY_REGISTERED,
        { toolName: name }
      );
    }

    // 注册Tool
    this.tools.set(name, { definition, handler });

    // 添加到分类
    const toolCategory = category || ToolCategory.COMMON;
    this.categories.get(toolCategory)?.add(name);

    logger.info(`Tool registered: ${name}`, { category: toolCategory });
  }

  /**
   * 批量注册Tools - 一次注册一堆
   */
  registerAll(tools: Array<{ definition: ToolDefinition; handler: ToolHandler }>): void {
    tools.forEach(({ definition, handler }) => {
      this.register(definition, handler);
    });
    logger.info(`Batch registered ${tools.length} tools`);
  }

  /**
   * 注销Tool - 把Tool从注册表里删掉
   */
  unregister(name: string): boolean {
    const tool = this.tools.get(name);
    if (!tool) {
      return false;
    }

    this.tools.delete(name);

    // 从分类中移除
    const category = tool.definition.category || ToolCategory.COMMON;
    this.categories.get(category)?.delete(name);

    logger.info(`Tool unregistered: ${name}`);
    return true;
  }

  /**
   * 获取Tool定义
   */
  getTool(name: string): RegisteredTool | undefined {
    return this.tools.get(name);
  }

  /**
   * 检查Tool是否存在
   */
  hasTool(name: string): boolean {
    return this.tools.has(name);
  }

  /**
   * 获取所有Tool定义 - MCP的tools/list用的
   */
  listTools(): ListToolsResponse {
    const tools = Array.from(this.tools.values()).map((t) => t.definition);
    return { tools };
  }

  /**
   * 按分类获取Tools
   */
  getToolsByCategory(category: ToolCategory): ToolDefinition[] {
    const toolNames = this.categories.get(category);
    if (!toolNames) {
      return [];
    }

    return Array.from(toolNames)
      .map((name) => this.tools.get(name)?.definition)
      .filter((t): t is ToolDefinition => t !== undefined);
  }

  /**
   * 调用Tool - 执行Tool的handler
   * 这是核心方法，处理Tool调用请求
   */
  async callTool(request: ToolCallRequest): Promise<ToolCallResult> {
    const { id, name, arguments: args } = request;
    const startTime = Date.now();

    logger.debug(`Calling tool: ${name}`, { id, args });

    // 检查Tool是否存在
    const tool = this.tools.get(name);
    if (!tool) {
      throw new ToolNotFoundError(name);
    }

    try {
      // 验证参数
      this.validateArguments(tool.definition, args);

      // 执行handler
      const result = await tool.handler(args);

      const duration = Date.now() - startTime;
      logger.info(`Tool executed: ${name}`, { id, duration, success: result.success });

      return {
        ...result,
        id,
      };
    } catch (error) {
      const duration = Date.now() - startTime;

      if (error instanceof McpError) {
        logger.error(`Tool execution failed: ${name}`, error, { id, duration });
        return {
          id,
          success: false,
          content: [{ type: 'text', text: error.message }],
          error: error.message,
        };
      }

      const execError = new ToolExecutionError(
        name,
        error instanceof Error ? error : new Error(String(error))
      );

      logger.error(`Tool execution failed: ${name}`, execError, { id, duration });

      return {
        id,
        success: false,
        content: [{ type: 'text', text: execError.message }],
        error: execError.message,
      };
    }
  }

  /**
   * 验证参数 - 检查必填参数是否都有
   */
  private validateArguments(
    definition: ToolDefinition,
    args: Record<string, unknown>
  ): void {
    const { inputSchema } = definition;
    const required = inputSchema.required || [];

    for (const param of required) {
      if (!(param in args) || args[param] === undefined || args[param] === null) {
        throw new InvalidParamsError(`Missing required parameter: ${param}`, {
          toolName: definition.name,
          missingParam: param,
        });
      }
    }

    // TODO: 可以加更详细的类型验证，但现在先这样，够用
  }

  /**
   * 获取Tool数量
   */
  get size(): number {
    return this.tools.size;
  }

  /**
   * 清空所有Tools - 测试用，生产环境别乱用
   */
  clear(): void {
    this.tools.clear();
    this.categories.forEach((set) => set.clear());
    logger.warn('All tools cleared');
  }

  /**
   * 创建Tool调用请求 - 辅助方法
   */
  static createRequest(
    name: string,
    args: Record<string, unknown>
  ): ToolCallRequest {
    return {
      id: uuidv4(),
      name,
      arguments: args,
    };
  }
}

// 导出单例
export const toolRegistry = ToolRegistry.getInstance();

/**
 * 装饰器：注册Tool
 * 用法：@RegisterTool(definition)
 */
export function RegisterTool(definition: ToolDefinition) {
  return function (
    _target: unknown,
    _propertyKey: string,
    descriptor: PropertyDescriptor
  ): PropertyDescriptor {
    const originalMethod = descriptor.value as ToolHandler;
    toolRegistry.register(definition, originalMethod);
    return descriptor;
  };
}

/**
 * 快捷注册函数
 */
export const registerTool = (
  definition: ToolDefinition,
  handler: ToolHandler
): void => {
  toolRegistry.register(definition, handler);
};

export default ToolRegistry;
