/**
 * Input: MCP 协议请求
 * Output: Tool 调用响应
 * Pos: MCP Server 核心实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * MCP Server实现 - 老王的MCP协议实现
 * 这是整个系统的核心，处理MCP协议通信
 * 基于stdio传输，Claude Desktop就是这么连的
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode as McpErrorCode,
  McpError as SdkMcpError,
} from '@modelcontextprotocol/sdk/types.js';

import { toolRegistry, ToolRegistry } from './tool-registry';
import { wpsClient } from '../client/wps-client';
import { ToolCallResult, ToolCategory } from '../types/tools';
import { allTools } from '../tools';
import { createChildLogger } from '../utils/logger';
import { McpError } from '../utils/error';

const logger = createChildLogger('McpServer');

/**
 * MCP Server配置
 */
export interface McpServerConfig {
  /** 服务器名称 */
  name: string;
  /** 服务器版本 */
  version: string;
  /** 是否启用调试模式 */
  debug?: boolean;
}

/**
 * 默认配置
 */
const DEFAULT_CONFIG: McpServerConfig = {
  name: 'wps-office-mcp',
  version: '1.0.0',
  debug: false,
};

/**
 * WPS MCP Server - 核心服务器类
 */
export class WpsMcpServer {
  private readonly config: McpServerConfig;
  private readonly server: Server;
  private readonly registry: ToolRegistry;
  private isRunning: boolean = false;

  // 跨应用数据缓存 - 解决macOS WPS无法跨应用操作的P0问题
  private static dataCache: Map<string, { data: unknown; timestamp: number; appType: string }> = new Map();

  constructor(config?: Partial<McpServerConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.registry = toolRegistry;

    // 创建MCP Server实例
    this.server = new Server(
      {
        name: this.config.name,
        version: this.config.version,
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    // 注册请求处理器
    this.setupRequestHandlers();

    // 错误处理
    this.server.onerror = (error) => {
      logger.error('MCP Server error', error);
    };

    logger.info('MCP Server created', {
      name: this.config.name,
      version: this.config.version,
    });
  }

  /**
   * 设置请求处理器
   */
  private setupRequestHandlers(): void {
    // 处理 tools/list 请求
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      logger.debug('Handling tools/list request');

      const { tools } = this.registry.listTools();

      logger.info(`Returning ${tools.length} tools`);

      return {
        tools: tools.map((tool) => ({
          name: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema,
        })),
      };
    });

    // 处理 tools/call 请求
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      logger.debug('Handling tools/call request', { name, args });

      // 检查Tool是否存在
      if (!this.registry.hasTool(name)) {
        throw new SdkMcpError(
          McpErrorCode.MethodNotFound,
          `Unknown tool: ${name}`
        );
      }

      try {
        // 创建调用请求并执行
        const callRequest = ToolRegistry.createRequest(name, args || {});
        const result: ToolCallResult = await this.registry.callTool(callRequest);

        return {
          content: result.content,
          isError: !result.success,
        };
      } catch (error) {
        logger.error('Tool call failed', error);

        if (error instanceof McpError) {
          return {
            content: [{ type: 'text', text: error.message }],
            isError: true,
          };
        }

        throw new SdkMcpError(
          McpErrorCode.InternalError,
          error instanceof Error ? error.message : 'Unknown error'
        );
      }
    });
  }

  /**
   * 注册内置Tools - 一些基础的WPS操作Tool
   */
  registerBuiltinTools(): void {
    logger.info('Registering built-in tools');

    // 连接状态检查Tool
    this.registry.register(
      {
        name: 'wps_check_connection',
        description: '检查WPS Office连接状态',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const connected = await wpsClient.checkConnection();
        const status = wpsClient.getStatus();

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                connected,
                status,
              }),
            },
          ],
        };
      }
    );

    // 获取当前文档信息
    this.registry.register(
      {
        name: 'wps_get_active_document',
        description: '获取当前打开的WPS文字文档信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.DOCUMENT,
      },
      async () => {
        const doc = await wpsClient.getActiveDocument();

        return {
          id: '',
          success: doc !== null,
          content: [
            {
              type: 'text',
              text: doc
                ? JSON.stringify(doc)
                : '没有打开的文档',
            },
          ],
        };
      }
    );

    // 在文档中插入文本
    this.registry.register(
      {
        name: 'wps_insert_text',
        description: '在当前文档中插入文本',
        inputSchema: {
          type: 'object',
          properties: {
            text: {
              type: 'string',
              description: '要插入的文本内容',
            },
            position: {
              type: 'number',
              description: '插入位置（可选，不指定则在光标处插入）',
            },
          },
          required: ['text'],
        },
        category: ToolCategory.DOCUMENT,
      },
      async (args) => {
        const text = args.text as string;
        const position = args.position as number | undefined;

        const success = await wpsClient.insertText(text, position);

        return {
          id: '',
          success,
          content: [
            {
              type: 'text',
              text: success ? '文本插入成功' : '文本插入失败',
            },
          ],
        };
      }
    );

    // 获取当前工作簿信息
    this.registry.register(
      {
        name: 'wps_get_active_workbook',
        description: '获取当前打开的WPS表格工作簿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.SPREADSHEET,
      },
      async () => {
        const workbook = await wpsClient.getActiveWorkbook();

        return {
          id: '',
          success: workbook !== null,
          content: [
            {
              type: 'text',
              text: workbook
                ? JSON.stringify(workbook)
                : '没有打开的工作簿',
            },
          ],
        };
      }
    );

    // 读取单元格值
    this.registry.register(
      {
        name: 'wps_get_cell_value',
        description: '读取指定单元格的值',
        inputSchema: {
          type: 'object',
          properties: {
            sheet: {
              type: 'string',
              description: '工作表名称或索引',
            },
            row: {
              type: 'number',
              description: '行号（从1开始）',
            },
            col: {
              type: 'number',
              description: '列号（从1开始）',
            },
          },
          required: ['sheet', 'row', 'col'],
        },
        category: ToolCategory.SPREADSHEET,
      },
      async (args) => {
        const sheet = args.sheet as string | number;
        const row = args.row as number;
        const col = args.col as number;

        const value = await wpsClient.getCellValue(sheet, row, col);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({ value }),
            },
          ],
        };
      }
    );

    // 设置单元格值
    this.registry.register(
      {
        name: 'wps_set_cell_value',
        description: '设置指定单元格的值',
        inputSchema: {
          type: 'object',
          properties: {
            sheet: {
              type: 'string',
              description: '工作表名称或索引',
            },
            row: {
              type: 'number',
              description: '行号（从1开始）',
            },
            col: {
              type: 'number',
              description: '列号（从1开始）',
            },
            value: {
              type: 'string',
              description: '要设置的值',
            },
          },
          required: ['sheet', 'row', 'col', 'value'],
        },
        category: ToolCategory.SPREADSHEET,
      },
      async (args) => {
        const sheet = args.sheet as string | number;
        const row = args.row as number;
        const col = args.col as number;
        const value = args.value;

        const success = await wpsClient.setCellValue(sheet, row, col, value);

        return {
          id: '',
          success,
          content: [
            {
              type: 'text',
              text: success ? '单元格值设置成功' : '单元格值设置失败',
            },
          ],
        };
      }
    );

    // 获取当前演示文稿信息
    this.registry.register(
      {
        name: 'wps_get_active_presentation',
        description: '获取当前打开的WPS演示文稿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.PRESENTATION,
      },
      async () => {
        const presentation = await wpsClient.getActivePresentation();

        return {
          id: '',
          success: presentation !== null,
          content: [
            {
              type: 'text',
              text: presentation
                ? JSON.stringify(presentation)
                : '没有打开的演示文稿',
            },
          ],
        };
      }
    );

    // 执行自定义WPS方法
    this.registry.register(
      {
        name: 'wps_execute_method',
        description: '执行自定义WPS API方法',
        inputSchema: {
          type: 'object',
          properties: {
            method: {
              type: 'string',
              description: 'API方法名',
            },
            params: {
              type: 'object',
              description: '方法参数',
            },
            appType: {
              type: 'string',
              description: '应用类型：wps（文字）、et（表格）、wpp（演示）',
              enum: ['wps', 'et', 'wpp'],
            },
          },
          required: ['method'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const method = args.method as string;
        const params = args.params as Record<string, unknown> | undefined;
        const appType = args.appType as string | undefined;

        const response = await wpsClient.executeMethod(
          method,
          params,
          appType as any
        );

        return {
          id: '',
          success: response.success,
          content: [
            {
              type: 'text',
              text: JSON.stringify(response),
            },
          ],
        };
      }
    );

    // ==================== 跨应用数据缓存工具 ====================
    // 解决macOS WPS加载项无法跨应用操作的P0问题
    // Excel读取数据 → 缓存到MCP Server → PPT获取缓存 → 创建演示文稿

    // 缓存数据
    this.registry.register(
      {
        name: 'wps_cache_data',
        description: '缓存数据到MCP Server，用于跨应用数据传递。例如：从Excel读取数据后缓存，然后在PPT中使用。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名，用于后续获取数据',
            },
            data: {
              type: 'object',
              description: '要缓存的数据（任意JSON对象）',
            },
            appType: {
              type: 'string',
              description: '数据来源应用类型：et（表格）、wps（文字）、wpp（演示）',
              enum: ['et', 'wps', 'wpp'],
            },
          },
          required: ['key', 'data'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const data = args.data;
        const appType = (args.appType as string) || 'unknown';

        WpsMcpServer.dataCache.set(key, {
          data,
          timestamp: Date.now(),
          appType,
        });

        logger.info(`Data cached: key=${key}, appType=${appType}`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                message: `数据已缓存，可在其他应用中通过 wps_get_cached_data 获取`,
                cacheSize: WpsMcpServer.dataCache.size,
              }),
            },
          ],
        };
      }
    );

    // 获取缓存数据
    this.registry.register(
      {
        name: 'wps_get_cached_data',
        description: '从MCP Server获取缓存的数据，用于跨应用数据传递。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名',
            },
          },
          required: ['key'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const cached = WpsMcpServer.dataCache.get(key);

        if (!cached) {
          return {
            id: '',
            success: false,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: false,
                  error: `缓存键 "${key}" 不存在`,
                  availableKeys: Array.from(WpsMcpServer.dataCache.keys()),
                }),
              },
            ],
          };
        }

        logger.info(`Cache hit: key=${key}, age=${Date.now() - cached.timestamp}ms`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                data: cached.data,
                appType: cached.appType,
                cachedAt: new Date(cached.timestamp).toISOString(),
              }),
            },
          ],
        };
      }
    );

    // 列出所有缓存
    this.registry.register(
      {
        name: 'wps_list_cache',
        description: '列出MCP Server中所有缓存的数据键名',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const cacheList = Array.from(WpsMcpServer.dataCache.entries()).map(([key, value]) => ({
          key,
          appType: value.appType,
          cachedAt: new Date(value.timestamp).toISOString(),
          ageMs: Date.now() - value.timestamp,
        }));

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                count: cacheList.length,
                caches: cacheList,
              }),
            },
          ],
        };
      }
    );

    // 清除缓存
    this.registry.register(
      {
        name: 'wps_clear_cache',
        description: '清除MCP Server中的缓存数据',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '要清除的缓存键名，不指定则清除所有缓存',
            },
          },
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string | undefined;

        if (key) {
          const deleted = WpsMcpServer.dataCache.delete(key);
          logger.info(`Cache cleared: key=${key}, deleted=${deleted}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: deleted ? `缓存 "${key}" 已清除` : `缓存 "${key}" 不存在`,
                }),
              },
            ],
          };
        } else {
          const count = WpsMcpServer.dataCache.size;
          WpsMcpServer.dataCache.clear();
          logger.info(`All cache cleared: count=${count}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `已清除所有缓存，共 ${count} 条`,
                }),
              },
            ],
          };
        }
      }
    );

    logger.info(`Registered ${this.registry.size} built-in tools`);
  }

  /**
   * 启动服务器
   */
  async start(): Promise<void> {
    if (this.isRunning) {
      logger.warn('Server is already running');
      return;
    }

    logger.info('Starting MCP Server...');

    // 注册内置Tools
    this.registerBuiltinTools();

    // 注册Excel、Word、PPT专业Tools - 这才是老王的核心功能
    this.registry.registerAll(allTools);
    logger.info(`Registered ${allTools.length} professional tools (Excel/Word/PPT)`);

    // 创建stdio传输层
    const transport = new StdioServerTransport();

    // 连接传输层
    await this.server.connect(transport);

    this.isRunning = true;
    logger.info('MCP Server started successfully');
  }

  /**
   * 停止服务器
   */
  async stop(): Promise<void> {
    if (!this.isRunning) {
      logger.warn('Server is not running');
      return;
    }

    logger.info('Stopping MCP Server...');

    await this.server.close();

    this.isRunning = false;
    logger.info('MCP Server stopped');
  }

  /**
   * 获取服务器状态
   */
  getStatus(): { running: boolean; toolCount: number } {
    return {
      running: this.isRunning,
      toolCount: this.registry.size,
    };
  }
}

// 导出单例创建函数
export const createMcpServer = (
  config?: Partial<McpServerConfig>
): WpsMcpServer => {
  return new WpsMcpServer(config);
};

export default WpsMcpServer;
