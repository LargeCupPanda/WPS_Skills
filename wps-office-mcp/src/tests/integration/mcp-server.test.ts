/**
 * MCP Server集成测试 - 陈十三出品
 *
 * 艹，这个测试模块是用来测试mcp-server.ts的集成功能
 * 测试服务启动、Tool调用流程、请求处理
 *
 * @author 陈十三
 * @date 2026-01-24
 */

import { WpsMcpServer } from '../../server/mcp-server';
import { ToolRegistry } from '../../server/tool-registry';
import { ToolDefinition, ToolCategory, ToolHandler } from '../../types/tools';

// Mock MCP SDK
jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
  Server: jest.fn().mockImplementation(() => ({
    setRequestHandler: jest.fn(),
    connect: jest.fn().mockResolvedValue(undefined),
    close: jest.fn().mockResolvedValue(undefined),
    onerror: null,
  })),
}));

jest.mock('@modelcontextprotocol/sdk/server/stdio.js', () => ({
  StdioServerTransport: jest.fn().mockImplementation(() => ({})),
}));

jest.mock('@modelcontextprotocol/sdk/types.js', () => ({
  CallToolRequestSchema: { method: 'tools/call' },
  ListToolsRequestSchema: { method: 'tools/list' },
  ErrorCode: { InternalError: -32603, InvalidParams: -32602 },
  McpError: class MockMcpError extends Error {
    constructor(public code: number, message: string) {
      super(message);
    }
  },
}));

// Mock logger
jest.mock('../../utils/logger', () => ({
  log: {
    info: jest.fn(),
    debug: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
  createChildLogger: jest.fn(() => ({
    info: jest.fn(),
    debug: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  })),
}));

// Mock wpsClient
jest.mock('../../client/wps-client', () => ({
  wpsClient: {
    checkConnection: jest.fn().mockResolvedValue(true),
    getStatus: jest.fn().mockReturnValue({ connected: true }),
  },
}));

// Mock错误类
jest.mock('../../utils/error', () => ({
  McpError: class McpError extends Error {
    constructor(message: string, public code: string, public details?: unknown) {
      super(message);
      this.name = 'McpError';
    }
  },
  ErrorCode: {
    TOOL_NOT_FOUND: 'TOOL_NOT_FOUND',
    INVALID_PARAMS: 'INVALID_PARAMS',
    TOOL_ALREADY_REGISTERED: 'TOOL_ALREADY_REGISTERED',
  },
  ToolNotFoundError: class ToolNotFoundError extends Error {
    constructor(toolName: string) {
      super(`Tool not found: ${toolName}`);
      this.name = 'ToolNotFoundError';
    }
  },
  ToolExecutionError: class ToolExecutionError extends Error {
    constructor(toolName: string, cause: Error) {
      super(`Tool execution failed: ${toolName} - ${cause.message}`);
      this.name = 'ToolExecutionError';
    }
  },
  InvalidParamsError: class InvalidParamsError extends Error {
    constructor(message: string, public details?: unknown) {
      super(message);
      this.name = 'InvalidParamsError';
    }
  },
}));

// Mock uuid
jest.mock('uuid', () => ({
  v4: jest.fn(() => 'mock-uuid-integration-test'),
}));

describe('WpsMcpServer集成测试', () => {
  let server: WpsMcpServer;
  let registry: ToolRegistry;

  // 测试用的Tool定义
  const testTools: Array<{ definition: ToolDefinition; handler: ToolHandler }> = [
    {
      definition: {
        name: 'integration.test.echo',
        description: '回显测试Tool',
        inputSchema: {
          type: 'object',
          properties: {
            message: { type: 'string', description: '要回显的消息' },
          },
          required: ['message'],
        },
        category: ToolCategory.COMMON,
      },
      handler: async (args) => ({
        id: 'echo-result',
        success: true,
        content: [{ type: 'text', text: `Echo: ${args.message}` }],
      }),
    },
    {
      definition: {
        name: 'integration.test.add',
        description: '加法测试Tool',
        inputSchema: {
          type: 'object',
          properties: {
            a: { type: 'number', description: '第一个数' },
            b: { type: 'number', description: '第二个数' },
          },
          required: ['a', 'b'],
        },
        category: ToolCategory.SPREADSHEET,
      },
      handler: async (args) => ({
        id: 'add-result',
        success: true,
        content: [{ type: 'text', text: `Result: ${Number(args.a) + Number(args.b)}` }],
      }),
    },
    {
      definition: {
        name: 'integration.test.fail',
        description: '故意失败的测试Tool',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      handler: async () => {
        throw new Error('Intentional failure for testing');
      },
    },
  ];

  beforeEach(() => {
    // 清理并重新注册
    registry = ToolRegistry.getInstance();
    registry.clear();

    // 注册测试Tools
    testTools.forEach(({ definition, handler }) => {
      registry.register(definition, handler);
    });
  });

  afterEach(() => {
    registry.clear();
  });

  describe('服务器创建', () => {
    it('应该使用默认配置创建服务器', () => {
      server = new WpsMcpServer();

      // 服务器应该被创建
      expect(server).toBeDefined();
    });

    it('应该接受自定义配置', () => {
      server = new WpsMcpServer({
        name: 'custom-server',
        version: '2.0.0',
        debug: true,
      });

      expect(server).toBeDefined();
    });
  });

  describe('Tool注册验证', () => {
    it('应该能获取所有已注册的Tools', () => {
      const { tools } = registry.listTools();

      expect(tools).toHaveLength(3);
      expect(tools.map((t) => t.name)).toContain('integration.test.echo');
      expect(tools.map((t) => t.name)).toContain('integration.test.add');
      expect(tools.map((t) => t.name)).toContain('integration.test.fail');
    });

    it('应该能按分类获取Tools', () => {
      const spreadsheetTools = registry.getToolsByCategory(ToolCategory.SPREADSHEET);
      const commonTools = registry.getToolsByCategory(ToolCategory.COMMON);

      expect(spreadsheetTools).toHaveLength(1);
      expect(spreadsheetTools[0].name).toBe('integration.test.add');

      expect(commonTools.length).toBeGreaterThanOrEqual(2);
    });
  });

  describe('Tool调用流程', () => {
    it('应该成功调用echo Tool', async () => {
      const request = ToolRegistry.createRequest('integration.test.echo', {
        message: 'Hello Integration Test',
      });

      const result = await registry.callTool(request);

      expect(result.success).toBe(true);
      expect(result.content).toHaveLength(1);
      expect(result.content[0].text).toBe('Echo: Hello Integration Test');
    });

    it('应该成功调用add Tool进行计算', async () => {
      const request = ToolRegistry.createRequest('integration.test.add', {
        a: 10,
        b: 20,
      });

      const result = await registry.callTool(request);

      expect(result.success).toBe(true);
      expect(result.content[0].text).toBe('Result: 30');
    });

    it('应该正确处理Tool执行失败', async () => {
      const request = ToolRegistry.createRequest('integration.test.fail', {});

      const result = await registry.callTool(request);

      expect(result.success).toBe(false);
      expect(result.error).toContain('Intentional failure');
    });

    it('应该拒绝调用不存在的Tool', async () => {
      const request = ToolRegistry.createRequest('nonexistent.tool', {});

      await expect(registry.callTool(request)).rejects.toThrow();
    });
  });

  describe('参数验证', () => {
    it('缺少必填参数应该返回失败结果', async () => {
      const request = ToolRegistry.createRequest('integration.test.echo', {});

      const result = await registry.callTool(request);

      expect(result.success).toBe(false);
      expect(result.error).toContain('Missing required parameter');
    });

    it('提供所有必填参数应该成功', async () => {
      const request = ToolRegistry.createRequest('integration.test.add', {
        a: 5,
        b: 3,
      });

      const result = await registry.callTool(request);

      expect(result.success).toBe(true);
    });
  });

  describe('并发调用测试', () => {
    it('应该能处理多个并发Tool调用', async () => {
      const requests = [
        ToolRegistry.createRequest('integration.test.echo', { message: 'First' }),
        ToolRegistry.createRequest('integration.test.echo', { message: 'Second' }),
        ToolRegistry.createRequest('integration.test.add', { a: 1, b: 2 }),
        ToolRegistry.createRequest('integration.test.add', { a: 3, b: 4 }),
      ];

      const results = await Promise.all(requests.map((req) => registry.callTool(req)));

      expect(results).toHaveLength(4);
      expect(results.every((r) => r.success)).toBe(true);
    });
  });

  describe('Tool动态注册和注销', () => {
    it('应该能在运行时动态注册新Tool', async () => {
      const dynamicTool: ToolDefinition = {
        name: 'dynamic.tool',
        description: '动态注册的Tool',
        inputSchema: {
          type: 'object',
          properties: { value: { type: 'string' } },
          required: ['value'],
        },
      };

      const dynamicHandler: ToolHandler = async (args) => ({
        id: 'dynamic',
        success: true,
        content: [{ type: 'text', text: `Dynamic: ${args.value}` }],
      });

      registry.register(dynamicTool, dynamicHandler);

      const request = ToolRegistry.createRequest('dynamic.tool', { value: 'test' });
      const result = await registry.callTool(request);

      expect(result.success).toBe(true);
      expect(result.content[0].text).toBe('Dynamic: test');
    });

    it('应该能注销已注册的Tool', () => {
      expect(registry.hasTool('integration.test.echo')).toBe(true);

      const unregistered = registry.unregister('integration.test.echo');

      expect(unregistered).toBe(true);
      expect(registry.hasTool('integration.test.echo')).toBe(false);
    });
  });
});

describe('端到端场景测试', () => {
  let registry: ToolRegistry;

  beforeEach(() => {
    registry = ToolRegistry.getInstance();
    registry.clear();
  });

  it('模拟完整的Excel公式设置流程', async () => {
    // 注册模拟的Excel Tool
    const setFormulaDefinition: ToolDefinition = {
      name: 'excel.setFormula',
      description: '设置Excel单元格公式',
      inputSchema: {
        type: 'object',
        properties: {
          range: { type: 'string', description: '单元格范围' },
          formula: { type: 'string', description: '公式' },
        },
        required: ['range', 'formula'],
      },
      category: ToolCategory.SPREADSHEET,
    };

    const setFormulaHandler: ToolHandler = async (args) => {
      const { range, formula } = args as { range: string; formula: string };

      // 验证公式格式
      if (!formula.startsWith('=')) {
        return {
          id: 'formula-error',
          success: false,
          content: [{ type: 'text', text: '公式必须以=开头，憨批' }],
          error: 'Invalid formula format',
        };
      }

      // 模拟成功设置公式
      return {
        id: 'formula-success',
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              range,
              formula,
              result: '计算结果会显示在这里',
            }),
          },
        ],
      };
    };

    registry.register(setFormulaDefinition, setFormulaHandler);

    // 测试正确的公式
    const successRequest = ToolRegistry.createRequest('excel.setFormula', {
      range: 'A1',
      formula: '=SUM(B1:B10)',
    });

    const successResult = await registry.callTool(successRequest);
    expect(successResult.success).toBe(true);

    // 测试错误的公式格式
    const failRequest = ToolRegistry.createRequest('excel.setFormula', {
      range: 'A1',
      formula: 'SUM(B1:B10)', // 缺少=
    });

    const failResult = await registry.callTool(failRequest);
    expect(failResult.success).toBe(false);
    expect(failResult.error).toContain('Invalid formula');
  });

  it('模拟完整的Word文档操作流程', async () => {
    // 注册模拟的Word Tools
    const insertTextDefinition: ToolDefinition = {
      name: 'word.insertText',
      description: '在Word文档中插入文本',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: '要插入的文本' },
          position: { type: 'number', description: '插入位置' },
        },
        required: ['text'],
      },
      category: ToolCategory.DOCUMENT,
    };

    const insertTextHandler: ToolHandler = async (args) => {
      const { text, position } = args as { text: string; position?: number };

      return {
        id: 'insert-result',
        success: true,
        content: [
          {
            type: 'text',
            text: `已在位置${position ?? '末尾'}插入${text.length}个字符`,
          },
        ],
      };
    };

    registry.register(insertTextDefinition, insertTextHandler);

    // 测试插入文本
    const request = ToolRegistry.createRequest('word.insertText', {
      text: '这是一段测试文本，用来验证Word文档操作功能。',
      position: 100,
    });

    const result = await registry.callTool(request);

    expect(result.success).toBe(true);
    expect(result.content[0].text).toContain('已在位置100插入');
  });
});
