/**
 * Tools定义和注册测试 - 陈十三出品
 *
 * 艹，这个测试模块是用来测试tool-registry.ts和tools类型的
 * 测试Tool参数验证、Tool注册、Tool调用
 *
 * @author 陈十三
 * @date 2026-01-24
 */

import { ToolRegistry } from '../../server/tool-registry';
import {
  ToolDefinition,
  ToolCallRequest,
  ToolCallResult,
  ToolCategory,
  ToolHandler,
} from '../../types/tools';

// Mock logger - 不想看到一堆日志输出
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

// Mock错误类
jest.mock('../../utils/error', () => ({
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
  McpError: class McpError extends Error {
    constructor(
      message: string,
      public code: string,
      public details?: unknown
    ) {
      super(message);
      this.name = 'McpError';
    }
  },
  ErrorCode: {
    TOOL_ALREADY_REGISTERED: 'TOOL_ALREADY_REGISTERED',
  },
}));

// Mock uuid
jest.mock('uuid', () => ({
  v4: jest.fn(() => 'mock-uuid-12345'),
}));

describe('ToolRegistry', () => {
  let registry: ToolRegistry;

  // 测试用的Tool定义
  const testToolDefinition: ToolDefinition = {
    name: 'test.tool',
    description: '这是个测试用的Tool，别TM乱用',
    inputSchema: {
      type: 'object',
      properties: {
        param1: { type: 'string', description: '参数1' },
        param2: { type: 'number', description: '参数2' },
      },
      required: ['param1'],
    },
    category: ToolCategory.COMMON,
  };

  // 测试用的Handler
  const testHandler: ToolHandler = jest.fn(async (args) => ({
    id: 'test-id',
    success: true,
    content: [{ type: 'text' as const, text: `Received: ${JSON.stringify(args)}` }],
  }));

  beforeEach(() => {
    // 每次测试前获取新的实例（单例模式需要清理）
    registry = ToolRegistry.getInstance();
    registry.clear();
  });

  describe('单例模式 - 全局唯一', () => {
    it('getInstance应该始终返回同一个实例', () => {
      const instance1 = ToolRegistry.getInstance();
      const instance2 = ToolRegistry.getInstance();

      expect(instance1).toBe(instance2);
    });
  });

  describe('Tool注册 - register', () => {
    it('应该成功注册Tool', () => {
      expect(() => {
        registry.register(testToolDefinition, testHandler);
      }).not.toThrow();

      expect(registry.hasTool('test.tool')).toBe(true);
      expect(registry.size).toBe(1);
    });

    it('重复注册同名Tool应该抛出错误', () => {
      registry.register(testToolDefinition, testHandler);

      expect(() => {
        registry.register(testToolDefinition, testHandler);
      }).toThrow();
    });

    it('应该正确分类Tool', () => {
      const excelTool: ToolDefinition = {
        name: 'excel.test',
        description: 'Excel测试Tool',
        inputSchema: { type: 'object', properties: {} },
        category: ToolCategory.SPREADSHEET,
      };

      registry.register(excelTool, testHandler);

      const spreadsheetTools = registry.getToolsByCategory(ToolCategory.SPREADSHEET);
      expect(spreadsheetTools).toHaveLength(1);
      expect(spreadsheetTools[0].name).toBe('excel.test');
    });

    it('没有指定category应该使用COMMON', () => {
      const noCategory: ToolDefinition = {
        name: 'no.category',
        description: '没有分类的Tool',
        inputSchema: { type: 'object', properties: {} },
      };

      registry.register(noCategory, testHandler);

      const commonTools = registry.getToolsByCategory(ToolCategory.COMMON);
      expect(commonTools.some((t) => t.name === 'no.category')).toBe(true);
    });
  });

  describe('批量注册 - registerAll', () => {
    it('应该一次注册多个Tools', () => {
      const tools = [
        { definition: { ...testToolDefinition, name: 'tool1' }, handler: testHandler },
        { definition: { ...testToolDefinition, name: 'tool2' }, handler: testHandler },
        { definition: { ...testToolDefinition, name: 'tool3' }, handler: testHandler },
      ];

      registry.registerAll(tools);

      expect(registry.size).toBe(3);
      expect(registry.hasTool('tool1')).toBe(true);
      expect(registry.hasTool('tool2')).toBe(true);
      expect(registry.hasTool('tool3')).toBe(true);
    });
  });

  describe('Tool注销 - unregister', () => {
    it('应该成功注销已注册的Tool', () => {
      registry.register(testToolDefinition, testHandler);
      expect(registry.hasTool('test.tool')).toBe(true);

      const result = registry.unregister('test.tool');

      expect(result).toBe(true);
      expect(registry.hasTool('test.tool')).toBe(false);
    });

    it('注销不存在的Tool应该返回false', () => {
      const result = registry.unregister('nonexistent.tool');
      expect(result).toBe(false);
    });
  });

  describe('获取Tool - getTool', () => {
    it('应该返回已注册Tool的完整信息', () => {
      registry.register(testToolDefinition, testHandler);

      const tool = registry.getTool('test.tool');

      expect(tool).toBeDefined();
      expect(tool?.definition).toEqual(testToolDefinition);
      expect(tool?.handler).toBe(testHandler);
    });

    it('获取不存在的Tool应该返回undefined', () => {
      const tool = registry.getTool('nonexistent.tool');
      expect(tool).toBeUndefined();
    });
  });

  describe('列出所有Tools - listTools', () => {
    it('应该返回所有Tool定义', () => {
      const tool1 = { ...testToolDefinition, name: 'list.tool1' };
      const tool2 = { ...testToolDefinition, name: 'list.tool2' };

      registry.register(tool1, testHandler);
      registry.register(tool2, testHandler);

      const result = registry.listTools();

      expect(result.tools).toHaveLength(2);
      expect(result.tools.map((t) => t.name)).toContain('list.tool1');
      expect(result.tools.map((t) => t.name)).toContain('list.tool2');
    });
  });

  describe('Tool调用 - callTool', () => {
    it('应该成功调用已注册的Tool', async () => {
      const mockHandler: ToolHandler = jest.fn(async () => ({
        id: 'result-id',
        success: true,
        content: [{ type: 'text' as const, text: 'Success!' }],
      }));

      registry.register(testToolDefinition, mockHandler);

      const request: ToolCallRequest = {
        id: 'call-123',
        name: 'test.tool',
        arguments: { param1: 'value1' },
      };

      const result = await registry.callTool(request);

      expect(result.success).toBe(true);
      expect(result.id).toBe('call-123');
      expect(mockHandler).toHaveBeenCalledWith({ param1: 'value1' });
    });

    it('调用不存在的Tool应该抛出错误', async () => {
      const request: ToolCallRequest = {
        id: 'call-456',
        name: 'nonexistent.tool',
        arguments: {},
      };

      await expect(registry.callTool(request)).rejects.toThrow();
    });

    it('缺少必填参数应该返回失败结果', async () => {
      registry.register(testToolDefinition, testHandler);

      const request: ToolCallRequest = {
        id: 'call-789',
        name: 'test.tool',
        arguments: { param2: 123 }, // 缺少必填的param1
      };

      const result = await registry.callTool(request);

      expect(result.success).toBe(false);
      expect(result.error).toContain('Missing required parameter');
    });

    it('Handler抛出错误应该返回失败结果', async () => {
      const errorHandler: ToolHandler = jest.fn(async () => {
        throw new Error('Handler exploded!');
      });

      registry.register({ ...testToolDefinition, name: 'error.tool' }, errorHandler);

      const request: ToolCallRequest = {
        id: 'call-error',
        name: 'error.tool',
        arguments: { param1: 'value' },
      };

      const result = await registry.callTool(request);

      expect(result.success).toBe(false);
      expect(result.error).toContain('Handler exploded');
    });
  });

  describe('静态方法 - createRequest', () => {
    it('应该创建有效的Tool调用请求', () => {
      const request = ToolRegistry.createRequest('my.tool', { arg1: 'value1' });

      expect(request.id).toBe('mock-uuid-12345');
      expect(request.name).toBe('my.tool');
      expect(request.arguments).toEqual({ arg1: 'value1' });
    });
  });

  describe('清空操作 - clear', () => {
    it('应该清空所有已注册的Tools', () => {
      registry.register({ ...testToolDefinition, name: 'clear.tool1' }, testHandler);
      registry.register({ ...testToolDefinition, name: 'clear.tool2' }, testHandler);

      expect(registry.size).toBe(2);

      registry.clear();

      expect(registry.size).toBe(0);
      expect(registry.hasTool('clear.tool1')).toBe(false);
      expect(registry.hasTool('clear.tool2')).toBe(false);
    });
  });
});

describe('ToolCategory枚举', () => {
  it('应该包含所有WPS应用类型', () => {
    expect(ToolCategory.DOCUMENT).toBe('document');
    expect(ToolCategory.SPREADSHEET).toBe('spreadsheet');
    expect(ToolCategory.PRESENTATION).toBe('presentation');
    expect(ToolCategory.COMMON).toBe('common');
  });
});

describe('Tool类型验证', () => {
  it('ToolDefinition应该有正确的结构', () => {
    const definition: ToolDefinition = {
      name: 'valid.tool',
      description: 'A valid tool',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string' },
        },
        required: ['text'],
      },
      category: ToolCategory.DOCUMENT,
    };

    // 类型检查通过即可
    expect(definition.name).toBeDefined();
    expect(definition.description).toBeDefined();
    expect(definition.inputSchema).toBeDefined();
    expect(definition.inputSchema.type).toBe('object');
  });

  it('ToolCallResult应该有正确的结构', () => {
    const result: ToolCallResult = {
      id: 'test-id',
      success: true,
      content: [
        { type: 'text', text: 'Hello' },
        { type: 'image', data: 'base64...', mimeType: 'image/png' },
      ],
    };

    expect(result).toBeValidToolResult();
  });
});
