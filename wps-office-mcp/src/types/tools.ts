/**
 * Tool类型定义 - 老王出品，必属精品
 * 定义MCP Tool的所有相关类型，别tm乱改
 */

/**
 * Tool参数的JSON Schema类型定义
 */
export interface ToolParameterSchema {
  type: 'string' | 'number' | 'boolean' | 'object' | 'array';
  description?: string;
  enum?: string[];
  items?: ToolParameterSchema;
  properties?: Record<string, ToolParameterSchema>;
  required?: string[];
  default?: unknown;
}

/**
 * Tool输入Schema定义
 */
export interface ToolInputSchema {
  type: 'object';
  properties: Record<string, ToolParameterSchema>;
  required?: string[];
}

/**
 * Tool定义 - MCP协议规范的Tool结构
 */
export interface ToolDefinition {
  /** Tool名称，唯一标识符 */
  name: string;
  /** Tool描述，告诉AI这玩意儿能干啥 */
  description: string;
  /** 输入参数Schema */
  inputSchema: ToolInputSchema;
  /** Tool分类，方便管理 */
  category?: ToolCategory;
}

/**
 * Tool分类枚举 - WPS三件套对应的分类
 */
export enum ToolCategory {
  /** 文档操作 - Word那一套 */
  DOCUMENT = 'document',
  /** 表格操作 - Excel那一套 */
  SPREADSHEET = 'spreadsheet',
  /** 演示操作 - PPT那一套 */
  PRESENTATION = 'presentation',
  /** 通用操作 - 跨应用的通用功能 */
  COMMON = 'common',
}

/**
 * Tool调用请求
 */
export interface ToolCallRequest {
  /** 请求ID，用于追踪 */
  id: string;
  /** Tool名称 */
  name: string;
  /** 调用参数 */
  arguments: Record<string, unknown>;
}

/**
 * Tool调用结果
 */
export interface ToolCallResult {
  /** 对应的请求ID */
  id: string;
  /** 是否成功 */
  success: boolean;
  /** 返回的内容 */
  content: ToolContent[];
  /** 错误信息（失败时） */
  error?: string;
}

/**
 * Tool返回内容类型
 */
export interface ToolContent {
  type: 'text' | 'image' | 'resource';
  text?: string;
  data?: string;
  mimeType?: string;
}

/**
 * Tool处理函数类型
 */
export type ToolHandler = (
  args: Record<string, unknown>
) => Promise<ToolCallResult>;

/**
 * 注册的Tool完整信息
 */
export interface RegisteredTool {
  definition: ToolDefinition;
  handler: ToolHandler;
}

/**
 * Tool列表响应
 */
export interface ListToolsResponse {
  tools: ToolDefinition[];
}

/**
 * Tool调用响应
 */
export interface CallToolResponse {
  content: ToolContent[];
  isError?: boolean;
}
