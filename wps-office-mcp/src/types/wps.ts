/**
 * WPS相关类型定义 - 老王的WPS客户端类型
 * 跟WPS加载项通信用的类型，别tm搞混了
 */

/**
 * WPS应用类型枚举
 */
export enum WpsAppType {
  /** WPS文字 - 对标Word */
  WRITER = 'wps',
  /** WPS表格 - 对标Excel */
  SPREADSHEET = 'et',
  /** WPS演示 - 对标PPT */
  PRESENTATION = 'wpp',
}

/**
 * WPS加载项API端点配置
 */
export interface WpsEndpointConfig {
  /** 基础URL，默认localhost:23333 */
  baseUrl: string;
  /** 超时时间（毫秒） */
  timeout: number;
}

/**
 * WPS API调用请求
 */
export interface WpsApiRequest {
  /** 调用的方法名 */
  method: string;
  /** 方法参数 */
  params?: Record<string, unknown>;
  /** 目标应用类型 */
  appType?: WpsAppType;
}

/**
 * WPS API调用响应
 */
export interface WpsApiResponse<T = unknown> {
  /** 是否成功 */
  success: boolean;
  /** 返回数据 */
  data?: T;
  /** 错误信息 */
  error?: string;
  /** 错误码 */
  errorCode?: number;
}

/**
 * 文档信息
 */
export interface DocumentInfo {
  /** 文档名称 */
  name: string;
  /** 文档路径 */
  path: string;
  /** 是否已保存 */
  saved: boolean;
  /** 是否只读 */
  readOnly: boolean;
}

/**
 * 工作簿信息
 */
export interface WorkbookInfo {
  /** 工作簿名称 */
  name: string;
  /** 工作簿路径 */
  path: string;
  /** 工作表列表 */
  sheets: SheetInfo[];
  /** 当前活动工作表索引 */
  activeSheetIndex: number;
}

/**
 * 工作表信息
 */
export interface SheetInfo {
  /** 工作表名称 */
  name: string;
  /** 工作表索引 */
  index: number;
  /** 是否可见 */
  visible: boolean;
}

/**
 * 单元格范围
 */
export interface CellRange {
  /** 起始行（1开始） */
  startRow: number;
  /** 起始列（1开始） */
  startCol: number;
  /** 结束行 */
  endRow: number;
  /** 结束列 */
  endCol: number;
}

/**
 * 单元格数据
 */
export interface CellData {
  /** 行号 */
  row: number;
  /** 列号 */
  col: number;
  /** 值 */
  value: string | number | boolean | null;
  /** 公式（如果有） */
  formula?: string;
  /** 格式化后的文本 */
  text?: string;
}

/**
 * 演示文稿信息
 */
export interface PresentationInfo {
  /** 演示文稿名称 */
  name: string;
  /** 演示文稿路径 */
  path: string;
  /** 幻灯片数量 */
  slideCount: number;
  /** 当前幻灯片索引 */
  currentSlideIndex: number;
}

/**
 * 幻灯片信息
 */
export interface SlideInfo {
  /** 幻灯片索引 */
  index: number;
  /** 幻灯片布局 */
  layout: string;
  /** 是否隐藏 */
  hidden: boolean;
}

/**
 * 文本操作参数
 */
export interface TextOperationParams {
  /** 要插入/替换的文本 */
  text: string;
  /** 位置（可选） */
  position?: number;
  /** 是否替换选中内容 */
  replaceSelection?: boolean;
}

/**
 * 格式化参数
 */
export interface FormatParams {
  /** 字体名称 */
  fontName?: string;
  /** 字体大小 */
  fontSize?: number;
  /** 是否加粗 */
  bold?: boolean;
  /** 是否斜体 */
  italic?: boolean;
  /** 是否下划线 */
  underline?: boolean;
  /** 字体颜色（RGB） */
  color?: string;
  /** 背景颜色（RGB） */
  backgroundColor?: string;
}

/**
 * WPS客户端状态
 */
export interface WpsClientStatus {
  /** 是否已连接 */
  connected: boolean;
  /** 当前活动应用 */
  activeApp?: WpsAppType;
  /** 最后心跳时间 */
  lastHeartbeat?: Date;
  /** 错误信息 */
  error?: string;
}
