/**
 * Input: 转换工具参数
 * Output: 文档转换结果
 * Pos: 通用文档转换工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 文档转换Tools - 二狗的跨应用转换神器
 * 这个SB模块负责把Word/Excel/PPT转成各种格式
 *
 * 包含：
 * - wps_convert_to_pdf: 转换为PDF格式
 * - wps_convert_format: 格式互转（docx<->doc, xlsx<->xls, pptx<->ppt等）
 *
 * 艹，转换这种活就该让机器干，人工导出多无聊
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

/**
 * 根据文件扩展名判断应该用哪个WPS应用类型
 * 别问我为什么要写这个，问就是WPS三件套太TM分裂了
 */
const getAppTypeByExtension = (filePath: string): WpsAppType | null => {
  const ext = filePath.toLowerCase().split('.').pop();

  // Word文档
  if (['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'wps', 'wpt'].includes(ext || '')) {
    return WpsAppType.WRITER;
  }

  // Excel表格
  if (['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'et', 'ett'].includes(ext || '')) {
    return WpsAppType.SPREADSHEET;
  }

  // PPT演示
  if (['ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt'].includes(ext || '')) {
    return WpsAppType.PRESENTATION;
  }

  return null;
};

/**
 * 根据输出格式获取对应的文件格式代码
 * WPS的格式代码和微软Office基本兼容，但也有自己的一套
 */
const getFormatCode = (format: string, appType: WpsAppType): number => {
  const formatLower = format.toLowerCase();

  switch (appType) {
    case WpsAppType.WRITER:
      // Word文档格式代码
      // wdFormatDocument = 0 (.doc)
      // wdFormatDocumentDefault = 16 (.docx)
      // wdFormatPDF = 17
      // wdFormatRTF = 6
      // wdFormatXPS = 18
      // wdFormatHTML = 8
      const wordFormats: Record<string, number> = {
        'doc': 0,
        'docx': 16,
        'pdf': 17,
        'rtf': 6,
        'xps': 18,
        'html': 8,
        'htm': 8,
        'txt': 2,  // wdFormatText
        'xml': 11, // wdFormatXML
      };
      return wordFormats[formatLower] ?? 16; // 默认docx

    case WpsAppType.SPREADSHEET:
      // Excel工作簿格式代码
      // xlWorkbookNormal = -4143 (.xls)
      // xlOpenXMLWorkbook = 51 (.xlsx)
      // xlTypePDF = 0 (导出PDF时用)
      // xlCSV = 6
      // xlHtml = 44
      const excelFormats: Record<string, number> = {
        'xls': -4143,
        'xlsx': 51,
        'xlsm': 52,
        'xlsb': 50,
        'csv': 6,
        'html': 44,
        'htm': 44,
        'pdf': 0, // 特殊处理
        'xps': 1, // 特殊处理
      };
      return excelFormats[formatLower] ?? 51; // 默认xlsx

    case WpsAppType.PRESENTATION:
      // PPT演示格式代码
      // ppSaveAsPresentation = 1 (.ppt)
      // ppSaveAsOpenXMLPresentation = 24 (.pptx)
      // ppSaveAsPDF = 32
      // ppSaveAsHTML = 12
      const pptFormats: Record<string, number> = {
        'ppt': 1,
        'pptx': 24,
        'pptm': 25,
        'pdf': 32,
        'xps': 33,
        'html': 12,
        'htm': 12,
        'png': 18,
        'jpg': 17,
        'jpeg': 17,
        'gif': 16,
        'bmp': 19,
      };
      return pptFormats[formatLower] ?? 24; // 默认pptx

    default:
      return -1;
  }
};

/**
 * 转换为PDF格式 - 老王都得佩服这个功能
 * 支持Word、Excel、PPT一键转PDF，省得用户到处找在线转换
 */
export const convertToPdfDefinition: ToolDefinition = {
  name: 'wps_convert_to_pdf',
  description: `将当前文档转换为PDF格式。支持Word、Excel、PPT文档。

使用场景：
- "把这个Word转成PDF"
- "导出PDF给客户看"
- "把表格保存成PDF格式"
- "PPT转PDF方便打印"

特点：
- 自动检测当前打开的文档类型
- 可以指定输出路径，不指定则使用原文件名.pdf
- 保持原文档格式和排版`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      outputPath: {
        type: 'string',
        description: 'PDF输出路径（包含文件名），如不指定则使用原文件路径，把扩展名改为.pdf',
      },
      openAfterExport: {
        type: 'boolean',
        description: '导出后是否自动打开PDF，默认false',
      },
    },
    required: [],
  },
};

export const convertToPdfHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { outputPath, openAfterExport } = args as {
    outputPath?: string;
    openAfterExport?: boolean;
  };

  try {
    // 调用WPS加载项执行转换
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      sourcePath: string;
      outputPath: string;
      appType: string;
      pageCount?: number;
    }>(
      'convertToPDF',
      {
        outputPath: outputPath || '',
        openAfterExport: openAfterExport || false,
      }
      // 不指定appType，让WPS加载项自动检测当前活动的应用
    );

    if (response.success && response.data) {
      const result = response.data;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `PDF导出成功！
源文件: ${result.sourcePath}
输出路径: ${result.outputPath}
文档类型: ${result.appType}${result.pageCount ? `\n页数: ${result.pageCount}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `PDF导出失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `PDF导出出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 格式互转 - 老王的格式转换大法
 * docx<->doc, xlsx<->xls, pptx<->ppt，想怎么转怎么转
 */
export const convertFormatDefinition: ToolDefinition = {
  name: 'wps_convert_format',
  description: `将当前文档转换为其他格式。

使用场景：
- "把docx转成doc格式"
- "保存为旧版Excel格式兼容老系统"
- "把PPT转成pptx"
- "导出为RTF格式"
- "转换成HTML网页格式"

支持的格式：
- Word: doc, docx, rtf, txt, html, xml
- Excel: xls, xlsx, xlsm, xlsb, csv, html
- PPT: ppt, pptx, pptm, html, png, jpg, gif, bmp

注意：
- 转换时会保留原文档，另存为新格式
- 某些格式转换可能会丢失部分效果（如doc不支持的新特性）`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      targetFormat: {
        type: 'string',
        description: '目标格式扩展名，如 doc, xlsx, ppt, rtf, csv, html 等',
      },
      outputPath: {
        type: 'string',
        description: '输出路径（包含文件名），如不指定则使用原文件名改为新扩展名',
      },
    },
    required: ['targetFormat'],
  },
};

export const convertFormatHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { targetFormat, outputPath } = args as {
    targetFormat: string;
    outputPath?: string;
  };

  if (!targetFormat || targetFormat.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '目标格式不能为空！你想转成啥格式得告诉我啊' }],
      error: '目标格式为空',
    };
  }

  try {
    // 调用WPS加载项执行格式转换
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      sourcePath: string;
      sourceFormat: string;
      targetFormat: string;
      outputPath: string;
      appType: string;
    }>(
      'convertFormat',
      {
        targetFormat: targetFormat.toLowerCase().replace(/^\./, ''), // 去掉开头的点
        outputPath: outputPath || '',
      }
      // 不指定appType，让WPS加载项自动检测
    );

    if (response.success && response.data) {
      const result = response.data;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `格式转换成功！
源文件: ${result.sourcePath}
源格式: ${result.sourceFormat}
目标格式: ${result.targetFormat}
输出路径: ${result.outputPath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `格式转换失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `格式转换出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有转换相关的Tools
 * 二狗出品，转换必备
 */
export const convertTools: RegisteredTool[] = [
  { definition: convertToPdfDefinition, handler: convertToPdfHandler },
  { definition: convertFormatDefinition, handler: convertFormatHandler },
];

// 导出工具函数，别的模块可能用得上
export { getAppTypeByExtension, getFormatCode };

export default convertTools;
