/**
 * Excel数据处理Tools - 老王的数据清洗神器
 * 处理那些乱七八糟的数据，去重、去空格、格式统一等
 *
 * 包含：
 * - wps_excel_read_range: 读取指定范围数据
 * - wps_excel_write_range: 写入数据到指定范围
 * - wps_excel_clean_data: 数据清洗（核心功能）
 * - wps_excel_remove_duplicates: 删除重复行
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
 * 读取指定范围的单元格数据
 * 先读后写，老王的原则是搞清楚数据长啥样再动手
 */
export const readRangeDefinition: ToolDefinition = {
  name: 'wps_excel_read_range',
  description: '读取Excel指定范围的单元格数据，返回二维数组格式的数据。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要读取的范围，如 A1:C10、B2:D5',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
      include_header: {
        type: 'boolean',
        description: '是否将第一行作为表头返回，默认false',
      },
    },
    required: ['range'],
  },
};

export const readRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, sheet, include_header } = args as {
    range: string;
    sheet?: string;
    include_header?: boolean;
  };

  try {
    const response = await wpsClient.getRangeData(sheet || 0, range);

    if (!response || response.length === 0) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `范围 ${range} 是空的，没有数据`,
          },
        ],
      };
    }

    // 格式化输出
    let output = `范围 ${range} 的数据（${response.length}行 x ${response[0]?.length || 0}列）：\n\n`;

    if (include_header && response.length > 0) {
      const headers = response[0] as unknown[];
      output += `表头: ${headers.join(' | ')}\n`;
      output += '-'.repeat(50) + '\n';

      for (let i = 1; i < response.length; i++) {
        output += `第${i}行: ${(response[i] as unknown[]).join(' | ')}\n`;
      }
    } else {
      response.forEach((row, index) => {
        output += `第${index + 1}行: ${(row as unknown[]).join(' | ')}\n`;
      });
    }

    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: output }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `读取数据出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 向指定范围写入数据
 * 批量写入数据，比一个个单元格设置快多了
 */
export const writeRangeDefinition: ToolDefinition = {
  name: 'wps_excel_write_range',
  description: '向Excel指定范围写入数据。数据格式为二维数组，从指定单元格开始向右下方填充。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '起始单元格地址，如 A1、B2',
      },
      data: {
        type: 'array',
        description: '二维数组数据，如 [["姓名","年龄"],["张三",25],["李四",30]]',
        items: {
          type: 'array',
          items: {
            type: 'string',
          },
        },
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'data'],
  },
};

export const writeRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, data, sheet } = args as {
    range: string;
    data: unknown[][];
    sheet?: string;
  };

  // 数据校验
  if (!Array.isArray(data) || data.length === 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '数据不能为空，给我整点实际的！' }],
      error: '数据为空',
    };
  }

  try {
    const success = await wpsClient.setRangeData(sheet || 0, range, data);

    if (success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `数据写入成功！\n起始位置: ${range}\n写入规模: ${data.length}行 x ${data[0]?.length || 0}列`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: '写入数据失败' }],
        error: '写入失败',
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `写入数据出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 数据清洗工具
 * 这是老王的得意之作，一键处理那些乱七八糟的脏数据
 */
export const cleanDataDefinition: ToolDefinition = {
  name: 'wps_excel_clean_data',
  description: `数据清洗工具，支持多种清洗操作的组合。

使用场景：
- "把A列的前后空格去掉" -> 使用 trim 操作
- "把日期格式统一成年-月-日" -> 使用 unify_date 操作
- "删除空行" -> 使用 remove_empty_rows 操作

支持的操作：
- trim: 去除单元格前后空格
- remove_duplicates: 删除重复行
- unify_date: 统一日期格式为 yyyy-mm-dd
- remove_empty_rows: 删除空行`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要清洗的数据范围，如 A1:D100',
      },
      operations: {
        type: 'array',
        description: '要执行的清洗操作列表',
        items: {
          type: 'string',
          enum: ['trim', 'remove_duplicates', 'unify_date', 'remove_empty_rows'],
        },
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'operations'],
  },
};

export const cleanDataHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, operations, sheet } = args as {
    range: string;
    operations: string[];
    sheet?: string;
  };

  // 校验操作列表
  const validOperations = ['trim', 'remove_duplicates', 'unify_date', 'remove_empty_rows'];
  const invalidOps = operations.filter((op) => !validOperations.includes(op));

  if (invalidOps.length > 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [
        {
          type: 'text',
          text: `不支持的操作: ${invalidOps.join(', ')}\n支持的操作: ${validOperations.join(', ')}`,
        },
      ],
      error: '无效的操作类型',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      range: string;
      operations: Array<{
        operation: string;
        success: boolean;
        message: string;
      }>;
      message: string;
    }>(
      'excel.cleanData',
      { range, operations, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `数据清洗失败: ${response.error}` }],
        error: response.error,
      };
    }

    const result = response.data;
    let output = `数据清洗完成！\n范围: ${result.range}\n\n操作结果：\n`;

    result.operations.forEach((op) => {
      const status = op.success ? '成功' : '失败';
      output += `- ${op.operation}: ${status} - ${op.message}\n`;
    });

    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: output }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `数据清洗出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 删除重复行
 * 单独拎出来，因为这个功能用得太多了
 */
export const removeDuplicatesDefinition: ToolDefinition = {
  name: 'wps_excel_remove_duplicates',
  description: '删除指定范围内的重复行。可以指定根据哪些列判断重复。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要处理的数据范围，如 A1:D100',
      },
      columns: {
        type: 'array',
        description: '用于判断重复的列，如 ["A", "B"]。不填则根据所有列判断',
        items: {
          type: 'string',
        },
      },
      has_header: {
        type: 'boolean',
        description: '第一行是否为表头，默认true',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range'],
  },
};

export const removeDuplicatesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, columns, has_header, sheet } = args as {
    range: string;
    columns?: string[];
    has_header?: boolean;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      originalCount: number;
      removedCount: number;
      remainingCount: number;
    }>(
      'excel.removeDuplicates',
      {
        range,
        columns: columns || [],
        hasHeader: has_header !== false, // 默认true
        sheet,
      },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除重复行失败: ${response.error}` }],
        error: response.error,
      };
    }

    const result = response.data;

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `删除重复行完成！
原始行数: ${result.originalCount}
删除行数: ${result.removedCount}
剩余行数: ${result.remainingCount}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除重复行出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有数据处理相关的Tools
 */
export const dataTools: RegisteredTool[] = [
  { definition: readRangeDefinition, handler: readRangeHandler },
  { definition: writeRangeDefinition, handler: writeRangeHandler },
  { definition: cleanDataDefinition, handler: cleanDataHandler },
  { definition: removeDuplicatesDefinition, handler: removeDuplicatesHandler },
];

export default dataTools;
