/**
 * Excel透视表Tools - 老王的数据分析神器
 * 透视表是数据分析的核心功能，别tm给我说你不会用透视表
 *
 * 包含：
 * - wps_excel_create_pivot_table: 创建透视表
 * - wps_excel_update_pivot_table: 更新透视表配置
 *
 * @author 马铁锤 (老王团队数据分析专家)
 * @date 2026-01-24
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
 * 值字段聚合类型
 * 这几个是透视表最常用的聚合方式，别问我为啥没有PRODUCT，用的人太少了
 */
type AggregationType = 'SUM' | 'COUNT' | 'AVERAGE' | 'MAX' | 'MIN';

/**
 * 透视表值字段配置
 */
interface PivotValueField {
  /** 字段名/列名 */
  field: string;
  /** 聚合方式 */
  aggregation: AggregationType;
  /** 自定义显示名称（可选） */
  displayName?: string;
}

/**
 * 透视表配置接口
 * 老王设计的，参数清晰明了，别给我传乱七八糟的
 */
interface PivotTableConfig {
  /** 数据源范围，如"A1:E100"或"Sheet1!A1:E100" */
  sourceRange: string;
  /** 透视表放置位置，如"G1"，这个单元格是透视表左上角 */
  destinationCell: string;
  /** 行字段列名列表，这些会作为透视表的行标签 */
  rowFields: string[];
  /** 列字段列名列表（可选），作为列标签，不填就没有列分组 */
  columnFields?: string[];
  /** 值字段配置列表，这是透视表的核心，汇总啥数据全靠它 */
  valueFields: PivotValueField[];
  /** 筛选字段列名列表（可选），放在透视表上方的筛选器 */
  filterFields?: string[];
  /** 透视表名称（可选），不填自动生成 */
  tableName?: string;
  /** 目标工作表名称（可选），不填就放当前工作表 */
  destinationSheet?: string;
}

/**
 * 创建透视表的Tool定义
 * 这个工具是数据分析的核心入口，用户说"帮我做个数据汇总"就靠它了
 */
export const createPivotTableDefinition: ToolDefinition = {
  name: 'wps_excel_create_pivot_table',
  description: `创建Excel透视表，用于数据汇总和分析。

使用场景：
- "帮我按部门统计销售额" -> 行字段=部门，值字段=销售额(SUM)
- "按月份和产品分析数量" -> 行字段=月份，列字段=产品，值字段=数量(SUM)
- "统计各地区的订单数" -> 行字段=地区，值字段=订单号(COUNT)

注意事项：
- sourceRange必须包含表头行
- rowFields、columnFields、valueFields中的field必须是表头中的列名
- 透视表会从destinationCell开始向右下方扩展`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sourceRange: {
        type: 'string',
        description: '数据源范围，如"A1:E100"。必须包含表头行，数据要连续无空行',
      },
      destinationCell: {
        type: 'string',
        description: '透视表放置位置（左上角单元格），如"G1"、"H5"',
      },
      rowFields: {
        type: 'array',
        description: '行字段列名列表，如["部门", "员工"]。这些字段会作为透视表的行标签',
        items: {
          type: 'string',
        },
      },
      columnFields: {
        type: 'array',
        description: '列字段列名列表（可选），如["月份"]。这些字段会作为透视表的列标签',
        items: {
          type: 'string',
        },
      },
      valueFields: {
        type: 'array',
        description: '值字段配置列表。每项包含field(字段名)和aggregation(聚合方式:SUM/COUNT/AVERAGE/MAX/MIN)',
        items: {
          type: 'object',
          properties: {
            field: {
              type: 'string',
              description: '字段名/列名',
            },
            aggregation: {
              type: 'string',
              description: '聚合方式：SUM(求和)、COUNT(计数)、AVERAGE(平均)、MAX(最大)、MIN(最小)',
              enum: ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'],
            },
          },
          required: ['field', 'aggregation'],
        },
      },
      filterFields: {
        type: 'array',
        description: '筛选字段列名列表（可选），这些字段会出现在透视表上方作为筛选器',
        items: {
          type: 'string',
        },
      },
      tableName: {
        type: 'string',
        description: '透视表名称（可选），不填则自动生成',
      },
      destinationSheet: {
        type: 'string',
        description: '目标工作表名称（可选），透视表将创建在此工作表，不填则使用当前工作表',
      },
    },
    required: ['sourceRange', 'destinationCell', 'rowFields', 'valueFields'],
  },
};

/**
 * 创建透视表的Handler
 */
export const createPivotTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const config = args as unknown as PivotTableConfig;

  // 参数校验 - 老王说必填项一个都不能少
  if (!config.sourceRange || config.sourceRange.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '数据源范围(sourceRange)不能为空，你想让我分析空气吗？' }],
      error: '参数错误：sourceRange为空',
    };
  }

  if (!config.destinationCell || config.destinationCell.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '放置位置(destinationCell)不能为空，透视表总得有个地方放吧' }],
      error: '参数错误：destinationCell为空',
    };
  }

  if (!config.rowFields || config.rowFields.length === 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行字段(rowFields)至少要有一个，不然透视表怎么分组？' }],
      error: '参数错误：rowFields为空',
    };
  }

  if (!config.valueFields || config.valueFields.length === 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '值字段(valueFields)至少要有一个，不然透视表汇总啥？' }],
      error: '参数错误：valueFields为空',
    };
  }

  // 校验聚合类型
  const validAggregations: AggregationType[] = ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'];
  for (const vf of config.valueFields) {
    if (!vf.field || vf.field.trim() === '') {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: '值字段的field不能为空，你要汇总哪个字段告诉我啊' }],
        error: '参数错误：valueFields中field为空',
      };
    }
    if (!vf.aggregation || !validAggregations.includes(vf.aggregation)) {
      return {
        id: uuidv4(),
        success: false,
        content: [{
          type: 'text',
          text: `聚合类型(aggregation)必须是${validAggregations.join('/')}之一，你给的"${vf.aggregation}"是什么鬼？`
        }],
        error: '参数错误：aggregation无效',
      };
    }
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      pivotTableName: string;
      location: string;
      rowCount: number;
      columnCount: number;
    }>(
      'excel.createPivotTable',
      {
        sourceRange: config.sourceRange.trim(),
        destinationCell: config.destinationCell.trim(),
        destinationSheet: config.destinationSheet,
        rowFields: config.rowFields,
        columnFields: config.columnFields || [],
        valueFields: config.valueFields,
        filterFields: config.filterFields || [],
        tableName: config.tableName,
      },
      WpsAppType.SPREADSHEET
    );

    if (response.success && response.data) {
      const result = response.data;
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `透视表创建成功！

透视表名称: ${result.pivotTableName}
放置位置: ${result.location}
数据维度: ${result.rowCount}行 x ${result.columnCount}列

配置信息:
- 数据源: ${config.sourceRange}
- 行字段: ${config.rowFields.join(', ')}
- 列字段: ${config.columnFields?.join(', ') || '无'}
- 值字段: ${config.valueFields.map(v => `${v.field}(${v.aggregation})`).join(', ')}
- 筛选字段: ${config.filterFields?.join(', ') || '无'}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建透视表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `艹，创建透视表出错了: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 更新透视表配置的Tool定义
 * 透视表创建之后经常需要调整，这个工具负责更新配置
 */
export const updatePivotTableDefinition: ToolDefinition = {
  name: 'wps_excel_update_pivot_table',
  description: `更新已有透视表的配置，包括添加/移除字段、修改聚合方式等。

使用场景：
- "给透视表加个筛选器" -> 使用addFilterFields
- "把销售额改成求平均" -> 使用updateValueFields
- "去掉月份这个列字段" -> 使用removeColumnFields
- "刷新透视表数据" -> 使用refresh=true

支持的操作：
- 添加/移除行字段
- 添加/移除列字段
- 添加/移除/更新值字段
- 添加/移除筛选字段
- 刷新数据`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      pivotTableName: {
        type: 'string',
        description: '透视表名称。如果不确定名称，可以先用其他工具查看工作表内容',
      },
      pivotTableCell: {
        type: 'string',
        description: '透视表所在的任意单元格地址（可选）。如果提供了pivotTableName则忽略此参数',
      },
      addRowFields: {
        type: 'array',
        description: '要添加的行字段列名列表',
        items: { type: 'string' },
      },
      removeRowFields: {
        type: 'array',
        description: '要移除的行字段列名列表',
        items: { type: 'string' },
      },
      addColumnFields: {
        type: 'array',
        description: '要添加的列字段列名列表',
        items: { type: 'string' },
      },
      removeColumnFields: {
        type: 'array',
        description: '要移除的列字段列名列表',
        items: { type: 'string' },
      },
      addValueFields: {
        type: 'array',
        description: '要添加的值字段配置列表',
        items: {
          type: 'object',
          properties: {
            field: { type: 'string', description: '字段名' },
            aggregation: {
              type: 'string',
              description: '聚合方式',
              enum: ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'],
            },
          },
          required: ['field', 'aggregation'],
        },
      },
      removeValueFields: {
        type: 'array',
        description: '要移除的值字段名列表',
        items: { type: 'string' },
      },
      updateValueFields: {
        type: 'array',
        description: '要更新的值字段配置（修改聚合方式）',
        items: {
          type: 'object',
          properties: {
            field: { type: 'string', description: '字段名' },
            aggregation: {
              type: 'string',
              description: '新的聚合方式',
              enum: ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'],
            },
          },
          required: ['field', 'aggregation'],
        },
      },
      addFilterFields: {
        type: 'array',
        description: '要添加的筛选字段列名列表',
        items: { type: 'string' },
      },
      removeFilterFields: {
        type: 'array',
        description: '要移除的筛选字段列名列表',
        items: { type: 'string' },
      },
      refresh: {
        type: 'boolean',
        description: '是否刷新透视表数据（当源数据变化时使用）',
      },
    },
    required: [],
  },
};

/**
 * 更新透视表配置的Handler
 */
export const updatePivotTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const {
    pivotTableName,
    pivotTableCell,
    addRowFields,
    removeRowFields,
    addColumnFields,
    removeColumnFields,
    addValueFields,
    removeValueFields,
    updateValueFields,
    addFilterFields,
    removeFilterFields,
    refresh,
  } = args as {
    pivotTableName?: string;
    pivotTableCell?: string;
    addRowFields?: string[];
    removeRowFields?: string[];
    addColumnFields?: string[];
    removeColumnFields?: string[];
    addValueFields?: PivotValueField[];
    removeValueFields?: string[];
    updateValueFields?: PivotValueField[];
    addFilterFields?: string[];
    removeFilterFields?: string[];
    refresh?: boolean;
  };

  // 至少要有一种方式定位透视表
  if (!pivotTableName && !pivotTableCell) {
    return {
      id: uuidv4(),
      success: false,
      content: [{
        type: 'text',
        text: '你要更新哪个透视表？给我透视表名称(pivotTableName)或者透视表所在单元格(pivotTableCell)啊！'
      }],
      error: '参数错误：未指定透视表',
    };
  }

  // 检查是否有实际的操作
  const hasOperations =
    (addRowFields && addRowFields.length > 0) ||
    (removeRowFields && removeRowFields.length > 0) ||
    (addColumnFields && addColumnFields.length > 0) ||
    (removeColumnFields && removeColumnFields.length > 0) ||
    (addValueFields && addValueFields.length > 0) ||
    (removeValueFields && removeValueFields.length > 0) ||
    (updateValueFields && updateValueFields.length > 0) ||
    (addFilterFields && addFilterFields.length > 0) ||
    (removeFilterFields && removeFilterFields.length > 0) ||
    refresh === true;

  if (!hasOperations) {
    return {
      id: uuidv4(),
      success: false,
      content: [{
        type: 'text',
        text: '你让我更新透视表但啥操作都没给，逗我玩呢？至少指定一个操作参数！'
      }],
      error: '参数错误：无操作',
    };
  }

  // 校验值字段的聚合类型
  const validAggregations: AggregationType[] = ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'];
  const validateValueFields = (fields: PivotValueField[] | undefined, fieldName: string): ToolCallResult | null => {
    if (!fields) return null;
    for (const vf of fields) {
      if (!vf.field || vf.field.trim() === '') {
        return {
          id: uuidv4(),
          success: false,
          content: [{ type: 'text', text: `${fieldName}中的field不能为空` }],
          error: `参数错误：${fieldName}中field为空`,
        };
      }
      if (!vf.aggregation || !validAggregations.includes(vf.aggregation)) {
        return {
          id: uuidv4(),
          success: false,
          content: [{
            type: 'text',
            text: `${fieldName}中的聚合类型必须是${validAggregations.join('/')}之一`
          }],
          error: `参数错误：${fieldName}中aggregation无效`,
        };
      }
    }
    return null;
  };

  const addValueFieldsError = validateValueFields(addValueFields, 'addValueFields');
  if (addValueFieldsError) return addValueFieldsError;

  const updateValueFieldsError = validateValueFields(updateValueFields, 'updateValueFields');
  if (updateValueFieldsError) return updateValueFieldsError;

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      pivotTableName: string;
      operations: Array<{
        operation: string;
        success: boolean;
        message: string;
      }>;
    }>(
      'excel.updatePivotTable',
      {
        pivotTableName: pivotTableName?.trim(),
        pivotTableCell: pivotTableCell?.trim(),
        addRowFields: addRowFields || [],
        removeRowFields: removeRowFields || [],
        addColumnFields: addColumnFields || [],
        removeColumnFields: removeColumnFields || [],
        addValueFields: addValueFields || [],
        removeValueFields: removeValueFields || [],
        updateValueFields: updateValueFields || [],
        addFilterFields: addFilterFields || [],
        removeFilterFields: removeFilterFields || [],
        refresh: refresh || false,
      },
      WpsAppType.SPREADSHEET
    );

    if (response.success && response.data) {
      const result = response.data;
      let output = `透视表"${result.pivotTableName}"更新完成！\n\n操作结果：\n`;

      for (const op of result.operations) {
        const status = op.success ? '成功' : '失败';
        output += `- ${op.operation}: ${status} - ${op.message}\n`;
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `更新透视表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `艹，更新透视表出错了: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有透视表相关的Tools
 * 整整齐齐，方便注册
 */
export const pivotTools: RegisteredTool[] = [
  { definition: createPivotTableDefinition, handler: createPivotTableHandler },
  { definition: updatePivotTableDefinition, handler: updatePivotTableHandler },
];

export default pivotTools;
