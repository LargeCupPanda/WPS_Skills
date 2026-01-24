/**
 * Excel公式相关Tools - 老王的公式神器
 * 这几个工具是解决用户"公式不会写"痛点的核心，别tm给我乱改
 *
 * 包含：
 * - wps_excel_set_formula: 设置公式到指定单元格
 * - wps_excel_generate_formula: 根据自然语言生成公式（这个是核心中的核心）
 * - wps_excel_diagnose_formula: 诊断公式错误，分析原因并提供修复建议
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
 * 设置公式到指定单元格
 * 这个工具是公式功能的执行者，生成完公式就靠它写进去
 */
export const setFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_set_formula',
  description: '在指定单元格设置Excel公式。公式必须以=开头，支持所有Excel内置函数。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '目标单元格地址，如 A1、B2:B10',
      },
      formula: {
        type: 'string',
        description: 'Excel公式，必须以=开头，如 =SUM(A1:A10)、=VLOOKUP(A1,B:C,2,0)',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'formula'],
  },
};

export const setFormulaHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, formula, sheet } = args as {
    range: string;
    formula: string;
    sheet?: string;
  };

  // 老王说：公式必须以=开头，不然就是憨批
  if (!formula.startsWith('=')) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '公式必须以=开头，别给我整那些没用的' }],
      error: '公式格式错误：必须以=开头',
    };
  }

  try {
    const response = await wpsClient.executeMethod(
      'excel.setFormula',
      { range, formula, sheet },
      WpsAppType.SPREADSHEET
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `公式设置成功！\n单元格: ${range}\n公式: ${formula}\n计算结果: ${JSON.stringify(response.data)}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `公式设置失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `艹，设置公式出错了: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 根据自然语言生成Excel公式
 * 这是核心中的核心！用户说"帮我写个公式查价格"，就靠这个生成VLOOKUP
 * 老王设计思路：先获取上下文，让AI理解表结构，然后生成公式
 */
export const generateFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_generate_formula',
  description: `根据自然语言描述生成Excel公式。这是解决用户"公式不会写"痛点的核心工具。

使用场景：
- 用户说"帮我写个公式查价格" -> 生成 VLOOKUP/XLOOKUP
- 用户说"如果大于100就显示达标" -> 生成 IF 公式
- 用户说"统计每个部门的人数" -> 生成 COUNTIF
- 用户说"求这列的平均值" -> 生成 AVERAGE

调用此工具会返回当前工作表的上下文信息，包括表头、选中区域等，便于生成准确的公式。`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      description: {
        type: 'string',
        description: '用户对公式需求的自然语言描述，如"查找产品名对应的价格"、"计算A列的总和"',
      },
      target_cell: {
        type: 'string',
        description: '目标单元格地址，如 B2。不填则返回上下文让用户确认',
      },
    },
    required: ['description'],
  },
};

export const generateFormulaHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { description, target_cell } = args as {
    description: string;
    target_cell?: string;
  };

  try {
    // 先获取工作表上下文，这样AI才能理解表结构
    const contextResponse = await wpsClient.executeMethod<{
      workbookName: string;
      currentSheet: string;
      allSheets: string[];
      selectedCell: string;
      headers: Array<{ column: string; value: string }>;
      usedRangeAddress: string;
    }>('excel.getContext', {}, WpsAppType.SPREADSHEET);

    if (!contextResponse.success || !contextResponse.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          {
            type: 'text',
            text: '获取工作表上下文失败，请确保WPS表格已打开并且有活动工作簿',
          },
        ],
        error: '无法获取工作表上下文',
      };
    }

    const context = contextResponse.data;

    // 返回上下文信息，让AI根据这些信息生成公式
    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表上下文信息：
工作簿: ${context.workbookName}
当前工作表: ${context.currentSheet}
所有工作表: ${context.allSheets.join(', ')}
当前选中: ${context.selectedCell}
数据范围: ${context.usedRangeAddress}
表头信息:
${context.headers.map((h) => `  ${h.column}列: ${h.value}`).join('\n')}

用户需求: ${description}
目标单元格: ${target_cell || '待确认'}

请根据上下文和用户描述生成公式，然后调用 wps_excel_set_formula 写入。`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取上下文出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 诊断公式错误
 * 用户看到#REF!、#N/A、#VALUE!这些报错就懵逼，这个工具帮他们分析原因
 */
export const diagnoseFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_diagnose_formula',
  description: `诊断公式错误，分析原因并提供修复建议。

使用场景：
- 用户说"这个公式报错了"
- 用户说"#REF! 是什么意思"
- 用户说"帮我看看公式哪里有问题"

支持诊断的错误类型：
- #REF! - 引用了不存在的单元格
- #N/A - 查找函数未找到匹配值
- #VALUE! - 参数类型错误
- #NAME? - 函数名称错误
- #DIV/0! - 除数为零
- #NUM! - 数值问题
- #NULL! - 交集为空`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      cell: {
        type: 'string',
        description: '包含错误公式的单元格地址，如 A1、B2',
      },
    },
    required: ['cell'],
  },
};

export const diagnoseFormulaHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { cell } = args as { cell: string };

  try {
    const response = await wpsClient.executeMethod<{
      cell: string;
      formula: string;
      currentValue: unknown;
      errorType: string | null;
      diagnosis: string;
      suggestion: string;
      precedents: string[];
    }>('excel.diagnoseFormula', { cell }, WpsAppType.SPREADSHEET);

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `诊断失败: ${response.error}` }],
        error: response.error,
      };
    }

    const diagnosis = response.data;

    // 如果没有错误
    if (!diagnosis.errorType) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `单元格 ${cell} 的公式没有错误！
公式: ${diagnosis.formula}
计算结果: ${JSON.stringify(diagnosis.currentValue)}`,
          },
        ],
      };
    }

    // 有错误，给出详细诊断
    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `公式诊断结果：

单元格: ${diagnosis.cell}
公式: ${diagnosis.formula}
错误类型: ${diagnosis.errorType}

错误原因: ${diagnosis.diagnosis}

修复建议: ${diagnosis.suggestion}

引用的单元格: ${diagnosis.precedents.length > 0 ? diagnosis.precedents.join(', ') : '无引用'}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `诊断公式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有公式相关的Tools
 * 老王说：整整齐齐，方便注册
 */
export const formulaTools: RegisteredTool[] = [
  { definition: setFormulaDefinition, handler: setFormulaHandler },
  { definition: generateFormulaDefinition, handler: generateFormulaHandler },
  { definition: diagnoseFormulaDefinition, handler: diagnoseFormulaHandler },
];

export default formulaTools;
