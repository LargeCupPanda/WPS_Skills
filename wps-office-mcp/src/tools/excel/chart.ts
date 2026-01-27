/**
 * Excel图表相关Tools - 刘大炮的可视化神器
 * 老王让我来搞这个图表模块，别tm给我乱传参数
 *
 * 包含：
 * - wps_excel_create_chart: 创建图表（柱状图、折线图、饼图、散点图等）
 * - wps_excel_update_chart: 更新图表属性（标题、颜色、图例等）
 *
 * @author 刘大炮
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
 * 支持的图表类型枚举
 * 这些是WPS Excel支持的主流图表类型，别给我瞎编类型
 */
export enum ChartType {
  /** 簇状柱形图 - 最常用的，比较数值大小 */
  COLUMN_CLUSTERED = 'column_clustered',
  /** 堆积柱形图 - 看整体和部分关系 */
  COLUMN_STACKED = 'column_stacked',
  /** 簇状条形图 - 横着的柱形图 */
  BAR_CLUSTERED = 'bar_clustered',
  /** 折线图 - 看趋势变化 */
  LINE = 'line',
  /** 带数据标记的折线图 */
  LINE_MARKERS = 'line_markers',
  /** 饼图 - 看占比 */
  PIE = 'pie',
  /** 环形图 - 空心饼图 */
  DOUGHNUT = 'doughnut',
  /** 散点图 - 看相关性 */
  SCATTER = 'scatter',
  /** 面积图 - 强调数量变化的程度 */
  AREA = 'area',
  /** 雷达图 - 多维度对比 */
  RADAR = 'radar',
}

/**
 * 图表类型到WPS常量的映射
 * WPS SDK用的是Excel的XlChartType常量
 */
const CHART_TYPE_MAP: Record<ChartType, number> = {
  [ChartType.COLUMN_CLUSTERED]: 51,    // xlColumnClustered
  [ChartType.COLUMN_STACKED]: 52,      // xlColumnStacked
  [ChartType.BAR_CLUSTERED]: 57,       // xlBarClustered
  [ChartType.LINE]: 4,                  // xlLine
  [ChartType.LINE_MARKERS]: 65,        // xlLineMarkers
  [ChartType.PIE]: 5,                   // xlPie
  [ChartType.DOUGHNUT]: -4120,         // xlDoughnut
  [ChartType.SCATTER]: -4169,          // xlXYScatter
  [ChartType.AREA]: 1,                  // xlArea
  [ChartType.RADAR]: -4151,            // xlRadar
};

/**
 * 创建图表工具
 * 这是核心功能，根据数据范围创建各种类型的图表
 */
export const createChartDefinition: ToolDefinition = {
  name: 'wps_excel_create_chart',
  description: `在Excel中创建图表。支持柱状图、折线图、饼图、散点图等多种类型。

使用场景：
- "帮我用A1:B10的数据画个柱状图" -> 创建 column_clustered
- "把这些数据做成折线图看趋势" -> 创建 line
- "显示各部门占比" -> 创建 pie 饼图
- "分析两个变量的相关性" -> 创建 scatter 散点图

支持的图表类型：
- column_clustered: 簇状柱形图（默认，最常用）
- column_stacked: 堆积柱形图
- bar_clustered: 簇状条形图
- line: 折线图
- line_markers: 带标记的折线图
- pie: 饼图
- doughnut: 环形图
- scatter: 散点图
- area: 面积图
- radar: 雷达图`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      data_range: {
        type: 'string',
        description: '数据范围，如 A1:C10。这是图表数据的来源，别给我传空的',
      },
      chart_type: {
        type: 'string',
        description: '图表类型，默认 column_clustered（簇状柱形图）',
        enum: Object.values(ChartType),
      },
      title: {
        type: 'string',
        description: '图表标题，如 "销售趋势图"',
      },
      position: {
        type: 'object',
        description: '图表位置，不填则自动放在数据右侧',
        properties: {
          left: {
            type: 'number',
            description: '左边距（像素）',
          },
          top: {
            type: 'number',
            description: '上边距（像素）',
          },
          width: {
            type: 'number',
            description: '图表宽度（像素），默认480',
          },
          height: {
            type: 'number',
            description: '图表高度（像素），默认300',
          },
        },
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
      has_header: {
        type: 'boolean',
        description: '数据第一行是否为表头，默认true',
      },
      show_legend: {
        type: 'boolean',
        description: '是否显示图例，默认true',
      },
      show_data_labels: {
        type: 'boolean',
        description: '是否显示数据标签，默认false',
      },
    },
    required: ['data_range'],
  },
};

export const createChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const {
    data_range,
    chart_type = ChartType.COLUMN_CLUSTERED,
    title,
    position,
    sheet,
    has_header = true,
    show_legend = true,
    show_data_labels = false,
  } = args as {
    data_range: string;
    chart_type?: ChartType;
    title?: string;
    position?: {
      left?: number;
      top?: number;
      width?: number;
      height?: number;
    };
    sheet?: string;
    has_header?: boolean;
    show_legend?: boolean;
    show_data_labels?: boolean;
  };

  // 校验数据范围格式，老王的规矩：参数必须严格校验
  if (!data_range || !/^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i.test(data_range)) {
    return {
      id: uuidv4(),
      success: false,
      content: [
        {
          type: 'text',
          text: `艹，数据范围格式不对！应该是类似 A1:C10 这样的格式，你给我传的是: ${data_range}`,
        },
      ],
      error: '数据范围格式无效',
    };
  }

  // 校验图表类型
  const validChartTypes = Object.values(ChartType);
  if (!validChartTypes.includes(chart_type as ChartType)) {
    return {
      id: uuidv4(),
      success: false,
      content: [
        {
          type: 'text',
          text: `不支持的图表类型: ${chart_type}\n支持的类型: ${validChartTypes.join(', ')}`,
        },
      ],
      error: '无效的图表类型',
    };
  }

  try {
    // 获取WPS图表类型常量
    const wpsChartType = CHART_TYPE_MAP[chart_type as ChartType];

    const response = await wpsClient.executeMethod<{
      chartName: string;
      chartIndex: number;
      dataRange: string;
      chartType: string;
      position: { left: number; top: number; width: number; height: number };
    }>(
      'createChart',
      {
        dataRange: data_range,
        chartType: wpsChartType,
        chartTypeName: chart_type,
        title: title || '',
        position: {
          left: position?.left,
          top: position?.top,
          width: position?.width || 480,
          height: position?.height || 300,
        },
        sheet,
        hasHeader: has_header,
        showLegend: show_legend,
        showDataLabels: show_data_labels,
      },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          { type: 'text', text: `创建图表失败: ${response.error || '未知错误'}` },
        ],
        error: response.error,
      };
    }

    const result = response.data;
    const pos = result.position || { left: 0, top: 0, width: 400, height: 300 };

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `图表创建成功！
图表名称: ${result.chartName || 'Chart'}
图表索引: ${result.chartIndex || 1}
数据范围: ${result.dataRange || data_range}
图表类型: ${chart_type}
位置: 左${pos.left}px, 上${pos.top}px
尺寸: ${pos.width}x${pos.height}px

提示: 可以使用 wps_excel_update_chart 工具修改图表的标题、颜色等属性`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建图表出错，艹: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 更新图表属性工具
 * 改标题、改颜色、改图例什么的都靠这个
 */
export const updateChartDefinition: ToolDefinition = {
  name: 'wps_excel_update_chart',
  description: `更新Excel图表的属性，包括标题、颜色、图例、数据标签等。

使用场景：
- "把图表标题改成销售报表" -> 更新 title
- "隐藏图例" -> 设置 show_legend: false
- "显示数据标签" -> 设置 show_data_labels: true
- "改变图表类型为折线图" -> 设置 chart_type: line

注意：需要先通过 wps_excel_create_chart 创建图表，或者指定已存在图表的名称/索引`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      chart_index: {
        type: 'number',
        description: '图表索引（从1开始），与chart_name二选一',
      },
      chart_name: {
        type: 'string',
        description: '图表名称，与chart_index二选一',
      },
      title: {
        type: 'string',
        description: '新的图表标题',
      },
      chart_type: {
        type: 'string',
        description: '更改图表类型',
        enum: Object.values(ChartType),
      },
      show_legend: {
        type: 'boolean',
        description: '是否显示图例',
      },
      legend_position: {
        type: 'string',
        description: '图例位置：bottom（下）、top（上）、left（左）、right（右）',
        enum: ['bottom', 'top', 'left', 'right'],
      },
      show_data_labels: {
        type: 'boolean',
        description: '是否显示数据标签',
      },
      data_range: {
        type: 'string',
        description: '更改数据源范围',
      },
      colors: {
        type: 'array',
        description: '系列颜色数组，如 ["#FF0000", "#00FF00", "#0000FF"]',
        items: {
          type: 'string',
        },
      },
      sheet: {
        type: 'string',
        description: '图表所在的工作表名称',
      },
    },
    required: [],
  },
};

export const updateChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const {
    chart_index,
    chart_name,
    title,
    chart_type,
    show_legend,
    legend_position,
    show_data_labels,
    data_range,
    colors,
    sheet,
  } = args as {
    chart_index?: number;
    chart_name?: string;
    title?: string;
    chart_type?: ChartType;
    show_legend?: boolean;
    legend_position?: 'bottom' | 'top' | 'left' | 'right';
    show_data_labels?: boolean;
    data_range?: string;
    colors?: string[];
    sheet?: string;
  };

  // 必须指定图表索引或名称
  if (chart_index === undefined && !chart_name) {
    return {
      id: uuidv4(),
      success: false,
      content: [
        {
          type: 'text',
          text: '憨批，你得告诉我要更新哪个图表！chart_index或chart_name至少填一个',
        },
      ],
      error: '未指定目标图表',
    };
  }

  // 校验图表类型（如果指定了）
  if (chart_type) {
    const validChartTypes = Object.values(ChartType);
    if (!validChartTypes.includes(chart_type)) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          {
            type: 'text',
            text: `不支持的图表类型: ${chart_type}\n支持的类型: ${validChartTypes.join(', ')}`,
          },
        ],
        error: '无效的图表类型',
      };
    }
  }

  // 校验颜色格式
  if (colors && colors.length > 0) {
    const colorRegex = /^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/;
    const invalidColors = colors.filter((c) => !colorRegex.test(c));
    if (invalidColors.length > 0) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          {
            type: 'text',
            text: `颜色格式不对！应该是十六进制格式如 #FF0000，无效的颜色: ${invalidColors.join(', ')}`,
          },
        ],
        error: '颜色格式无效',
      };
    }
  }

  try {
    // 构建更新参数
    const updateParams: Record<string, unknown> = {
      chartIndex: chart_index,
      chartName: chart_name,
      sheet,
    };

    // 只添加需要更新的属性
    if (title !== undefined) updateParams.title = title;
    if (chart_type !== undefined) {
      updateParams.chartType = CHART_TYPE_MAP[chart_type];
      updateParams.chartTypeName = chart_type;
    }
    if (show_legend !== undefined) updateParams.showLegend = show_legend;
    if (legend_position !== undefined) updateParams.legendPosition = legend_position;
    if (show_data_labels !== undefined) updateParams.showDataLabels = show_data_labels;
    if (data_range !== undefined) updateParams.dataRange = data_range;
    if (colors !== undefined) updateParams.colors = colors;

    const response = await wpsClient.executeMethod<{
      chartName: string;
      updatedProperties: string[];
    }>(
      'updateChart',
      updateParams,
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          { type: 'text', text: `更新图表失败: ${response.error || '未知错误'}` },
        ],
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
          text: `图表更新成功！
图表名称: ${result.chartName}
更新的属性: ${result.updatedProperties.join(', ')}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `更新图表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有图表相关的Tools
 * 刘大炮出品，质量保证
 */
export const chartTools: RegisteredTool[] = [
  { definition: createChartDefinition, handler: createChartHandler },
  { definition: updateChartDefinition, handler: updateChartHandler },
];

export default chartTools;
