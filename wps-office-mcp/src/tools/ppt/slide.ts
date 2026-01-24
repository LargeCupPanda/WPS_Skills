/**
 * PPT幻灯片Tools - 老王的演示神器
 * 处理幻灯片的添加、美化、字体统一等操作
 *
 * 包含：
 * - wps_ppt_add_slide: 添加新幻灯片
 * - wps_ppt_beautify: 美化幻灯片（核心功能）
 * - wps_ppt_unify_font: 统一字体
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
 * 添加新幻灯片
 * 可以指定布局、位置、标题和内容
 */
export const addSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_add_slide',
  description: `添加新幻灯片到演示文稿。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容（最常用）
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局

使用场景：
- "新建一页PPT"
- "添加一个标题页"
- "在第3页后面插入一页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      layout: {
        type: 'string',
        description: '幻灯片布局类型',
        enum: ['title', 'title_content', 'blank', 'two_column', 'comparison'],
      },
      position: {
        type: 'number',
        description: '插入位置（页码），不填则在末尾添加',
      },
      title: {
        type: 'string',
        description: '幻灯片标题',
      },
      content: {
        type: 'string',
        description: '幻灯片内容（针对有内容区域的布局）',
      },
    },
    required: [],
  },
};

export const addSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { layout, position, title, content } = args as {
    layout?: string;
    position?: number;
    title?: string;
    content?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      layout: string;
    }>(
      'slide.add',
      {
        layout: layout || 'title_content',
        position,
        title,
        content,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const layoutName = {
        title: '标题页',
        title_content: '标题+内容',
        blank: '空白页',
        two_column: '两栏内容',
        comparison: '对比布局',
      }[response.data.layout] || response.data.layout;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片添加成功！\n位置: 第 ${response.data.slideIndex} 页\n布局: ${layoutName}${title ? `\n标题: ${title}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 美化幻灯片
 * 这是老王最得意的功能，一键让PPT变得专业
 */
export const beautifyDefinition: ToolDefinition = {
  name: 'wps_ppt_beautify',
  description: `一键美化幻灯片，优化排版、配色、字体和间距。

支持的配色方案：
- business: 商务风（深蓝+灰色）
- tech: 科技风（蓝色+绿色）
- creative: 创意风（珊瑚红+金色）
- minimal: 简约风（黑白灰）

美化包含的操作：
- 统一字体
- 应用配色方案
- 对齐元素
- 优化间距

使用场景：
- "美化这页PPT"
- "用商务风格优化一下"
- "把PPT弄好看点"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slide_index: {
        type: 'number',
        description: '要美化的幻灯片页码，不填则美化当前页',
      },
      color_scheme: {
        type: 'string',
        description: '配色方案',
        enum: ['business', 'tech', 'creative', 'minimal'],
      },
      font: {
        type: 'string',
        description: '统一使用的字体，如 "微软雅黑"、"思源黑体"',
      },
      beautify_all: {
        type: 'boolean',
        description: '是否美化所有幻灯片，默认false只美化指定页',
      },
    },
    required: [],
  },
};

export const beautifyHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slide_index, color_scheme, font, beautify_all } = args as {
    slide_index?: number;
    color_scheme?: string;
    font?: string;
    beautify_all?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number | string;
      operations: Array<{
        operation: string;
        count: number;
        details?: string;
      }>;
    }>(
      'slide.beautify',
      {
        slideIndex: beautify_all ? 'all' : slide_index,
        style: {
          colorScheme: color_scheme || 'business',
          font: font || '微软雅黑',
        },
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const result = response.data;
      const schemeName = {
        business: '商务风',
        tech: '科技风',
        creative: '创意风',
        minimal: '简约风',
      }[color_scheme || 'business'] || color_scheme;

      let output = `幻灯片美化完成！\n`;
      output += `范围: ${beautify_all ? '所有幻灯片' : `第 ${result.slideIndex} 页`}\n`;
      output += `配色方案: ${schemeName}\n`;
      output += `字体: ${font || '微软雅黑'}\n\n`;
      output += `优化详情：\n`;

      result.operations.forEach((op) => {
        const opName = {
          unify_font: '统一字体',
          apply_color_scheme: '应用配色',
          align: '对齐元素',
          optimize_spacing: '优化间距',
        }[op.operation] || op.operation;
        output += `- ${opName}: 处理了 ${op.count} 个元素\n`;
      });

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `美化幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `美化幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 统一字体
 * 整个演示文稿的字体统一是PPT美观的基础
 */
export const unifyFontDefinition: ToolDefinition = {
  name: 'wps_ppt_unify_font',
  description: `统一演示文稿中所有幻灯片的字体。

使用场景：
- "把所有页面的字体都改成微软雅黑"
- "统一字体"
- "换个字体"

常用字体推荐：
- 微软雅黑：现代简洁，适合商务
- 思源黑体：开源免费，适合各种场合
- 黑体：传统正式
- 宋体：适合正式文档`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      font_name: {
        type: 'string',
        description: '要统一使用的字体名称，如 "微软雅黑"、"思源黑体"',
      },
      slide_index: {
        type: 'number',
        description: '只处理指定页，不填则处理所有页',
      },
      include_title: {
        type: 'boolean',
        description: '是否包含标题，默认true',
      },
      include_body: {
        type: 'boolean',
        description: '是否包含正文，默认true',
      },
    },
    required: ['font_name'],
  },
};

export const unifyFontHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { font_name, slide_index, include_title, include_body } = args as {
    font_name: string;
    slide_index?: number;
    include_title?: boolean;
    include_body?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      font: string;
      count: number;
      slideCount: number;
    }>(
      'slide.unifyFont',
      {
        fontName: font_name,
        slideIndex: slide_index,
        includeTitle: include_title !== false,
        includeBody: include_body !== false,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const result = response.data;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `字体统一完成！\n字体: ${result.font}\n处理幻灯片: ${result.slideCount} 页\n修改文本框: ${result.count} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `统一字体失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `统一字体出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有幻灯片相关的Tools
 */
export const slideTools: RegisteredTool[] = [
  { definition: addSlideDefinition, handler: addSlideHandler },
  { definition: beautifyDefinition, handler: beautifyHandler },
  { definition: unifyFontDefinition, handler: unifyFontHandler },
];

export default slideTools;
