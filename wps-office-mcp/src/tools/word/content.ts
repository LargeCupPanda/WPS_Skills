/**
 * Word内容操作Tools - 老王的文档编辑神器
 * 处理文档内容的插入、查找替换等操作
 *
 * 包含：
 * - wps_word_insert_text: 插入文本到文档
 * - wps_word_find_replace: 查找替换功能
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
 * 插入文本到文档
 * 可以在光标位置、文档开头或结尾插入文本
 */
export const insertTextDefinition: ToolDefinition = {
  name: 'wps_word_insert_text',
  description: `在Word文档中插入文本。

使用场景：
- "在文档开头加个标题"
- "在光标位置插入这段话"
- "在文档末尾添加备注"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要插入的文本内容',
      },
      position: {
        type: 'string',
        description: '插入位置，可选值: "cursor"(光标位置), "start"(文档开头), "end"(文档结尾)。默认cursor',
        enum: ['cursor', 'start', 'end'],
      },
      style: {
        type: 'string',
        description: '插入后应用的样式，如 "标题 1"、"正文"',
      },
      new_paragraph: {
        type: 'boolean',
        description: '插入后是否新起一段，默认false',
      },
    },
    required: ['text'],
  },
};

export const insertTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, position, style, new_paragraph } = args as {
    text: string;
    position?: string;
    style?: string;
    new_paragraph?: boolean;
  };

  if (!text || text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '要插入的文本不能为空！' }],
      error: '文本内容为空',
    };
  }

  try {
    // 处理换行符，如果需要新段落
    const finalText = new_paragraph ? text + '\n' : text;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      position: string;
      textLength: number;
    }>(
      'insertText',
      {
        text: finalText,
        position: position || 'cursor',
        style,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const positionText =
        position === 'start' ? '文档开头' :
        position === 'end' ? '文档结尾' : '光标位置';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本插入成功！\n位置: ${positionText}\n字符数: ${response.data.textLength}${style ? `\n应用样式: ${style}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 查找替换功能
 * 老王最爱用的功能之一，批量替换文本太爽了
 */
export const findReplaceDefinition: ToolDefinition = {
  name: 'wps_word_find_replace',
  description: `在Word文档中查找并替换文本。

使用场景：
- "把所有的'公司'替换成'集团'"
- "把文档里的错别字改过来"
- "批量替换某个词"

支持选项：
- 区分大小写
- 全字匹配
- 全部替换或仅替换一处`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      find_text: {
        type: 'string',
        description: '要查找的文本',
      },
      replace_text: {
        type: 'string',
        description: '替换为的文本，如果只是查找不替换，可以不填',
      },
      replace_all: {
        type: 'boolean',
        description: '是否全部替换，默认true',
      },
      match_case: {
        type: 'boolean',
        description: '是否区分大小写，默认false',
      },
      match_whole_word: {
        type: 'boolean',
        description: '是否全字匹配，默认false',
      },
    },
    required: ['find_text'],
  },
};

export const findReplaceHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { find_text, replace_text, replace_all, match_case, match_whole_word } = args as {
    find_text: string;
    replace_text?: string;
    replace_all?: boolean;
    match_case?: boolean;
    match_whole_word?: boolean;
  };

  if (!find_text || find_text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '查找文本不能为空！' }],
      error: '查找文本为空',
    };
  }

  try {
    // 如果没有替换文本，就只是查找
    const isReplaceMode = replace_text !== undefined && replace_text !== null;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      findText: string;
      replaceText: string;
      count: number;
    }>(
      'findReplace',
      {
        findText: find_text,
        replaceText: isReplaceMode ? replace_text : '',
        replaceAll: replace_all !== false, // 默认true
        matchCase: match_case || false,
        matchWholeWord: match_whole_word || false,
        replaceMode: isReplaceMode,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const result = response.data;

      if (isReplaceMode) {
        if (result.count === 0) {
          return {
            id: uuidv4(),
            success: true,
            content: [
              {
                type: 'text',
                text: `未找到 "${find_text}"，没有进行替换`,
              },
            ],
          };
        }
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `替换完成！\n查找: "${find_text}"\n替换为: "${replace_text}"\n替换了 ${result.count} 处`,
            },
          ],
        };
      } else {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `查找完成！\n"${find_text}" 在文档中出现了 ${result.count} 次`,
            },
          ],
        };
      }
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `查找替换失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `查找替换出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有内容操作相关的Tools
 */
export const contentTools: RegisteredTool[] = [
  { definition: insertTextDefinition, handler: insertTextHandler },
  { definition: findReplaceDefinition, handler: findReplaceHandler },
];

export default contentTools;
