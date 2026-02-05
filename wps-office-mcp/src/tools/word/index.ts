/**
 * Input: Word 工具定义
 * Output: Word 工具注册数组
 * Pos: Word Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word Tools入口 - 老王的文档工具箱
 * 把格式化和内容操作的Tools都整合到这里
 *
 * 这里导出的是所有Word相关的Tools，注册的时候一把梭哈
 */

import { RegisteredTool } from '../../types/tools';
import { formatTools } from './format';
import { contentTools } from './content';

/**
 * 所有Word相关的Tools
 * 包含：
 * - 格式化Tools: apply_style, set_font, generate_toc
 * - 内容Tools: insert_text, find_replace
 */
export const wordTools: RegisteredTool[] = [
  ...formatTools,
  ...contentTools,
];

// 分别导出，方便按需使用
export { formatTools } from './format';
export { contentTools } from './content';

// 导出单独的定义和处理器，方便测试
export {
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
} from './format';

export {
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
} from './content';

export default wordTools;
