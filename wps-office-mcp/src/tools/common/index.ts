/**
 * Common Tools入口 - 二狗的通用工具箱
 * 把跨应用的通用Tools都整合到这里
 *
 * 这里导出的是所有通用的Tools，Word/Excel/PPT都能用
 * 主要是文档转换、格式互转这些跨应用的功能
 */

import { RegisteredTool } from '../../types/tools';
import { convertTools } from './convert';

/**
 * 所有通用的Tools
 * 包含：
 * - 转换Tools: convert_to_pdf, convert_format
 *
 * 后续可以继续添加更多通用功能，比如：
 * - 文档比较
 * - 批量处理
 * - 云端同步
 * 等等，二狗先把转换搞定再说
 */
export const commonTools: RegisteredTool[] = [
  ...convertTools,
];

// 分别导出，方便按需使用
export { convertTools } from './convert';

// 导出单独的定义和处理器，方便测试
export {
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './convert';

export default commonTools;
