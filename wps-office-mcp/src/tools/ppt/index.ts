/**
 * PPT Tools入口 - 老王的演示工具箱
 * 目前主要是幻灯片操作相关的Tools
 *
 * 这里导出的是所有PPT相关的Tools，注册的时候一把梭哈
 */

import { RegisteredTool } from '../../types/tools';
import { slideTools } from './slide';

/**
 * 所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font
 */
export const pptTools: RegisteredTool[] = [
  ...slideTools,
];

// 分别导出，方便按需使用
export { slideTools } from './slide';

// 导出单独的定义和处理器，方便测试
export {
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './slide';

export default pptTools;
