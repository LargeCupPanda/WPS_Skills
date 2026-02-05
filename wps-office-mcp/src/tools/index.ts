/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - 老王的工具大集合
 * 把Excel、Word、PPT、Common的所有Tools都整合到这里
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 */

import { RegisteredTool } from '../types/tools';
import { excelTools } from './excel';
import { wordTools } from './word';
import { pptTools } from './ppt';
import { commonTools } from './common';

/**
 * 所有MCP Tools集合
 *
 * 包含的Tools：
 *
 * Excel (7个):
 * - wps_excel_set_formula: 设置公式
 * - wps_excel_generate_formula: 生成公式
 * - wps_excel_diagnose_formula: 诊断公式
 * - wps_excel_read_range: 读取数据
 * - wps_excel_write_range: 写入数据
 * - wps_excel_clean_data: 数据清洗
 * - wps_excel_remove_duplicates: 删除重复
 *
 * Word (5个):
 * - wps_word_apply_style: 应用样式
 * - wps_word_set_font: 设置字体
 * - wps_word_generate_toc: 生成目录
 * - wps_word_insert_text: 插入文本
 * - wps_word_find_replace: 查找替换
 *
 * PPT (3个):
 * - wps_ppt_add_slide: 添加幻灯片
 * - wps_ppt_beautify: 美化幻灯片
 * - wps_ppt_unify_font: 统一字体
 *
 * Common (2个):
 * - wps_convert_to_pdf: 转换为PDF
 * - wps_convert_format: 格式互转
 */
export const allTools: RegisteredTool[] = [
  ...excelTools,
  ...wordTools,
  ...pptTools,
  ...commonTools,
];

// 按应用类型分别导出
export { excelTools } from './excel';
export { wordTools } from './word';
export { pptTools } from './ppt';
export { commonTools } from './common';

// Excel相关导出
export {
  formulaTools,
  dataTools,
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
} from './excel';

// Word相关导出
export {
  formatTools,
  contentTools,
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
} from './word';

// PPT相关导出
export {
  slideTools,
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './ppt';

// Common相关导出
export {
  convertTools,
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './common';

/**
 * 获取所有Tool的数量
 */
export const getToolCount = (): number => allTools.length;

/**
 * 获取按应用分类的Tool数量
 */
export const getToolCountByApp = (): { excel: number; word: number; ppt: number; common: number } => ({
  excel: excelTools.length,
  word: wordTools.length,
  ppt: pptTools.length,
  common: commonTools.length,
});

export default allTools;
