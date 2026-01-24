/**
 * Excel Tools入口 - 老王的表格工具箱
 * 把公式、数据处理、透视表、图表的Tools都整合到这里
 *
 * 这里导出的是所有Excel相关的Tools，注册的时候一把梭哈
 */

import { RegisteredTool } from '../../types/tools';
import { formulaTools } from './formula';
import { dataTools } from './data';
import { pivotTools } from './pivot';
import { chartTools } from './chart';

/**
 * 所有Excel相关的Tools
 * 包含：
 * - 公式Tools: set_formula, generate_formula, diagnose_formula
 * - 数据Tools: read_range, write_range, clean_data, remove_duplicates
 * - 透视表Tools: create_pivot_table, update_pivot_table
 * - 图表Tools: create_chart, update_chart (刘大炮出品)
 */
export const excelTools: RegisteredTool[] = [
  ...formulaTools,
  ...dataTools,
  ...pivotTools,
  ...chartTools,
];

// 分别导出，方便按需使用
export { formulaTools } from './formula';
export { dataTools } from './data';
export { pivotTools } from './pivot';
export { chartTools } from './chart';

// 导出单独的定义和处理器，方便测试
export {
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
} from './formula';

export {
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
} from './data';

export {
  createPivotTableDefinition,
  createPivotTableHandler,
  updatePivotTableDefinition,
  updatePivotTableHandler,
} from './pivot';

export {
  createChartDefinition,
  createChartHandler,
  updateChartDefinition,
  updateChartHandler,
} from './chart';

export default excelTools;
