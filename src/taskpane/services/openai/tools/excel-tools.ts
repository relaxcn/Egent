/* global console */

/**
 * Excel 工具定义和验证器
 * 从 excel 模块导入并重新导出，提供更清晰的访问路径
 */

import OpenAI from "openai";

// 从 Excel 模块导入 Function 定义
export {
  INSERT_DATA_FUNCTION,
  SET_CELL_COLOR_FUNCTION,
  SET_CELL_COLORS_BATCH_FUNCTION,
  SET_CELL_FONT_FUNCTION,
  CREATE_CHART_FUNCTION,
  INSERT_FORMULA_FUNCTION,
  FORMAT_RANGE_FUNCTION,
  EXCEL_FUNCTIONS,
  EXCEL_FUNCTION_NAMES,
  getExcelFunction,
  validateFunctionCall,
} from "../../../excel/function-definitions";

/**
 * 获取所有可用的 Excel 工具定义
 * @returns OpenAI Function Calling 工具数组
 */
export function getAllExcelTools(): OpenAI.ChatCompletionTool[] {
  // 动态导入确保获取最新定义
  const { EXCEL_FUNCTIONS } = require("../../../excel/function-definitions");
  return EXCEL_FUNCTIONS;
}

/**
 * 根据工具名称获取单个工具定义
 * @param toolName 工具名称
 * @returns 工具定义或 undefined
 */
export function getExcelToolByName(toolName: string): OpenAI.ChatCompletionTool | undefined {
  const { getExcelFunction } = require("../../../excel/function-definitions");
  return getExcelFunction(toolName);
}

/**
 * 验证工具调用参数
 * @param toolName 工具名称
 * @param args 参数对象
 * @returns 验证结果
 */
export function validateToolCall(
  toolName: string,
  args: Record<string, any>
): { valid: boolean; error?: string } {
  const { validateFunctionCall } = require("../../../excel/function-definitions");
  return validateFunctionCall(toolName, args);
}

/**
 * 工具执行结果接口
 */
export interface ToolExecutionResult {
  success: boolean;
  data?: any;
  message: string;
  error?: string;
}
