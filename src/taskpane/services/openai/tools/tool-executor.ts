/* global console */

/**
 * 工具执行器
 * 负责执行 Function Calling 调用的 Excel 操作
 */

import { ExcelOperations, ExcelOperationResult } from "../../../excel";
import { ToolExecutionResult } from "./excel-tools";

/**
 * Excel 工具执行器类
 */
export class ToolExecutor {
  /**
   * 执行 Excel 操作工具
   * @param functionName 函数名称
   * @param args 参数对象
   * @returns 执行结果
   */
  static async execute(
    functionName: string,
    args: Record<string, any>
  ): Promise<ToolExecutionResult> {
    try {
      let result: ExcelOperationResult;

      switch (functionName) {
        case "insertData":
          result = await ExcelOperations.insertData({
            address: args.address,
            data: args.data,
            includeHeaders: args.includeHeaders,
            autoFitColumns: args.autoFitColumns,
            autoFitRows: args.autoFitRows,
            selectAfterInsert: args.selectAfterInsert,
          });
          break;

        case "setCellColor":
          result = await ExcelOperations.setCellColor(args.address, args.backgroundColor);
          break;

        case "setCellColorsBatch":
          result = await ExcelOperations.setCellColorsBatch({
            colorRanges: args.colorRanges,
          });
          break;

        case "setCellFont":
          result = await ExcelOperations.setCellFont(args.address, {
            fontColor: args.fontColor,
            fontSize: args.fontSize,
            fontName: args.fontName,
            bold: args.bold,
            italic: args.italic,
            underline: args.underline,
            backgroundColor: args.backgroundColor,
            horizontalAlignment: args.horizontalAlignment,
            verticalAlignment: args.verticalAlignment,
            numberFormat: args.numberFormat,
          });
          break;

        case "createChart":
          result = await ExcelOperations.createChart({
            type: args.type,
            dataRange: args.dataRange,
            position:
              args.left !== undefined && args.top !== undefined
                ? {
                    left: args.left,
                    top: args.top,
                    width: args.width,
                    height: args.height,
                  }
                : undefined,
            style: {
              title: args.title,
              showLegend: args.showLegend,
              showDataLabels: args.showDataLabels,
              xAxisTitle: args.xAxisTitle,
              yAxisTitle: args.yAxisTitle,
            },
          });
          break;

        case "insertFormula":
          result = await ExcelOperations.insertFormula({
            address: args.address,
            formula: args.formula,
            calculateImmediately: args.calculateImmediately,
          });
          break;

        case "formatRange":
          result = await ExcelOperations.formatRange(args.address, {
            fontColor: args.fontColor,
            fontSize: args.fontSize,
            fontName: args.fontName,
            bold: args.bold,
            italic: args.italic,
            underline: args.underline,
            backgroundColor: args.backgroundColor,
            horizontalAlignment: args.horizontalAlignment,
            verticalAlignment: args.verticalAlignment,
            numberFormat: args.numberFormat,
          });
          break;

        default:
          throw new Error(`未知的 Excel 操作: ${functionName}`);
      }

      return {
        success: result.success,
        data: result.data,
        message: result.message,
        error: result.error,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Excel 操作执行失败";
      console.error(`[ToolExecutor] 执行失败 (${functionName}):`, error);

      return {
        success: false,
        message: "",
        error: errorMessage,
      };
    }
  }

  /**
   * 批量执行多个工具
   * @param calls 工具调用数组
   * @returns 执行结果数组
   */
  static async executeBatch(
    calls: Array<{ functionName: string; args: Record<string, any> }>
  ): Promise<ToolExecutionResult[]> {
    const results: ToolExecutionResult[] = [];

    for (const call of calls) {
      const result = await this.execute(call.functionName, call.args);
      results.push(result);

      // 如果某个操作失败，可以选择继续或停止
      if (!result.success) {
        console.warn(`[ToolExecutor] 批量执行中遇到错误: ${result.error}`);
        // 继续执行剩余操作，但记录错误
      }
    }

    return results;
  }
}
