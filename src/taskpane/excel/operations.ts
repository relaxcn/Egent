/* global Excel, console */

import {
  ExcelOperationResult,
  ExcelRange,
  ExcelInsertDataOptions,
  ExcelInsertDataResult,
  ExcelCellFormat,
  ExcelChartOptions,
  ExcelChartResult,
  ExcelFormulaOptions,
  ExcelFormulaResult,
  ExcelWorksheetInfo,
  ExcelSelectionInfo,
  ExcelError,
  ExcelErrorType,
  ExcelOperationOptions,
  ExcelBatchColorOptions,
  ExcelBatchColorResult,
  ExcelColorRange
} from "./types";
import { ExcelValidator } from "./validation";

/**
 * Excel操作核心类
 * 提供所有Excel操作的统一接口，包含完整的错误处理和类型安全
 */
export class ExcelOperations {
  private static readonly DEFAULT_TIMEOUT = 30000; // 30秒
  private static readonly DEFAULT_RETRY_COUNT = 3;
  private static readonly DEFAULT_RETRY_DELAY = 1000; // 1秒

  // =============================================================================
  // 数据操作方法
  // =============================================================================

  /**
   * 向Excel指定范围插入数据
   * @param options 插入数据的配置选项
   * @returns 操作结果，包含插入的行列数等信息
   */
  static async insertData(
    options: ExcelInsertDataOptions
  ): Promise<ExcelOperationResult<ExcelInsertDataResult>> {
    const startTime = Date.now();

    try {
      // 验证参数
      const validation = ExcelValidator.validateAddress(options.address);
      if (!validation.isValid) {
        return this.createErrorResult(`地址格式无效: ${validation.error}`, startTime);
      }

      if (!options.data || !Array.isArray(options.data) || options.data.length === 0) {
        return this.createErrorResult("数据不能为空", startTime);
      }

      // 执行Excel操作
      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // 计算实际需要的范围地址
        const dataRows = options.data.length;
        const dataCols = Math.max(...options.data.map(row => row.length));

        let actualAddress = options.address;

        // 如果提供的是单个单元格地址，计算实际范围
        if (validation.type === 'cell' && (dataRows > 1 || dataCols > 1)) {
          actualAddress = ExcelValidator.calculateRangeFromData(options.address, dataRows, dataCols);
          console.log(`📍 自动计算范围: ${options.address} -> ${actualAddress} (${dataRows}行 ${dataCols}列)`);
        }

        const range = sheet.getRange(actualAddress);

        // 加载需要访问的属性
        range.load('address');

        // 设置数据
        range.values = options.data;

        // 自动调整列宽
        if (options.autoFitColumns !== false) {
          range.format.autofitColumns();
        }

        // 自动调整行高
        if (options.autoFitRows === true) {
          range.format.autofitRows();
        }

        // 选中插入的范围
        if (options.selectAfterInsert === true) {
          range.select();
        }

        await context.sync();

        const result: ExcelInsertDataResult = {
          rowCount: options.data.length,
          colCount: options.data[0]?.length || 0,
          actualRange: range.address
        };

        return this.createSuccessResult(
          `成功插入 ${result.rowCount} 行 ${result.colCount} 列数据到 ${actualAddress}`,
          result,
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `插入数据失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 设置单元格或范围的背景颜色
   * @param address 单元格地址，如 'A1' 或 'A1:B2'
   * @param backgroundColor 背景颜色
   * @returns 操作结果
   */
  static async setCellColor(
    address: string,
    backgroundColor: string
  ): Promise<ExcelOperationResult> {
    const startTime = Date.now();

    try {
      // 验证参数
      const validation = ExcelValidator.validateAddress(address);
      if (!validation.isValid) {
        return this.createErrorResult(`地址格式无效: ${validation.error}`, startTime);
      }

      if (!backgroundColor) {
        return this.createErrorResult("背景颜色不能为空", startTime);
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        // 设置背景颜色
        range.format.fill.color = backgroundColor;

        await context.sync();

        return this.createSuccessResult(
          `成功设置 ${address} 的背景颜色为 ${backgroundColor}`,
          { address, backgroundColor },
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `设置颜色失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 批量设置多个范围的背景颜色
   * @param options 批量颜色设置选项，包含范围和颜色的数组
   * @returns 操作结果，包含成功和失败的详细信息
   */
  static async setCellColorsBatch(
    options: ExcelBatchColorOptions
  ): Promise<ExcelOperationResult<ExcelBatchColorResult>> {
    const startTime = Date.now();

    try {
      // 验证参数
      if (!options.colorRanges || !Array.isArray(options.colorRanges) || options.colorRanges.length === 0) {
        return this.createErrorResult("颜色范围数组不能为空", startTime);
      }

      // 验证每个颜色范围项
      for (let i = 0; i < options.colorRanges.length; i++) {
        const item = options.colorRanges[i];
        if (!item.range || !item.color) {
          return this.createErrorResult(`第 ${i + 1} 项的范围或颜色不能为空`, startTime);
        }

        const validation = ExcelValidator.validateAddress(item.range);
        if (!validation.isValid) {
          return this.createErrorResult(`第 ${i + 1} 项地址格式无效: ${validation.error}`, startTime);
        }
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const successItems: Array<{ range: string; color: string }> = [];
        const failedItems: Array<{ range: string; color: string; error: string }> = [];

        // 批量处理每个颜色设置
        for (const colorRange of options.colorRanges) {
          try {
            const range = sheet.getRange(colorRange.range);
            range.format.fill.color = colorRange.color;

            successItems.push({
              range: colorRange.range,
              color: colorRange.color
            });
          } catch (error) {
            failedItems.push({
              range: colorRange.range,
              color: colorRange.color,
              error: error instanceof Error ? error.message : "未知错误"
            });
          }
        }

        await context.sync();

        const result: ExcelBatchColorResult = {
          totalCount: options.colorRanges.length,
          successCount: successItems.length,
          failedCount: failedItems.length,
          successItems,
          failedItems
        };

        const message = failedItems.length === 0
          ? `成功设置 ${successItems.length} 个范围的颜色`
          : `设置完成：${successItems.length} 个成功，${failedItems.length} 个失败`;

        return this.createSuccessResult(
          message,
          result,
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `批量设置颜色失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 设置单元格字体格式
   * @param address 单元格地址
   * @param format 字体格式配置
   * @returns 操作结果
   */
  static async setCellFont(
    address: string,
    format: ExcelCellFormat
  ): Promise<ExcelOperationResult> {
    const startTime = Date.now();

    try {
      // 验证参数
      const validation = ExcelValidator.validateAddress(address);
      if (!validation.isValid) {
        return this.createErrorResult(`地址格式无效: ${validation.error}`, startTime);
      }

      if (!format || Object.keys(format).length === 0) {
        return this.createErrorResult("格式配置不能为空", startTime);
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        // 设置字体格式
        if (format.fontColor) {
          range.format.font.color = format.fontColor;
        }
        if (format.fontSize) {
          range.format.font.size = format.fontSize;
        }
        if (format.fontName) {
          range.format.font.name = format.fontName;
        }
        if (format.bold !== undefined) {
          range.format.font.bold = format.bold;
        }
        if (format.italic !== undefined) {
          range.format.font.italic = format.italic;
        }
        if (format.underline !== undefined) {
          range.format.font.underline = format.underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
        }

        // 设置背景颜色
        if (format.backgroundColor) {
          range.format.fill.color = format.backgroundColor;
        }

        // 设置对齐方式
        if (format.horizontalAlignment) {
          range.format.horizontalAlignment = format.horizontalAlignment;
        }
        if (format.verticalAlignment) {
          range.format.verticalAlignment = format.verticalAlignment;
        }

        // 设置数字格式
        if (format.numberFormat) {
          range.numberFormat = [[format.numberFormat]];
        }

        await context.sync();

        const appliedFormats = Object.keys(format).join(', ');
        return this.createSuccessResult(
          `成功设置 ${address} 的格式: ${appliedFormats}`,
          { address, format },
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `设置字体失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 根据指定数据范围创建图表
   * @param options 图表创建配置
   * @returns 操作结果，包含图表信息
   */
  static async createChart(
    options: ExcelChartOptions
  ): Promise<ExcelOperationResult<ExcelChartResult>> {
    const startTime = Date.now();

    try {
      // 验证数据范围
      const validation = ExcelValidator.validateAddress(options.dataRange);
      if (!validation.isValid) {
        return this.createErrorResult(`数据范围无效: ${validation.error}`, startTime);
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const dataRange = sheet.getRange(options.dataRange);

        // 创建图表
        const chart = sheet.charts.add(
          options.type as Excel.ChartType,
          dataRange,
          "Auto"
        );

        // 设置图表位置
        if (options.position) {
          chart.left = options.position.left;
          chart.top = options.position.top;
          if (options.position.width) {
            chart.width = options.position.width;
          }
          if (options.position.height) {
            chart.height = options.position.height;
          }
        }

        // 设置图表样式
        if (options.style) {
          if (options.style.title) {
            chart.title.text = options.style.title;
            chart.title.visible = true;
          }

          if (options.style.showLegend !== undefined) {
            chart.legend.visible = options.style.showLegend;
          }

          if (options.style.legendPosition) {
            chart.legend.position = options.style.legendPosition;
          }

          if (options.style.showDataLabels !== undefined) {
            // 注意：Excel JavaScript API中数据标签的设置可能因图表类型而异
            try {
              chart.series.getItemAt(0).dataLabels.showValue = options.style.showDataLabels;
            } catch {
              // 某些图表类型可能不支持数据标签
            }
          }

          if (options.style.xAxisTitle) {
            chart.axes.categoryAxis.title.text = options.style.xAxisTitle;
          }

          if (options.style.yAxisTitle) {
            chart.axes.valueAxis.title.text = options.style.yAxisTitle;
          }
        }

        // 加载需要访问的图表属性
        chart.load(['name', 'id']);

        await context.sync();

        const result: ExcelChartResult = {
          chartId: chart.id,
          chartName: chart.name,
          chartType: options.type,
          dataRange: options.dataRange
        };

        return this.createSuccessResult(
          `成功创建 ${options.type} 图表，数据范围: ${options.dataRange}`,
          result,
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `创建图表失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 在指定单元格插入公式
   * @param options 公式插入配置
   * @returns 操作结果，包含公式和计算结果
   */
  static async insertFormula(
    options: ExcelFormulaOptions
  ): Promise<ExcelOperationResult<ExcelFormulaResult>> {
    const startTime = Date.now();

    try {
      // 验证参数
      const validation = ExcelValidator.validateAddress(options.address);
      if (!validation.isValid) {
        return this.createErrorResult(`地址格式无效: ${validation.error}`, startTime);
      }

      if (!options.formula) {
        return this.createErrorResult("公式不能为空", startTime);
      }

      if (!options.formula.startsWith('=')) {
        return this.createErrorResult("公式必须以 '=' 开头", startTime);
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(options.address);

        // 插入公式
        range.formulas = [[options.formula]];

        // 如果需要立即计算
        if (options.calculateImmediately !== false) {
          await context.sync();
          context.workbook.application.calculate(Excel.CalculationType.full);
        }

        await context.sync();

        // 获取计算结果
        range.load("values");
        await context.sync();

        const calculatedValue = range.values[0][0];

        const result: ExcelFormulaResult = {
          formula: options.formula,
          result: calculatedValue,
          resultType: typeof calculatedValue as any
        };

        return this.createSuccessResult(
          `成功在 ${options.address} 插入公式: ${options.formula}，结果: ${calculatedValue}`,
          result,
          startTime
        );
      });

    } catch (error) {
      return this.createErrorResult(
        `插入公式失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 对指定范围应用格式设置
   * @param address 范围地址
   * @param format 格式配置
   * @returns 操作结果
   */
  static async formatRange(
    address: string,
    format: ExcelCellFormat
  ): Promise<ExcelOperationResult> {
    const startTime = Date.now();

    try {
      // 复用setCellFont方法，因为功能相同
      return await this.setCellFont(address, format);
    } catch (error) {
      return this.createErrorResult(
        `格式化范围失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  // =============================================================================
  // 数据获取方法
  // =============================================================================

  /**
   * 获取当前选中的数据
   * @returns 选中区域的数据
   */
  static async getCurrentSelection(): Promise<ExcelOperationResult<ExcelSelectionInfo>> {
    const startTime = Date.now();

    try {
      return await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        const worksheet = context.workbook.worksheets.getActiveWorksheet();

        selection.load(["address", "values", "rowCount", "columnCount"]);
        worksheet.load("name");

        await context.sync();

        const result: ExcelSelectionInfo = {
          address: selection.address,
          values: selection.values,
          headers: selection.rowCount > 0 ? (selection.values[0] as string[]) : undefined,
          rowCount: selection.rowCount,
          columnCount: selection.columnCount,
          worksheetName: worksheet.name,
          cellCount: selection.rowCount * selection.columnCount,
          isSingleCell: selection.rowCount === 1 && selection.columnCount === 1
        };

        return this.createSuccessResult(
          `获取选中区域成功: ${result.address}`,
          result,
          startTime
        );
      });
    } catch (error) {
      return this.createErrorResult(
        `获取选中数据失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 获取当前活动工作表的所有数据
   * @returns 工作表数据
   */
  static async getActiveWorksheetData(): Promise<ExcelOperationResult<ExcelRange>> {
    const startTime = Date.now();

    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = worksheet.getUsedRange();

        if (!usedRange) {
          return this.createSuccessResult(
            "工作表为空",
            {
              address: "A1",
              values: [],
              headers: [],
              rowCount: 0,
              columnCount: 0
            },
            startTime
          );
        }

        usedRange.load(["address", "values", "rowCount", "columnCount"]);
        await context.sync();

        const result: ExcelRange = {
          address: usedRange.address,
          values: usedRange.values,
          headers: usedRange.rowCount > 0 ? (usedRange.values[0] as string[]) : undefined,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount
        };

        return this.createSuccessResult(
          `获取工作表数据成功: ${result.address}`,
          result,
          startTime
        );
      });
    } catch (error) {
      return this.createErrorResult(
        `获取工作表数据失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  /**
   * 获取指定范围的数据
   * @param address 范围地址
   * @returns 范围数据
   */
  static async getRangeData(address: string): Promise<ExcelOperationResult<ExcelRange>> {
    const startTime = Date.now();

    try {
      // 验证地址
      const validation = ExcelValidator.validateAddress(address);
      if (!validation.isValid) {
        return this.createErrorResult(`地址格式无效: ${validation.error}`, startTime);
      }

      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const range = worksheet.getRange(address);

        range.load(["address", "values", "rowCount", "columnCount"]);
        await context.sync();

        const result: ExcelRange = {
          address: range.address,
          values: range.values,
          headers: range.rowCount > 0 ? (range.values[0] as string[]) : undefined,
          rowCount: range.rowCount,
          columnCount: range.columnCount
        };

        return this.createSuccessResult(
          `获取范围数据成功: ${result.address}`,
          result,
          startTime
        );
      });
    } catch (error) {
      return this.createErrorResult(
        `获取范围数据失败: ${error instanceof Error ? error.message : "未知错误"}`,
        startTime
      );
    }
  }

  // =============================================================================
  // 实用工具方法
  // =============================================================================

  /**
   * 创建成功结果
   */
  private static createSuccessResult<T = any>(
    message: string,
    data?: T,
    startTime?: number
  ): ExcelOperationResult<T> {
    return {
      success: true,
      message,
      data,
      duration: startTime ? Date.now() - startTime : undefined
    };
  }

  /**
   * 创建错误结果
   */
  private static createErrorResult<T = any>(
    message: string,
    startTime?: number
  ): ExcelOperationResult<T> {
    return {
      success: false,
      message,
      error: message,
      duration: startTime ? Date.now() - startTime : undefined
    };
  }
}