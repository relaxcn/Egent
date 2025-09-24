/**
 * Excel操作模块统一导出
 * 提供完整的Excel操作功能，包括类型定义、操作函数、验证工具和Function定义
 */

// =============================================================================
// 核心操作类和函数
// =============================================================================
export { ExcelOperations } from "./operations";
export { ExcelValidator } from "./validation";

// 内部使用的导入
import { ExcelOperations } from "./operations";
import { ExcelValidator } from "./validation";

// =============================================================================
// 类型定义导出
// =============================================================================
export type {
  // 基础类型
  ExcelOperationResult,
  ExcelRange,
  ExcelOperationOptions,

  // 格式相关类型
  ExcelFontFormat,
  ExcelBackgroundFormat,
  ExcelBorderFormat,
  ExcelCellFormat,

  // 数据操作类型
  ExcelInsertDataOptions,
  ExcelInsertDataResult,

  // 图表相关类型
  ExcelChartType,
  ExcelChartPosition,
  ExcelChartStyle,
  ExcelChartOptions,
  ExcelChartResult,

  // 公式相关类型
  ExcelFormulaOptions,
  ExcelFormulaResult,

  // 工作表和选择相关类型
  ExcelWorksheetInfo,
  ExcelSelectionInfo,

  // 错误处理类型
  ExcelErrorType,
  ExcelError,

  // 验证相关类型
  ExcelAddressValidation
} from "./types";

// =============================================================================
// OpenAI Function定义导出
// =============================================================================
export {
  // 具体Function定义
  INSERT_DATA_FUNCTION,
  SET_CELL_COLOR_FUNCTION,
  SET_CELL_FONT_FUNCTION,
  CREATE_CHART_FUNCTION,
  INSERT_FORMULA_FUNCTION,
  FORMAT_RANGE_FUNCTION,

  // 集合和工具
  EXCEL_FUNCTIONS,
  EXCEL_FUNCTION_NAMES,
  getExcelFunction,
  validateFunctionCall
} from "./function-definitions";

// =============================================================================
// 便捷操作接口
// =============================================================================

/**
 * Excel操作的便捷接口类
 * 提供更简洁的方法调用方式，封装常用操作模式
 */
export class ExcelHelper {
  /**
   * 快速插入表格数据（包含标题行）
   * @param address 起始地址
   * @param headers 标题行数据
   * @param data 数据行
   */
  static async insertTable(
    address: string,
    headers: string[],
    data: any[][]
  ) {
    const tableData = [headers, ...data];
    return ExcelOperations.insertData({
      address,
      data: tableData,
      includeHeaders: true,
      autoFitColumns: true,
      selectAfterInsert: true
    });
  }

  /**
   * 快速创建带标题的柱状图
   * @param title 图表标题
   * @param dataRange 数据范围
   * @param position 图表位置（可选）
   */
  static async createColumnChart(
    title: string,
    dataRange: string,
    position?: { left: number; top: number }
  ) {
    return ExcelOperations.createChart({
      type: "ColumnClustered",
      dataRange,
      style: {
        title,
        showLegend: true,
        showDataLabels: false
      },
      position: position ? { ...position, width: 400, height: 300 } : undefined
    });
  }

  /**
   * 快速格式化表格标题行
   * @param address 标题行地址
   * @param backgroundColor 背景色（默认蓝色）
   */
  static async formatTableHeader(
    address: string,
    backgroundColor: string = "#4472C4"
  ) {
    return ExcelOperations.setCellFont(address, {
      bold: true,
      fontColor: "#FFFFFF",
      backgroundColor,
      horizontalAlignment: "Center" as any,
      verticalAlignment: "Center" as any
    });
  }

  /**
   * 快速插入汇总行
   * @param address 汇总行地址
   * @param dataRange 要汇总的数据范围
   * @param label 汇总标签（如"合计"）
   */
  static async insertSummaryRow(
    address: string,
    dataRange: string,
    label: string = "合计"
  ) {
    // 插入标签
    const labelAddress = address.split(':')[0]; // 获取起始地址
    await ExcelOperations.insertData({
      address: labelAddress,
      data: [[label]]
    });

    // 计算汇总地址（假设是右边一列）
    const cellMatch = labelAddress.match(/^([A-Z]+)(\d+)$/);
    if (cellMatch) {
      const nextCol = ExcelValidator.numberToColumn(
        ExcelValidator['columnToNumber'](cellMatch[1]) + 1
      );
      const summaryAddress = `${nextCol}${cellMatch[2]}`;

      // 插入求和公式
      await ExcelOperations.insertFormula({
        address: summaryAddress,
        formula: `=SUM(${dataRange})`
      });

      // 格式化汇总行
      await ExcelOperations.setCellFont(address, {
        bold: true,
        backgroundColor: "#F2F2F2"
      });
    }
  }

  /**
   * 快速创建数据验证下拉列表
   * @param address 目标地址
   * @param options 下拉选项数组
   */
  static async createDropdownList(
    address: string,
    options: string[]
  ) {
    // 注意：这里需要Excel的数据验证功能
    // 当前Excel JavaScript API可能不完全支持，这是一个示例实现
    console.warn(`下拉列表功能需要Excel高级API支持，当前可能不可用。目标地址: ${address}`);

    // 作为替代，可以在某个区域插入选项列表作为参考
    const optionsData = options.map(option => [option]);
    const helperAddress = "AA1:AA" + options.length;

    await ExcelOperations.insertData({
      address: helperAddress,
      data: optionsData
    });

    return {
      success: true,
      message: `选项列表已插入到 ${helperAddress}，请手动设置数据验证`
    };
  }
}

// =============================================================================
// 常用常量导出
// =============================================================================

/**
 * 常用的Excel颜色常量
 */
export const EXCEL_COLORS = {
  // 主题色
  PRIMARY_BLUE: "#4472C4",
  PRIMARY_RED: "#E15759",
  PRIMARY_GREEN: "#70AD47",
  PRIMARY_ORANGE: "#FFC000",
  PRIMARY_PURPLE: "#7030A0",

  // 状态色
  SUCCESS_GREEN: "#2ECC71",
  WARNING_YELLOW: "#F39C12",
  ERROR_RED: "#E74C3C",
  INFO_BLUE: "#3498DB",

  // 灰度色
  LIGHT_GRAY: "#F8F9FA",
  MEDIUM_GRAY: "#6C757D",
  DARK_GRAY: "#343A40",

  // 文本色
  TEXT_BLACK: "#000000",
  TEXT_WHITE: "#FFFFFF",
  TEXT_GRAY: "#666666"
} as const;

/**
 * 常用的Excel字体常量
 */
export const EXCEL_FONTS = {
  DEFAULT: "Arial",
  MICROSOFT_YAHEI: "微软雅黑",
  SONG_TI: "宋体",
  TIMES_NEW_ROMAN: "Times New Roman",
  CALIBRI: "Calibri"
} as const;

/**
 * 常用的Excel字体大小常量
 */
export const EXCEL_FONT_SIZES = {
  SMALL: 9,
  NORMAL: 11,
  MEDIUM: 12,
  LARGE: 14,
  TITLE: 16,
  HEADER: 18
} as const;