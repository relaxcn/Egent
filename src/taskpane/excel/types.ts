/* global Excel */

/**
 * Excel操作相关的类型定义
 * 提供完整的类型安全和接口定义
 */

// =============================================================================
// 基础类型定义
// =============================================================================

/**
 * Excel操作结果的统一接口
 */
export interface ExcelOperationResult<T = any> {
  /** 操作是否成功 */
  success: boolean;
  /** 操作消息（成功或失败描述） */
  message: string;
  /** 操作返回的数据 */
  data?: T;
  /** 错误信息（操作失败时） */
  error?: string;
  /** 操作耗时（毫秒） */
  duration?: number;
}

/**
 * Excel范围数据
 */
export interface ExcelRange {
  /** 范围地址，如 'A1:C3' */
  address: string;
  /** 范围中的值（二维数组） */
  values: any[][];
  /** 标题行（可选） */
  headers?: string[];
  /** 行数 */
  rowCount: number;
  /** 列数 */
  columnCount: number;
}

// =============================================================================
// 单元格格式相关类型
// =============================================================================

/**
 * 单元格字体格式配置
 */
export interface ExcelFontFormat {
  /** 字体颜色，支持 '#FF0000' 或 'red' 格式 */
  fontColor?: string;
  /** 字体大小，如 12, 14, 16 等 */
  fontSize?: number;
  /** 字体名称，如 'Arial', 'Times New Roman' 等 */
  fontName?: string;
  /** 是否加粗 */
  bold?: boolean;
  /** 是否斜体 */
  italic?: boolean;
  /** 是否下划线 */
  underline?: boolean;
}

/**
 * 单元格背景格式配置
 */
export interface ExcelBackgroundFormat {
  /** 背景颜色，支持 '#FF0000' 或 'red' 格式 */
  backgroundColor?: string;
}

/**
 * 单元格边框格式配置
 */
export interface ExcelBorderFormat {
  /** 边框颜色 */
  color?: string;
  /** 边框样式 */
  style?: Excel.BorderLineStyle;
  /** 边框粗细 */
  weight?: Excel.BorderWeight;
}

/**
 * 完整的单元格格式配置
 * 结合字体、背景、边框等所有格式选项
 */
export interface ExcelCellFormat extends ExcelFontFormat, ExcelBackgroundFormat {
  /** 边框格式 */
  borders?: {
    top?: ExcelBorderFormat;
    bottom?: ExcelBorderFormat;
    left?: ExcelBorderFormat;
    right?: ExcelBorderFormat;
  };
  /** 数字格式，如 '@'（文本）、'0.00'（小数）等 */
  numberFormat?: string;
  /** 水平对齐方式 */
  horizontalAlignment?: Excel.HorizontalAlignment;
  /** 垂直对齐方式 */
  verticalAlignment?: Excel.VerticalAlignment;
}

// =============================================================================
// 数据操作相关类型
// =============================================================================

/**
 * 单个颜色设置项：范围地址和颜色的对
 */
export interface ExcelColorRange {
  /** 范围地址，如 'A1' 或 'A1:B2' */
  range: string;
  /** 背景颜色，如 '#FF0000' 或 'red' */
  color: string;
}

/**
 * 批量设置颜色的配置选项
 */
export interface ExcelBatchColorOptions {
  /** 颜色范围数组，每个元素包含范围和颜色 */
  colorRanges: ExcelColorRange[];
}

/**
 * 批量颜色设置结果
 */
export interface ExcelBatchColorResult {
  /** 成功设置的范围数量 */
  successCount: number;
  /** 失败的范围数量 */
  failedCount: number;
  /** 总共处理的范围数量 */
  totalCount: number;
  /** 成功设置的范围详情 */
  successItems: Array<{ range: string; color: string }>;
  /** 失败的范围详情 */
  failedItems: Array<{ range: string; color: string; error: string }>;
}

/**
 * 插入数据的配置选项
 */
export interface ExcelInsertDataOptions {
  /** 目标范围地址，如 'A1:C3' */
  address: string;
  /** 要插入的数据（二维数组） */
  data: any[][];
  /** 是否包含标题行 */
  includeHeaders?: boolean;
  /** 是否自动调整列宽 */
  autoFitColumns?: boolean;
  /** 是否自动调整行高 */
  autoFitRows?: boolean;
  /** 插入后是否选中范围 */
  selectAfterInsert?: boolean;
}

/**
 * 数据插入结果
 */
export interface ExcelInsertDataResult {
  /** 插入的行数 */
  rowCount: number;
  /** 插入的列数 */
  colCount: number;
  /** 实际插入的范围地址 */
  actualRange: string;
}

// =============================================================================
// 图表相关类型
// =============================================================================

/**
 * 图表类型枚举（扩展Excel原有类型）
 */
export type ExcelChartType =
  | "ColumnClustered"
  | "ColumnStacked"
  | "ColumnStacked100"
  | "Line"
  | "LineMarkers"
  | "LineMarkersStacked"
  | "Pie"
  | "PieExploded"
  | "XYScatter"
  | "XYScatterLines"
  | "Area"
  | "AreaStacked"
  | "Doughnut"
  | "Radar";

/**
 * 图表位置配置
 */
export interface ExcelChartPosition {
  /** 左边距（像素） */
  left: number;
  /** 上边距（像素） */
  top: number;
  /** 宽度（像素，可选，默认自动计算） */
  width?: number;
  /** 高度（像素，可选，默认自动计算） */
  height?: number;
}

/**
 * 图表样式配置
 */
export interface ExcelChartStyle {
  /** 图表标题 */
  title?: string;
  /** 是否显示图例 */
  showLegend?: boolean;
  /** 图例位置 */
  legendPosition?: Excel.ChartLegendPosition;
  /** 是否显示数据标签 */
  showDataLabels?: boolean;
  /** X轴标题 */
  xAxisTitle?: string;
  /** Y轴标题 */
  yAxisTitle?: string;
}

/**
 * 创建图表的完整配置选项
 */
export interface ExcelChartOptions {
  /** 图表类型 */
  type: ExcelChartType;
  /** 数据范围，如 'A1:B5' */
  dataRange: string;
  /** 图表位置（可选，默认自动放置） */
  position?: ExcelChartPosition;
  /** 图表样式配置（可选） */
  style?: ExcelChartStyle;
}

/**
 * 图表创建结果
 */
export interface ExcelChartResult {
  /** 图表ID */
  chartId: string;
  /** 图表名称 */
  chartName: string;
  /** 图表类型 */
  chartType: ExcelChartType;
  /** 数据范围 */
  dataRange: string;
}

// =============================================================================
// 公式相关类型
// =============================================================================

/**
 * 插入公式的配置选项
 */
export interface ExcelFormulaOptions {
  /** 目标单元格地址，如 'A1' */
  address: string;
  /** 公式内容，如 '=SUM(A1:A10)' */
  formula: string;
  /** 是否立即计算结果 */
  calculateImmediately?: boolean;
}

/**
 * 公式插入结果
 */
export interface ExcelFormulaResult {
  /** 公式内容 */
  formula: string;
  /** 计算结果（如果可用） */
  result?: any;
  /** 结果类型 */
  resultType?: "number" | "string" | "boolean" | "error";
}

// =============================================================================
// 工作表和范围相关类型
// =============================================================================

/**
 * 工作表信息
 */
export interface ExcelWorksheetInfo {
  /** 工作表名称 */
  name: string;
  /** 工作表ID */
  id: string;
  /** 是否为活动工作表 */
  isActive: boolean;
  /** 已使用的范围 */
  usedRange?: string;
  /** 行数 */
  rowCount?: number;
  /** 列数 */
  columnCount?: number;
}

/**
 * 选中区域信息
 */
export interface ExcelSelectionInfo extends ExcelRange {
  /** 工作表名称 */
  worksheetName: string;
  /** 选中的单元格数量 */
  cellCount: number;
  /** 是否为单个单元格 */
  isSingleCell: boolean;
}

// =============================================================================
// 错误处理类型
// =============================================================================

/**
 * Excel操作错误类型
 */
export enum ExcelErrorType {
  /** 一般错误 */
  GENERAL = "GENERAL",
  /** 范围地址无效 */
  INVALID_RANGE = "INVALID_RANGE",
  /** 数据格式错误 */
  INVALID_DATA = "INVALID_DATA",
  /** 权限错误 */
  PERMISSION_DENIED = "PERMISSION_DENIED",
  /** 工作表不存在 */
  WORKSHEET_NOT_FOUND = "WORKSHEET_NOT_FOUND",
  /** 公式语法错误 */
  FORMULA_ERROR = "FORMULA_ERROR",
  /** 网络连接错误 */
  NETWORK_ERROR = "NETWORK_ERROR",
  /** 超时错误 */
  TIMEOUT_ERROR = "TIMEOUT_ERROR"
}

/**
 * Excel操作错误详细信息
 */
export interface ExcelError {
  /** 错误类型 */
  type: ExcelErrorType;
  /** 错误消息 */
  message: string;
  /** 错误代码（Excel API返回的） */
  code?: string;
  /** 出错的操作上下文 */
  context?: {
    operation: string;
    address?: string;
    data?: any;
  };
  /** 原始错误对象 */
  originalError?: Error;
}

// =============================================================================
// 实用工具类型
// =============================================================================

/**
 * Excel地址验证结果
 */
export interface ExcelAddressValidation {
  /** 是否有效 */
  isValid: boolean;
  /** 地址类型：'cell' | 'range' | 'column' | 'row' */
  type?: "cell" | "range" | "column" | "row";
  /** 规范化后的地址 */
  normalizedAddress?: string;
  /** 起始单元格 */
  startCell?: string;
  /** 结束单元格（范围时） */
  endCell?: string;
  /** 错误信息（无效时） */
  error?: string;
}

/**
 * Excel操作配置的通用选项
 */
export interface ExcelOperationOptions {
  /** 操作超时时间（毫秒），默认30秒 */
  timeout?: number;
  /** 是否显示进度指示器 */
  showProgress?: boolean;
  /** 操作失败时是否自动重试 */
  autoRetry?: boolean;
  /** 重试次数（默认3次） */
  retryCount?: number;
  /** 重试间隔（毫秒，默认1000ms） */
  retryDelay?: number;
}