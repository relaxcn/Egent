import { ExcelAddressValidation } from "./types";

/**
 * Excel地址验证工具类
 * 提供Excel地址格式验证和规范化功能
 */
export class ExcelValidator {
  // Excel地址的正则表达式模式
  private static readonly CELL_PATTERN = /^[A-Z]+[1-9]\d*$/;
  private static readonly RANGE_PATTERN = /^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$/;
  private static readonly COLUMN_PATTERN = /^[A-Z]+:[A-Z]+$/;
  private static readonly ROW_PATTERN = /^[1-9]\d*:[1-9]\d*$/;

  /**
   * 验证Excel地址格式
   * @param address 要验证的地址
   * @returns 验证结果，包含地址类型和规范化信息
   */
  static validateAddress(address: string): ExcelAddressValidation {
    if (!address || typeof address !== 'string') {
      return {
        isValid: false,
        error: '地址不能为空'
      };
    }

    // 去除空格并转换为大写
    const normalizedAddress = address.trim().toUpperCase();

    // 验证单个单元格（如：A1, B2, AA100）
    if (this.CELL_PATTERN.test(normalizedAddress)) {
      return {
        isValid: true,
        type: 'cell',
        normalizedAddress,
        startCell: normalizedAddress,
        endCell: normalizedAddress
      };
    }

    // 验证单元格范围（如：A1:B2, C1:D10）
    if (this.RANGE_PATTERN.test(normalizedAddress)) {
      const [startCell, endCell] = normalizedAddress.split(':');

      // 验证起始单元格是否在结束单元格之前
      if (this.compareCells(startCell, endCell) > 0) {
        return {
          isValid: false,
          error: '起始单元格不能大于结束单元格'
        };
      }

      return {
        isValid: true,
        type: 'range',
        normalizedAddress,
        startCell,
        endCell
      };
    }

    // 验证列范围（如：A:B, C:Z）
    if (this.COLUMN_PATTERN.test(normalizedAddress)) {
      const [startCol, endCol] = normalizedAddress.split(':');

      if (this.compareColumns(startCol, endCol) > 0) {
        return {
          isValid: false,
          error: '起始列不能大于结束列'
        };
      }

      return {
        isValid: true,
        type: 'column',
        normalizedAddress,
        startCell: `${startCol}1`,
        endCell: `${endCol}1048576` // Excel最大行数
      };
    }

    // 验证行范围（如：1:5, 10:20）
    if (this.ROW_PATTERN.test(normalizedAddress)) {
      const [startRow, endRow] = normalizedAddress.split(':').map(Number);

      if (startRow > endRow) {
        return {
          isValid: false,
          error: '起始行不能大于结束行'
        };
      }

      return {
        isValid: true,
        type: 'row',
        normalizedAddress,
        startCell: `A${startRow}`,
        endCell: `XFD${endRow}` // Excel最大列为XFD
      };
    }

    return {
      isValid: false,
      error: '地址格式无效。支持的格式：A1（单元格）、A1:B2（范围）、A:B（列）、1:5（行）'
    };
  }

  /**
   * 比较两个单元格的位置
   * @param cell1 第一个单元格
   * @param cell2 第二个单元格
   * @returns 比较结果：-1（cell1 < cell2）、0（相等）、1（cell1 > cell2）
   */
  private static compareCells(cell1: string, cell2: string): number {
    const { col: col1, row: row1 } = this.parseCellAddress(cell1);
    const { col: col2, row: row2 } = this.parseCellAddress(cell2);

    // 先比较行
    if (row1 !== row2) {
      return row1 - row2;
    }

    // 行相同，比较列
    return this.compareColumns(col1, col2);
  }

  /**
   * 比较两个列的位置
   * @param col1 第一个列
   * @param col2 第二个列
   * @returns 比较结果：-1（col1 < col2）、0（相等）、1（col1 > col2）
   */
  private static compareColumns(col1: string, col2: string): number {
    const num1 = this.columnToNumber(col1);
    const num2 = this.columnToNumber(col2);

    if (num1 < num2) return -1;
    if (num1 > num2) return 1;
    return 0;
  }

  /**
   * 解析单元格地址，分离列和行
   * @param cell 单元格地址（如：A1, B2）
   * @returns 解析结果，包含列字母和行数字
   */
  private static parseCellAddress(cell: string): { col: string; row: number } {
    const match = cell.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`无效的单元格地址: ${cell}`);
    }

    return {
      col: match[1],
      row: parseInt(match[2])
    };
  }

  /**
   * 将列字母转换为数字（A=1, B=2, ..., Z=26, AA=27, ...）
   * @param col 列字母
   * @returns 列数字
   */
  private static columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * 将数字转换为列字母（1=A, 2=B, ..., 26=Z, 27=AA, ...）
   * @param num 列数字
   * @returns 列字母
   */
  static numberToColumn(num: number): string {
    let result = '';
    while (num > 0) {
      num--;
      result = String.fromCharCode((num % 26) + 65) + result;
      num = Math.floor(num / 26);
    }
    return result;
  }

  /**
   * 验证数据是否适合指定的范围
   * @param address 目标范围地址
   * @param data 要插入的数据
   * @returns 验证结果
   */
  static validateDataForRange(address: string, data: any[][]): ExcelAddressValidation {
    const addressValidation = this.validateAddress(address);
    if (!addressValidation.isValid) {
      return addressValidation;
    }

    if (!data || !Array.isArray(data) || data.length === 0) {
      return {
        isValid: false,
        error: '数据不能为空'
      };
    }

    // 检查数据是否为二维数组
    if (!data.every(row => Array.isArray(row))) {
      return {
        isValid: false,
        error: '数据必须是二维数组'
      };
    }

    // 如果是单个单元格
    if (addressValidation.type === 'cell') {
      if (data.length !== 1 || data[0].length !== 1) {
        return {
          isValid: false,
          error: '单个单元格只能插入一个值'
        };
      }
      return { isValid: true };
    }

    // 如果是范围，计算范围大小
    if (addressValidation.type === 'range' && addressValidation.startCell && addressValidation.endCell) {
      const startCell = this.parseCellAddress(addressValidation.startCell);
      const endCell = this.parseCellAddress(addressValidation.endCell);

      const rangeRows = endCell.row - startCell.row + 1;
      const rangeCols = this.columnToNumber(endCell.col) - this.columnToNumber(startCell.col) + 1;

      const dataRows = data.length;
      const dataCols = Math.max(...data.map(row => row.length));

      if (dataRows > rangeRows || dataCols > rangeCols) {
        return {
          isValid: false,
          error: `数据大小 (${dataRows}×${dataCols}) 超过了范围大小 (${rangeRows}×${rangeCols})`
        };
      }
    }

    return { isValid: true };
  }

  /**
   * 生成范围地址
   * @param startCell 起始单元格
   * @param endCell 结束单元格
   * @returns 范围地址字符串
   */
  static createRangeAddress(startCell: string, endCell: string): string {
    return `${startCell}:${endCell}`;
  }

  /**
   * 根据起始位置和数据大小计算范围地址
   * @param startCell 起始单元格（如：A1）
   * @param rows 数据行数
   * @param cols 数据列数
   * @returns 计算出的范围地址
   */
  static calculateRangeFromData(startCell: string, rows: number, cols: number): string {
    const { col: startCol, row: startRow } = this.parseCellAddress(startCell);
    const startColNum = this.columnToNumber(startCol);

    const endRow = startRow + rows - 1;
    const endColNum = startColNum + cols - 1;
    const endCol = this.numberToColumn(endColNum);

    const endCell = `${endCol}${endRow}`;

    if (rows === 1 && cols === 1) {
      return startCell;
    }

    return this.createRangeAddress(startCell, endCell);
  }
}