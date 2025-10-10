/* global console */

/**
 * Excel 数据格式化工具
 */

import { ExcelData } from "../../../taskpane";

/**
 * 格式化 Excel 数据为 AI 可读的文本格式
 * @param dataList Excel 数据（单个或数组）
 * @returns 格式化后的文本
 */
export function formatExcelDataForAI(dataList: ExcelData | ExcelData[]): string {
  const dataArray = Array.isArray(dataList) ? dataList : [dataList];

  if (dataArray.length === 0) return "";

  let formatted =
    dataArray.length === 1 ? `Excel 数据:\n` : `Excel 数据 (${dataArray.length} 个数据集):\n`;

  dataArray.forEach((data, index) => {
    // 数据集标题
    if (dataArray.length > 1) {
      formatted += `\n### 数据集 ${index + 1}: ${data.address}\n`;
    } else {
      formatted += `数据来源: ${data.address}\n`;
    }

    // 带标题的数据
    if (data.headers) {
      formatted += `列标题: ${data.headers.join(", ")}\n`;
      formatted += `数据 (${data.values.length} 行):\n`;

      // 格式化为键值对形式
      data.values.slice(0, 20).forEach((row, rowIndex) => {
        formatted += `行${rowIndex + 1}: `;
        row.forEach((cell, cellIndex) => {
          const header = data.headers![cellIndex] || `列${cellIndex + 1}`;
          formatted += `${header}=${cell}, `;
        });
        formatted += "\n";
      });

      if (data.values.length > 20) {
        formatted += `... (还有 ${data.values.length - 20} 行数据)\n`;
      }
    }
    // 无标题的数据
    else {
      formatted += `数据 (${data.values.length} 行 x ${data.values[0]?.length || 0} 列):\n`;
      data.values.slice(0, 10).forEach((row, rowIndex) => {
        formatted += `行${rowIndex + 1}: [${row.join(", ")}]\n`;
      });

      if (data.values.length > 10) {
        formatted += `... (还有 ${data.values.length - 10} 行数据)\n`;
      }
    }

    // 数据集之间的分隔
    if (dataArray.length > 1 && index < dataArray.length - 1) {
      formatted += "\n";
    }
  });

  return formatted;
}

/**
 * 截断过长的文本
 * @param text 原始文本
 * @param maxLength 最大长度
 * @returns 截断后的文本
 */
export function truncateText(text: string, maxLength: number = 1000): string {
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + "...";
}

/**
 * 格式化工具调用信息
 * @param toolName 工具名称
 * @param args 参数对象
 * @returns 格式化后的描述
 */
export function formatToolCall(toolName: string, args: Record<string, any>): string {
  const argsStr = JSON.stringify(args, null, 2);
  return `调用工具: ${toolName}\n参数:\n${argsStr}`;
}
