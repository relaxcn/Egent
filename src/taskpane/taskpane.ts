/* global Excel console */

export interface ExcelData {
  address: string;
  values: any[][];
  headers?: string[];
}

export async function insertText(text: string) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:A10");
      // 创建一个包含10个相同文本的数组
      const values = Array(10).fill([text]);
      range.values = values;
      range.format.autofitColumns();
      range.format.autofitRows();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getSelectedData(): Promise<ExcelData | null> {
  try {
    return await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      selection.load(["address", "values", "rowCount", "columnCount"]);

      await context.sync();

      if (selection.rowCount === 0 || selection.columnCount === 0) {
        return null;
      }

      const data: ExcelData = {
        address: selection.address,
        values: selection.values as any[][],
      };

      // 如果选择的第一行看起来像标题行，将其作为headers
      if (
        selection.rowCount > 1 &&
        selection.values[0].every((cell) => typeof cell === "string" && cell.trim() !== "")
      ) {
        data.headers = selection.values[0] as string[];
        data.values = selection.values.slice(1) as any[][];
      }

      return data;
    });
  } catch (error) {
    console.error("Error reading Excel data:", error);
    return null;
  }
}

export async function getAllWorksheetData(): Promise<ExcelData | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();

      if (!usedRange) {
        return null;
      }

      usedRange.load(["address", "values", "rowCount", "columnCount"]);

      await context.sync();

      if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
        return null;
      }

      const data: ExcelData = {
        address: usedRange.address,
        values: usedRange.values as any[][],
      };

      // 如果第一行看起来像标题行，将其作为headers
      if (
        usedRange.rowCount > 1 &&
        usedRange.values[0].every((cell) => typeof cell === "string" && cell.trim() !== "")
      ) {
        data.headers = usedRange.values[0] as string[];
        data.values = usedRange.values.slice(1) as any[][];
      }

      return data;
    });
  } catch (error) {
    console.error("Error reading worksheet data:", error);
    return null;
  }
}
