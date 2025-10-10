/* global console */

/**
 * 提示词管理器
 * 负责加载和管理 Markdown 格式的系统提示词
 */

import { ChatMode } from "./types";

// 导入 Markdown 提示词内容
import baseSystemPrompt from "./prompts/base-system.md";
import agentSystemPrompt from "./prompts/agent-system.md";
import excelOperationsPrompt from "./prompts/excel-operations.md";

/**
 * 提示词管理器类
 */
export class PromptManager {
  private static basePrompt: string = "";
  private static agentPrompt: string = "";
  private static excelOpsPrompt: string = "";

  /**
   * 初始化提示词（从 Markdown 文件加载）
   */
  static async initialize(): Promise<void> {
    try {
      // 在浏览器环境中，直接使用导入的内容
      // 如果是字符串，直接使用；如果是模块，获取其 default 导出
      this.basePrompt = typeof baseSystemPrompt === "string"
        ? baseSystemPrompt
        : await this.fetchPrompt("base-system.md");

      this.agentPrompt = typeof agentSystemPrompt === "string"
        ? agentSystemPrompt
        : await this.fetchPrompt("agent-system.md");

      this.excelOpsPrompt = typeof excelOperationsPrompt === "string"
        ? excelOperationsPrompt
        : await this.fetchPrompt("excel-operations.md");

      console.log("✅ 提示词管理器初始化成功");
    } catch (error) {
      console.error("❌ 提示词管理器初始化失败:", error);
      // 使用后备提示词
      this.loadFallbackPrompts();
    }
  }

  /**
   * 从文件路径获取提示词内容
   */
  private static async fetchPrompt(filename: string): Promise<string> {
    const response = await fetch(`/prompts/${filename}`);
    if (!response.ok) {
      throw new Error(`Failed to fetch prompt: ${filename}`);
    }
    return await response.text();
  }

  /**
   * 加载后备提示词（当 Markdown 文件加载失败时）
   */
  private static loadFallbackPrompts(): void {
    this.basePrompt = `你是 Egent，一个专门帮助用户分析 Excel 数据的 AI 助手。你只能查看和分析数据，不能修改 Excel 文件。

请遵循以下规则：
1. 专注于数据分析、解释和见解
2. 提供清晰、易懂的回答
3. 如果需要，可以建议数据可视化的方法
4. 始终以中文回答
5. 保持友好和专业的语气

如果用户提供了 Excel 数据，请基于该数据进行分析和回答。`;

    this.agentPrompt = `你是 Egent Agent，一个专门帮助用户分析和操作 Excel 数据的 AI 助手。

请遵循以下规则：
1. 专注于数据分析、解释和见解，同时可以直接执行 Excel 操作
2. 提供清晰、易懂的回答
3. 如果需要，可以建议数据可视化的方法，并直接创建图表
4. 始终以中文回答
5. 保持友好和专业的语气
6. 当需要执行 Excel 操作时，主动使用可用的工具函数

你可以直接操作 Excel，包括以下功能：
- insertData: 插入数据到指定范围
- setCellColor: 设置单元格背景色
- setCellColorsBatch: 批量设置多个范围的背景色
- setCellFont: 设置字体格式
- createChart: 创建各种类型的图表
- insertFormula: 插入 Excel 公式
- formatRange: 批量格式化单元格范围

当用户需要 Excel 操作时，请主动调用相应的工具函数。`;

    this.excelOpsPrompt = "";
  }

  /**
   * 获取基础系统提示词（只读分析模式）
   */
  static getBaseSystemPrompt(): string {
    if (!this.basePrompt) {
      this.loadFallbackPrompts();
    }
    return this.basePrompt;
  }

  /**
   * 获取 Agent 系统提示词（操作模式）
   */
  static getAgentSystemPrompt(): string {
    if (!this.agentPrompt) {
      this.loadFallbackPrompts();
    }
    return this.agentPrompt;
  }

  /**
   * 获取 Excel 操作参考文档
   */
  static getExcelOperationsPrompt(): string {
    if (!this.excelOpsPrompt) {
      this.loadFallbackPrompts();
    }
    return this.excelOpsPrompt;
  }

  /**
   * 根据聊天模式获取相应的系统提示词
   * @param mode 聊天模式
   * @returns 系统提示词
   */
  static getSystemPromptByMode(mode: ChatMode): string {
    switch (mode) {
      case ChatMode.AGENT:
        return this.getAgentSystemPrompt();
      case ChatMode.ANALYSIS:
      default:
        return this.getBaseSystemPrompt();
    }
  }

  /**
   * 自定义系统提示词（允许运行时覆盖）
   * @param mode 模式
   * @param customPrompt 自定义提示词
   */
  static setCustomPrompt(mode: ChatMode, customPrompt: string): void {
    if (mode === ChatMode.AGENT) {
      this.agentPrompt = customPrompt;
    } else {
      this.basePrompt = customPrompt;
    }
  }

  /**
   * 重置为默认提示词
   */
  static reset(): void {
    this.basePrompt = "";
    this.agentPrompt = "";
    this.excelOpsPrompt = "";
    this.initialize();
  }
}

// 自动初始化
PromptManager.initialize();
