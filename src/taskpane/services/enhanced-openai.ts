/* global console */

import { OpenAIService, ChatMessage, StreamCallback } from "./openai";
import { ExcelData } from "../taskpane";
import { ApiSettings } from "../utils/settings";
import OpenAI from "openai";

// 导入Excel操作相关模块
import {
  ExcelOperations,
  ExcelOperationResult,
  EXCEL_FUNCTIONS,
  validateFunctionCall
} from "../excel";

/**
 * Excel操作结果的接口定义
 */
export interface ExcelActionResult {
  success: boolean;
  content: string;
  error?: string;
}

/**
 * 增强的流式回调接口，支持Excel操作反馈
 */
export interface EnhancedStreamCallback extends StreamCallback {
  onExcelAction?: (action: string, result: ExcelActionResult) => void;
}

/**
 * 增强的OpenAI服务
 * 直接集成Excel操作功能，使用LLM的function calling能力
 * 在浏览器环境中直接调用Excel API，无需外部依赖
 */
export class EnhancedOpenAIService extends OpenAIService {
  private enableExcelOperations: boolean = true;

  constructor(settings: ApiSettings) {
    super(settings);
  }

  /**
   * 启用/禁用Excel操作功能
   */
  setExcelOperationsEnabled(enabled: boolean): void {
    this.enableExcelOperations = enabled;
  }

  /**
   * 获取Excel操作的OpenAI函数定义
   */
  private getExcelFunctions(): OpenAI.ChatCompletionTool[] {
    return this.enableExcelOperations ? EXCEL_FUNCTIONS : [];
  }

  /**
   * 增强的流式消息发送
   * 使用LLM的function calling能力自动调用Excel操作
   */
  async sendMessageStreamWithExcel(
    userMessage: string,
    callbacks: EnhancedStreamCallback,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): Promise<void> {
    try {
      // 1. 获取可用的Excel函数定义
      const tools = this.getExcelFunctions();

      // 2. 构建消息
      const messages = this.buildEnhancedMessages(
        userMessage,
        excelData,
        conversationHistory,
        this.getSystemPromptWithExcelOperations()
      );

      console.log(`🔧 已加载 ${tools.length} 个Excel操作函数`);

      // 3. 调用OpenAI API，包含function calling
      const stream = await this.client.chat.completions.create({
        model: this.settings.model,
        messages,
        temperature: 0.7,
        max_tokens: 1500,
        stream: true,
        tools: tools.length > 0 ? tools : undefined,
        tool_choice: tools.length > 0 ? "auto" : undefined,
      });

      let functionCallBuffer = "";
      let isCollectingFunctionCall = false;
      let currentFunctionCall: { name?: string; arguments?: string } = {};

      for await (const chunk of stream) {
        const choice = chunk.choices[0];

        if (!choice) continue;

        // 处理工具调用
        if (choice.delta.tool_calls && choice.delta.tool_calls.length > 0) {
          const toolCall = choice.delta.tool_calls[0];

          if (toolCall.function) {
            if (toolCall.function.name) {
              currentFunctionCall.name = toolCall.function.name;
              isCollectingFunctionCall = true;
              console.log(`🔧 LLM请求执行Excel操作: ${toolCall.function.name}`);
            }

            if (toolCall.function.arguments) {
              functionCallBuffer += toolCall.function.arguments;
            }
          }
        }

        // 处理普通文本内容
        else if (choice.delta.content) {
          callbacks.onChunk(choice.delta.content);
        }

        // 如果工具调用结束
        if (choice.finish_reason === "tool_calls" && isCollectingFunctionCall) {
          try {
            const args = JSON.parse(functionCallBuffer);
            console.log(`🔧 Excel操作参数:`, args);

            // 验证函数调用参数
            const validation = validateFunctionCall(currentFunctionCall.name!, args);
            if (!validation.valid) {
              throw new Error(`参数验证失败: ${validation.error}`);
            }

            // 执行Excel操作
            const result = await this.executeExcelOperation(currentFunctionCall.name!, args);

            // 通知UI操作结果
            if (callbacks.onExcelAction) {
              callbacks.onExcelAction(
                `${currentFunctionCall.name}: ${JSON.stringify(args)}`,
                result
              );
            }

            console.log(result.success ? "✅" : "❌", result.content);

            // 将操作结果反馈给LLM
            const followUpMessages = [
              ...messages,
              {
                role: "assistant" as const,
                tool_calls: [
                  {
                    id: "call_" + Date.now(),
                    type: "function" as const,
                    function: {
                      name: currentFunctionCall.name!,
                      arguments: functionCallBuffer,
                    },
                  },
                ],
              },
              {
                role: "tool" as const,
                tool_call_id: "call_" + Date.now(),
                content: result.success
                  ? `操作成功: ${result.content}`
                  : `操作失败: ${result.error}`,
              },
            ];

            // 继续对话，让LLM基于操作结果生成回复
            const followUpStream = await this.client.chat.completions.create({
              model: this.settings.model,
              messages: followUpMessages,
              temperature: 0.7,
              max_tokens: 1000,
              stream: true,
            });

            for await (const followUpChunk of followUpStream) {
              const followUpContent = followUpChunk.choices[0]?.delta?.content;
              if (followUpContent) {
                callbacks.onChunk(followUpContent);
              }
            }
          } catch (error) {
            console.error("Excel操作执行失败:", error);
            const errorMessage = error instanceof Error ? error.message : "未知错误";
            callbacks.onChunk(`\n\n执行Excel操作时发生错误: ${errorMessage}`);

            // 通知UI操作失败
            if (callbacks.onExcelAction) {
              callbacks.onExcelAction(
                `${currentFunctionCall.name}: 执行失败`,
                {
                  success: false,
                  content: "",
                  error: errorMessage
                }
              );
            }
          }

          // 重置状态
          functionCallBuffer = "";
          isCollectingFunctionCall = false;
          currentFunctionCall = {};
          break; // 工具调用处理完成
        }
      }

      callbacks.onComplete();
    } catch (error) {
      console.error("Enhanced OpenAI API Streaming Error:", error);
      const errorMessage = error instanceof Error ? this.getErrorMessage(error) : "API 调用失败";
      callbacks.onError(errorMessage);
    }
  }

  /**
   * 执行Excel操作
   * @param functionName 函数名称
   * @param args 参数
   * @returns 操作结果
   */
  private async executeExcelOperation(
    functionName: string,
    args: Record<string, any>
  ): Promise<ExcelActionResult> {
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
            selectAfterInsert: args.selectAfterInsert
          });
          break;

        case "setCellColor":
          result = await ExcelOperations.setCellColor(
            args.address,
            args.backgroundColor
          );
          break;

        case "setCellColorsBatch":
          result = await ExcelOperations.setCellColorsBatch({
            colorRanges: args.colorRanges
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
            numberFormat: args.numberFormat
          });
          break;

        case "createChart":
          result = await ExcelOperations.createChart({
            type: args.type,
            dataRange: args.dataRange,
            position: (args.left !== undefined && args.top !== undefined) ? {
              left: args.left,
              top: args.top,
              width: args.width,
              height: args.height
            } : undefined,
            style: {
              title: args.title,
              showLegend: args.showLegend,
              showDataLabels: args.showDataLabels,
              xAxisTitle: args.xAxisTitle,
              yAxisTitle: args.yAxisTitle
            }
          });
          break;

        case "insertFormula":
          result = await ExcelOperations.insertFormula({
            address: args.address,
            formula: args.formula,
            calculateImmediately: args.calculateImmediately
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
            numberFormat: args.numberFormat
          });
          break;

        default:
          throw new Error(`未知的Excel操作: ${functionName}`);
      }

      return {
        success: result.success,
        content: result.message,
        error: result.error
      };

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Excel操作执行失败";
      return {
        success: false,
        content: "",
        error: errorMessage
      };
    }
  }

  /**
   * 构建增强的系统提示，包含Excel操作信息
   */
  private getSystemPromptWithExcelOperations(): string {
    let systemPrompt = `你是 Egent，一个专门帮助用户分析和操作 Excel 数据的 AI 助手。

请遵循以下规则：
1. 专注于数据分析、解释和见解，同时可以直接执行Excel操作
2. 提供清晰、易懂的回答
3. 如果需要，可以建议数据可视化的方法，并直接创建图表
4. 始终以中文回答
5. 保持友好和专业的语气
6. 当需要执行Excel操作时，主动使用可用的工具函数`;

    if (this.enableExcelOperations) {
      systemPrompt += `

你可以直接操作Excel，包括以下功能：
- insertData: 插入数据到指定范围，支持表格数据、列表等
- setCellColor: 设置单个单元格或范围的背景色，用于突出显示重要信息
- setCellColorsBatch: 批量设置多个范围的背景色，一次操作设置多个不同区域的颜色
- setCellFont: 设置字体格式，包括颜色、大小、样式、对齐等
- createChart: 创建各种类型的图表，支持柱状图、折线图、饼图等
- insertFormula: 插入Excel公式，进行自动计算
- formatRange: 批量格式化单元格范围

使用建议：
- 在插入数据时，优先使用包含标题行的格式
- 创建图表时，选择最适合数据特征的图表类型
- 格式化时，注意颜色搭配和可读性
- 插入公式时，确保公式语法正确且以'='开头
- 需要设置多个不同范围颜色时，优先使用setCellColorsBatch批量操作

当用户需要Excel操作时，请：
1. 分析用户需求，确定合适的操作类型
2. 准备正确的参数（地址、数据、格式等）
3. 执行相应的工具函数
4. 向用户说明操作结果和效果`;
    }

    return systemPrompt;
  }

  /**
   * 构建增强的消息列表
   */
  private buildEnhancedMessages(
    userMessage: string,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = [],
    systemPrompt?: string
  ): OpenAI.ChatCompletionMessageParam[] {
    const messages: OpenAI.ChatCompletionMessageParam[] = [
      {
        role: "system",
        content: systemPrompt || this.getSystemPromptWithExcelOperations(),
      },
      // 转换历史消息格式
      ...conversationHistory.map(
        (msg): OpenAI.ChatCompletionMessageParam => ({
          role: msg.role === "assistant" ? "assistant" : "user",
          content: msg.content,
        })
      ),
    ];

    // 构建用户消息
    let content = userMessage;

    // 添加Excel数据上下文
    if (excelData) {
      content = `${userMessage}\n\n${this.formatExcelDataForAI(excelData)}`;
    }

    messages.push({
      role: "user",
      content,
    });

    return messages;
  }

  /**
   * 手动执行Excel操作（提供给UI直接调用）
   */
  async executeManualExcelOperation(operation: {
    type: "insertData" | "setCellColor" | "setCellColorsBatch" | "setCellFont" | "createChart" | "insertFormula" | "formatRange";
    parameters: Record<string, any>;
  }): Promise<ExcelActionResult> {
    try {
      return await this.executeExcelOperation(operation.type, operation.parameters);
    } catch (error) {
      return {
        success: false,
        content: "",
        error: error instanceof Error ? error.message : "操作执行失败",
      };
    }
  }

  /**
   * 获取当前Excel数据（用于AI分析）
   */
  async getCurrentExcelData(type: "selection" | "worksheet" | "range", address?: string): Promise<any> {
    try {
      let result: ExcelOperationResult;

      switch (type) {
        case "selection":
          result = await ExcelOperations.getCurrentSelection();
          break;
        case "worksheet":
          result = await ExcelOperations.getActiveWorksheetData();
          break;
        case "range":
          if (!address) {
            throw new Error("获取范围数据时必须提供地址参数");
          }
          result = await ExcelOperations.getRangeData(address);
          break;
        default:
          throw new Error(`不支持的数据获取类型: ${type}`);
      }

      if (!result.success) {
        throw new Error(result.error || "获取Excel数据失败");
      }

      return result.data;
    } catch (error) {
      console.error("获取Excel数据失败:", error);
      throw error;
    }
  }

  /**
   * 获取Excel操作状态信息
   */
  getExcelOperationsStatus(): {
    enabled: boolean;
    availableOperations: string[];
    description: string;
  } {
    const operations = [
      "insertData - 插入数据",
      "setCellColor - 设置背景色",
      "setCellColorsBatch - 批量设置背景色",
      "setCellFont - 设置字体格式",
      "createChart - 创建图表",
      "insertFormula - 插入公式",
      "formatRange - 格式化范围"
    ];

    return {
      enabled: this.enableExcelOperations,
      availableOperations: operations,
      description: this.enableExcelOperations
        ? "Excel操作功能已启用，可以直接在浏览器中执行Excel操作"
        : "Excel操作功能已禁用"
    };
  }

  /**
   * 兼容性方法：保持与之前版本的接口一致
   * @deprecated 请使用 sendMessageStreamWithExcel 方法
   */
  async sendMessageStreamWithMCP(
    userMessage: string,
    callbacks: EnhancedStreamCallback,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): Promise<void> {
    // 直接调用新的Excel方法，保持向后兼容
    return this.sendMessageStreamWithExcel(userMessage, callbacks, excelData, conversationHistory);
  }

  /**
   * 兼容性方法：获取状态信息
   * @deprecated 请使用 getExcelOperationsStatus 方法
   */
  async getMCPStatus() {
    const status = this.getExcelOperationsStatus();
    return {
      connected: status.enabled,
      clientType: "direct" as const,
      availableTools: status.availableOperations.map(op => op.split(' -')[0]),
      description: status.description
    };
  }
}