/* global console */

/**
 * Function Calling 客户端
 * 支持 LLM 自动调用 Excel 操作的增强客户端
 */

import OpenAI from "openai";
import { BaseOpenAIClient } from "./base-client";
import { ApiSettings } from "../../utils/settings";
import { ExcelData } from "../../taskpane";
import {
  ChatMessage,
  EnhancedStreamCallback,
  ExcelActionResult,
  ChatMode,
  ToolCallState,
} from "./types";
import { PromptManager } from "./prompt-manager";
import { getAllExcelTools, validateToolCall } from "./tools/excel-tools";
import { ToolExecutor } from "./tools/tool-executor";
import { logError, getErrorMessage } from "./utils/error-handler";
import { formatToolCall } from "./utils/data-formatter";

/**
 * Function Calling 客户端类
 * 在基础客户端之上增加 Excel 操作能力
 */
export class FunctionCallingClient extends BaseOpenAIClient {
  private excelOperationsEnabled: boolean = true;

  constructor(settings: ApiSettings) {
    super(settings);
  }

  /**
   * 启用/禁用 Excel 操作功能
   */
  setExcelOperationsEnabled(enabled: boolean): void {
    this.excelOperationsEnabled = enabled;
    console.log(`🔧 Excel 操作功能已${enabled ? "启用" : "禁用"}`);
  }

  /**
   * 获取 Excel 操作工具定义
   */
  private getExcelTools(): OpenAI.ChatCompletionTool[] {
    return this.excelOperationsEnabled ? getAllExcelTools() : [];
  }

  /**
   * 增强的流式消息发送（支持 Function Calling）
   * @param userMessage 用户消息
   * @param callbacks 增强的流式回调
   * @param excelData Excel 数据（可选）
   * @param conversationHistory 对话历史（可选）
   */
  async sendMessageStreamWithFunctionCalling(
    userMessage: string,
    callbacks: EnhancedStreamCallback,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): Promise<void> {
    try {
      // 1. 获取可用的工具定义
      const tools = this.getExcelTools();

      // 2. 构建消息
      const messages = this.buildMessages({
        userMessage,
        excelData,
        conversationHistory,
        systemPrompt: PromptManager.getSystemPromptByMode(
          this.excelOperationsEnabled ? ChatMode.AGENT : ChatMode.ANALYSIS
        ),
      });

      console.log(`🔧 已加载 ${tools.length} 个 Excel 操作工具`);

      // 3. 调用 OpenAI API（包含 Function Calling）
      const stream = await this.client.chat.completions.create({
        model: this.settings.model,
        messages,
        temperature: 0.7,
        max_tokens: 1500,
        stream: true,
        tools: tools.length > 0 ? tools : undefined,
        tool_choice: tools.length > 0 ? "auto" : undefined,
      });

      // 4. 处理流式响应
      await this.processStream(stream, messages, callbacks);

      callbacks.onComplete();
    } catch (error) {
      logError("FunctionCallingClient.sendMessageStream", error);
      const errorMessage = error instanceof Error ? getErrorMessage(error) : "API 调用失败";
      callbacks.onError(errorMessage);
    }
  }

  /**
   * 处理流式响应
   */
  private async processStream(
    stream: AsyncIterable<OpenAI.ChatCompletionChunk>,
    originalMessages: OpenAI.ChatCompletionMessageParam[],
    callbacks: EnhancedStreamCallback
  ): Promise<void> {
    const toolCallState: ToolCallState = {
      arguments: "",
      isCollecting: false,
    };

    for await (const chunk of stream) {
      const choice = chunk.choices[0];
      if (!choice) continue;

      // 处理工具调用
      if (choice.delta.tool_calls && choice.delta.tool_calls.length > 0) {
        this.handleToolCallChunk(choice.delta.tool_calls[0], toolCallState);
      }
      // 处理普通文本内容
      else if (choice.delta.content) {
        callbacks.onChunk(choice.delta.content);
      }

      // 工具调用结束
      if (choice.finish_reason === "tool_calls" && toolCallState.isCollecting) {
        await this.executeToolCall(
          toolCallState,
          originalMessages,
          callbacks
        );
        break; // 工具调用处理完成
      }
    }
  }

  /**
   * 处理工具调用数据块
   */
  private handleToolCallChunk(
    toolCall: OpenAI.ChatCompletionChunk.Choice.Delta.ToolCall,
    state: ToolCallState
  ): void {
    if (toolCall.function) {
      if (toolCall.function.name) {
        state.name = toolCall.function.name;
        state.isCollecting = true;
        console.log(`🔧 LLM 请求执行 Excel 操作: ${toolCall.function.name}`);
      }

      if (toolCall.function.arguments) {
        state.arguments += toolCall.function.arguments;
      }
    }
  }

  /**
   * 执行工具调用
   */
  private async executeToolCall(
    toolCallState: ToolCallState,
    originalMessages: OpenAI.ChatCompletionMessageParam[],
    callbacks: EnhancedStreamCallback
  ): Promise<void> {
    try {
      const args = JSON.parse(toolCallState.arguments);
      console.log(`🔧 Excel 操作参数:`, args);
      console.log(formatToolCall(toolCallState.name!, args));

      // 验证函数调用参数
      const validation = validateToolCall(toolCallState.name!, args);
      if (!validation.valid) {
        throw new Error(`参数验证失败: ${validation.error}`);
      }

      // 执行 Excel 操作
      const result = await ToolExecutor.execute(toolCallState.name!, args);

      // 转换为 ExcelActionResult 格式
      const actionResult: ExcelActionResult = {
        success: result.success,
        content: result.message,
        error: result.error,
      };

      // 通知 UI 操作结果
      if (callbacks.onExcelAction) {
        callbacks.onExcelAction(
          `${toolCallState.name}: ${JSON.stringify(args)}`,
          actionResult
        );
      }

      console.log(result.success ? "✅" : "❌", result.message);

      // 将操作结果反馈给 LLM，让其继续生成回复
      await this.sendFollowUpResponse(
        originalMessages,
        toolCallState,
        result.success ? `操作成功: ${result.message}` : `操作失败: ${result.error}`,
        callbacks
      );
    } catch (error) {
      logError("FunctionCallingClient.executeToolCall", error);
      const errorMessage = error instanceof Error ? error.message : "未知错误";
      callbacks.onChunk(`\n\n执行 Excel 操作时发生错误: ${errorMessage}`);

      // 通知 UI 操作失败
      if (callbacks.onExcelAction) {
        callbacks.onExcelAction(`${toolCallState.name}: 执行失败`, {
          success: false,
          content: "",
          error: errorMessage,
        });
      }
    }
  }

  /**
   * 发送后续响应（将工具执行结果反馈给 LLM）
   */
  private async sendFollowUpResponse(
    originalMessages: OpenAI.ChatCompletionMessageParam[],
    toolCallState: ToolCallState,
    toolResult: string,
    callbacks: EnhancedStreamCallback
  ): Promise<void> {
    const toolCallId = "call_" + Date.now();

    const followUpMessages: OpenAI.ChatCompletionMessageParam[] = [
      ...originalMessages,
      {
        role: "assistant",
        tool_calls: [
          {
            id: toolCallId,
            type: "function",
            function: {
              name: toolCallState.name!,
              arguments: toolCallState.arguments,
            },
          },
        ],
      },
      {
        role: "tool",
        tool_call_id: toolCallId,
        content: toolResult,
      },
    ];

    // 继续对话，让 LLM 基于操作结果生成回复
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
  }

  /**
   * 手动执行 Excel 操作（提供给 UI 直接调用）
   */
  async executeManualExcelOperation(operation: {
    type: string;
    parameters: Record<string, any>;
  }): Promise<ExcelActionResult> {
    try {
      const result = await ToolExecutor.execute(operation.type, operation.parameters);
      return {
        success: result.success,
        content: result.message,
        error: result.error,
      };
    } catch (error) {
      return {
        success: false,
        content: "",
        error: error instanceof Error ? error.message : "操作执行失败",
      };
    }
  }

  /**
   * 获取 Excel 操作状态信息
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
      "formatRange - 格式化范围",
    ];

    return {
      enabled: this.excelOperationsEnabled,
      availableOperations: operations,
      description: this.excelOperationsEnabled
        ? "Excel 操作功能已启用，可以直接在浏览器中执行 Excel 操作"
        : "Excel 操作功能已禁用",
    };
  }
}
