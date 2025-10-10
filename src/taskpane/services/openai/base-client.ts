/* global console */

/**
 * 基础 OpenAI 客户端
 * 提供只读分析功能，不执行 Excel 操作
 */

import OpenAI from "openai";
import { ApiSettings } from "../../utils/settings";
import { ExcelData } from "../../taskpane";
import {
  ChatMessage,
  StreamCallback,
  ServiceResponse,
  ModelsResponse,
  MessageBuildOptions,
} from "./types";
import { getErrorMessage, logError } from "./utils/error-handler";
import { formatExcelDataForAI } from "./utils/data-formatter";
import { PromptManager } from "./prompt-manager";

/**
 * 基础 OpenAI 客户端类
 * 专注于数据分析和对话，不执行 Excel 操作
 */
export class BaseOpenAIClient {
  protected client: OpenAI;
  protected settings: ApiSettings;

  constructor(settings: ApiSettings) {
    this.settings = settings;
    this.client = new OpenAI({
      apiKey: settings.apiKey,
      baseURL: settings.baseUrl,
      dangerouslyAllowBrowser: true, // 允许在浏览器环境中使用
    });
  }

  /**
   * 构建消息列表
   */
  protected buildMessages(options: MessageBuildOptions): OpenAI.ChatCompletionMessageParam[] {
    const {
      userMessage,
      excelData,
      conversationHistory = [],
      systemPrompt,
    } = options;

    const messages: OpenAI.ChatCompletionMessageParam[] = [
      {
        role: "system",
        content: systemPrompt || PromptManager.getBaseSystemPrompt(),
      },
      // 转换历史消息
      ...conversationHistory.map(
        (msg): OpenAI.ChatCompletionMessageParam => ({
          role: msg.role,
          content: msg.content,
        })
      ),
    ];

    // 构建用户消息（包含 Excel 数据）
    let content = userMessage;
    if (excelData) {
      content = `${userMessage}\n\n${formatExcelDataForAI(excelData)}`;
    }

    messages.push({
      role: "user",
      content,
    });

    return messages;
  }

  /**
   * 发送流式消息
   * @param userMessage 用户消息
   * @param callbacks 流式回调
   * @param excelData Excel 数据（可选）
   * @param conversationHistory 对话历史（可选）
   */
  async sendMessageStream(
    userMessage: string,
    callbacks: StreamCallback,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): Promise<void> {
    try {
      const messages = this.buildMessages({
        userMessage,
        excelData,
        conversationHistory,
      });

      const stream = await this.client.chat.completions.create({
        model: this.settings.model,
        messages,
        temperature: 0.7,
        max_tokens: 1000,
        stream: true,
      });

      for await (const chunk of stream) {
        const content = chunk.choices[0]?.delta?.content;
        if (content) {
          callbacks.onChunk(content);
        }
      }

      callbacks.onComplete();
    } catch (error) {
      logError("BaseOpenAIClient.sendMessageStream", error);
      const errorMessage = error instanceof Error ? getErrorMessage(error) : "API 调用失败";
      callbacks.onError(errorMessage);
    }
  }

  /**
   * 测试 API 连接
   */
  async testConnection(): Promise<ServiceResponse> {
    try {
      const models = await this.client.models.list();

      if (models.data && models.data.length > 0) {
        return {
          success: true,
          message: `API 连接测试成功，发现 ${models.data.length} 个可用模型`,
        };
      } else {
        return {
          success: true,
          message: "API 连接成功，但未发现可用模型",
        };
      }
    } catch (error) {
      logError("BaseOpenAIClient.testConnection", error);
      return {
        success: false,
        error: error instanceof Error ? getErrorMessage(error) : "连接测试失败",
      };
    }
  }

  /**
   * 获取可用模型列表
   */
  async getAvailableModels(): Promise<ModelsResponse> {
    try {
      const response = await this.client.models.list();
      const modelNames = response.data.map((model) => model.id).sort();

      return {
        success: true,
        models: modelNames,
      };
    } catch (error) {
      logError("BaseOpenAIClient.getAvailableModels", error);
      return {
        success: false,
        error: error instanceof Error ? getErrorMessage(error) : "获取模型列表失败",
      };
    }
  }

  /**
   * 更新设置
   * @param newSettings 新的设置
   */
  updateSettings(newSettings: ApiSettings): void {
    this.settings = newSettings;
    this.client = new OpenAI({
      apiKey: newSettings.apiKey,
      baseURL: newSettings.baseUrl,
      dangerouslyAllowBrowser: true,
    });
  }

  /**
   * 获取当前设置
   */
  getSettings(): ApiSettings {
    return { ...this.settings };
  }
}
