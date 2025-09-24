/* global console */

import OpenAI from "openai";
import { ApiSettings } from "../utils/settings";
import { ExcelData } from "../taskpane";

export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
}

export interface StreamCallback {
  onChunk: (chunk: string) => void;
  onComplete: () => void;
  onError: (error: string) => void;
}

export interface ServiceResponse {
  success: boolean;
  message?: string;
  error?: string;
}

export interface ModelsResponse {
  success: boolean;
  models?: string[];
  error?: string;
}

const SYSTEM_PROMPT = `你是 Egent，一个专门帮助用户分析 Excel 数据的 AI 助手。你只能查看和分析数据，不能修改 Excel 文件。

请遵循以下规则：
1. 专注于数据分析、解释和见解
2. 提供清晰、易懂的回答
3. 如果需要，可以建议数据可视化的方法
4. 始终以中文回答
5. 保持友好和专业的语气

如果用户提供了 Excel 数据，请基于该数据进行分析和回答。`;

export class OpenAIService {
  protected client: OpenAI;
  protected settings: ApiSettings;

  constructor(settings: ApiSettings) {
    this.settings = settings;
    this.client = new OpenAI({
      apiKey: settings.apiKey,
      baseURL: settings.baseUrl,
      // 在浏览器环境中禁用默认的用户代理
      dangerouslyAllowBrowser: true,
    });
  }

  protected getErrorMessage(error: Error): string {
    if (error.message.includes("401") || error.message.includes("Unauthorized")) {
      return "API Key 无效，请检查配置";
    } else if (error.message.includes("429") || error.message.includes("Rate limit")) {
      return "API 调用次数超限，请稍后重试";
    } else if (error.message.includes("NetworkError") || error.message.includes("fetch")) {
      return "网络连接失败，请检查网络或 Base URL 配置";
    } else if (error.message.includes("model")) {
      return "模型不存在或不可用，请检查模型配置";
    } else {
      return error.message;
    }
  }

  protected formatExcelDataForAI(dataList: ExcelData | ExcelData[]): string {
    const dataArray = Array.isArray(dataList) ? dataList : [dataList];

    if (dataArray.length === 0) return "";

    let formatted =
      dataArray.length === 1 ? `Excel 数据:\n` : `Excel 数据 (${dataArray.length} 个数据集):\n`;

    dataArray.forEach((data, index) => {
      if (dataArray.length > 1) {
        formatted += `\n### 数据集 ${index + 1}: ${data.address}\n`;
      } else {
        formatted += `数据来源: ${data.address}\n`;
      }

      if (data.headers) {
        formatted += `列标题: ${data.headers.join(", ")}\n`;
        formatted += `数据 (${data.values.length} 行):\n`;

        // 格式化为表格
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
      } else {
        formatted += `数据 (${data.values.length} 行 x ${data.values[0]?.length || 0} 列):\n`;
        data.values.slice(0, 10).forEach((row, rowIndex) => {
          formatted += `行${rowIndex + 1}: [${row.join(", ")}]\n`;
        });

        if (data.values.length > 10) {
          formatted += `... (还有 ${data.values.length - 10} 行数据)\n`;
        }
      }

      if (dataArray.length > 1 && index < dataArray.length - 1) {
        formatted += "\n";
      }
    });

    return formatted;
  }

  private buildMessages(
    userMessage: string,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): OpenAI.ChatCompletionMessageParam[] {
    const messages: OpenAI.ChatCompletionMessageParam[] = [
      {
        role: "system",
        content: SYSTEM_PROMPT,
      },
      // 转换历史消息格式
      ...conversationHistory.map(
        (msg): OpenAI.ChatCompletionMessageParam => ({
          role: msg.role,
          content: msg.content,
        })
      ),
    ];

    // 构建用户消息
    let content = userMessage;
    if (excelData) {
      content = `${userMessage}\n\n${this.formatExcelDataForAI(excelData)}`;
    }

    messages.push({
      role: "user",
      content,
    });

    return messages;
  }

  async sendMessageStream(
    userMessage: string,
    callbacks: StreamCallback,
    excelData?: ExcelData | ExcelData[],
    conversationHistory: ChatMessage[] = []
  ): Promise<void> {
    try {
      const messages = this.buildMessages(userMessage, excelData, conversationHistory);

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
      console.error("OpenAI API Streaming Error:", error);
      const errorMessage = error instanceof Error ? this.getErrorMessage(error) : "API 调用失败";
      callbacks.onError(errorMessage);
    }
  }

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
      console.error("Connection test error:", error);
      return {
        success: false,
        error: error instanceof Error ? this.getErrorMessage(error) : "连接测试失败",
      };
    }
  }

  async getAvailableModels(): Promise<ModelsResponse> {
    try {
      const response = await this.client.models.list();
      const modelNames = response.data.map((model) => model.id).sort();

      return {
        success: true,
        models: modelNames,
      };
    } catch (error) {
      console.error("Get models error:", error);
      return {
        success: false,
        error: error instanceof Error ? this.getErrorMessage(error) : "获取模型列表失败",
      };
    }
  }
}
