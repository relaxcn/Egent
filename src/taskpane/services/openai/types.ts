/* global console */

/**
 * OpenAI 服务相关类型定义
 */

import { ExcelData } from "../../taskpane";

/**
 * 聊天消息接口
 */
export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
}

/**
 * 流式回调接口
 */
export interface StreamCallback {
  onChunk: (chunk: string) => void;
  onComplete: () => void;
  onError: (error: string) => void;
}

/**
 * Excel 操作结果接口
 */
export interface ExcelActionResult {
  success: boolean;
  content: string;
  error?: string;
}

/**
 * 增强的流式回调接口（支持 Excel 操作反馈）
 */
export interface EnhancedStreamCallback extends StreamCallback {
  onExcelAction?: (action: string, result: ExcelActionResult) => void;
}

/**
 * 服务响应接口
 */
export interface ServiceResponse {
  success: boolean;
  message?: string;
  error?: string;
}

/**
 * 模型列表响应接口
 */
export interface ModelsResponse {
  success: boolean;
  models?: string[];
  error?: string;
}

/**
 * 聊天模式枚举
 */
export enum ChatMode {
  /** 只读分析模式 */
  ANALYSIS = "analysis",
  /** Agent 操作模式 */
  AGENT = "agent",
}

/**
 * 消息构建选项
 */
export interface MessageBuildOptions {
  userMessage: string;
  excelData?: ExcelData | ExcelData[];
  conversationHistory?: ChatMessage[];
  systemPrompt?: string;
}

/**
 * 工具调用状态
 */
export interface ToolCallState {
  name?: string;
  arguments: string;
  isCollecting: boolean;
}

/**
 * OpenAI 客户端配置
 */
export interface OpenAIClientConfig {
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
}
