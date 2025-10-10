/**
 * OpenAI 服务统一导出
 * 提供清晰的模块化结构和访问路径
 */

// =============================================================================
// 核心客户端类
// =============================================================================
export { BaseOpenAIClient } from "./base-client";
export { FunctionCallingClient } from "./function-calling-client";

// =============================================================================
// 类型定义
// =============================================================================
export type {
  ChatMessage,
  StreamCallback,
  EnhancedStreamCallback,
  ExcelActionResult,
  ServiceResponse,
  ModelsResponse,
  MessageBuildOptions,
  ToolCallState,
  OpenAIClientConfig,
} from "./types";

export { ChatMode } from "./types";

// =============================================================================
// 提示词管理
// =============================================================================
export { PromptManager } from "./prompt-manager";

// =============================================================================
// 工具相关
// =============================================================================
export {
  getAllExcelTools,
  getExcelToolByName,
  validateToolCall,
} from "./tools/excel-tools";

export { ToolExecutor } from "./tools/tool-executor";

export type { ToolExecutionResult } from "./tools/excel-tools";

// =============================================================================
// 工具函数
// =============================================================================
export {
  getErrorMessage,
  logError,
  isRetryableError,
} from "./utils/error-handler";

export {
  formatExcelDataForAI,
  truncateText,
  formatToolCall,
} from "./utils/data-formatter";

// =============================================================================
// 便捷别名（向后兼容）
// =============================================================================

/**
 * OpenAIService - 基础服务类（只读分析）
 * @deprecated 请使用 BaseOpenAIClient
 */
export { BaseOpenAIClient as OpenAIService } from "./base-client";

/**
 * EnhancedOpenAIService - 增强服务类（支持 Excel 操作）
 * @deprecated 请使用 FunctionCallingClient
 */
export { FunctionCallingClient as EnhancedOpenAIService } from "./function-calling-client";

// =============================================================================
// 默认导出
// =============================================================================
import { BaseOpenAIClient } from "./base-client";
import { FunctionCallingClient } from "./function-calling-client";
import { PromptManager } from "./prompt-manager";
import { ToolExecutor } from "./tools/tool-executor";
import { ChatMode } from "./types";

export default {
  BaseOpenAIClient,
  FunctionCallingClient,
  PromptManager,
  ToolExecutor,
  ChatMode,
};
