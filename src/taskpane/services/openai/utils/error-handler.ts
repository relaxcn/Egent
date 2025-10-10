/* global console */

/**
 * OpenAI 错误处理工具
 */

/**
 * 将错误转换为用户友好的错误消息
 * @param error 错误对象
 * @returns 友好的错误消息
 */
export function getErrorMessage(error: Error): string {
  const message = error.message;

  // 401 未授权
  if (message.includes("401") || message.includes("Unauthorized")) {
    return "API Key 无效，请检查配置";
  }

  // 429 速率限制
  if (message.includes("429") || message.includes("Rate limit")) {
    return "API 调用次数超限，请稍后重试";
  }

  // 网络错误
  if (message.includes("NetworkError") || message.includes("fetch")) {
    return "网络连接失败，请检查网络或 Base URL 配置";
  }

  // 模型错误
  if (message.includes("model")) {
    return "模型不存在或不可用，请检查模型配置";
  }

  // 超时错误
  if (message.includes("timeout")) {
    return "请求超时，请稍后重试";
  }

  // 默认返回原始错误消息
  return message;
}

/**
 * 记录错误日志
 * @param context 错误上下文
 * @param error 错误对象
 */
export function logError(context: string, error: unknown): void {
  console.error(`[OpenAI Error] ${context}:`, error);
}

/**
 * 判断是否为可重试的错误
 * @param error 错误对象
 * @returns 是否可重试
 */
export function isRetryableError(error: Error): boolean {
  const message = error.message;
  return (
    message.includes("429") ||
    message.includes("timeout") ||
    message.includes("NetworkError") ||
    message.includes("ECONNRESET")
  );
}
