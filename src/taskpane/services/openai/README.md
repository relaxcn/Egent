# OpenAI 服务模块

本模块提供了与 OpenAI API 交互的完整功能，包括基础对话和 Function Calling 驱动的 Excel 操作。

## 📁 目录结构

```
openai/
├── base-client.ts              # 基础 OpenAI 客户端（只读分析）
├── function-calling-client.ts  # Function Calling 客户端（Excel 操作）
├── types.ts                    # 类型定义
├── prompt-manager.ts           # 提示词管理器
├── index.ts                    # 统一导出入口
├── tools/                      # Function Calling 工具
│   ├── excel-tools.ts         # Excel 工具定义和验证
│   └── tool-executor.ts       # 工具执行器
├── prompts/                    # 系统提示词（Markdown 格式）
│   ├── base-system.md         # 基础模式提示词
│   ├── agent-system.md        # Agent 模式提示词
│   └── excel-operations.md    # Excel 操作参考
└── utils/                      # 工具函数
    ├── error-handler.ts       # 错误处理
    └── data-formatter.ts      # 数据格式化
```

## 🎯 核心组件

### 1. BaseOpenAIClient（基础客户端）

**用途**：提供只读的数据分析功能

**特点**：
- ✅ 流式对话
- ✅ Excel 数据格式化
- ✅ API 连接测试
- ✅ 模型列表获取
- ❌ 不能操作 Excel

**使用示例**：
```typescript
import { BaseOpenAIClient } from './services/openai';

const client = new BaseOpenAIClient(settings);

await client.sendMessageStream(
  "分析这些数据",
  callbacks,
  excelData,
  conversationHistory
);
```

### 2. FunctionCallingClient（增强客户端）

**用途**：支持 LLM 自动调用 Excel 操作的智能客户端

**特点**：
- ✅ 继承基础客户端的所有功能
- ✅ OpenAI Function Calling 集成
- ✅ 7 种 Excel 操作（插入数据、设置颜色、创建图表等）
- ✅ 操作结果自动反馈给 LLM
- ✅ 可动态启用/禁用 Excel 操作

**使用示例**：
```typescript
import { FunctionCallingClient } from './services/openai';

const client = new FunctionCallingClient(settings);
client.setExcelOperationsEnabled(true); // 启用 Excel 操作

await client.sendMessageStreamWithFunctionCalling(
  "请将数据插入到 A1:C10，并创建柱状图",
  enhancedCallbacks,
  excelData,
  conversationHistory
);
```

### 3. PromptManager（提示词管理器）

**用途**：管理 Markdown 格式的系统提示词

**特点**：
- 📄 提示词以 Markdown 文件存储
- 🔄 自动加载和初始化
- 🔁 支持运行时覆盖
- 🎯 根据模式自动选择提示词

**使用示例**：
```typescript
import { PromptManager, ChatMode } from './services/openai';

// 获取提示词
const prompt = PromptManager.getSystemPromptByMode(ChatMode.AGENT);

// 自定义提示词
PromptManager.setCustomPrompt(ChatMode.AGENT, "自定义提示词内容");
```

### 4. ToolExecutor（工具执行器）

**用途**：执行 Function Calling 调用的 Excel 操作

**特点**：
- ⚙️ 统一的工具执行接口
- ✅ 完整的错误处理
- 📊 支持批量执行
- 🔍 详细的执行日志

**使用示例**：
```typescript
import { ToolExecutor } from './services/openai';

// 执行单个工具
const result = await ToolExecutor.execute('insertData', {
  address: 'A1',
  data: [['标题1', '标题2'], ['值1', '值2']]
});

// 批量执行
const results = await ToolExecutor.executeBatch([
  { functionName: 'insertData', args: {...} },
  { functionName: 'setCellColor', args: {...} }
]);
```

## 📝 可用的 Excel 操作

| 操作 | 功能 | 典型用途 |
|------|------|----------|
| `insertData` | 插入数据到指定范围 | 表格数据、列表 |
| `setCellColor` | 设置单个范围背景色 | 突出显示 |
| `setCellColorsBatch` | 批量设置多个范围背景色 | 颜色标记方案 |
| `setCellFont` | 设置字体格式 | 格式化标题 |
| `createChart` | 创建图表 | 数据可视化 |
| `insertFormula` | 插入公式 | 自动计算 |
| `formatRange` | 批量格式化 | 美化表格 |

详细的操作参数和使用说明，请参考 [prompts/excel-operations.md](./prompts/excel-operations.md)。

## 🔧 工具函数

### 错误处理（error-handler.ts）

```typescript
import { getErrorMessage, logError, isRetryableError } from './services/openai';

// 转换为友好的错误消息
const friendlyMessage = getErrorMessage(error);

// 记录错误日志
logError("操作上下文", error);

// 判断是否可重试
if (isRetryableError(error)) {
  // 执行重试逻辑
}
```

### 数据格式化（data-formatter.ts）

```typescript
import { formatExcelDataForAI, truncateText, formatToolCall } from './services/openai';

// 格式化 Excel 数据为 AI 可读格式
const formatted = formatExcelDataForAI(excelData);

// 截断过长的文本
const truncated = truncateText(longText, 1000);

// 格式化工具调用信息
const toolInfo = formatToolCall('insertData', args);
```

## 🚀 快速开始

### 1. 基础使用（只读分析）

```typescript
import { BaseOpenAIClient } from './services/openai';
import { getApiSettings } from './utils/settings';

// 创建客户端
const client = new BaseOpenAIClient(getApiSettings());

// 发送消息
await client.sendMessageStream(
  "分析这些销售数据",
  {
    onChunk: (chunk) => console.log(chunk),
    onComplete: () => console.log("完成"),
    onError: (error) => console.error(error)
  },
  excelData
);
```

### 2. 高级使用（Excel 操作）

```typescript
import { FunctionCallingClient } from './services/openai';

// 创建增强客户端
const client = new FunctionCallingClient(getApiSettings());
client.setExcelOperationsEnabled(true);

// 发送消息（支持 Excel 操作）
await client.sendMessageStreamWithFunctionCalling(
  "将销售数据插入 A1，并创建柱状图",
  {
    onChunk: (chunk) => console.log(chunk),
    onComplete: () => console.log("完成"),
    onError: (error) => console.error(error),
    onExcelAction: (action, result) => {
      console.log(`Excel 操作: ${action}`);
      console.log(`结果: ${result.success ? '成功' : '失败'}`);
    }
  },
  excelData
);
```

## 📚 类型定义

所有类型定义都在 [types.ts](./types.ts) 中，包括：

- `ChatMessage` - 聊天消息
- `StreamCallback` - 流式回调
- `EnhancedStreamCallback` - 增强的流式回调（支持 Excel 操作反馈）
- `ExcelActionResult` - Excel 操作结果
- `ServiceResponse` - 服务响应
- `ModelsResponse` - 模型列表响应
- `ChatMode` - 聊天模式枚举
- `ToolCallState` - 工具调用状态
- `OpenAIClientConfig` - 客户端配置

## 🔄 向后兼容

为了保持向后兼容，提供了以下别名：

```typescript
// 旧的类名仍然可用
import { OpenAIService, EnhancedOpenAIService } from './services/openai';

// 等价于
import { BaseOpenAIClient, FunctionCallingClient } from './services/openai';
```

## 🎨 自定义提示词

提示词以 Markdown 格式存储在 `prompts/` 目录中，你可以：

1. **直接编辑 Markdown 文件**（推荐）
   - 修改 `base-system.md` 或 `agent-system.md`
   - 重新构建项目

2. **运行时覆盖**
   ```typescript
   import { PromptManager, ChatMode } from './services/openai';

   PromptManager.setCustomPrompt(
     ChatMode.AGENT,
     "你是一个自定义的 AI 助手..."
   );
   ```

## 🛠️ 开发指南

### 添加新的 Excel 操作

1. 在 `excel/operations.ts` 中添加操作函数
2. 在 `excel/function-definitions.ts` 中添加 OpenAI Function 定义
3. 在 `tools/tool-executor.ts` 中添加执行逻辑
4. 更新 `prompts/agent-system.md` 中的操作说明

### 修改系统提示词

1. 编辑 `prompts/*.md` 文件
2. 运行 `npm run build` 重新构建
3. 提示词会自动加载到应用中

### 扩展错误处理

在 `utils/error-handler.ts` 中添加新的错误类型判断：

```typescript
export function getErrorMessage(error: Error): string {
  const message = error.message;

  // 添加新的错误类型
  if (message.includes("YOUR_ERROR_PATTERN")) {
    return "用户友好的错误消息";
  }

  // ... 其他错误类型
}
```

## 📖 相关文档

- [Excel 操作模块](../../excel/README.md)
- [提示词指南](./prompts/agent-system.md)
- [OpenAI Function Calling 文档](https://platform.openai.com/docs/guides/function-calling)

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT
