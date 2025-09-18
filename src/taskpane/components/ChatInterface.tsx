import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Send24Regular, Table24Regular, Person24Regular, Bot24Regular, Copy24Regular, Delete24Regular } from "@fluentui/react-icons";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { ExcelData } from "../taskpane";

export interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  excelData?: ExcelData;
  isDeletable?: boolean;  // 是否可以删除
  dataId?: string;        // 数据的唯一标识
}

interface ChatInterfaceProps {
  messages: ChatMessage[];
  onSendMessage: (message: string, excelData?: ExcelData) => void;
  onDeleteData?: (dataId: string) => void;  // 删除数据的回调
  isLoading: boolean;
  excelData: ExcelData | null;
}

const ChatInterface: React.FC<ChatInterfaceProps> = ({
  messages,
  onSendMessage,
  onDeleteData,
  isLoading,
  excelData
}) => {
  const [inputValue, setInputValue] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (inputValue.trim() && !isLoading) {
      onSendMessage(inputValue.trim(), excelData || undefined);
      setInputValue("");
    }
  };

  const handleCopyMessage = async (content: string, messageId: string) => {
    try {
      await navigator.clipboard.writeText(content);
      setCopiedMessageId(messageId);
      setTimeout(() => setCopiedMessageId(null), 2000);
    } catch (err) {
      console.error('Failed to copy text: ', err);
    }
  };

  const formatExcelData = (data: ExcelData) => {
    const maxRowsToShow = 5;
    const maxColumnsToShow = 4;

    return (
      <div className="mt-3 p-4 bg-blue-50 border border-blue-200 rounded-lg">
        <div className="flex items-center gap-2 mb-3">
          <Table24Regular className="text-blue-600" />
          <span className="text-sm font-medium text-blue-800">
            Excel 数据 ({data.address})
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full text-xs border-collapse">
            {data.headers && (
              <thead>
                <tr className="border-b border-blue-200">
                  {data.headers.slice(0, maxColumnsToShow).map((header, index) => (
                    <th key={index} className="text-left py-2 px-3 font-medium text-blue-800 bg-blue-100">
                      {header}
                    </th>
                  ))}
                  {data.headers.length > maxColumnsToShow && (
                    <th className="text-left py-2 px-3 font-medium text-blue-800 bg-blue-100">...</th>
                  )}
                </tr>
              </thead>
            )}
            <tbody>
              {data.values.slice(0, maxRowsToShow).map((row, rowIndex) => (
                <tr key={rowIndex} className={`border-b border-blue-100 ${
                  rowIndex % 2 === 0 ? "bg-white" : "bg-blue-25"
                }`}>
                  {row.slice(0, maxColumnsToShow).map((cell, cellIndex) => (
                    <td key={cellIndex} className="py-2 px-3 text-gray-700">
                      {String(cell)}
                    </td>
                  ))}
                  {row.length > maxColumnsToShow && (
                    <td className="py-2 px-3 text-gray-500">...</td>
                  )}
                </tr>
              ))}
              {data.values.length > maxRowsToShow && (
                <tr>
                  <td colSpan={Math.min(data.values[0]?.length || 1, maxColumnsToShow + 1)}
                      className="py-2 px-3 text-center text-gray-500 text-xs bg-blue-50">
                    ... 还有 {data.values.length - maxRowsToShow} 行
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const renderMessage = (message: ChatMessage) => {
    const isUser = message.type === 'user';
    const isSystem = message.type === 'system';

    return (
      <div
        key={message.id}
        className={`flex gap-4 ${isUser ? 'flex-row-reverse' : 'flex-row'} mb-6`}
      >
        {/* Avatar */}
        <div className={`flex-shrink-0 w-10 h-10 rounded-full flex items-center justify-center shadow-sm ${
          isUser
            ? 'bg-gradient-to-br from-blue-500 to-blue-600 text-white'
            : isSystem
            ? 'bg-gradient-to-br from-yellow-400 to-yellow-500 text-white'
            : 'bg-gradient-to-br from-green-500 to-green-600 text-white'
        }`}>
          {isUser ? (
            <Person24Regular className="w-5 h-5" />
          ) : isSystem ? (
            <Table24Regular className="w-5 h-5" />
          ) : (
            <Bot24Regular className="w-5 h-5" />
          )}
        </div>

        {/* Message Content */}
        <div className={`flex-1 max-w-[75%] ${
          isUser ? 'text-right' : 'text-left'
        }`}>
          {/* Message Bubble */}
          <div className={`group relative inline-block max-w-full ${
            isUser
              ? 'bg-blue-600 text-white rounded-2xl rounded-br-md'
              : isSystem
              ? 'bg-yellow-100 text-yellow-800 border border-yellow-200 rounded-2xl'
              : 'bg-white text-gray-800 border border-gray-200 rounded-2xl rounded-bl-md shadow-sm'
          } px-4 py-3`}>

            {/* Copy Button for AI messages */}
            {!isUser && !isSystem && (
              <button
                onClick={() => handleCopyMessage(message.content, message.id)}
                className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity p-1 rounded hover:bg-gray-100"
                title="复制消息"
              >
                <Copy24Regular className={`w-4 h-4 ${
                  copiedMessageId === message.id ? 'text-green-600' : 'text-gray-500'
                }`} />
              </button>
            )}

            {/* Delete Button for deletable system messages */}
            {isSystem && message.isDeletable && onDeleteData && message.dataId && (
              <button
                onClick={() => onDeleteData(message.dataId!)}
                className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity p-1 rounded hover:bg-red-100"
                title="删除此数据"
              >
                <Delete24Regular className="w-4 h-4 text-red-500 hover:text-red-700" />
              </button>
            )}

            {/* Message Text */}
            <div className="text-sm leading-relaxed">
              {isUser || isSystem ? (
                // 用户和系统消息直接显示文本
                <div className="whitespace-pre-wrap">{message.content}</div>
              ) : (
                // AI消息使用 Markdown 渲染
                <div className="prose prose-sm max-w-none prose-headings:text-gray-800 prose-p:text-gray-700 prose-strong:text-gray-800 prose-code:text-blue-600 prose-code:bg-blue-50 prose-code:px-1 prose-code:py-0.5 prose-code:rounded prose-pre:bg-gray-50 prose-pre:border prose-blockquote:border-l-4 prose-blockquote:border-blue-400 prose-blockquote:bg-blue-50">
                  <ReactMarkdown
                    remarkPlugins={[remarkGfm]}
                    components={{
                      // 自定义代码块样式
                      code: ({ inline, className, children, ...props }: any) => {
                        if (inline) {
                          return (
                            <code className="bg-blue-100 text-blue-800 px-1 py-0.5 rounded text-xs font-mono" {...props}>
                              {children}
                            </code>
                          );
                        }
                        return (
                          <pre className="bg-gray-100 border border-gray-200 rounded-lg p-3 overflow-x-auto">
                            <code className="text-sm font-mono text-gray-800" {...props}>
                              {children}
                            </code>
                          </pre>
                        );
                      },
                      // 自定义表格样式
                      table: ({ children }) => (
                        <div className="overflow-x-auto my-3">
                          <table className="min-w-full border border-gray-200 rounded-lg">
                            {children}
                          </table>
                        </div>
                      ),
                      th: ({ children }) => (
                        <th className="bg-gray-50 border-b border-gray-200 px-3 py-2 text-left text-xs font-medium text-gray-700">
                          {children}
                        </th>
                      ),
                      td: ({ children }) => (
                        <td className="border-b border-gray-100 px-3 py-2 text-sm text-gray-600">
                          {children}
                        </td>
                      ),
                    }}
                  >
                    {message.content}
                  </ReactMarkdown>
                </div>
              )}
            </div>

            {/* Excel Data Display */}
            {message.excelData && formatExcelData(message.excelData)}
          </div>

          {/* Timestamp */}
          <div className={`text-xs text-gray-500 mt-2 ${
            isUser ? 'text-right' : 'text-left'
          }`}>
            {message.timestamp.toLocaleTimeString()}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="flex flex-col h-full bg-gradient-to-b from-gray-50 to-white">
      {/* Messages Area */}
      <div className="flex-1 overflow-y-auto p-6">
        {messages.length === 0 && (
          <div className="text-center text-gray-500 mt-16">
            <div className="bg-white rounded-full w-20 h-20 mx-auto mb-4 flex items-center justify-center shadow-lg">
              <Bot24Regular className="text-3xl text-blue-600" />
            </div>
            <h3 className="text-lg font-medium text-gray-700 mb-2">开始与 Excel Agent 对话</h3>
            <p className="text-sm">选择 Excel 数据后开始智能分析对话</p>
          </div>
        )}

        {messages.map(renderMessage)}

        {/* Loading Indicator */}
        {isLoading && (
          <div className="flex gap-4 mb-6">
            <div className="flex-shrink-0 w-10 h-10 rounded-full bg-gradient-to-br from-green-500 to-green-600 text-white flex items-center justify-center shadow-sm">
              <Bot24Regular className="w-5 h-5" />
            </div>
            <div className="flex-1">
              <div className="inline-block bg-white border border-gray-200 rounded-2xl rounded-bl-md shadow-sm px-4 py-3">
                <div className="flex items-center gap-3">
                  <div className="flex space-x-1">
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce"></div>
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce" style={{ animationDelay: '0.1s' }}></div>
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
                  </div>
                  <span className="text-sm text-gray-600">AI 正在分析中...</span>
                </div>
              </div>
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div className="border-t border-gray-200 bg-white p-4">
        {excelData && (
          <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Table24Regular className="text-blue-600" />
                <span className="text-sm text-blue-800 font-medium">
                  已选择数据: {excelData.address}
                </span>
                <span className="text-xs text-blue-600 bg-blue-100 px-2 py-1 rounded-full">
                  {excelData.values.length} 行
                </span>
              </div>
            </div>
          </div>
        )}

        <form onSubmit={handleSubmit} className="flex gap-3">
          <div className="flex-1 relative">
            <input
              type="text"
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              placeholder="输入消息..."
              disabled={isLoading}
              className="w-full px-4 py-3 pr-12 border border-gray-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-sm"
            />
          </div>
          <button
            type="submit"
            disabled={!inputValue.trim() || isLoading}
            className="px-6 py-3 bg-blue-600 text-white rounded-xl hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition-all shadow-sm flex items-center gap-2"
          >
            <Send24Regular className="w-4 h-4" />
            <span className="hidden sm:inline">发送</span>
          </button>
        </form>
      </div>
    </div>
  );
};

export default ChatInterface;