import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Send24Regular, Table24Regular, Person24Regular, Bot24Regular, Delete24Regular, Document24Regular, ChevronDown24Regular, ChevronUp24Regular } from "@fluentui/react-icons";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { ExcelData } from "../taskpane";

export interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  excelData?: ExcelData;
  attachedData?: ExcelData[];  // 本次消息关联的 Excel 数据
  isDeletable?: boolean;  // 是否可以删除
  dataId?: string;        // 数据的唯一标识
  isStreaming?: boolean;  // 是否正在流式输出
}

interface ChatInterfaceProps {
  messages: ChatMessage[];
  onSendMessage: (message: string) => void; // 简化，不再需要传递单个数据
  onRemoveSelectedData?: (dataId: string) => void; // 从选中列表移除数据的回调
  onSelectExcelRange?: (address: string) => void; // 选中Excel范围的回调
  isLoading: boolean;
  selectedDataList: ExcelData[]; // 改为数组
}

const ChatInterface: React.FC<ChatInterfaceProps> = ({
  messages,
  onSendMessage,
  onRemoveSelectedData,
  onSelectExcelRange,
  isLoading,
  selectedDataList
}) => {
  const [inputValue, setInputValue] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const [collapsedData, setCollapsedData] = useState<{[key: string]: boolean}>({});

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (inputValue.trim() && !isLoading) {
      onSendMessage(inputValue.trim());
      setInputValue("");
    }
  };

  const getDataTypeIcon = (dataType: string) => {
    switch (dataType) {
      case 'excel':
        return <Table24Regular className="w-4 h-4" />;
      case 'file':
        return <Document24Regular className="w-4 h-4" />;
      default:
        return <Table24Regular className="w-4 h-4" />;
    }
  };

  const formatExcelData = (data: ExcelData) => {
    const maxRowsToShow = 5;
    const maxColumnsToShow = 4;

    return (
      <div className="data-card">
        <div className="flex items-center gap-2 mb-3">
          <Table24Regular className="text-blue-600" />
          <span className="text-sm font-medium text-blue-800">
            Excel 数据 ({data.address})
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="data-table">
            {data.headers && (
              <thead>
                <tr className="border-b border-blue-200">
                  {data.headers.slice(0, maxColumnsToShow).map((header, index) => (
                    <th key={index} className="data-table-header">
                      {header}
                    </th>
                  ))}
                  {data.headers.length > maxColumnsToShow && (
                    <th className="data-table-header">...</th>
                  )}
                </tr>
              </thead>
            )}
            <tbody>
              {data.values.slice(0, maxRowsToShow).map((row, rowIndex) => (
                <tr key={rowIndex} className={rowIndex % 2 === 0 ? "data-table-row-even" : "data-table-row-odd"}>
                  {row.slice(0, maxColumnsToShow).map((cell, cellIndex) => (
                    <td key={cellIndex} className="data-table-cell">
                      {String(cell)}
                    </td>
                  ))}
                  {row.length > maxColumnsToShow && (
                    <td className="data-table-cell text-gray-500">...</td>
                  )}
                </tr>
              ))}
              {data.values.length > maxRowsToShow && (
                <tr>
                  <td colSpan={Math.min(data.values[0]?.length || 1, maxColumnsToShow + 1)}
                      className="data-table-cell text-center text-gray-500 text-xs bg-blue-50">
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

  const renderAttachedDataList = (
    dataList: ExcelData[],
    messageId: string,
    isUserMessage: boolean
  ) => {
    if (!dataList || dataList.length === 0) {
      return null;
    }

    const dataKey = `${messageId}-data`;
    const isExpanded = collapsedData[dataKey] ?? false;
    const variantClass = isUserMessage ? "data-attachments-user" : "data-attachments-assistant";

    return (
      <div className={`data-attachments ${variantClass}`}>
        <button
          type="button"
          onClick={() => setCollapsedData(prev => ({
            ...prev,
            [dataKey]: !prev[dataKey]
          }))}
          className="data-attachments-toggle"
          title={isExpanded ? "折叠数据引用" : "展开查看使用的数据"}
        >
          <div className="data-attachments-toggle-main">
            <div className="data-attachments-icon">
              {getDataTypeIcon('excel')}
            </div>
            <div className="data-attachments-label">
              使用了 {dataList.length} 个数据源
            </div>
          </div>
          <div className="data-attachments-chevron">
            {isExpanded ?
              <ChevronUp24Regular className="data-attachments-chevron-icon" /> :
              <ChevronDown24Regular className="data-attachments-chevron-icon" />
            }
          </div>
        </button>

        {isExpanded && (
          <div className="data-attachments-list">
            {dataList.map((data, index) => (
              <button
                key={`${messageId}-${data.id ?? index}`}
                type="button"
                onClick={() => onSelectExcelRange?.(data.address)}
                className="data-attachments-item"
                title={`点击在Excel中高亮选中 ${data.address}`}
              >
                <div className="data-attachments-icon">
                  {getDataTypeIcon('excel')}
                </div>
                <div className="data-attachments-item-text">
                  <div className="data-attachments-item-title">{data.address}</div>
                  <div className="data-attachments-item-meta">
                    {`${data.values.length} 行`}
                    {data.headers ? " · 含标题" : ""}
                  </div>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    );
  };


  const renderMessage = (message: ChatMessage) => {
    const isUser = message.type === 'user';
    const isSystem = message.type === 'system';

    if (isUser) {
      // 用户消息：蓝色气泡，右对齐
      return (
        <div key={message.id} className="message-user slide-in-right">
          <div className="message-user-content">
            <div className="message-bubble message-bubble-user">
              <div className="text-sm whitespace-pre-wrap">{message.content}</div>
            </div>
            {message.attachedData &&
              message.attachedData.length > 0 && (
                <div className="message-user-attachments">
                  {renderAttachedDataList(message.attachedData, message.id, isUser)}
                </div>
              )}
          </div>
        </div>
      );
    } else if (isSystem) {

      // 系统消息：黄色背景
      return (
        <div key={message.id} className="message-system fade-in">
          <div className="message-bubble message-bubble-system text-center">
            <div className="text-sm">{message.content}</div>
          </div>
        </div>
      );
    } else {
      // AI回复：无气泡，左对齐，纯文本
      return (
        <div key={message.id} className="message-assistant slide-in-left">
          <div className="message-bubble message-bubble-assistant">
            <div className="markdown-content">
              <Markdown remarkPlugins={[remarkGfm]}>
                {message.content}
              </Markdown>
              {message.isStreaming && (
                <span className="streaming-cursor">▋</span>
              )}
            </div>
            {/* Excel Data Display */}
            {message.excelData && formatExcelData(message.excelData)}
          </div>
        </div>
      );
    }
  };

  return (
    <div className="chat-container">
      {/* Messages Area */}
      <div className="chat-messages">
        {messages.length === 0 && (
          <div className="text-center text-gray-500 mt-16 fade-in">
            <div className="card w-20 h-20 mx-auto mb-4 flex items-center justify-center">
              <Bot24Regular className="text-3xl text-blue-600" />
            </div>
            <h3 className="welcome-title text-lg mb-2">开始与 Excel Agent 对话</h3>
            <p className="welcome-description">选择 Excel 数据后开始智能分析对话</p>
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
      <div className="chat-input-container">
        {/* Selected Data List */}
        {selectedDataList.length > 0 && (
          <div className="mb-6">
            <div className="label flex items-center gap-2">
              <div className="w-2 h-2 bg-blue-500 rounded-full pulse"></div>
              已选中数据 ({selectedDataList.length} 个)
            </div>
            <div className="space-y-3 max-h-36 overflow-y-auto">
              {selectedDataList.map((data) => (
                <div
                  key={data.id}
                  className="data-card group flex items-center justify-between hover-lift cursor-pointer"
                  onClick={() => onSelectExcelRange?.(data.address)}
                  title={`点击在Excel中选中范围 ${data.address}`}
                >
                  <div className="flex items-center gap-3 flex-1 min-w-0">
                    <div className="flex-shrink-0 w-10 h-10 bg-gradient-to-br from-blue-500 to-indigo-600 rounded-xl flex items-center justify-center text-white shadow-md hover-glow">
                      <Table24Regular className="w-5 h-5" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm text-blue-900 font-semibold truncate">
                        {data.address}
                      </div>
                      <div className="text-xs text-blue-700/80 flex items-center gap-2">
                        <span>{data.values.length} 行</span>
                        {data.headers && <span className="flex items-center gap-1">• 包含标题</span>}
                        <span>• {data.timestamp.toLocaleTimeString()}</span>
                      </div>
                    </div>
                  </div>
                  {onRemoveSelectedData && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onRemoveSelectedData(data.id);
                      }}
                      className="button button-ghost p-2 hover:bg-red-100 hover-scale flex-shrink-0 opacity-70 hover:opacity-100"
                      title="移除此数据"
                    >
                      <Delete24Regular className="w-4 h-4 text-red-500 hover:text-red-600" />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}

        <form onSubmit={handleSubmit} className="relative">
          <div className="flex gap-4 items-end">
            <div className="flex-1 relative">
              <div className="absolute inset-0 bg-gradient-to-r from-blue-500/10 to-indigo-500/10 rounded-2xl blur-sm"></div>
              <div className="relative">
                <input
                  type="text"
                  value={inputValue}
                  onChange={(e) => setInputValue(e.target.value)}
                  placeholder="输入您的问题或指令..."
                  disabled={isLoading}
                  className="input px-6 py-4 pr-14 rounded-2xl bg-white/90 backdrop-blur-sm text-base"
                  style={{
                    minHeight: '56px',
                    boxShadow: '0 4px 20px rgba(59, 130, 246, 0.15), 0 1px 3px rgba(0, 0, 0, 0.1)'
                  }}
                />
                {/* 输入框内的装饰元素 */}
                <div className="absolute right-4 top-1/2 transform -translate-y-1/2 flex items-center gap-2">
                  {isLoading && (
                    <div className="flex space-x-1">
                      <div className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce"></div>
                      <div className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce" style={{ animationDelay: '0.1s' }}></div>
                      <div className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
                    </div>
                  )}
                </div>
              </div>
            </div>

            <button
              type="submit"
              disabled={!inputValue.trim() || isLoading}
              className="button button-primary group relative px-6 py-4 rounded-2xl flex items-center gap-3 disabled:hover:translate-y-0"
              style={{
                minHeight: '56px',
                background: inputValue.trim() && !isLoading
                  ? 'linear-gradient(135deg, #3b82f6 0%, #6366f1 100%)'
                  : undefined
              }}
            >
              <div className="absolute inset-0 bg-gradient-to-r from-white/20 to-transparent rounded-2xl opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
              <Send24Regular className="w-5 h-5 relative z-10 hover-scale" />
              <span className="hidden sm:inline relative z-10">发送</span>
              {inputValue.trim() && !isLoading && (
                <div className="absolute -inset-1 bg-gradient-to-r from-blue-600 to-indigo-600 rounded-2xl blur opacity-30 group-hover:opacity-50 transition-opacity duration-300"></div>
              )}
            </button>
          </div>

          {/* 底部提示文字 */}
          <div className="mt-3 text-xs text-gray-500 text-center">
            <span className="inline-flex items-center gap-1">
              按 Enter 发送消息
              <kbd className="px-2 py-1 bg-gray-100 border border-gray-200 rounded text-xs">⏎</kbd>
            </span>
          </div>
        </form>
      </div>
    </div>
  );
};

export default ChatInterface;
