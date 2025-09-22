import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Send24Regular, Table24Regular, Person24Regular, Bot24Regular, Delete24Regular, Document24Regular, ChevronDown24Regular, ChevronUp24Regular } from "@fluentui/react-icons";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { ExcelData } from "../taskpane";
import "./markdown.css";

export interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  excelData?: ExcelData;
  attachedData?: ExcelData[];  // 本次消息关联的 Excel 数据
  isDeletable?: boolean;  // 是否可以删除
  dataId?: string;        // 数据的唯一标识
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

    const buttonClass = isUserMessage
      ? "bg-white/15 border-white/30 hover:bg-white/25 text-blue-50"
      : "bg-blue-50 border-blue-200 hover:bg-blue-100 text-blue-800";
    const iconColor = isUserMessage ? "text-blue-50" : "text-blue-600";
    const rangeItemClass = isUserMessage
      ? "bg-white/10 border-white/20 hover:bg-white/20 text-blue-50"
      : "bg-white border-blue-100 hover:bg-blue-50 text-blue-700";

    return (
      <div className="mt-4">
        {/* Collapsed state: single button showing data usage */}
        <button
          type="button"
          onClick={() => setCollapsedData(prev => ({
            ...prev,
            [dataKey]: !prev[dataKey]
          }))}
          className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border transition-all shadow-sm ${buttonClass}`}
          title={isExpanded ? "折叠数据引用" : "展开查看使用的数据"}
        >
          <div className="flex items-center gap-3">
            <div className={`flex-shrink-0 ${iconColor}`}>
              {getDataTypeIcon('excel')}
            </div>
            <div className="text-sm font-medium">
              使用了 {dataList.length} 个数据源
            </div>
          </div>
          <div className={`${iconColor}`}>
            {isExpanded ?
              <ChevronUp24Regular className="w-4 h-4" /> :
              <ChevronDown24Regular className="w-4 h-4" />
            }
          </div>
        </button>

        {/* Expanded state: list of data ranges */}
        {isExpanded && (
          <div className="mt-2 space-y-1 pl-4">
            {dataList.map((data, index) => (
              <button
                key={`${messageId}-${data.id ?? index}`}
                type="button"
                onClick={() => onSelectExcelRange?.(data.address)}
                className={`w-full flex items-center gap-3 px-3 py-2 rounded-lg border transition-all text-left ${rangeItemClass}`}
                title={`点击在Excel中高亮选中 ${data.address}`}
              >
                <div className={`flex-shrink-0 ${iconColor}`}>
                  {getDataTypeIcon('excel')}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-medium truncate">{data.address}</div>
                  <div className="text-xs opacity-75">
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
        <div key={message.id} style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', marginBottom: '16px' }}>
          {/* 蓝色气泡 */}
          <div style={{
            maxWidth: '75%',
            backgroundColor: '#3b82f6',
            color: 'white',
            padding: '12px 16px',
            borderRadius: '16px 16px 4px 16px',
            boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)',
          }}>
            <div style={{ fontSize: '14px', whiteSpace: 'pre-wrap' }}>{message.content}</div>
          </div>
          {/* 引用的数据 - 在气泡外面，紧跟下方 */}
          {message.attachedData &&
            message.attachedData.length > 0 && (
              <div style={{ marginTop: '4px', maxWidth: '75%' }}>
                {renderAttachedDataList(message.attachedData, message.id, isUser)}
              </div>
            )}
        </div>
      );
    } else if (isSystem) {
      // 系统消息：黄色背景
      return (
        <div key={message.id} className="flex justify-center mb-4">
          <div className="max-w-[85%] bg-yellow-50 text-yellow-800 border border-yellow-200 px-4 py-3 rounded-lg text-center">
            <div className="text-sm">{message.content}</div>
          </div>
        </div>
      );
    } else {
      // AI回复：无气泡，左对齐，纯文本
      return (
        <div key={message.id} className="flex justify-start mb-6">
          <div className="max-w-[85%] text-gray-800">
            <div className="markdown-content">
              <Markdown remarkPlugins={[remarkGfm]}>{message.content}</Markdown>
            </div>
            {/* Excel Data Display */}
            {message.excelData && formatExcelData(message.excelData)}
          </div>
        </div>
      );
    }
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
        {/* Selected Data List */}
        {selectedDataList.length > 0 && (
          <div className="mb-4">
            <div className="text-sm text-gray-600 mb-2 font-medium">
              已选中数据 ({selectedDataList.length} 个)：
            </div>
            <div className="space-y-2 max-h-32 overflow-y-auto">
              {selectedDataList.map((data) => (
                <div
                  key={data.id}
                  className="flex items-center justify-between p-3 bg-blue-50 border border-blue-200 rounded-lg hover:bg-blue-100 transition-colors cursor-pointer"
                  onClick={() => onSelectExcelRange?.(data.address)}
                  title={`点击在Excel中选中范围 ${data.address}`}
                >
                  <div className="flex items-center gap-2 flex-1 min-w-0">
                    <Table24Regular className="text-blue-600 flex-shrink-0" />
                    <div className="flex-1 min-w-0">
                      <div className="text-sm text-blue-800 font-medium truncate">
                        {data.address}
                      </div>
                      <div className="text-xs text-blue-600">
                        {data.values.length} 行
                        {data.headers && ` • 包含标题`}
                        {` • ${data.timestamp.toLocaleTimeString()}`}
                      </div>
                    </div>
                  </div>
                  {onRemoveSelectedData && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation(); // 防止触发父元素的点击事件
                        onRemoveSelectedData(data.id);
                      }}
                      className="p-1 rounded hover:bg-red-100 transition-colors flex-shrink-0"
                      title="移除此数据"
                    >
                      <Delete24Regular className="w-4 h-4 text-red-500 hover:text-red-700" />
                    </button>
                  )}
                </div>
              ))}
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
