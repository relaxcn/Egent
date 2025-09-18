import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Send24Regular, Table24Regular, Person24Regular, Bot24Regular } from "@fluentui/react-icons";
import { ExcelData } from "../taskpane";

export interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  excelData?: ExcelData;
}

interface ChatInterfaceProps {
  messages: ChatMessage[];
  onSendMessage: (message: string, excelData?: ExcelData) => void;
  isLoading: boolean;
  excelData: ExcelData | null;
}

const ChatInterface: React.FC<ChatInterfaceProps> = ({
  messages,
  onSendMessage,
  isLoading,
  excelData
}) => {
  const [inputValue, setInputValue] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);

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

  const formatExcelData = (data: ExcelData) => {
    const maxRowsToShow = 5;
    const maxColumnsToShow = 4;

    return (
      <div className="mt-2 p-3 bg-blue-50 border border-blue-200 rounded-lg">
        <div className="flex items-center gap-2 mb-2">
          <Table24Regular className="text-blue-600" />
          <span className="text-sm font-medium text-blue-800">
            Excel 数据 ({data.address})
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full text-xs">
            {data.headers && (
              <thead>
                <tr className="border-b border-blue-200">
                  {data.headers.slice(0, maxColumnsToShow).map((header, index) => (
                    <th key={index} className="text-left py-1 px-2 font-medium text-blue-800">
                      {header}
                    </th>
                  ))}
                  {data.headers.length > maxColumnsToShow && (
                    <th className="text-left py-1 px-2 font-medium text-blue-800">...</th>
                  )}
                </tr>
              </thead>
            )}
            <tbody>
              {data.values.slice(0, maxRowsToShow).map((row, rowIndex) => (
                <tr key={rowIndex} className={rowIndex % 2 === 0 ? "bg-blue-25" : ""}>
                  {row.slice(0, maxColumnsToShow).map((cell, cellIndex) => (
                    <td key={cellIndex} className="py-1 px-2 text-gray-700">
                      {String(cell)}
                    </td>
                  ))}
                  {row.length > maxColumnsToShow && (
                    <td className="py-1 px-2 text-gray-500">...</td>
                  )}
                </tr>
              ))}
              {data.values.length > maxRowsToShow && (
                <tr>
                  <td colSpan={Math.min(data.values[0]?.length || 1, maxColumnsToShow + 1)}
                      className="py-1 px-2 text-center text-gray-500 text-xs">
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

  return (
    <div className="flex flex-col h-full bg-white">
      {/* Messages Area */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {messages.length === 0 && (
          <div className="text-center text-gray-500 mt-8">
            <Bot24Regular className="mx-auto mb-2 text-2xl" />
            <p>开始与 Excel Agent 对话</p>
            <p className="text-sm mt-1">选择 Excel 数据后开始聊天</p>
          </div>
        )}

        {messages.map((message) => (
          <div
            key={message.id}
            className={`flex gap-3 ${
              message.type === 'user' ? 'flex-row-reverse' : 'flex-row'
            }`}
          >
            <div className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${
              message.type === 'user'
                ? 'bg-blue-600 text-white'
                : message.type === 'assistant'
                ? 'bg-green-600 text-white'
                : 'bg-gray-500 text-white'
            }`}>
              {message.type === 'user' ? (
                <Person24Regular className="w-4 h-4" />
              ) : message.type === 'assistant' ? (
                <Bot24Regular className="w-4 h-4" />
              ) : null}
            </div>

            <div className={`flex-1 max-w-[80%] ${
              message.type === 'user' ? 'text-right' : 'text-left'
            }`}>
              <div className={`inline-block px-4 py-2 rounded-lg ${
                message.type === 'user'
                  ? 'bg-blue-600 text-white'
                  : message.type === 'assistant'
                  ? 'bg-gray-100 text-gray-800'
                  : 'bg-yellow-100 text-yellow-800'
              }`}>
                <div className="whitespace-pre-wrap text-sm">{message.content}</div>
                {message.excelData && formatExcelData(message.excelData)}
              </div>
              <div className="text-xs text-gray-500 mt-1">
                {message.timestamp.toLocaleTimeString()}
              </div>
            </div>
          </div>
        ))}

        {isLoading && (
          <div className="flex gap-3">
            <div className="flex-shrink-0 w-8 h-8 rounded-full bg-green-600 text-white flex items-center justify-center">
              <Bot24Regular className="w-4 h-4" />
            </div>
            <div className="flex-1">
              <div className="inline-block px-4 py-2 rounded-lg bg-gray-100 text-gray-800">
                <div className="flex items-center gap-2">
                  <div className="flex space-x-1">
                    <div className="w-2 h-2 bg-gray-500 rounded-full animate-bounce"></div>
                    <div className="w-2 h-2 bg-gray-500 rounded-full animate-bounce" style={{ animationDelay: '0.1s' }}></div>
                    <div className="w-2 h-2 bg-gray-500 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
                  </div>
                  <span className="text-sm text-gray-600">AI 正在思考...</span>
                </div>
              </div>
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div className="border-t border-gray-200 p-4">
        {excelData && (
          <div className="mb-3 p-2 bg-blue-50 border border-blue-200 rounded-lg">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Table24Regular className="text-blue-600" />
                <span className="text-sm text-blue-800">
                  已选择数据: {excelData.address} ({excelData.values.length} 行)
                </span>
              </div>
            </div>
          </div>
        )}

        <form onSubmit={handleSubmit} className="flex gap-2">
          <input
            type="text"
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            placeholder="输入消息..."
            disabled={isLoading}
            className="flex-1 px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed"
          />
          <button
            type="submit"
            disabled={!inputValue.trim() || isLoading}
            className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition-colors"
          >
            <Send24Regular className="w-4 h-4" />
          </button>
        </form>
      </div>
    </div>
  );
};

export default ChatInterface;