import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Button, Input, Space, Card, Badge, Collapse, Table } from "antd";
import {
  SendOutlined,
  TableOutlined,
  UserOutlined,
  RobotOutlined,
  DeleteOutlined,
  FileTextOutlined,
  DownOutlined,
  UpOutlined,
} from "@ant-design/icons";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { ExcelData } from "../taskpane";

const { TextArea } = Input;

export interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  excelData?: ExcelData;
  attachedData?: ExcelData[];
  isDeletable?: boolean;
  dataId?: string;
  isStreaming?: boolean;
}

interface ChatInterfaceProps {
  messages: ChatMessage[];
  onSendMessage: (message: string) => void;
  onRemoveSelectedData?: (dataId: string) => void;
  onSelectExcelRange?: (address: string) => void;
  isLoading: boolean;
  selectedDataList: ExcelData[];
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
  const textareaRef = useRef<any>(null);
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

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      if (inputValue.trim() && !isLoading) {
        onSendMessage(inputValue.trim());
        setInputValue("");
      }
    }
  };

  const getDataTypeIcon = (dataType: string) => {
    switch (dataType) {
      case 'excel':
        return <TableOutlined />;
      case 'file':
        return <FileTextOutlined />;
      default:
        return <TableOutlined />;
    }
  };

  const formatExcelData = (data: ExcelData) => {
    const maxRowsToShow = 5;
    const maxColumnsToShow = 4;

    const columns = data.headers
      ? data.headers.slice(0, maxColumnsToShow).map((header, index) => ({
          title: header,
          dataIndex: index.toString(),
          key: index.toString(),
          ellipsis: true,
        }))
      : data.values[0]?.slice(0, maxColumnsToShow).map((_, index) => ({
          title: `列 ${index + 1}`,
          dataIndex: index.toString(),
          key: index.toString(),
          ellipsis: true,
        }));

    if (data.headers && data.headers.length > maxColumnsToShow) {
      columns.push({
        title: '...',
        dataIndex: 'more',
        key: 'more',
        ellipsis: true,
      });
    }

    const dataSource = data.values.slice(0, maxRowsToShow).map((row, rowIndex) => {
      const rowData: any = { key: rowIndex };
      row.slice(0, maxColumnsToShow).forEach((cell, cellIndex) => {
        rowData[cellIndex.toString()] = String(cell);
      });
      if (row.length > maxColumnsToShow) {
        rowData.more = '...';
      }
      return rowData;
    });

    return (
      <Card
        size="small"
        style={{ marginTop: 12, backgroundColor: '#f0f5ff', border: '1px solid #adc6ff' }}
        title={
          <Space>
            <TableOutlined style={{ color: '#1677ff' }} />
            <span style={{ fontSize: 14 }}>Excel 数据 ({data.address})</span>
          </Space>
        }
      >
        <Table
          columns={columns}
          dataSource={dataSource}
          pagination={false}
          size="small"
          bordered
        />
        {data.values.length > maxRowsToShow && (
          <div style={{ textAlign: 'center', marginTop: 8, color: '#666', fontSize: 12 }}>
            ... 还有 {data.values.length - maxRowsToShow} 行
          </div>
        )}
      </Card>
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

    return (
      <Card
        size="small"
        style={{
          marginTop: 8,
          backgroundColor: isUserMessage ? '#e6f4ff' : '#f6ffed',
          border: isUserMessage ? '1px solid #91caff' : '1px solid #b7eb8f'
        }}
      >
        <div
          onClick={() => setCollapsedData(prev => ({
            ...prev,
            [dataKey]: !prev[dataKey]
          }))}
          style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}
        >
          <Space size="small">
            {getDataTypeIcon('excel')}
            <span style={{ fontSize: 13 }}>使用了 {dataList.length} 个数据源</span>
          </Space>
          {isExpanded ? <UpOutlined style={{ fontSize: 12 }} /> : <DownOutlined style={{ fontSize: 12 }} />}
        </div>

        {isExpanded && (
          <div style={{ marginTop: 12 }}>
            <Space direction="vertical" size="small" style={{ width: '100%' }}>
              {dataList.map((data, index) => (
                <Card
                  key={`${messageId}-${data.id ?? index}`}
                  size="small"
                  hoverable
                  onClick={() => onSelectExcelRange?.(data.address)}
                  style={{ cursor: 'pointer' }}
                >
                  <Space>
                    {getDataTypeIcon('excel')}
                    <div>
                      <div style={{ fontWeight: 500, fontSize: 13 }}>{data.address}</div>
                      <div style={{ fontSize: 11, color: '#666' }}>
                        {`${data.values.length} 行`}
                        {data.headers ? " · 含标题" : ""}
                      </div>
                    </div>
                  </Space>
                </Card>
              ))}
            </Space>
          </div>
        )}
      </Card>
    );
  };

  const renderMessage = (message: ChatMessage) => {
    const isUser = message.type === 'user';
    const isSystem = message.type === 'system';

    if (isUser) {
      return (
        <div key={message.id} className="message-user slide-in-right" style={{ marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
            <div style={{ maxWidth: '80%' }}>
              <Card
                size="small"
                style={{
                  backgroundColor: '#1677ff',
                  color: 'white',
                  borderRadius: 12,
                  border: 'none'
                }}
              >
                <div style={{ whiteSpace: 'pre-wrap', fontSize: 14 }}>{message.content}</div>
              </Card>
              {message.attachedData && message.attachedData.length > 0 && (
                <div style={{ marginTop: 8 }}>
                  {renderAttachedDataList(message.attachedData, message.id, isUser)}
                </div>
              )}
            </div>
          </div>
        </div>
      );
    } else if (isSystem) {
      return (
        <div key={message.id} className="message-system fade-in" style={{ marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'center' }}>
            <Card
              size="small"
              style={{
                backgroundColor: '#fff7e6',
                border: '1px solid #ffd591',
                borderRadius: 8,
                textAlign: 'center'
              }}
            >
              <div style={{ fontSize: 13, color: '#d46b08' }}>{message.content}</div>
            </Card>
          </div>
        </div>
      );
    } else {
      return (
        <div key={message.id} className="message-assistant slide-in-left" style={{ marginBottom: 16 }}>
          <Card
            size="small"
            style={{
              backgroundColor: 'white',
              borderRadius: 12,
              maxWidth: '85%'
            }}
          >
            <div className="markdown-content">
              <Markdown remarkPlugins={[remarkGfm]}>
                {message.content}
              </Markdown>
              {message.isStreaming && (
                <span className="streaming-cursor">▋</span>
              )}
            </div>
            {message.excelData && formatExcelData(message.excelData)}
          </Card>
        </div>
      );
    }
  };

  return (
    <div className="chat-container">
      <div className="chat-messages">
        {messages.length === 0 && (
          <div style={{ textAlign: 'center', marginTop: 64 }} className="fade-in">
            <Card
              style={{
                width: 80,
                height: 80,
                margin: '0 auto 16px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                borderRadius: 16
              }}
            >
              <RobotOutlined style={{ fontSize: 36, color: '#1677ff' }} />
            </Card>
            <h3 className="welcome-title" style={{ marginBottom: 12 }}>开始与 Excel Agent 对话</h3>
            <p className="welcome-description">选择 Excel 数据后开始智能分析对话</p>
          </div>
        )}

        {messages.map(renderMessage)}

        {isLoading && (
          <div style={{ marginBottom: 16 }}>
            <Space>
              <div style={{
                width: 40,
                height: 40,
                borderRadius: '50%',
                background: 'linear-gradient(135deg, #52c41a 0%, #73d13d 100%)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: 'white'
              }}>
                <RobotOutlined />
              </div>
              <Card size="small" style={{ borderRadius: 12 }}>
                <Space>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce"></div>
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce" style={{ animationDelay: '0.1s' }}></div>
                    <div className="w-2 h-2 bg-green-500 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
                  </div>
                  <span style={{ fontSize: 14, color: '#666' }}>AI 正在分析中...</span>
                </Space>
              </Card>
            </Space>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      <div className="chat-input-container">
        {selectedDataList.length > 0 && (
          <div style={{ marginBottom: 16 }}>
            <div style={{ marginBottom: 8, fontSize: 13, fontWeight: 500, display: 'flex', alignItems: 'center', gap: 8 }}>
              <Badge status="processing" />
              已选中数据 ({selectedDataList.length} 个)
            </div>
            <Space direction="vertical" size="small" style={{ width: '100%', maxHeight: 144, overflowY: 'auto' }}>
              {selectedDataList.map((data) => (
                <Card
                  key={data.id}
                  size="small"
                  hoverable
                  onClick={() => onSelectExcelRange?.(data.address)}
                  style={{ cursor: 'pointer' }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                    <Space>
                      <div style={{
                        width: 40,
                        height: 40,
                        borderRadius: 12,
                        background: 'linear-gradient(135deg, #1677ff 0%, #4096ff 100%)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: 'white'
                      }}>
                        <TableOutlined />
                      </div>
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 500, color: '#1677ff' }}>
                          {data.address}
                        </div>
                        <div style={{ fontSize: 12, color: '#666' }}>
                          <span>{data.values.length} 行</span>
                          {data.headers && <span> · 包含标题</span>}
                          <span> · {data.timestamp.toLocaleTimeString()}</span>
                        </div>
                      </div>
                    </Space>
                    {onRemoveSelectedData && (
                      <Button
                        type="text"
                        danger
                        size="small"
                        icon={<DeleteOutlined />}
                        onClick={(e) => {
                          e.stopPropagation();
                          onRemoveSelectedData(data.id);
                        }}
                      />
                    )}
                  </div>
                </Card>
              ))}
            </Space>
          </div>
        )}

        <form onSubmit={handleSubmit}>
          <Space.Compact style={{ width: '100%' }}>
            <TextArea
              ref={textareaRef}
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="输入您的问题或指令... (Shift+Enter 换行)"
              disabled={isLoading}
              autoSize={{ minRows: 1, maxRows: 6 }}
              style={{ flex: 1 }}
            />
            <Button
              type="primary"
              htmlType="submit"
              disabled={!inputValue.trim() || isLoading}
              icon={<SendOutlined />}
              style={{ height: 'auto' }}
            >
              发送
            </Button>
          </Space.Compact>
        </form>
      </div>
    </div>
  );
};

export default ChatInterface;
