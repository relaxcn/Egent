import * as React from "react";
import { useState, useEffect } from "react";
import { Button, Space, Card, Divider, Segmented, Badge } from "antd";
import {
  SettingOutlined,
  MessageOutlined,
  TableOutlined,
  HomeOutlined,
  PlusOutlined,
  DownloadOutlined,
} from "@ant-design/icons";
import ChatInterface, { ChatMessage } from "./ChatInterface";
import Settings from "./Settings";
import { getSelectedData, getAllWorksheetData, ExcelData, selectRangeByAddress } from "../taskpane";
import { EnhancedOpenAIService, EnhancedStreamCallback, ExcelActionResult } from "../services/enhanced-openai";
import { getApiSettings, validateApiSettings } from "../utils/settings";

interface AppProps {
  title: string;
}

type AppView = 'chat' | 'settings' | 'welcome';
type ChatMode = 'chat' | 'agent';

const App: React.FC<AppProps> = () => {
  const [currentView, setCurrentView] = useState<AppView>('welcome');
  const [chatMode, setChatMode] = useState<ChatMode>('chat');
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [selectedDataList, setSelectedDataList] = useState<ExcelData[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [apiConfigured, setApiConfigured] = useState(false);
  const [usedDataIds, setUsedDataIds] = useState<Set<string>>(new Set());
  const [excelOperations, setExcelOperations] = useState<ExcelActionResult[]>([]);

  useEffect(() => {
    const settings = getApiSettings();
    const validation = validateApiSettings(settings);
    setApiConfigured(validation.valid);

    if (validation.valid) {
      setCurrentView('chat');
    }
  }, []);

  const handleReadExcelData = async (readAll: boolean = false) => {
    try {
      const data = readAll ? await getAllWorksheetData() : await getSelectedData();

      if (data) {
        const isDuplicate = selectedDataList.some(existing =>
          existing.address === data.address &&
          JSON.stringify(existing.values) === JSON.stringify(data.values)
        );

        if (!isDuplicate) {
          setSelectedDataList(prev => [...prev, data]);
        }
      } else {
        const errorMessage: ChatMessage = {
          id: Date.now().toString(),
          type: 'system',
          content: '未找到数据，请选择包含数据的单元格区域',
          timestamp: new Date()
        };
        setMessages(prev => [...prev, errorMessage]);
      }
    } catch (error) {
      console.error('Error reading Excel data:', error);
      const errorMessage: ChatMessage = {
        id: Date.now().toString(),
        type: 'system',
        content: '读取 Excel 数据失败，请重试',
        timestamp: new Date()
      };
      setMessages(prev => [...prev, errorMessage]);
    }
  };

  const handleSendMessage = async (content: string) => {
    const usedDataSnapshot = selectedDataList.length > 0
      ? selectedDataList.map((data) => ({ ...data }))
      : undefined;

    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      type: 'user',
      content,
      timestamp: new Date(),
      attachedData: usedDataSnapshot,
    };

    setMessages(prev => [...prev, userMessage]);
    setIsLoading(true);

    const assistantMessageId = (Date.now() + 1).toString();
    const assistantMessage: ChatMessage = {
      id: assistantMessageId,
      type: 'assistant',
      content: '',
      timestamp: new Date(),
      isStreaming: true
    };

    setMessages(prev => [...prev, assistantMessage]);

    try {
      const settings = getApiSettings();
      const enhancedOpenAIService = new EnhancedOpenAIService(settings);
      enhancedOpenAIService.setExcelOperationsEnabled(chatMode === 'agent');

      const conversationHistory = messages.filter(m => m.type !== 'system').map(m => ({
        role: m.type,
        content: m.content
      }));

      setExcelOperations([]);

      const enhancedStreamCallbacks: EnhancedStreamCallback = {
        onChunk: (chunk: string) => {
          setMessages(prev => prev.map(msg =>
            msg.id === assistantMessageId
              ? { ...msg, content: msg.content + chunk }
              : msg
          ));
        },
        onComplete: () => {
          setMessages(prev => prev.map(msg =>
            msg.id === assistantMessageId
              ? { ...msg, isStreaming: false }
              : msg
          ));
          setIsLoading(false);

          if (usedDataSnapshot && usedDataSnapshot.length > 0) {
            const usedIds = usedDataSnapshot.map(data => data.id);
            setUsedDataIds(prev => {
              const newSet = new Set(prev);
              usedIds.forEach(id => newSet.add(id));
              return newSet;
            });

            setSelectedDataList([]);
          }
        },
        onError: (error: string) => {
          setMessages(prev => prev.map(msg =>
            msg.id === assistantMessageId
              ? { ...msg, content: `错误: ${error}`, isStreaming: false }
              : msg
          ));
          setIsLoading(false);
        },
        onExcelAction: (action: string, result: ExcelActionResult) => {
          setExcelOperations(prev => [...prev, result]);
          console.log(`🔧 Excel操作: ${action}`, result);

          if (result.success) {
            const operationMessage: ChatMessage = {
              id: `excel-${Date.now()}`,
              type: 'system',
              content: `✅ Excel操作完成: ${result.content}`,
              timestamp: new Date()
            };
            setMessages(prev => [...prev, operationMessage]);
          } else if (result.error) {
            const errorMessage: ChatMessage = {
              id: `excel-error-${Date.now()}`,
              type: 'system',
              content: `❌ Excel操作失败: ${result.error}`,
              timestamp: new Date()
            };
            setMessages(prev => [...prev, errorMessage]);
          }
        }
      };

      await enhancedOpenAIService.sendMessageStreamWithExcel(
        content,
        enhancedStreamCallbacks,
        usedDataSnapshot,
        conversationHistory
      );

    } catch (error) {
      console.error('Error sending message:', error);
      setMessages(prev => prev.map(msg =>
        msg.id === assistantMessageId
          ? { ...msg, content: '抱歉，发送消息时出现错误，请检查网络连接和 API 配置', isStreaming: false }
          : msg
      ));
      setIsLoading(false);
    }
  };

  const handleRemoveSelectedData = (dataId: string) => {
    setSelectedDataList(prev => prev.filter(data => data.id !== dataId));
  };

  const handleSelectExcelRange = async (address: string) => {
    try {
      const success = await selectRangeByAddress(address);
      if (!success) {
        const errorMessage: ChatMessage = {
          id: Date.now().toString(),
          type: 'system',
          content: `无法选中范围 ${address}，请确保该范围存在`,
          timestamp: new Date()
        };
        setMessages(prev => [...prev, errorMessage]);
      }
    } catch (error) {
      console.error('Error selecting Excel range:', error);
      const errorMessage: ChatMessage = {
        id: Date.now().toString(),
        type: 'system',
        content: '选中Excel范围时出现错误，请重试',
        timestamp: new Date()
      };
      setMessages(prev => [...prev, errorMessage]);
    }
  };

  const handleSettingsClose = () => {
    const settings = getApiSettings();
    const validation = validateApiSettings(settings);
    setApiConfigured(validation.valid);

    if (validation.valid) {
      setCurrentView('chat');
    } else {
      setCurrentView('welcome');
    }
  };

  const startNewChat = () => {
    setMessages([]);
    setSelectedDataList([]);
    setUsedDataIds(new Set());
  };

  const renderWelcomeScreen = () => (
    <div className="welcome-container fade-in">
      <div className="welcome-card scale-in">
        <Space direction="vertical" size="large" style={{ width: '100%' }}>
          <div style={{ textAlign: 'center' }}>
            <HomeOutlined style={{ fontSize: 48, color: '#1677ff', marginBottom: 16 }} />
            <h1 className="welcome-title">Egent</h1>
            <p className="welcome-subtitle">Excel AI 助手</p>
            <p className="welcome-description">
              专为 Excel 数据分析设计的智能助手，支持智能对话分析和Excel操作
            </p>
          </div>

          {!apiConfigured ? (
            <Space direction="vertical" size="middle" style={{ width: '100%' }}>
              <Card size="small" style={{ backgroundColor: '#fff7e6', border: '1px solid #ffd591' }}>
                <p style={{ margin: 0, fontSize: 14, color: '#d46b08' }}>
                  需要先配置 API 设置才能开始使用
                </p>
              </Card>
              <Button
                type="primary"
                size="large"
                block
                onClick={() => setCurrentView('settings')}
                icon={<SettingOutlined />}
              >
                配置 API 设置
              </Button>
            </Space>
          ) : (
            <Space direction="vertical" size="middle" style={{ width: '100%' }}>
              <Card size="small" style={{ backgroundColor: '#f6ffed', border: '1px solid #b7eb8f' }}>
                <p style={{ margin: 0, fontSize: 14, color: '#52c41a' }}>
                  ✅ API 已配置，可以开始使用
                </p>
              </Card>
              <Button
                type="primary"
                size="large"
                block
                onClick={() => setCurrentView('chat')}
                icon={<MessageOutlined />}
              >
                开始聊天
              </Button>
            </Space>
          )}

          <Divider />

          <Space direction="vertical" size="middle" style={{ width: '100%' }}>
            <Card hoverable size="small">
              <Space>
                <MessageOutlined style={{ fontSize: 24, color: '#1677ff' }} />
                <div>
                  <div style={{ fontWeight: 500, marginBottom: 4 }}>Chat 模式</div>
                  <p style={{ margin: 0, fontSize: 12, color: '#666' }}>
                    读取选中的 Excel 数据，与 AI 进行对话分析，不会修改表格内容
                  </p>
                </div>
              </Space>
            </Card>
            <Card hoverable size="small">
              <Space>
                <TableOutlined style={{ fontSize: 24, color: '#52c41a' }} />
                <div>
                  <div style={{ fontWeight: 500, marginBottom: 4 }}>Agent 模式</div>
                  <p style={{ margin: 0, fontSize: 12, color: '#666' }}>
                    智能分析用户意图，可直接执行Excel操作：插入数据、设置颜色、创建图表等
                  </p>
                </div>
              </Space>
            </Card>
          </Space>
        </Space>
      </div>
    </div>
  );

  const renderTopBar = () => (
    <div className="top-bar">
      <Space style={{ width: '100%', justifyContent: 'space-between' }}>
        <Space>
          {currentView === 'chat' && (
            <Space size="small">
              <span style={{ fontSize: 14, color: '#666' }}>模式:</span>
              <Segmented
                value={chatMode}
                onChange={(value) => setChatMode(value as ChatMode)}
                options={[
                  { label: 'Chat', value: 'chat', icon: <MessageOutlined /> },
                  { label: 'Agent', value: 'agent', icon: <TableOutlined /> }
                ]}
              />
            </Space>
          )}
        </Space>

        <Space size="small">
          {currentView === 'chat' && (
            <>
              <Button
                type="primary"
                size="small"
                icon={<TableOutlined />}
                onClick={() => handleReadExcelData(false)}
              >
                读取选中
              </Button>
              <Button
                size="small"
                icon={<DownloadOutlined />}
                onClick={() => handleReadExcelData(true)}
              >
                读取全部
              </Button>
              <Button
                size="small"
                icon={<PlusOutlined />}
                onClick={startNewChat}
              >
                新对话
              </Button>
            </>
          )}

          <Button
            type={currentView === 'settings' ? 'primary' : 'default'}
            size="small"
            icon={<SettingOutlined />}
            onClick={() => setCurrentView(currentView === 'settings' ? 'chat' : 'settings')}
          />
        </Space>
      </Space>
    </div>
  );

  return (
    <div className="h-screen flex flex-col bg-gradient-to-br from-gray-50 to-blue-50/30">
      {currentView !== 'welcome' && renderTopBar()}

      <div className="flex-1 overflow-hidden">
        {currentView === 'welcome' && renderWelcomeScreen()}

        {currentView === 'chat' && (
          <ChatInterface
            messages={messages}
            onSendMessage={handleSendMessage}
            onRemoveSelectedData={handleRemoveSelectedData}
            onSelectExcelRange={handleSelectExcelRange}
            isLoading={isLoading}
            selectedDataList={selectedDataList}
          />
        )}

        {currentView === 'settings' && (
          <Settings onClose={handleSettingsClose} />
        )}
      </div>
    </div>
  );
};

export default App;
