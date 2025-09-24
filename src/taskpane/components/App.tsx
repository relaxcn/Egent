import * as React from "react";
import { useState, useEffect } from "react";
import { Settings24Regular, Chat24Regular, Table24Regular, Home24Regular } from "@fluentui/react-icons";
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
  const [selectedDataList, setSelectedDataList] = useState<ExcelData[]>([]); // 改为数组
  const [isLoading, setIsLoading] = useState(false);
  const [apiConfigured, setApiConfigured] = useState(false);
  const [usedDataIds, setUsedDataIds] = useState<Set<string>>(new Set()); // 跟踪已使用的数据ID
  const [excelOperations, setExcelOperations] = useState<ExcelActionResult[]>([]); // 跟踪Excel操作结果

  useEffect(() => {
    // Check if API is configured on startup
    const settings = getApiSettings();
    const validation = validateApiSettings(settings);
    setApiConfigured(validation.valid);

    if (validation.valid) {
      setCurrentView('chat');
    }

    // // Disable right-click context menu
    // const handleContextMenu = (e: MouseEvent) => {
    //   e.preventDefault();
    //   return false;
    // };

    // // Disable F12 developer tools
    // const handleKeyDown = (e: KeyboardEvent) => {
    //   if (e.key === 'F12') {
    //     e.preventDefault();
    //     return false;
    //   }
    //   return true;
    // };

    // document.addEventListener('contextmenu', handleContextMenu);
    // document.addEventListener('keydown', handleKeyDown);

    // // Cleanup
    // return () => {
    //   document.removeEventListener('contextmenu', handleContextMenu);
    //   document.removeEventListener('keydown', handleKeyDown);
    // };
  }, []);

  const handleReadExcelData = async (readAll: boolean = false) => {
    try {
      const data = readAll ? await getAllWorksheetData() : await getSelectedData();

      if (data) {
        // 检查是否已经选中了相同的数据
        const isDuplicate = selectedDataList.some(existing =>
          existing.address === data.address &&
          JSON.stringify(existing.values) === JSON.stringify(data.values)
        );

        if (!isDuplicate) {
          setSelectedDataList(prev => [...prev, data]);
        }
        // 移除了系统消息的添加，只显示在列表中
      } else {
        // 只在没有数据时显示错误消息
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

    // 创建初始的助手消息（空内容，标记为streaming）
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

      // 根据模式设置Excel操作功能
      enhancedOpenAIService.setExcelOperationsEnabled(chatMode === 'agent');

      // Convert messages to API format
      const conversationHistory = messages.filter(m => m.type !== 'system').map(m => ({
        role: m.type,
        content: m.content
      }));

      // 重置Excel操作记录
      setExcelOperations([]);

      // 创建增强的流式回调
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

          // 标记所有使用过的数据
          if (usedDataSnapshot && usedDataSnapshot.length > 0) {
            const usedIds = usedDataSnapshot.map(data => data.id);
            setUsedDataIds(prev => {
              const newSet = new Set(prev);
              usedIds.forEach(id => newSet.add(id));
              return newSet;
            });

            // 清空选中的数据列表
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
          // 记录Excel操作结果
          setExcelOperations(prev => [...prev, result]);

          // 可选：在UI中显示操作反馈
          console.log(`🔧 Excel操作: ${action}`, result);

          // 如果操作成功，可以显示一个临时消息
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

      // 使用增强的流式发送消息
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

  // 简化删除函数，只需要从选中列表中移除
  const handleRemoveSelectedData = (dataId: string) => {
    setSelectedDataList(prev => prev.filter(data => data.id !== dataId));
  };

  // 处理选中Excel范围
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
    // Re-check API configuration
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
    setSelectedDataList([]); // 清空选中数据列表
    setUsedDataIds(new Set()); // 重置已使用数据ID
  };

  const renderWelcomeScreen = () => (
    <div className="welcome-container fade-in">
      <div className="welcome-card scale-in">
        <div className="mb-6">
          <Home24Regular className="welcome-icon" />
          <h1 className="welcome-title">Egent</h1>
          <p className="welcome-subtitle">Excel AI 助手</p>
          <p className="welcome-description">
            专为 Excel 数据分析设计的智能助手，支持智能对话分析和Excel操作
          </p>
        </div>

        {!apiConfigured ? (
          <div className="space-y-4">
            <div className="alert alert-warning">
              <p className="text-sm">
                需要先配置 API 设置才能开始使用
              </p>
            </div>
            <button
              onClick={() => setCurrentView('settings')}
              className="button button-primary button-large"
            >
              配置 API 设置
            </button>
          </div>
        ) : (
          <div className="space-y-4">
            <div className="alert alert-success">
              <p className="text-sm">✅ API 已配置，可以开始使用</p>
            </div>
            <button
              onClick={() => setCurrentView('chat')}
              className="button button-primary button-large"
            >
              开始聊天
            </button>
          </div>
        )}

        <div className="mt-8 grid grid-cols-1 gap-4 text-sm">
          <div className="card hover-lift">
            <div className="card-body">
              <div className="flex items-center gap-2 mb-2">
                <Chat24Regular className="text-blue-600" />
                <span className="font-medium">Chat 模式</span>
              </div>
              <p className="text-gray-600 text-xs">
                读取选中的 Excel 数据，与 AI 进行对话分析，不会修改表格内容
              </p>
            </div>
          </div>
          <div className="card hover-lift">
            <div className="card-body">
              <div className="flex items-center gap-2 mb-2">
                <Table24Regular className="text-green-600" />
                <span className="font-medium">Agent 模式</span>
              </div>
              <p className="text-gray-600 text-xs">
                智能分析用户意图，可直接执行Excel操作：插入数据、设置颜色、创建图表等
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  const renderTopBar = () => (
    <div className="top-bar">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          {currentView === 'chat' && (
            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-500">模式:</span>
              <div className="mode-selector">
                <button
                  onClick={() => setChatMode('chat')}
                  className={`mode-button ${
                    chatMode === 'chat'
                      ? 'mode-button-active'
                      : 'mode-button-inactive'
                  }`}
                >
                  Chat
                </button>
                <button
                  onClick={() => setChatMode('agent')}
                  className={`mode-button ${
                    chatMode === 'agent'
                      ? 'mode-button-active'
                      : 'mode-button-inactive'
                  }`}
                >
                  Agent (Excel操作)
                </button>
              </div>
            </div>
          )}
        </div>

        <div className="toolbar">
          {currentView === 'chat' && (
            <>
              <button
                onClick={() => handleReadExcelData(false)}
                className="toolbar-button toolbar-button-primary"
              >
                <Table24Regular className="w-4 h-4" />
                读取选中数据
              </button>
              <button
                onClick={() => handleReadExcelData(true)}
                className="toolbar-button toolbar-button-secondary"
              >
                读取全部数据
              </button>
              <button
                onClick={startNewChat}
                className="toolbar-button toolbar-button-secondary"
              >
                新对话
              </button>
            </>
          )}

          <button
            onClick={() => setCurrentView(currentView === 'settings' ? 'chat' : 'settings')}
            className={`button button-ghost p-2 ${
              currentView === 'settings'
                ? 'bg-blue-100 text-blue-700'
                : ''
            }`}
          >
            <Settings24Regular className="w-4 h-4" />
          </button>
        </div>
      </div>
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
