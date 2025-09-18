import * as React from "react";
import { useState, useEffect } from "react";
import { Settings24Regular, Chat24Regular, Table24Regular, Home24Regular } from "@fluentui/react-icons";
import ChatInterface, { ChatMessage } from "./ChatInterface";
import Settings from "./Settings";
import { getSelectedData, getAllWorksheetData, ExcelData, selectRangeByAddress } from "../taskpane";
import { OpenAIService } from "../services/openai";
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

  useEffect(() => {
    // Check if API is configured on startup
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
    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      type: 'user',
      content,
      timestamp: new Date(),
    };

    setMessages(prev => [...prev, userMessage]);
    setIsLoading(true);

    try {
      const settings = getApiSettings();
      const openAIService = new OpenAIService(settings);

      // Convert messages to API format
      const conversationHistory = messages.filter(m => m.type !== 'system').map(m => ({
        role: m.type,
        content: m.content
      }));

      // 使用所有选中的数据
      const response = await openAIService.sendMessage(
        content,
        selectedDataList.length > 0 ? selectedDataList : undefined,
        conversationHistory
      );

      // 标记所有使用过的数据
      if (selectedDataList.length > 0) {
        const usedIds = selectedDataList.map(data => data.id);
        setUsedDataIds(prev => {
          const newSet = new Set(prev);
          usedIds.forEach(id => newSet.add(id));
          return newSet;
        });

        // 清空选中的数据列表
        setSelectedDataList([]);
      }

      const assistantMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        type: 'assistant',
        content: response.success ? response.message! : `错误: ${response.error}`,
        timestamp: new Date()
      };

      setMessages(prev => [...prev, assistantMessage]);
    } catch (error) {
      console.error('Error sending message:', error);
      const errorMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        type: 'assistant',
        content: '抱歉，发送消息时出现错误，请检查网络连接和 API 配置',
        timestamp: new Date()
      };
      setMessages(prev => [...prev, errorMessage]);
    } finally {
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
    <div className="h-full flex flex-col items-center justify-center p-6 bg-gradient-to-b from-blue-50 to-white">
      <div className="text-center max-w-md">
        <div className="mb-6">
          <Home24Regular className="text-6xl text-blue-600 mx-auto mb-4" />
          <h1 className="text-3xl font-bold text-gray-800 mb-2">Egent</h1>
          <p className="text-lg text-gray-600 mb-4">Excel AI 助手</p>
          <p className="text-sm text-gray-500">
            专为 Excel 数据分析设计的智能助手，支持 Chat 模式进行数据对话分析
          </p>
        </div>

        {!apiConfigured ? (
          <div className="space-y-4">
            <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
              <p className="text-yellow-800 text-sm">
                需要先配置 API 设置才能开始使用
              </p>
            </div>
            <button
              onClick={() => setCurrentView('settings')}
              className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium"
            >
              配置 API 设置
            </button>
          </div>
        ) : (
          <div className="space-y-4">
            <div className="p-4 bg-green-50 border border-green-200 rounded-lg">
              <p className="text-green-800 text-sm">✅ API 已配置，可以开始使用</p>
            </div>
            <button
              onClick={() => setCurrentView('chat')}
              className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium"
            >
              开始聊天
            </button>
          </div>
        )}

        <div className="mt-8 grid grid-cols-1 gap-4 text-sm">
          <div className="p-3 bg-white border border-gray-200 rounded-lg">
            <div className="flex items-center gap-2 mb-2">
              <Chat24Regular className="text-blue-600" />
              <span className="font-medium">Chat 模式</span>
            </div>
            <p className="text-gray-600 text-xs">
              读取选中的 Excel 数据，与 AI 进行对话分析，不会修改表格内容
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  const renderTopBar = () => (
    <div className="bg-white border-b border-gray-200 px-4 py-3">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <h1 className="text-lg font-semibold text-gray-800">Egent</h1>

          {currentView === 'chat' && (
            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-500">模式:</span>
              <div className="flex bg-gray-100 rounded-lg p-1">
                <button
                  onClick={() => setChatMode('chat')}
                  className={`px-3 py-1 text-xs rounded-md transition-colors ${
                    chatMode === 'chat'
                      ? 'bg-white text-blue-600 shadow-sm'
                      : 'text-gray-600 hover:text-gray-800'
                  }`}
                >
                  Chat
                </button>
                <button
                  onClick={() => setChatMode('agent')}
                  disabled
                  className="px-3 py-1 text-xs rounded-md text-gray-400 cursor-not-allowed"
                >
                  Agent (即将推出)
                </button>
              </div>
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          {currentView === 'chat' && (
            <>
              <button
                onClick={() => handleReadExcelData(false)}
                className="px-3 py-2 text-sm bg-blue-100 text-blue-700 rounded-md hover:bg-blue-200 transition-colors flex items-center gap-1"
              >
                <Table24Regular className="w-4 h-4" />
                读取选中数据
              </button>
              <button
                onClick={() => handleReadExcelData(true)}
                className="px-3 py-2 text-sm bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors"
              >
                读取全部数据
              </button>
              <button
                onClick={startNewChat}
                className="px-3 py-2 text-sm bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors"
              >
                新对话
              </button>
            </>
          )}

          <button
            onClick={() => setCurrentView(currentView === 'settings' ? 'chat' : 'settings')}
            className={`p-2 rounded-md transition-colors ${
              currentView === 'settings'
                ? 'bg-blue-100 text-blue-700'
                : 'text-gray-600 hover:bg-gray-100'
            }`}
          >
            <Settings24Regular className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  );

  return (
    <div className="h-screen flex flex-col bg-gray-50">
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
