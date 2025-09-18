import * as React from "react";
import { useState, useEffect } from "react";
import { Settings24Regular, CheckmarkCircle24Regular, DismissCircle24Regular } from "@fluentui/react-icons";
import { OpenAIService } from "../services/openai";
import { ApiSettings, getApiSettings, saveApiSettings, validateApiSettings } from "../utils/settings";

interface SettingsProps {
  onClose: () => void;
}

const Settings: React.FC<SettingsProps> = ({ onClose }) => {
  const [settings, setSettings] = useState<ApiSettings>({
    baseUrl: "https://api.openai.com/v1",
    apiKey: "",
    model: "gpt-3.5-turbo",
  });
  const [isTestingConnection, setIsTestingConnection] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [availableModels, setAvailableModels] = useState<string[]>([]);
  const [isLoadingModels, setIsLoadingModels] = useState(false);

  useEffect(() => {
    // Load settings from localStorage
    const savedSettings = getApiSettings();
    setSettings(savedSettings);
  }, []);

  const handleSave = () => {
    const validation = validateApiSettings(settings);
    if (!validation.valid) {
      setTestResult({ success: false, message: validation.message || "配置验证失败" });
      return;
    }

    const success = saveApiSettings(settings);
    if (success) {
      onClose();
    } else {
      setTestResult({ success: false, message: "保存配置失败" });
    }
  };

  const handleInputChange = (field: keyof ApiSettings, value: string) => {
    setSettings(prev => ({
      ...prev,
      [field]: value
    }));
    // Clear test result when settings change
    setTestResult(null);
  };

  const handleTestConnection = async () => {
    const validation = validateApiSettings(settings);
    if (!validation.valid) {
      setTestResult({ success: false, message: validation.message || "请先完善配置信息" });
      return;
    }

    setIsTestingConnection(true);
    setTestResult(null);

    try {
      const service = new OpenAIService(settings);
      const result = await service.testConnection();
      setTestResult({
        success: result.success,
        message: result.success ? result.message! : result.error!
      });
    } catch (error) {
      setTestResult({
        success: false,
        message: error instanceof Error ? error.message : "连接测试失败"
      });
    } finally {
      setIsTestingConnection(false);
    }
  };

  const handleLoadModels = async () => {
    const validation = validateApiSettings(settings);
    if (!validation.valid) {
      setTestResult({ success: false, message: "请先完善配置信息并测试连接" });
      return;
    }

    setIsLoadingModels(true);
    try {
      const service = new OpenAIService(settings);
      const result = await service.getAvailableModels();

      if (result.success && result.models) {
        setAvailableModels(result.models);
        setTestResult({ success: true, message: `成功获取 ${result.models.length} 个可用模型` });
      } else {
        setTestResult({ success: false, message: result.error || "获取模型列表失败" });
      }
    } catch (error) {
      setTestResult({
        success: false,
        message: error instanceof Error ? error.message : "获取模型列表失败"
      });
    } finally {
      setIsLoadingModels(false);
    }
  };

  return (
    <div className="h-full bg-gray-50 p-6">
      <div className="max-w-2xl mx-auto">
        {/* Header */}
        <div className="flex items-center gap-3 mb-8">
          <Settings24Regular className="text-blue-600" />
          <h1 className="text-2xl font-semibold text-gray-800">API 配置</h1>
        </div>

        {/* Settings Card */}
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <div className="mb-6">
            <h2 className="text-lg font-medium text-gray-800 mb-2">OpenAI 兼容 API 设置</h2>
            <p className="text-sm text-gray-600">配置您的 AI 服务连接信息</p>
          </div>

          <div className="space-y-6">
            {/* Base URL */}
            <div>
              <label htmlFor="baseUrl" className="block text-sm font-medium text-gray-700 mb-2">
                Base URL <span className="text-red-500">*</span>
              </label>
              <input
                id="baseUrl"
                type="text"
                value={settings.baseUrl}
                onChange={(e) => handleInputChange("baseUrl", e.target.value)}
                placeholder="https://api.openai.com/v1"
                className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
              />
            </div>

            {/* API Key */}
            <div>
              <label htmlFor="apiKey" className="block text-sm font-medium text-gray-700 mb-2">
                API Key <span className="text-red-500">*</span>
              </label>
              <input
                id="apiKey"
                type="password"
                value={settings.apiKey}
                onChange={(e) => handleInputChange("apiKey", e.target.value)}
                placeholder="输入您的 API Key"
                className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
              />
              <p className="text-xs text-gray-500 mt-1">API Key 将安全存储在本地</p>
            </div>

            {/* Model */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <label htmlFor="model" className="block text-sm font-medium text-gray-700">
                  模型
                </label>
                <button
                  onClick={handleLoadModels}
                  disabled={isLoadingModels || !settings.baseUrl || !settings.apiKey}
                  className="px-3 py-1 text-xs bg-gray-100 text-gray-600 rounded hover:bg-gray-200 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                >
                  {isLoadingModels ? "加载中..." : "获取可用模型"}
                </button>
              </div>

              {availableModels.length > 0 ? (
                <select
                  id="model"
                  value={settings.model}
                  onChange={(e) => handleInputChange("model", e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
                >
                  {availableModels.map(model => (
                    <option key={model} value={model}>{model}</option>
                  ))}
                </select>
              ) : (
                <input
                  id="model"
                  type="text"
                  value={settings.model}
                  onChange={(e) => handleInputChange("model", e.target.value)}
                  placeholder="gpt-3.5-turbo"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
                />
              )}
            </div>

            {/* Test Connection */}
            <div>
              <button
                onClick={handleTestConnection}
                disabled={isTestingConnection || !settings.baseUrl || !settings.apiKey}
                className="w-full px-4 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 disabled:opacity-50 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2"
              >
                {isTestingConnection ? (
                  <>
                    <div className="w-4 h-4 border-2 border-gray-400 border-t-transparent rounded-full animate-spin"></div>
                    测试连接中...
                  </>
                ) : (
                  "测试 API 连接"
                )}
              </button>
            </div>

            {/* Test Result */}
            {testResult && (
              <div className={`p-3 rounded-lg flex items-center gap-2 ${
                testResult.success
                  ? "bg-green-50 border border-green-200 text-green-800"
                  : "bg-red-50 border border-red-200 text-red-800"
              }`}>
                {testResult.success ? (
                  <CheckmarkCircle24Regular className="text-green-600" />
                ) : (
                  <DismissCircle24Regular className="text-red-600" />
                )}
                <span className="text-sm">{testResult.message}</span>
              </div>
            )}
          </div>

          {/* Action Buttons */}
          <div className="flex gap-3 mt-8 pt-6 border-t border-gray-200">
            <button
              onClick={handleSave}
              className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition-colors font-medium"
            >
              保存设置
            </button>
            <button
              onClick={onClose}
              className="px-4 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 focus:ring-2 focus:ring-gray-500 focus:ring-offset-2 transition-colors"
            >
              取消
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Settings;