import * as React from "react";
import { useState, useEffect } from "react";
import { Form, Input, Button, Space, Card, Alert, Select, Spin, Divider } from "antd";
import {
  SettingOutlined,
  CheckCircleOutlined,
  CloseCircleOutlined,
  LoadingOutlined,
  ApiOutlined,
} from "@ant-design/icons";
import { OpenAIService } from "../services/openai"; // 使用向后兼容的别名
import { ApiSettings, getApiSettings, saveApiSettings, validateApiSettings } from "../utils/settings";

interface SettingsProps {
  onClose: () => void;
}

const Settings: React.FC<SettingsProps> = ({ onClose }) => {
  const [form] = Form.useForm();
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
    const savedSettings = getApiSettings();
    setSettings(savedSettings);
    form.setFieldsValue(savedSettings);
  }, [form]);

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
    <div className="h-full bg-gradient-to-br from-gray-50 to-blue-50/30 p-6">
      <div className="container max-w-2xl">
        <Space direction="vertical" size="large" style={{ width: '100%' }}>
          <Space>
            <SettingOutlined style={{ fontSize: 24, color: '#1677ff' }} />
            <h1 style={{ fontSize: 24, fontWeight: 600, margin: 0 }}>API 配置</h1>
          </Space>

          <Card className="fade-in" style={{ boxShadow: '0 2px 8px rgba(0,0,0,0.1)' }}>
            <Space direction="vertical" size="large" style={{ width: '100%' }}>
              <div>
                <h2 style={{ fontSize: 18, fontWeight: 500, marginBottom: 8 }}>
                  OpenAI 兼容 API 设置
                </h2>
                <p style={{ fontSize: 14, color: '#666', margin: 0 }}>
                  配置您的 AI 服务连接信息
                </p>
              </div>

              <Form
                form={form}
                layout="vertical"
                initialValues={settings}
                onValuesChange={(_, allValues) => setSettings(allValues)}
              >
                <Form.Item
                  label="Base URL"
                  name="baseUrl"
                  rules={[{ required: true, message: '请输入 Base URL' }]}
                >
                  <Input
                    prefix={<ApiOutlined />}
                    placeholder="https://api.openai.com/v1"
                    onChange={(e) => handleInputChange("baseUrl", e.target.value)}
                  />
                </Form.Item>

                <Form.Item
                  label="API Key"
                  name="apiKey"
                  rules={[{ required: true, message: '请输入 API Key' }]}
                  extra="API Key 将安全存储在本地"
                >
                  <Input.Password
                    placeholder="输入您的 API Key"
                    onChange={(e) => handleInputChange("apiKey", e.target.value)}
                  />
                </Form.Item>

                <Form.Item label="模型" name="model">
                  <Space.Compact style={{ width: '100%' }}>
                    {availableModels.length > 0 ? (
                      <Select
                        value={settings.model}
                        onChange={(value) => handleInputChange("model", value)}
                        options={availableModels.map(model => ({
                          label: model,
                          value: model
                        }))}
                        style={{ flex: 1 }}
                      />
                    ) : (
                      <Input
                        placeholder="gpt-3.5-turbo"
                        value={settings.model}
                        onChange={(e) => handleInputChange("model", e.target.value)}
                        style={{ flex: 1 }}
                      />
                    )}
                    <Button
                      onClick={handleLoadModels}
                      disabled={isLoadingModels || !settings.baseUrl || !settings.apiKey}
                      loading={isLoadingModels}
                    >
                      获取模型
                    </Button>
                  </Space.Compact>
                </Form.Item>

                <Form.Item>
                  <Button
                    type="default"
                    block
                    onClick={handleTestConnection}
                    disabled={isTestingConnection || !settings.baseUrl || !settings.apiKey}
                    icon={isTestingConnection ? <LoadingOutlined /> : <ApiOutlined />}
                  >
                    {isTestingConnection ? '测试连接中...' : '测试 API 连接'}
                  </Button>
                </Form.Item>

                {testResult && (
                  <Alert
                    message={testResult.message}
                    type={testResult.success ? 'success' : 'error'}
                    showIcon
                    icon={testResult.success ? <CheckCircleOutlined /> : <CloseCircleOutlined />}
                  />
                )}
              </Form>

              <Divider style={{ margin: 0 }} />

              <Space style={{ width: '100%', justifyContent: 'flex-end' }}>
                <Button onClick={onClose}>
                  取消
                </Button>
                <Button type="primary" onClick={handleSave}>
                  保存设置
                </Button>
              </Space>
            </Space>
          </Card>
        </Space>
      </div>
    </div>
  );
};

export default Settings;
