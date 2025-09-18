/* global localStorage, console, URL */

// API Settings utilities
export interface ApiSettings {
  baseUrl: string;
  apiKey: string;
  model: string;
}

const DEFAULT_SETTINGS: ApiSettings = {
  baseUrl: "https://api.openai.com/v1",
  apiKey: "",
  model: "gpt-3.5-turbo",
};

export const getApiSettings = (): ApiSettings => {
  try {
    const savedSettings = localStorage.getItem("egent-settings");
    if (savedSettings) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(savedSettings) };
    }
  } catch (error) {
    console.error("Failed to load settings:", error);
  }
  return DEFAULT_SETTINGS;
};

export const saveApiSettings = (settings: ApiSettings): boolean => {
  try {
    localStorage.setItem("egent-settings", JSON.stringify(settings));
    return true;
  } catch (error) {
    console.error("Failed to save settings:", error);
    return false;
  }
};

export const validateApiSettings = (
  settings: ApiSettings
): { valid: boolean; message?: string } => {
  if (!settings.baseUrl.trim()) {
    return { valid: false, message: "Base URL 不能为空" };
  }

  if (!settings.apiKey.trim()) {
    return { valid: false, message: "API Key 不能为空" };
  }

  if (!settings.model.trim()) {
    return { valid: false, message: "模型名称不能为空" };
  }

  // Validate URL format
  try {
    new URL(settings.baseUrl);
  } catch {
    return { valid: false, message: "Base URL 格式不正确" };
  }

  return { valid: true };
};
