import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { ConfigProvider, theme } from "antd";
import zhCN from "antd/locale/zh_CN";
import "../styles/main.css";

/* global document, Office, module, require, HTMLElement */

const title = "Egent - Excel AI 助手";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <ConfigProvider
      locale={zhCN}
      theme={{
        token: {
          colorPrimary: "#1677ff",
          borderRadius: 8,
          colorBgContainer: "#ffffff",
        },
        algorithm: theme.defaultAlgorithm,
      }}
    >
      <App title={title} />
    </ConfigProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
