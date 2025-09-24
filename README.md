# Egent - A Powerfull Excel Agent

这是一个针对 Microsoft Office Excel 的 Agent 插件，使用 add-in web + React + Typescript 构建。

它是一个 taskpane 插件，类似于目前大多数流行的 AI Agent 工具那样，能够为 Excel 用于提供丰富的 LLM AI 功能，使 Excel 使用者能够减少使用成本，提高工作效率。

## 界面

它是一个类似 Github Copilot in Vscode 中的那样的面板，底部是一个可以输入的聊天框，左下角可以选择需要的模式 (#模式)。

除了底部，其余都是内容框，可以看到历史的聊天记录等信息，也可以看到修改的文件，点进去可以看到修改的内容（类似 git diff 的界面）。

顶部的右上角可以选择新的聊天，或者历史聊天记录的按钮。

## 模式

它具备以下模式：

1. Chat
2. Agent

Chat 模式类似于 ChatGPT 那样的纯聊天的模式，它只会从用户选择的表格或者区域块读取数据，并返回给用户在聊天界面结果，而不会修改用户的表格中的任何数据。

Agent 是一个自动化的代理工具，它可以对 Excel 表格进行修改，但仅限于用户选择的区域或者位置（范围）。

### Agent

Agent 使用 OpenAI Function Calling，让 LLM AI 与 Excel API 进行交互。

本项目定义了一套用于修改 Excel 数据的标准函数，它们直接与 Excel JS API 进行交互。例如：

1. insertText(Range, values)
2. changeColor(Range, color)

等等，以上只是一个示例，并不是真正的代码。

目前，本项目只使用 OpenAI 兼容的 API。