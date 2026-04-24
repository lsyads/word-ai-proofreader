# AGENTS.md

本项目是一个面向出版社责任编辑的 Word AI 审校助手。

## 项目目标

先完成第一版最小闭环 MVP：

> 在 Word 中选中一段文字，点击插件按钮，调用 FastAPI 后端 AI 审校接口，返回结构化审校问题，并把审校建议作为 Word 批注插入当前选区。

## 当前目录结构

```text
.
├── backend/
└── word-addin/
```

- `backend/`：Python FastAPI 后端服务，负责接收文本、调用 AI 审校能力、返回结构化审校结果。
- `word-addin/`：Office.js + TypeScript + Webpack 的 Word 插件，负责读取 Word 选区、调用后端接口、把审校建议插入为 Word 批注。

## 开发约定

- 优先围绕 MVP 闭环做小步、可验证的改动，避免过早扩展复杂功能。
- 前后端通过 JSON API 交互，接口保持小而稳定。
- AI 返回内容必须先转换成结构化审校问题，再进入 Word 批注流程。
- Word 插件开发服务使用当前模板配置的 `https://localhost:3000/`。
- 后端默认使用 FastAPI，本地开发端口默认 `8000`。
- 不提交 API Key、`.env`、本地证书、调试缓存、依赖目录或构建产物。
- 开发完成后同步更新和补充文档。

## MVP 数据流

```text
Word 选区文本
  -> word-addin 调用后端 JSON API
  -> backend/FastAPI 调用 AI 审校
  -> 返回结构化审校问题
  -> word-addin 在当前选区插入 Word 批注
```
