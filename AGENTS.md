# AGENTS.md

本项目是一个面向出版社责任编辑的 Word AI 审校助手。

## 项目目标

在 Word 中选中一段文字，点击插件按钮，调用 FastAPI 后端 AI 审校接口，返回结构化审校问题，并把审校建议作为 Word 批注插入当前选区。
做到好用、易用：
- 后端在选中文本中搜索original原文片段的位置，同步返回给word插件，word插件在original原文片段批注suggestion修改建议，这样编辑后续可以在word中直接接受或者拒绝批注建议，避免二次修改
- 历史记录可以清空、另存为、导入
- 兼容chat和reponse的AI API，可以在word插件手动切换

## 当前目录结构

```text
.
├── backend/
├── docs/
└── word-addin/
```

- `backend/`：Python FastAPI 后端服务，负责接收文本、调用 AI 审校能力、返回结构化审校结果。
- `docs/`：项目补充文档；临时联调记录、临时排障说明放在 `docs/tmp/`。
- `word-addin/`：Office.js + TypeScript + Webpack 的 Word 插件，负责读取 Word 选区、调用后端接口、把审校建议插入为 Word 批注。

## 开发约定

- 优先围绕项目目标做小步、可验证的改动，避免过早扩展复杂功能。
- 前后端通过 JSON API 交互，接口保持小而稳定。
- AI 返回内容必须先转换成结构化审校问题，再进入 Word 批注流程。
- Word 插件开发服务使用当前模板配置的 `https://localhost:3000/`。
- 后端默认使用 FastAPI，本地开发端口默认 `8000`。
- 不提交 API Key、`.env`、本地证书、调试缓存、依赖目录或构建产物。
- 开发完成后同步新建或更新文档。
- `README.md` 只保留长期有效的项目说明、启动方式、联调流程和测试命令。
- 临时问题修复说明、阶段性排障记录不要写进 `README.md`，统一放在 `docs/tmp/`；问题稳定解决后可以删除或归档。
- 改接口契约时同步更新 `spec.md`；改运行方式或联调流程时同步更新 `README.md`；改协作约定时同步更新本文件。

## MVP 数据流

```text
Word 选区文本
  -> word-addin 调用后端 JSON API
  -> backend/FastAPI 调用 AI 审校
  -> 返回结构化审校问题
  -> word-addin 在当前选区插入 Word 批注
```
