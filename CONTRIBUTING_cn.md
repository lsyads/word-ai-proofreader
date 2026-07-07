# 贡献指南

语言：[English](CONTRIBUTING.md) | 简体中文

感谢你帮助改进 Word AI 审校助手。本项目是面向责任编辑的出版审校工作台，因此贡献应保留产品边界：Agent 产出建议和证据，编辑确认后才允许写回。

## 开发环境

后端：

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m pytest -q
```

前端：

```bash
cd word-addin
npm install
npm run lint
npm run build
```

使用 `.env.example` 作为模板。真实 Key 只放在本地 `.env` 中，不要写入跟踪文件。

## 贡献规则

- 改代码前先读 [AGENTS_cn.md](AGENTS_cn.md)。
- 保持 `agents/` 负责编排和人机确认状态，`services/` 负责可复用业务能力。
- 不绕过编辑确认直接写回。
- 默认不把完整稿件正文写入长期记忆。
- trace 和日志不得包含完整正文、API Key、Authorization header 或 Bearer token。
- 除非 V2 工作台目标确实需要，不要重写已验证的 DOCX、定位和写回服务。

## 文档同步

文档是双语的：

- 英文 `.md` 文件是主文档。
- 对应 `_cn.md` 文件必须保持内容一致。
- 英文文档链接英文文档。
- 中文文档链接 `_cn.md` 副本，语言切换链接除外。

改 API、schema、状态值或环境变量时，同步更新英文 `spec.md` 和中文 [spec_cn.md](spec_cn.md)。改入口、模块边界或数据流时，同步更新英文 `ARCHITECTURE.md` 和中文 [ARCHITECTURE_cn.md](ARCHITECTURE_cn.md)。改测试命令或验收口径时，同步更新英文 `TESTING.md` 和中文 [TESTING_cn.md](TESTING_cn.md)。

## Pull Request 清单

- 说明面向用户或开发者的行为变化。
- 说明运行过哪些测试。
- 不把无关重构混入 PR。
- 合同、命令或产品边界变化必须包含文档更新。
- 涉及配置、部署文档、日志或历史时，公开 PR 前运行 secret scan。
- 不包含真实稿件、API Key、`.env`、本地 SQLite 工作区、依赖目录或构建产物。

## 手工验证

有些行为无法完全由单元测试覆盖：

- Office.js 定位和当前选区写回需要 Word 桌面版。
- `npm run validate` 需要 Microsoft Office manifest validation service。
- 真实 AI provider 行为需要配置 Key 和兼容 endpoint。

如果 PR 跳过了某些手工验证，请明确说明。
