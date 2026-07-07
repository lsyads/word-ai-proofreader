# Word AI 审校助手测试说明

语言：[English](TESTING.md) | 简体中文

本文汇总开发和 AI coding 场景下的验证命令，是最小验证矩阵的主文档。产品验收口径见 [spec_cn.md](spec_cn.md)，Windows 试点部署验证见 [DEPLOYMENT_cn.md](DEPLOYMENT_cn.md)。

## 快速检查

后端测试，离线可跑：

```bash
cd backend
source .venv/bin/activate
python -m pytest -q
```

前端静态检查和构建，离线可跑：

```bash
cd word-addin
npm run lint
npm run build
```

Manifest 校验，需要访问 Microsoft Office manifest validation service：

```bash
cd word-addin
npm run validate
```

本地健康检查，需要后端已启动：

```bash
curl --noproxy 127.0.0.1 http://127.0.0.1:8000/health
```

Word 手工联调，需要 Word 桌面版、Office 插件 dev server 和可用后端：

```bash
cd word-addin
npm run dev-server
npm run start
```

## 最小验证矩阵

| 改动类型 | 最小检查 |
| --- | --- |
| API/schema/状态字段 | `cd backend && python -m pytest -q`；检查 [spec_cn.md](spec_cn.md)、`word-addin/src/taskpane/types.ts`、`api.ts` 是否同步 |
| V2 Agent run、状态推进、失败重试 | `backend/tests/test_v2_workspace.py`、`backend/tests/test_agents.py`，必要时跑完整后端测试 |
| 候选队列、分页、批量决策、latest run 语义 | `backend/tests/test_v2_workspace.py`；前端改动再跑 `npm run lint` 和 `npm run build` |
| DOCX 抽取、定位、批注/修订、下载 | `backend/tests/test_docx.py`、`backend/tests/test_v2_workspace.py`，有前端下载改动再跑前端检查 |
| AI profile、Responses/Chat、mock fallback、timeout | `backend/tests/test_ai_profiles.py`、`backend/tests/test_ai_client.py`、`backend/tests/test_api.py` |
| trace 脱敏、run events、当前分块 timeout | `backend/tests/test_agents.py`、`backend/tests/test_v2_workspace.py`；前端进度展示改动再跑前端检查 |
| 前端 UI、按钮状态、刷新、定位、写回 | `cd word-addin && npm run lint && npm run build`；关键 Office.js 行为需要 Word 手工联调 |
| 前端依赖或 lockfile 变更 | `cd word-addin && npm ci && npm audit --registry=https://registry.npmjs.org && npm run lint && npm run build`；audit 需要官方 npm registry，因为部分镜像不实现 audit endpoint |
| 环境变量、启动方式、部署文档 | `backend/tests/test_settings.py`、`backend/tests/test_env_example.py`；检查 `.env.example`、[README_cn.md](README_cn.md)、[DEPLOYMENT_cn.md](DEPLOYMENT_cn.md) 是否同步 |
| Markdown 文档 | 检查英文/中文成对文件、语言切换、内部链接和模板/个人路径残留 |

单个后端文件可这样跑：

```bash
cd backend
source .venv/bin/activate
python -m pytest -q tests/test_v2_workspace.py
```

## 检查类型说明

- 离线可跑：后端 pytest、前端 lint、前端 build。
- 需要网络：`npm run validate`，以及真实远程 AI API 连通性检查。
- 需要 npm security API：`cd word-addin && npm audit --registry=https://registry.npmjs.org`。
- 需要 Word 桌面版：旁加载插件、定位原文、当前选区写回、批注/修订模式人工确认。
- 需要真实 AI key：真实 Responses/Chat 调用、远程模型 timeout 和 provider 兼容性验证。
- 不需要真实 AI key：mock fallback、schema 校验、候选状态、项目存储、DOCX 写回单元测试。

## 手工联调清单

1. 启动后端并确认 `/health` 返回 `{"status":"ok"}`。
2. 启动 `word-addin` dev server，并旁加载到 Word。
3. 当前选区审校：选中正文，填写书名，启动审校，确认候选展示、定位、接受/忽略和写回。
4. DOCX 审校：选择 `.docx` 文件，启动审校，接受候选，执行后端写回并下载审校后文件。
5. 长任务场景：确认运行中分块进度、当前分块 timeout 倒计时、刷新当前项目和继续等待入口不把等待误报为失败。
6. 失败分块场景：确认 `retry-failed` 只重试失败 chunk，新增候选进入同一 run，已写回输出在需要时标记 `output_stale=true`。

## 文档检查

本项目没有单独配置 Markdown linter。文档改动至少手工检查：

- 相对链接可从仓库根目录打开。
- 命令和 `word-addin/package.json`、`backend/requirements.txt` 中的实际脚本一致。
- 新增 API 字段、状态值或环境变量时，同步 [spec_cn.md](spec_cn.md)、[README_cn.md](README_cn.md) 和相关测试说明。
- 英文主文档和对应 `_cn.md` 文件保持内容一致。
- 英文文档链接英文文档；中文文档链接 `_cn.md` 副本，语言切换链接除外。
