# Word 插件临时联调记录

本文记录 MVP 早期联调过程中遇到的临时问题。问题稳定解决后，可以删除或归档。

## 任务窗格只显示旁加载提示

原因：沿用了 Office 模板逻辑，`app-body` 默认 `display: none`，需要 `Office.onReady` 进入 Word 分支后才显示。

处理：已移除模板中的旁加载提示和 `display: none`，任务窗格主体默认显示。

## 点击“AI 审校”出现 `Script error.`

可能原因：旧 dev server 未加载新的 Webpack 配置，bundle 中仍存在 `process.env.WORD_ADDIN_API_BASE_URL`，而 Word WebView 中没有 `process`。

检查命令：

```bash
curl --noproxy localhost -k -s https://localhost:3000/taskpane.js | rg "process.env|API_BASE_URL"
```

处理：前端 MVP 暂不直接依赖 `process.env`；如修改 Webpack 配置，需要重启 dev server 并重新打开 Word 任务窗格。

## 点击“AI 审校”出现 `Load failed`

原因：Word 任务窗格页面是 `https://localhost:3000`，直接请求 `http://127.0.0.1:8000` 可能被 WebView 网络策略拦截。

处理：前端请求同源 `/api/proofread`，由 Webpack dev server 代理到 FastAPI：

```text
https://localhost:3000/api/proofread -> http://127.0.0.1:8000/api/proofread
```

修改 `word-addin/webpack.config.js` 后必须重启 dev server。

代理验证命令：

```bash
curl --noproxy localhost -k -X POST https://localhost:3000/api/proofread \
  -H 'Content-Type: application/json' \
  -d '{"text":"这是一段需要审校的文本。"}'
```
