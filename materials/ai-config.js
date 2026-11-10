// AI 智能问答配置。
//
// 当前为「浏览器直连」模式：apiKey 明文写在这里，会随 GitHub Pages 与仓库源码公开。
// 这是项目组在 2026-11-10 明确接受的取舍（见 README「AI 提问」）。
// 如需避免公开：按 worker/README.md 部署 Cloudflare Worker，把 proxyUrl 填成 Worker 地址，
// 并清空 apiKey；密钥只放 Worker 环境变量 DEEPSEEK_API_KEY。
// 如果这枚 key 之后要换掉，请到 https://platform.deepseek.com/api_keys 删除并重新生成——
// 只改文件不会让旧 key 失效。
window.ANTIFRAUD_AI_CONFIG = {
  "apiUrl": "https://api.deepseek.com/chat/completions",
  "proxyUrl": "",
  "apiKey": "sk-96c606c728d24b3eb202bbb637f7210b",
  "model": "deepseek-chat",
  "temperature": 0.3,
  "maxTokens": 900,
  "maxContextChunks": 5
};
