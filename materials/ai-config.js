// AI 智能问答配置。
//
// 安全约定（2026-11-05 起）：仓库与发布包中 apiKey 必须为空，密钥只存放在
// Cloudflare Worker 的环境变量 DEEPSEEK_API_KEY 里，前端只填 proxyUrl。
// 本地临时直连演示时再手动填入 key，且不要上传、不要截图、不要提交；
// 演示完立刻清空。若 key 曾经出现在文件或网页源码中，请到
// https://platform.deepseek.com/api_keys 删除并重新生成——删文件不等于撤销。
window.ANTIFRAUD_AI_CONFIG = {
  "apiUrl": "https://api.deepseek.com/chat/completions",
  "proxyUrl": "",
  "apiKey": "",
  "model": "deepseek-chat",
  "temperature": 0.3,
  "maxTokens": 900,
  "maxContextChunks": 5
};
