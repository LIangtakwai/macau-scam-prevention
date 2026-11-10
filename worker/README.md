# DeepSeek 代理 · 部署说明

纯静态网站藏不住密钥：`materials/ai-config.js` 里写 `apiKey`，任何人查看网页源码都能拿走。
所以把密钥放进这个 Cloudflare Worker，前端只请求 Worker。免费额度足够课程演示。

Worker 同时管**每日用量**：默认全站 200 次/天、单个 IP 30 次/天，澳门时间 0 点重置。

## 一、部署（约 8 分钟）

```bash
cd worker
npm install -g wrangler               # 只需一次
wrangler login                        # 浏览器里点一次授权
wrangler secret put DEEPSEEK_API_KEY  # 粘贴你的 key，回车；不会写进任何文件

wrangler kv namespace create QUOTA    # 建每日额度计数器
```

`kv namespace create` 会输出一段配置，把 `id` 填进 `wrangler.toml`，并取消注释：

```toml
[[kv_namespaces]]
binding = "QUOTA"
id = "上一步输出的 id"
```

然后部署：

```bash
wrangler deploy
```

部署成功后会得到一个地址，形如：

```
https://macau-antifraud-ai.<你的账号>.workers.dev
```

> 不建 KV 直接部署也能成功，但每次 AI 请求都会返回 `503 额度服务未配置`——
> 这是故意的：宁可报错，也不要让人以为设了上限其实没设。

## 二、接到网站上

打开 `materials/ai-config.js`，把 `proxyUrl` 填成上面的地址，`apiKey` 保持空：

```js
"proxyUrl": "https://macau-antifraud-ai.xxx.workers.dev",
"apiKey": "",
```

刷新页面，在「常见问题解答」输入问题点「AI 提问」，能出回答即成功。

## 三、自测

```bash
# 正常提问
curl -i -X POST https://macau-antifraud-ai.xxx.workers.dev \
  -H 'Content-Type: application/json' \
  -d '{"messages":[{"role":"user","content":"刷单是诈骗吗"}]}'

# 查看今天全站用了多少次
curl https://macau-antifraud-ai.xxx.workers.dev/quota
```

返回 JSON 里带 `choices` 即正常；`401` 说明 `DEEPSEEK_API_KEY` 没设对；
`429` 说明当天额度用完（响应头 `Retry-After` 是距澳门时间 0 点的秒数）；
`503` 说明 KV 没绑定或额度服务异常。

`/quota` 返回示例：

```json
{"date":"2026-11-10","timezone":"Asia/Macau","used":12,"limit":200,
 "remaining":188,"perIpLimit":30,"resetInSeconds":49669}
```

## 四、每日额度怎么规定

默认策略（都可以改，改完 `wrangler deploy` 生效）：

| 规则 | 默认值 | 配置位置 |
| --- | --- | --- |
| 全站每天最多调用 | 200 次 | `wrangler.toml` → `DAILY_LIMIT` |
| 单个 IP 每天最多调用 | 30 次 | `wrangler.toml` → `IP_DAILY_LIMIT` |
| 单次回答 token 上限 | 请求可传 1–1200，默认 900 | `worker.js` → `MAX_TOKENS` |
| 重置时间 | 澳门时间（UTC+8）每天 00:00 | `worker.js` → `MACAU_OFFSET_MS` |

为什么是 200/30：

- 200 次/天对课堂演示和答辩完全够用，按每次约 1–2 千 tokens 估算，一天最多几十万 tokens，
  费用可控；真被公开传播也不会一夜烧掉余额。
- 30 次/天/人：正常同学问几个问题远远用不到，但能挡住一个人把全站额度刷完。
- 课堂演示时全班可能共用校园网出口 IP（NAT），如果 30 次不够，临时改成 80 再部署即可。

计数规则：

- 每次**通过校验的 POST** 先计数、再请求 DeepSeek，所以上游报错、超时也占一次额度，
  防止有人靠刷失败请求绕过上限。
- 计数存在 KV：`quota:day:2026-11-10`（全站）和 `quota:ip:2026-11-10:<IP 哈希>`
  （单 IP，只存哈希，不存原始地址），保留 2 天后自动过期。
- 被拒绝的请求不计数；成功响应会带 `X-Quota-Limit / X-Quota-Used / X-Quota-Remaining`
  三个头，前端需要的话可以显示剩余次数。

常用运维命令：

```bash
# 改额度：编辑 wrangler.toml 后重新部署
wrangler deploy

# 查看今天的键
wrangler kv key list --binding=QUOTA

# 今天被刷爆了，想立刻清零（日期换成 /quota 返回的 date）
wrangler kv key delete --binding=QUOTA "quota:day:2026-11-10"

# 临时关闭 AI：把 materials/ai-config.js 的 proxyUrl 清空即可，不用动 Worker
```

注意事项：

- 免费版 KV 每天只有 1000 次写，而这里每个请求写 2 次（全站 + 单 IP）。
  所以 `DAILY_LIMIT` 建议不超过 400；要更大就升级套餐，或改用 Durable Objects 做原子计数。
- KV 是最终一致的，极端并发下可能多放行几次，误差在个位数；要严格硬上限用 Durable Objects。
- 想按“token 用量”而不是“次数”限额，可以在响应里读 `usage.total_tokens` 再累计，
  但每个请求会多一次 KV 写，课程项目不建议。

## 五、限制与轮换

- Worker 现在做四件事：限制来源域名、限制 `max_tokens`、只转发 `messages`、每日额度。
  它挡不住有人用 curl 伪造来源，真正兜底的是给这把 key 设消费上限、定期轮换。
- 换 key：`wrangler secret put DEEPSEEK_API_KEY` 重新粘贴即可，不用改前端。
- 如果旧 key 曾出现在网页或仓库里，先去 <https://platform.deepseek.com/api_keys>
  删除/重置——删文件不等于撤销，历史记录和缓存里还有。
- DeepSeek 平台本身没有“每天最多 N 次”的开关，余额就是最后一道保险；
  建议在平台里开启余额提醒，并给这个项目单独建一把 key。

## 六、本地演示

`file://` 打开的页面来源是 `null`，Worker 允许这种来源，双击 HTML 也能用；
来源限制只针对正式部署的域名。本地没有 KV 时额度检查会返回 503，
可以用 `wrangler dev --remote`（带 KV）或直接部署到测试环境验证。
