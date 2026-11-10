/**
 * 反诈网站 · DeepSeek 代理（Cloudflare Worker）
 *
 * 浏览器只请求这个 Worker，Worker 再带上环境变量里的密钥去请求 DeepSeek。
 * 密钥只存在于 Worker 设置中，前端代码和仓库里都不会出现。
 * 部署步骤见同目录 README.md。
 *
 * 每日额度（2026-11-10 增加）：
 *   - 全站每天最多 DAILY_LIMIT 次（默认 200）；
 *   - 单个 IP 每天最多 IP_DAILY_LIMIT 次（默认 30），防止一个人把全站额度用完；
 *   - 以澳门时间（UTC+8）0 点重置，不用管服务器时区；
 *   - 计数存在 KV 命名空间 QUOTA 里，每次调用先计数、再请求上游，
 *     所以失败重试也算一次，避免有人靠刷错误请求绕过额度；
 *   - GET /quota 可以查看今天全站用了多少次（不返回 IP、问题等内容）。
 */

const FALLBACK_ORIGINS = [
  'https://liangtakwai.github.io',
  'https://zhuolin272-dotcom.github.io',
  'http://localhost:8000',
  'http://127.0.0.1:8000',
  'http://localhost:5500',
  'http://127.0.0.1:5500',
];
const MAX_TOKENS = 1200;
const MAX_MESSAGES = 8;

// 额度默认值；可在 wrangler.toml 的 [vars] 或 Cloudflare 面板里覆盖。
const DEFAULT_DAILY_LIMIT = 200;
const DEFAULT_IP_DAILY_LIMIT = 30;
// 计数保留 2 天，过期自动清理，不用手动删。
const QUOTA_TTL_SECONDS = 2 * 24 * 60 * 60;
// 澳门时间 UTC+8，无夏令时。
const MACAU_OFFSET_MS = 8 * 60 * 60 * 1000;

function corsHeaders(origin, allowed) {
  const allow = origin && origin !== 'null' && allowed.includes(origin)
    ? origin
    : (origin === 'null' ? '*' : (allowed[0] || '*'));
  return {
    'Access-Control-Allow-Origin': allow,
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Expose-Headers': 'X-Quota-Limit, X-Quota-Used, X-Quota-Remaining',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin',
  };
}

function json(obj, status, headers) {
  return new Response(JSON.stringify(obj), {
    status: status,
    headers: Object.assign({}, headers, { 'Content-Type': 'application/json; charset=utf-8' }),
  });
}

function positiveInt(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : fallback;
}

// 澳门时区的“今天”，例如 2026-11-10。
function macauDay(now) {
  return new Date((now === undefined ? Date.now() : now) + MACAU_OFFSET_MS)
    .toISOString().slice(0, 10);
}

// 距离澳门时间下一个 0 点还有多少秒，用于 429 的 Retry-After。
function secondsUntilMacauMidnight(now) {
  const t = (now === undefined ? Date.now() : now) + MACAU_OFFSET_MS;
  const shifted = new Date(t);
  const next = Date.UTC(shifted.getUTCFullYear(), shifted.getUTCMonth(), shifted.getUTCDate() + 1);
  return Math.max(1, Math.ceil((next - t) / 1000));
}

// IP 只存哈希，KV 里不出现原始地址。
async function hashIp(ip) {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(ip));
  return Array.from(new Uint8Array(digest))
    .slice(0, 8)
    .map(function (b) { return b.toString(16).padStart(2, '0'); })
    .join('');
}

function quotaLimits(env) {
  return {
    daily: positiveInt(env.DAILY_LIMIT, DEFAULT_DAILY_LIMIT),
    perIp: positiveInt(env.IP_DAILY_LIMIT, DEFAULT_IP_DAILY_LIMIT),
  };
}

async function readUsage(env, key) {
  const raw = await env.QUOTA.get(key);
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

// 先读两个计数，都没超才各加 1；超了就不写，避免把被拒绝的请求也算进去。
async function checkAndCount(request, env, day) {
  const limits = quotaLimits(env);
  const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
  const dayKey = 'quota:day:' + day;
  const ipKey = 'quota:ip:' + day + ':' + (await hashIp(ip));

  const values = await Promise.all([readUsage(env, dayKey), readUsage(env, ipKey)]);
  const used = values[0];
  const ipUsed = values[1];

  if (used >= limits.daily) {
    return { ok: false, scope: 'day', limit: limits.daily, used: used, remaining: 0 };
  }
  if (ipUsed >= limits.perIp) {
    return { ok: false, scope: 'ip', limit: limits.perIp, used: ipUsed, remaining: 0 };
  }

  await Promise.all([
    env.QUOTA.put(dayKey, String(used + 1), { expirationTtl: QUOTA_TTL_SECONDS }),
    env.QUOTA.put(ipKey, String(ipUsed + 1), { expirationTtl: QUOTA_TTL_SECONDS }),
  ]);

  return {
    ok: true,
    scope: 'ok',
    limit: limits.daily,
    used: used + 1,
    remaining: limits.daily - used - 1,
    ipUsed: ipUsed + 1,
    ipLimit: limits.perIp,
  };
}

function quotaError(quota, headers) {
  const retry = secondsUntilMacauMidnight();
  const message = quota.scope === 'ip'
    ? '今天你的提问次数已达上限（' + quota.limit + ' 次/天），请明天再试；紧急情况请拨打 993。'
    : '今天 AI 问答总额度已用完（' + quota.limit + ' 次/天），请明天再试，或先用站内问答检索。';
  return json({
    error: {
      message: message,
      type: 'quota_exceeded',
      scope: quota.scope,
      limit: quota.limit,
      used: quota.used,
      resetIn: retry,
    },
  }, 429, Object.assign({}, headers, { 'Retry-After': String(retry) }));
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    const allowed = (env.ALLOWED_ORIGINS ? env.ALLOWED_ORIGINS.split(',') : FALLBACK_ORIGINS)
      .map(function (s) { return s.trim(); }).filter(Boolean);
    const headers = corsHeaders(origin, allowed);
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: headers });

    // 查看今日用量：GET /quota，只返回数字，不含任何用户信息。
    if (request.method === 'GET' && url.pathname === '/quota') {
      if (!env.QUOTA) {
        return json({ error: { message: '额度服务未配置：请绑定名为 QUOTA 的 KV 命名空间，见 worker/README.md' } }, 503, headers);
      }
      const day = macauDay();
      const limits = quotaLimits(env);
      try {
        const used = await readUsage(env, 'quota:day:' + day);
        return json({
          date: day,
          timezone: 'Asia/Macau',
          used: used,
          limit: limits.daily,
          remaining: Math.max(0, limits.daily - used),
          perIpLimit: limits.perIp,
          resetInSeconds: secondsUntilMacauMidnight(),
        }, 200, headers);
      } catch (e) {
        return json({ error: { message: '额度服务暂不可用，请稍后再试' } }, 503, headers);
      }
    }

    if (request.method !== 'POST') return json({ error: { message: '只接受 POST' } }, 405, headers);
    if (origin && origin !== 'null' && !allowed.includes(origin)) {
      return json({ error: { message: '来源不在允许列表' } }, 403, headers);
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return json({ error: { message: '请求体不是合法 JSON' } }, 400, headers);
    }

    const messages = Array.isArray(body.messages) ? body.messages.slice(0, MAX_MESSAGES) : [];
    if (!messages.length) return json({ error: { message: '缺少 messages' } }, 400, headers);

    // 每日额度：没绑定 KV 时直接拒绝，而不是悄悄放开，免得以为设了上限其实没有。
    if (!env.QUOTA) {
      return json({ error: { message: '额度服务未配置：请绑定名为 QUOTA 的 KV 命名空间，见 worker/README.md' } }, 503, headers);
    }

    let quota;
    try {
      quota = await checkAndCount(request, env, macauDay());
    } catch (e) {
      return json({ error: { message: '额度服务暂不可用，请稍后再试' } }, 503, headers);
    }
    if (!quota.ok) return quotaError(quota, headers);

    const payload = {
      model: env.MODEL || 'deepseek-chat',
      messages: messages,
      temperature: 0.3,
      max_tokens: Math.min(Math.max(Number(body.max_tokens) || 900, 1), MAX_TOKENS),
      stream: false,
    };

    let upstream;
    try {
      upstream = await fetch('https://api.deepseek.com/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + env.DEEPSEEK_API_KEY,
        },
        body: JSON.stringify(payload),
      });
    } catch (e) {
      return json({ error: { message: '上游请求失败：' + e.message } }, 502, headers);
    }

    const text = await upstream.text();
    return new Response(text, {
      status: upstream.status,
      headers: Object.assign({}, headers, {
        'Content-Type': 'application/json; charset=utf-8',
        'X-Quota-Limit': String(quota.limit),
        'X-Quota-Used': String(quota.used),
        'X-Quota-Remaining': String(quota.remaining),
      }),
    });
  },
};
