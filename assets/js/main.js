/* 界面逻辑与内容分离。全部宣传文字来自 materials/content.js。
   页面渲染只用 textContent 文本节点，材料内容不会作为 HTML 执行。 */
'use strict';
const content = window.ANTIFRAUD_CONTENT;
const main = document.querySelector('main');
const page = document.body.dataset.page;
// 创建 DOM 并写入文本。搜索高亮使用字面匹配，搜索词不作为正则或 HTML 执行。
function el(tag, cls, text) { const node = document.createElement(tag); if (cls) node.className = cls; if (text !== undefined) node.textContent = text; return node; }
function link(text, href, cls) { const a = el('a', cls, text); a.href = href; return a; }
function heading(title, subtitle, eyebrow) { main.append(el('p', 'eyebrow', eyebrow), el('h1', '', title), el('p', 'lead', subtitle)); }
function sectionHead(title, note, id) { const row = el('div', 'section-head'); if (id) row.id = id; row.append(el('h2', '', title), el('span', 'count', note)); main.append(row); return row; }
function highlight(node, text, keyword) {
  if (!keyword) { node.textContent = text; return; }
  const lower = text.toLocaleLowerCase(); let start = 0; let index;
  while ((index = lower.indexOf(keyword, start)) !== -1) {
    node.append(document.createTextNode(text.slice(start, index)), el('mark', '', text.slice(index, index + keyword.length)));
    start = index + keyword.length;
  }
  node.append(document.createTextNode(text.slice(start)));
}
function sourceLink(source) { const a = link('官方来源', source.url, 'source'); a.target = '_blank'; a.rel = 'noopener noreferrer'; a.setAttribute('aria-label', source.title + '（新窗口打开）'); return a; }
// 案例出处：详情页底部一行，链接与核对状态来自 materials/content.js。
function sourceLine(s) {
  const p = el('p', 'source-line'); p.append(el('b', '', '出处'));
  (s.links || []).forEach((l, i) => {
    const a = link(l.title, l.url, 'source');
    if (/^https?:/i.test(l.url)) { a.target = '_blank'; a.rel = 'noopener noreferrer'; }
    if (i) p.append(document.createTextNode(' · '));
    p.append(a);
  });
  if (s.note) p.append(document.createTextNode(' · ' + s.note));
  return p;
}

// ===== 智能问答（并入「搜索问题或答案」）：站内检索 + DeepSeek 生成回答 =====
// 大模型系统提示：约束回答只依据站内资料，防止编造。
let aiChunks = null; // 站内检索片段缓存，首次提问时构建
const AI_SYSTEM_PROMPT = '你是澳门大学蔡继有书院防诈宣传网站的 AI 防诈助手，服务对象是澳门在读大学生。\n' +
  '规则：\n' +
  '1. 只根据【参考资料】回答问题；资料中没有的内容，直接说明“站内资料暂未覆盖这个问题”，并建议拨打防诈骗查询热线 8800 7777，报案拨打 993。\n' +
  '2. 使用简体中文，回答简洁、分点列出，尽量不超过 200 字；不要使用 Markdown 格式；不要编造数据、法条或新闻。\n' +
  '3. 凡是涉及转账、垫付、保证金、解冻费的问题，先提醒“停止转账、核实身份”。\n' +
  '4. 引用参考资料时，在句末用 [编号] 标注来源。';
// 中文按二元组切词，英文数字按单词切分，用于本地相关性打分。
function aiTokenize(text) {
  const s = String(text).toLocaleLowerCase().replace(/[^\u4e00-\u9fa5a-z0-9]/g, ' ');
  const grams = new Set();
  s.split(/\s+/).forEach(function (w) {
    if (!w) return;
    if (/^[\u4e00-\u9fa5]{2,}$/.test(w)) { for (let i = 0; i < w.length - 1; i++) grams.add(w.slice(i, i + 2)); }
    else { grams.add(w); }
  });
  return grams;
}
// 把站内全部材料整理成可检索的片段（按需构建一次）。
function buildAiChunks() {
  const chunks = [];
  const push = function (cat, title, text) { if (text && text.trim()) chunks.push({ cat: cat, title: title, text: text.trim() }); };
  content.types.forEach(function (t) { push('诈骗类型', t.title, t.title + '。' + t.paragraphs.join('。')); });
  content.cases.forEach(function (c) { push('案例', c.title, c.title + '。' + c.paragraphs.join('。')); });
  content.faq.forEach(function (f) { push('常见问题', f.question, f.question + ' ' + f.answer); });
  (content.glossary || []).forEach(function (g) { push('名词解释', g.term, g.term + '：' + g.def); });
  (content.redFlags || []).forEach(function (r) { push('高危信号', '高危信号识别', r); });
  content.contacts.forEach(function (c) { push('求助渠道', c.name, c.name + '：' + c.number + '（' + c.detail + '）'); });
  content.college.forEach(function (c) { push('求助渠道', c.name, c.name + '：' + c.number + '（' + c.detail + '）'); });
  return chunks.map(function (c) {
    return { cat: c.cat, title: c.title, text: c.text, grams: aiTokenize(c.title + ' ' + c.text) };
  });
}
// 本地检索：二元组重合度 + 标题命中 + 原文包含加分，返回最相关片段。
const QUERY_STOP_GRAMS = ['怎么', '么办', '如何', '什么', '为什', '可以', '是否', '应该', '需要', '请问', '一下', '我的', '我要', '这个', '那个'];
function retrieveAi(query, topN) {
  if (!aiChunks) aiChunks = buildAiChunks();
  const q = aiTokenize(query);
  // 去掉「怎么办/如何/可以」这类疑问词，避免它们压过「被骗/兼职/验证码」等关键词。
  QUERY_STOP_GRAMS.forEach(function (g) { q.delete(g); });
  const scored = aiChunks.map(function (c) {
    let score = 0;
    q.forEach(function (g) { if (c.grams.has(g)) score += 1; });
    const lowerTitle = c.title.toLocaleLowerCase();
    q.forEach(function (g) { if (lowerTitle.includes(g)) score += 2; });
    if (query.trim().length >= 4 && c.text.includes(query.trim())) score += 6;
    if (c.cat === '常见问题') score *= 1.15;
    return { chunk: c, score: score };
  }).filter(function (x) { return x.score > 0; })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, topN)
    .map(function (x) { return x.chunk; });
  if (scored.length < 2) scored.push({ cat: '求助渠道', title: '紧急提醒', text: '怀疑遇到诈骗时：立即停止转账，保存聊天与转账凭证，拨打澳门司法警察局报案热线 993；防诈骗查询热线 8800 7777（24 小时）。' });
  return scored;
}
// 调用 DeepSeek：支持直连 apiUrl，也支持自建后端代理 proxyUrl（代理时不带密钥）。
async function askModel(query, refs) {
  const cfg = window.ANTIFRAUD_AI_CONFIG;
  const context = refs.map(function (c, i) { return '[' + (i + 1) + ']（' + c.cat + '）' + c.text; }).join('\n');
  const useProxy = !!(cfg && cfg.proxyUrl && String(cfg.proxyUrl).trim());
  const url = useProxy ? cfg.proxyUrl : cfg.apiUrl;
  const headers = { 'Content-Type': 'application/json' };
  if (!useProxy) headers.Authorization = 'Bearer ' + cfg.apiKey;
  const res = await fetch(url, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify({
      model: cfg.model || 'deepseek-chat',
      messages: [
        { role: 'system', content: AI_SYSTEM_PROMPT },
        { role: 'user', content: '【参考资料】\n' + context + '\n\n【问题】\n' + query }
      ],
      temperature: typeof cfg.temperature === 'number' ? cfg.temperature : 0.3,
      max_tokens: cfg.maxTokens || 900,
      stream: false
    })
  });
  if (!res.ok) {
    let detail = '';
    try { const err = await res.json(); if (err && err.error && err.error.message) detail = '：' + err.error.message; } catch (e) { /* 忽略解析失败 */ }
    throw new Error('AI 服务请求失败（HTTP ' + res.status + detail + '）');
  }
  const data = await res.json();
  const answer = data && data.choices && data.choices[0] && data.choices[0].message ? data.choices[0].message.content : '';
  if (!answer || !answer.trim()) throw new Error('AI 服务未返回有效回答');
  return answer;
}
// 本地兜底：AI 不可用时展示站内最相关的常见问题。
function localFallback(query) {
  const all = retrieveAi(query, 8);
  // 参考来源优先放问答，其次才是类型/案例，避免整段案例挤掉更贴题的问答。
  const faqRefs = all.filter(function (c) { return c.cat === '常见问题'; }).slice(0, 2);
  const otherRefs = all.filter(function (c) { return c.cat !== '常见问题'; });
  const refs = faqRefs.concat(otherRefs).slice(0, 3);
  const best = refs.find(function (c) { return c.cat === '常见问题'; }) || refs[0];
  let answer = best.text;
  // FAQ 片段的 text 以问题开头，卡片已有「问：」，这里去掉重复的问题。
  if (best.cat === '常见问题' && answer.indexOf(best.title) === 0) {
    answer = answer.slice(best.title.length).trim();
  }
  return { answer: answer, refs: refs };
}
// AI 回答卡片（问、答、参考来源、免责小字）。
function renderAiCard(query, answer, refs, fallback) {
  const card = el('article', 'ai-card');
  card.append(el('p', 'ai-question', '问：' + query));
  const body = el('div', 'ai-answer');
  answer.split(/\n+/).forEach(function (p) { const t = p.replace(/^[\s\-*•]+/, '').trim(); if (t) body.append(el('p', '', t)); });
  const srcTitle = el('p', 'ai-src-title', '参考站内资料：');
  const srcList = el('ul', 'ai-src');
  refs.forEach(function (r, i) { srcList.append(el('li', '', '[' + (i + 1) + '] ' + r.cat + ' · ' + r.title)); });
  const foot = el('p', 'ai-foot', fallback ? 'AI 服务暂时不可用，以上为站内资料直接匹配的结果。' : 'AI 生成内容仅供参考；紧急情况请拨打 993。');
  card.append(body, srcTitle, srcList, foot);
  return card;
}


// ===== 首页：整屏封面（居中标题、两个入口；不写导语、数据与小字） =====
function home() {
  const cover = el('section', 'cover');
  const inner = el('div', 'cover-inner');
  const actions = el('div', 'cover-actions');
  // 两个方向：了解骗局 / 求助（应急已并入求助页开头）
  actions.append(
    link('了解骗局', 'types.html', 'btn-main'),
    link('求助', 'help.html', 'btn-ghost')
  );
  // 封面上放两枚标志（深底白色版本），顶栏仍保留彩色标志
  const logos = el('div', 'cover-logos');
  const ck = el('img'); ck.src = 'materials/logos/ckyc-white.png'; ck.alt = '蔡继有书院标志'; ck.width = 68; ck.height = 76;
  const um = el('img'); um.src = 'materials/logos/um-white.png'; um.alt = '澳门大学标志'; um.width = 166; um.height = 61;
  logos.append(ck, um);
  inner.append(
    logos,
    el('span', 'cover-mark'),
    el('h1', '', '校园防诈骗宣传手册'),
    el('p', 'cover-en', (content.team && content.team.project) || 'SCAM PREVENTION CAMPAIGN'),
    el('div', 'rule'),
    actions
  );
  cover.append(inner);
  main.append(cover);
}

// ===== 认识诈骗类型：横向滑动的卡片轨道 + 全屏详情页 =====
// 一张图文卡片：图片封面 + 角标 + 两行标题 + 两行摘要
function mediaCard(opts) {
  const btn = el('button', 'media-card'); btn.type = 'button';
  const thumb = el('span', 'media-thumb');
  const img = el('img'); img.src = opts.image; img.alt = ''; img.loading = 'lazy';
  thumb.append(img);
  if (opts.tag) thumb.append(el('span', 'media-tag' + (opts.red ? ' red' : ''), opts.tag));
  if (opts.play) thumb.append(el('span', 'media-play', '▶'));
  const body = el('span', 'media-body');
  body.append(el('b', 'media-title', opts.title));
  if (opts.sub) body.append(el('span', 'media-sub', opts.sub));
  btn.append(thumb, body);
  btn.addEventListener('click', opts.onClick);
  return btn;
}
// 章节标题；传入轨道时右侧带左右翻页按钮
function blockHead(id, title, note, track) {
  const row = el('div', 'block-head'); row.id = id;
  const left = el('div', 'block-title');
  left.append(el('h2', '', title), el('span', '', note));
  row.append(left);
  if (track) {
    const nav = el('div', 'rail-nav');
    const prev = el('button', 'rail-arrow', '‹'); prev.type = 'button'; prev.setAttribute('aria-label', '向前翻');
    const next = el('button', 'rail-arrow', '›'); next.type = 'button'; next.setAttribute('aria-label', '向后翻');
    const step = () => Math.max(300, track.clientWidth * 0.85);
    prev.addEventListener('click', () => track.scrollBy({ left: -step(), behavior: 'smooth' }));
    next.addEventListener('click', () => track.scrollBy({ left: step(), behavior: 'smooth' }));
    const sync = () => {
      prev.disabled = track.scrollLeft < 8;
      next.disabled = track.scrollLeft + track.clientWidth >= track.scrollWidth - 8;
    };
    track.addEventListener('scroll', sync, { passive: true });
    window.addEventListener('resize', sync);
    setTimeout(sync, 60);
    nav.append(prev, next);
    row.append(nav);
  }
  main.append(row);
  return row;
}
// 横向轨道：鼠标拖动、触控板、按钮都能滑
function rail(cards, wide) {
  const track = el('div', 'rail' + (wide ? ' wide' : ''));
  track.setAttribute('role', 'list');
  cards.forEach(c => { c.setAttribute('role', 'listitem'); track.append(c); });
  return track;
}

// ---- 全屏详情页：铺满整屏，可上一个 / 下一个 / 返回 ----
let detailEl = null;
function renderDetail(build, label, nav) {
  const first = !detailEl;
  let page = detailEl, area;
  if (first) {
    page = el('div', 'detail');
    page.setAttribute('role', 'dialog'); page.setAttribute('aria-modal', 'true');
    page.setAttribute('aria-label', label || '详情');
    const bar = el('div', 'detail-bar');
    const back = el('button', 'detail-back', '← 返回'); back.type = 'button';
    const title = el('span', 'detail-bar-title', label || '');
    const navBox = el('div', 'detail-nav');
    const prev = el('button', 'detail-nav-btn', '‹'); prev.type = 'button'; prev.setAttribute('aria-label', '上一个');
    const next = el('button', 'detail-nav-btn', '›'); next.type = 'button'; next.setAttribute('aria-label', '下一个');
    prev.className = 'detail-nav-btn detail-prev';
    next.className = 'detail-nav-btn detail-next-btn';
    navBox.append(prev, next);
    bar.append(back, title, navBox);
    area = el('div', 'detail-scroll');
    page.append(bar, area);
    document.body.append(page);
    document.body.classList.add('u-lock');
    detailEl = page;
    try { history.pushState({ detail: true }, '', '#detail'); } catch (e) { /* file:// 下部分浏览器不允许改地址栏 */ }
    const onKey = e => {
      if (e.key === 'Escape') closeDetail();
      if (e.key === 'ArrowLeft' && detailEl && detailEl._prev) detailEl._prev();
      if (e.key === 'ArrowRight' && detailEl && detailEl._next) detailEl._next();
    };
    document.addEventListener('keydown', onKey);
    page._cleanup = () => document.removeEventListener('keydown', onKey);
    back.addEventListener('click', closeDetail);
    prev.addEventListener('click', () => { if (detailEl && detailEl._prev) detailEl._prev(); });
    next.addEventListener('click', () => { if (detailEl && detailEl._next) detailEl._next(); });
    page.tabIndex = -1; page.focus();
  } else {
    area = page.querySelector('.detail-scroll');
  }
  page.setAttribute('aria-label', label || '详情');
  page.querySelector('.detail-bar-title').textContent = label || '';
  area.replaceChildren();
  build(area);
  area.scrollTop = 0;
  page._prev = nav && nav.prev;
  page._next = nav && nav.next;
  page.querySelector('.detail-prev').disabled = !(nav && nav.prev);
  page.querySelector('.detail-next-btn').disabled = !(nav && nav.next);
}
function closeDetail(useHistory) {
  if (!detailEl) return;
  if (detailEl._cleanup) detailEl._cleanup();
  detailEl.remove(); detailEl = null;
  document.body.classList.remove('u-lock');
  if (useHistory !== false && history.state && history.state.detail) history.back();
}
window.addEventListener('popstate', () => {
  if (detailEl) {
    if (detailEl._cleanup) detailEl._cleanup();
    detailEl.remove(); detailEl = null;
    document.body.classList.remove('u-lock');
  }
});

// 详情页顶部大图 + 标题
function detailHero(parent, item, tag, red) {
  const hero = el('div', 'detail-hero');
  const img = el('img'); img.src = item.cover; img.alt = '';
  hero.append(img);
  if (tag) hero.append(el('span', 'media-tag' + (red ? ' red' : ''), tag));
  const text = el('div', 'detail-hero-text');
  text.append(el('h1', '', item.cardTitle || ''));
  if (item.origin) text.append(el('p', '', item.origin));
  hero.append(text);
  parent.append(hero);
}
function detailBody(parent) {
  const body = el('div', 'detail-body');
  parent.append(body);
  return body;
}
// 详情页底部的“下一个”入口
function detailNext(label, title, onClick) {
  const b = el('button', 'detail-next'); b.type = 'button';
  b.append(el('span', 'detail-next-label', label), el('b', '', title), el('span', 'detail-arrow', '→'));
  b.addEventListener('click', onClick);
  return b;
}

// B 站延伸视频列表（外链，新窗口打开）
function biliList(list) {
  const wrap = el('div', 'bili-wrap');
  wrap.append(el('p', 'bili-label', '延伸观看 · B 站'));
  const box = el('div', 'bili-list');
  list.forEach(b => {
    const a = el('a', 'bili-card');
    a.href = b.url; a.target = '_blank'; a.rel = 'noopener noreferrer';
    const txt = el('span', 'bili-text');
    txt.append(el('b', '', b.title), el('span', '', b.author + ' · ' + b.duration));
    a.append(el('span', 'bili-badge', 'B'), txt, el('span', 'bili-arrow', '↗'));
    box.append(a);
  });
  wrap.append(box);
  return wrap;
}

// 类型详情：可在 10 个类型之间直接前后切换
function showType(i) {
  const n = content.types.length;
  const t = content.types[i];
  const short = t.short || t.title.replace(/^[一二三四五六七八九十]+、/, '');
  renderDetail(area => {
    detailHero(area, { cover: t.cover, cardTitle: short, origin: t.title }, '类型 ' + String(i + 1).padStart(2, '0'));
    const body = detailBody(area);
    t.paragraphs.forEach(text => {
      const m = text.match(/^(类型名称|套路|受害人群|提醒)：([\s\S]*)$/);
      const row = el('div', 'detail-row');
      if (m) row.append(el('b', '', m[1]), el('p', '', m[2].trim()));
      else row.append(el('p', '', text));
      body.append(row);
    });
    if (t.sources) body.append(sourceLine({ links: t.sources }));
    if (t.bili) body.append(biliList(Array.isArray(t.bili) ? t.bili : [t.bili]));
    const nx = content.types[(i + 1) % n];
    body.append(detailNext('下一个类型', nx.short || '', () => showType((i + 1) % n)));
  }, '诈骗类型 ' + String(i + 1).padStart(2, '0') + ' · ' + short, {
    prev: () => showType((i - 1 + n) % n),
    next: () => showType((i + 1) % n)
  });
}

// 案例详情：可在 7 个案例之间前后切换
function showCase(i) {
  const n = content.cases.length;
  const c = content.cases[i];
  const loss = (c.paragraphs.find(p => p.startsWith('损失：')) || '').replace(/^损失：/, '').trim();
  const title = c.title.replace(/^案例\s*\d+[：:]\s*/, '');
  renderDetail(area => {
    detailHero(area, { cover: c.cover, cardTitle: title }, '损失 ' + (c.short || loss), true);
    const body = detailBody(area);
    c.paragraphs.forEach(text => {
      if (text.startsWith('损失：')) { const p = el('p', 'detail-loss'); p.append(el('b', '', text)); body.append(p); return; }
      const sub = /^(案情|诈骗套路重点|套路重点|防范提示)/.test(text);
      const point = /^要点：/.test(text);
      body.append(el('p', sub ? 'detail-h' : point ? 'detail-point' : '', point ? text.replace(/^要点：\s*/, '') : text));
    });
    if (c.source) body.append(sourceLine(c.source));
    body.append(detailNext('下一个案例', '案例 ' + ((i + 1) % n + 1), () => showCase((i + 1) % n)));
  }, '案例 ' + (i + 1), {
    prev: () => showCase((i - 1 + n) % n),
    next: () => showCase((i + 1) % n)
  });
}

// 视频详情：可在 4 段视频之间前后切换
function showVideo(i) {
  const n = content.videos.length;
  const v = content.videos[i];
  renderDetail(area => {
    const hero = el('div', 'detail-video');
    const vid = el('video'); vid.controls = true; vid.playsInline = true;
    vid.src = v.src; vid.poster = v.poster; vid.setAttribute('aria-label', v.title);
    hero.append(vid);
    const head = el('div', 'detail-head');
    head.append(el('h1', '', v.title), el('p', 'detail-origin', '本地宣传素材'));
    area.append(hero, head);
    const body = detailBody(area);
    body.append(el('p', '', '点击播放键观看完整影片；建议全屏观看。'));
    body.append(detailNext('下一段视频', content.videos[(i + 1) % n].title, () => showVideo((i + 1) % n)));
    vid.play().catch(() => { /* 浏览器阻止自动播放时，由用户点播放键 */ });
  }, '宣传视频 ' + String(i + 1).padStart(2, '0'), {
    prev: () => showVideo((i - 1 + n) % n),
    next: () => showVideo((i + 1) % n)
  });
}

function types() {
  heading('认识诈骗类型', '左右滑动查看，点开卡片进入完整内容。', '识骗学习');

  // 章节快捷入口
  const chips = el('nav', 'chip-nav'); chips.setAttribute('aria-label', '页面章节');
  [['fraud-types', '诈骗类型'], ['real-cases', '案例拆解'], ['red-flags', '高危信号'], ['videos', '宣传视频']]
    .forEach(([id, label]) => chips.append(link(label, '#' + id, 'chip')));
  main.append(chips);

  // ① 诈骗类型
  const typeCards = content.types.map((t, i) => {
    const short = t.short || t.title.replace(/^[一二三四五六七八九十]+、/, '');
    const sub = (t.paragraphs.find(p => p.startsWith('套路：')) || '').replace(/^套路：/, '').trim();
    return mediaCard({
      image: t.cover, tag: '类型 ' + String(i + 1).padStart(2, '0'), title: short, sub: sub,
      onClick: () => showType(i)
    });
  });
  const tTrack = rail(typeCards);
  blockHead('fraud-types', '诈骗类型', '10 类常见手法', tTrack);
  main.append(tTrack);

  // ② 案例拆解
  const caseCards = content.cases.map((c, i) => {
    const loss = (c.paragraphs.find(p => p.startsWith('损失：')) || '').replace(/^损失：/, '').trim();
    const title = c.title.replace(/^案例\s*\d+[：:]\s*/, '');
    return mediaCard({
      image: c.cover, tag: '损失 ' + (c.short || loss), red: true, title: title,
      onClick: () => showCase(i)
    });
  });
  const cTrack = rail(caseCards);
  blockHead('real-cases', '案例拆解', '7 个案例 · 每例标注出处与核对状态', cTrack);
  main.append(cTrack);

  // ③ 高危信号（纯文字，保持紧凑）
  blockHead('red-flags', '高危信号', content.redFlags.length + ' 条 · 对上任意一条都要警惕');
  const flags = el('div', 'flag-grid');
  (content.redFlags || []).forEach((r, i) => {
    const card = el('div', 'flag-card');
    card.append(el('span', 'flag-num', String(i + 1).padStart(2, '0')), el('p', '', r));
    flags.append(card);
  });
  main.append(flags);

  // ④ 宣传视频（放在页面最下面）
  const videoCards = content.videos.map((v, i) => mediaCard({
    image: v.cover || v.poster, tag: String(i + 1).padStart(2, '0'), play: true,
    title: v.title, sub: '点击播放',
    onClick: () => showVideo(i)
  }));
  const vTrack = rail(videoCards, true);
  blockHead('videos', '宣传视频', '4 段本地素材', vTrack);
  main.append(vTrack);

}

// ===== 常见问题解答：关键词搜索、问答、模拟帖子 =====
function faq() {
  heading('常见问题解答', '从学生常见疑问出发，找到下一步可以采取的行动。', '防诈问答');
  // 搜索 + AI 提问合并：输入即时筛选问答；回车或点「AI 提问」让 AI 结合站内资料回答
  const row = el('div', 'search-row');
  const box = el('div', 'search-box');
  const input = el('input'); input.type = 'search'; input.placeholder = '搜索问题或答案，如：转账、兼职、租房'; input.setAttribute('aria-label', '搜索防诈问题与答案');
  const clear = el('button', 'clear-search', '清除'); clear.type = 'button'; clear.hidden = true;
  box.append(input, clear);
  const ask = el('button', 'ai-ask', 'AI 提问'); ask.type = 'button'; ask.setAttribute('aria-label', '让 AI 结合站内资料回答');
  row.append(box, ask);
  const aiArea = el('div', 'ai-live'); aiArea.hidden = true; aiArea.setAttribute('aria-live', 'polite');
  const status = el('p', 'search-result'); status.setAttribute('role', 'status'); status.setAttribute('aria-live', 'polite');
  const results = el('div');
  main.append(row, aiArea, status, results);
  function render() {
    const keyword = input.value.trim().toLocaleLowerCase();
    let matches = content.faq.filter(x => (x.question + ' ' + x.answer).toLocaleLowerCase().includes(keyword));
    let fuzzy = false;
    // 自然语言问句（如「被骗了怎么办」）很少有完全匹配：用站内相关度算法补一轮推荐。
    if (keyword && !matches.length) {
      const titles = new Set(retrieveAi(keyword, 12)
        .filter(x => x.cat === '常见问题')
        .map(x => x.title));
      matches = content.faq.filter(x => titles.has(x.question));
      fuzzy = matches.length > 0;
    }
    clear.hidden = !keyword;
    status.textContent = keyword
      ? (matches.length ? `找到 ${matches.length} 个相关问答${fuzzy ? '（按相关度推荐）' : ''}` : '找到 0 个相关问答')
      : `全部 ${content.faq.length} 个问答`;
    results.replaceChildren();
    matches.forEach(item => {
      const d = el('details', 'faq-item'); d.open = false;
      const question = el('summary'); const answer = el('div', 'faq-answer');
      highlight(question, item.question, keyword); highlight(answer, item.answer, keyword);
      d.append(question, answer); results.append(d);
    });
    if (!matches.length) results.append(el('p', 'empty', aiArea.querySelector('.ai-card') ? '没有找到匹配的站内问答，AI 回答与参考来源见上方。' : '没有找到相关问答。试试“银行”“兼职”或“报案”等关键词。'));
  }
  input.addEventListener('input', render);
  // AI 提问：回车或点按钮触发；失败时自动降级为本地检索结果
  let aiBusy = false;
  async function askAI() {
    const query = input.value.trim();
    if (!query || aiBusy) return;
    const cfg = window.ANTIFRAUD_AI_CONFIG;
    aiArea.hidden = false; aiArea.replaceChildren();
    const aiStatus = el('p', 'ai-status', '正在检索站内资料并生成回答…');
    aiArea.append(aiStatus);
    if (!cfg) {
      const fb = localFallback(query);
      aiArea.replaceChildren(renderAiCard(query, fb.answer, fb.refs, true));
      aiArea.append(el('p', 'ai-error', '未加载 AI 配置（materials/ai-config.js），已切换为站内资料检索结果；紧急情况请拨打 993。'));
      render();
      return;
    }
    const viaProxy = !!(cfg.proxyUrl && String(cfg.proxyUrl).trim());
    if (!viaProxy && !String(cfg.apiKey || '').trim()) {
      const fb = localFallback(query);
      aiArea.replaceChildren(renderAiCard(query, fb.answer, fb.refs, true));
      aiArea.append(el('p', 'ai-error', 'AI 提问未开启（未配置密钥或代理），已切换为站内资料检索结果；紧急情况请拨打 993。'));
      render();
      return;
    }
    aiBusy = true; ask.disabled = true; ask.textContent = '思考中…';
    try {
      const refs = retrieveAi(query, 5);
      try {
        const answer = await askModel(query, refs);
        aiArea.replaceChildren(renderAiCard(query, answer, refs, false));
      } catch (err) {
        const fb = localFallback(query);
        aiArea.replaceChildren(renderAiCard(query, fb.answer, fb.refs, true));
        aiArea.append(el('p', 'ai-error', err.message + ' 已切换为站内本地检索结果。'));
      }
    } finally {
      aiBusy = false; ask.disabled = false; ask.textContent = 'AI 提问';
      render(); // AI 回答生成后刷新本地列表的空结果提示
    }
  }
  ask.addEventListener('click', askAI);
  input.addEventListener('keydown', e => { if (e.key === 'Enter') askAI(); });
  clear.addEventListener('click', () => { input.value = ''; aiArea.hidden = true; aiArea.replaceChildren(); render(); input.focus(); });
  render();
  // 模拟帖子
  sectionHead('同学提问');
  const posts = el('div', 'posts');
  content.posts.forEach(item => {
    const post = el('article', 'post');
    post.append(el('span', 'tag', item.tag), el('h3', '', item.question), el('p', '', '回复：' + item.answer));
    posts.append(post);
  });
  main.append(posts);
}

// ===== 求助：应急步骤、热线、书院渠道、证据清单、官方资源 =====
function contact(item, primary = false) {
  const card = el('article', 'contact' + (primary ? ' primary' : ''));
  card.append(el('h3', '', item.name), link(item.number, 'tel:' + item.tel, 'phone'), el('p', '', item.detail), sourceLink(item.source));
  return card;
}
function help() {
  heading('需要帮助？及时联系。', '怀疑遇到诈骗时，先停止转账与联系，保存证据，向警方及可信赖的人求助。', '求助渠道');
  const steps = el('div', 'steps');
  [['01 · 停止转账', '不要再支付保证金、解冻费或税费。'], ['02 · 保护账户', '联系银行或支付平台，保存聊天和付款记录。'], ['03 · 寻求协助', '尽快报案，并联系家人或书院人员。']].forEach(([t, p]) => {
    const block = el('div'); block.append(el('b', '', t), el('p', '', p)); steps.append(block);
  });
  main.append(steps);
  sectionHead('澳门本地与校园求助', '点击号码可发起拨号');
  const contacts = el('div', 'contacts');
  content.contacts.forEach((item, i) => contacts.append(contact(item, i === 0)));
  main.append(contacts);
  sectionHead('蔡继有书院求助渠道', '澳门大学 W12');
  const college = el('div', 'college-box');
  content.college.forEach(item => college.append(contact(item)));
  main.append(college);
  const mail = el('p', 'small-note');
  mail.append(document.createTextNode('书院邮箱：'), link(content.email, 'mailto:' + content.email), document.createTextNode('。也可向导师或辅导人员寻求协助。'));
  main.append(mail);
  // 官方资源链接
  sectionHead('官方资源链接', content.verified ? '核实日期 ' + content.verified : '');
  const resList = el('ul', 'resource-list');
  (content.resources || []).forEach(r => {
    const li = el('li'); const a = link(r.title, r.url, 'source'); a.target = '_blank'; a.rel = 'noopener noreferrer'; li.append(a); resList.append(li);
  });
  main.append(resList);
  // 内容出处与核对：取代原来的预留占位，页面不再出现"待补充"。
  if (content.casesNote) main.append(el('p', 'verified', content.casesNote + '案例出处核对日期 ' + (content.casesVerified || content.verified) + '。'));
  main.append(el('p', 'verified', '电话与开放时间以官方页面为准；拨号需设备支持。'));
}


// ===== 小彩蛋：连续点击封面标题或页脚左侧文字 5 次，显示制作团队 =====
function initEasterEgg() {
  const team = content.team;
  if (!team) return;
  let hits = 0, timer = null;
  function openTeam() {
    const wrap = el('div', 'egg');
    wrap.setAttribute('role', 'dialog');
    wrap.setAttribute('aria-modal', 'true');
    wrap.setAttribute('aria-label', '制作团队');
    const card = el('div', 'egg-card');
    const close = el('button', 'egg-close', '关闭'); close.type = 'button';
    const list = el('ul', 'egg-list');
    (team.members || []).forEach(m => {
      const li = el('li');
      li.append(el('b', '', m.name), el('span', '', m.role));
      list.append(li);
    });
    card.append(el('p', 'egg-kicker', (team.course || '') + ' · ' + (team.group || '')),
      el('h2', '', team.project || ''),
      el('p', 'egg-lead', '校园防诈骗宣传手册'),
      list, close);
    wrap.append(card);
    document.body.append(wrap);
    close.focus();
    const bye = () => { document.removeEventListener('keydown', onKey); wrap.remove(); };
    const onKey = e => { if (e.key === 'Escape') bye(); };
    document.addEventListener('keydown', onKey);
    close.addEventListener('click', bye);
    wrap.addEventListener('click', e => { if (e.target === wrap) bye(); });
  }
  function hit() {
    hits++;
    clearTimeout(timer);
    timer = setTimeout(() => { hits = 0; }, 1600);
    if (hits >= 5) { hits = 0; openTeam(); }
  }
  ['.cover h1', '.footer-inner'].forEach(sel => {
    const node = document.querySelector(sel);
    if (node) { node.style.cursor = 'default'; node.addEventListener('click', hit); }
  });
}

if (!content) { main.append(el('p', 'notice', '内容材料未加载，请确认 materials/content.js 文件存在。')); }
else { ({ home, types, faq, help }[page] || home)(); initEasterEgg(); }
