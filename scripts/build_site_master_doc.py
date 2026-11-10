# -*- coding: utf-8 -*-
"""生成《网站内容总表》：把 materials 里网站实际用到的内容整理成一份 Word。

数据来源：materials/content.json（网页实际读取的文字）+ materials 下的图片/视频/文档清单。
运行：python scripts/build_site_master_doc.py
"""
import json
import os
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'materials' / 'content.json').read_text(encoding='utf-8'))
OUT = ROOT / '网站内容总表.docx'
BLACK = '000000'

VIDEO_META = {
    'video-1.mp4': ('1:33.8', '1280×704', '14 MB'),
    'video-2.mp4': ('1:08.4', '1280×704', '10 MB'),
    'video-3.mp4': ('0:54.9', '1280×704', '4.7 MB'),
    'video-4.mp4': ('1:16.0', '1280×706', '14 MB'),
}

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.1)
sec.bottom_margin = Cm(2.0)


def set_cn(run, font='宋体'):
    run.font.name = font
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)


def run(p, text, size=10, bold=False, font='宋体'):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.font.color.rgb = RGBColor(0, 0, 0)
    set_cn(r, font)
    return r


def link(p, url, text, size=8.4):
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement('w:hyperlink'); hl.set(qn('r:id'), rid)
    r = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Calibri'); rf.set(qn('w:hAnsi'), 'Calibri'); rf.set(qn('w:eastAsia'), '宋体')
    rpr.append(rf)
    c = OxmlElement('w:color'); c.set(qn('w:val'), BLACK); rpr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rpr.append(sz)
    r.append(rpr)
    t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve'); r.append(t)
    hl.append(r); p._p.append(hl)


def P(text='', size=10, bold=False, font='宋体', before=0, after=3, left=0,
      first=0, line=1.42, align=None, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if left:
        pf.left_indent = Cm(left)
    if first:
        pf.first_line_indent = Cm(first)
    if align is not None:
        p.alignment = align
    if keep:
        pf.keep_with_next = True
    if text:
        run(p, text, size=size, bold=bold, font=font)
    return p


def H1(text):
    return P(text, size=14, bold=True, font='黑体', before=16, after=6, keep=True)


def H2(text):
    return P(text, size=11, bold=True, font='黑体', before=10, after=3, keep=True)


def label_line(label, body, size=9.8):
    p = P('', size=size, after=3)
    run(p, label + '：', size=size, bold=True)
    run(p, body, size=size)
    return p


def bullet(text, size=9.8, left=0.4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(left + 0.35); pf.first_line_indent = Cm(-0.35)
    pf.space_before = Pt(0); pf.space_after = Pt(2); pf.line_spacing = 1.38
    run(p, '· ', size=size)
    run(p, text, size=size)
    return p


def num(n, text, size=9.8):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.7); pf.first_line_indent = Cm(-0.45)
    pf.space_before = Pt(0); pf.space_after = Pt(2); pf.line_spacing = 1.38
    run(p, f'{n}. ', size=size, bold=True)
    run(p, text, size=size)
    return p


def url_line(url, size=8.4):
    p = P('', size=size, left=0.75, after=2)
    link(p, url, url, size=size)


def rule():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(8)
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '10')
    b.set(qn('w:space'), '1'); b.set(qn('w:color'), BLACK)
    bdr.append(b); ppr.append(bdr)


def img_size(path):
    try:
        with Image.open(path) as im:
            return f'{im.width}×{im.height}'
    except Exception:
        return '—'


def kb(path):
    try:
        return f'{os.path.getsize(path) // 1024} KB'
    except Exception:
        return '—'


# ---------------------------------------------------------------- 封面
P('澳门书院防诈宣传网站', size=13, bold=True, font='黑体', align=WD_ALIGN_PARAGRAPH.CENTER, before=40, after=2)
P('内容总表', size=26, bold=True, font='黑体', align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
P('CPED2001 第 11 组 · Scam Prevention Campaign', size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
P('整理日期：2026 年 11 月 10 日', size=9, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
rule()
P('说明：本表把 materials 中网站实际读取的文字、出处、联系方式与素材清单整理在一起，'
  '内容以 materials/content.json 为准；图片、视频与文档按文件实际信息列出。', size=9.5, after=4)

# ---------------------------------------------------------------- 目录
H1('目录')
for i, t in enumerate([
    '一、网站概况', '二、诈骗类型（10 类）', '三、案例（7 个）', '四、高危信号（14 条）',
    '五、名词解释（15 个）', '六、常见问题（35 个）', '七、模拟帖子（6 个）', '八、宣传视频（4 段）',
    '九、求助页信息', '十、AI 问答配置', '十一、素材与文件清单', '十二、未在页面展示的保留字段',
], 1):
    P(f'{i:02d}　{t}', size=9.5, after=1)

# ---------------------------------------------------------------- 一、概况
H1('一、网站概况')
team = DATA['team']
label_line('项目', team.get('project', ''))
label_line('课程 / 组别', f"{team.get('course', '')} · {team.get('group', '')}")
label_line('成员', '、'.join(f"{m['name']}（{m['role']}）" for m in team.get('members', [])))
label_line('介绍（保留字段）', DATA.get('intro', '（未使用）'))
P('网站共四个页面：首页（整屏封面，两个入口）、认识诈骗类型（10 类 + 7 个案例 + 14 条高危信号 + 4 段视频）、'
  '常见问题解答（35 个问答 + 6 个模拟帖子 + 站内检索/AI 提问）、求助（应急三步、联系方式、证据清单、官方链接）。', size=9.8, after=4)
label_line('内容数量', f"10 类诈骗 · 7 个案例 · 14 条高危信号 · 15 个名词 · 35 个问答 · 6 个模拟帖子 · 4 段视频")
label_line('核对日期', DATA.get('verified', ''))

# ---------------------------------------------------------------- 二、类型
H1('二、诈骗类型（10 类）')
for i, t in enumerate(DATA['types'], 1):
    H2(t["title"])
    if t.get('cover'):
        label_line('封面', t['cover'], size=8.8)
    for para in t['paragraphs']:
        m = para.split('：', 1)
        if len(m) == 2 and m[0] in ('类型名称', '套路', '受害人群', '提醒'):
            label_line(m[0], m[1])
        else:
            P(para, size=9.6, after=2)
    if t.get('bili'):
        P('延伸视频：', size=9.4, bold=True, after=1)
        for j, b in enumerate(t['bili'], 1):
            P(f'{j}. {b["title"]}（{b["author"]} · {b["duration"]}）', size=9.2, left=0.4, after=1)
            url_line(b['url'], size=8.2)

# ---------------------------------------------------------------- 三、案例
H1('三、案例（7 个）')
P('案例 4–6 为课程材料改编，已在出处中标注；其余为公开报道核对后的版本。', size=9.4, after=4)
for c in DATA['cases']:
    H2(c['title'])
    label_line('封面', c.get('cover', ''), size=8.8)
    for para in c['paragraphs']:
        if para in ('案情', '诈骗套路重点', '套路重点', '防范提示', '防范提示（网页展示用）'):
            P(para, size=9.6, bold=True, before=2, after=1)
        elif para.startswith('要点：'):
            bullet(para[3:], size=9.4)
        elif para.startswith('损失：'):
            label_line('损失', para[3:])
        else:
            P(para, size=9.6, after=2)
    src = c.get('source') or {}
    if src.get('links'):
        P('出处：', size=9.4, bold=True, after=1)
        for l in src['links']:
            P('· ' + l['title'], size=9.2, left=0.4, after=1)
            if l['url'].startswith('http'):
                url_line(l['url'], size=8.2)
            else:
                P('　站内文件：' + l['url'], size=8.6, left=0.75, after=1)
    if src.get('note'):
        P('说明：' + src['note'], size=8.8, left=0.4, after=4)

# ---------------------------------------------------------------- 四、信号
H1('四、高危信号（14 条）')
for i, r in enumerate(DATA['redFlags'], 1):
    num(i, r)

# ---------------------------------------------------------------- 五、名词
H1('五、名词解释（15 个）')
for i, g in enumerate(DATA['glossary'], 1):
    p = P('', size=9.8, after=3, left=0.35, first=-0.35)
    run(p, f'{i}. ', size=9.8, bold=True)
    run(p, g['term'] + '：', size=9.8, bold=True)
    run(p, g['def'], size=9.8)

# ---------------------------------------------------------------- 六、FAQ
H1('六、常见问题（35 个）')
for i, f in enumerate(DATA['faq'], 1):
    p = P('', size=9.8, before=2, after=1)
    run(p, f'Q{i}. ', size=9.8, bold=True)
    run(p, f['question'], size=9.8, bold=True)
    P('答：' + f['answer'], size=9.6, left=0.55, after=3)

# ---------------------------------------------------------------- 七、帖子
H1('七、模拟帖子（6 个）')
for i, post in enumerate(DATA['posts'], 1):
    p = P('', size=9.8, before=2, after=1)
    run(p, f'[{post.get("tag", "")}] ', size=9.2, bold=True)
    run(p, post['question'], size=9.8, bold=True)
    P('回复：' + post['answer'], size=9.6, left=0.55, after=3)

# ---------------------------------------------------------------- 八、视频
H1('八、宣传视频（4 段）')
P('四段均为本地 MP4，不自动播放；封面取自视频首帧。', size=9.4, after=3)
for v in DATA['videos']:
    name = Path(v['src']).name
    dur, res, size = VIDEO_META.get(name, ('—', '—', '—'))
    p = P('', size=9.8, before=2, after=1)
    run(p, v['title'] + '　', size=9.8, bold=True)
    run(p, f'{v["src"]}', size=9.4)
    label_line('时长 / 分辨率 / 大小', f'{dur} · {res} · {size}')
    label_line('封面 / 海报', f'{v.get("cover", "")} · {v.get("poster", "")}', size=8.8)

# ---------------------------------------------------------------- 九、求助
H1('九、求助页信息')
H2('9.1　应急三步')
for step, text in [
    ('停止转账', '不要再支付保证金、解冻费或税费。'),
    ('保护账户', '联系银行或支付平台，保存聊天和付款记录。'),
    ('寻求协助', '尽快报案，并联系家人或书院人员。'),
]:
    label_line(step, text)

H2('9.2　报案证据清单')
for item in DATA.get('helpChecklist', []):
    bullet(item)

H2('9.3　联系方式')
for ct in DATA['contacts']:
    p = P('', size=9.8, after=1)
    run(p, ct['name'] + '：', size=9.8, bold=True)
    run(p, f"{ct['number']}（{ct['detail']}）", size=9.8)
    if ct.get('source'):
        P('　来源：' + ct['source']['title'], size=8.6, left=0.5, after=1)
        url_line(ct['source']['url'], size=8.2)
for ct in DATA['college']:
    p = P('', size=9.8, after=1)
    run(p, ct['name'] + '：', size=9.8, bold=True)
    run(p, f"{ct['number']}（{ct['detail']}）", size=9.8)
    if ct.get('source'):
        P('　来源：' + ct['source']['title'], size=8.6, left=0.5, after=1)
        url_line(ct['source']['url'], size=8.2)
label_line('书院邮箱', DATA.get('email', ''))

H2('9.4　官方资源链接')
for i, res in enumerate(DATA['resources'], 1):
    P(f'{i}. {res["title"]}', size=9.6, left=0.4, after=1)
    url_line(res['url'], size=8.2)

H2('9.5　核对说明')
P(DATA.get('casesNote', ''), size=9.2, after=2)
label_line('案例出处核对日期', DATA.get('casesVerified', DATA.get('verified', '')), size=9)
label_line('电话与链接核对日期', DATA.get('verified', ''), size=9)

# ---------------------------------------------------------------- 十、AI
H1('十、AI 问答配置')
P('当前为浏览器直连：前端把问题与站内检索片段发给 DeepSeek，回答与参考来源显示在搜索框下方。'
  '密钥为明码写入（会随网页公开），本表不复制密钥内容。', size=9.6, after=3)
label_line('接口地址', 'https://api.deepseek.com/chat/completions')
label_line('模型', 'deepseek-chat；temperature 0.3；max_tokens 900；检索片段上限 5')
label_line('代理', 'proxyUrl 为空（直连）；也可按 worker/ 部署代理，改为不公开密钥的模式')
label_line('发送时机', '仅点击「AI 提问」或按回车时发送问题 + 站内片段；关键词筛选在本地完成')
label_line('降级', '请求失败或无密钥时，自动展示站内最相关问答，页面仍可使用')

# ---------------------------------------------------------------- 十一、素材
H1('十一、素材与文件清单')
H2('11.1　图片')
for pth in sorted((ROOT / 'materials' / 'covers').glob('*.jpg')):
    bullet(f'{pth.name}　{img_size(pth)}　{kb(pth)}　materials/covers/{pth.name}', size=9)
for pth in sorted((ROOT / 'materials' / 'logos').glob('*.png')):
    bullet(f'{pth.name}　{img_size(pth)}　{kb(pth)}　materials/logos/{pth.name}', size=9)
for pth in sorted((ROOT / 'materials' / 'videos').glob('poster-*.jpg')):
    bullet(f'{pth.name}　{img_size(pth)}　{kb(pth)}　materials/videos/{pth.name}', size=9)

H2('11.2　视频')
for pth in sorted((ROOT / 'materials' / 'videos').glob('*.mp4')):
    dur, res, size = VIDEO_META.get(pth.name, ('—', '—', '—'))
    bullet(f'{pth.name}　{dur}　{res}　{size}　materials/videos/{pth.name}', size=9)

H2('11.3　文字与配置文件')
for name, desc in [
    ('materials/content.json', '网页文字的 JSON 备份'),
    ('materials/content.js', '网页实际读取的文字材料'),
    ('materials/ai-config.js', 'AI 接口配置（含直连 key，本表不复制）'),
    ('materials/favicon.svg', '站点图标'),
    ('materials/案例.docx', '案例原始资料副本'),
    ('materials/诈骗类型.docx', '诈骗类型原始资料副本'),
]:
    pth = ROOT / name
    bullet(f'{name}　{kb(pth)}　{desc}', size=9)

H2('11.4　页面与脚本')
for name, desc in [
    ('index.html', '首页整屏封面'),
    ('types.html', '诈骗类型、案例、高危信号、视频'),
    ('faq.html', '常见问题、站内检索、AI 提问、模拟帖子'),
    ('help.html', '求助渠道、证据清单、官方资源'),
    ('assets/js/main.js', '页面渲染、搜索、AI 提问、视频切换'),
    ('assets/css/style.css', '配色、排版与响应式'),
    ('scripts/', '材料整理、来源修订、文档生成脚本（运行网站不需要）'),
    ('worker/', '可选的 AI 代理（Cloudflare Worker）'),
]:
    bullet(f'{name}　{desc}', size=9)

# ---------------------------------------------------------------- 十二、保留字段
_h12 = H1('十二、未在页面展示的保留字段')
_h12.paragraph_format.page_break_before = True
P('以下字段仍在 content.json/content.js 中，但当前页面未渲染，保留供以后使用。', size=9.4, after=4)
if DATA.get('slogan'):
    label_line('口号', DATA['slogan'])
if DATA.get('stats'):
    P('首页统计字段：', size=9.6, bold=True, after=1)
    for s in DATA['stats']:
        bullet(f"{s['value']}　{s['label']}", size=9.2)
if DATA.get('quickCards'):
    P('首页入口卡片字段：', size=9.6, bold=True, after=1)
    for q in DATA['quickCards']:
        bullet(f"{q['title']}：{q['text']}（{q['href']}）", size=9.2)
if DATA.get('quiz'):
    P('自测题：', size=9.6, bold=True, after=1)
    for i, q in enumerate(DATA['quiz'], 1):
        P(f'{i}. {q["question"]}', size=9.4, bold=True, after=1)
        P('　'.join(f'{"ABCD"[j]}. {opt}' for j, opt in enumerate(q['options'])), size=9.0, left=0.7, after=0)
        P(f'答案：{"ABCD"[q["answer"]]}　解释：{q["explain"]}', size=9, left=0.7, after=3)

# ---------------------------------------------------------------- 页脚
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(fp, '第 ', size=8)
r = fp.add_run()
f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'begin')
i = OxmlElement('w:instrText'); i.set(qn('xml:space'), 'preserve'); i.text = ' PAGE '
e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
r._r.append(f); r._r.append(i); r._r.append(e)
run(fp, ' 页', size=8)

cp = doc.core_properties
cp.title = '澳门书院防诈宣传网站 · 内容总表'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
