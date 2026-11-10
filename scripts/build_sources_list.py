# -*- coding: utf-8 -*-
"""生成黑白版《防诈宣传资料来源清单》。

只列来源与链接，不做彩色排版、章节分析或修订对照。
运行：python scripts/build_sources_list.py
"""
import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'materials' / 'content.json').read_text(encoding='utf-8'))
OUT = ROOT / '防诈宣传资料来源与核查报告.docx'
BLACK = '000000'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.2)
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


def link(p, url, text, size=8.5):
    part = p.part
    rid = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('r:id'), rid)
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Calibri')
    rf.set(qn('w:hAnsi'), 'Calibri')
    rf.set(qn('w:eastAsia'), '宋体')
    rpr.append(rf)
    c = OxmlElement('w:color'); c.set(qn('w:val'), BLACK); rpr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rpr.append(sz)
    r.append(rpr)
    t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve')
    r.append(t)
    hl.append(r)
    p._p.append(hl)


def P(text='', size=10, bold=False, font='宋体', before=0, after=2, left=0,
      first=0, line=1.3, align=None, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
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


def rule():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '12')
    b.set(qn('w:space'), '1'); b.set(qn('w:color'), BLACK)
    bdr.append(b)
    ppr.append(bdr)


def H1(text):
    p = P(text, size=14, bold=True, font='黑体', before=14, after=6, keep=True)
    return p


def H2(text):
    return P(text, size=10.5, bold=True, font='黑体', before=8, after=2, keep=True)


def source_block(title, items, note=None):
    """items: [(标题, url 或 None, 附加说明)]"""
    for i, (name, url, extra) in enumerate(items, 1):
        p = P('%d. %s' % (i, name), size=9.8, left=0.75, first=-0.45)
        if extra:
            run(p, '（' + extra + '）', size=9.2)
        if url:
            u = P('', size=8.5, left=1.35, after=3)
            link(u, url, url, size=8.5)
    if note:
        P('核对说明：' + note, size=9.2, left=0.75, after=6)


# ---------------------------------------------------------------- 标题
P('防诈宣传资料来源清单', size=22, bold=True, font='黑体', after=4)
P('CPED2001 第 11 组 · Scam Prevention Campaign', size=10, after=2)
P('核查日期：2026 年 11 月 5 日　　所有链接当日访问', size=9, after=6)
rule()
P('本清单列出网站使用的全部资料来源：7 个案例的原始报道、10 类诈骗的权威依据，以及官方电话、法规、'
  '媒体报道、宣传视频、图片、B 站视频与 AI 服务。案例 1、2、3、7 已找到原始报道；'
  '案例 4–6 为课程材料改编，列出相近报道供核对。', size=9.5, after=4)

# ---------------------------------------------------------------- 案例
H1('一、案例来源')
CASE_GRADE = {
    1: '已核实。金额、时间、手法与两家报道一致。',
    2: '已核实。40 个单位、约 210 万港元、三名学生 160,900 港元与报道一致。',
    3: '已核实。金额、手法、地点与报道一致；报道中受害人为两名中年非本地男子，并非大学生。',
    4: '部分吻合。公开报道中未见完全对应的案件，按课程材料保留。',
    5: '部分吻合。公开报道中未见完全对应的案件，按课程材料保留。',
    6: '未溯源。未检索到对应报道，按课程材料保留。',
    7: '已核实。8 名同校同学、逾 11 万澳门元、伪造 3 张转账截图与报道一致。',
}
for i, c in enumerate(DATA['cases'], 1):
    H2('案例 %d　%s' % (i, c['title'].split('：', 1)[1]))
    items = []
    for l in c['source']['links']:
        url = l['url'] if l['url'].startswith('http') else None
        extra = None if url else '站内文件'
        items.append((l['title'], url, extra))
    source_block('', items)
    P('说明：' + CASE_GRADE[i], size=9, left=0.75, after=8)

# ---------------------------------------------------------------- 类型
TYPE_CASE = [
    '2025 年 10 月，一名在澳就读的香港女大学生被冒充内地公安者以「视频办案」诈骗 800 万港元。',
    '2022 年 9 月，20 岁本地男大生刷单损失 148,532 澳门元；2023 年 7 月，女事主损失 118,000 澳门元。',
    '2023 年 10 月，本澳中年女子在网恋对象推荐下投资虚拟货币，损失逾 28 万港元。',
    '2024 年 1 月，两宗练功券案的三名事主共失逾 74 万人民币。',
    '2023 年 9 月，杨姓中介以「年付九折」收租，约 40 个单位、210 万港元，租客以在澳就读的内地大学生为主。',
    '2023 年 3 月，三宗游戏账号／皮肤交易骗案，三名学生合共损失逾 6 万澳门元。',
    '2026 年 9 月 18 日至 10 月 6 日，25 人因假冒抖音客服被骗，损失逾 540 万澳门元。',
    '2024 年 3 月，司警接获本地银行检举，指社交平台出现假冒该行的贷款广告与钓鱼链接。',
    '2024 年 3 月警情通告（假贷款广告、钓鱼网站）与 2023 年 3 月假购物网站案。',
    '站内以「视频通话借钱」为主；法规层面已有深度合成标识与同意要求。',
]

H1('二、诈骗类型来源')
for i, t in enumerate(DATA['types'], 1):
    H2('类型 %d　%s' % (i, t['short']))
    name = t['paragraphs'][0].replace('类型名称：', '')
    P('站内定义：' + name, size=9.5, left=0.55, after=3)
    P('来源：', size=9.5, left=0.55, after=1)
    for j, s in enumerate(t['sources'], 1):
        p = P('%d. %s' % (j, s['title']), size=9.5, left=1.0, first=-0.4, after=1)
        u = P('', size=8.5, left=1.45, after=3)
        link(u, s['url'], s['url'], size=8.5)
    P('本地案例：' + TYPE_CASE[i - 1], size=9.2, left=0.55, after=8)

# ---------------------------------------------------------------- 官方电话
H1('三、官方电话与链接')
phones = [
    ('司法警察局报案热线：993（24 小时）', 'https://www.pj.gov.mo/Web/Policia/hotline/'),
    ('司法警察局防诈骗查询热线：8800 7777（24 小时，只查询不报案）',
     'https://www.pj.gov.mo/Web/Policia/crime05.html?lang=en'),
    ('澳门大学保安紧急热线：8822 4000',
     'https://sfs.cmdo.um.edu.mo/our-services/security-services/'),
    ('蔡继有书院办公室：8822 9400', 'https://ckyc.rc.um.edu.mo/'),
    ('蔡继有书院保安值班：6332 4946',
     'https://ckyc.rc.um.edu.mo/student-life/student-caring-services/'),
    ('蔡继有书院邮箱：ckycollege@um.edu.mo', 'https://ckyc.rc.um.edu.mo/'),
    ('司法警察局反诈程式介绍（经微信公众号 PJ_Macao 进入）',
     'https://www.gov.mo/zh-hant/news/1056801/'),
]
for name, url in phones:
    P('· ' + name, size=9.5, left=0.55, first=-0.35, after=1)
    u = P('', size=8.5, left=1.0, after=3)
    link(u, url, url, size=8.5)
P('补充（官方页面列出，站内未显示）：澳大保安及设施服务处 8822 4126；'
  '书院学术人员当值 6392 9984；澳大学生事务部 8822 4901 / sao@um.edu.mo；'
  '紧急求助 999 / 110 / 112。', size=9, left=0.55, after=8)

# ---------------------------------------------------------------- 法规
H1('四、法规与政策')
laws = [
    ('第 20/2024 号法律《打击不法赌博犯罪法》，2024-10-29 生效；新增「经营为赌博的不法汇兑」罪，最高 5 年。',
     'https://www.gov.mo/zh-hant/news/1102576/'),
    ('第 16/2012 号法律《房地产中介业务法》，2013-07-01 生效；中介须持准照、签书面合同。',
     'https://www.gov.mo/zh-hant/services/ps-1911/'),
    ('《互联网信息服务深度合成管理规定》，2023-01-10 施行；人脸替换、合成人声须显著标识。',
     'https://www.gov.cn/zhengce/zhengceku/2022-12/12/content_5731431.htm'),
]
for text, url in laws:
    P('· ' + text, size=9.5, left=0.55, first=-0.35, after=1)
    u = P('', size=8.5, left=1.0, after=3)
    link(u, url, url, size=8.5)

# ---------------------------------------------------------------- 视频
H1('五、宣传视频')
P('四段视频均为课程材料提供的原始 MOV，经 scripts/prepare_videos.py 转为 H.264/AAC MP4，保留原声。', size=9.5, after=3)
for v, dur, res, size in [
    ('video-1.mp4', '1:33.8', '1280×704', '14 MB'),
    ('video-2.mp4', '1:08.4', '1280×704', '10 MB'),
    ('video-3.mp4', '0:54.9', '1280×704', '4.7 MB'),
    ('video-4.mp4', '1:16.0', '1280×706', '14 MB'),
]:
    P('· %s　时长 %s　%s　%s　materials/videos/%s' % (v, dur, res, size, v),
      size=9.5, left=0.55, first=-0.35, after=2)
P('视频封面：poster-1～4.jpg，取自视频首帧。原始素材与出镜授权请在公开发布前确认。', size=9, left=0.55, after=8)

# ---------------------------------------------------------------- 图片
H1('六、图片与标志')
P('· 类型封面 10 张、案例封面 7 张：均为宣传视频截帧，存放于 materials/covers。', size=9.5, left=0.55, first=-0.35, after=2)
P('　 尺寸：类型封面 1376×768（type-09/10 为 2730×1536）；案例封面 2730×1536（case-04 为 2730×650）。', size=9, left=0.9, after=3)
P('· 视频封面 4 张：800×450，取自视频首帧。', size=9.5, left=0.55, first=-0.35, after=2)
P('· 澳门大学标志 2 个（242×89）、蔡继有书院标志 2 个（161×181）：由课程材料提供，'
  '澳门大学标志深蓝底已抠成透明。使用须符合学校视觉识别规范。', size=9.5, left=0.55, first=-0.35, after=8)

# ---------------------------------------------------------------- B站
H1('七、B 站延伸视频（27 条）')
P('每条均为外链，网站不转载、不下载。核对日期 2026-10-30。', size=9, after=4)
n = 0
for t in DATA['types']:
    for b in t['bili']:
        n += 1
        P('%d. %s｜%s｜%s｜%s' % (n, t['short'], b['title'], b['author'], b['duration']),
          size=9.3, left=0.75, first=-0.45, after=1)
        u = P('', size=8.3, left=1.2, after=3)
        link(u, b['url'], b['url'], size=8.3)

# ---------------------------------------------------------------- AI
H1('八、AI 问答服务')
P('· 接口地址：https://api.deepseek.com/chat/completions', size=9.5, left=0.55, first=-0.35, after=2)
P('· 模型：deepseek-chat；temperature 0.3；max_tokens 900；检索片段上限 5。', size=9.5, left=0.55, first=-0.35, after=2)
P('· 代理：Cloudflare Worker（worker/），前端只填 proxyUrl；未部署时 AI 按钮提示不可用。', size=9.5, left=0.55, first=-0.35, after=2)
P('· 密钥：materials/ai-config.js 中为空；密钥只应存于 Worker 环境变量 DEEPSEEK_API_KEY。', size=9.5, left=0.55, first=-0.35, after=2)
P('· 发送时机：仅点击「AI 提问」时，发送问题与站内检索片段；关键词筛选在本地完成。', size=9.5, left=0.55, first=-0.35, after=8)

# ---------------------------------------------------------------- 技术
H1('九、技术依赖')
P('原生 HTML、CSS、JavaScript，无框架、无构建步骤；无 CDN、无第三方字体、无统计脚本；'
  '除 AI 提问外无外部请求。材料以 textContent 渲染，不执行内容中的 HTML。', size=9.5, after=8)

# ---------------------------------------------------------------- 备注
H1('十、备注')
P('1. 案例 4–6 为课程材料《案例.docx》改编，未找到完全对应的公开报道；如获原始通报或链接，可直接替换。', size=9.3, after=2)
P('2. 2026-11-05 核查发现 materials/ai-config.js 曾写入明码 DeepSeek API key，已清空；'
  '旧 key 请到 platform.deepseek.com/api_keys 删除并重新生成。', size=9.3, after=2)
P('3. 链接如失效，可按标题在对应媒体站内检索；B 站视频失效时按标题替换。', size=9.3, after=2)

# ---------------------------------------------------------------- 页脚页码
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(fp, '第 ', size=8)
for name in ('PAGE',):
    r = fp.add_run()
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'begin')
    i = OxmlElement('w:instrText'); i.set(qn('xml:space'), 'preserve'); i.text = ' ' + name + ' '
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
    r._r.append(f); r._r.append(i); r._r.append(e)
run(fp, ' 页', size=8)

cp = doc.core_properties
cp.title = '防诈宣传资料来源清单'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
