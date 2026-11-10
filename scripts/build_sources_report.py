# -*- coding: utf-8 -*-
"""生成《防诈宣传资料来源与核查报告.docx》。

数据来源：materials/content.json（站内现文）+ 2026-11-05 的公开报道核查记录。
运行：python scripts/build_sources_report.py
"""
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'materials' / 'content.json').read_text(encoding='utf-8'))
OUT = ROOT / '防诈宣传资料来源与核查报告_详细版.docx'

BLUE = '123A63'
MID = '1F4E79'
GREY = '595959'
LIGHT = 'F3F6FA'
WARN = 'C00000'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
W = 16.6


# ---------------------------------------------------------------- 基础工具
def set_cn(run, font='宋体'):
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn('w:eastAsia'), font)


def add_run(p, text, size=10.5, bold=False, color=None, font='宋体', italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    set_cn(r, font)
    return r


def add_hyperlink(paragraph, url, text, size=8.8, color=MID):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('r:id'), r_id)
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Calibri')
    rf.set(qn('w:hAnsi'), 'Calibri')
    rf.set(qn('w:eastAsia'), '宋体')
    rpr.append(rf)
    c = OxmlElement('w:color'); c.set(qn('w:val'), color); rpr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rpr.append(sz)
    r.append(rpr)
    t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve')
    r.append(t)
    hl.append(r)
    paragraph._p.append(hl)


def shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    tcpr.append(shd)


def shade_para(p, fill, border=True):
    ppr = p._p.get_or_add_pPr()
    if border:
        bdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single'); left.set(qn('w:sz'), '18')
        left.set(qn('w:space'), '8'); left.set(qn('w:color'), MID)
        bdr.append(left)
        ppr.append(bdr)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    ppr.append(shd)


def hrule(color=MID, size=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), color)
    bdr.append(bottom)
    ppr.append(bdr)


def P(text='', size=10.5, bold=False, color=None, align=None, before=0, after=4,
      indent=0, first=0, line=1.45, font='宋体', keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if indent:
        pf.left_indent = Cm(indent)
    if first:
        pf.first_line_indent = Cm(first)
    if align is not None:
        p.alignment = align
    if keep:
        pf.keep_with_next = True
    if text:
        add_run(p, text, size=size, bold=bold, color=color, font=font)
    return p


def H1(text, number=None, page_break=True):
    p = doc.add_paragraph(style='Heading 1')
    pf = p.paragraph_format
    if page_break:
        pf.page_break_before = True
    pf.space_before = Pt(6); pf.space_after = Pt(10); pf.keep_with_next = True
    label = (number + '　') if number else ''
    add_run(p, label + text, size=17, bold=True, color=BLUE, font='黑体')
    hrule('B7C7D9', 6)
    return p


def H2(text, number=None):
    p = doc.add_paragraph(style='Heading 2')
    pf = p.paragraph_format
    pf.space_before = Pt(14); pf.space_after = Pt(6); pf.keep_with_next = True
    label = (number + '　') if number and not text.startswith(str(number)) else ''
    add_run(p, label + text, size=13, bold=True, color=MID, font='黑体')
    return p


def H3(text):
    p = doc.add_paragraph(style='Heading 3')
    pf = p.paragraph_format
    pf.space_before = Pt(9); pf.space_after = Pt(4); pf.keep_with_next = True
    add_run(p, text, size=10.8, bold=True, color=BLUE, font='黑体')
    return p


def bullet(text, size=10.2, level=0, color=None, bold_lead=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.55 + level * 0.55)
    pf.first_line_indent = Cm(-0.36)
    pf.space_before = Pt(0); pf.space_after = Pt(2.5); pf.line_spacing = 1.35
    add_run(p, '• ' if level == 0 else '– ', size=size, color=MID)
    if bold_lead:
        add_run(p, bold_lead, size=size, bold=True, color=color or BLUE)
    if text:
        add_run(p, text, size=size, color=color)
    return p


def quote(text, cite):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.5); pf.right_indent = Cm(0.3)
    pf.space_before = Pt(4); pf.space_after = Pt(2); pf.line_spacing = 1.35
    add_run(p, text, size=9.3, color='333333')
    shade_para(p, 'F4F6F9')
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c.paragraph_format.space_after = Pt(8)
    add_run(c, '—— ' + cite, size=8.5, color=GREY)


def cell_write(cell, value, size=8.8, bold=False, color=None, align=None):
    cell.text = ''
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(1); pf.space_after = Pt(1); pf.line_spacing = 1.12
    if align is not None:
        p.alignment = align
    parts = value if isinstance(value, list) else [value]
    first = True
    for part in parts:
        if isinstance(part, tuple) and part[0] == 'link':
            add_hyperlink(p, part[1], part[2], size=size)
        else:
            for j, line in enumerate(str(part).split('\n')):
                if not first and j == 0:
                    p.add_run().add_break()
                add_run(p, line, size=size, bold=bold, color=color)
        first = False


def table(headers, rows, widths, size=8.8, head_size=8.8, zebra=True, aligns=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblpr = t._tbl.tblPr
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed'); tblpr.append(layout)
    tblw = OxmlElement('w:tblW')
    tblw.set(qn('w:w'), str(int(sum(widths) * 567)))
    tblw.set(qn('w:type'), 'dxa')
    tblpr.append(tblw)
    for i, w in enumerate(widths):
        t.columns[i].width = Cm(w)
    # 单元格内边距
    mar = OxmlElement('w:tblCellMar')
    for side, v in (('top', 40), ('left', 80), ('bottom', 40), ('right', 80)):
        e = OxmlElement('w:' + side); e.set(qn('w:w'), str(v)); e.set(qn('w:type'), 'dxa'); mar.append(e)
    tblpr.append(mar)

    hdr = t.rows[0]
    trpr = hdr._tr.get_or_add_trPr()
    th = OxmlElement('w:tblHeader'); trpr.append(th)
    cant = OxmlElement('w:cantSplit'); trpr.append(cant)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade(c, MID)
        cell_write(c, h, size=head_size, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)

    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        trpr = t.rows[-1]._tr.get_or_add_trPr()
        trpr.append(OxmlElement('w:cantSplit'))
        for i, val in enumerate(row):
            c = cells[i]
            c.width = Cm(widths[i])
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if zebra and ri % 2 == 1:
                shade(c, LIGHT)
            al = aligns[i] if aligns else None
            cell_write(c, val, size=size, color=None, align=al)

    if len(rows) <= 12:
        for row in t.rows[:-1]:
            for c in row.cells:
                for para in c.paragraphs:
                    para.paragraph_format.keep_with_next = True
    return t


def caption(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    add_run(p, text, size=9, bold=True, color=GREY)


def note(text, size=8.8):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(8)
    add_run(p, text, size=size, color=GREY)


def short_label(title):
    """把“媒体 + 日期 + 标题”的完整来源压成可点击的短标签。"""
    m = re.search(r'\d{4}-\d{2}-\d{2}', title)
    date = (' ' + m.group(0)) if m else ''
    if title.startswith('《'):
        end = title.find('》')
        return title[:end + 1] if end > 0 else title
    for mark in ('《', '〈'):
        if mark in title:
            prefix = title.split(mark)[0].strip()
            return prefix if (m and m.group(0) in prefix) else prefix + date
    return title


def link_line(items, size=8.6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(3)
    add_run(p, '来源链接：', size=size, bold=True, color=GREY)
    for i, (title, url) in enumerate(items):
        if i:
            add_run(p, '；', size=size, color=GREY)
        add_hyperlink(p, url, title, size=size)


# ---------------------------------------------------------------- 页眉页脚
def setup_header_footer():
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]
    hp.paragraph_format.space_after = Pt(2)
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(W), WD_ALIGN_PARAGRAPH.RIGHT)
    add_run(hp, '防诈宣传资料来源与核查报告', size=8.5, color=GREY)
    add_run(hp, '\tCPED2001 第 11 组 · 2026-11-05', size=8.5, color=GREY)
    ppr = hp._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom'); b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '4')
    b.set(qn('w:space'), '4'); b.set(qn('w:color'), 'BFBFBF')
    bdr.append(b); ppr.append(bdr)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(fp, '第 ', size=8.5, color=GREY)
    _field(fp, 'PAGE')
    add_run(fp, ' 页 / 共 ', size=8.5, color=GREY)
    _field(fp, 'NUMPAGES')
    add_run(fp, ' 页', size=8.5, color=GREY)


def _field(paragraph, name):
    r = paragraph.add_run()
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' ' + name + ' '
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    r._r.append(fld); r._r.append(instr); r._r.append(fld2)
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor.from_string(GREY)


# ---------------------------------------------------------------- 封面
def cover():
    for _ in range(3):
        P('', after=0)
    P('澳門大學 · 蔡繼有書院', size=12, color=MID, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    P('CPED2001 第 11 組 · Scam Prevention Campaign', size=10.5, color=GREY,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=18)
    hrule(BLUE, 12)
    P('', after=30)
    P('防诈宣传资料来源', size=30, bold=True, color=BLUE,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=2, font='黑体')
    P('与核查报告', size=30, bold=True, color=BLUE,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=16, font='黑体')
    P('7 个案例、10 类诈骗手法与全部外部信息的来源整理', size=12.5, color=GREY,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
    P('Sources & Verification Report', size=10.5, color='8EA9C1',
      align=WD_ALIGN_PARAGRAPH.CENTER, after=60)
    hrule('B7C7D9', 6)

    info = table(
        ['项目', '内容'],
        [
            ['报告对象', 'types.html、faq.html、help.html 所用材料，以及 materials/content.json、content.js、ai-config.js'],
            ['核查范围', '7 个案例 · 10 类诈骗 · 官方电话与网址 · 法规 · 媒体报道 · 视频与图片素材 · B 站延伸视频 27 条 · AI 问答'],
            ['核查日期', '2026 年 11 月 5 日（所有链接当日访问）'],
            ['版本', 'v1.0（对应网站内容 2026-11-05 修订）'],
            ['编制', 'CPED2001 第 11 组（资料来源整理）'],
        ],
        [3.0, 13.0], size=9
    )
    P('', after=24)
    P('说明：本报告只记录可公开核查的来源与比对结果，不构成法律意见；案件细节以司法机关公布为准。',
      size=8.8, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)


# ---------------------------------------------------------------- 目录
def toc():
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.paragraph_format.space_after = Pt(10)
    add_run(p, '目　录', size=18, bold=True, color=BLUE, font='黑体')
    hrule('B7C7D9', 6)
    items = [
        ('第一章', '文档说明', '目的、范围、方法与分级标准'),
        ('第二章', '核查结果概览', '案例与类型总表、8 项关键发现'),
        ('第三章', '案例来源逐条核查', '7 个案例的原始报道、原文摘录、事实对照与修订'),
        ('第四章', '诈骗类型来源', '10 类手法的权威依据与本地案例'),
        ('第五章', '外部信息清单', '电话、法规、报道、视频、图片、B 站、AI 服务'),
        ('第六章', '材料修订记录', '本次对内容、界面与配置的具体改动'),
        ('第七章', '风险与后续工作', '溯源、时效、版权、隐私与密钥'),
        ('附录 A', '参考文献总表', '[1]–[25] 完整文献信息与链接'),
        ('附录 B', '链接核查记录', '2026-11-05 逐条访问结果'),
    ]
    for num, title, desc in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(1)
        add_run(p, num + '　', size=10.5, bold=True, color=MID)
        add_run(p, title, size=11.5, bold=True, color=BLUE)
        q = doc.add_paragraph()
        q.paragraph_format.left_indent = Cm(1.9)
        q.paragraph_format.space_after = Pt(4)
        add_run(q, desc, size=9.2, color=GREY)
    note('各章独立起页；文中 [n] 对应附录 A 的参考文献编号。')


# ---------------------------------------------------------------- 第一章
def chapter1():
    H1('文档说明', '第一章')
    P('本报告是 CPED2001 第 11 组「Scam Prevention Campaign」防诈宣传网站的配套材料，'
      '用来回答三个问题：网站上的 7 个案例各自来自哪里；10 类诈骗手法的定义与本地案例依据是什么；'
      '网站引用的所有外部信息是否可核查。', first=0.74)

    H2('1.1　核查范围', '1.1')
    bullet('案例：types.html 案例轨道与 7 个案例详情页的全部文字、金额、出处。')
    bullet('类型：types.html 10 个类型详情页的定义、“套路”“受害人群”与延伸视频。')
    bullet('外部信息：官方电话与网址、法规与政策、媒体报道、4 段宣传视频、封面与标志、B 站延伸视频 27 条、AI 问答服务、网站技术依赖。')
    bullet('不包含：课程材料《案例.docx》《诈骗类型.docx》的创作过程；未公开的司警内部资料；任何个人姓名与联系方式。')

    H2('1.2　方法与判定标准', '1.2')
    P('核查按以下五步进行，每一步的结果都可以在后续章节复现。', after=6)
    table(
        ['步骤', '做法'],
        [
            ['① 提取关键词', '从每个案例提取金额、币种、日期、地点、手法、受害人身份（如“543,464 元人民币”“路氹酒店客房”“练功券”）。'],
            ['② 检索', '在《澳门日报》电子版、力报、广州日报大洋网、澳门特区政府入口网站（gov.mo）、司法警察局新闻稿与政府公报中检索；优先采用原始报道而非转载。'],
            ['③ 逐字比对', '把站内文字与报道逐项对照：金额是否精确到个位、币种是否相同、受害人身份是否一致、资金流向与话术是否一致、处理结果是否一致。'],
            ['④ 链接复核', '2026-11-05 逐条访问链接，只有返回 HTTP 200 的页面才写入报告；B 站因反爬限制单独说明。'],
            ['⑤ 分级', '按 A/B/C 三级标注一致性，并据此决定“更正”“保留并标注改编”或“标为未溯源”。'],
        ],
        [3.0, 13.6], size=9
    )
    caption('表 1　核查步骤')

    P('', after=4)
    table(
        ['等级', '定义', '处理方式'],
        [
            ['A 已核实', '金额、手法、地点等关键事实与原始报道一致，可逐项对应。', '按报道校正细节，详情页附报道链接。'],
            ['B 部分吻合', '能找到高度相近的报道，但受害人身份、平台、币种等细节不同。', '保留课程材料设定，标注“课程材料改编”，并列出相近报道。'],
            ['C 未溯源', '未检索到对应的公开报道，仅见于课程材料。', '保留但明确标注“未检索到完全对应报道”，建议团队补充来源或改为综合案例。'],
        ],
        [2.6, 7.4, 6.6], size=9
    )
    caption('表 2　一致性分级')

    H2('1.3　使用限制', '1.3')
    bullet('媒体报道是对司法案件的转述，不等同于法院判决；部分案件在报道时仍在侦查，后续可能有变化。')
    bullet('金额一律以报道原文为准；站内的人民币折算数字是按汇率估算，已在正文标明“按汇率折算”。')
    bullet('澳门媒体多用繁体，本报告引用时保留原文用字，叙述部分用简体。')
    bullet('本报告不复制任何 API 密钥；涉及密钥的位置只描述风险，不复述内容。')

    H2('1.4　报告结构', '1.4')
    bullet('第二章：一页看懂 7 个案例与 10 类诈骗的来源结构，并列 8 项关键发现。')
    bullet('第三章：逐条给出每个案例的站内现文、原始来源、报道原文、事实对照与修订结论。')
    bullet('第四章：说明 10 类手法的定义依据、本地案例与编辑建议。')
    bullet('第五章：把电话、法规、报道、视频、图片、B 站、AI 服务、技术依赖一次列全。')
    bullet('第六章与第七章：记录本次改了什么、还剩什么风险与后续工作。')
    bullet('附录 A 为 [1]–[25] 参考文献，附录 B 为 2026-11-05 的链接核查记录。')


# ---------------------------------------------------------------- 第二章
def chapter2():
    H1('核查结果概览', '第二章')
    P('一句话结论：7 个案例中，4 个可以逐项对应到原始报道（A），2 个只有相近报道（B），1 个未找到来源（C）；'
      '10 类诈骗均有官方通稿、法规或本地报道支撑；网站引用的官方电话与网址全部有效；'
      '另发现 1 处密钥泄露风险与 3 处案例事实错误，已在本次修订中处理。', first=0.74)

    H2('2.1　案例核查总表', '2.1')
    rows = [
        ['C1', '冒充公检法，港籍女生失 800 万', '《澳门日报》2025-10-26；《广州日报》2025-10-27', 'A', '无实质差异；报道未提伪造证件、办案背景'],
        ['C2', '内地生租房，中介卷租金', '《澳门日报》2023-09-02；力报 2023-09-01', 'A', '损失原稿 87,300 应为 160,900 港元'],
        ['C3', '跨境换钱，练功券冒充港币', '《澳门日报》2024-01-12', 'A', '受害人不是大学生，是两名中年非本地男子'],
        ['C4', '刷单兼职，失 12 万澳门元', '《澳门日报》2023-07-07（相近）', 'B', '11.8 万而非 12 万；本地女青年；佣金 8%'],
        ['C5', '杀猪盘，失 28 万澳门元', '《澳门日报》2023-10-24（相近）', 'B', '28 万港元；受害人为中年女子，非男大学生'],
        ['C6', '游戏账号，充值保证金', '《澳门日报》2023-03-17（相近）', 'C', '未找到对应报道；手法为假截图骗号'],
        ['C7', '校内群组换钱，8 人被骗', '《澳门日报》2024-04-24', 'A', '原稿“2 名研究生、4.2 万”有误，应为 8 人、逾 11 万'],
    ]
    table(['编号', '案例（简）', '原始来源', '等级', '主要差异'],
          rows, [1.2, 3.7, 5.9, 1.1, 4.7], size=8.6)
    caption('表 3　7 个案例的来源与一致性')

    H2('2.2　类型来源总表', '2.2')
    rows = [
        ['T1 冒充公检法', '司警局 2020-11-28 通稿；2025 年两报报道', '800 万港元案', '3'],
        ['T2 刷单兼职', '《澳门日报》2022-09-03、2023-07-07', '14.8 万、11.8 万澳门元案', '2'],
        ['T3 杀猪盘', '司警局 2020-11-28 网恋诈骗案；2023-10-24 报道', '28 万港元案', '2'],
        ['T4 私下换汇', '第 20/2024 号法律；司警局 2024-11-05、11-18 通稿', '练功券案 543,464 元', '3'],
        ['T5 租房诈骗', '第 16/2012 号法律；房屋局；2023-09-02 报道', '40 租户、210 万港元案', '2'],
        ['T6 游戏交易', '《澳门日报》2023-03-17（三宗游戏骗案）', '合共逾 6 万澳门元', '1'],
        ['T7 仿冒客服', '力报 2026-10-07；《澳门日报》2023-03-17', '25 人、540 万澳门元', '2'],
        ['T8 贷款诈骗', '司警局 2024-03-25 警情通告', '假冒银行贷款广告、钓鱼网站', '1'],
        ['T9 钓鱼网站', '司警局 2024-03-25 警情通告；2023-03-17 报道', '假购物网站、假贷款页', '2'],
        ['T10 AI 换脸', '《互联网信息服务深度合成管理规定》', '2023-01-10 施行', '1'],
    ]
    table(['类型', '主要权威依据', '本地案例佐证', '条数'],
          rows, [3.2, 7.3, 4.9, 1.2], size=8.6)
    caption('表 4　10 类诈骗的来源结构')

    H2('2.3　关键发现', '2.3')
    findings = [
        ('F1　4 个案例可精确对应。', '案例 1、2、3、7 的金额、手法、地点均能在《澳门日报》找到对应报道；案例 3 的 543,464 元人民币与报道逐位一致。'),
        ('F2　案例 2 损失金额错误。', '原稿 87,300 港元与报道的三名学生合计 160,900 港元相差近一倍，且与正文“十几万”自相矛盾。'),
        ('F3　案例 3 受害人身份错误。', '报道明确两名男事主“非本地人，同是中年”，并非“两名澳门高校男大学生”；事主是当场检查发现假钞，不是离开后才发现。'),
        ('F4　案例 7 案情方向相反。', '原稿写“学生先转人民币、等港币”，报道是同学转澳门元/人民币、等人民币；受害人数与金额也由 2 人 4.2 万更正为 8 人逾 11 万。'),
        ('F5　案例 4、5、6 只能标为改编。', '三者均未找到完全对应的公开报道，只能以“相近报道”佐证；如坚持写成真实个案，需要补充司警通报或采访记录。'),
        ('F6　官方联络信息有效，另有补充空间。', '站内 5 个电话、7 个官方链接全部有效；官方页面另有 8822 4126、6392 9984、8822 4901 与紧急号码 999/110/112，建议补充。'),
        ('F7　B 站链接数量与 README 不符。', '实际为 27 条（3+3+3+2+3+3+2+3+2+3），README 原写 25 条，已更正。'),
        ('F8　发现明码 API 密钥。', 'materials/ai-config.js 原含一枚可直接调用的 DeepSeek key，与 README“密钥为空”的说明不符；已清空，旧 key 须吊销。'),
    ]
    for lead, text in findings:
        bullet(text, bold_lead=lead)

    H2('2.4　本次修订摘要', '2.4')
    bullet('更正 3 个案例的关键事实（案例 2 金额、案例 3 受害人身份与发现方式、案例 7 受害人数与金额）。')
    bullet('为 7 个案例补齐原始报道链接与核对说明；案例 4–6 标注为课程材料改编。')
    bullet('为 10 个类型新增 sources 字段，详情页底部显示“出处”行。')
    bullet('把页面上的“真实案例”统一改为“案例拆解”，避免对未溯源案例作真实性承诺。')
    bullet('清空 ai-config.js 中的 API 密钥，并同步更新 README 与核对日期。')


# ---------------------------------------------------------------- 第三章
def src_rows(items):
    rows = []
    for name, desc, url, link_text in items:
        rows.append([name, desc])
        if url and url.startswith('http'):
            rows.append([name + '链接', ('link', url, link_text or url)])
        elif url:
            rows.append([name + '文件', url])
    rows.append(['访问日期', '2026-11-05（链接当日返回 HTTP 200）'])
    return rows


def case_section(no, grade_label, sources, quote_text, cite, compare, conclusions):
    c = DATA['cases'][no - 1]
    H2('案例 %d：%s' % (no, c['title'].split('：', 1)[1]), '3.%d' % no)
    p = P('一致性：', size=9.5, after=6)
    add_run(p, grade_label, size=9.5, bold=True, color=(WARN if grade_label.startswith('C') or grade_label.startswith('B') else MID))

    H3('3.%d.1　站内现文' % no)
    body = [x for x in c['paragraphs']
            if x not in ('案情', '诈骗套路重点', '套路重点', '防范提示（网页展示用）', '防范提示')]
    q = P('　'.join(body), size=9.2, color='333333', after=8)
    q.paragraph_format.left_indent = Cm(0.3)

    H3('3.%d.2　原始来源' % no)
    table(['项目', '内容'], src_rows(sources), [3.2, 13.4], size=8.6)

    H3('3.%d.3　报道原文摘录' % no)
    quote(quote_text, cite)

    H3('3.%d.4　事实对照' % no)
    table(['项目', '站内材料（修订前）', '原始报道', '判断'], compare, [2.4, 4.6, 6.6, 3.0], size=8.5)

    H3('3.%d.5　结论与修订' % no)
    for x in conclusions:
        bullet(x, size=9.8)


def chapter3():
    H1('案例来源逐条核查', '第三章')
    P('本章按网站顺序逐条核对 7 个案例。每条包含站内现文、原始来源、报道原文摘录、事实对照与修订结论。'
      '引用保留报道原文的繁体用字，“……”表示省略。', first=0.74)

    case_section(
        1, 'A 已核实（两家报道情节一致）',
        [
            ('第一来源', '《澳门日报》2025-10-26，A02 版《港女生墮「公檢法」失八百萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html', 'content_442031.html'),
            ('第二来源', '《广州日报》2025-10-27 09:05，记者曾毅、陈庆辉《又是「内地公安」骗局，澳门大学生被骗八百万港元》（大洋网）',
             'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm', 'news.dayoo.com/…54888366.htm'),
            ('补充', '司法警察局 2025-10 月接报调查，暂未有人被捕；报道未公布受害人姓名', None, None),
        ],
        '一名在本澳就讀大學的香港女大學生疑墮「公檢法」騙局，誤信「內地公安」指涉及詐騙，最終被騙去八百萬港元。……報稱今年八月上旬收到「內地執法人員」來電，指其內地手機號涉嫌發放大量詐騙短訊，電話隨後轉線至「公安」接聽，對方要求事主接受調查並視頻辦案，過程中要求交付四百萬港元「保證金」，並教唆事主以到外地升學為由，向父母索款。……對方又稱案情嚴重，要求再繳付多四百萬港元「保證金」。女事主再次相信，再從自己賬戶匯款四百萬港元到對方指定的香港銀行賬戶。……香港銀行通知女事主的家人指出匯款有可疑……本月廿三日向司警報案，報稱合共損失八百萬港元。',
        '《澳门日报》2025-10-26，A02 版',
        [
            ['受害人', '香港籍女大学生，在澳读大学', '香港居民，在澳读大学', '一致'],
            ['案发时间', '2025 年 8 月', '今年 8 月上旬', '一致'],
            ['索款方式', '向父母要 400 万 + 自己存款 400 万', '先向父母索 400 万，再从自己账户汇 400 万', '基本一致'],
            ['银行与报案', '银行监测异常跨境转账后联系父母', '银行通知家人汇款可疑；10 月 23 日报案', '一致，补日期'],
            ['损失', '800 万港元（约 733 万人民币）', '800 万港元', '金额一致，人民币为折算'],
            ['未见于报道', '伪造证件、假办案背景、分批转账', '—', '已从正文删去'],
        ],
        [
            '金额、时间、手法、报案经过与两家报道对应，判定 A。',
            '修订：把“自己名下存款 400 万”改为“再从自己账户汇出 400 万”；补充“2025 年 10 月 23 日报案”。',
            '修订：“银行监测到大额异常跨境转账”改为报道的“银行通知家人汇款有可疑”。',
            '报道未提及“伪造公安证件、假办案背景、分批转账”，站内已删去，避免把常见手法写成该案事实。',
            '人民币 733 万为折算数字，已标明“按汇率折算”。',
        ]
    )

    case_section(
        2, 'A 已核实（金额已更正）',
        [
            ('第一来源', '《澳门日报》2023-09-02，A09 版《中介呃租客預繳租金二百萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html', 'content_286009.html'),
            ('第二来源', '力报 2023-09-01《本澳一地產經紀呃租金去炒股 40租戶被騙逾210萬港元》',
             'https://exmoo.com/article/225209.html', 'exmoo.com/article/225209.html'),
        ],
        '上月廿九及三十日，司警先後接獲三名（十九至廿一歲）就讀本澳某大學的內地學生舉報……有人聲稱「一次過繳交一年租金可獲享九折優惠」……三人共合損失十六萬○九百港元。另外兩名事主……合共損失兩個月租金及按金合共五萬四千四百港元。……楊某承認……詐騙約四十個單位租客大約二百一十萬港元。全部騙款已用於個人股票投資，已經為大部分租客交回大約一百五十萬港元租金……還有十五個單位、涉及六十萬港元租金仍未交還。',
        '《澳门日报》2023-09-02，A09 版',
        [
            ['受害人', '3 名内地学生（合租，19–21 岁）', '3 名内地学生（19 至 21 岁），同校', '一致'],
            ['诱饵', '一次交一年租金打 9 折', '「一次過繳交一年租金可獲享九折優惠」', '一致'],
            ['损失', '原稿 87,300 港元', '三名学生合共 160,900 港元', '已更正'],
            ['其他受害人', '—', '另 2 名事主合共 54,400 港元', '已补充'],
            ['涉案总额', '约 40 个单位、约 210 万港元', '約四十個單位、約二百一十萬港元', '一致'],
            ['赃款去向', '炒股投资', '全部用于个人股票投资', '一致'],
            ['处理', '报警', '2023-08-31 拘捕，涉相当巨额诈骗、巨额信任之滥用', '已补充'],
        ],
        [
            '40 个单位、约 210 万港元、160,900 港元三个数字均可核对，判定 A。',
            '修订：损失由 87,300 改为 160,900 港元，并补充另两名事主 54,400 港元；卡片金额由“8.7 万”改为“16 万”。',
            '原稿正文“十几万”与报道一致，但损失行数字错误，属内部矛盾，已消除。',
            '案发时间改为“2023 年（租约自 2022 年 6 月起）”，与报道的租期一致。',
        ]
    )

    case_section(
        3, 'A 已核实（金额与手法；受害人身份已更正）',
        [
            ('第一来源', '《澳门日报》2024-01-12，A08 版《三男女練功券詐騙被捕》',
             'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html', 'content_313501.html'),
            ('补充', '同版报道另一起练功券案；两案共检获 400 张练功券', None, None),
        ],
        '上周五傍晚六時半，兩事主以社交軟件相約一名「換錢黨」到路氹某酒店客房兌換金錢。經協議，五十四萬三千四百六十四元人民幣可兌換五十八萬港元。……有人便將兩疊、二百張面額為一千港元的鈔票交予事主。事主檢查時發現全為「練功券」，通知酒店保安員代為報警。……司警表示此案兩名男事主為非本地人，同是中年。',
        '《澳门日报》2024-01-12，A08 版',
        [
            ['受害人', '两名澳门高校男大学生', '两名非本地男子，同是中年', '已更正'],
            ['金额', '543,464 元人民币', '543,464 元人民币兑 580,000 港元', '一致（精确到个位）'],
            ['地点', '路氹酒店房间', '路氹某酒店客房', '一致'],
            ['发现方式', '离开后才发现', '事主检查时发现，通知酒店保安报警', '已更正'],
            ['处理', '报警', '拘捕一男一女；另案拘一名男子，两案共检 400 张练功券', '已补充'],
        ],
        [
            '543,464 元人民币、路氹酒店客房、两叠 200 张千元港币均与报道一致，就金额与手法判定 A。',
            '受害人身份不符：报道为两名中年非本地男子，不是“两名澳门高校男大学生”。',
            '修订：受害人改为“两名非本地男子（中年）”；发现方式改为“事主检查时发现，通知酒店保安员报警”；标题加“练功券冒充港币”。',
            '建议：可保留为“换钱党”类型的本地案例，但不宜作为“学生受害”案例宣传。',
        ]
    )

    case_section(
        4, 'B 部分吻合（课程材料改编）',
        [
            ('课程材料', '《案例.docx》案例 4（站内文件）', 'materials/案例.docx', None),
            ('相近报道', '《澳门日报》2023-07-07，A05 版《兩男女兼職刷單被騙逾十二萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html', 'content_274031.html'),
        ],
        '兩名本地青年墮入兼職刷單騙局……女事主上月三十日在社交平台看到招聘廣告，與對方聯絡後參與刷單。對方要求事主先墊支貨款，隨後會聯同百分之八的佣金一併退回。……報稱合共損失十一萬八千澳門元。男事主被騙情節與女事主大致相同，被騙七千澳門元後向司警報案。',
        '《澳门日报》2023-07-07，A05 版',
        [
            ['受害人', '在澳女大学生（原“内地”）', '两名本地青年；女事主', '部分吻合'],
            ['金额', '120,000 澳门元', '118,000 + 7,000 澳门元', '接近，非同一案'],
            ['佣金', '约 8%（原 8%–10%）', '8%', '按报道改为 8%'],
            ['平台', '社交平台（原 Instagram）', '社交平台（未指明）', '已改为泛指'],
            ['结论', '课程材料改编', '—', '未见完全对应报道'],
        ],
        [
            '未找到与“120,000 澳门元、内地女大学生、Instagram、8%–10%”完全对应的报道，判定 B。',
            '相近报道：女事主损失 118,000 澳门元、佣金 8%，受害人为本地青年。',
            '修订：删去“内地”“Instagram”，佣金改为“约 8%”，标题改为“在澳女大学生”；保留 120,000（课程材料数字），出处注明“课程材料改编”。',
            '建议：如能找到原始司警通报，再决定是否把金额改为 118,000。',
        ]
    )

    case_section(
        5, 'B 部分吻合（课程材料改编）',
        [
            ('课程材料', '《案例.docx》案例 5（站内文件）', 'materials/案例.docx', None),
            ('相近报道', '《澳门日报》2023-10-24，A06 版《中年女網投被騙廿八萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html', 'content_296925.html'),
        ],
        '一名本澳女子疑墮網戀詐騙，被男網友遊說投資虛擬貨幣，但最終無法提現，損失逾二十八萬港元。……案中的本澳女事主是中年人，報稱本年七月下旬在網上認識一名男網友，未幾雙方發展成為線上情侶關係。……分多次轉帳合共十八萬港元，均顯示獲利，但無法提現。經詢問男網友後，對方稱需要支付十萬元「按金」才可提現……報稱損失超過二十八萬港元。',
        '《澳门日报》2023-10-24，A06 版',
        [
            ['受害人', '在澳男大学生（原“澳门大学内地”）', '本澳中年女子', '不一致'],
            ['金额', '280,000 澳门元', '逾 28 万港元', '数字相近、币种不同'],
            ['手法', '网恋 + 虚假投资网页 + 缴税提现', '网恋 + 虚拟货币 App + 10 万“按金”', '同属杀猪盘，细节不同'],
            ['结论', '课程材料改编', '—', '未见完全对应报道'],
        ],
        [
            '金额 28 万与 2023-10-24 报道相近，但报道为中年女子、港元、虚拟货币 App、10 万按金，判定 B。',
            '修订：删去“澳门大学”“内地”，改为“在澳就读的男大学生”；出处注明“课程材料改编”并列出相近报道。',
            '建议：若坚持“男大学生”设定，需要另找来源；否则应标为综合案例。',
        ]
    )

    case_section(
        6, 'C 未溯源（课程材料，标注为改编）',
        [
            ('课程材料', '《案例.docx》案例 6（站内文件）', 'materials/案例.docx', None),
            ('相近报道', '《澳门日报》2023-03-17，A05 版《七人墮電信詐騙失逾六十萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html', 'content_250629.html'),
        ],
        '三案涉及買賣遊戲帳戶或皮膚套裝騙案。受害人是二十歲本地男大學生、十八歲本地女子，來澳讀大學的十九歲內地女生，均接獲騙徒聲稱要購買他們的遊戲帳戶或皮膚套裝，之後騙徒指已轉帳並傳來轉帳截圖，三人信以為真，將遊戲帳戶或皮膚套裝過戶對方後，才發現沒有收到分毫，合共損失折合逾六萬澳門元。',
        '《澳门日报》2023-03-17，A05 版',
        [
            ['受害人', '澳门本地男大学生', '三宗：本地男大生、本地女子、内地女生', '类型相同、案件不同'],
            ['金额', '36,000 澳门元', '三宗合共逾 6 万澳门元', '不一致'],
            ['手法', '仿冒交易网站 + 保证金解冻', '假转账截图骗走账号/皮肤', '不同'],
            ['结论', '课程材料，未溯源', '—', '建议标为综合案例'],
        ],
        [
            '未找到对应报道；相近报道为假转账截图骗走账号/皮肤，与站内“仿冒平台充值保证金”手法不同，判定 C。',
            '修订：出处注明“课程材料；未检索到完全对应报道”，并列出相近报道。',
            '建议：改为“综合案例”，或补充司警关于游戏点卡/账号交易的通报。',
        ]
    )

    case_section(
        7, 'A 已核实（案情已更正）',
        [
            ('第一来源', '《澳门日报》2024-04-24，A01 版《大學生換錢呃同學11萬》',
             'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html', 'content_333259.html'),
            ('补充', '司警 2024-04-22 在北区截获嫌疑人，案件涉巨额诈骗、电脑伪造移送检察院', None, None),
        ],
        '本地十九歲「爛賭」大學生因欠債做「換錢黨」，瘋狂搵同學換人民幣，收錢後以「賬戶凍結、額度已滿」藉口拖延，八名同學被騙逾十一萬澳門元。……八名事主當中，一名本地人、六名內地人及一名南美人，與湯某就讀同一大學。……湯某在今年一月至四月期間，在大學聊天群組物色需要換錢的同學，訛稱可兌換人民幣。……湯某則以當日已夠額度無法轉賬、賬戶遭凍結等藉口拖延，並偽造三張轉賬截圖以博取事主信任。……合共損失逾十一萬澳門元。',
        '《澳门日报》2024-04-24，A01 版',
        [
            ['受害人', '8 名同校同学（1 本地、6 内地、1 南美）', '同左', '一致（已更正）'],
            ['金额', '逾 11 万澳门元', '逾 11 万澳门元', '一致'],
            ['人数', '8 人', '8 人', '一致（原稿 2 人）'],
            ['手法', '群组招揽、假截图、额度/冻结借口', '同左', '一致'],
            ['处理', '2024-04-22 截获，涉巨额诈骗、电脑伪造', '同左', '一致'],
        ],
        [
            '8 名受害人、逾 11 万、伪造 3 张截图、2024-04-22 拘捕均与报道一致，判定 A。',
            '修订：受害人由“2 名内地研究生”改为 8 名同校同学；损失由 4.2 万改为逾 11 万；资金方向改为“先转澳门元/人民币、等人民币”；删去“WhatsApp”，改为“大学聊天群组”。',
            '原稿与报道差异最大，修订后与报道一致；卡片金额改为“11 万”。',
        ]
    )

    H2('3.8　未完全溯源案例的处理建议', '3.8')
    P('案例 4、5、6 的共同问题是：课程材料只给出叙述，没有原始链接、日期或通报编号。'
      '在无法补证的情况下，最稳妥的做法不是删除，而是把“真实个案”改成“改编案例”，并在页面保留相近报道，让读者知道哪些是报道事实、哪些是教学化处理。', after=6)
    table(['案例', '可核对的相近报道', '不能核对的设定', '建议'],
          [
              ['C4', '2023-07-07 女事主 118,000 澳门元、佣金 8%', '内地身份、Instagram、120,000 整数', '标注改编；如获原始通报再改回'],
              ['C5', '2023-10-24 中年女子逾 28 万港元', '男大学生、澳门元、外汇网页、缴税提现', '标注改编或改写为综合案例'],
              ['C6', '2023-03-17 三宗游戏账号骗案', '3.6 万、仿冒平台保证金', '标为综合案例'],
          ], [1.4, 5.6, 5.4, 4.2], size=8.6)
    caption('表 5　未完全溯源案例的处理建议')


# ---------------------------------------------------------------- 第四章
TYPE_BASIS = [
    [
        '司警局 2020-11-28 通稿记录一宗假冒「公检法」案：冒充社工局职员→转「北京市公安」→指账户涉清洗黑钱→套取银行资料，损失逾 257 万港元；通稿列出「切勿尽信来电显示」等 9 条对策 [10]。',
        '2025 年 10 月澳门再发生同类案件，损失 800 万港元，「视频办案」「教唆向家人索款」「要求保密」等话术仍在用 [1][2]。',
    ],
    [
        '2022-09-03 报道：20 岁本地男大生刷单损失 148,532 澳门元；首单 1,300 元连本带利返还 1,429 元，随后被要求「多刷单才可取回本利」[25]。',
        '2023-07-07 报道：女事主损失 118,000 澳门元，佣金 8%，被要求「再做四次任务」[6]。',
    ],
    [
        '司警局 2020-11-28 网恋诈骗案：社交平台认识→线上情侣→投资虚拟货币 App→账户冻结→损失 33 万人民币 [10]。',
        '2023-10-24 报道：本澳中年女子损失逾 28 万港元，提现前被要求付 10 万「按金」[7]。',
    ],
    [
        '第 20/2024 号法律《打击不法赌博犯罪法》2024-10-29 生效，新增「经营为赌博的不法汇兑」罪，最高 5 年徒刑；用于不法汇兑的现金、筹码可被扣押；司警呼吁在合法金融机构兑换 [11][12]。',
        '2024-01-12 报道：两宗练功券案、三名事主共失逾 74 万人民币，检获 400 张练功券 [5]。',
    ],
    [
        '第 16/2012 号法律《房地产中介业务法》2013-07-01 生效，房屋局为主管实体；中介须持准照、在营业场所张贴准照与商业营业场所说明书、与客户签订书面中介合同 [14]。',
        '2023-09-01/02 报道：杨姓中介以「年付九折」收租约 210 万港元，约 40 个单位，租客以在澳就读的內地大学生为主 [3][4]。',
    ],
    [
        '2023-03-17 报道：三宗游戏账号/皮肤交易骗案，受害人为 20 岁本地男大生、18 岁本地女子、19 岁内地女生，均被假转账截图骗走账号，合共损失逾 6 万澳门元 [8]。',
    ],
    [
        '2026-10-07 力报：9 月 18 日至 10 月 6 日共 25 人报案，损失逾 540 万澳门元；骗子冒充抖音客服，称服务到期将自动续费，诱导共享屏幕、提供银行卡资料 [17]。',
        '2023-03-17 报道：假冒快递公司短讯，点链接输入信用卡资料后被盗刷 [8]。',
    ],
    [
        '司警局 2024-03-25 警情通告：社交平台出现假冒本地银行的「最新不良信用贷款条件」广告，链接跳转钓鱼网站；呼吁勿信网上融资广告、勿在不明网站输入个人资料、慎防以放款条件为名索取「借贷保证金」[15]。',
    ],
    [
        '同上通告：假贷款广告→钓鱼网站，是本地最常见的钓鱼入口之一 [15]。',
        '2023-03-17 报道：登入假购物网站购买音响，转账后收不到货 [8]。',
    ],
    [
        '《互联网信息服务深度合成管理规定》2023-01-10 施行：提供人脸、人声等生物识别信息编辑功能，应提示使用者告知被编辑人并取得单独同意（第 14 条）；人脸替换、合成人声等须显著标识（第 17 条）；不得删除、篡改标识（第 18 条）[18]。',
    ],
]

TYPE_CASE = [
    '2025 年 10 月，一名在澳就读的香港女大学生被冒充内地公安者以「视频办案」诈骗 800 万港元 [1][2]。',
    '2022 年 9 月与 2023 年 7 月，澳门分别有男大生损失 148,532 澳门元、女事主损失 118,000 澳门元的刷单案 [25][6]。',
    '2023 年 10 月，本澳中年女子在网恋对象推荐下投资虚拟货币，损失逾 28 万港元 [7]。',
    '2024 年 1 月，两宗练功券案的三名事主共失逾 74 万人民币；2024-10-29 起不法汇兑入刑 [5][11]。',
    '2023 年 9 月，杨姓中介以「年付九折」收租，约 40 个单位、210 万港元，租客以在澳就读的內地大学生为主 [3][4]。',
    '2023 年 3 月，三宗游戏账号/皮肤交易骗案，三名学生合共损失逾 6 万澳门元 [8]。',
    '2026 年 9 月 18 日至 10 月 6 日，25 人因假冒抖音客服被骗，损失逾 540 万澳门元 [17]。',
    '2024 年 3 月，司警接获本地银行检举，指社交平台出现假冒该行的贷款广告与钓鱼链接 [15]。',
    '2024 年 3 月警情通告与 2023 年 3 月假购物网站案，均属钓鱼网站/链接类型 [15][8]。',
    '目前站内以“视频通话借钱”为主；法规层面已有深度合成标识与同意要求 [18]。',
]

TYPE_ADVICE = [
    '可补一句“视像笔录不是法定程序，澳门和内地公职机关不会通过电话或视频要求转账”。',
    '保留“联单”“垫付”两个关键词，它们是最容易被忽略的转折点。',
    '强调“小额可提现”是养信任，不是平台真实；提现前的税费/按金是收网信号。',
    '区分“为赌博的不法汇兑”（刑事，最高 5 年）与一般私下换汇的民事、外汇风险。',
    '详情页已补“核验中介准照”一句；可再加“租金直接付业主并保留收据”。',
    '可与站内“仿冒平台 + 保证金”案例并列，说明“假截图骗号”与“假平台骗钱”两种变体。',
    '可补“取消订阅/自动续费”“共享屏幕”两个话术，它们是目前最高频的入口。',
    '把“验资、保证金、解冻费”写进类型说明，与刷单、游戏交易的话术打通。',
    '补充“学校系统/邮箱登录页”被仿冒的本地情境，比通用钓鱼更贴近学生。',
    '可加“视频通话也可能是合成的，回拨核实”；并说明标识要求不等于每条内容都有标识。',
]


def chapter4():
    H1('诈骗类型来源', '第四章')
    P('10 类诈骗按网站顺序排列。每类先列站内现文，再列权威依据与本地案例，最后给出编辑建议。'
      '类型不是司法罪名，而是防骗宣传中的手法分类，因此来源以官方通稿、法规与本地报道为主。', first=0.74)
    for i, t in enumerate(DATA['types']):
        H2(t['short'], '4.%d' % (i + 1))
        H3('站内定义')
        for para in t['paragraphs']:
            bullet(para, size=9.6)
        H3('来源与依据')
        for b in TYPE_BASIS[i]:
            bullet(b, size=9.6)
        H3('本地案例佐证')
        P(TYPE_CASE[i], size=9.8, indent=0.5)
        H3('编辑建议')
        P(TYPE_ADVICE[i], size=9.8, indent=0.5)
        link_line([(short_label(s['title']), s['url']) for s in t.get('sources', [])])


# ---------------------------------------------------------------- 第五章
def chapter5():
    H1('外部信息清单', '第五章')
    P('本章把网站涉及的“外界信息”一次列全：谁能提供帮助、依据哪条法规、报道在哪里、视频和图片从哪来、'
      'B 站看了什么、AI 问答把数据发到哪里、网站依赖什么。凡标注“站内：是”的，均已在页面出现。', first=0.74)

    H2('5.1　官方机构与联络方式', '5.1')
    table(['机构 / 服务', '号码或邮箱', '用途', '站内', '核对', '来源'],
          [
              ['司警局报案热线', '993', '24 小时报案', '是', '有效', '[19]'],
              ['司警局防诈骗查询热线', '8800 7777', '24 小时查询；2017-08-17 设立，只查询不报案', '是', '有效', '[20]'],
              ['澳大保安紧急热线', '8822 4000', '校内紧急求助', '是', '有效', '[21]'],
              ['澳大保安及设施服务处', '8822 4126', '门禁、CCTV、无人机申请', '否', '有效', '[21]'],
              ['蔡继有书院办公室', '8822 9400', '书院事务与转介', '是', '有效', '[22]'],
              ['蔡继有书院保安值班', '6332 4946', '书院内即时协助', '是', '有效', '[23]'],
              ['书院学术人员当值', '6392 9984', '学生关怀当值', '否', '有效', '[23]'],
              ['澳大学生事务部', '8822 4901 / sao@um.edu.mo', '学生支援', '否', '有效', '[24]'],
              ['紧急求助', '999 / 110 / 112', '警察、消防、救护', '否', '有效', '[23]'],
              ['书院邮箱', 'ckycollege@um.edu.mo', '一般查询', '是', '有效', '[22]'],
          ], [3.9, 3.6, 5.0, 1.1, 1.2, 1.8], size=8.5)
    caption('表 6　官方联络方式（2026-11-05 核对）')
    note('注：993 是司警报案热线，8800 7777 只作查询；999/110/112 为紧急求助。站内已收 5 个号码，建议把 8822 4126、6392 9984、8822 4901 与紧急号码补进 help.html。')

    H2('5.2　法律法规与政策文件', '5.2')
    table(['文件', '公布', '生效', '与本网站的关系', '来源'],
          [
              ['第 20/2024 号法律《打击不法赌博犯罪法》', '2024 年', '2024-10-29', '新增「经营为赌博的不法汇兑」罪，最高 5 年；现金、筹码可扣押；支撑“换钱党”类型', '[11][12][13]'],
              ['第 16/2012 号法律《房地产中介业务法》', '2012 年', '2013-07-01', '中介须持准照、张贴准照、签书面合同；支撑“租房诈骗”的核验建议', '[14]'],
              ['《互联网信息服务深度合成管理规定》', '2022-11-25', '2023-01-10', '人脸/人声编辑须告知并取得单独同意；合成内容须显著标识；支撑“AI 换脸”类型', '[18]'],
          ], [5.2, 2.2, 2.2, 5.6, 1.4], size=8.5)
    caption('表 7　法规与政策')

    H2('5.3　媒体报道', '5.3')
    media = [
        ['[1]', '澳门日报', '2025-10-26', 'A02', '港女生墮「公檢法」失八百萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html', '案例 1、T1'],
        ['[2]', '广州日报', '2025-10-27', '—', '又是「内地公安」骗局，澳门大学生被骗八百万港元',
         'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm', '案例 1 二次来源'],
        ['[3]', '澳门日报', '2023-09-02', 'A09', '中介呃租客預繳租金二百萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html', '案例 2、T5'],
        ['[4]', '力报', '2023-09-01', '—', '本澳一地產經紀呃租金去炒股 40租戶被騙逾210萬港元',
         'https://exmoo.com/article/225209.html', '案例 2、T5'],
        ['[5]', '澳门日报', '2024-01-12', 'A08', '三男女練功券詐騙被捕',
         'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html', '案例 3、T4'],
        ['[6]', '澳门日报', '2023-07-07', 'A05', '兩男女兼職刷單被騙逾十二萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html', '案例 4、T2'],
        ['[7]', '澳门日报', '2023-10-24', 'A06', '中年女網投被騙廿八萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html', '案例 5、T3'],
        ['[8]', '澳门日报', '2023-03-17', 'A05', '七人墮電信詐騙失逾六十萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html', '案例 6、T6/T7/T9'],
        ['[9]', '澳门日报', '2024-04-24', 'A01', '大學生換錢呃同學11萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html', '案例 7'],
        ['[17]', '力报', '2026-10-07', '—', '司警：提防抖音假客服詐騙 三周接獲25人報案失逾540萬',
         'https://www.exmoo.com/article/266439.html', 'T7'],
        ['[25]', '澳门日报', '2022-09-03', 'A04', '男大生墮刷單騙局失逾十四萬',
         'https://appimg.modaily.cn/app/szb/pc/content/202209/03/content_212804.html', 'T2'],
    ]
    rows = []
    for no, m, d, ed, title, url, use in media:
        rows.append([no, m, d, ed, title, use, ('link', url, '打开')])
    table(['编号', '媒体', '日期', '版次', '标题', '用途', '链接'],
          rows, [1.1, 1.7, 1.9, 1.0, 6.3, 2.9, 1.7], size=8.2)
    caption('表 8　媒体报道一览')

    H2('5.4　宣传视频', '5.4')
    table(['文件', '时长', '分辨率', '大小', '处理与来源', '站内用途'],
          [
              ['video-1.mp4', '1:33.8', '1280×704', '14 MB', '原始 MOV 经 scripts/prepare_videos.py 转 H.264/AAC，保留原声', '宣传视频'],
              ['video-2.mp4', '1:08.4', '1280×704', '10 MB', '同上', '宣传视频'],
              ['video-3.mp4', '0:54.9', '1280×704', '4.7 MB', '同上', '宣传视频'],
              ['video-4.mp4', '1:16.0', '1280×706', '14 MB', '同上；分辨率比其余三段高 2 px', '宣传视频'],
              ['poster-1～4.jpg', '—', '1280×704/706', '35–69 KB', '视频首帧，作播放前封面', '视频封面'],
          ], [3.2, 1.8, 2.4, 1.8, 5.2, 2.2], size=8.5)
    caption('表 9　4 段宣传视频')
    note('来源：课程材料提供的原始 MOV；README 记录“四段 MOV 已转换为 H.264/AAC MP4，保留原有声音并生成封面”。公开发布前须确认视频素材与出镜者的使用授权。')

    H2('5.5　图片与标志', '5.5')
    table(['类别', '数量', '尺寸', '来源', '授权提示'],
          [
              ['类型封面', '10 张', '1376×768（type-09/10 为 2730×1536）', '宣传视频截帧', '随视频授权；尺寸不统一，建议统一导出'],
              ['案例封面', '7 张', '2730×1536（case-04 为 2730×650）', '宣传视频截帧', '同上；case-04 为超宽裁切'],
              ['视频封面', '4 张', '800×450', '宣传视频截帧', '同上'],
              ['澳门大学标志', '2 个 PNG', '242×89', '课程材料提供；深蓝底已抠透明', '使用须符合澳大视觉识别规范'],
              ['蔡继有书院标志', '2 个 PNG', '161×181', '课程材料提供，含白色版', '使用须符合书院规范'],
          ], [3.2, 1.8, 4.4, 3.6, 3.6], size=8.5)
    caption('表 10　图片素材')
    note('封面均为视频截帧，不是新闻图片，因此不涉及转载媒体照片的版权问题；但视频本身的授权仍需确认。')

    H2('5.6　B 站延伸视频（27 条）', '5.6')
    rows = []
    n = 0
    for t in DATA['types']:
        for b in t['bili']:
            n += 1
            bv = b['url'].rstrip('/').split('/')[-1]
            rows.append([str(n), t['short'], b['title'], b['author'], b['duration'],
                         ('link', b['url'], bv), '10-30'])
    table(['#', '类型', '标题', 'UP 主', '时长', '链接', '核对'],
          rows, [0.9, 2.0, 6.2, 2.8, 1.3, 2.2, 1.2], size=7.9)
    caption('表 11　B 站延伸视频 27 条')
    note('核对说明：27 条由项目组于 2026-10-30 逐条打开确认标题、UP 主与时长；本次自动复核因 B 站反爬限制未能重跑，建议每学期人工复核一次并替换失效链接。')

    H2('5.7　AI 问答的外部服务', '5.7')
    table(['项目', '配置 / 说明', '位置'],
          [
              ['接口地址', 'https://api.deepseek.com/chat/completions', 'ai-config.js'],
              ['代理', 'Cloudflare Worker（worker/），前端只填 proxyUrl；未部署时 AI 按钮提示不可用', 'worker/README.md'],
              ['模型', 'deepseek-chat', 'ai-config.js'],
              ['参数', 'temperature 0.3；max_tokens 900；检索片段上限 5', 'ai-config.js'],
              ['密钥', '现为空；密钥只应存于 Worker 环境变量 DEEPSEEK_API_KEY', '本次核查'],
              ['发送时机', '仅点击「AI 提问」时，发送问题 + 站内检索片段', 'assets/js/main.js'],
              ['降级', '请求失败时展示站内最相关问答，页面仍可用', 'assets/js/main.js'],
              ['隐私边界', '关键词筛选在本地完成；AI 回答由大模型生成，仅供参考', '页面说明'],
          ], [3.4, 8.6, 4.6], size=8.5)
    caption('表 12　AI 问答配置')
    note('风险：2026-11-05 核查发现 ai-config.js 曾写入一枚明码 DeepSeek key（与 README“apiKey 为空”不符）。该 key 已从文件中清空；由于它曾以明文存在于工作区，必须到 DeepSeek 开放平台删除并重新生成，不能只删文件。')

    H2('5.8　技术依赖与离线可用性', '5.8')
    bullet('原生 HTML、CSS、JavaScript，无框架、无构建步骤，双击 index.html 即可打开。')
    bullet('无 CDN、无第三方字体、无统计脚本；除「AI 提问」外，页面不发起任何外部请求。')
    bullet('视频、图片、文字均为本地文件，断网时除 AI 外均可使用。')
    bullet('材料以 textContent 渲染，不执行内容中的 HTML；搜索高亮使用字面匹配，不作为正则执行。')
    bullet('支持键盘操作（方向键切换、Esc 返回）与手机响应式布局。')


# ---------------------------------------------------------------- 第六章
def chapter6():
    H1('材料修订记录', '第六章')
    P('本章记录 2026-11-05 对站点与材料文件的实际改动，便于回溯与课堂说明。', first=0.74)

    H2('6.1　内容修订（materials/content.json / content.js）', '6.1')
    table(['位置', '修订前', '修订后', '依据'],
          [
              ['案例 1　索款与汇款', '向父母要 400 万 + 自己名下存款 400 万', '先向父母索 400 万，再从自己账户汇 400 万', '[1][2]'],
              ['案例 1　报案经过', '银行监测异常跨境转账后联系父母', '银行通知家人汇款可疑；2025-10-23 报案', '[1]'],
              ['案例 1　细节', '伪造证件、假办案背景、分批转账', '删去（报道未提及）', '[1][2]'],
              ['案例 2　损失', '87,300 港元；卡片 8.7 万', '160,900 港元（另 2 人 54,400）；卡片 16 万', '[3][4]'],
              ['案例 2　时间', '2023 年', '2023 年（租约自 2022 年 6 月起）', '[3]'],
              ['案例 3　受害人', '两名澳门高校男大学生', '两名非本地男子（中年）', '[5]'],
              ['案例 3　发现方式', '离开后才发现', '事主检查时发现，通知酒店保安报警', '[5]'],
              ['案例 3　标题', '跨境换钱骗局', '跨境换钱骗局（练功券冒充港币）', '[5]'],
              ['案例 4　身份/平台/佣金', '内地女大学生；Instagram；8%–10%', '在澳女大学生；社交平台；约 8%', '[6]'],
              ['案例 5　身份', '澳门大学内地男大学生', '在澳就读的男大学生', '[7]'],
              ['案例 6　出处', '原始报道待核实', '课程材料；未检索到完全对应报道 + 相近报道', '[8]'],
              ['案例 7　受害人与金额', '2 名内地研究生；4.2 万澳门元', '8 名同校同学；逾 11 万澳门元', '[9]'],
              ['10 类诈骗', '无出处字段', '新增 sources（1–3 条）并在详情页显示', '本报告第四、五章'],
              ['T4 提醒', '私下换汇不受法律保护，银行卡冻结', '补第 20/2024 号法律：不法汇兑入刑，最高 5 年', '[11][12]'],
              ['T5 提醒', '无', '补第 16/2012 号法律：核验准照、书面合同', '[14]'],
              ['T10 提醒', '无', '补深度合成规定：人脸/人声编辑与显著标识', '[18]'],
              ['首页与案例区', '“真实案例”“7 个真实案例拆解”', '“案例拆解”“7 个案例拆解”', '本报告 2.3'],
              ['字段', 'casesNote / casesVerified / verified', '更新为 2026-11-05 与新的核对说明', '本次核查'],
          ], [4.0, 5.0, 5.2, 2.4], size=8.3)
    caption('表 13　内容修订')

    H2('6.2　界面、配置与文档', '6.2')
    table(['文件', '修改', '原因'],
          [
              ['assets/js/main.js', '类型详情页新增「出处」行（读取 sources）', '让类型来源在页面可见'],
              ['assets/js/main.js', '“真实案例”改为“案例拆解”；AI 检索标签改为“案例”', '避免对未溯源案例作真实性承诺'],
              ['materials/ai-config.js', 'apiKey 清空，补充安全注释', '消除明文密钥泄露风险'],
              ['README.md', '更新核对日期、案例说明、来源列表；B 站 25→27；补充报告与脚本', '与实际情况一致'],
              ['scripts/apply_source_revisions.py', '新增', '记录本次来源修订，便于复现'],
              ['scripts/build_sources_report.py', '新增', '生成本报告'],
          ], [4.6, 7.2, 4.8], size=8.5)
    caption('表 14　界面与文档修改')

    H2('6.3　未改动但需团队确认', '6.3')
    bullet('案例 4、5、6 的原始出处：是否还有课程老师的原始链接、司警通报编号或截图？若有，可把 B/C 升为 A。')
    bullet('案例 4 的金额：保留 120,000（课程材料）还是改为 118,000（报道）？建议以能找到的原始通报为准。')
    bullet('案例 5 的“男大学生”设定：若无来源，建议改为综合案例，避免把中年女性案件改写成学生案件。')
    bullet('视频与校徽授权：课程内部展示与公开发布的授权范围可能不同，发布前向书院/澳大确认。')
    bullet('B 站 27 条链接的时效：建议每学期复核，失效链接及时替换。')


# ---------------------------------------------------------------- 第七章
def chapter7():
    H1('风险与后续工作', '第七章')

    H2('7.1　案例 4–6 的溯源', '7.1')
    P('这三个案例目前只能标注为“课程材料改编”。如果期末展示要求“真实案例”，建议优先补齐案例 4 与案例 5 的司警通报，'
      '案例 6 改为“综合案例”并在页面写明“由多种游戏交易骗局综合而成”。这样做不会削弱宣传效果，反而能避免被追问出处时无法回答。', first=0.74)

    H2('7.2　链接时效', '7.2')
    bullet('《澳门日报》电子版链接含日期与稿件编号，稳定性较好；但旧稿可能改版，建议保留标题与日期，便于重新检索。')
    bullet('力报、广州日报链接为新闻页，长期有效概率较高；若失效，可用标题在站内搜索。')
    bullet('B 站视频可能下架或换稿；27 条建议每学期复核一次，并在 content.json 中整条替换。')
    bullet('官方电话可能调整；本站已写明“以官方页面为准”，建议每学期按 5.1 表重核。')

    H2('7.3　版权与授权', '7.3')
    bullet('校徽：澳门大学与蔡继有书院标志属官方标识，使用须符合视觉识别规范；公开网站建议先取得书院同意。')
    bullet('视频：4 段宣传视频来自课程材料，需确认拍摄/出镜/配乐授权；公开发布前不要默认“课程内部可用 = 可公开”。')
    bullet('封面：均为视频截帧，随视频授权；不涉及媒体新闻图片。')
    bullet('B 站：网站只外链、不转载、不下载，风险较低；但不应嵌入播放器绕过平台。')

    H2('7.4　隐私与合规', '7.4')
    bullet('案例只使用公开报道信息，不写全名；报道中的姓氏（如“汤某”“杨某”）在站内已省略。')
    bullet('AI 提问只在点击时发送“问题 + 站内片段”，不上传姓名、学号、电话；页面已作说明。')
    bullet('若 Worker 记录日志，应只保留必要字段并设置保留期限；不要把问题内容长期留存。')
    bullet('模拟帖子已标明“模拟内容”，不应伪装成真实学生发帖。')

    H2('7.5　密钥与代理（最高优先级）', '7.5')
    p = P('', after=4)
    p.paragraph_format.keep_together = True
    add_run(p, '必须立即处理：', size=10.5, bold=True, color=WARN)
    add_run(p, 'materials/ai-config.js 曾以明文保存一枚 DeepSeek API key。文件已清空，但密钥一旦以明文存在过，'
               '就应按“已泄露”处理。请到 platform.deepseek.com/api_keys 删除该 key 并重新生成，'
               '新 key 只放进 Cloudflare Worker 的环境变量 DEEPSEEK_API_KEY；不要把新 key 写回 ai-config.js、截图或聊天记录。', size=10.2)
    bullet('Worker 的 ALLOWED_ORIGINS 目前是占位域名（zhuolin272-dotcom.github.io）与本地地址，上线前改成实际域名。')
    bullet('若课程演示需要本地直连，可临时填入新 key，演示结束立即清空，并确认没有上传到任何公开位置。')

    H2('7.6　数据一致性', '7.6')
    bullet('content.js 是网页实际读取的文件，content.json 是备份；两者必须同步（运行 scripts/sync_content.py）。')
    bullet('封面尺寸不统一（1376×768、2730×1536、800×450、2730×650），建议统一导出规格，减少加载与裁切问题。')
    bullet('案例卡片金额、详情页损失、short 字段三处需一致；本次修订后已对齐。')


# ---------------------------------------------------------------- 附录
REFS = [
    ('[1]', '澳門日報. 港女生墮「公檢法」失八百萬[N]. 澳門日報, 2025-10-26(A02).', 'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html'),
    ('[2]', '曾毅, 陳慶輝. 又是「內地公安」騙局，澳門大學生被騙八百萬港元[N/OL]. 廣州日報大洋網, 2025-10-27.', 'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm'),
    ('[3]', '澳門日報. 中介呃租客預繳租金二百萬[N]. 澳門日報, 2023-09-02(A09).', 'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html'),
    ('[4]', '澳門力報. 本澳一地產經紀呃租金去炒股 40租戶被騙逾210萬港元[N/OL]. 2023-09-01.', 'https://exmoo.com/article/225209.html'),
    ('[5]', '澳門日報. 三男女練功券詐騙被捕[N]. 澳門日報, 2024-01-12(A08).', 'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html'),
    ('[6]', '澳門日報. 兩男女兼職刷單被騙逾十二萬[N]. 澳門日報, 2023-07-07(A05).', 'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html'),
    ('[7]', '澳門日報. 中年女網投被騙廿八萬[N]. 澳門日報, 2023-10-24(A06).', 'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html'),
    ('[8]', '澳門日報. 七人墮電信詐騙失逾六十萬[N]. 澳門日報, 2023-03-17(A05).', 'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html'),
    ('[9]', '澳門日報. 大學生換錢呃同學11萬[N]. 澳門日報, 2024-04-24(A01).', 'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html'),
    ('[10]', '司法警察局. 詐騙手法層出不窮 司警局呼籲市民小心防範[EB/OL]. 2020-11-28.', 'https://www.gov.mo/zh-hant/news/354373/'),
    ('[11]', '司法警察局. 司警局與博監局及博企舉行會議 共商配合《打擊不法賭博犯罪法》執法工作[EB/OL]. 2024-11-05.', 'https://www.gov.mo/zh-hant/news/1102576/'),
    ('[12]', '司法警察局. 司警局加強宣傳《打擊不法賭博犯罪法》 促警民合力打擊不法匯兌活動[EB/OL]. 2024-11-18.', 'https://www.gov.mo/zh-hant/news/1106939/'),
    ('[13]', '澳門特別行政區. 第20/2024號法律《打擊不法賭博犯罪法》[Z]. 2024（2024-10-29 生效；本報告經 [11] 引用，未單獨附連結）。', ''),
    ('[14]', '澳門特別行政區政府. 房地產中介業務（第16/2012號法律《房地產中介業務法》）[EB/OL].', 'https://www.gov.mo/zh-hant/services/ps-1911/'),
    ('[15]', '司法警察局. 警情通告：提防假冒本地銀行貸款廣告 警惕釣魚網站以免誤墮騙局[EB/OL]. 2024-03-25.', 'https://www.gov.mo/zh-hans/news/748837/'),
    ('[16]', '司法警察局. 司警局反詐程式正式上線 助公眾防範電訊網絡詐騙[EB/OL]. 2024-04-09.', 'https://www.gov.mo/zh-hant/news/1056801/'),
    ('[17]', '澳門力報. 司警：提防抖音假客服詐騙 三周接獲25人報案失逾540萬[N/OL]. 2026-10-07.', 'https://www.exmoo.com/article/266439.html'),
    ('[18]', '國家互聯網信息辦公室, 工業和信息化部, 公安部. 互聯網信息服務深度合成管理規定[Z]. 2022-11-25 公布, 2023-01-10 施行.', 'https://www.gov.cn/zhengce/zhengceku/2022-12/12/content_5731431.htm'),
    ('[19]', '司法警察局. 熱線電話[EB/OL].', 'https://www.pj.gov.mo/Web/Policia/hotline/'),
    ('[20]', '司法警察局. 防詐騙查詢熱線[EB/OL].', 'https://www.pj.gov.mo/Web/Policia/crime05.html?lang=en'),
    ('[21]', '澳門大學. 保安服務[EB/OL].', 'https://sfs.cmdo.um.edu.mo/our-services/security-services/'),
    ('[22]', '蔡繼有書院. 聯繫方式[EB/OL].', 'https://ckyc.rc.um.edu.mo/'),
    ('[23]', '蔡繼有書院. 學生關懷服務[EB/OL].', 'https://ckyc.rc.um.edu.mo/student-life/student-caring-services/'),
    ('[24]', '澳門大學學生事務部. 司法警察局－學校安全聯網通訊[EB/OL]. 2026-01-27.', 'https://e-bulletin.um.edu.mo/notice/486001/'),
    ('[25]', '澳門日報. 男大生墮刷單騙局失逾十四萬[N]. 澳門日報, 2022-09-03(A04).', 'https://appimg.modaily.cn/app/szb/pc/content/202209/03/content_212804.html'),
]


def appendix():
    H1('参考文献总表', '附录 A')
    P('按正文首次引用顺序编号。所有链接于 2026-11-05 访问；[13] 为法规本身，正文经 [11] 引用，故不单列链接。', after=6)
    rows = []
    for no, info, url in REFS:
        rows.append([no, info, ('link', url, '打开') if url else '—'])
    table(['编号', '文献信息', '链接'], rows, [1.2, 12.9, 2.5], size=8.0)
    note('表中“打开”为超链接，指向对应原文；所有链接均于 2026-11-05 访问。')

    H1('链接核查记录', '附录 B')
    P('下表为 2026-11-05 逐条访问的结果。B 站 27 条未列入：其接口对自动访问返回错误页，本次沿用项目组 2026-10-30 的人工核对记录。', after=6)
    checks = [
        ('《澳门日报》案例 1', 'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html'),
        ('《广州日报》案例 1', 'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm'),
        ('《澳门日报》案例 2', 'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html'),
        ('力报 案例 2', 'https://exmoo.com/article/225209.html'),
        ('《澳门日报》案例 3', 'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html'),
        ('《澳门日报》案例 4', 'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html'),
        ('《澳门日报》案例 5', 'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html'),
        ('《澳门日报》案例 6', 'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html'),
        ('《澳门日报》案例 7', 'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html'),
        ('《澳门日报》T2', 'https://appimg.modaily.cn/app/szb/pc/content/202209/03/content_212804.html'),
        ('司警局 2020-11-28 通稿', 'https://www.gov.mo/zh-hant/news/354373/'),
        ('司警局 2024-11-05 通稿', 'https://www.gov.mo/zh-hant/news/1102576/'),
        ('司警局 2024-11-18 通稿', 'https://www.gov.mo/zh-hant/news/1106939/'),
        ('司警局 2024-04-09 反詐程式', 'https://www.gov.mo/zh-hant/news/1056801/'),
        ('司警局 2024-03-25 警情通告', 'https://www.gov.mo/zh-hans/news/748837/'),
        ('房地产中介业务', 'https://www.gov.mo/zh-hant/services/ps-1911/'),
        ('深度合成管理规定', 'https://www.gov.cn/zhengce/zhengceku/2022-12/12/content_5731431.htm'),
        ('司警局 熱線電話', 'https://www.pj.gov.mo/Web/Policia/hotline/'),
        ('司警局 防詐騙查詢', 'https://www.pj.gov.mo/Web/Policia/crime05.html?lang=en'),
        ('澳大 保安服務', 'https://sfs.cmdo.um.edu.mo/our-services/security-services/'),
        ('蔡繼有書院 首頁', 'https://ckyc.rc.um.edu.mo/'),
        ('蔡繼有書院 學生關懷', 'https://ckyc.rc.um.edu.mo/student-life/student-caring-services/'),
        ('澳大 e-bulletin', 'https://e-bulletin.um.edu.mo/notice/486001/'),
        ('力报 2026-10-07', 'https://www.exmoo.com/article/266439.html'),
    ]
    rows = [[str(i + 1), name, ('link', url, '打开'), '200', '正常'] for i, (name, url) in enumerate(checks)]
    table(['#', '页面', '链接', 'HTTP', '结果'], rows, [0.9, 6.2, 4.2, 1.5, 1.6], size=8.0)
    note('说明：以上 24 条均返回 HTTP 200；另有 B 站 27 条仅作外链，未在本次自动复核。')


def main():
    setup_header_footer()
    cover()
    toc()
    chapter1()
    chapter2()
    chapter3()
    chapter4()
    chapter5()
    chapter6()
    chapter7()
    appendix()

    # 中文排版惯例：表题在表上方。把紧跟在表格后面的“表 X”说明移到表格之前。
    body = doc.element.body
    for p in list(body.iterchildren(qn('w:p'))):
        text = ''.join(t.text or '' for t in p.iter(qn('w:t')))
        prev = p.getprevious()
        if text.startswith('表 ') and prev is not None and prev.tag == qn('w:tbl'):
            prev.addprevious(p)

    cp = doc.core_properties
    cp.title = '防诈宣传资料来源与核查报告'
    cp.subject = 'CPED2001 第 11 组 · Scam Prevention Campaign'
    cp.author = 'CPED2001 第 11 组'
    cp.keywords = '防诈骗；资料来源；澳门；案例核查'
    doc.save(OUT)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
