# -*- coding: utf-8 -*-
"""生成黑白版《小组 Pre 分工与讲稿》。

按 5 人分工组织：开场/结尾、3 大类诈骗、Q&A。内容取自网站 2026-11-05 版。
运行：python scripts/build_pre_script.py
"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'Pre分工与讲稿.docx'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.4)
sec.top_margin = Cm(2.1)
sec.bottom_margin = Cm(1.9)


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


def P(text='', size=10, bold=False, font='宋体', before=0, after=3,
      left=0, first=0, line=1.4, align=None, keep=False):
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


def H1(text):
    return P(text, size=13, bold=True, font='黑体', before=12, after=5, keep=True)


def H2(text):
    return P(text, size=10.8, bold=True, font='黑体', before=8, after=3, keep=True)


def label(text, body, size=10):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.5
    run(p, text, size=size, bold=True)
    run(p, body, size=size)
    return p


def bullet(text, size=10, left=0.4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(left + 0.35)
    pf.first_line_indent = Cm(-0.35)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.4
    run(p, '· ', size=size)
    run(p, text, size=size)
    return p


def rule():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(10)
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '10')
    b.set(qn('w:space'), '1'); b.set(qn('w:color'), '000000')
    bdr.append(b)
    ppr.append(bdr)


# ---------------------------------------------------------------- 标题
P('小组 Pre 分工与讲稿', size=21, bold=True, font='黑体', after=3)
P('CPED2001 第 11 组 · Scam Prevention Campaign', size=9.5, after=2)
P('依据：小组分工（同学 1–5）+ 网站 2026-11-05 版材料', size=9, after=6)
rule()
P('说明：本稿按“开场 → 三大类 → 问题汇总 → 结尾”组织，每人一段，可直接照读或按要点讲。'
  '案例 4、5、6 为课程材料改编，讲稿中已注明口径；所有数字以网站详情页“出处”行为准。', size=9.5, after=4)

# ---------------------------------------------------------------- 分工
H1('一、分工与时间')
P('按 15–20 分钟设计，可按现场压缩。', size=9.5, after=4)
rows = [
    ('同学 1', '开场：主题引入、网站介绍、三大类导览', '2 分钟'),
    ('同学 2', '第一类：冒充官方机构（公检法、客服、钓鱼）', '3 分钟'),
    ('同学 3', '第二类：工作与生活需求（刷单、贷款、租房）', '3 分钟'),
    ('同学 4', '第三类：网络娱乐与社交媒体（杀猪盘、游戏、换钱、AI 换脸）', '4 分钟'),
    ('同学 5', '问题汇总：被骗后报警、止付、面子、账户出借等', '4 分钟'),
    ('同学 1', '结尾：求助电话、四句口号、整体总结', '3 分钟'),
]
for who, what, t in rows:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.35); pf.first_line_indent = Cm(-0.35)
    pf.space_after = Pt(2); pf.line_spacing = 1.4
    run(p, who + '　', size=10, bold=True)
    run(p, what + '　（' + t + '）', size=10)
P('合计约 19 分钟；如限时 15 分钟，每段压缩 1 分钟，Q&A 保留 3 分钟。', size=9, before=3, after=2)

# ---------------------------------------------------------------- 三大类
H1('二、三大类怎么分')
P('网站有 10 类诈骗、7 个案例。展示时压成 3 大类，每类讲 1–2 个代表案例，其余留给 Q&A。', size=9.5, after=4)

H2('A　冒充官方机构（同学 2）')
bullet('对应网站类型：1 冒充公检法、7 仿冒客服、9 钓鱼网站。')
bullet('主案例：案例 1（冒充公检法，800 万港元）。')
bullet('一句话：先用一个你不敢质疑的身份，让你怕、让你急，再让你转账或点链接。')

H2('B　工作与生活需求（同学 3）')
bullet('对应网站类型：2 刷单兼职、8 贷款诈骗、5 租房诈骗。')
bullet('主案例：案例 4（刷单，12 万澳门元）、案例 2（租房，16.09 万港元）。')
bullet('一句话：你正好想赚钱、缺钱或找房，对方就从这个需求进来。')

H2('C　网络娱乐与社交媒体（同学 4）')
bullet('对应网站类型：3 杀猪盘、6 游戏交易、4 私下换汇、10 AI 换脸。')
bullet('主案例：案例 5（杀猪盘，28 万澳门元）、案例 6（游戏，3.6 万澳门元）、案例 7（校内换钱，8 人逾 11 万澳门元）。')
bullet('一句话：信任来自“关系”和“群”，不是平台。')
P('注意：案例 3 的受害人是两名中年非本地男子，不是学生，不适合当学生案例；换钱主题用案例 7。',
  size=9, before=2, after=2)

# ---------------------------------------------------------------- 同学1
H1('三、同学 1：开场、结尾与总结')

H2('开场（2 分钟）')
label('【讲稿】', '大家好，我们是第 11 组，项目叫 Scam Prevention Campaign。开始前先问大家一个问题：'
                   '如果现在有人打电话，说得出你的姓名和学校，说你名下的手机号涉嫌诈骗，要你视频配合调查，你会不会犹豫一下？')
P('这不是假设。2025 年 10 月，澳门司警公布一宗案件：一名在澳门读大学的香港女生，被冒充内地公安的人骗走 800 万港元。'
  '对方让她视频办案、交保证金、不要告诉家人。', size=10, after=4)
P('我们做了一个网站，把澳门学生最常遇到的 10 类骗局、7 个案例、20 个问答和求助电话放在一起。'
  '今天不逐页讲，我们把 10 类压成 3 大类，由三位同学各讲一类，最后留时间回答问题。', size=10, after=4)
P('三类是：一、冒充官方机构；二、利用工作、兼职和生活需求；三、网络娱乐和社交媒体。'
  '先请同学 2 讲第一类。', size=10, after=5)
bullet('开场用问题 + 真实案件数字，不要先讲定义。', size=9.5)
bullet('网站四个页面：首页、认识诈骗类型、常见问题、求助。', size=9.5)

H2('结尾（2 分钟，Q&A 之后）')
label('【讲稿】', '最后，把最重要的电话留给大家。报案打 993；不确定是不是诈骗，打 8800 7777 查询；'
                   '校内紧急情况打 8822 4000；书院办公室 8822 9400，书院保安 6332 4946；'
                   '如果有人受伤或现场危险，打 999。')
P('记不住也没关系，网站“求助”页面都列好了，可以存下来。', size=10, after=4)
P('如果今天只记住一句话，希望是：任何让你保密、先转账、共享屏幕、点陌生链接的，都先停下来，'
  '用自己找到的官方号码核实。', size=10, after=5)

H2('总结（1 分钟）')
label('【讲稿】', '三类骗局看起来不一样，共同点是一样的：骗子先拿到你的信任或恐惧，'
                   '再让你把钱转到一个你控制不了的地方。防的方法也只有一个：停下来，自己核实。')
P('不轻信、不转账、不透露、先核实。谢谢大家。', size=10, bold=True, after=4)

# ---------------------------------------------------------------- 同学2
H1('四、同学 2：第一类　冒充官方机构（3 分钟）')
label('【开场】', '第一类，冒充官方机构。包括冒充公检法、海关、移民局，也包括冒充客服、银行，'
                   '还有伪装成学校系统的钓鱼网站。共同点是：先用一个你不敢质疑的身份，让你怕、让你急。')

H2('案例 1　冒充公检法，800 万港元')
P('2025 年 8 月，一名在澳门读大学的香港女生接到电话，对方说她名下的内地手机号发了大量诈骗短信，'
  '然后把电话转给“内地公安”，要求视频办案。对方要 400 万港元保证金，还教她跟父母说是出国升学需要钱；'
  '得手后又说案情严重，要再交 400 万。她前后共转出 800 万港元。直到香港的银行通知家人汇款可疑，'
  '她才在 10 月 23 日报案。', size=10, after=5)
label('【分析】', '她为什么信？三个点：来电显示、视频里的“办案”场景、还有“不能告诉任何人”的保密要求。'
                   '骗子最厉害的地方不是技术，是让你不敢跟别人核实。')

H2('同类的两个变体')
bullet('假客服：2026 年 9 月 18 日到 10 月 6 日，澳门 25 人因假冒抖音客服被骗，损失超过 540 万澳门元。'
       '套路是“服务到期会自动续费”，然后让你共享屏幕、提供银行卡资料。', size=9.8)
bullet('钓鱼网站：伪装成学校系统或银行登录页，骗账号密码。短信、IG、WhatsApp 群组里的链接都可能是入口。', size=9.8)

H2('防范（一句话收尾）')
P('挂断；用自己查到的官方号码回拨；不确定就打 8800 7777 或 993。'
  '公检法不会电话办案、不会视频做笔录、不会要求转账或保密。', size=10, after=4)
label('【过渡】', '第一类的核心是“怕”。第二类的核心是“想要”——想赚钱、缺钱、找房。请同学 3。')

# ---------------------------------------------------------------- 同学3
H1('五、同学 3：第二类　工作与生活需求（3 分钟）')
label('【开场】', '第二类，利用工作、兼职和生活需求。网站里的刷单兼职、贷款、租房都在这一类。'
                   '共同点：你正好想赚钱、缺钱或找房，对方就从这个需求进来。')

H2('案例 4　刷单兼职，12 万澳门元（课程材料案例）')
P('一名在澳女大学生想找线上兼职，看到“每单约 8% 佣金、本金全额返还”。前两单几百块，连本带利拿回来；'
  '第三单开始是“联单”，必须做完才能提现，不做完本金冻结。她前后投入 12 万澳门元。'
  '想提现时，对方又说任务未完成、账户要激活、要交保证金，最后拉黑。', size=10, after=5)

H2('案例 2　租房，16.09 万港元（真实报道）')
P('2023 年，3 名在澳读大学的内地学生通过社交群组认识一名本澳地产中介，对方说“一次交一年租金打九折”。'
  '他们转了年租、押金差价和中介费，合计损失 16.09 万港元。这个中介用同样方式骗了约 40 个单位、'
  '约 210 万港元，钱拿去炒股，最后是业主上门催租才揭发。', size=10, after=5)
label('【分析】', '共同点：先给你甜头——小额返利、九折——再让你一次性投入更多；等你发现，钱已经不在能退的地方。')
label('【防范】', '刷单要先垫付的，都是诈骗；租房先核验中介准照（第 16/2012 号法律），租金直接付业主、留收据，'
                   '不要一次预付全年；贷款要先交保证金、验资费的，也是诈骗。')
label('【过渡】', '第二类利用的是生活需求。第三类更隐蔽，因为它从恋爱、游戏、同学群里开始。')

# ---------------------------------------------------------------- 同学4
H1('六、同学 4：第三类　网络娱乐与社交媒体（4 分钟）')
label('【开场】', '第三类，网络娱乐和社交媒体。杀猪盘、游戏账号、私下换钱、AI 换脸都在这一类。'
                   '共同点：信任来自“关系”和“群”，不是平台。')

H2('案例 5　杀猪盘，28 万澳门元（课程材料案例）')
P('一名在澳男大学生在交友 App 认识一位“海外金融专业”的女生，发展成网恋。对方推荐一个“有内部漏洞”的投资网页，'
  '小额投入能提现；加大投入后，平台说要交“个税”才能提款，最后网页打不开、对方失联，损失 28 万澳门元。',
  size=10, after=5)

H2('案例 6　游戏账号，3.6 万澳门元（课程材料案例）')
P('一名本地男大学生想卖手游账号，买家让他去一个仿冒交易网站。页面提示“资金冻结，需要充值保证金解冻”，'
  '他充了 3 次，共 3.6 万澳门元，网站再也打不开。', size=10, after=5)

H2('案例 7　校内群组换钱，8 人逾 11 万澳门元（真实报道）')
P('2024 年 1 到 4 月，一名 19 岁本地大学生在学校聊天群组里说自己能换人民币，8 名同学转钱后一直没收到人民币。'
  '他用“额度已满、账户冻结”拖延，还伪造 3 张转账截图，共骗了逾 11 万澳门元，4 月 22 日被司警截获。',
  size=10, after=5)

H2('AI 换脸（30 秒带过）')
P('视频里的“同学”脸和声音都对，也可能是合成的。涉及借钱，挂断后用自己保存的号码回拨。', size=10, after=5)
label('【分析】', '这三类的共同点：对方先进入你的社交圈——恋爱、游戏群、学校群——再用“内部消息、优惠汇率、'
                   '保证金”让你转账。')
label('【防范】', '不向陌生人投资；游戏交易只走官方平台；换钱走银行或持牌机构；视频借钱先回拨。')
label('【过渡】', '三类讲完。接下来请同学 5，把大家最常问的问题集中回答。')

# ---------------------------------------------------------------- 同学5
H1('七、同学 5：问题汇总 Q&A（4 分钟）')
P('开场：我负责问题汇总。我们收集了同学最常问的 7 个问题，很多不是“怎么防”，而是“被骗之后怎么办”。',
  size=9.8, after=5)

qa = [
    ('Q1　被骗后觉得很丢脸，不想报警，怎么办？',
     '诈骗是犯罪，受害不是你的错。越早报案，银行越可能止付。可以先打 8800 7777 咨询，不一定要马上报案；'
     '但不要因为面子拖过黄金时间。'),
    ('Q2　只损失了几百块，值得报警吗？',
     '值得。金额小也可能和其他案件串起来，报案记录也会让银行和平台重视。打 993，或者直接到司警局。'),
    ('Q3　已经转账了，还能追回吗？',
     '马上联系银行或支付平台要求止付，保留转账记录，然后报案。能不能追回取决于时间和对方账户状态，'
     '没有人能保证；越早越好。'),
    ('Q4　怎么判断对方是不是真警察、真客服？',
     '挂断，用自己查到的官方号码回拨。公检法不会视频办案，不会要求转账或保密；客服不会让你共享屏幕，'
     '也不会要验证码。'),
    ('Q5　学校、书院会不会知道？会不会影响签证或学业？',
     '报案信息依法处理；书院可以提供协助和转介（8822 9400），保安 24 小时（6332 4946）。'
     '受害不会影响学籍，隐瞒反而更麻烦。'),
    ('Q6　朋友让我帮忙收钱、转账，会有事吗？',
     '可能涉及洗钱或帮助信息网络犯罪。不要出借账户，不要代收代转；发现可疑，拒绝并报警。'),
    ('Q7　网上有人说有“内部消息”、稳赚不赔，能信吗？',
     '一律不要信。真有稳赚的机会，对方不需要找你。'),
]
for q, a in qa:
    P(q, size=10, bold=True, before=5, after=1)
    P('答：' + a, size=10, left=0.3, after=2)
P('收尾：如果现场有问题，我们可以继续答；也可以到网站 FAQ 页面，20 个问答都在。下面把时间交回同学 1。',
  size=9.8, before=4, after=3)

# ---------------------------------------------------------------- 提醒
H1('八、讲稿使用提醒')
notes = [
    '数字口径：案例 2 是 16.09 万港元（不是 8.73 万）；案例 7 是 8 人、逾 11 万澳门元（不是 2 人 4.2 万）；'
    '案例 1 是 800 万港元，10 月 23 日报案。',
    '案例 4、5、6 是课程材料改编，讲的时候说“课程材料中的案例”，不要编具体学校、姓名和时间。',
    '案例 3 的受害人是两名中年男子，不是学生；换钱主题用案例 7。',
    '2026 年假客服数据来自司警 2026-10-07 通报：25 人、逾 540 万澳门元，时间 9 月 18 日至 10 月 6 日。',
    '所有数字和出处见网站各详情页“出处”行，或《防诈宣传资料来源与核查报告.docx》。',
]
for i, t in enumerate(notes, 1):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.55); pf.first_line_indent = Cm(-0.4)
    pf.space_after = Pt(2); pf.line_spacing = 1.4
    run(p, '%d. ' % i, size=9.8)
    run(p, t, size=9.8)

# ---------------------------------------------------------------- 对应
H1('九、网站对应速查')
for who, page in [
    ('同学 1', 'index.html（首页）；help.html（结尾求助电话）'),
    ('同学 2', 'types.html → 类型 1、7、9；案例 1'),
    ('同学 3', 'types.html → 类型 2、8、5；案例 4、2'),
    ('同学 4', 'types.html → 类型 3、6、4、10；案例 5、6、7'),
    ('同学 5', 'faq.html（20 个问答）；help.html（应急三步、求助渠道）'),
]:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.35); pf.first_line_indent = Cm(-0.35)
    pf.space_after = Pt(2); pf.line_spacing = 1.4
    run(p, who + '　', size=10, bold=True)
    run(p, page, size=10)

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
cp.title = '小组 Pre 分工与讲稿'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
