# -*- coding: utf-8 -*-
"""生成《外界信息来源清单》：网站用到的全部非原创信息与出处。

只列来源，不做排版、不做分析。数据：materials/content.json + 已核对的公开链接。
运行：python scripts/build_external_sources_doc.py
"""
import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'materials' / 'content.json').read_text(encoding='utf-8'))
OUT = ROOT / '外界信息来源清单.docx'
BLACK = '000000'
ACCESS = '2026-11-05'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.3)
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


def P(text='', size=10, bold=False, font='宋体', before=0, after=2, left=0,
      first=0, line=1.35, align=None, keep=False):
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
    return P(text, size=13, bold=True, font='黑体', before=10, after=4, keep=True)


def H2(text):
    return P(text, size=10.8, bold=True, font='黑体', before=8, after=3, keep=True)


def entry(no, text, size=9.6):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.95); pf.first_line_indent = Cm(-0.95)
    pf.space_before = Pt(0); pf.space_after = Pt(2); pf.line_spacing = 1.32
    run(p, f'[{no}] ', size=size, bold=True)
    run(p, text, size=size)
    return p


def bullet(text, size=9.6, left=0.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(left + 0.35); pf.first_line_indent = Cm(-0.35)
    pf.space_before = Pt(0); pf.space_after = Pt(2); pf.line_spacing = 1.35
    run(p, '· ', size=size)
    run(p, text, size=size)
    return p


# ---------------------------------------------------------------- 标题与说明
P('外界信息来源清单', size=19, bold=True, font='黑体', after=2)
P('用于学术诚信与来源核对', size=9.5, after=1)
P('澳门书院防诈宣传网站（CPED2001 第 11 组 · Scam Prevention Campaign）', size=9.5, after=1)
P('整理日期：2026-11-10　　检索/访问日期：2026-11-05（链接于 2026-11-10 复核）', size=9, after=8)
P('说明：', size=9.8, bold=True, after=2)
P('1. 本清单列出网站中全部非本组原创的信息及其出处，按“使用位置索引 → 关键数据 → 完整来源列表”排列；'
  '每条注明用在哪里，便于逐项核对。', size=9.6, after=2)
P('2. 未列入本清单的文字，包括常见问题（35 条）、高危信号（14 条）、名词解释（15 个）、模拟帖子（6 个）、'
  '自测题（5 道）、案例 4–6 的教学化改写、各案例的“套路重点/防范提示”、页面文案与代码设计，均为本组原创。', size=9.6, after=2)
P('3. 引用的事实、数字、法条、电话、图片、视频与外部服务均在此列明；本清单不复制长段原文，只给出可核查的出处。', size=9.6, after=8)

# ---------------------------------------------------------------- 一、索引
H1('一、按网站位置索引')

H2('1. 案例')
for i, c in enumerate(DATA['cases'], 1):
    refs = {
        1: '[1][2][26]', 2: '[3][4][26]', 3: '[5][26]', 4: '[6][26]',
        5: '[7][26]', 6: '[8][26]', 7: '[9][26]',
    }[i]
    bullet(f'案例 {i}：{c["title"].split("：", 1)[1]} → {refs}', size=9.4)

H2('2. 诈骗类型')
type_refs = {
    1: '[11][1][2][27]', 2: '[10][6][27]', 3: '[11][7][27]', 4: '[12][13][14][5][27]',
    5: '[15][3][27]', 6: '[8][27]', 7: '[18][8][27]', 8: '[16][27]', 9: '[16][8][27]',
    10: '[19][27]',
}
for i, t in enumerate(DATA['types'], 1):
    bullet(f'类型 {i}（{t["short"]}） → {type_refs[i]}', size=9.4)

H2('3. 求助页')
P('求助电话、邮箱与官方链接 → [20][21][22][23][24][25]；反诈程式介绍 → [17]；'
  '电话核对日期与“以官方页面为准”的说明 → [20]–[25]。', size=9.4, after=3)

H2('4. 素材与其他')
bullet('4 段宣传视频、封面与海报（视频截帧） → [29]', size=9.4)
bullet('澳门大学标志 → [30]；蔡继有书院标志 → [31]', size=9.4)
bullet('AI 问答（接口与模型） → [32]', size=9.4)
bullet('B 站延伸视频 27 条 → [33]–[59]（逐条见第三部分）', size=9.4)
bullet('案例与类型的原始文字底稿 → [26][27]；作业要求与页面结构 → [28]', size=9.4)

# ---------------------------------------------------------------- 二、关键数据
H1('二、关键数据与事实的来源')
facts = [
    ('800 万港元、报案日期 2025-10-23', '案例 1', '[1][2]'),
    ('三名学生 160,900 港元；另 2 人 54,400 港元；约 40 个单位、约 210 万港元', '案例 2', '[3][4]'),
    ('543,464 元人民币、200 张千元港币、两案共 400 张练功券', '案例 3', '[5]'),
    ('相近案例：118,000 + 7,000 澳门元、佣金 8%', '案例 4', '[6]'),
    ('相近案例：逾 28 万港元、10 万“按金”', '案例 5', '[7]'),
    ('相近案例：三宗游戏账号/皮肤骗案、合共逾 6 万澳门元', '案例 6', '[8]'),
    ('8 名同校同学、逾 11 万澳门元、伪造 3 张截图、2024-04-22 拘捕', '案例 7', '[9]'),
    ('148,532 澳门元、首单 1,300 元返 1,429 元', '类型 2', '[10]'),
    ('逾 257 万港元（2020 假冒公检法案）“切勿尽信来电显示”等 9 条对策', '类型 1', '[11]'),
    ('25 人、逾 540 万澳门元、2026-09-18 至 10-06', '类型 7', '[18]'),
    ('第 20/2024 号法律 2024-10-29 生效；最高 5 年；现金、筹码可扣押', '类型 4', '[12][14]'),
    ('第 16/2012 号法律 2013-07-01 生效；准照、书面合同', '类型 5', '[15]'),
    ('《深度合成管理规定》2023-01-10 施行；第 14、17、18 条', '类型 10', '[19]'),
    ('993 报案热线；8800 7777 查询热线（2017-08-17 设立）', '求助页', '[20][21]'),
    ('8822 4000 保安紧急热线；8822 4126 保安及设施服务处', '求助页', '[22]'),
    ('8822 9400 书院办公室；ckycollege@um.edu.mo', '求助页', '[23]'),
    ('6332 4946 书院保安值班；6392 9984 学术人员当值；999/110/112', '求助页', '[24]'),
    ('8822 4901、sao@um.edu.mo（学生事务部）', '求助页', '[25]'),
]
for text, where, refs in facts:
    bullet(f'{text}　—　{where}　{refs}', size=9.3)

# ---------------------------------------------------------------- 三、来源列表
H1('三、完整来源列表')

media = [
    '[1] 澳門日報. 港女生墮「公檢法」失八百萬[N]. 澳門日報, 2025-10-26(A02). '
    'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html（访问日期：%s）' % ACCESS,
    '[2] 曾毅, 陳慶輝. 又是「內地公安」騙局，澳門大學生被騙八百萬港元[N/OL]. 廣州日報大洋網, 2025-10-27. '
    'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm（访问日期：%s）' % ACCESS,
    '[3] 澳門日報. 中介呃租客預繳租金二百萬[N]. 澳門日報, 2023-09-02(A09). '
    'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html（访问日期：%s）' % ACCESS,
    '[4] 澳門力報. 本澳一地產經紀呃租金去炒股 40租戶被騙逾210萬港元[N/OL]. 2023-09-01. '
    'https://exmoo.com/article/225209.html（访问日期：%s）' % ACCESS,
    '[5] 澳門日報. 三男女練功券詐騙被捕[N]. 澳門日報, 2024-01-12(A08). '
    'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html（访问日期：%s）' % ACCESS,
    '[6] 澳門日報. 兩男女兼職刷單被騙逾十二萬[N]. 澳門日報, 2023-07-07(A05). '
    'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html（访问日期：%s）' % ACCESS,
    '[7] 澳門日報. 中年女網投被騙廿八萬[N]. 澳門日報, 2023-10-24(A06). '
    'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html（访问日期：%s）' % ACCESS,
    '[8] 澳門日報. 七人墮電信詐騙失逾六十萬[N]. 澳門日報, 2023-03-17(A05). '
    'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html（访问日期：%s）' % ACCESS,
    '[9] 澳門日報. 大學生換錢呃同學11萬[N]. 澳門日報, 2024-04-24(A01). '
    'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html（访问日期：%s）' % ACCESS,
    '[10] 澳門日報. 男大生墮刷單騙局失逾十四萬[N]. 澳門日報, 2022-09-03(A04). '
    'https://appimg.modaily.cn/app/szb/pc/content/202209/03/content_212804.html（访问日期：%s）' % ACCESS,
    '[11] 司法警察局. 詐騙手法層出不窮 司警局呼籲市民小心防範[EB/OL]. 2020-11-28. '
    'https://www.gov.mo/zh-hant/news/354373/（访问日期：%s）' % ACCESS,
    '[12] 司法警察局. 司警局與博監局及博企舉行會議 共商配合《打擊不法賭博犯罪法》執法工作[EB/OL]. 2024-11-05. '
    'https://www.gov.mo/zh-hant/news/1102576/（访问日期：%s）' % ACCESS,
    '[13] 司法警察局. 司警局加強宣傳《打擊不法賭博犯罪法》 促警民合力打擊不法匯兌活動[EB/OL]. 2024-11-18. '
    'https://www.gov.mo/zh-hant/news/1106939/（访问日期：%s）' % ACCESS,
    '[14] 澳門特別行政區. 第20/2024號法律《打擊不法賭博犯罪法》[Z]. 2024-10-29 生效. '
    '條文要點見 [12]；本清單未單列法規連結。',
    '[15] 澳門特別行政區政府. 房地產中介業務（第16/2012號法律《房地產中介業務法》）[EB/OL]. '
    'https://www.gov.mo/zh-hant/services/ps-1911/（访问日期：%s）' % ACCESS,
    '[16] 司法警察局. 警情通告：提防假冒本地銀行貸款廣告 警惕釣魚網站以免誤墮騙局[EB/OL]. 2024-03-25. '
    'https://www.gov.mo/zh-hans/news/748837/（访问日期：%s）' % ACCESS,
    '[17] 司法警察局. 司警局反詐程式正式上線 助公眾防範電訊網絡詐騙[EB/OL]. 2024-04-09. '
    'https://www.gov.mo/zh-hant/news/1056801/（访问日期：%s）' % ACCESS,
    '[18] 澳門力報. 司警：提防抖音假客服詐騙 三周接獲25人報案失逾540萬[N/OL]. 2026-10-07. '
    'https://www.exmoo.com/article/266439.html（访问日期：%s）' % ACCESS,
    '[19] 國家互聯網信息辦公室, 工業和信息化部, 公安部. 互聯網信息服務深度合成管理規定[Z]. '
    '2022-11-25 公布, 2023-01-10 施行. https://www.gov.cn/zhengce/zhengceku/2022-12/12/content_5731431.htm'
    '（访问日期：%s）' % ACCESS,
    '[20] 司法警察局. 熱線電話[EB/OL]. https://www.pj.gov.mo/Web/Policia/hotline/（访问日期：%s）' % ACCESS,
    '[21] 司法警察局. 防詐騙查詢熱線[EB/OL]. '
    'https://www.pj.gov.mo/Web/Policia/crime05.html?lang=en（访问日期：%s）' % ACCESS,
    '[22] 澳門大學. 保安服務[EB/OL]. '
    'https://sfs.cmdo.um.edu.mo/our-services/security-services/（访问日期：%s）' % ACCESS,
    '[23] 蔡繼有書院. 聯繫方式[EB/OL]. https://ckyc.rc.um.edu.mo/（访问日期：%s）' % ACCESS,
    '[24] 蔡繼有書院. 學生關懷服務[EB/OL]. '
    'https://ckyc.rc.um.edu.mo/student-life/student-caring-services/（访问日期：%s）' % ACCESS,
    '[25] 澳門大學學生事務部. 司法警察局－學校安全聯網通訊[EB/OL]. 2026-01-27. '
    'https://e-bulletin.um.edu.mo/notice/486001/（访问日期：%s）' % ACCESS,
    '[26] 課程材料. 案例.docx[Z]. 站內文件 materials/案例.docx；案例 1–7 的原始文字底稿，'
    '案例 1、2、3、7 另經 [1]–[5]、[9] 核對。',
    '[27] 課程材料. 詐騙類型.docx[Z]. 站內文件 materials/詐騙類型.docx；10 類詐騙的原始文字底稿。',
    '[28] 課程材料. 提示詞.docx[Z]. 課程作業要求與頁面結構；文件未隨站點發布。',
    '[29] 課程材料. 4 段宣傳視頻（原始 MOV，經 scripts/prepare_videos.py 轉為 H.264/AAC MP4）[Z]. '
    '站內文件 materials/videos/video-1.mp4–video-4.mp4；封面與海報為視頻截幀。',
    '[30] 澳門大學. 校徽[Z]. 由課程材料提供（materials/logos/um-mark.png、um-white.png）；'
    '版權歸澳門大學，使用須符合其視覺識別規範。',
    '[31] 蔡繼有書院. 院徽[Z]. 由課程材料提供（materials/logos/ckyc.png、ckyc-white.png）；'
    '版權歸蔡繼有書院，使用須符合其規範。',
    '[32] DeepSeek. DeepSeek API（模型 deepseek-chat）[EB/OL]. https://platform.deepseek.com/'
    '（访问日期：2026-11-10）；網站 AI 問答由該服務生成，非本組原創。',
]

for line in media:
    entry(line.split(']')[0].lstrip('['), line.split('] ', 1)[1] if '] ' in line else line)

# B 站视频
H2('B 站延伸视频（27 条）')
n = 32
for t in DATA['types']:
    for b in t['bili']:
        n += 1
        entry(n, f'{b["author"]}. {b["title"]}[EB/OL]. Bilibili, {b["duration"]}. {b["url"]}'
                  f'（项目组核对日期：2026-10-30）. 用于：类型“{t["short"]}”延伸视频。', size=9.3)

# ---------------------------------------------------------------- 四、原创声明
H1('四、原创部分声明')
P('以下内容为本组原创，未使用外部文字；其中涉及的事实与数字已在上文标注来源：', size=9.6, after=2)
for item in [
    '网站四个页面的 HTML、CSS、JavaScript 代码与界面设计；',
    '常见问题 35 条、高危信号 14 条、名词解释 15 个、模拟帖子 6 个、自测题 5 道；',
    '案例 4–6 的教学化改写，以及所有案例的“套路重点/防范提示”；',
    '各诈骗类型的“提醒”措辞、页面导语、按钮文案与结构安排；',
    '案例 1、2、3、7 中依据报道重写的叙述（事实来自对应报道，文字为本组改写）。',
]:
    bullet(item, size=9.4)

H1('五、技术与工具')
P('网站为原生 HTML/CSS/JavaScript，无第三方框架、无 CDN、无外部字体、无统计脚本；'
  '字体使用系统字体，视频转码与图片处理使用课程环境中的本地工具。AI 问答调用 DeepSeek API [32]。', size=9.5, after=4)

cp = doc.core_properties
cp.title = '外界信息来源清单'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
