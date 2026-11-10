# -*- coding: utf-8 -*-
"""生成《材料出处清单》（去重版）：同一出处只列一次，标注对应材料。

运行：python scripts/build_material_sources_doc.py
"""
import json
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'materials' / 'content.json').read_text(encoding='utf-8'))
OUT = ROOT / '材料出处清单.docx'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.0)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.7)

S = {
    'md20251026': ('《澳门日报》2025-10-26 A02《港女生墮「公檢法」失八百萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202510/26/content_442031.html'),
    'gz20251027': ('《广州日报》2025-10-27《又是「内地公安」骗局，澳门大学生被骗八百万港元》',
                   'https://news.dayoo.com/gzrbrmt/202510/27/170628_54888366.htm'),
    'md20230902': ('《澳门日报》2023-09-02 A09《中介呃租客預繳租金二百萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202309/02/content_286009.html'),
    'ex20230901': ('力报 2023-09-01《本澳一地產經紀呃租金去炒股 40租戶被騙逾210萬港元》',
                   'https://exmoo.com/article/225209.html'),
    'md20240112': ('《澳门日报》2024-01-12 A08《三男女練功券詐騙被捕》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202401/12/content_313501.html'),
    'md20230707': ('《澳门日报》2023-07-07 A05《兩男女兼職刷單被騙逾十二萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202307/07/content_274031.html'),
    'md20231024': ('《澳门日报》2023-10-24 A06《中年女網投被騙廿八萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202310/24/content_296925.html'),
    'md20230317': ('《澳门日报》2023-03-17 A05《七人墮電信詐騙失逾六十萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202303/17/content_250629.html'),
    'md20240424': ('《澳门日报》2024-04-24 A01《大學生換錢呃同學11萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202404/24/content_333259.html'),
    'md20220903': ('《澳门日报》2022-09-03 A04《男大生墮刷單騙局失逾十四萬》',
                   'https://appimg.modaily.cn/app/szb/pc/content/202209/03/content_212804.html'),
    'pj20201128': ('司法警察局《詐騙手法層出不窮 司警局呼籲市民小心防範》2020-11-28',
                   'https://www.gov.mo/zh-hant/news/354373/'),
    'pj20241105': ('司法警察局《司警局與博監局及博企舉行會議 共商配合〈打擊不法賭博犯罪法〉執法工作》2024-11-05',
                   'https://www.gov.mo/zh-hant/news/1102576/'),
    'pj20241118': ('司法警察局《司警局加強宣傳〈打擊不法賭博犯罪法〉 促警民合力打擊不法匯兌活動》2024-11-18',
                   'https://www.gov.mo/zh-hant/news/1106939/'),
    'law2024': ('第 20/2024 號法律《打擊不法賭博犯罪法》，2024-10-29 生效（要點見司警局 2024-11-05 通稿）', ''),
    'law2012': ('第 16/2012 號法律《房地產中介業務法》（澳門特別行政區政府「房地產中介業務」頁）',
                'https://www.gov.mo/zh-hant/services/ps-1911/'),
    'pj20240325': ('司法警察局《警情通告：提防假冒本地銀行貸款廣告 警惕釣魚網站以免誤墮騙局》2024-03-25',
                   'https://www.gov.mo/zh-hans/news/748837/'),
    'ex20261007': ('力报 2026-10-07《司警：提防抖音假客服詐騙 三周接獲25人報案失逾540萬》',
                   'https://www.exmoo.com/article/266439.html'),
    'deep2023': ('《互聯網信息服務深度合成管理規定》，2023-01-10 施行',
                 'https://www.gov.cn/zhengce/zhengceku/2022-12/12/content_5731431.htm'),
    'pjhotline': ('司法警察局「熱線電話」', 'https://www.pj.gov.mo/Web/Policia/hotline/'),
    'pjcrime05': ('司法警察局「防詐騙查詢熱線」',
                  'https://www.pj.gov.mo/Web/Policia/crime05.html?lang=en'),
    'umsec': ('澳門大學「保安服務」', 'https://sfs.cmdo.um.edu.mo/our-services/security-services/'),
    'ckyc': ('蔡繼有書院「聯繫方式」', 'https://ckyc.rc.um.edu.mo/'),
    'ckycCare': ('蔡繼有書院「學生關懷服務」',
                 'https://ckyc.rc.um.edu.mo/student-life/student-caring-services/'),
    'umBulletin': ('澳門大學學生事務部《司法警察局－學校安全聯網通訊》2026-01-27',
                   'https://e-bulletin.um.edu.mo/notice/486001/'),
    'pj20240409': ('司法警察局《司警局反詐程式正式上線 助公眾防範電訊網絡詐騙》2024-04-09',
                   'https://www.gov.mo/zh-hant/news/1056801/'),
    'caseDoc': ('課程材料《案例.docx》（站內文件 materials/案例.docx）', ''),
    'typeDoc': ('課程材料《詐騙類型.docx》（站內文件 materials/詐騙類型.docx）', ''),
}

CASE_SRC = {
    1: ['md20251026', 'gz20251027', 'caseDoc'],
    2: ['md20230902', 'ex20230901', 'caseDoc'],
    3: ['md20240112', 'caseDoc'],
    4: ['md20230707', 'caseDoc'],
    5: ['md20231024', 'caseDoc'],
    6: ['md20230317', 'caseDoc'],
    7: ['md20240424', 'caseDoc'],
}

TYPE_SRC = {
    1: ['pj20201128', 'md20251026', 'gz20251027', 'typeDoc'],
    2: ['md20220903', 'md20230707', 'typeDoc'],
    3: ['pj20201128', 'md20231024', 'typeDoc'],
    4: ['pj20241105', 'pj20241118', 'law2024', 'md20240112', 'typeDoc'],
    5: ['law2012', 'md20230902', 'typeDoc'],
    6: ['md20230317', 'typeDoc'],
    7: ['ex20261007', 'md20230317', 'typeDoc'],
    8: ['pj20240325', 'typeDoc'],
    9: ['pj20240325', 'md20230317', 'typeDoc'],
    10: ['deep2023', 'typeDoc'],
}

PHONE_KEYS = ['pjhotline', 'pjcrime05', 'umsec', 'ckyc', 'ckycCare', 'umBulletin', 'pj20240409']
ORDER = [
    'md20251026', 'gz20251027', 'md20230902', 'ex20230901', 'md20240112', 'md20230707',
    'md20231024', 'md20230317', 'md20240424', 'md20220903', 'pj20201128', 'pj20241105',
    'pj20241118', 'law2024', 'law2012', 'pj20240325', 'ex20261007', 'deep2023',
    'pjhotline', 'pjcrime05', 'umsec', 'ckyc', 'ckycCare', 'umBulletin', 'pj20240409',
    'caseDoc', 'typeDoc',
]

uses = {k: [] for k in ORDER}
for i, keys in CASE_SRC.items():
    for k in keys:
        label = f'案例 {i}'
        if label not in uses[k]:
            uses[k].append(label)
for i, keys in TYPE_SRC.items():
    for k in keys:
        label = f'类型 {i}'
        if label not in uses[k]:
            uses[k].append(label)
for k in PHONE_KEYS:
    uses[k].append('求助页')


def set_cn(run, font='宋体'):
    run.font.name = font
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)


def run(p, text, size=9.2, bold=False, font='宋体'):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.font.color.rgb = RGBColor(0, 0, 0)
    set_cn(r, font)
    return r


def P(text='', size=9.2, bold=False, font='宋体', before=0, after=2, left=0,
      first=0, line=1.28, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if left:
        pf.left_indent = Cm(left)
    if first:
        pf.first_line_indent = Cm(first)
    if keep:
        pf.keep_with_next = True
    if text:
        run(p, text, size=size, bold=bold, font=font)
    return p


def H1(text):
    return P(text, size=12.5, bold=True, font='黑体', before=8, after=2, keep=True)


def url_line(url, size=8.2):
    q = P('', size=size, left=1.0, after=2)
    r = q.add_run(url)
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0, 0, 0)
    set_cn(r)


P('材料出处清单', size=17, bold=True, font='黑体', after=2)
P('澳门书院防诈宣传网站（CPED2001 第 11 组 · Scam Prevention Campaign）', size=9, after=1)
P('整理日期：2026-11-10　　链接访问日期：2026-11-05', size=8.6, after=5)
P('说明：同一出处只列一次，后面标注对应的材料；未注明出处的内容为本组原创。', size=9, after=6)

H1('一、文字与官方来源')
n = 0
for key in ORDER:
    if not uses[key]:
        continue
    n += 1
    title, url = S[key]
    p = P('', size=9.0, left=0.9, first=-0.9, after=0)
    run(p, f'[{n}] ', size=9.0, bold=True)
    run(p, title + f'（对应：{"、".join(uses[key])}）', size=9.0)
    if url:
        url_line(url)

H1('二、B 站视频')
for t in DATA['types']:
    for b in t['bili']:
        n += 1
        p = P('', size=8.9, left=0.9, first=-0.9, after=0, before=2)
        run(p, f'[{n}] ', size=8.9, bold=True)
        run(p, f'《{b["title"]}》　{b["author"]}　{b["duration"]}（对应：类型「{t["short"]}」）', size=8.9)
        url_line(b['url'], size=7.9)

H1('三、宣传视频、封面与标志')
n += 1
p = P('', size=9.0, left=0.9, first=-0.9, after=0)
run(p, f'[{n}] ', size=9.0, bold=True)
run(p, '澳门官方防诈骗宣传网站（对应：视频 1–4，以及视频封面、类型封面、案例封面）', size=9.0)
for label, text in [
    ('澳门大学标志', '澳门大学校徽，由课程材料提供（materials/logos/um-mark.png、um-white.png）'),
    ('蔡继有书院标志', '蔡继有书院院徽，由课程材料提供（materials/logos/ckyc.png、ckyc-white.png）'),
]:
    n += 1
    p = P('', size=9.0, left=0.9, first=-0.9, after=0, before=2)
    run(p, f'[{n}] ', size=9.0, bold=True)
    run(p, label + '：' + text, size=9.0)

H1('四、AI 问答')
n += 1
p = P('', size=9.0, left=0.9, first=-0.9, after=0)
run(p, f'[{n}] ', size=9.0, bold=True)
run(p, 'DeepSeek API（模型 deepseek-chat），https://platform.deepseek.com/（访问日期：2026-11-10）', size=9.0)

H1('五、本组原创（无外部出处）')
P('常见问题 35 条、高危信号 14 条、名词解释 15 个、模拟帖子 6 个、自测题 5 道；'
  '案例 4–6 的教学化改写、各案例的“套路重点/防范提示”、各类型的“提醒”措辞；'
  '页面文案、四个页面的 HTML/CSS/JavaScript 代码与界面设计。', size=9.0, after=3)

cp = doc.core_properties
cp.title = '材料出处清单'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
