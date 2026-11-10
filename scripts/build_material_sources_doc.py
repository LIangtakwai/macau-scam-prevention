# -*- coding: utf-8 -*-
"""生成《材料出处清单》：只列网站材料与出处，不写内容、不做分析。

运行：python scripts/build_material_sources_doc.py
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
OUT = ROOT / '材料出处清单.docx'
BLACK = '000000'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.0)
sec.top_margin = Cm(1.9)
sec.bottom_margin = Cm(1.8)

# 常用来源：key -> (标题, 链接)
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
    'pj20240409': ('司法警察局《司警局反詐程式正式上線 助公眾防範電訊網絡詐騙》2024-04-09',
                   'https://www.gov.mo/zh-hant/news/1056801/'),
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
      first=0, line=1.3, keep=False):
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


def H2(text):
    return P(text, size=9.8, bold=True, font='黑体', before=4, after=1, keep=True)


def src(key, size=8.6):
    title, url = S[key]
    p = P('出处：' + title, size=size, left=0.5, after=0)
    if url:
        q = P('', size=8.2, left=0.9, after=1)
        r = q.add_run(url)
        r.font.size = Pt(8.2)
        r.font.color.rgb = RGBColor(0, 0, 0)
        set_cn(r)
    return p


# 标题
P('材料出处清单', size=17, bold=True, font='黑体', after=2)
P('澳门书院防诈宣传网站（CPED2001 第 11 组 · Scam Prevention Campaign）', size=9, after=1)
P('整理日期：2026-11-10　　链接访问日期：2026-11-05', size=8.6, after=6)
P('说明：本清单只列网站所用材料与出处，不写材料内容。未注明出处的文字为本组原创。', size=9, after=6)

H1('一、案例')
for i, c in enumerate(DATA['cases'], 1):
    H2(f'案例 {i}：{c["title"].split("：", 1)[1]}')
    for k in CASE_SRC[i]:
        src(k)

H1('二、诈骗类型')
for i, t in enumerate(DATA['types'], 1):
    H2(f'类型 {i}：{t["short"]}')
    for k in TYPE_SRC[i]:
        src(k)

H1('三、宣传视频与封面')
H2('视频 1–4（materials/videos/video-1.mp4–video-4.mp4）')
P('出处：澳门官方防诈骗宣传网站。', size=9, left=0.5, after=1)
P('原始 MOV 由课程材料提供，本站用 scripts/prepare_videos.py 转为 H.264/AAC MP4。', size=8.6, left=0.5, after=3)
H2('视频封面、类型封面、案例封面（materials/covers、materials/videos/poster-*.jpg）')
P('出处：同上，均为上述澳门官方防诈骗宣传视频的截帧。', size=9, left=0.5, after=3)

H1('四、B 站延伸视频（27 条）')
for t in DATA['types']:
    P(f'类型「{t["short"]}」：', size=8.8, bold=True, before=2, after=1)
    for b in t['bili']:
        P(f'《{b["title"]}》　{b["author"]}　{b["duration"]}', size=8.8, left=0.4, after=0)
        q = P('', size=8.0, left=0.8, after=2)
        r = q.add_run(b['url']); r.font.size = Pt(8.0); r.font.color.rgb = RGBColor(0, 0, 0); set_cn(r)

H1('五、图片与标志')
H2('澳门大学标志（materials/logos/um-mark.png、um-white.png）')
P('出处：澳门大学校徽，由课程材料提供；版权归澳门大学。', size=9, left=0.5, after=3)
H2('蔡继有书院标志（materials/logos/ckyc.png、ckyc-white.png）')
P('出处：蔡继有书院院徽，由课程材料提供；版权归蔡继有书院。', size=9, left=0.5, after=3)
H2('favicon.svg')
P('出处：本组制作。', size=9, left=0.5, after=3)

H1('六、官方电话与链接')
phones = [
    ('司警局报案热线 993', 'pjhotline'),
    ('司警局防诈骗查询热线 8800 7777', 'pjcrime05'),
    ('澳门大学保安紧急热线 8822 4000；保安及设施服务处 8822 4126', 'umsec'),
    ('蔡继有书院办公室 8822 9400；书院邮箱 ckycollege@um.edu.mo', 'ckyc'),
    ('蔡继有书院保安值班 6332 4946；学术人员当值 6392 9984；紧急求助 999/110/112', 'ckycCare'),
    ('澳大学生事务部 8822 4901、sao@um.edu.mo', 'umBulletin'),
    ('司警局反诈程式（经微信公众号 PJ_Macao 进入）', 'pj20240409'),
]
for name, key in phones:
    H2(name)
    src(key, size=8.4)

H1('七、法规')
for k in ['law2024', 'law2012', 'deep2023']:
    src(k)

H1('八、AI 问答')
H2('DeepSeek API（模型 deepseek-chat）')
P('出处：DeepSeek 开放平台 https://platform.deepseek.com/（访问日期：2026-11-10）。'
  '网站 AI 回答由该服务生成。', size=9, left=0.5, after=3)

H1('九、本组原创（无外部出处）')
for item in [
    '常见问题 35 条、高危信号 14 条、名词解释 15 个、模拟帖子 6 个、自测题 5 道；',
    '案例 4–6 的教学化改写，以及所有案例的“套路重点/防范提示”；',
    '各诈骗类型的“提醒”措辞、页面导语、按钮文案与结构安排；',
    '四个页面的 HTML、CSS、JavaScript 代码与界面设计；',
    '案例 1、2、3、7 中依据报道重写的叙述（事实来自对应报道，文字为本组改写）。',
]:
    p = P('· ' + item, size=9, left=0.45, first=-0.3, after=1)

cp = doc.core_properties
cp.title = '材料出处清单'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
