# -*- coding: utf-8 -*-
"""把 materials 下 5 个文件原文合并成一个紧凑排版的 Word。

文件：案例.docx、诈骗类型.docx、ai-config.js、content.js、content.json
运行：python scripts/build_materials_combined_doc.py
"""
from pathlib import Path
import json

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / '材料文件合集.docx'

BODY_SIZE = 7.2
CODE_SIZE = 6.3
HEAD_SIZE = 8.5
BLACK = RGBColor(0, 0, 0)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.3)
sec.top_margin = Cm(1.0)
sec.bottom_margin = Cm(1.0)

# 关闭文档网格，避免小字号被行网格撑高。
_grid = sec._sectPr.find(qn('w:docGrid'))
if _grid is not None:
    _grid.set(qn('w:type'), 'none')


def set_fonts(run, ascii_font='宋体', east='宋体', size=7.2, bold=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = BLACK
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn('w:ascii'), ascii_font)
    rpr.rFonts.set(qn('w:hAnsi'), ascii_font)
    rpr.rFonts.set(qn('w:eastAsia'), east)


def line(text='', size=BODY_SIZE, bold=False, ascii_font='宋体', east='宋体',
         before=0, after=0, line_spacing=1.0, left=0, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line_spacing
    if left:
        pf.left_indent = Cm(left)
    if keep:
        pf.keep_with_next = True
    # 段落标记也用小字号，否则行高会被默认 10.5pt 的段落标记撑大。
    ppr = p._p.get_or_add_pPr()
    mark = OxmlElement('w:rPr')
    for tag in ('w:sz', 'w:szCs'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), str(max(2, int(round(size * 2)))))
        mark.append(el)
    ppr.append(mark)
    if text:
        r = p.add_run(text)
        set_fonts(r, ascii_font=ascii_font, east=east, size=size, bold=bold)
    return p


def file_head(title, index):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(8)
    pf.space_after = Pt(2)
    pf.keep_with_next = True
    pf.line_spacing = Pt(10)
    r = p.add_run(f'{index}. {title}')
    set_fonts(r, ascii_font='黑体', east='黑体', size=HEAD_SIZE, bold=True)


def docx_lines(path):
    d = Document(path)
    lines = [p.text.rstrip() for p in d.paragraphs]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def raw_lines(path):
    return path.read_text(encoding='utf-8').rstrip('\n').split('\n')


def rule():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '4')
    b.set(qn('w:space'), '1'); b.set(qn('w:color'), '808080')
    bdr.append(b); ppr.append(bdr)


# 标题
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
r = p.add_run('材料文件合集')
set_fonts(r, ascii_font='黑体', east='黑体', size=12, bold=True)
line('案例.docx · 诈骗类型.docx · ai-config.js（文本）· content.js（文本）· content.json（文本）　（2026-11-10）', size=7.2, after=2)
rule()

def load_obj(path):
    """读取 js/json 文件里的对象（js 取第一个 { 到最后一个 }）。"""
    if path.suffix == '.json':
        return json.loads(path.read_text(encoding='utf-8'))
    text = path.read_text(encoding='utf-8')
    return json.loads(text[text.index('{'): text.rindex('}') + 1])


def _scalar(v):
    if isinstance(v, bool):
        return '是' if v else '否'
    if v is None:
        return '（空）'
    return str(v).replace('\n', ' ')


def dump(value, indent, key, out):
    """把 JSON 结构转成「键：值」文本，去掉大括号和引号。"""
    pad = '  ' * indent
    if isinstance(value, dict):
        if key is not None:
            out.append(f'{pad}{key}：')
            indent += 1
            pad = '  ' * indent
        for k, v in value.items():
            dump(v, indent, k, out)
    elif isinstance(value, list):
        if not value:
            out.append(f'{pad}{key}：（空）')
            return
        if all(not isinstance(x, (dict, list)) for x in value):
            joined = '；'.join(_scalar(x) for x in value)
            if len(joined) <= 120:
                out.append(f'{pad}{key}：{joined}')
            else:
                out.append(f'{pad}{key}：')
                for x in value:
                    out.append(f'{pad}  - {_scalar(x)}')
            return
        out.append(f'{pad}{key}：')
        for i, item in enumerate(value, 1):
            # 只在类型/案例/问答等大条目之间空一行；stats、members 等小条目保持紧凑。
            complex_item = isinstance(item, dict) and any(isinstance(v, (dict, list)) for v in item.values())
            if i > 1 and indent == 0 and (key in ('types', 'cases', 'faq', 'glossary', 'posts') or complex_item):
                out.append('')
            if isinstance(item, dict):
                simple = {k: v for k, v in item.items() if not isinstance(v, (dict, list))}
                complex_ = {k: v for k, v in item.items() if isinstance(v, (dict, list))}
                if simple:
                    head = '；'.join(f'{k}：{_scalar(v)}' for k, v in simple.items())
                    out.append(f'{pad}  {i}. {head}')
                else:
                    out.append(f'{pad}  {i}.')
                for k, v in complex_.items():
                    dump(v, indent + 2, k, out)
            elif isinstance(item, list):
                dump(item, indent + 1, f'{i}', out)
            else:
                out.append(f'{pad}  {i}. {_scalar(item)}')
    else:
        out.append(f'{pad}{key}：{_scalar(value)}')


FILES = [
    ('案例.docx', ROOT / 'materials' / '案例.docx', 'docx'),
    ('诈骗类型.docx', ROOT / 'materials' / '诈骗类型.docx', 'docx'),
    ('ai-config.js', ROOT / 'materials' / 'ai-config.js', 'code'),
    ('content.js', ROOT / 'materials' / 'content.js', 'code'),
    ('content.json', ROOT / 'materials' / 'content.json', 'code'),
]

for i, (name, path, kind) in enumerate(FILES, 1):
    file_head(name if kind == 'docx' else name + '（文本）', i)
    if kind == 'docx':
        for t in docx_lines(path):
            line(t, size=BODY_SIZE, line_spacing=Pt(8.0))
    else:
        out = []
        dump(load_obj(path), 0, None, out)
        for t in out:
            # 用不换行空格保留层级缩进，避免 Word 吞掉行首空格。
            indent = len(t) - len(t.lstrip(' '))
            text = '\u00A0' * indent + t[indent:]
            line(text, size=BODY_SIZE, line_spacing=Pt(8.0))

# 页脚页码
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run()
fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
r._r.append(fld); r._r.append(instr); r._r.append(end)
set_fonts(r, size=6.5)

cp = doc.core_properties
cp.title = '材料文件合集'
cp.author = 'CPED2001 第 11 组'
doc.save(OUT)
print('saved:', OUT)
