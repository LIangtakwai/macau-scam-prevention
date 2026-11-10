"""从 materials/content.js 提取 JSON 对象，同步生成 materials/content.json 备份。

content.js 中的对象必须是严格 JSON 格式（双引号键值、无尾逗号、内部无注释），
日常修改 content.js 后运行本脚本即可保持两个文件一致。
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'materials' / 'content.js'
OUT = ROOT / 'materials' / 'content.json'

text = SRC.read_text(encoding='utf-8')
match = re.search(r'window\.ANTIFRAUD_CONTENT\s*=\s*(\{.*\})\s*;?\s*$', text, re.S)
if not match:
    raise SystemExit('未在 content.js 中找到 ANTIFRAUD_CONTENT 对象')
data = json.loads(match.group(1))  # 解析失败会提示具体行列，便于修正格式
OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print('已同步 content.json：', len(data.get('types', [])), '类诈骗、',
      len(data.get('cases', [])), '个案例、', len(data.get('faq', [])), '个问答、',
      len(data.get('quiz', [])), '道自测题、', len(data.get('glossary', [])), '个名词')
