"""打包可直接打开或部署的静态网站，排除临时工具和原始大视频。"""
from pathlib import Path
import zipfile
import json
from html.parser import HTMLParser

ROOT = Path(__file__).resolve().parents[1]
class References(HTMLParser):
    def handle_starttag(self, tag, attrs):
        for key, value in attrs:
            if key in ('src', 'href') and value and not value.startswith(('#', 'http', 'tel:', 'mailto:', 'data:')):
                assert (ROOT / value).exists(), f'Missing reference: {value}'
for name in ('index.html', 'types.html', 'faq.html', 'help.html'):
    References().feed((ROOT / name).read_text(encoding='utf-8'))
data = json.loads((ROOT / 'materials/content.json').read_text(encoding='utf-8'))
assert len(data['types']) == 10 and len(data['cases']) == 7 and len(data['faq']) == 20
for video in data['videos']:
    assert (ROOT / video['src']).is_file() and (ROOT / video['poster']).is_file()
files = [ROOT / name for name in ('index.html', 'types.html', 'faq.html', 'help.html', 'README.md')]
for folder in ('assets', 'materials', 'scripts'):
    files += [p for p in (ROOT / folder).rglob('*') if p.is_file() and '__pycache__' not in p.parts]
target = ROOT / '澳门书院防诈网站.zip'
with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
    for file in files:
        archive.write(file, file.relative_to(ROOT))
print(f'Validated all pages and media references; packaged {len(files)} files, {target.stat().st_size / 1024 / 1024:.1f} MB')
