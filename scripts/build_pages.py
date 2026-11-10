"""生成五个独立 HTML 页面；共享页头、导航与资源引用。"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
# (文件名, data-page, 导航名称, 标题, 额外脚本)
pages = [
    ('index.html', 'home', '首页', '首页', ['assets/js/main.js']),
    ('types.html', 'types', '认识诈骗类型', '认识诈骗类型', ['assets/js/main.js']),
    ('faq.html', 'faq', '常见问题解答', '常见问题解答', ['assets/js/main.js', 'materials/ai-config.js']),
    ('help.html', 'help', '求助', '求助', ['assets/js/main.js']),
]
for filename, page, nav_label, title, extra_scripts in pages:
    nav = '\n'.join(
        f'<a href="{name}"' + (' aria-current="page"' if name == filename else '') + f'>{label}</a>'
        for name, _, label, _, _ in pages
    )
    scripts = '\n  '.join(f'<script src="{s}" defer></script>' for s in ['materials/content.js', *extra_scripts])
    html = f'''<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="蔡继有书院学生防诈宣传（Scam Prevention Campaign）：宣传视频、诈骗类型与案例、常见问题、AI 智能问答和澳门求助渠道。">
  <title>{title} | 蔡继有书院防诈宣传</title>
  <link rel="icon" type="image/svg+xml" href="materials/favicon.svg">
  <link rel="stylesheet" href="assets/css/style.css">
  {scripts}
</head>
<body data-page="{page}">
  <a class="skip" href="#main">跳至主要内容</a>
  <!-- 五页共用固定顶部导航；两枚标志均使用本地材料。 -->
  <header class="site-header">
    <div class="header-inner">
      <a class="brand" href="index.html">书院防诈宣传<small>CHOI KAI YAU COLLEGE</small></a>
      <div class="logos"><img src="materials/logos/ckyc.png" alt="蔡继有书院标志" width="49" height="55"><img src="materials/logos/um-mark.png" alt="澳门大学标志" width="143" height="55"></div>
    </div>
    <div class="nav-wrap"><nav aria-label="主要导航">{nav}</nav></div>
  </header>
  <main id="main"><noscript>请启用 JavaScript 以查看宣传材料和问答。您也可直接下载 <a href="materials/诈骗类型.docx">诈骗类型材料</a>、<a href="materials/案例.docx">案例材料</a>。报案热线：993；防诈骗查询：8800 7777。</noscript></main>
  <footer><div class="footer-inner"><span>蔡继有书院 · 澳门大学</span><span>提高防骗意识，守护校园生活</span><span>Scam Prevention Campaign · 第 11 组</span></div></footer>
</body>
</html>
'''
    (ROOT / filename).write_text(html, encoding='utf-8')
print('Generated four HTML pages')
