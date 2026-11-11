"""Embeds App.html into Code.gs as APP_HTML (copies self-update from Code.gs alone)."""
import json, re, pathlib
d = pathlib.Path(__file__).parent
code = (d / 'Code.gs').read_text()
html = (d / 'App.html').read_text()
marker = '// ---------- app page (generated from App.html by build.py — edit App.html, not this) ----------'
code = code.split(marker)[0].rstrip() + '\n\n' + marker + '\nconst APP_HTML = ' + json.dumps(html, ensure_ascii=False) + ';\n'
(d / 'Code.gs').write_text(code)
print('embedded', len(html), 'chars')
