"""Embeds App.html into Core.gs as APP_HTML (copies self-update from Core.gs alone)."""
import json, re, pathlib
d = pathlib.Path(__file__).parent
code = (d / 'Core.gs').read_text()
html = (d / 'App.html').read_text()
marker = '// ---------- app page (generated from App.html by build.py — edit App.html, not this) ----------'
code = code.split(marker)[0].rstrip() + '\n\n' + marker + '\nconst APP_HTML = ' + json.dumps(html, ensure_ascii=False) + ';\n'
(d / 'Core.gs').write_text(code)
print('embedded', len(html), 'chars')

# The loader must wrap exactly the functions Core.gs exposes.
core_entries = set(re.findall(r'\w+', re.search(r'const ENTRY_ = \{([^}]*)\}', code, re.S).group(1)))
loader = (d / 'Loader.gs').read_text()
loader_entries = set(re.findall(r"return ia_\('(\w+)'", loader))
assert core_entries == loader_entries, ('mismatch', core_entries ^ loader_entries)
print('loader wraps all', len(core_entries), 'entry points')
