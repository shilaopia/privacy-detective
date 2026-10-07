# -*- coding: utf-8 -*-
"""修复审查分片JSON中字符串值内的裸双引号（未转义）。"""
import re, json, sys

for path in sys.argv[1:]:
    lines = open(path, encoding='utf-8').read().splitlines()
    out = []
    for ln in lines:
        m = re.match(r'^(\s*"[^"]*"\s*:\s*")(.*)("\s*,?\s*)$', ln)
        if m and '"' in m.group(2):
            fixed = m.group(2).replace('\\', '\\\\').replace('"', '\\"')
            ln = m.group(1) + fixed + m.group(3)
        out.append(ln)
    txt = '\n'.join(out)
    json.loads(txt)  # 校验
    open(path, 'w', encoding='utf-8', newline='\n').write(txt)
    print('fixed OK:', path)
