#!/usr/bin/env bash
# Inlines website/assets/style.css, banner.js and nav.js into every page.
#
# The stylesheet was render-blocking: the browser had the new document but could
# not paint until the CSS arrived, which showed as a flash on every navigation.
# Inlining removes the round trip, so a page paints as soon as its HTML lands.
#
# assets/style.css and assets/banner.js remain the source of truth - edit those,
# then run this script. The files stay published so URLs in already-cached pages
# keep resolving.
set -euo pipefail
cd "$(dirname "$0")/website"
python3 - <<'PY'
import re, pathlib

css = pathlib.Path('assets/style.css').read_text()
js  = pathlib.Path('assets/banner.js').read_text() + '\n' + pathlib.Path('assets/nav.js').read_text()

pages = ['index.html','zmanim/index.html','member-database/index.html',
         'pledges-receipts/index.html','technologies/index.html',
         'terms/index.html','404.html']

style_block  = f'<!--STYLE-->\n<style>\n{css.strip()}\n</style>\n<!--/STYLE-->'
script_block = f'<!--SCRIPT-->\n<script>\n{js.strip()}\n</script>\n<!--/SCRIPT-->'

for f in pages:
    p = pathlib.Path(f); s = p.read_text()

    if '<!--STYLE-->' in s:
        s = re.sub(r'<!--STYLE-->.*?<!--/STYLE-->', lambda m: style_block, s, flags=re.S)
    else:
        s = re.sub(r'<link rel="stylesheet" href="/assets/style\.css[^"]*">', lambda m: style_block, s)

    if '<!--SCRIPT-->' in s:
        s = re.sub(r'<!--SCRIPT-->.*?<!--/SCRIPT-->', lambda m: script_block, s, flags=re.S)
    else:
        s = re.sub(r'<script src="/assets/banner\.js[^"]*"></script>', lambda m: script_block, s)

    p.write_text(s)
    print(f'  inlined -> {f}')
PY
