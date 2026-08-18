#!/usr/bin/env python3
import os, json

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.dirname(ROOT)  # output files go one level up (repo root), this script lives in site-src/
FRAG = os.path.join(ROOT, 'frag')
CONTENT = os.path.join(ROOT, 'content')

head_top = open(os.path.join(FRAG, 'head-top.html')).read()
head_bottom = open(os.path.join(FRAG, 'head-bottom.html')).read()
body_top = open(os.path.join(FRAG, 'body-top.html')).read()
body_bottom = open(os.path.join(FRAG, 'body-bottom.html')).read()

pages = json.load(open(os.path.join(ROOT, 'pages.json')))

for slug, meta in pages.items():
    content_path = os.path.join(CONTENT, slug + '.html')
    if not os.path.exists(content_path):
        print('MISSING CONTENT:', slug)
        continue
    content = open(content_path).read()

    canon_path = '' if meta['url'] == 'index.html' else meta['url']
    page_head = f'''<title>{meta['title']}</title>
<meta name="description" content="{meta['description']}">
<meta name="keywords" content="{meta['keywords']}">
<link rel="canonical" href="https://www.synmediatechnology.com/{canon_path}">
<meta property="og:url" content="https://www.synmediatechnology.com/{canon_path}">
<meta property="og:title" content="{meta['title']}">
<meta property="og:description" content="{meta['og_description']}">
<meta name="twitter:title" content="{meta['title']}">
<meta name="twitter:description" content="{meta['og_description']}">
'''
    extra_ld = meta.get('extra_ld', '')
    if extra_ld:
        page_head += extra_ld + '\n'

    nav_page = meta.get('nav_page', slug)
    body = body_top.replace('__PAGE__', nav_page)

    out = head_top + page_head + head_bottom + body + content + body_bottom
    out_path = os.path.join(OUT_DIR, meta['url'])
    with open(out_path, 'w') as f:
        f.write(out)
    print('built', meta['url'], len(out), 'bytes')
