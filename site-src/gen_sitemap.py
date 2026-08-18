import json, datetime, os

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.dirname(ROOT)  # sitemap.xml goes one level up (repo root), this script lives in site-src/

pages = json.load(open(os.path.join(ROOT, 'pages.json')))
today = '2026-08-18'

priority_map = {
    'index.html': '1.0',
}
def priority(url):
    if url == 'index.html': return '1.0'
    if url in ('services.html','locations.html','portfolio.html','contact.html'): return '0.9'
    if url.startswith('digital-marketing-') or url in ('seo-services.html','web-development.html','performance-marketing.html','social-media-marketing.html','branding-design.html','ai-automation.html'): return '0.85'
    if url in ('privacy-policy.html','terms-of-service.html'): return '0.3'
    return '0.6'

lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
for slug, meta in pages.items():
    url = meta['url']
    loc_path = '' if url == 'index.html' else url
    lines.append('  <url>')
    lines.append(f'    <loc>https://www.synmediatechnology.com/{loc_path}</loc>')
    lines.append(f'    <lastmod>{today}</lastmod>')
    lines.append('    <changefreq>weekly</changefreq>')
    lines.append(f'    <priority>{priority(url)}</priority>')
    lines.append('  </url>')
lines.append('</urlset>')

open(os.path.join(OUT_DIR, 'sitemap.xml'), 'w').write('\n'.join(lines) + '\n')
print('wrote sitemap.xml with', len(pages), 'urls')
