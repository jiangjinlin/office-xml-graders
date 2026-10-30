# -*- coding: utf-8 -*-
"""抓取 ooxml.org 的元素参考页，存成 markdown 供评分器开发时查证。

注意：ooxml.org 是前端渲染的站点，urllib 只能拿到空壳（正文在 JS 里）。
要么接一个带渲染的抓取（Playwright/IAB），要么直接用 WebFetch 单页取。
本脚本保留下来做批量入口，跑之前先确认单页能抓到正文。

用法: python tools/fetch_ooxml_reference.py
"""
import os
import re
import urllib.request

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reference', 'ooxml')

PAGES = {
    'wordml': ['w-sectpr', 'w-pgmar', 'w-pgsz', 'w-docgrid', 'w-rpr', 'w-ppr',
               'w-shd', 'w-bdr', 'w-tblborders', 'w-numpr', 'w-lvl', 'w-numfmt',
               'w-headerreference', 'w-instrtext', 'w-fldsimple', 'w-drawing',
               'w-hyperlink', 'w-bookmarkstart', 'w-styles', 'w-style'],
    # ooxml.org 的 SpreadsheetML / PresentationML 尚未放出（页面写 Coming soon），
    # Excel 的结构以 ECMA-376 Part 1 §18 和 learn.microsoft.com 为准。
    'spreadsheetml': [],
    'drawingml': ['a-solidfill', 'a-gradfill', 'a-latin', 'a-ea', 'a-rpr', 'a-blip'],
}

HEAD = ('<!-- 抓取自 https://ooxml.org/{path}/ ，仅作评分器开发参考；'
        '规范原文见 ECMA-376 / ISO-IEC 29500 -->\n')


def clean(html):
    html = re.sub(r'(?is)<(script|style|nav|footer|svg)[^>]*>.*?</\1>', ' ', html)
    html = re.sub(r'(?is)<h([1-4])[^>]*>(.*?)</h\1>',
                  lambda m: '\n' + '#' * int(m.group(1)) + ' ' + strip(m.group(2)) + '\n', html)
    html = re.sub(r'(?is)<li[^>]*>(.*?)</li>', lambda m: '- ' + strip(m.group(1)) + '\n', html)
    html = re.sub(r'(?is)</(p|div|tr|table)>', '\n', html)
    html = re.sub(r'(?is)<td[^>]*>', ' | ', html)
    html = re.sub(r'(?is)<br\s*/?>', '\n', html)
    text = strip(html)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def strip(s):
    s = re.sub(r'(?is)<[^>]+>', '', s)
    for a, b in (('&nbsp;', ' '), ('&amp;', '&'), ('&lt;', '<'), ('&gt;', '>'),
                 ('&quot;', '"'), ('&#39;', "'")):
        s = s.replace(a, b)
    return s


def main():
    os.makedirs(OUT, exist_ok=True)
    ok = 0
    for area, names in PAGES.items():
        d = os.path.join(OUT, area)
        os.makedirs(d, exist_ok=True)
        for name in names:
            path = '%s/%s/' % (area, name)
            url = 'https://ooxml.org/' + path
            dest = os.path.join(d, name + '.md')
            if os.path.exists(dest) and os.path.getsize(dest) > 400:
                ok += 1
                continue
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                html = urllib.request.urlopen(req, timeout=30).read().decode('utf-8', 'replace')
                body = clean(html)
                with open(dest, 'w', encoding='utf-8') as f:
                    f.write(HEAD.format(path=path))
                    f.write('# %s\n\n来源: %s\n\n' % (name, url))
                    f.write(body[:20000])
                ok += 1
                print('ok', name)
            except Exception as exc:
                print('fail', name, exc)
    print('done', ok)

    idx = ['# ooxml.org 元素参考（本地副本）', '',
           '来源 https://ooxml.org/ ，规范本体是 ECMA-376 / ISO-IEC 29500。',
           '这些文件是按需抓取的元素页（父/子元素、属性、ECMA 条款号），',
           '写评分器时用来核对结构，不替代规范原文。', '']
    for area in PAGES:
        idx.append('## ' + area)
        for name in PAGES[area]:
            idx.append('- [%s](./%s/%s.md)' % (name, area, name))
        idx.append('')
    with open(os.path.join(OUT, 'README.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(idx))
    print('index written')


if __name__ == '__main__':
    main()
