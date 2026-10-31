# -*- coding: utf-8 -*-
"""把 sitemap 里的 URL 主动推送给百度搜索资源平台。

用法：python tools/baidu_push.py <token>
token 在百度搜索资源平台 → 普通收录 → API 提交 里能看到（站点级，别外传）。
每次推送上限 2000 条，本站在 50 条以内，一次就够。
"""
import re
import sys
import urllib.parse
import urllib.request

SITE = 'https://web.xiaoshizi.asia'
API = 'http://data.zz.baidu.com/urls?site=%s&token=%s'


def sitemap_urls():
    with urllib.request.urlopen(SITE + '/sitemap.xml', timeout=30) as r:
        xml = r.read().decode('utf-8')
    return re.findall(r'<loc>([^<]+)</loc>', xml)


def main():
    if len(sys.argv) < 2:
        print('用法: python tools/baidu_push.py <token>')
        return 1
    token = sys.argv[1].strip()
    urls = sitemap_urls()
    body = '\n'.join(urls).encode('utf-8')
    req = urllib.request.Request(API % (SITE, token), data=body, method='POST')
    req.add_header('Content-Type', 'text/plain')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            print(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        print('推送失败', e.code, e.read().decode('utf-8', 'replace'))
        return 1
    print('共 %d 条 URL' % len(urls))
    return 0


if __name__ == '__main__':
    sys.exit(main())
