# -*- coding: utf-8 -*-
"""IndexNow 推送：告诉必应 / Yandex / Seznam 有新页面，比等蜘蛛快。

原理：key 文件公开放在站点根目录，内容就是 key 本身；
推送时带上 host + key + keyLocation，搜索引擎会先校验 key 文件再抓 URL。

用法：
    python tools/indexnow_push.py            # 推 sitemap 全部 URL
    python tools/indexnow_push.py <URL> ...  # 只推指定的
"""
import json
import os
import re
import sys
import urllib.error
import urllib.request

SITE = 'https://web.xiaoshizi.asia'
KEY = 'a7f3c9e14b8d4f62a1c05e7d3b9f8a26'
KEY_FILE = KEY + '.txt'
ENDPOINT = 'https://api.indexnow.org/indexnow'
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def sitemap_urls():
    with urllib.request.urlopen(SITE + '/sitemap.xml', timeout=30) as r:
        return re.findall(r'<loc>([^<]+)</loc>', r.read().decode('utf-8'))


def ensure_key_file():
    d = os.path.join(BASE, 'verify')
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, KEY_FILE)
    if not os.path.exists(p):
        with open(p, 'w', encoding='utf-8') as f:
            f.write(KEY)
    return p


def check_key():
    url = '%s/%s' % (SITE, KEY_FILE)
    with urllib.request.urlopen(url, timeout=20) as r:
        return r.read().decode('utf-8').strip() == KEY


def push(urls):
    payload = json.dumps({'host': SITE.split('//')[1], 'key': KEY,
                          'keyLocation': '%s/%s' % (SITE, KEY_FILE),
                          'urlList': urls}).encode('utf-8')
    req = urllib.request.Request(ENDPOINT, data=payload, method='POST')
    req.add_header('Content-Type', 'application/json; charset=utf-8')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, r.read().decode('utf-8', 'replace')
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode('utf-8', 'replace')


def main():
    ensure_key_file()
    if not check_key():
        print('key 文件还没生效：先 python deploy.py 把 verify/%s 发上去' % KEY_FILE)
        return 1
    urls = [a for a in sys.argv[1:] if a.startswith('http')] or sitemap_urls()
    code, body = push(urls[:10000])
    print('推送 %d 条 → HTTP %s %s' % (len(urls), code, body.strip()[:200]))
    print('200/202 都算成功；403 是 key 校验失败，422 是 URL 不属于该 host')
    return 0


if __name__ == '__main__':
    sys.exit(main())
