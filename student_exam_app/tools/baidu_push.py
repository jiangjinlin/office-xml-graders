# -*- coding: utf-8 -*-
"""把 sitemap 里的 URL 推送给百度普通收录（自动记录已推、按剩余额度停）。

用法：
    python tools/baidu_push.py            # 推到当天额度用完为止
    python tools/baidu_push.py --dry-run  # 只看还剩多少要推
    python tools/baidu_push.py <token>    # 临时指定 token

token 读取顺序：命令行参数 > 环境变量 BAIDU_TOKEN > ~/.workbuddy/secrets/baidu-push.json
已推记录：tools/baidu_pushed.json（下次跳过，避免重复消耗额度）
"""
import json
import os
import re
import sys
import urllib.error
import urllib.request

SITE = 'https://web.xiaoshizi.asia'
API = 'http://data.zz.baidu.com/urls?site=%s&token=%s'
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG = os.path.join(BASE, 'tools', 'baidu_pushed.json')
SECRET = os.path.expanduser('~/.workbuddy/secrets/baidu-push.json')


def token():
    for arg in sys.argv[1:]:
        if not arg.startswith('-'):
            return arg.strip()
    if os.environ.get('BAIDU_TOKEN'):
        return os.environ['BAIDU_TOKEN'].strip()
    if os.path.exists(SECRET):
        with open(SECRET, encoding='utf-8') as f:
            return json.load(f).get('token', '').strip()
    return ''


def sitemap_urls():
    with urllib.request.urlopen(SITE + '/sitemap.xml', timeout=30) as r:
        xml = r.read().decode('utf-8')
    return re.findall(r'<loc>([^<]+)</loc>', xml)


def load_log():
    if os.path.exists(LOG):
        with open(LOG, encoding='utf-8') as f:
            return set(json.load(f))
    return set()


def save_log(done):
    with open(LOG, 'w', encoding='utf-8') as f:
        json.dump(sorted(done), f, ensure_ascii=False, indent=1)


def push(urls, tok):
    """推一批，返回 (success, remain, raw)。site 原样拼，urlencode 会 400。"""
    body = '\n'.join(urls).encode('utf-8')
    req = urllib.request.Request(API % (SITE, tok), data=body, method='POST')
    req.add_header('Content-Type', 'text/plain')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = json.loads(r.read().decode('utf-8'))
            return data.get('success', 0), data.get('remain', 0), data
    except urllib.error.HTTPError as e:
        return 0, 0, {'http': e.code, 'body': e.read().decode('utf-8', 'replace')}


def main():
    dry = '--dry-run' in sys.argv
    tok = token()
    if not tok and not dry:
        print('没有 token：传参数、设 BAIDU_TOKEN，或写到', SECRET)
        return 1
    all_urls = sitemap_urls()
    done = load_log()
    pending = [u for u in all_urls if u not in done]
    print('sitemap %d 条，已推 %d 条，待推 %d 条' % (len(all_urls), len(done), len(pending)))
    if dry or not pending:
        for u in pending[:10]:
            print('  待推', u)
        return 0

    sent = 0
    while pending:
        batch, pending = pending[:10], pending[10:]
        ok, remain, raw = push(batch, tok)
        print('推 %d 条 → success=%s remain=%s %s' % (len(batch), ok, remain,
                                                     '' if ok else raw))
        if ok:
            done.update(batch[:ok])
            save_log(done)
            sent += ok
        if remain <= 0 or not ok:
            print('今日额度用完或出错，剩下的明天再跑（已推的不重复）')
            break
    print('本次成功 %d 条，累计 %d/%d' % (sent, len(done), len(all_urls)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
