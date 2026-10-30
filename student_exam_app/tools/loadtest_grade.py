# -*- coding: utf-8 -*-
"""模拟机房：同一出口 IP 并发提交 20 份，看有没有被限流误杀。"""
import json
import os
import time
import urllib.request
import concurrent.futures as cf
import sys

URL = 'https://web.xiaoshizi.asia/api/grade/ppt/7'
FILE = os.path.join(os.environ['TEMP'], 'p7.pptx')
N = int(sys.argv[1]) if len(sys.argv) > 1 else 20


def one(i):
    t0 = time.time()
    with open(FILE, 'rb') as f:
        data = f.read()
    boundary = '----wb%d' % i
    body = (
        ('--%s\r\nContent-Disposition: form-data; name="file"; filename="p%d.pptx"\r\n'
         'Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n'
         % (boundary, i)).encode() + data + ('\r\n--%s--\r\n' % boundary).encode())
    req = urllib.request.Request(URL, data=body, method='POST')
    req.add_header('Content-Type', 'multipart/form-data; boundary=%s' % boundary)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            d = json.loads(r.read().decode())
            return r.status, round(time.time() - t0, 1), d.get('score')
    except urllib.error.HTTPError as e:
        return e.code, round(time.time() - t0, 1), e.read()[:80].decode('utf-8', 'replace')
    except Exception as e:
        return 0, round(time.time() - t0, 1), str(e)[:60]


t0 = time.time()
with cf.ThreadPoolExecutor(max_workers=N) as pool:
    results = list(pool.map(one, range(N)))
codes = {}
for code, cost, _ in results:
    codes[code] = codes.get(code, 0) + 1
print('总耗时 %.1fs' % (time.time() - t0))
print('状态码分布:', codes)
print('耗时: 最短 %.1fs 最长 %.1fs' % (min(r[1] for r in results), max(r[1] for r in results)))
bad = [r for r in results if r[0] != 200]
for b in bad[:5]:
    print('失败样本:', b)
