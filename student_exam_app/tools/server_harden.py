#!/usr/bin/env python3
"""web.xiaoshizi.asia 加固：限流区、响应头、超时、禁 /uploads、TLS1.2+。只动本 vhost。"""
import time

NGX = '/www/server/nginx/conf/nginx.conf'
VH = '/www/server/panel/vhost/nginx/python_student_exam_app.conf'
stamp = time.strftime('%m%d%H%M')

s = open(NGX).read()
if 'zone=exam_api' not in s:
    s = s.replace('limit_conn_zone $server_name zone=perserver:10m;',
                  'limit_conn_zone $server_name zone=perserver:10m;\n'
                  '        limit_req_zone $binary_remote_addr zone=exam_api:10m rate=20r/m;\n'
                  '        limit_req_status 429;', 1)
    open(NGX + '.bak' + stamp, 'w').write(open(NGX).read())
    open(NGX, 'w').write(s)
    print('nginx.conf 加了 limit_req_zone')

v = open(VH).read()
backup = VH + '.bak' + stamp
open(backup, 'w').write(v)

v = v.replace('ssl_protocols TLSv1.1 TLSv1.2 TLSv1.3;', 'ssl_protocols TLSv1.2 TLSv1.3;')

proxy_headers = """        proxy_set_header Host $http_host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;"""

if 'location ^~ /uploads/' not in v:
    v = v.replace("""    # HTTP反向代理相关配置开始 >>>""",
"""    # 安全加固（本 vhost 专用）
    client_max_body_size 52m;
    add_header X-Content-Type-Options nosniff always;
    add_header X-Frame-Options SAMEORIGIN always;
    add_header Referrer-Policy same-origin always;
    add_header X-Permitted-Cross-Domain-Policies none always;
    proxy_hide_header X-Powered-By;
    proxy_hide_header Server;

    # 上传临时目录不对外
    location ^~ /uploads/ {
        return 404;
    }

    # 判分接口单 IP 限流（每分钟 20 次，突发 8）
    location ^~ /api/grade/ {
        limit_req zone=exam_api burst=8 nodelay;
        proxy_pass http://127.0.0.1:5000;
%s
        proxy_connect_timeout 10s;
        proxy_send_timeout 120s;
        proxy_read_timeout 120s;
    }

    # HTTP反向代理相关配置开始 >>>""" % proxy_headers, 1)

    v = v.replace("""        proxy_set_header Remote-Host $remote_addr;
    }""",
"""        proxy_set_header Remote-Host $remote_addr;
        proxy_connect_timeout 10s;
        proxy_send_timeout 120s;
        proxy_read_timeout 120s;
    }""", 1)

open(VH, 'w').write(v)
print('vhost 已加固，备份:', backup)
