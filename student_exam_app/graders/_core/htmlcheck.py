"""网页(HTML)评分辅助：GB2312/GBK/UTF-8 容错读取 + 宽松属性匹配。"""
import os
import re


def read_html(path):
    """按 gb2312→gbk→utf-8 顺序解码，返回 (text, encoding)。"""
    data = open(path, 'rb').read()
    for enc in ('gb2312', 'gbk', 'utf-8'):
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace'), 'utf-8'


def _tag_bodies(html, tag):
    """返回该标签所有完整开标签字符串（小写化），如 '<img src="a.gif" width="150">'。"""
    return re.findall(r'<%s\b[^>]*>' % tag, html, re.I)


def _attr_ok(tag_text, name, value):
    """属性匹配：值不区分大小写；value 为 True 表示只要求存在。"""
    if value is True:
        return re.search(r'\b%s\b\s*(?:=\s*[^ >]+)?' % re.escape(name), tag_text, re.I) is not None
    m = re.search(r'\b%s\s*=\s*("([^"]*)"|\'([^\']*)\'|([^\s>]+))' % re.escape(name),
                  tag_text, re.I)
    if not m:
        return False
    got = (m.group(2) or m.group(3) or m.group(4) or '').strip().strip('"\'')
    return got.lower() == str(value).strip().lower()


def find_tags(html, tag, **attrs):
    """找属性全部满足的开标签；attrs {'src':'image/a.gif','noshade':True}。"""
    hits = []
    for t in _tag_bodies(html, tag):
        if all(_attr_ok(t, k, v) for k, v in attrs.items()):
            hits.append(t)
    return hits


def has_tag(html, tag, **attrs):
    return bool(find_tags(html, tag, **attrs))


def title_of(html):
    m = re.search(r'<title[^>]*>(.*?)</title>', html, re.I | re.S)
    return re.sub(r'\s+', '', m.group(1)) if m else ''


def visible_text(html):
    """去 style/script 后剥标签，压缩空白。"""
    t = re.sub(r'<(style|script)\b.*?</\1>', ' ', html, flags=re.I | re.S)
    t = re.sub(r'<[^>]+>', ' ', t)
    t = t.replace('&nbsp;', ' ').replace('&gt;', '>').replace('&lt;', '<')
    return re.sub(r'\s+', '', t)


def css_text(html):
    m = re.findall(r'<style[^>]*>(.*?)</style>', html, re.I | re.S)
    return re.sub(r'\s+', '', '\n'.join(m)).lower()


def body_attr(html, name, value):
    for t in _tag_bodies(html, 'body'):
        if _attr_ok(t, name, value):
            return True
    return False


def frameset_attrs(html):
    return [dict(re.findall(r'([\w-]+)\s*=\s*"([^"]*)"', t)) for t in _tag_bodies(html, 'frameset')]


def frame_list(html):
    out = []
    for t in _tag_bodies(html, 'frame'):
        out.append(dict(re.findall(r'([\w-]+)\s*=\s*"([^"]*)"', t)))
    return out


def frame_srcs(html):
    return [f.get('src', '').lower() for f in frame_list(html)]


def select_info(html):
    """返回 (option 文本列表, selected 文本或 '')。"""
    m = re.search(r'<select\b[^>]*>(.*?)</select>', html, re.I | re.S)
    if not m:
        return [], ''
    block = m.group(1)
    opts = []
    selected = ''
    for om in re.finditer(r'<option\b([^>]*)>([^<]*)', block, re.I):
        attrs = om.group(1)
        text = re.sub(r'\s+', '', om.group(2))
        opts.append(text)
        if re.search(r'\bselected\b', attrs, re.I):
            selected = text
    return opts, selected


def anchor_names(html):
    return [a.get('name', '') for a in
            (dict(re.findall(r'([\w-]+)\s*=\s*"([^"]*)"', t)) for t in _tag_bodies(html, 'a'))
            if a.get('name')]


def collect_html_files(root):
    """递归收集 {文件名小写: 绝对路径}，用于按文件名定位页面。"""
    out = {}
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(('.html', '.htm')):
                out.setdefault(fn.lower(), os.path.join(dp, fn))
    return out


def locate_site_root(root):
    """找包含 index.html 的最浅目录作为站点根。"""
    best = None
    for dp, _, fns in os.walk(root):
        if any(fn.lower() == 'index.html' for fn in fns):
            depth = dp.replace(root, '').count(os.sep)
            if best is None or depth < best[0]:
                best = (depth, dp)
    return best[1] if best else root
