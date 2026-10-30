"""DW 网页题评分引擎：解压学生上传的 zip → 定位站点 → 按规范逐项检查。

规范(dw_specs.SPECS)每项:
  {'title': 报告标题, 'items': [
      {'name': 显示名, 'points': 分值,
       'checks': [检查元组...], 'labels': {检查元组: 失败提示}}]}

检查元组 (op, file, ...)：
  ('title',   f, value)            <title> 等于 value（去空白）
  ('text',    f, value)            可见文本包含 value
  ('css',     f, *frags)           <style> 内容包含全部片段
  ('body',    f, attr, value)      <body> 属性
  ('tag',     f, tag, **attrs)     标签存在且属性满足
  ('attr',    f, tag, attr, value) 某标签的某属性值相等（如 frameset rows）
  ('frame',   f, src)              <frame src> 含 src
  ('noframe', f, src)              <frame src> 不含 src
  ('select',  f, options, sel)     下拉选项与选中项；options 可为 []
  ('anchor',  f, name)             <a name=...>
  ('file',    name)                站点内存在该文件
"""
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import htmlcheck as H  # noqa: E402
from _core import Report, run, setup_stdout  # noqa: E402


def run_check(files, chk):
    op = chk[0]
    if op == 'file':
        return chk[1].lower() in files
    text, _ = _read(files, chk[1])
    if text is None:
        return False
    if op == 'title':
        return H.title_of(text) == re.sub(r'\s+', '', chk[2])
    if op == 'text':
        return chk[2] in H.visible_text(text)
    if op == 'css':
        css = H.css_text(text)
        return all(re.sub(r'\s+', '', f).lower() in css for f in chk[2:])
    if op == 'body':
        return H.body_attr(text, chk[2], chk[3])
    if op == 'tag':
        return H.has_tag(text, chk[2], **chk[3])
    if op == 'attr':
        want = str(chk[4]).strip().lower()
        if chk[2] == 'frameset':
            attrs_list = H.frameset_attrs(text)
        else:
            attrs_list = [dict(re.findall(r'([\w-]+)\s*=\s*"([^"]*)"', t))
                          for t in re.findall(r'<%s\b[^>]*>' % chk[2], text, re.I)]
        return any(str(a.get(chk[3], '')).strip().lower() == want for a in attrs_list)
    if op == 'frame':
        return any(chk[2].lower() in s for s in H.frame_srcs(text))
    if op == 'noframe':
        return all(chk[2].lower() not in s for s in H.frame_srcs(text))
    if op == 'select':
        want, sel = chk[2], chk[3]
        if isinstance(want, str):
            want = [want]
        for block in re.findall(r'<select\b[^>]*>.*?</select>', text, re.I | re.S):
            opts, selected = H.select_info(block)
            if want and not all(w in opts for w in want):
                continue
            if sel and selected != sel:
                continue
            if opts:
                return True
        return False
    if op == 'anchor':
        return chk[2] in H.anchor_names(text)
    if op == 'meta':
        for t in re.findall(r'<meta\b[^>]*>', text, re.I):
            if H._attr_ok(t, 'name', chk[2]) and H._attr_ok(t, 'content', chk[3]):
                return True
        return False
    if op == 'area':
        return H.has_tag(text, 'area', **chk[2])
    if op == 'date':
        return '#BeginDate' in text
    if op == 'a_text':
        for m in re.finditer(r'<a\b[^>]*>(.*?)</a>', text, re.I | re.S):
            tag, inner = m.group(0), m.group(1)
            if chk[3] and not H._attr_ok(tag, 'href', chk[3]):
                continue
            if chk[2] in H.visible_text(inner):
                return True
        return False
    if op == 'a_img':
        for m in re.finditer(r'<a\b[^>]*>(.*?)</a>', text, re.I | re.S):
            tag, inner = m.group(0), m.group(1)
            if not H._attr_ok(tag, 'href', chk[2]):
                continue
            if H.has_tag(inner, 'img', src=chk[3]):
                return True
        return False
    if op == 'css_any':
        css = H.css_text(text)
        return any(re.sub(r'\s+', '', f).lower() in css for f in chk[2:])
    raise ValueError('unknown check: %r' % (chk,))


def _read(files, name):
    path = files.get(name.lower())
    if not path:
        return None, None
    text, _enc = H.read_html(path)
    return text, path


class DwGrader:
    def __init__(self, path, spec, title=''):
        self.path = path
        self.spec = spec
        self.title = title or spec.get('title', '网页操作题评分')

    def grade(self):
        report = Report(self.title)
        work = self.path + '_x'
        shutil.rmtree(work, ignore_errors=True)
        try:
            root = self._site_root(work)
            if root is None:
                report.add(1, '站点文件', 0, 100,
                           ['上传内容里没有找到包含 index.html 的站点文件夹，'
                            '请把整个文件夹压缩成 zip 后上传'])
                return report
            files = H.collect_html_files(root)
            for idx, item in enumerate(self.spec['items'], 1):
                full = float(item['points'])
                per = full / len(item['checks'])
                got = 0.0
                detail = []
                for chk in item['checks']:
                    ok = run_check(files, chk)
                    if ok:
                        got += per
                    label = _label(item, chk)
                    detail.append('%s %s' % ('OK ' if ok else 'NG ', label))
                report.add(idx, item['name'], got, full, detail)
        finally:
            shutil.rmtree(work, ignore_errors=True)
        return report

    def _site_root(self, work):
        if zipfile.is_zipfile(self.path):
            with zipfile.ZipFile(self.path) as zf:
                zf.extractall(work)
            return H.locate_site_root(work)
        if os.path.isdir(self.path):
            return H.locate_site_root(self.path)
        return None


def _label(item, chk):
    """优先取 labels 里按位置对应的提示，否则用检查自身的简述。"""
    labels = item.get('labels') or []
    checks = item['checks']
    i = checks.index(chk) if chk in checks else -1
    if 0 <= i < len(labels):
        return labels[i]
    if chk[0] == 'file':
        return '文件 %s' % chk[1]
    if chk[0] == 'title':
        return '标题=%s' % chk[2]
    if chk[0] == 'text':
        return '包含文字「%s」' % chk[2]
    if chk[0] == 'css':
        return '样式含 %s' % '/'.join(chk[2:])
    if chk[0] == 'body':
        return 'body %s=%s' % (chk[2], chk[3])
    if chk[0] == 'tag':
        attrs = ' '.join('%s=%s' % (k, v) for k, v in chk[3].items()) or '存在'
        return '<%s> %s' % (chk[2], attrs)
    if chk[0] == 'attr':
        return '%s %s=%s' % (chk[2], chk[3], chk[4])
    if chk[0] == 'frame':
        return '框架指向 %s' % chk[2]
    if chk[0] == 'noframe':
        return '框架不再指向 %s' % chk[2]
    if chk[0] == 'select':
        return '下拉框选项/默认选中'
    if chk[0] == 'anchor':
        return '锚点 %s' % chk[2]
    if chk[0] == 'meta':
        return 'meta %s=%s' % (chk[2], chk[3])
    if chk[0] == 'area':
        return '热区 %s' % ' '.join('%s=%s' % (k, v) for k, v in chk[2].items())
    if chk[0] == 'date':
        return '自动更新日期'
    if chk[0] == 'a_text':
        return '「%s」链接 %s' % (chk[2], chk[3] or '')
    if chk[0] == 'a_img':
        return '图片链接 %s → %s' % (chk[3], chk[2])
    if chk[0] == 'css_any':
        return '样式含 %s' % ' / '.join(chk[2:])
    return str(chk)


def main(spec):
    # run() 内部会做 setup_stdout，这里不要重复包一层（会关掉底层 buffer）
    return run(lambda p: DwGrader(p, spec), title=spec.get('title', ''))
