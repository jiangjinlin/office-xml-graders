"""评分时的容错比较工具：同一要求允许多种实现方式。"""
import re


def norm_formula(text):
    """统一公式写法：去绝对引用符、空格、引号差异，转大写。"""
    if not text:
        return ''
    out = text.replace('$', '').replace(' ', '').replace('　', '')
    out = out.replace('"', '"').replace('"', '"')
    return out.upper()


def formula_has(formula, *keywords):
    text = norm_formula(formula)
    return any(k.upper() in text for k in keywords)


def func_name(formula):
    hit = re.match(r'=?\s*([A-Za-z_.]+)\s*\(', formula or '')
    return hit.group(1).upper() if hit else ''


def style_name(fmt):
    return (fmt.get('style_name') or '').replace(' ', '').replace('　', '')


def style_name_is(fmt, *names):
    target = style_name(fmt)
    return any(target == n.replace(' ', '') for n in names)


def as_float(value, default=None):
    try:
        return float(str(value).replace(',', '').strip())
    except (TypeError, ValueError):
        return default


def close(a, b, tol=0.01):
    fa, fb = as_float(a), as_float(b)
    if fa is None or fb is None:
        return False
    return abs(fa - fb) <= tol


def color_of(node):
    """返回 RRGGBB（大写），兼容 theme 色。"""
    if not node:
        return ''
    rgb = node.get('rgb') or ''
    if rgb:
        return rgb[-6:].upper()
    theme = node.get('theme')
    if theme is not None:
        return 'THEME%s' % theme
    return ''


def font_color(fmt):
    return color_of(fmt.get('font', {}).get('color'))


def fill_color(fmt):
    fill = fmt.get('fill', {})
    return color_of(fill.get('fg')) or color_of(fill.get('bg'))


def border_styles(fmt):
    """[(side, style, color)] 只保留有线的边。"""
    out = []
    for side, info in (fmt.get('border') or {}).items():
        if info.get('style'):
            out.append((side, info['style'], color_of(info.get('color'))))
    return out


def has_border(fmt, styles, colors=None):
    for _side, style, color in border_styles(fmt):
        if style in styles and (not colors or color in colors):
            return True
    return False


def norm_ref(ref):
    return (ref or '').replace('$', '').upper()


def ref_range_covers(ref, target):
    """判断 'A1:H21' 这类区域是否覆盖 target（简单矩形判断）。"""
    def parse(text):
        hit = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', norm_ref(text))
        if not hit:
            hit2 = re.match(r'([A-Z]+)(\d+)', norm_ref(text))
            if not hit2:
                return None
            col = col_num(hit2.group(1)); row = int(hit2.group(2))
            return (col, row, col, row)
        return (col_num(hit.group(1)), int(hit.group(2)),
                col_num(hit.group(3)), int(hit.group(4)))

    outer, inner = parse(ref), parse(target)
    if not outer or not inner:
        return False
    return (outer[0] <= inner[0] and outer[1] <= inner[1]
            and outer[2] >= inner[2] and outer[3] >= inner[3])


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    out = ''
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def has_field(xml_text, name='PAGE'):
    """判断是否含 Word 域代码（w:instrText 或 w:fldSimple/@w:instr）。

    Word 把 { PAGE } 存成 <w:instrText> PAGE </w:instrText>，
    或 <w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">，
    直接搜 'PAGE' 会命中 NUMPAGES、正文单词和自定义 XML。
    """
    name = name.upper()
    for code in re.findall(r'<(?:\w+:)?instrText[^>]*>([^<]*)</(?:\w+:)?instrText>', xml_text):
        if field_name(code) == name:
            return True
    for instr in re.findall(r'<(?:\w+:)?fldSimple[^>]*\binstr="([^"]*)"', xml_text):
        if field_name(instr) == name:
            return True
    return False


def field_name(code):
    """从域代码里取出域名：' PAGE \\* MERGEFORMAT ' -> 'PAGE'。"""
    parts = [p for p in re.split(r'[\s\\"]+', code or '') if p]
    for token in parts:
        if re.match(r'^[A-Za-z]+$', token):
            return token.upper()
    return ''


def _tags(xml_text, name):
    """prefix 无关地取标签，返回 (attrs_str, inner) 列表。"""
    pat = r'<(?:\w+:)?%s\b([^>]*)>(.*?)</(?:\w+:)?%s>' % (name, name)
    return re.findall(pat, xml_text, re.S)


def mail_merge_info(xml_text):
    """取 w:mailMerge 配置，返回 dict 或 None。

    settings.xml 里的结构：
      <w:mailMerge><w:mainDocumentType w:val="formLetters"/>
        <w:dataType w:val="native"/><w:connectString w:val="Provider=..."/>
        <w:query w:val="SELECT * FROM ..."/>...
    """
    m = re.search(r'<(?:\w+:)?mailMerge\b[^>]*>(.*?)</(?:\w+:)?mailMerge>', xml_text, re.S)
    if not m:
        return None
    block = m.group(1)
    out = {}
    for key in ('mainDocumentType', 'dataType', 'connectString', 'query',
                'dataSource', 'headerSource', 'linkToQuery', 'destination'):
        mm = re.search(r'<(?:\w+:)?%s\b[^>]*?\bval="([^"]*)"' % key, block, re.S)
        if mm:
            out[key] = mm.group(1)
    return out or {'present': True}


def merge_fields(xml_text):
    """取邮件合并域名称列表（MERGEFIELD 后面的域名）。"""
    names = []
    for _attrs, inner in _tags(xml_text, 'instrText'):
        if 'MERGEFIELD' not in inner.upper():
            continue
        m = re.search(r'MERGEFIELD\s+"?([^"\\*]+)', inner, re.I)
        if m:
            names.append(m.group(1).strip())
    for instr in re.findall(r'<(?:\w+:)?fldSimple\b[^>]*\binstr="([^"]*)"', xml_text):
        if 'MERGEFIELD' in instr.upper():
            m = re.search(r'MERGEFIELD\s+"?([^"\\*]+)', instr, re.I)
            if m:
                names.append(m.group(1).strip())
    return names


def pictures(xml_text):
    """收集文档里的图片，返回 [{rid, cx, cy, anchor, wrap}]。

    DrawingML: <w:drawing><wp:inline|wp:anchor><wp:extent cx cy><a:blip r:embed>
    VML(旧式/剪贴画): <w:pict><v:shape><v:imagedata r:id>
    cx/cy 单位 EMU，1 厘米 = 360000。
    """
    out = []
    for kind in ('inline', 'anchor'):
        for attrs, inner in _tags(xml_text, kind):
            rid = re.search(r'\br:embed="([^"]+)"', inner) or re.search(r'\br:id="([^"]+)"', inner)
            ext = re.search(r'<(?:\w+:)?extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"', inner)
            wrap = 'none'
            for w in ('wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'wrapNone'):
                if re.search(r'<(?:\w+:)?%s\b' % w, inner):
                    wrap = w
                    break
            out.append({'rid': rid.group(1) if rid else '',
                        'cx': int(ext.group(1)) if ext else 0,
                        'cy': int(ext.group(2)) if ext else 0,
                        'anchor': kind == 'anchor', 'wrap': wrap})
    for attrs, inner in _tags(xml_text, 'pict'):
        for sh_attrs, sh_inner in _tags(inner, 'shape'):
            img = re.search(r'<(?:\w+:)?imagedata\b[^>]*\br:id="([^"]+)"', sh_inner)
            if img:
                style = re.search(r'style="([^"]*)"', sh_attrs)
                out.append({'rid': img.group(1), 'cx': 0, 'cy': 0, 'anchor': True,
                            'wrap': 'vml', 'style': style.group(1) if style else ''})
    return out


def vml_shapes(xml_text):
    """取 VML 形状（文本框、艺术字、水印、椭圆标注都是这个）。"""
    out = []
    for attrs, inner in _tags(xml_text, 'shape'):
        out.append({'attrs': attrs, 'inner': inner,
                    'id': (re.search(r'id="([^"]*)"', attrs) or [None, ''])[1],
                    'style': (re.search(r'style="([^"]*)"', attrs) or [None, ''])[1],
                    'type': (re.search(r'type="([^"]*)"', attrs) or [None, ''])[1]})
    return out
