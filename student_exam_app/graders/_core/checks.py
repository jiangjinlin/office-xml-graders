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
