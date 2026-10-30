"""pptx 评分辅助：动画预设、切换、动作按钮、超链接、艺术字渐变。"""
import re

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

LAYOUT_MAP = {
    'slideLayout1': '标题幻灯片', 'slideLayout2': '标题和内容',
    'slideLayout3': '节标题', 'slideLayout4': '两栏内容',
    'slideLayout5': '比较', 'slideLayout6': '仅标题', 'slideLayout7': '空白',
    'slideLayout8': '内容与标题', 'slideLayout9': '图片与标题',
    'slideLayout10': '竖排标题与文本', 'slideLayout11': '标题和竖排文本',
}


def presets(slide):
    """set[(presetID, presetClass, presetSubtype)]"""
    return set(re.findall(r'presetID="(\d+)"[^>]*presetClass="(\w+)"[^>]*presetSubtype="([-\d]+)"',
                          slide.xml))


def has_preset(slide, pid=None, pclass=None, subtype=None):
    for pid_x, cls_x, sub_x in presets(slide):
        if pid is not None and str(pid) != str(pid_x):
            continue
        if pclass is not None and pclass != cls_x:
            continue
        if subtype is not None and str(subtype) != str(sub_x):
            continue
        return True
    return False


def filters(slide):
    return set(re.findall(r'filter="([^"]+)"', slide.xml))


def transition_types(slide):
    """兼容 mc:Choice/Fallback 与 p14 扩展。"""
    return re.findall(r'<(?:p:)?transition[^>]*>\s*<p(?:14)?:(\w+)', slide.xml)


def transition_attrs(slide):
    hit = re.search(r'<(?:p:)?transition([^>]*)>', slide.xml)
    return hit.group(1) if hit else ''


def actions(slide):
    return set(re.findall(r'action="([^"]+)"', slide.xml))


def shape_names(slide):
    return [sh.name for sh in slide.shapes]


def hlink_ids(slide):
    return re.findall(r'<a:hlinkClick[^>]*r:id="(rId\d+)"', slide.xml)


def has_gradfill(slide):
    return '<a:gradFill' in slide.xml


def auto_number(slide):
    return '<a:buAutoNum' in slide.xml


def line_spacing(slide):
    return re.findall(r'<a:lnSpc><a:spcPct val="(\d+)"', slide.xml)


def layout_short(deck, slide):
    part = slide.layout_of(deck)
    name = part.split('/')[-1].replace('.xml', '')
    return name, LAYOUT_MAP.get(name, name)


def placeholder_texts(deck, slide, *types):
    out = []
    for item in deck.resolved_shapes(slide):
        if item['placeholder'] in types and item['text']:
            out.append(item)
    return out


def title_text(deck, slide):
    for item in deck.resolved_shapes(slide):
        if item['placeholder'] in ('title', 'ctrTitle'):
            return item['text']
    return ''
