"""打印 pptx 的关键结构，用于对照题干写评分器。

用法: python tools/dump_pptx.py <文件>
"""
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, 'graders'))

from _core import PptDeck  # noqa: E402

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

LAYOUT_MAP = {
    'slideLayout1': '标题幻灯片', 'slideLayout2': '标题和内容',
    'slideLayout3': '节标题', 'slideLayout4': '两栏内容',
    'slideLayout5': '比较', 'slideLayout6': '仅标题',
    'slideLayout7': '空白', 'slideLayout8': '内容与标题',
    'slideLayout9': '图片与标题', 'slideLayout10': '竖排标题与文本',
}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    with PptDeck(sys.argv[1]) as deck:
        theme = deck.theme
        print('主题=%s 放映=%s 尺寸=%.1fx%.1f 页数=%d 主题字体=%s' % (
            theme.get('name'), deck.show_type, deck.slide_size[0], deck.slide_size[1],
            len(deck.slides), theme.get('fonts')))
        for slide in deck.slides:
            layout = slide.layout_of(deck)
            print('\n== #%d %s 版式=%s(%s) 切换=%s spd=%s' % (
                slide.index + 1, slide.part, LAYOUT_MAP.get(layout.split('/')[-1].replace('.xml', ''), ''),
                layout.split('/')[-1], slide.transition,
                (slide.transition_node.get('spd') if slide.transition_node is not None else '')))
            print('   占位符:', [(sh.name[:12], sh.placeholder, sh.placeholder_idx) for sh in slide.shapes if sh.placeholder])
            for item in deck.resolved_shapes(slide):
                if item['text']:
                    runs = item['runs']
                    print('   [%s] %r fonts=%s sizes=%s bold=%s' % (
                        item['placeholder'] or item['kind'],
                        item['text'][:24],
                        sorted({r['font'] for r in runs if r['font']}),
                        sorted({r['size'] for r in runs if r['size']}),
                        sorted({r['bold'] for r in runs})))
            images = deck.image_targets(slide)
            if images:
                print('   图片:', images)
            arts = [sh for sh in slide.shapes if sh.kind == 'sp' and not sh.placeholder
                    and sh.name and ('TextBox' in sh.name or '艺术' in sh.name)]
            timing = slide.timing
            if timing is not None:
                effects = []
                for node in timing.iter():
                    name = node.tag.split('}')[-1]
                    if name in ('animEffect', 'anim', 'animRot', 'animScale', 'animMotion'):
                        info = ''
                        for child in node.iter():
                            cn = child.tag.split('}')[-1]
                            if cn == 'filter':
                                info += ' filter=' + (child.get('filter') or '')
                        eff = node.find('.//{%s}cBhvr' % P)
                        preset = None
                        parent = node
                        effs = [n for n in node.iter('{%s}cTn' % P)]
                        for n in node.iter('{%s}cTn' % P):
                            pid = n.get('presetID')
                            if pid:
                                preset = (pid, n.get('presetClass'), n.get('presetSubtype'), n.get('nodeType'))
                        effects.append(name + info + (' preset=%s' % (preset,) if preset else ''))
                if effects:
                    print('   动画:', effects[:6])
            hyper = re.findall(r'hlinkClick[^>]*action="ppaction://hlinkshowjump\?jump=([^"]+)"', slide.xml)
            hlink2 = re.findall(r'<a:hlinkClick[^>]*r:id="(rId\d+)"', slide.xml)
            if hyper or hlink2:
                print('   链接:', hyper, hlink2)
            notes = deck.notes_text(slide)
            if notes:
                print('   备注:', notes[:40])
    return 0


if __name__ == '__main__':
    sys.exit(main())
