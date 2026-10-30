"""PPT 2010 评分器 - 第8套（前童古镇）"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (PptDeck, Report, run, ppt_presets, ppt_filters,
                   transition_types, layout_short, placeholder_texts,
                   title_text)


class PptGrader8:
    def __init__(self, path):
        self.deck = PptDeck(path)

    def grade(self):
        report = Report('PPT 第8套 · 前童古镇')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5, self.r6):
            report.add(*fn())
        return report

    def slide(self, idx):
        if 0 <= idx < len(self.deck.slides):
            return self.deck.slides[idx]
        return None

    # 1. 第1张 (18)
    def r1(self):
        notes, got = [], 0
        s = self.slide(0)
        if s is None:
            return (1, '第1张版式', 0, 18, ['没有幻灯片'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout1':
            got += 4; notes.append('标题幻灯片 OK')
        else:
            notes.append('版式=%s' % label)
        title = title_text(self.deck, s)
        if '前童古镇' in title:
            got += 5; notes.append('主标题 OK')
        else:
            notes.append('主标题=%r' % title)
        subs = placeholder_texts(self.deck, s, 'subTitle')
        if subs and '低调' in subs[0]['text']:
            got += 4; notes.append('副标题 OK')
        else:
            notes.append('副标题=%r' % (subs[0]['text'][:16] if subs else ''))
        sizes_t = {r['size'] for i in placeholder_texts(self.deck, s, 'ctrTitle') for r in i['runs']}
        sizes_s = {r['size'] for i in placeholder_texts(self.deck, s, 'subTitle') for r in i['runs']}
        if 50.0 in sizes_t:
            got += 3; notes.append('主标题50号 OK')
        else:
            notes.append('主标题字号=%s' % sorted(sizes_t))
        if 30.0 in sizes_s:
            got += 2; notes.append('副标题30号 OK')
        else:
            notes.append('副标题字号=%s' % sorted(sizes_s))
        return (1, '第1张版式', got, 18, notes)

    # 2. 第2-4张 两栏内容 (18)
    def r2(self):
        hit = 0
        for idx in range(1, 4):
            s = self.slide(idx)
            if s is not None and layout_short(self.deck, s)[0] == 'slideLayout4':
                hit += 1
        notes = ['%d/3 张为两栏内容' % hit]
        return (2, '第2-4张版式', hit * 6, 18, notes)

    # 3. 图片 (15)
    def r3(self):
        hit = 0
        for idx in range(1, 4):
            s = self.slide(idx)
            if s is not None and self.deck.image_targets(s):
                hit += 1
        notes = ['%d/3 张有图片' % hit]
        return (3, '插图', hit * 5, 15, notes)

    # 4. 动画 劈裂+飞入 (19)
    def r4(self):
        notes, got = [], 0
        barn = fly = timed = 0
        for idx in range(1, 4):
            s = self.slide(idx)
            if s is None:
                continue
            if s.timing is not None:
                timed += 1
            if 'barn' in ' '.join(ppt_filters(s)):
                barn += 1
            if any(p[0] == '2' and p[1] == 'entr' for p in ppt_presets(s)):
                fly += 1
        got = barn * 5 + fly * 5 + (1 if timed == 3 else 0)
        notes.append('文本劈裂 %d/3，图片飞入 %d/3，时序 %d/3' % (barn, fly, timed))
        return (4, '动画效果', got, 19, notes)

    # 5. 溶解切换 (15)
    def r5(self):
        hits = sum(1 for s in self.deck.slides if 'dissolve' in transition_types(s))
        notes = ['溶解 %d/%d 张' % (hits, len(self.deck.slides))]
        return (5, '切换效果', round(hits * 15.0 / max(1, len(self.deck.slides)), 1), 15, notes)

    # 6. 主题+放映 (15)
    def r6(self):
        notes, got = [], 0
        name = self.deck.theme.get('name', '')
        if '凤舞九天' in name:
            got += 8; notes.append('主题 OK')
        else:
            notes.append('主题=%r（要求 凤舞九天）' % name)
        if self.deck.show_type == 'browse':
            got += 7; notes.append('观众自行浏览 OK')
        else:
            notes.append('放映类型=%r' % self.deck.show_type)
        return (6, '主题与放映', got, 15, notes)


if __name__ == '__main__':
    sys.exit(run(PptGrader8, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'PPT 2010 操作题题库', '8', '1结果.pptx')))
