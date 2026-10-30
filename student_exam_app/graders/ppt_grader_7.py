"""PPT 2010 评分器 - 第7套（人工智能）"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (PptDeck, Report, run, ppt_presets, ppt_filters,
                   transition_types, has_gradfill, hlink_ids, line_spacing,
                   layout_short, title_text, placeholder_texts)


class PptGrader7:
    def __init__(self, path):
        self.deck = PptDeck(path)

    def grade(self):
        report = Report('PPT 第7套 · 人工智能')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9):
            report.add(*fn())
        return report

    def slide(self, idx):
        if 0 <= idx < len(self.deck.slides):
            return self.deck.slides[idx]
        return None

    # 1. 第1张 (15)
    def r1(self):
        notes, got = [], 0
        s = self.slide(0)
        if s is None:
            return (1, '第1张版式', 0, 15, ['没有幻灯片'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout1':
            got += 4; notes.append('标题幻灯片 OK')
        else:
            notes.append('版式=%s' % label)
        title = title_text(self.deck, s)
        if '人工智能' in title:
            got += 5; notes.append('主标题 OK')
        else:
            notes.append('主标题=%r' % title)
        subs = placeholder_texts(self.deck, s, 'subTitle')
        if subs and 'AI' in subs[0]['text']:
            got += 3; notes.append('副标题 AI OK')
        sizes = {r['size'] for item in placeholder_texts(self.deck, s, 'ctrTitle') for r in item['runs']}
        if 50.0 in sizes:
            got += 3; notes.append('主标题50号 OK')
        else:
            notes.append('主标题字号=%s' % sorted(sizes))
        return (1, '第1张版式', got, 15, notes)

    # 2. 第2张 两栏+标题+图片 (15)
    def r2(self):
        notes, got = [], 0
        s = self.slide(1)
        if s is None:
            return (2, '第2张版式', 0, 12, ['缺少第2张'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout4':
            got += 5; notes.append('两栏内容 OK')
        else:
            notes.append('版式=%s' % label)
        if '人工智能的介绍' in title_text(self.deck, s):
            got += 4; notes.append('标题 OK')
        else:
            notes.append('标题=%r' % title_text(self.deck, s))
        if self.deck.image_targets(s):
            got += 3; notes.append('图片存在 OK')
        else:
            notes.append('没有图片')
        return (2, '第2张版式', got, 12, notes)

    # 3. 第3张 空白+艺术字 (15)
    def r3(self):
        notes, got = [], 0
        s = self.slide(2)
        if s is None:
            return (3, '第3张艺术字', 0, 15, ['缺少第3张'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout7':
            got += 5; notes.append('空白版式 OK')
        else:
            notes.append('版式=%s' % label)
        if has_gradfill(s):
            got += 6; notes.append('艺术字渐变填充 OK')
            accent = re.search(r'<a:gradFill[^>]*>.{0,400}?schemeClr val="(accent\d)"', s.xml, re.S)
            if accent and accent.group(1) == 'accent1':
                got += 2; notes.append('强调文字颜色1 OK')
        else:
            notes.append('没有渐变填充')
        art = [item for item in self.deck.resolved_shapes(s) if item['text']]
        if art:
            got += 2; notes.append('艺术字文本=%r' % art[0]['text'][:20])
        return (3, '第3张艺术字', got, 15, notes)

    # 4. 动画 劈裂+飞入 (22)
    def r4(self):
        notes, got = [], 0
        s2 = self.slide(1)
        s3 = self.slide(2)
        barn = 0
        fly = 0
        for s in (s2, s3):
            if s is None:
                continue
            if 'barn' in ' '.join(ppt_filters(s)):
                barn += 1
            if any(p[0] == '2' and p[1] == 'entr' for p in ppt_presets(s)):
                fly += 1
        got = barn * 6 + fly * 5
        notes.append('文本劈裂 %d/2，图片飞入 %d/2' % (barn, fly))
        return (4, '动画效果', got, 17, notes)

    # 5. 主题加粗强调（素材链接为预置内容，不计分）(10)
    def r5(self):
        notes, got = [], 0
        s1 = self.slide(0)
        bold_title = False
        if s1 is not None:
            for item in self.deck.resolved_shapes(s1):
                if item['placeholder'] == 'ctrTitle' and item['runs']:
                    bold_title = all(r['bold'] for r in item['runs'])
        if bold_title:
            got += 10; notes.append('主标题加粗 OK')
        else:
            notes.append('主标题未加粗')
        return (5, '主标题加粗', got, 10, notes)

    # 6. 第3张 动画 (10)
    def r6(self):
        notes, got = [], 0
        s = self.slide(2)
        if s is None:
            return (6, '第3张动画', 0, 10, ['缺少第3张'])
        if s.timing is not None and any(p[1] == 'entr' for p in ppt_presets(s)):
            got += 10; notes.append('艺术字/内容进入动画 OK')
        else:
            notes.append('第3张没有进入动画')
        return (6, '第3张动画', got, 10, notes)

    # 7. 百叶窗切换 (8)
    def r7(self):
        hits = sum(1 for s in self.deck.slides if 'blinds' in transition_types(s))
        if hits == len(self.deck.slides) and hits:
            return (7, '切换效果', 6, 6, ['全部百叶窗 OK'])
        return (7, '切换效果', round(hits * 6.0 / max(1, len(self.deck.slides)), 1),
                6, ['百叶窗 %d/%d 张' % (hits, len(self.deck.slides))])

    # 8. 放映方式 (8)
    def r8(self):
        show = self.deck.show_type
        if show == 'browse':
            return (8, '放映方式', 7, 7, ['观众自行浏览 OK'])
        return (8, '放映方式', 0, 7, ['放映类型=%r' % show])

    # 9. 主题 (7)
    def r9(self):
        name = self.deck.theme.get('name', '')
        if '聚合' in name:
            return (9, '主题', 8, 8, ['主题=%s OK' % name])
        return (9, '主题', 0, 8, ['主题=%r（要求 聚合）' % name])


if __name__ == '__main__':
    sys.exit(run(PptGrader7, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        '2022年准易：Ppt 2010 操作题--素材', '7', '1结果.pptx')))
