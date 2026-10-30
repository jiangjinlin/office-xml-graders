"""PPT 2010 评分器 - 第9套（天宫课堂）"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (PptDeck, Report, run, ppt_presets, ppt_filters,
                   layout_short, placeholder_texts, title_text)


class PptGrader9:
    def __init__(self, path):
        self.deck = PptDeck(path)

    def grade(self):
        report = Report('PPT 第9套 · 天宫课堂')
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
            got += 3; notes.append('标题幻灯片 OK')
        else:
            notes.append('版式=%s' % label)
        title = title_text(self.deck, s)
        if '天宫课堂' in title:
            got += 4; notes.append('主标题 OK')
        else:
            notes.append('主标题=%r' % title)
        subs = placeholder_texts(self.deck, s, 'subTitle')
        if subs and '太空授课' in subs[0]['text']:
            got += 4; notes.append('副标题 OK')
        runs_t = [r for i in placeholder_texts(self.deck, s, 'ctrTitle') for r in i['runs']]
        runs_s = [r for i in placeholder_texts(self.deck, s, 'subTitle') for r in i['runs']]
        if any(r['font'] == '隶书' for r in runs_t):
            got += 3; notes.append('隶书 OK')
        else:
            notes.append('主标题字体=%s' % sorted({r['font'] for r in runs_t}))
        if 60.0 in {r['size'] for r in runs_t}:
            got += 2; notes.append('60号 OK')
        if any(r['font'] == '华文中宋' for r in runs_s) and 30.0 in {r['size'] for r in runs_s}:
            got += 2; notes.append('副标题华文中宋30 OK')
        return (1, '第1张版式', got, 18, notes)

    # 2. 第2张 图片+标题 (17)
    def r2(self):
        notes, got = [], 0
        s = self.slide(1)
        if s is None:
            return (2, '第2张插图', 0, 17, ['缺少第2张'])
        name, label = layout_short(self.deck, s)
        if '太空授课' in title_text(self.deck, s):
            got += 6; notes.append('标题 OK')
        else:
            notes.append('标题=%r' % title_text(self.deck, s))
        images = self.deck.image_targets(s)
        if images:
            got += 6; notes.append('图片 OK')
        else:
            notes.append('没有图片')
        if images and any(i.endswith('.jpg') for i in images):
            got += 5; notes.append('图片1.jpg OK')
        return (2, '第2张插图', got, 17, notes)

    # 3. 第2张 动画 (17)
    def r3(self):
        notes, got = [], 0
        s = self.slide(1)
        if s is None:
            return (3, '第2张动画', 0, 17, ['缺少第2张'])
        text_ok = any(p[0] == '15' and p[1] == 'entr' for p in ppt_presets(s)) \
            or 'box' in ' '.join(ppt_filters(s))
        if text_ok:
            got += 8; notes.append('文字"形状"动画 OK')
        else:
            notes.append('未找到文字形状动画 presets=%s' % sorted(ppt_presets(s)))
        img_ok = any(p[0] == '4' and p[1] == 'entr' for p in ppt_presets(s))
        if img_ok:
            got += 9; notes.append('图片螺旋飞入 OK')
        else:
            notes.append('未找到图片动画')
        return (3, '第2张动画', got, 17, notes)

    # 4. 新建第3张 (17)
    def r4(self):
        notes, got = [], 0
        s = self.slide(2)
        if s is None:
            return (4, '新建幻灯片', 0, 17, ['只有2张'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout2':
            got += 4; notes.append('标题和内容 OK')
        else:
            notes.append('版式=%s' % label)
        if '王亚平' in title_text(self.deck, s):
            got += 6; notes.append('标题 OK')
        else:
            notes.append('标题=%r' % title_text(self.deck, s))
        images = self.deck.image_targets(s)
        if images and any(i.endswith('.png') for i in images):
            got += 7; notes.append('图片2.png OK')
        elif images:
            got += 4; notes.append('有图片但=%s' % images)
        else:
            notes.append('没有图片')
        return (4, '新建幻灯片', got, 17, notes)

    # 5. 第3张 动画 (16)
    def r5(self):
        notes, got = [], 0
        s = self.slide(2)
        if s is None:
            return (5, '第3张动画', 0, 16, ['缺少第3张'])
        img_ok = any(p[0] == '2' and p[2] == '6' for p in ppt_presets(s))
        if img_ok:
            got += 8; notes.append('图片飞入(自右下部) OK')
        else:
            notes.append('未找到飞入自右下 presets=%s' % sorted(ppt_presets(s)))
        text_ok = any(p[0] == '47' for p in ppt_presets(s)) or 'fade' in ' '.join(ppt_filters(s))
        if text_ok:
            got += 8; notes.append('文字浮入 OK')
        else:
            notes.append('未找到浮入')
        return (5, '第3张动画', got, 16, notes)

    # 6. 主题+放映 (15)
    def r6(self):
        notes, got = [], 0
        name = self.deck.theme.get('name', '')
        if '聚合' in name:
            got += 8; notes.append('主题 OK')
        else:
            notes.append('主题=%r（要求 聚合）' % name)
        if self.deck.show_type == 'browse':
            got += 7; notes.append('观众自行浏览 OK')
        else:
            notes.append('放映类型=%r' % self.deck.show_type)
        return (6, '主题与放映', got, 15, notes)


if __name__ == '__main__':
    sys.exit(run(PptGrader9, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        '2022年准易：Ppt 2010 操作题--素材', '9', '1结果.pptx')))
