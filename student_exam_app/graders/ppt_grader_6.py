"""PPT 2010 评分器 - 第6套（北京冬奥）"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (PptDeck, Report, run, ppt_actions, shape_names, hlink_ids,
                   has_gradfill, auto_number, transition_types, layout_short,
                   placeholder_texts, title_text)


class PptGrader6:
    def __init__(self, path):
        self.deck = PptDeck(path)

    def grade(self):
        report = Report('PPT 第6套 · 北京冬奥')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9):
            report.add(*fn())
        return report

    def slide(self, idx):
        if 0 <= idx < len(self.deck.slides):
            return self.deck.slides[idx]
        return None

    # 1. 第1张 标题幻灯片+主标题 (12)
    def r1(self):
        notes, got = [], 0
        s = self.slide(0)
        if s is None:
            return (1, '第1张版式', 0, 12, ['没有幻灯片'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout1':
            got += 5; notes.append('标题幻灯片 OK')
        else:
            notes.append('版式=%s' % label)
        title = title_text(self.deck, s)
        if '北京冬奥' in title:
            got += 7; notes.append('主标题 OK')
        else:
            notes.append('主标题=%r' % title)
        return (1, '第1张版式', got, 12, notes)

    # 2. 第2张 标题和内容+编号 (13)
    def r2(self):
        notes, got = [], 0
        s = self.slide(1)
        if s is None:
            return (2, '第2张目录', 0, 13, ['缺少第2张'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout2':
            got += 4; notes.append('标题和内容 OK')
        else:
            notes.append('版式=%s' % label)
        if '冬奥会比赛项目' in title_text(self.deck, s):
            got += 4; notes.append('标题 OK')
        if auto_number(s):
            got += 5; notes.append('编号 1.2.3. OK')
        else:
            notes.append('没有自动编号')
        return (2, '第2张目录', got, 13, notes)

    # 3. 第3-7张 两栏内容 (15)
    def r3(self):
        notes, got = [], 0
        hit = 0
        for idx in range(2, 7):
            s = self.slide(idx)
            if s is None:
                continue
            name, _ = layout_short(self.deck, s)
            if name == 'slideLayout4':
                hit += 1
        got = hit * 3
        notes.append('%d/5 张为两栏内容' % hit)
        return (3, '第3-7张版式', got, 15, notes)

    # 4. 第3-7张 插图 (15)
    def r4(self):
        notes, got = [], 0
        hit = 0
        for idx in range(2, 7):
            s = self.slide(idx)
            if s is not None and self.deck.image_targets(s):
                hit += 1
        got = hit * 3
        notes.append('%d/5 张有图片' % hit)
        return (4, '第3-7张插图', got, 15, notes)

    # 5. 超链接 + 动作按钮 (12)
    def r5(self):
        notes, got = [], 0
        s2 = self.slide(1)
        if s2 is not None and 'hlinksldjump' in ' '.join(ppt_actions(s2)):
            got += 6; notes.append('目录超链接 OK')
        else:
            notes.append('第2张没有跳转链接')
        btn = 0
        for idx in range(2, 7):
            s = self.slide(idx)
            if s is None:
                continue
            names = shape_names(s)
            if any('动作按钮' in n and '后退' in n for n in names):
                btn += 1
                ext = re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"/>', s.xml)
                if any(x == '720000' and y == '720000' for x, y in ext):
                    notes.append('第%d张按钮 2cm OK' % (idx + 1))
        if btn == 5:
            got += 6; notes.append('5 张均有动作按钮 OK')
        else:
            notes.append('动作按钮 %d/5 张' % btn)
        return (5, '链接与按钮', got, 12, notes)

    # 6. 末张 空白+艺术字 (8)
    def r6(self):
        notes, got = [], 0
        s = self.slide(len(self.deck.slides) - 1)
        if s is None:
            return (6, '艺术字结尾', 0, 8, ['没有幻灯片'])
        name, label = layout_short(self.deck, s)
        if name == 'slideLayout7':
            got += 3; notes.append('空白版式 OK')
        else:
            notes.append('末张版式=%s' % label)
        if has_gradfill(s):
            got += 5; notes.append('艺术字渐变填充 OK')
        else:
            notes.append('末张没有渐变填充')
        return (6, '艺术字结尾', got, 8, notes)

    # 7. 动画效果 (8)
    def r7(self):
        notes, got = [], 0
        hit = 0
        from _core import ppt_presets
        for idx in range(2, 7):
            s = self.slide(idx)
            if s is not None and any(p[0] == '6' and p[1] == 'entr' for p in ppt_presets(s)):
                hit += 1
        got = 8 if hit == 5 else round(hit * 1.5, 1)
        notes.append('带进入动画 %d/5 张' % hit)
        return (7, '动画效果', got, 8, notes)

    # 8. 放映方式 (8)
    def r8(self):
        show = self.deck.show_type
        if show == 'browse':
            return (8, '放映方式', 8, 8, ['观众自行浏览 OK'])
        return (8, '放映方式', 0, 8, ['放映类型=%r' % show])

    # 9. 全部切换=门 (9)
    def r9(self):
        notes, got = [], 0
        doors = sum(1 for s in self.deck.slides if 'doors' in transition_types(s))
        if doors == len(self.deck.slides) and doors:
            got += 9; notes.append('8 张全部"门"切换 OK')
        else:
            got = round(doors * 9.0 / max(1, len(self.deck.slides)), 1)
            notes.append('"门"切换 %d/%d 张' % (doors, len(self.deck.slides)))
        return (9, '切换效果', got, 9, notes)


if __name__ == '__main__':
    sys.exit(run(PptGrader6, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        '2022年准易：Ppt 2010 操作题--素材', '6', '1结果.pptx')))
