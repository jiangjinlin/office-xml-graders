"""核心库自检：三种文档各取一个样本，打印关键结构。"""
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, 'graders'))

from _core import ExcelBook, WordDoc, PptDeck  # noqa: E402

ROOT = os.path.dirname(BASE)


def excel():
    path = os.path.join(ROOT, 'Excel 2010 操作题题库', '1', '1结果.xlsx')
    with ExcelBook(path) as book:
        print('[Excel] sheets=%s' % [(s.name, s.part) for s in book.sheets])
        print('  strings=%d styles.fonts=%d xfs=%d' % (
            len(book.strings), len(book.styles.fonts), len(book.styles.xfs)))
        for ref in ('A1', 'A3', 'C3', 'E3'):
            cell = book.cell(ref)
            if cell is None:
                continue
            fmt = book.styles.get(cell.style_index)
            font = fmt.get('font', {})
            align = fmt.get('alignment', {})
            print('  %s value=%r formula=%r font=%s/%s bold=%s align=%s/%s numfmt=%r' % (
                ref, cell.value[:14], cell.formula, font.get('name'), font.get('size'),
                font.get('bold'), align.get('horizontal'), align.get('vertical'),
                fmt.get('numfmt')))
        charts = book.charts()
        print('  charts=%s tables=%s pivots=%s' % (charts, book.tables(), book.pivot_tables()))


def word():
    path = os.path.join(ROOT, 'Word 2010 操作题题库', '2', '1结果.docx')
    with WordDoc(path) as doc:
        print('[Word] 段落=%d 表格=%d 节=%d' % (len(doc.paragraphs), len(doc.tables), len(doc.sections)))
        for p in doc.paragraphs[:4]:
            first = p.runs[0].props if p.runs else {}
            print('  #%d %r font=%s size=%s bold=%s jc=%s' % (
                p.index, p.text[:20], first.get('font'), first.get('size'),
                first.get('bold'), p.props.get('jc')))
        print('  页面=%s' % doc.page_setup())
        print('  页眉=%s' % [t[:12] for t in doc.header_texts()])
        print('  页脚=%s' % [t[:12] for t in doc.footer_texts()])
        print('  表格首行=%s' % (doc.tables[0].cell_texts()[0][:4] if doc.tables else []))


def ppt():
    path = os.path.join(ROOT, 'PPT 2010 操作题题库', '5', '1结果.pptx')
    with PptDeck(path) as deck:
        print('[PPT] 幻灯片=%d 尺寸=%.1fx%.1fcm 主题=%s 放映=%s' % (
            len(deck.slides), deck.slide_size[0], deck.slide_size[1],
            deck.theme.get('name'), deck.show_type))
        for s in deck.slides[:3]:
            print('  #%d %s title=%r transition=%s shapes=%d' % (
                s.index + 1, s.part, s.title[:16], s.transition, len(s.shapes)))
            for sh in deck.resolved_shapes(s)[:3]:
                if sh['runs']:
                    r = sh['runs'][0]
                    print('     %s ph=%s %.1f,%.1f %.1fx%.1f font=%s size=%s bold=%s' % (
                        sh['name'][:14], sh['placeholder'] or '-', sh['left'] or 0, sh['top'] or 0,
                        sh['width'] or 0, sh['height'] or 0, r['font'], r['size'], r['bold']))


if __name__ == '__main__':
    excel()
    print()
    word()
    print()
    ppt()
