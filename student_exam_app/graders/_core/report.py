"""统一的评分结果结构与输出。"""
import io
import json
import os
import sys


class Report:
    def __init__(self, title='', source=''):
        self.title = title
        self.source = source
        self.items = []

    def add(self, no, name, got, full, notes=None, status=None):
        got = round(min(float(got), float(full)), 1)
        if status is None:
            status = 'PASS' if got >= float(full) * 0.6 else 'FAIL'
        self.items.append({
            'no': no, 'name': name, 'got': got, 'full': float(full),
            'notes': [str(n) for n in (notes or [])], 'status': status,
        })
        return self

    @property
    def total(self):
        return round(sum(i['got'] for i in self.items), 1)

    @property
    def full(self):
        return round(sum(i['full'] for i in self.items), 1)

    @property
    def percent(self):
        return round(self.total / self.full * 100, 1) if self.full else 0.0

    @property
    def passed(self):
        return self.percent >= 60

    def to_dict(self):
        return {
            'title': self.title,
            'source': self.source,
            'score': self.percent,
            'raw': self.total,
            'full': self.full,
            'passed': self.passed,
            'items': self.items,
        }

    def to_text(self):
        lines = []
        lines.append(self.title or '操作题评分')
        if self.source:
            lines.append('文件: %s' % os.path.basename(self.source))
        lines.append('=' * 52)
        for item in self.items:
            lines.append('[%s] %2d. %-14s %5.1f / %-4g' % (
                item['status'], item['no'], item['name'], item['got'], item['full']))
            for note in item['notes']:
                lines.append('       ' + note)
        lines.append('=' * 52)
        lines.append('总分: %g / %g   得分率: %g%%' % (self.total, self.full, self.percent))
        return '\n'.join(lines)

    def to_lines(self):
        """兼容旧前端：以 pass/fail/info 前缀逐行输出。"""
        out = []
        for item in self.items:
            out.append('[%s] %2d. %s  %.1f/%g' % (
                'pass' if item['status'] == 'PASS' else 'fail',
                item['no'], item['name'], item['got'], item['full']))
            for note in item['notes']:
                mark = 'pass' if 'OK' in note or '通过' in note else 'info'
                out.append('[%s]      %s' % (mark, note))
        return out

    def write_json(self, path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)


def setup_stdout():
    if os.name == 'nt' and getattr(sys.stdout, 'buffer', None) is not None:
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except Exception:
            pass


def run(grader_cls, argv=None, title='操作题评分', default=None):
    """统一入口：python grader.py 学生文件 [--json 结果.json]"""
    argv = list(sys.argv[1:] if argv is None else argv)
    json_path = ''
    if '--json' in argv:
        i = argv.index('--json')
        json_path = argv[i + 1] if i + 1 < len(argv) else ''
        del argv[i:i + 2]
    if not argv and default:
        argv = [default]
    if not argv:
        print('用法: python %s 学生文件 [--json 结果.json]' % os.path.basename(sys.argv[0]))
        return 2

    setup_stdout()
    grader = grader_cls(argv[0])
    report = grader.grade()
    if not isinstance(report, Report):
        raise TypeError('grade() 必须返回 Report')
    report.source = argv[0]
    if title and not report.title:
        report.title = title
    print(report.to_text())
    if json_path:
        report.write_json(json_path)
        print('结果已写入 %s' % json_path)
    return 0
