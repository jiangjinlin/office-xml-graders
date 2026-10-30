"""xlsx 底层读取：工作表按 rels 定位、单元格索引、样式解析。"""
import re
from .pkg import Package

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

BUILTIN_NUMFMT = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
}


def _attr(el, name, default=''):
    return el.get(name, default) if el is not None else default


def _color(el):
    if el is None:
        return {}
    out = {}
    for key in ('rgb', 'theme', 'tint', 'indexed', 'auto'):
        val = el.get(key)
        if val is not None:
            out[key] = val
    return out


def split_ref(ref):
    m = re.match(r'([A-Za-z]+)(\d+)', ref or '')
    if not m:
        return '', 0
    return m.group(1).upper(), int(m.group(2))


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def local_name(tag):
    return tag.rsplit('}', 1)[-1]


def apply_tint(rgb, tint):
    """theme+tint → 实际 RGB（Excel 规则：tint<0 变暗，tint>0 变亮）。"""
    if not re.fullmatch(r'[0-9A-F]{6}', rgb or ''):
        return rgb
    out = []
    for i in (0, 2, 4):
        c = int(rgb[i:i + 2], 16)
        if tint < 0:
            c = c * (1 + tint)
        else:
            c = c * (1 - tint) + 255 * tint
        out.append('%02X' % max(0, min(255, int(round(c)))))
    return ''.join(out)


class Cell:
    __slots__ = ('ref', 'type', 'raw', 'value', 'formula', 'style_index')

    def __init__(self, ref, ctype, raw, value, formula, style_index):
        self.ref = ref
        self.type = ctype
        self.raw = raw
        self.value = value
        self.formula = formula
        self.style_index = style_index

    def __repr__(self):
        return 'Cell(%s=%r)' % (self.ref, self.value)


class Sheet:
    def __init__(self, name, part, state, root, xml):
        self.name = name
        self.part = part
        self.state = state
        self.root = root
        self.xml = xml
        self._cells = None

    def cells(self):
        if self._cells is None:
            self._cells = {}
            for c in self.root.iter('{%s}c' % S):
                ref = c.get('r', '')
                if ref:
                    self._cells[ref] = c
        return self._cells

    def rows(self):
        return list(self.root.iter('{%s}row' % S))

    def merged(self):
        out = []
        mc = self.root.find('{%s}mergeCells' % S)
        if mc is not None:
            for m in mc.iter('{%s}mergeCell' % S):
                out.append(m.get('ref', ''))
        return out


class Styles:
    def __init__(self, root=None):
        self.fonts, self.fills, self.borders, self.xfs = [], [], [], []
        self.numfmts = dict(BUILTIN_NUMFMT)
        if root is None:
            return
        fonts = root.find('{%s}fonts' % S)
        if fonts is not None:
            self.fonts = [self._font(f) for f in fonts.findall('{%s}font' % S)]
        fills = root.find('{%s}fills' % S)
        if fills is not None:
            self.fills = [self._fill(f) for f in fills.findall('{%s}fill' % S)]
        borders = root.find('{%s}borders' % S)
        if borders is not None:
            self.borders = [self._border(b) for b in borders.findall('{%s}border' % S)]
        nf = root.find('{%s}numFmts' % S)
        if nf is not None:
            for item in nf.iter('{%s}numFmt' % S):
                try:
                    self.numfmts[int(item.get('numFmtId', '0'))] = item.get('formatCode', '')
                except ValueError:
                    pass
        self.xfs = [self._xf(x) for x in self._xfs(root)]
        self.cell_style_names = {}
        cs = root.find('{%s}cellStyles' % S)
        if cs is not None:
            for item in cs.findall('{%s}cellStyle' % S):
                try:
                    self.cell_style_names[int(item.get('xfId', '0') or 0)] = item.get('name', '')
                except ValueError:
                    pass

    @staticmethod
    def _xfs(root):
        xfs = root.find('{%s}cellXfs' % S)
        return list(xfs) if xfs is not None else []

    @staticmethod
    def _font(el):
        return {
            'name': _attr(el.find('{%s}name' % S), 'val'),
            'size': _attr(el.find('{%s}sz' % S), 'val'),
            'bold': el.find('{%s}b' % S) is not None and _attr(el.find('{%s}b' % S), 'val', '1') not in ('0', 'false'),
            'italic': el.find('{%s}i' % S) is not None and _attr(el.find('{%s}i' % S), 'val', '1') not in ('0', 'false'),
            'underline': _attr(el.find('{%s}u' % S), 'val', '') if el.find('{%s}u' % S) is not None else '',
            'strike': el.find('{%s}strike' % S) is not None,
            'color': _color(el.find('{%s}color' % S)),
        }

    @staticmethod
    def _fill(el):
        pf = el.find('{%s}patternFill' % S)
        if pf is None:
            return {'pattern': '', 'fg': {}, 'bg': {}}
        return {
            'pattern': pf.get('patternType', ''),
            'fg': _color(pf.find('{%s}fgColor' % S)),
            'bg': _color(pf.find('{%s}bgColor' % S)),
        }

    @staticmethod
    def _border(el):
        out = {}
        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            node = el.find('{%s}%s' % (S, side))
            out[side] = {
                'style': node.get('style', '') if node is not None else '',
                'color': _color(node.find('{%s}color' % S)) if node is not None else {},
            }
        return out

    @staticmethod
    def _xf(el):
        align = el.find('{%s}alignment' % S)
        try:
            xf_id = int(el.get('xfId', '0') or 0)
        except ValueError:
            xf_id = 0
        return {
            'numFmtId': int(el.get('numFmtId', '0') or 0),
            'fontId': int(el.get('fontId', '0') or 0),
            'fillId': int(el.get('fillId', '0') or 0),
            'borderId': int(el.get('borderId', '0') or 0),
            'xfId': xf_id,
            'alignment': dict(align.attrib) if align is not None else {},
        }

    def get(self, index):
        if not self.xfs:
            index = 0
        if index < 0 or index >= len(self.xfs):
            index = 0
        xf = self.xfs[index]
        return {
            'font': self.fonts[xf['fontId']] if xf['fontId'] < len(self.fonts) else {},
            'fill': self.fills[xf['fillId']] if xf['fillId'] < len(self.fills) else {},
            'border': self.borders[xf['borderId']] if xf['borderId'] < len(self.borders) else {},
            'alignment': xf['alignment'],
            'numfmt_id': xf['numFmtId'],
            'numfmt': self.numfmts.get(xf['numFmtId'], ''),
            'style_name': self.cell_style_names.get(xf.get('xfId', 0), ''),
        }


class ExcelBook(Package):
    def __init__(self, path):
        super().__init__(path)
        self._ss = None
        self._styles = None
        self._index = {}
        self.workbook_part = self.main_part('xl/workbook.xml')
        base = self.workbook_part.rsplit('/', 1)[0] + '/' if '/' in self.workbook_part else ''
        self._base = base
        self.sheets = []
        self._load_sheets()

    def _load_sheets(self):
        rels = self.rels(self.workbook_part)
        wb = self.root(self.workbook_part)
        for sh in wb.iter('{%s}sheet' % S):
            rid = sh.get('{%s}id' % R, '')
            target = rels.get(rid, ('', ''))[1]
            if not target or not self.has(target):
                continue
            xml = self.read(target)
            sheet = Sheet(sh.get('name', ''), target, sh.get('state', 'visible'),
                          self.root(target), xml)
            self.sheets.append(sheet)
        if not self.sheets:
            for part in self.parts(self._base + 'worksheets/', exclude=('_rels/',)):
                sheet = Sheet(part.rsplit('/', 1)[-1], part, 'visible', self.root(part), self.read(part))
                self.sheets.append(sheet)

    @property
    def sheet_names(self):
        return [s.name for s in self.sheets]

    @property
    def strings(self):
        if self._ss is None:
            self._ss = []
            part = self._base + 'sharedStrings.xml'
            if self.has(part):
                for si in self.root(part).iter('{%s}si' % S):
                    self._ss.append(''.join(t.text or '' for t in si.iter('{%s}t' % S)))
        return self._ss

    @property
    def styles(self):
        if self._styles is None:
            part = self._base + 'styles.xml'
            self._styles = Styles(self.root(part)) if self.has(part) else Styles()
        return self._styles

    def theme_colors(self):
        """{0:rgb,...} 按样式表 theme 索引顺序：0=lt1,1=dk1,2=lt2,3=dk2,4..9=accent1..6。"""
        if not hasattr(self, '_theme'):
            self._theme = {}
            part = self._base + 'theme/theme1.xml'
            if self.has(part):
                scheme = self.root(part).find('.//{%s}clrScheme'
                                              % 'http://schemas.openxmlformats.org/drawingml/2006/main')
                if scheme is not None:
                    order = ['dk1', 'lt1', 'dk2', 'lt2',
                             'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
                    raw = {}
                    for child in scheme:
                        node = list(child)[0] if list(child) else None
                        if node is None:
                            continue
                        val = node.get('lastClr') or node.get('val') or ''
                        raw[local_name(child.tag)] = val[-6:].upper()
                    index = [raw.get('lt1', ''), raw.get('dk1', ''), raw.get('lt2', ''),
                             raw.get('dk2', '')] + [raw.get('accent%d' % i, '') for i in range(1, 7)]
                    self._theme = dict(enumerate(index))
        return self._theme

    def resolve_color(self, color):
        """{'rgb'|'theme'|'tint'} → RRGGBB。"""
        if not color:
            return ''
        rgb = color.get('rgb') or ''
        if rgb:
            out = rgb[-6:].upper()
        elif color.get('theme') is not None:
            try:
                out = self.theme_colors().get(int(color['theme']), '')
            except (TypeError, ValueError):
                out = ''
        else:
            out = ''
        tint = color.get('tint')
        if out and tint:
            try:
                t = float(tint)
            except ValueError:
                t = 0.0
            out = apply_tint(out, t)
        return out

    def sheet(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        for s in self.sheets:
            if name and name in s.name:
                return s
        return None

    def sheet_of_root(self, root):
        for s in self.sheets:
            if s.root is root:
                return s
        return None

    def cell(self, ref, sheet=None):
        if sheet is None:
            sheet = self.sheets[0] if self.sheets else None
        if isinstance(sheet, str):
            sheet = self.sheet(sheet)
        if sheet is None:
            return None
        node = sheet.cells().get(ref)
        return self._to_cell(node) if node is not None else None

    def cell_in(self, root, ref):
        node = self.node_in(root, ref)
        return self._to_cell(node) if node is not None else None

    def node_in(self, root, ref):
        sheet = self.sheet_of_root(root)
        if sheet is not None:
            return sheet.cells().get(ref)
        for c in root.iter('{%s}c' % S):
            if c.get('r', '') == ref:
                return c
        return None

    def _to_cell(self, node):
        ctype = node.get('t', 'n')
        raw = ''
        value = ''
        v = node.find('{%s}v' % S)
        if v is not None and v.text is not None:
            raw = v.text
            if ctype == 's':
                try:
                    idx = int(raw)
                    value = self.strings[idx] if 0 <= idx < len(self.strings) else ''
                except ValueError:
                    value = ''
            elif ctype == 'b':
                value = 'TRUE' if raw == '1' else 'FALSE'
            else:
                value = raw
        elif ctype == 'inlineStr':
            is_node = node.find('{%s}is' % S)
            if is_node is not None:
                value = ''.join(t.text or '' for t in is_node.iter('{%s}t' % S))
        f = node.find('{%s}f' % S)
        formula = f.text if f is not None and f.text else ''
        try:
            style_index = int(node.get('s', '0') or 0)
        except ValueError:
            style_index = 0
        return Cell(node.get('r', ''), ctype, raw, value, formula, style_index)

    def value_in(self, root, ref):
        cell = self.cell_in(root, ref)
        return cell.value if cell is not None else ''

    def formula_in(self, root, ref):
        cell = self.cell_in(root, ref)
        return cell.formula if cell is not None else ''

    def format_in(self, root, ref):
        cell = self.cell_in(root, ref)
        if cell is None:
            return {}
        return self.styles.get(cell.style_index)

    def charts(self):
        return self.parts(self._base + 'charts/', exclude=('_rels/',))

    def tables(self):
        return self.parts(self._base + 'tables/', exclude=('_rels/',))

    def pivot_tables(self):
        return self.parts(self._base + 'pivotTables/', exclude=('_rels/',))

    def all_xml(self, include=('sharedStrings', 'workbook', 'worksheets', 'styles')):
        chunks = []
        for part in self.names:
            if not part.startswith(self._base) or not part.endswith('.xml'):
                continue
            if any(key in part for key in include):
                chunks.append(self.read(part))
        return '\n'.join(chunks)
