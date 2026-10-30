"""docx 底层读取：段落/表格/节/页眉页脚/样式继承。"""
import re
import zipfile

from .pkg import Package

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
V = 'urn:schemas-microsoft-com:vml'
O = 'urn:schemas-microsoft-com:office:office'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'


def q(tag, ns=W):
    return '{%s}%s' % (ns, tag)


def attr(el, name, default=''):
    return el.get(name, default) if el is not None else default


def wattr(el, name, default=''):
    return el.get(q(name), default) if el is not None else default


def on(el):
    """ST_OnOff：元素存在即真，val=0/false 为假。"""
    if el is None:
        return False
    val = el.get(q('val'))
    return val not in ('0', 'false', 'off')


GRADING_PARTS = re.compile(
    r'^word/(document|styles|settings|numbering|footnotes|endnotes|fontTable|webSettings|people)\.xml$'
    r'|^word/(header|footer)\d*\.xml$'
    r'|^word/_rels/[^/]+\.rels$'
    r'|^word/(diagrams|theme|charts)/[^/]+\.xml$'
)


def extract_word_parts(docx_path, target_dir):
    """只解出评分需要的 XML，跳过 media，避免每次判分产生上百个临时文件。"""
    count = 0
    with zipfile.ZipFile(docx_path) as zf:
        for info in zf.infolist():
            if GRADING_PARTS.match(info.filename):
                zf.extract(info, target_dir)
                count += 1
    return count


class Run:
    __slots__ = ('text', 'node', 'props')

    def __init__(self, text, node, props):
        self.text = text
        self.node = node
        self.props = props


class Paragraph:
    def __init__(self, node, index, props, runs, style_id):
        self.node = node
        self.index = index
        self.props = props
        self.runs = runs
        self.style = style_id
        self.text = ''.join(r.text for r in runs)

    def __repr__(self):
        return 'Paragraph(%d, %r)' % (self.index, self.text[:20])


class Table:
    def __init__(self, node, index):
        self.node = node
        self.index = index
        self.props = node.find(q('tblPr'))
        self.grid = node.find(q('tblGrid'))
        self.rows = []
        for tr in node.findall(q('tr')):
            cells = []
            for tc in tr.findall(q('tc')):
                text = ''.join(t.text or '' for t in tc.iter(q('t')))
                cells.append({'node': tc, 'text': text})
            self.rows.append(cells)

    @property
    def text(self):
        return '\n'.join(''.join(c['text'] for c in row) for row in self.rows)

    def cell_texts(self):
        return [[c['text'] for c in row] for row in self.rows]


class WordDoc(Package):
    def __init__(self, path):
        super().__init__(path)
        self.part = self.main_part('word/document.xml')
        self.root_el = self.root(self.part)
        self.body = self.root_el.find(q('body'))
        self._styles = None
        self._numfmt = None
        self.paragraphs = []
        self._collect_paragraphs()
        self.tables = [Table(t, i) for i, t in enumerate(self.body.iter(q('tbl')))] if self.body is not None else []

    def _collect_paragraphs(self):
        if self.body is None:
            return
        idx = 0
        for p in self.body.iter(q('p')):
            ppr = p.find(q('pPr'))
            style = wattr(ppr.find(q('pStyle')) if ppr is not None else None, 'val')
            props = self.para_props(p)
            runs = []
            for r in p.findall(q('r')):
                text = ''.join(t.text or '' for t in r.iter(q('t')))
                runs.append(Run(text, r, self.run_props(r, p)))
            for link in p.findall(q('hyperlink')):
                for r in link.findall(q('r')):
                    text = ''.join(t.text or '' for t in r.iter(q('t')))
                    runs.append(Run(text, r, self.run_props(r, p)))
            self.paragraphs.append(Paragraph(p, idx, props, runs, style))
            idx += 1

    # ---------- 样式 ----------
    @property
    def styles(self):
        if self._styles is None:
            self._styles = {}
            part = self.part.rsplit('/', 1)[0] + '/styles.xml'
            if self.has(part):
                root = self.root(part)
                defaults = root.find('%s/%s/%s' % (q('docDefaults'), q('rPrDefault'), q('rPr')))
                self._styles['__default__'] = self._rpr_dict(defaults)
                for st in root.iter(q('style')):
                    sid = wattr(st, 'styleId')
                    self._styles[sid] = {
                        'type': wattr(st, 'type'),
                        'name': wattr(st.find(q('name')), 'val'),
                        'based': wattr(st.find(q('basedOn')), 'val'),
                        'rPr': self._rpr_dict(st.find(q('rPr'))),
                        'pPr': self._ppr_dict(st.find(q('pPr'))),
                    }
            else:
                self._styles['__default__'] = {}
        return self._styles

    def style_chain(self, style_id):
        chain = []
        seen = set()
        cur = style_id
        while cur and cur in self.styles and cur not in seen:
            seen.add(cur)
            chain.append(self.styles[cur])
            cur = self.styles[cur].get('based')
        return list(reversed(chain))

    @staticmethod
    def _rpr_dict(rpr):
        if rpr is None:
            return {}
        out = {}
        fonts = rpr.find(q('rFonts'))
        if fonts is not None:
            out['font'] = attr(fonts, q('eastAsia')) or attr(fonts, q('ascii')) or attr(fonts, q('hAnsi'))
            out['ascii'] = attr(fonts, q('ascii'))
        sz = rpr.find(q('sz'))
        if sz is not None:
            try:
                out['size'] = int(wattr(sz, 'val', '0'))
            except ValueError:
                pass
        for key in ('b', 'i', 'strike', 'caps', 'smallCaps'):
            node = rpr.find(q(key))
            if node is not None:
                out[key] = on(node)
        color = rpr.find(q('color'))
        if color is not None:
            out['color'] = wattr(color, 'val')
        u = rpr.find(q('u'))
        if u is not None:
            out['underline'] = wattr(u, 'val', 'single')
        shd = rpr.find(q('shd'))
        if shd is not None:
            out['shd'] = wattr(shd, 'fill')
        return out

    @staticmethod
    def _ppr_dict(ppr):
        if ppr is None:
            return {}
        out = {}
        jc = ppr.find(q('jc'))
        if jc is not None:
            out['jc'] = wattr(jc, 'val')
        spacing = ppr.find(q('spacing'))
        if spacing is not None:
            for key in ('before', 'after', 'line', 'lineRule', 'beforeLines', 'afterLines'):
                val = spacing.get(q(key))
                if val is not None:
                    out[key] = val
        ind = ppr.find(q('ind'))
        if ind is not None:
            for key in ('left', 'right', 'firstLine', 'firstLineChars', 'leftChars', 'hanging'):
                val = ind.get(q(key))
                if val is not None:
                    out[key] = val
        numpr = ppr.find(q('numPr'))
        if numpr is not None:
            out['numId'] = wattr(numpr.find(q('numId')), 'val')
            out['ilvl'] = wattr(numpr.find(q('ilvl')), 'val')
        return out

    def para_props(self, p):
        merged = {}
        ppr = p.find(q('pPr'))
        style = wattr(ppr.find(q('pStyle')) if ppr is not None else None, 'val')
        for st in self.style_chain(style):
            merged.update({k: v for k, v in st.get('pPr', {}).items()})
        merged.update(self._ppr_dict(ppr))
        return merged

    def run_props(self, run, para=None):
        merged = dict(self.styles.get('__default__', {}))
        if para is not None and hasattr(para, 'find'):
            ppr = para.find(q('pPr'))
            style = wattr(ppr.find(q('pStyle')) if ppr is not None else None, 'val')
            for st in self.style_chain(style):
                merged.update(st.get('rPr', {}))
            if ppr is not None:
                merged.update(self._rpr_dict(ppr.find(q('rPr'))))
        merged.update(self._rpr_dict(run.find(q('rPr'))))
        return merged

    # ---------- 节 / 页面 ----------
    @property
    def sections(self):
        return list(self.root_el.iter(q('sectPr')))

    def body_section(self):
        if self.body is None:
            return None
        return self.body.find(q('sectPr'))

    def page_setup(self, sect=None):
        sect = sect or self.body_section() or (self.sections[-1] if self.sections else None)
        out = {}
        if sect is None:
            return out
        sz = sect.find(q('pgSz'))
        if sz is not None:
            out['w'] = int(wattr(sz, 'w', '0') or 0)
            out['h'] = int(wattr(sz, 'h', '0') or 0)
            out['orient'] = wattr(sz, 'orient', 'portrait')
            out['code'] = wattr(sz, 'code')
        mar = sect.find(q('pgMar'))
        if mar is not None:
            for key in ('top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'):
                val = mar.get(q(key))
                if val is not None:
                    try:
                        out[key] = int(val)
                    except ValueError:
                        pass
        cols = sect.find(q('cols'))
        if cols is not None:
            out['cols'] = int(wattr(cols, 'num', '1') or 1)
            out['col_space'] = wattr(cols, 'space')
            out['col_sep'] = wattr(cols, 'sep')
        grid = sect.find(q('docGrid'))
        if grid is not None:
            out['linePitch'] = wattr(grid, 'linePitch')
            out['charSpace'] = wattr(grid, 'charSpace')
        return out

    # ---------- 页眉页脚 ----------
    def _ref_parts(self, tag):
        out = []
        for sect in self.sections:
            for ref in sect.findall(q(tag)):
                rid = ref.get('{%s}id' % R, '')
                rtype = wattr(ref, 'type', 'default')
                target = self.rel_target(self.part, rid)
                if target and self.has(target):
                    out.append({'type': rtype, 'part': target, 'root': self.root(target)})
        return out

    @property
    def headers(self):
        return self._ref_parts('headerReference')

    @property
    def footers(self):
        return self._ref_parts('footerReference')

    def header_texts(self):
        return [t for t in (self.part_text(h['part']) for h in self.headers) if t.strip()]

    def footer_texts(self):
        return [t for t in (self.part_text(f['part']) for f in self.footers) if t.strip()]

    # ---------- 文本 ----------
    def part_text(self, part):
        if not self.has(part):
            return ''
        lines = []
        for p in self.root(part).iter(q('p')):
            line = ''.join(t.text or '' for t in p.iter(q('t')))
            if line.strip():
                lines.append(line)
        return '\n'.join(lines)

    def text(self):
        return '\n'.join(p.text for p in self.paragraphs if p.text.strip())

    def text_all(self):
        parts = [self.part]
        for key in ('footnotes', 'endnotes'):
            part = self.part.rsplit('/', 1)[0] + '/%s.xml' % key
            if self.has(part):
                parts.append(part)
        parts += [h['part'] for h in self.headers] + [f['part'] for f in self.footers]
        chunks = [self.part_text(p) for p in parts]
        for t in self.tables:
            chunks.append(t.text)
        return '\n'.join(c for c in chunks if c)

    def find(self, needle):
        return needle in self.text_all()

    def instr_texts(self):
        out = []
        for part in self.parts(self.part.rsplit('/', 1)[0] + '/'):
            for it in self.root(part).iter(q('instrText')):
                if it.text:
                    out.append(it.text)
        return out

    # ---------- 编号 ----------
    @property
    def numbering(self):
        if self._numfmt is None:
            self._numfmt = {}
            part = self.part.rsplit('/', 1)[0] + '/numbering.xml'
            if not self.has(part):
                return self._numfmt
            root = self.root(part)
            abstract = {}
            for an in root.iter(q('abstractNum')):
                aid = wattr(an, 'abstractNumId')
                levels = {}
                for lvl in an.findall(q('lvl')):
                    ilvl = wattr(lvl, 'ilvl')
                    fmt = wattr(lvl.find(q('numFmt')), 'val')
                    text = wattr(lvl.find(q('lvlText')), 'val')
                    levels[ilvl] = (fmt, text)
                abstract[aid] = levels
            for num in root.iter(q('num')):
                nid = wattr(num, 'numId')
                aid = wattr(num.find(q('abstractNumId')), 'val')
                self._numfmt[nid] = abstract.get(aid, {})
        return self._numfmt

    def list_format(self, para):
        numid = para.props.get('numId')
        ilvl = para.props.get('ilvl', '0')
        if not numid:
            return ('', '')
        return self.numbering.get(numid, {}).get(ilvl, ('', ''))

    # ---------- 其他 ----------
    def drawings(self):
        return [d for d in self.root_el.iter(q('drawing'))]

    def shapes_vml(self):
        out = []
        for part in self.parts(self.part.rsplit('/', 1)[0] + '/'):
            out.extend(self.root(part).iter('{%s}shape' % V))
        return out

    def page_borders(self):
        out = []
        for sect in self.sections:
            pb = sect.find(q('pgBorders'))
            if pb is not None:
                out.append(pb)
        return out

    def background(self):
        return self.root_el.find(q('background'))

    def zip_media(self):
        return [n for n in self.names if n.startswith(self.part.rsplit('/', 1)[0] + '/media/')]

    def xml_all(self):
        base = self.part.rsplit('/', 1)[0] + '/'
        chunks = []
        for part in self.names:
            if part.startswith(base) and part.endswith('.xml') and '/_rels/' not in part:
                chunks.append(self.read(part))
        return '\n'.join(chunks)
