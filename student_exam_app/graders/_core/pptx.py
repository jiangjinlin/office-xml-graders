"""pptx 底层读取：按放映顺序定位幻灯片、形状/文本/字体继承/切换动画。"""
from .pkg import Package

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

EMU_CM = 360000.0
TRANSITION_SKIP = {'sndAc', 'stSnd', 'snd', 'extLst'}


def q(tag, ns=P):
    return '{%s}%s' % (ns, tag)


def aq(tag):
    return '{%s}%s' % (A, tag)


def local(tag):
    return tag.rsplit('}', 1)[-1]


def slide_no(path):
    """从 slideN.xml 取 N，用于按自然序排列（避免 slide10 排在 slide2 前）。"""
    name = path.rsplit('/', 1)[-1]
    digits = ''.join(ch for ch in name if ch.isdigit())
    return int(digits) if digits else 0


class Run:
    __slots__ = ('text', 'node', 'size', 'bold', 'italic', 'font', 'color')

    def __init__(self, text, node, size, bold, italic, font, color):
        self.text = text
        self.node = node
        self.size = size
        self.bold = bold
        self.italic = italic
        self.font = font
        self.color = color


class Shape:
    def __init__(self, node, kind):
        self.node = node
        self.kind = kind
        self.name = ''
        self.placeholder = ''
        self.placeholder_idx = None
        self.left = self.top = self.width = self.height = None
        self.runs = []
        self.paragraphs = []
        self._parse()

    def _parse(self):
        nv = None
        for tag in ('nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr', 'nvCxnSpPr'):
            nv = self.node.find(q(tag))
            if nv is not None:
                break
        if nv is not None:
            cnv = nv.find(q('cNvPr'))
            self.name = cnv.get('name', '') if cnv is not None else ''
            ph = nv.find('%s/%s' % (q('nvPr'), q('ph')))
            if ph is not None:
                self.placeholder = ph.get('type', 'body')
                try:
                    self.placeholder_idx = int(ph.get('idx', '0'))
                except ValueError:
                    self.placeholder_idx = None

        xfrm = self.node.find('%s/%s' % (q('spPr'), aq('xfrm')))
        if xfrm is None:
            xfrm = self.node.find('%s/%s' % (q('grpSpPr'), aq('xfrm')))
        if xfrm is not None:
            off = xfrm.find(aq('off'))
            ext = xfrm.find(aq('ext'))
            if off is not None:
                self.left = int(off.get('x', '0')) / EMU_CM
                self.top = int(off.get('y', '0')) / EMU_CM
            if ext is not None:
                self.width = int(ext.get('cx', '0')) / EMU_CM
                self.height = int(ext.get('cy', '0')) / EMU_CM

        body = self.node.find(q('txBody'))
        if body is None:
            gf = self.node.find('%s/%s' % (q('graphic'), aq('graphicData')))
            if gf is not None:
                body = gf.find('.//' + aq('txBody'))
        if body is None:
            return
        for para in body.findall(aq('p')):
            ppr = para.find(aq('pPr'))
            runs = []
            for r in para.findall(aq('r')):
                rpr = r.find(aq('rPr'))
                t = r.find(aq('t'))
                runs.append(self._run(t.text if t is not None else '', rpr, ppr))
            if not runs:
                fld = para.find(aq('fld'))
                if fld is not None:
                    t = fld.find(aq('t'))
                    runs.append(self._run(t.text if t is not None else '', None, ppr))
            self.paragraphs.append({'node': para, 'runs': runs, 'text': ''.join(x.text for x in runs)})
            self.runs.extend(runs)

    @staticmethod
    def _run(text, rpr, ppr):
        size = 0
        bold = italic = False
        font = ''
        color = ''
        for source in (rpr, ppr.find(aq('defRPr')) if ppr is not None else None):
            if source is None:
                continue
            if not size and source.get('sz'):
                try:
                    size = int(source.get('sz')) / 100.0
                except ValueError:
                    pass
            if not bold and source.get('b') in ('1', 'true'):
                bold = True
            if not italic and source.get('i') in ('1', 'true'):
                italic = True
            if not font:
                for tag in ('ea', 'latin', 'cs'):
                    node = source.find(aq(tag))
                    if node is not None and node.get('typeface'):
                        font = node.get('typeface')
                        break
            if not color:
                clr = source.find('%s/%s' % (aq('solidFill'), aq('srgbClr')))
                if clr is not None:
                    color = clr.get('val', '')
        return Run(text, rpr, size, bold, italic, font, color)

    @property
    def text(self):
        return '\n'.join(p['text'] for p in self.paragraphs)

    def __repr__(self):
        return 'Shape(%s, %r)' % (self.name, self.text[:20])


class Slide:
    def __init__(self, part, index, root, xml):
        self.part = part
        self.index = index
        self.root = root
        self.xml = xml
        self.shapes = []
        for node in root.iter():
            name = local(node.tag)
            if name in ('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'):
                self.shapes.append(Shape(node, name))

    @property
    def title(self):
        for sh in self.shapes:
            if sh.placeholder in ('title', 'ctrTitle'):
                return sh.text
        return ''

    @property
    def text(self):
        return '\n'.join(sh.text for sh in self.shapes if sh.text)

    def by_placeholder(self, *types):
        for sh in self.shapes:
            if sh.placeholder in types:
                return sh
        return None

    def has_placeholder(self, ptype):
        return any(sh.placeholder == ptype for sh in self.shapes)

    def layout_of(self, deck):
        for rtype, target in deck.rels(self.part).values():
            if rtype.endswith('/slideLayout'):
                return target
        return ''

    @property
    def transition(self):
        node = self.root.find(q('transition'))
        if node is None:
            return ''
        for child in node:
            name = local(child.tag)
            if name not in TRANSITION_SKIP:
                return name
        return ''

    @property
    def transition_node(self):
        return self.root.find(q('transition'))

    @property
    def timing(self):
        return self.root.find(q('timing'))

    def animations(self):
        timing = self.timing
        return list(timing.iter(q('animEffect'))) + list(timing.iter(q('anim'))) if timing is not None else []

    def animation_timing(self):
        out = []
        timing = self.timing
        if timing is None:
            return out
        for node in timing.iter():
            ntype = node.get('nodeType')
            if ntype:
                out.append(ntype)
        return out


class PptDeck(Package):
    def __init__(self, path):
        super().__init__(path)
        self.part = self.main_part('ppt/presentation.xml')
        self.root_el = self.root(self.part)
        self._base = self.part.rsplit('/', 1)[0] + '/'
        self.slides = []
        self._load_slides()
        self._theme = None

    def _load_slides(self):
        lst = self.root_el.find(q('sldIdLst'))
        ids = list(lst.findall(q('sldId'))) if lst is not None else []
        for i, sld in enumerate(ids):
            rid = sld.get('{%s}id' % R, '')
            target = self.rel_target(self.part, rid)
            if target and self.has(target):
                self.slides.append(Slide(target, i, self.root(target), self.read(target)))
        if not self.slides:
            parts = [p for p in self.parts(self._base + 'slides/', exclude=('_rels/',))
                     if p.rsplit('/', 1)[-1].startswith('slide')]
            for i, part in enumerate(parts):
                self.slides.append(Slide(part, i, self.root(part), self.read(part)))

    @property
    def slide_size(self):
        sz = self.root_el.find(q('sldSz'))
        if sz is None:
            return (25.4, 19.05)
        return (int(sz.get('cx', '9144000')) / EMU_CM, int(sz.get('cy', '6858000')) / EMU_CM)

    @property
    def slide_size_type(self):
        sz = self.root_el.find(q('sldSz'))
        return sz.get('type', '') if sz is not None else ''

    @property
    def theme(self):
        if self._theme is None:
            parts = self.parts(self._base + 'theme/', exclude=('_rels/',))
            info = {'name': '', 'colors': {}, 'fonts': {}}
            if parts:
                root = self.root(parts[0])
                info['name'] = root.get('name', '')
                scheme = root.find('%s/%s' % (aq('themeElements'), aq('clrScheme')))
                if scheme is not None:
                    for key in scheme:
                        child = list(key)
                        if child:
                            info['colors'][local(key.tag)] = child[0].get('val') or child[0].get('lastClr', '')
                fs = root.find('%s/%s' % (aq('themeElements'), aq('fontScheme')))
                if fs is not None:
                    for key in ('majorFont', 'minorFont'):
                        node = fs.find(aq(key))
                        if node is None:
                            continue
                        ea = node.find(aq('ea'))
                        latin = node.find(aq('latin'))
                        info['fonts'][key] = {
                            'ea': ea.get('typeface', '') if ea is not None else '',
                            'latin': latin.get('typeface', '') if latin is not None else '',
                        }
            self._theme = info
        return self._theme

    @property
    def show_type(self):
        part = self._base + 'presProps.xml'
        if not self.has(part):
            return ''
        show = self.root(part).find(q('showPr'))
        if show is None:
            return ''
        for child in show:
            name = local(child.tag)
            if name in ('present', 'browse', 'kiosk'):
                return name
        return ''

    def masters(self):
        return [p for p in self.parts(self._base + 'slideMasters/', exclude=('_rels/',))
                if p.rsplit('/', 1)[-1].startswith('slideMaster')]

    def layouts(self):
        return [p for p in self.parts(self._base + 'slideLayouts/', exclude=('_rels/',))
                if p.rsplit('/', 1)[-1].startswith('slideLayout')]

    def media(self):
        return self.parts(self._base + 'media/', suffix='', exclude=('_rels/',))

    def notes_of(self, slide):
        for rtype, target in self.rels(slide.part).values():
            if rtype.endswith('/notesSlide') and self.has(target):
                return target
        return ''

    def notes_text(self, slide):
        part = self.notes_of(slide)
        if not part:
            return ''
        lines = []
        for p in self.root(part).iter(aq('p')):
            line = ''.join(t.text or '' for t in p.iter(aq('t')))
            if line.strip():
                lines.append(line)
        return '\n'.join(lines)

    def image_targets(self, slide):
        out = []
        for rtype, target in self.rels(slide.part).values():
            if rtype.endswith('/image'):
                out.append(target)
        return out

    # ---------- 版式/母版继承 ----------
    def layout_part(self, slide):
        for rtype, target in self.rels(slide.part).values():
            if rtype.endswith('/slideLayout') and self.has(target):
                return target
        return ''

    def master_part(self, layout_part):
        if not layout_part:
            return ''
        for rtype, target in self.rels(layout_part).values():
            if rtype.endswith('/slideMaster') and self.has(target):
                return target
        return ''

    @staticmethod
    def _placeholders(part_root):
        table = {}
        if part_root is None:
            return table
        for sp in part_root.iter(q('sp')):
            ph = sp.find('%s/%s' % (q('nvSpPr'), '%s/%s' % (q('nvPr'), q('ph'))))
            if ph is None:
                continue
            ptype = ph.get('type', 'body')
            try:
                pidx = int(ph.get('idx', '0'))
            except ValueError:
                pidx = None
            entry = (sp, ptype, pidx)
            if pidx is not None:
                table.setdefault(('idx', pidx), entry)
            table.setdefault(('type', ptype), entry)
        return table

    def _placeholder_defaults(self, sp, ptype):
        out = {'size': 0, 'bold': False, 'font': ''}
        if sp is None:
            return out
        candidates = []
        lst = sp.find('%s/%s' % (q('txBody'), aq('lstStyle')))
        if lst is not None:
            lvl = lst.find(aq('lvl1pPr'))
            if lvl is not None:
                candidates.append(lvl.find(aq('defRPr')))
        for run in sp.iter(aq('r')):
            candidates.append(run.find(aq('rPr')))
        for node in candidates:
            if node is None:
                continue
            if not out['size'] and node.get('sz'):
                try:
                    out['size'] = int(node.get('sz')) / 100.0
                except ValueError:
                    pass
            if not out['bold'] and node.get('b') in ('1', 'true'):
                out['bold'] = True
            if not out['font']:
                for tag in ('ea', 'latin'):
                    face = node.find(aq(tag))
                    if face is not None and face.get('typeface'):
                        out['font'] = face.get('typeface')
                        break
        return out

    @staticmethod
    def _rpr_basics(node, out=None):
        out = out or {'size': 0, 'bold': False, 'font': ''}
        if node is None:
            return out
        if not out['size'] and node.get('sz'):
            try:
                out['size'] = int(node.get('sz')) / 100.0
            except ValueError:
                pass
        if not out['bold'] and node.get('b') in ('1', 'true'):
            out['bold'] = True
        if not out['font']:
            for tag in ('ea', 'latin'):
                face = node.find(aq(tag))
                if face is not None and face.get('typeface'):
                    out['font'] = face.get('typeface')
                    break
        return out

    def _tx_styles(self, master_root):
        out = {}
        if master_root is None:
            return out
        ts = master_root.find(q('txStyles'))
        if ts is None:
            return out
        for key, name in (('title', 'titleStyle'), ('body', 'bodyStyle'), ('other', 'otherStyle')):
            node = ts.find(q(name))
            if node is None:
                continue
            lvl = node.find(aq('lvl1pPr'))
            if lvl is not None:
                out[key] = self._rpr_basics(lvl.find(aq('defRPr')))
        return out

    def resolved_shapes(self, slide):
        """补全占位符的坐标与字体（来自版式、再母版、最后主题）。"""
        layout = self.layout_part(slide)
        master = self.master_part(layout)
        lay_ph = self._placeholders(self.root(layout) if layout else None)
        mas_ph = self._placeholders(self.root(master) if master else None)
        tx_styles = self._tx_styles(self.root(master) if master else None)
        theme_fonts = self.theme.get('fonts', {})
        out = []
        for sh in slide.shapes:
            item = {
                'name': sh.name, 'kind': sh.kind, 'placeholder': sh.placeholder,
                'left': sh.left, 'top': sh.top, 'width': sh.width, 'height': sh.height,
                'text': sh.text, 'runs': [], 'shape': sh,
            }
            srcs = []
            if sh.placeholder:
                for table in (lay_ph, mas_ph):
                    hit = None
                    if sh.placeholder_idx is not None:
                        hit = table.get(('idx', sh.placeholder_idx))
                    hit = hit or table.get(('type', sh.placeholder))
                    if hit is not None:
                        srcs.append(hit[0])
            for sp_node in srcs:
                xfrm = sp_node.find('%s/%s' % (q('spPr'), aq('xfrm')))
                if xfrm is None:
                    continue
                off = xfrm.find(aq('off'))
                ext = xfrm.find(aq('ext'))
                if item['left'] is None and off is not None:
                    item['left'] = int(off.get('x', '0')) / EMU_CM
                    item['top'] = int(off.get('y', '0')) / EMU_CM
                if item['width'] in (None, 0) and ext is not None:
                    item['width'] = int(ext.get('cx', '0')) / EMU_CM
                    item['height'] = int(ext.get('cy', '0')) / EMU_CM
            defaults = {'size': 0, 'bold': False, 'font': ''}
            for sp_node in reversed(srcs):
                got = self._placeholder_defaults(sp_node, sh.placeholder)
                if got['size']:
                    defaults['size'] = got['size']
                if got['bold']:
                    defaults['bold'] = True
                if got['font']:
                    defaults['font'] = got['font']
            style_key = ('title' if sh.placeholder in ('title', 'ctrTitle')
                         else 'body' if sh.placeholder in ('body', 'subTitle', 'obj')
                         else 'other')
            master_defaults = tx_styles.get(style_key) or {}
            for key in ('size', 'font'):
                if not defaults[key] and master_defaults.get(key):
                    defaults[key] = master_defaults[key]
            if master_defaults.get('bold'):
                defaults['bold'] = True
            theme_font = (theme_fonts.get('majorFont', {}).get('ea')
                          if sh.placeholder in ('title', 'ctrTitle')
                          else theme_fonts.get('minorFont', {}).get('ea'))
            for run in sh.runs:
                size = run.size
                bold = run.bold
                font = run.font
                if defaults:
                    size = size or defaults['size']
                    bold = bold or defaults['bold']
                    font = font or defaults['font']
                if font.startswith('+'):
                    major = 'mj' in font
                    group = theme_fonts.get('majorFont' if major else 'minorFont', {})
                    font = group.get('latin' if font.endswith('lt') else 'ea') or group.get('ea') or ''
                item['runs'].append({
                    'text': run.text, 'size': size, 'bold': bold,
                    'italic': run.italic, 'font': font or theme_font or '',
                    'color': run.color,
                })
            out.append(item)
        return out

    def all_text(self):
        return '\n'.join(s.text for s in self.slides)

    def xml_all(self):
        chunks = [self.read(self.part)]
        for part in self.slides:
            chunks.append(part.xml)
        for pattern in ('slideMasters/slideMaster1.xml', 'theme/theme1.xml'):
            part = self._base + pattern
            if self.has(part):
                chunks.append(self.read(part))
        return '\n'.join(chunks)
